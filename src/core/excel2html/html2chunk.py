from pathlib import Path
from bs4 import BeautifulSoup
import copy
import json
import re
import tiktoken

# 初始化 tokenizer（cl100k_base 用于 GPT-4/GPT-3.5-turbo）
_tokenizer = None

def get_tokenizer():
    """懒加载 tokenizer"""
    global _tokenizer
    if _tokenizer is None:
        _tokenizer = tiktoken.get_encoding("cl100k_base")
    return _tokenizer


def estimate_tokens(text: str) -> int:
    """使用 tiktoken 精确计算 token 数量"""
    return len(get_tokenizer().encode(text))


def extract_note_references(text):
    """从文本中提取注释引用
    
    支持格式：
    - [注1] 单个注释
    - [注1][注2] 连续多个注释
    - [注1、2、3] 或 [注1,2,3] 合并格式
    - 注1 无方括号格式
    """
    refs = set()
    
    # 1. 匹配合并格式: [注1、2、3] 或 [注1,2,3]
    multi_refs = re.findall(r'\[(注)([\d、,，]+)\]', text)
    for prefix, nums_str in multi_refs:
        nums = re.split(r'[、,，]', nums_str)
        for num in nums:
            num = num.strip()
            if num:
                refs.add(f"{prefix}{num}")
    
    # 2. 匹配单个方括号注释: [注1], [备注], [说明2] 等
    bracket_refs = re.findall(r'\[(注\s*\d*|备注\s*\d*|说明\s*\d*|注意\s*\d*)\s*\]', text)
    refs.update(ref.replace(' ', '') for ref in bracket_refs)
    
    # 3. 匹配无方括号的注释引用: 数值注1
    superscript_refs = re.findall(r'[^\[](注\d+)(?:[：:）\)]|$|\s)', text)
    refs.update(superscript_refs)
    
    # 4. 特殊符号
    if '*' in text:
        refs.add('*')
    if '※' in text:
        refs.add('※')
    
    return refs


def distribute_assets_and_chunk(
    long_html_content, max_rows_per_chunk: int = None, max_tokens_per_chunk: int = None,
    min_tokens_per_chunk: int = None, token_strategy: str = "prefer_max"
):
    """
    核心逻辑：将长 HTML 切分，并把全局资产（Context/Caption/Header）分发给每个片段

    参数:
        long_html_content: 完整 HTML 内容
        max_rows_per_chunk: 按行数切分（优先）
        max_tokens_per_chunk: 按 token 数切分（更精确）- 最大 token 限制
        min_tokens_per_chunk: 最小 token 数（可选）- chunk 至少要达到此值才切分
        token_strategy: 切分策略（仅在启用 min_tokens_per_chunk 时生效）
            - "prefer_max": 接近最大值 - 累加到接近 max_tokens 才切分（默认）
            - "prefer_min": 接近最小值 - 超过 min_tokens 就立即切分

    如果两个参数都未指定，默认 max_rows_per_chunk=8
    
    返回:
        dict: {
            "chunks": list[str],  # 切分后的 HTML 片段
            "warnings": list[dict],  # 超限警告信息
            "stats": dict  # 统计信息
        }
    """
    if max_rows_per_chunk is None and max_tokens_per_chunk is None:
        max_rows_per_chunk = 8
    
    # 验证 min_tokens_per_chunk 参数
    if min_tokens_per_chunk is not None:
        if max_tokens_per_chunk is None:
            raise ValueError("min_tokens_per_chunk 只能在 token 模式下使用，请同时设置 max_tokens_per_chunk")
        if min_tokens_per_chunk >= max_tokens_per_chunk:
            raise ValueError(f"min_tokens_per_chunk ({min_tokens_per_chunk}) 必须小于 max_tokens_per_chunk ({max_tokens_per_chunk})")

    soup = BeautifulSoup(long_html_content, "html.parser")

    # 1. 提取全局资产 (Context Div)
    context_div = soup.find("div", class_="rag-context")
    if not context_div:
        table_node = soup.find("table")
        if table_node and table_node.find_previous_sibling("div"):
            context_div = table_node.find_previous_sibling("div")

    # 1.1 提取注释元数据
    notes_meta_script = soup.find("script", class_="table-notes-meta")
    header_notes = {}
    conditional_notes = {}
    if notes_meta_script:
        try:
            notes_meta = json.loads(notes_meta_script.string)
            header_notes = notes_meta.get("header_notes", {})
            conditional_notes = notes_meta.get("conditional_notes", {})
        except (json.JSONDecodeError, AttributeError):
            pass

    # 2. 提取表格核心组件
    original_table = soup.find("table")
    if not original_table:
        return {
            "chunks": [long_html_content],
            "warnings": [],
            "stats": {"total_chunks": 1, "oversized_chunks": 0}
        }

    caption = original_table.find("caption")

    header_rows = []
    thead = original_table.find("thead")
    if thead:
        header_rows = thead.find_all("tr")
    else:
        header_rows = original_table.find_all("tr")[:1]

    # 3. 准备数据行（排除注释行）
    tbody = original_table.find("tbody")
    if tbody:
        all_body_rows = tbody.find_all("tr")
        data_rows = [row for row in all_body_rows if "table-note-row" not in row.get("class", [])]
    else:
        all_rows = original_table.find_all("tr")
        data_rows = [row for row in all_rows if row not in header_rows]

    # 4. 计算基础固定开销（用于 token 模式）
    fixed_parts = []
    if context_div:
        fixed_parts.append(str(context_div))
    if caption:
        fixed_parts.append(str(caption))
    for h_row in header_rows:
        fixed_parts.append(str(h_row))
    base_fixed_overhead = estimate_tokens("".join(fixed_parts))
    
    # 4.1 预计算 header_notes 的固定开销（每个 chunk 都会添加）
    header_notes_text = " | ".join(header_notes.values()) if header_notes else ""
    header_notes_overhead = estimate_tokens(header_notes_text) if header_notes_text else 0
    
    # 预计算表头文本（用于注释引用匹配）
    header_text = " ".join(str(row) for row in header_rows)

    chunks = []
    chunk_token_counts = []  # 记录每个 chunk 的实际 token 数
    warnings = []  # 超限警告
    current_chunk_data = []
    current_chunk_tokens = 0

    def calculate_notes_overhead(pending_rows):
        """动态计算当前 chunk 实际会匹配的注释 token 开销"""
        if not header_notes and not conditional_notes:
            return 0
        
        # 从待处理行和表头中提取注释引用
        chunk_text = " ".join(str(row) for row in pending_rows)
        all_text = chunk_text + " " + header_text
        chunk_refs = extract_note_references(all_text)
        
        # 收集实际会添加的注释（去重）
        actual_notes = []
        seen_notes = set()
        for note in header_notes.values():
            if note not in seen_notes:
                actual_notes.append(note)
                seen_notes.add(note)
        for key, note in conditional_notes.items():
            if key in chunk_refs and note not in seen_notes:
                actual_notes.append(note)
                seen_notes.add(note)
        
        if not actual_notes:
            return 0
        
        notes_text = " | ".join(actual_notes)
        return estimate_tokens(f" 【表格注释】{notes_text}")

    def should_split(row_count, row_tokens, pending_rows, new_row):
        """判断是否应该切分（动态计算注释开销）
        
        返回 True 表示：在加入 new_row 之前，先把 pending_rows 输出为一个 chunk
        
        策略说明：
        - prefer_max: 尽量累积到接近 max_tokens，只有加入新行会超过 max 时才切分
        - prefer_min: 只要当前已达到 min_tokens，就可以切分（但不能超过 max）
        """
        if max_tokens_per_chunk is not None:
            # 计算当前 chunk 的 token 数（不含新行）
            current_notes_overhead = calculate_notes_overhead(pending_rows)
            current_total = current_chunk_tokens + base_fixed_overhead + current_notes_overhead
            
            # 计算如果加入新行后的总 token 数
            test_rows = pending_rows + [new_row]
            notes_overhead = calculate_notes_overhead(test_rows)
            total_overhead = base_fixed_overhead + notes_overhead
            potential_total = current_chunk_tokens + row_tokens + total_overhead
            
            # 如果加入新行会超过最大限制，必须切分
            if potential_total > max_tokens_per_chunk:
                return True
            
            # 如果启用了最小 token 限制且使用 prefer_min 策略
            if min_tokens_per_chunk is not None and token_strategy == "prefer_min":
                # 当前 chunk 已达到最小值，可以切分
                if current_total >= min_tokens_per_chunk:
                    return True
            
            # 其他情况（prefer_max 或未达到 min）：继续累积
            return False
        else:
            return row_count >= max_rows_per_chunk

    def build_chunk(data_rows_for_chunk):
        """组装一个 chunk，智能添加匹配的注释"""
        new_soup = BeautifulSoup("<div></div>", "html.parser")
        wrapper_div = new_soup.div

        # 从数据行和表头中提取注释引用
        chunk_text = " ".join(str(row) for row in data_rows_for_chunk)
        header_text = " ".join(str(row) for row in header_rows)
        all_text = chunk_text + " " + header_text
        chunk_refs = extract_note_references(all_text)
        
        matched_notes = []
        seen_notes = set()  # 用于去重
        for key, note in header_notes.items():
            if note not in seen_notes:
                matched_notes.append(note)
                seen_notes.add(note)
        for key, note in conditional_notes.items():
            if key in chunk_refs and note not in seen_notes:
                matched_notes.append(note)
                seen_notes.add(note)
        
        if context_div:
            new_context = copy.copy(context_div)
            if matched_notes:
                notes_text = " | ".join(matched_notes)
                new_context.string = (new_context.get_text() or "") + f" 【表格注释】{notes_text}"
            wrapper_div.append(new_context)

        new_table = new_soup.new_tag("table")
        new_table.attrs = original_table.attrs
        new_table["border"] = "1"
        new_table["style"] = "border-collapse:collapse"

        if caption:
            new_table.append(copy.copy(caption))

        new_thead = new_soup.new_tag("thead")
        for h_row in header_rows:
            new_thead.append(copy.copy(h_row))
        new_table.append(new_thead)

        new_tbody = new_soup.new_tag("tbody")
        for d_row in data_rows_for_chunk:
            new_tbody.append(copy.copy(d_row))
        new_table.append(new_tbody)

        wrapper_div.append(new_table)
        return str(wrapper_div)

    # 5. 逐行累加切分
    for i, row in enumerate(data_rows):
        row_tokens = estimate_tokens(str(row))

        if current_chunk_data and should_split(len(current_chunk_data), row_tokens, current_chunk_data, row):
            # 记录当前 chunk 的 token 数
            final_notes_overhead = calculate_notes_overhead(current_chunk_data)
            final_total = current_chunk_tokens + base_fixed_overhead + final_notes_overhead
            chunk_token_counts.append(final_total)
            chunks.append(build_chunk(current_chunk_data))
            current_chunk_data = []
            current_chunk_tokens = 0

        current_chunk_data.append(row)
        current_chunk_tokens += row_tokens
        
        # 检查当前 chunk 是否已超限（处理单行超限的情况）
        if max_tokens_per_chunk is not None:
            current_notes_overhead = calculate_notes_overhead(current_chunk_data)
            current_total = current_chunk_tokens + base_fixed_overhead + current_notes_overhead
            if current_total > max_tokens_per_chunk:
                # 记录超限警告
                chunk_index = len(chunks)
                warnings.append({
                    "chunk_index": chunk_index,
                    "actual_tokens": current_total,
                    "limit": max_tokens_per_chunk,
                    "overflow": current_total - max_tokens_per_chunk,
                    "row_count": len(current_chunk_data),
                    "reason": "单行数据 + 固定开销 + 注释超过 token 限制" if len(current_chunk_data) == 1 else "累积数据超过 token 限制"
                })
                # 当前 chunk 已超限，立即输出
                chunk_token_counts.append(current_total)
                chunks.append(build_chunk(current_chunk_data))
                current_chunk_data = []
                current_chunk_tokens = 0

        if i == len(data_rows) - 1 and current_chunk_data:
            final_notes_overhead = calculate_notes_overhead(current_chunk_data)
            final_total = current_chunk_tokens + base_fixed_overhead + final_notes_overhead
            chunk_token_counts.append(final_total)
            chunks.append(build_chunk(current_chunk_data))

    # 构建统计信息
    stats = {
        "total_chunks": len(chunks),
        "oversized_chunks": len(warnings),
        "token_counts": chunk_token_counts,
        "max_token_count": max(chunk_token_counts) if chunk_token_counts else 0,
        "min_token_count": min(chunk_token_counts) if chunk_token_counts else 0,
        "avg_token_count": sum(chunk_token_counts) / len(chunk_token_counts) if chunk_token_counts else 0,
        "base_fixed_overhead": base_fixed_overhead,
    }
    
    if max_tokens_per_chunk:
        stats["token_limit"] = max_tokens_per_chunk
    if min_tokens_per_chunk:
        stats["min_token_limit"] = min_tokens_per_chunk
        stats["token_strategy"] = token_strategy

    return {
        "chunks": chunks,
        "warnings": warnings,
        "stats": stats
    }


def process_and_merge_html(file_path_str, separator="!!!_CHUNK_BREAK_!!!"):
    """读取文件 -> 切分 -> 合并 -> 保存"""
    source_path = Path(file_path_str)

    if not source_path.exists():
        print(f"❌ 错误：找不到文件 '{source_path}'")
        return

    print(f"📂 正在读取: {source_path.name}")
    content = source_path.read_text(encoding="utf-8")

    result = distribute_assets_and_chunk(content, max_rows_per_chunk=2)
    chunks = result["chunks"]
    warnings = result["warnings"]
    stats = result["stats"]

    print(f"🔪 切分完成：共生成 {stats['total_chunks']} 个片段")
    print(f"📊 Token 统计: 最小={stats['min_token_count']}, 最大={stats['max_token_count']}, 平均={stats['avg_token_count']:.1f}")
    
    # 输出超限警告
    if warnings:
        print(f"\n⚠️  警告：有 {len(warnings)} 个片段超过 token 限制：")
        for w in warnings:
            print(f"   - 片段 #{w['chunk_index']}: {w['actual_tokens']} tokens (超出 {w['overflow']})")
            print(f"     原因: {w['reason']}")

    formatted_separator = f"\n\n{separator}\n\n"
    merged_content = formatted_separator.join(chunks)

    new_filename = source_path.stem + "_chunk_merged" + source_path.suffix
    output_path = source_path.with_name(new_filename)

    try:
        output_path.write_text(merged_content, encoding="utf-8")
        print(f"\n✅ 合并成功！文件已保存至: {output_path.absolute()}")
        print(f"🔑 使用的分隔符: {separator}")
    except IOError as e:
        print(f"❌ 写入失败: {e}")


if __name__ == "__main__":
    my_input_file = "Files\\Excel\\本国子目注释调整表.html"
    process_and_merge_html(my_input_file)
