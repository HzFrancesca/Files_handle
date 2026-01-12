from pathlib import Path
from bs4 import BeautifulSoup
import copy


def estimate_tokens(text: str) -> int:
    """估算文本的 token 数量（中文约2.5字符=1token）"""
    return int(len(text) / 2.5)


def distribute_assets_and_chunk(
    long_html_content, max_rows_per_chunk: int = None, max_tokens_per_chunk: int = None
):
    """
    核心逻辑：将长 HTML 切分，并把全局资产（Context/Caption/Header）分发给每个片段

    参数:
        long_html_content: 完整 HTML 内容
        max_rows_per_chunk: 按行数切分（优先）
        max_tokens_per_chunk: 按 token 数切分（更精确）

    如果两个参数都未指定，默认 max_rows_per_chunk=8
    """
    # 默认值
    if max_rows_per_chunk is None and max_tokens_per_chunk is None:
        max_rows_per_chunk = 8

    soup = BeautifulSoup(long_html_content, "html.parser")

    # 1. 提取全局资产 (Context Div)
    context_div = soup.find("div", class_="rag-context")
    if not context_div:
        table_node = soup.find("table")
        if table_node and table_node.find_previous_sibling("div"):
            context_div = table_node.find_previous_sibling("div")

    # 2. 提取表格核心组件
    original_table = soup.find("table")
    if not original_table:
        return [long_html_content]

    caption = original_table.find("caption")

    header_rows = []
    thead = original_table.find("thead")
    if thead:
        header_rows = thead.find_all("tr")
    else:
        header_rows = original_table.find_all("tr")[:1]

    # 3. 准备数据行
    tbody = original_table.find("tbody")
    if tbody:
        data_rows = tbody.find_all("tr")
    else:
        all_rows = original_table.find_all("tr")
        data_rows = [row for row in all_rows if row not in header_rows]

    # 4. 计算固定开销（用于 token 模式）
    fixed_parts = []
    if context_div:
        fixed_parts.append(str(context_div))
    if caption:
        fixed_parts.append(str(caption))
    for h_row in header_rows:
        fixed_parts.append(str(h_row))
    fixed_overhead = estimate_tokens("".join(fixed_parts))

    chunks = []
    current_chunk_data = []
    current_chunk_tokens = 0

    def should_split(row_count, row_tokens):
        """判断是否应该切分"""
        if max_tokens_per_chunk is not None:
            # Token 模式：检查累计 token 是否超限
            return (
                current_chunk_tokens + row_tokens + fixed_overhead
            ) > max_tokens_per_chunk
        else:
            # 行数模式
            return row_count >= max_rows_per_chunk

    def build_chunk(data_rows_for_chunk):
        """组装一个 chunk"""
        new_soup = BeautifulSoup("<div></div>", "html.parser")
        wrapper_div = new_soup.div

        if context_div:
            wrapper_div.append(copy.copy(context_div))

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

        # 检查是否需要先切分（当前 chunk 非空且加入新行会超限）
        if current_chunk_data and should_split(len(current_chunk_data), row_tokens):
            chunks.append(build_chunk(current_chunk_data))
            current_chunk_data = []
            current_chunk_tokens = 0

        current_chunk_data.append(row)
        current_chunk_tokens += row_tokens

        # 最后一行，收尾
        if i == len(data_rows) - 1 and current_chunk_data:
            chunks.append(build_chunk(current_chunk_data))

    return chunks


def process_and_merge_html(file_path_str, separator="!!!_CHUNK_BREAK_!!!"):
    """
    读取文件 -> 切分 -> 合并 -> 保存
    """
    source_path = Path(file_path_str)

    if not source_path.exists():
        print(f"❌ 错误：找不到文件 '{source_path}'")
        return

    print(f"📂 正在读取: {source_path.name}")
    content = source_path.read_text(encoding="utf-8")

    # 1. 执行切分逻辑
    # 建议 max_rows_per_chunk 设置为 5-10，保证每个 chunk 不会因为加上表头和context后超过 Token 限制
    chunks = distribute_assets_and_chunk(content, max_rows_per_chunk=2)

    print(f"🔪 切分完成：共生成 {len(chunks)} 个片段")

    # 2. 执行合并逻辑
    # 我们在分隔符前后加换行符，确保结构清晰，不会粘连 HTML 标签
    formatted_separator = f"\n\n{separator}\n\n"
    merged_content = formatted_separator.join(chunks)

    # 3. 构建输出路径
    # 例子: input.html -> input_chunk_merged.html
    new_filename = source_path.stem + "_chunk_merged" + source_path.suffix
    output_path = source_path.with_name(new_filename)

    # 4. 写入文件
    try:
        output_path.write_text(merged_content, encoding="utf-8")
        print(f"✅ 合并成功！文件已保存至: {output_path.absolute()}")
        print(f"🔑 使用的分隔符: {separator}")
    except IOError as e:
        print(f"❌ 写入失败: {e}")


# --- 主程序入口 ---
if __name__ == "__main__":
    # 请将此处修改为你那个“已经处理过Context的长HTML”文件路径
    my_input_file = "Files\Excel\本国子目注释调整表.html"

    # 运行
    process_and_merge_html(my_input_file)
