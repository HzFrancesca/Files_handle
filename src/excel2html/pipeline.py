"""
Excel 转 HTML 完整流水线
输入 Excel 文件 -> 生成增强 HTML -> 切分为 Chunks

中间结果命名: 原文件名_converted.html
最终结果命名: 原文件名.html（与原文件同名，方便直接使用）
"""

from pathlib import Path
import argparse

from excel2html_openpyxl_enhanced import convert_excel_to_html
from html2chunk import distribute_assets_and_chunk, estimate_tokens


def estimate_rows_for_token_limit(html_content: str, target_tokens: int = 512) -> int:
    """
    根据目标 token 数估算每个 chunk 应该包含多少行

    参数:
        html_content: 完整的 HTML 内容
        target_tokens: 目标 token 数（默认 512）

    返回:
        建议的 max_rows_per_chunk
    """
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(html_content, "html.parser")

    # 计算固定开销（context + caption + thead）
    fixed_parts = []

    context_div = soup.find("div", class_="rag-context")
    if context_div:
        fixed_parts.append(str(context_div))

    table = soup.find("table")
    if not table:
        return 8  # 无表格，返回默认值

    caption = table.find("caption")
    if caption:
        fixed_parts.append(str(caption))

    thead = table.find("thead")
    if thead:
        fixed_parts.append(str(thead))

    fixed_overhead = estimate_tokens("".join(fixed_parts))

    # 计算每行平均 token
    tbody = table.find("tbody")
    if tbody:
        data_rows = tbody.find_all("tr")
    else:
        all_rows = table.find_all("tr")
        data_rows = all_rows[1:] if len(all_rows) > 1 else all_rows

    if not data_rows:
        return 8

    total_row_tokens = sum(estimate_tokens(str(row)) for row in data_rows)
    avg_tokens_per_row = total_row_tokens / len(data_rows)

    # 计算可用于数据行的 token 数
    available_tokens = target_tokens - fixed_overhead

    if available_tokens <= 0 or avg_tokens_per_row <= 0:
        return 1  # 极端情况，每个 chunk 只放 1 行

    suggested_rows = int(available_tokens / avg_tokens_per_row)

    # 限制在合理范围 [1, 20]
    return max(1, min(suggested_rows, 20))


def run_pipeline(
    excel_path: str,
    keywords: list = None,
    max_rows_per_chunk: int = None,
    target_tokens: int = 512,
    separator: str = "!!!_CHUNK_BREAK_!!!",
):
    """
    执行完整的 Excel -> HTML -> Chunks 流水线

    参数:
        excel_path: Excel 文件路径
        keywords: 关键检索词列表（用于幽灵标题）
        max_rows_per_chunk: 每个 chunk 的最大行数（如果指定，优先使用）
        target_tokens: 目标 token 数（当 max_rows 未指定时，自动计算行数）
        separator: chunk 之间的分隔符

    返回:
        dict: {
            'html_path': 中间 HTML 文件路径,
            'chunk_path': 最终 chunk 文件路径,
            'chunk_count': chunk 数量
        }
    """
    source_path = Path(excel_path)

    if not source_path.exists():
        print(f"❌ 错误：找不到文件 '{source_path}'")
        return None

    print("=" * 50)
    print(f"🚀 开始处理流水线: {source_path.name}")
    print("=" * 50)

    # === 第一步：Excel -> HTML ===
    print("\n📌 第一步：Excel 转 HTML（增强版）")
    html_path = convert_excel_to_html(
        excel_path=str(source_path),
        keywords=keywords,
        output_path=None,  # 默认保存到同目录
    )

    if not html_path:
        print("❌ 流水线中断：HTML 转换失败")
        return None

    # === 第二步：HTML -> Chunks ===
    print("\n📌 第二步：HTML 切分为 Chunks")
    html_content = Path(html_path).read_text(encoding="utf-8")

    # 自动计算或使用指定的行数
    if max_rows_per_chunk is None:
        # 使用 token 模式：逐行累加，精确控制每个 chunk 的 token 数
        print(f"📊 使用 token 模式，目标每 chunk ≤ {target_tokens} tokens")
        chunks = distribute_assets_and_chunk(
            html_content,
            max_rows_per_chunk=None,
            max_tokens_per_chunk=target_tokens
        )
    else:
        # 使用行数模式
        print(f"📊 使用行数模式，每 chunk {max_rows_per_chunk} 行")
        chunks = distribute_assets_and_chunk(
            html_content,
            max_rows_per_chunk=max_rows_per_chunk,
            max_tokens_per_chunk=None
        )
    print(f"🔪 切分完成：共生成 {len(chunks)} 个片段")

    # 保存 chunk 结果（最终结果与原文件同名，方便直接使用）
    chunk_path = source_path.with_suffix(".html")

    formatted_separator = f"\n\n{separator}\n\n"
    merged_content = formatted_separator.join(chunks)

    try:
        chunk_path.write_text(merged_content, encoding="utf-8")
        print(f"✅ Chunk 文件已保存: {chunk_path.absolute()}")
    except IOError as e:
        print(f"❌ 写入 Chunk 文件失败: {e}")
        return None

    # === 完成 ===
    print("\n" + "=" * 50)
    print("🎉 流水线执行完成！")
    print(f"   📄 中间结果 (HTML): {html_path}")
    print(f"   📄 最终结果 (Chunks): {chunk_path}")
    print(f"   🔢 Chunk 数量: {len(chunks)}")
    print(f"   🔑 分隔符: {separator}")
    print(f"   💡 提示: 最终结果与原文件同名，可直接使用")
    print("=" * 50)

    return {
        "html_path": html_path,
        "chunk_path": str(chunk_path),
        "chunk_count": len(chunks),
    }


def main():
    """命令行入口"""
    parser = argparse.ArgumentParser(
        description="Excel 转 HTML 完整流水线（增强版 + Chunk 切分）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python pipeline.py input.xlsx
  python pipeline.py input.xlsx -k "财务报表" "年度收入"
  python pipeline.py input.xlsx -t 512          # 基于 512 tokens 自动计算行数
  python pipeline.py input.xlsx -r 5            # 固定每 chunk 5 行
  python pipeline.py input.xlsx -t 1024 -s "---SPLIT---"
        """,
    )
    parser.add_argument("excel_file", help="要转换的 Excel 文件路径")
    parser.add_argument(
        "-k", "--keywords", nargs="+", help="关键检索词（用于幽灵标题）"
    )
    parser.add_argument(
        "-r",
        "--max-rows",
        type=int,
        default=None,
        help="每个 chunk 的最大数据行数（指定后忽略 -t 参数）",
    )
    parser.add_argument(
        "-t",
        "--target-tokens",
        type=int,
        default=512,
        help="目标 token 数，自动计算行数（默认: 512）",
    )
    parser.add_argument(
        "-s",
        "--separator",
        default="!!!_CHUNK_BREAK_!!!",
        help="chunk 之间的分隔符（默认: !!!_CHUNK_BREAK_!!!）",
    )

    args = parser.parse_args()

    run_pipeline(
        excel_path=args.excel_file,
        keywords=args.keywords,
        max_rows_per_chunk=args.max_rows,
        target_tokens=args.target_tokens,
        separator=args.separator,
    )


if __name__ == "__main__":
    main()
