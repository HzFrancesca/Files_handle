"""
Excel 转 HTML 完整流水线
输入 Excel 文件 -> 生成增强 HTML -> 切分为 Chunks
"""

from pathlib import Path
import argparse

from .excel2html_openpyxl_enhanced import convert_excel_to_html
from .html2chunk import distribute_assets_and_chunk, estimate_tokens


def estimate_rows_for_token_limit(html_content: str, target_tokens: int = 1024) -> int:
    """根据目标 token 数估算每个 chunk 应该包含多少行"""
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(html_content, "html.parser")

    fixed_parts = []

    context_div = soup.find("div", class_="rag-context")
    if context_div:
        fixed_parts.append(str(context_div))

    table = soup.find("table")
    if not table:
        return 8

    caption = table.find("caption")
    if caption:
        fixed_parts.append(str(caption))

    thead = table.find("thead")
    if thead:
        fixed_parts.append(str(thead))

    fixed_overhead = estimate_tokens("".join(fixed_parts))

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

    available_tokens = target_tokens - fixed_overhead

    if available_tokens <= 0 or avg_tokens_per_row <= 0:
        return 1

    suggested_rows = int(available_tokens / avg_tokens_per_row)

    return max(1, min(suggested_rows, 20))


def run_pipeline(
    excel_path: str,
    keywords: list = None,
    max_rows_per_chunk: int = None,
    target_tokens: int = 1024,
    separator: str = "!!!_CHUNK_BREAK_!!!",
):
    """执行完整的 Excel -> HTML -> Chunks 流水线"""
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
        output_path=None,
    )

    if not html_path:
        print("❌ 流水线中断：HTML 转换失败")
        return None

    # === 第二步：HTML -> Chunks ===
    print("\n📌 第二步：HTML 切分为 Chunks")
    html_content = Path(html_path).read_text(encoding="utf-8")

    if max_rows_per_chunk is None:
        print(f"📊 使用 token 模式，目标每 chunk ≤ {target_tokens} tokens")
        result = distribute_assets_and_chunk(
            html_content,
            max_rows_per_chunk=None,
            max_tokens_per_chunk=target_tokens
        )
    else:
        print(f"📊 使用行数模式，每 chunk {max_rows_per_chunk} 行")
        result = distribute_assets_and_chunk(
            html_content,
            max_rows_per_chunk=max_rows_per_chunk,
            max_tokens_per_chunk=None
        )
    
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

    chunk_path = source_path.parent / f"{source_path.stem}.html"

    formatted_separator = f"\n\n{separator}\n\n"
    merged_content = formatted_separator.join(chunks)

    try:
        chunk_path.write_text(merged_content, encoding="utf-8")
        print(f"✅ Chunk 文件已保存: {chunk_path.absolute()}")
    except IOError as e:
        print(f"❌ 写入 Chunk 文件失败: {e}")
        return None

    print("\n" + "=" * 50)
    print("🎉 流水线执行完成！")
    print(f"   📄 中间结果 (HTML): {html_path}")
    print(f"   📄 最终结果 (Chunks): {chunk_path}")
    print(f"   🔢 Chunk 数量: {len(chunks)}")
    print(f"   🔑 分隔符: {separator}")
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
  python pipeline.py input.xlsx -t 1024
  python pipeline.py input.xlsx -r 5
  python pipeline.py input.xlsx -t 2048 -s "---SPLIT---"
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
        help="每个 chunk 的最大数据行数",
    )
    parser.add_argument(
        "-t",
        "--target-tokens",
        type=int,
        default=1024,
        help="目标 token 数（默认: 1024）",
    )
    parser.add_argument(
        "-s",
        "--separator",
        default="!!!_CHUNK_BREAK_!!!",
        help="chunk 之间的分隔符",
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
