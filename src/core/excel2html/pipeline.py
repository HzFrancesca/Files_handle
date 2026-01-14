"""
Excel 转 HTML 完整流水线
输入 Excel 文件 -> 生成增强 HTML -> 切分为 Chunks
"""

import argparse
from dataclasses import dataclass
from pathlib import Path

from loguru import logger

from ..models import ChunkConfig, ConversionResult, SplitMode
from .chunker import HtmlChunker
from .converter import ExcelToHtmlConverter


@dataclass
class ConversionPipeline:
    """Excel 转 HTML 转换流水线"""

    keywords: list[str] | None = None
    max_rows_per_chunk: int | None = None
    target_tokens: int = 1024
    separator: str = "!!!_CHUNK_BREAK_!!!"

    def run(self, excel_path: Path) -> ConversionResult | None:
        """执行完整的转换流水线"""
        source_path = Path(excel_path) if not isinstance(excel_path, Path) else excel_path

        if not source_path.exists():
            logger.error(f"找不到文件 '{source_path}'")
            return None

        self._log_start(source_path)

        # 第一步：Excel -> HTML
        html_path = self._convert_to_html(source_path)
        if not html_path:
            return None

        # 第二步：HTML -> Chunks
        chunk_result = self._chunk_html(html_path, source_path)
        if not chunk_result:
            return None

        return chunk_result

    def _log_start(self, source_path: Path) -> None:
        """记录开始日志"""
        logger.info("=" * 50)
        logger.info(f"🚀 开始处理流水线: {source_path.name}")
        logger.info("=" * 50)

    def _convert_to_html(self, source_path: Path) -> Path | None:
        """执行 Excel 到 HTML 转换"""
        logger.info("📌 第一步：Excel 转 HTML（增强版）")

        converter = ExcelToHtmlConverter(keywords=self.keywords)
        html_path = converter.convert(source_path)

        if not html_path:
            logger.error("流水线中断：HTML 转换失败")
            return None

        return html_path

    def _chunk_html(self, html_path: Path, source_path: Path) -> ConversionResult | None:
        """执行 HTML 切分"""
        logger.info("📌 第二步：HTML 切分为 Chunks")

        html_content = html_path.read_text(encoding="utf-8")

        config = self._build_chunk_config()
        chunker = HtmlChunker(config=config)
        result = chunker.chunk(html_content)

        self._log_chunk_result(result)

        # 保存结果
        chunk_path = source_path.parent / f"{source_path.stem}.html"
        return self._save_chunks(chunk_path, result, html_path)

    def _build_chunk_config(self) -> ChunkConfig:
        """构建切分配置"""
        if self.max_rows_per_chunk is None:
            logger.info(f"📊 使用 token 模式，目标每 chunk ≤ {self.target_tokens} tokens")
            return ChunkConfig(
                split_mode=SplitMode.BY_TOKENS,
                max_tokens=self.target_tokens,
                separator=self.separator,
            )
        else:
            logger.info(f"📊 使用行数模式，每 chunk {self.max_rows_per_chunk} 行")
            return ChunkConfig(
                split_mode=SplitMode.BY_ROWS,
                max_rows=self.max_rows_per_chunk,
                separator=self.separator,
            )

    def _log_chunk_result(self, result) -> None:
        """记录切分结果"""
        stats = result.stats
        logger.info(f"🔪 切分完成：共生成 {stats.total_chunks} 个片段")
        logger.info(
            f"📊 Token 统计: 最小={stats.min_token_count}, "
            f"最大={stats.max_token_count}, 平均={stats.avg_token_count:.1f}"
        )

        if result.warnings:
            logger.warning(f"有 {len(result.warnings)} 个片段超过 token 限制：")
            for w in result.warnings:
                logger.warning(
                    f"   - 片段 #{w.chunk_index}: {w.actual_tokens} tokens (超出 {w.overflow})"
                )
                logger.warning(f"     原因: {w.reason}")

    def _save_chunks(self, chunk_path: Path, result, html_path: Path) -> ConversionResult | None:
        """保存切分结果"""
        formatted_separator = f"\n\n{self.separator}\n\n"
        merged_content = formatted_separator.join(result.chunks)

        try:
            chunk_path.write_text(merged_content, encoding="utf-8")
            logger.info(f"✅ Chunk 文件已保存: {chunk_path.absolute()}")
        except OSError as e:
            logger.error(f"写入 Chunk 文件失败: {e}")
            return None

        self._log_completion(html_path, chunk_path, result)

        return ConversionResult(
            html_path=html_path,
            chunk_path=chunk_path,
            chunk_count=len(result.chunks),
            status_message="处理完成",
            success=True,
        )

    def _log_completion(self, html_path: Path, chunk_path: Path, result) -> None:
        """记录完成日志"""
        logger.info("=" * 50)
        logger.info("🎉 流水线执行完成！")
        logger.info(f"   📄 中间结果 (HTML): {html_path}")
        logger.info(f"   📄 最终结果 (Chunks): {chunk_path}")
        logger.info(f"   🔢 Chunk 数量: {len(result.chunks)}")
        logger.info(f"   🔑 分隔符: {self.separator}")
        logger.info("=" * 50)


def run_pipeline(
    excel_path: str,
    keywords: list[str] | None = None,
    max_rows_per_chunk: int | None = None,
    target_tokens: int = 1024,
    separator: str = "!!!_CHUNK_BREAK_!!!",
) -> dict | None:
    """执行完整的 Excel -> HTML -> Chunks 流水线（兼容旧接口）"""
    pipeline = ConversionPipeline(
        keywords=keywords,
        max_rows_per_chunk=max_rows_per_chunk,
        target_tokens=target_tokens,
        separator=separator,
    )

    result = pipeline.run(Path(excel_path))

    if result is None:
        return None

    return {
        "html_path": str(result.html_path),
        "chunk_path": str(result.chunk_path),
        "chunk_count": result.chunk_count,
    }


def main() -> None:
    """命令行入口"""
    parser = argparse.ArgumentParser(
        description="Excel 转换流水线（支持 HTML/MD/CSV 格式）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python pipeline.py input.xlsx
  python pipeline.py input.xlsx -f md
  python pipeline.py input.xlsx -f csv --delimiter ";"
  python pipeline.py input.xlsx -k "财务报表" "年度收入"
  python pipeline.py input.xlsx -t 1024
  python pipeline.py input.xlsx -r 5
  python pipeline.py input.xlsx -t 2048 -s "---SPLIT---"
        """,
    )
    parser.add_argument("excel_file", nargs="+", help="要转换的 Excel 文件路径（支持多个）")
    parser.add_argument(
        "-f",
        "--format",
        choices=["html", "md", "csv"],
        default="html",
        help="输出格式（默认: html）",
    )
    parser.add_argument("-k", "--keywords", nargs="+", help="关键检索词（用于幽灵标题）")
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
    # CSV 特定参数
    parser.add_argument(
        "--delimiter",
        default=",",
        help="CSV 分隔符（默认: 逗号）",
    )
    parser.add_argument(
        "--encoding",
        default="utf-8",
        help="CSV 编码（默认: utf-8）",
    )

    args = parser.parse_args()

    # 批量处理多个文件
    for excel_file in args.excel_file:
        if args.format == "html":
            run_pipeline(
                excel_path=excel_file,
                keywords=args.keywords,
                max_rows_per_chunk=args.max_rows,
                target_tokens=args.target_tokens,
                separator=args.separator,
            )
        else:
            from ..unified_pipeline import run_unified_pipeline
            run_unified_pipeline(
                excel_path=excel_file,
                output_format=args.format,
                keywords=args.keywords,
                max_rows_per_chunk=args.max_rows,
                target_tokens=args.target_tokens,
                separator=args.separator,
                csv_delimiter=args.delimiter,
                csv_encoding=args.encoding,
            )


if __name__ == "__main__":
    main()
