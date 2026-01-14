"""
统一转换流水线
支持 HTML、Markdown、CSV 多种输出格式
"""

from dataclasses import dataclass
from pathlib import Path

from loguru import logger

from .base_converter import BaseExcelConverter
from .models import ChunkConfig, ConversionResult, OutputFormat, SplitMode


@dataclass
class UnifiedPipeline:
    """统一转换流水线"""

    output_format: OutputFormat = OutputFormat.HTML
    keywords: list[str] | None = None
    max_rows_per_chunk: int | None = None
    target_tokens: int = 1024
    separator: str = "!!!_CHUNK_BREAK_!!!"

    # MD 特定选项
    md_include_metadata: bool = True

    def run(self, excel_path: Path) -> ConversionResult | None:
        """执行转换流水线"""
        source_path = Path(excel_path) if not isinstance(excel_path, Path) else excel_path

        if not source_path.exists():
            logger.error(f"找不到文件 '{source_path}'")
            return None

        self._log_start(source_path)

        # 第一步：转换
        converter = self._create_converter()
        output_path = converter.convert(source_path)

        if not output_path:
            logger.error("流水线中断：转换失败")
            return None

        # 第二步：切分
        return self._process_with_chunking(output_path, source_path)

    def _log_start(self, source_path: Path) -> None:
        """记录开始日志"""
        logger.info("=" * 50)
        logger.info(f"🚀 开始处理流水线: {source_path.name}")
        logger.info(f"📄 输出格式: {self.output_format.value.upper()}")
        logger.info("=" * 50)

    def _create_converter(self) -> BaseExcelConverter:
        """创建对应格式的转换器"""
        if self.output_format == OutputFormat.HTML:
            from .excel2html.converter import ExcelToHtmlConverter
            return ExcelToHtmlConverter(keywords=self.keywords)
        else:
            from .excel2md.converter import MarkdownConverter
            return MarkdownConverter(
                keywords=self.keywords,
                include_metadata=self.md_include_metadata,
            )

    def _process_with_chunking(self, output_path: Path, source_path: Path) -> ConversionResult | None:
        """处理需要切分的格式"""
        logger.info("📌 第二步：切分为 Chunks")

        content = output_path.read_text(encoding="utf-8")
        config = self._build_chunk_config()

        if self.output_format == OutputFormat.HTML:
            from .excel2html.chunker import HtmlChunker
            chunker = HtmlChunker(config=config)
        else:
            from .excel2md.chunker import MarkdownChunker
            chunker = MarkdownChunker(config=config)

        result = chunker.chunk(content)
        self._log_chunk_result(result)

        # 保存结果
        ext = ".html" if self.output_format == OutputFormat.HTML else ".md"
        chunk_path = source_path.parent / f"{source_path.stem}{ext}"
        return self._save_chunks(chunk_path, result, output_path)

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

    def _save_chunks(self, chunk_path: Path, result, output_path: Path) -> ConversionResult | None:
        """保存切分结果"""
        formatted_separator = f"\n\n{self.separator}\n\n"
        merged_content = formatted_separator.join(result.chunks)

        try:
            chunk_path.write_text(merged_content, encoding="utf-8")
            logger.info(f"✅ Chunk 文件已保存: {chunk_path.absolute()}")
        except OSError as e:
            logger.error(f"写入 Chunk 文件失败: {e}")
            return None

        self._log_completion(output_path, chunk_path, len(result.chunks))

        return ConversionResult(
            output_path=output_path,
            chunk_path=chunk_path,
            chunk_count=len(result.chunks),
            status_message="处理完成",
            output_format=self.output_format,
            success=True,
            chunk_stats=result.stats,
        )

    def _log_completion(self, output_path: Path, chunk_path: Path, chunk_count: int) -> None:
        """记录完成日志"""
        logger.info("=" * 50)
        logger.info("🎉 流水线执行完成！")
        logger.info(f"   📄 中间结果: {output_path}")
        logger.info(f"   📄 最终结果: {chunk_path}")
        logger.info(f"   🔢 Chunk 数量: {chunk_count}")
        logger.info(f"   📝 输出格式: {self.output_format.value.upper()}")
        logger.info("=" * 50)


def run_unified_pipeline(
    excel_path: str,
    output_format: str = "html",
    keywords: list[str] | None = None,
    max_rows_per_chunk: int | None = None,
    target_tokens: int = 1024,
    separator: str = "!!!_CHUNK_BREAK_!!!",
) -> dict | None:
    """执行统一流水线（兼容函数接口）"""
    format_map = {
        "html": OutputFormat.HTML,
        "md": OutputFormat.MARKDOWN,
        "markdown": OutputFormat.MARKDOWN,
    }

    fmt = format_map.get(output_format.lower(), OutputFormat.HTML)

    pipeline = UnifiedPipeline(
        output_format=fmt,
        keywords=keywords,
        max_rows_per_chunk=max_rows_per_chunk,
        target_tokens=target_tokens,
        separator=separator,
    )

    result = pipeline.run(Path(excel_path))

    if result is None:
        return None

    return {
        "output_path": str(result.output_path),
        "chunk_path": str(result.chunk_path),
        "chunk_count": result.chunk_count,
        "output_format": result.output_format.value,
    }
