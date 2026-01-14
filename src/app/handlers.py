"""
业务处理器 - Excel 转换和预览
使用类封装状态，消除全局变量
"""

import shutil
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

from loguru import logger

from src.core.models import OutputFormat, ProcessingState, SplitMode
from src.core.unified_pipeline import UnifiedPipeline


@dataclass
class ExcelProcessHandler:
    """Excel 处理器"""

    state: ProcessingState = field(default_factory=ProcessingState)

    def process(
        self,
        excel_file,
        output_format: str,
        keywords_text: str,
        split_mode: str,
        max_rows: int,
        target_tokens: int,
        enable_min_tokens: bool,
        min_tokens: int,
        token_strategy: str,
        separator: str,
    ) -> tuple[str | None, str | None, str]:
        """处理 Excel 文件"""
        if excel_file is None:
            return None, None, "⚠️ 请先上传 Excel 文件"

        keywords = self._parse_keywords(keywords_text)
        temp_dir = Path(tempfile.mkdtemp())

        try:
            result = self._execute_conversion(
                excel_file,
                temp_dir,
                output_format,
                keywords,
                split_mode,
                max_rows,
                target_tokens,
                enable_min_tokens,
                min_tokens,
                token_strategy,
                separator,
            )
            return result
        except Exception as e:
            logger.exception("处理出错")
            self._reset_state()
            return None, None, f"❌ 处理出错: {e!s}"

    def _parse_keywords(self, keywords_text: str) -> list[str] | None:
        """解析关键词"""
        if not keywords_text.strip():
            return None
        return [k.strip() for k in keywords_text.split(",") if k.strip()]

    def _execute_conversion(
        self,
        excel_file,
        temp_dir: Path,
        output_format: str,
        keywords: list[str] | None,
        split_mode: str,
        max_rows: int,
        target_tokens: int,
        enable_min_tokens: bool,
        min_tokens: int,
        token_strategy: str,
        separator: str,
    ) -> tuple[str | None, str | None, str]:
        """执行转换流程"""
        # 复制文件到临时目录
        source_path = Path(excel_file.name)
        temp_excel = temp_dir / source_path.name
        shutil.copy(excel_file.name, temp_excel)

        # 解析输出格式
        fmt = OutputFormat(output_format)

        # 使用统一流水线
        pipeline = UnifiedPipeline(
            output_format=fmt,
            keywords=keywords,
            max_rows_per_chunk=max_rows if split_mode == SplitMode.BY_ROWS else None,
            target_tokens=target_tokens,
            separator=separator,
        )

        result = pipeline.run(temp_excel)

        if not result or not result.success:
            return None, None, "❌ 转换失败，请检查文件格式"

        # 重命名输出文件
        ext_map = {
            OutputFormat.HTML: ".html",
            OutputFormat.MARKDOWN: ".md",
        }
        ext = ext_map[fmt]

        # HTML/MD 格式有中间结果和最终结果
        middle_output_name = f"{source_path.stem}_middle{ext}"
        final_output_name = f"{source_path.stem}{ext}"
        middle_path = temp_dir / middle_output_name
        final_path = temp_dir / final_output_name

        if result.output_path and result.output_path.exists():
            # 避免复制到自身
            if result.output_path.resolve() != middle_path.resolve():
                shutil.copy(result.output_path, middle_path)
            else:
                middle_path = result.output_path
        if result.chunk_path and result.chunk_path.exists():
            # 避免复制到自身
            if result.chunk_path.resolve() != final_path.resolve():
                shutil.copy(result.chunk_path, final_path)
            else:
                final_path = result.chunk_path

        self.state.html_path = middle_path
        self.state.chunk_path = final_path
        self.state.output_format = fmt

        status = self._build_status_message(
            source_path,
            fmt,
            keywords,
            split_mode,
            enable_min_tokens,
            min_tokens,
            token_strategy,
            result.chunk_count,
            separator,
            result.chunk_stats,
        )

        return str(middle_path), str(final_path), status

    def _build_status_message(
        self,
        source_path: Path,
        output_format: OutputFormat,
        keywords: list[str] | None,
        split_mode: str,
        enable_min_tokens: bool,
        min_tokens: int,
        token_strategy: str,
        chunk_count: int,
        separator: str,
        chunk_stats=None,
    ) -> str:
        """构建状态消息"""
        format_names = {
            OutputFormat.HTML: "HTML",
            OutputFormat.MARKDOWN: "Markdown",
        }

        min_token_info = ""
        if split_mode == SplitMode.BY_TOKENS and enable_min_tokens:
            min_token_info = f"\n最小 Token：{min_tokens}，策略：{token_strategy}"

        chunking_info = f"\n切分模式：{split_mode}{min_token_info}"
        chunking_info += f"\n生成 Chunks：{chunk_count} 个"
        if chunk_stats:
            chunking_info += f"\nToken 统计：最小={chunk_stats.min_token_count}, 最大={chunk_stats.max_token_count}, 平均={chunk_stats.avg_token_count:.1f}"
        chunking_info += f"\n分隔符：{separator}"

        return f"""✅ 处理完成

源文件：{source_path.name}
输出格式：{format_names[output_format]}
关键词：{", ".join(keywords) if keywords else "无"}{chunking_info}"""

    def _reset_state(self) -> None:
        """重置状态"""
        self.state.html_path = None
        self.state.chunk_path = None
        self.state.output_format = OutputFormat.HTML

    def get_html_preview(self) -> str | None:
        """获取中间结果预览内容"""
        if self.state.html_path and self.state.html_path.exists():
            content = self.state.html_path.read_text(encoding="utf-8")
            is_markdown = self.state.output_format == OutputFormat.MARKDOWN
            return self._wrap_preview_html("中间结果预览", content, is_markdown)
        return None

    def get_chunk_preview(self) -> str | None:
        """获取最终结果预览内容"""
        if self.state.chunk_path and self.state.chunk_path.exists():
            content = self.state.chunk_path.read_text(encoding="utf-8")
            is_markdown = self.state.output_format == OutputFormat.MARKDOWN

            return self._wrap_preview_html("Chunks 预览", content, is_markdown)
        return None

    def _wrap_preview_html(self, title: str, content: str, is_markdown: bool = False) -> str:
        """包装预览 HTML"""
        if is_markdown:
            # Markdown 内容用 <pre> 包裹保留格式
            escaped_content = content.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            body_content = f'<pre class="markdown-preview">{escaped_content}</pre>'
        else:
            body_content = content

        return f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>{title}</title>
    <style>
        body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 40px; max-width: 1200px; margin: auto; }}
        table {{ border-collapse: collapse; width: 100%; margin: 20px 0; }}
        th, td {{ border: 1px solid #ddd; padding: 12px; text-align: left; }}
        th {{ background: #f5f5f5; font-weight: 600; }}
        tr:hover {{ background: #fafafa; }}
        .rag-context {{ background: #e3f2fd; padding: 16px; border-radius: 8px; margin-bottom: 20px; color: #1565c0; }}
        caption {{ font-size: 0.9rem; color: #666; margin-bottom: 12px; }}
        .chunk {{ margin-bottom: 20px; }}
        pre {{ background: #f5f5f5; padding: 16px; border-radius: 8px; overflow-x: auto; font-family: monospace; }}
        .markdown-preview {{ white-space: pre-wrap; word-wrap: break-word; line-height: 1.6; font-size: 14px; }}
    </style>
</head>
<body>
    <h2>📄 {title}</h2>
    {body_content}
</body>
</html>"""


# 全局处理器实例（用于 Gradio 回调）
_handler: ExcelProcessHandler | None = None


def _get_handler() -> ExcelProcessHandler:
    """获取处理器实例"""
    global _handler
    if _handler is None:
        _handler = ExcelProcessHandler()
    return _handler


def process_excel(
    excel_file,
    output_format: str,
    keywords_text: str,
    split_mode: str,
    max_rows: int,
    target_tokens: int,
    enable_min_tokens: bool,
    min_tokens: int,
    token_strategy: str,
    separator: str,
) -> tuple[str | None, str | None, str]:
    """处理 Excel 文件"""
    return _get_handler().process(
        excel_file,
        output_format,
        keywords_text,
        split_mode,
        max_rows,
        target_tokens,
        enable_min_tokens,
        min_tokens,
        token_strategy,
        separator,
    )


def get_html_preview() -> str | None:
    """获取中间结果预览"""
    return _get_handler().get_html_preview()


def get_chunk_preview() -> str | None:
    """获取最终结果预览"""
    return _get_handler().get_chunk_preview()
