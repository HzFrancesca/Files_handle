"""
业务处理器 - Excel 转换和预览
使用类封装状态，消除全局变量
"""

import shutil
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

from loguru import logger

from src.core.excel2html.chunker import distribute_assets_and_chunk
from src.core.excel2html.converter import convert_excel_to_html
from src.core.models import ProcessingState, SplitMode


@dataclass
class ExcelProcessHandler:
    """Excel 处理器"""

    state: ProcessingState = field(default_factory=ProcessingState)

    def process(
        self,
        excel_file,
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

        # Excel -> HTML
        html_path = convert_excel_to_html(
            excel_path=str(temp_excel),
            keywords=keywords,
            output_path=None,
        )

        if not html_path:
            return None, None, "❌ Excel 转换失败，请检查文件格式"

        html_content = Path(html_path).read_text(encoding="utf-8")

        # HTML -> Chunks
        result = self._chunk_html(
            html_content,
            split_mode,
            max_rows,
            target_tokens,
            enable_min_tokens,
            min_tokens,
            token_strategy,
        )

        if isinstance(result, str):  # 错误消息
            return None, None, result

        chunks, warnings, stats = result

        # 保存结果
        return self._save_results(
            source_path,
            temp_dir,
            html_path,
            chunks,
            warnings,
            stats,
            keywords,
            split_mode,
            enable_min_tokens,
            min_tokens,
            token_strategy,
            separator,
        )

    def _chunk_html(
        self,
        html_content: str,
        split_mode: str,
        max_rows: int,
        target_tokens: int,
        enable_min_tokens: bool,
        min_tokens: int,
        token_strategy: str,
    ) -> tuple[list[str], list[dict], dict] | str:
        """切分 HTML"""
        if split_mode == SplitMode.BY_TOKENS:
            min_tokens_value = min_tokens if enable_min_tokens else None

            if min_tokens_value is not None and min_tokens_value >= target_tokens:
                return (
                    f"❌ 最小 Token 数 ({min_tokens_value}) 必须小于最大 Token 数 ({target_tokens})"
                )

            strategy = "prefer_max" if token_strategy == "接近最大值" else "prefer_min"

            result = distribute_assets_and_chunk(
                html_content,
                max_rows_per_chunk=None,
                max_tokens_per_chunk=target_tokens,
                min_tokens_per_chunk=min_tokens_value,
                token_strategy=strategy,
            )
        else:
            result = distribute_assets_and_chunk(
                html_content,
                max_rows_per_chunk=max_rows,
                max_tokens_per_chunk=None,
            )

        return result["chunks"], result["warnings"], result["stats"]

    def _save_results(
        self,
        source_path: Path,
        temp_dir: Path,
        html_path: str,
        chunks: list[str],
        warnings: list[dict],
        stats: dict,
        keywords: list[str] | None,
        split_mode: str,
        enable_min_tokens: bool,
        min_tokens: int,
        token_strategy: str,
        separator: str,
    ) -> tuple[str, str, str]:
        """保存结果文件"""
        formatted_separator = f"\n\n{separator}\n\n"
        merged_content = formatted_separator.join(chunks)

        html_output_name = f"{source_path.stem}_middle.html"
        chunk_output_name = f"{source_path.stem}.html"

        chunk_path = temp_dir / chunk_output_name
        chunk_path.write_text(merged_content, encoding="utf-8")

        html_final_path = temp_dir / html_output_name
        if str(html_path) != str(html_final_path):
            shutil.copy(html_path, html_final_path)
            html_path = str(html_final_path)

        self.state.html_path = Path(html_path)
        self.state.chunk_path = chunk_path

        status = self._build_status_message(
            source_path,
            keywords,
            split_mode,
            enable_min_tokens,
            min_tokens,
            token_strategy,
            chunks,
            stats,
            warnings,
            separator,
        )

        return html_path, str(chunk_path), status

    def _build_status_message(
        self,
        source_path: Path,
        keywords: list[str] | None,
        split_mode: str,
        enable_min_tokens: bool,
        min_tokens: int,
        token_strategy: str,
        chunks: list[str],
        stats: dict,
        warnings: list[dict],
        separator: str,
    ) -> str:
        """构建状态消息"""
        warning_text = ""
        if warnings:
            warning_text = f"\n⚠️ 警告：{len(warnings)} 个片段超过 token 限制"
            for w in warnings:
                warning_text += (
                    f"\n   - 片段 #{w['chunk_index']}: "
                    f"{w['actual_tokens']} tokens (超出 {w['overflow']})"
                )

        min_token_info = ""
        if split_mode == SplitMode.BY_TOKENS and enable_min_tokens:
            min_token_info = f"\n最小 Token：{min_tokens}，策略：{token_strategy}"

        return f"""✅ 处理完成

源文件：{source_path.name}
关键词：{", ".join(keywords) if keywords else "无"}
切分模式：{split_mode}{min_token_info}
生成 Chunks：{len(chunks)} 个
Token 统计：最小={stats["min_token_count"]}, 最大={stats["max_token_count"]}, 平均={stats["avg_token_count"]:.1f}
分隔符：{separator}{warning_text}"""

    def _reset_state(self) -> None:
        """重置状态"""
        self.state.html_path = None
        self.state.chunk_path = None

    def get_html_preview(self) -> str | None:
        """获取 HTML 预览内容"""
        if self.state.html_path and self.state.html_path.exists():
            content = self.state.html_path.read_text(encoding="utf-8")
            return self._wrap_preview_html("中间结果预览", content)
        return None

    def get_chunk_preview(self) -> str | None:
        """获取 Chunk 预览内容"""
        if self.state.chunk_path and self.state.chunk_path.exists():
            content = self.state.chunk_path.read_text(encoding="utf-8")
            content = content.replace(
                "!!!_CHUNK_BREAK_!!!",
                '</div><hr style="border: 2px dashed #2563eb; margin: 40px 0;">'
                '<div style="background:#f8f9fa; padding: 8px 16px; border-radius: 4px; '
                'color: #666; font-size: 0.85rem; margin-bottom: 20px;">📦 Chunk 分隔</div>'
                '<div class="chunk">',
            )
            return self._wrap_preview_html("Chunks 预览", f'<div class="chunk">{content}</div>')
        return None

    def _wrap_preview_html(self, title: str, content: str) -> str:
        """包装预览 HTML"""
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
    </style>
</head>
<body>
    <h2>📄 {title}</h2>
    {content}
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
    keywords_text: str,
    split_mode: str,
    max_rows: int,
    target_tokens: int,
    enable_min_tokens: bool,
    min_tokens: int,
    token_strategy: str,
    separator: str,
) -> tuple[str | None, str | None, str]:
    """处理 Excel 文件（兼容旧接口）"""
    return _get_handler().process(
        excel_file,
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
    """获取 HTML 预览（兼容旧接口）"""
    return _get_handler().get_html_preview()


def get_chunk_preview() -> str | None:
    """获取 Chunk 预览（兼容旧接口）"""
    return _get_handler().get_chunk_preview()
