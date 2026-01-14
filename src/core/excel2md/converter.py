"""
Excel 转 Markdown 转换器
将 Excel 文件转换为 Markdown 表格格式

功能：
1. 生成有效的 Markdown 表格语法
2. 保留降维表头结构
3. 处理合并单元格（重复值）
4. 转义 Markdown 特殊字符
5. 包含 RAG 上下文元数据
"""

from dataclasses import dataclass
from pathlib import Path

from loguru import logger

from ..base_converter import BaseExcelConverter


@dataclass
class MarkdownConverter(BaseExcelConverter):
    """Markdown 格式转换器"""

    include_metadata: bool = True
    sheet_separator: str = "\n\n---\n\n"

    def _get_file_extension(self) -> str:
        """返回输出文件扩展名"""
        return ".md"

    def _log_features(self) -> None:
        """记录启用的功能"""
        features = "表头降维 ✓ | 合并单元格 ✓ | 特殊字符转义 ✓"
        if self.include_metadata:
            features += " | RAG 元数据 ✓"
        if self.keywords:
            features += f" | 关键词 ✓ ({len(self.keywords)}个)"
        logger.info(f"MD 增强功能: {features}")

    def _format_sheet(
        self,
        sheet,
        filename: str,
        flattened_headers: dict[int, str],
        header_rows: int,
        data_end_row: int,
    ) -> str:
        """生成 Markdown 表格"""
        lines: list[str] = []

        # 元数据块（HTML 注释格式）
        if self.include_metadata:
            lines.append(f"<!-- RAG Context: {filename} | Sheet: {sheet.title} -->")
            if self.keywords:
                lines.append(f"<!-- Keywords: {', '.join(self.keywords)} -->")
            lines.append("")

        # 表头行
        headers = [flattened_headers.get(i, "") for i in range(1, sheet.max_column + 1)]
        escaped_headers = [self._escape_md(h) for h in headers]
        lines.append("| " + " | ".join(escaped_headers) + " |")

        # 分隔行
        lines.append("| " + " | ".join(["---"] * len(headers)) + " |")

        # 数据行
        for row_idx in range(header_rows + 1, data_end_row + 1):
            row_values = self._get_row_values(sheet, row_idx)
            escaped_values = [self._escape_md(v) for v in row_values]
            lines.append("| " + " | ".join(escaped_values) + " |")

        return "\n".join(lines)

    def _join_sheets(self, sheet_contents: list[str]) -> str:
        """合并多个 sheet 的内容"""
        return self.sheet_separator.join(sheet_contents)

    def _escape_md(self, text: str) -> str:
        """转义 Markdown 特殊字符"""
        if not text:
            return ""
        # 转义顺序很重要：先转义反斜杠，再转义其他字符
        text = text.replace("\\", "\\\\")
        special_chars = ["|", "*", "_", "`", "[", "]"]
        for char in special_chars:
            text = text.replace(char, "\\" + char)
        # 处理换行符，替换为空格
        text = text.replace("\n", " ").replace("\r", "")
        return text


def convert_excel_to_md(
    excel_path: str,
    keywords: list[str] | None = None,
    output_path: str | None = None,
    include_metadata: bool = True,
) -> str | None:
    """将单个 Excel 文件转换为 Markdown（兼容函数接口）"""
    converter = MarkdownConverter(keywords=keywords, include_metadata=include_metadata)
    result = converter.convert(Path(excel_path), Path(output_path) if output_path else None)
    return str(result) if result else None
