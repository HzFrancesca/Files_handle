"""
Excel 转 Markdown 转换器
将 Excel 文件转换为 Markdown 表格格式

功能：
1. 生成有效的 Markdown 表格语法
2. 保留降维表头结构
3. 处理合并单元格（重复值）
4. 转义 Markdown 特殊字符
5. 包含 RAG 上下文元数据
6. 注释提取和分发（与 HTML 一致）
"""

import json
import re
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
        features = "表头降维 ✓ | 合并单元格 ✓ | 特殊字符转义 ✓ | 注释提取 ✓"
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

        # 提取注释
        footer_notes, _ = self._detect_footer_notes(sheet, header_rows)
        notes_dict = self._parse_notes_with_keys(footer_notes)
        header_text = " ".join(flattened_headers.values())
        header_note_refs = self._extract_note_references(header_text)

        # 元数据块（YAML front matter 格式）
        if self.include_metadata:
            lines.append("---")
            lines.append(f"source: {filename}")
            lines.append(f"sheet: {sheet.title}")
            if self.keywords:
                lines.append(f"keywords: {', '.join(self.keywords)}")
            # 添加注释元数据
            if notes_dict:
                header_notes = {k: v for k, v in notes_dict.items() if k in header_note_refs}
                other_notes = {k: v for k, v in notes_dict.items() if k not in header_note_refs}
                notes_meta = {"header_notes": header_notes, "conditional_notes": other_notes}
                notes_json = json.dumps(notes_meta, ensure_ascii=False)
                lines.append(f"notes_meta: {notes_json}")
            lines.append("---")
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

    def _parse_notes_with_keys(self, notes_list: list[str]) -> dict[str, str]:
        """解析注释列表，提取注释编号和内容"""
        notes_dict: dict[str, str] = {}
        for note in notes_list:
            note = note.strip()
            if not note:
                continue
            split_pattern = r"(?=\[注[\d、,，]+\]|\[备注\d*\]|\[说明\d*\])"
            parts = re.split(split_pattern, note)
            parts = [p.strip() for p in parts if p.strip()]
            if len(parts) > 1:
                for part in parts:
                    self._parse_single_note(part, notes_dict)
            else:
                self._parse_single_note(note, notes_dict)
        return notes_dict

    def _parse_single_note(self, note: str, notes_dict: dict[str, str]) -> None:
        """解析单个注释"""
        note = note.strip()
        if not note:
            return
        multi_num_match = re.match(r"^\[(注)([\d、,，]+)\]", note)
        if multi_num_match:
            prefix = multi_num_match.group(1)
            nums_str = multi_num_match.group(2)
            nums = re.split(r"[、,，]", nums_str)
            for num in nums:
                num = num.strip()
                if num:
                    notes_dict[f"{prefix}{num}"] = note
            return
        bracket_match = re.match(r"^\[([注备说][注明意]?\d*)\]", note)
        if bracket_match:
            notes_dict[bracket_match.group(1)] = note
            return
        paren_match = re.match(r"^[（\(]([注备说][注明意]?\d*)[）\)]", note)
        if paren_match:
            notes_dict[paren_match.group(1)] = note
            return
        plain_match = re.match(r"^(注\d*|备注\d*|说明\d*|注意\d*)[：:．.、]?\s*", note)
        if plain_match:
            key_match = re.match(r"^(注\d*|备注\d*|说明\d*|注意\d*)", note)
            if key_match:
                notes_dict[key_match.group(1)] = note
            return
        if note[0] in "*※●◆△▲":
            notes_dict[note[0]] = note
            return
        notes_dict[note[:10]] = note

    def _extract_note_references(self, text: str) -> set[str]:
        """从文本中提取注释引用"""
        refs: set[str] = set()
        multi_refs = re.findall(r"\[(注)([\d、,，]+)\]", text)
        for prefix, nums_str in multi_refs:
            nums = re.split(r"[、,，]", nums_str)
            for num in nums:
                num = num.strip()
                if num:
                    refs.add(f"{prefix}{num}")
        bracket_refs = re.findall(r"\[(注\d*|备注\d*|说明\d*|注意\d*)\]", text)
        refs.update(bracket_refs)
        superscript_refs = re.findall(r"[^\[](注\d+)(?:[：:）\)]|$|\s)", text)
        refs.update(superscript_refs)
        if "*" in text:
            refs.add("*")
        if "※" in text:
            refs.add("※")
        return refs

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
