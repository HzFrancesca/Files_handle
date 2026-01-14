"""
Excel 转 HTML 转换器 - RAG 增强版
将 Excel 文件转换为 RAG 优化的 HTML 格式

增强功能：
1. 上下文硬编码 - 注入文件名、Sheet名等元数据
2. 幽灵标题 - 添加同义词和关键检索词
3. 表头降维 - 把父级标题拼接到子级标题（针对多层表头）
4. 合并单元格智能处理
5. 注释提取和分发
"""

import json
import re
from dataclasses import dataclass
from pathlib import Path

from loguru import logger

from ..base_converter import BaseExcelConverter


@dataclass
class ExcelToHtmlConverter(BaseExcelConverter):
    """Excel 转 HTML 转换器"""

    def _get_file_extension(self) -> str:
        """返回输出文件扩展名"""
        return ".html"

    def _log_features(self) -> None:
        """记录启用的功能"""
        features = "上下文硬编码 ✓ | 表头降维 ✓ | 合并单元格 ✓ | 注释提取 ✓"
        if self.keywords:
            features += f" | 幽灵标题 ✓ ({len(self.keywords)}个关键词)"
        else:
            features += " | 幽灵标题 ✗ (未提供关键词)"
        logger.info(f"增强功能: {features}")

    def _format_sheet(
        self,
        sheet,
        filename: str,
        flattened_headers: dict[int, str],
        header_rows: int,
        data_end_row: int,
    ) -> str:
        """将单个 sheet 转换为 RAG 增强的 HTML 表格"""
        footer_notes, _ = self._detect_footer_notes(sheet, header_rows)
        notes_dict = self._parse_notes_with_keys(footer_notes)
        header_text = " ".join(flattened_headers.values())
        header_note_refs = self._extract_note_references(header_text)
        html_parts = self._build_html_parts(
            sheet, filename, flattened_headers, notes_dict,
            header_note_refs, header_rows, data_end_row,
        )
        return "\n".join(html_parts)

    def _join_sheets(self, sheet_contents: list[str]) -> str:
        """合并多个 sheet 的内容"""
        return "\n".join(sheet_contents)


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


    def _build_html_parts(
        self, sheet, filename: str, flattened_headers: dict[int, str],
        notes_dict: dict[str, str], header_note_refs: set[str],
        header_rows: int, data_end_row: int,
    ) -> list[str]:
        """构建 HTML 各部分"""
        html_parts: list[str] = []
        context_html = (
            f'<div class="rag-context">【文档上下文】来源：{filename} | 数据类型：表格数据</div>'
        )
        html_parts.append(context_html)
        if notes_dict:
            html_parts.append(self._build_notes_meta(notes_dict, header_note_refs))
        html_parts.append(
            f'<table border="1" style="border-collapse:collapse" '
            f'data-source="{filename}" data-sheet="{sheet.title}">'
        )
        if self.keywords:
            keyword_str = "，".join(self.keywords)
            html_parts.append(f"    <caption>关键检索词：{keyword_str}</caption>")
        html_parts.extend(self._build_thead(flattened_headers))
        html_parts.extend(self._build_tbody(sheet, header_rows, data_end_row))
        html_parts.append("</table>")
        return html_parts

    def _build_notes_meta(self, notes_dict: dict[str, str], header_note_refs: set[str]) -> str:
        """构建注释元数据"""
        header_notes = {k: v for k, v in notes_dict.items() if k in header_note_refs}
        other_notes = {k: v for k, v in notes_dict.items() if k not in header_note_refs}
        notes_meta = {"header_notes": header_notes, "conditional_notes": other_notes}
        notes_json = json.dumps(notes_meta, ensure_ascii=False)
        return f'<script type="application/json" class="table-notes-meta">{notes_json}</script>'

    def _build_thead(self, flattened_headers: dict[int, str]) -> list[str]:
        """构建表头"""
        parts = ["    <thead>", "        <tr>"]
        for col_idx in sorted(flattened_headers.keys()):
            flat_header = flattened_headers.get(col_idx, "")
            parts.append(f"            <th>{flat_header}</th>")
        parts.extend(["        </tr>", "    </thead>"])
        return parts

    def _build_tbody(self, sheet, header_rows: int, data_end_row: int) -> list[str]:
        """构建表体"""
        parts = ["    <tbody>"]
        for row_idx in range(header_rows + 1, sheet.max_row + 1):
            is_note_row = row_idx > data_end_row
            row_class = ' class="table-note-row"' if is_note_row else ""
            parts.append(f"        <tr{row_class}>")
            for col_idx in range(1, sheet.max_column + 1):
                cell_html = self._build_cell(sheet, row_idx, col_idx)
                if cell_html:
                    parts.append(cell_html)
            parts.append("        </tr>")
        parts.append("    </tbody>")
        return parts

    def _build_cell(self, sheet, row_idx: int, col_idx: int) -> str | None:
        """构建单元格"""
        info = self._merged_info.get((row_idx, col_idx))
        if info and info.skip:
            return None
        if info:
            span_attrs = []
            if info.rowspan > 1:
                span_attrs.append(f'rowspan="{info.rowspan}"')
            if info.colspan > 1:
                span_attrs.append(f'colspan="{info.colspan}"')
            span_str = " " + " ".join(span_attrs) if span_attrs else ""
            cell_content = info.value if info.value is not None else ""
        else:
            cell = sheet.cell(row=row_idx, column=col_idx)
            cell_content = self._format_cell_value(cell)
            span_str = ""
        return f"            <td{span_str}>{cell_content}</td>"


def convert_excel_to_html(
    excel_path: str,
    keywords: list[str] | None = None,
    output_path: str | None = None,
) -> str | None:
    """将单个 Excel 文件转换为 RAG 增强的 HTML（兼容旧接口）"""
    converter = ExcelToHtmlConverter(keywords=keywords)
    result = converter.convert(Path(excel_path), Path(output_path) if output_path else None)
    return str(result) if result else None
