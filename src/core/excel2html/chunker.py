"""
HTML 切分器模块
将长 HTML 表格切分为多个 chunk，并分发全局资产
"""

import copy
import json
import re
from dataclasses import dataclass, field

import tiktoken
from bs4 import BeautifulSoup

from ..models import ChunkConfig, ChunkResult, ChunkStats, ChunkWarning, SplitMode, TokenStrategy

# 类型别名
type RowList = list
type NotesDict = dict[str, str]


@dataclass
class HtmlChunker:
    """HTML 切分器"""

    config: ChunkConfig
    _tokenizer: tiktoken.Encoding | None = field(default=None, init=False, repr=False)

    def chunk(self, html_content: str) -> ChunkResult:
        """执行切分"""
        soup = BeautifulSoup(html_content, "html.parser")

        # 提取全局资产
        context_div = self._extract_context_div(soup)
        header_notes, conditional_notes = self._extract_notes_meta(soup)

        # 提取表格组件
        original_table = soup.find("table")
        if not original_table:
            return self._single_chunk_result(html_content)

        caption = original_table.find("caption")
        header_rows = self._extract_header_rows(original_table)
        data_rows = self._extract_data_rows(original_table, header_rows)

        # 规范化合并单元格
        header_rows = self._normalize_table_spans(header_rows, soup)
        data_rows = self._normalize_table_spans(data_rows, soup)

        # 计算固定开销
        base_overhead = self._calculate_base_overhead(context_div, caption, header_rows)
        header_text = " ".join(str(row) for row in header_rows)

        # 执行切分
        chunks, warnings, token_counts = self._split_rows(
            data_rows,
            context_div,
            caption,
            header_rows,
            original_table,
            header_notes,
            conditional_notes,
            header_text,
            base_overhead,
        )

        stats = self._build_stats(token_counts, warnings, base_overhead)
        return ChunkResult(chunks=chunks, warnings=warnings, stats=stats)

    def _get_tokenizer(self) -> tiktoken.Encoding:
        """懒加载 tokenizer"""
        if self._tokenizer is None:
            self._tokenizer = tiktoken.get_encoding("cl100k_base")
        return self._tokenizer

    def _estimate_tokens(self, text: str) -> int:
        """计算 token 数量"""
        return len(self._get_tokenizer().encode(text))

    def _extract_context_div(self, soup: BeautifulSoup):
        """提取上下文 div"""
        context_div = soup.find("div", class_="rag-context")
        if not context_div:
            table_node = soup.find("table")
            if table_node and table_node.find_previous_sibling("div"):
                context_div = table_node.find_previous_sibling("div")
        return context_div

    def _extract_notes_meta(self, soup: BeautifulSoup) -> tuple[NotesDict, NotesDict]:
        """提取注释元数据"""
        notes_meta_script = soup.find("script", class_="table-notes-meta")
        header_notes: NotesDict = {}
        conditional_notes: NotesDict = {}

        if notes_meta_script:
            try:
                notes_meta = json.loads(notes_meta_script.string)
                header_notes = notes_meta.get("header_notes", {})
                conditional_notes = notes_meta.get("conditional_notes", {})
            except (json.JSONDecodeError, AttributeError):
                pass

        return header_notes, conditional_notes

    def _extract_header_rows(self, table) -> RowList:
        """提取表头行"""
        thead = table.find("thead")
        if thead:
            return thead.find_all("tr")
        return table.find_all("tr")[:1]

    def _extract_data_rows(self, table, header_rows: RowList) -> RowList:
        """提取数据行（排除注释行）"""
        tbody = table.find("tbody")
        if tbody:
            all_body_rows = tbody.find_all("tr")
            return [row for row in all_body_rows if "table-note-row" not in row.get("class", [])]

        all_rows = table.find_all("tr")
        return [row for row in all_rows if row not in header_rows]

    def _normalize_table_spans(self, rows: RowList, soup: BeautifulSoup) -> RowList:
        """处理表格的 rowspan/colspan，展开合并单元格"""
        if not rows:
            return []

        occupied = self._build_occupied_matrix(rows)
        return self._rebuild_rows(rows, occupied, soup)

    def _build_occupied_matrix(self, rows: RowList) -> dict:
        """构建单元格占用矩阵"""
        occupied: dict = {}

        for row_idx, row in enumerate(rows):
            if row_idx not in occupied:
                occupied[row_idx] = {}

            col_idx = 0
            cells = row.find_all(["td", "th"])

            for cell in cells:
                while col_idx in occupied[row_idx]:
                    col_idx += 1

                rowspan = int(cell.get("rowspan", 1))
                colspan = int(cell.get("colspan", 1))

                for r_offset in range(rowspan):
                    target_row = row_idx + r_offset
                    if target_row not in occupied:
                        occupied[target_row] = {}
                    for c_offset in range(colspan):
                        target_col = col_idx + c_offset
                        is_origin = r_offset == 0 and c_offset == 0
                        occupied[target_row][target_col] = (cell, is_origin, rowspan, colspan)

                col_idx += colspan

        return occupied

    def _rebuild_rows(self, rows: RowList, occupied: dict, soup: BeautifulSoup) -> RowList:
        """根据占用矩阵重建行"""
        normalized_rows = []

        for row_idx, row in enumerate(rows):
            new_row = soup.new_tag("tr")
            for attr, value in row.attrs.items():
                new_row[attr] = value

            if row_idx not in occupied:
                normalized_rows.append(new_row)
                continue

            max_col = max(occupied[row_idx].keys()) if occupied[row_idx] else -1

            for col_idx in range(max_col + 1):
                cell = self._create_normalized_cell(occupied, row_idx, col_idx, soup)
                new_row.append(cell)

            normalized_rows.append(new_row)

        return normalized_rows

    def _create_normalized_cell(
        self, occupied: dict, row_idx: int, col_idx: int, soup: BeautifulSoup
    ):
        """创建规范化的单元格"""
        if col_idx not in occupied[row_idx]:
            return soup.new_tag("td")

        orig_cell, is_origin, _, _ = occupied[row_idx][col_idx]

        if is_origin:
            new_cell = copy.copy(orig_cell)
            if "rowspan" in new_cell.attrs:
                del new_cell["rowspan"]
            if "colspan" in new_cell.attrs:
                del new_cell["colspan"]
            return new_cell

        # 填充单元格
        fill_cell = soup.new_tag(orig_cell.name)
        for child in orig_cell.children:
            fill_cell.append(copy.copy(child))
        for attr, value in orig_cell.attrs.items():
            if attr not in ("rowspan", "colspan"):
                fill_cell[attr] = value

        existing_class = fill_cell.get("class", [])
        if isinstance(existing_class, str):
            existing_class = existing_class.split()
        fill_cell["class"] = existing_class + ["span-fill"]

        return fill_cell

    def _calculate_base_overhead(self, context_div, caption, header_rows: RowList) -> int:
        """计算基础固定开销"""
        fixed_parts = []
        if context_div:
            fixed_parts.append(str(context_div))
        if caption:
            fixed_parts.append(str(caption))
        for h_row in header_rows:
            fixed_parts.append(str(h_row))
        return self._estimate_tokens("".join(fixed_parts))

    def _single_chunk_result(self, html_content: str) -> ChunkResult:
        """返回单个 chunk 的结果"""
        stats = ChunkStats(total_chunks=1, oversized_chunks=0)
        return ChunkResult(chunks=[html_content], warnings=[], stats=stats)

    def _split_rows(
        self,
        data_rows: RowList,
        context_div,
        caption,
        header_rows: RowList,
        original_table,
        header_notes: NotesDict,
        conditional_notes: NotesDict,
        header_text: str,
        base_overhead: int,
    ) -> tuple[list[str], list[ChunkWarning], list[int]]:
        """执行行切分"""
        chunks: list[str] = []
        warnings: list[ChunkWarning] = []
        token_counts: list[int] = []

        current_chunk_data: RowList = []
        current_chunk_tokens = 0

        for i, row in enumerate(data_rows):
            row_tokens = self._estimate_tokens(str(row))

            if current_chunk_data and self._should_split(
                current_chunk_data,
                row_tokens,
                row,
                header_notes,
                conditional_notes,
                header_text,
                base_overhead,
                current_chunk_tokens,
            ):
                # 输出当前 chunk
                final_total = self._calculate_chunk_total(
                    current_chunk_data,
                    header_notes,
                    conditional_notes,
                    header_text,
                    base_overhead,
                    current_chunk_tokens,
                )
                token_counts.append(final_total)
                chunks.append(
                    self._build_chunk(
                        current_chunk_data,
                        context_div,
                        caption,
                        header_rows,
                        original_table,
                        header_notes,
                        conditional_notes,
                        header_text,
                    )
                )
                current_chunk_data = []
                current_chunk_tokens = 0

            current_chunk_data.append(row)
            current_chunk_tokens += row_tokens

            # 检查超限
            warning = self._check_overflow(
                current_chunk_data,
                header_notes,
                conditional_notes,
                header_text,
                base_overhead,
                current_chunk_tokens,
                len(chunks),
            )
            if warning:
                warnings.append(warning)
                token_counts.append(warning.actual_tokens)
                chunks.append(
                    self._build_chunk(
                        current_chunk_data,
                        context_div,
                        caption,
                        header_rows,
                        original_table,
                        header_notes,
                        conditional_notes,
                        header_text,
                    )
                )
                current_chunk_data = []
                current_chunk_tokens = 0

            # 最后一行
            if i == len(data_rows) - 1 and current_chunk_data:
                final_total = self._calculate_chunk_total(
                    current_chunk_data,
                    header_notes,
                    conditional_notes,
                    header_text,
                    base_overhead,
                    current_chunk_tokens,
                )
                token_counts.append(final_total)
                chunks.append(
                    self._build_chunk(
                        current_chunk_data,
                        context_div,
                        caption,
                        header_rows,
                        original_table,
                        header_notes,
                        conditional_notes,
                        header_text,
                    )
                )

        return chunks, warnings, token_counts

    def _should_split(
        self,
        pending_rows: RowList,
        row_tokens: int,
        new_row,
        header_notes: NotesDict,
        conditional_notes: NotesDict,
        header_text: str,
        base_overhead: int,
        current_tokens: int,
    ) -> bool:
        """判断是否应该切分"""
        if self.config.max_tokens is not None:
            current_total = self._calculate_chunk_total(
                pending_rows,
                header_notes,
                conditional_notes,
                header_text,
                base_overhead,
                current_tokens,
            )

            test_rows = pending_rows + [new_row]
            notes_overhead = self._calculate_notes_overhead(
                test_rows, header_notes, conditional_notes, header_text
            )
            potential_total = current_tokens + row_tokens + base_overhead + notes_overhead

            if potential_total > self.config.max_tokens:
                return True

            return (
                self.config.min_tokens is not None
                and self.config.token_strategy.value == "prefer_min"
                and current_total >= self.config.min_tokens
            )

        return len(pending_rows) >= (self.config.max_rows or 8)

    def _calculate_chunk_total(
        self,
        rows: RowList,
        header_notes: NotesDict,
        conditional_notes: NotesDict,
        header_text: str,
        base_overhead: int,
        current_tokens: int,
    ) -> int:
        """计算 chunk 总 token 数"""
        notes_overhead = self._calculate_notes_overhead(
            rows, header_notes, conditional_notes, header_text
        )
        return current_tokens + base_overhead + notes_overhead

    def _calculate_notes_overhead(
        self,
        rows: RowList,
        header_notes: NotesDict,
        conditional_notes: NotesDict,
        header_text: str,
    ) -> int:
        """计算注释开销"""
        if not header_notes and not conditional_notes:
            return 0

        chunk_text = " ".join(str(row) for row in rows)
        all_text = chunk_text + " " + header_text
        chunk_refs = self._extract_note_references(all_text)

        actual_notes = []
        seen_notes: set[str] = set()

        for note in header_notes.values():
            if note not in seen_notes:
                actual_notes.append(note)
                seen_notes.add(note)

        for key, note in conditional_notes.items():
            if key in chunk_refs and note not in seen_notes:
                actual_notes.append(note)
                seen_notes.add(note)

        if not actual_notes:
            return 0

        notes_text = " | ".join(actual_notes)
        return self._estimate_tokens(f" 【表格注释】{notes_text}")

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

        bracket_refs = re.findall(r"\[(注\s*\d*|备注\s*\d*|说明\s*\d*|注意\s*\d*)\s*\]", text)
        refs.update(ref.replace(" ", "") for ref in bracket_refs)

        superscript_refs = re.findall(r"[^\[](注\d+)(?:[：:）\)]|$|\s)", text)
        refs.update(superscript_refs)

        if "*" in text:
            refs.add("*")
        if "※" in text:
            refs.add("※")

        return refs

    def _check_overflow(
        self,
        rows: RowList,
        header_notes: NotesDict,
        conditional_notes: NotesDict,
        header_text: str,
        base_overhead: int,
        current_tokens: int,
        chunk_index: int,
    ) -> ChunkWarning | None:
        """检查是否超限"""
        if self.config.max_tokens is None:
            return None

        current_total = self._calculate_chunk_total(
            rows, header_notes, conditional_notes, header_text, base_overhead, current_tokens
        )

        if current_total > self.config.max_tokens:
            reason = (
                "单行数据 + 固定开销 + 注释超过 token 限制"
                if len(rows) == 1
                else "累积数据超过 token 限制"
            )
            return ChunkWarning(
                chunk_index=chunk_index,
                actual_tokens=current_total,
                limit=self.config.max_tokens,
                overflow=current_total - self.config.max_tokens,
                row_count=len(rows),
                reason=reason,
            )

        return None

    def _build_chunk(
        self,
        data_rows: RowList,
        context_div,
        caption,
        header_rows: RowList,
        original_table,
        header_notes: NotesDict,
        conditional_notes: NotesDict,
        header_text: str,
    ) -> str:
        """组装一个 chunk"""
        new_soup = BeautifulSoup("<div></div>", "html.parser")
        wrapper_div = new_soup.div

        # 提取注释引用
        chunk_text = " ".join(str(row) for row in data_rows)
        all_text = chunk_text + " " + header_text
        chunk_refs = self._extract_note_references(all_text)

        matched_notes = self._collect_matched_notes(header_notes, conditional_notes, chunk_refs)

        # 添加上下文
        if context_div:
            new_context = copy.copy(context_div)
            if matched_notes:
                notes_text = " | ".join(matched_notes)
                new_context.string = (new_context.get_text() or "") + f" 【表格注释】{notes_text}"
            wrapper_div.append(new_context)

        # 构建表格
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
        for d_row in data_rows:
            new_tbody.append(copy.copy(d_row))
        new_table.append(new_tbody)

        wrapper_div.append(new_table)
        return str(wrapper_div)

    def _collect_matched_notes(
        self,
        header_notes: NotesDict,
        conditional_notes: NotesDict,
        chunk_refs: set[str],
    ) -> list[str]:
        """收集匹配的注释"""
        matched_notes: list[str] = []
        seen_notes: set[str] = set()

        for note in header_notes.values():
            if note not in seen_notes:
                matched_notes.append(note)
                seen_notes.add(note)

        for key, note in conditional_notes.items():
            if key in chunk_refs and note not in seen_notes:
                matched_notes.append(note)
                seen_notes.add(note)

        return matched_notes

    def _build_stats(
        self,
        token_counts: list[int],
        warnings: list[ChunkWarning],
        base_overhead: int,
    ) -> ChunkStats:
        """构建统计信息"""
        stats = ChunkStats(
            total_chunks=len(token_counts),
            oversized_chunks=len(warnings),
            token_counts=token_counts,
            max_token_count=max(token_counts) if token_counts else 0,
            min_token_count=min(token_counts) if token_counts else 0,
            avg_token_count=(sum(token_counts) / len(token_counts) if token_counts else 0.0),
            base_fixed_overhead=base_overhead,
        )

        if self.config.max_tokens:
            stats.token_limit = self.config.max_tokens
        if self.config.min_tokens:
            stats.min_token_limit = self.config.min_tokens
            stats.token_strategy = self.config.token_strategy

        return stats


def distribute_assets_and_chunk(
    long_html_content: str,
    max_rows_per_chunk: int | None = None,
    max_tokens_per_chunk: int | None = None,
    min_tokens_per_chunk: int | None = None,
    token_strategy: str = "prefer_max",
) -> dict:
    """切分 HTML 并分发资产（兼容旧接口）"""
    if max_rows_per_chunk is None and max_tokens_per_chunk is None:
        max_rows_per_chunk = 8

    strategy = (
        TokenStrategy.PREFER_MAX if token_strategy == "prefer_max" else TokenStrategy.PREFER_MIN
    )

    config = ChunkConfig(
        split_mode=(SplitMode.BY_TOKENS if max_tokens_per_chunk else SplitMode.BY_ROWS),
        max_tokens=max_tokens_per_chunk,
        min_tokens=min_tokens_per_chunk,
        max_rows=max_rows_per_chunk,
        token_strategy=strategy,
    )

    chunker = HtmlChunker(config=config)
    result = chunker.chunk(long_html_content)

    # 转换为旧格式
    return {
        "chunks": result.chunks,
        "warnings": [
            {
                "chunk_index": w.chunk_index,
                "actual_tokens": w.actual_tokens,
                "limit": w.limit,
                "overflow": w.overflow,
                "row_count": w.row_count,
                "reason": w.reason,
            }
            for w in result.warnings
        ],
        "stats": {
            "total_chunks": result.stats.total_chunks,
            "oversized_chunks": result.stats.oversized_chunks,
            "token_counts": result.stats.token_counts,
            "max_token_count": result.stats.max_token_count,
            "min_token_count": result.stats.min_token_count,
            "avg_token_count": result.stats.avg_token_count,
            "base_fixed_overhead": result.stats.base_fixed_overhead,
            "token_limit": result.stats.token_limit,
            "min_token_limit": result.stats.min_token_limit,
            "token_strategy": (
                result.stats.token_strategy.value if result.stats.token_strategy else None
            ),
        },
    }


def estimate_tokens(text: str) -> int:
    """估算 token 数量（兼容旧接口）"""
    tokenizer = tiktoken.get_encoding("cl100k_base")
    return len(tokenizer.encode(text))
