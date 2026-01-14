"""
Markdown 切分器
将 Markdown 表格切分为多个 chunks，用于 RAG 场景
支持注释提取和分发（与 HTML chunker 一致）
"""

import json
import re
from dataclasses import dataclass

from loguru import logger

from ..models import ChunkConfig, ChunkResult, ChunkStats, ChunkWarning, SplitMode

# 类型别名
type NotesDict = dict[str, str]


def estimate_tokens(text: str) -> int:
    """估算文本的 token 数量（中文约2.5字符=1token）"""
    return int(len(text) / 2.5)


@dataclass
class MarkdownChunker:
    """Markdown 切分器"""

    config: ChunkConfig

    def chunk(self, md_content: str) -> ChunkResult:
        """切分 Markdown 内容"""
        # 解析 MD 结构
        front_matter, header_line, separator_line, data_lines, header_notes, conditional_notes = (
            self._parse_md_table(md_content)
        )

        if not header_line or not data_lines:
            # 无法解析，返回原内容作为单个 chunk
            return ChunkResult(
                chunks=[md_content],
                warnings=[],
                stats=ChunkStats(total_chunks=1, oversized_chunks=0),
            )

        # 提取表头文本用于注释匹配
        header_text = header_line

        # 计算固定开销（front matter + 表头）
        fixed_overhead = self._estimate_fixed_overhead(front_matter, header_line, separator_line)

        # 切分数据行
        chunks: list[str] = []
        warnings: list[ChunkWarning] = []
        current_chunk: list[str] = []
        current_tokens = fixed_overhead

        for i, line in enumerate(data_lines):
            line_tokens = estimate_tokens(line)

            # 计算加入新行后的注释开销
            test_chunk = current_chunk + [line]
            notes_overhead = self._calculate_notes_overhead(
                test_chunk, header_notes, conditional_notes, header_text
            )

            if current_chunk and self._should_split(
                len(current_chunk), current_tokens + line_tokens + notes_overhead
            ):
                chunks.append(
                    self._build_chunk(
                        front_matter, header_line, separator_line, current_chunk,
                        header_notes, conditional_notes, header_text
                    )
                )
                current_chunk = []
                current_tokens = fixed_overhead

            current_chunk.append(line)
            current_tokens += line_tokens

            # 检查单行是否超限
            if self.config.split_mode == SplitMode.BY_TOKENS and self.config.max_tokens:
                single_notes_overhead = self._calculate_notes_overhead(
                    [line], header_notes, conditional_notes, header_text
                )
                if line_tokens + fixed_overhead + single_notes_overhead > self.config.max_tokens:
                    warnings.append(
                        ChunkWarning(
                            chunk_index=len(chunks),
                            actual_tokens=line_tokens + fixed_overhead + single_notes_overhead,
                            limit=self.config.max_tokens,
                            overflow=line_tokens + fixed_overhead + single_notes_overhead - self.config.max_tokens,
                            row_count=1,
                            reason="单行数据 + 注释超过 token 限制",
                        )
                    )

        # 处理最后一个 chunk
        if current_chunk:
            chunks.append(
                self._build_chunk(
                    front_matter, header_line, separator_line, current_chunk,
                    header_notes, conditional_notes, header_text
                )
            )

        # 计算统计信息
        stats = self._calculate_stats(chunks, warnings, fixed_overhead)

        return ChunkResult(chunks=chunks, warnings=warnings, stats=stats)

    def _parse_md_table(
        self, md_content: str
    ) -> tuple[dict[str, str], str | None, str | None, list[str], NotesDict, NotesDict]:
        """解析 Markdown 表格结构，包括 YAML front matter"""
        lines = md_content.split("\n")

        front_matter: dict[str, str] = {}
        header_line: str | None = None
        separator_line: str | None = None
        data_lines: list[str] = []
        header_notes: NotesDict = {}
        conditional_notes: NotesDict = {}

        in_front_matter = False
        front_matter_done = False

        for line in lines:
            stripped = line.strip()

            # 解析 YAML front matter
            if stripped == "---" and not front_matter_done:
                if not in_front_matter:
                    in_front_matter = True
                    continue
                else:
                    in_front_matter = False
                    front_matter_done = True
                    continue

            if in_front_matter:
                if ":" in stripped:
                    key, _, value = stripped.partition(":")
                    front_matter[key.strip()] = value.strip()
                    # 提取注释元数据
                    if key.strip() == "notes_meta":
                        try:
                            notes_meta = json.loads(value.strip())
                            header_notes = notes_meta.get("header_notes", {})
                            conditional_notes = notes_meta.get("conditional_notes", {})
                        except json.JSONDecodeError:
                            pass
                continue

            # 兼容旧的 HTML 注释格式
            if stripped.startswith("<!--"):
                if "NOTES_META:" in stripped:
                    try:
                        json_start = stripped.index("NOTES_META:") + len("NOTES_META:")
                        json_end = stripped.rindex("-->")
                        json_str = stripped[json_start:json_end].strip()
                        notes_meta = json.loads(json_str)
                        header_notes = notes_meta.get("header_notes", {})
                        conditional_notes = notes_meta.get("conditional_notes", {})
                    except (ValueError, json.JSONDecodeError):
                        pass
                continue

            if not stripped:
                continue

            if stripped.startswith("|") and header_line is None:
                header_line = line
            elif stripped.startswith("|") and separator_line is None and "---" in stripped:
                separator_line = line
            elif stripped.startswith("|"):
                data_lines.append(line)

        return front_matter, header_line, separator_line, data_lines, header_notes, conditional_notes

    def _estimate_fixed_overhead(
        self, front_matter: dict[str, str], header: str | None, separator: str | None
    ) -> int:
        """估算固定开销的 token 数"""
        parts = []
        # front matter（不含 notes_meta）
        if front_matter:
            parts.append("---")
            for key, value in front_matter.items():
                if key != "notes_meta":
                    parts.append(f"{key}: {value}")
            parts.append("---")
        if header:
            parts.append(header)
        if separator:
            parts.append(separator)
        return estimate_tokens("\n".join(parts))

    def _should_split(self, row_count: int, total_tokens: int) -> bool:
        """判断是否应该切分"""
        if self.config.split_mode == SplitMode.BY_TOKENS and self.config.max_tokens:
            return total_tokens > self.config.max_tokens
        elif self.config.split_mode == SplitMode.BY_ROWS and self.config.max_rows:
            return row_count >= self.config.max_rows
        return False

    def _build_chunk(
        self,
        front_matter: dict[str, str],
        header: str | None,
        separator: str | None,
        data: list[str],
        header_notes: NotesDict,
        conditional_notes: NotesDict,
        header_text: str,
    ) -> str:
        """构建单个 chunk，包含匹配的注释"""
        parts: list[str] = []

        # 收集匹配的注释
        chunk_text = " ".join(data)
        all_text = chunk_text + " " + header_text
        chunk_refs = self._extract_note_references(all_text)
        matched_notes = self._collect_matched_notes(header_notes, conditional_notes, chunk_refs)

        # 构建 YAML front matter
        parts.append("---")
        for key, value in front_matter.items():
            if key != "notes_meta":
                parts.append(f"{key}: {value}")
        # 添加匹配的注释
        if matched_notes:
            notes_text = " | ".join(matched_notes)
            parts.append(f"notes: {notes_text}")
        parts.append("---")
        parts.append("")

        # 添加表头
        if header:
            parts.append(header)
        if separator:
            parts.append(separator)

        # 添加数据行
        parts.extend(data)

        return "\n".join(parts)

    def _calculate_notes_overhead(
        self,
        data_lines: list[str],
        header_notes: NotesDict,
        conditional_notes: NotesDict,
        header_text: str,
    ) -> int:
        """计算注释开销"""
        if not header_notes and not conditional_notes:
            return 0

        chunk_text = " ".join(data_lines)
        all_text = chunk_text + " " + header_text
        chunk_refs = self._extract_note_references(all_text)
        matched_notes = self._collect_matched_notes(header_notes, conditional_notes, chunk_refs)

        if not matched_notes:
            return 0

        notes_text = " | ".join(matched_notes)
        return estimate_tokens(f"<!-- 表格注释: {notes_text} -->")

    def _extract_note_references(self, text: str) -> set[str]:
        """从文本中提取注释引用"""
        refs: set[str] = set()

        # 处理 Markdown 转义字符
        text = text.replace(r"\[", "[").replace(r"\]", "]")

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

    def _collect_matched_notes(
        self,
        header_notes: NotesDict,
        conditional_notes: NotesDict,
        chunk_refs: set[str],
    ) -> list[str]:
        """收集匹配的注释"""
        matched_notes: list[str] = []
        seen_notes: set[str] = set()

        # 表头注释始终包含
        for note in header_notes.values():
            if note not in seen_notes:
                matched_notes.append(note)
                seen_notes.add(note)

        # 条件注释按引用匹配
        for key, note in conditional_notes.items():
            if key in chunk_refs and note not in seen_notes:
                matched_notes.append(note)
                seen_notes.add(note)

        return matched_notes

    def _calculate_stats(
        self, chunks: list[str], warnings: list[ChunkWarning], fixed_overhead: int
    ) -> ChunkStats:
        """计算切分统计信息"""
        token_counts = [estimate_tokens(chunk) for chunk in chunks]

        return ChunkStats(
            total_chunks=len(chunks),
            oversized_chunks=len(warnings),
            token_counts=token_counts,
            max_token_count=max(token_counts) if token_counts else 0,
            min_token_count=min(token_counts) if token_counts else 0,
            avg_token_count=sum(token_counts) / len(token_counts) if token_counts else 0.0,
            base_fixed_overhead=fixed_overhead,
            token_limit=self.config.max_tokens,
            min_token_limit=self.config.min_tokens,
            token_strategy=self.config.token_strategy,
        )


def chunk_markdown(
    md_content: str,
    max_tokens: int | None = 1024,
    max_rows: int | None = None,
    separator: str = "!!!_CHUNK_BREAK_!!!",
) -> ChunkResult:
    """切分 Markdown 内容（兼容函数接口）"""
    if max_rows is not None:
        config = ChunkConfig(
            split_mode=SplitMode.BY_ROWS,
            max_rows=max_rows,
            separator=separator,
        )
    else:
        config = ChunkConfig(
            split_mode=SplitMode.BY_TOKENS,
            max_tokens=max_tokens,
            separator=separator,
        )

    chunker = MarkdownChunker(config=config)
    return chunker.chunk(md_content)
