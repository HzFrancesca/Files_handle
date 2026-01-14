"""
Markdown 切分器
将 Markdown 表格切分为多个 chunks，用于 RAG 场景
"""

from dataclasses import dataclass

from loguru import logger

from ..models import ChunkConfig, ChunkResult, ChunkStats, ChunkWarning, SplitMode


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
        metadata_lines, header_line, separator_line, data_lines = self._parse_md_table(md_content)

        if not header_line or not data_lines:
            # 无法解析，返回原内容作为单个 chunk
            return ChunkResult(
                chunks=[md_content],
                warnings=[],
                stats=ChunkStats(total_chunks=1, oversized_chunks=0),
            )

        # 计算固定开销
        fixed_overhead = self._estimate_fixed_overhead(metadata_lines, header_line, separator_line)

        # 切分数据行
        chunks: list[str] = []
        warnings: list[ChunkWarning] = []
        current_chunk: list[str] = []
        current_tokens = fixed_overhead

        for i, line in enumerate(data_lines):
            line_tokens = estimate_tokens(line)

            if current_chunk and self._should_split(len(current_chunk), current_tokens + line_tokens):
                chunks.append(self._build_chunk(metadata_lines, header_line, separator_line, current_chunk))
                current_chunk = []
                current_tokens = fixed_overhead

            current_chunk.append(line)
            current_tokens += line_tokens

            # 检查单行是否超限
            if self.config.split_mode == SplitMode.BY_TOKENS and self.config.max_tokens:
                if line_tokens + fixed_overhead > self.config.max_tokens:
                    warnings.append(
                        ChunkWarning(
                            chunk_index=len(chunks),
                            actual_tokens=line_tokens + fixed_overhead,
                            limit=self.config.max_tokens,
                            overflow=line_tokens + fixed_overhead - self.config.max_tokens,
                            row_count=1,
                            reason="单行数据超过 token 限制",
                        )
                    )

        # 处理最后一个 chunk
        if current_chunk:
            chunks.append(self._build_chunk(metadata_lines, header_line, separator_line, current_chunk))

        # 计算统计信息
        stats = self._calculate_stats(chunks, warnings, fixed_overhead)

        return ChunkResult(chunks=chunks, warnings=warnings, stats=stats)

    def _parse_md_table(self, md_content: str) -> tuple[list[str], str | None, str | None, list[str]]:
        """解析 Markdown 表格结构"""
        lines = md_content.split("\n")

        metadata_lines: list[str] = []
        header_line: str | None = None
        separator_line: str | None = None
        data_lines: list[str] = []

        for line in lines:
            stripped = line.strip()
            if not stripped:
                continue

            if stripped.startswith("<!--"):
                metadata_lines.append(line)
            elif stripped.startswith("|") and header_line is None:
                header_line = line
            elif stripped.startswith("|") and separator_line is None and "---" in stripped:
                separator_line = line
            elif stripped.startswith("|"):
                data_lines.append(line)

        return metadata_lines, header_line, separator_line, data_lines

    def _estimate_fixed_overhead(
        self, metadata: list[str], header: str | None, separator: str | None
    ) -> int:
        """估算固定开销的 token 数"""
        parts = metadata.copy()
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
        self, metadata: list[str], header: str | None, separator: str | None, data: list[str]
    ) -> str:
        """构建单个 chunk"""
        parts: list[str] = []

        # 添加元数据
        if metadata:
            parts.extend(metadata)
            parts.append("")

        # 添加表头
        if header:
            parts.append(header)
        if separator:
            parts.append(separator)

        # 添加数据行
        parts.extend(data)

        return "\n".join(parts)

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
