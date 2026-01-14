"""
数据模型定义模块
包含枚举、dataclass 和 Pydantic 模型
"""

from dataclasses import dataclass, field
from enum import StrEnum
from pathlib import Path

from pydantic import BaseModel, Field, field_validator

# ============ 枚举定义 ============


class SplitMode(StrEnum):
    """切分模式枚举"""

    BY_TOKENS = "按 Token 数"
    BY_ROWS = "按行数"


class TokenStrategy(StrEnum):
    """Token 切分策略"""

    PREFER_MAX = "prefer_max"
    PREFER_MIN = "prefer_min"


# ============ 内部数据模型 (dataclass) ============


@dataclass
class MergedCellInfo:
    """合并单元格信息"""

    value: str | None
    rowspan: int
    colspan: int
    is_origin: bool
    skip: bool


@dataclass
class TableNote:
    """表格注释"""

    key: str
    content: str
    is_header_note: bool = False


@dataclass
class ChunkConfig:
    """切分配置"""

    split_mode: SplitMode = SplitMode.BY_TOKENS
    max_tokens: int | None = 1024
    min_tokens: int | None = None
    max_rows: int | None = None
    token_strategy: TokenStrategy = TokenStrategy.PREFER_MAX
    separator: str = "!!!_CHUNK_BREAK_!!!"


@dataclass
class ChunkWarning:
    """切分警告信息"""

    chunk_index: int
    actual_tokens: int
    limit: int
    overflow: int
    row_count: int
    reason: str


@dataclass
class ChunkStats:
    """切分统计"""

    total_chunks: int
    oversized_chunks: int
    token_counts: list[int] = field(default_factory=list)
    max_token_count: int = 0
    min_token_count: int = 0
    avg_token_count: float = 0.0
    base_fixed_overhead: int = 0
    token_limit: int | None = None
    min_token_limit: int | None = None
    token_strategy: TokenStrategy | None = None


@dataclass
class ChunkResult:
    """切分结果"""

    chunks: list[str]
    warnings: list[ChunkWarning]
    stats: ChunkStats


@dataclass
class ConversionResult:
    """转换结果"""

    html_path: Path | None
    chunk_path: Path | None
    chunk_count: int
    status_message: str
    success: bool = True


@dataclass
class ProcessingState:
    """处理状态（替代全局变量）"""

    html_path: Path | None = None
    chunk_path: Path | None = None


# ============ 外部输入验证模型 (Pydantic) ============


class ProcessRequest(BaseModel):
    """处理请求（外部输入验证）"""

    keywords: list[str] = Field(default_factory=list)
    split_mode: SplitMode = Field(default=SplitMode.BY_TOKENS)
    max_rows: int = Field(default=8, ge=1, le=100)
    target_tokens: int = Field(default=1024, ge=64, le=8192)
    min_tokens: int | None = Field(default=None, ge=64)
    enable_min_tokens: bool = Field(default=False)
    token_strategy: str = Field(default="接近最大值")
    separator: str = Field(default="!!!_CHUNK_BREAK_!!!")

    @field_validator("keywords", mode="before")
    @classmethod
    def parse_keywords(cls, v: str | list | None) -> list[str]:
        """解析关键词"""
        if v is None:
            return []
        if isinstance(v, str):
            return [k.strip() for k in v.split(",") if k.strip()]
        return v

    def to_chunk_config(self) -> ChunkConfig:
        """转换为 ChunkConfig"""
        strategy = (
            TokenStrategy.PREFER_MAX
            if self.token_strategy == "接近最大值"
            else TokenStrategy.PREFER_MIN
        )

        if self.split_mode == SplitMode.BY_TOKENS:
            return ChunkConfig(
                split_mode=self.split_mode,
                max_tokens=self.target_tokens,
                min_tokens=self.min_tokens if self.enable_min_tokens else None,
                max_rows=None,
                token_strategy=strategy,
                separator=self.separator,
            )
        else:
            return ChunkConfig(
                split_mode=self.split_mode,
                max_tokens=None,
                min_tokens=None,
                max_rows=self.max_rows,
                token_strategy=strategy,
                separator=self.separator,
            )
