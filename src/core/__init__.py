"""
Core 模块
"""

from .config import Settings, get_settings
from .excel2html import (
    ConversionPipeline,
    ExcelToHtmlConverter,
    HtmlChunker,
    convert_excel_to_html,
    distribute_assets_and_chunk,
    estimate_tokens,
    run_pipeline,
)
from .models import (
    ChunkConfig,
    ChunkResult,
    ChunkStats,
    ChunkWarning,
    ConversionResult,
    MergedCellInfo,
    ProcessingState,
    ProcessRequest,
    SplitMode,
    TableNote,
    TokenStrategy,
)

__all__ = [
    # 配置
    "Settings",
    "get_settings",
    # 转换器
    "ExcelToHtmlConverter",
    "HtmlChunker",
    "ConversionPipeline",
    # 数据模型
    "ChunkConfig",
    "ChunkResult",
    "ChunkStats",
    "ChunkWarning",
    "ConversionResult",
    "MergedCellInfo",
    "ProcessingState",
    "ProcessRequest",
    "SplitMode",
    "TableNote",
    "TokenStrategy",
    # 兼容函数
    "convert_excel_to_html",
    "distribute_assets_and_chunk",
    "estimate_tokens",
    "run_pipeline",
]
