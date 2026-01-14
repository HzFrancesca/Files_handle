"""
应用配置管理模块
使用 pydantic-settings 统一管理环境变量和配置
"""

from pathlib import Path

from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """应用配置"""

    model_config = SettingsConfigDict(
        env_prefix="EXCEL2HTML_",
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    # 临时文件目录
    temp_dir: Path = Field(default=Path("./temp"))

    # Token 相关配置
    default_max_tokens: int = Field(default=512, ge=64, le=8192)
    default_min_tokens: int = Field(default=256, ge=64, le=4096)
    default_max_rows: int = Field(default=8, ge=1, le=100)

    # 日志配置
    log_level: str = Field(default="INFO")
    log_format: str = Field(
        default="<green>{time:YYYY-MM-DD HH:mm:ss}</green> | "
        "<level>{level: <8}</level> | "
        "<cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> | "
        "<level>{message}</level>"
    )

    # 默认分隔符
    default_separator: str = Field(default="!!!_CHUNK_BREAK_!!!")


# 全局配置实例（懒加载）
_settings: Settings | None = None


def get_settings() -> Settings:
    """获取配置实例（单例模式）"""
    global _settings
    if _settings is None:
        _settings = Settings()
    return _settings
