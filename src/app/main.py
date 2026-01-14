"""
应用主入口
配置 loguru 日志和加载设置
"""

import sys

from loguru import logger

from src.core.config import get_settings

from .ui import create_ui


def _configure_logging() -> None:
    """配置 loguru 日志"""
    settings = get_settings()

    # 移除默认 handler
    logger.remove()

    # 添加控制台输出
    logger.add(
        sys.stderr,
        format=settings.log_format,
        level=settings.log_level,
        colorize=True,
    )

    logger.info("日志系统已初始化")


def create_app():
    """创建并返回 Gradio 应用"""
    _configure_logging()
    return create_ui()


def run_app() -> None:
    """运行 Gradio 应用"""
    app = create_app()
    logger.info("启动 Gradio 应用")
    app.launch()


if __name__ == "__main__":
    run_app()
