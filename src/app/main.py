"""
应用主入口
"""

from .styles import CUSTOM_CSS
from .ui import create_ui


def create_app():
    """创建并返回 Gradio 应用"""
    return create_ui()


def run_app():
    """运行 Gradio 应用"""
    app = create_app()
    app.launch()


if __name__ == "__main__":
    run_app()
