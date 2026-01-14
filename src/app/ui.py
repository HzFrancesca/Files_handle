"""
Gradio UI 组件定义
"""

import gradio as gr

from src.core.models import OutputFormat, SplitMode

from .handlers import get_chunk_preview, get_html_preview, process_excel
from .styles import CUSTOM_CSS


def _toggle_slider(mode: str) -> tuple:
    """切分模式切换时显示/隐藏对应滑块"""
    if mode == SplitMode.BY_TOKENS:
        return (
            gr.update(visible=True),  # target_tokens
            gr.update(visible=True),  # enable_min_tokens
            gr.update(visible=False),  # min_tokens
            gr.update(visible=False),  # max_rows
        )
    return (
        gr.update(visible=False),  # target_tokens
        gr.update(visible=False),  # enable_min_tokens
        gr.update(visible=False),  # min_tokens
        gr.update(visible=True),  # max_rows
    )


def _toggle_format_options(output_format: str) -> gr.update:
    """输出格式切换时显示/隐藏对应选项"""
    # HTML 和 MD 都支持切分选项
    return gr.update(visible=True)


def _toggle_min_tokens(enabled: bool) -> tuple:
    """启用最小 token 时显示滑块和策略选择"""
    return gr.update(visible=enabled), gr.update(visible=enabled)


def _create_input_column() -> tuple:
    """创建输入区域"""
    excel_input = gr.File(
        label="上传 Excel 文件",
        file_types=[".xlsx", ".xls"],
        elem_classes=["file-upload"],
    )

    gr.HTML('<div class="gap"></div>')

    # 输出格式选择
    output_format = gr.Radio(
        choices=["md", "html"],
        value="md",
        label="输出格式",
        elem_classes=["radio-group"],
    )

    gr.HTML('<div class="gap"></div>')

    keywords_input = gr.Textbox(
        label="关键检索词（可选）",
        placeholder="多个关键词用逗号分隔，如：财务报表, 年度收入, 利润分析",
        lines=3,
    )

    gr.HTML('<div class="gap"></div>')

    # 切分选项（HTML/MD 可用）
    with gr.Group(visible=True) as chunking_options:
        split_mode = gr.Radio(
            choices=[SplitMode.BY_TOKENS.value, SplitMode.BY_ROWS.value],
            value=SplitMode.BY_TOKENS.value,
            label="切分模式",
            elem_classes=["radio-group"],
        )

        gr.HTML('<div class="gap"></div>')

        target_tokens = gr.Slider(
            minimum=256,
            maximum=4096,
            value=512,
            step=128,
            label="最大 Token 数",
            visible=True,
        )

        enable_min_tokens = gr.Checkbox(
            label="启用最小 Token 限制",
            value=False,
            visible=True,
        )

        min_tokens = gr.Slider(
            minimum=64,
            maximum=2048,
            value=256,
            step=64,
            label="最小 Token 数",
            visible=False,
        )

        token_strategy = gr.Radio(
            choices=["接近最大值", "接近最小值"],
            value="接近最大值",
            label="切分策略",
            info="接近最大值：累加到接近 max 才切分 | 接近最小值：超过 min 就切分",
            visible=False,
        )

        max_rows = gr.Slider(
            minimum=1,
            maximum=20,
            value=8,
            step=1,
            label="每 Chunk 最大行数",
            visible=False,
        )

        gr.HTML('<div class="gap"></div>')

        separator_input = gr.Textbox(
            label="Chunk 分隔符",
            value="!!!_CHUNK_BREAK_!!!",
        )

    gr.HTML('<div class="gap"></div>')

    process_btn = gr.Button(
        "开始处理",
        variant="primary",
        elem_classes=["primary-btn"],
        size="lg",
    )

    return (
        excel_input,
        output_format,
        keywords_input,
        chunking_options,
        split_mode,
        target_tokens,
        enable_min_tokens,
        min_tokens,
        token_strategy,
        max_rows,
        separator_input,
        process_btn,
    )


def _create_output_column() -> tuple:
    """创建输出区域"""
    status_output = gr.Textbox(
        label="处理状态",
        lines=10,
        interactive=False,
        elem_classes=["status-box"],
        placeholder="处理结果将显示在这里...",
    )

    gr.HTML('<div class="gap"></div>')

    # 中间结果
    middle_result_label = gr.HTML('<p class="result-label">中间结果</p>')
    with gr.Row(elem_classes=["file-row"]):
        middle_output = gr.File(
            label=None,
            show_label=False,
            elem_classes=["file-download"],
        )
        middle_preview_btn = gr.Button(
            "预览",
            elem_classes=["preview-btn"],
            size="sm",
        )
        gr.HTML(
            '<span class="tooltip" title="预览为方便查看增加了美化，'
            '实际输出文件不含 CSS 样式，请以实际保存结果为准">ℹ️</span>'
        )

    gr.HTML('<div class="gap"></div>')

    # 最终结果
    final_result_label = gr.HTML('<p class="result-label">最终结果 (Chunks)</p>')
    with gr.Row(elem_classes=["file-row"]):
        chunk_output = gr.File(
            label=None,
            show_label=False,
            elem_classes=["file-download"],
        )
        chunk_preview_btn = gr.Button(
            "预览",
            elem_classes=["preview-btn"],
            size="sm",
        )
        gr.HTML(
            '<span class="tooltip" title="预览为方便查看增加了美化，'
            '实际输出文件不含 CSS 样式，请以实际保存结果为准">ℹ️</span>'
        )

    middle_preview_content = gr.HTML(visible=False)
    chunk_preview_content = gr.HTML(visible=False)

    return (
        status_output,
        middle_result_label,
        middle_output,
        final_result_label,
        chunk_output,
        middle_preview_btn,
        chunk_preview_btn,
        middle_preview_content,
        chunk_preview_content,
    )


def _create_usage_guide() -> None:
    """创建使用说明"""
    gr.HTML('<div class="gap"></div>')

    with gr.Accordion("使用说明", open=False, elem_classes=["accordion"]):
        gr.Markdown("""
**输出格式**：HTML / Markdown

**增强功能**：上下文注入、幽灵标题、表头扁平化、合并单元格处理、注释智能分发

**切分模式**：按 Token 数（精确控制大小）/ 按行数（固定行数）

**输出文件**：中间结果（增强后完整文件）、最终结果（切分后 chunks）

**预览**：点击预览按钮在新标签页查看，预览含美化样式，建议实际下载查看

**注意**：1. HTML 格式的标签会占用大量 Token，相同 Token 限制下实际有效内容会比 Markdown 格式少，且容易Token爆炸超出上下文长度;
2. 当前未实现表头自动清理,只有合并,建议先简单手动处理一次表头

""")


def _setup_preview_js(
    preview_btn: gr.Button,
    preview_fn,
    preview_content: gr.HTML,
) -> None:
    """设置预览按钮的 JavaScript"""
    preview_btn.click(
        fn=preview_fn,
        inputs=[],
        outputs=[preview_content],
    ).then(
        fn=None,
        inputs=[preview_content],
        outputs=[],
        js="""(content) => {
            if (content) {
                const blob = new Blob([content], {type: 'text/html;charset=utf-8'});
                const url = URL.createObjectURL(blob);
                window.open(url, '_blank');
            } else {
                alert('请先处理文件');
            }
        }""",
    )


def create_ui() -> gr.Blocks:
    """创建 Gradio 界面"""
    with gr.Blocks(css=CUSTOM_CSS) as app:
        gr.HTML('<h1 class="main-title">Excel 转换 RAG 增强工具</h1>')
        gr.HTML('<p class="sub-title">将 Excel 表格转换为 RAG 优化的 HTML/MD 格式，支持智能切分</p>')

        with gr.Row(equal_height=False):
            with gr.Column(scale=1):
                (
                    excel_input,
                    output_format,
                    keywords_input,
                    chunking_options,
                    split_mode,
                    target_tokens,
                    enable_min_tokens,
                    min_tokens,
                    token_strategy,
                    max_rows,
                    separator_input,
                    process_btn,
                ) = _create_input_column()

            with gr.Column(scale=1):
                (
                    status_output,
                    middle_result_label,
                    middle_output,
                    final_result_label,
                    chunk_output,
                    middle_preview_btn,
                    chunk_preview_btn,
                    middle_preview_content,
                    chunk_preview_content,
                ) = _create_output_column()

        # 绑定事件：输出格式切换
        output_format.change(
            fn=_toggle_format_options,
            inputs=[output_format],
            outputs=[chunking_options],
        )

        # 绑定事件：切分模式切换
        split_mode.change(
            fn=_toggle_slider,
            inputs=[split_mode],
            outputs=[target_tokens, enable_min_tokens, min_tokens, max_rows],
        )

        enable_min_tokens.change(
            fn=_toggle_min_tokens,
            inputs=[enable_min_tokens],
            outputs=[min_tokens, token_strategy],
        )

        process_btn.click(
            fn=process_excel,
            inputs=[
                excel_input,
                output_format,
                keywords_input,
                split_mode,
                max_rows,
                target_tokens,
                enable_min_tokens,
                min_tokens,
                token_strategy,
                separator_input,
            ],
            outputs=[middle_output, chunk_output, status_output],
            show_progress="minimal",
        )

        _setup_preview_js(middle_preview_btn, get_html_preview, middle_preview_content)
        _setup_preview_js(chunk_preview_btn, get_chunk_preview, chunk_preview_content)

        _create_usage_guide()

    return app
