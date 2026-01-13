"""
Gradio UI 组件定义
"""

import gradio as gr

from .handlers import process_excel, get_html_preview, get_chunk_preview
from .styles import CUSTOM_CSS


def create_ui():
    """创建 Gradio 界面"""
    with gr.Blocks(css=CUSTOM_CSS) as app:
        gr.HTML('<h1 class="main-title">Excel 转 HTML RAG 增强工具</h1>')
        gr.HTML('<p class="sub-title">将 Excel 表格转换为 RAG 优化的 HTML 片段，支持智能切分</p>')

        with gr.Row(equal_height=False):
            # 左侧：输入区
            with gr.Column(scale=1):
                excel_input = gr.File(
                    label="上传 Excel 文件",
                    file_types=[".xlsx", ".xls"],
                    elem_classes=["file-upload"],
                )
                
                gr.HTML('<div class="gap"></div>')
                
                keywords_input = gr.Textbox(
                    label="关键检索词（可选）",
                    placeholder="多个关键词用逗号分隔，如：财务报表, 年度收入, 利润分析",
                    lines=3,
                )

                gr.HTML('<div class="gap"></div>')

                split_mode = gr.Radio(
                    choices=["按 Token 数", "按行数"],
                    value="按 Token 数",
                    label="切分模式",
                    elem_classes=["radio-group"],
                )

                gr.HTML('<div class="gap"></div>')

                target_tokens = gr.Slider(
                    minimum=256,
                    maximum=4096,
                    value=1024,
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
                
                # 切分模式切换时显示/隐藏对应滑块
                def toggle_slider(mode):
                    if mode == "按 Token 数":
                        return (
                            gr.update(visible=True),  # target_tokens
                            gr.update(visible=True),  # enable_min_tokens
                            gr.update(visible=False), # min_tokens (由 checkbox 控制)
                            gr.update(visible=False), # max_rows
                        )
                    else:
                        return (
                            gr.update(visible=False), # target_tokens
                            gr.update(visible=False), # enable_min_tokens
                            gr.update(visible=False), # min_tokens
                            gr.update(visible=True),  # max_rows
                        )
                
                split_mode.change(
                    fn=toggle_slider,
                    inputs=[split_mode],
                    outputs=[target_tokens, enable_min_tokens, min_tokens, max_rows],
                )
                
                # 启用最小 token 时显示滑块和策略选择
                def toggle_min_tokens(enabled):
                    return gr.update(visible=enabled), gr.update(visible=enabled)
                
                enable_min_tokens.change(
                    fn=toggle_min_tokens,
                    inputs=[enable_min_tokens],
                    outputs=[min_tokens, token_strategy],
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

            # 右侧：输出区
            with gr.Column(scale=1):
                status_output = gr.Textbox(
                    label="处理状态",
                    lines=10,
                    interactive=False,
                    elem_classes=["status-box"],
                    placeholder="处理结果将显示在这里...",
                )
                
                gr.HTML('<div class="gap"></div>')
                
                # 中间结果 + 预览按钮
                gr.HTML('<p class="result-label">中间结果 (HTML)</p>')
                with gr.Row(elem_classes=["file-row"]):
                    html_output = gr.File(
                        label=None,
                        show_label=False,
                        elem_classes=["file-download"],
                    )
                    html_preview_btn = gr.Button(
                        "预览",
                        elem_classes=["preview-btn"],
                        size="sm",
                    )
                    gr.HTML('<span class="tooltip" title="预览为方便查看增加了美化，实际输出文件不含 CSS 样式，请以实际保存结果为准">ℹ️</span>')
                
                gr.HTML('<div class="gap"></div>')
                
                # 最终结果 + 预览按钮
                gr.HTML('<p class="result-label">最终结果 (Chunks)</p>')
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
                    gr.HTML('<span class="tooltip" title="预览为方便查看增加了美化，实际输出文件不含 CSS 样式，请以实际保存结果为准">ℹ️</span>')
                
                # 隐藏的 HTML 组件用于在新标签页打开预览
                html_preview_content = gr.HTML(visible=False)
                chunk_preview_content = gr.HTML(visible=False)

        # 绑定处理事件 - 禁用默认的 ETA 预估显示
        process_btn.click(
            fn=process_excel,
            inputs=[
                excel_input,
                keywords_input,
                split_mode,
                max_rows,
                target_tokens,
                enable_min_tokens,
                min_tokens,
                token_strategy,
                separator_input,
            ],
            outputs=[html_output, chunk_output, status_output],
            show_progress="minimal",
        )
        
        # 预览按钮 - 使用 JavaScript 在新标签页打开
        html_preview_btn.click(
            fn=get_html_preview,
            inputs=[],
            outputs=[html_preview_content],
        ).then(
            fn=None,
            inputs=[html_preview_content],
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
        
        chunk_preview_btn.click(
            fn=get_chunk_preview,
            inputs=[],
            outputs=[chunk_preview_content],
        ).then(
            fn=None,
            inputs=[chunk_preview_content],
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

        # 使用说明
        gr.HTML('<div class="gap"></div>')
        
        with gr.Accordion("使用说明", open=False, elem_classes=["accordion"]):
            gr.Markdown("""
**核心增强功能**

- **上下文硬编码**：自动注入文件名、Sheet 名等元数据，帮助 LLM 理解数据来源
- **幽灵标题**：通过关键检索词添加隐藏标题，提升 RAG 检索召回率
- **表头扁平化**：多层表头自动降维，将父级标题拼接到子级（如"销售额-Q1"）
- **合并单元格处理**：智能识别并正确处理跨行/跨列的合并单元格
- **注释智能分发**：自动提取表格末尾注释，按引用关系分发到对应 chunk

**切分模式**

- **按 Token 数**：根据目标 token 数动态切分，更精确控制每个 chunk 大小
- **按行数**：固定每个 chunk 的行数，适合行内容长度相近的表格

**Chunk 结构**

每个切分后的 chunk 都包含：
- 完整的文档上下文信息
- 扁平化后的表头
- 关键检索词（如有）
- 相关的表格注释（按引用自动匹配）

**输出文件**

- **中间结果**：增强后的完整 HTML（包含所有增强处理）
- **最终结果**：切分后的 chunks，用分隔符连接

**预览功能**

- 点击「预览」按钮可在新标签页查看渲染后的 HTML 效果
- 预览增加了美化样式，实际输出文件为纯净 HTML
""")

    return app
