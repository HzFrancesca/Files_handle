"""
Gradio UI 自定义 CSS 样式
"""

CUSTOM_CSS = """
/* 整体容器 - 放大宽度 */
.gradio-container {
    max-width: 1400px !important;
    margin: auto !important;
    padding: 40px 60px !important;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif !important;
}

/* 标题样式 */
.main-title {
    text-align: center;
    color: #1a1a1a;
    font-size: 2.2rem !important;
    font-weight: 600;
    margin-bottom: 8px;
    letter-spacing: -0.5px;
}

.sub-title {
    text-align: center;
    color: #666;
    font-size: 1.1rem !important;
    margin-bottom: 40px;
}

/* 统一标签样式 - 放大 */
label, .label-wrap span {
    font-size: 1rem !important;
    font-weight: 500 !important;
    color: #333 !important;
    margin-bottom: 8px !important;
}

/* 输入框样式 - 放大 */
textarea, input[type="text"] {
    font-size: 1rem !important;
    padding: 12px 16px !important;
    border-radius: 8px !important;
    border: 1px solid #d1d5db !important;
}

textarea:focus, input[type="text"]:focus {
    border-color: #2563eb !important;
    box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1) !important;
}

/* 主按钮样式 */
.primary-btn {
    background: #2563eb !important;
    border: none !important;
    font-size: 1.1rem !important;
    font-weight: 500 !important;
    padding: 16px 32px !important;
    border-radius: 10px !important;
    margin-top: 16px !important;
}

.primary-btn:hover {
    background: #1d4ed8 !important;
}

/* 预览按钮样式 */
.preview-btn {
    background: #f3f4f6 !important;
    border: 1px solid #d1d5db !important;
    color: #374151 !important;
    font-size: 0.9rem !important;
    padding: 8px 16px !important;
    border-radius: 6px !important;
    min-width: 80px !important;
}

.preview-btn:hover {
    background: #e5e7eb !important;
    border-color: #9ca3af !important;
}

/* 状态输出框 - 放大 */
.status-box textarea {
    font-family: "SF Mono", Monaco, "Cascadia Code", Consolas, monospace !important;
    font-size: 0.95rem !important;
    line-height: 1.8 !important;
    background: #f8f9fa !important;
    padding: 16px !important;
    min-height: 200px !important;
}

/* 文件上传区域 - 放大 */
.file-upload {
    min-height: 160px !important;
}

.file-upload .wrap {
    padding: 40px !important;
}

/* 滑块样式 */
input[type="range"] {
    accent-color: #2563eb !important;
    height: 6px !important;
}

/* 滑块数值显示 */
.slider-number input {
    font-size: 1rem !important;
    padding: 8px 12px !important;
}

/* Radio 按钮 - 放大间距 */
.radio-group {
    gap: 16px !important;
}

.radio-group label {
    font-size: 1rem !important;
    padding: 10px 20px !important;
}

/* 组件间距 */
.gap {
    margin-top: 24px !important;
}

/* 文件下载区域 */
.file-download {
    min-height: 80px !important;
}

.file-download .wrap {
    min-height: 60px !important;
    border: 1px dashed #d1d5db !important;
    border-radius: 8px !important;
    background: #fafafa !important;
}

/* 文件行容器 */
.file-row {
    display: flex;
    align-items: center;
    gap: 12px;
    margin-top: 16px;
}

.file-row > div:first-child {
    flex: 1;
}

/* Accordion 样式 */
.accordion {
    margin-top: 32px !important;
}

.accordion summary {
    font-size: 1rem !important;
    padding: 16px !important;
}

.accordion .prose {
    font-size: 0.95rem !important;
    line-height: 1.8 !important;
    padding: 20px !important;
}

/* 两列布局间距 */
.contain > .wrap {
    gap: 40px !important;
}

/* 强制双栏布局始终显示 */
.gradio-container .row {
    flex-wrap: nowrap !important;
}

.gradio-container .row > .column {
    min-width: 0 !important;
    flex: 1 1 50% !important;
}

/* 结果区块标题 */
.result-label {
    font-size: 0.95rem !important;
    font-weight: 500 !important;
    color: #333 !important;
    margin-bottom: 8px !important;
}

/* 提示图标样式 */
.tooltip {
    cursor: help;
    color: #6b7280;
    font-size: 0.9rem;
    margin-left: 4px;
}

.tooltip:hover {
    color: #2563eb;
}

/* 隐藏 Gradio 的预估时间显示，只保留已用时间 */
.eta-bar {
    display: none !important;
}

.progress-text::after {
    content: "" !important;
}

/* 隐藏进度条中的预估时间部分 (格式: 已用时间/预估时间) */
.meta-text-center::after,
.progress-bar span[data-testid="eta"] {
    display: none !important;
}
"""
