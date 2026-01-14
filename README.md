# Excel 转换 RAG 增强工具

一个专为 RAG（检索增强生成）场景优化的 Excel 表格转换工具，支持将 Excel 文件转换为 HTML 或 Markdown 格式，并提供智能切分功能。

## 功能特性

### 核心转换能力

- **多格式输出**：支持 HTML 和 Markdown 两种输出格式
- **智能切分**：按 Token 数或行数将大表格切分为多个 chunks
- **RAG 优化**：专为大语言模型检索场景设计的输出格式

### RAG 增强功能

| 功能 | 说明 |
|------|------|
| 上下文硬编码 | 自动注入文件名、Sheet 名等元数据，帮助 LLM 理解数据来源 |
| 幽灵标题 | 添加用户定义的关键检索词，提升检索召回率 |
| 表头降维 | 将多层表头扁平化为单行，父级标题拼接到子级（如 "财务-收入-Q1"） |
| 合并单元格处理 | 智能展开合并单元格，确保每个 chunk 数据完整 |
| 注释智能分发 | 自动识别表格注释，按引用关系分发到相关 chunks |

## 项目结构

```
file-handle/
├── src/
│   ├── app/                    # 应用层（Gradio UI）
│   │   ├── main.py            # 应用入口
│   │   ├── ui.py              # UI 组件定义
│   │   ├── handlers.py        # 业务处理器
│   │   └── styles.py          # CSS 样式
│   └── core/                   # 核心转换逻辑
│       ├── config.py          # 配置管理
│       ├── models.py          # 数据模型
│       ├── base_converter.py  # 转换器基类
│       ├── unified_pipeline.py # 统一流水线
│       ├── excel2html/        # HTML 转换模块
│       │   ├── converter.py   # Excel → HTML 转换
│       │   ├── chunker.py     # HTML 切分器
│       │   └── pipeline.py    # HTML 流水线
│       └── excel2md/          # Markdown 转换模块
│           ├── converter.py   # Excel → MD 转换
│           └── chunker.py     # MD 切分器
├── tests/                      # 测试目录
├── Files/                      # 示例文件
├── pyproject.toml             # 项目配置
└── start.ps1                  # 启动脚本
```

## 安装

### 环境要求

- Python >= 3.12
- uv（推荐）或 pip

### 安装步骤

```bash
# 使用 uv 安装依赖
uv sync

# 或使用 pip
pip install -e .
```

## 使用方式

### 方式一：Web 界面（推荐）

```bash
# 启动 Gradio 应用
python -m src.app.main

# 或使用启动脚本
./start.ps1
```

启动后访问 `http://localhost:7860` 使用图形界面。

### 方式二：命令行

```bash
# 基本用法
python -m src.core.excel2html.pipeline input.xlsx

# 指定输出格式
python -m src.core.excel2html.pipeline input.xlsx -f md

# 添加关键词
python -m src.core.excel2html.pipeline input.xlsx -k "财务报表" "年度收入"

# 设置 Token 限制
python -m src.core.excel2html.pipeline input.xlsx -t 1024

# 按行数切分
python -m src.core.excel2html.pipeline input.xlsx -r 5

# 自定义分隔符
python -m src.core.excel2html.pipeline input.xlsx -s "---SPLIT---"
```

### 方式三：Python API

```python
from pathlib import Path
from src.core.unified_pipeline import UnifiedPipeline
from src.core.models import OutputFormat

# 创建流水线
pipeline = UnifiedPipeline(
    output_format=OutputFormat.MARKDOWN,
    keywords=["财务", "报表"],
    target_tokens=512,
    separator="!!!_CHUNK_BREAK_!!!"
)

# 执行转换
result = pipeline.run(Path("input.xlsx"))

print(f"中间结果: {result.output_path}")
print(f"最终结果: {result.chunk_path}")
print(f"Chunk 数量: {result.chunk_count}")
```

## 处理流程

```
┌─────────────┐     ┌─────────────┐     ┌─────────────┐
│  Excel 文件  │ ──▶ │  增强转换    │ ──▶ │  智能切分    │
└─────────────┘     └─────────────┘     └─────────────┘
                           │                    │
                           ▼                    ▼
                    ┌─────────────┐     ┌─────────────┐
                    │ 中间结果     │     │ 最终结果     │
                    │ (完整文件)   │     │ (Chunks)    │
                    └─────────────┘     └─────────────┘
```

### 第一步：增强转换

1. **解析 Excel**：使用 openpyxl 读取工作簿
2. **提取合并单元格**：记录所有合并区域的位置和值
3. **检测表头行数**：根据合并单元格自动识别多层表头
4. **表头降维**：将多层表头扁平化为单行
5. **检测注释**：识别表格末尾的注释行
6. **生成输出**：根据目标格式生成 HTML 或 Markdown

### 第二步：智能切分

1. **解析结构**：提取表头、数据行、注释元数据
2. **计算开销**：估算固定部分（表头、元数据）的 Token 数
3. **逐行累加**：根据切分策略决定切分点
4. **注释分发**：将相关注释分发到对应 chunk
5. **组装输出**：每个 chunk 包含完整表头和匹配的注释

## 配置选项

### 切分模式

| 模式 | 说明 | 适用场景 |
|------|------|----------|
| 按 Token 数 | 精确控制每个 chunk 的大小 | LLM 上下文长度有限时 |
| 按行数 | 固定每个 chunk 的数据行数 | 数据行大小均匀时 |

### Token 策略（仅 Token 模式）

| 策略 | 说明 |
|------|------|
| 接近最大值 | 累加到接近 max_tokens 才切分，chunk 数量少 |
| 接近最小值 | 超过 min_tokens 就切分，chunk 大小更均匀 |

### 环境变量

通过环境变量或 `.env` 文件配置：

```bash
EXCEL2HTML_TEMP_DIR=./temp           # 临时文件目录
EXCEL2HTML_DEFAULT_MAX_TOKENS=512    # 默认最大 Token 数
EXCEL2HTML_DEFAULT_MIN_TOKENS=256    # 默认最小 Token 数
EXCEL2HTML_DEFAULT_MAX_ROWS=8        # 默认最大行数
EXCEL2HTML_LOG_LEVEL=INFO            # 日志级别
EXCEL2HTML_DEFAULT_SEPARATOR="!!!_CHUNK_BREAK_!!!"  # 默认分隔符
```

## 输出格式说明

### HTML 格式

```html
<div class="rag-context">【文档上下文】来源：报表.xlsx | 数据类型：表格数据 【表格注释】[注1]xxx</div>
<script type="application/json" class="table-notes-meta">{"header_notes":{},"conditional_notes":{}}</script>
<table border="1" data-source="报表.xlsx" data-sheet="Sheet1">
    <caption>关键检索词：财务，报表</caption>
    <thead>
        <tr><th>列1</th><th>列2</th></tr>
    </thead>
    <tbody>
        <tr><td>数据1</td><td>数据2</td></tr>
    </tbody>
</table>
```

### Markdown 格式

```markdown
---
source: 报表.xlsx
sheet: Sheet1
keywords: 财务, 报表
notes: [注1]xxx
---

| 列1 | 列2 |
| --- | --- |
| 数据1 | 数据2 |
```

## 注意事项

1. **Token 计算**：HTML 标签会占用大量 Token，相同限制下 Markdown 格式的有效内容更多
2. **表头处理**：当前仅支持表头合并降维，建议先手动简化复杂表头并剔除冗余内容
3. **大文件处理**：超大 Excel 文件建议先拆分后处理
4. **编码问题**：输出文件统一使用 UTF-8 编码

## 依赖说明

| 依赖 | 用途 |
|------|------|
| openpyxl | Excel 文件解析 |
| beautifulsoup4 | HTML 解析和操作 |
| tiktoken | Token 数量计算（OpenAI 编码） |
| gradio | Web 界面 |
| loguru | 日志记录 |
| pydantic | 数据验证和配置管理 |

## 开发

```bash
# 安装开发依赖
uv sync --extra dev

# 运行测试
pytest

# 代码检查
ruff check src/

# 代码格式化
ruff format src/
```

## License

MIT
