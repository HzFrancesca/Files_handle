# Excel 转 HTML RAG 增强流水线 - 完整技术文档

## 目录

1. [概述](#1-概述)
2. [流水线架构](#2-流水线架构)
3. [第一阶段：Excel 转 HTML（RAG 增强版）](#3-第一阶段excel-转-htmlrag-增强版)
4. [第二阶段：HTML 智能切分（Chunking）](#4-第二阶段html-智能切分chunking)
5. [流水线主控程序](#5-流水线主控程序)
6. [使用指南](#6-使用指南)
7. [RAG 增强技术详解](#7-rag-增强技术详解)
8. [Token 估算与动态切分](#8-token-估算与动态切分)

---

## 1. 概述

本流水线专为 RAG（Retrieval-Augmented Generation）场景设计，将 Excel 表格转换为经过优化的 HTML 片段，以提升向量检索的召回率和准确性。

### 1.1 核心问题

传统的 Excel 转 HTML 方案存在以下问题：

- **上下文丢失**：转换后的 HTML 缺少文件名、Sheet 名等元数据，检索时无法定位来源
- **多层表头问题**：复杂表头（如合并单元格形成的父子标题）在切分后丢失层级关系
- **检索词缺失**：表格内容可能不包含用户常用的检索词汇
- **切分粒度不当**：简单按行数切分可能导致 chunk 过大超出 token 限制，或过小丢失上下文

### 1.2 解决方案

本流水线通过以下增强技术解决上述问题：

| 增强技术 | 解决的问题 | 实现位置 |
|---------|-----------|---------|
| 上下文硬编码 | 上下文丢失 | `excel2html_openpyxl_enhanced.py` |
| 幽灵标题（Ghost Caption） | 检索词缺失 | `excel2html_openpyxl_enhanced.py` |
| 表头降维 | 多层表头问题 | `excel2html_openpyxl_enhanced.py` |
| 全局资产分发 | 切分后上下文丢失 | `html2chunk.py` |
| Token 动态切分 | 切分粒度不当 | `html2chunk.py` |

---

## 2. 流水线架构

### 2.1 整体流程图

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                           Excel 转 HTML RAG 增强流水线                        │
└─────────────────────────────────────────────────────────────────────────────┘

                                      │
                                      ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│  输入: Excel 文件 (.xlsx)                                                    │
│  参数: keywords（可选）, target_tokens（默认512）, separator                   │
└─────────────────────────────────────────────────────────────────────────────┘
                                      │
                    ┌─────────────────┴─────────────────┐
                    ▼                                   │
┌───────────────────────────────────┐                   │
│  第一阶段: Excel → HTML            │                   │
│  excel2html_openpyxl_enhanced.py  │                   │
│                                   │                   │
│  增强功能:                         │                   │
│  ✓ 上下文硬编码                    │                   │
│  ✓ 幽灵标题 (Ghost Caption)        │                   │
│  ✓ 表头降维                        │                   │
│  ✓ 合并单元格智能处理               │                   │
└───────────────────────────────────┘                   │
                    │                                   │
                    ▼                                   │
┌───────────────────────────────────┐                   │
│  中间产物: 增强 HTML 文件           │                   │
│  例: input.xlsx → input.html      │                   │
└───────────────────────────────────┘                   │
                    │                                   │
                    ▼                                   │
┌───────────────────────────────────┐                   │
│  第二阶段: HTML → Chunks           │                   │
│  html2chunk.py                    │                   │
│                                   │                   │
│  切分功能:                         │                   │
│  ✓ 全局资产提取与分发               │                   │
│  ✓ 行数模式切分                    │                   │
│  ✓ Token 模式动态切分              │                   │
└───────────────────────────────────┘                   │
                    │                                   │
                    ▼                                   │
┌───────────────────────────────────────────────────────┘
│  最终产物: Chunk 合并文件
│  例: input.xlsx → input_chunk_merged.html
│  格式: chunk1 + separator + chunk2 + separator + ...
└─────────────────────────────────────────────────────────
```

### 2.2 文件结构

```
src/excel2html/
├── excel2html_openpyxl_enhanced.py  # 第一阶段：Excel 转 HTML（RAG 增强版）
├── html2chunk.py                     # 第二阶段：HTML 智能切分
└── pipeline.py                       # 流水线主控程序
```

### 2.3 输出文件

假设输入文件为 `Files/Excel/本国子目注释调整表.xlsx`，流水线将生成：

| 文件 | 说明 |
|------|------|
| `Files/Excel/本国子目注释调整表.html` | 中间结果：增强后的完整 HTML |
| `Files/Excel/本国子目注释调整表_chunk_merged.html` | 最终结果：切分后的 chunks |

---

## 3. 第一阶段：Excel 转 HTML（RAG 增强版）

### 3.1 文件信息

- **文件路径**: `src/excel2html/excel2html_openpyxl_enhanced.py`
- **依赖库**: `openpyxl`, `pathlib`, `datetime`
- **主要功能**: 将 Excel 文件转换为 RAG 优化的 HTML 表格

### 3.2 数值格式化处理

#### 3.2.1 问题背景

Excel 里的数据存储值（Value）和显示值（Number Format）往往不一致：

| 数据类型 | 存储值 | 显示值 | 风险 |
|---------|--------|--------|------|
| 百分比 | `0.5` | `50%` | LLM 看到 0.5 无法理解"占比多少" |
| 货币 | `1000000` | `¥1,000,000` | 丢失货币符号和千分位 |
| 日期 | `44927`（Excel 序列号）或 `datetime` 对象 | `2023-01-01` | LLM 无法回答"2023年的数据" |
| 科学计数 | `0.000123` | `1.23E-04` | 格式丢失 |

如果提取工具只读取 `cell.value` 而不处理格式，LLM 可能会看到一堆原始数值，导致无法正确理解数据含义。

#### 3.2.2 解决方案

使用 `format_cell_value()` 函数，根据单元格的 `number_format` 属性返回格式化后的显示值：

```python
def format_cell_value(cell):
    """
    根据单元格的 number_format 返回格式化后的显示值
    解决 openpyxl 只返回存储值而非显示值的问题
    """
    value = cell.value
    if value is None:
        return ""

    number_format = cell.number_format or "General"

    # 日期/时间类型
    if isinstance(value, datetime):
        if "H" in number_format or "h" in number_format:
            return value.strftime("%Y-%m-%d %H:%M:%S")
        else:
            return value.strftime("%Y-%m-%d")

    # 非数字类型直接返回
    if not isinstance(value, (int, float)):
        return str(value)

    # 百分比格式
    if "%" in number_format:
        decimal_match = re.search(r"0\.(0+)%", number_format)
        decimals = len(decimal_match.group(1)) if decimal_match else 0
        return f"{value * 100:.{decimals}f}%"

    # 科学计数法
    if "E" in number_format.upper() and number_format != "General":
        decimal_match = re.search(r"0\.(0+)E", number_format, re.IGNORECASE)
        decimals = len(decimal_match.group(1)) if decimal_match else 2
        return f"{value:.{decimals}E}"

    # 货币和千分位格式
    if "#,##" in number_format or ",0" in number_format:
        decimal_match = re.search(r"0\.(0+)", number_format)
        decimals = len(decimal_match.group(1)) if decimal_match else 0
        formatted = f"{value:,.{decimals}f}"

        if "¥" in number_format or "￥" in number_format:
            return f"¥{formatted}"
        elif "$" in number_format:
            return f"${formatted}"
        else:
            return formatted

    # 默认：普通数字
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return str(value)
```

#### 3.2.3 格式化效果对比

| 数据类型 | 修复前输出 | 修复后输出 |
|---------|-----------|-----------|
| 百分比 | `0.5` | `50%` |
| 百分比(小数) | `0.1234` | `12.34%` |
| 货币(CNY) | `1000000` | `¥1,000,000` |
| 货币(USD) | `1234.56` | `$1,234.56` |
| 千分位 | `9876543` | `9,876,543` |
| 日期 | `2023-01-01 00:00:00` | `2023-01-01` |
| 科学计数 | `0.000123` | `1.23E-04` |

#### 3.2.4 重要配置

为了获取单元格的 `number_format` 属性，必须使用 `data_only=False` 加载 Excel：

```python
workbook = openpyxl.load_workbook(str(source_path), data_only=False)
```

**注意**: `data_only=True` 时可以读取公式的计算结果，但无法获取格式信息。当前实现选择保留格式信息，对于包含公式的单元格，会显示公式本身而非计算结果。如果需要同时获取公式结果和格式，可以考虑先用 `data_only=True` 读取值，再用 `data_only=False` 读取格式。

---

### 3.3 核心函数详解

#### 3.3.1 `get_merged_cell_info(sheet)`

**功能**: 获取所有合并单元格的信息，为后续处理提供基础数据。

**返回值**: 字典，键为 `(row, col)` 元组，值为单元格信息。

```python
def get_merged_cell_info(sheet):
    """
    获取所有合并单元格的信息
    返回: {(row, col): {'value': 值, 'rowspan': 行跨度, 'colspan': 列跨度, 'is_origin': 是否是左上角}}
    """
    merged_info = {}

    for merged_range in sheet.merged_cells.ranges:
        min_row, min_col = merged_range.min_row, merged_range.min_col
        max_row, max_col = merged_range.max_row, merged_range.max_col
        origin_value = sheet.cell(row=min_row, column=min_col).value

        rowspan = max_row - min_row + 1
        colspan = max_col - min_col + 1

        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                is_origin = r == min_row and c == min_col
                merged_info[(r, c)] = {
                    "value": origin_value,
                    "rowspan": rowspan if is_origin else 0,
                    "colspan": colspan if is_origin else 0,
                    "is_origin": is_origin,
                    "skip": not is_origin,  # 非左上角的单元格需要跳过
                }

    return merged_info
```

**数据结构示例**:

假设有一个 2x3 的合并单元格（从 A1 到 B3），返回的数据结构为：

```python
{
    (1, 1): {"value": "合并内容", "rowspan": 3, "colspan": 2, "is_origin": True, "skip": False},
    (1, 2): {"value": "合并内容", "rowspan": 0, "colspan": 0, "is_origin": False, "skip": True},
    (2, 1): {"value": "合并内容", "rowspan": 0, "colspan": 0, "is_origin": False, "skip": True},
    (2, 2): {"value": "合并内容", "rowspan": 0, "colspan": 0, "is_origin": False, "skip": True},
    (3, 1): {"value": "合并内容", "rowspan": 0, "colspan": 0, "is_origin": False, "skip": True},
    (3, 2): {"value": "合并内容", "rowspan": 0, "colspan": 0, "is_origin": False, "skip": True},
}
```

#### 3.2.2 `detect_header_rows(sheet, merged_info, max_check_rows=5)`

**功能**: 自动检测表头行数，通过识别合并单元格来判断多层表头。

**算法逻辑**:
1. 默认表头为 1 行
2. 遍历前 `max_check_rows` 行（默认 5 行）
3. 如果某行存在 `colspan > 1` 的合并单元格，说明该行是父级表头
4. 表头行数 = 最后一个有 colspan 的行号 + 1

```python
def detect_header_rows(sheet, merged_info, max_check_rows=5):
    """
    检测表头行数（通过合并单元格和内容特征判断）
    返回表头行数
    """
    header_rows = 1

    for row_idx in range(1, min(max_check_rows + 1, sheet.max_row + 1)):
        has_colspan = False
        for col_idx in range(1, sheet.max_column + 1):
            info = merged_info.get((row_idx, col_idx))
            if info and info.get("colspan", 1) > 1:
                has_colspan = True
                break

        if has_colspan:
            header_rows = max(header_rows, row_idx + 1)

    return min(header_rows, sheet.max_row)
```

**示例**:

原始表格：
```
┌─────────────────────────────┬─────────────────────────────┐
│          财务数据            │          人员数据            │  ← 第1行，有 colspan
├──────────┬──────────────────┼──────────┬──────────────────┤
│   收入    │       支出       │   在职    │       离职       │  ← 第2行，无 colspan
├──────────┼──────────────────┼──────────┼──────────────────┤
│   100    │        50        │    10    │        2         │  ← 数据行
└──────────┴──────────────────┴──────────┴──────────────────┘
```

检测结果：`header_rows = 2`

#### 3.2.3 `build_flattened_headers(sheet, merged_info, header_rows)`

**功能**: 将多层表头"降维"为单行表头，把父级标题拼接到子级标题。

**这是 RAG 增强的核心技术之一**，解决了切分后表头层级关系丢失的问题。

```python
def build_flattened_headers(sheet, merged_info, header_rows):
    """
    构建降维后的表头（把父级标题拼接到子级标题）
    返回: {col_idx: "父标题-子标题-..."}
    """
    if header_rows <= 1:
        # 单行表头，直接返回
        headers = {}
        for col_idx in range(1, sheet.max_column + 1):
            value = sheet.cell(row=1, column=col_idx).value
            headers[col_idx] = str(value) if value else f"列{col_idx}"
        return headers

    # 多行表头，需要降维
    # 先构建每列在每行的实际值（考虑合并单元格）
    col_values = {col: [] for col in range(1, sheet.max_column + 1)}

    for row_idx in range(1, header_rows + 1):
        for col_idx in range(1, sheet.max_column + 1):
            info = merged_info.get((row_idx, col_idx))
            if info:
                value = info["value"]
            else:
                value = sheet.cell(row=row_idx, column=col_idx).value

            col_values[col_idx].append(str(value) if value else "")

    # 拼接表头，去除重复和空值
    headers = {}
    for col_idx, values in col_values.items():
        # 去除空值和重复
        unique_values = []
        for v in values:
            v = v.strip()
            if v and (not unique_values or v != unique_values[-1]):
                unique_values.append(v)

        headers[col_idx] = "-".join(unique_values) if unique_values else f"列{col_idx}"

    return headers
```

**降维示例**:

原始多层表头：
```
┌─────────────────────────────┬─────────────────────────────┐
│          财务数据            │          人员数据            │
├──────────┬──────────────────┼──────────┬──────────────────┤
│   收入    │       支出       │   在职    │       离职       │
└──────────┴──────────────────┴──────────┴──────────────────┘
```

降维后的表头：
```
{
    1: "财务数据-收入",
    2: "财务数据-支出",
    3: "人员数据-在职",
    4: "人员数据-离职"
}
```

**去重逻辑说明**:

如果父级标题和子级标题相同（如某列只有一个标题跨越多行），会自动去重：
- 输入: `["总计", "总计", ""]` → 输出: `"总计"`
- 输入: `["财务", "收入", ""]` → 输出: `"财务-收入"`

#### 3.2.4 `sheet_to_enhanced_html(sheet, filename, keywords=None)`

**功能**: 将单个 Sheet 转换为 RAG 增强的 HTML 表格，这是第一阶段的核心函数。

```python
def sheet_to_enhanced_html(sheet, filename, keywords=None):
    """
    将单个 sheet 转换为 RAG 增强的 HTML 表格

    参数:
        sheet: openpyxl worksheet
        filename: 源文件名
        keywords: 可选的关键检索词列表
    """
    merged_info = get_merged_cell_info(sheet)
    header_rows = detect_header_rows(sheet, merged_info)
    flattened_headers = build_flattened_headers(sheet, merged_info, header_rows)

    html_parts = []

    # === 增强1: 上下文硬编码 ===
    update_time = datetime.now().strftime("%Y-%m-%d")
    context_html = f"""<div class="rag-context">【文档上下文】来源文件：{filename} | 工作表：{sheet.title} | 数据类型：表格数据 | 更新时间：{update_time}</div>"""
    html_parts.append(context_html)

    # 开始表格
    html_parts.append(
        f'<table border="1" style="border-collapse:collapse" data-source="{filename}" data-sheet="{sheet.title}">'
    )

    # === 增强2: 幽灵标题 (Ghost Caption) ===
    if keywords:
        keyword_str = "，".join(keywords)
        caption_html = f"    <caption>关键检索词：{keyword_str}。此表可能包含相关问题的答案。</caption>"
        html_parts.append(caption_html)

    # === 增强3: 表头降维 - 只保留扁平化的表头 ===
    html_parts.append("    <thead>")
    html_parts.append("        <tr>")
    for col_idx in range(1, sheet.max_column + 1):
        flat_header = flattened_headers.get(col_idx, "")
        html_parts.append(f"            <th>{flat_header}</th>")
    html_parts.append("        </tr>")
    html_parts.append("    </thead>")

    # 数据行
    html_parts.append("    <tbody>")
    for row_idx in range(header_rows + 1, sheet.max_row + 1):
        html_parts.append("        <tr>")
        for col_idx in range(1, sheet.max_column + 1):
            info = merged_info.get((row_idx, col_idx))

            if info and info.get("skip"):
                continue

            if info:
                value = info["value"]
                rowspan = info.get("rowspan", 1)
                colspan = info.get("colspan", 1)
                span_attrs = []
                if rowspan > 1:
                    span_attrs.append(f'rowspan="{rowspan}"')
                if colspan > 1:
                    span_attrs.append(f'colspan="{colspan}"')
                span_str = " " + " ".join(span_attrs) if span_attrs else ""
            else:
                value = sheet.cell(row=row_idx, column=col_idx).value
                span_str = ""

            cell_content = str(value) if value is not None else ""
            html_parts.append(f"            <td{span_str}>{cell_content}</td>")
        html_parts.append("        </tr>")

    html_parts.append("    </tbody>")
    html_parts.append("</table>")

    return "\n".join(html_parts)
```

**输出 HTML 结构**:

```html
<div class="rag-context">【文档上下文】来源文件：财务报表.xlsx | 工作表：Sheet1 | 数据类型：表格数据 | 更新时间：2025-01-12</div>
<table border="1" style="border-collapse:collapse" data-source="财务报表.xlsx" data-sheet="Sheet1">
    <caption>关键检索词：财务报表，年度收入，利润。此表可能包含相关问题的答案。</caption>
    <thead>
        <tr>
            <th>财务数据-收入</th>
            <th>财务数据-支出</th>
            <th>人员数据-在职</th>
            <th>人员数据-离职</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td>100</td>
            <td>50</td>
            <td>10</td>
            <td>2</td>
        </tr>
        <!-- 更多数据行... -->
    </tbody>
</table>
```

#### 3.2.5 `convert_excel_to_html(excel_path, keywords=None, output_path=None)`

**功能**: 主入口函数，将整个 Excel 文件转换为 HTML。

```python
def convert_excel_to_html(
    excel_path: str, keywords: list = None, output_path: str = None
):
    """
    将单个 Excel 文件转换为 RAG 增强的 HTML

    参数:
        excel_path: Excel 文件路径
        keywords: 可选的关键检索词列表，用于幽灵标题
        output_path: 可选的输出路径，默认与源文件同目录同名

    返回:
        成功返回输出文件路径，失败返回 None
    """
    source_path = Path(excel_path)

    if not source_path.exists():
        print(f"❌ 错误：找不到文件 '{source_path}'")
        return None

    if output_path:
        out_path = Path(output_path)
    else:
        out_path = source_path.with_suffix(".html")

    filename = source_path.name
    print(f"📄 正在处理: {filename}")
    print(f"   增强功能: 上下文硬编码 ✓ | 表头降维 ✓ | 合并单元格 ✓", end="")
    if keywords:
        print(f" | 幽灵标题 ✓ ({len(keywords)}个关键词)")
    else:
        print(" | 幽灵标题 ✗ (未提供关键词)")

    try:
        workbook = openpyxl.load_workbook(str(source_path), data_only=True)
    except Exception as e:
        print(f"❌ 解析失败: {e}")
        return None

    # 构建 HTML - 只输出核心内容，不包含文档外壳
    html_parts = []

    for sheet in workbook.worksheets:
        if sheet.max_row == 0 or sheet.max_column == 0:
            continue  # 跳过空 sheet

        html_parts.append(sheet_to_enhanced_html(sheet, filename, keywords))

    # 写入文件
    try:
        out_path.write_text("\n".join(html_parts), encoding="utf-8")
        print(f"✅ 转换成功！输出: {out_path.absolute()}")
        return str(out_path)
    except IOError as e:
        print(f"❌ 写入文件失败: {e}")
        return None
```

**关键设计决策**:

1. **`data_only=True`**: 读取公式单元格的计算结果值，而非公式本身
2. **跳过空 Sheet**: 避免生成无意义的空表格
3. **不包含 HTML 文档外壳**: 只输出核心内容（`<div>` + `<table>`），便于后续切分和嵌入

---

## 4. 第二阶段：HTML 智能切分（Chunking）

### 4.1 文件信息

- **文件路径**: `src/excel2html/html2chunk.py`
- **依赖库**: `beautifulsoup4`, `pathlib`, `copy`
- **主要功能**: 将长 HTML 表格切分为多个 chunks，并为每个 chunk 分发全局资产

### 4.2 核心概念：全局资产分发

**问题**: 当一个长表格被切分为多个 chunks 后，每个 chunk 都需要保留：
- 上下文信息（`<div class="rag-context">`）
- 幽灵标题（`<caption>`）
- 表头（`<thead>`）

否则，单独的 chunk 将失去上下文，无法被正确检索。

**解决方案**: 提取这些"全局资产"，在切分时复制到每个 chunk 中。

```
原始 HTML:
┌─────────────────────────────────────────┐
│  <div class="rag-context">...</div>     │  ← 全局资产
│  <table>                                │
│    <caption>...</caption>               │  ← 全局资产
│    <thead>...</thead>                   │  ← 全局资产
│    <tbody>                              │
│      <tr>Row 1</tr>                     │
│      <tr>Row 2</tr>                     │
│      <tr>Row 3</tr>                     │  ← 数据行（需要切分）
│      <tr>Row 4</tr>                     │
│      <tr>Row 5</tr>                     │
│      <tr>Row 6</tr>                     │
│    </tbody>                             │
│  </table>                               │
└─────────────────────────────────────────┘

切分后（每 chunk 2 行）:

Chunk 1:                          Chunk 2:                          Chunk 3:
┌─────────────────────┐           ┌─────────────────────┐           ┌─────────────────────┐
│  <div>context</div> │           │  <div>context</div> │           │  <div>context</div> │
│  <table>            │           │  <table>            │           │  <table>            │
│    <caption>...</>  │           │    <caption>...</>  │           │    <caption>...</>  │
│    <thead>...</>    │           │    <thead>...</>    │           │    <thead>...</>    │
│    <tbody>          │           │    <tbody>          │           │    <tbody>          │
│      <tr>Row 1</tr> │           │      <tr>Row 3</tr> │           │      <tr>Row 5</tr> │
│      <tr>Row 2</tr> │           │      <tr>Row 4</tr> │           │      <tr>Row 6</tr> │
│    </tbody>         │           │    </tbody>         │           │    </tbody>         │
│  </table>           │           │  </table>           │           │  </table>           │
└─────────────────────┘           └─────────────────────┘           └─────────────────────┘
```

### 4.3 核心函数详解

#### 4.3.1 `estimate_tokens(text)`

**功能**: 估算文本的 token 数量，用于动态切分。

```python
def estimate_tokens(text: str) -> int:
    """估算文本的 token 数量（中文约2.5字符=1token）"""
    return int(len(text) / 2.5)
```

**估算依据**:
- 中文字符：约 1-2 字符 = 1 token
- 英文单词：约 4 字符 = 1 token
- HTML 标签：按字符数估算
- 综合取值：`len(text) / 2.5`

**注意**: 这是简化估算，如需精确计算，可接入 `tiktoken` 库使用 OpenAI 的实际分词器。

#### 4.3.2 `distribute_assets_and_chunk(long_html_content, max_rows_per_chunk=None, max_tokens_per_chunk=None)`

**功能**: 核心切分函数，支持两种切分模式。

**参数说明**:
| 参数 | 类型 | 说明 |
|------|------|------|
| `long_html_content` | str | 完整的 HTML 内容 |
| `max_rows_per_chunk` | int | 行数模式：每个 chunk 的最大行数 |
| `max_tokens_per_chunk` | int | Token 模式：每个 chunk 的最大 token 数 |

**优先级**: 如果两个参数都未指定，默认使用 `max_rows_per_chunk=8`

```python
def distribute_assets_and_chunk(
    long_html_content,
    max_rows_per_chunk: int = None,
    max_tokens_per_chunk: int = None
):
    """
    核心逻辑：将长 HTML 切分，并把全局资产（Context/Caption/Header）分发给每个片段
    
    参数:
        long_html_content: 完整 HTML 内容
        max_rows_per_chunk: 按行数切分（优先）
        max_tokens_per_chunk: 按 token 数切分（更精确）
        
    如果两个参数都未指定，默认 max_rows_per_chunk=8
    """
    # 默认值
    if max_rows_per_chunk is None and max_tokens_per_chunk is None:
        max_rows_per_chunk = 8
    
    soup = BeautifulSoup(long_html_content, "html.parser")

    # 1. 提取全局资产 (Context Div)
    context_div = soup.find("div", class_="rag-context")
    if not context_div:
        table_node = soup.find("table")
        if table_node and table_node.find_previous_sibling("div"):
            context_div = table_node.find_previous_sibling("div")

    # 2. 提取表格核心组件
    original_table = soup.find("table")
    if not original_table:
        return [long_html_content]

    caption = original_table.find("caption")

    header_rows = []
    thead = original_table.find("thead")
    if thead:
        header_rows = thead.find_all("tr")
    else:
        header_rows = original_table.find_all("tr")[:1]

    # 3. 准备数据行
    tbody = original_table.find("tbody")
    if tbody:
        data_rows = tbody.find_all("tr")
    else:
        all_rows = original_table.find_all("tr")
        data_rows = [row for row in all_rows if row not in header_rows]

    # 4. 计算固定开销（用于 token 模式）
    fixed_parts = []
    if context_div:
        fixed_parts.append(str(context_div))
    if caption:
        fixed_parts.append(str(caption))
    for h_row in header_rows:
        fixed_parts.append(str(h_row))
    fixed_overhead = estimate_tokens("".join(fixed_parts))

    chunks = []
    current_chunk_data = []
    current_chunk_tokens = 0
```

**切分判断逻辑**:

```python
    def should_split(row_count, row_tokens):
        """判断是否应该切分"""
        if max_tokens_per_chunk is not None:
            # Token 模式：检查累计 token 是否超限
            return (current_chunk_tokens + row_tokens + fixed_overhead) > max_tokens_per_chunk
        else:
            # 行数模式
            return row_count >= max_rows_per_chunk
```

**Chunk 组装逻辑**:

```python
    def build_chunk(data_rows_for_chunk):
        """组装一个 chunk"""
        new_soup = BeautifulSoup("<div></div>", "html.parser")
        wrapper_div = new_soup.div

        if context_div:
            wrapper_div.append(copy.copy(context_div))

        new_table = new_soup.new_tag("table")
        new_table.attrs = original_table.attrs
        new_table["border"] = "1"
        new_table["style"] = "border-collapse:collapse"

        if caption:
            new_table.append(copy.copy(caption))

        new_thead = new_soup.new_tag("thead")
        for h_row in header_rows:
            new_thead.append(copy.copy(h_row))
        new_table.append(new_thead)

        new_tbody = new_soup.new_tag("tbody")
        for d_row in data_rows_for_chunk:
            new_tbody.append(copy.copy(d_row))
        new_table.append(new_tbody)

        wrapper_div.append(new_table)
        return str(wrapper_div)
```

**逐行累加切分**:

```python
    # 5. 逐行累加切分
    for i, row in enumerate(data_rows):
        row_tokens = estimate_tokens(str(row))
        
        # 检查是否需要先切分（当前 chunk 非空且加入新行会超限）
        if current_chunk_data and should_split(len(current_chunk_data), row_tokens):
            chunks.append(build_chunk(current_chunk_data))
            current_chunk_data = []
            current_chunk_tokens = 0

        current_chunk_data.append(row)
        current_chunk_tokens += row_tokens

        # 最后一行，收尾
        if i == len(data_rows) - 1 and current_chunk_data:
            chunks.append(build_chunk(current_chunk_data))

    return chunks
```

### 4.4 两种切分模式对比

| 特性 | 行数模式 | Token 模式 |
|------|---------|-----------|
| 参数 | `max_rows_per_chunk=N` | `max_tokens_per_chunk=N` |
| 切分依据 | 固定行数 | 累计 token 数 |
| 精确度 | 较低（行内容长度不一） | 较高（考虑实际内容） |
| 适用场景 | 行内容长度相近的表格 | 行内容长度差异大的表格 |
| 计算开销 | 低 | 略高（需要估算 token） |

**Token 模式的优势**:

假设目标 token 为 512，固定开销为 150 tokens：

```
行数模式（每 chunk 5 行）:
  Chunk 1: 150 + 50×5 = 400 tokens ✓
  Chunk 2: 150 + 200×5 = 1150 tokens ✗ 超限！

Token 模式（目标 512 tokens）:
  Chunk 1: 150 + 50 + 50 + 50 + 50 + 50 = 400 tokens ✓
  Chunk 2: 150 + 200 = 350 tokens ✓
  Chunk 3: 150 + 200 = 350 tokens ✓
  ...
```

---

## 5. 流水线主控程序

### 5.1 文件信息

- **文件路径**: `src/excel2html/pipeline.py`
- **依赖**: `excel2html_openpyxl_enhanced.py`, `html2chunk.py`
- **主要功能**: 串联两个阶段，提供统一的命令行接口

### 5.2 核心函数

#### 5.2.1 `estimate_rows_for_token_limit(html_content, target_tokens=512)`

**功能**: 根据目标 token 数预估每个 chunk 应该包含多少行（用于参考，实际切分使用逐行累加）。

```python
def estimate_rows_for_token_limit(html_content: str, target_tokens: int = 512) -> int:
    """
    根据目标 token 数估算每个 chunk 应该包含多少行

    参数:
        html_content: 完整的 HTML 内容
        target_tokens: 目标 token 数（默认 512）

    返回:
        建议的 max_rows_per_chunk
    """
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(html_content, "html.parser")

    # 计算固定开销（context + caption + thead）
    fixed_parts = []

    context_div = soup.find("div", class_="rag-context")
    if context_div:
        fixed_parts.append(str(context_div))

    table = soup.find("table")
    if not table:
        return 8  # 无表格，返回默认值

    caption = table.find("caption")
    if caption:
        fixed_parts.append(str(caption))

    thead = table.find("thead")
    if thead:
        fixed_parts.append(str(thead))

    fixed_overhead = estimate_tokens("".join(fixed_parts))

    # 计算每行平均 token
    tbody = table.find("tbody")
    if tbody:
        data_rows = tbody.find_all("tr")
    else:
        all_rows = table.find_all("tr")
        data_rows = all_rows[1:] if len(all_rows) > 1 else all_rows

    if not data_rows:
        return 8

    total_row_tokens = sum(estimate_tokens(str(row)) for row in data_rows)
    avg_tokens_per_row = total_row_tokens / len(data_rows)

    # 计算可用于数据行的 token 数
    available_tokens = target_tokens - fixed_overhead

    if available_tokens <= 0 or avg_tokens_per_row <= 0:
        return 1  # 极端情况，每个 chunk 只放 1 行

    suggested_rows = int(available_tokens / avg_tokens_per_row)

    # 限制在合理范围 [1, 20]
    return max(1, min(suggested_rows, 20))
```

#### 5.2.2 `run_pipeline(excel_path, keywords=None, max_rows_per_chunk=None, target_tokens=512, separator="!!!_CHUNK_BREAK_!!!")`

**功能**: 执行完整的流水线。

```python
def run_pipeline(
    excel_path: str,
    keywords: list = None,
    max_rows_per_chunk: int = None,
    target_tokens: int = 512,
    separator: str = "!!!_CHUNK_BREAK_!!!",
):
    """
    执行完整的 Excel -> HTML -> Chunks 流水线

    参数:
        excel_path: Excel 文件路径
        keywords: 关键检索词列表（用于幽灵标题）
        max_rows_per_chunk: 每个 chunk 的最大行数（如果指定，优先使用）
        target_tokens: 目标 token 数（当 max_rows 未指定时，自动计算行数）
        separator: chunk 之间的分隔符

    返回:
        dict: {
            'html_path': 中间 HTML 文件路径,
            'chunk_path': 最终 chunk 文件路径,
            'chunk_count': chunk 数量
        }
    """
    source_path = Path(excel_path)

    if not source_path.exists():
        print(f"❌ 错误：找不到文件 '{source_path}'")
        return None

    print("=" * 50)
    print(f"🚀 开始处理流水线: {source_path.name}")
    print("=" * 50)

    # === 第一步：Excel -> HTML ===
    print("\n📌 第一步：Excel 转 HTML（增强版）")
    html_path = convert_excel_to_html(
        excel_path=str(source_path),
        keywords=keywords,
        output_path=None,  # 默认保存到同目录
    )

    if not html_path:
        print("❌ 流水线中断：HTML 转换失败")
        return None

    # === 第二步：HTML -> Chunks ===
    print("\n📌 第二步：HTML 切分为 Chunks")
    html_content = Path(html_path).read_text(encoding="utf-8")

    # 自动计算或使用指定的行数
    if max_rows_per_chunk is None:
        # 使用 token 模式：逐行累加，精确控制每个 chunk 的 token 数
        print(f"📊 使用 token 模式，目标每 chunk ≤ {target_tokens} tokens")
        chunks = distribute_assets_and_chunk(
            html_content,
            max_rows_per_chunk=None,
            max_tokens_per_chunk=target_tokens
        )
    else:
        # 使用行数模式
        print(f"📊 使用行数模式，每 chunk {max_rows_per_chunk} 行")
        chunks = distribute_assets_and_chunk(
            html_content,
            max_rows_per_chunk=max_rows_per_chunk,
            max_tokens_per_chunk=None
        )
    print(f"🔪 切分完成：共生成 {len(chunks)} 个片段")

    # 保存 chunk 结果
    chunk_path = source_path.with_suffix("").with_name(
        source_path.stem + "_chunk_merged.html"
    )

    formatted_separator = f"\n\n{separator}\n\n"
    merged_content = formatted_separator.join(chunks)

    try:
        chunk_path.write_text(merged_content, encoding="utf-8")
        print(f"✅ Chunk 文件已保存: {chunk_path.absolute()}")
    except IOError as e:
        print(f"❌ 写入 Chunk 文件失败: {e}")
        return None

    # === 完成 ===
    print("\n" + "=" * 50)
    print("🎉 流水线执行完成！")
    print(f"   📄 中间结果 (HTML): {html_path}")
    print(f"   📄 最终结果 (Chunks): {chunk_path}")
    print(f"   🔢 Chunk 数量: {len(chunks)}")
    print(f"   🔑 分隔符: {separator}")
    print("=" * 50)

    return {
        "html_path": html_path,
        "chunk_path": str(chunk_path),
        "chunk_count": len(chunks),
    }
```

### 5.3 命令行接口

```python
def main():
    """命令行入口"""
    parser = argparse.ArgumentParser(
        description="Excel 转 HTML 完整流水线（增强版 + Chunk 切分）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python pipeline.py input.xlsx
  python pipeline.py input.xlsx -k "财务报表" "年度收入"
  python pipeline.py input.xlsx -t 512          # 基于 512 tokens 自动计算行数
  python pipeline.py input.xlsx -r 5            # 固定每 chunk 5 行
  python pipeline.py input.xlsx -t 1024 -s "---SPLIT---"
        """,
    )
    parser.add_argument("excel_file", help="要转换的 Excel 文件路径")
    parser.add_argument(
        "-k", "--keywords", nargs="+", help="关键检索词（用于幽灵标题）"
    )
    parser.add_argument(
        "-r",
        "--max-rows",
        type=int,
        default=None,
        help="每个 chunk 的最大数据行数（指定后忽略 -t 参数）",
    )
    parser.add_argument(
        "-t",
        "--target-tokens",
        type=int,
        default=512,
        help="目标 token 数，自动计算行数（默认: 512）",
    )
    parser.add_argument(
        "-s",
        "--separator",
        default="!!!_CHUNK_BREAK_!!!",
        help="chunk 之间的分隔符（默认: !!!_CHUNK_BREAK_!!!）",
    )

    args = parser.parse_args()

    run_pipeline(
        excel_path=args.excel_file,
        keywords=args.keywords,
        max_rows_per_chunk=args.max_rows,
        target_tokens=args.target_tokens,
        separator=args.separator,
    )
```

---

## 6. 使用指南

### 6.1 基本用法

```bash
# 进入项目目录
cd src/excel2html

# 基本转换（使用默认 512 tokens 切分）
python pipeline.py ../../Files/Excel/本国子目注释调整表.xlsx

# 带关键词（幽灵标题）
python pipeline.py ../../Files/Excel/本国子目注释调整表.xlsx -k "关税" "税则" "注释"

# 指定目标 token 数
python pipeline.py input.xlsx -t 1024

# 使用行数模式
python pipeline.py input.xlsx -r 5

# 自定义分隔符
python pipeline.py input.xlsx -s "---CHUNK_SEPARATOR---"

# 完整参数示例
python pipeline.py input.xlsx -k "财务" "报表" -t 512 -s "!!!SPLIT!!!"
```

### 6.2 参数说明

| 参数 | 短选项 | 长选项 | 默认值 | 说明 |
|------|--------|--------|--------|------|
| Excel 文件 | - | - | 必填 | 要转换的 Excel 文件路径 |
| 关键词 | `-k` | `--keywords` | 无 | 关键检索词列表，用于幽灵标题 |
| 最大行数 | `-r` | `--max-rows` | 无 | 每个 chunk 的最大行数（指定后忽略 -t） |
| 目标 Token | `-t` | `--target-tokens` | 512 | 每个 chunk 的目标 token 数 |
| 分隔符 | `-s` | `--separator` | `!!!_CHUNK_BREAK_!!!` | chunk 之间的分隔符 |

### 6.3 输出文件

假设输入文件为 `Files/Excel/财务报表.xlsx`：

| 文件 | 路径 | 说明 |
|------|------|------|
| 中间结果 | `Files/Excel/财务报表.html` | 增强后的完整 HTML |
| 最终结果 | `Files/Excel/财务报表_chunk_merged.html` | 切分后的 chunks |

### 6.4 输出示例

```
==================================================
🚀 开始处理流水线: 本国子目注释调整表.xlsx
==================================================

📌 第一步：Excel 转 HTML（增强版）
📄 正在处理: 本国子目注释调整表.xlsx
   增强功能: 上下文硬编码 ✓ | 表头降维 ✓ | 合并单元格 ✓ | 幽灵标题 ✓ (3个关键词)
✅ 转换成功！输出: C:\...\Files\Excel\本国子目注释调整表.html

📌 第二步：HTML 切分为 Chunks
📊 使用 token 模式，目标每 chunk ≤ 512 tokens
🔪 切分完成：共生成 15 个片段
✅ Chunk 文件已保存: C:\...\Files\Excel\本国子目注释调整表_chunk_merged.html

==================================================
🎉 流水线执行完成！
   📄 中间结果 (HTML): C:\...\Files\Excel\本国子目注释调整表.html
   📄 最终结果 (Chunks): C:\...\Files\Excel\本国子目注释调整表_chunk_merged.html
   🔢 Chunk 数量: 15
   🔑 分隔符: !!!_CHUNK_BREAK_!!!
==================================================
```

---

## 7. RAG 增强技术详解

### 7.1 增强技术 1：上下文硬编码

**问题**: 传统转换后的 HTML 缺少元数据，检索时无法定位来源。

**解决方案**: 在每个表格前注入上下文信息。

**实现代码**:
```python
update_time = datetime.now().strftime("%Y-%m-%d")
context_html = f"""<div class="rag-context">【文档上下文】来源文件：{filename} | 工作表：{sheet.title} | 数据类型：表格数据 | 更新时间：{update_time}</div>"""
```

**输出示例**:
```html
<div class="rag-context">【文档上下文】来源文件：财务报表.xlsx | 工作表：Sheet1 | 数据类型：表格数据 | 更新时间：2025-01-12</div>
```

**RAG 效果**:
- 用户查询"财务报表.xlsx 中的数据"时，可以精确匹配
- 用户查询"Sheet1 的内容"时，可以定位到具体工作表
- 提供时间戳，便于判断数据时效性

### 7.2 增强技术 2：幽灵标题（Ghost Caption）

**问题**: 表格内容可能不包含用户常用的检索词汇。

**解决方案**: 在 `<caption>` 中注入关键检索词。

**实现代码**:
```python
if keywords:
    keyword_str = "，".join(keywords)
    caption_html = f"    <caption>关键检索词：{keyword_str}。此表可能包含相关问题的答案。</caption>"
```

**输出示例**:
```html
<caption>关键检索词：关税，税则，注释。此表可能包含相关问题的答案。</caption>
```

**RAG 效果**:
- 用户查询"关税相关规定"时，即使表格内容中没有"关税"二字，也能被检索到
- 提供语义提示"此表可能包含相关问题的答案"，帮助 LLM 理解内容相关性

**使用建议**:
- 关键词应该是用户可能使用的检索词汇
- 包含同义词、缩写、常见问法
- 不宜过多，3-5 个为佳

### 7.3 增强技术 3：表头降维

**问题**: 多层表头在切分后丢失层级关系。

**解决方案**: 将多层表头"降维"为单行，把父级标题拼接到子级标题。

**原始多层表头**:
```
┌─────────────────────────────┬─────────────────────────────┐
│          财务数据            │          人员数据            │
├──────────┬──────────────────┼──────────┬──────────────────┤
│   收入    │       支出       │   在职    │       离职       │
└──────────┴──────────────────┴──────────┴──────────────────┘
```

**降维后**:
```html
<thead>
    <tr>
        <th>财务数据-收入</th>
        <th>财务数据-支出</th>
        <th>人员数据-在职</th>
        <th>人员数据-离职</th>
    </tr>
</thead>
```

**RAG 效果**:
- 每个 chunk 的表头都包含完整的层级信息
- 用户查询"财务数据的收入"时，可以精确匹配"财务数据-收入"列
- 避免了切分后"收入"列失去"财务数据"上下文的问题

### 7.4 增强技术 4：全局资产分发

**问题**: 切分后的 chunk 丢失上下文、表头等信息。

**解决方案**: 提取"全局资产"，在切分时复制到每个 chunk。

**全局资产包括**:
1. `<div class="rag-context">` - 上下文信息
2. `<caption>` - 幽灵标题
3. `<thead>` - 表头

**实现逻辑**:
```python
def build_chunk(data_rows_for_chunk):
    """组装一个 chunk"""
    new_soup = BeautifulSoup("<div></div>", "html.parser")
    wrapper_div = new_soup.div

    # 复制上下文
    if context_div:
        wrapper_div.append(copy.copy(context_div))

    new_table = new_soup.new_tag("table")
    new_table.attrs = original_table.attrs

    # 复制幽灵标题
    if caption:
        new_table.append(copy.copy(caption))

    # 复制表头
    new_thead = new_soup.new_tag("thead")
    for h_row in header_rows:
        new_thead.append(copy.copy(h_row))
    new_table.append(new_thead)

    # 添加数据行
    new_tbody = new_soup.new_tag("tbody")
    for d_row in data_rows_for_chunk:
        new_tbody.append(copy.copy(d_row))
    new_table.append(new_tbody)

    wrapper_div.append(new_table)
    return str(wrapper_div)
```

**RAG 效果**:
- 每个 chunk 都是自包含的，可以独立被检索和理解
- 即使只检索到一个 chunk，也能知道数据来源和列含义

### 7.5 增强技术 5：合并单元格智能处理

**问题**: 合并单元格在 HTML 中需要正确的 `rowspan`/`colspan` 属性。

**解决方案**: 遍历所有合并区域，记录每个单元格的跨度信息。

**实现逻辑**:
```python
for merged_range in sheet.merged_cells.ranges:
    min_row, min_col = merged_range.min_row, merged_range.min_col
    max_row, max_col = merged_range.max_row, merged_range.max_col
    origin_value = sheet.cell(row=min_row, column=min_col).value

    rowspan = max_row - min_row + 1
    colspan = max_col - min_col + 1

    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            is_origin = r == min_row and c == min_col
            merged_info[(r, c)] = {
                "value": origin_value,
                "rowspan": rowspan if is_origin else 0,
                "colspan": colspan if is_origin else 0,
                "is_origin": is_origin,
                "skip": not is_origin,
            }
```

**输出示例**:
```html
<td rowspan="3" colspan="2">合并单元格内容</td>
```

---

## 8. Token 估算与动态切分

### 8.1 Token 估算原理

**为什么需要估算 Token？**

RAG 系统中，每个 chunk 需要被 embedding 模型处理。大多数 embedding 模型有 token 限制（如 512 或 1024）。如果 chunk 过大，会被截断导致信息丢失。

**估算公式**:
```python
def estimate_tokens(text: str) -> int:
    """估算文本的 token 数量（中文约2.5字符=1token）"""
    return int(len(text) / 2.5)
```

**估算依据**:
| 内容类型 | 字符/Token 比例 |
|---------|----------------|
| 中文字符 | 约 1-2 字符 = 1 token |
| 英文单词 | 约 4 字符 = 1 token |
| HTML 标签 | 按字符数估算 |
| 综合取值 | 2.5 字符 = 1 token |

**精确计算方案**:

如需精确计算，可接入 `tiktoken` 库：

```python
import tiktoken

def estimate_tokens_precise(text: str, model: str = "cl100k_base") -> int:
    """使用 tiktoken 精确计算 token 数量"""
    encoding = tiktoken.get_encoding(model)
    return len(encoding.encode(text))
```

### 8.2 动态切分算法

**算法流程图**:

```
┌─────────────────────────────────────────────────────────────────┐
│  输入: HTML 内容 + 目标 token 数 (如 512)                        │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│  Step 1: 计算固定开销                                            │
│  fixed_overhead = tokens(context_div + caption + thead)         │
│  例如: 150 tokens                                               │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│  Step 2: 初始化                                                  │
│  current_chunk_data = []                                        │
│  current_chunk_tokens = 0                                       │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│  Step 3: 遍历每一行                                              │
│  for row in data_rows:                                          │
│      row_tokens = estimate_tokens(row)                          │
│                                                                 │
│      if (fixed_overhead + current_tokens + row_tokens) > 512:   │
│          → 切分！保存当前 chunk，重置计数器                        │
│                                                                 │
│      current_chunk_data.append(row)                             │
│      current_chunk_tokens += row_tokens                         │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│  输出: chunks 列表，每个 chunk 的 token ≤ target_tokens          │
└─────────────────────────────────────────────────────────────────┘
```

### 8.3 切分示例

**输入数据**:
- 目标 token: 512
- 固定开销: 150 tokens
- 数据行 token 分布: [50, 60, 200, 80, 300, 40, 50]

**切分过程**:

```
初始状态:
  current_chunk_tokens = 0
  current_chunk_data = []

处理 Row 1 (50 tokens):
  检查: 150 + 0 + 50 = 200 ≤ 512 ✓
  累加: current_chunk_tokens = 50, data = [Row1]

处理 Row 2 (60 tokens):
  检查: 150 + 50 + 60 = 260 ≤ 512 ✓
  累加: current_chunk_tokens = 110, data = [Row1, Row2]

处理 Row 3 (200 tokens):
  检查: 150 + 110 + 200 = 460 ≤ 512 ✓
  累加: current_chunk_tokens = 310, data = [Row1, Row2, Row3]

处理 Row 4 (80 tokens):
  检查: 150 + 310 + 80 = 540 > 512 ✗
  → 切分！保存 Chunk 1 = [Row1, Row2, Row3]
  重置: current_chunk_tokens = 80, data = [Row4]

处理 Row 5 (300 tokens):
  检查: 150 + 80 + 300 = 530 > 512 ✗
  → 切分！保存 Chunk 2 = [Row4]
  重置: current_chunk_tokens = 300, data = [Row5]

处理 Row 6 (40 tokens):
  检查: 150 + 300 + 40 = 490 ≤ 512 ✓
  累加: current_chunk_tokens = 340, data = [Row5, Row6]

处理 Row 7 (50 tokens):
  检查: 150 + 340 + 50 = 540 > 512 ✗
  → 切分！保存 Chunk 3 = [Row5, Row6]
  重置: current_chunk_tokens = 50, data = [Row7]

结束:
  → 保存 Chunk 4 = [Row7]

最终结果:
  Chunk 1: [Row1, Row2, Row3] → 150 + 310 = 460 tokens
  Chunk 2: [Row4]             → 150 + 80  = 230 tokens
  Chunk 3: [Row5, Row6]       → 150 + 340 = 490 tokens
  Chunk 4: [Row7]             → 150 + 50  = 200 tokens
```

### 8.4 参数选择建议

| 场景 | 建议参数 | 说明 |
|------|---------|------|
| OpenAI text-embedding-ada-002 | `-t 512` | 模型限制 8191 tokens，但较短 chunk 检索效果更好 |
| OpenAI text-embedding-3-small | `-t 512` | 同上 |
| 本地小模型 | `-t 256` | 较短 chunk 适合小模型 |
| 长文档检索 | `-t 1024` | 保留更多上下文 |
| 行内容长度相近 | `-r 5` | 行数模式更简单 |

---

## 附录 A：完整代码清单

### A.1 excel2html_openpyxl_enhanced.py

见 `src/excel2html/excel2html_openpyxl_enhanced.py`

### A.2 html2chunk.py

见 `src/excel2html/html2chunk.py`

### A.3 pipeline.py

见 `src/excel2html/pipeline.py`

---

## 附录 B：依赖安装

```bash
pip install openpyxl beautifulsoup4
```

或使用 uv：

```bash
uv add openpyxl beautifulsoup4
```

---

## 附录 C：常见问题

### Q1: 为什么选择 HTML 而不是 Markdown？

A: HTML 保留了更多结构信息（如 `rowspan`/`colspan`），且 `<caption>`、`<thead>` 等标签有明确的语义，便于后续处理。

### Q2: 如何处理超大 Excel 文件？

A: 当前实现会将整个文件加载到内存。对于超大文件，建议：
1. 使用 `openpyxl` 的 `read_only=True` 模式
2. 分 Sheet 处理
3. 流式写入输出文件

### Q3: 如何自定义 Token 估算？

A: 修改 `html2chunk.py` 中的 `estimate_tokens` 函数，可接入 `tiktoken` 或其他分词器。

### Q4: 分隔符的作用是什么？

A: 分隔符用于在后续处理中将 chunks 分开。选择一个不会出现在正常内容中的字符串即可。

---

*文档版本: 1.0*
*最后更新: 2025-01-12*
