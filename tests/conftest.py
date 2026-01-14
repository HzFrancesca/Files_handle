"""
pytest 共享 fixtures
"""

from pathlib import Path

import pytest


@pytest.fixture
def sample_excel_path() -> Path:
    """示例 Excel 文件路径"""
    return Path("Files/关税配额商品税目税率表.xlsx")


@pytest.fixture
def sample_html_content() -> str:
    """示例 HTML 内容"""
    return """
<div class="rag-context">【文档上下文】来源：test.xlsx | 数据类型：表格数据</div>
<table border="1" style="border-collapse:collapse" data-source="test.xlsx" data-sheet="Sheet1">
    <caption>关键检索词：测试</caption>
    <thead>
        <tr>
            <th>列1</th>
            <th>列2</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td>数据1</td>
            <td>数据2</td>
        </tr>
        <tr>
            <td>数据3</td>
            <td>数据4</td>
        </tr>
    </tbody>
</table>
"""


@pytest.fixture
def temp_dir(tmp_path: Path) -> Path:
    """临时目录"""
    return tmp_path


@pytest.fixture
def source_files() -> list[Path]:
    """获取所有源文件"""
    return list(Path("src").rglob("*.py"))
