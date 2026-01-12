from pathlib import Path
from unstructured.partition.xlsx import partition_xlsx
from unstructured.documents.elements import Table, Title


def convert_excel_to_html_file(excel_path):
    """读取 Excel，提取表格的 HTML 结构，并保存为同名 HTML 文件。"""
    source_path = excel_path if isinstance(excel_path, Path) else Path(excel_path)

    if not source_path.exists():
        print(f"❌ 错误：找不到文件 '{source_path}'")
        return

    output_path = source_path.with_suffix(".html")
    print(f"正在处理: {source_path.name}")

    try:
        elements = partition_xlsx(
            filename=str(source_path),
            mode="elements",
            include_metadata=True,
            infer_table_structure=True,
        )
    except Exception as e:
        print(f"❌ 解析失败: {e}")
        return

    html_content = []

    html_header = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>Excel Export</title>
</head>
<body>"""
    html_content.append(html_header)

    for element in elements:
        if isinstance(element, Table):
            if (
                hasattr(element.metadata, "text_as_html")
                and element.metadata.text_as_html
            ):
                html_content.append(element.metadata.text_as_html)
            else:
                html_content.append(f"<pre>{element.text}</pre>")
        elif isinstance(element, Title):
            html_content.append(f"<h2>{element.text}</h2>")
        else:
            html_content.append(f"<p>{element.text}</p>")

    html_content.append("</body></html>")

    try:
        output_path.write_text("\n".join(html_content), encoding="utf-8")
        print(f"✅ 转换成功！文件已保存至: {output_path.absolute()}")
    except IOError as e:
        print(f"❌ 写入文件失败: {e}")


def convert_folder(folder_path_str: str):
    """批量处理指定文件夹下的所有 Excel 文件"""
    folder = Path(folder_path_str)

    if not folder.exists():
        print(f"❌ 错误：找不到文件夹 '{folder}'")
        return

    if not folder.is_dir():
        print(f"❌ 错误：'{folder}' 不是一个文件夹")
        return

    excel_files = [
        f
        for f in list(folder.glob("*.xlsx")) + list(folder.glob("*.xls"))
        if not f.name.startswith("~$")
    ]

    if not excel_files:
        print(f"⚠️ 文件夹 '{folder}' 中没有找到 Excel 文件")
        return

    print(f"📁 找到 {len(excel_files)} 个 Excel 文件\n")

    for excel_file in excel_files:
        convert_excel_to_html_file(excel_file)

    print("\n🎉 处理完成！")


if __name__ == "__main__":
    target_folder = (
        r"C:\Users\Administrator\Desktop\玄通\通用知识库_handled\2026年税则调整"
    )
    convert_folder(target_folder)
