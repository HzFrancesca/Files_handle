"""
业务处理函数 - Excel 转换和预览
"""

import os
import shutil
import tempfile
from pathlib import Path

from src.core.excel2html import convert_excel_to_html, distribute_assets_and_chunk


# 全局变量存储当前结果路径
current_html_path = None
current_chunk_path = None


def process_excel(
    excel_file,
    keywords_text: str,
    split_mode: str,
    max_rows: int,
    target_tokens: int,
    separator: str,
):
    """处理 Excel 文件的主函数"""
    global current_html_path, current_chunk_path
    
    if excel_file is None:
        return None, None, "⚠️ 请先上传 Excel 文件"

    # 解析关键词
    keywords = None
    if keywords_text.strip():
        keywords = [k.strip() for k in keywords_text.split(",") if k.strip()]

    # 创建临时目录
    temp_dir = Path(tempfile.mkdtemp())
    
    try:
        # 复制上传的文件到临时目录
        source_path = Path(excel_file.name)
        temp_excel = temp_dir / source_path.name
        shutil.copy(excel_file.name, temp_excel)

        # 第一步：Excel -> HTML
        html_path = convert_excel_to_html(
            excel_path=str(temp_excel),
            keywords=keywords,
            output_path=None,
        )

        if not html_path:
            return None, None, "❌ Excel 转换失败，请检查文件格式"

        # 读取 HTML 内容
        html_content = Path(html_path).read_text(encoding="utf-8")

        # 第二步：HTML -> Chunks
        if split_mode == "按 Token 数":
            result = distribute_assets_and_chunk(
                html_content,
                max_rows_per_chunk=None,
                max_tokens_per_chunk=target_tokens,
            )
        else:
            result = distribute_assets_and_chunk(
                html_content,
                max_rows_per_chunk=max_rows,
                max_tokens_per_chunk=None,
            )

        chunks = result["chunks"]
        warnings = result["warnings"]
        stats = result["stats"]

        # 合并 chunks
        formatted_separator = f"\n\n{separator}\n\n"
        merged_content = formatted_separator.join(chunks)

        # 保存最终结果（使用原始文件名）
        html_output_name = f"{source_path.stem}_middle.html"
        chunk_output_name = f"{source_path.stem}.html"
        
        chunk_path = temp_dir / chunk_output_name
        chunk_path.write_text(merged_content, encoding="utf-8")

        # 将 HTML 文件复制到带有正确文件名的路径
        # 因为 convert_excel_to_html 可能生成在不同位置
        html_final_path = temp_dir / html_output_name
        if str(html_path) != str(html_final_path):
            shutil.copy(html_path, html_final_path)
            html_path = str(html_final_path)
        else:
            html_path = str(html_path)

        # 更新全局路径
        current_html_path = html_path
        current_chunk_path = str(chunk_path)

        # 生成状态信息
        warning_text = ""
        if warnings:
            warning_text = f"\n⚠️ 警告：{len(warnings)} 个片段超过 token 限制"
            for w in warnings:
                warning_text += f"\n   - 片段 #{w['chunk_index']}: {w['actual_tokens']} tokens (超出 {w['overflow']})"

        status = f"""✅ 处理完成

源文件：{source_path.name}
关键词：{', '.join(keywords) if keywords else '无'}
切分模式：{split_mode}
生成 Chunks：{len(chunks)} 个
Token 统计：最小={stats['min_token_count']}, 最大={stats['max_token_count']}, 平均={stats['avg_token_count']:.1f}
分隔符：{separator}{warning_text}"""

        return html_path, str(chunk_path), status

    except Exception as e:
        current_html_path = None
        current_chunk_path = None
        return None, None, f"❌ 处理出错: {str(e)}"


def get_html_preview():
    """获取 HTML 预览内容"""
    global current_html_path
    if current_html_path and os.path.exists(current_html_path):
        content = Path(current_html_path).read_text(encoding="utf-8")
        return f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>中间结果预览</title>
    <style>
        body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 40px; max-width: 1200px; margin: auto; }}
        table {{ border-collapse: collapse; width: 100%; margin: 20px 0; }}
        th, td {{ border: 1px solid #ddd; padding: 12px; text-align: left; }}
        th {{ background: #f5f5f5; font-weight: 600; }}
        tr:hover {{ background: #fafafa; }}
        .rag-context {{ background: #e3f2fd; padding: 16px; border-radius: 8px; margin-bottom: 20px; color: #1565c0; }}
        caption {{ font-size: 0.9rem; color: #666; margin-bottom: 12px; }}
    </style>
</head>
<body>
    <h2>📄 中间结果预览</h2>
    {content}
</body>
</html>"""
    return None


def get_chunk_preview():
    """获取 Chunk 预览内容"""
    global current_chunk_path
    if current_chunk_path and os.path.exists(current_chunk_path):
        content = Path(current_chunk_path).read_text(encoding="utf-8")
        content = content.replace(
            "!!!_CHUNK_BREAK_!!!",
            '</div><hr style="border: 2px dashed #2563eb; margin: 40px 0;"><div style="background:#f8f9fa; padding: 8px 16px; border-radius: 4px; color: #666; font-size: 0.85rem; margin-bottom: 20px;">📦 Chunk 分隔</div><div class="chunk">'
        )
        return f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Chunks 预览</title>
    <style>
        body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 40px; max-width: 1200px; margin: auto; }}
        table {{ border-collapse: collapse; width: 100%; margin: 20px 0; }}
        th, td {{ border: 1px solid #ddd; padding: 12px; text-align: left; }}
        th {{ background: #f5f5f5; font-weight: 600; }}
        tr:hover {{ background: #fafafa; }}
        .rag-context {{ background: #e3f2fd; padding: 16px; border-radius: 8px; margin-bottom: 20px; color: #1565c0; }}
        caption {{ font-size: 0.9rem; color: #666; margin-bottom: 12px; }}
        .chunk {{ margin-bottom: 20px; }}
    </style>
</head>
<body>
    <h2>📦 Chunks 预览</h2>
    <div class="chunk">
    {content}
    </div>
</body>
</html>"""
    return None
