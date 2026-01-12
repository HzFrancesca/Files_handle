# Excel2HTML RAG 工具启动脚本
$ErrorActionPreference = "Stop"

# 激活虚拟环境并启动应用
& ".\.venv\Scripts\Activate.ps1"
python -m src.app.main
