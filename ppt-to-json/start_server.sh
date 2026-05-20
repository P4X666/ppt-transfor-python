#!/bin/bash
# PPTX 转换服务启动脚本

# 设置虚拟环境路径
VENV_PATH="I:/wefor/ai/knowledge/venv"
APP_PATH="I:/wefor/ai/knowledge/ppt-to-json"

# 激活虚拟环境
source "$VENV_PATH/Scripts/activate"

# 切换到应用目录
cd "$APP_PATH"

# 启动 FastAPI 服务
echo "Starting PPTX Converter API Server..."
uvicorn main:app --host 0.0.0.0 --port 8000 --workers 4

echo "Server stopped."