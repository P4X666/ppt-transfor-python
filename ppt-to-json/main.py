#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPTX 转 JSON FastAPI 应用

提供基于 python-pptx + docling-core 的 PPTX 文档转换服务
"""

from fastapi import FastAPI, HTTPException, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.openapi.docs import get_swagger_ui_html, get_redoc_html
from fastapi.openapi.utils import get_openapi
from pathlib import Path
import tempfile
import shutil

from routes import pptx

# 创建应用实例
app = FastAPI(
    title="PPTX 转换服务",
    description="基于 python-pptx + docling-core 的 PPTX 文档转换 API 服务",
    version="1.0.0",
    docs_url=None,
    redoc_url=None,
)

# CORS 配置
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 注册路由
app.include_router(pptx.router, prefix="/api/v1/pptx", tags=["PPTX转换"])

# 健康检查端点
@app.get("/health", tags=["健康检查"])
async def health_check():
    """服务健康检查"""
    return {"status": "healthy", "service": "pptx-converter"}

# 自定义 OpenAPI 文档
def custom_openapi():
    if app.openapi_schema:
        return app.openapi_schema
    openapi_schema = get_openapi(
        title="PPTX 转换服务",
        version="1.0.0",
        description="基于 python-pptx + docling-core 的 PPTX 文档转换 API 服务",
        routes=app.routes,
    )
    app.openapi_schema = openapi_schema
    return app.openapi_schema

app.openapi = custom_openapi

# Swagger UI 端点
@app.get("/docs", include_in_schema=False)
async def custom_swagger_ui_html():
    return get_swagger_ui_html(
        openapi_url="/openapi.json",
        title="PPTX 转换服务 - Swagger UI",
        swagger_favicon_url="https://fastapi.tiangolo.com/img/favicon.png",
    )

# ReDoc 端点
@app.get("/redoc", include_in_schema=False)
async def redoc_html():
    return get_redoc_html(
        openapi_url="/openapi.json",
        title="PPTX 转换服务 - ReDoc",
    )

# 临时目录清理（应用启动时执行）
@app.on_event("startup")
async def startup_event():
    """启动时清理临时目录"""
    temp_dir = Path(tempfile.gettempdir())
    for item in temp_dir.glob("pptx_*"):
        try:
            if item.is_dir():
                shutil.rmtree(item)
            else:
                item.unlink()
        except Exception:
            pass

@app.on_event("shutdown")
async def shutdown_event():
    """关闭时清理临时目录"""
    temp_dir = Path(tempfile.gettempdir())
    for item in temp_dir.glob("pptx_*"):
        try:
            if item.is_dir():
                shutil.rmtree(item)
            else:
                item.unlink()
        except Exception:
            pass

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)