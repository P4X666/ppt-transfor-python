#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPTX 转换路由

提供 PPTX 文件上传、转换和结果下载功能
"""

from fastapi import APIRouter, File, UploadFile, HTTPException, status, BackgroundTasks
from fastapi.responses import JSONResponse, FileResponse
from pydantic import BaseModel, Field
from typing import Optional, Any
from pathlib import Path
import tempfile
import json
import shutil
import os

from services.pptx_converter import convert_pptx_to_json

router = APIRouter()

class ConversionResponse(BaseModel):
    """转换结果响应模型"""
    status: str = Field(..., description="转换状态")
    message: str = Field(..., description="状态消息")
    task_id: str = Field(..., description="任务ID")
    slide_count: Optional[int] = Field(None, description="幻灯片数量")
    image_count: Optional[int] = Field(None, description="图片数量")

class ConversionResult(BaseModel):
    """转换结果详细信息"""
    task_id: str = Field(..., description="任务ID")
    status: str = Field(..., description="转换状态")
    data: Optional[dict[str, Any]] = Field(None, description="转换后的数据")

# 存储正在处理的任务
tasks: dict[str, dict] = {}

@router.post("/convert", response_model=ConversionResponse, status_code=status.HTTP_200_OK)
async def convert_pptx(
    file: UploadFile = File(..., description="PPTX文件"),
    extract_images: bool = True,
    background_tasks: BackgroundTasks = None
):
    """
    上传并转换 PPTX 文件
    
    - **file**: 要转换的 PPTX 文件
    - **extract_images**: 是否提取图片（默认：True）
    
    返回任务ID和转换摘要信息
    """
    # 验证文件类型
    if not file.filename.lower().endswith(".pptx"):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="文件类型错误，仅支持 PPTX 格式"
        )
    
    # 验证文件大小（最大50MB）
    if file.size and file.size > 50 * 1024 * 1024:
        raise HTTPException(
            status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
            detail="文件大小超过限制（最大50MB）"
        )
    
    try:
        # 创建临时目录
        task_id = os.urandom(16).hex()
        temp_dir = Path(tempfile.mkdtemp(prefix=f"pptx_{task_id}_"))
        
        # 保存上传的文件
        input_path = temp_dir / "input.pptx"
        contents = await file.read()
        input_path.write_bytes(contents)
        
        # 执行转换
        result = convert_pptx_to_json(
            pptx_path=str(input_path),
            extract_images=extract_images,
            images_output_dir=str(temp_dir / "images")
        )
        
        # 保存结果到临时目录
        result_path = temp_dir / "result.json"
        result_path.write_text(
            json.dumps(result, ensure_ascii=False, indent=2),
            encoding="utf-8"
        )
        
        # 存储任务信息
        tasks[task_id] = {
            "status": "completed",
            "temp_dir": str(temp_dir),
            "slide_count": result["meta"]["slide_count"],
            "image_count": len(result["images"]),
            "result": result
        }
        
        return ConversionResponse(
            status="success",
            message="转换完成",
            task_id=task_id,
            slide_count=result["meta"]["slide_count"],
            image_count=len(result["images"])
        )
        
    except Exception as e:
        # 清理临时文件
        if 'temp_dir' in locals() and temp_dir.exists():
            shutil.rmtree(temp_dir, ignore_errors=True)
        
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"转换失败: {str(e)}"
        )

@router.get("/result/{task_id}", response_model=ConversionResult)
async def get_conversion_result(task_id: str):
    """
    获取转换结果
    
    - **task_id**: 任务ID
    
    返回完整的转换结果JSON数据
    """
    task = tasks.get(task_id)
    
    if not task:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="任务不存在或已过期"
        )
    
    return ConversionResult(
        task_id=task_id,
        status=task["status"],
        data=task["result"]
    )

@router.get("/download/{task_id}")
async def download_result(task_id: str):
    """
    下载转换结果文件（ZIP格式）
    
    - **task_id**: 任务ID
    
    返回包含JSON结果和图片的ZIP文件
    """
    task = tasks.get(task_id)
    
    if not task:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="任务不存在或已过期"
        )
    
    temp_dir = Path(task["temp_dir"])
    result_path = temp_dir / "result.json"
    
    if not result_path.exists():
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="结果文件不存在"
        )
    
    return FileResponse(
        path=str(result_path),
        filename=f"pptx_result_{task_id}.json",
        media_type="application/json"
    )

@router.get("/download/images/{task_id}")
async def download_images(task_id: str):
    """
    下载提取的图片文件（ZIP格式）
    
    - **task_id**: 任务ID
    
    返回包含所有图片的ZIP文件
    """
    task = tasks.get(task_id)
    
    if not task:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="任务不存在或已过期"
        )
    
    temp_dir = Path(task["temp_dir"])
    images_dir = temp_dir / "images"
    
    if not images_dir.exists() or not any(images_dir.iterdir()):
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="图片目录不存在或为空"
        )
    
    # 创建ZIP文件
    zip_path = temp_dir / "images.zip"
    shutil.make_archive(str(zip_path.with_suffix("")), 'zip', str(images_dir))
    
    return FileResponse(
        path=str(zip_path),
        filename=f"pptx_images_{task_id}.zip",
        media_type="application/zip"
    )

@router.delete("/task/{task_id}")
async def delete_task(task_id: str):
    """
    删除任务及其相关文件
    
    - **task_id**: 任务ID
    
    清理任务数据和临时文件
    """
    task = tasks.get(task_id)
    
    if not task:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="任务不存在"
        )
    
    # 清理临时目录
    temp_dir = Path(task["temp_dir"])
    if temp_dir.exists():
        shutil.rmtree(temp_dir, ignore_errors=True)
    
    # 移除任务记录
    del tasks[task_id]
    
    return {"status": "success", "message": "任务已删除"}

@router.get("/tasks")
async def list_tasks():
    """
    获取所有任务列表（仅ID和状态）
    """
    return [
        {"task_id": task_id, "status": task["status"], "slide_count": task.get("slide_count")}
        for task_id, task in tasks.items()
    ]