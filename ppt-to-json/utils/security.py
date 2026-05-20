#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
安全工具模块

提供路径验证、URL安全检查等安全相关功能
"""

import os
import re
from pathlib import Path


class SecurityError(Exception):
    """安全相关异常"""
    pass


def validate_local_file_under_base(file_path: str, base_dir: str) -> str:
    """
    验证文件路径是否在指定基础目录下，防止路径遍历攻击
    
    Args:
        file_path: 待验证的文件路径
        base_dir: 允许访问的基础目录
        
    Returns:
        安全的规范化路径
        
    Raises:
        SecurityError: 路径不在允许范围内
    """
    base_path = Path(base_dir).resolve()
    target_path = Path(file_path).resolve()
    
    if not str(target_path).startswith(str(base_path)):
        raise SecurityError(f"文件路径不在允许范围内: {file_path}")
    
    return str(target_path)


def validate_public_http_url(url: str) -> str:
    """
    验证URL是否为安全的公共HTTP(S) URL
    
    Args:
        url: 待验证的URL
        
    Returns:
        安全的URL
        
    Raises:
        SecurityError: URL不安全或格式不正确
    """
    # 检查基本格式
    if not (url.startswith("http://") or url.startswith("https://")):
        raise SecurityError(f"不支持的协议类型: {url}")
    
    # 禁止本地IP地址（防止SSRF攻击内网）
    local_patterns = [
        r"^http://127\.",
        r"^http://localhost",
        r"^http://0\.0\.0\.0",
        r"^http://192\.168\.",
        r"^http://10\.",
        r"^http://172\.(1[6-9]|2[0-9]|3[0-1])\.",
        r"^http://\[::1\]",
    ]
    
    for pattern in local_patterns:
        if re.match(pattern, url, re.IGNORECASE):
            raise SecurityError(f"禁止访问本地地址: {url}")
    
    return url