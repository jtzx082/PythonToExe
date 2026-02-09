#!/usr/bin/env python3
"""
简化的构建脚本，用于 GitHub Actions
"""

import sys
import os
from pathlib import Path
import subprocess

def build_executable():
    """构建可执行文件"""
    
    # 确保必要的目录存在
    Path("output").mkdir(exist_ok=True)
    Path("templates").mkdir(exist_ok=True)
    Path("config").mkdir(exist_ok=True)
    
    # 构建参数
    build_args = [
        sys.executable, "-m", "PyInstaller",
        "src/main.py",
        "--name=AcademicWriterPro",
        "--onefile",
        "--windowed",
        "--add-data=src:src",
        "--hidden-import=tkinter",
        "--hidden-import=tkinterdnd2",
        "--hidden-import=PIL",
        "--hidden-import=openai",
        "--hidden-import=requests",
        "--clean",
        "--noconfirm",
    ]
    
    # 根据平台调整参数
    if sys.platform == "win32":
        # Windows 特定设置
        build_args.extend(["--icon=assets/icon.ico"])
    elif sys.platform == "darwin":
        # macOS 特定设置
        build_args.extend(["--icon=assets/icon.icns"])
    
    print("Building with arguments:", build_args)
    
    try:
        result = subprocess.run(build_args, check=True, capture_output=True, text=True)
        print("Build output:", result.stdout)
        if result.stderr:
            print("Build errors:", result.stderr)
        
        # 检查输出文件
        dist_dir = Path("dist")
        if dist_dir.exists():
            files = list(dist_dir.iterdir())
            print(f"Build completed. Files in dist: {files}")
        else:
            print("ERROR: dist directory not created")
            sys.exit(1)
            
    except subprocess.CalledProcessError as e:
        print(f"Build failed with error: {e}")
        print(f"STDOUT: {e.stdout}")
        print(f"STDERR: {e.stderr}")
        sys.exit(1)

if __name__ == "__main__":
    build_executable()
