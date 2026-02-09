"""
打包为可执行文件
"""

import PyInstaller.__main__
import os
import sys
import shutil
from pathlib import Path

def build_executable():
    """构建可执行文件"""
    
    # 清理之前的构建文件
    build_dir = Path("build")
    dist_dir = Path("dist")
    
    if build_dir.exists():
        shutil.rmtree(build_dir)
    if dist_dir.exists():
        shutil.rmtree(dist_dir)
    
    # 创建必要的目录
    Path("output").mkdir(exist_ok=True)
    Path("templates").mkdir(exist_ok=True)
    Path("config").mkdir(exist_ok=True)
    
    # PyInstaller参数
    pyinstaller_args = [
        "src/main.py",
        "--name=AcademicWriterPro",
        "--onefile",
        "--windowed",
        "--icon=assets/icon.ico" if Path("assets/icon.ico").exists() else "",
        "--add-data=src;src",
        "--hidden-import=tkinter",
        "--hidden-import=tkinterdnd2",
        "--hidden-import=PIL",
        "--hidden-import=openai",
        "--hidden-import=requests",
        "--clean",
        "--noconfirm",
    ]
    
    # 过滤空参数
    pyinstaller_args = [arg for arg in pyinstaller_args if arg]
    
    try:
        print("开始构建可执行文件...")
        PyInstaller.__main__.run(pyinstaller_args)
        print("构建完成！")
        
        # 复制配置文件
        if dist_dir.exists():
            for dir_name in ["output", "templates", "config"]:
                src_dir = Path(dir_name)
                if src_dir.exists():
                    dst_dir = dist_dir / "AcademicWriterPro" / dir_name
                    if dst_dir.exists():
                        shutil.rmtree(dst_dir)
                    shutil.copytree(src_dir, dst_dir)
            
            print(f"可执行文件位于: {dist_dir / 'AcademicWriterPro'}")
            
    except Exception as e:
        print(f"构建失败: {e}")
        sys.exit(1)

if __name__ == "__main__":
    build_executable()
