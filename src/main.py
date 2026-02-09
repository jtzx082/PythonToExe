#!/usr/bin/env python3
"""
期刊论文撰写软件 - 主程序入口
支持：论文、计划、反思、案例、总结等文档撰写
"""

import sys
import os
import json
from pathlib import Path

# 添加src目录到Python路径
sys.path.insert(0, str(Path(__file__).parent))

from gui import AcademicWriterApp
from config import load_config, save_config

def main():
    """主函数"""
    # 确保必要的目录存在
    Path("output").mkdir(exist_ok=True)
    Path("templates").mkdir(exist_ok=True)
    Path("config").mkdir(exist_ok=True)
    
    # 加载配置
    config = load_config()
    
    # 创建并运行应用
    app = AcademicWriterApp(config)
    app.run()

if __name__ == "__main__":
    main()
