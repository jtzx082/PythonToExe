# 智能文稿撰写助手 - Academic Writer Pro

一个基于DeepSeek API的智能文稿撰写软件，支持期刊论文、研究计划、反思报告、案例分析、总结报告等多种文档类型。

## 功能特性

- 📝 **多文档类型支持**：期刊论文、计划、反思、案例、总结等
- 🤖 **AI智能生成**：基于DeepSeek API，生成高质量内容
- 🎯 **大纲导向**：先生成大纲，再完善内容，思路清晰
- 📊 **结构完整**：自动生成规范的学术文档结构
- 💾 **多格式导出**：支持Word、PDF、Markdown、纯文本格式
- 🌐 **多平台支持**：Windows、macOS、Linux全平台
- 🎨 **友好界面**：直观易用的图形用户界面

## 安装方法

### 方法一：直接运行源代码

1. 克隆仓库：
```bash
git clone https://github.com/yourusername/academic-writer.git
cd academic-writer
```
安装依赖：

bash
pip install -r requirements.txt
运行程序：

bash
python src/main.py

方法二：使用打包版本
从 Releases 页面下载对应平台的可执行文件。

配置说明
API密钥设置
获取DeepSeek API密钥：DeepSeek平台

在软件界面中输入API密钥

点击"测试连接"验证配置

文档类型
期刊论文：标准的学术论文结构

研究计划：科研项目申报书

反思报告：学习或工作反思

案例分析：商业或教育案例分析

总结报告：工作或项目总结

自定义类型：自定义文档结构

使用教程
基本流程
设置API密钥：在API设置区域输入您的DeepSeek API密钥

选择文档类型：在下拉框中选择要撰写的文档类型

输入标题：输入文档的标题

添加指令（可选）：输入额外的写作要求

生成大纲：点击"生成大纲"按钮，AI会自动生成文档大纲

编辑大纲：根据需要对大纲进行修改和调整

撰写文档：点击"撰写文档"按钮，生成完整文档

导出保存：将生成的文档导出为所需格式

高级功能
多标签编辑：可以分别编辑摘要、引言、方法等各部分

参数调整：调整温度、最大token数等生成参数

批量处理：支持批量生成多个文档

模板管理：保存和加载常用模板

开发指南
项目结构
text
academic-writer/
├── .github/workflows/    # GitHub Actions工作流
├── src/                  # 源代码
│   ├── main.py          # 主程序入口
│   ├── gui.py           # 图形界面
│   ├── api_client.py    # API客户端
│   ├── document_generator.py # 文档生成器
│   ├── config.py        # 配置文件
│   └── utils.py         # 工具函数
├── requirements.txt      # Python依赖
├── setup.py             # 安装配置
├── build_exe.py         # 打包脚本
└── README.md            # 说明文档
本地开发
bash
# 克隆项目
git clone https://github.com/yourusername/academic-writer.git
cd academic-writer

# 创建虚拟环境
python -m venv venv
source venv/bin/activate  # Linux/macOS
# 或
venv\Scripts\activate  # Windows

# 安装依赖
pip install -r requirements.txt

# 运行开发版本
python src/main.py
打包发布
bash
# 安装打包工具
pip install pyinstaller

# 打包为可执行文件
python build_exe.py
API使用说明
本软件使用DeepSeek API，需要有效的API密钥。以下是API调用的基本参数：

模型：deepseek-chat 或 deepseek-coder

温度：控制生成文本的随机性（0-2）

最大token数：控制生成文本的长度

注意事项
API限制：请遵守DeepSeek API的使用条款和限制

内容审查：生成的文档需要人工审查和修改

学术诚信：请遵守学术规范，合理使用AI辅助工具

数据安全：API密钥等敏感信息请妥善保管

贡献指南
欢迎提交Issue和Pull Request！以下是贡献步骤：

Fork本仓库

创建功能分支：git checkout -b feature/your-feature

提交更改：git commit -m 'Add some feature'

推送到分支：git push origin feature/your-feature

提交Pull Request

许可证
本项目采用MIT许可证。详见 LICENSE 文件。

联系方式
问题反馈：GitHub Issues

功能建议：通过Issue提交

邮箱：contact@example.com

更新日志
v1.0.0 (2024-01-01)
初始版本发布

支持基本的文档生成功能

图形用户界面

多平台打包支持

提示：本软件仅为辅助工具，生成的文档需要人工审核和修改，确保内容准确性和学术规范性。

text

## 使用说明

### 1. 准备工作
1. 获取DeepSeek API密钥：访问 [DeepSeek平台](https://platform.deepseek.com/)
2. 克隆或下载本项目代码

### 2. 安装运行
```bash
# 安装依赖
pip install -r requirements.txt

# 运行程序
python src/main.py
3. 首次使用
在软件界面中输入您的DeepSeek API密钥

选择文档类型（期刊论文、计划等）

输入文档标题和附加要求

点击"生成大纲"开始

4. 多平台打包
GitHub Actions会自动为三个平台构建可执行文件：

Windows: .exe 文件

macOS: .app 文件

Linux: 可执行文件

功能特点
智能大纲生成：根据标题和类型自动生成详细大纲

灵活编辑：可随时修改大纲结构

完整文档生成：基于大纲生成完整的学术文档

多格式导出：支持Word、PDF、Markdown等格式

多平台支持：Windows、macOS、Linux全平台兼容

自定义模板：支持用户自定义文档模板

这个完整的项目包含了所有必要的文件，您可以直接使用。记得在使用前配置好您的DeepSeek API密钥。
