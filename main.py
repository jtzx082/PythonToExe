"""
AI 写作助手 - 智能文稿创作平台
支持 Anthropic Claude、DeepSeek、OpenAI 及自定义兼容接口
支持学术论文、研究报告、工作计划、反思总结、案例分析、工作总结及自定义文稿
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import json
import os
import re
from datetime import datetime

# ── 引入 docx 相关库用于公文排版 ──────────────────────────────────────────────
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ── Markdown 转纯文本工具 ────────────────────────────────────────────────────
def md_to_plain(text: str) -> str:
    """将 Markdown 文本转换为干净的纯文本"""
    text = re.sub(r"```[\s\S]*?```", lambda m: m.group().replace("```", "").strip(), text)
    text = re.sub(r"`([^`]+)`", r"\1", text)
    text = re.sub(r"^#{1,6}\s+(.+)$", r"\1", text, flags=re.MULTILINE)
    text = re.sub(r"\*{1,3}([^*]+)\*{1,3}", r"\1", text)
    text = re.sub(r"_{1,3}([^_]+)_{1,3}", r"\1", text)
    text = re.sub(r"\[([^\]]+)\]\([^)]*\)", r"\1", text)
    text = re.sub(r"!\[([^\]]*)\]\([^)]*\)", r"\1", text)
    text = re.sub(r"^>+\s?", "", text, flags=re.MULTILINE)
    text = re.sub(r"^\s*[-*+]\s+", "", text, flags=re.MULTILINE)
    text = re.sub(r"^\s*\d+\.\s+", "", text, flags=re.MULTILINE)
    text = re.sub(r"^[-*_]{3,}\s*$", "", text, flags=re.MULTILINE)
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


# ── 公文格式化保存核心逻辑 ────────────────────────────────────────────────────
def save_as_docx(filepath: str, title: str, md_text: str):
    """
    将 Markdown 转换为符合《党政机关公文格式》标准的 Word 文档
    规范参考：GB/T 9704-2012
    """
    
    doc = Document()

    # ── 1. 页面设置 (Page Setup) ──
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.top_margin = Mm(37)
    section.bottom_margin = Mm(35)
    section.left_margin = Mm(28)
    section.right_margin = Mm(26)

    # 开启奇偶页页眉页脚不同
    doc.settings.odd_and_even_pages_header_footer = True

    # ── 2. 基础样式定义 (Styles) ──
    def set_run_font(run, font_cn, font_en='Times New Roman', size_pt=16, bold=False):
        run.font.name = font_en
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_cn)
        run.font.size = Pt(size_pt)
        run.font.bold = bold
        run.font.color.rgb = RGBColor(0, 0, 0)

    # 修改默认样式 'Normal' 为公文正文样式
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Times New Roman'
    style_normal.element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
    style_normal.font.size = Pt(16)
    style_normal.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    style_normal.paragraph_format.line_spacing = Pt(28)
    style_normal.paragraph_format.first_line_indent = Pt(32)

    # ── 3. 标题排版 (Main Title) ──
    head_p = doc.add_paragraph()
    head_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    head_p.paragraph_format.first_line_indent = Pt(0)
    head_p.paragraph_format.line_spacing = Pt(28)
    head_p.paragraph_format.space_before = Pt(0)
    head_p.paragraph_format.space_after = Pt(28) 

    run_title = head_p.add_run(title)
    set_run_font(run_title, '方正小标宋简体', size_pt=22, bold=False)

    # ── 4. 正文内容解析与转换 ──
    lines = md_text.splitlines()
    for line in lines:
        stripped = line.rstrip()
        
        if re.match(r"^[-*_]{3,}\s*$", stripped):
            continue

        # 识别标题 (#)
        heading_match = re.match(r"^(#{1,6})\s+(.*)", stripped)
        if heading_match:
            level = len(heading_match.group(1))
            text = _strip_inline(heading_match.group(2))
            
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = Pt(28)
            p.paragraph_format.first_line_indent = Pt(32)

            run = p.add_run(text)
            
            if level == 1:
                set_run_font(run, 'SimHei', size_pt=16) 
            elif level == 2:
                set_run_font(run, 'KaiTi', size_pt=16)
            else:
                set_run_font(run, '仿宋', size_pt=16, bold=True)
            continue
            
        if not stripped:
            continue

        # 普通段落 (正文)
        p = doc.add_paragraph()
        _add_inline_runs_styled(p, stripped)

    # ── 5. 页码设置 (Page Numbers) ──
    def create_page_number_xml(run):
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')
        run._element.append(fldChar1)

        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')
        instrText.text = "PAGE"
        run._element.append(instrText)

        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')
        run._element.append(fldChar2)

    def setup_footer(footer, alignment):
        p = footer.paragraphs[0]
        p.alignment = alignment
        p.paragraph_format.first_line_indent = 0
        r1 = p.add_run("— ") 
        set_run_font(r1, 'SimSun', size_pt=14)
        r2 = p.add_run()
        set_run_font(r2, 'SimSun', size_pt=14)
        create_page_number_xml(r2)
        r3 = p.add_run(" —")
        set_run_font(r3, 'SimSun', size_pt=14)

    setup_footer(section.footer, WD_ALIGN_PARAGRAPH.RIGHT)
    setup_footer(section.even_page_footer, WD_ALIGN_PARAGRAPH.LEFT)

    doc.save(filepath)


def _strip_inline(text: str) -> str:
    """去掉行内 Markdown 符号，只保留文字"""
    text = re.sub(r"\*{1,3}([^*]+)\*{1,3}", r"\1", text)
    text = re.sub(r"_{1,3}([^_]+)_{1,3}", r"\1", text)
    text = re.sub(r"`([^`]+)`", r"\1", text)
    text = re.sub(r"\[([^\]]+)\]\([^)]*\)", r"\1", text)
    return text


def _add_inline_runs_styled(paragraph, text: str):
    """解析 Markdown 行内格式并应用到 Docx Run"""
    from docx.oxml.ns import qn
    from docx.shared import Pt, RGBColor
    
    pattern = re.compile(r"(\*{1,3}[^*]+\*{1,3}|_{1,3}[^_]+_{1,3}|`[^`]+`)")
    last = 0
    
    def apply_style(run, bold=False, italic=False, code=False):
        run.font.name = 'Times New Roman'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
        run.font.size = Pt(16)
        run.font.color.rgb = RGBColor(0,0,0)
        
        if bold: run.font.bold = True
        if italic: run.font.italic = True
        if code:
             run.font.name = 'Courier New'

    for m in pattern.finditer(text):
        if m.start() > last:
            r = paragraph.add_run(text[last:m.start()])
            apply_style(r)
            
        token = m.group()
        if token.startswith("***") or token.startswith("___"):
            r = paragraph.add_run(token[3:-3])
            apply_style(r, bold=True, italic=True)
        elif token.startswith("**") or token.startswith("__"):
            r = paragraph.add_run(token[2:-2])
            apply_style(r, bold=True)
        elif token.startswith("*") or token.startswith("_"):
            r = paragraph.add_run(token[1:-1])
            apply_style(r, italic=True)
        elif token.startswith("`"):
            r = paragraph.add_run(token[1:-1])
            apply_style(r, code=True)
        last = m.end()
        
    if last < len(text):
        r = paragraph.add_run(text[last:])
        apply_style(r)


# ── 主题配置 ────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ── 常量定义 ────────────────────────────────────────────────────────────────
CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".ai_writer_config.json")
APP_VERSION = "v2.2.1"  # Updated version
APP_AUTHOR  = "Yu JinQuan"

# ── 服务商配置表 ────────────────────────────────────────────────────────────
PROVIDERS = {
    "Anthropic (Claude)": {
        "icon":     "🤖",
        "type":     "anthropic",
        "base_url": "[https://api.anthropic.com](https://api.anthropic.com)", # Default, can be overridden
        "key_hint": "sk-ant-api03-...",
        "models": [
            "claude-3-5-sonnet-20241022",
            "claude-3-opus-20240229",
            "claude-3-haiku-20240307",
        ],
        "default_model": "claude-3-5-sonnet-20241022",
    },
    "DeepSeek": {
        "icon":     "🐋",
        "type":     "openai_compat",
        "base_url": "[https://api.deepseek.com](https://api.deepseek.com)",
        "key_hint": "sk-...",
        "models": [
            "deepseek-chat",
            "deepseek-reasoner",
        ],
        "default_model": "deepseek-chat",
    },
    "OpenAI": {
        "icon":     "🌐",
        "type":     "openai_compat",
        "base_url": "[https://api.openai.com/v1](https://api.openai.com/v1)",
        "key_hint": "sk-...",
        "models": [
            "gpt-4o",
            "gpt-4o-mini",
            "o1-preview",
            "o1-mini",
        ],
        "default_model": "gpt-4o",
    },
    "自定义 (OpenAI 兼容)": {
        "icon":     "🔧",
        "type":     "openai_compat",
        "base_url": "",
        "key_hint": "API Key...",
        "models": [],
        "default_model": "",
    },
}

PROVIDER_NAMES = list(PROVIDERS.keys())

# ── 文稿类型 ────────────────────────────────────────────────────────────────
DOCUMENT_TYPES = [
    ("📄", "学术论文",  "含摘要、引言、方法、结果、讨论、参考文献"),
    ("📊", "研究报告",  "含背景、分析框架、结论与建议"),
    ("📋", "工作计划",  "含目标、阶段步骤、时间线、资源安排"),
    ("🔍", "反思总结",  "含经历回顾、收获、不足与改进方向"),
    ("🔬", "案例分析",  "含案例背景、问题呈现、深度分析、启示"),
    ("📝", "工作总结",  "含工作概述、核心成果、问题与展望"),
    ("✨", "自定义",    "根据您的描述自由定制文稿类型与结构"),
]

OUTLINE_SYSTEM = """你是一位资深写作顾问，擅长为各类专业文稿设计清晰、合理的结构大纲。

请根据用户提供的文稿类型、题目和要求，输出一份层次分明的大纲。

格式规范：
- 一级章节：1. 章节名称（简要说明本章核心内容
