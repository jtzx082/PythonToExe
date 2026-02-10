import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import json
import re
import time
import traceback
from docx import Document
from docx.shared import Cm, Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.oxml import OxmlElement

# --- 全局配置 ---
APP_NAME = "公文自动排版助手"
APP_VERSION = "v2.0.0 (Smart Structure AI)"
AUTHOR_INFO = "开发者：Python开发者\n基于 GB/T 9704-2012 标准"

DEFAULT_CONFIG = {
    "margins": {"top": 3.7, "bottom": 3.5, "left": 2.8, "right": 2.6},
    "line_spacing": 28,  # 固定值 28磅
    "fonts": {
        "title": "方正小标宋简体",
        "subtitle": "楷体_GB2312",
        "author": "楷体_GB2312",
        "abstract": "楷体_GB2312",
        "h1": "黑体",
        "h2": "楷体_GB2312",
        "h3": "仿宋_GB2312",
        "body": "仿宋_GB2312"
    },
    "sizes": {
        "title": 22,    # 二号
        "subtitle": 16, # 三号
        "author": 14,   # 四号
        "abstract": 14, # 四号
        "h1": 16,       # 三号
        "h2": 16,
        "h3": 16,
        "body": 16
    }
}

class GongWenFormatterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} {APP_VERSION}")
        self.geometry("1000x750")
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.config = self.load_config()
        self.file_list = []
        self.processed_docs = [] 
        self.process_queue = []
        
        self.setup_ui()

    def load_config(self):
        if os.path.exists("config.json"):
            try:
                with open("config.json", "r", encoding="utf-8") as f:
                    return json.load(f)
            except: pass
        return DEFAULT_CONFIG

    def save_config(self):
        try:
            with open("config.json", "w", encoding="utf-8") as f:
                json.dump(self.config, f, ensure_ascii=False, indent=4)
            messagebox.showinfo("成功", "配置已保存！")
        except Exception as e:
            messagebox.showerror("错误", f"保存配置失败: {e}")

    def setup_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.sidebar = ctk.CTkFrame(self, width=180, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        ctk.CTkLabel(self.sidebar, text=APP_NAME, font=ctk.CTkFont(size=18, weight="bold")).pack(pady=20)
        
        self.btn_home = ctk.CTkButton(self.sidebar, text="排版工作台", command=lambda: self.show_frame("home"))
        self.btn_home.pack(pady=10, padx=10)
        self.btn_settings = ctk.CTkButton(self.sidebar, text="参数设置", command=lambda: self.show_frame("settings"))
        self.btn_settings.pack(pady=10, padx=10)
        self.btn_about = ctk.CTkButton(self.sidebar, text="使用说明", command=lambda: self.show_frame("about"))
        self.btn_about.pack(pady=10, padx=10)

        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1)

        self.frames = {}
        self.create_home_frame()
        self.create_settings_frame()
        self.create_about_frame()
        self.show_frame("home")

    def create_home_frame(self):
        f = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.frames["home"] = f
        f.grid_columnconfigure(0, weight=1)
        f.grid_rowconfigure(1, weight=1)
        
        btn_box = ctk.CTkFrame(f, fg_color="transparent")
        btn_box.grid(row=0, column=0, sticky="ew", pady=10)
        
        self.btn_upload = ctk.CTkButton(btn_box, text="📂 1. 上传文档", command=self.upload_files, width=180)
        self.btn_upload.pack(side="left", padx=10)
        
        self.btn_process = ctk.CTkButton(btn_box, text="▶ 2. 开始排版", command=self.start_processing, width=180, fg_color="green", state="disabled")
        self.btn_process.pack(side="left", padx=10)
        
        self.btn_export = ctk.CTkButton(btn_box, text="💾 3. 导出结果", command=self.export_files, width=180, state="disabled")
        self.btn_export.pack(side="left", padx=10)

        self.log_box = ctk.CTkTextbox(f)
        self.log_box.grid(row=1, column=0, sticky="nsew", pady=10)
        self.log_box.insert("0.0", ">>> 欢迎使用 v2.0.0 智能版！请上传文档。\n")
        self.log_box.configure(state="disabled")

        self.progressbar = ctk.CTkProgressBar(f)
        self.progressbar.grid(row=2, column=0, sticky="ew", pady=10)
        self.progressbar.set(0)

    def create_settings_frame(self):
        f = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.frames["settings"] = f
        ctk.CTkLabel(f, text="排版参数设置 (重启软件生效)", font=("Arial", 20)).pack(pady=20)
        
        self.entries = {}
        settings = [
            ("上边距 (cm)", "top", self.config["margins"]["top"]),
            ("下边距 (cm)", "bottom", self.config["margins"]["bottom"]),
            ("左边距 (cm)", "left", self.config["margins"]["left"]),
            ("右边距 (cm)", "right", self.config["margins"]["right"]),
            ("行间距 (磅)", "line_spacing", self.config["line_spacing"])
        ]

        for label_text, key, val in settings:
            row = ctk.CTkFrame(f, fg_color="transparent")
            row.pack(fill="x", pady=5)
            ctk.CTkLabel(row, text=label_text, width=120).pack(side="left")
            entry = ctk.CTkEntry(row)
            entry.insert(0, str(val))
            entry.pack(side="left", fill="x", expand=True)
            self.entries[key] = entry

        ctk.CTkButton(f, text="保存设置", command=self.update_config).pack(pady=20)

    def create_about_frame(self):
        f = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.frames["about"] = f
        f.grid_columnconfigure(0, weight=1)
        f.grid_rowconfigure(0, weight=1)
        
        info = f"""{APP_NAME}\n{APP_VERSION}\n{AUTHOR_INFO}\n
【v2.0.0 核心升级】
1. 智能结构分析：准确识别 标题/副标题/作者/摘要/正文。
2. 告别误判：不再把“尊敬的领导”误认为大标题。
3. 纯净排版：自动清理原有格式，应用国标样式。

【排版规范】
- 大标题：方正小标宋简体，二号，居中。
- 正文：仿宋_GB2312，三号，首行缩进2字符。
- 一级标题：黑体，三号。
- 页面：上3.7 下3.5 左2.8 右2.6 (cm)。
"""
        lbl = ctk.CTkTextbox(f, font=("Arial", 14), wrap="word", width=600, height=500)
        lbl.insert("0.0", info)
        lbl.configure(state="disabled")
        lbl.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)

    def show_frame(self, name):
        for frame in self.frames.values(): frame.grid_forget()
        self.frames[name].grid(row=0, column=0, sticky="nsew")

    def log(self, text):
        print(f"[LOG] {text}")
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"{text}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.update_idletasks()

    def update_config(self):
        try:
            self.config["margins"]["top"] = float(self.entries["top"].get())
            self.config["margins"]["bottom"] = float(self.entries["bottom"].get())
            self.config["margins"]["left"] = float(self.entries["left"].get())
            self.config["margins"]["right"] = float(self.entries["right"].get())
            self.config["line_spacing"] = float(self.entries["line_spacing"].get())
            self.save_config()
        except ValueError:
            messagebox.showerror("错误", "输入必须是数字")

    def upload_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Word Document", "*.docx")])
        if files:
            self.file_list = list(files)
            self.processed_docs = [] 
            self.log(f"已加载 {len(files)} 个文件。请点击“开始排版”。")
            self.btn_process.configure(state="normal")
            self.btn_export.configure(state="disabled")

    # --- 流程控制 ---
    def start_processing(self):
        self.log(">>> 正在启动智能分析引擎...")
        self.btn_process.configure(state="disabled")
        self.btn_upload.configure(state="disabled")
        self.processed_docs = []
        self.process_queue = list(enumerate(self.file_list))
        self.total_files = len(self.file_list)
        self.success_count = 0
        self.update()
        self.after(100, self.process_next_file)

    def process_next_file(self):
        if not self.process_queue:
            self.on_process_finish(self.success_count)
            return

        index, file_path = self.process_queue.pop(0)
        filename = os.path.basename(file_path)
        self.progressbar.set(index / self.total_files)
        self.log(f"正在分析: {filename} ...")
        self.update() 

        try:
            doc = self.format_document(file_path)
            self.processed_docs.append((file_path, doc))
            self.success_count += 1
            self.log(f"✅ {filename} 排版成功")
        except Exception as e:
            error_msg = str(e)
            print(f"ERROR: {traceback.format_exc()}")
            self.log(f"❌ {filename} 失败: {error_msg}")
            messagebox.showerror("排版错误", f"文件：{filename}\n错误：{error_msg}")
        
        self.after(50, self.process_next_file)

    def on_process_finish(self, count):
        self.progressbar.set(1.0)
        self.btn_process.configure(state="normal")
        self.btn_upload.configure(state="normal")
        if count > 0:
            self.btn_export.configure(state="normal")
            messagebox.showinfo("完成", f"已完成 {count} 个文档的排版！\n请点击“导出结果”保存文件。")
        else:
            messagebox.showwarning("失败", "没有文档被成功处理。")

    def export_files(self):
        if not self.processed_docs: return
        save_dir = filedialog.askdirectory(title="选择导出文件夹")
        if not save_dir: return
        
        count = 0
        self.log(">>> 开始导出...")
        for original_path, doc in self.processed_docs:
            try:
                base_name = os.path.basename(original_path)
                name, ext = os.path.splitext(base_name)
                new_name = f"{name}_排版后{ext}"
                save_path = os.path.join(save_dir, new_name)
                doc.save(save_path)
                self.log(f"已保存: {new_name}")
                count += 1
            except Exception as e:
                self.log(f"保存失败 {base_name}: {e}")

        messagebox.showinfo("导出完成", f"成功导出 {count} 个文件。\n路径: {save_dir}")
        if os.name == 'nt':
            try: os.startfile(save_dir)
            except: pass

    # =========================================================
    #  核心排版逻辑 v2.0 (智能结构分析)
    # =========================================================
    
    def analyze_paragraph_type(self, text, index, is_body_started):
        """ 
        智能判断段落类型 
        返回: 'TITLE', 'SUBTITLE', 'AUTHOR', 'ABSTRACT', 'KEYWORD', 'H1', 'H2', 'H3', 'BODY'
        """
        text = text.strip()
        
        # 1. 强制正文标记（一旦出现，后续全是正文）
        if re.match(r"^(尊敬的|各位|亲爱的|大家好)", text):
            return 'BODY', True
        
        # 2. 如果已经进入正文区域，直接按正文逻辑处理
        if is_body_started:
            if re.match(r"^[一二三四五六七八九十]+、", text): return 'H1', True
            if re.match(r"^（[一二三四五六七八九十]+）", text): return 'H2', True
            if re.match(r"^\d+\.", text): return 'H3', True
            return 'BODY', True

        # 3. 尚未进入正文，分析头部信息
        
        # 关键词/摘要
        if re.match(r"^(摘要|【摘要】)", text): return 'ABSTRACT', False
        if re.match(r"^(关键词|【关键词】)", text): return 'KEYWORD', False

        # 副标题 (破折号开头，或被括号包围)
        if text.startswith("——") or text.startswith("--") or (text.startswith("（") and text.endswith("）")):
            return 'SUBTITLE', False

        # 主标题 (必须在最前面，且字数较少，且不含句号)
        if index < 3 and len(text) < 40 and ("。" not in text):
            return 'TITLE', False
        
        # 作者/单位信息 (字数极少，且不是标题)
        if len(text) < 25 and ("。" not in text):
            return 'AUTHOR', False

        # 默认兜底为正文（只要出现长难句，就视为正文开始）
        return 'BODY', True

    def format_document(self, file_path):
        if not os.path.exists(file_path): raise FileNotFoundError("文件不存在")
        try: doc = Document(file_path)
        except Exception as e: raise ValueError(f"文档损坏: {e}")

        cfg = self.config

        # 1. 页面设置
        try:
            for section in doc.sections:
                section.top_margin = Cm(cfg["margins"]["top"])
                section.bottom_margin = Cm(cfg["margins"]["bottom"])
                section.left_margin = Cm(cfg["margins"]["left"])
                section.right_margin = Cm(cfg["margins"]["right"])
                section.page_width = Cm(21)
                section.page_height = Cm(29.7)
        except: pass

        # 2. 遍历并排版
        is_body_started = False
        
        for i, paragraph in enumerate(doc.paragraphs):
            text = paragraph.text.strip()
            if not text: continue

            # --- 步骤A: 智能识别类型 ---
            p_type, is_body_now = self.analyze_paragraph_type(text, i, is_body_started)
            if is_body_now: is_body_started = True

            print(f"Debug: Line {i} [{p_type}] : {text[:10]}...") # 终端调试用

            # --- 步骤B: 清除旧格式 (重要！) ---
            try:
                paragraph.paragraph_format.first_line_indent = None
                paragraph.paragraph_format.left_indent = None
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            except: pass

            # --- 步骤C: 应用新样式 ---
            try:
                # 设定固定行距 (所有段落通用)
                paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                paragraph.paragraph_format.line_spacing = Pt(cfg["line_spacing"])
                
                # 网格对齐
                self.set_paragraph_grid_props(paragraph)

                if p_type == 'TITLE':
                    self.safe_set_font(paragraph, cfg["fonts"]["title"], cfg["sizes"]["title"], bold=False)
                    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    paragraph.paragraph_format.space_after = Pt(cfg["line_spacing"]) # 标题下空一行
                    self.set_indent_xml(paragraph, 0)

                elif p_type == 'SUBTITLE':
                    self.safe_set_font(paragraph, cfg["fonts"]["subtitle"], cfg["sizes"]["subtitle"], bold=False)
                    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    self.set_indent_xml(paragraph, 0)

                elif p_type == 'AUTHOR':
                    self.safe_set_font(paragraph, cfg["fonts"]["author"], cfg["sizes"]["author"], bold=False)
                    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    self.set_indent_xml(paragraph, 0)

                elif p_type == 'ABSTRACT' or p_type == 'KEYWORD':
                    self.safe_set_font(paragraph, cfg["fonts"]["abstract"], cfg["sizes"]["abstract"], bold=False)
                    self.set_indent_xml(paragraph, 2) # 摘要缩进

                elif p_type == 'H1':
                    self.safe_set_font(paragraph, cfg["fonts"]["h1"], cfg["sizes"]["h1"], bold=False) # 黑体
                    self.set_indent_xml(paragraph, 2)

                elif p_type == 'H2':
                    self.safe_set_font(paragraph, cfg["fonts"]["h2"], cfg["sizes"]["h2"], bold=False) # 楷体
                    self.set_indent_xml(paragraph, 2)

                elif p_type == 'H3':
                    self.safe_set_font(paragraph, cfg["fonts"]["h3"], cfg["sizes"]["h3"], bold=True) # 仿宋加粗
                    self.set_indent_xml(paragraph, 2)

                else: # BODY
                    self.safe_set_font(paragraph, cfg["fonts"]["body"], cfg["sizes"]["body"], bold=False)
                    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                    self.set_indent_xml(paragraph, 2) # 核心：正文首行缩进2字符

            except Exception as e:
                print(f"Warning: 段落排版出错 {e}")

        # 3. 表格处理
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        self.safe_set_font(p, "仿宋_GB2312", 14)
                        self.set_paragraph_grid_props(p)

        # 4. 页码
        try:
            footer = doc.sections[0].footer
            p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
            self.add_page_number(p)
        except: pass

        return doc

    # --- XML 辅助函数 ---
    def set_indent_xml(self, paragraph, chars=2):
        try:
            pPr = paragraph._p.get_or_add_pPr()
            ind = pPr.get_or_add_ind()
            if chars == 0:
                if 'w:firstLineChars' in ind.attrib: del ind.attrib['w:firstLineChars']
            else:
                ind.set(qn('w:firstLineChars'), str(int(chars * 100)))
                if 'w:firstLine' in ind.attrib: del ind.attrib['w:firstLine'] # 清除冲突
        except: pass

    def set_paragraph_grid_props(self, paragraph):
        try:
            pPr = paragraph._p.get_or_add_pPr()
            snap = pPr.find(qn('w:snapToGrid'))
            if snap is None:
                snap = OxmlElement('w:snapToGrid')
                pPr.append(snap)
            snap.set(qn('w:val'), '1')
            
            adj = pPr.find(qn('w:adjustRightInd'))
            if adj is None:
                adj = OxmlElement('w:adjustRightInd')
                pPr.append(adj)
            adj.set(qn('w:val'), '1')
        except: pass

    def safe_set_font(self, paragraph, font_name, font_size, bold=False):
        try:
            for run in paragraph.runs:
                run.font.name = font_name
                run.font.size = Pt(font_size)
                run.bold = bold
                run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
        except: pass

    def add_page_number(self, paragraph):
        try:
            paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run = paragraph.add_run()
            fldChar1 = OxmlElement('w:fldChar'); fldChar1.set(qn('w:fldCharType'), 'begin')
            instrText = OxmlElement('w:instrText'); instrText.set(qn('xml:space'), 'preserve'); instrText.text = "PAGE"
            fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'end')
            run._r.append(fldChar1); run._r.append(instrText); run._r.append(fldChar2)
            run.font.name = "宋体"; run.font.size = Pt(14)
        except: pass

if __name__ == "__main__":
    app = GongWenFormatterApp()
    app.mainloop()
