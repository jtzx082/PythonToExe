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
APP_VERSION = "v1.0.5 (Stable Loop)"
AUTHOR_INFO = "开发者：Python开发者\n基于 GB/T 9704-2012 标准"

DEFAULT_CONFIG = {
    "margins": {"top": 3.7, "bottom": 3.5, "left": 2.8, "right": 2.6},
    "line_spacing": 28,
    "fonts": {
        "title": "方正小标宋简体",
        "h1": "黑体",
        "h2": "楷体_GB2312",
        "h3": "仿宋_GB2312",
        "body": "仿宋_GB2312"
    },
    "sizes": {
        "title": 22,
        "h1": 16,
        "h2": 16,
        "h3": 16,
        "body": 16
    }
}

class GongWenFormatterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} {APP_VERSION}")
        self.geometry("900x700")
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.config = self.load_config()
        self.file_list = []
        self.processed_docs = [] 
        self.process_queue = [] # 任务队列
        self.is_processing = False

        self.setup_ui()

    def load_config(self):
        if os.path.exists("config.json"):
            try:
                with open("config.json", "r", encoding="utf-8") as f:
                    return json.load(f)
            except:
                pass
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

        self.sidebar = ctk.CTkFrame(self, width=140, corner_radius=0)
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
        
        ctk.CTkButton(btn_box, text="📂 1. 上传文档", command=self.upload_files, width=180).pack(side="left", padx=10)
        self.btn_process = ctk.CTkButton(btn_box, text="▶ 2. 开始排版", command=self.start_processing, width=180, fg_color="green", state="disabled")
        self.btn_process.pack(side="left", padx=10)
        self.btn_export = ctk.CTkButton(btn_box, text="💾 3. 导出结果", command=self.export_files, width=180, state="disabled")
        self.btn_export.pack(side="left", padx=10)

        self.log_box = ctk.CTkTextbox(f)
        self.log_box.grid(row=1, column=0, sticky="nsew", pady=10)
        self.log_box.insert("0.0", ">>> 欢迎使用！请先上传 Word 文档。\n")
        self.log_box.configure(state="disabled")

        self.progressbar = ctk.CTkProgressBar(f)
        self.progressbar.grid(row=2, column=0, sticky="ew", pady=10)
        self.progressbar.set(0)

    def create_settings_frame(self):
        f = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.frames["settings"] = f
        ctk.CTkLabel(f, text="排版参数设置", font=("Arial", 20)).pack(pady=20)
        
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
        
        info = f"{APP_NAME}\n{APP_VERSION}\n{AUTHOR_INFO}\n\n【常见问题】\n1. 请确保文档未被占用。\n2. 导出时会自动添加“_排版后”后缀。\n3. 如果排版失败，请检查文档是否加密。"
        lbl = ctk.CTkTextbox(f, font=("Arial", 14), wrap="word")
        lbl.insert("0.0", info)
        lbl.configure(state="disabled")
        lbl.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)

    def show_frame(self, name):
        for frame in self.frames.values(): frame.grid_forget()
        self.frames[name].grid(row=0, column=0, sticky="nsew")

    def log(self, text):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"{text}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        # 强制刷新UI，确保卡死前能看到最后一句话
        self.log_box.update_idletasks()

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

    # --- 核心修改：移除 Thread，改用递归调用 ---
    def start_processing(self):
        self.btn_process.configure(state="disabled")
        self.btn_upload.configure(state="disabled")
        self.processed_docs = []
        
        # 准备任务队列
        self.process_queue = list(enumerate(self.file_list))
        self.total_files = len(self.file_list)
        self.success_count = 0
        
        self.log(">>> 排版引擎已启动 (v1.0.5)")
        
        # 延迟100ms开始处理第一个，给UI刷新的机会
        self.after(100, self.process_next_file)

    def process_next_file(self):
        if not self.process_queue:
            # 队列为空，全部完成
            self.on_process_finish(self.success_count)
            return

        # 取出下一个任务
        index, file_path = self.process_queue.pop(0)
        filename = os.path.basename(file_path)
        
        self.progressbar.set(index / self.total_files)
        self.log(f"正在分析: {filename} ...")
        
        # 使用 update() 强制刷新界面，防止假死
        self.update() 

        try:
            # 执行排版 (这是耗时操作)
            doc = self.format_document(file_path)
            self.processed_docs.append((file_path, doc))
            self.success_count += 1
            self.log(f"✅ {filename} 排版成功")
        except Exception as e:
            error_msg = str(e)
            print(traceback.format_exc()) # 控制台打印详细错误
            self.log(f"❌ {filename} 失败: {error_msg}")
        
        # 调度下一个任务，间隔 50ms，保证 UI 响应
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
                self.log(f"已导出: {save_path}")
                count += 1
            except Exception as e:
                self.log(f"导出失败 {base_name}: {e}")

        messagebox.showinfo("导出完成", f"成功导出 {count} 个文件。\n路径: {save_dir}")
        if os.name == 'nt':
            try: os.startfile(save_dir)
            except: pass

    # --- 核心排版逻辑 (增强健壮性) ---
    def format_document(self, file_path):
        if not os.path.exists(file_path):
            raise FileNotFoundError("找不到文件")

        try:
            doc = Document(file_path)
        except Exception as e:
            raise ValueError(f"文档损坏或格式不支持: {e}")

        cfg = self.config

        # 1. 页面设置
        for section in doc.sections:
            try:
                section.top_margin = Cm(cfg["margins"]["top"])
                section.bottom_margin = Cm(cfg["margins"]["bottom"])
                section.left_margin = Cm(cfg["margins"]["left"])
                section.right_margin = Cm(cfg["margins"]["right"])
                section.page_width = Cm(21)
                section.page_height = Cm(29.7)
            except Exception as e:
                print(f"页面设置警告: {e}")

        # 2. 字体设置
        try:
            style = doc.styles['Normal']
            style.font.name = 'Times New Roman'
            style.font.size = Pt(cfg["sizes"]["body"])
            style._element.rPr.rFonts.set(qn('w:eastAsia'), cfg["fonts"]["body"])
        except Exception: pass

        # 3. 遍历排版
        for paragraph in doc.paragraphs:
            text = paragraph.text.strip()
            if not text: continue

            try:
                paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                paragraph.paragraph_format.line_spacing = Pt(cfg["line_spacing"])
            except: pass
            
            # 标题识别
            if paragraph == doc.paragraphs[0] and len(text) < 50:
                self.set_font(paragraph, cfg["fonts"]["title"], cfg["sizes"]["title"], bold=False)
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                try: paragraph.paragraph_format.space_after = Pt(cfg["line_spacing"])
                except: pass
                continue

            if re.match(r"^[一二三四五六七八九十]+、", text):
                self.set_font(paragraph, cfg["fonts"]["h1"], cfg["sizes"]["h1"], bold=False)
                try: paragraph.paragraph_format.first_line_indent = Pt(cfg["sizes"]["h1"] * 2)
                except: pass
                continue

            if re.match(r"^（[一二三四五六七八九十]+）", text):
                self.set_font(paragraph, cfg["fonts"]["h2"], cfg["sizes"]["h2"], bold=False)
                try: paragraph.paragraph_format.first_line_indent = Pt(cfg["sizes"]["h2"] * 2)
                except: pass
                continue

            if re.match(r"^\d+\.", text):
                self.set_font(paragraph, cfg["fonts"]["h3"], cfg["sizes"]["h3"], bold=True)
                try: paragraph.paragraph_format.first_line_indent = Pt(cfg["sizes"]["h3"] * 2)
                except: pass
                continue

            self.set_font(paragraph, cfg["fonts"]["body"], cfg["sizes"]["body"])
            try:
                paragraph.paragraph_format.first_line_indent = Pt(cfg["sizes"]["body"] * 2)
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
            except: pass

        # 4. 表格
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        self.set_font(p, "仿宋_GB2312", 14) 

        # 5. 页码
        try:
            footer = doc.sections[0].footer
            p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
            self.add_page_number(p)
        except: pass

        return doc

    def set_font(self, paragraph, font_name, font_size, bold=False):
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = Pt(font_size)
            run.bold = bold
            try:
                run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
            except: pass

    def add_page_number(self, paragraph):
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = paragraph.add_run()
        fldChar1 = OxmlElement('w:fldChar'); fldChar1.set(qn('w:fldCharType'), 'begin')
        instrText = OxmlElement('w:instrText'); instrText.set(qn('xml:space'), 'preserve'); instrText.text = "PAGE"
        fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'end')
        run._r.append(fldChar1); run._r.append(instrText); run._r.append(fldChar2)
        run.font.name = "宋体"; run.font.size = Pt(14)

if __name__ == "__main__":
    app = GongWenFormatterApp()
    app.mainloop()
