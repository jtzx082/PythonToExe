import customtkinter as ctk
import threading
from openai import OpenAI
import os
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from tkinter import filedialog
import json
import time
import re  # 引入正则库，用于清洗Markdown符号

# --- 配置区域 ---
APP_VERSION = "v5.0.0 (Clean Text Final)"
DEV_NAME = "俞晋全"
DEV_ORG = "俞晋全高中化学名师工作室"
# ----------------

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class PaperWriterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"期刊论文撰写系统 (纯净版) - {DEV_NAME}")
        self.geometry("1100x850")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.api_config = {
            "api_key": "",
            "base_url": "https://api.deepseek.com", 
            "model": "deepseek-chat"
        }
        self.load_config()

        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
        self.tab_info = self.tabview.add("1. 论文参数")
        self.tab_write = self.tabview.add("2. 深度撰写")
        self.tab_settings = self.tabview.add("3. 系统设置")

        self.setup_info_tab()
        self.setup_write_tab()
        self.setup_settings_tab()

        self.status_label = ctk.CTkLabel(self, text="就绪", text_color="gray")
        self.status_label.grid(row=1, column=0, pady=5)
        
        self.progressbar = ctk.CTkProgressBar(self, mode="determinate")
        self.progressbar.grid(row=2, column=0, padx=20, pady=(0, 10), sticky="ew")
        self.progressbar.set(0)

    # === Tab 1: 信息设定 ===
    def setup_info_tab(self):
        t = self.tab_info
        t.grid_columnconfigure(1, weight=1)

        # 题目
        ctk.CTkLabel(t, text="论文题目:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=0, column=0, padx=10, pady=10, sticky="e")
        self.entry_title = ctk.CTkEntry(t, placeholder_text="例如：高中化学虚拟仿真实验教学的价值与策略研究", height=35)
        self.entry_title.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        # 作者信息
        ctk.CTkLabel(t, text="作者姓名:", font=("Microsoft YaHei UI", 12)).grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.entry_author = ctk.CTkEntry(t, placeholder_text="俞晋全")
        self.entry_author.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        ctk.CTkLabel(t, text="单位信息:", font=("Microsoft YaHei UI", 12)).grid(row=2, column=0, padx=10, pady=5, sticky="e")
        self.entry_org = ctk.CTkEntry(t, placeholder_text="甘肃省金塔县中学, 甘肃金塔 735399")
        self.entry_org.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        # 字数控制
        ctk.CTkLabel(t, text="期望字数:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=3, column=0, padx=10, pady=5, sticky="e")
        self.entry_word_count = ctk.CTkEntry(t, placeholder_text="5000")
        self.entry_word_count.insert(0, "4500") 
        self.entry_word_count.grid(row=3, column=1, padx=10, pady=5, sticky="ew")

        # 大纲预览
        ctk.CTkLabel(t, text="结构大纲 (自动生成/手动修改):", font=("Microsoft YaHei UI", 12, "bold")).grid(row=4, column=0, padx=10, pady=(10,0), sticky="nw")
        self.txt_outline = ctk.CTkTextbox(t, height=250, font=("Microsoft YaHei UI", 13))
        self.txt_outline.grid(row=4, column=1, padx=10, pady=10, sticky="nsew")
        
        self.btn_gen_outline = ctk.CTkButton(t, text="生成标准期刊大纲", command=self.run_gen_outline, fg_color="#1F6AA5")
        self.btn_gen_outline.grid(row=5, column=1, pady=10, sticky="e")

    # === Tab 2: 深度撰写 ===
    def setup_write_tab(self):
        t = self.tab_write
        t.grid_columnconfigure(0, weight=1)
        t.grid_rowconfigure(1, weight=1)

        info = "提示：导出时会自动清除 Markdown 格式 (*, #)，生成纯净的 Word 文档，方便您直接排版。"
        ctk.CTkLabel(t, text=info, text_color="gray").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        
        self.txt_paper = ctk.CTkTextbox(t, font=("Microsoft YaHei UI", 14))
        self.txt_paper.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)

        btn_frame = ctk.CTkFrame(t, fg_color="transparent")
        btn_frame.grid(row=2, column=0, pady=10)
        
        self.btn_gen_paper = ctk.CTkButton(btn_frame, text="开始深度撰写", command=self.run_deep_write, 
                                           width=200, height=40, font=("Microsoft YaHei UI", 14, "bold"))
        self.btn_gen_paper.pack(side="left", padx=20)
        
        self.btn_save_word = ctk.CTkButton(btn_frame, text="导出纯净 Word 文档", command=self.save_to_word,
                                           fg_color="#2CC985", width=200, height=40)
        self.btn_save_word.pack(side="left", padx=20)

    # === Tab 3: 设置 ===
    def setup_settings_tab(self):
        t = self.tab_settings
        ctk.CTkLabel(t, text="API Key:").pack(pady=(20, 5))
        self.entry_key = ctk.CTkEntry(t, width=400, show="*")
        self.entry_key.insert(0, self.api_config.get("api_key", ""))
        self.entry_key.pack(pady=5)
        ctk.CTkLabel(t, text="Base URL:").pack(pady=5)
        self.entry_url = ctk.CTkEntry(t, width=400)
        self.entry_url.insert(0, self.api_config.get("base_url", ""))
        self.entry_url.pack(pady=5)
        ctk.CTkLabel(t, text="Model:").pack(pady=5)
        self.entry_model = ctk.CTkEntry(t, width=400)
        self.entry_model.insert(0, self.api_config.get("model", ""))
        self.entry_model.pack(pady=5)
        ctk.CTkButton(t, text="保存配置", command=self.save_config).pack(pady=20)

    # --- 逻辑核心 ---

    def get_client(self):
        key = self.api_config.get("api_key")
        base = self.api_config.get("base_url")
        if not key:
            self.status_label.configure(text="错误：请配置 API Key", text_color="red")
            return None
        return OpenAI(api_key=key, base_url=base)

    def run_gen_outline(self):
        title = self.entry_title.get()
        if not title: return
        threading.Thread(target=self.thread_gen_outline, args=(title,), daemon=True).start()

    def thread_gen_outline(self, title):
        client = self.get_client()
        if not client: return
        self.status_label.configure(text="正在构建大纲...", text_color="#1F6AA5")
        
        prompt = f"""
        请为高中化学教学论文《{title}》设计一份大纲。
        【格式要求】：
        1. 必须包含：摘要、关键词、一、引言；二、理论/价值；三、策略/实践；四、结语；参考文献。
        2. 正文标题使用汉字数字（一、二...）。
        3. 请直接输出大纲文本，不要使用 Markdown 格式（不要用 ** 或 #）。
        """
        try:
            response = client.chat.completions.create(
                model=self.api_config.get("model"),
                messages=[{"role": "user", "content": prompt}],
                stream=True
            )
            self.txt_outline.delete("0.0", "end")
            for chunk in response:
                if chunk.choices[0].delta.content:
                    self.txt_outline.insert("end", chunk.choices[0].delta.content)
            self.status_label.configure(text="大纲已生成", text_color="green")
        except Exception as e:
            self.status_label.configure(text=f"API 错误: {str(e)}", text_color="red")

    def run_deep_write(self):
        title = self.entry_title.get()
        outline = self.txt_outline.get("0.0", "end").strip()
        try: total_words = int(self.entry_word_count.get().strip())
        except: total_words = 4500
        
        if len(outline) < 10: return
        threading.Thread(target=self.thread_deep_write, args=(title, outline, total_words), daemon=True).start()

    def thread_deep_write(self, title, outline, total_words):
        client = self.get_client()
        if not client: return

        self.btn_gen_paper.configure(state="disabled", text="正在撰写纯净文本...")
        self.txt_paper.delete("0.0", "end")
        self.progressbar.set(0)

        # 动态分配字数
        w_intro = int(total_words * 0.15)
        w_theory = int(total_words * 0.2)
        w_practice = int(total_words * 0.55) # 重点
        w_concl = int(total_words * 0.1)

        sections = [
            ("摘要与关键词", f"请撰写【摘要】（300字）和【关键词】。纯文本格式，不要加粗符号。"),
            ("一、引言与背景", f"撰写论文第一部分。背景与意义。字数约 {w_intro} 字。"),
            ("二、核心价值", f"撰写论文理论价值部分。结合核心素养。字数约 {w_theory} 字。"),
            ("三、教学策略与实践（重点）", f"撰写核心实践部分。必须包含具体的化学实验案例、师生互动细节。字数约 {w_practice} 字。"),
            ("四、结语与参考文献", f"撰写结语和参考文献（5-8条）。")
        ]

        full_text = ""
        total = len(sections)

        try:
            for i, (name, instruction) in enumerate(sections):
                self.status_label.configure(text=f"正在撰写：{name}...", text_color="#1F6AA5")
                self.progressbar.set(i / total)
                
                # 核心 System Prompt：严禁 Markdown
                system_prompt = """
                你是一位专业的高中化学教师。
                【绝对禁止】：
                1. 禁止使用 Markdown 格式（严禁使用 **加粗**、## 标题、- 列表）。
                2. 禁止使用空洞的套话。
                3. 所有段落必须是纯文本，首行不要缩进（留给Word处理）。
                
                【写作要求】：
                1. 使用汉字数字作为标题（一、二、三）。
                2. 内容务实，多结合具体化学教材知识点。
                """
                
                user_prompt = f"""
                题目：{title}
                大纲：{outline}
                当前任务：{instruction}
                请直接输出纯文本内容。
                """

                response = client.chat.completions.create(
                    model=self.api_config.get("model"),
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    stream=True,
                    temperature=0.7
                )

                # 添加内部标记方便用户阅读，导出时会清洗
                self.txt_paper.insert("end", f"\n\n【{name}】\n") 
                
                for chunk in response:
                    if chunk.choices[0].delta.content:
                        content = chunk.choices[0].delta.content
                        self.txt_paper.insert("end", content)
                        self.txt_paper.see("end")
                        full_text += content
                
                self.progressbar.set((i + 1) / total)
                time.sleep(1)

            self.status_label.configure(text="撰写完成！请点击导出 Word。", text_color="green")

        except Exception as e:
            self.status_label.configure(text=f"错误: {str(e)}", text_color="red")
        finally:
            self.btn_gen_paper.configure(state="normal", text="开始深度撰写")

    def save_to_word(self):
        content = self.txt_paper.get("0.0", "end").strip()
        if not content: return
        
        file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
        if file_path:
            doc = Document()
            
            # 设置一个通用的中文字体基础（宋体），方便用户打开
            doc.styles['Normal'].font.name = u'Times New Roman'
            doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
            
            # 1. 写入题目（居中，稍大）
            p_title = doc.add_paragraph()
            p_title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run_title = p_title.add_run(self.entry_title.get())
            run_title.font.size = Pt(16)
            run_title.bold = True
            
            # 2. 写入作者（居中）
            p_author = doc.add_paragraph()
            p_author.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            p_author.add_run(f"{self.entry_author.get()}\n({self.entry_org.get()})")

            doc.add_paragraph() # 空一行

            # 3. 正文清洗与写入
            lines = content.split('\n')
            for line in lines:
                line = line.strip()
                if not line: continue
                
                # 清洗步骤：去除系统生成的【章节名】标记，避免干扰正文
                if line.startswith("【") and line.endswith("】"):
                    continue

                # 清洗步骤：强力去除残留的 Markdown 符号
                # 去除 **bold**, ## Header, - List, > Quote
                clean_line = re.sub(r'\*\*|##|__|```', '', line) 
                # 去除行首的列表符 (如 "- ", "* ")
                if clean_line.startswith("- ") or clean_line.startswith("* "):
                    clean_line = clean_line[2:]
                
                # 写入纯文本段落
                p = doc.add_paragraph(clean_line)
                
                # 简单的格式优化：如果是“一、”开头的，稍微加粗一下，方便识别
                # 但不应用任何Word样式，保持“干净”
                if clean_line.startswith("一、") or clean_line.startswith("二、") or clean_line.startswith("三、") or clean_line.startswith("四、"):
                     p.runs[0].bold = True
                
                # 首行缩进 (2字符)，符合中文习惯，方便排版
                p.paragraph_format.first_line_indent = Pt(24) 

            doc.save(file_path)
            self.status_label.configure(text=f"已导出纯净版: {os.path.basename(file_path)}", text_color="green")

    def load_config(self):
        try:
            with open("config.json", "r") as f: self.api_config = json.load(f)
        except: pass
    def save_config(self):
        self.api_config["api_key"] = self.entry_key.get().strip()
        self.api_config["base_url"] = self.entry_url.get().strip()
        self.api_config["model"] = self.entry_model.get().strip()
        with open("config.json", "w") as f: json.dump(self.api_config, f)

if __name__ == "__main__":
    app = PaperWriterApp()
    app.mainloop()
