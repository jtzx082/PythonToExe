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
import re

# --- 配置区域 ---
APP_VERSION = "v6.0.0 (Word Count Fix)"
DEV_NAME = "俞晋全"
DEV_ORG = "俞晋全高中化学名师工作室"
# ----------------

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class PaperWriterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"期刊论文撰写系统 (精准控字版) - {DEV_NAME}")
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

        ctk.CTkLabel(t, text="论文题目:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=0, column=0, padx=10, pady=10, sticky="e")
        self.entry_title = ctk.CTkEntry(t, placeholder_text="例如：高中化学虚拟仿真实验教学的价值与策略研究", height=35)
        self.entry_title.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        ctk.CTkLabel(t, text="作者姓名:", font=("Microsoft YaHei UI", 12)).grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.entry_author = ctk.CTkEntry(t, placeholder_text="俞晋全")
        self.entry_author.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        ctk.CTkLabel(t, text="单位信息:", font=("Microsoft YaHei UI", 12)).grid(row=2, column=0, padx=10, pady=5, sticky="e")
        self.entry_org = ctk.CTkEntry(t, placeholder_text="甘肃省金塔县中学, 甘肃金塔 735300")
        self.entry_org.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        # 字数控制
        ctk.CTkLabel(t, text="期望字数:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=3, column=0, padx=10, pady=5, sticky="e")
        self.entry_word_count = ctk.CTkEntry(t, placeholder_text="4000")
        self.entry_word_count.insert(0, "4000") 
        self.entry_word_count.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
        
        hint = ctk.CTkLabel(t, text="提示：系统已内置'防爆字'算法，您输入多少，最终结果就会非常接近多少。", text_color="gray", font=("Arial", 10))
        hint.grid(row=4, column=1, sticky="w", padx=10)

        ctk.CTkLabel(t, text="结构大纲 (预览/修改):", font=("Microsoft YaHei UI", 12, "bold")).grid(row=5, column=0, padx=10, pady=(10,0), sticky="nw")
        self.txt_outline = ctk.CTkTextbox(t, height=220, font=("Microsoft YaHei UI", 13))
        self.txt_outline.grid(row=5, column=1, padx=10, pady=10, sticky="nsew")
        
        self.btn_gen_outline = ctk.CTkButton(t, text="生成标准大纲", command=self.run_gen_outline, fg_color="#1F6AA5")
        self.btn_gen_outline.grid(row=6, column=1, pady=10, sticky="e")

    # === Tab 2: 深度撰写 ===
    def setup_write_tab(self):
        t = self.tab_write
        t.grid_columnconfigure(0, weight=1)
        t.grid_rowconfigure(1, weight=1)

        info = "提示：生成内容为纯净无格式文本，导出 Word 后请全选直接设置字体。"
        ctk.CTkLabel(t, text=info, text_color="gray").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        
        self.txt_paper = ctk.CTkTextbox(t, font=("Microsoft YaHei UI", 14))
        self.txt_paper.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)

        btn_frame = ctk.CTkFrame(t, fg_color="transparent")
        btn_frame.grid(row=2, column=0, pady=10)
        
        self.btn_gen_paper = ctk.CTkButton(btn_frame, text="精准控字撰写", command=self.run_deep_write, 
                                           width=200, height=40, font=("Microsoft YaHei UI", 14, "bold"))
        self.btn_gen_paper.pack(side="left", padx=20)
        
        self.btn_save_word = ctk.CTkButton(btn_frame, text="导出纯净 Word", command=self.save_to_word,
                                           fg_color="#2CC985", width=150, height=40)
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
        请为高中化学教学论文《{title}》设计大纲。
        要求：
        1. 包含：摘要、关键词、一、引言；二、理论/价值；三、策略/实践；四、结语；参考文献。
        2. 正文标题使用汉字数字（一、二...）。
        3. 直接输出文本，不要使用 Markdown 符号。
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
        except: total_words = 4000
        
        if len(outline) < 10: return
        threading.Thread(target=self.thread_deep_write, args=(title, outline, total_words), daemon=True).start()

    def thread_deep_write(self, title, outline, target_total_words):
        client = self.get_client()
        if not client: return

        self.btn_gen_paper.configure(state="disabled", text="正在撰写(已启用控字算法)...")
        self.txt_paper.delete("0.0", "end")
        self.progressbar.set(0)

        # === 核心修改：字数压制算法 ===
        # AI 通常会超写 50%~80%，所以我们在 Prompt 中请求的字数要打折
        # 衰减系数 0.65：如果用户要 1000 字，我们只叫 AI 写 650 字，AI 发挥完正好 1000 左右
        dampening_factor = 0.65 
        
        adjusted_total = target_total_words * dampening_factor

        # 重新分配各章节的“Prompt 请求字数”
        w_intro = int(adjusted_total * 0.15)
        w_theory = int(adjusted_total * 0.20)
        w_practice = int(adjusted_total * 0.55) # 重点部分
        w_concl = int(adjusted_total * 0.10)

        sections = [
            ("摘要与关键词", f"请撰写【摘要】（控制在250-300字）和【关键词】。不要超字数。"),
            ("一、引言与背景", f"撰写论文引言。分析背景痛点。请严格控制字数在 {w_intro} 字左右，切勿啰嗦。"),
            ("二、核心价值", f"撰写理论价值部分。逻辑要紧凑。请严格控制字数在 {w_theory} 字左右。"),
            ("三、教学策略与实践（重点）", f"撰写实践部分。包含具体化学实验案例。请严格控制字数在 {w_practice} 字左右，不要写成长篇大论。"),
            ("四、结语与参考文献", f"撰写简短的结语和参考文献。")
        ]

        full_text = ""
        total_steps = len(sections)

        try:
            for i, (name, instruction) in enumerate(sections):
                self.status_label.configure(text=f"正在撰写：{name}...", text_color="#1F6AA5")
                self.progressbar.set(i / total_steps)
                
                system_prompt = """
                你是一位高中化学教师。
                【严格约束】：
                1. 严禁使用 Markdown 格式（不要加粗，不要标题符）。
                2. 严禁字数超标！请简练、务实地写作。
                3. 直接输出纯文本段落。
                """
                
                user_prompt = f"""
                题目：{title}
                大纲：{outline}
                当前任务：{instruction}
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

                self.txt_paper.insert("end", f"\n\n【{name}】\n") 
                
                for chunk in response:
                    if chunk.choices[0].delta.content:
                        content = chunk.choices[0].delta.content
                        self.txt_paper.insert("end", content)
                        self.txt_paper.see("end")
                        full_text += content
                
                self.progressbar.set((i + 1) / total_steps)
                time.sleep(1)

            # 最终统计
            actual_len = len(full_text)
            self.status_label.configure(text=f"撰写完成！目标: {target_total_words}, 实际: {actual_len} 字。", text_color="green")

        except Exception as e:
            self.status_label.configure(text=f"错误: {str(e)}", text_color="red")
        finally:
            self.btn_gen_paper.configure(state="normal", text="精准控字撰写")

    def save_to_word(self):
        content = self.txt_paper.get("0.0", "end").strip()
        if not content: return
        
        file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
        if file_path:
            doc = Document()
            
            # 设置基础字体（宋体）
            doc.styles['Normal'].font.name = u'Times New Roman'
            doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
            
            # 头部信息
            p_title = doc.add_paragraph()
            p_title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run_title = p_title.add_run(self.entry_title.get())
            run_title.font.size = Pt(16)
            run_title.bold = True
            
            p_author = doc.add_paragraph()
            p_author.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            p_author.add_run(f"{self.entry_author.get()}\n({self.entry_org.get()})")

            doc.add_paragraph()

            # 正文清洗与写入
            lines = content.split('\n')
            for line in lines:
                line = line.strip()
                if not line: continue
                
                # 过滤系统标记
                if line.startswith("【") and line.endswith("】"): continue

                # 强力清洗 Markdown
                clean_line = re.sub(r'\*\*|##|__|```', '', line) 
                if clean_line.startswith("- ") or clean_line.startswith("* "):
                    clean_line = clean_line[2:]
                
                p = doc.add_paragraph(clean_line)
                
                # 简单加粗一级标题
                if clean_line.startswith("一、") or clean_line.startswith("二、") or clean_line.startswith("三、") or clean_line.startswith("四、"):
                     if p.runs: p.runs[0].bold = True
                
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
