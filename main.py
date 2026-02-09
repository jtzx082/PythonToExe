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
APP_VERSION = "v10.0.0 (Universal Teacher Assistant)"
DEV_NAME = "俞晋全"
DEV_ORG = "俞晋全高中化学名师工作室"
# ----------------

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# 文体配置库 (不同文体对应不同的 AI 人设和策略)
MODE_CONFIG = {
    "期刊论文": {
        "persona": "你是一位资深的高中化学特级教师，擅长写发表在核心期刊的学术论文。风格严谨、务实，结合核心素养。",
        "structure_prompt": "请设计一份标准期刊论文大纲（摘要、引言、理论、实践、结语）。",
        "temp": 0.85
    },
    "教学案例": {
        "persona": "你是一位善于观察的高中化学教师。请用‘叙事研究’的笔法，生动地描写课堂上发生的真实故事。多写对话、动作、冲突。",
        "structure_prompt": "请设计一份教学案例结构（背景、案例描述/教学片段、分析与反思）。",
        "temp": 0.95  # 案例需要高创造性
    },
    "教学反思": {
        "persona": "你是一位正在深夜备课的化学老教师。请用第一人称‘我’，诚恳地剖析自己教学中的得失。不要说套话，要说心里话。",
        "structure_prompt": "请设计一份深度教学反思结构（教学初衷、实际现象、问题归因、改进设想）。",
        "temp": 0.9
    },
    "工作总结": {
        "persona": "你是一位教学主任或骨干教师。请写一份条理清晰、数据详实的工作总结。既要展示成绩，也要分析不足。",
        "structure_prompt": "请设计一份工作总结结构（工作概况、重点成绩、存在问题、未来规划）。",
        "temp": 0.7   # 总结需要稳重
    },
    "工作计划": {
        "persona": "你是一位思维缜密的教研组长。请写一份可执行性强的工作计划。包含具体目标、实施步骤、时间节点。",
        "structure_prompt": "请设计一份工作计划结构（指导思想、工作目标、具体措施、行事历）。",
        "temp": 0.7
    }
}

class UniversalWriterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"全能化学教师写作助手 - {DEV_NAME}")
        self.geometry("1200x850")
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
        
        self.tab_write = self.tabview.add("1. 智能写作")
        self.tab_settings = self.tabview.add("2. 系统设置")

        self.setup_write_tab()
        self.setup_settings_tab()

        self.status_label = ctk.CTkLabel(self, text="就绪", text_color="gray")
        self.status_label.grid(row=1, column=0, pady=5)
        
        self.progressbar = ctk.CTkProgressBar(self, mode="determinate")
        self.progressbar.grid(row=2, column=0, padx=20, pady=(0, 10), sticky="ew")
        self.progressbar.set(0)

    # === Tab 1: 智能写作 (核心界面) ===
    def setup_write_tab(self):
        t = self.tab_write
        t.grid_columnconfigure(1, weight=1)
        t.grid_rowconfigure(4, weight=1) # 让文本框自适应

        # 1. 文体选择
        ctk.CTkLabel(t, text="文体类型:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=0, column=0, padx=10, pady=10, sticky="e")
        self.combo_mode = ctk.CTkComboBox(t, values=list(MODE_CONFIG.keys()), width=200)
        self.combo_mode.set("期刊论文")
        self.combo_mode.grid(row=0, column=1, padx=10, pady=10, sticky="w")

        # 2. 核心主题
        ctk.CTkLabel(t, text="文章标题/主题:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.entry_topic = ctk.CTkEntry(t, placeholder_text="例如：高一化学《钠及其化合物》教学反思 / 2026年春季学期工作计划", width=500)
        self.entry_topic.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        # 3. 具体指令 (Prompt Injection)
        ctk.CTkLabel(t, text="具体指令与要求:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=2, column=0, padx=10, pady=5, sticky="ne")
        self.txt_instructions = ctk.CTkTextbox(t, height=80, font=("Microsoft YaHei UI", 12))
        self.txt_instructions.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        self.txt_instructions.insert("0.0", "例如：重点分析学生在实验操作中的安全意识问题；字数不要太长；语气要诚恳；多举具体的例子。")

        # 4. 字数控制
        ctk.CTkLabel(t, text="期望字数:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=3, column=0, padx=10, pady=5, sticky="e")
        self.entry_length = ctk.CTkEntry(t, placeholder_text="3000", width=100)
        self.entry_length.insert(0, "3000")
        self.entry_length.grid(row=3, column=1, padx=10, pady=5, sticky="w")

        # 5. 结果展示区
        self.txt_output = ctk.CTkTextbox(t, font=("Microsoft YaHei UI", 14))
        self.txt_output.grid(row=4, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

        # 6. 按钮区
        btn_frame = ctk.CTkFrame(t, fg_color="transparent")
        btn_frame.grid(row=5, column=0, columnspan=2, pady=10)
        
        self.btn_run = ctk.CTkButton(btn_frame, text="开始智能撰写", command=self.run_writing, 
                                     width=200, height=40, font=("Microsoft YaHei UI", 14, "bold"), fg_color="#1F6AA5")
        self.btn_run.pack(side="left", padx=20)
        
        self.btn_export = ctk.CTkButton(btn_frame, text="导出纯净 Word", command=self.save_to_word,
                                        width=150, height=40, fg_color="#2CC985")
        self.btn_export.pack(side="left", padx=20)

    # === Tab 2: 设置 ===
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

    # --- 核心逻辑 ---

    def get_client(self):
        key = self.api_config.get("api_key")
        base = self.api_config.get("base_url")
        if not key:
            self.status_label.configure(text="错误：请配置 API Key", text_color="red")
            return None
        return OpenAI(api_key=key, base_url=base)

    def run_writing(self):
        topic = self.entry_topic.get().strip()
        mode = self.combo_mode.get()
        instructions = self.txt_instructions.get("0.0", "end").strip()
        
        try: target_len = int(self.entry_length.get())
        except: target_len = 3000

        if not topic:
            self.status_label.configure(text="请输入标题或主题！", text_color="red")
            return

        threading.Thread(target=self.thread_write_process, args=(mode, topic, instructions, target_len), daemon=True).start()

    def thread_write_process(self, mode, topic, instructions, target_len):
        client = self.get_client()
        if not client: return

        self.btn_run.configure(state="disabled", text="正在规划结构...")
        self.txt_output.delete("0.0", "end")
        self.progressbar.set(0)

        # 1. 获取该文体的配置
        config = MODE_CONFIG.get(mode, MODE_CONFIG["期刊论文"])
        
        # 2. 第一步：生成动态结构（大纲）
        self.status_label.configure(text=f"正在为【{mode}】设计结构...", text_color="#1F6AA5")
        
        structure_prompt = f"""
        任务：为《{topic}》写一份【{mode}】的写作大纲。
        用户的特殊要求：{instructions}
        
        要求：
        1. 必须包含4-6个主要章节（一级标题）。
        2. 请直接输出标题列表，每行一个。例如：
           一、工作回顾
           二、主要成绩
           ...
        3. 不要包含任何 Markdown 符号。
        """
        
        try:
            resp = client.chat.completions.create(
                model=self.api_config.get("model"),
                messages=[{"role": "user", "content": structure_prompt}],
                temperature=0.7
            )
            outline_raw = resp.choices[0].message.content
            
            # 简单的解析，提取章节
            sections = []
            for line in outline_raw.split('\n'):
                line = line.strip()
                # 过滤空行或太短的行
                if len(line) > 2 and (line[0].isdigit() or line[0] in ['一','二','三','四','五','六','七','八'] or "摘要" in line):
                     sections.append(line)
            
            # 如果提取失败，使用兜底结构
            if len(sections) < 2:
                sections = ["一、背景与目的", "二、过程与实施", "三、成效与分析", "四、问题与反思", "五、未来展望"]

            # 3. 第二步：分段撰写
            full_text = ""
            total_sections = len(sections)
            # 计算每段大概字数 (打个折，防止AI写超)
            chunk_len = int(target_len * 0.7 / total_sections) 

            for i, section_title in enumerate(sections):
                self.status_label.configure(text=f"正在撰写 ({i+1}/{total_sections}): {section_title}...", text_color="#1F6AA5")
                self.progressbar.set(i / total_sections)
                
                # 构建核心 System Prompt (人设 + 要求)
                sys_prompt = f"""
                {config['persona']}
                
                【写作铁律 - 严禁AI味】：
                1. 严禁使用 Markdown 格式。
                2. 严禁使用“综上所述、总而言之、多维互动”等词。
                3. 请严格遵守用户的【具体指令】：{instructions}
                4. 当前文体是【{mode}】，请确保语体风格正确（例如：教学反思要主观，工作总结要客观）。
                """
                
                user_prompt = f"""
                文章主题：{topic}
                当前章节：{section_title}
                参考字数：约 {chunk_len} 字
                
                请直接撰写本章节的正文内容。不要重复标题。
                """

                stream_resp = client.chat.completions.create(
                    model=self.api_config.get("model"),
                    messages=[
                        {"role": "system", "content": sys_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    stream=True,
                    temperature=config['temp'] # 使用不同文体的随机性设置
                )

                # 插入章节标头
                self.txt_output.insert("end", f"\n\n【{section_title}】\n")
                
                for chunk in stream_resp:
                    if chunk.choices[0].delta.content:
                        content = chunk.choices[0].delta.content
                        self.txt_output.insert("end", content)
                        self.txt_output.see("end")
                        full_text += content
                
                time.sleep(1) # 防封

            self.status_label.configure(text=f"撰写完成！总字数: {len(full_text)}", text_color="green")
            self.progressbar.set(1)

        except Exception as e:
            self.status_label.configure(text=f"API 错误: {str(e)}", text_color="red")
        finally:
            self.btn_run.configure(state="normal", text="开始智能撰写")

    def save_to_word(self):
        content = self.txt_output.get("0.0", "end").strip()
        if not content: return
        
        file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
        if file_path:
            doc = Document()
            # 设置中文字体基础
            doc.styles['Normal'].font.name = u'Times New Roman'
            doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
            
            # 标题
            p_title = doc.add_paragraph()
            p_title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run_title = p_title.add_run(self.entry_topic.get())
            run_title.font.size = Pt(16)
            run_title.bold = True
            
            doc.add_paragraph()

            # 正文
            lines = content.split('\n')
            for line in lines:
                line = line.strip()
                if not line: continue
                if line.startswith("【") and line.endswith("】"): continue

                # 清洗 Markdown
                clean_line = re.sub(r'\*\*|##|__|```', '', line) 
                if clean_line.startswith("- ") or clean_line.startswith("* "): clean_line = clean_line[2:]
                
                p = doc.add_paragraph(clean_line)
                
                # 简单识别标题
                if clean_line.startswith("一、") or clean_line.startswith("二、") or clean_line.startswith("三、") or clean_line.startswith("四、"):
                     if p.runs: p.runs[0].bold = True
                
                p.paragraph_format.first_line_indent = Pt(24)

            doc.save(file_path)
            self.status_label.configure(text=f"已导出: {os.path.basename(file_path)}", text_color="green")

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
    app = UniversalWriterApp()
    app.mainloop()
