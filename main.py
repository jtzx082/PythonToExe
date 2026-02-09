import customtkinter as ctk
import threading
from openai import OpenAI
import os
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from tkinter import filedialog, messagebox
import json
import time
import re

# --- 配置区域 ---
APP_VERSION = "v12.0.0 (Outline-First Workflow)"
DEV_NAME = "俞晋全"
DEV_ORG = "俞晋全高中化学名师工作室"
# ----------------

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# === 动态预设库 ===
# 当用户切换文体时，自动填充这些默认值
PRESET_CONFIGS = {
    "期刊论文": {
        "topic": "高中化学虚拟仿真实验教学的价值与策略研究",
        "instruction": "要求：\n1. 结构包含：摘要、关键词、引言、理论价值、教学策略、结语、参考文献。\n2. 重点写‘教学策略’，结合具体的《氯气》实验案例。\n3. 语气严谨学术。",
        "words": "4000"
    },
    "教学案例": {
        "topic": "《钠与水反应》教学案例分析",
        "instruction": "要求：\n1. 采用叙事风格，描写课堂真实发生的冲突和师生对话。\n2. 重点反思实验演示中出现的意外现象。\n3. 包含：背景、过程描述、分析与反思。",
        "words": "2500"
    },
    "教学反思": {
        "topic": "高三化学二轮复习课后的深刻反思",
        "instruction": "要求：\n1. 使用第一人称‘我’。\n2. 深刻剖析复习课‘满堂灌’的弊端。\n3. 提出具体的改进措施，如‘学生讲题’模式。",
        "words": "1500"
    },
    "工作计划": {
        "topic": "2026年春季学期高二化学备课组工作计划",
        "instruction": "要求：\n1. 条理清晰，多用数据指标。\n2. 包含：指导思想、工作目标、具体措施（教研、培优、实验）、行事历。\n3. 务实可行。",
        "words": "2000"
    },
    "工作总结": {
        "topic": "2025年度个人教学工作总结",
        "instruction": "要求：\n1. 总结本年度的教学成绩、科研成果、班主任工作。\n2. 分析存在的不足。\n3. 数据详实，态度诚恳。",
        "words": "3000"
    }
}

class InteractiveWriterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"全能写作助手 (交互式大纲版) - {DEV_NAME}")
        self.geometry("1200x900")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.api_config = {
            "api_key": "",
            "base_url": "https://api.deepseek.com", 
            "model": "deepseek-chat"
        }
        self.load_config()
        self.stop_event = threading.Event() # 用于控制停止

        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
        self.tab_write = self.tabview.add("智能写作工作台")
        self.tab_settings = self.tabview.add("系统设置")

        self.setup_write_tab()
        self.setup_settings_tab()

        self.status_label = ctk.CTkLabel(self, text="就绪", text_color="gray")
        self.status_label.grid(row=1, column=0, pady=5)
        
        self.progressbar = ctk.CTkProgressBar(self, mode="determinate")
        self.progressbar.grid(row=2, column=0, padx=20, pady=(0, 10), sticky="ew")
        self.progressbar.set(0)

    # === Tab 1: 写作工作台 ===
    def setup_write_tab(self):
        t = self.tab_write
        t.grid_columnconfigure(1, weight=1)
        t.grid_rowconfigure(6, weight=1) # 让正文区自适应

        # 1. 文体选择 (带回调)
        ctk.CTkLabel(t, text="选择文体:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=0, column=0, padx=10, pady=10, sticky="e")
        self.combo_mode = ctk.CTkComboBox(t, values=list(PRESET_CONFIGS.keys()), width=250, command=self.on_mode_change)
        self.combo_mode.set("期刊论文")
        self.combo_mode.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        
        # 2. 标题
        ctk.CTkLabel(t, text="标题/主题:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.entry_topic = ctk.CTkEntry(t, width=500)
        self.entry_topic.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        # 3. 具体指令
        ctk.CTkLabel(t, text="指令要求:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=2, column=0, padx=10, pady=5, sticky="ne")
        self.txt_instructions = ctk.CTkTextbox(t, height=80, font=("Microsoft YaHei UI", 12))
        self.txt_instructions.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        # 4. 字数
        ctk.CTkLabel(t, text="目标字数:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=3, column=0, padx=10, pady=5, sticky="e")
        self.entry_words = ctk.CTkEntry(t, width=150)
        self.entry_words.grid(row=3, column=1, padx=10, pady=5, sticky="w")

        # --- 分割线 ---
        ctk.CTkFrame(t, height=2, fg_color="gray").grid(row=4, column=0, columnspan=2, sticky="ew", padx=10, pady=10)

        # 5. 双面板布局 (左大纲，右正文)
        self.paned_frame = ctk.CTkFrame(t, fg_color="transparent")
        self.paned_frame.grid(row=5, column=0, columnspan=2, sticky="nsew", padx=5)
        self.paned_frame.grid_columnconfigure(0, weight=1) # 左侧权重 1
        self.paned_frame.grid_columnconfigure(1, weight=2) # 右侧权重 2
        self.paned_frame.grid_rowconfigure(1, weight=1)

        # 左侧：大纲区
        ctk.CTkLabel(self.paned_frame, text="第一步：生成并修改大纲", text_color="#1F6AA5", font=("bold", 12)).grid(row=0, column=0, sticky="w", padx=5)
        self.txt_outline = ctk.CTkTextbox(self.paned_frame, height=300, font=("Microsoft YaHei UI", 13))
        self.txt_outline.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        
        btn_outline_frame = ctk.CTkFrame(self.paned_frame, fg_color="transparent")
        btn_outline_frame.grid(row=2, column=0, sticky="ew")
        self.btn_gen_outline = ctk.CTkButton(btn_outline_frame, text="1. 生成大纲", command=self.run_gen_outline, fg_color="#1F6AA5", width=120)
        self.btn_gen_outline.pack(side="left", padx=5, pady=5)
        ctk.CTkButton(btn_outline_frame, text="清空大纲", command=lambda: self.txt_outline.delete("0.0", "end"), fg_color="gray", width=80).pack(side="right", padx=5)

        # 右侧：正文区
        ctk.CTkLabel(self.paned_frame, text="第二步：基于左侧大纲撰写全文", text_color="#2CC985", font=("bold", 12)).grid(row=0, column=1, sticky="w", padx=5)
        self.txt_content = ctk.CTkTextbox(self.paned_frame, height=300, font=("Microsoft YaHei UI", 14))
        self.txt_content.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)
        
        btn_write_frame = ctk.CTkFrame(self.paned_frame, fg_color="transparent")
        btn_write_frame.grid(row=2, column=1, sticky="ew")
        
        self.btn_run_write = ctk.CTkButton(btn_write_frame, text="2. 按大纲撰写全文", command=self.run_full_write, fg_color="#2CC985", font=("bold", 14))
        self.btn_run_write.pack(side="left", padx=5, pady=5)
        
        self.btn_stop = ctk.CTkButton(btn_write_frame, text="🔴 紧急停止", command=self.stop_writing, fg_color="#C0392B", width=100)
        self.btn_stop.pack(side="left", padx=5)

        self.btn_clear_all = ctk.CTkButton(btn_write_frame, text="🧹 清空全部", command=self.clear_all, fg_color="gray", width=100)
        self.btn_clear_all.pack(side="right", padx=5)
        
        self.btn_export = ctk.CTkButton(btn_write_frame, text="导出 Word", command=self.save_to_word, width=100)
        self.btn_export.pack(side="right", padx=5)

        # 初始化默认值
        self.on_mode_change("期刊论文")

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

    # --- 交互逻辑 ---

    def on_mode_change(self, choice):
        """当文体改变时，自动更新预设文本"""
        preset = PRESET_CONFIGS.get(choice, PRESET_CONFIGS["期刊论文"])
        
        self.entry_topic.delete(0, "end")
        self.entry_topic.insert(0, preset["topic"])
        
        self.txt_instructions.delete("0.0", "end")
        self.txt_instructions.insert("0.0", preset["instruction"])
        
        self.entry_words.delete(0, "end")
        self.entry_words.insert(0, preset["words"])

    def clear_all(self):
        """一键清空"""
        self.txt_outline.delete("0.0", "end")
        self.txt_content.delete("0.0", "end")
        self.status_label.configure(text="已清空所有内容", text_color="gray")
        self.progressbar.set(0)

    def stop_writing(self):
        """紧急停止"""
        self.stop_event.set()
        self.status_label.configure(text="已发送停止指令，正在中断...", text_color="red")
        self.btn_run_write.configure(state="normal")
        self.btn_gen_outline.configure(state="normal")

    def get_client(self):
        key = self.api_config.get("api_key")
        base = self.api_config.get("base_url")
        if not key:
            self.status_label.configure(text="错误：请配置 API Key", text_color="red")
            return None
        return OpenAI(api_key=key, base_url=base)

    # --- 核心任务：生成大纲 ---
    def run_gen_outline(self):
        self.stop_event.clear()
        topic = self.entry_topic.get().strip()
        mode = self.combo_mode.get()
        instr = self.txt_instructions.get("0.0", "end").strip()
        
        if not topic:
            self.status_label.configure(text="请先输入标题！", text_color="red")
            return

        threading.Thread(target=self.thread_outline, args=(mode, topic, instr), daemon=True).start()

    def thread_outline(self, mode, topic, instr):
        client = self.get_client()
        if not client: return

        self.btn_gen_outline.configure(state="disabled", text="生成中...")
        self.status_label.configure(text="正在构思大纲结构...", text_color="#1F6AA5")
        
        prompt = f"""
        任务：为《{topic}》写一份【{mode}】的大纲。
        用户的特殊指令：{instr}
        
        要求：
        1. 请列出文章的章节标题（每行一个）。
        2. 不要包含任何 Markdown 符号（如 # 或 *）。
        3. 确保结构完整（如论文需包含摘要、引言、正文各章、结语、参考文献）。
        """
        
        try:
            resp = client.chat.completions.create(
                model=self.api_config.get("model"),
                messages=[{"role": "user", "content": prompt}],
                stream=True
            )
            
            self.txt_outline.delete("0.0", "end")
            full_text = ""
            for chunk in resp:
                if self.stop_event.is_set(): break
                if chunk.choices[0].delta.content:
                    c = chunk.choices[0].delta.content
                    self.txt_outline.insert("end", c)
                    self.txt_outline.see("end")
                    full_text += c
            
            self.status_label.configure(text="大纲生成完毕！请在左侧文本框手动修改，满意后点击'撰写全文'。", text_color="green")

        except Exception as e:
            self.status_label.configure(text=f"API 错误: {str(e)}", text_color="red")
        finally:
            self.btn_gen_outline.configure(state="normal", text="1. 生成大纲")

    # --- 核心任务：撰写全文 ---
    def run_full_write(self):
        self.stop_event.clear()
        
        # 1. 获取用户修改后的大纲
        outline_raw = self.txt_outline.get("0.0", "end").strip()
        if len(outline_raw) < 5:
            self.status_label.configure(text="大纲为空！请先生成或手写大纲。", text_color="red")
            return
            
        # 2. 解析大纲（按行分割）
        sections = [line.strip() for line in outline_raw.split('\n') if line.strip()]
        if not sections: return

        # 3. 获取其他参数
        topic = self.entry_topic.get().strip()
        mode = self.combo_mode.get()
        instr = self.txt_instructions.get("0.0", "end").strip()
        try: total_words = int(self.entry_words.get())
        except: total_words = 3000
        
        threading.Thread(target=self.thread_write, args=(sections, mode, topic, instr, total_words), daemon=True).start()

    def thread_write(self, sections, mode, topic, instr, total_words):
        client = self.get_client()
        if not client: return

        self.btn_run_write.configure(state="disabled", text="写作中...")
        self.txt_content.delete("0.0", "end")
        self.progressbar.set(0)
        
        # 计算每段字数
        avg_words = int(total_words / len(sections))
        
        full_doc = ""
        total_steps = len(sections)

        try:
            for i, section_title in enumerate(sections):
                # 检查是否停止
                if self.stop_event.is_set():
                    self.status_label.configure(text="写作已强制终止。", text_color="red")
                    break

                self.status_label.configure(text=f"正在撰写 ({i+1}/{total_steps}): {section_title}...", text_color="#1F6AA5")
                self.progressbar.set(i / total_steps)

                # 插入标题
                self.txt_content.insert("end", f"\n\n【{section_title}】\n")
                self.txt_content.see("end")

                # 构建 Prompt
                system_prompt = f"""
                你是一位专业的高中化学教师文秘。
                当前任务：根据大纲，撰写文章的【{section_title}】部分。
                文体类型：{mode}
                
                【写作铁律】：
                1. 严禁 Markdown 格式。输出纯文本。
                2. 严格遵守用户指令：{instr}
                3. 内容要务实，多结合具体的化学教学案例或数据。
                """
                
                user_prompt = f"""
                文章标题：{topic}
                当前章节：{section_title}
                参考字数：本章节约 {avg_words} 字
                
                请直接输出正文。
                """

                resp = client.chat.completions.create(
                    model=self.api_config.get("model"),
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    stream=True,
                    temperature=0.8
                )

                chunk_text = ""
                for chunk in resp:
                    if self.stop_event.is_set(): break
                    if chunk.choices[0].delta.content:
                        c = chunk.choices[0].delta.content
                        self.txt_content.insert("end", c)
                        self.txt_content.see("end")
                        chunk_text += c
                
                full_doc += chunk_text
                time.sleep(0.5) 

            if not self.stop_event.is_set():
                self.status_label.configure(text=f"撰写完成！总字数: {len(full_doc)}", text_color="green")
                self.progressbar.set(1)

        except Exception as e:
            self.status_label.configure(text=f"API 错误: {str(e)}", text_color="red")
        finally:
            self.btn_run_write.configure(state="normal", text="2. 按大纲撰写全文")
            self.btn_gen_outline.configure(state="normal")

    def save_to_word(self):
        content = self.txt_content.get("0.0", "end").strip()
        if not content: return
        
        file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
        if file_path:
            doc = Document()
            doc.styles['Normal'].font.name = u'Times New Roman'
            doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
            
            # 标题
            p_title = doc.add_paragraph()
            p_title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run_title = p_title.add_run(self.entry_topic.get())
            run_title.font.size = Pt(16)
            run_title.bold = True
            run_title.font.name = u'黑体'
            run_title._element.rPr.rFonts.set(qn('w:eastAsia'), u'黑体')
            
            doc.add_paragraph()

            lines = content.split('\n')
            for line in lines:
                line = line.strip()
                if not line: continue

                # 识别标题标记
                if line.startswith("【") and line.endswith("】"):
                    header = line.replace("【", "").replace("】", "")
                    p = doc.add_paragraph()
                    p.paragraph_format.space_before = Pt(12)
                    run = p.add_run(header)
                    run.bold = True
                    run.font.size = Pt(14)
                    run.font.name = u'黑体'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), u'黑体')
                else:
                    clean_line = re.sub(r'\*\*|##|__|```', '', line)
                    if clean_line.startswith("- ") or clean_line.startswith("* "): clean_line = clean_line[2:]
                    p = doc.add_paragraph(clean_line)
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
    app = InteractiveWriterApp()
    app.mainloop()
