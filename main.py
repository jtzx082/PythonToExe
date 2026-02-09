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
APP_VERSION = "v11.0.0 (Hard-Coded Structure)"
DEV_NAME = "俞晋全"
DEV_ORG = "俞晋全高中化学名师工作室"
# ----------------

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# === 核心：文体结构强制模板 ===
# 这里定义了每种文体必须包含的“骨架”。AI 无法跳过，必须填空。
TEMPLATE_CONFIG = {
    "期刊论文 (标准学术)": [
        {"title": "摘要与关键词", "prompt": "请写一段300字的【摘要】，概括研究背景、方法、结论。紧接着列出3-5个【关键词】。格式：\n摘要：...\n关键词：..."},
        {"title": "一、问题的提出", "prompt": "请撰写论文的第一部分（引言）。分析当前教学的痛点或背景，引出本文的研究意义。"},
        {"title": "二、核心概念与价值", "prompt": "请撰写论文的理论部分。结合高中化学核心素养或具体教育理论，阐述本研究的教学价值。"},
        {"title": "三、教学策略与实践", "prompt": "这是论文的重点。请分点阐述具体的教学策略（例如：1. ...; 2. ...）。必须结合具体的化学教学案例或实验细节，写深写透。"},
        {"title": "四、成效与反思", "prompt": "请撰写教学成效（学生的改变）以及教学反思（存在的不足）。"},
        {"title": "参考文献", "prompt": "请列出5-8条规范的参考文献（GB/T 7714格式）。"}
    ],
    "教学案例 (叙事风格)": [
        {"title": "一、案例背景", "prompt": "介绍这节课的教材分析、学情分析以及教学目标。"},
        {"title": "二、案例描述 (教学片段)", "prompt": "请用生动的语言描述课堂上发生的真实情境。包括师生对话、实验现象、突发状况。像讲故事一样写。"},
        {"title": "三、案例分析", "prompt": "针对上述片段进行深入分析。为什么会出现这种情况？体现了什么教育理念？"},
        {"title": "四、教学反思", "prompt": "作为教师，通过这个案例，你得到了什么启示？后续如何改进？"}
    ],
    "教学反思 (个人独白)": [
        {"title": "一、教学意图与设想", "prompt": "简述这节课原本的设计思路是怎样的。"},
        {"title": "二、教学过程中的亮点", "prompt": "反思这节课哪里上得好？学生的哪些反应超出了预期？"},
        {"title": "三、存在的问题与遗憾", "prompt": "诚恳地剖析这节课的败笔。是时间没把控好？还是实验演示失败了？"},
        {"title": "四、改进措施", "prompt": "如果重上这节课，我会怎么做？"}
    ],
    "工作计划 (行政公文)": [
        {"title": "一、指导思想", "prompt": "简述本学期工作的指导思想，贯彻什么教育方针。"},
        {"title": "二、工作目标", "prompt": "列出具体的量化目标（如：及格率、优分率、教研活动次数）。"},
        {"title": "三、重点工作与措施", "prompt": "分条列出具体要做的事情。例如：1. 抓好常规教学... 2. 推进实验改革..."},
        {"title": "四、行事历与时间节点", "prompt": "按月份列出大致的工作安排（9月干什么，10月干什么...）。"}
    ],
    "工作总结 (汇报材料)": [
        {"title": "一、工作概况", "prompt": "回顾本学期/本年度的基本情况。"},
        {"title": "二、主要成绩与经验", "prompt": "重点写做成了哪些事？有哪些亮点？（多用数据说话）。"},
        {"title": "三、存在的问题", "prompt": "客观分析当前工作中遇到的困难和不足。"},
        {"title": "四、下一步打算", "prompt": "简述未来的工作方向。"}
    ]
}

class StructureWriterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"全能结构化写作助手 - {DEV_NAME}")
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
        
        self.tab_write = self.tabview.add("1. 结构化写作")
        self.tab_settings = self.tabview.add("2. 系统设置")

        self.setup_write_tab()
        self.setup_settings_tab()

        self.status_label = ctk.CTkLabel(self, text="就绪", text_color="gray")
        self.status_label.grid(row=1, column=0, pady=5)
        
        self.progressbar = ctk.CTkProgressBar(self, mode="determinate")
        self.progressbar.grid(row=2, column=0, padx=20, pady=(0, 10), sticky="ew")
        self.progressbar.set(0)

    # === Tab 1: 写作界面 ===
    def setup_write_tab(self):
        t = self.tab_write
        t.grid_columnconfigure(1, weight=1)
        t.grid_rowconfigure(5, weight=1)

        # 1. 文体选择
        ctk.CTkLabel(t, text="选择文体:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=0, column=0, padx=10, pady=10, sticky="e")
        self.combo_mode = ctk.CTkComboBox(t, values=list(TEMPLATE_CONFIG.keys()), width=250)
        self.combo_mode.set("期刊论文 (标准学术)")
        self.combo_mode.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        
        # 2. 标题/主题
        ctk.CTkLabel(t, text="文章标题:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.entry_topic = ctk.CTkEntry(t, placeholder_text="例如：核心素养视域下的高中化学大单元教学设计", width=500)
        self.entry_topic.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        # 3. 具体要求
        ctk.CTkLabel(t, text="具体指令:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=2, column=0, padx=10, pady=5, sticky="ne")
        self.txt_instructions = ctk.CTkTextbox(t, height=60, font=("Microsoft YaHei UI", 12))
        self.txt_instructions.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        self.txt_instructions.insert("0.0", "例如：要结合具体的《氯气》实验案例；语气要严谨；字数要充足。")

        # 4. 字数系数
        ctk.CTkLabel(t, text="篇幅控制:", font=("Microsoft YaHei UI", 12, "bold")).grid(row=3, column=0, padx=10, pady=5, sticky="e")
        self.combo_length = ctk.CTkComboBox(t, values=["标准篇幅 (约3000字)", "长篇深度 (约5000字)", "短篇精简 (约1500字)"], width=200)
        self.combo_length.set("标准篇幅 (约3000字)")
        self.combo_length.grid(row=3, column=1, padx=10, pady=5, sticky="w")

        # 5. 提示
        info = ctk.CTkLabel(t, text="说明：系统将严格按照【摘要-引言-正文-参考文献】等固定板块逐一撰写，绝不混淆。", text_color="#1F6AA5")
        info.grid(row=4, column=1, sticky="w", padx=10)

        # 6. 输出区
        self.txt_output = ctk.CTkTextbox(t, font=("Microsoft YaHei UI", 14))
        self.txt_output.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

        # 7. 按钮
        btn_frame = ctk.CTkFrame(t, fg_color="transparent")
        btn_frame.grid(row=6, column=0, columnspan=2, pady=10)
        
        self.btn_run = ctk.CTkButton(btn_frame, text="按标准结构撰写", command=self.run_structured_writing, 
                                     width=200, height=40, font=("Microsoft YaHei UI", 14, "bold"), fg_color="#1F6AA5")
        self.btn_run.pack(side="left", padx=20)
        
        self.btn_save = ctk.CTkButton(btn_frame, text="导出格式化 Word", command=self.save_to_word,
                                        width=150, height=40, fg_color="#2CC985")
        self.btn_save.pack(side="left", padx=20)

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

    # --- 逻辑核心 ---

    def get_client(self):
        key = self.api_config.get("api_key")
        base = self.api_config.get("base_url")
        if not key:
            self.status_label.configure(text="错误：请配置 API Key", text_color="red")
            return None
        return OpenAI(api_key=key, base_url=base)

    def run_structured_writing(self):
        topic = self.entry_topic.get().strip()
        mode = self.combo_mode.get()
        instructions = self.txt_instructions.get("0.0", "end").strip()
        length_opt = self.combo_length.get()
        
        if not topic:
            self.status_label.configure(text="请输入标题！", text_color="red")
            return

        threading.Thread(target=self.thread_write, args=(mode, topic, instructions, length_opt), daemon=True).start()

    def thread_write(self, mode, topic, instructions, length_opt):
        client = self.get_client()
        if not client: return

        self.btn_run.configure(state="disabled", text="正在分板块撰写...")
        self.txt_output.delete("0.0", "end")
        self.progressbar.set(0)

        # 1. 获取强制模板
        template_sections = TEMPLATE_CONFIG.get(mode)
        if not template_sections:
            self.status_label.configure(text="错误：未找到该文体的模板", text_color="red")
            return

        # 2. 字数系数
        length_factor = 1.0
        if "长篇" in length_opt: length_factor = 1.5
        if "短篇" in length_opt: length_factor = 0.5

        full_text = ""
        total_steps = len(template_sections)

        # 3. 循环执行每一个板块（物理隔离，确保结构不乱）
        try:
            for i, section in enumerate(template_sections):
                section_title = section["title"]
                section_prompt = section["prompt"]
                
                self.status_label.configure(text=f"正在撰写 ({i+1}/{total_steps}): {section_title}...", text_color="#1F6AA5")
                self.progressbar.set(i / total_steps)

                # 插入显眼的标题分隔符
                self.txt_output.insert("end", f"\n\n【{section_title}】\n")
                self.txt_output.see("end")

                # 构建 Prompt
                system_prompt = f"""
                你是一位专业的高中化学教师文秘。
                当前任务：撰写文章的【{section_title}】部分。
                
                【绝对规则】：
                1. 只写这一个部分的内容，不要写其他部分的。
                2. 严禁使用 Markdown（**加粗**等）。
                3. 必须输出纯文本段落。
                4. 严格遵守用户的具体指令：{instructions}
                """
                
                user_prompt = f"""
                文章标题：{topic}
                当前板块：{section_title}
                板块要求：{section_prompt}
                篇幅要求：请根据内容需要，写 {int(500 * length_factor)} 字左右。
                """

                # 请求 AI
                response = client.chat.completions.create(
                    model=self.api_config.get("model"),
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    stream=True,
                    temperature=0.8
                )

                chunk_text = ""
                for chunk in response:
                    if chunk.choices[0].delta.content:
                        content = chunk.choices[0].delta.content
                        self.txt_output.insert("end", content)
                        self.txt_output.see("end")
                        chunk_text += content
                
                full_text += chunk_text
                time.sleep(1) # 休息防封

            self.status_label.configure(text="撰写完成！结构已强制锁定。", text_color="green")
            self.progressbar.set(1)

        except Exception as e:
            self.status_label.configure(text=f"API 错误: {str(e)}", text_color="red")
        finally:
            self.btn_run.configure(state="normal", text="按标准结构撰写")

    def save_to_word(self):
        content = self.txt_output.get("0.0", "end").strip()
        if not content: return
        
        file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
        if file_path:
            doc = Document()
            # 设置基础字体
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

                # 识别强制插入的标题标记 【XXX】
                if line.startswith("【") and line.endswith("】"):
                    # 提取标题文字
                    header_text = line.replace("【", "").replace("】", "")
                    
                    # 创建一级标题 (Heading 1)
                    p = doc.add_paragraph()
                    p.paragraph_format.space_before = Pt(12)
                    p.paragraph_format.space_after = Pt(6)
                    run = p.add_run(header_text)
                    run.bold = True
                    run.font.size = Pt(14)
                    run.font.name = u'黑体'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), u'黑体')
                else:
                    # 普通正文
                    # 清洗 markdown
                    clean_line = re.sub(r'\*\*|##|__|```', '', line)
                    if clean_line.startswith("- ") or clean_line.startswith("* "): clean_line = clean_line[2:]
                    
                    p = doc.add_paragraph(clean_line)
                    p.paragraph_format.first_line_indent = Pt(24) # 首行缩进

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
    app = StructureWriterApp()
    app.mainloop()
