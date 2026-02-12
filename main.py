import customtkinter as ctk
from openai import OpenAI
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
from docx.shared import Pt

ctk.set_appearance("system")
ctk.set_default_color_theme("blue")

class WritingAssistant(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("AI 写作助手 - 论文/计划/反思/总结/自定义")
        self.geometry("1200x900")
        self.client = None

        self.create_widgets()

    def create_widgets(self):
        # === API 设置区 ===
        api_frame = ctk.CTkFrame(self)
        api_frame.pack(pady=10, padx=20, fill="x")

        ctk.CTkLabel(api_frame, text="API Key:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.key_entry = ctk.CTkEntry(api_frame, width=350, show="*")
        self.key_entry.grid(row=0, column=1, padx=5, pady=5)

        ctk.CTkLabel(api_frame, text="Base URL:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.url_entry = ctk.CTkEntry(api_frame, width=300)
        self.url_entry.grid(row=0, column=3, padx=5, pady=5)
        self.url_entry.insert(0, "https://api.openai.com/v1")  # 默认 OpenAI

        ctk.CTkLabel(api_frame, text="模型:").grid(row=0, column=4, padx=5, pady=5, sticky="w")
        self.model_combo = ctk.CTkComboBox(api_frame, width=200, values=[
            "gpt-4o", "gpt-4o-mini", "gpt-3.5-turbo",
            "claude-3-5-sonnet-20241022", "claude-3-opus-20240229",
            "llama3-70b-8192", "llama3-8b-8192", "mixtral-8x7b-32768",
            "grok-beta", "deepseek-chat"
        ])
        self.model_combo.set("gpt-4o-mini")
        self.model_combo.grid(row=0, column=5, padx=5, pady=5)

        ctk.CTkButton(api_frame, text=""保存 API 设置", command=self.save_api).grid(row=0, column=6, padx=10, pady=5)

        # === 输入区 ===
        input_frame = ctk.CTkFrame(self)
        input_frame.pack(pady=10, padx=20, fill="x")

        ctk.CTkLabel(input_frame, text="写作类型:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.type_combo = ctk.CTkComboBox(input_frame, values=[
            "期刊论文", "项目计划", "个人反思", "案例分析", "工作总结", "自定义"
        ], command=self.toggle_custom_prompt)
        self.type_combo.set("期刊论文")
        self.type_combo.grid(row=0, column=1, padx=5, pady=5)

        ctk.CTkLabel(input_frame, text="题目/主题:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.title_entry = ctk.CTkEntry(input_frame, width=450)
        self.title_entry.grid(row=0, column=3, padx=5, pady=5)

        ctk.CTkButton(input_frame, text="生成大纲", command=self.generate_outline).grid(row=0, column=4, padx=10, pady=5)

        # 附加参考文献区（可选）
        refs_frame = ctk.CTkFrame(self)
        refs_frame.pack(pady=10, padx=20, fill="x")
        ctk.CTkLabel(refs_frame, text="附加参考文献或材料（可选，会自动引用）:").pack(anchor="w", padx=10)
        self.refs_text = ctk.CTkTextbox(refs_frame, height=100)
        self.refs_text.pack(fill="x", padx=10, pady=5)

        # === 大纲区 ===
        outline_frame = ctk.CTkFrame(self)
        outline_frame.pack(pady=10, padx=20, fill="both", expand=True)

        btn_frame1 = ctk.CTkFrame(outline_frame)
        btn_frame1.pack(fill="x", pady=5)
        ctk.CTkLabel(btn_frame1, text="大纲（可直接编辑）:").pack(side="left", padx=10)
        ctk.CTkButton(btn_frame1, text="清空大纲", command=lambda: self.outline_text.delete("1.0", "end")).pack(side="right", padx=10)

        self.outline_text = ctk.CTkTextbox(outline_frame)
        self.outline_text.pack(fill="both", expand=True, padx=10, pady=5)

        ctk.CTkButton(outline_frame, text="根据大纲生成全文", command=self.generate_full).pack(pady=10)

        # === 结果区 ===
        result_frame = ctk.CTkFrame(self)
        result_frame.pack(pady=10, padx=20, fill="both", expand=True)

        btn_frame2 = ctk.CTkFrame(result_frame)
        btn_frame2.pack(fill="x", pady=5)
        ctk.CTkLabel(btn_frame2, text="生成结果:").pack(side="left", padx=10)
        ctk.CTkButton(btn_frame2, text="清空结果", command=lambda: self.result_text.delete("1.0", "end")).pack(side="right", padx=5)
        ctk.CTkButton(btn_frame2, text="导出 Word", command=self.export_word).pack(side="right", padx=5)
        ctk.CTkButton(btn_frame2, text="导出 Markdown", command=self.export_md).pack(side="right", padx=5)
        ctk.CTkButton(btn_frame2, text="导出 TXT", command=self.export_txt).pack(side="right", padx=5)

        self.result_text = ctk.CTkTextbox(result_frame)
        self.result_text.pack(fill="both", expand=True, padx=10, pady=5)

        # 自定义提示词区
        self.custom_prompt = ctk.CTkTextbox(self, height=120)
        self.custom_prompt.insert("1.0", "在此输入你的详细写作要求和结构...")
        self.toggle_custom_prompt(self.type_combo.get())

    def toggle_custom_prompt(self, choice):
        if choice == "自定义":
            self.custom_prompt.pack(pady=10, padx=20, fill="x")
        else:
            self.custom_prompt.pack_forget()

    def save_api(self):
        api_key = self.key_entry.get().strip()
        base_url = self.url_entry.get().strip() or None
        if not api_key:
            messagebox.showerror("错误", "请填写 API Key")
            return
        self.client = OpenAI(api_key=api_key, base_url=base_url)
        self.model = self.model_combo.get()
        messagebox.showinfo("成功", f"API 设置保存成功\n模型: {self.model}\nBase URL: {base_url or '默认 OpenAI'}")

    def generate_outline(self):
        if not self.client:
            messagebox.showerror("错误", "请先保存 API 设置")
            return
        title = self.title_entry.get().strip()
        if not title:
            messagebox.showwarning("提示", "请填写题目/主题")
            return

        writing_type = self.type_combo.get()
        prompt = self.build_prompt(writing_type, title, is_outline=True)
        self.call_api(prompt, self.outline_text)

    def generate_full(self):
        if not self.client:
            messagebox.showerror("错误", "请先保存 API 设置")
            return
        outline = self.outline_text.get("1.0", "end").strip()
        if not outline:
            messagebox.showwarning("提示", "大纲为空，请先生成或填写大纲")
            return

        title = self.title_entry.get().strip()
        writing_type = self.type_combo.get()
        refs = self.refs_text.get("1.0", "end").strip()
        custom = self.custom_prompt.get("1.0", "end").strip() if writing_type == "自定义" else None

        prompt = self.build_prompt(writing_type, title, is_outline=False, outline=outline, refs=refs, custom=custom)
        self.call_api(prompt, self.result_text, max_tokens=8000)

    def call_api(self, prompt, textbox, max_tokens=2000):
        textbox.delete("1.0", "end")
        textbox.insert("1.0", "正在生成，请稍候...")
        self.update_idletasks()

        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.7 if "outline" in textbox._name else 0.8,
                max_tokens=max_tokens
            )
            content = response.choices[0].message.content.strip()
            textbox.delete("1.0", "end")
            textbox.insert("1.0", content)
        except Exception as e:
            messagebox.showerror("生成失败", str(e))

    def build_prompt(self, writing_type, title, is_outline, outline=None, refs=None, custom=None):
        refs_part = f"\n\n附加参考材料（请在正文中适当位置使用规范引用，如 APA、GB/T 7714 或编号格式）:\n{refs}" if refs else ""

        prompts = {
            "期刊论文": {
                "outline": f"请为题目《{title}》生成一个详细的学术期刊论文大纲。要求使用中文，结构清晰，包括：1. 标题 2. 摘要 3. 关键词 4. 引言 5. 文献综述 6. 研究方法 7. 结果与分析 8. 讨论 9. 结论与展望 10. 参考文献。每节给出简要描述。",
                "full": f"请为题目《{title}》撰写一篇完整的学术期刊论文，语言正式、逻辑严谨、学术规范。严格按照以下大纲撰写，每节内容充实、论证充分：\n\n{outline}{refs_part}"
            },
            "项目计划": {
                "outline": f"请为项目《{title}》制定详细的项目执行计划大纲，包括：背景、目标、范围、阶段划分、时间表、资源需求、风险分析、预算等。",
                "full": f"请为项目《{title}》撰写完整的项目执行计划书，内容专业、结构完整，严格按照以下大纲：\n\n{outline}{refs_part}"
            },
            "个人反思": {
                "outline": f"请针对《{title}》写一篇个人反思的大纲，包括：事件背景、个人感受、具体经历、收获与不足、未来改进等。",
                "full": f"请针对《{title}》撰写一篇深入、真挚的个人反思文章，情感真实、逻辑清晰，严格按照以下大纲：\n\n{outline}{refs_part}"
            },
            "案例分析": {
                "outline": f"请对案例《{title}》进行全面分析的大纲，包括：案例背景、问题描述、分析框架、具体分析、结论与建议等。",
                "full": f"请对案例《{title}》撰写完整的案例分析报告，分析深入、逻辑严密，严格按照以下大纲：\n\n{outline}{refs_part}"
            },
            "工作总结": {
                "outline": f"请为《{title}》撰写工作总结的大纲，包括：工作概述、完成情况、经验教训、存在问题、改进措施等。",
                "full": f"请为《{title}》撰写一份完整的工作总结报告，语言客观、数据详实，严格按照以下大纲：\n\n{outline}{refs_part}"
            },
            "自定义": {
                "outline": custom or title,
                "full": f"{custom}\n\n请严格按照以下大纲/要求撰写完整内容：\n\n{outline}{refs_part}"
            }
        }

        key = "outline" if is_outline else "full"
        return prompts[writing_type][key]

    # === 导出功能 ===
    def export_word(self):
        text = self.result_text.get("1.0", "end").strip()
        if not text:
            return
        file = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word 文件", "*.docx")])
        if file:
            doc = Document()
            doc.add_heading(self.title_entry.get() or "未命名文档", 0)
            for paragraph in text.split("\n\n"):
                p = doc.add_paragraph(paragraph.strip())
                p.style = 'Normal'
            doc.save(file)
            messagebox.showinfo("成功", f"已保存 Word 文件：{file}")

    def export_md(self):
        text = self.result_text.get("1.0", "end").strip()
        if not text:
            return
        file = filedialog.asksaveasfilename(defaultextension=".md", filetypes=[("Markdown 文件", "*.md")])
        if file:
            with open(file, "w", encoding="utf-8") as f:
                f.write(f"# {self.title_entry.get() or '未命名文档'}\n\n{text}")
            messagebox.showinfo("成功", f"已保存 Markdown 文件：{file}")

    def export_txt(self):
        text = self.result_text.get("1.0", "end").strip()
        if not text:
            return
        file = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("文本文件", "*.txt")])
        if file:
            with open(file, "w", encoding="utf-8") as f:
                f.write(text)
            messagebox.showinfo("成功", f"已保存 TXT 文件：{file}")

if __name__ == "__main__":
    app = WritingAssistant()
    app.mainloop()
