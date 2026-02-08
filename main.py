import customtkinter as ctk
import threading
from openai import OpenAI
import os
from docx import Document
from tkinter import filedialog
import json

# --- 配置区域 ---
APP_VERSION = "v1.0.0 (Paper AI)"
DEV_NAME = "俞晋全"
DEV_ORG = "俞晋全高中化学名师工作室"
# ----------------

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class PaperWriterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"AI 期刊论文撰写助手 - {DEV_NAME}")
        self.geometry("900x700")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # 默认配置 (推荐使用 DeepSeek，因为它便宜且中文能力强)
        self.api_config = {
            "api_key": "",
            "base_url": "https://api.deepseek.com", 
            "model": "deepseek-chat"
        }
        self.load_config()

        # --- 主选项卡 ---
        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
        self.tab_outline = self.tabview.add("1. 生成大纲")
        self.tab_write = self.tabview.add("2. 撰写正文")
        self.tab_settings = self.tabview.add("3. 系统设置")

        self.setup_outline_tab()
        self.setup_write_tab()
        self.setup_settings_tab()

        # 状态栏
        self.status_label = ctk.CTkLabel(self, text="就绪 - 请先在设置中配置 API Key", text_color="gray")
        self.status_label.grid(row=1, column=0, pady=10)

    # === Tab 1: 大纲生成 ===
    def setup_outline_tab(self):
        t = self.tab_outline
        t.grid_columnconfigure(0, weight=1)
        t.grid_rowconfigure(2, weight=1) # 让文本框自动伸缩

        ctk.CTkLabel(t, text="请输入论文题目:", font=("Microsoft YaHei UI", 14, "bold")).grid(row=0, column=0, sticky="w", padx=10, pady=(10,0))
        
        self.entry_title = ctk.CTkEntry(t, placeholder_text="例如: 高中化学“素养为本”的教学案例研究", height=40)
        self.entry_title.grid(row=1, column=0, sticky="ew", padx=10, pady=10)

        self.txt_outline = ctk.CTkTextbox(t, font=("Microsoft YaHei UI", 14), height=300)
        self.txt_outline.grid(row=2, column=0, sticky="nsew", padx=10, pady=10)
        self.txt_outline.insert("0.0", "（此处将显示生成的论文大纲，您可以直接修改...）")

        self.btn_gen_outline = ctk.CTkButton(t, text="自动生成大纲", command=self.run_gen_outline, height=40, font=("Microsoft YaHei UI", 14, "bold"))
        self.btn_gen_outline.grid(row=3, column=0, pady=10)

    # === Tab 2: 正文撰写 ===
    def setup_write_tab(self):
        t = self.tab_write
        t.grid_columnconfigure(0, weight=1)
        t.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(t, text="AI 将基于前一页的大纲撰写全文:", font=("Microsoft YaHei UI", 12)).grid(row=0, column=0, sticky="w", padx=10, pady=10)
        
        self.txt_paper = ctk.CTkTextbox(t, font=("Microsoft YaHei UI", 14))
        self.txt_paper.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)

        btn_frame = ctk.CTkFrame(t, fg_color="transparent")
        btn_frame.grid(row=2, column=0, pady=10)
        
        self.btn_gen_paper = ctk.CTkButton(btn_frame, text="开始撰写全文", command=self.run_gen_paper, fg_color="#2CC985", hover_color="#229966")
        self.btn_gen_paper.pack(side="left", padx=10)
        
        self.btn_save_word = ctk.CTkButton(btn_frame, text="导出为 Word", command=self.save_to_word)
        self.btn_save_word.pack(side="left", padx=10)

    # === Tab 3: 设置 ===
    def setup_settings_tab(self):
        t = self.tab_settings
        
        ctk.CTkLabel(t, text="API Key (推荐使用 DeepSeek):").pack(pady=(20, 5))
        self.entry_key = ctk.CTkEntry(t, width=400, show="*")
        self.entry_key.insert(0, self.api_config.get("api_key", ""))
        self.entry_key.pack(pady=5)

        ctk.CTkLabel(t, text="Base URL (例如 https://api.deepseek.com):").pack(pady=(10, 5))
        self.entry_url = ctk.CTkEntry(t, width=400)
        self.entry_url.insert(0, self.api_config.get("base_url", "https://api.deepseek.com"))
        self.entry_url.pack(pady=5)
        
        ctk.CTkLabel(t, text="模型名称 (例如 deepseek-chat):").pack(pady=(10, 5))
        self.entry_model = ctk.CTkEntry(t, width=400)
        self.entry_model.insert(0, self.api_config.get("model", "deepseek-chat"))
        self.entry_model.pack(pady=5)

        ctk.CTkButton(t, text="保存配置", command=self.save_config).pack(pady=20)
        
        help_text = "说明：本软件需要调用大模型 API。\n推荐去 deepseek.com 申请 API Key，价格极低且中文写作能力强。\n如果您有 OpenAI Key，也可以填入并修改 URL。"
        ctk.CTkLabel(t, text=help_text, text_color="gray").pack(pady=20)

    # --- 核心逻辑区 ---
    
    def get_client(self):
        key = self.api_config.get("api_key")
        base = self.api_config.get("base_url")
        if not key:
            self.status_label.configure(text="错误：请先在设置中填写 API Key", text_color="red")
            return None
        return OpenAI(api_key=key, base_url=base)

    def run_gen_outline(self):
        title = self.entry_title.get()
        if not title:
            self.status_label.configure(text="请输入论文题目！", text_color="red")
            return
        threading.Thread(target=self.thread_gen_outline, args=(title,), daemon=True).start()

    def thread_gen_outline(self, title):
        client = self.get_client()
        if not client: return

        self.btn_gen_outline.configure(state="disabled", text="正在思考大纲...")
        self.status_label.configure(text="AI 正在构思大纲，请稍候...", text_color="#1F6AA5")
        
        prompt = f"""
        你是一位资深的学术期刊编辑。请为题目《{title}》写一份详细的论文大纲。
        要求：
        1. 结构符合标准学术论文规范（摘要、引言、正文各章节、结论、参考文献）。
        2. 层级清晰，列出二级标题。
        3. 直接输出大纲内容，不要有多余的寒暄。
        """

        try:
            response = client.chat.completions.create(
                model=self.api_config.get("model"),
                messages=[{"role": "user", "content": prompt}],
                stream=True
            )
            
            self.txt_outline.delete("0.0", "end")
            full_content = ""
            for chunk in response:
                if chunk.choices[0].delta.content:
                    content = chunk.choices[0].delta.content
                    full_content += content
                    self.txt_outline.insert("end", content)
                    self.txt_outline.see("end") # 自动滚动
            
            self.status_label.configure(text="大纲生成完毕，请修改后点击下一步", text_color="green")
            self.tabview.set("1. 生成大纲") # 确保视口在这里

        except Exception as e:
            self.status_label.configure(text=f"API 错误: {str(e)}", text_color="red")
        finally:
            self.btn_gen_outline.configure(state="normal", text="自动生成大纲")

    def run_gen_paper(self):
        outline = self.txt_outline.get("0.0", "end").strip()
        title = self.entry_title.get()
        if len(outline) < 20:
            self.status_label.configure(text="大纲内容太少，请先生成或输入大纲", text_color="red")
            return
        threading.Thread(target=self.thread_gen_paper, args=(title, outline), daemon=True).start()

    def thread_gen_paper(self, title, outline):
        client = self.get_client()
        if not client: return

        self.btn_gen_paper.configure(state="disabled", text="正在疯狂码字中...")
        self.status_label.configure(text="AI 正在撰写全文，这可能需要几分钟...", text_color="#1F6AA5")
        self.txt_paper.delete("0.0", "end")

        prompt = f"""
        你是一位专业的学术研究员。请根据以下题目和大纲，撰写一篇完整的学术论文。
        
        题目：{title}
        大纲：
        {outline}
        
        要求：
        1. 语言学术、严谨，逻辑性强。
        2. 内容要丰满，扩展大纲中的每一个点。
        3. 篇幅要足够长，适合发表。
        4. 使用 Markdown 格式方便阅读。
        """

        try:
            response = client.chat.completions.create(
                model=self.api_config.get("model"),
                messages=[{"role": "user", "content": prompt}],
                stream=True,
                temperature=0.7
            )
            
            full_content = ""
            for chunk in response:
                if chunk.choices[0].delta.content:
                    content = chunk.choices[0].delta.content
                    full_content += content
                    self.txt_paper.insert("end", content)
                    self.txt_paper.see("end")
            
            self.status_label.configure(text="论文撰写完成！您可以导出为 Word。", text_color="green")

        except Exception as e:
            self.status_label.configure(text=f"API 错误: {str(e)}", text_color="red")
        finally:
            self.btn_gen_paper.configure(state="normal", text="开始撰写全文")

    def save_to_word(self):
        content = self.txt_paper.get("0.0", "end").strip()
        if not content:
            return
        
        file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
        if file_path:
            doc = Document()
            doc.add_heading(self.entry_title.get(), 0)
            
            # 简单的 Markdown 转 Word 处理
            for line in content.split('\n'):
                line = line.strip()
                if line.startswith('### '):
                    doc.add_heading(line.replace('### ', ''), level=3)
                elif line.startswith('## '):
                    doc.add_heading(line.replace('## ', ''), level=2)
                elif line.startswith('# '):
                    doc.add_heading(line.replace('# ', ''), level=1)
                else:
                    if line: doc.add_paragraph(line)
            
            doc.save(file_path)
            self.status_label.configure(text=f"已保存至: {os.path.basename(file_path)}", text_color="green")

    def load_config(self):
        try:
            with open("config.json", "r") as f:
                self.api_config = json.load(f)
        except:
            pass

    def save_config(self):
        self.api_config["api_key"] = self.entry_key.get().strip()
        self.api_config["base_url"] = self.entry_url.get().strip()
        self.api_config["model"] = self.entry_model.get().strip()
        
        with open("config.json", "w") as f:
            json.dump(self.api_config, f)
        
        self.status_label.configure(text="配置已保存！现在可以生成大纲了。", text_color="green")

if __name__ == "__main__":
    app = PaperWriterApp()
    app.mainloop()
