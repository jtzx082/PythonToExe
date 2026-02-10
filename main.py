import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import os
import json
from datetime import datetime
from openai import OpenAI
# --- 扩展功能库 ---
from duckduckgo_search import DDGS
import pypdf
from docx import Document

# --- 配置区域 ---
APP_NAME = "DeepSeek Pro 桌面版"
APP_VERSION = "v1.0.0"
DEV_INFO = "开发者：Yu Jinquan\n基于 DeepSeek-V3/R1 API"

# 默认配置
DEFAULT_CONFIG = {
    "api_key": "",
    "model": "deepseek-chat",  # deepseek-chat (V3) 或 deepseek-reasoner (R1)
    "temperature": 1.3,
    "use_search": False,
    "system_prompt": "你是一个乐于助人的AI助手。"
}

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class DeepSeekApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} {APP_VERSION}")
        self.geometry("1100x800")
        
        self.config = self.load_config()
        self.chat_history = [] # 存储对话上下文
        self.client = None
        self.stop_generation = False
        self.attached_content = "" # 附件内容缓存

        self.setup_ui()
        
        # 如果有Key，预初始化
        if self.config["api_key"]:
            self.init_client()

    def load_config(self):
        if os.path.exists("config.json"):
            try:
                with open("config.json", "r") as f:
                    return json.load(f)
            except: pass
        return DEFAULT_CONFIG.copy()

    def save_config(self):
        with open("config.json", "w") as f:
            json.dump(self.config, f)

    def init_client(self):
        if not self.config["api_key"]: return
        self.client = OpenAI(
            api_key=self.config["api_key"],
            base_url="https://api.deepseek.com"
        )

    def setup_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # === 左侧边栏 (设置与说明) ===
        self.sidebar = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        
        ctk.CTkLabel(self.sidebar, text=APP_NAME, font=("Arial", 18, "bold")).pack(pady=20)
        
        # 模型选择
        ctk.CTkLabel(self.sidebar, text="模型选择:").pack(padx=10, anchor="w")
        self.model_var = ctk.StringVar(value=self.config["model"])
        self.model_combo = ctk.CTkComboBox(self.sidebar, values=["deepseek-chat", "deepseek-reasoner"], variable=self.model_var, command=self.update_settings)
        self.model_combo.pack(padx=10, pady=5, fill="x")
        
        # 联网搜索开关
        self.search_var = ctk.BooleanVar(value=self.config["use_search"])
        self.search_switch = ctk.CTkSwitch(self.sidebar, text="联网搜索", variable=self.search_var, command=self.update_settings)
        self.search_switch.pack(padx=10, pady=15, anchor="w")

        # API Key 设置
        ctk.CTkLabel(self.sidebar, text="API Key:").pack(padx=10, anchor="w")
        self.entry_key = ctk.CTkEntry(self.sidebar, show="*")
        self.entry_key.insert(0, self.config["api_key"])
        self.entry_key.pack(padx=10, pady=5, fill="x")
        ctk.CTkButton(self.sidebar, text="保存 Key", command=self.save_key).pack(padx=10, pady=5)

        # 功能按钮
        ctk.CTkButton(self.sidebar, text="🧹 新对话", fg_color="gray", command=self.clear_chat).pack(padx=10, pady=(20, 5), fill="x")
        ctk.CTkButton(self.sidebar, text="ℹ️ 关于/说明", command=self.show_about).pack(padx=10, pady=5, fill="x")

        # 底部开发者信息
        ctk.CTkLabel(self.sidebar, text=DEV_INFO, font=("Arial", 10), text_color="gray").pack(side="bottom", pady=20)

        # === 右侧主区域 ===
        self.main_area = ctk.CTkFrame(self, fg_color="transparent")
        self.main_area.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        self.main_area.grid_rowconfigure(0, weight=1)
        self.main_area.grid_columnconfigure(0, weight=1)

        # 1. 聊天显示区 (使用 Textbox 模拟流式输出)
        self.chat_display = ctk.CTkTextbox(self.main_area, font=("Microsoft YaHei UI", 14), wrap="word")
        self.chat_display.grid(row=0, column=0, sticky="nsew", pady=(0, 10))
        self.chat_display.insert("0.0", "👋 你好！我是 DeepSeek 智能助手。\n请在设置中输入 API Key 开始对话。\n支持 PDF/Word 读取和联网搜索。\n\n")
        self.chat_display.configure(state="disabled")

        # 2. 思考过程显示区 (仿网页版，默认隐藏，有深度思考时显示)
        self.thought_frame = ctk.CTkFrame(self.main_area, fg_color=("gray85", "gray20"), height=0)
        self.thought_display = ctk.CTkTextbox(self.thought_frame, font=("Arial", 12), text_color="gray", height=100, wrap="word")
        self.thought_display.pack(fill="both", expand=True, padx=5, pady=5)
        self.thought_label = ctk.CTkLabel(self.thought_frame, text="🧠 深度思考中...", font=("Arial", 12, "bold"), text_color="gray")
        self.thought_label.pack(anchor="w", padx=5)
        # 初始不布局，需要时 grid

        # 3. 输入区
        input_frame = ctk.CTkFrame(self.main_area, fg_color="transparent")
        input_frame.grid(row=2, column=0, sticky="ew")
        input_frame.grid_columnconfigure(1, weight=1)

        self.btn_attach = ctk.CTkButton(input_frame, text="📎", width=40, command=self.upload_file)
        self.btn_attach.grid(row=0, column=0, padx=(0, 5), sticky="s")

        self.entry_msg = ctk.CTkTextbox(input_frame, height=60, font=("Microsoft YaHei UI", 14))
        self.entry_msg.grid(row=0, column=1, sticky="ew")
        # 绑定回车发送
        self.entry_msg.bind("<Shift-Return>", lambda e: "break") # 换行
        self.entry_msg.bind("<Return>", self.on_enter_press)

        self.btn_send = ctk.CTkButton(input_frame, text="发送", width=80, command=self.send_message)
        self.btn_send.grid(row=0, column=2, padx=(5, 0), sticky="s")
        
        self.lbl_file_status = ctk.CTkLabel(input_frame, text="", text_color="green", font=("Arial", 10))
        self.lbl_file_status.grid(row=1, column=1, sticky="w")

    # --- 逻辑处理 ---

    def update_settings(self, choice=None):
        self.config["model"] = self.model_var.get()
        self.config["use_search"] = self.search_var.get()
        self.save_config()

    def save_key(self):
        key = self.entry_key.get().strip()
        if not key:
            messagebox.showerror("错误", "API Key 不能为空")
            return
        self.config["api_key"] = key
        self.save_config()
        self.init_client()
        messagebox.showinfo("成功", "API Key 已保存")

    def upload_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Documents", "*.pdf *.docx *.txt")])
        if not filepath: return
        
        try:
            text = ""
            ext = os.path.splitext(filepath)[1].lower()
            if ext == ".pdf":
                reader = pypdf.PdfReader(filepath)
                for page in reader.pages: text += page.extract_text() + "\n"
            elif ext == ".docx":
                doc = Document(filepath)
                text = "\n".join([p.text for p in doc.paragraphs])
            else:
                with open(filepath, "r", encoding="utf-8") as f: text = f.read()
            
            if not text.strip(): raise ValueError("文件内容为空")
            
            self.attached_content = f"【附件内容】：\n{text[:10000]}\n(内容过长已截断)\n----------------\n"
            self.lbl_file_status.configure(text=f"已加载附件: {os.path.basename(filepath)}")
        except Exception as e:
            messagebox.showerror("读取失败", str(e))

    def perform_web_search(self, query):
        """ 使用 DuckDuckGo 进行搜索 """
        try:
            with DDGS() as ddgs:
                results = list(ddgs.text(query, max_results=3))
                if results:
                    context = "\n".join([f"- {r['title']}: {r['body']}" for r in results])
                    return f"【联网搜索结果】：\n{context}\n----------------\n"
        except Exception as e:
            print(f"搜索失败: {e}")
        return ""

    def on_enter_press(self, event):
        if not event.state & 0x0001: # 如果没有按 Shift
            self.send_message()
            return "break"

    def clear_chat(self):
        self.chat_history = []
        self.chat_display.configure(state="normal")
        self.chat_display.delete("0.0", "end")
        self.chat_display.configure(state="disabled")
        self.thought_frame.grid_forget()
        self.attached_content = ""
        self.lbl_file_status.configure(text="")

    def show_about(self):
        info = """【DeepSeek Pro 桌面版】
版本：v1.0.0
开发者：Yu Jinquan

【功能说明】
1. 深度思考：选择 'deepseek-reasoner' 模型即可触发，展示思维链。
2. 联网搜索：勾选开启，AI 会先搜索相关信息再回答（会增加等待时间）。
3. 附件上传：支持 PDF/Word/Txt，自动提取文字作为上下文。
4. 连续对话：软件会自动记忆上下文。

【注意】
API Key 必须开通 DeepSeek 官方服务。
联网搜索使用 DuckDuckGo 接口，需确保网络畅通。
"""
        messagebox.showinfo("关于", info)

    def append_chat(self, role, text, tag=None):
        self.chat_display.configure(state="normal")
        timestamp = datetime.now().strftime("%H:%M")
        header = "🧑 我" if role == "user" else "🤖 DeepSeek"
        
        self.chat_display.insert("end", f"\n{header} ({timestamp}):\n", "header")
        self.chat_display.insert("end", f"{text}\n", tag if tag else "body")
        self.chat_display.see("end")
        self.chat_display.configure(state="disabled")

    def send_message(self):
        user_input = self.entry_msg.get("0.0", "end").strip()
        if not user_input: return
        if not self.client:
            messagebox.showerror("错误", "请先配置 API Key")
            return

        # 1. UI更新
        self.entry_msg.delete("0.0", "end")
        self.append_chat("user", user_input)
        self.btn_send.configure(state="disabled", text="生成中...")
        
        # 隐藏旧的思考框
        self.thought_frame.grid_forget()
        self.thought_display.configure(state="normal")
        self.thought_display.delete("0.0", "end")
        self.thought_display.configure(state="disabled")

        # 2. 开启线程处理
        threading.Thread(target=self.process_generation, args=(user_input,), daemon=True).start()

    def process_generation(self, user_input):
        full_context = ""
        
        # A. 处理附件
        if self.attached_content:
            full_context += self.attached_content
            self.attached_content = "" # 消耗掉
            self.after(0, lambda: self.lbl_file_status.configure(text=""))

        # B. 处理联网搜索
        if self.search_var.get():
            self.after(0, lambda: self.chat_display.configure(state="normal"))
            self.after(0, lambda: self.chat_display.insert("end", "🔍 正在联网搜索...\n"))
            self.after(0, lambda: self.chat_display.configure(state="disabled"))
            
            search_res = self.perform_web_search(user_input)
            if search_res:
                full_context += search_res

        # C. 组装消息
        final_prompt = full_context + user_input
        self.chat_history.append({"role": "user", "content": final_prompt})

        try:
            # D. 调用 API (流式)
            response = self.client.chat.completions.create(
                model=self.config["model"],
                messages=[
                    {"role": "system", "content": self.config["system_prompt"]},
                    *self.chat_history
                ],
                stream=True
            )

            # 准备UI接收流
            is_reasoning = False
            ai_content = ""
            ai_reasoning = ""
            
            self.after(0, lambda: self.chat_display.configure(state="normal"))
            self.after(0, lambda: self.chat_display.insert("end", f"\n🤖 DeepSeek ({datetime.now().strftime('%H:%M')}):\n", "header"))
            
            for chunk in response:
                delta = chunk.choices[0].delta
                
                # 1. 处理深度思考 (Reasoning)
                if hasattr(delta, 'reasoning_content') and delta.reasoning_content:
                    if not is_reasoning:
                        is_reasoning = True
                        # 显示思考框
                        self.after(0, lambda: self.thought_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=5))
                    
                    content = delta.reasoning_content
                    ai_reasoning += content
                    self.after(0, self.update_textbox, self.thought_display, content)

                # 2. 处理正文
                if hasattr(delta, 'content') and delta.content:
                    content = delta.content
                    ai_content += content
                    self.after(0, self.update_textbox, self.chat_display, content)

            # 记录历史 (去掉附件和搜索的大段文本，只存核心，或者存全部取决于Token限制)
            # 这里为了省钱，建议只存用户原始问题，或者精简版
            # 但为了连续对话准确，暂存全部。
            self.chat_history.append({"role": "assistant", "content": ai_content})

        except Exception as e:
            self.after(0, lambda: messagebox.showerror("API 错误", str(e)))
        
        finally:
            self.after(0, self.finish_generation)

    def update_textbox(self, widget, text):
        widget.configure(state="normal")
        widget.insert("end", text)
        widget.see("end")
        widget.configure(state="disabled")

    def finish_generation(self):
        self.btn_send.configure(state="normal", text="发送")
        self.chat_display.configure(state="normal")
        self.chat_display.insert("end", "\n------------------------------------------------\n")
        self.chat_display.configure(state="disabled")

if __name__ == "__main__":
    app = DeepSeekApp()
    app.mainloop()
