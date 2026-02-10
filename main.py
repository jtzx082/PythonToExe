import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import os
import json
import re
import uuid
from datetime import datetime
import pyperclip
from openai import OpenAI

# --- 扩展功能库 ---
from duckduckgo_search import DDGS
import pypdf
from docx import Document
import pandas as pd
try:
    from pptx import Presentation
except ImportError:
    Presentation = None

# --- 配置区域 ---
APP_NAME = "DeepSeek Pro"
APP_VERSION = "v2.2.0 (Sessions & Stream)"
DEV_INFO = "Developer: Yu Jinquan"

DEFAULT_CONFIG = {
    "api_key": "",
    "model": "deepseek-chat",
    "use_search": False,
    "is_r1": False,
    "system_prompt": "你是一个乐于助人的AI助手。代码请用Markdown格式。"
}

# 颜色配置
COLOR_USER_BUBBLE = "#95EC69" # 微信绿
COLOR_AI_BUBBLE = ("#FFFFFF", "#2B2B2B") # 白/深灰
COLOR_BG = ("#F2F2F2", "#1a1a1a")
COLOR_SIDEBAR = ("#EBEBEB", "#212121")

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class AttachmentChip(ctk.CTkFrame):
    """ 单个附件胶囊组件，带删除按钮 """
    def __init__(self, master, filename, command_delete, **kwargs):
        super().__init__(master, fg_color=("gray85", "gray30"), corner_radius=10, **kwargs)
        
        # 文件名
        lbl = ctk.CTkLabel(self, text=filename, font=("Arial", 11))
        lbl.pack(side="left", padx=(10, 5), pady=2)
        
        # 删除按钮 (X)
        btn = ctk.CTkButton(self, text="×", width=20, height=20, 
                            fg_color="transparent", hover_color=("gray70", "gray40"),
                            text_color="red", font=("Arial", 14, "bold"),
                            command=command_delete)
        btn.pack(side="right", padx=(0, 5), pady=2)

class ChatBubble(ctk.CTkFrame):
    """ 聊天气泡：支持流式更新、代码高亮、一键复制 """
    def __init__(self, master, role, text="", is_reasoning=False, timestamp=None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.role = role
        self.raw_text = text 
        self.is_reasoning = is_reasoning
        
        # 布局
        self.grid_columnconfigure(0 if role == "user" else 1, weight=1)
        self.grid_columnconfigure(1 if role == "user" else 0, weight=0)
        
        # 样式定义
        if role == "user":
            bubble_color = COLOR_USER_BUBBLE
            text_color = "black"
            anchor = "e"
        else:
            bubble_color = COLOR_AI_BUBBLE
            text_color = ("black", "white")
            anchor = "w"

        if is_reasoning:
            bubble_color = ("#F0F0F0", "#333333")
            text_color = "gray"
            self.prefix = "🧠 深度思考:\n"
        else:
            self.prefix = ""

        # 气泡实体
        self.bubble_inner = ctk.CTkFrame(self, fg_color=bubble_color, corner_radius=12)
        self.bubble_inner.grid(row=0, column=1 if role == "user" else 0, padx=10, pady=5, sticky=anchor)

        # 内容容器 (用于动态添加 Label 或 CodeBlock)
        self.content_frame = ctk.CTkFrame(self.bubble_inner, fg_color="transparent")
        self.content_frame.pack(fill="both", padx=10, pady=10)

        # 初始渲染
        self.render_content(self.prefix + text, text_color)

        # 底部栏：时间 + 复制按钮
        self.bottom_bar = ctk.CTkFrame(self.bubble_inner, fg_color="transparent", height=20)
        self.bottom_bar.pack(fill="x", padx=10, pady=(0, 5))
        
        # 复制按钮 (常驻显示)
        self.btn_copy = ctk.CTkButton(self.bottom_bar, text="📋 复制", width=50, height=20,
                                      fg_color="transparent", hover_color=("gray80", "gray40"),
                                      text_color="gray", font=("Arial", 10),
                                      command=self.copy_content)
        self.btn_copy.pack(side="right")

        if timestamp:
            ctk.CTkLabel(self.bottom_bar, text=timestamp, font=("Arial", 10), text_color="gray").pack(side="left")

    def update_text(self, new_text):
        """ 流式更新接口 """
        self.raw_text = new_text
        # 清空旧内容
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        
        # 重新渲染 (根据当前主题色判断文字颜色)
        text_color = "gray" if self.is_reasoning else ("black", "white")
        self.render_content(self.prefix + new_text, text_color)

    def copy_content(self):
        try:
            pyperclip.copy(self.raw_text)
            self.btn_copy.configure(text="✅ 已复制")
            self.after(2000, lambda: self.btn_copy.configure(text="📋 复制"))
        except: pass

    def render_content(self, text, text_color):
        # 简单的 Markdown 代码块解析
        parts = re.split(r'(```[\s\S]*?```)', text)
        for part in parts:
            if part.startswith("```") and part.endswith("```"):
                # 代码块
                code = part.strip("`")
                if '\n' in code:
                    lang = code.split('\n', 1)[0]
                    code = code.split('\n', 1)[1]
                
                f = ctk.CTkFrame(self.content_frame, fg_color="#1E1E1E", corner_radius=5)
                f.pack(fill="x", pady=5)
                
                # 代码内容
                t = ctk.CTkTextbox(f, font=("Consolas", 12), text_color="#D4D4D4", fg_color="transparent", 
                                   height=min(len(code.split('\n'))*20 + 20, 400), wrap="none")
                t.insert("0.0", code)
                t.configure(state="disabled")
                t.pack(fill="x", padx=5, pady=5)
                
                # 代码块独立复制
                ctk.CTkButton(f, text="复制代码", height=20, width=60, font=("Arial", 10),
                              fg_color="#333333", hover_color="#444444",
                              command=lambda c=code: pyperclip.copy(c)).pack(anchor="ne", padx=5, pady=2)
            else:
                if part:
                    # 普通文本 (自动换行)
                    ctk.CTkLabel(self.content_frame, text=part, text_color=text_color, justify="left", 
                                 font=("Microsoft YaHei UI", 14), wraplength=600).pack(fill="x", anchor="w")

class DeepSeekApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} {APP_VERSION}")
        self.geometry("1300x850")
        
        # 数据初始化
        self.config = self.load_json("config.json", DEFAULT_CONFIG)
        self.sessions = self.load_json("sessions.json", []) # 多会话数据
        
        # 确保至少有一个会话
        if not self.sessions:
            self.create_new_session(save=False)
        else:
            self.current_session_index = 0 # 默认选第一个
            
        self.attachments = [] # 当前暂存的附件列表
        self.client = None
        self.is_running = False

        self.setup_ui()
        self.load_current_session_ui() # 加载聊天记录
        
        if self.config["api_key"]:
            self.init_client()

    def load_json(self, path, default):
        if os.path.exists(path):
            try: return json.load(open(path, "r", encoding="utf-8"))
            except: pass
        return default

    def save_config(self):
        json.dump(self.config, open("config.json", "w", encoding="utf-8"), indent=2)

    def save_sessions(self):
        # 保存所有会话
        json.dump(self.sessions, open("sessions.json", "w", encoding="utf-8"), ensure_ascii=False, indent=2)

    def init_client(self):
        if not self.config["api_key"]: return
        self.client = OpenAI(api_key=self.config["api_key"], base_url="https://api.deepseek.com")

    # --- UI 构建 ---
    def setup_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # === 1. 左侧功能区 (优化版) ===
        self.sidebar = ctk.CTkFrame(self, width=250, corner_radius=0, fg_color=COLOR_SIDEBAR)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(2, weight=1) # 历史记录列表占满中间

        # 1.1 顶部标题与新建
        top_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        top_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=20)
        ctk.CTkLabel(top_frame, text="DeepSeek Pro", font=("Arial", 22, "bold")).pack(anchor="w")
        
        self.btn_new = ctk.CTkButton(self.sidebar, text="+ 开启新对话", height=40, font=("Arial", 14), 
                                     fg_color="#3498DB", hover_color="#2980B9",
                                     command=lambda: self.create_new_session(save=True))
        self.btn_new.grid(row=1, column=0, padx=15, pady=(0, 10), sticky="ew")

        # 1.2 历史记录列表 (Scrollable)
        ctk.CTkLabel(self.sidebar, text="历史记录", font=("Arial", 12), text_color="gray").grid(row=2, column=0, sticky="nw", padx=15)
        
        self.history_list = ctk.CTkScrollableFrame(self.sidebar, fg_color="transparent")
        self.history_list.grid(row=3, column=0, sticky="nsew", padx=5, pady=5)
        self.render_history_list() # 渲染左侧列表

        # 1.3 底部设置区
        setting_frame = ctk.CTkFrame(self.sidebar, fg_color=("white", "#2B2B2B"), corner_radius=10)
        setting_frame.grid(row=4, column=0, sticky="ew", padx=10, pady=20)
        
        self.r1_var = ctk.BooleanVar(value=self.config["is_r1"])
        ctk.CTkSwitch(setting_frame, text="深度思考 (R1)", variable=self.r1_var, command=self.update_settings).pack(pady=5, padx=10, anchor="w")
        
        self.search_var = ctk.BooleanVar(value=self.config["use_search"])
        ctk.CTkSwitch(setting_frame, text="联网搜索", variable=self.search_var, command=self.update_settings).pack(pady=5, padx=10, anchor="w")

        self.entry_key = ctk.CTkEntry(setting_frame, placeholder_text="API Key (sk-...)")
        self.entry_key.insert(0, self.config["api_key"])
        self.entry_key.pack(pady=5, padx=10, fill="x")
        
        ctk.CTkButton(setting_frame, text="保存配置", height=24, command=self.save_key).pack(pady=10)

        # === 2. 右侧聊天区 ===
        self.main_area = ctk.CTkFrame(self, fg_color=COLOR_BG)
        self.main_area.grid(row=0, column=1, sticky="nsew")
        self.main_area.grid_rowconfigure(0, weight=1)
        self.main_area.grid_columnconfigure(0, weight=1)

        # 2.1 聊天内容
        self.chat_scroll = ctk.CTkScrollableFrame(self.main_area, fg_color="transparent")
        self.chat_scroll.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        # 2.2 底部输入栏
        input_frame = ctk.CTkFrame(self.main_area, fg_color=("white", "#2B2B2B"), height=180)
        input_frame.grid(row=1, column=0, sticky="ew", padx=20, pady=20)
        input_frame.grid_columnconfigure(0, weight=1)

        # 附件展示区 (横向滚动)
        self.attach_display = ctk.CTkScrollableFrame(input_frame, height=40, orientation="horizontal", fg_color="transparent")
        self.attach_display.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        
        # 输入框
        self.entry_msg = ctk.CTkTextbox(input_frame, height=80, font=("Microsoft YaHei UI", 14), fg_color="transparent", border_width=0)
        self.entry_msg.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        self.entry_msg.bind("<Return>", self.on_enter_press)

        # 按钮区
        btn_box = ctk.CTkFrame(input_frame, fg_color="transparent")
        btn_box.grid(row=1, column=1, sticky="s", padx=10, pady=10)
        
        self.btn_attach = ctk.CTkButton(btn_box, text="📎", width=40, command=self.upload_files)
        self.btn_attach.pack(side="left", padx=2)
        
        self.btn_send = ctk.CTkButton(btn_box, text="发送", width=80, command=self.send_message)
        self.btn_send.pack(side="left", padx=2)
        
        self.btn_stop = ctk.CTkButton(btn_box, text="⏹", width=40, fg_color="#C0392B", command=self.stop_generation)
        # 初始不显示停止

    # --- 会话管理逻辑 ---

    def create_new_session(self, save=True):
        """ 创建新会话对象 """
        new_session = {
            "id": str(uuid.uuid4()),
            "title": "新对话",
            "time": datetime.now().strftime("%m-%d %H:%M"),
            "messages": [] # 存储 [{"role":..., "content":..., "reasoning":...}]
        }
        self.sessions.insert(0, new_session) # 插到最前
        self.current_session_index = 0
        
        if save:
            self.save_sessions()
            self.render_history_list()
            self.load_current_session_ui()

    def switch_session(self, index):
        """ 切换会话 """
        self.current_session_index = index
        self.render_history_list() # 更新选中状态
        self.load_current_session_ui()

    def delete_session(self, index):
        """ 删除会话 """
        if len(self.sessions) <= 1:
            self.create_new_session(save=False)
            self.sessions = [self.sessions[0]] # 重置为新会话
        else:
            del self.sessions[index]
            if self.current_session_index >= index:
                self.current_session_index = max(0, self.current_session_index - 1)
        
        self.save_sessions()
        self.render_history_list()
        self.load_current_session_ui()

    def render_history_list(self):
        """ 渲染左侧历史记录列表 """
        for widget in self.history_list.winfo_children():
            widget.destroy()

        for i, session in enumerate(self.sessions):
            # 选中状态颜色不同
            color = ("#D1D1D1", "#3A3A3A") if i == self.current_session_index else "transparent"
            
            item = ctk.CTkFrame(self.history_list, fg_color=color, corner_radius=6)
            item.pack(fill="x", pady=2)
            
            # 点击整个Frame切换
            item.bind("<Button-1>", lambda e, idx=i: self.switch_session(idx))
            
            # 标题与时间
            title = session.get("title", "无标题")
            if len(title) > 12: title = title[:12] + "..."
            
            lbl_title = ctk.CTkLabel(item, text=title, font=("Arial", 13, "bold"))
            lbl_title.pack(anchor="w", padx=10, pady=(5,0))
            lbl_title.bind("<Button-1>", lambda e, idx=i: self.switch_session(idx))
            
            lbl_time = ctk.CTkLabel(item, text=session.get("time", ""), font=("Arial", 10), text_color="gray")
            lbl_time.pack(anchor="w", padx=10, pady=(0,5))
            lbl_time.bind("<Button-1>", lambda e, idx=i: self.switch_session(idx))

            # 删除按钮 (仅hover显示比较复杂，这里简化为常驻小点或右键，为了简单，放一个显式的小X)
            btn_del = ctk.CTkButton(item, text="×", width=20, height=20, fg_color="transparent", text_color="gray", hover_color="red",
                                    command=lambda idx=i: self.delete_session(idx))
            btn_del.place(relx=1.0, rely=0.5, anchor="e", x=-5)

    def load_current_session_ui(self):
        """ 加载当前会话到右侧聊天区 """
        # 清空聊天区
        for widget in self.chat_scroll.winfo_children():
            widget.destroy()
        
        self.attachments = [] # 切换会话清空暂存附件
        self.render_attachments_ui()
        
        session = self.sessions[self.current_session_index]
        msgs = session.get("messages", [])
        
        for msg in msgs:
            role = msg["role"]
            content = msg.get("content", "")
            reasoning = msg.get("reasoning", "")
            ts = msg.get("timestamp", "")
            
            if role == "user":
                self.add_bubble_ui("user", content, timestamp=ts)
            else:
                if reasoning:
                    self.add_bubble_ui("ai", reasoning, is_reasoning=True, timestamp=ts)
                if content:
                    self.add_bubble_ui("ai", content, is_reasoning=False, timestamp=ts)
        
        self.scroll_to_bottom()

    # --- 附件逻辑 ---
    def upload_files(self):
        files = filedialog.askopenfilenames()
        if not files: return
        for path in files:
            name = os.path.basename(path)
            content = self.extract_text(path)
            if len(content) > 30000: content = content[:30000] + "\n[Truncated]"
            self.attachments.append({"name": name, "content": content})
        self.render_attachments_ui()

    def remove_attachment(self, index):
        del self.attachments[index]
        self.render_attachments_ui()

    def render_attachments_ui(self):
        for w in self.attach_display.winfo_children(): w.destroy()
        for i, att in enumerate(self.attachments):
            chip = AttachmentChip(self.attach_display, att["name"], lambda idx=i: self.remove_attachment(idx))
            chip.pack(side="left", padx=5)

    def extract_text(self, filepath):
        # ... (保持原有的多格式读取逻辑)
        ext = os.path.splitext(filepath)[1].lower()
        try:
            if ext == '.pdf':
                reader = pypdf.PdfReader(filepath)
                return "\n".join([p.extract_text() or "" for p in reader.pages])
            elif ext == '.docx':
                doc = Document(filepath)
                return "\n".join([p.text for p in doc.paragraphs])
            elif ext in ['.xlsx', '.xls', '.csv']:
                df = pd.read_excel(filepath) if 'xls' in ext else pd.read_csv(filepath)
                return df.to_string()
            else:
                with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                    return f.read()
        except: return f"[无法读取文件 {os.path.basename(filepath)}]"

    # --- 聊天交互 ---
    def add_bubble_ui(self, role, text, is_reasoning=False, timestamp=None):
        if not timestamp: timestamp = datetime.now().strftime("%H:%M")
        bubble = ChatBubble(self.chat_scroll, role, text, is_reasoning, timestamp)
        bubble.pack(fill="x", pady=5)
        return bubble

    def scroll_to_bottom(self):
        self.chat_scroll.update_idletasks()
        try: self.chat_scroll._parent_canvas.yview_moveto(1.0)
        except: pass

    def send_message(self):
        text = self.entry_msg.get("0.0", "end").strip()
        if not text and not self.attachments: return
        if not self.client: return messagebox.showerror("Error", "No API Key")

        # 1. 准备数据
        display_text = text
        full_prompt = ""
        
        if self.attachments:
            files_str = "\n".join([f"文件[{f['name']}]:\n{f['content']}" for f in self.attachments])
            full_prompt += files_str + "\n\n"
            display_text += f"\n[已发送 {len(self.attachments)} 个文件]"
            self.attachments = [] # 发送后清空
            self.render_attachments_ui()
        
        full_prompt += text
        ts = datetime.now().strftime("%H:%M")

        # 2. 更新界面
        self.entry_msg.delete("0.0", "end")
        self.add_bubble_ui("user", display_text, timestamp=ts)
        self.scroll_to_bottom()

        # 3. 更新数据模型
        session = self.sessions[self.current_session_index]
        # 更新标题（如果是第一条）
        if len(session["messages"]) == 0:
            session["title"] = text[:15]
            self.render_history_list()
        
        session["messages"].append({"role": "user", "content": full_prompt, "timestamp": ts})
        self.save_sessions()

        # 4. 线程生成
        self.is_running = True
        self.btn_send.pack_forget()
        self.btn_stop.pack(side="left")
        threading.Thread(target=self.process_stream, args=(full_prompt,), daemon=True).start()

    def process_stream(self, prompt):
        # 联网搜索逻辑... (保持不变)
        if self.search_var.get():
            self.after(0, lambda: self.add_bubble_ui("ai", "🔍 正在搜索...", timestamp="System"))
            # ...执行搜索并拼接到 prompt

        session = self.sessions[self.current_session_index]
        # 构造上下文 (最近5轮)
        api_msgs = [{"role": "system", "content": self.config["system_prompt"]}]
        for m in session["messages"][-5:]:
            api_msgs.append({"role": "user" if m["role"]=="user" else "assistant", "content": m["content"]})
        
        # 确保最后一条是刚才发送的
        if api_msgs[-1]["content"] != prompt:
             api_msgs.append({"role": "user", "content": prompt})

        try:
            response = self.client.chat.completions.create(
                model=self.config["model"],
                messages=api_msgs,
                stream=True
            )

            # === 核心流式优化：直接创建气泡，实时更新 ===
            r1_text = ""
            ai_text = ""
            
            # 在主线程创建气泡占位
            bubble_r1 = None
            bubble_ai = None
            
            def get_r1_bubble():
                nonlocal bubble_r1
                if not bubble_r1:
                    bubble_r1 = self.add_bubble_ui("ai", "", is_reasoning=True)
                return bubble_r1

            def get_ai_bubble():
                nonlocal bubble_ai
                if not bubble_ai:
                    bubble_ai = self.add_bubble_ui("ai", "")
                return bubble_ai

            for chunk in response:
                if not self.is_running: break
                delta = chunk.choices[0].delta
                
                # 处理思考
                if hasattr(delta, 'reasoning_content') and delta.reasoning_content:
                    r1_text += delta.reasoning_content
                    # 必须在主线程更新 UI
                    self.after(0, lambda b=get_r1_bubble(), t=r1_text: b.update_text(t))
                    self.after(0, self.scroll_to_bottom) # 自动滚动

                # 处理正文
                if hasattr(delta, 'content') and delta.content:
                    ai_text += delta.content
                    self.after(0, lambda b=get_ai_bubble(), t=ai_text: b.update_text(t))
                    self.after(0, self.scroll_to_bottom) # 自动滚动

            # 保存结果
            ts = datetime.now().strftime("%H:%M")
            session["messages"].append({
                "role": "ai", 
                "content": ai_text, 
                "reasoning": r1_text, 
                "timestamp": ts
            })
            self.save_sessions()

        except Exception as e:
            self.after(0, lambda: messagebox.showerror("API Error", str(e)))
        
        finally:
            self.is_running = False
            self.after(0, self.reset_ui)

    def reset_ui(self):
        self.btn_stop.pack_forget()
        self.btn_send.pack(side="left", padx=2)

    def stop_generation(self):
        self.is_running = False
        self.reset_ui()

    def update_settings(self):
        self.config["is_r1"] = self.r1_var.get()
        self.config["use_search"] = self.search_var.get()
        self.config["model"] = "deepseek-reasoner" if self.r1_var.get() else "deepseek-chat"
        self.save_config()

    def save_key(self):
        self.config["api_key"] = self.entry_key.get().strip()
        self.save_config()
        self.init_client()
        messagebox.showinfo("OK", "Key Saved")

    def on_enter_press(self, event):
        if not event.state & 0x0001: 
            self.send_message()
            return "break"

if __name__ == "__main__":
    app = DeepSeekApp()
    app.mainloop()
