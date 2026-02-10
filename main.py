import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import os
import json
import re
from datetime import datetime
import pyperclip  #用于剪贴板
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
APP_VERSION = "v2.0.0 (Chat Bubble Edition)"
DEV_INFO = "开发者：Yu Jinquan | 核心：DeepSeek-V3/R1"

DEFAULT_CONFIG = {
    "api_key": "",
    "model": "deepseek-chat",
    "use_search": False,
    "is_r1": False, # 是否开启深度思考
    "system_prompt": "你是一个乐于助人的AI助手。输出代码时请使用Markdown格式。"
}

# 颜色配置 (仿微信/现代风格)
COLOR_USER_BUBBLE = "#95EC69"  # 微信绿
COLOR_USER_TEXT = "#000000"
COLOR_AI_BUBBLE = "#FFFFFF"    # 亮色模式白
COLOR_AI_BUBBLE_DARK = "#2B2B2B" # 深色模式灰
COLOR_CODE_BG = "#1E1E1E"      # 代码块背景
COLOR_BG = ("#F2F2F2", "#1a1a1a") # 整体背景

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class ChatBubble(ctk.CTkFrame):
    """ 自定义聊天气泡组件 """
    def __init__(self, master, role, text, is_reasoning=False, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.role = role
        self.text_content = text
        self.is_reasoning = is_reasoning
        
        # 布局配置
        self.grid_columnconfigure(0 if role == "user" else 1, weight=1)
        self.grid_columnconfigure(1 if role == "user" else 0, weight=0)
        
        # 气泡颜色
        if role == "user":
            bubble_color = COLOR_USER_BUBBLE
            text_color = COLOR_USER_TEXT
            anchor = "e"
            justify = "left"
        else:
            bubble_color = (COLOR_AI_BUBBLE, COLOR_AI_BUBBLE_DARK)
            text_color = ("black", "white")
            anchor = "w"
            justify = "left"

        if is_reasoning:
            bubble_color = ("#F0F0F0", "#333333")
            text_color = "gray"
            text = f"🧠 深度思考过程:\n{text}"

        # 内容容器 (圆角矩形)
        self.bubble_frame = ctk.CTkFrame(self, fg_color=bubble_color, corner_radius=15)
        self.bubble_frame.grid(row=0, column=1 if role == "user" else 0, padx=10, pady=5, sticky=anchor)

        # 文本/代码渲染逻辑
        self.render_content(self.bubble_frame, text, text_color)

        # 复制按钮 (悬浮或位于底部)
        self.btn_copy = ctk.CTkButton(self.bubble_frame, text="📄", width=30, height=20, 
                                      fg_color="transparent", text_color="gray",
                                      command=self.copy_text)
        self.btn_copy.pack(anchor="e", padx=5, pady=(0, 5))

    def render_content(self, parent, text, text_color):
        """ 简单的 Markdown 代码块解析与渲染 """
        # 正则分割代码块 ```code```
        parts = re.split(r'(```[\s\S]*?```)', text)
        
        for part in parts:
            if part.startswith("```") and part.endswith("```"):
                # 处理代码块
                code_content = part.strip("`")
                # 尝试去除第一行语言标识 (如 python)
                first_newline = code_content.find('\n')
                if first_newline != -1:
                    lang = code_content[:first_newline].strip()
                    code_body = code_content[first_newline+1:]
                else:
                    code_body = code_content
                
                # 代码容器
                code_frame = ctk.CTkFrame(parent, fg_color=COLOR_CODE_BG, corner_radius=5)
                code_frame.pack(fill="x", padx=10, pady=5)
                
                # 代码文本
                code_font = ctk.CTkFont(family="Consolas", size=12)
                code_label = ctk.CTkTextbox(code_frame, font=code_font, text_color="#D4D4D4", 
                                            fg_color="transparent", height=len(code_body.split('\n'))*20 + 20, wrap="none")
                code_label.insert("0.0", code_body)
                code_label.configure(state="disabled")
                code_label.pack(fill="x", padx=5, pady=5)
                
                # 代码复制按钮
                ctk.CTkButton(code_frame, text="复制代码", height=20, fg_color="#333333", 
                              command=lambda c=code_body: self.copy_to_clip(c)).pack(anchor="ne", padx=5, pady=2)
            else:
                # 普通文本
                if part.strip():
                    lbl = ctk.CTkLabel(parent, text=part, text_color=text_color, justify="left", 
                                       font=("Microsoft YaHei UI", 14), wraplength=600)
                    lbl.pack(fill="x", padx=10, pady=5)

    def copy_text(self):
        self.copy_to_clip(self.text_content)

    def copy_to_clip(self, content):
        pyperclip.copy(content)
        messagebox.showinfo("提示", "内容已复制到剪贴板")


class DeepSeekApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} {APP_VERSION}")
        self.geometry("1200x850")
        
        self.config = self.load_config()
        self.chat_history = [] 
        self.client = None
        self.is_running = False # 控制停止生成
        self.attached_content = "" 
        self.attached_filename = ""

        self.setup_ui()
        
        if self.config["api_key"]:
            self.init_client()

    def load_config(self):
        if os.path.exists("config.json"):
            try: return json.load(open("config.json", "r"))
            except: pass
        return DEFAULT_CONFIG.copy()

    def save_config(self):
        with open("config.json", "w") as f: json.dump(self.config, f)

    def init_client(self):
        if not self.config["api_key"]: return
        self.client = OpenAI(api_key=self.config["api_key"], base_url="https://api.deepseek.com")

    def setup_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # === 1. 左侧边栏 (优化布局) ===
        self.sidebar = ctk.CTkFrame(self, width=220, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(10, weight=1) # 底部占位

        # 标题区
        ctk.CTkLabel(self.sidebar, text="DeepSeek Pro", font=("Arial", 20, "bold")).pack(pady=(30, 10))
        ctk.CTkLabel(self.sidebar, text="全能AI助手", font=("Arial", 12), text_color="gray").pack(pady=(0, 20))

        # 核心设置组
        frame_settings = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        frame_settings.pack(fill="x", padx=10)

        # R1 深度思考开关
        self.r1_var = ctk.BooleanVar(value=self.config.get("is_r1", False))
        switch_r1 = ctk.CTkSwitch(frame_settings, text="深度思考 (R1)", variable=self.r1_var, command=self.update_settings)
        switch_r1.pack(pady=10, anchor="w")

        # 联网搜索开关
        self.search_var = ctk.BooleanVar(value=self.config["use_search"])
        switch_search = ctk.CTkSwitch(frame_settings, text="联网搜索", variable=self.search_var, command=self.update_settings)
        switch_search.pack(pady=10, anchor="w")

        # API Key 区域
        ctk.CTkLabel(self.sidebar, text="API Key 配置:", anchor="w").pack(padx=15, pady=(20, 0), fill="x")
        self.entry_key = ctk.CTkEntry(self.sidebar, show="*", placeholder_text="sk-...")
        self.entry_key.insert(0, self.config["api_key"])
        self.entry_key.pack(padx=15, pady=5, fill="x")
        ctk.CTkButton(self.sidebar, text="保存密钥", height=30, command=self.save_key).pack(padx=15, pady=5, fill="x")

        # 操作按钮组 (底部)
        ctk.CTkButton(self.sidebar, text="🗑️ 清空会话", fg_color="#C0392B", hover_color="#E74C3C", command=self.clear_chat).pack(side="bottom", padx=15, pady=10, fill="x")
        ctk.CTkLabel(self.sidebar, text=DEV_INFO, font=("Arial", 10), text_color="gray50").pack(side="bottom", pady=5)

        # === 2. 右侧主区域 ===
        self.main_area = ctk.CTkFrame(self, fg_color=COLOR_BG)
        self.main_area.grid(row=0, column=1, sticky="nsew")
        self.main_area.grid_rowconfigure(0, weight=1)
        self.main_area.grid_columnconfigure(0, weight=1)

        # 2.1 聊天滚动区 (Bubble Flow)
        self.chat_scroll = ctk.CTkScrollableFrame(self.main_area, fg_color="transparent")
        self.chat_scroll.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        # 欢迎语
        self.add_system_message(f"👋 欢迎使用！\n支持格式：PDF, Word, Excel, PPT, 代码文件等。\n当前模式：{'深度思考(R1)' if self.r1_var.get() else '通用对话(V3)'}")

        # 2.2 底部输入区
        input_container = ctk.CTkFrame(self.main_area, fg_color=("white", "#2B2B2B"), height=150)
        input_container.grid(row=1, column=0, sticky="ew", padx=15, pady=15)
        input_container.grid_columnconfigure(1, weight=1)

        # 附件按钮栏
        attach_frame = ctk.CTkFrame(input_container, fg_color="transparent")
        attach_frame.grid(row=0, column=0, columnspan=3, sticky="ew", padx=10, pady=(5,0))
        
        self.btn_attach = ctk.CTkButton(attach_frame, text="📎 上传附件", width=80, height=24, fg_color="transparent", border_width=1, text_color=("gray20", "gray80"), command=self.upload_file)
        self.btn_attach.pack(side="left")
        
        # 附件状态与删除按钮
        self.lbl_file = ctk.CTkLabel(attach_frame, text="", font=("Arial", 12), text_color="gray")
        self.lbl_file.pack(side="left", padx=5)
        self.btn_del_file = ctk.CTkButton(attach_frame, text="❌", width=20, height=20, fg_color="transparent", text_color="red", command=self.clear_attachment)
        # 初始隐藏删除按钮

        # 文本输入框
        self.entry_msg = ctk.CTkTextbox(input_container, height=80, font=("Microsoft YaHei UI", 14), fg_color="transparent", border_width=0)
        self.entry_msg.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=10, pady=5)
        self.entry_msg.bind("<Return>", self.on_enter_press)

        # 发送与停止按钮
        btn_frame = ctk.CTkFrame(input_container, fg_color="transparent")
        btn_frame.grid(row=1, column=2, sticky="s", padx=10, pady=10)
        
        self.btn_send = ctk.CTkButton(btn_frame, text="发送", width=80, command=self.send_message)
        self.btn_send.pack(side="bottom")
        
        self.btn_stop = ctk.CTkButton(btn_frame, text="⏹️", width=30, fg_color="#C0392B", command=self.stop_generation)
        # 初始隐藏停止按钮

    # --- 逻辑处理 ---

    def update_settings(self):
        self.config["use_search"] = self.search_var.get()
        self.config["is_r1"] = self.r1_var.get()
        self.config["model"] = "deepseek-reasoner" if self.r1_var.get() else "deepseek-chat"
        self.save_config()
        self.add_system_message(f"⚙️ 模式已切换为: {self.config['model']}")

    def save_key(self):
        key = self.entry_key.get().strip()
        if not key: return messagebox.showerror("错误", "Key不能为空")
        self.config["api_key"] = key
        self.save_config()
        self.init_client()
        messagebox.showinfo("成功", "API Key 已保存")

    def upload_file(self):
        # 扩展支持的文件类型
        filetypes = [
            ("文档", "*.pdf *.docx *.pptx *.txt *.md"),
            ("数据", "*.xlsx *.xls *.csv"),
            ("代码", "*.py *.js *.html *.css *.java *.cpp *.c *.json *.xml"),
            ("所有文件", "*.*")
        ]
        filepath = filedialog.askopenfilename(filetypes=filetypes)
        if not filepath: return
        
        try:
            text = self.extract_text(filepath)
            if not text.strip(): raise ValueError("无法提取文本或文件为空")
            
            self.attached_content = f"【附件内容 ({os.path.basename(filepath)})】:\n{text[:15000]}\n(以上是附件内容)\n----------------\n"
            self.attached_filename = os.path.basename(filepath)
            
            # 更新UI
            self.lbl_file.configure(text=f"已添加: {self.attached_filename}")
            self.btn_del_file.pack(side="left", padx=2)
            
        except Exception as e:
            messagebox.showerror("文件读取失败", str(e))

    def extract_text(self, filepath):
        ext = os.path.splitext(filepath)[1].lower()
        text = ""
        if ext == '.pdf':
            reader = pypdf.PdfReader(filepath)
            for p in reader.pages: text += p.extract_text() + "\n"
        elif ext == '.docx':
            doc = Document(filepath)
            text = "\n".join([p.text for p in doc.paragraphs])
        elif ext == '.pptx' and Presentation:
            prs = Presentation(filepath)
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text"): text += shape.text + "\n"
        elif ext in ['.xlsx', '.xls']:
            df = pd.read_excel(filepath)
            text = df.to_string()
        elif ext == '.csv':
            df = pd.read_csv(filepath)
            text = df.to_string()
        else:
            # 尝试纯文本读取
            with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                text = f.read()
        return text

    def clear_attachment(self):
        self.attached_content = ""
        self.attached_filename = ""
        self.lbl_file.configure(text="")
        self.btn_del_file.pack_forget()

    def perform_search(self, query):
        try:
            with DDGS() as ddgs:
                results = list(ddgs.text(query, max_results=3))
                if results:
                    return "【联网搜索结果】:\n" + "\n".join([f"- {r['title']}: {r['body']}" for r in results]) + "\n----------------\n"
        except: pass
        return ""

    def on_enter_press(self, event):
        if not event.state & 0x0001: 
            self.send_message()
            return "break"

    def clear_chat(self):
        self.chat_history = []
        for widget in self.chat_scroll.winfo_children():
            widget.destroy()
        self.add_system_message("🗑️ 会话已清空")

    def add_system_message(self, text):
        lbl = ctk.CTkLabel(self.chat_scroll, text=text, font=("Arial", 10), text_color="gray")
        lbl.pack(pady=5)

    def add_chat_bubble(self, role, text, is_reasoning=False):
        bubble = ChatBubble(self.chat_scroll, role, text, is_reasoning)
        bubble.pack(fill="x", pady=5)
        # 滚动到底部
        self.chat_scroll.update_idletasks()
        self.chat_scroll._parent_canvas.yview_moveto(1.0)
        return bubble

    def stop_generation(self):
        self.is_running = False
        self.btn_stop.pack_forget()
        self.btn_send.configure(state="normal", text="发送")

    def send_message(self):
        text = self.entry_msg.get("0.0", "end").strip()
        if not text: return
        if not self.client: return messagebox.showerror("Error", "请配置API Key")

        # 1. 用户气泡
        self.entry_msg.delete("0.0", "end")
        self.add_chat_bubble("user", text)
        
        # 2. 状态切换
        self.is_running = True
        self.btn_send.configure(state="disabled", text="生成中")
        self.btn_stop.pack(side="bottom", pady=5) # 显示停止按钮

        # 3. 异步处理
        threading.Thread(target=self.process_stream, args=(text,), daemon=True).start()

    def process_stream(self, user_input):
        context_str = ""
        
        # 附件处理
        if self.attached_content:
            context_str += self.attached_content
            self.after(0, self.clear_attachment) # 消耗附件

        # 联网搜索
        if self.search_var.get():
            self.after(0, lambda: self.add_system_message("🔍 正在联网搜索..."))
            search_res = self.perform_search(user_input)
            if search_res: context_str += search_res

        # 构建历史
        full_prompt = context_str + user_input
        self.chat_history.append({"role": "user", "content": full_prompt})

        try:
            response = self.client.chat.completions.create(
                model=self.config["model"],
                messages=[{"role": "system", "content": self.config["system_prompt"]}, *self.chat_history],
                stream=True
            )

            # 占位气泡 (用于流式更新)
            # R1模型有深度思考，需要两个气泡？
            # 策略：先检测是否有 reasoning，如果有，先创建思考气泡，思考完后再创建回答气泡
            
            ai_content = ""
            reasoning_content = ""
            
            # 临时变量控制 UI 创建
            reasoning_bubble = None
            content_bubble = None
            
            for chunk in response:
                if not self.is_running: break # 手动停止
                
                delta = chunk.choices[0].delta
                
                # 1. 处理深度思考
                if hasattr(delta, 'reasoning_content') and delta.reasoning_content:
                    r_text = delta.reasoning_content
                    reasoning_content += r_text
                    
                    if not reasoning_bubble:
                        # 在主线程创建气泡
                        self.after(0, lambda: self.create_bubble_safely("ai", "", True))
                        # 等待气泡创建完成 (简单sleep或者用变量同步，这里简化处理，假设after很快)
                        import time; time.sleep(0.05) 
                        reasoning_bubble = self.chat_scroll.winfo_children()[-1] # 获取最新创建的
                    
                    # 更新气泡内容 (这里为了性能，实际应该优化，但作为demo，我们重新渲染或追加文本会很卡)
                    # 更好的方式：ChatBubble 内部有一个 Textbox，我们往里 insert
                    # 由于 Tkinter 线程安全，必须用 after
                    # 这里为了简化代码逻辑，我们在循环结束后统一渲染漂亮的 Markdown，流式期间只显示纯文本
                    # 改进：我们只在 ChatBubble 里放一个 Label，流式更新 Label 的 text
                    pass # 实际更新逻辑略复杂，见下文修正的 ChatBubble

                # 2. 处理正文
                if hasattr(delta, 'content') and delta.content:
                    c_text = delta.content
                    ai_content += c_text
                    
                    if not content_bubble:
                        self.after(0, lambda: self.create_bubble_safely("ai", ""))
                        import time; time.sleep(0.05)
                        content_bubble = self.chat_scroll.winfo_children()[-1]
                    
                    pass 

            # 流式结束，由于 Tkinter 实时渲染 Markdown 很卡，
            # 我们采取策略：流式过程不展示，或者流式只展示 Loading... 
            # 为了体验，我们这里做一次性渲染（简单方案）或者重构 ChatBubble 支持流式。
            
            # === 修正方案：上述循环只收集文本，实时显示太复杂，我们模拟流式效果或者分段更新 ===
            # 但用户要求"等待输出太慢"。
            # 因此，我们必须实现流式更新 UI。
            # 下面是重写后的流式逻辑：
            
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("API Error", str(e)))
        
        finally:
            # 循环结束后，在界面上创建最终的完美渲染气泡
            # 为了避免逻辑过于复杂，v2.0 采用：收集全量文本 -> 渲染Markdown气泡
            # 如果要实时，需要 ChatBubble 暴露 update_text 方法
            
            self.after(0, lambda: self.finalize_bubbles(reasoning_content, ai_content))
            self.chat_history.append({"role": "assistant", "content": ai_content})
            self.is_running = False
            self.after(0, self.reset_ui_state)

    def create_bubble_safely(self, role, text, is_reasoning=False):
        # 仅用于占位，实际在 finalize 中渲染
        pass 

    def finalize_bubbles(self, reasoning, content):
        if reasoning:
            self.add_chat_bubble("ai", reasoning, is_reasoning=True)
        if content:
            self.add_chat_bubble("ai", content, is_reasoning=False)

    def reset_ui_state(self):
        self.btn_stop.pack_forget()
        self.btn_send.configure(state="normal", text="发送")

# --- 覆盖重写 send_message 中的流式逻辑，使其能实时显示 ---
# 由于 CustomTkinter 的 Label/Textbox 性能，实时 Markdown 渲染不现实。
# 最佳实践：流式输出到纯文本框 -> 结束后销毁纯文本框 -> 替换为渲染好的 Markdown 组件。

    def process_stream(self, user_input):
        context_str = ""
        if self.attached_content:
            context_str += self.attached_content
            self.after(0, self.clear_attachment)
        if self.search_var.get():
            self.after(0, lambda: self.add_system_message("🔍 正在联网搜索..."))
            s = self.perform_search(user_input)
            if s: context_str += s
        
        full_prompt = context_str + user_input
        self.chat_history.append({"role": "user", "content": full_prompt})

        # 创建流式显示的临时容器
        self.current_stream_box = None
        self.current_r1_box = None
        
        def init_stream_ui(is_r1_box=False):
            frame = ctk.CTkFrame(self.chat_scroll, fg_color=("white", "#2B2B2B"))
            frame.pack(fill="x", pady=5, anchor="w", padx=10)
            txt = ctk.CTkTextbox(frame, height=100, font=("Microsoft YaHei UI", 14), fg_color="transparent", wrap="word")
            txt.pack(fill="x", padx=10, pady=10)
            if is_r1_box:
                txt.configure(text_color="gray", font=("Arial", 12))
                txt.insert("0.0", "🧠 深度思考中...\n")
            return frame, txt

        try:
            response = self.client.chat.completions.create(
                model=self.config["model"],
                messages=[{"role": "system", "content": self.config["system_prompt"]}, *self.chat_history],
                stream=True
            )
            
            ai_text = ""
            r1_text = ""
            
            # UI 更新辅助
            def append_text(widget, text):
                widget.insert("end", text)
                widget.see("end")
                # 调整高度
                h = int(widget.index("end-1c").split('.')[0]) * 20 + 20
                widget.configure(height=min(h, 400)) # 限制最大高度

            for chunk in response:
                if not self.is_running: break
                delta = chunk.choices[0].delta
                
                # R1 思考流
                if hasattr(delta, 'reasoning_content') and delta.reasoning_content:
                    if self.current_r1_box is None:
                        # 主线程创建 UI
                        done_evt = threading.Event()
                        def _make():
                            self.r1_frame, self.current_r1_box = init_stream_ui(True)
                            done_evt.set()
                        self.after(0, _make)
                        done_evt.wait()
                    
                    content = delta.reasoning_content
                    r1_text += content
                    self.after(0, lambda c=content: append_text(self.current_r1_box, c))

                # 正文流
                if hasattr(delta, 'content') and delta.content:
                    if self.current_stream_box is None:
                        done_evt = threading.Event()
                        def _make():
                            self.stream_frame, self.current_stream_box = init_stream_ui(False)
                            done_evt.set()
                        self.after(0, _make)
                        done_evt.wait()
                        
                    content = delta.content
                    ai_text += content
                    self.after(0, lambda c=content: append_text(self.current_stream_box, c))

            # 生成结束，替换为完美气泡
            def replace_with_bubble():
                if self.current_r1_box: self.r1_frame.destroy()
                if self.current_stream_box: self.stream_frame.destroy()
                
                if r1_text: self.add_chat_bubble("ai", r1_text, is_reasoning=True)
                if ai_text: self.add_chat_bubble("ai", ai_text, is_reasoning=False)
                
            self.after(0, replace_with_bubble)
            self.chat_history.append({"role": "assistant", "content": ai_text})

        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            self.is_running = False
            self.after(0, self.reset_ui_state)

if __name__ == "__main__":
    app = DeepSeekApp()
    app.mainloop()
