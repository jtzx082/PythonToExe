import os
import sys
import asyncio
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import docx  # python-docx
import edge_tts
from openai import OpenAI # 用于调用 DeepSeek
from moviepy.editor import AudioFileClip, ColorClip

# --- 配置部分 ---
# 请在环境变量中设置 DEEPSEEK_API_KEY，或者直接在下方填入（不推荐直接填入代码中）
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "") 
DEEPSEEK_BASE_URL = "https://api.deepseek.com"

class TTSApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DeepSeek 智能语音合成助手")
        self.root.geometry("800x600")
        
        # 状态变量
        self.is_playing = False
        self.temp_audio_file = "temp_preview.mp3"
        self.loop = asyncio.new_event_loop()
        
        # 启动异步事件循环线程
        threading.Thread(target=self.start_loop, daemon=True).start()

        self.create_ui()

    def start_loop(self):
        asyncio.set_event_loop(self.loop)
        self.loop.run_forever()

    def create_ui(self):
        # 顶部按钮区：文件操作
        frame_top = tk.Frame(self.root)
        frame_top.pack(pady=10, fill=tk.X, padx=10)
        
        tk.Button(frame_top, text="📂 导入文本/Word", command=self.import_file).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_top, text="🧹 清空内容", command=self.clear_text).pack(side=tk.LEFT, padx=5)
        
        # 中间：文本输入区
        self.text_area = scrolledtext.ScrolledText(self.root, font=("Arial", 12))
        self.text_area.pack(expand=True, fill=tk.BOTH, padx=10, pady=5)
        
        # DeepSeek 功能区
        frame_ai = tk.Frame(self.root)
        frame_ai.pack(pady=5, fill=tk.X, padx=10)
        tk.Label(frame_ai, text="AI 辅助:", fg="blue").pack(side=tk.LEFT)
        tk.Button(frame_ai, text="✨ 使用 DeepSeek 润色文本", command=self.run_deepseek_polish, bg="#e1f5fe").pack(side=tk.LEFT, padx=5)
        
        # 底部：控制与导出
        frame_bottom = tk.Frame(self.root)
        frame_bottom.pack(pady=15, fill=tk.X, padx=10)
        
        tk.Button(frame_bottom, text="▶️ 生成并播放", command=self.play_audio, bg="#e8f5e9", width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_bottom, text="⏹️ 停止播放", command=self.stop_audio, bg="#ffebee").pack(side=tk.LEFT, padx=5)
        
        tk.Label(frame_bottom, text="|").pack(side=tk.LEFT, padx=10)
        
        tk.Button(frame_bottom, text="💾 导出 MP3", command=lambda: self.export_audio("mp3")).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_bottom, text="🎬 导出 WMV", command=lambda: self.export_audio("wmv")).pack(side=tk.LEFT, padx=5)
        
        # 状态栏
        self.status_label = tk.Label(self.root, text="就绪", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

    def update_status(self, text):
        self.status_label.config(text=text)
        self.root.update_idletasks()

    # --- 文件处理 ---
    def import_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Text/Word", "*.txt *.docx")])
        if not file_path: return
        
        content = ""
        try:
            if file_path.endswith(".txt"):
                with open(file_path, "r", encoding="utf-8") as f:
                    content = f.read()
            elif file_path.endswith(".docx"):
                doc = docx.Document(file_path)
                content = "\n".join([para.text for para in doc.paragraphs])
            
            self.text_area.delete("1.0", tk.END)
            self.text_area.insert(tk.END, content)
            self.update_status(f"已导入: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("错误", f"无法读取文件: {str(e)}")

    def clear_text(self):
        self.text_area.delete("1.0", tk.END)
        self.update_status("已清空")

    # --- DeepSeek API 调用 ---
    def run_deepseek_polish(self):
        text = self.text_area.get("1.0", tk.END).strip()
        if not text:
            messagebox.showwarning("提示", "请输入需要润色的内容")
            return
            
        if not DEEPSEEK_API_KEY:
            # 尝试弹窗让用户输入 Key
            key = tk.simpledialog.askstring("DeepSeek API Key", "请输入你的 DeepSeek API Key:", show="*")
            if not key: return
            globals()["DEEPSEEK_API_KEY"] = key

        threading.Thread(target=self._deepseek_thread, args=(text,)).start()

    def _deepseek_thread(self, text):
        self.update_status("正在连接 DeepSeek 进行润色...")
        try:
            client = OpenAI(api_key=DEEPSEEK_API_KEY, base_url=DEEPSEEK_BASE_URL)
            response = client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": "你是一个专业的文本润色助手。请将用户的输入修改为更自然、流畅、适合朗读的口语化文本。保持原意，但修正语病。直接输出润色后的文本，不要包含解释。"},
                    {"role": "user", "content": text},
                ],
                stream=False
            )
            polished_text = response.choices[0].message.content
            
            # 回到主线程更新 UI
            self.root.after(0, lambda: self.text_area.delete("1.0", tk.END))
            self.root.after(0, lambda: self.text_area.insert(tk.END, polished_text))
            self.root.after(0, lambda: self.update_status("DeepSeek 润色完成"))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("API 错误", str(e)))
            self.root.after(0, lambda: self.update_status("API 调用失败"))

    # --- 语音合成逻辑 (Edge-TTS) ---
    async def _generate_audio(self, text, output_file):
        # 使用中文语音，可根据需要修改为 zh-CN-YunjianNeural 等
        voice = "zh-CN-XiaoxiaoNeural" 
        communicate = edge_tts.Communicate(text, voice)
        await communicate.save(output_file)

    def play_audio(self):
        text = self.text_area.get("1.0", tk.END).strip()
        if not text: return
        
        self.stop_audio() # 先停止之前的
        self.update_status("正在生成语音...")
        
        def run_gen():
            future = asyncio.run_coroutine_threadsafe(
                self._generate_audio(text, self.temp_audio_file), self.loop
            )
            try:
                future.result() # 等待完成
                self.root.after(0, self._play_sound_file)
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("错误", str(e)))

        threading.Thread(target=run_gen).start()

    def _play_sound_file(self):
        import pygame
        pygame.mixer.init()
        pygame.mixer.music.load(self.temp_audio_file)
        pygame.mixer.music.play()
        self.is_playing = True
        self.update_status("正在播放...")

    def stop_audio(self):
        import pygame
        try:
            pygame.mixer.init()
            if pygame.mixer.music.get_busy():
                pygame.mixer.music.stop()
                pygame.mixer.music.unload()
        except:
            pass
        self.is_playing = False
        self.update_status("已停止")

    # --- 导出功能 ---
    def export_audio(self, fmt):
        text = self.text_area.get("1.0", tk.END).strip()
        if not text: return

        file_types = [("MP3 Audio", "*.mp3")] if fmt == "mp3" else [("WMV Video", "*.wmv")]
        save_path = filedialog.asksaveasfilename(defaultextension=f".{fmt}", filetypes=file_types)
        if not save_path: return

        self.update_status(f"正在导出 {fmt}...")

        def run_export():
            try:
                # 1. 先生成 MP3
                temp_mp3 = "temp_export.mp3"
                future = asyncio.run_coroutine_threadsafe(
                    self._generate_audio(text, temp_mp3), self.loop
                )
                future.result()

                # 2. 如果是 WMV，进行转换
                if fmt == "wmv":
                    audio = AudioFileClip(temp_mp3)
                    # 创建一个黑色背景的视频，时长等于音频时长
                    video = ColorClip(size=(640, 480), color=(0,0,0), duration=audio.duration)
                    video = video.set_audio(audio)
                    # 导出 WMV (使用 wmv 编码器或 libx264)
                    video.write_videofile(save_path, fps=1, codec="libx264", audio_codec="aac")
                    audio.close()
                    video.close()
                else:
                    # 如果是 MP3，直接重命名或移动
                    import shutil
                    shutil.move(temp_mp3, save_path)

                self.root.after(0, lambda: messagebox.showinfo("成功", f"文件已导出到: {save_path}"))
                self.root.after(0, lambda: self.update_status("导出完成"))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("导出错误", str(e)))
                self.root.after(0, lambda: self.update_status("导出失败"))

        threading.Thread(target=run_export).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = TTSApp(root)
    root.mainloop()
