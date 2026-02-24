import os
import sys
import platform
import subprocess
import threading
import queue
import customtkinter as ctk
from tkinter import filedialog, messagebox

# 界面初始化配置
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class PyPackagerPro(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("PyPackager Pro - Ubuntu 跨平台打包引擎")
        self.geometry("900x750")
        self.minsize(800, 700)
        
        self.assets_list = []
        self.log_queue = queue.Queue()
        self.after(100, self.process_log_queue) # 启动安全队列
        
        # ============ 界面布局 ============
        self.title_label = ctk.CTkLabel(self, text="PyPackager Pro", font=ctk.CTkFont(size=28, weight="bold"))
        self.title_label.pack(pady=(20, 10))

        self.tabview = ctk.CTkTabview(self, width=850, height=350)
        self.tabview.pack(padx=20, pady=10, fill="x")
        
        self.tab_basic = self.tabview.add("基础配置")
        self.tab_env = self.tabview.add("环境与依赖 (高级)")
        self.tab_assets = self.tabview.add("资源与数据")
        self.tab_cloud = self.tabview.add("云端跨平台 (CI/CD)")

        self.setup_basic_tab()
        self.setup_env_tab()
        self.setup_assets_tab()
        self.setup_cloud_tab()

        self.log_label = ctk.CTkLabel(self, text="实时终端日志输出:", font=ctk.CTkFont(weight="bold"))
        self.log_label.pack(padx=20, pady=(10, 0), anchor="w")

        self.log_textbox = ctk.CTkTextbox(self, state="disabled", wrap="word", height=150, font=ctk.CTkFont(family="Consolas", size=12))
        self.log_textbox.pack(padx=20, pady=5, fill="both", expand=True)

        self.build_btn = ctk.CTkButton(self, text="🚀 启动智能打包", font=ctk.CTkFont(size=18, weight="bold"), height=50, command=self.start_build_thread)
        self.build_btn.pack(padx=20, pady=20, fill="x")

    # ------------------ UI 布局搭建 ------------------
    def setup_basic_tab(self):
        ctk.CTkLabel(self.tab_basic, text="Python 主程序 (.py):").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.script_entry = ctk.CTkEntry(self.tab_basic, width=500)
        self.script_entry.grid(row=0, column=1, padx=10, pady=10)
        ctk.CTkButton(self.tab_basic, text="浏览", width=80, command=lambda: self.select_file(self.script_entry, [("Python", "*.py")])).grid(row=0, column=2, padx=10, pady=10)

        ctk.CTkLabel(self.tab_basic, text="软件图标 (.ico/.icns):").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.icon_entry = ctk.CTkEntry(self.tab_basic, width=500)
        self.icon_entry.grid(row=1, column=1, padx=10, pady=10)
        ctk.CTkButton(self.tab_basic, text="浏览", width=80, command=lambda: self.select_file(self.icon_entry, [("Icon", "*.ico *.icns")])).grid(row=1, column=2, padx=10, pady=10)
        
        ctk.CTkLabel(self.tab_basic, text="输出软件名称 (可选):").grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.name_entry = ctk.CTkEntry(self.tab_basic, width=500, placeholder_text="默认与主程序同名")
        self.name_entry.grid(row=2, column=1, padx=10, pady=10)

        self.frame_options = ctk.CTkFrame(self.tab_basic, fg_color="transparent")
        self.frame_options.grid(row=3, column=0, columnspan=3, pady=20, sticky="w")
        
        self.opt_onefile = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(self.frame_options, text="打包为单文件 (-F)", variable=self.opt_onefile).pack(side="left", padx=10)
        self.opt_windowed = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(self.frame_options, text="隐藏控制台 (GUI程序适用 -w)", variable=self.opt_windowed).pack(side="left", padx=10)

    def setup_env_tab(self):
        self.opt_venv = ctk.BooleanVar(value=True)
        ctk.CTkSwitch(self.tab_env, text="启用纯净虚拟环境打包 (推荐开启)", variable=self.opt_venv, font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=20, pady=20)
        
        frame = ctk.CTkFrame(self.tab_env, fg_color="transparent")
        frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(frame, text="依赖清单 (requirements.txt):").pack(side="left")
        self.req_entry = ctk.CTkEntry(frame, width=400, placeholder_text="如果不填，将只打包标准库...")
        self.req_entry.pack(side="left", padx=10)
        ctk.CTkButton(frame, text="浏览", width=80, command=lambda: self.select_file(self.req_entry, [("Text", "*.txt")])).pack(side="left")

    def setup_assets_tab(self):
        ctk.CTkLabel(self.tab_assets, text="附加资源 (图片、音频等)：").pack(anchor="w", padx=20, pady=10)
        self.assets_textbox = ctk.CTkTextbox(self.tab_assets, height=120)
        self.assets_textbox.pack(fill="x", padx=20, pady=5)
        self.assets_textbox.insert("end", "当前未添加任何附加文件。\n")
        self.assets_textbox.configure(state="disabled")
        
        btn_frame = ctk.CTkFrame(self.tab_assets, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=5)
        
        ctk.CTkButton(btn_frame, text="添加文件", command=self.add_asset_file).pack(side="left", padx=(0, 10))
        ctk.CTkButton(btn_frame, text="添加文件夹", command=self.add_asset_folder).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="清空", fg_color="darkred", hover_color="red", command=self.clear_assets).pack(side="right")

    def setup_cloud_tab(self):
        ctk.CTkLabel(self.tab_cloud, text="GitHub Actions 自动打包配置生成器", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=20, pady=10)
        ctk.CTkButton(self.tab_cloud, text="生成 Workflow (.yml)", height=40).pack(anchor="w", padx=20, pady=20)

    # ------------------ 辅助逻辑 ------------------
    def select_file(self, entry_widget, filetypes):
        path = filedialog.askopenfilename(filetypes=filetypes)
        if path:
            entry_widget.delete(0, "end")
            entry_widget.insert(0, path)

    def add_asset_file(self):
        paths = filedialog.askopenfilenames()
        for path in paths: self.assets_list.append((path, "."))
        self.update_assets_display()

    def add_asset_folder(self):
        path = filedialog.askdirectory()
        if path: self.assets_list.append((path, os.path.basename(path)))
        self.update_assets_display()

    def clear_assets(self):
        self.assets_list.clear()
        self.update_assets_display()

    def update_assets_display(self):
        self.assets_textbox.configure(state="normal")
        self.assets_textbox.delete("1.0", "end")
        for src, dest in self.assets_list: self.assets_textbox.insert("end", f"源: {src}  --->  目标文件夹: {dest}\n")
        self.assets_textbox.configure(state="disabled")

    # ================== 队列处理与线程安全 ==================
    def log_message(self, message):
        self.log_queue.put(message)

    def process_log_queue(self):
        try:
            logs = []
            while True: logs.append(self.log_queue.get_nowait())
        except queue.Empty: pass
            
        if logs:
            self.log_textbox.configure(state="normal")
            for log in logs: self.log_textbox.insert("end", log + "\n")
            self.log_textbox.see("end")
            self.log_textbox.configure(state="disabled")
            
        self.after(100, self.process_log_queue)

    def restore_button_state(self):
        self.build_btn.configure(state="normal", text="🚀 启动智能打包")

    # ================== 核心修复：针对 Linux/Ubuntu 的进程调度器 ==================
    def execute_command(self, cmd, cwd=None, prefix=""):
        # 获取当前系统的环境变量副本
        custom_env = os.environ.copy()
        
        # 针对 Ubuntu 的致命一击 1：斩断 pip 唤起系统 Keyring 密码弹窗的途径！
        custom_env["PYTHON_KEYRING_BACKEND"] = "keyring.backends.null.Keyring"
        
        # 针对 Ubuntu 的致命一击 2：强制 Linux 管道无缓冲，防止假死死锁！
        custom_env["PYTHONUNBUFFERED"] = "1"

        kwargs = {
            'stdout': subprocess.PIPE,
            'stderr': subprocess.STDOUT,
            'stdin': subprocess.PIPE, # 关闭输入，防止在后台偷偷要求按 Y/N
            'text': True,
            'errors': 'replace',
            'env': custom_env
        }
        
        if platform.system() == "Windows":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            kwargs['startupinfo'] = startupinfo
            kwargs['creationflags'] = 0x08000000 
            
        if cwd: kwargs['cwd'] = cwd

        process = subprocess.Popen(cmd, **kwargs)
        process.stdin.close() 

        for line in process.stdout:
            if line.strip(): self.log_message(f"{prefix}{line.strip()}")
                
        process.wait()
        return process.returncode

    # ------------------ 核心打包流程 ------------------
    def start_build_thread(self):
        script_path = self.script_entry.get()
        if not script_path or not os.path.exists(script_path):
            messagebox.showerror("错误", "请先选择 Python 主程序！")
            return

        self.build_btn.configure(state="disabled", text="⚙️ 引擎正在全力打包中...")
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")
        
        threading.Thread(target=self.run_build_process, args=(script_path,), daemon=True).start()

    def run_build_process(self, script_path):
        work_dir = os.path.dirname(script_path)
        os_type = platform.system()
        pyinstaller_exe = "pyinstaller"
        
        try:
            if self.opt_venv.get():
                self.log_message("[*] ================= 初始化纯净虚拟环境 =================")
                venv_dir = os.path.join(work_dir, "build_env")
                
                if os_type == "Windows":
                    python_exe = os.path.join(venv_dir, "Scripts", "python.exe")
                    pyinstaller_exe = os.path.join(venv_dir, "Scripts", "pyinstaller.exe")
                else:
                    # Ubuntu / Linux 环境路径
                    python_exe = os.path.join(venv_dir, "bin", "python")
                    pyinstaller_exe = os.path.join(venv_dir, "bin", "pyinstaller")

                if not os.path.exists(venv_dir):
                    self.log_message(f"[*] 正在创建虚拟环境...")
                    # Ubuntu 特殊提醒：如果这里报错，说明系统缺包
                    ret = self.execute_command([sys.executable, "-m", "venv", venv_dir], prefix="[系统] ")
                    if ret != 0: 
                        self.log_message("[x] 严重错误：Ubuntu 中可能未安装 venv 模块。")
                        self.log_message("[!] 请打开您的 Ubuntu 终端，手动执行一次：sudo apt install python3-venv")
                        raise Exception("虚拟环境创建失败。")
                else:
                    self.log_message("[*] 发现现有虚拟环境，正在复用...")

                self.log_message("[*] 正在安装底层打包引擎 (PyInstaller)...")
                self.execute_command([python_exe, "-m", "pip", "install", "pyinstaller", "--quiet"], prefix="[PIP] ")
                
                req_file = self.req_entry.get()
                if req_file and os.path.exists(req_file):
                    self.log_message(f"[*] 正在安装依赖 (requirements.txt)...")
                    self.execute_command([python_exe, "-m", "pip", "install", "-r", req_file], prefix="[PIP] ")

            self.log_message("[*] ================= 准备启动引擎打包 =================")
            cmd = [pyinstaller_exe, "-y", "--clean"]
            
            if self.opt_onefile.get(): cmd.append("--onefile")
            if self.opt_windowed.get(): cmd.append("--windowed")
                
            app_name = self.name_entry.get()
            if app_name: cmd.extend(["--name", app_name])
                
            icon_path = self.icon_entry.get()
            if icon_path and os.path.exists(icon_path):
                cmd.append(f"--icon={icon_path}")
                
            if self.assets_list:
                separator = ";" if os_type == "Windows" else ":"
                for src, dest in self.assets_list:
                    cmd.append(f"--add-data={src}{separator}{dest}")
                    
            cmd.append(script_path)
            self.log_message(f"[*] 执行指令: {' '.join(cmd)}\n")

            self.log_message("[*] 🚀 编译正式开始，这在 Ubuntu 上可能需要一两分钟...")
            
            retcode = self.execute_command(cmd, cwd=work_dir, prefix="[打包器] ")

            if retcode == 0:
                dist_dir = os.path.join(work_dir, 'dist')
                self.log_message(f"\n[+] 🎉 恭喜！Ubuntu 版本打包大功告成！")
                self.log_message(f"[+] 您的可执行文件已输出至: {dist_dir}")
                # 适配 Ubuntu 的自动打开文件夹命令
                try:
                    if os_type == "Windows": os.startfile(dist_dir)
                    elif os_type == "Darwin": subprocess.run(["open", dist_dir])
                    elif os_type == "Linux": subprocess.run(["xdg-open", dist_dir])
                except: pass
            else:
                self.log_message("\n[x] ⚠️ 打包失败，请往上翻阅日志查看具体的 Error 信息。")

        except Exception as e:
            self.log_message(f"\n[x] 发生严重错误: {str(e)}")
        finally:
            self.after(0, self.restore_button_state)

if __name__ == "__main__":
    app = PyPackagerPro()
    app.mainloop()
