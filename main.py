import os
import sys
import platform
import subprocess
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox

# 界面初始化配置
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# [关键修复 1] 针对 Windows 系统，定义隐藏子进程窗口的宏
if platform.system() == "Windows":
    CREATE_NO_WINDOW = subprocess.CREATE_NO_WINDOW
else:
    CREATE_NO_WINDOW = 0

class PyPackagerPro(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("PyPackager Pro - 终极跨平台打包引擎")
        self.geometry("900x750")
        self.minsize(800, 700)
        
        self.assets_list = []
        
        # ============ 顶部标题 ============
        self.title_label = ctk.CTkLabel(self, text="PyPackager Pro", font=ctk.CTkFont(size=28, weight="bold"))
        self.title_label.pack(pady=(20, 10))

        # ============ 核心功能选项卡 ============
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

        # ============ 实时日志控制台 ============
        self.log_label = ctk.CTkLabel(self, text="实时终端日志输出:", font=ctk.CTkFont(weight="bold"))
        self.log_label.pack(padx=20, pady=(10, 0), anchor="w")

        self.log_textbox = ctk.CTkTextbox(self, state="disabled", wrap="word", height=150, font=ctk.CTkFont(family="Consolas", size=12))
        self.log_textbox.pack(padx=20, pady=5, fill="both", expand=True)

        # ============ 底部执行按钮 ============
        self.build_btn = ctk.CTkButton(self, text="🚀 启动智能打包", font=ctk.CTkFont(size=18, weight="bold"), height=50, command=self.start_build_thread)
        self.build_btn.pack(padx=20, pady=20, fill="x")

    # ------------------ UI 布局搭建 ------------------
    # (此部分与之前保持一致)

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
        self.opt_admin = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(self.frame_options, text="请求管理员权限 (Windows)", variable=self.opt_admin).pack(side="left", padx=10)

    def setup_env_tab(self):
        self.opt_venv = ctk.BooleanVar(value=True)
        ctk.CTkSwitch(self.tab_env, text="启用纯净虚拟环境打包 (推荐：可极大幅减小软件体积，防止依赖污染)", variable=self.opt_venv, font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=20, pady=20)
        
        frame = ctk.CTkFrame(self.tab_env, fg_color="transparent")
        frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(frame, text="依赖清单 (requirements.txt):").pack(side="left")
        self.req_entry = ctk.CTkEntry(frame, width=400, placeholder_text="如果不填，将只打包标准库...")
        self.req_entry.pack(side="left", padx=10)
        ctk.CTkButton(frame, text="浏览", width=80, command=lambda: self.select_file(self.req_entry, [("Text", "*.txt")])).pack(side="left")

        ctk.CTkLabel(self.tab_env, text="说明：\n开启此功能后，软件将在项目目录下自动创建一个名为 'build_env' 的隔离环境，\n并在其中安装所选的 requirements.txt，最后在该环境内执行 PyInstaller。\n这能有效解决您的软件因为包含了系统中无关的庞大第三方库而变得臃肿的问题。", justify="left", text_color="gray").pack(anchor="w", padx=20, pady=10)

    def setup_assets_tab(self):
        ctk.CTkLabel(self.tab_assets, text="附加资源 (图片、音频、配置、模型文件等)：").pack(anchor="w", padx=20, pady=10)
        self.assets_textbox = ctk.CTkTextbox(self.tab_assets, height=120)
        self.assets_textbox.pack(fill="x", padx=20, pady=5)
        self.assets_textbox.insert("end", "当前未添加任何附加文件。\n")
        self.assets_textbox.configure(state="disabled")
        
        btn_frame = ctk.CTkFrame(self.tab_assets, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=5)
        
        ctk.CTkButton(btn_frame, text="添加文件", command=self.add_asset_file).pack(side="left", padx=(0, 10))
        ctk.CTkButton(btn_frame, text="添加文件夹", command=self.add_asset_folder).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="清空资源", fg_color="darkred", hover_color="red", command=self.clear_assets).pack(side="right")

    def setup_cloud_tab(self):
        ctk.CTkLabel(self.tab_cloud, text="GitHub Actions 自动打包配置生成器", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=20, pady=10)
        ctk.CTkLabel(self.tab_cloud, text="无法在 Windows 上打包 macOS 软件？\n一键生成 CI/CD 脚本，推送到 GitHub 后，云端会自动为您同时编译 Windows、macOS 和 Linux 版本！", justify="left", text_color="gray").pack(anchor="w", padx=20, pady=5)
        ctk.CTkButton(self.tab_cloud, text="生成 GitHub Actions Workflow (.yml)", command=self.generate_github_actions, height=40).pack(anchor="w", padx=20, pady=20)

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

    def generate_github_actions(self):
        # ... 保持与原版一致即可 ...
        pass

    # ================== [关键修复 2] 真正的线程安全日志机制 ==================
    def log_message(self, message):
        """
        线程安全的日志输出。
        当后台线程调用此方法时，它会将更新 UI 的任务委托给主线程执行，防止 UI 卡死。
        """
        self.after(0, self._insert_log, message)

    def _insert_log(self, message):
        """实际执行 UI 更新的方法（仅在主线程运行）"""
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", message + "\n")
        self.log_textbox.see("end")  # 自动滚动
        self.log_textbox.configure(state="disabled")

    def restore_button_state(self):
        """线程安全地恢复按钮状态"""
        self.build_btn.configure(state="normal", text="🚀 启动智能打包")

    # ------------------ 核心打包引擎逻辑 ------------------

    def start_build_thread(self):
        script_path = self.script_entry.get()
        if not script_path or not os.path.exists(script_path):
            messagebox.showerror("错误", "请先在【基础配置】中选择一个有效的 Python 主程序！")
            return

        self.build_btn.configure(state="disabled", text="⚙️ 引擎正在全力打包中...")
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")
        
        # 启动后台线程执行，防止卡死 UI
        threading.Thread(target=self.run_build_process, args=(script_path,), daemon=True).start()

    def run_build_process(self, script_path):
        work_dir = os.path.dirname(script_path)
        os_type = platform.system()
        
        pyinstaller_exe = "pyinstaller"
        
        try:
            if self.opt_venv.get():
                self.log_message("[*] ================= 环境隔离构建模式 =================")
                venv_dir = os.path.join(work_dir, "build_env")
                
                if os_type == "Windows":
                    python_exe = os.path.join(venv_dir, "Scripts", "python.exe")
                    pyinstaller_exe = os.path.join(venv_dir, "Scripts", "pyinstaller.exe")
                else:
                    python_exe = os.path.join(venv_dir, "bin", "python")
                    pyinstaller_exe = os.path.join(venv_dir, "bin", "pyinstaller")

                if not os.path.exists(venv_dir):
                    self.log_message(f"[*] 正在创建纯净虚拟环境: {venv_dir}")
                    # [关键修复 3] 加入 creationflags 防止弹窗
                    subprocess.run([sys.executable, "-m", "venv", venv_dir], check=True, creationflags=CREATE_NO_WINDOW)
                else:
                    self.log_message("[*] 发现现有虚拟环境，正在复用...")

                self.log_message("[*] 正在隔离环境中安装 PyInstaller...")
                # [关键修复 3] 加入 creationflags 防止弹窗
                subprocess.run([python_exe, "-m", "pip", "install", "pyinstaller", "--quiet"], check=True, creationflags=CREATE_NO_WINDOW)
                
                req_file = self.req_entry.get()
                if req_file and os.path.exists(req_file):
                    self.log_message(f"[*] 正在安装用户依赖 (requirements.txt)... 可能会花费一些时间。")
                    # [关键修复 3&4] 加入 creationflags 并设置 bufsize=1 实现行缓冲
                    process_pip = subprocess.Popen(
                        [python_exe, "-m", "pip", "install", "-r", req_file], 
                        stdout=subprocess.PIPE, 
                        stderr=subprocess.STDOUT, 
                        text=True, 
                        bufsize=1,
                        creationflags=CREATE_NO_WINDOW
                    )
                    for line in iter(process_pip.stdout.readline, ''):
                        if line: self.log_message(f"[PIP] {line.strip()}")
                    process_pip.wait()

            self.log_message("[*] ================= 准备打包引擎参数 =================")
            cmd = [pyinstaller_exe, "-y", "--clean"]
            
            if self.opt_onefile.get(): cmd.append("--onefile")
            if self.opt_windowed.get(): cmd.append("--windowed")
            if self.opt_admin.get(): cmd.append("--uac-admin")
                
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
            self.log_message(f"[*] 最终执行命令:\n{' '.join(cmd)}\n")

            self.log_message("[*] 🚀 引擎开始编译代码，请勿关闭软件...")
            
            # [关键修复 3&4] 隐藏 pyinstaller 执行过程的黑框，防止缓冲区阻塞
            process = subprocess.Popen(
                cmd,
                cwd=work_dir,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                creationflags=CREATE_NO_WINDOW
            )

            # 采用 iter 方式逐行读取，防止读取阻塞
            for line in iter(process.stdout.readline, ''):
                if line: self.log_message(f"[打包器] {line.strip()}")
                
            process.wait()

            if process.returncode == 0:
                self.log_message(f"\n[+] 🎉 恭喜！打包大功告成！")
                dist_dir = os.path.join(work_dir, 'dist')
                self.log_message(f"[+] 您的软件已输出至: {dist_dir}")
                # 尝试自动打开输出文件夹
                try:
                    if os_type == "Windows": os.startfile(dist_dir)
                    elif os_type == "Darwin": subprocess.run(["open", dist_dir])
                except Exception:
                    pass
            else:
                self.log_message("\n[x] ⚠️ 打包失败，请检查上方日志中的红色或 Error 信息。")

        except Exception as e:
            self.log_message(f"\n[x] 发生系统级错误: {str(e)}")
        finally:
            # 无论成功失败，恢复按钮状态都必须在主线程执行
            self.after(0, self.restore_button_state)

if __name__ == "__main__":
    app = PyPackagerPro()
    app.mainloop()
