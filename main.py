import os
import sys
import shutil
import platform
import subprocess
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox

# 界面初始化配置
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class PyPackagerPro(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("PyPackager Pro - 终极跨平台打包引擎")
        self.geometry("900x750")
        self.minsize(800, 700)
        
        self.assets_list = []  # 存储附加数据文件/文件夹的列表
        
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

    def setup_basic_tab(self):
        """基础配置选项卡"""
        # 主脚本
        ctk.CTkLabel(self.tab_basic, text="Python 主程序 (.py):").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.script_entry = ctk.CTkEntry(self.tab_basic, width=500)
        self.script_entry.grid(row=0, column=1, padx=10, pady=10)
        ctk.CTkButton(self.tab_basic, text="浏览", width=80, command=lambda: self.select_file(self.script_entry, [("Python", "*.py")])).grid(row=0, column=2, padx=10, pady=10)

        # 图标
        ctk.CTkLabel(self.tab_basic, text="软件图标 (.ico/.icns):").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.icon_entry = ctk.CTkEntry(self.tab_basic, width=500)
        self.icon_entry.grid(row=1, column=1, padx=10, pady=10)
        ctk.CTkButton(self.tab_basic, text="浏览", width=80, command=lambda: self.select_file(self.icon_entry, [("Icon", "*.ico *.icns")])).grid(row=1, column=2, padx=10, pady=10)
        
        # 软件名称
        ctk.CTkLabel(self.tab_basic, text="输出软件名称 (可选):").grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.name_entry = ctk.CTkEntry(self.tab_basic, width=500, placeholder_text="默认与主程序同名")
        self.name_entry.grid(row=2, column=1, padx=10, pady=10)

        # 打包选项
        self.frame_options = ctk.CTkFrame(self.tab_basic, fg_color="transparent")
        self.frame_options.grid(row=3, column=0, columnspan=3, pady=20, sticky="w")
        
        self.opt_onefile = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(self.frame_options, text="打包为单文件 (-F)", variable=self.opt_onefile).pack(side="left", padx=10)
        self.opt_windowed = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(self.frame_options, text="隐藏控制台 (GUI程序适用 -w)", variable=self.opt_windowed).pack(side="left", padx=10)
        self.opt_admin = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(self.frame_options, text="请求管理员权限 (Windows)", variable=self.opt_admin).pack(side="left", padx=10)

    def setup_env_tab(self):
        """虚拟环境配置选项卡 - 解决打包体积过大的核心"""
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
        """资源文件选项卡 - 解决打包后找不到图片、模型等文件的问题"""
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
        """云端跨平台 CI/CD 生成器"""
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
        for path in paths:
            self.assets_list.append((path, ".")) # 默认放到根目录
        self.update_assets_display()

    def add_asset_folder(self):
        path = filedialog.askdirectory()
        if path:
            folder_name = os.path.basename(path)
            self.assets_list.append((path, folder_name)) # 保持文件夹结构
        self.update_assets_display()

    def clear_assets(self):
        self.assets_list.clear()
        self.update_assets_display()

    def update_assets_display(self):
        self.assets_textbox.configure(state="normal")
        self.assets_textbox.delete("1.0", "end")
        for src, dest in self.assets_list:
            self.assets_textbox.insert("end", f"源: {src}  --->  目标文件夹: {dest}\n")
        self.assets_textbox.configure(state="disabled")

    def log_message(self, message):
        """线程安全的日志输出"""
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", message + "\n")
        self.log_textbox.see("end")
        self.log_textbox.configure(state="disabled")

    def generate_github_actions(self):
        script_name = os.path.basename(self.script_entry.get()) if self.script_entry.get() else "main.py"
        req_line = "pip install -r requirements.txt" if self.req_entry.get() else ""
        
        yml_content = f"""name: Build Multi-Platform Python App
on: [push, pull_request]

jobs:
  build:
    runs-on: ${{{{ matrix.os }}}}
    strategy:
      matrix:
        os: [ubuntu-latest, macos-latest, windows-latest]
    steps:
    - uses: actions/checkout@v3
    - name: Set up Python
      uses: actions/setup-python@v4
      with:
        python-version: '3.10'
    - name: Install dependencies
      run: |
        python -m pip install --upgrade pip
        pip install pyinstaller
        {req_line}
    - name: Build with PyInstaller
      run: pyinstaller -y --onefile {"--windowed " if self.opt_windowed.get() else ""}{script_name}
    - name: Upload Artifact
      uses: actions/upload-artifact@v3
      with:
        name: executable-${{{{ matrix.os }}}}
        path: dist/
"""
        save_path = filedialog.asksaveasfilename(defaultextension=".yml", initialfile="build.yml", title="保存 GitHub Actions 配置文件")
        if save_path:
            with open(save_path, "w", encoding="utf-8") as f:
                f.write(yml_content)
            messagebox.showinfo("成功", f"CI/CD 脚本已保存至:\n{save_path}\n请将其放置在您的项目仓库的 .github/workflows/ 目录下！")

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
        
        # 1. 环境准备阶段 (Virtual Environment)
        pyinstaller_exe = "pyinstaller" # 默认使用系统全局环境变量
        
        if self.opt_venv.get():
            self.log_message("[*] ================= 环境隔离构建模式 =================")
            venv_dir = os.path.join(work_dir, "build_env")
            
            # 判断不同系统的 venv 路径
            if os_type == "Windows":
                python_exe = os.path.join(venv_dir, "Scripts", "python.exe")
                pip_exe = os.path.join(venv_dir, "Scripts", "pip.exe")
                pyinstaller_exe = os.path.join(venv_dir, "Scripts", "pyinstaller.exe")
            else:
                python_exe = os.path.join(venv_dir, "bin", "python")
                pip_exe = os.path.join(venv_dir, "bin", "pip")
                pyinstaller_exe = os.path.join(venv_dir, "bin", "pyinstaller")

            # 创建或清理 venv
            if not os.path.exists(venv_dir):
                self.log_message(f"[*] 正在创建纯净虚拟环境: {venv_dir}")
                subprocess.run([sys.executable, "-m", "venv", venv_dir], check=True)
            else:
                self.log_message("[*] 发现现有虚拟环境，正在复用...")

            # 安装依赖
            self.log_message("[*] 正在隔离环境中安装 PyInstaller...")
            subprocess.run([python_exe, "-m", "pip", "install", "pyinstaller", "--quiet"], check=True)
            
            req_file = self.req_entry.get()
            if req_file and os.path.exists(req_file):
                self.log_message(f"[*] 正在安装用户依赖 (requirements.txt)... 可能会花费一些时间。")
                process_pip = subprocess.Popen([python_exe, "-m", "pip", "install", "-r", req_file], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
                for line in process_pip.stdout:
                    self.log_message(f"[PIP] {line.strip()}")
                process_pip.wait()

        # 2. 构建 PyInstaller 命令行
        self.log_message("[*] ================= 准备打包引擎参数 =================")
        cmd = [pyinstaller_exe, "-y", "--clean"]
        
        if self.opt_onefile.get(): cmd.append("--onefile")
        if self.opt_windowed.get(): cmd.append("--windowed")
        if self.opt_admin.get(): cmd.append("--uac-admin")
            
        app_name = self.name_entry.get()
        if app_name:
            cmd.extend(["--name", app_name])
            
        icon_path = self.icon_entry.get()
        if icon_path and os.path.exists(icon_path):
            cmd.append(f"--icon={icon_path}")
            
        # 处理资源数据映射 (跨平台分隔符)
        if self.assets_list:
            separator = ";" if os_type == "Windows" else ":"
            for src, dest in self.assets_list:
                cmd.append(f"--add-data={src}{separator}{dest}")
                
        cmd.append(script_path)
        self.log_message(f"[*] 最终执行命令:\n{' '.join(cmd)}\n")

        # 3. 执行打包并捕获日志
        try:
            self.log_message("[*] 🚀 引擎开始编译代码，请勿关闭软件...")
            process = subprocess.Popen(
                cmd,
                cwd=work_dir,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW if os_type == 'Windows' else 0
            )

            for line in process.stdout:
                self.log_message(f"[打包器] {line.strip()}")
            process.wait()

            if process.returncode == 0:
                self.log_message(f"\n[+] 🎉 恭喜！打包大功告成！")
                self.log_message(f"[+] 您的软件已输出至: {os.path.join(work_dir, 'dist')}")
                # 打包成功后尝试自动打开文件夹 (限Windows/macOS)
                if os_type == "Windows": os.startfile(os.path.join(work_dir, 'dist'))
                elif os_type == "Darwin": subprocess.run(["open", os.path.join(work_dir, 'dist')])
            else:
                self.log_message("\n[x] ⚠️ 打包失败，请检查上方日志中的红色或 Error 信息。")

        except Exception as e:
            self.log_message(f"\n[x] 发生系统级错误: {str(e)}")
        finally:
            self.build_btn.configure(state="normal", text="🚀 启动智能打包")


if __name__ == "__main__":
    app = PyPackagerPro()
    app.mainloop()
