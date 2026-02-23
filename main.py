import os
import sys
import json
import shutil
import subprocess
import threading
import multiprocessing
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledText

AUTO_CONFIG_FILE = "pyinstaller_gui_history.json"

class PyInstallerGUI(ttk.Window):
    def __init__(self):
        super().__init__(themename="lumen")
        self.title("PyInstaller 打包工具 v5.6 (完美典藏版)")
        self.geometry("820x800")
        self.minsize(750, 650)
        
        self.process = None
        self.current_theme = "lumen"

        self._init_vars()
        self._create_menu()
        self._create_layout()

        self.load_config(AUTO_CONFIG_FILE, silent=True)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def _init_vars(self):
        self.var_req = tk.StringVar()
        self.var_script = tk.StringVar()
        self.var_outdir = tk.StringVar()
        self.var_outname = tk.StringVar()
        self.var_icon = tk.StringVar()
        
        self.var_onefile = tk.BooleanVar(value=True)
        self.var_console = tk.BooleanVar(value=True) 
        self.var_clean = tk.BooleanVar(value=True)
        self.var_upx = tk.BooleanVar(value=False)
        self.var_uac = tk.BooleanVar(value=False)
        
        self.var_add_data = tk.StringVar()
        self.var_hidden_imports = tk.StringVar()
        self.var_exclude_modules = tk.StringVar()
        
        self.var_use_venv = tk.BooleanVar(value=True) 

    def _create_menu(self):
        menubar = tk.Menu(self)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="导入配置...", command=self.import_config)
        file_menu.add_command(label="导出配置...", command=self.export_config)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.on_closing)
        menubar.add_cascade(label="文件", menu=file_menu)
        self.config(menu=menubar)

    def _create_layout(self):
        toolbar = ttk.Frame(self)
        toolbar.pack(fill=X, padx=10, pady=(10, 0))
        ttk.Label(toolbar, text="🚀 Python GUI & 脚本自动化打包引擎", font=("", 12, "bold")).pack(side=LEFT)
        ttk.Button(toolbar, text="🌓 切换主题", bootstyle=(SECONDARY, OUTLINE), command=self.toggle_theme).pack(side=RIGHT)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=BOTH, expand=False, padx=10, pady=10)
        
        self.tab_basic = ttk.Frame(self.notebook)
        self.tab_advanced = ttk.Frame(self.notebook)
        self.tab_env = ttk.Frame(self.notebook)
        self.tab_about = ttk.Frame(self.notebook)
        
        self.notebook.add(self.tab_basic, text="📦 基础配置")
        self.notebook.add(self.tab_advanced, text="🛠️ 高级设置")
        self.notebook.add(self.tab_env, text="🌱 依赖与隔离环境 (推荐)")
        self.notebook.add(self.tab_about, text="📖 关于与说明")
        
        self._build_basic_tab()
        self._build_advanced_tab()
        self._build_env_tab()
        self._build_about_tab()

        bottom_frame = ttk.Frame(self)
        bottom_frame.pack(fill=BOTH, expand=True, padx=10, pady=(0, 10))
        
        btn_bar = ttk.Frame(bottom_frame)
        btn_bar.pack(fill=X, pady=5)
        
        self.btn_open_dir = ttk.Button(btn_bar, text="打开输出目录", bootstyle=INFO, state=DISABLED, command=self.open_output_dir)
        self.btn_open_dir.pack(side=LEFT)
        
        # 新增：一键清空按钮
        self.btn_clear = ttk.Button(btn_bar, text="🧹 一键清空", bootstyle=(SECONDARY, OUTLINE), command=self.clear_all_inputs)
        self.btn_clear.pack(side=LEFT, padx=(10, 0))
        
        self.btn_cancel = ttk.Button(btn_bar, text="取消操作", bootstyle=DANGER, command=self.cancel_process, state=DISABLED)
        self.btn_cancel.pack(side=RIGHT, padx=(5, 0))
        
        self.btn_start = ttk.Button(btn_bar, text="一键执行打包", bootstyle=PRIMARY, command=self.start_build_thread)
        self.btn_start.pack(side=RIGHT)

        self.progress = ttk.Progressbar(bottom_frame, mode='indeterminate', bootstyle=INFO)
        self.progress.pack(fill=X, pady=(5, 10))
        
        frame_console = ttk.Labelframe(bottom_frame, text="实时日志终端", padding=5)
        frame_console.pack(fill=BOTH, expand=True)
        self.console_text = ScrolledText(frame_console, wrap=WORD, font=("Consolas", 10))
        self.console_text.pack(fill=BOTH, expand=True)

    def _build_basic_tab(self):
        f_script = ttk.Labelframe(self.tab_basic, text="主程序 (必填)", padding=10)
        f_script.pack(fill=X, pady=10, padx=10)
        ttk.Entry(f_script, textvariable=self.var_script).pack(side=LEFT, fill=X, expand=True, padx=5)
        ttk.Button(f_script, text="浏览...", command=self.browse_script).pack(side=LEFT)

        f_out = ttk.Labelframe(self.tab_basic, text="输出与外观 (可选)", padding=10)
        f_out.pack(fill=X, pady=5, padx=10)
        
        ttk.Label(f_out, text="输出目录:").grid(row=0, column=0, sticky=W, pady=5)
        ttk.Entry(f_out, textvariable=self.var_outdir, bootstyle="info").grid(row=0, column=1, sticky=EW, padx=5, pady=5)
        ttk.Button(f_out, text="浏览...", command=self.browse_outdir).grid(row=0, column=2, pady=5)
        
        ttk.Label(f_out, text="应用名称:").grid(row=1, column=0, sticky=W, pady=5)
        ttk.Entry(f_out, textvariable=self.var_outname).grid(row=1, column=1, sticky=EW, padx=5, pady=5)
        
        ttk.Label(f_out, text="应用图标:").grid(row=2, column=0, sticky=W, pady=5)
        ttk.Entry(f_out, textvariable=self.var_icon).grid(row=2, column=1, sticky=EW, padx=5, pady=5)
        ttk.Button(f_out, text="浏览...", command=self.browse_icon).grid(row=2, column=2, pady=5)
        f_out.columnconfigure(1, weight=1)

        f_opt = ttk.Labelframe(self.tab_basic, text="核心模式", padding=10)
        f_opt.pack(fill=X, pady=5, padx=10)
        ttk.Checkbutton(f_opt, text="打包为单文件 (-F)", variable=self.var_onefile).pack(side=LEFT, padx=15)
        ttk.Checkbutton(f_opt, text="隐藏控制台黑框 (-w, 适合 GUI 程序)", variable=self.var_console).pack(side=LEFT, padx=15)

    def _build_advanced_tab(self):
        f_data = ttk.Labelframe(self.tab_advanced, text="资源与依赖管理", padding=10)
        f_data.pack(fill=X, pady=10, padx=10)
        
        ttk.Label(f_data, text="附加数据:").grid(row=0, column=0, sticky=W, pady=5)
        ttk.Entry(f_data, textvariable=self.var_add_data).grid(row=0, column=1, sticky=EW, padx=5, pady=5)
        ttk.Button(f_data, text="添加...", command=self.browse_add_data).grid(row=0, column=2, pady=5)
        
        ttk.Label(f_data, text="隐式导入:").grid(row=1, column=0, sticky=W, pady=5)
        ttk.Entry(f_data, textvariable=self.var_hidden_imports).grid(row=1, column=1, columnspan=2, sticky=EW, padx=5, pady=5)

        ttk.Label(f_data, text="排除模块:").grid(row=2, column=0, sticky=W, pady=5)
        ttk.Entry(f_data, textvariable=self.var_exclude_modules).grid(row=2, column=1, columnspan=2, sticky=EW, padx=5, pady=5)
        f_data.columnconfigure(1, weight=1)

        f_build = ttk.Labelframe(self.tab_advanced, text="构建参数", padding=10)
        f_build.pack(fill=X, pady=5, padx=10)
        ttk.Checkbutton(f_build, text="打包后清理临时文件 (--clean)", variable=self.var_clean).pack(anchor=W, pady=3)
        ttk.Checkbutton(f_build, text="使用 UPX 极致压缩 (--upx-dir)", variable=self.var_upx).pack(anchor=W, pady=3)
        ttk.Checkbutton(f_build, text="请求管理员权限 (Windows 提权)", variable=self.var_uac).pack(anchor=W, pady=3)

    def _build_env_tab(self):
        f_env = ttk.Labelframe(self.tab_env, text="沙盒隔离打包 (极限压缩体积)", padding=20)
        f_env.pack(fill=X, pady=20, padx=20)
        
        desc = ("建议启用【纯净虚拟环境】！工具会在后台创建一个隔离的沙盒，"
                "并仅安装必要的依赖进行打包，彻底杜绝生成的 exe 体积臃肿问题。")
        desc_lbl = ttk.Label(f_env, text=desc, justify=LEFT)
        desc_lbl.pack(anchor=W, pady=(0, 15), fill=X)
        desc_lbl.bind('<Configure>', lambda e: e.widget.config(wraplength=e.width))
        
        ttk.Checkbutton(f_env, text="启用纯净虚拟环境打包 (.pack_venv)", variable=self.var_use_venv, bootstyle="success-round-toggle").pack(anchor=W, pady=(0, 15))
        
        row = ttk.Frame(f_env)
        row.pack(fill=X)
        ttk.Label(row, text="指定专属依赖 (requirements.txt):").pack(side=LEFT)
        ttk.Entry(row, textvariable=self.var_req).pack(side=LEFT, fill=X, expand=True, padx=5)
        ttk.Button(row, text="浏览...", command=self.browse_req).pack(side=LEFT, padx=(0, 5))

    def _build_about_tab(self):
        f_guide = ttk.Labelframe(self.tab_about, text="💡 软件使用说明", padding=15)
        f_guide.pack(fill=X, pady=10, padx=20)
        
        guide_text = (
            "1. 基础配置：选择您编写的 Python 主程序 (.py/.pyw 文件)。如果是带有图形界面的程序，建议保持勾选“隐藏控制台黑框”。\n\n"
            "2. 极限压缩（推荐）：切换到【🌱 依赖与隔离环境】标签页，勾选“启用纯净虚拟环境”。如果代码使用了第三方库，请务必指定 requirements.txt 文件。工具将在沙盒中独立打包，杜绝体积臃肿。\n\n"
            "3. 解决报错：如果打包生成的软件在运行时闪退并提示 'ModuleNotFoundError'，请在【🛠️ 高级设置】的“隐式导入”中填入报错缺失的模块名，然后重新打包即可解决。\n\n"
            "4. 一键执行：配置完成后，点击右下角按钮，静待终端输出“🎉 打包圆满完成”的提示即可提取软件。"
        )
        guide_lbl = ttk.Label(f_guide, text=guide_text, justify=LEFT)
        guide_lbl.pack(anchor=W, fill=X)
        guide_lbl.bind('<Configure>', lambda e: e.widget.config(wraplength=e.width)) 

        f_author = ttk.Labelframe(self.tab_about, text="👨‍💻 关于作者", padding=15)
        f_author.pack(fill=X, pady=10, padx=20)
        
        author_text = (
            "开发与维护：俞晋全\n"
            "个人博客：电子云\n\n"
            "本工具致力于为广大的 Python 开发者、教师同仁提供一款轻量且强大的跨平台打包解决方案。无论是开发日常的教学辅助脚本、成绩统计分析软件，还是复杂的应用系统，都能通过自动化的沙盒纯净打包机制，彻底告别环境污染和软件体积臃肿的烦恼。"
        )
        author_lbl = ttk.Label(f_author, text=author_text, justify=LEFT)
        author_lbl.pack(anchor=W, fill=X)
        author_lbl.bind('<Configure>', lambda e: e.widget.config(wraplength=e.width)) 

    # --- 界面控制与配置 ---
    def toggle_theme(self):
        if self.current_theme == "lumen":
            self.style.theme_use("cyborg")
            self.current_theme = "cyborg"
        else:
            self.style.theme_use("lumen")
            self.current_theme = "lumen"

    def open_output_dir(self):
        out_dir = self.var_outdir.get() or os.path.join(os.path.dirname(self.var_script.get()), "dist")
        if os.path.exists(out_dir):
            if sys.platform == "win32": os.startfile(out_dir)
            elif sys.platform == "darwin": subprocess.Popen(["open", out_dir])
            else: subprocess.Popen(["xdg-open", out_dir])
        else: messagebox.showwarning("提示", "输出目录不存在！")

    def clear_all_inputs(self):
        """一键清空所有输入框和选项，恢复初始状态"""
        if messagebox.askyesno("确认清空", "确定要清空当前所有填写的路径和配置参数吗？\n(此操作方便您准备打包下一个新项目)"):
            # 清空文本路径
            self.var_req.set("")
            self.var_script.set("")
            self.var_outdir.set("")
            self.var_outname.set("")
            self.var_icon.set("")
            self.var_add_data.set("")
            self.var_hidden_imports.set("")
            self.var_exclude_modules.set("")
            
            # 恢复默认勾选项
            self.var_onefile.set(True)
            self.var_console.set(True)
            self.var_clean.set(True)
            self.var_use_venv.set(True)
            self.var_upx.set(False)
            self.var_uac.set(False)
            
            # 清空控制台
            self.console_text.delete(1.0, END)
            self.log_console("✨ 所有配置已清空，您可以开始配置下一个打包项目了。\n")

    def get_current_config(self):
        return {
            "req_path": self.var_req.get(), "script_path": self.var_script.get(),
            "outdir": self.var_outdir.get(), "outname": self.var_outname.get(),
            "icon": self.var_icon.get(), "add_data": self.var_add_data.get(),
            "hidden_imports": self.var_hidden_imports.get(), "exclude_modules": self.var_exclude_modules.get(),
            "onefile": self.var_onefile.get(), "console": self.var_console.get(),
            "clean": self.var_clean.get(), "upx": self.var_upx.get(), "uac": self.var_uac.get(),
            "use_venv": self.var_use_venv.get()
        }

    def save_config(self, filepath, silent=False):
        try:
            with open(filepath, 'w', encoding='utf-8') as f: json.dump(self.get_current_config(), f, indent=4, ensure_ascii=False)
            if not silent: messagebox.showinfo("成功", "配置导出成功！")
        except: pass

    def load_config(self, filepath, silent=False):
        if not os.path.exists(filepath): return
        try:
            with open(filepath, 'r', encoding='utf-8') as f: cfg = json.load(f)
            self.var_req.set(cfg.get("req_path", ""))
            self.var_script.set(cfg.get("script_path", ""))
            self.var_outdir.set(cfg.get("outdir", ""))
            self.var_outname.set(cfg.get("outname", ""))
            self.var_icon.set(cfg.get("icon", ""))
            self.var_add_data.set(cfg.get("add_data", ""))
            self.var_hidden_imports.set(cfg.get("hidden_imports", ""))
            self.var_exclude_modules.set(cfg.get("exclude_modules", ""))
            self.var_onefile.set(cfg.get("onefile", True))
            self.var_console.set(cfg.get("console", True)) 
            self.var_clean.set(cfg.get("clean", True))
            self.var_upx.set(cfg.get("upx", False))
            self.var_uac.set(cfg.get("uac", False))
            self.var_use_venv.set(cfg.get("use_venv", True))
        except: pass

    def export_config(self):
        p = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if p: self.save_config(p)

    def import_config(self):
        p = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if p: self.load_config(p)

    def on_closing(self):
        self.save_config(AUTO_CONFIG_FILE, silent=True)
        if self.process: self.process.terminate()
        self.destroy()

    # --- 浏览文件 ---
    def browse_req(self):
        p = filedialog.askopenfilename(filetypes=[("Text", "*.txt")])
        if p: self.var_req.set(p)

    def browse_script(self):
        p = filedialog.askopenfilename(filetypes=[("Python", "*.py *.pyw")])
        if p: self.var_script.set(p)

    def browse_outdir(self):
        p = filedialog.askdirectory()
        if p: self.var_outdir.set(p)

    def browse_icon(self):
        p = filedialog.askopenfilename(filetypes=[("Icon", "*.ico *.icns")])
        if p: self.var_icon.set(p)

    def browse_add_data(self):
        p = filedialog.askdirectory(title="选择要包含的文件夹")
        if p: 
            sep = ";" if os.name == 'nt' else ":"
            self.var_add_data.set(f"{self.var_add_data.get()} {p}{sep}{os.path.basename(p)}".strip())

    # --- 环境自检逻辑 ---
    def get_system_python(self):
        if os.name == 'nt':
            return "python" if shutil.which("python") else None
        else:
            if shutil.which("python3"): return "python3"
            if shutil.which("python"): return "python"
            return None

    # --- 核心打包逻辑 ---
    def log_console(self, text):
        self.console_text.insert(END, text)
        self.console_text.see(END)

    def _lock_ui(self):
        self.btn_start.config(state=DISABLED)
        self.btn_cancel.config(state=NORMAL)
        self.btn_open_dir.config(state=DISABLED)
        self.btn_clear.config(state=DISABLED) # 打包时锁定清空按钮
        self.progress.start(10)

    def _unlock_ui(self):
        self.progress.stop()
        self.btn_start.config(state=NORMAL)
        self.btn_cancel.config(state=DISABLED)
        self.btn_open_dir.config(state=NORMAL) 
        self.btn_clear.config(state=NORMAL) # 恢复清空按钮
        self.process = None

    def start_build_thread(self):
        if not self.var_script.get():
            messagebox.showwarning("警告", "请先在基础配置中选择需要打包的 Python 脚本！")
            return
            
        sys_python = self.get_system_python()
        if not sys_python:
            messagebox.showerror(
                "环境缺失", 
                "⚠️ 未检测到本机的 Python 环境！\n\n本工具依赖底层 Python 解释器运行打包逻辑，请先在此电脑上安装 Python 并配置环境变量。"
            )
            return

        self._lock_ui()
        self.console_text.delete(1.0, END)
        self.save_config(AUTO_CONFIG_FILE, silent=True) 
        threading.Thread(target=self._run_build_pipeline, args=(sys_python,), daemon=True).start()

    def _run_cmd_blocking(self, cmd):
        try:
            kwargs = {}
            if os.name == 'nt': kwargs['creationflags'] = subprocess.CREATE_NO_WINDOW
            self.process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1, **kwargs)
            for line in self.process.stdout: self.log_console(line)
            self.process.wait()
            return self.process.returncode == 0
        except Exception as e:
            self.log_console(f"\n❌ 执行异常: {str(e)}\n")
            return False

    def _run_build_pipeline(self, system_python):
        script_dir = os.path.dirname(self.var_script.get())
        pyinstaller_exe = "pyinstaller"
        
        if self.var_use_venv.get():
            venv_dir = os.path.join(script_dir, ".pack_venv")
            self.log_console(f"🌱 [阶段 1/2] 正在调用系统环境构建纯净沙盒...\n路径: {venv_dir}\n")
            
            if not self._run_cmd_blocking([system_python, "-m", "venv", venv_dir, "--clear"]):
                self.log_console("\n❌ 虚拟环境创建失败！\n(提示: Ubuntu 等 Linux 系统请确保已通过终端执行过 sudo apt install python3-venv)\n")
                self.after(0, self._unlock_ui)
                return
                
            if sys.platform == "win32":
                v_python = os.path.join(venv_dir, "Scripts", "python.exe")
                pyinstaller_exe = os.path.join(venv_dir, "Scripts", "pyinstaller.exe")
            else:
                v_python = os.path.join(venv_dir, "bin", "python")
                pyinstaller_exe = os.path.join(venv_dir, "bin", "pyinstaller")
                
            self.log_console("\n📦 正在沙盒中静默安装 PyInstaller 核心库...\n")
            if not self._run_cmd_blocking([v_python, "-m", "pip", "install", "pyinstaller"]):
                self.log_console("\n❌ 核心库安装失败，终止打包。\n")
                self.after(0, self._unlock_ui)
                return
                
            req_path = self.var_req.get()
            if req_path and os.path.exists(req_path):
                self.log_console(f"\n📥 正在沙盒中注入专属依赖 ({os.path.basename(req_path)})...\n")
                if not self._run_cmd_blocking([v_python, "-m", "pip", "install", "-r", req_path]):
                    self.log_console("\n❌ 专属依赖安装失败，终止打包。\n")
                    self.after(0, self._unlock_ui)
                    return

        self.log_console(f"\n🚀 [阶段 2/2] 启动打包引擎...\n{'-'*40}\n")
        cmd = [pyinstaller_exe, "-y"] 
        
        if self.var_onefile.get(): cmd.append("-F")
        if self.var_console.get(): cmd.append("-w") 
        if self.var_clean.get(): cmd.append("--clean")
        if self.var_upx.get(): cmd.append("--upx-dir=.") 
        if self.var_uac.get() and sys.platform == "win32": cmd.append("--uac-admin")
        
        if self.var_outdir.get(): cmd.extend(["--distpath", self.var_outdir.get()])
        if self.var_outname.get(): cmd.extend(["-n", self.var_outname.get()])
        if self.var_icon.get(): cmd.extend(["-i", self.var_icon.get()])
            
        add_data = self.var_add_data.get().strip()
        if add_data:
            for data in add_data.split(): cmd.extend(["--add-data", data])
                
        default_hidden = ["PIL._tkinter_finder"]
        for d_imp in default_hidden:
            cmd.extend(["--hidden-import", d_imp])
            
        hidden_imports = self.var_hidden_imports.get().strip()
        if hidden_imports:
            for imp in hidden_imports.replace(" ", "").split(","):
                if imp and imp not in default_hidden: 
                    cmd.extend(["--hidden-import", imp])
                
        exclude_modules = self.var_exclude_modules.get().strip()
        if exclude_modules:
            for exc in exclude_modules.replace(" ", "").split(","):
                if exc: cmd.extend(["--exclude-module", exc])
                
        cmd.append(self.var_script.get())
        
        success = self._run_cmd_blocking(cmd)
        
        if success:
            self.log_console("\n🎉 打包圆满完成！(生成的程序体积已得到极限优化)\n您可以点击左下角打开输出目录查看。\n")
        else:
            self.log_console("\n❌ 操作失败或被强制取消。\n")
            
        self.after(0, self._unlock_ui)

    def cancel_process(self):
        if self.process:
            self.process.terminate()
            self.log_console("\n🛑 正在强制终止进程...\n")

if __name__ == "__main__":
    multiprocessing.freeze_support()
    app = PyInstallerGUI()
    app.mainloop()
