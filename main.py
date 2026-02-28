import os
import sys
import shutil
import platform
import datetime
import threading
import subprocess
import ast
import tkinter as tk
from tkinter import ttk, filedialog
import customtkinter as ctk

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class PyInstallerGUI(ctk.CTk):
    def __init__(self):
        super().__init__()

        os_name = platform.system()
        self.title(f"Python 终极打包工作站 - {os_name} 适用版")
        self.geometry("950x850")
        self.minsize(900, 800)

        self.font_main = ctk.CTkFont(family="Microsoft YaHei", size=13)
        self.font_title = ctk.CTkFont(family="Microsoft YaHei", size=13, weight="bold")
        self.font_log = ctk.CTkFont(family="Consolas", size=12)

        self.create_widgets()
        self.safe_log(f"✅ 系统初始化完成，当前操作系统识别为: {os_name}")

    def create_widgets(self):
        self.tabview = ctk.CTkTabview(self, font=self.font_title)
        self.tabview.pack(fill="x", padx=15, pady=(10, 5))

        self.tabview.add("📄 基础配置")
        self.tabview.add("🌱 环境与清理")
        self.tabview.add("🚀 高级/专业特性")

        self._build_tab_basic()
        self._build_tab_env()
        self._build_tab_advanced()

        self.btn_pack = ctk.CTkButton(self, text="⚡ 开始极速打包", fg_color="#2E7D32", hover_color="#1B5E20", 
                                      font=ctk.CTkFont(family="Microsoft YaHei", size=16, weight="bold"), 
                                      height=45, command=self.start_pack_thread)
        self.btn_pack.pack(fill="x", padx=15, pady=10)

        log_frame = ctk.CTkFrame(self)
        log_frame.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        ctk.CTkLabel(log_frame, text="📜 构建日志面板", font=self.font_title).pack(anchor="w", padx=10, pady=(5, 0))
        
        self.textbox_log = ctk.CTkTextbox(log_frame, font=self.font_log, fg_color="#1e1e1e", text_color="#d4d4d4")
        self.textbox_log.pack(fill="both", expand=True, padx=10, pady=(5, 10))

    def _build_tab_basic(self):
        tab = self.tabview.tab("📄 基础配置")
        
        row1 = ctk.CTkFrame(tab, fg_color="transparent")
        row1.pack(fill="x", padx=10, pady=10)
        ctk.CTkLabel(row1, text="选择 Python 脚本：", font=self.font_main, width=130, anchor="w").pack(side="left")
        self.entry_script = ctk.CTkEntry(row1, font=self.font_main)
        self.entry_script.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(row1, text="浏览", width=80, font=self.font_main, command=self.browse_script).pack(side="left")

        row2 = ctk.CTkFrame(tab, fg_color="transparent")
        row2.pack(fill="x", padx=10, pady=(5, 10))
        target_name = "输出 EXE 名称：" if platform.system() == "Windows" else "输出程序名称："
        ctk.CTkLabel(row2, text=target_name, font=self.font_main, width=130, anchor="w").pack(side="left")
        self.entry_name = ctk.CTkEntry(row2, font=self.font_main)
        self.entry_name.pack(side="left", fill="x", expand=True, padx=(0, 20))
        
        ctk.CTkLabel(row2, text="程序图标 (.ico/.icns)：", font=self.font_main).pack(side="left", padx=(0, 10))
        self.entry_icon = ctk.CTkEntry(row2, font=self.font_main)
        self.entry_icon.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(row2, text="浏览", width=80, font=self.font_main, command=self.browse_icon).pack(side="left")

        row3 = ctk.CTkFrame(tab, fg_color="transparent")
        row3.pack(fill="x", padx=10, pady=10)
        self.var_single_file = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(row3, text="打包为单文件 (-F)", variable=self.var_single_file, font=self.font_main).pack(side="left", padx=(0, 30))

        self.var_noconsole = ctk.IntVar(value=1)
        ctk.CTkRadioButton(row3, text="控制台程序 (带黑框)", variable=self.var_noconsole, value=0, font=self.font_main).pack(side="left", padx=(0, 15))
        ctk.CTkRadioButton(row3, text="纯 GUI 程序 (-w 无黑框)", variable=self.var_noconsole, value=1, font=self.font_main).pack(side="left")

    def _build_tab_env(self):
        tab = self.tabview.tab("🌱 环境与清理")

        env_frame = ctk.CTkFrame(tab)
        env_frame.pack(fill="x", padx=10, pady=5)
        self.var_use_venv = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(env_frame, text="使用独立虚拟环境打包 (推荐：隔离系统庞杂库，减小体积)", variable=self.var_use_venv, font=self.font_title).pack(anchor="w", padx=15, pady=(15, 5))
        
        self.var_auto_deps = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(env_frame, text="智能推导并安装依赖 (调用 pipreqs 扫描代码)", variable=self.var_auto_deps, font=self.font_main).pack(anchor="w", padx=40, pady=(5, 15))

        clean_frame = ctk.CTkFrame(tab)
        clean_frame.pack(fill="x", padx=10, pady=10)
        self.var_clean_build = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(clean_frame, text="每次打包前彻底清理残留 (删除旧 build/dist/spec 及重建虚拟环境)", 
                        variable=self.var_clean_build, font=self.font_title, text_color="#EF5350").pack(anchor="w", padx=15, pady=15)

    def _build_tab_advanced(self):
        tab = self.tabview.tab("🚀 高级/专业特性")

        self.var_smart_fix = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(tab, text="开启 AST 智能防丢包修复 (自动补全 CustomTkinter/PyQt 等依赖)", variable=self.var_smart_fix, font=self.font_main).pack(anchor="w", padx=15, pady=(15, 10))

        if platform.system() == "Windows":
            self.var_uac = ctk.BooleanVar(value=False)
            ctk.CTkCheckBox(tab, text="请求管理员权限 (注入 UAC 盾牌，适合系统管理工具)", variable=self.var_uac, font=self.font_main).pack(anchor="w", padx=15, pady=10)
        else:
            self.var_uac = ctk.BooleanVar(value=False)

        splash_frame = ctk.CTkFrame(tab, fg_color="transparent")
        splash_frame.pack(fill="x", padx=10, pady=5)
        self.var_splash = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(splash_frame, text="添加加载启动屏 (Splash Image)：", variable=self.var_splash, font=self.font_main).pack(side="left", padx=(5, 10))
        self.entry_splash = ctk.CTkEntry(splash_frame, font=self.font_main, placeholder_text="选择 .png 或 .jpg")
        self.entry_splash.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(splash_frame, text="浏览", width=80, font=self.font_main, command=self.browse_splash).pack(side="left")

    def safe_log(self, message):
        self.after(0, self._append_log, message)

    def _append_log(self, message):
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.textbox_log.insert("end", f"[{timestamp}] {message}\n")
        self.textbox_log.see("end")

    def browse_script(self):
        filename = filedialog.askopenfilename(title="选择Python脚本", filetypes=[("Python Files", "*.py")])
        if filename:
            self.entry_script.delete(0, "end")
            self.entry_script.insert(0, filename)
            self.entry_name.delete(0, "end")
            self.entry_name.insert(0, os.path.splitext(os.path.basename(filename))[0])

    def browse_icon(self):
        ext = "*.ico" if platform.system() == "Windows" else "*.icns"
        filename = filedialog.askopenfilename(title="选择图标", filetypes=[("Icon Files", ext), ("All Files", "*.*")])
        if filename:
            self.entry_icon.delete(0, "end")
            self.entry_icon.insert(0, filename)

    def browse_splash(self):
        filename = filedialog.askopenfilename(title="选择启动屏图片", filetypes=[("Image Files", "*.png;*.jpg;*.jpeg")])
        if filename:
            self.entry_splash.delete(0, "end")
            self.entry_splash.insert(0, filename)
            self.var_splash.set(True)

    def start_pack_thread(self):
        script = self.entry_script.get()
        if not script:
            self.safe_log("❌ 请先选择要打包的脚本！")
            return
        self.btn_pack.configure(state="disabled", text="引擎运转中...")
        threading.Thread(target=self._pack_process, daemon=True).start()

    def _run_subprocess(self, cmd_list, cwd=None):
        creationflags = subprocess.CREATE_NO_WINDOW if platform.system() == 'Windows' else 0
        try:
            process = subprocess.Popen(cmd_list, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, cwd=cwd, creationflags=creationflags, encoding='utf-8', errors='replace')
            for line in process.stdout:
                self.safe_log(line.strip())
            process.wait()
            return process.returncode
        except Exception as e:
            self.safe_log(f"❌ 命令执行失败: {str(e)}")
            return 1

    def _cleanup_old_builds(self, work_dir, exe_name, venv_dir):
        self.safe_log("🧹 正在执行深度清理，扫除历史残留文件...")
        dirs_to_remove = [os.path.join(work_dir, "build"), os.path.join(work_dir, "dist")]
        if self.var_use_venv.get():
            dirs_to_remove.append(venv_dir)
            
        for d in dirs_to_remove:
            if os.path.exists(d):
                try:
                    shutil.rmtree(d)
                    self.safe_log(f"   已删除目录: {os.path.basename(d)}")
                except Exception as e:
                    self.safe_log(f"   ⚠️ 删除目录 {d} 失败: {e}")

        spec_file = os.path.join(work_dir, f"{exe_name}.spec")
        if os.path.exists(spec_file):
            try:
                os.remove(spec_file)
            except Exception:
                pass

    def _pack_process(self):
        try:
            script_path = self.entry_script.get()
            work_dir = os.path.dirname(script_path)
            exe_name = self.entry_name.get()
            venv_dir = os.path.join(work_dir, "smart_build_venv")
            
            if self.var_clean_build.get():
                self._cleanup_old_builds(work_dir, exe_name, venv_dir)

            if self.var_use_venv.get():
                python_exe = sys.executable
                if not os.path.exists(venv_dir):
                    self.safe_log("🌱 [环境] 正在初始化全新虚拟环境 (请耐心等待)...")
                    self._run_subprocess([python_exe, "-m", "venv", venv_dir])
                else:
                    self.safe_log("🌱 [环境] 检测到已有虚拟环境，直接复用。")

                if platform.system() == 'Windows':
                    active_python = os.path.join(venv_dir, "Scripts", "python")
                    active_pip = os.path.join(venv_dir, "Scripts", "pip")
                    active_pyinstaller = os.path.join(venv_dir, "Scripts", "pyinstaller")
                else:
                    active_python = os.path.join(venv_dir, "bin", "python")
                    active_pip = os.path.join(venv_dir, "bin", "pip")
                    active_pyinstaller = os.path.join(venv_dir, "bin", "pyinstaller")

                self._run_subprocess([active_python, "-m", "pip", "install", "--upgrade", "pip", "-q"])
                self._run_subprocess([active_pip, "install", "pyinstaller", "-q"])

                if self.var_auto_deps.get():
                    self.safe_log("🤖 [依赖] 调用 pipreqs 分析项目所需库...")
                    self._run_subprocess([active_pip, "install", "pipreqs", "-q"])
                    pipreqs_cmd = os.path.join(venv_dir, "Scripts" if platform.system() == 'Windows' else "bin", "pipreqs")
                    req_path = os.path.join(work_dir, "auto_requirements.txt")
                    self._run_subprocess([pipreqs_cmd, work_dir, "--force", "--savepath", req_path])
                    
                    if os.path.exists(req_path):
                        self.safe_log("⏳ [依赖] 正在安装业务所需模块，由于网络原因可能较慢...")
                        self._run_subprocess([active_pip, "install", "-r", req_path])
            else:
                self.safe_log("⚡ [环境] 警告：已关闭虚拟环境，将使用系统主环境直接打包！")
                active_pyinstaller = "pyinstaller"

            self.safe_log("🚀 正在构建最终打包参数...")
            cmd = [active_pyinstaller, "-y"]
            
            if exe_name: cmd.extend(["-n", exe_name])
            if self.var_single_file.get(): cmd.append("-F")
            if self.var_noconsole.get() == 1: cmd.append("-w")
            
            icon = self.entry_icon.get()
            if icon: cmd.extend(["-i", icon])

            if self.var_uac.get():
                cmd.append("--uac-admin")
                self.safe_log("🛡️ [特性] 已注入管理员权限申请 (UAC)")
            
            if self.var_splash.get() and self.entry_splash.get():
                cmd.extend(["--splash", self.entry_splash.get()])
                self.safe_log("🖼️ [特性] 已加入启动屏特效")

            if self.var_smart_fix.get():
                self.safe_log("🔍 [AST] 正在扫描代码漏洞，注入补丁...")
                try:
                    with open(script_path, "r", encoding="utf-8") as f:
                        tree = ast.parse(f.read(), filename=script_path)
                    for node in ast.walk(tree):
                        if isinstance(node, ast.Import):
                            for alias in node.names:
                                if "customtkinter" in alias.name: cmd.extend(["--collect-all", "customtkinter"])
                                if "pandas" in alias.name: cmd.extend(["--hidden-import", "pandas"])
                        elif isinstance(node, ast.ImportFrom) and node.module:
                            if "customtkinter" in node.module: cmd.extend(["--collect-all", "customtkinter"])
                except Exception as e:
                    self.safe_log(f"⚠️ AST扫描跳过: {e}")

            cmd.append(script_path)

            return_code = self._run_subprocess(cmd, cwd=work_dir)
            
            if return_code == 0:
                self.safe_log("🎉 [大功告成] 打包已完美完成！")
                dist_dir = os.path.join(work_dir, "dist")
                if platform.system() == "Windows":
                    os.startfile(dist_dir)
                elif platform.system() == "Darwin":
                    subprocess.call(["open", dist_dir])
                else:
                    subprocess.call(["xdg-open", dist_dir])
            else:
                self.safe_log("❌ [打包失败] 请检查上方日志。")

        except Exception as e:
            self.safe_log(f"❌ 发生致命异常: {str(e)}")
        finally:
            self.after(0, lambda: self.btn_pack.configure(state="normal", text="⚡ 开始极速打包"))

if __name__ == "__main__":
    app = PyInstallerGUI()
    app.mainloop()