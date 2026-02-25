import os
import sys
import threading
import subprocess
import shutil
import shlex
import time
import customtkinter as ctk
from tkinter import filedialog, END
from tkinterdnd2 import TkinterDnD, DND_FILES

# --- 让 CustomTkinter 支持完美拖拽 ---
class TkinterDnD_CTk(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

class PackagerApp(TkinterDnD_CTk):
    def __init__(self):
        super().__init__()
        self.title("Python脚本打包工具 - 智能守护终极版")
        self.geometry("860x920")
        self.minsize(800, 800)

        lbl_title = ctk.CTkLabel(self, text="Python脚本打包 “EXE” 工具", font=("Microsoft YaHei UI", 22, "bold"), text_color="#1f538d")
        lbl_title.pack(pady=(15, 10))

        # ==================== 1. 文件与配置 ====================
        self.frame_files = ctk.CTkFrame(self, corner_radius=10)
        self.frame_files.pack(pady=5, padx=15, fill="x")
        ctk.CTkLabel(self.frame_files, text="📁 核心配置 (支持拖拽文件输入)", font=("Microsoft YaHei UI", 15, "bold")).grid(row=0, column=0, columnspan=3, padx=15, pady=8, sticky="w")

        self.entry_name = ctk.CTkEntry(self.frame_files, placeholder_text="可选: 自动提取或自定义程序名 (如: 我的软件)")
        
        self.entry_script = self.create_file_row(self.frame_files, "选择脚本(*):", 1, "必须: 支持拖拽主 .py 文件", self.browse_script)
        self.entry_req = self.create_file_row(self.frame_files, "依赖文件:", 2, "可选: requirements.txt (自动安装依赖)", self.browse_req)
        
        ctk.CTkLabel(self.frame_files, text="程序命名:").grid(row=3, column=0, padx=15, pady=6, sticky="e")
        self.entry_name.grid(row=3, column=1, columnspan=2, padx=5, pady=6, sticky="ew")

        ctk.CTkLabel(self.frame_files, text="额外参数:").grid(row=4, column=0, padx=15, pady=6, sticky="e")
        self.entry_extra = ctk.CTkEntry(self.frame_files, placeholder_text="可选: 输入额外的指令 (有了智能修复，通常这里可留空)")
        self.entry_extra.grid(row=4, column=1, columnspan=2, padx=5, pady=6, sticky="ew")
        
        ctk.CTkFrame(self.frame_files, height=2, fg_color="gray80").grid(row=5, column=0, columnspan=3, sticky="ew", padx=15, pady=10)

        self.entry_icon = self.create_file_row(self.frame_files, "程序图标:", 6, "可选: .ico 或 .icns 格式", self.browse_icon)
        self.entry_outdir = self.create_file_row(self.frame_files, "输出目录:", 7, "可选: 默认当前目录下的 dist 文件夹", self.browse_dir)
        self.entry_adddata = self.create_file_row(self.frame_files, "附加资源:", 8, "可选: 需要打包的额外文件/文件夹", self.browse_adddata)

        # ==================== 2. 打包选项 ====================
        self.frame_opts = ctk.CTkFrame(self, corner_radius=10)
        self.frame_opts.pack(pady=10, padx=15, fill="x")
        
        ctk.CTkLabel(self.frame_opts, text="⚙️ 环境与智能选项", font=("Microsoft YaHei UI", 15, "bold")).pack(anchor="w", padx=15, pady=(10, 5))

        grid_frame = ctk.CTkFrame(self.frame_opts, fg_color="transparent")
        grid_frame.pack(fill="x", padx=15, pady=5)

        self.var_onefile = ctk.BooleanVar(value=True)
        self.var_noconsole = ctk.BooleanVar(value=True)
        self.var_admin = ctk.BooleanVar(value=False)
        self.var_venv = ctk.BooleanVar(value=True)
        self.var_auto_fix = ctk.BooleanVar(value=True) 
        self.var_open_folder = ctk.BooleanVar(value=True)

        ctk.CTkCheckBox(grid_frame, text="单文件模式 (-F)", variable=self.var_onefile).grid(row=0, column=0, padx=(0, 20), pady=10, sticky="w")
        ctk.CTkCheckBox(grid_frame, text="隐藏控制台 (-w)", variable=self.var_noconsole).grid(row=0, column=1, padx=(0, 20), pady=10, sticky="w")
        ctk.CTkCheckBox(grid_frame, text="请求管理员权限", variable=self.var_admin).grid(row=0, column=2, padx=(0, 20), pady=10, sticky="w")
        
        ctk.CTkCheckBox(grid_frame, text="🟢 每次新建干净环境", variable=self.var_venv, text_color="green").grid(row=1, column=0, padx=(0, 20), pady=10, sticky="w")
        ctk.CTkCheckBox(grid_frame, text="🤖 智能防报错修复", variable=self.var_auto_fix, text_color="#d97706").grid(row=1, column=1, padx=(0, 20), pady=10, sticky="w")
        ctk.CTkCheckBox(grid_frame, text="📂 打包完打开目录", variable=self.var_open_folder, text_color="#1f538d").grid(row=1, column=2, padx=(0, 20), pady=10, sticky="w")

        adv_frame = ctk.CTkFrame(self.frame_opts, fg_color="transparent")
        adv_frame.pack(fill="x", padx=15, pady=(5, 15))
        ctk.CTkLabel(adv_frame, text="🚫 排除模块:").pack(side="left", padx=(0, 10))
        self.entry_exclude = ctk.CTkEntry(adv_frame, placeholder_text="输入要排除的库名，用逗号分隔 (如: numpy,pandas)")
        self.entry_exclude.pack(side="left", fill="x", expand=True)

        # ==================== 3. 按钮区 ====================
        self.frame_btns = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_btns.pack(pady=5, padx=20, fill="x")

        self.btn_pack = ctk.CTkButton(self.frame_btns, text="🚀 开始智能隔离打包", font=("Microsoft YaHei UI", 16, "bold"), fg_color="#28a745", hover_color="#218838", height=45, command=self.start_pack)
        self.btn_pack.pack(side="left", expand=True, fill="x", padx=(0, 10))

        ctk.CTkButton(self.frame_btns, text="🗑️ 清空日志", font=("Microsoft YaHei UI", 16), fg_color="#dc3545", hover_color="#c82333", height=45, width=120, command=self.clear_log).pack(side="right")

        # ==================== 4. 日志区 ====================
        self.frame_log = ctk.CTkFrame(self, corner_radius=10)
        self.frame_log.pack(pady=10, padx=15, fill="both", expand=True) 
        self.txt_log = ctk.CTkTextbox(self.frame_log, font=("Consolas", 12))
        self.txt_log.pack(padx=10, pady=10, fill="both", expand=True)

    def create_file_row(self, parent, label_text, row, placeholder, btn_cmd):
        ctk.CTkLabel(parent, text=label_text).grid(row=row, column=0, padx=15, pady=5, sticky="e")
        entry = ctk.CTkEntry(parent, placeholder_text=placeholder)
        entry.grid(row=row, column=1, padx=5, pady=5, sticky="ew")
        parent.columnconfigure(1, weight=1)
        ctk.CTkButton(parent, text="浏览", width=70, command=btn_cmd).grid(row=row, column=2, padx=15, pady=5)
        entry.drop_target_register(DND_FILES)
        entry.dnd_bind('<<Drop>>', lambda e: self.on_drop(e, entry))
        return entry

    def on_drop(self, event, entry_widget):
        file_path = event.data.strip('{}')
        entry_widget.delete(0, END)
        entry_widget.insert(0, file_path)
        if getattr(self, 'entry_script', None) and entry_widget == self.entry_script:
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            self.entry_name.delete(0, END)
            self.entry_name.insert(0, base_name)

    def browse_script(self):
        f = filedialog.askopenfilename(filetypes=[("Python Files", "*.py")])
        if f: 
            self.entry_script.delete(0, END)
            self.entry_script.insert(0, f)
            base_name = os.path.splitext(os.path.basename(f))[0]
            self.entry_name.delete(0, END)
            self.entry_name.insert(0, base_name)

    def browse_req(self):
        f = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
        if f: self.entry_req.delete(0, END); self.entry_req.insert(0, f)

    def browse_icon(self):
        f = filedialog.askopenfilename(filetypes=[("Icon Files", "*.ico;*.icns")])
        if f: self.entry_icon.delete(0, END); self.entry_icon.insert(0, f)

    def browse_dir(self):
        d = filedialog.askdirectory()
        if d: self.entry_outdir.delete(0, END); self.entry_outdir.insert(0, d)

    def browse_adddata(self):
        f = filedialog.askopenfilename()
        if f: self.entry_adddata.delete(0, END); self.entry_adddata.insert(0, f)

    def log(self, message):
        self.txt_log.insert(END, message + "\n")
        self.txt_log.see(END)

    def clear_log(self):
        self.txt_log.delete("1.0", END)

    def get_system_python(self):
        if getattr(sys, 'frozen', False):
            if sys.platform == "darwin":
                good_paths = [
                    "/usr/local/bin/python3",
                    "/Library/Frameworks/Python.framework/Versions/Current/bin/python3"
                ]
                for p in good_paths:
                    if os.path.exists(p): return p
                return shutil.which("python3") or shutil.which("python")
            else:
                return shutil.which("python") or shutil.which("python3")
        else:
            return sys.executable or shutil.which("python3") or shutil.which("python")

    def get_clean_env(self):
        env = os.environ.copy()
        env.pop('_MEIPASS2', None)
        env.pop('PYARMOR_LICENSE', None)
        env.pop('PYTHONPATH', None)
        env.pop('LD_LIBRARY_PATH', None)
        env.pop('DYLD_LIBRARY_PATH', None)
        return env

    def run_cmd_with_log(self, cmd_list, cwd=None):
        startupinfo = None
        if os.name == 'nt':
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            
        try:
            process = subprocess.Popen(
                cmd_list, 
                stdout=subprocess.PIPE, 
                stderr=subprocess.STDOUT, 
                text=True, 
                encoding='utf-8', 
                errors='replace', 
                startupinfo=startupinfo, 
                cwd=cwd, 
                env=self.get_clean_env()
            )
            for line in process.stdout:
                self.log(line.strip())
            process.wait()
            return process.returncode == 0
        except Exception as e:
            self.log(f"执行命令时出错: {e}")
            return False

    def open_output_folder(self, path):
        try:
            if not os.path.exists(path): return
            if os.name == 'nt':
                os.startfile(path)
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', path])
            else:
                subprocess.Popen(['xdg-open', path])
        except Exception as e:
            self.log(f"无法自动打开文件夹: {e}")

    def bring_window_to_front(self):
        try:
            if self.state() == 'iconic':
                self.deiconify() 
            self.attributes('-topmost', True)
            self.focus_force()
            self.update()
            self.attributes('-topmost', False)
        except Exception as e:
            pass

    def smart_analyze_dependencies(self, script_path, req_path):
        auto_args_set = set()
        content = ""
        
        if script_path and os.path.exists(script_path):
            try:
                with open(script_path, 'r', encoding='utf-8', errors='ignore') as f:
                    content += f.read()
            except Exception: pass
            
        if req_path and os.path.exists(req_path):
            try:
                with open(req_path, 'r', encoding='utf-8', errors='ignore') as f:
                    content += "\n" + f.read()
            except Exception: pass

        if "ttkbootstrap" in content:
            auto_args_set.add(("--collect-all", "ttkbootstrap"))
            auto_args_set.add(("--hidden-import", "PIL._tkinter_finder"))
            
        if "customtkinter" in content:
            auto_args_set.add(("--collect-all", "customtkinter"))
            auto_args_set.add(("--hidden-import", "PIL._tkinter_finder"))

        if "PIL" in content or "Pillow" in content or "pillow" in content:
            auto_args_set.add(("--hidden-import", "PIL._tkinter_finder"))
            
        if "tkinterdnd2" in content:
            auto_args_set.add(("--collect-all", "tkinterdnd2"))
            
        if "pyttsx3" in content:
            auto_args_set.add(("--hidden-import", "pyttsx3.drivers"))
            auto_args_set.add(("--hidden-import", "pyttsx3.drivers.sapi5"))
            auto_args_set.add(("--hidden-import", "pyttsx3.drivers.nsss"))
            auto_args_set.add(("--hidden-import", "pyttsx3.drivers.dummy"))
            
        if "pandas" in content:
            auto_args_set.add(("--hidden-import", "pandas._libs.tslibs.timedeltas"))

        # 🌟 修改点：放弃容易失败的 --collect-all，仅做基本引入，重任交给后面的物理外挂
        if "azure.cognitiveservices.speech" in content or "azure" in content:
            auto_args_set.add(("--hidden-import", "azure.cognitiveservices.speech"))

        final_args = []
        for flag, val in auto_args_set:
            final_args.extend([flag, val])
            
        return final_args

    def start_pack(self):
        self.btn_pack.configure(state="disabled", text="⏳ 打包进行中 (请勿关闭)...")
        self.log("="*60)
        self.log("🚀 开始全自动智能打包流程...")
        threading.Thread(target=self.orchestrate_packaging, daemon=True).start()

    def orchestrate_packaging(self):
        try:
            script = self.entry_script.get().strip()
            if not script or not os.path.exists(script):
                self.log("❌ 错误: 找不到指定的 Python 脚本！")
                return

            req_file = self.entry_req.get().strip()
            app_name = self.entry_name.get().strip()
            script_dir = os.path.dirname(script)
            
            sys_py = self.get_system_python()
            if not sys_py:
                self.log("❌ 致命错误: 系统环境中找不到 Python！")
                return

            run_py = sys_py

            if self.var_venv.get():
                venv_dir = os.path.join(script_dir, ".pack_venv")
                self.log(f"📦 启用纯净虚拟环境。目标位置: {venv_dir}")
                
                if os.path.exists(venv_dir):
                    self.log("🧹 发现历史残留的虚拟环境，正在执行深度清理，请稍候...")
                    for _ in range(3):
                        try:
                            shutil.rmtree(venv_dir, ignore_errors=True)
                            if not os.path.exists(venv_dir): break
                            time.sleep(1)
                        except: pass
                    
                    if os.path.exists(venv_dir):
                        self.log("⚠️ 警告：无法彻底删除旧环境（可能被占用），将尝试直接覆盖。")
                    else:
                        self.log("✨ 历史环境清理完毕！")

                if os.name == 'nt':
                    venv_py = os.path.join(venv_dir, "Scripts", "python.exe")
                else:
                    venv_py = os.path.join(venv_dir, "bin", "python")

                self.log(f"⏳ 正在从零创建全新的隔离虚拟环境...")
                success = self.run_cmd_with_log([sys_py, "-m", "venv", venv_dir], cwd=script_dir)
                if not success or not os.path.exists(venv_py):
                    self.log("❌ 虚拟环境创建失败！")
                    return
                self.log("✅ 纯净虚拟环境创建成功！")

                run_py = venv_py 

                if req_file and os.path.exists(req_file):
                    self.log(f"⬇️ 正在从 {os.path.basename(req_file)} 挂载全新依赖...")
                    self.run_cmd_with_log([run_py, "-m", "pip", "install", "-r", req_file, "--disable-pip-version-check"], cwd=script_dir)

                self.log("⬇️ 正在为当前环境安装 PyInstaller 引擎...")
                self.run_cmd_with_log([run_py, "-m", "pip", "install", "pyinstaller", "--disable-pip-version-check"], cwd=script_dir)

            self.log(f"\n⚙️ 环境部署就绪，开始执行构建...")
            
            cmd = [run_py, "-m", "PyInstaller", "--noconfirm", "--clean"]

            is_macos_app = sys.platform == "darwin" and self.var_noconsole.get()
            if self.var_onefile.get() and not is_macos_app: 
                cmd.append("-F")
            elif is_macos_app and self.var_onefile.get():
                self.log("🍎 [Mac专属防呆] 侦测到正在生成 macOS .app 包，自动禁用冲突的单文件(-F)模式。")

            if self.var_noconsole.get(): cmd.append("-w")
            if self.var_admin.get(): cmd.append("--uac-admin")
            
            if app_name: cmd.extend(["-n", app_name])

            icon = self.entry_icon.get().strip()
            if icon: cmd.extend(["-i", icon])

            outdir = self.entry_outdir.get().strip()
            final_outdir = outdir if outdir else os.path.join(script_dir, "dist")
            cmd.extend(["--distpath", final_outdir])

            adddata = self.entry_adddata.get().strip()
            if adddata: 
                sep = ";" if os.name == 'nt' else ":"
                cmd.extend(["--add-data", f"{adddata}{sep}."])

            excludes = self.entry_exclude.get().strip()
            if excludes:
                for mod in excludes.split(","):
                    cmd.extend(["--exclude-module", mod.strip()])

            if self.var_auto_fix.get():
                self.log("🤖 正在进行代码深度扫描，寻找常见报错库...")
                smart_fixes = self.smart_analyze_dependencies(script, req_file)
                if smart_fixes:
                    self.log(f"✨ 检测到易错库，已自动注入免疫补丁: {' '.join(smart_fixes)}")
                    cmd.extend(smart_fixes)
                else:
                    self.log("✨ 扫描完毕，代码很干净，无需补丁。")
                    
                # ================= 🌟 物理寻址外挂防御体系 =================
                self.log("🤖 [动态探测] 正在扫描隐蔽的 C++ 底层依赖库...")
                check_code = "try:\n import azure.cognitiveservices.speech as az\n print(az.__path__[0])\nexcept:\n pass"
                try:
                    res = subprocess.run([run_py, "-c", check_code], capture_output=True, text=True, env=self.get_clean_env())
                    if res.returncode == 0 and res.stdout.strip():
                        az_path = res.stdout.strip()
                        sep = ";" if os.name == 'nt' else ":"
                        # 物理强制将这三个平台的库全都绑进去，不管三七二十一
                        cmd.extend(["--add-binary", f"{az_path}/*.dll{sep}azure/cognitiveservices/speech"])
                        cmd.extend(["--add-binary", f"{az_path}/*.so{sep}azure/cognitiveservices/speech"])
                        cmd.extend(["--add-binary", f"{az_path}/*.dylib{sep}azure/cognitiveservices/speech"])
                        self.log("✨ [终极防御] 成功定位并物理提取 Azure C++ 核心动态库，已强行捆绑至打包配方！")
                except Exception:
                    pass
                # ========================================================

            extra = self.entry_extra.get().strip()
            if extra:
                cmd.extend(shlex.split(extra))

            cmd.append(script)
            
            success = self.run_cmd_with_log(cmd, cwd=script_dir)
            
            if success:
                self.log(f"\n✅ 打包大功告成！文件已输出至: {final_outdir}")
                
                target_name = app_name if app_name else os.path.splitext(os.path.basename(script))[0]
                
                spec_path = os.path.join(script_dir, f"{target_name}.spec")
                if os.path.exists(spec_path):
                    try:
                        os.remove(spec_path)
                        self.log("🧹 [无痕清理] 已自动删除临时的 .spec 配置文件。")
                    except Exception:
                        pass
                
                if sys.platform == "darwin" and self.var_noconsole.get():
                    raw_folder_path = os.path.join(final_outdir, target_name)
                    app_bundle_path = os.path.join(final_outdir, f"{target_name}.app")
                    
                    if os.path.exists(app_bundle_path) and os.path.exists(raw_folder_path) and os.path.isdir(raw_folder_path):
                        try:
                            shutil.rmtree(raw_folder_path, ignore_errors=True)
                            self.log("🧹 [无痕清理] 已自动为您删除 macOS 底层多余的同名文件夹，输出目录仅保留纯净的 .app 包！")
                        except Exception:
                            pass

                if self.var_open_folder.get():
                    self.log("📂 正在为您打开输出文件夹...")
                    self.open_output_folder(final_outdir)
            else:
                self.log("\n❌ 打包失败，请向上滚动查看红色错误日志。")

        except Exception as e:
            self.log(f"\n❌ 发生严重异常: {str(e)}")
            
        finally:
            self.btn_pack.configure(state="normal", text="🚀 开始智能隔离打包")
            self.log("\n✨ 任务彻底结束，工具已释放！")
            self.bring_window_to_front()

if __name__ == "__main__":
    app = PackagerApp()
    app.mainloop()
