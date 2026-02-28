import os
import sys
import re
import shutil
import hashlib
import subprocess
import platform
import configparser
import time
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import List, Optional, Tuple, Callable

from PySide6.QtCore import Qt, QObject, Signal, QThread, QSize, QUrl
from PySide6.QtGui import QAction, QFont, QGuiApplication, QImage, QPainter, QColor, QLinearGradient, QDesktopServices, QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFileDialog, QMessageBox,
    QHBoxLayout, QVBoxLayout, QFormLayout, QLineEdit, QPushButton, QTextEdit,
    QComboBox, QCheckBox, QLabel, QPlainTextEdit, QGroupBox,
    QSplitter, QScrollArea, QTabWidget, QFrame, QSpinBox
)

# -----------------------------
# 全局常量与智能免疫规则库
# -----------------------------
APP_NAME = "MultiPlatform Py Packer"
APP_VERSION = "3.5.0 Ultimate"  # 🚀 新增 ttkbootstrap 免疫，强化环境隔离，保持每次开启纯净
BUILD_ROOT_NAME = ".mpbuild"
DEFAULT_OUTPUT_DIRNAME = "dist_out"

IS_WIN = sys.platform.startswith("win")
IS_MAC = sys.platform == "darwin"
IS_LINUX = sys.platform.startswith("linux")

# 🧠 终极智能免疫知识库
SMART_HEURISTICS = {
    "ttkbootstrap": {
        "collect_all": ["ttkbootstrap"], 
        "hidden_imports": ["PIL._tkinter_finder"]
    },
    "azure-cognitiveservices-speech": {"collect_all": ["azure.cognitiveservices.speech"]},
    "customtkinter": {"collect_all": ["customtkinter"], "hidden_imports": ["PIL._tkinter_finder"], "nuitka_plugins": ["tk-inter"]},
    "pandas": {"hidden_imports": ["pandas._libs.tslibs.timedeltas"], "nuitka_plugins": ["numpy"]},
    "numpy": {"nuitka_plugins": ["numpy"]},
    "opencv-python": {"collect_all": ["cv2"]},
    "scipy": {"collect_all": ["scipy"]},
    "matplotlib": {"collect_all": ["matplotlib"]},
    "playwright": {"collect_all": ["playwright"]},
    "moviepy": {"hidden_imports": ["pkg_resources.py2_warn"]},
    "pyqt5": {"collect_all": ["PyQt5"]},
    "pyside6": {"collect_all": ["PySide6"], "nuitka_plugins": ["pyside6"]},
    "tiktoken": {"collect_all": ["tiktoken"]},
    "torch": {"collect_all": ["torch", "torchaudio", "torchvision"]},
    "transformers": {"collect_all": ["transformers"]},
    "soundfile": {"collect_all": ["soundfile"]},
    "librosa": {"collect_all": ["librosa"]}
}

# -----------------------------
# 基础工具
# -----------------------------
def normpath(p: str) -> str: return str(Path(p).expanduser().resolve())
def sha1_text(s: str) -> str: return hashlib.sha1(s.encode("utf-8")).hexdigest()
def safe_mkdir(p: Path): p.mkdir(parents=True, exist_ok=True)
def rm_tree(p: Path): shutil.rmtree(p, ignore_errors=True) if p.exists() else None
def which_in_venv(venv_dir: Path, exe_name: str) -> Path: return venv_dir / "Scripts" / (exe_name + ".exe") if IS_WIN else venv_dir / "bin" / exe_name
def python_in_venv(venv_dir: Path) -> Path: return which_in_venv(venv_dir, "python")
def quote_arg(a: str) -> str: return f'"{a}"' if " " in a or "\t" in a else a
def format_cmd(cmd: List[str]) -> str: return " ".join(quote_arg(x) for x in cmd)
def split_lines(s: str) -> List[str]: return [x.strip() for x in s.splitlines() if x.strip()]
def home_desktop_dir() -> Path: return Path.home() / "Desktop" if (Path.home() / "Desktop").exists() else Path.home()
def is_frozen_app() -> bool: return bool(getattr(sys, "frozen", False))

def ensure_writable_directory(target: Path, fallback: Path) -> Tuple[Path, Optional[str]]:
    try:
        safe_mkdir(target)
        (target / ".__write_probe__").write_text("ok", encoding="utf-8")
        (target / ".__write_probe__").unlink(missing_ok=True)
        return target, None
    except Exception:
        safe_mkdir(fallback)
        return fallback, f"目录不可写，已切换至桌面：{fallback}"

def sanitize_name(s: str) -> str:
    s = re.sub(r"[^A-Za-z0-9_\-\.]+", "_", s.strip())
    return s.strip("._-") or "MyApp"

def guess_app_name(project_dir: Path) -> str:
    setup_cfg = project_dir / "setup.cfg"
    if setup_cfg.exists():
        cp = configparser.ConfigParser()
        try:
            cp.read(setup_cfg, encoding="utf-8")
            if cp.has_section("metadata") and cp.has_option("metadata", "name"):
                return sanitize_name(cp.get("metadata", "name"))
        except Exception: pass
    return sanitize_name(project_dir.name)

# -----------------------------
# Host Python (强校验)
# -----------------------------
def _is_valid_host_python(py: Path) -> bool:
    if not py.exists() or py.is_dir(): return False
    if is_frozen_app() and py.resolve() == Path(sys.executable).resolve(): return False
    if IS_MAC and "Contents/MacOS" in py.parts: return False
    try:
        r = subprocess.run([str(py), "-c", "print('PYTHON_CORE_OK')"], capture_output=True, text=True, timeout=3)
        return "PYTHON_CORE_OK" in r.stdout
    except Exception:
        return False

def _rank_macos_python(p: Path) -> int:
    try:
        parts = p.parts
        if "Versions" in parts:
            ver_str = parts[parts.index("Versions") + 1]
            if "3.12" in ver_str or "3.11" in ver_str: return 10
            if "3.10" in ver_str or "3.9" in ver_str: return 8
            if "3." in ver_str: return 5
    except: pass
    return 0

def find_host_python() -> Path:
    candidates = []
    if IS_MAC:
        fw_base = Path("/Library/Frameworks/Python.framework/Versions")
        if fw_base.exists():
            cands = [v / "bin" / "python3" for v in fw_base.iterdir() if (v / "bin" / "python3").exists()]
            cands.sort(key=_rank_macos_python, reverse=True)
            candidates.extend(cands)
        candidates += [Path("/opt/homebrew/bin/python3"), Path("/usr/local/bin/python3")]

    for name in ("python3", "python"):
        if p := shutil.which(name): candidates.append(Path(p))

    if IS_MAC: candidates.append(Path("/usr/bin/python3"))
    if getattr(sys, "_base_executable", None): candidates.append(Path(sys._base_executable))

    seen = set()
    for c in candidates:
        cs = str(c.resolve())
        if cs not in seen:
            seen.add(cs)
            if _is_valid_host_python(c):
                return c
    raise RuntimeError("未在系统中探测到有效的 Python 3 环境，请手动浏览选择。")

# -----------------------------
# 子进程执行 (支持硬核中断)
# -----------------------------
class BuildCancelledError(Exception): pass

def run_subprocess_stream(cmd: List[str], cwd: Optional[Path], env: Optional[dict], log_cb: Callable[[str], None], check_cancel: Callable[[], bool]):
    proc = subprocess.Popen(
        cmd, cwd=str(cwd) if cwd else None, env=env,
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        text=True, bufsize=1, universal_newlines=True
    )
    last_out = time.time()
    try:
        while True:
            if check_cancel():
                proc.kill()
                raise BuildCancelledError("用户主动终止了构建进程。")
            line = proc.stdout.readline()
            if line:
                last_out = time.time()
                log_cb(line.rstrip("\n"))
            else:
                if proc.poll() is not None: break
                if time.time() - last_out >= 8.0:
                    last_out = time.time()
                    log_cb("[INFO] 正在执行中...")
                time.sleep(0.05)
    except BuildCancelledError:
        proc.wait(); raise
    except Exception as e:
        proc.kill(); raise RuntimeError(f"进程异常: {e}")
    proc.wait()
    return proc.returncode

def generate_default_splash_png(path: Path, app_name: str):
    w, h = 780, 460
    img = QImage(w, h, QImage.Format_ARGB32); img.fill(QColor("#FFFFFF"))
    painter = QPainter(img); painter.setRenderHint(QPainter.Antialiasing, True)
    grad = QLinearGradient(0, 0, w, h); grad.setColorAt(0.0, QColor("#EEF2FF")); grad.setColorAt(1.0, QColor("#F5F7FB"))
    painter.fillRect(0, 0, w, h, grad)
    painter.setPen(QColor("#D6DCE8")); painter.setBrush(QColor("#FFFFFF"))
    painter.drawRoundedRect(60, 70, w - 120, h - 140, 18, 18)
    painter.setPen(QColor("#0F172A")); f = QFont("PingFang SC" if IS_MAC else "Arial", 26); f.setBold(True); painter.setFont(f)
    painter.drawText(90, 150, app_name)
    painter.setPen(QColor("#64748B")); f2 = QFont("PingFang SC" if IS_MAC else "Arial", 13); painter.setFont(f2)
    painter.drawText(90, 185, "正在启动，请稍候…")
    painter.end()
    safe_mkdir(path.parent); img.save(str(path), "PNG")

# -----------------------------
# 自定义 UI 组件 (支持拖拽)
# -----------------------------
class DropLineEdit(QLineEdit):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setAcceptDrops(True)
    def dragEnterEvent(self, e: QDragEnterEvent): e.accept() if e.mimeData().hasUrls() else e.ignore()
    def dropEvent(self, e: QDropEvent): self.setText(normpath(e.mimeData().urls()[0].toLocalFile()))

# -----------------------------
# BuildConfig
# -----------------------------
@dataclass
class BuildConfig:
    project_dir: str; entry_script: str; output_dir: str; app_name: str
    builder: str; onefile: bool; windowed: bool; icon_path: str; host_python: str
    splash_enable: bool; splash_path: str; optimize_level: int
    use_requirements: bool; requirements_path: str; pip_mirror: str; upgrade_pip: bool; force_reinstall: bool; no_pip_cache: bool
    hidden_imports: List[str]; collect_all: List[str]; add_data: List[str]; nuitka_plugins: List[str]; use_upx: bool
    clean_build_dirs: bool; purge_venv: bool; purge_pyinstaller_cache: bool; extra_args: str

def default_build_config() -> BuildConfig:
    return BuildConfig(
        project_dir="", entry_script="", output_dir="", app_name="MyApp",
        builder="pyinstaller", onefile=True, windowed=True, icon_path="", host_python="",
        splash_enable=not IS_MAC, splash_path="", optimize_level=1,
        use_requirements=True, requirements_path="", pip_mirror="",
        upgrade_pip=True, force_reinstall=False, no_pip_cache=False,
        hidden_imports=[], collect_all=[], add_data=[], nuitka_plugins=[], use_upx=False,
        clean_build_dirs=True, purge_venv=False, purge_pyinstaller_cache=False, extra_args=""
    )

# -----------------------------
# 构建线程
# -----------------------------
class BuildWorker(QObject):
    log = Signal(str); done = Signal(bool, str, str); stage = Signal(str)

    def __init__(self, cfg: BuildConfig):
        super().__init__()
        self.cfg = cfg
        self._is_cancelled = False

    def cancel(self): self._is_cancelled = True
    def _check_cancel(self) -> bool: return self._is_cancelled
    def _emit(self, s: str): self.log.emit(s)
    def _header(self, title: str): self._emit(f"\n{'='*80}\n[{title}]\n{'='*80}")

    def run(self):
        out_dir = ""
        try:
            out_dir = self._run_impl()
            if not self._is_cancelled: self.done.emit(True, "构建成功 ✅", out_dir)
        except BuildCancelledError as e: self._emit(f"[STOP] {e}"); self.done.emit(False, "构建已终止 🛑", out_dir)
        except Exception as e: self._emit(f"[FATAL] {e}"); self.done.emit(False, f"构建失败：{e}", out_dir)

    def _run_cmd(self, cmd: List[str], cwd: Path, msg: str = ""):
        if msg: self._emit(msg)
        
        # 🛡️ 终极环境隔离：防止打包器自身的运行环境污染目标项目的编译环境
        clean_env = os.environ.copy()
        for key in ["PYTHONPATH", "PYTHONHOME", "DYLD_LIBRARY_PATH", "LD_LIBRARY_PATH"]:
            clean_env.pop(key, None)
            
        if cmd and "python" in Path(cmd[0]).name.lower():
            venv_bin_dir = str(Path(cmd[0]).parent)
            clean_env["PATH"] = f"{venv_bin_dir}{os.pathsep}{clean_env.get('PATH', '')}"

        if run_subprocess_stream(cmd, cwd, clean_env, self._emit, self._check_cancel) != 0:
            raise RuntimeError(f"命令执行失败: {format_cmd(cmd)}")

    def _apply_smart_heuristics(self, freeze_path: Path):
        self.stage.emit("智能诊断"); self._header("🤖 智能依赖诊断与自动免疫")
        try: freeze_text = freeze_path.read_text(encoding="utf-8").lower()
        except: return self._emit("[WARN] 无法读取依赖，跳过智能诊断。")

        applied = []
        for pkg, rules in SMART_HEURISTICS.items():
            if re.search(rf"^{pkg}(==|>=|<=|~=|>|<|$)", freeze_text, re.MULTILINE):
                for key, target_list in [("collect_all", self.cfg.collect_all), ("hidden_imports", self.cfg.hidden_imports)]:
                    if key in rules:
                        for item in rules[key]:
                            if item not in target_list: target_list.append(item); applied.append(f"[{key}] {item} (触发: {pkg})")
                if "nuitka_plugins" in rules and self.cfg.builder == "nuitka":
                    for plg in rules["nuitka_plugins"]:
                        if plg not in self.cfg.nuitka_plugins: self.cfg.nuitka_plugins.append(plg); applied.append(f"[nuitka_plugin] {plg} (触发: {pkg})")

        if applied:
            self._emit("✅ 已自动注入以下免疫补丁：")
            for p in applied: self._emit("  - " + p)
        else: self._emit("未检测到高危依赖。")

    def _run_impl(self) -> str:
        cfg = self.cfg
        proj_dir, entry_py = Path(normpath(cfg.project_dir)), Path(normpath(cfg.entry_script))
        if not proj_dir.exists() or not entry_py.exists(): raise RuntimeError("项目目录或入口脚本不存在")

        if IS_MAC and cfg.windowed: cfg.onefile = False

        out_dir, _ = ensure_writable_directory(Path(normpath(cfg.output_dir)) if cfg.output_dir else proj_dir / DEFAULT_OUTPUT_DIRNAME, home_desktop_dir() / DEFAULT_OUTPUT_DIRNAME)
        final_export = str(out_dir / f"{cfg.app_name}_export")
        
        host_py = Path(cfg.host_python) if cfg.host_python and _is_valid_host_python(Path(cfg.host_python)) else find_host_python()
        work_root = proj_dir / BUILD_ROOT_NAME / sha1_text(json.dumps(asdict(cfg), ensure_ascii=False, sort_keys=True))[:12]
        venv_dir, build_dir, dist_dir, logs_dir = work_root/"venv", work_root/"build", work_root/"dist", work_root/"logs"
        for d in (work_root, build_dir, dist_dir, logs_dir): safe_mkdir(d)

        self._header("环境准备")
        if cfg.clean_build_dirs: rm_tree(build_dir); rm_tree(dist_dir); safe_mkdir(build_dir); safe_mkdir(dist_dir)
        if cfg.purge_venv: rm_tree(venv_dir)

        if not venv_dir.exists(): self._run_cmd([str(host_py), "-m", "venv", str(venv_dir)], proj_dir, f"基于 {host_py} 创建虚拟环境...")
        vpy = python_in_venv(venv_dir)
        pip_cmd = [str(vpy), "-m", "pip", "install"]
        if cfg.pip_mirror: pip_cmd += ["-i", cfg.pip_mirror]
        if cfg.no_pip_cache: pip_cmd += ["--no-cache-dir"]

        self.stage.emit("安装依赖")
        if cfg.upgrade_pip: self._run_cmd(pip_cmd + ["--upgrade", "pip", "setuptools", "wheel"], proj_dir)
        self._run_cmd(pip_cmd + ["--upgrade", cfg.builder], proj_dir, f"安装引擎 {cfg.builder}...")

        if cfg.use_requirements and (req := Path(cfg.requirements_path) if cfg.requirements_path else proj_dir / "requirements.txt").exists():
            cmd = pip_cmd + ["--upgrade", "--force-reinstall", "-r", str(req)] if cfg.force_reinstall else pip_cmd + ["-r", str(req)]
            self._run_cmd(cmd, proj_dir, "安装项目依赖...")

        freeze_path = logs_dir / "freeze.txt"
        freeze_path.write_text(subprocess.run([str(vpy), "-m", "pip", "freeze"], stdout=subprocess.PIPE, text=True).stdout, encoding="utf-8")
        self._apply_smart_heuristics(freeze_path)

        self.stage.emit("编译打包")
        self._header(f"开始 {cfg.builder.upper()} 打包")
        
        if cfg.builder == "pyinstaller":
            cmd = [str(vpy), "-m", "PyInstaller", str(entry_py), "--noconfirm", "--clean", "--name", cfg.app_name]
            cmd += ["--distpath", str(dist_dir), "--workpath", str(build_dir), "--specpath", str(build_dir)]
            cmd += ["--onefile"] if cfg.onefile else ["--onedir"]
            if cfg.windowed: cmd += ["--windowed"]
            if cfg.icon_path: cmd += ["--icon", str(Path(cfg.icon_path).resolve())]
            if cfg.optimize_level > 0: cmd += [f"--optimize={cfg.optimize_level}"]
            if not cfg.use_upx: cmd += ["--noupx"]
            
            sep = ";" if IS_WIN else ":"
            for item in cfg.add_data: cmd += ["--add-data", item] if sep in item else []
            for h in cfg.hidden_imports: cmd += ["--hidden-import", h]
            for p in cfg.collect_all: cmd += ["--collect-all", p]

            if cfg.splash_enable and not IS_MAC:
                splash = Path(cfg.splash_path) if cfg.splash_path else work_root / "splash.png"
                if not splash.exists(): generate_default_splash_png(splash, cfg.app_name)
                cmd += ["--splash", str(splash)]
        else:
            cmd = [str(vpy), "-m", "nuitka", "--assume-yes-for-downloads", f"--output-dir={dist_dir}"]
            cmd += ["--onefile"] if cfg.onefile else ["--standalone"]
            if cfg.windowed: cmd += ["--macos-create-app-bundle"] if IS_MAC else ["--windows-disable-console"] if IS_WIN else ["--disable-console"]
            if cfg.icon_path: cmd += [f"--macos-app-icon={cfg.icon_path}"] if IS_MAC else [f"--windows-icon-from-ico={cfg.icon_path}"] if IS_WIN else [f"--linux-icon={cfg.icon_path}"]
            
            for h in cfg.hidden_imports: cmd += [f"--include-module={h}"]
            for p in cfg.collect_all: cmd += [f"--include-package={p}"]
            sep = ";" if IS_WIN else ":"
            for item in cfg.add_data:
                if sep in item: src, dest = item.split(sep, 1); cmd += [f"--include-data-dir={src}={dest}"]
            for plg in cfg.nuitka_plugins: cmd += [f"--enable-plugin={plg}"]

        if cfg.extra_args: cmd += [x for x in cfg.extra_args.split() if x]
        if cfg.builder == "nuitka": cmd += [str(entry_py)]

        self._emit(format_cmd(cmd))
        self._run_cmd(cmd, proj_dir)

        self.stage.emit("导出产物")
        rm_tree(Path(final_export)); safe_mkdir(Path(final_export))
        for it in dist_dir.glob("*"): shutil.copytree(it, Path(final_export) / it.name) if it.is_dir() else shutil.copy2(it, Path(final_export) / it.name)
        
        return final_export

# -----------------------------
# UI 样式
# -----------------------------
def qss_fluent_light() -> str:
    return """
    QMainWindow { background: #F0F2F5; }
    QWidget { color: #1E293B; font-size: 13px; font-family: -apple-system, BlinkMacSystemFont, "PingFang SC", "Microsoft YaHei"; }
    QTabWidget::pane { border: 1px solid #E2E8F0; border-radius: 10px; background: #FFFFFF; top: -1px; }
    QTabBar::tab { background: #F8FAFC; border: 1px solid #E2E8F0; border-bottom: none; padding: 10px 20px; margin-right: 4px; border-top-left-radius: 8px; border-top-right-radius: 8px; color: #64748B; font-weight: bold; }
    QTabBar::tab:selected { background: #FFFFFF; color: #2563EB; border-color: #E2E8F0; border-top: 3px solid #2563EB; }
    QTabBar::tab:hover:!selected { background: #F1F5F9; color: #0F172A; }
    QGroupBox { border: 1px solid #CBD5E1; border-radius: 8px; margin-top: 14px; padding: 16px 12px 12px 12px; background: #FFFFFF; }
    QGroupBox::title { subcontrol-origin: margin; left: 16px; padding: 0 6px; color: #0F172A; font-weight: 800; }
    QLineEdit, QPlainTextEdit, QComboBox, QSpinBox { background: #F8FAFC; border: 1px solid #CBD5E1; border-radius: 6px; padding: 8px 10px; color: #0F172A; }
    QLineEdit:focus, QPlainTextEdit:focus, QComboBox:focus, QSpinBox:focus { border: 1.5px solid #3B82F6; background: #FFFFFF; }
    QTextEdit#LogTerminal { background-color: #0F172A; color: #38BDF8; border: 1px solid #334155; border-radius: 8px; padding: 10px; font-family: "Menlo", "Consolas", monospace; font-size: 12px; line-height: 1.4; }
    QPushButton { background: #FFFFFF; color: #334155; border: 1px solid #CBD5E1; border-radius: 6px; padding: 8px 16px; font-weight: 600; }
    QPushButton:hover { background: #F1F5F9; border-color: #94A3B8; color: #0F172A; }
    QPushButton:pressed { background: #E2E8F0; }
    QPushButton:disabled { background: #F8FAFC; color: #94A3B8; border-color: #E2E8F0; }
    QPushButton#primaryButton { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #3B82F6, stop:1 #2563EB); color: #FFFFFF; border: none; }
    QPushButton#primaryButton:hover { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #2563EB, stop:1 #1D4ED8); }
    QPushButton#dangerButton { background: #FEF2F2; color: #DC2626; border: 1px solid #FECACA; }
    QPushButton#dangerButton:hover { background: #FEE2E2; border-color: #F87171; }
    QCheckBox { spacing: 8px; }
    QCheckBox::indicator { width: 18px; height: 18px; border-radius: 4px; border: 1px solid #CBD5E1; background: #FFFFFF; }
    QCheckBox::indicator:checked { background: #2563EB; border: 1px solid #2563EB; }
    """

def wrap_scroll(widget: QWidget) -> QScrollArea:
    area = QScrollArea(); area.setWidgetResizable(True); area.setFrameShape(QScrollArea.NoFrame); area.setWidget(widget)
    return area

# -----------------------------
# 主窗口
# -----------------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} v{APP_VERSION}")
        self.setMinimumSize(QSize(1250, 800))
        self._thread: Optional[QThread] = None
        self._worker: Optional[BuildWorker] = None
        self._is_syncing = False

        self._init_ui()
        self.setStyleSheet(qss_fluent_light())
        self._init_defaults()

    def _init_ui(self):
        root = QWidget(); self.setCentralWidget(root)
        splitter = QSplitter(Qt.Horizontal)

        left_panel = QWidget(); left_layout = QVBoxLayout(left_panel); left_layout.setContentsMargins(16, 16, 16, 16); left_layout.setSpacing(14)
        header = QWidget(); hl = QVBoxLayout(header); hl.setContentsMargins(0,0,0,0); hl.setSpacing(4)
        title = QLabel(f"📦 {APP_NAME}"); title.setStyleSheet("font-size: 22px; font-weight: 800; color: #0F172A;")
        sub = QLabel("支持拖拽文件 • 每次开启全新纯净状态 • 硬核杀进程 • 终极免疫"); sub.setStyleSheet("color: #64748B; font-size: 13px;")
        hl.addWidget(title); hl.addWidget(sub); left_layout.addWidget(header)

        tabs = QTabWidget(); tabs.setDocumentMode(True)

        def make_row(edit: QWidget, btn_txt: str, fn) -> QWidget:
            w = QWidget(); l = QHBoxLayout(w); l.setContentsMargins(0,0,0,0); l.setSpacing(8)
            btn = QPushButton(btn_txt); btn.clicked.connect(fn)
            l.addWidget(edit, 1); l.addWidget(btn); return w

        # --- Tab 1: 基础 ---
        tab_basic = QWidget(); l_b = QVBoxLayout(tab_basic); l_b.setContentsMargins(16,16,16,16)
        gb_proj = QGroupBox("项目与输出 (支持拖拽)"); fb_proj = QFormLayout(gb_proj)
        self.ed_proj, self.ed_entry, self.ed_out, self.ed_app, self.ed_hostpy = DropLineEdit(), DropLineEdit(), DropLineEdit(), QLineEdit("MyApp"), DropLineEdit()
        fb_proj.addRow("📂 项目目录：", make_row(self.ed_proj, "选择目录...", self._choose_project))
        fb_proj.addRow("📄 入口脚本：", make_row(self.ed_entry, "选择入口...", self._choose_entry))
        fb_proj.addRow("🎯 输出目录：", make_row(self.ed_out, "选择输出...", self._choose_output))
        fb_proj.addRow("📝 应用名称：", self.ed_app)
        
        w_py = QWidget(); l_py = QHBoxLayout(w_py); l_py.setContentsMargins(0,0,0,0); l_py.setSpacing(8)
        btn_auto_py = QPushButton("🔄 自动"); btn_auto_py.setToolTip("自动探测系统中最佳的 Python 3 路径"); btn_auto_py.clicked.connect(self._auto_detect_python)
        btn_pick_py = QPushButton("📂 浏览..."); btn_pick_py.clicked.connect(self._choose_host_python)
        l_py.addWidget(self.ed_hostpy, 1); l_py.addWidget(btn_auto_py); l_py.addWidget(btn_pick_py)
        fb_proj.addRow("🐍 宿主 Python：", w_py)
        
        gb_build = QGroupBox("构建引擎与输出模式"); fb_build = QFormLayout(gb_build)
        self.cb_builder = QComboBox(); self.cb_builder.addItems(["pyinstaller", "nuitka"]); self.cb_builder.currentTextChanged.connect(self._sync_ui_state)
        
        self.ck_onefile = QCheckBox("打包为单文件 (Unix 可执行文件)")
        self.ck_windowed = QCheckBox("GUI 模式 (.app 程序包)")
        if IS_MAC:
            self.ck_onefile.setText("Unix 单文件 (双击会弹终端黑框，不支持内置图标)")
            self.ck_windowed.setText("macOS .app 程序包 (无终端，支持图标，强烈推荐)")
            self.ck_onefile.clicked.connect(lambda: self._mac_mutex(True))
            self.ck_windowed.clicked.connect(lambda: self._mac_mutex(False))
        
        self.ed_icon = DropLineEdit()
        self.ed_icon.setPlaceholderText("Mac 系统建议使用 .icns 格式图标")
        
        fb_build.addRow("⚙️ 核心引擎：", self.cb_builder)
        fb_build.addRow("🖼️ 应用图标：", make_row(self.ed_icon, "选择图标...", self._choose_icon))
        fb_build.addRow("", self.ck_onefile); fb_build.addRow("", self.ck_windowed)
        l_b.addWidget(gb_proj); l_b.addWidget(gb_build); l_b.addStretch(1)
        tabs.addTab(wrap_scroll(tab_basic), "📌 基础配置")

        # --- Tab 2: 依赖 ---
        tab_dep = QWidget(); l_d = QVBoxLayout(tab_dep); l_d.setContentsMargins(16,16,16,16)
        gb_dep = QGroupBox("环境与依赖管理"); fd_dep = QFormLayout(gb_dep)
        self.ck_use_req = QCheckBox("根据 requirements.txt 自动安装依赖"); self.ck_use_req.setChecked(True)
        self.ed_req = DropLineEdit()
        self.cb_mirror = QComboBox(); self.cb_mirror.addItems(["(默认官方源)", "https://pypi.tuna.tsinghua.edu.cn/simple", "https://mirrors.aliyun.com/pypi/simple/"])
        self.ck_upg_pip = QCheckBox("构建前升级 pip/setuptools"); self.ck_upg_pip.setChecked(True)
        self.ck_force = QCheckBox("强制重新安装所有依赖 (--force-reinstall)")
        self.ck_nocache = QCheckBox("禁用 pip 缓存 (--no-cache-dir)")
        fd_dep.addRow("", self.ck_use_req); fd_dep.addRow("📄 依赖文件：", make_row(self.ed_req, "选择文件...", self._choose_req))
        fd_dep.addRow("🚀 pip 镜像源：", self.cb_mirror); fd_dep.addRow("", self.ck_upg_pip); fd_dep.addRow("", self.ck_force); fd_dep.addRow("", self.ck_nocache)
        l_d.addWidget(gb_dep); l_d.addStretch(1)
        tabs.addTab(wrap_scroll(tab_dep), "📦 依赖管理")

        # --- Tab 3: 高级 ---
        tab_adv = QWidget(); l_a = QVBoxLayout(tab_adv); l_a.setContentsMargins(16,16,16,16)
        gb_adv = QGroupBox("高级参数 (每行一个)"); fa_adv = QFormLayout(gb_adv)
        self.pt_hidden, self.pt_collect, self.pt_data, self.pt_extra = QPlainTextEdit(), QPlainTextEdit(), QPlainTextEdit(), QPlainTextEdit()
        for pt in (self.pt_hidden, self.pt_collect, self.pt_data, self.pt_extra): pt.setMaximumHeight(70)
        self.ck_upx = QCheckBox("使用 UPX 极致压缩产物体积 (仅限 PyInstaller, 需预装 upx)")
        fa_adv.addRow("🛡️ 隐式导入：", self.pt_hidden); fa_adv.addRow("🧲 强制收集包：", self.pt_collect)
        fa_adv.addRow("📁 数据文件(src:dst)：", self.pt_data); fa_adv.addRow("🔧 其它参数：", self.pt_extra)
        fa_adv.addRow("", self.ck_upx)
        info_lb = QLabel("💡 工具已内置终极 AI与多媒体库 免疫字典，无需手动填写常见高危依赖。")
        info_lb.setStyleSheet("color: #059669; font-weight: bold;")
        l_a.addWidget(gb_adv); l_a.addWidget(info_lb); l_a.addStretch(1)
        tabs.addTab(wrap_scroll(tab_adv), "🛠️ 高级与优化")

        # --- Tab 4: 体验 ---
        tab_ux = QWidget(); l_u = QVBoxLayout(tab_ux); l_u.setContentsMargins(16,16,16,16)
        gb_splash = QGroupBox("启动体验优化"); fu_splash = QFormLayout(gb_splash)
        self.ck_splash = QCheckBox("启用加载启动画面 (仅 PyInstaller + 非 Mac 平台生效)")
        self.ck_splash.stateChanged.connect(self._sync_ui_state)
        self.ed_splash = DropLineEdit(); self.ed_splash.setPlaceholderText("支持拖拽图片，留空则自动生成")
        self.spin_opt = QSpinBox(); self.spin_opt.setRange(0, 2); self.spin_opt.setValue(1)
        fu_splash.addRow("", self.ck_splash); fu_splash.addRow("🖼️ 自定义图片：", make_row(self.ed_splash, "浏览...", self._choose_splash))
        fu_splash.addRow("⚡ 字节码优化级：", self.spin_opt)
        l_u.addWidget(gb_splash); l_u.addStretch(1)
        tabs.addTab(wrap_scroll(tab_ux), "✨ 体验增强")

        left_layout.addWidget(tabs, 1)

        # Footer Buttons
        footer = QWidget(); fl = QHBoxLayout(footer); fl.setContentsMargins(0,0,0,0)
        self.btn_stop = QPushButton("🛑 紧急强制终止"); self.btn_stop.setObjectName("dangerButton"); self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self._stop_build)
        self.btn_build = QPushButton("🚀 启动编译打包"); self.btn_build.setObjectName("primaryButton"); self.btn_build.setMinimumHeight(38)
        self.btn_build.clicked.connect(self._start_build)
        fl.addStretch(1); fl.addWidget(self.btn_stop); fl.addWidget(self.btn_build); left_layout.addWidget(footer)

        # --- Right Panel: Console ---
        right_panel = QWidget(); rl = QVBoxLayout(right_panel); rl.setContentsMargins(0, 16, 16, 16)
        gb_log = QGroupBox("🖥️ 编译日志终端"); ll = QVBoxLayout(gb_log); ll.setSpacing(8)
        topbar = QWidget(); tbl = QHBoxLayout(topbar); tbl.setContentsMargins(0,0,0,0)
        btn_clear = QPushButton("🧹 清屏"); btn_clear.clicked.connect(lambda: self.log_view.clear())
        tbl.addStretch(1); tbl.addWidget(btn_clear)
        self.log_view = QTextEdit(); self.log_view.setReadOnly(True); self.log_view.setObjectName("LogTerminal")
        ll.addWidget(topbar); ll.addWidget(self.log_view, 1); rl.addWidget(gb_log, 1)

        splitter.addWidget(left_panel); splitter.addWidget(right_panel); splitter.setSizes([650, 600])
        layout = QHBoxLayout(root); layout.setContentsMargins(0,0,0,0); layout.addWidget(splitter)
        self.statusBar().showMessage("准备就绪")

    def _init_defaults(self):
        """🌟 初始化纯净状态"""
        self.ed_out.setText(str(home_desktop_dir() / DEFAULT_OUTPUT_DIRNAME))
        self._auto_detect_python(show_msg=False)
        
        if IS_MAC:
            self._is_syncing = True
            self.ck_windowed.setChecked(True)
            self.ck_onefile.setChecked(False)
            self._is_syncing = False
        else:
            self.ck_windowed.setChecked(True)
            self.ck_onefile.setChecked(True)
            
        self._sync_ui_state()

    def _mac_mutex(self, is_onefile: bool):
        if not IS_MAC or self._is_syncing: return
        self._is_syncing = True
        try:
            if is_onefile and self.ck_onefile.isChecked(): self.ck_windowed.setChecked(False)
            elif not is_onefile and self.ck_windowed.isChecked(): self.ck_onefile.setChecked(False)
        finally: self._is_syncing = False

    def _sync_ui_state(self):
        builder = self.cb_builder.currentText()
        is_splash_supported = (builder == "pyinstaller" and not IS_MAC)
        self.ck_splash.setEnabled(is_splash_supported)
        if not is_splash_supported: self.ck_splash.setChecked(False)
        self.ed_splash.setEnabled(is_splash_supported and self.ck_splash.isChecked())

    def _auto_detect_python(self, show_msg=True):
        try:
            best_py = str(find_host_python())
            self.ed_hostpy.setText(best_py)
            if show_msg: self.statusBar().showMessage(f"已自动选择最佳 Python: {best_py}", 4000)
        except Exception as e:
            if show_msg: QMessageBox.warning(self, "自动检测失败", str(e))

    def _choose_project(self):
        if d := QFileDialog.getExistingDirectory(self, "选择项目目录", str(Path.home())):
            proj = Path(normpath(d)); self.ed_proj.setText(str(proj)); self.ed_app.setText(guess_app_name(proj))
            for f, ed in [("requirements.txt", self.ed_req), ("main.py", self.ed_entry)]:
                if (proj / f).exists(): ed.setText(str(proj / f))
            out, _ = ensure_writable_directory(proj / DEFAULT_OUTPUT_DIRNAME, home_desktop_dir() / DEFAULT_OUTPUT_DIRNAME)
            self.ed_out.setText(str(out))

    def _choose_entry(self):
        if f := QFileDialog.getOpenFileName(self, "选择入口", self.ed_proj.text() or str(Path.home()), "Python (*.py)")[0]: self.ed_entry.setText(normpath(f))
    def _choose_output(self):
        if d := QFileDialog.getExistingDirectory(self, "选择输出", self.ed_out.text() or str(Path.home())): self.ed_out.setText(normpath(d))
    def _choose_icon(self):
        if f := QFileDialog.getOpenFileName(self, "选择图标", str(Path.home()), "Icon (*.ico *.icns *.png *.jpg)")[0]: self.ed_icon.setText(normpath(f))
    def _choose_req(self):
        if f := QFileDialog.getOpenFileName(self, "选择依赖文件", self.ed_proj.text() or str(Path.home()), "Text (*.txt)")[0]: self.ed_req.setText(normpath(f))
    def _choose_host_python(self):
        if f := QFileDialog.getOpenFileName(self, "选择宿主 Python", str(Path.home()), "All (*)")[0]: self.ed_hostpy.setText(normpath(f))
    def _choose_splash(self):
        if f := QFileDialog.getOpenFileName(self, "选择 Splash", str(Path.home()), "Image (*.png *.jpg)")[0]: self.ed_splash.setText(normpath(f))

    def _collect_cfg(self) -> BuildConfig:
        cfg = default_build_config()
        cfg.project_dir, cfg.entry_script, cfg.output_dir = self.ed_proj.text().strip(), self.ed_entry.text().strip(), self.ed_out.text().strip()
        cfg.app_name, cfg.builder = sanitize_name(self.ed_app.text().strip() or "MyApp"), self.cb_builder.currentText().strip()
        cfg.onefile, cfg.windowed, cfg.icon_path, cfg.host_python = self.ck_onefile.isChecked(), self.ck_windowed.isChecked(), self.ed_icon.text().strip(), self.ed_hostpy.text().strip()
        cfg.splash_enable, cfg.splash_path, cfg.optimize_level = self.ck_splash.isChecked(), self.ed_splash.text().strip(), int(self.spin_opt.value())
        cfg.use_requirements, cfg.requirements_path = self.ck_use_req.isChecked(), self.ed_req.text().strip()
        cfg.upgrade_pip, cfg.force_reinstall, cfg.no_pip_cache = self.ck_upg_pip.isChecked(), self.ck_force.isChecked(), self.ck_nocache.isChecked()
        cfg.pip_mirror = "" if "默认" in self.cb_mirror.currentText() else self.cb_mirror.currentText()
        cfg.hidden_imports, cfg.collect_all, cfg.add_data = split_lines(self.pt_hidden.toPlainText()), split_lines(self.pt_collect.toPlainText()), split_lines(self.pt_data.toPlainText())
        cfg.extra_args, cfg.use_upx = self.pt_extra.toPlainText().replace("\n", " "), self.ck_upx.isChecked()
        return cfg

    def _start_build(self):
        cfg = self._collect_cfg()
        if not cfg.project_dir or not cfg.entry_script: return QMessageBox.warning(self, "提示", "请确保已填写【项目目录】与【入口脚本】。")
        self.log_view.clear(); self._append_log(f"> 启动引擎：{cfg.builder.upper()} ..."); self.statusBar().showMessage("构建中...")
        self.btn_build.setEnabled(False); self.btn_stop.setEnabled(True)

        self._thread = QThread()
        self._worker = BuildWorker(cfg)
        self._worker.moveToThread(self._thread)

        self._thread.started.connect(self._worker.run)
        self._worker.log.connect(self._append_log)
        self._worker.stage.connect(lambda s: self.statusBar().showMessage(f"当前阶段：{s}"))
        self._worker.done.connect(self._on_done)
        self._worker.done.connect(self._thread.quit)
        self._thread.finished.connect(self._thread.deleteLater)
        self._thread.start()

    def _stop_build(self):
        if self._worker:
            self._worker.cancel()
            self._append_log("\n[WARN] 已发送终止信号，正在强行中断底层的编译进程...")
            self.btn_stop.setEnabled(False)

    def _on_done(self, ok: bool, msg: str, out_dir: str):
        self.btn_build.setEnabled(True); self.btn_stop.setEnabled(False)
        self.statusBar().showMessage("任务完成 ✅" if ok else "任务已终止/失败 🛑")
        if ok and out_dir and Path(out_dir).exists():
            if QMessageBox.information(self, "成功", f"{msg}\n是否立即打开输出目录？", QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
                QDesktopServices.openUrl(QUrl.fromLocalFile(out_dir))
        else: QMessageBox.information(self, "构建结果", msg)
        self._worker, self._thread = None, None

    def _append_log(self, s: str):
        self.log_view.append(s); self.log_view.ensureCursorVisible()

def main():
    try: QGuiApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    except Exception: pass
    app = QApplication(sys.argv)
    app.setFont(QFont("PingFang SC" if IS_MAC else "Microsoft YaHei UI", 10))
    w = MainWindow()
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
