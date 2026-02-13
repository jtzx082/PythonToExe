import sys
import os
import json
import requests
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QTextEdit, QPushButton, QComboBox,
    QFileDialog, QMessageBox, QDialog, QFormLayout
)
from PyQt6.QtCore import Qt, QLocale
from PyQt6.QtGui import QFont, QInputMethod
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# ===================== 配置文件路径 =====================
CONFIG_PATH = "config.json"
# ======================================================

class ConfigManager:
    """配置文件管理：保存/加载API Key"""
    @staticmethod
    def load_config():
        if os.path.exists(CONFIG_PATH):
            try:
                with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                    return json.load(f)
            except:
                return {"deepseek_api_key": ""}
        return {"deepseek_api_key": ""}

    @staticmethod
    def save_api_key(api_key):
        config = {"deepseek_api_key": api_key.strip()}
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)

class APISettingDialog(QDialog):
    """API Key 设置弹窗（修复中文输入）"""
    def __init__(self, current_key):
        super().__init__()
        self.setWindowTitle("API 设置")
        self.setFixedSize(500, 180)
        self.api_key = current_key
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        form_layout = QFormLayout()

        # API Key 输入框（修复中文输入）
        self.key_input = QLineEdit()
        self.key_input.setPlaceholderText("请输入 DeepSeek API Key（支持中文粘贴）")
        self.key_input.setText(self.api_key)
        self.key_input.setEchoMode(QLineEdit.EchoMode.Password)
        # 强制启用中文输入
        self.key_input.setAttribute(Qt.WidgetAttribute.WA_InputMethodEnabled, True)
        self.key_input.setLocale(QLocale(QLocale.Language.Chinese, QLocale.Country.China))
        form_layout.addRow("DeepSeek API Key：", self.key_input)

        # 验证按钮 + 保存按钮
        btn_layout = QHBoxLayout()
        self.check_btn = QPushButton("🔍 验证API有效性")
        self.check_btn.clicked.connect(self.check_api_valid)
        self.save_btn = QPushButton("✅ 保存并应用")
        self.save_btn.clicked.connect(self.save_key)
        btn_layout.addWidget(self.check_btn)
        btn_layout.addWidget(self.save_btn)
        form_layout.addRow("", btn_layout)

        layout.addLayout(form_layout)
        self.setLayout(layout)

    def check_api_valid(self):
        """验证API Key是否有效"""
        key = self.key_input.text().strip()
        if not key:
            QMessageBox.warning(self, "提示", "API Key 不能为空")
            return
        
        headers = {
            "Authorization": f"Bearer {key}",
            "Content-Type": "application/json"
        }
        data = {
            "model": "deepseek-chat",
            "messages": [{"role": "user", "content": "测试"}],
            "temperature": 0.1
        }
        try:
            resp = requests.post(
                "https://api.deepseek.com/v1/chat/completions",
                json=data,
                headers=headers,
                timeout=30
            )
            if resp.status_code == 200:
                QMessageBox.information(self, "成功", "API Key 有效！")
            elif resp.status_code == 401:
                QMessageBox.critical(self, "错误", "API Key 无效或已过期！")
            else:
                QMessageBox.critical(self, "错误", f"验证失败：{resp.status_code}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"网络异常：{str(e)}")

    def save_key(self):
        key = self.key_input.text().strip()
        if not key:
            QMessageBox.warning(self, "提示", "API Key 不能为空")
            return
        ConfigManager.save_api_key(key)
        QMessageBox.information(self, "成功", "API Key 已保存，下次启动自动加载！")
        self.accept()

class PaperWriter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = ConfigManager.load_config()
        self.DEEPSEEK_API_KEY = self.config.get("deepseek_api_key", "")
        self.DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions"
        self.setWindowTitle("智能公文/论文撰写工具 | API可配置 | 标准Word导出")
        self.setMinimumSize(950, 780)
        # 全局启用中文输入
        self.setAttribute(Qt.WidgetAttribute.WA_InputMethodEnabled, True)
        self.setLocale(QLocale(QLocale.Language.Chinese, QLocale.Country.China))
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # ========== 顶部：API 设置按钮 ==========
        top_layout = QHBoxLayout()
        self.api_status_label = QLabel()
        self.update_api_status()
        self.setting_btn = QPushButton("⚙️ API 设置")
        self.setting_btn.clicked.connect(self.open_api_setting)
        top_layout.addWidget(self.api_status_label)
        top_layout.addStretch()
        top_layout.addWidget(self.setting_btn)
        layout.addLayout(top_layout)

        # ========== 文稿类型 ==========
        type_layout = QHBoxLayout()
        type_label = QLabel("文稿类型：")
        self.type_combo = QComboBox()
        # 修复ComboBox中文显示
        self.type_combo.setAttribute(Qt.WidgetAttribute.WA_InputMethodEnabled, True)
        self.type_combo.addItems([
            "期刊论文", "工作计划", "工作总结", "学习反思", "教学案例", "汇报材料", "自定义"
        ])
        type_layout.addWidget(type_label)
        type_layout.addWidget(self.type_combo)
        layout.addLayout(type_layout)

        # ========== 题目输入（修复中文输入） ==========
        title_layout = QHBoxLayout()
        title_label = QLabel("题目/要求：")
        self.title_input = QLineEdit()
        self.title_input.setPlaceholderText("输入完整题目或详细要求，例如：2026年度部门工作总结")
        # 强制启用中文输入
        self.title_input.setAttribute(Qt.WidgetAttribute.WA_InputMethodEnabled, True)
        self.title_input.setLocale(QLocale(QLocale.Language.Chinese, QLocale.Country.China))
        title_layout.addWidget(title_label)
        title_layout.addWidget(self.title_input)
        layout.addLayout(title_layout)

        # ========== 生成大纲 ==========
        self.outline_btn = QPushButton("📌 生成标准公文大纲")
        self.outline_btn.clicked.connect(self.generate_outline)
        layout.addWidget(self.outline_btn)

        # ========== 大纲编辑区（修复中文输入） ==========
        layout.addWidget(QLabel("📝 大纲（纯文本公文层级，可直接修改）："))
        self.outline_edit = QTextEdit()
        self.outline_edit.setPlaceholderText("大纲格式：一、 →（一）→1. →（1），禁止使用Markdown")
        self.outline_edit.setAttribute(Qt.WidgetAttribute.WA_InputMethodEnabled, True)
        self.outline_edit.setLocale(QLocale(QLocale.Language.Chinese, QLocale.Country.China))
        layout.addWidget(self.outline_edit)

        # ========== 撰写全文 ==========
        self.write_btn = QPushButton("🚀 按公文格式撰写完整文稿")
        self.write_btn.clicked.connect(self.generate_full_text)
        layout.addWidget(self.write_btn)

        # ========== 文稿展示 ==========
        layout.addWidget(QLabel("📄 完整文稿（纯文本无格式）："))
        self.result_text = QTextEdit()
        self.result_text.setAttribute(Qt.WidgetAttribute.WA_InputMethodEnabled, True)
        self.result_text.setLocale(QLocale(QLocale.Language.Chinese, QLocale.Country.China))
        layout.addWidget(self.result_text)

        # ========== 导出Word ==========
        self.export_btn = QPushButton("📄 导出【国家标准公文格式】Word文档")
        self.export_btn.clicked.connect(self.export_word)
        layout.addWidget(self.export_btn)

    def update_api_status(self):
        """更新API状态显示"""
        if self.DEEPSEEK_API_KEY:
            self.api_status_label.setText("✅ API Key 已配置")
            self.api_status_label.setStyleSheet("color:green;")
        else:
            self.api_status_label.setText("❌ 未设置 API Key，请先配置")
            self.api_status_label.setStyleSheet("color:red;")

    def open_api_setting(self):
        """打开API设置弹窗"""
        dialog = APISettingDialog(self.DEEPSEEK_API_KEY)
        if dialog.exec():
            self.config = ConfigManager.load_config()
            self.DEEPSEEK_API_KEY = self.config.get("deepseek_api_key", "")
            self.update_api_status()

    def check_api_key(self):
        """检查API是否配置"""
        if not self.DEEPSEEK_API_KEY:
            QMessageBox.critical(self, "错误", "请先点击右上角【API 设置】配置 DeepSeek Key！")
            return False
        return True

    def call_deepseek(self, prompt):
        """调用DeepSeek API（带详细错误处理）"""
        if not self.check_api_key():
            return "API未配置，请先设置"
        
        headers = {
            "Authorization": f"Bearer {self.DEEPSEEK_API_KEY}",
            "Content-Type": "application/json"
        }
        data = {
            "model": "deepseek-chat",
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.2
        }
        try:
            resp = requests.post(self.DEEPSEEK_API_URL, json=data, timeout=90)
            
            # 详细错误处理
            if resp.status_code == 401:
                return f"API调用失败：401未授权\n原因：API Key无效/过期/格式错误\n请重新配置API Key"
            elif resp.status_code == 403:
                return f"API调用失败：403禁止访问\n原因：账号余额不足/权限限制"
            elif resp.status_code == 429:
                return f"API调用失败：429请求频繁\n原因：超出API调用频率限制，请稍后再试"
            elif resp.status_code != 200:
                return f"API调用失败：{resp.status_code}\n响应内容：{resp.text}"
            
            return resp.json()["choices"][0]["message"]["content"].strip()
        
        except requests.exceptions.ConnectionError:
            return "API调用失败：网络连接异常，请检查网络"
        except requests.exceptions.Timeout:
            return "API调用失败：请求超时，请重试"
        except Exception as e:
            return f"API调用失败：{str(e)}"

    def generate_outline(self):
        if not self.check_api_key(): return
        doc_type = self.type_combo.currentText()
        title = self.title_input.text().strip()
        if not title:
            QMessageBox.warning(self, "提示", "请输入题目或要求")
            return
        prompt = f"""
        你是专业公文写作助手，请为【{doc_type}】生成大纲。
        题目：{title}
        要求：
        1. 纯文本，绝对禁止任何Markdown、符号、表格、代码
        2. 严格使用国家标准公文层级：一、 →（一）→1. →（1）
        3. 结构清晰，可直接用于正式文稿
        只输出大纲，不要多余解释。
        """
        outline = self.call_deepseek(prompt)
        self.outline_edit.setPlainText(outline)

    def generate_full_text(self):
        if not self.check_api_key(): return
        doc_type = self.type_combo.currentText()
        title = self.title_input.text().strip()
        outline = self.outline_edit.toPlainText().strip()
        if not title or not outline:
            QMessageBox.warning(self, "提示", "请先生成并完善大纲")
            return
        prompt = f"""
        你是专业公文撰稿人，请按【{doc_type}】正式文体写作。
        题目：{title}
        大纲：{outline}
        要求：
        1. 纯文本，无任何Markdown、格式符、特殊符号
        2. 严格使用公文层级：一、 （一） 1. （1）
        3. 语言正式、逻辑严谨、内容完整
        4. 直接输出正文，不要前言、说明、解释
        """
        full_text = self.call_deepseek(prompt)
        self.result_text.setPlainText(full_text)

    def export_word(self):
        """导出国家标准公文格式Word（GB/T 9704-2012）"""
        title = self.title_input.text().strip()
        content = self.result_text.toPlainText().strip()
        if not title or not content:
            QMessageBox.warning(self, "提示", "请先生成完整文稿")
            return
        save_path, _ = QFileDialog.getSaveFileName(
            self, "导出Word", f"{title}.docx", "Word文档 (*.docx)"
        )
        if not save_path:
            return
        try:
            doc = Document()
            # A4公文页面设置
            section = doc.sections[0]
            section.page_height = Cm(29.7)
            section.page_width = Cm(21.0)
            section.left_margin = Cm(2.8)
            section.right_margin = Cm(2.6)
            section.top_margin = Cm(3.7)
            section.bottom_margin = Cm(3.5)

            # 公文标题：二号小标宋体、居中
            title_p = doc.add_paragraph()
            title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            title_run = title_p.add_run(title)
            title_run.font.size = Pt(22)
            title_run.font.bold = True
            title_run.font.name = "小标宋体"
            title_run._element.rPr.rFonts.set(qn('w:eastAsia'), '小标宋体')
            doc.add_paragraph()

            # 正文按公文层级自动排版
            lines = content.splitlines()
            for line in lines:
                line = line.strip()
                if not line: continue
                p = doc.add_paragraph()
                run = p.add_run(line)
                run.font.size = Pt(16)  # 三号字

                # 一级标题：一、 黑体
                if line.startswith(("一、","二、","三、","四、","五、")):
                    run.font.name = "黑体"
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
                    run.font.bold = True
                    p.paragraph_format.first_line_indent = Cm(0)
                # 二级标题：（一） 楷体
                elif line.startswith(("（一）","（二）","（三）")):
                    run.font.name = "楷体_GB2312"
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '楷体_GB2312')
                    p.paragraph_format.first_line_indent = Cm(0)
                # 三级标题：1.  加粗
                elif line.startswith(("1.","2.","3.")):
                    run.font.bold = True
                    p.paragraph_format.first_line_indent = Cm(0)
                # 正文：仿宋_GB2312 + 首行缩进
                else:
                    run.font.name = "仿宋_GB2312"
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋_GB2312')
                    p.paragraph_format.first_line_indent = Cm(0.74)
                p.paragraph_format.line_spacing = 1.25

            doc.save(save_path)
            QMessageBox.information(self, "成功", "已按【国家标准公文格式】导出Word！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导出失败：{str(e)}")

if __name__ == "__main__":
    # 全局启用中文输入
    app = QApplication(sys.argv)
    app.setLocale(QLocale(QLocale.Language.Chinese, QLocale.Country.China))
    window = PaperWriter()
    window.show()
    sys.exit(app.exec())
