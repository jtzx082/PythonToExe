import sys
import os
import requests
import docx
import PyPDF2
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QTextEdit, QLabel, QLineEdit, QFileDialog, QProgressBar, QMessageBox)
from PyQt6.QtCore import QThread, pyqtSignal, Qt
from PyQt6.QtGui import QFont
from pptx import Presentation
from pptx.util import Inches, Pt
from openai import OpenAI

# ================= 线程类：调用DeepSeek生成大纲 =================
class OutlineWorker(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, api_key, content):
        super().__init__()
        self.api_key = api_key
        self.content = content

    def run(self):
        try:
            client = OpenAI(api_key=self.api_key, base_url="https://api.deepseek.com")
            prompt = f"""
            请根据以下内容/主题生成一份PPT演示文稿大纲。
            要求格式严格遵守以下Markdown规范，以便后续程序解析：
            每个幻灯片以 '# ' 开头作为标题。
            幻灯片的内容要点以 '- ' 开头。
            在每个幻灯片的最后，提供一个用于生成配图的英文关键词，格式为 '[Image Keyword: 关键词]'。
            
            内容/主题：{self.content}
            
            示例：
            # PPT封面：人工智能的未来
            - 探索AI的无限可能
            - 演讲者：张三
            [Image Keyword: Artificial Intelligence Future]
            """
            
            response = client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": "你是一个专业的PPT大纲设计师。"},
                    {"role": "user", "content": prompt}
                ]
            )
            outline = response.choices[0].message.content
            self.finished.emit(outline)
        except Exception as e:
            self.error.emit(str(e))

# ================= 线程类：生成PPT =================
class PPTWorker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, outline_text, template_path, output_path):
        super().__init__()
        self.outline_text = outline_text
        self.template_path = template_path
        self.output_path = output_path

    def parse_outline(self, text):
        slides = []
        current_slide = None
        for line in text.split('\n'):
            line = line.strip()
            if line.startswith('# '):
                if current_slide:
                    slides.append(current_slide)
                current_slide = {'title': line[2:], 'bullets': [], 'keyword': 'presentation'}
            elif line.startswith('- '):
                if current_slide:
                    current_slide['bullets'].append(line[2:])
            elif line.startswith('[Image Keyword:'):
                if current_slide:
                    current_slide['keyword'] = line.split(':')[1].strip()[:-1]
        if current_slide:
            slides.append(current_slide)
        return slides

    def run(self):
        try:
            slides_data = self.parse_outline(self.outline_text)
            if not slides_data:
                raise ValueError("大纲格式错误，未找到幻灯片内容。请确保包含'#'标题。")

            # 加载模板或创建空白
            if self.template_path and os.path.exists(self.template_path):
                prs = Presentation(self.template_path)
            else:
                prs = Presentation()

            total = len(slides_data)
            for idx, slide_data in enumerate(slides_data):
                # 尝试使用"标题和内容"排版 (索引一般为1)
                layout = prs.slide_layouts[1] if len(prs.slide_layouts) > 1 else prs.slide_layouts[0]
                slide = prs.slides.add_slide(layout)
                
                # 填入标题
                if slide.shapes.title:
                    slide.shapes.title.text = slide_data['title']
                
                # 填入要点内容
                if len(slide.placeholders) > 1:
                    tf = slide.placeholders[1].text_frame
                    tf.text = ""
                    for bullet in slide_data['bullets']:
                        p = tf.add_paragraph()
                        p.text = bullet
                        p.level = 0
                
                # AI配图 (使用免费免key的pollinations API生图)
                try:
                    img_url = f"https://image.pollinations.ai/prompt/{slide_data['keyword']}?width=400&height=300&nologo=true"
                    img_data = requests.get(img_url, timeout=10).content
                    img_path = f"temp_img_{idx}.jpg"
                    with open(img_path, 'wb') as handler:
                        handler.write(img_data)
                    
                    # 将图片插入到幻灯片右侧
                    left = Inches(5)
                    top = Inches(2)
                    slide.shapes.add_picture(img_path, left, top, width=Inches(4.5))
                    os.remove(img_path) # 清理临时图片
                except Exception as img_e:
                    print(f"无法生成图片: {img_e}")

                self.progress.emit(int(((idx + 1) / total) * 100))

            prs.save(self.output_path)
            self.finished.emit(self.output_path)
        except Exception as e:
            self.error.emit(str(e))

# ================= 主窗口 GUI =================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AI PPT Master - DeepSeek 智能生成器")
        self.resize(900, 700)
        self.setStyleSheet("""
            QMainWindow { background-color: #f4f5f7; }
            QLabel { font-size: 14px; font-weight: bold; color: #333; }
            QTextEdit, QLineEdit { background-color: white; border: 1px solid #ccc; border-radius: 5px; padding: 8px; font-size: 14px; }
            QPushButton { background-color: #0052cc; color: white; border-radius: 5px; padding: 10px; font-size: 14px; font-weight: bold; }
            QPushButton:hover { background-color: #0043a6; }
            QPushButton:disabled { background-color: #a5b4fc; }
            QProgressBar { text-align: center; border: 1px solid #ccc; border-radius: 5px; }
            QProgressBar::chunk { background-color: #0052cc; }
        """)
        
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)

        # 1. API配置区
        api_layout = QHBoxLayout()
        api_layout.addWidget(QLabel("DeepSeek API Key:"))
        self.api_input = QLineEdit()
        self.api_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.api_input.setPlaceholderText("sk-...")
        api_layout.addWidget(self.api_input)
        layout.addLayout(api_layout)

        # 2. 输入区
        input_label = QLabel("输入主题或上传文件 (TXT/DOCX/PDF):")
        layout.addWidget(input_label)
        
        self.input_text = QTextEdit()
        self.input_text.setPlaceholderText("在此输入PPT主题，或点击右侧按钮解析文档内容...")
        
        btn_layout = QVBoxLayout()
        self.btn_upload = QPushButton("📂 上传解析文件")
        self.btn_upload.clicked.connect(self.upload_file)
        self.btn_gen_outline = QPushButton("✨ 第一步: AI 生成大纲")
        self.btn_gen_outline.clicked.connect(self.generate_outline)
        
        btn_layout.addWidget(self.btn_upload)
        btn_layout.addWidget(self.btn_gen_outline)
        btn_layout.addStretch()

        input_box = QHBoxLayout()
        input_box.addWidget(self.input_text, 4)
        input_box.addLayout(btn_layout, 1)
        layout.addLayout(input_box)

        # 3. 大纲编辑区
        layout.addWidget(QLabel("PPT 大纲 (支持手动调整修改):"))
        self.outline_text = QTextEdit()
        self.outline_text.setPlaceholderText("生成的Markdown大纲将显示在这里，您可以随意修改标题、要点和图片关键词...")
        layout.addWidget(self.outline_text)

        # 4. 生成区
        bottom_layout = QHBoxLayout()
        self.btn_template = QPushButton("🎨 选择本地PPT模板 (可选)")
        self.btn_template.clicked.connect(self.select_template)
        self.template_path = ""
        
        self.btn_generate_ppt = QPushButton("🚀 第二步: 一键生成PPT")
        self.btn_generate_ppt.clicked.connect(self.generate_ppt)
        
        bottom_layout.addWidget(self.btn_template)
        bottom_layout.addWidget(self.btn_generate_ppt)
        layout.addLayout(bottom_layout)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

    def upload_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择文件", "", "文档 (*.txt *.docx *.pdf)")
        if not file_path:
            return
        
        ext = file_path.split('.')[-1].lower()
        content = ""
        try:
            if ext == 'txt':
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
            elif ext == 'docx':
                doc = docx.Document(file_path)
                content = "\n".join([para.text for para in doc.paragraphs])
            elif ext == 'pdf':
                with open(file_path, 'rb') as f:
                    reader = PyPDF2.PdfReader(f)
                    content = "\n".join([page.extract_text() for page in reader.pages if page.extract_text()])
            self.input_text.setText(content)
        except Exception as e:
            QMessageBox.critical(self, "读取失败", f"无法解析文件: {str(e)}")

    def generate_outline(self):
        api_key = self.api_input.text().strip()
        content = self.input_text.toPlainText().strip()
        
        if not api_key:
            QMessageBox.warning(self, "错误", "请输入 DeepSeek API Key!")
            return
        if not content:
            QMessageBox.warning(self, "错误", "请输入主题或上传文件内容!")
            return

        self.btn_gen_outline.setEnabled(False)
        self.btn_gen_outline.setText("生成中，请稍候...")
        self.progress_bar.setValue(30)

        self.outline_worker = OutlineWorker(api_key, content)
        self.outline_worker.finished.connect(self.on_outline_finished)
        self.outline_worker.error.connect(self.on_outline_error)
        self.outline_worker.start()

    def on_outline_finished(self, text):
        self.outline_text.setText(text)
        self.btn_gen_outline.setEnabled(True)
        self.btn_gen_outline.setText("✨ 第一步: AI 生成大纲")
        self.progress_bar.setValue(100)
        QMessageBox.information(self, "成功", "大纲已生成，请在文本框中检查并修改！")

    def on_outline_error(self, err):
        self.btn_gen_outline.setEnabled(True)
        self.btn_gen_outline.setText("✨ 第一步: AI 生成大纲")
        self.progress_bar.setValue(0)
        QMessageBox.critical(self, "API请求失败", str(err))

    def select_template(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择PPT模板", "", "PPTX 文件 (*.pptx)")
        if path:
            self.template_path = path
            self.btn_template.setText(f"🎨 已选: {os.path.basename(path)}")

    def generate_ppt(self):
        outline = self.outline_text.toPlainText().strip()
        if not outline:
            QMessageBox.warning(self, "错误", "大纲为空！请先生成或手动输入。")
            return

        save_path, _ = QFileDialog.getSaveFileName(self, "保存PPT", "AI_Presentation.pptx", "PPTX 文件 (*.pptx)")
        if not save_path:
            return

        self.btn_generate_ppt.setEnabled(False)
        self.btn_generate_ppt.setText("正在合成PPT与配图...")
        self.progress_bar.setValue(0)

        self.ppt_worker = PPTWorker(outline, self.template_path, save_path)
        self.ppt_worker.progress.connect(self.progress_bar.setValue)
        self.ppt_worker.finished.connect(self.on_ppt_finished)
        self.ppt_worker.error.connect(self.on_ppt_error)
        self.ppt_worker.start()

    def on_ppt_finished(self, path):
        self.btn_generate_ppt.setEnabled(True)
        self.btn_generate_ppt.setText("🚀 第二步: 一键生成PPT")
        QMessageBox.information(self, "成功", f"PPT生成完毕！\n保存位置: {path}")

    def on_ppt_error(self, err):
        self.btn_generate_ppt.setEnabled(True)
        self.btn_generate_ppt.setText("🚀 第二步: 一键生成PPT")
        QMessageBox.critical(self, "生成失败", str(err))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
