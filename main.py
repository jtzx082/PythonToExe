import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import pptx
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import requests
import json
import threading
from PIL import Image, ImageTk
import io
import re
import openai
from pptx.enum.dml import MSO_THEME_COLOR

class PPTMakerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AI智能PPT制作工具")
        self.root.geometry("1200x800")
        
        # 设置样式
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # 创建主框架
        main_frame = ttk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 左侧控制面板
        control_frame = ttk.LabelFrame(main_frame, text="控制面板", width=300)
        control_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        control_frame.pack_propagate(False)
        
        # 主题输入
        ttk.Label(control_frame, text="PPT主题:").pack(pady=5)
        self.topic_var = tk.StringVar()
        topic_entry = ttk.Entry(control_frame, textvariable=self.topic_var, width=35)
        topic_entry.pack(pady=5)
        
        # 文件上传
        ttk.Label(control_frame, text="上传文档:").pack(pady=5)
        self.file_path_var = tk.StringVar()
        file_frame = ttk.Frame(control_frame)
        file_frame.pack(pady=5)
        ttk.Entry(file_frame, textvariable=self.file_path_var, width=25).pack(side=tk.LEFT)
        ttk.Button(file_frame, text="浏览", command=self.browse_file).pack(side=tk.LEFT, padx=5)
        
        # API密钥输入
        ttk.Label(control_frame, text="DeepSeek API Key:").pack(pady=5)
        self.api_key_var = tk.StringVar()
        api_entry = ttk.Entry(control_frame, textvariable=self.api_key_var, show="*", width=35)
        api_entry.pack(pady=5)
        
        # 生成按钮
        ttk.Button(control_frame, text="生成大纲", command=self.generate_outline).pack(pady=10)
        ttk.Button(control_frame, text="生成PPT", command=self.generate_ppt).pack(pady=5)
        
        # 进度条
        self.progress = ttk.Progressbar(control_frame, mode='indeterminate')
        self.progress.pack(fill=tk.X, pady=10)
        
        # 中间大纲编辑区
        outline_frame = ttk.LabelFrame(main_frame, text="PPT大纲编辑", width=500)
        outline_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        # 大纲树形控件
        columns = ('title', 'content')
        self.outline_tree = ttk.Treeview(outline_frame, columns=columns, show='tree headings', height=20)
        self.outline_tree.heading('#0', text='幻灯片')
        self.outline_tree.heading('title', text='标题')
        self.outline_tree.heading('content', text='内容')
        self.outline_tree.column('#0', width=100)
        self.outline_tree.column('title', width=150)
        self.outline_tree.column('content', width=300)
        
        # 滚动条
        tree_scroll = ttk.Scrollbar(outline_frame, orient=tk.VERTICAL, command=self.outline_tree.yview)
        self.outline_tree.configure(yscrollcommand=tree_scroll.set)
        
        self.outline_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 添加/删除按钮
        btn_frame = ttk.Frame(outline_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        ttk.Button(btn_frame, text="添加幻灯片", command=self.add_slide).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="删除选中", command=self.delete_slide).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="上移", command=self.move_up).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="下移", command=self.move_down).pack(side=tk.LEFT, padx=5)
        
        # 右侧预览区
        preview_frame = ttk.LabelFrame(main_frame, text="预览", width=400)
        preview_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # 预览Canvas
        self.preview_canvas = tk.Canvas(preview_frame, bg='white', width=350, height=600)
        preview_scroll = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.preview_canvas.yview)
        self.preview_canvas.configure(yscrollcommand=preview_scroll.set)
        
        self.preview_frame = ttk.Frame(self.preview_canvas)
        self.preview_window = self.preview_canvas.create_window((0, 0), window=self.preview_frame, anchor="nw")
        
        self.preview_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        preview_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 绑定滚动事件
        self.preview_frame.bind("<Configure>", self.on_preview_configure)
        
        # 初始化大纲示例
        self.init_example_outline()
        
        # 存储API响应
        self.generated_outline = []
        
    def on_preview_configure(self, event):
        self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all"))
    
    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="选择文档",
            filetypes=[
                ("Text files", "*.txt"),
                ("Word documents", "*.docx"),
                ("PDF files", "*.pdf"),
                ("All files", "*.*")
            ]
        )
        if filename:
            self.file_path_var.set(filename)
    
    def init_example_outline(self):
        """初始化示例大纲"""
        example_slides = [
            {"title": "欢迎页", "content": "演示文稿标题\n副标题或作者信息"},
            {"title": "目录", "content": "1. 背景介绍\n2. 问题分析\n3. 解决方案\n4. 实施计划\n5. 总结"},
            {"title": "背景介绍", "content": "项目背景\n市场需求\n技术趋势"},
            {"title": "问题分析", "content": "现状分析\n存在问题\n影响因素"},
            {"title": "解决方案", "content": "核心方案\n实施步骤\n预期效果"},
            {"title": "总结", "content": "要点回顾\n未来展望\n致谢"}
        ]
        
        for i, slide in enumerate(example_slides):
            self.outline_tree.insert('', 'end', text=f'幻灯片 {i+1}', values=(slide['title'], slide['content']))
    
    def generate_outline(self):
        """生成PPT大纲"""
        topic = self.topic_var.get().strip()
        file_path = self.file_path_var.get().strip()
        api_key = self.api_key_var.get().strip()
        
        if not api_key:
            messagebox.showerror("错误", "请输入API密钥")
            return
        
        if not topic and not file_path:
            messagebox.showerror("错误", "请输入主题或上传文件")
            return
        
        # 启动进度条
        self.progress.start()
        
        # 在新线程中执行API调用
        thread = threading.Thread(target=self._generate_outline_thread, args=(topic, file_path, api_key))
        thread.daemon = True
        thread.start()
    
    def _generate_outline_thread(self, topic, file_path, api_key):
        try:
            # 准备提示词
            prompt = f"请为'{topic}'这个主题生成一个详细的PPT大纲，包含至少6个幻灯片。每个幻灯片应包含标题和详细内容。以JSON格式返回，格式如下：[{{'title': '幻灯片标题', 'content': '幻灯片内容'}}, ...]"
            
            # 如果有上传文件，读取内容并加入提示词
            if file_path:
                content = self.read_file_content(file_path)
                prompt = f"根据以下内容为'{topic}'这个主题生成一个详细的PPT大纲：\n{content}\n\n请返回JSON格式：[{{'title': '幻灯片标题', 'content': '幻灯片内容'}}, ...]"
            
            # 调用DeepSeek API
            headers = {
                'Content-Type': 'application/json',
                'Authorization': f'Bearer {api_key}'
            }
            
            data = {
                "model": "deepseek-chat",
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0.7
            }
            
            response = requests.post(
                "https://api.deepseek.com/chat/completions",
                headers=headers,
                json=data
            )
            
            if response.status_code == 200:
                result = response.json()
                content = result['choices'][0]['message']['content']
                
                # 提取JSON部分
                json_match = re.search(r'\[(.*?)\]', content, re.DOTALL)
                if json_match:
                    json_str = '[' + json_match.group(1) + ']'
                    self.generated_outline = json.loads(json_str)
                    
                    # 在主线程中更新界面
                    self.root.after(0, self._update_outline_ui)
                else:
                    # 尝试直接解析
                    try:
                        self.generated_outline = json.loads(content)
                        self.root.after(0, self._update_outline_ui)
                    except:
                        messagebox.showerror("错误", "无法解析API返回结果")
            else:
                error_msg = response.json().get('error', {}).get('message', '未知错误')
                self.root.after(0, lambda: messagebox.showerror("API错误", f"请求失败: {error_msg}"))
        
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", f"生成大纲时出错: {str(e)}"))
        finally:
            self.root.after(0, lambda: self.progress.stop())
    
    def _update_outline_ui(self):
        """在主线程中更新大纲UI"""
        # 清空现有内容
        for item in self.outline_tree.get_children():
            self.outline_tree.delete(item)
        
        # 添加新大纲
        for i, slide in enumerate(self.generated_outline):
            self.outline_tree.insert('', 'end', text=f'幻灯片 {i+1}', 
                                   values=(slide['title'], slide['content']))
        
        messagebox.showinfo("成功", "大纲生成完成！您可以进一步编辑大纲内容。")
    
    def read_file_content(self, file_path):
        """读取文件内容"""
        ext = os.path.splitext(file_path)[1].lower()
        
        if ext == '.txt':
            with open(file_path, 'r', encoding='utf-8') as f:
                return f.read()
        elif ext == '.docx':
            from docx import Document
            doc = Document(file_path)
            full_text = []
            for para in doc.paragraphs:
                full_text.append(para.text)
            return '\n'.join(full_text)
        elif ext == '.pdf':
            import PyPDF2
            with open(file_path, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                full_text = []
                for page in reader.pages:
                    full_text.append(page.extract_text())
                return '\n'.join(full_text)
        else:
            return ""
    
    def add_slide(self):
        """添加幻灯片"""
        item_id = self.outline_tree.selection()[0] if self.outline_tree.selection() else ''
        self.outline_tree.insert('', 'end', text='新幻灯片', values=('新标题', '新内容'))
    
    def delete_slide(self):
        """删除选中幻灯片"""
        selected_items = self.outline_tree.selection()
        if selected_items:
            for item in selected_items:
                self.outline_tree.delete(item)
    
    def move_up(self):
        """上移选中项"""
        selected = self.outline_tree.selection()
        if selected:
            item = selected[0]
            prev_item = self.outline_tree.prev(item)
            if prev_item:
                self.outline_tree.move(item, self.outline_tree.parent(item), 
                                     self.outline_tree.index(prev_item))
    
    def move_down(self):
        """下移选中项"""
        selected = self.outline_tree.selection()
        if selected:
            item = selected[0]
            next_item = self.outline_tree.next(item)
            if next_item:
                self.outline_tree.move(item, self.outline_tree.parent(item), 
                                     self.outline_tree.index(next_item)+1)
    
    def generate_ppt(self):
        """生成PPT文件"""
        slides_data = []
        for item in self.outline_tree.get_children():
            values = self.outline_tree.item(item, 'values')
            slides_data.append({
                'title': values[0],
                'content': values[1]
            })
        
        if not slides_data:
            messagebox.showwarning("警告", "请先生成或编辑PPT大纲")
            return
        
        # 选择保存位置
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pptx",
            filetypes=[("PowerPoint files", "*.pptx"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
        
        # 启动进度条
        self.progress.start()
        
        # 在新线程中生成PPT
        thread = threading.Thread(target=self._generate_ppt_thread, args=(slides_data, file_path))
        thread.daemon = True
        thread.start()
    
    def _generate_ppt_thread(self, slides_data, file_path):
        try:
            # 创建PPT
            prs = pptx.Presentation()
            
            # 设置主题颜色
            prs.slide_master.background.fill.solid()
            prs.slide_master.background.fill.fore_color.rgb = RGBColor(255, 255, 255)
            
            # 为每个幻灯片数据创建幻灯片
            for i, slide_data in enumerate(slides_data):
                # 根据内容类型选择布局
                if i == 0:  # 第一张通常是标题页
                    slide_layout = prs.slide_layouts[0]  # 标题幻灯片
                elif "目录" in slide_data['title'] or "概览" in slide_data['title']:
                    slide_layout = prs.slide_layouts[1]  # 标题和内容
                elif len(slide_data['content'].split('\n')) > 3:
                    slide_layout = prs.slide_layouts[1]  # 标题和内容
                else:
                    slide_layout = prs.slide_layouts[1]  # 标题和内容
            
                slide = prs.slides.add_slide(slide_layout)
                
                # 获取占位符
                for shape in slide.placeholders:
                    if shape.placeholder_format.type == 0:  # 标题
                        title = shape
                        title.text = slide_data['title']
                        title.text_frame.paragraphs[0].font.size = Pt(32)
                        title.text_frame.paragraphs[0].font.bold = True
                        title.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 51, 102)
                    elif shape.placeholder_format.type == 1:  # 内容
                        content = shape
                        content.text = slide_data['content']
                        content.text_frame.paragraphs[0].font.size = Pt(18)
                        # 设置行间距
                        for paragraph in content.text_frame.paragraphs:
                            paragraph.line_spacing = 1.2
            
            # 保存文件
            prs.save(file_path)
            
            self.root.after(0, lambda: [
                self.progress.stop(),
                messagebox.showinfo("成功", f"PPT已保存到: {file_path}")
            ])
        
        except Exception as e:
            self.root.after(0, lambda: [
                self.progress.stop(),
                messagebox.showerror("错误", f"生成PPT时出错: {str(e)}")
            ])

def main():
    root = tk.Tk()
    app = PPTMakerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
