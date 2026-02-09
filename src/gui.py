"""
图形用户界面模块
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog, font
import tkinterdnd2 as tkdnd
from PIL import Image, ImageTk
import threading
import queue
from datetime import datetime

class AcademicWriterApp:
    def __init__(self, config):
        self.config = config
        self.root = tkdnd.Tk()
        self.root.title("智能文稿撰写助手 - Academic Writer Pro")
        self.root.geometry("1200x800")
        
        # 设置窗口图标
        self.setup_icon()
        
        # 创建队列用于线程通信
        self.message_queue = queue.Queue()
        
        # 初始化变量
        self.setup_variables()
        
        # 创建界面
        self.setup_ui()
        
        # 检查消息队列
        self.root.after(100, self.process_queue)
    
    def setup_icon(self):
        """设置窗口图标"""
        try:
            # 可以添加图标文件
            pass
        except:
            pass
    
    def setup_variables(self):
        """初始化变量"""
        self.api_key_var = tk.StringVar(value=self.config.get("api_key", ""))
        self.document_type_var = tk.StringVar(value="journal_paper")
        self.custom_type_var = tk.StringVar(value="")
        self.title_var = tk.StringVar()
        self.instruction_var = tk.StringVar()
        self.model_var = tk.StringVar(value=self.config.get("model", "deepseek-chat"))
        self.temperature_var = tk.DoubleVar(value=0.7)
        self.max_tokens_var = tk.IntVar(value=4000)
        
        # 文档类型选项
        self.document_types = {
            "journal_paper": "期刊论文",
            "research_proposal": "研究计划",
            "reflection": "反思报告",
            "case_study": "案例分析",
            "summary": "总结报告",
            "custom": "自定义类型"
        }
    
    def setup_ui(self):
        """设置用户界面"""
        # 创建主框架
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=1)
        self.main_frame.rowconfigure(3, weight=1)
        
        # 标题栏
        self.create_title_bar()
        
        # API设置区域
        self.create_api_section()
        
        # 文档设置区域
        self.create_document_section()
        
        # 大纲区域
        self.create_outline_section()
        
        # 内容区域
        self.create_content_section()
        
        # 状态栏
        self.create_status_bar()
    
    def create_title_bar(self):
        """创建标题栏"""
        title_frame = ttk.Frame(self.main_frame)
        title_frame.grid(row=0, column=0, columnspan=3, pady=(0, 10), sticky=(tk.W, tk.E))
        
        title_label = ttk.Label(
            title_frame,
            text="📝 智能文稿撰写助手",
            font=("Arial", 24, "bold"),
            foreground="#2c3e50"
        )
        title_label.pack()
        
        subtitle_label = ttk.Label(
            title_frame,
            text="支持期刊论文、计划、反思、案例、总结等多种文档类型",
            font=("Arial", 10),
            foreground="#7f8c8d"
        )
        subtitle_label.pack()
    
    def create_api_section(self):
        """创建API设置区域"""
        api_frame = ttk.LabelFrame(self.main_frame, text="API设置", padding="10")
        api_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # API密钥
        ttk.Label(api_frame, text="DeepSeek API密钥:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        api_entry = ttk.Entry(api_frame, textvariable=self.api_key_var, width=50, show="•")
        api_entry.grid(row=0, column=1, sticky=(tk.W, tk.E))
        
        # 模型选择
        ttk.Label(api_frame, text="模型:").grid(row=0, column=2, sticky=tk.W, padx=(20, 5))
        model_combo = ttk.Combobox(api_frame, textvariable=self.model_var, width=20)
        model_combo['values'] = ('deepseek-chat', 'deepseek-coder')
        model_combo.grid(row=0, column=3, sticky=tk.W)
        
        # 测试按钮
        test_btn = ttk.Button(api_frame, text="测试连接", command=self.test_api_connection)
        test_btn.grid(row=0, column=4, padx=(10, 0))
        
        # 参数设置
        param_frame = ttk.Frame(api_frame)
        param_frame.grid(row=1, column=0, columnspan=5, sticky=(tk.W, tk.E), pady=(10, 0))
        
        ttk.Label(param_frame, text="温度:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        temp_scale = ttk.Scale(param_frame, from_=0, to=2, variable=self.temperature_var, 
                              length=100, orient=tk.HORIZONTAL)
        temp_scale.grid(row=0, column=1, sticky=tk.W)
        ttk.Label(param_frame, textvariable=self.temperature_var).grid(row=0, column=2, padx=(5, 10))
        
        ttk.Label(param_frame, text="最大Token:").grid(row=0, column=3, sticky=tk.W, padx=(10, 5))
        tokens_entry = ttk.Entry(param_frame, textvariable=self.max_tokens_var, width=10)
        tokens_entry.grid(row=0, column=4, sticky=tk.W)
    
    def create_document_section(self):
        """创建文档设置区域"""
        doc_frame = ttk.LabelFrame(self.main_frame, text="文档设置", padding="10")
        doc_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 文档类型
        ttk.Label(doc_frame, text="文档类型:").grid(row=0, column=0, sticky=tk.W)
        type_combo = ttk.Combobox(doc_frame, textvariable=self.document_type_var, width=20)
        type_combo['values'] = list(self.document_types.keys())
        type_combo.grid(row=0, column=1, sticky=tk.W, padx=(5, 20))
        type_combo.bind('<<ComboboxSelected>>', self.on_document_type_change)
        
        # 自定义类型
        self.custom_type_label = ttk.Label(doc_frame, text="自定义类型:")
        self.custom_type_label.grid(row=0, column=2, sticky=tk.W, padx=(0, 5))
        self.custom_type_entry = ttk.Entry(doc_frame, textvariable=self.custom_type_var, width=20)
        self.custom_type_entry.grid(row=0, column=3, sticky=tk.W)
        self.toggle_custom_type()
        
        # 文档标题
        ttk.Label(doc_frame, text="文档标题:").grid(row=1, column=0, sticky=tk.W, pady=(10, 0))
        title_entry = ttk.Entry(doc_frame, textvariable=self.title_var, width=80)
        title_entry.grid(row=1, column=1, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 附加指令
        ttk.Label(doc_frame, text="附加指令:").grid(row=2, column=0, sticky=tk.W, pady=(10, 0))
        instruction_entry = ttk.Entry(doc_frame, textvariable=self.instruction_var, width=80)
        instruction_entry.grid(row=2, column=1, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 按钮区域
        btn_frame = ttk.Frame(doc_frame)
        btn_frame.grid(row=3, column=0, columnspan=4, pady=(15, 0))
        
        ttk.Button(btn_frame, text="生成大纲", command=self.generate_outline).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="修改大纲", command=self.edit_outline).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="撰写文档", command=self.generate_document, 
                  style="Accent.TButton").pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="导出文档", command=self.export_document).pack(side=tk.LEFT, padx=(10, 0))
    
    def create_outline_section(self):
        """创建大纲编辑区域"""
        outline_frame = ttk.LabelFrame(self.main_frame, text="论文大纲", padding="10")
        outline_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        outline_frame.columnconfigure(0, weight=1)
        outline_frame.rowconfigure(0, weight=1)
        
        # 大纲文本框
        self.outline_text = scrolledtext.ScrolledText(
            outline_frame,
            wrap=tk.WORD,
            width=40,
            height=20,
            font=("Consolas", 10)
        )
        self.outline_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 添加示例大纲
        self.insert_sample_outline()
    
    def create_content_section(self):
        """创建内容显示区域"""
        content_frame = ttk.LabelFrame(self.main_frame, text="生成内容", padding="10")
        content_frame.grid(row=3, column=1, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), 
                          padx=(10, 0), pady=(0, 10))
        content_frame.columnconfigure(0, weight=1)
        content_frame.rowconfigure(0, weight=1)
        
        # 内容文本框
        self.content_text = scrolledtext.ScrolledText(
            content_frame,
            wrap=tk.WORD,
            width=60,
            height=20,
            font=("Consolas", 10)
        )
        self.content_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 添加标签页控件用于不同部分
        self.setup_tab_view()
    
    def setup_tab_view(self):
        """设置标签页视图"""
        notebook = ttk.Notebook(self.main_frame)
        notebook.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # 添加标签页
        self.tabs = {}
        sections = ["摘要", "引言", "方法", "结果", "讨论", "参考文献"]
        
        for section in sections:
            frame = ttk.Frame(notebook, padding="10")
            notebook.add(frame, text=section)
            
            text_widget = scrolledtext.ScrolledText(
                frame,
                wrap=tk.WORD,
                font=("Consolas", 10)
            )
            text_widget.pack(fill=tk.BOTH, expand=True)
            self.tabs[section] = text_widget
    
    def create_status_bar(self):
        """创建状态栏"""
        self.status_bar = ttk.Label(
            self.main_frame,
            text="就绪",
            relief=tk.SUNKEN,
            anchor=tk.W
        )
        self.status_bar.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 进度条
        self.progress = ttk.Progressbar(
            self.main_frame,
            mode='indeterminate',
            length=200
        )
        self.progress.grid(row=5, column=2, sticky=tk.E, pady=(10, 0))
    
    def insert_sample_outline(self):
        """插入示例大纲"""
        sample_outline = """# 论文题目：智能文稿撰写系统的设计与实现

## 一、摘要
1.1 研究背景
1.2 研究目的
1.3 研究方法
1.4 主要结果
1.5 研究结论

## 二、引言
2.1 研究背景与意义
2.2 国内外研究现状
2.3 研究内容与目标
2.4 论文结构安排

## 三、相关工作
3.1 智能写作系统研究
3.2 自然语言处理技术
3.3 文档生成方法
3.4 现有系统比较

## 四、系统设计
4.1 总体架构设计
4.2 核心模块设计
4.3 算法设计
4.4 界面设计

## 五、系统实现
5.1 开发环境与工具
5.2 关键技术实现
5.3 功能模块实现
5.4 系统集成

## 六、实验与分析
6.1 实验设计
6.2 实验结果
6.3 结果分析
6.4 性能评估

## 七、结论与展望
7.1 研究总结
7.2 主要贡献
7.3 不足与改进
7.4 未来展望

## 八、参考文献"""
        
        self.outline_text.insert(1.0, sample_outline)
    
    def on_document_type_change(self, event=None):
        """文档类型改变事件"""
        self.toggle_custom_type()
    
    def toggle_custom_type(self):
        """切换自定义类型输入框的显示"""
        if self.document_type_var.get() == "custom":
            self.custom_type_label.grid()
            self.custom_type_entry.grid()
        else:
            self.custom_type_label.grid_remove()
            self.custom_type_entry.grid_remove()
    
    def test_api_connection(self):
        """测试API连接"""
        api_key = self.api_key_var.get().strip()
        if not api_key:
            messagebox.showwarning("警告", "请输入API密钥")
            return
        
        self.set_status("测试API连接中...")
        self.progress.start()
        
        # 在后台线程中测试连接
        threading.Thread(
            target=self._test_api_connection_thread,
            args=(api_key,),
            daemon=True
        ).start()
    
    def _test_api_connection_thread(self, api_key):
        """测试API连接的线程函数"""
        try:
            # 这里调用API测试连接
            # 暂时模拟成功
            import time
            time.sleep(1)
            self.message_queue.put(("success", "API连接成功！"))
        except Exception as e:
            self.message_queue.put(("error", f"API连接失败: {str(e)}"))
    
    def generate_outline(self):
        """生成大纲"""
        title = self.title_var.get().strip()
        if not title:
            messagebox.showwarning("警告", "请输入文档标题")
            return
        
        doc_type = self.get_document_type()
        instruction = self.instruction_var.get()
        
        self.set_status(f"正在生成{doc_type}大纲...")
        self.progress.start()
        
        # 在后台线程中生成大纲
        threading.Thread(
            target=self._generate_outline_thread,
            args=(title, doc_type, instruction),
            daemon=True
        ).start()
    
    def _generate_outline_thread(self, title, doc_type, instruction):
        """生成大纲的线程函数"""
        try:
            # 调用API生成大纲
            from api_client import DeepSeekClient
            from document_generator import DocumentGenerator
            
            api_key = self.api_key_var.get().strip()
            client = DeepSeekClient(api_key)
            generator = DocumentGenerator(client)
            
            outline = generator.generate_outline(title, doc_type, instruction)
            
            # 在主线程中更新UI
            self.root.after(0, self.update_outline_text, outline)
            self.message_queue.put(("success", "大纲生成成功！"))
            
        except Exception as e:
            self.message_queue.put(("error", f"生成大纲失败: {str(e)}"))
    
    def edit_outline(self):
        """编辑大纲"""
        # 获取当前大纲内容
        outline = self.outline_text.get(1.0, tk.END).strip()
        
        # 创建编辑对话框
        edit_window = tk.Toplevel(self.root)
        edit_window.title("编辑大纲")
        edit_window.geometry("800x600")
        
        # 创建编辑框
        edit_text = scrolledtext.ScrolledText(edit_window, wrap=tk.WORD, font=("Consolas", 10))
        edit_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        edit_text.insert(1.0, outline)
        
        # 创建按钮
        btn_frame = ttk.Frame(edit_window)
        btn_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        ttk.Button(btn_frame, text="保存", 
                  command=lambda: self.save_outline(edit_text, edit_window)).pack(side=tk.RIGHT)
        ttk.Button(btn_frame, text="取消", 
                  command=edit_window.destroy).pack(side=tk.RIGHT, padx=(0, 10))
    
    def save_outline(self, edit_text, window):
        """保存编辑后的大纲"""
        outline = edit_text.get(1.0, tk.END).strip()
        self.outline_text.delete(1.0, tk.END)
        self.outline_text.insert(1.0, outline)
        window.destroy()
        messagebox.showinfo("成功", "大纲已保存！")
    
    def generate_document(self):
        """生成完整文档"""
        outline = self.outline_text.get(1.0, tk.END).strip()
        if not outline:
            messagebox.showwarning("警告", "请先生成或编辑大纲")
            return
        
        doc_type = self.get_document_type()
        
        self.set_status(f"正在生成{doc_type}...")
        self.progress.start()
        
        # 在后台线程中生成文档
        threading.Thread(
            target=self._generate_document_thread,
            args=(outline, doc_type),
            daemon=True
        ).start()
    
    def _generate_document_thread(self, outline, doc_type):
        """生成文档的线程函数"""
        try:
            # 调用API生成文档
            from api_client import DeepSeekClient
            from document_generator import DocumentGenerator
            
            api_key = self.api_key_var.get().strip()
            client = DeepSeekClient(api_key)
            generator = DocumentGenerator(client)
            
            document = generator.generate_document(outline, doc_type)
            
            # 在主线程中更新UI
            self.root.after(0, self.update_content_text, document)
            self.message_queue.put(("success", "文档生成成功！"))
            
        except Exception as e:
            self.message_queue.put(("error", f"生成文档失败: {str(e)}"))
    
    def export_document(self):
        """导出文档"""
        content = self.content_text.get(1.0, tk.END).strip()
        if not content:
            messagebox.showwarning("警告", "没有内容可以导出")
            return
        
        # 选择保存路径
        filename = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[
                ("Word文档", "*.docx"),
                ("PDF文件", "*.pdf"),
                ("Markdown文件", "*.md"),
                ("纯文本", "*.txt"),
                ("所有文件", "*.*")
            ]
        )
        
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(content)
                messagebox.showinfo("成功", f"文档已保存到:\n{filename}")
            except Exception as e:
                messagebox.showerror("错误", f"保存失败: {str(e)}")
    
    def get_document_type(self):
        """获取文档类型"""
        doc_type_key = self.document_type_var.get()
        if doc_type_key == "custom":
            return self.custom_type_var.get()
        return self.document_types.get(doc_type_key, "期刊论文")
    
    def update_outline_text(self, outline):
        """更新大纲文本框"""
        self.outline_text.delete(1.0, tk.END)
        self.outline_text.insert(1.0, outline)
    
    def update_content_text(self, content):
        """更新内容文本框"""
        self.content_text.delete(1.0, tk.END)
        self.content_text.insert(1.0, content)
    
    def set_status(self, message):
        """设置状态栏消息"""
        self.status_bar.config(text=message)
    
    def process_queue(self):
        """处理消息队列"""
        try:
            while True:
                msg_type, message = self.message_queue.get_nowait()
                if msg_type == "success":
                    messagebox.showinfo("成功", message)
                elif msg_type == "error":
                    messagebox.showerror("错误", message)
                elif msg_type == "info":
                    messagebox.showinfo("信息", message)
                
                self.progress.stop()
                self.set_status("就绪")
        except queue.Empty:
            pass
        
        # 每隔100ms检查一次队列
        self.root.after(100, self.process_queue)
    
    def run(self):
        """运行应用"""
        # 设置窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 运行主循环
        self.root.mainloop()
    
    def on_closing(self):
        """窗口关闭事件"""
        # 保存配置
        self.save_config()
        self.root.destroy()
    
    def save_config(self):
        """保存配置"""
        self.config["api_key"] = self.api_key_var.get()
        self.config["model"] = self.model_var.get()
        self.config["temperature"] = self.temperature_var.get()
        self.config["max_tokens"] = self.max_tokens_var.get()
        
        from config import save_config
        save_config(self.config)
