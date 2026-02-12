#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
期刊论文撰写软件
支持自动生成大纲、编辑大纲、生成完整文稿
支持多种文稿类型：论文、计划、反思、案例、总结等
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import json
import os
from openai import OpenAI
from typing import List, Dict, Optional
import threading


class PaperWriterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("智能文稿撰写助手")
        self.root.geometry("1200x800")
        self.root.configure(bg='#f0f0f0')
        
        # DeepSeek API配置
        self.api_key = os.getenv("DEEPSEEK_API_KEY", "")
        self.base_url = "https://api.deepseek.com/v1"
        self.client = None
        
        # 文稿类型
        self.document_types = {
            "学术论文": "学术论文",
            "工作计划": "工作计划",
            "反思总结": "反思总结",
            "案例分析": "案例分析",
            "工作总结": "工作总结",
            "自定义": "自定义"
        }
        
        # 当前文稿数据
        self.current_title = ""
        self.current_type = "学术论文"
        self.current_outline = []
        self.custom_type_description = ""
        
        self.setup_ui()
        self.load_config()
        
    def setup_ui(self):
        """设置用户界面"""
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # API配置区域
        api_frame = ttk.LabelFrame(main_frame, text="API配置", padding="10")
        api_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        api_frame.columnconfigure(1, weight=1)
        
        ttk.Label(api_frame, text="DeepSeek API Key:").grid(row=0, column=0, padx=(0, 5))
        self.api_key_entry = ttk.Entry(api_frame, width=50, show="*")
        self.api_key_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        if self.api_key:
            self.api_key_entry.insert(0, self.api_key)
        
        ttk.Button(api_frame, text="保存配置", command=self.save_config).grid(row=0, column=2)
        
        # 文稿信息区域
        info_frame = ttk.LabelFrame(main_frame, text="文稿信息", padding="10")
        info_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        info_frame.columnconfigure(1, weight=1)
        
        ttk.Label(info_frame, text="文稿类型:").grid(row=0, column=0, padx=(0, 5), sticky=tk.W)
        self.type_var = tk.StringVar(value=self.current_type)
        type_combo = ttk.Combobox(info_frame, textvariable=self.type_var, 
                                  values=list(self.document_types.keys()), 
                                  state="readonly", width=15)
        type_combo.grid(row=0, column=1, sticky=tk.W, padx=(0, 20))
        type_combo.bind("<<ComboboxSelected>>", self.on_type_change)
        
        self.custom_type_frame = ttk.Frame(info_frame)
        self.custom_type_frame.grid(row=0, column=2, sticky=(tk.W, tk.E))
        ttk.Label(self.custom_type_frame, text="自定义类型描述:").pack(side=tk.LEFT, padx=(0, 5))
        self.custom_type_entry = ttk.Entry(self.custom_type_frame, width=30)
        self.custom_type_entry.pack(side=tk.LEFT)
        
        ttk.Label(info_frame, text="文稿标题:").grid(row=1, column=0, padx=(0, 5), sticky=tk.W, pady=(10, 0))
        self.title_entry = ttk.Entry(info_frame, width=60)
        self.title_entry.grid(row=1, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        ttk.Button(button_frame, text="生成大纲", command=self.generate_outline, 
                  width=15).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="保存大纲", command=self.save_outline, 
                  width=15).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="加载大纲", command=self.load_outline, 
                  width=15).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="开始撰写", command=self.generate_document, 
                  width=15).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="保存文稿", command=self.save_document, 
                  width=15).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="清空内容", command=self.clear_all, 
                  width=15).pack(side=tk.LEFT)
        
        # 内容区域（使用Notebook）
        content_notebook = ttk.Notebook(main_frame)
        content_notebook.grid(row=2, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(10, 0))
        
        # 大纲编辑标签页
        outline_frame = ttk.Frame(content_notebook, padding="10")
        content_notebook.add(outline_frame, text="大纲编辑")
        outline_frame.columnconfigure(0, weight=1)
        outline_frame.rowconfigure(0, weight=1)
        
        self.outline_text = scrolledtext.ScrolledText(outline_frame, wrap=tk.WORD, 
                                                      font=("Microsoft YaHei", 11))
        self.outline_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 文稿预览标签页
        document_frame = ttk.Frame(content_notebook, padding="10")
        content_notebook.add(document_frame, text="文稿预览")
        document_frame.columnconfigure(0, weight=1)
        document_frame.rowconfigure(0, weight=1)
        
        self.document_text = scrolledtext.ScrolledText(document_frame, wrap=tk.WORD, 
                                                        font=("Microsoft YaHei", 11))
        self.document_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                               relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E))
        
    def on_type_change(self, event=None):
        """文稿类型改变事件"""
        self.current_type = self.type_var.get()
        if self.current_type == "自定义":
            self.custom_type_frame.grid()
        else:
            self.custom_type_frame.grid_remove()
    
    def get_client(self):
        """获取API客户端"""
        api_key = self.api_key_entry.get().strip()
        if not api_key:
            messagebox.showerror("错误", "请先配置DeepSeek API Key")
            return None
        
        if not self.client or self.api_key != api_key:
            self.api_key = api_key
            self.client = OpenAI(
                api_key=self.api_key,
                base_url=self.base_url
            )
        return self.client
    
    def get_document_type_description(self):
        """获取文稿类型描述"""
        if self.current_type == "自定义":
            desc = self.custom_type_entry.get().strip()
            return desc if desc else "通用文稿"
        return self.document_types[self.current_type]
    
    def generate_outline(self):
        """生成大纲"""
        title = self.title_entry.get().strip()
        if not title:
            messagebox.showerror("错误", "请输入文稿标题")
            return
        
        client = self.get_client()
        if not client:
            return
        
        doc_type = self.get_document_type_description()
        self.status_var.set("正在生成大纲，请稍候...")
        
        def generate():
            try:
                prompt = f"""请为以下{doc_type}生成详细的大纲结构。

标题：{title}

请按照以下格式输出大纲：
1. 一级标题
   1.1 二级标题
   1.2 二级标题
2. 一级标题
   2.1 二级标题
   2.2 二级标题
...

请确保大纲结构清晰、完整，适合撰写{doc_type}。"""
                
                response = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[
                        {"role": "system", "content": f"你是一位专业的{doc_type}撰写专家，擅长构建清晰、逻辑严密的文稿结构。"},
                        {"role": "user", "content": prompt}
                    ],
                    temperature=0.7,
                    max_tokens=2000
                )
                
                outline = response.choices[0].message.content
                
                self.root.after(0, lambda: self.update_outline(outline))
                self.root.after(0, lambda: self.status_var.set("大纲生成完成"))
                
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("错误", f"生成大纲失败：{str(e)}"))
                self.root.after(0, lambda: self.status_var.set("大纲生成失败"))
        
        threading.Thread(target=generate, daemon=True).start()
    
    def update_outline(self, outline):
        """更新大纲显示"""
        self.outline_text.delete(1.0, tk.END)
        self.outline_text.insert(1.0, outline)
        self.current_title = self.title_entry.get().strip()
    
    def save_outline(self):
        """保存大纲"""
        outline = self.outline_text.get(1.0, tk.END).strip()
        if not outline:
            messagebox.showwarning("警告", "大纲内容为空")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
            initialfile=f"{self.current_title}_大纲.txt" if self.current_title else "大纲.txt"
        )
        
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(outline)
                messagebox.showinfo("成功", "大纲已保存")
            except Exception as e:
                messagebox.showerror("错误", f"保存失败：{str(e)}")
    
    def load_outline(self):
        """加载大纲"""
        filename = filedialog.askopenfilename(
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    outline = f.read()
                self.outline_text.delete(1.0, tk.END)
                self.outline_text.insert(1.0, outline)
                messagebox.showinfo("成功", "大纲已加载")
            except Exception as e:
                messagebox.showerror("错误", f"加载失败：{str(e)}")
    
    def generate_document(self):
        """生成完整文稿"""
        title = self.title_entry.get().strip()
        outline = self.outline_text.get(1.0, tk.END).strip()
        
        if not title:
            messagebox.showerror("错误", "请输入文稿标题")
            return
        
        if not outline:
            messagebox.showerror("错误", "请先生成或编辑大纲")
            return
        
        client = self.get_client()
        if not client:
            return
        
        doc_type = self.get_document_type_description()
        self.status_var.set("正在生成文稿，请稍候（这可能需要几分钟）...")
        self.document_text.delete(1.0, tk.END)
        self.document_text.insert(1.0, "正在生成，请稍候...\n\n")
        
        def generate():
            try:
                prompt = f"""请根据以下标题和大纲，撰写一篇完整的{doc_type}。

标题：{title}

大纲：
{outline}

要求：
1. 严格按照提供的大纲结构撰写
2. 内容要充实、专业、逻辑清晰
3. 每个章节都要有详细的内容
4. 如果是学术论文，请确保引用格式规范
5. 字数要充足，确保每个章节都有实质性内容
6. 使用中文撰写

请直接输出完整的文稿内容，不需要额外的说明。"""
                
                response = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[
                        {"role": "system", "content": f"你是一位专业的{doc_type}撰写专家，擅长撰写高质量、结构清晰、内容充实的文稿。"},
                        {"role": "user", "content": prompt}
                    ],
                    temperature=0.8,
                    max_tokens=8000
                )
                
                document = response.choices[0].message.content
                
                self.root.after(0, lambda: self.update_document(document))
                self.root.after(0, lambda: self.status_var.set("文稿生成完成"))
                
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("错误", f"生成文稿失败：{str(e)}"))
                self.root.after(0, lambda: self.status_var.set("文稿生成失败"))
        
        threading.Thread(target=generate, daemon=True).start()
    
    def update_document(self, document):
        """更新文稿显示"""
        self.document_text.delete(1.0, tk.END)
        self.document_text.insert(1.0, document)
    
    def save_document(self):
        """保存文稿"""
        document = self.document_text.get(1.0, tk.END).strip()
        if not document:
            messagebox.showwarning("警告", "文稿内容为空")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("Markdown文件", "*.md"), ("所有文件", "*.*")],
            initialfile=f"{self.current_title}.txt" if self.current_title else "文稿.txt"
        )
        
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(document)
                messagebox.showinfo("成功", "文稿已保存")
            except Exception as e:
                messagebox.showerror("错误", f"保存失败：{str(e)}")
    
    def clear_all(self):
        """清空所有内容"""
        if messagebox.askyesno("确认", "确定要清空所有内容吗？"):
            self.title_entry.delete(0, tk.END)
            self.outline_text.delete(1.0, tk.END)
            self.document_text.delete(1.0, tk.END)
            self.status_var.set("已清空")
    
    def save_config(self):
        """保存配置"""
        api_key = self.api_key_entry.get().strip()
        config = {
            "api_key": api_key,
            "document_type": self.type_var.get()
        }
        
        config_file = os.path.join(os.path.expanduser("~"), ".paper_writer_config.json")
        try:
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("成功", "配置已保存")
        except Exception as e:
            messagebox.showerror("错误", f"保存配置失败：{str(e)}")
    
    def load_config(self):
        """加载配置"""
        config_file = os.path.join(os.path.expanduser("~"), ".paper_writer_config.json")
        if os.path.exists(config_file):
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    if "api_key" in config:
                        self.api_key_entry.delete(0, tk.END)
                        self.api_key_entry.insert(0, config["api_key"])
                        self.api_key = config["api_key"]
                    if "document_type" in config:
                        self.type_var.set(config["document_type"])
                        self.current_type = config["document_type"]
            except Exception as e:
                print(f"加载配置失败：{e}")


def main():
    root = tk.Tk()
    app = PaperWriterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
