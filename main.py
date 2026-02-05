import customtkinter as ctk
import pandas as pd
import numpy as np
import threading
import os
import sys
from tkinter import filedialog, messagebox

# 设置外观模式 (System, Dark, Light)
ctk.set_appearance_mode("System")  
# 设置颜色主题 (blue, dark-blue, green)
ctk.set_default_color_theme("blue")  

class GaokaoApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # 1. 窗口基础设置
        self.title("甘肃新高考赋分系统 Pro | 俞晋全名师工作室")
        self.geometry("1100x800")
        self.minsize(900, 700)
        
        # 数据存储变量
        self.file_path = None
        self.df_raw = None     # 原始读取的数据
        self.df_working = None # 当前工作表的数据
        self.sheet_names = []
        
        # 布局容器配置
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # === 左侧边栏 (控制区) ===
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(8, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="高考赋分工具", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        # 步骤1: 导入文件
        self.btn_load = ctk.CTkButton(self.sidebar_frame, text="1. 导入Excel成绩表", command=self.load_file_action)
        self.btn_load.grid(row=1, column=0, padx=20, pady=10)

        # 步骤2: 选择工作表 (Sheet)
        self.lbl_sheet = ctk.CTkLabel(self.sidebar_frame, text="选择工作表:", anchor="w")
        self.lbl_sheet.grid(row=2, column=0, padx=20, pady=(10, 0), sticky="w")
        self.sheet_dropdown = ctk.CTkOptionMenu(self.sidebar_frame, values=[], command=self.change_sheet_event)
        self.sheet_dropdown.grid(row=3, column=0, padx=20, pady=(5, 10))
        self.sheet_dropdown.set("请先导入文件")
        self.sheet_dropdown.configure(state="disabled")

        # 步骤3: 选择班级列
        self.lbl_class = ctk.CTkLabel(self.sidebar_frame, text="选择班级列 (用于班排):", anchor="w")
        self.lbl_class.grid(row=4, column=0, padx=20, pady=(10, 0), sticky="w")
        self.class_col_dropdown = ctk.CTkOptionMenu(self.sidebar_frame, values=[])
        self.class_col_dropdown.grid(row=5, column=0, padx=20, pady=(5, 10))
        self.class_col_dropdown.set("等待加载...")

        # 底部操作按钮
        self.btn_calc = ctk.CTkButton(self.sidebar_frame, text="开始计算", fg_color="green", command=self.start_calculation)
        self.btn_calc.grid(row=9, column=0, padx=20, pady=10)
        self.btn_calc.configure(state="disabled")

        self.btn_export = ctk.CTkButton(self.sidebar_frame, text="导出结果", command=self.export_file)
        self.btn_export.grid(row=10, column=0, padx=20, pady=(0, 20))
        self.btn_export.configure(state="disabled")

        # === 右侧主内容区 ===
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        
        # 顶部状态栏
        self.status_label = ctk.CTkLabel(self.main_frame, text="就绪 - 请导入Excel文件", anchor="w", font=("Arial", 14))
        self.status_label.pack(fill="x", pady=(0, 10))

        # 滚动区域 (用于放置多选框)
        self.scroll_frame = ctk.CTkScrollableFrame(self.main_frame, label_text="科目设置 (自动识别列名)")
        self.scroll_frame.pack(fill="both", expand=True)

        # 内部容器：原始分科目
        self.lbl_raw = ctk.CTkLabel(self.scroll_frame, text="【原始计入科目】(语数外+首选):", anchor="w", font=("Arial", 12, "bold"))
        self.lbl_raw.pack(fill="x", pady=(5, 0))
        self.raw_checkboxes_frame = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
        self.raw_checkboxes_frame.pack(fill="x", pady=5)
        self.raw_checkboxes = [] # 存储复选框对象

        # 内部容器：赋分科目
        self.lbl_assign = ctk.CTkLabel(self.scroll_frame, text="【等级赋分科目】(再选科目):", anchor="w", font=("Arial", 12, "bold"))
        self.lbl_assign.pack(fill="x", pady=(20, 0))
        self.assign_checkboxes_frame = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
        self.assign_checkboxes_frame.pack(fill="x", pady=5)
        self.assign_checkboxes = [] # 存储复选框对象

        # 进度条
        self.progressbar = ctk.CTkProgressBar(self.main_frame)
        self.progressbar.pack(fill="x", pady=(20, 0))
        self.progressbar.set(0)

    # --- 逻辑处理 ---

    def load_file_action(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return
        
        self.file_path = file_path
        self.status_label.configure(text=f"正在读取: {os.path.basename(file_path)}...")
        self.progressbar.start()
        
        # 开启线程读取，防止界面卡顿
        threading.Thread(target=self.read_excel_sheets).start()

    def read_excel_sheets(self):
        try:
            excel_file = pd.ExcelFile(self.file_path)
            self.sheet_names = excel_file.sheet_names
            
            # 回到主线程更新UI
            self.after(0, self.update_sheet_ui)
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("错误", f"读取文件失败: {e}"))
            self.after(0, self.progressbar.stop)

    def update_sheet_ui(self):
        self.progressbar.stop()
        self.progressbar.set(1)
        self.status_label.configure(text=f"已加载: {os.path.basename(self.file_path)}")
        
        # 更新Sheet下拉框
        self.sheet_dropdown.configure(values=self.sheet_names, state="normal")
        self.sheet_dropdown.set(self.sheet_names[0])
        
        # 自动加载第一个Sheet的数据
        self.change_sheet_event(self.sheet_names[0])

    def change_sheet_event(self, sheet_name):
        self.status_label.configure(text=f"正在加载工作表: {sheet_name}...")
        try:
            # 读取数据
            self.df_raw = pd.read_excel(self.file_path, sheet_name=sheet_name)
            columns = self.df_raw.columns.tolist()
            
            # 更新班级下拉框
            self.class_col_dropdown.configure(values=columns)
            # 智能猜测班级列
            for col in columns:
                if "班" in str(col):
                    self.class_col_dropdown.set(col)
                    break
            else:
                self.class_col_dropdown.set(columns[0] if columns else "")

            # 生成科目复选框
            self.create_subject_checkboxes(columns)
            
            self.btn_calc.configure(state="normal")
            self.status_label.configure(text=f"工作表 {sheet_name} 加载完成，请选择科目。")
            
        except Exception as e:
            messagebox.showerror("错误", f"加载工作表失败: {e}")

    def create_subject_checkboxes(self, columns):
        # 清除旧的复选框
        for cb in self.raw_checkboxes: cb.destroy()
        for cb in self.assign_checkboxes: cb.destroy()
        self.raw_checkboxes.clear()
        self.assign_checkboxes.clear()
        
        # 常见的科目名称，用于自动勾选
        common_raw = ["语文", "数学", "英语", "物理", "历史", "外语"]
        common_assign = ["化学", "生物", "地理", "政治", "思想政治"]

        # 创建原始分复选框 (网格布局)
        for i, col in enumerate(columns):
            cb = ctk.CTkCheckBox(self.raw_checkboxes_frame, text=col)
            cb.grid(row=i//4, column=i%4, sticky="w", padx=10, pady=5)
            # 自动勾选
            if any(n in str(col) for n in common_raw):
                cb.select()
            self.raw_checkboxes.append(cb)

        # 创建赋分复选框
        for i, col in enumerate(columns):
            cb = ctk.CTkCheckBox(self.assign_checkboxes_frame, text=col)
            cb.grid(row=i//4, column=i%4, sticky="w", padx=10, pady=5)
            # 自动勾选
            if any(n in str(col) for n in common_assign):
                cb.select()
            self.assign_checkboxes.append(cb)

    # --- 核心计算逻辑 ---
    
    def start_calculation(self):
        # 获取用户选择
        self.selected_raw = [cb.cget("text") for cb in self.raw_checkboxes if cb.get() == 1]
        self.selected_assign = [cb.cget("text") for cb in self.assign_checkboxes if cb.get() == 1]
        self.selected_class_col = self.class_col_dropdown.get()

        if not self.selected_raw and not self.selected_assign:
            messagebox.showwarning("提示", "请至少选择一个科目！")
            return

        self.btn_calc.configure(state="disabled")
        self.status_label.configure(text="正在进行赋分计算...")
        self.progressbar.start()
        
        threading.Thread(target=self.run_math_logic).start()

    def run_math_logic(self):
        try:
            df = self.df_raw.copy()
            
            # 定义赋分标准
            grade_configs = [
                {'grade': 'A', 'percent': 0.15, 't_max': 100, 't_min': 86},
                {'grade': 'B', 'percent': 0.35, 't_max': 85,  't_min': 71},
                {'grade': 'C', 'percent': 0.35, 't_max': 70,  't_min': 56},
                {'grade': 'D', 'percent': 0.13, 't_max': 55,  't_min': 41},
                {'grade': 'E', 'percent': 0.02, 't_max': 40,  't_min': 30},
            ]

            def calculate_assigned_score(series):
                series_numeric = pd.to_numeric(series, errors='coerce')
                valid_scores = series_numeric.dropna()
                total_count = len(valid_scores)
                if total_count == 0: return pd.Series(index=series.index, dtype=float)

                sorted_scores = valid_scores.sort_values(ascending=False)
                assigned_result = pd.Series(index=valid_scores.index, dtype=float)
                current_idx = 0
                
                for cfg in grade_configs:
                    count = int(np.round(total_count * cfg['percent']))
                    if cfg['grade'] == 'E': count = total_count - current_idx
                    if count <= 0: continue
                    
                    end_idx = min(current_idx + count, total_count)
                    if current_idx >= end_idx: break

                    grade_indices = sorted_scores.iloc[current_idx : end_idx].index
                    grade_raw_scores = sorted_scores.iloc[current_idx : end_idx]
                    
                    Y2, Y1 = grade_raw_scores.max(), grade_raw_scores.min()
                    T2, T1 = cfg['t_max'], cfg['t_min']
                    
                    def calc_single(Y):
                        return (T2 + T1) / 2 if Y2 == Y1 else T1 + ((Y - Y1) * (T2 - T1)) / (Y2 - Y1)

                    assigned_result.loc[grade_indices] = grade_raw_scores.apply(calc_single)
                    current_idx = end_idx
                return assigned_result.round()

            def calc_ranks(dframe, col, class_col):
                if col not in dframe.columns: return
                dframe[f"{col}_年排"] = dframe[col].rank(ascending=False, method='min')
                if class_col and class_col in dframe.columns:
                    dframe[f"{col}_班排"] = dframe.groupby(class_col)[col].rank(ascending=False, method='min')

            final_score_cols = []
            
            # 1. 处理原始分
            for sub in self.selected_raw:
                df[sub] = pd.to_numeric(df[sub], errors='coerce')
                calc_ranks(df, sub, self.selected_class_col)
                final_score_cols.append(sub)

            # 2. 处理赋分
            for sub in self.selected_assign:
                assigned_col = f"{sub}_赋分"
                df[assigned_col] = calculate_assigned_score(df[sub])
                calc_ranks(df, assigned_col, self.selected_class_col)
                final_score_cols.append(assigned_col)

            # 3. 计算总分
            df["总分"] = df[final_score_cols].sum(axis=1, min_count=1)
            calc_ranks(df, "总分", self.selected_class_col)
            
            # 默认按总分排序
            df = df.sort_values("总分_年排")

            # 4. 优化列顺序 (整理表格)
            # 基础信息列 (除去成绩和排名之外的列)
            processed_cols = []
            for sub in self.selected_raw:
                processed_cols.extend([sub, f"{sub}_年排", f"{sub}_班排"])
            for sub in self.selected_assign:
                assigned_col = f"{sub}_赋分"
                processed_cols.extend([assigned_col, f"{assigned_col}_年排", f"{assigned_col}_班排"])
            
            total_cols = ["总分", "总分_年排", "总分_班排"]
            
            # 找出所有涉及的计算列，剩下的即为基础信息列(姓名、考号等)
            all_calc_cols = set(processed_cols + total_cols + self.selected_assign) 
            base_cols = [c for c in df.columns if c not in all_calc_cols]
            
            # 最终顺序: 基础信息 + (原始分+排名) + (赋分+排名) + 总分+排名
            # 注意：这里我们只保留赋分后的列，不保留赋分前的原始列(避免混淆)，如果需要保留请修改此处
            final_order = base_cols + processed_cols + total_cols
            
            # 过滤掉不存在的列（以防万一）
            final_order = [c for c in final_order if c in df.columns]
            
            self.df_result = df[final_order]

            self.after(0, self.finish_calculation)

        except Exception as e:
            self.after(0, lambda: messagebox.showerror("计算错误", str(e)))
            self.after(0, self.stop_loading_ui)

    def finish_calculation(self):
        self.stop_loading_ui()
        self.status_label.configure(text="计算完成！请点击导出。")
        self.btn_export.configure(state="normal", fg_color="green")
        messagebox.showinfo("成功", "赋分及排名计算完成！\n请点击下方【导出结果】按钮保存文件。")

    def stop_loading_ui(self):
        self.progressbar.stop()
        self.btn_calc.configure(state="normal")

    def export_file(self):
        save_path = filedialog.asksaveasfilename(
            title="保存结果",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="赋分排名结果.xlsx"
        )
        if save_path:
            try:
                self.df_result.to_excel(save_path, index=False)
                messagebox.showinfo("保存成功", f"文件已保存至:\n{save_path}")
                os.startfile(os.path.dirname(save_path))
            except Exception as e:
                messagebox.showerror("保存失败", str(e))

if __name__ == "__main__":
    app = GaokaoApp()
    app.mainloop()
