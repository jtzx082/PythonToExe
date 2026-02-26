import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk as tk_ttk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# 解决跨平台中文字体显示问题
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'Arial Unicode MS', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False

class ElectronCloudGaokaoAnalyzer:
    def __init__(self, master):
        self.master = master
        self.master.title("电子云 - 甘肃新高考(3+1+2)数据分析舱 (全功能旗舰版)")
        self.master.geometry("1300x850")
        
        # 核心数据流状态
        self.df = pd.DataFrame()
        self.cleaned_df = pd.DataFrame()
        self.raw_subjects = []       # 语数外等原始分科目
        self.assign_subjects = []    # 化生政地等需赋分科目
        self.tracks = []             # 选科方向 (物理类/历史类)
        self.thresholds = {}         # 各科类达线阈值

        self.setup_ui()

    def setup_ui(self):
        self.notebook = ttk.Notebook(self.master, bootstyle="info")
        self.notebook.pack(fill=BOTH, expand=YES, padx=10, pady=10)

        # 五大核心功能舱
        self.tab_data = ttk.Frame(self.notebook, padding=10)
        self.tab_threshold = ttk.Frame(self.notebook, padding=10)
        self.tab_report = ttk.Frame(self.notebook, padding=10)
        self.tab_chart = ttk.Frame(self.notebook, padding=10)
        self.tab_export = ttk.Frame(self.notebook, padding=10)

        self.notebook.add(self.tab_data, text=" 📂 1. 数据引擎与赋分 ")
        self.notebook.add(self.tab_threshold, text=" 🎯 2. 划线与上线率 ")
        self.notebook.add(self.tab_report, text=" 📝 3. 质量诊断大表 ")
        self.notebook.add(self.tab_chart, text=" 📊 4. 可视化大屏 ")
        self.notebook.add(self.tab_export, text=" 📤 5. 分发与导出中心 ")

        self._build_data_tab()
        self._build_threshold_tab()
        self._build_report_tab()
        self._build_chart_tab()
        self._build_export_tab()

    # ================= UI 构建层 =================

    def _build_data_tab(self):
        ctrl_frame = ttk.Labelframe(self.tab_data, text="操作面板：数据导入与赋分初始化", padding=15)
        ctrl_frame.pack(fill=X, pady=(0, 10))

        ttk.Button(ctrl_frame, text="导入教务原始成绩单 (Excel)", icon="📂", bootstyle=PRIMARY, command=self.load_data).pack(side=LEFT, padx=5)
        ttk.Button(ctrl_frame, text="执行 3+1+2 等级赋分与统算", bootstyle=SUCCESS, command=self.clean_and_compute).pack(side=LEFT, padx=5)
        
        self.data_status = ttk.Label(ctrl_frame, text="等待导入数据...", foreground="gray")
        self.data_status.pack(side=RIGHT, padx=10)

        self.tv_data = ttk.Treeview(self.tab_data, show="headings", height=20)
        self.tv_data.pack(fill=BOTH, expand=YES)

    def _build_threshold_tab(self):
        ctrl_frame = ttk.Labelframe(self.tab_threshold, text="设定各科类达线标准 (如一本线/本科线)", padding=15)
        ctrl_frame.pack(fill=X, pady=(0, 10))
        
        self.threshold_inputs_frame = ttk.Frame(ctrl_frame)
        self.threshold_inputs_frame.pack(side=LEFT, fill=X, expand=YES)
        
        ttk.Button(ctrl_frame, text="计算各班上线指标", bootstyle=WARNING, command=self.calculate_thresholds).pack(side=RIGHT, padx=15)

        self.tv_threshold = ttk.Treeview(self.tab_threshold, show="headings", height=20)
        self.tv_threshold.pack(fill=BOTH, expand=YES)

    def _build_report_tab(self):
        ctrl_frame = ttk.Frame(self.tab_report)
        ctrl_frame.pack(fill=X, pady=(0, 10))
        
        ttk.Label(ctrl_frame, text="选科方向:").pack(side=LEFT, padx=5)
        self.report_track_var = tk.StringVar()
        self.cb_report_track = ttk.Combobox(ctrl_frame, textvariable=self.report_track_var, state="readonly", width=15)
        self.cb_report_track.pack(side=LEFT, padx=5)
        
        ttk.Button(ctrl_frame, text="生成班级全科均分横向对比表", bootstyle=INFO, command=self.generate_report).pack(side=LEFT, padx=15)

        self.report_text = ttk.Text(self.tab_report, font=("Consolas", 11), padding=15)
        self.report_text.pack(fill=BOTH, expand=YES)

    def _build_chart_tab(self):
        ctrl_frame = ttk.Frame(self.tab_chart)
        ctrl_frame.pack(fill=X, pady=(0, 10))
        
        ttk.Label(ctrl_frame, text="科类:").pack(side=LEFT, padx=5)
        self.chart_track_var = tk.StringVar()
        self.cb_chart_track = ttk.Combobox(ctrl_frame, textvariable=self.chart_track_var, state="readonly", width=12)
        self.cb_chart_track.pack(side=LEFT, padx=5)

        ttk.Label(ctrl_frame, text="指标:").pack(side=LEFT, padx=5)
        self.chart_metric_var = tk.StringVar(value="3+1+2总分")
        self.cb_chart_metric = ttk.Combobox(ctrl_frame, textvariable=self.chart_metric_var, state="readonly", width=12)
        self.cb_chart_metric.pack(side=LEFT, padx=5)
        
        ttk.Button(ctrl_frame, text="一键渲染对比大图", bootstyle=SUCCESS, command=self.draw_chart).pack(side=LEFT, padx=15)

        self.canvas_frame = ttk.Frame(self.tab_chart)
        self.canvas_frame.pack(fill=BOTH, expand=YES)
        self.figure, self.ax = plt.subplots(figsize=(10, 5))
        self.figure.patch.set_facecolor('#f8f9fa')
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.canvas_frame)
        self.canvas.get_tk_widget().pack(fill=BOTH, expand=YES)

    def _build_export_tab(self):
        ctrl_frame = ttk.Labelframe(self.tab_export, text="批量分发工具 (按班主任拆分成绩单)", padding=20)
        ctrl_frame.pack(fill=BOTH, expand=YES, padx=50, pady=50)

        info_lbl = ttk.Label(ctrl_frame, text="将当前经过赋分和排名的总表，一键拆分为每个班级独立的 Excel 文件，方便下发给各班班主任核对。", font=("Microsoft YaHei", 10), wraplength=600)
        info_lbl.pack(pady=20)

        self.export_btn = ttk.Button(ctrl_frame, text="🚀 一键拆分并导出各班成绩单", bootstyle=(SUCCESS, OUTLINE), width=30, command=self.export_class_files)
        self.export_btn.pack(pady=20)

        self.export_status = ttk.Label(ctrl_frame, text="", font=("Consolas", 10), foreground="blue")
        self.export_status.pack(pady=10)

    # ================= 数据与 3+1+2 赋分逻辑 =================

    def assign_score_logic(self, series):
        """甘肃新高考等级赋分标准算法"""
        s = series.replace(0, np.nan).dropna()
        if len(s) == 0: return series

        pct = s.rank(method='min', ascending=False) / len(s)
        conditions = [
            pct <= 0.15,
            (pct > 0.15) & (pct <= 0.50),
            (pct > 0.50) & (pct <= 0.85),
            (pct > 0.85) & (pct <= 0.98),
            pct > 0.98
        ]
        assigned_ranges = [(86, 100), (71, 85), (56, 70), (41, 55), (30, 40)]
        
        result = pd.Series(index=s.index, dtype=float)
        for cond, (Y1, Y2) in zip(conditions, assigned_ranges):
            group = s[cond]
            if len(group) == 0: continue
            
            T1, T2 = group.min(), group.max()
            if T1 == T2:
                result[group.index] = round((Y1 + Y2) / 2)
            else:
                assigned = ((group - T1) / (T2 - T1)) * (Y2 - Y1) + Y1
                result[group.index] = assigned.round()

        final_series = series.copy()
        final_series.loc[result.index] = result
        return final_series.fillna(0)

    def load_data(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if not filepath: return
        try:
            self.df = pd.read_excel(filepath)
            self.data_status.config(text=f"🟢 已加载: {os.path.basename(filepath)} | 共 {len(self.df)} 条", foreground="green")
            self._update_treeview(self.tv_data, self.df.head(50))
        except Exception as e:
            messagebox.showerror("读取错误", f"无法读取文件:\n{str(e)}")

    def clean_and_compute(self):
        if self.df.empty:
            messagebox.showwarning("提示", "请先导入教务原始数据！")
            return
        try:
            df = self.df.copy()
            df.columns = df.columns.str.strip()
            
            if '科类' not in df.columns or '班级' not in df.columns:
                messagebox.showerror("规范错误", "Excel表头必须包含 '班级' 与 '科类'。")
                return

            all_num_cols = [col for col in df.columns if pd.api.types.is_numeric_dtype(df[col]) and col not in ['学号', '考号', '班级排名', '年级排名', '总分']]
            
            # 自动识别需赋分科目
            target_assign_names = ['化学', '生物', '政治', '地理']
            self.assign_subjects = [col for col in all_num_cols if any(name in col for name in target_assign_names)]
            self.raw_subjects = [col for col in all_num_cols if col not in self.assign_subjects]

            df[all_num_cols] = df[all_num_cols].fillna(0)

            # 1. 等级赋分转换
            calc_cols = []
            for sub in self.assign_subjects:
                new_col_name = f"{sub}_赋分"
                df[new_col_name] = self.assign_score_logic(df[sub])
                calc_cols.append(new_col_name)

            # 2. 合成 3+1+2 总分
            calc_cols.extend(self.raw_subjects)
            df['3+1+2总分'] = df[calc_cols].sum(axis=1)

            # 3. 双轨独立排名
            df['科类'] = df['科类'].fillna('未分科').astype(str)
            df['班级'] = df['班级'].astype(str)
            df['科类统考排名'] = df.groupby('科类')['3+1+2总分'].rank(method='min', ascending=False).astype(int)
            df['班级内排名'] = df.groupby('班级')['3+1+2总分'].rank(method='min', ascending=False).astype(int)

            self.cleaned_df = df.sort_values(['科类', '科类统考排名'])
            
            # 联动 UI 组件
            self.tracks = list(self.cleaned_df['科类'].unique())
            self.cb_report_track['values'] = self.tracks
            self.cb_chart_track['values'] = self.tracks
            if self.tracks:
                self.cb_report_track.current(0)
                self.cb_chart_track.current(0)
            
            # 图表指标加入赋分列
            chart_metrics = ['3+1+2总分'] + self.raw_subjects + [f"{sub}_赋分" for sub in self.assign_subjects]
            self.cb_chart_metric['values'] = chart_metrics

            cols_to_show = ['班级', '姓名', '科类', '3+1+2总分', '科类统考排名', '班级内排名'] + self.raw_subjects + [f"{sub}_赋分" for sub in self.assign_subjects]
            exist_cols = [c for c in cols_to_show if c in self.cleaned_df.columns]
            self._update_treeview(self.tv_data, self.cleaned_df[exist_cols])
            
            self._generate_threshold_inputs()
            messagebox.showinfo("引擎启动成功", "赋分与排名计算完毕！数据已就绪。")
        except Exception as e:
            messagebox.showerror("引擎异常", f"处理失败:\n{str(e)}")

    # ================= 业务分析逻辑 =================

    def _generate_threshold_inputs(self):
        for widget in self.threshold_inputs_frame.winfo_children():
            widget.destroy()
        self.threshold_entries = {}
        for track in self.tracks:
            frame = ttk.Frame(self.threshold_inputs_frame)
            frame.pack(side=LEFT, padx=10)
            ttk.Label(frame, text=f"{track} 目标线:").pack(side=LEFT)
            ent = ttk.Entry(frame, width=8)
            ent.insert(0, "450")
            ent.pack(side=LEFT, padx=5)
            self.threshold_entries[track] = ent

    def calculate_thresholds(self):
        if self.cleaned_df.empty: return
        try:
            for track, ent in self.threshold_entries.items():
                self.thresholds[track] = float(ent.get())
        except ValueError:
            messagebox.showerror("格式错误", "分数线必须为数字！")
            return

        df = self.cleaned_df.copy()
        df['是否达线'] = df.apply(lambda row: 1 if row['3+1+2总分'] >= self.thresholds.get(row['科类'], 0) else 0, axis=1)
        
        stats = df.groupby(['科类', '班级']).agg(班级参考人数=('3+1+2总分', 'count'), 达线人数=('是否达线', 'sum')).reset_index()
        stats['达线率'] = (stats['达线人数'] / stats['班级参考人数'] * 100).map('{:.1f}%'.format)
        stats = stats.sort_values(by=['科类', '达线人数'], ascending=[True, False])
        self._update_treeview(self.tv_threshold, stats)

    def generate_report(self):
        if self.cleaned_df.empty: return
        track = self.report_track_var.get()
        if not track: return

        self.report_text.delete(1.0, END)
        track_df = self.cleaned_df[self.cleaned_df['科类'] == track]
        
        report = f"【{track}】各平行班 全科均分横向大比武 (含赋分转换)\n"
        report += "="*90 + "\n"
        
        agg_dict = {'3+1+2总分': 'mean'}
        for sub in self.raw_subjects:
            if track_df[sub].sum() > 0: agg_dict[sub] = 'mean'
        for sub in self.assign_subjects:
            assigned_col = f"{sub}_赋分"
            if track_df[assigned_col].sum() > 0: agg_dict[assigned_col] = 'mean'
            
        class_compare = track_df.groupby('班级').agg(agg_dict).reset_index()
        for col in class_compare.columns[1:]:
            class_compare[col] = class_compare[col].map('{:.2f}'.format)
            
        class_compare = class_compare.sort_values(by='3+1+2总分', ascending=False)
        report += class_compare.to_string(index=False) + "\n\n"
        self.report_text.insert(END, report)

    def draw_chart(self):
        if self.cleaned_df.empty: return
        track = self.chart_track_var.get()
        metric = self.chart_metric_var.get()
        if not track or not metric: return

        track_df = self.cleaned_df[self.cleaned_df['科类'] == track]
        if track_df[metric].sum() == 0:
            messagebox.showwarning("无数据", f"该科类没有【{metric}】的有效成绩。")
            return

        class_means = track_df.groupby('班级')[metric].mean().sort_values(ascending=False)
        self.ax.clear()
        
        bars = self.ax.bar(class_means.index.astype(str), class_means.values, color=ttk.Style().colors.primary, alpha=0.85, width=0.6)
        self.ax.set_title(f"{track} - 各班级【{metric}】平均分", fontsize=15, pad=20, fontweight='bold', color='#333333')
        self.ax.set_ylabel("平均分", fontsize=12)
        self.ax.spines['top'].set_visible(False)
        self.ax.spines['right'].set_visible(False)
        self.ax.bar_label(bars, fmt='%.1f', padding=4)
        
        self.figure.tight_layout()
        self.canvas.draw()

    # ================= 批量导出模块 (NEW) =================
    
    def export_class_files(self):
        if self.cleaned_df.empty:
            messagebox.showwarning("提示", "长官，请先在第一步完成数据导入和赋分计算！")
            return

        # 选择保存目录
        export_dir = filedialog.askdirectory(title="选择成绩单保存文件夹")
        if not export_dir: return
        
        try:
            self.export_btn.config(state=DISABLED)
            self.export_status.config(text="正在切割数据，请稍候...", foreground="orange")
            self.master.update()

            classes = self.cleaned_df['班级'].unique()
            
            # 为了下发给班主任更清晰，我们重新排列一下导出的列顺序
            cols_to_export = ['班级', '姓名', '科类', '3+1+2总分', '班级内排名', '科类统考排名'] + self.raw_subjects + self.assign_subjects + [f"{sub}_赋分" for sub in self.assign_subjects]
            exist_cols = [c for c in cols_to_export if c in self.cleaned_df.columns]

            for cls in classes:
                # 提取特定班级数据
                class_data = self.cleaned_df[self.cleaned_df['班级'] == cls][exist_cols]
                # 按班级内排名升序排列
                class_data = class_data.sort_values('班级内排名')
                
                filename = os.path.join(export_dir, f"高二_{cls}班_成绩单.xlsx")
                class_data.to_excel(filename, index=False)

            self.export_status.config(text=f"✅ 成功！已将 {len(classes)} 个班级的成绩单导出至:\n{export_dir}", foreground="green")
            messagebox.showinfo("导出完毕", f"完美拆分！共生成 {len(classes)} 份独立的 Excel 班级成绩单。")
            
        except Exception as e:
            self.export_status.config(text="❌ 导出过程发生错误", foreground="red")
            messagebox.showerror("导出错误", f"文件导出失败，请检查文件夹权限或是否文件被占用。\n{str(e)}")
        finally:
            self.export_btn.config(state=NORMAL)

    def _update_treeview(self, tree, df):
        tree.delete(*tree.get_children())
        tree["columns"] = list(df.columns)
        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, width=80, anchor=CENTER)
        for index, row in df.iterrows():
            tree.insert("", "end", values=list(row))

if __name__ == "__main__":
    app = ttk.Window(themename="cosmo") 
    ElectronCloudGaokaoAnalyzer(app)
    app.mainloop()
