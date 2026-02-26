import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk as tk_ttk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from openpyxl.styles import Alignment, Border, Side 

# 解决跨平台中文字体显示问题
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'Arial Unicode MS', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False

class ReverieOfCopperSulfateAnalyzer:
    def __init__(self, master):
        self.master = master
        self.master.title("硫酸铜的遐想 - 甘肃新高考(3+1+2)数据分析舱 (全维典藏版)")
        self.master.geometry("1550x900") 
        
        self.df = pd.DataFrame()
        self.cleaned_df = pd.DataFrame()
        
        self.tracks = []             
        self.thresholds = {}         
        self.top_n_target = 50       
        self.score_bins_list = [0, 400, 450, 500, 550, 600] 
        self.margin_tekong = 15 
        self.margin_benke = 20  
        self.exist_cols = []
        
        self.track_valid_cols_map = {}
        self.track_raw_subjects = {}
        self.track_assign_subjects = {}
        self.track_calc_cols = {}

        self.assign_rules = [
            {"level": "A", "pct": 15, "min": 86, "max": 100},
            {"level": "B", "pct": 35, "min": 71, "max": 85},
            {"level": "C", "pct": 35, "min": 56, "max": 70},
            {"level": "D", "pct": 13, "min": 41, "max": 55},
            {"level": "E", "pct": 2,  "min": 30, "max": 40}
        ]

        self.setup_ui()

    def setup_ui(self):
        self.notebook = ttk.Notebook(self.master, bootstyle="info")
        self.notebook.pack(fill=BOTH, expand=YES, padx=20, pady=20)

        self.tab_data = ttk.Frame(self.notebook, padding=25)
        self.tab_kpi = ttk.Frame(self.notebook, padding=25)
        self.tab_report = ttk.Frame(self.notebook, padding=25)
        self.tab_chart = ttk.Frame(self.notebook, padding=25)
        self.tab_export = ttk.Frame(self.notebook, padding=25)
        self.tab_help = ttk.Frame(self.notebook, padding=25) 

        self.notebook.add(self.tab_data, text=" 📂 1. 数据洗算引擎 ")
        self.notebook.add(self.tab_kpi, text=" 🎯 2. 双线参数总控 ")
        self.notebook.add(self.tab_report, text=" 📝 3. 质量诊断大表 ")
        self.notebook.add(self.tab_chart, text=" 📊 4. 可视化大屏 ")
        self.notebook.add(self.tab_export, text=" 📤 5. 商业报表导出 ")
        self.notebook.add(self.tab_help, text=" 📖 6. 关于与算法释义 ")

        self._build_data_tab()
        self._build_kpi_tab()
        self._build_report_tab()
        self._build_chart_tab()
        self._build_export_tab()
        self._build_help_tab()

    # ================= UI 构建层 =================

    def _build_data_tab(self):
        ctrl_frame = ttk.Labelframe(self.tab_data, text=" 第一步：数据导入与动态赋分引擎 ", padding=20, bootstyle="info")
        ctrl_frame.pack(fill=X, pady=(0, 20))

        btn_frame = ttk.Frame(ctrl_frame)
        btn_frame.pack(side=LEFT)

        ttk.Button(btn_frame, text="📂 导入教务原始成绩单", bootstyle=PRIMARY, width=22, command=self.load_data).pack(side=LEFT, padx=10)
        ttk.Button(btn_frame, text="🔧 自定义赋分比例", bootstyle=WARNING, width=18, command=self.open_assign_rules_dialog).pack(side=LEFT, padx=10)
        ttk.Button(btn_frame, text="🚀 设定各科类赋分规则并统算", bootstyle=SUCCESS, width=28, command=self.open_config_dialog).pack(side=LEFT, padx=10)
        
        self.data_status = ttk.Label(ctrl_frame, text="🟢 等待导入数据...", font=("Microsoft YaHei", 11), foreground="gray")
        self.data_status.pack(side=RIGHT, padx=20)

        tv_frame = ttk.Frame(self.tab_data)
        tv_frame.pack(fill=BOTH, expand=YES)
        
        x_scroll = ttk.Scrollbar(tv_frame, orient=HORIZONTAL)
        y_scroll = ttk.Scrollbar(tv_frame, orient=VERTICAL)
        self.tv_data = ttk.Treeview(tv_frame, show="headings", bootstyle="info", xscrollcommand=x_scroll.set, yscrollcommand=y_scroll.set)
        
        x_scroll.config(command=self.tv_data.xview)
        y_scroll.config(command=self.tv_data.yview)
        
        x_scroll.pack(fill=X, side=BOTTOM)
        y_scroll.pack(fill=Y, side=RIGHT)
        self.tv_data.pack(fill=BOTH, expand=YES)

    def _build_kpi_tab(self):
        ctrl_frame = ttk.Labelframe(self.tab_kpi, text=" 第二步：双线全局配置 (严密对齐的参数矩阵) ", padding=25, bootstyle="info")
        ctrl_frame.pack(fill=X, pady=(0, 20))
        
        self.threshold_inputs_frame = ttk.Frame(ctrl_frame)
        self.threshold_inputs_frame.grid(row=0, column=0, sticky=W, padx=(0, 20))
        
        ttk.Separator(ctrl_frame, orient=VERTICAL).grid(row=0, column=1, sticky=NS, padx=20)
        
        top_frame = ttk.Frame(ctrl_frame)
        top_frame.grid(row=0, column=2, sticky=W)
        
        ttk.Label(top_frame, text="⚠️ 特控边缘生(±分):", font=("Microsoft YaHei", 10)).grid(row=0, column=0, padx=5, pady=5, sticky=E)
        self.ent_margin_tekong = ttk.Entry(top_frame, width=8, justify=CENTER)
        self.ent_margin_tekong.insert(0, "15")
        self.ent_margin_tekong.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(top_frame, text="本科边缘生(±分):", font=("Microsoft YaHei", 10)).grid(row=1, column=0, padx=5, pady=5, sticky=E)
        self.ent_margin_benke = ttk.Entry(top_frame, width=8, justify=CENTER)
        self.ent_margin_benke.insert(0, "20")
        self.ent_margin_benke.grid(row=1, column=1, padx=5, pady=5)

        ttk.Separator(ctrl_frame, orient=VERTICAL).grid(row=0, column=3, sticky=NS, padx=20)

        right_frame = ttk.Frame(ctrl_frame)
        right_frame.grid(row=0, column=4, sticky=W)

        ttk.Label(right_frame, text="🏆 统计前 N 名:", font=("Microsoft YaHei", 10)).grid(row=0, column=0, padx=5, pady=5, sticky=E)
        self.ent_top_n = ttk.Entry(right_frame, width=20)
        self.ent_top_n.insert(0, "50")
        self.ent_top_n.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(right_frame, text="📶 分数段切分点:", font=("Microsoft YaHei", 10)).grid(row=1, column=0, padx=5, pady=5, sticky=E)
        self.ent_score_bins = ttk.Entry(right_frame, width=20)
        self.ent_score_bins.insert(0, "400, 450, 500, 550, 600")
        self.ent_score_bins.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Button(ctrl_frame, text="📊 刷新大盘指标", bootstyle=WARNING, width=20, command=self.calculate_kpi).grid(row=0, column=5, padx=40, sticky=E)

        self.tv_kpi = ttk.Treeview(self.tab_kpi, show="headings", bootstyle="info")
        self.tv_kpi.pack(fill=BOTH, expand=YES)

    def _build_report_tab(self):
        ctrl_frame = ttk.Frame(self.tab_report, padding=(0,0,0,20))
        ctrl_frame.pack(fill=X)
        
        ttk.Label(ctrl_frame, text="当前诊断视图:", font=("Microsoft YaHei", 11, "bold")).pack(side=LEFT, padx=(0,10))
        self.report_track_var = tk.StringVar()
        self.cb_report_track = ttk.Combobox(ctrl_frame, textvariable=self.report_track_var, state="readonly", width=20, font=("Microsoft YaHei", 11))
        self.cb_report_track.pack(side=LEFT, padx=5)
        
        ttk.Button(ctrl_frame, text="📝 生成班级多维质量诊断报告", bootstyle=INFO, command=self.generate_report).pack(side=LEFT, padx=25)

        text_frame = ttk.Frame(self.tab_report)
        text_frame.pack(fill=BOTH, expand=YES)
        
        self.report_text = ttk.Text(text_frame, font=("Consolas", 12), padx=20, pady=20, wrap="none", relief=FLAT, bg="#f8f9fa")
        x_scroll = ttk.Scrollbar(text_frame, orient=HORIZONTAL, command=self.report_text.xview)
        y_scroll = ttk.Scrollbar(text_frame, orient=VERTICAL, command=self.report_text.yview)
        
        self.report_text.configure(xscrollcommand=x_scroll.set, yscrollcommand=y_scroll.set)
        x_scroll.pack(fill=X, side=BOTTOM)
        y_scroll.pack(fill=Y, side=RIGHT)
        self.report_text.pack(fill=BOTH, expand=YES)

    def _build_chart_tab(self):
        ctrl_frame = ttk.Labelframe(self.tab_chart, text=" 可视化参数配置 ", padding=20, bootstyle="info")
        ctrl_frame.pack(fill=X, pady=(0, 20))
        
        ttk.Label(ctrl_frame, text="科类:", font=("Microsoft YaHei", 11)).pack(side=LEFT, padx=(10,5))
        self.chart_track_var = tk.StringVar()
        self.cb_chart_track = ttk.Combobox(ctrl_frame, textvariable=self.chart_track_var, state="readonly", width=15)
        self.cb_chart_track.pack(side=LEFT, padx=5)
        self.cb_chart_track.bind("<<ComboboxSelected>>", self._on_chart_track_change)
        
        ttk.Label(ctrl_frame, text="图表类型:", font=("Microsoft YaHei", 11)).pack(side=LEFT, padx=(25,5))
        self.chart_type_var = tk.StringVar(value="各班均分横向对比(柱状图)")
        self.cb_chart_type = ttk.Combobox(ctrl_frame, textvariable=self.chart_type_var, state="readonly", width=28, 
                                          values=["各班均分横向对比(柱状图)", "年级总分分层分布(直方图)"])
        self.cb_chart_type.pack(side=LEFT, padx=5)

        ttk.Label(ctrl_frame, text="学科指标:", font=("Microsoft YaHei", 11)).pack(side=LEFT, padx=(25,5))
        self.chart_metric_var = tk.StringVar(value="3+1+2总分")
        self.cb_chart_metric = ttk.Combobox(ctrl_frame, textvariable=self.chart_metric_var, state="readonly", width=15)
        self.cb_chart_metric.pack(side=LEFT, padx=5)
        
        ttk.Button(ctrl_frame, text="📈 渲染大屏", bootstyle=SUCCESS, width=20, command=self.draw_chart).pack(side=RIGHT, padx=20)

        canvas_border = ttk.Frame(self.tab_chart, bootstyle="secondary", padding=2)
        canvas_border.pack(fill=BOTH, expand=YES)
        self.canvas_frame = ttk.Frame(canvas_border)
        self.canvas_frame.pack(fill=BOTH, expand=YES)
        self.figure, self.ax = plt.subplots(figsize=(10, 5))
        self.figure.patch.set_facecolor('#ffffff') 
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.canvas_frame)
        self.canvas.get_tk_widget().pack(fill=BOTH, expand=YES)

    def _build_export_tab(self):
        container = ttk.Frame(self.tab_export)
        container.pack(expand=YES, fill=BOTH)
        
        card = ttk.Frame(container, padding=50)
        card.pack(expand=YES)

        title_lbl = ttk.Label(card, text="商业级教务全矩阵报表中心", font=("Microsoft YaHei", 26, "bold"), bootstyle=PRIMARY)
        title_lbl.pack(pady=(0, 35))

        info_text = (
            "🚀 硫酸铜的遐想·专属定制流：点击下方按钮，系统将为你瞬间输出高阶全系数据：\n\n"
            "  👤 1. [全息学生成绩单]：含无下划线精美排名、及【优势/薄弱】靶向诊断。\n"
            "  📊 2. [双线考核KPI]：特控/本科双线达标率及 Top N 贡献榜。\n"
            "  ⚠️ 3. [干预追踪雷达]：双轨临界生名单独立提取，高亮待提分科目。\n"
            "  📈 4. [教学离散透视]：囊括及格率、优秀率、均分、标准差及极值的分析矩阵。\n"
            "  📁 5. [自动化分发夹]：全自动生成按班级切割的汇报材料夹。\n"
            "  ✨ 6. [商业级排版化]：自动全局居中、自适应列宽、自动换行、全单元格边框封装！"
        )
        info_lbl = ttk.Label(card, text=info_text, font=("Microsoft YaHei", 12), wraplength=850, justify=LEFT)
        info_lbl.pack(pady=(0, 40))

        self.export_btn = ttk.Button(card, text="⬇ 一键生成排版级商业大表及拆分文件", bootstyle=SUCCESS, width=45, command=self.export_all_reports)
        self.export_btn.pack(pady=15)

        self.export_status = ttk.Label(card, text="准备就绪...", font=("Consolas", 11), foreground="gray")
        self.export_status.pack(pady=20)

    def _build_help_tab(self):
        main_frame = ttk.Frame(self.tab_help)
        main_frame.pack(fill=BOTH, expand=YES, padx=40, pady=20)

        header = ttk.Frame(main_frame)
        header.pack(fill=X, pady=(0, 10))
        ttk.Label(header, text="硫酸铜的遐想", font=("Microsoft YaHei", 28, "bold"), foreground="#0078D7").pack(anchor=W)
        ttk.Label(header, text="甘肃新高考(3+1+2)数据分析舱 · 考务终极引擎", font=("Microsoft YaHei", 14), foreground="gray").pack(anchor=W, pady=(5,0))
        ttk.Separator(main_frame).pack(fill=X, pady=10)

        content = ttk.Frame(main_frame)
        content.pack(fill=BOTH, expand=YES)
        
        y_scroll_help = ttk.Scrollbar(content, orient=VERTICAL)
        txt = tk.Text(content, font=("Microsoft YaHei", 10), wrap=WORD, bg="#f8f9fa", relief=FLAT, padx=20, pady=15, spacing2=6, yscrollcommand=y_scroll_help.set)
        y_scroll_help.config(command=txt.yview)
        
        txt.pack(side=LEFT, fill=BOTH, expand=YES, pady=5)
        y_scroll_help.pack(side=RIGHT, fill=Y, pady=5)

        guide_text = (
            "👨‍💻 【开发者信息】\n"
            "• 核心架构：俞晋全 (俞晋全高中化学名师工作室)\n"
            "• 官方博客：硫酸铜的遐想\n"
            "• 研发寄语：用数据驱动精准教学，用代码解放教务生产力。\n\n"
            "======================================================================\n\n"
            "📊 【附录一：核心教研指标与算法释义】\n"
            "为了确保年级统一考评标准的科学性，本系统采用标准统计学算法：\n\n"
            "1. 三率计算 (及格率 / 优秀率)\n"
            "   • 系统自动识别满分标准：语、数、外(英)系统默认满分计 150 分，其余物理、历史、化学等学科满分计 100 分。\n"
            "   • 及格线标准：得分 ≥ 卷面满分的 60% (即150分制及格线为90分，100分制及格线为60分)。\n"
            "   • 优秀线标准：得分 ≥ 卷面满分的 80% (即150分制优秀线为120分，100分制优秀线为80分)。\n"
            "   • 公式：及格/优秀率 = (达到该标准人数 ÷ 该科实际参考有效人数) × 100%\n\n"
            "2. 标准差 (教学离散度分析)\n"
            "   • 物理意义：标准差反映了一个班级内学生成绩的“两极分化”程度，是极其重要的教学诊断指标。\n"
            "   • 诊断指南：\n"
            "     - 标准差越【小】：说明该班学生该科成绩紧密围绕平均分，整体水平整齐，未出现断层。\n"
            "     - 标准差越【大】：说明该班学生成绩高度分散，高分与低分差距极其悬殊，班级内“偏科严重”或“尾巴过长”。班主任与任课教师应重点关注此指标，适时调整培优补差策略。\n\n"
            "3. 优势/薄弱学科 靶向诊断引擎\n"
            "   • 系统通过计算学生单科在全年级(同科类)中的“百分比击败率 (Percentile Rank)”，而非简单的绝对分数高低来评判。\n"
            "   • 诊断原理：某科击败的年级人数比例最高的学科，系统判定为【优势学科】；击败比例最低的学科，判定为【薄弱学科(亟待提升)】。此算法完美排除了各科试卷难度不同造成的分数误差。\n\n"
            "======================================================================\n\n"
            "📝 【附录二：甘肃省新高考(3+1+2)等级赋分原理说明】\n"
            "系统内置的等级赋分机制，严格遵循甘肃等省份的高考标准：\n\n"
            "1. 位次定等：将该选考科目考生的原始卷面分从高到低排序，按规定比例划分至 A、B、C、D、E 五个等级。\n"
            "   • 默认比例：A(15%)、B(35%)、C(35%)、D(13%)、E(2%)。支持在系统面板内修改。\n\n"
            "2. 确定区间：每个等级对应一个法定的赋分区间，满分为100分，起点分为30分。\n"
            "   • 默认区间：A(100~86)、B(85~71)、C(70~56)、D(55~41)、E(40~30)。\n\n"
            "3. 等比例换算：采用线性等比例法则，将考生的原始分映射到所在等级的赋分区间，四舍五入取整。\n"
            "   • 核心公式：(T2 - T) / (T - T1) = (Y2 - Y) / (Y - Y1)\n"
            "   • 变量释义：T 为考生的原始分；T1、T2 分别为该生所在等级内全体考生的最低、最高原始分；Y 为换算后的最终赋分；Y1、Y2 分别为该等级规定的最低、最高赋分值。\n"
            "   • 结论：同等级内，原始分越高，赋分越高。赋分机制消除了各学科试题难度差异导致的不公，全省位次才是成绩的核心体现。"
        )
        
        txt.insert(END, guide_text)
        txt.configure(state=DISABLED)

    # ================= 🚀 多Sheet选择与导入逻辑 =================

    def load_data(self):
        filepath = filedialog.askopenfilename(
            title="选择成绩单数据",
            filetypes=[("Excel Files", "*.xlsx"), ("Excel 97-2003", "*.xls"), ("All Files", "*.*")]
        )
        if not filepath: return
        
        try:
            xls = pd.ExcelFile(filepath)
            sheet_names = xls.sheet_names
            
            if len(sheet_names) > 1:
                self._open_sheet_selector(xls, filepath, sheet_names)
            else:
                self._execute_load_dataframe(xls, filepath, sheet_names[0])
        except Exception as e:
            messagebox.showerror("读取错误", f"无法解析该 Excel 文件:\n{str(e)}")

    def _open_sheet_selector(self, xls, filepath, sheet_names):
        dialog = tk.Toplevel(self.master)
        dialog.title("检测到多个工作表 (Sheet)")
        dialog.geometry("500x280")
        dialog.grab_set()

        ttk.Label(dialog, text="📄 该 Excel 文件包含多个工作表", font=("Microsoft YaHei", 12, "bold"), bootstyle=PRIMARY).pack(pady=(25, 10))
        ttk.Label(dialog, text="请在下方选择您要分析的成绩单所在 Sheet：", font=("Microsoft YaHei", 10)).pack(pady=5)

        sheet_var = tk.StringVar(value=sheet_names[0])
        cb = ttk.Combobox(dialog, textvariable=sheet_var, values=sheet_names, state="readonly", font=("Microsoft YaHei", 11), width=30)
        cb.pack(pady=15)

        def on_confirm():
            selected_sheet = sheet_var.get()
            dialog.destroy()
            self._execute_load_dataframe(xls, filepath, selected_sheet)

        ttk.Button(dialog, text="✔ 确认选择并导入", bootstyle=SUCCESS, width=25, command=on_confirm).pack(pady=20)

    def _execute_load_dataframe(self, xls, filepath, sheet_name):
        try:
            self.df = pd.read_excel(xls, sheet_name=sheet_name)
            filename = os.path.basename(filepath)
            self.data_status.config(text=f"🟢 已加载: {filename} [Sheet: {sheet_name}] | 共 {len(self.df)} 条", foreground="green")
            self._update_treeview(self.tv_data, self.df.head(50))
        except Exception as e:
            messagebox.showerror("读取错误", f"无法读取指定工作表:\n{str(e)}")

    # ================= 🔧 赋分参数与弹窗引擎 =================

    def open_assign_rules_dialog(self):
        dialog = tk.Toplevel(self.master)
        dialog.title("🔧 自定义赋分比例与区间参数")
        dialog.geometry("600x420")
        dialog.grab_set()

        top_lbl = ttk.Label(dialog, text="请根据当年考试院最新政策调整赋分模型", font=("Microsoft YaHei", 11, "bold"), bootstyle=PRIMARY)
        top_lbl.pack(pady=15)

        form_frame = ttk.Frame(dialog)
        form_frame.pack(padx=20, pady=10)

        headers = ["等级", "人数比例(%)", "赋分下限", "赋分上限"]
        for col, h in enumerate(headers):
            ttk.Label(form_frame, text=h, font=("Microsoft YaHei", 10, "bold")).grid(row=0, column=col, padx=15, pady=10)

        self.rule_entries = []
        for row, rule in enumerate(self.assign_rules, start=1):
            lbl_level = ttk.Label(form_frame, text=f"【 {rule['level']} 】", font=("Microsoft YaHei", 10, "bold"), bootstyle=INFO)
            lbl_level.grid(row=row, column=0, pady=8)
            
            ent_pct = ttk.Entry(form_frame, width=10, justify=CENTER)
            ent_pct.insert(0, str(rule['pct']))
            ent_pct.grid(row=row, column=1, pady=8)
            
            ent_min = ttk.Entry(form_frame, width=10, justify=CENTER)
            ent_min.insert(0, str(rule['min']))
            ent_min.grid(row=row, column=2, pady=8)
            
            ent_max = ttk.Entry(form_frame, width=10, justify=CENTER)
            ent_max.insert(0, str(rule['max']))
            ent_max.grid(row=row, column=3, pady=8)
            
            self.rule_entries.append((rule['level'], ent_pct, ent_min, ent_max))
            
        def save_rules():
            try:
                new_rules = []
                total_pct = 0
                for level, epct, emin, emax in self.rule_entries:
                    p, mi, ma = float(epct.get()), float(emin.get()), float(emax.get())
                    if mi > ma: raise ValueError(f"[{level}]等级的下限不能大于上限！")
                    total_pct += p
                    new_rules.append({"level": level, "pct": p, "min": mi, "max": ma})
                
                if abs(total_pct - 100) > 0.1:
                    messagebox.showwarning("比例警告", f"注意：当前比例总和为 {total_pct}%，非100%，请确保这是您的意图。", parent=dialog)
                    
                self.assign_rules = new_rules
                messagebox.showinfo("保存成功", "自定义赋分参数已保存！请执行统算生效。", parent=dialog)
                dialog.destroy()
            except ValueError as e:
                messagebox.showerror("输入错误", f"格式不正确：\n{str(e)}", parent=dialog)

        ttk.Button(dialog, text="💾 保存并应用参数", bootstyle=SUCCESS, width=30, command=save_rules).pack(pady=20)


    def open_config_dialog(self):
        if self.df.empty:
            messagebox.showwarning("提示", "俞老师，请先导入成绩单数据！")
            return

        df = self.df.copy()
        df.columns = df.columns.astype(str).str.strip() 
        if '科类' not in df.columns or '班级' not in df.columns:
            messagebox.showerror("规范错误", "Excel表头必须包含 '班级' 与 '科类'。")
            return

        tracks = df['科类'].fillna('未分科').astype(str).unique()
        exclude_cols = ['班级', '姓名', '学号', '考号', '性别', '科类', '总分', '班级排名', '年级排名', '科类排名', '班级内排名', '科类统考排名', '优势学科', '薄弱学科']
        potential_cols = [c for c in df.columns if c not in exclude_cols and not c.endswith('班排') and not c.endswith('级排')]
        
        self.rule_vars = {}
        self.track_valid_cols_map = {}
        
        dialog = tk.Toplevel(self.master)
        dialog.title("高阶统算：因科制宜 - 设定计分与赋分规则")
        dialog.geometry("800x650")
        dialog.grab_set()

        header = ttk.Frame(dialog, padding=15)
        header.pack(fill=X)
        ttk.Label(header, text="⚙️ 考务定制化计分模型", font=("Microsoft YaHei", 14, "bold"), bootstyle=PRIMARY).pack(anchor=W)
        ttk.Label(header, text="系统已为您自动剥离无人选考的无效科目，请分别为下方科目指定计算模式。", font=("Microsoft YaHei", 10), foreground="gray").pack(anchor=W, pady=(5,0))

        notebook = ttk.Notebook(dialog, bootstyle="info")
        notebook.pack(fill=BOTH, expand=YES, padx=20, pady=5)
        
        for t in tracks:
            track_df = df[df['科类'] == t]
            t_cols = [c for c in potential_cols if pd.to_numeric(track_df[c], errors='coerce').sum() > 0]
            self.track_valid_cols_map[t] = t_cols
            
            frame = ttk.Frame(notebook, padding=20)
            notebook.add(frame, text=f" {t} 规则配置 ")
            self.rule_vars[t] = {}
            
            for sub in t_cols:
                row_frame = ttk.Frame(frame)
                row_frame.pack(fill=X, pady=6)
                ttk.Label(row_frame, text=f"【{sub}】", width=12, font=("Microsoft YaHei", 11, "bold")).pack(side=LEFT)
                
                var = tk.StringVar()
                if any(n in sub for n in ['化学', '生物', '政治', '地理']): var.set("等级赋分")
                else: var.set("直接计分")
                self.rule_vars[t][sub] = var
                
                ttk.Radiobutton(row_frame, text="原始分 (直接计入总分)", variable=var, value="直接计分", bootstyle="primary").pack(side=LEFT, padx=15)
                ttk.Radiobutton(row_frame, text="转换分 (执行等级赋分)", variable=var, value="等级赋分", bootstyle="success").pack(side=LEFT, padx=15)

        ttk.Button(dialog, text="🚀 确认规则并启动全景统算引擎", bootstyle=SUCCESS, width=40, 
                   command=lambda: self._execute_computation(dialog, df, tracks)).pack(pady=20)

    # ================= 动态统算逻辑 =================

    def assign_score_logic(self, series):
        s = series.replace(0, np.nan).dropna()
        if len(s) == 0: return series

        pct = s.rank(method='min', ascending=False) / len(s)
        conditions, assigned_ranges = [], []
        cum_pct = 0.0
        
        for i, rule in enumerate(self.assign_rules):
            lower_bound = cum_pct
            cum_pct += float(rule['pct']) / 100.0
            if i == len(self.assign_rules) - 1: cond = pct > lower_bound
            elif i == 0: cond = pct <= cum_pct
            else: cond = (pct > lower_bound) & (pct <= cum_pct)
            conditions.append(cond)
            assigned_ranges.append((float(rule['min']), float(rule['max'])))
            
        result = pd.Series(index=s.index, dtype=float)
        for cond, (Y1, Y2) in zip(conditions, assigned_ranges):
            group = s[cond]
            if len(group) == 0: continue
            T1, T2 = group.min(), group.max()
            if T1 == T2: result[group.index] = round((Y1 + Y2) / 2)
            else: result[group.index] = (((group - T1) / (T2 - T1)) * (Y2 - Y1) + Y1).round()

        final_series = series.copy()
        final_series.loc[result.index] = result
        return final_series.fillna(0)

    def _execute_computation(self, dialog, df, tracks):
        track_rules = {}
        for t in tracks:
            track_rules[t] = {sub: self.rule_vars[t][sub].get() for sub in self.track_valid_cols_map[t]}
        
        dialog.destroy()
        
        try:
            processed_dfs = []
            self.track_calc_cols = {}
            self.track_raw_subjects = {}
            self.track_assign_subjects = {}
            self.tracks = list(tracks)
            
            for track in self.tracks:
                t_cols = self.track_valid_cols_map[track]
                track_df = df[df['科类'] == track].copy()
                
                for c in t_cols:
                    track_df[c] = pd.to_numeric(track_df[c], errors='coerce').fillna(0)
                    
                calc_cols, t_raw, t_assign = [], [], []
                
                for sub in t_cols:
                    rule = track_rules[track][sub]
                    if rule == "等级赋分":
                        assigned_col = f"{sub}赋分"
                        track_df[assigned_col] = self.assign_score_logic(track_df[sub])
                        calc_cols.append(assigned_col)
                        t_assign.append(sub)
                    else:
                        calc_cols.append(sub)
                        t_raw.append(sub)
                        
                self.track_calc_cols[track] = calc_cols
                self.track_raw_subjects[track] = t_raw
                self.track_assign_subjects[track] = t_assign
                
                track_df['3+1+2总分'] = track_df[calc_cols].sum(axis=1)
                track_df['科类统考排名'] = track_df['3+1+2总分'].rank(method='min', ascending=False).astype(int)
                track_df['班级内排名'] = track_df.groupby('班级')['3+1+2总分'].rank(method='min', ascending=False).astype(int)
                
                for col in calc_cols:
                    track_df['temp_sub'] = track_df[col].replace(0, np.nan)
                    track_df[f'{col}级排'] = track_df['temp_sub'].rank(method='min', ascending=False).fillna(9999).astype(int)
                    track_df[f'{col}班排'] = track_df.groupby('班级')['temp_sub'].rank(method='min', ascending=False).fillna(9999).astype(int)
                    track_df[f'{col}_pct'] = track_df['temp_sub'].rank(pct=True, ascending=True)

                def get_diagnostics(row):
                    pcts = {c: row[f'{c}_pct'] for c in calc_cols if pd.notna(row[f'{c}_pct']) and row[c] > 0}
                    if not pcts or len(pcts) < 3: return "无", "无"
                    best_sub = max(pcts, key=pcts.get).replace('赋分', '')
                    worst_sub = min(pcts, key=pcts.get).replace('赋分', '')
                    return best_sub, worst_sub

                track_df[['优势学科', '薄弱学科']] = track_df.apply(lambda r: pd.Series(get_diagnostics(r)), axis=1)
                track_df.drop(columns=[f'{col}_pct' for col in calc_cols] + ['temp_sub'], inplace=True, errors='ignore')
                
                processed_dfs.append(track_df)
                
            self.cleaned_df = pd.concat(processed_dfs).sort_values(['科类', '科类统考排名'])
            
            self.cb_report_track['values'] = self.tracks
            self.cb_chart_track['values'] = self.tracks
            if self.tracks:
                self.cb_report_track.current(0)
                self.cb_chart_track.current(0)
                self._on_chart_track_change()

            base_cols = ['班级', '姓名', '科类', '3+1+2总分', '班级内排名', '科类统考排名', '优势学科', '薄弱学科']
            all_display_cols = []
            
            for t in self.tracks:
                t_raw = self.track_raw_subjects.get(t, [])
                t_assign = self.track_assign_subjects.get(t, [])
                for sub in t_raw:
                    if sub not in all_display_cols: all_display_cols.extend([sub, f"{sub}班排", f"{sub}级排"])
                for sub in t_assign:
                    if sub not in all_display_cols: all_display_cols.extend([sub, f"{sub}赋分", f"{sub}赋分班排", f"{sub}赋分级排"])
            
            final_preview_cols = base_cols[:]
            seen = set(base_cols)
            for c in all_display_cols:
                if c not in seen and c in self.cleaned_df.columns:
                    final_preview_cols.append(c)
                    seen.add(c)
                    
            self.exist_cols = final_preview_cols 
            
            preview_df = self.cleaned_df[final_preview_cols].copy()
            for c in preview_df.columns:
                if c.endswith('班排') or c.endswith('级排'):
                    preview_df[c] = preview_df[c].replace(9999, '')
            self._update_treeview(self.tv_data, preview_df.head(50))
            
            self._generate_threshold_inputs()
            messagebox.showinfo("超级引擎完毕", "定制规则统算已完美落地！\n无用科目已剔除，0分未考者已剔除排名。前往后续页签体验高阶分析。")
        except Exception as e:
            messagebox.showerror("引擎异常", f"处理失败:\n{str(e)}")

    def _on_chart_track_change(self, event=None):
        track = self.chart_track_var.get()
        if track in self.track_calc_cols:
            metrics = ['3+1+2总分'] + self.track_calc_cols[track]
            self.cb_chart_metric['values'] = metrics
            self.cb_chart_metric.current(0)

    # ================= 双线KPI参数 =================

    def _generate_threshold_inputs(self):
        for widget in self.threshold_inputs_frame.winfo_children(): widget.destroy()
        self.threshold_entries = {}
        for row_idx, track in enumerate(self.tracks):
            ttk.Label(self.threshold_inputs_frame, text=f"[{track}] 特控:", font=("Microsoft YaHei", 10, "bold")).grid(row=row_idx, column=0, padx=5, pady=8)
            ent_tk = ttk.Entry(self.threshold_inputs_frame, width=6, justify=CENTER)
            ent_tk.insert(0, "500")
            ent_tk.grid(row=row_idx, column=1, padx=(0, 15))
            self.threshold_entries[f"{track}_特控"] = ent_tk
            
            ttk.Label(self.threshold_inputs_frame, text="本科:", font=("Microsoft YaHei", 10, "bold")).grid(row=row_idx, column=2, padx=5, pady=8)
            ent_bk = ttk.Entry(self.threshold_inputs_frame, width=6, justify=CENTER)
            ent_bk.insert(0, "430")
            ent_bk.grid(row=row_idx, column=3, padx=(0, 10))
            self.threshold_entries[f"{track}_本科"] = ent_bk

    def calculate_kpi(self):
        if self.cleaned_df.empty: return
        try:
            for key, ent in self.threshold_entries.items():
                self.thresholds[key] = float(ent.get())
            self.top_n_target = int(self.ent_top_n.get())
            self.margin_tekong = int(self.ent_margin_tekong.get())
            self.margin_benke = int(self.ent_margin_benke.get())
            
            bin_str = self.ent_score_bins.get()
            raw_bins = [int(x.strip()) for x in bin_str.split(',')]
            if 0 not in raw_bins: raw_bins.append(0)
            if 1500 not in raw_bins: raw_bins.append(1500) 
            self.score_bins_list = sorted(list(set(raw_bins)))
        except ValueError:
            messagebox.showerror("错误", "参数框内必须输入纯数字！")
            return

        df = self.cleaned_df.copy()
        
        def check_line(row, line_type):
            target = self.thresholds.get(f"{row['科类']}_{line_type}", 0)
            return 1 if row['3+1+2总分'] >= target else 0

        df['特控达线'] = df.apply(lambda r: check_line(r, '特控'), axis=1)
        df['本科达线'] = df.apply(lambda r: check_line(r, '本科'), axis=1)
        df['是否尖子生'] = df.apply(lambda row: 1 if row['科类统考排名'] <= self.top_n_target else 0, axis=1)
        
        stats = df.groupby(['科类', '班级']).agg(
            班级参考人数=('3+1+2总分', 'count'), 
            特控达线人数=('特控达线', 'sum'),
            本科达线人数=('本科达线', 'sum'),
            尖子生人数=('是否尖子生', 'sum')
        ).reset_index()
        
        stats['特控达线率'] = (stats['特控达线人数'] / stats['班级参考人数'] * 100).map('{:.1f}%'.format)
        stats['本科达线率'] = (stats['本科达线人数'] / stats['班级参考人数'] * 100).map('{:.1f}%'.format)
        
        stats = stats[['科类', '班级', '班级参考人数', '特控达线人数', '特控达线率', '本科达线人数', '本科达线率', '尖子生人数']]
        stats = stats.rename(columns={'尖子生人数': f'特优生(前{self.top_n_target})'})
        stats = stats.sort_values(by=['科类', '特控达线人数'], ascending=[True, False])
        
        self._update_treeview(self.tv_kpi, stats)

    # ================= 多维质量诊断 =================

    def generate_report(self):
        if self.cleaned_df.empty: return
        track = self.report_track_var.get()
        if not track: return

        try:
            self.report_text.delete(1.0, END)
            track_df = self.cleaned_df[self.cleaned_df['科类'] == track].copy()
            
            report = f"【{track}】多维质量诊断大表 (均分、标准差与最高分)\n"
            report += "="*140 + "\n"
            
            track_df['3+1+2总分'] = track_df['3+1+2总分'].astype(float)
            agg_dict = {'3+1+2总分': ['mean', 'std', 'max']}
            
            calc_cols = self.track_calc_cols.get(track, [])
            for sub in calc_cols:
                track_df[sub] = track_df[sub].astype(float)
                if track_df[sub].sum() > 0: agg_dict[sub] = ['mean', 'max']
                
            class_compare = track_df.groupby('班级').agg(agg_dict)
            class_compare.columns = ['_'.join(col).strip() for col in class_compare.columns.values]
            class_compare = class_compare.reset_index()
            
            rename_map = {'3+1+2总分_mean': '总分均分', '3+1+2总分_std': '总分标准差', '3+1+2总分_max': '总分极值'}
            for c in class_compare.columns:
                if c.endswith('_mean') and c not in rename_map: rename_map[c] = c.replace('_mean', '均分')
                if c.endswith('_max') and c not in rename_map: rename_map[c] = c.replace('_max', '最高')
            
            class_compare = class_compare.rename(columns=rename_map).sort_values(by='总分均分', ascending=False)
            for col in class_compare.columns:
                if col != '班级': class_compare[col] = class_compare[col].map('{:.2f}'.format)
                
            report += class_compare.to_string(index=False) + "\n\n"
            self.report_text.insert(END, report)
        except Exception as e: pass

    def draw_chart(self):
        if self.cleaned_df.empty: return
        track = self.chart_track_var.get()
        metric = self.chart_metric_var.get()
        chart_type = self.chart_type_var.get()
        if not track or not metric: return

        try:
            track_df = self.cleaned_df[self.cleaned_df['科类'] == track].copy()
            if metric not in track_df.columns: return
            track_df[metric] = track_df[metric].astype(float)
            if track_df[metric].sum() == 0: return

            self.ax.clear()

            if "柱状图" in chart_type:
                class_means = track_df.groupby('班级')[metric].mean().sort_values(ascending=False)
                bars = self.ax.bar(class_means.index.astype(str), class_means.values, color='#0078D7', alpha=0.85, width=0.6)
                self.ax.set_title(f"{track} - 各班级【{metric}】平均分", fontsize=14, pad=15, fontweight='bold', color='#333333')
                self.ax.set_ylabel("平均分", fontsize=11)
                self.ax.bar_label(bars, fmt='%.1f', padding=3)
                
            elif "直方图" in chart_type:
                scores = track_df[track_df[metric] > 0][metric] 
                self.ax.hist(scores, bins=15, color='#28A745', edgecolor='white', alpha=0.8)
                self.ax.set_title(f"{track} - 全年级【{metric}】分层分布直方图", fontsize=14, pad=15, fontweight='bold')
                self.ax.set_xlabel("分数区间", fontsize=11)
                self.ax.set_ylabel("人数", fontsize=11)

            self.ax.spines['top'].set_visible(False)
            self.ax.spines['right'].set_visible(False)
            self.figure.tight_layout()
            self.canvas.draw()
        except: pass

    # ================= 🚀 商业级 Excel 格式化与导出 =================
    
    def _format_excel_sheet(self, ws):
        """核心：为导出的 Excel Sheet 施加专业排版魔法（全居中、换行、全边框）"""
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                             top=Side(style='thin'), bottom=Side(style='thin'))
        center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        for row in ws.iter_rows():
            for cell in row:
                cell.border = thin_border
                cell.alignment = center_alignment
                
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter 
            for cell in col:
                try: 
                    val_str = str(cell.value)
                    visual_len = sum(2 if ord(c) > 127 else 1 for c in val_str)
                    if visual_len > max_length:
                        max_length = visual_len
                except: pass
            adjusted_width = max_length + 2
            if adjusted_width > 25: adjusted_width = 25
            elif adjusted_width < 10: adjusted_width = 10
            ws.column_dimensions[column].width = adjusted_width

    def export_all_reports(self):
        if self.cleaned_df.empty:
            messagebox.showwarning("提示", "请先完成数据导入！")
            return

        export_dir = filedialog.askdirectory(title="选择报表保存路径")
        if not export_dir: return
        
        try:
            self.export_btn.config(state=DISABLED)
            self.export_status.config(text="正在进行底层排版渲染与综合大表合并，请稍候...", foreground="orange")
            self.master.update()

            def clean_ranks(df_to_clean):
                for c in df_to_clean.columns:
                    if c.endswith('班排') or c.endswith('级排'):
                        df_to_clean[c] = df_to_clean[c].replace(9999, '')
                return df_to_clean

            base_cols = ['班级', '姓名', '科类', '3+1+2总分', '班级内排名', '科类统考排名', '优势学科', '薄弱学科']

            # ---------------- 任务 A：独立导出成绩单 ----------------
            class_dir = os.path.join(export_dir, "各班级独立成绩单_供分发")
            os.makedirs(class_dir, exist_ok=True)
            classes = self.cleaned_df['班级'].unique()
            
            for cls in classes:
                cls_df = self.cleaned_df[self.cleaned_df['班级'] == cls].sort_values('班级内排名')
                cls_tracks = cls_df['科类'].unique()
                if len(cls_tracks) > 0:
                    pt = cls_tracks[0]
                    t_raw = self.track_raw_subjects.get(pt, [])
                    t_assign = self.track_assign_subjects.get(pt, [])
                    
                    cls_export_cols = base_cols[:]
                    for sub in t_raw:
                        if sub in cls_df.columns: cls_export_cols.extend([sub, f"{sub}班排", f"{sub}级排"])
                    for sub in t_assign:
                        if f"{sub}赋分" in cls_df.columns: cls_export_cols.extend([sub, f"{sub}赋分", f"{sub}赋分班排", f"{sub}赋分级排"])
                    
                    class_data = cls_df[cls_export_cols]
                    class_data = clean_ranks(class_data)
                    
                    filepath = os.path.join(class_dir, f"高二_{cls}班_全维成绩单.xlsx")
                    with pd.ExcelWriter(filepath, engine='openpyxl') as w:
                        class_data.to_excel(w, index=False, sheet_name="成绩单")
                        self._format_excel_sheet(w.sheets["成绩单"])

            # ---------------- 任务 B：编译大一统教务表 ----------------
            master_file_path = os.path.join(export_dir, "【综合考务报告】年级统考全维数据矩阵.xlsx")
            
            with pd.ExcelWriter(master_file_path, engine='openpyxl') as writer:
                
                current_thresholds = {k: float(ent.get()) if ent.get().replace('.','',1).isdigit() else 0.0 for k, ent in self.threshold_entries.items()}
                df_kpi = self.cleaned_df.copy()
                
                def check_line(row, line_type):
                    target = current_thresholds.get(f"{row['科类']}_{line_type}", 0)
                    return 1 if row['3+1+2总分'] >= target else 0

                df_kpi['特控达线'] = df_kpi.apply(lambda r: check_line(r, '特控'), axis=1)
                df_kpi['本科达线'] = df_kpi.apply(lambda r: check_line(r, '本科'), axis=1)
                df_kpi['是否尖子生'] = df_kpi.apply(lambda row: 1 if row['科类统考排名'] <= self.top_n_target else 0, axis=1)
                
                stats = df_kpi.groupby(['科类', '班级']).agg(
                    班级人数=('3+1+2总分', 'count'), 
                    特控达线人数=('特控达线', 'sum'),
                    本科达线人数=('本科达线', 'sum'),
                    特优生人数=('是否尖子生', 'sum')
                ).reset_index()
                
                stats['特控达线率'] = (stats['特控达线人数'] / stats['班级人数'] * 100).map('{:.1f}%'.format)
                stats['本科达线率'] = (stats['本科达线人数'] / stats['班级人数'] * 100).map('{:.1f}%'.format)
                stats.rename(columns={'特优生人数': f'特优生(前{self.top_n_target})贡献'}, inplace=True)
                stats = stats[['科类', '班级', '班级人数', '特控达线人数', '特控达线率', '本科达线人数', '本科达线率', f'特优生(前{self.top_n_target})贡献']]
                stats = stats.sort_values(by=['科类', '特控达线人数'], ascending=[True, False])
                stats.to_excel(writer, sheet_name="大盘上线与考核", index=False)
                self._format_excel_sheet(writer.sheets["大盘上线与考核"])

                for track in self.tracks:
                    track_df = self.cleaned_df[self.cleaned_df['科类'] == track].copy()
                    calc_cols = self.track_calc_cols.get(track, [])
                    t_raw = self.track_raw_subjects.get(track, [])
                    t_assign = self.track_assign_subjects.get(track, [])
                    
                    valid_track_cols = base_cols[:]
                    for sub in t_raw: valid_track_cols.extend([sub, f"{sub}班排", f"{sub}级排"])
                    for sub in t_assign: valid_track_cols.extend([sub, f"{sub}赋分", f"{sub}赋分班排", f"{sub}赋分级排"])
                        
                    track_board = track_df[valid_track_cols].sort_values('科类统考排名')
                    track_board = clean_ranks(track_board)
                    track_board.to_excel(writer, sheet_name=f"{track}-全面总榜", index=False)
                    self._format_excel_sheet(writer.sheets[f"{track}-全面总榜"])

                    top_board = track_board[track_board['科类统考排名'] <= self.top_n_target]
                    top_board.to_excel(writer, sheet_name=f"{track}-Top{self.top_n_target}光荣榜", index=False)
                    self._format_excel_sheet(writer.sheets[f"{track}-Top{self.top_n_target}光荣榜"])
                    
                    for line_type, margin in [('特控', self.margin_tekong), ('本科', self.margin_benke)]:
                        target = current_thresholds.get(f"{track}_{line_type}", 0)
                        if target == 0: continue
                        
                        border_df = track_df[(track_df['3+1+2总分'] >= target - margin) & (track_df['3+1+2总分'] <= target + margin)].copy()
                        border_df[f'距{line_type}分差'] = border_df['3+1+2总分'] - target
                        border_df['薄弱学科(亟待提升)'] = border_df['薄弱学科']
                        
                        border_cols = ['班级', '姓名', '3+1+2总分', f'距{line_type}分差', '薄弱学科(亟待提升)', '优势学科', '科类统考排名']
                        other_cols = [c for c in valid_track_cols if c not in border_cols and c != '薄弱学科']
                        border_cols.extend(other_cols)
                        
                        border_df = border_df[border_cols].sort_values(['班级', f'距{line_type}分差'], ascending=[True, False])
                        border_df = clean_ranks(border_df)
                        border_df.to_excel(writer, sheet_name=f"{track}-{line_type}临界生", index=False)
                        self._format_excel_sheet(writer.sheets[f"{track}-{line_type}临界生"])

                    rate_dfs = []
                    for sub in calc_cols:
                        max_s = 150 if any(n in sub for n in ['语', '数', '外', '英']) else 100
                        sub_df = track_df[track_df[sub] > 0].groupby('班级')[sub].agg(
                            均分='mean',
                            及格率=lambda x, m=max_s: (x >= m*0.6).mean(),
                            优秀率=lambda x, m=max_s: (x >= m*0.8).mean()
                        ).reset_index()
                        sub_df = sub_df.rename(columns={'均分': f'{sub}均分', '及格率': f'{sub}及格率', '优秀率': f'{sub}优秀率'})
                        rate_dfs.append(sub_df.set_index('班级'))
                    
                    if rate_dfs:
                        final_rate_df = pd.concat(rate_dfs, axis=1).reset_index()
                        for col in final_rate_df.columns:
                            if '率' in col: final_rate_df[col] = (final_rate_df[col]*100).map('{:.1f}%'.format)
                            elif '均分' in col: final_rate_df[col] = final_rate_df[col].map('{:.2f}'.format)
                        final_rate_df.to_excel(writer, sheet_name=f"{track}-单科三率矩阵", index=False)
                        self._format_excel_sheet(writer.sheets[f"{track}-单科三率矩阵"])

                    agg_dict = {'3+1+2总分': ['mean', 'std', 'max']}
                    for sub in calc_cols:
                        track_df[sub] = track_df[sub].astype(float)
                        if track_df[sub].sum() > 0: agg_dict[sub] = ['mean', 'max']
                    class_compare = track_df.groupby('班级').agg(agg_dict)
                    class_compare.columns = ['_'.join(col).strip() for col in class_compare.columns.values]
                    class_compare = class_compare.reset_index()
                    rename_map = {'3+1+2总分_mean': '总分均分', '3+1+2总分_std': '总分标准差(离散)', '3+1+2总分_max': '班级最高分'}
                    for c in class_compare.columns:
                        if c.endswith('_mean') and c not in rename_map: rename_map[c] = c.replace('_mean', '均分')
                        if c.endswith('_max') and c not in rename_map: rename_map[c] = c.replace('_max', '最高分')
                    class_compare = class_compare.rename(columns=rename_map).sort_values(by='总分均分', ascending=False)
                    for col in class_compare.columns:
                        if col != '班级': class_compare[col] = class_compare[col].map('{:.2f}'.format)
                    class_compare.to_excel(writer, sheet_name=f"{track}-综合教学诊断", index=False)
                    self._format_excel_sheet(writer.sheets[f"{track}-综合教学诊断"])

                    bins = self.score_bins_list
                    labels = []
                    for i in range(len(bins)-1):
                        if i == len(bins) - 2: labels.append(f"{bins[i]}分及以上")
                        elif i == 0: labels.append(f"{bins[i+1]-1}分及以下")
                        else: labels.append(f"{bins[i]}-{bins[i+1]-1}分")
                    
                    track_df['分数段'] = pd.cut(track_df['3+1+2总分'], bins=bins, labels=labels, right=False)
                    band_stats = pd.crosstab(track_df['班级'], track_df['分数段'])
                    band_stats = band_stats[band_stats.columns[::-1]].reset_index()
                    band_stats.to_excel(writer, sheet_name=f"{track}-分数段分层矩阵", index=False)
                    self._format_excel_sheet(writer.sheets[f"{track}-分数段分层矩阵"])

            self.export_status.config(text=f"✅ 完美！全景排版级商业大表已生成至:\n{export_dir}", foreground="green")
            messagebox.showinfo("超级引擎导出完毕", "🎓 天花板级数据引擎统算及排版渲染完毕！\n所有Excel大表已实现：自动居中、自动适应列宽、自动换行及全框线包裹。请前往文件夹检阅您的作品！")
            
        except Exception as e:
            self.export_status.config(text="❌ 导出过程发生错误", foreground="red")
            messagebox.showerror("导出错误", f"文件导出失败，请确认 Excel 没有被占用。\n详细: {str(e)}")
        finally:
            self.export_btn.config(state=NORMAL)

    def _update_treeview(self, tree, df):
        tree.delete(*tree.get_children())
        tree["columns"] = list(df.columns)
        for col in df.columns:
            tree.heading(col, text=col)
            w = 85
            if '排' in col or len(str(col)) <= 2: w = 65
            elif '学科' in col: w = 100
            tree.column(col, width=w, anchor=CENTER)
        for index, row in df.iterrows():
            tree.insert("", "end", values=list(row))

# ================= 🚀 商业级跨平台防漂移授权模块 =================
import hashlib
import uuid
import platform
import subprocess

SECRET_SALT = "LiuSuanTong_Chem_2026_@TopSecret!" 

def get_stable_machine_code():
    system = platform.system()
    hw_id = ""
    try:
        if system == "Windows":
            hw_id = subprocess.check_output('wmic baseboard get serialnumber').decode().split('\n')[1].strip()
        elif system == "Darwin":
            hw_id = subprocess.check_output("ioreg -rd1 -c IOPlatformExpertDevice | grep -E '(UUID)'", shell=True).decode().split('"')[3]
        elif system == "Linux":
            with open('/etc/machine-id', 'r') as f: hw_id = f.read().strip()
    except: pass
        
    if not hw_id: hw_id = str(uuid.getnode())
    raw_str = hw_id + platform.machine()
    return hashlib.md5(raw_str.encode('utf-8')).hexdigest().upper()[:16]

def get_license_file_path():
    home_dir = os.path.expanduser('~')
    return os.path.join(home_dir, ".liusuantong_auth.key")

def verify_license(machine_code, input_key):
    expected_hash = hashlib.sha256((machine_code + SECRET_SALT).encode('utf-8')).hexdigest().upper()[:20]
    expected_key = "-".join([expected_hash[i:i+4] for i in range(0, 20, 4)])
    return input_key.strip() == expected_key

def check_local_auth():
    key_path = get_license_file_path()
    if os.path.exists(key_path):
        try:
            with open(key_path, 'r') as f:
                saved_key = f.read().strip()
                mc = get_stable_machine_code()
                if verify_license(mc, saved_key): return True
        except: pass
    return False

def show_activation_window(root):
    auth_win = tk.Toplevel(root)
    auth_win.title("软件未授权 - 硫酸铜的遐想")
    auth_win.geometry("550x380")
    auth_win.resizable(False, False)
    
    def on_close():
        root.destroy()
        
    auth_win.protocol("WM_DELETE_WINDOW", on_close) 
    
    mc = get_stable_machine_code()
    
    ttk.Label(auth_win, text="🔒 系统级商业授权", font=("Microsoft YaHei", 22, "bold"), bootstyle=PRIMARY).pack(pady=(30, 10))
    ttk.Label(auth_win, text="检测到当前设备尚未激活《甘肃新高考数据分析舱》", font=("Microsoft YaHei", 11)).pack(pady=5)
    
    mc_frame = ttk.Frame(auth_win, padding=15, bootstyle="secondary")
    mc_frame.pack(pady=15, fill=X, padx=40)
    
    ttk.Label(mc_frame, text="本机硬件特征码：", font=("Microsoft YaHei", 10, "bold")).pack(side=LEFT)
    mc_ent = ttk.Entry(mc_frame, width=25, font=("Consolas", 12, "bold"), bootstyle=INFO)
    mc_ent.insert(0, mc)
    mc_ent.configure(state="readonly")
    mc_ent.pack(side=LEFT, padx=10)
    
    ttk.Label(auth_win, text="👇 请联系开发者【俞老师】获取专权注册码：", font=("Microsoft YaHei", 10)).pack(pady=(10, 5))
    
    key_ent = ttk.Entry(auth_win, width=35, font=("Consolas", 13), justify=CENTER)
    key_ent.pack(pady=5)
    
    def on_activate():
        input_key = key_ent.get().strip()
        if verify_license(mc, input_key):
            try:
                with open(get_license_file_path(), 'w') as f: f.write(input_key)
                messagebox.showinfo("激活成功", "🎉 恭喜！设备数字签名绑定成功！\n\n欢迎使用【硫酸铜的遐想】专属教务引擎。", parent=auth_win)
                auth_win.destroy()
            except Exception as e:
                messagebox.showerror("写入失败", f"无法保存授权文件。\n错误: {str(e)}", parent=auth_win)
        else:
            messagebox.showerror("激活失败", "❌ 授权码无效！请确认输入无误，且为本机专属授权码。", parent=auth_win)
            
    ttk.Button(auth_win, text="🔑 立即验证并激活", bootstyle=SUCCESS, width=25, command=on_activate).pack(pady=20)
    return auth_win

if __name__ == "__main__":
    app = ttk.Window(themename="yeti") 
    app.withdraw() 
    
    if check_local_auth():
        ReverieOfCopperSulfateAnalyzer(app)
        app.deiconify()
        app.mainloop()
    else:
        auth_win = show_activation_window(app)
        app.wait_window(auth_win) 
        
        try:
            if app.winfo_exists() and check_local_auth():
                ReverieOfCopperSulfateAnalyzer(app)
                app.deiconify()
                app.mainloop()
        except tk.TclError:
            pass
