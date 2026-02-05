import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import os

# --------------------------
# 核心算法逻辑
# --------------------------

def get_grade_config():
    """定义赋分等级和区间标准 (甘肃/通用 3+1+2)"""
    return [
        {'grade': 'A', 'percent': 0.15, 't_max': 100, 't_min': 86},
        {'grade': 'B', 'percent': 0.35, 't_max': 85,  't_min': 71},
        {'grade': 'C', 'percent': 0.35, 't_max': 70,  't_min': 56},
        {'grade': 'D', 'percent': 0.13, 't_max': 55,  't_min': 41},
        {'grade': 'E', 'percent': 0.02, 't_max': 40,  't_min': 30},
    ]

def calculate_assigned_score(series):
    """
    对单科原始成绩进行赋分计算
    会自动忽略空值（即没选该科的学生），只对有分数的学生群体进行赋分
    """
    # 1. 强制转换为数字类型，非数字变为NaN
    series_numeric = pd.to_numeric(series, errors='coerce')
    
    # 2. 过滤掉没选该科的学生（NaN）
    valid_scores = series_numeric.dropna()
    
    total_count = len(valid_scores)
    
    # 如果没人选这科，返回全空
    if total_count == 0:
        return pd.Series(index=series.index, dtype=float)

    # 3. 排序：降序
    sorted_scores = valid_scores.sort_values(ascending=False)
    
    # 4. 划分等级并计算
    assigned_result = pd.Series(index=valid_scores.index, dtype=float)
    
    current_idx = 0
    configs = get_grade_config()
    
    for cfg in configs:
        count = int(np.round(total_count * cfg['percent']))
        if cfg['grade'] == 'E':
            count = total_count - current_idx
        
        if count <= 0:
            continue

        end_idx = min(current_idx + count, total_count)
        if current_idx >= end_idx:
            break

        grade_indices = sorted_scores.iloc[current_idx : end_idx].index
        grade_raw_scores = sorted_scores.iloc[current_idx : end_idx]
        
        Y2 = grade_raw_scores.max()
        Y1 = grade_raw_scores.min()
        T2 = cfg['t_max']
        T1 = cfg['t_min']
        
        def calculate_single(Y):
            if Y2 == Y1: 
                return (T2 + T1) / 2
            else:
                return T1 + ((Y - Y1) * (T2 - T1)) / (Y2 - Y1)

        assigned_vals = grade_raw_scores.apply(calculate_single)
        assigned_result.loc[grade_indices] = assigned_vals
        
        current_idx = end_idx

    return assigned_result.round()

def calculate_rankings(df, score_col, class_col, rank_col_name_grade, rank_col_name_class):
    """
    计算排名（年级排 + 班级排）
    score_col: 要排名的分数列名
    class_col: 班级列名
    """
    # 1. 年级排名 (method='min' 表示并列第1，下一个人第3)
    df[rank_col_name_grade] = df[score_col].rank(ascending=False, method='min')
    
    # 2. 班级排名 (分组计算)
    if class_col in df.columns:
        df[rank_col_name_class] = df.groupby(class_col)[score_col].rank(ascending=False, method='min')
    else:
        df[rank_col_name_class] = None # 如果没找到班级列，留空

# --------------------------
# GUI 与 业务流程
# --------------------------

def run_app():
    root = tk.Tk()
    root.withdraw()

    try:
        messagebox.showinfo("甘肃新高考赋分工具 Pro", 
                            "版本更新说明：\n"
                            "1. 支持选课组合差异化处理（没分数的科目自动跳过）。\n"
                            "2. 总分自动合成（原始分科目 + 赋分后科目）。\n"
                            "3. 自动生成单科及总分的【年级排名】和【班级排名】。")

        # 1. 选择文件
        file_path = filedialog.askopenfilename(title="选择学生成绩表 (Excel)", filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path: return

        df = pd.read_excel(file_path)

        # 2. 识别关键列名
        # 2.1 班级列
        class_col = simpledialog.askstring("列名配置", "请输入Excel中【班级】所在的列名：\n(例如：班级 / 班号 / Class)", initialvalue="班级")
        if not class_col or class_col not in df.columns:
            messagebox.showwarning("警告", f"未找到列名'{class_col}'，将无法计算班级排名，仅计算年级排名。")
            class_col = None

        # 2.2 原始分科目 (语数外 + 物理/历史)
        raw_subs_str = simpledialog.askstring("科目配置 1/2", 
            "请输入【直接计入总分】的原始分科目：\n(包括：语文、数学、英语、物理、历史)\n用逗号分隔", 
            initialvalue="语文,数学,英语,物理,历史")
        if not raw_subs_str: return
        raw_subs = [s.strip() for s in raw_subs_str.replace("，", ",").split(",") if s.strip() and s.strip() in df.columns]

        # 2.3 赋分科目 (化生政地)
        assign_subs_str = simpledialog.askstring("科目配置 2/2", 
            "请输入需【等级赋分】的科目：\n(通常为：化学、生物、政治、地理)\n用逗号分隔", 
            initialvalue="化学,生物,政治,地理")
        if not assign_subs_str: return
        assign_subs = [s.strip() for s in assign_subs_str.replace("，", ",").split(",") if s.strip() and s.strip() in df.columns]

        # 3. 开始处理
        output_df = df.copy()
        
        # 用于存储最终用于加总分的列名
        final_score_columns = []

        # --- A. 处理原始分科目 ---
        for sub in raw_subs:
            # 确保是数字，处理空值
            output_df[sub] = pd.to_numeric(output_df[sub], errors='coerce')
            final_score_columns.append(sub) # 原始分直接加入总分计算列表
            
            # 计算原始分排名
            calculate_rankings(output_df, sub, class_col, f"{sub}_年排", f"{sub}_班排")

        # --- B. 处理赋分科目 ---
        for sub in assign_subs:
            # 1. 计算赋分
            assigned_col_name = f"{sub}_赋分"
            try:
                # 赋分逻辑已经包含空值处理
                output_df[assigned_col_name] = calculate_assigned_score(df[sub])
            except Exception as e:
                messagebox.showerror("错误", f"计算 {sub} 失败: {e}")
                return
            
            # 2. 将赋分后的列加入总分计算列表
            # 注意：如果某学生没选这科，这里是NaN，sum的时候会自动忽略
            final_score_columns.append(assigned_col_name)

            # 3. 对“赋分后”的成绩进行排名 (用户要求：赋分科目只对赋分后成绩排名)
            calculate_rankings(output_df, assigned_col_name, class_col, f"{sub}_赋分_年排", f"{sub}_赋分_班排")

        # --- C. 计算总分 ---
        # 核心逻辑：将所有相关列（原始分列 + 赋分后列）相加
        # min_count=1 确保如果所有科目都是NaN（缺考），总分也是NaN，而不是0
        output_df["总分"] = output_df[final_score_columns].sum(axis=1, min_count=1)

        # --- D. 计算总分排名 ---
        calculate_rankings(output_df, "总分", class_col, "总分_年排", "总分_班排")
        
        # 排序：默认按总分年级排名排序
        output_df = output_df.sort_values("总分_年排")

        # 4. 整理列顺序 (可选：把排名列放对应的成绩后面，这里为了简单，直接全部输出)
        # 如果想让表格好看一点，可以简单重排一下列，这里保留所有列以防数据丢失

        # 5. 导出
        save_path = filedialog.asksaveasfilename(title="保存结果", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], initialfile="赋分排名结果.xlsx")
        
        if save_path:
            output_df.to_excel(save_path, index=False)
            messagebox.showinfo("成功", f"处理完成！\n已包含单科及总分的班排、年排。\n文件保存至: {save_path}")
            try:
                os.startfile(os.path.dirname(save_path))
            except:
                pass

    except Exception as e:
        import traceback
        messagebox.showerror("系统错误", f"发生意外错误:\n{str(e)}\n\n{traceback.format_exc()}")

if __name__ == "__main__":
    run_app()
