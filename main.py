import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import os
import sys

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
    :param series: Pandas Series (原始分数)
    :return: Pandas Series (赋分后分数)
    """
    # 过滤掉缺考或0分数据（视具体教务规定而定，这里假设参与排名的都有分）
    valid_scores = series.dropna()
    total_count = len(valid_scores)
    
    if total_count == 0:
        return series

    # 1. 排序：降序
    sorted_scores = valid_scores.sort_values(ascending=False)
    
    # 2. 划分等级并计算
    # 创建一个空的Series来存储结果，索引与原数据一致
    assigned_result = pd.Series(index=valid_scores.index, dtype=float)
    
    current_idx = 0
    configs = get_grade_config()
    
    for cfg in configs:
        # 计算该等级的人数
        count = int(np.round(total_count * cfg['percent']))
        
        # 修正人数误差（确保最后E等级包含剩余所有人）
        if cfg['grade'] == 'E':
            count = total_count - current_idx
        
        if count <= 0:
            continue

        # 获取该等级内的所有学生索引
        grade_indices = sorted_scores.iloc[current_idx : current_idx + count].index
        grade_raw_scores = sorted_scores.iloc[current_idx : current_idx + count]
        
        # 获取该等级原始分的 Max (Y2) 和 Min (Y1)
        Y2 = grade_raw_scores.max()
        Y1 = grade_raw_scores.min()
        T2 = cfg['t_max']
        T1 = cfg['t_min']
        
        # 对该等级内的每个学生进行线性插值赋分
        # 公式变换: T = T1 + ((Y - Y1) * (T2 - T1)) / (Y2 - Y1)
        
        def calculate_single(Y):
            if Y2 == Y1: # 如果该区间所有人都同分，取赋分区间中点
                return (T2 + T1) / 2
            else:
                return T1 + ((Y - Y1) * (T2 - T1)) / (Y2 - Y1)

        assigned_vals = grade_raw_scores.apply(calculate_single)
        assigned_result.loc[grade_indices] = assigned_vals
        
        current_idx += count

    # 四舍五入保留整数（高考赋分通常取整）
    return assigned_result.round().astype(int)

# --------------------------
# GUI 与 业务流程
# --------------------------

def run_app():
    root = tk.Tk()
    root.withdraw() # 隐藏主窗口

    messagebox.showinfo("甘肃新高考赋分工具", "欢迎使用！\n请准备好Excel文件，包含：姓名、考号、以及各科成绩。")

    # 1. 选择文件
    file_path = filedialog.askopenfilename(
        title="选择学生成绩表 (Excel)",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    
    if not file_path:
        return

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        messagebox.showerror("错误", f"无法读取文件: {e}")
        return

    # 2. 识别科目列
    # 简单的逻辑：让用户输入哪些列是赋分科目，哪些是直接计入科目
    # 假设前几列是基础信息，后面是成绩
    all_columns = df.columns.tolist()
    
    # 弹出简单对话框让用户输入列名（用逗号分隔）
    # 实际项目中可以用复选框列表，这里为了代码简洁使用输入框
    msg = f"检测到以下列名:\n{all_columns}\n\n请输入需要【赋分】的科目名称（用中文逗号或英文逗号分隔）:\n例如: 化学,生物"
    
    assigned_subs_str = simpledialog.askstring("选择赋分科目", msg)
    if not assigned_subs_str:
        return
        
    assigned_subs = [s.strip() for s in assigned_subs_str.replace("，", ",").split(",") if s.strip()]
    
    # 3. 处理数据
    output_df = df.copy()
    
    # 验证列是否存在
    for sub in assigned_subs:
        if sub not in df.columns:
            messagebox.showerror("错误", f"找不到列名: {sub}")
            return

    # 3.1 计算赋分
    for sub in assigned_subs:
        # 新增一列：科目_赋分
        new_col = f"{sub}_赋分"
        try:
            output_df[new_col] = calculate_assigned_score(df[sub])
        except Exception as e:
            messagebox.showerror("计算错误", f"计算科目 {sub} 时出错: {str(e)}")
            return

    # 3.2 计算总分
    # 询问哪些是原始计入科目 (语数外 + 物理/历史)
    raw_subs_str = simpledialog.askstring("选择原始计入科目", 
        f"请输入【直接计入总分】的原始科目 (语数外+首选科目):\n例如: 语文,数学,英语,物理")
    
    if raw_subs_str:
        raw_subs = [s.strip() for s in raw_subs_str.replace("，", ",").split(",") if s.strip()]
        
        # 开始计算总分：原始科目 + 赋分后的科目
        total_score_col = output_df[raw_subs].sum(axis=1) # 原始分部分
        
        for sub in assigned_subs:
            total_score_col += output_df[f"{sub}_赋分"] # 加上赋分部分
            
        output_df["总分"] = total_score_col
        
        # 3.3 排名
        output_df["排名"] = output_df["总分"].rank(ascending=False, method='min')
        output_df = output_df.sort_values("排名")

    # 4. 导出
    save_path = filedialog.asksaveasfilename(
        title="保存结果",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        initialfile="赋分结果.xlsx"
    )
    
    if save_path:
        output_df.to_excel(save_path, index=False)
        messagebox.showinfo("成功", f"处理完成！\n文件已保存至: {save_path}")
        
        # 尝试打开文件夹
        try:
            os.startfile(os.path.dirname(save_path))
        except:
            pass

if __name__ == "__main__":
    run_app()
