import customtkinter as ctk
import webbrowser
from tkinter import Canvas

# --- 配置区域 ---
APP_VERSION = "v1.0.0 (Periodic Table)"
DEV_NAME = "俞晋全"
DEV_ORG = "俞晋全高中化学名师工作室"
DEV_SCHOOL = "金塔县中学"
COPYRIGHT_YEAR = "2026"

# --- 核心数据：前 118 号元素完整教学数据 ---
# 格式: [原子序数, 符号, 中文名, 英文名, 相对原子质量, 类别, 电子排布(简化), 常见化合价]
# 类别色块映射: 
# 1=碱金属(红), 2=碱土金属(橙), 3=过渡金属(黄), 4=主族金属(灰), 
# 5=类金属(绿), 6=非金属(蓝), 7=卤素(深蓝), 8=稀有气体(紫), 9=镧系/锕系(粉)
ELEMENTS_DB = {
    1: ["H", "氢", "Hydrogen", "1.008", 6, "1s1", "+1, -1"],
    2: ["He", "氦", "Helium", "4.0026", 8, "1s2", "0"],
    3: ["Li", "锂", "Lithium", "6.94", 1, "[He] 2s1", "+1"],
    4: ["Be", "铍", "Beryllium", "9.0122", 2, "[He] 2s2", "+2"],
    5: ["B", "硼", "Boron", "10.81", 5, "[He] 2s2 2p1", "+3"],
    6: ["C", "碳", "Carbon", "12.011", 6, "[He] 2s2 2p2", "+4, -4, +2"],
    7: ["N", "氮", "Nitrogen", "14.007", 6, "[He] 2s2 2p3", "+5, +3, -3"],
    8: ["O", "氧", "Oxygen", "15.999", 6, "[He] 2s2 2p4", "-2"],
    9: ["F", "氟", "Fluorine", "18.998", 7, "[He] 2s2 2p5", "-1"],
    10: ["Ne", "氖", "Neon", "20.180", 8, "[He] 2s2 2p6", "0"],
    11: ["Na", "钠", "Sodium", "22.990", 1, "[Ne] 3s1", "+1"],
    12: ["Mg", "镁", "Magnesium", "24.305", 2, "[Ne] 3s2", "+2"],
    13: ["Al", "铝", "Aluminium", "26.982", 4, "[Ne] 3s2 3p1", "+3"],
    14: ["Si", "硅", "Silicon", "28.085", 5, "[Ne] 3s2 3p2", "+4, -4"],
    15: ["P", "磷", "Phosphorus", "30.974", 6, "[Ne] 3s2 3p3", "+5, +3, -3"],
    16: ["S", "硫", "Sulfur", "32.06", 6, "[Ne] 3s2 3p4", "+6, +4, -2"],
    17: ["Cl", "氯", "Chlorine", "35.45", 7, "[Ne] 3s2 3p5", "+7, +5, +1, -1"],
    18: ["Ar", "氩", "Argon", "39.948", 8, "[Ne] 3s2 3p6", "0"],
    19: ["K", "钾", "Potassium", "39.098", 1, "[Ar] 4s1", "+1"],
    20: ["Ca", "钙", "Calcium", "40.078", 2, "[Ar] 4s2", "+2"],
    21: ["Sc", "钪", "Scandium", "44.956", 3, "[Ar] 3d1 4s2", "+3"],
    22: ["Ti", "钛", "Titanium", "47.867", 3, "[Ar] 3d2 4s2", "+4, +3"],
    23: ["V", "钒", "Vanadium", "50.942", 3, "[Ar] 3d3 4s2", "+5, +4, +3"],
    24: ["Cr", "铬", "Chromium", "51.996", 3, "[Ar] 3d5 4s1", "+6, +3"],
    25: ["Mn", "锰", "Manganese", "54.938", 3, "[Ar] 3d5 4s2", "+7, +4, +2"],
    26: ["Fe", "铁", "Iron", "55.845", 3, "[Ar] 3d6 4s2", "+3, +2"],
    27: ["Co", "钴", "Cobalt", "58.933", 3, "[Ar] 3d7 4s2", "+3, +2"],
    28: ["Ni", "镍", "Nickel", "58.693", 3, "[Ar] 3d8 4s2", "+3, +2"],
    29: ["Cu", "铜", "Copper", "63.546", 3, "[Ar] 3d10 4s1", "+2, +1"],
    30: ["Zn", "锌", "Zinc", "65.38", 3, "[Ar] 3d10 4s2", "+2"],
    31: ["Ga", "镓", "Gallium", "69.723", 4, "[Ar] 3d10 4s2 4p1", "+3"],
    32: ["Ge", "锗", "Germanium", "72.630", 5, "[Ar] 3d10 4s2 4p2", "+4"],
    33: ["As", "砷", "Arsenic", "74.922", 5, "[Ar] 3d10 4s2 4p3", "+5, +3, -3"],
    34: ["Se", "硒", "Selenium", "78.971", 6, "[Ar] 3d10 4s2 4p4", "+6, +4, -2"],
    35: ["Br", "溴", "Bromine", "79.904", 7, "[Ar] 3d10 4s2 4p5", "+5, -1"],
    36: ["Kr", "氪", "Krypton", "83.798", 8, "[Ar] 3d10 4s2 4p6", "0, +2"],
    # ... (为了代码简洁，省略部分中间元素，实际使用时建议补全至118) ...
    47: ["Ag", "银", "Silver", "107.87", 3, "[Kr] 4d10 5s1", "+1"],
    50: ["Sn", "锡", "Tin", "118.71", 4, "[Kr] 4d10 5s2 5p2", "+4, +2"],
    53: ["I", "碘", "Iodine", "126.90", 7, "[Kr] 4d10 5s2 5p5", "+7, +5, -1"],
    54: ["Xe", "氙", "Xenon", "131.29", 8, "[Kr] 4d10 5s2 5p6", "0, +2, +4"],
    78: ["Pt", "铂", "Platinum", "195.08", 3, "[Xe] 4f14 5d9 6s1", "+4, +2"],
    79: ["Au", "金", "Gold", "196.97", 3, "[Xe] 4f14 5d10 6s1", "+3, +1"],
    80: ["Hg", "汞", "Mercury", "200.59", 3, "[Xe] 4f14 5d10 6s2", "+2, +1"],
    82: ["Pb", "铅", "Lead", "207.2", 4, "[Xe] 4f14 5d10 6s2 6p2", "+4, +2"],
}

# 补充一个简单的生成器，防止演示时空白（实际发布版本应填满）
def get_element_data(atomic_num):
    if atomic_num in ELEMENTS_DB:
        return ELEMENTS_DB[atomic_num]
    else:
        # 默认填充数据，防止报错
        return ["?", "未知", "Unknown", "(???)", 3, "n/a", "?"]

# 颜色配置 (亮色/暗色模式)
CATEGORY_COLORS = {
    1: ("#FF6B6B", "#8B0000"), # 碱金属
    2: ("#FFD93D", "#B8860B"), # 碱土金属
    3: ("#F7F7F7", "#505050"), # 过渡金属 (白/灰)
    4: ("#A0E7E5", "#2F4F4F"), # 主族金属
    5: ("#95E1D3", "#228B22"), # 类金属
    6: ("#74BDCB", "#1E90FF"), # 非金属
    7: ("#EFA8E4", "#4B0082"), # 卤素
    8: ("#B19CD9", "#483D8B"), # 稀有气体
    9: ("#FFC0CB", "#C71585"), # 镧系/锕系
}

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class ElementButton(ctk.CTkButton):
    def __init__(self, master, atomic_num, symbol, category_id, command_callback):
        # 获取颜色
        colors = CATEGORY_COLORS.get(category_id, ("gray", "gray"))
        
        super().__init__(master, 
                         text=f"{atomic_num}\n{symbol}", 
                         width=45, height=50, 
                         fg_color=colors[1], # 默认用深色模式的颜色，看起来更专业
                         hover_color=colors[0],
                         font=("Arial", 12, "bold"),
                         command=lambda: command_callback(atomic_num))

class PeriodicTableApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title(f"高中化学动态元素周期表 - {DEV_NAME}")
        self.geometry("1100x700")
        self.minsize(1000, 600)

        # 布局：左侧是表格(可滚动)，右侧是信息面板
        self.grid_columnconfigure(0, weight=3) # 左侧宽
        self.grid_columnconfigure(1, weight=1) # 右侧窄
        self.grid_rowconfigure(0, weight=1)

        # --- 左侧：周期表容器 ---
        self.table_frame = ctk.CTkFrame(self)
        self.table_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        
        # 提示标签
        self.lbl_hint = ctk.CTkLabel(self.table_frame, text="点击元素查看详情", text_color="gray")
        self.lbl_hint.pack(pady=5)

        # 真正的网格区域
        self.grid_area = ctk.CTkFrame(self.table_frame, fg_color="transparent")
        self.grid_area.pack(expand=True, fill="both", padx=10, pady=10)

        # --- 右侧：信息面板 ---
        self.info_frame = ctk.CTkFrame(self, corner_radius=15)
        self.info_frame.grid(row=0, column=1, padx=(0, 10), pady=10, sticky="nsew")
        
        # 信息面板内容初始化
        self.setup_info_panel()

        # --- 绘制周期表 ---
        self.draw_periodic_table()

    def setup_info_panel(self):
        # 标题区域
        self.lbl_symbol = ctk.CTkLabel(self.info_frame, text="H", font=("Times New Roman", 64, "bold"))
        self.lbl_symbol.pack(pady=(40, 0))
        
        self.lbl_name_cn = ctk.CTkLabel(self.info_frame, text="氢", font=("Microsoft YaHei UI", 32))
        self.lbl_name_cn.pack(pady=(0, 10))
        
        self.lbl_name_en = ctk.CTkLabel(self.info_frame, text="Hydrogen", font=("Arial", 16), text_color="gray")
        self.lbl_name_en.pack()

        # 分割线
        ctk.CTkFrame(self.info_frame, height=2, fg_color="gray").pack(fill="x", padx=20, pady=20)

        # 详细数据区
        self.detail_container = ctk.CTkFrame(self.info_frame, fg_color="transparent")
        self.detail_container.pack(fill="x", padx=20)

        # 数据行构建函数
        def create_row(parent, label, value_id):
            f = ctk.CTkFrame(parent, fg_color="transparent")
            f.pack(fill="x", pady=5)
            ctk.CTkLabel(f, text=label, width=80, anchor="w", font=("Microsoft YaHei UI", 12, "bold")).pack(side="left")
            l = ctk.CTkLabel(f, text="---", anchor="w", font=("Arial", 12))
            l.pack(side="left", fill="x", expand=True)
            return l

        self.val_atomic_num = create_row(self.detail_container, "原子序数:", "num")
        self.val_mass = create_row(self.detail_container, "相对原子质量:", "mass")
        self.val_config = create_row(self.detail_container, "电子排布:", "config")
        self.val_category = create_row(self.detail_container, "元素类别:", "cat")
        self.val_valence = create_row(self.detail_container, "常见化合价:", "val")

        # 底部按钮
        self.btn_wiki = ctk.CTkButton(self.info_frame, text="查看详细百科 (Baike)", 
                                      command=self.open_wiki, fg_color="transparent", border_width=1)
        self.btn_wiki.pack(side="bottom", pady=20)
        
        self.current_atomic_num = 1
        self.update_info_panel(1) # 默认显示氢

    def draw_periodic_table(self):
        # 周期表坐标映射 (Row, Col) - 这是一个手工校准的布局
        # 1-18列，1-7行
        positions = {
            1: (1, 1), 2: (1, 18), # Period 1
            3: (2, 1), 4: (2, 2), 5: (2, 13), 6: (2, 14), 7: (2, 15), 8: (2, 16), 9: (2, 17), 10: (2, 18), # Period 2
            11: (3, 1), 12: (3, 2), 13: (3, 13), 14: (3, 14), 15: (3, 15), 16: (3, 16), 17: (3, 17), 18: (3, 18), # Period 3
        }
        
        # 自动填充第4-7周期 (简化逻辑，也可以完全手写)
        # Period 4
        for i in range(19, 37):
            col = i - 18
            if i > 18: positions[i] = (4, i - 18)
        # Period 5
        for i in range(37, 55):
            positions[i] = (5, i - 36)
        # Period 6 (La系特殊处理)
        positions[55] = (6, 1); positions[56] = (6, 2)
        # 57-71 是镧系，按下不表或单独放底部
        for i in range(72, 87): positions[i] = (6, i - 72 + 4)
        # Period 7
        positions[87] = (7, 1); positions[88] = (7, 2)
        
        # 镧系锕系 (放在 Row 9, 10)
        lanthanides = range(57, 72)
        actinides = range(89, 104)
        
        for idx, atom_num in enumerate(lanthanides):
            positions[atom_num] = (9, 4 + idx) # 放在下方
        for idx, atom_num in enumerate(actinides):
            positions[atom_num] = (10, 4 + idx)

        # 开始绘制按钮
        # 获取所有已定义数据的元素，或者循环 1-118
        for atomic_num in range(1, 104): # 演示版暂只画到103
            if atomic_num in positions:
                r, c = positions[atomic_num]
                data = get_element_data(atomic_num)
                symbol = data[0]
                category = data[4]
                
                btn = ElementButton(self.grid_area, atomic_num, symbol, category, self.update_info_panel)
                btn.grid(row=r, column=c, padx=1, pady=1, sticky="nsew")

    def update_info_panel(self, atomic_num):
        self.current_atomic_num = atomic_num
        data = get_element_data(atomic_num)
        # Data结构: [符号, 中文, 英文, 质量, 类别ID, 排布, 化合价]
        
        self.lbl_symbol.configure(text=data[0])
        self.lbl_name_cn.configure(text=data[1])
        self.lbl_name_en.configure(text=data[2])
        
        self.val_atomic_num.configure(text=str(atomic_num))
        self.val_mass.configure(text=str(data[3]))
        self.val_config.configure(text=data[5])
        
        # 类别名称映射
        cat_names = {1:"碱金属", 2:"碱土金属", 3:"过渡金属", 4:"主族金属", 5:"类金属", 6:"非金属", 7:"卤素", 8:"稀有气体", 9:"镧系/锕系"}
        self.val_category.configure(text=cat_names.get(data[4], "其他"))
        self.val_valence.configure(text=data[6])
        
        # 动态改变边框颜色以匹配元素
        color = CATEGORY_COLORS.get(data[4], ("gray", "gray"))[1]
        self.info_frame.configure(border_color=color, border_width=2)

    def open_wiki(self):
        data = get_element_data(self.current_atomic_num)
        name = data[1] # 中文名
        url = f"https://baike.baidu.com/item/{name}"
        webbrowser.open(url)

if __name__ == "__main__":
    app = PeriodicTableApp()
    app.mainloop()
