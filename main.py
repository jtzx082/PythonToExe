import customtkinter as ctk
import webbrowser

# --- 配置区域 ---
APP_VERSION = "v2.0.0 (High School Ed.)"
DEV_NAME = "俞晋全"
DEV_ORG = "俞晋全高中化学名师工作室"
DEV_SCHOOL = "金塔县中学"

# --- 核心数据：高中常考元素精选库 ---
# 格式: [原子序数, 符号, 中文名, 英文名, 相对原子质量, 类别ID, 电子排布, 化合价, 详细用途/性质]
# 类别ID: 1=碱金属, 2=碱土金属, 3=过渡金属, 4=主族金属, 5=类金属, 6=非金属, 7=卤素, 8=稀有气体, 9=镧系, 10=锕系
ELEMENTS_DB = {
    1: ["H", "氢", "Hydrogen", "1.008", 6, "1s1", "+1, -1", 
        "【性质】密度最小的气体，具有可燃性和还原性。\n【用途】合成氨工业原料；清洁能源（氢燃料电池）；冶炼金属（还原剂）。"],
    2: ["He", "氦", "Helium", "4.0026", 8, "1s2", "0", 
        "【性质】化学性质极不活泼。\n【用途】保护气（焊接金属）；填充探空气球；液氦用于超低温环境（磁悬浮列车）。"],
    3: ["Li", "锂", "Lithium", "6.94", 1, "[He] 2s1", "+1", 
        "【性质】密度最小的金属，保存于石蜡油中。\n【用途】锂电池核心材料；原子反应堆导热剂；制轻质合金。"],
    4: ["Be", "铍", "Beryllium", "9.012", 2, "[He] 2s2", "+2", 
        "【用途】原子能工业（中子源）；航空航天合金材料。"],
    5: ["B", "硼", "Boron", "10.81", 5, "[He] 2s2 2p1", "+3", 
        "【用途】硼砂、硼酸；核工业屏蔽材料；半导体掺杂剂。"],
    6: ["C", "碳", "Carbon", "12.011", 6, "[He] 2s2 2p2", "+4, +2, -4", 
        "【同素异形体】金刚石、石墨、C60。\n【用途】冶炼金属；燃料；有机物骨架；碳纤维材料。"],
    7: ["N", "氮", "Nitrogen", "14.007", 6, "[He] 2s2 2p3", "+5, +3, -3", 
        "【性质】空气中含量最高(78%)，化学性质稳定。\n【用途】合成氨、硝酸；保护气；液氮冷冻剂。"],
    8: ["O", "氧", "Oxygen", "15.999", 6, "[He] 2s2 2p4", "-2", 
        "【用途】供给呼吸；支持燃烧；炼钢；医疗急救。"],
    9: ["F", "氟", "Fluorine", "18.998", 7, "[He] 2s2 2p5", "-1", 
        "【性质】氧化性最强的非金属单质。\n【用途】制冷剂(氟利昂)；特种塑料(特氟龙)；牙膏添加剂(NaF)。"],
    10: ["Ne", "氖", "Neon", "20.180", 8, "[He] 2s2 2p6", "0", 
        "【用途】霓虹灯（发红光）；激光技术。"],
    11: ["Na", "钠", "Sodium", "22.990", 1, "[Ne] 3s1", "+1", 
        "【性质】质软、银白色，焰色反应为黄色。保存于煤油中。\n【用途】制备过氧化钠；原子反应堆导热剂；高压钠灯（透雾性强）。"],
    12: ["Mg", "镁", "Magnesium", "24.305", 2, "[Ne] 3s2", "+2", 
        "【性质】燃烧发出耀眼白光。\n【用途】制造信号弹、照明弹；镁铝合金（航空材料）。"],
    13: ["Al", "铝", "Aluminium", "26.982", 4, "[Ne] 3s2 3p1", "+3", 
        "【性质】两性金属，既能与酸反应又能与强碱反应。地壳中含量最高的金属。\n【用途】铝合金门窗；导线；铝热反应炼铁。"],
    14: ["Si", "硅", "Silicon", "28.085", 5, "[Ne] 3s2 3p2", "+4, -4", 
        "【性质】亲氧元素，自然界主要以氧化物和硅酸盐形式存在。\n【用途】半导体芯片（晶体硅）；光伏电池；光导纤维（SiO2）。"],
    15: ["P", "磷", "Phosphorus", "30.974", 6, "[Ne] 3s2 3p3", "+5, -3", 
        "【同素异形体】白磷（剧毒）、红磷。\n【用途】制造火柴（红磷）；磷肥；农药。"],
    16: ["S", "硫", "Sulfur", "32.06", 6, "[Ne] 3s2 3p4", "+6, +4, -2", 
        "【性质】淡黄色固体，俗称硫磺。\n【用途】制造硫酸；硫化橡胶；黑火药；杀菌剂。"],
    17: ["Cl", "氯", "Chlorine", "35.45", 7, "[Ne] 3s2 3p5", "+7, +5, +1, -1", 
        "【性质】黄绿色有毒气体。\n【用途】制盐酸、漂白粉；自来水消毒；合成PVC塑料。"],
    18: ["Ar", "氩", "Argon", "39.948", 8, "[Ne] 3s2 3p6", "0", 
        "【用途】焊接保护气；灯泡填充气。"],
    19: ["K", "钾", "Potassium", "39.098", 1, "[Ar] 4s1", "+1", 
        "【性质】焰色反应为紫色（透过蓝色钴玻璃）。\n【用途】钾肥；制备钾盐。"],
    20: ["Ca", "钙", "Calcium", "40.078", 2, "[Ar] 4s2", "+2", 
        "【用途】建筑材料（石灰石）；炼钢脱氧剂。"],
    26: ["Fe", "铁", "Iron", "55.845", 3, "[Ar] 3d6 4s2", "+3, +2", 
        "【性质】工业部门的基础，黑色金属。\n【用途】炼钢；机械制造；人体血红蛋白核心元素。"],
    29: ["Cu", "铜", "Copper", "63.546", 3, "[Ar] 3d10 4s1", "+2, +1", 
        "【性质】紫红色金属，导电性仅次于银。\n【用途】电线电缆；铜合金（黄铜、青铜）；杀菌剂（波尔多液）。"],
    35: ["Br", "溴", "Bromine", "79.904", 7, "[Ar] 3d10 4s2 4p5", "+5, -1", 
        "【性质】唯一的液态非金属，深红棕色，易挥发。\n【用途】制药；阻燃剂；感光材料（AgBr）。"],
    53: ["I", "碘", "Iodine", "126.90", 7, "[Kr] 4d10 5s2 5p5", "+7, +5, -1", 
        "【性质】紫黑色固体，易升华，遇淀粉变蓝。\n【用途】加碘食盐（KIO3）；碘酒消毒；人工降雨（AgI）。"],
    # ... 您可以按照此格式继续补充 ...
}

# 颜色配置 (背景色, 悬停色) - 使用柔和的教学配色
COLORS = {
    1: ("#FF9AA2", "#E07A82"), # 碱金属(红)
    2: ("#FFB7B2", "#E09792"), # 碱土金属(浅红)
    3: ("#E2F0CB", "#C2D0AB"), # 过渡金属(浅黄/白)
    4: ("#B5EAD7", "#95CAB7"), # 主族金属(浅绿)
    5: ("#C7CEEA", "#A7AEC8"), # 类金属(浅紫)
    6: ("#E0F7FA", "#B0D7DA"), # 非金属(蓝)
    7: ("#FFDAC1", "#DFBAC1"), # 卤素(橙)
    8: ("#FFFFD8", "#DFDFB8"), # 稀有气体
    9: ("#FF9AA2", "#E07A82"), # 镧系
    10: ("#FFB7B2", "#E09792"), # 锕系
    "default": ("#EEEEEE", "#CCCCCC")
}

ctk.set_appearance_mode("Light") # 教学常用亮色背景更清晰
ctk.set_default_color_theme("blue")

class PeriodicTableApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"高中化学数字化元素周期表 - {DEV_NAME}")
        self.geometry("1280x800")
        
        # 布局: 左侧表格(3/4)，右侧详情(1/4)
        self.grid_columnconfigure(0, weight=4)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # === 左侧: 周期表绘制区 ===
        self.frame_table = ctk.CTkFrame(self, fg_color="white")
        self.frame_table.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        
        # 标题
        ctk.CTkLabel(self.frame_table, text="元素周期表 (The Periodic Table of Elements)", 
                     font=("Microsoft YaHei UI", 24, "bold"), text_color="#333").pack(pady=10)
        
        # 按钮网格容器
        self.grid_container = ctk.CTkFrame(self.frame_table, fg_color="white")
        self.grid_container.pack(expand=True, fill="both", padx=20, pady=10)

        # 绘制元素
        self.draw_elements()
        self.draw_legend()

        # === 右侧: 详情面板 ===
        self.frame_info = ctk.CTkFrame(self, fg_color="#F9F9F9", corner_radius=10)
        self.frame_info.grid(row=0, column=1, padx=(0,10), pady=10, sticky="nsew")
        self.setup_info_panel()
        
        # 默认显示氢
        self.show_details(1)

    def draw_elements(self):
        # 循环 1 到 118 号元素
        for i in range(1, 119):
            r, c = self.get_coordinates(i)
            
            # 获取数据，如果数据库没有，显示默认
            data = ELEMENTS_DB.get(i, [str(i), "", "...", "", 3, "", "", ""])
            symbol = data[1]
            cat_id = data[4]
            
            # 颜色
            bg_color, hover_color = COLORS.get(cat_id, COLORS["default"])
            
            # 按钮
            btn = ctk.CTkButton(self.grid_container, 
                                text=f"{i}\n{symbol}", 
                                width=40, height=45,
                                fg_color=bg_color, 
                                text_color="#333",
                                hover_color=hover_color,
                                font=("Arial", 10, "bold"),
                                command=lambda aid=i: self.show_details(aid))
            
            # 放置 (注意: 镧系锕系有额外的间距)
            pady_val = 1
            if r >= 8: pady_val = (15, 1) if r==8 else 1 # 镧系上面空一点
            
            btn.grid(row=r, column=c, padx=1, pady=pady_val, sticky="nsew")

    def get_coordinates(self, atomic_num):
        # --- 核心算法：教材版长式周期表坐标 ---
        # 1. 镧系 (57-71) -> 放到底部 Row 8
        if 57 <= atomic_num <= 71:
            return 8, (atomic_num - 57) + 3 # Col 3-17
        
        # 2. 锕系 (89-103) -> 放到底部 Row 9
        if 89 <= atomic_num <= 103:
            return 9, (atomic_num - 89) + 3 # Col 3-17
        
        # 3. 主表逻辑
        if atomic_num == 1: return 1, 1
        if atomic_num == 2: return 1, 18
        
        if 3 <= atomic_num <= 4: return 2, atomic_num - 2
        if 5 <= atomic_num <= 10: return 2, atomic_num + 8
        
        if 11 <= atomic_num <= 12: return 3, atomic_num - 10
        if 13 <= atomic_num <= 18: return 3, atomic_num + 2
        
        if 19 <= atomic_num <= 36: return 4, atomic_num - 18
        if 37 <= atomic_num <= 54: return 5, atomic_num - 36
        
        if 55 <= atomic_num <= 56: return 6, atomic_num - 54
        if 72 <= atomic_num <= 86: return 6, atomic_num - 68 # 跳过镧系
        
        if 87 <= atomic_num <= 88: return 7, atomic_num - 86
        if 104 <= atomic_num <= 118: return 7, atomic_num - 100 # 跳过锕系
        
        return 10, 1 # Fallback

    def draw_legend(self):
        # 底部图例
        legend_frame = ctk.CTkFrame(self.frame_table, fg_color="transparent")
        legend_frame.pack(pady=10)
        
        legends = [("碱金属", 1), ("碱土金属", 2), ("过渡金属", 3), ("非金属", 6), ("卤素", 7), ("稀有气体", 8)]
        for name, cid in legends:
            color = COLORS[cid][0]
            l = ctk.CTkLabel(legend_frame, text=f"■ {name}  ", text_color=color, font=("Microsoft YaHei UI", 12, "bold"))
            l.pack(side="left")

    def setup_info_panel(self):
        # 大大的符号
        self.lbl_big_symbol = ctk.CTkLabel(self.frame_info, text="H", font=("Times New Roman", 70, "bold"), text_color="#333")
        self.lbl_big_symbol.pack(pady=(30, 0))
        
        self.lbl_cn_name = ctk.CTkLabel(self.frame_info, text="氢", font=("Microsoft YaHei UI", 30), text_color="#555")
        self.lbl_cn_name.pack()
        
        self.lbl_en_name = ctk.CTkLabel(self.frame_info, text="Hydrogen", font=("Arial", 14), text_color="#888")
        self.lbl_en_name.pack(pady=(0, 20))

        # 属性列表
        self.info_grid = ctk.CTkFrame(self.frame_info, fg_color="transparent")
        self.info_grid.pack(fill="x", padx=20)
        
        self.val_anum = self.create_info_row("原子序数:")
        self.val_mass = self.create_info_row("相对原子质量:")
        self.val_conf = self.create_info_row("电子排布:")
        self.val_valc = self.create_info_row("常见化合价:")

        # 用途/详细介绍 (新增)
        ctk.CTkLabel(self.frame_info, text="性质与用途", font=("Microsoft YaHei UI", 14, "bold"), text_color="#333", anchor="w").pack(fill="x", padx=20, pady=(20, 5))
        
        self.txt_uses = ctk.CTkTextbox(self.frame_info, height=200, fg_color="white", text_color="#333", font=("Microsoft YaHei UI", 13))
        self.txt_uses.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
        # 链接按钮
        ctk.CTkButton(self.frame_info, text="查看网络百科", command=self.open_wiki).pack(pady=20)

    def create_info_row(self, label):
        f = ctk.CTkFrame(self.info_grid, fg_color="transparent")
        f.pack(fill="x", pady=2)
        ctk.CTkLabel(f, text=label, width=90, anchor="w", font=("Microsoft YaHei UI", 12, "bold"), text_color="#555").pack(side="left")
        v = ctk.CTkLabel(f, text="--", anchor="w", font=("Arial", 12), text_color="#333")
        v.pack(side="left", fill="x")
        return v

    def show_details(self, aid):
        self.curr_aid = aid
        # 获取数据 (如果不存在则显示默认)
        raw = ELEMENTS_DB.get(aid, [str(aid), "??", "未知", "Unknown", 0, 3, "未知", "?", "暂无数据，请在代码中补充..."])
        
        self.lbl_big_symbol.configure(text=raw[1])
        self.lbl_cn_name.configure(text=raw[2])
        self.lbl_en_name.configure(text=raw[3])
        
        self.val_anum.configure(text=str(raw[0]))
        self.val_mass.configure(text=str(raw[3]))
        self.val_conf.configure(text=raw[5])
        self.val_valc.configure(text=raw[6])
        
        # 更新文本框 (需先启用编辑再禁用)
        self.txt_uses.configure(state="normal")
        self.txt_uses.delete("0.0", "end")
        self.txt_uses.insert("0.0", raw[7])
        self.txt_uses.configure(state="disabled")

    def open_wiki(self):
        name = self.lbl_cn_name.cget("text")
        webbrowser.open(f"https://baike.baidu.com/item/{name}")

if __name__ == "__main__":
    app = PeriodicTableApp()
    app.mainloop()
