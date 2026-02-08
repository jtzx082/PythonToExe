import customtkinter as ctk
import pubchempy as pcp
from rdkit import Chem
from rdkit.Chem import AllChem
import webbrowser
import os
from deep_translator import GoogleTranslator
import threading
from tkinter import filedialog # 用于打开本地文件

# --- 配置区域 ---
APP_VERSION = "v2.0.0 (Crystal Ed.)"
DEV_NAME = "俞晋全"
DEV_ORG = "俞晋全高中化学名师工作室"
DEV_SCHOOL = "金塔县中学"
COPYRIGHT_YEAR = "2026"
# ----------------

# --- 内置高中常见晶体 CIF 数据 (简化教学演示) ---
CRYSTAL_PRESETS = {
    "氯化钠 (NaCl) - 离子晶体": """
    data_NaCl
    _cell_length_a 5.64
    _cell_length_b 5.64
    _cell_length_c 5.64
    _cell_angle_alpha 90
    _cell_angle_beta 90
    _cell_angle_gamma 90
    _symmetry_space_group_name_H-M 'F m -3 m'
    loop_
    _atom_site_label
    _atom_site_fract_x
    _atom_site_fract_y
    _atom_site_fract_z
    Na 0.00000 0.00000 0.00000
    Na 0.50000 0.50000 0.00000
    Na 0.50000 0.00000 0.50000
    Na 0.00000 0.50000 0.50000
    Cl 0.50000 0.50000 0.50000
    Cl 0.00000 0.00000 0.50000
    Cl 0.00000 0.50000 0.00000
    Cl 0.50000 0.00000 0.00000
    """,
    "氯化铯 (CsCl) - 离子晶体": """
    data_CsCl
    _cell_length_a 4.123
    _cell_length_b 4.123
    _cell_length_c 4.123
    _cell_angle_alpha 90
    _cell_angle_beta 90
    _cell_angle_gamma 90
    _symmetry_space_group_name_H-M 'P m -3 m'
    loop_
    _atom_site_label
    _atom_site_fract_x
    _atom_site_fract_y
    _atom_site_fract_z
    Cs 0.5 0.5 0.5
    Cl 0 0 0
    """,
    "金刚石 (C) - 共价晶体": """
    data_Diamond
    _cell_length_a 3.567
    _cell_length_b 3.567
    _cell_length_c 3.567
    _cell_angle_alpha 90
    _cell_angle_beta 90
    _cell_angle_gamma 90
    _symmetry_space_group_name_H-M 'F d -3 m'
    loop_
    _atom_site_label
    _atom_site_fract_x
    _atom_site_fract_y
    _atom_site_fract_z
    C 0.00000 0.00000 0.00000
    C 0.25000 0.25000 0.25000
    """,
    "铜 (Cu) - 面心立方堆积": """
    data_Cu
    _cell_length_a 3.615
    _cell_length_b 3.615
    _cell_length_c 3.615
    _cell_angle_alpha 90
    _cell_angle_beta 90
    _cell_angle_gamma 90
    _symmetry_space_group_name_H-M 'F m -3 m'
    loop_
    _atom_site_label
    _atom_site_fract_x
    _atom_site_fract_y
    _atom_site_fract_z
    Cu 0.00000 0.00000 0.00000
    """,
    "铁 (Fe) - 体心立方堆积": """
    data_Fe
    _cell_length_a 2.866
    _cell_length_b 2.866
    _cell_length_c 2.866
    _cell_angle_alpha 90
    _cell_angle_beta 90
    _cell_angle_gamma 90
    _symmetry_space_group_name_H-M 'I m -3 m'
    loop_
    _atom_site_label
    _atom_site_fract_x
    _atom_site_fract_y
    _atom_site_fract_z
    Fe 0 0 0
    """,
    "镁 (Mg) - 六方最密堆积": """
    data_Mg
    _cell_length_a 3.21
    _cell_length_b 3.21
    _cell_length_c 5.21
    _cell_angle_alpha 90
    _cell_angle_beta 90
    _cell_angle_gamma 120
    _symmetry_space_group_name_H-M 'P 63/m m c'
    loop_
    _atom_site_label
    _atom_site_fract_x
    _atom_site_fract_y
    _atom_site_fract_z
    Mg 0.333333 0.666667 0.250000
    Mg 0.666667 0.333333 0.750000
    """,
     "二氧化碳 (干冰) - 分子晶体": """
    data_CO2
    _cell_length_a 5.624
    _cell_length_b 5.624
    _cell_length_c 5.624
    _cell_angle_alpha 90
    _cell_angle_beta 90
    _cell_angle_gamma 90
    _symmetry_space_group_name_H-M 'P a -3'
    loop_
    _atom_site_label
    _atom_site_fract_x
    _atom_site_fract_y
    _atom_site_fract_z
    C 0.00000 0.00000 0.00000
    O 0.11540 0.11540 0.11540
    """
}

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class AboutWindow(ctk.CTkToplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title("关于软件")
        self.geometry("400x300")
        self.attributes("-topmost", True)
        self.label_title = ctk.CTkLabel(self, text="化学结构 3D 教学系统", font=("Microsoft YaHei UI", 18, "bold"))
        self.label_title.pack(pady=(20, 10))
        self.label_ver = ctk.CTkLabel(self, text=f"版本: {APP_VERSION}", font=("Arial", 12))
        self.label_ver.pack(pady=0)
        self.frame_line = ctk.CTkFrame(self, height=2, fg_color="gray")
        self.frame_line.pack(fill="x", padx=50, pady=15)
        info_text = f"开发者: {DEV_NAME}\n单位: {DEV_SCHOOL}\n{DEV_ORG}"
        self.label_dev = ctk.CTkLabel(self, text=info_text, font=("Microsoft YaHei UI", 14), justify="center")
        self.label_dev.pack(pady=10)
        credits_text = "技术栈: Python, RDKit, PubChemPy, 3Dmol.js\n自动构建: GitHub Actions"
        self.label_credits = ctk.CTkLabel(self, text=credits_text, font=("Arial", 10), text_color="gray")
        self.label_credits.pack(side="bottom", pady=20)

class MoleculeViewerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title(f"化学结构 3D 演示平台 - {DEV_NAME}作品")
        self.geometry("700x600")
        self.grid_columnconfigure(0, weight=1)
        self.toplevel_window = None

        # --- 顶部 ---
        self.frame_top = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_top.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        self.label_title = ctk.CTkLabel(self.frame_top, text="化学结构 3D 建模", font=("Microsoft YaHei UI", 24, "bold"))
        self.label_title.pack(side="left")
        self.btn_about = ctk.CTkButton(self.frame_top, text="关于 / About", width=80, height=24, 
                                       fg_color="transparent", border_width=1, 
                                       text_color=("gray10", "gray90"), command=self.open_about)
        self.btn_about.pack(side="right")

        # --- 主要功能区 (Tabs) ---
        self.tabview = ctk.CTkTabview(self, width=650, height=400)
        self.tabview.grid(row=1, column=0, padx=20, pady=10)
        
        self.tab_mol = self.tabview.add("有机分子检索")
        self.tab_cryst = self.tabview.add("晶体结构展示")

        # === Tab 1: 有机分子 ===
        self.setup_molecule_tab()

        # === Tab 2: 晶体结构 ===
        self.setup_crystal_tab()

        # --- 底部 ---
        self.status_label = ctk.CTkLabel(self, text="系统就绪", text_color="gray")
        self.status_label.grid(row=2, column=0, pady=5)
        self.label_footer = ctk.CTkLabel(self, text=f"© {COPYRIGHT_YEAR} {DEV_ORG}", font=("Microsoft YaHei UI", 10), text_color="gray50")
        self.label_footer.grid(row=3, column=0, pady=(0, 10))

    def setup_molecule_tab(self):
        t = self.tab_mol
        t.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(t, text="输入有机物名称、分子式或 SMILES", font=("Microsoft YaHei UI", 14)).grid(row=0, column=0, pady=20)
        
        self.entry_chem = ctk.CTkEntry(t, placeholder_text="例如: 苯酚, C2H5OH, Aspirin", width=400, height=40)
        self.entry_chem.grid(row=1, column=0, pady=10)
        self.entry_chem.bind("<Return>", lambda e: self.start_thread(self.generate_molecule))

        self.style_var = ctk.StringVar(value="stick")
        radio_frame = ctk.CTkFrame(t, fg_color="transparent")
        radio_frame.grid(row=2, column=0, pady=10)
        ctk.CTkRadioButton(radio_frame, text="球棍模型", variable=self.style_var, value="stick").pack(side="left", padx=10)
        ctk.CTkRadioButton(radio_frame, text="比例模型", variable=self.style_var, value="sphere").pack(side="left", padx=10)

        self.btn_gen_mol = ctk.CTkButton(t, text="生成分子模型", command=lambda: self.start_thread(self.generate_molecule), height=40)
        self.btn_gen_mol.grid(row=3, column=0, pady=20)

    def setup_crystal_tab(self):
        t = self.tab_cryst
        t.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(t, text="选择高中常见典型晶体", font=("Microsoft YaHei UI", 14)).grid(row=0, column=0, pady=15)

        # 预设下拉框
        self.crystal_combo = ctk.CTkComboBox(t, values=list(CRYSTAL_PRESETS.keys()), width=300, height=35)
        self.crystal_combo.set("氯化钠 (NaCl) - 离子晶体")
        self.crystal_combo.grid(row=1, column=0, pady=10)

        # 晶胞重复设置
        repeat_frame = ctk.CTkFrame(t, fg_color="transparent")
        repeat_frame.grid(row=2, column=0, pady=10)
        ctk.CTkLabel(repeat_frame, text="堆积范围 (Supercell):").pack(side="left", padx=5)
        self.repeat_val = ctk.CTkSegmentedButton(repeat_frame, values=["1x1x1", "2x2x2", "3x3x3"])
        self.repeat_val.set("2x2x2") # 默认 2x2x2 看起来更像晶体
        self.repeat_val.pack(side="left", padx=10)

        # 按钮区
        btn_frame = ctk.CTkFrame(t, fg_color="transparent")
        btn_frame.grid(row=3, column=0, pady=20)

        self.btn_gen_cryst = ctk.CTkButton(btn_frame, text="加载预设晶体", 
                                           command=lambda: self.start_thread(self.generate_preset_crystal), width=150)
        self.btn_gen_cryst.pack(side="left", padx=10)

        ctk.CTkLabel(t, text="或者", text_color="gray").grid(row=4, column=0)

        self.btn_load_cif = ctk.CTkButton(t, text="导入本地 .CIF 文件", fg_color="transparent", border_width=1,
                                          command=lambda: self.start_thread(self.load_local_cif), width=150)
        self.btn_load_cif.grid(row=5, column=0, pady=10)


    def open_about(self):
        if self.toplevel_window is None or not self.toplevel_window.winfo_exists():
            self.toplevel_window = AboutWindow(self)
        else:
            self.toplevel_window.focus()

    def start_thread(self, target_func):
        threading.Thread(target=target_func, daemon=True).start()

    # --- 逻辑：生成有机分子 ---
    def generate_molecule(self):
        user_input = self.entry_chem.get().strip()
        if not user_input:
            self.update_status("请输入有机物名称！", "red")
            return

        self.update_status(f"正在搜索 '{user_input}' ...", "#1F6AA5")
        self.btn_gen_mol.configure(state="disabled")

        try:
            # 翻译
            search_query = user_input
            if self.is_contains_chinese(user_input):
                try:
                    search_query = GoogleTranslator(source='auto', target='en').translate(user_input)
                except: pass 
            
            # 搜索
            compounds = pcp.get_compounds(search_query, 'name')
            if not compounds:
                compounds = pcp.get_compounds(search_query, 'formula')
            
            if not compounds:
                self.update_status("未找到该化合物。", "orange")
                self.btn_gen_mol.configure(state="normal")
                return

            # RDKit 处理
            target_compound = compounds[0]
            smiles = target_compound.canonical_smiles
            mol = Chem.MolFromSmiles(smiles)
            mol_with_h = Chem.AddHs(mol)
            AllChem.EmbedMolecule(mol_with_h, AllChem.ETKDG())
            mol_block = Chem.MolToMolBlock(mol_with_h)

            self.create_html_viewer(user_input, mol_block, "mol", self.style_var.get())
            self.update_status(f"成功展示: {user_input}", "green")

        except Exception as e:
            self.update_status(f"错误: {str(e)}", "red")
            print(e)
        finally:
            self.btn_gen_mol.configure(state="normal")

    # --- 逻辑：生成预设晶体 ---
    def generate_preset_crystal(self):
        selected = self.crystal_combo.get()
        cif_data = CRYSTAL_PRESETS.get(selected)
        if cif_data:
            self.update_status(f"正在加载晶体: {selected}", "#1F6AA5")
            repeat = self.repeat_val.get() # "2x2x2"
            self.create_html_viewer(selected, cif_data, "cif", "crystal", repeat)
            self.update_status(f"加载完成: {selected}", "green")

    # --- 逻辑：加载本地 CIF ---
    def load_local_cif(self):
        file_path = filedialog.askopenfilename(filetypes=[("Crystallographic Information File", "*.cif"), ("All Files", "*.*")])
        if file_path:
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    cif_data = f.read()
                name = os.path.basename(file_path)
                self.update_status(f"加载文件: {name}", "#1F6AA5")
                repeat = self.repeat_val.get()
                self.create_html_viewer(name, cif_data, "cif", "crystal", repeat)
                self.update_status(f"成功加载: {name}", "green")
            except Exception as e:
                self.update_status(f"文件读取失败: {str(e)}", "red")

    def update_status(self, text, color):
        self.status_label.configure(text=text, text_color=color)

    def is_contains_chinese(self, strs):
        for _char in strs:
            if '\u4e00' <= _char <= '\u9fa5': return True
        return False

    def create_html_viewer(self, title, data, format_type, style_mode, repeat="1x1x1"):
        # repeat string "2x2x2" -> replicate object {x:2, y:2, z:2}
        rep_x, rep_y, rep_z = map(int, repeat.split('x'))

        # JS 逻辑配置
        js_logic = ""
        
        if style_mode == "crystal":
            # 晶体渲染逻辑
            js_logic = f"""
                let config = {{ backgroundColor: '#f5f7fa' }};
                let viewer = $3Dmol.createViewer(element, config);
                let data = `{data}`;
                
                // 加载晶体模型
                let m = viewer.addModel(data, "{format_type}");
                
                // 设置晶体样式 (通常显示晶胞框)
                m.setStyle({{}}, {{sphere: {{scale: 0.3, colorscheme: 'Jmol'}}, stick: {{radius: 0.1, colorscheme: 'Jmol'}}}});
                
                // 核心：复制晶胞 (Supercell)
                viewer.replicateUnitCell({rep_x}, {rep_y}, {rep_z}, m);
                
                // 添加单位晶胞框线
                viewer.addUnitCell(m, {{box:{{color:'black'}}, split:10}});
            """
        else:
            # 有机分子渲染逻辑
            style_config = ""
            if style_mode == "stick":
                style_config = "viewer.setStyle({}, {stick: {radius: 0.14, colorscheme: 'Jmol'}, sphere: {scale: 0.23, colorscheme: 'Jmol'}});"
            else:
                style_config = "viewer.setStyle({}, {sphere: {colorscheme: 'Jmol'}});"
            
            js_logic = f"""
                let config = {{ backgroundColor: '#f5f7fa' }};
                let viewer = $3Dmol.createViewer(element, config);
                let data = `{data}`;
                viewer.addModel(data, "{format_type}");
                {style_config}
            """

        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <title>{title} - {DEV_NAME} 3D</title>
            <script src="https://3Dmol.org/build/3Dmol-min.js"></script>
            <style>
                body {{ margin: 0; padding: 0; overflow: hidden; font-family: "Microsoft YaHei"; }}
                #container {{ width: 100vw; height: 100vh; position: relative; }}
                #info {{ 
                    position: absolute; top: 20px; left: 20px; z-index: 10; 
                    background: rgba(255, 255, 255, 0.9); padding: 15px; border-radius: 8px; 
                    box-shadow: 0 4px 10px rgba(0,0,0,0.1); border-left: 5px solid #3B8ED0;
                }}
            </style>
        </head>
        <body>
            <div id="info">
                <h2 style="margin:0;color:#333">{title}</h2>
                <p style="margin:5px 0;color:#666">操作: 左键旋转 | 滚轮缩放 | 右键平移</p>
                <div style="font-size:12px;color:#999;margin-top:5px">模式: {style_mode} ({repeat if style_mode=='crystal' else '单分子'})</div>
                <div style="font-size:12px;color:#ccc;text-align:right;margin-top:10px">Design by {DEV_ORG}</div>
            </div>
            <div id="container"></div>
            <script>
                let element = document.getElementById('container');
                {js_logic}
                viewer.zoomTo();
                viewer.render();
            </script>
        </body>
        </html>
        """
        
        filename = "structure_view.html"
        with open(filename, "w", encoding="utf-8") as f:
            f.write(html_content)
        webbrowser.open('file://' + os.path.realpath(filename))

if __name__ == "__main__":
    app = MoleculeViewerApp()
    app.mainloop()
