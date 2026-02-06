import customtkinter as ctk
import pubchempy as pcp
from rdkit import Chem
from rdkit.Chem import AllChem
import webbrowser
import os
from deep_translator import GoogleTranslator
import threading

# 设置外观模式
ctk.set_appearance_mode("System")  # 默认跟随系统
ctk.set_default_color_theme("blue")

class MoleculeViewerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("有机分子结构 3D 展示工具 - 教学版")
        self.geometry("600x450")
        self.grid_columnconfigure(0, weight=1)

        # 标题
        self.label_title = ctk.CTkLabel(self, text="有机化学分子 3D 建模助手", font=("Microsoft YaHei UI", 24, "bold"))
        self.label_title.grid(row=0, column=0, padx=20, pady=(20, 10))

        # 说明
        self.label_desc = ctk.CTkLabel(self, text="输入中文名称、英文名称、分子式或 SMILES (例如: 乙醇, Benzene)", font=("Microsoft YaHei UI", 14))
        self.label_desc.grid(row=1, column=0, padx=20, pady=10)

        # 输入框
        self.entry_chem = ctk.CTkEntry(self, placeholder_text="在此输入有机物名称...", width=400, height=40)
        self.entry_chem.grid(row=2, column=0, padx=20, pady=10)
        self.entry_chem.bind("<Return>", self.start_generation_thread)

        # 样式选择
        self.style_var = ctk.StringVar(value="stick")
        self.radio_frame = ctk.CTkFrame(self)
        self.radio_frame.grid(row=3, column=0, pady=10)
        
        ctk.CTkRadioButton(self.radio_frame, text="球棍模型 (Ball & Stick)", variable=self.style_var, value="stick").pack(side="left", padx=10)
        ctk.CTkRadioButton(self.radio_frame, text="空间填充 (Sphere)", variable=self.style_var, value="sphere").pack(side="left", padx=10)

        # 按钮
        self.btn_generate = ctk.CTkButton(self, text="生成并查看 3D 模型", command=self.start_generation_thread, height=40, font=("Microsoft YaHei UI", 16, "bold"))
        self.btn_generate.grid(row=4, column=0, padx=20, pady=20)

        # 状态显示
        self.status_label = ctk.CTkLabel(self, text="就绪", text_color="gray")
        self.status_label.grid(row=5, column=0, pady=10)

    def start_generation_thread(self, event=None):
        # 使用线程防止界面卡顿
        threading.Thread(target=self.generate_model, daemon=True).start()

    def generate_model(self):
        user_input = self.entry_chem.get().strip()
        if not user_input:
            self.status_label.configure(text="请输入内容！", text_color="red")
            return

        self.status_label.configure(text=f"正在搜索 '{user_input}' ...", text_color="orange")
        self.btn_generate.configure(state="disabled")

        try:
            # 1. 尝试翻译中文到英文 (PubChem 对中文支持不佳)
            search_query = user_input
            if self.is_contains_chinese(user_input):
                try:
                    search_query = GoogleTranslator(source='auto', target='en').translate(user_input)
                    print(f"Translating: {user_input} -> {search_query}")
                except Exception as e:
                    print(f"Translation failed: {e}")
                    # 如果翻译失败，仍尝试用原词搜索
            
            # 2. 在 PubChem 中搜索
            compounds = pcp.get_compounds(search_query, 'name')
            if not compounds:
                # 尝试当作分子式搜索
                compounds = pcp.get_compounds(search_query, 'formula')
            
            if not compounds:
                self.status_label.configure(text="未找到该化合物，请检查拼写。", text_color="red")
                self.btn_generate.configure(state="normal")
                return

            target_compound = compounds[0]
            smiles = target_compound.canonical_smiles
            cid = target_compound.cid
            name = user_input # 使用用户输入的名称作为文件名

            # 3. RDKit 处理：生成 3D 坐标
            mol = Chem.MolFromSmiles(smiles)
            mol = Chem.AddHs(mol) # 关键：加上氢原子
            AllChem.EmbedMolecule(mol, AllChem.ETKDG()) # 计算 3D 构象
            # 生成 PDB 数据块
            pdb_block = Chem.MolToPDBBlock(mol)

            # 4. 生成 HTML 文件 (使用 3Dmol.js)
            self.create_html_viewer(name, pdb_block, self.style_var.get())
            
            self.status_label.configure(text=f"成功！已在浏览器中打开 {name}", text_color="green")

        except Exception as e:
            self.status_label.configure(text=f"发生错误: {str(e)}", text_color="red")
            print(e)
        finally:
            self.btn_generate.configure(state="normal")

    def is_contains_chinese(self, strs):
        for _char in strs:
            if '\u4e00' <= _char <= '\u9fa5':
                return True
        return False

    def create_html_viewer(self, title, pdb_data, style):
        # 根据选择设置样式
        style_json = ""
        if style == "stick":
            style_json = "stick: {radius: 0.15, colorscheme: 'Jmol'}, sphere: {scale: 0.25, colorscheme: 'Jmol'}" # 球棍模型
        else:
            style_json = "sphere: {colorscheme: 'Jmol'}" # 填充模型

        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <title>{title} - 3D Structure</title>
            <script src="https://3Dmol.org/build/3Dmol-min.js"></script>
            <style>
                body {{ margin: 0; padding: 0; overflow: hidden; background-color: #f0f2f5; font-family: sans-serif; }}
                #container {{ width: 100vw; height: 100vh; position: relative; }}
                #info {{ position: absolute; top: 10px; left: 10px; z-index: 10; background: rgba(255,255,255,0.8); padding: 10px; border-radius: 5px; }}
            </style>
        </head>
        <body>
            <div id="info">
                <h2>{title}</h2>
                <p>左键旋转 | 滚轮缩放 | 右键平移</p>
            </div>
            <div id="container" class="mol-container"></div>
            <script>
                let element = document.getElementById('container');
                let config = {{ backgroundColor: 'white' }};
                let viewer = $3Dmol.createViewer(element, config);
                let pdb = `{pdb_data}`;
                
                viewer.addModel(pdb, "pdb");
                viewer.setStyle({{}}, {{{style_json}}});
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
