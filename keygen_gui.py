import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import hashlib

# 必须与 main.py 中的盐值绝对一致！
SECRET_SALT = "LiuSuanTong_Chem_2026_@TopSecret!"

class LicenseGenerator:
    def __init__(self, master):
        self.master = master
        self.master.title("硫酸铜的遐想 - 商业授权注册机")
        self.master.geometry("550x400")
        self.master.resizable(False, False)
        self.setup_ui()

    def setup_ui(self):
        container = ttk.Frame(self.master, padding=30)
        container.pack(fill=BOTH, expand=YES)

        ttk.Label(container, text="🔑 核心商业授权注册机", font=("Microsoft YaHei", 22, "bold"), bootstyle=SUCCESS).pack(pady=(0, 5))
        ttk.Label(container, text="开发者专属配套工具，请妥善保管，严禁外传", font=("Microsoft YaHei", 10), foreground="gray").pack(pady=(0, 25))

        input_frame = ttk.Frame(container)
        input_frame.pack(fill=X, pady=10)
        ttk.Label(input_frame, text="1. 请输入客户发来的【机器码】：", font=("Microsoft YaHei", 11, "bold")).pack(anchor=W)
        
        self.ent_machine_code = ttk.Entry(input_frame, font=("Consolas", 15), justify=CENTER)
        self.ent_machine_code.pack(fill=X, pady=10)

        ttk.Button(container, text="⚙️ 极速生成专属授权码", bootstyle=PRIMARY, width=30, command=self.generate_key).pack(pady=15)

        output_frame = ttk.Frame(container)
        output_frame.pack(fill=X, pady=10)
        ttk.Label(output_frame, text="2. 生成的 20 位专属授权码：", font=("Microsoft YaHei", 11, "bold")).pack(anchor=W)

        self.ent_license_key = ttk.Entry(output_frame, font=("Consolas", 15, "bold"), justify=CENTER, bootstyle=INFO)
        self.ent_license_key.pack(fill=X, pady=10)

        ttk.Button(container, text="📋 一键复制授权码", bootstyle=(SUCCESS, OUTLINE), width=20, command=self.copy_to_clipboard).pack(pady=5)

    def generate_key(self):
        mc = self.ent_machine_code.get().strip()
        if not mc:
            ttk.dialogs.dialogs.Messagebox.show_error("错误", "机器码不能为空！")
            return
        expected_hash = hashlib.sha256((mc + SECRET_SALT).encode('utf-8')).hexdigest().upper()[:20]
        expected_key = "-".join([expected_hash[i:i+4] for i in range(0, 20, 4)])

        self.ent_license_key.delete(0, tk.END)
        self.ent_license_key.insert(0, expected_key)

    def copy_to_clipboard(self):
        key = self.ent_license_key.get().strip()
        if key:
            self.master.clipboard_clear()
            self.master.clipboard_append(key)
            self.master.update() 
            ttk.dialogs.dialogs.Messagebox.show_info("复制成功", "授权码已复制到剪贴板，可直接在微信中粘贴发给客户！")

if __name__ == "__main__":
    # 采用炫酷的暗黑黑客主题，与主程序的亮色办公风彻底区分！
    app = ttk.Window(themename="darkly") 
    LicenseGenerator(app)
    app.mainloop()
