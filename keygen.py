# keygen.py (硫酸铜的遐想 - 专属注册机)
import hashlib

# ⚠️ 核心机密：加密盐。任何人只要不知道这串字符，就绝对算不出正确的注册码！
SECRET_SALT = "LiuSuanTong_Chem_2026_@TopSecret!"

def generate_license_key(machine_code):
    """根据用户的机器码，生成20位授权码"""
    # 算法：将机器码与加密盐拼接，进行 SHA256 哈希计算，然后截取前 20 位
    raw_str = machine_code + SECRET_SALT
    license_key = hashlib.sha256(raw_str.encode('utf-8')).hexdigest().upper()[:20]
    # 格式化一下，变成 XXXX-XXXX-XXXX-XXXX-XXXX 的精美格式
    return "-".join([license_key[i:i+4] for i in range(0, 20, 4)])

if __name__ == "__main__":
    print("="*50)
    print(" 🌟 硫酸铜的遐想 - 软件授权注册机 🌟")
    print("="*50)
    user_mc = input("请输入用户发给您的【机器码】: ").strip()
    if user_mc:
        key = generate_license_key(user_mc)
        print("\n✅ 生成成功！请将以下【注册码】发送给该用户：")
        print(f"\n      {key}\n")
    print("="*50)
    input("按回车键退出...")