import tkinter as tk
from tkinter import messagebox, filedialog
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import requests
from RSA import get_unique_id

# 退出/关闭/结束程序
def on_exit(root):
    if root.winfo_exists():  # 安全检查
        root.destroy()  # 正确关闭方式

# 验证注册码


# 服务端地址
server_url = "http://10.142.35.199:8000" 


def show_login_window(root, on_success_callback):
    def verify_license():
        email = email_var.get().strip()
        license_key = license_var.get().strip()
        device_id = get_unique_id()
        print(f"email: {email}, license_key: {license_key}, device_id: {device_id}")
        
        if not email or not license_key:
            messagebox.showwarning("输入错误", "请填写完整的邮箱和注册码")
            return

        try:
            response = requests.post(
                f"{server_url}/verify_license",json={
                "email": email, 
                "license_key": license_key,
                "device_id": device_id
                }
            )

            data = response.json()
            msg = data.get("msg")

            if response.status_code == 200 and data.get("status") == "success":
                messagebox.showinfo("验证成功", "注册码验证通过，可以使用软件！")
                login_win.destroy()       # 关闭验证窗口
                on_success_callback()     # 启动主 UI
            else:
                messagebox.showerror("验证失败", msg, parent=login_win)
        except Exception as e:
            messagebox.showerror("连接失败", f"无法连接服务器：{str(e)}")

    # 创建登录窗口toplevel
    login_win = ttk.Toplevel(root)
    login_win.title("用户登录验证")
    login_win.geometry("500x250")
    login_win.resizable(False, False)

    frame = ttk.Frame(login_win, padding=20)
    frame.pack(fill="both", expand=True)

    email_var = tk.StringVar()
    license_var = tk.StringVar()

    ttk.Label(frame, text="邮箱：").grid(row=0, column=0, pady=10, sticky="e")
    ttk.Entry(frame, textvariable=email_var, width=30).grid(row=0, column=1, pady=10)

    ttk.Label(frame, text="注册码：").grid(row=1, column=0, pady=10, sticky="e")
    ttk.Entry(frame, textvariable=license_var, width=30, show="*").grid(row=1, column=1, pady=10)

    ttk.Button(frame, text="登录验证", command=verify_license).grid(row=2, column=1, pady=20)
    login_win.protocol("WM_DELETE_WINDOW", lambda: (root.destroy(), sys.exit()))
    

