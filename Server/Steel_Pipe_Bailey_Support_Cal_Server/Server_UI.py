import tkinter as tk
from tkinter import ttk, messagebox
import threading
import subprocess
import sys
import os
import signal

server_process = None

def start_server():
    global server_process
    if server_process is not None:
        return

    server_dir = os.path.dirname(os.path.abspath(__file__))
    python_exec = sys.executable

    # 如果是打包后的环境，sys.executable 会指向 UI.exe
    # 这时我们改用用户本机的 python 解释器
    if python_exec.lower().endswith(".exe") and "server_ui" in python_exec.lower():
        # 自动找到系统 Python
        import shutil
        python_exec = shutil.which("python") or shutil.which("python3")
    
    print("使用的解释器：", python_exec)

    server_process = subprocess.Popen(
        [python_exec, "run_server.py"],
        cwd=server_dir,
        creationflags=subprocess.CREATE_NEW_CONSOLE
    )

    status_var.set("服务端运行中")
    status_label.configure(foreground="green")

def stop_server():
    global server_process
    if server_process is None:
        messagebox.showinfo("提示", "服务端没有运行。")
        return

    try:
        os.kill(server_process.pid, signal.SIGTERM)
    except Exception:
        pass
    server_process = None
    status_var.set("已停止")
    status_label.configure(foreground="red")

def on_close():
    stop_server()
    root.destroy()

root = tk.Tk()
root.title("数智桥施 服务端启动器")
root.geometry("320x180")
root.resizable(False, False)

status_var = tk.StringVar(value="已停止")

ttk.Label(root, text="服务端状态:", font=("微软雅黑", 10)).pack(pady=10)
status_label = ttk.Label(root, textvariable=status_var, font=("微软雅黑", 12, "bold"), foreground="red")
status_label.pack()

ttk.Button(root, text="启动服务端", command=start_server).pack(pady=10)
ttk.Button(root, text="停止服务端", command=stop_server).pack()

root.protocol("WM_DELETE_WINDOW", on_close)
root.mainloop()