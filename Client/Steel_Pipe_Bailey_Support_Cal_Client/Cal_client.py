import ttkbootstrap as ttk
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from ttkbootstrap.constants import *
import requests
import time
import os
import json
import threading

# 定义传输全流程
# 定义服务端口

server_url = "http://10.142.35.199:8000" 
task_in_progress = False

# 定义非阻塞式提示框
def show_info(title, message):
    win = tk.Toplevel()
    win.title(title)
    win.geometry("300x150")
    win.attributes("-topmost", True)  # 置顶窗口

    tk.Label(win, text=message, wraplength=280).pack(pady=20)
    tk.Button(win, text="确定", command=win.destroy).pack(pady=10)

# 检测服务端是否可访问
def check_server_alive(server_url):
    try:
        ping_url = f"{server_url}/ping"
        r = requests.get(ping_url, timeout=5)
        if r.status_code == 200 and r.json().get("status") == "ok":
            print("服务端响应正常")
            return True
        else:
            print("服务端响应异常")
            return False
    except requests.exceptions.RequestException as e:
        print(f"无法连接服务端：{e}")
        return False

# 检查上传数据完整性
def check_upload_data(mcb_path, txt_path, foundation_name, foundation_standard, foundation_values):
    if not mcb_path or not os.path.exists(mcb_path):
        show_info("上传失败", "未选择或找不到 MCB 文件")
        return False
    if not txt_path or not os.path.exists(txt_path):
        show_info("上传失败", "未选择或找不到参数 TXT 文件")
        return False
    if not foundation_name.strip():
        show_info("上传失败", "基础形式未选择")
        return False
    if not foundation_standard.strip():
        show_info("上传失败", "基础规范未选择")
        return False
    if not isinstance(foundation_values, dict) or not any(foundation_values.values()):
        show_info("上传失败", "基础参数输入错误")
        return False
    return True

# 上传mbc文件及基础数据
def submit_task(mcb_path, txt_path, foundation_name, foundation_standard, foundation_values):
    global task_in_progress
    if task_in_progress:
        show_info("提示", "当前有任务正在进行，请稍后再试")
        return None

    if not check_upload_data(mcb_path, txt_path, foundation_name, foundation_standard, foundation_values):
        return None  
    try:
        task_in_progress = True  # 设置任务状态为运行中
        with open(mcb_path, "rb") as mcb_file, open(txt_path, "rb") as txt_file:
            files = {
                "mcb_file": (os.path.basename(mcb_path), mcb_file, "application/octet-stream"),
                "txt_file": (os.path.basename(txt_path), txt_file, "text/plain")
                }
            data = {
                "foundation_name": foundation_name,
                "foundation_standard": foundation_standard,
                "foundation_values": json.dumps(foundation_values)  
            }
            response = requests.post(f"{server_url}/submit_task", files=files, data=data)
            
            if response.status_code == 200:
                task_id = response.json()["task_id"]
                show_info("上传成功",f"已提交任务，任务号：{task_id}")
                # messagebox.showinfo("上传成功",f"已提交任务，任务号：{task_id}")
                return task_id
            else:
                print("上传失败", f"{response.status_code}, {response.text}")
    except Exception as e:
        print("上传错误：", str(e))

# 定义状态获取函数
def poll_task_status(task_id):
    def _poll():
        global task_in_progress
        last_status = None
        while True:
            try:
                print(f"正在检查任务 {task_id} 的状态...")  # 调试信息：检查状态
                response = requests.post(f"{server_url}/check_status", data={"task_id": task_id})
                response.raise_for_status()  # 确保请求成功
                data = response.json()
                status = data["status"]
                msg = data["msg"]
                print(f"当前状态：{data['status']} - {data['msg']}")

                # 如果状态有变化或是第一次提示，则弹窗提示
                if status != last_status:
                    last_status = status
                    if status == "done":
                        show_info("完成", "计算完成，可以下载计算书了")
                        # messagebox.showinfo("完成", "计算完成，可以下载计算书了")
                        task_in_progress = False
                        break
                    elif status == "error":
                        show_info("错误", f"计算失败：{msg}")
                        # messagebox.showerror("错误", f"计算失败：{msg}")
                        task_in_progress = False
                        break
                    elif status == "running":
                        show_info("处理中", "已开始计算，请耐心等待")
                        # messagebox.showinfo("处理中", "已开始计算，请耐心等待")
                    elif status == "queued":
                        show_info("排队中", msg)
                        # messagebox.showinfo("排队中", msg)
                    else:
                        print(f"未知状态：{status}")
            except Exception as e:
                print(f"状态检查失败: {e}")
                task_in_progress = False
                break

            time.sleep(5)

    threading.Thread(target=_poll, daemon=True).start()  

# 启动运算进程
def run_task(task_id):
    if not task_id:
        show_info("错误", "请先上传模型文件")
        # messagebox.showerror("错误", "请先上传模型文件")
        return

    try:
        response = requests.post(f"{server_url}/run_analysis",  data={"task_id": task_id})
        response.raise_for_status()
        data = response.json()
        msg = data.get("msg", "")
        print("服务器返回消息", msg)

        if "无需重复计算" in msg:
            show_info("提示", "计算已完成，无需重复计算")
            return
        elif "计算失败" in msg:
            show_info("错误", f"计算失败，请重新上传：{msg}")
        elif "任务当前状态为" in msg:
            show_info("提示", msg)
            return
        elif "任务已排队执行" in msg:
            poll_task_status(task_id)
        else:
            show_info("提示", msg)  

    except Exception as e:
        print("计算启动失败", str(e))
        show_info("错误", f"计算启动失败：{e}")

# 下载计算书文件
def download_docx(task_id):
    if not task_id:
        show_info("错误", "请先上传模型文件")
        # messagebox.showerror("错误", "请先上传模型文件")
        return   
    try:
        response = requests.post(f"{server_url}/check_status", data={"task_id": task_id})
        response.raise_for_status()
        data = response.json()

        if data["status"] == "done":
            # 状态为 done，允许下载
            # 弹出保存路径选择窗口
            save_path = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word 文档", "*.docx")],
                title="保存计算书为..."
            )
            if not save_path:
                show_info("取消保存", "用户取消了保存操作")
                # messagebox.showinfo("取消保存", "用户取消了保存操作")
                return

            # 发起下载请求
            download_response = requests.post(f"{server_url}/download_docx", data={"task_id": task_id})
            download_response.raise_for_status()

            # 保存文件
            with open(save_path, "wb") as f:
                f.write(download_response.content)
            show_info("下载完成", f"计算书已保存至：\n{save_path}")
            # messagebox.showinfo("下载完成", f"计算书已保存至：\n{save_path}")
    
        elif data["status"] == "error":
            show_info("错误", f"计算失败：{data['msg']}")
            # messagebox.showerror("错误", f"计算失败：{data['msg']}")

        else:
            # 状态不是 done，禁止下载
            show_info("错误", "计算未完成，请稍后再试")
            # messagebox.showwarning("计算未完成", "计算尚未完成，请稍后再试")

    except Exception as e:
        show_info("下载失败", str(e))
        # messagebox.showerror("下载失败", str(e))

# 返回删除临时目录信息
# def delete_temp_dir(task_id):
#===========================================================================================================================================

