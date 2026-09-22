# 1. 标准库
import os
import sys
import winreg
import tkinter as tk
from contextlib import contextmanager
from tkinter import filedialog


# 同时屏蔽 Python 层 sys.stderr 与 OS 文件描述符 2(stderr)，
# 覆盖 Python print 与 C/C++ DLL(fprintf) 两类噪音输出。
# 用法: with suppress_dialog_stderr(): <原生对话框调用>
@contextmanager
def suppress_dialog_stderr():
    _saved_sys = sys.stderr
    _saved_fd = None
    _devnull_fd = None
    _devnull_sys = open(os.devnull, 'w')
    try:
        try:
            _saved_fd = os.dup(2)
            _devnull_fd = os.open(os.devnull, os.O_WRONLY)
            os.dup2(_devnull_fd, 2)
        except OSError:
            _saved_fd = _devnull_fd = None
        sys.stderr = _devnull_sys
        yield
    finally:
        sys.stderr = _saved_sys
        _devnull_sys.close()
        if _devnull_fd is not None:
            try:
                os.dup2(_saved_fd, 2)
            except OSError:
                pass
            os.close(_devnull_fd)
            os.close(_saved_fd)


# 通过查询注册表获取不同电脑上的桌面路径
def get_desktop_path_from_registry():
    try:
        # 打开注册表键
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
        )
        # 读取 Desktop 的值
        value, regtype = winreg.QueryValueEx(key, "Desktop")
        winreg.CloseKey(key)
        
        # 展开环境变量（如 %USERPROFILE%）
        return os.path.expandvars(value)
    except Exception as e:
        print(f"注册表读取失败: {e}")
        return os.path.join(os.environ['USERPROFILE'], 'Desktop')

    
# 选择文件
def qucik_open_file_dialog():
    # 获取文件完整路径
    root = tk.Tk()
    root.withdraw() # 隐藏主窗口
    # 原生对话框会加载第三方 DLL（Qt/AdSync 等）向 stderr 打印噪音提示，临时屏蔽
    with suppress_dialog_stderr():
        FileName = filedialog.askopenfilename() # 打开文件对话框,选择打开什么文件，返回文件名
    root.destroy()# 结束程序
    return FileName


def quick_open_folder_dialog():
    root = tk.Tk()
    root.withdraw()
    # 原生对话框会加载第三方 DLL 向 stderr 打印噪音提示，临时屏蔽
    with suppress_dialog_stderr():
        folder_path = filedialog.askdirectory(title="选择文件夹")
    root.destroy()
    return folder_path

# 保存文件
def qucik_save_file_dialog(extend):
    root = tk.Tk()
    root.withdraw() # 隐藏主窗口
    # 原生对话框会加载第三方 DLL 向 stderr 打印噪音提示，临时屏蔽
    with suppress_dialog_stderr():
        fileName = filedialog.asksaveasfilename(
            title="保存模型文件",
            initialdir=os.getcwd(),  # 初始目录为当前工作目录
            filetypes=[("", f"*{extend}")],# 设置可选择的文件类型，格式为 [(描述, 扩展名), (描述, 扩展名),...]
            defaultextension=extend #当用户未输入文件扩展名时，自动添加的默认扩展名
            ) # 打开文件对话框,让用户指定保存文件的路径和名称。
    root.destroy()# 结束程序
    return fileName


# 保存文件所在的文件夹路径, 不限制文件名
def quick_save_folder_dialog():
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    # 原生对话框会加载第三方 DLL 向 stderr 打印噪音提示，临时屏蔽
    with suppress_dialog_stderr():
        folder_path = filedialog.askdirectory(
            title="选择保存文件夹",
            initialdir=os.getcwd()  # 初始目录为当前工作目录
        )
    root.destroy()
    return folder_path


# 在指定路径下根据指定内容生成指定格式的文件
def write_file_within_path(savepath, txt, encoding_mode):
    try:
        with open(savepath, 'w', encoding = encoding_mode) as f:
            f.write(txt)
        print(f"文件已成功生成: {savepath}")
    except Exception as e:
        print(f"文件生成失败: {e}")


# 修改文件格式，路径不变
def file_extension_Modified(file_path, file_extension):
    # 获取目录路径和文件名
    path = os.path.dirname(file_path) # 文件目录路径
    name = os.path.basename(file_path) # 文件名
    # 拆分为文件名+扩展名
    root, ext = os.path.splitext(name)
    # 新的文件名
    new_name = root + file_extension
    new_path = os.path.join(path, new_name)
    return new_path
