# 1. 标准库
import os
import sys
import winreg
import tkinter as tk
from pathlib import Path
from tkinter import messagebox

# 2. 第三方库
import ttkbootstrap as tb
from ttkbootstrap.constants import *

# 3. 本地模块
sys.path.append(str(Path(__file__).parent.parent.parent))
from General.UIHandle       import if_Reg
from General.ProgramMonitor import Program_State_Monitoring
from General.Midas          import MidasURL, Open_Midas_civil
from General.FilePath       import quick_open_folder_dialog, qucik_open_file_dialog

sys.path.append(str(Path(__file__).parent))
from CalRpt_MSWord.Steel_Sheet_Pile_CofferDam_Excel_Batch.CofferDam_Excel_Generate_Batch import make_all_doc_batch

# ======================================================================================
# UI 组件（与 Steel_Sheet_Pile_CofferDam_Cal_GUI 一致）
# ======================================================================================

def on_exit(root):
    if root.winfo_exists():
        root.destroy()


def dialog_headr(root, txt):
    style = tb.Style()
    style.configure("SimSun.TLabel", font=("SimSun", 10, "bold"))
    container = tb.Frame(root)
    container.pack(fill=X, expand=YES)
    hdr = tb.Label(master=root, text=txt, width=50, style="SimSun.TLabel")
    hdr.pack(fill=X, pady=10)


def dialog_boxline3(root, label_txt, label_length, entry_var, entry_length, button_txt, button_command):
    """标签 + 编辑框(disabled) + 按钮 —— 与原 GUI 一致"""
    container = tb.Frame(root)
    container.pack(fill=X, expand=YES, pady=5)
    label = tb.Label(master=container, text=label_txt.title(), width=label_length)
    label.pack(side=LEFT, padx=0)
    ent = tb.Entry(master=container, textvariable=entry_var, width=entry_length, state="disabled")
    ent.pack(side=LEFT, padx=20, fill=X, expand=YES)
    but = tb.Button(master=container, text=button_txt, command=button_command, width=15)
    but.pack(side=RIGHT, padx=10)


def dialog_boxline_end(root, button_txt1, button_command1, button_txt2, button_command2, button_txt3=None, button_command3=None):
    container = tb.Frame(root)
    container.pack(fill=X, expand=YES, pady=10)
    but = tb.Button(master=container, text=button_txt2, command=button_command2, width=6)
    but.pack(side=RIGHT, padx=20, pady=10)
    sub_btn = tb.Button(master=container, text=button_txt1, command=button_command1, bootstyle=SUCCESS, width=6)
    sub_btn.pack(side=RIGHT, pady=10)
    # 可选的第三个按钮, 位于确认按钮左边
    if button_txt3:
        cfg_btn = tb.Button(master=container, text=button_txt3, command=button_command3, width=10)
        cfg_btn.pack(side=RIGHT, padx=20, pady=10)


class CalConfigDialog:
    """计算配置 — 选择需要生成的计算书模块, 回到主流程时只传结果 dict"""
    def __init__(self, parent, init_config=None):
        win = tb.Toplevel(parent)
        win.title("计算配置")
        win.geometry(f"300x340+{parent.winfo_rootx()+50}+{parent.winfo_rooty()+50}")
        win.resizable(False, False)
        win.transient(parent)
        win.grab_set()
        self.win = win
        self.result = None

        init = init_config if init_config else {}

        # ===== 状态变量 =====
        self.ckb_pressure   = tk.BooleanVar(value=init.get("基坑内外土压力计算", True))    # 基坑内外土压力计算（固定）
        self.ckb_k_coef     = tk.BooleanVar(value=init.get("土层水平反力系数计算", True))  # 土层水平反力系数计算（固定）
        self.ckb_read_model = tk.BooleanVar(value=init.get("模型结果读取", True))          # 模型结果读取（固定）
        self.ckb_sliding    = tk.BooleanVar(value=init.get("滑动稳定性分析", True))        # 滑动稳定性分析
        self.ckb_embedment  = tk.BooleanVar(value=init.get("嵌固稳定性分析", True))        # 嵌固稳定性分析
        self.ckb_heave      = tk.BooleanVar(value=init.get("抗隆起稳定性分析", True))      # 抗隆起稳定性分析

        # ===== UI =====
        tb.Label(win, text="请选择需要生成的计算书模块").pack(anchor="w", padx=20, pady=(15, 10))

        # 虚线分隔：标题 | 勾选框列表（样式同主UI）
        _sep = tk.Canvas(win, height=2, highlightthickness=0, bg=win.cget("background"))
        _sep.pack(fill=X, padx=20, pady=(0, 10))
        _sep.create_line(0, 1, 2000, 1, dash=(4, 4), fill="gray")

        rows = [
            ("基坑内外土压力计算",   self.ckb_pressure,   True),
            ("土层水平反力系数计算", self.ckb_k_coef,     True),
            ("模型结果读取",         self.ckb_read_model, True),
            ("滑动稳定性分析",       self.ckb_sliding,    False),
            ("嵌固稳定性分析",       self.ckb_embedment,  False),
            ("抗隆起稳定性分析",     self.ckb_heave,      False),
        ]
        for text, var, fixed in rows:
            ckb = tb.Checkbutton(win, text=text, variable=var)
            if fixed:
                ckb.configure(state="disabled")
            ckb.pack(anchor="w", padx=25, pady=5)

        # ===== 底部按钮（居中, 确认在取消左边） =====
        bottom = tb.Frame(win)
        bottom.pack(pady=15)
        tb.Button(bottom, text="确认", bootstyle=SUCCESS, width=8, command=self._on_confirm).pack(side="left", padx=5)
        tb.Button(bottom, text="取消", bootstyle=SECONDARY, width=8, command=self._on_cancel).pack(side="left", padx=5)

    def _on_confirm(self):
        self.result = {
            "基坑内外土压力计算":   self.ckb_pressure.get(),
            "土层水平反力系数计算": self.ckb_k_coef.get(),
            "模型结果读取":         self.ckb_read_model.get(),
            "滑动稳定性分析":       self.ckb_sliding.get(),
            "嵌固稳定性分析":       self.ckb_embedment.get(),
            "抗隆起稳定性分析":     self.ckb_heave.get(),
        }
        self.win.destroy()

    def _on_cancel(self):
        self.result = None
        self.win.destroy()

    def show(self):
        self.win.wait_window()
        return self.result


# ======================================================================================
# 文件/文件夹选择
# ======================================================================================

def choose_excel_file(entry_var):
    openfile_path = qucik_open_file_dialog().replace("/", "\\")
    if "xlsx" not in openfile_path:
        messagebox.showinfo("提示", "未选择有效 xlsx 文件，请重新选择。")
        entry_var.set("未选取任何文件")
    else:
        entry_var.set(openfile_path)


def choose_model_folder(entry_var1, entry_var2):
    folder_path = quick_open_folder_dialog().replace("/", "\\")
    if folder_path:
        entry_var1.set(folder_path)
        entry_var2.set(folder_path)


def choose_save_folder(entry_var):
    folder_path = quick_open_folder_dialog().replace("/", "\\")
    if folder_path:
        entry_var.set(folder_path)


# ======================================================================================
# 运行
# ======================================================================================

def on_submit(excel_var, model_var, save_var, root_window, status_var, status_label, progress_bar, cal_config=None):
    excel_path = excel_var.get()
    model_folder = model_var.get()
    save_folder = save_var.get()

    if not excel_path or "xlsx" not in excel_path:
        messagebox.showinfo("提示", "请先选择有效的 Excel 批量参数表。")
        return
    if not model_folder or model_folder in ("未选取任何文件夹", ""):
        model_folder = os.path.dirname(excel_path)
    if not save_folder or save_folder in ("未选取任何文件夹", ""):
        save_folder = os.path.dirname(excel_path)

    if not messagebox.askokcancel("确认运行",
        f"即将批量生成计算书：\n\n"
        f"参数表：{excel_path}\n"
        f"模型文件夹：{model_folder}\n"
        f"输出文件夹：{save_folder}\n\n"
        f"请确认 Midas Civil 已启动。\n继续？"):
        return

    # 重置进度条
    progress_bar['value'] = 0
    progress_bar.configure(bootstyle=SUCCESS + STRIPED)
    status_var.set("正在初始化...")
    status_label.configure(bootstyle=INFO)
    root_window.update_idletasks()

    make_all_doc_batch(
        excel_path=excel_path,
        model_folder=model_folder,
        pbar=progress_bar,
        svar=status_var,
        root=root_window,
        cal_config=cal_config,
    )
    progress_bar.configure(bootstyle=SUCCESS)
    status_var.set("批量生成完成！")
    status_label.configure(bootstyle=SUCCESS)
    messagebox.showinfo("完成", f"批量计算书生成完成！\n结果已保存至：\n{save_folder}")


# ======================================================================================
# 主对话框
# ======================================================================================

def creat_dialog(parent=None, on_close=None):
    if parent is None:
        root = tb.Window("矩形围堰批量计算书exe")
    else:
        root = tb.Toplevel(parent)
        root.title("矩形围堰批量计算书exe")

    # --- 初始化 Midas API ---
    MidasURL()

    # --- 注册表读取 Midas 安装路径 ---
    reg_path2 = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"SOFTWARE\MIDAS\CVLwNX_CH\PATH")
    civilnx_path = winreg.QueryValueEx(reg_path2, "Installed Path")[0]
    partial_title = "MIDAS CIVIL NX"

    # --- 变量 ---
    program_state_var  = tb.StringVar(value=Program_State_Monitoring(civilnx_path, partial_title))
    excel_entry_var    = tb.StringVar(value="未选取任何文件")
    model_entry_var    = tb.StringVar(value="")
    save_entry_var     = tb.StringVar(value="")

    # 计算配置（需生成的计算书模块, 默认全选）
    cal_config = {
        "基坑内外土压力计算":   True,
        "土层水平反力系数计算": True,
        "模型结果读取":         True,
        "滑动稳定性分析":       True,
        "嵌固稳定性分析":       True,
        "抗隆起稳定性分析":     True,
    }

    # 计算配置（持久化存储）
    def open_cal_config():
        dlg = CalConfigDialog(root, init_config=cal_config)
        result = dlg.show()
        if result:
            cal_config.update(result)

    # --- 标题：文件操作 ---
    dialog_headr(root, "文件操作")

    # 程序状态 + 打开 Midas
    dialog_boxline3(root, "程序状态", 10, program_state_var, 40,
                    "打开Midas civil NX",
                    lambda: Open_Midas_civil(civilnx_path, program_state_var))

    # 选择 Excel 批量参数表
    dialog_boxline3(root, "参数表格", 10, excel_entry_var, 40,
                    "选择Excel批量表",
                    lambda: choose_excel_file(excel_entry_var))

    # 模型文件所在文件夹
    dialog_boxline3(root, "模型文件夹", 10, model_entry_var, 40,
                    "选择模型文件夹",
                    lambda: choose_model_folder(model_entry_var, save_entry_var))

    # 计算书输出文件夹
    dialog_boxline3(root, "结果保存", 10, save_entry_var, 40,
                    "选择输出文件夹",
                    lambda: choose_save_folder(save_entry_var))

    # --- 进度可视化区域 ---
    progress_frame = tb.Frame(root, padding=(20, 10))
    progress_frame.pack(fill=X, expand=YES)

    status_var = tb.StringVar(value="请完善上方文件配置...")
    status_label = tb.Label(progress_frame, textvariable=status_var, bootstyle=INFO)
    status_label.pack(side=TOP, anchor=W)

    progress_bar = tb.Progressbar(
        progress_frame,
        bootstyle=SUCCESS + STRIPED,
        maximum=100,
        value=0
    )
    progress_bar.pack(fill=X, pady=5)

    # --- 底部按钮: 计算配置 | 运行 | 关闭 ---
    dialog_boxline_end(root,
        "运行",
        lambda: on_submit(excel_entry_var, model_entry_var, save_entry_var,
                          root, status_var, status_label, progress_bar, cal_config),
        "关闭",
        lambda: on_exit(root),
        "计算配置",
        open_cal_config)

    # --- 状态校验 ---
    def check_ready(*args):
        state = program_state_var.get()
        excel = excel_entry_var.get()
        if "已打开" in state and "xlsx" in excel:
            status_var.set("准备就绪")
            status_label.configure(bootstyle=SUCCESS)
        else:
            progress_bar['value'] = 0
            status_var.set("请完善上方文件配置...")
            status_label.configure(bootstyle=INFO)

    program_state_var.trace_add("write", check_ready)
    excel_entry_var.trace_add("write", check_ready)
    check_ready()

    if parent is None:
        root.mainloop()
    elif on_close:
        def _on_root_destroy(event, _root=root):
            if event.widget is _root:
                on_close()
        root.bind('<Destroy>', _on_root_destroy)


# ======================================================================================
# 入口
# ======================================================================================

def CofferDam_Cal_Report_Main_Batch(parent=None, on_close=None):
    if not if_Reg(parent):
        if on_close:
            on_close()
        return

    creat_dialog(parent, on_close)
    

# CofferDam_Cal_Report_Main_Batch()