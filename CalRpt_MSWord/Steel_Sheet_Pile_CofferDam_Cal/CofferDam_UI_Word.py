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
from General.DataUtils      import read_Register
from General.ProgramMonitor import Program_State_Monitoring
from General.UIHandle       import dialog_headr, on_exit, if_Reg
from General.Midas          import MidasURL, get_current_midas_file, Open_Midas_civil

sys.path.append(str(Path(__file__).parent))
from CalRpt_MSWord.Steel_Sheet_Pile_CofferDam_Cal.CofferDam_Cal_Report_Generate import make_all_doc
from CalRpt_MSWord.Steel_Sheet_Pile_CofferDam_Cal.CofferDam_Cal_File_Manipulation import Open_Midas_mcb, Cal_path_SaveAs, load_all_projects_for_cal, get_excel_path


class SeepageSettingsDialog:
    """渗透稳定性验算参数设置 — 二级窗口, 回到主流程时只传结果 dict"""
    def __init__(self, parent):
        win = tb.Toplevel(parent)
        win.title("设置渗透稳定性验算参数")
        win.geometry(f"600x450+{parent.winfo_rootx()+50}+{parent.winfo_rooty()+50}")
        win.resizable(False, False)
        win.transient(parent)
        win.grab_set()
        self.win = win
        self.result = None

        # ===== 状态变量 =====
        self.ckb_heave  = tk.BooleanVar(value=True)  # 抗突涌
        self.ckb_flow   = tk.BooleanVar(value=True)  # 流土
        self.ckb_piping = tk.BooleanVar(value=False)  # 管涌
        self.d_heave    = tk.StringVar(value="3.0")   # 坑底至承压水顶面间距
        self.confined   = tk.BooleanVar(value=True)   # 承压水/潜水
        self.d_flow     = tk.StringVar(value="5.0")   # 承压水层顶面到基坑顶间距

        # ===== UI =====
        notebook = tb.Notebook(win)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # 抗突涌
        f0 = tb.Frame(notebook)
        notebook.add(f0, text="抗突涌稳定性验算")
        ck0 = tb.Checkbutton(f0, text="启用抗突涌稳定性验算", variable=self.ckb_heave, bootstyle="success-round-toggle")
        ck0.pack(anchor="w", padx=15, pady=(15, 10))
        frame0 = tb.Frame(f0)
        frame0.pack(fill="both", expand=True, padx=15, pady=5)
        # 左侧示意
        lf0 = tb.Labelframe(frame0, text="抗突涌稳定性验算示意图", bootstyle=SECONDARY, padding=5)
        lf0.pack(side="left", fill="both", expand=True, padx=(0, 10))
        # 右侧参数
        rf0 = tb.Frame(frame0)
        rf0.pack(side="left", fill="x")
        tb.Label(rf0, text="坑底至承压水顶面间距 D (m):").pack(anchor="w", pady=5)
        tb.Entry(rf0, textvariable=self.d_heave, width=12).pack(anchor="w")

        # 流土
        f1 = tb.Frame(notebook)
        notebook.add(f1, text="流土稳定性验算")
        ck1 = tb.Checkbutton(f1, text="启用流土稳定性验算", variable=self.ckb_flow, bootstyle="success-round-toggle")
        ck1.pack(anchor="w", padx=15, pady=(15, 10))
        frame1 = tb.Frame(f1)
        frame1.pack(fill="both", expand=True, padx=15, pady=5)
        lf1 = tb.Labelframe(frame1, text="流土稳定性验算示意图", bootstyle=SECONDARY, padding=5)
        lf1.pack(side="left", fill="both", expand=True, padx=(0, 10))
        rf1 = tb.Frame(frame1)
        rf1.pack(side="left", fill="x")
        tb.Label(rf1, text="地下水类型:").pack(anchor="w", pady=5)
        radio_frame1 = tb.Frame(rf1)
        radio_frame1.pack(anchor="w")
        tb.Radiobutton(radio_frame1, text="承压水", variable=self.confined, value=True).pack(side="left", padx=5)
        tb.Radiobutton(radio_frame1, text="潜水", variable=self.confined, value=False).pack(side="left", padx=5)
        tb.Label(rf1, text="承压水层顶面到基坑顶间距 (m):").pack(anchor="w", pady=(10, 5))
        entry_flow = tb.Entry(rf1, textvariable=self.d_flow, width=12)
        entry_flow.pack(anchor="w")
        self.confined.trace_add("write", lambda *_: entry_flow.configure(state="normal" if self.confined.get() else "disabled"))
        entry_flow.configure(state="normal" if self.confined.get() else "disabled")

        # ---- 页2: 管涌 ----
        f2 = tb.Frame(notebook)
        notebook.add(f2, text="管涌稳定性验算")
        ck2 = tb.Checkbutton(f2, text="启用管涌稳定性验算", variable=self.ckb_piping, bootstyle="success-round-toggle")
        ck2.pack(anchor="w", padx=15, pady=(15, 10))
        tb.Label(f2, text="当前计算选项待开发，仅提供参考公式",
                  foreground="gray", font=("微软雅黑", 10, "italic")).pack(pady=40)

        # ===== 底部按钮 =====
        bottom = tb.Frame(win)
        bottom.pack(fill="x", padx=15, pady=15)
        tb.Button(bottom, text="确认", bootstyle=SUCCESS, width=10, command=self._on_confirm).pack(side="right", padx=5)
        tb.Button(bottom, text="取消", bootstyle=SECONDARY, width=10, command=self._on_cancel).pack(side="right", padx=5)

    def _on_confirm(self):
        seepage_vars = []
        if self.ckb_heave.get():
            seepage_vars.append("抗突涌稳定性验算")
        if self.ckb_flow.get():
            seepage_vars.append("流土稳定性验算")
        if self.ckb_piping.get():
            seepage_vars.append("管涌稳定性验算")
        self.result = {
            'seepage_vars': seepage_vars,
            'if_confined': self.confined.get(),
            'd_confined': float(self.d_flow.get() or 0),
        }
        self.win.destroy()

    def _on_cancel(self):
        self.result = {}
        self.win.destroy()

    def show(self):
        self.win.wait_window()
        return self.result


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


# 点击运行
def click_confirm(Applocation, openfile_path, saveas_path, pbar, svar, root, all_res_var, seepage_params=None, callback=None, proj_key=None, cal_config=None):
    if "mcb" in openfile_path:
        pbar['value'] = 0
        pbar.configure(bootstyle=SUCCESS + STRIPED)
        svar.set("正在初始化...")
        root.update_idletasks()
        if "docx" in saveas_path:
            make_all_doc(Applocation, openfile_path, saveas_path, pbar=pbar, svar=svar, root=root, all_res_var=all_res_var, seepage_params=seepage_params, proj_key=proj_key, cal_config=cal_config)
            if callback and all_res_var.get():
                callback()
        else:
            print("未输入保存路径")
            messagebox.showwarning("警告", "未输入保存路径")
    else:
        messagebox.showerror("错误", "未识别到有效mcb文件")


# 根据下拉列表的值控制编辑框的编辑状态和编辑框的值
def update_entry_state(combo_var, ent):
    # 根据下拉列表选择更新编辑框状态
    if combo_var.get() == "自定义":
        ent.config(state="normal")
        ent.focus()
    if combo_var.get() != "自定义":
        ent.config(state="disabled")
        ent.focus()

#=====================================================================================================================================================================


# 标签 + 边界框 + 按钮
def dialog_boxline3(root, label_txt, label_length, entry_var, entry_length, button_txt, button_command):
    container = tb.Frame(root)
    container.pack(fill=X, expand=YES, pady=5)
    # 标签
    Label = tb.Label(master=container, text=label_txt.title(), width=label_length)
    Label.pack(side=LEFT, padx=0)
    # 编辑框
    ent = tb.Entry(master=container, textvariable=entry_var, width=entry_length, state="disabled")
    ent.pack(side=LEFT, padx=20, fill=X, expand=YES)
    # 按钮
    but = tb.Button(master=container, text=button_txt, command = button_command, width = 15)
    but.pack(side=RIGHT, padx=10)


# 最后一行
def dialog_boxline_end(root, button_txt1, button_command1, button_txt2, button_command2):
    container = tb.Frame(root)
    container.pack(fill=X, expand=YES, pady=10)
    # 退出按钮
    but = tb.Button(master=container, text=button_txt2, command = button_command2, width=6)
    but.pack(side=RIGHT, padx=20, pady=10)
    # 确认按钮
    sub_btn = tb.Button(master=container, text=button_txt1, command = button_command1, bootstyle=SUCCESS, width=6)
    sub_btn.pack(side=RIGHT, pady=10)

#=====================================================================================================================================================================

# 含主对话框全局变量的函数
# 对话框总函数
def creat_dialog(root, Applocation):

    # 获取base_url和reg_Key
    MidasURL()
    # 注册表路径
    reg_path2 = winreg.OpenKey(winreg.HKEY_CURRENT_USER,r"SOFTWARE\MIDAS\CVLwNX_CH\PATH")
    civilnx_path = winreg.QueryValueEx(reg_path2,"Installed Path")[0]
    # 窗口名
    partial_title = "MIDAS CIVIL NX"  # 你可以根据需要修改字符串
    # 自动识别当前模型
    current_active_mcb = get_current_midas_file()
    
    # 变量设置
    Set_File_entry_var1 = tb.StringVar(value = Program_State_Monitoring(civilnx_path, partial_title)) # 程序状态
    Set_File_entry_var2 = tb.StringVar(value = current_active_mcb if current_active_mcb else "未选取任何模型") # 选取模型
    Set_File_entry_var3 = tb.StringVar(value = current_active_mcb.replace(".mcb", ".docx") if current_active_mcb else os.path.join(Applocation, "Custom\\Untitled_Cal.docx")) # 另存为路径

    # 计算配置（需生成的计算书模块, 默认全选）
    cal_config = {
        "基坑内外土压力计算":   True,
        "土层水平反力系数计算": True,
        "模型结果读取":         True,
        "滑动稳定性分析":       True,
        "嵌固稳定性分析":       True,
        "抗隆起稳定性分析":     True,
    }

    # ===== 项目选择 =====
    dialog_headr(root, "项目选择")

    proj_frame = tb.Frame(root)
    proj_frame.pack(fill=X, expand=YES, pady=5)
    tb.Label(proj_frame, text="项目名称:").pack(side=LEFT, padx=(0, 10))
    project_name_combo = tb.Combobox(proj_frame, state="readonly", width=15)
    project_name_combo.pack(side=LEFT, padx=(0, 15))
    tb.Label(proj_frame, text="围堰编号:").pack(side=LEFT, padx=(0, 5))
    project_no_combo = tb.Combobox(proj_frame, state="readonly", width=15)
    project_no_combo.pack(side=LEFT, fill=X, expand=True, padx=(0, 10))

    # 读取 Excel 项目列表
    _current_proj_key = [None]
    try:
        _excel_path = get_excel_path()
        _load_result = load_all_projects_for_cal(_excel_path) if _excel_path else {"projects_dict": {}, "project_groups": {}, "last_proj_key": None}
    except Exception as e:
        print(f"[CalRpt] 读取 Excel 项目列表失败: {e}")
        _load_result = {"projects_dict": {}, "project_groups": {}, "last_proj_key": None}
    _projects_dict = _load_result["projects_dict"]
    _project_groups = _load_result["project_groups"]
    _last_proj_key = _load_result["last_proj_key"]

    # 填充项目名称下拉框
    project_name_combo['values'] = list(_project_groups.keys())

    def _on_project_name_change(event=None):
        pname = project_name_combo.get()
        nos = _project_groups.get(pname, [])
        project_no_combo['values'] = nos
        if nos:
            project_no_combo.set(nos[0])
            _current_proj_key[0] = f"{pname}{nos[0]}"
        else:
            project_no_combo.set("")
            _current_proj_key[0] = pname

    def _on_project_no_change(event=None):
        pname = project_name_combo.get()
        pno = project_no_combo.get()
        _current_proj_key[0] = f"{pname}{pno}" if pno else pname

    project_name_combo.bind("<<ComboboxSelected>>", _on_project_name_change)
    project_no_combo.bind("<<ComboboxSelected>>", _on_project_no_change)

    # 初始化选中最后一个项目
    if _last_proj_key and _last_proj_key in _projects_dict:
        _proj = _projects_dict[_last_proj_key]
        _pname = _proj.get("项目名称", "")
        _pno = _proj.get("围堰编号", "")
        project_name_combo.set(_pname)
        _nos = _project_groups.get(_pname, [])
        project_no_combo['values'] = _nos
        if _pno in _nos:
            project_no_combo.set(_pno)
        _current_proj_key[0] = _last_proj_key

    # 虚线分隔：项目选择 | 文件操作
    _sep1 = tk.Canvas(root, height=2, highlightthickness=0, bg=root.cget("background"))
    _sep1.pack(fill=X, padx=20, pady=(12, 8))
    _sep1.create_line(0, 1, 2000, 1, dash=(4, 4), fill="gray")

    # 标题：文件操作
    dialog_headr(root, "文件操作")

    # 程序状态、模型名称、保存路径
    dialog_boxline3(root, "程序状态", 10, Set_File_entry_var1, 40, "打开Midas civil NX", lambda: Open_Midas_civil(civilnx_path, Set_File_entry_var1))
    dialog_boxline3(root, "模型名称", 10, Set_File_entry_var2, 40, "选取mcb文件并打开", lambda: Open_Midas_mcb(Set_File_entry_var2,Set_File_entry_var3))
    dialog_boxline3(root, "保存路径", 10, Set_File_entry_var3, 40, "计算书另存为", lambda: Cal_path_SaveAs(Set_File_entry_var3, Applocation))

    all_res_var = tb.StringVar(value="等待计算完成...")

    # 渗透稳定性验算参数（持久化存储）
    seepage_params = {'seepage_vars': [], 'if_confined': True, 'd_confined': 5.0}
    def open_seepage_settings():
        dlg = SeepageSettingsDialog(root)
        result = dlg.show()
        if result:
            seepage_params.update(result)

    # 计算配置（持久化存储）
    def open_cal_config():
        dlg = CalConfigDialog(root, init_config=cal_config)
        result = dlg.show()
        if result:
            cal_config.update(result)

    # 虚线分隔：文件操作 | 进度
    _sep2 = tk.Canvas(root, height=2, highlightthickness=0, bg=root.cget("background"))
    _sep2.pack(fill=X, padx=20, pady=(12, 8))
    _sep2.create_line(0, 1, 2000, 1, dash=(4, 4), fill="gray")

    # 进度可视化区域
    progress_frame = tb.Frame(root, padding=(20, 10))
    progress_frame.pack(fill=X, expand=YES)
    status_var = tb.StringVar(value="请完善上方文件配置...")
    status_label = tb.Label(progress_frame, textvariable=status_var, bootstyle=INFO)
    status_label.pack(side=TOP, anchor=W)
    progress_bar = tb.Progressbar(progress_frame, bootstyle=SUCCESS + STRIPED, maximum=100, value=0)
    progress_bar.pack(fill=X, pady=5)

    # 底部按钮行: 计算配置 | 运行 | 关闭
    btn_container = tb.Frame(root)
    btn_container.pack(fill=X, expand=YES, pady=10)
    tb.Button(btn_container, text="关闭", command=lambda: on_exit(root), width=6, bootstyle = DANGER).pack(side=RIGHT, padx=20, pady=10)
    tb.Button(btn_container, text="运行",
               command=lambda: click_confirm(Applocation,
                   Set_File_entry_var2.get(), Set_File_entry_var3.get(),
                   progress_bar, status_var, root, all_res_var,
                   seepage_params=seepage_params, callback=show_res,
                   proj_key=_current_proj_key[0], cal_config=cal_config),
               bootstyle=SUCCESS, width=6).pack(side=RIGHT, pady=10)
    tb.Button(btn_container, text="计算配置", command=open_cal_config, width=10).pack(side=RIGHT, padx=20, pady=10)
    # ttk.Button(btn_container, text="设置渗透稳定性验算参数", command=open_seepage_settings,
    #            bootstyle=INFO, width=20).pack(side=RIGHT, padx=20, pady=10)

    def show_res():
        # 如果结论为空，则不执行更新显示的操作
        if not all_res_var.get():
            return
        raw_data = all_res_var.get()
        if "|TIME_STAMP|" in raw_data:
            # 拆分内容和时间
            content, start_time_str = raw_data.split("|TIME_STAMP|")
            # 重新设置变量内容（去掉时间标签，保证显示干净）
            all_res_var.set(content)
        else:
            start_time_str = ""
        # 展示结果
        # 创建独立弹窗
        top = tb.Toplevel(title="简要计算结果")
        # 弹窗位置
        root_x = root.winfo_x()
        root_y = root.winfo_y()
        root_width = root.winfo_width()
        top.geometry(f"600x600+{root_x + root_width + 5}+{root_y}")  # 调整位置到主窗口右侧
        # 弹窗内部布局
        # 标题
        full_path = Set_File_entry_var2.get()
        file_name = file_name = os.path.splitext(os.path.basename(full_path))[0] if "mcb" in full_path.lower() else "未知模型"
        window_title = f"{file_name} - 计算结果({start_time_str})"
        header = tb.Label(top, text=window_title, font=("微软雅黑", 12, "bold"))
        header.pack(fill=X, pady=10)
        # 内容
        from tkinter.scrolledtext import ScrolledText
        text_area = ScrolledText(top, font=("微软雅黑", 10), padx=15, pady=15, relief=FLAT)
        text_area.pack(fill=BOTH, expand=YES)
        text_area.insert("1.0", all_res_var.get())
        text_area.config(state=DISABLED)
        top.focus_set()
    
    # 校验函数
    def check_ready(*args):
        state = Set_File_entry_var1.get()
        model = Set_File_entry_var2.get()
        save_path = Set_File_entry_var3.get()
        if "已打开" in state and "mcb" in model and "docx" in save_path:
            status_var.set("准备就绪")
            status_label.configure(bootstyle=SUCCESS) 
        else:
            progress_bar['value'] = 0
            status_var.set("请完善上方文件配置...")
            status_label.configure(bootstyle=INFO)
    Set_File_entry_var1.trace_add("write", check_ready)
    Set_File_entry_var2.trace_add("write", check_ready)
    Set_File_entry_var3.trace_add("write", check_ready)
    check_ready()


def CofferDam_Cal_Report_Main(parent=None, on_close=None):
    """钢板桩围堰计算书自动化助手 — 主入口函数

    职责：
        1. 检查软件注册状态，未注册则直接退出
        2. 读取注册表中的软件安装路径
        3. 根据是否有父窗口，创建独立窗口(tb.Window)或子窗口(tb.Toplevel)
        4. 构建对话框界面并启动主事件循环

    Args:
        parent   : 父窗口对象，为 None 时作为独立窗口运行
        on_close : 子窗口关闭时的回调函数，仅在 parent 不为 None 时生效
    """
    # 检查软件是否已注册，未注册则回调 on_close 并退出
    if not if_Reg(parent):
        if on_close:
            on_close()
        return
    # 从注册表读取软件安装路径，供后续对话框加载资源使用
    Applocation = read_Register('Software\\ShuZhiQiaoShi', 'Applocation')
    # 根据是否有父窗口决定创建独立窗口还是子窗口
    if parent == None:
        root = tb.Window("钢板桩围堰计算书自动化助手")
    else:
        root = tb.Toplevel(parent)
        root.title("钢板桩围堰计算书自动化助手")
    # 构建对话框界面（包含文件配置、参数设置等控件）
    creat_dialog(root, Applocation)
    # 独立窗口模式：启动主事件循环，阻塞直到窗口关闭
    if parent == None:
        root.mainloop()
    # 子窗口模式：绑定窗口销毁事件，触发回调通知父窗口
    elif on_close:
        def _on_root_destroy(event, _root=root):
            if event.widget is _root:
                on_close()
        root.bind('<Destroy>', _on_root_destroy)


# CofferDam_Cal_Report_Main()