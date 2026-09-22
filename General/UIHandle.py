# 1. 标准库
import os
import sys
import gc
import time
import ctypes
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox

# 2. 第三方库
import rsa
import ttkbootstrap as tb
import matplotlib.pyplot as _plt
from ttkbootstrap import Style, Menu
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledFrame
from PIL import Image, ImageTk

# 3. 本地模块
from General.TextHandle import read_file, update_file
from General.DataUtils  import if_Expired, get_unique_id, read_Register, Recover_code, delete_registry_value, write_Register, get_revit_install_path


class ToolTip(object):
    TOOLTIP_FONT = ("Microsoft YaHei", 9, "bold")  
    HOVER_DELAY_MS = 0
    OFFSET_X = 5   # 相对鼠标指针右下偏移
    OFFSET_Y = 5
    def __init__(self, widget):
        self.widget = widget
        self.tipwindow = None
        self._hover_job = None
        self._x = None
        self._y = None
    def showtip(self, text):
        self.text = text
    def _showtip_now(self):
        if self.tipwindow or not self.text:
            return
        try:
            # 默认跟随鼠标指针；若未捕获指针坐标则回落控件右上角
            if self._x is None or self._y is None:
                x = self.widget.winfo_rootx() + self.widget.winfo_width() + 4
                y = self.widget.winfo_rooty() - 4
            else:
                x = self._x + self.OFFSET_X
                y = self._y + self.OFFSET_Y
        except Exception:
            return
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.attributes("-topmost", True)
        lbl = tb.Label(tw, text=self.text, justify=tk.LEFT, bootstyle="inverse-primary", font=self.TOOLTIP_FONT)
        lbl.pack(ipadx=8, ipady=5)
        # 先应用位置，再量测尺寸做轻微防出界
        tw.update_idletasks()
        try:
            screen_w = self.widget.winfo_screenwidth()
            screen_h = self.widget.winfo_screenheight()
            tip_w = tw.winfo_reqwidth()
            tip_h = tw.winfo_reqheight()
            if x + tip_w > screen_w:
                x = screen_w - tip_w - 8
            if y + tip_h > screen_h:
                y = screen_h - tip_h - 8
        except Exception:
            pass
        tw.wm_geometry("+%d+%d" % (x, y))
    def _schedule_show(self, event=None):
        self._cancel_show()
        # 记录鼠标指针位置
        if event is not None:
            self._x = getattr(event, "x_root", self._x)
            self._y = getattr(event, "y_root", self._y)
        if self.text:
            self._hover_job = self.widget.after(self.HOVER_DELAY_MS, self._showtip_now)
    def _cancel_show(self):
        if self._hover_job is not None:
            self.widget.after_cancel(self._hover_job) 
            self._hover_job = None
    def hidetip(self):
        self._cancel_show()
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()


def createToolTip(widget, text):
    toolTip = ToolTip(widget)
    toolTip.showtip(text)
    widget.bind('<Enter>', lambda e: toolTip._schedule_show(e))
    widget.bind('<Leave>', lambda e: toolTip.hidetip())
    widget.bind('<ButtonPress>', lambda e: toolTip.hidetip(), add="+")
    widget._tooltip = toolTip
    return toolTip


# 重新启用主窗口的按钮，并销毁二级窗口
def enable_button(secondary_window, button):
    """重新启用主窗口的按钮，并销毁二级窗口"""
    button.config(state='normal')
    secondary_window.destroy()


def on_exit(root):
    # 1. 遍历所有子控件，清除可能存在的 image 引用
    for widget in root.winfo_children():
        # 递归处理子控件（如果控件内部还有子控件）
        _clear_image_recursive(widget)
    # 2. 主动触发一次垃圾回收
    gc.collect()
    # 3. 关闭 matplotlib 图
    _plt.close('all')
    # 4. 销毁窗口
    root.destroy()


def _clear_image_recursive(widget):
    """递归清除控件及其子控件的 image 属性"""
    # 检查当前控件是否有 image 属性
    if hasattr(widget, 'image') and widget.image:
        widget.image = None
    # 递归处理子控件（如果是容器类控件）
    for child in widget.winfo_children():
        _clear_image_recursive(child)


# 标题栏
def dialog_headr(root, txt):
    # 宋体加粗
    style = tb.Style()
    style.configure("SimSun.TLabel", font=("SimSun", 10, "bold"))
    # container = tb.Frame(root)
    # container.pack(fill=X, expand=YES)
    hdr = tb.Label(master=root, text=txt, width=50, style="SimSun.TLabel")
    hdr.pack(fill=X, pady=10)


def show_warning_dialog(massage):
    """使用 Windows API 显示警告对话框"""
    result = ctypes.windll.user32.MessageBoxW(
        0,  # 无父窗口
        massage,  # 消息内容
        "Warning",  # 标题
        0x30 | 0x0  # 警告图标 + 确定按钮 (MB_ICONWARNING | MB_OK)
    )
    return result


def open_smart_recommend_ui():
    """启动本项目（AIrecommender）的智能推荐桌面UI

    通过子进程启动 AIrecommender.ui.app，避免阻塞当前UI；
    不能用 `python -m AIrecommender.ui.app`：打包后 app.py 被编译为 .pyd，
    runpy 无法执行扩展模块（No code object available），
    因此改用 `-c` 直接导入模块并调用 main()：
    - PyStand 环境：PyStand.int 识别 -c 参数后执行
    - 普通 Python 环境：python.exe 原生支持 -c
    若推荐系统已在运行（单例端口被占用）则只提示、不重复启动。
    """
    import socket
    from pathlib import Path

    # 推荐系统单例探测端口，需与 AIrecommender/ui/app.py 中 main() 的绑定端口保持一致
    SINGLETON_PORT = 58901

    # 1. 定位 AIrecommender 入口：环境变量优先，否则自动探测
    ai_home = os.environ.get('BRIDGE_TEMP_AI_HOME')
    if ai_home:
        ai_root = Path(ai_home)
    else:
        # 自动探测：从 __file__ 向上查找 AIrecommender/ui
        file_dir = Path(__file__).resolve().parent  # General/
        if (file_dir.parent / 'AIrecommender' / 'ui').exists():
            ai_root = file_dir.parent
        else:
            show_warning_dialog(f'未找到智能推荐系统入口\n(AIrecommender/ui)\n\n可通过环境变量 BRIDGE_TEMP_AI_HOME 指定项目根目录。')
            return

    # 确定 cwd（使 AIrecommender 包可导入）
    cwd_dir = str(ai_root)

    # 2. 已有实例则不重复启动
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        sock.settimeout(0.3)
        if sock.connect_ex(('127.0.0.1', SINGLETON_PORT)) == 0:
            show_warning_dialog('智能推荐系统已在运行，请查看其窗口。')
            return
    finally:
        sock.close()

    # 3. 后台启动推荐系统UI
    #    子进程可执行文件：PyStand 环境为 PyStand.exe（识别 -c 参数执行），普通环境为 python.exe
    #    额外把项目根与 site-packages 加入 PYTHONPATH，保证普通 python（含 run.bat 调试模式）能解析到依赖
    try:
        env = dict(os.environ)
        extra_paths = [str(ai_root / 'site-packages'), str(ai_root)]
        extra_paths = [p for p in extra_paths if os.path.isdir(p)]
        if extra_paths:
            old_pp = env.get('PYTHONPATH', '')
            env['PYTHONPATH'] = os.pathsep.join(extra_paths + ([old_pp] if old_pp else []))
        subprocess.Popen(
            [sys.executable, '-c', 'from AIrecommender.ui.app import main; main()'],
            cwd=cwd_dir,
            env=env,
        )
    except Exception as e:
        show_warning_dialog(f'启动智能推荐系统失败:\n{e}')


# 读取注册表中的值判断是否有注册过
def if_Reg(parent=None):
    # 时间验证
    if_Expired_True_False, if_Expired_error_massage = if_Expired()
    print(if_Expired_True_False, if_Expired_error_massage)
    if not if_Expired_True_False:
        show_warning_dialog(if_Expired_error_massage)
        return False
    # 公钥
    privkey = rsa.PrivateKey(9956464146951211552826016701021265250479789780306620429905373908593176984175566297278625676443067589165683650782333253548337395952065689382179059836314329, 65537, 2438488274146114357300614829303940194629768005305423875374386325386103783505890155683159258207620917851664381866099860954152147801879153757810913380629505, 7063515761135576043584250450282154072071044441059338431153376108851625369603348313, 1409562105280918435173821236903389828790022619992026188852424363074140033)
    # 注册表路径、主键、键值
    Reg_path = "Software\\ShuZhiQiaoShi"
    Reg_key = "RequestCode_Python"
    Reg_key_real = "RegistrationCode_Python"
    Reg_value = get_unique_id()
    # 读取注册码, 没有则返回NotFound
    zhucema_value = read_Register(Reg_path, Reg_key_real)
    if zhucema_value == "NotFound":
        # 没有注册码，打开注册程序
        print("未注册, 请先进行注册")
        return _show_reg_dialog(parent, Reg_path, Reg_value, privkey, Reg_key_real)
    else:
        # 有注册码, 判断注册码还原后是不是与原信息一致
        if Recover_code(zhucema_value, privkey) == Reg_value:
            print("已注册")
            return True
        else:
            print("注册码已失效, 请重新注册")
            delete_registry_value(Reg_path, Reg_key_real)
            return _show_reg_dialog(parent, Reg_path, Reg_value, privkey, Reg_key_real)
        

def _show_reg_dialog(parent, Reg_path, Reg_value, privkey, Reg_key_real):
    
    def _on_closing_reg(root, result):
        """点X关闭"""
        result[0] = False
        root.destroy()

    """显示注册对话框，返回 True(注册成功) 或 False(取消)"""
    result = [False]  # 用列表承载结果，供回调修改
    if parent is None:
        root = tk.Tk()
    else:
        root = tk.Toplevel(parent)
    root.title("注册对话框")
    root.protocol("WM_DELETE_WINDOW", lambda: _on_closing_reg(root, result))
    # 创建标签和输入框
    label_application = tk.Label(root, text="申请码:")
    label_application.grid(row=0, column=0, padx=10, pady=10, sticky="w")
    entry_application = tk.Entry(root, width=50)
    entry_application.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
    entry_application.insert(0, Reg_value)
    entry_application.config(state="readonly")
    label_registration = tk.Label(root, text="注册码:")
    label_registration.grid(row=1, column=0, padx=10, pady=10, sticky="w")
    entry_registration = tk.Entry(root, width=50)
    entry_registration.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
    # 创建一个Frame容器来放置按钮
    button_frame = tk.Frame(root)
    button_frame.grid(row=2, column=0, columnspan=2, pady=10)
    # 创建确定和取消按钮
    button_confirm = tk.Button(button_frame, text="确定", command=lambda: _on_confirm_reg(root, entry_registration.get(), privkey, Reg_path, Reg_key_real, result))
    button_confirm.pack(side="left", padx=10)
    button_cancel = tk.Button(button_frame, text="取消", command=lambda: _on_closing_reg(root, result))
    button_cancel.pack(side="left", padx=10)
    # 设置列权重，使输入框随窗口拉伸
    root.columnconfigure(1, weight=1)
    if parent is None:
        root.mainloop()
    else:
        root.grab_set()        # 模态对话框
        root.wait_window()     # 阻塞直到对话框关闭
    return result[0]


def _on_confirm_reg(root, registration_code, privkey, Reg_path, Reg_key_real, result):
    """确定按钮：验证注册码"""
    try:
        Recover_code(registration_code, privkey)
        write_Register(Reg_path, Reg_key_real, registration_code)
        result[0] = True
        root.destroy()
    except:
        messagebox.showerror("Error", "注册码不正确")


def show_register_window(parent=None):
    """直接弹出注册窗口（用于帮助-注册）。返回 True(注册成功) 或 False(取消/关闭)"""
    privkey = rsa.PrivateKey(9956464146951211552826016701021265250479789780306620429905373908593176984175566297278625676443067589165683650782333253548337395952065689382179059836314329, 65537, 2438488274146114357300614829303940194629768005305423875374386325386103783505890155683159258207620917851664381866099860954152147801879153757810913380629505, 7063515761135576043584250450282154072071044441059338431153376108851625369603348313, 1409562105280918435173821236903389828790022619992026188852424363074140033)
    Reg_path = "Software\\ShuZhiQiaoShi"
    Reg_key_real = "RegistrationCode_Python"
    Reg_value = get_unique_id()
    return _show_reg_dialog(parent, Reg_path, Reg_value, privkey, Reg_key_real)


# 恢复并置顶主窗口
def restore_main_window(root):
    """恢复并置顶主窗口"""
    root.deiconify()  # 重新显示窗口
    # 显式居中窗口（避免 withdraw→deiconify 后窗口管理器放到左上角）
    root.update_idletasks()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    main_width = root.winfo_width()
    main_height = root.winfo_height()
    x = int((screen_width - main_width) / 2)
    y = int((screen_height - main_height - 300) / 2)
    root.geometry(f"+{x}+{y}")
    root.lift()       # 将窗口提升到最上层
    root.focus_force()  # 强制获取焦点
    root.attributes('-topmost', True)  # 临时置顶
    # 短暂置顶后恢复正常状态
    root.after(100, lambda: root.attributes('-topmost', False))


# 监控外部进程状态的线程函数
def monitor_exe_process(process, root, button):
    """监控外部进程状态的线程函数"""
    while process.poll() is None:  # 进程仍在运行
        time.sleep(0.5)  # 每0.5秒检查一次
    # 进程已结束，重新显示主窗口
    button.config(state='normal') # 按钮恢复
    root.after(0, lambda: restore_main_window(root))


# 最小化按钮
def button_withdraw(root, mini):
    root.withdraw()
    mini[0] = True


# 更新按钮的状态
def reset_button_state(button):
    button.state = 'disabled'


# 标头栏
def Main_Headline(self, txt):
    # 宋体加粗
    style = tb.Style()
    style.configure("SimSun.TLabel", font=("SimSun", 10, "bold"))
    container = tb.Frame(self)
    container.pack(fill=X, pady=0)
    hdr = tb.Label(master=self, text=txt, width=50, style="SimSun.TLabel")
    hdr.pack(fill=X, padx=10, pady=10)


def parse_blocks_from_lines(lines):
    """
    参数: lines: list of str，文本文件的每一行（末尾换行符已去除）
    返回: dict: { block_name: [list_of_data_lines] }
    """
    blocks = {}
    current_key = None
    current_data = []
    for line in lines:
        line = line.strip()
        # 空行作为块结束分隔符
        if line == '':
            if current_key is not None:
                blocks[current_key] = current_data
                current_key = None
                current_data = []
            continue
        # 标题行：以 '*' 开头
        if line.startswith('*'):
            if current_key is not None:
                blocks[current_key] = current_data
            current_key = line.lstrip('*').strip()
            current_data = []
        else:
            if current_key is not None:
                current_data.append(line)  # 或 current_data.append(line.split())
    # 处理最后可能没有结尾空行的情况
    if current_key is not None:
        blocks[current_key] = current_data
    return blocks


def Boxline0(self, label, variable, length):
    container = tb.Frame(self)
    container.pack(fill=X, padx=10, pady=10)
    lbl = tb.Label(master=container, text=label.title(), width=length)
    lbl.pack(side=LEFT, padx=10)
    ent = tb.Entry(master=container, textvariable=variable, width=7)
    ent.pack(side=RIGHT, padx=10, fill=X)


def BoxLine1(self, label1, label2, variable1, variable2, length1, length2):
    container = tb.Frame(self)
    container.pack(fill=X, pady=10)
    lbl1 = tb.Label(master=container, text=label1.title(), width=length1)
    lbl1.pack(side=LEFT, padx=10)
    ent1 = tb.Entry(master=container, textvariable=variable1, width=7)
    ent1.pack(side=LEFT, padx=10, fill=X)
    ent2 = tb.Entry(master=container, textvariable=variable2, width=7)
    ent2.pack(side=RIGHT, padx=10, fill=X)
    lbl2 = tb.Label(master=container, text=label2.title(), width=length2)
    lbl2.pack(side=RIGHT, padx=10)


def BoxLine2(self, label, lst, secondary_window_func, combo_value, value_dict):
    string_vars_dict = {}
    # 获取选择值
    def on_select(event):
        combo_value[0] = combo.get()
        # 清除之前的对话框内容
        for widget in dialog.winfo_children():
            widget.destroy()
        if combo_value[0] != "自定义":
            string_vars_dict.clear()
            new_string_vars = secondary_window_func(dialog, combo_value[0], value_dict)
            string_vars_dict.update(new_string_vars)
        else:
            label_hint = tb.Label(dialog, text="该荷载无添加", font=('Arial', 10), foreground="blue")
            label_hint.pack(pady=0)
    container = tb.LabelFrame(self, text = label)
    container.pack(fill=X, padx=10, pady=10)
    frame1 = tb.Frame(container)
    frame1.pack(fill=X, pady=10)
    frame2 = tb.Frame(container)
    frame2.pack(fill=X, pady=0)
    # 下拉列表
    combo = tb.Combobox(master=frame1)
    combo.pack(side=LEFT, padx=20, fill=X)
    # 设置下拉列表的选项
    combo['values'] = lst
    combo.set(combo_value[0])
    # 绑定选择事件
    combo.bind('<<ComboboxSelected>>', on_select)
    dialog = tb.Frame(frame2)
    dialog.pack(fill=X, padx=10, pady=10)
    if combo_value[0] != "自定义":
        string_vars_dict.update(secondary_window_func(dialog, combo_value[0], value_dict))
    else:
        label_hint = tb.Label(dialog, text="该荷载无添加", font=('Arial', 10), foreground="blue")
        label_hint.pack(pady=0)
    return combo, string_vars_dict


def BoxLine3(self, labelTXT, labellst, variablelst, length):
    container = tb.LabelFrame(self, text = labelTXT)
    container.pack(fill=X, padx=10, pady=10)
    for i, label in enumerate(labellst):
        variable = variablelst[i]
        Boxline0(container, label, variable, length)


def create_Feng_dialog(Feng_dialog, standardname, dialog_values):
    Vars_dict = {}
    # 随规范变化的参数名称
    if standardname == "公路桥梁抗风设计规范":
        name_Feng = "基本风速U10(m/s)："
        name_K = "地形条件系数kt："
    # 下拉列表
    DBFL_lst = ["A:海面、海岸、开阔水面", "B:田野、乡村、丛林、平坦开阔地", "C:树木及地层建筑密集区、平缓丘陵地", "D:中高层建筑密集区、起伏较大的丘陵地"]
    formula_lst = ["Ud = kf·kt·kh·U10", "Ud = kf·(Z/10)^α0·Us10"]
    # 检查dialog_values中是否已经存在过数值
    for key, key_str in dialog_values.items():
        Vars_dict[key] = tb.StringVar(value=key_str)
    # 基本风速
    frame_label_entry(Feng_dialog, name_Feng, Vars_dict['wind_entry1'])
    # 地表分类
    frame_combobox1(Feng_dialog, "地表分类：", DBFL_lst, Vars_dict['wind_combo1'])
    # 地形修正
    frame_label_entry(Feng_dialog, name_K, Vars_dict['wind_entry2'])
    # 计算高度
    frame_label_entry(Feng_dialog, "主梁基准高度(m)：", Vars_dict['wind_entry3'])
    # if standardname == "公路桥梁抗风设计规范":
    # 水平加载长度
    frame_label_entry(Feng_dialog, "水平加载长度(m)：", Vars_dict['wind_entry4'])
    # 迎风贝雷梁片数
    frame_label_entry(Feng_dialog, "迎风贝雷梁片数：", Vars_dict['wind_entry5'])
    # 主梁横向力系数
    frame_label_entry(Feng_dialog, "主梁横向力系数：", Vars_dict['wind_entry6'])
    # 设计基准风速计算公式
    frame_combobox2(Feng_dialog, "设计基准风速计算公式", formula_lst, Vars_dict['wind_combo2'])
    return Vars_dict


# 风荷载计算对话框
# 对于港口工程荷载规范，水流力的基本参数为环境ENV，设计流速V，水位高程H
def create_Shui_dialog(Shui_dialog, standardname, dialog_values):
    Vars_dict = {}
    # 随规范变化的参数名称
    if standardname == "港口工程荷载规范":
        name_ENV = "水密度(t/m3)："
        name_V = "水流设计流速(m/s)："
        name_H = "水面相对高程(m)："
    # 检查dialog_values中是否已经存在过数值
    for key, key_str in dialog_values.items():
        Vars_dict[key] = tb.StringVar(value=key_str)
    # 水密度
    frame_label_entry(Shui_dialog, name_ENV, Vars_dict['flow_entry1'])
    # 设计流速
    frame_label_entry(Shui_dialog, name_V, Vars_dict['flow_entry2'])
    # 水面高程
    frame_label_entry(Shui_dialog, name_H, Vars_dict['flow_entry3'])
    return Vars_dict


# 标签+编辑框
def frame_label_entry(dialog, txt, variable):
    # frame
    frame = tb.Frame(dialog)
    frame.pack(fill=tb.BOTH)
    # 标签
    label = tb.Label(master=frame, text=txt.title(), width=20)
    label.pack(side=LEFT, padx=10, pady=10)
    # 编辑框
    entry = tb.Entry(master=frame, textvariable=variable, width=20)
    entry.pack(side=RIGHT, padx=10, pady=10, fill=X)


# 下拉列表1
def frame_combobox1(dialog, txt, lst, combo_var1):
    # frame
    frame = tb.Frame(dialog)
    frame.pack(fill=tb.BOTH)
    # 标签
    label = tb.Label(master=frame, text=txt.title(), width=20)
    label.pack(side=LEFT, padx=10, pady=10)
    # 下拉列表
    combo1 = tb.Combobox(master=frame, textvariable=combo_var1, width=18)
    combo1.pack(side=RIGHT, padx=10, pady=10)
    # 设置下拉列表的选项
    combo1['values'] = lst
    # 设置默认值
    if combo_var1.get() == "":
        combo1.current(0)


# 下拉列表2
def frame_combobox2(dialog, txt, lst, combo_var2):
    # frame
    frame = tb.Frame(dialog)
    frame.pack(fill=tb.BOTH)
    # 标签
    label = tb.Label(master=frame, text=txt.title(), width=20)
    label.pack(side=LEFT, padx=10, pady=10)
    # 下拉列表
    combo2 = tb.Combobox(master=frame, textvariable=combo_var2, width=18)
    combo2.pack(side=RIGHT, padx=10, pady=10)
    # 设置下拉列表的选项
    combo2['values'] = lst
    # 设置默认值
    if combo_var2.get() == "":
        combo2.current(0)


# 支架基本参数设置
def Basic_Param_Set_ZJ(root, Applocation):
    path = os.path.join(Applocation, 'Support\\basic_param\\Basic_Param_Set_ZJ.txt')
    basic_param_lst = read_file(path)
    Self_values_lst = basic_param_lst[0] # -1.0, 26.5, 2.0, 2.5
    FengGui_selected_value = basic_param_lst[1] # ["自定义"] # 选择规范默认值
    ShuiGui_selected_value = basic_param_lst[3] # ["自定义"] # 选择规范默认值
    Feng_values_lst = basic_param_lst[2] # ['24.5', '1.0', '20', '50', '8', '1.3', 'A:海面、海岸、开阔水面', 'Ud = kf·kt·kh·U10']
    Shui_values_lst = basic_param_lst[4] # ['1.025', '2.0', '-5.0']
    toplevel = tb.Toplevel(root)
    toplevel.title("支架基本参数设置")
    toplevel.geometry("600x800")
    scrollbar = ScrolledFrame(toplevel, autohide=True, height=200)
    scrollbar.pack(fill=BOTH, expand=YES, padx=10, pady=10)
    SelfWeight = tb.DoubleVar(value=float(Self_values_lst[0]))
    ConcreteG = tb.DoubleVar(value=float(Self_values_lst[1]))
    SGLoad = tb.DoubleVar(value=float(Self_values_lst[2]))
    MBLoad = tb.DoubleVar(value=float(Self_values_lst[3]))
    FengGui_lst = ["自定义", "公路桥梁抗风设计规范"]
    ShuiGui_lst = ["自定义", "港口工程荷载规范"]
    Feng_values_dict = {"wind_entry1":'24.5', "wind_entry2":'1.0', "wind_entry3":'20', "wind_entry4":'50', "wind_entry5":'8', "wind_entry6":'1.3', "wind_combo1":'A:海面、海岸、开阔水面', "wind_combo2":'Ud = kf·kt·kh·U10'} # 风荷载参数词典
    Shui_values_dict = {"flow_entry1":'1.025', "flow_entry2":'2.0', "flow_entry3":'-5.0'} # 水流力参数词典
    if FengGui_selected_value != ['自定义']:
        for i,key in enumerate(Feng_values_dict.keys()):
            Feng_values_dict[key] = Feng_values_lst[i]
    if ShuiGui_selected_value != ['自定义']:
        for i,key in enumerate(Shui_values_dict.keys()):
            Shui_values_dict[key] = Shui_values_lst[i]
    Main_Headline(scrollbar, "基本参数定义：")
    BoxLine1(scrollbar, "自重系数:", "混凝土容重(kN/m3):", SelfWeight, ConcreteG, 15, 15)
    BoxLine1(scrollbar, "施工荷载(kN/m2):", "模板荷载(kN/m2):", SGLoad, MBLoad, 15, 15)
    combo1, string_vars_dict1 = BoxLine2(scrollbar, "风荷载", FengGui_lst, create_Feng_dialog, FengGui_selected_value, Feng_values_dict)
    combo2, string_vars_dict2 = BoxLine2(scrollbar, "水流力", ShuiGui_lst, create_Shui_dialog, ShuiGui_selected_value, Shui_values_dict)
    # 创建保存按钮的函数
    def save_parameters():
        """保存所有参数到文件"""
        if combo1.get() == "自定义":
            Feng_values = ['无']
        else:
            # 获取所有参数值
            Feng_values = []
            for key in Feng_values_dict.keys():
                if key in string_vars_dict1:
                    value = string_vars_dict1[key].get()
                    Feng_values.append(value)
                    Feng_values_dict[key] = value  # 同时更新字典
                else:
                    Feng_values.append("")
        if combo2.get() == "自定义":
            Shui_values = ['无']
        else:
            Shui_values = []
            for key in Shui_values_dict.keys():
                if key in string_vars_dict2:
                    value = string_vars_dict2[key].get()
                    Shui_values.append(value)
                    Shui_values_dict[key] = value  # 同时更新字典
                else:
                    Shui_values.append("")
        # 构造保存的内容
        save_content = [[SelfWeight.get(), ConcreteG.get(), SGLoad.get(), MBLoad.get()], [combo1.get()], Feng_values, [combo2.get()], Shui_values]
        print(save_content)
        update_file(path, save_content)
    button_frame = tb.Frame(toplevel)
    button_frame.pack(padx=10, pady=10)
    # 添加取消/关闭按钮
    cancel_button = tb.Button(button_frame, text="关闭",command=toplevel.destroy,width=5)
    cancel_button.pack(side=RIGHT, padx=10)
    # 添加保存按钮
    save_button = tb.Button(button_frame, text="保存", command=save_parameters, bootstyle=SUCCESS, width=5)
    save_button.pack(side=RIGHT, padx=10)
    toplevel.mainloop()


# 围堰基本参数设置
def Basic_Param_Set_WY(root, Applocation):
    path = os.path.join(Applocation, 'Support\\basic_param\\Basic_Param_Set_WY2.txt')
    with open(path, 'r', encoding='utf-8') as f:
        raw_lines = f.readlines()
    # 去除每行的换行符
    raw_lines = [line.rstrip('\n\r') for line in raw_lines]
    blocks = parse_blocks_from_lines(raw_lines)

    # 提取规范信息 (*Formula 块)
    Formula_data = blocks.get('Formula', [])
    Formula_data_lst = [row.split() for row in Formula_data]

    # 提取参数信息 (*Param1 块)
    Param4_data = blocks.get('Param1', [])
    Param4_data_lst = [row.split() for row in Param4_data]

    CofferDam_Selected_Formula = Formula_data_lst[0]
    CofferDam_Basic_paramlst = Param4_data_lst[0]
    CofferDam_Formula = ['建筑基坑支护技术规程']
    Importance_Factor_lst = ['一级(结构重要性系数γ0=1.1)', '二级(结构重要性系数γ0=1.0)', '三级(结构重要性系数γ0=0.9)']
    Method_for_Pressures_lst = ["水土合算", "水土分算"]
    CofferDam_Formula_var = tb.StringVar(value=CofferDam_Selected_Formula[0])
    Importance_Factor_var = tb.StringVar(value=CofferDam_Basic_paramlst[0])
    Coefficient_Self_Weight = tb.DoubleVar(value=CofferDam_Basic_paramlst[1])
    Ea_sigma0 = tb.DoubleVar(value=CofferDam_Basic_paramlst[2])
    Method_for_Pressures_var = tb.StringVar(value=CofferDam_Basic_paramlst[3])
    toplevel = tb.Toplevel(root)
    toplevel.title("围堰基本参数设置")
    def BoxLine(param, label_txt, combo_var, combo_lst, botton_txt, button_command):
        container = tb.Frame(toplevel)
        container.pack(fill=X, expand=YES, pady=5)
        label1 = tb.Label(master=container, text=label_txt, width=15)
        label1.pack(side=LEFT, padx=10, pady=10, fill=X, expand=YES)
        if param == "combo":
            combo = tb.Combobox(master=container, textvariable=combo_var, values=combo_lst, width=25)
            combo.pack(side=RIGHT, padx=10, pady=10, fill=X, expand=YES)
        elif param == 'button':
            button = tb.Button(master=container, text=botton_txt, command=button_command, width=26)
            button.pack(side=RIGHT, padx=10, pady=10, fill=X, expand=YES)
        elif param == 'entry':
            entry = tb.Entry(master=container, textvariable=combo_var, width=27)
            entry.pack(side=LEFT, padx=10, pady=10, fill=X, expand=YES)
    def save_parameters():
        save_content = [
            ['*Formula'],
            [CofferDam_Formula_var.get()],
            [''],
            ['*Param1'],
            [Importance_Factor_var.get(), Coefficient_Self_Weight.get(), Ea_sigma0.get(), Method_for_Pressures_var.get()],
            [''],
            [''],
        ]
        # print(save_content)
        update_file(path, save_content, ' ')
    Main_Headline(toplevel, "基本参数定义：")
    # 荷载计算规范选取
    BoxLine("combo", "荷载计算规范", CofferDam_Formula_var, CofferDam_Formula, '', '')
    # 围堰安全等级
    BoxLine("combo", "围堰安全等级", Importance_Factor_var, Importance_Factor_lst, '', '')
    # 荷载自重系数
    BoxLine("entry", "荷载自重系数", Coefficient_Self_Weight, '', '', '')
    # 围堰外初始附加应力
    BoxLine("entry", "初始附加应力(kPa)", Ea_sigma0, '', '', '')
    # 水土压力计算方法
    BoxLine("combo", "荷载计算方法", Method_for_Pressures_var, Method_for_Pressures_lst, '', '')
    button_frame = tb.Frame(toplevel)
    button_frame.pack(padx=10, pady=10)
    # 添加取消/关闭按钮
    cancel_button = tb.Button(button_frame,text="关闭",command=toplevel.destroy,width=5)
    cancel_button.pack(side=RIGHT, padx=10)
    # 添加保存按钮
    save_button = tb.Button(button_frame, text="保存", command=save_parameters, bootstyle=SUCCESS, width=5)
    save_button.pack(side=RIGHT, padx=10)
    toplevel.mainloop()


# 栈桥基本参数设置
def Basic_Param_Set_ZQ(root, Applocation):
    path = os.path.join(Applocation, 'Support\\basic_param\\Basic_Param_Set_ZQ.txt')
    basic_param_lst = read_file(path)
    Self_values_lst = basic_param_lst[0] # -1.0
    FengGui_selected_value = basic_param_lst[1] # ["自定义"] # 选择规范默认值
    ShuiGui_selected_value = basic_param_lst[3] # ["自定义"] # 选择规范默认值
    Feng_values_lst = basic_param_lst[2] # ['24.5', '1.0', '20', '50', '8', '1.3', 'A:海面、海岸、开阔水面', 'Ud = kf·kt·kh·U10']
    Shui_values_lst = basic_param_lst[4] # ['1.025', '2.0', '-5.0']
    toplevel = tb.Toplevel(root)
    toplevel.title("支架基本参数设置")
    toplevel.geometry("600x800")
    scrollbar = ScrolledFrame(toplevel, autohide=True, height=200)
    scrollbar.pack(fill=BOTH, expand=YES, padx=10, pady=10)
    SelfWeight = tb.DoubleVar(value=float(Self_values_lst[0]))
    FengGui_lst = ["自定义", "公路桥梁抗风设计规范"]
    ShuiGui_lst = ["自定义", "港口工程荷载规范"]
    Feng_values_dict = {"wind_entry1":'24.5', "wind_entry2":'1.0', "wind_entry3":'20', "wind_entry4":'50', "wind_entry5":'8', "wind_entry6":'1.3', "wind_combo1":'A:海面、海岸、开阔水面', "wind_combo2":'Ud = kf·kt·kh·U10'} # 风荷载参数词典
    Shui_values_dict = {"flow_entry1":'1.025', "flow_entry2":'2.0', "flow_entry3":'-5.0'} # 水流力参数词典
    if FengGui_selected_value != ['自定义']:
        for i,key in enumerate(Feng_values_dict.keys()):
            Feng_values_dict[key] = Feng_values_lst[i]
    if ShuiGui_selected_value != ['自定义']:
        for i,key in enumerate(Shui_values_dict.keys()):
            Shui_values_dict[key] = Shui_values_lst[i]
    Main_Headline(scrollbar, "基本参数定义：")
    container = tb.Frame(scrollbar)
    container.pack(fill=X, padx=20, pady=10)
    lbl = tb.Label(master=container, text='自重系数', width=12)
    lbl.pack(side=LEFT, padx=10)
    ent = tb.Entry(master=container, textvariable=SelfWeight, width=7)
    ent.pack(side=RIGHT, padx=10, fill=X)
    combo1, string_vars_dict1 = BoxLine2(scrollbar, "风荷载", FengGui_lst, create_Feng_dialog, FengGui_selected_value, Feng_values_dict)
    combo2, string_vars_dict2 = BoxLine2(scrollbar, "水流力", ShuiGui_lst, create_Shui_dialog, ShuiGui_selected_value, Shui_values_dict)
    # 创建保存按钮的函数
    def save_parameters():
        """保存所有参数到文件"""
        if combo1.get() == "自定义":
            Feng_values = ['无']
        else:
            # 获取所有参数值
            Feng_values = []
            for key in Feng_values_dict.keys():
                if key in string_vars_dict1:
                    value = string_vars_dict1[key].get()
                    Feng_values.append(value)
                    Feng_values_dict[key] = value  # 同时更新字典
                else:
                    Feng_values.append("")
        if combo2.get() == "自定义":
            Shui_values = ['无']
        else:
            Shui_values = []
            for key in Shui_values_dict.keys():
                if key in string_vars_dict2:
                    value = string_vars_dict2[key].get()
                    Shui_values.append(value)
                    Shui_values_dict[key] = value  # 同时更新字典
                else:
                    Shui_values.append("")
        # 构造保存的内容
        save_content = [[SelfWeight.get()], [combo1.get()], Feng_values, [combo2.get()], Shui_values]
        print(save_content)
        update_file(path, save_content)
    button_frame = tb.Frame(toplevel)
    button_frame.pack(padx=10, pady=10)
    # 添加取消/关闭按钮
    cancel_button = tb.Button(button_frame,text="关闭",command=toplevel.destroy,width=5)
    cancel_button.pack(side=RIGHT, padx=10)
    # 添加保存按钮
    save_button = tb.Button(button_frame, text="保存", command=save_parameters, bootstyle=SUCCESS, width=5)
    save_button.pack(side=RIGHT, padx=10)
    toplevel.mainloop()


# 平台基本参数设置
def Basic_Param_Set_PT(root, Applocation):
    path = os.path.join(Applocation, 'Support\\basic_param\\Basic_Param_Set_PT.txt')
    basic_param_lst = read_file(path)
    Self_values_lst = basic_param_lst[0] # -1.0
    FengGui_selected_value = basic_param_lst[1] # ["自定义"] # 选择规范默认值
    ShuiGui_selected_value = basic_param_lst[3] # ["自定义"] # 选择规范默认值
    Feng_values_lst = basic_param_lst[2] # ['24.5', '1.0', '20', '50', '8', '1.3', 'A:海面、海岸、开阔水面', 'Ud = kf·kt·kh·U10']
    Shui_values_lst = basic_param_lst[4] # ['1.025', '2.0', '-5.0']
    Dril_values_lst = basic_param_lst[5] # ['XR550D', '180', '6870', '1000', '5000', '5250']
    toplevel = tb.Toplevel(root)
    toplevel.title("支架基本参数设置")
    toplevel.geometry("600x800")
    scrollbar = ScrolledFrame(toplevel, autohide=True, height=200)
    scrollbar.pack(fill=BOTH, expand=YES, padx=10, pady=10)
    SelfWeight = tb.DoubleVar(value=float(Self_values_lst[0]))
    FengGui_lst = ["自定义", "公路桥梁抗风设计规范"]
    ShuiGui_lst = ["自定义", "港口工程荷载规范"]
    Feng_values_dict = {"wind_entry1":'24.5', "wind_entry2":'1.0', "wind_entry3":'20', "wind_entry4":'50', "wind_entry5":'8', "wind_entry6":'1.3', "wind_combo1":'A:海面、海岸、开阔水面', "wind_combo2":'Ud = kf·kt·kh·U10'} # 风荷载参数词典
    Shui_values_dict = {"flow_entry1":'1.025', "flow_entry2":'2.0', "flow_entry3":'-5.0'} # 水流力参数词典
    Dril_values_list = ['XR550D', '180', '6870', '1000', '5000', '5250']
    if FengGui_selected_value != ['自定义']:
        for i,key in enumerate(Feng_values_dict.keys()):
            Feng_values_dict[key] = Feng_values_lst[i]
    if ShuiGui_selected_value != ['自定义']:
        for i,key in enumerate(Shui_values_dict.keys()):
            Shui_values_dict[key] = Shui_values_lst[i]
    Main_Headline(scrollbar, "基本参数定义：")
    container = tb.Frame(scrollbar)
    container.pack(fill=X, padx=20, pady=10)
    lbl = tb.Label(master=container, text='自重系数', width=12)
    lbl.pack(side=LEFT, padx=10)
    ent = tb.Entry(master=container, textvariable=SelfWeight, width=7)
    ent.pack(side=RIGHT, padx=10, fill=X)
    combo1, string_vars_dict1 = BoxLine2(scrollbar, "风荷载", FengGui_lst, create_Feng_dialog, FengGui_selected_value, Feng_values_dict)
    combo2, string_vars_dict2 = BoxLine2(scrollbar, "水流力", ShuiGui_lst, create_Shui_dialog, ShuiGui_selected_value, Shui_values_dict)
    lablelst1 = ['钻机名称', '钻机工作重量(t)', '履带长度(mm)', '履带宽度(mm)', '履带轴距(mm)', '钻孔半径(mm)']
    variablelst1 = [
        tb.StringVar(value = Dril_values_lst[0]),
        tb.StringVar(value = Dril_values_lst[1]),
        tb.StringVar(value = Dril_values_lst[2]),
        tb.StringVar(value = Dril_values_lst[3]),
        tb.StringVar(value = Dril_values_lst[4]),
        tb.StringVar(value = Dril_values_lst[5]),
        ]
    BoxLine3(scrollbar, "钻机", lablelst1, variablelst1, 15)
    # 创建保存按钮的函数
    def save_parameters():
        """保存所有参数到文件"""
        if combo1.get() == "自定义":
            Feng_values = ['无']
        else:
            # 获取所有参数值
            Feng_values = []
            for key in Feng_values_dict.keys():
                if key in string_vars_dict1:
                    value = string_vars_dict1[key].get()
                    Feng_values.append(value)
                    Feng_values_dict[key] = value  # 同时更新字典
                else:
                    Feng_values.append("")
        if combo2.get() == "自定义":
            Shui_values = ['无']
        else:
            Shui_values = []
            for key in Shui_values_dict.keys():
                if key in string_vars_dict2:
                    value = string_vars_dict2[key].get()
                    Shui_values.append(value)
                    Shui_values_dict[key] = value  # 同时更新字典
                else:
                    Shui_values.append("")
        Dril_values = [x.get() for x in variablelst1]
        # 构造保存的内容
        save_content = [[SelfWeight.get()], [combo1.get()], Feng_values, [combo2.get()], Shui_values, Dril_values]
        print(save_content)
        update_file(path, save_content)
    button_frame = tb.Frame(toplevel)
    button_frame.pack(padx=10, pady=10)
    # 添加取消/关闭按钮
    cancel_button = tb.Button(button_frame,text="关闭",command=toplevel.destroy,width=5)
    cancel_button.pack(side=RIGHT, padx=10)
    # 添加保存按钮
    save_button = tb.Button(button_frame, text="保存", command=save_parameters, bootstyle=SUCCESS, width=5)
    save_button.pack(side=RIGHT, padx=10)
    toplevel.mainloop()


# 创建图标按钮（图标在上，加粗文字在下）
def create_icon_button(parent, image, text, command, pad=8, txt_height=9):
    """创建图标按钮：图标在上，加粗文字在下。pad 控制按钮内边距"""
    frame = tb.Frame(parent)
    # 图标按钮（无文字，通过 padding 控制大小）
    btn = tb.Button(master=frame, image=image, style="Custom.TButton", command=command)
    btn.configure(padding=pad)
    btn.pack(side=TOP, padx=5, pady=(2, 0))
    btn.image = image  # 保持图片引用
    # 加粗文字标签
    lbl = tb.Label(master=frame, text=text,
                   font=("Microsoft YaHei UI", txt_height, "bold"), foreground="#0d6efd")
    lbl.pack(side=TOP, padx=5, pady=(0, 2))
    return frame, btn


def create_label_entry(parent, label_text, row, default_val="", placeholder="",
                       label_width=25, entry_width=20, padx=5, pady=5):
    """标签+输入框 (grid布局)"""
    ttk.Label(parent, text=label_text, width=label_width, bootstyle=PRIMARY).grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
    entry = ttk.Entry(parent, width=entry_width)
    entry.grid(row=row, column=1, padx=padx, sticky="w")
    if placeholder:
        entry.insert(0, placeholder)
        entry.config(foreground="gray")
    else:
        if default_val:
            entry.insert(0, default_val)
    return entry


def make_card_frame(parent, title=None):
    outer = ttk.Frame(parent)
    outer.pack(fill=X, padx=12, pady=6)
    if title:
        lbl = ttk.Label(outer, text=title, font=("Microsoft YaHei UI", 10, "bold"),
                        bootstyle=PRIMARY)
        lbl.pack(anchor=W, padx=18, pady=(10, 5))
    content = ttk.Frame(outer)
    content.pack(fill=X, padx=18, pady=(0, 5))
    return outer, content


def make_form_row(parent, label_text, widget, label_width=18, row_pady=4):
    """标签+控件行"""
    row_frame = ttk.Frame(parent)
    row_frame.pack(fill=X, pady=row_pady, padx=5)
    lbl = ttk.Label(row_frame, text=label_text, width=label_width, bootstyle=PRIMARY)
    lbl.pack(side=LEFT, padx=(10, 5))
    widget.pack(side=LEFT, padx=5)
    return row_frame


class SplitEntry:
    """全局通用：双 Entry 组合框（无 placeholder，直接输入）"""
    def __init__(self, parent, row, default_val="", placeholders=("起始间距", "布置间距"),
                 entry_width=40, padx=(0, 5), pady=3, on_focus=None):
        self.inner_frame = ttk.Frame(parent)
        self.inner_frame.grid(row=row, column=1, sticky="w", padx=padx, pady=pady)

        w1 = (entry_width - 3) // 2
        w2 = entry_width - 3 - w1

        self.e1 = ttk.Entry(self.inner_frame, width=w1)
        self.e1.pack(side="left")
        ttk.Label(self.inner_frame, text=" , ").pack(side="left", padx=1)
        self.e2 = ttk.Entry(self.inner_frame, width=w2)
        self.e2.pack(side="left")

        self.placeholders = placeholders
        parts = str(default_val).split(",") if default_val else []
        v1 = parts[0].strip() if len(parts) > 0 else ""
        v2 = parts[1].strip() if len(parts) > 1 else ""

        if v1:
            self.e1.insert(0, v1)
        if v2:
            self.e2.insert(0, v2)

        # 外部高亮回调（_on_diagram_focus）
        if on_focus:
            self.e1.bind("<FocusIn>", lambda e: on_focus("left", True), add="+")
            self.e1.bind("<FocusOut>", lambda e: on_focus("left", False), add="+")
            self.e2.bind("<FocusIn>", lambda e: on_focus("right", True), add="+")
            self.e2.bind("<FocusOut>", lambda e: on_focus("right", False), add="+")

    @property
    def entry1(self): return self.e1
    @property
    def entry2(self): return self.e2

    def get(self):
        """外部调用获取组合值：英文逗号组装，过滤两端空白"""
        val1 = self.e1.get().strip()
        val2 = self.e2.get().strip()
        return f"{val1},{val2}" if val2 else val1

    def bind(self, sequence, func, add="+"):
        """重写 bind 方法，默认追加事件(add='+')"""
        self.e1.bind(sequence, func, add=add)
        self.e2.bind(sequence, func, add=add)

    def configure(self, **kwargs):
        self.e1.configure(**kwargs)
        self.e2.configure(**kwargs)


def create_scrollable_frame(parent, pack_kwargs=None):
    """创建带滚动条的 Canvas + 内部 Frame，返回 (canvas, inner_frame)。

    使用 pack 布局，支持鼠标滚轮。调用方通过返回的 inner_frame 添加内容。
    """
    scrollbar = ttk.Scrollbar(parent, orient="vertical")
    scrollbar.pack(side="right", fill="y")
    canvas = tk.Canvas(parent, borderwidth=0, highlightthickness=0)
    canvas.pack(side="left", fill="both", expand=True, **(pack_kwargs or {}))
    canvas.configure(yscrollcommand=scrollbar.set)
    scrollbar.config(command=canvas.yview)

    inner = tk.Frame(canvas)
    inner_id = canvas.create_window((0, 0), window=inner, anchor="nw")
    canvas.bind("<Configure>", lambda e: canvas.itemconfig(inner_id, width=e.width))
    inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    def _on_mousewheel(event):
        widget = event.widget
        depth = 0
        while widget is not None and depth < 20:
            if isinstance(widget, tk.Canvas):
                try:
                    widget.yview_scroll(int(-1 * (event.delta / 120)), "units")
                except Exception:
                    pass
                return
            try:
                widget = widget.master
            except AttributeError:
                return
            depth += 1
    canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
    canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

    return canvas, inner

# ═══════════════════════════════════════════════════════════════════
# PileBoundaryFrame — 桩底边界定义（D/R 六向约束 checkbox 组）
# ═══════════════════════════════════════════════════════════════════

class PileBoundaryFrame(ttk.Frame):
    """桩底边界定义 — D(Dx/Dy/Dz) + R(Rx/Ry/Rz) 六向约束 checkbox 组。

    支持两种使用模式：
    - 嵌入模式：直接 pack/grid 到父容器
    - 二级窗口模式：调用 open_dialog() 类方法

    参数:
        parent: 父容器
        boundary_var: tk.StringVar，6位字符串如 "111111"，前3位平动D，后3位转动R
        on_change: fn() -> None，值变化回调（如联动图示刷新）
    """

    def __init__(self, parent, boundary_var=None, on_change=None):
        super().__init__(parent)
        self._boundary_var = boundary_var or tk.StringVar(value="111111")
        self._on_change = on_change
        self._setup_ui()
        self._refresh_from_var()
        self._boundary_var.trace_add("write", lambda *_: self._refresh_from_var())

    def _setup_ui(self):
        boundary_str = self._boundary_var.get()
        if len(boundary_str) < 6:
            boundary_str = "111111"

        # 注册加粗样式
        style = tb.Style()
        style.configure("Bold.primary.TCheckbutton", font=("Microsoft YaHei UI", 9, "bold"))

        # D 组（左侧）
        d_frame = tb.Frame(self)
        d_frame.pack(side=LEFT, fill=BOTH, expand=True, pady=10, padx=10)

        self._d_vars = []
        self._d_all_var = tb.BooleanVar(value=(boundary_str[:3] == "111"))

        tb.Checkbutton(d_frame, text="D-ALL", variable=self._d_all_var,
                       style="Bold.primary.TCheckbutton",
                       command=self._toggle_d_all).pack(anchor=W, pady=(0, 5), padx=20)

        d_subframe = tb.Frame(d_frame)
        d_subframe.pack(anchor=W, fill=X)

        for d_idx, d_name in enumerate(["Dx", "Dy", "Dz"]):
            dv = tb.BooleanVar(value=(boundary_str[d_idx] == '1'))
            self._d_vars.append(dv)
            tb.Checkbutton(d_subframe, text=d_name, variable=dv, bootstyle="primary",
                           command=lambda didx=d_idx: self._on_d_single(didx)).pack(
                               side=LEFT, pady=10, padx=(20, 20))

        # 垂直分隔线
        tb.Separator(self, orient=VERTICAL).pack(side=LEFT, fill=Y, padx=15, pady=5)

        # R 组（右侧）
        r_frame = tb.Frame(self)
        r_frame.pack(side=LEFT, fill=BOTH, expand=True, pady=10, padx=10)

        self._r_vars = []
        self._r_all_var = tb.BooleanVar(value=(boundary_str[3:] == "111"))

        tb.Checkbutton(r_frame, text="R-ALL", variable=self._r_all_var,
                       style="Bold.primary.TCheckbutton",
                       command=self._toggle_r_all).pack(anchor=W, pady=(0, 5), padx=10)

        r_subframe = tb.Frame(r_frame)
        r_subframe.pack(anchor=W, fill=X)

        for r_idx, r_name in enumerate(["Rx", "Ry", "Rz"]):
            rv = tb.BooleanVar(value=(boundary_str[r_idx + 3] == '1'))
            self._r_vars.append(rv)
            tb.Checkbutton(r_subframe, text=r_name, variable=rv, bootstyle="primary",
                           command=lambda ridx=r_idx: self._on_r_single(ridx)).pack(
                               side=LEFT, pady=10, padx=(10, 20))

    def _toggle_d_all(self):
        val = self._d_all_var.get()
        for dv in self._d_vars:
            dv.set(val)
        self._update_boundary_str()

    def _on_d_single(self, didx):
        self._d_all_var.set(all(dv.get() for dv in self._d_vars))
        self._update_boundary_str()

    def _toggle_r_all(self):
        val = self._r_all_var.get()
        for rv in self._r_vars:
            rv.set(val)
        self._update_boundary_str()

    def _on_r_single(self, ridx):
        self._r_all_var.set(all(rv.get() for rv in self._r_vars))
        self._update_boundary_str()

    def _update_boundary_str(self):
        s = "".join(["1" if v.get() else "0" for v in self._d_vars + self._r_vars])
        self._boundary_var.set(s)
        if self._on_change:
            self._on_change()

    def _refresh_from_var(self):
        bs = self._boundary_var.get()
        if len(bs) < 6:
            bs = "111111"
        self._d_all_var.set(bs[:3] == "111")
        for idx, dv in enumerate(self._d_vars):
            dv.set(bs[idx] == "1")
        self._r_all_var.set(bs[3:] == "111")
        for idx, rv in enumerate(self._r_vars):
            rv.set(bs[idx + 3] == "1")

    @classmethod
    def open_dialog(cls, parent, boundary_var=None, title="桩底边界设置"):
        """二级窗口模式。返回 {"action":"save","boundary":"111111"} 或 None"""
        dialog = tb.Toplevel(parent)
        dialog.title(title)
        dialog.geometry("600x200")
        dialog.transient(parent)
        dialog.grab_set()

        var = tk.StringVar(value=boundary_var.get() if boundary_var else "111111")
        frame = cls(dialog, boundary_var=var)
        frame.pack(fill=BOTH, expand=True, padx=10, pady=10)

        result = [None]

        def on_ok():
            if boundary_var:
                boundary_var.set(var.get())
            result[0] = {"action": "save", "boundary": var.get()}
            dialog.destroy()

        btn_row = tb.Frame(dialog)
        btn_row.pack(fill=X, padx=10, pady=5)
        tb.Button(btn_row, text="确定", bootstyle=SUCCESS, command=on_ok).pack(side=RIGHT, padx=5)
        tb.Button(btn_row, text="取消", bootstyle=SECONDARY, command=dialog.destroy).pack(side=RIGHT, padx=5)

        dialog.wait_window()
        return result[0]


# ═══════════════════════════════════════════════════════════════════
# ConnectionSettingFrame — 单个连接项设置（类型选择 + 刚度参数）
# ═══════════════════════════════════════════════════════════════════

class ConnectionSettingFrame(ttk.Frame):
    """单个弹性连接项的设置UI。

    封装一个连接项的：连接类型下拉 + 自定义名称 + 刚度参数区。
    不包含多连接项的按钮循环，多连接项的组织由调用方负责。

    支持两种使用模式：
    - 嵌入模式：直接 pack/grid 到父容器
    - 二级窗口模式：调用 open_dialog() 类方法

    参数:
        parent: 父容器
        preset_type: 预设连接类型名，如 "共节点" / "铰接1"
        stiffness: 预设刚度 dict {SDx, SDy, SDz, SRx, SRy, SRz, NSDx}
        conn_options: 下拉可选项列表（不含"自定义"/"共节点"，由类内部自动补充）
        node_only: True=仅支持共节点，显示label而非下拉
        editable: 是否允许修改连接类型（False时下拉disabled）
        on_change: fn() -> None，值变化回调
    """

    STIFFNESS_KEYS = ["SDx", "SDy", "SDz", "SRx", "SRy", "SRz", "NSDx"]
    STIFFNESS_LABELS = [
        ("SDx(N/mm)", 0), ("SDy(N/mm)", 1), ("SDz(N/mm)", 2),
        ("SRx(N·mm/rad)", 3), ("SRy(N·mm/rad)", 4), ("SRz(N·mm/rad)", 5),
    ]

    def __init__(self, parent, preset_type=None, stiffness=None,
                 conn_options=None, node_only=False, editable=True,
                 on_change=None):
        super().__init__(parent)
        self._node_only = node_only
        self._editable = editable
        self._on_change = on_change
        self._conns = []  # 连接数据列表（含 name + 刚度值）
        self._conn_options = list(conn_options) if conn_options else []
        self._original_type = preset_type or ("共节点" if node_only else "自定义")
        self._original_stiffness = dict(stiffness) if stiffness else {}

        # StringVar 初始化
        self._type_var = tk.StringVar(value=self._original_type)
        self._custom_name_var = tk.StringVar()
        self._stiff_vars = [tk.StringVar(value="") for _ in range(7)]

        self._setup_ui()
        self._load_preset()

    def _build_conn_names(self):
        """构建连接类型下拉列表：自定义 → conn_options"""
        names = ["自定义"]
        for name in self._conn_options:
            if name and name not in names:
                names.append(name)
        return names

    def _setup_ui(self):
        if self._node_only:
            # 仅共节点模式：显示 label
            row = tb.Frame(self)
            row.pack(fill=X, pady=2, padx=5)
            tb.Label(row, text="连接类型:", width=12).pack(side=LEFT, padx=(5, 2))
            tb.Label(row, text="共节点", font=("Microsoft YaHei UI", 10, "bold"),
                     bootstyle=PRIMARY).pack(side=LEFT, padx=5)
            return

        # 完整模式：类型选择 + 刚度参数
        # 连接类型行
        type_row = tb.Frame(self)
        type_row.pack(fill=X, pady=2, padx=5)

        tb.Label(type_row, text="连接类型:", width=12).pack(side=LEFT, padx=(5, 2))
        self._type_combo = tb.Combobox(type_row, textvariable=self._type_var,
                                        values=self._build_conn_names(),
                                        state='readonly' if self._editable else 'disabled',
                                        width=22)
        self._type_combo.pack(side=LEFT, padx=2)

        # 自定义名称区（仅选"自定义"时显示）
        self._name_frame = tb.Frame(type_row)
        tb.Label(self._name_frame, text="自定义名称:", width=10).pack(side=LEFT, padx=(5, 2))
        self._custom_name_entry = tb.Entry(self._name_frame, textvariable=self._custom_name_var, width=18)
        self._custom_name_entry.pack(side=LEFT, padx=2)
        self._save_conn_btn = tb.Button(self._name_frame, text="保存", bootstyle=SUCCESS, width=8,
                                         command=self._save_custom_conn)

        # 刚度参数区
        self._stiffness_container = tb.Frame(self)
        self._stiffness_container.pack(fill=X, pady=3, padx=5)

        stiff_frame = tb.Frame(self._stiffness_container)
        stiff_frame.pack(fill=X, pady=3)

        left_frame = tb.Frame(stiff_frame)
        left_frame.pack(side=LEFT, fill=X, expand=True)
        self._stiffness_entries = []

        trans_frame = tb.Frame(left_frame)
        trans_frame.pack(fill=X, pady=(3, 5))
        for var_idx, (label, idx) in enumerate(self.STIFFNESS_LABELS[:3]):
            tb.Label(trans_frame, text=label, bootstyle=PRIMARY, width=14).pack(side=LEFT, padx=3)
            e = tb.Entry(trans_frame, textvariable=self._stiff_vars[idx], width=10)
            e.pack(side=LEFT, padx=3)
            self._stiffness_entries.append(e)

        rot_frame = tb.Frame(left_frame)
        rot_frame.pack(fill=X, pady=3)
        for var_idx, (label, idx) in enumerate(self.STIFFNESS_LABELS[3:6]):
            tb.Label(rot_frame, text=label, bootstyle=PRIMARY, width=14).pack(side=LEFT, padx=3)
            e = tb.Entry(rot_frame, textvariable=self._stiff_vars[idx], width=10)
            e.pack(side=LEFT, padx=3)
            self._stiffness_entries.append(e)

        tb.Separator(stiff_frame, orient=VERTICAL).pack(side=LEFT, fill=Y, padx=8)

        right_frame = tb.Frame(stiff_frame)
        right_frame.pack(side=LEFT, fill=X, expand=True)
        nsdx_frame = tb.Frame(right_frame)
        nsdx_frame.pack(fill=X, pady=1)
        tb.Label(nsdx_frame, text="NSDx(N/mm)", bootstyle=PRIMARY, width=14).pack(side=LEFT, padx=3)
        nsdx_entry = tb.Entry(nsdx_frame, textvariable=self._stiff_vars[6], width=10)
        nsdx_entry.pack(side=LEFT, padx=3)
        self._stiffness_entries.append(nsdx_entry)

        self._format_hint = tb.Label(self._stiffness_container,
                                      text="* 弹性连接参数格式: 100000 或 1e5",
                                      font=("Microsoft YaHei UI", 8), foreground="#6b7280")
        self._format_hint.pack(anchor=W, padx=10, pady=2)

        # 联动
        self._type_var.trace_add("write", self._on_type_changed)
        self._toggle_custom_conn()

    def _load_preset(self):
        """从预设值加载"""
        t = self._original_type
        if t == "共节点" or self._node_only:
            return
        if t and t != "自定义":
            for vi, key in enumerate(self.STIFFNESS_KEYS):
                self._stiff_vars[vi].set(str(self._original_stiffness.get(key, "")))
        elif t == "自定义":
            for vi, key in enumerate(self.STIFFNESS_KEYS):
                val = self._original_stiffness.get(key, "/")
                self._stiff_vars[vi].set(str(val) if val != "/" else "")

    def _toggle_custom_conn(self, *_):
        if self._node_only:
            return
        if self._type_var.get() == "自定义":
            self._name_frame.pack(side=LEFT, fill=X, expand=True)
            self._save_conn_btn.pack(side=RIGHT, padx=5)
        else:
            self._name_frame.pack_forget()
            self._save_conn_btn.pack_forget()

    def _on_type_changed(self, *_):
        if self._node_only:
            return
        name = self._type_var.get()
        is_custom = (name == "自定义")

        # 显隐刚度区
        if name == "共节点":
            self._stiffness_container.pack_forget()
        else:
            self._stiffness_container.pack(fill=X, pady=3, padx=5)

        # 加载刚度值
        if name not in ("自定义", "共节点"):
            for c in self._conns:
                if c.get("name") == name:
                    for vi, key in enumerate(self.STIFFNESS_KEYS):
                        self._stiff_vars[vi].set(str(c.get(key, "")))
                    break
        elif name == "共节点":
            for v in self._stiff_vars:
                v.set("")

        # Entry 编辑状态
        for w in self._stiffness_entries:
            try:
                w.configure(state='normal' if is_custom else 'readonly')
            except Exception:
                pass

        # 自定义名称区显隐
        self._toggle_custom_conn()
        if self._on_change:
            self._on_change()

    def _save_custom_conn(self):
        name = self._custom_name_var.get().strip()
        if not name:
            messagebox.showwarning("提示", "连接类型名称不能为空", parent=self)
            return
        if name in self._build_conn_names():
            messagebox.showwarning("提示", "该名称已存在", parent=self)
            return
        vals = [v.get().strip() for v in self._stiff_vars]
        if not all(vals[:6]):
            messagebox.showwarning("提示", "请填写全部刚度参数", parent=self)
            return
        new_conn = {"name": name}
        for vi, key in enumerate(self.STIFFNESS_KEYS):
            new_conn[key] = vals[vi]
        self._conns.append(new_conn)
        self._conn_options.append(name)
        self._type_combo.configure(values=self._build_conn_names())
        self._type_var.set(name)

    # ===== 对外接口 =====

    def get_type(self):
        """获取当前连接类型名"""
        return self._type_var.get()

    def get_stiffness(self):
        """获取当前刚度值 dict"""
        if self._node_only or self._type_var.get() in ("共节点", ""):
            return {}
        result = {}
        for vi, key in enumerate(self.STIFFNESS_KEYS):
            val = self._stiff_vars[vi].get().strip()
            if val:
                result[key] = val
        return result

    def set_type(self, type_name):
        """设置连接类型（项目切换时调用）"""
        if self._node_only:
            return
        self._type_var.set(type_name if type_name else "自定义")

    def set_stiffness(self, stiffness_dict):
        """设置刚度值（项目切换时调用）"""
        if self._node_only:
            return
        for vi, key in enumerate(self.STIFFNESS_KEYS):
            val = stiffness_dict.get(key, "")
            self._stiff_vars[vi].set(str(val) if val else "")

    def set_conn_options(self, options):
        """更新下拉可选项列表"""
        self._conn_options = list(options) if options else []
        if not self._node_only:
            self._type_combo.configure(values=self._build_conn_names())

    def set_conns_data(self, conns_list):
        """设置完整连接数据列表（含 name + 刚度值），用于类型切换时自动加载"""
        self._conns = list(conns_list) if conns_list else []

    @classmethod
    def open_dialog(cls, parent, preset_type=None, stiffness=None,
                    conn_options=None, node_only=False, editable=True,
                    title="连接设置"):
        """二级窗口模式。返回 {"action":"save","type":...,"stiffness":{...}} 或 None"""
        dialog = tb.Toplevel(parent)
        dialog.title(title)
        dialog.geometry("550x300")
        dialog.transient(parent)
        dialog.grab_set()

        result = [None]
        frame = cls(dialog, preset_type=preset_type, stiffness=stiffness,
                    conn_options=conn_options, node_only=node_only, editable=editable)
        frame.pack(fill=BOTH, expand=True, padx=10, pady=10)

        def on_ok():
            result[0] = {
                "action": "save",
                "type": frame.get_type(),
                "stiffness": frame.get_stiffness(),
            }
            dialog.destroy()

        btn_row = tb.Frame(dialog)
        btn_row.pack(fill=X, padx=10, pady=5)
        tb.Button(btn_row, text="确定", bootstyle=SUCCESS, command=on_ok).pack(side=RIGHT, padx=5)
        tb.Button(btn_row, text="取消", bootstyle=SECONDARY, command=dialog.destroy).pack(side=RIGHT, padx=5)

        dialog.wait_window()
        return result[0]
    

# ═══════════════════════════════════════════════════════════════════
# BeamReleaseFrame — 梁端约束释放设置（2行×7列 checkbox + 预设按钮）
# ═══════════════════════════════════════════════════════════════════

class BeamReleaseFrame(ttk.Frame):
    """梁端约束释放设置 — i/j 节点各7自由度 checkbox，支持预设按钮。

    数据格式: "0000000,0000100" (i端,j端，1=释放/自由，0=约束)

    参数:
        parent: 父容器
        release_var: tk.StringVar，格式 "iiiiiii,jjjjjjj"
        on_change: fn() -> None，值变化回调
    """

    COL_LABELS = ["Fx", "Fy", "Fz", "Mx", "My", "Mz", "Mb"]
    COL_WIDTH = 6

    def __init__(self, parent, release_var=None, on_change=None):
        super().__init__(parent)
        self._release_var = release_var or tk.StringVar(value="0000000,0000000")
        self._on_change = on_change
        self._check_vars = [[], []]
        self._setup_ui()
        self._refresh_from_var()
        self._release_var.trace_add("write", lambda *_: self._refresh_from_var())

    def _setup_ui(self):
        # ── 表格区域  ──
        table_frame = ttk.Frame(self)
        table_frame.pack(fill="x", pady=2)
        # 配置表格列的权重和统一宽度
        table_frame.columnconfigure(0, minsize=60)
        for ci in range(7):
            table_frame.columnconfigure(ci + 1, weight=1, uniform="br_col")
        # ── 列标题 (Grid 第 0 行) ──
        ttk.Label(table_frame, text="", width=8).grid(row=0, column=0)
        for ci, label in enumerate(self.COL_LABELS):
            ttk.Label(table_frame, text=label, anchor="center",
                      font=("Microsoft YaHei UI", 9, "bold"),
                      bootstyle=PRIMARY).grid(row=0, column=ci + 1, padx=2, pady=(0, 5), sticky="ew")
        # ── 数据行 (Grid 第 1, 2 行) ──
        # 传入 table_frame 作为父容器，并指定在网格中的行号
        self._build_row(table_frame, "i-节点", data_idx=0, grid_row=1)
        self._build_row(table_frame, "j-节点", data_idx=1, grid_row=2)

        # ── 预设按钮行 ──
        preset_frame = ttk.Frame(self)
        preset_frame.pack(fill="x", pady=(10, 2))
        for text, cmd in [
            ("铰-铰",   lambda: self._apply_preset(i_released=[4,5], j_released=[4,5])),
            ("铰-刚接", lambda: self._apply_preset(i_released=[4,5], j_released=[])),
            ("刚接-铰", lambda: self._apply_preset(i_released=[], j_released=[4,5])),
            ("刚接-刚接", lambda: self._apply_preset(i_released=[], j_released=[])),
        ]:
            ttk.Button(preset_frame, text=text, bootstyle="primary-outline",
                       command=cmd).pack(side="left", padx=2, fill="x", expand=True)

    def _build_row(self, parent_frame, label, data_idx, grid_row):
        # 放置行首文本 (居右对齐)
        ttk.Label(parent_frame, text=label, anchor="e",
                  font=("Microsoft YaHei UI", 9)).grid(row=grid_row, column=0, padx=(0, 5), pady=3, sticky="e")
        
        # 放置复选框 (默认居中对齐)
        for col in range(7):
            var = tk.BooleanVar(value=False)
            self._check_vars[data_idx].append(var)
            cb = ttk.Checkbutton(parent_frame, variable=var, bootstyle="primary",
                                 command=self._update_release_str)
            cb.grid(row=grid_row, column=col + 1, padx=(4, 0), pady=3)

    def _apply_preset(self, i_released, j_released):
        """应用预设：i_released/j_released 为需要释放（置1）的列索引"""
        for col in range(7):
            self._check_vars[0][col].set(col in i_released)
            self._check_vars[1][col].set(col in j_released)
        self._update_release_str()

    def _update_release_str(self):
        i_str = "".join(["1" if v.get() else "0" for v in self._check_vars[0]])
        j_str = "".join(["1" if v.get() else "0" for v in self._check_vars[1]])
        self._release_var.set(f"{i_str},{j_str}")
        if self._on_change:
            self._on_change()

    def _refresh_from_var(self):
        raw = self._release_var.get()
        parts = raw.split(",")
        if len(parts) >= 2 and len(parts[0]) == 7 and len(parts[1]) == 7:
            i_str, j_str = parts[0], parts[1]
        else:
            i_str = j_str = "0000000"
        for col in range(7):
            self._check_vars[0][col].set(i_str[col] == "1" if col < len(i_str) else False)
            self._check_vars[1][col].set(j_str[col] == "1" if col < len(j_str) else False)


# ====================================================================
# 通用土层参数编辑组件（tksheet）
# ====================================================================
class SoilSettingsFrame(ttk.Frame):
    """通用土层参数编辑框架 — 支持子frame嵌入和二级窗口模式

    == 三层架构 ==
    第1层（控件层）: tksheet + 按钮 + "选择已有"参数来源区
    第2层（数据层）: load_data / get_data / has_changes
    第3层（业务层）: 由调用方通过 on_save 回调实现

    用法:
        # 作为子frame嵌入（无按钮，由父容器控制）
        frame = SoilSettingsFrame(parent, data=soil_dict, show_buttons=False)
        frame.pack(fill="both", expand=True)
        ...
        result = frame.get_data()         # → {name: [vals]}
        if frame.has_changes(): ...

        # 作为二级窗口（阻塞式）
        result = SoilSettingsFrame.open_dialog(parent, data=soil_dict,
                    title="1# 土层参数设置", on_save=my_callback)
    """
    HEADERS = ["土层名称", "土层性质", "层厚(m)", "重度(kN/m3)",
               "黏聚力c(kPa)", "内摩擦角φ(°)", "qsik(kPa)", "qpk(kPa)"]

    # ── 对话框模式 ────────────────────────────────
    @classmethod
    def open_dialog(cls, parent, data=None, headers=None, title="土层参数设置",
                    row_count=20, existing_options=None, on_save=None,
                    dialog_width=None, dialog_height=500, col_widths=None, column_dropdowns=None, on_existing_selected=None):
        """以对话框形式打开（阻塞，返回保存结果）

        Args:
            parent: 父窗口
            data: dict {name: [vals]} 或 list[list]
            headers: 表头列表
            title: 窗口标题
            existing_options: 已有土层参数名称列表（供"选择已有"下拉框）
            on_save: 保存回调 fn(result_dict) → None
        Returns:
            {"action":"save", "data":{name:[vals]}} 或
            {"action":"reference", "target":"已有土层名"} 或
            None（取消）
        """
        win = tb.Toplevel(parent)
        win.title(title)
        if dialog_width:
            dw = dialog_width
        else:
            pw = parent.winfo_width() if parent.winfo_width() > 100 else 900
            dw = max(900, pw)
        win.geometry(f"{dw}x{dialog_height}+{parent.winfo_rootx()+50}+{parent.winfo_rooty()+50}")
        win.resizable(True, True)
        win.transient(parent)
        win.grab_set()

        result_holder = []
        def _on_save(res):
            result_holder.append(res)
            if on_save:
                on_save(res)
            win.destroy()

        frame = cls(win, data=data, headers=headers, row_count=row_count,
                    existing_options=existing_options, on_save=_on_save,
                    col_widths=col_widths, column_dropdowns=column_dropdowns,
                    on_existing_selected=on_existing_selected,
                    show_buttons=True)
        frame.pack(fill="both", expand=True)
        win.protocol("WM_DELETE_WINDOW", win.destroy)

        # 取消按钮（对话框独有）
        btn_frame = frame.btn_frame
        tb.Button(btn_frame, text="取消", bootstyle="secondary", width=10,
                  command=win.destroy).pack(side="right", padx=5)

        parent.wait_window(win)
        return result_holder[0] if result_holder else None

    # ── 初始化 ────────────────────────────────────
    def __init__(self, parent, data=None, headers=None, row_count=20,
                 existing_options=None, on_save=None, show_buttons=True, height=None,
                 col_widths=None, column_dropdowns=None, on_existing_selected=None):
        """height: 内嵌模式下的固定高度（px），不设则自动撑满
        col_widths: 每列宽度列表，长度需与 headers 一致，不设则统一 100
        column_dropdowns: {列索引: [选项列表]}，如 {6: ["水土合算", "水土分算"]}
        on_existing_selected: 选择已有参数时的回调 fn(sheet_name)"""
        super().__init__(parent)
        self._headers = headers or list(self.HEADERS)
        self._col_count = len(self._headers)
        self._col_widths = col_widths
        self._column_dropdowns = column_dropdowns or {}
        self._row_count = row_count
        self._existing_options = existing_options or []
        self._on_existing_selected_cb = on_existing_selected
        self._on_save = on_save
        self._show_buttons = show_buttons
        self._fixed_height = height
        self.result = None
        self._original_data = None

        self._setup_ui()
        self.load_data(data)

    # ── 第1层：控件布局 ───────────────────────────

    def _setup_ui(self):
        import tksheet

        # 1. 参数来源区（参考已有土层参数）
        self._setup_source_frame()

        # 分隔线
        ttk.Separator(self, orient=HORIZONTAL).pack(fill=X, padx=15, pady=(2, 0))

        # 2. tksheet（外框包裹）
        sheet_border = tk.Frame(self, bg="#cbd5e1", highlightthickness=0)
        if self._fixed_height:
            sheet_border.pack(fill=X, padx=15, pady=(6, 5))
            sheet_border.configure(height=self._fixed_height)
            sheet_border.pack_propagate(False)
        else:
            sheet_border.pack(fill=BOTH, expand=True, padx=15, pady=(6, 5))
        sheet_inner = tk.Frame(sheet_border, bg="white")
        sheet_inner.pack(fill=BOTH, expand=True, padx=1, pady=1)

        self.sheet = tksheet.Sheet(
            sheet_inner,
            data=[["" for _ in range(self._col_count)] for _ in range(self._row_count)],
            show_row_index=True,
            row_index_width=30,
            header_height=30,
            row_height=30,
            header_font=("Microsoft YaHei", 9, "bold"),
            font=("Microsoft YaHei", 9, "normal"),
            theme="light",
            header_bg="#f8fafc",
            header_fg="#1e293b",
            grid_color="#e2e8f0",
        )
        self.sheet.pack(fill="both", expand=True)
        self.sheet.headers(self._headers)
        self.sheet.enable_bindings(
            "single_select", "drag_select", "row_select", "column_select",
            "arrowkeys", "right_click_popup_menu", "rc_select", "rc_insert_row",
            "rc_delete_row", "copy", "cut", "paste", "undo", "edit_cell"
        )
        self.sheet.MT.bind("<MouseWheel>", self._on_mouse_wheel)

        # 分隔线
        ttk.Separator(self, orient=HORIZONTAL).pack(fill=X, padx=15, pady=(2, 0))
        
        # 3. 底部按钮
        self.btn_frame = ttk.Frame(self)
        self.btn_frame.pack(side="bottom", fill="x", padx=10, pady=8)
        if self._show_buttons:
            tb.Button(self.btn_frame, text="保存", bootstyle="success", width=10,
                      command=self._on_save_clicked).pack(side="right", padx=5)

    def _on_mouse_wheel(self, event):
        # 允许滚动，但在滚动后调用 refresh
        self.sheet.refresh()

    def _setup_source_frame(self):
        """参考已有土层参数的 UI（无边框，直接放置在界面内）"""
        self.top_frame = ttk.Frame(self)
        self.top_frame.pack(fill="x", padx=10, pady=(8, 2))
        self.use_existing = tk.BooleanVar(value=False)
        cb = tb.Checkbutton(self.top_frame, text="参考已有土层参数",
                            variable=self.use_existing,
                            command=self._toggle_existing, bootstyle="round-toggle")
        cb.pack(side="left")
        self.existing_combo = ttk.Combobox(self.top_frame, state="readonly", width=30)
        self.existing_var = tb.StringVar(value="")
        self.existing_combo.configure(textvariable=self.existing_var)
        self.existing_combo.pack(side="left", padx=(10, 0))
        self.existing_combo.bind("<<ComboboxSelected>>", self._on_existing_selected)
        if self._existing_options:
            self.existing_combo['values'] = self._existing_options
        self.existing_combo.pack_forget()  # 默认隐藏

    def _toggle_existing(self):
        if self.use_existing.get():
            self.existing_combo.pack(side="left", padx=(10, 0))
        else:
            self.existing_combo.pack_forget()

    def _on_existing_selected(self, event):
        """参考已有土层参数时触发——调用外部回调"""
        if self._on_existing_selected_cb:
            sheet = self.existing_combo.get()
            if sheet:
                self._on_existing_selected_cb(sheet)
    def load_data(self, data):
        """加载数据到 tksheet。支持 dict 或 list[list] 格式

        Args:
            data: dict {layer_name: [val1, val2, ...]}
                  或 list[list] [[name, val1, val2, ...], ...]
                  或 None（清空）
        """
        if data is None:
            rows = [["" for _ in range(self._col_count)] for _ in range(self._row_count)]
        elif isinstance(data, dict):
            rows = []
            for name, vals in data.items():
                row = [name] + list(vals)
                while len(row) < self._col_count:
                    row.append("")
                rows.append(row)
            while len(rows) < self._row_count:
                rows.append(["" for _ in range(self._col_count)])
        elif isinstance(data, list):
            rows = [list(row) for row in data]
            while len(rows) < self._row_count:
                rows.append(["" for _ in range(self._col_count)])
        else:
            rows = [["" for _ in range(self._col_count)] for _ in range(self._row_count)]

        # 先应用列布局（含下拉），再设数据，避免 create_dropdown 覆盖
        self._original_data = [list(r) for r in rows]
        self.sheet.set_sheet_data(rows)
        self._original_data = [list(r) for r in rows]
        self.after(50, self._apply_column_layout)
    def get_data(self):
        """获取编辑后的数据 → dict {name: [vals]}（已过滤空行）"""
        data = self.sheet.get_sheet_data()
        result = {}
        for row in data:
            name = str(row[0]).strip() if row[0] not in (None, "") else ""
            vals = [str(v).strip() if v not in (None, "") else ""
                    for v in row[1:self._col_count]]
            if not name and not any(v != "" for v in vals):
                continue
            result[name] = vals
        return result

    def has_changes(self):
        """检查数据是否被修改"""
        if self._original_data is None:
            return True
        return self.sheet.get_sheet_data() != self._original_data

    def set_existing_options(self, options):
        """设置已有土层参数下拉选项"""
        self._existing_options = options
        self.existing_combo['values'] = options
        self.existing_var.set("")

    def _apply_column_layout(self):
        self.sheet.align_columns(columns=list(range(self._col_count)), align="center")
        for idx in range(self._col_count):
            w = self._col_widths[idx] if self._col_widths and idx < len(self._col_widths) else 100
            self.sheet.column_width(column=idx, width=w)
        # 列下拉选项（先存值再建下拉，避免 create_dropdown 覆盖数据）
        for col_idx, options in self._column_dropdowns.items():
            if col_idx < self._col_count:
                saved = [self.sheet.get_cell_data(r, col_idx) for r in range(len(self.sheet.get_sheet_data()))]
                self.sheet.create_dropdown("all", col_idx, values=options)
                for r, val in enumerate(saved):
                    self.sheet.set_cell_data(r, col_idx, val)

    def _on_save_clicked(self):
        """保存按钮点击处理"""
        data = self.get_data()
        chosen = self.existing_combo.get() if self.use_existing.get() else ""
        edited = self.has_changes()

        if chosen and not edited:
            self.result = {"action": "reference", "target": chosen}
        else:
            self.result = {"action": "save", "data": data}

        if self._on_save:
            self._on_save(self.result)

    def get_save_result(self):
        """获取保存结果（供外部手动调用时使用，不触发 on_save）"""
        data = self.get_data()
        chosen = self.existing_combo.get() if self.use_existing.get() else ""
        edited = self.has_changes()
        if chosen and not edited:
            return {"action": "reference", "target": chosen}
        return {"action": "save", "data": data}

# ====================================================================
# 通用截面参数组件（Excel数据源）
# ====================================================================
_GEOM_COLS = {"H","B","tw","tf","tf1","tf2","r1","r2","D","d","C","dx"}

def load_section_db(filepath, families=None):
    """加载 Excel 截面数据, 保留几何尺寸列"""
    db = {}
    import openpyxl
    wb = openpyxl.load_workbook(filepath, data_only=True)
    for sn in wb.sheetnames:
        if families is not None and sn not in families:
            continue
        ws = wb[sn]
        headers = [ws.cell(1, c).value for c in range(1, ws.max_column+1)]
        for r in range(2, ws.max_row+1):
            row = {headers[c]: ws.cell(r, c+1).value for c in range(len(headers))}
            name = row.get("型号")
            if not name:
                continue
            geo = {k: v for k, v in row.items() if k in _GEOM_COLS and v is not None}
            db[name] = {"family": sn, "params": geo}
    return db


class SectionSetupFrame(ttk.Frame):
    """截面参数设置

    用法:
        # 作为子frame嵌入
        frame = SectionSetupFrame(parent, initial_data=..., callback=cb)
        frame.pack(...)

        # 作为对话框
        result = SectionSetupFrame.open_dialog(parent, initial_data=...)
    """
    _FAMILY_LABEL = {"C":"槽钢截面","I":"工字钢截面","HM":"HM型钢截面","HN":"HN型钢截面","HW":"HW型钢截面",
        "2C":"双拼槽钢截面","2I":"双拼工字钢截面","2HM":"双拼HM型钢截面","2HN":"双拼HN型钢截面","2HW":"双拼HW型钢截面"
        ,"O":"圆钢管","∠":"角钢截面"}
    _IMG_MAP = {"槽钢截面":"C_beam.png","工字钢截面":"H_beam.png","HM型钢截面":"H_beam.png","HN型钢截面":"H_beam.png","HW型钢截面":"H_beam.png",
        "双拼槽钢截面":"double_C_beam.png","双拼工字钢截面":"box_section.png","双拼HM型钢截面":"box_section.png","双拼HN型钢截面":"box_section.png","双拼HW型钢截面":"box_section.png",
        "圆钢管":"chs.png"}
    _NO_DOUBLE = {"∠","圆钢管"}
    _LABEL_TO_FAMILY = {v:k for k,v in _FAMILY_LABEL.items()}
    _FAMILY_COLS = {
        "I":  ["A","Ix","Iy","Wx","Wy","ix","iy","每延米重","H","B","tw","tf","r1","r2"],
        "2I": ["A","Ix","Iy","Wx","Wy","ix","iy","每延米重","H","B","tw","tf1","C","tf2"],
        "HM": ["A","Ix","Iy","Wx","Wy","ix","iy","每延米重","H","B","tw","tf","r1"],
        "2HM":["A","Ix","Iy","Wx","Wy","ix","iy","每延米重","H","B","tw","tf1","C","tf2"],
        "HN": ["A","Ix","Iy","Wx","Wy","ix","iy","每延米重","H","B","tw","tf","r1"],
        "2HN":["A","Ix","Iy","Wx","Wy","ix","iy","每延米重","H","B","tw","tf1","C","tf2"],
        "HW": ["A","Ix","Iy","Wx","Wy","ix","iy","每延米重","H","B","tw","tf","r1"],
        "2HW":["A","Ix","Iy","Wx","Wy","ix","iy","每延米重","H","B","tw","tf1","C","tf2"],
        "C":  ["A","Ix","Iy","Wx","Wy","ix","iy","每延米重","H","B","tw","tf1","dx"],
        "2C": ["A","Ix","Iy","Wx","Wy","ix","iy","每延米重","H","B","tw","tf1","C"],
        "∠": ["A","Ix","Iy","Wx","Wy","ix","iy","每延米重","H","B","tw","r1","r2"],
        "圆钢管":["A","I","W","i","D","d"],
    }

    @classmethod
    def open_dialog(cls, parent, initial_data=None, families=None):
        win = tb.Toplevel(parent); px=parent.winfo_rootx(); py=parent.winfo_rooty()
        win.title("截面参数设置"); win.geometry(f"700x550+{px}+{py}")
        win.resizable(False,False); win.transient(parent); win.grab_set()
        frame = cls(win, initial_data=initial_data, families=families)
        frame.pack(fill="both", expand=True)
        rh = []
        def _cb(d): rh.append(d); win.destroy()
        frame.set_on_confirm(_cb)
        win.protocol("WM_DELETE_WINDOW", win.destroy)
        parent.wait_window(win)
        return rh[0] if rh else None

    def __init__(self, parent, initial_data=None, families=None, callback=None):
        super().__init__(parent)
        self.families = families or list(self._FAMILY_LABEL.keys())
        self._callback = callback; self.result = None
        Applocation = read_Register("Software\\ShuZhiQiaoShi","Applocation")
        sp = os.path.join(Applocation, "Support","basic_param", "properties_parameter.xlsx")
        if not os.path.exists(sp):
            sp = os.path.join(os.path.dirname(__file__), "properties_parameter.xlsx")
        self._db = load_section_db(sp, families=self.families)
        self.png_path = os.path.join(Applocation, "Support", "png_picture")
        self.section_var = tk.StringVar(); self.section_name_var = tk.StringVar()
        self.param_entries = {}
        self._custom_cache = {}
        self._setup_ui(); self._load_initial_data(initial_data)

    def _type_labels(self):
        all_labels = [self._FAMILY_LABEL.get(f, f) for f in self.families]
        special_sections = {"圆钢管", "角钢截面"}
        normal_set = []; double_set = []; tail_set = []
        for lbl in all_labels:
            if lbl in special_sections: tail_set.append(lbl)
            elif "双拼" in lbl: double_set.append(lbl)
            else: normal_set.append(lbl)
        return sorted(normal_set, key=len) + sorted(double_set, key=len) + tail_set

    def _models_for(self, lbl):
        fams = [f for f,lb in self._FAMILY_LABEL.items() if lb == lbl]
        return [n for n in self._db if self._db[n]["family"] in fams] + ["自定义"]

    def set_on_confirm(self, cb): self._callback = cb

    def _setup_ui(self):
        bf = ttk.Frame(self); bf.pack(fill="x", padx=20, pady=15,side="bottom")
        tf = ttk.Frame(self); tf.pack(fill="x", padx=20, pady=(20,10))
        ttk.Label(tf, text="选择截面类型：").pack(side="left")
        self.combo_type = ttk.Combobox(tf, values=self._type_labels(), textvariable=self.section_var, width=15, state="readonly")
        self.combo_type.pack(side="left", padx=10)
        self.combo_type.bind("<<ComboboxSelected>>", self._on_type_change)
        mf = ttk.Frame(self); mf.pack(fill="both", expand=True, padx=20, pady=5)
        mf.columnconfigure(0, weight=4); mf.columnconfigure(1, weight=5)
        pf = ttk.LabelFrame(mf, text="型号选择/参数输入", padding=10)
        pf.grid(row=0, column=0, sticky="nsew", padx=(0,10))
        self.r1 = ttk.Frame(pf); self.r1.pack(fill="x", pady=5)
        ttk.Label(self.r1, text="截面型号：").pack(side="left")
        self.combo_name = ttk.Combobox(self.r1, textvariable=self.section_name_var, width=15, state="readonly")
        self.combo_name.pack(side="left", padx=5)
        self.combo_name.bind("<<ComboboxSelected>>", self._on_name_change)
        self.r2_custom = ttk.Frame(pf)
        ttk.Label(self.r2_custom, text="截面名称：").pack(side="left")
        self.entry_custom_name = ttk.Entry(self.r2_custom, width=15)
        self.entry_custom_name.pack(side="left", padx=5)
        self.dpc = ttk.Frame(pf); self.dpc.pack(fill="both", expand=True, pady=10)
        self.preview_frame = ttk.LabelFrame(mf, text="截面示意图", padding=10)
        self.preview_frame.grid(row=0, column=1, sticky="nsew")
        self.preview_label = ttk.Label(self.preview_frame, text="暂无图片", anchor="center")
        self.preview_label.pack(expand=True, fill="both")
        tb.Button(bf, text="确认", bootstyle=SUCCESS, width=10, command=self.on_confirm).pack(side="right")
        tb.Button(bf, text="取消", bootstyle=SECONDARY, width=10, command=self._on_cancel).pack(side="right", padx=10)

    def _load_initial_data(self, data):
        if not data:
            self.section_var.set(self._type_labels()[0]); self._on_type_change(); return
        sn = data.get("section_name","")
        params = data.get("params", {})
        is_std = sn in self._db
        if is_std:
            fam = self._db[sn]["family"]
        else:
            st = data.get("section_type","")
            # st 可能是代码（如 "O"）或标签（如 "圆钢管"），需要统一处理
            # 先检查是否是代码
            if st in self._FAMILY_LABEL:
                fam = st
            else:
                # 否则当作标签，转换为代码
                rev = {v:k for k,v in self._FAMILY_LABEL.items()}
                fam = rev.get(st, self.families[0])
        lbl = self._FAMILY_LABEL.get(fam, self._type_labels()[0])
        self.section_var.set(lbl); self._on_type_change()
        if is_std:
            self.section_name_var.set(sn)
        else:
            self.section_name_var.set("自定义")
            custom_lbl = data.get("custom_label","") or sn
            if custom_lbl and custom_lbl != "自定义":
                self.entry_custom_name.delete(0,"end")
                self.entry_custom_name.insert(0, custom_lbl)
            if params:
                self._custom_cache[lbl] = dict(params)
        self._on_name_change()

    def _on_type_change(self, event=None):
        models = self._models_for(self.section_var.get())
        self.combo_name["values"] = models
        if models: self.section_name_var.set(models[0])
        self._on_name_change()

    def _on_name_change(self, event=None):
        name = self.section_name_var.get(); is_c = (name == "自定义")
        is_pipe = (self.section_var.get() == "圆钢管")
        if is_c and not is_pipe:
            self.r2_custom.pack(fill="x", pady=(10,5), after=self.r1)
        else:
            self.r2_custom.pack_forget()
        self._update_preview(is_c); self._refresh_params(is_c)

    def _update_preview(self, is_c=False):
        st = self.section_var.get()
        im = self._IMG_MAP.get(st)
        if im:
            p = os.path.join(self.png_path, im)
            if os.path.exists(p):
                try:
                    img = Image.open(p).resize((250,200), Image.LANCZOS)
                    self.preview_imgtk = ImageTk.PhotoImage(img)
                    self.preview_label.config(image=self.preview_imgtk, text=""); return
                except: pass
        self.preview_label.config(text="暂无图片")

    def _refresh_params(self, is_c=False):
        prev_custom = getattr(self, '_was_custom', False)
        st_label = self.section_var.get()
        if prev_custom:
            for k, e in self.param_entries.items():
                self._custom_cache.setdefault(st_label, {})[k] = e.get()
        for w in self.dpc.winfo_children(): w.destroy()
        self.param_entries.clear()
        name = self.section_name_var.get(); is_c = (name == "自定义")
        self._was_custom = is_c
        if is_c:
            fam = [f for f,lb in self._FAMILY_LABEL.items() if lb==st_label][0]
            first = next((n for n in self._db if self._db[n]["family"]==fam), None)
            keys = list(self._db[first]["params"].keys()) if first else ["H","B","tw","tf"]
            cached = self._custom_cache.get(st_label, {})
        else:
            info = self._db.get(name, {}); keys = list(info.get("params",{}).keys())
            cached = {}
        for i,key in enumerate(keys):
            ttk.Label(self.dpc, text=f"{key}:", width=14).grid(row=i,column=0,sticky="e",pady=3,padx=(0,5))
            e = ttk.Entry(self.dpc, width=10)
            e.grid(row=i,column=1,sticky="w",pady=3)
            if is_c:
                if key in cached and cached[key] != "":
                    e.insert(0, cached[key])
            else:
                v = self._db.get(name, {}).get("params", {}).get(key)
                if v is not None:
                    e.insert(0, str(v))
            e.config(state="normal" if is_c else "normal")
            self.param_entries[key] = e

    def _collect(self):
        n = self.section_name_var.get()
        custom_lbl = self.entry_custom_name.get().strip() if n=="自定义" else ""
        if n=="自定义" and not custom_lbl:
            if self.section_var.get() == "圆钢管":
                D = self.param_entries.get("D",self.param_entries.get("D(mm)"))
                d = self.param_entries.get("d",self.param_entries.get("d(mm)"))
                if D and d:
                    custom_lbl = f"{D.get().strip()}×{d.get().strip()}"
                else:
                    custom_lbl = "自定义截面"
            else:
                custom_lbl = "自定义截面"
        return {"section_type": self.section_var.get(), "section_name": n,
                "custom_label": custom_lbl if n=="自定义" else "",
                "params": {k:v.get().strip() for k,v in self.param_entries.items()}}

    def _validate(self):
        """验证截面数据，返回错误信息列表"""
        errors = []
        n = self.section_name_var.get()

        # 1. 自定义截面时检查名称
        if n == "自定义":
            custom_lbl = self.entry_custom_name.get().strip()
            if not custom_lbl:
                # 圆钢管自动生成名称
                if self.section_var.get() == "圆钢管":
                    D = self.param_entries.get("D", self.param_entries.get("D(mm)"))
                    d = self.param_entries.get("d", self.param_entries.get("d(mm)"))
                    if not D or not d or not D.get().strip() or not d.get().strip():
                        errors.append("圆钢管截面必须输入 D 和 d 参数")
                    else:
                        custom_lbl = f"{D.get().strip()}×{d.get().strip()}"
                else:
                    errors.append("自定义截面必须输入截面名称")
            
            # 检查名称是否与截面库重名
            if custom_lbl and custom_lbl in self._db:
                errors.append(f"截面名称 '{custom_lbl}' 已存在，请重命名")

        # 2. 检查参数是否为空或非数值
        for key, entry in self.param_entries.items():
            val = entry.get().strip()
            if not val:
                errors.append(f"参数 '{key}' 不能为空")
            else:
                try:
                    float(val)
                except ValueError:
                    errors.append(f"参数 '{key}' 必须为数值，当前值: '{val}'")

        return errors

    def _close_win(self):
        p = self.master
        if isinstance(p, tk.Toplevel): p.destroy()
        else: self.destroy()

    def on_confirm(self):
        # 数据验证
        errors = self._validate()
        if errors:
            messagebox.showerror("输入错误", "\n".join(errors), parent=self.winfo_toplevel())
            return
        self.result = self._collect()
        if self._callback: self._callback(self.result)
        else: self._close_win()

    def _on_cancel(self):
        self.result = None
        if self._callback: self._callback(None)
        else: self._close_win()


def create_section_selector(parent, label_text, row, initial_data=None, textvariable=None, families=None, on_change=None, on_focus=None, label_width=20, entry_width=25, padx=5, pady=5, mode="dialog", on_select=None):
    """统一截面选择入口.

    mode="dialog" : Label + 按钮, 点击弹窗 (默认)
    mode="embed"  : Label + SectionSetupFrame 直接嵌入
    """
    def _summary(data):
        if not data: return "请选择截面"
        sn = data.get("custom_label","") or data.get("section_name","")
        if sn == "自定义" or not sn:
            return "自定义截面" if sn == "自定义" else "请选择截面"
        return sn if sn != "/" else "请选择截面"

    # --- 嵌入模式 ---
    if mode == "embed":
        frame = SectionSetupFrame(parent, initial_data=initial_data, families=families)
        ttk.Label(parent, text=label_text, width=label_width, bootstyle=PRIMARY).grid(row=row, column=0, sticky="nw", padx=padx, pady=pady)
        frame.grid(row=row, column=1, columnspan=2, sticky="nsew", padx=padx, pady=pady)
        frame.get_section_data = frame._collect
        return frame

    # --- 弹窗模式 ---
    btn = tb.Button(parent, textvariable=textvariable if textvariable else None,
                    text="请选择截面" if not textvariable else None,
                    takefocus=False, bootstyle="outline-primary")

    if label_text:
        ttk.Label(parent, text=label_text, width=label_width, bootstyle=PRIMARY).grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
        btn.grid(row=row, column=1, columnspan=2, sticky="we", padx=padx, pady=pady)
    else:
        btn.pack(padx=2, pady=2, fill="both", expand=True)

    if isinstance(initial_data, str):
        initial_data = {"section_name": initial_data}
    btn.section_data = initial_data
    if on_focus:
        btn.bind("<FocusIn>", lambda e: on_focus(btn))

    def _refresh():
        txt = _summary(btn.section_data)
        if textvariable:
            textvariable.set(txt)
        else:
            btn.configure(text=txt)

    def _open():
        # 无效截面数据（空/"/"）不传入弹窗，让它加载第一个标准截面
        sd = btn.section_data
        if not sd or (isinstance(sd, dict) and sd.get("section_name", "") in ("", "/")):
            sd = None
        result = SectionSetupFrame.open_dialog(parent.winfo_toplevel(), initial_data=sd, families=families)
        if result:
            # 将组装后的截面名称回写到 section_name，保证 get_section_data() 返回正确名称
            display = _summary(result)
            if display and display not in ("请选择截面", "自定义截面"):
                result["section_name"] = display
            btn.section_data = result
            _refresh()
            if on_change: on_change(result)
            if on_select:
                st = result.get("section_type", "")
                name = result.get("section_name", "")
                params = result.get("params", {})
                _LABEL_TO_FAMILY = SectionSetupFrame._LABEL_TO_FAMILY
                type_code = _LABEL_TO_FAMILY.get(st, st)
                on_select({"type": type_code, "name": name, "params": params})

    btn.configure(command=_open)
    _refresh()  # 初始化显示
    btn.get_section_data = lambda: btn.section_data
    return btn