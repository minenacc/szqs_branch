# 1. 标准库
import os
import sys
import winreg
from pathlib import Path

# 2. 第三方库
import ttkbootstrap as tb
from tkinter import messagebox
from ttkbootstrap.constants import *

# 3. 本地模块
sys.path.append(str(Path(__file__).parent.parent.parent))
from General.Midas import MidasURL, Open_Midas_civil, get_current_midas_file
from General.ProgramMonitor import Program_State_Monitoring
from General.UIHandle import if_Reg, make_card_frame
from General.DataUtils import read_Register

sys.path.append(str(Path(__file__).parent))
from Trestle_File_Cal import Open_Midas_mcb, Cal_path_SaveAs, get_trestle_projects, read_trestle_params_from_excel
from Trestle_Generate_Cal import make_all_doc

# 当前py文件对应总对话框函数

#=====================================================================================================================================================================

# 点击运行
def Make_Doc(Applocation, model_post_result, openfile_path, foundation_name, foundation_standard, foundation_values, saveas_path, param_excel_path, project_name, trestle_no):
    if "mcb" in openfile_path:
        if "docx" in saveas_path:
            # 从Excel读取参数
            try:
                excel_params = read_trestle_params_from_excel(param_excel_path, project_name, trestle_no)
                print("Excel参数读取成功:", excel_params)
            except Exception as e:
                print(f"Excel参数读取失败: {e}")
                return
            if model_post_result[0] == {}:
                make_all_doc(True, Applocation, openfile_path, model_post_result, foundation_name, foundation_standard, foundation_values, saveas_path, excel_params=excel_params)
            else:
                response = messagebox.askyesno("确认", "检测到已存在模型结果，是否重新读取")
                make_all_doc(response, Applocation, openfile_path, model_post_result, foundation_name, foundation_standard, foundation_values, saveas_path, excel_params=excel_params)
        else:
            print("未输入保存路径")
    else:
        print("未识别到有效mcb文件, 请重新选择")
    

# 退出/关闭/结束程序
def on_exit(root):
    if root.winfo_exists():  # 安全检查
        root.destroy()  # 正确关闭方式

#=====================================================================================================================================================================

# 标题栏
def dialog_headr(root, txt):
    # 宋体加粗
    style = tb.Style()
    style.configure("SimSun.TLabel", font=("SimSun", 10, "bold"))
    container = tb.Frame(root)
    container.pack(fill=X, expand=YES)
    hdr = tb.Label(master=root, text=txt, width=50, style="SimSun.TLabel")
    hdr.pack(fill=X, pady=10)


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


# 两个标签 + 两个下拉列表 + 一个按钮
def dialog_boxline4(root, label_txt1, label_txt2, combo_lst1, combo_lst2, button_txt, button_command):
    container = tb.Frame(root)
    container.pack(fill=X, expand=YES, pady=5)
    # 标签1
    tb.Label(master=container, text=label_txt1.title(), width = 10).pack(side="left", padx=5)
    # 第一个下拉列表
    combo1 = tb.Combobox(master=container, width = 15)
    combo1.pack(side=LEFT, padx=10, fill=X)
    combo1['values'] = combo_lst1
    combo1.current(0)
    # 标签2
    tb.Label(master=container, text=label_txt2.title(), width = 8).pack(side="left", padx=5)
    # 第二个下拉列表
    combo2 = tb.Combobox(master=container, width = 15)
    combo2.pack(side=LEFT, padx=0, fill=X)

    def update_second_combobox(event):
        if combo1.get() == "无":
            combo2['values'] = ["无"]
        else:
            combo2['values'] = combo_lst2
        combo2.current(0)
        
    # 初始设置第二个下拉列表
    update_second_combobox(None)
    # 绑定事件
    combo1.bind("<<ComboboxSelected>>", update_second_combobox)
    # 按钮
    button1 = tb.Button(master=container, text=button_txt, command = button_command, width = 10)
    button1.pack(side=RIGHT, padx=10)
    
    return combo1, combo2, button1


# 标签 + 编辑框
def dialog_boxline5(root, label_txt, variable):
    container = tb.Frame(root)
    container.pack(fill=X, expand=YES, pady=5)
    tb.Label(master=container, text=label_txt.title(), width=22).pack(side=LEFT, padx=30, pady=0, fill=X, expand=YES)
    tb.Entry(master=container, textvariable=variable, width=20).pack(side=LEFT, padx=30, pady=0, fill=X, expand=YES)


# 标签 + 按钮
def dialog_boxline6(root, label_txt, button_txt, button_command):
    container = tb.Frame(root)
    container.pack(fill=X, expand=YES, pady=5)
    tb.Label(master=container, text=label_txt.title(), width=22).pack(side=LEFT, padx=30, pady=0, fill=X, expand=YES)
    button1 = tb.Button(master=container, text=button_txt, command = button_command, width=19)
    button1.pack(side=RIGHT, padx=30)


# 标签 + 下拉列表 + 按钮
def dialog_boxline7(root, label_txt, combo_lst, combo_length, button_txt, button_command):
    container = tb.Frame(root)
    container.pack(fill=X, expand=YES, pady=5)
    # 标签
    tb.Label(master=container, text=label_txt.title(), width = 11).pack(side="left", padx=5)
    # 下拉列表
    combo = tb.Combobox(master=container, width = 15)
    combo.pack(side=LEFT, padx=0, fill=X)
    combo['values'] = combo_lst
    combo.current(0)
    # 按钮
    button = tb.Button(master=container, text=button_txt, command = button_command, width = 10)
    button.pack(side=RIGHT, padx=10)
    return combo, button



# 标签 + 编辑框 + 标签 + 编辑框
def dialog_boxline8(root, label_txt1, label_length1, entry_var1, entry_length1, label_txt2, label_length2, entry_var2, entry_length2):
    container = tb.Frame(root)
    container.pack(fill=X, expand=YES, pady=5)
    # 标签
    Label1 = tb.Label(master=container, text=label_txt1.title(), width=label_length1)
    Label1.pack(side=LEFT, padx=0)
    # 编辑框
    ent1 = tb.Entry(master=container, textvariable=entry_var1, width=entry_length1)
    ent1.pack(side=LEFT, padx=20, fill=X, expand=YES)
    # 标签
    Label2 = tb.Label(master=container, text=label_txt2.title(), width=label_length2)
    Label2.pack(side=LEFT, padx=0)
    # 编辑框
    ent2 = tb.Entry(master=container, textvariable=entry_var2, width=entry_length2)
    ent2.pack(side=LEFT, padx=20, fill=X, expand=YES)


# 最后一行
def dialog_boxline_end(root, button_txt1, button_command1, button_txt2, button_command2):
    container = tb.Frame(root)
    container.pack(fill=X, expand=YES, pady=10)
    # 分隔线
    sep_line = tb.Separator(master=container, orient=HORIZONTAL)
    sep_line.pack(fill=X)
    # 退出按钮
    but = tb.Button(master=container, text=button_txt2, command = button_command2, width=8)
    but.pack(side=RIGHT, padx=10, pady=10)
    # 确认按钮
    sub_btn = tb.Button(master=container, text=button_txt1, command = button_command1, bootstyle=SUCCESS, width=8)
    sub_btn.pack(side=RIGHT, pady=10)


# 二级对话框的确认和取消
def secondary_dialog_submit_quit(dialog, button, dialog_values, var_dict, var_lst, dict_key, foundation_name, param_keys):
    frame = tb.Frame(dialog)
    frame.pack(fill=X, expand=YES, pady=(15, 10))
    # 分隔线
    sep_line = tb.Separator(master=frame, orient=HORIZONTAL)
    sep_line.pack(fill=X, pady=5)
    # 退出按钮
    cnl_btn = tb.Button(master=frame,text="退出", command = lambda: enable_button(dialog, button), bootstyle = PRIMARY, width = 8)
    cnl_btn.pack(side=RIGHT, padx=10)
    # 确认按钮：保存参数后销毁窗口并重新启用主按钮
    sub_btn = tb.Button(master=frame, text="确认",
                         command=lambda: _save_and_close(dialog, button, dialog_values, var_dict, var_lst, dict_key, foundation_name, param_keys),
                         bootstyle=SUCCESS, width=8)
    sub_btn.pack(side=RIGHT, padx=5)
    sub_btn.focus_set()


def _save_and_close(dialog, button, dialog_values, var_dict, var_lst, dict_key, foundation_name, param_keys):
    """保存二级对话框参数，然后销毁窗口并重新启用主按钮。"""
    secondary_dialog_set_save_params(dialog_values, var_dict, var_lst, dict_key, foundation_name, param_keys)
    enable_button(dialog, button)


# 二级对话框确认时，对主对话框传进来的变量赋值，保存二级对话框的参数
def secondary_dialog_set_save_params(dialog_values, var_dict, var_lst, dict_key, foundation_name, param_keys):
    """保存二级对话框参数为字典。

    参数:
        dialog_values:  主对话框传入的共享字典（foundation_secondary_dialog_values）
        var_dict:       主对话框传入的参数字典（foundation_params_dict），下次打开时回填默认值
        var_lst:        UI Variable 对象列表（顺序与 param_keys 一一对应）
        dict_key:       写入 dialog_values 的键（"深基础"/"浅基础"）
        foundation_name: 写入 var_dict 的键（"打入桩"/"钻孔灌注桩"/"扩大基础"/"条形基础"）
        param_keys:     参数名列表，与 var_lst 顺序一致
    """
    values = [var.get() if hasattr(var, 'get') else var for var in var_lst]
    saved = dict(zip(param_keys, values))
    dialog_values.__setitem__(dict_key, saved)
    var_dict[foundation_name] = saved
    print(f"[{foundation_name}] 参数已保存: {saved}")


# 重新启用主窗口的按钮，并销毁二级窗口
def enable_button(secondary_window, button):
    """重新启用主窗口的按钮，并销毁二级窗口"""
    button.config(state=tb.NORMAL)
    secondary_window.destroy()   


CONCRETE_GRADE_LIST = ['C15', 'C20', 'C25', 'C30', 'C35', 'C40', 'C45', 'C50',
                       'C55', 'C60', 'C65', 'C70', 'C75', 'C80']
REBAR_GRADE_LIST = ['HPB300', 'HRB335', 'HRB400', 'HRBF400', 'RRB400',
                    'HRB500', 'HRBF500']


def _param_row(parent, label_text, variable, entry_width=12):
    """平铺对齐行：标签 + 编辑框（与 _dual_param_row 一致的 padding）"""
    row = tb.Frame(parent)
    row.pack(fill=X, pady=3)
    tb.Label(row, text=label_text, width=18, font=("Microsoft YaHei UI", 9)).pack(side=LEFT, padx=(10, 5))
    tb.Entry(row, textvariable=variable, width=entry_width).pack(side=LEFT, padx=(0, 20))


def _dual_param_row(parent, row_idx, label1, var1, label2, var2, entry_width=12):
    """两两一排：grid 布局，左 label+entry | 右 label+entry"""
    tb.Label(parent, text=label1, width=18, font=("Microsoft YaHei UI", 9)).grid(
        row=row_idx, column=0, sticky="w", padx=(10, 5), pady=3)
    tb.Entry(parent, textvariable=var1, width=entry_width).grid(
        row=row_idx, column=1, sticky="w", padx=(0, 20), pady=3)
    tb.Label(parent, text=label2, width=18, font=("Microsoft YaHei UI", 9)).grid(
        row=row_idx, column=2, sticky="e", padx=(10, 5), pady=3)
    tb.Entry(parent, textvariable=var2, width=entry_width).grid(
        row=row_idx, column=3, sticky="w", padx=(0, 20), pady=3)




# ── 二级窗口：打入桩 ──────────────────────────────────────────
_DRIVEN_PILE_KEYS = ['x0a']

def Deep_Foundation_secondary_dialog_steelpile(root, button, secondary_dialog_values, var_dict, dict_key):
    secondary_dialog = tb.Toplevel(root)
    secondary_dialog.title("打入桩 - 参数设定")
    defaults = var_dict.get("打入桩") or {'x0a': 10}

    _, card = make_card_frame(secondary_dialog, "桩基参数")
    x0a     = tb.DoubleVar(value=defaults.get('x0a', 10))
    _param_row(card, "水平位移值 x0a(mm)", x0a)

    var_lst = [x0a]
    secondary_dialog_submit_quit(secondary_dialog, button, secondary_dialog_values,
                                 var_dict, var_lst, dict_key, "打入桩", _DRIVEN_PILE_KEYS)
    button.config(state=tb.DISABLED)
    secondary_dialog.protocol("WM_DELETE_WINDOW", lambda: enable_button(secondary_dialog, button))


# ── 二级窗口：钻孔灌注桩 ──────────────────────────────────────────
_DRILLED_SHAFT_KEYS = ['concrete_grade', 'rebar_grade', 'pile_rho_g', 'cover_thick', 'x0a']

def Deep_Foundation_secondary_dialog_boredpile(root, button, secondary_dialog_values, var_dict, dict_key):
    secondary_dialog = tb.Toplevel(root)
    secondary_dialog.title("钻孔灌注桩 - 参数设定")
    defaults = var_dict.get("钻孔灌注桩") or {'concrete_grade': 'C30', 'rebar_grade': 'HRB400', 'pile_rho_g': 0.5, 'cover_thick': 50, 'x0a': 10}

    _, card = make_card_frame(secondary_dialog, "桩基参数")
    concrete_grade = tb.StringVar(value=defaults.get('concrete_grade', 'C30'))
    rebar_grade    = tb.StringVar(value=defaults.get('rebar_grade', 'HRB400'))
    pile_rho_g     = tb.DoubleVar(value=defaults.get('pile_rho_g', 0.5))
    cover_thick    = tb.DoubleVar(value=defaults.get('cover_thick', 50))
    x0a            = tb.DoubleVar(value=defaults.get('x0a', 10))

    # 混凝土等级
    row_conc = tb.Frame(card)
    row_conc.pack(fill=X, pady=3)
    tb.Label(row_conc, text="混凝土等级", width=18, font=("Microsoft YaHei UI", 9)).pack(side=LEFT, padx=(10, 5))
    combo_conc = tb.Combobox(row_conc, values=CONCRETE_GRADE_LIST, state="readonly", width=10)
    combo_conc.pack(side=LEFT, padx=(0, 20))
    combo_conc.set(concrete_grade.get())
    combo_conc.bind("<<ComboboxSelected>>", lambda _: concrete_grade.set(combo_conc.get()))

    # 钢筋等级
    row_rebar = tb.Frame(card)
    row_rebar.pack(fill=X, pady=3)
    tb.Label(row_rebar, text="钢筋等级", width=18, font=("Microsoft YaHei UI", 9)).pack(side=LEFT, padx=(10, 5))
    combo_rebar = tb.Combobox(row_rebar, values=REBAR_GRADE_LIST, state="readonly", width=10)
    combo_rebar.pack(side=LEFT, padx=(0, 20))
    combo_rebar.set(rebar_grade.get())
    combo_rebar.bind("<<ComboboxSelected>>", lambda _: rebar_grade.set(combo_rebar.get()))

    _param_row(card, "桩身正截面配筋率(%)", pile_rho_g)
    _param_row(card, "钢筋保护层厚度(mm)", cover_thick)
    _param_row(card, "水平位移值 x0a(mm)", x0a)

    var_lst = [concrete_grade, rebar_grade, pile_rho_g, cover_thick, x0a]
    secondary_dialog_submit_quit(secondary_dialog, button, secondary_dialog_values,
                                 var_dict, var_lst, dict_key, "钻孔灌注桩", _DRILLED_SHAFT_KEYS)
    button.config(state=tb.DISABLED)
    secondary_dialog.protocol("WM_DELETE_WINDOW", lambda: enable_button(secondary_dialog, button))


# ── 二级窗口：扩大基础 ──────────────────────────────────────────
_SPREAD_KEYS = [
    'side_A', 'side_B', 'height', 'gc',
    'fa', 'h2_top', 'theta', 'faz', 'has_weak',
    'Ks', 'conc_grd', 'rebar_grd', 'cover', 'space_min', 'space_max', 'dia_min', 'd_max'
]
_SPREAD_DEFAULTS = {
    'side_A': 2.0, 'side_B': 2.0, 'height': 1.0, 'gc': 26.5,
    'fa': 200, 'h2_top': 0, 'theta': 0, 'faz': 0, 'has_weak': False,
    'Ks': 1.35, 'conc_grd': 'C30', 'rebar_grd': 'HRB400', 'cover': 50,
    'space_min': 100, 'space_max': 200, 'dia_min': 12, 'd_max': 25
}

def Spread_Foundation_secondary_dialog(root, button, secondary_dialog_values, var_dict, dict_key):
    secondary_dialog = tb.Toplevel(root)
    secondary_dialog.title("扩大基础 - 参数设定")
    defaults = var_dict.get("扩大基础") or _SPREAD_DEFAULTS

    # 基础基本参数
    _, card1 = make_card_frame(secondary_dialog, "基础基本参数")
    side_A = tb.DoubleVar(value=defaults.get('side_A', 2.0))
    side_B = tb.DoubleVar(value=defaults.get('side_B', 2.0))
    height = tb.DoubleVar(value=defaults.get('height', 1.0))
    gc     = tb.DoubleVar(value=defaults.get('gc', 26.5))
    _dual_param_row(card1, 0, "边长A(m):", side_A, "边长B(m):", side_B)
    _dual_param_row(card1, 1, "基础高度h(m):", height, "混凝土容重(kN/m³):", gc)
    card1.columnconfigure(0, weight=0)
    card1.columnconfigure(1, weight=0)
    card1.columnconfigure(2, weight=1)
    card1.columnconfigure(3, weight=0)

    # 基底承载力验算
    _, card2 = make_card_frame(secondary_dialog, "基底承载力验算")
    fa     = tb.DoubleVar(value=defaults.get('fa', 200))
    h2_top = tb.DoubleVar(value=defaults.get('h2_top', 0))
    theta  = tb.DoubleVar(value=defaults.get('theta', 0))
    faz    = tb.DoubleVar(value=defaults.get('faz', 0))
    has_weak = tb.BooleanVar(value=defaults.get('has_weak', False))

    # row 0: fa
    tb.Label(card2, text="地基承载力fa(kPa)", width=18, font=("Microsoft YaHei UI", 9)).grid(
        row=0, column=0, sticky="w", padx=(10, 5), pady=3)
    tb.Entry(card2, textvariable=fa, width=12).grid(
        row=0, column=1, sticky="w", padx=(0, 20), pady=3)

    # row 1: checkbox + 扩散角
    chk_weak = tb.Checkbutton(card2, text="考虑软弱下卧层", variable=has_weak,
                               bootstyle="round-toggle")
    chk_weak.grid(row=1, column=0, sticky="w", padx=(10, 5), pady=3)
    tb.Label(card2, text="扩散角θ(°)", width=18, font=("Microsoft YaHei UI", 9)).grid(
        row=1, column=2, sticky="e", padx=(10, 5), pady=3)
    ent_theta = tb.Entry(card2, textvariable=theta, width=12)
    ent_theta.grid(row=1, column=3, sticky="w", padx=(0, 20), pady=3)

    # row 2: 下卧层顶标高 + 下卧层承载力
    tb.Label(card2, text="下卧层顶标高(m)", width=18, font=("Microsoft YaHei UI", 9)).grid(
        row=2, column=0, sticky="w", padx=(10, 5), pady=3)
    ent_h2 = tb.Entry(card2, textvariable=h2_top, width=12)
    ent_h2.grid(row=2, column=1, sticky="w", padx=(0, 20), pady=3)
    tb.Label(card2, text="下卧层承载力faz(kPa)", width=18, font=("Microsoft YaHei UI", 9)).grid(
        row=2, column=2, sticky="e", padx=(10, 5), pady=3)
    ent_faz = tb.Entry(card2, textvariable=faz, width=12)
    ent_faz.grid(row=2, column=3, sticky="w", padx=(0, 20), pady=3)

    card2.columnconfigure(0, weight=0)
    card2.columnconfigure(1, weight=0)
    card2.columnconfigure(2, weight=1)
    card2.columnconfigure(3, weight=0)

    # checkbox 联动：未勾选时禁用下卧层参数
    _weak_widgets = [ent_theta, ent_h2, ent_faz]

    def _toggle_weak(*_):
        state = "normal" if has_weak.get() else "disabled"
        for w in _weak_widgets:
            w.config(state=state)
        if not has_weak.get():
            h2_top.set(0)
            theta.set(0)
            faz.set(0)

    has_weak.trace_add("write", _toggle_weak)
    _toggle_weak()  # 初始化状态

    # tips
    tb.Label(card2, text="*下卧层顶标高按照以地面为+0标高输入",
              font=("Microsoft YaHei UI", 8), foreground="gray").grid(
        row=3, column=0, columnspan=4, sticky="w", padx=(10, 5), pady=(0, 3))

    # 基础结构强度验算
    _, card3 = make_card_frame(secondary_dialog, "基础结构强度验算")
    Ks        = tb.DoubleVar(value=defaults.get('Ks', 1.35))
    conc_grd  = tb.StringVar(value=defaults.get('conc_grd', 'C30'))
    rebar_grd = tb.StringVar(value=defaults.get('rebar_grd', 'HRB400'))
    cover     = tb.DoubleVar(value=defaults.get('cover', 50))
    space_min = tb.DoubleVar(value=defaults.get('space_min', 100))
    space_max = tb.DoubleVar(value=defaults.get('space_max', 200))
    dia_min   = tb.DoubleVar(value=defaults.get('dia_min', 12))
    d_max     = tb.DoubleVar(value=defaults.get('d_max', 25))

    _dual_param_row(card3, 0, "组合转换系数Ks:", Ks, "钢筋保护层厚度(mm):", cover)

    # 混凝土等级 + 钢筋等级
    tb.Label(card3, text="混凝土等级:", width=18, font=("Microsoft YaHei UI", 9)).grid(
        row=1, column=0, sticky="w", padx=(10, 5), pady=3)
    combo_conc = tb.Combobox(card3, values=CONCRETE_GRADE_LIST, state="readonly", width=10)
    combo_conc.grid(row=1, column=1, sticky="w", padx=(0, 20), pady=3)
    combo_conc.set(conc_grd.get())
    combo_conc.bind("<<ComboboxSelected>>", lambda _: conc_grd.set(combo_conc.get()))

    tb.Label(card3, text="钢筋等级:", width=18, font=("Microsoft YaHei UI", 9)).grid(
        row=1, column=2, sticky="e", padx=(10, 5), pady=3)
    combo_rebar = tb.Combobox(card3, values=REBAR_GRADE_LIST, state="readonly", width=10)
    combo_rebar.grid(row=1, column=3, sticky="w", padx=(0, 20), pady=3)
    combo_rebar.set(rebar_grd.get())
    combo_rebar.bind("<<ComboboxSelected>>", lambda _: rebar_grd.set(combo_rebar.get()))

    _dual_param_row(card3, 2, "钢筋间距最小值(mm):", space_min, "钢筋间距最大值(mm):", space_max)
    _dual_param_row(card3, 3, "钢筋直径最小值(mm):", dia_min, "钢筋直径最大值(mm):", d_max)
    card3.columnconfigure(0, weight=0)
    card3.columnconfigure(1, weight=0)
    card3.columnconfigure(2, weight=1)
    card3.columnconfigure(3, weight=0)

    var_lst = [side_A, side_B, height, gc,
               fa, h2_top, theta, faz, has_weak,
               Ks, conc_grd, rebar_grd, cover, space_min, space_max, dia_min, d_max]
    secondary_dialog_submit_quit(secondary_dialog, button, secondary_dialog_values,
                                 var_dict, var_lst, dict_key, "扩大基础",
                                 param_keys=_SPREAD_KEYS)
    button.config(state=tb.DISABLED)
    secondary_dialog.protocol("WM_DELETE_WINDOW", lambda: enable_button(secondary_dialog, button))


# ── 二级窗口：条形基础 ──────────────────────────────────────────
_STRIP_KEYS = [
    'width_b', 'height', 'gc',
    'fa', 'h2_top', 'theta', 'faz', 'has_weak',
    'Ks', 'conc_grd', 'rebar_grd', 'cover', 'space_min', 'space_max', 'dia_min', 'dia_max'
]
_STRIP_DEFAULTS = {
    'width_b': 1.0, 'height': 0.8, 'gc': 26.5,
    'fa': 200, 'h2_top': 0, 'theta': 0, 'faz': 200, 'has_weak': True,
    'Ks': 1.35, 'conc_grd': 'C30', 'rebar_grd': 'HRB400', 'cover': 50,
    'space_min': 100, 'space_max': 200, 'dia_min': 12, 'dia_max': 25
}

def Strip_Foundation_secondary_dialog(root, button, secondary_dialog_values, var_dict, dict_key):
    secondary_dialog = tb.Toplevel(root)
    secondary_dialog.title("条形基础 - 参数设定")
    defaults = var_dict.get("条形基础") or _STRIP_DEFAULTS

    # 基础基本参数（宽度 + 高度一行，容重单独一行）
    _, card1 = make_card_frame(secondary_dialog, "基础基本参数")
    width_b = tb.DoubleVar(value=defaults.get('width_b', 1.0))
    height  = tb.DoubleVar(value=defaults.get('height', 0.8))
    gc      = tb.DoubleVar(value=defaults.get('gc', 26.5))
    _dual_param_row(card1, 0, "基础宽度b(m):", width_b, "基础高度h(m):", height)
    tb.Label(card1, text="混凝土容重(kN/m³):", width=18, font=("Microsoft YaHei UI", 9)).grid(
        row=1, column=0, sticky="w", padx=(10, 5), pady=3)
    tb.Entry(card1, textvariable=gc, width=12).grid(
        row=1, column=1, sticky="w", padx=(0, 20), pady=3)
    card1.columnconfigure(0, weight=0)
    card1.columnconfigure(1, weight=0)
    card1.columnconfigure(2, weight=1)
    card1.columnconfigure(3, weight=0)

    # 基底承载力验算
    _, card2 = make_card_frame(secondary_dialog, "基底承载力验算")
    fa     = tb.DoubleVar(value=defaults.get('fa', 200))
    h2_top = tb.DoubleVar(value=defaults.get('h2_top', 0))
    theta  = tb.DoubleVar(value=defaults.get('theta', 0))
    faz    = tb.DoubleVar(value=defaults.get('faz', 200))
    has_weak = tb.BooleanVar(value=defaults.get('has_weak', True))

    # row 0: fa
    tb.Label(card2, text="地基承载力fa(kPa)", width=18, font=("Microsoft YaHei UI", 9)).grid(
        row=0, column=0, sticky="w", padx=(10, 5), pady=3)
    tb.Entry(card2, textvariable=fa, width=12).grid(
        row=0, column=1, sticky="w", padx=(0, 20), pady=3)

    # row 1: checkbox + 扩散角
    chk_weak = tb.Checkbutton(card2, text="考虑软弱下卧层", variable=has_weak,
                               bootstyle="round-toggle")
    chk_weak.grid(row=1, column=0, sticky="w", padx=(10, 5), pady=3)
    tb.Label(card2, text="扩散角θ(°)", width=18, font=("Microsoft YaHei UI", 9)).grid(
        row=1, column=2, sticky="e", padx=(10, 5), pady=3)
    ent_theta = tb.Entry(card2, textvariable=theta, width=12)
    ent_theta.grid(row=1, column=3, sticky="w", padx=(0, 20), pady=3)

    # row 2: 下卧层顶标高 + 下卧层承载力
    tb.Label(card2, text="下卧层顶标高(m)", width=18, font=("Microsoft YaHei UI", 9)).grid(
        row=2, column=0, sticky="w", padx=(10, 5), pady=3)
    ent_h2 = tb.Entry(card2, textvariable=h2_top, width=12)
    ent_h2.grid(row=2, column=1, sticky="w", padx=(0, 20), pady=3)
    tb.Label(card2, text="下卧层承载力faz(kPa)", width=18, font=("Microsoft YaHei UI", 9)).grid(
        row=2, column=2, sticky="e", padx=(10, 5), pady=3)
    ent_faz = tb.Entry(card2, textvariable=faz, width=12)
    ent_faz.grid(row=2, column=3, sticky="w", padx=(0, 20), pady=3)

    card2.columnconfigure(0, weight=0)
    card2.columnconfigure(1, weight=0)
    card2.columnconfigure(2, weight=1)
    card2.columnconfigure(3, weight=0)

    # checkbox 联动：未勾选时禁用下卧层参数
    _weak_widgets = [ent_theta, ent_h2, ent_faz]

    def _toggle_weak(*_):
        state = "normal" if has_weak.get() else "disabled"
        for w in _weak_widgets:
            w.config(state=state)
        if not has_weak.get():
            h2_top.set(0)
            theta.set(0)
            faz.set(0)

    has_weak.trace_add("write", _toggle_weak)
    _toggle_weak()

    # tips
    tb.Label(card2, text="*下卧层顶标高按照以地面为+0标高输入",
              font=("Microsoft YaHei UI", 8), foreground="gray").grid(
        row=3, column=0, columnspan=4, sticky="w", padx=(10, 5), pady=(0, 3))

    # 基础结构强度验算
    _, card3 = make_card_frame(secondary_dialog, "基础结构强度验算")
    Ks        = tb.DoubleVar(value=defaults.get('Ks', 1.35))
    conc_grd  = tb.StringVar(value=defaults.get('conc_grd', 'C30'))
    rebar_grd = tb.StringVar(value=defaults.get('rebar_grd', 'HRB400'))
    cover     = tb.DoubleVar(value=defaults.get('cover', 50))
    space_min = tb.DoubleVar(value=defaults.get('space_min', 100))
    space_max = tb.DoubleVar(value=defaults.get('space_max', 200))
    dia_min   = tb.DoubleVar(value=defaults.get('dia_min', 12))
    dia_max   = tb.DoubleVar(value=defaults.get('dia_max', 25))

    _dual_param_row(card3, 0, "组合转换系数Ks:", Ks, "钢筋保护层厚度(mm):", cover)

    # 混凝土等级 + 钢筋等级
    tb.Label(card3, text="混凝土等级:", width=18, font=("Microsoft YaHei UI", 9)).grid(
        row=1, column=0, sticky="w", padx=(10, 5), pady=3)
    combo_conc = tb.Combobox(card3, values=CONCRETE_GRADE_LIST, state="readonly", width=10)
    combo_conc.grid(row=1, column=1, sticky="w", padx=(0, 20), pady=3)
    combo_conc.set(conc_grd.get())
    combo_conc.bind("<<ComboboxSelected>>", lambda _: conc_grd.set(combo_conc.get()))

    tb.Label(card3, text="钢筋等级:", width=18, font=("Microsoft YaHei UI", 9)).grid(
        row=1, column=2, sticky="e", padx=(10, 5), pady=3)
    combo_rebar = tb.Combobox(card3, values=REBAR_GRADE_LIST, state="readonly", width=10)
    combo_rebar.grid(row=1, column=3, sticky="w", padx=(0, 20), pady=3)
    combo_rebar.set(rebar_grd.get())
    combo_rebar.bind("<<ComboboxSelected>>", lambda _: rebar_grd.set(combo_rebar.get()))

    _dual_param_row(card3, 2, "钢筋间距最小值(mm):", space_min, "钢筋间距最大值(mm):", space_max)
    _dual_param_row(card3, 3, "钢筋直径最小值(mm):", dia_min, "钢筋直径最大值(mm):", dia_max)
    card3.columnconfigure(0, weight=0)
    card3.columnconfigure(1, weight=0)
    card3.columnconfigure(2, weight=1)
    card3.columnconfigure(3, weight=0)

    var_lst = [width_b, height, gc,
               fa, h2_top, theta, faz, has_weak,
               Ks, conc_grd, rebar_grd, cover, space_min, space_max, dia_min, dia_max]
    secondary_dialog_submit_quit(secondary_dialog, button, secondary_dialog_values,
                                 var_dict, var_lst, dict_key, "条形基础",
                                 param_keys=_STRIP_KEYS)
    button.config(state=tb.DISABLED)
    secondary_dialog.protocol("WM_DELETE_WINDOW", lambda: enable_button(secondary_dialog, button))


# 基础参数设定的二级对话框（路由分发）
def foundation_secondary_dialog(root, foundation_name, formula_name, button, secondary_dialog_values, var_dict):
    if foundation_name == "打入桩" and formula_name != '无':
        Deep_Foundation_secondary_dialog_steelpile(root, button, secondary_dialog_values,
                                                   var_dict, "深基础")
    elif foundation_name == "钻孔灌注桩" and formula_name != '无':
        Deep_Foundation_secondary_dialog_boredpile(root, button, secondary_dialog_values,
                                                   var_dict, "深基础")
    elif foundation_name == "扩大基础":
        Spread_Foundation_secondary_dialog(root, button, secondary_dialog_values, var_dict, "浅基础")
    elif foundation_name == "条形基础":
        Strip_Foundation_secondary_dialog(root, button, secondary_dialog_values, var_dict, "浅基础")
    else:
        print('请选择基础形式及对应计算规范')

#=====================================================================================================================================================================

# 对话框总函数
def creat_dialog(root, Applocation):

    # 用于判断是否已经读取过模型结果
    model_post_result = [{}]

    # 文件操作大类=============================================================================================================

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
    Set_File_entry_var3 = tb.StringVar(value = current_active_mcb.replace(".mcb", ".docx") if current_active_mcb else os.path.join(Applocation, 'Custom', "ZQ_Untitled_Cal.docx")) # 另存为路径

    # 项目选择顶部栏=========================================================================================================

    # 从Excel读取项目列表
    param_excel_path = os.path.join(Applocation, 'Support', 'program_param', 'Trestle_ParamTable.xlsx')
    trestle_projects = get_trestle_projects(param_excel_path)
    project_names = sorted(set(p[0] for p in trestle_projects))

    selected_project = tb.StringVar()
    selected_trestle_no = tb.StringVar()

    # 顶部栏
    top_bar = tb.Frame(root)
    top_bar.pack(side=TOP, fill=X)
    proj_inner = tb.Frame(top_bar, padding=(15, 10))
    proj_inner.pack(fill=X)
    # 左侧：项目选择 + 栈桥编号
    tb.Label(proj_inner, text="项目选择", font=("Microsoft YaHei UI", 10, "bold")).pack(side=LEFT, padx=(0, 10))
    combo_project = tb.Combobox(proj_inner, textvariable=selected_project, values=project_names, state="readonly", width=15)
    combo_project.pack(side=LEFT, padx=(0, 5))
    if project_names:
        combo_project.current(0)
    tb.Label(proj_inner, text="栈桥编号", font=("Microsoft YaHei UI", 9)).pack(side=LEFT, padx=(0, 5))
    combo_trestle_no = tb.Combobox(proj_inner, textvariable=selected_trestle_no, state="readonly")
    combo_trestle_no.pack(side=LEFT, fill=X, expand=True, padx=(0, 10))
    # 右侧：查看基本参数按钮
    def show_params_dialog():
        """二级窗口：查看当前项目的基本参数（只读 tksheet 表格）"""
        import tksheet
        proj = selected_project.get()
        trestle_no = selected_trestle_no.get()
        if not proj or not trestle_no:
            return
        try:
            params = read_trestle_params_from_excel(param_excel_path, proj, trestle_no)
        except Exception as e:
            messagebox.showerror("错误", f"读取参数失败: {e}")
            return

        # 车辆参数中文标签映射
        vehicle_param_labels = {
            'axle_name': '标准荷载轴名', 'forces': '轴重(kN)', 'distances': '轴距(m)',
            'weight': '自重(kN)', 'capacity': '吊重(kN)', 'boom': '吊臂(m)',
            'track_length': '履带尺寸(m)', 'track_width': '履带宽度(m)',
            'torque': '扭矩(kN)',
        }

        # 构建表格数据：[类别, 参数名, 值]
        rows = []
        row_types = []

        # 风荷载参数
        wind_items = [
            ("规范", params.get('wind_standard', '')),
            ("地表分类", params.get('ground_class', '')),
            ("地形修正系数", params.get('terrain_factor', '')),
            ("主梁基准高度(m)", params.get('ref_height', '')),
            ("水平加载长度系数", params.get('length_factor', '')),
            ("设计基准风速公式", params.get('wind_formula', '')),
            ("设计风速(m/s)", params.get('design_wind_speed', '')),
        ]
        for i, (name, val) in enumerate(wind_items):
            rows.append(["风荷载" if i == 0 else "", name, str(val if val is not None else '')])
            row_types.append('cat' if i == 0 else 'param')

        # 水流力参数
        flow_items = [
            ("规范", params.get('flow_standard', '')),
            ("设计水流速(m/s)", params.get('design_flow_speed', '')),
        ]
        for i, (name, val) in enumerate(flow_items):
            rows.append(["水流力" if i == 0 else "", name, str(val if val is not None else '')])
            row_types.append('cat' if i == 0 else 'param')

        # 车辆荷载参数
        vehicle_list = params.get('vehicle_params_list', [])
        print("车辆参数:", vehicle_list)  # 调试输出
        if vehicle_list:
            first = True
            for v in vehicle_list:
                rows.append(["车辆荷载" if first else "", v['name'], ""])
                row_types.append('vehicle')
                first = False
                # 车辆类型
                rows.append(["", "  车辆荷载类型", v['params'].get('type', '')])
                row_types.append('sub')
                # 轴距
                rows.append(["", "  轴距(m)", str(v['wheelbase'] if v['wheelbase'] is not None else '')])
                row_types.append('sub')
                # 各参数
                for k, val in v['params'].items():
                    if k in ('type', 'weight', 'forces_formula', 'distances_formula'):
                        continue
                    if val is not None and val != '':
                        # 履带尺寸特殊处理：合并长×宽
                        if k == 'track_length':
                            tw = v['params'].get('track_width', '')
                            val_str = f"{val}×{tw}" if tw else str(val)
                            rows.append(["", "  履带尺寸(m)", val_str])
                        elif k == 'track_width':
                            continue  # 已在 track_length 中合并
                        else:
                            label = vehicle_param_labels.get(k, k)
                            rows.append(["", f"  {label}", str(val)])
                        row_types.append('sub')
        else:
            rows.append(["车辆荷载", "(无)", ""])
            row_types.append('cat')

        # 创建窗口
        dlg = tb.Toplevel(root)
        dlg.title(f"基本参数 - {proj} / {trestle_no}")
        dlg.geometry("620x520")
        dlg.transient(root)
        dlg.grab_set()

        # 标题
        tb.Label(dlg, text=f"{proj} - {trestle_no}",
                  font=("Microsoft YaHei UI", 11, "bold"), bootstyle=PRIMARY).pack(pady=(10, 5))

        # tksheet
        sheet_frame = tb.Frame(dlg)
        sheet_frame.pack(fill=BOTH, expand=True, padx=10, pady=(0, 5))
        sheet = tksheet.Sheet(
            sheet_frame,
            data=rows,
            headers=["类别", "参数名", "值"],
            header_height=30,
            row_height=26,
            font=("Microsoft YaHei UI", 9, "normal"),
            header_font=("Microsoft YaHei UI", 9, "bold"),
            theme="light",
            header_bg="#f0f4f8",
            header_fg="#1e293b",
            grid_color="#e2e8f0",
            show_row_index=False,
        )
        sheet.pack(fill=BOTH, expand=True)
        # 列宽
        sheet.column_width(column=0, width=80)
        sheet.column_width(column=1, width=200)
        sheet.column_width(column=2, width=300)
        # 禁用编辑
        sheet.disable_bindings(
            "edit_cell", "edit_header", "edit_index",
            "rc_insert_row", "rc_delete_row", "rc_delete_column",
            "cut", "paste", "delete", "undo",
        )
        sheet.readonly = True

        # 样式
        for r, rt in enumerate(row_types):
            if rt == 'vehicle':
                # 车辆名行：参数名列蓝底蓝字
                sheet.highlight_cells(row=r, column=1, bg="#e8f0fe", fg="#1a56db")
            else:
                # 其余所有行：参数名列灰底
                sheet.highlight_cells(row=r, column=1, bg="#f8f9fa", fg="#333333")

        # 关闭按钮
        tb.Button(dlg, text="关闭", command=dlg.destroy, width=8).pack(pady=8)

    btn_frame = tb.Frame(proj_inner)
    btn_frame.pack(side=RIGHT, padx=5)
    tb.Button(btn_frame, text="查看基本参数", command=show_params_dialog, width=12).pack(side=LEFT, padx=2)

    # 编号联动
    def update_trestle_no_list(event=None):
        proj = selected_project.get()
        numbers = [p[1] for p in trestle_projects if p[0] == proj]
        combo_trestle_no['values'] = numbers
        if numbers:
            combo_trestle_no.current(0)

    combo_project.bind("<<ComboboxSelected>>", update_trestle_no_list)
    update_trestle_no_list()

    # 分隔线
    tb.Separator(top_bar).pack(fill=X)

    # 文件操作卡片===========================================================================================================

    _, file_card = make_card_frame(root, "文件操作")
    # 每行：标签 + 编辑框 + 按钮，横向排列
    def _form_row(parent, label_text, entry_var, btn_text, btn_cmd):
        row = tb.Frame(parent)
        row.pack(fill=X, pady=4, padx=10)
        tb.Label(row, text=label_text, width=10, font=("Microsoft YaHei UI", 9)).pack(side=LEFT)
        tb.Entry(row, textvariable=entry_var, width=35, state="disabled").pack(side=LEFT, padx=(5, 10), fill=X, expand=True)
        tb.Button(row, text=btn_text, width=16, command=btn_cmd).pack(side=RIGHT)

    _form_row(file_card, "程序状态", Set_File_entry_var1, "打开Midas civil NX",
              lambda: Open_Midas_civil(civilnx_path, Set_File_entry_var1))
    _form_row(file_card, "模型名称", Set_File_entry_var2, "选取mcb文件并打开",
              lambda: Open_Midas_mcb(Set_File_entry_var2, Set_File_entry_var3))
    _form_row(file_card, "保存路径", Set_File_entry_var3, "计算书另存为",
              lambda: Cal_path_SaveAs(Set_File_entry_var3))

    # 基础计算参数设定卡片=====================================================================================================

    # 基础形式 → 支持的计算规范映射
    foundation_standard_map = {
        "打入桩": ["建筑桩基技术规范"],
        "钻孔灌注桩": ["建筑桩基技术规范"],
        "扩大基础":   ["建筑地基基础设计规范"],
        "条形基础":   ["建筑地基基础设计规范"],
    }
    foundation_types = list(foundation_standard_map.keys())
    # 各基础类型的默认参数，用户未打开二级窗口时直接采用
    _FOUNDATION_DEFAULTS = {
        "打入桩": {'x0a': 10},
        "钻孔灌注桩": {'concrete_grade': 'C30', 'rebar_grade': 'HRB400', 'pile_rho_g': 0.5, 'cover_thick': 50, 'x0a': 10},
        "扩大基础": dict(_SPREAD_DEFAULTS),
        "条形基础": dict(_STRIP_DEFAULTS),
    }
    foundation_params_dict = {ftype: _FOUNDATION_DEFAULTS.get(ftype, {}) for ftype in foundation_types}

    _, param_card = make_card_frame(root, "基础计算参数设定")

    # 第一行：基础形式 + 计算规范 + 设定参数按钮
    row1 = tb.Frame(param_card)
    row1.pack(fill=X, pady=4, padx=10)
    tb.Label(row1, text="基础形式", width=10, font=("Microsoft YaHei UI", 9)).pack(side=LEFT)
    selected_foundation_type = tb.StringVar(value=foundation_types[0])
    combo_foundation = tb.Combobox(row1, textvariable=selected_foundation_type,
                                    values=foundation_types, state="readonly", width=15)
    combo_foundation.pack(side=LEFT, padx=(5, 10))

    tb.Label(row1, text="计算规范", width=10, font=("Microsoft YaHei UI", 9)).pack(side=LEFT)
    selected_foundation_standard = tb.StringVar()
    combo_standard = tb.Combobox(row1, textvariable=selected_foundation_standard,
                                  state="readonly", width=20)
    combo_standard.pack(side=LEFT, padx=(5, 10))

    def _on_foundation_type_change(event=None):
        ftype = selected_foundation_type.get()
        standards = foundation_standard_map.get(ftype, [])
        combo_standard['values'] = standards
        if standards:
            combo_standard.current(0)

    combo_foundation.bind("<<ComboboxSelected>>", _on_foundation_type_change)
    _on_foundation_type_change()

    # 二级窗口共享状态
    foundation_secondary_dialog_values = {}

    def _open_foundation_params():
        """打开对应基础类型的参数设定二级窗口"""
        ftype = selected_foundation_type.get()
        fstd  = selected_foundation_standard.get()
        foundation_secondary_dialog(root, ftype, fstd, btn_set_params,
                                    foundation_secondary_dialog_values,
                                    foundation_params_dict)


    btn_set_params = tb.Button(row1, text="设定参数", width=12, command=_open_foundation_params)
    btn_set_params.pack(side=RIGHT)

    # 运行和关闭===============================================================================================================

    dialog_boxline_end(root, "运行", lambda: Make_Doc(Applocation, model_post_result, Set_File_entry_var2.get(), selected_foundation_type.get(), selected_foundation_standard.get(), foundation_params_dict, Set_File_entry_var3.get(), param_excel_path, selected_project.get(), selected_trestle_no.get()), "关闭", lambda: on_exit(root))


#=====================================================================================================================================================================

def Trestle_Cal_Report_Main(parent=None, on_close=None):
    if not if_Reg(parent):
        if on_close:
            on_close()
        return
    
    Applocation = read_Register('Software\\ShuZhiQiaoShi', 'Applocation')
    if parent == None:
        root = tb.Window("施工栈桥计算书自动化助手")
    else:
        root = tb.Toplevel(parent)
        root.title("施工栈桥计算书自动化助手")
    creat_dialog(root, Applocation)
    if parent == None:
        root.mainloop()
    elif on_close:
        def _on_root_destroy(event, _root=root):
            if event.widget is _root:
                on_close()
        root.bind('<Destroy>', _on_root_destroy)

if __name__ == "__main__":
    Trestle_Cal_Report_Main()
