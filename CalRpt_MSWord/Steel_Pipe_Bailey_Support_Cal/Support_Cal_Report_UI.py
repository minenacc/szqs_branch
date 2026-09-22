# 1. 标准库
import os
import sys
import winreg
from pathlib import Path
from tkinter import messagebox

# 2. 第三方库
import ttkbootstrap as tb
from ttkbootstrap.constants import *

# 3. 本地模块
sys.path.append(str(Path(__file__).parent.parent.parent))
from General.TextHandle     import read_file
from General.DataUtils      import read_Register
from General.ProgramMonitor import Program_State_Monitoring
from General.ExcelHandle    import ZJ_get_soild_excel_data_tolst
from General.Midas          import MidasAPI, MidasURL, Open_Midas_civil
from General.FilePath       import get_desktop_path_from_registry
from General.UIHandle       import if_Reg, enable_button, dialog_headr, on_exit

sys.path.append(str(Path(__file__).parent))
from CalRpt_MSWord.Steel_Pipe_Bailey_Support_Cal.Support_Cal_Report_Generate   import make_all_doc
from CalRpt_MSWord.Steel_Pipe_Bailey_Support_Cal.Support_Cal_File_Manipulation import Open_Midas_mcb, Cal_path_SaveAs
from CalRpt_MSWord.Steel_Pipe_Bailey_Support_Cal.Support_Midas_Model_Pre       import get_node_ele_str, midas_group_output

#=====================================================================================================================================================================

# 点击运行
def Make_Doc(Applocation, model_post_result, openfile_path, soleplate_values, foundation_name, foundation_standard, foundation_values, saveas_path):
    if "mcb" in openfile_path:
        # 判断用户是否需要重新读取模型结果，对于已经读取过的模型，节省模型读取的时间
        if "docx" in saveas_path:
            if model_post_result[0] == {}:
                make_all_doc(True, Applocation, openfile_path, model_post_result, soleplate_values, foundation_name, foundation_standard, foundation_values, saveas_path)
            else:
                response = messagebox.askyesno("确认", "检测到已存在模型结果，是否重新读取")
                make_all_doc(response, Applocation, openfile_path, model_post_result, soleplate_values, foundation_name, foundation_standard, foundation_values, saveas_path)
        else:
            print("未输入保存路径")
    else:
        print("未识别到有效mcb文件, 请重新选择")
    # print(model_post_result, openfile_path, foundation_name, foundation_standard, foundation_values)
    # print(soleplate_values)


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



# 标签 + 编辑框 + 下拉列表
def dialog_boxline1(root, label_txt, label_length, entry_var, entry_length, button_txt, button_command):
    container = tb.Frame(root)
    container.pack(fill=X, expand=YES, pady=5)
    # 标签
    Label = tb.Label(master=container, text=label_txt.title(), width=label_length)
    Label.pack(side=LEFT, padx=0)
    # 编辑框
    ent = tb.Entry(master=container, textvariable=entry_var, width=entry_length)
    ent.pack(side=LEFT, padx=20, fill=X, expand=YES)
    # 按钮
    but = tb.Button(master=container, text=button_txt, command = button_command, width = 15)
    but.pack(side=RIGHT, padx=10)


# 标签 + 下拉列表 + 编辑框 + 按钮
def dialog_boxline2(root, label_txt2, label_length2, combo_var, combo_lst, combo_length, entry_var2, entry_length2, button_txt, button_command):
    container = tb.Frame(root)
    container.pack(fill=X, expand=YES, pady=5)
    # 标签2
    Label2 = tb.Label(master=container, text=label_txt2.title(), width=label_length2)
    Label2.pack(side=LEFT, padx=0)
    # 下拉列表
    combo = tb.Combobox(master=container, textvariable=combo_var, values=combo_lst, width=combo_length)
    combo.pack(side=LEFT, padx=20)
    # 设置默认值
    combo.current(0)
    # 编辑框2
    ent2 = tb.Entry(master=container, textvariable=entry_var2, width=entry_length2, state="disabled")
    ent2.pack(side=LEFT, padx=20, fill=X, expand=YES)
    # 将状态更新函数绑定到编辑框的<FocusIn>事件
    ent2.bind("<FocusIn>", lambda e : update_entry_state(combo_var, ent2))
    # 同时保留下拉列表的绑定
    combo.bind("<<ComboboxSelected>>", lambda e : update_entry_state(combo_var, ent2))
    # 按钮
    but = tb.Button(master=container, text=button_txt, command = button_command, width = 15)
    but.pack(side=RIGHT, padx=10)


# 标签 + 编辑框 + 按钮
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
def dialog_boxline4(root, label_txt1, label_txt2, combo_lst1, combo_lst2, combo_lst3, button_txt, button_command):
    container = tb.Frame(root)
    container.pack(fill=X, expand=YES, pady=5)
    # 标签1
    tb.Label(master=container, text=label_txt1.title()).pack(side="left", padx=5)
    # 第一个下拉列表
    combo1 = tb.Combobox(master=container)
    combo1.pack(side=LEFT, padx=0, fill=X)
    combo1['values'] = combo_lst1
    combo1.current(0)
    # 标签2
    tb.Label(master=container, text=label_txt2.title()).pack(side="left", padx=5)
    # 第二个下拉列表
    combo2 = tb.Combobox(master=container)
    combo2.pack(side=LEFT, padx=0, fill=X)

    def update_second_combobox(event):
        if combo1.get() == "根据模型信息":
            combo2['values'] = ["根据模型信息"]
        elif combo1.get() == "扩大基础":
            combo2['values'] = combo_lst2
        else:
            combo2['values'] = combo_lst3
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
    tb.Label(master=container, text=label_txt.title()).pack(side="left", padx=5)
    # 下拉列表
    combo = tb.Combobox(master=container)
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
    container.pack(fill=X, expand=YES, padx=0, pady=5)
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


# 标签 + 编辑框
def dialog_boxline9(root, label_txt, label_length, variable, entry_length):
    container = tb.Frame(root)
    container.pack(padx=0, pady=5, anchor='w')
    label1 = tb.Label(master=container, text=label_txt.title(), width=label_length)
    label1.pack(side=LEFT, padx=0, pady=0)
    entry1 = tb.Entry(master=container, textvariable=variable, width=entry_length)
    entry1.pack(side=LEFT, padx=0, pady=0)


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

# 在扩大基础模块中用户需要输入的数据
# 共有参数
# 1、基础的长 —— foundation_x
# 2、基础的宽 —— foundation_y
# 3、基础的高 —— foundation_z
# 4、混凝土的容重 —— gc
# 5、基础底标高 —— Foundation_Base_Level
# 基底验算
# 6、基础上土体平均容重 —— gs
# 7、修正后地基承载力 —— foundation_fa
# 8、软弱下卧层顶标高 —— Weak_Layer_Horizon
# 9、软弱下卧层以上土体扩散角 —— Dispersion_Angle
# 10、软弱下卧层承载力 —— foundation_faz
# 基础检算
# 11、标准组合和基本组合之间的转换系数 —— ksi
# 12、混凝土等级 —— concrete_grade
# 13、钢筋牌号 —— rebar_grade
# 14、钢筋保护层厚度 —— as0
# 15、钢筋间距范围最小值 —— steel_s1
# 16、钢筋间距范围最大值 —— steel_s2
# 17、钢筋直径范围最小值 —— steel_d1
# 18、钢筋直径范围最大值 —— steel_d2


# 二级对话框的确认和取消
def secondary_dialog_submit_quit(dialog, button, dialog_values, var_dict, var_lst, dict_key, foundation_name):
    frame = tb.Frame(dialog)
    frame.pack(fill=X, expand=YES, pady=(15, 10))
    # 退出按钮
    cnl_btn = tb.Button(master=frame,text="退出", command = lambda: enable_button(dialog, button))
    cnl_btn.pack(side=RIGHT, padx=30)
    # 确认按钮
    # sub_btn = tb.Button(master=frame,text="确认", command = lambda: [dialog_values.__setitem__(dict_key, [var.get() for var in var_lst])])
    sub_btn = tb.Button(master=frame,text="确认", command = lambda: secondary_dialog_set_save_params(dialog_values, var_dict, var_lst, dict_key, foundation_name))
    sub_btn.pack(side=RIGHT, padx=5)
    sub_btn.focus_set()


# 二级对话框确认时，对主对话框传进来的变量赋值，保存二级对话框的参数
def secondary_dialog_set_save_params(dialog_values, var_dict, var_lst, dict_key, foundation_name):
    # 对主对话框传进来的变量赋值
    dialog_values.__setitem__(dict_key, [var.get() if isinstance(var, (tb.StringVar, tb.IntVar, tb.DoubleVar, tb.BooleanVar)) else var for var in var_lst])
    # 保存二级对话框的参数
    var_dict[foundation_name] = [var.get() if isinstance(var, (tb.StringVar, tb.IntVar, tb.DoubleVar, tb.BooleanVar)) else var for var in var_lst]
    print("参数信息已保存")


# 浅基础的参数设定二级对话框
def Shallow_Foundation_secondary_dialog(root, button, secondary_dialog_values, var_dict, dict_key):
    secondary_dialog = tb.Toplevel(root)
    secondary_dialog.title("浅基础参数设定框")
    # 基础的长
    foundation_x = tb.DoubleVar(value = var_dict["扩大基础"][0])
    # 基础的宽
    foundation_y = tb.DoubleVar(value = var_dict["扩大基础"][1])
    # 基础的高
    foundation_z = tb.DoubleVar(value = var_dict["扩大基础"][2])
    # 基础上立柱计算直径
    Pile_D = tb.DoubleVar(value = var_dict["扩大基础"][3])
    # 混凝土的容重
    concrete_gc = tb.DoubleVar(value = var_dict["扩大基础"][4])
    # 基础底标高
    Foundation_Base_Level = tb.DoubleVar(value = var_dict["扩大基础"][5])
    # 基础上土体平均容重
    soild_gs = tb.DoubleVar(value = var_dict["扩大基础"][6])
    # 修正后地基承载力
    foundation_fa = tb.DoubleVar(value = var_dict["扩大基础"][7])
    # 软弱下卧层顶标高
    Weak_Layer_Horizon = tb.DoubleVar(value = var_dict["扩大基础"][8])
    # 软弱下卧层以上土体扩散角
    Dispersion_Angle = tb.DoubleVar(value = var_dict["扩大基础"][9])
    # 软弱下卧层承载力
    foundation_faz = tb.DoubleVar(value = var_dict["扩大基础"][10])
    # 标准组合和基本组合之间的转换系数
    ksi = tb.DoubleVar(value = var_dict["扩大基础"][11])
    # 混凝土等级
    concrete_grade = tb.StringVar(value = var_dict["扩大基础"][12])
    # 钢筋牌号
    rebar_grade = tb.StringVar(value = var_dict["扩大基础"][13])
    # 钢筋保护层厚度
    as0 = tb.DoubleVar(value = var_dict["扩大基础"][14])
    # 钢筋间距范围最小值
    steel_s1 = tb.DoubleVar(value = var_dict["扩大基础"][15])
    # 钢筋间距范围最大值
    steel_s2 = tb.DoubleVar(value = var_dict["扩大基础"][16])
    # 钢筋直径范围最小值
    steel_d1 = tb.DoubleVar(value = var_dict["扩大基础"][17])
    # 钢筋直径范围最大值
    steel_d2 = tb.DoubleVar(value = var_dict["扩大基础"][18])
    # 每一行参数
    dialog_headr(secondary_dialog, "基础基本参数")
    dialog_boxline5(secondary_dialog, "基础与Mx对应边长A(m)", foundation_x)
    dialog_boxline5(secondary_dialog, "基础与My对应边长B(m)", foundation_y)
    dialog_boxline5(secondary_dialog, "基础高度(m)", foundation_z)
    dialog_boxline5(secondary_dialog, "立柱计算直径(m)", Pile_D)
    dialog_boxline5(secondary_dialog, "混凝土容重(kN/m3)", concrete_gc)
    dialog_boxline5(secondary_dialog, "基础底标高(m)", Foundation_Base_Level)
    dialog_headr(secondary_dialog, "基底承载力验算")
    dialog_boxline5(secondary_dialog, "基础上土体平均容重(kN/m3)", soild_gs)
    dialog_boxline5(secondary_dialog, "修正后地基承载力(kPa)", foundation_fa)
    dialog_boxline5(secondary_dialog, "软弱下卧层顶标高(m)", Weak_Layer_Horizon)
    dialog_boxline5(secondary_dialog, "软弱下卧层以上土体扩散角(°)", Dispersion_Angle)
    dialog_boxline5(secondary_dialog, "软弱下卧层承载力(kPa)", foundation_faz)
    dialog_headr(secondary_dialog, "基础结构强度验算")
    dialog_boxline5(secondary_dialog, "组合转换系数", ksi)
    dialog_boxline5(secondary_dialog, "混凝土等级", concrete_grade)
    dialog_boxline5(secondary_dialog, "钢筋牌号", rebar_grade)
    dialog_boxline5(secondary_dialog, "钢筋保护层厚度(mm)", as0)
    dialog_boxline5(secondary_dialog, "钢筋间距范围最小值(mm)", steel_s1)
    dialog_boxline5(secondary_dialog, "钢筋间距范围最大值(mm)", steel_s2)
    dialog_boxline5(secondary_dialog, "钢筋直径范围最小值(mm)", steel_d1)
    dialog_boxline5(secondary_dialog, "钢筋直径范围最大值(mm)", steel_d2) 

    # 确认和退出，将参数从二级对话框返回给主对话框
    var_lst = [foundation_x, foundation_y, foundation_z, Pile_D, concrete_gc,
               Foundation_Base_Level, soild_gs, foundation_fa, Weak_Layer_Horizon, Dispersion_Angle, foundation_faz,
               ksi, concrete_grade, rebar_grade, as0, steel_s1, steel_s2, steel_d1, steel_d2
            ]
    secondary_dialog_submit_quit(secondary_dialog, button, secondary_dialog_values, var_dict, var_lst, dict_key, "扩大基础")
    # 禁用主窗口的按钮
    button.config(state=tb.DISABLED)
    # 当二级窗口关闭时，重新启用主窗口的按钮
    secondary_dialog.protocol("WM_DELETE_WINDOW", lambda: enable_button(secondary_dialog, button))


# 深基础的参数设定二级对话框——钢管打入桩
def Deep_Foundation_secondary_dialog_steelpile(root, excel_path, button, secondary_dialog_values, var_dict, Soil_Layer_Parameter_Table, dict_key):
    # 生成模板
    # open_csv_saveas_xlsx(doc_save_path, Applocation)
    secondary_dialog = tb.Toplevel(root)
    secondary_dialog.title("深基础参数设定框")
    # 钢管桩直径
    Pile_D = tb.DoubleVar(value = var_dict["钢管打入桩"][0])
    # 钢管桩壁厚
    Pile_t = tb.DoubleVar(value = var_dict["钢管打入桩"][1])
    # 水平位移值
    Soil_vb = tb.DoubleVar(value = var_dict["钢管打入桩"][2])
    # 水平抗力计算土层
    # Soil_Layer_n = tb.IntVar(value = var_dict["钢管打入桩"][3])
    # 土层参数表的值
    # Soil_Layer_Parameter_Table = {"土层分类":[], "土层厚度":[], "黏聚力":[], "内摩擦角":[], "侧摩阻":[], "端阻":[]}
    # 对话框内容
    dialog_headr(secondary_dialog, "钢管打入桩参数")
    dialog_boxline5(secondary_dialog, "钢管桩直径(mm)", Pile_D)
    dialog_boxline5(secondary_dialog, "钢管桩壁厚(mm)", Pile_t)
    dialog_boxline5(secondary_dialog, "水平位移值(mm)", Soil_vb)
    dialog_boxline6(secondary_dialog, "读取土层参数表", "读取Excel文件", lambda: ZJ_get_soild_excel_data_tolst(Soil_Layer_Parameter_Table, excel_path))
    # dialog_boxline5(secondary_dialog, "水平抗力计算土层", Soil_Layer_n)
    # 确认和退出，将参数从二级对话框返回给主对话框
    # var_lst = [Pile_D, Pile_t, Soil_vb, Soil_Layer_n, Soil_Layer_Parameter_Table]
    var_lst = [Pile_D, Pile_t, Soil_vb, Soil_Layer_Parameter_Table]

    secondary_dialog_submit_quit(secondary_dialog, button, secondary_dialog_values, var_dict, var_lst, dict_key, "钢管打入桩")
    # 禁用主窗口的按钮
    button.config(state=tb.DISABLED)
    # 当二级窗口关闭时，重新启用主窗口的按钮
    secondary_dialog.protocol("WM_DELETE_WINDOW", lambda: enable_button(secondary_dialog, button))


# 深基础的参数设定二级对话框——混凝土预制桩
def Deep_Foundation_secondary_dialog_precastpile(root, excel_path, button, secondary_dialog_values, var_dict, Soil_Layer_Parameter_Table, dict_key):
    # 生成模板
    # open_csv_saveas_xlsx(desktop_path, Applocation)
    secondary_dialog = tb.Toplevel(root)
    secondary_dialog.title("深基础参数设定框")
    # 预制桩直径
    Pile_D = tb.DoubleVar(value = var_dict["混凝土预制桩"][0])
    # 预制桩壁厚
    Pile_t = tb.DoubleVar(value = var_dict["混凝土预制桩"][1])
    # 混凝土等级
    concrete_grade = tb.StringVar(value = var_dict["混凝土预制桩"][2])
    # 水平位移值
    Soil_vb = tb.DoubleVar(value = var_dict["混凝土预制桩"][3])
    # 水平抗力计算土层
    # Soil_Layer_n = tb.IntVar(value = var_dict["混凝土预制桩"][4])
    # 土层参数表的值
    # Soil_Layer_Parameter_Table = {"土层分类":[], "土层厚度":[], "黏聚力":[], "内摩擦角":[], "侧摩阻":[], "端阻":[]}
    # 对话框内容
    dialog_headr(secondary_dialog, "混凝土预制桩参数")
    dialog_boxline5(secondary_dialog, "预制桩直径(mm)", Pile_D)
    dialog_boxline5(secondary_dialog, "预制桩壁厚(mm)", Pile_t)
    dialog_boxline5(secondary_dialog, "混凝土等级", concrete_grade)
    dialog_boxline5(secondary_dialog, "水平位移值(mm)", Soil_vb)
    dialog_boxline6(secondary_dialog, "读取土层参数表", "读取Excel文件", lambda: ZJ_get_soild_excel_data_tolst(Soil_Layer_Parameter_Table, excel_path))
    # dialog_boxline5(secondary_dialog, "水平抗力计算土层", Soil_Layer_n)
    # 确认和退出，将参数从二级对话框返回给主对话框
    var_lst = [Pile_D, Pile_t, concrete_grade, Soil_vb, Soil_Layer_Parameter_Table]
    secondary_dialog_submit_quit(secondary_dialog, button, secondary_dialog_values, var_dict, var_lst, dict_key, "混凝土预制桩")
    # 禁用主窗口的按钮
    button.config(state=tb.DISABLED)
    # 当二级窗口关闭时，重新启用主窗口的按钮
    secondary_dialog.protocol("WM_DELETE_WINDOW", lambda: enable_button(secondary_dialog, button))


# 深基础的参数设定二级对话框——混凝土灌注桩
def Deep_Foundation_secondary_dialog_concretepile(root, excel_path, button, secondary_dialog_values, var_dict, Soil_Layer_Parameter_Table, dict_key):
    # 生成模板
    # open_csv_saveas_xlsx(desktop_path, Applocation)
    secondary_dialog = tb.Toplevel(root)
    secondary_dialog.title("深基础参数设定框")
    # 灌注桩直径
    Pile_D = tb.DoubleVar(value = var_dict["混凝土灌注桩"][0])
    # 混凝土等级
    concrete_grade = tb.StringVar(value = var_dict["混凝土灌注桩"][1])
    # 钢筋等级
    rebar_grade = tb.StringVar(value = var_dict["混凝土灌注桩"][2])
    # 桩身正截面配筋率
    Pile_p = tb.DoubleVar(value = var_dict["混凝土灌注桩"][3])
    # 钢筋保护层厚度
    Pile_as = tb.DoubleVar(value = var_dict["混凝土灌注桩"][4])
    # 水平位移值
    Soil_vb = tb.DoubleVar(value = var_dict["混凝土灌注桩"][5])
    # 水平抗力计算土层
    # Soil_Layer_n = tb.IntVar(value = var_dict["混凝土灌注桩"][6])
    # 土层参数表的值
    # Soil_Layer_Parameter_Table = {"土层分类":[], "土层厚度":[], "黏聚力":[], "内摩擦角":[], "侧摩阻":[], "端阻":[]}
    # 对话框内容
    dialog_headr(secondary_dialog, "混凝土灌注桩参数")
    dialog_boxline5(secondary_dialog, "预制桩直径(mm)", Pile_D)
    dialog_boxline5(secondary_dialog, "混凝土等级", concrete_grade)
    dialog_boxline5(secondary_dialog, "钢筋等级", rebar_grade)
    dialog_boxline5(secondary_dialog, "桩身正截面配筋率(%)", Pile_p)
    dialog_boxline5(secondary_dialog, "钢筋保护层厚度(mm)", Pile_as)
    dialog_boxline5(secondary_dialog, "水平位移值(mm)", Soil_vb)
    dialog_boxline6(secondary_dialog, "读取土层参数表", "读取Excel文件", lambda: ZJ_get_soild_excel_data_tolst(Soil_Layer_Parameter_Table, excel_path))
    # dialog_boxline5(secondary_dialog, "水平抗力计算土层", Soil_Layer_n)
    # 确认和退出，将参数从二级对话框返回给主对话框
    # var_lst = [Pile_D, concrete_grade, rebar_grade, Pile_p, Pile_as, Soil_vb, Soil_Layer_n, Soil_Layer_Parameter_Table]
    var_lst = [Pile_D, concrete_grade, rebar_grade, Pile_p, Pile_as, Soil_vb, Soil_Layer_Parameter_Table]

    secondary_dialog_submit_quit(secondary_dialog, button, secondary_dialog_values, var_dict, var_lst, dict_key, "混凝土灌注桩")
    # 禁用主窗口的按钮
    button.config(state=tb.DISABLED)
    # 当二级窗口关闭时，重新启用主窗口的按钮
    secondary_dialog.protocol("WM_DELETE_WINDOW", lambda: enable_button(secondary_dialog, button))


# 基础参数设定的二级对话框
def foundation_secondary_dialog(root, foundation_name, formula_name, button, secondary_dialog_values, var_dict, Soil_Layer_Parameter_Table, excel_path):
    if foundation_name == "扩大基础" and formula_name != '无':
        Shallow_Foundation_secondary_dialog(root, button, secondary_dialog_values, var_dict, "浅基础")
    elif foundation_name == "钢管打入桩" and formula_name != '无':
        Deep_Foundation_secondary_dialog_steelpile(root, excel_path, button, secondary_dialog_values, var_dict, Soil_Layer_Parameter_Table, "深基础")
    elif foundation_name == "混凝土预制桩" and formula_name != '无':
        Deep_Foundation_secondary_dialog_precastpile(root, excel_path, button, secondary_dialog_values, var_dict, Soil_Layer_Parameter_Table, "深基础")
    elif foundation_name == "混凝土灌注桩" and formula_name != '无':
        Deep_Foundation_secondary_dialog_concretepile(root, excel_path, button, secondary_dialog_values, var_dict, Soil_Layer_Parameter_Table, "深基础")
    else:
        print('请选择基础形式及对应计算规范')


def ZJB_FM_soleplate_secondary_dialog(root, button, soleplate_name, secondary_dialog_values, var_dict, Self_values_lst, dict_key):
    '''
    '''
    secondary_dialog = tb.Toplevel(root)
    secondary_dialog.title("竹胶板+方木模板参数设定框")
    ZhuJiaoBan_thickness = tb.IntVar(value = var_dict[soleplate_name][0]) # 竹胶板厚度
    FangMu_SEC = tb.StringVar(value = var_dict[soleplate_name][1]) # 方木截面
    Concrete_Cal_Height1 = tb.DoubleVar(value = var_dict[soleplate_name][2]) # 截面腹板高度
    Concrete_Cal_Height2 = tb.DoubleVar(value = var_dict[soleplate_name][3]) # 截面顶底板高度
    FangMu_Gap1 = tb.IntVar(value = var_dict[soleplate_name][4]) # 方木腹板下间距
    FangMu_Gap2 = tb.IntVar(value = var_dict[soleplate_name][5]) # 方木顶底板下间距
    Rib_SEC = tb.StringVar(value = var_dict[soleplate_name][6]) # 大肋截面
    Rib_Gap = tb.IntVar(value = var_dict[soleplate_name][7]) # 大肋间距
    Vibration_Load = tb.DoubleVar(value = var_dict[soleplate_name][8]) # 振捣荷载
    Concrete_Unit = tb.DoubleVar(value = Self_values_lst[1])
    Construction_Load = tb.DoubleVar(value = Self_values_lst[2])
    Formwork_Load = tb.DoubleVar(value = Self_values_lst[3])
    dialog_headr(secondary_dialog, "模板计算参数")
    dialog_boxline8(secondary_dialog, '竹胶板厚度(mm)', 15, ZhuJiaoBan_thickness, 10, '方木截面', 16, FangMu_SEC, 15)
    dialog_boxline8(secondary_dialog, '大肋截面', 15, Rib_SEC, 10, '大肋间距(mm)', 16, Rib_Gap, 15)
    dialog_headr(secondary_dialog, "截面计算参数")
    dialog_boxline8(secondary_dialog, '腹板计算高度(m)', 15, Concrete_Cal_Height1, 10, '顶底板计算高度(m)', 16, Concrete_Cal_Height2, 15)
    dialog_boxline8(secondary_dialog, '腹板方木间距(mm)', 15, FangMu_Gap1, 10, '顶底板方木间距(mm)', 16, FangMu_Gap2, 15)
    dialog_headr(secondary_dialog, "荷载计算参数")
    dialog_boxline9(secondary_dialog, '振捣荷载(kN/m2)', 17, Vibration_Load, 10)
    var_lst = [ZhuJiaoBan_thickness, FangMu_SEC, Concrete_Cal_Height1, Concrete_Cal_Height2, FangMu_Gap1, FangMu_Gap2, Rib_SEC, Rib_Gap, Concrete_Unit, Construction_Load, Vibration_Load, Formwork_Load]
    secondary_dialog_submit_quit(secondary_dialog, button, secondary_dialog_values, var_dict, var_lst, dict_key, soleplate_name)
    # 禁用主窗口的按钮
    button.config(state=tb.DISABLED)
    # 当二级窗口关闭时，重新启用主窗口的按钮
    secondary_dialog.protocol("WM_DELETE_WINDOW", lambda: enable_button(secondary_dialog, button))


# 模板参数设定的二级对话框
def soleplate_secondary_dialog(root, button, soleplate_name, secondary_dialog_values, var_dict, Self_values_lst):
    if soleplate_name == '竹胶板+方木':
        ZJB_FM_soleplate_secondary_dialog(root, button, soleplate_name, secondary_dialog_values, var_dict, Self_values_lst, '模板')
    if soleplate_name == '钢模板':
        messagebox.showinfo("提示", "暂不支持，功能开发中")


#=====================================================================================================================================================================

# 含主对话框全局变量的函数
# 获取节点和单元信息
def get_node_elem_dict(entry_var, NODE_dict_res, ELEM_dict_res):
    # global NODE_dict_res, ELEM_dict_res
    NODE_dict_res[0] = MidasAPI("GET", "/db/NODE") # 节点信息
    ELEM_dict_res[0] = MidasAPI("GET", "/db/ELEM") # 单元信息
    for node_key, elem_key in zip(NODE_dict_res[0].keys(), ELEM_dict_res[0].keys()):
        # print(node_key, elem_key)
        if node_key == 'NODE' and elem_key == 'ELEM':
            entry_var.set("成功获取模型节点和单元信息")
        else:
            entry_var.set("模型节点和单元信息获取失败, 请检查API连接") 

#=====================================================================================================================================================================

# 对话框总函数
def creat_dialog(root):
    
    # 用于判断是否已经读取过模型结果
    model_post_result = [{}]
    # 桌面默认路径
    desktop_path = get_desktop_path_from_registry().replace('/', '\\')
    # 注册表路径
    Applocation = read_Register('Software\\ShuZhiQiaoShi', 'Applocation')
    # print(Applocation)
    path = os.path.join(Applocation, 'Support\\basic_param\\Basic_Param_Set_ZJ.txt')
    basic_param_lst = read_file(path)
    Self_values_lst = basic_param_lst[0] # -1.0, 26.5, 2.0, 2.5

    # 文件操作大类=============================================================================================================

    # 获取base_url和reg_Key
    MidasURL()
    # 注册表路径
    reg_path2 = winreg.OpenKey(winreg.HKEY_CURRENT_USER,r"SOFTWARE\MIDAS\CVLwNX_CH\PATH")
    civilnx_path = winreg.QueryValueEx(reg_path2,"Installed Path")[0]
    # 窗口名
    partial_title = "MIDAS CIVIL NX"  # 你可以根据需要修改字符串
    
    # 变量设置
    Set_File_entry_var1 = tb.StringVar(value = Program_State_Monitoring(civilnx_path, partial_title)) # 程序状态
    Set_File_entry_var2 = tb.StringVar(value = "未选取任何模型") # 选取模型
    # Set_File_entry_var3 = tb.StringVar(value = "未选取任何路径") # 另存为路径
    # Set_File_entry_var3 = tb.StringVar(value = os.path.join(desktop_path, "Untitled_Cal.docx")) # 另存为路径
    Set_File_entry_var3 = tb.StringVar(value = os.path.join(Applocation, 'Custom', "ZJ_Untitled_Cal.docx")) # 另存为路径
    excel_path = os.path.join(Applocation, 'Custom\\Pile_Solid_Table.xlsx')

    # 标题
    dialog_headr(root, "文件操作")
    # 程序状态
    dialog_boxline3(root, "程序状态", 10, Set_File_entry_var1, 10, "打开Midas civil NX", lambda: Open_Midas_civil(civilnx_path, Set_File_entry_var1))
    # 模型名
    dialog_boxline3(root, "模型名称", 10, Set_File_entry_var2, 10, "选取mcb文件并打开", lambda: Open_Midas_mcb(Set_File_entry_var2))
    # 另存为路径
    dialog_boxline3(root, "保存路径", 10, Set_File_entry_var3, 10, "计算书另存为", lambda: Cal_path_SaveAs(Set_File_entry_var3, Applocation))

    # 模型操作大类=============================================================================================================

    # 一些用于储存对话框信息的变量
    mcb_combo_lst1 = ["小肋", "贝雷梁弦杆", "贝雷梁竖杆", "贝雷梁斜杆", "分配梁", "钢管桩", "联结系", "自定义"] # 结构组
    mcb_combo_lst2 = ["桩基础", "扩大基础", "承台", "自定义"] # 边界组
    NODE_dict_res = [[]] # 节点信息
    ELEM_dict_res = [[]] # 单元信息
    # 所有节点和单元
    mcb_entry_var1 = tb.StringVar(value="未读取节点和单元信息")# 编辑框的值
    # 结构组变量
    mcb_combo_var1 = tb.StringVar(value="")# 所选的下拉列表的值
    mcb_entry_var2 = tb.StringVar(value="")# 编辑框的值
    mcb_entry_var3 = tb.StringVar(value="/")# 编辑框的值
    # 边界组变量
    mcb_combo_var2 = tb.StringVar(value="")# 所选的下拉列表的值
    mcb_entry_var4 = tb.StringVar(value="")# 编辑框的值
    mcb_entry_var5 = tb.StringVar(value="/")# 编辑框的值

    # 标题
    dialog_headr(root, "手动添加必要的模型组(对于非自动化模型)")
    # 获取模型得节点和单元信息
    dialog_boxline3(root, "节点单元信息", 10, mcb_entry_var1, 10, "获取节点单元信息", lambda: get_node_elem_dict(mcb_entry_var1, NODE_dict_res, ELEM_dict_res))
    # 获取用户选择的单元号列表
    dialog_boxline1(root, "已选单元号", 10, mcb_entry_var2, 30, "选择单元并读取", lambda: get_node_ele_str(mcb_entry_var2, "ELEM"))
    # 结构组赋值
    dialog_boxline2(root, "结构组命名", 10, mcb_combo_var1, mcb_combo_lst1, 10, mcb_entry_var3, 15, "输入Midas", lambda: midas_group_output(mcb_combo_var1, mcb_entry_var3, NODE_dict_res, ELEM_dict_res, "Structure Group"))
    # 获取用户选择的节点号列表
    dialog_boxline1(root, "已选节点号", 10, mcb_entry_var4, 30, "选择节点并读取", lambda: get_node_ele_str(mcb_entry_var4, "NODE"))
    # 边界组赋值
    dialog_boxline2(root, "基础组命名", 10, mcb_combo_var2, mcb_combo_lst2, 10, mcb_entry_var5, 15, "输入Midas", lambda: midas_group_output(mcb_combo_var2, mcb_entry_var5, NODE_dict_res, ELEM_dict_res, "Boundary Group"))
    
    # 参数设定大类=============================================================================================================  
   
    # 二级对话框返回值
    Soil_Layer_Parameter_Table = {"土层分类":[], "土层厚度":[], "黏聚力":[], "内摩擦角":[], "侧摩阻":[], "端阻":[]} # 土层参数表信息
    foundation_secondary_dialog_values = {"浅基础":[], "深基础":[]} # 计算书需要用到的用户输入的信息
    # 下拉列表值
    # combo_lst1 = ["扩大基础", "钢管打入桩", "混凝土预制桩", "混凝土灌注桩"]
    combo_lst1 = ["根据模型信息", "扩大基础", "钢管打入桩", "混凝土灌注桩"]
    combo_lst2 = ["建筑地基基础设计规范"]
    combo_lst3 = ["建筑桩基技术规范"]
   
    # 参数预设定值
    var_lst1 = [2.0, 3.0, 0.5, 0.820, 26.5, -0.5, 20.0, 150.0, 0, 22.0, 100.0, 1.35, "C30", "HRB400", 50, 200, 200, 16, 18] # 扩大基础二级对话框
    # var_lst2 = [820, 10, 10, 1, {}] # 钢管打入桩二级对话框
    var_lst2 = [820, 10, 10, {}] # 钢管打入桩二级对话框
    # var_lst3 = [600, 110, "C30", 10, 1, {}] # 混凝土预制桩二级对话框
    var_lst3 = [600, 110, "C30", 10, {}] # 混凝土预制桩二级对话框
    # var_lst4 = [1000, "C30", "HRB400", 0.65, 50, 10, 1, {}] # 混凝土灌注桩二级对话框
    var_lst4 = [1000, "C30", "HRB400", 0.65, 50, 10, {}] # 混凝土灌注桩二级对话框
    var_dict1 = {"扩大基础":var_lst1, "钢管打入桩":var_lst2, "混凝土预制桩":var_lst3, "混凝土灌注桩":var_lst4}

    soleplate_secondary_dialog_values = {"模板":[]} # 计算书需要用到的用户输入的信息
    combo_lst4 = ['竹胶板+方木', '钢模板']
    var_lst5 = [15, '100mm×100mm', 3.5, 1.52, 200, 300, 'I10', 600, 2.5] # 竹胶板+方木参数
    var_lst6 = [10, 'I10', 600, 2.5] # 钢模板参数
    var_dict2 = {'竹胶板+方木':var_lst5, '钢模板':var_lst6}
    
    # 标题
    dialog_headr(root, "参数设定")
    # 基础
    combo0, button0 = dialog_boxline7(root, '模板结构', combo_lst4, 10, '设定参数', lambda: soleplate_secondary_dialog(root, button0, combo0.get(), soleplate_secondary_dialog_values, var_dict2, Self_values_lst))
    combo1, combo2, button1 = dialog_boxline4(root, "基础形式", "计算规范", combo_lst1, combo_lst2, combo_lst3, "设定参数", lambda: foundation_secondary_dialog(root, combo1.get(), combo2.get(), button1, foundation_secondary_dialog_values, var_dict1, Soil_Layer_Parameter_Table, excel_path))
    
    # 运行和关闭===============================================================================================================

    # dialog_boxline_end(root, "运行", lambda: print(secondary_dialog_values), "关闭", lambda: on_exit(root))
    dialog_boxline_end(root, "运行", lambda: Make_Doc(Applocation, model_post_result, Set_File_entry_var2.get(), soleplate_secondary_dialog_values, combo1.get(), combo2.get(), foundation_secondary_dialog_values, Set_File_entry_var3.get()), "关闭", lambda: on_exit(root))

    # 运行root
    # root.mainloop()

# if __name__ == "__main__":
def BaileySupport_Cal_Report_Main(parent=None, on_close=None):
    if not if_Reg(parent):
        if on_close:
            on_close()
        return
    if parent == None:
        root = tb.Window("钢管贝雷现浇支架计算书自动化助手")
    else:
        root = tb.Toplevel(parent)
        root.title("钢管贝雷现浇支架计算书自动化助手")
    creat_dialog(root)
    if parent == None:
        root.mainloop()
    elif on_close:
        def _on_root_destroy(event, _root=root):
            if event.widget is _root:
                on_close()
        root.bind('<Destroy>', _on_root_destroy)
#=====================================================================================================================================================================


# BaileySupport_Cal_Report_Main()

