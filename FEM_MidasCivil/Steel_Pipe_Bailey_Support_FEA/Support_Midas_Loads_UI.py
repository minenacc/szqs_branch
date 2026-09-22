# 1. 标准库
import tkinter as tk

# 2. 第三方库
import ttkbootstrap as tb
from ttkbootstrap.constants import *

# 3. 本地模块
from General.UIHandle import enable_button

# 风荷载计算对话框
# 对于桥梁抗风设计规范，风荷载的基本参数有基本风速U10，地表分类DBFL，地形条件系数kt，高度GD_Z，水平加载长度L，假定迎风贝雷排数n，设计基准风速计算公式
# 对于工程结构通用规范，风荷载的基本参数有基本风速v0，地表分类DBFL， 地形修正系数k，高度GD_Z，假定迎风贝雷排数n
# 对于港口工程荷载规范，风荷载的基本参数有平均最大风速V，地表分类DBFL，基本风压增大或降低系数k，高度GD_Z，假定迎风贝雷排数n

def create_Feng_dialog(root, standardname, dialog_values, button):
    if standardname != "自定义":
        # print(standardname)
        # 随规范变化的参数名称
        if standardname == "公路桥梁抗风设计规范":
            name_Feng = "基本风速U10(m/s)："
            name_K = "地形条件系数kt："
        elif standardname == "工程结构通用规范":
            name_Feng = "基本风速v0(m/s)："
            name_K = "地形修正系数k："
        elif standardname == "港口工程荷载规范":
            name_Feng = "平均最大风速V(m/s)："
            name_K = "基本风压增大或降低系数k："

        # 下拉列表
        DBFL_lst = ["A:海面、海岸、开阔水面", "B:田野、乡村、丛林、平坦开阔地", "C:树木及地层建筑密集区、平缓丘陵地", "D:中高层建筑密集区、起伏较大的丘陵地"]
        formula_lst = ["Ud = kf·kt·kh·U10", "Ud = kf·(Z/10)^α0·Us10"]

        # 检查dialog_values中是否已经存在过数值
        for key, key_str in dialog_values.items():
            # 基本风速
            if key == "wind_entry1":
                if key_str != "":
                    wind_entry1 = tk.StringVar(value=key_str)
                else:
                    wind_entry1 = tk.StringVar(value="24.5")
            # 地形修正
            if key == "wind_entry2":
                if key_str != "":
                    wind_entry2 = tk.StringVar(value=key_str)
                else:
                    wind_entry2 = tk.StringVar(value="1.0")
            # 计算高度
            if key == "wind_entry3":
                if key_str != "":
                    wind_entry3 = tk.StringVar(value=key_str)
                else:
                    wind_entry3 = tk.StringVar(value="20.0")
            # 水平加载长度
            if key == "wind_entry4":
                if key_str != "":
                    wind_entry4 = tk.StringVar(value=key_str)
                else:
                    wind_entry4 = tk.StringVar(value="50.0")
            # 迎风贝雷梁片数
            if key == "wind_entry5":
                if key_str != "":
                    wind_entry5 = tk.StringVar(value=key_str)
                else:
                    wind_entry5 = tk.StringVar(value="8")
            # 主梁横向力系数
            if key == "wind_entry6":
                if key_str != "":
                    wind_entry6 = tk.StringVar(value=key_str)
                else:
                    wind_entry6 = tk.StringVar(value="1.3")
            # 地表分类DBFL
            if key == "wind_combo1":
                if key_str != "":
                    wind_combo1 = tk.StringVar(value=key_str)
                else:
                    wind_combo1 = tk.StringVar()
            # 设计基准风速计算公式
            if key == "wind_combo2":
                if key_str != "":
                    wind_combo2 = tk.StringVar(value=key_str)
                else:
                    wind_combo2 = tk.StringVar()

        # root
        Feng_dialog = tk.Toplevel(root)
        Feng_dialog.title("风荷载参数表")
        # 基本风速
        frame_label_entry(Feng_dialog, name_Feng, wind_entry1)
        # 地表分类
        frame_combobox1(Feng_dialog, "地表分类：", DBFL_lst, wind_combo1)
        # 地形修正
        frame_label_entry(Feng_dialog, name_K, wind_entry2)
        # 计算高度
        frame_label_entry(Feng_dialog, "主梁基准高度(m)：", wind_entry3)
        if standardname == "公路桥梁抗风设计规范":
            # 水平加载长度
            frame_label_entry(Feng_dialog, "水平加载长度(m)：", wind_entry4)
            # 迎风贝雷梁片数
            frame_label_entry(Feng_dialog, "迎风贝雷梁片数：", wind_entry5)
            # 主梁横向力系数
            frame_label_entry(Feng_dialog, "主梁横向力系数：", wind_entry6)
            # 设计基准风速计算公式
            frame_combobox2(Feng_dialog, "设计基准风速计算公式", formula_lst, wind_combo2) 
        elif standardname == "工程结构通用规范" or standardname == "港口工程荷载规范":
            wind_entry4.set("None")
            # 迎风贝雷梁片数
            frame_label_entry(Feng_dialog, "迎风贝雷梁片数：", wind_entry5)
            wind_entry6.set("None")
            wind_combo1.set("None")
            wind_combo2.set("None")
        
        # 确认和取消
        Feng_frame_check_button(Feng_dialog, dialog_values, wind_entry1, wind_entry2, wind_entry3, wind_entry4, wind_entry5, wind_entry6, wind_combo1, wind_combo2, Feng_dialog, button)

        # 禁用主窗口的按钮
        button.config(state=tk.DISABLED)

        # 当二级窗口关闭时，重新启用主窗口的按钮
        Feng_dialog.protocol("WM_DELETE_WINDOW", lambda: enable_button(Feng_dialog, button))
    else:
        print("请选择风荷载规范")


# 风荷载计算对话框
# 对于港口工程荷载规范，水流力的基本参数为环境ENV，设计流速V，水位高程H

def create_Shui_dialog(root, standardname, dialog_values, button):
    if standardname != "自定义":
        # print(standardname)
        # 随规范变化的参数名称
        if standardname == "港口工程荷载规范":
            name_ENV = "水密度(t/m3)："
            name_V = "水流设计流速(m/s)："
            name_H = "水面相对高程(m)："

        # 检查dialog_values中是否已经存在过数值
        for key, key_str in dialog_values.items():
            # 水密度
            if key == "flow_entry1":
                if key_str != "":
                    flow_entry1 = tk.StringVar(value=key_str)
                else:
                    flow_entry1 = tk.StringVar(value="1.025")
            # 设计流速
            if key == "flow_entry2":
                if key_str != "":
                    flow_entry2 = tk.StringVar(value=key_str)
                else:
                    flow_entry2 = tk.StringVar(value="2.0")
            # 水面高程
            if key == "flow_entry3":
                if key_str != "":
                    flow_entry3 = tk.StringVar(value=key_str)
                else:
                    flow_entry3 = tk.StringVar(value="-5.0")

        # root
        Shui_dialog = tk.Toplevel(root)
        Shui_dialog.title("水流力参数表")
        # 水密度
        frame_label_entry(Shui_dialog, name_ENV, flow_entry1)
        # 设计流速
        frame_label_entry(Shui_dialog, name_V, flow_entry2)
        # 水面高程
        frame_label_entry(Shui_dialog, name_H, flow_entry3)
        
        # 确认和取消
        Shui_frame_check_button(Shui_dialog, dialog_values, flow_entry1, flow_entry2, flow_entry3, Shui_dialog, button)

        # 禁用主窗口的按钮
        button.config(state=tk.DISABLED)

        # 当二级窗口关闭时，重新启用主窗口的按钮
        Shui_dialog.protocol("WM_DELETE_WINDOW", lambda: enable_button(Shui_dialog, button))
    else:
        print("请选择水流力规范")

# 标签+编辑框
def frame_label_entry(dialog, txt, variable):
    # frame
    frame = tb.Frame(dialog)
    frame.pack(fill=tk.BOTH, expand=True)
    # 标签
    label = tb.Label(master=frame, text=txt.title(), width=20)
    label.pack(side=LEFT, padx=0, pady=10)
    # 编辑框
    entry = tb.Entry(master=frame, textvariable=variable, width=20)
    entry.pack(side=RIGHT, padx=10, pady=10, fill=X, expand=YES)

# 下拉列表1
def frame_combobox1(dialog, txt, lst, combo_var1):
    # frame
    frame = tb.Frame(dialog)
    frame.pack(fill=tk.BOTH, expand=True)
    # 标签
    label = tb.Label(master=frame, text=txt.title(), width=20)
    label.pack(side=LEFT, padx=0, pady=10)
    # 下拉列表
    combo1 = tb.Combobox(master=frame, textvariable=combo_var1)
    combo1.pack(side=RIGHT, padx=10, pady=10, fill=X)
    # 设置下拉列表的选项
    combo1['values'] = lst
    # 设置默认值
    if combo_var1.get() == "":
        combo1.current(0)

# 下拉列表2
def frame_combobox2(dialog, txt, lst, combo_var2):
    # frame
    frame = tb.Frame(dialog)
    frame.pack(fill=tk.BOTH, expand=True)
    # 标签
    label = tb.Label(master=frame, text=txt.title(), width=20)
    label.pack(side=LEFT, padx=0, pady=10)
    # 下拉列表
    combo2 = tb.Combobox(master=frame, textvariable=combo_var2)
    combo2.pack(side=RIGHT, padx=10, pady=10, fill=X)
    # 设置下拉列表的选项
    combo2['values'] = lst
    # 设置默认值
    if combo_var2.get() == "":
        combo2.current(0)

# 风荷载确认和取消
def Feng_frame_check_button(dialog, dialog_values, entry1, entry2, entry3, entry4, entry5, entry6, combo1, combo2, secondary_window, button):
    frame = tb.Frame(dialog)
    frame.pack(fill=X, expand=YES, pady=(15, 10))
    cnl_btn = tb.Button(master=frame,text="退出", command = lambda: enable_button(secondary_window, button))
    cnl_btn.pack(side=RIGHT, padx=5)
    sub_btn = tb.Button(master=frame,text="确认", command = lambda: Feng_on_submit(dialog_values, entry1, entry2, entry3, entry4, entry5, entry6, combo1, combo2))
    sub_btn.pack(side=RIGHT, padx=5)
    sub_btn.focus_set()

# 水流力确认和取消
def Shui_frame_check_button(dialog, dialog_values, entry1, entry2, entry3, secondary_window, button):
    frame = tb.Frame(dialog)
    frame.pack(fill=X, expand=YES, pady=(15, 10))
    cnl_btn = tb.Button(master=frame,text="退出", command = lambda: enable_button(secondary_window, button))
    cnl_btn.pack(side=RIGHT, padx=5)
    sub_btn = tb.Button(master=frame,text="确认", command = lambda: Shui_on_submit(dialog_values, entry1, entry2, entry3))
    sub_btn.pack(side=RIGHT, padx=5)
    sub_btn.focus_set()

# 风荷载提交
def Feng_on_submit(dialog_values, entry1, entry2, entry3, entry4, entry5, entry6, combo1, combo2):
    dialog_values["wind_entry1"] = entry1.get()
    dialog_values["wind_entry2"] = entry2.get()
    dialog_values["wind_entry3"] = entry3.get()
    dialog_values["wind_entry4"] = entry4.get()
    dialog_values["wind_entry5"] = entry5.get()
    dialog_values["wind_entry6"] = entry6.get()
    dialog_values["wind_combo1"] = combo1.get()
    dialog_values["wind_combo2"] = combo2.get()
    print("风荷载计算参数已保存")
    # print(entry1.get(), entry2.get(), entry3.get(), entry4.get(), entry5.get(), entry6.get(), combo1.get(), combo2.get())

# 水流力提交
def Shui_on_submit(dialog_values, entry1, entry2, entry3):
    dialog_values["flow_entry1"] = entry1.get()
    dialog_values["flow_entry2"] = entry2.get()
    dialog_values["flow_entry3"] = entry3.get()
    print("水流力计算参数已保存")
