import ttkbootstrap as ttk
import tkinter as tk
from tkinter import messagebox, filedialog
from ttkbootstrap.constants import *
import requests
from Cal_client import check_server_alive, submit_task, run_task, download_docx, task_in_progress

# 当前py文件对应总对话框函数

#=====================================================================================================================================================================

# 退出/关闭/结束程序
def on_exit(root):
    print("点击关闭按钮")
    if task_in_progress:
        messagebox.showwarning("提示", "任务正在进行中，请等待任务完成后再退出！")
        return
    else:
        if root.winfo_exists():  # 安全检查
            print("未进行任务，关闭窗口")
            root.destroy()  # 正确关闭方式


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

# 标题栏
def dialog_headr(root, txt):
    # 宋体加粗
    style = ttk.Style()
    style.configure("SimSun.TLabel", font=("SimSun", 10, "bold"))
    container = ttk.Frame(root)
    container.pack(fill=X, expand=YES)
    hdr = ttk.Label(master=root, text=txt, width=50, style="SimSun.TLabel")
    hdr.pack(fill=X, pady=10)


# 标签 + 编辑框 + 下拉列表
def dialog_boxline1(root, label_txt, label_length, entry_var, entry_length, button_txt, button_command):
    container = ttk.Frame(root)
    container.pack(fill=X, expand=YES, pady=5)
    # 标签
    Label = ttk.Label(master=container, text=label_txt.title(), width=label_length)
    Label.pack(side=LEFT, padx=0)
    # 编辑框
    ent = ttk.Entry(master=container, textvariable=entry_var, width=entry_length)
    ent.pack(side=LEFT, padx=20, fill=X, expand=YES)
    # 按钮
    but = ttk.Button(master=container, text=button_txt, command = button_command, width = 10)
    but.pack(side=RIGHT, padx=10)


# 标签 + 下拉列表 + 编辑框 + 按钮
def dialog_boxline2(root, label_txt1, label_length1, entry_var1, entry_length1, label_txt2, label_length2, combo_var, combo_lst, combo_length, entry_var2, entry_length2, button_txt, button_command):
    container = ttk.Frame(root)
    container.pack(fill=X, expand=YES, pady=5)
    # 标签1
    Label1 = ttk.Label(master=container, text=label_txt1.title(), width=label_length1)
    Label1.pack(side=LEFT, padx=0)
    # 编辑框1
    ent1 = ttk.Entry(master=container, textvariable=entry_var1, width=entry_length1)
    ent1.pack(side=LEFT, padx=20, fill=X, expand=YES)
    # 标签2
    Label2 = ttk.Label(master=container, text=label_txt2.title(), width=label_length2)
    Label2.pack(side=LEFT, padx=0)
    # 下拉列表
    combo = ttk.Combobox(master=container, textvariable=combo_var, values=combo_lst, width=combo_length)
    combo.pack(side=LEFT, padx=0)
    # 设置默认值
    combo.current(0)
    # 编辑框2
    ent2 = ttk.Entry(master=container, textvariable=entry_var2, width=entry_length2, state="disabled")
    ent2.pack(side=LEFT, padx=20, fill=X, expand=YES)
    # 将状态更新函数绑定到编辑框的<FocusIn>事件
    ent2.bind("<FocusIn>", lambda e : update_entry_state(combo_var, ent2))
    # 同时保留下拉列表的绑定
    combo.bind("<<ComboboxSelected>>", lambda e : update_entry_state(combo_var, ent2))
    # 按钮
    but = ttk.Button(master=container, text=button_txt, command = button_command, width = 10)
    but.pack(side=RIGHT, padx=10)


# 标签 + 边界框 + 按钮
def dialog_boxline3(root, label_txt, label_length, entry_var, entry_length, button_txt, button_command):
    container = ttk.Frame(root)
    container.pack(fill=X, expand=YES, pady=5)
    # 标签
    Label = ttk.Label(master=container, text=label_txt.title(), width=label_length)
    Label.pack(side=LEFT, padx=0)
    # 编辑框
    ent = ttk.Entry(master=container, textvariable=entry_var, width=entry_length, state="disabled")
    ent.pack(side=LEFT, padx=20, fill=X, expand=YES)
    # 按钮
    but = ttk.Button(master=container, text=button_txt, command = button_command, width = 15)
    but.pack(side=RIGHT, padx=10)


# 两个标签 + 两个下拉列表 + 一个按钮
def dialog_boxline4(root, label_txt1, label_txt2, combo_lst1, combo_lst2, combo_lst3, button_txt, button_command):
    container = ttk.Frame(root)
    container.pack(fill=X, expand=YES, pady=5)
    # 标签1
    ttk.Label(master=container, text=label_txt1.title()).pack(side="left", padx=5)
    # 第一个下拉列表
    combo1 = ttk.Combobox(master=container)
    combo1.pack(side=LEFT, padx=0, fill=X)
    combo1['values'] = combo_lst1
    combo1.current(0)
    # 标签2
    ttk.Label(master=container, text=label_txt2.title()).pack(side="left", padx=5)
    # 第二个下拉列表
    combo2 = ttk.Combobox(master=container)
    combo2.pack(side=LEFT, padx=0, fill=X)

    def update_second_combobox(event):
        if combo1.get() == "扩大基础":
            combo2['values'] = combo_lst2
        else:
            combo2['values'] = combo_lst3
        combo2.current(0)
        
    # 初始设置第二个下拉列表
    update_second_combobox(None)
    # 绑定事件
    combo1.bind("<<ComboboxSelected>>", update_second_combobox)
    # 按钮
    button1 = ttk.Button(master=container, text=button_txt, command = button_command, width = 10)
    button1.pack(side=RIGHT, padx=10)
    
    return combo1, combo2, button1


# 标签 + 编辑框
def dialog_boxline5(root, label_txt, variable):
    container = ttk.Frame(root)
    container.pack(fill=X, expand=YES, pady=5)
    ttk.Label(master=container, text=label_txt.title(), width=22).pack(side=LEFT, padx=30, pady=0, fill=X, expand=YES)
    ttk.Entry(master=container, textvariable=variable, width=20).pack(side=LEFT, padx=30, pady=0, fill=X, expand=YES)


# 标签 + 按钮
def dialog_boxline6(root, label_txt, button_txt, button_command):
    container = ttk.Frame(root)
    container.pack(fill=X, expand=YES, pady=5)
    ttk.Label(master=container, text=label_txt.title(), width=22).pack(side=LEFT, padx=30, pady=0, fill=X, expand=YES)
    button1 = ttk.Button(master=container, text=button_txt, command = button_command, width=19)
    button1.pack(side=RIGHT, padx=30)

# 土层参数输入表格
def open_soil_layer_editor(parent, Soil_Layer_Parameter_Table,target_window):
    win = tk.Toplevel(parent)
    win.title("土层参数表")
    win.geometry("1000x450")

    # 自定义样式：增强边框线
    style = ttk.Style()
    style.configure("Treeview", rowheight=28, font=("微软雅黑", 10))
    style.map("Treeview")

    data_columns = ["土层分类", "土层厚度(m)", "黏聚力(kPa)", "内摩擦角(°)", "侧摩阻(kPa)", "端阻(kPa)"]
    m_col = "水平抗力系数m"
    columns = ["#"] + data_columns + [m_col]  # 加上行号列

    tree = ttk.Treeview(win, columns=columns, show="headings", selectmode="browse")
    tree.pack(fill="both", expand=True)

    tree.heading("#", text="行号")
    tree.column("#", width=50, anchor="center")

    for col in data_columns:
        tree.heading(col, text=col)
        tree.column(col, width=130, anchor="center")
    
    tree.heading(m_col, text=m_col)
    tree.column(m_col, width=140, anchor="center")

    tree.pack(fill="both", expand=True)

    # 示例数据
    sample_data = [
        ["砂土", "3.0", "20", "30", "5", "3"],
        ["黏土", "2.5", "25", "28", "6", "4"],
    ]
    # 插入数据时自动添加行号
    for i, row_values in enumerate(sample_data, start=1):
        tree.insert("", "end", values=[str(i)] + row_values)

    # 当前 Entry 编辑器
    editor = tk.Entry(win)
    editor.place_forget()

    def update_row_numbers():
        for idx, item in enumerate(tree.get_children()):
            tree.set(item, "#", str(idx + 1))

    def on_double_click(event):
        nonlocal editor
        selected_item = tree.focus()
        if not selected_item:
            return

        region = tree.identify_region(event.x, event.y)
        if region != "cell":
            return

        column = tree.identify_column(event.x)
        col_index = int(column[1:]) - 1

        if col_index == 0 or col_index == len(columns) - 1:
            return  # 不允许编辑行号列及m值列

        x, y, width, height = tree.bbox(selected_item, column)
        value = tree.set(selected_item, columns[col_index])

        editor.delete(0, tk.END)
        editor.insert(0, value)
        editor.place(x=x, y=y + tree.winfo_y(), width=width, height=height)
        editor.focus_set()

        def save_edit(event=None):
            new_val = editor.get()
            tree.set(selected_item, columns[col_index], new_val)
            editor.place_forget()

        editor.bind("<Return>", save_edit)
        editor.bind("<FocusOut>", save_edit)

    tree.bind("<Double-1>", on_double_click)

    def apply_row_striping():
        for i, row in enumerate(tree.get_children()):
            tag = 'evenrow' if i % 2 == 0 else 'oddrow'
            tree.item(row, tags=(tag,))
        tree.tag_configure('evenrow', background="#f4f4f4")
        tree.tag_configure('oddrow', background="#ffffff")

    apply_row_striping()

    btn_frame = ttk.Frame(win)
    btn_frame.pack(pady=10, fill="x")

    def add_row():
        tree.insert("", "end", values=[""] * len(columns))
        update_row_numbers()
        apply_row_striping()

    def del_row():
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showwarning("提示", "请选择要删除的行")
            return
        tree.delete(selected_item)
        update_row_numbers()
        apply_row_striping()

    def save_and_close():
        key_map = {
            "土层分类": "土层分类",
            "土层厚度(m)": "土层厚度",
            "黏聚力(kPa)": "黏聚力",
            "内摩擦角(°)": "内摩擦角",
            "侧摩阻(kPa)": "侧摩阻",
            "端阻(kPa)": "端阻",
            # 不包含 "水平抗力系数m"
        }

        numeric_keys = {"土层厚度", "黏聚力", "内摩擦角", "侧摩阻", "端阻"}
        # 先清空字典内旧数据（只清data_columns对应键）
        for key in key_map.values():
            Soil_Layer_Parameter_Table[key] = []

        # 遍历Treeview所有行
        for item in tree.get_children():
            row_values = tree.item(item)["values"]
            for i, col_name in enumerate(columns[1:], start=1):  # 跳过行号列
                dict_key = key_map.get(col_name)
                if dict_key:
                    raw_val = row_values[i]
                    if dict_key in numeric_keys:
                        try:
                            val = float(raw_val)
                        except (ValueError, TypeError):
                            val = 0.0  # 或者你可以 raise Error 或者 messagebox 提醒
                    else:
                        val = raw_val
                    Soil_Layer_Parameter_Table[dict_key].append(val)

        messagebox.showinfo("保存成功", "已保存土层参数")
        win.destroy()

    def calculate_m_value():
    # 定义经验匹配表
        m_value_map = {
            "砂土": 15000,
            "中密砂": 20000,
            "密实砂": 30000,
            "黏土": 8000,
            "硬黏土": 15000,
            "软黏土": 5000,
            "碎石土": 25000,
            "粉土": 10000,
        }

        missing_count = 0

        for item in tree.get_children():
            vals = tree.item(item)["values"]
            soil_type = str(vals[1]).strip()  # 土层分类列索引 1

            # 尝试匹配经验值
            m_val = m_value_map.get(soil_type)

            if m_val is None:
                tree.set(item, m_col, "未知")
                missing_count += 1
            else:
                tree.set(item, m_col, f"{m_val:.0f}")  # 保留整数

        if missing_count > 0:
            messagebox.showwarning("提示", f"{missing_count} 行的“土层分类”未识别，m值已标为“未知”")
        else:
            messagebox.showinfo("计算完成", "所有 m 值已成功计算")

    ttk.Button(btn_frame, text="添加一行", command=add_row).grid(row=0, column=0, padx=10)
    ttk.Button(btn_frame, text="删除选中行", command=del_row).grid(row=0, column=1, padx=10)
    ttk.Button(btn_frame, text="计算m值", command=calculate_m_value).grid(row=0, column=2, padx=10)
    ttk.Button(btn_frame, text="保存并关闭", command=save_and_close).grid(row=0, column=3, padx=20)

# 最后一行
def dialog_boxline_end(root,):
    container = ttk.Frame(root)
    container.pack(fill=X, expand=YES, pady=10)

    btn1 = ttk.Button(container, text="关闭", width=6)
    btn1.pack(side=RIGHT, padx=10, pady=10)

    btn2 = ttk.Button(container, text="下载", bootstyle=SUCCESS, width=6)
    btn2.pack(side=RIGHT, padx=10, pady=10)

    btn3 = ttk.Button(container, text="计算", bootstyle=SUCCESS, width=6)
    btn3.pack(side=RIGHT, padx=10, pady=10)

    btn4 = ttk.Button(container, text="上传", bootstyle=SUCCESS,width=6)
    btn4.pack(side=RIGHT, padx=10, pady=10)
    return btn1, btn2, btn3, btn4
    
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
    frame = ttk.Frame(dialog)
    frame.pack(fill=X, expand=YES, pady=(15, 10))
    # 退出按钮
    cnl_btn = ttk.Button(master=frame,text="退出", command = lambda: enable_button(dialog, button))
    cnl_btn.pack(side=RIGHT, padx=30)
    # 确认按钮
    # sub_btn = ttk.Button(master=frame,text="确认", command = lambda: [dialog_values.__setitem__(dict_key, [var.get() for var in var_lst])])
    sub_btn = ttk.Button(master=frame,text="确认", command = lambda: secondary_dialog_set_save_params(dialog_values, var_dict, var_lst, dict_key, foundation_name))
    sub_btn.pack(side=RIGHT, padx=5)
    sub_btn.focus_set()
    return cnl_btn, sub_btn

# 二级对话框确认时，对主对话框传进来的变量赋值，保存二级对话框的参数
def secondary_dialog_set_save_params(dialog_values, var_dict, var_lst, dict_key, foundation_name):
    # 对主对话框传进来的变量赋值
    dialog_values.__setitem__(dict_key, [var.get() if isinstance(var, (ttk.StringVar, ttk.IntVar, ttk.DoubleVar, ttk.BooleanVar)) else var for var in var_lst])
    # 保存二级对话框的参数
    var_dict[foundation_name] = [var.get() if isinstance(var, (ttk.StringVar, ttk.IntVar, ttk.DoubleVar, ttk.BooleanVar)) else var for var in var_lst]

# 重新启用主窗口的按钮，并销毁二级窗口
def enable_button(secondary_window, button):
    """重新启用主窗口的按钮，并销毁二级窗口"""
    button.config(state=ttk.NORMAL)
    secondary_window.destroy()

# 浅基础的参数设定二级对话框
def Shallow_Foundation_secondary_dialog(root, button, secondary_dialog_values, var_dict, dict_key):
    secondary_dialog = ttk.Toplevel(root)
    secondary_dialog.title("浅基础参数设定框")
    # 基础的长
    foundation_x = ttk.DoubleVar(value = var_dict["扩大基础"][0])
    # 基础的宽
    foundation_y = ttk.DoubleVar(value = var_dict["扩大基础"][1])
    # 基础的高
    foundation_z = ttk.DoubleVar(value = var_dict["扩大基础"][2])
    # 基础上立柱计算直径
    Pile_D = ttk.DoubleVar(value = var_dict["扩大基础"][3])
    # 混凝土的容重
    concrete_gc = ttk.DoubleVar(value = var_dict["扩大基础"][4])
    # 基础底标高
    Foundation_Base_Level = ttk.DoubleVar(value = var_dict["扩大基础"][5])
    # 基础上土体平均容重
    soild_gs = ttk.DoubleVar(value = var_dict["扩大基础"][6])
    # 修正后地基承载力
    foundation_fa = ttk.DoubleVar(value = var_dict["扩大基础"][7])
    # 软弱下卧层顶标高
    Weak_Layer_Horizon = ttk.DoubleVar(value = var_dict["扩大基础"][8])
    # 软弱下卧层以上土体扩散角
    Dispersion_Angle = ttk.DoubleVar(value = var_dict["扩大基础"][9])
    # 软弱下卧层承载力
    foundation_faz = ttk.DoubleVar(value = var_dict["扩大基础"][10])
    # 标准组合和基本组合之间的转换系数
    ksi = ttk.DoubleVar(value = var_dict["扩大基础"][11])
    # 混凝土等级
    concrete_grade = ttk.StringVar(value = var_dict["扩大基础"][12])
    # 钢筋牌号
    rebar_grade = ttk.StringVar(value = var_dict["扩大基础"][13])
    # 钢筋保护层厚度
    as0 = ttk.DoubleVar(value = var_dict["扩大基础"][14])
    # 钢筋间距范围最小值
    steel_s1 = ttk.DoubleVar(value = var_dict["扩大基础"][15])
    # 钢筋间距范围最大值
    steel_s2 = ttk.DoubleVar(value = var_dict["扩大基础"][16])
    # 钢筋直径范围最小值
    steel_d1 = ttk.DoubleVar(value = var_dict["扩大基础"][17])
    # 钢筋直径范围最大值
    steel_d2 = ttk.DoubleVar(value = var_dict["扩大基础"][18])
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
    button.config(state=ttk.DISABLED)
    # 当二级窗口关闭时，重新启用主窗口的按钮
    secondary_dialog.protocol("WM_DELETE_WINDOW", lambda: enable_button(secondary_dialog, button))


# 深基础的参数设定二级对话框——钢管打入桩
def Deep_Foundation_secondary_dialog_steelpile(root, button, secondary_dialog_values, var_dict, Soil_Layer_Parameter_Table, dict_key):
    secondary_dialog = ttk.Toplevel(root)
    secondary_dialog.title("深基础参数设定框")
    # 钢管桩直径
    Pile_D = ttk.DoubleVar(value = var_dict["钢管打入桩"][0])
    # 钢管桩壁厚
    Pile_t = ttk.DoubleVar(value = var_dict["钢管打入桩"][1])
    # 水平位移值
    Soil_vb = ttk.DoubleVar(value = var_dict["钢管打入桩"][2])
    # 水平抗力计算土层
    Soil_Layer_n = ttk.IntVar(value = var_dict["钢管打入桩"][3])
    # 土层参数表的值
    # Soil_Layer_Parameter_Table = {"土层分类":[], "土层厚度":[], "黏聚力":[], "内摩擦角":[], "侧摩阻":[], "端阻":[]}
    # 对话框内容
    dialog_headr(secondary_dialog, "钢管打入桩参数")
    dialog_boxline5(secondary_dialog, "钢管桩直径(mm)", Pile_D)
    dialog_boxline5(secondary_dialog, "钢管桩壁厚(mm)", Pile_t)
    dialog_boxline5(secondary_dialog, "水平位移值(mm)", Soil_vb)
    dialog_boxline6(secondary_dialog, "输入土层参数表", "打开编辑窗口", lambda: open_soil_layer_editor(secondary_dialog, Soil_Layer_Parameter_Table, secondary_dialog))
    dialog_boxline5(secondary_dialog, "水平抗力计算土层", Soil_Layer_n)
    # 确认和退出，将参数从二级对话框返回给主对话框
    var_lst = [Pile_D, Pile_t, Soil_vb, Soil_Layer_n, Soil_Layer_Parameter_Table]
    secondary_dialog_submit_quit(secondary_dialog, button, secondary_dialog_values, var_dict, var_lst, dict_key, "钢管打入桩")
    # 禁用主窗口的按钮
    button.config(state=ttk.DISABLED)
    # 当二级窗口关闭时，重新启用主窗口的按钮
    secondary_dialog.protocol("WM_DELETE_WINDOW", lambda: enable_button(secondary_dialog, button))


# # 深基础的参数设定二级对话框——混凝土预制桩
# def Deep_Foundation_secondary_dialog_precastpile(root, button, secondary_dialog_values, var_dict, Soil_Layer_Parameter_Table, dict_key):
#     secondary_dialog = ttk.Toplevel(root)
#     secondary_dialog.title("深基础参数设定框")
#     # 预制桩直径
#     Pile_D = ttk.DoubleVar(value = var_dict["混凝土预制桩"][0])
#     # 预制桩壁厚
#     Pile_t = ttk.DoubleVar(value = var_dict["混凝土预制桩"][1])
#     # 混凝土等级
#     concrete_grade = ttk.StringVar(value = var_dict["混凝土预制桩"][2])
#     # 水平位移值
#     Soil_vb = ttk.DoubleVar(value = var_dict["混凝土预制桩"][3])
#     # 水平抗力计算土层
#     Soil_Layer_n = ttk.IntVar(value = var_dict["混凝土预制桩"][4])
#     # 土层参数表的值
#     # Soil_Layer_Parameter_Table = {"土层分类":[], "土层厚度":[], "黏聚力":[], "内摩擦角":[], "侧摩阻":[], "端阻":[]}
#     # 对话框内容
#     dialog_headr(secondary_dialog, "混凝土预制桩参数")
#     dialog_boxline5(secondary_dialog, "预制桩直径(mm)", Pile_D)
#     dialog_boxline5(secondary_dialog, "预制桩壁厚(mm)", Pile_t)
#     dialog_boxline5(secondary_dialog, "混凝土等级", concrete_grade)
#     dialog_boxline5(secondary_dialog, "水平位移值(mm)", Soil_vb)
#     dialog_boxline6(secondary_dialog, "输入土层参数表", "打开编辑窗口", lambda: open_soil_layer_editor(secondary_dialog, Soil_Layer_Parameter_Table, secondary_dialog))
#     dialog_boxline5(secondary_dialog, "水平抗力计算土层", Soil_Layer_n)
#     # 确认和退出，将参数从二级对话框返回给主对话框
#     var_lst = [Pile_D, Pile_t, concrete_grade, Soil_vb, Soil_Layer_n, Soil_Layer_Parameter_Table]
#     secondary_dialog_submit_quit(secondary_dialog, button, secondary_dialog_values, var_dict, var_lst, dict_key, "混凝土预制桩")
#     # 禁用主窗口的按钮
#     button.config(state=ttk.DISABLED)
#     # 当二级窗口关闭时，重新启用主窗口的按钮
#     secondary_dialog.protocol("WM_DELETE_WINDOW", lambda: enable_button(secondary_dialog, button))


# 深基础的参数设定二级对话框——混凝土灌注桩
def Deep_Foundation_secondary_dialog_concretepile(root, button, secondary_dialog_values, var_dict, Soil_Layer_Parameter_Table, dict_key):
    secondary_dialog = ttk.Toplevel(root)
    secondary_dialog.title("深基础参数设定框")
    # 灌注桩直径
    Pile_D = ttk.DoubleVar(value = var_dict["混凝土灌注桩"][0])
    # 混凝土等级
    concrete_grade = ttk.StringVar(value = var_dict["混凝土灌注桩"][1])
    # 钢筋等级
    rebar_grade = ttk.StringVar(value = var_dict["混凝土灌注桩"][2])
    # 桩身正截面配筋率
    Pile_p = ttk.DoubleVar(value = var_dict["混凝土灌注桩"][3])
    # 钢筋保护层厚度
    Pile_as = ttk.DoubleVar(value = var_dict["混凝土灌注桩"][4])
    # 水平位移值
    Soil_vb = ttk.DoubleVar(value = var_dict["混凝土灌注桩"][5])
    # 水平抗力计算土层
    Soil_Layer_n = ttk.IntVar(value = var_dict["混凝土灌注桩"][6])
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
    dialog_boxline6(secondary_dialog, "输入土层参数表", "打开编辑窗口", lambda: open_soil_layer_editor(secondary_dialog, Soil_Layer_Parameter_Table, secondary_dialog))
    dialog_boxline5(secondary_dialog, "水平抗力计算土层", Soil_Layer_n)
    # 确认和退出，将参数从二级对话框返回给主对话框
    var_lst = [Pile_D, concrete_grade, rebar_grade, Pile_p, Pile_as, Soil_vb, Soil_Layer_n, Soil_Layer_Parameter_Table]
    secondary_dialog_submit_quit(secondary_dialog, button, secondary_dialog_values, var_dict, var_lst, dict_key, "混凝土灌注桩")
    # 禁用主窗口的按钮
    button.config(state=ttk.DISABLED)
    # 当二级窗口关闭时，重新启用主窗口的按钮
    secondary_dialog.protocol("WM_DELETE_WINDOW", lambda: enable_button(secondary_dialog, button))


# 基础参数设定的二级对话框
def secondary_dialog(root, combo_var, button, secondary_dialog_values, var_dict, Soil_Layer_Parameter_Table):
    if combo_var == "扩大基础":
        Shallow_Foundation_secondary_dialog(root, button, secondary_dialog_values, var_dict, "浅基础")
    elif combo_var == "钢管打入桩":
        Deep_Foundation_secondary_dialog_steelpile(root, button, secondary_dialog_values, var_dict, Soil_Layer_Parameter_Table, "深基础")
    # elif combo_var == "混凝土预制桩":
    #     Deep_Foundation_secondary_dialog_precastpile(root, button, secondary_dialog_values, var_dict, Soil_Layer_Parameter_Table, "深基础")
    elif combo_var == "混凝土灌注桩":
        Deep_Foundation_secondary_dialog_concretepile(root, button, secondary_dialog_values, var_dict, Soil_Layer_Parameter_Table, "深基础")

# 获取excel的数据
def get_excel_data_tolst(Soil_Layer_Parameter_Table):
    # 连接对应表
    openfile_path = qucik_open_file_dialog()
    folder_path = os.path.dirname(openfile_path) # 文件目录路径
    excel_path = os.path.join(folder_path, '土层参数表.xlsx')
    df = pd.read_excel(excel_path, sheet_name='Sheet1')
    # 生成计算书所需要的参数
    column_data_1 = df['土层分类']
    column_data_2 = df['土层厚度(m)']
    column_data_3 = df['黏聚力(kPa)']
    column_data_4 = df['内摩擦角(°)']
    column_data_5 = df['侧摩阻(kPa)']
    column_data_6 = df['端阻(kPa)']
    # 赋值
    Soil_type_lst = column_data_1.tolist() # 土层分类
    Soil_Layer_Hlst = column_data_2.tolist() # 土层厚度
    Soil_Cohesion_lst = column_data_3.tolist() # 黏聚力
    Soil_friction_lst = column_data_4.tolist() # 内摩擦角
    Soil_qsk_lst = column_data_5.tolist() # 侧摩阻
    Soil_psk_lst = column_data_6.tolist() # 端阻

    Soil_Layer_Parameter_Table["土层分类"] = Soil_type_lst
    Soil_Layer_Parameter_Table["土层厚度"] = Soil_Layer_Hlst
    Soil_Layer_Parameter_Table["黏聚力"] = Soil_Cohesion_lst
    Soil_Layer_Parameter_Table["内摩擦角"] = Soil_friction_lst
    Soil_Layer_Parameter_Table["侧摩阻"] = Soil_qsk_lst
    Soil_Layer_Parameter_Table["端阻"] = Soil_psk_lst

    print(Soil_Layer_Parameter_Table)

# 打开Midas模型文件
# 选择文件
def qucik_open_file_dialog():
    # 获取文件完整路径
    # root = tk.Tk()
    # root.withdraw() # 隐藏主窗口
    FileName = filedialog.askopenfilename() # 打开文件对话框,选择打开什么文件，返回文件名
    # root.destroy()# 结束程序
    return FileName
# print(qucik_open_file_dialog())

def Open_Midas_mcb(entry_var):
    # 选择路径
    openfile_path = qucik_open_file_dialog()
    # 更新对应编辑框的值
    entry_var.set(openfile_path)
    print(f"你选取的路径为：{entry_var.get()}")
    # # 文件目录路径
    # folder_path = os.path.dirname(openfile_path)
    # print(folder_path)
    # # 文件名
    # folder_name = os.path.basename(openfile_path)
    # print(folder_name)
    # 打开模型
    if "mcb" in openfile_path:
        arguments = {"Argument" : openfile_path}
    else:
        print("未识别到有效mcb文件, 请重新选择")

def upload_txt_file(entry_var):
    # 选择路径
    openfile_path = qucik_open_file_dialog()
    # 更新对应编辑框的值
    entry_var.set(openfile_path)
    print(f"你选取的路径为：{entry_var.get()}")
    # # 文件目录路径
    # folder_path = os.path.dirname(openfile_path)
    # print(folder_path)
    # # 文件名
    # folder_name = os.path.basename(openfile_path)
    # print(folder_name)
    # 打开模型
    if "txt" in openfile_path:
        arguments = {"Argument" : openfile_path}
    else:
        print("未识别到有效txt文件, 请重新选择")

#===========================================================================================================================================
        
# 对话框总函数
def creat_dialog(root):
    root.title("计算书自动化程序")
    root.deiconify()
    # 文件操作大类=============================================================================================================

    # 变量设置
    Set_File_entry_var2 = ttk.StringVar(value = "未选取任何模型") # 选取模型

    # 标题
    dialog_headr(root, "文件操作")
 
    # 模型名
    dialog_boxline3(root, "模型名称", 8, Set_File_entry_var2, 10, "选取Midas模型", lambda: Open_Midas_mcb(Set_File_entry_var2))

    # 选取参数文件（txt）
    Set_txt_entry_var = ttk.StringVar(value="未选取任何参数文件")
    dialog_boxline3(root, "参数文件", 8, Set_txt_entry_var, 10, "选取参数txt", lambda: upload_txt_file(Set_txt_entry_var))

    # 参数设定大类=============================================================================================================
    
    # # 二级对话框返回值
    # Soil_Layer_Parameter_Table = {"土层分类":[], "土层厚度":[], "黏聚力":[], "内摩擦角":[], "侧摩阻":[], "端阻":[]} # 土层参数表信息
    # secondary_dialog_values = {"浅基础":[], "深基础":[]} # 计算书需要用到的用户输入的信息
    # # 下拉列表值
    # combo_lst1 = ["扩大基础", "钢管打入桩", "混凝土预制桩", "混凝土灌注桩"]
    # combo_lst2 = ["建筑地基基础设计规范"]
    # combo_lst3 = ["建筑桩基技术规范"]
    # # 参数预设定值
    # var_lst1 = [2.0, 3.0, 0.5, 0.820, 26.5, -0.5, 20.0, 150.0, 0, 22.0, 100.0, 1.35, "C30", "HRB400", 50, 200, 200, 16, 18] # 扩大基础二级对话框
    # var_lst2 = [820, 10, 10, 1, {}] # 钢管打入桩二级对话框
    # var_lst3 = [600, 110, "C30", 10, 1, {}] # 混凝土预制桩二级对话框
    # var_lst4 = [1000, "C30", "HRB400", 0.65, 50, 10, 1, {}] # 混凝土灌注桩二级对话框
    # var_dict = {"扩大基础":var_lst1, "钢管打入桩":var_lst2, "混凝土预制桩":var_lst3, "混凝土灌注桩":var_lst4}

    # 二级对话框返回值
    Soil_Layer_Parameter_Table = {"土层分类":[], "土层厚度":[], "黏聚力":[], "内摩擦角":[], "侧摩阻":[], "端阻":[]} # 土层参数表信息
    secondary_dialog_values = {"浅基础":[], "深基础":[]} # 计算书需要用到的用户输入的信息
    # 下拉列表值
    combo_lst1 = ["扩大基础", "钢管打入桩", "混凝土灌注桩"]
    combo_lst2 = ["建筑地基基础设计规范"]
    combo_lst3 = ["建筑桩基技术规范"]
    # 参数预设定值
    var_lst1 = [2.0, 3.0, 0.5, 0.820, 26.5, -0.5, 20.0, 150.0, 0, 22.0, 100.0, 1.35, "C30", "HRB400", 50, 200, 200, 16, 18] # 扩大基础二级对话框
    var_lst2 = [820, 10, 10, 1, {}] # 钢管打入桩二级对话框
    var_lst4 = [1000, "C30", "HRB400", 0.65, 50, 10, 1, {}] # 混凝土灌注桩二级对话框
    var_dict = {"扩大基础":var_lst1, "钢管打入桩":var_lst2,  "混凝土灌注桩":var_lst4}

    # 标题
    dialog_headr(root, "参数设定") 
    # 基础
    combo1, combo2, button = dialog_boxline4(root, "基础形式", "计算规范", combo_lst1, combo_lst2, combo_lst3, "设定参数", lambda: secondary_dialog(root, combo1.get(), button, secondary_dialog_values, var_dict, Soil_Layer_Parameter_Table))
    
    # 运行和关闭===============================================================================================================

    # 绑定最后一行按钮（因task_id需要传递运行，故分开编写）
    # 创建按钮
    btn_close, btn_download, btn_run, btn_upload = dialog_boxline_end(root)

    # 定义task_id存储容器
    task_id_holder = {"value": None}
    # 绑定上传按钮：上传后存储 task_id
    def upload_action():
        global task_id
        task_id = submit_task(Set_File_entry_var2.get(), Set_txt_entry_var.get(),combo1.get(),combo2.get(), secondary_dialog_values)
        if task_id:
            task_id_holder["value"] = task_id  # 保存 task_id 给后续按钮使用

    # 绑定计算按钮
    def run_action():
        task_id = task_id_holder.get("value")
        if task_id:
            run_task(task_id)
        else:
            messagebox.showerror("错误", "请先上传模型文件")

    # 绑定下载按钮
    def download_action():
        task_id = task_id_holder.get("value")
        if task_id:
            download_docx(task_id)
        else:
            messagebox.showerror("错误", "请先上传模型文件")

    # 绑定关闭按钮
    def close_action():
        on_exit(root)

    # 实际绑定
    btn_upload.config(command=upload_action)
    btn_run.config(command=run_action)
    btn_download.config(command=download_action)
    btn_close.config(command=close_action)

    # 运行root
    root.mainloop()

#=====================================================================================================================================================================


# 启动程序
if __name__ == "__main__":
    try:
        check_server_alive(server_url = "http://10.142.35.199:8000" )
        creat_dialog(root)
    except Exception as e:
        import traceback
        with open("error.log", "w", encoding="utf-8") as f:
            f.write("程序异常退出:\n")
            traceback.print_exc(file=f)
        raise
