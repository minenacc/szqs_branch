# 1. 标准库
import os
import winreg
import tkinter as tk

# 2. 第三方库
import ttkbootstrap as tb
from ttkbootstrap.constants import *

# 3. 本地模块
from General.TextHandle     import read_file
from General.ExcelHandle    import write_csv
from General.Geometry       import get_distance_from_xlst
from General.ProgramMonitor import Program_State_Monitoring
from General.UIHandle       import createToolTip, enable_button
from General.Midas          import MidasURL, Open_Midas_civil, importmct_to_mcb
from General.DataUtils      import update_or_append, midasdisttolst, SET_CLIP_STRING
from General.AutoCAD        import SelectOnScreen, get_block_reference_xdata, get_SECinfo_fromcad
from General.FilePath       import write_file_within_path, file_extension_Modified, qucik_save_file_dialog
from General.Formula        import cal_FengYa_JTG_T_3360_01_2018, cal_FengYa_GB_55001_2021, cal_FengYa_JTS_144_1_2010

from FEM_MidasCivil.Steel_Pipe_Bailey_Support_FEA.Support_Midas_Model_Generate import (
    get_heightlst_fromcad, make_Bailey_origin_points, Cross_Bailey_Rib, draw_Bailey_incad, creat_Bailey_node_element_group_brnd,
    FPL_section, draw_FPL_incad, Reg_FPL_PM, Cross_Bailey_FPL, update_FPLdict_with_SP, creat_FPL_node_element_group_brnd, get_sp_info_fromcad, draw_SP1_incad, 
    create_SP_LJX_node_element_grup_brnd, rib_sections, get_rib_basedata, make_rib_nodes_elements_loads, draw_rib_blocks_incad
)
from FEM_MidasCivil.Steel_Pipe_Bailey_Support_FEA.Support_Utils_function import support_section_data, get_point_and_make_xline, group_consecutive_elements, loadcomb, MATERIAL


class DataEntryForm(tb.Frame):
    def __init__(self, master, Applocation, **kwargs):

        self.root = tb.Frame(master)
        
        path = os.path.join(Applocation, 'Support\\basic_param\\Basic_Param_Set_ZJ.txt')
        basic_param_lst = read_file(path)
        Self_values_lst = basic_param_lst[0] # -1.0, 26.5, 2.0, 2.5
        FengGui_selected_value_lst = basic_param_lst[1] # ["自定义"] # 选择规范默认值
        ShuiGui_selected_value_lst = basic_param_lst[3] # ["自定义"] # 选择规范默认值
        Feng_values_lst = basic_param_lst[2] # ['24.5', '1.0', '20', '50', '8', '1.3', 'A:海面、海岸、开阔水面', 'Ud = kf·kt·kh·U10']
        Shui_values_lst = basic_param_lst[4] # ['1.025', '2.0', '-5.0']

        self.entryname13 = tb.StringVar(value=Self_values_lst[0]) # 自重系数
        self.entryname12 = tb.StringVar(value=Self_values_lst[1]) # 混凝土容重
        self.entryname10 = tb.StringVar(value=Self_values_lst[2]) # 施工荷载
        self.entryname11 = tb.StringVar(value=Self_values_lst[3]) # 模板荷载
        
        self.entryname1 = tb.StringVar(value="")
        self.entryname2 = tb.StringVar(value="")
        self.entryname3 = tb.StringVar(value="820") # 钢管桩直径
        self.entryname4 = tb.StringVar(value="10") # 钢管桩壁厚
        self.entryname5 = tb.StringVar(value="12000")# 钢管桩水平投影长度(桩顶桩底z坐标差值)
        self.entryname6 = tb.StringVar(value="")
        self.entryname7 = tb.StringVar(value="")
        self.entryname8 = tb.StringVar(value="")
        self.entryname9 = tb.StringVar(value="21@1000") # 小肋数量及间距
        # self.entryname10 = tb.StringVar(value="2.0") # 施工荷载
        # self.entryname11 = tb.StringVar(value="2.5") # 模板荷载
        # self.entryname12 = tb.StringVar(value="26.5") # 混凝土荷载
        # self.entryname13 = tb.StringVar(value="-1") # 自重系数
        self.entryname14 = tb.StringVar(value="")
        self.entryname15 = tb.StringVar(value="1000") # 小肋悬臂长度
        self.entryname16 = tb.StringVar(value="")
        self.entryname17 = tb.StringVar(value="")
        self.entryname18 = tb.StringVar(value="")
        self.entryname19 = tb.StringVar(value="0") # 小肋横桥向起始坐标

        self.User_Orignal_Point_inCAD = []#用户选取的坐标原点

        self.selected_value1 = tb.StringVar()#分配梁的选择值
        self.selected_value2 = tb.StringVar()#小肋的选择值

        self.current_row = 0#行数的初始值
        self.ribstartnum = 1#选中的小肋的起点序号
        self.ribendnum = 1#选中的小肋的终点序号

        self.section_xlst1 = []#截面所有直线端点、曲线端点中点、贝雷梁插入点的x坐标
        self.section_xlst2 = []#截面下对应的贝雷梁和截面端点、圆弧中点的x坐标
        self.section_hlst = []#左右微小移动的x坐标对应的截面h
        self.secetyname_lst = []#二次对话框中用户输入或识别到的文本
        self.get_rib_basedata_resultlist = []
        self.rib_Bailey_y_dict = {}#小肋与贝雷梁弹性连接的y坐标词典

        self.Bailey_dict = []#贝雷梁词典
        self.Distance_between_section_Bailey_Ry_lst = []#记录每一组截面距离贝雷梁的边线的距离
        # self.Bailey_y_dist_lst = []#贝雷梁纵向布置
        self.SEC_basic_info_dict = {}#截面基本信息词典
        self.Sec_info_states = []#截面录入状态——未录入，已录入，1

        self.rib_lst = [] # 小肋截面信息——sec文件读取结果
        self.FPL_lst = [] # 分配梁截面信息——sec文件读取结果

        # self.FengGui_selected_value = "自定义"# 选择规范默认值
        # self.ShuiGui_selected_value = "自定义"# 选择规范默认值
        # self.Feng_values = {"wind_entry1":"", "wind_entry2":"", "wind_entry3":"", "wind_entry4":"", "wind_entry5":"", "wind_entry6":"", "wind_combo1":"", "wind_combo2":""}# 风荷载参数词典
        # self.Shui_values = {"flow_entry1":"", "flow_entry2":"", "flow_entry3":""}# 水流力参数词典
        self.FengGui_selected_value = FengGui_selected_value_lst[0]# 选择规范默认值
        self.ShuiGui_selected_value = ShuiGui_selected_value_lst[0]# 选择规范默认值
        self.Feng_values = {"wind_entry1":Feng_values_lst[0], "wind_entry2":Feng_values_lst[1], "wind_entry3":Feng_values_lst[2], "wind_entry4":Feng_values_lst[3], 
                            "wind_entry5":Feng_values_lst[4], "wind_entry6":Feng_values_lst[5], "wind_combo1":Feng_values_lst[6], "wind_combo2":Feng_values_lst[7]}# 风荷载参数词典
        self.Shui_values = {"flow_entry1":Shui_values_lst[0], "flow_entry2":Shui_values_lst[1], "flow_entry3":Shui_values_lst[2]}# 水流力参数词典


        globals()[f"sec_entryname{self.current_row}"] = tb.StringVar(value="")
        super().__init__(master, padding=(20, 10), **kwargs)
        self.pack(fill=BOTH, expand=YES)

        # self.mct_savepath = tb.StringVar(value = os.path.join(get_desktop_path_from_registry(), "Untitled.mct")) # mct保存路径
        self.reg_path = winreg.OpenKey(winreg.HKEY_CURRENT_USER,r"SOFTWARE\MIDAS\CVLwNX_CH\PATH")
        self.civilnx_path = winreg.QueryValueEx(self.reg_path,"Installed Path")[0]
        self.Set_File_entry_var = tb.StringVar(value = Program_State_Monitoring(self.civilnx_path, "MIDAS CIVIL NX")) # 程序状态
        self.mct_savepath = tb.StringVar(value = os.path.join(Applocation, r"Custom\Support_Untitled.mct")) # mct保存路径

    # 标头栏
    def headr_line(self, txt):
        # 宋体加粗
        style = tb.Style()
        style.configure("SimSun.TLabel", font=("SimSun", 10, "bold"))
        container = tb.Frame(self)
        container.pack(fill=X, expand=YES, pady=10)
        hdr = tb.Label(master=self, text=txt, width=50, style="SimSun.TLabel")
        hdr.pack(fill=X, padx=10, pady=0)

    # 创建1个标签、1个编辑框
    def Label_Entry_line(self, label, variable, length, tips):
        container = tb.Frame(self)
        container.pack(fill=X, expand=YES, pady=10)
        lbl = tb.Label(master=container, text=label.title(), width=length)
        lbl.pack(side=LEFT, padx=10)
        ent = tb.Entry(master=container, textvariable=variable)
        ent.pack(side=RIGHT, padx=10, fill=X, expand=YES)
        createToolTip(ent, tips)

    # 创建2个标签、2个编辑框
    def Two_Label_Two_Entry_line(self, label1, variable1, length1, label2, variable2, length2, tips):
        container = tb.Frame(self)
        container.pack(fill=X, expand=YES, pady=10)
        lbl1 = tb.Label(master=container, text=label1.title(), width=length1)
        lbl1.pack(side=LEFT, padx=10)
        ent1 = tb.Entry(master=container, textvariable=variable1, width=10)
        ent1.pack(side=LEFT, padx=30, fill=X, expand=YES)
        # createToolTip(ent1, tips)
        ent2 = tb.Entry(master=container, textvariable=variable2, width=10)
        ent2.pack(side=RIGHT, padx=10, fill=X, expand=YES)
        lbl2 = tb.Label(master=container, text=label2.title(), width=length2)
        lbl2.pack(side=RIGHT, padx=10)
        # createToolTip(ent2, tips)
    
    # 创建1个标签、1个按钮
    def Entry_Button_line(self, label, length, button, button_command):
        container = tb.Frame(self)
        container.pack(fill=X, expand=YES, pady=10)
        lbl = tb.Label(master=container, text=label.title(), width=length)
        lbl.pack(side=LEFT, padx=10)
        but = tb.Button(master=container, text=button, command=button_command)
        but.pack(side=RIGHT, padx=10)

    # # 创建1个标签、2个按钮
    # def create_boxline3(self, label, length, button1, button2, button_command1, button_command2):
    #     container = tb.Frame(self)
    #     container.pack(fill=X, expand=YES, pady=10)
    #     lbl = tb.Label(master=container, text=label.title(), width=length)
    #     lbl.pack(side=LEFT, padx=0)
    #     but2 = tb.Button(master=container, text=button2, command=button_command2)
    #     but2.pack(side=RIGHT, padx=10)
    #     but1 = tb.Button(master=container, text=button1, command=button_command1)
    #     but1.pack(side=RIGHT, padx=10)

    # 创建1个标签、3个按钮
    def Lebal_Three_Button_line(self, label, length, button1, button2, button3, acadapp, button_command1, button_command2, button_command3):
        container = tb.Frame(self)
        container.pack(fill=X, expand=YES, pady=10)
        lbl = tb.Label(master=container, text=label.title(), width=length)
        lbl.pack(side=LEFT, padx=10)
        but2 = tb.Button(master=container, text=button2, command = lambda: button_command2(acadapp, but2))
        but2.pack(side=RIGHT, padx=10)
        but1 = tb.Button(master=container, text=button1, command = button_command1)
        but1.pack(side=RIGHT, padx=10)
        but3 = tb.Button(master=container, text=button3, command = button_command3)
        but3.pack(side=RIGHT, padx=10)

    # 包含2个标签、2个编辑框
    def Two_Lebal_Two_Entry_line(self, label1, label2, variable1, variable2, length1, length2, tips1, tips2):
        container = tb.Frame(self)
        container.pack(fill=X, expand=YES, pady=10)
        lbl1 = tb.Label(master=container, text=label1.title(), width=length1)
        lbl1.pack(side=LEFT, padx=10)
        ent1 = tb.Entry(master=container, textvariable=variable1, width=7)
        ent1.pack(side=LEFT, padx=10, fill=X, expand=YES)
        createToolTip(ent1, tips1)
        ent2 = tb.Entry(master=container, textvariable=variable2, width=7)
        ent2.pack(side=RIGHT, padx=10, fill=X, expand=YES)
        createToolTip(ent2, tips2)
        lbl2 = tb.Label(master=container, text=label2.title(), width=length2)
        lbl2.pack(side=RIGHT, padx=10)
        
    # 包含3个标签、3个编辑框
    def Three_Lebal_Three_Entry_line(self, label1, label2, label3, variable1, variable2, variable3, length1, length2, length3, tips1, tips2, tips3):
        container = tb.Frame(self)
        container.pack(fill=X, expand=YES, pady=10)
        lbl1 = tb.Label(master=container, text=label1.title(), width=length1)
        lbl1.pack(side=LEFT, padx=10)
        ent1 = tb.Entry(master=container, textvariable=variable1, width=7)
        ent1.pack(side=LEFT, padx=20, fill=X, expand=YES)
        createToolTip(ent1, tips1)
        lbl2 = tb.Label(master=container, text=label2.title(), width=length2)
        lbl2.pack(side=LEFT, padx=10)
        ent2 = tb.Entry(master=container, textvariable=variable2, width=7)
        ent2.pack(side=LEFT, padx=20, fill=X, expand=YES)
        createToolTip(ent2, tips2)
        lbl3 = tb.Label(master=container, text=label3.title(), width=length3)
        lbl3.pack(side=LEFT, padx=10)
        ent3 = tb.Entry(master=container, textvariable=variable3, width=8)
        ent3.pack(side=RIGHT, padx=10, fill=X, expand=YES)
        createToolTip(ent3, tips3)

    # 标签 + 边界框 + 按钮
    def dialog_boxline(self, label_txt, label_length, entry_var, entry_length, button_txt, button_command):
        container = tb.Frame(self)
        container.pack(fill=X, expand=YES, pady=5)
        # 标签
        Label = tb.Label(master=container, text=label_txt.title(), width=label_length)
        Label.pack(side=LEFT, padx=10)
        # 编辑框
        ent = tb.Entry(master=container, textvariable=entry_var, width=entry_length, state="disabled")
        ent.pack(side=LEFT, padx=20, fill=X, expand=YES)
        # 按钮
        but = tb.Button(master=container, text=button_txt, command = button_command, width = 15)
        but.pack(side=RIGHT, padx=10)

    # 标签 + 编辑框 + 按钮
    def dialog_boxline(self, label_txt, label_length, entry_var, entry_length, button_txt, button_command):
        container = tb.Frame(self)
        container.pack(fill=X, expand=YES, pady=5)
        # 标签
        Label = tb.Label(master=container, text=label_txt.title(), width=label_length)
        Label.pack(side=LEFT, padx=10)
        # 编辑框
        ent = tb.Entry(master=container, textvariable=entry_var, width=entry_length, state="disabled")
        ent.pack(side=LEFT, padx=20, fill=X, expand=YES)
        # 按钮
        but = tb.Button(master=container, text=button_txt, command = button_command, width = 15)
        but.pack(side=RIGHT, padx=10)

    # 风荷载
    def combobox_Feng_line(self, label, lst, button, length, secondary_window_func):
        # 获取选择值
        def on_select(event):
            self.FengGui_selected_value = combo.get()

        container = tb.Frame(self)
        container.pack(fill=X, expand=YES, pady=5)
        lbl = tb.Label(master=container, text=label.title(), width=length)
        lbl.pack(side=LEFT, padx=10)
        # 下拉列表
        combo = tb.Combobox(master=container)
        combo.pack(side=LEFT, padx=10, fill=X)
        # 设置下拉列表的选项
        combo['values'] = lst
        # 设置默认值
        combo.current(0)
        # 绑定选择事件
        combo.bind('<<ComboboxSelected>>', on_select)
        but = tb.Button(master=container, text=button, command=lambda: secondary_window_func(self.root, self.FengGui_selected_value, self.Feng_values, but))
        but.pack(side=RIGHT, padx=10)
    
    # 水荷载
    def combobox_Shui_line(self, label, lst, button, length, secondary_window_func):
        # 获取选择值
        def on_select(event):
            self.ShuiGui_selected_value = combo.get()
        container = tb.Frame(self)
        container.pack(fill=X, expand=YES, pady=5)
        lbl = tb.Label(master=container, text=label.title(), width=length)
        lbl.pack(side=LEFT, padx=10)
        # 下拉列表
        combo = tb.Combobox(master=container)
        combo.pack(side=LEFT, padx=10, fill=X)
        # 设置下拉列表的选项
        combo['values'] = lst
        # 设置默认值
        combo.current(0)
        # 绑定选择事件
        combo.bind('<<ComboboxSelected>>', on_select)
        but = tb.Button(master=container, text=button, command=lambda: secondary_window_func(self.root, self.ShuiGui_selected_value, self.Shui_values, but))
        but.pack(side=RIGHT, padx=10)
    
    def FPL_combo_line(self,lst1,lst2,length):
        self.options1 = {"单截面":lst1,"双截面":lst2}
        # 创建单选按钮变量
        self.selected_option1 = tb.StringVar()
        # 创建单选按钮和下拉列表的容器
        self.radio_frame1 = tb.Frame(self)
        self.radio_frame1.pack(fill='x', padx=10, pady=10)
        lbl1 = tb.Label(self.radio_frame1, text="型钢类别：".title(), width=length)
        lbl1.pack(side=LEFT, padx=10, fill=X, expand=YES)
        # 创建单选按钮并绑定值
        for option_text, values in self.options1.items():
            rb = tb.Radiobutton(self.radio_frame1, text=option_text, value=option_text, variable=self.selected_option1, command=self.update_FPL_combo)
            rb.pack(side=LEFT, padx=20, fill=X, expand=YES)
        # 创建下拉列表
        self.combobox1 = tb.Combobox(self.radio_frame1,textvariable=self.selected_value1)
        self.combobox1.pack(side=RIGHT, padx=0, fill=X, expand=YES)
        lbl2 = tb.Label(self.radio_frame1, text="截面选择：".title(), width=length)
        lbl2.pack(side=RIGHT, padx=15, fill=X, expand=YES)
        # 设置默认的单选按钮值
        self.selected_option1.set(list(self.options1.keys())[0])
        self.update_FPL_combo()

    def update_FPL_combo(self):
        # 获取当前选中的单选按钮的值
        selected_option = self.selected_option1.get()
        # 更新下拉列表的选项
        self.combobox1['values'] = self.options1[selected_option]
        # 如果下拉列表中有选项，则选择第一个选项
        if self.combobox1['values']:
            self.combobox1.current(0)
    
    def rib_combo_line(self, lst1, lst2, length):
        self.options2 = {"单截面":lst1,"双截面":lst2}
        # 创建单选按钮变量
        self.selected_option2 = tb.StringVar()
        # 创建单选按钮和下拉列表的容器
        self.radio_frame2 = tb.Frame(self)
        self.radio_frame2.pack(fill='x', padx=10, pady=10)
        lbl1 = tb.Label(self.radio_frame2, text="型钢类别：".title(), width=length)
        lbl1.pack(side=LEFT, padx=10, fill=X, expand=YES)
        # 创建单选按钮并绑定值
        for option_text, values in self.options2.items():
            rb = tb.Radiobutton(self.radio_frame2, text=option_text, value=option_text, variable=self.selected_option2, command=self.update_rib_combo)
            rb.pack(side=LEFT, padx=20, fill=X, expand=YES)
        # 创建下拉列表
        self.combobox2 = tb.Combobox(self.radio_frame2,textvariable=self.selected_value2)
        self.combobox2.pack(side=RIGHT, padx=0, fill=X, expand=YES)
        lbl2 = tb.Label(self.radio_frame2, text="截面选择：".title(), width=length)
        lbl2.pack(side=RIGHT, padx=0, fill=X, expand=YES)
        # 设置默认的单选按钮值
        self.selected_option2.set(list(self.options2.keys())[0])
        self.update_rib_combo()
        
    def update_rib_combo(self):
        # 获取当前选中的单选按钮的值
        selected_option = self.selected_option2.get()
        # 更新下拉列表的选项
        self.combobox2['values'] = self.options2[selected_option]
        # 如果下拉列表中有选项，则选择第一个选项
        if self.combobox2['values']:
            self.combobox2.current(0)
    
    def endline(self):
        container = tb.Frame(self)
        container.pack(fill=X, expand=YES, pady=(15, 10))
        self.chVarEn1 = tb.IntVar(value=1)
        self.chVarEn2 = tb.IntVar(value=0)
        self.chVarEn3 = tb.IntVar(value=0)
        check1 = tb.Checkbutton(master=container, text="Midas模型", variable = self.chVarEn1, bootstyle="success-round-toggle", command = lambda: self.CheckButtonValue("Midas"))
        # check2 = tb.Checkbutton(master=container, text="CAD布置图", variable = self.chVarEn2, bootstyle="success-round-toggle", command = lambda: self.CheckButtonValue("CAD"))
        # check3 = tb.Checkbutton(master=container, text="Word计算书", variable = self.chVarEn3, bootstyle="success-round-toggle", command = lambda: self.CheckButtonValue("Word"))
        check1.pack(side=LEFT, padx=10)
        # check2.pack(side=LEFT, padx=5)
        # check3.pack(side=LEFT, padx=5) 
        cnl_btn = tb.Button(master=container,text="退出",command=self.on_cancel,bootstyle=DANGER,width=6)
        cnl_btn.pack(side=RIGHT, padx=10)
        sub_btn = tb.Button(master=container,text="确认",command=lambda:self.on_submit(),bootstyle=SUCCESS,width=6)
        sub_btn.pack(side=RIGHT, padx=10)
        sub_btn.focus_set()

    def CheckButtonValue(self, param):
        if param == "Midas":
            self.chVarEn1.set(True)
            self.chVarEn2.set(False)
            self.chVarEn3.set(False)

        if param == "CAD":
            self.chVarEn1.set(False)
            self.chVarEn2.set(True)
            self.chVarEn3.set(False)

        if param == "Word":
            self.chVarEn1.set(False)
            self.chVarEn2.set(False)
            self.chVarEn3.set(True)

    def creat_dialog(self, acadapp, Applocation):
        # 获取脚本所在目录的绝对路径
        # script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        ribsteel1_filepath = os.path.join(Applocation, 'Support', 'basic_param', 'ribsteel1.sec')  # ribsteel1文件在同目录
        ribsteel2_filepath = os.path.join(Applocation, 'Support', 'basic_param', 'ribsteel2.sec')  # ribsteel2文件在同目录
        FPL1_filepath = os.path.join(Applocation, 'Support', 'basic_param', 'FPL1.sec')  # FPL11文件在同目录
        FPL2_filepath = os.path.join(Applocation, 'Support', 'basic_param', 'FPL2.sec')  # FPL2文件在同目录
        self.rib_lst = rib_sections(ribsteel1_filepath, ribsteel2_filepath)
        self.FPL_lst = FPL_section(FPL1_filepath, FPL2_filepath)
        distributed_list1 = self.FPL_lst[0]
        self.FPL_seclst1 = self.FPL_lst[1]# 变量未引用
        distributed_list2 = self.FPL_lst[2]
        self.FPL_seclst2 = self.FPL_lst[3]# 变量未引用
        rib_list1 = self.rib_lst[0]
        rib_list2 = self.rib_lst[2]
        FengGui_lst = ["自定义", "公路桥梁抗风设计规范", "工程结构通用规范", "港口工程荷载规范"]
        ShuiGui_lst = ["自定义", "港口工程荷载规范"] 
        

        # self.creatheadr ("固定参数调试及图元生成：")
        # self.create_boxline1 ("小肋数量及间距布置(mm)：", self.entryname9, 20, "数据格式: 数量1@间距1 + ... + 数量n@间距n")
        # self.creatcombobox2 (rib_list1, rib_list2, 8)
        # self.create_boxline2 ("小肋顺桥向起始横坐标(mm)：", self.entryname19, 22, "小肋横桥向悬臂长度(mm)：", self.entryname15, 22, "提示")
        # self.create_boxline3 ("绘制小肋及贝雷片套组：", 18, "绘制小肋布置图", "生成贝雷片套组", lambda:draw_rib_blocks_incad(acadapp, self.selected_value2.get(), self.entryname9.get()), lambda:draw_Bailey_incad(acadapp))

        # self.creatheadr ("动态参数调试及图元生成：")
        # self.creatcombobox1 (distributed_list1, distributed_list2, 8)
        # self.createntry ("钢管桩外直径(mm):", "钢管桩壁厚(mm):", "钢管桩桩长(mm):", self.entryname3, self.entryname4, self.entryname5, 15, 13, 13, "钢管桩外直径", "钢管桩壁厚", "钢管桩桩长")
        # self.create_boxline3 ("绘制分配梁及钢管桩：", 18, "绘制分配梁图块", "绘制钢管桩图块", lambda:draw_FPL_incad(acadapp, self.selected_value1.get()), lambda:draw_SP1_incad(acadapp, self.entryname3.get(), self.entryname4.get(), self.entryname5.get()))

        # self.creatheadr ("识别交互：")
        # self.create_boxline4 ("平面布置图识别及小肋组截面赋予：", 25, "识别平面布置图", "小肋组截面赋予", "定位坐标原点", acadapp, lambda:self.Reg_All_Bailey_Fpl_SP(acadapp), self.create_secondary_dialog, lambda:self.get_User_Orignal_Point_inCAD(acadapp))

        # self.creatheadr ("荷载定义：")
        # self.createntry ("混凝土荷载(kN/m3):", "施工荷载(kN/m2):", "模板荷载(kN/m2):", self.entryname12, self.entryname10, self.entryname11, 15, 13, 13, "默认值为26.5kN/m3", "默认值为2.0kN/m2", "默认值为2.5kN/m2")

        # self.creatcombobox_Feng("风荷载规范：", FengGui_lst, "风荷载参数", 15, create_Feng_dialog)
        # self.creatcombobox_Shui("水流力规范：", ShuiGui_lst, "水流力参数", 15, create_Shui_dialog)

        # self.creatheadr ("保存路径：")
        # self.dialog_savepath("MCT文件保存路径", 13, self.mct_savepath, 10, "指定保存路径", lambda: mct_path_SaveAs(self.mct_savepath))

        # self.create_endlinebuttonbox()
        # self.root.mainloop()

        self.headr_line ("小肋定义：")
        self.Label_Entry_line ("小肋数量及间距布置(mm)：", self.entryname9, 20, "数据格式: 数量1@间距1 + ... + 数量n@间距n")
        self.rib_combo_line (rib_list1, rib_list2, 8)
        self.Two_Label_Two_Entry_line ("小肋顺桥向起始横坐标(mm)：", self.entryname19, 22, "小肋横桥向悬臂长度(mm)：", self.entryname15, 22, "提示")
        self.Entry_Button_line ("根据小肋数量、间距布置及小肋型钢截面参数, 在CAD中指定点生成小肋布置", 60, "绘制小肋布置", lambda:draw_rib_blocks_incad(acadapp, self.selected_value2.get(), self.entryname9.get()))

        self.headr_line ("贝雷梁定义：")
        self.Entry_Button_line ("在CAD中指定点生成3m、1.5m贝雷片的三视图", 60, "绘制贝雷片组", lambda:draw_Bailey_incad(acadapp))
        
        self.headr_line ("分配梁定义：")
        self.FPL_combo_line (distributed_list1, distributed_list2, 8)
        self.Entry_Button_line ("根据分配梁截面类型参数, 在CAD中指定点生成分配梁图块", 60, "绘制分配梁图块", lambda:draw_FPL_incad(acadapp, self.selected_value1.get()))

        self.headr_line ("钢管桩定义：")
        self.Three_Lebal_Three_Entry_line ("钢管桩外直径(mm):", "钢管桩壁厚(mm):", "钢管桩桩长(mm):", self.entryname3, self.entryname4, self.entryname5, 15, 13, 13, "钢管桩外直径", "钢管桩壁厚", "钢管桩桩长")
        self.Entry_Button_line ("根据钢管桩直径、壁厚、桩长参数, 在CAD中指定点生成钢管桩图块", 60, "绘制钢管桩图块", lambda:draw_SP1_incad(acadapp, self.entryname3.get(), self.entryname4.get(), self.entryname5.get()))

        self.headr_line ("支架识别：")
        self.Lebal_Three_Button_line ("根据提示与CAD交互：", 25, "识别平面布置图", "添加主梁线荷载", "定位坐标原点", acadapp, lambda:self.Reg_All_Bailey_Fpl_SP(acadapp), self.create_secondary_dialog, lambda:self.get_User_Orignal_Point_inCAD(acadapp))

        # self.headr_line ("荷载定义：")
        # self.Two_Lebal_Two_Entry_line ("自重系数:", "混凝土容重(kN/m3):", self.entryname13, self.entryname12, 15, 15, "默认值为-1", "默认值为26.5kN/m3")
        # self.Two_Lebal_Two_Entry_line ("施工荷载(kN/m2):", "模板荷载(kN/m2):", self.entryname10, self.entryname11, 15, 15, "默认值为2.0kN/m2", "默认值为2.5kN/m2")

        # self.combobox_Feng_line("风荷载规范：", FengGui_lst, "风荷载参数", 15, create_Feng_dialog)
        # self.combobox_Shui_line("水流力规范：", ShuiGui_lst, "水流力参数", 15, create_Shui_dialog)

        self.headr_line ("文件操作：")
        # 程序状态
        self.dialog_boxline("程序状态", 13, self.Set_File_entry_var, 10, "打开Midas civil NX", lambda: Open_Midas_civil(self.civilnx_path, self.Set_File_entry_var))
        self.dialog_boxline("MCT文件保存路径", 13, self.mct_savepath, 10, "更改文件保存路径", lambda: mct_path_SaveAs(self.mct_savepath))

        self.endline()
        # self.root.mainloop()

    # # 水流力计算对话框
    # def cal_Shui_dialog(self):
    #     print(self.ShuiGui_selected_value)

    # 小肋-荷载对话框
    def create_secondary_dialog(self, acadapp, button):
        secondary_dialog = tk.Toplevel(self.root)
        secondary_dialog.title("小肋—混凝土荷载参数表")
        frame1 = tb.Frame(secondary_dialog)
        frame1.pack(fill=tk.BOTH, expand=True)
        frame2 = tb.Frame(secondary_dialog)
        frame2.pack(fill=tk.BOTH, expand=True)
        # global current_row
        # current_row = 0
        if self.current_row == None:
            self.current_row = 0

        def add_entry_row(acadapp, tips1, tips2, tips3, param):
            if param == 0:
                globals()[f"sec_entryname{self.current_row}"] = tb.StringVar(value="")
                globals()[f"Ecc{self.current_row}"] = tb.StringVar(value="0")
                globals()[f"sec_info{self.current_row}"] = tb.StringVar(value="未录入")
            else:
                globals()[f"sec_entryname{self.current_row}"] = tb.StringVar(value=globals()[f"sec_entryname{self.current_row}"].get())
                globals()[f"Ecc{self.current_row}"] = tb.StringVar(value=globals()[f"Ecc{self.current_row}"].get())
                globals()[f"sec_info{self.current_row}"] = tb.StringVar(value=globals()[f"sec_info{self.current_row}"].get())

            globals()[f"sec_info{self.current_row}"].trace_add("write", lambda *args, t="{}".format(self.current_row+1): update_button_state(t, *args))
            # 文本内容，说明小肋的序号组
            globals()[f"lable{self.current_row}"] = tb.Label(master=frame1,text="{} {}".format("小肋组序号",self.current_row+1))
            globals()[f"lable{self.current_row}"].grid(row=self.current_row, column=0, padx=5, pady=5, sticky="e")
            # 编辑框内容，用来存储一组小肋的 x to y 或单根序号 n 的信息，并设置提示文本
            globals()[f"entry{self.current_row}"] = tb.Entry(frame1,textvariable=globals()[f"sec_entryname{self.current_row}"],width=15)
            globals()[f"entry{self.current_row}"].grid(row=self.current_row, column=1, padx=5, pady=5, sticky="ew")
            createToolTip(globals()[f"entry{self.current_row}"], tips1)
            # 按钮，返回值出现在编辑框中
            globals()[f"button1{self.current_row}"] = tb.Button(frame1, text="{} {}".format("识别小肋组",self.current_row+1),
                                                                 command=lambda t="{}".format(self.current_row+1): (self.get_buttontext_toentry(acadapp,t)))
            globals()[f"button1{self.current_row}"].grid(row=self.current_row, column=2, padx=5, pady=5, sticky="ew")
            # 文本内容，说明截面偏心
            globals()[f"second_lable{self.current_row}"] = tb.Label(master=frame1,text="{} {}".format("截面偏心",self.current_row+1))
            globals()[f"second_lable{self.current_row}"].grid(row=self.current_row, column=3, padx=5, pady=5, sticky="e")
            # 编辑框内容，用来储存截面偏心值Ecc，并设置提示文本
            globals()[f"second_entry{self.current_row}"] = tb.Entry(frame1,textvariable=globals()[f"Ecc{self.current_row}"],width=8)
            globals()[f"second_entry{self.current_row}"].grid(row=self.current_row, column=4, padx=5, pady=5, sticky="ew")
            createToolTip(globals()[f"second_entry{self.current_row}"], tips2)
            
            # 按钮，生成截面相关信息 xlst 和 hlst ，生成的变量均为全局变量
            globals()[f"button2{self.current_row}"] = tb.Button(frame1, text="{} {}".format("识别截面组",self.current_row+1),
                                                                 command=lambda t="{}".format(self.current_row+1): (self.get_sectiondata_tolist(acadapp,t)))
            globals()[f"button2{self.current_row}"].grid(row=self.current_row, column=5, padx=5, pady=5, sticky="ew")
            # 编辑框，显示截面是否录入完成，并设置提示文本
            globals()[f"third_entry{self.current_row}"] = tb.Entry(frame1,textvariable=globals()[f"sec_info{self.current_row}"],width=8)
            globals()[f"third_entry{self.current_row}"].grid(row=self.current_row, column=6, padx=5, pady=5, sticky="ew")
            createToolTip(globals()[f"third_entry{self.current_row}"], tips3)
            
            self.current_row += 1
        
        def delete_entry_row():
            # 从网格中删除Entry控件和删除按钮
            #global current_row
            current_row2 = self.current_row-1
            globals()[f"lable{current_row2}"].grid_forget()
            globals()[f"entry{current_row2}"].grid_forget()
            globals()[f"button1{current_row2}"].grid_forget()
            globals()[f"second_lable{current_row2}"].grid_forget()
            globals()[f"second_entry{current_row2}"].grid_forget()
            globals()[f"button2{current_row2}"].grid_forget()
            globals()[f"third_entry{current_row2}"].grid_forget()
            #globals()[f"button3{current_row2}"].grid_forget()
            # 减少行号
            self.current_row -= 1

        def last_row(acadapp):
            addline_button = tb.Button(frame2, text="添加一行", command = lambda: add_entry_row(acadapp, "请使用识别小肋按钮进行识别" , "请输入截面定位线至平面布置定位线的距离", "输入1表示截面同上", 0))
            addline_button.pack(side=LEFT, padx=10, pady=10)
            delline_button = tb.Button(frame2, text="删除一行", command = delete_entry_row)
            delline_button.pack(side=LEFT, padx=10, pady=10)
            # numline_button = tb.Button(frame2, text="识别小肋序号")
            # numline_button.pack(side=LEFT, padx=10, pady=10)
            close_button = tb.Button(frame2, text="退出", command = lambda: enable_button(secondary_dialog, button))
            close_button.pack(side=RIGHT, padx=10, pady=10)
            accline_button = tb.Button(frame2, text="确认", command = comfirm_secondarydialog_list)
            accline_button.pack(side=RIGHT, padx=10, pady=10)

        def update_button_state(text, *args):
            # 获取编辑框的值
            # print(text)
            entry_value = globals()[f"sec_info{int(text)-1}"].get()
            # print(entry_value)
            # 根据编辑框的值更新按钮状态
            if entry_value == "1":
                globals()[f"button2{int(text)-1}"].state(["disabled"])  # 禁用按钮
            else:
                globals()[f"button2{int(text)-1}"].state(["!disabled"])  # 启用按钮

        # 确认按钮
        def comfirm_secondarydialog_list():
            print("小肋-截面信息已保存。")

        # 第一次进行截面赋予
        if self.current_row == 0:
            add_entry_row(acadapp, "请使用识别小肋按钮进行识别" , "请输入截面定位线至平面布置定位线的距离", "输入1表示截面同上", 0)
            last_row(acadapp)
            secondary_dialog.grid_columnconfigure(1, weight=1)
        # 根据截面赋予信息中的全局变量和行数重生成一张表
        elif self.current_row != 0:
            loop_a = self.current_row
            self.current_row = 0
            for loop_row in range(loop_a):
                add_entry_row(acadapp, "请使用识别小肋按钮进行识别" , "请输入截面定位线至平面布置定位线的距离", "输入1表示截面同上", 1)
            last_row(acadapp)
            secondary_dialog.grid_columnconfigure(1, weight=1)

        # 禁用主窗口的按钮
        button.config(state=tk.DISABLED)
        # 当二级窗口关闭时，重新启用主窗口的按钮
        secondary_dialog.protocol("WM_DELETE_WINDOW", lambda: enable_button(secondary_dialog, button))

    def get_global_secentryname_value(self,num):
        key = f"sec_entryname{num}"
        if key in globals():
            return globals()[key]
        else:
            return None
    
    # 将小肋的数量转换为 x to y 的文本
    def get_buttontext_toentry(self,acadapp,text):
        ribnum=self.reg_ribnum_incad(acadapp)
        if text=='1':
            self.ribstartnum=1
            self.ribendnum=ribnum
            if ribnum==1:
                globals()[f"sec_entryname{int(text)-1}"].set("{}".format(self.ribendnum))
            else:
                globals()[f"sec_entryname{int(text)-1}"].set("{} {} {}".format(self.ribstartnum,"to",self.ribendnum))
            self.ribstartnum=self.ribendnum+1
        else:
            self.ribendnum=self.ribstartnum+ribnum-1
            if ribnum==1:
                globals()[f"sec_entryname{int(text)-1}"].set("{}".format(self.ribendnum))
            else:
                globals()[f"sec_entryname{int(text)-1}"].set("{} {} {}".format(self.ribstartnum,"to",self.ribendnum))
            self.ribstartnum=self.ribendnum+1
    
    # 识别得到小肋的数量
    def reg_ribnum_incad(self,acadapp):
        rib_num = 0
        try:
            SelectOnScreenlist = SelectOnScreen(acadapp, "请选择一组小肋")
            acaddoc = SelectOnScreenlist[0]#acaddoc = acadapp.ActiveDocument
            slt = SelectOnScreenlist[2]#返回<COMObject Add>，是一个集合
            # print(slt.Count)
            if slt.Count > 0 :
                # for x in ss_to_lst(slt):
                #     objname = acaddoc.HandleToObject(x).ObjectName
                for obj in slt:
                    objname = obj.EntityName
                    try:
                        block_xdata = get_block_reference_xdata(obj)
                        # print(block_xdata)
                        xdata = block_xdata[1][1]
                        # print(xdata)
                        if objname == 'AcDbBlockReference' and "EX_XL" in xdata:#如果是块返回AcDbBlockReference
                        # 查询扩展属性是否是小肋相关
                            rib_num = rib_num+1
                    except:
                        pass
            else:
                print('捕捉到0个图形。')
            if rib_num == 0:
                print("Warning: 未识别到小肋相关属性，捕捉小肋失败。")
        except:
            print("小肋捕捉失败，请重试。")
            
        return rib_num

    # 从贝雷梁中获取小肋的基本信息，包括贝雷梁的纵向排布
    def Get_Ribinfo_from_Bailey(self):
        # 得到贝雷梁词典
        self.Bailey_dict = Cross_Bailey_Rib(self.entryname9.get(), float(self.entryname19.get()), self.Bailey_CAD_to_lst)
        # self.Bailey_y_dist_lst = self.Bailey_dict["所有y坐标"]
        

    # 实现贝雷梁和分配梁的交互
    def Get_FPLinfo_from_Bailey(self):
        dicts = Cross_Bailey_FPL(self.Bailey_dict, self.FPL_CAD_to_lst)
        # 贝雷梁包含和小肋以及分配梁有关信息的词典
        self.Bailey_Rib_FPL_dict = dicts[0]
        # 分配梁包含和贝雷梁有关信息的词典
        self.FPL_Bailey_dict = dicts[1]
        

    # 捕捉钢管桩桩顶桩底的坐标和直径，以及每一道联结系在哪两根钢管桩之间
    def Get_SPinfo_from_CAD(self):
        # 获取基本信息
        self.SP_dict1 = self.SP_CAD_to_lst.copy()
        

    # 返回识别截面的基本信息
    def get_sectiondata_tolist(self,acadapp,text):
        # text表示行数，初始值从"1"开始
        # print(text)
        # 截面信息初始化
        try:
            if text == "1":
                self.SEC_basic_info_dict = {}
            acaddoc = acadapp.ActiveDocument
            acadmsp = acaddoc.ModelSpace
            get_SECinfo_fromcad_reslst = get_SECinfo_fromcad(acadapp, acaddoc, acadmsp)
            acaddoc = get_SECinfo_fromcad_reslst[0]
            acadmsp = get_SECinfo_fromcad_reslst[1]
            xlst = get_SECinfo_fromcad_reslst[2]
            coord_y_lst = get_SECinfo_fromcad_reslst[3]
            coord_h_lst = get_SECinfo_fromcad_reslst[4]

            # 判断函数是否返回结果，如果返回，且编辑框内文本内容不为"1"，则将未录入显示为已录入
            if get_SECinfo_fromcad_reslst and globals()[f"sec_info{int(text)-1}"].get() != "1":
                globals()[f"sec_info{int(text)-1}"].set("已录入")

                # 将截面基本信息词典更新，text可以确定本次点击的是第几行，text从"1"开始
                # 设置初始值
                if f"第{text}组截面基本信息" not in self.SEC_basic_info_dict:
                    self.SEC_basic_info_dict[f"第{text}组截面基本信息"] = []
                
                # 将需要的信息放入词典
                self.SEC_basic_info_dict[f"第{text}组截面基本信息"] = [acaddoc, acadmsp, xlst, coord_y_lst, coord_h_lst]

            # 如果没有截面信息且当前编辑框内不为"1"，则还是显示未录入
            else:
                globals()[f"sec_info{int(text)-1}"].set("未录入")
            
            # print(self.SEC_basic_info_dict)
        except:
            print("捕捉截面失败。可能的原因：用户取消操作、截面并非多段线、多段线不在同一高度、多段线未闭合。")
        

    def return_secondarydialog_list(self):
        print("开始整合截面信息，请等待...")

        self.secetyname_lst = []
        self.Distance_between_section_Bailey_Ry_lst = []
        self.Sec_info_states = []

        for x in range(self.current_row):

            # 得到第一个编辑框中所有的小肋序号
            update_or_append(self.secetyname_lst, x, self.get_global_secentryname_value(x).get())

            # 得到第二个编辑框中的所有截面偏心值
            Distance_between_section_Bailey_Ry = float(globals()[f"Ecc{x}"].get())
            update_or_append(self.Distance_between_section_Bailey_Ry_lst, x, Distance_between_section_Bailey_Ry)

            # 得到第三个编辑框中的所有截面录入状态
            Sec_info_state = globals()[f"sec_info{x}"].get()
            update_or_append(self.Sec_info_states, x, Sec_info_state)

        # 得到小肋的基本信息
        self.get_rib_basedata_resultlist = get_rib_basedata(self.selected_value2.get(), self.selected_option2.get(), self.entryname9.get(), self.entryname19.get(), self.rib_lst)
        # print(self.get_rib_basedata_resultlist)
        rib_xlst = [round(x, 1) for x in self.get_rib_basedata_resultlist[2]]
        # print(rib_xlst)
        # print(self.Bailey_dict)
        
        self.rib_Bailey_y_dict = {}
        # 创建小肋与贝雷梁弹性连接的y坐标词典
        for x in rib_xlst:
            if f"{x}" not in self.rib_Bailey_y_dict:
                self.rib_Bailey_y_dict[f"{x}"] = []
            for key, value in self.Bailey_dict.items():
                if "与小肋弹连坐标" in key:
                    for pt in value:
                        if abs(pt[0] - x) < 1:
                            self.rib_Bailey_y_dict[f"{x}"].append(pt[1])
        # print("Midas中小肋与贝雷的弹连y坐标：", self.rib_Bailey_y_dict)
        
        # 根据行数依次将信息一行一行结合
        for y in range(self.current_row):

            # 如果当前行文本信息是已录入，则说明是新截面
            if self.Sec_info_states[y] == "已录入": 

                acaddoc = self.SEC_basic_info_dict[f"第{y+1}组截面基本信息"][0]
                acadmsp = self.SEC_basic_info_dict[f"第{y+1}组截面基本信息"][1]
                xlst = self.SEC_basic_info_dict[f"第{y+1}组截面基本信息"][2]
                coord_y_lst = self.SEC_basic_info_dict[f"第{y+1}组截面基本信息"][3]
                coord_h_lst = self.SEC_basic_info_dict[f"第{y+1}组截面基本信息"][4]

                # 获取当前单个截面相关信息, 返回小肋y坐标和截面高度
                get_heightlst_fromcad_reslst = get_heightlst_fromcad(xlst, coord_h_lst, self.secetyname_lst[y], rib_xlst, self.rib_Bailey_y_dict, self.Distance_between_section_Bailey_Ry_lst[y])

                globals()[f"section_xlst1{y}"]=get_heightlst_fromcad_reslst[0]
                globals()[f"section_xlst2{y}"]=get_heightlst_fromcad_reslst[1]
                globals()[f"section_hlst{y}"]=get_heightlst_fromcad_reslst[2]

                # print(globals()[f"section_xlst1{y}"], globals()[f"section_xlst2{y}"], globals()[f"section_hlst{y}"])

                # 组合
                update_or_append(self.section_xlst1, y, globals()[f"section_xlst1{y}"])
                update_or_append(self.section_xlst2, y, globals()[f"section_xlst2{y}"])
                update_or_append(self.section_hlst, y, globals()[f"section_hlst{y}"])

                # 储存最近更新的截面信息
                SEC_basic_info_lst = [acaddoc, acadmsp, xlst, coord_y_lst, coord_h_lst]

            # 如果当前行文本信息为"1"，说明该变量和上一个变量的截面基本信息一致
            if  self.Sec_info_states[y] == "1":

                # 读取最近更新的截面信息
                acaddoc = SEC_basic_info_lst[0]
                acadmsp = SEC_basic_info_lst[1]
                xlst = SEC_basic_info_lst[2]
                coord_y_lst = SEC_basic_info_lst[3]
                coord_h_lst = SEC_basic_info_lst[4]

                # 获取当前单个截面相关信息, 返回小肋y坐标和截面高度
                get_heightlst_fromcad_reslst = get_heightlst_fromcad(xlst, coord_h_lst, self.secetyname_lst[y], rib_xlst, self.rib_Bailey_y_dict, self.Distance_between_section_Bailey_Ry_lst[y])
                
                globals()[f"section_xlst1{y}"]=get_heightlst_fromcad_reslst[0]
                globals()[f"section_xlst2{y}"]=get_heightlst_fromcad_reslst[1]
                globals()[f"section_hlst{y}"]=get_heightlst_fromcad_reslst[2]

                # print(globals()[f"section_xlst1{y}"], globals()[f"section_xlst2{y}"], globals()[f"section_hlst{y}"])

                # 组合
                update_or_append(self.section_xlst1, y, globals()[f"section_xlst1{y}"])
                update_or_append(self.section_xlst2, y, globals()[f"section_xlst2{y}"])
                update_or_append(self.section_hlst, y, globals()[f"section_hlst{y}"])

            print(f"当前整合第{y+1}个截面信息。")
        
        # print("编辑框中所有的小肋序号：", self.secetyname_lst)
        # print("编辑框中的所有截面偏心值：", self.Distance_between_section_Bailey_Ry_lst)
        # print("编辑框中的所有截面录入状态：", self.Sec_info_states)
        # print('所有的截面xlst1：', self.section_xlst1)
        # print('所有截面的xlst2：', self.section_xlst2)
        # print('所有截面的hlst：', self.section_hlst)
        print("Success: 截面信息整合完成。")
        # 返回所有的小肋序号字符串， 所有的截面xlst1，所有截面的xlst2，所有截面的hlst
        return self.secetyname_lst, self.section_xlst1, self.section_xlst2, self.section_hlst
    #返回用户所有输入或识别的小肋序号组、结合贝雷梁后的截面x坐标、结合贝雷梁后的截面左右0.01的x坐标、结合贝雷梁后的x坐标下的截面对应高度


    # 贝雷梁、分配梁、钢管桩、联结系的全部平面识别
    def Reg_All_Bailey_Fpl_SP(self, acadapp):
        try:
            # 捕获当前小肋间距及数量的entry对应的变量值
            self.rib_info_str1 = self.entryname9.get()
            self.rib_info_str2 = self.entryname19.get()

            # 重置图形信息
            self.Bailey_CAD_to_lst = []
            self.FPL_CAD_to_lst = []
            self.SP_CAD_to_lst = []

            # 获取当前Cad对应的活动文档、模型空间、贝雷梁平面图选择集
            Select_lst = SelectOnScreen(acadapp, "请选择平面布置图")
            Select_slt = Select_lst[2]
            Bailey_select = [] # Select_slt 集合中和贝雷有关的图块
            Fenpeiliang_select = [] # Select_slt 集合中和分配梁有关的图块
            Gangguanzhuang_select1 = [] # Select_slt 集合中和钢管桩桩顶有关的图块
            Gangguanzhuang_select2 = [] # Select_slt 集合中和钢管桩桩底有关的图块
            Lianjiexi_select = [] # Select_slt 集合中和联结系有关的图块
            # 先对 Select_slt 分类, 属于贝雷、分配梁、钢管桩的图块
            for obj in Select_slt:
                if obj.EntityName == "AcDbBlockReference":
                    xdata = get_block_reference_xdata(obj)
                    if (obj.Name == "1.5m贝雷平面block" or obj.Name == "3m贝雷平面block") and xdata[1][0] == "SZQS" and (xdata[1][1] == "EX_BL3_PM" or xdata[1][1] == "EX_BL1_PM"):
                        Bailey_select.append(obj)
                    if xdata[1][0] == "SZQS" and "EX_FPL_PM" in xdata[1][1]:
                        Fenpeiliang_select.append(obj)
                    if "SP1_Block" in obj.Name:
                        Gangguanzhuang_select1.append(obj)
                elif obj.EntityName == "AcDbCircle" and obj.Layer == 'Defpoints':
                    Gangguanzhuang_select2.append(obj)
                elif obj.EntityName == "AcDbLine" and obj.Layer == "3中心线":
                    Lianjiexi_select.append(obj)
            if Lianjiexi_select != []:
                # 捕捉贝雷梁信息
                self.Bailey_CAD_to_lst = make_Bailey_origin_points(acadapp, Bailey_select, self.User_Orignal_Point_inCAD)
                print("Success: 完成贝雷梁平面布置识别")
                # 捕捉分配梁信息
                self.FPL_CAD_to_lst = Reg_FPL_PM(Fenpeiliang_select, self.Bailey_CAD_to_lst)
                print("Success: 完成分配梁平面布置识别")
                # 捕捉钢管桩信息
                self.SP_CAD_to_lst = get_sp_info_fromcad(Gangguanzhuang_select1, Gangguanzhuang_select2, Lianjiexi_select)
                print("Success: 完成钢管桩平面布置识别")
            else:
                print('未添加联结系, 请用3中心线绘制联结系')
        except:
            print("平面捕捉失败，请重试。")

    def get_User_Orignal_Point_inCAD(self, acadapp):
        self.User_Orignal_Point_inCAD = get_point_and_make_xline(acadapp)

    def on_submit_rib(self):
        
        nodeid = '1'
        elementid = '1'
        make_rib_element_resultlist = []#用于判断是否初始化
        all_nodes = []#用于储存所有节点信息
        all_elements = []#用于储存所有单元信息
        all_beamloads = []#用于储存所有混凝土荷载信息
        all_constructionloads = []#施工荷载
        all_templateloads = []#模板荷载
        nodeidlst = []
        elementidlst = []
        all_group = []
        all_elink = []

        #结合识别截面按钮获取的截面信息和编辑框输入的信息，循环生成小肋组
        for x in range(len(self.secetyname_lst)):

            if make_rib_element_resultlist != []:
                nodeid = make_rib_element_resultlist[0]
                elementid = make_rib_element_resultlist[1]

            make_rib_element_resultlist = make_rib_nodes_elements_loads(

                # self.get_rib_basedata_resultlist[0],#小肋截面信息 rib_info
                [round(y, 1) for y in self.get_rib_basedata_resultlist[2]],#小肋的所有x坐标 rib_xlist
                self.get_rib_basedata_resultlist[3],#小肋荷载宽度表 ribload_distance

                self.secetyname_lst[x],#小肋组号,"x to y" ribi_batchnum

                self.section_xlst1[x],#一组小肋除去悬臂的y坐标 ribi_ylist1
                self.section_xlst2[x],#一组小肋除去截面外贝雷梁的y坐标 ribi_ylist2
                self.section_hlst[x],#一组小肋除去悬臂的y坐标对应的截面有效高度h ribi_hlist

                # self.Distance_between_section_Bailey_Ry_lst[x],#一组小肋中截面距离贝雷梁边线的距离 D
                # self.Bailey_y_dist_lst,#贝雷梁的全部y坐标 Bailey_ylst
                self.rib_Bailey_y_dict,#小肋与贝雷梁弹性连接的y坐标词典，用于判断弹性连接
                
                self.entryname15.get(),#小肋悬臂
                # self.entryname19.get(),#小肋整体移动的距离
                self.entryname12.get(),#混凝土容重
                self.entryname10.get(),#施工荷载
                self.entryname11.get(),#模板荷载

                nodeid,#一组小肋的初始节点号
                elementid,#一组小肋的初始单元号
                )
            
            all_nodes = all_nodes + make_rib_element_resultlist[2]
            all_elements = all_elements + make_rib_element_resultlist[3]

            all_beamloads = all_beamloads + make_rib_element_resultlist[4]
            all_constructionloads = all_constructionloads + make_rib_element_resultlist[5]
            all_templateloads = all_templateloads + make_rib_element_resultlist[6]

            all_group = all_group + make_rib_element_resultlist[7]
            all_elink = all_elink + make_rib_element_resultlist[8]

            nodeidlst.append(make_rib_element_resultlist[0])
            elementidlst.append(make_rib_element_resultlist[1])
        
        rib_group = ["小肋, , {} to {}, 0".format(all_group[0], all_group[-1])]
        print("Success: 成功生成小肋节点、单元、结构组。")
            
        return all_nodes, all_elements, all_beamloads, all_constructionloads, all_templateloads, nodeidlst[-1], elementidlst[-1], rib_group, all_elink
    
    def on_submit_Baiey(self, rib_h, id1, id2, rib_elink_lst, Bailey_Fengya):
        nodeid = id1
        elementid = id2
        # print("贝雷梁词典：", self.Bailey_Rib_FPL_dict)
        # print("分配梁词典：", self.FPL_Bailey_dict)
        reslst = creat_Bailey_node_element_group_brnd(rib_h, nodeid, elementid, self.Bailey_Rib_FPL_dict, self.FPL_Bailey_dict, rib_elink_lst, Bailey_Fengya)
        return reslst[0], reslst[1], reslst[2], reslst[3], reslst[4], reslst[5], reslst[6], reslst[7], reslst[8], reslst[9], reslst[10], reslst[11]

    def on_submit_FPL(self, Bailey_FPL_elink_dict, startnodeid, startelementid, startelinkid):
        # self.SP_dict1 为原始钢管桩数据
        # self.Bailey_Rib_FPL_dict 为贝雷梁词典
        # self.FPL_Bailey_dict 为分配梁词典 
        # 根据 贝雷梁词典的原点 以及 分配梁词典 将钢管桩信息生成新的钢管桩词典，并且更新分配梁词典
        SP_dict = self.SP_dict1.copy()
        reslst1 = update_FPLdict_with_SP(self.Bailey_Rib_FPL_dict, self.FPL_Bailey_dict, SP_dict, self.FPL_seclst1, self.FPL_seclst2)
        self.FPL_Bailey_SP_dict = reslst1[0]
        self.SP_dict2 = reslst1[1]
        self.SEC_LST = reslst1[2]
        # print(self.SP_dict2)
        # 根据更新后的分配梁词典建立分配梁节点、单元、结构组、弹性连接
        reslst2 = creat_FPL_node_element_group_brnd(self.FPL_Bailey_SP_dict, Bailey_FPL_elink_dict, self.SEC_LST, startnodeid, startelementid, startelinkid)
        return reslst2[0], reslst2[1], reslst2[2], reslst2[3], reslst2[4], reslst2[5], reslst2[6], reslst2[7]
    
    def on_submit_SP(self, nodeid, elementid, elinkid, FPL_dict, SP_Fengya, SP_Shui_info):
        reslst = create_SP_LJX_node_element_grup_brnd(self.SP_dict2, nodeid, elementid, elinkid, FPL_dict, self.SEC_LST, SP_Fengya, SP_Shui_info)
        return reslst[0], reslst[1], reslst[2], reslst[3], reslst[4], reslst[5], reslst[6], reslst[7], reslst[8], reslst[9]

    def on_submit(self):
        # 各个复选按钮的控件值
        Midas_CheckButtonVlaue = self.chVarEn1.get()
        CAD_CheckButtonVlaue = self.chVarEn2.get()
        Word_CheckButtonVlaue = self.chVarEn3.get()
        # 当任意按钮被选择时
        if Midas_CheckButtonVlaue != 0 or CAD_CheckButtonVlaue != 0 or Word_CheckButtonVlaue != 0:
            if self.chVarEn1.get() == 1:
                self.on_submit_Midas()
                # print("Midas模型")
            if self.chVarEn2.get() == 1:
                print("功能未开放。")
            if self.chVarEn3.get() == 1:
                print("功能未开放。")
        # 当没有按钮被选择时
        else:
            print("未选择任何内容。")

    def on_submit_Midas(self, file_extension = ".txt"):
        csv_dict = {
            'bailey_Cantilever_length':0,
            'bailey_along_layout_namelst': [],
            'bailey_along_layout_gaplst':[],
            'bailey_cross_layout_gaplst':[],
            'rib_sec_name':'',
            'rib_along_layout_gaplst':[],
            'fenpeiliang_sec_name':[],
            'ganggz_along_layout_gaplst':[],
            'ganggz_cross_layout_gaplst':[],

        }
        csv_dict['rib_along_layout_gaplst'] = group_consecutive_elements(midasdisttolst(self.entryname9.get()))
        csv_dict['rib_sec_name'] = self.selected_value2.get()
        csv_dict['fenpeiliang_sec_name'] = self.selected_value1.get()
        # print(self.Bailey_CAD_to_lst)
        # print(self.FPL_CAD_to_lst)
        # print(self.SP_CAD_to_lst)
        ganggz_along_layout_gaplst_xlst = []
        ganggz_cross_layout_gaplst_ylst = []
        for key, value in self.SP_CAD_to_lst.items():
            if '钢管桩桩顶' in key:
                ganggz_along_layout_gaplst_xlst.append(round(value[1][0],1))
                ganggz_cross_layout_gaplst_ylst.append(round(value[1][1],1))
        ganggz_along_layout_gaplst_xlst = sorted(list(set(ganggz_along_layout_gaplst_xlst)))
        ganggz_cross_layout_gaplst_ylst = sorted(list(set(ganggz_cross_layout_gaplst_ylst)))
        csv_dict['ganggz_along_layout_gaplst'] = group_consecutive_elements(get_distance_from_xlst(ganggz_along_layout_gaplst_xlst))
        csv_dict['ganggz_cross_layout_gaplst'] = group_consecutive_elements(get_distance_from_xlst(ganggz_cross_layout_gaplst_ylst))
        first_fenpeiliang_x = self.FPL_CAD_to_lst['第1根分配梁'][0][0]
        first_bailey_x = self.Bailey_CAD_to_lst['第1排贝雷上弦杆接头坐标'][0][0]
        csv_dict['bailey_Cantilever_length'] = first_fenpeiliang_x-first_bailey_x
        csv_dict['bailey_along_layout_gaplst'] = group_consecutive_elements(get_distance_from_xlst([pt[0] for pt in self.Bailey_CAD_to_lst['第1排贝雷上弦杆接头坐标']]))
        csv_dict['bailey_cross_layout_gaplst'] = group_consecutive_elements(get_distance_from_xlst(sorted(list(set([value[0][1] for key, value in self.Bailey_CAD_to_lst.items() if '排贝雷' in key])))))
        for gaplst in csv_dict['bailey_along_layout_gaplst']:
            if gaplst[0] == 1500:
                csv_dict['bailey_along_layout_namelst'].append('1.5m贝雷梁')
            if gaplst[0] == 3000:
                csv_dict['bailey_along_layout_namelst'].append('贝雷梁')

        # 图形捕捉的数据检查
        # print(self.Bailey_CAD_to_lst)
        # print(self.FPL_CAD_to_lst)
        # print(self.SP_CAD_to_lst)

        # 风荷载参数
        # print(self.Feng_values)
        # 风荷载规范
        # print(self.FengGui_selected_value)
        # 根据规范计算风荷载
        try:
            if self.FengGui_selected_value == "公路桥梁抗风设计规范":
                Feng_result_lst = cal_FengYa_JTG_T_3360_01_2018(self.Feng_values["wind_entry1"], self.Feng_values["wind_entry2"], self.Feng_values["wind_entry3"], self.Feng_values["wind_entry4"], self.Feng_values["wind_entry5"], self.Feng_values["wind_entry6"], self.Feng_values["wind_combo1"], self.Feng_values["wind_combo2"])
                FengYa = Feng_result_lst[10:13]
            elif self.FengGui_selected_value == "工程结构通用规范":
                Feng_result_lst = cal_FengYa_GB_55001_2021(self.Feng_values["wind_entry1"], self.Feng_values["wind_entry2"], self.Feng_values["wind_entry3"], self.Feng_values["wind_entry5"], self.Feng_values["wind_combo1"])
                FengYa = Feng_result_lst[0:3]
            elif self.FengGui_selected_value == "港口工程荷载规范":
                Feng_result_lst = cal_FengYa_JTS_144_1_2010(self.Feng_values["wind_entry1"], self.Feng_values["wind_entry2"], self.Feng_values["wind_entry3"], self.Feng_values["wind_entry5"], self.Feng_values["wind_combo1"])
                FengYa = Feng_result_lst[0:3]
            elif self.FengGui_selected_value == "自定义":
                Feng_result_lst = []
                FengYa = [0, 0, 0]
            # print("贝雷和墩柱风压大小为(kpa):", FengYa)
        except:
            # 模板、贝雷和墩柱风压大小均设为0
            FengYa = [0, 0, 0]
            print("风压计算失败, 请检查输入的参数是否正确")


        # 水荷载参数
        # print(self.Shui_values)
        # 水流力规范
        # print(self.ShuiGui_selected_value)
        if self.ShuiGui_selected_value == "港口工程荷载规范":
            SP_Shui_info = [key_str for key, key_str in self.Shui_values.items()]
        elif self.ShuiGui_selected_value == "自定义":
            SP_Shui_info = ['']

        # 处理结构交叉信息
        try:
            # 结合小肋和贝雷梁
            self.Get_Ribinfo_from_Bailey()
        except:
            print("Warning: 小肋与贝雷信息交互失败。")
        try:
            # 结合贝雷梁和分配梁
            self.Get_FPLinfo_from_Bailey()
        except:
            print("Warning: 贝雷与分配梁信息交互失败。")
        try:
            # 结合分配梁和钢管桩
            self.Get_SPinfo_from_CAD()
        except:
            print("Warning: 分配梁与钢管桩信息交互失败。")

        # 处理截面赋予对话框中的截面信息
        try:
            self.return_secondarydialog_list()
        except:
            print("Warning: 小肋上截面信息处理失败。")

        # 初始化
        list1 = []
        list2 = []
        list3 = []
        list4 = []
        list5 = []

        # 小肋
        list1 = self.on_submit_rib()
        # 贝雷
        list2 = self.on_submit_Baiey(self.get_rib_basedata_resultlist[0][0].split(',')[13], list1[5], list1[6], list1[8], FengYa[1])
        # 分配梁
        list3 = self.on_submit_FPL(list2[7], list2[8], list2[9], list2[10])
        # 钢管桩
        list4 = self.on_submit_SP(list3[4], list3[5], list3[6], list3[7], FengYa[2], SP_Shui_info)
        # 荷载组合
        list5 = loadcomb()

        # 材料
        MATERIAL_lst = MATERIAL()
        # 截面
        SECTION_lst = support_section_data(self.get_rib_basedata_resultlist[0]) + self.SEC_LST
        # 节点
        NODE_lst = list1[0] + list2[0] + list3[0] + list4[0]
        # 单元
        ELEMENT_lst = list1[1] + list2[1] + list3[1] + list4[1]
        # 组
        GROUP_lst = list1[7] + list2[2] + list2[3] + list2[4] + list3[2] + list4[2] + list4[5] + [x[0] for x in list4[6]]
        # 弹性连接
        ELASTICLINK_lst = list2[6] + list3[3] + list4[3]
        # 释放梁端约束
        FRAME_lst = list2[5]
        # 桩底固结
        CONSTRAINT_lst = list4[4]

        # print(MATERIAL_lst)
        # print(SECTION_lst)
        # print(list1[0])
        # print(list1[1])

        # 汇总表
        mct_lst = [
            # 版本号，进行各个版本Midas的mct文件测试
            ["*VERSION", "8.6.5"],
            # 单位系
            ["*UNIT", "N,MM,KJ,C"],
            # 材料
            ["*MATERIAL"] + MATERIAL_lst,
            # 截面
            ["*SECTION"] + SECTION_lst,
            # 节点
            ["*NODE"] + NODE_lst,
            # 单元
            ["*ELEMENT"] + ELEMENT_lst,
            # 结构组
            ["*GROUP"] + GROUP_lst,
            # 约束组
            ["*BNDR-GROUP", "弹性连接, 0", "桩底反力, 0", "释放梁端约束, 0"],
            # 弹性连接
            ["*ELASTICLINK"] + ELASTICLINK_lst,
            # 释放梁端约束
            ["*FRAME-RLS"] + FRAME_lst,
            # 桩底固结
            ["*CONSTRAINT"] + CONSTRAINT_lst,
            # 荷载组
            ["*LOAD-GROUP", "自重", "混凝土荷载", "施工荷载", "模板荷载"],
            # 荷载说明
            ["*STLDCASE", "自重, USER,", "混凝土荷载, USER,", "施工荷载, USER,", "模板荷载, USER,", "风荷载, USER,", "水流力, USER,"],
            # 自重
            ["*USE-STLD, 自重", "*SELFWEIGHT, 0, 0, {}, 自重".format(self.entryname13.get())],
            # 混凝土荷载
            ["*USE-STLD, 混凝土荷载", "*BEAMLOAD"] + list1[2],
            # 施工荷载
            ["*USE-STLD, 施工荷载", "*BEAMLOAD"] + list1[3],
            # 模板荷载
            ["*USE-STLD, 模板荷载", "*BEAMLOAD"] + list1[4],
            # 风荷载
            ["*USE-STLD, 风荷载", "*BEAMLOAD"] + list2[11] + list4[7],
            # 水流力
            ["*USE-STLD, 水流力", "*BEAMLOAD"] + list4[8],
            # 荷载组合
            ["*LOADCOMB"] + list5,
        ]
        print("Success: 已生成MCT命令。")
        # 换行 
        mct_lst_tolines = "\n".join([item for sublist in mct_lst for item in sublist])
        # 复制粘贴
        # pyperclip.copy(mct_lst_tolines)
        SET_CLIP_STRING(mct_lst_tolines)
        print("mct命令流已粘贴到剪切板中")
        # 根据指定路径生成mct文件
        write_file_within_path(self.mct_savepath.get(), mct_lst_tolines, "gbk")
        # 生成mcb模型
        MidasURL()
        mcb_savepath = file_extension_Modified(self.mct_savepath.get(), '.mcb')
        importmct_to_mcb(self.mct_savepath.get(), mcb_savepath)

        # print("粘贴板已复制")
        # 得到静态编辑框的值
        # print(self.entryname9.get())
        # print(self.entryname10.get())
        # print(self.entryname11.get())
        # print(self.entryname12.get())
        # print(self.entryname19.get())
        # 得到动态编辑框的值，Ecc是截面偏心值
        # print(globals()['Ecc0'].get())
        # print(globals()['Ecc1'].get())
        # print(globals()['Ecc2'].get())

        # txt文件文本内容
        txt_lst = [
            # 风荷载规范
            ['*风荷载规范'] + [self.FengGui_selected_value], 
            # 水流力规范
            ['*水流力规范'] + [self.ShuiGui_selected_value], 
            # 自重参数
            ['*自重参数'] + [self.entryname13.get()], 
            # 混凝土荷载参数
            ['*混凝土荷载参数'] + [self.entryname12.get()], 
            # 模板荷载参数
            ['*模板荷载参数'] + [self.entryname11.get()], 
            # 施工荷载参数
            ['*施工荷载参数'] + [self.entryname10.get()], 
            # 风荷载参数
            # ['*风荷载参数'] + [', '.join([x for x in self.Feng_values.values()]+[str(y) for y in Feng_result_lst])], 
            ['*风荷载参数', "None"] if self.FengGui_selected_value == "自定义" else ['*风荷载参数', ', '.join(str(y) for x in [self.Feng_values.values(), Feng_result_lst] for y in x )], 
            # 水流力参数
            # ['*水流力参数'] + [', '.join([x for x in self.Shui_values.values()]+ [str(y) for y in list4[9]])], 
            ['*水流力参数', "None"] if self.ShuiGui_selected_value == "自定义" else ['*水流力参数', ', '.join(str(y) for x in [self.Shui_values.values(), list4[9]] for y in x )], 
            # 贝雷梁材料特性参数
            ['*贝雷梁材料特性参数'] + ['None'], 
            # 钢材材料特性参数
            ['*钢材材料特性参数'] + ['None'], 
            # 混凝土材料特性参数
            ['*混凝土材料特性参数'] + ['None'], 
            # 钢筋材料特性参数
            ['*钢筋材料特性参数'] + ['None'], 
            # 螺栓材料特性参数
            ['*螺栓材料特性参数'] + ['None'], 
            # 焊缝材料特性参数
            ['*焊缝材料特性参数'] + ['None'], 
            # 竹胶板材料特性参数
            ['*竹胶板材料特性参数'] + ['None'], 
            # 方木材料特性参数
            ['*方木材料特性参数'] + ['None'], 
        ]
        # 列表换行处理
        txt_lst_tolines = "\n".join([item for sublist in txt_lst for item in sublist])
        # txt文件保存完整路径
        # txt_savepath = file_extension_Modified(self.mct_savepath.get(), ".txt")
        # txt_savepath = file_extension_Modified(self.mct_savepath.get(), ".cfg")
        txt_savepath = file_extension_Modified(self.mct_savepath.get(), file_extension)
        # 根据指定路径生成txt文件
        write_file_within_path(txt_savepath, txt_lst_tolines, "utf-8")

        bailey_Cantilever_length, bailey_along_layout_namelst, bailey_along_layout_gaplst, bailey_cross_layout_gaplst, rib_sec_name, rib_along_layout_gaplst, fenpeiliang_sec_name, ganggz_along_layout_gaplst, ganggz_cross_layout_gaplst = [value for value in csv_dict.values()]
        csv_lst = CSV_Bracket(bailey_Cantilever_length, bailey_along_layout_namelst, bailey_along_layout_gaplst, bailey_cross_layout_gaplst,
                              rib_sec_name, rib_along_layout_gaplst, fenpeiliang_sec_name, ganggz_along_layout_gaplst, ganggz_cross_layout_gaplst
                             )
        write_csv('Bracket.csv', csv_lst)


    def BIMBase(MATERIAL_lst, SECTION_lst, NODE_lst, ELEMENT_lst):
        # BIMBase信息流
        # 将材质整合成字典，形式为"材质号":["Q235"]
        MATERIAL_lst = {"1":["Q235"], "2":["16Mn"]}
        # 将截面整合成字典，形式为"截面号":[截面信息]
        SECTION_dict = {}
        # 将节点信息整合成字典，形式为"节点号":[0,0,0]
        NODE_dict = {}
        for sstr in NODE_lst:
            # 去除字符串两端的引号和空格
            sstr_strip = sstr.strip("'")
            # 分割字符串得到各部分
            sstr_parts = [p.strip() for p in sstr_strip.split(',')]
            # 提取节点号和坐标
            node_id = sstr_parts[0]
            coordinates = [float(p) for p in sstr_parts[1:]]
            # 添加到字典
            NODE_dict[node_id] = coordinates
        # 将单元信息整合成字典，形式为"单元号":["材质号","截面号","节点号1","节点号2","旋转角度"]
        ELEMENT_dict = {}
        for sstr in ELEMENT_lst:
            # 去除字符串两端的引号和空格
            sstr_strip = sstr.strip("'")
            # 分割字符串得到各部分
            sstr_parts = [p.strip() for p in sstr_strip.split(',')]
            # 提取节点号和坐标
            elemrnt_id = sstr_parts[0]
            # material_id = sstr_parts[1]
            # section_id = sstr_parts[2]
            # node_id1 = sstr_parts[3]
            # node_id2 = sstr_parts[3]
            # rotate_angle = sstr_parts[5]
            # 添加到字典
            ELEMENT_dict[elemrnt_id] = [p for p in sstr_parts[1:]]

    def on_cancel(self):
        # self.quit()
        self.winfo_toplevel().destroy()

    def on_save(self):
        pass


def CSV_Bracket(bailey_Cantilever_length, bailey_along_layout_namelst, bailey_along_layout_gaplst, bailey_cross_layout_gaplst,
                rib_sec_name, rib_along_layout_gaplst, fenpeiliang_sec_name, ganggz_along_layout_gaplst, ganggz_cross_layout_gaplst, 
               ):
    '''
    支架的csv文件格式
    '''
    line1 = ['贝雷梁', '伸出长度：']
    line2 = ['', '类型：']
    line3 = ['', '纵向：']
    line4 = ['', '横向：']
    line5 = ['大肋', '名称：']
    line6 = ['', '间距：']
    line7 = ['分配梁', '名称：']
    line8 = ['钢管桩', '纵向：']
    line9 = ['', '横向：']
    line10 = ['横向联结系', '竖向：']
    line11 = ['纵向联结系', '竖向：']
    # 写入数据
    # line1
    # bailey_Cantilever_length = 1500 # 贝雷悬出长度
    line1.append(bailey_Cantilever_length)
    # line2
    # bailey_along_layout_namelst = ['1.5m贝雷梁', '贝雷梁']
    for x in bailey_along_layout_namelst:
        line2.extend([str(x), '', '', ''])
    # line3
    # bailey_along_layout_gaplst  = [[1500], [3000, 3000, 3000, 3000, 3000, 3000, 3000, 3000]]
    for x in bailey_along_layout_gaplst:
        line3.extend([str(x[0]), '@', str(len(x)), ','])
    line3[-1] = '/'
    # line4
    # bailey_cross_layout_gaplst  = [[900, 900], [450, 450, 450], [900, 900], [450]]
    for x in bailey_cross_layout_gaplst:
        line4.extend([str(x[0]), '@', str(len(x)), ','])
    line4[-1] = '/'
    # line5
    # rib_sec_name = 'I10'
    line5.append(rib_sec_name)
    # line6
    # rib_along_layout_gaplst  = [[1000, 1000, 1000, 1000, 1000], [500, 500, 500, 500, 500]]
    for x in rib_along_layout_gaplst:
        line6.extend([str(x[0]), '@', str(len(x)), ','])
    line6[-1] = '/'
    # line7
    # fenpeiliang_sec_name = '2I32a'
    line7.append(fenpeiliang_sec_name)
    # line8
    # ganggz_along_layout_gaplst  = [[12000, 12000], [15000, 15000, 15000]]
    for x in ganggz_along_layout_gaplst:
        line8.extend([str(x[0]), '@', str(len(x)), ','])
    line8[-1] = '/'
    # line9
    # ganggz_cross_layout_gaplst  = [[3000, 3000]]
    for xlst in ganggz_cross_layout_gaplst:
        for x in xlst:
            line9.extend([str(x), '@', '1', ','])
    line9[-1] = '/'
    # line10
    line10.extend(['2000', '@', '1'])
    # line11
    line11.extend(['2000', '@', '1'])
    # 同化格式
    linelst = [line1, line2, line3, line4, line5, line6, line7, line8, line9, line10, line11]
    csv_line_length = max(len(line1), len(line2), len(line3), len(line4), len(line5), len(line6), len(line7), len(line8), len(line9), len(line10), len(line11))
    for lst in linelst:
        lsti = csv_line_length - len(lst)
        for i in range(lsti):
            lst.append('')
    # for x in linelst:
    #     print(x)
    return linelst


# 计算书另存为得路径
def mct_path_SaveAs(entry_var):
    # 选择路径
    savefile_path = qucik_save_file_dialog('.mct')
    # 更新对应编辑框的值
    entry_var.set(savefile_path)
