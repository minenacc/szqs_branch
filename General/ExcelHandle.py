# 1. 标准库
import os
import csv
import time
import winreg
from tkinter import messagebox

# 2. 第三方库
import win32com.client

# 3. 本地模块
from General.ProgramMonitor import is_excel_open


# 与excel建立连接
def ExcelApp_Dispatch(State = False):
    # 建立连接
    exApp = win32com.client.Dispatch("Excel.Application")
    exApp.Visible = State
    exApp.DisplayAlerts = State
    return exApp


# 退出与excel的连接
def ExcelApp_Quit(exApp):
    exApp.Quit()


# 用win32来与excel交互
def read_excel_to_dict(excel, excel_path, State = False):
    excel.Visible = State
    data_list = []
    try:
        # 打开工作表格
        wb = excel.Workbooks.Open(os.path.abspath(excel_path))
        sheet = wb.ActiveSheet  # 获取活动工作表
        # 获取已使用区域
        used_range = sheet.UsedRange
        # 转成元组
        data_tuple = used_range.Value
        # 转成列表
        data_list = [list(x) for x in data_tuple]
        if wb:
            wb.Close(False)  # 不保存更改
    except:
        print("发生未知错误，请重新尝试")
    return data_list


def read_excel_to_dict_by_Sheet(WorkBook, sheet_specifier=None):
    """
    读取 Excel 文件指定工作表的内容，返回二维列表。
    参数:
        sheet_specifier: 工作表标识，可以是：
            - 字符串：工作表名称，如 "Sheet1"
            - 整数：工作表索引（从1开始），如 1
            - None：使用当前活动工作表（默认行为）
    返回:
        二维列表，每行是一个列表
    """
    data_list = []
    # 根据 sheet_specifier 获取工作表
    if sheet_specifier is None:
        sheet = WorkBook.ActiveSheet
    elif isinstance(sheet_specifier, str):
        sheet = WorkBook.Sheets(sheet_specifier)  # 按名称获取
    elif isinstance(sheet_specifier, int):
        sheet = WorkBook.Sheets(sheet_specifier)  # 按索引获取（从1开始）
    used_range = sheet.UsedRange
    data_tuple = used_range.Value
    data_list = [list(x) for x in data_tuple]

    return data_list


def write_csv(csv_filename, csv_lst):
    # 读取注册表路径
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"SOFTWARE\ShuZhiQiaoShi", 0, winreg.KEY_READ) as key:
        parent_path, value_type = winreg.QueryValueEx(key, 'Applocation')
    # 指定csv保存路径
    csv_savepath = os.path.join(parent_path, r'BIM_Revit\SZQS-excel', csv_filename)
    # print(csv_savepath)
    with open(csv_savepath, 'w', newline='', encoding='gbk') as file:
        writer = csv.writer(file)
        for lst in csv_lst:
            writer.writerow(lst)


# 获取土层参数表信息
def get_soild_excel_data_tolst(Steel_Sheet_Pile_CofferDam_Load_dict, excel_path):
    if os.path.isfile(excel_path):
        print(f'文件存在:{excel_path}')
        excel = ExcelApp_Dispatch()
        excel_data_lst = []
        response = messagebox.askyesno("确认", "成功读取到文件，是否打开编辑")
        if response:
            excel.Visible = True
            wb = excel.Workbooks.Open(os.path.abspath(excel_path))
            while True:
                # 每隔0.5s判断excel是否关闭
                time.sleep(0.5)
                drive, rest = os.path.splitdrive(excel_path)
                excel_path = drive.upper() + rest
                if not is_excel_open(excel_path):
                    print('excel已关闭, 程序正在尝试读取表格信息')
                    break
        else:
            excel.Visible = False
        excel_data_lst = read_excel_to_dict(excel, excel_path)
        print(excel_data_lst)    
        # ecxel数据字典
        excel_dict = {}
        # 去掉第一行
        layer_data = excel_data_lst[1:]
        for i in range(len(layer_data)):
            excel_dict[f'第{i+1}层土'] = {'土层名称':layer_data[i][0], '土层性质':layer_data[i][1], '层厚':layer_data[i][2], '重度':layer_data[i][3], '黏聚力':layer_data[i][4], '内摩擦角':layer_data[i][5]}
        print(excel_dict)
        Steel_Sheet_Pile_CofferDam_Load_dict['Excel'] = excel_dict
        print('Excel结果已读取')
        ExcelApp_Quit(excel)
    else:
        print(f'文件不存在:{excel_path}, 请检查资源')


# 获取土层参数表信息(现浇支架土层信息表)
def ZJ_get_soild_excel_data_tolst(Soil_Layer_Parameter_Table, excel_path):
    if os.path.isfile(excel_path):
        print(f'文件存在:{excel_path}')
        excel = ExcelApp_Dispatch()
        excel_data_lst = []
        response = messagebox.askyesno("确认", "成功读取到文件，是否打开编辑")
        if response:
            excel.Visible = True
            wb = excel.Workbooks.Open(os.path.abspath(excel_path))
            while True:
                # 每隔0.5s判断excel是否关闭
                time.sleep(0.5)
                drive, rest = os.path.splitdrive(excel_path)
                excel_path = drive.upper() + rest
                if not is_excel_open(excel_path):
                    print('excel已关闭, 程序正在尝试读取表格信息')
                    break
        else:
            excel.Visible = False
        excel_data_lst = read_excel_to_dict(excel, excel_path)
        # 表不为空
        if excel_data_lst != []:
            for key in Soil_Layer_Parameter_Table.keys():
                for i, table_str in enumerate(excel_data_lst[0]):
                    if key in table_str:
                        Soil_Layer_Parameter_Table[key] = [x[i] for x in excel_data_lst[1:]]
            # 检查赋值后的列表中是否存在空值
            if not all(None not in value for value in Soil_Layer_Parameter_Table.values()):
                print("未获取有效信息或矩阵存在空值, 请确保文件或数据正确")
                print(Soil_Layer_Parameter_Table)
            else:
                print(Soil_Layer_Parameter_Table)
        else:
            print("Excel读取结果为空表, 请确保文件或数据正确")  
        ExcelApp_Quit(excel)
    else:
        print(f'文件不存在:{excel_path}, 请检查资源')
