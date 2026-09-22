# 1. 标准库
import os

# 3. 本地模块
from General.Midas    import MidasAPI
from General.FilePath import qucik_open_file_dialog, qucik_save_file_dialog


# 在Midas已经启动的前提下打开选定的文件
def Open_Midas_mcb(entry_var):
    # 选择路径
    openfile_path = qucik_open_file_dialog()
    # 更新对应编辑框的值
    entry_var.set(openfile_path)
    if "mcb" in openfile_path:
        arguments = {"Argument" : openfile_path}
        MidasAPI("POST" , "/doc/open" , arguments)
    else:
        print("未识别到有效mcb文件, 请重新选择")


# 判断当前mcb文件所在文件夹内是否有计算结果文件，没有则运行
def Midas_Analysis(folder_name, folder_path):
    # 根据文件名生成结果文件名
    Analysis_result_file_name1 = folder_name.replace(".mcb", ".OUT")
    Analysis_result_file_name2 = folder_name.replace(".mcb", ".CA1")
    print(Analysis_result_file_name1)
    print(Analysis_result_file_name2)
    # 生成对应的路径
    Analysis_result_file_path1 = os.path.join(folder_path, Analysis_result_file_name1)
    Analysis_result_file_path2 = os.path.join(folder_path, Analysis_result_file_name2)
    # 文件运行
    if os.path.exists(Analysis_result_file_path1) and os.path.exists(Analysis_result_file_path2):
        # 两个结果文件存在，跳过不运行
        print("检测到结果文件，正在读取...")
    else:
        # 两个结果文件不存在，运行
        print("未检测到结果文件，开始运行。")
        Analysis_post = MidasAPI("POST", "/doc/Anal", {})


# 计算书另存为得路径
def Cal_path_SaveAs(entry_var, Applocation):
    # 选择路径
    savefile_path = qucik_save_file_dialog('.docx').replace("/", "\\")
    # open_csv_saveas_xlsx(savefile_path, Applocation)
    # 更新对应编辑框的值
    entry_var.set(savefile_path)
