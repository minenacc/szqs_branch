import time
import pickle
import os
import traceback
import threading
import pythoncom
import winreg
from pathlib import Path
from win32com.client.gencache import EnsureDispatch

# 动态路径
import sys
# 获取当前文件的路径
current_path = Path(__file__).resolve()
# 获取上级目录的路径
parent_path = current_path.parent
# 获取所需目录的路径
Word_directory = Path(r'E:\SZQS\CalRpt_MSWord\Steel_Pipe_Bailey_Support_Cal')
# 将目录添加到sys.path
sys.path.append(str(Word_directory))
print(sys.path)

from Cal_File_Manipulation import MidasURL, MidasAPI, Program_State_Monitoring
from Cal_Report_Generate import *

TASK_FOLDER = os.path.join(os.path.dirname(__file__), "tasks")
INDEX_FILE = os.path.join(TASK_FOLDER, "task_index.pkl")
os.makedirs(TASK_FOLDER, exist_ok=True)
index_lock = threading.Lock()
POLL_INTERVAL = 5  # 每5秒检查一次是否有新任务


# 保存单个任务
def save_task(task):
    task_file = os.path.join(TASK_FOLDER, f"task_{task['task_id']}.pkl")
    with open(task_file, "wb") as f:
        pickle.dump(task, f)

# 加载单个任务
def load_task(task_id):
    task_file = os.path.join(TASK_FOLDER, f"task_{task_id}.pkl")
    if os.path.exists(task_file):
        with open(task_file, "rb") as f:
            return pickle.load(f)
    return None

# 加载任务索引列表
def load_task_index():
    if os.path.exists(INDEX_FILE):
        with open(INDEX_FILE, "rb") as f:
            return pickle.load(f)
    return []

# 保存任务索引
def save_task_index(index_list):
    with open(INDEX_FILE, "wb") as f:
        pickle.dump(index_list, f)

# 添加任务到索引（上传任务时调用）
def add_task_id(task_id):
    with index_lock:
        index_list = load_task_index()
        if task_id not in index_list:
            index_list.append(task_id)
            save_task_index(index_list)

# 打开指定MCB文件
def open_midas_mcb_file(mcb_path: str):
    if mcb_path.lower().endswith(".mcb") and os.path.exists(mcb_path):
        arguments = {"Argument": mcb_path}
        res = MidasAPI("POST", "/doc/open", arguments)
        # 等待模型加载
        print(f"已发送打开 {mcb_path} 的指令给 Midas Civil NX。")
        time.sleep(10)
        return res
    else:
        print("未识别到有效的 .mcb 文件路径，请检查文件路径是否正确。")

# 任务执行函数
def run_task(task):
    task_id = task["task_id"]
    try:
        # 确保为绝对路径
        mcb_path = mcb_abs_path = os.path.abspath(task["mcb_path"])
        foundation_name = task["foundation_name"]
        foundation_standard = task["foundation_standard"]
        foundation_values = task["foundation_values"]
        model_post_result = [{}]  # 用于传递模型读取结果

        temp_dir = os.path.abspath(os.path.join("task", task_id))  # 保证是绝对路径
        os.makedirs(temp_dir, exist_ok=True)

        mcb_name = os.path.splitext(os.path.basename(mcb_path))[0]
        docx_output_path = os.path.join(temp_dir, f"{mcb_name}_cal_doc.docx")
        print("计算书另存为路径：", docx_output_path)

        # 获取base_url和reg_Key
        result = MidasURL()
        if not result:
            raise RuntimeError("无法获取 MIDAS Civil 注册表信息，请检查是否正确安装并打开 Civil")
        base_url, reg_key = result
        
        # 注册表路径
        reg_path2 = winreg.OpenKey(winreg.HKEY_CURRENT_USER,r"SOFTWARE\MIDAS\CVLwNX_CH\PATH")
        civilnx_path = winreg.QueryValueEx(reg_path2,"Installed Path")[0]
        # 窗口名
        partial_title = "MIDAS CIVIL NX"  
        # Midas进程检测，未安装/已打开/未打开
        Program_State_Monitoring(civilnx_path, partial_title)

        # 打开Midas文件
        open_midas_mcb_file(mcb_path)
        pythoncom.CoInitialize()
        # 执行计算函数
        # 初始化COM

        # 调用计算函数
        make_all_doc(
            response = True,
            openfile_path = mcb_path, 
            model_post_result = model_post_result, 
            foundation_name = foundation_name,
            foundation_standard = foundation_standard,
            foundation_value_dict = foundation_values,
            saveas_path = docx_output_path, 
            docx_save_path = "", 
            file_extension = ".txt", 
        )
        # 释放COM
        pythoncom.CoUninitialize()
        # 判断结果是否成功
        if model_post_result[0] == {}:
            raise ValueError("模型结果读取失败")
        else:
            print("模型结果读取成功")

        # 计算书路径整理及状态变更
        task["status"] = "done"
        task["msg"] = "计算完成，计算书已生成"
        task["docx_path"] = docx_output_path
        save_task(task)
        

    except Exception as e:
        task["status"] = "error"
        task["msg"] = "任务执行失败"
        task["error"] = traceback.format_exc()
        save_task(task)
        print(f"[任务错误] {task_id}: {e}")

    finally:
        # 保存并关闭Midas当前任务窗口
        res = MidasAPI("POST", "/doc/save", {})
        time.sleep(5)
        res = MidasAPI("POST", "/doc/close", {})
        print("Midas任务窗口已关闭。")

        # 删除任务临时目录
        #shutil.rmtree(task["temp_dir"])
        #print(f"[清理] 已删除任务文件夹：{task['temp_dir']}")
        
def run_task_loop():
    print("任务处理线程启动中...")
    while True:
        try:
            index_list = load_task_index()
            for task_id in index_list:
                task = load_task(task_id)
                if not task or task["status"] != "queued":
                    continue

                print(f"[任务线程] 正在处理任务 {task_id}")

                # 标记为 running 并保存
                task["status"] = "running"
                task["msg"] = "正在分析模型并生成计算书..."
                save_task(task)

                try:
                    run_task(task)  # 主计算逻辑
                    task["status"] = "done"
                    task["msg"] = "计算完成，计算书已生成"
                    save_task(task)
                except Exception as e:
                    task["status"] = "error"
                    task["msg"] = "任务执行失败"
                    task["error"] = traceback.format_exc()
                    save_task(task)

                time.sleep(1)  # 防止过于频繁
        except Exception as e:
            print(f"[任务线程] 错误: {e}")
            traceback.print_exc()
       
        time.sleep(POLL_INTERVAL) # 轮询间隔        
      