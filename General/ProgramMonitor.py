# 2. 第三方库
import psutil
import win32gui
import win32process

# 检测excel是否在打开
def is_excel_open(file_path):
    for proc in psutil.process_iter():
        try:
            if"EXCEL.EXE" in proc.name():
                for item in proc.open_files():
                    if file_path in item.path:
                        return True
        except:
            pass
    return False


# 获取前台应用程序
def get_foreground_windows_name():
    # 获取前台应用程序窗口的句柄和pid
    def callback(hwnd, hwnd_list: list):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowTextLength(hwnd) > 0:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            hwnd_list.append((hwnd, pid))
        return True
    hwnd_list = []
    win32gui.EnumWindows(callback, hwnd_list)
    # 根据句柄查询应用窗口名
    hwnd_name_list = [win32gui.GetWindowText(hwndx[0]) for hwndx in hwnd_list]
    return hwnd_name_list


# 检查是否有CAD窗口
def if_AutoCAD():
    windows_name_list = get_foreground_windows_name()
    # print(windows_name_list)
    error_massage = ''
    for name in windows_name_list:
        if 'AutoCAD' in name:
            print('检测到CAD应用程序窗口。')
            return [True, error_massage]
    error_massage = '未检测到CAD应用程序窗口，请先打开CAD并选择文件'   
    return [False, error_massage]


# 检查是否有窗口标题包含 ""
def is_window_open_with_partial_title(partial_title):
    """
    检查是否有窗口的标题包含指定的字符串
    :param partial_title: 要查找的部分标题字符串
    :return: 如果找到包含该字符串的窗口，返回 True；否则返回 False
    """
    def enum_windows_callback(hwnd, extra):
        if win32gui.IsWindowVisible(hwnd):  # 只检查可见的窗口
            window_title = win32gui.GetWindowText(hwnd)
            if partial_title in window_title:  # 判断标题是否包含指定字符串
                extra.append(hwnd)  # 将匹配的窗口句柄保存到列表中
        return True
    matching_windows = []  # 用于保存匹配的窗口句柄
    win32gui.EnumWindows(enum_windows_callback, matching_windows)  # 枚举所有窗口
    return len(matching_windows) > 0  # 如果有匹配的窗口，返回 True


# 进程检测，未安装/已打开/未打开
def Program_State_Monitoring(civilnx_path, partial_title):
    if not civilnx_path:
        print("注册表检索失败，请检查是否安装Midas！")
        return "未安装"
    # 打开
    if is_window_open_with_partial_title(partial_title):
        print("窗口已找到，退出等待并继续执行其他操作...")
        return "已打开"
    else:
        print("窗口未找到，请选择下一步操作...")
        return "未打开"