# 1. 标准库
import os
import sys
from pathlib import Path

# 2. 第三方库
import win32com.client

# 3. 本地模块
sys.path.append(str(Path(__file__).parent.parent.parent))
from General.DataUtils import read_Register
from General.AutoCAD     import find_installed_autocad_versions, AutoCAD_VersionR_tans_to_COM_ProgID, Import_Blocks


def BaileySupport_Drawing_AutoCad_Main(parent=None, on_close=None):
    FileName = 'ShuZhiQiaoShi'
    Applocation = read_Register('Software\\{}'.format(FileName), 'Applocation')
    # current_dir = os.path.dirname(os.path.abspath(sys.argv[0])) # 当前py文件所在的文件夹的路径
    fas_path = os.path.join(Applocation, 'Drawing_AutoCad', 'Steel_Pipe_Bailey_Support_CAD', 'DraftFlow.fas')

    print("正在搜索已安装的AutoCAD版本...")
    VersionR_lst = find_installed_autocad_versions()
    versionR_lst = sorted(VersionR_lst, key = lambda x : x.replace('R', ''), reverse = True)
    print(versionR_lst)

    # 确保文件存在
    if not os.path.exists(fas_path):
        raise FileNotFoundError(f"FAS文件不存在: {fas_path}")
    # 连接到AutoCAD
    try:
        for Rx in versionR_lst:
            try:
                ProgID = AutoCAD_VersionR_tans_to_COM_ProgID(Rx)
                print(ProgID)
                acadapp = win32com.client.Dispatch(ProgID)
                acadapp.Visible = True
                print(f'AutoCAD\'{Rx}\'连接成功')
                break
            except:
                pass
        Import_Blocks(Applocation, acadapp)
    except:
        print('AutoCAD版本号检索失败, 尝试直接启动连接')
        try:
            acadapp = win32com.client.Dispatch('AutoCAD.Application')
            acadapp.Visible = True
            print('AutoCAD连接成功')
            Import_Blocks(Applocation, acadapp)
        except:
            print('AutoCAD连接失败')
            print('请手动启动AutoCAD') 
    # 获取当前文档
    try:
        doc = acadapp.ActiveDocument
    except:
        print("未成功获取ActiveDocument对象, 请新建文件")
    # 发送LISP命令加载FAS文件
    try:
        # load命令加载
        lisp_command = f'(load "{fas_path.replace("\\", "/")}")\n'
        doc.SendCommand(lisp_command)
        doc.SendCommand('C2M\n')
        print(f"成功加载FAS文件: {fas_path}")
    except Exception as e:
        print("加载FAS文件失败")
    # 无二级窗口，完成后立即回调恢复主窗口
    if on_close:
        on_close()
