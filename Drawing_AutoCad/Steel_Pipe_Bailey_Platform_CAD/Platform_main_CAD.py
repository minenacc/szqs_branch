# 1. 标准库
import sys
import traceback
from pathlib import Path
from tkinter import messagebox

# 2. 第三方库
import win32com.client
import ttkbootstrap as tb
from ttkbootstrap.constants import *

# 3. 本地模块
sys.path.append(str(Path(__file__).parent.parent.parent))
from General.AutoCAD        import Import_Blocks, find_installed_autocad_versions, AutoCAD_VersionR_tans_to_COM_ProgID
from General.DataUtils      import read_Register
from General.UIHandle       import if_Reg, show_warning_dialog

sys.path.append(str(Path(__file__).parent))
from Drawing_AutoCad.Steel_Pipe_Bailey_Platform_CAD.Platform_CAD_draw import All
from Drawing_AutoCad.Steel_Pipe_Bailey_Platform_CAD.Platform_CAD_UI import PlatformUI


# 展开嵌套列表
def flatten_to_str_lines(nested):
    """递归展开任意层嵌套列表，并转为字符串"""
    lines = []
    for item in nested:
        if isinstance(item, (list, tuple)):
            lines.extend(flatten_to_str_lines(item))
        else:
            lines.append(str(item))
    return lines


def run_cad(app, acadapp):
    print("CAD 绘图")
    try:
        All(app, acadapp)
    except Exception as e:
        messagebox.showerror("错误", f"CAD 绘图失败：{e}")
        traceback.print_exc()


# if __name__ == "__main__":
def PlatForm_Drawing_AutoCad_Main(parent=None, on_close=None):
    if not if_Reg(parent):
        if on_close:
            on_close()
        return
    FileName = 'ShuZhiQiaoShi'
    Applocation = read_Register('Software\\{}'.format(FileName), 'Applocation')
    print("正在搜索已安装的AutoCAD版本...")
    VersionR_lst = find_installed_autocad_versions()
    versionR_lst = sorted(VersionR_lst, key = lambda x : x.replace('R', ''), reverse = True)
    print(versionR_lst)

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
        if parent == None:
            root = tb.Window(themename="cosmo")
        else:
            root = tb.Toplevel(parent)
        app = PlatformUI(root)
        app.run_cad_btn.config(command= lambda:run_cad(app, acadapp))
        if parent == None:
            root.mainloop()
        elif on_close:
            def _on_root_destroy(event, _root=root):
                if event.widget is _root:
                    on_close()
            root.bind('<Destroy>', _on_root_destroy)
    except:
        print('AutoCAD版本号检索失败, 尝试直接启动连接')
        try:
            acadapp = win32com.client.Dispatch('AutoCAD.Application')
            acadapp.Visible = True
            print('AutoCAD连接成功')
            Applocation = read_Register('Software\\ShuZhiQiaoShi', 'Applocation') # 软件安装路径
            Import_Blocks(Applocation, acadapp)
            if parent == None:
                root = tb.Window(themename="cosmo")
            else:
                root = tb.Toplevel(parent)
            app = PlatformUI(root)
            app.run_cad_btn.config(command= lambda:run_cad(app, acadapp))
            if parent == None:
                root.mainloop()
            elif on_close:
                def _on_root_destroy2(event, _root=root):
                    if event.widget is _root:
                        on_close()
                root.bind('<Destroy>', _on_root_destroy2)
        except:
            print('AutoCAD连接失败')
            print('请手动启动AutoCAD')
            if on_close:
                on_close()

# PlatForm_Drawing_AutoCad_Main()