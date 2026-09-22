# 1. 标准库
import sys
from pathlib import Path

# 2. 第三方库
import win32com.client

# 3. 本地模块
sys.path.append(str(Path(__file__).parent.parent.parent))
from General.UIHandle  import if_Reg
from General.DataUtils import read_Register
from General.AutoCAD   import find_installed_autocad_versions, AutoCAD_VersionR_tans_to_COM_ProgID, Import_Blocks

sys.path.append(str(Path(__file__).parent))
from FEM_MidasCivil.Steel_Pipe_Bailey_Support_FEA.Support_Midas_Model_UI import DataEntryForm, tb


def BaileySupport_FEM_MidasCivil_Main(parent=None, on_close=None):
    # 注册程序
    if not if_Reg(parent):
        if on_close:
            on_close()
        return
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
        # 外部引用dwg文件
        Applocation = read_Register('Software\\ShuZhiQiaoShi', 'Applocation') # 软件安装路径
        Import_Blocks(Applocation, acadapp)
        # 创建窗口   
        if parent == None:
            root = tb.Window("钢管贝雷现浇支架自动化建模程序")
        else:
            root = tb.Toplevel(parent)
            root.title("钢管贝雷现浇支架自动化建模程序")
        # 对话框类
        creat_dialog = DataEntryForm(root, Applocation)  
        # 调用对话框类中的对话框主函数
        creat_dialog.creat_dialog(acadapp, Applocation)
        if parent is None:
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
            # 外部引用dwg文件
            Applocation = read_Register('Software\\ShuZhiQiaoShi', 'Applocation') # 软件安装路径
            Import_Blocks(Applocation, acadapp)
            # 创建窗口
            if parent == None:
                root = tb.Window("钢管贝雷梁支架参数化")
            else:
                root = tb.Toplevel(parent)
            # 对话框类
            creat_dialog = DataEntryForm(root, Applocation)
            # 调用对话框类中的对话框主函数
            creat_dialog.creat_dialog(acadapp, Applocation)
            if parent is None:
                root.mainloop()
            elif on_close:
                def _on_root_destroy2(event, _root=root):
                    if event.widget is _root:
                        on_close()
                root.bind('<Destroy>', _on_root_destroy2)
        except:
            print('AutoCAD连接失败')
            print('请手动启动AutoCAD后重试')
            if on_close:
                on_close()


# BaileySupport_FEM_MidasCivil_Main()