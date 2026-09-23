# 1. 标准库
import sys
import winreg
from pathlib import Path

# 2. 第三方库
import ttkbootstrap as tb

# 3. 本地模块
sys.path.append(str(Path(__file__).parent))
from General.UIFloating import FloatingBallApp_ttk

# ================================================================
# 模块引入按 UI 分类组织：功能为页面（建模/画图/计算书），结构为入口
# ================================================================

# ---------- 建模功能（Midas，页面：建模） ----------
from FEM_MidasCivil.Steel_Pipe_Bailey_Support_FEA.Support_Midas_Main              import BaileySupport_FEM_MidasCivil_Main                 # 钢管贝雷梁现浇支架
from FEM_MidasCivil.Steel_Pipe_Bailey_Trestle_FEA.Trestle_UI_FEM                  import Trestle_FEM_MidasCivil_Main                       # 上承式桁架/型钢纵梁直线栈桥
from FEM_MidasCivil.Steel_Pipe_Bailey_Platform_FEA.Platform_main_FEM              import PlatForm_FEM_MidasCivil_Main                      # 钢管贝雷梁作业平台
from FEM_MidasCivil.Steel_Sheet_Pile_CofferDam_FEA.CofferDam_UI_FEM               import CofferDam_FEM_MidasCivil_Main                     # 钢板桩/锁扣钢管桩矩形围堰
from FEM_MidasCivil.Steel_Sheet_Pile_CofferDam_FEA_Batch.CofferDam_UI_FEM_Batch   import CofferDam_FEM_MidasCivil_Main_Batch               # 批量钢板桩/锁扣钢管桩矩形围堰

# ---------- 画图功能（AutoCAD，页面：画图） ----------
from Drawing_AutoCad.Steel_Pipe_Bailey_Support_CAD.C2M_command                    import BaileySupport_Drawing_AutoCad_Main                 # 钢管贝雷梁现浇支架
from Drawing_AutoCad.Steel_Pipe_Bailey_Trestle_CAD.Trestle_UI_CAD                 import Trestle_Drawing_AutoCad_Main                       # 上承式桁架/型钢纵梁直线栈桥
from Drawing_AutoCad.Steel_Pipe_Bailey_Platform_CAD.Platform_main_CAD             import PlatForm_Drawing_AutoCad_Main                      # 钢管贝雷梁作业平台
from Drawing_AutoCad.Steel_Sheet_Pile_CofferDam_CAD.CofferDam_UI_CAD              import CofferDam_Drawing_AutoCad_Main                     # 钢板桩/锁扣钢管桩矩形围堰

# ---------- 计算书功能（Word/Excel，页面：计算书） ----------
from CalRpt_MSWord.Steel_Pipe_Bailey_Support_Cal.Support_Cal_Report_UI            import BaileySupport_Cal_Report_Main                      # 钢管贝雷梁现浇支架
from CalRpt_MSWord.Steel_Pipe_Bailey_Trestle_Cal.Trestle_UI_Cal                   import Trestle_Cal_Report_Main                            # 上承式桁架/型钢纵梁直线栈桥
# 钢管贝雷梁作业平台计算书：暂无实现
from CalRpt_MSWord.Steel_Sheet_Pile_CofferDam_Cal.CofferDam_UI_Word               import CofferDam_Cal_Report_Main                          # 钢板桩/锁扣钢管桩矩形围堰
from CalRpt_MSWord.Steel_Sheet_Pile_CofferDam_Excel_Batch.CofferDam_UI_Excel_Batch import CofferDam_Cal_Report_Main_Batch                  # 批量钢板桩/锁扣钢管桩矩形围堰(Excel)


def run(resource_dir=None):
    """PyStand 启动入口"""
    # 读取注册表路径
    if resource_dir:
        parent_path = resource_dir
    else:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"SOFTWARE\ShuZhiQiaoShi", 0, winreg.KEY_READ) as key:
            parent_path, value_type = winreg.QueryValueEx(key, 'Applocation')
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"SOFTWARE\ShuZhiQiaoShi", 0, winreg.KEY_READ) as key:
            try:
                Revit_Version, value_type = winreg.QueryValueEx(key, 'Revit_Version')
            except:
                print('未检索到Revit版本号')
                Revit_Version = None
    except:
        Revit_Version = None

    # ================== 按功能分组的功能字典（结构与UI内容一致） ==================
    Midas_Func_Dict = {
        # 建模：结构 → Midas 建模函数
        '钢管贝雷梁现浇支架': BaileySupport_FEM_MidasCivil_Main,
        '上承式桁架/型钢纵梁直线栈桥': Trestle_FEM_MidasCivil_Main,
        '钢管贝雷梁作业平台': PlatForm_FEM_MidasCivil_Main,
        '钢板桩/锁扣钢管桩矩形围堰': CofferDam_FEM_MidasCivil_Main,
        '批量钢板桩/锁扣钢管桩矩形围堰': CofferDam_FEM_MidasCivil_Main_Batch,
    }
    CAD_Func_Dict = {
        # 画图：结构 → AutoCAD 绘图函数
        '钢管贝雷梁现浇支架': BaileySupport_Drawing_AutoCad_Main,
        '上承式桁架/型钢纵梁直线栈桥': Trestle_Drawing_AutoCad_Main,
        '钢管贝雷梁作业平台': PlatForm_Drawing_AutoCad_Main,
        '钢板桩/锁扣钢管桩矩形围堰': CofferDam_Drawing_AutoCad_Main,
    }
    Word_Func_Dict = {
        # 计算书：结构 → 计算书生成函数
        '钢管贝雷梁现浇支架': BaileySupport_Cal_Report_Main,
        '上承式桁架/型钢纵梁直线栈桥': Trestle_Cal_Report_Main,
        '钢管贝雷梁作业平台': None,  # 暂未实现
        '钢板桩/锁扣钢管桩矩形围堰': CofferDam_Cal_Report_Main,
        '钢板桩/锁扣钢管桩矩形围堰批量结果': CofferDam_Cal_Report_Main_Batch,  # Excel 批量结果
    }
    print('test: one message for git test')
    print('test: 2nd message for git test')
    print('test: 3rd message for git test')
<<<<<<< HEAD
    print('test: main branch commit upon once')
=======
    print('test: hotfix message for branch:hotfix')
>>>>>>> hotfix

    # 使用 ttkbootstrap UI（功能为页面、结构为入口）
    app = FloatingBallApp_ttk(
        parent_path, Revit_Version,
        Midas_Func_Dict, CAD_Func_Dict, Word_Func_Dict
    )


if __name__ == "__main__":
    run()
