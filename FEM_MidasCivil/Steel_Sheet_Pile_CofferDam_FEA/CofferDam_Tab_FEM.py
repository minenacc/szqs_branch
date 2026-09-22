# =============================================================================
# 术语表（全文件统一约定）
#   SKGGZ  : 锁扣钢管桩
#   Waler  : 围檩；Strut: 内支撑（对撑 DC / 斜撑 XC）
#   Cap    : 承台；Corbel/Bracket: 牛腿
#   sheet_refs : 跨 Tab 共享的引用字典，承载各 Tab 的 cache 与回调
#   cache  : 各 Tab 数据缓存（basic_cache/substructure_cache/boundary_cache/
#            surcharge_cache/condition_cache），供 UI.combine_cofferdam_dict 读取
#   工况类型 : 取土/抽水/加撑/拆撑/封底/垫层/辅助加撑/辅助换撑/辅助加圈梁
#   单位约定 : 界面输入长度/标高 m，截面参数 mm；土层表按列单位（见 Excel_io）
# =============================================================================

# 1. 标准库
import sys
import re
import tkinter as tk
from pathlib import Path
from tksheet import Sheet
from tkinter import messagebox

# 2. 第三方库
import ttkbootstrap as tb
from ttkbootstrap.constants import *

# 3. 本地模块
sys.path.append(str(Path(__file__).parent.parent.parent))
from General.DataUtils      import format_value
from General.UIHandle       import createToolTip, enable_button, SoilSettingsFrame, PileBoundaryFrame, ConnectionSettingFrame

sys.path.append(str(Path(__file__).parent))
from FEM_MidasCivil.Steel_Sheet_Pile_CofferDam_FEA.CofferDam_Generate_FEM  import Cal_Stage
from FEM_MidasCivil.Steel_Sheet_Pile_CofferDam_FEA.CofferDam_Calculate_FEM import calc_skggz_center_spacing
from FEM_MidasCivil.Steel_Sheet_Pile_CofferDam_FEA.CofferDam_Diagram_FEM   import SKGGZCrossSectionDiagram, SurchargeDiagram
from FEM_MidasCivil.Steel_Sheet_Pile_CofferDam_FEA.CofferDam_Excel_io_FEM  import get_excel_path, get_soil_sheet_names, load_soil_sheet

# 4. 共享工具函数（从主 UI 导入，均在模块级定义，无循环依赖）
from FEM_MidasCivil.Steel_Sheet_Pile_CofferDam_FEA.CofferDam_UI_FEM import make_card_frame, toggle_control, auto_calculate_pile_offsets


# ===== 主题配色 =====
TEXT_PRIMARY = "#2c3e50"
TEXT_SECONDARY = "#6b7280"
BORDER_COLOR = "#e5e7eb"

# 基本规定——规范名称（统一使用全称，兼容项目数据中存储的简称）
CODE_STANDARD_FULL = "《建筑基坑支护技术规程》（JGJ 120-2012）"
CODE_STANDARD_SHORT = "建筑基坑支护技术规程"

# ===================================================================
# 页面1+2：基本设置 + 支护桩布置
# ===================================================================

def _resolve_steel_std(std_val, steel_dict):
    """将主表"钢结构规范"值匹配到钢材材料库的 key。

    先精确匹配，再尝试提取书名号内名称，最后忽略书名号与空白后比较。

    Args:
        std_val (str): 主表中的钢结构规范名称（可能为简称或全称）。
        steel_dict (dict): 钢材材料库 {规范名: {...}}。

    Returns:
        str|None: 匹配到的规范 key；无匹配返回 None。
    """
    if not steel_dict or not std_val:
        return None
    std_val = str(std_val).strip()
    if std_val in steel_dict:
        return std_val
    m = re.search(r'《([^》]+)》', std_val)
    if m and m.group(1).strip() in steel_dict:
        return m.group(1).strip()
    norm = re.sub(r'[《》\s]', '', std_val)
    for k in steel_dict:
        if re.sub(r'[《》\s]', '', k) == norm:
            return k
    return None

def _resolve_concrete_std(std_val, concrete_dict):
    """将主表"混凝土规范"值匹配到混凝土材料库的 key。

    先精确匹配，再尝试提取书名号内名称，最后忽略书名号与空白后比较。

    Args:
        std_val (str): 主表中的混凝土规范名称。
        concrete_dict (dict): 混凝土材料库 {规范名: {...}}。

    Returns:
        str|None: 匹配到的规范 key；无匹配返回 None。
    """
    if not concrete_dict or not std_val:
        return None
    std_val = str(std_val).strip()
    if std_val in concrete_dict:
        return std_val
    m = re.search(r'《([^》]+)》', std_val)
    if m and m.group(1).strip() in concrete_dict:
        return m.group(1).strip()
    norm = re.sub(r'[《》\s]', '', std_val)
    for k in concrete_dict:
        if re.sub(r'[《》\s]', '', k) == norm:
            return k
    return None

def Basic_Setting_GUI(table, setup_path, Solid_Level, Water_Level, Cap_X, Cap_Y,
                         X_offset, Y_offset, Cap_Bottom_Level, Cap_H, widget_map=None,
                         CofferDam_Top_Level=None, CofferDam_L=None, Sheet_Pile_namevar=None,
                         Sheet_Pile_namelst=None, Pile_SEC_info_dict=None, Pile_sec_values=None,
                         root=None, Sheet_Pile_typevar=None, Sheet_Pile_typelst=None,
                         SKGGZ_D=None, SKGGZ_t=None, SKGGZ_Gap=None, applocation_sec=None,
                         sheet_refs=None, diagram=None, sections=None, mode="basic",
                         SKGGZ_LockWidth=None, SKGGZ_SheetCount=None, SKGGZ_SheetWidth=None,
                         steel_dict=None, projects_dict=None, current_proj_key_ref=None):
    """构建"基本设置 + 支护桩布置"页（Tab 1/2），并注册控件与缓存刷新。

    含基本规定、环境参数、承台参数、支护桩类型/截面/材质、围堰尺寸、
    边距与内边尺寸自动计算，以及钢板桩/锁扣钢管桩详细设置二级窗口。

    Args:
        table (tk.Widget): 页面容器。
        setup_path (str): 程序安装目录。
        Solid_Level, Water_Level (tk.Variable): 土顶/水面标高（m）。
        Cap_X, Cap_Y, X_offset, Y_offset (tk.Variable): 承台长宽与预留边距（m）。
        Cap_Bottom_Level, Cap_H (tk.Variable): 承台底标高/高（m）。
        widget_map (dict|None): 控件注册表（高亮联动用），就地写入。
        CofferDam_Top_Level, CofferDam_L (tk.Variable): 围堰顶标高/高（m）。
        Sheet_Pile_namevar, Sheet_Pile_namelst: 支护桩截面名变量与候选列表。
        Pile_SEC_info_dict (dict): 支护桩截面信息 {'SEC': [...]}。
        Pile_sec_values (dict): 内置截面参数字典。
        root (tk.Widget): 根窗口（二级窗口父级）。
        Sheet_Pile_typevar, Sheet_Pile_typelst: 支护类型变量与候选列表。
        SKGGZ_D, SKGGZ_t, SKGGZ_Gap (tk.Variable): 钢管桩直径/壁厚/中心间距（mm）。
        applocation_sec (str): 安装目录（截面资源用）。
        sheet_refs (dict|None): 跨 Tab 共享引用字典。
        diagram (RealTimeDiagram|None): 实时示意图实例。
        sections (dict|None): 截面库 {名称: {'type':..., 'params':...}}。
        mode (str): 'basic'/'sheet_pile'/'all'，控制渲染范围。
        SKGGZ_LockWidth, SKGGZ_SheetCount, SKGGZ_SheetWidth (tk.Variable): 锁扣参数。
        steel_dict (dict|None): 钢材材料库。
        projects_dict (dict|None): 全部项目字典。
        current_proj_key_ref (list|None): 当前项目键的引用（可变单元素列表）。

    Returns:
        None: 副作用为构建 UI 并写入 widget_map/sheet_refs。
    """
    if widget_map is None:
        widget_map = {}
    # ===== 基本设置（仅 mode=basic/all 时渲染） =====
    if mode in ("basic", "all"):
        # 基本规定frame（最上方）
        _, regulation_card = make_card_frame(table, "基本规定")
        regulation_row = tb.Frame(regulation_card); regulation_row.pack(fill=X, pady=5)
        code_standard_var = tb.StringVar(value=CODE_STANDARD_FULL)
        safety_grade_var = tb.StringVar(value="二级")
        tb.Label(regulation_row, text="规范:", width=9, bootstyle=PRIMARY).pack(side=LEFT, padx=10)
        code_combo = tb.Combobox(regulation_row, textvariable=code_standard_var,
                                  values=[CODE_STANDARD_FULL], state="readonly", width=35)
        code_combo.pack(side=LEFT, padx=5)
        safety_combo = tb.Combobox(regulation_row, textvariable=safety_grade_var,
                                    values=["一级", "二级", "三级"], state="readonly", width=14)
        safety_combo.pack(side=RIGHT, padx=10)

        safety_label = tb.Label(regulation_row, text="围堰安全等级:", width=16, bootstyle=PRIMARY)
        safety_label.pack(side=RIGHT, padx=10)
        def _on_code_change(*_):
            """规范切换时显隐围堰安全等级控件（仅当前规范支持时显示）。"""
            if CODE_STANDARD_SHORT in code_standard_var.get():
                safety_combo.pack(side=RIGHT, padx=10)
                safety_label.pack(side=RIGHT, padx=10)
            else:
                safety_label.pack_forget()
                safety_combo.pack_forget()
        code_standard_var.trace_add("write", _on_code_change)
        widget_map["code_standard"] = code_standard_var
        widget_map["safety_grade"] = safety_grade_var

        _, card = make_card_frame(table, "基本环境参数")
        container1 = tb.Frame(card)
        container1.pack(fill=X, pady=5)
        label_ground_level = tb.Label(master=container1, text="土顶标高(m): ", width=16, bootstyle=PRIMARY)
        label_ground_level.pack(side=LEFT, padx=10, pady=5)
        entry1 = tb.Entry(master=container1, textvariable=Solid_Level, width=15)
        entry1.pack(side=LEFT, padx=10, pady=5)
        entry2 = tb.Entry(master=container1, textvariable=Water_Level, width=15)
        entry2.pack(side=RIGHT, padx=10, pady=5)
        label_water_level = tb.Label(master=container1, text="水面标高(m): ", width=16, bootstyle=PRIMARY)
        label_water_level.pack(side=RIGHT, padx=10, pady=5)
        widget_map["ground_level"] = entry1
        widget_map["water_level"] = entry2

        _, card2 = make_card_frame(table, "承台基本参数")
        container2 = tb.Frame(card2)
        container2.pack(fill=X, pady=5)
        label_cap_length = tb.Label(master=container2, text="承台长(m): ", width=16, bootstyle=PRIMARY)
        label_cap_length.pack(side=LEFT, padx=10, pady=5)
        entry3 = tb.Entry(master=container2, textvariable=Cap_X, width=15)
        entry3.pack(side=LEFT, padx=10, pady=5)
        entry4 = tb.Entry(master=container2, textvariable=Cap_Y, width=15)
        entry4.pack(side=RIGHT, padx=10, pady=5)
        label_cap_width = tb.Label(master=container2, text="承台宽(m): ", width=16, bootstyle=PRIMARY)
        label_cap_width.pack(side=RIGHT, padx=10, pady=5)
        container3 = tb.Frame(card2)
        container3.pack(fill=X, pady=5)
        label_cap_bottom = tb.Label(master=container3, text="承台底标高(m): ", width=16, bootstyle=PRIMARY)
        label_cap_bottom.pack(side=LEFT, padx=10, pady=5)
        entry5 = tb.Entry(master=container3, textvariable=Cap_Bottom_Level, width=15)
        entry5.pack(side=LEFT, padx=10, pady=5)
        entry6 = tb.Entry(master=container3, textvariable=Cap_H, width=15)
        entry6.pack(side=RIGHT, padx=10, pady=5)
        label_cap_height = tb.Label(master=container3, text="承台高度(m): ", width=16, bootstyle=PRIMARY)
        label_cap_height.pack(side=RIGHT, padx=10, pady=5)
        widget_map["cap_x"] = entry3
        widget_map["cap_y"] = entry4
        widget_map["cap_bottom"] = entry5
        widget_map["cap_h"] = entry6
        widget_map["cap_bottom"] = entry5

        # 承台顶标高 = 承台底标高 + 承台高（自动计算，只读）
        _cap_top_var = tb.StringVar(value="0.000")
        def _update_cap_top(*_):
            """根据承台底标高与承台高自动刷新只读的承台顶标高。"""
            try:
                cb = float(Cap_Bottom_Level.get())
                ch = float(Cap_H.get())
                _cap_top_var.set(f"{cb + ch:.3f}")
            except: _cap_top_var.set("0.000")
        Cap_Bottom_Level.trace_add("write", _update_cap_top)
        Cap_H.trace_add("write", _update_cap_top)
        container4 = tb.Frame(card2)
        container4.pack(fill=X, pady=5)
        tb.Label(container4, text="承台顶标高(m): ", width=16, bootstyle=PRIMARY).pack(side=LEFT, padx=10, pady=5)
        tb.Entry(container4, textvariable=_cap_top_var, width=15, state=tk.DISABLED).pack(side=LEFT, padx=10, pady=5)
        _update_cap_top()  # 初始计算

        # 承台边距自动计算触发（移到 pile_param 中）

    # ===== 支护桩布置 =====
    if mode in ("sheet_pile", "all"):
    # ===== 支护桩布置参数 =====
        if widget_map is None:
            widget_map = {}

        # 从 sections 构建下拉选项池
        gbz_names = ["自定义"]
        skggz_names = ["自定义"]
        if sections:
            for sn, sv in sections.items():
                if sv["type"] == "钢板桩":
                    gbz_names.append(sn)
                elif sv["type"] == "锁扣钢管桩":
                    skggz_names.append(sn)
        if sheet_refs is not None:
            sheet_refs["_gbz_names"] = gbz_names
            sheet_refs["_skggz_names"] = skggz_names
            sheet_refs["_sections"] = sections or {}
        
        _, pile_param = make_card_frame(table, "围堰基本参数")

        # 第一行：截面类型选择与设置
        type_row = tb.Frame(pile_param)
        type_row.pack(fill=X, pady=5)
    
        tb.Label(type_row, text="支护桩类型:", bootstyle=PRIMARY, width=16).pack(side=LEFT, padx=10, pady=5)
        combo_type = tb.Combobox(type_row, textvariable=Sheet_Pile_typevar, values=Sheet_Pile_typelst, width=14, state="readonly")
        combo_type.pack(side=LEFT, padx=10)
        # 切换至锁扣钢管桩时自动加载第一个标准型号
        def _on_type_to_skggz(*_):
            """支护类型切换为锁扣钢管桩时，自动加载首个标准型号参数。"""
            if Sheet_Pile_typevar.get() == "锁扣钢管桩" and sections:
                for name, sec in sections.items():
                    if sec.get("type") == "锁扣钢管桩":
                        p = sec.get("params", {})
                        if p.get("D"):
                            SKGGZ_D.set(p["D"])
                            SKGGZ_t.set(p.get("t", 0))
                            SKGGZ_LockWidth.set(p.get("锁扣宽度", 35))
                            SKGGZ_SheetCount.set(p.get("钢板桩数量", 1))
                            SKGGZ_SheetWidth.set(p.get("钢板桩宽度", 600))
                            SKGGZ_Gap.set(calc_skggz_center_spacing(
                                p["D"], p.get("锁扣宽度", 35),
                                p.get("钢板桩数量", 1), p.get("钢板桩宽度", 600)))
                        break
        Sheet_Pile_typevar.trace_add("write", _on_type_to_skggz)

        sec_setup_btn = tb.Button(type_row, text="详细截面设置", bootstyle=PRIMARY, width=16)
        sec_setup_btn.pack(side=LEFT, padx=10)

        # 围堰高
        entry2 = tb.Entry(master=type_row, textvariable=CofferDam_L, width=15)
        entry2.pack(side=RIGHT, padx=10, pady=5)
        label_pile_length = tb.Label(master=type_row, text="围堰高(m): ", width=16, bootstyle=PRIMARY)
        label_pile_length.pack(side=RIGHT, padx=10, pady=5)

        # 围堰顶标高 / 围堰底标高
        container1 = tb.Frame(pile_param)
        container1.pack(fill=X, pady=5)
        label_pile_top = tb.Label(master=container1, text="围堰顶标高(m): ", width=16, bootstyle=PRIMARY)
        label_pile_top.pack(side=LEFT, padx=10, pady=5)
        entry1 = tb.Entry(master=container1, textvariable=CofferDam_Top_Level, width=15)
        entry1.pack(side=LEFT, padx=10, pady=5)
        # 围堰底标高 = 围堰顶标高 - 围堰高（自动计算，只读）
        _pit_bottom_var = tb.StringVar(value="0.000")
        def _update_pit_bottom(*_):
            """根据围堰顶标高与围堰高自动刷新只读的围堰底标高。"""
            try:
                pt = float(CofferDam_Top_Level.get())
                pl = float(CofferDam_L.get())
                _pit_bottom_var.set(f"{pt - pl:.3f}")
            except: _pit_bottom_var.set("0.000")
        CofferDam_Top_Level.trace_add("write", _update_pit_bottom)
        CofferDam_L.trace_add("write", _update_pit_bottom)
        _pit_entry = tb.Entry(master=container1, textvariable=_pit_bottom_var, width=15, state=tk.DISABLED)
        _pit_entry.pack(side=RIGHT, padx=10, pady=5)
        tb.Label(master=container1, text="围堰底标高(m): ", width=16, bootstyle=PRIMARY).pack(side=RIGHT, padx=10, pady=5)
        _update_pit_bottom()
        widget_map["pile_top"] = entry1
        widget_map["pile_length"] = entry2

        # 边距变化时更新内边长/边宽（定义提前，供下方回调用）
        _inner_length_var = tb.StringVar(value="0.000")
        _inner_width_var = tb.StringVar(value="0.000")
        def _update_inner_dims(*_):
            """先自动计算边距，再刷新只读的围堰内边长/边宽。"""
            # 先自动计算边距，再算内边长/边宽
            auto_calculate_pile_offsets(Cap_X, Cap_Y, X_offset, Y_offset,
                                         Sheet_Pile_typevar, Pile_SEC_info_dict,
                                         SKGGZ_D, SKGGZ_Gap)
            try:
                cx = float(Cap_X.get())
                cy = float(Cap_Y.get())
                xo = float(X_offset.get())
                yo = float(Y_offset.get())
                _inner_length_var.set(f"{cx + 2 * xo:.3f}")
                _inner_width_var.set(f"{cy + 2 * yo:.3f}")
            except (ValueError, tk.TclError):
                _inner_length_var.set("0.000")
                _inner_width_var.set("0.000")

        # 承台长边边距 / 承台短边边距（同一行，左右对齐）
        offset_row = tb.Frame(pile_param)
        offset_row.pack(fill=X, pady=5)
        tb.Label(offset_row, text="承台长边边距(m): ", bootstyle=PRIMARY, width=16).pack(side=LEFT, padx=10, pady=5)
        tb.Entry(offset_row, textvariable=X_offset, width=15).pack(side=LEFT, padx=10, pady=5)
        tb.Entry(offset_row, textvariable=Y_offset, width=15).pack(side=RIGHT, padx=10, pady=5)
        tb.Label(offset_row, text="承台短边边距(m): ", bootstyle=PRIMARY, width=16).pack(side=RIGHT, padx=10, pady=5)
        # 围堰内边长 / 围堰内边宽（同一行，左右对齐，只读）
        inner_row = tb.Frame(pile_param)
        inner_row.pack(fill=X, pady=5)
        tb.Label(inner_row, text="围堰内边长(m): ", bootstyle=PRIMARY, width=16).pack(side=LEFT, padx=10, pady=5)
        tb.Entry(inner_row, textvariable=_inner_length_var, width=15, state=tk.DISABLED).pack(side=LEFT, padx=10, pady=5)
        tb.Entry(inner_row, textvariable=_inner_width_var, width=15, state=tk.DISABLED).pack(side=RIGHT, padx=10, pady=5)
        tb.Label(inner_row, text="围堰内边宽(m): ", bootstyle=PRIMARY, width=16).pack(side=RIGHT, padx=10, pady=5)
        # 自动计算边距 + 内边长/边宽
        def _recalc_with_offsets(*_):
            """自动重算承台边距后再刷新内边尺寸（承台/桩型变化时触发）。"""
            auto_calculate_pile_offsets(Cap_X, Cap_Y, X_offset, Y_offset,
                                         Sheet_Pile_typevar, Pile_SEC_info_dict,
                                         SKGGZ_D, SKGGZ_Gap)
            _update_inner_dims()
        def _recalc_dims_only(*_):
            """仅刷新内边尺寸（边距被手动修改时触发）。"""
            _update_inner_dims()
        Cap_X.trace_add("write", _recalc_with_offsets)
        Cap_Y.trace_add("write", _recalc_with_offsets)
        if Sheet_Pile_typevar:
            try: Sheet_Pile_typevar.trace_add("write", _recalc_with_offsets)
            except: pass
        X_offset.trace_add("write", _recalc_dims_only)
        Y_offset.trace_add("write", _recalc_dims_only)
        _recalc_with_offsets()
        if sheet_refs is not None:
            sheet_refs["_inner_length_var"] = _inner_length_var
            sheet_refs["_inner_width_var"] = _inner_width_var

        hint = tb.Label(table, text="* 边距为承台边线至桩内侧的距离，根据承台尺寸及支护桩设置自动计算",
                         font=("Microsoft YaHei UI", 8), foreground=TEXT_SECONDARY)
        hint.pack(anchor=W, padx=20, pady=(5, 5))

        # ==========================================
        # 数据字典与变量定义 (供二级窗口使用)
        # ==========================================
        _EXCEL_KEY_MAP = {
            'H(mm)': 'H', 'B(mm)': 'B', 'As(mm^2)': 'A',
            'Asy(mm^2)': 'Asy', 'Asz(mm^2)': 'Asz',
            'Ixx(mm^4)': 'Ixx', 'Iyy(mm^4)': 'Iyy', 'Izz(mm^4)': 'Izz',
            'Cyp(mm)': 'Cyp', 'Cym(mm)': 'Cym', 'Czp(mm)': 'Czp', 'Czm(mm)': 'Czm',
            'Qyb(mm^2)': 'Qyb', 'Qzb(mm^2)': 'Qzb',
            'Peri:O(mm)': 'PeriO', 'Peri:I(mm)': 'PeriI',
            'Cent:y(mm)': 'Centy', 'Cent:z(mm)': 'Centz',
            'y1(mm)': 'y1', 'z1(mm)': 'z1', 'y2(mm)': 'y2', 'z2(mm)': 'z2',
            'y3(mm)': 'y3', 'z3(mm)': 'z3', 'y4(mm)': 'y4', 'z4(mm)': 'z4',
            'Zyy(mm^3)': 'Zyy', 'Zzz(mm^3)': 'Zzz',
        }
        variable_names = ['H', 'B', 'A', 'Asy', 'Asz', 'Ixx', 'Iyy', 'Izz',
                          'Cyp', 'Cym', 'Czp', 'Czm', 'Qyb', 'Qzb',
                          'PeriO', 'PeriI', 'Centy', 'Centz',
                          'y1', 'z1', 'y2', 'z2', 'y3', 'z3', 'y4', 'z4', 'Zyy', 'Zzz']
        units = ["mm", "mm", "mm²", "mm²", "mm²", "mm⁴", "mm⁴", "mm⁴",
                 "mm", "mm", "mm", "mm", "mm²", "mm²",
                 "mm", "mm", "mm", "mm",
                 "mm", "mm", "mm", "mm", "mm", "mm", "mm", "mm", "mm³", "mm³"]

        # ==========================================
        # 二级窗口：钢板桩详细设置
        # ==========================================
        def _open_gbz_dialog():
            """打开钢板桩详细截面设置二级窗口（选型/材质/28 项截面特性）。"""
            dialog_window = tb.Toplevel(root)
            dialog_window.title("钢板桩截面设置")
            dialog_window.geometry("760x550+%d+%d" % (root.winfo_rootx()+80, root.winfo_rooty()+80))
            dialog_window.transient(root); dialog_window.grab_set()
            dialog_window.resizable(False, False)

            SEC_LABEL_W = 15
            SEC_INPUT_W = 15
            MAT_LABEL_W = 15
            MAT_INPUT_W = 8

            # 底部按钮
            bottom_button_frame = tb.Frame(dialog_window)
            bottom_button_frame.pack(side=BOTTOM, fill=X, padx=15, pady=10)

            # 顶部配置
            top_config_frame = tb.Frame(dialog_window)
            top_config_frame.pack(side=TOP, fill=X, padx=15, pady=(10, 5))

            model_select_row = tb.Frame(top_config_frame); model_select_row.pack(fill=X, pady=3)
            tb.Label(model_select_row, text="选取钢板桩型号:", width=SEC_LABEL_W).pack(side=LEFT, padx=5)

            default_gbz = sheet_refs.get("_last_gbz_section", "")
            if not default_gbz or default_gbz not in sheet_refs.get("_gbz_names", []):
                gbz_opts_local = [x for x in sheet_refs.get("_gbz_names", ["自定义"]) if x != "自定义"]
                default_gbz = gbz_opts_local[0] if gbz_opts_local else "自定义"
            selected_model_var = tb.StringVar(value=default_gbz)
            model_combobox = tb.Combobox(model_select_row, textvariable=selected_model_var, width=SEC_INPUT_W, state="readonly")
            model_combobox.pack(side=LEFT, padx=5)
            model_combobox.configure(values=sheet_refs.get("_gbz_names", ["自定义"]))
            material_var = tb.StringVar(value='Q295')
            material_combobox = tb.Combobox(model_select_row, textvariable=material_var, width=MAT_INPUT_W, state="readonly")
            material_combobox.pack(side=RIGHT, padx=10)
            material_combobox.configure(values=['Q295'])
            material_combobox._var = material_var
            material_combobox.set('Q295')
            sheet_refs["_gbz_material_var"] = material_var
            tb.Label(model_select_row, text="选取钢板桩材质:", width=MAT_LABEL_W).pack(side=RIGHT, padx=5)

            custom_name_row = tb.Frame(top_config_frame)
            tb.Label(custom_name_row, text="截面名称:", width=SEC_LABEL_W).pack(side=LEFT, padx=5)
            custom_name_var = tb.StringVar()
            custom_name_entry = tb.Entry(custom_name_row, textvariable=custom_name_var, width=SEC_INPUT_W)
            custom_name_entry.pack(side=LEFT, padx=5)
            def toggle_custom_name_input(*_):
                """选中"自定义"时显示截面名称输入行，否则隐藏。"""
                if selected_model_var.get() == "自定义":
                    custom_name_row.pack(fill=X, pady=3, after=model_select_row)
                else:
                    custom_name_row.pack_forget()
            selected_model_var.trace_add("write", toggle_custom_name_input)

            # 表格部分
            from tksheet import Sheet
            half_property_count = len(variable_names) // 2
        
            sheet_border = tk.Frame(dialog_window, bg="#E2E8F0")
            sheet_border.pack(fill=BOTH, expand=True, padx=15, pady=5)
            sheet_border.pack_propagate(False)
            sheet_inner = tk.Frame(sheet_border, bg="white")
            sheet_inner.pack(fill=BOTH, expand=True, padx=1, pady=1)
            property_sheet = Sheet(
                sheet_inner,
                font=("Microsoft YaHei UI", 9, "normal"),
                header_font=("Microsoft YaHei UI", 9, "bold"),
                show_row_index=False, show_column_index=False, show_top_left=False,
                editable=True, headers=False, row_height=30, table_bg="white",
                header_bg="white", header_fg="#0F172A", index_bg="white", top_left_bg="white",
                empty_horizontal=0, empty_vertical=0,
                selected_row_bg="#FFF7ED", selected_cell_bg="#FFF7ED",
            )
            try: property_sheet.hide("header")
            except: pass
        
            property_sheet.set_column_widths([60, 90, 40, 60, 90, 40])
            property_sheet.pack(fill=BOTH, expand=True)
            # 激活原有交互事件
            property_sheet.enable_bindings()

            # 加载与权限控制
            def load_section_properties(selected_option):
                """把选中型号的截面特性加载到表格，并设置自定义/标准型的可编辑权限。

                Args:
                    selected_option (str): 选中的截面名称，"自定义"表示手工输入。

                Returns:
                    None
                """
                property_values = []
                is_custom = (selected_option == "自定义")
                # 清空旧的锁定状态，再重新设置
                try: property_sheet.readonly_columns([0, 1, 2, 3, 4, 5], readonly=False)
                except: pass

                if is_custom:
                    try: property_sheet.readonly_columns([0, 2, 3, 5])
                    except: pass
                else:
                    try: property_sheet.readonly_columns([0, 1, 2, 3, 4, 5])
                    except: pass
                    # 读取数据
                    if sections and selected_option in sections:
                        section_params = sections[selected_option].get("params", {})
                        property_values = [section_params.get(ek, 0) for ek in _EXCEL_KEY_MAP]
                    else:
                        property_values = list(Pile_sec_values.get(selected_option, [0]*28))

                sheet_display_data = []
                for i in range(half_property_count):
                    j = i + half_property_count
                    val1 = "" if is_custom else format_value(property_values[i])
                    val2 = "" if is_custom else format_value(property_values[j])
                    sheet_display_data.append([variable_names[i], val1, units[i], variable_names[j], val2, units[j]])

                property_sheet.set_sheet_data(sheet_display_data)
            
                # 原版底色高亮
                for row_idx in range(len(sheet_display_data)):
                    property_sheet.highlight_cells(row_idx, 0, bg="#DBEAFE", fg="#0F172A")
                    property_sheet.highlight_cells(row_idx, 3, bg="#DBEAFE", fg="#0F172A")
                    property_sheet.highlight_cells(row_idx, 2, fg="#94A3B8")
                    property_sheet.highlight_cells(row_idx, 5, fg="#94A3B8")

            model_combobox.bind("<<ComboboxSelected>>", lambda e: load_section_properties(model_combobox.get()))
            load_section_properties(selected_model_var.get())

            # ==========================================
            # 保存逻辑
            # ==========================================
            def save_section_settings():
                """保存钢板桩截面设置：新建/更新截面库并同步变量、边距与图示。"""
                current_model_name = selected_model_var.get()
                current_material_name = material_var.get()

                if current_model_name == "自定义":
                    current_model_name = custom_name_var.get().strip()
                    if not current_model_name: 
                        messagebox.showwarning("保存失败", "截面名称不能为空", parent=dialog_window); return
                    if sections and current_model_name in sections: 
                        messagebox.showwarning("保存失败", f"截面 '{current_model_name}' 已存在", parent=dialog_window); return
                
                    current_sheet_data = property_sheet.get_sheet_data()
                    extracted_values = []
                    for row in current_sheet_data:
                        for col_idx in (1, 4):
                            try: extracted_values.append(float(row[col_idx]))
                            except: extracted_values.append(0.0)
                        
                    if not any(extracted_values): 
                        messagebox.showwarning("保存失败", "截面特性值未填写完整", parent=dialog_window); return
                    
                    if sections is not None:
                        sections[current_model_name] = {
                            "type": "钢板桩", 
                            "params": {ek: extracted_values[i] if i < len(extracted_values) else 0 for i, ek in enumerate(_EXCEL_KEY_MAP)}
                        }
                    updated_gbz_names = list(sheet_refs.get("_gbz_names", ["自定义"]))
                    updated_gbz_names.append(current_model_name)
                    sheet_refs["_gbz_names"] = updated_gbz_names
            
                sheet_refs["_last_gbz_section"] = current_model_name
                sheet_refs["_last_gbz_material"] = current_material_name
                if projects_dict and current_proj_key_ref and current_proj_key_ref[0] in projects_dict:
                    projects_dict[current_proj_key_ref[0]]["支护桩材质"] = current_material_name
                Sheet_Pile_namevar.set(current_model_name)
                if sections and current_model_name in sections:
                    section_params = sections[current_model_name].get("params", {})
                    Pile_SEC_info_dict["SEC"] = [section_params.get(ek, 0) for ek in _EXCEL_KEY_MAP]
                    # 截面参数变化 → 重算承台边距
                    try: _recalc_with_offsets()
                    except: pass
            
                dialog_window.destroy()
                if diagram: diagram.refresh()

            tb.Button(bottom_button_frame, text="保存", bootstyle=SUCCESS, width=10, command=save_section_settings).pack(side=RIGHT, padx=5)
            tb.Button(bottom_button_frame, text="取消", bootstyle=SECONDARY, width=10, command=dialog_window.destroy).pack(side=RIGHT, padx=5)

        # ==========================================
        # 二级窗口：锁扣钢管桩详细设置
        # ==========================================
        def _open_skggz_dialog():
            """打开锁扣钢管桩详细截面设置二级窗口（型号/材质/几何参数/截面俯视图）。"""
            dialog_window = tb.Toplevel(root)
            dialog_window.title("锁扣钢管桩截面设置")
            dialog_window.geometry("850x500+%d+%d" % (root.winfo_rootx()+80, root.winfo_rooty()+80))
            dialog_window.transient(root)
            dialog_window.grab_set()
            # dialog_window.resizable(False, False)

            LABEL_W = 18
            INPUT_W = 16

            # 底部按钮区域 (先 pack 到底部，避免后续内容挤压)
            bottom_button_frame = tb.Frame(dialog_window)
            bottom_button_frame.pack(side=BOTTOM, fill=X, padx=15, pady=10)

            # 主内容区域
            main_content_frame = tb.Frame(dialog_window)
            main_content_frame.pack(fill=BOTH, expand=True, padx=15, pady=10)
            
            # 左侧输入区
            left_input_frame = tb.Frame(main_content_frame)
            left_input_frame.pack(side=LEFT, fill=Y, padx=5)

            # 垂直分隔线 (Separator)
            # padx 提供左右间距，fill=Y 让线垂直拉满
            sep = tb.Separator(main_content_frame, orient=VERTICAL)
            sep.pack(side=LEFT, fill=Y, padx=15)

            # 右侧图示区
            right_diagram_frame = tb.Frame(main_content_frame)
            right_diagram_frame.pack(side=LEFT, fill=BOTH, expand=True, padx=5)

            # 1. 预先创建所有行
            model_select_row = tb.Frame(left_input_frame)
            model_select_row.pack(fill=X, pady=3)
            
            custom_name_row = tb.Frame(left_input_frame)
            # custom_name_row 先不 pack，后面逻辑控制

            # 2. 逻辑：处理型号选择
            skg_opts_local = sheet_refs.get("_skggz_names", ["自定义"])
            last_skg = sheet_refs.get("_last_skggz_section", "")
            default_skg_val = last_skg if last_skg in skg_opts_local else (skg_opts_local[0] if skg_opts_local else "自定义")
            
            selected_model_var = tb.StringVar(value=default_skg_val)
            tb.Label(model_select_row, text="钢管桩型号:", width=LABEL_W).pack(side=LEFT, padx=5)
            model_combobox = tb.Combobox(model_select_row, textvariable=selected_model_var, width=INPUT_W, state="readonly")
            model_combobox.pack(side=LEFT, padx=5)
            model_combobox.configure(values=skg_opts_local)

            # 3. 自定义名称逻辑
            tb.Label(custom_name_row, text="型号名称:", width=LABEL_W).pack(side=LEFT, padx=5)
            custom_name_var = tb.StringVar()
            custom_name_entry = tb.Entry(custom_name_row, textvariable=custom_name_var, width=INPUT_W)
            custom_name_entry.pack(side=LEFT, padx=5)
            
            def toggle_custom_name_input(*_):
                """选中"自定义"时显示型号名称输入行，否则隐藏。"""
                if selected_model_var.get() == "自定义":
                    custom_name_row.pack(fill=X, pady=3, after=model_select_row)
                else:
                    custom_name_row.pack_forget()
            selected_model_var.trace_add("write", toggle_custom_name_input)
            toggle_custom_name_input() # 初始化执行一次

            # 4. 几何参数 (变量初始化)
            diameter_var = tb.StringVar(value=str(SKGGZ_D.get()))
            thickness_var = tb.StringVar(value=str(SKGGZ_t.get()))
            lock_width_var = tb.StringVar(value=str(SKGGZ_LockWidth.get()))
            sheet_count_var = tb.StringVar(value=str(int(SKGGZ_SheetCount.get())))
            sheet_width_var = tb.StringVar(value=str(SKGGZ_SheetWidth.get()))
            gap_spacing_var = tb.StringVar(value=str(SKGGZ_Gap.get()))

            # 钢板桩数量限制为整数
            def _validate_sheet_count(new_val):
                """输入校验：钢板桩数量只允许整数（空串允许，便于编辑）。"""
                if new_val == "":
                    return True
                try:
                    int(new_val)
                    return True
                except ValueError:
                    return False
            _vcmd_sheet_count = dialog_window.register(_validate_sheet_count)

            # 核心计算逻辑
            def _recalc_center_spacing(*_):
                """根据 D/锁扣宽度/钢板桩数量与宽度重算并刷新钢管桩中心间距。"""
                try:
                    D = float(diameter_var.get() or "0")
                    lock = float(lock_width_var.get() or "0")
                    cnt = float(sheet_count_var.get() or "0")
                    w = float(sheet_width_var.get() or "0")
                    gap_spacing_var.set(str(calc_skggz_center_spacing(D, lock, cnt, w)))
                except: pass

            # 5. 生成输入项
            entry_widgets = []
            entry_configs = [
                ("钢管桩直径(mm):", diameter_var, True),
                ("钢管桩壁厚(mm):", thickness_var, True),
                ("锁扣宽度(mm):", lock_width_var, True),
                ("钢板桩数量:", sheet_count_var, True),
                ("钢板桩宽度(mm):", sheet_width_var, True),
                ("钢管桩中心间距(mm):", gap_spacing_var, False), 
            ]
            for idx, (label_text, target_var, editable) in enumerate(entry_configs):
                input_row = tb.Frame(left_input_frame)
                input_row.pack(fill=X, pady=3)
                tb.Label(input_row, text=label_text, width=LABEL_W).pack(side=LEFT, padx=5)
                if idx == 3:  # 钢板桩数量：限制整数输入
                    entry_widget = tb.Entry(input_row, textvariable=target_var, width=INPUT_W,
                                            validate="key", validatecommand=(_vcmd_sheet_count, "%P"))
                else:
                    entry_widget = tb.Entry(input_row, textvariable=target_var, width=INPUT_W, state="readonly" if not editable else "normal")
                entry_widget.pack(side=LEFT, padx=5)
                entry_widgets.append(entry_widget)
                if editable: target_var.trace_add("write", _recalc_center_spacing)

            # 6. 型号切换联动 (更新只读状态)
            def load_skggz_properties(*_):
                """切换型号时加载其几何参数，并按自定义/标准型设置输入框可编辑状态。"""
                is_custom = (selected_model_var.get() == "自定义")
                if not is_custom and sections and selected_model_var.get() in sections:
                    sp = sections[selected_model_var.get()].get("params", {})
                    diameter_var.set(sp.get("D", "")); thickness_var.set(sp.get("t", ""))
                    lock_width_var.set(sp.get("锁扣宽度", "35"))
                    sheet_count_var.set(str(int(float(sp.get("钢板桩数量", "1")))))
                    sheet_width_var.set(sp.get("钢板桩宽度", "600"))
                
                for i, ew in enumerate(entry_widgets):
                    if i < 5: ew.configure(state="normal" if is_custom else "readonly")
                _recalc_center_spacing()

            model_combobox.bind("<<ComboboxSelected>>", load_skggz_properties)
            load_skggz_properties()

            # 6.5 钢管桩材质（图片上方）
            _mat_options = []
            _default_mat = ""
            if projects_dict and current_proj_key_ref and current_proj_key_ref[0] in projects_dict:
                _cur_proj = projects_dict[current_proj_key_ref[0]]
                _mat_std = _resolve_steel_std(_cur_proj.get("钢结构规范", ""), steel_dict or {})
                if _mat_std and steel_dict and _mat_std in steel_dict:
                    _mat_options = list(steel_dict[_mat_std].get("牌号参数", {}).keys())
                _default_mat = str(_cur_proj.get("支护桩材质", "")).strip()
                if _default_mat not in _mat_options:
                    _default_mat = _mat_options[0] if _mat_options else ""
            mat_row = tb.Frame(right_diagram_frame)
            mat_row.pack(fill=X, pady=(0, 5))
            tb.Label(mat_row, text="钢管桩材质:", width=LABEL_W).pack(side=LEFT, padx=5)
            material_var = tb.StringVar(value=_default_mat)
            material_combobox = tb.Combobox(mat_row, textvariable=material_var, values=_mat_options,
                                            state="readonly", width=INPUT_W)
            material_combobox.pack(side=LEFT, padx=5)
            material_combobox._var = material_var
            material_combobox.set(_default_mat)
            sheet_refs["_skggz_material_var"] = material_var

            # 7. 右侧绘图
            skggz_diag = SKGGZCrossSectionDiagram(right_diagram_frame, diameter_var, thickness_var, gap_spacing_var)
            skggz_diag.draw()  # 初始化绘制

            # 8. 保存与取消
            def save_skggz_settings():
                """保存锁扣钢管桩截面设置：新建/更新截面库并同步变量、边距与图示。"""
                current_model_name = selected_model_var.get()
                if current_model_name == "自定义":
                    current_model_name = custom_name_var.get().strip()
                    if not current_model_name:
                        messagebox.showwarning("保存失败", "型号名称不能为空", parent=dialog_window); return
                    if sections and current_model_name in sections:
                        messagebox.showwarning("保存失败", f"型号 '{current_model_name}' 已存在", parent=dialog_window); return
                    if sections is not None:
                        sections[current_model_name] = {
                            "type": "锁扣钢管桩",
                            "params": {
                                "D": float(diameter_var.get() or 0),
                                "t": float(thickness_var.get() or 0),
                                "锁扣宽度": float(lock_width_var.get() or 35),
                                "钢板桩数量": int(float(sheet_count_var.get() or 1)),
                                "钢板桩宽度": float(sheet_width_var.get() or 600),
                            }
                        }
                    updated_skggz_names = list(sheet_refs.get("_skggz_names", ["自定义"]))
                    updated_skggz_names.append(current_model_name)
                    sheet_refs["_skggz_names"] = updated_skggz_names
                else:
                    # 已有截面：回写修改到 sections
                    if sections and current_model_name in sections:
                        sections[current_model_name]["params"] = {
                            "D": float(diameter_var.get() or 0),
                            "t": float(thickness_var.get() or 0),
                            "锁扣宽度": float(lock_width_var.get() or 35),
                            "钢板桩数量": int(float(sheet_count_var.get() or 1)),
                            "钢板桩宽度": float(sheet_width_var.get() or 600),
                        }
                SKGGZ_D.set(diameter_var.get())
                SKGGZ_t.set(thickness_var.get())
                SKGGZ_Gap.set(gap_spacing_var.get())
                Sheet_Pile_namevar.set(current_model_name)
                sheet_refs["_last_skggz_section"] = current_model_name
                sheet_refs["_last_skggz_material"] = material_var.get()
                if projects_dict and current_proj_key_ref and current_proj_key_ref[0] in projects_dict:
                    projects_dict[current_proj_key_ref[0]]["支护桩材质"] = material_var.get()
                # 截面参数变化 → 重算承台边距
                try: _recalc_with_offsets()
                except: pass
                dialog_window.destroy()
                if diagram: diagram.refresh()
            tb.Button(bottom_button_frame, text="保存", bootstyle=SUCCESS, command=save_skggz_settings).pack(side=RIGHT, padx=5)
            tb.Button(bottom_button_frame, text="取消", bootstyle=SECONDARY, command=dialog_window.destroy).pack(side=RIGHT, padx=5)

        # ==========================================
        # 事件绑定与接口兼容
        # ==========================================
        def _on_sec_setup(*_):
            """"详细截面设置"按钮回调：按支护类型打开对应二级窗口。"""
            if Sheet_Pile_typevar.get() == "钢板桩":
                _open_gbz_dialog()
            elif Sheet_Pile_typevar.get() == "锁扣钢管桩":
                _open_skggz_dialog()
            else:
                pass
            
        sec_setup_btn.config(command=_on_sec_setup)

        # 兼容外部项目切换时的图示刷新需求
        def _dummy_frame_change(*_):
            """项目切换/下拉变化时重算边距并刷新图示（兼容外部调用接口）。"""
            try: _recalc_with_offsets()
            except: pass
            if diagram is not None:
                diagram.refresh()
            
        if sheet_refs is not None:
            # 保留兼容字典键值，防止 CofferDam_UI.py 中的 _switch_project 报错
            sheet_refs["_on_section_frame_change"] = _dummy_frame_change

    # ---- basic cache ----
    def _refresh_basic_cache():
        """把基本设置页当前值写入 sheet_refs['basic_cache']，供保存时读取。"""
        if sheet_refs is None: return
        code_var = widget_map.get("code_standard")
        safety_var = widget_map.get("safety_grade")
        sheet_refs["basic_cache"] = {
            "ground_level": str(Solid_Level.get()),
            "water_level": str(Water_Level.get()),
            "pile_type": str(Sheet_Pile_typevar.get()),
            "section_ref": str(Sheet_Pile_namevar.get()),
            "pile_top_level": str(CofferDam_Top_Level.get()),
            "pile_length": str(CofferDam_L.get()),
            "cap_length": str(Cap_X.get()),
            "cap_width": str(Cap_Y.get()),
            "cap_height": str(Cap_H.get()),
            "cap_bottom_level": str(Cap_Bottom_Level.get()),
            "cap_offset_long": str(X_offset.get()),
            "cap_offset_short": str(Y_offset.get()),
            "code_standard": str(code_var.get()) if code_var else "",
            "safety_grade": str(safety_var.get()) if safety_var else "",
            "skggz_D": str(SKGGZ_D.get()),
            "skggz_t": str(SKGGZ_t.get()),
            "skggz_gap": str(SKGGZ_Gap.get()),
            "skggz_lock_width": str(SKGGZ_LockWidth.get()),
            "skggz_sheet_count": str(SKGGZ_SheetCount.get()),
            "skggz_sheet_width": str(SKGGZ_SheetWidth.get()),
        }
    if sheet_refs is not None:
        sheet_refs["refresh_basic"] = _refresh_basic_cache
        
        combo_type.bind("<<ComboboxSelected>>", _dummy_frame_change)

# ===== 页面3: 支撑及封底 =====



# ===== 页面3: 支撑及封底 =====

def Waler_Strut_Concrete_Setting_GUI(table, setup_path, rows, blocks,
                                     Waler_section_dict, Strut_section_dict, Walers_Strut_selected_dict,
                                     Concrete_Plug_check_var, Concrete_Plug_thickness_var,
                                     Concrete_Plug_grade_var, Concrete_Plug_grade_lst,
                                     Concrete_Blinding_check_var, Concrete_Blinding_thickness_var,
                                     Concrete_Blinding_grade_var, Concrete_Blinding_grade_lst,
                                     widget_map=None, on_waler_change=None, _waler_actions=None,
                                     plug_support_d_var=None,
                                     plug_support_long_var=None,
                                     plug_support_short_var=None,
                                     plug_bond_var=None, plug_casing_count_var=None,
                                     plug_casing_diam_var=None,
                                     pile_coords_var=None, sheet_refs=None,
                                     steel_dict=None, projects_dict=None, current_proj_key_ref=None,
                                     concrete_dict=None):
    """构建"支撑及封底"页（Tab 3）：围檩/内支撑布置、垫层封底与封底参数。

    支持动态增删围檩层、每层选择材质与截面、配置对撑/斜撑布置、编辑桩坐标，
    并在 sheet_refs 注册 substructure_cache 刷新与材质/混凝土等级选项刷新。

    Args:
        table (tk.Widget): 页面容器。
        setup_path (str): 程序安装目录。
        rows (list[dict]): 围檩行控件集合，就地追加。
        blocks (dict): 内支撑布置变量 {层号: {'DC_X','DC_Y','XC_X','XC_Y'}}。
        Waler_section_dict, Strut_section_dict (dict): 围檩/内支撑截面库。
        Walers_Strut_selected_dict (dict): 初始选型数据。
        Concrete_Plug_check_var 等: 封底/垫层的开关、厚度、等级变量。
        widget_map (dict|None): 控件注册表（高亮联动）。
        on_waler_change (callable|None): 围檩行变化回调。
        _waler_actions (dict|None): 接收 add/del 回调，供项目切换时增删层。
        plug_support_*、plug_bond_var、plug_casing_*: 封底参数变量。
        pile_coords_var (tk.Variable|None): 桩坐标字符串变量。
        sheet_refs (dict|None): 跨 Tab 共享引用字典。
        steel_dict, concrete_dict (dict|None): 钢材/混凝土材料库。
        projects_dict (dict|None): 全部项目字典。
        current_proj_key_ref (list|None): 当前项目键引用。

    Returns:
        None: 副作用为构建 UI 并写入 sheet_refs/widget_map。
    """
    State = {'True': 'normal', 'False': 'disabled'}
    Waler_Strut_valuelst = []
    for value in Walers_Strut_selected_dict.values():
        Waler_Strut_valuelst.append({
            "entry": value['间距'], "combo1": value['围檩长边截面'],
            "combo2": value['对撑截面'], "combo3": value['斜撑截面']
        })

    _, card = make_card_frame(table, "内支撑及围檩布置")

    # 工具栏（框外）
    toolbar = tb.Frame(card)
    toolbar.pack(fill=X, pady=5)
    add_btn = tb.Button(toolbar, text="+ 增加一层", bootstyle=(OUTLINE, SUCCESS), width=12)
    add_btn.pack(side=LEFT, padx=5)
    del_btn = tb.Button(toolbar, text="- 删除一层", bootstyle=(OUTLINE, DANGER), width=12)
    del_btn.pack(side=LEFT, padx=5)

    # 输入内容区
    waler_inner = tk.Frame(card, bg="white", highlightbackground="#bbbbbb", highlightcolor="#60a5fa", highlightthickness=0.5)
    waler_inner.pack(fill=X, padx=5, pady=2)

    canvas = tk.Canvas(waler_inner, height=180, bg="white", highlightthickness=0)
    v_scrollbar = tb.Scrollbar(waler_inner, orient=VERTICAL, command=canvas.yview)
    h_scrollbar = tb.Scrollbar(card, orient=HORIZONTAL, command=canvas.xview)

    scrollable_frame = tk.Frame(canvas, bg="white")
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

    canvas.pack(side=LEFT, fill=BOTH, expand=True, pady=5)
    v_scrollbar.pack(side=RIGHT, fill=Y, pady=5)
    h_scrollbar.pack(side=BOTTOM, fill=X, padx=5)

    def _waler_on_mousewheel(event):
        """围檩区滚轮纵向滚动（到顶不越界）。"""
        if event.delta > 0 and canvas.yview()[0] <= 0:
            return
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _waler_on_enter(_):
        """鼠标进入围檩区时绑定全局滚轮事件。"""
        canvas.bind_all("<MouseWheel>", _waler_on_mousewheel)

    def _waler_on_leave(_):
        """鼠标离开围檩区时解绑全局滚轮事件。"""
        canvas.unbind_all("<MouseWheel>")

    canvas.bind("<Enter>", _waler_on_enter)
    canvas.bind("<Leave>", _waler_on_leave)
    current_row = [0]
    Walers_namelst = list(Waler_section_dict.keys())  # 围檩型号
    Struts_namelst = list(Strut_section_dict.keys())  # 内支撑型号
    # Walers_namelst.insert(0, '/')
    Struts_namelst.insert(0, '/')

    def _steel_material_options():
        """按当前项目的钢结构规范返回可选钢材牌号列表。

        Returns:
            list[str]: 牌号列表；无法解析时返回 []。
        """
        if not projects_dict or not current_proj_key_ref:
            return []
        _cur = projects_dict.get(current_proj_key_ref[0], {})
        _std = _resolve_steel_std(_cur.get("钢结构规范", ""), steel_dict or {})
        if _std and steel_dict and _std in steel_dict:
            return list(steel_dict[_std].get("牌号参数", {}).keys())
        return []

    def _concrete_grade_options():
        """按当前项目的混凝土规范返回可选混凝土强度等级列表。

        Returns:
            list[str]: 强度等级列表；无法解析时返回 []。
        """
        if not projects_dict or not current_proj_key_ref:
            return []
        _cur = projects_dict.get(current_proj_key_ref[0], {})
        _std = _resolve_concrete_std(_cur.get("混凝土规范", ""), concrete_dict or {})
        if _std and concrete_dict and _std in concrete_dict:
            return list(concrete_dict[_std].get("强度等级参数", {}).keys())
        return []
    combo_long_lst = [tb.StringVar(value=dict_i['combo1']) for dict_i in Waler_Strut_valuelst]
    combo_short_lst = [tb.StringVar(value=dict_i['combo1']) for dict_i in Waler_Strut_valuelst]
    combo_var3_lst = [tb.StringVar(value=dict_i['combo2']) for dict_i in Waler_Strut_valuelst]
    combo_var4_lst = [tb.StringVar(value=dict_i['combo3']) for dict_i in Waler_Strut_valuelst]
    entry_var1_lst = [tb.StringVar(value=dict_i['entry']) for dict_i in Waler_Strut_valuelst]
    Strut1_Xlst = [value['X对撑布置'] for value in Walers_Strut_selected_dict.values()] # ['/', '1000+2000']
    Strut1_Ylst = [value['Y对撑布置'] for value in Walers_Strut_selected_dict.values()] # ['/', '1500']
    Strut2_Xlst = [value['X斜撑布置'] for value in Walers_Strut_selected_dict.values()] # ['4000', '4500']
    Strut2_Ylst = [value['Y斜撑布置'] for value in Walers_Strut_selected_dict.values()] # ['2000', '2500']
    def add_row():
        """新增一层围檩/内支撑行，创建控件并绑定联动逻辑。

        Returns:
            None: 副作用为向 rows/blocks 追加数据并刷新缓存。
        """
        # 初始化
        current_i = current_row[0]
        # print('开始行数:', current_i)
        try:
            entry1_var = entry_var1_lst[current_i]
            combo_long_var = combo_long_lst[current_i]
            combo_short_var = combo_short_lst[current_i]
            combo3_var = combo_var3_lst[current_i]
            combo4_var = combo_var4_lst[current_i]
        except:
            entry1_var = tb.StringVar(value='')
            combo_long_var = tb.StringVar(value='')
            combo_short_var = tb.StringVar(value='')
            combo3_var = tb.StringVar(value='')
            combo4_var = tb.StringVar(value='')
        try:
            Strut1_X = tb.StringVar(value=Strut1_Xlst[current_i])
            Strut1_Y = tb.StringVar(value=Strut1_Ylst[current_i])
            Strut2_X = tb.StringVar(value=Strut2_Xlst[current_i])
            Strut2_Y = tb.StringVar(value=Strut2_Ylst[current_i])
        except:
            Strut1_X = tb.StringVar(value='')
            Strut1_Y = tb.StringVar(value='')
            Strut2_X = tb.StringVar(value='')
            Strut2_Y = tb.StringVar(value='')
        # 行
        row_frame = tb.Frame(scrollable_frame)
        row_frame.pack(fill=X, pady=6)
        mat_opts = _steel_material_options()
        strut_mat_opts = ['/'] + mat_opts
        # 间距
        label1 = tb.Label(row_frame, text=f"间距{current_i+1}(m):", width=8)
        label1.pack(side=LEFT)
        entry1 = tb.Entry(row_frame, textvariable=entry1_var, width=8)
        entry1.pack(side=LEFT, padx=5)
        createToolTip(entry1, '内支撑层距, 仅第一层为桩顶距, 单位m')
        # 围檩长边材质/截面
        mat_long_var = tb.StringVar(value='')
        label2m = tb.Label(row_frame, text=f"围檩长边{current_i+1}材质:", width=12)
        label2m.pack(side=LEFT)
        mat_long_w = tb.Combobox(row_frame, textvariable=mat_long_var, values=mat_opts, width=8, state="readonly")
        mat_long_w.pack(side=LEFT, padx=2)
        label2a = tb.Label(row_frame, text=f"围檩长边{current_i+1}截面:", width=12)
        label2a.pack(side=LEFT)
        combo_long = tb.Combobox(row_frame, textvariable=combo_long_var, values=Walers_namelst, width=10, state="readonly")
        combo_long.pack(side=LEFT, padx=2)
        # 围檩短边材质/截面
        mat_short_var = tb.StringVar(value='')
        label2bm = tb.Label(row_frame, text=f"围檩短边{current_i+1}材质:", width=12)
        label2bm.pack(side=LEFT)
        mat_short_w = tb.Combobox(row_frame, textvariable=mat_short_var, values=mat_opts, width=8, state="readonly")
        mat_short_w.pack(side=LEFT, padx=2)
        label2b = tb.Label(row_frame, text=f"围檩短边{current_i+1}截面:", width=12)
        label2b.pack(side=LEFT)
        combo_short = tb.Combobox(row_frame, textvariable=combo_short_var, values=Walers_namelst, width=10, state="readonly")
        combo_short.pack(side=LEFT, padx=2)
        # 对撑材质/截面
        mat_dc_var = tb.StringVar(value='')
        label3m = tb.Label(row_frame, text=f"对撑{current_i+1}材质:", width=8)
        label3m.pack(side=LEFT)
        mat_dc_w = tb.Combobox(row_frame, textvariable=mat_dc_var, values=strut_mat_opts, width=8, state="readonly")
        mat_dc_w.pack(side=LEFT, padx=2)
        label3 = tb.Label(row_frame, text=f"对撑{current_i+1}截面:", width=8)
        label3.pack(side=LEFT)
        combo3 = tb.Combobox(row_frame, textvariable=combo3_var, values=Struts_namelst, width=10, state="readonly")
        combo3.pack(side=LEFT, padx=5)
        # 斜撑材质/截面
        mat_xc_var = tb.StringVar(value='')
        label4m = tb.Label(row_frame, text=f"斜撑{current_i+1}材质:", width=8)
        label4m.pack(side=LEFT)
        mat_xc_w = tb.Combobox(row_frame, textvariable=mat_xc_var, values=strut_mat_opts, width=8, state="readonly")
        mat_xc_w.pack(side=LEFT, padx=2)
        label4 = tb.Label(row_frame, text=f"斜撑{current_i+1}截面:", width=8)
        label4.pack(side=LEFT)
        combo4 = tb.Combobox(row_frame, textvariable=combo4_var, values=Struts_namelst, width=10, state="readonly")
        combo4.pack(side=LEFT, padx=5)
        button1 = tb.Button(row_frame, text="支撑{}布置".format(current_i+1), width=10, bootstyle=INFO, command=lambda: Add_Strut_Place_Setting(table, button1, blocks, combo3, combo4, Strut1_X, Strut1_Y, Strut2_X, Strut2_Y, current_i))
        button1.pack(side=RIGHT)
        rows.append({"frame": row_frame, "entry": entry1_var, "combo_long": combo_long_var, "combo_short": combo_short_var, "combo2": combo3_var, "combo3": combo4_var,
                     "mat_long": mat_long_var, "mat_short": mat_short_var, "mat_dc": mat_dc_var, "mat_xc": mat_xc_var,
                     "entry_w": entry1, "combo1_w": combo_long, "combo_short_w": combo_short, "combo2_w": combo3, "combo3_w": combo4, "btn_w": button1,
                     "mat_long_w": mat_long_w, "mat_short_w": mat_short_w, "mat_dc_w": mat_dc_w, "mat_xc_w": mat_xc_w})
        blocks[current_i] = {'DC_X':Strut1_X, 'DC_Y':Strut1_Y, 'XC_X':Strut2_X, 'XC_Y':Strut2_Y}

        # 对撑/斜撑选择 "/" 时，自动将布置填入 "/"
        def _make_dc_trace(combo_value, strut_x_var, strut_y_var):
            """生成对撑截面 trace 回调：截面为 '/' 时清空其长短边布置。"""
            def _cb(*_):
                """对撑截面变化时的实际回调。"""
                if combo_value.get() == '/':
                    strut_x_var.set('/'); strut_y_var.set('/')
            return _cb
        def _make_xc_trace(combo_value, strut_x_var, strut_y_var):
            """生成斜撑截面 trace 回调：截面为 '/' 时清空其长短边布置。"""
            def _cb(*_):
                """斜撑截面变化时的实际回调。"""
                if combo_value.get() == '/':
                    strut_x_var.set('/'); strut_y_var.set('/')
            return _cb
        combo3_var.trace_add('write', _make_dc_trace(combo3_var, Strut1_X, Strut1_Y))
        combo4_var.trace_add('write', _make_xc_trace(combo4_var, Strut2_X, Strut2_Y))

        # 对撑/斜撑截面变化时重新计算内支撑层数，更新牛腿+连接定义状态
        def _on_strut_sec_change(*_):
            """对撑/斜撑截面变化时重算内支撑层数并更新牛腿/连接定义启用状态。"""
            update_fn = sheet_refs.get("_update_waler_dependent_state")
            if update_fn:
                n_strut = sum(1 for r in rows
                              if r.get("combo2", tb.StringVar()).get() not in ("", "/")
                              or r.get("combo3", tb.StringVar()).get() not in ("", "/"))
                update_fn(len(rows), n_strut)
        combo3_var.trace_add('write', _on_strut_sec_change)
        combo4_var.trace_add('write', _on_strut_sec_change)

        current_row[0] = current_row[0] + 1
        entry1.bind("<KeyRelease>", lambda e: on_waler_change())
        if on_waler_change:
            on_waler_change()

    def del_row():
        """删除最后一层围檩/内支撑行并刷新相关缓存与图示。"""
        if rows:
            row = rows.pop()
            row["frame"].destroy()
            current_row[0] = current_row[0] - 1
            if on_waler_change:
                on_waler_change()

    add_btn.config(command=add_row)
    del_btn.config(command=del_row)
    if _waler_actions is not None:
        _waler_actions['add'] = add_row
        _waler_actions['del'] = del_row
    for i in range(len(Waler_Strut_valuelst)):
        current_row[0] = i
        add_row()

    # 垫层/封底
    _, concrete_card = make_card_frame(table, "垫层及封底设置")

    # 垫层行
    pad_row = tb.Frame(concrete_card)
    pad_row.pack(fill=X, pady=4)
    pad_row.columnconfigure(1, weight=1)  # 弹性空间，将右侧控件推至右端
    check_blinding = tb.Checkbutton(pad_row, text='混凝土垫层', variable=Concrete_Blinding_check_var,
                          bootstyle=PRIMARY,
                          command=lambda: toggle_control(Concrete_Blinding_check_var,
                                                         Concrete_Plug_check_var,
                                                         [entry_blinding_thickness, combo_blinding_grade], [entry_plug_thickness, combo_plug_grade]))
    check_blinding.grid(row=0, column=0, sticky=W, padx=10)
    tb.Label(pad_row, text="混凝土等级:", width=12, anchor=E, bootstyle=SECONDARY).grid(row=0, column=2, padx=5)
    combo_blinding_grade = tb.Combobox(pad_row, textvariable=Concrete_Blinding_grade_var,
                      values=_concrete_grade_options(), width=8,
                      state=State[str(Concrete_Blinding_check_var.get())])
    combo_blinding_grade.grid(row=0, column=3, padx=5)
    tb.Label(pad_row, text="垫层厚度(m):", width=15, anchor=E, bootstyle=SECONDARY).grid(row=0, column=4, padx=5)
    entry_blinding_thickness = tb.Entry(pad_row, textvariable=Concrete_Blinding_thickness_var, width=10,
                   state=State[str(Concrete_Blinding_check_var.get())])
    entry_blinding_thickness.grid(row=0, column=5, padx=5)
    createToolTip(entry_blinding_thickness, '承台底标高-垫层厚度=基坑底标高')

    # 封底行
    plug_row = tb.Frame(concrete_card)
    plug_row.pack(fill=X, pady=4)
    plug_row.columnconfigure(1, weight=1)  # 弹性空间，将右侧控件推至右端
    check_plug = tb.Checkbutton(plug_row, text='封底混凝土', variable=Concrete_Plug_check_var,
                          bootstyle=PRIMARY,
                          command=lambda: toggle_control(Concrete_Plug_check_var,
                                                         Concrete_Blinding_check_var,
                                                         [entry_plug_thickness, combo_plug_grade], [entry_blinding_thickness, combo_blinding_grade]))
    check_plug.grid(row=0, column=0, sticky=W, padx=10)
    tb.Label(plug_row, text="混凝土等级:", width=12, anchor=E, bootstyle=SECONDARY).grid(row=0, column=2, padx=5)
    combo_plug_grade = tb.Combobox(plug_row, textvariable=Concrete_Plug_grade_var,
                      values=_concrete_grade_options(), width=8,
                      state=State[str(Concrete_Plug_check_var.get())])
    combo_plug_grade.grid(row=0, column=3, padx=5)
    tb.Label(plug_row, text="封底厚度(m):", width=15, anchor=E, bootstyle=SECONDARY).grid(row=0, column=4, padx=5)
    entry_plug_thickness = tb.Entry(plug_row, textvariable=Concrete_Plug_thickness_var, width=10,
                   state=State[str(Concrete_Plug_check_var.get())])
    entry_plug_thickness.grid(row=0, column=5, padx=5)

    # 封底参数 frame（仅封底勾选时显示）
    plug_param_frame = tb.LabelFrame(concrete_card, text="封底参数", bootstyle=PRIMARY)

    def open_pile_coords_dialog():
        """打开桩坐标编辑窗口（tksheet 增删行），保存为 "(x,y), (x,y)" 字符串。"""
        import ast
        dialog = tb.Toplevel(table)
        dialog.title("桩坐标设置")
        dialog.geometry("400x400+%d+%d" % (table.winfo_rootx()+80, table.winfo_rooty()+80))
        dialog.transient(table); dialog.grab_set()

        # ===== 底部按钮 =====
        btn_frame = tb.Frame(dialog)
        btn_frame.pack(side=BOTTOM, fill=X, padx=15, pady=(6, 10))
        tb.Button(btn_frame, text="保存", bootstyle=SUCCESS, width=10,
                    command=lambda: [pile_coords_var and pile_coords_var.set(
                        ", ".join([f"({r[0]},{r[1]})" for r in sheet.get_sheet_data()
                                    if len(r) >= 2 and str(r[0]).strip() and str(r[1]).strip()])),
                                        dialog.destroy()]).pack(side=RIGHT, padx=5)
        tb.Button(btn_frame, text="取消", bootstyle=SECONDARY, width=10,
                    command=dialog.destroy).pack(side=RIGHT, padx=5)

        # ===== 顶部工具栏 =====
        tool_frame = tb.Frame(dialog)
        tool_frame.pack(side=TOP, fill=X, padx=15, pady=(10, 2))
        tb.Button(tool_frame, text="+ 新增行", bootstyle=(OUTLINE, SUCCESS), width=10,
                    command=lambda: sheet.insert_row()).pack(side=LEFT, padx=2)
        tb.Button(tool_frame, text="- 删除行", bootstyle=(OUTLINE, DANGER), width=10,
                    command=lambda: [sheet.delete_row(r) for r in reversed(sorted(sheet.get_selected_rows()))] if sheet.get_selected_rows() else None).pack(side=LEFT, padx=2)

        # ===== tksheet =====
        from tksheet import Sheet
        sheet_frame = tk.Frame(dialog, bg="white", highlightbackground="#d0d4d9", highlightthickness=1)
        sheet_frame.pack(fill=BOTH, expand=True, padx=15, pady=(6, 0))

        coords = []
        if pile_coords_var:
            cs = pile_coords_var.get()
            if cs:
                for item in cs.split(", "):
                    try:
                        c = ast.literal_eval(item)
                        if isinstance(c, (tuple, list)) and len(c) == 2:
                            coords.append(c)
                    except:
                        pass

        sheet = Sheet(sheet_frame,
                        data=[[str(c[0]), str(c[1])] for c in coords],
                        headers=["x 坐标", "y 坐标"],
                        font=("Microsoft YaHei UI", 10, "normal"),
                        header_font=("Microsoft YaHei UI", 10, "bold"),
                        show_row_index=True, row_index_width=40,
                        header_height=30, row_height=28,
                        theme="light",
                        default_column_width=140,
                        border_color="#cbd5e1",
                        grid_color="#e2e8f0",
                        header_bg="#f1f5f9",
                        header_fg="#1e293b",
                        )
        sheet.pack(fill=BOTH, expand=True)
        sheet.enable_bindings(
            "single_select", "row_select", "arrowkeys",
            "rc_delete_row", "delete", "cut", "paste",
            "edit_cell")
        dialog.wait_window()
        
    # 第一行：导入DXF按钮行
    plug_btn_row = tb.Frame(plug_param_frame)
    plug_btn_row.pack(fill=X, pady=(10, 2), padx=10)
    tb.Button(plug_btn_row, text="导入桩位布置DXF", bootstyle=PRIMARY, width=18).pack(side=LEFT, pady=5,padx=5)
    tb.Button(plug_btn_row, text="编辑桩坐标", bootstyle=(OUTLINE, PRIMARY), width=15, command=open_pile_coords_dialog).pack(side=RIGHT, padx=5, pady=5)
    # 支撑间距拆分为长短边（使用模块级变量）
    # 同步 Plug_support_d_var (兼容旧代码)
    def _sync_plug_d_to_long_short(*_):
        """把封底长边支撑间距同步到兼容变量 Plug_support_d_var。"""
        if plug_support_long_var:
            plug_support_d_var.set(round(plug_support_long_var.get(), 3))

    # 第二、三行
    P_LBL_W, P_ENT_W = 17, 12
    # Row 1: 封底长边支撑间距(左) | 封底短边支撑间距(右)
    plug_row1 = tb.Frame(plug_param_frame)
    plug_row1.pack(fill=X, pady=5, padx=10)
    tb.Label(plug_row1, text="封底长边支撑间距(m):", width=P_LBL_W, bootstyle=PRIMARY).pack(side=LEFT, padx=5)
    tb.Entry(plug_row1, textvariable=plug_support_long_var, width=P_ENT_W).pack(side=LEFT, padx=5)
    tb.Entry(plug_row1, textvariable=plug_support_short_var, width=P_ENT_W).pack(side=RIGHT, padx=5)
    tb.Label(plug_row1, text="封底短边支撑间距(m):", width=P_LBL_W, bootstyle=PRIMARY).pack(side=RIGHT, padx=5)
    if plug_support_long_var:
        plug_support_long_var.trace_add("write", _sync_plug_d_to_long_short)
    if plug_support_short_var:
        plug_support_short_var.trace_add("write", _sync_plug_d_to_long_short)
    # Row 2: 钢混黏聚(左) | (自动撑开) | 护筒数量(右) | 护筒直径(右)
    # Row 2: 钢混黏聚 | 护筒数量(居中) | 护筒直径 — 3列等宽
    plug_row2 = tb.Frame(plug_param_frame)
    plug_row2.pack(fill=X, pady=(5, 10), padx=10)
    plug_row2.grid_columnconfigure((0, 1, 2), weight=1, uniform="plug2")
    f_a = tb.Frame(plug_row2); f_a.grid(row=0, column=0, sticky="ew", padx=5)
    tb.Label(f_a, text="钢混黏聚(kPa):", width=P_LBL_W, bootstyle=PRIMARY).pack(side=LEFT)
    tb.Entry(f_a, textvariable=plug_bond_var, width=P_ENT_W).pack(side=LEFT, padx=10)
    f_b = tb.Frame(plug_row2); f_b.grid(row=0, column=1, sticky="ew", padx=5)
    tb.Label(f_b, text="护筒数量:", width=10, bootstyle=PRIMARY).pack(side=LEFT)
    tb.Entry(f_b, textvariable=plug_casing_count_var, width=P_ENT_W).pack(side=LEFT, padx=10)
    f_c = tb.Frame(plug_row2); f_c.grid(row=0, column=2, sticky="ew", padx=5)
    tb.Entry(f_c, textvariable=plug_casing_diam_var, width=P_ENT_W).pack(side=RIGHT)
    tb.Label(f_c, text="护筒直径(m):", width=P_LBL_W, bootstyle=PRIMARY).pack(side=RIGHT, padx=10)
    def _toggle_plug_param(*_):
        """勾选封底时显示"封底参数"区域，否则隐藏。"""
        if Concrete_Plug_check_var.get():
            plug_param_frame.pack(fill=X, padx=10, pady=5, after=plug_row)
        else:
            plug_param_frame.pack_forget()
    Concrete_Plug_check_var.trace_add("write", _toggle_plug_param)
    _toggle_plug_param()

    if widget_map is not None:
        widget_map["concrete_blinding"] = [entry_blinding_thickness, combo_blinding_grade]
        widget_map["concrete_plug"] = [entry_plug_thickness, combo_plug_grade]
        widget_map["_waler_entries"] = waler_entry_widgets = []
        for row in rows:
            if "entry_w" in row:
                waler_entry_widgets.append(row["entry_w"])

    # ---- substructure cache ----
    def _refresh_substructure_cache():
        """把围檩/内支撑/垫层封底当前值写入 sheet_refs['substructure_cache']。"""
        if sheet_refs is None: return
        waler_data = []
        for i, row in enumerate(rows):
            b = blocks.get(i, {})
            waler_data.append({
                "spacing": str(row['entry'].get()),
                "sec_long": str(row['combo_long'].get()),
                "sec_short": str(row['combo_short'].get()),
                "mat_long": str(row['mat_long'].get()),
                "mat_short": str(row['mat_short'].get()),
                "dc_sec": str(row['combo2'].get()),
                "xc_sec": str(row['combo3'].get()),
                "mat_dc": str(row['mat_dc'].get()),
                "mat_xc": str(row['mat_xc'].get()),
                "dc_x": str(b.get('DC_X', tb.StringVar()).get()),
                "dc_y": str(b.get('DC_Y', tb.StringVar()).get()),
                "xc_x": str(b.get('XC_X', tb.StringVar()).get()),
                "xc_y": str(b.get('XC_Y', tb.StringVar()).get()),
            })
        sheet_refs["substructure_cache"] = {
            "blinding_grade": str(Concrete_Blinding_grade_var.get()),
            "blinding_thickness": str(Concrete_Blinding_thickness_var.get()),
            "plug_grade": str(Concrete_Plug_grade_var.get()),
            "plug_thickness": str(Concrete_Plug_thickness_var.get()),
            "plug_support_d": str(plug_support_d_var.get()) if plug_support_d_var else "0",
            "plug_support_long": str(plug_support_long_var.get()) if plug_support_long_var else "0",
            "plug_support_short": str(plug_support_short_var.get()) if plug_support_short_var else "0",
            "plug_bond": str(plug_bond_var.get()),
            "plug_casing_count": str(plug_casing_count_var.get()),
            "plug_casing_diam": str(plug_casing_diam_var.get()),
            "waler_rows": waler_data,
        }
    if sheet_refs is not None:
        sheet_refs["refresh_substructure"] = _refresh_substructure_cache

        def _refresh_waler_material_options():
            """按当前项目钢结构规范刷新围檩/内支撑材质下拉选项。"""
            opts = _steel_material_options()
            strut_opts = ['/'] + opts
            for row in rows:
                for k in ("mat_long_w", "mat_short_w"):
                    w = row.get(k)
                    if w is not None and w.winfo_exists():
                        w.configure(values=opts)
                for k in ("mat_dc_w", "mat_xc_w"):
                    w = row.get(k)
                    if w is not None and w.winfo_exists():
                        w.configure(values=strut_opts)

        sheet_refs["_refresh_waler_material_options"] = _refresh_waler_material_options

        def _refresh_concrete_grade_options():
            """按当前项目混凝土规范刷新垫层/封底混凝土等级下拉选项。"""
            opts = _concrete_grade_options()
            for w in (combo_blinding_grade, combo_plug_grade):
                if w is not None and w.winfo_exists():
                    w.configure(values=opts)

        sheet_refs["_refresh_concrete_grade_options"] = _refresh_concrete_grade_options

        sheet_refs["_get_steel_material_options"] = _steel_material_options

def Add_Strut_Place_Setting(table, button, blocks, DC_combo, XC_combo, Strut1_X, Strut1_Y, Strut2_X, Strut2_Y, current_i):
    """打开某层内支撑布置参数二级窗口（对撑/斜撑长短边间距）。

    根据对撑/斜撑截面是否为 "/" 决定输入框启用状态；关闭窗口时恢复入口按钮。

    Args:
        table (tk.Widget): 父容器。
        button (tk.Button): 触发本窗口的按钮（打开时禁用、关闭时恢复）。
        blocks (dict): 内支撑布置变量 {层号: {...}}。
        DC_combo, XC_combo (tk.Variable): 对撑/斜撑截面变量。
        Strut1_X, Strut1_Y (tk.Variable): 对撑长边/短边布置（m）。
        Strut2_X, Strut2_Y (tk.Variable): 斜撑长边/短边布置（m）。
        current_i (int): 当前层号（0 基）。

    Returns:
        None
    """
    DC_state = 'normal' if DC_combo.get() != '/' else 'disabled'
    XC_state = 'normal' if XC_combo.get() != '/' else 'disabled'
    dialog = tb.Toplevel(table)
    dialog.title("内支撑参数表")
    container = tb.Frame(dialog)
    container.pack(fill=X, pady=5)
    labelframe1 = tb.Labelframe(container, text="第{}层对撑布置".format(current_i+1), bootstyle=PRIMARY)
    labelframe1.pack(side=LEFT, fill=X, padx=15, pady=10)
    row1_frame = tb.Frame(labelframe1)
    row1_frame.pack(fill=X, pady=5)
    label_dc_long = tb.Label(master=row1_frame, text="长边对撑间距布置(m): ", width=17)
    label_dc_long.pack(side=LEFT, padx=15, pady=10, fill=X)
    entry3 = tb.Entry(master=row1_frame, textvariable=Strut1_X, width=15, state=DC_state)
    entry3.pack(side=LEFT, padx=10, pady=10, fill=X)
    row2_frame = tb.Frame(labelframe1)
    row2_frame.pack(fill=X, pady=5)
    label_dc_short = tb.Label(master=row2_frame, text="宽边对撑间距布置(m): ", width=17)
    label_dc_short.pack(side=LEFT, padx=15, pady=10, fill=X)
    entry5 = tb.Entry(master=row2_frame, textvariable=Strut1_Y, width=15, state=DC_state)
    entry5.pack(side=LEFT, padx=10, pady=10, fill=X)
    labelframe2 = tb.Labelframe(container, text="第{}层斜撑布置".format(current_i+1), bootstyle=PRIMARY)
    labelframe2.pack(side=RIGHT, fill=X, padx=15, pady=10)
    row3_frame = tb.Frame(labelframe2)
    row3_frame.pack(fill=X, pady=5)
    label_xc_long = tb.Label(master=row3_frame, text="长边斜撑间距布置(m): ", width=17)
    label_xc_long.pack(side=LEFT, padx=15, pady=10, fill=X)
    entry4 = tb.Entry(master=row3_frame, textvariable=Strut2_X, width=15, state=XC_state)
    entry4.pack(side=LEFT, padx=10, pady=10, fill=X)
    row4_frame = tb.Frame(labelframe2)
    row4_frame.pack(fill=X, pady=5)
    label_xc_short = tb.Label(master=row4_frame, text="宽边斜撑间距布置(m): ", width=17)
    label_xc_short.pack(side=LEFT, padx=15, pady=10, fill=X)
    entry6 = tb.Entry(master=row4_frame, textvariable=Strut2_Y, width=15, state=XC_state)
    entry6.pack(side=LEFT, padx=10, pady=10, fill=X)
    blocks[current_i] = {'DC_X':Strut1_X, 'DC_Y':Strut1_Y, 'XC_X':Strut2_X, 'XC_Y':Strut2_Y}

    # 禁用主窗口的按钮
    button.config(state=tk.DISABLED)
    # 当二级窗口关闭时，重新启用主窗口的按钮
    dialog.protocol("WM_DELETE_WINDOW", lambda:enable_button(dialog, button))

# ===== 页面4: 土层及边界 =====
def Soil_Info_Setting_GUI(table, Steel_Sheet_Pile_CofferDam_Load_dict,
                         on_soil_change=None, sheet_refs=None,
                         spring_thickness_var=None, consider_stress_path_var=None):
    """构建"土层信息"页（Tab 4）：土层分层表、加权平均参数与土弹簧设置。

    Args:
        table (tk.Widget): 页面容器。
        Steel_Sheet_Pile_CofferDam_Load_dict (dict): 土层信息字典（含 'Excel'）。
        on_soil_change (callable|None): 土层变化回调（刷新图示）。
        sheet_refs (dict|None): 跨 Tab 共享引用字典。
        spring_thickness_var (tk.Variable|None): 土弹簧分层厚度（m）。
        consider_stress_path_var (tk.Variable|None): 是否考虑土的应力路径。

    Returns:
        None
    """
    # ===== 土层信息 card（嵌入 SoilSettingsFrame） =====
    _, card = make_card_frame(table, "地质参数")

    soil_headers = [
        "土层名称", "层厚(m)", "重度(kN/m3)",
        "黏聚力c(kPa)", "内摩擦角φ(°)", "计算方法",
        "渗透系数(m/d)", "土层顶承压水头(m)"
    ]

    # 获取全部土层 sheet 名，供"选择已有"下拉
    soil_sheet_names = get_soil_sheet_names(get_excel_path()) if get_excel_path() else []

    def _on_existing_soil_selected(sheet_name):
        """选择已有土层 sheet 时加载其数据到土层表并更新引用与基准。

        Args:
            sheet_name (str): 选中的土层 sheet 名。

        Returns:
            None
        """
        ep = get_excel_path()
        if not ep:
            return
        sheet_data = load_soil_sheet(ep, sheet_name)
        if not sheet_data:
            return
        # 转为 Steel_Sheet_Pile_CofferDam_Load_dict['Excel'] 格式
        new_excel = {}
        for i, (name, vals) in enumerate(sheet_data.items()):
            entry = {"土层名称": name}
            field_keys = ["层厚", "重度", "黏聚力", "内摩擦角",
                         "计算方法", "渗透系数(m/d)", "土层顶承压水头(m)"]
            for j, k in enumerate(field_keys):
                entry[k] = vals[j] if j < len(vals) else ""
            new_excel[f"第{i+1}层土"] = entry
        Steel_Sheet_Pile_CofferDam_Load_dict["Excel"] = new_excel
        _load_soil_into_frame(soil_frame, Steel_Sheet_Pile_CofferDam_Load_dict)
        # 更新 ref
        if sheet_refs is not None:
            sheet_refs["cofferdam_soil_ref"] = sheet_name
        # 更新原始数据基准（用于检测后续是否修改）
        if sheet_refs is not None:
            sheet_refs["cofferdam_soil_ref"] = sheet_name
            # 将参照数据转为 list 格式存储，与 tksheet 返回格式一致
            _ref_soil_data = {}
            for name, vals in sheet_data.items():
                _ref_soil_data[name] = list(vals) if not isinstance(vals, list) else vals
            proj_key = sheet_refs.get("_current_proj_key", [None])[0]
            if proj_key:
                sheet_refs["_cofferdam_soil_original"][proj_key] = _ref_soil_data
            # 土层变更后刷新 RealTimeDiagram
            if on_soil_change:
                try: on_soil_change()
                except: pass

    # ── 加权平均土层参数 ──
    _weighted_check = tb.BooleanVar(value=False)
    _weighted_row = tb.Frame(card)
    _weighted_row.pack(fill=X, padx=15, pady=(10, 0))
    tb.Checkbutton(_weighted_row, text="采用加权平均土层参数",
                    variable=_weighted_check, bootstyle="round-toggle").pack(side=LEFT,padx=10)
    _weighted_entries = tb.Frame(card)
    _weighted_entries.pack(fill=X, padx=15, pady=(10, 5))
    _weighted_entries.grid_columnconfigure((0, 1, 2), weight=1, uniform="wavg")
    _w_重度 = tb.DoubleVar(value=0.0)
    _w_黏聚力 = tb.DoubleVar(value=0.0)
    _w_内摩擦角 = tb.DoubleVar(value=0.0)
    for _ci, (_label, _var, _attr) in enumerate([
        ("加权平均重度(kN/m3):", _w_重度, "_w_entry_重度"),
        ("加权平均黏聚力c(kPa):", _w_黏聚力, "_w_entry_黏聚力"),
        ("加权平均内摩擦角φ(°):", _w_内摩擦角, "_w_entry_内摩擦角"),
    ]):
        grp = tb.Frame(_weighted_entries)
        grp.grid(row=0, column=_ci, sticky="ew", padx=5)
        tb.Label(grp, text=_label, bootstyle=PRIMARY).pack(side=LEFT, padx=(0, 5), pady=5)
        e = tb.Entry(grp, textvariable=_var, width=10, state=tk.DISABLED)
        e.pack(side=LEFT, fill=X, expand=True, pady=5)
        if _ci == 0: _w_entry_重度 = e
        elif _ci == 1: _w_entry_黏聚力 = e
        else: _w_entry_内摩擦角 = e
    def _toggle_weighted_entries(*_):
        """勾选"采用加权平均土层参数"时启用三个加权输入框，否则禁用。"""
        st = tk.NORMAL if _weighted_check.get() else tk.DISABLED
        _w_entry_重度.configure(state=st)
        _w_entry_黏聚力.configure(state=st)
        _w_entry_内摩擦角.configure(state=st)
    _weighted_check.trace_add("write", _toggle_weighted_entries)
    # 存入 sheet_refs 以便保存
    if sheet_refs is not None:
        sheet_refs["_weighted_重度"] = _w_重度
        sheet_refs["_weighted_黏聚力"] = _w_黏聚力
        sheet_refs["_weighted_内摩擦角"] = _w_内摩擦角
        sheet_refs["_weighted_check"] = _weighted_check

    soil_frame = SoilSettingsFrame(
        card,
        headers=soil_headers,
        row_count=20,
        show_buttons=False,
        height=300,
        col_widths=[100, 80, 100, 110, 110, 100, 120, 160],
        column_dropdowns={5: ["水土合算", "水土分算"]},
        existing_options=soil_sheet_names,
        on_existing_selected=_on_existing_soil_selected,
    )
    soil_frame.pack(fill=BOTH, expand=True, padx=10, pady=5)

    # 初始加载数据
    _load_soil_into_frame(soil_frame, Steel_Sheet_Pile_CofferDam_Load_dict)

    if sheet_refs is not None:
        sheet_refs["soil_frame"] = soil_frame
    # 土层编辑后刷新 RealTimeDiagram 立面图（绑定 tksheet 编辑完成事件）
    if on_soil_change:
        def _sync_soil_and_refresh(*_):
            """土层表编辑完成后同步到 Load_dict 并刷新立面图。"""
            try:
                raw = soil_frame.sheet.get_sheet_data()
                new_dict = {}
                soil_idx = 0
                for row in raw:
                    if not row or not row[0]: continue
                    name = str(row[0]).strip()
                    if not name: continue
                    soil_idx += 1
                    new_dict[f"第{soil_idx}层土"] = {
                        "土层名称": name,
                        "层厚": str(row[1]) if len(row) > 1 else "/",
                        "重度": str(row[2]) if len(row) > 2 else "/",
                        "黏聚力": str(row[3]) if len(row) > 3 else "/",
                        "内摩擦角": str(row[4]) if len(row) > 4 else "/",
                        "计算方法": str(row[5]) if len(row) > 5 else "/",
                        "渗透系数(m/d)": str(row[6]) if len(row) > 6 else "/",
                        "土层顶承压水头(m)": str(row[7]) if len(row) > 7 else "/",
                    }
                Steel_Sheet_Pile_CofferDam_Load_dict["Excel"] = new_dict
                if on_soil_change: on_soil_change()
            except: pass
        try:
            soil_frame.sheet.extra_bindings("end_edit_cell", _sync_soil_and_refresh)
        except:
            pass

    # 土弹簧分层厚度
    spring_row = tb.Frame(card)
    spring_row.pack(fill=X, padx=5, pady=(2, 8))
    tb.Label(spring_row, text="土弹簧分层厚度(m):", bootstyle=PRIMARY, width=16).pack(side=LEFT, padx=10)
    tb.Entry(spring_row, textvariable=spring_thickness_var, width=12).pack(side=LEFT)
    # 是否考虑土的应力路径
    if consider_stress_path_var is not None:
        stress_path_frame = tb.Frame(spring_row)
        stress_path_frame.pack(side=RIGHT, padx=10)
        tb.Label(stress_path_frame, text="考虑土的应力路径:", bootstyle=PRIMARY, width=16).pack(side=LEFT)
        _stress_path_check_var = tk.BooleanVar(value=(consider_stress_path_var.get() == "是"))
        def _toggle_stress_path(_name=None, _index=None, _mode=None):
            """StringVar 变化时同步复选框状态。"""
            val = consider_stress_path_var.get()
            _stress_path_check_var.set(val == "是")
        def _on_check_toggle():
            """复选框切换时回写"是/否"到 StringVar。"""
            consider_stress_path_var.set("是" if _stress_path_check_var.get() else "否")
        # 双向同步：StringVar 变化 → 更新复选框
        consider_stress_path_var.trace_add("write", _toggle_stress_path)
        tb.Checkbutton(stress_path_frame, variable=_stress_path_check_var,
                       bootstyle="round-toggle", command=_on_check_toggle).pack(side=LEFT, padx=5)

def Boundary_Setting_GUI(table,
                        pile_boundary_var=None,
                        conn_data=None,
                        corbel_long_var=None, corbel_short_var=None,
                        skip_template_count=0,
                        sheet_refs=None):
    """构建"边界定义"页（Tab 5）：桩底边界、牛腿数量与四类弹性连接定义。

    连接定义用按钮切换四个 ConnectionSettingFrame；根据围檩/内支撑层数动态
    启用/禁用牛腿与连接控件，并注册 boundary_cache 刷新。

    Args:
        table (tk.Widget): 页面容器。
        pile_boundary_var (tk.Variable|None): 桩底边界编码变量。
        conn_data (list|None): 连接形式数据（来自 Excel）。
        corbel_long_var, corbel_short_var (tk.Variable|None): 长/短边牛腿数量。
        skip_template_count (int): 跳过连接形式列表中开头的模板项数。
        sheet_refs (dict|None): 跨 Tab 共享引用字典。

    Returns:
        None
    """
    # ===== 桩底边界 & 牛腿 =====
    _, pile_boundary_card = make_card_frame(table, "桩底边界定义")
    boundary_frame = tb.LabelFrame(pile_boundary_card, text="一般支承设置", bootstyle=PRIMARY)
    boundary_frame.pack(fill=X, pady=5, padx=5)
    PileBoundaryFrame(boundary_frame, boundary_var=pile_boundary_var).pack(fill=X, padx=5, pady=5)

    _, corbel_card = make_card_frame(table, "牛腿设置")
    corbel_row = tb.Frame(corbel_card)
    corbel_row.pack(fill=X, pady=5, padx=10)
    tb.Label(corbel_row, text="长边布置数量:", bootstyle=PRIMARY, width=16).pack(side=LEFT, padx=5)
    corbel_long_entry = tb.Entry(corbel_row, textvariable=corbel_long_var, width=12)
    corbel_long_entry.pack(side=LEFT, padx=5)
    corbel_short_entry = tb.Entry(corbel_row, textvariable=corbel_short_var, width=12)
    corbel_short_entry.pack(side=RIGHT, padx=(5, 0))
    tb.Label(corbel_row, text="短边布置数量:", bootstyle=PRIMARY, width=16).pack(side=RIGHT, padx=5)
    
    if sheet_refs is not None:
        sheet_refs["corbel_long_entry"] = corbel_long_entry
        sheet_refs["corbel_short_entry"] = corbel_short_entry

    # ===== 连接定义模块（重构版） =====
    _, stiffness_card = make_card_frame(table, "连接定义")
    CONN_DEFS = [
        {"name": "牛腿与支护桩连接", "key": "牛腿与支护桩连接",  "node_only": True,  "editable": False},
        {"name": "牛腿与围檩连接",   "key": "牛腿与围檩连接", "node_only": False, "editable": False},
        {"name": "围檩与支护桩连接", "key": "围檩与支护桩连接",   "node_only": False, "editable": True},
        {"name": "围檩与内支撑连接", "key": "围檩与内支撑连接",  "node_only": True,  "editable": False},
    ]

    conn_options = [c.get("name", "") for c in (conn_data[skip_template_count:] if skip_template_count > 0 else conn_data) if c.get("name")] if conn_data else []

    conn_btn_row = tb.Frame(stiffness_card)
    conn_btn_row.pack(fill=X, pady=(5, 2), padx=10)

    # 弹性连接设置（共节点：无滚动条；弹性连接：横向滚动条）
    conn_settings_outer = tb.LabelFrame(stiffness_card, text="弹性连接设置", bootstyle=PRIMARY, padding=0)
    conn_settings_outer.pack(fill=X, padx=10, pady=5)

    # 始终用 Canvas 作为容器，frame 放在 Canvas 中
    _conn_canvas = tk.Canvas(conn_settings_outer, highlightthickness=0, height=180)
    _conn_hscroll = tb.Scrollbar(conn_settings_outer, orient=HORIZONTAL, command=_conn_canvas.xview)
    _conn_canvas.configure(xscrollcommand=_conn_hscroll.set)

    conn_settings_frame = tb.Frame(_conn_canvas)
    _conn_window_id = _conn_canvas.create_window((0, 0), window=conn_settings_frame, anchor=NW)

    _conn_canvas.pack(side=TOP, fill=X)
    # 滚动条初始隐藏

    def _on_conn_frame_cfg(e):
        """连接内容尺寸变化时更新画布滚动区域。"""
        _conn_canvas.configure(scrollregion=_conn_canvas.bbox("all"))
    conn_settings_frame.bind("<Configure>", _on_conn_frame_cfg)

    def _on_conn_canvas_cfg(e):
        """画布尺寸变化时同步内容窗口宽度，避免内容被压缩。"""
        fw = conn_settings_frame.winfo_reqwidth()
        _conn_canvas.itemconfig(_conn_window_id, width=max(e.width, fw))
    _conn_canvas.bind("<Configure>", _on_conn_canvas_cfg)

    conn_frames, conn_btns, _active_conn_idx = [], [], [0]

    def _switch_conn(idx):
        """切换显示第 idx 类连接定义，并按是否共节点显隐滚动条与调整高度。"""
        if conn_frames:
            conn_frames[_active_conn_idx[0]].pack_forget()
            _active_conn_idx[0] = idx
            conn_frames[idx].pack(anchor=NW, padx=5, pady=5)
            # 弹性连接：显示滚动条；共节点：隐藏滚动条
            if CONN_DEFS[idx]["node_only"]:
                _conn_hscroll.pack_forget()
            else:
                _conn_hscroll.pack(side=BOTTOM, fill=X)
            # 延迟刷新：等布局完成后获取实际高度
            def _adjust_height():
                """布局完成后按内容高度调整画布（共节点自适应，弹性连接固定高度）。"""
                _conn_canvas.update_idletasks()
                fw = conn_settings_frame.winfo_reqheight()
                if CONN_DEFS[idx]["node_only"]:
                    _conn_canvas.configure(height=fw + 10)
                else:
                    _conn_canvas.configure(height=180)
                cw = _conn_canvas.winfo_width()
                _conn_canvas.itemconfig(_conn_window_id, width=max(cw, conn_settings_frame.winfo_reqwidth()))
                _conn_canvas.configure(scrollregion=_conn_canvas.bbox("all"))
            _conn_canvas.after(10, _adjust_height)
            for bi, btn in enumerate(conn_btns):
                btn.configure(bootstyle=PRIMARY if bi == idx else (PRIMARY, OUTLINE))

    for i, cdef in enumerate(CONN_DEFS):
        btn = tb.Button(conn_btn_row, text=cdef["name"], bootstyle=(PRIMARY, OUTLINE), command=lambda idx=i: _switch_conn(idx))
        btn.pack(side=LEFT, padx=2, fill=X, expand=True)
        conn_btns.append(btn)
        frame = ConnectionSettingFrame(conn_settings_frame, preset_type=None, stiffness=None, conn_options=conn_options, node_only=cdef["node_only"], editable=cdef["editable"])
        conn_frames.append(frame)

    _switch_conn(0)
    if sheet_refs:
        sheet_refs.update({"conn_frames": conn_frames, "conn_defs": CONN_DEFS})

    # ===== 状态更新逻辑 =====
    _waler_tip_label = tb.Label(table, text="* 未设置围檩及内支撑，无需设置牛腿参数和连接定义", font=("Microsoft YaHei UI", 9), bootstyle=WARNING)
    _strut_tip_label = tb.Label(table, text="* 未设置内支撑，无需设置围檩与内支撑连接定义", font=("Microsoft YaHei UI", 9), bootstyle=WARNING)

    def _update_waler_dependent_state(waler_count, strut_count=0):
        """（旧版）按围檩/内支撑层数启用牛腿与连接控件并显隐提示。

        注：随后存在同名新实现会覆盖本函数。

        Args:
            waler_count (int): 围檩层数。
            strut_count (int): 内支撑层数。

        Returns:
            None
        """
        no_both = (waler_count == 0 and strut_count == 0)
        no_strut = (strut_count == 0)
        
        # 更新控件状态
        state = "disabled" if waler_count == 0 else "normal"
        corbel_long_entry.configure(state=state)
        corbel_short_entry.configure(state=state)

        # 显隐设置区
        if no_both:
            conn_settings_outer.pack_forget()
            _waler_tip_label.pack(fill=X, padx=15, pady=(2, 8))
        else:
            conn_settings_outer.pack(fill=X, padx=10, pady=5)
            _waler_tip_label.pack_forget()
            if no_strut: _strut_tip_label.pack(fill=X, padx=15, pady=(2, 8))
            else: _strut_tip_label.pack_forget()

        for i, btn in enumerate(conn_btns):
            btn.configure(state="disabled" if (no_both or (no_strut and i == 3)) else "normal")

    if sheet_refs: sheet_refs["_update_waler_dependent_state"] = _update_waler_dependent_state

    # ===== 围檩/内支撑层数变化时禁用相关控件 =====
    _waler_tip_label = tb.Label(table, text="* 未设置围檩及内支撑，无需设置牛腿参数和连接定义",
                                 font=("Microsoft YaHei UI", 9, "normal"), bootstyle=WARNING)
    _waler_tip_label.pack(fill=X, padx=15, pady=(2, 8))
    _waler_tip_label.pack_forget()

    _strut_tip_label = tb.Label(table, text="* 未设置内支撑，无需设置围檩与内支撑连接定义",
                                 font=("Microsoft YaHei UI", 9, "normal"), bootstyle=WARNING)
    _strut_tip_label.pack(fill=X, padx=15, pady=(2, 8))
    _strut_tip_label.pack_forget()

    def _update_waler_dependent_state(waler_count, strut_count=0):
        """围檩/内支撑层数变化时，启用/禁用牛腿和连接定义控件

        - 无围檩且无内支撑：禁用牛腿，隐藏整个弹性连接设置区
        - 有围檩无内支撑：启用牛腿，仅禁用「围檩与内支撑」按钮
        - 有围檩有内支撑：全部启用
        """
        no_waler = (waler_count == 0)
        no_strut = (strut_count == 0)
        no_both = no_waler and no_strut

        # ---- 牛腿 entry ----
        corbel_state = "disabled" if no_waler else "normal"
        corbel_long_entry.configure(state=corbel_state)
        corbel_short_entry.configure(state=corbel_state)
        if no_waler:
            corbel_long_var.set("/")
            corbel_short_var.set("/")

        # ---- 连接按钮 ----
        for i, btn in enumerate(conn_btns):
            if no_both:
                btn.configure(state="disabled")
            elif no_strut and i == 3:
                # 有围檩无内支撑 → 仅禁用「围檩与内支撑」
                btn.configure(state="disabled")
            else:
                btn.configure(state="normal")

        # ---- 弹性连接设置区 + tips ----
        _waler_tip_label.pack_forget()
        _strut_tip_label.pack_forget()

        if no_both:
            # 无围檩无内支撑：隐藏整个弹性连接设置区
            conn_settings_outer.pack_forget()
            for frame in conn_frames:
                frame.pack_forget()
            _waler_tip_label.pack(fill=X, padx=15, pady=(2, 8))
        elif no_strut:
            # 有围檩无内支撑：显示弹性连接设置区，但提示无需设置围檩与内支撑连接
            conn_settings_outer.pack(fill=X, padx=10, pady=5)
            conn_frames[_active_conn_idx[0]].pack(fill=X, padx=5, pady=5)
            _strut_tip_label.pack(fill=X, padx=15, pady=(2, 8))
        else:
            # 有围檩有内支撑：全部正常显示
            conn_settings_outer.pack(fill=X, padx=10, pady=5)
            conn_frames[_active_conn_idx[0]].pack(fill=X, padx=5, pady=5)

    _update_waler_dependent_state(0)  # 初始状态，后续由 _switch_project 更新

    if sheet_refs is not None:
        sheet_refs["_update_waler_dependent_state"] = _update_waler_dependent_state

    def _collect_boundary_data():
        """收集边界定义页全部数据（桩底边界、牛腿数量、连接类型与刚度）。

        Returns:
            dict: 结构形式
                {'pile_bottom_boundary': str, 'corbel_long_num': str,
                 'corbel_short_num': str,
                 'connections': [{'key','name','type'}, ...]}。
        """
        result = {}
        result["pile_bottom_boundary"] = pile_boundary_var.get() if pile_boundary_var else ""
        result["corbel_long_num"] = corbel_long_var.get() if corbel_long_var else "/"
        result["corbel_short_num"] = corbel_short_var.get() if corbel_short_var else "/"
        
        # 显式组装主类 combine_cofferdam_dict 需要的 connections 列表
        connections_list = []
        form_dict = {}
        stiffness_dict = {}
        
        for i, cdef in enumerate(CONN_DEFS):
            if i < len(conn_frames):
                conn_type = conn_frames[i].get_type()
                stiffness = conn_frames[i].get_stiffness()
            else:
                conn_type = ""
                stiffness = {}
                
            form_dict[cdef["name"]] = conn_type
            
            # 将当前层弹性连接的 key 和最新的 type 存入 connections 列表
            connections_list.append({
                "key": cdef["key"],
                "name": cdef["name"],
                "type": conn_type
            })
            
            if conn_type and conn_type != "共节点" and stiffness:
                gen_keys = ["SDx", "SDy", "SDz", "SRx", "SRy", "SRz"]
                comp_keys = ["NSDx"]
                gen = {k.lower(): stiffness.get(k, "/") for k in gen_keys}
                comp = {k.lower(): stiffness.get(k, "/") for k in comp_keys}
                stiffness_dict[conn_type] = {"GEN": gen, "COMP": comp}
                
        result["connections"] = connections_list  # 写入 connections 键
        result["cofferdam_elasticlink_form_dict"] = form_dict
        result["sheet_elasticlink_dict"] = stiffness_dict
        return result

    # ---- boundary cache ----
    def _refresh_boundary_cache():
        """收集边界数据并写入 sheet_refs['boundary_cache']，供保存时读取。"""
        data = _collect_boundary_data()
        if sheet_refs is not None:
            sheet_refs["boundary_cache"] = {
                "pile_bottom_boundary": data["pile_bottom_boundary"],
                "corbel_long_num": data["corbel_long_num"],
                "corbel_short_num": data["corbel_short_num"],
                "connections": data["connections"],
                "form_dict": data["cofferdam_elasticlink_form_dict"],
                "stiffness_dict": data["sheet_elasticlink_dict"],
            }

    if sheet_refs is not None:
        sheet_refs["refresh_boundary"] = _refresh_boundary_cache
        sheet_refs["_collect_boundary_data"] = _collect_boundary_data

    
def _load_soil_into_frame(soil_frame, load_dict):
    """将 Load_dict['Excel'] 中的土层数据载入 SoilSettingsFrame 表格。

    Args:
        soil_frame (SoilSettingsFrame): 土层表控件，提供 load_data 方法。
        load_dict (dict): 土层信息字典，结构形式
            {'Excel': {层名: {'土层名称','层厚','重度','黏聚力','内摩擦角',
                              '计算方法','渗透系数(m/d)','土层顶承压水头(m)'}}}。

    Returns:
        None
    """
    excel_dict = load_dict.get('Excel', {})
    field_keys = ["层厚", "重度", "黏聚力", "内摩擦角",
                  "计算方法", "渗透系数(m/d)", "土层顶承压水头(m)"]
    data = {}
    for layer in excel_dict.values():
        name = layer.get('土层名称', '')
        if not name:
            continue
        vals = [str(layer.get(k, "")) for k in field_keys]
        data[name] = vals
    soil_frame.load_data(data) if data else soil_frame.load_data(None)


# ===== 页面5: 荷载参数 =====
def LoadSetting_GUI(table, sigma_k_dict, sheet_refs=None, diagram=None):
    """构建"荷载参数"页（Tab 6）：超载（附加荷载）参数表与加载示意图。

    支持均布/矩形局部/条形局部三类荷载的四面独立配置，参数表编辑实时同步
    到 sigma_k_dict 并刷新加载示意图；注册 surcharge_cache 刷新。

    Args:
        table (tk.Widget): 页面容器。
        sigma_k_dict (dict): 附加荷载信息字典，就地初始化与修改。
        sheet_refs (dict|None): 跨 Tab 共享引用字典。
        diagram (RealTimeDiagram|None): 主示意图，供加载示意图复用轮廓。

    Returns:
        None
    """

    TYPE_LABELS = ["均布附加荷载", "矩形局部附加荷载", "条形局部附加荷载"]
    FIELD_KEYS = ["p", "angle", "b", "a", "d", "l", "p2", "c"]
    COL_HEADERS = ["面", "是否加载", "p(kPa)", "θ(°)", "b(m)", "a(m)", "d(m)", "l(m)", "p2(kPa)", "c(m)"]
    COL_WEIGHTS = [2, 2, 2, 1, 2, 2, 2, 2, 2, 2]
    FACE_NAMES = ["①", "②", "③", "④"]
    

    def _ensure_init():
        """确保 sigma_k_dict 具备三类荷载的完整初始结构（缺省时补齐）。"""
        if "active_type" not in sigma_k_dict:
            sigma_k_dict["active_type"] = "均布附加荷载"
        if "均布附加荷载" not in sigma_k_dict:
            sigma_k_dict["均布附加荷载"] = {str(i): {"q0": "/"} for i in range(1, 5)}
        if "矩形局部附加荷载" not in sigma_k_dict:
            sigma_k_dict["矩形局部附加荷载"] = {str(i): {"angle": "45", "p0": "/", "b": "/", "a": "/", "d": "/", "l": "/", "p2": "/", "c": "/"} for i in range(1, 5)}
        if "条形局部附加荷载" not in sigma_k_dict:
            sigma_k_dict["条形局部附加荷载"] = {str(i): {"angle": "45", "p0": "/", "b": "/", "a": "/", "d": "/"} for i in range(1, 5)}


    # ==========================================
    # UI 构建 
    # ==========================================
    # 1. 最外层 table 权重分配
    table.grid_rowconfigure(0, weight=7) # Row 0 (超载与图示) 拿走所有拉伸空间
    table.grid_rowconfigure(1, weight=3) # Row 1 (环境参数) 保持最小所需高度
    table.grid_columnconfigure(0, weight=1)

    # ══════════════════════════════════════════
    # A. 超载参数区
    # ══════════════════════════════════════════
    surcharge_lf = tb.LabelFrame(table, text="超载参数", bootstyle=PRIMARY, padding=5)
    surcharge_lf.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10, 5))

    # --- 1. 荷载类型按钮区 (固定高度) ---
    type_toolbar = tb.Frame(surcharge_lf)
    type_toolbar.pack(side=TOP, fill=X, pady=10, padx=10)
    type_buttons = []
    
    # --- 2. 表格区 (固定高度) ---
    container = tb.Frame(surcharge_lf)
    container.pack(side=TOP, fill=X, padx=10, pady=5)

    canvas = tk.Canvas(container, borderwidth=0, highlightthickness=0, bg="white", height=200)
    canvas.pack(side=LEFT, fill=X, expand=True) 
    inner = tk.Frame(canvas, bg="white")
    inner_id = canvas.create_window((0, 0), window=inner, anchor="nw")
    canvas.bind("<Configure>", lambda e: canvas.itemconfig(inner_id, width=e.width))

    # 表头渲染
    header_bgs = ["#e8f4fd", "#e8f4fd"] + ["#f2f2f2"] * 8
    for col_idx, (hdr, w) in enumerate(zip(COL_HEADERS, COL_WEIGHTS)):
        tk.Label(inner, text=hdr, font=("Microsoft YaHei UI", 9, "bold"),
                 bg=header_bgs[col_idx], relief="solid", borderwidth=1, padx=10, pady=6)\
            .grid(row=0, column=col_idx, sticky="nsew")
        inner.grid_columnconfigure(col_idx, weight=w, uniform="sc_cols")

    # 数据行渲染
    check_var_list = []
    _handling_check = [False]  
    _highlight_key = [None]
    entry_matrix = []
    
    for face_idx in range(4):
        widgets_in_row = []
        
        # 面 Label
        lbl = tk.Label(inner, text=FACE_NAMES[face_idx], font=("Microsoft YaHei UI", 9, "bold"),
                       bg="white", relief="solid", borderwidth=1, padx=10, pady=8)
        lbl.grid(row=face_idx + 1, column=0, sticky="nsew")
        widgets_in_row.append(lbl)

        # Checkbutton
        chk_frame = tk.Frame(inner, relief="solid", borderwidth=1, bg="white", padx=5, pady=5)
        chk_frame.grid(row=face_idx + 1, column=1, sticky="nsew")
        bool_var = tk.BooleanVar(value=False)
        check_var_list.append(bool_var)
        chk = tb.Checkbutton(chk_frame, text="", variable=bool_var, bootstyle="round-toggle", 
                             command=lambda fidx=face_idx: handle_checkbox(fidx))
        chk.pack(expand=True)
        widgets_in_row.append(chk_frame)

        # 参数 Entries
        row_entries = []
        for col_idx in range(8):
            e = tk.Entry(inner, font=("Microsoft YaHei UI", 9), relief="solid", borderwidth=1, justify="center", bd=1)
            e.grid(row=face_idx + 1, column=col_idx + 2, sticky="nsew", padx=0, pady=0)
            e.bind("<FocusOut>", lambda e, fidx=face_idx, cidx=col_idx: handle_entry_edit(fidx, cidx))
            e.bind("<FocusIn>", lambda e, fidx=face_idx, cidx=col_idx: _on_focus_entry(fidx, cidx))
            row_entries.append(e)
            widgets_in_row.append(e)
            
        entry_matrix.append(row_entries)

    # --- 3. 底部绘图区 ---
    plot_frame = tb.Frame(surcharge_lf)
    plot_frame.pack(side=TOP, fill=BOTH, expand=True, padx=5, pady=(0, 5)) 
    
    plot_inner = tb.Frame(plot_frame)
    plot_inner.pack(fill=BOTH, expand=True, padx=10, pady=(0, 10))
    surcharge_diagram = SurchargeDiagram(plot_inner)


    # ══════════════════════════════════════════
    # B. 环境参数区
    # ══════════════════════════════════════════
    env_lf = tb.LabelFrame(table, text="环境参数", bootstyle=PRIMARY, padding=5)
    env_lf.grid(row=1, column=0, sticky="ew", padx=10, pady=(10, 5))

    env_btn_frame = tb.Frame(env_lf)
    env_btn_frame.pack(fill=X, padx=5, pady=(5, 0))
    env_var = tb.StringVar(value="水流力")
    env_btns = []
    for label in ["水流力", "波浪力"]:
        btn = tb.Button(env_btn_frame, text=label, bootstyle=(PRIMARY, OUTLINE),
                        command=lambda v=label: _switch_env(v))
        btn.pack(side=LEFT, padx=2, fill=X, expand=True)
        env_btns.append(btn)

    env_content = tb.Frame(env_lf)
    env_content.pack(fill=BOTH, expand=True, padx=5, pady=5)
    env_placeholder = tb.Label(env_content, text="其余荷载参数设置待开发",
                               font=("Microsoft YaHei UI", 10), bootstyle="secondary")
    env_placeholder.pack(expand=True)

    def _switch_env(selected):
        """（初版）切换环境荷载类型并更新按钮高亮。"""
        env_var.set(selected)
        for btn in env_btns:
            btn.config(bootstyle=(PRIMARY, OUTLINE) if btn.cget("text") != selected else PRIMARY)

    _switch_env("水流力")
    # ==========================================
    # 3. 业务与渲染逻辑
    # ==========================================
    def switch_type(type_index):
        """切换当前编辑的附加荷载类型并刷新表格与示意图。

        Args:
            type_index (int): 0/1/2 分别对应均布/矩形局部/条形局部。

        Returns:
            None
        """
        _ensure_init()
        sigma_k_dict["active_type"] = TYPE_LABELS[type_index]
        for i, btn in enumerate(type_buttons):
            btn.config(bootstyle=(PRIMARY, OUTLINE) if i != type_index else PRIMARY)
        refresh_ui()

    for i, label_text in enumerate(TYPE_LABELS):
        btn = tb.Button(type_toolbar, text=label_text, bootstyle=(PRIMARY, OUTLINE), style="LoadType.TButton",
                        command=lambda idx=i: switch_type(idx))
        btn.pack(side=LEFT, padx=2, fill=X, expand=True)
        type_buttons.append(btn)


    def _on_focus_entry(fidx, cidx):
        """输入框聚焦时设置示意图高亮项并重绘（d/angle 不高亮）。

        Args:
            fidx (int): 面索引（0~3）。
            cidx (int): 列索引（对应 FIELD_KEYS）。

        Returns:
            None
        """
        FIELD_KEYS = ["p", "angle", "b", "a", "d", "l", "p2", "c"]
        # p/p2 高亮文字标注，b/a/l/c 高亮尺寸标注，d/angle 不高亮
        if FIELD_KEYS[cidx] in ("d", "angle"): return
        _highlight_key[0] = (fidx+1, FIELD_KEYS[cidx])
        refresh_ui()
    
    def refresh_ui():
        """按当前荷载类型把 sigma_k_dict 数据回填到表格，并刷新加载示意图。

        同时根据荷载类型设置各列输入框的启用/只读/禁用状态。

        Returns:
            None
        """
        _ensure_init()
        active_type = sigma_k_dict["active_type"]
        
        for face_idx in range(4):
            face_str = str(face_idx + 1)
            f_data = sigma_k_dict[active_type][face_str]
            
            p_key = 'q0' if active_type == "均布附加荷载" else 'p0'
            p_val = str(f_data.get(p_key, '/')).strip()
            is_active = (p_val not in ('/', '', 'None'))
            check_var_list[face_idx].set(is_active)
            
            for col_idx, key in enumerate(FIELD_KEYS):
                entry = entry_matrix[face_idx][col_idx]
                entry.config(state="normal")
                entry.delete(0, tk.END)
                
                display_val = '/'
                if is_active:
                    if key == 'p': 
                        display_val = f_data.get(p_key, '/')
                    elif active_type == "均布附加荷载":
                        display_val = '/' 
                    elif active_type == "条形局部附加荷载" and key == 'l':
                        display_val = '/' 
                    else:
                        display_val = f_data.get(key, '/')
                        
                entry.insert(0, str(display_val))
                
                if not is_active:
                    entry.config(state="disabled", bg="#f0f0f0")
                else:
                    if active_type == "均布附加荷载" and key != 'p':
                        entry.config(state="readonly", bg="#fcfcfc")
                    elif active_type == "条形局部附加荷载" and key in ('l', 'p2', 'c'):
                        entry.config(state="readonly", bg="#fcfcfc")
                    elif active_type == "矩形局部附加荷载" and key in ('p2', 'c'):
                        entry.config(state="normal", bg="white")
                    else:
                        entry.config(state="normal", bg="white")
                        
        surcharge_diagram.draw(sigma_k_dict, rtd=diagram, highlight_key=_highlight_key[0])

    def handle_checkbox(face_idx):
        """勾选/取消某面荷载：启用时写入默认参数，取消时重置为 "/"。

        Args:
            face_idx (int): 面索引（0~3）。

        Returns:
            None
        """
        if _handling_check[0]: return
        _handling_check[0] = True
        _ensure_init()
        active_type = sigma_k_dict["active_type"]
        face_str = str(face_idx + 1)
        is_checked = check_var_list[face_idx].get()
        p_key = "q0" if active_type == "均布附加荷载" else "p0"
        if is_checked:
            sigma_k_dict[active_type][face_str][p_key] = "20"
            if active_type != "均布附加荷载":
                sigma_k_dict[active_type][face_str].update({"angle": "45", "b": "1", "a": "1", "d": "0"})
                if active_type == "矩形局部附加荷载":
                    sigma_k_dict[active_type][face_str].update({"l": "5", "p2": "/", "c": "/"})
        else:
            sigma_k_dict[active_type][face_str][p_key] = "/"
            if active_type != "均布附加荷载":
                base = {"angle": "/", "b": "/", "a": "/", "d": "/", "l": "/"}
                if active_type == "矩形局部附加荷载":
                    base.update({"p2": "/", "c": "/"})
                sigma_k_dict[active_type][face_str].update(base)
        refresh_ui()
        _handling_check[0] = False

    def handle_entry_edit(face_idx, col_idx):
        """输入框失焦时把值写回 sigma_k_dict 并刷新示意图。

        Args:
            face_idx (int): 面索引（0~3）。
            col_idx (int): 列索引（对应 FIELD_KEYS）。

        Returns:
            None
        """
        _ensure_init()
        active_type = sigma_k_dict["active_type"]
        face_str = str(face_idx + 1)
        key = FIELD_KEYS[col_idx]
        new_val = entry_matrix[face_idx][col_idx].get().strip()

        # 各类型不可编辑的字段
        if active_type == "均布附加荷载" and key != 'p': return
        if active_type == "条形局部附加荷载" and key in ('l', 'p2', 'c'): return

        dict_key = ('q0' if active_type == "均布附加荷载" else 'p0') if key == 'p' else key
        if not new_val: new_val = '/'
        sigma_k_dict[active_type][face_str][dict_key] = new_val
        _highlight_key[0] = None
        surcharge_diagram.draw(sigma_k_dict, rtd=diagram, highlight_key=_highlight_key[0])
        
    if sheet_refs is not None:
        sheet_refs["sigma_k_dict"] = sigma_k_dict
        sheet_refs["surcharge_diagram"] = surcharge_diagram
        sheet_refs["_surcharge_refresh"] = refresh_ui

    # ---- surcharge cache ----
    def _refresh_surcharge_cache():
        """把当前附加荷载数据快照写入 sheet_refs['surcharge_cache']。"""
        if sheet_refs is not None:
            sheet_refs["surcharge_cache"] = dict(sigma_k_dict)
    if sheet_refs is not None:
        sheet_refs["refresh_surcharge"] = _refresh_surcharge_cache

    # 首次渲染触发
    switch_type(0)
    def _switch_env(selected):
        """（最终版）切换环境荷载类型并更新按钮高亮，预留内容切换逻辑。"""
        env_var.set(selected)
        for btn in env_btns:
            btn.config(bootstyle=(PRIMARY, OUTLINE) if btn.cget("text") != selected else PRIMARY)
        # 预留：根据 selected 切换 env_content 内容

    _switch_env("水流力")  # 初始状态


# ===== 页面5: 工况分析 =====

def _parse_assist_bracket_data(raw_str):
    """解析辅助换撑数据的括号分隔字符串为字典列表。

    格式：(name:connected,mat_dc,dc_sec,mat_xc,xc_sec,dc_x,dc_y,xc_x,xc_y,
    proj_long,proj_short,diff_long,diff_short),...

    Args:
        raw_str (str): 括号分隔字符串；空或 "/" 表示无数据。

    Returns:
        list[dict]: 每项含 name/connected 及 12 个布置字段（缺失补 "/"）。
    """
    if not raw_str or raw_str.strip() in ("", "/"):
        return []
    result = []
    # 用正则匹配每个 (...) 块
    import re
    blocks = re.findall(r'\(([^)]+)\)', raw_str)
    for block in blocks:
        parts = block.split(",")
        if len(parts) < 2:
            continue
        # 第一项是 "name:connected"
        name_conn = parts[0].split(":", 1)
        name = name_conn[0].strip()
        connected = name_conn[1].strip() if len(name_conn) > 1 else "否"
        # 剩余字段
        fields = [p.strip() for p in parts[1:]]
        while len(fields) < 12:
            fields.append("/")
        result.append({
            "name": name,
            "connected": connected,
            "mat_dc": fields[0], "dc_sec": fields[1],
            "mat_xc": fields[2], "xc_sec": fields[3],
            "dc_x": fields[4], "dc_y": fields[5],
            "xc_x": fields[6], "xc_y": fields[7],
            "proj_long": fields[8], "proj_short": fields[9],
            "diff_long": fields[10], "diff_short": fields[11],
        })
    return result


def _serialize_assist_bracket_data(data_list):
    """将辅助换撑字典列表序列化为括号分隔字符串（_parse 的逆操作）。

    Args:
        data_list (list[dict]): 每项含 name/connected 及 12 个布置字段。

    Returns:
        str: 形如 "(name:connected,字段...),..." 的字符串；无数据返回 "/"。
    """
    if not data_list:
        return "/"
    parts = []
    for item in data_list:
        name = item.get("name", "")
        connected = item.get("connected", "否")
        fields = [
            item.get("mat_dc", "/"), item.get("dc_sec", "/"),
            item.get("mat_xc", "/"), item.get("xc_sec", "/"),
            item.get("dc_x", "/"), item.get("dc_y", "/"),
            item.get("xc_x", "/"), item.get("xc_y", "/"),
            item.get("proj_long", "/"), item.get("proj_short", "/"),
            item.get("diff_long", "/"), item.get("diff_short", "/"),
        ]
        parts.append(f"({name}:{connected},{','.join(fields)})")
    return ",".join(parts)


def _parse_assist_add_data(waler_raw, strut_raw=None):
    """解析辅助加撑数据：分别解析围檩列与内支撑列后按 layer_no 合并。

    Args:
        waler_raw (str): 辅助加围檩列字符串，格式
            (layer_no:elevation,mat_long,sec_long,mat_short,sec_short),...
        strut_raw (str|None): 辅助加内支撑列字符串，格式
            (layer_no:connected,mat_dc,dc_sec,mat_xc,xc_sec,dc_x,dc_y,xc_x,xc_y,
             proj_long,proj_short,diff_long,diff_short),...

    Returns:
        list[dict]: 每项含 layer_no 及 18 个完整字段（缺失补 "/"）。
    """
    import re
    # 解析围檩列
    waler_map = {}  # {layer_no: {elevation, mat_long, sec_long, mat_short, sec_short}}
    if waler_raw and waler_raw.strip() not in ("", "/"):
        blocks = re.findall(r'\(([^)]+)\)', waler_raw)
        for block in blocks:
            parts = [p.strip() for p in block.split(",")]
            if len(parts) < 1:
                continue
            layer_elev = parts[0].split(":", 1)
            layer_no = layer_elev[0].strip()
            elevation = layer_elev[1].strip() if len(layer_elev) > 1 else "/"
            mat_long = parts[1] if len(parts) > 1 else "/"
            sec_long = parts[2] if len(parts) > 2 else "/"
            mat_short = parts[3] if len(parts) > 3 else "/"
            sec_short = parts[4] if len(parts) > 4 else "/"
            waler_map[layer_no] = {
                "elevation": elevation,
                "mat_long": mat_long,
                "sec_long": sec_long,
                "mat_short": mat_short,
                "sec_short": sec_short,
            }
    # 解析内支撑列
    strut_map = {}  # {layer_no: {connected, mat_dc, dc_sec, ...}}
    if strut_raw and strut_raw.strip() not in ("", "/"):
        blocks = re.findall(r'\(([^)]+)\)', strut_raw)
        for block in blocks:
            parts = [p.strip() for p in block.split(",")]
            if len(parts) < 1:
                continue
            layer_conn = parts[0].split(":", 1)
            layer_no = layer_conn[0].strip()
            connected = layer_conn[1].strip() if len(layer_conn) > 1 else "否"
            fields = parts[1:]
            while len(fields) < 12:
                fields.append("/")
            strut_map[layer_no] = {
                "connected": connected,
                "mat_dc": fields[0], "dc_sec": fields[1],
                "mat_xc": fields[2], "xc_sec": fields[3],
                "dc_x": fields[4], "dc_y": fields[5],
                "xc_x": fields[6], "xc_y": fields[7],
                "proj_long": fields[8], "proj_short": fields[9],
                "diff_long": fields[10], "diff_short": fields[11],
            }
    # 按围檩列的 layer_no 顺序合并（围檩列为主）
    result = []
    for layer_no in waler_map:
        w = waler_map[layer_no]
        s = strut_map.get(layer_no, {})
        result.append({
            "layer_no": layer_no,
            "elevation": w.get("elevation", "/"),
            "mat_long": w.get("mat_long", "/"),
            "sec_long": w.get("sec_long", "/"),
            "mat_short": w.get("mat_short", "/"),
            "sec_short": w.get("sec_short", "/"),
            "connected": s.get("connected", "否"),
            "mat_dc": s.get("mat_dc", "/"),
            "dc_sec": s.get("dc_sec", "/"),
            "mat_xc": s.get("mat_xc", "/"),
            "xc_sec": s.get("xc_sec", "/"),
            "dc_x": s.get("dc_x", "/"),
            "dc_y": s.get("dc_y", "/"),
            "xc_x": s.get("xc_x", "/"),
            "xc_y": s.get("xc_y", "/"),
            "proj_long": s.get("proj_long", "/"),
            "proj_short": s.get("proj_short", "/"),
            "diff_long": s.get("diff_long", "/"),
            "diff_short": s.get("diff_short", "/"),
        })
    # 如果内支撑列有围檩列中不存在的 layer_no，也追加（防御性）
    for layer_no in strut_map:
        if layer_no not in waler_map:
            s = strut_map[layer_no]
            result.append({
                "layer_no": layer_no,
                "elevation": "/",
                "mat_long": "/", "sec_long": "/",
                "mat_short": "/", "sec_short": "/",
                "connected": s.get("connected", "否"),
                "mat_dc": s.get("mat_dc", "/"),
                "dc_sec": s.get("dc_sec", "/"),
                "mat_xc": s.get("mat_xc", "/"),
                "xc_sec": s.get("xc_sec", "/"),
                "dc_x": s.get("dc_x", "/"),
                "dc_y": s.get("dc_y", "/"),
                "xc_x": s.get("xc_x", "/"),
                "xc_y": s.get("xc_y", "/"),
                "proj_long": s.get("proj_long", "/"),
                "proj_short": s.get("proj_short", "/"),
                "diff_long": s.get("diff_long", "/"),
                "diff_short": s.get("diff_short", "/"),
            })
    return result


def _serialize_assist_add_data(data_list):
    """序列化辅助加撑数据（兼容旧调用，实际输出围檩部分）。

    Args:
        data_list (list[dict]): 辅助加撑数据列表。

    Returns:
        str: 围檩部分括号分隔字符串；无数据返回 "/"。
    """
    if not data_list:
        return "/"
    waler_str = _serialize_assist_add_waler(data_list)
    return waler_str


def _serialize_assist_add_waler(data_list):
    """序列化辅助加撑的围檩部分。

    格式：(layer_no:elevation,mat_long,sec_long,mat_short,sec_short),...

    Args:
        data_list (list[dict]): 辅助加撑数据列表。

    Returns:
        str: 围檩部分括号分隔字符串；无数据返回 "/"。
    """
    if not data_list:
        return "/"
    parts = []
    for item in data_list:
        layer_no = item.get("layer_no", "")
        elevation = item.get("elevation", "/")
        mat_long = item.get("mat_long", "/")
        sec_long = item.get("sec_long", "/")
        mat_short = item.get("mat_short", "/")
        sec_short = item.get("sec_short", "/")
        parts.append(f"({layer_no}:{elevation},{mat_long},{sec_long},{mat_short},{sec_short})")
    return ",".join(parts)


def _serialize_assist_add_strut(data_list):
    """序列化辅助加撑的内支撑部分（连接、材质、截面、承台投影等）。

    格式：(layer_no:connected,mat_dc,dc_sec,mat_xc,xc_sec,dc_x,dc_y,xc_x,xc_y,
    proj_long,proj_short,diff_long,diff_short),...

    Args:
        data_list (list[dict]): 辅助加撑数据列表。

    Returns:
        str: 内支撑部分括号分隔字符串；无数据返回 "/"。
    """
    if not data_list:
        return "/"
    parts = []
    for item in data_list:
        layer_no = item.get("layer_no", "")
        connected = item.get("connected", "否")
        fields = [
            item.get("mat_dc", "/"), item.get("dc_sec", "/"),
            item.get("mat_xc", "/"), item.get("xc_sec", "/"),
            item.get("dc_x", "/"), item.get("dc_y", "/"),
            item.get("xc_x", "/"), item.get("xc_y", "/"),
            item.get("proj_long", "/"), item.get("proj_short", "/"),
            item.get("diff_long", "/"), item.get("diff_short", "/"),
        ]
        parts.append(f"({layer_no}:{connected},{','.join(fields)})")
    return ",".join(parts)

def Construct_Stage(table, line_data, rows, Drawdown_height,
                            Concrete_Plug_check_var, Solid_Level, Water_Level,
                            Concrete_Blinding_check_var=None,
                            sheet_refs=None):
    """构建"工况分析"页的施工阶段定义表格（Canvas 多列表格）。

    每行代表一个施工阶段，可按工况类型动态启用/禁用列、切换支撑层数下拉、
    联动施工阶段图示，并把结果同步为 stage_dict 与 line_data。

    Args:
        table (tk.Widget): 页面容器。
        line_data (list[list[str]]): 旧格式阶段数据（后向兼容，就地更新）。
        rows (list[dict]): 围檩行集合，用于确定支撑层号。
        Drawdown_height (tk.Variable): 超挖深度（m）。
        Concrete_Plug_check_var, Concrete_Blinding_check_var (tk.Variable):
            封底/垫层开关，决定可用工况类型。
        Solid_Level, Water_Level (tk.Variable): 土顶/水面标高（m）。
        sheet_refs (dict|None): 跨 Tab 共享引用字典。

    工况类型: 取土, 加撑, 抽水, 封底, 垫层, 辅助加圈梁, 拆撑, 辅助换撑, 辅助加撑

    Returns:
        None
    """
    HEADER_BG = "#1e3a5f"
    HEADER_FG = "#ffffff"
    BORDER_COLOR = "#c8d2de"
    DISABLED_BG = "#e9ecef"

    STAGE_TYPES = [
        "取土", "加撑", "抽水", "封底", "垫层",
        "辅助加圈梁", "拆撑", "辅助换撑", "辅助加撑"
    ]
    HEADERS = ["序号", "各工况名称", "各工况基坑内土顶标高(m)", "各工况基坑内水面标高(m)", "各工况支撑变动"]

    # ==================== 布局（全部 grid，图示固定底部，表格 weight=1 撑满）====================
    table.grid_rowconfigure(0, weight=0)  # 标题
    table.grid_rowconfigure(1, weight=0)  # 按钮行
    table.grid_rowconfigure(2, weight=1)  # 表格
    table.grid_rowconfigure(3, weight=0)  # 图示（固定）
    table.grid_columnconfigure(0, weight=1)

    # 标题 (row 0)
    stage_label = tb.Frame(table)
    stage_label.grid(row=0, column=0, sticky="ew", pady=5)
    tb.Label(stage_label, text='施工阶段定义', bootstyle=PRIMARY,
              font=("Microsoft YaHei UI", 10, "bold")).pack(side=LEFT, padx=15, pady=5)
    # 按钮控件行 (row 1)
    container1 = tb.Frame(table)
    container1.grid(row=1, column=0, sticky="ew", padx=15, pady=5)
    # 辅助换撑/辅助加撑缓存
    if sheet_refs is not None and "assist_replace_data" not in sheet_refs:
        sheet_refs["assist_replace_data"] = []
    if sheet_refs is not None and "assist_add_data" not in sheet_refs:
        sheet_refs["assist_add_data"] = []
    # 辅助按钮（右：辅助加撑 / 辅助换撑）
    btn_assist_add = tb.Button(container1, text='辅助加撑', bootstyle=(PRIMARY,OUTLINE), width=11)
    btn_assist_replace = tb.Button(container1, text='辅助换撑', bootstyle=(PRIMARY,OUTLINE), width=11)
    btn_assist_replace.pack(side=RIGHT, padx=(0, 15), pady=5)
    btn_assist_add.pack(side=RIGHT, padx=5, pady=5)

    def _get_steel_mat_options():
        """从 sheet_refs 取钢材牌号选项提供函数并返回其结果。

        Returns:
            list[str]: 钢材牌号列表；无提供者时返回 []。
        """
        fn = sheet_refs.get("_get_steel_material_options") if sheet_refs else None
        return list(fn()) if fn else []

    # ==================== 内部数据（stage_dict 格式）====================
    # {1: {'工况类型':'取土', '基坑内土顶标高':'/', '基坑内水面标高':'/', '支撑层数':'/'}, ...}
    _stage_dict = {}

    def _convert_old_line_data():
        """将 line_data（旧格式 list[list[str]]）转为 stage_dict"""
        nonlocal _stage_dict
        _stage_dict = {}
        for i, row in enumerate(line_data):
            if not row or not row[0]:
                continue
            name = str(row[0]).strip()
            if name in ("", "请点击自动生成按钮"):
                continue
            stype = name
            support = "/"
            if "安装" in name and "围檩" in name:
                stype = "加撑"
                for p in name.split("第"):
                    for p2 in p.split("层"):
                        try: support = str(int(float(p2.strip()))); break
                        except: pass
            elif "开挖" in name:
                stype = "取土"
            elif "封底" in name:
                stype = "封底"
            elif "抽水" in name:
                stype = "抽水"
            _stage_dict[i + 1] = {
                "工况类型": stype,
                "基坑内土顶标高": "/",
                "基坑内水面标高": "/",
                "支撑层数": support,
            }
    _convert_old_line_data()

    # ==================== 表格区域（Canvas）====================
    border_frame = tk.Frame(table, bg=BORDER_COLOR)
    border_frame.grid(row=2, column=0, sticky="nsew", padx=15, pady=5)

    table_frame = tb.Frame(border_frame)
    table_frame.pack(fill=BOTH, expand=True, padx=1, pady=1)

    v_scroll = tb.Scrollbar(table_frame, orient="vertical")
    v_scroll.pack(side=RIGHT, fill=Y)

    canvas = tk.Canvas(table_frame, borderwidth=0, highlightthickness=0,
                       yscrollcommand=v_scroll.set)
    canvas.pack(side=LEFT, fill=BOTH, expand=True)
    v_scroll.config(command=canvas.yview)

    inner = tk.Frame(canvas)
    inner_id = canvas.create_window((0, 0), window=inner, anchor="nw")
    inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    # 让 inner 始终撑满 canvas 宽度（列比例生效的前提）
    def _on_canvas_resize(event):
        """画布尺寸变化时让内容撑满宽度，使列比例生效。"""
        canvas.itemconfig(inner_id, width=event.width)
    canvas.bind("<Configure>", _on_canvas_resize)

    def _on_mousewheel(event):
        """内容高于视口时允许滚轮纵向滚动。"""
        if inner.winfo_height() > canvas.winfo_height():
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    def _on_enter(_):
        """鼠标进入表格区时绑定全局滚轮事件。"""
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
    def _on_leave(_):
        """鼠标离开表格区时解绑全局滚轮事件。"""
        canvas.unbind_all("<MouseWheel>")
    canvas.bind("<Enter>", _on_enter)
    canvas.bind("<Leave>", _on_leave)

    # 表头（按比例分配列宽，minsize 保证最小宽度）
    STAGE_COL_WEIGHTS = [2, 2, 3, 3, 2]
    STAGE_COL_MINSIZE = [40, 80, 130, 130, 80]
    for ci, h in enumerate(HEADERS):
        lbl = tk.Label(inner, text=h, font=("Microsoft YaHei UI", 9, "bold"),
                       bg=HEADER_BG, fg=HEADER_FG, relief="solid", borderwidth=1,
                       padx=5, pady=5, anchor="center")
        lbl.grid(row=0, column=ci, sticky="nsew")
        inner.grid_columnconfigure(ci, weight=STAGE_COL_WEIGHTS[ci],
                                   minsize=STAGE_COL_MINSIZE[ci])

    # ==================== 工况类型可用性 ====================

    def _get_available_types():
        """按当前水位/封底/垫层/辅助数据返回可用的工况类型列表。

        Returns:
            list[str]: 工况类型列表（如 取土/加撑/抽水/封底/垫层/...）。
        """
        types = ["取土", "加撑"]
        # 水比土高时允许抽水工况
        try:
            if float(Water_Level.get()) > float(Solid_Level.get()):
                types.append("抽水")
        except: pass
        if Concrete_Plug_check_var and Concrete_Plug_check_var.get():
            types.append("封底")
        if Concrete_Blinding_check_var and Concrete_Blinding_check_var.get():
            types.append("垫层")
        types.append("辅助加圈梁")
        types.append("拆撑")
        if sheet_refs and sheet_refs.get("assist_replace_data"):
            types.append("辅助换撑")
        if sheet_refs and sheet_refs.get("assist_add_data"):
            types.append("辅助加撑")
        return types

    # ==================== Per-cell 规则 ====================
    def _col1_enabled(stype):
        """判断该工况类型下"基坑内土顶标高"列是否可编辑。"""
        return stype in ("取土", "辅助加圈梁", "抽水")

    def _col2_enabled(stype):
        """判断该工况类型下"基坑内水面标高"列是否可编辑。"""
        return stype in ("取土", "抽水")

    def _get_support_options(stype):
        """按工况类型返回"支撑层数"列的可选值列表。

        Args:
            stype (str): 工况类型。

        Returns:
            list[str]: 支撑层号/换撑名/辅助加撑层号等选项。
        """
        if stype in ("加撑", "拆撑"):
            cnt = len(rows) if rows else 0
            opts = [str(i + 1) for i in range(cnt)]
            # 拆撑时追加辅助加撑定义的层号
            if stype == "拆撑" and sheet_refs:
                for d in sheet_refs.get("assist_add_data", []):
                    ln = d.get("layer_no", "")
                    if ln and ln not in opts:
                        opts.append(ln)
            return opts
        if stype == "辅助换撑" and sheet_refs:
            return [d.get("name", "") for d in sheet_refs.get("assist_replace_data", [])]
        if stype == "辅助加撑" and sheet_refs:
            return [d.get("layer_no", "") for d in sheet_refs.get("assist_add_data", [])]
        return []

    def _support_is_combo(stype):
        """判断该工况类型的"支撑层数"列是否使用下拉框。"""
        return stype in ("加撑", "拆撑", "辅助换撑", "辅助加撑")

    # ==================== 行高亮 ====================
    _selected_row = -1

    def select_row(idx):
        """选中某行：浅蓝高亮并联动刷新施工阶段图示。

        Args:
            idx (int): 行索引（0 基）。

        Returns:
            None
        """
        nonlocal _selected_row
        _selected_row = idx
        row_count = len(_row_widgets)
        for r in range(1, row_count + 1):
            bg = "#e3f2fd" if r - 1 == idx else "white"
            for w in inner.grid_slaves(row=r):
                if isinstance(w, (tk.Label, tk.Frame)):
                    try: w.config(bg=bg)
                    except: pass
        # 联动施工阶段图示
        if sheet_refs is not None:
            sd = sheet_refs.get("stage_diagram")
            if sd is not None and hasattr(sd, "draw"):
                try: sd.draw(idx)
                except: pass

    # ==================== 行控件 ====================
    _row_widgets = []

    def _apply_row_states(idx):
        """按该行工况类型启用/禁用各列，并重建"支撑层数"控件。

        Args:
            idx (int): 行索引（0 基）。

        Returns:
            None
        """
        if idx >= len(_row_widgets):
            return
        rw = _row_widgets[idx]
        stype = rw["type_var"].get()

        # 土顶标高 (col 1)
        e1 = _col1_enabled(stype)
        if rw["entry_soil"]:
            if e1:
                rw["entry_soil"].configure(state="normal", bg="white")
            else:
                rw["entry_soil"].configure(state="disabled", bg=DISABLED_BG)
                rw["soil_var"].set("/")

        # 水面标高 (col 2)
        e2 = _col2_enabled(stype)
        if rw["entry_water"]:
            if e2:
                rw["entry_water"].configure(state="normal", bg="white")
            else:
                rw["entry_water"].configure(state="disabled", bg=DISABLED_BG)
                rw["water_var"].set("/")

        # 支撑层数 (col 4): 销毁重创建
        for w in inner.grid_slaves(row=idx + 1, column=4):
            w.destroy()
        if _support_is_combo(stype):
            opts = _get_support_options(stype)
            # 保留原有值（如"2"），仅当不在选项中时才回退为首项
            current_val = rw["support_var"].get()
            if current_val not in opts:
                current_val = opts[0] if opts else "/"
            rw["support_var"] = tk.StringVar(value=current_val)
            cb = tb.Combobox(inner, textvariable=rw["support_var"], values=opts,
                           state="readonly", width=8)
            cb.grid(row=idx + 1, column=4, sticky="nsew", padx=1, pady=1)
            cb.bind("<Button-1>", lambda e, i=idx: select_row(i), add="+")
            cb.bind("<FocusIn>", lambda e, i=idx: select_row(i), add="+")
            cb.bind("<<ComboboxSelected>>", lambda e, i=idx: _refresh_diagram(i), add="+")
            cb.bind("<FocusOut>", lambda e, i=idx: _refresh_diagram(i), add="+")
            cb.bind("<MouseWheel>", lambda e: "break", add="+")
        else:
            rw["support_var"] = tk.StringVar(value="/")
            e = tk.Entry(inner, textvariable=rw["support_var"], state="disabled",
                        font=("Microsoft YaHei UI", 9), justify="center",
                        relief="solid", borderwidth=1,
                        disabledbackground=DISABLED_BG, disabledforeground="#999999")
            e.grid(row=idx + 1, column=4, sticky="nsew", padx=1, pady=1)

    def _refresh_diagram(idx):
        """控件失焦后刷新施工阶段图示（无高亮联动）。

        Args:
            idx (int): 行索引（0 基）。

        Returns:
            None
        """
        if sheet_refs is not None:
            sd = sheet_refs.get("stage_diagram")
            if sd is not None and hasattr(sd, "draw"):
                try: sd.draw(idx)
                except: pass

    def _create_row_widgets(r_idx, data_dict):
        """创建一行的控件（序号/工况名称/土顶标高/水面标高/支撑层数）并绑定事件。

        Args:
            r_idx (int): 行索引（0 基）。
            data_dict (dict): 该行初始数据 {'工况类型','基坑内土顶标高',
                '基坑内水面标高','支撑层数'}。

        Returns:
            dict: 该行控件与变量集合。
        """
        row_num = r_idx + 1
        row = {}

        # 序号 (Label)
        lbl_idx = tk.Label(inner, text=str(row_num), font=("Microsoft YaHei UI", 9),
                           bg="white", relief="solid", borderwidth=1, anchor="center")
        lbl_idx.grid(row=row_num, column=0, sticky="nsew", padx=1, pady=1)
        lbl_idx.bind("<Button-1>", lambda e, idx=r_idx: select_row(idx))
        row["idx_label"] = lbl_idx

        # 工况类型 (Combobox)
        available = _get_available_types()
        type_var = tk.StringVar(value=data_dict.get("工况类型", "/"))
        cb_type = tb.Combobox(inner, textvariable=type_var, values=available,
                            state="readonly", width=12)
        cb_type.grid(row=row_num, column=1, sticky="nsew", padx=1, pady=1)
        cb_type.bind("<<ComboboxSelected>>",
                    lambda e, idx=r_idx: (_on_type_changed(idx), _refresh_diagram(idx)))
        cb_type.bind("<Button-1>", lambda e, idx=r_idx: select_row(idx), add="+")
        cb_type.bind("<FocusIn>", lambda e, idx=r_idx: select_row(idx), add="+")
        cb_type.bind("<FocusOut>", lambda e, idx=r_idx: _refresh_diagram(idx), add="+")
        cb_type.bind("<MouseWheel>", lambda e: "break", add="+")
        row["type_var"] = type_var

        def _fmt3(val):
            """将数值格式化为3位小数，非数值原样返回"""
            try: return f"{float(val):.3f}"
            except: return str(val)
        # 土顶标高 (Entry)
        soil_var = tk.StringVar(value=_fmt3(data_dict.get("基坑内土顶标高", "/")))
        e_soil = tk.Entry(inner, textvariable=soil_var,
                         font=("Microsoft YaHei UI", 9), justify="center",
                         relief="solid", borderwidth=1)
        e_soil.grid(row=row_num, column=2, sticky="nsew", padx=1, pady=1)
        e_soil.bind("<FocusIn>", lambda e, idx=r_idx: select_row(idx))
        e_soil.bind("<FocusOut>", lambda e, idx=r_idx: _refresh_diagram(idx))
        row["soil_var"] = soil_var
        row["entry_soil"] = e_soil

        # 水面标高 (Entry)
        water_var = tk.StringVar(value=_fmt3(data_dict.get("基坑内水面标高", "/")))
        e_water = tk.Entry(inner, textvariable=water_var,
                          font=("Microsoft YaHei UI", 9), justify="center",
                          relief="solid", borderwidth=1)
        e_water.grid(row=row_num, column=3, sticky="nsew", padx=1, pady=1)
        e_water.bind("<FocusIn>", lambda e, idx=r_idx: select_row(idx))
        e_water.bind("<FocusOut>", lambda e, idx=r_idx: _refresh_diagram(idx))
        row["water_var"] = water_var
        row["entry_water"] = e_water

        # 支撑层数 (占位)
        support_var = tk.StringVar(value=data_dict.get("支撑层数", "/"))
        row["support_var"] = support_var

        _row_widgets.append(row)
        _apply_row_states(r_idx)
        return row

    def _on_type_changed(idx):
        """工况类型变化时刷新该行列状态并同步 line_data。"""
        if idx >= len(_row_widgets):
            return
        _apply_row_states(idx)
        _update_line_data()

    def _destroy_all_rows():
        """销毁表格内所有数据行控件并清空行集合。"""
        for w in inner.grid_slaves():
            g = w.grid_info()
            if g and g.get("row") and int(g["row"]) > 0:
                w.destroy()
        _row_widgets.clear()

    def _rebuild_all_rows():
        """根据 _stage_dict 重建所有数据行并同步 line_data。"""
        _destroy_all_rows()
        for k in sorted(_stage_dict.keys()):
            _create_row_widgets(k - 1, _stage_dict[k])
        canvas.configure(scrollregion=canvas.bbox("all"))
        _update_line_data()

    # ==================== 数据同步 ====================
    def _build_stage_dict():
        """从当前表格控件值构建施工阶段字典。

        Returns:
            dict: {序号(int): {'工况类型','基坑内土顶标高','基坑内水面标高','支撑层数'}}。
        """
        result = {}
        for i, rw in enumerate(_row_widgets):
            stype = rw["type_var"].get()
            if not stype:
                continue
            result[i + 1] = {
                "工况类型": stype,
                "基坑内土顶标高": rw["soil_var"].get(),
                "基坑内水面标高": rw["water_var"].get(),
                "支撑层数": rw["support_var"].get(),
            }
        return result

    def _update_line_data():
        """把当前 stage_dict 同步到 line_data（后向兼容的旧格式）。"""
        sd = _build_stage_dict()
        line_data.clear()
        for k in sorted(sd.keys()):
            d = sd[k]
            line_data.append([d["工况类型"],
                             d.get("基坑内土顶标高", "/"),
                             d.get("基坑内水面标高", "/"),
                             d.get("支撑层数", "/")])
        nonlocal _stage_dict
        _stage_dict = sd

    # ==================== 加载 / 刷新 ====================
    _rebuild_all_rows()
    select_row(0)  # 默认高亮第一行

    def load_from_old_format(old_list):
        """从旧版 Cal_Stage 输出（工况名列表）加载并重建表格。

        Args:
            old_list (list[str]): 旧格式工况名称列表。

        Returns:
            None
        """
        new_dict = {}
        for i, name in enumerate(old_list):
            nm = name.strip()
            stype = nm
            support = "/"
            if "安装" in nm and "围檩" in nm:
                stype = "加撑"
                for p in nm.split("第"):
                    for p2 in p.split("层"):
                        try: support = str(int(float(p2.strip()))); break
                        except: pass
            elif "开挖" in nm:
                stype = "取土"
            elif "封底" in nm:
                stype = "封底"
            elif "抽水" in nm:
                stype = "抽水"
            new_dict[i + 1] = {
                "工况类型": stype,
                "基坑内土顶标高": "/",
                "基坑内水面标高": "/",
                "支撑层数": support,
            }
        nonlocal _stage_dict
        _stage_dict = new_dict
        _rebuild_all_rows()
        select_row(0)

    def load_from_stage_dict(new_dict):
        """从 stage_dict 加载并重建表格。

        Args:
            new_dict (dict): 施工阶段字典。

        Returns:
            None
        """
        nonlocal _stage_dict
        _stage_dict = dict(new_dict) if new_dict else {}
        _rebuild_all_rows()
        select_row(0)

    def get_stage_dict():
        """返回当前施工阶段字典（先同步表格数据）。

        Returns:
            dict: {序号: {'工况类型','基坑内土顶标高','基坑内水面标高','支撑层数'}}。
        """
        _update_line_data()
        return dict(_stage_dict)

    # ==================== 新增 / 删除行 ====================
    def _add_row():
        """在选中行之后插入一个新的空施工阶段并重建表格。"""
        nonlocal _stage_dict
        # 在选中行后插入，无选中则插末尾
        insert_pos = _selected_row + 1 if _selected_row >= 0 else len(_stage_dict)
        # 转列表插入
        items = [(k, v) for k, v in sorted(_stage_dict.items())]
        new_item = {"工况类型": "/", "基坑内土顶标高": "/",
                    "基坑内水面标高": "/", "支撑层数": "/"}
        items.insert(insert_pos, (0, new_item))  # key=0 临时占位，重编后被丢弃
        # 重编号
        _stage_dict.clear()
        for i, (_, v) in enumerate(items):
            _stage_dict[i + 1] = v
        _rebuild_all_rows()
        select_row(insert_pos)

    def _delete_row():
        """删除当前选中的施工阶段行并重建表格。"""
        nonlocal _selected_row
        if not _stage_dict or _selected_row < 0:
            return
        keys = sorted(_stage_dict.keys())
        del_key = keys[_selected_row]
        del _stage_dict[del_key]
        # 重编号 _stage_dict
        old = dict(_stage_dict)
        _stage_dict.clear()
        for i, k in enumerate(sorted(old.keys())):
            _stage_dict[i + 1] = old[k]
        # 完全重建表格（grid 行号自动对齐 _row_widgets）
        _rebuild_all_rows()
        _selected_row = min(_selected_row, len(_row_widgets) - 1)
        if _selected_row >= 0:
            select_row(_selected_row)

    # 操作按钮（左：新增行 / 删除选中行）
    tb.Button(container1, text="+ 新增行", bootstyle=SUCCESS, width=10,
              command=_add_row).pack(side=LEFT, padx=(0, 5), pady=5)
    tb.Button(container1, text="- 删除选中行", bootstyle=DANGER, width=12,
              command=_delete_row).pack(side=LEFT, padx=5, pady=5)

    # ==================== Assist Windows ====================
    # (辅助换撑 + 辅助加撑 — 完整保留原代码)
    # ---- 辅助换撑二级窗口 ----
    def _open_assist_replace_support():
        """打开"辅助换撑设置"二级窗口（增删行、校验并写回 assist_replace_data）。

        Returns:
            None
        """
        win = tb.Toplevel(table.winfo_toplevel())
        win.title("辅助换撑设置")
        win.geometry("1050x420+%d+%d" % (table.winfo_toplevel().winfo_rootx()+80,
                                         table.winfo_toplevel().winfo_rooty()+80))
        win.transient(table.winfo_toplevel())
        win.grab_set()
        win.resizable(False, False)

        # 操作函数
        def _add_row():
            """在辅助换撑窗口末尾新增一行，自动按层号生成换撑名称。"""
            existing_layers = {}
            for d in _data:
                wl = d.get("waler_layer", "")
                if wl: existing_layers[wl] = existing_layers.get(wl, 0) + 1
            new_layer = layer_options[0] if layer_options else "1"
            new_j = existing_layers.get(new_layer, 0) + 1
            new_data = {
                "waler_layer": new_layer, "name": f"{new_layer}-{new_j}", "connected": False,
                "mat_dc": mat_options[0], "dc_sec": sec_options[0],
                "mat_xc": mat_options[0], "xc_sec": sec_options[0],
                "dc_x": "/", "dc_y": "/", "xc_x": "/", "xc_y": "/",
                "proj_long": "/", "proj_short": "/", "diff_long": "/", "diff_short": "/",
            }
            _data.append(new_data)
            widgets = _create_row_widgets(len(_data) - 1, new_data)
            _row_widgets.append(widgets)
            _apply_entry_states(len(_data) - 1)
            _refresh_names()
            cv.configure(scrollregion=cv.bbox("all"))

        def _delete_row():
            """删除辅助换撑窗口的最后一行并刷新编号与名称。"""
            if not _data: return
            sel_r = len(_data) - 1
            for w in inner_ass.grid_slaves():
                g = w.grid_info()
                if g and g.get("row") and int(g["row"]) == sel_r + 1:
                    w.destroy()
            _data.pop(sel_r)
            _row_widgets.pop(sel_r)
            _refresh_numbers()
            _refresh_names()
            cv.configure(scrollregion=cv.bbox("all"))

        def _save():
            """校验辅助换撑各行并写回 sheet_refs['assist_replace_data']，随后关闭窗口。"""
            for ri, wg in enumerate(_row_widgets):
                rn = ri + 1
                if len(wg) < 16: continue
                # 支撑层 (index 1, StringVar)
                if str(wg[1].get()).strip() in ("", "/"):
                    messagebox.showwarning("数据检查", f"第{rn}行：必须选择支撑层")
                    return
                # 对撑 (index 5→截面, 8→X, 9→Y)
                has_dc = str(wg[5].get()).strip() not in ("", "/")
                if has_dc and all(str(wg[i].get()).strip() in ("", "/") for i in (8, 9)):
                    messagebox.showwarning("数据检查", f"第{rn}行：已选择对撑截面，请至少输入一个对撑间距")
                    return
                # 斜撑 (index 7→截面, 10→X, 11→Y)
                has_xc = str(wg[7].get()).strip() not in ("", "/")
                if has_xc:
                    for ci, xlbl in [(10, "X方向"), (11, "Y方向")]:
                        if str(wg[ci].get()).strip() in ("", "/"):
                            messagebox.showwarning("数据检查", f"第{rn}行：已选择斜撑截面，{xlbl}斜撑间距不能为空")
                            return
                # 与承台连接 (index 3→BooleanVar, 12-15→投影高差)
                if wg[3].get():
                    for ci, clbl in [(12,"承台长边投影"),(13,"承台短边投影"),
                                      (14,"承台长边高差"),(15,"承台短边高差")]:
                        if str(wg[ci].get()).strip() in ("", "/"):
                            messagebox.showwarning("数据检查", f"第{rn}行：已连接承台，{clbl}不能为空")
                            return
            sheet_refs["assist_replace_data"].clear()
            for ri, wg in enumerate(_row_widgets):
                sheet_refs["assist_replace_data"].append({
                    "waler_layer": str(wg[1].get()) if len(wg) > 1 else "",
                    "name": str(wg[2].cget("text")) if len(wg) > 2 and hasattr(wg[2], "cget") else "",
                    "connected": "是" if (len(wg) > 3 and wg[3].get()) else "否",
                    "mat_dc": str(wg[4].get()) if len(wg) > 4 else "/",
                    "dc_sec": str(wg[5].get()) if len(wg) > 5 else "/",
                    "mat_xc": str(wg[6].get()) if len(wg) > 6 else "/",
                    "xc_sec": str(wg[7].get()) if len(wg) > 7 else "/",
                    "dc_x": str(wg[8].get()) if len(wg) > 8 else "/",
                    "dc_y": str(wg[9].get()) if len(wg) > 9 else "/",
                    "xc_x": str(wg[10].get()) if len(wg) > 10 else "/",
                    "xc_y": str(wg[11].get()) if len(wg) > 11 else "/",
                    "proj_long": str(wg[12].get()) if len(wg) > 12 else "/",
                    "proj_short": str(wg[13].get()) if len(wg) > 13 else "/",
                    "diff_long": str(wg[14].get()) if len(wg) > 14 else "/",
                    "diff_short": str(wg[15].get()) if len(wg) > 15 else "/",
                })
            # 辅助换撑数据变更后刷新工况类型可用性
            _refresh_type_options()
            win.destroy()

        bottom_frame = tb.Frame(win)
        bottom_frame.pack(side=BOTTOM, fill=X, padx=10, pady=10)
        tb.Button(bottom_frame, text="取消", bootstyle=SECONDARY, width=10,
                  command=win.destroy).pack(side=RIGHT, padx=5)
        tb.Button(bottom_frame, text="保存", bootstyle=SUCCESS, width=10,
                  command=lambda: _save()).pack(side=RIGHT, padx=5)

        top_btn_frame = tb.Frame(win)
        top_btn_frame.pack(side=TOP, fill=X, padx=10, pady=(10, 5))
        tb.Button(top_btn_frame, text="+ 新增行", bootstyle=SUCCESS, width=10,
                  command=_add_row).pack(side=LEFT, padx=5)
        tb.Button(top_btn_frame, text="- 删除行", bootstyle=DANGER, width=10,
                  command=_delete_row).pack(side=LEFT, padx=5)

        HEADERS_A = ["序号", "选择已有支撑层", "换撑形态名称", "是否与承台连接",
                   "对撑材质", "对撑截面", "斜撑材质", "斜撑截面",
                   "对撑长边布置(m)", "对撑短边布置(m)",
                   "斜撑长边布置(m)", "斜撑短边布置(m)",
                   "承台长边投影(m)", "承台短边投影(m)",
                   "承台长边高差(m)", "承台短边高差(m)"]
        COL_W_A = [40, 110, 100, 110, 90, 100, 90, 100, 120, 120, 120, 120, 120, 120, 120, 120]
        COL_CONNECTED = 3
        COL_MAT_DC = 4
        COL_DC_SEC = 5
        COL_MAT_XC = 6
        COL_XC_SEC = 7
        COL_DC_X = 8
        COL_DC_Y = 9
        COL_XC_X = 10
        COL_XC_Y = 11
        COL_PROJ_START = 12
        COL_PROJ_END = 15

        wc = len(rows) if rows else 0
        layer_options = [str(i+1) for i in range(wc)]
        strut_sec_list = sheet_refs.get("strut_section_list", []) if sheet_refs else []
        sec_options = ["/"] + (strut_sec_list if strut_sec_list else [])
        mat_options = ["/"] + _get_steel_mat_options()

        _data = []
        _row_widgets = []

        for item in sheet_refs.get("assist_replace_data", []):
            _data.append({
                "waler_layer": item.get("waler_layer", ""),
                "name": item.get("name", ""),
                "connected": item.get("connected", "否") == "是",
                "mat_dc": item.get("mat_dc", "/"), "dc_sec": item.get("dc_sec", "/"),
                "mat_xc": item.get("mat_xc", "/"), "xc_sec": item.get("xc_sec", "/"),
                "dc_x": item.get("dc_x", "/"), "dc_y": item.get("dc_y", "/"),
                "xc_x": item.get("xc_x", "/"), "xc_y": item.get("xc_y", "/"),
                "proj_long": item.get("proj_long", "/"), "proj_short": item.get("proj_short", "/"),
                "diff_long": item.get("diff_long", "/"), "diff_short": item.get("diff_short", "/"),
            })

        tf = tb.Frame(win)
        tf.pack(fill=BOTH, expand=True, padx=10, pady=5)
        vs = tb.Scrollbar(tf, orient="vertical")
        vs.pack(side=RIGHT, fill=Y)
        hs = tb.Scrollbar(tf, orient="horizontal")
        hs.pack(side=BOTTOM, fill=X)
        cv = tk.Canvas(tf, borderwidth=0, highlightthickness=0,
                       xscrollcommand=hs.set, yscrollcommand=vs.set)
        cv.pack(side=LEFT, fill=BOTH, expand=True)
        hs.config(command=cv.xview); vs.config(command=cv.yview)
        inner_ass = tk.Frame(cv)
        cv.create_window((0, 0), window=inner_ass, anchor="nw")
        inner_ass.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        def _mw_rep(e):
            """辅助换撑窗口滚轮纵向滚动（内容高于视口时）。"""
            if inner_ass.winfo_height() > cv.winfo_height():
                cv.yview_scroll(int(-1 * (e.delta / 120)), "units")
        win.bind("<MouseWheel>", _mw_rep)

        for c, (h, w) in enumerate(zip(HEADERS_A, COL_W_A)):
            tk.Label(inner_ass, text=h, font=("Microsoft YaHei UI", 9, "bold"),
                     bg="#1e3a5f", fg="white", relief="solid", borderwidth=1,
                     width=w // 10, anchor="center").grid(row=0, column=c, sticky="nsew")

        def _refresh_names():
            """按层号顺序重新生成各行换撑名称（层号-序号）。"""
            lc = {}
            for d in _data:
                wl = d.get("waler_layer", "")
                if not wl: continue
                lc[wl] = lc.get(wl, 0) + 1
                d["name"] = f"{wl}-{lc[wl]}"
            for i, wg in enumerate(_row_widgets):
                if i < len(_data): wg[2].config(text=_data[i]["name"])

        def _refresh_numbers():
            """刷新辅助换撑各行的序号显示。"""
            for i, wg in enumerate(_row_widgets): wg[0].config(text=str(i + 1))

        def _auto_select_layer(idx):
            """根据换撑名称自动回填该行的支撑层号下拉值。"""
            if idx >= len(_data): return
            nm = _data[idx].get("name", "")
            if "-" in nm:
                _data[idx]["waler_layer"] = nm.split("-")[0]
                _row_widgets[idx][1].set(nm.split("-")[0])

        def _apply_entry_states(idx):
            """按对撑/斜撑截面与是否连接承台，启用/禁用辅助换撑该行输入框。"""
            if idx >= len(_row_widgets): return
            conn = _row_widgets[idx][COL_CONNECTED].get()
            dcs = _row_widgets[idx][COL_DC_SEC].get()
            xcs = _row_widgets[idx][COL_XC_SEC].get()

            # 斜撑截面始终可用（不再因与承台连接而禁用）
            ch_xc = inner_ass.grid_slaves(row=idx + 1, column=COL_XC_SEC)
            if ch_xc and isinstance(ch_xc[0], tb.Combobox):
                ch_xc[0].configure(state="readonly")

            for ci, vr in [(COL_DC_X, _row_widgets[idx][COL_DC_X]),
                           (COL_DC_Y, _row_widgets[idx][COL_DC_Y])]:
                ch = inner_ass.grid_slaves(row=idx + 1, column=ci)
                if ch and isinstance(ch[0], tk.Entry):
                    if dcs != "/": ch[0].configure(state="normal")
                    else: ch[0].configure(state="disabled"); vr.set("/")
            for ci, vr in [(COL_XC_X, _row_widgets[idx][COL_XC_X]),
                           (COL_XC_Y, _row_widgets[idx][COL_XC_Y])]:
                ch = inner_ass.grid_slaves(row=idx + 1, column=ci)
                if ch and isinstance(ch[0], tk.Entry):
                    if xcs != "/": ch[0].configure(state="normal")
                    else: ch[0].configure(state="disabled"); vr.set("/")
            for ci in range(COL_PROJ_START, COL_PROJ_END + 1):
                vr = _row_widgets[idx][ci]
                ch = inner_ass.grid_slaves(row=idx + 1, column=ci)
                if ch and isinstance(ch[0], tk.Entry):
                    if conn: ch[0].configure(state="normal")
                    else: ch[0].configure(state="disabled"); vr.set("/")

        def _on_checkbox_toggle(idx):
            """辅助换撑行"是否连接承台"勾选变化时刷新输入框状态。"""
            if idx >= len(_row_widgets): return
            _apply_entry_states(idx)

        def _on_combo_change(idx):
            """辅助换撑行下拉变化时更新层号、重算名称并刷新输入框状态。"""
            if idx >= len(_row_widgets) or idx >= len(_data): return
            _data[idx]["waler_layer"] = _row_widgets[idx][1].get()
            _refresh_names()
            _apply_entry_states(idx)

        def _create_row_widgets(r_idx, data_row):
            """创建辅助换撑窗口的一行控件（16 列）并返回其变量/控件列表。"""
            wg = []
            rn = r_idx + 1
            lbl_num = tk.Label(inner_ass, text=str(rn), relief="solid", borderwidth=1,
                               bg="white", font=("Microsoft YaHei UI", 9), padx=5, pady=3)
            lbl_num.grid(row=rn, column=0, sticky="nsew"); wg.append(lbl_num)
            lv = tk.StringVar(value=data_row.get("waler_layer", ""))
            cb_l = tb.Combobox(inner_ass, textvariable=lv, values=layer_options, width=8, state="readonly")
            cb_l.grid(row=rn, column=1, sticky="nsew", padx=1, pady=1)
            cb_l.bind("<<ComboboxSelected>>", lambda e, idx=r_idx: _on_combo_change(idx))
            wg.append(lv)
            lbl_n = tk.Label(inner_ass, text=data_row.get("name", ""), relief="solid", borderwidth=1,
                             bg="white", font=("Microsoft YaHei UI", 9), padx=5, pady=3)
            lbl_n.grid(row=rn, column=2, sticky="nsew"); wg.append(lbl_n)
            cv_ = tk.BooleanVar(value=data_row.get("connected", False))
            cf = tk.Frame(inner_ass, relief="solid", borderwidth=1, bg="white")
            cf.grid(row=rn, column=3, sticky="nsew")
            tb.Checkbutton(cf, variable=cv_, bootstyle="round-toggle",
                           command=lambda idx=r_idx: _on_checkbox_toggle(idx)).pack(expand=True)
            wg.append(cv_)
            mdv = tk.StringVar(value=data_row.get("mat_dc", "/"))
            cmd = tb.Combobox(inner_ass, textvariable=mdv, values=mat_options, width=9, state="readonly")
            cmd.grid(row=rn, column=4, sticky="nsew", padx=1, pady=1)
            cmd.bind("<<ComboboxSelected>>", lambda e, idx=r_idx: _on_combo_change(idx))
            wg.append(mdv)
            dv = tk.StringVar(value=data_row.get("dc_sec", "/"))
            cd = tb.Combobox(inner_ass, textvariable=dv, values=sec_options, width=10, state="readonly")
            cd.grid(row=rn, column=5, sticky="nsew", padx=1, pady=1)
            cd.bind("<<ComboboxSelected>>", lambda e, idx=r_idx: _on_combo_change(idx))
            wg.append(dv)
            mxv = tk.StringVar(value=data_row.get("mat_xc", "/"))
            cmx = tb.Combobox(inner_ass, textvariable=mxv, values=mat_options, width=9, state="readonly")
            cmx.grid(row=rn, column=6, sticky="nsew", padx=1, pady=1)
            cmx.bind("<<ComboboxSelected>>", lambda e, idx=r_idx: _on_combo_change(idx))
            wg.append(mxv)
            xv = tk.StringVar(value=data_row.get("xc_sec", "/"))
            cx = tb.Combobox(inner_ass, textvariable=xv, values=sec_options, width=10, state="readonly")
            cx.grid(row=rn, column=7, sticky="nsew", padx=1, pady=1)
            cx.bind("<<ComboboxSelected>>", lambda e, idx=r_idx: _on_combo_change(idx))
            wg.append(xv)
            for ci, k in [(8, "dc_x"), (9, "dc_y")]:
                vr = tk.StringVar(value=data_row.get(k, "/"))
                tk.Entry(inner_ass, textvariable=vr, width=6, font=("Microsoft YaHei UI", 9),
                         relief="solid", borderwidth=1, justify="center").grid(
                    row=rn, column=ci, sticky="nsew", padx=1, pady=1)
                wg.append(vr)
            for ci, k in [(10, "xc_x"), (11, "xc_y")]:
                vr = tk.StringVar(value=data_row.get(k, "/"))
                tk.Entry(inner_ass, textvariable=vr, width=6, font=("Microsoft YaHei UI", 9),
                         relief="solid", borderwidth=1, justify="center").grid(
                    row=rn, column=ci, sticky="nsew", padx=1, pady=1)
                wg.append(vr)
            for ci, k in [(12, "proj_long"), (13, "proj_short"), (14, "diff_long"), (15, "diff_short")]:
                vr = tk.StringVar(value=data_row.get(k, "/"))
                tk.Entry(inner_ass, textvariable=vr, width=7, font=("Microsoft YaHei UI", 9),
                         relief="solid", borderwidth=1, justify="center").grid(
                    row=rn, column=ci, sticky="nsew", padx=1, pady=1)
                wg.append(vr)
            return wg

        for i, d in enumerate(_data):
            _row_widgets.append(_create_row_widgets(i, d))
        for i in range(len(_data)):
            _apply_entry_states(i)
            _auto_select_layer(i)
        _refresh_names()

    btn_assist_replace.config(command=_open_assist_replace_support)

    # ---- 辅助加撑二级窗口 ----
    def _open_assist_add_support():
        """打开"辅助加撑设置"二级窗口（增删行、校验并写回 assist_add_data）。

        Returns:
            None
        """
        win = tb.Toplevel(table.winfo_toplevel())
        win.title("辅助加撑设置")
        win.geometry("1050x420+%d+%d" % (table.winfo_toplevel().winfo_rootx()+80,
                                          table.winfo_toplevel().winfo_rooty()+80))
        win.transient(table.winfo_toplevel())
        win.grab_set()
        win.resizable(False, False)

        def _add_row():
            """在辅助加撑窗口末尾新增一行，自动分配未占用的层号。"""
            used = set()
            for d in _data:
                ln = d.get("layer_no", "")
                if ln and ln.strip():
                    try: used.add(int(ln.strip()))
                    except: pass
            nl = next_layer
            while nl in used: nl += 1
            new_data = {
                "layer_no": str(nl), "elevation": "/",
                "mat_long": waler_mat_default, "sec_long": waler_sec_options[0],
                "mat_short": waler_mat_default, "sec_short": waler_sec_options[0],
                "connected": False,
                "mat_dc": "/", "dc_sec": strut_sec_options[0],
                "mat_xc": "/", "xc_sec": strut_sec_options[0],
                "dc_x": "/", "dc_y": "/", "xc_x": "/", "xc_y": "/",
                "proj_long": "/", "proj_short": "/", "diff_long": "/", "diff_short": "/",
            }
            _data.append(new_data)
            wg = _create_row_widgets(len(_data) - 1, new_data)
            _row_widgets.append(wg)
            _apply_entry_states(len(_data) - 1)
            cv.configure(scrollregion=cv.bbox("all"))

        def _delete_row():
            """删除辅助加撑窗口的最后一行并刷新编号。"""
            if not _data: return
            sr = len(_data) - 1
            for w in inner_add.grid_slaves():
                g = w.grid_info()
                if g and g.get("row") and int(g["row"]) == sr + 1:
                    w.destroy()
            _data.pop(sr)
            _row_widgets.pop(sr)
            _refresh_numbers()
            cv.configure(scrollregion=cv.bbox("all"))

        def _save():
            """校验辅助加撑各行并写回 sheet_refs['assist_add_data']，随后关闭窗口。"""
            # 从 StringVar 读取当前值（_data 存的是初始值，不实时更新）
            for ri, wg in enumerate(_row_widgets):
                rn = ri + 1
                if len(wg) < 20: continue
                # 1. 标高 (index 2)
                elev_sv = wg[2]
                try:
                    ev = str(elev_sv.get()).strip()
                    if ev in ("", "/"): raise ValueError
                    float(ev)
                except:
                    messagebox.showwarning("数据检查", f"第{rn}行：标高必须为有效数值")
                    return
                # 2. 围檩截面 (index 4,6)
                for idx, label in [(4, "长边截面"), (6, "宽边截面")]:
                    if str(wg[idx].get()).strip() in ("", "/"):
                        messagebox.showwarning("数据检查", f"第{rn}行：围檩{label}不能为空")
                        return
                # 3. 对撑 (index 9→截面, 12→X, 13→Y)
                has_dc = str(wg[9].get()).strip() not in ("", "/")
                if has_dc and all(str(wg[i].get()).strip() in ("", "/") for i in (12, 13)):
                    messagebox.showwarning("数据检查", f"第{rn}行：已选择对撑截面，请至少输入一个对撑间距")
                    return
                # 4. 斜撑 (index 11→截面, 14→X, 15→Y)
                has_xc = str(wg[11].get()).strip() not in ("", "/")
                if has_xc:
                    for ci, xlbl in [(14, "X方向"), (15, "Y方向")]:
                        if str(wg[ci].get()).strip() in ("", "/"):
                            messagebox.showwarning("数据检查", f"第{rn}行：已选择斜撑截面，{xlbl}斜撑间距不能为空")
                            return
                # 5. 与承台连接 (index 7→BooleanVar, 16-19→投影高差)
                if wg[7].get():  # connected
                    for ci, clbl in [(16,"承台长边投影"),(17,"承台短边投影"),(18,"承台长边高差"),(19,"承台短边高差")]:
                        if str(wg[ci].get()).strip() in ("", "/"):
                            messagebox.showwarning("数据检查", f"第{rn}行：已连接承台，{clbl}不能为空")
                            return
            sheet_refs["assist_add_data"].clear()
            for ri, wg in enumerate(_row_widgets):
                sheet_refs["assist_add_data"].append({
                    "layer_no": _data[ri]["layer_no"] if ri < len(_data) else "",
                    "elevation": str(wg[2].get()) if len(wg) > 2 else "/",
                    "mat_long": str(wg[3].get()) if len(wg) > 3 else "/",
                    "sec_long": str(wg[4].get()) if len(wg) > 4 else "/",
                    "mat_short": str(wg[5].get()) if len(wg) > 5 else "/",
                    "sec_short": str(wg[6].get()) if len(wg) > 6 else "/",
                    "connected": "是" if (len(wg) > 7 and wg[7].get()) else "否",
                    "mat_dc": str(wg[8].get()) if len(wg) > 8 else "/",
                    "dc_sec": str(wg[9].get()) if len(wg) > 9 else "/",
                    "mat_xc": str(wg[10].get()) if len(wg) > 10 else "/",
                    "xc_sec": str(wg[11].get()) if len(wg) > 11 else "/",
                    "dc_x": str(wg[12].get()) if len(wg) > 12 else "/",
                    "dc_y": str(wg[13].get()) if len(wg) > 13 else "/",
                    "xc_x": str(wg[14].get()) if len(wg) > 14 else "/",
                    "xc_y": str(wg[15].get()) if len(wg) > 15 else "/",
                    "proj_long": str(wg[16].get()) if len(wg) > 16 else "/",
                    "proj_short": str(wg[17].get()) if len(wg) > 17 else "/",
                    "diff_long": str(wg[18].get()) if len(wg) > 18 else "/",
                    "diff_short": str(wg[19].get()) if len(wg) > 19 else "/",
                })
            _refresh_type_options()
            win.destroy()

        bottom_frame = tb.Frame(win)
        bottom_frame.pack(side=BOTTOM, fill=X, padx=10, pady=10)
        tb.Button(bottom_frame, text="取消", bootstyle=SECONDARY, width=10,
                  command=win.destroy).pack(side=RIGHT, padx=5)
        tb.Button(bottom_frame, text="保存", bootstyle=SUCCESS, width=10,
                  command=lambda: _save()).pack(side=RIGHT, padx=5)

        top_btn_frame = tb.Frame(win)
        top_btn_frame.pack(side=TOP, fill=X, padx=10, pady=(10, 5))
        tb.Button(top_btn_frame, text="+ 新增行", bootstyle=SUCCESS, width=10,
                  command=_add_row).pack(side=LEFT, padx=5)
        tb.Button(top_btn_frame, text="- 删除行", bootstyle=DANGER, width=10,
                  command=_delete_row).pack(side=LEFT, padx=5)

        HEADERS_A = ["序号", "辅助支撑层号", "标高(m)", "围檩长边材质", "围檩长边截面", "围檩宽边材质", "围檩宽边截面",
                   "是否与承台连接", "对撑材质", "对撑截面", "斜撑材质", "斜撑截面",
                   "对撑长边布置(m)", "对撑短边布置(m)", "斜撑长边布置(m)", "斜撑短边布置(m)",
                   "承台长边投影(m)", "承台短边投影(m)", "承台长边高差(m)", "承台短边高差(m)"]
        COL_W_A = [40, 80, 70, 90, 100, 90, 100, 80, 90, 100, 90, 100, 90, 90, 90, 90, 80, 80, 70, 70]
        COL_CONNECTED = 7
        COL_MAT_DC = 8
        COL_DC_SEC = 9
        COL_MAT_XC = 10
        COL_XC_SEC = 11
        COL_DC_X = 12
        COL_DC_Y = 13
        COL_XC_X = 14
        COL_XC_Y = 15
        COL_PROJ_START = 16
        COL_PROJ_END = 19

        wc = len(rows) if rows else 0
        next_layer = wc + 1
        waler_sec_list = sheet_refs.get("waler_section_list", []) if sheet_refs else []
        waler_sec_options = ["/"] + (waler_sec_list if waler_sec_list else [])
        strut_sec_list = sheet_refs.get("strut_section_list", []) if sheet_refs else []
        strut_sec_options = ["/"] + (strut_sec_list if strut_sec_list else [])
        waler_mat_options = _get_steel_mat_options()
        strut_mat_options = ["/"] + waler_mat_options
        waler_mat_default = waler_mat_options[0] if waler_mat_options else ""

        _data = []
        _row_widgets = []

        for item in sheet_refs.get("assist_add_data", []):
            _data.append({
                "layer_no": item.get("layer_no", ""), "elevation": item.get("elevation", "/"),
                "mat_long": item.get("mat_long", "/"), "sec_long": item.get("sec_long", "/"),
                "mat_short": item.get("mat_short", "/"), "sec_short": item.get("sec_short", "/"),
                "connected": item.get("connected", "否") == "是",
                "mat_dc": item.get("mat_dc", "/"), "dc_sec": item.get("dc_sec", "/"),
                "mat_xc": item.get("mat_xc", "/"), "xc_sec": item.get("xc_sec", "/"),
                "dc_x": item.get("dc_x", "/"), "dc_y": item.get("dc_y", "/"),
                "xc_x": item.get("xc_x", "/"), "xc_y": item.get("xc_y", "/"),
                "proj_long": item.get("proj_long", "/"), "proj_short": item.get("proj_short", "/"),
                "diff_long": item.get("diff_long", "/"), "diff_short": item.get("diff_short", "/"),
            })

        tf = tb.Frame(win)
        tf.pack(fill=BOTH, expand=True, padx=10, pady=5)
        vs = tb.Scrollbar(tf, orient="vertical")
        vs.pack(side=RIGHT, fill=Y)
        hs = tb.Scrollbar(tf, orient="horizontal")
        hs.pack(side=BOTTOM, fill=X)
        cv = tk.Canvas(tf, borderwidth=0, highlightthickness=0,
                       xscrollcommand=hs.set, yscrollcommand=vs.set)
        cv.pack(side=LEFT, fill=BOTH, expand=True)
        hs.config(command=cv.xview); vs.config(command=cv.yview)
        inner_add = tk.Frame(cv)
        cv.create_window((0, 0), window=inner_add, anchor="nw")
        inner_add.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        def _mw_add(e):
            """辅助加撑窗口滚轮纵向滚动（内容高于视口时）。"""
            if inner_add.winfo_height() > cv.winfo_height():
                cv.yview_scroll(int(-1 * (e.delta / 120)), "units")
        win.bind("<MouseWheel>", _mw_add)

        for c, (h, w) in enumerate(zip(HEADERS_A, COL_W_A)):
            inner_add.grid_columnconfigure(c, minsize=w, weight=0)
            tk.Label(inner_add, text=h, font=("Microsoft YaHei UI", 9, "bold"),
                     bg="#1e3a5f", fg="white", relief="solid", borderwidth=1,
                     padx=5, pady=5, anchor="center").grid(row=0, column=c, sticky="nsew")

        def _refresh_numbers():
            """刷新辅助加撑各行的序号显示。"""
            for i, wg in enumerate(_row_widgets):
                if wg:
                    ch = inner_add.grid_slaves(row=i+1, column=0)
                    if ch: ch[0].config(text=str(i + 1))

        def _apply_entry_states(idx):
            """按对撑/斜撑截面与是否连接承台，启用/禁用辅助加撑该行输入框。"""
            if idx >= len(_row_widgets): return
            conn = _row_widgets[idx][COL_CONNECTED].get()
            dcs = _row_widgets[idx][COL_DC_SEC].get()
            xcs = _row_widgets[idx][COL_XC_SEC].get()

            # 斜撑截面始终可用（不再因与承台连接而禁用）
            ch_xc = inner_add.grid_slaves(row=idx + 1, column=COL_XC_SEC)
            if ch_xc and isinstance(ch_xc[0], tb.Combobox):
                ch_xc[0].configure(state="readonly")

            for ci, vr in [(COL_DC_X, _row_widgets[idx][COL_DC_X]),
                           (COL_DC_Y, _row_widgets[idx][COL_DC_Y])]:
                ch = inner_add.grid_slaves(row=idx + 1, column=ci)
                if ch and isinstance(ch[0], tk.Entry):
                    if dcs != "/": ch[0].configure(state="normal")
                    else: ch[0].configure(state="disabled"); vr.set("/")
            for ci, vr in [(COL_XC_X, _row_widgets[idx][COL_XC_X]),
                           (COL_XC_Y, _row_widgets[idx][COL_XC_Y])]:
                ch = inner_add.grid_slaves(row=idx + 1, column=ci)
                if ch and isinstance(ch[0], tk.Entry):
                    if xcs != "/": ch[0].configure(state="normal")
                    else: ch[0].configure(state="disabled"); vr.set("/")
            for ci in range(COL_PROJ_START, COL_PROJ_END + 1):
                vr = _row_widgets[idx][ci]
                ch = inner_add.grid_slaves(row=idx + 1, column=ci)
                if ch and isinstance(ch[0], tk.Entry):
                    if conn: ch[0].configure(state="normal")
                    else: ch[0].configure(state="disabled"); vr.set("/")

        def _on_checkbox_toggle(idx):
            """辅助加撑行"是否连接承台"勾选变化时刷新输入框状态。"""
            if idx >= len(_row_widgets): return
            _apply_entry_states(idx)

        def _on_combo_change(idx):
            """辅助加撑行下拉变化时刷新输入框状态。"""
            if idx >= len(_row_widgets) or idx >= len(_data): return
            _apply_entry_states(idx)

        def _create_row_widgets(r_idx, data_row):
            """创建辅助加撑窗口的一行控件（20 列）并返回其变量/控件列表。"""
            wg = []
            rn = r_idx + 1
            lbl_num = tk.Label(inner_add, text=str(rn), relief="solid", borderwidth=1,
                               bg="white", font=("Microsoft YaHei UI", 9), padx=5, pady=3)
            lbl_num.grid(row=rn, column=0, sticky="nsew"); wg.append(lbl_num)
            lv = data_row.get("layer_no", "")
            tk.Label(inner_add, text=str(lv), relief="solid", borderwidth=1,
                     bg="white", font=("Microsoft YaHei UI", 9), padx=5, pady=3).grid(
                row=rn, column=1, sticky="nsew"); wg.append(None)
            ev = tk.StringVar(value=data_row.get("elevation", "/"))
            tk.Entry(inner_add, textvariable=ev, font=("Microsoft YaHei UI", 9),
                     relief="solid", borderwidth=1, justify="center", width=6).grid(
                row=rn, column=2, sticky="nsew", padx=1, pady=1); wg.append(ev)
            mlv = tk.StringVar(value=data_row.get("mat_long", waler_mat_default))
            tb.Combobox(inner_add, textvariable=mlv, values=waler_mat_options, state="readonly", width=9).grid(
                row=rn, column=3, sticky="nsew", padx=1, pady=1); wg.append(mlv)
            sv = tk.StringVar(value=data_row.get("sec_long", "/"))
            tb.Combobox(inner_add, textvariable=sv, values=waler_sec_options, state="readonly", width=10).grid(
                row=rn, column=4, sticky="nsew", padx=1, pady=1); wg.append(sv)
            msv = tk.StringVar(value=data_row.get("mat_short", waler_mat_default))
            tb.Combobox(inner_add, textvariable=msv, values=waler_mat_options, state="readonly", width=9).grid(
                row=rn, column=5, sticky="nsew", padx=1, pady=1); wg.append(msv)
            sv2 = tk.StringVar(value=data_row.get("sec_short", "/"))
            tb.Combobox(inner_add, textvariable=sv2, values=waler_sec_options, state="readonly", width=10).grid(
                row=rn, column=6, sticky="nsew", padx=1, pady=1); wg.append(sv2)
            cv_ = tk.BooleanVar(value=data_row.get("connected", False))
            cf = tk.Frame(inner_add, relief="solid", borderwidth=1, bg="white")
            cf.grid(row=rn, column=7, sticky="nsew")
            tb.Checkbutton(cf, variable=cv_, bootstyle="round-toggle",
                           command=lambda idx=r_idx: _on_checkbox_toggle(idx)).pack(expand=True)
            wg.append(cv_)
            mdcv = tk.StringVar(value=data_row.get("mat_dc", "/"))
            tb.Combobox(inner_add, textvariable=mdcv, values=strut_mat_options, state="readonly", width=9).grid(
                row=rn, column=8, sticky="nsew", padx=1, pady=1); wg.append(mdcv)
            dcv = tk.StringVar(value=data_row.get("dc_sec", "/"))
            cd = tb.Combobox(inner_add, textvariable=dcv, values=strut_sec_options, state="readonly", width=10)
            cd.grid(row=rn, column=9, sticky="nsew", padx=1, pady=1)
            cd.bind("<<ComboboxSelected>>", lambda e, idx=r_idx: _on_combo_change(idx))
            wg.append(dcv)
            mxcv = tk.StringVar(value=data_row.get("mat_xc", "/"))
            tb.Combobox(inner_add, textvariable=mxcv, values=strut_mat_options, state="readonly", width=9).grid(
                row=rn, column=10, sticky="nsew", padx=1, pady=1); wg.append(mxcv)
            xcv = tk.StringVar(value=data_row.get("xc_sec", "/"))
            cx = tb.Combobox(inner_add, textvariable=xcv, values=strut_sec_options, state="readonly", width=10)
            cx.grid(row=rn, column=11, sticky="nsew", padx=1, pady=1)
            cx.bind("<<ComboboxSelected>>", lambda e, idx=r_idx: _on_combo_change(idx))
            wg.append(xcv)
            for ci, k in [(12, "dc_x"), (13, "dc_y"), (14, "xc_x"), (15, "xc_y")]:
                vr = tk.StringVar(value=data_row.get(k, "/"))
                tk.Entry(inner_add, textvariable=vr, font=("Microsoft YaHei UI", 9),
                         relief="solid", borderwidth=1, justify="center", width=5).grid(
                    row=rn, column=ci, sticky="nsew", padx=1, pady=1)
                wg.append(vr)
            for ci, k in [(16, "proj_long"), (17, "proj_short"), (18, "diff_long"), (19, "diff_short")]:
                vr = tk.StringVar(value=data_row.get(k, "/"))
                tk.Entry(inner_add, textvariable=vr, font=("Microsoft YaHei UI", 9),
                         relief="solid", borderwidth=1, justify="center", width=6).grid(
                    row=rn, column=ci, sticky="nsew", padx=1, pady=1)
                wg.append(vr)
            return wg

        for i, d in enumerate(_data):
            _row_widgets.append(_create_row_widgets(i, d))
        for i in range(len(_data)):
            _apply_entry_states(i)
        _refresh_numbers()

    btn_assist_add.config(command=_open_assist_add_support)

    # ==================== 辅助窗口保存后刷新主表 ====================
    def _refresh_type_options():
        """辅助换撑/加撑数据变更后刷新各行的工况类型下拉与支撑层数。

        Returns:
            None
        """
        available = _get_available_types()
        for i, rw in enumerate(_row_widgets):
            # 刷新工况类型下拉选项
            for w in inner.grid_slaves(row=i + 1, column=1):
                if isinstance(w, tb.Combobox):
                    w.configure(values=available)
                    break
            # 刷新支撑层数（当前行类型变化时）
            _apply_row_states(i)

    # ==================== 返回控制器 ====================
    # 返回控制器对象（兼容旧 tksheet API + 新接口）
    class _Controller:
        """施工阶段表格控制器：向外部提供统一的读写接口（兼容旧 tksheet API）。"""

        def get_stage_dict(self):
            """返回当前施工阶段字典。"""
            return get_stage_dict()
        def get_sheet_data(self):
            """后向兼容：返回含表头与数据行的二维列表。

            Returns:
                list[list[str]]: 首行表头，其后每行
                    [工况类型, 土顶标高, 水面标高, 支撑层数]。
            """
            sd = get_stage_dict()
            rows = [["各工况名称", "各工况基坑内土顶标高(m)", "各工况基坑内水面标高(m)", "各工况支撑变动"]]
            for k in sorted(sd.keys()):
                d = sd[k]
                rows.append([d["工况类型"], d["基坑内土顶标高"],
                            d["基坑内水面标高"], d["支撑层数"]])
            return rows
        def load_from_old_format(self, old_list):
            """从旧格式工况名列表加载数据。"""
            load_from_old_format(old_list)
        def load_from_stage_dict(self, new_dict):
            """从 stage_dict 加载数据。"""
            load_from_stage_dict(new_dict)
        def redraw(self):
            """空实现（兼容旧接口）。"""
            pass
        def focus_set(self):
            """空实现（兼容旧接口）。"""
            pass
        def refresh_type_options(self):
            """刷新工况类型下拉选项。"""
            _refresh_type_options()

    ctrl = _Controller()
    if sheet_refs is not None:
        sheet_refs["construct_stage"] = ctrl

    # ==================== 施工阶段图示 ====================
    # 图示区域（grid row=3，固定高度在底部）
    diagram_outer = tb.Frame(table)
    diagram_outer.grid(row=3, column=0, sticky="sew", padx=15, pady=(0, 5))
    diagram_outer.pack_propagate(False)
    diagram_outer.configure(height=350)

    # 上分隔线
    tb.Separator(diagram_outer, orient=HORIZONTAL).pack(fill=X, pady=(5, 2))

    # 标题
    _diagram_title = tb.Label(diagram_outer, text="施工阶段图示",
                               font=("Microsoft YaHei UI", 9, "bold"),
                               anchor='center')
    _diagram_title.pack(fill=X, pady=2)

    # 下分隔线
    tb.Separator(diagram_outer, orient=HORIZONTAL).pack(fill=X, pady=(2, 5))

    # 图示画布容器（在固定高度的 diagram_outer 内自然填充）
    diagram_inner = tb.Frame(diagram_outer)
    diagram_inner.pack(fill=BOTH, expand=True)

    try:
        from CofferDam_Diagram_FEM import StageDiagram
        waler_sec_dict = sheet_refs.get("waler_sec_dict", {}) if sheet_refs else {}
        strut_sec_dict = sheet_refs.get("strut_sec_dict", {}) if sheet_refs else {}
        stage_diagram = StageDiagram(diagram_inner, sheet_refs, ctrl,
                                     waler_sec_dict, strut_sec_dict,
                                     title_label=_diagram_title)
        if sheet_refs is not None:
            sheet_refs["stage_diagram"] = stage_diagram
        # 先刷新缓存确保有真实数据
        for _rk in ["refresh_basic", "refresh_substructure"]:
            _rf = sheet_refs.get(_rk)
            if _rf: _rf()
        stage_diagram.draw(0)
    except Exception as e:
        tb.Label(diagram_inner, text=f"图示加载失败: {e}",
                 font=("Microsoft YaHei UI", 9), foreground="red").pack(expand=True)

    return ctrl

def Load_Combo(table, Steel_Sheet_Pile_CofferDam_Load_dict, Load_Combo_Line_Data, sheet_refs=None):
    """构建荷载组合定义表（tksheet），并按土层计算方法自动生成组合行。

    水土分算时额外生成"基坑外水压力/基坑底水压力"组合，系数从 sheet_refs 的
    combo_factors 读取；表格编辑实时同步回 Load_Combo_Line_Data。

    Args:
        table (tk.Widget): 页面容器。
        Steel_Sheet_Pile_CofferDam_Load_dict (dict): 土层信息字典。
        Load_Combo_Line_Data (list[list]): 荷载组合数据（就地更新）。
        sheet_refs (dict|None): 跨 Tab 共享引用字典。

    Returns:
        tksheet.Sheet: 荷载组合表格控件。
    """
    # ==========================================
    # 1. 基础 UI 框架与配色初始化 (保持原样)
    # ==========================================
    container1 = tb.Frame(table)
    container1.pack(fill=X, pady=5)
    tb.Label(container1, text='荷载组合定义', bootstyle=PRIMARY, font=("Microsoft YaHei UI", 10, "bold")).pack(side=LEFT, padx=15, pady=5)

    border_frame = tk.Frame(table, bg="#c8d2de")
    border_frame.pack(fill=BOTH, expand=True, padx=15, pady=5)
    container2 = tb.Frame(border_frame)
    container2.pack(fill=BOTH, expand=True, padx=1, pady=1)

    sheet = Sheet(container2, font=("Microsoft YaHei UI", 9, "normal"), row_height=36, show_row_index=False, show_column_index=False, headers=False)
    sheet.pack(expand=True, fill="both")
    sheet.hide("header")

    # ==========================================
    # 2. 解耦后的核心控制逻辑
    # ==========================================
    def _get_calc_method():
        """从已绑定的土层框架控件中获取最新的计算方法。

        Returns:
            str: '水土合算' 或 '水土分算'（无法判断时默认水土分算）。
        """
        # 采用加权平均土层时，实际生成用单层"加权平均土层"，计算方法固定为水土合算
        if sheet_refs:
            _w = sheet_refs.get("_weighted_check")
            if _w is not None and _w.get():
                return '水土合算'
        if sheet_refs and "soil_frame" in sheet_refs:
            try:
                sf = sheet_refs["soil_frame"]
                raw_data = sf.sheet.get_sheet_data() if hasattr(sf, 'sheet') else sf.get_sheet_data()
                # 过滤出名称非空的有效土层行，提取计算方法
                methods = [str(r[5]).strip() for r in raw_data if len(r) > 5 and str(r[0]).strip()]
                if methods: 
                    return '水土合算' if any(m == '水土合算' for m in methods) else '水土分算'
            except Exception: 
                pass
        return '水土分算' # 默认兜底

    def _refresh_combo_by_soil(force=False):
        """按当前土层计算方法重建荷载组合行。

        Args:
            force (bool): True 时忽略与现有组合的一致性检查强制刷新。

        Returns:
            None
        """
        current_method = _get_calc_method()
        
        # 1. 确定预期的荷载组合列表
        expected_loads = ['自重', '基坑外主动土压力', '基坑底初始土反力']
        if current_method == '水土分算':
            expected_loads.insert(2, '基坑外水压力')
            expected_loads.append('基坑底水压力')

        # 2. 如果现有荷载项与预期完全一致，且未触发强刷，直接保留用户手填的系数
        current_loads = [row[1] for row in Load_Combo_Line_Data if len(row) > 1]
        if not force and current_loads == expected_loads:
            return 

        # 3. 核心解耦点：从传入的桥梁字典 sheet_refs 中读取当前项目的 4 个分项系数值
        f = sheet_refs.get('combo_factors', {}) if sheet_refs else {}
        bp = str(f.get('basic_partial', '1.25')).strip() or "1.25"
        bc = str(f.get('basic_combo', '1.0')).strip() or "1.0"
        sp = str(f.get('standard_partial', '1.0')).strip() or "1.0"
        sc = str(f.get('standard_combo', '1.0')).strip() or "1.0"

        # 4. 重置底层列表并推给表格
        Load_Combo_Line_Data.clear()
        for i, load in enumerate(expected_loads):
            Load_Combo_Line_Data.append([str(i+1), load, bp, bc, sp, sc])
            
        _apply_data_and_style()

    def _apply_data_and_style():
        """把荷载组合数据填充到表格并设置表头/斑马纹/锁定列等样式。"""
        head_data = [["序号", "荷载组合", "基本", "组合", "标准", "组合"], 
                     ["", "", "分项系数", "组合值", "分项系数", "组合值"]]
        data = head_data + Load_Combo_Line_Data
        sheet.set_sheet_data(data, redraw=False)
        
        sheet.highlight_rows(0, bg="#1e3a5f", fg="#ffffff")
        sheet.highlight_rows(1, bg="#e6edf4", fg="#1e3a5f")
        for i in range(2, len(data)):
            sheet.highlight_rows(i, bg="#ffffff" if i % 2 == 0 else "#f6f8fb", fg="#2c3e50")
            sheet.highlight_cells(row=i, column=0, bg="#f0f3f8")
            
        sheet.align_columns(columns=[0,1,2,3,4,5], align="center")
        sheet.align_cells(row=0, column=2, align="e")  # "基本" 靠右 (East)
        sheet.align_cells(row=0, column=3, align="w")  # "组合" 靠左 (West)
        sheet.align_cells(row=0, column=4, align="e")  # "标准" 靠右 (East)
        sheet.align_cells(row=0, column=5, align="w")  # "组合" 靠左 (West)
        sheet.redraw()
        try:
            # 锁定表头
            sheet.readonly_rows([0, 1])
            # 锁定前两列 (序号和荷载组合名称)
            sheet.readonly_columns([0, 1])
            sheet.enable_edit(columns=[2, 3, 4, 5]) 
        except: 
            pass

    # ==========================================
    # 3. 接口暴露与事件绑定
    # ==========================================
    if sheet_refs is not None:
        sheet_refs["refresh_load_combo"] = _refresh_combo_by_soil
        # 勾选/取消"采用加权平均土层参数"时即时重建荷载组合表
        _weighted_check_ref = sheet_refs.get("_weighted_check")
        if _weighted_check_ref is not None:
            try:
                _weighted_check_ref.trace_add("write", lambda *_: _refresh_combo_by_soil(True))
            except Exception:
                pass
    _refresh_combo_by_soil() # 初始化首次渲染

    def on_data_changed(event=None):
        """表格数据变化时同步回 Load_Combo_Line_Data（跳过两行表头）。"""
        try:
            all_data = sheet.get_sheet_data()
            if len(all_data) > 2:
                new_data = all_data[2:]
                Load_Combo_Line_Data.clear()
                Load_Combo_Line_Data.extend(new_data)
        except Exception as e:
            print(f"数据同步错误: {e}")

    def on_combo_resize(event):
        """容器尺寸变化时按比例调整荷载组合表各列宽度。"""
        total_w = max(event.width - 25, 600) 
        ratios = [80, 300, 120, 120, 120]
        widths = [int(total_w * (r / 860)) for r in ratios]
        sheet.set_column_widths(widths + [total_w - sum(widths)])
    container2.bind("<Configure>", on_combo_resize)
    sheet.enable_bindings("single_select", "row_select", "column_select", "edit_cell")
    sheet.set_options(edit_cell_validation=True)
    sheet.focus_set()
    return sheet

def Condition_Analysis(table, none_construction_stage_check_var, construction_stage_check_var,
                       Drawdown_height, Waterdown_height, Excavation_face_dewater,
                       Steel_Sheet_Pile_CofferDam_Load_dict, Load_Combo_Line_Data,
                       Construct_Stage_Line_Data, rows, Concrete_Plug_check_var,
                       Concrete_Blinding_check_var,
                       Solid_Level, Water_Level, sheet_refs=None):
    """构建"工况分析"页（Tab 7）：整体模型/施工阶段切换、参数与自动生成。

    整体模型模式显示荷载组合表；施工阶段模式显示施工阶段定义表与图示，
    并提供"自动生成"按钮（调用 Cal_Stage）。

    Args:
        table (tk.Widget): 页面容器。
        none_construction_stage_check_var, construction_stage_check_var (tk.Variable):
            整体模型/施工阶段模式开关。
        Drawdown_height, Waterdown_height, Excavation_face_dewater (tk.Variable):
            整体模型的超挖/基坑底降水/开挖面降水（m）。
        Steel_Sheet_Pile_CofferDam_Load_dict (dict): 土层信息字典。
        Load_Combo_Line_Data (list[list]): 荷载组合数据。
        Construct_Stage_Line_Data (list[list]): 施工阶段数据。
        rows (list[dict]): 围檩行集合。
        Concrete_Plug_check_var, Concrete_Blinding_check_var (tk.Variable): 封底/垫层开关。
        Solid_Level, Water_Level (tk.Variable): 土顶/水面标高（m）。
        sheet_refs (dict|None): 跨 Tab 共享引用字典。

    Returns:
        None
    """
    _, card = make_card_frame(table, "工况分析")

    chk_row = tb.Frame(card)
    chk_row.pack(fill=X, pady=5)
    check_overall = tb.Checkbutton(chk_row, text='生成整体模型', variable=none_construction_stage_check_var,
                          bootstyle=PRIMARY,
                          command=lambda: switch_table(none_construction_stage_check_var,
                                                       construction_stage_check_var,
                                                       frame_combo, frame_stage, 0))
    check_overall.pack(side=LEFT, padx=15, pady=5)
    check_stage = tb.Checkbutton(chk_row, text='生成施工阶段', variable=construction_stage_check_var,
                          bootstyle=PRIMARY,
                          command=lambda: switch_table(construction_stage_check_var,
                                                       none_construction_stage_check_var,
                                                       frame_stage, frame_combo, 1))
    check_stage.pack(side=RIGHT, padx=15, pady=5)
    
    # ---- 中继保存：施工阶段临时参数 ----
    _condition_temp_data = {"drawdown": "", "waterdown": "", "excavation_face_dewater": ""}

    # ---- 整体模型参数 frame ----
    param_row_overall = tb.LabelFrame(card, text="整体模型工况参数设置", bootstyle=PRIMARY, padding=5)
    param_row_overall.columnconfigure((0, 1, 2), weight=1)

    f1o = tb.Frame(param_row_overall)
    f1o.grid(row=0, column=0, sticky="ew", padx=10)
    tb.Label(f1o, text="超挖深度(m):", bootstyle=PRIMARY).pack(side=LEFT, padx=(0, 5), pady=5)
    tb.Entry(f1o, textvariable=Drawdown_height, width=12).pack(side=LEFT, pady=5)

    f2o = tb.Frame(param_row_overall)
    f2o.grid(row=0, column=1, sticky="n", padx=10)
    tb.Label(f2o, text="基坑底降水深度(m):", bootstyle=PRIMARY).pack(side=LEFT, padx=(0, 5), pady=5)
    tb.Entry(f2o, textvariable=Waterdown_height, width=12).pack(side=LEFT, pady=5)

    f3o = tb.Frame(param_row_overall)
    f3o.grid(row=0, column=2, sticky="ew", padx=10)
    tb.Frame(f3o).pack(side=LEFT, fill=X, expand=True)
    tb.Label(f3o, text="开挖面降水深度(m):", bootstyle=PRIMARY).pack(side=LEFT, padx=(0, 5), pady=5)
    tb.Entry(f3o, textvariable=Excavation_face_dewater, width=12).pack(side=LEFT, pady=5)

    # ---- 施工阶段参数 frame ----
    param_row_stage = tb.LabelFrame(card, text="自动工况参数设置", bootstyle=PRIMARY, padding=5)

    # 施工阶段专用变量（独立于整体模型）
    _stage_drawdown = tb.DoubleVar(value=0.0)
    _stage_waterdown = tb.DoubleVar(value=0.0)
    _stage_excavation_face_dewater = tb.DoubleVar(value=0.0)

    # 参数入口行（grid 布局，包在子 frame 中）
    _stage_entry_frame = tb.Frame(param_row_stage)
    _stage_entry_frame.pack(fill=X)
    _stage_entry_frame.columnconfigure((0, 1, 2), weight=1)

    f1s = tb.Frame(_stage_entry_frame)
    f1s.grid(row=0, column=0, sticky="ew", padx=10)
    tb.Label(f1s, text="超挖深度(m):", bootstyle=PRIMARY).pack(side=LEFT, padx=(0, 5), pady=5)
    tb.Entry(f1s, textvariable=_stage_drawdown, width=12).pack(side=LEFT, pady=5)

    f2s = tb.Frame(_stage_entry_frame)
    f2s.grid(row=0, column=1, sticky="n", padx=10)
    tb.Label(f2s, text="基坑底降水深度(m):", bootstyle=PRIMARY).pack(side=LEFT, padx=(0, 5), pady=5)
    tb.Entry(f2s, textvariable=_stage_waterdown, width=12).pack(side=LEFT, pady=5)

    f3s = tb.Frame(_stage_entry_frame)
    f3s.grid(row=0, column=2, sticky="ew", padx=10)
    tb.Frame(f3s).pack(side=LEFT, fill=X, expand=True)
    tb.Label(f3s, text="开挖面降水深度(m):", bootstyle=PRIMARY).pack(side=LEFT, padx=(0, 5), pady=5)
    tb.Entry(f3s, textvariable=_stage_excavation_face_dewater, width=12).pack(side=LEFT, pady=5)

    # 分隔线 + 自动生成按钮 + tip（pack 在 param_row_stage 内部）
    tb.Separator(param_row_stage, orient=HORIZONTAL).pack(fill=X, padx=10, pady=(5, 0))
    _stage_bottom_row = tb.Frame(param_row_stage)
    _stage_bottom_row.pack(fill=X, padx=10, pady=(5, 0))
    tb.Label(_stage_bottom_row, text="* 自动生成仅支持陆地钢板桩围堰或无封底的浅水围堰",
             font=("Microsoft YaHei UI", 9), foreground="#e67e22").pack(side=LEFT, padx=10)
    tb.Button(_stage_bottom_row, text="自动生成", bootstyle=PRIMARY, width=11,
              command=lambda: _auto_generate_stage()).pack(side=RIGHT)

    def _auto_generate_stage():
        """自动生成施工阶段（stage_dict 格式，含辅助步骤）。

        读取施工阶段专用参数与缓存后调用 Cal_Stage，并加载到施工阶段表格。

        Returns:
            None
        """
        # 先读取施工阶段专用参数（后清零）
        overdig_depth = _stage_drawdown.get()
        waterdown_at_bottom = _stage_waterdown.get()
        excavation_face_dewater = _stage_excavation_face_dewater.get()
        _stage_drawdown.set(0.0)
        _stage_waterdown.set(0.0)
        _stage_excavation_face_dewater.set(0.0)

        # 刷新缓存确保数据最新
        if sheet_refs:
            for k in ["refresh_substructure", "refresh_basic"]:
                fn = sheet_refs.get(k)
                if fn: fn()

        # 收集基本参数（先获取桩顶标高用于间距→标高转换）
        basic = sheet_refs.get("basic_cache", {}) if sheet_refs else {}
        try:
            pile_top_level = float(str(basic.get("pile_top_level", "0")).strip() or "0")
        except (ValueError, TypeError):
            pile_top_level = 0.0
        try:
            design_water_level = float(str(basic.get("water_level", "0")).strip() or "0")
        except (ValueError, TypeError):
            design_water_level = None
        try:
            ground_level = float(str(basic.get("ground_level", "0")).strip() or "0")
        except (ValueError, TypeError):
            ground_level = None

        # 收集围檩标高（间距=距上一道围檩的相对距离，累减计算绝对标高）
        substructure_cache = sheet_refs.get("substructure_cache", {}) if sheet_refs else {}
        waler_rows = substructure_cache.get("waler_rows", [])
        waler_elevations = []
        current_elev = pile_top_level
        for waler_row in waler_rows:
            try:
                spacing = float(str(waler_row.get("spacing", "0")).strip())
                current_elev -= spacing
                waler_elevations.append(current_elev)
            except (ValueError, TypeError):
                pass

        # 基坑底 = 承台底 - 垫层/封底厚度
        try:
            cap_bottom_level = float(str(basic.get("cap_bottom_level", "0")).strip() or "0")
        except (ValueError, TypeError):
            cap_bottom_level = 0.0
        try:
            blinding_thickness = float(str(substructure_cache.get("blinding_thickness", "0")).strip() or "0")
        except (ValueError, TypeError):
            blinding_thickness = 0.0
        try:
            plug_thickness = float(str(substructure_cache.get("plug_thickness", "0")).strip() or "0")
        except (ValueError, TypeError):
            plug_thickness = 0.0
        has_plug = Concrete_Plug_check_var.get()
        has_blinding = Concrete_Blinding_check_var.get() if Concrete_Blinding_check_var else False
        bottom_extra_depth = blinding_thickness if has_blinding else plug_thickness if has_plug else 0.0

        try:
            cap_height = float(str(basic.get("cap_height", "0")).strip() or "0")
        except:
            cap_height = 0.0
        stage_dict = Cal_Stage(waler_elevations, cap_bottom_level, bottom_extra_depth,
                               overdig_depth, excavation_face_dewater, waterdown_at_bottom,
                               cap_height=cap_height,
                               design_water_level=design_water_level,
                               ground_level=ground_level,
                               has_plug=has_plug, has_blinding=has_blinding)

        stage_controller = sheet_refs.get("construct_stage") if sheet_refs else None
        if stage_controller and hasattr(stage_controller, "load_from_stage_dict"):
            stage_controller.load_from_stage_dict(stage_dict)

    # 初始显示：整体模型参数
    param_row_overall.pack(fill=X, padx=10, pady=5)

    content_frame = tb.Frame(table)
    content_frame.pack(fill=BOTH, expand=True, padx=12, pady=5)
    frame_combo = tb.Frame(content_frame)
    frame_stage = tb.Frame(content_frame)

    combo_sheet = Load_Combo(frame_combo, Steel_Sheet_Pile_CofferDam_Load_dict, Load_Combo_Line_Data, sheet_refs)
    stage_sheet = Construct_Stage(frame_stage, Construct_Stage_Line_Data, rows,
                            Drawdown_height, Concrete_Plug_check_var, Solid_Level, Water_Level,
                            Concrete_Blinding_check_var=Concrete_Blinding_check_var,
                            sheet_refs=sheet_refs)
    if sheet_refs is not None:
        sheet_refs["load_combo"] = combo_sheet
        sheet_refs["construct_stage"] = stage_sheet
        sheet_refs["frame_combo"] = frame_combo
        sheet_refs["frame_stage"] = frame_stage

    if construction_stage_check_var.get():
        frame_stage.pack(fill=BOTH, expand=True)
        frame_combo.pack_forget()
    else:
        frame_combo.pack(fill=BOTH, expand=True)
        frame_stage.pack_forget()

    def switch_table(check_var1, check_var2, frame1, frame2, param):
        """在整体模型与施工阶段之间切换显示与参数区。

        Args:
            check_var1, check_var2 (tk.Variable): 目标/对侧模式开关。
            frame1, frame2 (tk.Frame): 目标/对侧内容区。
            param (int): 0 表示整体模型，1 表示施工阶段。

        Returns:
            None
        """
        check_var1.set(True)
        check_var2.set(False)
        frame2.pack_forget()
        frame1.pack(fill=BOTH, expand=True)
        # 切换参数 frame 显示
        if param == 0:  # 整体模型
            param_row_stage.pack_forget()
            param_row_overall.pack(fill=X, padx=10, pady=5)
            combo_sheet = sheet_refs.get("load_combo")
            if combo_sheet:
                combo_sheet.focus_set()
                combo_sheet.redraw()
        else:  # 施工阶段
            param_row_overall.pack_forget()
            param_row_stage.pack(fill=X, padx=10, pady=5)
            stage_sheet = sheet_refs.get("construct_stage")
            if stage_sheet:
                stage_sheet.focus_set()
                stage_sheet.redraw()
            # 刷新缓存+图示
            for _rk in ["refresh_basic", "refresh_substructure"]:
                _rf = sheet_refs.get(_rk)
                if _rf: _rf()
            sd = sheet_refs.get("stage_diagram")
            if sd and hasattr(sd, "draw"):
                try: sd.draw(0)
                except: pass

    def switch_table_from_project():
        """项目切换时根据 construction_stage_check_var 自动切换显示。

        Returns:
            None
        """
        if construction_stage_check_var.get():
            frame_combo.pack_forget()
            frame_stage.pack(fill=BOTH, expand=True)
            param_row_overall.pack_forget()
            param_row_stage.pack(fill=X, padx=10, pady=5)
        else:
            frame_stage.pack_forget()
            frame_combo.pack(fill=BOTH, expand=True)
            param_row_stage.pack_forget()
            param_row_overall.pack(fill=X, padx=10, pady=5)
    if sheet_refs is not None:
        sheet_refs["switch_table_from_project"] = switch_table_from_project

    # ---- condition cache ----
    def _refresh_condition_cache():
        """把工况分析页当前参数写入 sheet_refs['condition_cache']，供保存读取。"""
        if sheet_refs is None: return
        sheet_refs["condition_cache"] = {
            "overdig_depth": str(Drawdown_height.get()),
            "drawdown": str(Waterdown_height.get()),
            "excavation_face_dewater": str(Excavation_face_dewater.get()),
            "has_construction_stages": str(construction_stage_check_var.get()),
        }
    if sheet_refs is not None:
        sheet_refs["refresh_condition"] = _refresh_condition_cache


# ===== 运行入口 =====


