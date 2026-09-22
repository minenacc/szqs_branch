# 1. 标准库
import os
import re
import copy

# 2. 第三方库
import openpyxl
from openpyxl.styles import Alignment


# ============================================================
# 路径工具
# ============================================================

def get_excel_path():
    """解析 Trestle_ParamTable.xlsx 的完整路径"""
    try:
        import winreg
        registry_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                      r"Software\ShuZhiQiaoShi")
        applocation, _ = winreg.QueryValueEx(registry_key, "Applocation")
        winreg.CloseKey(registry_key)
    except Exception:
        applocation = None
    if applocation:
        new_path = os.path.join(applocation, "Support", "program_param", "Trestle_ParamTable.xlsx")
        if os.path.exists(new_path):
            return new_path
    fallback = os.path.join(os.path.dirname(__file__), "Trestle_ParamTable.xlsx")
    if os.path.exists(fallback):
        return fallback
    return None


def get_properties_path():
    """获取 properties_parameter.xlsx 的完整路径"""
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\ShuZhiQiaoShi")
        applocation, _ = winreg.QueryValueEx(key, "Applocation")
        winreg.CloseKey(key)
    except Exception:
        applocation = None
    if applocation:
        path = os.path.join(applocation, "Support", "basic_param", "properties_parameter.xlsx")
        if os.path.exists(path):
            return path
    return None


# ============================================================
# 通用工具（读写共用）
# ============================================================

_CENTER_ALIGN = Alignment(horizontal='center', vertical='center')


def _replace_empty_with_slash(val):
    """空值返回 '/'，非空返回 str(val)"""
    if val is None:
        return "/"
    s = str(val).strip()
    return s if s else "/"


def _center_align_all_cells(wb):
    """遍历所有 sheet 的所有非空 cell，统一设置居中对齐"""
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row,
                                min_col=1, max_col=ws.max_column):
            for cell in row:
                if cell.value is not None:
                    cell.alignment = _CENTER_ALIGN


def _safe_str(val, default=""):
    """
    安全地将值转换为字符串：处理 None、去除首尾空格。
    
    :param val: 原始值
    :param default: 遇到 None 时的默认返回值
    """
    if val is None:
        return default
    # 统一转换为字符串并去除首尾空白符
    res = str(val).strip()
    # 将中文逗号(，)和顿号(、)替换为英文半角逗号(,)
    res = res.replace("，", ",").replace("、", ",")
    return res


def _split_and_strip_by_comma(val):
    """以逗号分割字符串，去除每个元素的空格，并过滤空字符串"""
    if not val:
        return []
    return [x.strip() for x in val.split(",") if x.strip()]


def _safe_to_float(val, default=0.0):
    """安全转换为浮点数，转换失败则返回默认值"""
    if val is None:
        return default
    try:
        return float(val)
    except (ValueError, TypeError):
        return default


def _join_with_commas(items):
    """将序列元素转换为字符串并用逗号拼接"""
    return ",".join(str(x) for x in items)


def _split_bracket(val):
    """按 [] 分组拆分，支持逗号分隔的多个 [group] 串"""
    if not val:
        return []
    return [x.strip() for x in re.split(r'(?<=\])\s*,\s*(?=\[)', val) if x.strip()]


def _format_as_bracketed_list(items):
    """将序列格式化为带方括号的字符串，如 [1,2,3]"""
    if not items:
        return "[]"
    return "[" + ",".join(str(x) for x in items) + "]"


def _build_header_map(ws, header_row=1):
    """构建 {表头名: 列号(1-based)} 映射"""
    hmap = {}
    for c in range(1, ws.max_column + 1):
        val = ws.cell(header_row, c).value
        if val is not None:
            h = str(val).strip()
            if h:
                hmap[h] = c
    return hmap


def _cell(ws, row, header_name, hmap):
    """按表头名读取单元格值"""
    col = hmap.get(header_name)
    if col is None:
        return None
    return ws.cell(row, col).value


def _find_sheet(wb, names):
    """按名称列表依次查找 sheet，返回第一个匹配的 sheet 对象"""
    if isinstance(names, str):
        names = [names]  # 防字符串被按字符迭代（'库' 会误匹配其它 sheet）
    for name in names:
        if name in wb.sheetnames:
            return wb[name]
    for sn in wb.sheetnames:
        for name in names:
            if name in sn:
                return wb[sn]
    return None

# def _clear_from_row(ws, start_row):
#     """清空 start_row 以下全部内容"""
#     for r in range(start_row, ws.max_row + 1):
#         for c in range(1, ws.max_column + 1):
#             ws.cell(r, c).value = None


def _norm_sec_name(name):
    """归一化截面名：×/x/Ｘ/* 统一为大写X，去空格"""
    if not name:
        return ""
    return str(name).strip().replace("×", "X").replace("Ｘ", "X").replace("x", "X").replace("*", "X")


# ============================================================
# 读取 — 基础库（properties_parameter.xlsx）
# ============================================================

def load_section_library(section_xlsx_path):
    """读取 properties_parameter.xlsx 截面库

    properties_parameter.xlsx 包含多个 Sheet，每个代表一类截面：
      I/2I/HM/2HM/HN/2HN/HW/2HW — 工字钢/H型钢（C10=H, C11=B, C12=tw, C13=tf1, C14=r1, C15=tf2）
      O — 钢管（C7=D, C8=d）
      C/2C — 槽钢（C10=H, C11=B, C12=tw, C13=tf1, C14=tf2/C）
      ∠ — 角钢（C10=H, C11=B, C12=tw, C13=r1, C14=r2）

    Returns:
        {section_name: {'params': {键: 值}, 'type': str, 'A': float, 'Ix': float, ..., 'MCT': str}}
    """
    if not section_xlsx_path or not os.path.exists(section_xlsx_path):
        return {}

    # 列头映射：每个 sheet_type 的 dim_cols → params 键名
    _DIM_KEYS = {
        'I':   ['H', 'B', 'tw', 'tf1', 'r1', 'r2'],
        '2I':  ['H', 'B', 'tw', 'tf1', 'C',  'tf2'],
        'HM':  ['H', 'B', 'tw', 'tf1', 'r1'],
        '2HM': ['H', 'B', 'tw', 'tf1', 'C',  'tf2'],
        'HN':  ['H', 'B', 'tw', 'tf1', 'r1'],
        '2HN': ['H', 'B', 'tw', 'tf1', 'C',  'tf2'],
        'HW':  ['H', 'B', 'tw', 'tf1', 'r1'],
        '2HW': ['H', 'B', 'tw', 'tf1', 'C',  'tf2'],
        'C':   ['H', 'B', 'tw', 'tf1', 'tf2'],
        '2C':  ['H', 'B', 'tw', 'tf1', 'C'],
        '∠':   ['H', 'B', 'tw', 'r1', 'r2'],
        'O':   ['D', 'd'],
    }

    SHEET_COL_CONFIG = {
        'I':   {'hdr_row': 1, 'dim_cols': [10, 11, 12, 13, 14, 15]},
        '2I':  {'hdr_row': 1, 'dim_cols': [10, 11, 12, 13, 14, 15]},
        'HM':  {'hdr_row': 1, 'dim_cols': [10, 11, 12, 13, 14]},
        '2HM': {'hdr_row': 1, 'dim_cols': [10, 11, 12, 13, 14, 15]},
        'HN':  {'hdr_row': 1, 'dim_cols': [10, 11, 12, 13, 14]},
        '2HN': {'hdr_row': 1, 'dim_cols': [10, 11, 12, 13, 14, 15]},
        'HW':  {'hdr_row': 1, 'dim_cols': [10, 11, 12, 13, 14]},
        '2HW': {'hdr_row': 1, 'dim_cols': [10, 11, 12, 13, 14, 15]},
        'C':   {'hdr_row': 1, 'dim_cols': [10, 11, 12, 13, 14]},
        '2C':  {'hdr_row': 1, 'dim_cols': [10, 11, 12, 13, 14]},
        '∠':  {'hdr_row': 1, 'dim_cols': [10, 11, 12, 13, 14]},
        'O':   {'hdr_row': 1, 'dim_cols': [7, 8]},
    }

    wb = openpyxl.load_workbook(section_xlsx_path, data_only=True)
    library = {}

    for sheet_type, config in SHEET_COL_CONFIG.items():
        if sheet_type not in wb.sheetnames:
            continue
        ws = wb[sheet_type]
        hm = _build_header_map(ws, header_row=config['hdr_row'])
        dim_keys = _DIM_KEYS.get(sheet_type, [])
        for r in range(2, ws.max_row + 1):
            name = _safe_str(ws.cell(r, 1).value)
            if not name:
                continue
            # 读取截面尺寸 → 构建 params 字典
            params = {}
            for i, dc in enumerate(config['dim_cols']):
                v = ws.cell(r, dc).value
                try:
                    fv = float(v) if v is not None else 0.0
                except (ValueError, TypeError):
                    fv = 0.0
                if i < len(dim_keys) and dim_keys[i]:
                    params[dim_keys[i]] = fv
            # 读取力学特性
            library[name] = {
                'params': params,
                'type': sheet_type,
                'A': _safe_to_float(_cell(ws, r, 'A', hm)),
                'Ix': _safe_to_float(_cell(ws, r, 'Ix', hm)) if sheet_type != 'O' else _safe_to_float(_cell(ws, r, 'I', hm)),
                'Iy': _safe_to_float(_cell(ws, r, 'Iy', hm)) if sheet_type != 'O' else _safe_to_float(_cell(ws, r, 'I', hm)),
                'Wx': _safe_to_float(_cell(ws, r, 'Wx', hm)) if sheet_type != 'O' else _safe_to_float(_cell(ws, r, 'W', hm)),
                'Wy': _safe_to_float(_cell(ws, r, 'Wy', hm)) if sheet_type != 'O' else _safe_to_float(_cell(ws, r, 'W', hm)),
                'ix': _safe_to_float(_cell(ws, r, 'ix', hm)) if sheet_type != 'O' else _safe_to_float(_cell(ws, r, 'i', hm)),
                'iy': _safe_to_float(_cell(ws, r, 'iy', hm)) if sheet_type != 'O' else _safe_to_float(_cell(ws, r, 'i', hm)),
                '每延米重': _safe_to_float(_cell(ws, r, '每延米重', hm)),
            }

    wb.close()
    return library


def load_material_library(section_xlsx_path):
    """读取 properties_parameter.xlsx 中 '材质' sheet 的钢材/混凝土/钢筋参数

    Returns:
        (steel_dict, concrete_dict, rebar_dict)

    steel_dict: {规范名: {"钢结构规范编号": str, "Midas对应钢结构规范编号": str, "牌号参数": {牌号: 参数dict}}}
    concrete_dict: {规范名: {"混凝土规范编号": str, "Midas对应混凝土规范编号": str, "强度等级参数": {等级: 参数dict}}}
    rebar_dict: {规范名: {"钢筋规范标号": str, "牌号参数": {牌号: 参数dict}}}
    """
    if not section_xlsx_path or not os.path.exists(section_xlsx_path):
        return {}, {}, {}

    wb = openpyxl.load_workbook(section_xlsx_path, data_only=True)
    if '材质' not in wb.sheetnames:
        wb.close()
        return {}, {}, {}

    ws = wb['材质']
    # 用 R2 的表头名建 map（R1 是合并单元格大类标题，无效）
    hm = _build_header_map(ws, header_row=2)

    # ── 钢材 ──
    # 每列存 "f,fv,fcc,fy,fu" 逗号分隔，5列对应5档厚度
    steel_thk_keys = [16, 40, 63, 80, 100]
    steel_thk_headers = [
        "钢材厚度小于等于16mm",
        "钢材厚度大于16mm小于等于40mm",
        "钢材厚度大于40mm小于等于63mm",
        "钢材厚度大于63mm小于等于80mm",
        "钢材厚度大于80mm小于等于100mm",
    ]
    steel_strength_keys = ["f", "fv", "fcc", "fy", "fu"]

    steel_dict = {}
    for r in range(3, ws.max_row + 1):
        std_name = _safe_str(_cell(ws, r, "钢结构规范名", hm))
        if not std_name:
            continue
        if std_name not in steel_dict:
            steel_dict[std_name] = {
                "钢结构规范编号": _safe_str(_cell(ws, r, "钢结构规范编号", hm)),
                "Midas对应钢结构规范编号": _safe_str(_cell(ws, r, "Midas对应钢结构规范编号", hm)),
                "牌号参数": {},
            }
        grade = _safe_str(_cell(ws, r, "钢材牌号", hm))
        if not grade:
            continue
        # 初始化每个力学属性的厚度字典
        entry = {sk: {} for sk in steel_strength_keys}
        for tv, thk_hdr in zip(steel_thk_keys, steel_thk_headers):
            raw = _safe_str(_cell(ws, r, thk_hdr, hm))
            vals = [x.strip() for x in raw.split(",") if x.strip()]
            for i, sk in enumerate(steel_strength_keys):
                entry[sk][tv] = float(vals[i]) if i < len(vals) else 0.0
        entry["E"] = _safe_to_float(_cell(ws, r, "钢材弹性模量", hm))
        entry["G"] = _safe_to_float(_cell(ws, r, "钢材剪切变形模量", hm))
        entry["a"] = _safe_to_float(_cell(ws, r, "钢材线膨胀系数", hm))
        entry["p"] = _safe_to_float(_cell(ws, r, "钢材质量密度", hm))
        steel_dict[std_name]["牌号参数"][grade] = entry

    # ── 混凝土 ──
    concrete_dict = {}
    for r in range(3, ws.max_row + 1):
        std_name = _safe_str(_cell(ws, r, "混凝土规范名", hm))
        if not std_name:
            continue
        if std_name not in concrete_dict:
            concrete_dict[std_name] = {
                "混凝土规范编号": _safe_str(_cell(ws, r, "混凝土规范编号", hm)),
                "Midas对应混凝土规范编号": _safe_str(_cell(ws, r, "Midas对应混凝土规范编号", hm)),
                "强度等级参数": {},
            }
        grade = _safe_str(_cell(ws, r, "混凝土强度等级", hm))
        if not grade:
            continue
        concrete_dict[std_name]["强度等级参数"][grade] = {
            "fck": _safe_to_float(_cell(ws, r, "混凝土轴心抗压强度标准值", hm)),
            "ftk": _safe_to_float(_cell(ws, r, "混凝土轴心抗拉强度标准值", hm)),
            "fc":  _safe_to_float(_cell(ws, r, "混凝土轴心抗压强度设计值", hm)),
            "ft":  _safe_to_float(_cell(ws, r, "混凝土轴心抗拉强度设计值", hm)),
            "Ec":  _safe_to_float(_cell(ws, r, "混凝土弹性模量", hm)),
            "Gc":  _safe_to_float(_cell(ws, r, "混凝土剪切变形模量", hm)),
            "a":   _safe_to_float(_cell(ws, r, "混凝土线膨胀系数", hm)),
            "v":   _safe_to_float(_cell(ws, r, "混凝土泊松比", hm)),
        }

    # ── 钢筋 ──
    rebar_dict = {}
    for r in range(3, ws.max_row + 1):
        std_name = _safe_str(_cell(ws, r, "钢筋规范名", hm))
        if not std_name:
            continue
        if std_name not in rebar_dict:
            rebar_dict[std_name] = {
                "钢筋规范标号": _safe_str(_cell(ws, r, "钢筋规范标号", hm)),
                "牌号参数": {},
            }
        grade = _safe_str(_cell(ws, r, "钢筋牌号", hm))
        if not grade:
            continue
        rebar_dict[std_name]["牌号参数"][grade] = {
            "fy":    _safe_to_float(_cell(ws, r, "钢筋抗拉强度设计值", hm)),
            "fy_":   _safe_to_float(_cell(ws, r, "钢筋抗压强度设计值", hm)),
            "fyk":   _safe_to_float(_cell(ws, r, "钢筋屈服强度标准值", hm)),
            "fstk":  _safe_to_float(_cell(ws, r, "钢筋极限强度标准值", hm)),
            "E":     _safe_to_float(_cell(ws, r, "钢筋弹性模量", hm)),
            "delta": _safe_to_float(_cell(ws, r, "钢筋总伸长率限值", hm)),
        }

    wb.close()
    return steel_dict, concrete_dict, rebar_dict


def _parse_material_value(val):
    """解析 '规范编号,牌号' → (spec_id, brand)

    返回 (None, None) 如果 val 为空或 '/'
    兼容旧格式（无规范前缀）：spec_id 返回 None
    """
    if not val or val == "/":
        return None, None
    parts = val.split(",", 1)
    if len(parts) == 2:
        return parts[0].strip(), parts[1].strip()
    return None, val.strip()


def _resolve_section_ref(val, section_db):
    """将截面编号/名称解析为完整截面数据"""
    if val is None or str(val).strip() in ("", "/") or not section_db:
        return val
    try:
        seq = int(str(val).strip())
        if seq in section_db:
            rec = section_db[seq]
            return {
                "section_name": rec.get("name", ""),
                "section_type": rec.get("type", ""),
                "params": rec.get("params", {}),
            }
    except (ValueError, TypeError):
        pass
    return val


def _sec_dim(db_entry, key, default):
    """从截面库条目获取尺寸（统一 params 字典格式）"""
    if "params" in db_entry:
        return float(db_entry["params"].get(key, default))
    return float(default)


# ============================================================
# 读取 — 项目参数（Trestle_ParamTable.xlsx）
# ============================================================

PROJECT_TEMPLATE = {
    # ---- 项目身份 ----
    "project_name": "",
    "project_no": "",
    "bridge_type": "上承式桁架梁栈桥",

    # ---- 基本参数 ----
    "basic": {"span": "15", "bridge_width": "6", "water_level": "+12.0", "deck_level": "+12.5", "x_origin": ""},

    # ---- 下部结构（cache 格式：逐支撑展开）----
    "substructure": {
        "0": {
            "type": "单排桩",
            "pile_type": "单排桩",
            "pile_sec": "630X8",
            "pile_sec_detail": {"section_type": "O", "section_name": "630X8", "params": {"D": "630", "d": "8"}},
            "pile_len": "10",
            "pile_trans_space": "1000,4000",
            "pile_long_space": "/",
            "expansion_len": "/",
            "dist_trans_sec": "2I45a",
            "dist_trans_sec_detail": {"section_type": "2I", "section_name": "2I45a", "params": {"H": "450", "B": "150", "tw": "11.5", "tf1": "18", "C": "100", "tf2": "18"}},
            "dist_trans_len": "6000",
            "dist_long_sec": "/",
            "dist_long_len": "/",
            "dist_trans_offset": "3000",
            "dist_long_center_offset": "/",
            "brace_form": "-",
            "brace_sec": "219X6",
            "brace_sec_detail": {"section_type": "O", "section_name": "219X6", "params": {"D": "219", "d": "6"}},
            "brace_sec_tilt": "/",
            "brace_space": "620,/",
            "brace_height": "2500",
            "pile_material": "",
            "dist_trans_material": "",
            "dist_long_material": "/",
            "brace_material": "",
            "soil_data": {},
        },
        "1": {
            "type": "单排桩",
            "pile_type": "单排桩",
            "pile_sec": "630X8",
            "pile_sec_detail": {"section_type": "O", "section_name": "630X8", "params": {"D": "630", "d": "8"}},
            "pile_len": "10",
            "pile_trans_space": "1000,4000",
            "pile_long_space": "/",
            "expansion_len": "/",
            "dist_trans_sec": "2I45a",
            "dist_trans_sec_detail": {"section_type": "2I", "section_name": "2I45a", "params": {"H": "450", "B": "150", "tw": "11.5", "tf1": "18", "C": "100", "tf2": "18"}},
            "dist_trans_len": "6000",
            "dist_long_sec": "/",
            "dist_long_len": "/",
            "dist_trans_offset": "3000",
            "dist_long_center_offset": "/",
            "brace_form": "-",
            "brace_sec": "219X6",
            "brace_sec_detail": {"section_type": "O", "section_name": "219X6", "params": {"D": "219", "d": "6"}},
            "brace_sec_tilt": "/",
            "brace_space": "620,/",
            "brace_height": "2500",
            "pile_material": "",
            "dist_trans_material": "",
            "dist_long_material": "/",
            "brace_material": "",
            "soil_data": {},
        },
    },
    # ---- 土层标签（项目命名空间前缀，派生 {label}地形表 / {label}土层表）----
    "soil_label": "",

    # ---- 上部结构（components 统一结构）----
    "bridge_deck_type_ref": "1",
    "components": {
        "上承式桁架梁栈桥": {
            "deck_type": "单层钢面板",
            "deck_thickness": "10",
            "deck_material": "",
            "deck_trans_ecc": "3000",
            "deck_end_ext": "0,0",
            "rib_trans_sec": "I14",
            "rib_trans_sec_detail": {"section_type": "I", "section_name": "I14", "params": {"H": "140", "B": "80", "tw": "5.5", "tf1": "9.1", "r1": "7.5", "r2": "3.8"}},
            "rib_trans_space": "0,20@750",
            "rib_trans_material": "",
            "rib_long_sec": "/",
            "rib_long_space": "/",
            "rib_long_material": "/",
            "panel_length": "/",
            "concrete_grade": "/",
            "truss_beam_type": "321型贝雷梁",
            "beam_space": "6@900",
            "beam_cant": "0,0",
            "beam_first_space": "2700",
            "beam_material": "/",
            "beam_sec": "/",
        }
    },

    # ---- 环境荷载（wind/water/wave 三层合一）----
    "environment": {
        "wind": {
            "spec": "《公路桥梁抗风设计规范》",
            "design_wind_speed": "24.5",
            "ground_category": "A:海面、海岸、开阔水面",
            "terrain_factor": "1.0",
            "girder_height": "6.4",
            "transverse_coeff": "1.17",
            "formula": "Ud = kf·kt·kh·U10",
        },
        "water": {
            "spec": "《港口工程荷载规范》",
            "flow_velocity": "1.025",
        },
        "wave": {},
    },

    # ---- 连接形式（简单字符串，与 Excel 读取一致） ----
    "connection": {
        "桥面系横肋与纵梁": "铰接",
        "纵梁与支撑架": "连接1",
        "纵梁与横向分配梁": "连接1",
        "桩顶分配梁与桩顶": "铰接",
        "贝雷梁释放梁端约束": "0000000,0000100",
        "桩底": "111111",
        "土弹簧分层厚度(m)": "/",
    },
    "connection_stiffness": {},

    # ---- 荷载----
    "vehicle": {},
    "move_load_cases": {},
    "static_load_cases": {},
    "load_combination": {},
    "section": {},
    "quick_defaults": {},

    # ---- 地面线 ----
    "ground_line": [],
}


def load_all_params(excel_path=None, row=3):
    """从 栈桥参数表.xlsx 读取指定行参数，返回缓存格式 dict。

    所有读取均通过表头名映射，禁止硬编码列号。
    row: 主表数据行号（默认3，兼容单项目调用）
    """
    # 默认值（缓存格式）
    basic = {"span": "", "bridge_width": "", "water_level": "", "deck_level": "", "x_origin": ""}
    bridge_type = "上承式桁架梁栈桥"
    project_name = ""
    project_no = ""
    bridge_deck_type_ref = "1"
    deck_trans_ecc = ""
    deck_end_ext = "0,0"
    rib_trans_space = ""
    beam_material = ""
    beam_sec = ""
    beam_sec_detail = {}
    beam_first_space = ""
    quick_defaults = {}
    connection = {}
    connection_stiffness = {}
    deck_system = {}
    beam = {}
    env_wind = {}
    env_water = {}
    env_wave = {}
    move_load_cases = []
    static_load_cases = []
    load_combinations = []
    substructure_defaults = {}
    support_nos = ""
    pile_lens = ""
    soil_labels_csv = ""
    vehicle_list = []
    seq_to_name = {}
    section_db = {}
    ground_line = []

    if excel_path is None:
        excel_path = get_excel_path()
    if not excel_path or not os.path.exists(excel_path):
        print(f"[excel_io] 警告: 找不到 {excel_path}，使用内置默认值")
        return _build_cache_result(project_name, project_no, bridge_type, basic, bridge_deck_type_ref,
                                   deck_trans_ecc, deck_end_ext, deck_system, beam, beam_first_space,
                                   beam_material, beam_sec, beam_sec_detail, rib_trans_space,
                                   connection, connection_stiffness, env_wind, env_water, env_wave,
                                   vehicle_list, move_load_cases, static_load_cases, load_combinations,
                                   section_db, support_nos, pile_lens, soil_labels_csv, substructure_defaults,
                                   quick_defaults, ground_line, [])

    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
    except Exception as e:
        print(f"[excel_io] 无法打开 Excel: {e}，使用内置默认值")
        return _build_cache_result(project_name, project_no, bridge_type, basic, bridge_deck_type_ref,
                                   deck_trans_ecc, deck_end_ext, deck_system, beam, beam_first_space,
                                   beam_material, beam_sec, beam_sec_detail, rib_trans_space,
                                   connection, connection_stiffness, env_wind, env_water, env_wave,
                                   vehicle_list, move_load_cases, static_load_cases, load_combinations,
                                   section_db, support_nos, pile_lens, soil_labels_csv, substructure_defaults,
                                   quick_defaults, ground_line, [])

    # ======================== Sheet "栈桥参数" — 主表 ========================
    ws1 = _find_sheet(wb, ["栈桥参数", "1"])
    if ws1 is not None:
        hm = _build_header_map(ws1, header_row=2)
        R = lambda hdr: _cell(ws1, row, hdr, hm)

        # --- 基本参数 ---
        project_name = _safe_str(R("项目名称"))
        project_no = _safe_str(R("编号"))
        bridge_type = _safe_str(R("栈桥类型")) or bridge_type
        basic["span"] = _safe_str(R("计算跨径(m)"))
        basic["bridge_width"] = _safe_str(R("计算宽度(m)"))
        basic["water_level"] = _safe_str(R("设防水位(m)"))
        basic["deck_level"] = _safe_str(R("桥面高程(m)"))
        basic["x_origin"] = _safe_str(R("栈桥起始定位线(0#)x坐标(m)"))
        support_nos = _safe_str(R("下部结构编号"))
        pile_lens = _safe_str(R("桩长(m)"))
        soil_labels_csv = _safe_str(R("详细土层地形信息"))

        # --- 桥面系参数区块 ---
        bridge_deck_type_ref = _safe_str(R("桥面系编号"))
        deck_trans_ecc = _safe_str(R("桥面左端相对线路中心线横向间距(mm)"))
        deck_end_ext = _safe_str(R("桥面两端相对桥面系横肋纵向伸出距离(mm)"))

        # --- 纵梁参数区块 ---
        beam["truss_beam_type"] = _safe_str(R("桁架纵梁类型"))
        beam["beam_space"] = _safe_str(R("纵梁横向布置(mm)"))
        beam_first_space = _safe_str(R("纵梁第一排相对线路中心线横向间距(mm)"))
        beam["beam_cant"] = _safe_str(R("纵梁两端纵向伸出距离(mm)"))

        # --- 桥面系横肋纵向布置(主表) ---
        rib_trans_space = _safe_str(R("桥面系横肋纵向布置(mm)"))

        # --- 纵梁材质 / 型钢截面 ---
        beam_material = _safe_str(R("纵梁材质"))
        beam_sec = _safe_str(R("纵梁截面"))

        # --- 快捷设置（通行荷载/施工静载，"/" 视为空）---
        quick = {}
        if "通行荷载" in hm:
            val = _safe_str(R("通行荷载"))
            if val and val != "/":
                quick["traffic"] = val
        if "施工静载" in hm:
            val = _safe_str(R("施工静载"))
            if val and val != "/":
                quick["static"] = val
        if quick:
            quick_defaults = quick

        # --- 移动荷载工况区块 ---
        if "移动荷载工况名称" in hm:
            move_veh = [v for v in _split_and_strip_by_comma(_safe_str(R("移动荷载车辆/设备"))) if v not in ("", "/")]
            if move_veh:
                move_names = _split_and_strip_by_comma(_safe_str(R("移动荷载工况名称")))
                move_ecc = _split_and_strip_by_comma(_safe_str(R("车道偏心(m)")))
                for j, mn in enumerate(move_names):
                    move_load_cases.append({
                        "name": mn,
                        "vehicle": move_veh[j] if j < len(move_veh) else "",
                        "lane_ecc": move_ecc[j] if j < len(move_ecc) else "0.0",
                    })

        # --- 静载工况区块 ---
        if "静载工况名称" in hm:
            def _parse_bracket_name(s):
                s = s.strip()
                if s.startswith("[") and s.endswith("]"):
                    return [x.strip() for x in s[1:-1].split(",") if x.strip()]
                return [s]
            def _parse_bracket_list(s):
                s = s.strip()
                if s.startswith("[") and s.endswith("]"):
                    inner = s[1:-1].strip()
                    if not inner:
                        return []
                    return [int(x.strip()) for x in inner.split(",") if x.strip()]
                return []
            static_veh = [v for v in _split_and_strip_by_comma(_safe_str(R("静载车辆/设备"))) if v not in ("", "/")]
            if static_veh:
                static_name_groups = _split_bracket(_safe_str(R("静载工况名称")))
                static_types_raw = _split_bracket(_safe_str(R("静载工况类型")))
                static_types = [[t.strip() for t in x.strip("[]").split(",") if t.strip()] for x in static_types_raw]
                static_left = [_parse_bracket_list(x) for x in _split_bracket(_safe_str(R("跨间左侧加载")))]
                static_mid = [_parse_bracket_list(x) for x in _split_bracket(_safe_str(R("跨中加载")))]
                static_right = [_parse_bracket_list(x) for x in _split_bracket(_safe_str(R("跨间右侧加载")))]
                static_pile = [_parse_bracket_list(x) for x in _split_bracket(_safe_str(R("桩顶加载")))]
                for j, name_group in enumerate(static_name_groups):
                    veh = static_veh[j] if j < len(static_veh) else ""
                    types = static_types[j] if j < len(static_types) else []
                    left = static_left[j] if j < len(static_left) else []
                    mid = static_mid[j] if j < len(static_mid) else []
                    right = static_right[j] if j < len(static_right) else []
                    pile = static_pile[j] if j < len(static_pile) else []
                    for sn in _parse_bracket_name(name_group):
                        static_load_cases.append({
                            "name": sn, "vehicle": veh, "types": types,
                            "span_left": left, "span_mid": mid, "span_right": right, "pile": pile,
                        })

        # --- 连接形式区块 ---
        for conn_hdr in ("桥面系横肋与纵梁", "纵梁与支撑架", "纵梁与横向分配梁",
                         "桩顶分配梁与桩顶", "贝雷梁释放梁端约束", "桩底", "土弹簧分层厚度(m)"):
            v = R(conn_hdr)
            if v is not None:
                connection[conn_hdr] = _safe_str(v)

        # --- 环境/水流/波浪参数 ---
        env_wind["spec"] = _safe_str(R("风荷载规范"))
        env_wind["design_wind_speed"] = _safe_str(R("设计风速(m/s)"))
        env_wind["ground_category"] = _safe_str(R("地表分类"))
        env_wind["terrain_factor"] = _safe_str(R("地形条件系数"))
        env_wind["girder_height"] = _safe_str(R("主梁基准高度(m)"))
        env_wind["transverse_coeff"] = _safe_str(R("主梁横向力系数"))
        env_wind["formula"] = _safe_str(R("设计基准风速计算公式"))
        env_water["spec"] = _safe_str(R("水流力规范"))
        env_water["flow_velocity"] = _safe_str(R("设计水流速(m/s)") or R("设计流速(m/s)"))
        wave_morison = _safe_str(R("Morison波浪力参数-周期(s)/波长(m)/波高(m)"))
        if wave_morison and wave_morison != "/":
            parts = [x.strip() for x in wave_morison.split(",")]
            if len(parts) >= 3:
                env_wave["wave_period"] = parts[0]
                env_wave["wave_length"] = parts[1]
                env_wave["wave_height"] = parts[2]
            elif len(parts) == 2:
                env_wave["wave_period"] = parts[0]
                env_wave["wave_height"] = parts[1]
            elif len(parts) == 1:
                env_wave["wave_period"] = parts[0]

        # --- 荷载组合区块（从主表读取） ---
        comb_names_str = _safe_str(R("组合名"))
        comb_types_str = _safe_str(R("组合类型"))
        comb_cases_str = _safe_str(R("荷载工况"))
        comb_gamma_str = _safe_str(R("分项系数"))
        comb_psi_str = _safe_str(R("组合值系数"))
        comb_names = [v for v in _split_and_strip_by_comma(comb_names_str) if v not in ("", "/")]
        if comb_names:
            comb_types = _split_and_strip_by_comma(comb_types_str)
            comb_cases_groups = _split_bracket(comb_cases_str) if comb_cases_str else []
            comb_gamma_groups = _split_bracket(comb_gamma_str) if comb_gamma_str else []
            comb_psi_groups = _split_bracket(comb_psi_str) if comb_psi_str else []
            for j, cn in enumerate(comb_names):
                ctype = comb_types[j] if j < len(comb_types) else "相加"
                cases = []
                if j < len(comb_cases_groups):
                    case_names = _split_and_strip_by_comma(comb_cases_groups[j].strip("[]"))
                    gamma_vals = _split_and_strip_by_comma(comb_gamma_groups[j].strip("[]")) if j < len(comb_gamma_groups) else []
                    psi_vals = _split_and_strip_by_comma(comb_psi_groups[j].strip("[]")) if j < len(comb_psi_groups) else []
                    for ci, case_name in enumerate(case_names):
                        gamma = gamma_vals[ci] if ci < len(gamma_vals) else "1.0"
                        psi = psi_vals[ci] if ci < len(psi_vals) else "1.0"
                        cases.append({"case_name": case_name, "gamma": gamma, "psi": psi})
                load_combinations.append({"name": cn, "type": ctype, "cases": cases})


    # ======================== Sheet "连接形式" ========================
    ws_conn = _find_sheet(wb, ["连接形式"])
    if ws_conn is not None:
        conn_hm = _build_header_map(ws_conn, header_row=2)
        for r in range(3, ws_conn.max_row + 1):
            name = _safe_str(_cell(ws_conn, r, "连接形式", conn_hm))
            if not name:
                continue
            connection_stiffness[name] = {
                "SDx": _safe_str(_cell(ws_conn, r, "SDx(kN/m)", conn_hm)),
                "SDy": _safe_str(_cell(ws_conn, r, "SDy(kN/m)", conn_hm)),
                "SDz": _safe_str(_cell(ws_conn, r, "SDz(N/mm)", conn_hm)),
                "SRx": _safe_str(_cell(ws_conn, r, "SRx(N·mm/rad)", conn_hm)),
                "SRy": _safe_str(_cell(ws_conn, r, "SRy(N·mm/rad)", conn_hm)),
                "SRz": _safe_str(_cell(ws_conn, r, "SRz(N·mm/rad)", conn_hm)),
                "NSDx": _safe_str(_cell(ws_conn, r, "NSDx(N/mm)", conn_hm)),
            }
        # 连接刚度记录数（不逐行打印）


    # ======================== Sheet "上部结构-桥面系库" ========================
    ws_deck = _find_sheet(wb, ["上部结构-桥面系库", "桥面板库", "桥面板", "桥面"])
    if ws_deck is not None:
        deck_hm = _build_header_map(ws_deck, header_row=1)
        deck_db = {}
        for r in range(2, ws_deck.max_row + 1):
            seq = _cell(ws_deck, r, "桥面系编号", deck_hm) or _cell(ws_deck, r, "编号", deck_hm)
            if seq is None:
                continue
            seq = int(seq) if str(seq).isdigit() else seq
            deck_db[seq] = {
                "deck_type": _safe_str(_cell(ws_deck, r, "桥面系类型", deck_hm)) or "单层钢面板",
                "deck_thickness": _safe_str(_cell(ws_deck, r, "桥面板厚(mm)", deck_hm)),
                "rib_trans_sec": _safe_str(_cell(ws_deck, r, "横肋截面", deck_hm)),
                "rib_long_sec": _safe_str(_cell(ws_deck, r, "纵肋截面", deck_hm)),
                "rib_long_space": _safe_str(_cell(ws_deck, r, "纵肋横向布置(mm)", deck_hm)),
                "panel_length": _safe_str(_cell(ws_deck, r, "单块面板长度(m)", deck_hm)),
                "concrete_grade": _safe_str(_cell(ws_deck, r, "混凝土标号", deck_hm)),
                "deck_material": _safe_str(_cell(ws_deck, r, "桥面板材质", deck_hm)),
                "rib_trans_material": _safe_str(_cell(ws_deck, r, "横肋材质", deck_hm)),
                "rib_long_material": _safe_str(_cell(ws_deck, r, "纵肋材质", deck_hm)),
            }

        ref = bridge_deck_type_ref
        try:
            ref = int(ref)
        except:
            ref = 1
        deck_system = deck_db.get(ref, deck_db.get(1, {}))


    # ======================== 纵梁 — 从主表直接读取 ========================
    # 桁架纵梁类型/布置已在主表区块中读取至 params["beam"]

    # ======================== Sheet "下部结构库" ========================
    ws_sub = _find_sheet(wb, ["下部结构库", "下部结构", "下部"])
    if ws_sub is not None:
        sub_hm = _build_header_map(ws_sub, header_row=1)
        for r in range(2, ws_sub.max_row + 1):
            seq = _safe_str(_cell(ws_sub, r, "下部结构编号", sub_hm) or _cell(ws_sub, r, "支撑编号", sub_hm))
            if not seq:
                continue
            sup = {
                "type": _safe_str(_cell(ws_sub, r, "结构类型", sub_hm)),
                "pile_type": _safe_str(_cell(ws_sub, r, "桩类型", sub_hm)),
                "pile_sec": _safe_str(_cell(ws_sub, r, "桩截面", sub_hm)),
                "pile_trans_space": _safe_str(_cell(ws_sub, r, "桩横向布置(mm)", sub_hm)),
                "pile_long_space": _safe_str(_cell(ws_sub, r, "桩纵向布置(mm)", sub_hm)),
                "expansion_len": _safe_str(_cell(ws_sub, r, "伸缩缝(mm)", sub_hm)),
                "dist_trans_sec": _safe_str(_cell(ws_sub, r, "分配梁(横)截面", sub_hm)),
                "dist_trans_len": _safe_str(_cell(ws_sub, r, "分配梁(横)长度(mm)", sub_hm)),
                "dist_trans_offset": _safe_str(_cell(ws_sub, r, "分配梁(横)左端相对线路中心线横向间距(mm)", sub_hm)),
                "dist_long_center_offset": _safe_str(_cell(ws_sub, r, "分配梁(纵)左端相对联中心线横向间距(mm)", sub_hm)),
                "dist_long_sec": _safe_str(_cell(ws_sub, r, "分配梁(纵)截面", sub_hm)),
                "dist_long_len": _safe_str(_cell(ws_sub, r, "分配梁(纵)长度(mm)", sub_hm)),
                "brace_form": _safe_str(_cell(ws_sub, r, "联结系形式", sub_hm)),
                "brace_sec": _safe_str(_cell(ws_sub, r, "联结系截面", sub_hm)),
                "brace_sec_tilt": _safe_str(_cell(ws_sub, r, "联结系斜杆/竖杆截面", sub_hm)),
                "brace_space": _safe_str(_cell(ws_sub, r, "联结系间距(mm)", sub_hm)),
                "brace_height": _safe_str(_cell(ws_sub, r, "联结系高度(mm)", sub_hm)),
                "pile_material": _safe_str(_cell(ws_sub, r, "桩材质", sub_hm)),
                "dist_trans_material": _safe_str(_cell(ws_sub, r, "分配梁(横)材质", sub_hm)),
                "dist_long_material": _safe_str(_cell(ws_sub, r, "分配梁(纵)材质", sub_hm)),
                "brace_material": _safe_str(_cell(ws_sub, r, "联结系材质", sub_hm)),
            }
            sup["pile_len"] = "/"
            substructure_defaults[seq] = sup

    # ======================== Sheet "车辆设备库" ========================
    _ref_veh_names = set()
    for ml in move_load_cases:
        v = ml.get("vehicle", "").strip()
        if v:
            _ref_veh_names.add(v)
    for sl in static_load_cases:
        v = sl.get("vehicle", "").strip()
        if v:
            _ref_veh_names.add(v)

    ws_veh = _find_sheet(wb, ["车辆设备库", "车辆", "设备"])
    if ws_veh is not None:
        veh_hm = _build_header_map(ws_veh, header_row=2)
        _all_veh_rows = []
        for r in range(3, ws_veh.max_row + 1):
            seq = _cell(ws_veh, r, "车辆设备编号", veh_hm) or _cell(ws_veh, r, "序号", veh_hm)
            if seq is None:
                continue
            name_val = _safe_str(_cell(ws_veh, r, "车辆/设备名称", veh_hm))   # 名称（主表引用此列）
            vehicle_tag = _safe_str(_cell(ws_veh, r, "车辆/设备", veh_hm))     # 标签（区分自定义/标准）
            lt_val = _safe_str(_cell(ws_veh, r, "车辆荷载类型", veh_hm))
            if not name_val or not lt_val:
                continue
            veh_params = []
            if lt_val in ("一般车辆", "自定义-一般车辆"):
                veh_hm2 = _build_header_map(ws_veh, header_row=2)
                axle_dist = _safe_str(_cell(ws_veh, r, "车轮间距/轴距(m)", veh_hm))
                loads_csv = _safe_str(_cell(ws_veh, r, "轴重(kN)", veh_hm2))
                dists_csv = _safe_str(_cell(ws_veh, r, "距前轴(m)", veh_hm2) or
                               _cell(ws_veh, r, "距前轴(mm)", veh_hm2))
                veh_params = [axle_dist, loads_csv, dists_csv]
            else:
                for c in range(5, min(ws_veh.max_column + 1, 37)):
                    cv = ws_veh.cell(r, c).value
                    if cv is not None and str(cv).strip() not in ("", "/"):
                        veh_params.append(str(cv).strip())
            _all_veh_rows.append((len(_all_veh_rows) + 1, name_val, lt_val, vehicle_tag, veh_params))

        seq_to_name = {str(row_seq): name for row_seq, name, *_ in _all_veh_rows}
        _resolved_names = set()
        for v in _ref_veh_names:
            _resolved_names.add(seq_to_name.get(v, v))

        for row_seq, name_val, lt_val, vehicle_tag, veh_params in _all_veh_rows:
            if name_val not in _resolved_names:
                continue
            vehicle_list.append({
                "name": name_val,
                "vehicle_tag": vehicle_tag if vehicle_tag else name_val,
                "load_type": lt_val,
                "display_name": name_val,
                "params": veh_params,
            })

    # ======================== Sheet "截面库" ========================
    ws_sec = _find_sheet(wb, ["截面库", "截面"])
    if ws_sec is not None:
        sec_hm = _build_header_map(ws_sec, header_row=1)
        sec_col_to_hdr = {c: h for h, c in sec_hm.items()}
        for r in range(2, ws_sec.max_row + 1):
            seq = _cell(ws_sec, r, "截面编号", sec_hm) or _cell(ws_sec, r, "序号", sec_hm)
            if seq is None:
                continue
            name_val = _safe_str(_cell(ws_sec, r, "截面名称", sec_hm))
            type_val = _safe_str(_cell(ws_sec, r, "截面类型", sec_hm))
            material_val = _safe_str(_cell(ws_sec, r, "材料", sec_hm))
            if not name_val or not type_val:
                continue
            p = {}
            for c in range(4, ws_sec.max_column + 1):
                v = ws_sec.cell(r, c).value
                hdr = sec_col_to_hdr.get(c, f"col{c}")
                if hdr == "材料":
                    continue
                if v is not None and str(v).strip() not in ("", "/"):
                    p[hdr] = v
            rec = {"type": type_val, "name": name_val, "params": p}
            section_db[int(seq)] = rec

        # 解析截面引用（编号→名称+detail）
        ref_keys = ["pile_sec", "dist_trans_sec", "dist_long_sec", "brace_sec", "brace_sec_tilt"]
        for sd in substructure_defaults.values():
            for rk in ref_keys:
                if rk in sd:
                    resolved = _resolve_section_ref(sd[rk], section_db)
                    if isinstance(resolved, dict):
                        sd[rk + "_detail"] = resolved
                        sd[rk] = resolved.get("section_name", "")
                    else:
                        sd[rk] = resolved
        for rk in ("rib_trans_sec", "rib_long_sec"):
            if rk in deck_system:
                resolved = _resolve_section_ref(deck_system[rk], section_db)
                if isinstance(resolved, dict):
                    deck_system[rk + "_detail"] = resolved
                    deck_system[rk] = resolved.get("section_name", "")
                else:
                    deck_system[rk] = resolved
        if beam_sec:
            _bs_r = _resolve_section_ref(beam_sec, section_db)
            if isinstance(_bs_r, dict):
                beam_sec_detail = _bs_r
                beam_sec = _bs_r.get("section_name", "")
            else:
                beam_sec = _bs_r

    # ======================== 后处理：解析车辆引用（序号→名称） ========================
    if seq_to_name:
        for ml in move_load_cases:
            veh = ml.get("vehicle", "")
            if veh in seq_to_name:
                ml["vehicle"] = seq_to_name[veh]
        for sl in static_load_cases:
            veh = sl.get("vehicle", "")
            if veh in seq_to_name:
                sl["vehicle"] = seq_to_name[veh]

    wb_sheets = list(wb.sheetnames)
    wb.close()

    # ======================== 构建缓存格式结果 ========================
    return _build_cache_result(project_name, project_no, bridge_type, basic, bridge_deck_type_ref,
                               deck_trans_ecc, deck_end_ext, deck_system, beam, beam_first_space,
                               beam_material, beam_sec, beam_sec_detail, rib_trans_space,
                               connection, connection_stiffness, env_wind, env_water, env_wave,
                               vehicle_list, move_load_cases, static_load_cases, load_combinations,
                               section_db, support_nos, pile_lens, soil_labels_csv, substructure_defaults,
                               quick_defaults, ground_line, wb_sheets)


def _build_cache_result(project_name, project_no, bridge_type, basic, bridge_deck_type_ref,
                        deck_trans_ecc, deck_end_ext, deck_system, beam, beam_first_space,
                        beam_material, beam_sec, beam_sec_detail, rib_trans_space,
                        connection, connection_stiffness, env_wind, env_water, env_wave,
                        vehicle_list, move_load_cases, static_load_cases, load_combinations,
                        section_db, support_nos, pile_lens, soil_labels_csv, substructure_defaults,
                        quick_defaults, ground_line, wb_sheets):
    """将读取的各模块数据组装为统一缓存格式"""

    # ---- substructure：refs + defaults 展开为逐支撑 dict ----
    def _csv(s):
        s = (s or "").replace("，", ",")
        return [x.strip() for x in s.split(",") if x.strip()]

    nos = _csv(support_nos)
    lens = _csv(pile_lens)
    soil_lbls = _csv(soil_labels_csv)
    soil_label = soil_lbls[0] if soil_lbls else ""
    substructure = {}
    for i, no in enumerate(nos):
        src = substructure_defaults.get(str(no), {})
        init_type = src.get("type", "单排桩")
        entry = get_default_support_data(init_type)
        SEC_REF_KEYS = {"pile_sec", "dist_trans_sec", "dist_long_sec", "brace_sec", "brace_sec_tilt"}
        for k, v in src.items():
            if v is not None:
                if k.endswith("_detail") and isinstance(v, dict):
                    entry[k] = copy.deepcopy(v)
                elif str(v).strip() == "/":
                    entry[k] = "/"
                elif str(v).strip() == "" and k in SEC_REF_KEYS:
                    # 截面引用字段为空 → 视同未设置，置 "/"（避免回退模板默认截面）
                    entry[k] = "/"
                elif str(v).strip() not in ("", "/"):
                    entry[k] = str(v).strip()
        entry["type"] = init_type
        if i < len(lens) and lens[i] not in ("", "/"):
            entry["pile_len"] = lens[i]
        entry["soil_data"] = {}
        substructure[str(i)] = entry

    # ---- components ----
    deck_type = deck_system.get("deck_type", "") or "单层钢面板"
    comp = {
        "deck_type": deck_type,
        "deck_thickness": deck_system.get("deck_thickness", ""),
        "deck_material": deck_system.get("deck_material", ""),
        "deck_trans_ecc": deck_trans_ecc,
        "deck_end_ext": deck_end_ext,
        "rib_trans_sec": deck_system.get("rib_trans_sec", ""),
        "rib_trans_sec_detail": copy.deepcopy(deck_system.get("rib_trans_sec_detail", {})),
        "rib_trans_space": rib_trans_space,
        "rib_trans_material": deck_system.get("rib_trans_material", ""),
        "rib_long_sec": deck_system.get("rib_long_sec", ""),
        "rib_long_sec_detail": copy.deepcopy(deck_system.get("rib_long_sec_detail", {})),
        "rib_long_space": deck_system.get("rib_long_space", ""),
        "rib_long_material": deck_system.get("rib_long_material", ""),
        "panel_length": deck_system.get("panel_length", ""),
        "concrete_grade": deck_system.get("concrete_grade", ""),
        "truss_beam_type": beam.get("truss_beam_type", ""),
        "beam_space": beam.get("beam_space", ""),
        "beam_cant": beam.get("beam_cant", ""),
        "beam_first_space": beam_first_space,
        "beam_material": beam_material,
        "beam_sec": beam_sec,
        "beam_sec_detail": copy.deepcopy(beam_sec_detail),
    }

    # 根据 deck_type 清除不适用的字段，避免误用
    if deck_type == "单层钢面板":
        # 单层钢面板无纵肋
        for k in ("rib_long_sec", "rib_long_sec_detail", "rib_long_space", "rib_long_material"):
            comp[k] = "" if k != "rib_long_sec_detail" else {}
        # 单层钢面板无混凝土参数
        comp["concrete_grade"] = ""
        comp["panel_length"] = ""
    elif deck_type == "混凝土桥面板":
        # 混凝土桥面板无横肋/纵肋
        for k in ("rib_trans_sec", "rib_trans_sec_detail", "rib_trans_space", "rib_trans_material",
                   "rib_long_sec", "rib_long_sec_detail", "rib_long_space", "rib_long_material"):
            comp[k] = "" if k.endswith("_sec") or k.endswith("_space") or k.endswith("_material") else {}

    # ---- environment ----
    environment = {
        "wind": {
            "spec": env_wind.get("spec", ""),
            "design_wind_speed": env_wind.get("design_wind_speed", ""),
            "ground_category": env_wind.get("ground_category", ""),
            "terrain_factor": env_wind.get("terrain_factor", ""),
            "girder_height": env_wind.get("girder_height", ""),
            "transverse_coeff": env_wind.get("transverse_coeff", ""),
            "formula": env_wind.get("formula", ""),
        },
        "water": {
            "spec": env_water.get("spec", ""),
            "flow_velocity": env_water.get("flow_velocity", ""),
        },
        "wave": {
            "spec": env_wave.get("spec", ""),
            "wave_height": env_wave.get("wave_height", ""),
            "wave_period": env_wave.get("wave_period", ""),
            "wave_length": env_wave.get("wave_length", ""),
        },
    }

    # ---- vehicle：list → dict ----
    vehicle = {}
    for vd in vehicle_list:
        n = vd.get("name", "")
        if n:
            vehicle[n] = {
                "vehicle_var": n,
                "vehicle_tag": vd.get("vehicle_tag", n),
                "vehicle_type": vd.get("load_type", ""),
                "vehicle_params": vd.get("params", []),
            }

    # ---- move_load_cases：list → dict ----
    ml_dict = {}
    for ml in move_load_cases:
        n = ml.get("name", "")
        if n:
            try:
                ecc = float(ml.get("lane_ecc", 0) or 0)
            except (TypeError, ValueError):
                ecc = 0.0
            ml_dict[n] = {"vehicle": ml.get("vehicle", ""), "lane_ecc": ecc, "direction": "往返", "factor": 1.0}

    # ---- static_load_cases：展开 type ----
    sl_dict = {}
    for sl in static_load_cases:
        veh = sl.get("vehicle", "")
        if not veh:
            continue
        types = sl.get("types", []) or [""]
        base_name = sl.get("name", "") or veh
        for t in types:
            cn = f"{base_name}{t}" if t else base_name
            _dup = 2
            while cn in sl_dict:
                cn = f"{base_name}{t}{_dup}"
                _dup += 1
            sl_dict[cn] = {
                "vehicle": veh, "type": t,
                "跨间左侧": list(sl.get("span_left", [])),
                "跨中": list(sl.get("span_mid", [])),
                "跨间右侧": list(sl.get("span_right", [])),
                "桩顶": list(sl.get("pile", [])),
            }

    # ---- load_combination：dict 化 ----
    lc_dict = {}
    for lc in load_combinations:
        n = lc.get("name", "")
        if not n:
            continue
        cases = {}
        for c in lc.get("cases", []):
            cn = c.get("case_name", "")
            if not cn:
                continue
            try:
                g = float(c.get("gamma", 1.0))
            except (TypeError, ValueError):
                g = 1.0
            try:
                ps = float(c.get("psi", 1.0))
            except (TypeError, ValueError):
                ps = 1.0
            cases[cn] = {"gamma": g, "psi": ps}
        lc_dict[n] = {"type": lc.get("type", "相加"), "cases": cases}

    # ---- section：int key → 名称 key ----
    section = {}
    for seq in sorted(section_db.keys(), key=lambda x: int(x) if str(x).isdigit() else 0):
        rec = section_db[seq]
        if not isinstance(rec, dict):
            continue
        key = _norm_sec_name(rec.get("name", "")) or str(seq)
        if key not in section:
            section[key] = rec

    # ---- 组装结果 ----
    print_load_summary(project_name, project_no, bridge_type, basic, substructure,
                               beam, beam_material, beam_sec, deck_system, quick_defaults,
                               connection, env_wind, env_water, env_wave,
                               move_load_cases, static_load_cases, load_combinations,
                               vehicle_list, section_db)

    return {
        "project_name": project_name,
        "project_no": project_no,
        "bridge_type": bridge_type,
        "basic": basic,
        "substructure": substructure,
        "components": {bridge_type: comp},
        "bridge_deck_type_ref": bridge_deck_type_ref,
        "connection": connection,
        "connection_stiffness": connection_stiffness,
        "environment": environment,
        "vehicle": vehicle,
        "move_load_cases": ml_dict,
        "static_load_cases": sl_dict,
        "load_combination": lc_dict,
        "section": section,
        "soil_label": soil_label,
        "quick_defaults": quick_defaults,
        "ground_line": ground_line,
        "_wb_sheets": wb_sheets,
    }


def print_load_summary(project_name, project_no, bridge_type, basic, substructure,
                       beam, beam_material, beam_sec, deck_system, quick_defaults,
                       connection, env_wind, env_water, env_wave,
                       move_load_cases, static_load_cases, load_combinations,
                       vehicle_list, section_db):
    """简洁打印加载完成的数据结构概览"""
    print(f"  项目: {project_name}-{project_no}  桥型：{bridge_type}")
    print(f"  基本: 跨径={basic.get('span','')}  宽度={basic.get('bridge_width','')}  设防水位={basic.get('water_level','')}  桥面高程={basic.get('deck_level','')}")
    if basic.get("x_origin"):
        print(f"        起始定位线x0={basic.get('x_origin','')}")

    # 下部结构逐桩打印
    if substructure:
        print(f"  下部结构({len(substructure)}个):")
        for k in sorted(substructure, key=lambda x: int(x) if str(x).isdigit() else 0):
            s = substructure[k]
            print(f"    {k}#: {s.get('type',''):6s}  桩={s.get('pile_sec','/')}"
                  f"  桩横={s.get('pile_trans_space','/')}  桩纵={s.get('pile_long_space','/')}"
                  f"  brace={s.get('brace_form','-')}  offset={s.get('dist_trans_offset','/')}")

    if "型钢" in bridge_type:
        print(f"  纵梁: 型钢  材质={beam_material}  截面={beam_sec}")
    elif beam.get("truss_beam_type"):
        print(f"  纵梁: {beam.get('truss_beam_type','')}  布置={beam.get('beam_space','')}")
        if beam.get("beam_cant"):
            print(f"        伸出={beam.get('beam_cant','')}  第一排距中心={beam.get('beam_first_space','')}")

    if deck_system:
        print(f"  桥面系: 类型={deck_system.get('deck_type', '单层钢面板')}  面板厚={deck_system.get('deck_thickness','')}mm")
        if deck_system.get('deck_type') == '混凝土桥面板':
            print(f"          面板长度={deck_system.get('panel_length','')}m  混凝土标号={deck_system.get('concrete_grade','')}")
        elif deck_system.get('deck_type') == '双层钢面板':
            _seg = deck_system.get('seg_enabled', False)
            print(f"          横肋={deck_system.get('rib_trans_sec','')}  纵肋={deck_system.get('rib_long_sec','')}"
                  + (f"  分节={deck_system.get('panel_length','')}m" if _seg else ""))
        else:
            print(f"          横肋={deck_system.get('rib_trans_sec','')}  纵肋={deck_system.get('rib_long_sec','')}")

    if quick_defaults:
        print(f"  车辆荷载快捷设置: 通行荷载={quick_defaults.get('traffic', '/')}  施工静载={quick_defaults.get('static', '/')}")

    if connection:
        print(f"  连接形式: {connection}")

    env_line = f"  环境: 风速={env_wind.get('design_wind_speed','')}  水流速={env_water.get('flow_velocity','')}"
    wave_period = env_wave.get('wave_period', '')
    wave_length = env_wave.get('wave_length', '')
    wave_height = env_wave.get('wave_height', '')
    if wave_period and wave_period != "/":
        env_line += f"  波浪周期={wave_period}  波长={wave_length}  波高={wave_height}"
    else:
        env_line += "  不计算波浪力"
    print(env_line)

    print(f"  车辆：{len(vehicle_list)}辆  组合：{len(load_combinations)}个")
    print(f"\n{'='*60}")


def load_all_projects(excel_path=None):
    """从 栈桥参数表.xlsx 读取所有项目，返回 projects_dict/project_groups/last_proj_key。

    每行数据以 "{项目名称}{编号}" 为 key 存入 projects_dict。
    返回值已经是 cache 格式（统一数据格式）。
    """
    if excel_path is None:
        excel_path = get_excel_path()
    if not excel_path or not os.path.exists(excel_path):
        print(f"[excel_io] 警告: 找不到 {excel_path}，使用单项目默认值")
        p = load_all_params(excel_path, row=3)
        key = p.get("project_name", "") + p.get("project_no", "")
        if not key:
            key = "默认项目"
        return {"projects_dict": {key: p}, "project_groups": {p.get("project_name", "默认"): [p.get("project_no", "")]}, "last_proj_key": key}

    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
    except Exception as e:
        print(f"[excel_io] 无法打开 Excel: {e}，使用单项目默认值")
        p = load_all_params(excel_path, row=3)
        key = p.get("project_name", "") + p.get("project_no", "")
        if not key:
            key = "默认项目"
        return {"projects_dict": {key: p}, "project_groups": {p.get("project_name", "默认"): [p.get("project_no", "")]}, "last_proj_key": key}

    ws1 = _find_sheet(wb, ["栈桥参数", "1"])
    if ws1 is None:
        wb.close()
        p = load_all_params(excel_path, row=3)
        key = p.get("project_name", "") + p.get("project_no", "")
        if not key:
            key = "默认项目"
        return {"projects_dict": {key: p}, "project_groups": {p.get("project_name", "默认"): [p.get("project_no", "")]}, "last_proj_key": key}

    hm = _build_header_map(ws1, header_row=2)
    has_name_col = "项目名称" in hm

    print(f"\n{'='*60}")
    print(f"[excel_io] 读取文件: {excel_path}")
    print(f"  工作表: {wb.sheetnames}")
    print(f"\n{'='*60}")

    projects_dict = {}
    project_groups = {}
    last_proj_key = ""

    if has_name_col:
        for r in range(3, ws1.max_row + 1):
            pname = _safe_str(_cell(ws1, r, "项目名称", hm))
            pno = _safe_str(_cell(ws1, r, "编号", hm))
            # 跳过无效行：项目名称为空、"/"或编号为"/"（避免伪项目被加载）
            if not pname or pname == "/" or pno == "/":
                continue
            proj_key = pname + pno
            # 用 load_all_params 读取该行（共享 sheet 已被 load_all_params 内部读取）
            p = load_all_params(excel_path, row=r)
            projects_dict[proj_key] = p
            project_groups.setdefault(pname, [])
            if pno not in project_groups[pname]:
                project_groups[pname].append(pno)
            last_proj_key = proj_key
    else:
        # 无"项目名称"列，按单项目处理
        p = load_all_params(excel_path, row=3)
        key = p.get("project_name", "") + p.get("project_no", "")
        if not key:
            key = "默认项目"
        projects_dict[key] = p
        project_groups[p.get("project_name", "默认")] = [p.get("project_no", "")]
        last_proj_key = key

    wb.close()

    if not projects_dict:
        p = load_all_params(excel_path, row=3)
        key = p.get("project_name", "") + p.get("project_no", "")
        if not key:
            key = "默认项目"
        projects_dict[key] = p
        project_groups[p.get("project_name", "默认")] = [p.get("project_no", "")]
        last_proj_key = key

    return {"projects_dict": projects_dict, "project_groups": project_groups, "last_proj_key": last_proj_key}


def get_default_support_data(sup_type):
    """根据类型生成默认字典（硬编码原始值，用于 Excel 未覆盖的新增支撑）"""
    if sup_type in ['重力式桥台', '简易桥台']:
        return {
            "type": sup_type,
            "pile_type": "/",  "pile_sec": "/", "pile_len": "/","pile_trans_space": "/", "pile_long_space": "/",
            "expansion_len": "/",
            "dist_trans_sec": "/","dist_trans_len": "/",
            "dist_long_sec": "/","dist_long_len": "/","dist_trans_offset": "/","dist_long_center_offset": "/",
            "brace_form": "/", "brace_sec": "/","brace_sec_tilt":"/","brace_space": "/", "brace_height": "/",
            "pile_material": "/", "dist_trans_material": "/", "dist_long_material": "/", "brace_material": "/",
        }
    elif sup_type == '单排桩':
        return {
            "type": sup_type,
            "pile_type": "单排桩",  "pile_sec": "630X8", "pile_len": "15","pile_trans_space": "750,2@3000", "pile_long_space": "/", "expansion_len": "/",
            "pile_sec_detail": {"section_type": "O", "section_name": "630X8", "params": {"D": "630", "d": "8"}},
            "dist_trans_sec":"2I40a","dist_trans_len": "7500",
            "dist_trans_sec_detail": {"section_type": "2I", "section_name": "2I40a", "params": {"H": "400", "B": "142", "tw": "10.5", "tf1": "16.5", "C": "100", "tf2": "16.5"}},
            "dist_long_sec": "/","dist_long_len": "/","dist_trans_offset": "3750","dist_long_center_offset": "/",
            "brace_form": "-", "brace_sec": "377X6","brace_sec_tilt":"/","brace_space": "620,/", "brace_height": "2500",
            "brace_sec_detail": {"section_type": "O", "section_name": "377X6", "params": {"D": "377", "d": "6"}},
            "pile_material": "", "dist_trans_material": "", "dist_long_material": "/", "brace_material": "",
        }
    elif sup_type == '制动墩':
        return {
            "type": sup_type,
            "pile_type": "双排桩",  "pile_sec": "630X8", "pile_len": "15","pile_trans_space": "750,2@3000", "pile_long_space": "750,3000", "expansion_len": "200",
            "pile_sec_detail": {"section_type": "O", "section_name": "630X8", "params": {"D": "630", "d": "8"}},
            "dist_trans_sec":"2I40a","dist_trans_len": "7500","dist_trans_offset": "3750",
            "dist_trans_sec_detail": {"section_type": "2I", "section_name": "2I40a", "params": {"H": "400", "B": "142", "tw": "10.5", "tf1": "16.5", "C": "100", "tf2": "16.5"}},
            "dist_long_sec": "2HM588X300","dist_long_len": "4500","dist_long_center_offset": "2250",
            "dist_long_sec_detail": {"section_type": "2HM", "section_name": "2HM588X300", "params": {"H": "588", "B": "300", "tw": "12", "tf1": "20", "C": "300", "tf2": "20"}},
            "brace_form": "X", "brace_sec": "377X6","brace_sec_tilt":"219X6","brace_space": "620,/", "brace_height": "2500",
            "brace_sec_detail": {"section_type": "O", "section_name": "377X6", "params": {"D": "377", "d": "6"}},
            "brace_sec_tilt_detail": {"section_type": "O", "section_name": "219X6", "params": {"D": "219", "d": "6"}},
            "pile_material": "", "dist_trans_material": "", "dist_long_material": "", "brace_material": "",
        }


ALL_PROJECTS = None


def get_projects(excel_path=None):
    """返回 {"projects_dict": {...}, "project_groups": {...}, "last_proj_key": "..."}"""
    global ALL_PROJECTS
    if ALL_PROJECTS is None:
        ALL_PROJECTS = load_all_projects(excel_path)
    return ALL_PROJECTS


# ============================================================
# 写入 — 主入口 + 主表
# ============================================================

def save_trestle_to_excel(ui_dict, mct_path, template_path=None):
    """将 projects_dict（多项目）写回参数表

    所有写入均通过表头名映射，禁止硬编码列号。
    主表按项目名+编号定位行：已有行覆写，新行追加；子 sheet 汇总写入。
    """
    if template_path is None:
        template_path = get_excel_path()
    if not template_path or not os.path.exists(template_path):
        print(f"[excel_io] 错误: 找不到模板 {template_path}")
        return None

    save_dir = os.path.dirname(mct_path)
    save_path = os.path.join(save_dir, "Trestle_ParamTable.xlsx")
    veh_conflicts = []

    try:
        wb = openpyxl.load_workbook(template_path)
    except Exception as e:
        print(f"[excel_io] 无法打开模板: {e}")
        return None

    # ====== 先写截面库，获取截面编号映射 ======
    section_no_map = write_section_sheet(wb, ui_dict)

    # ====== 先写桥面系库，获取桥面系编号映射 ======
    deck_no_map = write_deck_sheet(wb, ui_dict, section_no_map)

    # ====== Sheet "栈桥参数" 主表 ======
    # 按项目名+编号定位行，已有行覆写，新行追加
    ws_main = _find_sheet(wb, ["栈桥参数", "1"])
    if ws_main is not None:
        hm_main = _build_header_map(ws_main, header_row=2)
        # 读取现有行的 (项目名称+编号) → 行号 映射
        existing_rows = {}
        for r in range(3, ws_main.max_row + 1):
            pname = _safe_str(_cell(ws_main, r, "项目名称", hm_main))
            pno = _safe_str(_cell(ws_main, r, "编号", hm_main))
            if pname or pno:
                existing_rows[pname + pno] = r
        next_row = max(existing_rows.values(), default=2) + 1

        for proj_dict in ui_dict.values():
            pname = proj_dict.get("project_name", "")
            pno = proj_dict.get("project_no", "")
            combined = pname + pno
            # 所有项目均为统一 cache 格式（load_all_params 读取时即转换）
            # 用 deck_no_map 更新桥面系编号
            deck_key = _build_deck_dedup_key(proj_dict, section_no_map)
            if deck_key in deck_no_map:
                proj_dict["bridge_deck_type_ref"] = str(deck_no_map[deck_key])

            subs_no_map = write_substructure_sheet(wb, proj_dict, section_no_map)
            if combined in existing_rows:
                # 已有项目：覆写到原行（不写身份列）
                write_sheet_main(wb, proj_dict,
                                  section_no_map=section_no_map,
                                  subs_no_map=subs_no_map,
                                  row=existing_rows[combined])
            else:
                # 新项目：追加到末尾（写入身份列）
                write_sheet_main(wb, proj_dict,
                                  section_no_map=section_no_map,
                                  subs_no_map=subs_no_map,
                                  row=next_row, write_identity=True)
                next_row += 1
            # C2：全项目写 {soil_label}土层表 per-pile 覆盖表（无数据项目自动跳过）
            try:
                write_pile_soil_sheet(wb, proj_dict)
            except Exception as e:
                print(f"[excel_io] 写土层表覆盖表失败: {e}")
    # 共享表
    veh_conflicts = write_vehicle_sheet(wb, ui_dict) or []
    write_connection_sheet(wb, ui_dict)

    # 统一居中对齐
    _center_align_all_cells(wb)

    # 保存到用户路径 + 模板路径
    saved = None
    template_ok = False
    try:
        wb.save(save_path)
        print(f"[excel_io] 已保存: {save_path}")
        saved = save_path
    except PermissionError:
        print(f"[excel_io] 权限错误: {save_path} 被占用")
    except Exception as e:
        print(f"[excel_io] 保存失败: {e}")

    if template_path != save_path:
        try:
            wb.save(template_path)
            print(f"[excel_io] 已覆盖模板: {template_path}")
            template_ok = True
        except PermissionError:
            print(f"[excel_io] 警告: 无法写入模板 {template_path}")

    return (saved, template_ok, veh_conflicts)


def write_sheet_main(wb, d, subs_no_map=None, section_no_map=None, row=3, write_identity=False):
    """写 Sheet '栈桥参数' 指定行（通过表头名映射）

    subs_no_map: 来自 write_substructure_sheet 的 {key: 编号} 映射，
                 用于"下部结构编号"列。
    section_no_map: 来自 write_section_sheet 的 {截面名称: 截面编号} 映射，
                 用于将主表中的截面引用列转换为编号写入（读侧再按编号反查名称）。
    row: 写入的行号（默认3，多项目时递增）
    write_identity: True 时写入项目名称/编号列（新增/另存为时使用）
    """
    ws = _find_sheet(wb, ["栈桥参数", "1"])
    if ws is None:
        return
    R = row
    hm = _build_header_map(ws, header_row=2)

    def wh(hdr, val, hmap=None):
        m = hmap or hm
        col = m.get(hdr)
        if col is not None:
            ws.cell(R, col, _replace_empty_with_slash(val))

    basic = d.get("basic", {})
    substructure = d.get("substructure", {})
    components = None
    for v in d.get("components", {}).values():
        components = v
        break
    env = d.get("environment", {}).get("wind", {})
    water = d.get("environment", {}).get("water", {})
    wave = d.get("environment", {}).get("wave", {})

    # --- 基本参数 ---
    # 项目名称/编号：仅新增/另存为时写入，编辑已有项目时保持 Excel 原值
    if write_identity:
        wh("项目名称", d.get("project_name", ""))
        wh("编号", d.get("project_no", ""))
    wh("栈桥类型", d.get("bridge_type", "") or (list(d.get("components", {}).keys())[0] if d.get("components") else ""))
    wh("计算跨径(m)", basic.get("span", ""))
    wh("计算宽度(m)", basic.get("bridge_width", ""))
    wh("桥面高程(m)", basic.get("deck_level", ""))
    wh("设防水位(m)", basic.get("water_level", ""))
    wh("栈桥起始定位线(0#)x坐标(m)", basic.get("x_origin", ""))

    # 支撑编号 / 桩长 / 详细土层地形信息（使用下部结构库查重后的编号映射）
    if subs_no_map:
        nos = _join_with_commas(str(subs_no_map.get(k, int(k) + 1)) for k in sorted(substructure.keys(), key=lambda x: int(x) if str(x).isdigit() else 0))
    else:
        nos = _join_with_commas(str(int(k) + 1) for k in sorted(substructure.keys(), key=lambda x: int(x) if str(x).isdigit() else 0))
    wh("下部结构编号", nos)
    lens = ",".join(str(substructure[k].get("pile_len", "")) for k in sorted(substructure.keys(), key=lambda x: int(x) if str(x).isdigit() else 0))
    wh("桩长(m)", lens)
    # 详细土层地形信息：仅新增/另存为时写入（后续由地形插件管理）
    if write_identity:
        _sl = str(d.get("soil_label", "")).strip()
        if _sl:
            wh("详细土层地形信息", _sl)

    # --- 桥面系参数区块 ---
    wh("桥面系编号", d.get("bridge_deck_type_ref", "1"))
    wh("桥面左端相对线路中心线横向间距(mm)", components.get("deck_trans_ecc", "2700") if components else "2700")
    wh("桥面两端相对桥面系横肋纵向伸出距离(mm)", components.get("deck_end_ext", "0") if components else "0")

    # --- 纵梁参数区块（统一从 components 读取）---
    wh("桁架纵梁类型", components.get("truss_beam_type", "") if components else "")
    wh("纵梁横向布置(mm)", components.get("beam_space", "") if components else "")
    wh("纵梁第一排相对线路中心线横向间距(mm)", components.get("beam_first_space", "0") if components else "0")
    wh("纵梁两端纵向伸出距离(mm)", components.get("beam_cant", "") if components else "")

    # --- 桥面系横肋纵向布置(主表) ---
    wh("桥面系横肋纵向布置(mm)", components.get("rib_trans_space", "") if components else "")

    # --- 纵梁材质 / 型钢截面 ---
    wh("纵梁材质", components.get("beam_material", "") if components else "")
    # 纵梁截面：先写截面库取编号，主表写截面编号（与桥面系库/下部结构库一致，读侧反查名称）
    _beam_sec = components.get("beam_sec", "") if components else ""
    if section_no_map and _beam_sec and _beam_sec != "/":
        _id_to_name = {str(sid): rec.get("name", "") for sid, rec in (d.get("section") or {}).items() if isinstance(rec, dict)}
        _beam_name = _id_to_name.get(str(_beam_sec), _beam_sec)
        _beam_sec = str(section_no_map.get(_beam_name, _beam_name))
    wh("纵梁截面", _beam_sec)

    # --- 快捷设置（通行荷载/施工静载）：运行时参数，不持久化 ---
    if "通行荷载" in hm:
        wh("通行荷载", "/")
    if "施工静载" in hm:
        wh("施工静载", "/")

    # --- 移动荷载工况区块 ---
    ml = d.get("move_load_cases", {})
    wh("移动荷载工况名称", _join_with_commas(list(ml.keys())))
    wh("移动荷载车辆/设备", _join_with_commas([v.get("vehicle", "") for v in ml.values()]))
    wh("车道偏心(m)", _join_with_commas([str(v.get("lane_ecc", "0")) for v in ml.values()]))

    # --- 静载工况区块 ---
    # 按 vehicle 分组（同一车辆多类型合并）
    sl = d.get("static_load_cases", {})
    veh_groups = {}
    for cname, centry in sl.items():
        veh = centry.get("vehicle", "")
        if veh not in veh_groups:
            veh_groups[veh] = {"vehicle": veh, "types": [], "cases": [],
                               "跨间左侧": [], "跨中": [], "跨间右侧": [], "桩顶": []}
        veh_groups[veh]["types"].append(centry.get("type", ""))
        veh_groups[veh]["cases"].append(cname)
        for pk in ("跨间左侧", "跨中", "跨间右侧", "桩顶"):
            vals = centry.get(pk, [])
            if vals:
                veh_groups[veh][pk].extend(vals)
    for g in veh_groups.values():
        g["types"] = list(dict.fromkeys(g["types"]))  # 去重保序
        for pk in ("跨间左侧", "跨中", "跨间右侧", "桩顶"):
            if g[pk]:
                g[pk] = sorted(set(g[pk]))
    # 同车辆的工况名/类型/位置各写一个 []，车辆名只写一次
    wh("静载工况名称", _join_with_commas([_format_as_bracketed_list(g["cases"]) for g in veh_groups.values()]))
    wh("静载车辆/设备", _join_with_commas([g["vehicle"] for g in veh_groups.values()]))
    wh("静载工况类型", _join_with_commas([_format_as_bracketed_list(g["types"]) for g in veh_groups.values()]))
    wh("跨间左侧加载", _join_with_commas([_format_as_bracketed_list(g["跨间左侧"]) for g in veh_groups.values()]))
    wh("跨中加载", _join_with_commas([_format_as_bracketed_list(g["跨中"]) for g in veh_groups.values()]))
    wh("跨间右侧加载", _join_with_commas([_format_as_bracketed_list(g["跨间右侧"]) for g in veh_groups.values()]))
    wh("桩顶加载", _join_with_commas([_format_as_bracketed_list(g["桩顶"]) for g in veh_groups.values()]))

    # --- 连接形式区块 ---
    conn = d.get("connection", {})
    for conn_hdr in ("桥面系横肋与纵梁", "纵梁与支撑架", "纵梁与横向分配梁",
                     "桩顶分配梁与桩顶", "贝雷梁释放梁端约束", "桩底", "土弹簧分层厚度(m)"):
        val = conn.get(conn_hdr, "/")
        if isinstance(val, dict):
            val = val.get("type", "/")
        wh(conn_hdr, val)

    # --- 环境/水流/波浪参数 ---
    wh("风荷载规范", env.get("spec", env.get("wind_spec", "")))
    wh("设计风速(m/s)", env.get("design_wind_speed", ""))
    wh("地表分类", env.get("ground_category", ""))
    wh("地形条件系数", env.get("terrain_factor", ""))
    wh("主梁基准高度(m)", env.get("girder_height", ""))
    wh("主梁横向力系数", env.get("transverse_coeff", ""))
    wh("设计基准风速计算公式", env.get("formula", env.get("wind_formula", "")))
    wh("水流力规范", water.get("spec", water.get("water_spec", "")))
    wh("设计水流速(m/s)", water.get("flow_velocity", ""))
    wave_parts = []
    if wave.get("wave_period") and wave["wave_period"] != "/":
        wave_parts.append(str(wave["wave_period"]))
    if wave.get("wave_length") and wave["wave_length"] != "/":
        wave_parts.append(str(wave["wave_length"]))
    if wave.get("wave_height") and wave["wave_height"] != "/":
        wave_parts.append(str(wave["wave_height"]))
    wh("Morison波浪力参数-周期(s)/波长(m)/波高(m)", ",".join(wave_parts) if wave_parts else "/")

    # --- 荷载组合区块 ---
    combs = d.get("load_combination", {})
    comb_names = []
    comb_types = []
    comb_cases = []
    comb_gammas = []
    comb_psis = []
    for cname, cdata in combs.items():
        comb_names.append(cname)
        comb_types.append(cdata.get("type", "相加"))
        cases = cdata.get("cases", {})
        case_items = [f"{n}" for n in cases.keys()]
        gamma_items = [str(v.get("gamma", "1.0")) if isinstance(v, dict) else "1.0" for v in cases.values()]
        psi_items = [str(v.get("psi", "1.0")) if isinstance(v, dict) else "1.0" for v in cases.values()]
        comb_cases.append("[" + ",".join(case_items) + "]")
        comb_gammas.append("[" + ",".join(gamma_items) + "]")
        comb_psis.append("[" + ",".join(psi_items) + "]")
    wh("组合名", _join_with_commas(comb_names))
    wh("组合类型", _join_with_commas(comb_types))
    wh("荷载工况", ",".join(comb_cases))
    wh("分项系数", ",".join(comb_gammas))
    wh("组合值系数", ",".join(comb_psis))


# ============================================================
# 写入 — 共享库（辅助函数紧跟使用者）
# ============================================================

def write_section_sheet(wb, d):
    """写 Sheet '截面库' — 查重复用模式

    d: 可以是单个项目 dict，也可以是 projects_dict（多项目）

    Returns:
        section_no_map: dict {截面归一化名称: 截面编号}
    """
    ws = _find_sheet(wb, ["截面库", "截面"])
    if ws is None:
        return {}

    hm = _build_header_map(ws, header_row=1)

    def wh(row, hdr, val):
        col = hm.get(hdr)
        if col is not None:
            ws.cell(row, col, _replace_empty_with_slash(val))

    seq_key = "截面编号" if "截面编号" in hm else "序号"

    # 读取现有行，构建 dedup map: {(name, type): (行号, 序号)}
    existing_map = {}
    max_num = 0
    for r in range(2, ws.max_row + 1):
        seq_val = _cell(ws, r, seq_key, hm)
        if seq_val is None:
            continue
        try:
            num = int(str(seq_val).strip())
        except ValueError:
            continue
        max_num = max(max_num, num)
        name = _safe_str(_cell(ws, r, "截面名称", hm)) or ""
        stype = _safe_str(_cell(ws, r, "截面类型", hm)) or ""
        existing_map[(name, stype)] = (r, num)

    # 收集所有项目的 sections
    all_sections = {}
    projects = d.values() if isinstance(d, dict) and not d.get("section") and not d.get("components") else [d]
    for proj in projects:
        for rec in proj.get("section", {}).values():
            name = rec.get("name", "")
            stype = rec.get("type", "")
            if (name, stype) not in all_sections:
                all_sections[(name, stype)] = rec

    existing_count = len(existing_map)
    new_count = 0
    reuse_count = 0

    for (name, stype), rec in all_sections.items():
        dedup_key = (name, stype)
        params = rec.get("params", {})

        if dedup_key in existing_map:
            reuse_count += 1
            row_idx = existing_map[dedup_key][0]
        else:
            # 新增行：找序号列第一个空行
            new_count += 1
            max_num += 1
            seq_col = hm.get(seq_key)
            row_idx = 2
            if seq_col:
                for r in range(2, ws.max_row + 2):
                    if ws.cell(r, seq_col).value is None:
                        row_idx = r
                        break
            else:
                row_idx = ws.max_row + 1 if ws.max_row >= 2 else 2
            existing_map[dedup_key] = (row_idx, max_num)
            wh(row_idx, seq_key, max_num)
            wh(row_idx, "截面名称", name)
            wh(row_idx, "截面类型", stype)
            # 材料列已移至构件级别（桩材质、分配梁材质等），截面库不再写入
            # if "材料" in hm:
            #     wh(row_idx, "材料", rec.get("material", "Q235"))

        # 更新参数列（不在 params 中的列写 "/"）
        for col, hdr in sorted([(c, h) for h, c in hm.items()], key=lambda x: x[0]):
            if hdr in (seq_key, "截面名称", "截面类型", "材料"):
                continue
            ws.cell(row_idx, col, _replace_empty_with_slash(params.get(hdr)))

    print(f"[截面库] 现有{existing_count}条，新增{new_count}条，复用{reuse_count}条")

    # 构建返回的映射：截面名称 → 编号
    section_no_map = {}
    for (name, stype), (_, num) in existing_map.items():
        section_no_map[name] = num
    return section_no_map


def _dedup_val(val):
    """去重键字段归一化：None/空串统一为 '/'，保证读写两侧格式一致"""
    if val is None or str(val).strip() == "":
        return "/"
    return str(val).strip()


def _build_deck_dedup_key(d, section_no_map=None):
    """构建桥面系去重键，与 write_deck_sheet 中逻辑保持一致。"""
    raw = d.get("components", {})
    # components 可能是 {"1": {实际数据}} 或直接 {实际数据}
    components = None
    for v in raw.values():
        components = v
        break
    if not components:
        return ("", "", "", "", "")
    proj = d.get("project_info", {})
    sec_dict = proj.get("section", d.get("section", {}))
    id_to_name = {sid: rec.get("name", "") for sid, rec in sec_dict.items() if isinstance(rec, dict)}

    deck_type = _dedup_val(components.get("deck_type", "单层钢面板"))
    thickness = _dedup_val(components.get("deck_thickness", ""))
    rib_trans = components.get("rib_trans_sec", "")
    rib_long = components.get("rib_long_sec", "")
    # dedup 统一用截面名称比较（与 write_deck_sheet 读侧一致），不转编号
    rib_trans_name = _dedup_val(id_to_name.get(str(rib_trans), str(rib_trans)))
    rib_long_name = _dedup_val(id_to_name.get(str(rib_long), str(rib_long)))
    rib_long_space = _dedup_val(components.get("rib_long_space", ""))
    deck_material = _dedup_val(components.get("deck_material", ""))
    rib_trans_material = _dedup_val(components.get("rib_trans_material", ""))
    rib_long_material = _dedup_val(components.get("rib_long_material", ""))
    panel_length = _dedup_val(components.get("panel_length", ""))
    return (deck_type, thickness, rib_trans_name, rib_long_name, rib_long_space,
            deck_material, rib_trans_material, rib_long_material, panel_length)


def write_deck_sheet(wb, ui_data, section_no_map=None):
    """写 Sheet '上部结构-桥面系库' — 查重复用模式

    section_no_map: 截面名称→编号映射，用于将截面名称转换为编号写入
    返回: existing_map {dedup_key: deck_number} 用于更新 bridge_deck_type_ref
    """
    ws = _find_sheet(wb, ["上部结构-桥面系库", "桥面板库", "桥面板", "桥面"])
    if ws is None:
        return {}

    hm = _build_header_map(ws, header_row=1)

    def wh(row, hdr, val):
        col = hm.get(hdr)
        if col is not None:
            ws.cell(row, col, _replace_empty_with_slash(val))

    seq_key = "桥面系编号" if "桥面系编号" in hm else "编号"

    # 读取现有行，构建 dedup map（统一用截面名称比较，不用编号）
    existing_map = {}  # {dedup_key: row_num}
    max_num = 0
    # 先构建 编号→名称 反向映射（用于将 sheet 中的截面编号还原为名称）
    _no_to_name = {}
    if section_no_map:
        _no_to_name = {str(v): k for k, v in section_no_map.items()}
    for r in range(2, ws.max_row + 1):
        row_num = _cell(ws, r, seq_key, hm)
        if row_num is None:
            continue
        try:
            num = int(str(row_num).strip())
        except ValueError:
            continue
        max_num = max(max_num, num)
        # 构建 dedup key（截面字段用名称，空值统一为 "/"，与写侧保持一致）
        deck_type = _dedup_val(_cell(ws, r, "桥面系类型", hm))
        thickness = _dedup_val(_cell(ws, r, "桥面板厚(mm)", hm))
        rib_trans_raw = _dedup_val(_cell(ws, r, "横肋截面", hm))
        rib_long_raw = _dedup_val(_cell(ws, r, "纵肋截面", hm))
        rib_trans = _no_to_name.get(rib_trans_raw, rib_trans_raw)
        rib_long = _no_to_name.get(rib_long_raw, rib_long_raw)
        rib_long_space = _dedup_val(_cell(ws, r, "纵肋横向布置(mm)", hm))
        deck_material = _dedup_val(_cell(ws, r, "桥面板材质", hm))
        rib_trans_material = _dedup_val(_cell(ws, r, "横肋材质", hm))
        rib_long_material = _dedup_val(_cell(ws, r, "纵肋材质", hm))
        panel_length = _dedup_val(_cell(ws, r, "单块面板长度(m)", hm))
        # 混凝土标号已并入桥面板材质，不再单独入 key
        dedup_key = (deck_type, thickness, rib_trans, rib_long, rib_long_space,
                     deck_material, rib_trans_material, rib_long_material, panel_length)
        existing_map[dedup_key] = num

    # 收集所有项目的 components
    projects = ui_data.values() if isinstance(ui_data, dict) and not ui_data.get("components") else [ui_data]
    existing_count = len(existing_map)
    new_count = 0
    reuse_count = 0

    for proj in projects:
        components = None
        for v in proj.get("components", {}).values():
            components = v
            break
        if not components:
            continue

        # 构建 临时截面ID → 截面名称 映射
        sec_dict = proj.get("section", ui_data.get("section", {}))
        id_to_name = {sid: rec.get("name", "") for sid, rec in sec_dict.items() if isinstance(rec, dict)}

        deck_type = _dedup_val(components.get("deck_type", "单层钢面板"))
        thickness = _dedup_val(components.get("deck_thickness", ""))
        rib_trans = components.get("rib_trans_sec", "")
        rib_long = components.get("rib_long_sec", "")
        # 先将临时ID解析为截面名称（dedup 统一用名称比较，不转编号）
        rib_trans_name = _dedup_val(id_to_name.get(str(rib_trans), str(rib_trans)))
        rib_long_name = _dedup_val(id_to_name.get(str(rib_long), str(rib_long)))
        rib_trans_no = str(section_no_map.get(rib_trans_name, rib_trans_name)) if section_no_map else rib_trans_name
        rib_long_no = str(section_no_map.get(rib_long_name, rib_long_name)) if section_no_map else rib_long_name
        rib_long_space = _dedup_val(components.get("rib_long_space", ""))
        deck_material = _dedup_val(components.get("deck_material", ""))
        rib_trans_material = _dedup_val(components.get("rib_trans_material", ""))
        rib_long_material = _dedup_val(components.get("rib_long_material", ""))
        panel_length = _dedup_val(components.get("panel_length", ""))
        # 混凝土标号已并入桥面板材质，不再单独入 key
        dedup_key = (deck_type, thickness, rib_trans_name, rib_long_name, rib_long_space,
                     deck_material, rib_trans_material, rib_long_material, panel_length)

        if dedup_key in existing_map:
            reuse_count += 1
            continue  # 已存在，复用

        # 新增行：找序号列第一个空行
        new_count += 1
        max_num += 1
        existing_map[dedup_key] = max_num
        seq_col = hm.get(seq_key)
        r = 2
        if seq_col:
            for r in range(2, ws.max_row + 2):
                if ws.cell(r, seq_col).value is None:
                    break
        else:
            r = ws.max_row + 1 if ws.max_row >= 2 else 2
        wh(r, seq_key, max_num)
        wh(r, "桥面系类型", deck_type)
        wh(r, "桥面板厚(mm)", thickness)
        wh(r, "横肋截面", rib_trans_no)
        wh(r, "纵肋截面", rib_long_no)
        wh(r, "纵肋横向布置(mm)", components.get("rib_long_space", ""))
        if "单块面板长度(m)" in hm:
            _pl = components.get("panel_length", "")
            # 双层钢面板：未启用分节时写 "/"，启用时写实际值
            if deck_type == "双层钢面板" and not components.get("seg_enabled", False):
                _pl = "/"
            wh(r, "单块面板长度(m)", _pl)
        if "桥面板材质" in hm:
            wh(r, "桥面板材质", components.get("deck_material", ""))
        if "横肋材质" in hm:
            wh(r, "横肋材质", components.get("rib_trans_material", ""))
        if "纵肋材质" in hm:
            wh(r, "纵肋材质", components.get("rib_long_material", ""))

    print(f"[桥面系库] 现有{existing_count}条，新增{new_count}条，复用{reuse_count}条")
    return existing_map


def _read_substruct_rows(ws, hm, HDR_KEY):
    """读取下部结构库现有行 → [(行号, {参数字典}), ...]"""
    rows = []
    for r in range(2, ws.max_row + 1):
        row_num = _cell(ws, r, "下部结构编号", hm)
        if row_num is None:
            continue
        try:
            num = int(str(row_num).strip())
        except ValueError:
            continue
        params = {}
        for hdr, dk in HDR_KEY:
            if dk is not None:
                val = _cell(ws, r, hdr, hm)
                params[dk] = _safe_str(val) if val is not None else "/"
        rows.append((num, params))
    return rows


def _substruct_params(entry, HDR_KEY, section_no_map=None, id_to_name=None):
    """从 entry 构建用于查重的参数字典（仅含 HDR_KEY 中的字段）

    Args:
        entry: 下部结构条目数据
        HDR_KEY: 表头到字段的映射列表
        section_no_map: 截面名称→编号映射，用于将截面名称转换为编号
        id_to_name: 临时截面ID→截面名称映射，用于解析临时ID

    Returns:
        params: 参数字典
    """
    # 截面引用字段
    SEC_REF_KEYS = {"pile_sec", "dist_trans_sec", "dist_long_sec", "brace_sec", "brace_sec_tilt"}

    params = {}
    for _, dk in HDR_KEY:
        if dk is not None:
            v = entry.get(dk, "/")
            v = str(v).strip() if v is not None else "/"
            # 如果是截面引用字段且有映射，先解析临时ID为名称，再转换为编号
            if dk in SEC_REF_KEYS and section_no_map and v and v != "/":
                # v 可能是临时ID（如"3"），先解析为截面名称
                sec_name = v
                if id_to_name and v in id_to_name:
                    sec_name = id_to_name[v]
                # 用名称查正式编号
                if sec_name in section_no_map:
                    v = str(section_no_map[sec_name])
            params[dk] = v
    return params


def write_substructure_sheet(wb, d, section_no_map=None):
    """写 Sheet '下部结构库' — 查重，已有条目复用编号，新增无重叠条目

    Args:
        wb: 工作簿
        d: 项目数据字典
        section_no_map: 截面名称→编号映射，用于将截面名称转换为编号写入

    Returns:
        subs_no_map: dict {substructure_key: 下部结构编号}
    """
    ws = _find_sheet(wb, ["下部结构库", "下部结构", "下部"])
    if ws is None:
        return {}

    hm = _build_header_map(ws, header_row=1)

    def wh(row, hdr, val):
        col = hm.get(hdr)
        if col is not None:
            ws.cell(row, col, _replace_empty_with_slash(val))

    HDR_KEY = [
        ("下部结构编号", None),
        ("结构类型", "type"),
        ("桩类型", "pile_type"),
        ("桩截面", "pile_sec"),
        ("桩横向布置(mm)", "pile_trans_space"),
        ("桩纵向布置(mm)", "pile_long_space"),
        ("伸缩缝(mm)", "expansion_len"),
        ("分配梁(横)截面", "dist_trans_sec"),
        ("分配梁(横)长度(mm)", "dist_trans_len"),
        ("分配梁(横)左端相对线路中心线横向间距(mm)", "dist_trans_offset"),
        ("分配梁(纵)左端相对联中心线横向间距(mm)", "dist_long_center_offset"),
        ("分配梁(纵)截面", "dist_long_sec"),
        ("分配梁(纵)长度(mm)", "dist_long_len"),
        ("联结系形式", "brace_form"),
        ("联结系截面", "brace_sec"),
        ("联结系斜杆/竖杆截面", "brace_sec_tilt"),
        ("联结系间距(mm)", "brace_space"),
        ("联结系高度(mm)", "brace_height"),
        ("桩材质", "pile_material"),
        ("分配梁(横)材质", "dist_trans_material"),
        ("分配梁(纵)材质", "dist_long_material"),
        ("联结系材质", "brace_material"),
    ]

    # 1. 读取现有行
    existing = _read_substruct_rows(ws, hm, HDR_KEY)
    # 建立参数字典→行号的映射（键=冻结的参数字典）
    existing_map = {}  # frozenset(params.items()) → row_num
    max_num = 0
    for num, params in existing:
        key = frozenset(params.items())
        existing_map[key] = num
        if num > max_num:
            max_num = num

    # 2. 遍历当前模型条目，查重
    substructure = d.get("substructure", {})
    # 构建 临时截面ID → 截面名称 映射
    id_to_name = {sid: rec.get("name", "") for sid, rec in d.get("section", {}).items() if isinstance(rec, dict)}
    keys = sorted(substructure.keys(), key=lambda x: int(x) if str(x).isdigit() else 0)
    subs_no_map = {}      # {key: 下部结构编号}
    new_rows = []         # [(assigned_num, params_dict), ...] 待写入的新行

    for k in keys:
        entry = substructure[k]
        entry_params = _substruct_params(entry, HDR_KEY, section_no_map, id_to_name)
        entry_key = frozenset(entry_params.items())

        if entry_key in existing_map:
            # 匹配到已有条目，复用编号
            subs_no_map[k] = existing_map[entry_key]
        else:
            # 新增条目
            max_num += 1
            subs_no_map[k] = max_num
            new_rows.append((max_num, entry_params))
            existing_map[entry_key] = max_num  # 防同项目内重复

    # 3. 写入新增行（不清除原有数据）
    if new_rows:
        seq_col = hm.get(HDR_KEY[0][0])  # 下部结构编号列
        next_row = 2
        if seq_col:
            for r in range(2, ws.max_row + 2):
                if ws.cell(r, seq_col).value is None:
                    next_row = r
                    break
        else:
            next_row = ws.max_row + 1 if ws.max_row >= 2 else 2
        for num, params in new_rows:
            for hdr, dk in HDR_KEY:
                if dk is None:
                    wh(next_row, hdr, num)
                else:
                    wh(next_row, hdr, params.get(dk, "/"))
            next_row += 1

    print(f"[下部结构库] 现有{len(existing)}条，新增{len(new_rows)}条，复用{len(keys)-len(new_rows)}条")
    return subs_no_map


def write_vehicle_sheet(wb, d):
    """写 Sheet '车辆设备库' — 查重复用模式

    d: 可以是单个项目 dict，也可以是 projects_dict（多项目）

    注：标准库（vehicle_parameter.xlsx）仅供 UI 下拉选车时填入参数，
    与保存逻辑无关 —— 项目车辆数据独立管理（可能有意与标准库不同），
    保存时只做"同名同类型 → 复用旧行"，不做参数比对。
    """
    ws = _find_sheet(wb, ["车辆设备库"])
    if ws is None:
        return

    hm = _build_header_map(ws, header_row=2)

    def wh(row, hdr, val):
        col = hm.get(hdr)
        if col is not None:
            ws.cell(row, col, _replace_empty_with_slash(val))

    seq_key = "车辆设备编号"

    # 读取现有行，构建 dedup map: {dedup_key: (行号, 序号)}
    existing_map = {}
    max_num = 0
    for r in range(3, ws.max_row + 1):
        seq_val = _cell(ws, r, seq_key, hm)
        if seq_val is None:
            continue
        try:
            num = int(str(seq_val).strip())
        except ValueError:
            continue
        max_num = max(max_num, num)
        vname = _safe_str(_cell(ws, r, "车辆/设备名称", hm)) or ""
        vtype = _safe_str(_cell(ws, r, "车辆荷载类型", hm)) or ""
        existing_map[(vname, vtype)] = (r, num)

    # 收集所有项目的 vehicles
    all_vehicles = {}
    projects = d.values() if isinstance(d, dict) and not d.get("vehicle") and not d.get("components") else [d]
    for proj in projects:
        for vname, vdata in proj.get("vehicle", {}).items():
            if vname not in all_vehicles:
                all_vehicles[vname] = vdata

    existing_count = len(existing_map)
    new_count = 0
    reuse_count = 0
    conflicts = []  # 保留返回结构（恒空：不再与标准库比对）

    for vname, vdata in all_vehicles.items():
        lt = vdata.get("vehicle_type", "")
        dedup_key = (vname, lt)

        params = vdata.get("vehicle_params", [])

        if dedup_key in existing_map:
            # 同名同类型已存在于 Excel → 复用旧行（项目数据独立管理，不做参数比对）
            reuse_count += 1
            row_idx = existing_map[dedup_key][0]
        else:
            # 新增行
            new_count += 1
            max_num += 1
            seq_col = hm.get(seq_key)
            row_idx = 3
            if seq_col:
                for r in range(3, ws.max_row + 2):
                    if ws.cell(r, seq_col).value is None:
                        row_idx = r
                        break
            else:
                row_idx = ws.max_row + 1 if ws.max_row >= 3 else 3
            existing_map[dedup_key] = (row_idx, max_num)
            wh(row_idx, seq_key, max_num)
            wh(row_idx, "车辆/设备", vdata.get("vehicle_tag", vname))
            wh(row_idx, "车辆荷载类型", lt)
            wh(row_idx, "车辆/设备名称", vdata.get("vehicle_var", vname))

        if params:
            wh(row_idx, "车轮间距/轴距(m)", params[0])

        if lt == "公路标准荷载":
            if len(params) > 1:
                col_std = hm.get("标准荷载类型")
                if col_std:
                    ws.cell(row_idx, col_std, _replace_empty_with_slash(params[1]))
        elif "吊装设备" in lt:
            # params: [轴距, 自重, 吊重, 吊臂, 接地长度, 接地宽度]
            for hdr, pi in [("吊装设备自重(kN)", 1), ("吊重(kN)", 2), ("吊臂(m)", 3),
                            ("吊装设备接地长度(m)", 4), ("吊装设备接地宽度(m)", 5)]:
                col = hm.get(hdr)
                if col and pi < len(params):
                    ws.cell(row_idx, col, _replace_empty_with_slash(params[pi]))
        elif "钻孔设备" in lt:
            # params: [轴距, 自重, 扭矩反力, 接地长度, 接地宽度]
            for hdr, pi in [("钻孔设备自重(kN)", 1), ("扭矩反力(kN)", 2),
                            ("钻孔设备接地长度(m)", 3), ("钻孔设备接地宽度(m)", 4)]:
                col = hm.get(hdr)
                if col and pi < len(params):
                    ws.cell(row_idx, col, _replace_empty_with_slash(params[pi]))
        elif "一般车辆" in lt and len(params) >= 3:
            col_axle = hm.get("轴重(kN)")
            col_dist = hm.get("距前轴(m)")
            if col_axle:
                ws.cell(row_idx, col_axle, _replace_empty_with_slash(params[1]))
            if col_dist:
                ws.cell(row_idx, col_dist, _replace_empty_with_slash(params[2]))

    print(f"[车辆设备库] 现有{existing_count}条，新增{new_count}条，复用{reuse_count}条"
          + (f"，冲突{len(conflicts)}条" if conflicts else ""))
    return conflicts


def write_connection_sheet(wb, d):
    """写入 '连接形式' Sheet — 查重复用模式

    每行以 type 名标识，SDx/SDy/SDz/SRx/SRy/SRz/NSDx 各一列。
    d: 可以是单个项目 dict，也可以是 projects_dict（多项目）
    """
    ws = _find_sheet(wb, ["连接形式"])
    if ws is None:
        return

    hm = _build_header_map(ws, header_row=2)

    def wh(row, hdr, val):
        col = hm.get(hdr)
        if col is not None:
            ws.cell(row, col, _replace_empty_with_slash(val))

    # 读取现有行，构建 dedup map {ctype: row_num}
    existing_map = {}
    for r in range(3, ws.max_row + 1):
        ctype = _safe_str(_cell(ws, r, "连接形式", hm))
        if ctype:
            existing_map[ctype] = r

    # 收集所有项目的 connection entries（按 type 去重）
    all_conns = {}  # {ctype: entry}
    projects = d.values() if isinstance(d, dict) and not d.get("connection") and not d.get("components") else [d]
    for proj in projects:
        for entry in proj.get("connection", {}).values():
            if not isinstance(entry, dict):
                continue
            ctype = entry.get("type", "")
            if not ctype or ctype in ("/", "自定义", ""):
                continue
            if ctype not in all_conns:
                all_conns[ctype] = entry

    existing_count = len(existing_map)
    new_count = 0
    reuse_count = 0

    for ctype, entry in all_conns.items():
        if ctype in existing_map:
            reuse_count += 1
            row_idx = existing_map[ctype]
        else:
            new_count += 1
            conn_col = hm.get("连接形式")
            row_idx = 3
            if conn_col:
                for r in range(3, ws.max_row + 2):
                    if ws.cell(r, conn_col).value is None:
                        row_idx = r
                        break
            else:
                row_idx = ws.max_row + 1 if ws.max_row >= 3 else 3
            existing_map[ctype] = row_idx

        wh(row_idx, "连接形式", ctype)
        # 刚度列：缓存键为短键（SDx/SDy/SDz/SRx/SRy/SRz/NSDx），
        # Excel 表头为带单位长键（SDx(kN/m) 等），需映射读取、按表头写入。
        for cache_key, hdr in [
            ("SDx", "SDx(kN/m)"), ("SDy", "SDy(kN/m)"), ("SDz", "SDz(N/mm)"),
            ("SRx", "SRx(N·mm/rad)"), ("SRy", "SRy(N·mm/rad)"), ("SRz", "SRz(N·mm/rad)"),
            ("NSDx", "NSDx(N/mm)"),
        ]:
            v = entry.get(cache_key, "/")
            wh(row_idx, hdr, v if v and v not in ("",) else "/")

    print(f"[连接形式] 现有{existing_count}条，新增{new_count}条，复用{reuse_count}条")


# ============================================================
# 土层覆盖表（项目级读写）
# ============================================================

def make_soil_label(project_name, project_no):
    """生成项目唯一的 soil_label（命名空间前缀）：{项目名}{编号}

    两张关联 sheet 由前缀派生：{label}地形表（地质插件）、{label}土层表（逐桩土层覆盖表）。
    做 Excel sheet 名合法化（去非法字符、去首尾空格、截断 ≤31 字符）。
    唯一性依据：项目名+编号在项目创建时查重（combined_key），故 label 全局唯一。
    """
    label = f"{project_name}{project_no}"
    for ch in ('\\', '/', '?', '*', '[', ']', ':', "'"):
        label = label.replace(ch, '')
    label = label.strip()
    if len(label) > 31:
        label = label[:31]
    return label


def copy_soil_sheet(excel_path, src_label, dst_label):
    """把源项目地形表（{src_label}地形表）及土层表（{src_label}土层表，若存在）
    复制为 dst_label 对应的同名 sheet。

    openpyxl 无原生 sheet 复制，逐 cell 复制值（格式由 _center_align_all_cells 统一）。
    新 sheet 用 create_sheet(title, index) 插入到源 sheet 之后（紧贴其后）。
    源 sheet 不存在 → 返回 False；其它异常捕获并返回 False（不阻断另存为）。
    """
    if not excel_path or not os.path.exists(excel_path) or not src_label or not dst_label:
        return False
    copied = False
    try:
        wb = openpyxl.load_workbook(excel_path)
        for src_name, dst_name in ((src_label + "地形表", dst_label + "地形表"),
                                   (src_label + "土层表", dst_label + "土层表")):
            if src_name not in wb.sheetnames:
                continue
            src_ws = wb[src_name]
            src_idx = wb.sheetnames.index(src_name)
            if dst_name in wb.sheetnames:
                del wb[dst_name]
            dst_ws = wb.create_sheet(dst_name, index=src_idx + 1)
            for r in range(1, src_ws.max_row + 1):
                for c in range(1, src_ws.max_column + 1):
                    dst_ws.cell(r, c, src_ws.cell(r, c).value)
            copied = True
        _center_align_all_cells(wb)
        wb.save(excel_path)
        if copied:
            print(f"[soil] 复制地形 sheet: {src_label} → {dst_label} (含土层表)")
    except Exception as e:
        print(f"[soil] 复制地形 sheet 失败: {e}")
        return False
    return copied


PILE_SOIL_HEADERS = ["桩号", "土顶标高(m)", "土层名称", "土层厚度(m)", "土层重度(kN/m3)", "黏聚力c(kPa)", "内摩擦角φ(°)", "桩极限侧阻力标准值qsik(kPa)", "桩极限端阻力标准值qpk(kPa)"]


def write_pile_soil_sheet(wb, proj_dict):
    """把项目 per-pile 土层写入 {soil_label}土层表 覆盖表（每桩一行、列内逗号分层）。

    - 插入位置：对应地形表（{soil_label}地形表）之后（create_sheet index）；无地形表追加末尾（不依赖地形表存在）。
    - 写回范围：全项目（由 save_trestle_to_excel 遍历调用）；某桩 soil_data 为空则整项目跳过（避免空表覆盖）。
    - 列格式：桩号 | 土顶标高 | 土层名称 | 层厚(m) | 重度 | c | φ | qsik | qpk；多层用英文逗号分隔。
    """
    label = str(proj_dict.get("soil_label", "") or "").strip()
    if not label:
        return False
    sub = proj_dict.get("substructure", {})
    terrain_derived = None  # 惰性：仅当某桩内存 soil_data 为空时才从地形插值推导（支持纯 Excel 输入）
    rows = []
    for sk in sorted(sub, key=lambda k: int(k) if k.isdigit() else 0):
        sd = sub.get(sk, {}).get("soil_data") or {}
        if not sd:
            if terrain_derived is None:
                terrain_derived = derive_pile_soil_from_terrain(wb, proj_dict)
            sd = (terrain_derived.get(sk) or {}).get("soil", {})
        if not sd:
            continue
        names = [str(n) for n in sd.keys()]
        thicks, weights, cs, phis, qsiks, qpks = [], [], [], [], [], []
        for v in sd.values():
            vv = [str(x) for x in (v or [])]
            thicks.append(vv[0] if len(vv) > 0 else "")
            weights.append(vv[1] if len(vv) > 1 else "")
            cs.append(vv[2] if len(vv) > 2 else "")
            phis.append(vv[3] if len(vv) > 3 else "")
            qsiks.append(vv[4] if len(vv) > 4 else "")
            qpks.append(vv[5] if len(vv) > 5 else "")
        ge = sub.get(sk, {}).get("ground_elevation")
        if ge in (None, "", "/") and terrain_derived:
            ge = (terrain_derived.get(sk) or {}).get("ground_level")
        top = str(ge) if ge not in (None, "", "/") else ""
        rows.append([f"{int(sk)}#", top, ",".join(names), ",".join(thicks),
                     ",".join(weights), ",".join(cs), ",".join(phis),
                     ",".join(qsiks), ",".join(qpks)])
    if not rows:
        return False
    target = label + "土层表"
    if target in wb.sheetnames:
        wb.remove(wb[target])
    terrain_sheet = label + "地形表"
    if terrain_sheet in wb.sheetnames:
        idx = wb.sheetnames.index(terrain_sheet) + 1
    else:
        idx = len(wb.sheetnames)
    ws = wb.create_sheet(target, index=min(idx, len(wb.sheetnames)))
    for ci, h in enumerate(PILE_SOIL_HEADERS, 1):
        ws.cell(1, ci, h)
    for ri, row in enumerate(rows, 2):
        for ci, v in enumerate(row, 1):
            ws.cell(ri, ci, v)
    return True


def derive_pile_soil_from_terrain(wb, proj_dict):
    """无土层表覆盖 / 内存 soil_data 时，从地形表 + 土层参数库插值推导各桩 per-pile 土层。

    支持"仅 Excel 输入地形"场景：即使项目从未在 UI 中访问，保存时也能生成土层表覆盖表。
    返回 {sk: {"soil": {层名: [层厚,重度,c,φ,qsik,qpk]}, "ground_level": float}}；无地形/不可推导返回空 dict。

    纯几何插值复用 General.Geometry.polyline_y_at_x；跨径解析复用 General.DataUtils.midasdisttolst。
    """
    from General.Geometry import polyline_y_at_x
    from General.DataUtils import midasdisttolst

    label = str(proj_dict.get("soil_label", "") or "").strip()
    terrain_sheet = label + "地形表" if label else ""
    if not terrain_sheet or terrain_sheet not in wb.sheetnames:
        return {}
    lib = {}
    if "土层参数库" in wb.sheetnames:
        ws_lib = wb["土层参数库"]
        for r in range(2, ws_lib.max_row + 1):
            lid = str(ws_lib.cell(r, 1).value or "").strip()
            if lid:
                lib[lid] = {
                    "name": str(ws_lib.cell(r, 2).value or "").strip(),
                    "weight": ws_lib.cell(r, 3).value or 0,
                    "cohesion": ws_lib.cell(r, 4).value or 0,
                    "friction_angle": ws_lib.cell(r, 5).value or 0,
                    "qsik": ws_lib.cell(r, 6).value or 0,
                    "qpk": ws_lib.cell(r, 7).value or 0,
                }
    ws = wb[label + "地形表"]
    segments, y_bottom = [], None
    for r in range(2, ws.max_row + 1):
        raw = str(ws.cell(r, 1).value or "").strip()
        if raw == "计算底边线y坐标":
            bv = ws.cell(r, 2).value
            if bv is not None:
                try:
                    y_bottom = float(bv)
                except (ValueError, TypeError):
                    pass
            continue
        if not raw:
            continue
        coords = []
        for c in range(2, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v is None or str(v).strip() in ("", "/"):
                continue
            m = re.match(r"\(\s*([\d.-]+)\s*,\s*([\d.-]+)\s*\)", str(v))
            if m:
                coords.append((float(m.group(1)), float(m.group(2))))
        if len(coords) >= 2:
            segments.append((raw, coords))
    if not segments:
        return {}
    basic = proj_dict.get("basic", {})
    try:
        x0 = float(basic.get("x_origin", 0) or 0)
    except (ValueError, TypeError):
        x0 = 0.0
    pile_xs = [x0]
    for sp in midasdisttolst(basic.get("span", "")):
        pile_xs.append(pile_xs[-1] + sp)
    sub = proj_dict.get("substructure", {})
    result = {}
    for sk in sorted(sub, key=lambda k: int(k) if k.isdigit() else 0):
        try:
            idx = int(sk)
        except (ValueError, TypeError):
            continue
        if idx >= len(pile_xs):
            continue
        px = pile_xs[idx]
        layers = []
        for sid, coords in segments:
            y = polyline_y_at_x(px, coords)
            if y is not None:
                layers.append((sid, y))
        layers.sort(key=lambda t: -t[1])
        if not layers:
            continue
        soil = {}
        for i, (sid, top_y) in enumerate(layers):
            bot_y = layers[i + 1][1] if i + 1 < len(layers) else (y_bottom if y_bottom is not None else top_y)
            th = top_y - bot_y
            if th <= 0:
                continue
            p = lib.get(sid, {})
            soil[p.get("name", f"层{sid}")] = [round(th, 3), p.get("weight", 0),
                                               p.get("cohesion", 0), p.get("friction_angle", 0),
                                               p.get("qsik", 0), p.get("qpk", 0)]
        if soil:
            result[sk] = {"soil": soil, "ground_level": round(layers[0][1], 3)}
    return result


# ============================================================
# 项目级辅助（行删除 / 项目截面库读取）
# ============================================================


def delete_project_row(project_name, project_no, excel_path):
    """立即从 Excel 主表删除指定项目的行（不影响其他 sheet）"""
    if not excel_path or not os.path.exists(excel_path):
        return
    try:
        wb = openpyxl.load_workbook(excel_path)
        ws = _find_sheet(wb, ["栈桥参数", "1"])
        if ws is None:
            wb.close()
            return
        hm = _build_header_map(ws, header_row=2)

        # 找到目标行
        target_row = None
        data_rows = 0
        for r in range(3, ws.max_row + 1):
            pname = _safe_str(_cell(ws, r, "项目名称", hm))
            pno = _safe_str(_cell(ws, r, "编号", hm))
            if pname or pno:
                data_rows += 1
                if pname == project_name and pno == project_no:
                    target_row = r

        if target_row is None:
            wb.close()
            return

        if data_rows <= 1:
            wb.close()
            return  # 至少保留一行

        # 删除行（上移）
        ws.delete_rows(target_row)
        wb.save(excel_path)
        wb.close()
        print(f"[excel_io] 已从 Excel 删除项目行: {project_name}{project_no}")
    except Exception as e:
        print(f"[excel_io] 删除行失败: {e}")


def load_project_section_library(excel_path):
    """从项目参数表 Excel 截面库 sheet 读取截面数据（name → {type, params}）

    与 load_section_library（properties_parameter.xlsx 基础库）互补：
    本函数读取项目参数表中用户新增/修改的截面。
    """
    # 每种截面类型对应的参数键
    _DIM_KEYS = {
        'I':   ['H', 'B', 'tw', 'tf1', 'r1', 'r2'],
        '2I':  ['H', 'B', 'tw', 'tf1', 'C',  'tf2'],
        'HM':  ['H', 'B', 'tw', 'tf1', 'r1'],
        '2HM': ['H', 'B', 'tw', 'tf1', 'C',  'tf2'],
        'HN':  ['H', 'B', 'tw', 'tf1', 'r1'],
        '2HN': ['H', 'B', 'tw', 'tf1', 'C',  'tf2'],
        'HW':  ['H', 'B', 'tw', 'tf1', 'r1'],
        '2HW': ['H', 'B', 'tw', 'tf1', 'C',  'tf2'],
        'C':   ['H', 'B', 'tw', 'tf1', 'tf2'],
        '2C':  ['H', 'B', 'tw', 'tf1', 'C'],
        '∠':   ['H', 'B', 'tw', 'r1', 'r2'],
        'O':   ['D', 'd'],
    }
    result = {}
    if not excel_path or not os.path.exists(excel_path):
        return result
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        ws = _find_sheet(wb, ["截面库", "截面"])
        if ws is None:
            wb.close()
            return result
        hm = _build_header_map(ws, header_row=1)
        for r in range(2, ws.max_row + 1):
            seq = _cell(ws, r, "截面编号", hm)
            if seq is None:
                continue
            name_val = _safe_str(_cell(ws, r, "截面名称", hm))
            type_val = _safe_str(_cell(ws, r, "截面类型", hm))
            if not name_val or not type_val:
                continue
            # 按截面类型过滤参数键
            valid_keys = set(_DIM_KEYS.get(type_val, []))
            p = {}
            for hdr, col in hm.items():
                if hdr in ("截面编号", "截面名称", "截面类型", "材料"):
                    continue
                # 只保留该截面类型对应的参数
                if valid_keys and hdr not in valid_keys:
                    continue
                v = ws.cell(r, col).value
                if v is not None:
                    p[hdr] = v
            result[name_val] = {"type": type_val, "params": p}
        wb.close()
    except Exception as e:
        print(f"[excel_io] 加载项目截面库失败: {e}")
    print(f"[excel_io] 项目截面库: path={repr(excel_path)}, loaded={len(result)} sections")
    return result

