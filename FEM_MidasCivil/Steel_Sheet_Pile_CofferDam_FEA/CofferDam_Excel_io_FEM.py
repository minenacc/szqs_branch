"""
围堰 Excel 参数读写模块 — 表头名方式

数据源：CofferDam_ParamTable.xlsx
- 矩形围堰设计参数 sheet: R2=表头, R3起=每个项目一行, 121列
- 支护桩参数 sheet: R1=表头, R2起=截面数据
- 数据选项 sheet: 每列一个类别
- 连接形式 sheet: R1=大分类, R2=表头, R3起=数据
- 土层-XXX sheet: R1=表头, R2起=土层分层

术语表:
    SKGGZ  : 锁扣钢管桩
    Waler  : 围檩（水平向围护梁）
    Strut  : 内支撑（对撑/斜撑）
    SEC    : 截面尺寸列表，元素含义随截面类型而变
    MCT    : MIDAS Civil 命令文本
    proj / proj_key : 单个项目参数字典 / 项目键（"项目名称+围堰编号"）
    hmap   : 表头名 -> 列号 的映射字典 {str: int}

单位约定:
    标高/厚度/长度列名中已带单位(m)；Excel 存储值一般为 m；
    力学参数 c(kPa)、重度(kN/m3)、连接刚度(N/mm、N·mm/rad) 等按表头单位。
"""

import os, warnings
import openpyxl
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# 全局连接形式列表（从Excel读取，供新建项目使用）
global_connections = []


# ═══════════════════════════════════════════════════════════════
# 新增项目模板数据（不含土层/荷载/工况）
# ═══════════════════════════════════════════════════════════════

PROJECT_TEMPLATE = {
    "初始地面标高(m)": "0.5", "初始水位标高(m)": "0",
    "地质参数": "",
    "加权平均重度+粘聚力+摩擦角": "/",
    "支护桩类型": "钢板桩", "支护桩材质":"Q235","支护桩截面编号": "1",
    "围堰高": "9", "围堰顶标高": "1.0", "围堰底标高": "-8.0",
    "承台长(m)": "9.5", "承台宽(m)": "4.5", "承台高(m)": "2.0",
    "承台底标高(m)": "-3.5", "承台顶标高(m)": "-1.5",
    "围堰内边长(m)": "0", "围堰内边宽(m)": "0",
    "规范": "《建筑基坑支护技术规程》（JGJ 120-2012）", "围堰安全等级": "二级",
    "钢结构规范": "《钢结构设计标准》 （GB 50017-2017）", "混凝土规范": "《混凝土结构设计标准》（GB/T 50010-2010）",
    "垫层混凝土等级": "C20", "垫层厚度(m)": "0.2",
    "封底混凝土等级": "C20", "封底厚度(m)": "0",
    "封底支撑间距(m)": "0,0", "钢与土摩擦力(kPa)": "/", "钢与混凝土黏聚力(kPa)": "0",
    "护筒数量": "0", "护筒直径(m)": "0",
    "支护桩底边界": "001001",
    "牛腿长边布置数量": "2", "牛腿短边布置数量": "2",
    "牛腿与支护桩连接": "", "牛腿与围檩连接": "牛腿连接",
    "围檩与支护桩连接": "仅受压1", "围檩与内支撑连接": "",
    "土弹簧分层厚度(m)": "0.5",
    "考虑土的应力路径": "否",
    "超挖深度(m)": "0", "水面至基坑底距离(m)": "0", "开挖面降水高度(m)": "0",
    "是否生成施工阶段": "否",
    "各工况名称": "", "各工况基坑内土顶标高(m)": "",
    "各工况基坑内水面标高(m)": "", "各工况支撑变动": "",
    "辅助换撑": "/", "辅助加围檩i": "/", "辅助加内支撑i": "/",
    "基本组合分项系数": "1.25", "基本组合组合值系数": "1.0",
    "标准组合分项系数": "1.0",
    "标准组合组合值系数": "1.0",
}

# 主表列头顺序（直接用中文表头名作为 key）
HEADERS_IN_ORDER = [
    "项目名称", "围堰编号",
    "初始地面标高(m)", "初始水位标高(m)",
    "地质参数", "加权平均重度+粘聚力+摩擦角",
    "均布附加荷载",
    "矩形局部附加荷载", "条形局部附加荷载",
    "承台长(m)", "承台宽(m)",
    "承台高(m)", "承台底标高(m)",
    "承台顶标高(m)",
    "钢护筒坐标(m)",
    "承台长边预留边距(m)",
    "承台宽边预留边距(m)",
    "围堰内边长(m)", "围堰内边宽(m)",
    "围堰高", "围堰顶标高", "围堰底标高",
    "支护桩类型", "支护桩材质", "支护桩截面编号",
    "土弹簧分层厚度(m)", "考虑土的应力路径",
    "垫层混凝土等级", "垫层厚度(m)",
    "封底支撑间距(m)", "钢与土摩擦力(kPa)", "钢与混凝土黏聚力(kPa)",
    "护筒数量", "护筒直径(m)",
    "封底混凝土等级", "封底厚度(m)",
    "规范", "钢结构规范", "混凝土规范", "围堰安全等级",
    "支护桩底边界",
    "牛腿长边布置数量", "牛腿短边布置数量",
    "牛腿与支护桩连接", "牛腿与围檩连接",
    "围檩与支护桩连接", "围檩与内支撑连接",
    "超挖深度(m)", "水面至基坑底距离(m)", "开挖面降水高度(m)",
    "是否生成施工阶段",
    "各工况名称", "各工况基坑内土顶标高(m)",
    "各工况基坑内水面标高(m)", "各工况支撑变动",
    "辅助换撑", "辅助加围檩i", "辅助加内支撑i",
    "基本组合分项系数", "基本组合组合值系数",
    "标准组合分项系数", "标准组合组合值系数",
]
# 围檩/支撑 1~5
for _i in range(1, 6):
    HEADERS_IN_ORDER.extend([f"围檩{_i}间距(m)", f"围檩{_i}长边材质", f"围檩{_i}长边截面", f"围檩{_i}宽边材质", f"围檩{_i}宽边截面"])
for _i in range(1, 6):
    HEADERS_IN_ORDER.extend([
        f"内支撑{_i}对撑材质", f"内支撑{_i}对撑截面",
        f"内支撑{_i}斜撑材质", f"内支撑{_i}斜撑截面",
        f"内支撑{_i}对撑长边布置(m)", f"内支撑{_i}对撑短边布置(m)",
        f"内支撑{_i}斜撑长边布置(m)", f"内支撑{_i}斜撑短边布置(m)",
        ])


# ═══════════════════════════════════════════════════════════════
# 工具函数
# ═══════════════════════════════════════════════════════════════

def _s(val, default=""):
    """安全转字符串并去首尾空白；None 返回 default。

    Args:
        val: 任意值，可能为 None。
        default (str): val 为 None 时的返回值。

    Returns:
        str: 去空白后的字符串或 default。
    """
    if val is None: return default
    return str(val).strip()

def _ss(val, default=""):
    """先转字符串去空白，再把中文逗号/顿号统一为英文逗号。

    Args:
        val: 任意值，可能为 None。
        default (str): val 为 None 时的返回值。

    Returns:
        str: 归一化后的字符串。
    """
    return _normalize_commas(_s(val, default))

def _normalize_commas(val):
    """将中文逗号"，"、顿号"、"统一替换为英文逗号","。

    Args:
        val: 任意值，可能为 None。

    Returns:
        str: 替换后的字符串；None 返回 ""。
    """
    if val is None: return ""
    return str(val).replace("，", ",").replace("、", ",")

def _split(val):
    """按英文逗号切分并去除空项，用于解析逗号分隔字符串。

    Args:
        val (str|None): 逗号分隔字符串。

    Returns:
        list[str]: 非空的去空白元素列表。
    """
    if not val: return []
    return [x.strip() for x in val.split(",") if x.strip()]

def _f(val, default=0.0):
    """安全转 float，失败返回 default。

    Args:
        val: 任意值。
        default (float): 转换失败时的返回值。

    Returns:
        float: 转换结果或 default。
    """
    if val is None: return default
    try: return float(val)
    except: return default

def _build_header_map(ws, header_row=1):
    """扫描指定行，建立"表头名 -> 列号"映射。

    Args:
        ws (openpyxl.worksheet.worksheet.Worksheet): 工作表对象。
        header_row (int): 表头所在行号（1 基）。

    Returns:
        dict: {表头名(str): 列号(int)}，空表头不收录。
    """
    hmap = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if v is not None:
            h = str(v).strip()
            if h: hmap[h] = c
    return hmap

def _header_map_in_range(ws, header_row, start_col, end_col):
    """在指定列区间内建立"表头名 -> 列号"映射。

    Args:
        ws: 工作表对象。
        header_row (int): 表头行号。
        start_col (int): 起始列号（含）。
        end_col (int): 结束列号（含）。

    Returns:
        dict: {表头名(str): 列号(int)}。
    """
    hmap = {}
    for c in range(start_col, end_col + 1):
        v = ws.cell(header_row, c).value
        if v is not None:
            h = str(v).strip()
            if h: hmap[h] = c
    return hmap

def _cell(ws, row, header_name, hmap):
    """按表头名取指定行的单元格值。

    Args:
        ws: 工作表对象。
        row (int): 行号。
        header_name (str): 表头名。
        hmap (dict): {表头名: 列号} 映射。

    Returns:
        单元格原始值；表头不存在时返回 None。
    """
    col = hmap.get(header_name)
    if col is None: return None
    return ws.cell(row, col).value

def _find_sheet(wb, names):
    """按候选名查找工作表：先精确匹配，再模糊（包含）匹配。

    Args:
        wb (openpyxl.Workbook): 工作簿对象。
        names (list[str]): 候选 sheet 名，按优先级排列。

    Returns:
        匹配到的工作表对象；均未匹配返回 None。
    """
    for name in names:
        if name in wb.sheetnames: return wb[name]
    for sn in wb.sheetnames:
        for name in names:
            if name in sn: return wb[sn]
    return None


# ═══════════════════════════════════════════════════════════════
# 数据读取
# ═══════════════════════════════════════════════════════════════

def get_excel_path():
    """从注册表读取程序安装目录，拼接并返回参数表 Excel 路径。

    读取 HKCU\\Software\\ShuZhiQiaoShi 下的 Applocation，拼接
    Support/program_param/CofferDam_ParamTable.xlsx，并校验文件存在。

    Returns:
        str|None: 参数表 Excel 的绝对路径；注册表缺失或文件不存在时返回 None。

    Example:
        >>> get_excel_path()
        'C:\\\\...\\\\Support\\\\program_param\\\\CofferDam_ParamTable.xlsx'
    """
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\ShuZhiQiaoShi")
        app, _ = winreg.QueryValueEx(key, "Applocation")
        winreg.CloseKey(key)
        path = os.path.join(app, "Support", "program_param", "CofferDam_ParamTable.xlsx")
        if os.path.exists(path): return path
    except:
        pass
    return None


def load_all_params(excel_path=None):
    """读取 CofferDam_ParamTable.xlsx 的全部项目参数、分组、截面与土层数据。

    依次读取主表（矩形围堰设计参数）、支护桩参数、连接形式、各项目关联的
    土层 sheet，并把 sections/connections/soil_data 等共享数据挂到每个项目上。

    Args:
        excel_path (str|None): Excel 路径；None 时调用 get_excel_path() 获取。

    Returns:
        dict: 结构形式
            {
                "projects_dict": {proj_key: proj_dict, ...},
                "project_groups": {"项目名称": ["编号1", "编号2", ...], ...},
                "last_proj_key": str|None,
            }
            其中 proj_dict 为按 HEADERS_IN_ORDER 组织的字段字典，额外含：
            - "_raw_name"/"_raw_no": 原始项目名/编号；
            - "sections": {截面名称: {"type": str, "params": dict}}；
            - "section_seq_to_name": {截面号(int): 截面名称(str)}；
            - "connections": list[dict]，连接形式及 6 向刚度；
            - "soil_data": {土层名称: [层厚, 重度, c, φ, 计算方法, 渗透系数, 承压水头]}。
            找不到文件时返回 {}（注意：此时结构与正常返回不同）。

    Example:
        >>> data = load_all_params()
        >>> data["last_proj_key"]
        '盐宜DT103#'
    """
    if excel_path is None: excel_path = get_excel_path()
    all_projects = {}
    project_groups = {}    # {"盐宜": ["DT103#", ...], ...} 按行序归类
    proj_keys_in_order = [] # 保持 Excel 行顺序

    if not excel_path or not os.path.exists(excel_path):
        print(f"[CofferDam_Excel] 找不到 {excel_path}")
        return all_projects

    wb = openpyxl.load_workbook(excel_path, data_only=True)

    # ==================== 主表：矩形围堰设计参数 ====================
    ws_main = _find_sheet(wb, ["矩形围堰设计参数"])
    if ws_main is not None:
        hm = _build_header_map(ws_main, header_row=2)

        for r in range(3, ws_main.max_row + 1):
            pname = _s(_cell(ws_main, r, "项目名称", hm))
            if not pname:
                continue
            pno = _s(_cell(ws_main, r, "围堰编号", hm))
            proj_key = f"{pname}{pno}" if pno else pname

            proj = {}
            for hdr in HEADERS_IN_ORDER:
                val = _cell(ws_main, r, hdr, hm)
                # 数值型字段转 float
                if hdr in ("土弹簧分层厚度(m)", "超挖深度(m)",
                           "基本组合分项系数", "基本组合组合值系数",
                           "标准组合分项系数", "标准组合组合值系数"):
                    proj[hdr] = _f(val)
                else:
                    proj[hdr] = _s(val)
            proj["_raw_name"] = pname
            proj["_raw_no"] = pno
            all_projects[proj_key] = proj
            proj_keys_in_order.append(proj_key)
            # 按项目名称归类编号（保持行序）
            if pname not in project_groups:
                project_groups[pname] = []
            if pno:
                project_groups[pname].append(pno)

        print(f"[CofferDam_Excel] 主表共读取 {len(all_projects)} 个项目: {list(all_projects.keys())}")

    # ==================== 支护桩参数 ====================
    ws_sec = _find_sheet(wb, ["支护桩参数", "支护桩库", "支护桩"])
    if ws_sec is not None:
        # 支护桩参数 sheet: R1=类别合并行, R2=表头行
        # 获取 支护桩参数 sheet数据
        sec_hm = _build_header_map(ws_sec, header_row=2)
        sec_col_to_hdr = {c: h for h, c in sec_hm.items()}
        sections = {}
        seq_to_name = {}
        for r in range(3, ws_sec.max_row + 1):
            name = _s(_cell(ws_sec, r, "截面名称", sec_hm))
            seq = _cell(ws_sec, r, "截面号", sec_hm)
            if not name: continue
            if seq is not None:
                try: seq_to_name[int(seq)] = name
                except: pass
            d_val = _cell(ws_sec, r, "D(mm)", sec_hm)
            d_str = str(d_val).strip() if d_val is not None else ""
            if d_str and d_str not in ("", "/", "."):
                sections[name] = {
                    "type": "锁扣钢管桩",
                    "params": {
                        "D": _f(_cell(ws_sec, r, "D(mm)", sec_hm)),
                        "t": _f(_cell(ws_sec, r, "t(mm)", sec_hm)),
                        "锁扣宽度": _f(_cell(ws_sec, r, "锁扣宽度(mm)", sec_hm), 35),
                        "钢板桩数量": _f(_cell(ws_sec, r, "钢板桩数量", sec_hm), 1),
                        "钢板桩宽度": _f(_cell(ws_sec, r, "钢板桩宽度(mm)", sec_hm), 600),
                    }
                }
            else:
                params = {}
                for c in range(4, ws_sec.max_column + 1):
                    v = ws_sec.cell(r, c).value
                    if v is not None and str(v).strip() not in ("", "/"):
                        hdr = sec_col_to_hdr.get(c, f"col{c}")
                        if hdr not in ("截面号", "截面名称"):
                            params[hdr] = _f(v)
                sections[name] = {"type": "钢板桩", "params": params}

        for proj in all_projects.values():
            proj["sections"] = sections
            proj["section_seq_to_name"] = seq_to_name

    # ==================== 连接形式 ====================
    ws_conn = _find_sheet(wb, ["连接形式", "连接方式"])
    if ws_conn is not None:
        conn_hm = _build_header_map(ws_conn, header_row=2)
        connections = []
        for r in range(3, ws_conn.max_row + 1):
            name = _s(ws_conn.cell(r, 1).value)
            if not name: continue
            connections.append({
                "name": name,
                "SDx": _s(_cell(ws_conn, r, "SDx(N/mm)", conn_hm)),
                "SDy": _s(_cell(ws_conn, r, "SDy(N/mm)", conn_hm)),
                "SDz": _s(_cell(ws_conn, r, "SDz(N/mm)", conn_hm)),
                "SRx": _s(_cell(ws_conn, r, "SRx(N·mm/rad)", conn_hm)),
                "SRy": _s(_cell(ws_conn, r, "SRy(N·mm/rad)", conn_hm)),
                "SRz": _s(_cell(ws_conn, r, "SRz(N·mm/rad)", conn_hm)),
                "NSDx": _s(_cell(ws_conn, r, "NSDx(N/mm)", conn_hm)),
            })
        for proj in all_projects.values():
            proj["connections"] = connections
        # 存储全局连接列表，供新建项目使用
        global_connections.clear()
        global_connections.extend(connections)

    # ==================== 土层数据 ====================
    for proj_key, proj in all_projects.items():
        soil_ref_raw = proj.get("地质参数", "")
        soil_ref = ""
        # 新格式: "1,盐宜DT41#" — 编码逗号sheet名；旧格式: sheet名
        if soil_ref_raw:
            parts = soil_ref_raw.split(",")
            for p in parts:
                p = p.strip()
                if p in wb.sheetnames:
                    soil_ref = p
                    break
            if not soil_ref:
                soil_ref = parts[0].strip() if parts and parts[0].strip() in wb.sheetnames else ""
        soil_data = {}
        if soil_ref and soil_ref in wb.sheetnames:
            ws_soil = wb[soil_ref]
            soil_hm = _build_header_map(ws_soil, header_row=1)
            SOIL_HDRS = ["土层名称", "层厚(m)", "重度(kN/m3)",
                         "黏聚力c(kPa)", "内摩擦角φ(°)",
                         "计算方法", "渗透系数(m/d)", "土层顶承压水头(m)"]
            for r in range(2, ws_soil.max_row + 1):
                name = _s(_cell(ws_soil, r, "土层名称", soil_hm))
                if not name: continue
                vals = []
                for hdr in SOIL_HDRS[1:]:
                    v = _cell(ws_soil, r, hdr, soil_hm)
                    try: vals.append(float(v) if v is not None else 0.0)
                    except: vals.append(_s(v))
                soil_data[name] = vals
        proj["soil_data"] = soil_data
    wb.close()
    last_proj_key = proj_keys_in_order[-1] if proj_keys_in_order else None
    return {
        "projects_dict": all_projects,
        "project_groups": project_groups,
        "last_proj_key": last_proj_key,
    }


# ═══════════════════════════════════════════════════════════════
# 数据写入
# ═══════════════════════════════════════════════════════════════

def _clear_from_row(ws, start_row):
    """清空从 start_row 起的所有单元格（保留表头）。

    Args:
        ws: 工作表对象。
        start_row (int): 起始行号（含）。

    Returns:
        None
    """
    for row in ws.iter_rows(min_row=start_row, max_row=ws.max_row):
        for cell in row:
            cell.value = None


def _clear_columns_from_row(ws, start_row, columns):
    """仅清除指定列号集合中的单元格，保留其他列的数据。

    Args:
        ws: 工作表对象。
        start_row (int): 起始行号（含）。
        columns (set[int]|list[int]): 需要清空的列号集合。

    Returns:
        None
    """
    for row in ws.iter_rows(min_row=start_row, max_row=ws.max_row):
        for cell in row:
            if cell.column in columns:
                cell.value = None


def _write_cofferdam_soil_sheet(wb, sheet_name, soil_data):
    """写入围堰土层 sheet：不存在则新建，存在则覆写（保留表头）。

    Args:
        wb (openpyxl.Workbook): 工作簿对象。
        sheet_name (str): 目标 sheet 名。
        soil_data (dict): 结构形式
            {土层名称: [层厚, 重度, c, φ, 计算方法, 渗透系数, 承压水头]}。

    Returns:
        None
    """
    SOIL_HDRS = ["土层名称", "层厚(m)", "重度(kN/m3)", "黏聚力c(kPa)",
                 "内摩擦角φ(°)", "计算方法", "渗透系数(m/d)", "土层顶承压水头(m)"]
    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(sheet_name)
        for ci, h in enumerate(SOIL_HDRS, 1):
            ws.cell(1, ci, h)
    else:
        ws = wb[sheet_name]
        _clear_from_row(ws, 2)  # 覆写：保留表头
    for ri, (name, vals) in enumerate(soil_data.items(), 2):
        ws.cell(ri, 1, name)
        for ci, v in enumerate(vals, 2):
            ws.cell(ri, ci, str(v) if v is not None else "")


def _write_sections_to_sheet(wb, all_projects):
    """将内存中新增/修改的截面写入 '支护桩参数' sheet，并更新项目中的截面编号。

    所有项目共享同一份截面库：先收集全局 sections，找出 sheet 中尚不存在的
    新截面并追加（按类型补全表头、写参数），最后把项目中"截面名称"形式的
    引用转换为截面号。

    Args:
        wb (openpyxl.Workbook): 工作簿对象。
        all_projects (dict): 结构形式 {proj_key: proj_dict}，
            每个 proj_dict 需含 "sections" 与 "section_seq_to_name"。

    Returns:
        None: 就地修改 wb（新增截面行）与 all_projects（更新截面编号）。
    """
    # 1. 收集全局 sections（所有项目共享同一份）
    global_sections = {}
    global_seq_to_name = {}
    for proj in all_projects.values():
        if proj.get("sections"):
            global_sections = proj["sections"]
        if proj.get("section_seq_to_name"):
            global_seq_to_name = proj["section_seq_to_name"]
        if global_sections:
            break
    if not global_sections:
        return

    # 2. 找到或创建支护桩参数 sheet
    ws_sec = _find_sheet(wb, ["支护桩参数", "支护桩库", "支护桩"])
    if ws_sec is None:
        ws_sec = wb.create_sheet("支护桩参数")
        ws_sec.cell(1, 1, "支护桩截面")
        ws_sec.cell(2, 1, "截面号")
        ws_sec.cell(2, 2, "截面名称")
        ws_sec.cell(2, 3, "材料标号")

    sec_hm = _build_header_map(ws_sec, header_row=2)

    # 3. 找出已有截面名称（sheet 中已存在的）
    existing_names = set()
    for r in range(3, ws_sec.max_row + 1):
        name = ws_sec.cell(r, sec_hm.get("截面名称", 2)).value
        if name:
            existing_names.add(str(name).strip())

    # 4. 找出需要新增的截面（在 sections 中但不在 sheet 中）
    new_sections = []
    for name, sec_data in global_sections.items():
        if name not in existing_names and name not in ("自定义",):
            new_sections.append((name, sec_data))

    if not new_sections:
        # 没有新截面需要写入，但仍需将 proj 中的截面名转为编号
        for proj in all_projects.values():
            proj["section_seq_to_name"] = global_seq_to_name
            sec_ref = proj.get("支护桩截面编号", "")
            if sec_ref and not str(sec_ref).isdigit():
                for seq, n in global_seq_to_name.items():
                    if n == sec_ref:
                        proj["支护桩截面编号"] = str(seq)
                        break
        return

    # 5. 确定当前最大截面号
    max_seq = 0
    for r in range(3, ws_sec.max_row + 1):
        seq_val = ws_sec.cell(r, sec_hm.get("截面号", 1)).value
        if seq_val is not None:
            try:
                max_seq = max(max_seq, int(seq_val))
            except (ValueError, TypeError):
                pass

    # 6. 确保关键表头存在（按用户指定的列名），不存在则追加
    _header_col_cache = dict(sec_hm)  # 本地副本，可追加

    def _ensure_header(header_name):
        """确保 sheet 中存在指定表头，不存在则在末尾追加，返回其列号。

        Args:
            header_name (str): 表头名。

        Returns:
            int: 表头所在列号。
        """
        if header_name in _header_col_cache:
            return _header_col_cache[header_name]
        new_col = ws_sec.max_column + 1
        ws_sec.cell(2, new_col, header_name)
        _header_col_cache[header_name] = new_col
        return new_col

    # 材料标号列
    mat_col = _ensure_header("材料标号")

    # 锁扣钢管桩列
    skggz_cols = {
        "D(mm)": _ensure_header("D(mm)"),
        "t(mm)": _ensure_header("t(mm)"),
        "锁扣宽度(mm)": _ensure_header("锁扣宽度(mm)"),
        "钢板桩数量": _ensure_header("钢板桩数量"),
    }
    skggz_cols["钢板桩宽度(mm)"] = _ensure_header("钢板桩宽度(mm)")

    # 钢板桩列：使用 Excel 中已有的表头名
    gbz_col_map = {}
    for hdr in ["H(mm)", "B(mm)", "As(mm^2)", "Asy(mm^2)", "Asz(mm^2)",
                "Ixx(mm^4)", "Iyy(mm^4)", "Izz(mm^4)",
                "Cyp(mm)", "Cym(mm)", "Czp(mm)", "Czm(mm)",
                "Qyb(mm^2)", "Qzb(mm^2)", "Peri:O(mm)", "Peri:I(mm)",
                "Cent:y(mm)", "Cent:z(mm)",
                "y1(mm)", "z1(mm)", "y2(mm)", "z2(mm)",
                "y3(mm)", "z3(mm)", "y4(mm)", "z4(mm)",
                "Zyy(mm^3)", "Zzz(mm^3)"]:
        gbz_col_map[hdr] = _ensure_header(hdr)

    # 7. 写入新截面行
    for name, sec_data in new_sections:
        max_seq += 1
        next_row = ws_sec.max_row + 1
        sec_type = sec_data.get("type", "")
        params = sec_data.get("params", {})

        # 截面号、截面名称
        ws_sec.cell(next_row, _header_col_cache.get("截面号", 1), max_seq)
        ws_sec.cell(next_row, _header_col_cache.get("截面名称", 2), name)

        if sec_type == "锁扣钢管桩":
            # 材料标号
            ws_sec.cell(next_row, mat_col, "Q235B")
            # 参数列
            ws_sec.cell(next_row, skggz_cols["D(mm)"], params.get("D", ""))
            ws_sec.cell(next_row, skggz_cols["t(mm)"], params.get("t", ""))
            ws_sec.cell(next_row, skggz_cols["锁扣宽度(mm)"], params.get("锁扣宽度", 35))
            ws_sec.cell(next_row, skggz_cols["钢板桩数量"], int(params.get("钢板桩数量", 1)))
            ws_sec.cell(next_row, skggz_cols["钢板桩宽度(mm)"], params.get("钢板桩宽度", 600))
            # 钢板桩列填 /
            for hdr, col in gbz_col_map.items():
                ws_sec.cell(next_row, col, "/")
        elif sec_type == "钢板桩":
            # 材料标号
            ws_sec.cell(next_row, mat_col, "Q295B")
            # 锁扣钢管桩列填 /
            for col in skggz_cols.values():
                ws_sec.cell(next_row, col, "/")
            # 钢板桩参数列
            for hdr, col in gbz_col_map.items():
                val = params.get(hdr, "/")
                ws_sec.cell(next_row, col, val if val not in (None, "", 0, 0.0) else "/")

        # 更新内存中的 seq_to_name
        global_seq_to_name[max_seq] = name

    # 8. 更新所有项目中的 section_seq_to_name，并将截面名转为编号
    for proj in all_projects.values():
        proj["section_seq_to_name"] = global_seq_to_name
        sec_ref = proj.get("支护桩截面编号", "")
        if sec_ref and not str(sec_ref).isdigit():
            for seq, n in global_seq_to_name.items():
                if n == sec_ref:
                    proj["支护桩截面编号"] = str(seq)
                    break


def _to_float_0(val):
    """将值安全转为 float，空串/None/非法值均返回 0.0。

    Args:
        val: 任意值。

    Returns:
        float: 转换结果或 0.0。
    """
    try:
        return float(str(val).strip() or "0")
    except Exception:
        return 0.0


def _normalize_row(proj):
    """保存前对单个项目行做格式归一化（就地修改 proj）。

    规则：
    1) 封底协助字段恒为 "/"；
    2) 垫层/封底互斥：厚度>0 视为激活（封底优先），未激活方的厚度与等级写 "/"；
    3) 施工阶段：生成=是 -> 超挖/水面至基坑底/开挖面降水 = "0.0"；
       =否 -> 辅助换撑等施工阶段七列 = "/"。

    Args:
        proj (dict): 单个项目参数字典（含中文表头键）。

    Returns:
        None: 就地修改 proj。
    """
    for k in ["封底支撑间距(m)", "钢与土摩擦力(kPa)", "钢与混凝土黏聚力(kPa)",
              "护筒数量", "护筒直径(m)"]:
        proj[k] = "/"
    seal_t = _to_float_0(proj.get("封底厚度(m)", 0))
    cushion_t = _to_float_0(proj.get("垫层厚度(m)", 0))
    if seal_t > 0:
        proj["垫层厚度(m)"] = "/"
        proj["垫层混凝土等级"] = "/"
    elif cushion_t > 0:
        proj["封底厚度(m)"] = "/"
        proj["封底混凝土等级"] = "/"
    else:
        proj["封底厚度(m)"] = "/"
        proj["垫层厚度(m)"] = "/"
        proj["封底混凝土等级"] = "/"
        proj["垫层混凝土等级"] = "/"
    if str(proj.get("是否生成施工阶段", "")).strip() == "是":
        proj["超挖深度(m)"] = "0.0"
        proj["水面至基坑底距离(m)"] = "0.0"
        proj["开挖面降水高度(m)"] = "0.0"
    else:
        for k in ["辅助换撑", "辅助加围檩i", "辅助加内支撑i", "各工况名称",
                  "各工况基坑内土顶标高(m)", "各工况基坑内水面标高(m)", "各工况支撑变动"]:
            proj[k] = "/"


def save_to_excel(excel_path, all_projects, project_keys_in_order=None,
                  soil_owner=None, soil_data_cache=None):
    """将全部项目数据写回 CofferDam_ParamTable.xlsx。

    流程：新增自定义截面 -> 仅清除会被重写的列（R3 起）-> 逐项目写入主表
    字段并自动计算"承台顶标高/围堰底标高"-> 写入各项目自有的土层 sheet ->
    保存工作簿。

    Args:
        excel_path (str): 目标 Excel 路径。
        all_projects (dict): 结构形式 {proj_key: proj_dict}，与
            load_all_params 的 "projects_dict" 一致。
        project_keys_in_order (list[str]|None): 写入顺序；None 时用 dict 顺序。
        soil_owner (dict|None): {proj_key: bool}，True 表示该项目拥有独立土层 sheet。
        soil_data_cache (dict|None): {proj_key: {土层名称: [...]}} 土层数据缓存。

    Returns:
        bool: 保存成功返回 True；文件不存在/找不到主 sheet/保存异常返回 False。

    Example:
        >>> save_to_excel(path, projects, ["盐宜DT103#"], soil_owner, cache)
        True
    """
    # print(list(all_projects.keys()))
    # for k, v in all_projects.items():
    #     print(k)
    #     print(v)
    # print(project_keys_in_order)
    # print(soil_data_cache)
    # print(soil_owner)
    
    if not os.path.exists(excel_path):
        print(f"[CofferDam_Excel] 错误: 找不到 {excel_path}")
        return False

    wb = openpyxl.load_workbook(excel_path)
    ws_main = _find_sheet(wb, ["矩形围堰设计参数", "钢板桩围堰设计参数", "基坑围堰设计参数"])
    if ws_main is None:
        print("[CofferDam_Excel] 找不到主 sheet")
        return False

    hm = _build_header_map(ws_main, header_row=2)

    def wh(row, hdr, val):
        """按表头名把值写入指定行：None/空串统一写为 "/"。

        Args:
            row (int): 目标行号。
            hdr (str): 表头名。
            val: 待写入的值。

        Returns:
            None
        """
        col = hm.get(hdr)
        if col is not None:
            if val is None:
                ws_main.cell(row, col, "/")
            else:
                s = str(val).strip()
                ws_main.cell(row, col, "/" if s == "" else s)

    # 收集所有会被写入的列号，仅清除这些列，保留 Excel 中的其他列数据
    write_cols = set()
    for hdr in HEADERS_IN_ORDER:
        col = hm.get(hdr)
        if col is not None:
            write_cols.add(col)
    for i in range(1, 6):
        for suffix in [f"围檩{i}间距(m)", f"围檩{i}长边材质", f"围檩{i}长边截面", f"围檩{i}宽边材质", f"围檩{i}宽边截面",
                       f"内支撑{i}对撑材质", f"内支撑{i}对撑截面", f"内支撑{i}斜撑材质", f"内支撑{i}斜撑截面",
                       f"内支撑{i}对撑长边布置(m)", f"内支撑{i}对撑短边布置(m)",
                       f"内支撑{i}斜撑长边布置(m)", f"内支撑{i}斜撑短边布置(m)"]:
            col = hm.get(suffix)
            if col is not None:
                write_cols.add(col)
    for hdr in ["承台顶标高(m)", "围堰底标高"]:
        col = hm.get(hdr)
        if col is not None:
            write_cols.add(col)
    # 先写入支护桩参数sheet（新增自定义截面），并将 proj 中的截面名转为编号
    _write_sections_to_sheet(wb, all_projects)

    # 仅清除会被重写的列（R3起）
    _clear_columns_from_row(ws_main, 3, write_cols)

    keys = project_keys_in_order or list(all_projects.keys())
    # print(keys)
    for idx, proj_key in enumerate(keys):
        r = 3 + idx
        proj = all_projects.get(proj_key, {})
        _normalize_row(proj)
        # 直接写入全部表头（缺失的 key 也送入 wh，统一归一化为 "/"）
        for hdr in HEADERS_IN_ORDER:
            wh(r, hdr, proj.get(hdr))
        # 承台顶标高：自动计算 = 承台底标高 + 承台高
        try:
            cap_bottom = float(str(proj.get("承台底标高(m)", 0)).strip() or "0")
            cap_height = float(str(proj.get("承台高(m)", 0)).strip() or "0")
            wh(r, "承台顶标高(m)", str(round(cap_bottom + cap_height, 3)))
        except (ValueError, TypeError):
            wh(r, "承台顶标高(m)", "/")
        # 围堰底标高 = 围堰顶 - 围堰高（自动计算）
        try:
            waler_top = float(str(proj.get("围堰顶标高", 0)).strip() or "0")
            waler_height = float(str(proj.get("围堰高", 0)).strip() or "0")
            wh(r, "围堰底标高", str(round(waler_top - waler_height, 3)))
        except (ValueError, TypeError):
            wh(r, "围堰底标高", "/")

        # 围檩 1~5（从 proj 中文key读取，与 combine_cofferdam_dict 一致）
        for i in range(1, 6):
            wh(r, f"围檩{i}间距(m)", proj.get(f"围檩{i}间距(m)"))
            wh(r, f"围檩{i}长边材质", proj.get(f"围檩{i}长边材质"))
            wh(r, f"围檩{i}长边截面", proj.get(f"围檩{i}长边截面"))
            wh(r, f"围檩{i}宽边材质", proj.get(f"围檩{i}宽边材质"))
            wh(r, f"围檩{i}宽边截面", proj.get(f"围檩{i}宽边截面"))

        # 支撑 1~5
        for i in range(1, 6):
            wh(r, f"内支撑{i}对撑材质", proj.get(f"内支撑{i}对撑材质"))
            wh(r, f"内支撑{i}对撑截面", proj.get(f"内支撑{i}对撑截面"))
            wh(r, f"内支撑{i}斜撑材质", proj.get(f"内支撑{i}斜撑材质"))
            wh(r, f"内支撑{i}斜撑截面", proj.get(f"内支撑{i}斜撑截面"))
            wh(r, f"内支撑{i}对撑长边布置(m)", proj.get(f"内支撑{i}对撑长边布置(m)"))
            wh(r, f"内支撑{i}对撑短边布置(m)", proj.get(f"内支撑{i}对撑短边布置(m)"))
            wh(r, f"内支撑{i}斜撑长边布置(m)", proj.get(f"内支撑{i}斜撑长边布置(m)"))
            wh(r, f"内支撑{i}斜撑短边布置(m)", proj.get(f"内支撑{i}斜撑短边布置(m)"))

    # 写入土层sheet（自有sheet或修改后的参照sheet）
    if soil_owner and soil_data_cache:
        for proj_key, is_owner in soil_owner.items():
            if not is_owner:
                continue  # 仅自有sheet需要写入
            soil_data = soil_data_cache.get(proj_key, {})
            if not soil_data:
                continue
            # sheet名 = 项目key（如"盐宜DT103#"）
            _write_cofferdam_soil_sheet(wb, proj_key, soil_data)

    try:
        wb.save(excel_path)
        print(f"[CofferDam_Excel] 已保存: {excel_path}")
        return True
    except Exception as e:
        print(f"[CofferDam_Excel] 保存失败: {e}")
        return False
    finally:
        wb.close()


def load_soil_template(excel_path):
    """从 Excel sheet '0' 读取模板土层数据，用于新建项目或无有效土层引用时的回退。

    Args:
        excel_path (str): Excel 路径。

    Returns:
        dict: 结构形式
            {土层名称: [层厚(m), 重度, c, φ, 计算方法, 渗透系数, 承压水头]}，
            与 load_all_params 的 soil_data 格式一致；无模板时返回 {}。
    """
    if not excel_path or not os.path.exists(excel_path):
        return {}
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    if "0" not in wb.sheetnames:
        wb.close()
        return {}
    ws = wb["0"]
    hm = _build_header_map(ws, header_row=1)
    SOIL_HDRS = ["层厚(m)", "重度(kN/m3)",
                 "黏聚力c(kPa)", "内摩擦角φ(°)",
                 "计算方法", "渗透系数(m/d)", "土层顶承压水头(m)"]
    data = {}
    for r in range(2, ws.max_row + 1):
        name = _s(ws.cell(r, hm.get("土层名称", 0)).value) if hm.get("土层名称") else ""
        if not name:
            continue
        vals = []
        for hdr in SOIL_HDRS:
            col = hm.get(hdr)
            v = ws.cell(r, col).value if col else None
            if v is not None:
                s = str(v).strip()
                try:
                    vals.append(float(s))
                except:
                    vals.append(s)
            else:
                vals.append("")
        data[name] = vals
    wb.close()
    return data


# ═══════════════════════════════════════════════════════════════
# 施工阶段数据解析/序列化
# ═══════════════════════════════════════════════════════════════

def parse_stage_data(stage_name_str, soil_top_str, water_level_str, support_count_str):
    """将 Excel 中 4 列逗号分隔字符串解析为施工阶段字典。

    以最长列为准对齐，缺失项分别回退为 "" 或 "/"。

    Args:
        stage_name_str (str): 工况类型，如 "取土,加撑,取土"。
        soil_top_str (str): 基坑内土顶标高（m），如 "-1,/,-4"。
        water_level_str (str): 基坑内水面标高（m），如 "-1,-1,-4"。
        support_count_str (str): 支撑层数，如 "/,1,/,2"。

    Returns:
        dict: 结构形式
            {序号(int): {"工况类型": str, "基坑内土顶标高": str,
                        "基坑内水面标高": str, "支撑层数": str}, ...}。

    Example:
        >>> parse_stage_data("取土,加撑", "-1,", "-1,-1", "/,1")
        {1: {'工况类型': '取土', '基坑内土顶标高': '-1',
             '基坑内水面标高': '-1', '支撑层数': '/'},
         2: {'工况类型': '加撑', '基坑内土顶标高': '/',
             '基坑内水面标高': '-1', '支撑层数': '1'}}
    """
    names = _split(stage_name_str)
    soil_tops = _split(soil_top_str)
    water_levels = _split(water_level_str)
    support_counts = _split(support_count_str)
    max_len = max(len(names), len(soil_tops), len(water_levels), len(support_counts))
    result = {}
    for i in range(max_len):
        result[i + 1] = {
            "工况类型": names[i] if i < len(names) else "",
            "基坑内土顶标高": soil_tops[i] if i < len(soil_tops) else "/",
            "基坑内水面标高": water_levels[i] if i < len(water_levels) else "/",
            "支撑层数": support_counts[i] if i < len(support_counts) else "/",
        }
    return result


def serialize_stage_data(stage_dict):
    """将施工阶段字典序列化为 Excel 的 4 列逗号分隔字符串。

    与 parse_stage_data 互为逆操作，按序号升序输出。

    Args:
        stage_dict (dict): 结构形式
            {序号(int): {"工况类型": str, "基坑内土顶标高": str,
                        "基坑内水面标高": str, "支撑层数": str}, ...}。

    Returns:
        tuple[str, str, str, str]: 依次为
            (工况类型串, 土顶标高串, 水面标高串, 支撑层数串)；
            stage_dict 为空时返回 ("", "", "", "")。

    Example:
        >>> serialize_stage_data({1: {"工况类型": "取土", "基坑内土顶标高": "-1",
        ...                           "基坑内水面标高": "-1", "支撑层数": "/"}})
        ('取土', '-1', '-1', '/')
    """
    if not stage_dict:
        return "", "", "", ""
    names = []
    soil_tops = []
    water_levels = []
    support_counts = []
    for i in sorted(stage_dict.keys()):
        d = stage_dict[i]
        names.append(d.get("工况类型", ""))
        soil_tops.append(d.get("基坑内土顶标高", "/"))
        water_levels.append(d.get("基坑内水面标高", "/"))
        support_counts.append(d.get("支撑层数", "/"))
    return ",".join(names), ",".join(soil_tops), ",".join(water_levels), ",".join(support_counts)


# ═══════════════════════════════════════════════════════════════
# 土层 sheet 查询
# ═══════════════════════════════════════════════════════════════

_NON_SOIL_SHEETS = {"矩形围堰设计参数", "支护桩参数", "数据选项", "连接形式", "0", "1"}


def section_mct_trans(section_type, section_name, section_info):
    """按截面类型生成对应的 MIDAS Civil DBUSER 截面命令文本。

    Args:
        section_type (str): 截面类型，取值
            'HW'/'HM'/'HN'/'I'（单 H/工字钢）、'C'（单槽钢）、
            'O'（钢管）、'2HW'/'2HM'/'2HN'/'2I'/'2C'（双拼）。
        section_name (str): 截面名称；钢管类型会自动追加"钢管桩"后缀。
        section_info (list[float]): 截面尺寸列表，长度随类型而变
            （H 型钢一般 5 项 [H,B,tw,tf1,tf2]，钢管 2 项 [D,d]，双拼 6 项）。

    Returns:
        str: 拼装好的 MCT 命令字符串；未知类型返回 ""。
    """
    if section_type == 'HW':
        mct_str = 'DBUSER, {}, CC, 0, 0, 0, 0, 0, 0, YES, NO, H , 2, {}, {}, {}, {}, 0, 0, {}, 0, 0, 0'.format(section_name, section_info[0], section_info[1], section_info[2], section_info[3], section_info[4])
    elif section_type == 'HM':
        mct_str = 'DBUSER, {}, CC, 0, 0, 0, 0, 0, 0, YES, NO, H , 2, {}, {}, {}, {}, 0, 0, {}, 0, 0, 0'.format(section_name, section_info[0], section_info[1], section_info[2], section_info[3], section_info[4])
    elif section_type == 'HN':
        mct_str = 'DBUSER, {}, CC, 0, 0, 0, 0, 0, 0, YES, NO, H , 2, {}, {}, {}, {}, 0, 0, {}, 0, 0, 0'.format(section_name, section_info[0], section_info[1], section_info[2], section_info[3], section_info[4])
    elif section_type == 'I':
        mct_str = 'DBUSER, {}, CC, 0, 0, 0, 0, 0, 0, YES, NO, H , 2, {}, {}, {}, {}, 0, 0, {}, 0, 0, 0'.format(section_name, section_info[0], section_info[1], section_info[2], section_info[3], section_info[4])
    elif section_type == 'C':
        mct_str = 'DBUSER, {}, CC, 0, 0, 0, 0, 0, 0, YES, NO, C , 2, {}, {}, {}, {}, 0, {}, 0, 0, 0, 0'.format(section_name, section_info[0], section_info[1], section_info[2], section_info[3], section_info[3])
    elif section_type == 'O':
        mct_str = 'DBUSER, {}, CC, 0, 0, 0, 0, 0, 0, YES, NO, P , 2, {}, {}, 0, 0, 0, 0, 0, 0, 0, 0'   .format(section_name + '钢管桩', section_info[0], section_info[1])
    elif section_type == '2HW':
        mct_str = 'DBUSER, {}, CC, 0, 0, 0, 0, 0, 0, YES, NO, B , 2, {}, {}, {}, {}, {}, {}, 0, 0, 0, 0, NO, 0, NO, 0, 0'.format(section_name, section_info[0], section_info[1], section_info[2], section_info[3], section_info[4], section_info[5])
    elif section_type == '2HM':
        mct_str = 'DBUSER, {}, CC, 0, 0, 0, 0, 0, 0, YES, NO, B , 2, {}, {}, {}, {}, {}, {}, 0, 0, 0, 0, NO, 0, NO, 0, 0'.format(section_name, section_info[0], section_info[1], section_info[2], section_info[3], section_info[4], section_info[5])
    elif section_type == '2HN':
        mct_str = 'DBUSER, {}, CC, 0, 0, 0, 0, 0, 0, YES, NO, B , 2, {}, {}, {}, {}, {}, {}, 0, 0, 0, 0, NO, 0, NO, 0, 0'.format(section_name, section_info[0], section_info[1], section_info[2], section_info[3], section_info[4], section_info[5])
    elif section_type == '2I':
        mct_str = 'DBUSER, {}, CC, 0, 0, 0, 0, 0, 0, YES, NO, B , 2, {}, {}, {}, {}, {}, {}, 0, 0, 0, 0, NO, 0, NO, 0, 0'.format(section_name, section_info[0], section_info[1], section_info[2], section_info[3], section_info[4], section_info[5])
    elif section_type == '2C':
        mct_str = 'DBUSER, {}, CC, 0, 0, 0, 0, 0, 0, YES, NO, 2C, 2, {}, {}, {}, {}, {}, 0, 0, 0, 0, 0'.format(section_name, section_info[0], section_info[1], section_info[2], section_info[3], section_info[4])
    else:
        mct_str = ''
    return mct_str


def load_section_library(section_xlsx_path):
    """读取 properties_parameter.xlsx 截面库。

    每个 Sheet 代表一类截面，列配置由 SHEET_COL_CONFIG 决定：
      I/2I/HM/2HM/HN/2HN/HW/2HW — 工字钢/H 型钢（C10=H, C11=B, C12=tw, C13=tf1）
      O — 钢管（C7=D, C8=d）
      C/2C — 槽钢；∠ — 角钢。

    Args:
        section_xlsx_path (str): properties_parameter.xlsx 路径。

    Returns:
        dict: 结构形式
            {截面名称: {"SEC": list[float], "type": str, "A": float,
                       "Ix": float, "Iy": float, "Wx": float, "Wy": float,
                       "ix": float, "iy": float, "每延米重": float, "MCT": str}}；
            路径无效或文件不存在时返回 {}。

    Example:
        >>> lib = load_section_library("properties_parameter.xlsx")
        >>> lib["HW400x400"]["SEC"]
        [400.0, 400.0, 13.0, 21.0, ...]
    """
    if not section_xlsx_path or not os.path.exists(section_xlsx_path):
        return {}

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
        for r in range(2, ws.max_row + 1):
            name = _s(ws.cell(r, 1).value)
            if not name:
                continue
            # 读取截面尺寸
            sec = []
            for dc in config['dim_cols']:
                v = ws.cell(r, dc).value
                try:
                    sec.append(float(v) if v is not None else 0.0)
                except (ValueError, TypeError):
                    sec.append(0.0)
            # 读取力学特性
            library[name] = {
                'SEC': sec,
                'type': sheet_type,
                'A': _f(_cell(ws, r, 'A', hm)),
                'Ix': _f(_cell(ws, r, 'Ix', hm)) if sheet_type != 'O' else _f(_cell(ws, r, 'I', hm)),
                'Iy': _f(_cell(ws, r, 'Iy', hm)) if sheet_type != 'O' else _f(_cell(ws, r, 'I', hm)),
                'Wx': _f(_cell(ws, r, 'Wx', hm)) if sheet_type != 'O' else _f(_cell(ws, r, 'W', hm)),
                'Wy': _f(_cell(ws, r, 'Wy', hm)) if sheet_type != 'O' else _f(_cell(ws, r, 'W', hm)),
                'ix': _f(_cell(ws, r, 'ix', hm)) if sheet_type != 'O' else _f(_cell(ws, r, 'i', hm)),
                'iy': _f(_cell(ws, r, 'iy', hm)) if sheet_type != 'O' else _f(_cell(ws, r, 'i', hm)),
                '每延米重': _f(_cell(ws, r, '每延米重', hm)),
            }
            sec_mct_string = section_mct_trans(sheet_type, name, sec)
            library[name]['MCT'] = sec_mct_string

    wb.close()
    return library


def load_waler_strut_sections(excel_path):
    """从"数据选项"sheet 按表头读取围檩/内支撑截面列表，并从截面库补齐参数。

    读取逻辑：第 1 行为表头(key)，第 2 行起为数据；同一列非空、非 "/" 的值
    收集为 list，以表头为 key 组成字典，再通过固定表头取出围檩与内支撑列表，
    最后调用 load_section_library 补齐每个截面的尺寸/力学参数。

    Args:
        excel_path (str): 参数表 Excel 路径（同目录下查找 properties_parameter.xlsx）。

    Returns:
        tuple: (waler_list, strut_list, waler_dict, strut_dict)
            - waler_list/strut_list (list[str]): 可选截面名称列表；
            - waler_dict/strut_dict (dict): 结构形式
              {截面名称: {"SEC": list[float], "Ix": float, "Wx": float, ...}}，
              未在截面库中找到时对应值为 {}。
            路径无效时返回 ([], [], {}, {})。

    Example:
        >>> wl, sl, wd, sd = load_waler_strut_sections("CofferDam_ParamTable.xlsx")
        >>> wl[:1]
        ['HW400x400']
    """
    if not excel_path or not os.path.exists(excel_path):
        return [], [], {}, {}
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = _find_sheet(wb, ["数据选项"])
    if ws is None:
        wb.close()
        return [], [], {}, {}

    # 读取截面库
    section_dir = os.path.dirname(excel_path).replace('program_param', 'basic_param')
    section_path = os.path.join(section_dir, 'properties_parameter.xlsx')
    if not os.path.exists(section_path):
        # 回退：从注册表路径查找
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\ShuZhiQiaoShi')
            app, _ = winreg.QueryValueEx(key, 'Applocation')
            winreg.CloseKey(key)
            section_path = os.path.join(app, 'Support', 'basic_param', 'properties_parameter.xlsx')
        except:
            pass
    library = load_section_library(section_path)

    def _lookup_section(name):
        """在截面库中查找截面：依次尝试直接匹配、去"钢管桩"后缀、去空格。

        Args:
            name (str): 截面名称。

        Returns:
            dict|None: 命中时返回截面参数字典，否则返回 None。
        """
        clean = str(name).strip()
        if not clean:
            return None
        # 1. 直接匹配
        if clean in library:
            return library[clean]
        # 2. 钢管桩：去掉"钢管桩"后缀
        if clean.endswith('钢管桩'):
            base = clean.replace('钢管桩', '').strip()
            if base in library:
                return library[base]
        # 3. 去掉空格再试
        no_space = clean.replace(' ', '')
        if no_space in library:
            return library[no_space]
        return None

    # ---- 按表头动态读取各列 ----
    # 第1行：表头
    headers = {}
    for c in range(1, ws.max_column + 1):
        val = ws.cell(1, c).value
        if val is not None and str(val).strip():
            headers[c] = str(val).strip()

    # 第2行起：按列收集非空值
    columns_data = {h: [] for h in headers.values()}
    for r in range(2, ws.max_row + 1):
        for c, header_name in headers.items():
            val = ws.cell(r, c).value
            if val is not None:
                s = str(val).strip()
                if s and s != "/":
                    columns_data[header_name].append(s)

    # ---- 通过固定 key 取出围檩和内支撑列表 ----
    WALER_KEY = "围檩截面可选类型"
    STRUT_KEY = "内支撑截面可选类型"
    waler_secs = columns_data.get(WALER_KEY, [])
    strut_secs = columns_data.get(STRUT_KEY, [])

    waler_dict = {}
    for name in waler_secs:
        props = _lookup_section(name)
        waler_dict[name] = props if props else {}

    strut_dict = {}
    for name in strut_secs:
        props = _lookup_section(name)
        strut_dict[name] = props if props else {}

    wb.close()
    return waler_secs, strut_secs, waler_dict, strut_dict


def load_material_library(_excel_path):
    """读取 properties_parameter.xlsx 中 '材质' sheet 的钢材/混凝土/钢筋参数。

    钢材按 5 档厚度分别存储 f/fv/fcc/fy/fu；混凝土存强度标准值/设计值及弹模等；
    钢筋存抗拉/抗压设计值、屈服/极限强度、弹模与伸长率。

    Args:
        _excel_path (str): 参数表 Excel 路径（用于定位同目录的截面库）。

    Returns:
        tuple: (steel_dict, concrete_dict, rebar_dict)，结构形式分别为
            - steel_dict: {钢结构规范名: {"钢结构规范编号": str,
                  "Midas对应钢结构规范编号": str,
                  "牌号参数": {牌号: {"f": {厚度档: float, ...}, "fv": {...},
                                    "fcc": {...}, "fy": {...}, "fu": {...},
                                    "E": float, "G": float, "a": float, "p": float}}}}；
            - concrete_dict: {混凝土规范名: {"混凝土规范编号": str,
                  "Midas对应混凝土规范编号": str,
                  "强度等级参数": {等级: {"fck","ftk","fc","ft","Ec","Gc","a","v": float}}}}；
            - rebar_dict: {钢筋规范名: {"钢筋规范标号": str,
                  "牌号参数": {牌号: {"fy","fy_","fyk","fstk","E","delta": float}}}}。
            找不到文件或 '材质' sheet 时三者均为 {}。

    Example:
        >>> steel, concrete, rebar = load_material_library("CofferDam_ParamTable.xlsx")
        >>> list(steel.keys())[0]
        '《钢结构设计标准》 （GB 50017-2017）'
    """
    if not _excel_path or not os.path.exists(_excel_path):
        return {}, {}, {}

    _section_dir = os.path.dirname(_excel_path).replace('program_param', 'basic_param') if _excel_path else ''
    _section_path = os.path.join(_section_dir, 'properties_parameter.xlsx') if _section_dir else ''
    if not _section_path or not os.path.exists(_section_path):
        try:
            import winreg
            _key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\ShuZhiQiaoShi')
            _app, _ = winreg.QueryValueEx(_key, 'Applocation')
            winreg.CloseKey(_key)
            _section_path = os.path.join(_app, 'Support', 'basic_param', 'properties_parameter.xlsx')
        except:
            pass

    wb = openpyxl.load_workbook(_section_path, data_only=True)
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
        std_name = _s(_cell(ws, r, "钢结构规范名", hm))
        if not std_name:
            continue
        if std_name not in steel_dict:
            steel_dict[std_name] = {
                "钢结构规范编号": _s(_cell(ws, r, "钢结构规范编号", hm)),
                "Midas对应钢结构规范编号": _s(_cell(ws, r, "Midas对应钢结构规范编号", hm)),
                "牌号参数": {},
            }
        grade = _s(_cell(ws, r, "钢材牌号", hm))
        if not grade:
            continue
        # 初始化每个力学属性的厚度字典
        entry = {sk: {} for sk in steel_strength_keys}
        for tv, thk_hdr in zip(steel_thk_keys, steel_thk_headers):
            raw = _s(_cell(ws, r, thk_hdr, hm))
            vals = [x.strip() for x in raw.split(",") if x.strip()]
            for i, sk in enumerate(steel_strength_keys):
                entry[sk][tv] = float(vals[i]) if i < len(vals) else 0.0
        entry["E"] = _f(_cell(ws, r, "钢材弹性模量", hm))
        entry["G"] = _f(_cell(ws, r, "钢材剪切变形模量", hm))
        entry["a"] = _f(_cell(ws, r, "钢材线膨胀系数", hm))
        entry["p"] = _f(_cell(ws, r, "钢材质量密度", hm))
        steel_dict[std_name]["牌号参数"][grade] = entry

    # ── 混凝土 ──
    concrete_dict = {}
    for r in range(3, ws.max_row + 1):
        std_name = _s(_cell(ws, r, "混凝土规范名", hm))
        if not std_name:
            continue
        if std_name not in concrete_dict:
            concrete_dict[std_name] = {
                "混凝土规范编号": _s(_cell(ws, r, "混凝土规范编号", hm)),
                "Midas对应混凝土规范编号": _s(_cell(ws, r, "Midas对应混凝土规范编号", hm)),
                "强度等级参数": {},
            }
        grade = _s(_cell(ws, r, "混凝土强度等级", hm))
        if not grade:
            continue
        concrete_dict[std_name]["强度等级参数"][grade] = {
            "fck": _f(_cell(ws, r, "混凝土轴心抗压强度标准值", hm)),
            "ftk": _f(_cell(ws, r, "混凝土轴心抗拉强度标准值", hm)),
            "fc":  _f(_cell(ws, r, "混凝土轴心抗压强度设计值", hm)),
            "ft":  _f(_cell(ws, r, "混凝土轴心抗拉强度设计值", hm)),
            "Ec":  _f(_cell(ws, r, "混凝土弹性模量", hm)),
            "Gc":  _f(_cell(ws, r, "混凝土剪切变形模量", hm)),
            "a":   _f(_cell(ws, r, "混凝土线膨胀系数", hm)),
            "v":   _f(_cell(ws, r, "混凝土泊松比", hm)),
        }

    # ── 钢筋 ──
    rebar_dict = {}
    for r in range(3, ws.max_row + 1):
        std_name = _s(_cell(ws, r, "钢筋规范名", hm))
        if not std_name:
            continue
        if std_name not in rebar_dict:
            rebar_dict[std_name] = {
                "钢筋规范标号": _s(_cell(ws, r, "钢筋规范标号", hm)),
                "牌号参数": {},
            }
        grade = _s(_cell(ws, r, "钢筋牌号", hm))
        if not grade:
            continue
        rebar_dict[std_name]["牌号参数"][grade] = {
            "fy":    _f(_cell(ws, r, "钢筋抗拉强度设计值", hm)),
            "fy_":   _f(_cell(ws, r, "钢筋抗压强度设计值", hm)),
            "fyk":   _f(_cell(ws, r, "钢筋屈服强度标准值", hm)),
            "fstk":  _f(_cell(ws, r, "钢筋极限强度标准值", hm)),
            "E":     _f(_cell(ws, r, "钢筋弹性模量", hm)),
            "delta": _f(_cell(ws, r, "钢筋总伸长率限值", hm)),
        }

    wb.close()
    return steel_dict, concrete_dict, rebar_dict


def get_soil_sheet_names(excel_path):
    """获取 Excel 中全部土层 sheet 名（排除主表/选项/模板等非土层 sheet）。

    Args:
        excel_path (str): Excel 路径。

    Returns:
        list[str]: 土层 sheet 名列表；文件不存在时返回 []。
    """
    if not excel_path or not os.path.exists(excel_path):
        return []
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    names = [sn for sn in wb.sheetnames if sn not in _NON_SOIL_SHEETS]
    wb.close()
    return names


def load_soil_sheet(excel_path, sheet_name):
    """读取指定土层 sheet，返回土层分层数据。

    Args:
        excel_path (str): Excel 路径。
        sheet_name (str): 土层 sheet 名。

    Returns:
        dict: 结构形式
            {土层名称: [层厚(m), 重度, c(kPa), φ(°), 计算方法,
                        渗透系数, 土层顶承压水头]}；文件/sheet 不存在时返回 {}。
    """
    if not excel_path or not os.path.exists(excel_path):
        return {}
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        return {}
    ws = wb[sheet_name]
    hm = _build_header_map(ws, header_row=1)
    SOIL_HDRS = ["层厚(m)", "重度(kN/m3)",
                 "黏聚力c(kPa)", "内摩擦角φ(°)",
                 "计算方法", "渗透系数(m/d)", "土层顶承压水头(m)"]
    data = {}
    for r in range(2, ws.max_row + 1):
        name = _s(ws.cell(r, hm.get("土层名称", 0)).value) if hm.get("土层名称") else ""
        if not name:
            continue
        vals = []
        for hdr in SOIL_HDRS:
            col = hm.get(hdr)
            v = ws.cell(r, col).value if col else None
            if v is not None:
                s = str(v).strip()
                try:
                    vals.append(float(s))
                except:
                    vals.append(s)
            else:
                vals.append("")
        data[name] = vals
    wb.close()
    return data
