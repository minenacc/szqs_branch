"""
围堰 Excel 参数读写模块 — 表头名方式

数据源：CofferDam_ParamTable.xlsx
- 矩形围堰设计参数 sheet: R2=表头, R3起=每个项目一行, 121列
- 支护桩参数 sheet: R1=表头, R2起=截面数据
- 数据选项 sheet: 每列一个类别
- 连接形式 sheet: R1=大分类, R2=表头, R3起=数据
- 土层-XXX sheet: R1=表头, R2起=土层分层
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
    "牛腿长边布置数量": "/", "牛腿短边布置数量": "/",
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
    HEADERS_IN_ORDER.extend([f"内支撑{_i}对撑材质", f"内支撑{_i}对撑截面", f"内支撑{_i}斜撑材质", f"内支撑{_i}斜撑截面",
        f"内支撑{_i}对撑长边布置(m)", f"内支撑{_i}对撑短边布置(m)",
        f"内支撑{_i}斜撑长边布置(m)", f"内支撑{_i}斜撑短边布置(m)"])


# ═══════════════════════════════════════════════════════════════
# 工具函数
# ═══════════════════════════════════════════════════════════════

def _s(val, default=""):
    if val is None: return default
    return str(val).strip()

def _ss(val, default=""):
    return _normalize_commas(_s(val, default))

def _normalize_commas(val):
    if val is None: return ""
    return str(val).replace("，", ",").replace("、", ",")

def _split(val):
    if not val: return []
    return [x.strip() for x in val.split(",") if x.strip()]

def _f(val, default=0.0):
    if val is None: return default
    try: return float(val)
    except: return default

def _build_header_map(ws, header_row=1):
    hmap = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if v is not None:
            h = str(v).strip()
            if h: hmap[h] = c
    return hmap

def _header_map_in_range(ws, header_row, start_col, end_col):
    hmap = {}
    for c in range(start_col, end_col + 1):
        v = ws.cell(header_row, c).value
        if v is not None:
            h = str(v).strip()
            if h: hmap[h] = c
    return hmap

def _cell(ws, row, header_name, hmap):
    col = hmap.get(header_name)
    if col is None: return None
    return ws.cell(row, col).value

def _find_sheet(wb, names):
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
    """获取 CofferDam_ParamTable.xlsx 路径"""
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
    """读取 CofferDam_ParamTable.xlsx 全部参数

    Returns:
        {
            "projects_dict": {proj_key: proj_dict, ...},
            "project_groups": {"盐宜": ["DT103#", ...], ...},
            "last_proj_key": "盐宜DT103#",
        }
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


def load_material_library(_excel_path):
    """读取 properties_parameter.xlsx 中 '材质' sheet 的钢材/混凝土/钢筋参数

    Returns:
        (steel_dict, concrete_dict, rebar_dict)
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


# ═══════════════════════════════════════════════════════════════
# 数据写入
# ═══════════════════════════════════════════════════════════════

def _clear_from_row(ws, start_row):
    for row in ws.iter_rows(min_row=start_row, max_row=ws.max_row):
        for cell in row:
            cell.value = None


def _clear_columns_from_row(ws, start_row, columns):
    """仅清除指定列号集合中的单元格，保留其他列的数据"""
    for row in ws.iter_rows(min_row=start_row, max_row=ws.max_row):
        for cell in row:
            if cell.column in columns:
                cell.value = None


def _write_cofferdam_soil_sheet(wb, sheet_name, soil_data):
    """写入围堰土层sheet：不存在则新建，存在则覆写（保留表头）"""
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
    """将内存中新增/修改的截面写入 '支护桩参数' sheet，并更新项目中的截面编号

    Args:
        wb: openpyxl workbook 对象
        all_projects: {proj_key: project_dict}，需包含 sections / section_seq_to_name
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
        """确保 sheet 中有指定表头，返回列号"""
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
    try:
        return float(str(val).strip() or "0")
    except Exception:
        return 0.0


def _normalize_row(proj):
    """保存格式归一化（对每个项目行生效）：
    1) 封底协助字段恒为 /
    2) 垫层/封底互斥：厚度>0视为激活（封底优先），未激活方厚度与等级写 /
    3) 施工阶段规则：生成=是 -> 超挖/水面基坑/开挖面降水 = 0.0；=否 -> 施工阶段七列 = /
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
    """将全部项目数据写回 CofferDam_ParamTable.xlsx

    Args:
        excel_path: 文件路径
        all_projects: {proj_key: project_dict}（由 load_all_params 格式）
        project_keys_in_order: 项目写入顺序，None则用 dict 顺序
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
    """从 Excel sheet '0' 读取模板土层数据 → {name: [层厚, 重度, ...]}

    用于新建项目或无有效 soil_sheet_ref 时的回退模板。
    格式与 load_all_params 的 soil_data 一致。
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

# def parse_stage_data(stage_name_str, soil_top_str, water_level_str, support_count_str):
#     """将 Excel 逗号分隔的 4 列字符串解析为 stage_dict

#     Args:
#         stage_name_str: "取土,加撑,取土,…"
#         soil_top_str:   "-1,/,-4,…"
#         water_level_str:"-1,-1,-4,…"
#         support_count_str: "/,1,/,2,…"

#     Returns:
#         {1: {'工况类型':'取土', '基坑内土顶标高':'-1',
#              '基坑内水面标高':'-1', '支撑层数':'/'}, ...}
#     """
#     names = _split(stage_name_str)
#     soil_tops = _split(soil_top_str)
#     water_levels = _split(water_level_str)
#     support_counts = _split(support_count_str)
#     max_len = max(len(names), len(soil_tops), len(water_levels), len(support_counts))
#     result = {}
#     for i in range(max_len):
#         result[i + 1] = {
#             "工况类型": names[i] if i < len(names) else "",
#             "基坑内土顶标高": soil_tops[i] if i < len(soil_tops) else "/",
#             "基坑内水面标高": water_levels[i] if i < len(water_levels) else "/",
#             "支撑层数": support_counts[i] if i < len(support_counts) else "/",
#         }
#     return result


# def serialize_stage_data(stage_dict):
#     """将 stage_dict 序列化为 Excel 4 列逗号分隔字符串

#     Args:
#         stage_dict: {1: {'工况类型':'取土', ...}, 2: {...}, ...}

#     Returns:
#         (stage_name_str, soil_top_str, water_level_str, support_count_str)
#     """
#     if not stage_dict:
#         return "", "", "", ""
#     names = []
#     soil_tops = []
#     water_levels = []
#     support_counts = []
#     for i in sorted(stage_dict.keys()):
#         d = stage_dict[i]
#         names.append(d.get("工况类型", ""))
#         soil_tops.append(d.get("基坑内土顶标高", "/"))
#         water_levels.append(d.get("基坑内水面标高", "/"))
#         support_counts.append(d.get("支撑层数", "/"))
#     return ",".join(names), ",".join(soil_tops), ",".join(water_levels), ",".join(support_counts)


# ═══════════════════════════════════════════════════════════════
# 土层 sheet 查询
# ═══════════════════════════════════════════════════════════════

_NON_SOIL_SHEETS = {"矩形围堰设计参数", "支护桩参数", "数据选项", "连接形式", "0", "1"}


def section_mct_trans(section_type, section_name, section_info):
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
    """读取 properties_parameter.xlsx 截面库

    properties_parameter.xlsx 包含多个 Sheet，每个代表一类截面：
      I/2I/HM/2HM/HN/2HN/HW/2HW — 工字钢/H型钢（C10=H, C11=B, C12=tw, C13=tf1）
      O — 钢管（C7=D, C8=d）
      C/2C — 槽钢（C10=H, C11=B, C12=tw, C13=tf1）
      ∠ — 角钢（C10=H, C11=B, C12=tw）

    Returns:
        {section_name: {'SEC': [H, B, tw, tf], 'Ix': float, 'Wx': float, ...}}
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
    """从数据选项 sheet 按表头动态读取各列截面列表，
    并从 properties_parameter.xlsx 截面库补齐截面尺寸参数。

    读取逻辑：第1行为表头(key)，第2行起为数据。
    同一列非空、非"/"的值收集为 list[str]，以表头为 key 组成 dict。
    然后通过固定的表头 key 取出围檩和内支撑列表。

    Returns: (waler_list: list[str], strut_list: list[str],
              waler_dict: dict, strut_dict: dict)
    其中 dict 格式为 {name: {'SEC': [H,B,tw,tf], 'Ix':..., 'Wx':...}}
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
        """在截面库中查找名称，返回完整 props 或降级解析"""
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


def get_soil_sheet_names(excel_path):
    """获取 Excel 中全部土层 sheet 名（排除主表/选项/模板后的所有 sheet）"""
    if not excel_path or not os.path.exists(excel_path):
        return []
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    names = [sn for sn in wb.sheetnames if sn not in _NON_SOIL_SHEETS]
    wb.close()
    return names


def load_soil_sheet(excel_path, sheet_name):
    """读取指定土层 sheet → {name: [层厚, 重度, ...]}"""
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
