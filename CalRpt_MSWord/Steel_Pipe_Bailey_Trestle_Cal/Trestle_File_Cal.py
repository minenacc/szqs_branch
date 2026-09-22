"""
栈桥计算书 - 文件操作模块
包含：Midas操作、Excel读取、文件对话框等
"""

# 1. 标准库
import os
import sys
import openpyxl
from pathlib import Path

# 3. 本地模块
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from General.Midas import MidasAPI
from General.DataUtils import midasdisttolst
from General.Geometry import polyline_y_at_x
from General.FilePath import qucik_open_file_dialog, qucik_save_file_dialog


#=====================================================================================================================================================================
# Midas相关操作
#=====================================================================================================================================================================

def Open_Midas_mcb(entry_var, save_var=None):
    """打开mcb文件，自动同步保存路径"""
    openfile_path = qucik_open_file_dialog()
    entry_var.set(openfile_path)
    if save_var is not None:
        save_var.set(openfile_path.replace(".mcb", ".docx"))
    if "mcb" in openfile_path:
        arguments = {"Argument": openfile_path}
        MidasAPI("POST", "/doc/open", arguments)
    else:
        print("未识别到有效mcb文件，请重新选择")


def Midas_Analysis(folder_name, folder_path):
    """运行Midas分析"""
    Analysis_result_file_name1 = folder_name.replace(".mcb", ".OUT")
    Analysis_result_file_name2 = folder_name.replace(".mcb", ".CA1")
    print(Analysis_result_file_name1)
    print(Analysis_result_file_name2)
    # 生成对应的路径
    Analysis_result_file_path1 = os.path.join(folder_path, Analysis_result_file_name1)
    Analysis_result_file_path2 = os.path.join(folder_path, Analysis_result_file_name2)
    # 文件运行
    if os.path.exists(Analysis_result_file_path1) and os.path.exists(Analysis_result_file_path2):
        # 两个结果文件存在，跳过不运行
        print("检测到结果文件，正在读取...")
    else:
        # 两个结果文件不存在，运行
        print("未检测到结果文件，开始运行。")
        Analysis_post = MidasAPI("POST", "/doc/Anal", {})


#=====================================================================================================================================================================
# 文件对话框
#=====================================================================================================================================================================

def Cal_path_SaveAs(entry_var):
    """计算书另存为"""
    savefile_path = qucik_save_file_dialog('.docx').replace("/", "\\")
    entry_var.set(savefile_path)


#=====================================================================================================================================================================
# Excel相关操作
#=====================================================================================================================================================================


def get_trestle_projects(excel_path):
    """从Excel读取所有(项目名称, 栈桥编号)组合，供UI下拉列表使用

    Returns:
        list of (project_name, trestle_number)
    """
    wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    ws = wb['栈桥参数']
    projects = []
    for row in ws.iter_rows(min_row=3, max_col=2, values_only=True):
        project, number = row[0], row[1]
        if project and number:
            projects.append((str(project), str(number)))
    wb.close()
    return projects


# ========================= 材质通用库读取 ================================

def _matlib_cell(ws, row, header_name, hmap):
    col = hmap.get(header_name)
    if col is None: return None
    return ws.cell(row, col).value

def _matlib_f(val, default=0.0):
    if val is None: return default
    try: return float(val)
    except: return default

def _matlib_s(val, default=""):
    if val is None: return default
    return str(val).strip()

def _matlib_build_header_map(ws, header_row=1):
    hmap = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if v is not None:
            h = str(v).strip()
            if h: hmap[h] = c
    return hmap

def load_material_library(properties_xlsx_path):
    """读取 properties_parameter.xlsx 中 '材质' sheet 的钢材/混凝土/钢筋参数

    Returns:
        (steel_dict, concrete_dict, rebar_dict)
    """
    if not properties_xlsx_path or not os.path.exists(properties_xlsx_path):
        return {}, {}, {}

    wb = openpyxl.load_workbook(properties_xlsx_path, data_only=True)
    if '材质' not in wb.sheetnames:
        wb.close()
        return {}, {}, {}

    ws = wb['材质']
    hm = _matlib_build_header_map(ws, header_row=2)

    # ── 钢材 ──
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
        std_name = _matlib_s(_matlib_cell(ws, r, "钢结构规范名", hm))
        if not std_name:
            continue
        if std_name not in steel_dict:
            steel_dict[std_name] = {
                "钢结构规范编号": _matlib_s(_matlib_cell(ws, r, "钢结构规范编号", hm)),
                "Midas对应钢结构规范编号": _matlib_s(_matlib_cell(ws, r, "Midas对应钢结构规范编号", hm)),
                "牌号参数": {},
            }
        grade = _matlib_s(_matlib_cell(ws, r, "钢材牌号", hm))
        if not grade:
            continue
        entry = {sk: {} for sk in steel_strength_keys}
        for tv, thk_hdr in zip(steel_thk_keys, steel_thk_headers):
            raw = _matlib_s(_matlib_cell(ws, r, thk_hdr, hm))
            vals = [x.strip() for x in raw.split(",") if x.strip()]
            for i, sk in enumerate(steel_strength_keys):
                entry[sk][tv] = float(vals[i]) if i < len(vals) else 0.0
        entry["E"] = _matlib_f(_matlib_cell(ws, r, "钢材弹性模量", hm))
        entry["G"] = _matlib_f(_matlib_cell(ws, r, "钢材剪切变形模量", hm))
        entry["a"] = _matlib_f(_matlib_cell(ws, r, "钢材线膨胀系数", hm))
        entry["p"] = _matlib_f(_matlib_cell(ws, r, "钢材质量密度", hm))
        steel_dict[std_name]["牌号参数"][grade] = entry

    # ── 混凝土 ──
    concrete_dict = {}
    for r in range(3, ws.max_row + 1):
        std_name = _matlib_s(_matlib_cell(ws, r, "混凝土规范名", hm))
        if not std_name:
            continue
        if std_name not in concrete_dict:
            concrete_dict[std_name] = {
                "混凝土规范编号": _matlib_s(_matlib_cell(ws, r, "混凝土规范编号", hm)),
                "Midas对应混凝土规范编号": _matlib_s(_matlib_cell(ws, r, "Midas对应混凝土规范编号", hm)),
                "强度等级参数": {},
            }
        grade = _matlib_s(_matlib_cell(ws, r, "混凝土强度等级", hm))
        if not grade:
            continue
        concrete_dict[std_name]["强度等级参数"][grade] = {
            "fck": _matlib_f(_matlib_cell(ws, r, "混凝土轴心抗压强度标准值", hm)),
            "ftk": _matlib_f(_matlib_cell(ws, r, "混凝土轴心抗拉强度标准值", hm)),
            "fc":  _matlib_f(_matlib_cell(ws, r, "混凝土轴心抗压强度设计值", hm)),
            "ft":  _matlib_f(_matlib_cell(ws, r, "混凝土轴心抗拉强度设计值", hm)),
            "Ec":  _matlib_f(_matlib_cell(ws, r, "混凝土弹性模量", hm)),
            "Gc":  _matlib_f(_matlib_cell(ws, r, "混凝土剪切变形模量", hm)),
            "a":   _matlib_f(_matlib_cell(ws, r, "混凝土线膨胀系数", hm)),
            "v":   _matlib_f(_matlib_cell(ws, r, "混凝土泊松比", hm)),
        }

    # ── 钢筋 ──
    rebar_dict = {}
    for r in range(3, ws.max_row + 1):
        std_name = _matlib_s(_matlib_cell(ws, r, "钢筋规范名", hm))
        if not std_name:
            continue
        if std_name not in rebar_dict:
            rebar_dict[std_name] = {
                "钢筋规范标号": _matlib_s(_matlib_cell(ws, r, "钢筋规范标号", hm)),
                "牌号参数": {},
            }
        grade = _matlib_s(_matlib_cell(ws, r, "钢筋牌号", hm))
        if not grade:
            continue
        rebar_dict[std_name]["牌号参数"][grade] = {
            "fy":    _matlib_f(_matlib_cell(ws, r, "钢筋抗拉强度设计值", hm)),
            "fy_":   _matlib_f(_matlib_cell(ws, r, "钢筋抗压强度设计值", hm)),
            "fyk":   _matlib_f(_matlib_cell(ws, r, "钢筋屈服强度标准值", hm)),
            "fstk":  _matlib_f(_matlib_cell(ws, r, "钢筋极限强度标准值", hm)),
            "E":     _matlib_f(_matlib_cell(ws, r, "钢筋弹性模量", hm)),
            "delta": _matlib_f(_matlib_cell(ws, r, "钢筋总伸长率限值", hm)),
        }

    wb.close()
    return steel_dict, concrete_dict, rebar_dict


def _read_pile_soil_perpile(wb, label):
    """从 {label}土层表 覆盖表读取 per-pile 土层（每桩一行、列内逗号分层）。

    覆盖表自带每层内联参数（重度/c/φ/qsik/qpk），无需再查土层参数库。
    返回 (soil_param, soil_param_index)；无该 sheet 或无可解析行时返回 (None, None)：
      soil_param:       [{'name','weight','c','phi','qsik','qpk'}, ...] 去重保序
      soil_param_index: [{'index','name','thickness','ground_level'}, ...]
    """
    target = (label + "土层表") if label else ""
    if not target or target not in wb.sheetnames:
        return None, None
    ws = wb[target]
    hm = {}
    for c in ws[1]:
        if c.value:
            hm[str(c.value).strip()] = c.column

    def _split(v):
        return [x.strip() for x in str(v or "").split(",") if x.strip()]

    def _f(v):
        try:
            return float(v) if v not in (None, "") else 0
        except (ValueError, TypeError):
            return 0

    soil_param_index = []
    soil_param_map = {}
    for r in range(2, ws.max_row + 1):
        pile_tag = str(ws.cell(r, hm.get("桩号", 1)).value or "").strip()
        if not pile_tag or not pile_tag.rstrip("#").isdigit():
            continue
        idx = pile_tag.rstrip("#")
        names = _split(ws.cell(r, hm.get("土层名称", 3)).value)
        if not names:
            continue
        thicks = _split(ws.cell(r, hm.get("土层厚度(m)", 4)).value)
        weights = _split(ws.cell(r, hm.get("土层重度(kN/m3)", 5)).value)
        cs = _split(ws.cell(r, hm.get("黏聚力c(kPa)", 6)).value)
        phis = _split(ws.cell(r, hm.get("内摩擦角φ(°)", 7)).value)
        qsiks = _split(ws.cell(r, hm.get("桩极限侧阻力标准值qsik(kPa)", 8)).value)
        qpks = _split(ws.cell(r, hm.get("桩极限端阻力标准值qpk(kPa)", 9)).value)
        ground_raw = ws.cell(r, hm.get("土顶标高(m)", 2)).value
        ground_level = _f(ground_raw)
        for i, nm in enumerate(names):
            key = (nm,
                   weights[i] if i < len(weights) else "",
                   cs[i] if i < len(cs) else "",
                   phis[i] if i < len(phis) else "",
                   qsiks[i] if i < len(qsiks) else "",
                   qpks[i] if i < len(qpks) else "")
            if key not in soil_param_map:
                soil_param_map[key] = {
                    "name": nm,
                    "weight": _f(key[1]), "c": _f(key[2]), "phi": _f(key[3]),
                    "qsik": _f(key[4]), "qpk": _f(key[5]),
                }
        soil_param_index.append({
            "index": idx,
            "name": names,
            "thickness": thicks,
            "ground_level": ground_level,
        })
    if not soil_param_index:
        return None, None
    return list(soil_param_map.values()), soil_param_index


def read_trestle_params_from_excel(excel_path, project_name, trestle_no):
    """从Excel读取栈桥计算书所需的全部参数

    Args:
        excel_path: Trestle_ParamTable.xlsx 路径
        project_name: 项目名称（如"合湛"）
        trestle_no: 栈桥编号（如"卖皂河165#-167#"）

    Returns:
        dict，包含：
            project_name: 项目名称
            trestle_no: 栈桥编号
            beam_type: 桁架纵梁类型
            wind_standard: 风荷载规范名称
            ground_class: 地表分类
            terrain_factor: 地形修正系数
            ref_height: 主梁基准高度(m)
            length_factor: 水平加载长度系数
            wind_formula: 设计基准风速计算公式
            design_wind_speed: 设计风速(m/s)
            flow_standard: 水流力规范名称
            design_flow_speed: 设计水流速(m/s)
            soil_param: 土层参数列表[{name, weight, c, phi, qsik, qpk}, ...]
            vehicle_params_list: 车辆参数列表[{name, tag, type, wheelbase, params}, ...]
    """
    wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)

    # ── 1. 定位目标行 + 读取表头映射 ──
    ws_main = wb['栈桥参数']
    target_row = None
    for row in ws_main.iter_rows(min_row=3, values_only=False):
        if str(row[0].value) == project_name and str(row[1].value) == trestle_no:
            target_row = row
            break
    if target_row is None:
        wb.close()
        raise ValueError(f"未找到项目[{project_name}]编号[{trestle_no}]的参数行")

    col_map = {}
    for row in ws_main.iter_rows(min_row=2, max_row=2, values_only=False):
        for c in row:
            if c.value is not None:
                col_map[str(c.value)] = c.column - 1  # 转为0-based index
        break

    def cell(header):
        """按表头名称获取指定列的值"""
        return target_row[col_map[header]].value

    result = {
        # 基本信息
        'project_name': project_name,
        'trestle_no': trestle_no,
        'beam_type': str(cell('桁架纵梁类型') or ''),
        'deck_level': cell('桥面高程(m)'),
        # 风荷载参数
        'wind_standard': str(cell('风荷载规范') or ''),
        'ground_class': str(cell('地表分类') or ''),
        'terrain_factor': cell('地形条件系数'),
        'ref_height': cell('主梁基准高度(m)'),
        'length_factor': cell('主梁横向力系数'),
        'wind_formula': str(cell('设计基准风速计算公式') or ''),
        'design_wind_speed': cell('设计风速(m/s)'),
        # 水流力参数
        'flow_standard': str(cell('水流力规范') or ''),
        'design_flow_speed': cell('设计水流速(m/s)'),
    }

    # ── 3. 从地形表读取土层参数（"详细土层地形信息"列存 soil_label 前缀，地形表 = {前缀}地形表） ──
    soil_sheet_name = ''
    if '详细土层地形信息' in col_map:
        soil_sheet_name = str(cell('详细土层地形信息') or '')
    result['soil_param'] = []
    if soil_sheet_name and (soil_sheet_name + '地形表') in wb.sheetnames:
        # 3a. 从地形表读取土层编号标签（从第2行到"计算底边线y坐标"前一行），去重保序
        ws_soil = wb[soil_sheet_name + '地形表']
        soil_ids = []
        for row in ws_soil.iter_rows(min_row=2, max_col=1, values_only=True):
            val = str(row[0] or '')
            if val == '计算底边线y坐标':
                break
            if val:
                soil_ids.append(val)
        unique_soil_ids = list(dict.fromkeys(soil_ids))  # 去重保序

        # 3b. 从"土层参数库"查找对应参数
        def _f(v):
            try:
                return float(v) if v not in (None, '') else 0
            except (ValueError, TypeError):
                return 0

        ws_soil_lib = wb['土层参数库']
        soil_db = {}
        for row in ws_soil_lib.iter_rows(min_row=2, values_only=True):
            sid = str(row[0] or '')
            if sid:
                soil_db[sid] = {
                    'name': str(row[1] or ''),
                    'weight': _f(row[2]),   # 重度(kN/m3)
                    'c': _f(row[3]),        # 黏聚力c(kPa)
                    'phi': _f(row[4]),      # 内摩擦角φ(°)
                    'qsik': _f(row[5]),     # 桩极限侧阻力标准值qsik(kPa)
                    'qpk': _f(row[6]),      # 桩极限端阻力标准值qpk(kPa)
                }

        soil_param = []
        for sid in unique_soil_ids:
            if sid in soil_db:
                soil_param.append(soil_db[sid])
        result['soil_param'] = soil_param

    # ── 3c.优先读 {label}土层表（每桩一行、列内逗号分层）；无则按地形插值 ──
    result['soil_param_index'] = []
    override_param, override_index = _read_pile_soil_perpile(wb, soil_sheet_name)
    if override_index:
        result['soil_param'] = override_param
        result['soil_param_index'] = override_index
        print(f"[read_excel] 使用 {soil_sheet_name}土层表 覆盖表: {len(override_index)} 个桩号")
    else:
        try:
            x0 = cell('栈桥起始定位线(0#)x坐标(m)')
            spans_raw = str(cell('计算跨径(m)') or '')
            sub_raw = str(cell('下部结构编号') or '')

            if x0 is not None and spans_raw and sub_raw:
                x0 = float(x0)
                spans = midasdisttolst(spans_raw)
                sub_ids = [s.strip() for s in sub_raw.split(',') if s.strip()]
                n_piles = len(sub_ids)

                # 各桩号 X 坐标：0#=x0, 1#=x0+span[0], 2#=x0+span[0]+span[1], ...
                pile_x = [x0]
                for sp in spans:
                    pile_x.append(pile_x[-1] + sp)
                print(f"[read_excel] 桩号X坐标: {[round(x, 3) for x in pile_x[:n_piles]]}")

                # 读取地形参数矩阵表（{前缀}地形表）
                _topo_sheet = str(cell('详细土层地形信息') or '') if '详细土层地形信息' in col_map else ''
                _topo_sheet = (_topo_sheet + '地形表') if _topo_sheet else ''
                if _topo_sheet and _topo_sheet in wb.sheetnames:
                    ws_topo = wb[_topo_sheet]

                    # 读取表头行的 X 采样坐标（第1行，第2列起）
                    x_samples = []
                    for c_idx in range(2, ws_topo.max_column + 1):
                        hdr_val = ws_topo.cell(1, c_idx).value
                        if hdr_val is not None:
                            try:
                                x_samples.append(float(hdr_val))
                            except (ValueError, TypeError):
                                pass

                    # 读取各土层的坐标数据（第2行起，A列=土层库ID，B列起=(x,y)字符串）
                    topo_layers = []  # [{'id': str, 'coords': [(x, y), ...]}, ...]
                    bottom_y_raw = None  # 计算底边线y坐标
                    for r in range(2, ws_topo.max_row + 1):
                        layer_id = str(ws_topo.cell(r, 1).value or '').strip()
                        if layer_id == '计算底边线y坐标':
                            # 读取底边线Y值（第2列）
                            bv = ws_topo.cell(r, 2).value
                            if bv is not None:
                                try:
                                    bottom_y_raw = float(bv)
                                except (ValueError, TypeError):
                                    pass
                            print(f"[read_excel] 计算底边线y坐标: {bottom_y_raw}")
                            break
                        if not layer_id:
                            continue
                        coords = []
                        for c_idx in range(2, ws_topo.max_column + 1):
                            cell_val = ws_topo.cell(r, c_idx).value
                            if cell_val and cell_val != '/':
                                # 解析 "(x,y)" 格式
                                try:
                                    parts = str(cell_val).strip('()').split(',')
                                    coords.append((float(parts[0]), float(parts[1])))
                                except (ValueError, IndexError):
                                    pass
                        if coords:
                            topo_layers.append({'id': layer_id, 'coords': coords})

                    # 对每个桩号，插值计算各土层在该X处的顶面Y值（复用通用折线插值）
                    soil_param_index = []
                    for i in range(min(n_piles, len(pile_x))):
                        px = pile_x[i]
                        layer_ys = []
                        for layer in topo_layers:
                            y_val = polyline_y_at_x(px, layer['coords'])
                            if y_val is not None:
                                layer_ys.append({'id': layer['id'], 'y': y_val})

                        # 按 y 从大到小排序（从上到下）
                        layer_ys.sort(key=lambda item: item['y'], reverse=True)

                        # 构建 name 和 thickness
                        names = []
                        thicknesses = []
                        for j, ly in enumerate(layer_ys):
                            layer_name = soil_db.get(ly['id'], {}).get('name', ly['id'])
                            names.append(layer_name)
                            if j < len(layer_ys) - 1:
                                thicknesses.append(round(ly['y'] - layer_ys[j + 1]['y'], 3))
                            elif bottom_y_raw is not None:
                                # 最后一层：用计算底边线y坐标作为底界
                                thicknesses.append(round(ly['y'] - bottom_y_raw, 3))

                        ground_level = layer_ys[0]['y'] if layer_ys else 0

                        soil_param_index.append({
                            'index': str(i),
                            'name': names,
                            'thickness': thicknesses,
                            'ground_level': round(ground_level, 3),
                        })

                    result['soil_param_index'] = soil_param_index
                    print(f"[read_excel] soil_param_index 共 {len(soil_param_index)} 个桩号")

        except Exception as e:
            import traceback
            print(f"[read_excel] soil_param_index 构建失败: {e}")
            traceback.print_exc()

    # ── 4. 从"车辆设备库"sheet读取车辆参数 ──
    # 从主表获取移动荷载和静载车辆名称（存的是 车辆/设备名称 列的值）
    moving_vehicles_raw = str(cell('移动荷载车辆/设备') or '')
    static_vehicles_raw = str(cell('静载车辆/设备') or '')
    # 合并两列车辆名，去重保序
    vehicle_names_raw = []
    for raw in [moving_vehicles_raw, static_vehicles_raw]:
        for item in raw.split(','):
            item = item.strip().strip('[]')
            if item:
                vehicle_names_raw.append(item)
    unique_vehicle_names = list(dict.fromkeys(vehicle_names_raw))

    # 从车辆设备库读取完整参数
    ws_vehicle = wb['车辆设备库']
    vcol = {}
    for row in ws_vehicle.iter_rows(min_row=2, max_row=2, values_only=False):
        for c in row:
            if c.value is not None:
                vcol[str(c.value)] = c.column - 1
        break

    def vcell(row, header):
        """按表头名称获取车辆设备库某行指定列的值"""
        return row[vcol[header]].value

    vehicle_db = {}
    for row in ws_vehicle.iter_rows(min_row=3, values_only=False):
        name = str(vcell(row, '车辆/设备名称') or '')      # 名称（主表引用此列）
        vehicle_tag = str(vcell(row, '车辆/设备') or '')    # 标签（区分自定义/标准）
        if name:
            vehicle_db[name] = {
                'tag': vehicle_tag,
                'type': str(vcell(row, '车辆荷载类型') or ''),
                'wheelbase': vcell(row, '车轮间距/轴距(m)'),
                'std_axle_name': str(vcell(row, '标准荷载类型') or ''),
                'crane_weight': vcell(row, '吊装设备自重(kN)'),
                'crane_capacity': vcell(row, '吊重(kN)'),
                'crane_boom': vcell(row, '吊臂(m)'),
                'crane_track_length': vcell(row, '吊装设备接地长度(m)'),
                'crane_track_width': vcell(row, '吊装设备接地宽度(m)'),
                'rotary_weight': vcell(row, '钻孔设备自重(kN)'),
                'rotary_torque': vcell(row, '扭矩反力(kN)'),
                'rotary_track_length': vcell(row, '钻孔设备接地长度(m)'),
                'rotary_track_width': vcell(row, '钻孔设备接地宽度(m)'),
                'std_forces': str(vcell(row, '轴重(kN)') or ''),
                'std_distances': str(vcell(row, '距前轴(m)') or ''),
            }

    # 构建车辆参数列表
    vehicle_params_list = []
    for v_name in unique_vehicle_names:
        if v_name in vehicle_db:
            info = vehicle_db[v_name]
            # 按类型提取非空参数
            params = {}
            if info['type'] == '公路标准荷载':
                params.update({
                    'axle_name': info['std_axle_name'],
                    'forces': info['std_forces'],
                    'distances': info['std_distances'],
                })
            elif info['type'] == '吊装设备':
                params.update({
                    'weight': info['crane_weight'],
                    'capacity': info['crane_capacity'],
                    'boom': info['crane_boom'],
                    'track_length': info['crane_track_length'],
                    'track_width': info['crane_track_width'],
                })
            elif info['type'] == '钻孔设备':
                params.update({
                    'weight': info['rotary_weight'],
                    'torque': info['rotary_torque'],
                    'track_length': info['rotary_track_length'],
                    'track_width': info['rotary_track_width'],
                })
            elif info['type'] == '一般车辆':
                raw_forces = str(info.get('std_forces', ''))
                raw_distances = str(info.get('std_distances', ''))
                force_items = [f.strip() for f in raw_forces.split(',') if f.strip()]
                dist_items = [d.strip() for d in raw_distances.split(',') if d.strip()][:-1]
                # 计算自重
                total_weight = int(sum(float(f) for f in force_items))
                # 拼接计算书显示字符
                forces_formula = f"({'+'.join(force_items)})" if force_items else ""
                distances_formula = f"({'+'.join(dist_items)})" if dist_items else ""
                params.update({
                    'weight': total_weight,                   # 比如: 500
                    'forces_formula': forces_formula,         # 比如: "(80+80+170+170)"
                    'distances_formula': distances_formula,   # 比如: "(1.7+3.5+1.35)"
                    'forces': raw_forces,                     
                    'distances': raw_distances,              
                })
            vehicle_params_list.append({
                'name': v_name,
                'tag': info.get('tag', ''),
                'type': info['type'],
                'wheelbase': info['wheelbase'],
                'params': params,
            })
    result['vehicle_params_list'] = vehicle_params_list

    wb.close()

    # ── 加载材质通用库 ──
    properties_path = os.path.join(os.path.dirname(excel_path), '..', 'basic_param', 'properties_parameter.xlsx')
    properties_path = os.path.normpath(properties_path)
    steel_dict, concrete_dict, rebar_dict = load_material_library(properties_path)
    result['steel_dict'] = steel_dict
    result['concrete_dict'] = concrete_dict
    result['rebar_dict'] = rebar_dict
    print(f"[材质库] 钢材规范 {len(steel_dict)} 本, 混凝土规范 {len(concrete_dict)} 本, 钢筋规范 {len(rebar_dict)} 本")

    return result
