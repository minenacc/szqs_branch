# 1. 标准库
import os
import re
import math
import warnings

# 2. 第三方库
import openpyxl
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# 3. 本地模块
from General.Midas import MidasAPI
from General.FilePath import qucik_open_file_dialog, qucik_save_file_dialog

#=====================================================================================================================================================================

# 在Midas已经启动的前提下打开选定的文件
def Open_Midas_mcb(entry_var, entry_var2):
    # 选择路径
    openfile_path = qucik_open_file_dialog()
    # 更新对应编辑框的值
    entry_var.set(openfile_path)
    entry_var2.set(openfile_path.replace(".mcb", ".docx"))
    # 打开模型
    if "mcb" in openfile_path:
        arguments = {"Argument" : openfile_path}
        MidasAPI("POST" , "/doc/open" , arguments)
    else:
        print("未识别到有效mcb文件, 请重新选择")


# 判断当前mcb文件所在文件夹内是否有计算结果文件，没有则运行
def Midas_Analysis(folder_name, folder_path):
    # 根据文件名生成结果文件名
    Analysis_result_file_name1 = folder_name.replace(".mcb", ".OUT")
    Analysis_result_file_name2 = folder_name.replace(".mcb", ".CA1")
    # print(Analysis_result_file_name1)
    # print(Analysis_result_file_name2)
    # 生成对应的路径
    Analysis_result_file_path1 = os.path.join(folder_path, Analysis_result_file_name1)
    Analysis_result_file_path2 = os.path.join(folder_path, Analysis_result_file_name2)
    MidasAPI("POST", "/doc/Anal", {})


# 计算书另存为得路径
def Cal_path_SaveAs(entry_var, Applocation):
    # 选择路径
    savefile_path = qucik_save_file_dialog('.docx').replace("/", "\\")
    # open_csv_saveas_xlsx(savefile_path, Applocation)
    # 更新对应编辑框的值
    entry_var.set(savefile_path)
    

# 读取截面sec文件，返回sec_db字典
def Load_Sec_Properties(file_path):
    sec_db = {}
    with open(file_path, 'r', encoding='utf-8') as f:
        for line in f:
            parts = line.strip().split()
            if not parts: continue
            name = parts[0]
            # 根据参数个数判断截面类型
            if len(parts) == 6:  # 钢管桩: 名称, D, t, Ix, Wx, A
                sec_db[name] = {
                    'type': 'PIPE',
                    'D': float(parts[1]),
                    't': float(parts[2]),
                    'Ix': float(parts[3]),
                    'Wx': float(parts[4]),
                    'A': float(parts[5])
                }
            elif len(parts) == 9: # 单拼型钢: 名称, H, B, tw, tf, r, Ix, Wx, A
                sec_db[name] = {
                    'type': 'TYPE_BEAM',
                    'H': float(parts[1]),
                    'B': float(parts[2]),
                    'tw': float(parts[3]),
                    'tf': float(parts[4]),
                    'Ix': float(parts[6]),
                    'Wx': float(parts[7]),
                    'A': float(parts[8])
                }
            elif len(parts) == 10: # 双拼型钢: 名称, H, B_total, tw, tf_up, dist, tf_down, Ix, Wx, A
                sec_db[name] = {
                    'type': 'DOUBLE_TYPE_BEAM',
                    'H': float(parts[1]),
                    'B': float(parts[2]),
                    'tw': float(parts[3]),
                    'tf_up': float(parts[4]),
                    'dist': float(parts[5]), # 腹板间距
                    'tf_down': float(parts[6]),
                    'Ix': float(parts[7]),
                    'Wx': float(parts[8]),
                    'A': float(parts[9])
                }
    return sec_db


def parse_blocks_from_lines(lines):
    """
    参数: lines: list of str，文本文件的每一行（末尾换行符已去除）
    返回: dict: { block_name: [list_of_data_lines] }
    """
    blocks = {}
    current_key = None
    current_data = []
    for line in lines:
        line = line.strip()
        # 空行作为块结束分隔符
        if line == '':
            if current_key is not None:
                blocks[current_key] = current_data
                current_key = None
                current_data = []
            continue
        # 标题行：以 '*' 开头
        if line.startswith('*'):
            if current_key is not None:
                blocks[current_key] = current_data
            current_key = line.lstrip('*').strip()
            current_data = []
        else:
            if current_key is not None:
                current_data.append(line)  # 或 current_data.append(line.split())
    # 处理最后可能没有结尾空行的情况
    if current_key is not None:
        blocks[current_key] = current_data
    return blocks


# ═══════════════════════════════════════════════════════════════
# Excel 参数读取（替代 txt 读取）
# ═══════════════════════════════════════════════════════════════

# gamma_0 推导表
GAMMA_0_MAP = {"一级": 1.1, "二级": 1.0, "三级": 0.9}

# 支护桩参数 sheet 的 key list（按 header 名称检索，不受列顺序影响）
STEEL_SHEET_PILE_KEYS = [
    "H(mm)", "B(mm)", "As(mm^2)", "Asy(mm^2)", "Asz(mm^2)",
    "Ixx(mm^4)", "Iyy(mm^4)", "Izz(mm^4)",
    "Cyp(mm)", "Cym(mm)", "Czp(mm)", "Czm(mm)",
    "Qyb(mm^2)", "Qzb(mm^2)",
    "Peri:O(mm)", "Peri:I(mm)",
    "Cent:y(mm)", "Cent:z(mm)",
    "y1(mm)", "z1(mm)", "y2(mm)", "z2(mm)",
    "y3(mm)", "z3(mm)", "y4(mm)", "z4(mm)",
    "Zyy(mm^3)", "Zzz(mm^3)"
]

LOCK_PIPE_PILE_KEYS = [
    "D(mm)", "t(mm)", "锁扣宽度", "钢板桩数量", "钢板桩宽度(mm)"
]


# ---- 工具函数 ----

def _s(val, default=""):
    if val is None: return default
    return str(val).strip()

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


# ---- 项目列表读取（供 GUI 使用）----

def load_all_projects_for_cal(excel_path=None):
    """读取所有项目列表，供 GUI 下拉框使用

    Returns:
        {
            "projects_dict": {proj_key: {"项目名称": ..., "围堰编号": ..., ...}, ...},
            "project_groups": {"项目名称": ["编号1", "编号2", ...], ...},
            "last_proj_key": "项目名称围堰编号",
        }
    """
    if excel_path is None: excel_path = get_excel_path()
    all_projects = {}
    project_groups = {}
    proj_keys_in_order = []

    if not excel_path or not os.path.exists(excel_path):
        print(f"[CalRpt] 找不到 {excel_path}")
        return {"projects_dict": all_projects, "project_groups": project_groups, "last_proj_key": None}

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws_main = _find_sheet(wb, ["矩形围堰设计参数"])
    if ws_main is not None:
        hm = _build_header_map(ws_main, header_row=2)
        for r in range(3, ws_main.max_row + 1):
            pname = _s(_cell(ws_main, r, "项目名称", hm))
            if not pname: continue
            pno = _s(_cell(ws_main, r, "围堰编号", hm))
            proj_key = f"{pname}{pno}" if pno else pname
            proj = {"项目名称": pname, "围堰编号": pno}
            all_projects[proj_key] = proj
            proj_keys_in_order.append(proj_key)
            if pname not in project_groups:
                project_groups[pname] = []
            if pno:
                project_groups[pname].append(pno)
    wb.close()
    last_proj_key = proj_keys_in_order[-1] if proj_keys_in_order else None
    return {"projects_dict": all_projects, "project_groups": project_groups, "last_proj_key": last_proj_key}


# ---- 单项目参数读取（供计算书使用）----

def load_cofferdam_params_for_cal(excel_path, proj_key, acc = 3):
    """读取指定项目的全部围堰参数，供计算书生成使用

    Args:
        excel_path: CofferDam_ParamTable.xlsx 路径
        proj_key: 项目键，如 "盐宜DT103#"

    Returns:
        dict: 包含所有 CalRpt 需要的参数
    """
    if not excel_path or not os.path.exists(excel_path):
        raise FileNotFoundError(f"找不到 Excel 文件: {excel_path}")

    wb = openpyxl.load_workbook(excel_path, data_only=True)

    # ==================== 1. 主表读取 ====================
    ws_main = _find_sheet(wb, ["矩形围堰设计参数"])
    if ws_main is None:
        wb.close()
        raise ValueError("Excel 中找不到 '矩形围堰设计参数' sheet")

    hm = _build_header_map(ws_main, header_row=2)

    # 定位目标行
    target_row = None
    for r in range(3, ws_main.max_row + 1):
        pname = _s(_cell(ws_main, r, "项目名称", hm))
        pno = _s(_cell(ws_main, r, "围堰编号", hm))
        key = f"{pname}{pno}" if pno else pname
        if key == proj_key:
            target_row = r
            break

    if target_row is None:
        wb.close()
        raise ValueError(f"在 Excel 中找不到项目: {proj_key}")

    r = target_row

    # ---- 基本参数 ----
    SECtype = _s(_cell(ws_main, r, "支护桩类型", hm))
    pile_material = _s(_cell(ws_main, r, "支护桩材质", hm))
    Solid_Level = _f(_cell(ws_main, r, "初始地面标高(m)", hm))
    # Water_Level = _f(_cell(ws_main, r, "初始水位标高(m)", hm))
    Cap_X = _f(_cell(ws_main, r, "承台长(m)", hm))
    Cap_Y = _f(_cell(ws_main, r, "承台宽(m)", hm))
    Cap_H = _f(_cell(ws_main, r, "承台高(m)", hm))
    Cap_Bottom_Level = _f(_cell(ws_main, r, "承台底标高(m)", hm))
    CofferDam_Top_Level = _f(_cell(ws_main, r, "围堰顶标高", hm))
    CofferDam_Bottom_Level = _f(_cell(ws_main, r, "围堰底标高", hm))
    CofferDam_L = _f(_cell(ws_main, r, "围堰高", hm))
    CofferDam_a = _f(_cell(ws_main, r, "围堰内边长(m)", hm))
    CofferDam_b = _f(_cell(ws_main, r, "围堰内边宽(m)", hm))
    Ea_p0_raw = _cell(ws_main, r, "均布附加荷载", hm)
    # 初始水位标高'/'处理
    Water_Level = _cell(ws_main, r, "初始水位标高(m)", hm)
    print(Water_Level)
    Water_Level = _f(Water_Level) if Water_Level != "/" else round(float(CofferDam_Bottom_Level) - 1.0, acc)
    print(Water_Level)
    # 解析格式："(1,2,3,4), 20" → 提取最后一个数值作为 Ea_p0
    if isinstance(Ea_p0_raw, str):
        # 按逗号分割，取最后一个非空部分
        parts = [p.strip() for p in Ea_p0_raw.split(",")]
        # 从后往前找第一个能转为数字的部分
        for p in reversed(parts):
            p = p.rstrip(")")
            try:
                Ea_p0 = float(p)
                break
            except ValueError:
                continue
        else:
            Ea_p0 = 0.0
    else:
        Ea_p0 = _f(Ea_p0_raw)
    CofferDam_Level = _s(_cell(ws_main, r, "围堰安全等级", hm))
    has_stage_str = _s(_cell(ws_main, r, "是否生成施工阶段", hm))
    Drawdown_height = _f(_cell(ws_main, r, "超挖深度(m)", hm))
    Waterdown_height = _f(_cell(ws_main, r, "水面至基坑底距离(m)", hm))
    Excavation_face_dewater = _f(_cell(ws_main, r, "开挖面降水高度(m)", hm))
    Basic_Solid_Layer_Thickness = _f(_cell(ws_main, r, "土弹簧分层厚度(m)", hm))

    # ---- gamma_0 推导 ----
    gamma_0 = GAMMA_0_MAP.get(CofferDam_Level, 1.0)

    # ---- 垫层/封底判断 ----
    seal_t = _f(_cell(ws_main, r, "封底厚度(m)", hm))
    cushion_t = _f(_cell(ws_main, r, "垫层厚度(m)", hm))
    if seal_t != 0:
        conc_type = "混凝土封底"
        conc_grade = _s(_cell(ws_main, r, "封底混凝土等级", hm))
        conc_t = seal_t
    else:
        conc_type = "混凝土垫层"
        conc_grade = _s(_cell(ws_main, r, "垫层混凝土等级", hm))
        conc_t = cushion_t

    # ---- 截面编号 ----
    section_seq_str = _s(_cell(ws_main, r, "支护桩截面编号", hm))
    try:
        section_seq = int(section_seq_str)
    except:
        section_seq = 1

    # ==================== 2. 支护桩参数 sheet ====================
    ws_sec = _find_sheet(wb, ["支护桩参数", "支护桩库", "支护桩"])
    sec_hm = _build_header_map(ws_sec, header_row=2) if ws_sec else {}

    # 通过截面号查找截面名称
    gbz_name = ""
    gbz_sec_info_dict = {}
    pile_b = 0.0
    ggz_D = ggz_t = ggz_name = ggz_A = ggz_I = ggz_W = None

    if ws_sec is not None:
        # 根据支护桩类型选择 key list
        if SECtype == "钢板桩":
            pile_keys = STEEL_SHEET_PILE_KEYS
        elif SECtype == "锁扣钢管桩":
            pile_keys = LOCK_PIPE_PILE_KEYS

        for sr in range(3, ws_sec.max_row + 1):
            seq = _cell(ws_sec, sr, "截面号", sec_hm)
            if seq is not None:
                try:
                    if int(seq) == section_seq:
                        gbz_name = _s(_cell(ws_sec, sr, "截面名称", sec_hm))
                        # 按 header 名称检索读取参数，不受列顺序影响
                        sec_params = []
                        for key in pile_keys:
                            val = _f(_cell(ws_sec, sr, key, sec_hm))
                            sec_params.append(val)
                        gbz_sec_info_dict = {'SEC': sec_params}

                        if SECtype == "钢板桩":
                            pile_b = sec_params[1]  # B(mm)
                        elif SECtype == "锁扣钢管桩":
                            ggz_D = sec_params[0]  # D(mm)
                            ggz_t = sec_params[1]  # t(mm)
                            ggz_name = f"φ{ggz_D}x{ggz_t}"
                            ggz_A = round((math.pi / 4 * (ggz_D**2 - (ggz_D - 2*ggz_t)**2)), 1)
                            ggz_I = round((math.pi / 64 * (ggz_D**4 - (ggz_D - 2*ggz_t)**4)), 1)
                            ggz_W = round((ggz_I / (ggz_D / 2)), 1)
                            pile_b = ggz_D  # 锁扣钢管桩 pile_b = 直径D
                        break
                except (ValueError, TypeError):
                    continue

    # ==================== 3. 土层数据读取 ====================
    soil_ref_raw = _s(_cell(ws_main, r, "地质参数", hm))
    soil_ref = ""
    if soil_ref_raw:
        parts = soil_ref_raw.split(",")
        for p in parts:
            p = p.strip()
            if p in wb.sheetnames:
                soil_ref = p
                break
        if not soil_ref:
            soil_ref = parts[0].strip() if parts and parts[0].strip() in wb.sheetnames else ""

    Solid_data_dict = {}
    Excel_Data = [['土层名称', '层厚(m)', '重度(kN/m3)', '黏聚力c(kPa)', '内摩擦角φ(°)', '计算方法', '渗透系数(m/d)', '土层顶承压水头(m)']]
    solid_layer_lst = []
    Method_for_Pressures_var = "水土分算"  # 默认值
    calc_methods = []

    # 加权平均土层参数：加权列可解析出完整 γ,c,φ 且 γ>0 时，采用加权单层土（固定水土合算）
    def _parse_wavg(raw):
        if raw is None or str(raw).strip() in ('', '/', 'None'):
            return None
        parts = [p.strip() for p in str(raw).replace('，', ',').split(',')]
        if len(parts) >= 3:
            try:
                return (float(parts[0]), float(parts[1]), float(parts[2]))
            except (ValueError, TypeError):
                return None
        return None

    _wavg = _parse_wavg(_s(_cell(ws_main, r, "加权平均重度+粘聚力+摩擦角", hm)))
    if _wavg is not None and _wavg[0] <= 0:
        print(f"[Cal] {proj_key} 加权平均重度({_wavg[0]})不合法(须>0)，已退回按分层土层计算")
        _wavg = None

    if _wavg is not None:
        _wg, _wc_, _wf_ = _wavg
        _thick = round(Solid_Level - CofferDam_Bottom_Level + 1, acc)
        _wavg_name = '加权平均土层'
        Solid_data_dict['第1层土'] = {
            '土层名称': _wavg_name,
            '层厚': _thick,
            '重度': round(_wg, acc),
            '黏聚力': round(_wc_, acc),
            '内摩擦角': round(_wf_, acc),
            '计算方法': '水土合算',
            '渗透系数': 0,
            '土层顶承压水头': 0,
        }
        Excel_Data.append([_wavg_name, _thick, round(_wg, acc), round(_wc_, acc), round(_wf_, acc), '水土合算', 0, 0])
        solid_layer_lst.append([_thick, round(_wg, acc), round(_wc_, acc), round(_wf_, acc), '水土合算', 0, 0])
        Method_for_Pressures_var = "水土合算"
    elif soil_ref and soil_ref in wb.sheetnames:
        ws_soil = wb[soil_ref]
        soil_hm = _build_header_map(ws_soil, header_row=1)
        for sr in range(2, ws_soil.max_row + 1):
            name = _s(_cell(ws_soil, sr, "土层名称", soil_hm))
            if not name: continue
            thickness = _f(_cell(ws_soil, sr, "层厚(m)", soil_hm))
            unit_weight = _f(_cell(ws_soil, sr, "重度(kN/m3)", soil_hm))
            cohesion = _f(_cell(ws_soil, sr, "黏聚力c(kPa)", soil_hm))
            friction = _f(_cell(ws_soil, sr, "内摩擦角φ(°)", soil_hm))
            calc_method = _s(_cell(ws_soil, sr, "计算方法", soil_hm))
            hydraulic_conductivity = _s(_cell(ws_soil, sr, "渗透系数(m/d)", soil_hm))
            soil_water_head = _s(_cell(ws_soil, sr, "土层顶承压水头(m)", soil_hm))

            layer_idx = len(Solid_data_dict) + 1
            Solid_data_dict[f'第{layer_idx}层土'] = {
                '土层名称': name,
                '层厚': thickness,
                '重度': unit_weight,
                '黏聚力': cohesion,
                '内摩擦角': friction,
                '计算方法': calc_method,
                '渗透系数': hydraulic_conductivity,
                '土层顶承压水头': soil_water_head,

            }
            Excel_Data.append([name, thickness, unit_weight, cohesion, friction, calc_method, hydraulic_conductivity, soil_water_head])
            solid_layer_lst.append([thickness, unit_weight, cohesion, friction, calc_method, hydraulic_conductivity, soil_water_head])

            if calc_method:
                calc_methods.append(calc_method)

        # Method_for_Pressures_var 推导
        if any(m == "水土合算" for m in calc_methods):
            Method_for_Pressures_var = "水土合算"

    # ==================== 4. 施工阶段读取 ====================
    stage_name_str = _s(_cell(ws_main, r, "各工况名称", hm))
    stage_soil_str = _s(_cell(ws_main, r, "各工况基坑内土顶标高(m)", hm))
    cstage_water_str = _s(_cell(ws_main, r, "各工况基坑内水面标高(m)", hm))
    stage_name_lst = stage_name_str.split(',')
    stage_soil_lst = stage_soil_str.split(',')
    cstage_water_lst = cstage_water_str.split(',')
    stage_dict = {}
    for i in range(len(stage_name_lst)):
        stage_dict[i+1] = {
            '工况类型': stage_name_lst[i],
            '基坑内土顶标高': stage_soil_lst[i],
            '基坑内水面标高': cstage_water_lst[i],
        }
    for k, v in stage_dict.items():
        print(k)
        print(v)
        if v['基坑内水面标高'] == "/":
            v['基坑内水面标高'] = round(float(CofferDam_Bottom_Level) - 1.0, acc)
    for k, v in stage_dict.items():
        print(k)
        print(v)

    # ==================== 5. 规范读取 ====================
    steel_standard = _s(_cell(ws_main, r, "钢结构规范", hm))
    steel_standard_name = re.search(r'《([^》]+)》', steel_standard).group(1)

    wb.close()
    # ==================== 派生量计算 ====================
    Concrete_Bottom_Level = round((Cap_Bottom_Level - conc_t), acc)  # 混凝土底标高
    Ld = round((Concrete_Bottom_Level - CofferDam_Bottom_Level), acc)  # 嵌固深度
    Foundation_Pit_Depth = round((max(Solid_Level, Water_Level) - Concrete_Bottom_Level), acc)  # 基坑深度
    slice_width = 0.4  # 圆弧滑动稳定性条分宽度，单位m

    return {
        # 基本参数
        "SECtype": SECtype,
        "pile_material": pile_material,
        "Solid_Level": Solid_Level,
        "Water_Level": Water_Level,
        "Cap_X": Cap_X,
        "Cap_Y": Cap_Y,
        "Cap_H": Cap_H,
        "Cap_Bottom_Level": Cap_Bottom_Level,
        "CofferDam_Top_Level": CofferDam_Top_Level,
        "CofferDam_Bottom_Level": CofferDam_Bottom_Level,
        "CofferDam_L": CofferDam_L,
        "CofferDam_a": CofferDam_a,
        "CofferDam_b": CofferDam_b,
        "Ea_p0": Ea_p0,
        "CofferDam_Level": CofferDam_Level,
        "Drawdown_height": Drawdown_height,
        "Waterdown_height": Waterdown_height,
        "Excavation_face_dewater": Excavation_face_dewater,
        "conc_type": conc_type,
        "conc_grade": conc_grade,
        "conc_t": conc_t,
        "Basic_Solid_Layer_Thickness": Basic_Solid_Layer_Thickness,
        "gamma_0": gamma_0,
        "Method_for_Pressures_var": Method_for_Pressures_var,
        # 派生量
        "Concrete_Bottom_Level": Concrete_Bottom_Level,
        "Ld": Ld,
        "Foundation_Pit_Depth": Foundation_Pit_Depth,
        "slice_width": slice_width,
        # 截面参数
        "gbz_name": gbz_name,
        "gbz_sec_info_dict": gbz_sec_info_dict,
        "ggz_D": ggz_D,
        "ggz_t": ggz_t,
        "ggz_name": ggz_name,
        "ggz_A": ggz_A,
        "ggz_I": ggz_I,
        "ggz_W": ggz_W,
        "pile_b": pile_b,
        # 土层数据
        "Solid_data_dict": Solid_data_dict,
        "Excel_Data": Excel_Data,
        "solid_layer_lst": solid_layer_lst,
        # 施工阶段
        "stage_dict": stage_dict,
        # 规范
        "steel_standard": steel_standard_name,
    }


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