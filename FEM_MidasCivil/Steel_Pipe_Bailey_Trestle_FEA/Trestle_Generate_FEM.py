# 1. 标准库
import re
import sys
from pathlib import Path

# 2. 第三方库

# 3. 本地模块
sys.path.append(str(Path(__file__).parent.parent.parent))
from General.Formula import cal_FengYa_JTG_T_3360_01_2018, cal_m, cal_Ks
from General.Midas import MidasURL, importmct_to_mcb, group_arithmetic_sequences
from General.FilePath import file_extension_Modified, write_file_within_path
from General.DataUtils import midasdisttolst, midasdisttolst2, SET_CLIP_STRING
from Trestle_Excel_io_FEM import _parse_material_value, _norm_sec_name

# =====================================================================================================================================================================================================================
# =====================================================================================================================================================================================================================
# =====================================================================================================================================================================================================================



# ============================== 0.辅助函数 ==============================

def _get_bridge_type(components_dict):
    """取components的第一个键作为当前桥型名（消除硬编码）"""
    if not components_dict:
        raise KeyError("components dict is empty")
    return next(iter(components_dict))


# 桥型分类：型钢 / 上承式桁架 / 下承式桁架（预留）
BRIDGE_CATEGORY = {
    "上承式桁架梁栈桥": "truss_upper",
    "下承式桁架梁栈桥": "truss_lower",   # 暂不支持，预留分支
    "型钢栈桥":          "steel",
}


def _bridge_category(bridge_type):
    """桥型分类：steel / truss_upper / truss_lower（未知类型按上承式桁架兜底）"""
    return BRIDGE_CATEGORY.get(bridge_type, "truss_upper")


# 上承式桁架梁栈桥纵梁类型配置：几何/截面/材质（当前仅 321 型贝雷梁，新增类型在此扩展）
TRUSS_BEAM_TYPES = {
    "321型贝雷梁": {
        "chord_sec": "2C10",       # 弦杆截面
        "web_sec": "I8",           # 竖杆/斜杆截面
        "truss_height": 1400,      # 桁架高度 mm
        "chord_height": 100,       # 弦杆截面高 mm（Z 定位、下弦半高用）
        "material": "16Mn",        # 桁架材质
        "piece": {                 # 贝雷片几何（3m/1.5m）
            "3m":   {"j0": [0, 3000], "j90": [90, 1500, 2910], "j795": [795, 2205]},
            "1.5m": {"j0": [0, 1500], "j90": [90, 1410], "j750": [750]},
        },
    },
}


def _truss_beam_cfg(truss_beam_type):
    """取纵梁类型配置；未知类型回退 321 型贝雷梁"""
    return TRUSS_BEAM_TYPES.get(truss_beam_type, TRUSS_BEAM_TYPES["321型贝雷梁"])


def _get_component(ui_dict, field, default=""):
    """从当前桥型的components中取字段值"""
    bridge_type = _get_bridge_type(ui_dict.get('components', {}))
    comp = ui_dict.get('components', {}).get(bridge_type, {})
    return comp.get(field, default)


def _parse_a_bx(value, default_a=0):
    """解析 'a,b@x' 格式字符串 → (a, b@x_part)

    新UI中 pile_trans_space, pile_long_space, rib_trans_space, rib_long_space
    统一使用此格式：
      '750,2@3000' → (750.0, '2@3000')
      '2@3000'     → (default_a, '2@3000')
      '/', None    → (default_a, '/')
    """
    if not value or str(value).strip() in ('', '/'):
        return float(default_a), '/'
    s = str(value).strip()
    if ',' in s:
        parts = s.split(',', 1)
        try:
            return float(parts[0]), parts[1]
        except ValueError:
            return float(default_a), parts[1]
    return float(default_a), s


def _parse_parenthesized_groups(txt):
    """解析括号分组 "(0,22@750)+(0,20@750)" → ["0,22@750", "0,20@750"]"""
    if '(' not in txt:
        return [txt]
    groups = []
    depth = 0
    buf = ""
    for ch in txt:
        if ch == '(':
            depth = 1
            buf = ""
        elif ch == ')':
            if depth == 1 and buf:
                groups.append(buf)
            depth = 0
            buf = ""
        elif depth == 1:
            buf += ch
    # 无括号匹配时返回原始字符串
    if not groups:
        return [txt]
    return groups


def _get_sec_id(sec_dict, sec_ref, default=1):
    """从 component/substructure 字段值查 MCT 截面序号

    section dict 统一以归一化名称为 key（load_all_params 转换，保持 Excel 编号顺序），
    SECTIONS() 按插入位置 1,2,3... 编号，因此：
    - 截面名称 '630X8' → 匹配 key 或记录 name 字段 → 返回插入位置
    - '/' 或 '' → 返回 default
    """
    if not sec_ref or str(sec_ref).strip() in ('', '/'):
        return default
    # 名称匹配（key 或记录 name 字段）
    norm = _norm_sec_name(str(sec_ref)).upper()
    for pos, (k, rec) in enumerate(sec_dict.items(), start=1):
        if _norm_sec_name(k).upper() == norm or _norm_sec_name(rec.get('name', '')).upper() == norm:
            return pos
    print(f"  [警告] 截面 '{sec_ref}' 未在截面库中找到，使用默认序号 {default}")
    return default


def _get_sec_H(sec_dict, sec_ref, default=200):
    """查截面高度(mm)，ref 支持名称或编号（编号=插入位置）"""
    sid = _get_sec_id(sec_dict, sec_ref)
    recs = list(sec_dict.values())
    if 1 <= sid <= len(recs):
        try:
            return float(recs[sid - 1]['params'].get('H', default))
        except (KeyError, ValueError, TypeError):
            pass
    return float(default)


def _sec_D(sections_dict, sec_ref, default=820):
    """查截面外径 D(mm)。ref 支持名称或编号（键兼容 str/int 两种格式）"""
    v = str(sec_ref).strip() if sec_ref else ""
    if not v or v == "/":
        return float(default)
    rec = None
    # 1. 编号直查（键可能是 str 或 int）
    for k, r in sections_dict.items():
        if str(k) == v and isinstance(r, dict):
            rec = r
            break
    # 2. 名称匹配（归一化 ×/x/＊ 为 X）
    if rec is None:
        norm = v.replace("×", "X").replace("Ｘ", "X").replace("x", "X").replace("*", "X")
        for r in sections_dict.values():
            if isinstance(r, dict) and str(r.get("name", "")).replace(
                    "×", "X").replace("Ｘ", "X").replace("x", "X").replace("*", "X").upper() == norm.upper():
                rec = r
                break
    if rec is not None:
        try:
            return float(rec.get("params", {}).get("D", default))
        except (ValueError, TypeError):
            pass
    return float(default)


_STEEL_MCT = {
    '16Mn': 'STEEL, 16Mn, 0, 0, , C, NO, 0.02, 1, JTJ(S), , 16Mn, NO, 2.1414e+07',
    'Q235': 'STEEL, Q235, 0, 0, , C, NO, 0.02, 1, GB50017-17(S), , Q235, NO, 2.10062e+07',
    'Q295': 'STEEL, Q295, 0, 0, , C, NO, 0.02, 1, GB50017-17(S), , Q295, NO, 2.10062e+07',
    'Q345': 'STEEL, Q345, 0, 0, , C, NO, 0.02, 1, GB50017-17(S), , Q345, NO, 2.10062e+07',
}
# 混凝土标号 → 弹性模量(MPa)
_CONCRETE_E = {
    'C30': 3.00e4, 'C35': 3.15e4, 'C40': 3.25e4, 'C45': 3.35e4, 
    'C50': 3.45e4, 'C55': 3.55e4, 'C60': 3.60e4,
}


def _extract_brand(material_val):
    """从 "规范编号,牌号" 格式中提取牌号；空值返回 None。"""
    if not material_val:
        return None
    _, brand = _parse_material_value(material_val)
    return brand or None


def _mat_id(trestle_ui_dict, brand, default='2'):
    """从 material_id_map 查材质 ID，回退到默认值。"""
    if not brand:
        return default
    m = trestle_ui_dict.get('_material_id_map', {})
    return m.get(brand, default)


def MATERIAL(sections_dict, concrete_grade=None,
             steel_dict=None, concrete_dict=None,
             component_brands=None):
    """动态构建材质列表。

    ID 分配规则：
      1 → 16Mn —— SECTIONS() 末尾硬编码的贝雷片(2C10/I8)和虚拟梁
                   对应元素生成处写死 mat_id='1'，故必须占首位
      2+ → 用户截面中出现的其它钢材（按首次出现顺序）
      末尾 → 混凝土（如有）

    Args:
        sections_dict: trestle_ui_dict['section']，每条含 'type', 'name', 'params' 字段
        concrete_grade: 混凝土标号，如 'C30'；None 表示不使用混凝土
        steel_dict: 从 load_material_library() 获取的钢材字典（可选，优先使用）
        concrete_dict: 从 load_material_library() 获取的混凝土字典（可选，优先使用）

    Returns:
        (MATERIAL_lst, material_id_map)
        material_id_map: {材料名: str(ID)}，供下游 _get_mat_id 查询
    """
    material_id_map = {}
    MATERIAL_lst = []

    # 16Mn 固定 ID=1（硬编码截面：贝雷片 2C10、I8、虚拟梁）
    material_id_map['16Mn'] = '1'
    MATERIAL_lst.append('1, ' + _STEEL_MCT['16Mn'])

    # 2. 收集构件级别材质（从 component_brands 获取）
    next_id = 2
    for mat_name in (component_brands or []):
        if mat_name not in material_id_map:
            mct_body = _build_steel_mct_from_db(mat_name, steel_dict)
            if mct_body is None:
                mct_body = _STEEL_MCT.get(mat_name)
            if mct_body is None:
                print(f"[材质] 未知钢材 '{mat_name}'，按 Q235 处理")
                mct_body = _STEEL_MCT['Q235']
                mat_name = 'Q235'
            if mat_name not in material_id_map:
                material_id_map[mat_name] = str(next_id)
                MATERIAL_lst.append('{}, {}'.format(next_id, mct_body))
                next_id += 1

    # 3. 收集实际用到的混凝土
    if concrete_grade:
        mct_body = _build_concrete_mct_from_db(concrete_grade, concrete_dict)
        if mct_body is None:
            # 回退到硬编码字典
            if concrete_grade in _CONCRETE_E:
                mct_body = 'CONC, 0, 0, , C, NO, 0.05, 1, TB10092-17(RC), , {}, NO, {}'.format(
                    concrete_grade, _CONCRETE_E[concrete_grade])
        if mct_body:
            material_id_map['CONCRETE'] = str(next_id)
            MATERIAL_lst.append('{}, {}'.format(next_id, mct_body))

    return MATERIAL_lst, material_id_map


def _build_steel_mct_from_db(brand, steel_dict):
    """从 steel_dict 查找牌号，构建钢材 MCT 行。未找到返回 None。"""
    if not steel_dict:
        return None
    for spec_data in steel_dict.values():
        if brand in spec_data.get("牌号参数", {}):
            params = spec_data["牌号参数"][brand]
            midas_code = spec_data.get("Midas对应钢结构规范编号", "GB50017-17(S)")
            E = params.get("E", 206000)
            return 'STEEL, {}, 0, 0, , C, NO, 0.02, 1, {}, , {}, NO, {}'.format(
                brand, midas_code, brand, float(E))
    return None


def _build_concrete_mct_from_db(grade, concrete_dict):
    """从 concrete_dict 查找等级，构建混凝土 MCT 行。未找到返回 None。"""
    if not concrete_dict:
        return None
    for spec_data in concrete_dict.values():
        if grade in spec_data.get("强度等级参数", {}):
            params = spec_data["强度等级参数"][grade]
            midas_code = spec_data.get("Midas对应混凝土规范编号", "TB10092-17(RC)")
            Ec = params.get("Ec", 30000)
            return 'CONC, 0, 0, , C, NO, 0.05, 1, {}, , {}, NO, {}'.format(
                midas_code, grade, float(Ec))
    return None

# ============================== 2.截面 ============================== 
# 截面的mct格式
def section_mct_trans(section_dict):
    section_type = section_dict['type']
    section_name = section_dict['name']
    section_info = section_dict['params']
    if section_type in ('HW', 'HM', 'HN', 'I'):
        mct_str = 'DBUSER, {}, CC, 0, 0, 0, 0, 0, 0, YES, NO, H , 2, {}, {}, {}, {}, 0, 0, {}, 0, 0, 0'.format(section_name, section_info['H'], section_info['B'], section_info['tw'], section_info['tf1'], section_info['r1'])
    elif section_type == 'C':
        mct_str = 'DBUSER, {}, CC, 0, 0, 0, 0, 0, 0, YES, NO, C , 2, {}, {}, {}, {}, 0, {}, 0, 0, 0, 0'.format(section_name, section_info['H'], section_info['B'], section_info['tw'], section_info['tf1'], section_info['tf2'])
    elif section_type == 'O':
        mct_str = 'DBUSER, {}, CC, 0, 0, 0, 0, 0, 0, YES, NO, P , 2, {}, {}, 0, 0, 0, 0, 0, 0, 0, 0'   .format(section_name, section_info['D'], section_info['d'])
    elif section_type in ('2HW', '2HM', '2HN', '2I'):
        mct_str = 'DBUSER, {}, CC, 0, 0, 0, 0, 0, 0, YES, NO, B , 2, {}, {}, {}, {}, {}, {}, 0, 0, 0, 0, NO, 0, NO, 0, 0'.format(section_name, section_info['H'], section_info['B'], section_info['tw'], section_info['tf1'], section_info['C'], section_info['tf2'])
    elif section_type == '2C':
        mct_str = 'DBUSER, {}, CC, 0, 0, 0, 0, 0, 0, YES, NO, 2C, 2, {}, {}, {}, {}, {}, 0, 0, 0, 0, 0'.format(section_name, section_info['H'], section_info['B'], section_info['tw'], section_info['tf1'], section_info['C'])
    return mct_str


# 截面
def SECTIONS(sections_dict):
    section_list = []
    # 用户选择的截面
    i = 1
    for value in sections_dict.values():
        mct_str = section_mct_trans(value)
        section_list.append(str(i) + ', ' + mct_str)
        i += 1
    # 默认贝雷片截面
    section_list.append("{}, DBUSER, 2C10 , CC, 0, 0, 0, 0, 0, 0, YES, NO, 2C , 2, 100, 48, 5.3, 8.5, 80, 0, 0, 0, 0, 0".format(i))
    section_list.append("{}, DBUSER, I8   , CC, 0, 0, 0, 0, 0, 0, YES, NO, H , 2, 80, 50, 4.5, 6.5, 0, 0, 0, 0, 0, 0".format(i+1))
    # 添加虚拟梁截面
    section_list.append('{}, DBUSER, 虚拟梁, CC, 0, 0, 0, 0, 0, 0, YES, NO, SR, 2, 10, 0, 0, 0, 0, 0, 0, 0, 0, 0'.format(i+2))
    return section_list

# ============================== 3.跨径计算 ============================== 
def trestle_span(trestle_ui_dict):
    '''
    根据制动墩判断栈桥联数 并 获取每一联的墩号 以及 每一联的跨度
    '''
    # span 变量 3+4@12+3@9
    span = trestle_ui_dict['basic']['span']
    # print(span)
    # 转换成跨径 [3, 12, 12, 12, 12, 9, 9, 9]
    span_lst = midasdisttolst(span)
    # print(span_lst)
    # 每一个序号对应的下部结构类型 ['桥台', '单排桩', '制动墩', '单排桩', '单排桩', '制动墩', '单排桩', '单排桩', '桥台']
    substructure_namelst = [value['type'] for value in trestle_ui_dict['substructure'].values()]
    # print(substructure_namelst)
    # 制动墩所在的序号 [2, 5]
    zdd_index_lst = [i for i, x in enumerate(substructure_namelst) if x == '制动墩']
    # print(zdd_index_lst)
    # 根据制动墩划分墩号 {第1联墩号:[0,2], 第2联墩号:[2,5], 第3联墩号:[5, 8]}
    max_index = len(substructure_namelst)
    boundary_lst = [0] + zdd_index_lst + [max_index - 1]
    pier_dict = {}
    for idx in range(len(boundary_lst) - 1):
        pier_dict[f'第{idx + 1}联墩号'] = [boundary_lst[idx], boundary_lst[idx + 1]]
    # print(pier_dict)
    # 根据制动墩获取每一联的长度 {第1联长度:15, 第2联长度:36, 第3联长度:27}
    length_dict = {}
    for idx in range(len(boundary_lst) - 1):
        start = boundary_lst[idx]
        end = boundary_lst[idx + 1]
        length_dict[f'第{idx + 1}联长度'] = sum(span_lst[start:end])
    # print(length_dict)
    # 汇总信息 {第1联:{墩号:[], 长度:[]}, ...}
    span_dict = {}
    for idx in range(len(boundary_lst) - 1):
        key = f'第{idx + 1}联'
        span_dict[key] = {
            '墩号': pier_dict[f'{key}墩号'],
            '长度': length_dict[f'{key}长度']
        }
    return span_dict

# ============================== 4.节点/单元/结构组mct处理 ============================== 
# 生成节点的mct命令流
def node_mctstring_handle(nodei, x, y, z):
    nodestring = '{}, {}, {}, {}'.format(nodei, x, y, z)
    return nodestring


# 生成梁单元的mct命令流
def beam_elem_mctstring_handle(elemi, elem_type, mat_id, sec_id, node1, node2, angle, sub = 0):
    elemstring = '{}, {}, {}, {}, {}, {}, {}, {}'.format(elemi, elem_type, mat_id, sec_id, node1, node2, angle, sub)
    return elemstring


# 生成板单元的mct命令流
def plate_elem_mctstring_handle(elemi, elem_type, mat_id, sec_id, node1, node2, node3, node4):
    # 4457, PLATE, 2, 1, 12, 23, 24, 13, 1, 0
    # 4460, PLATE, 2, 1, 34, 45, 46, 35, 1, 0
    elemstring = '{}, {}, {}, {}, {}, {}, {}, {}, 1, 0'.format(elemi, elem_type, mat_id, sec_id, node1, node2, node3, node4)
    return elemstring


# 生成结构组的mct命令流
def grup_mctstring_handle(group_name, node_str, elem_str):
    groupstring = '{}, {}, {}, 0'.format(group_name, node_str, elem_str)
    return groupstring


# 生成桩底固结mct命令流
def constraint_mctstring_handle(node_id, constraint="111101", name=""):
    """*CONSTRAINT 行: {node_id}, {constraint},{name}"""
    return f"{node_id}, {constraint}, {name}"


# 生成释放梁端约束的mct命令流
def frame_rls_mctstring_handle(elemid, _if, i, ifx, ify, ifz, imx, imy, imz, j, jfx, jfy, jfz, jmx, jmy, jmz, boundary_name):
    # '{}, NO, 000000, 0, 0, 0, 0, 0, 0 \n 000010, 0, 0, 0, 0, 0, 0, 释放梁端约束'
    framerlsstring = '{}, {}, {}, {}, {}, {}, {}, {}, {} \n {}, {}, {}, {}, {}, {}, {}, {}'.format(elemid, _if, i, ifx, ify, ifz, imx, imy, imz, j, jfx, jfy, jfz, jmx, jmy, jmz, boundary_name)
    return framerlsstring

def plate_rls_mctstring_handle(elemid, n1, n2, n3, n4, boundary_name):
    """生成板单元释放约束mct命令流 (*PLATE-RLS)

    Parameters
    ----------
    elemid : int
        板单元号
    n1~n4 : str
        各角点释放标志（5位，对应 Fx,Fy,Fz,Mx,My），如 "00000" "00001"
    boundary_name : str
        边界组名
    """
    return '{}, {}, {}, {}, {}, {}'.format(elemid, n1, n2, n3, n4, boundary_name)

# 生成弹性连接mct命令流
def elastic_link_mctstring_handle(link_id, n1, n2, type, lctype, sfx, sfy, sfz, smx, smy, smz, kx, ky, kz, krx, kry, krz, boundary, damp1, damp2, name):
    # {id},{n1},{n2},{type},{lctype},{sfx},...,{smz},{kx},...,{krz},{boundary},{damp1},{damp2},{name}
    elinkstring = '{}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}'.format(link_id, n1, n2, type, lctype, sfx, sfy, sfz, smx, smy, smz, kx, ky, kz, krx, kry, krz, boundary, damp1, damp2, name)
    return elinkstring

# 生成一般支承mct命令流
def support_mctstring_handle(support_id, dx, dy, dz, rx, ry, rz):
    supportstring = '{}, {}, {}, {}, {}, {},{}'.format(support_id, dx, dy, dz, rx, ry, rz)
    return supportstring


def _conn_stiff(conn_settings):
    """将 UI 连接设置 dict 转换为 build_elastic_links 所需的刚度参数。

    Parameters
    ----------
    conn_settings : dict
        ConnectionTab 缓存的单条连接设置，包含 type/SDx/SDy/SDz/SRx/SRy/SRz。

    Returns
    -------
    dict : {link_type, kx, ky, kz, krx, kry, krz}
    """
    if not conn_settings:
        return {}
    lt = conn_settings.get("type", "自定义")
    link_type = "GEN" if lt != "刚接" else "RIGID"
    return {
        "link_type": link_type,
        "kx":  str(conn_settings.get("SDx", "1e+07")),
        "ky":  str(conn_settings.get("SDy", "1e+05")),
        "kz":  str(conn_settings.get("SDz", "1e+05")),
        "krx": str(conn_settings.get("SRx", "1")),
        "kry": str(conn_settings.get("SRy", "1")),
        "krz": str(conn_settings.get("SRz", "1")),
    }


# 通用弹性连接：匹配(x,y)相同节点对，按z从高到低生成*ELASTICLINK行
def build_elastic_links(dict_a, dict_b, start_link_id=1,
                         link_type="GEN", lctype="0",
                         sfx="NO", sfy="NO", sfz="NO", smx="NO", smy="NO", smz="NO",
                         kx="1e+07", ky="1e+05", kz="1e+05",
                         krx="1", kry="1", krz="1",
                         boundary="NO", damp1="0.5", damp2="0.5",
                         name="弹性连接", xy_tol=6,
                         stiffness=None):
    """
    MCT行格式: {id},{n1},{n2},{type},{lctype},{sfx},...,{smz},{kx},...,{krz},{boundary},{damp1},{damp2},{name}

    stiffness : dict, optional — 由 _conn_stiff() 生成的刚度参数，覆盖 link_type/kx~krz 默认值。

    Returns
    -------
    elink_lines : list[str]
    next_link_id : int
    """
    if stiffness:
        link_type = stiffness.get("link_type", link_type)
        kx  = stiffness.get("kx", kx)
        ky  = stiffness.get("ky", ky)
        kz  = stiffness.get("kz", kz)
        krx = stiffness.get("krx", krx)
        kry = stiffness.get("kry", kry)
        krz = stiffness.get("krz", krz)
    # 构建 dict_b 的 (x,y) 查找表
    b_map = {}
    for nid, coord in dict_b.items():
        x, y, z = coord[0], coord[1], coord[2]
        key = (round(x, xy_tol), round(y, xy_tol))
        if key not in b_map or z > b_map[key][1]:
            b_map[key] = (nid, z)

    elink_lines = []
    lid = start_link_id

    for nid_a, coord_a in dict_a.items():
        xa, ya, za = coord_a[0], coord_a[1], coord_a[2]
        key = (round(xa, xy_tol), round(ya, xy_tol))
        if key not in b_map:
            continue
        nid_b, zb = b_map[key]
        n1, n2 = (nid_a, nid_b) if za >= zb else (nid_b, nid_a)
        line = elastic_link_mctstring_handle(
            lid, n1, n2, link_type, lctype,
            sfx, sfy, sfz, smx, smy, smz,
            kx, ky, kz, krx, kry, krz,
            boundary, damp1, damp2, name)
        elink_lines.append(line)
        lid += 1

    return elink_lines, lid

# ============================== 5.虚拟梁 ============================== 
# 生成虚拟梁单元
def make_virtual_elem(SECTION_name_search_dict, virtual_nodes_lst, elem_id):
    virtual_elems_mctlst = [] # 用于虚拟梁单元mct生成
    virtual_elems_lst = [] #  用于虚拟梁结构组生成
    virtual_section_id = SECTION_name_search_dict['虚拟梁'][0]
    for i in range(len(virtual_nodes_lst) - 1):
        node1 = virtual_nodes_lst[i]
        node2 = virtual_nodes_lst[i+1]
        mctstring = beam_elem_mctstring_handle(elem_id, 'BEAM', '2', virtual_section_id, node1, node2, '0')
        virtual_elems_mctlst.append(mctstring)
        virtual_elems_lst.append(elem_id)
        elem_id += 1
    return virtual_elems_lst, virtual_elems_mctlst, elem_id


# 生成虚拟梁结构组
def make_virtual_group(virtual_nodes_lst, virtual_elems_lst):
    virtual_group_mctlst = []
    node_str  = group_arithmetic_sequences(virtual_nodes_lst)
    elem_str  = group_arithmetic_sequences(virtual_elems_lst)
    mctstring = grup_mctstring_handle('虚拟梁', node_str, elem_str)
    virtual_group_mctlst.append(mctstring)
    return virtual_group_mctlst

# ============================== 6.桥面系 ============================== 
# 计算桥面参数
def calc_deck_geometry(trestle_ui_dict, span_dict):
    """计算桥面系几何参数，供纵肋/横肋节点生成共用。

    Returns
    -------
    rib_info : dict
        包含 is_double, h_rib_long, h_rib_trans, rib_trans_z, rib_beam_z,
        rib_x_lst, rib_start_y, rib_end_y, rib_trans_middle_y,
        beam_y_lst, rib_long_ylst, span_ranges
    """
    bridge_type = _get_bridge_type(trestle_ui_dict.get('components', {}))
    comp = trestle_ui_dict['components'][bridge_type]
    # 双层判断 + 各构件Z坐标
    deck_type = comp.get("deck_type", "单层钢面板")
    is_double = (deck_type == "双层钢面板")
    sections = trestle_ui_dict.get('section', {})
    rib_trans_sec_ref = comp.get('rib_trans_sec')
    h_rib_trans = _get_sec_H(sections, rib_trans_sec_ref, 140)
    h_rib_long = 0
    if is_double:
        rib_long_sec_ref = comp.get('rib_long_sec')
        h_rib_long = _get_sec_H(sections, rib_long_sec_ref, 140)
    # Z坐标体系（z=0=桥面顶）：
    # 单层横肋顶=0；双层横肋中心=-(h_long+h_trans)/2
    if is_double:
        rib_trans_z = -(h_rib_long + h_rib_trans) / 2
    else:
        rib_trans_z = 0
    rib_beam_z = rib_trans_z - (100 + h_rib_trans) / 2
    bridge_width_m = float(trestle_ui_dict.get('bridge_width', 6.0))
    bridge_width_mm = bridge_width_m * 1000
    deck_trans_ecc_mm = float(comp.get('deck_trans_ecc', 0) or 0)
    beam_space = comp.get('beam_space', '5@900')
    beam_first_space = float(comp.get('beam_first_space', 0))
    rib_trans_layout = comp.get('rib_trans_space', '750,68@750')
    # 横桥向定位：以概念中心线为基准
    rib_start_y = -deck_trans_ecc_mm      # 桥面左端距中心线
    rib_end_y = rib_start_y + bridge_width_mm  # 桥面右端
    rib_trans_middle_y = (rib_start_y + rib_end_y) / 2
    # 计算贝雷梁的y坐标（以概念中心线为基准）
    beam_y_lst = midasdisttolst2(-beam_first_space, beam_space)
    # 混合小肋中点坐标以及贝雷梁的y坐标后去重并排序
    beam_y_lst = list(sorted(set(beam_y_lst + [rib_trans_middle_y])))
    # 双层时：纵肋Y坐标（用于横肋节点补充，确保纵↔横弹连点齐全）
    rib_long_ylst = []
    if is_double:
        rib_long_layout = comp.get('rib_long_space', '200,14@400')
        rib_d2_mm, rib_long_space_str = _parse_a_bx(rib_long_layout, 200)
        rib_long_start_y = rib_start_y + rib_d2_mm
        rib_long_ylst = midasdisttolst2(rib_long_start_y, rib_long_space_str)
        rib_long_ylst = [y for y in rib_long_ylst if y < rib_end_y]

    # ---- 联数范围 ----
    span_ranges = {}
    current_x = 0
    for span_key, span_info in span_dict.items():
        span_length = int(span_info['长度'] * 1000)  # 米转换为毫米
        span_start = current_x
        span_end = current_x + span_length
        span_ranges[span_key] = [span_start, span_end]
        current_x = span_end
    for i, (key, value) in enumerate(span_ranges.items()):
        span_start, span_end = value
        span_ranges[key] = [span_start + i * 200, span_end + i * 200]

    # 纵梁伸出扩展：首联起点左移，尾联终点右移
    beam_cant_str = _get_component(trestle_ui_dict, 'beam_cant', '0,0')
    cant_parts = [x.strip() for x in str(beam_cant_str).split(",")]
    try: cant_left = int(float(cant_parts[0])) if cant_parts[0] else 0
    except (ValueError, TypeError): cant_left = 0
    try: cant_right = int(float(cant_parts[1])) if len(cant_parts) > 1 and cant_parts[1] else 0
    except (ValueError, TypeError): cant_right = 0
    span_keys = list(span_ranges.keys())
    if span_keys:
        span_ranges[span_keys[0]][0] -= cant_left
        span_ranges[span_keys[-1]][1] += cant_right

    # ---- 横肋X坐标：支持分联/不分联格式 ----
    # 分联格式: "(0,22@750)+(0,20@750)" 每联独立定义
    # 不分联格式: "0,22@750" 全局定义（向后兼容）
    if '(' in rib_trans_layout:
        groups = _parse_parenthesized_groups(rib_trans_layout)
        sorted_spans = sorted(span_ranges.items())
        rib_x_lst = []
        for gi, (span_key, (s_start, s_end)) in enumerate(sorted_spans):
            if gi < len(groups):
                offset_str = groups[gi]
            else:
                offset_str = groups[-1]  # 超出部分复用最后一组
            offset, space_str = _parse_a_bx(offset_str, 0)
            base = s_start + offset
            coords = midasdisttolst2(base, space_str)
            rib_x_lst.extend(coords)
    else:
        beam_d2_mm, rib_trans_space_str = _parse_a_bx(rib_trans_layout, 750)
        # 不分联格式：起点 = beam_d2_mm - cant_left（对齐贝雷梁端头）
        rib_x_lst = midasdisttolst2(beam_d2_mm - cant_left, rib_trans_space_str)

    return {
        "is_double": is_double,
        "h_rib_long": h_rib_long,
        "h_rib_trans": h_rib_trans,
        "rib_trans_z": rib_trans_z,
        "rib_beam_z": rib_beam_z,
        "rib_x_lst": rib_x_lst,
        "rib_start_y": rib_start_y,
        "rib_end_y": rib_end_y,
        "rib_trans_middle_y": rib_trans_middle_y,
        "beam_y_lst": beam_y_lst,
        "rib_long_ylst": rib_long_ylst,
        "span_ranges": span_ranges,
    }

# 创建纵肋节点
def make_rib_long_nodes(trestle_ui_dict, rib_info, node_id):
    """生成纵肋节点（双层桥面系）

    纵肋沿 X 方向（纵桥向）布置，位于横肋之上(z=0)。
    rib_long_space 格式 "a,b@x": a=原横肋悬臂d2(桥面左端偏移), b@x=纵肋横向间距
    """
    # 判断是否有纵肋
    if not rib_info.get("is_double"):
        return {}, {}, [], node_id

    rib_long_layout = _get_component(trestle_ui_dict, 'rib_long_space', '200,14@400')
    rib_d2_mm, rib_long_space_str = _parse_a_bx(rib_long_layout, 200)

    rib_x_lst = rib_info["rib_x_lst"]
    rib_start_y = rib_info["rib_start_y"]
    rib_end_y = rib_info["rib_end_y"]

    # 纵肋Y坐标：从桥面左端 + d2 起，按间距排布
    rib_long_start_y = rib_start_y + rib_d2_mm
    rib_long_ylst = midasdisttolst2(rib_long_start_y, rib_long_space_str)
    rib_long_ylst = [y for y in rib_long_ylst if y < rib_end_y]

    # 生成纵肋节点：每根纵肋 = 沿X方向的所有横肋位置
    rib_long_nodes_dict = {}
    rib_long_nodes_lst = []
    for j, ly in enumerate(rib_long_ylst):
        key = f"纵肋{j+1}"
        nodes = []
        for rx in rib_x_lst:
            nodes.append([node_id, [rx, ly, 0]])  # z=0（纵肋在横肋之上）
            rib_long_nodes_lst.append(node_id)
            node_id += 1
        rib_long_nodes_dict[key] = nodes

    return rib_long_nodes_dict, rib_long_nodes_lst, rib_long_ylst, node_id

def make_rib_long_nodes_mct(rib_long_nodes_dict):
    """纵肋节点 → MCT NODE 行"""
    mct_lst = []
    for nodes in rib_long_nodes_dict.values():
        for nid, coord in nodes:
            mct_lst.append(node_mctstring_handle(nid, coord[0], coord[1], coord[2]))
    return mct_lst

def make_rib_long_elems(trestle_ui_dict, rib_long_nodes_dict, elem_id, seg_boundaries=None):
    """生成纵肋BEAM单元（沿X方向相邻横肋位置之间）

    seg_boundaries: {纵肋名: [bx1, bx2, ...]} 分节边界X坐标(mm)
    """
    rib_long_sec_ref = _get_component(trestle_ui_dict, 'rib_long_sec')
    sec_id = _get_sec_id(trestle_ui_dict.get('section', {}), rib_long_sec_ref, 1)
    _rib_long_mat = _mat_id(trestle_ui_dict, _extract_brand(
        _get_component(trestle_ui_dict, 'rib_long_material')))
    elems_lst = []
    elems_mctlst = []

    for _, nodes in rib_long_nodes_dict.items():
        for i in range(len(nodes) - 1):
            n1, n2 = nodes[i][0], nodes[i+1][0]
            x1, x2 = nodes[i][1][0], nodes[i+1][1][0]

            # 物理断开：跳过 bx-2 → bx+2 的短梁单元
            if (x2 - x1) < 4.1:
                continue

            mct = beam_elem_mctstring_handle(elem_id, 'BEAM', _rib_long_mat, str(sec_id), n1, n2, '0')
            elems_mctlst.append(mct)
            elems_lst.append(elem_id)

            elem_id += 1

    return elems_lst, elems_mctlst, elem_id

def make_rib_long_group(rib_long_nodes_lst, rib_long_elems_lst):
    """纵肋结构组"""
    return _cmp_group('纵肋', rib_long_nodes_lst, rib_long_elems_lst)

def make_rib_long_trans_elink(rib_long_nodes_dict, span_rib_nodes_dict, elink_id, stiffness=None):
    """纵肋(z=0) ↔ 横肋(z=-h_rib_long) 弹性连接

    在每根纵肋与横肋的交叉点创建弹性连接。
    匹配 (x, y) 坐标（忽略 z 差异）。
    stiffness : dict, optional — 由 _conn_stiff() 生成的刚度参数。
    """
    if not rib_long_nodes_dict or not span_rib_nodes_dict:
        return [], elink_id

    _lt = stiffness.get("link_type", "GEN") if stiffness else "GEN"
    _kx = stiffness.get("kx", "1e+07") if stiffness else "1e+07"
    _ky = stiffness.get("ky", "1e+05") if stiffness else "1e+05"
    _kz = stiffness.get("kz", "1e+05") if stiffness else "1e+05"
    _krx = stiffness.get("krx", "1") if stiffness else "1"
    _kry = stiffness.get("kry", "1") if stiffness else "1"
    _krz = stiffness.get("krz", "1") if stiffness else "1"

    # 构建纵肋节点 (x,y)→nid 查找表
    rib_long_by_xy = {}
    for nodes in rib_long_nodes_dict.values():
        for nid, coord in nodes:
            x, y = coord[0], coord[1]
            rib_long_by_xy[(round(x, 1), round(y, 1))] = nid

    elink_lines = []
    for span_nodes in span_rib_nodes_dict.values():
        for trans_nodes in span_nodes.values():
            for trans_nid, trans_coord in trans_nodes:
                xy = (round(trans_coord[0], 1), round(trans_coord[1], 1))
                if xy in rib_long_by_xy:
                    rib_long_nid = rib_long_by_xy[xy]
                    # 纵肋 z=0 → 横肋 z=-h_rib_long
                    line = elastic_link_mctstring_handle(
                        elink_id, rib_long_nid, trans_nid,
                        _lt, '0', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO',
                        _kx, _ky, _kz, _krx, _kry, _krz,
                        'NO', '0.5', '0.5', '弹性连接')
                    elink_lines.append(line)
                    elink_id += 1

    return elink_lines, elink_id


# 分节建模辅助函数
def make_rib_long_seg_nodes(rib_long_nodes_dict, panel_length, node_id):
    """在纵肋分节处生成新节点（双层桥面板分节建模，物理断开）

    断开处删除原有节点，替换为 bx-2 和 bx+2 两个新节点，确保：
    1. 断开处不建立弹连（原有节点被删除，不会与横肋匹配）
    2. 纵肋单元跳过间隙（bx-2 → bx+2 间距 4mm < 4.1mm 阈值）

    Parameters
    ----------
    rib_long_nodes_dict : dict
        {"纵肋1": [[nid, [x, y, z]], ...], ...}
    panel_length : float
        节段长度（米）
    node_id : int
        当前最大节点号

    Returns
    -------
    seg_boundaries : dict
        {纵肋名: [boundary_x1, boundary_x2, ...]} 分节边界X坐标列表（mm）
    new_node_ids : list[int]
        新生成的节点号列表（需加入结构组）
    node_id : int
        更新后的节点号
    """
    if not rib_long_nodes_dict or not panel_length:
        return {}, [], node_id

    seg_boundaries = {}
    new_node_ids = []
    panel_length_mm = panel_length * 1000  # 米转毫米
    TOL = 1.0  # 1mm容差

    for rib_key, nodes in rib_long_nodes_dict.items():
        if len(nodes) < 2:
            seg_boundaries[rib_key] = []
            continue

        # 按X排序
        nodes_sorted = sorted(nodes, key=lambda n: n[1][0])
        x_coords = [n[1][0] for n in nodes_sorted]
        x_min, x_max = x_coords[0], x_coords[-1]

        # 计算边界X坐标
        boundaries = []
        x = x_min + panel_length_mm
        while x < x_max - TOL:
            boundaries.append(round(x, 1))
            x += panel_length_mm

        seg_boundaries[rib_key] = boundaries

        if not boundaries:
            continue

        # 纵肋Y/Z坐标（常量）
        rib_y = nodes_sorted[0][1][1]
        rib_z = nodes_sorted[0][1][2]

        # 构建边界集合，用于快速查找
        boundary_set = set(boundaries)

        # 处理每个节点：删除边界位置的原有节点，替换为 bx-2 和 bx+2
        new_nodes = []
        for nid, coord in nodes_sorted:
            nx = coord[0]
            # 检查是否在边界位置（容差 1mm）
            matched_bx = None
            for bx in boundary_set:
                if abs(nx - bx) < TOL:
                    matched_bx = bx
                    break
            if matched_bx is not None:
                # 删除原有节点，替换为 bx-2 和 bx+2
                new_nodes.append([node_id, [matched_bx - 2, rib_y, rib_z]])
                new_node_ids.append(node_id)
                node_id += 1
                new_nodes.append([node_id, [matched_bx + 2, rib_y, rib_z]])
                new_node_ids.append(node_id)
                node_id += 1
            else:
                new_nodes.append([nid, coord])

        # 更新纵肋节点列表
        rib_long_nodes_dict[rib_key] = sorted(new_nodes, key=lambda n: n[1][0])

    return seg_boundaries, new_node_ids, node_id

# 创建横肋节点
def make_rib_trans_nodes(rib_info, span_dict, node_id):
    '''创建横肋节点（几何参数由 calc_deck_geometry 预计算）

    双层桥面系：纵肋在上(z=0)，横肋下移(z=-h_rib_long)
    '''
    is_double = rib_info["is_double"]
    rib_trans_z = rib_info["rib_trans_z"]
    rib_x_lst = rib_info["rib_x_lst"]
    rib_start_y = rib_info["rib_start_y"]
    rib_end_y = rib_info["rib_end_y"]
    rib_trans_middle_y = rib_info["rib_trans_middle_y"]
    beam_y_lst = rib_info["beam_y_lst"]
    rib_long_ylst = rib_info["rib_long_ylst"]
    span_ranges = rib_info["span_ranges"]
    # 储存虚拟梁节点号 + 所有节点号
    virtual_nodes_lst = []
    rib_trans_nodes_lst = []
    # 汇总小肋的节点信息
    rib_nodes_dict = {}
    rib_elastic_nodes_dict = {}
    for i, rib_x in enumerate(rib_x_lst):
        rib_key = f"小肋{i+1}"
        elastic_key = f"小肋{i+1}"
        rib_nodes = []
        elastic_nodes = []
        # 起始点节点
        rib_nodes.append([node_id, [rib_x, rib_start_y, rib_trans_z]])
        rib_trans_nodes_lst.append(node_id)
        node_id += 1
        # 贝雷梁交点节点（仅在小肋范围内的贝雷梁）
        for beam_y in beam_y_lst:
            if rib_start_y < beam_y < rib_end_y:
                rib_nodes.append([node_id, [rib_x, beam_y, rib_trans_z]])
                elastic_nodes.append([node_id, [rib_x, beam_y, rib_trans_z]])  # 弹连节点与小肋节点相同
                rib_trans_nodes_lst.append(node_id)
                if abs(beam_y - rib_trans_middle_y) < 0.5:
                    virtual_nodes_lst.append(node_id)
                node_id += 1
        # 双层时：纵肋Y位置交点（为纵肋↔横肋弹连预留）
        if is_double:
            for rib_long_y in rib_long_ylst:
                if rib_start_y < rib_long_y < rib_end_y:
                    # 检查是否已存在（避免与beam_y重复产生同xy节点）
                    already = any(abs(rib_long_y - by) < 0.5 for by in beam_y_lst)
                    if not already:
                        rib_nodes.append([node_id, [rib_x, rib_long_y, rib_trans_z]])
                        elastic_nodes.append([node_id, [rib_x, rib_long_y, rib_trans_z]])
                        rib_trans_nodes_lst.append(node_id)
                        node_id += 1
        # 终止点节点
        rib_nodes.append([node_id, [rib_x, rib_end_y, rib_trans_z]])
        rib_trans_nodes_lst.append(node_id)
        node_id += 1
        # 按 Y 坐标排序，确保板单元节点顺序正确（纵肋交点可能插入在贝雷梁节点之间）
        rib_nodes.sort(key=lambda n: n[1][1])
        elastic_nodes.sort(key=lambda n: n[1][1])
        rib_nodes_dict[rib_key] = rib_nodes
        rib_elastic_nodes_dict[elastic_key] = elastic_nodes
    # 联数（span_ranges 已在前面计算）
    span_rib_nodes_dict = {}
    span_rib_elastic_nodes_dict = {}
    for key, value in span_ranges.items():
        span_rib_nodes_dict[key] = {}
        span_rib_elastic_nodes_dict[key] = {}
    sorted_span_keys = sorted(span_ranges.keys())
    for key, value in rib_nodes_dict.items():
        x = value[0][1][0] # 第n根小肋的x坐标
        assigned = False
        for i, (k, v) in enumerate(span_ranges.items()):
            if x >= v[0] and x < v[1]:
                span_rib_nodes_dict[k][key] = value
                assigned = True
                break
        if not assigned:
            # 落在联间隙中，分配给最近的联（优先后一联）
            best_k = sorted_span_keys[-1]
            for i, (k, v) in enumerate(span_ranges.items()):
                if x < v[0]:
                    best_k = k
                    break
            span_rib_nodes_dict[best_k][key] = value
    for key, value in rib_elastic_nodes_dict.items():
        x = value[0][1][0] # 第n根小肋的x坐标
        assigned = False
        for i, (k, v) in enumerate(span_ranges.items()):
            if x >= v[0] and x < v[1]:
                span_rib_elastic_nodes_dict[k][key] = value
                assigned = True
                break
        if not assigned:
            best_k = sorted_span_keys[-1]
            for i, (k, v) in enumerate(span_ranges.items()):
                if x < v[0]:
                    best_k = k
                    break
            span_rib_elastic_nodes_dict[best_k][key] = value
    return rib_trans_nodes_lst, span_rib_nodes_dict, span_rib_elastic_nodes_dict, virtual_nodes_lst, rib_trans_middle_y, node_id

# 生成横肋节点mct命令流所需的list
def make_rib_trans_nodes_mct(nodes_dict):
    rib_nodes_mctlst = []
    for sublst in nodes_dict.values():
        for lst in sublst:
            nodei = lst[0]
            x,y,z = lst[1]
            mctstring = node_mctstring_handle(nodei, x, y, z)
            rib_nodes_mctlst.append(mctstring)
    return rib_nodes_mctlst

# 创建横肋单元
def make_rib_trans_elems(trestle_ui_dict, span_rib_nodes_dict, elem_id):
    rib_sec_ref = _get_component(trestle_ui_dict, 'rib_trans_sec')
    rib_section_id = _get_sec_id(trestle_ui_dict.get('section', {}), rib_sec_ref, 1)
    _rib_trans_mat = _mat_id(trestle_ui_dict, _extract_brand(
        _get_component(trestle_ui_dict, 'rib_trans_material')))
    rib_elems_mctlst = [] # 用于小肋单元mct生成
    rib_elems_lst = [] # 用于小肋结构组生成
    for nodes_dict in span_rib_nodes_dict.values():
        for v in nodes_dict.values(): # v = '[[节点号, 坐标], [节点号, 坐标], ...]'
            for i in range(len(v) - 1):
                node1 = v[i][0]
                node2 = v[i+1][0]
                mctstring = beam_elem_mctstring_handle(elem_id, 'BEAM', _rib_trans_mat, rib_section_id, node1, node2, '0')
                rib_elems_mctlst.append(mctstring)
                rib_elems_lst.append(elem_id)
                elem_id += 1
    return rib_elems_lst, rib_elems_mctlst, elem_id

# 生成横肋结构组
def make_rib_trans_group(rib_trans_nodes_lst, rib_elems_lst):
    return _cmp_group('横肋', rib_trans_nodes_lst, rib_elems_lst)


# 生成桥面板结构组（节点来源：单层=横肋节点，双层=纵肋节点）
def make_plate_group(plate_nodes_lst, plate_elems_lst):
    return _cmp_group('桥面板', plate_nodes_lst, plate_elems_lst)

# 创建桥面板板单元（单层）
def make_plate_elems(span_rib_nodes_dict, elem_id, mat_id='2'):
    span_plate_elems_mctlst = []
    span_plate_elems_lst = []
    for span_key, rib_nodes_dict in span_rib_nodes_dict.items():
        rib_trans_nodes_lst = list(rib_nodes_dict.values())
        for i in range(len(rib_trans_nodes_lst)-1):
            rib1_node_lst = rib_trans_nodes_lst[i]
            rib2_node_lst = rib_trans_nodes_lst[i+1]
            for j in range(len(rib1_node_lst)-1):
                node1 = rib1_node_lst[j][0]
                node2 = rib2_node_lst[j][0]
                node3 = rib1_node_lst[j+1][0]
                node4 = rib2_node_lst[j+1][0]
                mctstring = plate_elem_mctstring_handle(elem_id, 'PLATE', mat_id, '1', node1, node2, node4, node3)
                span_plate_elems_mctlst.append(mctstring)
                span_plate_elems_lst.append(elem_id)
                elem_id += 1
    return span_plate_elems_lst, span_plate_elems_mctlst, elem_id

# 创建桥面板板单元（双层）
def make_plate_elems_by_long(rib_long_nodes_dict, elem_id, seg_boundaries=None, mat_id='2'):
    """双层桥面系：PLATE 单元连接相邻纵肋之间

    rib_long_nodes_dict: {"纵肋1": [[nid, [x,y,z]], ...], ...}
    每根纵肋节点按X排序，相邻纵肋同X节点间生成PLATE。
    节点顺序：N1(x小y小) → N2(x大y小) → N3(x大y大) → N4(x小y大)

    seg_boundaries: {纵肋名: [bx1, bx2, ...]} 分节边界X坐标(mm)
    """
    rib_long_keys = list(rib_long_nodes_dict.keys())
    if len(rib_long_keys) < 2:
        return [], [], [], elem_id

    span_plate_elems_mctlst = []
    span_plate_elems_lst = []
    plate_rls_mctlst = []  # PLATE-RLS 行（暂不启用）

    for i in range(len(rib_long_keys) - 1):
        nodes1 = rib_long_nodes_dict[rib_long_keys[i]]      # 纵肋i 的节点（y较小）
        nodes2 = rib_long_nodes_dict[rib_long_keys[i + 1]]  # 纵肋i+1 的节点（y较大）
        for j in range(len(nodes1) - 1):
            x_j  = nodes1[j][1][0]
            x_j1 = nodes1[j + 1][1][0]

            # 物理断开：跳过间隙内的板单元（宽度≈4mm的短板单元）
            if (x_j1 - x_j) < 4.1:
                continue

            n1 = nodes1[j][0]       # N1: x小, y小
            n2 = nodes1[j + 1][0]   # N2: x大, y小
            n3 = nodes2[j + 1][0]   # N3: x大, y大
            n4 = nodes2[j][0]       # N4: x小, y大
            mct = plate_elem_mctstring_handle(elem_id, 'PLATE', mat_id, '1', n1, n2, n3, n4)
            span_plate_elems_mctlst.append(mct)
            span_plate_elems_lst.append(elem_id)
            elem_id += 1

    return span_plate_elems_lst, span_plate_elems_mctlst, plate_rls_mctlst, elem_id


# ============================== 7.贝雷梁 ==============================
# 生成一片贝雷梁的点，参数param——0代表3m贝雷梁，1代表1.5m贝雷梁, i代表第i排贝雷梁, insert_P为插入点坐标, dict为词典
def bailey_piece_point(start_pt, param = '0', cfg=None):
    """按纵梁类型配置生成贝雷片节点（桁高/片几何取自 cfg，不再硬编码 321 尺寸）"""
    cfg = cfg or _truss_beam_cfg("321型贝雷梁")
    truss_height = cfg["truss_height"]
    piece = cfg["piece"]
    start_x, start_y, start_z = start_pt
    bailey_piece_dict = {
        '上弦杆接头':[], '上弦杆竖杆':[], '上弦杆斜杆':[],
        '下弦杆接头':[], '下弦杆竖杆':[], '下弦杆斜杆':[], '竖杆':[],
    }
    if param == "0":  # 3m贝雷片
        p = piece["3m"]
        for x in p["j0"]:
            bailey_piece_dict['上弦杆接头'].append([start_x + x, start_y, start_z])
            bailey_piece_dict['下弦杆接头'].append([start_x + x, start_y, start_z - truss_height])
        for x in p["j90"]:
            bailey_piece_dict['上弦杆竖杆'].append([start_x + x, start_y, start_z])
            bailey_piece_dict['下弦杆竖杆'].append([start_x + x, start_y, start_z - truss_height])
            bailey_piece_dict['竖杆'].append([start_x + x, start_y, start_z - truss_height / 2])
        for x in p["j795"]:
            bailey_piece_dict['上弦杆斜杆'].append([start_x + x, start_y, start_z])
            bailey_piece_dict['下弦杆斜杆'].append([start_x + x, start_y, start_z - truss_height])
    elif param == "1":  # 1.5m贝雷片
        p = piece["1.5m"]
        for x in p["j0"]:
            bailey_piece_dict['上弦杆接头'].append([start_x + x, start_y, start_z])
            bailey_piece_dict['下弦杆接头'].append([start_x + x, start_y, start_z - truss_height])
        for x in p["j90"]:
            bailey_piece_dict['上弦杆竖杆'].append([start_x + x, start_y, start_z])
            bailey_piece_dict['下弦杆竖杆'].append([start_x + x, start_y, start_z - truss_height])
            bailey_piece_dict['竖杆'].append([start_x + x, start_y, start_z - truss_height / 2])
        for x in p["j750"]:
            bailey_piece_dict['上弦杆斜杆'].append([start_x + x, start_y, start_z])
            bailey_piece_dict['下弦杆斜杆'].append([start_x + x, start_y, start_z - truss_height])
    return bailey_piece_dict


# 生成贝雷梁节点
def make_bailey_nodes(trestle_ui_dict, span_dict, span_rib_elastic_nodes_dict, fpl_elink_x, node_id, bailey_z=None, cfg=None, rib_info=None):
    cfg = cfg or _truss_beam_cfg("321型贝雷梁")
    if bailey_z is None:
        bailey_z = -70  # fallback（正常由调用方按 cfg 计算传入）
    bailey_space = _get_component(trestle_ui_dict, 'beam_space', '5@900')
    bailey_first_space = float(_get_component(trestle_ui_dict, 'beam_first_space', 0))
    bailey_ylst = midasdisttolst2(-bailey_first_space, bailey_space)
    # 坐标
    bailey_shangxiangan_coord_dict = {} # 上弦杆端头节点
    bailey_xiaxiangan_coord_dict = {} # 下弦杆端头节点
    bailey_shangshugan_coord_dict = {} # 上弦杆对应竖杆节点
    bailey_xiashugan_coord_dict = {} # 下弦杆对应竖杆节点
    bailey_shangxiegan_coord_dict = {} # 上弦杆对应斜杆节点
    bailey_xiaxiegan_coord_dict = {} # 下弦杆对应斜杆节点
    bailey_middleshugan_coord_dict = {} # 中间竖杆对应节点
    bailey_shangxiangancross_coord_dict = {} # 上弦杆和小肋的交点
    bailey_xiaxiangancross_coord_dict = {} # 下弦杆和分配梁的交点
    # 节点号
    bailey_shangxiangan_nodes_lst = []
    bailey_xiaxiangan_nodes_lst = []
    bailey_shugan_nodes_lst = []
    bailey_xiegan_nodes_lst = []
    # 联数：复用 rib_info["span_ranges"]（已含伸出扩展 + 制动墩偏移）
    if rib_info and "span_ranges" in rib_info:
        span_ranges = {k: list(v) for k, v in rib_info["span_ranges"].items()}
    else:
        span_ranges = {}
        current_x = 0
        for span_key, span_info in span_dict.items():
            span_length = int(span_info['长度'] * 1000)
            span_start = current_x
            span_end = current_x + span_length
            span_ranges[span_key] = [span_start, span_end]
            current_x = span_end
        for i, (key, value) in enumerate(span_ranges.items()):
            span_start, span_end = value
            span_ranges[key] = [span_start + i * 200, span_end + i * 200]
    span_fpl_elastic_xdict = {}
    for key, (start, end) in span_ranges.items():
        # 使用列表推导式筛选属于该区间的 x
        span_fpl_elastic_xdict[key] = [x for x in fpl_elink_x if start <= x <= end]
    # # 结构组
    # Bailey_XianGan_group_lst = [] # 弦杆
    # Bailey_ShuGan_group_lst = [] # 竖杆
    # Bailey_XieGan_group_lst = [] # 斜杆
    # 计算所有贝雷片的起点
    for key, value in span_ranges.items():
        bailey_shangxiangan_coord_dict[key] = {}
        bailey_xiaxiangan_coord_dict[key] = {}
        bailey_shangshugan_coord_dict[key] = {}
        bailey_xiashugan_coord_dict[key] = {}
        bailey_shangxiegan_coord_dict[key] = {}
        bailey_xiaxiegan_coord_dict[key] = {}
        bailey_middleshugan_coord_dict[key] = {}
        bailey_shangxiangancross_coord_dict[key] = {}
        bailey_xiaxiangancross_coord_dict[key] = {}
        span_bailey_xlst = list(range(value[0], value[1], 3000)) # 一联内贝雷片起点x坐标
        rib_cross_xlst = []
        for rib_cross_node_lst in span_rib_elastic_nodes_dict[key].values():
            rib_cross_xlst.extend([rib_cross_node_lst[0][1][0]]) # 一联内小肋的x坐标
        fpl_cross_xlst = span_fpl_elastic_xdict[key]
        for i in range(len(bailey_ylst)):
            # print(f'第{i+1}排贝雷')
            bailey_shangxiangan_coord_dict[key][f'第{i+1}排贝雷'] = []
            bailey_xiaxiangan_coord_dict[key][f'第{i+1}排贝雷'] = []
            bailey_shangshugan_coord_dict[key][f'第{i+1}排贝雷'] = []
            bailey_xiashugan_coord_dict[key][f'第{i+1}排贝雷'] = []
            bailey_shangxiegan_coord_dict[key][f'第{i+1}排贝雷'] = []
            bailey_xiaxiegan_coord_dict[key][f'第{i+1}排贝雷'] = []
            bailey_middleshugan_coord_dict[key][f'第{i+1}排贝雷'] = []
            bailey_shangxiangancross_coord_dict[key][f'第{i+1}排贝雷'] = []
            bailey_xiaxiangancross_coord_dict[key][f'第{i+1}排贝雷'] = []
            for x in span_bailey_xlst:
                start_pt = (x, bailey_ylst[i], bailey_z)
                bailey_piece_dict = bailey_piece_point(start_pt, cfg=cfg)
                # print(bailey_piece_dict) # 一片贝雷的默认节点
                bailey_shangxiangan_coord_dict[key][f'第{i+1}排贝雷'].extend(bailey_piece_dict['上弦杆接头'])
                bailey_xiaxiangan_coord_dict[key][f'第{i+1}排贝雷'].extend(bailey_piece_dict['下弦杆接头'])
                bailey_shangshugan_coord_dict[key][f'第{i+1}排贝雷'].extend(bailey_piece_dict['上弦杆竖杆'])
                bailey_xiashugan_coord_dict[key][f'第{i+1}排贝雷'].extend(bailey_piece_dict['下弦杆竖杆'])
                bailey_shangxiegan_coord_dict[key][f'第{i+1}排贝雷'].extend(bailey_piece_dict['上弦杆斜杆'])
                bailey_xiaxiegan_coord_dict[key][f'第{i+1}排贝雷'].extend(bailey_piece_dict['下弦杆斜杆'])
                bailey_middleshugan_coord_dict[key][f'第{i+1}排贝雷'].extend(bailey_piece_dict['竖杆'])
            for x in rib_cross_xlst:
                bailey_shangxiangancross_coord_dict[key][f'第{i+1}排贝雷'].append([x, bailey_ylst[i], bailey_z])
            for x in fpl_cross_xlst:
                bailey_xiaxiangancross_coord_dict[key][f'第{i+1}排贝雷'].append([x, bailey_ylst[i], bailey_z - 1400])

        
    # 节点
    bailey_shangxiangan_node_dict = {}
    bailey_xiaxiangan_node_dict = {}
    bailey_shugan_node_dict = {}
    # bailey_xiegan_node_dict = {}
    bailey_rib_elink_node_dict = {}
    bailey_fpl_elink_node_dict = {}

    for key, value in span_ranges.items():
        bailey_shangxiangan_node_dict[key] = {}
        bailey_xiaxiangan_node_dict[key] = {}
        bailey_shugan_node_dict[key] = {}
        # bailey_xiegan_node_dict[key] = {}
        bailey_rib_elink_node_dict[key] = {}
        bailey_fpl_elink_node_dict[key] = {}

    # 计算每一联 每一排贝雷 上下弦杆的所有节点
    for key, value in span_ranges.items():
        for i in range(len(bailey_ylst)):
            bailey_k = f'第{i+1}排贝雷'
            bailey_shangxiangan_node_dict[key][bailey_k] = []
            bailey_xiaxiangan_node_dict[key][bailey_k] = []
            bailey_shugan_node_dict[key][bailey_k] = []
            bailey_rib_elink_node_dict[key][bailey_k] = []
            bailey_fpl_elink_node_dict[key][bailey_k] = []
            # 上弦杆所有节点
            sum_lst = bailey_shangxiangan_coord_dict[key][bailey_k] + bailey_shangshugan_coord_dict[key][bailey_k] + bailey_shangxiegan_coord_dict[key][bailey_k] + bailey_shangxiangancross_coord_dict[key][bailey_k]
            sum_lst = list(sorted(set([tuple(x) for x in sum_lst]), key = lambda x:x[0]))
            # print(bailey_shangxiangancross_coord_dict[key][bailey_k])
            # print(sum_lst)
            for j in range(len(sum_lst)):
                bailey_shangxiangan_node_dict[key][bailey_k].append([node_id, sum_lst[j]])
                bailey_shangxiangan_nodes_lst.append(node_id)
                if list(sum_lst[j]) in bailey_shangxiangancross_coord_dict[key][bailey_k]:
                    bailey_rib_elink_node_dict[key][bailey_k].append([node_id, sum_lst[j]])
                node_id += 1
            # 下弦杆所有节点
            sum_lst = bailey_xiaxiangan_coord_dict[key][bailey_k] + bailey_xiashugan_coord_dict[key][bailey_k] + bailey_xiaxiegan_coord_dict[key][bailey_k] + bailey_xiaxiangancross_coord_dict[key][bailey_k]
            sum_lst = list(sorted(set([tuple(x) for x in sum_lst]), key = lambda x:x[0]))
            for j in range(len(sum_lst)):
                bailey_xiaxiangan_node_dict[key][bailey_k].append([node_id, sum_lst[j]])
                bailey_xiaxiangan_nodes_lst.append(node_id)
                if list(sum_lst[j]) in bailey_xiaxiangancross_coord_dict[key][bailey_k]:
                    bailey_fpl_elink_node_dict[key][bailey_k].append([node_id, sum_lst[j]])
                node_id += 1
            # 竖杆所有节点
            sum_lst = bailey_middleshugan_coord_dict[key][bailey_k]
            sum_lst = list(sorted(set([tuple(x) for x in sum_lst]), key = lambda x:x[0]))
            for j in range(len(sum_lst)):
                bailey_shugan_node_dict[key][bailey_k].append([node_id, sum_lst[j]])
                node_id += 1

    return bailey_shangxiangan_coord_dict, bailey_shangxiangan_nodes_lst, bailey_shangxiangan_node_dict, bailey_xiaxiangan_coord_dict, bailey_xiaxiangan_nodes_lst, bailey_xiaxiangan_node_dict, bailey_shugan_node_dict, bailey_shangxiegan_coord_dict, bailey_xiaxiegan_coord_dict, bailey_rib_elink_node_dict, bailey_fpl_elink_node_dict, node_id


# 生成贝雷梁上弦杆节点mct命令流所需的list
def make_bailey_shangxiangan_nodes_mct(nodes_dict):
    span_bailey_shangxiangan_nodes_mctlst = []
    for sublst in nodes_dict.values():
        for lst in sublst:
            nodei = lst[0]
            x,y,z = lst[1]
            mctstring = node_mctstring_handle(nodei, x, y, z)
            span_bailey_shangxiangan_nodes_mctlst.append(mctstring)
    return span_bailey_shangxiangan_nodes_mctlst


# 生成贝雷梁上弦杆节点mct命令流所需的list
def make_bailey_xiaxiangan_nodes_mct(nodes_dict):
    span_bailey_xiaxiangan_nodes_mctlst = []
    for sublst in nodes_dict.values():
        for lst in sublst:
            nodei = lst[0]
            x,y,z = lst[1]
            mctstring = node_mctstring_handle(nodei, x, y, z)
            span_bailey_xiaxiangan_nodes_mctlst.append(mctstring)
    return span_bailey_xiaxiangan_nodes_mctlst


# 生成贝雷梁上弦杆节点mct命令流所需的list
def make_bailey_shugan_nodes_mct(nodes_dict):
    span_bailey_shugan_nodes_mctlst = []
    for sublst in nodes_dict.values():
        for lst in sublst:
            nodei = lst[0]
            x,y,z = lst[1]
            mctstring = node_mctstring_handle(nodei, x, y, z)
            span_bailey_shugan_nodes_mctlst.append(mctstring)
    return span_bailey_shugan_nodes_mctlst


# 生成贝雷梁上弦杆单元
def make_bailey_shangxiangan_elem(SECTION_name_search_dict, bailey_shangxiangan_coord_dict, span_bailey_shangxiangan_nodes_lst, elem_id, cfg=None, bailey_mat_id='1'):
    cfg = cfg or _truss_beam_cfg("321型贝雷梁")
    bailey_shangxiangan_elems_mctlst = [] # 用于贝雷梁上弦杆单元mct生成
    bailey_shangxiangan_elems_lst = [] #  用于贝雷梁上弦杆结构组生成
    bailey_shangxiangan_frame_rls_elemlst = [] # 用于贝雷梁上弦杆释放梁端约束
    bailey_shangxiangan_section_id = SECTION_name_search_dict[cfg["chord_sec"]][0]

    for span_key, nodes_dict in span_bailey_shangxiangan_nodes_lst.items():
        shangxiangan_jietou_xlst = list(sorted(set([point[0] for point in  bailey_shangxiangan_coord_dict[span_key]['第1排贝雷']])))
        shangxiangan_jietou_frame_rls_xlst = shangxiangan_jietou_xlst[1:-1] # 去掉一联首尾端头坐标
        for v in nodes_dict.values(): # v = '[[节点号, 坐标], [节点号, 坐标], ...]'
            for i in range(len(v) - 1):
                node1 = v[i][0]
                node2 = v[i+1][0]
                node_coord1 = v[i][1]
                node_coord2 = v[i+1][1]
                mctstring = beam_elem_mctstring_handle(elem_id, 'BEAM', bailey_mat_id, bailey_shangxiangan_section_id, node1, node2, '0')
                bailey_shangxiangan_elems_mctlst.append(mctstring)
                bailey_shangxiangan_elems_lst.append(elem_id)
                # 判断梁端约束
                if node_coord2[0] in shangxiangan_jietou_frame_rls_xlst:
                    bailey_shangxiangan_frame_rls_elemlst.append(elem_id)
                elem_id += 1
    # print(bailey_shangxiangan_frame_rls__elemlst)
    return bailey_shangxiangan_elems_lst, bailey_shangxiangan_elems_mctlst, bailey_shangxiangan_frame_rls_elemlst, elem_id


# 生成贝雷梁下弦杆单元
def make_bailey_xiaxiangan_elem(SECTION_name_search_dict, bailey_xiaxiangan_coord_dict, bailey_xiaxiangan_nodes_lst, elem_id, cfg=None, bailey_mat_id='1'):
    cfg = cfg or _truss_beam_cfg("321型贝雷梁")
    bailey_xiaxiangan_elems_mctlst = [] # 用于虚拟梁单元mct生成
    bailey_xiaxiangan_elems_lst = [] #  用于虚拟梁结构组生成
    bailey_xiaxiangan_frame_rls_elemlst = [] # 用于贝雷梁下弦杆释放梁端约束
    bailey_xiaxiangan_section_id = SECTION_name_search_dict[cfg["chord_sec"]][0]

    for span_key, nodes_dict in bailey_xiaxiangan_nodes_lst.items():
        xiaxiangan_jietou_xlst = list(sorted(set([point[0] for point in  bailey_xiaxiangan_coord_dict[span_key]['第1排贝雷']])))
        xiaxiangan_jietou_frame_rls_xlst = xiaxiangan_jietou_xlst[1:-1] # 去掉一联首尾端头坐标
        for v in nodes_dict.values(): # v = '[[节点号, 坐标], [节点号, 坐标], ...]'
            for i in range(len(v) - 1):
                node1 = v[i][0]
                node2 = v[i+1][0]
                node_coord1 = v[i][1]
                node_coord2 = v[i+1][1]
                mctstring = beam_elem_mctstring_handle(elem_id, 'BEAM', bailey_mat_id, bailey_xiaxiangan_section_id, node1, node2, '0')
                bailey_xiaxiangan_elems_mctlst.append(mctstring)
                bailey_xiaxiangan_elems_lst.append(elem_id)
                # 判断梁端约束
                if node_coord2[0] in xiaxiangan_jietou_frame_rls_xlst:
                    bailey_xiaxiangan_frame_rls_elemlst.append(elem_id)
                elem_id += 1
    return bailey_xiaxiangan_elems_lst, bailey_xiaxiangan_elems_mctlst, bailey_xiaxiangan_frame_rls_elemlst, elem_id


# 生成贝雷梁竖杆单元
def make_bailey_shugan_elem(SECTION_name_search_dict, span_bailey_shangxiangan_node_dict, span_bailey_xiaxiangan_node_dict, span_bailey_shugan_node_dict, elem_id, cfg=None, bailey_mat_id='1'):
    cfg = cfg or _truss_beam_cfg("321型贝雷梁")
    bailey_shugan_elems_mctlst = [] # 用于虚拟梁单元mct生成
    bailey_shugan_nodes_lst = []
    bailey_shugan_elems_lst = [] #  用于虚拟梁结构组生成
    bailey_shugan_section_id = SECTION_name_search_dict[cfg["web_sec"]][0]

    for key, value in span_bailey_shugan_node_dict.items():
        for k, v in value.items():
            # 第i联内的第j排贝雷梁
            shugan_x_nodei_dict = {}
            shangxiangan_x_nodei_dict = {}
            xiaxiangan_x_nodei_dict = {}
            for nodei, (x, y, z) in v:
                shugan_x_nodei_dict[x] = nodei
            for nodei, (x, y, z) in span_bailey_shangxiangan_node_dict[key][k]:
                shangxiangan_x_nodei_dict[x] = nodei
            for nodei, (x, y, z) in span_bailey_xiaxiangan_node_dict[key][k]:
                xiaxiangan_x_nodei_dict[x] = nodei
            # 一排内所有x相同的坐标对应的
            for coordx, nodeid in shugan_x_nodei_dict.items():
                shangxiangan_nodeid = shangxiangan_x_nodei_dict[coordx]
                xiaxiangan_nodeid = xiaxiangan_x_nodei_dict[coordx]
                shugan_nodeid_lst = [shangxiangan_nodeid, nodeid, xiaxiangan_nodeid]
                bailey_shugan_nodes_lst.extend(shugan_nodeid_lst)
                mctstring = beam_elem_mctstring_handle(elem_id, 'BEAM', bailey_mat_id, bailey_shugan_section_id, shangxiangan_nodeid, nodeid, '90')
                bailey_shugan_elems_mctlst.append(mctstring)
                bailey_shugan_elems_lst.append(elem_id)
                elem_id += 1
                mctstring = beam_elem_mctstring_handle(elem_id, 'BEAM', bailey_mat_id, bailey_shugan_section_id, nodeid, xiaxiangan_nodeid, '90')
                bailey_shugan_elems_mctlst.append(mctstring)
                bailey_shugan_elems_lst.append(elem_id)
                elem_id += 1

    return bailey_shugan_nodes_lst, bailey_shugan_elems_lst, bailey_shugan_elems_mctlst, elem_id


# 生成贝雷梁斜杆单元
def make_bailey_xiegan_elem(SECTION_name_search_dict, span_bailey_shangxiangan_node_dict, span_bailey_xiaxiangan_node_dict, span_bailey_shugan_node_dict, bailey_shangxiegan_coord_dict, bailey_xiaxiegan_coord_dict, elem_id, cfg=None, bailey_mat_id='1'):
    cfg = cfg or _truss_beam_cfg("321型贝雷梁")
    bailey_xiegan_elems_mctlst = [] # 用于虚拟梁单元mct生成
    bailey_xiegan_nodes_lst = []
    bailey_xiegan_elems_lst = [] #  用于虚拟梁结构组生成
    bailey_xiegan_section_id = SECTION_name_search_dict[cfg["web_sec"]][0]

    bailey_shangxiegan_x_dict = {}
    bailey_xiaxiegan_x_dict = {}

    for key, value in span_bailey_shugan_node_dict.items():
        # 斜杆在上弦杆对应的x坐标
        bailey_shangxiegan_x_dict[key] = [pt[0] for pt in bailey_shangxiegan_coord_dict[key]['第1排贝雷']]
        # 斜杆在下弦杆对应的x坐标
        bailey_xiaxiegan_x_dict[key] = [pt[0] for pt in bailey_xiaxiegan_coord_dict[key]['第1排贝雷']]
        for k, v in value.items():
            # 第i联内的第j排贝雷梁
            shugan_x_nodei_dict = {}
            shangxiangan_x_nodei_dict = {}
            xiaxiangan_x_nodei_dict = {}
            for nodei, (x, y, z) in v:
                shugan_x_nodei_dict[x] = nodei
            for nodei, (x, y, z) in span_bailey_shangxiangan_node_dict[key][k]:
                shangxiangan_x_nodei_dict[x] = nodei
            for nodei, (x, y, z) in span_bailey_xiaxiangan_node_dict[key][k]:
                xiaxiangan_x_nodei_dict[x] = nodei
            shangxiangan_nodeid_lst = []
            xiaxiangan_nodeid_lst = []
            shangshugan_nodeid_lst = []
            xiashugan_nodeid_lst = []
            # 斜杆在上弦杆的节点号以及上半斜杆在竖杆的节点号
            for coordx in bailey_shangxiegan_x_dict[key]:
                shangxiangan_nodeid = shangxiangan_x_nodei_dict[coordx]
                shangxiangan_nodeid_lst.append(shangxiangan_nodeid)
                shugan_nodeid = shugan_x_nodei_dict[coordx-705]
                shangshugan_nodeid_lst.append(shugan_nodeid)
                shugan_nodeid = shugan_x_nodei_dict[coordx+705]
                shangshugan_nodeid_lst.append(shugan_nodeid)
            # 斜杆在下弦杆的节点号以及下半斜杆在竖杆的节点号
            for coordx in bailey_xiaxiegan_x_dict[key]:
                xiaxiangan_nodeid = xiaxiangan_x_nodei_dict[coordx]
                xiaxiangan_nodeid_lst.append(xiaxiangan_nodeid)
                shugan_nodeid = shugan_x_nodei_dict[coordx-705]
                xiashugan_nodeid_lst.append(shugan_nodeid)
                shugan_nodeid = shugan_x_nodei_dict[coordx+705]
                xiashugan_nodeid_lst.append(shugan_nodeid)
            # print(key)
            # print(k)
            # print(shangxiangan_nodeid_lst)
            # print(xiaxiangan_nodeid_lst)
            # print(shangshugan_nodeid_lst)
            # print(xiashugan_nodeid_lst)
            
            bailey_xiegan_nodes_lst.extend(shangxiangan_nodeid_lst)
            bailey_xiegan_nodes_lst.extend(xiaxiangan_nodeid_lst)
            bailey_xiegan_nodes_lst.extend(shangshugan_nodeid_lst)
            
            for i in range(len(shangxiangan_nodeid_lst)):
                node1 = shangxiangan_nodeid_lst[i]
                node2 = shangshugan_nodeid_lst[2*i]
                mctstring = beam_elem_mctstring_handle(elem_id, 'BEAM', bailey_mat_id, bailey_xiegan_section_id, node1, node2, '90')
                bailey_xiegan_elems_mctlst.append(mctstring)
                bailey_xiegan_elems_lst.append(elem_id)
                elem_id += 1
                node1 = shangxiangan_nodeid_lst[i]
                node2 = shangshugan_nodeid_lst[2*i+1]
                mctstring = beam_elem_mctstring_handle(elem_id, 'BEAM', bailey_mat_id, bailey_xiegan_section_id, node1, node2, '90')
                bailey_xiegan_elems_mctlst.append(mctstring)
                bailey_xiegan_elems_lst.append(elem_id)
                elem_id += 1
            for i in range(len(xiaxiangan_nodeid_lst)):
                node1 = xiaxiangan_nodeid_lst[i]
                node2 = xiashugan_nodeid_lst[2*i]
                mctstring = beam_elem_mctstring_handle(elem_id, 'BEAM', bailey_mat_id, bailey_xiegan_section_id, node1, node2, '90')
                bailey_xiegan_elems_mctlst.append(mctstring)
                bailey_xiegan_elems_lst.append(elem_id)
                elem_id += 1
                node1 = xiaxiangan_nodeid_lst[i]
                node2 = xiashugan_nodeid_lst[2*i+1]
                mctstring = beam_elem_mctstring_handle(elem_id, 'BEAM', bailey_mat_id, bailey_xiegan_section_id, node1, node2, '90')
                bailey_xiegan_elems_mctlst.append(mctstring)
                bailey_xiegan_elems_lst.append(elem_id)
                elem_id += 1

    return bailey_xiegan_nodes_lst, bailey_xiegan_elems_lst, bailey_xiegan_elems_mctlst, elem_id


# 生成贝雷梁上弦杆结构组
def _cmp_group(name, nodes, elems):
    """生成压缩后的结构组 MCT 行（排序去重避免行超长）"""
    ns = group_arithmetic_sequences(sorted(set(nodes))) if nodes else ""
    es = group_arithmetic_sequences(elems) if elems else ""
    return [grup_mctstring_handle(name, ns, es)]


def make_bailey_shangxiangan_group(bailey_shangxiangan_nodes_lst, bailey_shangxiangan_elems_lst):
    return _cmp_group('贝雷梁上弦杆', bailey_shangxiangan_nodes_lst, bailey_shangxiangan_elems_lst)


def make_bailey_xiaxiangan_group(bailey_xiaxiangan_nodes_lst, bailey_xiaxiangan_elems_lst):
    return _cmp_group('贝雷梁下弦杆', bailey_xiaxiangan_nodes_lst, bailey_xiaxiangan_elems_lst)

def make_bailey_xiangan_group(bailey_shangxiangan_nodes_lst, bailey_shangxiangan_elems_lst, bailey_xiaxiangan_nodes_lst, bailey_xiaxiangan_elems_lst):
    """生成贝雷梁弦杆结构组（上弦杆 + 下弦杆全部节点和单元）"""
    all_nodes = list(bailey_shangxiangan_nodes_lst) + list(bailey_xiaxiangan_nodes_lst)
    all_elems = list(bailey_shangxiangan_elems_lst) + list(bailey_xiaxiangan_elems_lst)
    return _cmp_group('贝雷梁弦杆', all_nodes, all_elems)

# 生成贝雷梁竖杆结构组
def make_bailey_shugan_group(bailey_shugan_nodes_lst, bailey_shugan_elems_lst):
    return _cmp_group('贝雷梁竖杆', bailey_shugan_nodes_lst, bailey_shugan_elems_lst)


# 生成贝雷梁斜杆结构组
def make_bailey_xiegan_group(bailey_xiegan_nodes_lst, bailey_xiegan_elems_lst):
    return _cmp_group('贝雷梁斜杆', bailey_xiegan_nodes_lst, bailey_xiegan_elems_lst)

# 生成贝雷梁结构组
def make_bailey_group(beam_type, bailey_shangxiangan_nodes_lst, bailey_shangxiangan_elems_lst, bailey_xiaxiangan_nodes_lst, bailey_xiaxiangan_elems_lst, bailey_shugan_nodes_lst, bailey_shugan_elems_lst, bailey_xiegan_nodes_lst, bailey_xiegan_elems_lst):
    all_nodes = list(bailey_shangxiangan_nodes_lst) + list(bailey_xiaxiangan_nodes_lst) + list(bailey_shugan_nodes_lst) + list(bailey_xiegan_nodes_lst)
    all_elems = list(bailey_shangxiangan_elems_lst) + list(bailey_xiaxiangan_elems_lst) + list(bailey_shugan_elems_lst) + list(bailey_xiegan_elems_lst)
    return _cmp_group(beam_type, all_nodes, all_elems)

# 生成贝雷释放梁端约束
def make_bailey_frame_rls(bailey_shangxiangan_frame_rls_elemlst, bailey_xiaxiangan_frame_rls_elemlst):
    bailey_frame_rls_mctlst = []
    bailey_frame_rls_elemlst = sorted(bailey_shangxiangan_frame_rls_elemlst + bailey_xiaxiangan_frame_rls_elemlst)
    for elemid in bailey_frame_rls_elemlst:
        mctstring = frame_rls_mctstring_handle(elemid, 'NO', '000000', '0', '0', '0', '0', '0', '0', '000010', '0', '0', '0', '0', '0', '0', '释放梁端约束')
        bailey_frame_rls_mctlst.append(mctstring)
    return bailey_frame_rls_mctlst


# 创建纵梁与小肋弹性连接
def make_rib_trans_beam_elink(span_rib_elastic_nodes_dict, span_bailey_rib_elink_node_dict, elink_id, stiffness=None):
    """stiffness : dict, optional — 由 _conn_stiff() 生成的刚度参数。"""
    _lt = stiffness.get("link_type", "GEN") if stiffness else "GEN"
    _kx = stiffness.get("kx", "1e+07") if stiffness else "1e+07"
    _ky = stiffness.get("ky", "1e+05") if stiffness else "1e+05"
    _kz = stiffness.get("kz", "1e+05") if stiffness else "1e+05"
    _krx = stiffness.get("krx", "1") if stiffness else "1"
    _kry = stiffness.get("kry", "1") if stiffness else "1"
    _krz = stiffness.get("krz", "1") if stiffness else "1"

    rib_bailey_elink_mctlst = []
    all_intersections = []
    for span_key, rib_elink_node_dict in span_rib_elastic_nodes_dict.items():
        for ribi_elink_node_lst in rib_elink_node_dict.values():
            bailey_elink_node_dict = span_bailey_rib_elink_node_dict[span_key]
            bailey_elinl_node_lst = list(bailey_elink_node_dict.values())
            for bei_lei in bailey_elinl_node_lst:
                for xl_node, xl_coord in ribi_elink_node_lst:
                    for bl_node, bl_coord in bei_lei:
                        # 匹配x和y坐标
                        if xl_coord[0] == bl_coord[0] and xl_coord[1] == bl_coord[1]:
                            intersection = [[xl_node, tuple(xl_coord)], [bl_node, bl_coord]]
                            all_intersections.append(intersection)
    for lst in all_intersections:
        node1 = lst[0][0] # 弹连第一个节点
        node2 = lst[1][0] # 弹连第二个节点
        mctstring = elastic_link_mctstring_handle(elink_id, node1, node2,
            _lt, '0', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO',
            _kx, _ky, _kz, _krx, _kry, _krz,
            'NO', '0.5', '0.5', '弹性连接')
        rib_bailey_elink_mctlst.append(mctstring)
        elink_id += 1
    return rib_bailey_elink_mctlst, elink_id


# ============================== 7b.型钢纵梁 ==============================
def make_steel_beam_nodes(rib_info, h_steel, fpl_elink_x, node_id):
    """
    生成型钢纵梁节点：简单直线梁，X方向通长。

    Z = rib_trans_z - (h_rib_trans + h_steel) / 2
    Y = 复用 beam_y_lst
    X = 每联起终点 + 分配梁弹连点位置

    Returns
    -------
    steel_node_dict : {span_key: {beam_name: [[node_id, (x,y,z)], ...]}}
    steel_nodes_lst : [node_id, ...]
    steel_rib_elink_node_dict : {span_key: {beam_name: [[node_id, (x,y,z)], ...]}}  — 与横肋弹连用
    steel_fpl_elink_node_dict : {span_key: {beam_name: [[node_id, (x,y,z)], ...]}}  — 与分配梁弹连用
    node_id : int
    """
    rib_trans_z = rib_info["rib_trans_z"]
    h_rib_trans = rib_info["h_rib_trans"]
    beam_y_lst = rib_info["beam_y_lst"]
    span_ranges = rib_info["span_ranges"]
    rib_x_lst = rib_info["rib_x_lst"]

    steel_z = rib_trans_z - (h_rib_trans + h_steel) / 2

    steel_node_dict = {}
    steel_nodes_lst = []
    steel_rib_elink_node_dict = {}
    steel_fpl_elink_node_dict = {}

    for span_key, (span_start, span_end) in span_ranges.items():
        steel_node_dict[span_key] = {}
        steel_rib_elink_node_dict[span_key] = {}
        steel_fpl_elink_node_dict[span_key] = {}

        # 收集该联内的X坐标：联起终点 + 横肋位置 + 分配梁弹连位置
        xs = [span_start, span_end]
        for rx in rib_x_lst:
            if span_start <= rx <= span_end:
                xs.append(rx)
        for fx in fpl_elink_x:
            if span_start <= fx <= span_end:
                xs.append(fx)
        xs = sorted(set(xs))

        # 横肋X集合（用于判断弹连点）
        rib_xs_set = set(int(x) for x in rib_x_lst if span_start <= x <= span_end)
        # 分配梁弹连X集合
        fpl_xs_set = set(int(x) for x in fpl_elink_x if span_start <= x <= span_end)

        for i, y in enumerate(beam_y_lst):
            beam_name = f'第{i+1}排型钢'
            nodes = []
            rib_elink_nodes = []
            fpl_elink_nodes = []
            for x in xs:
                nodes.append([node_id, (x, y, steel_z)])
                steel_nodes_lst.append(node_id)
                if int(x) in rib_xs_set:
                    rib_elink_nodes.append([node_id, (x, y, steel_z)])
                if int(x) in fpl_xs_set:
                    fpl_elink_nodes.append([node_id, (x, y, steel_z)])
                node_id += 1
            steel_node_dict[span_key][beam_name] = nodes
            steel_rib_elink_node_dict[span_key][beam_name] = rib_elink_nodes
            steel_fpl_elink_node_dict[span_key][beam_name] = fpl_elink_nodes

    return steel_node_dict, steel_nodes_lst, steel_rib_elink_node_dict, steel_fpl_elink_node_dict, node_id


def make_steel_beam_nodes_mct(steel_node_dict):
    """型钢纵梁节点 → MCT NODE 行"""
    mct_lst = []
    for span_dict in steel_node_dict.values():
        for beam_nodes in span_dict.values():
            for nid, (x, y, z) in beam_nodes:
                mct_lst.append(node_mctstring_handle(nid, x, y, z))
    return mct_lst


def make_steel_beam_elems(steel_node_dict, section_name_search_dict, steel_section_name, elem_id, mat_id='1'):
    """
    生成型钢梁单元

    Parameters
    ----------
    mat_id : str — 材质ID（默认 '1'）

    Returns
    -------
    steel_elems_lst : [elem_id, ...]
    steel_elems_mctlst : [mctstring, ...]
    elem_id : int
    """
    steel_elems_lst = []
    steel_elems_mctlst = []

    # 查找截面ID：SECTION_name_search_dict 结构为 {截面名称: [截面编号, ...]}，value[0] 才是编号
    _target_norm = _norm_sec_name(steel_section_name).upper()
    steel_sec_id = None
    for sec_name, sec_value in section_name_search_dict.items():
        if _norm_sec_name(sec_name).upper() == _target_norm:
            steel_sec_id = sec_value[0] if sec_value else None
            break
    if steel_sec_id is None:
        # fallback: 使用第一个截面的编号（不能取键名）
        _first = next(iter(section_name_search_dict.values()), None)
        steel_sec_id = _first[0] if _first else 1

    for span_dict in steel_node_dict.values():
        for beam_nodes in span_dict.values():
            for j in range(len(beam_nodes) - 1):
                ni = beam_nodes[j][0]
                nj = beam_nodes[j + 1][0]
                mctstring = beam_elem_mctstring_handle(elem_id, 'BEAM', mat_id, steel_sec_id, ni, nj, '0')
                steel_elems_mctlst.append(mctstring)
                steel_elems_lst.append(elem_id)
                elem_id += 1

    return steel_elems_lst, steel_elems_mctlst, elem_id


def make_steel_beam_group(steel_nodes_lst, steel_elems_lst):
    """型钢纵梁结构组 MCT 行"""
    ns = group_arithmetic_sequences(steel_nodes_lst) if steel_nodes_lst else ""
    es = group_arithmetic_sequences(steel_elems_lst) if steel_elems_lst else ""
    return [grup_mctstring_handle('型钢纵梁', ns, es)]


# ============================== 8.分配梁 ==============================
# ---- 分配梁节点 → MCT 命令流 ----
def make_dist_nodes_mct(dist_trans_dict, dist_long_dict):
    """展平 dist_trans_dict + dist_long_dict 中的所有节点为 MCT NODE 行"""
    mct_lst = []
    for nodes in list(dist_trans_dict.values()) + list(dist_long_dict.values()):
        for nid, coord in nodes:
            mct_lst.append(node_mctstring_handle(nid, coord[0], coord[1], coord[2]))
    return mct_lst


# 确定分配梁与贝雷梁下弦杆弹连点、贝雷梁下弦杆固结点x坐标
def beam_bottom_elink_x(span_dict, substructure_dict, span_lst=None):
    """
    确定分配梁与贝雷梁下弦杆弹连点、贝雷梁下弦杆固结点x坐标

    前提：建模以第一根小肋x坐标为x原点0。
    桥台处取贝雷梁最外侧竖杆x坐标为固结点(起始侧90, 终点侧total_len-90)，
    单排桩处取桩位中心线x为弹连点，
    制动墩处前一联弹连点x-90, 后一联弹连点x+200+90。
    """
    # 按数字键排序构建0-indexed支撑类型列表
    sorted_keys = sorted(substructure_dict.keys(), key=int)
    types = [substructure_dict[k]['type'] for k in sorted_keys]
    n = len(types)

    # 从span_dict提取联边界位置
    lian_keys = sorted(span_dict.keys())
    bounds = []
    for k in lian_keys:
        r = span_dict[k]['墩号']
        if not bounds:
            bounds.append(r[0])
        bounds.append(r[1])
    # bounds = [0, a, b, ..., last]

    # span_lst回退：从span_dict均匀反推（无明细时近似）
    if span_lst is None:
        span_lst = []
        for i, k in enumerate(lian_keys):
            start, end = bounds[i], bounds[i + 1]
            n_spans = end - start
            lian_len = span_dict[k]['长度']
            span_lst.extend([lian_len / n_spans] * n_spans)

    # 累计跨径到每个支撑位置(m)
    cum_spans = [0]
    for s in span_lst:
        cum_spans.append(cum_spans[-1] + s)

    elink = []       # 单排桩弹连点
    brake_elink = [] # 制动墩弹连点

    for i in range(n):
        cum_m = cum_spans[i]                           # 到第i个支撑的累计跨径(m)
        brakes_before = sum(1 for bp in bounds[1:-1] if bp < i)  # 前制动墩数
        x = cum_m * 1000 + brakes_before * 200          # 桩位中心线x(mm)

        t = types[i]
        if t in ("桥台", "重力式桥台", "简易桥台"):
            # 桥台：固结点由调用者通过总长计算(起始侧90, 终点侧total_len-90)
            pass
        elif t == "单排桩":
            elink.append(x)
        elif t == "制动墩":
            # 前一联弹连点 x-90, 后一联弹连点 x+200+90
            brake_elink.append(x - 90)
            brake_elink.append(x + 200 + 90)

    return elink, brake_elink


def _calc_pile_ys_bx(y_start, pile_cantilever, pile_trans_space_str):
    """
    计算桩y坐标：从 y_start + pile_cantilever 起，按 spacing 累加。
    用于新UI：pile_trans_space 已通过 _parse_a_bx 解析，传入的pile_trans_space_str不含前缀a。
    """
    spaces = midasdisttolst(pile_trans_space_str)
    if not spaces:
        return []
    ys = [y_start + pile_cantilever]
    for d in spaces:
        ys.append(ys[-1] + d)
    return ys


def _calc_pile_xs_bx(x_start, pile_cantilever, pile_long_space_str):
    """
    计算制动墩多排桩x坐标：从 x_start + pile_cantilever 起，按 spacing 累加。
    """
    spaces = midasdisttolst(pile_long_space_str)
    if not spaces:
        return []
    xs = [x_start + pile_cantilever]
    for d in spaces:
        xs.append(xs[-1] + d)
    return xs


def make_dist_nodes(trestle_ui_dict, elink_x, brake_elink_x,
                     beam_dist_elink_nodes, node_id, beam_half_height=50):
    """
    生成分配梁节点（横向dist_trans + 纵向dist_long）。

    - z_beam_bottom 和 all_beam_ys 由 beam_dist_elink_nodes 提供。
      beam_dist_elink_nodes 的 z 约定为"纵梁截面中心 z"（贝雷梁下弦杆节点、型钢纵梁节点均为中心），
      beam_half_height 为从该中心到梁底面的偏移（贝雷梁下弦杆=50，型钢=h_steel/2）。
      横向分配梁中心 z = 梁底面 z - H_trans/2。
    各支撑类型生成的构件：
    - 桥台：pass
    - 单排桩：1根横向分配梁 (dist_trans)
    - 制动墩：2根横向分配梁 (dist_trans) + N根纵向分配梁 (dist_long)
    """
    # ========== 从 beam_dist_elink_nodes 提取 z 和 y ==========
    if not beam_dist_elink_nodes:
        raise ValueError("beam_dist_elink_nodes 为空，无法获取贝雷梁下弦杆与分配梁弹连点坐标")
    # 从纵梁截面中心z到梁底面的偏移（贝雷梁下弦杆半高50；型钢由调用方传 h_steel/2）
    # z_beam_bottom 名称保留（历史），实为纵梁截面中心z
    z_beam_bottom = next(iter(beam_dist_elink_nodes.values()))[2]
    # 不重复y值列表
    all_beam_ys = sorted(set(coord[1] for coord in beam_dist_elink_nodes.values()))

    # ========== 提取数据 ==========
    sections_dict = trestle_ui_dict['section']
    sub = trestle_ui_dict['substructure']
    sorted_sub_keys = sorted(sub.keys(), key=int)
    types = [sub[k]['type'] for k in sorted_sub_keys]
    n = len(types)

    # ========== 结果容器 ==========
    dist_nodes_lst = []
    dist_trans_dict = {}
    dist_long_dict = {}
    dist_beam_elink = {}
    dist_pile_elink = {}
    pile_xs = []
    pile_ys = []

    elink_idx = 0
    brake_idx = 0

    # ========== 遍历每个支撑位置 ==========
    for i in range(n):
        t = types[i]

        # -------------------------------------------------------
        # ① 桥台（不在此生成节点）
        # -------------------------------------------------------
        if t in ("桥台", "重力式桥台", "简易桥台"):
            pass

        # -------------------------------------------------------
        # ② 单排桩 → 1根横向分配梁 (dist_trans)
        # -------------------------------------------------------
        elif t == "单排桩":
            x = elink_x[elink_idx]
            elink_idx += 1

            p = sub[sorted_sub_keys[i]]
            trans_len = float(p.get('dist_trans_len', 7500))
            trans_offset = float(p.get('dist_trans_offset', 3750))  # 分配梁左端距中心线
            trans_sec = p.get('dist_trans_sec', '2I40a')
            # 解析 pile_trans_space "a,b@x": a=原cantilever1(桩偏移), b@x=桩横向布置
            pile_trans = p.get('pile_trans_space', '750,2@3000')
            cant1, pile_space_part = _parse_a_bx(pile_trans, 750)

            # y向范围（以中心线为基准）
            y_start = -trans_offset
            y_end = y_start + trans_len

            # 桩y坐标
            pit_ys = _calc_pile_ys_bx(y_start, cant1, pile_space_part)
            pile_ys.extend(pit_ys)

            # z坐标：横向分配梁中心 = 梁底面z - H_trans/2
            H = _get_sec_H(sections_dict, trans_sec, 400)
            z_trans = z_beam_bottom - beam_half_height - H / 2

            # 整合y坐标
            y_vals = sorted(set([y_start, y_end] + all_beam_ys + pit_ys))

            # 生成节点
            beam_nodes = []
            for y in y_vals:
                coord = [x, y, z_trans]
                beam_nodes.append([node_id, coord])
                dist_nodes_lst.append(node_id)
                if y in all_beam_ys:
                    dist_beam_elink[node_id] = coord
                if y in pit_ys:
                    dist_pile_elink[node_id] = coord
                node_id += 1

            dist_trans_dict[sorted_sub_keys[i]] = beam_nodes

        # -------------------------------------------------------
        # ③ 制动墩 → 2根横向(dist_trans) + N根纵向(dist_long)
        # -------------------------------------------------------
        elif t == "制动墩":
            x_prev = brake_elink_x[brake_idx]
            x_next = brake_elink_x[brake_idx + 1]
            brake_idx += 2

            p = sub[sorted_sub_keys[i]]

            # -- 横向参数 --
            trans_len = float(p.get('dist_trans_len', 7500))
            trans_offset = float(p.get('dist_trans_offset', 3750))
            trans_sec = p.get('dist_trans_sec', '2I40a')
            pile_trans = p.get('pile_trans_space', '750,2@3000')
            cant1, pile_space_part = _parse_a_bx(pile_trans, 750)

            # -- 纵向参数 --
            long_len = float(p.get('dist_long_len', 4500))
            long_offset = float(p.get('dist_long_center_offset', 2250))  # 纵梁左端距中心线
            long_sec = p.get('dist_long_sec', '2HM588X300')
            pile_long = p.get('pile_long_space', '750,3000')
            long_cant1, long_space_part = _parse_a_bx(pile_long, 750)

            # 截面高度
            H_trans = _get_sec_H(sections_dict, trans_sec, 400)
            H_long = _get_sec_H(sections_dict, long_sec, 588)

            # z坐标：横向分配梁中心 = 梁底面z - H_trans/2；纵向分配梁在其下方
            z_trans = z_beam_bottom - beam_half_height - H_trans / 2
            z_long = z_trans - H_trans / 2 - H_long / 2

            # --- 横向分配梁（前后各1根）---
            y_start = -trans_offset
            y_end = y_start + trans_len
            pit_ys = _calc_pile_ys_bx(y_start, cant1, pile_space_part)
            pile_ys.extend(pit_ys)
            y_vals = sorted(set([y_start, y_end] + all_beam_ys + pit_ys))

            for x_pos, tag in [(x_prev, '前'), (x_next, '后')]:
                beam_nodes = []
                for y in y_vals:
                    coord = [x_pos, y, z_trans]
                    beam_nodes.append([node_id, coord])
                    dist_nodes_lst.append(node_id)
                    if y in all_beam_ys:
                        dist_beam_elink[node_id] = coord
                    if y in pit_ys:
                        dist_pile_elink[node_id] = coord
                    node_id += 1
                dist_trans_dict[f"{sorted_sub_keys[i]}_{tag}"] = beam_nodes

            # --- 纵向分配梁 ---
            # 联中心线 = 两销轴中点，纵梁以联中心线为基准偏移
            joint_center = (x_prev + x_next) / 2
            x_long_start = joint_center - long_offset  # 纵梁左端
            x_long_end = x_long_start + long_len    # dist_long右端

            # 多排桩x坐标
            pit_xs = _calc_pile_xs_bx(x_long_start, long_cant1, long_space_part)
            pile_xs.extend(pit_xs)

            # x_vals 须包含 dist_trans 位置(前后弹连点)
            x_vals = sorted(set([x_long_start, x_long_end] + pit_xs + [x_prev, x_next]))

            for j, y_pile in enumerate(pit_ys):
                beam_nodes = []
                for xv in x_vals:
                    coord = [xv, y_pile, z_long]
                    beam_nodes.append([node_id, coord])
                    dist_nodes_lst.append(node_id)
                    if xv in pit_xs:
                        dist_pile_elink[node_id] = coord
                    node_id += 1
                dist_long_dict[f"{sorted_sub_keys[i]}#{j + 1}"] = beam_nodes

    # 去重（同一桩的y/x可能被多个支撑重复添加）
    pile_ys = sorted(set(pile_ys))
    pile_xs = sorted(set(pile_xs))

    return (dist_trans_dict, dist_long_dict,
            dist_beam_elink, dist_pile_elink,
            pile_xs, pile_ys, node_id)

def make_dist_elems(trestle_ui_dict, dist_trans_dict, dist_long_dict,
                    SECTION_name_search_dict, elem_id):
    """
    生成分配梁BEAM单元。

    遍历每根 dist_trans（沿y相邻节点）和 dist_long（沿x相邻节点），
    从 substructure 查截面，按命名规则生成 group_name。

    Returns
    -------
    dist_elem_lst : list[int] — 所有分配梁单元ID
    dist_elem_dict : dict — {group_name: [elem_id, ...]}
    dist_nodes_dict : dict — {group_name: [node_id, ...]}（供结构组用）
    dist_elem_mctlst : list[str] — MCT *ELEMENT 行
    elem_id : int — 更新后的单元号
    """
    sub = trestle_ui_dict['substructure']
    sections_dict = trestle_ui_dict.get('section', {})
    dist_elem_lst = []
    dist_elem_dict = {}
    dist_nodes_dict = {}
    dist_elem_mctlst = []

    # ====== 1. dist_trans（横向分配梁）======
    # 按桩号分组，同桩号内按x排序
    trans_by_sup = {}
    for key, nodes in dist_trans_dict.items():
        sup_key = key.split('_')[0]  # "3_前" → "3", 纯数字不变
        if not sup_key.isdigit():
            continue
        x0 = nodes[0][1][0]
        trans_by_sup.setdefault(sup_key, []).append((x0, key, nodes))

    for sup_key in sorted(trans_by_sup, key=int):
        beams = sorted(trans_by_sup[sup_key], key=lambda b: b[0])  # x↑
        sec_name = sub[sup_key].get('dist_trans_sec', '2I40a')
        sec_id = _get_sec_id(sections_dict, sec_name, 3)
        _dt_mat = _mat_id(trestle_ui_dict, _extract_brand(sub[sup_key].get('dist_trans_material')))

        for idx, (x0, key, nodes) in enumerate(beams):
            # 组名——单根无编号，多根编号
            if len(beams) == 1 and '_' not in key:
                gname = f"{sup_key}-横向分配梁"
            else:
                gname = f"{sup_key}-横向分配梁{idx + 1}"

            beam_elems = []
            beam_node_ids = [n[0] for n in nodes]
            for i in range(len(nodes) - 1):
                n1, n2 = nodes[i][0], nodes[i + 1][0]
                mct = beam_elem_mctstring_handle(elem_id, 'BEAM', _dt_mat, str(sec_id), n1, n2, '0')
                dist_elem_mctlst.append(mct)
                beam_elems.append(elem_id)
                elem_id += 1

            dist_elem_lst.extend(beam_elems)
            dist_elem_dict[gname] = beam_elems
            dist_nodes_dict[gname] = beam_node_ids

    # ====== 2. dist_long（纵向分配梁）======
    for key in sorted(dist_long_dict, key=lambda k: (int(k.split('#')[0]), k)):
        nodes = dist_long_dict[key]
        parts = key.split('#')
        if len(parts) != 2 or not parts[0].isdigit():
            continue
        sup_key = parts[0]
        seq = parts[1]

        sec_name = sub[sup_key].get('dist_long_sec', '2HM588X300')
        sec_id = _get_sec_id(sections_dict, sec_name, 6)
        _dl_mat = _mat_id(trestle_ui_dict, _extract_brand(sub[sup_key].get('dist_long_material')))
        gname = f"{sup_key}-纵向分配梁{seq}"

        beam_elems = []
        beam_node_ids = [n[0] for n in nodes]
        for i in range(len(nodes) - 1):
            n1, n2 = nodes[i][0], nodes[i + 1][0]
            mct = beam_elem_mctstring_handle(elem_id, 'BEAM', _dl_mat, str(sec_id), n1, n2, '0')
            dist_elem_mctlst.append(mct)
            beam_elems.append(elem_id)
            elem_id += 1

        dist_elem_lst.extend(beam_elems)
        dist_elem_dict[gname] = beam_elems
        dist_nodes_dict[gname] = beam_node_ids

    return dist_elem_dict, dist_nodes_dict, dist_elem_mctlst, elem_id


def make_dist_group(dist_elem_dict, dist_nodes_dict):
    """
    生成分配梁结构组。

    - "分配梁" 组包含所有分配梁节点+单元
    - 每根梁单独生成结构组（命名见 §4 规则）
    """
    groups = []
    # 全组
    all_nodes = sorted(set(n for nl in dist_nodes_dict.values() for n in nl))
    all_elems = sorted(set(e for el in dist_elem_dict.values() for e in el))
    groups.extend(_cmp_group('分配梁', all_nodes, all_elems))
    # 每根梁单独组
    for gname in sorted(dist_elem_dict):
        nodes = sorted(dist_nodes_dict.get(gname, []))
        elems = sorted(dist_elem_dict[gname])
        groups.extend(_cmp_group(gname, nodes, elems))
    return groups

# ============================== 9.钢管桩 ==============================
# ---- 钢管桩节点 → MCT 命令流 ----
def make_pile_nodes_mct(pile_head_nodes, pile_bottom_nodes, pile_water_top_nodes, pile_water_force_nodes, pile_brace_conn_nodes, pile_spring_nodes=None):
    """展平桩头/桩底/水流力顶点/联结系连接点/土弹簧所有节点"""
    mct_lst = []
    node_lst = []
    seen_ids = set()
    node_lst_set = []
    dicts = (pile_head_nodes, pile_bottom_nodes, pile_water_top_nodes, pile_water_force_nodes, pile_brace_conn_nodes)
    if pile_spring_nodes:
        dicts = dicts + (pile_spring_nodes,)
    for d in dicts:
        for nid, coord in d.items():
            node_lst.append((nid, [round(coord[0], 3), round(coord[1], 3), round(coord[2], 3)]))
    # print(node_lst)
    for item in node_lst:
        if item[0] not in seen_ids:
            seen_ids.add(item[0])
            node_lst_set.append(item)
    # print(node_lst_set)
    for nid, coord in  node_lst_set:
        mct_lst.append(node_mctstring_handle(nid, coord[0], coord[1], coord[2]))
    return mct_lst

def make_pile_nodes(trestle_ui_dict, dist_pile_elink, pile_xs, elink_x, brake_elink_x,
                    node_id):
    """
    生成钢管桩全部节点（桩头、水流力顶点、联结系连接点、桩底）。

    每根桩上按 z 从高到低生成：head → water_top → brace_conn_tops/bots → bottom。
    同 (x,y) 内的节点去重（z 容差 0.5mm），各类型分别保留在独立 dict 中供下游使用。

    Returns
    -------
    pile_head_nodes : dict {node_id: [x,y,z]}
    pile_bottom_nodes : dict {node_id: [x,y,z]}
    pile_water_force_nodes : dict {node_id: [x,y,z]}
    pile_water_force_nodes : dict {node_id: [x,y,z]} — 水流力合力点(入水深度1/3处)
    pile_brace_conn_nodes : dict {node_id: [x,y,z]} — 联结系连接点（供 brace 模块用）
    pile_all_node_ids : list[int]
    pile_by_support : dict {sup_key: {head:[], water_top:[], brace_conn:[[(top,bot),...]], bottom:[]}}
    node_id : int
    """
    sections_dict = trestle_ui_dict['section']
    pile_xs_set = set(pile_xs)
    sub = trestle_ui_dict['substructure']
    
    # 水位 → 模型z
    wl = float(trestle_ui_dict['basic']['water_level'].lstrip('+'))
    dl = float(trestle_ui_dict['basic']['deck_level'].lstrip('+'))
    z_water_surface = (wl - dl) * 1000

    # ---- 去重 dist_pile_elink 为唯一 (x,y) 桩位 ----
    uniq = {}
    for nid, coord in dist_pile_elink.items():
        xy = (coord[0], round(coord[1], 1))
        if xy not in uniq or coord[2] < uniq[xy][1][2]:
            uniq[xy] = (nid, coord)

    # ---- 构建 (x,y) → sup_key 映射 ----
    sorted_keys = sorted(sub.keys(), key=int)
    types = [sub[k]['type'] for k in sorted_keys]
    n = len(types)
    elink_idx = 0
    brake_idx = 0
    xy_to_sup = {}

    for i in range(n):
        t = types[i]
        if t in ("桥台", "重力式桥台", "简易桥台"):
            continue
        sup_key = sorted_keys[i]
        p = sub[sup_key]

        if t == "单排桩":
            x = elink_x[elink_idx]
            elink_idx += 1
            trans_offset = float(p.get('dist_trans_offset', 3750))
            pile_trans = p.get('pile_trans_space', '750,2@3000')
            cant1, pile_space_part = _parse_a_bx(pile_trans, 750)
            y_start = -trans_offset
            pit_ys = _calc_pile_ys_bx(y_start, cant1, pile_space_part)
            for py in pit_ys:
                xy_to_sup[(x, round(py, 1))] = sup_key

        elif t == "制动墩":
            x_prev = brake_elink_x[brake_idx]
            x_next = brake_elink_x[brake_idx + 1]  # noqa
            brake_idx += 2
            long_offset = float(p.get('dist_long_center_offset', 2250))
            pile_long = p.get('pile_long_space', '750,3000')
            long_cant1, long_space_part = _parse_a_bx(pile_long, 750)
            joint_center = (x_prev + x_next) / 2  # 联中心线
            x_long_start = joint_center - long_offset
            pit_xs = _calc_pile_xs_bx(x_long_start, long_cant1, long_space_part)

            trans_offset = float(p.get('dist_trans_offset', 3750))
            pile_trans = p.get('pile_trans_space', '750,2@3000')
            cant1, pile_space_part = _parse_a_bx(pile_trans, 750)
            y_start = -trans_offset
            pit_ys = _calc_pile_ys_bx(y_start, cant1, pile_space_part)
            for px in pit_xs:
                for py in pit_ys:
                    xy_to_sup[(round(px, 1), round(py, 1))] = sup_key
    pile_head_nodes = {}
    pile_bottom_nodes = {}
    pile_water_top_nodes = {}
    pile_water_force_nodes = {}
    pile_water_top_nodes = {}
    pile_water_force_nodes = {}
    pile_brace_conn_nodes = {}
    pile_spring_nodes = {}  # 土弹簧节点 {nid: [x, y, z]}
    pile_all_node_ids = []
    pile_by_support = {}

    # 用于后续 z 查询的合并 dict
    _all_pile_coords = {}  # node_id → [x,y,z]

    for (x, y), (_, coord) in uniq.items():
        z_beam_center = coord[2]
        sup_key = xy_to_sup.get((x, y))
        if sup_key is None:
            continue
        p = sub[sup_key]
        pile_len = float(p.get('pile_len', '15')) * 1000

        is_brake = x in pile_xs_set
        if is_brake:
            H = _get_sec_H(sections_dict, p.get('dist_long_sec', '2HM588X300'), 588)
        else:
            H = _get_sec_H(sections_dict, p.get('dist_trans_sec', '2I40a'), 400)
        z_head = z_beam_center - H / 2
        z_bottom = z_head - pile_len
        z_water_top = min(z_water_surface, z_head)

        # ---- 收集该桩所有候选 z 坐标 ----
        candidates = {}  # z → (set_of_tags, nid)
        def _add_candidate(z_val, tag):
            nonlocal node_id
            for ez, (tags, eid) in candidates.items():
                if abs(z_val - ez) < 0.5:
                    tags.add(tag)
                    return eid
            nid = node_id
            node_id += 1
            candidates[z_val] = ({tag}, nid)
            return nid

        # 1. 桩头
        nid_head = _add_candidate(z_head, 'head')

        # 2. 水流力顶点
        nid_wtop = _add_candidate(z_water_top, 'water_top')
        z_wf = z_bottom + (z_water_top - z_bottom) / 3
        nid_wforce = _add_candidate(z_wf, 'water_force')

        # 3. 联结系连接点（"/" 和 0 均表示无下一层）
        brace_form = p.get('brace_form', 'X')
        brace_layers = []  # [(top_nid, bot_nid), ...]
        if brace_form != '/':
            bs_raw = p.get('brace_space', '620')
            bs_filtered = ",".join(x for x in bs_raw.split(",") if x.strip() not in ("", "/", "0"))
            brace_spaces = midasdisttolst(bs_filtered) if bs_filtered.strip() else [0]
            try:
                brace_height = float(p.get('brace_height', '2500'))
            except ValueError:
                brace_height = 2500
            for spacing in brace_spaces:
                z_ct = z_head - spacing
                nid_ct = _add_candidate(z_ct, 'brace_conn')
                # "-" 型仅上弦杆，不需要下弦杆节点
                if brace_form == '-' or brace_height <= 0:
                    nid_cb = None
                else:
                    z_cb = z_ct - brace_height
                    nid_cb = _add_candidate(z_cb, 'brace_conn')
                brace_layers.append((nid_ct, nid_cb))

        # 4. 桩底
        nid_bot = _add_candidate(z_bottom, 'bottom')

        # 5. 土弹簧中间节点（地面线到桩底，按分层厚度插入）
        conn = trestle_ui_dict.get('connection', {})
        soil_spring_setting = conn.get("土弹簧", "/")
        is_soil_spring = isinstance(soil_spring_setting, dict) and soil_spring_setting != "/"
        soil_spring_nids = []
        if is_soil_spring:
            layer_thickness = 1.0
            try:
                layer_thickness = float(conn.get("土弹簧", {}).get("thickness", "1.0"))
            except (ValueError, TypeError, AttributeError):
                pass
            ground_y_m = float(p.get('ground_elevation', 0) or 0)
            # 模型坐标系：z = (实际标高 - 桥面标高) * 1000
            dl = float(trestle_ui_dict['basic']['deck_level'].lstrip('+'))
            z_ground = (ground_y_m - dl) * 1000
            # 从地面线向下按分层厚度插入节点
            z_insert = z_ground - layer_thickness * 1000
            while z_insert > z_bottom + 1:  # 容差 1mm
                nid_spring = _add_candidate(z_insert, 'soil_spring')
                soil_spring_nids.append(nid_spring)
                z_insert -= layer_thickness * 1000

        # ---- 写入各类型 dict ----
        for z_val, (tags, nid) in candidates.items():
            coord_xyz = [x, y, z_val]
            if 'head' in tags:
                pile_head_nodes[nid] = coord_xyz
            if 'water_top' in tags:
                pile_water_top_nodes[nid] = coord_xyz
            if 'water_force' in tags:
                pile_water_force_nodes[nid] = coord_xyz
            if 'brace_conn' in tags:
                pile_brace_conn_nodes[nid] = coord_xyz
            if 'bottom' in tags:
                pile_bottom_nodes[nid] = coord_xyz
            if 'soil_spring' in tags:
                pile_spring_nodes[nid] = coord_xyz
            pile_all_node_ids.append(nid)
            _all_pile_coords[nid] = coord_xyz
        # 确保 water_top 写入（与 head 同z时被去重）
        if nid_wtop not in pile_water_force_nodes and nid_wtop in pile_head_nodes:
            pile_water_force_nodes[nid_wtop] = pile_head_nodes[nid_wtop]

        # ---- 按支撑分组 ----
        ps = pile_by_support.setdefault(sup_key, {
            'head': [], 'water_top': [], 'water_force': [], 'brace_conn': [], 'bottom': [], 'soil_spring': []
        })
        ps['head'].append(nid_head)
        ps['water_top'].append(nid_wtop)
        ps['water_force'].append(nid_wforce)
        ps['brace_conn'].append(brace_layers)
        ps['bottom'].append(nid_bot)
        ps['soil_spring'].append(soil_spring_nids)

    return (pile_head_nodes, pile_bottom_nodes, pile_water_top_nodes,
            pile_water_force_nodes, pile_brace_conn_nodes, pile_spring_nodes,
            pile_all_node_ids, pile_by_support, node_id)


def make_pile_elems(trestle_ui_dict, pile_by_support,
                    SECTION_name_search_dict, elem_id,
                    pile_head_nodes, pile_bottom_nodes,
                    pile_water_top_nodes, pile_water_force_nodes,
                    pile_brace_conn_nodes, pile_spring_nodes=None):
    """
    钢管桩 BEAM 单元：对每根桩，将其下所有节点（同 x,y）按 z 排序后相邻连接。
    仅做竖向连接，不跨桩。
    """
    sub = trestle_ui_dict['substructure']
    sections_dict = trestle_ui_dict.get('section', {})
    pile_elem_dict = {}
    pile_elem_mctlst = []
    pile_bottom_ids = []

    # 合并坐标查询表（包括土弹簧节点）
    all_coords = {}
    for d in (pile_head_nodes, pile_bottom_nodes, pile_water_top_nodes,
              pile_water_force_nodes, pile_brace_conn_nodes):
        all_coords.update(d)
    if pile_spring_nodes:
        all_coords.update(pile_spring_nodes)

    for sup_key in sorted(pile_by_support, key=int):
        ps = pile_by_support[sup_key]
        sec_id = _get_sec_id(sections_dict, sub[sup_key].get('pile_sec', '2'), 2)
        _pile_mat = _mat_id(trestle_ui_dict, _extract_brand(sub[sup_key].get('pile_material')))
        gname = f"{sup_key}-钢管桩"
        beam_elems = []

        # zip 每根桩的各类节点（包括土弹簧节点）
        spring_nids_list = ps.get('soil_spring', [])
        for pile_idx, (h, wt, wf, brace_layers, b) in enumerate(zip(
            ps['head'], ps['water_top'], ps['water_force'], ps['brace_conn'], ps['bottom'])):
            # 收集该桩所有节点 (nid, z)
            raw = [(h, all_coords[h][2]),
                   (wt, all_coords[wt][2]),
                   (wf, all_coords[wf][2]),
                   (b, all_coords[b][2])]
            for top_nid, bot_nid in brace_layers:
                raw.append((top_nid, all_coords[top_nid][2]))
                if bot_nid is not None:
                    raw.append((bot_nid, all_coords[bot_nid][2]))
            # 仅添加当前桩的土弹簧节点
            if pile_idx < len(spring_nids_list):
                for nid in spring_nids_list[pile_idx]:
                    if nid in all_coords:
                        raw.append((nid, all_coords[nid][2]))

            # 按 z 从高到低排序，去重（z 容差 0.5mm）
            raw.sort(key=lambda it: it[1], reverse=True)
            uniq = []
            for nid, zv in raw:
                if not uniq or abs(zv - uniq[-1][1]) > 0.5:
                    uniq.append((nid, zv))

            for i in range(len(uniq) - 1):
                mct = beam_elem_mctstring_handle(elem_id, 'BEAM', _pile_mat, str(sec_id),
                                             uniq[i][0], uniq[i + 1][0], '0')
                pile_elem_mctlst.append(mct)
                beam_elems.append(elem_id)
                elem_id += 1

            pile_bottom_ids.append(b)

        pile_elem_dict[gname] = beam_elems

    return (pile_elem_dict, pile_elem_mctlst,
            pile_bottom_ids, elem_id)


def make_pile_group(pile_elem_dict, pile_all_node_ids, pile_by_support):
    """
    钢管桩结构组。
    - '钢管桩'：全组
    - '{sup_key}-钢管桩'：单支撑（含该支撑下所有桩节点）
    """
    groups = []
    all_elems = sorted(set(e for el in pile_elem_dict.values() for e in el))
    groups.extend(_cmp_group('钢管桩', pile_all_node_ids, all_elems))
    for gname in sorted(pile_elem_dict):
        elems = sorted(pile_elem_dict[gname])
        sup_key = gname.split('-')[0]
        ps = pile_by_support.get(sup_key, {})
        node_ids = []
        for nid in ps.get('head', []):
            node_ids.append(nid)
        for nid in ps.get('bottom', []):
            node_ids.append(nid)
        for nid in ps.get('water_top', []):
            node_ids.append(nid)
        for nid in ps.get('water_force', []):
            node_ids.append(nid)
        for layers in ps.get('brace_conn', []):
            for top_nid, bot_nid in layers:
                node_ids.append(top_nid)
                if bot_nid is not None:
                    node_ids.append(bot_nid)
        for nids in ps.get('soil_spring', []):
            node_ids.extend(nids)
        groups.extend(_cmp_group(gname, sorted(set(node_ids)), elems))
    return groups

# ============================== 土弹簧 ==============================
def make_soil_springs(trestle_ui_dict, pile_by_support, all_coords):
    """根据土层参数生成土弹簧 MCT 命令流（m 法）

    每个土弹簧节点独立计算刚度（因为不同节点深度不同，Ks 不同）。
    节点已在 make_pile_nodes 中按分层厚度插入。

    Parameters
    ----------
    trestle_ui_dict : dict
        完整 UI 数据字典
    pile_by_support : dict
        {sup_key: {'head': [...], 'bottom': [...], 'soil_spring': [[nid, ...], ...]}}
    all_coords : dict
        {node_id: [x, y, z]} 所有已生成节点坐标

    Returns
    -------
    list : MCT SPRING 行列表
    """
    sub = trestle_ui_dict.get('substructure', {})
    sections_dict = trestle_ui_dict.get('section', {})
    dl = float(trestle_ui_dict['basic']['deck_level'].lstrip('+'))

    spring_lines = []
    spring_debug = {}  # {sup_key: {nid: {z, depth, m, Ks, layer}}}

    for sup_key in sorted(pile_by_support, key=int):
        ps = pile_by_support[sup_key]
        p = sub.get(sup_key, {})
        soil_data = p.get('soil_data', {})
        ground_y_m = float(p.get('ground_elevation', 0) or 0)

        if not soil_data or not ground_y_m:
            print(f"  [警告] 桩 {sup_key}: soil_data={len(soil_data)} 层, ground_y={ground_y_m}m, 跳过")
            continue

        # 桩计算宽度 b0（ref 可能为名称或编号，_sec_D 兼容）
        sec_ref = p.get("pile_sec", "")
        d_mm = _sec_D(sections_dict, sec_ref, 820)
        d_m = d_mm / 1000.0
        b0 = 0.9 * (d_m + 1)

        # 构建土层列表（模型坐标系 z）
        soil_layers = []
        current_y_m = ground_y_m
        for layer_name, vals in soil_data.items():
            try:
                thickness = float(vals[0])
                c = float(vals[3]) if len(vals) > 3 else 0
                phi = float(vals[4]) if len(vals) > 4 else 0
            except (ValueError, TypeError, IndexError):
                continue
            bot_y_m = current_y_m - thickness
            z_top = (current_y_m - dl) * 1000
            z_bot = (bot_y_m - dl) * 1000
            soil_layers.append((z_top, z_bot, c, phi, layer_name))
            current_y_m = bot_y_m

        if not soil_layers:
            print(f"  [警告] 桩 {sup_key}: 土层列表为空")
            continue

        z_ground = (ground_y_m - dl) * 1000

        # 遍历该支撑每根桩的土弹簧节点
        spring_nids_list = ps.get('soil_spring', [])
        print(f"  桩 {sup_key}: soil_spring 分组数={len(spring_nids_list)}")
        for pile_idx, spring_nids in enumerate(spring_nids_list):
            print(f"    桩{sup_key}-{pile_idx}: 土弹簧节点数={len(spring_nids)}")
            pile_spring_info = {}
            for nid in spring_nids:
                if nid not in all_coords:
                    print(f"      [警告] nid={nid} 不在 all_coords 中")
                    continue
                nz = all_coords[nid][2]
                depth = z_ground - nz  # 地面以下深度（正值）

                # 找到该节点所在土层
                c_val, phi_val, layer_name = 0, 0, ""
                for zt, zb, c, phi, lname in soil_layers:
                    if zt >= nz >= zb:
                        c_val, phi_val, layer_name = c, phi, lname
                        break

                m_val = cal_m(phi_val, c_val)
                Ks = cal_Ks(m_val, depth / 1000, b0)  # depth 转为 m

                if Ks > 0:
                    spring_lines.append(
                        f"{nid}, LINEAR, NO, NO, NO, NO, NO, NO, "
                        f"{Ks}, {Ks}, 0, 0, 0, 0, NO, 0, 0, 0, 0, 0, 0, 土弹簧, 0, 0, 0, 0, 0")
                    pile_spring_info[nid] = {
                        "z": round(nz, 1),
                        "depth_m": round(depth / 1000, 3),
                        "m": m_val,
                        "Ks": Ks,
                        "layer": layer_name
                    }
            spring_debug[sup_key] = pile_spring_info
            if pile_spring_info:
                print(f"    → 生成 {len(pile_spring_info)} 个土弹簧")

    # 写入 trestle_ui_dict 供调试
    trestle_ui_dict["_soil_spring_debug"] = spring_debug
    return spring_lines


# ============================== 10.联结系 ==============================
def make_brace_nodes_mct(brace_center_nodes):
    """展平联结系中心节点为 MCT NODE 行"""
    mct_lst = []
    for nid, coord in brace_center_nodes.items():
        mct_lst.append(node_mctstring_handle(nid, coord[0], coord[1], coord[2]))
    return mct_lst

def make_brace_nodes(trestle_ui_dict, pile_head_nodes, pile_brace_conn_nodes,
                     pile_by_support, node_id):
    """
    生成联结系中心节点和分组。

    桩上连接点已在 make_pile_nodes 中生成并传入 pile_brace_conn_nodes。
    本函数根据 brace_form 按需创建节点：
      '-'  → 不创建中心节点（仅上弦杆）
      'X'  → 创建 X 交叉中心节点
      'X|X'→ 创建 X 交叉中心节点（供 make_brace_elems 补充中竖杆）
      'Z'  → 不创建中心节点（交替斜杆）

    Returns
    -------
    brace_center_nodes : dict {node_id: [x,y,z]} — X/X|X 形中心节点
    brace_groups : dict {gid: [up_l, dn_l, center|None, up_r, dn_r]}
    brace_all_node_ids : list[int]
    brace_by_support : dict {sup_key: [node_id, ...]}
    node_id : int
    """
    sub = trestle_ui_dict.get('substructure', {})
    brace_center_nodes = {}
    brace_groups = {}
    brace_all_node_ids = []
    brace_by_support = {}

    for sup_key, ps in sorted(pile_by_support.items(), key=lambda x: int(x[0])):
        if not ps['brace_conn'] or len(ps['head']) < 2:
            continue

        brace_form = sub.get(sup_key, {}).get('brace_form', 'X')
        if brace_form == '/':
            continue

        # 构建 pile_items：每个桩的 {x, y, layers}
        pile_items = []
        for pile_idx, brace_layers in enumerate(ps['brace_conn']):
            nid_head = ps['head'][pile_idx]
            x, y, _ = pile_head_nodes[nid_head]
            layers = []
            for top_nid, bot_nid in brace_layers:
                top_coord = pile_brace_conn_nodes[top_nid]
                bot_coord = pile_brace_conn_nodes[bot_nid] if bot_nid is not None else None
                layers.append({'top': (top_nid, top_coord),
                               'bot': (bot_nid, bot_coord) if bot_nid is not None else None})
            pile_items.append({'x': round(x, 1), 'y': round(y, 1), 'layers': layers})
            brace_by_support.setdefault(sup_key, []).append(top_nid)
            if bot_nid is not None:
                brace_by_support[sup_key].append(bot_nid)

        # 按 x 数量判断是否为制动墩
        xs = sorted(set(pi['x'] for pi in pile_items))
        if len(xs) > 1:
            mid_x = (xs[0] + xs[-1]) / 2
            left = [pi for pi in pile_items if pi['x'] <= mid_x]
            right = [pi for pi in pile_items if pi['x'] > mid_x]
            for side in [left, right]:
                _, node_id = _build_brace_pairs(
                    side, brace_groups, brace_center_nodes,
                    brace_all_node_ids, node_id, is_brake=False,
                    brace_form=brace_form)
            # 纵向（同 y、邻 x）：跨所有排
            _, node_id = _build_brace_pairs(
                pile_items, brace_groups, brace_center_nodes,
                brace_all_node_ids, node_id, long_only=True,
                brace_form=brace_form)
        else:
            _, node_id = _build_brace_pairs(
                pile_items, brace_groups, brace_center_nodes,
                brace_all_node_ids, node_id, is_brake=False,
                brace_form=brace_form)

    all_ids = brace_all_node_ids + list(brace_center_nodes.keys())
    if all_ids:
        node_id = max(all_ids) + 1

    return (brace_center_nodes, brace_groups, brace_all_node_ids,
            brace_by_support, node_id)


def _build_brace_pairs(pile_items, brace_groups, brace_center_nodes,
                       brace_all_node_ids, node_id, is_brake=False,
                       long_only=False, brace_form='X'):
    """
    在 pile_items 列表中生成联结系分组。
    - 非制动墩：仅横向（同 x 内邻 y）
    - 制动墩：横向（同 x 邻 y）+ 纵向（同 y 邻 x）
    brace_form 控制节点创建：
      '-'  → 不创建中心节点
      'X'  → 创建 X 交叉中心节点
      'X|X'→ 创建 X 交叉中心节点（供 make_brace_elems 补充中竖杆）
      'Z'  → 不创建中心节点
    返回 (next_gid, next_node_id)。
    """
    gid = len(brace_groups) + 1

    def _pair(p1, p2):
        nonlocal gid, node_id
        n_layers = len(p1['layers'])
        for li in range(n_layers):
            l1 = p1['layers'][li]
            l2 = p2['layers'][li]
            up_l, dn_l = l1['top'], l1['bot']
            up_r, dn_r = l2['top'], l2['bot']
            if dn_l is None or dn_r is None:
                brace_groups[gid] = [up_l, None, None, up_r, None]
            elif brace_form in ('X', 'X|X'):
                z_mid = (up_l[1][2] + dn_l[1][2] + up_r[1][2] + dn_r[1][2]) / 4
                y_mid = (p1['y'] + p2['y']) / 2
                x_mid = p1['x'] if abs(p1['x'] - p2['x']) < 0.5 else (p1['x'] + p2['x']) / 2
                coord_mid = [x_mid, y_mid, z_mid]
                brace_center_nodes[node_id] = coord_mid
                brace_all_node_ids.append(node_id)
                brace_groups[gid] = [up_l, dn_l, (node_id, coord_mid), up_r, dn_r]
                node_id += 1
            else:
                brace_groups[gid] = [up_l, dn_l, None, up_r, dn_r]
            gid += 1

    if not long_only:
        # 横向：同 x、邻 y
        by_x = {}
        for pi in pile_items:
            by_x.setdefault(pi['x'], []).append(pi)
        for x in sorted(by_x):
            lst = sorted(by_x[x], key=lambda pi: pi['y'])
            for i in range(len(lst) - 1):
                _pair(lst[i], lst[i + 1])

    if is_brake or long_only:
        # 纵向：同 y、邻 x
        by_y = {}
        for pi in pile_items:
            by_y.setdefault(pi['y'], []).append(pi)
        for y in sorted(by_y):
            lst = sorted(by_y[y], key=lambda pi: pi['x'])
            for i in range(len(lst) - 1):
                _pair(lst[i], lst[i + 1])

    return gid, node_id


def make_brace_elems(trestle_ui_dict, brace_groups, brace_by_support,
                     SECTION_name_search_dict, elem_id, node_id=None):
    """
    生成联结系 BEAM 单元。

    每组 5 节点 [up_l, dn_l, center, up_r, dn_r]，按 brace_form 选择：
    -:  仅上弦杆
    X:  弦杆(上下) + X 形斜杆 2 根
    X|X:弦杆(上下) + 左右各 X + 中竖杆
    Z:  弦杆(上下) + 单斜杆交替方向
    """
    sub = trestle_ui_dict['substructure']
    sections_dict = trestle_ui_dict.get('section', {})
    brace_elem_dict = {}
    brace_elem_mctlst = []
    brace_xx_node_mctlst = []

    for sup_key in sorted(brace_by_support, key=int):
        p = sub[sup_key]
        brace_form = p.get('brace_form', 'X')
        if brace_form == '/':
            continue
        sec_chord = _get_sec_id(sections_dict, p.get('brace_sec', '4'), 4)
        sec_diag = _get_sec_id(sections_dict, p.get('brace_sec_tilt', '5'), 5)
        # brace_material \u540c\u65f6\u63a7\u5236\u5f26\u6746\u548c\u659c\u6746/\u7ad6\u6746
        _brace_mat = _mat_id(trestle_ui_dict, _extract_brand(p.get('brace_material')))
        gname = f"{sup_key}-\u8054\u7ed3\u7cfb"
        g_elems = []
        bay_idx = 0

        for gid, nodes in sorted(brace_groups.items()):
            if nodes[0] is None:
                continue
            nid_check = nodes[0][0]
            found = False
            for sk, ids in brace_by_support.items():
                if nid_check in ids:
                    found = (sk == sup_key)
                    break
            if not found:
                continue

            up_l, dn_l, mid, up_r, dn_r = nodes

            if brace_form == '-':
                mct = beam_elem_mctstring_handle(elem_id, 'BEAM', _brace_mat, str(sec_chord), up_l[0], up_r[0], '0')
                brace_elem_mctlst.append(mct); g_elems.append(elem_id); elem_id += 1

            elif brace_form == 'X':
                pairs = [
                    (up_l[0], up_r[0]), (dn_l[0], dn_r[0]),
                    (up_l[0], mid[0]), (mid[0], dn_r[0]),
                    (dn_l[0], mid[0]), (mid[0], up_r[0]),
                ]
                for i, (n1, n2) in enumerate(pairs):
                    sec = sec_chord if i < 2 else sec_diag
                    mct = beam_elem_mctstring_handle(elem_id, 'BEAM', _brace_mat, str(sec), n1, n2, '0')
                    brace_elem_mctlst.append(mct); g_elems.append(elem_id); elem_id += 1

            elif brace_form == 'X|X':
                # 中竖杆顶/底节点（上下弦杆中点）
                x_mid = (up_l[1][0] + up_r[1][0]) / 2
                y_mid = (up_l[1][1] + up_r[1][1]) / 2
                z_top = (up_l[1][2] + up_r[1][2]) / 2
                z_bot = (dn_l[1][2] + dn_r[1][2]) / 2
                mt_nid = node_id; node_id += 1
                mb_nid = node_id; node_id += 1
                brace_xx_node_mctlst.append(node_mctstring_handle(mt_nid, x_mid, y_mid, z_top))
                brace_xx_node_mctlst.append(node_mctstring_handle(mb_nid, x_mid, y_mid, z_bot))
                # X 交叉中心节点（mid 由 _build_brace_pairs 按 X|X 创建）
                cx_nid = mid[0] if mid else None
                if cx_nid is not None:
                    pairs = [
                        (up_l[0], mt_nid), (mt_nid, up_r[0]),       # 上弦杆分段
                        (dn_l[0], mb_nid), (mb_nid, dn_r[0]),       # 下弦杆分段
                        (mt_nid, cx_nid), (cx_nid, mb_nid),         # 中竖杆分段
                        (up_l[0], cx_nid), (cx_nid, dn_r[0]),       # X 斜杆 ①
                        (dn_l[0], cx_nid), (cx_nid, up_r[0]),       # X 斜杆 ②
                    ]
                    for i, (n1, n2) in enumerate(pairs):
                        sec = sec_chord if i < 6 else sec_diag
                        mct = beam_elem_mctstring_handle(elem_id, 'BEAM', _brace_mat, str(sec), n1, n2, '0')
                        brace_elem_mctlst.append(mct); g_elems.append(elem_id); elem_id += 1
                else:
                    # 兜底：无中心节点时仅建弦杆+中竖杆
                    pairs = [
                        (up_l[0], mt_nid), (mt_nid, up_r[0]),
                        (dn_l[0], mb_nid), (mb_nid, dn_r[0]),
                        (mt_nid, mb_nid),
                    ]
                    for i, (n1, n2) in enumerate(pairs):
                        mct = beam_elem_mctstring_handle(elem_id, 'BEAM', _brace_mat, str(sec_chord), n1, n2, '0')
                        brace_elem_mctlst.append(mct); g_elems.append(elem_id); elem_id += 1

            elif brace_form == 'Z':
                if bay_idx % 2 == 0:
                    mct = beam_elem_mctstring_handle(elem_id, 'BEAM', _brace_mat, str(sec_diag), up_l[0], dn_r[0], '0')
                else:
                    mct = beam_elem_mctstring_handle(elem_id, 'BEAM', _brace_mat, str(sec_diag), dn_l[0], up_r[0], '0')
                brace_elem_mctlst.append(mct); g_elems.append(elem_id); elem_id += 1
                for n1, n2 in [(up_l[0], up_r[0]), (dn_l[0], dn_r[0])]:
                    mct = beam_elem_mctstring_handle(elem_id, 'BEAM', _brace_mat, str(sec_chord), n1, n2, '0')
                    brace_elem_mctlst.append(mct); g_elems.append(elem_id); elem_id += 1

            bay_idx += 1

        if g_elems:
            brace_elem_dict[gname] = g_elems

    return brace_elem_dict, brace_elem_mctlst, elem_id, brace_xx_node_mctlst, node_id

def make_brace_group(brace_elem_dict, brace_all_node_ids, brace_by_support):
    """
    联结系结构组。
    - '联结系'：全组（含联结系中心节点 + 联结系连接的桩头/连接节点）
    - '{sup_key}-联结系'：单支撑
    """
    # 联结系连接的桩头节点和连接节点
    brace_conn_node_ids = set()
    for sup_key, nids in brace_by_support.items():
        brace_conn_node_ids.update(nids)
    all_node_ids = sorted(set(brace_all_node_ids) | brace_conn_node_ids)

    groups = []
    all_elems = sorted(set(e for el in brace_elem_dict.values() for e in el))
    groups.extend(_cmp_group('联结系', all_node_ids, all_elems))
    for gname in sorted(brace_elem_dict):
        elems = sorted(brace_elem_dict[gname])
        sup_key = gname.split('-')[0]
        conn_ids = brace_by_support.get(sup_key, [])
        sub_nodes = sorted(set(brace_all_node_ids) | set(conn_ids))
        groups.extend(_cmp_group(gname, sub_nodes, elems))
    return groups

# ============================== 12、风荷载 ==============================
def make_wind_load(trestle_ui_dict, pile_head_nodes, brake_pile_head_nodes=None):
    """
    生成风荷载 *CONLOAD MCT 行。

    读取环境参数中的设计风速/地形/高度/系数，
    调用 cal_FengYa_JTG_T_3360_01_2018 计算风压，
    均分到所有桩顶节点，仅 Y 方向（横桥向），生成单一风荷载。

    Returns
    -------
    wind_lines : list[str] — ["*USE-STLD, 风荷载", "*CONLOAD", ...]
    """
    env = trestle_ui_dict.get("environment", {})
    wind_params = env.get("wind", {})

    # 1. 读取风参数（新UI：单一设计风速）
    try:
        design_v = float(wind_params.get("design_wind_speed", 20.0))
        kt = float(wind_params.get("terrain_factor", 1.0))
        Z = float(wind_params.get("girder_height", 6.4))
        CH = float(wind_params.get("transverse_coeff", 1.17))
        formula = wind_params.get("formula", "Ud = kf·kt·kh·U10")
        raw_gc = wind_params.get("ground_category", "A")
        gc_map = {"A":"A:海面、海岸、开阔水面","B":"B:田野、乡村、丛林、平坦开阔地",
        "C":"C:树木及地层建筑密集区、平缓丘陵地","D":"D:中高层建筑密集区、起伏较大的丘陵地"}
        prefix = raw_gc.split(":")[0].strip() if ":" in raw_gc else raw_gc
        dbfl = gc_map.get(prefix, gc_map.get(raw_gc, "A:海面、海岸、开阔水面"))
    except (ValueError, TypeError):
        print("  [警告] 风荷载参数读取失败，跳过")
        return []

    # 2. 贝雷梁总长
    span_str = trestle_ui_dict.get("basic", {}).get("span", "4@12")
    span_lst_m = midasdisttolst(span_str)
    L = sum(s * 1000 for s in span_lst_m)
    beam_count = len(span_lst_m)

    # 3. 计算风压（单一设计风速）
    _, _, _, _, _, _, _, _, _, _, Pa, _, _, _, _ = \
            cal_FengYa_JTG_T_3360_01_2018(design_v, kt, Z, L, beam_count, CH, dbfl, formula)

    # 总风力(N) = 风压(kPa) × 高度(m) × 1000
    F_total = Pa * Z * 1000

    # 4. 收集桩顶节点
    node_ids = list(pile_head_nodes.keys())
    if brake_pile_head_nodes:
        node_ids.extend(list(brake_pile_head_nodes.keys()))
    if not node_ids:
        print("  [信息] 无钢管桩，风荷载不加载")
        return []

    # 5. 均分到各桩顶 (Y方向负向)
    Fy_per_node = -F_total / len(node_ids)
    conload_lines = [
        f"{nid}, 0, {Fy_per_node:.2f}, 0, 0, 0, 0, , "
        for nid in sorted(node_ids)
    ]
    return conload_lines
  
# ============================== 13、水流力 ==============================
def make_water_load(trestle_ui_dict, pile_water_force_nodes, pile_bottom_nodes, pile_by_support=None):
    """
    生成水流力 *CONLOAD MCT 行。
    每个支撑独立计算 Cw 链（第一排 0.73），同一支撑内 y 从大到小递减。
    """
    env = trestle_ui_dict.get("environment", {})
    water_params = env.get("water", {})
    sections_dict = trestle_ui_dict.get("section", {})
    sub = trestle_ui_dict.get("substructure", {})
    v = float(water_params.get("flow_velocity", 2.7))
    if not pile_water_force_nodes:
        return []

    def _m1(lam):
        pts = [1,2,3,4,6,8,12,16,18,20]; ms = [-0.38,0.25,0.54,0.66,0.78,0.82,0.86,0.88,0.90,1.00]
        if lam <= pts[0]: return ms[0]
        if lam >= pts[-1]: return ms[-1]
        for i in range(len(pts)-1):
            if pts[i] <= lam <= pts[i+1]:
                k = (ms[i+1]-ms[i])/(pts[i+1]-pts[i])
                return ms[i] + k*(lam-pts[i])
        return ms[-1]

    z_top = max(c[2] for c in pile_water_force_nodes.values())
    z_bot = min(c[2] for c in pile_bottom_nodes.values())
    h = abs(z_bot - z_top) / 1000.0
    rho = 1000.0

    water_lines = []

    if pile_by_support is None:
        pile_by_support = {}
    for sup_key in sorted(pile_by_support, key=int):
        ps = pile_by_support[sup_key]
        p = sub.get(sup_key, {})
        # 桩径（ref 可能为名称或编号，_sec_D 兼容）
        sec_ref = p.get("pile_sec", "")
        d_m = _sec_D(sections_dict, sec_ref, 630) / 1000.0
        # 排间距（先剥离 a,b@x 格式的前缀 a）
        pts = p.get("pile_trans_space", "")
        if pts and pts != "/":
            _, space_part = _parse_a_bx(pts, 750)
            spacings = [sp for sp in midasdisttolst(space_part) if sp > 0]
        else:
            spacings = []
        if not spacings:
            continue
        # Cw 链
        Cw = 0.73; Cw_list = [0.73]
        for L in spacings:
            lam = (L - d_m*1000) / (d_m*1000) if d_m > 0 else 0
            Cw = round(Cw * _m1(lam), 2)
            Cw_list.append(Cw)
        # 该支撑的 water_top 节点按 y 分组
        y_nodes = {}
        for nid in ps['water_force']:
            if nid in pile_water_force_nodes:
                c = pile_water_force_nodes[nid]; yk = round(c[1], 1)
                y_nodes.setdefault(yk, []).append((nid, c))
        for idx, y in enumerate(sorted(y_nodes, reverse=True)):
            cw = Cw_list[idx] if idx < len(Cw_list) else Cw_list[-1]
            fw = round(0.5 * rho * cw * v * v * d_m * h / 1000, 2)
            for nid, _ in y_nodes[y]:
                water_lines.append(f"{nid}, 0, -{fw:.2f}, 0, 0, 0, 0, , ")
    return water_lines


# ============================== 14、移动荷载 ==============================
def build_vehicle_lines(vehicle_dict):
    """
    生成 *VEHICLE MCT 行。

    三类车辆映射：
      - 公路标准荷载 → NAME={name}, 1, {std}
      - 吊装设备/钻孔设备 → NAME={name}, 2, TRUCK, 4, {dw1}, {dd1}
      - 一般车辆 → NAME={name}, 2, TRUCK, 1, 0, 0, 0 + 荷载-间距行
    """
    lines = ["*VEHICLE"]
    for name, info in sorted(vehicle_dict.items()):
        vt = info.get("vehicle_type", "")
        vp = info.get("vehicle_params", [])
        if vt == "公路标准荷载":
            std = vp[1] if len(vp) > 1 else "CH-CD"
            lines.append(f"   NAME={name}, 1, {std}, JTGB01-2014")

        elif vt == "吊装设备":
            # params[1]=自重(kN), params[4]=接地长(m), params[5]=接地宽(m)
            self_wt = float(vp[1]) if len(vp) > 1 else 1500  # kN
            contact_len = float(vp[4]) if len(vp) > 4 else 7.182  # m
            dd1 = contact_len * 1000  # mm
            dw1 = (self_wt * 1000) / dd1  # N/mm （线荷载）
            lines.append(f"   NAME={name}, 2, TRUCK, 4, {dw1:.3f}, {dd1:.3f}")
        elif vt == "钻孔设备":
            # params[1]=自重(kN), params[3]=接地长(m), params[4]=接地宽(m)
            self_wt = float(vp[1]) if len(vp) > 1 else 1500  # kN
            contact_len = float(vp[3]) if len(vp) > 3 else 7.182  # m
            dd1 = contact_len * 1000  # mm
            dw1 = (self_wt * 1000) / dd1  # N/mm （线荷载）
            lines.append(f"   NAME={name}, 2, TRUCK, 4, {dw1:.3f}, {dd1:.3f}")
        elif vt == "一般车辆":
            lines.append(f"   NAME={name}, 2, TRUCK, 1, 0, 0, 0")
            # params = [轴距, 轴重CSV, 距前轴CSV]
            pairs = []
            if len(vp) >= 3:
                loads = [x.strip() for x in vp[1].split(",") if x.strip()]
                dists = [x.strip() for x in vp[2].split(",") if x.strip()]
                for ld, dd in zip(loads, dists):
                    load_n = float(ld) * 1000      # kN → N
                    dist_mm = float(dd) * 1000     # m → mm
                    pairs.append(f"{load_n:.0f}, {dist_mm:.3f}")
            lines.append("      " + ", ".join(pairs))

    return lines


def build_line_lanes(move_load_cases, vehicle_dict=None, virtual_elem_ids=None):
    """
    生成 *LINELANE(CH) MCT 行（CROSS 格式，基于虚拟梁单元）。

    每条工况对应一条车道，格式：
      NAME=车道{n}, CROSS, 横向联系梁, 0, 0, {方向}, {轮距}, 3000, NO, 3
        {elem1}, {ecc}, 1.5, YES, 1, {elem2}, {ecc}, 1.5, NO, 1
        ...

    轮距（两轮中心距mm）：
      - 公路标准荷载：vehicle_params[0]（轴距,m）×1000
      - 吊装设备/钻孔设备：vehicle_params[5]（接地宽,m）×1000
      - 一般车辆：默认 2000mm

    方向：往返→forward, 向前→forward, 向后→backward
    """
    if not virtual_elem_ids:
        return []

    lines = ["*LINELANE(CH)"]
    for cname, cinfo in sorted(move_load_cases.items()):
        ecc_m = float(cinfo.get("lane_ecc", 0))   # m
        ecc = ecc_m * 1000                         # → mm
        direc = {"向前": "forward", "向后": "backward"}.get(
            cinfo.get("direction", ""), "forward")

        # 轮距（mm）：每次循环初始化默认值
        wheel_dist = 2000
        veh_name = cinfo.get("vehicle", "")
        if vehicle_dict and veh_name in vehicle_dict:
            vp = vehicle_dict[veh_name].get("vehicle_params", [])
            vt = vehicle_dict[veh_name].get("vehicle_type", "")
            if vt == "公路标准荷载" and vp:
                wheel_dist = float(vp[0]) * 1000
            elif vt in ("吊装设备", "钻孔设备") and vp:
                wheel_dist = float(vp[0]) * 1000
            elif vt == "一般车辆" and vp and vp[0]:
                wheel_dist = float(vp[0]) * 1000

        # 车道头部行（车道名 = 工况名）
        lines.append(
            f"   NAME={cname},  CROSS, 横向联系梁, 0, 0, {direc}, {wheel_dist:.0f}, 3000, NO, 3000")

        # 虚拟梁单元行（两个单元一行，第一个 YES 其余 NO）
        n = len(virtual_elem_ids)
        for j in range(0, n, 2):
            eid1 = virtual_elem_ids[j]
            tag1 = "YES" if j == 0 else "NO"
            if j + 1 < n:
                eid2 = virtual_elem_ids[j + 1]
                lines.append(f"    {eid1}, {ecc:.0f}, 1500, {tag1}, 1, {eid2}, {ecc:.0f}, 1500, NO, 1")
            else:
                lines.append(f"    {eid1}, {ecc:.0f}, 1500, {tag1}, 1")

    return lines


def build_mvldcase_lines(move_load_cases):
    """
    生成 *MVLDCASE(CH) MCT 行。

    每条工况一行 NAME + 影响线系数(固定) + VL 车辆-车道关联。
    """
    lines = ["*MVLDCASE(CH)"]
    for cname, cinfo in sorted(move_load_cases.items()):
        veh = cinfo.get("vehicle", "")
        fac = float(cinfo.get("factor", 1.0))
        lines.append(f"   NAME={cname}, , NO, 0, 1, 0")
        lines.append("        1, 1, 0.8, 0.67, 0.6, 0.55, 0.55, 0.55")
        lines.append("        1, 1, 0.78, 0.67, 0.6, 0.55, 0.52, 0.5")
        lines.append(f"        {fac:.2f}, 1, 0.78, 0.67, 0.6, 0.55, 0.52, 0.5")
        lines.append(f"        VL, {veh}, 1, 1, 1, {cname}")
    return lines


def build_move_load(trestle_ui_dict, virtual_elem_ids=None):
    """
    组装完整的移动荷载 MCT 段。

    Parameters
    ----------
    trestle_ui_dict : dict
    virtual_elem_ids : list[int], optional — 虚拟梁单元ID列表（供车道CROSS定义）

    Returns
    -------
    move_load_lst : list[str] — 工况名列表
    move_lines : list[str] — *VEHICLE + *LINELANE + *MVLDCASE 全部 MCT 行
    """
    vd = trestle_ui_dict.get("vehicle", {})
    ml = trestle_ui_dict.get("move_load_cases", {})
    move_load_lst = list(ml.keys())
    if not ml:
        return move_load_lst, []

    veh_lines = build_vehicle_lines(vd)
    lane_lines = build_line_lanes(ml, vd, virtual_elem_ids)
    mvl_lines = build_mvldcase_lines(ml)

    move_lines = veh_lines + lane_lines + mvl_lines
    print(f"[MCT] 移动荷载: {len(move_load_lst)} 个工况, {len(move_lines)} 行 MCT")
    return move_load_lst, move_lines


# ============================== 15、静载工况 ==============================
# 计算履带吊
def crawler_crane_load_cal(selfweight, lifting_load, track_length, track_width):
    # 履带吊正吊及侧吊计算
    '''
    track_gauge  履带中心间距
    selfweight   履带吊自重
    lifting_load 吊重
    boom_length  吊臂
    track_length 单侧履带长
    track_width  单侧履带宽
    '''
    # 正吊
    pkmax = (selfweight+lifting_load)/(track_length*track_width)
    pkmin = 0
    zd_pk_lst = [round(pkmax,1), round(pkmin,1)]
    # 侧吊
    pkmax = 0.7*(selfweight+lifting_load)/(track_length*track_width)
    pkmin = 0.3*(selfweight+lifting_load)/(track_length*track_width)
    cd_pk_lst = [round(pkmax,1), round(pkmin,1)]
    return zd_pk_lst, cd_pk_lst


# 生成静载工况所需要的信息
def make_static_load_cases_info(trestle_ui_dict):
    stldcase_info_dict = {}
    static_load_vehicle_dict = trestle_ui_dict['static_load_cases']

    # cache 已展开为逐个 span+position 的独立条目
    # 按 (vehicle, type) 分组，每组只需计算一次 pk_lst / pnloadtype_namelst
    vt_groups = {}       # (vehicle, type) → group_index
    vt_case_lists = {}   # (vehicle, type) → [case_name, ...]

    for key, value in static_load_vehicle_dict.items():
        vehicle = value['vehicle']
        vehicle_type = value['type']
        group_key = (vehicle, vehicle_type)
        if group_key not in vt_groups:
            vt_groups[group_key] = len(vt_groups)
            vt_case_lists[group_key] = []
        vt_case_lists[group_key].append(key)

    for group_key, case_names in vt_case_lists.items():
        vehicle, vehicle_type = group_key
        pk_lst = []
        load_area_lst = []
        pnloadtype_namelst = []
        vehicle_param_lst = trestle_ui_dict['vehicle'][vehicle]['vehicle_params']
        if '履带吊' in vehicle:
            track_gauge, selfweight, lifting_load, boom_length, track_length, track_width = [float(x) for x in vehicle_param_lst]
            if '正吊' in vehicle_type:
                pk_lst = crawler_crane_load_cal(selfweight, lifting_load, track_length, track_width)[0]
                pnloadtype_namelst = [vehicle + vehicle_type]
            if '侧吊' in vehicle_type:
                pk_lst = crawler_crane_load_cal(selfweight, lifting_load, track_length, track_width)[1]
                pnloadtype_namelst = [vehicle + vehicle_type + '(max)', vehicle + vehicle_type + '(min)']
            load_area_lst = [track_gauge, track_length, track_width]

        stldcase_info_dict[vt_groups[group_key]] = {
            'vehicle': vehicle,
            '分配面荷载类型': pnloadtype_namelst,
            '荷载': pk_lst,
            '分布': load_area_lst,
            '工况': case_names,  # 直接用 cache key 作为工况名
        }

    return stldcase_info_dict


# 生成静载工况的mct命令流
def make_static_load_cases(stldcase_info_dict):
    stldcase_mctstring = []
    for value in stldcase_info_dict.values():
        for case_name in value['工况']:
            stldcase_mctstring.append('{}, USER, '.format(case_name))
    return stldcase_mctstring


# 生成分配面荷载类型
def make_pnloadtype(stldcase_info_dict):
    pnloadtype_mctstring = []
    for value in stldcase_info_dict.values():
        s, l, d = [x*1000 for x in value['分布']]
        pkmax, pkmin = [round(x/1000, 6) for x in value['荷载']]
        for name in value['分配面荷载类型']:
            if '正向' in name:
                pnloadtype_mctstring.append('NAME={}, AREA, \n CP_Y={} \n DATA=NO, NO, {}, {}, {},  {}, {}, {},  {}, {}, {},  {}, {}, {}'.format(name, s, 0, 0, pkmin*-1, 0, d, pkmin*-1, l, d, pkmax*-1, l, 0, pkmax*-1))
            elif '反向' in name:
                pnloadtype_mctstring.append('NAME={}, AREA, \n CP_Y={} \n DATA=NO, NO, {}, {}, {},  {}, {}, {},  {}, {}, {},  {}, {}, {}'.format(name, s, 0, 0, pkmax*-1, 0, d, pkmax*-1, l, d, pkmin*-1, l, 0, pkmin*-1))
            elif 'max' in name:
                pnloadtype_mctstring.append('NAME={}, AREA, \n DATA=YES, NO, {}, {}, {},  {}, {}, {},  {}, {}, {},  {}, {}, {}'.format(name, 0, 0, pkmax*-1, 0, d, 0, l, d, 0, l, 0, 0))
            elif 'min' in name:
                pnloadtype_mctstring.append('NAME={}, AREA, \n DATA=YES, NO, {}, {}, {},  {}, {}, {},  {}, {}, {},  {}, {}, {}'.format(name, 0, 0, pkmin*-1, 0, d, 0, l, d, 0, l, 0, 0))
    return pnloadtype_mctstring


# 生成分配面荷载对应静载工况
def make_vehicle_static_load(trestle_ui_dict, stldcase_info_dict):
    # 获取墩号对应的x坐标, 对于制动墩取两个x
    def process_x_lst(x_lst, types):
        result = []
        offset = 0
        for x, t in zip(x_lst, types):
            if t == '制动墩':
                offset += 200
            adjusted_x = x + offset
            if t == '制动墩':
                result.append([adjusted_x - 290, adjusted_x + 90])
            else:
                result.append([adjusted_x])
        return result

    def calc_span_coords(dun_x_lst):
        """
        根据墩坐标计算跨数坐标
        - 单坐标墩：该坐标既作为前一跨终点，也作为后一跨起点
        - 双坐标墩（制动墩）：左坐标作为前一跨终点，右坐标作为后一跨起点
        """
        span_dict = {}
        span_num = 1
        prev_end = None
        for coords in dun_x_lst:
            if len(coords) == 1:
                # 单坐标墩
                if prev_end is None:
                    prev_end = coords[0]  # 第一个墩，记录起点
                else:
                    # 完成前一跨，并开始下一跨（该坐标既是终点也是起点）
                    span_dict[str(span_num)] = [prev_end, coords[0]]
                    span_num += 1
                    prev_end = coords[0]
            else:
                # 双坐标墩（制动墩）
                # 左坐标作为前一跨终点
                if prev_end is not None:
                    span_dict[str(span_num)] = [prev_end, coords[0]]
                    span_num += 1
                # 右坐标作为后一跨起点
                prev_end = coords[1]
        return span_dict

    # ── 计算贝雷梁分布中心 Y ──
    bridge_type = _get_bridge_type(trestle_ui_dict.get('components', {}))
    comp = trestle_ui_dict['components'][bridge_type]
    beam_space = comp.get('beam_space', '5@900')
    beam_first_space = float(comp.get('beam_first_space', 0))
    beam_y_lst = midasdisttolst2(-beam_first_space, beam_space)
    middle_y = (min(beam_y_lst) + max(beam_y_lst)) / 2  # 贝雷梁分布中心 Y（mm）

    dun_dict = {}
    planeload_mctstring = []
    pile_numlst = trestle_ui_dict['substructure'].keys()
    span = trestle_ui_dict['basic']['span']
    x_lst = [x*1000 for x in midasdisttolst2(0, span)]
    type_lst = []
    for v in trestle_ui_dict['substructure'].values():
        type_lst.append(v['type'])
    dun_x_lst = process_x_lst(x_lst, type_lst)
    for i, x in enumerate(pile_numlst):
        dun_dict[x] = dun_x_lst[i]
    kua_dict = calc_span_coords(dun_x_lst)

    for value in stldcase_info_dict.values():
        static_load_cases_lst = value['工况']
        pnloadtype_lst = value['分配面荷载类型']
        s, l, d = [x*1000 for x in value['分布']]
        for static_load_case_name in static_load_cases_lst:
            y_center = middle_y
            # 一个 name 是一个静载工况
            # 从"跨N"或"桩N"中提取跨号/墩号，避免车辆名中的数字干扰
            _m = re.search(r'跨(\d+)|桩(\d+)', static_load_case_name)
            num = str(int(_m.group(1) or _m.group(2))) if _m else ''
            if '左侧' in static_load_case_name:
                for pnloadtype_name in pnloadtype_lst:
                    if 'max' in pnloadtype_name:
                        x1, y1, z1 = kua_dict[num][0], y_center - s/2 - d/2, 0
                    elif 'min' in pnloadtype_name:
                        x1, y1, z1 = kua_dict[num][0], y_center + s/2 - d/2, 0
                    else:
                        x1, y1, z1 = kua_dict[num][0], y_center - s/2 - d/2, 0
                    x2, y2, z2 = x1 + 1000, y1, 0
                    x3, y3, z3 = x2, y1 + 1000, 0
                    mctstring = '{}, {}, PLATE, \n LPLANE, , 1, NLP, NO, \n {}, {}, {}, {}, {}, {},  {}, {}, {}, 1, NO'.format(static_load_case_name, pnloadtype_name, x1, y1, z1, x2, y2, z2, x3, y3, z3)
                    planeload_mctstring.append(mctstring)
            elif '右侧' in static_load_case_name:
                for pnloadtype_name in pnloadtype_lst:
                    if 'max' in pnloadtype_name:
                        x1, y1, z1 = kua_dict[num][-1] - l, y_center - s/2 - d/2, 0
                    elif 'min' in pnloadtype_name:
                        x1, y1, z1 = kua_dict[num][-1] - l, y_center + s/2 - d/2, 0
                    else:
                        x1, y1, z1 = kua_dict[num][-1] - l, y_center - s/2 - d/2, 0
                    x2, y2, z2 = x1 + 1000, y1, 0
                    x3, y3, z3 = x2, y1 + 1000, 0
                    mctstring = '{}, {}, PLATE, \n LPLANE, , 1, NLP, NO, \n {}, {}, {}, {}, {}, {},  {}, {}, {}, 1, NO'.format(static_load_case_name, pnloadtype_name, x1, y1, z1, x2, y2, z2, x3, y3, z3)
                    planeload_mctstring.append(mctstring)
            elif '跨中' in static_load_case_name:
                for pnloadtype_name in pnloadtype_lst:
                    if 'max' in pnloadtype_name:
                        x1, y1, z1 = (kua_dict[num][0] + kua_dict[num][-1])/2 - l/2, y_center - s/2 - d/2, 0
                    elif 'min' in pnloadtype_name:
                        x1, y1, z1 = (kua_dict[num][0] + kua_dict[num][-1])/2 - l/2, y_center + s/2 - d/2, 0
                    else:
                        x1, y1, z1 = (kua_dict[num][0] + kua_dict[num][-1])/2 - l/2, y_center - s/2 - d/2, 0
                    x2, y2, z2 = x1 + 1000, y1, 0
                    x3, y3, z3 = x2, y1 + 1000, 0
                    mctstring = '{}, {}, PLATE, \n LPLANE, , 1, NLP, NO, \n {}, {}, {}, {}, {}, {},  {}, {}, {}, 1, NO'.format(static_load_case_name, pnloadtype_name, x1, y1, z1, x2, y2, z2, x3, y3, z3)
                    planeload_mctstring.append(mctstring)
            elif '桩' in static_load_case_name:
                for pnloadtype_name in pnloadtype_lst:
                    if 'max' in pnloadtype_name:
                        x1, y1, z1 = (dun_dict[num][0] + dun_dict[num][-1])/2 - l/2, y_center - s/2 - d/2, 0
                    elif 'min' in pnloadtype_name:
                        x1, y1, z1 = (dun_dict[num][0] + dun_dict[num][-1])/2 - l/2, y_center + s/2 - d/2, 0
                    else:
                        x1, y1, z1 = (dun_dict[num][0] + dun_dict[num][-1])/2 - l/2, y_center - s/2 - d/2, 0
                    x2, y2, z2 = x1 + 1000, y1, 0
                    x3, y3, z3 = x2, y1 + 1000, 0
                    mctstring = '{}, {}, PLATE, \n LPLANE, , 1, NLP, NO, \n {}, {}, {}, {}, {}, {},  {}, {}, {}, 1, NO'.format(static_load_case_name, pnloadtype_name, x1, y1, z1, x2, y2, z2, x3, y3, z3)      
                    planeload_mctstring.append(mctstring)
    return planeload_mctstring

# ============================== 16、荷载组合 ==============================
def load_combinations(trestle_ui_dict):
    """生成 *LOADCOMB MCT 命令流

    格式：
      NAME=组合名, GEN/ENV, ACTIVE, 0, iTYPE, , 0, 0, 0, 1
          前缀, 工况名, 系数, 前缀, 工况名, 系数, ...

    类型: 相加→iTYPE=0(GEN), 包络→iTYPE=1(ENV)
    前缀: MV=移动荷载工况, ST=其他静力工况
    """
    comb_dict = trestle_ui_dict.get('load_combination', {})
    if not comb_dict:
        print("  [信息] 无荷载组合定义，跳过 *LOADCOMB")
        return []

    # 收集名称集用于判断前缀
    move_names = set(trestle_ui_dict.get('move_load_cases', {}).keys())
    comb_names = set(comb_dict.keys())

    lines = []
    for comb_name, comb_info in comb_dict.items():
        ctype = comb_info.get('type', '相加')
        itype = '1' if ctype == '包络' else '0'
        cases = comb_info.get('cases', {})

        lines.append(f"NAME={comb_name}, GEN, ACTIVE, 0, {itype}, , 0, 0, 0, 1")

        # 前缀：MV=移动, CB=其他组合, ST=静力
        parts = []
        for case_name, coeff in cases.items():
            # 兼容新旧格式：新格式为 dict {gamma, psi}，旧格式为 float
            if isinstance(coeff, dict):
                gamma = float(coeff.get("gamma", 1.0))
                psi = float(coeff.get("psi", 1.0))
                factor = round(gamma * psi, 4)
            else:
                factor = float(coeff)
            if case_name in move_names:
                prefix = 'MV'
            elif case_name in comb_names:
                prefix = 'CB'
            else:
                prefix = 'ST'
            parts.append(f"{prefix}, {case_name}, {factor}")
        if parts:
            # 每4个一组换行，避免单行过长
            for i in range(0, len(parts), 4):
                lines.append("    " + ", ".join(parts[i:i+4]))

    print(f"[MCT] 荷载组合 {len(comb_dict)} 个")
    return lines








# ============================== 总函数 ==============================
# ── mcb_generate 分段模块（全量模块化）──

def build_section_material_module(trestle_ui_dict):
    """截面/材质/板厚模块：SECTIONS + MATERIAL + THICKNESS + 截面名检索字典"""
    sections_dict = trestle_ui_dict['section']
    SECTION_lst = SECTIONS(sections_dict)
    _steel_dict = trestle_ui_dict.get("steel_dict")
    _concrete_dict = trestle_ui_dict.get("concrete_dict")
    _sub = trestle_ui_dict.get('substructure', {})
    bridge_type = _get_bridge_type(trestle_ui_dict.get('components', {}))
    _comp = trestle_ui_dict['components'][bridge_type]
    deck_type = _comp.get('deck_type', '单层钢面板')
    _concrete_grade = _comp.get('concrete_grade', 'C30') if deck_type == "混凝土桥面板" else None

    _mat_brands = []  # 去重有序
    _seen = set()
    def _add_brand(val):
        b = _extract_brand(val)
        if b and b not in _seen:
            _seen.add(b)
            _mat_brands.append(b)
    _add_brand(_comp.get('deck_material'))
    _add_brand(_comp.get('rib_trans_material'))
    _add_brand(_comp.get('rib_long_material'))
    _add_brand(_comp.get('beam_material'))
    for sup_data in _sub.values():
        _add_brand(sup_data.get('pile_material'))
        _add_brand(sup_data.get('dist_trans_material'))
        _add_brand(sup_data.get('dist_long_material'))
        _add_brand(sup_data.get('brace_material'))

    MATERIAL_lst, material_id_map = MATERIAL(sections_dict, _concrete_grade,
                                              steel_dict=_steel_dict,
                                              concrete_dict=_concrete_dict,
                                              component_brands=_mat_brands)
    trestle_ui_dict['_material_id_map'] = material_id_map
    print(f"[MCT] 材质 {len(MATERIAL_lst)} 条: {material_id_map}")
    print(f"[MCT] 截面 {len(SECTION_lst)} 条")

    SECTION_name_search_dict = {}
    for line in SECTION_lst:
        parts = [item.strip() for item in line.split(',')]
        key = parts[2]
        value = parts[:2] + parts[3:]
        SECTION_name_search_dict[key] = value

    thickness = _get_component(trestle_ui_dict, 'deck_thickness', '10')
    THICKNESS_lst = ['1, VALUE, 1, YES, {}, 0,  NO, 0, 0'.format(thickness)]
    if deck_type == "混凝土桥面板":
        THICKNESS_lst.append('2, VALUE, 1, YES, {}, 0,  NO, 0, 0'.format(thickness))

    return {
        'SECTION_lst': SECTION_lst, 'MATERIAL_lst': MATERIAL_lst, 'THICKNESS_lst': THICKNESS_lst,
        'SECTION_name_search_dict': SECTION_name_search_dict, 'sections_dict': sections_dict,
        'material_id_map': material_id_map,
        'bridge_type': bridge_type, 'deck_type': deck_type, 'comp': _comp,
    }


def build_geometry_module(trestle_ui_dict):
    """跨径/桥面系几何/弹连点模块：span_dict + rib_info + 分配梁与纵梁弹连点X"""
    bridge_type = _get_bridge_type(trestle_ui_dict.get('components', {}))
    _comp = trestle_ui_dict['components'][bridge_type]
    deck_type = _comp.get('deck_type', '单层钢面板')
    # 双层钢面板分节参数
    _seg_enabled = _comp.get('seg_enabled', False) if deck_type == "双层钢面板" else False
    _panel_length = _comp.get('panel_length', None) if _seg_enabled else None
    if _panel_length:
        try:
            _panel_length = float(_panel_length)
        except (ValueError, TypeError):
            _panel_length = None
            _seg_enabled = False

    span_dict = trestle_span(trestle_ui_dict)
    print(f"[MCT] {len(span_dict)} 跨: {list(span_dict.keys())}")

    rib_info = calc_deck_geometry(trestle_ui_dict, span_dict)
    is_double = rib_info["is_double"]

    # 分配梁与纵梁弹连点（被 beam/dist/pile 三模块共用）
    span_str = trestle_ui_dict['basic']['span']
    span_lst = midasdisttolst(span_str)
    pile_elink_x, brake_elink_x = beam_bottom_elink_x(span_dict, trestle_ui_dict['substructure'], span_lst)
    fpl_elink_x = sorted(pile_elink_x + brake_elink_x)

    return {
        'span_dict': span_dict, 'rib_info': rib_info, 'is_double': is_double,
        'seg_enabled': _seg_enabled, 'panel_length': _panel_length,
        'pile_elink_x': pile_elink_x, 'brake_elink_x': brake_elink_x, 'fpl_elink_x': fpl_elink_x,
    }


def build_deck_superstructure_module(trestle_ui_dict, span_dict, rib_info, is_double,
                                     _seg_enabled, _panel_length, _comp,
                                     SECTION_name_search_dict, node_id, elem_id):
    """桥面系上部结构模块：纵肋/横肋节点、桥面板PLATE、结构组、虚拟梁、横向联系/桥面系结构组

    弹连点产出：span_rib_elastic_nodes_dict（横肋侧）、rib_long_nodes_dict/span_rib_nodes_dict（纵肋↔横肋）。
    ELASTICLINK 行由 mcb_generate 弹连铺平生成，不在本模块生成。
    """
    NODE_lst, ELEM_lst, GRUP_lst = [], [], []
    rib_long_nodes_dict = {}
    rib_long_nodes_lst = []
    rib_long_elems_lst = []
    seg_boundaries = {}
    if is_double:
        _nid_start = node_id
        rib_long_nodes_dict, rib_long_nodes_lst, _, node_id = make_rib_long_nodes(trestle_ui_dict, rib_info, node_id)

        # 分节建模：在纵肋边界处插入新节点（物理断开）
        if _seg_enabled and _panel_length:
            seg_boundaries, seg_new_nids, node_id = make_rib_long_seg_nodes(
                rib_long_nodes_dict, _panel_length, node_id)
            rib_long_nodes_lst.extend(seg_new_nids)
            print(f"[MCT] 分节建模(物理断开) 面板长{_panel_length}m, 新增节点 {len(seg_new_nids)} 个")

        NODE_lst.extend(make_rib_long_nodes_mct(rib_long_nodes_dict))
        rib_long_elems_lst, rib_long_elems_mctlst, elem_id = make_rib_long_elems(
            trestle_ui_dict, rib_long_nodes_dict, elem_id,
            seg_boundaries=seg_boundaries if _seg_enabled else None)
        ELEM_lst.extend(rib_long_elems_mctlst)
        print(f"[MCT] 纵肋节点 {len(rib_long_nodes_lst)} 个 (node {_nid_start}~{node_id-1})")

    # 横肋节点（单层时同时用于桥面板 PLATE）
    _nid_start = node_id
    rib_trans_nodes_lst, span_rib_nodes_dict, span_rib_elastic_nodes_dict, virtual_nodes_lst, rib_trans_middle_y, node_id = make_rib_trans_nodes(rib_info, span_dict, node_id)
    for nodes_dict in span_rib_nodes_dict.values():
        NODE_lst.extend(make_rib_trans_nodes_mct(nodes_dict))

    # 横肋 BEAM 单元
    rib_elems_lst, span_rib_elems_mctlst, elem_id = make_rib_trans_elems(trestle_ui_dict, span_rib_nodes_dict, elem_id)
    ELEM_lst.extend(span_rib_elems_mctlst)

    # 桥面板 PLATE 单元（单层连横肋间，双层连纵肋间）
    _deck_mat = _mat_id(trestle_ui_dict, _extract_brand(_comp.get('deck_material')))
    if is_double:
        span_plate_elems_lst, span_plate_elems_mctlst, _, elem_id = make_plate_elems_by_long(
            rib_long_nodes_dict, elem_id,
            seg_boundaries=seg_boundaries if _seg_enabled else None,
            mat_id=_deck_mat)
    else:
        span_plate_elems_lst, span_plate_elems_mctlst, elem_id = make_plate_elems(span_rib_nodes_dict, elem_id, mat_id=_deck_mat)
    ELEM_lst.extend(span_plate_elems_mctlst)

    # 结构组：桥面板/纵肋/横肋
    plate_nodes = rib_long_nodes_lst if is_double else rib_trans_nodes_lst
    GRUP_lst.extend(make_plate_group(plate_nodes, span_plate_elems_lst))
    if is_double:
        GRUP_lst.extend(make_rib_long_group(rib_long_nodes_lst, rib_long_elems_lst))
    GRUP_lst.extend(make_rib_trans_group(rib_trans_nodes_lst, rib_elems_lst))

    # 虚拟梁
    virtual_elems_lst, virtual_elems_mctlst, elem_id = make_virtual_elem(SECTION_name_search_dict, virtual_nodes_lst, elem_id)
    ELEM_lst.extend(virtual_elems_mctlst)
    GRUP_lst.extend(make_virtual_group(virtual_nodes_lst, virtual_elems_lst))

    # 横向联系梁结构组（供移动荷载车道定义用）
    lane_nodes = sorted(set(rib_trans_nodes_lst + rib_long_nodes_lst))
    lane_elems = sorted(set(rib_elems_lst + rib_long_elems_lst))
    ns = group_arithmetic_sequences(lane_nodes) if lane_nodes else ""
    es = group_arithmetic_sequences(lane_elems) if lane_elems else ""
    GRUP_lst.append(grup_mctstring_handle('横向联系梁', ns, es))

    # 桥面系结构组（桥面板+横肋+纵肋）
    deck_nodes = sorted(set(rib_trans_nodes_lst + rib_long_nodes_lst))
    deck_elems = sorted(set(span_plate_elems_lst + rib_elems_lst + rib_long_elems_lst))
    ns = group_arithmetic_sequences(deck_nodes) if deck_nodes else ""
    es = group_arithmetic_sequences(deck_elems) if deck_elems else ""
    GRUP_lst.append(grup_mctstring_handle('桥面系', ns, es))

    return {
        'nodes': NODE_lst, 'elems': ELEM_lst, 'groups': GRUP_lst,
        'rib_long_nodes_dict': rib_long_nodes_dict, 'rib_long_nodes_lst': rib_long_nodes_lst,
        'rib_long_elems_lst': rib_long_elems_lst, 'rib_trans_nodes_lst': rib_trans_nodes_lst,
        'span_rib_nodes_dict': span_rib_nodes_dict, 'span_rib_elastic_nodes_dict': span_rib_elastic_nodes_dict,
        'virtual_nodes_lst': virtual_nodes_lst, 'virtual_elems_lst': virtual_elems_lst,
        'span_plate_elems_lst': span_plate_elems_lst, 'seg_boundaries': seg_boundaries,
        'node_id': node_id, 'elem_id': elem_id,
    }


def build_steel_beam_module(trestle_ui_dict, rib_info, fpl_elink_x,
                            SECTION_name_search_dict, sections_dict, _comp,
                            node_id, elem_id, elink_id):
    """型钢栈桥纵梁模块：节点/单元/结构组/桥台约束 + 纵梁侧弹连点

    弹连点产出：beam_rib_elink_node_dict（横肋弹连）、beam_dist_elink_nodes（分配梁弹连）。
    ELASTICLINK 行由 mcb_generate 弹连铺平生成。
    """
    NODE_lst, ELEM_lst, GRUP_lst, CONSTRAINT_lst, BNDR_lst = [], [], [], [], []
    _nid_start = node_id
    steel_sec_ref = _get_component(trestle_ui_dict, 'beam_sec', '')
    h_steel = _get_sec_H(sections_dict, steel_sec_ref, 300)

    steel_node_dict, steel_nodes_lst, steel_rib_elink_node_dict, steel_fpl_elink_node_dict, node_id = \
        make_steel_beam_nodes(rib_info, h_steel, fpl_elink_x, node_id)
    NODE_lst.extend(make_steel_beam_nodes_mct(steel_node_dict))
    print(f"[MCT] 型钢纵梁节点 {len(steel_nodes_lst)} 个 (node {_nid_start}~{node_id-1})")

    _steel_mat_id = _mat_id(trestle_ui_dict, _extract_brand(_comp.get('beam_material')), '1')
    steel_elems_lst, steel_elems_mctlst, elem_id = make_steel_beam_elems(
        steel_node_dict, SECTION_name_search_dict, steel_sec_ref, elem_id, mat_id=_steel_mat_id)
    ELEM_lst.extend(steel_elems_mctlst)
    print(f"[MCT] 型钢纵梁单元 {len(steel_elems_lst)} 个")

    GRUP_lst.extend(make_steel_beam_group(steel_nodes_lst, steel_elems_lst))

    # 桥台反力约束（型钢最外侧节点）
    has_abutment = any(
        t in ('桥台', '重力式桥台', '简易桥台')
        for t in [v['type'] for v in trestle_ui_dict['substructure'].values()]
    )
    if has_abutment:
        BNDR_lst.append("桥台反力, 0")
        support_types = [v['type'] for v in trestle_ui_dict['substructure'].values()]
        is_left_brake = support_types[0] == '制动墩' if support_types else False
        is_right_brake = support_types[-1] == '制动墩' if support_types else False
        left_node_ids = []
        right_node_ids = []
        sorted_span_keys = sorted(steel_node_dict.keys())
        if not is_left_brake:
            for sk in sorted_span_keys:
                for beam_nodes in steel_node_dict[sk].values():
                    if beam_nodes:
                        left_node_ids.append(beam_nodes[0][0])
                if left_node_ids:
                    break
        if not is_right_brake:
            for sk in reversed(sorted_span_keys):
                for beam_nodes in steel_node_dict[sk].values():
                    if beam_nodes:
                        right_node_ids.append(beam_nodes[-1][0])
                if right_node_ids:
                    break
        for nid in left_node_ids + right_node_ids:
            CONSTRAINT_lst.append(constraint_mctstring_handle(nid, name='桥台反力'))
        print(f'[MCT] 桥台反力约束(型钢): {len(left_node_ids + right_node_ids)} 个节点')

    # 分配梁弹连节点（型钢截面中心z；半高经 beam_half_height 传 make_dist_nodes）
    beam_dist_elink_nodes = {}
    for span_dict_inner in steel_fpl_elink_node_dict.values():
        for beam_nodes in span_dict_inner.values():
            for item in beam_nodes:
                nid, coord = item
                beam_dist_elink_nodes[nid] = list(coord)
    dist_beam_half_height = h_steel / 2

    return {
        'nodes': NODE_lst, 'elems': ELEM_lst, 'groups': GRUP_lst, 'frames': [],
        'constraints': CONSTRAINT_lst, 'bndr': BNDR_lst,
        'beam_rib_elink_node_dict': steel_rib_elink_node_dict,
        'beam_dist_elink_nodes': beam_dist_elink_nodes,
        'dist_beam_half_height': dist_beam_half_height,
        'node_id': node_id, 'elem_id': elem_id, 'elink_id': elink_id,
    }


def build_truss_beam_module(trestle_ui_dict, span_dict, rib_info, span_rib_elastic_nodes_dict,
                            fpl_elink_x, SECTION_name_search_dict,
                            node_id, elem_id, elink_id, truss_beam_type):
    """上承式桁架梁栈桥纵梁模块（贝雷桁架，按 truss_beam_type 适配）

    与型钢模块结构完全一致（上下弦杆/竖杆/斜杆、释放梁端约束、弹连、边界均同），
    仅材质/高度/尺寸随纵梁类型变化（TRUSS_BEAM_TYPES，S8 参数化）。
    弹连点产出：beam_rib_elink_node_dict（横肋弹连）、beam_dist_elink_nodes（分配梁弹连）。
    """
    cfg = _truss_beam_cfg(truss_beam_type)
    # 上弦杆中心 Z：rib_trans_z - (chord_height + h_rib_trans)/2（按类型弦杆高计算，替代 321 固定值）
    bailey_z = rib_info["rib_trans_z"] - (cfg["chord_height"] + rib_info["h_rib_trans"]) / 2
    bailey_mat_id = _mat_id(trestle_ui_dict, cfg["material"], '1')
    NODE_lst, ELEM_lst, GRUP_lst, FRAME_lst, CONSTRAINT_lst, BNDR_lst = [], [], [], [], [], []
    _nid_start = node_id
    bailey_shangxiangan_coord_dict, bailey_shangxiangan_nodes_lst, span_bailey_shangxiangan_node_dict, bailey_xiaxiangan_coord_dict, bailey_xiaxiangan_nodes_lst, span_bailey_xiaxiangan_node_dict, span_bailey_shugan_node_dict, bailey_shangxiegan_coord_dict, bailey_xiaxiegan_coord_dict, span_bailey_rib_elink_node_dict, bailey_fpl_elink_node_dict, node_id = make_bailey_nodes(trestle_ui_dict, span_dict, span_rib_elastic_nodes_dict, fpl_elink_x, node_id, bailey_z=bailey_z, cfg=cfg, rib_info=rib_info)
    for nodes_dict in span_bailey_shangxiangan_node_dict.values():
        NODE_lst.extend(make_bailey_shangxiangan_nodes_mct(nodes_dict))
    for nodes_dict in span_bailey_xiaxiangan_node_dict.values():
        NODE_lst.extend(make_bailey_xiaxiangan_nodes_mct(nodes_dict))
    for nodes_dict in span_bailey_shugan_node_dict.values():
        NODE_lst.extend(make_bailey_shugan_nodes_mct(nodes_dict))
    _n_bailey = len(bailey_shangxiangan_nodes_lst) + len(bailey_xiaxiangan_nodes_lst) + sum(len(v) for d in span_bailey_shugan_node_dict.values() for v in d.values())
    print(f"[MCT] 贝雷梁节点 {_n_bailey} 个 (上弦{len(bailey_shangxiangan_nodes_lst)}+下弦{len(bailey_xiaxiangan_nodes_lst)}) (node {_nid_start}~{node_id-1})")

    bailey_shangxiangan_elems_lst, bailey_shangxiangan_elems_mctlst, bailey_shangxiangan_frame_rls_elemlst, elem_id = make_bailey_shangxiangan_elem(SECTION_name_search_dict, bailey_shangxiangan_coord_dict, span_bailey_shangxiangan_node_dict, elem_id, cfg=cfg, bailey_mat_id=bailey_mat_id)
    ELEM_lst.extend(bailey_shangxiangan_elems_mctlst)
    bailey_xiaxiangan_elems_lst, bailey_xiaxiangan_elems_mctlst, bailey_xiaxiangan_frame_rls_elemlst, elem_id = make_bailey_xiaxiangan_elem(SECTION_name_search_dict, bailey_xiaxiangan_coord_dict, span_bailey_xiaxiangan_node_dict, elem_id, cfg=cfg, bailey_mat_id=bailey_mat_id)
    ELEM_lst.extend(bailey_xiaxiangan_elems_mctlst)
    bailey_shugan_nodes_lst, bailey_shugan_elems_lst, bailey_shugan_elems_mctlst, elem_id = make_bailey_shugan_elem(SECTION_name_search_dict, span_bailey_shangxiangan_node_dict, span_bailey_xiaxiangan_node_dict, span_bailey_shugan_node_dict, elem_id, cfg=cfg, bailey_mat_id=bailey_mat_id)
    ELEM_lst.extend(bailey_shugan_elems_mctlst)
    bailey_xiegan_nodes_lst, bailey_xiegan_elems_lst, bailey_xiegan_elems_mctlst, elem_id = make_bailey_xiegan_elem(SECTION_name_search_dict, span_bailey_shangxiangan_node_dict, span_bailey_xiaxiangan_node_dict, span_bailey_shugan_node_dict, bailey_shangxiegan_coord_dict, bailey_xiaxiegan_coord_dict, elem_id, cfg=cfg, bailey_mat_id=bailey_mat_id)
    ELEM_lst.extend(bailey_xiegan_elems_mctlst)

    GRUP_lst.extend(make_bailey_shangxiangan_group(bailey_shangxiangan_nodes_lst, bailey_shangxiangan_elems_lst))
    GRUP_lst.extend(make_bailey_xiaxiangan_group(bailey_xiaxiangan_nodes_lst, bailey_xiaxiangan_elems_lst))
    GRUP_lst.extend(make_bailey_xiangan_group(bailey_shangxiangan_nodes_lst, bailey_shangxiangan_elems_lst, bailey_xiaxiangan_nodes_lst, bailey_xiaxiangan_elems_lst))
    GRUP_lst.extend(make_bailey_shugan_group(bailey_shugan_nodes_lst, bailey_shugan_elems_lst))
    GRUP_lst.extend(make_bailey_xiegan_group(bailey_xiegan_nodes_lst, bailey_xiegan_elems_lst))
    GRUP_lst.extend(make_bailey_group(truss_beam_type, bailey_shangxiangan_nodes_lst, bailey_shangxiangan_elems_lst, bailey_xiaxiangan_nodes_lst, bailey_xiaxiangan_elems_lst, bailey_shugan_nodes_lst, bailey_shugan_elems_lst, bailey_xiegan_nodes_lst, bailey_xiegan_elems_lst))

    FRAME_lst.extend(make_bailey_frame_rls(bailey_shangxiangan_frame_rls_elemlst, bailey_xiaxiangan_frame_rls_elemlst))

    # 桥台反力约束
    has_abutment = any(
        t in ('桥台', '重力式桥台', '简易桥台')
        for t in [v['type'] for v in trestle_ui_dict['substructure'].values()]
    )
    if has_abutment:
        BNDR_lst.append("桥台反力, 0")
        lian_keys = sorted(span_bailey_xiaxiangan_node_dict.keys())
        support_types = [v['type'] for v in trestle_ui_dict['substructure'].values()]
        is_left_brake = support_types[0] == '制动墩' if support_types else False
        is_right_brake = support_types[-1] == '制动墩' if support_types else False
        left_node_ids = []
        if not is_left_brake:
            for lk in lian_keys:
                for v in span_bailey_xiaxiangan_node_dict[lk].values():
                    if v:
                        left_node_ids.append(v[0][0])
                if left_node_ids:
                    break
        right_node_ids = []
        if not is_right_brake:
            for lk in reversed(lian_keys):
                for v in span_bailey_xiaxiangan_node_dict[lk].values():
                    if v:
                        right_node_ids.append(v[-1][0])
                if right_node_ids:
                    break
        for nid in left_node_ids + right_node_ids:
            CONSTRAINT_lst.append(constraint_mctstring_handle(nid, name='桥台反力'))
        print(f'[MCT] 桥台反力约束: {len(left_node_ids + right_node_ids)} 个节点')

    # 分配梁弹连节点（贝雷梁下弦杆）
    beam_dist_elink_nodes = {}
    for n_dict in bailey_fpl_elink_node_dict.values():
        for n_lst in n_dict.values():
            for item in n_lst:
                nid, coord_tuple = item
                beam_dist_elink_nodes[nid] = list(coord_tuple)
    dist_beam_half_height = cfg["chord_height"] / 2  # 下弦杆半截面高（按类型弦杆高）

    return {
        'nodes': NODE_lst, 'elems': ELEM_lst, 'groups': GRUP_lst, 'frames': FRAME_lst,
        'constraints': CONSTRAINT_lst, 'bndr': BNDR_lst,
        'beam_rib_elink_node_dict': span_bailey_rib_elink_node_dict,
        'beam_dist_elink_nodes': beam_dist_elink_nodes,
        'dist_beam_half_height': dist_beam_half_height,
        'node_id': node_id, 'elem_id': elem_id, 'elink_id': elink_id,
    }


def build_beam_module(trestle_ui_dict, span_dict, rib_info, span_rib_elastic_nodes_dict,
                      fpl_elink_x, SECTION_name_search_dict, sections_dict, _comp,
                      node_id, elem_id, elink_id, bridge_type, truss_beam_type):
    """纵梁模块派发：steel → 型钢；truss_upper → 上承式桁架；truss_lower → 下承式桁架（暂不支持）"""
    category = _bridge_category(bridge_type)
    if category == 'steel':
        return build_steel_beam_module(trestle_ui_dict, rib_info, fpl_elink_x,
                                       SECTION_name_search_dict, sections_dict, _comp,
                                       node_id, elem_id, elink_id)
    elif category == 'truss_upper':
        return build_truss_beam_module(trestle_ui_dict, span_dict, rib_info, span_rib_elastic_nodes_dict,
                                       fpl_elink_x, SECTION_name_search_dict,
                                       node_id, elem_id, elink_id, truss_beam_type)
    else:
        raise NotImplementedError("下承式桁架梁栈桥暂不支持：上部结构建模未实现，下部结构可复用")


def build_dist_module(trestle_ui_dict, pile_elink_x, brake_elink_x, beam_dist_elink_nodes,
                      dist_beam_half_height, SECTION_name_search_dict, node_id, elem_id):
    """分配梁模块：横向/纵向分配梁节点、单元、结构组

    弹连点产出：dist_beam_elink（分配梁↔纵梁）、dist_pile_elink（分配梁↔钢管桩）。
    ELASTICLINK 行由 mcb_generate 弹连铺平生成。
    """
    NODE_lst, ELEM_lst, GRUP_lst = [], [], []
    _nid_start = node_id
    (dist_trans_dict, dist_long_dict,
     dist_beam_elink, dist_pile_elink,
     pile_xs, pile_ys, node_id) = make_dist_nodes(
        trestle_ui_dict, pile_elink_x, brake_elink_x,
        beam_dist_elink_nodes, node_id,
        beam_half_height=dist_beam_half_height)
    _n_dist = sum(len(v) for v in dist_trans_dict.values()) + sum(len(v) for v in dist_long_dict.values())
    NODE_lst.extend(make_dist_nodes_mct(dist_trans_dict, dist_long_dict))
    print(f"[MCT] 分配梁节点 {_n_dist} 个 (node {_nid_start}~{node_id-1})")
    dist_elem_dict, dist_nodes_dict, dist_elem_mctlst, elem_id = make_dist_elems(trestle_ui_dict, dist_trans_dict, dist_long_dict,
                        SECTION_name_search_dict, elem_id)
    ELEM_lst.extend(dist_elem_mctlst)
    dist_group_mctlst = make_dist_group(dist_elem_dict, dist_nodes_dict)
    GRUP_lst.extend(dist_group_mctlst)

    return {
        'nodes': NODE_lst, 'elems': ELEM_lst, 'groups': GRUP_lst,
        'dist_trans_dict': dist_trans_dict, 'dist_long_dict': dist_long_dict,
        'dist_beam_elink': dist_beam_elink, 'dist_pile_elink': dist_pile_elink,
        'pile_xs': pile_xs, 'pile_ys': pile_ys,
        'node_id': node_id, 'elem_id': elem_id,
    }


def build_pile_module(trestle_ui_dict, dist_pile_elink, pile_xs, pile_elink_x, brake_elink_x,
                      SECTION_name_search_dict, node_id, elem_id):
    """钢管桩模块：节点、单元、结构组、桩底边界/土弹簧

    弹连点产出：pile_head_nodes（桩顶弹连点）。ELASTICLINK 行由 mcb_generate 弹连铺平生成。
    """
    NODE_lst, ELEM_lst, GRUP_lst, CONSTRAINT_lst, BNDR_lst = [], [], [], [], []
    SPRING_lst = []
    _nid_start = node_id
    (pile_head_nodes, pile_bottom_nodes, pile_water_top_nodes,
     pile_water_force_nodes, pile_brace_conn_nodes, pile_spring_nodes,
     pile_all_node_ids, pile_by_support, node_id) = make_pile_nodes(
        trestle_ui_dict, dist_pile_elink, pile_xs, pile_elink_x, brake_elink_x, node_id)
    print(f"[MCT] 钢管桩节点 {len(pile_all_node_ids)} 个 (含土弹簧 {len(pile_spring_nodes)} 个) (node {_nid_start}~{node_id-1})")

    NODE_lst.extend(make_pile_nodes_mct(pile_head_nodes, pile_bottom_nodes,
                                         pile_water_top_nodes, pile_water_force_nodes,
                                         pile_brace_conn_nodes, pile_spring_nodes))
    (pile_elem_dict, pile_elem_mctlst,
     pile_bottom_ids, elem_id) = make_pile_elems(
        trestle_ui_dict, pile_by_support, SECTION_name_search_dict, elem_id,
        pile_head_nodes, pile_bottom_nodes, pile_water_top_nodes,
        pile_water_force_nodes, pile_brace_conn_nodes, pile_spring_nodes)
    ELEM_lst.extend(pile_elem_mctlst)

    # 桩底边界条件：桩底固结/土弹簧
    conn = trestle_ui_dict.get('connection', {})
    soil_spring_setting = conn.get("土弹簧", "/")
    bnd_method = "土弹簧" if isinstance(soil_spring_setting, dict) else "桩底边界"
    soil_thickness = soil_spring_setting.get("thickness", "1.0") if isinstance(soil_spring_setting, dict) else "1.0"
    print(f"[MCT] 桩边界: {bnd_method}" + (f", 分层厚度={soil_thickness}m" if bnd_method == "土弹簧" else ""))

    sub_data = trestle_ui_dict.get('substructure', {})
    for sup_key in sorted(pile_by_support, key=int):
        p = sub_data.get(sup_key, {})
        ground_y = p.get('ground_elevation', 0)
        soil_data = p.get('soil_data', {})
        print(f"  桩 {sup_key}: 地面标高={ground_y}m, 土层={len(soil_data)} 层")

    if bnd_method == "桩底边界":
        pile_bnd_code = conn.get("桩底", "111111")
        for nid in pile_bottom_ids:
            CONSTRAINT_lst.append(constraint_mctstring_handle(nid, constraint=pile_bnd_code, name="桩底反力"))
        print(f"[MCT] 桩底边界({pile_bnd_code}): {len(pile_bottom_ids)} 个节点")
    else:
        all_coords = {}
        for d in (pile_head_nodes, pile_bottom_nodes, pile_water_top_nodes,
                  pile_water_force_nodes, pile_brace_conn_nodes, pile_spring_nodes):
            all_coords.update(d)
        SPRING_lst = make_soil_springs(
            trestle_ui_dict, pile_by_support, all_coords)
        BNDR_lst.append("土弹簧, 0")
        for nid in pile_bottom_ids:
            CONSTRAINT_lst.append(constraint_mctstring_handle(nid, constraint="001000", name="桩底反力"))
        print(f"[MCT] 土弹簧: {len(SPRING_lst)} 个, 桩底z向支承: {len(pile_bottom_ids)} 个")

    GRUP_lst.extend(make_pile_group(pile_elem_dict, pile_all_node_ids, pile_by_support))

    return {
        'nodes': NODE_lst, 'elems': ELEM_lst, 'groups': GRUP_lst,
        'constraints': CONSTRAINT_lst, 'bndr': BNDR_lst, 'springs': SPRING_lst,
        'pile_head_nodes': pile_head_nodes, 'pile_bottom_nodes': pile_bottom_nodes,
        'pile_water_top_nodes': pile_water_top_nodes, 'pile_water_force_nodes': pile_water_force_nodes,
        'pile_brace_conn_nodes': pile_brace_conn_nodes, 'pile_spring_nodes': pile_spring_nodes,
        'pile_all_node_ids': pile_all_node_ids, 'pile_by_support': pile_by_support,
        'pile_bottom_ids': pile_bottom_ids,
        'node_id': node_id, 'elem_id': elem_id,
    }


def build_brace_module(trestle_ui_dict, pile_head_nodes, pile_brace_conn_nodes, pile_by_support,
                       SECTION_name_search_dict, node_id, elem_id):
    """联结系模块：中心节点、XX竖杆节点、单元、结构组"""
    NODE_lst, ELEM_lst, GRUP_lst = [], [], []
    _nid_start = node_id
    brace_center_nodes, brace_groups, brace_all_node_ids, brace_by_support, node_id = make_brace_nodes(trestle_ui_dict, pile_head_nodes, pile_brace_conn_nodes, pile_by_support, node_id)
    NODE_lst.extend(make_brace_nodes_mct(brace_center_nodes))
    (brace_elem_dict, brace_elem_mctlst, elem_id, brace_xx_node_mctlst, node_id) = make_brace_elems(
        trestle_ui_dict, brace_groups, brace_by_support,
        SECTION_name_search_dict, elem_id, node_id)
    NODE_lst.extend(brace_xx_node_mctlst)
    print(f"[MCT] 联结系节点 中心{len(brace_center_nodes)}+XX竖杆{len(brace_xx_node_mctlst)} (node {_nid_start}~{node_id-1})")
    for line in brace_xx_node_mctlst:
        nid = int(line.split(',')[0].strip())
        brace_all_node_ids.append(nid)
    ELEM_lst.extend(brace_elem_mctlst)
    brace_group_mctlst = make_brace_group(brace_elem_dict, brace_all_node_ids, brace_by_support)
    GRUP_lst.extend(brace_group_mctlst)
    print(f"[MCT] 联结系节点 {len(brace_all_node_ids)} 个, 单元 {len(brace_elem_mctlst)} 个")

    return {'nodes': NODE_lst, 'elems': ELEM_lst, 'groups': GRUP_lst,
            'node_id': node_id, 'elem_id': elem_id}


def build_load_module(trestle_ui_dict, pile_head_nodes, pile_water_force_nodes, pile_bottom_nodes,
                      pile_by_support, virtual_elems_lst):
    """荷载模块：风/水流/移动/静载/面荷载/荷载组合"""
    wind_load_mctlst = make_wind_load(trestle_ui_dict, pile_head_nodes)
    water_load_mctlst = make_water_load(trestle_ui_dict, pile_water_force_nodes, pile_bottom_nodes, pile_by_support)
    print(f"[MCT] 环境荷载: 风 {len(wind_load_mctlst)} 行, 水流 {len(water_load_mctlst)} 行")

    move_load_lst, move_load_mctlst = build_move_load(
        trestle_ui_dict, virtual_elem_ids=virtual_elems_lst)
    stldcase_info_dict = make_static_load_cases_info(trestle_ui_dict)
    stldcase_mctstring = make_static_load_cases(stldcase_info_dict)
    pnloadtype_mctstring = make_pnloadtype(stldcase_info_dict)
    planeload_mctstring = make_vehicle_static_load(trestle_ui_dict, stldcase_info_dict)
    load_comb_lines = load_combinations(trestle_ui_dict)

    return {
        'wind_load_mctlst': wind_load_mctlst, 'water_load_mctlst': water_load_mctlst,
        'move_load_lst': move_load_lst, 'move_load_mctlst': move_load_mctlst,
        'stldcase': list(stldcase_mctstring), 'pnloadtype': list(pnloadtype_mctstring),
        'planeload': list(planeload_mctstring), 'load_comb': load_comb_lines,
    }


def assemble_mct_output(node_id, NODE_lst, ELEM_lst, GRUP_lst, BNDR_lst, ELASTICLINK_lst,
                        FRAME_lst, CONSTRAINT_lst, SPRING_lst, STLDCASE_lst, PNLOADTYPE_lst, PLANELOAD_lst,
                        MATERIAL_lst, SECTION_lst, THICKNESS_lst,
                        move_load_lst, move_load_mctlst, wind_load_mctlst, water_load_mctlst,
                        load_comb_lines, mct_savepath):
    """汇总输出：节点排序验证 + mct 汇总表 + 写文件/剪贴板"""
    # 按节点号排序，确保 Midas Civil 节点输入区域有序（各 make_*_nodes_mct 函数
    # 按构件类型分组输出，但 node_id 是跨类型递增的，导致 NODE_lst 顺序错乱）
    _node_ids_in_lst = [int(s.split(',')[0]) for s in NODE_lst]
    _expected = set(range(1, node_id))
    _actual = set(_node_ids_in_lst)
    _missing = sorted(_expected - _actual)
    _dupes = sorted(set(nid for nid in _node_ids_in_lst if _node_ids_in_lst.count(nid) > 1))
    if _missing:
        print(f"[警告] NODE_lst 缺失 {len(_missing)} 个节点号: {_missing[:30]}{'...' if len(_missing)>30 else ''}")
        print(f"  node_id 终值={node_id}, NODE_lst 条目={len(NODE_lst)}")
    if _dupes:
        print(f"[警告] NODE_lst 有 {len(_dupes)} 个重复节点号: {_dupes[:20]}")
    if not _missing and not _dupes:
        print(f"[MCT] 节点验证通过: {len(NODE_lst)} 个节点, node_id 终值={node_id}")
    NODE_lst.sort(key=lambda s: int(s.split(',')[0]))

    mct_lst = [
        # 版本号，进行各个版本Midas的mct文件测试
        # ["*VERSION", "8.6.5"],
        # 单位系
        ["*UNIT", "N,MM,KJ,C"],
        # 材料
        ["*MATERIAL"] + MATERIAL_lst,
        # 截面
        ["*SECTION"] + SECTION_lst,
        # 板厚
        ["*THICKNESS"] + THICKNESS_lst,
        # 节点
        ["*NODE"] + NODE_lst,
        # 单元
        ["*ELEMENT"] + ELEM_lst,
        # # 结构组
        ["*GROUP"] + GRUP_lst,
        # # 约束组
        ["*BNDR-GROUP", "弹性连接, 0", "桩底反力, 0", "释放梁端约束, 0"] + BNDR_lst,
        # 弹性连接
        ["*ELASTICLINK"] + ELASTICLINK_lst,
        # 释放梁端约束
        ["*FRAME-RLS"] + FRAME_lst,
        # 桩底固结 / 土弹簧
        ["*CONSTRAINT"] + CONSTRAINT_lst,
        ["*SPRING"] + SPRING_lst if SPRING_lst else [],
        # 荷载组
        ["*LOAD-GROUP", "自重", "风荷载"] + move_load_lst,
        # 荷载工况
        ["*STLDCASE", "自重, USER,", "风荷载, USER,", "水流力, USER,"] + STLDCASE_lst,
        # 自重
        ["*USE-STLD, 自重", "*SELFWEIGHT", "0, 0, {}, 自重".format(-1)],
        # 风荷载
        ["*USE-STLD, 风荷载", "*CONLOAD"] + wind_load_mctlst,
        # 水流力
        ["*USE-STLD, 水流力", "*CONLOAD"] + water_load_mctlst,
        # 移动荷载
        ["*MVLDCODE", "  CODE=CHINA"] + move_load_mctlst,
        # 分配面荷载类型
        ["*PNLOADTYPE"] + PNLOADTYPE_lst,
        # 分配面荷载加载
        ["*PLANELOAD"] + PLANELOAD_lst,
        # 荷载组合
        ["*LOADCOMB"] + load_comb_lines
    ]
    print("[MCT] 汇总完成，正在输出...")
    mct_lst_tolines = "\n".join([item for sublist in mct_lst for item in sublist])
    SET_CLIP_STRING(mct_lst_tolines)
    print("[MCT] 已复制到剪贴板")
    write_file_within_path(mct_savepath, mct_lst_tolines, "gbk")
    MidasURL()
    mcb_savepath = file_extension_Modified(mct_savepath, '.mcb')
    importmct_to_mcb(mct_savepath, mcb_savepath)


# 生成全部的mct命令流
def mcb_generate(trestle_ui_dict, mct_savepath = None):
    # 起始节点号、单元号
    node_id = 1
    elem_id = 1
    elink_id = 1
    NODE_lst = []
    ELEM_lst = []
    GRUP_lst = []
    FRAME_lst = []
    PLATE_lst = []
    THICKNESS_lst = []
    ELASTICLINK_lst = []
    CONSTRAINT_lst = []
    BNDR_lst = []
    SPRING_lst = []
    STLDCASE_lst = []
    PNLOADTYPE_lst = []
    PLANELOAD_lst = []

    # 弹连刚度
    conn = trestle_ui_dict.get('connection', {})
    stiff_beam_dist = _conn_stiff(conn.get("纵梁与横向分配梁"))
    stiff_dist_pile   = _conn_stiff(conn.get("桩顶分配梁与桩顶"))
    stiff_deck_beam = _conn_stiff(conn.get("桥面系横肋与纵梁"))

    # ── 1、截面/材质/板厚 + 跨径/几何/弹连点 ──
    o1 = build_section_material_module(trestle_ui_dict)
    SECTION_lst = o1['SECTION_lst']; MATERIAL_lst = o1['MATERIAL_lst']
    THICKNESS_lst = o1['THICKNESS_lst']; SECTION_name_search_dict = o1['SECTION_name_search_dict']
    sections_dict = o1['sections_dict']
    bridge_type = o1['bridge_type']
    _comp = o1['comp']
    o2 = build_geometry_module(trestle_ui_dict)
    span_dict = o2['span_dict']; rib_info = o2['rib_info']; is_double = o2['is_double']
    _seg_enabled = o2['seg_enabled']; _panel_length = o2['panel_length']
    pile_elink_x = o2['pile_elink_x']; brake_elink_x = o2['brake_elink_x']; fpl_elink_x = o2['fpl_elink_x']

    # ── 2、桥面系上部结构（纵肋/横肋/桥面板/结构组/虚拟梁）──
    o3 = build_deck_superstructure_module(trestle_ui_dict, span_dict, rib_info, is_double,
                                          _seg_enabled, _panel_length, _comp,
                                          SECTION_name_search_dict, node_id, elem_id)
    NODE_lst.extend(o3['nodes']); ELEM_lst.extend(o3['elems']); GRUP_lst.extend(o3['groups'])
    rib_long_nodes_dict = o3['rib_long_nodes_dict']
    span_rib_nodes_dict = o3['span_rib_nodes_dict']
    virtual_elems_lst = o3['virtual_elems_lst']
    node_id = o3['node_id']; elem_id = o3['elem_id']

    # ── 3、纵梁（型钢/桁架等模块派发）──
    truss_beam_type = _get_component(trestle_ui_dict, 'truss_beam_type', '321型贝雷梁')
    o4 = build_beam_module(trestle_ui_dict, span_dict, rib_info, o3['span_rib_elastic_nodes_dict'],
                           fpl_elink_x, SECTION_name_search_dict, sections_dict, _comp,
                           node_id, elem_id, elink_id, bridge_type, truss_beam_type)
    NODE_lst.extend(o4['nodes']); ELEM_lst.extend(o4['elems']); GRUP_lst.extend(o4['groups'])
    FRAME_lst.extend(o4['frames']); CONSTRAINT_lst.extend(o4['constraints']); BNDR_lst.extend(o4['bndr'])
    beam_rib_elink_node_dict = o4['beam_rib_elink_node_dict']
    beam_dist_elink_nodes = o4['beam_dist_elink_nodes']
    dist_beam_half_height = o4['dist_beam_half_height']
    node_id = o4['node_id']; elem_id = o4['elem_id']; elink_id = o4['elink_id']

    # ── 4、分配梁 / 钢管桩 / 联结系 / 荷载 ──
    o5 = build_dist_module(trestle_ui_dict, pile_elink_x, brake_elink_x, beam_dist_elink_nodes,
                           dist_beam_half_height, SECTION_name_search_dict, node_id, elem_id)
    NODE_lst.extend(o5['nodes']); ELEM_lst.extend(o5['elems']); GRUP_lst.extend(o5['groups'])
    dist_trans_dict = o5['dist_trans_dict']; dist_long_dict = o5['dist_long_dict']
    dist_beam_elink = o5['dist_beam_elink']; dist_pile_elink = o5['dist_pile_elink']
    pile_xs = o5['pile_xs']; pile_ys = o5['pile_ys']
    node_id = o5['node_id']; elem_id = o5['elem_id']

    o6 = build_pile_module(trestle_ui_dict, dist_pile_elink, pile_xs, pile_elink_x, brake_elink_x,
                           SECTION_name_search_dict, node_id, elem_id)
    NODE_lst.extend(o6['nodes']); ELEM_lst.extend(o6['elems']); GRUP_lst.extend(o6['groups'])
    CONSTRAINT_lst.extend(o6['constraints']); BNDR_lst.extend(o6['bndr'])
    SPRING_lst.extend(o6['springs'])
    pile_head_nodes = o6['pile_head_nodes']; pile_bottom_nodes = o6['pile_bottom_nodes']
    pile_water_force_nodes = o6['pile_water_force_nodes']
    pile_brace_conn_nodes = o6['pile_brace_conn_nodes']; 
    pile_by_support = o6['pile_by_support']
    node_id = o6['node_id']; elem_id = o6['elem_id']

    o7 = build_brace_module(trestle_ui_dict, pile_head_nodes, pile_brace_conn_nodes, pile_by_support,
                            SECTION_name_search_dict, node_id, elem_id)
    NODE_lst.extend(o7['nodes']); ELEM_lst.extend(o7['elems']); GRUP_lst.extend(o7['groups'])
    node_id = o7['node_id']; elem_id = o7['elem_id']

    o8 = build_load_module(trestle_ui_dict, pile_head_nodes, pile_water_force_nodes, pile_bottom_nodes,
                           pile_by_support, virtual_elems_lst)
    wind_load_mctlst = o8['wind_load_mctlst']; water_load_mctlst = o8['water_load_mctlst']
    move_load_lst = o8['move_load_lst']; move_load_mctlst = o8['move_load_mctlst']
    STLDCASE_lst.extend(o8['stldcase']); PNLOADTYPE_lst.extend(o8['pnloadtype'])
    PLANELOAD_lst.extend(o8['planeload']); load_comb_lines = o8['load_comb']

    # ── 5、弹连生成：单独铺平（各弹连刚度/名称/节点对零散，不封装）──
    # 纵肋↔横肋（双层）
    if is_double and rib_long_nodes_dict:
        rib_long_trans_elink_mctlst, elink_id = make_rib_long_trans_elink(
            rib_long_nodes_dict, span_rib_nodes_dict, elink_id, stiffness=stiff_deck_beam)
        ELASTICLINK_lst.extend(rib_long_trans_elink_mctlst)
    # 横肋↔纵梁
    rib_beam_elink_mctlst, elink_id = make_rib_trans_beam_elink(
        o3['span_rib_elastic_nodes_dict'], beam_rib_elink_node_dict, elink_id, stiffness=stiff_deck_beam)
    ELASTICLINK_lst.extend(rib_beam_elink_mctlst)
    # 纵梁↔分配梁
    elink_lines, elink_id = build_elastic_links(
        beam_dist_elink_nodes, dist_beam_elink, elink_id,
        name="弹性连接", stiffness=stiff_beam_dist)
    ELASTICLINK_lst.extend(elink_lines)
    # 横向分配梁↔纵向分配梁（制动墩，复用桩顶分配梁与桩顶刚度）
    if dist_long_dict:
        trans_all = {}
        for nodes in dist_trans_dict.values():
            for nid, coord in nodes:
                trans_all[nid] = coord
        long_all = {}
        for nodes in dist_long_dict.values():
            for nid, coord in nodes:
                long_all[nid] = coord
        elink_lines, elink_id = build_elastic_links(
            long_all, trans_all, elink_id,
            name="弹性连接", stiffness=stiff_dist_pile)
        ELASTICLINK_lst.extend(elink_lines)
    # 分配梁↔钢管桩
    elink_lines, elink_id = build_elastic_links(
        dist_pile_elink, pile_head_nodes, elink_id,
        name="弹性连接", stiffness=stiff_dist_pile)
    ELASTICLINK_lst.extend(elink_lines)

    # ── 6、模块：节点排序验证 + 汇总输出 ──
    assemble_mct_output(node_id, NODE_lst, ELEM_lst, GRUP_lst, BNDR_lst, ELASTICLINK_lst,
                        FRAME_lst, CONSTRAINT_lst, SPRING_lst, STLDCASE_lst, PNLOADTYPE_lst, PLANELOAD_lst,
                        MATERIAL_lst, SECTION_lst, THICKNESS_lst,
                        move_load_lst, move_load_mctlst, wind_load_mctlst, water_load_mctlst,
                        load_comb_lines, mct_savepath)
