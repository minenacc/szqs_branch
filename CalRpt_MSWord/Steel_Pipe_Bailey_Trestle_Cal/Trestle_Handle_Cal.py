"""
栈桥计算书 - 计算和Word子模板渲染模块
包含：荷载计算、上部结构计算、稳定性计算、Word子模板渲染等
"""

# 1. 标准库
import os
import sys
import math
import re
import tempfile
from pathlib import Path

# 2. 第三方库
from docxtpl import DocxTemplate, InlineImage
from docx import Document
from docx.shared import Cm
from collections import Counter

# 3. 本地模块
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from General.Formula import *
from General.Midas import MidasAPI, get_midas_version, Value_Deformed_ByNodes
from General.DataUtils import inear_interpolation
from General.TextHandle import format_scientific_latex
from General.WordHandle import replace_placeholder_with_formula
from Trestle_Post_Cal import (
    Post_processing_rib_long,Post_processing_rib_trans, Post_processing_bailey,
    Post_processing_steel_beam,
    Post_processing_fenpeiliang, Post_processing_lianjiexi, Post_processing_gangguanzhuang,
    Post_processing_reaction,
    Picture_Beamstress
)


# ========================= 辅助函数 ===============================
def if_Boundary_Group_name_exist(group_name):
    """检查边界组是否存在"""
    try:
        result = MidasAPI("GET", "/db/BNGR")
        print(result)
        if result and 'BNGR' in result:
            for key, value in result['BNGR'].items():
                if value.get('NAME') == group_name:
                    return True
    except:
        pass
    return False

def _classify_brace_angle(elem_data, NODE_dict):
    """按单元两端节点坐标判断：弦杆(Z相同) / 竖杆(XY相同Z不同) / 斜杆(其余)"""
    nids = elem_data.get('NODE', [])
    if len(nids) < 2:
        return '弦杆'
    n1 = NODE_dict.get('NODE', {}).get(str(nids[0]), {})
    n2 = NODE_dict.get('NODE', {}).get(str(nids[1]), {})
    dx = n2.get('X', 0) - n1.get('X', 0)
    dy = n2.get('Y', 0) - n1.get('Y', 0)
    dz = n2.get('Z', 0) - n1.get('Z', 0)
    if abs(dz) < 1e-6:
        return '弦杆'
    elif abs(dx) < 1e-6 and abs(dy) < 1e-6:
        return '竖杆'
    return '斜杆'

def _read_sect_mat(sect_id, matl_id, grp_name, SECT_dict, MATL_dict, THIK_dict):
    """读取截面名称、SHAPE、vSIZE、材料名称和类型"""
    sect_name = ''
    vsize = []
    shape = ''
    if sect_id is not None:
        if grp_name == '桥面板':
            thik_data = THIK_dict.get('THIK', {}).get(str(sect_id), {})
            sect_name = thik_data.get('NAME', '')
            vsize = [thik_data.get('T_IN', 0)]
        else:
            sect_data = SECT_dict.get('SECT', {}).get(str(sect_id), {})
            sect_name = sect_data.get('SECT_NAME', '')
            vsize = sect_data.get('SECT_BEFORE', {}).get('SECT_I', {}).get('vSIZE', [])
            shape = str(sect_data.get('SECT_BEFORE', {}).get('SHAPE', '') or '').strip()
    matl_name = ''
    matl_type = ''
    if matl_id is not None:
        matl_data = MATL_dict.get('MATL', {}).get(str(matl_id), {})
        matl_name = matl_data.get('NAME', '')
        matl_type = matl_data.get('TYPE', '')
    return {'sect_name': sect_name, 'matl_name': matl_name, 'matl_type': matl_type, 'vsize': vsize, 'shape': shape}


def _get_kh_coefficient(D, ground_category):
    category = ground_category.split(":")[0].strip().upper()
    table_data = {
        "Z": [5, 10, 15, 20, 30, 40, 50, 60, 70, 80, 90, 100, 150, 200, 250, 300, 350, 400, 450],
        "A": [1.08, 1.17, 1.23, 1.28, 1.34, 1.39, 1.42, 1.46, 1.48, 1.51, 1.53, 1.55, 1.62, 1.68, 1.73, 1.77, 1.77, 1.77, 1.77],
        "B": [1.00, 1.00, 1.07, 1.12, 1.19, 1.25, 1.29, 1.33, 1.36, 1.40, 1.42, 1.45, 1.54, 1.62, 1.67, 1.72, 1.77, 1.77, 1.77],
        "C": [0.86, 0.86, 0.86, 0.92, 1.00, 1.06, 1.12, 1.16, 1.20, 1.24, 1.27, 1.30, 1.42, 1.52, 1.59, 1.66, 1.71, 1.77, 1.77],
        "D": [0.79, 0.79, 0.79, 0.79, 0.85, 0.85, 0.91, 0.96, 1.01, 1.05, 1.09, 1.13, 1.27, 1.39, 1.48, 1.57, 1.64, 1.71, 1.77]
    }
    z_list = table_data["Z"]
    v_list = table_data[category]

    # 边界条件处理 (超出表格范围时，取边缘最值)
    if D <= z_list[0]:
        return v_list[0]
    if D >= z_list[-1]:
        return v_list[-1]
    # 查找区间并进行计算
    for i in range(len(z_list) - 1):
        if z_list[i] <= D <= z_list[i+1]:
            # 命中确切值，直接返回，避免不必要的插值计算
            if D == z_list[i]:
                return v_list[i]
            if D == z_list[i+1]:
                return v_list[i+1]
            # 处于区间内，插值函数
            # x1, x2 对应高度 Z; h1, h2 对应系数 kh; p 对应目标高度 D
            result = inear_interpolation(
                x1=z_list[i], x2=z_list[i+1], 
                h1=v_list[i], h2=v_list[i+1], 
                p=D
            )
            return round(result, 2)

def _get_kf_coefficient(wind_v):
    if wind_v <= 24.5:
        kf = 1.0
    elif wind_v > 24.5 and wind_v <= 32.6:
        kf = 1.02
    else:
        kf = 1.05
    return kf
 
def _get_a0_kc_coefficient(ground_category):
    ground_category = ground_category.split(":")[0].strip().upper()
    if ground_category == 'A':
        return 0.12, 1.174
    elif ground_category == 'B':
        return 0.16, 1.0
    elif ground_category == 'C':
        return 0.22, 0.785
    elif ground_category == 'D':
        return 0.30, 0.564
    else:
        return 0.12, 1.174

def _chinese_to_arabic(chinese_str):
    """中文数字转阿拉伯数字（支持一~九十九）"""
    digit_map = {'零': 0, '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9}
    if not chinese_str:
        return 0
    result = 0
    temp = 0
    for ch in chinese_str:
        if ch == '十':
            if temp == 0:
                temp = 1
            result += temp * 10
            temp = 0
        elif ch in digit_map:
            temp = digit_map[ch]
    result += temp
    return result if result > 0 else temp


def _get_concrete_f(grade):
    """根据混凝土强度等级返回 (fc, ft, Ec)，单位 N/mm²。
    fc — 轴心抗压强度设计值
    ft — 轴心抗拉强度设计值
    Ec — 弹性模量
    """
    table = {
        'C15': (7.2,  0.91, 22000),
        'C20': (9.6,  1.10, 25500),
        'C25': (11.9, 1.27, 28000),
        'C30': (14.3, 1.43, 30000),
        'C35': (16.7, 1.57, 31500),
        'C40': (19.1, 1.71, 32500),
        'C45': (21.1, 1.80, 33500),
        'C50': (23.1, 1.89, 34500),
        'C55': (25.3, 1.96, 35500),
        'C60': (27.5, 2.04, 36000),
        'C65': (29.7, 2.09, 36500),
        'C70': (31.8, 2.14, 37000),
        'C75': (33.8, 2.18, 37500),
        'C80': (35.9, 2.22, 38000),
    }
    return table.get(grade, ('', '', ''))

def _get_rebar_f(grade):
    """根据钢筋牌号返回 (fy, Es)，单位 N/mm²。
    fy — 抗拉强度设计值
    Es — 弹性模量
    """
    table = {
        'HPB300':  (270, 210000),
        'HRB335':  (300, 200000),
        'HRB400':  (360, 200000),
        'HRBF400': (360, 200000),
        'RRB400':  (360, 200000),
        'HRB500':  (435, 200000),
        'HRBF500': (435, 200000),
    }
    return table.get(grade, ('', ''))


# ========================= 材质通用库查询 ================================

def _lookup_steel(brand, steel_dict):
    """从 steel_dict 查找钢材牌号参数，返回完整 entry 或 None"""
    for spec_data in steel_dict.values():
        if brand in spec_data.get("牌号参数", {}):
            return spec_data["牌号参数"][brand]
    return None

def _pick_from_thickness_table(thickness_table, thickness):
    """从分档表 {16: val, 40: val, 63: val, 80: val, 100: val} 中
    按实际厚度取对应档位值。超出范围取边界值。"""
    if not thickness_table:
        return 0
    sorted_keys = sorted(thickness_table.keys())
    for k in sorted_keys:
        if thickness <= k:
            return thickness_table[k]
    return thickness_table[sorted_keys[-1]]

def _lookup_allowable_stress(brand, flange_thickness, web_thickness, steel_dict):
    """查找钢材容许应力（按厚度分档，组合应力取翼缘厚，剪应力取腹板厚）。

    参数:
        brand:            牌号（如 "Q235"）
        flange_thickness: 翼缘厚 tf1 (mm)，用于查容许组合应力 f
        web_thickness:    腹板厚 tw (mm)，用于查容许剪应力 fv
        steel_dict:       材质库字典

    返回:
        {'f': float, 'fv': float, 'E': float} — 牌号不在库中时返回 Q235 t≤16 兜底值
    """
    entry = _lookup_steel(brand, steel_dict)
    if not entry:
        # Q235 t≤16 兜底（GB50017-2017 表3.4.1-1）
        return {'f': 215.0, 'fv': 125.0, 'E': 206000.0}
    return {
        'f':  _pick_from_thickness_table(entry.get('f', {}), flange_thickness),
        'fv': _pick_from_thickness_table(entry.get('fv', {}), web_thickness),
        'E':  entry.get('E', 206000),
    }


def _get_flange_web_t(vsize, shape):
    """按 SHAPE 从 vSIZE 分拆 (flange_t, web_t) mm。

    SHAPE：P=圆管, H=工字钢, C=槽钢, 2C=双槽钢, B=箱型(含双拼I型), SR=虚拟梁。
    未知 SHAPE 回退 max(vsize)。
    """
    if not vsize:
        return 0, 0
    v = [float(x) for x in vsize]
    s = (shape or '').upper().strip()
    if s == 'P':           # 圆管: vSIZE = [D, t, ...]
        t = v[1] if len(v) > 1 else 0
        return t, t
    if s in ('H', 'C', '2C', 'B'):  # H型/槽钢/双槽钢/箱型: vSIZE = [H, B, tw, tf1, ...]
        if len(v) >= 4:
            return v[3], v[2]  # tf1, tw
    # SR/板型/未知: 单值取 v[0]
    t = v[0] if v else 0
    return t, t

def _lookup_concrete(grade, concrete_dict):
    """从 concrete_dict 查找混凝土等级参数，返回 {fc, ft, Ec, ...} 或 None"""
    for spec_data in concrete_dict.values():
        if grade in spec_data.get("强度等级参数", {}):
            return spec_data["强度等级参数"][grade]
    return None

def _lookup_rebar(grade, rebar_dict):
    """从 rebar_dict 查找钢筋牌号参数，返回 {fy, E, ...} 或 None"""
    for spec_data in rebar_dict.values():
        if grade in spec_data.get("牌号参数", {}):
            return spec_data["牌号参数"][grade]
    return None


def _get_sec_mat_by_ele(ele_id, ELEM_dict, SECT_dict, MATL_dict):
    """根据单元号查找截面名、材料名、vSIZE、SHAPE"""
    elem_data = ELEM_dict.get('ELEM', {}).get(str(ele_id), {})
    sect_id = elem_data.get('SECT')
    matl_id = elem_data.get('MATL')
    sect_name = ''
    vsize = []
    shape = ''
    if sect_id is not None:
        sect_data = SECT_dict.get('SECT', {}).get(str(sect_id), {})
        sect_name = sect_data.get('SECT_NAME', '')
        vsize = sect_data.get('SECT_BEFORE', {}).get('SECT_I', {}).get('vSIZE', [])
        shape = str(sect_data.get('SECT_BEFORE', {}).get('SHAPE', '') or '').strip()
    matl_name = ''
    if matl_id is not None:
        matl_name = MATL_dict.get('MATL', {}).get(str(matl_id), {}).get('NAME', '')
    return {'sect_name': sect_name, 'matl_name': matl_name, 'vsize': vsize, 'shape': shape}

def _get_group_support_coords(group_name, GRUP_dict_res, ELEM_dict_res, NODE_dict_res, coord_axis):
    """从结构组中提取支撑坐标列表。
    coord_axis: 'X' 或 'Y'，表示取节点的哪个方向坐标。
    返回: (sorted_coords, coord_to_nodes)
        sorted_coords:   去重排序后的坐标列表
        coord_to_nodes:  {coord: [node_id, ...]} 坐标到节点ID列表的映射
    """
    coord_to_nodes = {}
    for grp in GRUP_dict_res.get('GRUP', {}).values():
        if grp.get('NAME', '') != group_name:
            continue
        for eid in grp.get('E_LIST', []):
            elem_data = ELEM_dict_res.get('ELEM', {}).get(str(eid), {})
            for nid in elem_data.get('NODE', []):
                node = NODE_dict_res.get('NODE', {}).get(str(nid), {})
                val = node.get(coord_axis)
                if val is not None:
                    coord_to_nodes.setdefault(val, []).append(str(nid))
        break
    return sorted(coord_to_nodes.keys()), coord_to_nodes

def _find_span_length(target_coord, support_coords, coord_min=None, coord_max=None):
    """在 support_coords 中找到 target_coord 两侧最近的支撑，返回间距。
    若两侧均有支撑: L = 两侧最近支撑点距离。
    若仅一侧有支撑（悬臂）: L = 该侧最近支撑点与构件边缘的距离。
    若无支撑点返回 0。

    参数:
        target_coord:   变形最大点的坐标
        support_coords: 支撑点坐标列表
        coord_min:      构件坐标最小值（悬臂时确定边缘）
        coord_max:      构件坐标最大值（悬臂时确定边缘）

    返回:
        (L, is_cantilever)
    """
    lower = [c for c in support_coords if c < target_coord]
    upper = [c for c in support_coords if c > target_coord]

    if lower and upper:
        # target 在两个支撑点之间 → L = 两侧最近支撑点距离
        return min(upper) - max(lower), False

    # target 恰好落在支撑点上 → 向两侧外扩找最近支撑
    if not lower and not upper and target_coord in support_coords:
        strictly_below = [c for c in support_coords if c < target_coord]
        strictly_above = [c for c in support_coords if c > target_coord]
        if strictly_below and strictly_above:
            return min(strictly_above) - max(strictly_below), False
        lower = strictly_below
        upper = strictly_above

    if lower:   # 仅有左侧支撑 → 悬臂右端
        edge = coord_max if coord_max is not None else target_coord
        return edge - max(lower), True
    if upper:   # 仅有右侧支撑 → 悬臂左端
        edge = coord_min if coord_min is not None else target_coord
        return min(upper) - edge, True

    return 0, False


def _compute_relative_deformation(max_Dz, max_Dz_node, support_coords, coord_to_nodes, target_coord, load_case_name, NODE_dict_res=None, match_axis=None, match_val=None):
    """计算相对变形: 找两侧最近支撑节点，查变形表，取差值最大的。

    max_Dz:          构件最大变形值
    max_Dz_node:     最大变形节点号
    support_coords:  支撑点坐标排序列表
    coord_to_nodes:  {coord: [node_id, ...]} 映射
    target_coord:    最大变形节点的坐标
    load_case_name:  荷载组合名
    NODE_dict_res:   节点字典（match_axis 过滤时需要）
    match_axis:      垂直方向轴名 'X' 或 'Y'（用于过滤同排支撑节点）
    match_val:       垂直方向坐标值（与构件最大变形节点一致）
    """
    lower = [c for c in support_coords if c < target_coord]
    upper = [c for c in support_coords if c > target_coord]

    def _pick_node(coord):
        """从 coord 对应的支撑节点中，取 match_axis 方向上最接近 match_val 的那个。"""
        nids = coord_to_nodes.get(coord, [])
        if not match_axis or match_val is None or not NODE_dict_res:
            return nids[0] if nids else None
        best, best_dist = None, float('inf')
        for nid in nids:
            node = NODE_dict_res.get('NODE', {}).get(str(nid), {})
            v = node.get(match_axis)
            if v is not None and abs(v - match_val) < best_dist:
                best, best_dist = nid, abs(v - match_val)
        return best

    query_nids = []
    if lower:
        nid = _pick_node(max(lower))
        if nid:
            query_nids.append(nid)
    if upper:
        nid = _pick_node(min(upper))
        if nid:
            query_nids.append(nid)
    print(query_nids)

    if not query_nids:
        print(f"[_compute_relative_deformation] 未找到支撑节点，使用绝对变形 {max_Dz}")
        return max_Dz

    # 逐个节点查变形
    candidates = []
    for nid in query_nids:
        res = Value_Deformed_ByNodes("node_displacement", int(nid), load_case_name)
        data = res.get('node_displacement', {}).get('DATA', [])
        print(data)
        if data:
            dz = float(data[0][-1])
            print(dz)
            candidates.append(dz)
            print(f"  支撑节点{nid} Dz={dz}")

    if not candidates:
        print(f"[_compute_relative_deformation] 支撑节点无变形数据，使用绝对变形 {max_Dz}")
        return max_Dz

    worst_support_dz = max(candidates, key=lambda v: abs(max_Dz - v))
    relative_dz = abs(max_Dz - worst_support_dz)
    print(f"[_compute_relative_deformation] 节点{max_Dz_node} Dz={max_Dz}, "
          f"支撑Dz={worst_support_dz}, 相对变形={relative_dz}")
    return relative_dz


def _rib_space_cal(GRUP_dict_res, NODE_dict_res, rib_type, cal_axis):
    coords = []
    for group_data in GRUP_dict_res.get('GRUP', {}).values():
        if group_data.get('NAME') == rib_type:
            rib_node_ids = group_data.get('N_LIST', [])
            break
    for nid in rib_node_ids:
        node_data = NODE_dict_res.get('NODE', {}).get(str(nid), {})
        if 'Y' in node_data:
            coords.append(node_data[cal_axis])
    spacings = [round(coords[i+1] - coords[i], 1) for i in range(len(coords)-1)]
    spacings = [s for s in spacings if s > 0]
    spacing = Counter(spacings).most_common(1)[0][0]
    return spacing


# ========================= 基础计算函数 ================================
# 输入参数计算基础抗剪V、抗冲切F、抗弯M承载力及配筋
# 剪切计算主函数
def Vs_Verification(x, y, z, as0, at, bt, D, pmax, ft):
    # 是否需要抗剪验算
    Vs_lst = foundation_Vs(x, y, at, bt, D, pmax)
    # 抗剪切力
    Allowable_Vs_lst = foundation_Allowable_Vs(x, y, z, as0, ft)
    # 需要抗剪验算
    if Vs_lst != None:
        # print("基础底面短边尺寸小于等于柱宽加两倍基础有效高度，需要抗剪计算")
        # 剪切力
        As1, As2, Vs1, Vs2 = Vs_lst
        # 抗剪切力
        Bhs, A01, A02, Allowable_Vs1, Allowable_Vs2 = Allowable_Vs_lst
        print("剪切力: ", Vs1, Vs2)
        print("抗剪力: ", Allowable_Vs1, Allowable_Vs2)
        # 是否包络
        if Vs1 <= Allowable_Vs1:
            if_Vs_satisfied1 = "满足"
        else:
            if_Vs_satisfied1 = "不满足"
        if Vs2 <= Allowable_Vs2:
            if_Vs_satisfied2 = "满足"
        else:
            if_Vs_satisfied2 = "不满足"
    # else:
    #     print("基础底面短边尺寸大于柱宽加两倍基础有效高度，不需要抗剪计算")
    return [As1, As2, Vs1, Vs2, Bhs, A01, A02, Allowable_Vs1, Allowable_Vs2, if_Vs_satisfied1, if_Vs_satisfied2]


# 冲切计算主函数
def Fl_Verification(x, y, z, as0, at, bt, D, pmax, ft):
    # 抗冲切力
    h0, Bhp, am, Allowable_Fl = foundation_Allowable_Fl(ft, at, bt, z, D, as0)
    print("抗冲切力: ", Allowable_Fl)
    # 冲切力
    Al1, Al2, Fl1, Fl2 = foundation_Fl(pmax, at, bt, D, x, y, z, as0)
    print("冲切力: ", Fl1, Fl2)
    # 是否包络
    if Fl1 < Allowable_Fl:
        if_Fl_satisfied1 = "满足"
    else:
        if_Fl_satisfied1 = "不满足"
    if Fl2 < Allowable_Fl:
        if_Fl_satisfied2 = "满足"
    else:
        if_Fl_satisfied2 = "不满足"
    return [Al1, Al2, Fl1, Fl2, h0, Bhp, am, Allowable_Fl, if_Fl_satisfied1, if_Fl_satisfied2]


# 扩大基础弯矩及配筋计算主函数
def M_Verification(x, y, z, as0, at, bt, D, p, pmax, pmin, fy, steel_s1, steel_s2, steel_d1, steel_d2, type_M):
    # 弯矩计算
    a1, aa, bb = cal_M_param(x, y, at, bt, D, type_M)
    print(a1, aa, bb)
    M1, M2 = cal_M(x, y, a1, aa, bb, p, pmax, pmin, type_M)
    print(M1, M2)
    # 配筋面积计算
    As1 = max([cal_As(M1, fy, x, y, z, as0, type_M), min_p_As(z, as0)])
    As2 = max([cal_As(M2, fy, x, y, z, as0, type_M), min_p_As(z, as0)])
    print("弯矩作用方向最小配筋面积: ", As1)# 弯矩方向
    print("垂直弯矩方向最小配筋面积: ", As2)# 垂直弯矩方向

    # 钢筋直径和间距计算
    rebar_arrangement_lst1 = cal_rebar_arrangement(z, as0, steel_s1, steel_s2, steel_d1, steel_d2, As1)
    rebar_arrangement_lst2 = cal_rebar_arrangement(z, as0, steel_s1, steel_s2, steel_d1, steel_d2, As2)
    if rebar_arrangement_lst1 != None:
        s1, d1, p1, A1 = rebar_arrangement_lst1
        print(f"弯矩作用方向配筋规格: {d1}@{s1}mm", ", ", f"配筋率{round(p1*100,3)}% > 0.15%", ", ", f"钢筋截面面积{round(A1,1)}mm2 > {As1}mm2")
    else:
        s1, d1, p1, A1 = 0, 0, 0, 0
        print(f"弯矩作用方向配筋未成功")
    if rebar_arrangement_lst2 != None:
        s2, d2, p2, A2 = rebar_arrangement_lst2
        print(f"垂直弯矩方向配筋规格: {d2}@{s2}mm", ", ", f"配筋率{round(p2*100,3)}% > 0.15%", ", ", f"钢筋截面面积{round(A2,1)}mm2 > {As2}mm2")
    else:
        s2, d2, p2, A2 = 0, 0, 0, 0
        print(f"垂直弯矩方向配筋未成功")
    
    if A1 >= As1:
        if_As_satisfied1 = "满足"
    else: 
        if_As_satisfied1 = "不满足"
    if s1 == 200:
        if_gouzao1 = "是"
    else:
        if_gouzao1 = "否"
    if A2 >= As2:
        if_As_satisfied2 = "满足"
    else: 
        if_As_satisfied2 = "不满足"
    if s2 == 200:
        if_gouzao2 = "是"
    else:
        if_gouzao2 = "否"

    return [a1, aa, bb, M1, M2, As1, As2, s1, d1, round(p1*100,3), round(A1,1), s2, d2, round(p2*100,3), round(A2,1), if_As_satisfied1, if_As_satisfied2, if_gouzao1, if_gouzao2]



# ========================= 处理函数 ================================
# 对Midas模型进行结果的读取并截图
def Post_processing_Midas_Model(folder_path, NODE_dict_res, ELEM_dict_res, SECT_dict_res, MATL_dict_res, GRUP_dict_res):
    print("正在获取模型计算结果...")
    # 后处理——纵肋（如有）；无纵肋时为 None（model_cal_params 按 deck_type 判断跳过）
    list0 = None
    for grp in GRUP_dict_res['GRUP'].values():
        if "纵肋" in grp['NAME']:
            print("纵肋信息读取中...")
            list0 = Post_processing_rib_long(folder_path)
    # 后处理——横肋
    print("横肋信息读取中...")
    list1 = Post_processing_rib_trans(folder_path)
    # 后处理——纵梁（按结构组派发）
    _grp_names = {grp['NAME'] for grp in GRUP_dict_res.get('GRUP', {}).values()}
    if '型钢纵梁' in _grp_names:
        print("型钢纵梁信息读取中...")
        list2 = Post_processing_steel_beam(folder_path)
    elif '321型贝雷梁' in _grp_names:
        print("贝雷梁信息读取中...")
        list2 = Post_processing_bailey(folder_path)
    # elif '上弦杆' in _grp_names:  # 下承式预留
    #     list2 = Post_processing_truss_lower(folder_path)
    else:
        print("未识别到纵梁结构组，跳过纵梁后处理")
        list2 = [0, 0, 0, 0, '']
    # 后处理——分配梁
    print("分配梁信息读取中...")
    list3 = Post_processing_fenpeiliang(folder_path)
    # 后处理——联结系
    print("联结系信息读取中...")
    list4 = Post_processing_lianjiexi(folder_path)
    # 后处理——钢管桩
    print("钢管桩信息读取中...")
    list5 = Post_processing_gangguanzhuang(folder_path, NODE_dict_res, ELEM_dict_res, SECT_dict_res, MATL_dict_res, GRUP_dict_res)
    # 后处理——反力
    print("反力信息读取中...")
    list6 = Post_processing_reaction(folder_path)
    # 组成字典
    model_post_dict = {"model_post_rib_long":list0,
                       "model_post_rib_trans":list1,
                       "model_post_beam":list2,
                       "model_post_fenpeiliang":list3,
                       "model_post_lianjiexi":list4,
                       "model_post_gangguanzhuang":list5,
                       "model_post_reaction":list6,
                    }
    print("模型信息读取完成")
    print(model_post_dict)
    return model_post_dict

# 各构件信息读取
def get_component_info(GRUP_dict_res, NODE_dict_res, ELEM_dict_res, SECT_dict_res, MATL_dict_res, THIK_dict_res):
    """
    从各结构组集中读取截面和材料信息。
    桥面板/纵肋/横肋：按名称精确匹配，取任一单元。
    钢管桩/分配梁/联结系：按 'N-名称' 模式匹配，联结系逐单元按角度分类。
    返回:
        component_info: dict，键为结构组名或分类名，
                        值为 {'sect_name', 'matl_name', 'matl_type', 'vsize'}
                        含 has_brace、has_rib_long 两个布尔键
    """
    component_info = {}
    has_brace = False
    has_rib_long = False
    has_steel_beam = False
    has_bailey = False

    # 已知分配梁类型（按长度降序避免子串误匹配）
    _distrib_types = sorted(['横向分配梁', '纵向分配梁'], key=len, reverse=True)

    def _parse_group_name(grp_name):
        """从 '2-横向分配梁1' 提取 ('横向分配梁', 2)，从 '1-联结系' 提取 ('联结系', 1)"""
        m = re.match(r'^(\d+)-(.+)$', grp_name)
        if not m:
            return None, None
        num = int(m.group(1))
        raw = m.group(2)
        base = re.sub(r'\d+$', '', raw)
        for dtype in _distrib_types:
            if base.startswith(dtype):
                return dtype, num
        return base, num

    for grp in GRUP_dict_res.get('GRUP', {}).values():
        grp_name = grp.get('NAME', '')
        elem_ids = grp.get('E_LIST') or []

        # ---- 固定名称结构组：桥面板 / 纵肋 / 横肋 ----
        if grp_name in ('桥面板', '纵肋', '横肋', '型钢纵梁'):
            if grp_name == '纵肋':
                has_rib_long = True
            if grp_name == '型钢纵梁':
                has_steel_beam = True
            sample_eid = elem_ids[0] if elem_ids else None
            if sample_eid is None:
                continue
            elem_data = ELEM_dict_res.get('ELEM', {}).get(str(sample_eid), {})
            component_info[grp_name] = _read_sect_mat(
                elem_data.get('SECT'), elem_data.get('MATL'),
                grp_name, SECT_dict_res, MATL_dict_res, THIK_dict_res
            )
            continue
        # ---- 桁架梁结构组判断 ----
        if grp_name in ('321型贝雷梁',):
            has_bailey = True
            continue

        # ---- 带桩号的结构组：钢管桩 / 分配梁 / 联结系 ----
        base_type, pile_num = _parse_group_name(grp_name)
        if base_type is None:
            continue

        if base_type == '联结系':
            has_brace = True
            # 弦杆(水平)和竖杆/斜杆(竖直或倾斜)各取一个即可
            found_horiz = False
            found_non_horiz = False
            for eid in elem_ids:
                if found_horiz and found_non_horiz:
                    break
                elem_data = ELEM_dict_res.get('ELEM', {}).get(str(eid), {})
                brace_type = _classify_brace_angle(elem_data, NODE_dict_res)
                if brace_type == '弦杆':
                    if found_horiz:
                        continue
                    found_horiz = True
                    key = f"{pile_num}-联结系弦杆"
                else:
                    if found_non_horiz:
                        continue
                    found_non_horiz = True
                    key = f"{pile_num}-联结系竖杆/斜杆"
                component_info[key] = _read_sect_mat(
                    elem_data.get('SECT'), elem_data.get('MATL'),
                    '联结系', SECT_dict_res, MATL_dict_res, THIK_dict_res
                )
        elif base_type == '钢管桩':
            key = f"{pile_num}-钢管桩"
            if key not in component_info and elem_ids:
                elem_data = ELEM_dict_res.get('ELEM', {}).get(str(elem_ids[0]), {})
                info = _read_sect_mat(
                    elem_data.get('SECT'), elem_data.get('MATL'),
                    '钢管桩', SECT_dict_res, MATL_dict_res, THIK_dict_res
                )
                # 桩长 + 节点ID列表：遍历结构组全部单元
                node_z_set = set()
                node_id_set = set()
                for eid in elem_ids:
                    ed = ELEM_dict_res.get('ELEM', {}).get(str(eid), {})
                    for nid in ed.get('NODE', []):
                        node_id_set.add(str(nid))
                        z = NODE_dict_res.get('NODE', {}).get(str(nid), {}).get('Z')
                        if z is not None:
                            node_z_set.add(z)
                info['node_ids'] = node_id_set
                if node_z_set:
                    info['L'] = round(abs(max(node_z_set) - min(node_z_set)), 3)
                    info['pile_bottom_level'] = round(min(node_z_set), 3)
                else:
                    info['L'] = 0
                    info['pile_bottom_level'] = 0
                component_info[key] = info
        else:
            # 分配梁：N-横向分配梁 / N-纵向分配梁
            key = f"{pile_num}-{base_type}"
            if key not in component_info and elem_ids:
                elem_data = ELEM_dict_res.get('ELEM', {}).get(str(elem_ids[0]), {})
                component_info[key] = _read_sect_mat(
                    elem_data.get('SECT'), elem_data.get('MATL'),
                    '分配梁', SECT_dict_res, MATL_dict_res, THIK_dict_res
                )

    # ---- 判断纵梁类型 ----
    if has_steel_beam:
        beam_type = '型钢纵梁'
    elif has_bailey:
        beam_type = '贝雷梁'
    else:
        beam_type = '贝雷梁'  # 兜底
    component_info['has_brace'] = has_brace
    component_info['has_rib_long'] = has_rib_long
    component_info['beam_type'] = beam_type
    return component_info


# 模型参数读取
def get_model_params(component_info, GRUP_dict_res, NODE_dict_res, MATL_dict_res,
                     steel_dict=None, concrete_dict=None):
    """
    从模型中读取计算参数。
    component_info 由主函数调用 get_component_info 获取后传入。
    steel_dict / concrete_dict: 材质通用库（来自 load_material_library），可选
    """
    has_rib_long = component_info.get('has_rib_long', False)
    beam_type = component_info.get('beam_type', '贝雷梁')

    # 1、自重系数
    self_weight_param = MidasAPI("GET","/db/BODF")
    if self_weight_param == {'message': ''} or 'error' in self_weight_param:
        self_weight = "No Self_Weight Defined"
    else:
        self_weight = self_weight_param['BODF']['1']['FV'][2]
    print(f"自重系数: {self_weight}")

    # 2、纵梁排数
    if beam_type == '型钢纵梁':
        beam_node_ids = []
        for group_data in GRUP_dict_res.get('GRUP', {}).values():
            if group_data.get('NAME') == '型钢纵梁':
                beam_node_ids = group_data.get('N_LIST', [])
                break
        beam_y_coords = []
        for nid in beam_node_ids:
            node_data = NODE_dict_res.get('NODE', {}).get(str(nid), {})
            if 'Y' in node_data:
                beam_y_coords.append(node_data['Y'])
        beam_y_coords.sort()
        # 按Y坐标分组
        beam_y_groups = []
        for y in beam_y_coords:
            if not beam_y_groups or abs(y - beam_y_groups[-1]) > 0.5:
                beam_y_groups.append(y)
        beam_n = len(beam_y_groups)
    elif beam_type == '贝雷梁':
    # 从结构组"贝雷梁竖杆"的节点Y坐标分组计数
        bailey_shugan_node_ids = []
        for group_data in GRUP_dict_res.get('GRUP', {}).values():
            if group_data.get('NAME') == '贝雷梁竖杆':
                bailey_shugan_node_ids = group_data.get('N_LIST', [])
                break
        bailey_y_coords = []
        for nid in bailey_shugan_node_ids:
            node_data = NODE_dict_res.get('NODE', {}).get(str(nid), {})
            if 'Y' in node_data:
                bailey_y_coords.append(node_data['Y'])
        bailey_y_coords.sort()
        # 按Y坐标分组
        bailey_y_groups = []
        for y in bailey_y_coords:
            if not bailey_y_groups or abs(y - bailey_y_groups[-1]) > 0.5:
                bailey_y_groups.append(y)
        beam_n = len(bailey_y_groups)
    print(f"纵梁排数: {beam_n}")

    # 3、钢管桩直径列表：从全部钢管桩结构组的 vSIZE 取直径
    pile_d_set = set()
    for key, info in component_info.items():
        if key.endswith('钢管桩') and isinstance(info, dict):
            vsize = info.get('vsize', [])
            if vsize and vsize[0]:
                pile_d_set.add(int(vsize[0]))
    pile_d_lst = sorted(pile_d_set)
    print(f"钢管桩直径列表: {pile_d_lst}")

    # 4、桥面系类型：由 component_info 中桥面板的材料类型 + 是否有纵肋判断
    deck_matl_type = component_info.get('桥面板', {}).get('matl_type', '')
    if deck_matl_type == 'CONC':
        deck_type = '混凝土桥面板'
    elif deck_matl_type == 'STEEL':
        deck_type = '双层钢面板' if has_rib_long else '单层钢面板'
    else:
        deck_type = '单层钢面板'
    print(f"桥面系类型: {deck_type}")

    # 5、材料参数（优先从材质通用库查询，兜底用硬编码函数）
    steel_lst = []
    concrete_lst = []
    rebar_lst = []
    for matl_data in MATL_dict_res.get('MATL', {}).values():
        matl_name = matl_data.get('NAME', '')
        matl_type = matl_data.get('TYPE', '')
        if matl_type == 'STEEL' and matl_name != '16Mn':
            entry = _lookup_steel(matl_name, steel_dict) if steel_dict else None
            if entry:
                f_table = entry.get('f', {})
                fv_table = entry.get('fv', {})
                steel_lst.append({
                    'grade': matl_name,
                    'f_1': f_table.get(16, 0),
                    'fv_1': fv_table.get(16, 0),
                    'f_2': f_table.get(40, 0),
                    'fv_2': fv_table.get(40, 0),
                    'E': entry.get('E', 206000),
                })
            else:
                # Q235 兜底（材质库查不到时，GB50017-2017 表3.4.1-1）：t≤16: f=215/fv=125, 16<t≤40: f=205/fv=120
                steel_lst.append({
                    'grade': matl_name,
                    'f_1': 215, 'fv_1': 125,
                    'f_2': 205, 'fv_2': 120,
                    'E': 206000,
                })
        elif matl_type == 'CONC':
            conc_entry = _lookup_concrete(matl_name, concrete_dict) if concrete_dict else None
            if conc_entry:
                concrete_lst.append({
                    'grade': matl_name,
                    'fc': conc_entry.get('fc', 0),
                    'ft': conc_entry.get('ft', 0),
                    'E': conc_entry.get('Ec', 0),
                })
            else:
                fc, ft, ec = _get_concrete_f(matl_name)
                concrete_lst.append({
                    'grade': matl_name, 'fc': fc, 'ft': ft, 'E': ec,
                })
        elif matl_type == 'REBAR':
            rebar_lst.append({
                'grade': matl_name,
                'fy': 0, 'E': 0,
            })
    material = {
        'steel_lst': steel_lst,
        'concrete_lst': concrete_lst,
        'rebar_lst': rebar_lst,
        'has_concrete': len(concrete_lst) > 0,
        'has_rebar': len(rebar_lst) > 0,
    }
    print("材料参数:", material)

    # 6、获取midas版本号
    midas_version = get_midas_version()
    print(f"midas版本号: {midas_version}")

    return self_weight, beam_n, pile_d_lst, deck_type, material, midas_version


def get_foundation_material(foundation_name, foundation_value_dict, material,
                           concrete_dict=None, rebar_dict=None):
    """从 UI 二级窗口参数中提取基础材料等级，合并模型材料属性。

    参数:
        foundation_name:      基础类型名称（"打入桩"/"钻孔灌注桩"/"扩大基础"/"条形基础"）
        foundation_value_dict: UI二级窗口参数字典（键为基础类型名，值为参数列表）
        material:             get_model_params 返回的 material 字典
        concrete_dict:        材质通用库-混凝土（来自 load_material_library），可选
        rebar_dict:           材质通用库-钢筋（来自 load_material_library），可选

    返回:
        dict，包含基础材料等级和属性：
            concrete_grade: 混凝土等级（str）
            concrete_fc:    混凝土抗压强度设计值（MPa）
            concrete_ft:    混凝土抗拉强度设计值（MPa）
            concrete_Ec:    混凝土弹性模量（MPa）
            rebar_grade:    钢筋等级（str）
            rebar_fy:       钢筋抗拉强度设计值（MPa）
            rebar_Es:       钢筋弹性模量（MPa）
            steel_lst:      钢材列表（来自模型）
    """
    # 默认值
    result = {
        'concrete_grade': '',
        'concrete_fc': 0,
        'concrete_ft': 0,
        'concrete_Ec': 0,
        'rebar_grade': '',
        'rebar_fy': 0,
        'rebar_Es': 0,
    }

    # 从 foundation_value_dict 提取材料等级（UI 已改为 dict 存储）
    values = foundation_value_dict.get(foundation_name, {})
    conc_grade = ''
    rebar_grade = ''

    if foundation_name == '钻孔灌注桩' and isinstance(values, dict):
        conc_grade = str(values.get('concrete_grade', ''))
        rebar_grade = str(values.get('rebar_grade', ''))
    elif foundation_name in ('扩大基础', '条形基础') and isinstance(values, dict):
        conc_grade = str(values.get('conc_grd', ''))
        rebar_grade = str(values.get('rebar_grd', ''))
    # 打入桩：无混凝土/钢筋，钢材来自模型

    # 混凝土参数：优先从材质通用库查询，兜底用硬编码函数
    if conc_grade:
        conc_entry = _lookup_concrete(conc_grade, concrete_dict) if concrete_dict else None
        if conc_entry:
            result['concrete_grade'] = conc_grade
            result['concrete_fc'] = conc_entry.get('fc', 0)
            result['concrete_ft'] = conc_entry.get('ft', 0)
            result['concrete_Ec'] = conc_entry.get('Ec', 0)
        else:
            # 模型中未找到，使用硬编码查询
            result['concrete_grade'] = conc_grade
            fc, ft, Ec = concrete_f(conc_grade)
            result['concrete_fc'] = fc
            result['concrete_ft'] = ft
            result['concrete_Ec'] = Ec

    # 钢筋参数：优先从材质通用库查询，兜底用硬编码函数
    if rebar_grade:
        rebar_entry = _lookup_rebar(rebar_grade, rebar_dict) if rebar_dict else None
        if rebar_entry:
            result['rebar_grade'] = rebar_grade
            result['rebar_fy'] = rebar_entry.get('fy', 0)
            result['rebar_Es'] = rebar_entry.get('E', 0)
        else:
            result['rebar_grade'] = rebar_grade
            fy, Es = rebar_f(rebar_grade)
            result['rebar_fy'] = fy
            result['rebar_Es'] = Es

    # 钢材列表直接从模型 material 传入
    result['steel_lst'] = material.get('steel_lst', [])

    print("[get_foundation_material] 基础材料:", result)
    return result


def environment_params(excel_params, beam_n, pile_d_lst, beam_type='贝雷梁', component_info=None, bridge_type='上承式桁架梁栈桥'):
    # 1、风荷载参数
    wind_standard = str(excel_params.get('wind_standard', ''))
    wind_formula = str(excel_params.get('wind_formula', ''))
    ground_category = str(excel_params.get('ground_class', ''))
    D = float(excel_params.get('ref_height', ''))
    Gv = float(excel_params.get('terrain_factor', ''))
    U10 = float(excel_params.get('design_wind_speed', ''))
    kt = float(excel_params.get('length_factor', ''))
    kf = _get_kf_coefficient(U10)
    kh = _get_kh_coefficient(float(D), ground_category)
    
    # 获取型钢高度（型钢时使用）
    steel_H = 0
    if beam_type == '型钢纵梁' and component_info:
        beam_vsize = component_info.get('型钢纵梁', {}).get('vsize', [])
        if len(beam_vsize) > 1:
            steel_H = float(beam_vsize[0])  # H = vsize[0]（翼缘高度）

    # 计算迎风面积 A（贝雷梁用透风系数，型钢用翼缘宽度）
    if beam_type == '贝雷梁':
        A = round((1-0.66 ** int(beam_n))*0.3*4.5/(1-0.66),2)
    else:
        # 型钢：A = n * B * L（B为翼缘宽度，L为计算长度4.5m）
        A = round(beam_n * steel_H / 1000 * 4.5, 2)  # B转为米
    
    if wind_standard == '《公路桥梁抗风设计规范》':
        wind = {
            'wind_standard': wind_standard,
            'wind_formula':wind_formula,
            'wind_v': str(excel_params.get('design_wind_speed', '')),
            'ground_category': ground_category,
            'D': str(round(D, 2)),
            'Gv': str(round(Gv, 2)),
            'kf': str(round(kf, 2)),
            'kt': str(round(kt, 2)),
            'kh': str(kh),
            'beam_n': str(beam_n),
            'beam_type': beam_type,
            'bridge_type': bridge_type,
        }
        if wind_formula == 'Ud = kf·kt·kh·U10':
            Ud = round(kf*kt*kh*U10,2)
            Ug = round(Gv*Ud,2)
            
            if beam_type == '贝雷梁':
                # 贝雷梁：FgD = 0.5 * ρ * Ug^2 * CH，CH=1.7
                ch = 1.7
                wind['ch'] = ch
                FgD = round(0.5*1.25*Ug**2*ch,2)
                F = round(FgD*A, 2)
                wind['formula_registry'] = {
                    '{Ud_Formula_3360}': JTG_3360_01_2018_4_2_6_2 + f'={str(kf)}' + r'\times' + str(kt) + r'\times' + str(kh) + r'\times' + f'{U10}={Ud}m/s',
                    '{Ug_Formula_3360}': JTG_3360_01_2018_5_2_1 + f'={Gv}' + r'\times' + f'{Ud}={Ug}m/s',
                    '{wind_FgD_cal}': JTG_3360_01_2018_5_3_1 + r'=0.5\times1.25' + r'\times{' + str(Ug) + r'}^2\times' + f'{ch}={FgD}kPa',
                    '{wind_A_cal}': r'A=\frac{1-{0.66}^' + str(beam_n) + r'}{1-0.66}\times0.3\times4.5=' + f'{A}m^2',
                    '{wind_F_cal}':r'F=\frac{F_g}{D}A=' + str(FgD) + r'\times' + str(round(A,2)) + r'=' + f'{F}kN'
                }
            else:
                # 型钢：F_g = 0.5 * ρ * V_g^2 * C_H * D，D=型钢翼缘高度H
                ch = 1.3
                wind['ch'] = ch
                D_m = steel_H / 1000  # H转为米
                Fg = round(0.5*1.25*Ug**2*ch*D_m/1000, 2)  # F_g (kN/m)
                wind['formula_registry'] = {
                    '{Ud_Formula_3360}': JTG_3360_01_2018_4_2_6_2 + f'={str(kf)}' + r'\times' + str(kt) + r'\times' + str(kh) + r'\times' + f'{U10}={Ud}m/s',
                    '{Ug_Formula_3360}': JTG_3360_01_2018_5_2_1 + f'={Gv}' + r'\times' + f'{Ud}={Ug}m/s',
                    '{wind_F_cal}': r'F_{g}=\frac{1}{2}\rho V_{g}^{2}C_{H}D=' + r'=0.5\times1.25' + r'\times{' + str(Ug) + r'}^2\times' + str(ch) + r'\times' + f'{D_m}={Fg}kN/m'
                }
                
        elif wind_formula == 'Ud = kf·(Z/10)^α0·Us10':
            a0, kc = _get_a0_kc_coefficient(ground_category)
            Us10 = round(kc * U10, 2)
            Ud = round(kf * (D / 10) ** 0.5 * U10, 2)
            Ug = round(Gv * Ud, 2)
            
            if beam_type == '贝雷梁':
                ch = 1.7
                FgD = round(0.5 * 1.25 * Ug ** 2 * ch, 2)
                F = round(FgD * A, 2)
                wind['ch'] = ch
                wind['kc'] = kc
                wind['formula_registry'] = {
                    '{Uc_Formula_3360}': JTG_3360_01_2018_4_2_4 + f'={kc}' + r'\times' + f'{U10}={Us10}m/s',
                    '{Ud_Formula_3360}': JTG_3360_01_2018_4_2_6_1 + f'={kf}' +r'\times' + str(D/10) + '^{' + str(a0) + r'}\times' + f'{Us10}={Ud}m/s',
                    '{Ug_Formula_3360}': JTG_3360_01_2018_5_2_1 + f'={Gv}' + r'\times' + f'{Ud}={Ug}m/s',
                    '{wind_FgD_cal}': JTG_3360_01_2018_5_3_1 + r'=0.5\times1.25' + r'\times{' + str(Ug) + r'}^2\times' + f'{ch}={FgD}kPa',
                    '{wind_A_cal}': r'A=\frac{1-{0.66}^' + str(beam_n) + r'}{1-0.66}\times0.3\times4.5=' + f'{A}m^2',
                    '{wind_F_cal}':r'F=\frac{F_g}{D}A=' + str(FgD) + r'\times' + str(round(A,2)) + r'=' + f'{F}kN'
                }
            else:
                # 型钢
                ch = 1.3
                D_m = steel_H / 1000
                Fg = round(0.5*1.25*Ug**2*ch*D_m/1000, 2)
                wind['ch'] = ch
                wind['kc'] = kc
                wind['formula_registry'] = {
                    '{Uc_Formula_3360}': JTG_3360_01_2018_4_2_4 + f'={kc}' + r'\times' + f'{U10}={Us10}m/s',
                    '{Ud_Formula_3360}': JTG_3360_01_2018_4_2_6_1 + f'={kf}' +r'\times' + str(D/10) + '^{' + str(a0) + r'}\times' + f'{Us10}={Ud}m/s',
                    '{Ug_Formula_3360}': JTG_3360_01_2018_5_2_1 + f'={Gv}' + r'\times' + f'{Ud}={Ug}m/s',
                    '{wind_F_cal}': JTG_3360_01_2018_5_3_1 + r'=0.5\times1.25' + r'\times{' + str(Ug) + r'}^2\times' + str(ch) + r'\times' + f'{D_m}={Fg}kN/m'
                }

    # 2、水流力参数
    water_standard = str(excel_params.get('flow_standard', ''))
    water_v = float(excel_params.get('design_flow_speed', ''))
    p = 1.0
    water = {
        'water_standard': water_standard,
        'water_v': str(round(water_v,2)),
        'pile_d_lst': pile_d_lst,
        'formula_registry': {},
        }
    for d in pile_d_lst:
        d_m = d / 1000
        Fw = round(0.5 * 0.73 * p * water_v ** 2 * d_m, 2)
        water['formula_registry'][f'water_force_cal_{d}'] =  JTS_144_1_2010_13_0_1 + r'=0.5\times0.73\times' + str(p) + r'\times{' + str(water_v) + r'}^2\times' + str(d) + r'\times1.0=' + f'{Fw}kN'
        

    # 3、土层参数
    soil_param = excel_params.get('soil_param', [])

    return wind, water, soil_param


def get_conditions_info(COMP_dict_res):
    """从荷载组合数据中解析工况信息。

    参数:
        COMP_dict_res: MidasAPI("GET", "/db/LCOM-GEN") 返回的荷载组合字典

    返回:
        (conditions_table, conditions)
        conditions_table: list[dict]，每项含 text_index, type, param, factor
        conditions:       list[dict]，每项含 text_index, param（+ 连接字符串）
    """
    conditions_table = []
    lcom = COMP_dict_res.get('LCOM-GEN', {})

    for comb_data in lcom.values():
        name = comb_data.get('NAME', '')
        vcomb = comb_data.get('vCOMB', [])

        # 仅解析"工况X标准组合"/"工况X基本组合"（名称含"工况"）
        if '工况' not in name:
            continue

        if '标准' in name:
            comb_type = '标准'
        elif '基本' in name:
            comb_type = '基本'
        else:
            continue

        # 提取工况编号（"工况"与"标准"/"基本"之间的中文数字）
        m = re.search(r'工况(.+?)(?:标准|基本)', name)
        if not m:
            continue
        text_index = '工况' + m.group(1)

        param = []
        factor = []
        for item in vcomb:
            param.append(item.get('LCNAME', ''))
            factor.append(item.get('FACTOR', 1))

        conditions_table.append({
            'text_index': text_index,
            'type': comb_type,
            'param': param,
            'factor': factor,
        })

    # 按工况编号和类型排序（标准在前，基本在后）
    def _sort_key(item):
        m = re.search(r'工况(.+)', item['text_index'])
        num = _chinese_to_arabic(m.group(1)) if m else 0
        return (num, 0 if item['type'] == '标准' else 1)
    conditions_table.sort(key=_sort_key)

    # conditions：取每个工况的标准组合，param 拼接为 + 连接字符串
    conditions = []
    seen = set()
    for item in conditions_table:
        if item['text_index'] in seen:
            continue
        seen.add(item['text_index'])
        conditions.append({
            'text_index': item['text_index'],
            'param': '+'.join(item['param']),
        })

    print("工况表:", conditions_table)
    print("工况:", conditions)
    return conditions_table, conditions


def vehicle_subdoc_render(main_tpl, vehicle_params_list, docx_open_path, output_dir=''):
    """渲染车辆设备参数子模板并返回 subdoc 对象。
    直接将 vehicle_params_list 传入模板，由模板内 if 块匹配渲染。

    参数:
        main_tpl:           主模板 DocxTemplate 实例（用于创建 subdoc）
        vehicle_params_list: 车辆参数列表（来自 excel_params）
        docx_open_path:     模板目录路径
        output_dir:         临时文件输出目录（默认与模板目录相同）

    返回:
        subdoc 对象，可直接插入主模板 context（如 context['vehicle_param'] = subdoc）
    """
    sub_tpl_path = os.path.join(docx_open_path, 'Trestle_cal_template_vehicle_param.docx')
    sub_tpl = DocxTemplate(sub_tpl_path)
    sub_tpl.render({'vehicle_params_list': vehicle_params_list})
    # new_subdoc 需要文件路径，先保存渲染结果为临时文件
    save_dir = output_dir or docx_open_path
    tmp_path = os.path.join(save_dir, '_vehicle_subdoc_tmp.docx')
    sub_tpl.save(tmp_path)
    subdoc = main_tpl.new_subdoc(tmp_path)
    # os.remove(tmp_path)  # 暂时保留临时文件
    return subdoc


def conditions_table_subdoc_render(main_tpl, conditions_table, vehicle_params_list, docx_open_path, output_dir=''):
    """动态生成荷载工况组合表并返回 subdoc 对象。

    表格结构：
        列 = 工况 + 组合 + 自重 + 风荷载 + 水流力 + 各车辆设备
        行 = 工况数 × 2（标准/基本），工况列纵向合并

    参数:
        main_tpl:             主模板 DocxTemplate 实例
        conditions_table:     get_conditions_info 返回的工况表数据
        vehicle_params_list:  车辆参数列表（来自 excel_params）
        docx_open_path:       模板目录路径（此函数未用，保留接口一致性）
        output_dir:           临时文件输出目录（默认系统临时目录）

    返回:
        subdoc 对象，可直接插入主模板 context
    """

    # ── 1、构建列定义 ──
    fixed_columns = ['工况', '组合', '自重', '风荷载', '水流力']
    vehicle_names = [v['name'] for v in vehicle_params_list]
    all_columns = fixed_columns + vehicle_names
    num_cols = len(all_columns)

    # ── 2、准备数据映射 ──
    # 固定荷载名 → 列索引
    fixed_map = {'自重': 2, '风荷载': 3, '水流力': 4}
    # 车辆名 → 列索引
    vehicle_col_map = {name: 5 + i for i, name in enumerate(vehicle_names)}

    # ── 3、收集工况顺序（保持 conditions_table 中的顺序） ──
    condition_order = []
    seen = set()
    for item in conditions_table:
        if item['text_index'] not in seen:
            seen.add(item['text_index'])
            condition_order.append(item['text_index'])

    # ── 4、构建行数据 ──
    # 每个工况 2 行：[标准, 基本]
    # rows_data = [{col_index: value}, ...]
    num_rows = 1 + len(condition_order) * 2  # 1 表头 + N×2 数据行
    rows_data = [{} for _ in range(num_rows)]

    for item in conditions_table:
        ci = item['text_index']
        row_offset = condition_order.index(ci) * 2
        row_idx = 1 + row_offset + (0 if item['type'] == '标准' else 1)

        # 组合列
        rows_data[row_idx][1] = item['type']

        # 遍历 param/factor 填充各列
        for pname, pfac in zip(item['param'], item['factor']):
            if pname in fixed_map:
                rows_data[row_idx][fixed_map[pname]] = pfac
            else:
                # 车辆匹配：param 包含车辆名即可
                for vname, col_idx in vehicle_col_map.items():
                    if vname in pname:
                        rows_data[row_idx][col_idx] = pfac

    # ── 5、从主模板副本创建表格（继承自定义样式） ──
    from docx import Document
    tpl_src = os.path.join(docx_open_path, 'Trestle_cal_template_main.docx')
    tmp_doc = Document(tpl_src)
    # 清空正文内容，保留样式定义和 section 属性
    for child in list(tmp_doc.element.body):
        if child.tag != '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr':
            tmp_doc.element.body.remove(child)

    table = tmp_doc.add_table(rows=num_rows, cols=num_cols)
    table.style = '无格式表格 11'


    # 填写表头
    for col_idx, col_name in enumerate(all_columns):
        cell = table.cell(0, col_idx)
        cell.text = col_name
        # 赋予段落样式
        cell.paragraphs[0].style = 'c.图文表文'

    # 填写数据行
    for row_idx in range(1, num_rows):
        for col_idx in range(num_cols):
            value = rows_data[row_idx].get(col_idx, '/')
            cell = table.cell(row_idx, col_idx)
            cell.text = str(value)
            # 赋予段落样式
            cell.paragraphs[0].style = 'c.图文表文'

    # ── 6、纵向合并工况列 ──
    for i, ci in enumerate(condition_order):
        row_start = 1 + i * 2
        row_end = row_start + 1
        cell_start = table.cell(row_start, 0)
        cell_end = table.cell(row_end, 0)
        cell_start.merge(cell_end)
        cell_start.text = ci
        # 给合并后的单元格重新赋予样式
        cell_start.paragraphs[0].style = 'c.图文表文'

    # ── 7、保存临时文件并生成 subdoc ──
    save_dir = output_dir or tempfile.gettempdir()
    tmp_path = os.path.join(save_dir, '_conditions_table_tmp.docx')
    tmp_doc.save(tmp_path)
    subdoc = main_tpl.new_subdoc(tmp_path)
    # os.remove(tmp_path)  # 暂时保留临时文件

    print("工况组合表 subdoc 生成完成")
    return subdoc


def plate_cal_params(component_info, excel_params, material, conditions, GRUP_dict_res, NODE_dict_res, deck_type):
    # 1、桥面板信息
    deck_info = component_info.get('桥面板', {})
    deck_t = deck_info.get('vsize', [0])[0]
    deck_grade = deck_info.get('matl_name', '')
    deck_E = 0
    for steel in material.get('steel_lst', []):
        if steel.get('grade') == deck_grade:
            deck_sigma = steel.get('f_1', 0) 
            deck_E = steel.get('E', 0)
            break
    # 2、桥面板下间距及支撑构件翼缘宽度
    if deck_type == '双层钢面板':
        deck_rib_type = component_info.get('纵肋', {}).get('sect_name', '')
        spacing = _rib_space_cal(GRUP_dict_res, NODE_dict_res, '纵肋', 'Y')
        # 读取纵肋翼缘宽度（H型钢 vsize=[H, B, tw, tf1, ...]，B=vsize[1]）
        rib_long_vsize = component_info.get('纵肋', {}).get('vsize', [])
        B1_half = float(rib_long_vsize[1]) / 2 if len(rib_long_vsize) > 1 else 0
        B2_half = B1_half  # 当前单层肋只有一种截面，预留不同截面入口
    elif deck_type == '单层钢面板':
        deck_rib_type = component_info.get('横肋', {}).get('sect_name', '')
        spacing = _rib_space_cal(GRUP_dict_res, NODE_dict_res, '横肋', 'X')
        # 读取横肋翼缘宽度（H型钢 vsize=[H, B, tw, tf1, ...]，B=vsize[1]）
        rib_trans_vsize = component_info.get('横肋', {}).get('vsize', [])
        B1_half = float(rib_trans_vsize[1]) / 2 if len(rib_trans_vsize) > 1 else 0
        B2_half = B1_half  # 当前单层肋只有一种截面，预留不同截面入口
    # 计算净跨度
    net_spacing = spacing - B1_half - B2_half
    # 3、最不利工况
    vehicle_params_list = excel_params.get('vehicle_params_list', [])
    condition_texts = [c['param'] for c in conditions]

    def _in_conditions(name):
        """检查车辆名是否出现在任一工况中"""
        return any(name in c_text for c_text in condition_texts)
    def _is_side_lift(name):
        """检查吊装设备在工况中是否为侧吊（车辆名后紧跟'侧吊'）"""
        for c_text in condition_texts:
            if name in c_text:
                suffix = c_text.split(name)[-1][:4]
                if '侧吊' in suffix:
                    return True
        return False

    q_max = 0
    condition_max = ''
    vehicle_force_width_max = 0
    for v in vehicle_params_list:
        v_name = v.get('name', '')
        v_type = v.get('type', '')
        v_params = v.get('params', {})
        q = 0
        vehicle_force_width = 0
        if v_type == '公路标准荷载':
            # 暂时写死
            if _in_conditions(v_name):
                q = 583.3
                if deck_type == '双层钢面板':
                    vehicle_force_width = 0.2
                elif deck_type == '单层钢面板':
                    vehicle_force_width = 0.6
            
        elif v_type == '吊装设备':
            weight = float(v_params.get('weight', 0))
            capacity = float(v_params.get('capacity', 0))
            track_length = float(v_params.get('track_length', 0))
            track_width = float(v_params.get('track_width', 0))
            if track_length > 0 and track_width > 0:
                q_prime = (weight + capacity) / track_length / track_width / 2
                # 判断侧吊/正吊
                if _in_conditions(v_name):
                    if deck_type == '双层钢面板':
                        vehicle_force_width = track_width
                    elif deck_type == '单层钢面板':
                        vehicle_force_width = track_length
                    if _is_side_lift(v_name):
                        q = round(0.7 * q_prime, 1)
                    else:
                        q = round(0.5 * q_prime, 1)
                        
        elif v_type == '钻孔设备':
            weight = float(v_params.get('weight', 0))
            track_length = float(v_params.get('track_length', 0))
            track_width = float(v_params.get('track_width', 0))
            if track_length > 0 and track_width > 0:
                q_prime = weight / track_length / track_width / 2
                if _in_conditions(v_name):
                    q = round(q_prime, 1)
                    if deck_type == '双层钢面板':
                        vehicle_force_width = track_width
                    elif deck_type == '单层钢面板':
                        vehicle_force_width = track_length      

        elif v_type == '一般车辆':
            # # 暂时写死：8方/12方混凝土罐车
            # if '8方混凝土' in v_name or '12方混凝土' in v_name:
            #     q = 708.3 if _in_conditions(v_name) else 0
            # else:
            forces = v_params.get('forces', [])
            if isinstance(forces, str):
                forces = [float(x) for x in forces.split(',') if x.strip()]
            fore_force = min(forces) if forces else 0
            back_force = max(forces) if forces else 0
            if _in_conditions(v_name) :
                q_fore = round(fore_force / 2 / 0.3 / 0.2, 1)
                q_back = round(back_force / 2 / 0.6 / 0.2, 1)
                if q_fore > q_back:
                    q = q_fore
                    if deck_type == '双层钢面板':
                        vehicle_force_width = 0.2
                    elif deck_type == '单层钢面板':
                        vehicle_force_width = 0.3
                else:
                    q = q_back
                    if deck_type == '双层钢面板':
                        vehicle_force_width = 0.2
                    elif deck_type == '单层钢面板':
                        vehicle_force_width = 0.6

        # 更新最不利
        if q > q_max:
            q_max = q
            condition_max = v_name + '荷载'
            vehicle_force_width_max = vehicle_force_width

    # 4、判断计算跨数及计算加载宽度
    spacing_m = spacing / 1000
    vehicle_force_width = vehicle_force_width_max
    if vehicle_force_width > spacing_m:
        cal_span = '三跨'
    else:
        cal_span = '单跨'
    if vehicle_force_width > 1.0:
        vehicle_force_width = 1.0
    
    deck = {
        't': deck_t,
        'rib_type': deck_rib_type,
        'rib_dist': spacing,
        'condition': condition_max,
        'grade': deck_grade,
        'vehicle_force_width': vehicle_force_width,
        'cal_span': cal_span,
        'if_satisfied': '满足',
    }

    
    # 5、计算（先用原始间距，不满足时用净跨度重算，最终都写净跨度结果）
    def _calc_deck(span, span_m, label=''):
        """计算桥面板应力和挠度，返回 (sigma, f, formula_dict)"""
        sigma_daxiao = r'\leq '
        f_daxiao = r'\leq '

        h = int(deck_t)
        E = int(deck_E)
        sigma_max = round(float(deck_sigma), 0)
        f_max_local = round(span / 400, 2)

        q_val = round(q_max * vehicle_force_width, 1)
        b_val = round(vehicle_force_width * 1000, 1)
        
        I_val = round(b_val * deck_t ** 3 / 12, 1)
        W_val = round(b_val * deck_t ** 2 / 6, 1)
        if cal_span == '三跨':
            M_val = round(1.4 * q_val * span_m ** 2 / 10, 1)
        else:
            M_val = round(1.4 * q_val * span_m ** 2 / 8, 1)
        sigma_val = round(M_val * 10 ** 6 / W_val, 1)
        f_val = round(0.677 * q_val * span ** 4 / 100 / E / I_val, 1)
        
        if sigma_val > sigma_max:
            sigma_daxiao = '>'
        if f_val > f_max_local:
            f_daxiao = '>'

        cal_q = 'q=' + str(q_max) + r'\times' + f'{vehicle_force_width}={q_val}kN/m'
        cal_I = r'I=\frac{bh^{3}}{12}=\frac{' + str(b_val) + r'\times' + str(h) + r'^{3}}{12}=' + str(format_scientific_latex(I_val, 2)) + r'mm^{4}'
        cal_W = r'W=\frac{bh^{2}}{6}=\frac{' + str(b_val) + r'\times' + str(h) + r'^{2}}{6}=' + str(format_scientific_latex(W_val, 2)) + r'mm^{3}'
        if cal_span == '三跨':
            cal_M = r'M=0.1ql^{2}=1.4\times' + str(q_val) + r'\times' + str(span_m) + r'^{2}=' + str(M_val) + r'kN\bullet m'
        else:
            cal_M = r'M=0.125ql^{2}=1.4\times' + str(q_val) + r'\times' + str(span_m) + r'^{2}=' + str(M_val) + r'kN\bullet m'
        cal_sigma = r'\mathrm{\sigma}=\frac{M}{W}=\frac{' + str(M_val) + r'\times{10}^{6}}{' + str(W_val) + r'}=' + str(sigma_val) + 'MPa' + sigma_daxiao + r'f_{m}=' + str(sigma_max) + r'MPa'
        cal_f = r'f=0.677\times\frac{ql^{4}}{100EI}=0.677\times\frac{' + str(q_val) + r'\times' + str(span) + r'^{4}}{100\times' + str(E) + r'\times' + str(format_scientific_latex(I_val, 2)) + r'}=' + str(f_val) + 'mm' + f_daxiao + r'\frac{' + str(span) + r'}{400}=' + str(f_max_local) + r'mm'

        satisfied = sigma_val <= sigma_max and f_val <= f_max_local
        formulas = {
            '{deck_cal_q}': cal_q,
            '{deck_cal_I}': cal_I,
            '{deck_cal_W}': cal_W,
            '{deck_cal_M}': cal_M,
            '{deck_cal_sigma}': cal_sigma,
            '{deck_cal_f}': cal_f,
        }
        return sigma_val, f_val, sigma_max, f_max_local, satisfied, formulas

    # 第一次计算：用原始间距
    sigma1, f1, sigma_max1, f_max1, satisfied1, formulas1 = _calc_deck(spacing, spacing_m, '原始间距')
    print(f'[桥面板] 原始间距={spacing}mm, sigma={sigma1}/{sigma_max1}, f={f1}/{f_max1}, 满足={satisfied1}')

    # 如果不满足，用净跨度重新计算
    if not satisfied1:
        net_spacing_m = net_spacing / 1000
        sigma2, f2, sigma_max2, f_max2, satisfied2, formulas2 = _calc_deck(net_spacing, net_spacing_m, '净跨度')
        print(f'[桥面板] 净跨度={net_spacing}mm, sigma={sigma2}/{sigma_max2}, f={f2}/{f_max2}, 满足={satisfied2}')
        # 使用净跨度结果
        deck['rib_dist'] = net_spacing
        deck['if_satisfied'] = '满足' if satisfied2 else '不满足'
        formula_registry = formulas2
    else:
        # 原始间距满足，但仍用净跨度重新计算（最终写净跨度结果）
        net_spacing_m = net_spacing / 1000
        sigma2, f2, sigma_max2, f_max2, satisfied2, formulas2 = _calc_deck(net_spacing, net_spacing_m, '净跨度')
        print(f'[桥面板] 净跨度={net_spacing}mm, sigma={sigma2}/{sigma_max2}, f={f2}/{f_max2}, 满足={satisfied2}')
        deck['rib_dist'] = net_spacing
        deck['if_satisfied'] = '满足' if satisfied2 else '不满足'
        formula_registry = formulas2

    return deck, formula_registry

def model_cal_params(model_post_dict, folder_path, component_info, ELEM_dict_res, SECT_dict_res, MATL_dict_res, GRUP_dict_res, NODE_dict_res, deck_type, steel_dict=None):
    """
    将 model_post_dict 后处理结果转换为模板所需的 deck/rib_long/rib_trans/dist/beam/pile/brace 字典。
    model_post_dict 键:
        model_post_rib_long:        [max_cb, max_shear, max_Dz, max_Dz_node]
        model_post_rib_trans:       [max_cb, max_shear, max_Dz, max_Dz_node]
        model_post_beam:             贝雷梁: [xiangan_zhouli, shugan_zhouli, xiegan_zhouli, max_Dz, max_Dz_node]
                                    型钢纵梁: [max_cb, max_shear, max_Dz, max_Dz_node]
        model_post_fenpeiliang:     [max_cb, max_shear, max_Dz, max_cb_ele, max_Dz_node]
        model_post_lianjiexi:       [max_cb, max_shear, max_cb_ele]
        model_post_gangguanzhuang:  [ele, Fx, My, Mz, D, t, L, sect_name, matl_name]
        model_post_reaction:        [Fz, Fxy, Mx, My]
    """
    # 图片路径辅助函数 ──
    def pic(filename):
        return os.path.join(folder_path, filename)
    # 确保切换至后处理模式
    Picture_Beamstress(folder_path, "分配梁组合应力图.jpg", ["分配梁"], "基本组合", "Combined")
    # 支撑坐标（供 L 值计算使用）及坐标→节点映射（供相对变形计算使用）
    _pile_coords_y, _pile_c2n_y = _get_group_support_coords('钢管桩', GRUP_dict_res, ELEM_dict_res, NODE_dict_res, 'Y')
    _pile_coords_x, _pile_c2n_x = _get_group_support_coords('钢管桩', GRUP_dict_res, ELEM_dict_res, NODE_dict_res, 'X')
    _henglei_coords_x, _henglei_c2n_x = _get_group_support_coords('横肋', GRUP_dict_res, ELEM_dict_res, NODE_dict_res, 'X')
    # 横肋支撑坐标（按 beam_type 派发）
    beam_type = component_info.get('beam_type', '贝雷梁')
    if beam_type == '型钢纵梁':
        _rib_support_coords_y, _rib_support_c2n_y = _get_group_support_coords(
            '型钢纵梁', GRUP_dict_res, ELEM_dict_res, NODE_dict_res, 'Y')
    elif beam_type == '贝雷梁':
        _rib_support_coords_y, _rib_support_c2n_y = _get_group_support_coords(
            '贝雷梁上弦杆', GRUP_dict_res, ELEM_dict_res, NODE_dict_res, 'Y')
    # elif beam_type == '下承式桁架梁':  # 下承式预留
    #     _rib_support_coords_y, _rib_support_c2n_y = _get_group_support_coords(
    #         '下弦杆', GRUP_dict_res, ELEM_dict_res, NODE_dict_res, 'Y')
    else:
        _rib_support_coords_y, _rib_support_c2n_y = [], {}

    # 预设 rib_long（无纵肋/单层/混凝土桥面板时为 {}，双层面板分支会重新赋值）
    rib_long = {}

    if deck_type == '双层钢面板':
        # 2、纵肋验算结果（支撑在横肋上，L 取最大变形节点两侧最近横肋 X 坐标差）
        rib_long_max_cb, rib_long_max_shear, rib_long_max_Dz, rib_long_max_Dz_node = model_post_dict.get('model_post_rib_long', [0, 0, 0, ''])
        rib_long_node_x = NODE_dict_res.get('NODE', {}).get(str(rib_long_max_Dz_node), {}).get('X', 0)
        rib_long_L, rib_long_cantilever = _find_span_length(rib_long_node_x, _henglei_coords_x)
        rib_long_div = 250 if rib_long_cantilever else 400
        print(f"[纵肋] 节点X={rib_long_node_x}, 横肋X坐标={_henglei_coords_x}, L={rib_long_L}, 悬臂={rib_long_cantilever}")
        rib_long_dz_max = round(rib_long_L / rib_long_div, 2) if rib_long_L > 0 else 0
        rib_long_node_y = NODE_dict_res.get('NODE', {}).get(str(rib_long_max_Dz_node), {}).get('Y', 0)
        rib_long_rel_Dz = round(_compute_relative_deformation(
            float(rib_long_max_Dz), rib_long_max_Dz_node,
            _henglei_coords_x, _henglei_c2n_x, rib_long_node_x, '标准组合(CB:all)',
            NODE_dict_res, 'Y', rib_long_node_y), 2)
        rib_long_info = component_info.get('纵肋', {})
        rib_long_sec = rib_long_info.get('sect_name', '')
        rib_long_matl = rib_long_info.get('matl_name', '')
        rib_long_vsize = rib_long_info.get('vsize', [])
        rib_long_flange_t, rib_long_web_t = _get_flange_web_t(rib_long_vsize, rib_long_info.get('shape', ''))
        rib_long_res = _lookup_allowable_stress(rib_long_matl, rib_long_flange_t, rib_long_web_t, steel_dict)
        rib_long_cs_max = rib_long_res['f']
        rib_long_ssz_max = rib_long_res['fv']
        rib_long = {
            'cs_pic': pic('纵肋组合应力图.jpg'),
            'ssz_pic': pic('纵肋剪应力图.jpg'),
            'dz_pic': pic('纵肋竖向变形图.jpg'),
            'sec': rib_long_sec,
            'cs': rib_long_max_cb,
            'ssz': rib_long_max_shear,
            'cs_max': rib_long_cs_max,
            'ssz_max': rib_long_ssz_max,
            'dz': rib_long_rel_Dz,
            'dz_max': rib_long_dz_max,
            'if_satisfied': '满足' if abs(float(rib_long_max_cb)) <= rib_long_cs_max and abs(float(rib_long_max_shear)) <= rib_long_ssz_max and abs(float(rib_long_rel_Dz)) <= rib_long_dz_max else '不满足',
        }

    # 3、横肋验算结果（支撑在贝雷梁上弦杆/纵梁上，L 取最大变形节点两侧最近贝雷梁上弦杆 Y 坐标差）
    rib_trans_max_cb, rib_trans_max_shear, rib_trans_max_Dz, rib_trans_max_Dz_node = model_post_dict.get('model_post_rib_trans', [0, 0, 0, ''])
    rib_trans_node_y = NODE_dict_res.get('NODE', {}).get(str(rib_trans_max_Dz_node), {}).get('Y', 0)
    rib_trans_L, rib_trans_cantilever = _find_span_length(rib_trans_node_y, _rib_support_coords_y)
    rib_trans_div = 250 if rib_trans_cantilever else 400
    print(f"[横肋] 节点Y={rib_trans_node_y}, 纵梁支撑Y坐标={_rib_support_coords_y}, L={rib_trans_L}, 悬臂={rib_trans_cantilever}")
    rib_trans_dz_max = round(rib_trans_L / rib_trans_div, 2) if rib_trans_L > 0 else 0
    # 相对变形 = |横肋max_Dz - 纵梁支撑同X坐标节点Dz|
    rib_trans_node_x = NODE_dict_res.get('NODE', {}).get(str(rib_trans_max_Dz_node), {}).get('X', 0)
    rib_trans_rel_Dz = round(_compute_relative_deformation(
        float(rib_trans_max_Dz), rib_trans_max_Dz_node,
        _rib_support_coords_y, _rib_support_c2n_y, rib_trans_node_y, '标准组合(CB:all)',
        NODE_dict_res, 'X', rib_trans_node_x), 2)
    rib_trans_info = component_info.get('横肋', {})
    rib_trans_sec = rib_trans_info.get('sect_name', '')
    rib_trans_matl = rib_trans_info.get('matl_name', '')
    rib_trans_vsize = rib_trans_info.get('vsize', [])
    rib_trans_flange_t, rib_trans_web_t = _get_flange_web_t(rib_trans_vsize, rib_trans_info.get('shape', ''))
    rib_trans_res = _lookup_allowable_stress(rib_trans_matl, rib_trans_flange_t, rib_trans_web_t, steel_dict)
    rib_trans_cs_max = rib_trans_res['f']
    rib_trans_ssz_max = rib_trans_res['fv']
    rib_trans = {
        'cs_pic': pic('横肋组合应力图.jpg'),
        'ssz_pic': pic('横肋剪应力图.jpg'),
        'dz_pic': pic('横肋竖向变形图.jpg'),
        'sec': rib_trans_sec,
        'cs': rib_trans_max_cb,
        'ssz': rib_trans_max_shear,
        'cs_max': rib_trans_cs_max,
        'ssz_max': rib_trans_ssz_max,
        'dz': rib_trans_rel_Dz,
        'dz_max': rib_trans_dz_max,
        'if_satisfied': '满足' if abs(float(rib_trans_max_cb)) <= rib_trans_cs_max and abs(float(rib_trans_max_shear)) <= rib_trans_ssz_max and abs(float(rib_trans_rel_Dz)) <= rib_trans_dz_max else '不满足',
    }

    # 4、分配梁验算结果（多截面：用最大应力单元号查截面）
    # 横向分配梁支撑在钢管桩上（L 取 Y 坐标差），纵向分配梁支撑在钢管桩上（L 取 X 坐标差）
    dist_max_cb, dist_max_shear, dist_max_Dz, dist_max_ele, dist_max_Dz_node = model_post_dict.get('model_post_fenpeiliang', [0, 0, 0, '', ''])
    _dist = _get_sec_mat_by_ele(dist_max_ele, ELEM_dict_res, SECT_dict_res, MATL_dict_res) if dist_max_ele else {}
    dist_sec, dist_matl, dist_vsize, dist_shape = _dist.get('sect_name', ''), _dist.get('matl_name', ''), _dist.get('vsize', []), _dist.get('shape', '')
    # 判断是横向还是纵向分配梁（用截面 vsize 匹配 component_info key）
    dist_is_longitudinal = False
    if dist_max_ele:
        for _key in component_info:
            if '纵向分配梁' in _key and isinstance(component_info[_key], dict):
                if component_info[_key].get('vsize') == dist_vsize:
                    dist_is_longitudinal = True
                    break
    if dist_is_longitudinal:
        dist_node_coord = NODE_dict_res.get('NODE', {}).get(str(dist_max_Dz_node), {}).get('X', 0)
        dist_L, dist_cantilever = _find_span_length(dist_node_coord, _pile_coords_x)
        print(f"[纵向分配梁] 节点X={dist_node_coord}, 钢管桩X坐标={_pile_coords_x}, L={dist_L}, 悬臂={dist_cantilever}")
        # 相对变形 = |纵向分配梁max_Dz - 钢管桩同Y坐标节点Dz|
        dist_match_val = NODE_dict_res.get('NODE', {}).get(str(dist_max_Dz_node), {}).get('Y', 0)
        dist_rel_Dz = round(_compute_relative_deformation(
            float(dist_max_Dz), dist_max_Dz_node,
            _pile_coords_x, _pile_c2n_x, dist_node_coord, '标准组合(CB:all)',
            NODE_dict_res, 'Y', dist_match_val), 2)
    else:
        dist_node_coord = NODE_dict_res.get('NODE', {}).get(str(dist_max_Dz_node), {}).get('Y', 0)
        dist_L, dist_cantilever = _find_span_length(dist_node_coord, _pile_coords_y)
        print(f"[横向分配梁] 节点Y={dist_node_coord}, 钢管桩Y坐标={_pile_coords_y}, L={dist_L}, 悬臂={dist_cantilever}")
        # 相对变形 = |横向分配梁max_Dz - 钢管桩同X坐标节点Dz|
        dist_match_val = NODE_dict_res.get('NODE', {}).get(str(dist_max_Dz_node), {}).get('X', 0)
        dist_rel_Dz = round(_compute_relative_deformation(
            float(dist_max_Dz), dist_max_Dz_node,
            _pile_coords_y, _pile_c2n_y, dist_node_coord, '标准组合(CB:all)',
            NODE_dict_res, 'X', dist_match_val), 2)
    dist_div = 250 if dist_cantilever else 400
    dist_dz_max = round(dist_L / dist_div, 2) if dist_L > 0 else 0
    dist_flange_t, dist_web_t = _get_flange_web_t(dist_vsize, dist_shape)
    dist_res = _lookup_allowable_stress(dist_matl, dist_flange_t, dist_web_t, steel_dict)
    dist_cs_max = dist_res['f']
    dist_ssz_max = dist_res['fv']
    dist = {
        'cs_pic': pic('分配梁组合应力图.jpg'),
        'ssz_pic': pic('分配梁剪应力图.jpg'),
        'dz_pic': pic('分配梁竖向变形图.jpg'),
        'sec': dist_sec,
        'cs': dist_max_cb,
        'ssz': dist_max_shear,
        'cs_max': dist_cs_max,
        'ssz_max': dist_ssz_max,
        'dz': dist_rel_Dz,
        'dz_max': dist_dz_max,
        'if_satisfied': '满足' if abs(float(dist_max_cb)) <= dist_cs_max and abs(float(dist_max_shear)) <= dist_ssz_max and abs(float(dist_rel_Dz)) <= dist_dz_max else '不满足',
    }

    # 5、纵梁验算结果（按 beam_type 派发）
    if beam_type == '贝雷梁':
        # 贝雷梁验算：弦杆/竖杆/斜杆轴力（支撑在钢管桩上，L 取最大变形节点两侧最近钢管桩 X 坐标差 = 跨径）
        bailey_xiangan, bailey_shugan, bailey_xiegan, bailey_Dz, bailey_max_Dz_node = model_post_dict.get('model_post_beam', [0, 0, 0, 0, ''])
        bailey_node_x = NODE_dict_res.get('NODE', {}).get(str(bailey_max_Dz_node), {}).get('X', 0)
        bailey_L, bailey_cantilever = _find_span_length(bailey_node_x, _pile_coords_x)
        bailey_div = 250 if bailey_cantilever else 400
        print(f"[贝雷梁] 节点X={bailey_node_x}, 钢管桩X坐标={_pile_coords_x}, L={bailey_L}, 悬臂={bailey_cantilever}")
        bailey_dz_max = round(bailey_L / bailey_div, 2) if bailey_L > 0 else 0
        bailey_chord_f_allow = 560
        bailey_vertical_f_allow = 210
        bailey_diagonal_f_allow = 171
        beam = {
            'chord_fx_pic': pic('贝雷弦杆轴力图.jpg'),
            'vertical_fx_pic': pic('贝雷竖杆轴力图.jpg'),
            'diagonal_fx_pic': pic('贝雷斜杆轴力图.jpg'),
            'dz_pic': pic('贝雷竖向变形图.jpg'),
            'chord_fx': bailey_xiangan,
            'vertical_fx': bailey_shugan,
            'diagonal_fx': bailey_xiegan,
            'chord_f_allow': bailey_chord_f_allow,
            'vertical_f_allow': bailey_vertical_f_allow,
            'diagonal_f_allow': bailey_diagonal_f_allow,
            'chord_satisfied': '满足' if abs(float(bailey_xiangan)) <= bailey_chord_f_allow else '不满足',
            'vertical_satisfied': '满足' if abs(float(bailey_shugan)) <= bailey_vertical_f_allow else '不满足',
            'diagonal_satisfied': '满足' if abs(float(bailey_xiegan)) <= bailey_diagonal_f_allow else '不满足',
            'dz': bailey_dz_max,
            'dz_max': bailey_dz_max,
        }

    elif beam_type == '型钢纵梁':
        # 型钢纵梁验算：组合应力/剪应力/变形（支撑在钢管桩上，L 取最大变形节点两侧最近钢管桩 X 坐标差）
        steel_beam_cb, steel_beam_shear, steel_beam_Dz, steel_beam_Dz_node = \
            model_post_dict.get('model_post_beam', [0, 0, 0, ''])
        steel_beam_node_x = NODE_dict_res.get('NODE', {}).get(str(steel_beam_Dz_node), {}).get('X', 0)
        steel_beam_L, steel_beam_cantilever = _find_span_length(steel_beam_node_x, _pile_coords_x)
        steel_beam_div = 250 if steel_beam_cantilever else 400
        print(f"[型钢纵梁] 节点X={steel_beam_node_x}, 钢管桩X坐标={_pile_coords_x}, L={steel_beam_L}, 悬臂={steel_beam_cantilever}")
        steel_beam_dz_max = round(steel_beam_L / steel_beam_div, 2) if steel_beam_L > 0 else 0
        steel_beam_info = component_info.get('型钢纵梁', {})
        steel_beam_sec = steel_beam_info.get('sect_name', '')
        steel_beam_matl = steel_beam_info.get('matl_name', '')
        steel_beam_flange_t, steel_beam_web_t = _get_flange_web_t(
            steel_beam_info.get('vsize', []), steel_beam_info.get('shape', ''))
        steel_beam_res = _lookup_allowable_stress(
            steel_beam_matl, steel_beam_flange_t, steel_beam_web_t, steel_dict)
        beam = {
            'sec': steel_beam_sec,
            'cs_pic': pic('型钢纵梁组合应力图.jpg'),
            'ssz_pic': pic('型钢纵梁剪应力图.jpg'),
            'dz_pic': pic('型钢纵梁竖向变形图.jpg'),
            'cs': steel_beam_cb,
            'ssz': steel_beam_shear,
            'cs_max': steel_beam_res['f'],
            'ssz_max': steel_beam_res['fv'],
            'dz': steel_beam_Dz,
            'dz_max': steel_beam_dz_max,
            'if_satisfied': (
                '满足' if abs(float(steel_beam_cb)) <= steel_beam_res['f']
                and abs(float(steel_beam_shear)) <= steel_beam_res['fv']
                and abs(float(steel_beam_Dz)) <= steel_beam_dz_max
                else '不满足'
            ),
        }

    # elif beam_type == '下承式桁架梁':  # 下承式预留
    #     raise NotImplementedError("下承式桁架梁验算暂未实现")

    else:
        beam = {}

    # 6、钢管桩验算结果
    pile_ele, pile_Fx, pile_My, pile_Mz, pile_D, pile_t, pile_L, pile_sec, pile_matl = model_post_dict.get('model_post_gangguanzhuang', [0, 0, 0, 0, 0, 0, 0, '', ''])
    # 钢管桩为圆管（SHAPE='P'），翼缘=腹板=壁厚
    pile_res = _lookup_allowable_stress(pile_matl, pile_t, pile_t, steel_dict)
    pile_cs_max = pile_res['f']
    pile_ssz_max = pile_res['fv']
    # 从模型材料库读取弹性模量 E (MPa)
    pile_E = 206000  # 默认钢材弹性模量
    for _mid, _mdata in MATL_dict_res.get('MATL', {}).items():
        if _mdata.get('NAME', '') == pile_matl:
            pile_E = _mdata.get('E', 206000)
            break
    # 桩稳定性验算 (GB50017-2017 §8.2.1 压弯构件平面内稳定性)
    _pile_Fx = abs(float(pile_Fx))
    _pile_My = abs(float(pile_My))
    _pile_Mz = abs(float(pile_Mz))
    _pile_D = float(pile_D)
    _pile_t = float(pile_t)
    _pile_L = float(pile_L)
    if _pile_D > 0 and _pile_t > 0 and _pile_L > 0:
        _pile_Area = math.pi * (_pile_D ** 2 - (_pile_D - 2 * _pile_t) ** 2) / 4
        _pile_Ix = math.pi * (_pile_D ** 4 - (_pile_D - 2 * _pile_t) ** 4) / 64
        _pile_ix = math.sqrt(_pile_Ix / _pile_Area)
        _pile_Phi = stability_coefficient("b", _pile_L, _pile_ix, 235, pile_E)
        _pile_Mm = math.sqrt(_pile_My ** 2 + _pile_Mz ** 2)
        _pile_Wx = _pile_Ix / (_pile_D / 2)
        _pile_lambda = _pile_L / _pile_ix
        _pile_Nex = (math.pi ** 2) * pile_E * _pile_Area / (1.1 * _pile_lambda ** 2)
        _pile_sigma = (_pile_Fx * 1000 / (_pile_Phi * _pile_Area)) + (_pile_Mm * 1000000 / (1.15 * _pile_Wx * (1 - 0.8 * _pile_Fx * 1000 / _pile_Nex)))
        if _pile_sigma <= pile_cs_max:
            pile_stability_satisfied = '满足'
            pile_sigma_daxiao = r'\leq '
        else:
            pile_stability_satisfied = '不满足'
            pile_sigma_daxiao = r'>'
    else:
        _pile_sigma = 0
        pile_stability_satisfied = '满足'
    # 钢管桩稳定性公式 LaTeX
    pile_formula_registry = {}
    if _pile_D > 0 and _pile_t > 0 and _pile_L > 0:
        _pile_Nex_latex = GB50017_2017_8_2_1_2 + r'=\frac{\pi^2\times' + format_scientific_latex(pile_E) + r'\times' + str(round(_pile_Area, 1)) + r'}{1.1\times{' + str(round(_pile_lambda, 1)) + r'}^2}=' + str(round(_pile_Nex, 1)) + 'N'
        _pile_sigma_latex = GB50017_2017_8_2_1_1 + r'=\frac{' + str(round(_pile_Fx, 1)) + r'\times{10}^3}{' + str(round(_pile_Phi, 3)) + r'\times' + str(round(_pile_Area, 1)) + r'}+\frac{1\times' + str(round(_pile_Mm, 1)) + r'\times{10}^6}{1.15\times' + str(round(_pile_Wx, 1)) + r'\times\left(1-0.8\times\frac{' + str(round(_pile_Fx, 1)) + r'\times{10}^3}{' + str(round(_pile_Nex, 1)) + r'}\right)}=' + str(round(_pile_sigma, 1)) + 'MPa' + pile_sigma_daxiao + str(pile_cs_max)
        pile_formula_registry['{pile_Nex_cal}'] = _pile_Nex_latex
        pile_formula_registry['{pile_sigma_cal}'] = _pile_sigma_latex

    pile = {
        'fx_pic': pic('钢管桩轴力图.jpg'),
        'm_pic': pic('钢管桩弯矩图.jpg'),
        'sec': pile_sec,
        'fx_max': pile_Fx,
        'm_max': max(abs(float(pile_My)), abs(float(pile_Mz))),
        'D': pile_D,
        't': pile_t,
        'L': pile_L,
        'cs_max': pile_cs_max,
        'ssz_max': pile_ssz_max,
        'stability_satisfied': pile_stability_satisfied,
        'formula_registry': pile_formula_registry,
    }

    # 7、联结系验算结果（多截面：用最大应力单元号查截面）
    brace_max_cb, brace_max_shear, brace_max_ele = model_post_dict.get('model_post_lianjiexi', [0, 0, ''])
    brace_sec, brace_matl, brace_vsize = '', '', []
    brace_shape = ''
    if brace_max_ele:
        _brace = _get_sec_mat_by_ele(brace_max_ele, ELEM_dict_res, SECT_dict_res, MATL_dict_res)
        brace_sec, brace_matl, brace_vsize, brace_shape = _brace.get('sect_name', ''), _brace.get('matl_name', ''), _brace.get('vsize', []), _brace.get('shape', '')
    brace_flange_t, brace_web_t = _get_flange_web_t(brace_vsize, brace_shape)
    brace_res = _lookup_allowable_stress(brace_matl, brace_flange_t, brace_web_t, steel_dict)
    brace_cs_max = brace_res['f']
    brace_ssz_max = brace_res['fv']
    brace = {
        'cs_pic': pic('联结系组合应力图.jpg'),
        'ssz_pic': pic('联结系剪应力图.jpg'),
        'sec': brace_sec,
        'cs': brace_max_cb,
        'ssz': brace_max_shear,
        'cs_max': brace_cs_max,
        'ssz_max': brace_ssz_max,
        'if_satisfied': '满足' if abs(float(brace_max_cb)) <= brace_cs_max and abs(float(brace_max_shear)) <= brace_ssz_max else '不满足',
    }

    return rib_long, rib_trans, dist, beam, pile, brace




#=====================================================================================================================================================================

# 基础计算函数

def match_soil_entry(reaction_dict, component_info, soil_param_index):
    """根据反力节点匹配 soil_param_index 中的土层条目。

    匹配规则：从 component_info 找到反力节点所属桩组（如 "1-钢管桩"），
    取编号 1 直接索引 soil_param_index[1]。
    若无匹配，取 component_info 中第一个桩组。
    若仍无，取 soil_param_index 第一条。

    返回: (soil_entry_dict, pile_key_or_None)
    """
    if not soil_param_index:
        return None, None

    # 从 Fz_max 取反力节点号
    fzc = reaction_dict.get('Fz_max', ['0', '0', '', 0, 0, 0, 0, 0, 0])
    nodeid = str(fzc[1]) if len(fzc) > 1 else ''

    # 遍历 component_info 找反力节点所属桩组
    pile_key = None
    for key, info in component_info.items():
        if not key.endswith('钢管桩') or not isinstance(info, dict):
            continue
        if nodeid in info.get('node_ids', set()):
            pile_key = key
            break

    # 兜底：取第一个桩组
    if pile_key is None:
        for key, info in component_info.items():
            if key.endswith('钢管桩') and isinstance(info, dict):
                pile_key = key
                break

    if pile_key:
        m = re.match(r'^(\d+)-', pile_key)
        idx = int(m.group(1)) if m else -1
        if 0 <= idx < len(soil_param_index):
            print(f"[match_soil_entry] {pile_key} → soil_param_index[{idx}]")
            return soil_param_index[idx], pile_key
        else:
            print(f"[match_soil_entry] {pile_key} 索引{idx}越界(0~{len(soil_param_index)-1})")

    # 最终兜底：取第一条
    print(f"[match_soil_entry] 无桩组匹配，取 soil_param_index[0]")
    return soil_param_index[0], pile_key


def compute_average_gs(soil_entry, foundation_base_level):
    """从地面到基底标高之间的加权平均土容重 (kN/m³)。

    参数:
        soil_entry:             单条土层条目 dict，含 {'layers': [...], 'ground_level': float}
        foundation_base_level:  基底标高 (m, 负值, 绝对标高)

    算法：逐层累计 (weight × thickness)，最后 ÷ 总深度。
    若基底低于最深土层底，取最后一层参数继续计算。
    默认返回 18 kN/m³（无土层数据时）。
    """
    if not soil_entry:
        print("[compute_average_gs] soil_entry 为空，使用默认 18 kN/m³")
        return 18.0

    layers = soil_entry.get('layers', [])
    gl = float(soil_entry.get('ground_level', 0))

    if not layers:
        print("[compute_average_gs] 土层为空，使用默认 18 kN/m³")
        return 18.0

    # 从地面到基底的总深度
    total_depth = gl - float(foundation_base_level)  # 正值 (m)
    if total_depth <= 0:
        print(f"[compute_average_gs] 基底高于地面 (ground={gl}, base={foundation_base_level})，使用默认 18")
        return 18.0

    accumulated = 0.0   # 累计 weight × thickness
    remaining = total_depth
    last_weight = 18.0

    for layer in layers:
        if remaining <= 0:
            break
        thickness = float(layer.get('thickness', 0))
        weight = float(layer.get('weight', 18))
        last_weight = weight
        if thickness <= 0:
            continue
        used = min(thickness, remaining)
        accumulated += weight * used
        remaining -= used

    # 若基底低于最深土层底，取最后一层参数继续
    if remaining > 0:
        accumulated += last_weight * remaining
        print(f"[compute_average_gs] 基底低于土层底 {remaining:.2f}m，使用最后一层 weight={last_weight}")

    gs = round(accumulated / total_depth, 2)
    print(f"[compute_average_gs] ground={gl}, base={foundation_base_level}, depth={total_depth:.2f}m, "
          f"{len(layers)}层, gs={gs} kN/m³")
    return gs


# xyz是基础平面尺寸, gc是基础容重, h1为基础底标高, h2为软弱下卧层顶标高, gs为土容重, angle为扩散角
def foundation_pressure(x, y, z, gc, Fk1, Fk2, Mk, h1, h2, gs, angle, fa, faz, type_M):
    # Ak, 基础底面面积
    # Gk, 基础自重
    # pk, 基础底面处的平均压力值
    # Wx, 基础抗弯截面模量
    # pkmax, 基础底面边缘的最大压力值
    # pcz, 软弱下卧层顶面处土的自重压力值
    # pc, 基础底面处土的自重压力值
    # Az, 基础底面扩散影响面积
    # pz, 软弱下卧层顶面处的附加压力值
    # 初始化
    Ak, Gk, pk, Mk_e, Mk_a, Eccentricity, Wk, pkmax, pcz, pc, pz, if_satisfied_fa1, if_satisfied_fa2, if_satisfied_faz = [0, 0, 0, 0, 0, "", 0, 0, 0, 0, 0, "", "", ""]
    # Gk
    Gk = foundation_Gk(x, y, z, h1, gc, gs)
    # pk
    Ak, pk = foundation_pk(x, y, Gk, Fk1)
    if pk < fa:
        if_satisfied_fa1 = "满足"
    else:
        if_satisfied_fa1 = "不满足"
    # pkmax
    if Mk != 0:
        # 偏心检算
        Mk_e, Mk_a, Eccentricity = cal_Eccentricity(Gk, Fk2, Mk, x, y, type_M)
        if Eccentricity == "大偏心":
            pkmax = foundation_pkmax_Large_Eccentricity(Gk, Fk2, x, y, Mk_a, type_M)
            pkmin = 0
        elif Eccentricity == "小偏心":
            Wk, pkmax, pkmin = foundation_pkmax_pkmin(x, y, Gk, Fk2, Mk, type_M)
        if pkmax < 1.2 * fa:
            if_satisfied_fa2 = "满足" 
        else:
            if_satisfied_fa2 = "不满足"
    # pcz/pz
    if h2 != 0:
        pcz = foundation_pcz(h2, gs)
        pc, pz = foundation_pz(x, y, z, pk, h1, h2, gs, angle)
        if pcz + pz < faz:
            if_satisfied_faz = "满足"
        else:
            if_satisfied_faz = "不满足"
            
    return [Ak, Gk, pk, Mk_e, Mk_a, Eccentricity, Wk, pkmax, pkmin, pcz, pc, pz, if_satisfied_fa1, if_satisfied_fa2, if_satisfied_faz, type_M]


# =====================================================================================================================================================================
# 基础计算分项函数
def calc_driven_pile(excel_params, component_info, secondary_dialog_values, reaction_dict, material=None, pile_soil=None):
    """打入桩计算总函数。

    参数:
        excel_params:           Excel参数表读取结果（含 soil_param）
        pile_soil:              匹配后的桩组土层信息（含 layers/ground_level/pile_bottom_depth）
        component_info:         模型截面/材料信息（含 各桩 vsize/L）
        secondary_dialog_values: UI二级窗口参数
        reaction_dict:          见下方格式说明

    reaction_dict 格式（来自 Post_processing_reaction 的 list6）:
        'Fz_max':   [序号, nodeid, 组合名, Fx, Fy, Fz, Mx, My, Mz]  最大竖向力
        'Fxy_max':  [序号, nodeid, 组合名, Fx, Fy, Fz, Mx, My, Mz]  最大水平力
        'pile_max': [ele, Fx, My, Mz]                                钢管桩最不利内力

    返回:
        drivenplie: dict，预设 context 变量，字典结构后续可编辑
    """
    # ── 反力 ──
    Fz_max = reaction_dict.get('Fz_max', ['0', '0', '', 0, 0, 0, 0, 0, 0])
    Fxy_max = reaction_dict.get('Fxy_max', ['0', '0', '', 0, 0, 0, 0, 0, 0])
    reaction_nodeid = str(Fz_max[1])
    Fz = abs(float(Fz_max[5]))      # Fz 分量
    Fx = abs(float(Fxy_max[3]))     # Fx 分量（水平力最大工况）
    Fy = abs(float(Fxy_max[4]))     # Fy 分量
    Fxy = math.sqrt(Fx**2 + Fy**2)

    # ── 桩截面参数（根据反力节点号匹配所属桩组） ──
    pile_info = {}
    for key, info in component_info.items():
        if not key.endswith('钢管桩') or not isinstance(info, dict):
            continue
        node_ids = info.get('node_ids', [])
        if reaction_nodeid in node_ids:
            pile_info = info
            break
    # 若未找到精确匹配，取第一个桩组兜底
    if not pile_info:
        pile_key = next((k for k in component_info if k.endswith('钢管桩') and isinstance(component_info[k], dict)), None)
        if pile_key:
            pile_info = component_info[pile_key]

    vsize = pile_info.get('vsize', [])
    d = float(vsize[0]) if vsize else 0              # 直径 mm
    t = float(vsize[1]) if len(vsize) > 1 else 0     # 壁厚 mm
    L = float(pile_info.get('L', 0))                 # 桩长 mm

    # ── 桩入土深度：取 pile_soil 的 pile_bottom_depth（m） ──
    pile_bottom_depth = pile_soil.get('pile_bottom_depth', 0.0) if pile_soil else 0.0
    print(f"[calc_driven_pile] pile_info: L={L}mm, d={d}mm, t={t}mm, pile_bottom_depth={pile_bottom_depth}m")

    # ── 土层参数 ──
    if not pile_soil or not pile_soil.get('layers'):
        raise ValueError("[calc_driven_pile] pile_soil 为空，请检查 soil_param_index 是否正确传入")
    soil_layers = pile_soil['layers']
    print(f"[calc_driven_pile] 桩底深={pile_soil.get('pile_bottom_depth')}m")
    s_layer_h = [layer.get('thickness', 0) for layer in soil_layers]
    s_qsk = [layer.get('qsik', 0) for layer in soil_layers]
    s_psk = [layer.get('qpk', 0) for layer in soil_layers]
    friction = [layer.get('phi', 0) for layer in soil_layers]
    Cohesion = [layer.get('c', 0) for layer in soil_layers]
    print(f"[calc_driven_pile] 土层参数: {len(s_layer_h)}层, s_qsk={s_qsk}, s_psk={s_psk}, friction={friction}, Cohesion={Cohesion}")

    # ── UI 参数 ──
    # 打入桩 dict: {'pile_d': ..., 'x0a': ...}
    vals = secondary_dialog_values if isinstance(secondary_dialog_values, dict) else {}
    vb = float(vals.get('x0a', 10))  # x0a = 水平位移允许值 mm
    print(f'[calc_driven_pile] vals={vals}, vb={vb}mm')
    print(f'[calc_driven_pile] 土层参数: s_layer_h={s_layer_h}, s_qsk={s_qsk}, s_psk={s_psk}, friction={friction}, Cohesion={Cohesion}')

    # ── 竖向承载力 JGJ94-2008 5.3.7 ──
    Qsk = JGJ94_Qsk_DrivenPile(d / 1000, s_qsk, s_layer_h)
    Qpk = JGJ94_Qpk_steelpile(d / 1000, s_psk[-1] if s_psk else 0, s_layer_h[-1] if s_layer_h else 0)
    Quk = JGJ94_Quk(Qsk, Qpk)
    Ra = JGJ94_Ra(Quk)

    # ── 水平承载力 JGJ94-2008 5.7.2-2（钢桩适用） ──
    phi = friction[0] if friction else 0
    c = Cohesion[0] if Cohesion else 0
    # 从 material['steel_lst'] 查找桩材弹性模量
    pile_matl = pile_info.get('matl_name', '')
    E = 206000  # 默认值 MPa
    if material and pile_matl:
        for s in material.get('steel_lst', []):
            if s.get('grade', '') == pile_matl:
                E = s.get('E', 206000)
                break
    E = E * 1000 #KPa
    I = math.pi * ((d / 1000)**4 - ((d / 1000) - 2 * (t / 1000))**4) / 64  # m4
    print(f"[calc_driven_pile] 桩材弹性模量 E={E} kPa, 桩截面抗弯模量 I={I} m4")
    if d / 1000 > 1:
        b0 = 0.9 * (d / 1000 + 1)
    else:
        b0 = 0.9 * (1.5 * d / 1000 + 0.5)
    m_lst = Subgrade_Reaction_Coefficient(friction, Cohesion, vb)
    m = next((v for v in m_lst if v != 0), 1000)
    alpha = Displacement_Coefficient(m, b0, E, I)
    Vm, Vx, alpha_h = JGJ94_Table_Vm_Vs(alpha, pile_bottom_depth)
    Rha = JGJ94_Rha2(alpha, E, I, Vx, vb)

    # ── 判定 ──
    if_Nk_satisfied = '满足' if Ra >= Fz else '不满足'
    if_Hk_satisfied = '满足' if Rha >= Fxy else '不满足'

    # ── 材料类型（模板 drivenplie.material 分支） ──
    pile_material_type = '钢管桩'  # 当前仅支持打入桩；若后续扩展预制桩，由 UI 或 component_info 判定

    # ── 综合判定 ──
    satisfied = '满足' if if_Nk_satisfied == '满足' and if_Hk_satisfied == '满足' else '不满足'

    # ── 预设 context 字典（键名与模板 {{ drivenplie.xxx }} 对应） ──
    drivenplie = {
        # 模板直接引用
        'd': d,                                 # {{ drivenplie.d }}
        't': t,                                 # {{ drivenplie.t }}
        'h': pile_bottom_depth,                  # {{ drivenplie.h }}  入土深度 m
        'Nk': round(Fz, 1),                     # {{ drivenplie.Nk }}
        'Hk': round(Fxy, 1),                    # {{ drivenplie.Hk }}
        'vb': vb,                               # {{ drivenplie.vb }}  水平位移限值 mm
        'b0': round(b0, 3),                     # {{ drivenplie.b0 }}
        'alpha_h': alpha_h,                     # {{ drivenplie.alpha_h }}
        'vm': Vm,                               # {{ drivenplie.vm }}
        'vx': Vx,                               # {{ drivenplie.vx }}
        'material': pile_material_type,          # {{ drivenplie.material }}
        'concrete_grade': '',                   # {{ drivenplie.concrete_grade }}（钢管桩为空）
        'satisfied': satisfied,                 # {{ drivenplie.satisfied }}
        # 计算值
        'Qsk': Qsk,
        'Qpk': Qpk,
        'Quk': Quk,
        'Ra': Ra,
        'phi': phi,
        'c': c,
        'm': round(m, 1),
        'EI': round(E * I, 1),
        'alpha': alpha,
        'Rha': Rha,
        'if_Nk_satisfied': if_Nk_satisfied,
        'if_Hk_satisfied': if_Hk_satisfied,
    }

    # ── 公式 LaTeX 构建 ──
    # 竖向承载力公式
    latex_Quk = JGJ94_2008_5_3_7_1 + '=' + str(Qsk) + '+' + str(Qpk) + '=' + str(Quk) + 'kN'
    if Ra >= Fz:
        latex_Ra = JGJ94_2008_5_2_2 + r'=\frac{1}{2}\times' + str(Quk) + '=' + str(Ra) + 'kN>N_k=' + str(round(Fz, 1)) + 'kN'
    else:
        latex_Ra = JGJ94_2008_5_2_2 + r'=\frac{1}{2}\times' + str(Quk) + '=' + str(Ra) + 'kN<N_k=' + str(round(Fz, 1)) + 'kN'
    # 水平承载力公式（Rha2 路径）
    latex_m = JGJ_120_2012_4_1_6 + r'=\frac{0.2\times{' + str(phi) + r'}^2-' + str(phi) + '+' + str(c) + r'}{' + str(vb) + r'}\times1000=' + str(round(m, 1)) + 'kN/m^4'
    latex_alpha = JGJ94_2008_5_7_5 + r'=\sqrt[5]{\frac{' + str(round(m, 1)) + r'\times' + str(round(b0, 3)) + r'}{' + str(round(E * I, 1)) + r'}}=' + str(alpha) + r'\left(1/m\right)'
    if Rha >= Fxy:
        latex_Rha = JGJ94_2008_5_7_2_2 + r'=0.75\times\frac{{' + str(alpha) + r'}^3\times' + str(round(E * I, 1)) + r'}{' + str(Vx) + r'}\times' + str(round(vb / 1000, 3)) + '=' + str(Rha) + 'kN>H_k=' + str(round(Fxy, 1)) + 'kN'
    else:
        latex_Rha = JGJ94_2008_5_7_2_2 + r'=0.75\times\frac{{' + str(alpha) + r'}^3\times' + str(round(E * I, 1)) + r'}{' + str(Vx) + r'}\times' + str(round(vb / 1000, 3)) + '=' + str(Rha) + 'kN<H_k=' + str(round(Fxy, 1)) + 'kN'

    formula_registry = {
        '{JGJ94_Quk}': latex_Quk,
        '{JGJ94_Ra}': latex_Ra,
        '{JGJ94_Subgrade_Reaction_Coefficient}': latex_m,
        '{JGJ94_Displacement_Coefficient}': latex_alpha,
        '{JGJ94_Lateral_Capacity}': latex_Rha,
    }
    drivenplie['formula_registry'] = formula_registry

    print("[calc_driven_pile] 计算完成:", drivenplie)
    return drivenplie


def calc_drilled_shaft(excel_params, component_info, secondary_dialog_values, reaction_dict, pile_soil=None):
    """钻孔灌注桩计算总函数。

    参数:
        excel_params:           Excel参数表读取结果（含 soil_param）
        pile_soil:              匹配后的桩组土层信息（含 layers/ground_level/pile_bottom_depth）
        component_info:         模型截面/材料信息（含 各桩 vsize/L）
        secondary_dialog_values: UI二级窗口参数
        reaction_dict:          见下方格式说明

    reaction_dict 格式（来自 Post_processing_reaction 的 list6）:
        'Fz_max':   [序号, nodeid, 组合名, Fx, Fy, Fz, Mx, My, Mz]  最大竖向力
        'Fxy_max':  [序号, nodeid, 组合名, Fx, Fy, Fz, Mx, My, Mz]  最大水平力
        'pile_max': [ele, Fx, My, Mz]                                钢管桩最不利内力

    返回:
        drilledshaft: dict，预设 context 变量，字典结构后续可编辑
    """
    # ── 反力 ──
    Fz_max = reaction_dict.get('Fz_max', ['0', '0', '', 0, 0, 0, 0, 0, 0])
    Fxy_max = reaction_dict.get('Fxy_max', ['0', '0', '', 0, 0, 0, 0, 0, 0])
    reaction_nodeid = str(Fz_max[1])
    Fz = abs(float(Fz_max[5]))      # Fz 分量
    Fx = abs(float(Fxy_max[3]))     # Fx 分量（水平力最大工况）
    Fy = abs(float(Fxy_max[4]))     # Fy 分量
    Fxy = math.sqrt(Fx**2 + Fy**2)

    # ── 桩截面参数（根据反力节点号匹配所属桩组） ──
    pile_info = {}
    for key, info in component_info.items():
        if not key.endswith('钢管桩') or not isinstance(info, dict):
            continue
        node_ids = info.get('node_ids', [])
        if reaction_nodeid in node_ids:
            pile_info = info
            break
    if not pile_info:
        pile_key = next((k for k in component_info if k.endswith('钢管桩') and isinstance(component_info[k], dict)), None)
        if pile_key:
            pile_info = component_info[pile_key]

    vsize = pile_info.get('vsize', [])
    d = float(vsize[0]) if vsize else 0
    L = float(pile_info.get('L', 0))
    pile_bottom_depth = pile_soil.get('pile_bottom_depth', 0.0) if pile_soil else 0.0

    # ── 土层参数（必须由 pile_soil 按桩位匹配） ──
    if not pile_soil or not pile_soil.get('layers'):
        raise ValueError("[calc_drilled_shaft] pile_soil 为空，请检查 soil_param_index 是否正确传入")
    soil_layers = pile_soil['layers']
    print(f"[calc_drilled_shaft] 使用pile_soil: {len(soil_layers)}层, 桩底深={pile_soil.get('pile_bottom_depth')}m")
    s_layer_h = [layer.get('thickness', 0) for layer in soil_layers]
    s_qsk = [layer.get('qsik', 0) for layer in soil_layers]
    s_psk = [layer.get('qpk', 0) for layer in soil_layers]
    friction = [layer.get('phi', 0) for layer in soil_layers]
    Cohesion = [layer.get('c', 0) for layer in soil_layers]
    def _soil_category(name):
        """将土层名称映射为 JGJ94 公式需要的 '黏土'/'砂土' 类别。
        含'黏'→黏土，含'砂'→砂土，默认砂土（偏保守）。"""
        if '黏' in name:
            return '黏土'
        return '砂土'

    s_type = [_soil_category(layer.get('name', '')) for layer in soil_layers]

    # ── UI 参数 ──
    # 钻孔灌注桩 dict: {'concrete_grade', 'rebar_grade', 'pile_rho_g', 'cover_thick', 'x0a'}
    vals = secondary_dialog_values if isinstance(secondary_dialog_values, dict) else {}
    concrete_grade = str(vals.get('concrete_grade', 'C30'))
    rebar_grade = str(vals.get('rebar_grade', 'HRB400'))
    pg = float(vals.get('pile_rho_g', 0.65))       # 配筋率 %
    d_as = float(vals.get('cover_thick', 50))       # 保护层厚度 mm
    vb = float(vals.get('x0a', 10))                 # x0a = 水平位移允许值 mm
    N = Fz  # 竖向力用于 Rha1 公式

    # ── 材料参数 ──
    fc, ft, Ec = concrete_f(concrete_grade)
    fy, Es = rebar_f(rebar_grade)

    # ── 竖向承载力 JGJ94-2008 5.3.6 ──
    param_alst = Side_Resistance_Factor(s_type, d / 1000)
    param_blst = Base_Resistance_Factor(s_type, d / 1000)
    Qsk = JGJ94_Qsk_DrilledShaft(d / 1000, s_qsk, s_layer_h, param_alst)
    Qpk = JGJ94_Qpk_DrilledShaft(d / 1000, s_psk[-1] if s_psk else 0, param_blst)
    Quk = JGJ94_Quk(Qsk, Qpk)
    Ra = JGJ94_Ra(Quk)

    # ── 水平承载力 ──
    if d / 1000 > 1:
        b0 = 0.9 * (d / 1000 + 1)
    else:
        b0 = 0.9 * (1.5 * d / 1000 + 0.5)
    m_lst = Subgrade_Reaction_Coefficient(friction, Cohesion, vb)
    m = next((v for v in m_lst if v != 0), 1000)
    phi = friction[0] if friction else 0
    c = Cohesion[0] if Cohesion else 0

    if pg < 0.65:
        # 配筋率 <0.65%：JGJ94-2008 公式 5.7.2-1
        W0 = Section_Modulus_of_Transformed_Section_at_Tension_Edge_of_Pile_Shaft(d, Ec, Es, pg, d_as)
        An = Transformed_Section_Area_of_Pile_Shaft(d, Ec, Es, pg)
        E_pile = 0.85 * Ec * 1000  # kPa
        I = (W0 * 1e-9) * ((d - 2 * d_as) / 1000) / 2  # m4
        alpha = Displacement_Coefficient(m, b0, E_pile, I)
        Vm, Vx, alpha_h = JGJ94_Table_Vm_Vs(alpha, pile_bottom_depth)
        Rha = JGJ94_Rha1(alpha, 2, ft, W0, Vm, pg, N, An)
    else:
        # 配筋率 ≥0.65%：JGJ94-2008 公式 5.7.2-2
        E_pile = 0.85 * Ec * 1000
        I = math.pi * (d / 1000)**4 / 64
        alpha = Displacement_Coefficient(m, b0, E_pile, I)
        Vm, Vx, alpha_h = JGJ94_Table_Vm_Vs(alpha, pile_bottom_depth)
        Rha = JGJ94_Rha2(alpha, E_pile, I, Vx, vb)
        W0, An = 0, 0

    # ── 判定 ──
    if_Nk_satisfied = '满足' if Ra >= Fz else '不满足'
    if_Hk_satisfied = '满足' if Rha >= Fxy else '不满足'

    # ── 综合判定 ──
    type_satisfied = '满足' if if_Nk_satisfied == '满足' and if_Hk_satisfied == '满足' else '不满足'

    # ── 预设 context 字典（键名与模板 {{ drilledshaft.xxx }} 对应） ──
    drilledshaft = {
        # 模板直接引用
        'concrete_grade': concrete_grade,       # {{ drilledshaft.concrete_grade }}
        'd': d,                                 # {{ drilledshaft.d }}
        'as0': d_as,                            # {{ drilledshaft.as0 }}
        'pg': pg,                               # {{ drilledshaft.pg }}
        'h': L,                                 # {{ drilledshaft.h }}  桩长 m
        'Nk': round(Fz, 1),                     # {{ drilledshaft.Nk }}
        'Hk': round(Fxy, 1),                    # {{ drilledshaft.Hk }}
        'vb': vb,                               # {{ drilledshaft.vb }}
        'b0': round(b0, 3),                     # {{ drilledshaft.b0 }}
        'alpha_h': alpha_h,                     # {{ drilledshaft.alpha_h }}
        'vm': Vm,                               # {{ drilledshaft.vm }}
        'vx': Vx,                               # {{ drilledshaft.vx }}
        # type.satisfied 单独构建（模板用 {{ type.satisfied }}）
        # 计算中间值（供 LaTeX 或后续引用）
        'rebar_grade': rebar_grade,
        'Qsk': Qsk,
        'Qpk': Qpk,
        'Quk': Quk,
        'Ra': Ra,
        'phi': phi,
        'c': c,
        'm': round(m, 1),
        'EI': round(E_pile * I, 1),
        'alpha': alpha,
        'Es': Es,
        'Ec': Ec,
        'W0': W0,
        'An': An,
        'Rha': Rha,
        'if_Nk_satisfied': if_Nk_satisfied,
        'if_Hk_satisfied': if_Hk_satisfied,
    }
    # type.satisfied 模板用独立变量 {{ type.satisfied }}
    type_dict = {'satisfied': type_satisfied}

    # ── 公式 LaTeX 构建 ──
    # 竖向承载力公式（灌注桩用 JGJ94_2008_5_3_6）
    latex_Quk = JGJ94_2008_5_3_6 + '=' + str(Qsk) + '+' + str(Qpk) + '=' + str(Quk) + 'kN'
    if Ra >= Fz:
        latex_Ra = JGJ94_2008_5_2_2 + r'=\frac{1}{2}\times' + str(Quk) + '=' + str(Ra) + 'kN>N_k=' + str(round(Fz, 1)) + 'kN'
    else:
        latex_Ra = JGJ94_2008_5_2_2 + r'=\frac{1}{2}\times' + str(Quk) + '=' + str(Ra) + 'kN<N_k=' + str(round(Fz, 1)) + 'kN'
    # 水平承载力公式
    latex_m = JGJ_120_2012_4_1_6 + r'=\frac{0.2\times{' + str(phi) + r'}^2-' + str(phi) + '+' + str(c) + r'}{' + str(vb) + r'}\times1000=' + str(round(m, 1)) + 'kN/m^4'
    latex_alpha = JGJ94_2008_5_7_5 + r'=\sqrt[5]{\frac{' + str(round(m, 1)) + r'\times' + str(round(b0, 3)) + r'}{' + str(round(E_pile * I, 1)) + r'}}=' + str(alpha) + r'\left(1/m\right)'

    formula_registry = {
        '{JGJ94_Quk}': latex_Quk,
        '{JGJ94_Ra}': latex_Ra,
        '{JGJ94_Subgrade_Reaction_Coefficient}': latex_m,
        '{JGJ94_Displacement_Coefficient}': latex_alpha,
    }

    if pg < 0.65:
        # Rha1 路径：含 W0, An 公式
        latex_W0 = JGJ94_2008_5_7_2_1_W0 + r'=\frac{' + str(d) + r'\pi}{32}\times\left[{' + str(d) + r'}^2+2\left(\frac{' + str(Es) + '}{' + str(Ec) + r'}-1\right)\times' + str(pg / 100) + r'\times\left(' + str(d - d_as) + r'\right)^2\right]=' + str(W0) + 'mm^3'
        latex_An = JGJ94_2008_5_7_2_1_An + r'=\frac{{' + str(d) + r'}^2\pi}{4}\times\left[1+\left(\frac{' + str(Es) + '}{' + str(Ec) + r'}-1\right)\times' + str(pg / 100) + r'\right]=' + str(An) + 'mm^2'
        if Rha >= Fxy:
            latex_Rha = JGJ94_2008_5_7_2_1 + '=' + str(Rha) + 'kN>H_k=' + str(round(Fxy, 1)) + 'kN'
        else:
            latex_Rha = JGJ94_2008_5_7_2_1 + '=' + str(Rha) + 'kN<H_k=' + str(round(Fxy, 1)) + 'kN'
        formula_registry['{JGJ94_Transformed_Modulus}'] = latex_W0
        formula_registry['{JGJ94_Transformed_Area}'] = latex_An
    else:
        # Rha2 路径
        if Rha >= Fxy:
            latex_Rha = JGJ94_2008_5_7_2_2 + r'=0.75\times\frac{{' + str(alpha) + r'}^3\times' + str(round(E_pile * I, 1)) + r'}{' + str(Vx) + r'}\times' + str(round(vb / 1000, 3)) + '=' + str(Rha) + 'kN>H_k=' + str(round(Fxy, 1)) + 'kN'
        else:
            latex_Rha = JGJ94_2008_5_7_2_2 + r'=0.75\times\frac{{' + str(alpha) + r'}^3\times' + str(round(E_pile * I, 1)) + r'}{' + str(Vx) + r'}\times' + str(round(vb / 1000, 3)) + '=' + str(Rha) + 'kN<H_k=' + str(round(Fxy, 1)) + 'kN'

    formula_registry['{JGJ94_Lateral_Capacity}'] = latex_Rha
    drilledshaft['formula_registry'] = formula_registry

    print("[calc_drilled_shaft] 计算完成:", drilledshaft)
    return drilledshaft, type_dict


def calc_spread_foundation(secondary_dialog_values, reaction_dict, pile_diameter=0, soil_gs=None, h1=0):
    """扩大基础计算总函数。

    参数:
        secondary_dialog_values: UI二级窗口参数
        reaction_dict:          见下方格式说明
        pile_diameter:          钢管桩直径 (mm)，从模型获取，用于柱边等效尺寸 0.8D
        soil_gs:                土层加权平均容重 (kN/m³)，None 时使用默认值 18

    reaction_dict 格式（来自 Post_processing_reaction 的 list6）:
        'Fz_max': [序号, nodeid, 组合名, Fx, Fy, Fz, Mx, My, Mz]  最大竖向力
        'Mx_max': [序号, nodeid, 组合名, Fx, Fy, Fz, Mx, My, Mz]  最大Mx
        'My_max': [序号, nodeid, 组合名, Fx, Fy, Fz, Mx, My, Mz]  最大My

    返回:
        spread:              dict，扩大基础全部 context 变量（含地基承载力 + 结构计算）
        formula_registry: dict，{占位符: LaTeX} 用于 post-render 公式替换
    """
    # ── 反力（取最不利工况） ──
    # 格式: [序号, nodeid, 组合名, Fx, Fy, Fz, Mx, My, Mz]
    _default = ['0', '0', '', 0, 0, 0, 0, 0, 0]
    Fz_max = reaction_dict.get('Fz_max', _default)
    Mx_max = reaction_dict.get('Mx_max', _default)
    My_max = reaction_dict.get('My_max', _default)

    # 选择控制工况：比较 Mx 和 My 的绝对值
    if abs(float(My_max[7])) >= abs(float(Mx_max[6])):
        Fk1 = abs(float(Fz_max[5]))      # 最大竖向力工况的 Fz
        Fk2 = abs(float(My_max[5]))       # 最大My工况的 Fz
        Mk = abs(float(My_max[7]))        # My
        type_M = 'My'
    else:
        Fk1 = abs(float(Fz_max[5]))
        Fk2 = abs(float(Mx_max[5]))       # 最大Mx工况的 Fz
        Mk = abs(float(Mx_max[6]))        # Mx
        type_M = 'Mx'
    
    # ── 读取 UI 参数 ──
    # 扩大基础 dict: {'side_A', 'side_B', 'height', 'gc', 'fa', 'h2_top', 'theta', 'faz', 'has_weak',
    #                  'Ks', 'conc_grd', 'rebar_grd', 'cover', 'space_min', 'space_max', 'dia_min', 'd_max'}
    vals = secondary_dialog_values if isinstance(secondary_dialog_values, dict) else {}
    print(f"[calc_spread_foundation] vals={vals}")

    # ── 基础参数 ──
    x = float(vals.get('side_A', 0))     # 边长A
    y = float(vals.get('side_B', 0))     # 边长B
    z = float(vals.get('height', 0))     # 基础高度
    gc = float(vals.get('gc', 25))

    # ── 地基参数 ──
    fa = float(vals.get('fa', 0))
    # h1 由调用方从土层数据计算传入（基底标高 = 地面标高 - 桩入土深度）
    h2_raw = float(vals.get('h2_top', 0))  # Weak_Layer_Horizon（UI 原始值）
    angle_raw = float(vals.get('theta', 0))  # Dispersion_Angle（UI 原始值）
    faz_raw = float(vals.get('faz', 0))   # 下卧层承载力（UI 原始值）
    has_weak = bool(vals.get('has_weak', False))  # 是否考虑软弱下卧层
    gs = float(soil_gs) if soil_gs else 18.0  # 土容重 kN/m³（由 soil_param_index 加权平均计算）

    # h2_top=0 表示用户未填写软弱下卧层深度，强制关闭
    if has_weak and h2_raw == 0:
        has_weak = False

    # 根据 has_weak 决定是否启用软弱下卧层参数
    if has_weak:
        h2 = h2_raw
        angle = angle_raw
        faz = faz_raw
    else:
        h2 = 0
        angle = 0
        faz = 0

    # ── 结构强度参数 ──
    Ksi = float(vals.get('Ks', 1.35))
    concrete_grade = str(vals.get('conc_grd', 'C30'))
    rebar_grade = str(vals.get('rebar_grd', 'HRB400'))
    as0 = float(vals.get('cover', 50))
    at = 0   # 柱截面X向边长 (m)，扩大基础+钢管桩时由 D 等效
    bt = 0   # 柱截面Y向边长 (m)，同上
    D = float(pile_diameter) / 1000  # 钢管桩直径 m（模型取值mm→转m）
    fy_value, _ = rebar_f(rebar_grade)
    s1 = float(vals.get('space_min', 100))    # 钢筋间距下限 (mm)
    s2 = float(vals.get('space_max', 200))    # 钢筋间距上限 (mm)
    d1 = float(vals.get('dia_min', 12))       # 钢筋直径下限 (mm)
    d2 = float(vals.get('d_max', 12))         # 钢筋直径上限 (mm)

    # ── 地基承载力验算 ──
    Ak, Gk, pk_val, Mk_e, Mk_a, Eccentricity, Wk, pkmax, pkmin, pcz, pc, pz, \
        if_satisfied_fa1, if_satisfied_fa2, if_satisfied_faz, _ = \
        foundation_pressure(x, y, z, gc, Fk1, Fk2, Mk, h1, h2, gs, angle, fa, faz, type_M)

    print(f"Fk1={Fk1}, Fk2={Fk2}, Mk={Mk}")
    print(f"x={x}, y={y}, z={z}")
    print(f"pkmax={pkmax}")
    
    # ── 基础强度验算 ──
    ft = concrete_f(concrete_grade)[1]
    # 抗剪
    As1, As2, Vs1, Vs2, Bhs, A01, A02, Allowable_Vs1, Allowable_Vs2, if_Vs1, if_Vs2 = 0, 0, 0, 0, 0, 0, 0, 0, 0, '', ''
    print(f"[DEBUG 抗剪前] pkmax={pkmax}, type={type(pkmax)}")
    if z > 0 and as0 > 0 and ft:
        vs_result = Vs_Verification(x, y, z, as0, at, bt, D, pkmax, ft)
        if vs_result:
            As1, As2, Vs1, Vs2, Bhs, A01, A02, Allowable_Vs1, Allowable_Vs2, if_Vs1, if_Vs2 = vs_result

    # 抗冲切
    Al1, Al2, Fl1, Fl2, h0, Bhp, am_val, Allowable_Fl, if_Fl1, if_Fl2 = 0, 0, 0, 0, 0, 0, 0, 0, '', ''
    if z > 0 and as0 > 0 and ft:
        fl_result = Fl_Verification(x, y, z, as0, at, bt, D, pkmax, ft)
        if fl_result:
            Al1, Al2, Fl1, Fl2, h0, Bhp, am_val, Allowable_Fl, if_Fl1, if_Fl2 = fl_result

    # 抗弯配筋
    p = Ksi * (pk_val - Gk / Ak) if Ak > 0 else 0
    pmax = Ksi * (pkmax - Gk / Ak) if Ak > 0 else 0
    pmin = Ksi * (pkmin - Gk / Ak) if Ak > 0 and pkmin > 0 else 0
    print(f"[DEBUG 抗弯前] p={p}, pmax={pmax}, pmin={pmin}, pk_val={pk_val}, Gk/Ak={Gk/Ak if Ak>0 else 0}, Ksi={Ksi}")
    a1, aa, bb, M1, M2, As1, As2 = 0, 0, 0, 0, 0, 0, 0
    s1_val, d1_val, p1, A1, s2_val, d2_val, p2_val, A2 = 0, 0, 0, 0, 0, 0, 0, 0
    if_As1, if_As2, if_gouzao1, if_gouzao2 = '', '', '', ''
    if z > 0 and fy_value:
        m_result = M_Verification(x, y, z, as0, at, bt, D, p, pmax, pmin, fy_value, s1, s2, d1, d2, type_M)
        if m_result:
            a1, aa, bb, M1, M2, As1, As2, s1_val, d1_val, p1, A1, s2_val, d2_val, p2_val, A2, \
                if_As1, if_As2, if_gouzao1, if_gouzao2 = m_result

    # ── 模板分支判定字段 ──
    forcemode = '轴心荷载' if if_satisfied_fa2 == '' else '偏心荷载'
    has_weakSubstr = has_weak
    # Eccentricity: 模板期望"大偏心"/"小偏心"文字描述
    if Eccentricity == '大偏心':
        eccentricity_text = '大偏心'
    elif Eccentricity == '小偏心':
        eccentricity_text = '小偏心'
    else:
        eccentricity_text = ''

    # ── 结构计算综合判定（6 项：Vs1/Vs2/Fl1/Fl2/As1/As2） ──
    _all_sf = [if_Vs1, if_Vs2, if_Fl1, if_Fl2, if_As1, if_As2]
    if _all_sf and all(s == '满足' for s in _all_sf):
        if_foundation_Vs_FL_M_satisfied = '满足'
    elif any(s == '不满足' for s in _all_sf):
        if_foundation_Vs_FL_M_satisfied = '不满足'
    else:
        if_foundation_Vs_FL_M_satisfied = '满足'

    # ── 地基承载力综合判定 ──
    _bearing_sf = [if_satisfied_fa1]
    if if_satisfied_fa2:
        _bearing_sf.append(if_satisfied_fa2)
    if if_satisfied_faz:
        _bearing_sf.append(if_satisfied_faz)
    if all(s == '满足' for s in _bearing_sf):
        if_satisfied_foundation_pk = '满足'
    elif any(s == '不满足' for s in _bearing_sf):
        if_satisfied_foundation_pk = '不满足'
    else:
        if_satisfied_foundation_pk = '不满足'

    # ── 项目综合判定（地基承载力 + 结构验算，任一不满足即为不满足） ──
    _project_sf = [if_satisfied_foundation_pk, if_foundation_Vs_FL_M_satisfied]
    if any(s == '不满足' for s in _project_sf):
        project_satisfied = '不满足'
    elif all(s == '满足' for s in _project_sf):
        project_satisfied = '满足'
    else:
        project_satisfied = '满足'

    # ── 预设 context 字典（键名与模板 {{ spread.xxx }} 对应） ──
    spread = {
        # 承载力验算
        'forcemode': forcemode,                          # {{ pk.forcemode }}
        'has_weakSubstr': has_weakSubstr,                # {{ pk.has_weakSubstr }}
        'reaction_Fzmax': round(Fk1, 1),                 # {{ pk.reaction_Fzmax }}
        'biaozhun_reaction_Mymax': round(Mk, 1),         # {{ pk.biaozhun_reaction_Mymax }}
        'biaozhun_reaction_FFzmax': round(Fk2, 1),       # {{ pk.biaozhun_reaction_FFzmax }}
        'foundation_x': x,                               # {{ pk.foundation_x }}
        'foundation_y': y,                               # {{ pk.foundation_y }}
        'foundation_z': z,                               # {{ pk.foundation_z }}
        'Foundation_Base_Level': h1,                     # {{ pk.Foundation_Base_Level }}
        'foundation_fa': round(fa, 1),                   # {{ pk.foundation_fa }}
        'foundation_faz': round(faz, 1),                 # {{ pk.foundation_faz }}
        'Weak_Layer_Horizon': round(abs(h2), 1),         # {{ pk.Weak_Layer_Horizon }}
        'Dispersion_Angle': round(angle, 1),             # {{ pk.Dispersion_Angle }}
        'Eccentricity': eccentricity_text,               # {{ pk.Eccentricity }}
        'if_foundation_pk_satisfied': 
        if_satisfied_foundation_pk,                      # {{ pk.if_satisfied }}
        # 材料
        'concrete_grade': concrete_grade,                # {{ pk.concrete_grade }}
        'steel_protective_layer': round(as0, 1),         # {{ pk.steel_protective_layer }}
        # 结构计算综合判定
        'if_foundation_Vs_FL_M_satisfied': if_foundation_Vs_FL_M_satisfied,  # {{ pk.if_foundation_Vs_FL_M_satisfied }}
        # 项目综合判定
        'satisfied': project_satisfied,                     # {{ pk.satisfied }}
        # 计算值
        'Ak': Ak,
        'Gk': Gk,
        'pk': pk_val,
        'fa': fa,
        'if_satisfied_fa1': if_satisfied_fa1,
        'e': Mk_e,
        'a': Mk_a,
        'Wk': Wk,
        'pkmax': pkmax,
        'pkmin': pkmin,
        'if_satisfied_fa2': if_satisfied_fa2,
        'pcz': pcz,
        'pc': pc,
        'pz': pz,
        'faz': faz,
        'if_satisfied_faz': if_satisfied_faz,
        # ── 抗剪验算变量（模板 T0 表） ──
        'JCAv1': round(As1, 3),           # 受剪面积 Av1 (m²)
        'JCAv2': round(As2, 3),           # 受剪面积 Av2 (m²)
        'JCV1': round(Vs1, 1),            # 剪力设计值 V1 (kN)
        'JCV2': round(Vs2, 1),            # 剪力设计值 V2 (kN)
        'BHS1': round(Bhs, 2),            # 截面高度影响系数 βhs
        'BHS2': round(Bhs, 2),            # 同上（两个方向共用）
        'JCV_ft1': round(ft, 2),          # 混凝土抗拉强度 ft
        'JCV_ft2': round(ft, 2),          # 同上
        'JCA01': round(A01, 1),           # 有效面积 A01 (mm²)
        'JCA02': round(A02, 1),           # 有效面积 A02 (mm²)
        'JCV1_V': round(Allowable_Vs1, 1),   # 抗剪承载力允许值
        'JCV2_V': round(Allowable_Vs2, 1),   # 抗剪承载力允许值
        'JCV_sf1': if_Vs1,                # 抗剪判定1
        'JCV_sf2': if_Vs2,                # 抗剪判定2
        # ── 抗冲切验算变量（模板 T1 表） ──
        'JCAl1': round(Al1, 3),           # 冲切面积 Al1 (m²)
        'JCAl2': round(Al2, 3),           # 冲切面积 Al2 (m²)
        'JCFl1': round(Fl1, 1),           # 冲切力 Fl1 (kN)
        'JCFl2': round(Fl2, 1),           # 冲切力 Fl2 (kN)
        'BHP1': round(Bhp, 3),            # 截面高度影响系数 βhp
        'BHP2': round(Bhp, 3),            # 同上
        'JCF_ft1': round(ft, 2),          # 混凝土抗拉强度 ft
        'JCF_ft2': round(ft, 2),          # 同上
        'JCF_am1': round(am_val, 1),      # 冲切破坏锥体最不利一侧计算长度 am
        'JCF_am2': round(am_val, 1),      # 同上
        'JC_h01': round(h0, 1),           # 有效高度 h0 (mm)
        'JC_h02': round(h0, 1),           # 同上
        'JCF1_F': round(Allowable_Fl, 1), # 抗冲切承载力允许值
        'JCF2_F': round(Allowable_Fl, 1), # 同上
        'JCF_sf1': if_Fl1,                # 抗冲切判定1
        'JCF_sf2': if_Fl2,                # 抗冲切判定2
        # ── 弯矩计算变量（模板 T2 表） ──
        'JC_a1': round(a1, 3),            # 悬挑长度 a1 (m)
        'JC_aa': round(aa, 3),            # 扩散长度 aa (m)
        'JC_bb': round(bb, 3),            # 扩散长度 bb (m)
        'JC_M1': round(M1, 1),            # 弯矩 M1 (kN·m)
        'JC_M2': round(M2, 1),            # 弯矩 M2 (kN·m)
        # ── 抗弯配筋变量（模板 T3 表） ──
        'rebar_grade': rebar_grade,       
        'JCM1': round(M1, 1),             # 弯矩设计值 M1
        'JCM2': round(M2, 1),             # 弯矩设计值 M2
        'MAs1': round(As1, 1),            # 需要钢筋面积 As1 (mm²/m)
        'MAs2': round(As2, 1),            # 需要钢筋面积 As2 (mm²/m)
        'JCGJ1': f'E{d1_val}@{s1_val}',   # 配筋规格1（如 E12@150）
        'JCGJ2': f'E{d2_val}@{s2_val}',   # 配筋规格2
        'JCAs1': round(A1, 1),            # 实际配筋面积1 (mm²/m)
        'JCAs2': round(A2, 1),            # 实际配筋面积2 (mm²/m)
        'JCGJ_p1': round(p1, 3),          # 配筋率1
        'JCGJ_p2': round(p2_val, 3),       # 配筋率2
        'JCM_if1': if_gouzao1,            # 构造配筋判定1
        'JCM_if2': if_gouzao2,            # 构造配筋判定2
        'JCM_sf1': if_As1,                # 配筋判定1
        'JCM_sf2': if_As2,                # 配筋判定2
    }

    # ── 公式 LaTeX 构建 ──
    hs = max(round(abs(h1) - z, 1), 0)

    # --- 地基承载力公式 ---
    # {foundation_Gk}
    if hs != 0:
        latex_Gk = 'G_k=' + str(x) + r'\times' + str(y) + r'\times' + str(z) + r'\times' + str(gc) + '+' + str(x) + r'\times' + str(y) + r'\times' + str(hs) + r'\times' + str(gs) + f'={Gk}kN'
    else:
        latex_Gk = 'G_k=' + str(x) + r'\times' + str(y) + r'\times' + str(z) + r'\times' + str(gc) + f'={Gk}kN'
    # {foundation_pk}
    latex_pk = GB50007_2011_5_2_2_1 + r'=\frac{\left(' + f'{Fk1}+{Gk}' + r'\right)}{' + str(Ak) + r'}=' + f'{pk_val}kPa'
    # {if_satisfied_foundation_pk}
    if if_satisfied_fa1 == "":
        latex_if_pk = ''
    elif if_satisfied_fa1 == "满足":
        latex_if_pk = f"p_k={pk_val}kPa<f_a={fa}kPa"
    else:
        latex_if_pk = f"p_k={pk_val}kPa>f_a={fa}kPa"

    # --- 偏心受压公式 ---
    if if_satisfied_fa2 != "":
        b_val = y  # 基础宽度
        if Mk_e >= b_val / 6:
            latex_ecc = GB50007_2011_Eccentricity + r'=\frac{' + str(Mk) + r'}{\left(' + f'{Fk2}+{Gk}' + r'\right)}=' + str(Mk_e) + r'>\frac{' + str(b_val) + r'}{6}'
            latex_pkmax = GB50007_2011_5_2_2_4 + r'=\frac{2\times\left(' + f'{Fk2}+{Gk}' + r'\right)}{3\times' + str(x) + r'\times' + str(Mk_a) + r'}=' + str(pkmax) + 'kPa'
            latex_pkmin = r'p_{kmin}=0kPa'
        else:
            latex_ecc = GB50007_2011_Eccentricity + r'=\frac{' + str(Mk) + r'}{\left(' + f'{Fk2}+{Gk}' + r'\right)}=' + str(Mk_e) + r'<\frac{' + str(b_val) + r'}{6}'
            latex_pkmax = GB50007_2011_5_2_2_2 + r'=\frac{\left(' + f'{Fk2}+{Gk}' + r'\right)}{' + str(Ak) + r'}+\frac{' + str(Mk) + '}{' + str(Wk) + '}=' + f'{pkmax}kPa'
            latex_pkmin = GB50007_2011_5_2_2_3 + r'=\frac{\left(' + f'{Fk2}+{Gk}' + r'\right)}{' + str(Ak) + r'}-\frac{' + str(Mk) + '}{' + str(Wk) + '}=' + f'{pkmin}kPa'
        # {if_satisfied_foundation_pkmax}
        if if_satisfied_fa2 == "满足":
            latex_if_pkmax = r'p_{kmax}=' + f'{pkmax}kPa<1.2f_a={round(1.2 * fa, 1)}kPa'
        else:
            latex_if_pkmax = r'p_{kmax}=' + f'{pkmax}kPa>1.2f_a={round(1.2 * fa, 1)}kPa'
    else:
        latex_ecc = ''
        latex_pkmax = ''
        latex_pkmin = ''
        latex_if_pkmax = ''

    # --- 软弱下卧层公式 ---
    if if_satisfied_faz != "":
        latex_pcz = r'p_{cz}=\gamma h_{cz}=' + str(gs) + r'\times' + f'{abs(h2)}={pcz}kPa'
        if hs != 0:
            latex_pc = r'p_c=\gamma h_c=' + str(gs) + r'\times' + f'{abs(hs)}={pc}kPa'
        else:
            latex_pc = r'p_c=\gamma\ h_c=0kPa'
        latex_pz = GB50007_2011_5_2_7_3 + f'={pz}kPa'
        if if_satisfied_faz == "满足":
            latex_if_pz = r'p_z+p_{cz}=' + f'{pz}+{pcz}={round(pz + pcz, 1)}' + r'kPa<f_{az}=' + f'{faz}kPa'
        else:
            latex_if_pz = r'p_z+p_{cz}=' + f'{pz}+{pcz}={round(pz + pcz, 1)}' + r'kPa>f_{az}=' + f'{faz}kPa'
    else:
        latex_pcz = ''
        latex_pc = ''
        latex_pz = ''
        latex_if_pz = ''

    # --- 基底净反力公式 ---
    latex_pj = GB50007_2011_5_2_8_pj + '=' + str(Ksi) + r'\times\left(' + str(pk_val) + r'-\frac{' + str(Gk) + '}{' + str(x) + r'\times' + str(y) + r'}\right)=' + f'{p}kPa'
    latex_pjmax = GB50007_2011_5_2_11_pjmax + '=' + str(Ksi) + r'\times\left(' + str(pkmax) + r'-\frac{' + str(Gk) + '}{' + str(x) + r'\times' + str(y) + r'}\right)=' + f'{pmax}kPa'
    if pmin > 0:
        latex_pjmin = GB50007_2011_5_2_11_pjmin + '=' + str(Ksi) + r'\times\left(' + str(pkmin) + r'-\frac{' + str(Gk) + '}{' + str(x) + r'\times' + str(y) + r'}\right)=' + f'{pmin}kPa'
    else:
        latex_pjmin = r'p_{jmin}=0kPa'

    formula_registry = {
        # 地基承载力公式
        '{foundation_Gk}': latex_Gk,
        '{foundation_pk}': latex_pk,
        '{if_satisfied_foundation_pk}': latex_if_pk,
        '{foundation_Eccentricity}': latex_ecc,
        '{foundation_pkmax}': latex_pkmax,
        '{foundation_pkmin}': latex_pkmin,
        '{if_satisfied_foundation_pkmax}': latex_if_pkmax,
        '{foundation_pcz}': latex_pcz,
        '{foundation_pc}': latex_pc,
        '{foundation_pz}': latex_pz,
        '{if_satisfied_foundation_pk_pz}': latex_if_pz,
        # 基底净反力公式
        '{spread_foundation_pj}': latex_pj,
        '{spread_foundation_pjmax}': latex_pjmax,
        '{spread_foundation_pjmin}': latex_pjmin,
    }

    print("[calc_spread_foundation] 计算完成")
    print("  spread:", spread)
    return spread, formula_registry


def calc_strip_foundation(secondary_dialog_values, reaction_dict, pile_diameter=0, soil_gs=None, h1=0):
    """条形基础计算总函数。

    与扩大基础的主要区别：
    - 长度按 1m 单位长度计算（y = 1.0）
    - 只需计算一个方向的抗剪（上下方向）
    - 不计算抗冲切
    - 抗弯公式：M = 1/6 × a1² × (2pmax + p - 3G/A)

    参数:
        secondary_dialog_values: UI二级窗口参数
        reaction_dict:          荷载反力数据（同 calc_spread_foundation）
        pile_diameter:          钢管桩直径 (mm)
        soil_gs:                土层加权平均容重 (kN/m³)，None 时使用默认值 18

    返回:
        strip:              dict，条形基础全部 context 变量（含地基承载力 + 结构计算）
        formula_registry:   dict，{占位符: LaTeX} 用于 post-render 公式替换
    """
    # 反力（取最不利工况）
    _default = ['0', '0', '', 0, 0, 0, 0, 0, 0]
    Fz_max = reaction_dict.get('Fz_max', _default)
    Mx_max = reaction_dict.get('Mx_max', _default)
    My_max = reaction_dict.get('My_max', _default)

    if abs(float(My_max[7])) >= abs(float(Mx_max[6])):
        Fk1 = abs(float(Fz_max[5]))
        Fk2 = abs(float(My_max[5]))
        Mk = abs(float(My_max[7]))
        type_M = 'My'
    else:
        Fk1 = abs(float(Fz_max[5]))
        Fk2 = abs(float(Mx_max[5]))
        Mk = abs(float(Mx_max[6]))
        type_M = 'Mx'

    # ── 读取 UI 参数 ──
    # 条形基础 dict: {'width_b', 'height', 'gc', 'fa', 'h2_top', 'theta', 'faz', 'has_weak',
    #                  'Ks', 'conc_grd', 'rebar_grd', 'cover', 'space_min', 'space_max', 'dia_min', 'dia_max'}
    vals = secondary_dialog_values if isinstance(secondary_dialog_values, dict) else {}

    x = float(vals.get('width_b', 0))    # 基础宽度 b
    y = 1.0                               # 单位长度
    z = float(vals.get('height', 0))      # 基础高度 h
    gc = float(vals.get('gc', 25))

    fa = float(vals.get('fa', 0))
    # h1 由调用方从土层数据计算传入（基底标高 = 地面标高 - 桩入土深度）
    h2_raw = float(vals.get('h2_top', 0))
    angle_raw = float(vals.get('theta', 0))
    faz_raw = float(vals.get('faz', 0))
    has_weak = bool(vals.get('has_weak', False))
    gs = float(soil_gs) if soil_gs else 18.0  # 土容重 kN/m³（由 soil_param_index 加权平均计算）
    # h2_top=0 表示用户未填写软弱下卧层深度，强制关闭
    if has_weak and h2_raw == 0:
        has_weak = False
    if has_weak:
        h2, angle, faz = h2_raw, angle_raw, faz_raw
    else:
        h2, angle, faz = 0, 0, 0

    Ksi = float(vals.get('Ks', 1.35))
    concrete_grade = str(vals.get('conc_grd', 'C30'))
    rebar_grade = str(vals.get('rebar_grd', 'HRB400'))
    as0 = float(vals.get('cover', 50))
    at, bt = 0, 0
    D = float(pile_diameter) / 1000  # 桩径 m（模型取值mm→转m）
    fy_value, _ = rebar_f(rebar_grade)
    s1 = float(vals.get('space_min', 100))
    s2 = float(vals.get('space_max', 200))
    d1 = float(vals.get('dia_min', 12))
    d2 = float(vals.get('dia_max', 25))

    # ── 地基承载力验算 ── 
    Ak, Gk, pk_val, Mk_e, Mk_a, Eccentricity, Wk, pkmax, pkmin, pcz, pc, pz, \
        if_satisfied_fa1, if_satisfied_fa2, if_satisfied_faz, _ = \
        foundation_pressure(x, y, z, gc, Fk1, Fk2, Mk, h1, h2, gs, angle, fa, faz, type_M)

    ft = concrete_f(concrete_grade)[1]

    # ── 基础结构验算 ──
    # 抗剪
    As1, As2, Vs1, Vs2, Bhs, A01, A02, Allowable_Vs1, Allowable_Vs2, if_Vs1, if_Vs2 = \
        0, 0, 0, 0, 0, 0, 0, 0, 0, '', ''
    if z > 0 and as0 > 0 and ft:
        vs_result = Vs_Verification(x, y, z, as0, at, bt, D, pkmax, ft)
        if vs_result:
            As1, As2, Vs1, Vs2, Bhs, A01, A02, Allowable_Vs1, Allowable_Vs2, if_Vs1, if_Vs2 = vs_result

    # 不计算抗冲切

    # 抗弯配筋
    p = Ksi * (pk_val - Gk / Ak) if Ak > 0 else 0
    pmax = Ksi * (pkmax - Gk / Ak) if Ak > 0 else 0
    pmin = Ksi * (pkmin - Gk / Ak) if Ak > 0 and pkmin > 0 else 0
    a1, aa, bb, M1, As_calc = 0, 0, 0, 0, 0
    s1_val, d1_val, p1, A1 = 0, 0, 0, 0
    if_As1, if_gouzao1 = '', ''
    if z > 0 and fy_value:
        a1, aa, bb = cal_M_param(x, y, at, bt, D, type_M)
        # 条形基础抗弯公式：M = 1/6 × a1² × (2pmax + p - 3G/A)
        G_over_A = Gk / Ak if Ak > 0 else 0
        M1 = round(1.0 / 6.0 * a1 ** 2 * (2 * pmax + p - 3 * G_over_A), 1)
        As_calc = max(cal_As(M1, fy_value, x, y, z, as0, type_M), min_p_As(z, as0))
        rebar_result = cal_rebar_arrangement(z, as0, s1, s2, d1, d2, As_calc)
        if rebar_result:
            s1_val, d1_val, p1, A1 = rebar_result
            if_As1 = "满足" if A1 >= As_calc else "不满足"
            if_gouzao1 = "是" if s1_val == 200 else "否"

    forcemode = '轴心荷载' if if_satisfied_fa2 == '' else '偏心荷载'
    has_weakSubstr = has_weak
    eccentricity_text = Eccentricity if Eccentricity in ('大偏心', '小偏心') else ''

    _all_sf = [if_Vs1, if_As1]
    if _all_sf and all(s == '满足' for s in _all_sf):
        if_foundation_Vs_FL_M_satisfied = '满足'
    elif any(s == '不满足' for s in _all_sf):
        if_foundation_Vs_FL_M_satisfied = '不满足'
    else:
        if_foundation_Vs_FL_M_satisfied = '满足'

    _bearing_sf = [if_satisfied_fa1]
    if if_satisfied_fa2: _bearing_sf.append(if_satisfied_fa2)
    if if_satisfied_faz: _bearing_sf.append(if_satisfied_faz)
    if all(s == '满足' for s in _bearing_sf):
        if_satisfied_foundation_pk = '满足'
    elif any(s == '不满足' for s in _bearing_sf):
        if_satisfied_foundation_pk = '不满足'
    else:
        if_satisfied_foundation_pk = '不满足'

    # ── 项目综合判定（地基承载力 + 结构验算，任一不满足即为不满足） ──
    _project_sf = [if_satisfied_foundation_pk, if_foundation_Vs_FL_M_satisfied]
    if any(s == '不满足' for s in _project_sf):
        project_satisfied = '不满足'
    elif all(s == '满足' for s in _project_sf):
        project_satisfied = '满足'
    else:
        project_satisfied = '满足'

    strip = {
        'forcemode': forcemode, 'has_weakSubstr': has_weakSubstr,
        'reaction_Fzmax': round(Fk1, 1), 'biaozhun_reaction_Mymax': round(Mk, 1),
        'biaozhun_reaction_FFzmax': round(Fk2, 1),
        'foundation_x': round(x, 3), 'foundation_z': round(z, 3),
        'Foundation_Base_Level': round(h1, 1), 'foundation_fa': round(fa, 1),
        'foundation_faz': round(faz, 1), 'Weak_Layer_Horizon': round(abs(h2), 1),
        'Dispersion_Angle': round(angle, 1), 'Eccentricity': eccentricity_text,
        'if_foundation_pk_satisfied': if_satisfied_foundation_pk,
        'concrete_grade': concrete_grade, 'steel_protective_layer': round(as0, 1),
        'rebar_grade': rebar_grade,
        'if_foundation_Vs_FL_M_satisfied': if_foundation_Vs_FL_M_satisfied,
        'satisfied': project_satisfied,
        'Ak': Ak, 'Gk': Gk, 'pk': pk_val, 'fa': fa,
        'if_satisfied_fa1': if_satisfied_fa1, 'e': Mk_e, 'a': Mk_a, 'Wk': Wk,
        'pkmax': pkmax, 'pkmin': pkmin, 'if_satisfied_fa2': if_satisfied_fa2,
        'pcz': pcz, 'pc': pc, 'pz': pz, 'faz': faz, 'if_satisfied_faz': if_satisfied_faz,
        'JCAv1': round(As1, 3), 'JCV1': round(Vs1, 1), 'BHS1': round(Bhs, 2),
        'JCV_ft1': round(ft, 2), 'JCA01': round(A01, 1), 'JCV1_V': round(Allowable_Vs1, 1),
        'JCV_sf1': if_Vs1,
        'JC_a1': round(a1, 3), 'JC_M1': round(M1, 1),
        'JCM1': round(M1, 1), 'MAs1': round(As_calc, 1),
        'JCGJ1': f'E{d1_val}@{s1_val}', 'JCAs1': round(A1, 1),
        'JCGJ_p1': round(p1, 3), 'JCM_if1': if_gouzao1, 'JCM_sf1': if_As1,
    }

    # 公式 LaTeX 构建
    hs = max(round(abs(h1) - z, 1), 0)
    if hs != 0:
        latex_Gk = 'G_k=' + str(x) + r'\times' + str(y) + r'\times' + str(z) + r'\times' + str(gc) + '+' + str(x) + r'\times' + str(y) + r'\times' + str(hs) + r'\times' + str(gs) + f'={Gk}kN'
    else:
        latex_Gk = 'G_k=' + str(x) + r'\times' + str(y) + r'\times' + str(z) + r'\times' + str(gc) + f'={Gk}kN'
    latex_pk = GB50007_2011_5_2_2_1 + r'=\frac{\left(' + f'{Fk1}+{Gk}' + r'\right)}{' + str(Ak) + r'}=' + f'{pk_val}kPa'
    latex_if_pk = '' if if_satisfied_fa1 == '' else (f"p_k={pk_val}kPa<f_a={fa}kPa" if if_satisfied_fa1 == '满足' else f"p_k={pk_val}kPa>f_a={fa}kPa")

    if if_satisfied_fa2 != '':
        if Mk_e >= x / 6:
            latex_ecc = GB50007_2011_Eccentricity + r'=\frac{' + str(Mk) + r'}{\left(' + f'{Fk2}+{Gk}' + r'\right)}=' + str(Mk_e) + r'>\frac{' + str(x) + r'}{6}'
            latex_pkmax = GB50007_2011_5_2_2_4 + r'=\frac{2\times\left(' + f'{Fk2}+{Gk}' + r'\right)}{3\times' + str(x) + r'\times' + str(Mk_a) + r'}=' + str(pkmax) + 'kPa'
            latex_pkmin = r'p_{kmin}=0kPa'
        else:
            latex_ecc = GB50007_2011_Eccentricity + r'=\frac{' + str(Mk) + r'}{\left(' + f'{Fk2}+{Gk}' + r'\right)}=' + str(Mk_e) + r'<\frac{' + str(x) + r'}{6}'
            latex_pkmax = GB50007_2011_5_2_2_2 + r'=\frac{\left(' + f'{Fk2}+{Gk}' + r'\right)}{' + str(Ak) + r'}+\frac{' + str(Mk) + '}{' + str(Wk) + '}=' + f'{pkmax}kPa'
            latex_pkmin = GB50007_2011_5_2_2_3 + r'=\frac{\left(' + f'{Fk2}+{Gk}' + r'\right)}{' + str(Ak) + r'}-\frac{' + str(Mk) + '}{' + str(Wk) + '}=' + f'{pkmin}kPa'
        latex_if_pkmax = r'p_{kmax}=' + f'{pkmax}kPa' + ('<1.2f_a=' if if_satisfied_fa2 == '满足' else '>1.2f_a=') + f'{round(1.2 * fa, 1)}kPa'
    else:
        latex_ecc = latex_pkmax = latex_pkmin = latex_if_pkmax = ''

    if if_satisfied_faz != '':
        latex_pcz = r'p_{cz}=\gamma h_{cz}=' + str(gs) + r'\times' + f'{abs(h2)}={pcz}kPa'
        latex_pc = r'p_c=\gamma h_c=' + str(gs) + r'\times' + f'{abs(hs)}={pc}kPa' if hs != 0 else r'p_c=\gamma\ h_c=0kPa'
        latex_pz = GB50007_2011_5_2_7_3 + f'={pz}kPa'
        latex_if_pz = r'p_z+p_{cz}=' + f'{pz}+{pcz}={round(pz + pcz, 1)}' + r'kPa' + ('<f_{az}=' if if_satisfied_faz == '满足' else '>f_{az}=') + f'{faz}kPa'
    else:
        latex_pcz = latex_pc = latex_pz = latex_if_pz = ''

    latex_pj = GB50007_2011_5_2_8_pj + '=' + str(Ksi) + r'\times\left(' + str(pk_val) + r'-\frac{' + str(Gk) + '}{' + str(x) + r'\times' + str(y) + r'}\right)=' + f'{p}kPa'
    latex_pjmax = GB50007_2011_5_2_11_pjmax + '=' + str(Ksi) + r'\times\left(' + str(pkmax) + r'-\frac{' + str(Gk) + '}{' + str(x) + r'\times' + str(y) + r'}\right)=' + f'{pmax}kPa'
    latex_pjmin = (GB50007_2011_5_2_11_pjmin + '=' + str(Ksi) + r'\times\left(' + str(pkmin) + r'-\frac{' + str(Gk) + '}{' + str(x) + r'\times' + str(y) + r'}\right)=' + f'{pmin}kPa') if pmin > 0 else r'p_{jmin}=0kPa'

    # latex_M = (GB50007_2011_8_2_12_M
    #     + r'=\frac{1}{6}a_1^2\left(2p_{max}+p_j-\frac{3G}{A}\right)='
    #     + r'\frac{1}{6}\times' + f'{a1}^2\\times(2\\times{pmax}+{p}-3\\times{round(G_over_A, 1)})='
    #     + f'{M1}kN' + r'\cdot m')

    formula_registry = {
        '{foundation_Gk}': latex_Gk, 
        '{foundation_pk}': latex_pk,
        '{if_satisfied_foundation_pk}': latex_if_pk,
        '{foundation_Eccentricity}': latex_ecc, 
        '{foundation_pkmax}': latex_pkmax,
        '{foundation_pkmin}': latex_pkmin, 
        '{if_satisfied_foundation_pkmax}': latex_if_pkmax,
        '{foundation_pcz}': latex_pcz, 
        '{foundation_pc}': latex_pc,
        '{foundation_pz}': latex_pz, 
        '{if_satisfied_foundation_pk_pz}': latex_if_pz,
        '{sf_foundation_pj}': latex_pj, 
        '{sf_foundation_pjmax}': latex_pjmax,
        '{sf_foundation_pjmin}': latex_pjmin,
    }

    print("[calc_strip_foundation] 计算完成")
    print("  strip:", strip)
    return strip, formula_registry

def match_pile_soil(reaction_dict, component_info, soil_param_index, soil_param, deck_level):
    """根据最不利反力节点匹配桩组土层，计算桩底深度，合并土层参数。

    参数:
        reaction_dict:       {'Fz_max': [...], 'Fxy_max': [...]}
        component_info:      模型截面/材料信息
        soil_param_index:    按桩号分组的土层列表（来自 excel_params）
        soil_param:          全局土层参数列表（来自 excel_params，含 qsik/qpk 等）
        deck_level:          桥面高程 (m)

    返回:
        pile_soil: dict，包含：
            'layers':       [{'name', 'thickness', 'weight', 'c', 'phi', 'qsik', 'qpk'}, ...]
            'ground_level': 地面标高 (m)
            'pile_bottom_depth': 桩底距地面深度 (m)
        若无法匹配则返回 None
    """
    if not soil_param_index:
        print("[match_pile_soil] soil_param_index 为空，跳过")
        return None

    # 用 match_soil_entry 按模型桩号匹配土层条目
    spi, pile_key = match_soil_entry(reaction_dict, component_info, soil_param_index)
    if spi is None:
        print("[match_pile_soil] 无法匹配土层条目")
        return None

    # 从 component_info 取桩底标高
    pile_bottom_level = 0
    if pile_key and pile_key in component_info:
        pile_bottom_level = component_info[pile_key].get('pile_bottom_level', 0)
    ground_level = float(spi.get('ground_level', 0))

    # 桩底绝对标高 = 桥面高程 + 模型桩底相对标高(mm→m)
    pile_bottom_abs = float(deck_level) + pile_bottom_level / 1000
    pile_bottom_depth = round(ground_level - pile_bottom_abs, 3)
    m_idx = re.match(r'^(\d+)-', pile_key) if pile_key else None
    pile_idx = int(m_idx.group(1)) if m_idx else '?'
    print(f"[match_pile_soil] {pile_key}: 桩号索引{pile_idx}, 地面={ground_level}m, "
          f"桩底绝对={pile_bottom_abs}m, 桩底深={pile_bottom_depth}m")

    # 构建 soil_param 属性查找表 {name: {weight, c, phi, qsik, qpk}}
    soil_prop_map = {}
    for sp in (soil_param or []):
        soil_prop_map[sp.get('name', '')] = sp

    # 合并土层名称 → 对应属性（统一转 float，防止 Excel 文本格式）
    def _f(v):
        try:
            return float(v) if v not in (None, '') else 0
        except (ValueError, TypeError):
            return 0

    layers = []
    for j, name in enumerate(spi.get('name', [])):
        prop = soil_prop_map.get(name, {})
        layer = {
            'name':          name,
            'thickness':     _f(spi['thickness'][j]) if j < len(spi.get('thickness', [])) else 0,
            'weight':        _f(prop.get('weight')),
            'c':             _f(prop.get('c')),
            'phi':           _f(prop.get('phi')),
            'qsik':          _f(prop.get('qsik')),
            'qpk':           _f(prop.get('qpk')),
        }
        layers.append(layer)

    pile_soil = {
        'layers':             layers,
        'ground_level':       ground_level,
        'pile_bottom_depth':  pile_bottom_depth,
    }
    print(f"[match_pile_soil] 土层数={len(layers)}, 各层={[l['name'] for l in layers]}")
    return pile_soil


# =====================================================================================================================================================================
# 基础计算总函数——用于 docxtpl 模板渲染
# 输入：excel_params（含土层信息）、component_info（含管桩信息）、secondary_dialog_values（UI参数）、reaction_dict（反力）
# 输出：context 字典，供 docxtpl 渲染

def foundation_subdoc_render(foundation_type, foundation_standard, foundation_value_dict, model_post_dict,
                   excel_params, component_info, docx_open_path, main_tpl, pile_diameter=0, material=None,
                   output_dir='', midas_path='', soil_param_index=None):
    """基础计算分发函数。

    根据基础类型和规范调用对应的计算函数，渲染基础计算子模板并返回 subdoc 和公式注册表。

    支持的基础类型:
        - "打入桩"  → calc_driven_pile
        - "钻孔灌注桩"  → calc_drilled_shaft
        - "扩大基础"    → calc_spread_foundation
        - "条形基础"    → calc_strip_foundation

    参数:
        foundation_type:        基础类型名称
        foundation_standard:    计算规范名称（如 "建筑地基基础设计规范"）
        foundation_value_dict:  UI二级窗口参数字典（键为基础类型名，值为参数列表）
        model_post_dict:        Post_processing_Midas_Model 返回的结果字典
        excel_params:           Excel参数表读取结果（含 soil_param）
        component_info:         模型截面/材料信息（含 各桩 vsize/L）
        docx_open_path:         模板目录路径
        main_tpl:               主模板 DocxTemplate 实例（用于创建 subdoc）
        pile_diameter:          钢管桩直径 (mm)
        material:               材料字典（含 steel_lst/concrete_lst/rebar_lst）
        output_dir:             临时文件输出目录（默认与模板目录相同）

    返回:
        (subdoc, formula_registry)
        subdoc:             基础计算子模板渲染结果，插入主模板 context
        formula_registry:   dict，{占位符: LaTeX} 用于 post-render 公式替换
    """
    if foundation_type == "无":
        print("基础形式为无，跳过基础计算")
        return '', {}, ''

    if excel_params is None:
        excel_params = {}

    # ── 反力数据 ──
    model_post_reaction = model_post_dict.get("model_post_reaction", [])
    _default_react = ['0', '0', '', 0, 0, 0, 0, 0, 0]

    # ── 分发计算 ──
    print(f"[foundation_subdoc_render] foundation_type={foundation_type}, foundation_standard={foundation_standard}")
    print(f"[foundation_subdoc_render] foundation_value_dict={foundation_value_dict}")
    # UI 未打开二级窗口时值为 None，统一替换为空 dict 防止 .get() 报错
    if foundation_value_dict is None:
        foundation_value_dict = {}
    else:
        foundation_value_dict = {k: (v if v is not None else {}) for k, v in foundation_value_dict.items()}
    context_dict = {}
    type_dict = {}
    formula_registry = {}
    foundation_satisfied = ''  # 基础验算综合判定，由各分支赋值

    if foundation_type == "打入桩":
        if foundation_standard != "建筑桩基技术规范":
            print(f"打入桩规范应为建筑桩基技术规范，当前: {foundation_standard}")
            return '', {}, ''

        reaction_dict = {
            'Fz_max': model_post_reaction[0] if len(model_post_reaction) > 0 else _default_react,
            'Fxy_max': model_post_reaction[1] if len(model_post_reaction) > 1 else _default_react,
        }
        gangguanzhuang = model_post_dict.get("model_post_gangguanzhuang", [0, 0, 0, 0])
        reaction_dict['pile_max'] = gangguanzhuang[:4] if len(gangguanzhuang) >= 4 else [0, 0, 0, 0]

        # ── 匹配最不利反力桩组 → 取对应土层 ──
        pile_soil = match_pile_soil(
            reaction_dict, component_info, soil_param_index,
            excel_params.get('soil_param', []), excel_params.get('deck_level', 0)
        )

        secondary_dialog_values = foundation_value_dict.get(foundation_type, {})
        context_dict = calc_driven_pile(
            excel_params, component_info, secondary_dialog_values, reaction_dict, material, pile_soil
        )
        formula_registry = context_dict.pop('formula_registry', {})
        foundation_satisfied = context_dict.get('satisfied', '')


    elif foundation_type == "钻孔灌注桩":
        if foundation_standard != "建筑桩基技术规范":
            print(f"钻孔灌注桩规范应为建筑桩基技术规范，当前: {foundation_standard}")
            return '', {}, ''

        reaction_dict = {
            'Fz_max': model_post_reaction[0] if len(model_post_reaction) > 0 else _default_react,
            'Fxy_max': model_post_reaction[1] if len(model_post_reaction) > 1 else _default_react,
        }
        gangguanzhuang = model_post_dict.get("model_post_gangguanzhuang", [0, 0, 0, 0])
        reaction_dict['pile_max'] = gangguanzhuang[:4] if len(gangguanzhuang) >= 4 else [0, 0, 0, 0]

        # ── 匹配最不利反力桩组 → 取对应土层 ──
        pile_soil = match_pile_soil(
            reaction_dict, component_info, soil_param_index,
            excel_params.get('soil_param', []), excel_params.get('deck_level', 0)
        )

        secondary_dialog_values = foundation_value_dict.get(foundation_type, {})
        context_dict, type_dict = calc_drilled_shaft(
            excel_params, component_info, secondary_dialog_values, reaction_dict, pile_soil
        )
        formula_registry = context_dict.pop('formula_registry', {})
        foundation_satisfied = type_dict.get('satisfied', '')

    elif foundation_type == "扩大基础":
        if foundation_standard != "建筑地基基础设计规范":
            print(f"扩大基础规范应为建筑地基基础设计规范，当前: {foundation_standard}")
            return '', {}, ''

        reaction_dict = {
            'Fz_max': model_post_reaction[0] if len(model_post_reaction) > 0 else _default_react,
            'Mx_max': model_post_reaction[2] if len(model_post_reaction) > 2 else _default_react,
            'My_max': model_post_reaction[3] if len(model_post_reaction) > 3 else _default_react,
        }
        secondary_dialog_values = foundation_value_dict.get(foundation_type, {})

        # ── 从最不利反力节点匹配桩径 ──
        _Fz_node = str(reaction_dict['Fz_max'][1])
        _Mx_node = str(reaction_dict['Mx_max'][1])
        _My_node = str(reaction_dict['My_max'][1])
        _ctrl_node = _My_node if abs(float(reaction_dict['My_max'][7] or 0)) >= abs(float(reaction_dict['Mx_max'][6] or 0)) else _Mx_node
        _matched_d = 0
        for _key, _info in component_info.items():
            if not _key.endswith('钢管桩') or not isinstance(_info, dict):
                continue
            if _ctrl_node in _info.get('node_ids', []):
                _vsize = _info.get('vsize', [])
                if _vsize and _vsize[0]:
                    _matched_d = float(_vsize[0])
                break
        pile_diameter = _matched_d if _matched_d > 0 else pile_diameter
        print(f"[扩大基础] 控制节点={_ctrl_node}, 匹配桩径={pile_diameter}mm")

        # ── 从土层数据计算加权平均容重（按桩号匹配） ──
        soil_gs = None
        h1 = 0
        if soil_param_index:
            _spi_entry, _ = match_soil_entry(reaction_dict, component_info, soil_param_index)
            if _spi_entry:
                gl = float(_spi_entry.get('ground_level', 0))
                # soil_param_index 原始格式：name/thickness 为平行列表，需转为 layers
                _thicknesses = _spi_entry.get('thickness', [])
                _names = _spi_entry.get('name', [])
                if not _spi_entry.get('layers') and _names:
                    _soil_param = excel_params.get('soil_param', [])
                    _prop_map = {sp.get('name', ''): sp for sp in (_soil_param or [])}
                    _layers = []
                    for _j, _n in enumerate(_names):
                        _prop = _prop_map.get(_n, {})
                        _layers.append({
                            'name': _n,
                            'thickness': float(_thicknesses[_j]) if _j < len(_thicknesses) else 0,
                            'weight': float(_prop.get('weight', 18)) if _prop.get('weight') else 18.0,
                        })
                    _spi_entry = dict(_spi_entry, layers=_layers)
                _depths = [layer.get('thickness', 0) for layer in _spi_entry.get('layers', [])]
                pile_bottom_depth = sum(_depths)
                h1 = gl - pile_bottom_depth  # 基底标高（绝对标高）
                soil_gs = compute_average_gs(_spi_entry, h1)

        context_dict, formula_registry = calc_spread_foundation(
            secondary_dialog_values, reaction_dict, pile_diameter, soil_gs, h1
        )
        foundation_satisfied = context_dict.get('satisfied', '')

    elif foundation_type == "条形基础":
        if foundation_standard != "建筑地基基础设计规范":
            print(f"条形基础规范应为建筑地基基础设计规范，当前: {foundation_standard}")
            return '', {}, ''

        reaction_dict = {
            'Fz_max': model_post_reaction[0] if len(model_post_reaction) > 0 else _default_react,
            'Mx_max': model_post_reaction[2] if len(model_post_reaction) > 2 else _default_react,
            'My_max': model_post_reaction[3] if len(model_post_reaction) > 3 else _default_react,
        }
        secondary_dialog_values = foundation_value_dict.get(foundation_type, {})

        # ── 从最不利反力节点匹配桩径 ──
        _Fz_node = str(reaction_dict['Fz_max'][1])
        _Mx_node = str(reaction_dict['Mx_max'][1])
        _My_node = str(reaction_dict['My_max'][1])
        _ctrl_node = _My_node if abs(float(reaction_dict['My_max'][7] or 0)) >= abs(float(reaction_dict['Mx_max'][6] or 0)) else _Mx_node
        _matched_d = 0
        for _key, _info in component_info.items():
            if not _key.endswith('钢管桩') or not isinstance(_info, dict):
                continue
            if _ctrl_node in _info.get('node_ids', []):
                _vsize = _info.get('vsize', [])
                if _vsize and _vsize[0]:
                    _matched_d = float(_vsize[0])
                break
        pile_diameter = _matched_d if _matched_d > 0 else pile_diameter
        print(f"[条形基础] 控制节点={_ctrl_node}, 匹配桩径={pile_diameter}mm")

        # ── 从土层数据计算加权平均容重（按桩号匹配） ──
        soil_gs = None
        h1 = 0
        if soil_param_index:
            _spi_entry, _ = match_soil_entry(reaction_dict, component_info, soil_param_index)
            if _spi_entry:
                gl = float(_spi_entry.get('ground_level', 0))
                # soil_param_index 原始格式：name/thickness 为平行列表，需转为 layers
                _thicknesses = _spi_entry.get('thickness', [])
                _names = _spi_entry.get('name', [])
                if not _spi_entry.get('layers') and _names:
                    _soil_param = excel_params.get('soil_param', [])
                    _prop_map = {sp.get('name', ''): sp for sp in (_soil_param or [])}
                    _layers = []
                    for _j, _n in enumerate(_names):
                        _prop = _prop_map.get(_n, {})
                        _layers.append({
                            'name': _n,
                            'thickness': float(_thicknesses[_j]) if _j < len(_thicknesses) else 0,
                            'weight': float(_prop.get('weight', 18)) if _prop.get('weight') else 18.0,
                        })
                    _spi_entry = dict(_spi_entry, layers=_layers)
                _depths = [layer.get('thickness', 0) for layer in _spi_entry.get('layers', [])]
                pile_bottom_depth = sum(_depths)
                h1 = gl - pile_bottom_depth  # 基底标高（绝对标高）
                soil_gs = compute_average_gs(_spi_entry, h1)

        context_dict, formula_registry = calc_strip_foundation(
            secondary_dialog_values, reaction_dict, pile_diameter, soil_gs, h1
        )
        foundation_satisfied = context_dict.get('satisfied', '')

    else:
        print(f"未知基础类型: {foundation_type}")
        return '', {}, '满足'

    # ── 渲染子模板 ──
    sub_tpl_path = os.path.join(docx_open_path, 'Trestle_cal_template_foundation.docx')
    if not os.path.exists(sub_tpl_path):
        print(f"基础模板不存在: {sub_tpl_path}")
        return '', formula_registry, '满足'

    sub_tpl = DocxTemplate(sub_tpl_path)

    def _img(pic_path):
        if pic_path and os.path.exists(str(pic_path)):
            return InlineImage(sub_tpl, str(pic_path), width=Cm(16), height=Cm(10))
        return ''

    render_context = {
        'foundation_type': foundation_type,
        'drivenplie': context_dict if foundation_type == "打入桩" else {},
        'drilledshaft': context_dict if foundation_type == "钻孔灌注桩" else {},
        'type': type_dict if foundation_type == "钻孔灌注桩" else {},
        'spread': context_dict if foundation_type == "扩大基础" else {},
        'strip': context_dict if foundation_type == "条形基础" else {},
        'Midas_reaction_picture_biaozhun_Fxyz': _img(os.path.join(midas_path, '基础竖向反力图.jpg')),
        'Midas_reaction_picture_biaozhun_Mxyz': _img(os.path.join(midas_path, '基础弯矩图.jpg')),
    }
    sub_tpl.render(render_context)

    # 保存临时文件
    save_dir = output_dir or docx_open_path
    tmp_path = os.path.join(save_dir, '_foundation_subdoc_tmp.docx')
    sub_tpl.save(tmp_path)

    # 公式替换：将 LaTeX 占位符转为 Word OMML 公式
    if formula_registry:
        tmp_doc = Document(tmp_path)
        for placeholder, latex in formula_registry.items():
            if latex:
                replace_placeholder_with_formula(tmp_doc, placeholder, latex)
        tmp_doc.save(tmp_path)
        print(f"[foundation_subdoc_render] 公式替换完成: {len(formula_registry)} 个")

    subdoc = main_tpl.new_subdoc(tmp_path)
    # os.remove(tmp_path)  # 暂时保留临时文件

    print(f"[foundation_subdoc_render] {foundation_type} 子模板渲染完成")
    return subdoc, foundation_satisfied

