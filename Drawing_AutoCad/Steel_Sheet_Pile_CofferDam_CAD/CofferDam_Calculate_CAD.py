# 1. 标准库
import math
from itertools import accumulate

# 3. 本地模块
from General.Geometry import find_circle_line_intersection
from Drawing_AutoCad.Steel_Sheet_Pile_CofferDam_CAD.CofferDam_Resource_CAD import new_length_divisible, get_equidistant_points


def calc_skggz_center_spacing(diameter, lock_width, sheet_count, sheet_width):
    """计算锁扣钢管桩中心间距
    公式: 中心间距 = D + 锁扣宽度×2 + 钢板桩数量×钢板桩宽度

    Args:
        diameter: 钢管桩直径(mm)
        lock_width: 锁扣宽度(mm)
        sheet_count: 钢板桩数量
        sheet_width: 钢板桩宽度(mm)

    Returns:
        float: 中心间距(mm)
    """
    return diameter + lock_width * 2 + sheet_count * sheet_width


# ============================================================
# 几何计算函数（从 FEA 移植，支持双类型桩）
# ============================================================

def Cap_Pile_PM(Cap_X, Cap_Y, X_offset, Y_offset, Sheet_Pile_typevar,
                Pile_SEC_info_dict, recognize_Cap_result_dict,
                SKGGZ_D=0, SKGGZ_Gap=0):
    """
    计算承台角点和支护桩位置（平面图）
    边距定义：承台边线至桩内侧边线的距离
    """
    O_point = (0, 0, 0)
    print('定位原点:', O_point)
    Cap_X = Cap_X * 1000  # m → mm
    Cap_Y = Cap_Y * 1000
    X_offset = X_offset * 1000  # m → mm
    Y_offset = Y_offset * 1000

    # 承台角点（闭合多边形）
    Cap_corner_ptlst = [
        (O_point[0] + Cap_X / 2 * dx, O_point[1] + Cap_Y / 2 * dy, 0)
        for dx, dy in [(1, 1), (-1, 1), (-1, -1), (1, -1), (1, 1)]
    ]

    if Sheet_Pile_typevar == '钢板桩':
        Sheet_Pile_X = Pile_SEC_info_dict['SEC'][1]  # B（宽度，沿围堰周长方向）
        Sheet_Pile_Y = Pile_SEC_info_dict['SEC'][0]  # H（高度，垂直围堰周长方向）
        SEC_D = Sheet_Pile_X # 支护桩宽度
        SEC_H = Sheet_Pile_Y if Sheet_Pile_Y > 400 else Sheet_Pile_Y + 30 # 支护桩高度, 拉森四图块高度为370mm, 计算高度为340mm, 拉森六均为420mm
        l_n = (Cap_X + 2*X_offset)/SEC_D
        b_n = (Cap_Y + 2*Y_offset)/SEC_D
        l_n = math.ceil(l_n)
        b_n = math.ceil(b_n)
        l_n = l_n + (1 if l_n % 2 == 0 else 0) # 长边保持奇数
        b_n = b_n + (0 if b_n % 2 == 0 else 1) # 宽边保持偶数
        X_offset_mm = ((l_n*SEC_D) - Cap_X) / 2
        Y_offset_mm = ((b_n*SEC_D) - Cap_Y) / 2
    elif Sheet_Pile_typevar == '锁扣钢管桩':
        SEC_D = SKGGZ_D
        SEC_H = SKGGZ_D
        l_n = (Cap_X + 2*X_offset + SKGGZ_D)/SKGGZ_Gap
        b_n = (Cap_Y + 2*Y_offset + SKGGZ_D)/SKGGZ_Gap
        l_n = math.ceil(l_n)
        b_n = math.ceil(b_n)
        X_offset_mm = ((l_n*SKGGZ_Gap) - Cap_X - SKGGZ_D) / 2
        Y_offset_mm = ((b_n*SKGGZ_Gap) - Cap_Y - SKGGZ_D) / 2
    else:
        raise ValueError(f"未知截面类型: {Sheet_Pile_typevar}")
    
    print('支护桩高度:', SEC_H)
    print('支护桩宽度:', SEC_D)
    print("X方向预留边距:", X_offset_mm)
    print("Y方向预留边距:", Y_offset_mm)

    # 定位中心线角点坐标（使用原始输入边距，不用取整后的值）
    # 承台是轴对齐矩形，偏移方向应垂直于各边，不能用对角方向
    position_line_lst = []
    for i in range(len(Cap_corner_ptlst)):
        cap_corner_pt = Cap_corner_ptlst[i]
        # 用角点坐标的正负号确定垂直向外方向（不依赖角度计算）
        nx = 1 if cap_corner_pt[0] > 0 else (-1 if cap_corner_pt[0] < 0 else 0)
        ny = 1 if cap_corner_pt[1] > 0 else (-1 if cap_corner_pt[1] < 0 else 0)
        position_line_lst.append((
            cap_corner_pt[0] + nx * (X_offset_mm + SEC_H / 2),
            cap_corner_pt[1] + ny * (Y_offset_mm + SEC_H / 2),
            cap_corner_pt[2],
        ))

    # 支护桩坐标
    Pile_ptlst = []
    for i in range(len(position_line_lst) - 1):
        pt1 = position_line_lst[i]
        pt2 = position_line_lst[i + 1]
        if Sheet_Pile_typevar == '钢板桩':
            if i % 2 == 0: # 长边
                pt3 = find_circle_line_intersection(pt2, pt1, (SEC_D-SEC_H) / 2, 'EX')
                pt4 = find_circle_line_intersection(pt1, pt2, (SEC_D-SEC_H) / 2, 'EX')
            else: # 宽边
                pt3 = find_circle_line_intersection(pt2, pt1, SEC_H / 2, 'ON')
                pt4 = find_circle_line_intersection(pt1, pt2, SEC_H / 2, 'ON')
            Pile_ptlst.append(get_equidistant_points(pt3, pt4, SEC_D))
        elif Sheet_Pile_typevar == '锁扣钢管桩':
            Pile_ptlst.append(get_equidistant_points(pt1, pt2, SKGGZ_Gap))

    # 记录结果
    recognize_Cap_result_dict['O_Point'] = O_point
    recognize_Cap_result_dict['Cap_corner_ptlst'] = Cap_corner_ptlst
    recognize_Cap_result_dict['position_line_lst'] = position_line_lst
    recognize_Cap_result_dict['Pile_ptlst'] = Pile_ptlst
    recognize_Cap_result_dict['CapX'] = Cap_X
    recognize_Cap_result_dict['CapY'] = Cap_Y
    recognize_Cap_result_dict['X_offset_mm'] = X_offset_mm
    recognize_Cap_result_dict['Y_offset_mm'] = Y_offset_mm

    print('坐标原点:', O_point)
    print('承台角点:', Cap_corner_ptlst)
    print('支护桩中心线角点:', position_line_lst)
    print('支护桩图块坐标:')
    for x in Pile_ptlst:
        print(x)
    print('承台及支护桩图形计算完成')


def Waler_PM(CofferDam_Top_Level, Pile_SEC_info_dict, recognize_Cap_result_dict,
             recognize_Waler_result_dict, Walers_Strut_selected_dict,
             Waler_section_dict, Sheet_Pile_typevar, SKGGZ_D):
    """
    计算各层围檩中心线角点坐标
    """
    if Sheet_Pile_typevar == '钢板桩':
        Sheet_Pile_Y = Pile_SEC_info_dict['SEC'][0]  # Sheet_Pile_Y
        SEC_H = Sheet_Pile_Y if Sheet_Pile_Y > 400 else Sheet_Pile_Y + 30 # 拉森四图块高度为370mm, 计算高度为340mm, 拉森六均为420mm
    elif Sheet_Pile_typevar == '锁扣钢管桩':
        SEC_H = SKGGZ_D
    else:
        SEC_H = 0
    position_line_lst = recognize_Cap_result_dict['position_line_lst']
    # 围檩外围定位线（桩内侧，往承台方向偏移）
    Waler_line_lst = []
    for i in range(len(position_line_lst)):
        position_line_pt = position_line_lst[i]
        nx = -1 if position_line_pt[0] > 0 else (1 if position_line_pt[0] < 0 else 0)
        ny = -1 if position_line_pt[1] > 0 else (1 if position_line_pt[1] < 0 else 0)
        Waler_line_lst.append((
            position_line_pt[0] + nx * (SEC_H / 2),
            position_line_pt[1] + ny * (SEC_H / 2),
            position_line_pt[2],
        ))
    # 逐层计算围檩中心线
    recognize_Waler_result_dict['section'] = []
    recognize_Waler_result_dict['position_line_lst'] = []
    z = CofferDam_Top_Level  # 从桩顶开始
    for value in Walers_Strut_selected_dict.values():
        waler = value['围檩长边截面']
        spacing = float(value['间距'])
        z -= spacing  # 间距是从上往下累加的距离
        waler_H = float(Waler_section_dict[waler]['SEC'][0])  # 围檩截面高度
        Waler_position_line_lst = []
        for j in range(len(Waler_line_lst)):
            Waler_line_pt = Waler_line_lst[j]
            # 围檩在桩内侧、承台外侧，中心线往承台方向（内侧）偏移
            nx = -1 if Waler_line_pt[0] > 0 else (1 if Waler_line_pt[0] < 0 else 0)
            ny = -1 if Waler_line_pt[1] > 0 else (1 if Waler_line_pt[1] < 0 else 0)
            Waler_position_line_lst.append((
                Waler_line_pt[0] + nx * (waler_H / 2),
                Waler_line_pt[1] + ny * (waler_H / 2),
                z * 1000,
            ))
        recognize_Waler_result_dict['section'].append(waler)
        recognize_Waler_result_dict['position_line_lst'].append(Waler_position_line_lst)
    print('围檩图形计算完成')


def Strut_PM(Strut_blocks_dict, recognize_Waler_result_dict, recognize_Strut_result_dict):
    """
    计算各层内支撑（对撑+斜撑）线段坐标
    Strut_blocks_dict: {layer_idx: {'DC_X': str, 'DC_Y': str, 'XC_X': str, 'XC_Y': str}}
    """
    recognize_Strut_result_dict['position_line_duicheng_lst'] = []
    recognize_Strut_result_dict['position_line_xiecheng_lst'] = []

    if not recognize_Waler_result_dict.get('position_line_lst'):
        print('无围檩, 无内支撑计算')
        return

    for key, value in Strut_blocks_dict.items():
        Strut1_Xlst = value['DC_X']
        Strut2_Xlst = value['XC_X']
        Strut1_Ylst = value['DC_Y']
        Strut2_Ylst = value['XC_Y']

        # 将 '1.5+2.5'（m）转换成 [1500, 4000]（mm）
        acclst = []
        for stri in [Strut1_Xlst, Strut2_Xlst, Strut1_Ylst, Strut2_Ylst]:
            strlst = stri.split('+')
            if strlst[0] != '/':
                acclst.append(list(accumulate([abs(float(i)) * 1000 for i in strlst])))
            else:
                acclst.append(strlst)

        Strut1_X_acclst, Strut2_X_acclst, Strut1_Y_acclst, Strut2_Y_acclst = acclst

        # 获取该层围檩中心线角点
        if key >= len(recognize_Waler_result_dict['position_line_lst']):
            print(f'第{key+1}层围檩数据不存在，跳过')
            recognize_Strut_result_dict['position_line_duicheng_lst'].append([])
            recognize_Strut_result_dict['position_line_xiecheng_lst'].append([])
            continue

        position_line_lst = recognize_Waler_result_dict['position_line_lst'][key]
        Waler_X = abs(position_line_lst[0][0] - position_line_lst[1][0])
        Waler_Y = abs(position_line_lst[1][1] - position_line_lst[2][1])

        # 对撑
        Strut1_ptlst = []
        if Strut1_X_acclst[0] != '/':
            for x in Strut1_X_acclst:
                if x != 0:
                    Strut1_ptlst.append([(x, Waler_Y / 2, 0), (x, -Waler_Y / 2, 0)])
                    Strut1_ptlst.append([(-x, Waler_Y / 2, 0), (-x, -Waler_Y / 2, 0)])
                else:
                    Strut1_ptlst.append([(x, Waler_Y / 2, 0), (x, -Waler_Y / 2, 0)])
        if Strut1_Y_acclst[0] != '/':
            for y in Strut1_Y_acclst:
                if y != 0:
                    Strut1_ptlst.append([(Waler_X / 2, y, 0), (-Waler_X / 2, y, 0)])
                    Strut1_ptlst.append([(Waler_X / 2, -y, 0), (-Waler_X / 2, -y, 0)])
                else:
                    Strut1_ptlst.append([(Waler_X / 2, y, 0), (-Waler_X / 2, y, 0)])

        # 斜撑
        Strut2_ptlst = []
        if Strut2_X_acclst[0] != '/' and Strut2_Y_acclst[0] != '/':
            for x, y in zip(Strut2_X_acclst, Strut2_Y_acclst):
                Strut2_ptlst.extend([
                    [(x * i, Waler_Y / 2 * j, 0), (Waler_X / 2 * i, y * j, 0)]
                    for i, j in [(1, 1), (-1, 1), (-1, -1), (1, -1)]
                ])

        recognize_Strut_result_dict['position_line_duicheng_lst'].append(Strut1_ptlst)
        recognize_Strut_result_dict['position_line_xiecheng_lst'].append(Strut2_ptlst)

    print('内支撑计算完成')