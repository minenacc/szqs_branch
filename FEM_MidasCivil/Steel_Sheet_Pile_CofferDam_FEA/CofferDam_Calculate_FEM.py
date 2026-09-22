# =============================================================================
# 术语表（全文件统一约定）
#   SKGGZ    : 锁扣钢管桩
#   Waler    : 围檩（水平向的围护梁）
#   Strut    : 内支撑（含对撑 DuiCheng、斜撑 XieCheng）
#   Cap      : 承台
#   Bracket  : 牛腿
#   DC / XC  : 对撑 / 斜撑
#   PM       : 位置模型（Position Model），本文件用于计算各类构件的几何坐标
#   坐标系   : 原点位于承台中心，x/y 为水平方向，z 竖直向上
#   单位约定 : 输入 Cap_X/Cap_Y 为 m，内部换算为 mm；坐标、间距、截面高均为 mm；
#              标高 CofferDam_Top_Level 为 m，写入坐标时换算为 mm
# =============================================================================

# 1. 标准库
import re
from itertools import accumulate

# 3. 本地模块
from General.Geometry import find_circle_line_intersection


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


def new_length_divisible(length, param1, param2):
    """将线段长度调整为 param1 的整数倍，并控制份数的奇偶性。

    先取 length//param1 作为最大份数 divisor，再根据 param2 调整其奇偶：
    param2=2 时使 divisor 为偶数（份数为双数），param2=1 时使 divisor 为奇数
    （份数为单数），最终返回 param1*divisor。

    Args:
        length (float|int): 线段原始长度（mm）。
        param1 (float|int): 分段长度/模数（mm），如钢板桩宽度。
        param2 (int): 奇偶控制，2 表示要求份数为双数，1 表示要求为单数。

    Returns:
        float|int: 调整后的线段长度（mm），为 param1 的整数倍。

    Example:
        >>> new_length_divisible(10000, 400, 2)
        10400
    """
    divisor = length//param1 # 最大除数
    #最大除数不可以被param2整除, 最大除数+1, 最后结果是双数
    if param2 == 2:
        if divisor%2 != 0: 
            divisor += 1
    #最大除数可以被param2整除, 最大除数+1, 最后结果是单数
    if param2 == 1:
        if divisor%2 == 0: 
            divisor += 1
    new_length = param1*divisor
    return new_length


def new_length_divisible2(length, param1):
    """将线段长度向上取整为 param1 的整数倍（有余数则补一段）。

    Args:
        length (float|int): 线段原始长度（mm）。
        param1 (float|int): 分段长度/模数（mm），如锁扣钢管桩中心间距。

    Returns:
        float|int: 不小于 length 且为 param1 整数倍的长度（mm）。

    Example:
        >>> new_length_divisible2(10000, 371)
        10388
    """
    divisor = length//param1
    n = length%param1
    if n != 0:
        divisor += 1
    new_length = param1*divisor
    return new_length


def get_equidistant_points(pointA, pointB, spacing):
    """计算线段 AB 上所有间距为 spacing 的等分点坐标（含起点）。

    Args:
        pointA (tuple): 起点坐标 (x1, y1) 或 (x1, y1, z1)，仅取前两维。
        pointB (tuple): 终点坐标 (x2, y2) 或 (x2, y2, z2)，仅取前两维。
        spacing (float): 点间距（mm）。

    Returns:
        list[tuple]: 等分点坐标列表 [(x, y, 0), ...]，首点为 pointA，
            末点约等于 pointB（由 round(length/spacing) 决定）。

    Example:
        >>> get_equidistant_points((0, 0), (1000, 0), 500)
        [(0.0, 0.0, 0), (500.0, 0.0, 0), (1000.0, 0.0, 0)]
    """
    import math
    # 计算线段长度
    x1, y1 = pointA[0:2]
    x2, y2 = pointB[0:2]
    length = math.sqrt((x2 - x1)**2 + (y2 - y1)**2)
    # 计算等分点数量
    n = int(round(length / spacing))
    num_points = n + 1
    # 计算方向向量
    dx = (x2 - x1) / length
    dy = (y2 - y1) / length
    # 生成等分点坐标
    points = []
    for i in range(num_points):
        distance = i * spacing
        x = x1 + distance * dx
        y = y1 + distance * dy
        points.append((x, y, 0))
    return points


def Cap_Pile_PM(Cap_X, Cap_Y, X_offset, Y_offset, Sheet_Pile_typevar, Pile_SEC_info_dict, recognize_Cap_result_dict, SKGGZ_D, SKGGZ_Gap):
    """根据承台尺寸与支护类型，计算承台角点、定位中心线及支护桩点位。

    先按支护类型把承台外扩后的长度调整为截面宽度的整数倍（自动调整预留
    边距），再生成四条边的支护桩中心线，最后沿线按间距生成所有桩点。

    Args:
        Cap_X (float): 承台 X 方向长度（m）。
        Cap_Y (float): 承台 Y 方向长度（m）。
        X_offset (float): X 方向预留边距（m）。
        Y_offset (float): Y 方向预留边距（m）。
        Sheet_Pile_typevar (str): 支护类型，'钢板桩' 或 '锁扣钢管桩'。
        Pile_SEC_info_dict (dict): 支护桩截面信息，结构形式
            {'SEC': [高, 宽, ...]}（钢板桩用 SEC[1]/SEC[0]）。
        recognize_Cap_result_dict (dict): 承台识别结果字典，函数就地写入以下键：
            - 'O_Point': tuple，定位原点 (0, 0, 0)；
            - 'Cap_corner_ptlst': list[tuple]，承台 5 个角点（闭合）；
            - 'position_line_lst': list[tuple]，定位中心线角点；
            - 'Pile_ptlst': list[list[tuple]]，每条边的桩点列表。
        SKGGZ_D (float): 锁扣钢管桩直径（mm）。
        SKGGZ_Gap (float): 锁扣钢管桩中心间距（mm）。

    Returns:
        None: 结果通过就地修改 recognize_Cap_result_dict 返回。

    Example:
        >>> d = {}
        >>> Cap_Pile_PM(20, 15, 0.5, 0.5, '锁扣钢管桩', {'SEC': [0.8]}, d, 800, 1000)
        >>> sorted(d.keys())
        ['Cap_corner_ptlst', 'O_Point', 'Pile_ptlst', 'position_line_lst']
    """
    O_point = (0,0,0)
    print('定位原点:', O_point)
    Cap_X = Cap_X*1000 # m变mm
    Cap_Y = Cap_Y*1000 # m变mm
    # 绘制承台
    Cap_corner_ptlst = [(O_point[0] + Cap_X/2 * dx, O_point[1] + Cap_Y/2 * dy, 0) for dx, dy in [(1, 1), (-1, 1), (-1, -1), (1, -1), (1, 1)]] # 承台角点相对于原点的坐标
    if Sheet_Pile_typevar == '钢板桩':
        # 绘制钢板桩
        Sheet_Pile_X = Pile_SEC_info_dict['SEC'][1]
        Sheet_Pile_Y = Pile_SEC_info_dict['SEC'][0]
        Cap_A = new_length_divisible(Cap_X + X_offset*2*1000 - Sheet_Pile_X, Sheet_Pile_X, 1) + Sheet_Pile_X
        Cap_B = new_length_divisible(Cap_Y + Y_offset*2*1000 - Sheet_Pile_X, Sheet_Pile_X, 1) + Sheet_Pile_X
        X_offset_adjust = round((Cap_A-Cap_X)/2, 3)
        Y_offset_adjust = round((Cap_B-Cap_Y)/2, 3)
        SEC_H = Sheet_Pile_Y
    elif Sheet_Pile_typevar == '锁扣钢管桩':
        Cap_A = new_length_divisible2(Cap_X + X_offset*2*1000 + SKGGZ_D, SKGGZ_Gap)
        Cap_B = new_length_divisible2(Cap_Y + Y_offset*2*1000 + SKGGZ_D, SKGGZ_Gap)
        X_offset_adjust = round((Cap_A-Cap_X)/2 - SKGGZ_D/2, 3)
        Y_offset_adjust = round((Cap_B-Cap_Y)/2 - SKGGZ_D/2, 3)
        SEC_H = SKGGZ_D
    print("自动调整后X方向预留边距:", X_offset_adjust)
    print("自动调整后Y方向预留边距:", Y_offset_adjust)
    print("自动调整后承台预留边距后X长度:", Cap_A)
    print("自动调整后承台预留边距后Y长度:", Cap_B)
    # 定位中心线角点坐标
    position_line_lst = []
    nx_ny_lst = [(1,1), (-1,1), (-1,-1), (1,-1), (1,1)] # 缩放比例系数
    for i in range(len(Cap_corner_ptlst)):
        cap_corner_pt = Cap_corner_ptlst[i] # 承台第i个角点
        nx, ny = nx_ny_lst[i]
        position_line_lst.append((cap_corner_pt[0]+nx*(X_offset_adjust+SEC_H/2), cap_corner_pt[1]+ny*(Y_offset_adjust+SEC_H/2), cap_corner_pt[2]))
    print('定位中心线角点坐标')
    print(position_line_lst)
    # 钢板桩坐标
    Pile_ptlst = []
    for i in range(len(position_line_lst)-1):
        pt1 = position_line_lst[i] # 第i条边对应的钢板桩中心线的端点1
        pt2 = position_line_lst[i+1] # 第i条边对应的钢板桩中心线的端点2
        if Sheet_Pile_typevar == '钢板桩':
            pt3 = find_circle_line_intersection(pt2, pt1, Sheet_Pile_Y/2, 'ON') # 第i条边对应的钢板桩中心起点1
            pt4 = find_circle_line_intersection(pt1, pt2, Sheet_Pile_Y/2, 'ON') # 第i条边对应的钢板桩中心起点2
            Pile_ptlst.append(get_equidistant_points(pt3, pt4, Sheet_Pile_X))
        elif Sheet_Pile_typevar == '锁扣钢管桩':
            Pile_ptlst.append(get_equidistant_points(pt1, pt2, SKGGZ_Gap))
    # print(Pile_ptlst)
    # 将承台和钢板桩信息记录
    recognize_Cap_result_dict['O_Point'] = O_point # 定位原点
    recognize_Cap_result_dict['Cap_corner_ptlst'] = Cap_corner_ptlst # 承台角点
    recognize_Cap_result_dict['position_line_lst'] = position_line_lst # 定位中心线角点坐标
    recognize_Cap_result_dict['Pile_ptlst'] = Pile_ptlst # 所有钢板桩的点
    print('承台及支护桩图形计算完成')


def Waler_PM(CofferDam_Top_Level, Pile_SEC_info_dict, recognize_Cap_result_dict, recognize_Waler_result_dict, Walers_Strut_selected_dict, assist_add_dict, Waler_section_dict, Sheet_Pile_typevar, SKGGZ_D):
    """计算各层围檩的中心线角点坐标（含初始围檩与辅助加撑围檩）。

    以支护桩定位中心线为基准，按围檩截面高向外偏移得到围檩外围定位线，
    再逐层向内偏移半截面高，得到各层围檩中心线角点。

    Args:
        CofferDam_Top_Level (float): 围堰顶标高（m），作为第 1 层围檩定位起点。
        Pile_SEC_info_dict (dict): 支护桩截面信息 {'SEC': [...]}。
        recognize_Cap_result_dict (dict): Cap_Pile_PM 的输出，需含 'position_line_lst'。
        recognize_Waler_result_dict (dict): 围檩结果字典，就地写入：
            - 'index': list[int]，围檩层号；
            - 'section': list[str]，各层围檩截面名称；
            - 'position_line_lst': list[list[tuple]]，各层中心线角点（mm）。
        Walers_Strut_selected_dict (dict): 围檩/内支撑选型表，结构形式
            {'1行': {'围檩长边截面': str, '间距': str, ...}, ...}。
        assist_add_dict (dict): 辅助加撑围檩定义，结构形式
            {'层号': {'围檩长边截面': str, '标高': str, ...}, ...}。
        Waler_section_dict (dict): 围檩截面库，见 Resource.Waler_sections。
        Sheet_Pile_typevar (str): 支护类型，'钢板桩' 或 '锁扣钢管桩'。
        SKGGZ_D (float): 锁扣钢管桩直径（mm）。

    Returns:
        None: 结果通过就地修改 recognize_Waler_result_dict 返回。

    Example:
        >>> d = {}
        >>> Waler_PM(3.0, {'SEC': [0.8]}, cap_dict, d, sel_dict, {}, sec_dict, '锁扣钢管桩', 800)
        >>> d['index']
        [1, 2, 3]
    """
    if Sheet_Pile_typevar == '钢板桩':
        Sheet_Pile_Y = Pile_SEC_info_dict['SEC'][0]
        SEC_H = Sheet_Pile_Y
    elif Sheet_Pile_typevar == '锁扣钢管桩':
        SEC_H = SKGGZ_D
    # 钢板桩定位中心线角点坐标
    position_line_lst = recognize_Cap_result_dict['position_line_lst']
    # 围檩外围定位线
    Waler_line_lst = []
    nx_ny_lst = [(-1,-1), (1,-1), (1,1), (-1,1), (-1,-1)] # 缩放比例系数
    for i in range(len(position_line_lst)):
        position_line_pt = position_line_lst[i] # 钢板桩定位中心线第i个角点
        nx, ny = nx_ny_lst[i]
        Waler_line_lst.append((position_line_pt[0]+nx*(SEC_H/2), position_line_pt[1]+ny*(SEC_H/2), position_line_pt[2]))
    # 计算所有围檩的中心线角点坐标
    # print(Walers_Strut_selected_dict)
    # {'1行': {'围檩': 'HW400X400', '内支撑': '377X6钢管桩', '标高': '-0.5'}, '2行': {'围檩': 'HM588X300', '内支撑': '377X6钢管桩', '标高': '-1.5'}}
    recognize_Waler_result_dict['index'] = []
    recognize_Waler_result_dict['section'] = []
    recognize_Waler_result_dict['position_line_lst'] = []
    z = CofferDam_Top_Level # 初始定位标高
    # 支撑与垫层中定义的初始围檩
    for waler_i, value in enumerate(Walers_Strut_selected_dict.values()):
        waler = value['围檩长边截面'] # 第i行围檩截面
        z = round(z - float(value['间距']), 3) # 第i行围檩标高
        waler_H = float(Waler_section_dict[waler]['SEC'][0]) # 第i行围檩截面高
        Waler_position_line_lst = [] # 第i行围檩角点坐标
        for i in range(len(Waler_line_lst)):
            Waler_line_pt = Waler_line_lst[i] # 围檩外围定位线第i个角点
            nx, ny = nx_ny_lst[i]
            Waler_position_line_lst.append((round(Waler_line_pt[0]+nx*(waler_H/2), 3), round(Waler_line_pt[1]+ny*(waler_H/2), 3), round(z*1000 ,3)))
        recognize_Waler_result_dict['index'].append(waler_i+1)
        recognize_Waler_result_dict['section'].append(waler)
        recognize_Waler_result_dict['position_line_lst'].append(Waler_position_line_lst)
    # 辅助加撑中定义的围檩
    for k, v in assist_add_dict.items():
        waler = v['围檩长边截面'] # 第i行围檩截面
        waler_H = float(Waler_section_dict[waler]['SEC'][0]) # 第i行围檩截面高
        waler_z = float(v['标高']) # 第i行围檩截面标高
        Waler_position_line_lst = []
        for i in range(len(Waler_line_lst)):
            Waler_line_pt = Waler_line_lst[i] # 围檩外围定位线第i个角点
            nx, ny = nx_ny_lst[i]
            Waler_position_line_lst.append((round(Waler_line_pt[0]+nx*(waler_H/2), 3), round(Waler_line_pt[1]+ny*(waler_H/2), 3), round(waler_z*1000, 3)))
        recognize_Waler_result_dict['index'].append(int(k))
        recognize_Waler_result_dict['section'].append(waler)
        recognize_Waler_result_dict['position_line_lst'].append(Waler_position_line_lst)
    print('围檩外围定位线角点坐标:', Waler_line_lst)
    print('围檩中心线角点计算结果:')
    for k, v in recognize_Waler_result_dict.items():
        for x in v:
            print(k, x)
    print('围檩图形计算完成')


def Strut_PM(Strut_blocks_dict, assist_add_dict, assist_replace_dict, stage_dict, recognize_Waler_result_dict, recognize_Strut_result_dict, recognize_Strut_Replace_result_dict):
    """计算初始内支撑、辅助加撑、辅助换撑的对撑/斜撑线段端点坐标。

    流程分三段：
    1) 初始内支撑：按各层围檩矩形边长与"间距串"累加位置，生成 X/Y 向对撑
       与四象限斜撑的线段端点；
    2) 辅助加撑：逻辑同上，额外支持"是否与承台连接"（连接时对撑端点向
       承台投影并带高差）；
    3) 辅助换撑：仅计算施工阶段实际引用到的换撑，并对"与承台连接且相对
       上一工况为单元纯增"的对撑做继承优化，避免重复/丢失。

    Args:
        Strut_blocks_dict (dict): 初始内支撑布置，结构形式
            {'层号': {'DC_X': str, 'DC_Y': str, 'XC_X': str, 'XC_Y': str}, ...}，
            间距串形如 '1500+2500'，'/' 表示该向无支撑。
        assist_add_dict (dict): 辅助加撑定义，键为层号，值含
            {'对撑截面': str, '斜撑截面': str, '对撑长边布置(m)': str, ...,
             '是否与承台连接': '是'/'否', '承台长边投影(m)': str, ...}。
        assist_replace_dict (dict): 辅助换撑定义，结构形式
            {层号: {换撑序号: {同上辅助加撑字段}, ...}, ...}。
        stage_dict (dict): 施工阶段字典，用于筛选实际用到的换撑，
            结构形式 {'工况名': {'工况类型': str, '支撑层数': str, ...}, ...}。
        recognize_Waler_result_dict (dict): Waler_PM 的输出，需含
            'position_line_lst'（各层围檩中心线角点）。
        recognize_Strut_result_dict (dict): 初始/辅助内支撑结果，就地写入：
            - 'position_index_duicheng_lst': list[int]，对撑层号；
            - 'position_index_xiecheng_lst': list[int]，斜撑层号；
            - 'position_line_duicheng_lst': list[list[list[tuple]]]，对撑线段端点；
            - 'position_line_xiecheng_lst': list[list[list[tuple]]]，斜撑线段端点。
        recognize_Strut_Replace_result_dict (dict): 换撑结果，结构形式
            {层号: {换撑序号: {'position_line_duicheng_lst': [...],
                              'position_line_xiecheng_lst': [...]}}}。

    Returns:
        None: 结果通过就地修改两个 recognize_* 字典返回。

    Example:
        >>> Strut_PM(blocks, add, replace, stages, waler_d, strut_d, replace_d)
        >>> strut_d['position_index_duicheng_lst']
        [1, 2]
    """
    recognize_Strut_result_dict['position_index_duicheng_lst'] = [] # 第i层对撑对应的层号
    recognize_Strut_result_dict['position_index_xiecheng_lst'] = [] # 第i层斜撑对应的层号
    recognize_Strut_result_dict['position_line_duicheng_lst'] = [] # 第i层对撑对应的线段端点
    recognize_Strut_result_dict['position_line_xiecheng_lst'] = [] # 第i层斜撑对应的线段端点

    # recognize_Strut_Replace_result_dict['position_line_duicheng_lst'] = {} # 第i层对撑对应的线段端点
    # recognize_Strut_Replace_result_dict['position_line_xiecheng_lst'] = {} # 第i层斜撑对应的线段端点

    # 加撑计算/辅助加撑计算
    print('')
    print('初始内支撑计算==============================================')
    if recognize_Waler_result_dict['position_line_lst'] != []:
        for strut_i, (key, value) in enumerate(Strut_blocks_dict.items()):
            # 将'1500+2500'转换成1500,3500
            Strut1_Xlst = value['DC_X']
            Strut2_Xlst = value['XC_X']
            Strut1_Ylst = value['DC_Y']
            Strut2_Ylst = value['XC_Y']
            acclst = []
            for stri in [Strut1_Xlst, Strut2_Xlst, Strut1_Ylst, Strut2_Ylst]:
                strlst = stri.split('+')
                if strlst[0] != '/':
                    acclst.append(list(accumulate([abs(float(i)*1000) for i in strlst])))
                else:
                    acclst.append(strlst)
            print('内支撑间距累加计算:', acclst)
            Strut1_X_acclst, Strut2_X_acclst, Strut1_Y_acclst, Strut2_Y_acclst = acclst
            # 第i层围檩中心线角点坐标
            position_line_lst = recognize_Waler_result_dict['position_line_lst'][key]
            Waler_X = abs(position_line_lst[0][0] - position_line_lst[1][0]) # 中心线所构成的矩形X边长
            Waler_Y = abs(position_line_lst[1][1] - position_line_lst[2][1]) # 中心线所构成的矩形Y边长
            # 所有对撑的节点[[(), ()], [(), ()]]
            Strut1_ptlst = []
            # 所有对撑的节点[[(), ()], [(), ()]]
            Strut2_ptlst = []
            # X向对撑间距表
            if Strut1_X_acclst[0] != '/':
                for x in Strut1_X_acclst:
                    if x != 0:
                        Strut1_ptlst.append([(x, Waler_Y/2, 0), (x, -1*Waler_Y/2, 0)])
                        Strut1_ptlst.append([(-x, Waler_Y/2, 0), (-x, -1*Waler_Y/2, 0)])
                    else:
                        Strut1_ptlst.append([(x, Waler_Y/2, 0), (x, -1*Waler_Y/2, 0)])
            else:
                print('无X边对撑')
            # Y向对撑间距表
            if Strut1_Y_acclst[0] != '/':
                for y in Strut1_Y_acclst:
                    if y != 0:
                        Strut1_ptlst.append([(Waler_X/2, y, 0), (-1*Waler_X/2, y, 0)])
                        Strut1_ptlst.append([(Waler_X/2, -y, 0), (-1*Waler_X/2, -y, 0)])
                    else:
                        Strut1_ptlst.append([(Waler_X/2, y, 0), (-1*Waler_X/2, y, 0)])
            else:
                print('无Y边对撑')
            # 斜撑间距表
            if Strut2_X_acclst[0] != '/' and Strut2_Y_acclst[0] != '/':
                for x, y in zip(Strut2_X_acclst, Strut2_Y_acclst):
                    print(x, y)
                    Strut2_ptlst.extend([[(x*i, Waler_Y/2*j, 0), (Waler_X/2*i, y*j, 0)] for i, j in [(1,1), (-1,1), (-1,-1), (1,-1)]])
            else:
                print('无斜撑')
            recognize_Strut_result_dict['position_index_duicheng_lst'].append(strut_i+1)
            recognize_Strut_result_dict['position_index_xiecheng_lst'].append(strut_i+1)
            recognize_Strut_result_dict['position_line_duicheng_lst'].append(Strut1_ptlst)
            recognize_Strut_result_dict['position_line_xiecheng_lst'].append(Strut2_ptlst)
            print(f'第{strut_i+1}层初始内支撑计算完成')
            print('')
    else:
        print('无围檩, 无初始内支撑计算')
        print('')

    print('辅助加撑计算==============================================')
    if assist_add_dict != {}:
        for k, v in assist_add_dict.items():
            if v['对撑截面'] == '/' and v['斜撑截面'] == '/':
                print(f'第{k}层辅助加撑无内支撑')
            else:
                print(f'第{k}层辅助加撑存在内支撑')
                Strut1_Xlst = v['对撑长边布置(m)']
                Strut2_Xlst = v['斜撑长边布置(m)']
                Strut1_Ylst = v['对撑短边布置(m)']
                Strut2_Ylst = v['斜撑短边布置(m)']
                acclst = []
                for stri in [Strut1_Xlst, Strut2_Xlst, Strut1_Ylst, Strut2_Ylst]:
                    strlst = stri.split('+')
                    if strlst[0] != '/':
                        acclst.append(list(accumulate([abs(float(i)*1000) for i in strlst])))
                    else:
                        acclst.append(strlst)
                print('内支撑间距累加计算:', acclst)
                Strut1_X_acclst, Strut2_X_acclst, Strut1_Y_acclst, Strut2_Y_acclst = acclst
                # 第i层围檩中心线角点坐标
                position_line_lst = recognize_Waler_result_dict['position_line_lst'][int(k)-1]
                print(position_line_lst)
                Waler_X = abs(position_line_lst[0][0] - position_line_lst[1][0]) # 中心线所构成的矩形X边长
                Waler_Y = abs(position_line_lst[1][1] - position_line_lst[2][1]) # 中心线所构成的矩形Y边长
                # 所有对撑的节点[[(), ()], [(), ()]]
                Strut1_ptlst = []
                # 所有对撑的节点[[(), ()], [(), ()]]
                Strut2_ptlst = []
                # X向对撑间距表
                if Strut1_X_acclst[0] != '/' and v['是否与承台连接'] == '否':
                    for x in Strut1_X_acclst:
                        if x != 0:
                            Strut1_ptlst.append([(x, Waler_Y/2, 0), (x, -1*Waler_Y/2, 0)])
                            Strut1_ptlst.append([(-x, Waler_Y/2, 0), (-x, -1*Waler_Y/2, 0)])
                        else:
                            Strut1_ptlst.append([(x, Waler_Y/2, 0), (x, -1*Waler_Y/2, 0)])
                elif Strut1_X_acclst[0] != '/' and v['是否与承台连接'] == '是':
                    strut1_x_l = float(v['承台长边投影(m)'])
                    strut1_x_h = float(v['承台长边高差(m)'])
                    for x in Strut1_X_acclst:
                        if x != 0:
                            Strut1_ptlst.append([(x, Waler_Y/2, 0), (x, Waler_Y/2 - strut1_x_l * 1000, 0 - strut1_x_h * 1000)])
                            Strut1_ptlst.append([(x, -1*Waler_Y/2, 0), (x, -1*Waler_Y/2 + strut1_x_l * 1000, 0 - strut1_x_h * 1000)])
                            Strut1_ptlst.append([(-x, Waler_Y/2, 0), (-x, Waler_Y/2 - strut1_x_l * 1000, 0 - strut1_x_h * 1000)])
                            Strut1_ptlst.append([(-x, -1*Waler_Y/2, 0), (-x, -1*Waler_Y/2 + strut1_x_l * 1000, 0 - strut1_x_h * 1000)])
                        else:
                            Strut1_ptlst.append([(x, Waler_Y/2, 0), (x, Waler_Y/2 - strut1_x_l * 1000, 0 - strut1_x_h * 1000)])
                            Strut1_ptlst.append([(x, -1*Waler_Y/2, 0), (x, -1*Waler_Y/2 + strut1_x_l * 1000, 0 - strut1_x_h * 1000)])
                else:
                    print('无X边对撑')
                
                # Y向对撑间距表
                if Strut1_Y_acclst[0] != '/' and v['是否与承台连接'] == '否':
                    for y in Strut1_Y_acclst:
                        if y != 0:
                            Strut1_ptlst.append([(Waler_X/2, y, 0), (-1*Waler_X/2, y, 0)])
                            Strut1_ptlst.append([(Waler_X/2, -y, 0), (-1*Waler_X/2, -y, 0)])
                        else:
                            Strut1_ptlst.append([(Waler_X/2, y, 0), (-1*Waler_X/2, y, 0)])
                elif Strut1_Y_acclst[0] != '/' and v['是否与承台连接'] == '是':
                    strut1_y_l = float(v['承台短边投影(m)'])
                    strut1_y_h = float(v['承台短边高差(m)'])
                    for y in Strut1_Y_acclst:
                        if y != 0:
                            Strut1_ptlst.append([(Waler_X/2, y, 0), (Waler_X/2 - strut1_y_l * 1000, y, 0 - strut1_y_h * 1000)])
                            Strut1_ptlst.append([(-1*Waler_X/2, y, 0), (-1*Waler_X/2 + strut1_y_l * 1000, y, 0 - strut1_y_h * 1000)])
                            Strut1_ptlst.append([(Waler_X/2, -y, 0), (Waler_X/2 - strut1_y_l * 1000, -y, 0 - strut1_y_h * 1000)])
                            Strut1_ptlst.append([(-1*Waler_X/2, -y, 0), (-1*Waler_X/2 + strut1_y_l * 1000, -y, 0 - strut1_y_h * 1000)])
                        else:
                            Strut1_ptlst.append([(Waler_X/2, y, 0), (Waler_X/2 - strut1_y_l * 1000, y, 0 - strut1_y_h * 1000)])
                            Strut1_ptlst.append([(-1*Waler_X/2, y, 0), (-1*Waler_X/2 + strut1_y_l * 1000, y, 0 - strut1_y_h * 1000)])
                else:
                    print('无X边对撑')

                # 斜撑间距表
                if Strut2_X_acclst[0] != '/' and Strut2_Y_acclst[0] != '/':
                    for x, y in zip(Strut2_X_acclst, Strut2_Y_acclst):
                        print(x, y)
                        Strut2_ptlst.extend([[(x*i, Waler_Y/2*j, 0), (Waler_X/2*i, y*j, 0)] for i, j in [(1,1), (-1,1), (-1,-1), (1,-1)]])
                else:
                    print('无斜撑')
                recognize_Strut_result_dict['position_index_duicheng_lst'].append(int(k))
                recognize_Strut_result_dict['position_index_xiecheng_lst'].append(int(k))
                recognize_Strut_result_dict['position_line_duicheng_lst'].append(Strut1_ptlst)
                recognize_Strut_result_dict['position_line_xiecheng_lst'].append(Strut2_ptlst)
                print(f'第{k}层辅助内支撑计算完成')
                print('')
    else:
        print('无辅助加撑, 无辅助加撑计算')
        print('')

    print('加撑及辅助加撑信息:')
    for k, v in recognize_Strut_result_dict.items():
        print(k)
        for x in v:
            print(x)
    print('')

    # 辅助换撑计算
    print('辅助换撑计算==============================================')
    assit_replace_layer_in_stage_lst = [v['支撑层数'] for v in stage_dict.values() if v['工况类型'] == '辅助换撑'] # 确认施工阶段中实际用到的换撑信息，不计算全部换撑信息
    if assist_replace_dict != {}:
        for layer_replace_key, layer_replace_value in assist_replace_dict.items():
            recognize_Strut_Replace_result_dict[layer_replace_key] = {}
            for k, v in layer_replace_value.items():
                if f'{layer_replace_key}-{k}' not in assit_replace_layer_in_stage_lst:
                    continue
                recognize_Strut_Replace_result_dict[layer_replace_key][k] = {}
                recognize_Strut_Replace_result_dict[layer_replace_key][k]['position_line_duicheng_lst'] = []
                recognize_Strut_Replace_result_dict[layer_replace_key][k]['position_line_xiecheng_lst'] = []
                if v['对撑截面'] == '/' and v['斜撑截面'] == '/':
                    print(f'第{layer_replace_key}层第{k}次换撑内支撑截面为/, 即清空所有内支撑')
                    print()
                else:
                    print(f'第{layer_replace_key}层第{k}次换撑:')
                    acclst = []
                    Strut1_Xlst = v['对撑长边布置(m)']
                    Strut2_Xlst = v['斜撑长边布置(m)']
                    Strut1_Ylst = v['对撑短边布置(m)']
                    Strut2_Ylst = v['斜撑短边布置(m)']
                    

                    for stri in [Strut1_Xlst, Strut2_Xlst, Strut1_Ylst, Strut2_Ylst]:
                        strlst = stri.split('+')
                        if strlst[0] != '/':
                            acclst.append(list(accumulate([abs(float(i)*1000) for i in strlst])))
                        else:
                            acclst.append(strlst)
                    print('内支撑间距累加计算:', acclst)
                    Strut1_X_acclst, Strut2_X_acclst, Strut1_Y_acclst, Strut2_Y_acclst = acclst
                    # 第i层围檩中心线角点坐标
                    position_line_lst = recognize_Waler_result_dict['position_line_lst'][int(layer_replace_key)-1]
                    print(position_line_lst)
                    Waler_X = abs(position_line_lst[0][0] - position_line_lst[1][0]) # 中心线所构成的矩形X边长
                    Waler_Y = abs(position_line_lst[1][1] - position_line_lst[2][1]) # 中心线所构成的矩形Y边长
                    waler_z = position_line_lst[0][2] # 第i层围檩z坐标(mm), 与普通内支撑标高一致
                    # 所有对撑的节点[[(), ()], [(), ()]]
                    Strut1_ptlst = []
                    # 所有对撑的节点[[(), ()], [(), ()]]
                    Strut2_ptlst = []
                    # X向对撑间距表
                    if Strut1_X_acclst[0] != '/' and v['是否与承台连接'] == '否':
                        for x in Strut1_X_acclst:
                            if x != 0:
                                Strut1_ptlst.append([(x, Waler_Y/2, waler_z), (x, -1*Waler_Y/2, waler_z)])
                                Strut1_ptlst.append([(-x, Waler_Y/2, waler_z), (-x, -1*Waler_Y/2, waler_z)])
                            else:
                                Strut1_ptlst.append([(x, Waler_Y/2, waler_z), (x, -1*Waler_Y/2, waler_z)])
                    elif Strut1_X_acclst[0] != '/' and v['是否与承台连接'] == '是':
                        strut1_x_l = float(v['承台长边投影(m)'])
                        strut1_x_h = float(v['承台长边高差(m)'])
                        for x in Strut1_X_acclst:
                            if x != 0:
                                Strut1_ptlst.append([(x, Waler_Y/2, waler_z), (x, Waler_Y/2 - strut1_x_l * 1000, waler_z - strut1_x_h * 1000)])
                                Strut1_ptlst.append([(x, -1*Waler_Y/2, waler_z), (x, -1*Waler_Y/2 + strut1_x_l * 1000, waler_z - strut1_x_h * 1000)])
                                Strut1_ptlst.append([(-x, Waler_Y/2, waler_z), (-x, Waler_Y/2 - strut1_x_l * 1000, waler_z - strut1_x_h * 1000)])
                                Strut1_ptlst.append([(-x, -1*Waler_Y/2, waler_z), (-x, -1*Waler_Y/2 + strut1_x_l * 1000, waler_z - strut1_x_h * 1000)])
                            else:
                                Strut1_ptlst.append([(x, Waler_Y/2, waler_z), (x, Waler_Y/2 - strut1_x_l * 1000, waler_z - strut1_x_h * 1000)])
                                Strut1_ptlst.append([(x, -1*Waler_Y/2, waler_z), (x, -1*Waler_Y/2 + strut1_x_l * 1000, waler_z - strut1_x_h * 1000)])
                    else:
                        print('无X边对撑')

                    # Y向对撑间距表
                    if Strut1_Y_acclst[0] != '/' and v['是否与承台连接'] == '否':
                        for y in Strut1_Y_acclst:
                            if y != 0:
                                Strut1_ptlst.append([(Waler_X/2, y, waler_z), (-1*Waler_X/2, y, waler_z)])
                                Strut1_ptlst.append([(Waler_X/2, -y, waler_z), (-1*Waler_X/2, -y, waler_z)])
                            else:
                                Strut1_ptlst.append([(Waler_X/2, y, waler_z), (-1*Waler_X/2, y, waler_z)])
                    elif Strut1_Y_acclst[0] != '/' and v['是否与承台连接'] == '是':
                        strut1_y_l = float(v['承台短边投影(m)'])
                        strut1_y_h = float(v['承台短边高差(m)'])
                        for y in Strut1_Y_acclst:
                            if y != 0:
                                Strut1_ptlst.append([(Waler_X/2, y, waler_z), (Waler_X/2 - strut1_y_l * 1000, y, waler_z - strut1_y_h * 1000)])
                                Strut1_ptlst.append([(-1*Waler_X/2, y, waler_z), (-1*Waler_X/2 + strut1_y_l * 1000, y, waler_z - strut1_y_h * 1000)])
                                Strut1_ptlst.append([(Waler_X/2, -y, waler_z), (Waler_X/2 - strut1_y_l * 1000, -y, waler_z - strut1_y_h * 1000)])
                                Strut1_ptlst.append([(-1*Waler_X/2, -y, waler_z), (-1*Waler_X/2 + strut1_y_l * 1000, -y, waler_z - strut1_y_h * 1000)])
                            else:
                                Strut1_ptlst.append([(Waler_X/2, y, waler_z), (Waler_X/2 - strut1_y_l * 1000, y, waler_z - strut1_y_h * 1000)])
                                Strut1_ptlst.append([(-1*Waler_X/2, y, waler_z), (-1*Waler_X/2 + strut1_y_l * 1000, y, waler_z - strut1_y_h * 1000)])
                    else:
                        print('无X边对撑')

                    # 斜撑间距表
                    if Strut2_X_acclst[0] != '/' and Strut2_Y_acclst[0] != '/':
                        for x, y in zip(Strut2_X_acclst, Strut2_Y_acclst):
                            print(x, y)
                            Strut2_ptlst.extend([[(x*i, Waler_Y/2*j, waler_z), (Waler_X/2*i, y*j, waler_z)] for i, j in [(1,1), (-1,1), (-1,-1), (1,-1)]])
                    else:
                        print('无斜撑')
                    recognize_Strut_Replace_result_dict[layer_replace_key][k]['position_line_duicheng_lst'] = Strut1_ptlst
                    recognize_Strut_Replace_result_dict[layer_replace_key][k]['position_line_xiecheng_lst'] = Strut2_ptlst
                    print(f'第{layer_replace_key}层第{k}次换撑计算完成')
                    print('')
        # 对辅助换撑信息进行对撑判断
        def is_pure_addition(Elem_A, Elem_B):
            """判断集合 B 是否为集合 A 的"纯增"（A 是 B 的子集且 B 更大）。

            用于换撑优化：只有当前工况的对撑是上一工况的纯增（不含删减）时，
            才把上一工况的对撑继承到当前工况。

            Args:
                Elem_A (list): 上一工况的单元/坐标集合。
                Elem_B (list): 当前工况的单元/坐标集合。

            Returns:
                bool: 若 A 是 B 的子集且 len(B) > len(A) 返回 True，否则 False。
            """
            # 用 frozenset 消除单元方向影响：(i,j) 和 (j,i) 视为同一单元
            A_set = {e for e in Elem_A}
            B_set = {e for e in Elem_B}
            return A_set.issubset(B_set) and len(B_set) > len(A_set)
        
        for k, v in recognize_Strut_Replace_result_dict.items():
            print(k)
            print(v)
        # 确定施工阶段中辅助换撑的顺序
        assit_replace_layer_order = {}
        for stage_label in assit_replace_layer_in_stage_lst:
            _layer_i, _form_j = [int(x) for x in re.findall(r'\d+', stage_label)]
            if _layer_i not in assit_replace_layer_order.keys():
                assit_replace_layer_order[_layer_i] = []
            assit_replace_layer_order[_layer_i].append(_form_j)
        for k, v in assit_replace_layer_order.items():
            print(k)
            print(v)
        # 确定初始层的对撑布置
        origin_dc_dict = {}
        for _layer_i in assit_replace_layer_order.keys():
            origin_dc_dict[_layer_i] = recognize_Strut_result_dict['position_line_duicheng_lst'][int(_layer_i)-1]
        for k, v in origin_dc_dict.items():
            print(k)
            print(v)
        print('')
        # 判断当前工况的上一工况是否存在X对撑或Y对撑
        for _layer_i, _form_j_lst in assit_replace_layer_order.items():
            position_line_lst = recognize_Waler_result_dict['position_line_lst'][int(_layer_i)-1]
            _layer_z = position_line_lst[0][2]
            print(f'第{_layer_i}层标高: {_layer_z}')
            for index_i, _form_j in enumerate(_form_j_lst):
                print(f'第{_layer_i}层第{index_i+1}次换撑')
                print(f'对应换撑工况{_layer_i}-{_form_j}')
                
                connect_with_cap = True if assist_replace_dict[_layer_i][_form_j]['是否与承台连接'] == '是' else False
                if index_i == 0: # 当第一次换撑的时候, 上一工况是初始对撑层
                    last_dc = origin_dc_dict[_layer_i]
                else:
                    last_form_j = _form_j_lst[index_i-1] # 上一工况的换撑编号
                    last_dc = recognize_Strut_Replace_result_dict[_layer_i][last_form_j]['position_line_duicheng_lst']
                print('上一工况对撑节点坐标', last_dc)
                present_dc = recognize_Strut_Replace_result_dict[_layer_i][_form_j]['position_line_duicheng_lst']
                print('当前工况对撑节点坐标', present_dc)
                print('')
                # 将上一工况和当前工况的所有对撑分类为x对撑和y对撑
                last_dc_x = []
                last_dc_y = []
                present_dc_x = []
                present_dc_y = []
                for ptlst in last_dc:
                    pt1, pt2 = ptlst
                    x1, y1, x2, y2 = pt1[0], pt1[1], pt2[0], pt2[1]
                    if abs(x1 - x2) < 0.1:
                        last_dc_x.append(ptlst)
                    elif abs(y1 - y2) < 0.1:
                        last_dc_y.append(ptlst)
                for ptlst in present_dc:
                    pt1, pt2 = ptlst
                    x1, y1, x2, y2 = pt1[0], pt1[1], pt2[0], pt2[1]
                    if abs(x1 - x2) < 0.1:
                        present_dc_x.append(ptlst)
                    elif abs(y1 - y2) < 0.1:
                        present_dc_y.append(ptlst)
                # 判断单元是否纯增
                last_dc_x_lst = [ptlst[0][0] for ptlst in last_dc_x]
                last_dc_y_lst = [ptlst[0][1] for ptlst in last_dc_y]
                present_dc_x_lst = [ptlst[0][0] for ptlst in present_dc_x]
                present_dc_y_lst = [ptlst[0][1] for ptlst in present_dc_y]
                dc_x_if_add = is_pure_addition(last_dc_x_lst, present_dc_x_lst)
                dc_y_if_add = is_pure_addition(last_dc_y_lst, present_dc_y_lst)
                print('上一工况X对撑节点坐标', last_dc_x)
                print('上一工况Y对撑节点坐标', last_dc_y)
                print('上一工况X对撑节点X坐标', last_dc_x_lst)
                print('上一工况Y对撑节点Y坐标', last_dc_y_lst)
                print('')
                print('当前工况X对撑节点坐标', present_dc_x)
                print('当前工况Y对撑节点坐标', present_dc_y)
                print('当前工况X对撑节点X坐标', present_dc_x_lst)
                print('当前工况Y对撑节点Y坐标', present_dc_y_lst)
                print('')
                print('当前工况是否激活承台连接', connect_with_cap)
                print('当前工况相比于上一工况是否是单元纯增工况(X)', dc_x_if_add)
                print('当前工况相比于上一工况是否是单元纯增工况(Y)', dc_y_if_add)
                print('')
                # 当存在X对撑或Y对撑时, 当前工况计算的对撑信息要去除x相等或y相等的对撑, 并把上一个工况的对应x或y的对撑写入当前工况进行替换
                new_dc_x = []
                new_dc_y = []
                if last_dc_x != [] and connect_with_cap and dc_x_if_add: # 同时满足: 上一工况存在对撑, 当前工况选择激活承台连接, 当前工况节点是对上一工况的补充(单元纯增, 不包括单元既增又减)
                    remove_x = [ptlst[0][0] for ptlst in last_dc_x]
                    for present_ptlst in present_dc_x:
                        present_x = present_ptlst[0][0]
                        decrease_lst = [abs(present_x - x) for x in remove_x]
                        equal = any(x < 0.1 for x in decrease_lst)
                        if not equal:
                            new_dc_x.append(present_ptlst)
                    new_dc_x = new_dc_x + [[(ptlst[0][0], ptlst[0][1], _layer_z), (ptlst[1][0], ptlst[1][1], _layer_z)] for ptlst in last_dc_x]
                    new_dc_x = sorted(new_dc_x, key = lambda x : x[0][0])
                    print('当前工况优化后X对撑节点坐标', new_dc_x)
                if last_dc_y != [] and connect_with_cap and dc_y_if_add: # 同时满足: 上一工况存在对撑, 当前工况选择激活承台连接, 当前工况节点是对上一工况的补充(单元纯增, 不包括单元既增又减)
                    remove_y = [ptlst[0][1] for ptlst in last_dc_y]
                    for present_ptlst in present_dc_y:
                        present_y = present_ptlst[0][1]
                        decrease_lst = [abs(present_y - y) for y in remove_y]
                        equal = any(y < 0.1 for y in decrease_lst)
                        if not equal:
                            new_dc_y.append(present_ptlst)
                    new_dc_y = new_dc_y + [[(ptlst[0][0], ptlst[0][1], _layer_z), (ptlst[1][0], ptlst[1][1], _layer_z)] for ptlst in last_dc_y]
                    new_dc_y = sorted(new_dc_y, key = lambda x : x[0][0])
                    print('当前工况优化后Y对撑节点坐标', new_dc_y)
                new_dc_x = new_dc_x if new_dc_x != [] else present_dc_x
                new_dc_y = new_dc_y if new_dc_y != [] else present_dc_y
                print(new_dc_x)
                print(new_dc_y)
                new_dc = new_dc_x + new_dc_y
                print(new_dc)
                recognize_Strut_Replace_result_dict[_layer_i][_form_j]['position_line_duicheng_lst'] = new_dc
    else:
        print('无辅助换撑, 无辅助换撑计算')
        print('')

    print('辅助换撑信息:')
    for k0, v0 in recognize_Strut_Replace_result_dict.items():
        for k1, v1 in v0.items():
            print(f'第{k0}层第{k1}次换撑')
            for k2, v2 in v1.items():
                print(k2, v2)
            print('')

    print('内支撑信息计算完成==============================================')
    print('')

def Bracket_PM(recognize_Cap_result_dict, recognize_Bracket_result_dict, cofferdam_longside_bracket_num, cofferdam_shortside_bracket_num):
    """根据牛腿数量和支护桩坐标，找到牛腿所在的支护桩节点。

    牛腿均布在围堰各边并关于围堰中心线对称；同类边（长边/短边）使用相同
    基准坐标，确保上下/左右一致，最终吸附到距离理论位置最近的支护桩节点。

    Args:
        recognize_Cap_result_dict (dict): Cap_Pile_PM 的输出，需含
            'Pile_ptlst'（每条边的支护桩点列表）。
        recognize_Bracket_result_dict (dict): 牛腿结果字典，就地写入
            'bracket_points_by_side' 键。
        cofferdam_longside_bracket_num (int|str): 长边牛腿数量，"/" 或空表示未设置。
        cofferdam_shortside_bracket_num (int|str): 短边牛腿数量，"/" 或空表示未设置。

    Returns:
        dict: 就地修改后的 recognize_Bracket_result_dict，结构形式
            {'bracket_points_by_side': {side_idx: [point, ...]}}，
            其中 point 为 (x, y, z) 坐标；两侧均未设置时该键为空 dict。

    Example:
        >>> Bracket_PM(cap_d, bracket_d, 3, 2)
        {'bracket_points_by_side': {0: [(x, y, 0), ...], 1: [...]}}
    """
    # 处理"/"值（未设置牛腿数量的情况）
    def _parse_bracket_num(val):
        """将牛腿数量输入解析为整数，非法或未设置（None/""/"/"）返回 0。

        Args:
            val (int|float|str|None): 牛腿数量原始值。

        Returns:
            int: 解析后的牛腿数量，无法解析时为 0。
        """
        if val is None or str(val).strip() in ("", "/"):
            return 0
        try:
            return int(float(str(val)))
        except (ValueError, TypeError):
            return 0

    long_num = _parse_bracket_num(cofferdam_longside_bracket_num)
    short_num = _parse_bracket_num(cofferdam_shortside_bracket_num)

    # 如果两边都未设置牛腿，返回空结果
    if long_num <= 0 and short_num <= 0:
        return {'bracket_points': [], 'bracket_points_by_side': {}}

    cofferdam_pile_ptlst = recognize_Cap_result_dict['Pile_ptlst']

    def calc_bracket_coords(piles, n_brackets, coord_index):
        """计算某类边的牛腿基准坐标（只算一次，确保同类边一致）。

        按均布理论位置吸附到最近桩坐标，中心位置与正半轴位置分别处理，
        正半轴结果镜像到负半轴以保证关于中心线对称。

        Args:
            piles (list[tuple]): 基准边的桩坐标列表 [(x, y, z), ...]。
            n_brackets (int): 牛腿数量。
            coord_index (int): 取用的坐标索引（0=x, 1=y）。

        Returns:
            list[float]: 牛腿基准坐标值列表（升序）。
        """
        coords = sorted([p[coord_index] for p in piles])
        coord_min, coord_max = coords[0], coords[-1]
        L = coord_max - coord_min

        # 计算理论牛腿位置（均布）
        step = L / (n_brackets + 1)
        theoretical_positions = [coord_min + (i + 1) * step for i in range(n_brackets)]

        # 分离中心位置和正半轴位置
        center_pos = None
        positive_positions = []

        for pos in theoretical_positions:
            if abs(pos) < 1e-6:  # 中心位置（接近0）
                center_pos = pos
            elif pos > 0:
                positive_positions.append(pos)

        bracket_coords = []

        # 处理中心位置（如果存在，即n为奇数）
        if center_pos is not None:
            nearest = min(piles, key=lambda p: abs(p[coord_index] - center_pos))
            bracket_coords.append(nearest[coord_index])

        # 处理正半轴位置，并镜像到负半轴确保对称
        for pos in positive_positions:
            # 找正半轴最近的桩
            nearest = min(piles, key=lambda p: abs(p[coord_index] - pos))
            bracket_coords.append(nearest[coord_index])
            # 镜像
            bracket_coords.append(-nearest[coord_index])

        return sorted(bracket_coords)

    # 以顶边（side_idx=0）为基准计算长边牛腿x坐标
    long_side_bracket_x = calc_bracket_coords(cofferdam_pile_ptlst[0], long_num, 0) if long_num > 0 else []

    # 以左边（side_idx=1）为基准计算短边牛腿y坐标
    short_side_bracket_y = calc_bracket_coords(cofferdam_pile_ptlst[1], short_num, 1) if short_num > 0 else []

    # 应用到各边
    bracket_points = []
    bracket_points_by_side = {}

    for side_idx, piles in enumerate(cofferdam_pile_ptlst):
        side_bracket_points = []

        if side_idx == 0 or side_idx == 2:  # 长边（顶边/底边）
            for x_coord in long_side_bracket_x:
                nearest = min(piles, key=lambda p: abs(p[0] - x_coord))
                side_bracket_points.append(nearest)
        else:  # 短边（左边/右边）
            for y_coord in short_side_bracket_y:
                nearest = min(piles, key=lambda p: abs(p[1] - y_coord))
                side_bracket_points.append(nearest)

        bracket_points.extend(side_bracket_points)
        bracket_points_by_side[side_idx] = side_bracket_points

    recognize_Bracket_result_dict['bracket_points_by_side'] = bracket_points_by_side
    return recognize_Bracket_result_dict

