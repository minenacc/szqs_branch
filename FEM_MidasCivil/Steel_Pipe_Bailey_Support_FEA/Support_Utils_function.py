# 1. 标准库
import math

# 3. 本地模块
from General.AutoCAD import vtpnt


def duplicate(lst):
    new_list = list(set(lst))
    return new_list


# 获取捕捉点并在捕捉点位置生成一条x轴一条y轴直线
def get_point_and_make_xline(acad):
    try:
        doc = acad.ActiveDocument
        msp = doc.ModelSpace
        # 用户指定点
        pt = doc.Utility.Getpoint()
        # 竖直线
        direction_pt_y = [pt[0], pt[1] + 1, pt[2]]
        # 水平线
        direction_pt_x = [pt[0] + 1, pt[1], pt[2]]
        # 生成直线
        msp.AddXLine(vtpnt(pt), vtpnt(direction_pt_y))
        msp.AddXLine(vtpnt(pt), vtpnt(direction_pt_x))
        print("已成功生成定位坐标轴。")
        return pt
    except Exception as e:
        # print(e)
        try:
            excepinfo = e.excepinfo
            scode, source, desc, helpfile, helpcontext, scode2 = excepinfo
            if desc == None:
                print("错误！可能的原因：用户取消操作或其他。")
            else:
                print("错误！可能的原因：未成功与CAD连接，请重新尝试。")
        except:
            print("错误！可能的原因：未成功与CAD连接，请重新尝试。")


# 判断点对是否连续，连续返回none，不连续返回两个断点组成的点对
def if_continuity(point_pairs):
    discontinuous_points = []
    
    for i in range(len(point_pairs) - 1):
        current_pair = point_pairs[i]
        next_pair = point_pairs[i + 1]
        
        if current_pair[1] != next_pair[0]:
            discontinuous_points.append([current_pair[1], next_pair[0]])
    
    return discontinuous_points if discontinuous_points else None


# 两个节点之间存在单元的判断
def is_element_between_nodes(x1, x2, intervals):
    # 确保x1 <= x2
    if x1 > x2:
        x1, x2 = x2, x1
    for interval in intervals:
        start, end = interval
        # 检查两个x坐标形成的区间是否与当前区间重叠
        if not (x2 <= start or x1 >= end):
            return False
    if intervals == []:
        return True
    return True


# 判断一个点在点群中最近的点并返回该最近点在点群中的序号
def find_nearest_point_index(P, PL):
    x0 = P[0]
    y0 = P[1]
    Distance_min = float('inf')
    for i, point in enumerate(PL):
        if len(point) == 1:
            x1 = point[0]
            y1 = point[1]
        elif len(point) > 1:
            x1 = point[0][0]
            y1 = point[0][1]
        Distance = math.sqrt((x1-x0)**2 + (y1-y0)**2)
        if Distance < Distance_min:
            Distance_min = Distance
            if len(point) == 1:
                nearest_point = point
            elif len(point) > 1:
                nearest_point = point[0]
            nearest_index = i

    return (nearest_index + 1, nearest_point)


# 材料定义函数
def MATERIAL():
    materials = ["1, STEEL, Q235, 0, 0, , C, NO, 0.02, 1, GB03(S), , Q235, NO, 206",
		         "2, STEEL, 贝雷梁弦杆, 0, 0, , C, NO, 0.02, 1, JTJ(S), , 16Mn, NO, 210",
		         "3, STEEL, 贝雷梁竖杆, 0, 0, , C, NO, 0.02, 1, JTJ(S), , 16Mn, NO, 210",
		         "4, STEEL, 贝雷梁斜杆, 0, 0, , C, NO, 0.02, 1, JTJ(S), , 16Mn, NO, 210",
		        ]
    return materials


# 截面定义函数，需要的参数有：
# 小肋信息，钢管桩直径和壁厚，分配梁信息
def support_section_data(selected_rib_info):
    section_list=[
        "{} {}".format("1, ",selected_rib_info[0]),#小肋
        "2, DBUSER, 2C10, CC, 0, 0, 0, 0, 0, 0, YES, NO, 2C , 2, 100, 48, 5.3, 8.5, 80, 0, 0, 0, 0, 0",
        "3, DBUSER, I8, CC, 0, 0, 0, 0, 0, 0, YES, NO, H , 2, 80, 50, 4.5, 6.5, 0, 0, 0, 0, 0, 0"
    ]
    return section_list


# 荷载组合
def loadcomb():
    str1 = "NAME=标准组合, GEN, ACTIVE, 0, 0, , 0, 0"
    str2 = "ST, 自重, 1, ST, 混凝土荷载, 1, ST, 施工荷载, 0.7, ST, 模板荷载, 0.7, ST, 风荷载, 0.6, ST, 水流力, 0.7"
    str3 = "NAME=基本组合, GEN, ACTIVE, 0, 0, , 0, 0"
    str4 = "ST, 自重, 1.2, ST, 混凝土荷载, 1.4, ST, 施工荷载, 0.98, ST, 模板荷载, 0.98, ST, 风荷载, 0.84, ST, 水流力, 0.98"

    str_comb = [str1, str2, str3, str4]
    print("Success: 荷载组合创建成功。")
    return str_comb


# 根据水流力大小Fw、水面高程H1、冲刷线高程H2，计算水流力梯形荷载大小
def ShuiLiuLi_Line_Loads(Fw, H):
    if H != 0:
        # 水位线处水流力线荷载大小
        F1 = 2 * Fw / H
        # 冲刷线处水流力线荷载大小
        F2 = 0.0
    else:
        F1 = F2 = 0.0
    # 返回高程及对应线荷载大小，直线内插
    return round(F1,1)


# 对一连串的数据进行整理，碰到相同的则合并
def group_consecutive_elements(lst):
    if not lst:
        return []
    
    result = []
    current_group = [lst[0]]
    
    for i in range(1, len(lst)):
        if lst[i] == lst[i-1]:
            current_group.append(lst[i])
        else:
            result.append(current_group)
            current_group = [lst[i]]
    
    result.append(current_group)
    return result
