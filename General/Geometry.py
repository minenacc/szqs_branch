# 1. 标准库
import math

# 3. 本地模块
from General.DataUtils import Bailey_pt_close_to_ptinlst


# 求平面直线交点 (直线) 
def line_inter(line1, line2, param = 1):
    xdiff = (line1[0][0] - line1[1][0], line2[0][0] - line2[1][0])
    ydiff = (line1[0][1] - line1[1][1], line2[0][1] - line2[1][1])
    def det(a, b):
        return a[0] * b[1] - a[1] * b[0]
    div = det(xdiff, ydiff)
    if div == 0:
       return None
    d = (det(*line1), det(*line2))
    x = round(det(d, xdiff) / div, param)
    y = round(det(d, ydiff) / div, param)
    return x, y, 0


# 求平面直线交点 (线段) 
def line_intersection(line1, line2):
    x1 = line1[0][0]
    y1 = line1[0][1]
    x2 = line1[1][0]
    y2 = line1[1][1]
    x3 = line2[0][0]
    y3 = line2[0][1]
    x4 = line2[1][0]
    y4 = line2[1][1]
    # 计算分母
    denominator = (x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4)
    # 如果分母为0，则线段平行，没有交点
    if denominator == 0:
        return None
    # 计算分子
    t_numerator = (x1 - x3) * (y3 - y4) - (y1 - y3) * (x3 - x4)
    s_numerator = (x1 - x3) * (y1 - y2) - (y1 - y3) * (x1 - x2)
    # 计算 t 和 s
    t = t_numerator / denominator
    s = s_numerator / denominator
    # 检查 t 和 s 是否在 [0, 1] 之间
    if 0 <= t <= 1 and 0 <= s <= 1:
        # 计算交点坐标
        x = round(x1 + t * (x2 - x1), 1)
        y = round(y1 + t * (y2 - y1), 1)
        return (x, y, 0)
    else:
        return None
    

# 判断单元上荷载的非负值
def non_negative_elem_load(load_lst, height_lst):
    '''
    load_lst:[节点1线荷载, 节点2线荷载]
    height_lst:[节点1标高, 节点2标高]
    存在中间节点：
    返回[[节点1标高, 中间节点标高, 节点2标高], [0, 相对值, 1], [节点1线荷载, 0, 节点2线荷载]]
    不存在中间节点：
    返回[[节点1标高, 节点2标高], [0, 1], [节点1线荷载, 节点2线荷载]]
    全是负值:
    返回[[节点1标高, 节点2标高], [0, 1], [0, 0]]
    '''
    pt1_LineLoad, pt2_LineLoad = load_lst
    pt1_CoordZ, pt2_CoordZ = height_lst
    h = pt1_CoordZ-pt2_CoordZ
    line1 = [(0,0), (0, h)] # 两节点构成的竖直线段
    line2 = [(pt2_LineLoad, 0), (pt1_LineLoad, h)] # 两个线荷载节点构成的斜线
    # 求交点
    if pt1_LineLoad<=0 and pt2_LineLoad<=0: # 全是负数或是一0一负
        return [height_lst, [0, 1], [0, 0]]
    else:
        crosspt = line_inter(line1, line2, 3)
        if crosspt == None: # 无交点, pt1_LineLoad == pt2_LineLoad
            return [height_lst, [0, 1], load_lst]
        else:
            if is_point_on_segment(crosspt, line2, 0.001) and (pt1_LineLoad*pt2_LineLoad != 0): # 交点在线段上, 一正一负
                return [[pt1_CoordZ, round(pt2_CoordZ+crosspt[1],3), pt2_CoordZ], [0, round(1-crosspt[1]/h,3), 1], [round(pt1_LineLoad,3), 0, round(pt2_LineLoad,3)]]
            else: # 交点不在线段上, 全是正数或者是一0一正
                return [height_lst, [0, 1], [round(pt1_LineLoad,3), round(pt2_LineLoad,3)]]
            

def calculate_full_angle(A, B, C):
    """
    计算从向量AB到向量AC的旋转角度（0-360度）
    """
    # 创建向量AB和AC
    AB = (B[0] - A[0], B[1] - A[1])
    AC = (C[0] - A[0], C[1] - A[1])
    # 计算AB的方向角（与x轴的夹角）
    angle_AB = math.atan2(AB[1], AB[0])
    # 计算AC的方向角
    angle_AC = math.atan2(AC[1], AC[0])
    # 计算角度差（弧度）
    angle_rad = angle_AC - angle_AB
    # 将角度差规范化到0-2π范围
    if angle_rad < 0:
        angle_rad += 2 * math.pi
    # 转换为角度
    angle_deg = math.degrees(angle_rad)
    return round(angle_deg, 1)


# 距离起点长度L的直线上一点的坐标
def point_coordinate_on_line(start, end, L):
    """
    计算直线上距离起点为 L 长度的点坐标
    :param start: 起点坐标 (x1, y1, [z1])
    :param end: 终点坐标 (x2, y2, [z2])
    :param L: 距离起点的长度
    :return: 所求点的坐标
    """
    # 计算直线方向向量
    dx = end[0] - start[0]
    dy = end[1] - start[1]
    # 处理三维点（如果提供 z 坐标）
    dz = 0
    if len(start) > 2 and len(end) > 2:
        dz = end[2] - start[2]
    # 计算直线长度
    line_length = math.sqrt(dx**2 + dy**2 + dz**2)
    # 处理特殊情况
    if line_length == 0:
        return start  # 起点和终点重合
    # 计算单位方向向量
    ux = dx / line_length
    uy = dy / line_length
    uz = dz / line_length
    # 计算所求点坐标
    if L <= 0:
        return start  # 距离为0或负值，返回起点
    elif L >= line_length:
        return end  # 距离超过直线长度，返回终点
    else:
        # 计算所求点坐标
        x = start[0] + ux * L
        y = start[1] + uy * L
        # 处理三维点
        if len(start) > 2 and len(end) > 2:
            z = start[2] + uz * L
            return (x, y, z)
        else:
            return (x, y)
        

def find_circle_line_intersection(A, B, r, param = 'BOTH'):
    """
    计算直线AB上, 以B为圆心半径为r的圆与AB的两个交点
    param = 'BOTH':
    return: 两个角点
    param = 'ON':
    return: 线段上的一个点坐标
    param = 'EX':
    return: 延长线上的一个点坐标
    """
    # 计算向量AB
    dx = B[0] - A[0]
    dy = B[1] - A[1]
    # 计算向量AB的长度
    d = math.sqrt(dx**2 + dy**2)
    # 如果A和B重合，则没有交点或无限多个交点
    if d == 0:
        return None
    # 计算单位向量
    ux = dx / d
    uy = dy / d
    # 计算两个交点
    intersection1 = (B[0] + r * ux, B[1] + r * uy)
    intersection2 = (B[0] - r * ux, B[1] - r * uy)
    # 使用向量点积判断哪个点在线段AB的延长线上
    # 对于点P，如果向量BP与向量BA的点积为负，则P在BA的延长线上（不在AB上）
    def is_extended(point):
        bx = point[0] - B[0]
        by = point[1] - B[1]
        # 计算向量BP与向量BA的点积
        dot_product = bx * (A[0] - B[0]) + by * (A[1] - B[1])
        return dot_product > 0  # 如果点积为正，则点在线段AB的延长线上
    # 返回在线段AB延长线上的点
    if param == 'BOTH':
        return [intersection1, intersection2]
    elif param == 'ON':
        if is_extended(intersection1):
            return intersection1
        else:
            return intersection2
    elif param == 'EX':
        if is_extended(intersection1):
            return intersection2
        else:
            return intersection1


def point_to_line_projection(P0, P1, P2):
    """
    计算点P0(x, y, z)在线段AB（端点P1(x1,y1,z1)和P2(x2,y2,z2)）上的投影点坐标
    
    参数:
    x, y, z: 目标点的坐标
    x1, y1, z1: 线段端点A的坐标
    x2, y2, z2: 线段端点B的坐标
    
    返回:
    (px, py, pz): 投影点坐标
    """
    # 计算向量AB
    x,y,z = P0
    x1,y1,z1 = P1
    x2,y2,z2 = P2
    abx = x2 - x1
    aby = y2 - y1
    abz = z2 - z1
    # 计算向量AP
    apx = x - x1
    apy = y - y1
    apz = z - z1
    # 计算AB长度的平方
    ab_len_sq = abx**2 + aby**2 + abz**2
    # 处理线段长度为0的情况（A和B重合）
    if ab_len_sq < 1e-10:
        return (x1, y1, z1)
    # 计算AP在AB上的投影长度比例 (点积)
    dot_product = apx * abx + apy * aby + apz * abz
    t = dot_product / ab_len_sq
    # 限制t在[0,1]范围内，确保投影点在线段上
    t = max(0, min(1, t))
    # 计算投影点坐标
    px = x1 + t * abx
    py = y1 + t * aby
    pz = z1 + t * abz
    return (px, py, pz)


# 判断点是否在线段上，Line为二维点表
def is_point_on_segment(Point, Line, acc=0.5):
    px = Point[0]
    py = Point[1]
    ax = Line[0][0]
    ay = Line[0][1] 
    bx = Line[1][0]
    by = Line[1][1]
    # 判断点是否在线段端点上
    if (abs(px-ax)<acc and abs(py-ay)<acc) or (abs(px-bx)<acc and abs(py-by)<acc): # 模糊处理, 当点的距离过近时认为两点为同一点, acc为坐标差值
        return True
    # 计算向量 AP 和 AB
    else:
        AB = (bx - ax, by - ay)
        AP = (px - ax, py - ay)
        BP = (px - bx, py - by)
        # 计算点积
        dot_product = AP[0] * AB[0] + AP[1] * AB[1]
        # 计算向量 AP 和 AB 的模
        AP_magnitude = math.sqrt(AP[0]**2 + AP[1]**2)
        AB_magnitude = math.sqrt(AB[0]**2 + AB[1]**2)
        BP_magnitude = math.sqrt(BP[0]**2 + BP[1]**2)
        # print(AP_magnitude, AB_magnitude)
        # 计算余弦值
        cos_theta = dot_product / (AP_magnitude * AB_magnitude)
        # 检查余弦值是否接近 1
        is_on_segment = abs(cos_theta - 1) < 0.00001
        # print(abs(cos_theta - 1))
        # 判断点是否超出了线段
        if AP_magnitude <= AB_magnitude and BP_magnitude <= AB_magnitude: # 当夹角为0, 线段AP和BP的长度均小于等于线段AB的长度时, 点在线上
            return is_on_segment
        else:
            return False


# Centroid([[-1,0],[2,0],[2,3],[-1,3]])-->[0.5, 1.5]
# 多边形的重心，多边形的顶点坐标为Pts
def vxs(v, s):
    return list(map(lambda n: n * s, v))


def Centroid(Pts):
    return vxs(list(map(lambda *args: sum(args), *Pts)), 1.0 / len(Pts))


# 求两点中点函数
def midpoint(p1, p2):
    return [(x1 + x2) / 2 for x1, x2 in zip(p1, p2)]


# 找到首尾相连且在角度相同的的线段
def Bailey_connect_line_with_point_angle(endpoint, angle, lineptlst):
    # 端点初始值
    # 创建空表
    res_lst = []
    # 寻找连接点
    # 当一片贝雷梁的终点出现在剩余贝雷梁的起点表中时，
    while Bailey_pt_close_to_ptinlst(endpoint, lineptlst, angle):
        # 找到下一片贝雷梁对应的信息
        lst = Bailey_pt_close_to_ptinlst(endpoint, lineptlst, angle)
        # 替换终点
        endpoint = lst[4]
        res_lst.append(lst)
    return res_lst


# 判断点是否在多边形内部
# point为点的坐标，形如[1,1]
# polygon为多边形的顶点坐标，形如[[1,1],[2,2],[3,3],[1,1]]
def is_point_inside_polygon(point, polygon):
    num_intersections = 0
    x, y = point[0], point[1]
    for i in range(len(polygon)):
        p1 = polygon[i]
        p2 = polygon[(i + 1) % len(polygon)]
        if (p1[1] > y) != (p2[1] > y):
            if x < (p2[0] - p1[0]) * (y - p1[1]) / (p2[1] - p1[1]) + p1[0]:
                num_intersections += 1
    return num_intersections % 2 == 1


def isPolylineWithinPolyline(polyline1, polyline2):
    # 判断polyline1是否在polyline2内部
    for point in polyline1:
        if not is_point_inside_polygon(point, polyline2):
            return False
    return True


# 判断一群点是否有任意一个点落在一个范围内
def if_ptlst_have_pt_in_range(pointlst, point1, point2):
    x1 = point1[0]
    x2 = point2[0]
    xmin = min(x1, x2)
    xmax = max(x1, x2)
    xlst = [coor[0] for coor in pointlst]
    reslst = []
    for x in xlst:
        if x-xmin >= 0 and x-xmax <= 0:
            reslst.append("T")
    return reslst 


def is_point_inside_polyline(point, polyline):
    # 检查点是否在多段线内部
    return polyline.IsPointInside(point.X, point.Y, 0)


def check_polyline_intersection(polyline1, polyline2):
    # 检查两个几何对象是否有交点
    intersection_points = polyline1.IntersectWith(polyline2, 0)
    # 如果交点列表不为空，则表示有交点
    if len(intersection_points) > 0:
        # print(intersection_points)
        return intersection_points
    else:
        return False
    

# 计算两三维坐标点所构成的线段上的一点，该点的z坐标与最高点的z坐标相差H
def point_in_line_Z_scale(p1, p2, z):
    if p1[2] > p2[2]:
        top_p = p1
        but_p = p2
    elif p1[2] < p2[2]:
        top_p = p2
        but_p = p1
    elif p1[2] == p2[2]:
        return None
    x1, y1, z1 = top_p
    x2, y2, z2 = but_p
    # 比例
    scale = z / abs(z1 - z2)
    # 平面坐标
    if x2 - x1 >= 0:
        x = x1 + scale * abs(x1 - x2)
    else:
        x = x1 - scale * abs(x1 - x2)
    if y2 - y1 >= 0:
        y = y1 + scale * abs(y1 - y2)
    else:
        y = y1 - scale * abs(y1 - y2)
    z = z1 - z
    return (x, y, z)


# 两点之间的线段长
def distance(point1, point2):
    x1, y1, z1 = point1
    x2, y2, z2 = point2
    return math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2 + (z2 - z1) ** 2)


#获取圆弧的中点坐标
def midptcoordonarc(startptcoord , endptcoord , bulge , accuracy):
    x1, y1 = startptcoord[0] , startptcoord[1]
    x2, y2 = endptcoord[0] , endptcoord[1]
    dist = math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)#圆弧起点和终点的距离
    radius = (1 + bulge ** 2) * dist / ( math.fabs(bulge) * 4)#圆弧的半径
    c = (1 / bulge - bulge) / 2 #圆弧的中心角
    center = [(x1 + x2 - (y2 - y1) * c) / 2 ,(y1 + y2 + (x2 - x1) * c) / 2] #圆弧的圆心坐标
    midpoint = [(x1 + x2) / 2, (y1 + y2) / 2]#圆弧的起点和终点的中点坐标
    vector = [midpoint[0] - center[0], midpoint[1] - center[1]]#圆心到midpoint的向量
    dist2 = math.sqrt(vector[0] ** 2 + vector[1] ** 2) #圆心到中点的距离
    vector2 = [vector[0] * radius / dist2, vector[1] * radius / dist2]#圆心到中点的向量
    midpoint2 = [round(center[0] + vector2[0],accuracy), round(center[1] + vector2[1],accuracy) ]#圆弧的中点坐标
    return midpoint2


# 两线段求交，line1 = (pt1, pt2), line2 = (pt3, pt4)
def intersection_point_of_two_line_segment(line1, line2):
    x1, y1, z1 = line1[0]
    x2, y2, z2 = line1[1]
    x3, y3, z3 = line2[0]
    x4, y4, z4 = line2[1]
    # 计算斜率和截距
    if x2 != x1:
        m1 = (y2 - y1) / (x2 - x1)
        b1 = y1 - m1 * x1
    else:
        m1 = None
        b1 = x1
    if x4 != x3:
        m2 = (y4 - y3) / (x4 - x3)
        b2 = y3 - m2 * x3
    else:
        m2 = None
        b2 = x3
    # 检查斜率是否相同
    if m1 == m2:
        return None  # 平行或重合
    # 计算交点
    if m1 is None:
        x = b1
        y = m2 * x + b2
    elif m2 is None:
        x = b2
        y = m1 * x + b1
    else:
        x = (b2 - b1) / (m1 - m2)
        y = m1 * x + b1
    # 检查交点是否在线段上
    if (min(x1, x2) <= x <= max(x1, x2) and
        min(x3, x4) <= x <= max(x3, x4) and
        min(y1, y2) <= y <= max(y1, y2) and
        min(y3, y4) <= y <= max(y3, y4)):
        return (round(x, 1), round(y, 1))
    else:
        return None
    

# 垂直距离函数
def perdistToline(pt, p1, p2):
    x1, y1, z1 = pt
    x2, y2, z2 = p1
    x3, y3, z3 = p2
    # Calculate the vector from p1 to p2
    v = (x3 - x2, y3 - y2, z3 - z2)
    # Calculate the vector from p1 to pt
    w = (x1 - x2, y1 - y2, z1 - z2)
    # Calculate the dot product of v and w
    dot_product = v[0] * w[0] + v[1] * w[1] + v[2] * w[2]
    # Calculate the squared length of v
    v_length_squared = v[0] ** 2 + v[1] ** 2 + v[2] ** 2
    # Calculate the parameter t
    t = dot_product / v_length_squared
    # Calculate the closest point on the line to pt
    closest_point = (x2 + t * v[0], y2 + t * v[1], z2 + t * v[2])
    # Calculate the distance between pt and the closest point
    distance = ((x1 - closest_point[0]) ** 2 + (y1 - closest_point[1]) ** 2 + (z1 - closest_point[2]) ** 2) ** 0.5
    return distance


# 数轴上各数之间的差值 (间距) 
def get_distance_from_xlst(xlst, acc=1):
    appendlst = []
    for i in range(len(xlst)-1):
        x1 = xlst[i]
        x2 = xlst[i+1]
        distance = round(abs(x1-x2),acc)
        appendlst.append(distance)
    return appendlst


def Line_Offset(Line, Offset, param = 'BOTH'):
    """
    直线段偏移一定距离后产生的两条新直线, param控制正负偏移, 以向右为正
    """
    x1, y1, z1 = Line[0]
    x2, y2, z2 = Line[1]
    if x1 == x2: # 竖直线
        x_increment = Offset # x增量
        y_increment = 0      # y增量
        nx, ny = 1, 1
    elif y1 == y2: # 水平线
        x_increment = 0      # x增量
        y_increment = Offset # y增量
        nx, ny = 1, 1
    else:
        delta_x = abs(x1-x2) # x方向长度
        delta_y = abs(y1-y2) # y方向长度
        Distance = math.sqrt((delta_x)**2+(delta_y)**2) # 线段长度
        x_increment = Offset / Distance * delta_y # x增量
        y_increment = Offset / Distance * delta_x # y增量
        # 两点中y坐标比较小的那一点
        pt_sorted = sorted([(x1,y1,z1), (x2,y2,z2)], key = lambda x : x[1])
        ptA = pt_sorted[0]
        ptB = (ptA[0]+1, ptA[1], ptA[2])
        ptC = pt_sorted[1]
        angle = calculate_full_angle(ptA, ptB, ptC)
        if angle < 90:
            nx, ny = 1, -1
        else:
            nx, ny = 1, 1
    line1 = [(x1+nx*x_increment, y1+ny*y_increment, z1), (x2+nx*x_increment, y2+ny*y_increment, z2)]
    line2 = [(x1-nx*x_increment, y1-ny*y_increment, z1), (x2-nx*x_increment, y2-ny*y_increment, z2)]
    if param == 'BOTH':
        return [line1, line2]
    elif param == 'Positive':
        return [line1]
    elif param == 'Negative':
        return [line2]
    

# 主动土压力的0点和非负值
def insert_zero_and_clip_negative(depths, pressures):
    """
    在压力值异号的相邻点之间插入零点（压力为0），并将所有负压力值改为0。
    
    参数:
        depths: list[float] 深度列表（从浅到深，可能包含重复值）
        pressures: list[float] 对应深度的主动土压力值
    
    返回:
        new_depths, new_pressures: 插入零点且负值归零后的新列表
    """
    new_depths = []
    new_pressures = []
    n = len(depths)
    if n == 0:
        return new_depths, new_pressures
    # 先加入第一个点
    new_depths.append(depths[0])
    new_pressures.append(pressures[0])
    for i in range(n - 1):
        d1, d2 = depths[i], depths[i + 1]
        p1, p2 = pressures[i], pressures[i + 1]
        # 如果区间长度不为零且两端压力异号（严格一正一负），则插入零点
        if d1 != d2 and p1 * p2 < 0:
            # 线性插值计算零点深度
            d0 = round(d1 + (0 - p1) * (d2 - d1) / (p2 - p1), 3)
            # 确保插值点位于区间内（防止浮点误差）
            if (d1 <= d0 <= d2) or (d2 <= d0 <= d1):
                new_depths.append(d0)
                new_pressures.append(0.0)
        # 加入下一个点
        new_depths.append(d2)
        new_pressures.append(p2)
    # 将所有负压力值改为0
    new_pressures = [max(0, p) for p in new_pressures]
    return new_depths, new_pressures


# 通用荷载几何积分函数
def calculate_pressure_resultant(altitudes, layer_pressures, reference_alt):
    """
    通用压力几何积分函数 (适配嵌套列表输入)
    :param altitudes: 标高列表 [H0, H1, H2, ...]，长度为 N
    :param layer_pressures: 层顶层底压力列表 [[P0_top, P0_bot], [P1_top, P1_bot], ...]，长度为 N-1
    :param reference_alt: 参考基准标高 (计算力臂用)
    :return: (total_force, final_arm)
    """
    total_force = 0.0
    total_moment = 0.0
    # 遍历每一层 (逐层处理)
    for i in range(len(layer_pressures)):
        h_t = altitudes[i]      # 当前层顶标高
        h_b = altitudes[i+1]    # 当前层底标高
        p_t = layer_pressures[i][0]  # 该层顶压力
        p_b = layer_pressures[i][1]  # 该层底压力
        # 零点切割与有效段识别
        active_sub_segments = []
        if p_t >= 0 and p_b >= 0:
            # 情况A：全正压段 (梯形)
            active_sub_segments.append((h_t, h_b, p_t, p_b))
        elif p_t < 0 and p_b > 0:
            # 情况B：跨零点 (由负变正)，截取下部三角形
            # 相似三角形求零点标高: h_z = h_t - |p_t|/(|p_t|+|p_b|) * 层厚
            h_z = h_t - (abs(p_t) / (abs(p_t) + abs(p_b))) * (h_t - h_b)
            active_sub_segments.append((h_z, h_b, 0.0, p_b))
        elif p_t > 0 and p_b < 0:
            # 情况C：跨零点 (由正变负)，截取上部三角形
            h_z = h_t - (abs(p_t) / (abs(p_t) + abs(p_b))) * (h_t - h_b)
            active_sub_segments.append((h_t, h_z, p_t, 0.0))
        # 几何积分 (静矩法)
        for sz1, sz2, sv1, sv2 in active_sub_segments:
            h = abs(sz1 - sz2)
            if h < 1e-8: continue
            # 计算该子段面积 (合力)
            f_segment = (sv1 + sv2) * h / 2.0
            # 计算该子段形心位置 (相对于子段顶部的距离)
            # 梯形形心公式: d = (h/3) * (v1 + 2*v2) / (v1 + v2)
            if (sv1 + sv2) == 0: continue
            d_from_top = (h / 3.0) * (sv1 + 2 * sv2) / (sv1 + sv2)
            # 换算为绝对标高并累加力矩 (带符号净弯矩, 跨参考标高时方向相减)
            centroid_alt = sz1 - d_from_top
            total_force += f_segment
            total_moment += f_segment * (centroid_alt - reference_alt)
    # 合成总形心距离 (合力点), 力臂取净弯矩绝对值以保证非负
    final_arm = abs(total_moment) / total_force if total_force > 1e-8 else 0.0
    return total_force, final_arm


def polyline_y_at_x(x_target, coords):
    """在折线坐标序列 [(x0,y0), (x1,y1), ...] 中线性插值 x_target 处的 y 值。

    纯几何函数，不涉及 Excel/UI：
      - x 在折线 x 范围内 → 返回对应 y（线性插值）；
      - x 超出范围 → 返回 None（调用方据此判断该土层在此桩号不存在）。
    coords 无序亦可（内部按 x 排序）。
    """
    if not coords:
        return None
    sp = sorted(coords, key=lambda p: p[0])
    xs = [p[0] for p in sp]
    ys = [p[1] for p in sp]
    if x_target < xs[0] or x_target > xs[-1]:
        return None
    for i in range(len(xs) - 1):
        if xs[i] <= x_target <= xs[i + 1]:
            if abs(xs[i + 1] - xs[i]) < 1e-10:
                return ys[i]
            return ys[i] + (ys[i + 1] - ys[i]) * (x_target - xs[i]) / (xs[i + 1] - xs[i])
    return None
