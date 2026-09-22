# 1. 标准库
import math
import operator
from functools import reduce

# 2. 第三方库
import numpy as np


#根据1m弧长分割圆弧
def arc_points_by_length(arc_length, radius, start_angle, end_angle, center=(0, 0)):
    # 计算圆弧的总角度
    angle_degrees = abs(end_angle - start_angle)
    # 将角度转换为弧度
    angle_radians = math.radians(angle_degrees)
    # 将弧度转换成角度
    # angle_degrees = math.degrees(angle_radians)
    # 计算圆弧的长度
    arc_length = min(arc_length, radius * angle_radians)  # 防止arc_length大于圆周长
    # 计算每份的弧度：1弧度 = 180/3.14 = 57.3度
    segment_length = arc_length / radius
    # 生成所有分割点的坐标
    points = []
    current_angle = start_angle
    while current_angle < end_angle:
        x = center[0] + radius * math.cos(math.radians(current_angle))
        y = center[1] + radius * math.sin(math.radians(current_angle))
        points.append((round(x,1), round(y,1)))
        current_angle += segment_length * 180 / math.pi
    # 添加终止点
    x = center[0] + radius * math.cos(math.radians(end_angle))
    y = center[1] + radius * math.sin(math.radians(end_angle))
    points.append((round(x,1), round(y,1)))
    points = set(points)
    return points
# # 使用示例
# arc_length = 1000  # 将圆弧分割成弧长为1000mm的份数
# radius = 15000  # 假设圆的半径
# start_angle = 0  # 起始角度
# end_angle = 90  # 终止角度
# points = arc_points_by_length(arc_length, radius, start_angle, end_angle)
# for point in points:
#     print(point)


#原点、第一点、第二点
def angle_between_points_relative_to_origin(point0, point1, point2):
    # 计算向量AB和AC
    x1 = point0[0]
    y1 = point0[1]
    x2 = point1[0]
    y2 = point1[1]
    x3 = point2[0]
    y3 = point2[1]
    vec_ab = (x2 - x1, y2 - y1)
    vec_ac = (x3 - x1, y3 - y1)
    # 计算点积
    dot_product = vec_ab[0] * vec_ac[0] + vec_ab[1] * vec_ac[1]
    # 计算向量模长
    magnitude_ab = math.sqrt(vec_ab[0]**2 + vec_ab[1]**2)
    magnitude_ac = math.sqrt(vec_ac[0]**2 + vec_ac[1]**2)
    # 计算夹角的余弦值
    cos_theta = dot_product / (magnitude_ab * magnitude_ac)
    # 计算夹角的角度值
    angle_degrees = math.degrees(math.acos(cos_theta))
    return angle_degrees


# 任意两点组成的射线与x正轴的夹角，point1为射线原点
def angle_between_line_and_X_axis(point1, point2):
    # 使用反正切函数计算夹角
    angle_radians = math.atan2(point2[1] - point1[1], point2[0] - point1[0])
    # 将弧度转换为度
    angle_degrees = math.degrees(angle_radians)
    return angle_degrees
# 使用示例
# A = [-1,-1,0]
# B = [1,1,0]
# angle = angle_between_line_and_x_axis(A, B)
# print(f"两点连成的线与x轴的夹角: {angle}度")


#点按逆时针排序
def Points_Counterclockwise(coords):
    center = tuple(map(operator.truediv, reduce(lambda x, y: map(operator.add, x, y), coords), [len(coords)] * 2))
    return sorted(coords, key=lambda coord: (-135 - math.degrees(math.atan2(*tuple(map(operator.sub, coord, center))[::-1]))) % 360, reverse=True)


#重新排列表格
def change_list_order(point,ptlist):
    position = ptlist.index(point)
    if position != 0:
        new_list = ptlist[position:]+ptlist[0:position]
    else:
        new_list = ptlist
    return new_list


#去除列表重复元素并保留原始顺序
def ListToTuple_with_OriginalSequence(originlst):
    seen = set()
    tuple_list = []
    for item in originlst:
        if item not in seen:
            seen.add(item)
            tuple_list.append(item)
    return tuple_list


# 计算point_list中与point_x最短距离的点，因为在函数中起点是先x后y坐标最大的点，规定连线方向为顺时针，所以下一个点的y坐标必须小于该点的y坐标
def find_closest_point(point, point_list):
    min_distance = math.inf
    closest_point = None
    for p in point_list:
        #if p[1] < point[1]:
            distance = math.sqrt((p[0] - point[0])**2 + (p[1] - point[1])**2)
            if distance < min_distance:
                min_distance = distance
                closest_point = p
    return closest_point


#判断两张表中x坐标或y坐标相同的点，并将这些点赋予到第一张表中
def LookForTheSame_XorY(lst1,lst2):
    #得到lst1中所有的x和y
    lst1_X = list(set([x for x,y,z in lst1]))
    lst1_X.sort()
    lst1_Y = list(set([y for x,y,z in lst1]))
    lst1_Y.sort()
    lst2_X = []
    lst2_Y = []
    for group in lst2:
        for x in group:
            #对于每个点，如果x坐标相等
            if x[0] in lst1_X and abs(x[1]) <= abs(lst1_Y[0]):
                lst2_X.append(x)
            elif x[1] in lst1_Y and abs(x[0]) <= abs(lst1_X[0]):
                lst2_Y.append(x)
        #返回x和y坐标相同的点的表
    return lst2_X+lst2_Y


#根据表和正负处理数据：paramA,paramB为x和y轴正负的参数，1代表正，0代表负
def Offset_list(lst,O,paramA,paramB):
    x_coords = [point[0] for point in lst]
    y_coords = [point[1] for point in lst]
    new_points_lst = []
    for point in lst:
        x, y, z = point
        if paramA == 1 and paramB == 1:
            x_max = max(x_coords)
            y_max = max(y_coords)
            if x == x_max:
                new_points = [x+O, y, z]
                new_points_lst.append(new_points)
            if y == y_max:
                new_points = [x, y+O, z]
                new_points_lst.append(new_points)
        if paramA == 0 and paramB == 1:
            x_max = min(x_coords)
            y_max = max(y_coords)
            if x == x_max:
                new_points = [x-O, y, z]
                new_points_lst.append(new_points)
            if y == y_max:
                new_points = [x, y+O, z]
                new_points_lst.append(new_points)
        if paramA == 0 and paramB == 0:
            x_max = min(x_coords)
            y_max = min(y_coords)
            if x == x_max:
                new_points = [x-O, y, z]
                new_points_lst.append(new_points)
            if y == y_max:
                new_points = [x, y-O, z]
                new_points_lst.append(new_points)
        if paramA == 1 and paramB == 0:
            x_max = max(x_coords)
            y_max = min(y_coords)
            if x == x_max:
                new_points = [x+O, y, z]
                new_points_lst.append(new_points)
            if y == y_max:
                new_points = [x, y-O, z]
                new_points_lst.append(new_points)
    return new_points_lst,x_max,y_max 


#对一个图形上的所有点进行偏移
def Offset_points(oldlst,elder_lst,O_dist):
    quadrants = {
        'First_Quadrant': [],#第一象限的点
        'Second_quadrant': [],#第二象限的点
        'Third_Quadrant': [],#第三象限的点
        'Fourth_Quadrant': [],#第四象限的点
        'X_Positive': [],#x正轴上的点
        'X_Negative': [],#x负轴上的点
        'Y_Positive': [],#y正轴上的点
        'Y_Negative': [],#y负轴上的点
        }
    for point in oldlst:
        x, y, z = point
        # 根据x和y的值判断象限
        if x > 0 and y > 0:
            quadrants['First_Quadrant'].append(point)#第一象限的点
        elif x < 0 and y > 0:
            quadrants['Second_quadrant'].append(point)#第二象限的点
        elif x < 0 and y < 0:
            quadrants['Third_Quadrant'].append(point)#第三象限的点
        elif x > 0 and y < 0:
            quadrants['Fourth_Quadrant'].append(point)#第四象限的点
        elif y == 0 and x > 0:
            quadrants['X_Positive'].append(point)#x正轴上的点
        elif y == 0 and x < 0:
            quadrants['X_Negative'].append(point)#x负轴上的点
        elif x == 0 and y > 0:
            quadrants['Y_Positive'].append(point)#y正轴上的点
        elif x == 0 and y < 0:
            quadrants['Y_Negative'].append(point)#y负轴上的点
        else:
            pass
    #返回后的表中的坐标已经是根据象限偏心过的了
    list1 = Offset_list(quadrants['First_Quadrant'],O_dist,1,1)
    list2 = Offset_list(quadrants['Second_quadrant'],O_dist,0,1)
    list3 = Offset_list(quadrants['Third_Quadrant'],O_dist,0,0)
    list4 = Offset_list(quadrants['Fourth_Quadrant'],O_dist,1,0)
    list5 = [[point[0]+O_dist,point[1],point[2]] for point in quadrants['X_Positive']]
    list6 = [[point[0]-O_dist,point[1],point[2]] for point in quadrants['X_Negative']]
    list7 = [[point[0],point[1]+O_dist,point[2]] for point in quadrants['Y_Positive']]
    list8 = [[point[0],point[1]-O_dist,point[2]] for point in quadrants['Y_Negative']]
    #给四个角点增加偏移量
    ptlst = []
    for pt in elder_lst:
        if pt[0]>0 and pt[1]>0:
            ptlst.append([pt[0]+O_dist,pt[1]+O_dist,pt[2]])
        elif pt[0]>0 and pt[1]<0:
            ptlst.append([pt[0]+O_dist,pt[1]-O_dist,pt[2]])
        elif pt[0]<0 and pt[1]<0:
            ptlst.append([pt[0]-O_dist,pt[1]-O_dist,pt[2]])
        else:
            ptlst.append([pt[0]-O_dist,pt[1]+O_dist,pt[2]])
    #将所有表相加并顺时针排序
    list_apply = list1[0]+list2[0]+list3[0]+list4[0]+list5+list6+list7+list8+ptlst
    list_apply = list(set(tuple(point) for point in list_apply))
    layer_pointlst_sort = []
    list_apply.sort(key=lambda p: (-p[0], -p[1]))
    layer_startpoint = list_apply[0]
    layer_pointlst_sort = change_list_order(layer_startpoint,Points_Counterclockwise(list_apply))
    return layer_pointlst_sort


#4个坐标的中间坐标
def get_midpoint_within_4points(alst):
    x = 0
    y = 0
    z = 0
    for x0 in alst:
        x = x+x0[0]
        y = y+x0[1]
        z = 0
    #结果均保留一位小数
    x = round(x/4,1)
    y = round(y/4,1)
    point = (x,y,z)
    return point





def normalize(v):
    norm = np.linalg.norm(v)
    return v / norm if norm != 0 else v


def transform_to_new_coordinate_system(lst, p):
    # 计算新的基向量
    p1 = np.array(lst[0])
    p2 = np.array(lst[1])
    p3 = np.array(lst[2])
    x_prime = p2 - p1
    x_prime = normalize(x_prime)
    z_prime = np.cross(p2 - p1, p3 - p1)
    z_prime = normalize(z_prime)
    y_prime = np.cross(z_prime, x_prime)
    y_prime = normalize(y_prime)
    # 计算转换矩阵
    transformation_matrix = np.array([x_prime, y_prime, z_prime]).T
    # 将点 p 转换到新坐标系
    p_relative = p - p1
    p_new = np.dot(transformation_matrix.T, p_relative)
    #return p_new, transformation_matrix, x_prime, y_prime, z_prime
    #把向量表示的矩阵点坐标转换为元组点坐标
    return tuple(p_new)