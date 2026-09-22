# 1. 标准库
import csv
from io import StringIO

# 2. 第三方库
import pyperclip

# 3. 本地模块
from FEM_MidasCivil.Double_Walls_CofferDam_FEA.supports import transform_to_new_coordinate_system
from FEM_MidasCivil.Double_Walls_CofferDam_FEA.get_data_from_cad import Handle_Cad_PlanView_LinesToPoints,Handle_Cad_ElevationView_LinesToPoints,Points_Counterclockwise,change_list_order


#材料定义函数
def MATERIAL():
    materials = [
        "1, STEEL, Q235, 0, 0, , C, NO, 0.02, 1, GB03(S), , Q235, NO, 206",
	]
    return materials


#截面定义函数
def support_section_data():
    section_list=[
        "1, DBUSER, 水平环板, LC, 0, 0, 0, 0, 0, 0, YES, NO, SB , 2, 20, 500, 0, 0, 0, 0, 0, 0, 0, 0",
        "2, DBUSER, 水平桁架, CC, 0, 0, 0, 0, 0, 0, YES, NO, 2C , 1, GB-YB05, 2C 10",
        "3, DBUSER, 内支撑, CC, 0, 0, 0, 0, 0, 0, YES, NO, P  , 2, 1000, 12, 0, 0, 0, 0, 0, 0, 0, 0",
    ]
    return section_list


#定义板厚函数：1, VALUE, YES, 6, 0,  NO, 0, 0
def support_plate_data():
    section_list=[
        "1, STIFFENED, USER, LOWER, 20",
        "NO, , 0,  0, 0, 0, 0, 0, 0",
        "YES, L, 300,  63, 63, 6, 6, 0, 0",
        #"2, VALUE, YES, 14, 0, NO, 0, 0",
        "2, STIFFENED, USER, LOWER, 14",
        "NO, , 0,  0, 0, 0, 0, 0, 0",
        "YES, L, 300,  63, 63, 6, 6, 0, 0",
    ]
    return section_list


#生成所有点的坐标
def make_point_element(acadapp):
    #与平面图交互得到点的信息
    #返回layer1_pointlst_sort, layer2_pointlst_sort，外和内壁板的按线连续的点坐标
    Handle_Cad_PlanView_LinesToPoints_resultlst = Handle_Cad_PlanView_LinesToPoints(acadapp)
    #与立面图交互得到点的信息
    #返回apply_ylist，围堰立面图上壁板、内支撑和水平环板的对应的所有的y坐标
    Handle_Cad_ElevationView_LinesToPoints_resultlst = Handle_Cad_ElevationView_LinesToPoints(acadapp)
    ###
    list_p1 = Handle_Cad_PlanView_LinesToPoints_resultlst[0]#一层外壁板的所有点
    list_p2 = Handle_Cad_PlanView_LinesToPoints_resultlst[1]#一层内壁版的所有点
    list_p3 = Handle_Cad_PlanView_LinesToPoints_resultlst[2]#内支撑的点，嵌套列表
    list_p4 = Handle_Cad_PlanView_LinesToPoints_resultlst[3]#隔舱板的点，嵌套列表
    list_p5 = Handle_Cad_PlanView_LinesToPoints_resultlst[4]#大内撑的点，嵌套列表
    list_p6 = Handle_Cad_PlanView_LinesToPoints_resultlst[5]#底层刃脚的点
    list_p7 = Handle_Cad_PlanView_LinesToPoints_resultlst[6]#中层刃脚的点
    Outer_plate_ConnorPoints = Handle_Cad_PlanView_LinesToPoints_resultlst[7]#外壁板的角点
    Inner_plate_ConnorPoints = Handle_Cad_PlanView_LinesToPoints_resultlst[8]#内壁板的角点
    list_z = Handle_Cad_ElevationView_LinesToPoints_resultlst[0]#高度上的z坐标
    list_z1 = Handle_Cad_ElevationView_LinesToPoints_resultlst[2]#围堰水平环板处的高度上的z坐标
    list_z2 = Handle_Cad_ElevationView_LinesToPoints_resultlst[3]#围堰大内撑处的高度上的z坐标
    list_z3 = Handle_Cad_ElevationView_LinesToPoints_resultlst[4]#围堰底层刃脚的z坐标
    list_z4 = Handle_Cad_ElevationView_LinesToPoints_resultlst[5]#围堰中层刃脚的z坐标
    nodei = "1"
    elementi = "1"
    startnodei = "1"
    #points = []
    nodes = []
    Outer_plate_nodei_lst = []#外壁板第一层的节点号
    Inner_plate_nodei_lst = []#内壁板第一层的节点号
    Inner_plate_connor_nodei_lst = []#内壁板角点的节点号
    elements = []
    nodei_RingPlate_LST1 = []#外壁板上每层水平环板的节点号
    nodei_RingPlate_LST2 = []#内壁板上每层水平环板的节点号
    #生成外壁板节点
    for z_times in range(len(list_z)):
        for pt in list_p1:
            point = [pt[0],pt[1],list_z[z_times]]
            nodestring = f"{nodei} {","} {point[0]} {","} {point[1]} {","} {point[2]}"
            #如果y坐标是水平环板处的y坐标
            if point[2] in list_z1:
                nodei_RingPlate_LST1.append(nodei)
            if point[2] == 0:#外壁板第一层，用于刃脚单元生成
                Outer_plate_nodei_lst.append(nodei)
            #points.append(point)
            nodes.append(nodestring)
            nodei = str(int(nodei)+1)
    #生成外壁板的板单元
    for z in range(len(list_z)-1):
         present_startnodei = startnodei#记录下每一层循环开始时的起始节点号
         for x in range(len(list_p1)):
            #不是最后一次闭合图形时,PLATE代表板单元，1代表材料号，1代表截面号，2代表thin/1代表thick
            if x != len(list_p1)-1:
                elementstring = f"{elementi} {", PLATE, 1, 1,"} {startnodei} {","} {str(int(startnodei)+1)} {","} {str(int(startnodei)+len(list_p1)+1)} {","} {str(int(startnodei)+len(list_p1))} {", 1, 0"}"
                elements.append(elementstring)
                elementi = str(int(elementi)+1)
                startnodei = str(int(startnodei)+1)
            #最后一次闭合时  
            else:
                elementstring = f"{elementi} {", PLATE, 1, 1,"} {startnodei} {","} {present_startnodei} {","} {str(int(present_startnodei)+len(list_p1))} {","} {str(int(startnodei)+len(list_p1))} {", 1, 0"}"
                elements.append(elementstring)
                elementi = str(int(elementi)+1)
                startnodei = str(int(startnodei)+1)
    #生成外壁板上的水平环板
    nodei_RingPlate_LST1 = [list(group) for group in zip(*([iter(nodei_RingPlate_LST1)] * (len(nodei_RingPlate_LST1)//len(list_z1))))]
    for nodeilist in nodei_RingPlate_LST1:
        for nodei_x in range(len(nodeilist)):
            if nodei_x != len(nodeilist)-1:
                elementstring = f"{elementi} {", BEAM, 1, 1,"} {nodeilist[nodei_x]} {","} {nodeilist[nodei_x+1]} {", 0, 0"}"
                elements.append(elementstring)
                elementi = str(int(elementi)+1)
            else:
                elementstring = f"{elementi} {", BEAM, 1, 1,"} {nodeilist[nodei_x]} {","} {nodeilist[0]} {", 0, 0"}"
                elements.append(elementstring)
                elementi = str(int(elementi)+1)
    #生成内壁板节点
    startnodei = nodei#将上一个循环结束后的初始节点号设置为单元的开始节点号
    for z_times in range(len(list_z)):
        for pt in list_p2:
            point = [pt[0],pt[1],list_z[z_times]]
            nodestring = f"{nodei} {","} {point[0]} {","} {point[1]} {","} {point[2]}"
            #如果y坐标是水平环板处的y坐标
            if point[2] in list_z1:
                nodei_RingPlate_LST2.append(nodei)
            if point[2] == 0:#内壁板第一层，用于刃脚单元生成
                Inner_plate_nodei_lst.append(nodei)
                if tuple(point) in Inner_plate_ConnorPoints:#元组和列表形式需要转化，否则点相同也为False
                    Inner_plate_connor_nodei_lst.append(nodei)#角点节点号记录
            #points.append(point)
            nodes.append(nodestring)
            nodei = str(int(nodei)+1)
    #生成内壁板的板单元
    for z in range(len(list_z)-1):
         present_startnodei = startnodei#记录下每一层循环开始时的起始节点号
         for x in range(len(list_p2)):
            #不是最后一次闭合图形时,PLATE代表板单元，1代表材料号，1代表截面号，2代表thin/1代表thick
            if x != len(list_p2)-1:
                elementstring = f"{elementi} {", PLATE, 1, 1,"} {startnodei} {","} {str(int(startnodei)+1)} {","} {str(int(startnodei)+len(list_p2)+1)} {","} {str(int(startnodei)+len(list_p2))} {", 1, 0"}"
                elements.append(elementstring)
                elementi = str(int(elementi)+1)
                startnodei = str(int(startnodei)+1)
            #最后一次闭合时  
            else:
                elementstring = f"{elementi} {", PLATE, 1, 1,"} {startnodei} {","} {present_startnodei} {","} {str(int(present_startnodei)+len(list_p2))} {","} {str(int(startnodei)+len(list_p2))} {", 1, 0"}"
                elements.append(elementstring)
                elementi = str(int(elementi)+1)
                startnodei = str(int(startnodei)+1)
    #生成内壁版上的水平环板
    nodei_RingPlate_LST2 = [list(group) for group in zip(*([iter(nodei_RingPlate_LST2)] * (len(nodei_RingPlate_LST2)//len(list_z1))))]
    for nodeilist in nodei_RingPlate_LST2:
        for nodei_x in range(len(nodeilist) - 1, -1, -1):
            if nodei_x != len(nodeilist)-1:
                elementstring = f"{elementi} {", BEAM, 1, 1,"} {nodeilist[nodei_x+1]} {","} {nodeilist[nodei_x]} {", 0, 0"}"
                elements.append(elementstring)
                elementi = str(int(elementi)+1)
            else:
                elementstring = f"{elementi} {", BEAM, 1, 1,"} {nodeilist[0]} {","} {nodeilist[nodei_x]} {", 0, 0"}"
                elements.append(elementstring)
                elementi = str(int(elementi)+1)
    # #生成中层和底层刃脚的节点
    Inner_Cutting_Edge_bottom_nodei_lst = []#内壁板下刃脚底层
    Inner_Cutting_Edge_middle_nodei_lst = []#内壁板下刃脚中间层
    Outer_Cutting_Edge_bottom_nodei_lst = []#外壁板下刃脚底层
    Outer_Cutting_Edge_middle_nodei_lst = []#外壁板下刃脚中间层
    #内壁板下中间层节点
    for pt in list_p7:
        point = [pt[0],pt[1],list_z4]
        nodestring = f"{nodei} {","} {point[0]} {","} {point[1]} {","} {point[2]}"
        Inner_Cutting_Edge_middle_nodei_lst.append(nodei)
        nodes.append(nodestring)
        nodei = str(int(nodei)+1)
    #内壁板下底层节点
    for pt in list_p6:
        point = [pt[0],pt[1],list_z3]
        nodestring = f"{nodei} {","} {point[0]} {","} {point[1]} {","} {point[2]}"
        Inner_Cutting_Edge_bottom_nodei_lst.append(nodei)
        nodes.append(nodestring)
        nodei = str(int(nodei)+1)
    #外壁板下节点
    for pt in list_p1:
        for z in [list_z4,list_z3]:
            point = [pt[0],pt[1],z]
            nodestring = f"{nodei} {","} {point[0]} {","} {point[1]} {","} {point[2]}"
            #中间层
            if z == list_z4:
                Outer_Cutting_Edge_middle_nodei_lst.append(nodei)
            #底层
            else:
                Outer_Cutting_Edge_bottom_nodei_lst.append(nodei)
            nodes.append(nodestring)
            nodei = str(int(nodei)+1)
    #建立刃脚单元
    #获取内壁板第一层、外壁板第一层、内壁板下刃脚中间层、内壁板下刃脚底层、外壁板下刃脚中间层，外壁板下刃脚底层的节点号
    #Inner_plate_nodei_lst, Outer_plate_nodei_lst,
    #Inner_Cutting_Edge_middle_nodei_lst, Inner_Cutting_Edge_bottom_nodei_lst, 
    #Outer_Cutting_Edge_middle_nodei_lst, Outer_Cutting_Edge_bottom_nodei_lst,
    #先建立内壁板下的刃脚板单元：
    lst2_node_id_i = 0#第二张表中节点号的定位位置
    for node_id_i1 in range(len(Inner_plate_nodei_lst)):
        #如果该节点号是第一层内壁板的角点节点号
        if Inner_plate_nodei_lst[node_id_i1] in Inner_plate_connor_nodei_lst:
            #第一个角点
            if node_id_i1 == 0:
                elementstring = f"{elementi} {", PLATE, 1, 1,"} {Inner_plate_nodei_lst[0]} {","} {Inner_Cutting_Edge_middle_nodei_lst[0]} {","} {Inner_Cutting_Edge_middle_nodei_lst[-1]} {", 0, 1, 0"}"
                elements.append(elementstring)
                elementi = str(int(elementi)+1)
                elementstring = f"{elementi} {", PLATE, 1, 1,"} {Inner_plate_nodei_lst[0]} {","} {Inner_Cutting_Edge_middle_nodei_lst[0]} {","} {Inner_Cutting_Edge_middle_nodei_lst[1]} {", 0, 1, 0"}"
                elements.append(elementstring)
                elementi = str(int(elementi)+1)
                lst2_node_id_i = lst2_node_id_i+1
            #其余角点
            else:
                elementstring = f"{elementi} {", PLATE, 1, 1,"} {Inner_plate_nodei_lst[node_id_i1]} {","} {Inner_Cutting_Edge_middle_nodei_lst[lst2_node_id_i]} {","} {Inner_Cutting_Edge_middle_nodei_lst[lst2_node_id_i+1]} {", 0, 1, 0"}"
                elements.append(elementstring)
                elementi = str(int(elementi)+1)
                lst2_node_id_i = lst2_node_id_i+1
                elementstring = f"{elementi} {", PLATE, 1, 1,"} {Inner_plate_nodei_lst[node_id_i1]} {","} {Inner_Cutting_Edge_middle_nodei_lst[lst2_node_id_i]} {","} {Inner_Cutting_Edge_middle_nodei_lst[lst2_node_id_i+1]} {", 0, 1, 0"}"
                elements.append(elementstring)
                elementi = str(int(elementi)+1)
                lst2_node_id_i = lst2_node_id_i+1
        #非角点号
        #非末尾
        if node_id_i1 != len(Inner_plate_nodei_lst)-1:
            elementstring = f"{elementi} {", PLATE, 1, 1,"} {Inner_plate_nodei_lst[node_id_i1]} {","} {Inner_Cutting_Edge_middle_nodei_lst[lst2_node_id_i]} {","} {Inner_Cutting_Edge_middle_nodei_lst[lst2_node_id_i+1]} {","} {Inner_plate_nodei_lst[node_id_i1+1]} {", 1, 0"}"
            elements.append(elementstring)
            elementi = str(int(elementi)+1)
            lst2_node_id_i = lst2_node_id_i+1
        #末尾
        else:
            elementstring = f"{elementi} {", PLATE, 1, 1,"} {Inner_plate_nodei_lst[node_id_i1]} {","} {Inner_Cutting_Edge_middle_nodei_lst[lst2_node_id_i]} {","} {Inner_Cutting_Edge_middle_nodei_lst[lst2_node_id_i+1]} {","} {Inner_plate_nodei_lst[0]} {", 1, 0"}"
            elements.append(elementstring)
            elementi = str(int(elementi)+1)
            lst2_node_id_i = lst2_node_id_i+1
    #内壁板下刃脚底层和中层的板单元：
    make_plate_with_list_resultlst1 = make_plate_with_list(Inner_Cutting_Edge_middle_nodei_lst,Inner_Cutting_Edge_bottom_nodei_lst,elementi,1,1)
    elements = elements + make_plate_with_list_resultlst1[0]#合并单元信息
    elementi = make_plate_with_list_resultlst1[1]#已经+1后的起始单元号
    #外壁板下的刃脚板单元：
    make_plate_with_list_resultlst2 = make_plate_with_list(Outer_plate_nodei_lst,Outer_Cutting_Edge_middle_nodei_lst,elementi,1,1)
    elements = elements + make_plate_with_list_resultlst2[0]#合并单元信息
    elementi = make_plate_with_list_resultlst2[1]#已经+1后的起始单元号
    #外壁板下刃脚底层和中层的板单元：
    make_plate_with_list_resultlst3 = make_plate_with_list(Outer_Cutting_Edge_middle_nodei_lst,Outer_Cutting_Edge_bottom_nodei_lst,elementi,1,1)
    elements = elements + make_plate_with_list_resultlst3[0]#合并单元信息
    elementi = make_plate_with_list_resultlst3[1]#已经+1后的起始单元号
    #刃脚底板：
    make_plate_with_list_resultlst4 = make_plate_with_list(Inner_Cutting_Edge_bottom_nodei_lst,Outer_Cutting_Edge_bottom_nodei_lst,elementi,1,1)
    elements = elements + make_plate_with_list_resultlst4[0]#合并单元信息
    elementi = make_plate_with_list_resultlst4[1]#已经+1后的起始单元号
    #得到[节点号,x,y,z]的表
    #去空格
    nodes_string_list = [s.replace(' ', '') for s in nodes]
    nodes_table = []
    # 使用csv.reader解析字符串
    for stringi in nodes_string_list:
        # 使用StringIO将字符串转换为文件对象
        fake_file = StringIO(stringi)
        reader = csv.reader(fake_file, delimiter=',')
        # 读取转换后的行数据
        for row in reader:
            nodes_table.append(row)
    #生成内支撑
    first_nodei_lst2_1 = []
    for ptlst in list_p3:
        first_nodei_lst_1 = []
        for pt in ptlst:
            first_nodei_1 = []
            #点的x和y坐标能对应并且y坐标是水平环板的y坐标
            # [first_nodei_1.append(sublist[0]) for sublist in nodes_table if sublist[1:3] == [str(pt[0]),str(pt[1])] and float(sublist[-1]) in list_z1]
            [first_nodei_1.append(sublist[0]) for sublist in nodes_table if abs(float(sublist[1]) - pt[0]) < 5 and abs(float(sublist[2]) - pt[1]) < 5 and float(sublist[-1]) in list_z1]
            first_nodei_lst_1.append(first_nodei_1)
        first_nodei_lst2_1.append(first_nodei_lst_1)
    #有几段多段线就有几组，一组中的每张表代表了在该位置上的所有竖向的节点号
    #[['85', '155'], ['339', '387'], ['84', '154'], ['338', '386'], ['83', '153']]这一组就代表了一根多段线，list['85', '155']代表在该点的坐标下有两层内支撑
    #85代表该层内支撑的顶点和壁板相交处的节点号，将list[i]依次相连则得到顺次的内支撑单元连接顺序，85,339,84,338,83和155,387,154,386,153就是同一竖直投影下不同高度的两层内支撑
    for x in first_nodei_lst2_1:#x = [['85', '155'], ['339', '387'], ['84', '154'], ['338', '386'], ['83', '153']]
        # print(x)
        for i in range(len(x[0])):
            for y in range(len(x)-1):
                elementstring = f"{elementi} {", BEAM, 1, 2,"} {x[y][i]} {","} {x[y+1][i]} {", 90, 0"}"
                elements.append(elementstring)
                elementi = str(int(elementi)+1)
    #生成隔舱板
    #确定一个隔舱板的两个端点，以x和y坐标相同去和nodes_table比较获取节点号，根据两排节点号依次生成板单元
    first_nodei_lst2_2 = []
    Outer_Cutting_Edge_bottom_nodei_lst_p4 = []
    Outer_Cutting_Edge_middle_nodei_lst_p4 = []
    Inner_Cutting_Edge_bottom_nodei_lst_p4 = []
    Inner_Cutting_Edge_middle_nodei_lst_p4 = []
    for pt_group in list_p4:#pt_group = [(0.0, -9600.0, 0.0), (0.0, -11600.0, 0.0)]
        first_nodei_lst_2 = []
        for ptx in pt_group:#x = (0.0, -9600.0, 0.0)
            first_nodei_2 = []
            #点的x和y坐标能对应并且y坐标是水平环板的y坐标
            #[first_nodei_2.append(sublist[0]) for sublist in nodes_table if sublist[1:3] == [str(ptx[0]),str(ptx[1])]]
            for sublist in nodes_table:
                if sublist[1:3] == [str(ptx[0]), str(ptx[1])]:
                    first_nodei_2.append(sublist[0])#得到节点号
                    if sublist[3] == str(list_z3):#刃脚底层
                        Outer_Cutting_Edge_bottom_nodei_lst_p4.append(sublist[0])#记录隔舱板在刃脚外壁板底层的节点号组
            #[print(ptx,sublist) for sublist in nodes_table if sublist[1:3] == [str(ptx[0]),str(ptx[1])]]
            first_nodei_lst_2.append(first_nodei_2)
        first_nodei_lst2_2.append(first_nodei_lst_2)
    #得到刃脚和隔舱板有关的节点并插入到原表中
    for x in Outer_Cutting_Edge_bottom_nodei_lst_p4:
        nodei_position = Outer_Cutting_Edge_bottom_nodei_lst.index(x)
        Outer_Cutting_Edge_middle_nodei_lst_p4.append(Outer_Cutting_Edge_middle_nodei_lst[nodei_position])
        Inner_Cutting_Edge_bottom_nodei_lst_p4.append(Inner_Cutting_Edge_bottom_nodei_lst[nodei_position])
        Inner_Cutting_Edge_middle_nodei_lst_p4.append(Inner_Cutting_Edge_middle_nodei_lst[nodei_position])
    for x in range(len(first_nodei_lst2_2)):#len(first_nodei_lst2_2)=隔舱板数量
        for y in range(len(first_nodei_lst2_2[x])):#len(x)=2
            if y == 0:
                first_nodei_lst2_2[x][y].insert(0,Inner_Cutting_Edge_middle_nodei_lst_p4[x])
                first_nodei_lst2_2[x][y].insert(0,Inner_Cutting_Edge_bottom_nodei_lst_p4[x])
            elif y ==1:
                first_nodei_lst2_2[x][y].insert(0,Outer_Cutting_Edge_middle_nodei_lst_p4[x])
                first_nodei_lst2_2[x][y].insert(0,Outer_Cutting_Edge_bottom_nodei_lst_p4[x])
    #print(first_nodei_lst2_2)
    for x in first_nodei_lst2_2:#x = [['626', '674', '722'], ['96', '166', '236']]，3点说明有两块板，顺序为626-674-166-96，674-722-236-166
        for i in range(len(x[0])-1):
            for y in range(len(x)-1):
                elementstring = f"{elementi} {", PLATE, 1, 2,"} {x[y][i]} {","} {x[y][i+1]} {","} {x[y+1][i+1]} {","} {x[y+1][i]} {", 2, 0"}"
                elements.append(elementstring)
                elementi = str(int(elementi)+1)
    #生成大内撑
    first_nodei_lst2_3 = []
    for ptlst in list_p5:
        first_nodei_lst_3 = []
        for pt in ptlst:
            first_nodei_3 = []
            #点的x和y坐标能对应并且y坐标是水平环板的y坐标
            [first_nodei_3.append(sublist[0]) for sublist in nodes_table if sublist[1:3] == [str(pt[0]),str(pt[1])] and float(sublist[-1]) in list_z2]
            first_nodei_lst_3.append(first_nodei_3)
        first_nodei_lst2_3.append(first_nodei_lst_3)
    #生成单元
    for x in first_nodei_lst2_3:
        for i in range(len(x[0])):
            for y in range(len(x)-1):
                elementstring = f"{elementi} {", BEAM, 1, 3,"} {x[y][i]} {","} {x[y+1][i]} {", 0, 0"}"
                elements.append(elementstring)
                elementi = str(int(elementi)+1)
    #外壁板的四个角点+刃脚底层z坐标+外壁板顶层z坐标=八个角点的点坐标
    #面数
    Faces = len(Outer_plate_ConnorPoints)
    #最左下角点
    first_node_in_Outer_plate_ConnorPoints = min(Outer_plate_ConnorPoints, key=lambda point: point[0] + point[1])
    #以最左下角点为起点的逆时针排序
    Outer_plate_ConnorPoints = change_list_order(first_node_in_Outer_plate_ConnorPoints,Points_Counterclockwise(Outer_plate_ConnorPoints))
    #得到围堰最低点和最高点的z坐标
    zmin_max = [list_z3,max(list_z)]
    #根据面数生成每个面上的四个角点的坐标
    #先得到最上层和最下层的角点的坐标
    Outer_plate_ConnorPoints_bottom_top_lst = []
    for layer in zmin_max:#层数，一般为两层，最顶层和最底层
        Outer_plate_ConnorPoints_bottom_top = []
        for pt in Outer_plate_ConnorPoints:
            Outer_plate_ConnorPoints_bottom_top.append([pt[0],pt[1],layer])
        Outer_plate_ConnorPoints_bottom_top_lst.append(Outer_plate_ConnorPoints_bottom_top)
    #根据面数重排列坐标，转换为一个竖面上的角点坐标,总面数=Faces
    faces_node_list = []
    for face in range(Faces):#面数
        for layer in range(len(Outer_plate_ConnorPoints_bottom_top_lst)-1):#层数，一般为两层
            if face != Faces-1:
                faces_node = [Outer_plate_ConnorPoints_bottom_top_lst[layer][face],
                              Outer_plate_ConnorPoints_bottom_top_lst[layer][face+1],
                              Outer_plate_ConnorPoints_bottom_top_lst[layer+1][face+1],
                              Outer_plate_ConnorPoints_bottom_top_lst[layer+1][face],
                            ]
            else:
                faces_node = [Outer_plate_ConnorPoints_bottom_top_lst[layer][face],
                              Outer_plate_ConnorPoints_bottom_top_lst[layer][0],
                              Outer_plate_ConnorPoints_bottom_top_lst[layer+1][0],
                              Outer_plate_ConnorPoints_bottom_top_lst[layer+1][face],
                            ]
        faces_node_list.append(faces_node)
    #将这些点的坐标转换成一个平面上的坐标，左下角点为坐标原点
    #faces_node_list_transinplate = [list(map(lambda point: [p - h for p, h in zip(point, table[0])], table)) for table in faces_node_list]
    faces_node_list_transinplate = [list(map(lambda point: transform_to_new_coordinate_system(table,point), table)) for table in faces_node_list]
    #根据面数创建全局变量
    #hydrostatic_pressure_i：第i面上的静水压力
    HydrostaticPressure = (zmin_max[-1]-zmin_max[0])/-100000#荷载大小
    PNLOADTYPE_lst = []
    PLANELOAD_lst = []
    for i in range(len(faces_node_list)):
        #荷载类型名
        variable_name = f"{"HydrostaticPressure_"}{i+1}"
        PNLOADTYPE_lst = PNLOADTYPE_lst + PNLOADTYPE(variable_name,faces_node_list_transinplate[i],HydrostaticPressure)
        PLANELOAD_lst = PLANELOAD_lst + PLANELOAD(variable_name,faces_node_list[i])
    return nodes,elements,PNLOADTYPE_lst,PLANELOAD_lst


#定义面荷载类型
#需要确定的参数有外壁板最顶层四角点的坐标以及外壁板刃脚底层四角点的坐标
def PNLOADTYPE(loadname,lst,load):
    lst = [
        f"{"NAME="} {loadname} {", AREA,"}",
        f"{"DATA=NO, NO,"} {lst[0][0]} {","} {lst[0][1]} {","} {load} {","} {lst[1][0]} {","} {lst[1][1]} {","} {load} {","} {lst[2][0]} {","} {lst[2][1]} {", 0"} {","} {lst[3][0]} {","} {lst[3][1]} {", 0"}",
    ]
    return lst


#定义加载面的平面位置
#需要确定的参数有原点坐标，x轴上一点，以及平面上任意一点
def PLANELOAD(loadname,lst):
    lst = [
        f"{"水压力荷载,"} {loadname} {", PLATE, 水压力荷载"}",
        "LPLANE, , 1, NLP, NO, ",
        f"{lst[0][0]} {","} {lst[0][1]} {","} {lst[0][2]} {","} {lst[1][0]} {","} {lst[1][1]} {","} {lst[1][2]} {","} {lst[2][0]} {","} {lst[2][1]} {","} {lst[2][2]} {",  1e-006, NO"}",
    ]
    return lst


#全部汇总，生成mct格式信息流
def makeup_finallst(acadapp):
        make_point_element_resultlst = make_point_element(acadapp)
        finallst = [
            ["*VERSION","8.6.5"],
            ["*UNIT","N,MM,KJ,C"],
            ["*MATERIAL"] + MATERIAL(),
            ["*SECTION"] + support_section_data(),
            ["*THICKNESS"] + support_plate_data(),
            ["*NODE"] + make_point_element_resultlst[0],
            ["*ELEMENT"] + make_point_element_resultlst[1],
            ["*LOAD-GROUP","自重","水压力荷载"],
            ["*STLDCASE","自重, USER,","水压力荷载, USER,"],
            ["*USE-STLD, 自重","*SELFWEIGHT, 0, 0, -1, 自重"],
            #面荷载类型
            ["*PNLOADTYPE"] + make_point_element_resultlst[2],
            #定义面荷载
            ["*PLANELOAD"] + make_point_element_resultlst[3],
        ]
        finallst_to_copy = "\n".join([item for sublist in finallst for item in sublist])
        pyperclip.copy(finallst_to_copy)
        print("粘贴板已复制")


#根据两个有顺序的节点表组成板单元，单元方向为逆时针
def make_plate_with_list(lista,listb,elei,sec_num,thick_or_thin):
    list2 = []
    for ia,ib in zip(range(len(lista)),range(len(listb))):
        if ia == len(lista)-1 and ib == len(listb)-1:
            list1 =[lista[ia],listb[ib],listb[0],lista[0]]
            list2.append(list1)
        else:
            list1 =[lista[ia],listb[ib],listb[ib+1],lista[ia+1]]
            list2.append(list1)
    #print(list2)
    elements = []
    for node_id_lst in list2:
        elementstring = f"{elei} {", PLATE, 1,"} {sec_num} {","} {node_id_lst[0]} {","} {node_id_lst[1]} {","} {node_id_lst[2]} {","} {node_id_lst[3]} {","} {thick_or_thin} {", 0"}"
        elements.append(elementstring)
        elei = str(int(elei)+1)
    #传出单元信息流和最后一个单元号
    return elements,elei