# 1. 标准库
import re
import math
from collections import OrderedDict

# 3. 本地模块
from General.Formula   import cal_ShuiLiuLi_JTS_144_1_2010
from General.AutoCAD   import insert_TXT_incad, get_block_reference_xdata, get_lines_from_block, vtint, vtpnt, vtvariant
from General.DataUtils import inear_interpolation, midasdisttolst, midasdisttolst2, ribdistancelist_trans, trim0, remove_nearby_points, time_str
from General.Geometry  import Bailey_connect_line_with_point_angle, point_coordinate_on_line, intersection_point_of_two_line_segment, if_ptlst_have_pt_in_range, distance, line_intersection, is_point_on_segment, point_in_line_Z_scale

from FEM_MidasCivil.Steel_Pipe_Bailey_Support_FEA.Support_Utils_function import if_continuity, is_element_between_nodes, find_nearest_point_index


# xlst1：需要插值的点组成的表
# xlst2：原点表
# hlst：与原点表对应的高度表
def Linear_interpolation_with_two_points(xlst1, xlst2, hlst):
    # print(xlst1)
    # print(xlst2)
    # print(hlst)
    lst = []
    table1 = []
    table2 = []
    # 定义域
    lim1 = xlst2[0]
    lim2 = xlst2[-1]
    # 判断xlst1在xlst2中的位置
    for i in range(len(xlst1)):
        p = xlst1[i]
        # 当该点处于定义域时
        if p >= lim1 and p <= lim2 and p not in xlst2:
            # print(p)
            for j in range(len(xlst2)):
                # print(xlst2[j])
                if p < xlst2[j]:
                    x1 = xlst2[j - 1]
                    x2 = xlst2[j]
                    h1 = hlst[2*(j - 1) + 1]
                    h2 = hlst[2*j]
                    # print(x1, x2, h1, h2)
                    p1 = p - 0.1
                    p2 = p + 0.1
                    # 求得p点的±0.1的内插高度ph1和ph2
                    ph1 = round(inear_interpolation(x1, x2, h1, h2, p1), 1)
                    ph2 = round(inear_interpolation(x1, x2, h1, h2, p2), 1)
                    
                    lst.append([p, ph1, ph2])
                    # 打断
                    break        
    # print(lst)
    # 循环完成后如果lst不为空，将点和点对应的高度插入原点表和原高度表中对应的位置得到新点表和新高度表
    if lst != []:
        # 插入值
        table1 = xlst2.copy()
        table2 = hlst.copy()
        for pi, ph1i, ph2i in lst:
            # 找到插入位置
            insert_index = 0
            while insert_index < len(table1) and table1[insert_index] < pi:
                insert_index += 1
            # print(insert_index)
            # 插入到第一张表
            table1.insert(insert_index, pi)
            # 插入到第二张表
            table2.insert(insert_index * 2, ph2i)
            table2.insert(insert_index * 2, ph1i)
        # print(table2)
        return table1, table2
    return None


# 根据贝雷梁的y坐标和定位距离再细分截面并得到截面有效高度
def get_heightlst_fromcad(x_lst, coord_h_lst, ribi_batchnum, rib_xlst, rib_Bailey_y_dict, D):

    # 一根小肋：x_lst_1, x_lst_2, heightlsts
    # 一组小肋：x_lst_1_lst, x_lst_2_lst, heightlsts_lst
    x_lst_1_lst = []
    x_lst_2_lst = []
    heightlsts_lst = []
    
    # 根据小肋组序号 ribi_batchnum 从 rib_xlst 中获得对应小肋的x坐标
    if len(ribi_batchnum)<4:
        rib_batch_x = [rib_xlst[int(ribi_batchnum)-1]]#只有单根小肋时，小肋的x坐标
    else:
        ribnumbers_str = ribi_batchnum.split('to')#有两根及以上的小肋时，小肋的x坐标
        ribnum1 = int(ribnumbers_str[0])-1
        ribnum2 = int(ribnumbers_str[1])
        rib_batch_x = rib_xlst[ribnum1:ribnum2]#midas中一组小肋的x坐标，增加了整体定位的距离。指定了小肋的初始位置
    # print("一组小肋的x坐标为：", rib_batch_x)

    # 根据贝雷梁y坐标和定位距离将vertexxcoord中的CAD坐标转化为Midas坐标
    #将截面x坐标转化成Midas中的y坐标
    x_lst_0 = [round(x - x_lst[0] + D, 1) for x in x_lst] 
    
    for rib_x in rib_batch_x:

        # rib_x 为一根小肋的x坐标，他的y坐标为 rib_Bailey_y_dict[f"{rib_x}"]
        rib_y_lst = rib_Bailey_y_dict[f"{rib_x}"]

        # 先检查贝雷梁y坐标和截面y坐标有没有相近的点，如果有相近的点，将该点截面y坐标替换为贝雷梁的y坐标，近似值范围取1
        for i in range(len(x_lst_0)):
            for j in range(len(rib_y_lst)):
                if abs(x_lst_0[i] - rib_y_lst[j]) < 1:
                    x_lst_0[i] = rib_y_lst[j]

        # x_lst_1为小肋y坐标除去悬臂的部分，包含了截面对应的y坐标和贝雷梁对应的y坐标
        x_lst_1 = list(sorted (set([round(x, 1) for x in x_lst_0] + rib_y_lst)))
        # x_lst_2为截面加截面下贝雷梁的坐标，因为有一部分贝雷梁可能在截面外，所以这部分要去除
        # x_lst_2 = [x for x in x_lst_1 if x-D>=0 and x - x_lst_0[-1] <= 0]
        x_lst_1_lst.append(x_lst_1)
        # x_lst_2_lst.append(x_lst_2)

        # 求贝雷y坐标的线性内插截面高度，返回插好的y坐标表和截面高度表
        # print(rib_y_lst)
        # print(x_lst_0)
        # print(coord_h_lst)
        reslst = Linear_interpolation_with_two_points(rib_y_lst, x_lst_0, coord_h_lst)
        if reslst != None:
            x_lst_2_lst.append(reslst[0])
            heightlsts_lst.append(reslst[1])
        # print(reslst[0])
        # print(reslst[1])
    
    # print(heightlsts_lst)

    # 返回小肋除去悬臂部分在Midas中的y坐标，截面及截面下贝雷梁的y坐标，截面有效高度
    return x_lst_1_lst, x_lst_2_lst, heightlsts_lst



# 生成小肋============================================================================================================================================================
# 生成小肋============================================================================================================================================================
# 生成小肋============================================================================================================================================================

#小肋的所有截面信息
def rib_sections(ribsteel1_filepath, ribsteel2_filepath):
    rib_namelist1=[]
    rib_section1=[]
    rib_namelist2=[]
    rib_section2=[]
    #ribsteel1小肋单截面sec文件，ribsteel2小肋双截面sec文件
    with open(ribsteel1_filepath, 'r', encoding='utf-8') as file1:
        # 逐行读取并打印
        for line in file1:
            original_str1 = line.strip()
            original_parts1 = original_str1.split()
            target_format_parts1=["DBUSER"]
            #first_str1：型钢名字如I10
            first_str1 = original_parts1[0]
            rib_namelist1.append(first_str1)
            target_format_parts1.extend([first_str1])
            if first_str1[0] == 'I':
                target_format_parts1.extend(["CC", "0", "0", "0", "0", "0", "0", "YES", "NO", "H", "2"])
                target_format_parts1.extend(original_parts1[4:8])
                target_format_parts1.extend(["0", "0"])
                target_format_parts1.extend(original_parts1[-2:])
                target_format_parts1.extend(["0", "0"])
            else:
                target_format_parts1.extend(["CC", "0", "0", "0", "0", "0", "0", "YES", "NO", "C", "2"])
                target_format_parts1.extend(original_parts1[4:8])
                target_format_parts1.extend([original_parts1[5],original_parts1[7]])
                target_format_parts1.extend(original_parts1[-2:])
                target_format_parts1.extend(["0", "0"])
            string1 = ",".join(target_format_parts1)
            rib_section1.append(string1)
    with open(ribsteel2_filepath, 'r', encoding='utf-8') as file2:
        # 逐行读取并打印
        for line in file2:
            original_str2 = line.strip()
            original_parts2 = original_str2.split()
            target_format_parts2=["DBUSER"]
            first_str2 = original_parts2[0]
            rib_namelist2.append(first_str2)
            target_format_parts2.extend([first_str2])
            target_format_parts2.extend(["CC", "0", "0", "0", "0", "0", "0", "YES", "NO", "2C", "2"])
            target_format_parts2.extend(original_parts2[4:8])
            target_format_parts2.extend([original_parts2[5],original_parts2[7]])
            target_format_parts2.extend(original_parts2[-2:])
            target_format_parts2.extend(["0", "0"])
            string2 = ",".join(target_format_parts2)
            rib_section2.append(string2)
    return rib_namelist1, rib_section1, rib_namelist2, rib_section2

#根据用户的输入先得到一些小肋的基础信息，可以减少重复获取数据的步骤
#参数依次为：选择的小肋截面型号如'I10'、单截面还是双截面str、编辑框中的小肋间距
def get_rib_basedata(selected_rib_section, section_type, rib9, rib19, rib_seclst):
    if section_type == '单截面' :
        rib_info = list(filter(lambda s: selected_rib_section in s, rib_seclst[1]))#根据小肋截面的型号选择获得对应小肋的mct信息
    else:
        rib_info = list(filter(lambda s: selected_rib_section in s, rib_seclst[3]))#根据小肋截面的型号选择获得对应小肋的mct信息

    #得到小肋的x坐标信息
    rib_spacing_distance_list = midasdisttolst(rib9)#小肋间距表
    rib_xlist = midasdisttolst2(float(rib19), rib9)#小肋所有x坐标
    rib_xlist = [round(x, 1) for x in rib_xlist]
    # print(rib_xlist)
    ribload_distance = ribdistancelist_trans(rib_spacing_distance_list)#小肋承担混凝土荷载宽度表

    return rib_info, rib_spacing_distance_list, rib_xlist, ribload_distance

#创建小肋单元需要小肋的节点坐标，函数传递的参数应包括：
#rib_info是mct格式
def make_rib_nodes_elements_loads(ribi_xlist, ribiload_distance, ribi_batchnum, ribi_ylist1, ribi_ylist2, ribi_hlist, rib_Bailey_y_dict, ribi_cantilever, g_concrete, g_construction, g_template, nodei, elementi):

    #先处理小肋的组号并根据组号得到小肋的x坐标和小肋的荷载宽度
    if len(ribi_batchnum)<4:
        rib_batch_x = [ribi_xlist[int(ribi_batchnum)-1]]#只有单根小肋时，小肋的x坐标
        rib_batch_load = [ribiload_distance[int(ribi_batchnum)-1]]#只有单根小肋时，小肋的荷载宽度
    else:
        ribnumbers_str = ribi_batchnum.split('to')#有两根及以上的小肋时，小肋的x坐标
        ribnum1 = int(ribnumbers_str[0])-1
        ribnum2 = int(ribnumbers_str[1])
        rib_batch_x = ribi_xlist[ribnum1:ribnum2]#midas中一组小肋的x坐标，增加了整体定位的距离。指定了小肋的初始位置
        rib_batch_load = ribiload_distance[ribnum1:ribnum2]#有两根及以上的小肋时，小肋的荷载宽度

    # print("一组小肋的x坐标：", rib_batch_x)
    # print("一组小肋截面的y坐标和贝雷的y坐标：", ribi_ylist1)
    # print("一组小肋截面的y坐标和截面下贝雷的y坐标：", ribi_ylist2)
    # print("ribi_ylist2对应的截面高度：", ribi_hlist)

    #再处理y坐标，ribi_ylist1包含了一组小肋的除悬臂外的y坐标
    #加上小肋悬臂的长度
    rib_batch_y = ribi_ylist1.copy()
    
    if float(ribi_cantilever) != 0:
        [ylst.insert(0, ylst[0] - float(ribi_cantilever)) for ylst in rib_batch_y]
        [ylst.append(ylst[-1] + float(ribi_cantilever)) for ylst in rib_batch_y]
    # print("ribi_ylist1加上悬臂长度等于全部的小肋y坐标：", rib_batch_y)

    #生成小肋节点
    nodes = []
    node_elink_lst = []
    startnodei = int(nodei)#用于单元的起始节点号指定
    for i in range(len(rib_batch_x)):
        for y in rib_batch_y[i]:
            nodestring = f"{nodei} {","} {rib_batch_x[i]} {","} {y} {",0"}"#z小肋z坐标默认为0
            nodes.append(nodestring)
            # 如果y坐标和贝雷有关系，先构建一张初级的弹连筛选表
            if y in rib_Bailey_y_dict[f"{rib_batch_x[i]}"]:
                node_elink_lst.append([int(nodei), [rib_batch_x[i], y, 0]])
            nodei = str(int(nodei)+1)
    # print(nodes)
    # print(node_elink_lst)

    #生成小肋单元
    elements = []
    beamloads = []#混凝土荷载
    constructionloads = []#施工荷载
    templateloads = []#模板荷载

    # print(ribi_ylist2)
    # print(rib_batch_x)
    # print(len(ribi_ylist2))
    # print(len(rib_batch_x))

    for x in range(len(rib_batch_x)):#一组小肋的根数
        i = 0#用于控制混凝土荷载的选取
        # 根据ribi_ylist2在rib_batch_y中的范围确定一根小肋中第n个节点至第n个节点之间有混凝土荷载, ribi_ylist2代表截面下含贝雷梁的小肋y坐标,那么混凝土荷载y坐标开始于ribi_ylist2[0],即截面对应的第一个y坐标
        ribi_ylist2_starty = ribi_ylist2[x][0]
        ribi_ylist2_endy = ribi_ylist2[x][-1]
        rib_beamload_startnodeid = rib_batch_y[x].index(ribi_ylist2_starty)
        rib_beamload_endnodeid = rib_batch_y[x].index(ribi_ylist2_endy)
        rib_beamload_starteleid = rib_beamload_startnodeid
        rib_beamload_endeleid = rib_beamload_endnodeid-1
        # print("截面荷载起始节点号：",rib_beamload_startnodeid, "截面荷载结束单元号：",rib_beamload_endnodeid)
        # print("截面荷载起始单元号：",rib_beamload_starteleid, "截面荷载结束单元号：",rib_beamload_endeleid)

        #根据小肋荷载宽度和混凝土截面高度处理混凝土荷载
        rib_batch_surfaceload = list(map(lambda x: x * float(g_concrete) * 0.000001, rib_batch_load))#面荷载
        ribi_hlist_fix = trim0(ribi_hlist[x],0)#去掉首尾0
        # print("去掉首尾0后的截面高度表：",ribi_hlist_fix)
        # print("去掉首尾0后的截面高度表的个数：",len(ribi_hlist_fix))

        for y in range(len(rib_batch_y[x])-1):#一根小肋的单元数
            #单元
            elementstring = f"{elementi} {",BEAM,1,1,"} {str(startnodei)} {","} {str(startnodei+1)} {",0,0"}"
            if x == len(rib_batch_x) - 1 and y == len(rib_batch_y[x]) - 2:
                last_rib_elementid = int(elementi)
            elif x == 0 and y == 0:
                first_rib_elementid = int(elementi)
            elements.append(elementstring)
            #混凝土荷载
            if y >= rib_beamload_starteleid and y <= rib_beamload_endeleid:
                #i端和j端的混凝土荷载
                beamload0 = -1*ribi_hlist_fix[2*i]*rib_batch_surfaceload[x]
                beamload1 = -1*ribi_hlist_fix[2*i+1]*rib_batch_surfaceload[x]
                # print(y, ribi_hlist_fix[2*i], ribi_hlist_fix[2*i+1])
                if beamload0 != 0 and beamload1 != 0:
                    beamloadstring = f"{elementi} {",LINE,UNILOAD,GZ,NO,NO,aDir[1], , , ,0,"} {beamload0} {",1,"} {beamload1} {",0,0,0,0,混凝土荷载,NO,0,0,NO,"}"
                    beamloads.append(beamloadstring)
                i = i+1
            #施工荷载
            constructionload0 = float(g_construction)*rib_batch_load[x]*-0.001
            constructionloadstring = f"{elementi} {",LINE,UNILOAD,GZ,NO,NO,aDir[1], , , ,0,"} {constructionload0} {",1,"} {constructionload0} {",0,0,0,0,施工荷载,NO,0,0,NO,"}"
            constructionloads.append(constructionloadstring)
            #模板荷载
            templateload0 = float(g_template)*rib_batch_load[x]*-0.001
            templateloadstring = f"{elementi} {",LINE,UNILOAD,GZ,NO,NO,aDir[1], , , ,0,"} {templateload0} {",1,"} {templateload0} {",0,0,0,0,模板荷载,NO,0,0,NO,"}"
            templateloads.append(templateloadstring)
            #单元号增加
            startnodei = startnodei+1
            elementi = str(int(elementi)+1)
        startnodei = startnodei+1
    #print(elements)
    # 小肋结构组
    gruop_lst = [first_rib_elementid, last_rib_elementid]
    return nodei, elementi, nodes, elements, beamloads, constructionloads, templateloads, gruop_lst, node_elink_lst


# 通过读取小肋布置和小肋的截面进行小肋块的绘制，同时生成编号
def draw_rib_blocks_incad(acadapp, rib_name, rib_entry):
    try:
        # 获取活动文档和模型空间
        doc = acadapp.ActiveDocument
        msp = doc.ModelSpace

        # 输入指定坐标
        point = doc.Utility.Getpoint()
        # 用户输入小肋间距转化为小肋x坐标
        rib_xlst = midasdisttolst2(0, rib_entry)
        # 根据插入点转化成cad中的x坐标
        cad_plst = [[point[0]+x, point[1], point[2]] for x in rib_xlst]
        # 小肋序号初始值以及编号文字的字高
        xl_num = 1
        text_height = 100
        # 根据每个点的坐标绘制小肋和序号
        for point in cad_plst:
            # 调用ActiveX方法时坐标的数据类型为变体
            txt_point = [point[0]+100,point[1],point[2]]
            insert_point = vtpnt(point)
            txt_insert_point = vtpnt(txt_point)
            # 绘制小肋
            insert_block_incad_with_type_name(msp, insert_point, "小肋", rib_name)
            # 插入编号
            insert_TXT_incad(msp, txt_insert_point, xl_num, text_height)
            xl_num = xl_num + 1 
        print("小肋绘制完成。")

    except Exception as e:
        # print(e)
        try:
            # 获取程序的报错信息
            excepinfo = e.excepinfo
            scode, source, desc, helpfile, helpcontext, scode2 = excepinfo
            # print(desc)
            # 进行条件判断
            if desc == "文件处理器错误" :
                print("错误！可能的原因：未检索到指定小肋图块，无法生成图形。")
            elif desc == "AutoCAD 主窗口不可见":
                print("错误！可能的原因：AutoCAD主窗口不可见，可能是中途切换当前CAD窗口或CAD关闭导致，请重新尝试。")
            elif desc == None:
                print("错误！可能的原因：用户取消操作或其他。")
            else:
                print("错误！可能的原因：未知。")
        except:
            print("错误！可能的原因：未知。")

# 生成贝雷============================================================================================================================================================
# 生成贝雷============================================================================================================================================================
# 生成贝雷============================================================================================================================================================

# 获取平面视图上贝雷梁图块的插入点
def get_Bailey_Pingmian_InsertionPoints(acadapp, Bailey_select, User_Orignal_Point_inCAD):
    # # 获取当前Cad对应的活动文档、模型空间、贝雷梁平面图选择集
    doc = acadapp.ActiveDocument
    msp = doc.ModelSpace

    # 如果选择集中的子单元是图块并且图块名是指定图块名，返回一个包含图块名和插入点信息的元组
    insertion_points = []
    for obj in Bailey_select:
        insertion_points.append((obj.Name, obj.InsertionPoint, obj.Rotation, obj.XScaleFactor))
    # print(insertion_points)
    # 得到所有的x和所有的y，将最小的x和最小的y作为坐标原点
    insertion_O_points_x = sorted([Dict_Tuple[1][0] for Dict_Tuple in insertion_points])[0]
    insertion_O_points_y = sorted([Dict_Tuple[1][1] for Dict_Tuple in insertion_points])[0]

    # []则说明用户未选取，则生成默认点并生成直线
    if User_Orignal_Point_inCAD == []:
        insertion_O_points = [insertion_O_points_x, insertion_O_points_y]
        insertion_O_points_vtpnt = [insertion_O_points_x, insertion_O_points_y, 0]
        msp.AddXLine(vtpnt(insertion_O_points_vtpnt), vtpnt([insertion_O_points_vtpnt[0] + 1, insertion_O_points_vtpnt[1], insertion_O_points_vtpnt[2]]))
        msp.AddXLine(vtpnt(insertion_O_points_vtpnt), vtpnt([insertion_O_points_vtpnt[0], insertion_O_points_vtpnt[1] + 1, insertion_O_points_vtpnt[2]]))
    # 不为[]则说明用户选取了一个坐标原点
    elif User_Orignal_Point_inCAD != []:
        insertion_O_points = [User_Orignal_Point_inCAD[0], User_Orignal_Point_inCAD[1]]
    # print(insertion_O_points)

    # 处理数据，把z坐标归0，x和y坐标只保留一位小数并且还原到原点
    # insertion_points = [(name, ((round(points[0]-insertion_O_points[0], 1)), (round(points[1]-insertion_O_points[1], 1)), 0.0)) for name, points in insertion_points]
    insertion_points_lst = []
    for name, points, rotation, XScaleFactor in insertion_points:
        point_coordinate_x = (round(points[0]-insertion_O_points[0], 1))
        point_coordinate_y = (round(points[1]-insertion_O_points[1], 1))
        # 把-0.0去掉
        if point_coordinate_y == 0.0:
            insertion_points_lst.append((name, (point_coordinate_x, 0.0, 0.0), rotation, XScaleFactor))
        else:
            insertion_points_lst.append((name, (point_coordinate_x, point_coordinate_y, 0.0), rotation, XScaleFactor))
    # print(insertion_points_lst)

    # 根据插入点、旋转角度以及贝雷梁的长度，将一片贝雷梁的起点和终点坐标计算出来
    insertion_points_lst2 = []
    for Name, InsertionPoint, Rotation, XScaleFactor  in insertion_points_lst:
        if '3m' in Name:
            Bailey_L = 3000
        elif '1.5m' in Name:
            Bailey_L = 1500
        # if abs(rotation - 3.14159) < 0.001:
        #     rotation = 0 # 在cad中不管绕x轴镜像还是y轴镜像, 块的角度均会旋转180°
        StartPoint = [(round(x,1), round(y,1), abs(round(z,1))) for x, y, z in [InsertionPoint]]
        EndPoint = [(round(x+Bailey_L*XScaleFactor*math.cos(Rotation),1), round(y+Bailey_L*XScaleFactor*math.sin(Rotation),1), abs(round(z,1))) for x, y, z in [InsertionPoint]]
        # 重排列, 使x坐标小的为起点
        StartPoint, EndPoint = sorted([StartPoint, EndPoint], key=lambda point: point[0])
        insertion_points_lst2.append((Name, StartPoint[0], rotation, XScaleFactor, EndPoint[0]))
    # print(insertion_points_lst2)

    # 建立所有贝雷梁的起点表
    Bailey_start_point_lst = [(start_x, start_y, start_z) for name, (start_x, start_y, start_z), rotation, XScaleFactor, (end_x, end_y, end_z) in insertion_points_lst2]
    # 建立所有贝雷梁的终点表
    Bailey_end_point_lst = [(end_x, end_y, end_z) for name, (start_x, start_y, start_z), rotation, XScaleFactor, (end_x, end_y, end_z) in insertion_points_lst2]
    # print(Bailey_end_point_lst)
    # 找到每一段贝雷梁的起点，有几个起点说明有几条线段，每一条线段中有若干贝雷梁
    insertion_startpoints_lst = insertion_points_lst2.copy()
    for name, (start_x, start_y, start_z), rotation, XScaleFactor, (end_x, end_y, end_z) in insertion_points_lst2:
        for (x, y, z) in Bailey_end_point_lst:
            # 如果一个起始点的坐标可以在终点表中找到两点坐标接近的一点，则说明该点不能作为一段贝雷梁的起点坐标
            if abs(start_x - x) < 1 and abs(start_y - y) < 1:
                # print((name, (start_x, start_y, start_z), rotation, XScaleFactor, (end_x, end_y, end_z)))
                insertion_startpoints_lst.remove((name, (start_x, start_y, start_z), rotation, XScaleFactor, (end_x, end_y, end_z)))

    # 将所有起点按照先y后x进行排序
    insertion_startpoints_lst = sorted(insertion_startpoints_lst, key = lambda x: (x[1][1], x[1][0]))
    # print(insertion_startpoints_lst)
    # 根据一段贝雷梁的起点，找到一段贝雷梁中的剩余贝雷梁的起点
    coordinate_groups_dict = {}
    # 获取所有 y 坐标，排序次序为从小到大
    y_coordinates = sorted(list([start_y for name, (start_x, start_y, start_z), rotation, XScaleFactor, (end_x, end_y, end_z) in insertion_startpoints_lst]))
    # 将所有 y 坐标添加到字典中，信息头为 "ylst"
    coordinate_groups_dict["ylst"] = y_coordinates
    # 写入原坐标
    coordinate_groups_dict["O_Point"] = insertion_O_points
    # 根据起点贝雷梁的信息分类
    list1 = insertion_points_lst2.copy()
    for i, (name, pt1, rotation, XScaleFactor, pt2) in enumerate(insertion_startpoints_lst):
        if f"第{i+1}排贝雷梁端点信息表" not in coordinate_groups_dict:
            coordinate_groups_dict[f"第{i+1}排贝雷梁端点信息表"] = [(name, pt1, rotation, XScaleFactor, pt2)]
        # 得到起点对应的所有贝雷梁线段
        ptlst = Bailey_connect_line_with_point_angle(pt2, rotation, insertion_points_lst2)
        if ptlst != []:
            coordinate_groups_dict[f"第{i+1}排贝雷梁端点信息表"] = coordinate_groups_dict[f"第{i+1}排贝雷梁端点信息表"] + ptlst
    # print(coordinate_groups_dict)
    # 贝雷梁分组完成
    return coordinate_groups_dict

# 生成一片贝雷梁的点，参数param——0代表3m贝雷梁，1代表1.5m贝雷梁, i代表第i排贝雷梁, insert_P为插入点坐标, dict为词典
def single_Bailey_points(param, i , start_P, end_P, Dict):
    A = f"第{i}排贝雷上弦杆接头坐标"
    B = f"第{i}排贝雷上弦杆竖杆坐标"
    C = f"第{i}排贝雷上弦杆斜杆坐标"
    if A not in Dict:
        Dict[A] = []
    if B not in Dict:
        Dict[B] = []
    if C not in Dict:
        Dict[C] = []
    
    if param == "0":
        increase_lst = [0,90,795,1500,2205,2910,3000]
        for x in increase_lst:
            # # x增量
            # delta_x = round(x * Xfactor * math.cos(angle), 1)
            # # y增量
            # delta_y = round(x * Xfactor * math.sin(angle), 1)
            if x == 0 or x == 3000:
                # Dict[A].append((round(start_P[0] + delta_x,1), round(start_P[1] + delta_y, 1), start_P[2]))
                Dict[A].append(point_coordinate_on_line(start_P, end_P, x))
            if x == 90 or x == 1500 or x == 2910:
                # Dict[B].append((round(start_P[0] + delta_x,1), round(start_P[1] + delta_y, 1), start_P[2]))
                Dict[B].append(point_coordinate_on_line(start_P, end_P, x))
            if x == 795 or x == 2205:
                # Dict[C].append((round(start_P[0] + delta_x,1), round(start_P[1] + delta_y, 1), start_P[2]))
                Dict[C].append(point_coordinate_on_line(start_P, end_P, x))
        # [point_lst.append([round(insert_P[0]+x,1), insert_P[1], insert_P[2]]) for x in increase_lst]
    elif param == "1":
        increase_lst = [0,90,750,1410,1500]
        for x in increase_lst:
            # # x增量
            # delta_x = round(x * Xfactor * math.cos(angle), 1)
            # # y增量
            # delta_y = round(x * Xfactor * math.sin(angle), 1)
            if x == 0 or x == 1500:
                # Dict[A].append((round(start_P[0] + delta_x,1), round(start_P[1] + delta_y, 1), start_P[2]))
                Dict[A].append(point_coordinate_on_line(start_P, end_P, x))
            if x == 90 or x == 1410:
                # Dict[B].append((round(start_P[0] + delta_x,1), round(start_P[1] + delta_y, 1), start_P[2]))
                Dict[B].append(point_coordinate_on_line(start_P, end_P, x))
            if x == 750:
                # Dict[C].append((round(start_P[0] + delta_x,1), round(start_P[1] + delta_y, 1), start_P[2]))
                Dict[C].append(point_coordinate_on_line(start_P, end_P, x))
        # [point_lst.append([round(insert_P[0]+x,1), insert_P[1], insert_P[2]]) for x in increase_lst]
    return Dict

# 根据块名和插入点生成所有贝雷梁弦杆原生的点
def make_Bailey_origin_points(acadapp, Bailey_select, User_Orignal_Point_inCAD):

    # 插入点词典和以贝雷定位的图纸原点
    insert_points_dict = get_Bailey_Pingmian_InsertionPoints(acadapp, Bailey_select, User_Orignal_Point_inCAD)
    O_Point = [round(coor,1) for coor in insert_points_dict["O_Point"]]
    # print(O_Point)
    # 所有贝雷梁的y坐标
    Bailey_ylst = insert_points_dict["ylst"]

    # 所有贝雷坐标点的词典
    Bailey_origin_points_dict = {}

    for i in range(len(Bailey_ylst)):
        # if Bailey_ylst[i] == 0:
        for lst in insert_points_dict[f"第{i+1}排贝雷梁端点信息表"]:
            if lst[0] == '3m贝雷平面block':
                Bailey_origin_points_dict = single_Bailey_points("0", i+1, lst[1], lst[4], Bailey_origin_points_dict)
            elif lst[0] == '1.5m贝雷平面block':
                Bailey_origin_points_dict = single_Bailey_points("1", i+1, lst[1], lst[4], Bailey_origin_points_dict)
    
    # 对每一张表去重并排序
    Bailey_origin_points_dict = {group_name: sorted(remove_nearby_points(list(set(coordinates)))) for group_name, coordinates in Bailey_origin_points_dict.items()}

    Bailey_origin_points_dict["图纸坐标原点"] = O_Point
    Bailey_origin_points_dict["所有y坐标"] = Bailey_ylst
    # print(Bailey_ylst)
    
    # print(Bailey_origin_points_dict)
    return Bailey_origin_points_dict


# 根据小肋的布置参数，将小肋与贝雷的点放入词典的新的头中，命名为弹连点，entry_txt是小肋布置的参数，start_x是小肋的起点横坐标
def Cross_Bailey_Rib(rib9, rib19, Bailey_CAD_to_lst):

    # 包含上弦杆的接头、竖杆、弦杆的点坐标的词典
    Bailey_Dict = Bailey_CAD_to_lst.copy()
    # #最左侧贝雷梁的x坐标
    # left_Bailey_x = min(value[0][0] for key, value in Bailey_Dict.items() if "上弦杆接头" in key)
    
    # 得到小肋的x坐标
    rib_xlst = midasdisttolst2(float(rib19), rib9)

    # 对于每一个小肋的x坐标，与第i排贝雷的首尾接头进行判断，如果坐标包含在接头首尾坐标范围内，则将该点放入第i排贝雷的弹连组中
    # 这一步可以先把端部空缺的贝雷梁排除，得到一个初始的弹连词典
    for i in range(len(Bailey_Dict["所有y坐标"])):
        Bailey_X_range = [Bailey_Dict[f"第{i+1}排贝雷上弦杆接头坐标"][0][0], Bailey_Dict[f"第{i+1}排贝雷上弦杆接头坐标"][-1][0]]
        if f"第{i+1}排贝雷与小肋弹连坐标" not in Bailey_Dict:
            Bailey_Dict[f"第{i+1}排贝雷与小肋弹连坐标"] = []
        for x in rib_xlst:
            if x - Bailey_X_range[0] >= 0 and x - Bailey_X_range[1] <= 0:
                line1 = [Bailey_Dict[f"第{i+1}排贝雷上弦杆接头坐标"][0], Bailey_Dict[f"第{i+1}排贝雷上弦杆接头坐标"][-1]]
                line2 = [[x, -1e10, 0], [x, 1e10, 0]]
                Bailey_Dict[f"第{i+1}排贝雷与小肋弹连坐标"].append(intersection_point_of_two_line_segment(line1, line2))
    # [print(key, ": ", value) for key, value in Bailey_Dict.items() if '贝雷与小肋弹连坐标' in key]
            
    # 返回带弹连信息的贝雷梁词典
    # print(Bailey_Dict)
    return Bailey_Dict

# 判断空洞贝雷梁
def Empty_Bailey(Bailey_Dict):
    # 分析贝雷梁内部的空洞位置
    coor_dict = {}
    for dict_i in range(len(Bailey_Dict["所有y坐标"])):
        # 先得到贝雷梁的信息，根据竖杆是否在接头坐标内判断一排贝雷梁是否连续
        a = Bailey_Dict[f"第{dict_i+1}排贝雷上弦杆接头坐标"]
        b = Bailey_Dict[f"第{dict_i+1}排贝雷上弦杆竖杆坐标"]
        coor_lst = []
        for coor_i in range(len(a)-1):
            if if_ptlst_have_pt_in_range(b, a[coor_i], a[coor_i + 1]) != []:
                coor_lst.append([a[coor_i], a[coor_i + 1]])
            coor_i = coor_i +1
        coor_dict[f"第{dict_i+1}排贝雷单片贝雷坐标表"] = coor_lst
    # print(coor_dict)

    # 第i排贝雷中的空洞贝雷
    empty_Bailey_dict = {}
    for dict_i in range(len(coor_dict)):
        coor_lst = coor_dict[f"第{dict_i+1}排贝雷单片贝雷坐标表"]
        # print(coor_lst)
        if_continuity_res = if_continuity(coor_lst)
        if if_continuity_res != None:
            empty_Bailey_dict[f"第{dict_i+1}排空洞贝雷梁两端点坐标"] = if_continuity_res
        elif if_continuity_res == None:
            empty_Bailey_dict[f"第{dict_i+1}排空洞贝雷梁两端点坐标"] = []
    # print(empty_Bailey_dict)

    return empty_Bailey_dict

# 生成3m和1.5m贝雷梁的三视图图块
def draw_Bailey_incad(acadapp):
    try:
        # 获取活动文档和模型空间
        doc = acadapp.ActiveDocument
        msp = doc.ModelSpace
        # 输入插入3m贝雷梁正立面图块的指定坐标
        point1 = doc.Utility.Getpoint()
        insert_point1 = vtpnt(point1)
        # 3m贝雷梁侧立面、平面的插入点坐标
        point2 = [point1[0] + 4000, point1[1], point1[2]]
        insert_point2 = vtpnt(point2)
        point3 = [point1[0], point1[1] - 2000, point1[2]]
        insert_point3 = vtpnt(point3)
        # 3m贝雷梁正立面、侧立面、平面的插入点坐标
        point4 = [point3[0], point3[1] - 600, point3[2]]
        insert_point4 = vtpnt(point4)
        point5 = [point4[0] + 4000, point4[1], point4[2]]
        insert_point5 = vtpnt(point5)
        point6 = [point4[0], point4[1] - 2000, point4[2]]
        insert_point6 = vtpnt(point6)

        insert_point_lst = [["贝雷正立面" , "3", insert_point1],
                            ["贝雷侧立面" , "3", insert_point2],
                            ["贝雷平面" , "3", insert_point3],
                            ["贝雷正立面" , "1.5", insert_point4],
                            ["贝雷侧立面" , "1.5", insert_point5],
                            ["贝雷平面" , "1.5", insert_point6],
                            ]
        for str1, str2, insertpoint in insert_point_lst:
            insert_block_incad_with_type_name(msp, insertpoint, str1, str2)
        print("贝雷片套组已成功载入。")
    except Exception as e:
        # print(e)
        try:
            # 获取程序的报错信息
            excepinfo = e.excepinfo
            scode, source, desc, helpfile, helpcontext, scode2 = excepinfo
            # print(desc)
            # 进行条件判断
            if desc == "文件处理器错误" :
                print("错误！可能的原因：未检索到指定贝雷片图块，无法生成图形。")
            elif desc == "AutoCAD 主窗口不可见":
                print("错误！可能的原因：AutoCAD主窗口不可见，可能是中途切换当前CAD窗口或CAD关闭导致，请重新尝试。")
            elif desc == None:
                print("错误！可能的原因：用户取消操作或其他。")
            else:
                print("错误！可能的原因：未知。")
        except:
            print("错误！可能的原因：未知。")
# acadapp = win32com.client.Dispatch("AutoCAD.Application")
# draw_Bailey_incad(acadapp)


# 建立贝雷梁单元

# 需要的变量：
# 贝雷梁词典
# 分配梁词典
# 节点和单元的初始号码
# 小肋的截面高度

# 需要得到的结果：
# 建立节点信息流
# 建立单元信息流
# 建立弹连信息流
# 建立结构组信息流
# 建立释放梁端约束信息流

def creat_Bailey_node_element_group_brnd(RIB_H, startid1, startid2, Bailey_dict, FPL_dict, rib_elink_lst, Bailey_Fengya):
    # print(Bailey_dict)
    # print(FPL_dict)

    # 根据贝雷梁词典索引字符串 '所有y坐标' 来确定贝雷梁的排数
    Bailey_ylst = Bailey_dict["所有y坐标"]
    startnodeid = int(startid1)
    startelementid = int(startid2)
    # 贝雷梁节点号词典
    Bailey_ShangXianGan_nodeid_dict ={}
    Bailey_XiaXianGan_nodeid_dict ={}
    Bailey_ShuGan_nodeid_dict ={}
    # Bailey_XieGan_nodeid_dict ={}
    nodeid = "初始化节点号"
    # 结构组
    Bailey_XianGan_group_lst = []
    Bailey_ShuGan_group_lst = []
    Bailey_XieGan_group_lst = []

    # 根据 '第i排贝雷上弦杆接头坐标' '第i排贝雷上弦杆竖杆坐标' '第i排贝雷上弦杆斜杆坐标' '第i排贝雷与小肋弹连坐标' 创建上弦杆节点词典
    for i in range(len(Bailey_ylst)):
        if nodeid == "初始化节点号":
            nodeid = startnodeid
        nodelst1 = [(point[0], point[1], round(0-(float(RIB_H)/2)-50,1)) for point in Bailey_dict[f"第{i+1}排贝雷上弦杆接头坐标"]]
        nodelst2 = [(point[0], point[1], round(0-(float(RIB_H)/2)-50,1)) for point in Bailey_dict[f"第{i+1}排贝雷上弦杆竖杆坐标"]]
        nodelst3 = [(point[0], point[1], round(0-(float(RIB_H)/2)-50,1)) for point in Bailey_dict[f"第{i+1}排贝雷上弦杆斜杆坐标"]]
        nodelst4 = [(point[0], point[1], round(0-(float(RIB_H)/2)-50,1)) for point in Bailey_dict[f"第{i+1}排贝雷与小肋弹连坐标"]]
        # 对合并后的列表进行排序
        sorted_list = sorted(set(nodelst1 + nodelst2 + nodelst3 + nodelst4))
        # 为每个坐标分配序号
        indexed_list = [(i + nodeid, coord) for i, coord in enumerate(sorted_list)]
        # 创建一个字典来存储每个坐标及其序号
        index_dict = {coord: index for index, coord in indexed_list}
        # 将带有序号的坐标重新分配到四个列表
        Bailey_ShangXianGan_nodeid_dict[f"第{i+1}排贝雷上弦杆接头节点号"] = [(index_dict[coord], coord) for coord in nodelst1]
        Bailey_ShangXianGan_nodeid_dict[f"第{i+1}排贝雷上弦杆竖杆节点号"] = [(index_dict[coord], coord) for coord in nodelst2]
        Bailey_ShangXianGan_nodeid_dict[f"第{i+1}排贝雷上弦杆斜杆节点号"] = [(index_dict[coord], coord) for coord in nodelst3]
        Bailey_ShangXianGan_nodeid_dict[f"第{i+1}排贝雷上弦杆弹连节点号"] = [(index_dict[coord], coord) for coord in nodelst4]

        nodeid = Bailey_ShangXianGan_nodeid_dict[f"第{i+1}排贝雷上弦杆接头节点号"][-1][0] + 1
    # print(Bailey_ShangXianGan_nodeid_dict)

    # 同理，创建下弦杆、竖杆的节点词典
    # 下弦杆的接头坐标、竖杆斜杆对应的坐标与上弦杆相同，弹连表根据Bailey_dict中的 "第i排贝雷与分配梁的交点" 索引
    for i in range(len(Bailey_ylst)):
        nodelst1 = [(point[0], point[1], round(0-(float(RIB_H)/2)-1450,1)) for point in Bailey_dict[f"第{i+1}排贝雷上弦杆接头坐标"]]
        nodelst2 = [(point[0], point[1], round(0-(float(RIB_H)/2)-1450,1)) for point in Bailey_dict[f"第{i+1}排贝雷上弦杆竖杆坐标"]]
        nodelst3 = [(point[0], point[1], round(0-(float(RIB_H)/2)-1450,1)) for point in Bailey_dict[f"第{i+1}排贝雷上弦杆斜杆坐标"]]
        nodelst4 = [(point[0], point[1], round(0-(float(RIB_H)/2)-1450,1)) for point in Bailey_dict[f"第{i+1}排贝雷与分配梁的交点"]]
        # 对合并后的列表进行排序
        sorted_list = sorted(set(nodelst1 + nodelst2 + nodelst3 + nodelst4))
        # 为每个坐标分配序号
        indexed_list = [(i + nodeid, coord) for i, coord in enumerate(sorted_list)]
        # 创建一个字典来存储每个坐标及其序号
        index_dict = {coord: index for index, coord in indexed_list}
        # 将带有序号的坐标重新分配到四个列表
        Bailey_XiaXianGan_nodeid_dict[f"第{i+1}排贝雷下弦杆接头节点号"] = [(index_dict[coord], coord) for coord in nodelst1]
        Bailey_XiaXianGan_nodeid_dict[f"第{i+1}排贝雷下弦杆竖杆节点号"] = [(index_dict[coord], coord) for coord in nodelst2]
        Bailey_XiaXianGan_nodeid_dict[f"第{i+1}排贝雷下弦杆斜杆节点号"] = [(index_dict[coord], coord) for coord in nodelst3]
        Bailey_XiaXianGan_nodeid_dict[f"第{i+1}排贝雷下弦杆弹连节点号"] = [(index_dict[coord], coord) for coord in nodelst4]

        nodeid = Bailey_XiaXianGan_nodeid_dict[f"第{i+1}排贝雷下弦杆接头节点号"][-1][0] + 1
    # print(Bailey_XiaXianGan_nodeid_dict)

    # 竖杆需要建立中间节点
    for i in range(len(Bailey_ylst)):
        nodelst1 = [(point[0], point[1], round(0-(float(RIB_H)/2)-750,1)) for point in Bailey_dict[f"第{i+1}排贝雷上弦杆竖杆坐标"]]
        # 为每个坐标分配序号
        indexed_list = [(i + nodeid, coord) for i, coord in enumerate(nodelst1)]
        # 创建一个字典来存储每个坐标及其序号
        index_dict = {coord: index for index, coord in indexed_list}
        # 将带有序号的坐标重新分配进列表
        Bailey_ShuGan_nodeid_dict[f"第{i+1}排贝雷竖杆中间节点号"] = [(index_dict[coord], coord) for coord in nodelst1]

        nodeid = Bailey_ShuGan_nodeid_dict[f"第{i+1}排贝雷竖杆中间节点号"][-1][0] + 1
    # print(Bailey_ShuGan_nodeid_dict)

    # 根据以上词典生成贝雷梁节点mct信息流
    Bailey_node_string_lst = []

    # 建立所有贝雷梁节点
    for nodeid_dict in [Bailey_ShangXianGan_nodeid_dict, Bailey_XiaXianGan_nodeid_dict, Bailey_ShuGan_nodeid_dict]:
        str_lst = ["{}, {}, {}, {}".format(nodeid_lst[0], nodeid_lst[1][0], nodeid_lst[1][1], nodeid_lst[1][2]) for key, value in nodeid_dict.items() for nodeid_lst in value]
        Bailey_node_string_lst = Bailey_node_string_lst + str_lst

    Bailey_node_string_lst = list(OrderedDict.fromkeys(Bailey_node_string_lst))
    
    # 开始建立单元

    elementid = "初始化单元号"
    Bailey_ShangXianGan_elementid_dict ={}
    Bailey_ShangXianGan_JT_elementid_dict ={}
    Bailey_XiaXianGan_elementid_dict ={}
    Bailey_XiaXianGan_JT_elementid_dict ={}
    Bailey_ShuGan_elementid_dict ={}
    Bailey_XieGan_elementid_dict ={}

    # 首先创建单元词典，包含 第i排贝雷梁上弦杆接头单元、 上弦杆单元、 竖杆单元、 斜杆单元、 下弦杆接头单元、 下弦杆单元
    # 单元的连接需要考虑空洞贝雷梁的存在，即节点和节点之间可能没有单元

    empty_Bailey_dict = Empty_Bailey(Bailey_dict)
    # print(empty_Bailey_dict)
    # 列表全空返回True，任一不为空返回False
    if_all_empty = all(not value for value in empty_Bailey_dict.values())
    
    dict1 = {}
    for i in range(len(Bailey_ylst)):
        # 贝雷上弦杆节点表整合
        for key, value in Bailey_ShangXianGan_nodeid_dict.items():
            node = int(re.findall(r'\d+', key)[0])
            if i+1 == node:
                if f"第{i+1}排贝雷上弦杆全部节点表" not in dict1:
                    dict1[f"第{i+1}排贝雷上弦杆全部节点表"] = []
                dict1[f"第{i+1}排贝雷上弦杆全部节点表"].append(value)
    for keys, values in dict1.items():
        dict1[keys] = sorted(set(item for itemlst in values for item in itemlst), key=lambda x: x[1][0])


    dict2 = {}
    for i in range(len(Bailey_ylst)):
        # 贝雷上弦杆节点表整合
        for key, value in Bailey_XiaXianGan_nodeid_dict.items():
            node = int(re.findall(r'\d+', key)[0])
            if i+1 == node:
                if f"第{i+1}排贝雷下弦杆全部节点表" not in dict2:
                    dict2[f"第{i+1}排贝雷下弦杆全部节点表"] = []
                dict2[f"第{i+1}排贝雷下弦杆全部节点表"].append(value)
    for keys, values in dict2.items():
        dict2[keys] = sorted(set(item for itemlst in values for item in itemlst), key=lambda x: x[1][0])

    # 根据空洞贝雷梁的端点x坐标以及弦杆节点号表来生成单元
    for i in range(len(Bailey_ylst)):

        # 初始化单元号
        if elementid == "初始化单元号":
            elementid = startelementid
        # 获取第i排贝雷梁空洞贝雷梁的端点坐标
        if if_all_empty:
            empty_xlst = []
            empty_xlst2 = []
        else:
            empty_xlst = [[ptlst[0][0], ptlst[1][0]] for ptlst in empty_Bailey_dict[f"第{i+1}排空洞贝雷梁两端点坐标"]]
            empty_xlst2 = [item for sublist in empty_xlst for item in sublist]
        # 获取第i排贝雷梁接头的坐标
        JieTou_xlst = [pt[0] for pt in Bailey_dict[f"第{i+1}排贝雷上弦杆接头坐标"]]

        # 上弦杆
        for j in range(len(dict1[f"第{i+1}排贝雷上弦杆全部节点表"])-1):
            # 初始化
            if f"第{i+1}排贝雷上弦杆单元号" not in Bailey_ShangXianGan_elementid_dict:
                Bailey_ShangXianGan_elementid_dict[f"第{i+1}排贝雷上弦杆单元号"] = []
            # 获取节点号和节点x坐标
            node_lst1 = dict1[f"第{i+1}排贝雷上弦杆全部节点表"][j]
            node_lst2 = dict1[f"第{i+1}排贝雷上弦杆全部节点表"][j+1]
            node_i1 = node_lst1[0]
            node_x1 = node_lst1[1][0]
            node_i2 = node_lst2[0]
            node_x2 = node_lst2[1][0]
            # if i == 0:
            #     print([elementid, node_i1, node_i2])
            # 当满足两个节点之间存在单元时，返回值为True, 建立上弦杆单元，单元号+1
            if is_element_between_nodes(node_x1, node_x2, empty_xlst):
                Bailey_ShangXianGan_elementid_dict[f"第{i+1}排贝雷上弦杆单元号"].append([elementid, node_i1, node_i2])
                # 记录第i排贝雷梁上弦杆单元号
                if j == 0:
                    Bailey_ShangXianGan_first_elementid = elementid
                elif j == len(dict1[f"第{i+1}排贝雷上弦杆全部节点表"])-2:
                    Bailey_ShangXianGan_last_elementid = elementid
                # 建立单元时，若满足当前单元为接头，且是中间贝雷梁时，记录接头单元号
                if (node_x1 in JieTou_xlst) and (node_x1 not in empty_xlst2) and (j != 0):
                    if f"第{i+1}排贝雷上弦杆前接头单元号" not in Bailey_ShangXianGan_JT_elementid_dict:
                        Bailey_ShangXianGan_JT_elementid_dict[f"第{i+1}排贝雷上弦杆前接头单元号"] = []
                    Bailey_ShangXianGan_JT_elementid_dict[f"第{i+1}排贝雷上弦杆前接头单元号"].append(elementid)
                elif (node_x2 in JieTou_xlst) and (node_x2 not in empty_xlst2) and (j != len(dict1[f"第{i+1}排贝雷上弦杆全部节点表"])-2):
                    if f"第{i+1}排贝雷上弦杆后接头单元号" not in Bailey_ShangXianGan_JT_elementid_dict:
                        Bailey_ShangXianGan_JT_elementid_dict[f"第{i+1}排贝雷上弦杆后接头单元号"] = []
                    Bailey_ShangXianGan_JT_elementid_dict[f"第{i+1}排贝雷上弦杆后接头单元号"].append(elementid)
                elementid = elementid + 1
        # 弦杆结构组
        Bailey_XianGan_group_lst.append([Bailey_ShangXianGan_first_elementid, Bailey_ShangXianGan_last_elementid])

        # 下弦杆
        for j in range(len(dict2[f"第{i+1}排贝雷下弦杆全部节点表"])-1):
            # 初始化
            if f"第{i+1}排贝雷下弦杆单元号" not in Bailey_XiaXianGan_elementid_dict:
                Bailey_XiaXianGan_elementid_dict[f"第{i+1}排贝雷下弦杆单元号"] = []
            # 获取节点号和节点x坐标
            node_lst1 = dict2[f"第{i+1}排贝雷下弦杆全部节点表"][j]
            node_lst2 = dict2[f"第{i+1}排贝雷下弦杆全部节点表"][j+1]
            node_i1 = node_lst1[0]
            node_x1 = node_lst1[1][0]
            node_i2 = node_lst2[0]
            node_x2 = node_lst2[1][0]
            # 当满足两个节点之间存在单元时，返回值为True, 建立上弦杆单元，单元号+1
            if is_element_between_nodes(node_x1, node_x2, empty_xlst):
                Bailey_XiaXianGan_elementid_dict[f"第{i+1}排贝雷下弦杆单元号"].append([elementid, node_i1, node_i2])
                # 记录第i排贝雷梁下弦杆单元号
                if j == 0:
                    Bailey_XiaXianGan_first_elementid = elementid
                elif j == len(dict2[f"第{i+1}排贝雷下弦杆全部节点表"])-2:
                    Bailey_XiaXianGan_last_elementid = elementid
                # 建立单元时，若满足当前单元为接头，且是中间贝雷梁时，记录接头单元号
                if (node_x1 in JieTou_xlst) and (node_x1 not in empty_xlst2) and (j != 0):
                    if f"第{i+1}排贝雷下弦杆前接头单元号" not in Bailey_XiaXianGan_JT_elementid_dict:
                        Bailey_XiaXianGan_JT_elementid_dict[f"第{i+1}排贝雷下弦杆前接头单元号"] = []
                    Bailey_XiaXianGan_JT_elementid_dict[f"第{i+1}排贝雷下弦杆前接头单元号"].append(elementid)
                elif (node_x2 in JieTou_xlst) and (node_x2 not in empty_xlst2) and (j != len(dict2[f"第{i+1}排贝雷下弦杆全部节点表"])-2):
                    if f"第{i+1}排贝雷下弦杆后接头单元号" not in Bailey_XiaXianGan_JT_elementid_dict:
                        Bailey_XiaXianGan_JT_elementid_dict[f"第{i+1}排贝雷下弦杆后接头单元号"] = []
                    Bailey_XiaXianGan_JT_elementid_dict[f"第{i+1}排贝雷下弦杆后接头单元号"].append(elementid)
                elementid = elementid + 1
        # 弦杆结构组
        Bailey_XianGan_group_lst.append([Bailey_XiaXianGan_first_elementid, Bailey_XiaXianGan_last_elementid])

        # 竖杆
        # 竖杆之间的节点读取上弦杆竖杆节点、竖杆中间节点、下弦杆竖杆节点
        ShuGan_nodeid_lst1 = [value[0] for value in Bailey_ShangXianGan_nodeid_dict[f"第{i+1}排贝雷上弦杆竖杆节点号"]]
        ShuGan_nodeid_lst2 = [value[0] for value in Bailey_ShuGan_nodeid_dict[f"第{i+1}排贝雷竖杆中间节点号"]]
        ShuGan_nodeid_lst3 = [value[0] for value in Bailey_XiaXianGan_nodeid_dict[f"第{i+1}排贝雷下弦杆竖杆节点号"]]
        if f"第{i+1}排贝雷竖杆单元号" not in Bailey_ShuGan_elementid_dict:
            Bailey_ShuGan_elementid_dict[f"第{i+1}排贝雷竖杆单元号"] = []
        # print(ShuGan_nodeid_lst1)
        # print(ShuGan_nodeid_lst2)
        # print(ShuGan_nodeid_lst3)
        # 两两节点号表之间建立竖杆单元
        for k in range(len(ShuGan_nodeid_lst1)):
            Bailey_ShuGan_elementid_dict[f"第{i+1}排贝雷竖杆单元号"].append([elementid, ShuGan_nodeid_lst1[k], ShuGan_nodeid_lst2[k]])
            if k == 0:
                Bailey_ShuGan_first_elementid = elementid
            elementid = elementid + 1

        for k in range(len(ShuGan_nodeid_lst2)):
            Bailey_ShuGan_elementid_dict[f"第{i+1}排贝雷竖杆单元号"].append([elementid, ShuGan_nodeid_lst2[k], ShuGan_nodeid_lst3[k]])
            if k == len(ShuGan_nodeid_lst2) - 1:
                Bailey_ShuGan_last_elementid = elementid
            elementid = elementid + 1
        # 竖杆结构组
        Bailey_ShuGan_group_lst.append([Bailey_ShuGan_first_elementid, Bailey_ShuGan_last_elementid])

        # print(Bailey_ShuGan_elementid_dict)

        # 斜杆
        # 斜杆之间的节点读取上弦杆斜杆节点、竖杆中间节点、下弦杆斜杆节点
        XieGan_nodeid_lst1 = [value[0] for value in Bailey_ShangXianGan_nodeid_dict[f"第{i+1}排贝雷上弦杆斜杆节点号"]]
        XieGan_nodeid_lst2 = [value for value in Bailey_ShuGan_nodeid_dict[f"第{i+1}排贝雷竖杆中间节点号"]]
        XieGan_nodeid_lst3 = [value[0] for value in Bailey_XiaXianGan_nodeid_dict[f"第{i+1}排贝雷下弦杆斜杆节点号"]]
        if f"第{i+1}排贝雷斜杆单元号" not in Bailey_XieGan_elementid_dict:
            Bailey_XieGan_elementid_dict[f"第{i+1}排贝雷斜杆单元号"] = []
        # 上下弦杆中每个节点对应竖杆中间节点的两个节点
        k0 = 0
        for k in range(len(XieGan_nodeid_lst2) - 1):
            # 如果中间节点两个节点之间间距是1410(3m贝雷片)或1320(1.5m贝雷片)，则分别连上下弦杆的两个节点，否则不连
            # x1 = XieGan_nodeid_lst2[k][1][0]
            # x2 = XieGan_nodeid_lst2[k + 1][1][0]
            pt1 = XieGan_nodeid_lst2[k][1]
            pt2 = XieGan_nodeid_lst2[k + 1][1]
            pt_distance = distance(pt1, pt2)
            if abs(pt_distance - 1410) < 0.1 or abs(pt_distance - 1320) < 0.1:
                # print(abs(x1 - x2))
                # print(XieGan_nodeid_lst1[k0])
                if k == 0:
                    Bailey_XieGan_first_elementid = elementid
                Bailey_XieGan_elementid_dict[f"第{i+1}排贝雷斜杆单元号"].append([elementid, XieGan_nodeid_lst1[k0], XieGan_nodeid_lst2[k][0]])
                elementid = elementid + 1
                Bailey_XieGan_elementid_dict[f"第{i+1}排贝雷斜杆单元号"].append([elementid, XieGan_nodeid_lst1[k0], XieGan_nodeid_lst2[k + 1][0]])
                elementid = elementid + 1
                Bailey_XieGan_elementid_dict[f"第{i+1}排贝雷斜杆单元号"].append([elementid, XieGan_nodeid_lst3[k0], XieGan_nodeid_lst2[k][0]])
                elementid = elementid + 1
                Bailey_XieGan_elementid_dict[f"第{i+1}排贝雷斜杆单元号"].append([elementid, XieGan_nodeid_lst3[k0], XieGan_nodeid_lst2[k + 1][0]])
                if k == len(XieGan_nodeid_lst2) - 2:
                    Bailey_XieGan_last_elementid = elementid
                elementid = elementid + 1
                # print(XieGan_nodeid_lst1[k0])
                # print(XieGan_nodeid_lst2[k], XieGan_nodeid_lst2[k+1])
                # print(XieGan_nodeid_lst3[k0])
                # print(k0)
                k0 = k0 + 1
        # 斜杆结构组
        Bailey_XieGan_group_lst.append([Bailey_XieGan_first_elementid, Bailey_XieGan_last_elementid])

    # print("贝雷梁上弦杆单元号词典：", Bailey_ShangXianGan_elementid_dict, end="\n\n")
    # print("贝雷梁上弦杆接头单元号词典：", Bailey_ShangXianGan_JT_elementid_dict, end="\n\n")
    # print("贝雷梁下弦杆单元号词典：", Bailey_XiaXianGan_elementid_dict, end="\n\n")
    # print("贝雷梁下弦杆接头单元号词典：", Bailey_XiaXianGan_JT_elementid_dict, end="\n\n")
    # print("贝雷梁竖杆单元号词典：", Bailey_ShuGan_elementid_dict, end="\n\n")
    # print("贝雷梁斜杆单元号词典：", Bailey_XieGan_elementid_dict, end="\n\n")

    # 根据各个词典生成mct单元命令流，弦杆材料号/截面号/旋转角度 = 2,2,0  竖杆材料号/截面号/旋转角度 = 3,3,90  斜杆材料号/截面号/旋转角度 = 4,3,90
    Bailey_element_string_lst = []

    for dict_lst in [
        [Bailey_ShangXianGan_elementid_dict, 2,2,0], 
        [Bailey_XiaXianGan_elementid_dict, 2,2,0],
        [Bailey_ShuGan_elementid_dict, 3,3,90],
        [Bailey_XieGan_elementid_dict, 4,3,90],
        ]:
        str_lst = ["{}, BEAM, {}, {}, {}, {}, {}, 0".format(
            elementid_lst[0], # 单元号
            dict_lst[1], # 材料号
            dict_lst[2], # 截面号
            elementid_lst[1],# 节点号i
            elementid_lst[2],# 节点号j
            dict_lst[3],# 旋转角度
            ) 
            for key, value in dict_lst[0].items() for elementid_lst in value]
        Bailey_element_string_lst = Bailey_element_string_lst + str_lst
    
    Bailey_element_string_lst = list(OrderedDict.fromkeys(Bailey_element_string_lst))

    # 将结构组转化为字符串的形式
    group_str1 = ' '.join(f"{start}to{end}" for start, end in Bailey_XianGan_group_lst)
    Bailey_XianGan_group = [f"贝雷梁弦杆, , {group_str1}, 0"]
    group_str2 = ' '.join(f"{start}to{end}" for start, end in Bailey_ShuGan_group_lst)
    Bailey_ShuGan_group = [f"贝雷梁竖杆, , {group_str2}, 0"]
    group_str3 = ' '.join(f"{start}to{end}" for start, end in Bailey_XieGan_group_lst)
    Bailey_XieGan_group = [f"贝雷梁斜杆, , {group_str3}, 0"]

    print("Success: 成功生成贝雷梁节点、单元及结构组。")

    # 释放梁端约束
    FRAME_RLS_element_lst1 = [id for key, value in Bailey_ShangXianGan_JT_elementid_dict.items() if "后接头" in key for id in value]
    FRAME_RLS_element_lst2 = [id for key, value in Bailey_XiaXianGan_JT_elementid_dict.items() if "后接头" in key for id in value]
    FRAME_RLS_lst = ["{}, NO, 000000, 0, 0, 0, 0, 0, 0 \n 000010, 0, 0, 0, 0, 0, 0, 释放梁端约束".format(x) for x in FRAME_RLS_element_lst1 + FRAME_RLS_element_lst2]
    print("Success: 释放梁端约束成功。")

    # 小肋和贝雷的弹性连接
    # print(rib_elink_lst)
    # print(Bailey_ShangXianGan_nodeid_dict)
    reslst = ELASTICLINK_between_rib_Bailey(rib_elink_lst, Bailey_ShangXianGan_nodeid_dict)
    # print(reslst)
    elink_lst = []
    elink_i = 1
    for sublist in reslst:
        point1, point2 = sublist
        elink_lst.append("{}, {}, {}, GEN, 0, 1e+007, 1e+006, 1e+006, 1e+005, 1e+005, 1e+005, NO, 0.5, 0.5, 弹性连接".format(elink_i, point1[0], point2[0]))
        elink_i = elink_i + 1
    print("Success: 贝雷与小肋弹性连接创建成功。")

    # 贝雷片上的风荷载由压强Bailey_Wind和迎风贝雷梁排数n决定，转化为贝雷梁杆件上的线荷载
    wind_ylst = sorted(list(set(Bailey_ylst)))
    wind_ydict = {}
    for i, y in enumerate(Bailey_ylst):
        if str(y) not in wind_ydict:
            wind_ydict[str(y)] = []
        wind_ydict[str(y)].append(i+1)
    # print(wind_ylst)
    # print(wind_ydict)
    # 风压计算值和遮挡系数
    Wind_Beamload = []
    Bailey_Feng_kpa = Bailey_Fengya
    if Bailey_Feng_kpa != 0:
        η = 0.66
        # 弦杆迎风宽度为0.1m，竖杆斜杆迎风宽度为0.05m
        # 计算第i排贝雷梁弦杆、竖杆、斜杆上的线荷载大小
        XianGan_Feng_Line_Load_dict = {}
        ShuXieGan_Feng_Line_Load_dict = {}
        for i in range(len(wind_ylst)):
            # i对应的y坐标及y坐标对应的排数
            for j in wind_ydict[str(wind_ylst[i])]:
                if f"第{j}排迎风贝雷上下弦杆线荷载" not in XianGan_Feng_Line_Load_dict:
                    XianGan_Feng_Line_Load_dict[f"第{j}排迎风贝雷上下弦杆线荷载"] = round(pow(η, i) * 0.1 * Bailey_Feng_kpa, 5)
                if f"第{j}排迎风贝雷竖杆斜杆线荷载" not in ShuXieGan_Feng_Line_Load_dict:
                    ShuXieGan_Feng_Line_Load_dict[f"第{j}排迎风贝雷竖杆斜杆线荷载"] = round(pow(η, i) * 0.05 * Bailey_Feng_kpa, 5)
        # print(XianGan_Feng_Line_Load_dict)
        # print(ShuXieGan_Feng_Line_Load_dict)
        # 以线荷载形式添加弦杆、竖杆、斜杆上的风荷载，当弦杆上线荷载小于等于0.01kN/m、竖杆斜杆上线荷载小于等于0.005kN/m时忽略，即风压为0.1kpa
        # 上弦杆风荷载
        for key, value in Bailey_ShangXianGan_elementid_dict.items():
            if "贝雷上弦杆单元号" in key:
                element_lst = [elelst[0] for elelst in value]
                num = re.findall(r'\d+', key)[0]
                beamload1 = XianGan_Feng_Line_Load_dict[f"第{int(num)}排迎风贝雷上下弦杆线荷载"]
                if beamload1 > 0.01:
                    for elementi in element_lst:
                        beamloadstring = f"{elementi} {",LINE,UNILOAD,GY,NO,NO,aDir[1], , , ,0,"} {beamload1} {",1,"} {beamload1} {",0,0,0,0,,NO,0,0,NO,"}"
                        Wind_Beamload.append(beamloadstring)
        # 下弦杆风荷载
        for key, value in Bailey_XiaXianGan_elementid_dict.items():
            if "贝雷下弦杆单元号" in key:
                element_lst = [elelst[0] for elelst in value]
                num = re.findall(r'\d+', key)[0]
                beamload1 = XianGan_Feng_Line_Load_dict[f"第{int(num)}排迎风贝雷上下弦杆线荷载"]
                if beamload1 > 0.01:
                    for elementi in element_lst:
                        beamloadstring = f"{elementi} {",LINE,UNILOAD,GY,NO,NO,aDir[1], , , ,0,"} {beamload1} {",1,"} {beamload1} {",0,0,0,0,,NO,0,0,NO,"}"
                        Wind_Beamload.append(beamloadstring)
        # 竖杆风荷载
        for key, value in Bailey_ShuGan_elementid_dict.items():
            if "贝雷竖杆单元号" in key:    
                element_lst = [elelst[0] for elelst in value]
                num = re.findall(r'\d+', key)[0]
                beamload2 = ShuXieGan_Feng_Line_Load_dict[f"第{int(num)}排迎风贝雷竖杆斜杆线荷载"]
                if beamload2 > 0.005:
                    for elementi in element_lst:
                        beamloadstring = f"{elementi} {",LINE,UNILOAD,GY,NO,NO,aDir[1], , , ,0,"} {beamload2} {",1,"} {beamload2} {",0,0,0,0,,NO,0,0,NO,"}"
                        Wind_Beamload.append(beamloadstring)
        # 斜杆风荷载
        for key, value in Bailey_XieGan_elementid_dict.items():
            if "贝雷斜杆单元号" in key:
                element_lst = [elelst[0] for elelst in value]
                num = re.findall(r'\d+', key)[0]
                beamload2 = ShuXieGan_Feng_Line_Load_dict[f"第{int(num)}排迎风贝雷竖杆斜杆线荷载"]
                if beamload2 > 0.005:
                    for elementi in element_lst:
                        beamloadstring = f"{elementi} {",LINE,UNILOAD,GY,NO,NO,aDir[1], , , ,0,"} {beamload2} {",1,"} {beamload2} {",0,0,0,0,,NO,0,0,NO,"}"
                        Wind_Beamload.append(beamloadstring)
        print("Success: 贝雷梁上风荷载成功生成。")
    else:
        print("Warning: 贝雷梁上风荷载未生成。")

    return Bailey_node_string_lst, Bailey_element_string_lst, Bailey_XianGan_group, Bailey_ShuGan_group, Bailey_XieGan_group, FRAME_RLS_lst, elink_lst, Bailey_XiaXianGan_nodeid_dict, nodeid, elementid, elink_i, Wind_Beamload


# 对于弹性连接，小肋与贝雷梁有的数据有小肋的节点mct流，贝雷梁词典 Bailey_ShangXianGan_nodeid_dict[f"第{i+1}排贝雷上弦杆弹连节点号"]
def ELASTICLINK_between_rib_Bailey(lst1, dict1):
    ELASTICLINK_lst = [find_rib_elink_nodeid(lst1, ptlst) for key, value in dict1.items() if "弹连" in key for ptlst in value]
    return ELASTICLINK_lst
    

# 在小肋节点mct信息流中找到与指定坐标相关联的节点号和坐标，返回[节点号，坐标]
def find_rib_elink_nodeid(lst1, lst2):
    nodeid2 = lst2[0]
    coord2 = lst2[1]
    for sublist in lst1:
        nodeid1, coord1 = sublist
        if abs(coord1[0] - coord2[0]) < 1 and abs(coord1[1] - coord2[1]) < 1:
            return [sublist, lst2]

# 生成分配梁============================================================================================================================================================
# 生成分配梁============================================================================================================================================================
# 生成分配梁============================================================================================================================================================

# 返回分配梁词典 "第i根分配梁"：[[插入点], 长度]，词典根据先x坐标后y坐标、先小后大的原则排序，用户使用块时，插入点需要注意位置和方向
def Reg_FPL_PM(Fenpeiliang_select, Bailey_dict):
    # 指定图层名称
    layer_name = "Defpoints"
    #分配梁信息词典
    layer1_line_dict = {}
    FPL_i = 1

    # 获取与指定图层相关的直线
    for i in range(len(Fenpeiliang_select)):

        obj = Fenpeiliang_select[i]
    
        # 查询分配梁的扩展属性，对于平面分配梁，其扩展属性为EX_FPL_PM
        # if obj.EntityName == "AcDbBlockReference":
        xdata = get_block_reference_xdata(obj)
        if xdata[1][0] == "SZQS" and "EX_FPL_PM" in xdata[1][1]:
            # obj_name = obj.Name
            # obj_point = [round(coor,1) for coor in obj.InsertionPoint]
            # 读取扩展属性中记录的该根分配梁的截面类型
            FPL_str_name = xdata[1][1].replace("EX_FPL_PM_", "")
            # print(FPL_str_name)
            # 复制块
            # new_obj = obj.Copy()
            # 如果是需要的块，获取该块的“Defpoints”图层的长度(一般是中心线)
            get_lines_from_block_reslst = get_lines_from_block(obj, layer_name)
            FPL_length = get_lines_from_block_reslst[0]
            FPL_Start_point = [round(coor,1) for coor in get_lines_from_block_reslst[1]]
            FPL_End_point = [round(coor,1) for coor in get_lines_from_block_reslst[2]]
            
            if f"第{FPL_i}根分配梁" not in layer1_line_dict:
                if FPL_Start_point[1] < FPL_End_point[1]:
                    layer1_line_dict[f"第{FPL_i}根分配梁"] = [FPL_Start_point, FPL_End_point, round(FPL_length,1), FPL_str_name]
                if FPL_Start_point[1] > FPL_End_point[1]:
                    layer1_line_dict[f"第{FPL_i}根分配梁"] = [FPL_End_point, FPL_Start_point, round(FPL_length,1), FPL_str_name]

            FPL_i = FPL_i + 1

    layer1_line_dict_sorted_values = sorted(layer1_line_dict.values(), key=lambda item: (item[0][0], item[0][1]))

    # 重新构建字典，保持键不变
    layer1_line_dict_sorted = {key: value for key, value in zip(layer1_line_dict.keys(), layer1_line_dict_sorted_values)}

    # 输出排序后的结果
    # print(layer1_line_dict_sorted)
    # 获取坐标原点
    O_Point = Bailey_dict["图纸坐标原点"]+[0]
    # 将分配梁的图纸坐标转换为模型坐标
    for key, value in layer1_line_dict_sorted.items():
        layer1_line_dict_sorted[key][0] = [round(value[0][i] - O_Point[i],1) for i in range(3)]
        layer1_line_dict_sorted[key][1] = [round(value[1][i] - O_Point[i],1) for i in range(3)]

    return layer1_line_dict_sorted


# 读取分配梁的截面数据文件，对应到选择分配梁截面的下拉列表中
def FPL_section(FPL1_filepath, FPL2_filepath):
    FPL_namelist1 = []
    FPL_section1 = []
    FPL_namelist2 = []
    FPL_section2 = []
    # 打开分配梁单截面sec文件
    with open(FPL1_filepath, 'r', encoding='utf-8') as file1:
        # 逐行读取并打印
        for line in file1:
            original_str1 = line.strip()
            original_parts1 = original_str1.split()
            # 设置开头
            target_format_parts1=["DBUSER"]
            # 添加mct文件截面的型钢名，如I10
            first_str1 = original_parts1[0]
            FPL_namelist1.append(first_str1)
            target_format_parts1.extend([first_str1])
            # 如果是单工字钢
            if first_str1[0] == 'I':
                target_format_parts1.extend(["CC", "0", "0", "0", "0", "0", "0", "YES", "NO", "H", "2"])
                target_format_parts1.extend(original_parts1[4:8])
                target_format_parts1.extend(["0", "0"])
                target_format_parts1.extend(original_parts1[-2:])
                target_format_parts1.extend(["0", "0"])
            # 如果是 HM 或者 HN 型钢
            elif first_str1[0] == 'H':
                target_format_parts1.extend(["CC", "0", "0", "0", "0", "0", "0", "YES", "NO", "H", "2"])
                target_format_parts1.extend(original_parts1[2:6])
                target_format_parts1.extend(["0", "0"])
                target_format_parts1.extend([original_parts1[6]])
                target_format_parts1.extend(["0", "0", "0"])
            # 逐条合并字节流
            string1 = ",".join(target_format_parts1)
            FPL_section1.append(string1)
    # 打开分配梁双截面sec文件
    with open(FPL2_filepath, 'r', encoding='utf-8') as file2:
        # 逐行读取并打印
        for line in file2:
            original_str2 = line.strip()
            original_parts2 = original_str2.split()
            # 设置开头
            target_format_parts2=["DBUSER"]
            # 添加mct文件截面的型钢名，如I10
            first_str2 = original_parts2[0]
            FPL_namelist2.append(first_str2)
            target_format_parts2.extend([first_str2])
            target_format_parts2.extend(["CC", "0", "0", "0", "0", "0", "0", "YES", "NO", "B", "2"])
            target_format_parts2.extend(original_parts2[2:8])
            target_format_parts2.extend(["0", "0", "0", "0"])
            # 逐条合并字节流
            string2 = ",".join(target_format_parts2)
            FPL_section2.append(string2)
    return FPL_namelist1, FPL_section1, FPL_namelist2, FPL_section2


# 插入分配梁对应截面的块
def draw_FPL_incad(acadapp, FPL_name_str):
    try:
        # 获取活动文档和模型空间
        doc = acadapp.ActiveDocument
        msp = doc.ModelSpace
        # 输入指定坐标
        point = doc.Utility.Getpoint()
        insert_point = vtpnt(point)
        # 插入指定分配梁对应截面的块
        insert_block_incad_with_type_name(msp, insert_point, "分配梁立面", FPL_name_str)
        # 插入指定分配梁对应平面的块
        insert_block_incad_with_type_name(msp, insert_point, "分配梁平面", FPL_name_str)
    except Exception as e:
        # print(e)
        try:
            # 获取程序的报错信息
            excepinfo = e.excepinfo
            scode, source, desc, helpfile, helpcontext, scode2 = excepinfo
            # print(desc)
            # 进行条件判断
            if desc == "文件处理器错误" :
                print("错误！可能的原因：未检索到指定分配梁图块，无法生成图形。")
            elif desc == "AutoCAD 主窗口不可见":
                print("错误！可能的原因：AutoCAD主窗口不可见，可能是中途切换当前CAD窗口或CAD关闭导致，请重新尝试。")
            elif desc == None:
                print("错误！可能的原因：用户取消操作或其他。")
            else:
                print("错误！可能的原因：未知。")
        except:
            print("错误！可能的原因：未知。")
# acadapp = win32com.client.Dispatch("AutoCAD.Application")
# draw_FPL_incad(acadapp, "2HM588")


# 根据起点终点和长度以及贝雷梁的具体布置得到分配梁和贝雷梁的交点坐标
# 分配梁词典的结构形式[起点、终点、长度]
# 贝雷梁词典的结构形式[第i排贝雷梁的节点坐标]
def Cross_Bailey_FPL(Bailey_dict, FPL_dict):

    # 建立贝雷梁和分配梁的交点词典
    Bailey_cross_pt_dict = {}
    FPL_cross_pt_dict = {}

    # print(Bailey_dict)
    # print(FPL_dict)

    # 根据分配梁的坐标和贝雷梁的坐标得到相交的点
    for Bailey_dict_i in range(len(Bailey_dict["所有y坐标"])):

        # 第i排贝雷梁的起点和终点
        Bailey_i_pt1 = Bailey_dict[f"第{Bailey_dict_i+1}排贝雷上弦杆接头坐标"][0]
        Bailey_i_pt2 = Bailey_dict[f"第{Bailey_dict_i+1}排贝雷上弦杆接头坐标"][-1]
        # 防止为空
        if f"第{Bailey_dict_i+1}排贝雷与分配梁的交点" not in Bailey_cross_pt_dict:
            Bailey_cross_pt_dict[f"第{Bailey_dict_i+1}排贝雷与分配梁的交点"] = []

        # 将第i排贝雷梁线段与每一根分配梁求交
        for FPL_i in range(len(FPL_dict)):
            # 防止为空
            if f"第{FPL_i+1}排分配梁与贝雷梁的交点" not in FPL_cross_pt_dict:
                FPL_cross_pt_dict[f"第{FPL_i+1}排分配梁与贝雷梁的交点"] = []
            # 求的交点
            cross_pt = line_intersection([Bailey_i_pt1, Bailey_i_pt2], [FPL_dict[f"第{FPL_i+1}根分配梁"][0], FPL_dict[f"第{FPL_i+1}根分配梁"][1]])
            # 记录交点
            if cross_pt != None:
                Bailey_cross_pt_dict[f"第{Bailey_dict_i+1}排贝雷与分配梁的交点"].append([round(coor,1) for coor in cross_pt])
                FPL_cross_pt_dict[f"第{FPL_i+1}排分配梁与贝雷梁的交点"].append([round(coor,1) for coor in cross_pt])
            
    # 对词典去重
    for dict_i in Bailey_cross_pt_dict:
        set_lst = [tuple(x) for x in Bailey_cross_pt_dict[dict_i]]
        Bailey_cross_pt_dict[dict_i] = sorted(list(set(set_lst)))
    for dict_i in FPL_cross_pt_dict:
        set_lst = [tuple(x) for x in FPL_cross_pt_dict[dict_i]]
        FPL_cross_pt_dict[dict_i] = sorted(list(set(set_lst)))

    # 更新 贝雷梁 和 分配梁 词典
    Bailey_Dict = {**Bailey_dict, **Bailey_cross_pt_dict}
    FPL_Dict = {**FPL_dict, **FPL_cross_pt_dict}
    # print(Bailey_Dict)
    # print(FPL_Dict)

    return Bailey_Dict, FPL_Dict


# 根据贝雷梁词典、分配梁词典、钢管桩词典 将分配梁的词典更新
def update_FPLdict_with_SP(dict1, dict2, dict3, seclst1, seclst2):
    # 贝雷梁词典中的坐标原点信息
    O_point = dict1["图纸坐标原点"]
    # 对钢管桩词典的坐标进行同化
    for key, value in dict3.items():
        if "钢管桩桩顶" in key: 
            dict3[key] = (value[0], (round(value[1][0]-O_point[0], 1), round(value[1][1]-O_point[1], 1), 0), value[2], value[3], value[4])
        elif "钢管桩桩底" in key:
            dict3[key] = (value[0], (round(value[1][0]-O_point[0], 1), round(value[1][1]-O_point[1], 1), 0))
        elif "联结系" in key:
            dict3[key] = ((value[0][0], (round(value[0][1][0]-O_point[0], 1), round(value[0][1][1]-O_point[1], 1), 0)), (value[1][0], (round(value[1][1][0]-O_point[0], 1), round(value[1][1][1]-O_point[1], 1), 0)))
    # print(dict1)
    # print(dict2)
    # print(dict3)
    # 将桩顶坐标分配进分配梁
    # 先得到所有分配梁的线段端点坐标组成的表
    FPL_ptlst = [[key, value[0], value[1]] for key, value in dict2.items() if len(key) == 6 or len(key) == 7]
    # print(FPL_ptlst)
    SP_FPL_numlst = []
    for key, value in dict3.items():
        if "钢管桩桩顶" in key:
            SP_point = value[1]
            for ilst in FPL_ptlst:
                # 判断当前钢管桩桩顶坐标在哪根分配梁上
                if is_point_on_segment(SP_point, [ilst[1], ilst[2]]):
                    # 将桩顶所在第几根分配梁写入词典
                    dict3[key] = (value[0], (value[1][0], value[1][1], 0), value[2], value[3], value[4], ilst[0])
                    # 记录分配梁号码和钢管桩号码[分配梁号码, 钢管桩号码]
                    SP_FPL_numlst.append([int(re.findall(r'\d+', ilst[0])[0]), value[0]])
    # print(dict3)
    # print(SP_FPL_numlst)
    # 将钢管桩的坐标加入到分配梁词典中
    for ilst in SP_FPL_numlst:
        if f"第{ilst[0]}根分配梁和钢管桩的交点" not in dict2:
            dict2[f"第{ilst[0]}根分配梁和钢管桩的交点"] = []
        SP_point = dict3[f"第{ilst[1]}根钢管桩桩顶"][1]
        dict2[f"第{ilst[0]}根分配梁和钢管桩的交点"].append(SP_point)
    # print(dict2)
    # 返回分配梁截面和钢管桩截面，并返回截面的mct格式，分配梁截面起始编号为4，联结系截面为固定的C20a/C14a
    sec_str_lst1 = set([(f"{value[3]}分配梁", value[3]) for key, value in dict2.items() if len(key) == 6])
    sec_str_lst2 = set([(f"{value[2]}×{value[3]}钢管桩", value[2], value[3]) for key, value in dict3.items() if "钢管桩桩顶" in key])
    # print(sec_str_lst1)
    # print(sec_str_lst2)
    secid = 4
    section_lst = []
    for sec_lst in sec_str_lst1:
        sec = sec_lst[1]
        if sec[0] == "2":
            FPL_info = list(filter(lambda s: sec in s, seclst2))
        else:
            FPL_info = list(filter(lambda s: sec in s, seclst1))
        section_lst.append("{}, {}".format(secid, FPL_info[0]))
        secid = secid + 1
    # print(section_lst)
    for sec_lst in sec_str_lst2:
        SP_info = "{}, DBUSER, 钢管桩{}×{}, CC, 0, 0, 0, 0, 0, 0, YES, NO, P, 2, {}, {}, 0, 0, 0, 0, 0, 0, 0, 0".format(secid, sec_lst[1], sec_lst[2], sec_lst[1], sec_lst[2])
        section_lst.append(SP_info)
        secid = secid + 1
    # print(section_lst)
    L_lst = [["2C20a", "2C", "2C 20a"], ["2C14a", "2C", "2C 14a"], ["C 20a", "C", "C 20a"], ["C 14a", "C", "C 14a"]]
    for i in range(len(L_lst)):
        L_info = "{}, DBUSER, {}, CC, 0, 0, 0, 0, 0, 0, YES, NO, {}, 1, GB-YB05, {}".format(secid, L_lst[i][0], L_lst[i][1], L_lst[i][2])
        section_lst.append(L_info)
        secid = secid + 1
    # print(section_lst)
    # print(dict2)

    # 贝雷梁词典没变化，分配梁和钢管桩词典均有变化，附带剩余截面的mct信息流
    return dict2, dict3, section_lst


# 创建分配梁节点、单元、结构组、弹性连接
# 需要分配梁词典 以及 贝雷梁下弦杆的节点信息 以获取 弹连节点号
# 需要分配梁的截面型号
# 需要初始节点号和单元号
def creat_FPL_node_element_group_brnd(FPL_dict, elink_dict, sec_lst, startnodeid, startelementid, startelinkid):
    # print(FPL_dict)
    # print(elink_dict)
    # print(sec_lst)
    # print(startnodeid)
    # print(startelementid)
    # 初始化数据
    nodeid = startnodeid
    elementid = startelementid
    elinkid = startelinkid
    Z0 = elink_dict['第1排贝雷下弦杆接头节点号'][0][1][2]
    # 获取分配梁截面编号
    sec_str_lst = set([value[3] for key, value in FPL_dict.items() if len(key) == 6])
    sec_num_lst = []
    for sec_str in sec_str_lst:
        for sec in sec_lst:
            lst1 = sec.split(",")
            if sec_str in lst1:
                sec_num_lst.append([lst1[0], lst1[2]])
    # 获取分配梁的数量
    FPL_quantity = 0
    for key, value in FPL_dict.items():
        if len(key) == 6 or len(key) == 7:
            FPL_quantity = FPL_quantity + 1
    # 获取贝雷梁下弦杆弹连节点坐标和节点号
    Bailey_elink_lst = [x for key, value in elink_dict.items() if "弹连" in key for x in value]

    # 对 FPL_dict 中的坐标先进行编号
    FPL_nodeid_dict1 = {}
    for i in range(FPL_quantity):
        # 分配梁截面类型字符串 "HM588X300"
        sec_str = FPL_dict[f"第{i+1}根分配梁"][3]
        # print(sec_str)
        # print(sec_num_lst)
        # 分配梁截面高度 588
        if sec_str[0] == "2":
            if sec_str[1] == "I":
                FPL_H = int(re.findall(r'\d+', sec_str)[1])*10
            else:
                FPL_H = int(re.findall(r'\d+', sec_str)[1])
        else:
            if sec_str[0] == "I":
                FPL_H = int(re.findall(r'\d+', sec_str)[0])*10
            else:
                FPL_H = int(re.findall(r'\d+', sec_str)[0])
        # 分配梁截面对应的截面编号
        sec_num = [x[0] for x in sec_num_lst if x[1] == sec_str]
        # 生成三类坐标的表
        nodelst1 = [(x[0], x[1], round(Z0 - 50 - float(FPL_H)/2, 1)) for x in FPL_dict[f"第{i+1}根分配梁"] if type(x) in (list, tuple)]
        nodelst2 = [(x[0], x[1], round(Z0 - 50 - float(FPL_H)/2, 1)) for x in FPL_dict[f"第{i+1}排分配梁与贝雷梁的交点"]]
        nodelst3 = [(x[0], x[1], round(Z0 - 50 - float(FPL_H)/2, 1)) for x in FPL_dict[f"第{i+1}根分配梁和钢管桩的交点"]]
        # 去重并根据y坐标排序
        sorted_list = sorted(set(nodelst1 + nodelst2 + nodelst3), key = lambda x: x[1])
        # 排序后按顺序赋予节点号
        indexed_list = [(i + nodeid, coord) for i, coord in enumerate(sorted_list)]
        # 创建坐标和节点号一一对应的词典
        index_dict = {coord: index for index, coord in indexed_list}
        # 根据词典中的坐标，将节点号分配给原表中的坐标
        FPL_nodeid_dict1[f"第{i+1}根分配梁截面"] = [sec_num[0], sec_str, FPL_H]
        FPL_nodeid_dict1[f"第{i+1}根分配梁端点节点号"] = [(index_dict[coord], coord) for coord in nodelst1]
        FPL_nodeid_dict1[f"第{i+1}根分配梁贝雷交点节点号"] = [(index_dict[coord], coord) for coord in nodelst2]
        FPL_nodeid_dict1[f"第{i+1}根分配梁管桩交点节点号"] = [(index_dict[coord], coord) for coord in nodelst3]
        # 下一个循环的开始节点号
        nodeid = FPL_nodeid_dict1[f"第{i+1}根分配梁端点节点号"][-1][0] + 1
    # print(FPL_nodeid_dict1)

    # 将 FPL_nodeid_dict 中的节点信息整理为每根分配梁的节点信息
    FPL_nodeid_dict2 = {}
    for i in range(FPL_quantity):
        FPL_nodeid_dict2[f'第{i+1}根分配梁截面'] = FPL_nodeid_dict1[f'第{i+1}根分配梁截面']
        FPL_nodeid_dict2[f'第{i+1}根分配梁节点号'] = sorted(set(FPL_nodeid_dict1[f"第{i+1}根分配梁端点节点号"] + FPL_nodeid_dict1[f"第{i+1}根分配梁管桩交点节点号"] + FPL_nodeid_dict1[f"第{i+1}根分配梁贝雷交点节点号"]), key = lambda x: x[0])
    # print(FPL_nodeid_dict2)

    # 开始建立节点信息
    nodelst = ["{}, {}, {}, {}".format(x[0], x[1][0], x[1][1],x[1][2])for key, value in FPL_nodeid_dict2.items() if "截面" not in key for x in value]
    # for x in nodelst:
    #     print(x)

    # 建立单元信息的同时，创建结构组词典
    elementlst = []
    FPL_group_dict = {}
    for i in range(FPL_quantity):
        sec_num = int(FPL_nodeid_dict2[f'第{i+1}根分配梁截面'][0])
        nodeid_lst = [x[0] for x in FPL_nodeid_dict2[f'第{i+1}根分配梁节点号']]
        for j in range(nodeid_lst[-1]-nodeid_lst[0]):
            # 单元、材料、截面、节点i、节点j、旋转角度
            elementlst.append("{}, BEAM, 1, {}, {}, {}, 0, 0".format(elementid, sec_num, nodeid_lst[j], nodeid_lst[j+1]))
            # 建立结构组
            if j == 0 or j == (nodeid_lst[-1]-nodeid_lst[0]) - 1:
                if f"第{i+1}根分配梁结构组单元号对" not in FPL_group_dict:
                    FPL_group_dict[f"第{i+1}根分配梁结构组单元号对"] = []
                FPL_group_dict[f"第{i+1}根分配梁结构组单元号对"].append(elementid)
            elementid = elementid + 1
    # for x in elementlst:
    #     print(x)

    # 添加结构组
    group_elementid1 = FPL_group_dict["第1根分配梁结构组单元号对"][0]
    group_elementid2 = FPL_group_dict[f"第{FPL_quantity}根分配梁结构组单元号对"][1]
    FPL_gplst = [f"分配梁, , {group_elementid1} to {group_elementid2}, 0"]

    print("Success: 成功生成分配梁节点、单元及结构组。")

    # 添加和贝雷梁的弹性连接
    # Bailey_elink_lst 包括了所有下弦杆与分配梁的弹连节点号和坐标
    # FPL_nodeid_dict1[f"第{i+1}根分配梁贝雷交点节点号"] 中记录了分配梁和贝雷相交的节点号和坐标
    # print(Bailey_elink_lst)
    # print(FPL_nodeid_dict1)
    elink_ptlst = []
    elink_lst = []
    for key, value in FPL_nodeid_dict1.items():
        if "贝雷" in key:
            for x in value:
                # 分配梁的弹连节点号
                elink_nodeid1 = x[0]
                # 搜索贝雷梁满足条件的弹连节点号
                elink_nodeid2 = [y[0] for y in Bailey_elink_lst if x[1][0] == y[1][0] and x[1][1] == y[1][1]][0]
                elink_ptlst.append([elink_nodeid1, elink_nodeid2])
    for x in elink_ptlst:
        elink_lst.append("{}, {}, {}, GEN, 0, 1e+007, 1e+006, 1e+006, 1e+005, 1e+005, 1e+005, NO, 0.5, 0.5, 弹性连接".format(elinkid, x[0], x[1]))
        elinkid = elinkid + 1
    print("Success: 分配梁与贝雷弹性连接创建成功。")
    return nodelst, elementlst, FPL_gplst, elink_lst, nodeid, elementid, elinkid, FPL_nodeid_dict1

# 生成钢管桩============================================================================================================================================================
# 生成钢管桩============================================================================================================================================================
# 生成钢管桩============================================================================================================================================================

# 识别钢管桩桩顶桩顶和直径
def get_sp_info_fromcad(Gangguanzhuang_select1, Gangguanzhuang_select2, Lianjiexi_select):
    sp_dict = sp(Gangguanzhuang_select1, Gangguanzhuang_select2, Lianjiexi_select)
    return sp_dict

# 根据用户输入的直径壁厚和桩长生成钢管桩平面图
def draw_SP1_incad(acadapp, D, r, L):
    # 获取活动文档和模型空间
    doc = acadapp.ActiveDocument
    msp = doc.ModelSpace
    try:
        while True:
            # 输入指定坐标
            point = doc.Utility.Getpoint()
            insert_point = vtpnt(point)
            # 根据时间定义块名
            block_name = f"SP1_Block_{time_str()}"
            # 生成块
            block = doc.Blocks.Add(insert_point, block_name)
            # 生成圆
            circle1 = block.AddCircle(insert_point, float(D)/2)
            circle1.Layer = "1粗实线"
            circle2 = block.AddCircle(insert_point, (float(D)-2*float(r))/2)
            circle2.Layer = "4虚线"
            # 插入文本
            txt = block.AddText(f"{D}×{r}×{L}", insert_point, float(D)/20)
            txt.Layer = "Defpoints"
            # 插入中心十字线
            startpoint1 = [point[0]-float(D)/10, point[1], point[2]]
            endpoint1 = [point[0]+float(D)/10, point[1], point[2]]
            var_startpoint1 = vtpnt(startpoint1)
            var_endpoint1 = vtpnt(endpoint1)
            line1 = block.AddLine(var_startpoint1, var_endpoint1)
            line1.Layer = "Defpoints"
            startpoint2 = [point[0], point[1]-float(D)/10, point[2]]
            endpoint2 = [point[0], point[1]+float(D)/10, point[2]]
            var_startpoint2 = vtpnt(startpoint2)
            var_endpoint2 = vtpnt(endpoint2)
            line2 = block.AddLine(var_startpoint2, var_endpoint2)
            line2.Layer = "Defpoints"
            # 插入块
            insert_block = msp.InsertBlock(insert_point, block_name, 1, 1, 1, 0)
            # 写入扩展属性
            app_name = "SZQS"
            ex_data = f"{D}×{r}×{L}:{block_name}"
            data_type = vtint([1001, 1000])
            data_lst = vtvariant([app_name, ex_data])
            insert_block.SetXData(data_type, data_lst)
            # 生成Defpoints线型的桩底圆
            circle3 = msp.AddCircle(insert_point, float(D)/2)
            circle3.Layer = "Defpoints"
            circle3.SetXData(data_type, data_lst)
        # if acadapp.GetInput() == "ESC":
        #     break
    except Exception as e:
        # print(f"发生错误: {e}")
        print("已退出当前钢管桩绘图")

# 桩顶平面
def sp(Gangguanzhuang_select1, Gangguanzhuang_select2, Lianjiexi_select):
    # 创建词典和表格
    sp_dict = {}
    circle1_lst = []# 桩顶圆
    circle2_lst = []# 桩底圆
    line_lst = []# 联结系
    # 对于圆获取直径和坐标，对于中心线获取坐标
    for obj in Gangguanzhuang_select1:
        # if obj.EntityName == "AcDbBlockReference" and "SP1_Block" in obj.Name:
        # obj是桩顶平面图块
        Circle1_Center = obj.InsertionPoint
        # 查询块的扩展属性
        block_xdata = get_block_reference_xdata(obj)
        xdata = re.search(r'(\d+)\D+(\d+)\D+(\d+):.*?(\d+)$', block_xdata[1][1])
        # [圆心，直径]
        circle1_lst.append(((round(Circle1_Center[0], 1), round(Circle1_Center[1], 1), 0), xdata.group(1), xdata.group(2), xdata.group(3), xdata.group(4)))
    for obj in Gangguanzhuang_select2:
        # elif obj.EntityName == "AcDbCircle" and obj.Layer == circle_layer_name:
        # obj是桩底平面图块
        Circle2_Center = obj.Center
        # 查询块的扩展属性
        block_xdata = get_block_reference_xdata(obj)
        xdata = re.search(r'(\d+)\D+(\d+)\D+(\d+):.*?(\d+)$', block_xdata[1][1])
        # [圆心，直径]
        circle2_lst.append(((round(Circle2_Center[0], 1), round(Circle2_Center[1], 1), 0), xdata.group(1), xdata.group(2), xdata.group(3), xdata.group(4)))
    for obj in Lianjiexi_select:
        # elif obj.EntityName == "AcDbLine" and obj.Layer == line_layer_name:
        # obj是联结系线段
        StartPoint = obj.StartPoint
        EndPoint = obj.EndPoint
        # [起点，终点]
        line_lst.append(((round(StartPoint[0],1), round(StartPoint[1],1), 0), (round(EndPoint[0],1), round(EndPoint[1],1), 0)))

    # 圆表去重并排序
    set_dict = {}
    for item in circle1_lst:
        key = tuple(item[0][:2])  # 使用 x 和 y 作为键
        if key not in set_dict:
            set_dict[key] = item
    set_lst = list(set_dict.values())
    circle1_lst = sorted(set_lst, key=lambda item: (item[0][0], item[0][1]))
    # print(circle1_lst)
    set_dict = {}
    for item in circle2_lst:
        key = tuple(item[0][:2])  # 使用 x 和 y 作为键
        if key not in set_dict:
            set_dict[key] = item
    set_lst = list(set_dict.values())
    circle2_lst = sorted(set_lst, key=lambda item: (item[0][0], item[0][1]))
    # print(circle2_lst)
    # 线表去重没必要排序
    set_tuple = set()
    set_lst = []
    for item in line_lst:
        # 将两个坐标对转换为元组，并排序，以便忽略顺序
        pair = tuple(sorted(item))
        if pair not in set_tuple:
            set_tuple.add(pair)
            set_lst.append(item)
    line_lst = set_lst
    # 将表格信息排序后放入dict
    for i in range(len(circle1_lst)):
        sp_dict[f"第{i+1}根钢管桩桩顶"] = (i+1, circle1_lst[i][0], int(circle1_lst[i][1]), int(circle1_lst[i][2]), round(float(circle1_lst[i][3]),1), circle1_lst[i][4])
    for i in range(len(circle2_lst)):
        time_mark = sp_dict[f"第{i+1}根钢管桩桩顶"][5]
        sp2_mark = [x[0] for x in circle2_lst if x[4] == time_mark]
        sp_dict[f"第{i+1}根钢管桩桩底"] = (i+1, sp2_mark[0])
    for i in range(len(line_lst)):
        sp_dict[f"第{i+1}道联结系"] = (find_nearest_point_index(line_lst[i][0], circle1_lst), find_nearest_point_index(line_lst[i][1], circle1_lst))
    
    # print(sp_dict)
    return sp_dict


# 生成一根钢管桩桩顶和桩底以及桩顶往下隔1m的3m高联结系坐标以及节点号，将节点分为桩顶节点、桩底节点、联结系上节点，联结系下节点
# 确定一根钢管桩所对应的分配梁截面以确定弹连节点号和z坐标高度
# 联结系节点根据联结系的长度，最多划分为三份
def create_SP_LJX_node_element_grup_brnd(SP_dict, startnodeid, startelementid, startelinkid, FPL_dict, sec_mct, SP_Fengya, SP_Shui_info):
    # print("钢管桩词典", SP_dict)
    # print("分配梁词典", FPL_dict)
    # print("起始节点号", startnodeid)
    # print("起始单元号", startelementid)
    # print("起始弹连号", startelinkid)

    nodeid = startnodeid
    elementid = startelementid
    elinkid = startelinkid
    
    # 生成每一根桩的所有节点，并且将节点分类
    # 钢管桩总数
    SP_quantity = max([int(re.findall(r'\d+', key)[0]) for key in SP_dict if "桩顶" in key])
    # 联结系总数
    LJX_quantity = max([int(re.findall(r'\d+', key)[0]) for key in SP_dict if "联结系" in key])
    # print("钢管桩总数", SP_quantity, ",", "联结系总数", LJX_quantity)
    # 钢管桩节点号词典、单元号词典、受水流力影响的单元号词典
    SP_node_dict = {}
    SP_element_dict = {}
    Flow_SP_element_dict = {}
    # 将截面mct信息转化成表形式
    seclst = [astr.split(",") for astr in sec_mct]

    # 水面对应z坐标
    if '' in SP_Shui_info:
        Flow_Z0 = -1000 * 1000
    else:
        Flow_Z0 = float(SP_Shui_info[2]) * 1000

    # 根据钢管桩总数建立钢管桩
    for i in range(SP_quantity):
        # 当前钢管桩在词典中对应的所有信息
        dict_info1 = SP_dict[f"第{i+1}根钢管桩桩顶"]
        dict_info2 = SP_dict[f"第{i+1}根钢管桩桩底"]
        # 钢管桩所属的分配梁
        # print(SP_dict)
        # print(dict_info1)
        SP_FPL_string = dict_info1[5]
        # 去分配梁词典中查询对应的分配梁的截面高度和节点z坐标
        SP_FPL_sec_height = FPL_dict[SP_FPL_string + "截面"][2]
        SP_FPL_node_z = FPL_dict[SP_FPL_string + "端点节点号"][0][1][2]
        # 当前钢管桩直径、壁厚、桩长
        SP_D = dict_info1[2]
        SP_s = dict_info1[3]
        SP_H = dict_info1[4]
        # 钢管桩桩顶桩底节点坐标
        SP_coord1_Z = SP_FPL_node_z - SP_FPL_sec_height/2
        SP_coord2_Z = SP_FPL_node_z - SP_FPL_sec_height/2 - SP_H
        SP_coord1 = (dict_info1[1][0], dict_info1[1][1], SP_coord1_Z)
        SP_coord2 = (dict_info2[1][0], dict_info2[1][1], SP_coord2_Z)
        # print(SP_coord1, SP_coord2)
        # 不管是不是斜桩，从桩顶往下计算z坐标减1000和4000的坐标作为当前钢管桩的联结系节点，联结系节点分为上节点和下节点
        LJX_coord1 = point_in_line_Z_scale(SP_coord1, SP_coord2, 1000)
        LJX_coord2 = point_in_line_Z_scale(SP_coord1, SP_coord2, 3000)
        # print(LJX_coord1, LJX_coord2)
        Flow_coord = None
        # 当水面高度在桩顶和桩底之间时，计算水面对应的坐标，否则全部都有水流力作用，或全部都没有水流力作用
        if Flow_Z0 < SP_coord1_Z and Flow_Z0 > SP_coord2_Z:
            distance_flowz0_to_SPcoord1Z = abs(SP_coord1_Z - Flow_Z0)
            # 水面对应节点坐标
            Flow_coord = point_in_line_Z_scale(SP_coord1, SP_coord2, distance_flowz0_to_SPcoord1Z)

        # 查询对应的截面号
        for sublst in seclst:
            if "P" in sublst[12] and (str(SP_D) in sublst[14] and str(SP_s) in sublst[15]):
                secid = sublst[0]

        # 将坐标信息写入钢管桩的节点号词典
        if Flow_coord != None:
            SP_coordlst = sorted([SP_coord1, LJX_coord1, LJX_coord2, SP_coord2, Flow_coord], key = lambda x: x[2],  reverse=True)
        elif Flow_coord == None:
            SP_coordlst = sorted([SP_coord1, LJX_coord1, LJX_coord2, SP_coord2], key = lambda x: x[2],  reverse=True)
        if f"第{i+1}根钢管桩节点号" not in SP_node_dict:
            SP_node_dict[f"第{i+1}根钢管桩节点号"] = []
        for coord in SP_coordlst:
            SP_node_dict[f"第{i+1}根钢管桩节点号"].append((nodeid, coord))
            nodeid = nodeid + 1
        SP_node_dict[f"第{i+1}根钢管桩截面号"] = secid
    # print(SP_node_dict)
    # 创建钢管桩节点mct信息流
    nodelst = ["{}, {}, {}, {}".format(sublst[0], sublst[1][0], sublst[1][1], sublst[1][2]) for key, value in SP_node_dict.items() if "截面号" not in key for sublst in value]

    # 创建单元号词典
    for i in range(SP_quantity):
        if f"第{i+1}根钢管桩单元号" not in SP_element_dict:
            SP_element_dict[f"第{i+1}根钢管桩单元号"] = []
        if f"第{i+1}根钢管桩受水流力单元号" not in SP_element_dict:
            Flow_SP_element_dict[f"第{i+1}根钢管桩受水流力单元号"] = []
        lst1 = SP_node_dict[f"第{i+1}根钢管桩节点号"]
        # SP_element_dict[f"第{i+1}根钢管桩截面号"] = SP_node_dict[f"第{i+1}根钢管桩截面号"]
        for j in range(len(lst1) - 1):
            SP_element_dict[f"第{i+1}根钢管桩单元号"].append((elementid, (lst1[j][0], lst1[j][1], lst1[j+1][0], lst1[j+1][1])))
            # 当单元的i端节点z坐标小于水面时，将该单元号放入水流力加载的单元表中
            if Flow_Z0 >= lst1[j][1][2]:
                Flow_SP_element_dict[f"第{i+1}根钢管桩受水流力单元号"].append((elementid, (lst1[j][0], lst1[j][1], lst1[j+1][0], lst1[j+1][1])))
            elementid = elementid + 1
    # print(SP_element_dict)
    # print(Flow_SP_element_dict)

    # 创建单元mct信息流
    elementlst = []
    for i in range(SP_quantity):
        sec_id = SP_node_dict[f"第{i+1}根钢管桩截面号"]
        for j in SP_element_dict[f"第{i+1}根钢管桩单元号"]:
            elementlst.append("{}, BEAM, 1, {}, {}, {}, 0, 0".format(j[0], sec_id, j[1][0], j[1][2]))
    # print(elementlst)

    # 创建钢管桩的结构组，弹连，桩底固结
    # 结构组
    SP_num = 1# 钢管桩编号，用于生成组中的单根钢管桩
    SP_gplsts = []# 单根钢管桩的组所组成的表
    for key, value in SP_element_dict.items():
        if "根钢管桩单元号" in key:
            SP_gplsts.append([f"钢管桩{SP_num}, , {value[0][0]} to {value[-1][0]}, 0"])
            SP_num = SP_num + 1
            # print(value)
    group_elementid1 = SP_element_dict["第1根钢管桩单元号"][0][0]
    group_elementid2 = SP_element_dict[f"第{SP_quantity}根钢管桩单元号"][-1][0]
    SP_gplst = [f"钢管桩, , {group_elementid1} to {group_elementid2}, 0"]
    print("Success: 成功生成钢管桩节点、单元及结构组。")
    # 弹性连接
    elink_lst = []
    for i in range(SP_quantity):
        # 钢管桩桩顶节点信息(节点号，坐标)
        SP_elink_lst = SP_node_dict[f"第{i+1}根钢管桩节点号"][0]
        SP_elink_id = SP_elink_lst[0]# 钢管桩弹连节点号
        SP_elink_coord = SP_elink_lst[1]# 该点的坐标
        # print(SP_elink_id)
        # 当前钢管桩对应的分配梁
        FPL_string = SP_dict[f"第{i+1}根钢管桩桩顶"][5]
        # 找到对应的分配梁和钢管桩弹连的所有节点
        FPL_elink_lst = FPL_dict[FPL_string + "管桩交点节点号"]
        # 根据坐标去寻找对应的弹性连接节点号
        for sublst in FPL_elink_lst:
            FPL_elink_coord = sublst[1] 
            if abs(FPL_elink_coord[0] - SP_elink_coord[0]) < 0.1 and abs(FPL_elink_coord[1] - SP_elink_coord[1]) < 0.1:
                FPL_elink_id = sublst[0]
                # print(FPL_elink_id)
                elink_lst.append("{}, {}, {}, GEN, 0, 1e+007, 1e+006, 1e+006, 1e+005, 1e+005, 1e+005, NO, 0.5, 0.5, 弹性连接".format(elinkid, SP_elink_id, FPL_elink_id))
                elinkid = elinkid + 1
    print("Success: 钢管桩与分配梁弹性连接创建成功。")         
    # 桩底固结
    SP_cons_nodeid_lst = [value[-1][0] for key, value in SP_node_dict.items() if "节点" in key]
    cons_lst = [' '.join(map(str, SP_cons_nodeid_lst)) + ", 111111, 桩底反力"]
    # print(cons_lst)
    print("Success: 固端约束添加成功。")

    LJX_dict = {}
    # 获取每一道联结系对应的两根钢管桩并且获取钢管桩上的联结系坐标
    for i in range(LJX_quantity):
        SP_id1 = SP_dict[f"第{i+1}道联结系"][0][0]# 第1根钢管桩的序号
        SP_id2 = SP_dict[f"第{i+1}道联结系"][1][0]# 第2根钢管桩的序号
        # SP_DD = max(SP_dict[f"第{SP_id1}根钢管桩桩顶"][2], SP_dict[f"第{SP_id2}根钢管桩桩顶"][2])# 两根钢管桩直径取大值
        LJX_lst1_above = SP_node_dict[f"第{SP_id1}根钢管桩节点号"][1]# 左上(节点号，坐标)
        LJX_lst1_below = SP_node_dict[f"第{SP_id1}根钢管桩节点号"][2]# 左下(节点号，坐标)
        LJX_lst2_above = SP_node_dict[f"第{SP_id2}根钢管桩节点号"][1]# 右上(节点号，坐标)
        LJX_lst2_below = SP_node_dict[f"第{SP_id2}根钢管桩节点号"][2]# 右下(节点号，坐标)

        # # 根据一道联结系的4个角点，判断中间会生成的节点，分为above, below
        LJX_L = distance(LJX_lst1_above[1], LJX_lst2_above[1])
        # print(LJX_lst1_above[1], LJX_lst2_above[1], LJX_L)
        # # 联结系分割数目,9m分成3个
        # a = math.ceil(L/3)
        # # 中间生成的节点数 = 联结系数目a - 1

        if f"第{i+1}道联结系斜杆单元号" not in LJX_dict:
            LJX_dict[f"第{i+1}道联结系斜杆单元号"] = [] 

        LJX_dict[f"第{i+1}道联结系上弦杆单元号"] = [[elementid, LJX_lst1_above[0], LJX_lst2_above[0]]]
        elementid = elementid + 1
        LJX_dict[f"第{i+1}道联结系下弦杆单元号"] = [[elementid, LJX_lst1_below[0], LJX_lst2_below[0]]]
        elementid = elementid + 1
        LJX_dict[f"第{i+1}道联结系斜杆单元号"].append([elementid, LJX_lst1_above[0], LJX_lst2_below[0]])
        elementid = elementid + 1
        LJX_dict[f"第{i+1}道联结系斜杆单元号"].append([elementid, LJX_lst2_above[0], LJX_lst1_below[0]])
        elementid = elementid + 1

        # 根据钢管桩直径判断联结系的种类(C20\C14)并查询上弦杆和下弦杆的截面号
        if LJX_L < 6000:
            LJX_sec1 = "2C14a"
            LJX_sec2 = "C 14a"
        elif LJX_L >= 6000:
            LJX_sec1 = "2C20a"
            LJX_sec2 = "C 20a"
        for sublst in seclst:
            if "2C" in sublst[12] and LJX_sec1 in sublst[2]:
                secid1 = sublst[0]
            elif "C" in sublst[12] and "2" not in sublst[12] and LJX_sec2 in sublst[2]:
                secid2 = sublst[0]

        LJX_dict[f"第{i+1}道联结系弦杆截面"] = [secid1, LJX_sec1]
        LJX_dict[f"第{i+1}道联结系斜杆截面"] = [secid2, LJX_sec2]
    
    # 创建联结系单元mct信息流
    for i in range(LJX_quantity):
        # 弦杆
        elementlst1 = [valuelst for key, value in LJX_dict.items() if f"第{i+1}道联结系上弦杆单元号" == key or f"第{i+1}道联结系下弦杆单元号" == key for valuelst in value]
        sec_id = LJX_dict[f"第{i+1}道联结系弦杆截面"][0]
        for j in elementlst1:
            elementlst.append("{}, BEAM, 1, {}, {}, {}, 0, 0".format(j[0], sec_id, j[1], j[2]))
        # 斜杆
        elementlst2 = [valuelst for key, value in LJX_dict.items() if f"第{i+1}道联结系斜杆单元号" == key for valuelst in value]
        sec_id = LJX_dict[f"第{i+1}道联结系斜杆截面"][0]
        for j in elementlst2:
            elementlst.append("{}, BEAM, 1, {}, {}, {}, 0, 0".format(j[0], sec_id, j[1], j[2]))

    # 联结系结构组
    group_elementid1 = min([valuelst[0] for key, value in LJX_dict.items() if "1" in key and "截面" not in key for valuelst in value])
    group_elementid2 = max([valuelst[0] for key, value in LJX_dict.items() if str(LJX_quantity) in key and "截面" not in key for valuelst in value])
    LJX_gplst = [f"联结系, , {group_elementid1} to {group_elementid2}, 0"]
    print("Success: 成功生成联结系单元及结构组。")

    # 钢管桩风荷载线荷载大小为风压kpa乘以直径m
    Wind_Beamload = []
    SP_Feng_kpa = SP_Fengya
    if SP_Feng_kpa >= 0.1:
        for key, value in SP_element_dict.items():
            # 直径
            num = re.findall(r'\d+', key)[0]
            D = SP_dict[f"第{int(num)}根钢管桩桩顶"][2] / 1000
            # 线荷载
            beamload = SP_Feng_kpa * D
            for elelst in value:
                # 单元号
                elementi = elelst[0]
                beamloadstring = f"{elementi} {",LINE,UNILOAD,GY,NO,NO,aDir[1], , , ,0,"} {beamload} {",1,"} {beamload} {",0,0,0,0,,NO,0,0,NO,"}"
                Wind_Beamload.append(beamloadstring)
        print("Success: 钢管桩上风荷载成功生成。")
    else:
        print("Warning: 钢管桩上风荷载未生成。")

    # 独立计算每一根钢管桩的水流力作用，考虑钢管桩与钢管桩之间的间距
    # 获取每一根钢管桩的桩顶坐标并将其分为第一排桩、第二排桩、...、第n排桩，计算第一排桩的无系数水流力，后续桩的水流力乘以遮流影响系数和水深影响系数，将水流力转化成线荷载加载到单元上
    Flow_Beamload = []
    flow_PileD_lst = [] # 所有水流力和对应的钢管桩直径
    if '' in SP_Shui_info:
        Flow_Beamload = []
        flow_PileD_lst = []
        max_flow_lst = []
        print("Warning: 钢管桩上水流力未生成。")
    else:
        flow_SP_grouped_data = {}
        for key, value in SP_dict.items():
            if "桩顶" in key:
                x_coord = value[1][0]  # 获取 x 坐标
                if x_coord not in flow_SP_grouped_data:
                    flow_SP_grouped_data[x_coord] = []  # 初始化一个列表
                flow_SP_grouped_data[x_coord].append(key)  # 将桩号存入列表
        flow_SP_group_dict = {}
        group_counter = 1  # 组号计数器
        for x_coord, SP_s in flow_SP_grouped_data.items():
            group_name = f"组{group_counter}"  # 组名，例如 "组1"
            flow_SP_group_dict[group_name] = SP_s  # 将桩号列表存入结果字典
            group_counter += 1  # 组号递增
        # print(flow_SP_group_dict)
        for i in range(len(flow_SP_group_dict)):
            Flow_lst = flow_SP_group_dict[f"组{i+1}"]
            for j in range(len(Flow_lst)):
                # 桩号
                Flow_num = int(re.findall(r'\d+', Flow_lst[j])[0])
                # 桩顶平面坐标
                Flow_SP_p1 = SP_dict[f"第{Flow_num}根钢管桩桩顶"][1]
                # 桩底平面坐标
                Flow_SP_p2 = SP_dict[f"第{Flow_num}根钢管桩桩底"][1]
                # 直径
                Flow_SP_D = SP_dict[f"第{Flow_num}根钢管桩桩顶"][2]
                # 桩长
                Flow_SP_H = SP_dict[f"第{Flow_num}根钢管桩桩顶"][4]
                # 如果不是第一根，计算其上一根钢管桩的间距
                if j != 0:
                    Flow_SP_p0 = SP_dict[f"第{Flow_num-1}根钢管桩桩顶"][1]
                    Flow_SP_L = abs(Flow_SP_p0[1] - Flow_SP_p1[1])
                else:
                    Flow_SP_L = 0
                # 查询分配梁
                SP_FPL_string = SP_dict[f"第{Flow_num}根钢管桩桩顶"][5]
                # 去分配梁词典中查询对应的分配梁的截面高度和节点z坐标
                SP_FPL_sec_height = FPL_dict[SP_FPL_string + "截面"][2]
                SP_FPL_node_z =   FPL_dict[SP_FPL_string + "端点节点号"][0][1][2]
                # 桩顶z坐标
                Flow_SP_Z1 = SP_FPL_node_z - SP_FPL_sec_height/2
                # 桩底z坐标
                Flow_SP_Z2 = Flow_SP_Z1 - Flow_SP_H
                # 计算该根桩的水流力，返回水面处的线荷载大小
                Fw, Fw_lineload = cal_ShuiLiuLi_JTS_144_1_2010(SP_Shui_info[0], SP_Shui_info[1], SP_Shui_info[2], Flow_SP_Z1/1000, Flow_SP_Z2/1000, Flow_SP_D/1000, Flow_SP_L/1000, j+1)
                flow_PileD_lst.append([Fw, Flow_SP_D/1000])
                # 第i根钢管桩受水流力单元号
                Flow_element_lsts = Flow_SP_element_dict[f"第{Flow_num}根钢管桩受水流力单元号"]
                if Flow_element_lsts != []:
                    for Flow_element_lst in Flow_element_lsts:
                        # 单元号及节点号及节点坐标
                        Flow_element = Flow_element_lst[0]
                        Flow_element_node1_z = Flow_element_lst[1][1][2]  
                        Flow_element_node2_z = Flow_element_lst[1][3][2]
                        # 根据节点z坐标和水流力计算单元i端和j端对应荷载大小
                        Flow_element_node1_lineload = inear_interpolation(Flow_Z0, Flow_SP_Z2, Fw_lineload, 0, Flow_element_node1_z)
                        Flow_element_node2_lineload = inear_interpolation(Flow_Z0, Flow_SP_Z2, Fw_lineload, 0, Flow_element_node2_z)
                        # 该单元的线荷载
                        beamloadstring = f"{Flow_element} {",LINE,UNILOAD,GY,NO,NO,aDir[1], , , ,0,"} {Flow_element_node1_lineload} {",1,"} {Flow_element_node2_lineload} {",0,0,0,0,,NO,0,0,NO,"}"
                        Flow_Beamload.append(beamloadstring)
        # 判断最大的水流力和对应的钢管桩直径
        max_flow_lst = max(flow_PileD_lst, key=lambda x: abs(x[0]))
        print("Success: 钢管桩上水流力成功生成。")

    # print(LJX_dict)
    return nodelst, elementlst, SP_gplst, elink_lst, cons_lst, LJX_gplst, SP_gplsts, Wind_Beamload, Flow_Beamload, max_flow_lst


# 根据选择的截面在cad中插入对应的块，插入点可以自己选择
def insert_block_incad_with_type_name(msp, insert_point, Type, name_str):
    # 比例因子（x, y, z）
    Scale_Factor = [1, 1, 1]
    # 旋转角度（以度为单位）
    RO_Angle = 0
    # 判断块是小肋还是贝雷梁还是分配梁还是钢管桩
    if Type == "贝雷正立面":
        # 贝雷的参数有 "1.5" 和 "3" 
        if name_str == "3":
            block_name = "3m贝雷block"
            # 3m贝雷的扩展属性
            ex_data = "EX_BL3_ZLM"
        elif name_str == "1.5":
            block_name = "1.5m贝雷block"
            # 1.5m贝雷的扩展属性
            ex_data = "EX_BL1_ZLM"
    elif Type == "贝雷平面":
        # 贝雷的参数有 "1.5" 和 "3" 
        if name_str == "3":
            block_name = "3m贝雷平面block"
            # 3m贝雷的扩展属性
            ex_data = "EX_BL3_PM"
        elif name_str == "1.5":
            block_name = "1.5m贝雷平面block"
            # 1.5m贝雷的扩展属性
            ex_data = "EX_BL1_PM"
    elif Type == "贝雷侧立面":
        # 贝雷的参数有 "1.5" 和 "3" 
        if name_str == "3":
            block_name = "3m贝雷立面block"
            # 3m贝雷的扩展属性
            ex_data = "EX_BL3_CLM"
        elif name_str == "1.5":
            block_name = "1.5m贝雷立面block"
            # 1.5m贝雷的扩展属性
            ex_data = "EX_BL1_CLM"
    elif Type == "小肋":
        # 小肋的参数和 I、C、2C 有关，获取cad中对应小肋的块名
        if name_str[0] == "I":
            block_num = name_str.replace("I", "")
            block_name = f"工{block_num}小肋block"
        elif name_str[0] == "C":
            block_num = name_str.replace("C", "")
            block_name = f"槽{block_num}小肋block"
        elif name_str[0] == "2":
            block_num = name_str.replace("2C", "")
            block_name = f"双槽{block_num}小肋block"
        # 小肋的扩展属性
        ex_data = f"EX_XL_{name_str}"
    elif Type == "分配梁立面":
        # 分配梁的参数和 工、HM、HN、2工、2HM、2HN有关
        if name_str[0] == "I":
            block_num = name_str.replace("I", "")
            block_name = f"工{block_num}分配梁block"
        elif name_str[1] == "M":
            block_num = name_str.replace("HM", "")
            block_name = f"HM{block_num}分配梁block"
        elif name_str[1] == "N":
            block_num = name_str.replace("HN", "")
            block_name = f"HN{block_num}分配梁block"
        elif name_str[1] == "I":
            block_num = name_str.replace("2I", "")
            block_name = f"2工{block_num}分配梁block"
        elif name_str[2] == "M":
            block_num = name_str.replace("2HM", "")
            block_name = f"2HM{block_num}分配梁block"
        elif name_str[2] == "N":
            block_num = name_str.replace("2HN", "")
            block_name = f"2HN{block_num}分配梁block"
        # 分配梁的扩展属性
        ex_data = f"EX_FPL_LM_{name_str}"
    elif Type == "分配梁平面":
        # 分配梁的参数和 工、HM、HN、2工、2HM、2HN有关
        if name_str[0] == "I":
            block_num = name_str.replace("I", "")
            block_name = f"工{block_num}分配梁平面block"
        elif name_str[1] == "M":
            block_num = name_str.replace("HM", "")
            block_name = f"HM{block_num}分配梁平面block"
        elif name_str[1] == "N":
            block_num = name_str.replace("HN", "")
            block_name = f"HN{block_num}分配梁平面block"
        elif name_str[1] == "I":
            block_num = name_str.replace("2I", "")
            block_name = f"2工{block_num}分配梁平面block"
        elif name_str[2] == "M":
            block_num = name_str.replace("2HM", "")
            block_name = f"2HM{block_num}分配梁平面block"
        elif name_str[2] == "N":
            block_num = name_str.replace("2HN", "")
            block_name = f"2HN{block_num}分配梁平面block"
        # 分配梁的扩展属性
        ex_data = f"EX_FPL_PM_{name_str}"
    elif Type == "钢管桩":
        # 钢管桩的参数和 半径 有关
        pass
    # 创建对应截面类型的块
    block_ref = msp.InsertBlock(insert_point, block_name, Scale_Factor[0], Scale_Factor[1], Scale_Factor[2], RO_Angle)
    # 扩展属性的对应应用程序名
    app_name = "SZQS"  
    # 扩展属性组码，1001代表应用名称，1000代表字符串数据，数据类型均为变体
    data_type = vtint([1001, 1000])
    data_lst = vtvariant([app_name, ex_data])
    # 添加扩展属性
    block_ref.SetXData(data_type, data_lst)
    # 检查扩展属性
    # xdata = get_block_reference_xdata(block_ref)
    # print(xdata)
