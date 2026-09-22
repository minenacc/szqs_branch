# 3. 本地模块
from General.AutoCAD import SelectOnScreen, ss_to_lst, getptsfromplwithcricle

from FEM_MidasCivil.Double_Walls_CofferDam_FEA.supports import get_midpoint_within_4points, LookForTheSame_XorY, change_list_order, Points_Counterclockwise, Offset_points, ListToTuple_with_OriginalSequence


#得到平面视图上外壁板、内壁版、内支撑、隔舱板的节点坐标，与Midas中节点的x、y坐标有关
def Get_Cad_PlanView_LinesToPoints(acadapp):
    SelectOnScreenlist = SelectOnScreen(acadapp,"请选择围堰平面简布图：")
    acaddoc1 = SelectOnScreenlist[0]
    acadmsp1 = SelectOnScreenlist[1]
    slt1 = SelectOnScreenlist[2]
    handle_lst = ss_to_lst(slt1)#所有捕捉的图元的句柄
    handle_layer1_resultlst = []
    handle_layer2_resultlst = []
    handle_layer3_resultlst = []
    handle_layer4_resultlst = []
    handle_layer5_resultlst = []
    for x in handle_lst:
        layer = acaddoc1.HandleToObject(x)#处理当前文档中的句柄字符串
        layer_name = layer.Layer#图层的名字，如“6标注线”
        if layer_name =="1粗实线":
            #1粗实线代表外壁板
            #线段或多段线组成的矩形外壁板的四个角点
            handle_layer1_resultlst = handle_layer1_resultlst + handle_layer_type1(layer)
        handle_layer1_resultlst = list(set(tuple(point) for point in handle_layer1_resultlst))#变成元组后删除重复元素
        if layer_name =="2细实线":
            #2细实线代表内壁板
            #线段或多段线组成的矩形内壁板的四个角点
            handle_layer2_resultlst = handle_layer2_resultlst + handle_layer_type1(layer)
        handle_layer2_resultlst = list(set(tuple(point) for point in handle_layer2_resultlst))#变成元组后删除重复元素
        if layer_name =="3中心线":
            #3中心线代表内支撑，一段连续的内支撑必须由多段线表示
            #多段线组成的内支撑的所有顶点
            handle_layer3_resultlst.append(handle_layer_type1(layer))
        if layer_name =="4虚线":
            #4虚线代表隔舱板
            #线段组成的隔舱板的端点
            handle_layer4_resultlst.append(handle_layer_type1(layer))
        if layer_name =="5文本":
            #5文本代表大内撑
            #线段所代表的大内撑的所有顶点
            handle_layer5_resultlst.append(handle_layer_type1(layer))
    #外壁板四角点坐标、内壁板四角点坐标、内支撑所有顶点坐标(分组)、隔舱板所有端点坐标(分组)
    return handle_layer1_resultlst, handle_layer2_resultlst, handle_layer3_resultlst, handle_layer4_resultlst,handle_layer5_resultlst


#处理外壁板，内壁板、内支撑、隔舱板同理
def handle_layer_type1(handlei):
    if handlei.ObjectName == 'AcDbLine':
        #一条线段上的端点和终点组成的表
        handle_points_list = []
        #线段的端点
        handle_points_list.append([round(handlei.StartPoint[0], 1), round(handlei.StartPoint[1], 1), round(handlei.StartPoint[2], 1)])
        #线段的终点
        handle_points_list.append([round(handlei.EndPoint[0], 1), round(handlei.EndPoint[1], 1), round(handlei.EndPoint[2], 1)])
    else:
        #获得多段线上的端点
        handle_points_list = getptsfromplwithcricle(handlei,1)
    #线段或多段线的各个顶点
    return handle_points_list


#处理Get_Cad_PlanView_LinesToPoints(acadapp)的返回值，分别是:
#外壁板四角点坐标
#内壁板四角点坐标
#内支撑所有顶点坐标
#隔舱板所有端点坐标
def Handle_Cad_PlanView_LinesToPoints(acadapp):
    Get_Cad_PlanView_LinesToPoints_resultlst = Get_Cad_PlanView_LinesToPoints(acadapp)
    # print(Get_Cad_PlanView_LinesToPoints_resultlst)
    #先得到内外壁板的中间点并判断是否在同一个点上
    # print(Get_Cad_PlanView_LinesToPoints_resultlst[0])
    # print(Get_Cad_PlanView_LinesToPoints_resultlst[1])
    midpoint1 = get_midpoint_within_4points(Get_Cad_PlanView_LinesToPoints_resultlst[0])
    midpoint2 = get_midpoint_within_4points(Get_Cad_PlanView_LinesToPoints_resultlst[1])
    # print(midpoint1)
    # print(midpoint2)
    if midpoint1 == midpoint2:
        #print("内外壁板中心点在同一个坐标原点")
        midpoint = midpoint1
    #得到减去中心原点后的各个点坐标
    for x0 in range(len(Get_Cad_PlanView_LinesToPoints_resultlst)):
        if x0 == 0:
            layer1_pointlst = [(round(x - midpoint[0],1), round(y - midpoint[1],1), round(z - midpoint[2],1)) for x, y, z in Get_Cad_PlanView_LinesToPoints_resultlst[x0]]
        elif x0 == 1:
            layer2_pointlst = [(round(x - midpoint[0],1), round(y - midpoint[1],1), round(z - midpoint[2],1)) for x, y, z in Get_Cad_PlanView_LinesToPoints_resultlst[x0]]
        elif x0 == 2:
            layer3_pointlst = [[(round(x - midpoint[0],1), round(y - midpoint[1],1), round(z - midpoint[2],1)) for x, y, z in group ] for group in Get_Cad_PlanView_LinesToPoints_resultlst[x0]]
        elif x0 == 3:
            layer4_pointlst = [[(round(x - midpoint[0],1), round(y - midpoint[1],1), round(z - midpoint[2],1)) for x, y, z in group ] for group in Get_Cad_PlanView_LinesToPoints_resultlst[x0]]
        else:
            layer5_pointlst = [[(round(x - midpoint[0],1), round(y - midpoint[1],1), round(z - midpoint[2],1)) for x, y, z in group ] for group in Get_Cad_PlanView_LinesToPoints_resultlst[x0]]
    # print(layer3_pointlst)
    layer1_pointlst_y = [y[1] for y in layer1_pointlst]
    layer2_pointlst_y = [y[1] for y in layer2_pointlst]
    layer1_pointlst_x = [x[0] for x in layer1_pointlst]
    layer2_pointlst_x = [x[0] for x in layer2_pointlst]
    layer1_ymax = max(layer1_pointlst_y)
    layer1_xmax = max(layer1_pointlst_x)
    layer2_ymax = max(layer2_pointlst_y)
    layer2_xmax = max(layer2_pointlst_x)
    layer1_layer2_dist = abs(layer1_ymax-layer2_ymax)#内壁板和外壁板的间距
    #内支撑的点中和外壁板重复的点，分x重合和y重合
    layer3_points_inlayer1 = LookForTheSame_XorY(layer1_pointlst,layer3_pointlst)
    #隔舱板的点中和外壁板重复的点，分x重合和y重合
    layer4_points_inlayer1 = LookForTheSame_XorY(layer1_pointlst,layer4_pointlst)
    #内支撑的点中和内壁板重复的点，分x重合和y重合
    layer3_points_inlayer2 = LookForTheSame_XorY(layer2_pointlst,layer3_pointlst)
    #隔舱板的点中和内壁板重复的点，分x重合和y重合
    layer4_points_inlayer2 = LookForTheSame_XorY(layer2_pointlst,layer4_pointlst)
    #大内撑的在内壁板上的点
    layer5_points_inlayer2 = []
    [layer5_points_inlayer2.append(y) for x in range(len(layer5_pointlst)) for y in layer5_pointlst[x]]
    #将所有坐标集合，则得到外壁板和内壁版的所有点
    layer1_pointlst_apply = list(set(layer1_pointlst+layer3_points_inlayer1+layer4_points_inlayer1))
    layer2_pointlst_apply = list(set(layer2_pointlst+layer3_points_inlayer2+layer4_points_inlayer2+layer5_points_inlayer2))
    layer1_pointlst_apply_allx = [pt[0] for pt in layer1_pointlst_apply]
    layer2_pointlst_apply_allx = [pt[0] for pt in layer2_pointlst_apply]
    layer1_pointlst_apply_ally = [pt[1] for pt in layer1_pointlst_apply]
    layer2_pointlst_apply_ally = [pt[1] for pt in layer2_pointlst_apply]
    layer_pointlst_apply_allx = sorted(list(set(x for x in layer1_pointlst_apply_allx+layer2_pointlst_apply_allx)))
    layer_pointlst_apply_ally = sorted(list(set(y for y in layer1_pointlst_apply_ally+layer2_pointlst_apply_ally)))
    layer1_pointlst_overlay = []
    layer2_pointlst_overlay = []
    for x in layer_pointlst_apply_allx:
        layer1_pointlst_overlay.append([x, layer1_ymax, 0])
        layer1_pointlst_overlay.append([x,layer1_ymax*-1,0])
        if abs(x)-layer2_xmax<=0 :
            layer2_pointlst_overlay.append([x,layer2_ymax,0])
            layer2_pointlst_overlay.append([x,layer2_ymax*-1,0])
    for y in layer_pointlst_apply_ally:
        layer1_pointlst_overlay.append([layer1_xmax,y,0])
        layer1_pointlst_overlay.append([layer1_xmax*-1,y,0])                
        if abs(y)-layer2_ymax<=0 :
            layer2_pointlst_overlay.append([layer2_xmax,y,0])
            layer2_pointlst_overlay.append([layer2_xmax*-1,y,0])      
    #开始按照两点之间直线最短的规律排序点
    #外壁板按连线排序的点 
    layer1_pointlst_overlay.sort(key=lambda p: (-p[0], -p[1]))
    layer2_pointlst_overlay.sort(key=lambda p: (-p[0], -p[1]))
    layer1_startpoint = layer1_pointlst_overlay[0]
    layer2_startpoint = layer2_pointlst_overlay[0]
    layer1_pointlst_inorder = change_list_order(layer1_startpoint,Points_Counterclockwise(layer1_pointlst_overlay))
    layer2_pointlst_inorder = change_list_order(layer2_startpoint,Points_Counterclockwise(layer2_pointlst_overlay))
    layer1_pointlst_sort = [item for index, item in enumerate(layer1_pointlst_inorder) if index == layer1_pointlst_inorder.index(item)]
    layer2_pointlst_sort = [item for index, item in enumerate(layer2_pointlst_inorder) if index == layer2_pointlst_inorder.index(item)]
    # #根据刃脚的宽度得到平面图上各点的x和y的坐标，坐标和内壁板对应
    acaddoc1 = acadapp.ActiveDocument
    CutEdge_w = acaddoc1.Utility.getstring ( 1, "请输入刃脚底板宽度:")
    #得到顺时针排序后的底层刃脚的点
    CutEdge_ptlst1 = Offset_points(layer2_pointlst_sort,layer2_pointlst,(layer1_layer2_dist-float(CutEdge_w)))
    # #得到顺时针排序后的中层刃脚的点
    CutEdge_ptlst2 = Offset_points(layer2_pointlst_sort,layer2_pointlst,(layer1_layer2_dist-float(CutEdge_w))/2)
    # #外壁板按连线排序的点、内壁板按连线排序的点、内支撑按多段线分类的点、隔舱板按线段分类的点、大内撑的点、底层刃脚的点、中层刃脚的点、外壁板的角点、内壁板的角点
    return layer1_pointlst_sort, layer2_pointlst_sort,layer3_pointlst,layer4_pointlst,layer5_pointlst,CutEdge_ptlst1,CutEdge_ptlst2,layer1_pointlst,layer2_pointlst        


#从立面图中获取围堰的接高段数，围堰每段的接高长度以及内支撑、水平环板的位置
#立面图壁板用1粗实线，内支撑和水平环板用3中心线表示
def Get_Cad_ElevationView_LinesToPoints(acadapp):
    #handle_layer1_resultlst = []
    handle_layer1_infolst = []
    handle_layer3_resultlst = []
    handle_layer5_resultlst = []
    while True:
        SelectOnScreenlist = SelectOnScreen(acadapp,"请按从低往高的顺序选择围堰节段立面简布图(输入'n'结束选择): ")
        acaddoc1 = SelectOnScreenlist[0]
        acadmsp1 = SelectOnScreenlist[1]
        slt1 = SelectOnScreenlist[2]
        handle_lst = ss_to_lst(slt1)#所有捕捉的图元的句柄
        #handle_layer1_ptlst = []#记录一段围堰的图层中的点
        handle_layer1_info = []#记录一段围堰的高度和起始y坐标
        handle_layer3_ptlst = []#记录一段围堰的水平环板的y坐标
        handle_layer5_ptlst = []#记录一段围堰的大内撑的y坐标
        for x in handle_lst:
            layer = acaddoc1.HandleToObject(x)#处理当前文档中的句柄字符串
            layer_name = layer.Layer#图层的名字，如“6标注线”
            if layer_name =="1粗实线":
                #1粗实线代表外壁板
                #两道板的四个点
                layer1_ptlst = handle_layer_type1(layer)#一条竖线的两个端点
                h1 = round(abs(layer1_ptlst[0][1]-layer1_ptlst[1][1]),1)#该段围堰的高度
                start_y1 = min(layer1_ptlst[0][1],layer1_ptlst[1][1])#cad中的起始y坐标
                #handle_layer1_ptlst = handle_layer1_ptlst+layer1_ptlst
                handle_layer1_info = handle_layer1_info+[h1,start_y1]
            if layer_name =="3中心线":
                #3中心线代表内支撑和水平环板的位置
                #去获取y坐标的位置并排序
                layer3_ptlst = handle_layer_type1(layer)#一条水平线的两个端点
                handle_layer3_ptlst.append(layer3_ptlst[0][1])#只需要y坐标
            if layer_name =="5文本":
                #5文本代表大内撑的位置
                #去获取y坐标的位置并排序
                layer5_ptlst = handle_layer_type1(layer)#一条水平线的两个端点
                handle_layer5_ptlst.append(layer5_ptlst[0][1])#只需要y坐标
        #有x组数据说明有x段，每一段表中的包含该段的所有符合要求的节点
        #handle_layer1_resultlst.append(handle_layer1_ptlst)
        #对于外壁板，得到由围堰节段高度和起始y坐标的表，也是分围堰段数的
        handle_layer1_infolst.append(ListToTuple_with_OriginalSequence(handle_layer1_info))
        #也是分段的，每一段代表一段围堰的内支撑和水平环板的所有y坐标位置，每一段中的一个数字代表一个y坐标
        handle_layer3_resultlst.append(sorted(handle_layer3_ptlst))
        handle_layer5_resultlst.append(sorted(handle_layer5_ptlst))
        #通过用户在cad中的输入来判断是否继续选取
        user_input = acaddoc1.Utility.getstring ( 1, "是否继续选取围堰接高段(Y/N):")
        if user_input.lower() == 'n':#兼容大小写输入
            break
    #一段围堰节段高度和起始y坐标的表、一段围堰的内支撑和水平环板的所有y坐标、一段围堰的大内撑的所有y坐标
    return handle_layer1_infolst, handle_layer3_resultlst,handle_layer5_resultlst


#处理y坐标的信息
def Handle_Cad_ElevationView_LinesToPoints(acadapp):
    acaddoc1 = acadapp.ActiveDocument
    CutEdge_h = acaddoc1.Utility.getstring ( 1, "请输入刃脚高度:")
    CutEdge_h1 = float(CutEdge_h)*-1#刃脚底层高度
    CutEdge_h2 = float(CutEdge_h)/2*-1#刃脚中层高度
    Get_Cad_ElevationView_LinesToPoints_resultlst = Get_Cad_ElevationView_LinesToPoints(acadapp)
    a0 = Get_Cad_ElevationView_LinesToPoints_resultlst[0]
    a1 = Get_Cad_ElevationView_LinesToPoints_resultlst[1]
    a2 = Get_Cad_ElevationView_LinesToPoints_resultlst[2]
    a0lst = [0]
    a1lst = []
    a2lst = []
    H = 0#围堰节段的起始y坐标
    for x in range(len(a1)):
        a1lst = a1lst + [round(y-a0[x][1]+H,1) for y in a1[x]]
        H = H + a0[x][0]
        a0lst.append(H)
    H = 0
    for x in range(len(a2)):
        a2lst = a2lst + [round(y-a0[x][1]+H,1) for y in a2[x]]
        H = H + a0[x][0]
    ylist = list(set(a0lst+a1lst+a2lst))
    ylist.sort()
    #所有y坐标、围堰节段对应y坐标、水平环板对应y坐标
    return ylist,a0lst,a1lst,a2lst,CutEdge_h1,CutEdge_h2