# 1. 标准库
import re
import math

# 3. 本地模块
from General.Midas import API_DIST_FORCE_Unit, MidasAPI


# 获取模型的所有节点、单元、截面、结构组的Json格式
def get_node_elem_sect_grup_dict():

    # 设置单位默认为N和mm
    API_DIST_FORCE_Unit("N", 'MM')

    NODE_dict_res = MidasAPI("GET", "/db/NODE") # 节点信息
    ELEM_dict_res = MidasAPI("GET", "/db/ELEM") # 单元信息
    SECT_dict_res = MidasAPI("GET", "/db/SECT") # 截面信息
    get_structure_group_res = MidasAPI("GET", "/db/GRUP")# 组信息

    return [NODE_dict_res, ELEM_dict_res, SECT_dict_res, get_structure_group_res]


#=====================================================================================================================================================================
# 围檩参数获取（从模型）
def Get_Waler_info_from_Model(Matl_res, SECT_res, Node_res, Elem_res, Structure_group_res, Structure_group_namelst):
    # 设置单位
    API_DIST_FORCE_Unit("N", 'MM')
    # 结构组名
    waler_group_name_pattern = r'^第\d+层围檩$'
    WL_group_namelst = sorted([g for g in Structure_group_namelst if re.match(waler_group_name_pattern, g)],key=lambda x: int(re.findall(r'\d+', x)[0]) if re.findall(r'\d+', x) else 0)
    # 获取该组对应的单元
    WL_group_elem_dict = {}
    for grup_name in WL_group_namelst:
        i = Structure_group_namelst.index(grup_name)
        Elem_id = Structure_group_res['GRUP'][str(i+1)]['E_LIST'][0] # 该结构组的第一个单元号
        WL_group_elem_dict[grup_name] = Elem_id
    # 获取单元对应的截面号和节点
    WL_group_node_dict = {} # 每一层围檩的第一个单元号的任意一个节点
    WL_group_secnum_dict = {} # 每一层围檩的第一个单元号的对应截面号
    Wl_group_material_dict = {}
    for key, elem in WL_group_elem_dict.items():
        WL_group_node_dict[key] = Elem_res['ELEM'][str(elem)]['NODE'][0]
        WL_group_secnum_dict[key] = Elem_res['ELEM'][str(elem)]['SECT']
        Wl_group_material_dict[key] = Elem_res['ELEM'][str(elem)]['MATL']
    # 获取单元对应的材质信息
    WL_group_matl_dict = {} # 每一层围檩的截面名称
    for key, matnum in Wl_group_material_dict.items():
        WL_group_matl_dict[key] = Matl_res['MATL'][str(matnum)]['NAME']
    WL_MATL_lst = list(WL_group_matl_dict.values())
    # 获取单元对应的截面信息
    WL_group_sect_dict = {} # 每一层围檩的截面名称
    for key, secnum in WL_group_secnum_dict.items():
        WL_group_sect_dict[key] = SECT_res['SECT'][str(secnum)]['SECT_NAME']
    WL_SECT_lst = list(WL_group_sect_dict.values())
    # 获取节点对应坐标
    WL_group_node_coor_dict = {} # 每一层围檩的第一个单元号的任意一个节点的坐标
    for key, node in WL_group_node_dict.items():
        WL_group_node_coor_dict[key] = [Node_res['NODE'][str(node)]]
    # 获取z坐标lst
    Waler_Strut_Zlst = [round((val[0]['Z'])/1000, 2) for val in WL_group_node_coor_dict.values()]
    print('第i层围檩的结构组名称:', WL_group_namelst)
    print('第i层围檩的z坐标:', Waler_Strut_Zlst)
    print('第i层围檩的截面:', WL_SECT_lst)
    print('第i层围檩的材质:', WL_MATL_lst)
    print('')
    return WL_group_namelst, Waler_Strut_Zlst, WL_SECT_lst, WL_MATL_lst


# 内支撑参数获取（从模型）
def Get_Strut_info_from_Model(Matl_res, SECT_res, Node_res, Elem_res, Structure_group_res, Structure_group_namelst):
    # 设置单位
    API_DIST_FORCE_Unit("N", 'MM')
    # 结构组名
    XC_group_name_pattern = r'^第\d+层斜撑$'
    DC_group_name_pattern = r'^第\d+层对撑$'
    XC_group_namelst = sorted([g for g in Structure_group_namelst if re.match(XC_group_name_pattern, g)],key=lambda x: int(re.findall(r'\d+', x)[0]) if re.findall(r'\d+', x) else 0)
    DC_group_namelst = sorted([g for g in Structure_group_namelst if re.match(DC_group_name_pattern, g)],key=lambda x: int(re.findall(r'\d+', x)[0]) if re.findall(r'\d+', x) else 0)
    # 获取该组对应的单元
    XC_group_elem_dict = {} # 每一层斜撑的第一个单元号
    DC_group_elem_dict = {} # 每一层对撑的第一个单元号
    for grup_name in XC_group_namelst:
        i = Structure_group_namelst.index(grup_name)
        Elem_id = Structure_group_res['GRUP'][str(i+1)]['E_LIST'][0] # 该结构组的第一个单元号
        XC_group_elem_dict[grup_name] = Elem_id
    for grup_name in DC_group_namelst:
        i = Structure_group_namelst.index(grup_name)
        Elem_id = Structure_group_res['GRUP'][str(i+1)]['E_LIST'][0] # 该结构组的第一个单元号
        DC_group_elem_dict[grup_name] = Elem_id
    # print('每一层斜撑的第一个单元号', XC_group_elem_dict)
    # print('每一层对撑的第一个单元号', DC_group_elem_dict)
    # 获取单元对应的截面号和节点
    XC_group_node_dict = {} # 每一层斜撑的第一个单元号的对应两个节点
    DC_group_node_dict = {} # 每一层对撑的第一个单元号的对应两个节点
    XC_group_secnum_dict = {} # 每一层斜撑的第一个单元号的对应截面号
    DC_group_secnum_dict = {} # 每一层对撑的第一个单元号的对应截面号
    XC_group_matnum_dict = {} # 每一层斜撑的第一个单元号的对应材质号
    DC_group_matnum_dict = {} # 每一层对撑的第一个单元号的对应材质号
    for key, elem in XC_group_elem_dict.items():
        XC_group_node_dict[key] = Elem_res['ELEM'][str(elem)]['NODE'][0:2]
        XC_group_secnum_dict[key] = Elem_res['ELEM'][str(elem)]['SECT']
        XC_group_matnum_dict[key] = Elem_res['ELEM'][str(elem)]['MATL']
    for key, elem in DC_group_elem_dict.items():
        DC_group_node_dict[key] = Elem_res['ELEM'][str(elem)]['NODE'][0:2]
        DC_group_secnum_dict[key] = Elem_res['ELEM'][str(elem)]['SECT']
        DC_group_matnum_dict[key] = Elem_res['ELEM'][str(elem)]['MATL']
    # print('每一层斜撑的第一个单元号的对应两个节点号', XC_group_node_dict)
    # print('每一层对撑的第一个单元号的对应两个节点号', DC_group_node_dict)
    # print('每一层斜撑的第一个单元号的对应截面号', XC_group_secnum_dict)
    # print('每一层对撑的第一个单元号的对应截面号', DC_group_secnum_dict)
    # 获取单元对应的材质信息
    XC_group_matl_dict = {} # 每一层斜撑的材质
    DC_group_matl_dict = {} # 每一层对撑的材质
    for key, matnum in XC_group_matnum_dict.items():
        XC_group_matl_dict[key] = Matl_res['MATL'][str(matnum)]['NAME']
    for key, matnum in DC_group_matnum_dict.items():
        DC_group_matl_dict[key] = Matl_res['MATL'][str(matnum)]['NAME']
    # 获取单元对应的截面信息
    XC_group_sect_name_dict = {} # 每一层斜撑的截面名称
    DC_group_sect_name_dict = {} # 每一层对撑的截面名称
    for key, secnum in XC_group_secnum_dict.items():
        XC_group_sect_name_dict[key] = SECT_res['SECT'][str(secnum)]['SECT_NAME']
    for key, secnum in DC_group_secnum_dict.items():
        DC_group_sect_name_dict[key] = SECT_res['SECT'][str(secnum)]['SECT_NAME']
    # print('每一层斜撑的截面名称', XC_group_sect_name_dict)
    # print('每一层对撑的截面名称', DC_group_sect_name_dict)
    # 获取节点对应坐标
    XC_group_node_coor_dict = {} # 每一层斜撑的第一个单元号的对应两个节点的坐标
    DC_group_node_coor_dict = {} # 每一层对撑的第一个单元号的对应两个节点的坐标
    for key, nodelst in XC_group_node_dict.items():
        XC_group_node_coor_dict[key] = [Node_res['NODE'][str(node)] for node in nodelst]
    for key, nodelst in DC_group_node_dict.items():
        DC_group_node_coor_dict[key] = [Node_res['NODE'][str(node)] for node in nodelst]
    # print('每一层斜撑的第一个单元号的对应两个节点的坐标', XC_group_node_coor_dict)
    # print('每一层对撑的第一个单元号的对应两个节点的坐标', DC_group_node_coor_dict)
    # 获取长度
    XC_group_len_dict = {}
    DC_group_len_dict = {}
    for key, points in XC_group_node_coor_dict.items():
        p1, p2 = points
        XC_len = round(math.sqrt((p2['X'] - p1['X'])**2 + (p2['Y'] - p1['Y'])**2 + (p2['Z'] - p1['Z'])**2)/1000,3)
        XC_group_len_dict[key] = XC_len
    for key, points in DC_group_node_coor_dict.items():
        p1, p2 = points
        DC_len = round(math.sqrt((p2['X'] - p1['X'])**2 + (p2['Y'] - p1['Y'])**2 + (p2['Z'] - p1['Z'])**2)/1000,3)
        DC_group_len_dict[key] = DC_len
    # print('每一层斜撑的长度', XC_group_len_dict)
    # print('每一层对撑的长度', DC_group_len_dict)
    print('第i层对撑的结构组名称:', DC_group_namelst)
    print('第i层对撑的截面:', DC_group_sect_name_dict)
    print('第i层对撑的长度:', DC_group_len_dict)
    print('第i层对撑的材质:', DC_group_matl_dict)
    print('第i层斜撑的结构组名称:', XC_group_namelst)
    print('第i层斜撑的截面:', XC_group_sect_name_dict)
    print('第i层斜撑的长度:', XC_group_len_dict)
    print('第i层斜撑的材质:', XC_group_matl_dict)

    print('')
    return DC_group_namelst, XC_group_namelst, DC_group_sect_name_dict, XC_group_sect_name_dict, DC_group_len_dict, XC_group_len_dict, DC_group_matl_dict, XC_group_matl_dict