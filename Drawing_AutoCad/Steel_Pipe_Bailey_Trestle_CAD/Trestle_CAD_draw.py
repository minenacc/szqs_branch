# 1. 标准库
import math

# 3. 本地模块
from General.AutoCAD import (
    vtint, vtvariant, vtpnt, pointlist_extend, insert_block_incad_with_block_name, Load_Font, Load_Standrad_Style,
    make_line, make_polyline, Add_Header, Add_Annotation_Linear, Add_Annotation_MLeader, Add_Text, Add_MText
)

from Drawing_AutoCad.Steel_Pipe_Bailey_Trestle_CAD.Trestle_MCT_format_CAD import build_mct_from_ui


# 转换坐标形式
def strip_node_ids(group_nodes: dict) -> dict:
    grouped_coords = {}
    for group_id, nodes in group_nodes.items():
        # 先把 (x, y) 作为键分组
        xy_dict = {}
        for _, coord in nodes:
            x, y, z = coord
            xy_key = (round(x, 6), round(y, 6)) 
            xy_dict.setdefault(xy_key, []).append(coord)
        # 将每根桩的节点列表加入组
        grouped_coords[group_id] = list(xy_dict.values())
    return grouped_coords

def merge_brake_pile_groups(brake_pile_group_nodes: dict) -> dict:
    """
    将原始的 brake_pile_group_nodes（按 x 分组）：
        {1: [[(x,y,z),...], [(x,y,z),...]], 2: [[(x,y,z),...], ...]}
    转换为统一一组的结构：
        {1: [[(x,y,z), (x,y,z), ...], [(x,y,z), (x,y,z), ...], ...]}

    即：把所有组里的“单根桩”汇总到一个大组内，供 CAD 绘图使用。
    """
    merged = {1: []}  # 合并后只有一个组编号
    for gid, pile_list in brake_pile_group_nodes.items():
        for single_pile in pile_list:
            merged[1].append(single_pile)
    return merged

# 小肋
xl_single_nodes = {
    1: [(1, (200.0, 500.0, 0.0)), (2, (200.0, 1400.0, 0.0)), (3, (200.0, 2300.0, 0.0)), (4, (200.0, 3200.0, 0.0)), (5, (200.0, 4100.0, 0.0)), (6, (200.0, 5000.0, 0.0))],
    2: [(7, (600.0, 500.0, 0.0)), (8, (600.0, 1400.0, 0.0)), (9, (600.0, 2300.0, 0.0)), (10, (600.0, 3200.0, 0.0)), (11, (600.0, 4100.0, 0.0)), (12, (600.0, 5000.0, 0.0))], 
    3: [(13, (1000.0, 500.0, 0.0)), (14, (1000.0, 1400.0, 0.0)), (15, (1000.0, 2300.0, 0.0)), (16, (1000.0, 3200.0, 0.0)), (17, (1000.0, 4100.0, 0.0)), (18, (1000.0, 5000.0, 0.0))], 
    4: [(19, (1400.0, 500.0, 0.0)), (20, (1400.0, 1400.0, 0.0)), (21, (1400.0, 2300.0, 0.0)), (22, (1400.0, 3200.0, 0.0)), (23, (1400.0, 4100.0, 0.0)), (24, (1400.0, 5000.0, 0.0))], 
    5: [(25, (1800.0, 500.0, 0.0)), (26, (1800.0, 1400.0, 0.0)), (27, (1800.0, 2300.0, 0.0)), (28, (1800.0, 3200.0, 0.0)), (29, (1800.0, 4100.0, 0.0)), (30, (1800.0, 5000.0, 0.0))],
}
# 贝雷梁
bailey_single_nodes = {
    1:[(0, 1800, 0), (3000, 1800, 0), (6000, 1800, 0)], # 单排的每一个起点
    2:[(0, 2700, 0), (3000, 2700, 0), (6000, 2700, 0)],
    3:[(0, 3600, 0), (3000, 3600, 0), (6000, 3600, 0)],
}
# 分配梁
fenpei_single_nodes = {
    1: [(7, (3600.0, 500.0, -2000.0)), (10, (3600.0, 3200.0, -2000.0)), (11, (3600.0, 4100.0, -2000.0)), (12, (3600.0, 5000.0, -2000.00))], 
    2: [(13, (5100.0, 500.0, -2000.0)), (16, (5100.0, 3200.0, -2000.0)), (17, (5100.0, 4100.0, -2000.0)), (18, (5100.0, 5000.0, -2000.0))], 
    3: [(13, (0.0, 500.0, -2000.0)), (16, (0.0, 3200.0, -2000.0)), (17, (0.0, 4100.0, -2000.0)), (18, (0.0, 5000.0, -2000.0))]
}
# 钢管桩
ganggz_single_nodes = {
    1:[[(3600, 1500, -2500), (3600, 1500, -12500)], [(3600, 2750, -2500), (3600, 2750, -12500)], [(3600, 4000, -2500), (3600, 4000, -12500)]], # 一排有3根
    2:[[(5100, 1500, -2500), (5100, 1500, -12500)], [(5100, 4000, -2500), (5100, 4000, -12500)]], # 一排有两根
    3: [[(0, 1500, -2500), (0, 1500, -12500)], [(0, 4000, -2500), (0, 4000, -12500)]]
} # 2排非制动墩钢管桩
# 制动墩顶分配梁1(上面的)
abutmt_fenpei1_single_nodes = {
    1: [(13, (8900.0, 500.0, -2000.0)), (16, (8900.0, 3200.0, -2000.0)), (17, (8900.0, 4100.0, -2000.0)), (18, (8900.0, 5000.0, -2000.0))], 
}
# abutmt_fenpei1_single_nodes = {  
#     1: [(13, (8900.0, 500.0, -2000.0)), (16, (8900.0, 3200.0, -2000.0)), (17, (8900.0, 4100.0, -2000.0)), (18, (8900.0, 5000.0, -2000.0))], 
#     # 2: [(13, (0.0, 500.0, -2000.0)), (16, (0.0, 3200.0, -2000.0)), (17, (0.0, 4100.0, -2000.0)), (18, (0.0, 5000.0, -2000.0))], 
# } # 两个制动墩
# abutmt_fenpei1_single_nodes = {}
# 制动墩顶分配梁2(下面的)
abutmt_fenpei2_single_nodes = {
    1: [(1, (7000.0, 1250.0, -2500.0)), (4, (8000.0, 1250.0, -2500.0)), (5, (9000.0, 1250.0, -2500.0)), (6, (10800.0, 1250.0, -2500.0))],
    2: [(7, (7000.0, 4250.0, -2000.0)), (10, (8000.0, 4250.0, -2000.0)), (11, (9000.0, 4250.0, -2000.0)), (12, (10800.0, 4250.0, -2000.00))], 
}
# abutmt_fenpei2_single_nodes = {
#     1: [(1, (7600.0, 1250.0, -2500.0)), (4, (8000.0, 1250.0, -2500.0)), (5, (9000.0, 1250.0, -2500.0)), (6, (10200.0, 1250.0, -2500.0))],
#     2: [(7, (7600.0, 4250.0, -2000.0)), (10, (8000.0, 4250.0, -2000.0)), (11, (9000.0, 4250.0, -2000.0)), (12, (10200.0, 4250.0, -2000.00))], 
#     3: [(1, (-3000.0, 1250.0, -2500.0)), (4, (1000.0, 1250.0, -2500.0)), (5, (2000.0, 1250.0, -2500.0)), (6, (3000.0, 1250.0, -2500.0))],
#     4: [(7, (-3000.0, 4250.0, -2000.0)), (10, (1000.0, 4250.0, -2000.0)), (11, (2000.0, 4250.0, -2000.0)), (12, (3000.0, 4250.0, -2000.00))], 
# } # 两个制动墩
# abutmt_fenpei2_single_nodes = {}
abutmt_ganggz_single_nodes = {
    # 1:[
    #     [(-1000, 3000, -2500), (-1000, 3000, -12500)],
    #     [(-1000, 4300, -2500), (-1000, 4300, -12500)],
    #     [(1000, 3000, -2500), (1000, 3000, -12500)],
    #     [(1000, 4300, -2500), (1000, 4300, -12500)],
    # ], # 每个制动墩有4根桩
    1:[
        [(7900, 1250, -2500), (7900, 1250, -12500)],
        [(7900, 4250, -2500), (7900, 4250, -12500)],
        [(9900, 1250, -2500), (9900, 1250, -12500)],
        [(9900, 4250, -2500), (9900, 4250, -12500)],
    ],
} # 两个制动墩
# abutmt_ganggz_single_nodes = {}

# 截面信息
xl_section_cad = {'工字钢截面': {'H': 180, 'B': 1, 'tw': 1, 'tf': 1, 'r': 1}}
xl_section_name = '工字钢截面'
distribute_section_cad = {'双拼工字钢截面': {'H': 320, 'B': 100, 'tw': 1, 'tf': 1, 'C': 1}}
distribute_section_name = '双拼工字钢截面'
gz_section_cad = {'管型截面': {'D': 630, 'd': 8}}
gz_section_name = '管型截面'
zd_f1_section_cad = {'双拼工字钢截面': {'H': 360, 'B': 100, 'tw': 1, 'tf': 16, 'C': 1}}
zd_f1_section_name = '双拼工字钢截面'
zd_f2_section_cad = {'双拼HM截面': {'H': 588, 'B': 600, 'tw': 1, 'tf': 20, 'r': 1, 'C': 300}}
zd_f2_section_name = '双拼HM截面'
zd_section_cad = {'管型截面': {'D': 630, 'd': 8}}
zd_section_name = '管型截面'

# 延米重
# xl_section_weight_per_meter = 30
# distribute_section_per_meter = 40
# gz_section_weight_per_meter = 50
# zd_f1_section_weight_per_meter = 60
# zd_f2_section_weight_per_meter = 70
# zd_section_weight_per_meter = 80
section_weight_per_meter_lst = [30, 40, 50, 60, 70]

# {'工字钢截面': {'H': 1, 'B': 1, 'tw': 1, 'tf1': 1, 'r': 1}}
# {'截面类型': '槽钢截面', '参数': ['H(mm): 1', 'B(mm): 1', 'tw(mm): 1', 'tf1(mm): 1', 'tf2(mm): 1', 'r(mm): 1']}
# {'截面类型': '管型截面', '参数': ['D(mm): 1', 'd(mm): 1']}
# {'截面类型': 'H型截面', '参数': ['H(mm): 1', 'B(mm): 1', 'tw(mm): 1', 'tf1(mm): 1', 'r(mm): 1']}
# {'截面类型': '双拼工字钢截面', '参数': ['H(mm): 1', 'B(mm): 1', 'tw(mm): 1', 'tf1(mm): 1', 'C(mm): 1']}
# {'截面类型': '双拼槽钢截面', '参数': ['H(mm): 1', 'B(mm): 1', 'tw(mm): 1', 'tf1(mm): 1', 'tf2(mm): ', 'r(mm): 1', 'C(mm): 1']}
# {'截面类型': '双拼H型截面', '参数': ['H(mm): 1', 'B(mm): 1', 'tw(mm): 1', 'tf1(mm): 1', 'r(mm): 1', 'C(mm): 1']}
# 钢管桩



# 转换坐标形式
def strip_node_ids(group_nodes: dict) -> dict:
    grouped_coords = {}
    for group_id, nodes in group_nodes.items():
        # 先把 (x, y) 作为键分组
        xy_dict = {}
        for _, coord in nodes:
            x, y, z = coord
            xy_key = (round(x, 6), round(y, 6)) 
            xy_dict.setdefault(xy_key, []).append(coord)
        # 将每根桩的节点列表加入组
        grouped_coords[group_id] = list(xy_dict.values())
    return grouped_coords

def merge_brake_pile_groups(brake_pile_group_nodes: dict) -> dict:
    """
    将原始的 brake_pile_group_nodes（按 x 分组）：
        {1: [[(x,y,z),...], [(x,y,z),...]], 2: [[(x,y,z),...], ...]}
    转换为统一一组的结构：
        {1: [[(x,y,z), (x,y,z), ...], [(x,y,z), (x,y,z), ...], ...]}

    即：把所有组里的“单根桩”汇总到一个大组内，供 CAD 绘图使用。
    """
    merged = {1: []}  # 合并后只有一个组编号
    for gid, pile_list in brake_pile_group_nodes.items():
        for single_pile in pile_list:
            merged[1].append(single_pile)
    return merged


# 根据选择的截面在cad中插入对应的块，插入点可以自己选择
def insert_block_incad_with_type_name(msp, insert_point, Type, name_str, Scale_Factor = [1,1,1], RO_Angle = 0):
    # 比例因子（x, y, z）
    # Scale_Factor = [1, 1, 1]
    # 旋转角度（以度为单位）
    # RO_Angle = 0
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


def int_or_float(Dividend, Divisor, Accuracy):
    if Dividend % Divisor != 0:
        return round(Dividend / Divisor, Accuracy)
    else:
        return int(Dividend / Divisor)


def SEC_trans(SEC_CAD, SEC_NAME):
    '''
    SEC_CAD: 截面信息表
    SEC_NAME: 截面名称
    '''
    H = SEC_CAD[SEC_NAME]['H'] # 截面高度
    B = SEC_CAD[SEC_NAME]['B'] # 截面宽度
    tf = SEC_CAD[SEC_NAME]['tf'] # 翼缘高度
    if SEC_NAME == '工字钢截面':
        format_str = int_or_float(H, 10, 1)
        if H >  180: Block_name = f'I{format_str}a'
        if H <= 180: Block_name = f'I{format_str}'
    elif SEC_NAME == '槽钢截面':
        format_str = int_or_float(H, 10, 1)
        if H >= 140: Block_name = f'C{format_str}a'
        if H <  140: Block_name = f'C{format_str}'
    elif SEC_NAME == 'HM截面' or SEC_NAME == 'HN截面':
        Block_name = f'{SEC_NAME.replace('截面', '')}{H}X{B}'
    elif SEC_NAME == '双拼工字钢截面':
        format_str = int_or_float(H, 10, 1)
        if H >  180: Block_name = f'2I{format_str}a'
        if H <= 180: Block_name = f'2I{format_str}'
    elif SEC_NAME == '双拼槽钢截面':
        format_str = int_or_float(H, 10, 1)
        if H >= 140: Block_name = f'2C{format_str}a'
        if H <  140: Block_name = f'2C{format_str}'
    elif SEC_NAME == '双拼HM截面' or SEC_NAME == '双拼HN截面':
        C = SEC_CAD[SEC_NAME]['C'] 
        format_str_1 = int(H)
        format_str_2 = int(C)
        Block_name = f'2{SEC_NAME[2:4]}{format_str_1}X{format_str_2}'
        print(Block_name)
    return Block_name, H, tf


# 插入圆（使用圆心和半径）
def make_circle(mspace, center, radius, layer):
    """
    mspace: 模型空间
    center: 圆心坐标
    radius: 半径
    layer: 图层名称
    """
    # 绘制圆
    circle = mspace.AddCircle(vtpnt(center), radius)
    
    # 设置圆属性
    try:
        circle.Layer = layer
    except:
        circle.Layer = '0'
    
    return circle  # 返回Obj




def draw_Trestle_LM_incad(aMSP, CAD_O_point, XY_range_dict, Global_Scale, if_ZD, xl_section_weight_per_meter, 
                          xl_single_nodes, xl_section_cad, xl_section_name, 
                          bailey_single_nodes, 
                          fenpei_single_nodes, distribute_section_cad, distribute_section_name, 
                          ganggz_single_nodes, gz_section_cad, gz_section_name,
                          abutmt_fenpei1_single_nodes, zd_f1_section_cad, zd_f1_section_name, 
                          abutmt_fenpei2_single_nodes, zd_f2_section_cad, zd_f2_section_name, 
                          abutmt_ganggz_single_nodes, zd_section_cad, zd_section_name, 
                          LM_Annotation_Dimension_Xdict, LM_Annotation_Dimension_Ydict, LM_Annotation_Mleader_Potdict, Material_dict, 
                         ):
    LM_Xlst, LM_Ylst = [], [] # 立面图中的XY的min,max值. 先将所有可能的值写进去, 最后从小到大排列取首尾
    O_point = CAD_O_point
    if_bz_ganggz = True # 默认为标准钢管桩, 控制标注的位置
    if_bz_abutmt_ganggz = True # 默认为标准钢管桩, 控制标注的位置  
    # 小肋
    rib_xlst = [value[0][1][0] for value in xl_single_nodes.values()] # 小肋的x坐标
    rib_length = abs(xl_single_nodes[1][0][1][1]-xl_single_nodes[1][-1][1][1])  # 小肋的长度
    Rib_XL_section, Rib_H, Rib_tf  = SEC_trans(xl_section_cad, xl_section_name)
    for x in rib_xlst:
        rib_insert_point = [O_point[0]+x, O_point[1], 0]
        insert_block_incad_with_type_name(aMSP, vtpnt(rib_insert_point), '小肋', Rib_XL_section)
    Material_dict['小肋']['数量'].append(len(rib_xlst)) # 小肋的数量
    Material_dict['小肋']['单重'].append(round(rib_length*xl_section_weight_per_meter/1000,2)) # 小肋的单重
    Material_dict['小肋']['总重'].append(round(rib_length*xl_section_weight_per_meter/1000*len(rib_xlst),1)) # 小肋的总重  
    # 桥面板
    deck_thickness = 10
    deck_x1, deck_x2 = O_point[0] + rib_xlst[0], O_point[0] + rib_xlst[-1]
    deck_y1, deck_y2 = O_point[1] + Rib_H, O_point[1]+Rib_H+deck_thickness
    deck_ptlst = [(deck_x1, deck_y1, 0), (deck_x2, deck_y1, 0), (deck_x2, deck_y2, 0), (deck_x1, deck_y2,  0), (deck_x1, deck_y1, 0)]
    flat_points = pointlist_extend(deck_ptlst)
    make_polyline(aMSP, flat_points, '6标注线')
    LM_Annotation_Mleader_Potdict['桥面板'].append([(deck_x2, deck_y1, 0), '桥面板', (1, 1)]) # 引线点, 文本内容, 标注方向(X和Y方向, 1表示正, -1表示负)
    Material_dict['桥面板']['规格'].append(f'□{int(abs(deck_x2-deck_x1))}×{int(deck_thickness)}×{int(rib_length)}') #   桥面板的长度
    deck_weight = round((abs(int(deck_x2-deck_x1))/1000)*(int(deck_thickness)/1000)*(int(rib_length)/1000)*7.85*1000, 2)
    Material_dict['桥面板']['单重'].append(deck_weight) # 桥面板的单重
    Material_dict['桥面板']['总重'].append(round(deck_weight,1)) # 桥面板的总重
    # 贝雷片    
    bailey_insertpointy = O_point[1]-50 # 贝雷片插入点的y坐标
    bailey_insertpointx = [pt[0] for pt in bailey_single_nodes[1]]
    for x in bailey_insertpointx:
        bailey_insert_point = [O_point[0]+x, bailey_insertpointy, 0]
        insert_block_incad_with_type_name(aMSP, vtpnt(bailey_insert_point), '贝雷正立面', '3')
    LM_Annotation_Mleader_Potdict['贝雷'].append([(O_point[0] + bailey_insertpointx[-1] + 3000, bailey_insertpointy - 750, 0), '贝雷梁', (1, 1)]) # 引线点, 文本内容
    # 分配梁
    Fenpei_section, fenpei_H, fenpei_tf = SEC_trans(distribute_section_cad, distribute_section_name)
    fenpei_xlst = [value[0][1][0] for value in fenpei_single_nodes.values()] # 分配梁的x坐标
    fenpei_insertpointy = bailey_insertpointy-1450-fenpei_H # 分配梁插入点的y坐标
    for x in fenpei_xlst:
        fenpei_insert_point = [O_point[0]+x, fenpei_insertpointy, 0]
        insert_block_incad_with_type_name(aMSP, vtpnt(fenpei_insert_point), '分配梁立面', Fenpei_section)
        LM_Annotation_Mleader_Potdict['分配梁'].append([fenpei_insert_point, '分配梁F1', (1, -1)]) # 引线点, 文本内容
    Material_dict['分配梁']['名称'].append('分配梁F1')
    Material_dict['分配梁']['备注'].append('')
    # 钢管桩
    ganggz_xlst = fenpei_xlst # 非制动墩的钢管桩x坐标与分配梁一致
    ganggz_insertpointy = fenpei_insertpointy
    ganggz_length = abs(ganggz_single_nodes[1][0][0][2] - ganggz_single_nodes[1][0][-1][2]) # 钢管桩桩长
    ganggz_D, ganggz_d = gz_section_cad[gz_section_name]['D'], gz_section_cad[gz_section_name]['d'] # 直径和壁厚
    for x in ganggz_xlst:
        try:
            insert_block_incad_with_block_name(aMSP, vtpnt([O_point[0]+x, ganggz_insertpointy, 0]), f'φ{ganggz_D}开槽柱头截断线钢管桩2block', f'EX_φ{ganggz_D}')
            LM_Annotation_Mleader_Potdict['钢管桩'].append([(O_point[0]+x, ganggz_insertpointy-4500, 0), '钢管桩G1', (1, -1)]) # 引线点, 文本内容, 4500是钢管桩图块长7000, 固定值
        except:
            print('未匹配非制动墩钢管桩直径, 生成钢管桩简图')
            ganggz_ptlst1 = [[O_point[0]+x+nx*ganggz_D/2, ganggz_insertpointy+ny*ganggz_length, 0] for nx, ny in [(-1,0),(-1,-1),(1,-1),(1,0),(-1,0)]]
            flat_points = pointlist_extend(ganggz_ptlst1)
            make_polyline(aMSP, flat_points, '1粗实线')
            ganggz_ptlst2 = [[O_point[0]+x+nx*ganggz_D/2-ganggz_d, ganggz_insertpointy+ny*ganggz_length, 0] for nx, ny in [(-1,0),(-1,-1),(1,-1),(1,0),(-1,0)]]
            flat_points = pointlist_extend(ganggz_ptlst2)
            make_polyline(aMSP, flat_points, '2细实线')
            make_line(aMSP, [(O_point[0]+x, ganggz_insertpointy, 0), (O_point[0]+x, ganggz_insertpointy-ganggz_length, 0)], '3中心线')
            LM_Annotation_Mleader_Potdict['钢管桩'].append([(O_point[0]+x, ganggz_insertpointy-ganggz_length/2, 0), '钢管桩G1', (1, -1)]) # 引线点, 文本内容
            if_bz_ganggz = False
    Material_dict['钢管桩']['名称'].append('钢管桩G1')
    Material_dict['钢管桩']['单重'].append(round((math.pi*ganggz_D**2/4-math.pi*(ganggz_D-2*ganggz_d)**2/4)/1000000*7.85*1000*ganggz_length/1000, 2))
    Material_dict['钢管桩']['备注'].append('')
    # 制动墩分配梁1
    abutmt_fenpei1_xlst = []
    if if_ZD == True:
        abutmt_fenpei1_section, abutmt_fenpei1_H, abutmt_fenpei1_tf = SEC_trans(zd_f1_section_cad, zd_f1_section_name) # 截面名称、截面高度、翼缘厚度
        for value in abutmt_fenpei1_single_nodes.values():
            abutmt_fenpei1_xlst.append(value[0][1][0]) # x坐标
        abutmt_fenpei1_xlst = list(set(abutmt_fenpei1_xlst))
        abutmt_fenpei1_insertpointy = bailey_insertpointy-1450-abutmt_fenpei1_H
        for x in abutmt_fenpei1_xlst:
            abutmt_fenpei1_insert_point = [O_point[0]+x, abutmt_fenpei1_insertpointy, 0]
            insert_block_incad_with_type_name(aMSP, vtpnt(abutmt_fenpei1_insert_point), '分配梁立面', abutmt_fenpei1_section)
            LM_Annotation_Mleader_Potdict['分配梁1'].append([abutmt_fenpei1_insert_point, '分配梁F2', (1, 1)]) # 引线点, 文本内容
        Material_dict['分配梁']['名称'].append('分配梁F2')
        Material_dict['分配梁']['备注'].append('')
    # 制动墩分配梁2
    abutmt_fenpei2_xlst = []
    if if_ZD == True:
        abutmt_fenpei2_section, abutmt_fenpei2_H, abutmt_fenpei2_tf = SEC_trans(zd_f2_section_cad, zd_f2_section_name) # 截面名称和截面高度
        for value in abutmt_fenpei2_single_nodes.values():
            abutmt_fenpei2_xlst.append((value[0][1][0], value[-1][1][0]))
        abutmt_fenpei2_xlst = list(set(abutmt_fenpei2_xlst))
        abutmt_fenpei2_insertpointy = abutmt_fenpei1_insertpointy
        for xlst in abutmt_fenpei2_xlst:
            abutmt_fenpei2_ptlst1 = [(O_point[0]+xlst[0], abutmt_fenpei2_insertpointy, 0),
                                     (O_point[0]+xlst[1], abutmt_fenpei2_insertpointy, 0),
                                     (O_point[0]+xlst[1], abutmt_fenpei2_insertpointy-abutmt_fenpei2_H, 0),
                                     (O_point[0]+xlst[0], abutmt_fenpei2_insertpointy-abutmt_fenpei2_H, 0),
                                     (O_point[0]+xlst[0], abutmt_fenpei2_insertpointy, 0)
                                    ]
            flat_points = pointlist_extend(abutmt_fenpei2_ptlst1)
            make_polyline(aMSP, flat_points, '1粗实线')
            abutmt_fenpei2_ptlst2 = [(O_point[0]+xlst[0], abutmt_fenpei2_insertpointy-abutmt_fenpei2_tf, 0),
                                     (O_point[0]+xlst[1], abutmt_fenpei2_insertpointy-abutmt_fenpei2_tf, 0),
                                     (O_point[0]+xlst[1], abutmt_fenpei2_insertpointy-abutmt_fenpei2_H+abutmt_fenpei2_tf, 0),
                                     (O_point[0]+xlst[0], abutmt_fenpei2_insertpointy-abutmt_fenpei2_H+abutmt_fenpei2_tf, 0),
                                     (O_point[0]+xlst[0], abutmt_fenpei2_insertpointy-abutmt_fenpei2_tf, 0)
                                    ]
            flat_points = pointlist_extend(abutmt_fenpei2_ptlst2)
            make_polyline(aMSP, flat_points, '2细实线')
            LM_Annotation_Mleader_Potdict['分配梁2'].append([((xlst[0]+xlst[1])/2+O_point[0], abutmt_fenpei2_insertpointy-abutmt_fenpei2_H, 0), '分配梁F3', (1, -1)]) # 引线点, 文本内容
        Material_dict['分配梁']['名称'].append('分配梁F3')
        Material_dict['分配梁']['备注'].append('')
    # 制动墩钢管桩
    abutmt_ganggz_xlst = []
    if if_ZD == True:
        for value in abutmt_ganggz_single_nodes.values():
            abutmt_ganggz_xlst.append(sorted(list(set(tuple([ptlst[0][0] for ptlst in value])))))
        abutmt_ganggz_length = abs(abutmt_ganggz_single_nodes[1][0][0][2] - abutmt_ganggz_single_nodes[1][0][-1][2]) # 制动墩钢管桩桩长
        abutmt_ganggz_D, abutmt_ganggz_d = zd_section_cad[zd_section_name]['D'], zd_section_cad[zd_section_name]['d'] # 直径和壁厚
        abutmt_ganggz_insertpointy = abutmt_fenpei2_insertpointy-abutmt_fenpei2_H # 钢管桩插入点Y坐标
        for xlst in abutmt_ganggz_xlst:
            for x in xlst:
                try:
                    insert_block_incad_with_block_name(aMSP, vtpnt([O_point[0]+x, abutmt_ganggz_insertpointy, 0]), f'φ{ganggz_D}开槽柱头截断线钢管桩1block', f'EX_φ{ganggz_D}')
                    LM_Annotation_Mleader_Potdict['制动桩'].append([(O_point[0]+x, ganggz_insertpointy-4500, 0), '钢管桩G2', (1, -1)]) # 引线点, 文本内容, 4500是钢管桩图块长7000, 固定值
                except:
                    print('未匹配制动墩钢管桩直径, 生成钢管桩简图')
                    abutmt_ganggz_ptlst1 = [[O_point[0]+x+nx*abutmt_ganggz_D/2, abutmt_ganggz_insertpointy+ny*abutmt_ganggz_length, 0] for nx, ny in [(-1,0),(-1,-1),(1,-1),(1,0),(-1,0)]]
                    flat_points = pointlist_extend(abutmt_ganggz_ptlst1)
                    make_polyline(aMSP, flat_points, '1粗实线')
                    abutmt_ganggz_ptlst2 = [[O_point[0]+x+nx*abutmt_ganggz_D/2-abutmt_ganggz_d, abutmt_ganggz_insertpointy+ny*abutmt_ganggz_length, 0] for nx, ny in [(-1,0),(-1,-1),(1,-1),(1,0),(-1,0)]]
                    flat_points = pointlist_extend(abutmt_ganggz_ptlst2)
                    make_polyline(aMSP, flat_points, '2细实线')
                    make_line(aMSP, [(O_point[0]+x, abutmt_ganggz_insertpointy, 0), (O_point[0]+x, abutmt_ganggz_insertpointy-abutmt_ganggz_length, 0)], '3中心线')
                    LM_Annotation_Mleader_Potdict['制动桩'].append([(O_point[0]+x, ganggz_insertpointy-ganggz_length/2, 0), '钢管桩G2', (1, -1)]) # 引线点, 文本内容
            # 制动墩联结系
            ljx_y1, ljx_y2 = abutmt_ganggz_insertpointy-1000, abutmt_ganggz_insertpointy-3000
            ljx_pt1 = (O_point[0]+xlst[0], ljx_y1, 0)
            ljx_pt2 = (O_point[0]+xlst[1], ljx_y1, 0)
            ljx_pt3 = (O_point[0]+xlst[0], ljx_y2, 0)
            ljx_pt4 = (O_point[0]+xlst[1], ljx_y2, 0)
            try:
                ljx_length = int(abs(ljx_pt1[0] - ljx_pt2[0])/1000)
                insert_block_incad_with_block_name(aMSP, vtpnt([(ljx_pt1[0]+ljx_pt2[0])/2, abutmt_ganggz_insertpointy-2000, 0]), f'φ{abutmt_ganggz_D}型钢{ljx_length}m联结系block', f'EX_φ{abutmt_ganggz_D}')
            except:
                print('未匹配已有联结系长度, 生成联结系简图')
                make_line(aMSP, [ljx_pt1, ljx_pt2], '3中心线')
                make_line(aMSP, [ljx_pt3, ljx_pt4], '3中心线')
                make_line(aMSP, [ljx_pt1, ljx_pt4], '3中心线')
                make_line(aMSP, [ljx_pt2, ljx_pt3], '3中心线')
        Material_dict['钢管桩']['名称'].append('钢管桩G2')  
        Material_dict['钢管桩']['单重'].append(round((math.pi*abutmt_ganggz_D**2/4-math.pi*(abutmt_ganggz_D-2*abutmt_ganggz_d)**2/4)/1000000*7.85*1000*abutmt_ganggz_length/1000, 2))
        Material_dict['钢管桩']['备注'].append('')

    # 常规结构标注点的XY坐标
    bailey_X1 = O_point[0] + bailey_insertpointx[0] # 贝雷起点
    bailey_X2 = O_point[0] + bailey_insertpointx[-1] + 3000 # 贝雷终点
    ganggz_Xlst = sorted([O_point[0] + ptx for ptx in [x for xlst in abutmt_ganggz_xlst for x in xlst] + ganggz_xlst]) # 非制动墩和制动墩的全部X坐标
    deck_rib_ylst = [O_point[1], deck_y2] # 小肋与桥面板在CAD中的y坐标
    bailey_ylst = [O_point[1], O_point[1]-1500] # 贝雷在CAD中的y坐标
    fenpei_ylst = [fenpei_insertpointy, fenpei_insertpointy+fenpei_H] # 分配梁在CAD中的y坐标
    ganggz_ylst = [ganggz_insertpointy, ganggz_insertpointy-ganggz_length] if not if_bz_ganggz else [ganggz_insertpointy, ganggz_insertpointy-7000]  # 钢管桩在CAD中的y坐标
    # 标注信息
    LM_Annotation_Dimension_Xdict['贝雷'].extend([bailey_X1, bailey_X2])
    LM_Annotation_Dimension_Xdict['钢管桩'].extend(ganggz_Xlst)
    LM_Annotation_Dimension_Ydict['桥面板'].extend(deck_rib_ylst)
    LM_Annotation_Dimension_Ydict['贝雷'].extend(bailey_ylst)
    LM_Annotation_Dimension_Ydict['分配梁'].extend(fenpei_ylst)
    LM_Annotation_Dimension_Ydict['钢管桩'].extend(ganggz_ylst)
    # 可能没有的结构的标注点的XY坐标
    if if_ZD == True:
        abutmt_fenpei1_ylst = [abutmt_fenpei1_insertpointy, abutmt_fenpei1_insertpointy+abutmt_fenpei1_H]
        LM_Annotation_Dimension_Ydict['分配梁1'].extend(abutmt_fenpei1_ylst)
        abutmt_fenpei2_ylst = [abutmt_fenpei2_insertpointy, abutmt_fenpei2_insertpointy-abutmt_fenpei2_H]
        LM_Annotation_Dimension_Ydict['分配梁2'].extend(abutmt_fenpei2_ylst)
        abutmt_ganggz_ylst = [abutmt_ganggz_insertpointy, abutmt_ganggz_insertpointy-abutmt_ganggz_length] if not if_bz_abutmt_ganggz else [abutmt_ganggz_insertpointy, abutmt_ganggz_insertpointy-7000]
        LM_Annotation_Dimension_Ydict['制动桩'].extend(abutmt_ganggz_ylst)
        abutmt_ljx_ylst = [ljx_y1, ljx_y2]
        LM_Annotation_Dimension_Ydict['联结系'].extend(abutmt_ljx_ylst)
    # 图形边界处理
    LM_Xlst.append(O_point[0]) # 小肋X
    LM_Xlst.append(bailey_X1)  # 第一片贝雷插入点X
    LM_Xlst.append(bailey_X2) # 最后一片贝雷插入点X + 3000mm
    LM_Ylst.append(O_point[1] + Rib_H) # 小肋Y+小肋高
    LM_Ylst.append(ganggz_ylst[1]) # 非制动墩桩底
    if if_ZD == True:
        LM_Xlst.extend([O_point[0] + x for xlst in abutmt_fenpei2_xlst for x in xlst]) # 制动墩分配梁2X
        LM_Ylst.append(abutmt_ganggz_ylst[1]) # 制动墩桩底
    LM_Xlst = sorted(LM_Xlst)
    LM_Ylst = sorted(LM_Ylst)
    XY_range_dict['立面图']['X'].extend([round(LM_Xlst[0],1), round(LM_Xlst[-1],1)])
    XY_range_dict['立面图']['Y'].extend([round(LM_Ylst[0],1), round(LM_Ylst[-1],1)])
    print(XY_range_dict)
    


def draw_Trestle_LM_Annotation_incad(aCAD, aMSP, Global_Scale, XY_range_dict, if_ZD, ganggz_single_nodes, abutmt_ganggz_single_nodes, 
                                     LM_Annotation_Dimension_Xdict, LM_Annotation_Dimension_Ydict, LM_Annotation_Mleader_Potdict,
                                    ):
    # 尺寸标注
    Y1_keylst = ['桥面板', '贝雷', '分配梁', '钢管桩']
    Y2_keylst = ['桥面板', '贝雷', '分配梁1', '分配梁2', '制动桩', '联结系']
    LM_Annotation_Dimension_X1 = sorted(set([x for x in LM_Annotation_Dimension_Xdict['贝雷']]))
    LM_Annotation_Dimension_X2 = sorted(set([x for x in LM_Annotation_Dimension_Xdict['钢管桩']])) 
    LM_Annotation_Dimension_Y1 = sorted(set([y for key,value in LM_Annotation_Dimension_Ydict.items() if key in Y1_keylst for y in value]), reverse = True)
    LM_Annotation_Dimension_Y2 = sorted(set([y for key,value in LM_Annotation_Dimension_Ydict.items() if key in Y2_keylst for y in value]), reverse = True)
    Dimension_O_X1, Dimension_O_X2 = XY_range_dict['立面图']['X'] # 标注的定位X
    Dimension_O_Y1, Dimension_O_Y2 = XY_range_dict['立面图']['Y'] # 标注的定位Y
    ganggz_length = abs(ganggz_single_nodes[1][0][0][2] - ganggz_single_nodes[1][0][-1][2]) # 钢管桩桩长
    # 上标注
    for i in range(len(LM_Annotation_Dimension_X1)-1):
        X1 = LM_Annotation_Dimension_X1[i]
        X2 = LM_Annotation_Dimension_X1[i+1]
        Y0 = Dimension_O_Y2
        Add_Annotation_Linear(aMSP, (X1,Y0,0), (X2,Y0,0), (X1,Y0+6*Global_Scale,0), 0) # 两点之间标注
    # 下标注
    for i in range(len(LM_Annotation_Dimension_X2)-1):
        X1 = LM_Annotation_Dimension_X2[i]
        X2 = LM_Annotation_Dimension_X2[i+1]
        Y0 = Dimension_O_Y1
        Add_Annotation_Linear(aMSP, (X1,Y0,0), (X2,Y0,0), (X1,Y0-6*Global_Scale,0), 0) # 两点之间标注
    # 下总标注
    X1 = LM_Annotation_Dimension_X2[0]
    X2 = LM_Annotation_Dimension_X2[-1]
    Y0 = Dimension_O_Y1
    Add_Annotation_Linear(aMSP, (X1,Y0,0), (X2,Y0,0), (X1,Y0-12*Global_Scale,0), 0) # 两点之间标注
    # 左标注
    for i in range(len(LM_Annotation_Dimension_Y1)-1):
        Y1 = LM_Annotation_Dimension_Y1[i]
        Y2 = LM_Annotation_Dimension_Y1[i+1]
        X0 = Dimension_O_X1-5*Global_Scale
        if i != len(LM_Annotation_Dimension_Y1)-2:
            Add_Annotation_Linear(aMSP, (X0,Y1,0), (X0,Y2,0), (X0-6*Global_Scale,Y1,0), 90) # 两点之间标注
        else:
            Annotation_Linear_txt = ganggz_length
            Add_Annotation_Linear(aMSP, (X0,Y1,0), (X0,Y2,0), (X0-6*Global_Scale,Y1,0), 90, '6标注线', Annotation_Linear_txt) # 钢管桩最后一个标注
    # 总左标注
    Y1 = LM_Annotation_Dimension_Y1[0]
    Y2 = LM_Annotation_Dimension_Y1[-1]
    X0 = Dimension_O_X1-5*Global_Scale
    Annotation_Linear_txt = abs(Y1-Y2) + (ganggz_length - 7000)
    Add_Annotation_Linear(aMSP, (X0,Y1,0), (X0,Y2,0), (X0-12*Global_Scale,Y1,0), 90, '6标注线', Annotation_Linear_txt) # 两点之间标注
    if if_ZD == True:
        abutmt_ganggz_length = abs(abutmt_ganggz_single_nodes[1][0][0][2] - abutmt_ganggz_single_nodes[1][0][-1][2]) # 制动墩钢管桩桩长
        # 右标注
        for i in range(len(LM_Annotation_Dimension_Y2)-1):
            Y1 = LM_Annotation_Dimension_Y2[i]
            Y2 = LM_Annotation_Dimension_Y2[i+1]
            X0 = Dimension_O_X2+5*Global_Scale
            if i != len(LM_Annotation_Dimension_Y2)-2:
                Add_Annotation_Linear(aMSP, (X0,Y1,0), (X0,Y2,0), (X0+6*Global_Scale,Y1,0), 90) # 两点之间标注
            else:
                Annotation_Linear_txt = abutmt_ganggz_length - 3000
                Add_Annotation_Linear(aMSP, (X0,Y1,0), (X0,Y2,0), (X0+6*Global_Scale,Y1,0), 90, '6标注线', Annotation_Linear_txt) # 钢管桩最后一个标注
        # 总右标注
        Y1 = LM_Annotation_Dimension_Y2[0]
        Y2 = LM_Annotation_Dimension_Y2[-1]
        X0 = Dimension_O_X2+5*Global_Scale
        Annotation_Linear_txt = abs(Y1-Y2) + (abutmt_ganggz_length - 7000)
        Add_Annotation_Linear(aMSP, (X0,Y1,0), (X0,Y2,0), (X0+12*Global_Scale,Y1,0), 90, '6标注线', Annotation_Linear_txt) # 两点之间标注
    # 引线标注
    for value in LM_Annotation_Mleader_Potdict.values():
        for lst in value:
            startpoint = lst[0]
            txt = lst[1]
            nx, ny = lst[2]
            endpoint = [startpoint[0]+3*nx*Global_Scale, startpoint[1]+6*ny*Global_Scale, startpoint[2]]
            Add_Annotation_MLeader(aCAD, aMSP, Global_Scale, startpoint, endpoint, '5文本', 'Song_07', txt)
    # 标题
    Header_Point = [(LM_Annotation_Dimension_X1[0]+LM_Annotation_Dimension_X1[-1])/2, Dimension_O_Y2+24*Global_Scale, 0]
    Add_Header(aMSP, Global_Scale, Header_Point, '栈桥立面图', 'Song_07')


# 生成材料表
def draw_Material_List_in_cad(aMSP, CAD_O_point, Global_Scale, section_weight_per_meter_lst, 
                              XY_range_dict, PM_Annotation_Dimension_Xdict, PM_Annotation_Dimension_Ydict, Material_dict,
                             ):
    '''
    编号; 名称; 规格; 材质/牌号; 数量; 单重; 总重; 备注
    '''
    # 表格宽度
    Serial_Number_Width = 10*Global_Scale # 编号宽度
    Name_Width = 20*Global_Scale # 名称宽度
    Specifications_Width = 35*Global_Scale # 规格宽度
    Material_Width = 20*Global_Scale # 材质宽度
    Quantity_Width = 10*Global_Scale # 数量宽度
    Singlet_Width = 15*Global_Scale # 单重宽度
    Total_Weight_Width = 15*Global_Scale # 总重宽度
    Note_Width = 25*Global_Scale # 备注宽度
    # 生成材料表的线
    Width_lst = [Serial_Number_Width, Name_Width, Specifications_Width, Material_Width, Quantity_Width, Singlet_Width, Total_Weight_Width, Note_Width]
    Width  = sum(Width_lst) # 总宽
    row = len(Width_lst) # 总列数
    line = len([name for item in Material_dict.values() for name in item['名称']]) # 构件行数
    Height = (line+2) * 6 * Global_Scale
    firstline_txt_insert_X = [] # 第一行txt的X插入点
    firstline_txt_insert_Y = [] # 第一行txt的Y插入点
    txt_insert_X = [] # txt的x坐标插入点
    txt_insert_Y = [] # txt的Y坐标插入点
    # 定位点计算
    position_line_xlst = XY_range_dict['立面图']['X'] # 边线X坐标
    position_line_ylst = XY_range_dict['平面图']['Y'] # 边线Y坐标
    X = round(abs(position_line_xlst[0]-position_line_xlst[1]),1)
    O_point = [CAD_O_point[0]+36*Global_Scale+X+Width/2, (position_line_ylst[0]+position_line_ylst[1])/2, CAD_O_point[2]] # 定位原点
    # 标题
    Header_Point = [O_point[0], O_point[1]+Height/2+10*Global_Scale, 0]
    Add_Header(aMSP, Global_Scale, Header_Point, '材料表', 'Song_07')
    # 外轮廓线
    ptlst  = [(O_point[0]+nx*Width/2, O_point[1]+ny*Height/2, 0)  for nx,ny in [(1,1),(-1,1),(-1,-1),(1,-1),(1,1)]]
    flat_points = pointlist_extend(ptlst)
    make_polyline(aMSP, flat_points, '1粗实线', 0.4*Global_Scale) # 外轮廓线
    # 右上角文字
    Add_Text(aMSP, (ptlst[0][0], ptlst[0][1]+1.5*Global_Scale, ptlst[0][2]), '以单个栈桥计', 2.5*Global_Scale, 'Song_07', '5文本', 14)
    # 细实线
    ptstartX, ptendX = O_point[0]-Width/2 , O_point[0]+Width/2  # 从左往右
    ptstartY, ptendY = O_point[1]+Height/2, O_point[1]-Height/2 # 从上往下
    for i in range(line+2):
        ptY = ptstartY-(i+1)*6*Global_Scale
        if i != line+1:
            pt1 = (ptstartX, ptY, 0)
            pt2 = (ptendX  , ptY, 0)
            make_line(aMSP, [pt1, pt2], '2细实线')
        txt_insert_Y.extend([ptY+3*Global_Scale])
        firstline_txt_insert_Y.extend([ptY+3*Global_Scale])
    ptX = ptstartX # 初始量
    for i in range(row):
        if i != 0:
            pt1 = (ptX, ptstartY, 0)
            pt2 = (ptX, ptendY+6*Global_Scale, 0)
            make_line(aMSP, [pt1, pt2], '2细实线')
        if i == 2:
            txt_insert_X.append(ptX + 2*Global_Scale)
        else:
            txt_insert_X.append(ptX + Width_lst[i]/2)
        firstline_txt_insert_X.append(ptX + Width_lst[i]/2)
        ptX = ptX + Width_lst[i]
    # 插入文字
    txt_dict = {0:['编号', '名称', '规    格', '材质/牌号', '数量', '单重(kg)', '总重(kg)', '备注']}
    txt_line = 1
    print(Material_dict)
    for key, value in Material_dict.items():
        namelst = value['名称']
        if namelst != []:
            for i in range(len(value['名称'])):
                txt_dict[txt_line] = [txt_line, value['名称'][i], value['规格'][i], value['材质'][i], value['数量'][i], value['单重'][i], value['总重'][i], value['备注'][i]]
                txt_line += 1
    print(txt_dict)
    # 除汇总行外的所有行
    for iy in range(len(txt_insert_Y)-1):
        for ix in range(len(txt_insert_X)):
            txt = txt_dict[iy][ix]
            if iy == 0:
                pt = (firstline_txt_insert_X[ix],firstline_txt_insert_Y[iy],0)
            else:
                pt = (txt_insert_X[ix],txt_insert_Y[iy],0)
            if ix == 2 and iy > 0: AlignNum = 9 # 规格列的对齐方式为左对齐
            else: AlignNum = 4 # 其余列的对齐方式为居中对齐
            Add_Text(aMSP, pt, txt, 2.5*Global_Scale, 'Song_07', '5文本', AlignNum)
    # 汇总行
    Steel_all_weight = sum([value[6] for key, value in txt_dict.items() if key != 0]) # 钢材总重
    all_weight_line_txt = '合计钢材 {} kg。'.format(Steel_all_weight)
    insert_pt = (ptstartX+sum(Width_lst[:-1]), O_point[1]-Height/2+3*Global_Scale, 0)
    Add_Text(aMSP, insert_pt, all_weight_line_txt, 2.5*Global_Scale, 'Song_07', '5文本', 11)
    
    # for key, value in Material_dict.items():
    #     print(key)
    #     print(value)

def draw_Trestle_PM_incad(aMSP, CAD_O_point, XY_range_dict, Global_Scale, if_ZD, distribute_section_per_meter, zd_f1_section_weight_per_meter, zd_f2_section_weight_per_meter, 
                          xl_single_nodes, xl_section_cad, xl_section_name, 
                          bailey_single_nodes, 
                          fenpei_single_nodes, distribute_section_cad, distribute_section_name, 
                          ganggz_single_nodes, gz_section_cad, gz_section_name,
                          abutmt_fenpei1_single_nodes, zd_f1_section_cad, zd_f1_section_name, 
                          abutmt_fenpei2_single_nodes, zd_f2_section_cad, zd_f2_section_name, 
                          abutmt_ganggz_single_nodes, zd_section_cad, zd_section_name, 
                          PM_Annotation_Dimension_Xdict, PM_Annotation_Dimension_Ydict, PM_Annotation_Mleader_Potdict, Material_dict, 
                         ):
    GAP_Y = 46*Global_Scale # 平面图与立面图之间的Y间距  
    # 小肋
    rib_length = abs(xl_single_nodes[1][0][1][1]-xl_single_nodes[1][-1][1][1])  # 小肋的长度
    O_point = (CAD_O_point[0], XY_range_dict['立面图']['Y'][0]-GAP_Y-rib_length, CAD_O_point[2]) # 平面图中的定位点
    # 贝雷梁
    basic_x, basic_y, basic_z = xl_single_nodes[1][0][1] # 小肋节点y坐标在midas中的第一个点
    bailey_insertpoint_lst = [pt for value in bailey_single_nodes.values() for pt in value] # 所有贝雷片的起点
    for pt in bailey_insertpoint_lst:
        insertpt = (O_point[0]+pt[0], O_point[1]+pt[1]-basic_y, 0)
        PM_Annotation_Dimension_Ydict['贝雷'].extend([O_point[1]+pt[1]-basic_y]) # 贝雷Y标注
        insert_block_incad_with_type_name(aMSP, vtpnt(insertpt), '贝雷平面', '3')
    bailey_mleader_insertpointx = (bailey_insertpoint_lst[0][0] + bailey_insertpoint_lst[-1][0])/2 + O_point[0] + 1500
    bailey_mleader_insertpointy = min([pt[1] for pt in bailey_insertpoint_lst])+O_point[1]-basic_y
    PM_Annotation_Mleader_Potdict['贝雷'].append([(bailey_mleader_insertpointx, bailey_mleader_insertpointy, 0), '贝雷梁', (1, -1)])
    Material_dict['贝雷']['数量'].append(len(bailey_insertpoint_lst))
    Material_dict['贝雷']['总重'].append(round(len(bailey_insertpoint_lst)*270,1))
    # 分配梁
    fenpei_insertpoint_lst = [[value[0][1], value[-1][1]] for value in fenpei_single_nodes.values()] # 所有分配梁的首尾端点
    Fenpei_section, fenpei_H, fenpei_tf = SEC_trans(distribute_section_cad, distribute_section_name)
    for ptlst in fenpei_insertpoint_lst:
        pt1, pt2 = ptlst # 每根分配梁的端点
        fenpei_length = math.sqrt((pt1[0]-pt2[0])**2+(pt1[1]-pt2[1])**2) # 每根分配梁的长度
        scale_y = fenpei_length/5000 # 分配梁在平面图中y方向上拉伸的比例, 图块的默认长度为5000
        insertpt = (O_point[0]+pt2[0], O_point[1]+pt2[1]-basic_y, 0)
        PM_Annotation_Dimension_Xdict['分配梁'].extend([O_point[0]+pt2[0]]) # 分配梁X标注
        PM_Annotation_Dimension_Ydict['分配梁'].extend([O_point[1]+pt1[1]-basic_y, O_point[1]+pt2[1]-basic_y]) # 分配梁Y标注
        insert_block_incad_with_type_name(aMSP, vtpnt(insertpt), '分配梁平面', Fenpei_section, [1,scale_y,1], 0)
        PM_Annotation_Mleader_Potdict['分配梁'].append([insertpt, '分配梁F1', (1, 1)])
    Material_dict['分配梁']['规格'].append(f'{Fenpei_section}×{int(fenpei_length)}')
    Material_dict['分配梁']['材质'].append('Q235B')
    Material_dict['分配梁']['数量'].append(len(fenpei_insertpoint_lst))
    Material_dict['分配梁']['单重'].append(round(fenpei_length*distribute_section_per_meter/1000,2)) 
    Material_dict['分配梁']['总重'].append(round(fenpei_length*distribute_section_per_meter*len(fenpei_insertpoint_lst)/1000,1)) 
    # 钢管桩
    ganggz_insertpoint_lst = [ptlst[0] for value in ganggz_single_nodes.values() for ptlst in value]
    ganggz_D, ganggz_d = gz_section_cad[gz_section_name]['D'], gz_section_cad[gz_section_name]['d'] # 直径和壁厚
    for pt in ganggz_insertpoint_lst:
        insertpt = (O_point[0]+pt[0], O_point[1]+pt[1]-basic_y, 0)
        PM_Annotation_Dimension_Ydict['钢管桩'].extend([O_point[1]+pt[1]-basic_y]) # 钢管桩Y标注
        make_circle(aMSP, insertpt, ganggz_D/2, '1粗实线')
        make_circle(aMSP, insertpt, ganggz_D/2-ganggz_d, '4虚线')
    Material_dict['钢管桩']['数量'].append(len(ganggz_insertpoint_lst))
    Material_dict['钢管桩']['材质'].append('Q235B')
    # 制动墩分配梁  
    if if_ZD == True:
        abutmt_fenpei1_insertpoint_lst = [[value[0][1], value[-1][1]] for value in abutmt_fenpei1_single_nodes.values()] # 所有分配梁的首尾端点
        abutmt_Fenpei1_section, abutmt_fenpei1_H, abutmt_fenpei1_tf = SEC_trans(zd_f1_section_cad, zd_f1_section_name)
        for ptlst in abutmt_fenpei1_insertpoint_lst:
            pt1, pt2 = ptlst # 每根分配梁的端点
            abutmt_fenpei1_length = math.sqrt((pt1[0]-pt2[0])**2+(pt1[1]-pt2[1])**2) # 每根分配梁的长度
            scale_y = abutmt_fenpei1_length/5000 # 分配梁在平面图中y方向上拉伸的比例, 图块的默认长度为5000
            insertpt = (O_point[0]+pt2[0], O_point[1]+pt2[1]-basic_y, 0)
            PM_Annotation_Dimension_Xdict['分配梁1'].extend([O_point[0]+pt2[0]]) # 分配梁X标注
            PM_Annotation_Dimension_Ydict['分配梁1'].extend([O_point[1]+pt1[1]-basic_y, O_point[1]+pt2[1]-basic_y]) # 分配梁Y标注
            insert_block_incad_with_type_name(aMSP, vtpnt(insertpt), '分配梁平面', abutmt_Fenpei1_section, [1,scale_y,1], 0)
            PM_Annotation_Mleader_Potdict['分配梁1'].append([insertpt, '分配梁F2', (1, 1)])
        Material_dict['分配梁']['规格'].append(f'{abutmt_Fenpei1_section}×{int(abutmt_fenpei1_length)}')
        Material_dict['分配梁']['材质'].append('Q235B')
        Material_dict['分配梁']['数量'].append(len(abutmt_fenpei1_insertpoint_lst))
        Material_dict['分配梁']['单重'].append(round(abutmt_fenpei1_length*zd_f1_section_weight_per_meter/1000,2)) 
        Material_dict['分配梁']['总重'].append(round(abutmt_fenpei1_length*zd_f1_section_weight_per_meter*len(abutmt_fenpei1_insertpoint_lst)/1000,1)) 
        abutmt_fenpei2_insertpoint_lst = [[value[0][1], value[-1][1]] for value in abutmt_fenpei2_single_nodes.values()] # 所有分配梁的首尾端点
        abutmt_Fenpei2_section, abutmt_fenpei2_H, abutmt_fenpei2_tf = SEC_trans(zd_f2_section_cad, zd_f2_section_name)
        for ptlst in abutmt_fenpei2_insertpoint_lst:
            pt1, pt2 = ptlst # 每根分配梁的端点
            abutmt_fenpei2_length = math.sqrt((pt1[0]-pt2[0])**2+(pt1[1]-pt2[1])**2) # 每根分配梁的长度
            scale_y = abutmt_fenpei2_length/5000 # 分配梁在平面图中y方向上拉伸的比例, 图块的默认长度为5000
            insertpt = (O_point[0]+pt1[0], O_point[1]+pt1[1]-basic_y, 0)
            PM_Annotation_Dimension_Xdict['分配梁2'].extend([O_point[0]+pt1[0], O_point[0]+pt2[0]]) # 分配梁X标注
            PM_Annotation_Dimension_Ydict['分配梁2'].append(insertpt[1]) # 分配梁Y标注
            insert_block_incad_with_type_name(aMSP, vtpnt(insertpt), '分配梁平面', abutmt_Fenpei2_section, [1,scale_y,1], math.pi/2) 
            PM_Annotation_Mleader_Potdict['分配梁2'].append([insertpt, '分配梁F3', (-1, 1)])   
        Material_dict['分配梁']['规格'].append(f'{abutmt_Fenpei2_section}×{int(abutmt_fenpei2_length)}')
        Material_dict['分配梁']['材质'].append('Q235B')
        Material_dict['分配梁']['数量'].append(len(abutmt_fenpei2_insertpoint_lst)) 
        Material_dict['分配梁']['单重'].append(round(abutmt_fenpei2_length*zd_f2_section_weight_per_meter/1000,2)) 
        Material_dict['分配梁']['总重'].append(round(abutmt_fenpei2_length*zd_f2_section_weight_per_meter*len(abutmt_fenpei2_insertpoint_lst)/1000,1))  
    # 制动墩钢管桩
    if if_ZD == True:
        abutmt_ganggz_inserpoint_lst = [ptlst[0] for value in abutmt_ganggz_single_nodes.values() for ptlst in value]
        abutmt_ganggz_D, abutmt_ganggz_d = zd_section_cad[zd_section_name]['D'], zd_section_cad[zd_section_name]['d'] # 直径和壁厚
        for pt in abutmt_ganggz_inserpoint_lst:
            insertpt = (O_point[0]+pt[0], O_point[1]+pt[1]-basic_y, 0)
            PM_Annotation_Dimension_Ydict['制动桩'].extend([O_point[1]+pt[1]-basic_y]) # 制动墩钢管桩Y标注
            make_circle(aMSP, insertpt, abutmt_ganggz_D/2, '1粗实线')
            make_circle(aMSP, insertpt, abutmt_ganggz_D/2-abutmt_ganggz_d, '4虚线')
        Material_dict['钢管桩']['数量'].append(len(abutmt_ganggz_inserpoint_lst))
        Material_dict['钢管桩']['材质'].append('Q235B')
    Material_dict['钢管桩']['总重'].extend([round(x*y,1) for x,y in zip(Material_dict['钢管桩']['数量'],Material_dict['钢管桩']['单重'])])
    


def draw_Trestle_PM_Annotation_incad(aCAD, aMSP, Global_Scale, XY_range_dict, if_ZD, 
                                     PM_Annotation_Dimension_Xdict, PM_Annotation_Dimension_Ydict, PM_Annotation_Mleader_Potdict,
                                    ):
    # 尺寸标注
    X1_keylst = ['贝雷']
    X2_keylst = ['分配梁', '分配梁1', '分配梁2', '钢管桩']
    Y1_keylst = ['贝雷', '分配梁', '钢管桩']
    Y2_keylst = ['分配梁1', '分配梁2', '制动桩']
    PM_Annotation_Dimension_X1 = sorted(set([x for key,value in PM_Annotation_Dimension_Xdict.items() if key in X1_keylst for x in value]))
    PM_Annotation_Dimension_X2 = sorted(set([x for key,value in PM_Annotation_Dimension_Xdict.items() if key in X2_keylst for x in value]))
    PM_Annotation_Dimension_Y1 = sorted(set([y for key,value in PM_Annotation_Dimension_Ydict.items() if key in Y1_keylst for y in value]))
    PM_Annotation_Dimension_Y2 = sorted(set([y for key,value in PM_Annotation_Dimension_Ydict.items() if key in Y2_keylst for y in value]))
    Dimension_O_X1, Dimension_O_X2 = XY_range_dict['立面图']['X'] # 标注的定位X
    Dimension_O_Y1, Dimension_O_Y2 = min(PM_Annotation_Dimension_Y1+PM_Annotation_Dimension_Y2), max(PM_Annotation_Dimension_Y1+PM_Annotation_Dimension_Y2) # 标注的定位Y
    # 上标注
    for i in range(len(PM_Annotation_Dimension_X1)-1):
        X1 = PM_Annotation_Dimension_X1[i]
        X2 = PM_Annotation_Dimension_X1[i+1]
        Y0 = Dimension_O_Y2
        Add_Annotation_Linear(aMSP, (X1,Y0,0), (X2,Y0,0), (X1,Y0+12*Global_Scale,0), 0) # 两点之间标注
    # 下标注
    for i in range(len(PM_Annotation_Dimension_X2)-1):
        X1 = PM_Annotation_Dimension_X2[i]
        X2 = PM_Annotation_Dimension_X2[i+1]
        Y0 = Dimension_O_Y1
        Add_Annotation_Linear(aMSP, (X1,Y0,0), (X2,Y0,0), (X1,Y0-6*Global_Scale,0), 0) # 两点之间标注
    # 下总标注
    X1 = PM_Annotation_Dimension_X2[0]
    X2 = PM_Annotation_Dimension_X2[-1]
    Y0 = Dimension_O_Y1
    Add_Annotation_Linear(aMSP, (X1,Y0,0), (X2,Y0,0), (X1,Y0-12*Global_Scale,0), 0) # 两点之间标注
    # 左标注
    for i in range(len(PM_Annotation_Dimension_Y1)-1):
        Y1 = PM_Annotation_Dimension_Y1[i]
        Y2 = PM_Annotation_Dimension_Y1[i+1]
        X0 = Dimension_O_X1-5*Global_Scale
        Add_Annotation_Linear(aMSP, (X0,Y1,0), (X0,Y2,0), (X0-6*Global_Scale,Y1,0), 90) # 两点之间标注
    # 总左标注
    Y1 = PM_Annotation_Dimension_Y1[0]
    Y2 = PM_Annotation_Dimension_Y1[-1]
    X0 = Dimension_O_X1-5*Global_Scale
    Add_Annotation_Linear(aMSP, (X0,Y1,0), (X0,Y2,0), (X0-12*Global_Scale,Y1,0), 90) # 两点之间标注
    # 右标注
    if if_ZD == True:
        for i in range(len(PM_Annotation_Dimension_Y2)-1):
            Y1 = PM_Annotation_Dimension_Y2[i]
            Y2 = PM_Annotation_Dimension_Y2[i+1]
            X0 = Dimension_O_X2+5*Global_Scale
            Add_Annotation_Linear(aMSP, (X0,Y1,0), (X0,Y2,0), (X0+6*Global_Scale,Y1,0), 90) # 两点之间标注
        # 总右标注
        Y1 = PM_Annotation_Dimension_Y2[0]
        Y2 = PM_Annotation_Dimension_Y2[-1]
        X0 = Dimension_O_X2+5*Global_Scale
        Add_Annotation_Linear(aMSP, (X0,Y1,0), (X0,Y2,0), (X0+12*Global_Scale,Y1,0), 90) # 两点之间标注
    # 引线标注
    for value in PM_Annotation_Mleader_Potdict.values():
        for lst in value:
            startpoint = lst[0]
            txt = lst[1]
            nx, ny = lst[2]
            endpoint = [startpoint[0]+3*nx*Global_Scale, startpoint[1]+6*ny*Global_Scale, startpoint[2]]
            Add_Annotation_MLeader(aCAD, aMSP, Global_Scale, startpoint, endpoint, '5文本', 'Song_07', txt)
    # 标题
    Header_Point = [(PM_Annotation_Dimension_X1[0]+PM_Annotation_Dimension_X1[-1])/2, Dimension_O_Y2+24*Global_Scale, 0]
    Add_Header(aMSP, Global_Scale, Header_Point, '1---1', 'Song_07')
    # 边界
    XY_range_dict['平面图']['Y'].extend([Dimension_O_Y1, Dimension_O_Y2])

def draw_Trestle_DM_incad(aMSP, CAD_O_point, XY_range_dict, Global_Scale, if_ZD,
                          xl_single_nodes, xl_section_cad, xl_section_name, 
                          bailey_single_nodes, 
                          fenpei_single_nodes, distribute_section_cad, distribute_section_name, 
                          ganggz_single_nodes, gz_section_cad, gz_section_name,
                          abutmt_fenpei1_single_nodes, zd_f1_section_cad, zd_f1_section_name, 
                          abutmt_fenpei2_single_nodes, zd_f2_section_cad, zd_f2_section_name, 
                          abutmt_ganggz_single_nodes, zd_section_cad, zd_section_name, 
                          DM_Annotation_Dimension_Xdict1, DM_Annotation_Dimension_Ydict1, DM_Annotation_Mleader_Potdict1,
                          DM_Annotation_Dimension_Xdict2, DM_Annotation_Dimension_Ydict2, DM_Annotation_Mleader_Potdict2, Material_dict, 
                         ):
    GAP_X = 52*Global_Scale # 断面图与其他图之间的X间距  
    rib_length = abs(xl_single_nodes[1][0][1][1]-xl_single_nodes[1][-1][1][1])  # 小肋的长度
    O_point1 = (XY_range_dict['立面图']['X'][1]+GAP_X, CAD_O_point[1], CAD_O_point[2]) # 断面图1中的定位点
    if abutmt_ganggz_single_nodes != {} and abutmt_fenpei1_single_nodes != {} and abutmt_fenpei2_single_nodes != {}:
        O_point2 = (XY_range_dict['立面图']['X'][1]+2*GAP_X+rib_length, CAD_O_point[1], CAD_O_point[2]) # 断面图2中的定位点
    else:
        O_point2 = None
    if_bz_ganggz = True
    if_bz_abutmt_ganggz = True
    # 小肋
    Rib_XL_section, Rib_H, Rib_tf  = SEC_trans(xl_section_cad, xl_section_name)
    Material_dict['小肋']['规格'].append(f'{Rib_XL_section}×{int(rib_length)}')
    try:
        for O_point in [O_point1, O_point2]:
            rib_x1, rib_x2 = O_point[0], O_point[0] + rib_length
            rib_y1, rib_y2 = O_point[1], O_point[1] + Rib_H
            if O_point == O_point1:
                DM_Annotation_Dimension_Xdict1['桥面板'].extend([rib_x1, rib_x2])   # 桥面板X标注
            if O_point == O_point2:
                DM_Annotation_Dimension_Xdict2['桥面板'].extend([rib_x1, rib_x2])   # 桥面板X标注
            rib_ptlst = [(rib_x1, rib_y1, 0), (rib_x2, rib_y1, 0), (rib_x2, rib_y2, 0), (rib_x1, rib_y2, 0), (rib_x1, rib_y1, 0)]
            flat_points = pointlist_extend(rib_ptlst)
            make_polyline(aMSP, flat_points, '2细实线')
    except:
        pass
    # 桥面板
    deck_thickness = 10
    try:
        for O_point in [O_point1, O_point2]:
            deck_x1, deck_x2 = O_point[0], O_point[0] + rib_length
            deck_y1, deck_y2 = O_point[1] + Rib_H, O_point[1] + Rib_H + deck_thickness
            if O_point == O_point1:
                DM_Annotation_Dimension_Ydict1['桥面板'].extend([deck_y1-Rib_H, deck_y2])   # 桥面板X标注
                DM_Annotation_Mleader_Potdict1['桥面板'].append([(deck_x2, deck_y1, 0), '桥面板', (1, 1)]) # 引线点, 文本内容, 标注方向(X和Y方向, 1表示正, -1表示负)
            if O_point == O_point2:
                DM_Annotation_Dimension_Ydict2['桥面板'].extend([deck_y1-Rib_H, deck_y2])   # 桥面板X标注
                DM_Annotation_Mleader_Potdict2['桥面板'].append([(deck_x2, deck_y1, 0), '桥面板', (1, 1)]) # 引线点, 文本内容, 标注方向(X和Y方向, 1表示正, -1表示负)
            deck_ptlst = [(deck_x1, deck_y1, 0), (deck_x2, deck_y1, 0), (deck_x2, deck_y2, 0), (deck_x1, deck_y2, 0), (deck_x1, deck_y1, 0)]
            flat_points = pointlist_extend(deck_ptlst)
            make_polyline(aMSP, flat_points, '6标注线')
    except:
        pass
    # 贝雷梁
    basic_x = xl_single_nodes[1][0][1][1] # 小肋节点y坐标在midas中的ymin
    bailey_xlst = [value[0][1] for value in bailey_single_nodes.values()] # 模型中贝雷梁的y坐标
    bailey_insertpointy = CAD_O_point[1]-50 # 图纸中贝雷立面的插入点y坐标
    try:
        for O_point in [O_point1, O_point2]:
            if O_point == O_point1:
                DM_Annotation_Dimension_Xdict1['贝雷'].extend([O_point[0]+x-basic_x for x in bailey_xlst]) # 贝雷X标注
                DM_Annotation_Dimension_Ydict1['贝雷'].extend([CAD_O_point[1], CAD_O_point[1]-1500])       # 贝雷Y标注
                DM_Annotation_Mleader_Potdict1['贝雷'].append([(O_point[0]+bailey_xlst[-1]-basic_x, bailey_insertpointy-750, 0), '贝雷梁', (1, 1)]) # 引线点, 文本内容, 标注方向(X和Y方向, 1表示正, -1表示负)
            if O_point == O_point2:
                DM_Annotation_Dimension_Xdict2['贝雷'].extend([O_point[0]+x-basic_x for x in bailey_xlst]) # 贝雷X标注
                DM_Annotation_Dimension_Ydict2['贝雷'].extend([CAD_O_point[1], CAD_O_point[1]-1500])       # 贝雷Y标注
                DM_Annotation_Mleader_Potdict2['贝雷'].append([(O_point[0]+bailey_xlst[-1]-basic_x, bailey_insertpointy-750, 0), '贝雷梁', (1, 1)]) # 引线点, 文本内容, 标注方向(X和Y方向, 1表示正, -1表示负)
            for x in bailey_xlst:
                insert_block_incad_with_type_name(aMSP, vtpnt((O_point[0]+x-basic_x, bailey_insertpointy, 0)), '贝雷侧立面', '3')
    except:
        pass
    # 非制动墩分配梁
    Fenpei_section, fenpei_H, fenpei_tf = SEC_trans(distribute_section_cad, distribute_section_name)
    fenpei_xlst = [fenpei_single_nodes[1][0][1][1], fenpei_single_nodes[1][-1][1][1]] # 模型中分配梁端点y坐标
    fenpei_length = abs(fenpei_xlst[0]-fenpei_xlst[1])  # 分配梁的长度
    fenpei_x1, fenpei_x2 = O_point1[0]+fenpei_xlst[0]-basic_x , O_point1[0]+fenpei_xlst[1]-basic_x
    fenpei_y1, fenpei_y2 = bailey_insertpointy-1450, bailey_insertpointy-1450-fenpei_H
    fenpei_y3, fenpei_y4 = bailey_insertpointy-1450-fenpei_tf, bailey_insertpointy-1450-fenpei_H+fenpei_tf
    fenpei_ptlst1 = [(fenpei_x1, fenpei_y1, 0),(fenpei_x2, fenpei_y1, 0),(fenpei_x2, fenpei_y2, 0),(fenpei_x1, fenpei_y2, 0),(fenpei_x1, fenpei_y1, 0)]
    flat_points = pointlist_extend(fenpei_ptlst1)
    make_polyline(aMSP, flat_points, '1粗实线')
    fenpei_ptlst2 = [(fenpei_x1, fenpei_y3, 0),(fenpei_x2, fenpei_y3, 0),(fenpei_x2, fenpei_y4, 0),(fenpei_x1, fenpei_y4, 0),(fenpei_x1, fenpei_y3, 0)]
    flat_points = pointlist_extend(fenpei_ptlst2)
    make_polyline(aMSP, flat_points, '2细实线')
    DM_Annotation_Dimension_Xdict1['分配梁'].extend([fenpei_x1, fenpei_x2]) # 分配梁X标注
    DM_Annotation_Dimension_Ydict1['分配梁'].extend([fenpei_y1, fenpei_y2]) # 分配梁Y标注
    DM_Annotation_Mleader_Potdict1['分配梁'].append([(fenpei_x2, fenpei_y2, 0), '分配梁F1', (1, -1)]) # 引线点, 文本内容, 标注方向(X和Y方向, 1表示正, -1表示负)
    # 制动墩分配梁
    if if_ZD == True:
        abutmt_Fenpei1_section, abutmt_fenpei1_H, abutmt_fenpei1_tf = SEC_trans(zd_f1_section_cad, zd_f1_section_name)
        abutmt_fenpei1_xlst = [abutmt_fenpei1_single_nodes[1][0][1][1], abutmt_fenpei1_single_nodes[1][-1][1][1]] # 模型中制动墩上的分配梁1端点y坐标
        abutmt_fenpei1_length = abs(abutmt_fenpei1_xlst[0]-abutmt_fenpei1_xlst[1])  # 分配梁1的长度
        abutmt_fenpei1_x1, abutmt_fenpei1_x2 = O_point2[0]+abutmt_fenpei1_xlst[0]-basic_x , O_point2[0]+abutmt_fenpei1_xlst[1]-basic_x
        abutmt_fenpei1_y1, abutmt_fenpei1_y2 = bailey_insertpointy-1450, bailey_insertpointy-1450-abutmt_fenpei1_H
        abutmt_fenpei1_y3, abutmt_fenpei1_y4 = bailey_insertpointy-1450-abutmt_fenpei1_tf, bailey_insertpointy-1450-abutmt_fenpei1_H+abutmt_fenpei1_tf
        abutmt_fenpei1_ptlst1 = [(abutmt_fenpei1_x1, abutmt_fenpei1_y1, 0),(abutmt_fenpei1_x2, abutmt_fenpei1_y1, 0),(abutmt_fenpei1_x2, abutmt_fenpei1_y2, 0),(abutmt_fenpei1_x1, abutmt_fenpei1_y2, 0),(abutmt_fenpei1_x1, abutmt_fenpei1_y1, 0)]
        flat_points = pointlist_extend(abutmt_fenpei1_ptlst1)
        make_polyline(aMSP, flat_points, '1粗实线')
        abutmt_fenpei1_ptlst2 = [(abutmt_fenpei1_x1, abutmt_fenpei1_y3, 0),(abutmt_fenpei1_x2, abutmt_fenpei1_y3, 0),(abutmt_fenpei1_x2, abutmt_fenpei1_y4, 0),(abutmt_fenpei1_x1, abutmt_fenpei1_y4, 0),(abutmt_fenpei1_x1, abutmt_fenpei1_y3, 0)]
        flat_points = pointlist_extend(abutmt_fenpei1_ptlst2)
        make_polyline(aMSP, flat_points, '2细实线')
        DM_Annotation_Dimension_Xdict2['分配梁1'].extend([abutmt_fenpei1_x1, abutmt_fenpei1_x2]) # 分配梁1X标注
        DM_Annotation_Dimension_Ydict2['分配梁1'].extend([abutmt_fenpei1_y1, abutmt_fenpei1_y2]) # 分配梁1Y标注
        DM_Annotation_Mleader_Potdict2['分配梁1'].append([(abutmt_fenpei1_x2, abutmt_fenpei1_y2, 0), '分配梁F2', (1, -1)]) # 引线点, 文本内容, 标注方向(X和Y方向, 1表示正, -1表示负)
        abutmt_Fenpei2_section, abutmt_fenpei2_H, abutmt_fenpei2_tf = SEC_trans(zd_f2_section_cad, zd_f2_section_name)
        abutmt_Fenpei2_xlst = sorted(list(set(tuple([value[0][1][1] for value in abutmt_fenpei2_single_nodes.values()]))))
        abutmt_fenpei2_insertpointy = abutmt_fenpei1_y2-abutmt_fenpei2_H
        for x in abutmt_Fenpei2_xlst:
            abutmt_fenpei2_insert_point = [O_point2[0]+x-basic_x, abutmt_fenpei2_insertpointy, 0]
            insert_block_incad_with_type_name(aMSP, vtpnt(abutmt_fenpei2_insert_point), '分配梁立面', abutmt_Fenpei2_section)
            DM_Annotation_Mleader_Potdict2['分配梁2'].append([(O_point2[0]+x-basic_x, abutmt_fenpei2_insertpointy+abutmt_fenpei2_H/2, 0), '分配梁F3', (1, -1)]) # 引线点, 文本内容, 标注方向(X和Y方向, 1表示正, -1表示负)
        DM_Annotation_Dimension_Xdict2['分配梁2'].extend([O_point2[0]+x-basic_x for x in abutmt_Fenpei2_xlst]) # 分配梁2X标注
        DM_Annotation_Dimension_Ydict2['分配梁2'].extend([abutmt_fenpei2_insertpointy, abutmt_fenpei2_insertpointy+abutmt_fenpei2_H]) # 分配梁2Y标注
    # 非制动墩钢管桩
    ganggz_xlst = [ptlst[0][1] for ptlst in ganggz_single_nodes[1]] # 模型中钢管桩的y坐标
    ganggz_insertpointy = fenpei_y2
    ganggz_length = abs(ganggz_single_nodes[1][0][0][2] - ganggz_single_nodes[1][0][-1][2]) # 钢管桩桩长
    ganggz_D, ganggz_d = gz_section_cad[gz_section_name]['D'], gz_section_cad[gz_section_name]['d'] # 直径和壁厚
    for x in ganggz_xlst:
        try:
            insert_block_incad_with_block_name(aMSP, vtpnt([O_point1[0]+x-basic_x, ganggz_insertpointy, 0]), f'φ{ganggz_D}开槽柱头截断线钢管桩1block', f'EX_φ{ganggz_D}')
            DM_Annotation_Mleader_Potdict1['钢管桩'].append([(O_point1[0]+x-basic_x, ganggz_insertpointy-4500, 0), '钢管桩G1', (1, -1)]) # 引线点, 文本内容, 4500是钢管桩图块长7000, 固定值
        except:
            print('未匹配非制动墩钢管桩直径, 生成钢管桩简图')
            ganggz_ptlst1 = [[O_point1[0]+x-basic_x+nx*ganggz_D/2, ganggz_insertpointy+ny*ganggz_length, 0] for nx, ny in [(-1,0),(-1,-1),(1,-1),(1,0),(-1,0)]]
            flat_points = pointlist_extend(ganggz_ptlst1)
            make_polyline(aMSP, flat_points, '1粗实线')
            ganggz_ptlst2 = [[O_point1[0]+x-basic_x+nx*ganggz_D/2-ganggz_d, ganggz_insertpointy+ny*ganggz_length, 0] for nx, ny in [(-1,0),(-1,-1),(1,-1),(1,0),(-1,0)]]
            flat_points = pointlist_extend(ganggz_ptlst2)
            make_polyline(aMSP, flat_points, '2细实线')
            make_line(aMSP, [(O_point1[0]+x-basic_x, ganggz_insertpointy, 0), (O_point1[0]+x-basic_x, ganggz_insertpointy-ganggz_length, 0)], '3中心线')
            DM_Annotation_Mleader_Potdict1['钢管桩'].append([(O_point1[0]+x-basic_x, ganggz_insertpointy-ganggz_length/2, 0), '钢管桩G1', (1, -1)]) # 引线点, 文本内容, 标注方向(X和Y方向, 1表示正, -1表示负)
            if_bz_ganggz = False
    DM_Annotation_Dimension_Xdict1['钢管桩'].extend([O_point1[0]+x-basic_x for x in ganggz_xlst]) # 钢管桩X标注
    if if_bz_ganggz:
        DM_Annotation_Dimension_Ydict1['钢管桩'].extend([ganggz_insertpointy, ganggz_insertpointy-7000]) # 钢管桩Y标注
    else:
        DM_Annotation_Dimension_Ydict1['钢管桩'].extend([ganggz_insertpointy, ganggz_insertpointy-ganggz_length]) # 钢管桩Y标注
    Material_dict['钢管桩']['规格'].append(f'∅{ganggz_D}×{ganggz_d}×{int(ganggz_length)}')
    # 联结系
    ljx_xlst = ganggz_xlst
    ljx_ylst = [ganggz_insertpointy-1000, ganggz_insertpointy-3000]
    for i in range(len(ljx_xlst)-1):
        ptx1, ptx2 = O_point1[0]+ljx_xlst[i]-basic_x, O_point1[0]+ljx_xlst[i+1]-basic_x
        pty1, pty2 = ljx_ylst
        try:
            ljx_length = int(abs(ptx1-ptx2)/1000)
            insert_block_incad_with_block_name(aMSP, vtpnt([(ptx1+ptx2)/2, ganggz_insertpointy-2000, 0]), f'φ{ganggz_D}型钢{ljx_length}m联结系block', f'EX_φ{ganggz_D}')
        except:
            print('未匹配已有联结系长度, 生成联结系简图')
            make_line(aMSP, [(ptx1, pty1, 0), (ptx2, pty1, 0)], '3中心线')
            make_line(aMSP, [(ptx1, pty2, 0), (ptx2, pty2, 0)], '3中心线')
            make_line(aMSP, [(ptx1, pty1, 0), (ptx2, pty2, 0)], '3中心线')
            make_line(aMSP, [(ptx1, pty2, 0), (ptx2, pty1, 0)], '3中心线')
    DM_Annotation_Dimension_Ydict1['联结系'].extend(ljx_ylst) # 联结系Y标注
    # 制动墩钢管桩
    if if_ZD == True:
        abutmt_ganggz_xlst = sorted(list(set(tuple([ptlst[0][1] for ptlst in abutmt_ganggz_single_nodes[1]])))) # 模型中制动墩钢管桩的y坐标
        abutmt_ganggz_length = abs(abutmt_ganggz_single_nodes[1][0][0][2] - abutmt_ganggz_single_nodes[1][0][-1][2]) # 制动墩钢管桩桩长
        abutmt_ganggz_D, abutmt_ganggz_d = zd_section_cad[zd_section_name]['D'], zd_section_cad[zd_section_name]['d'] # 直径和壁厚
        abutmt_ganggz_insertpointy = abutmt_fenpei2_insertpointy # 制动墩钢管桩插入点Y坐标
        for x in abutmt_ganggz_xlst:
            try:
                insert_block_incad_with_block_name(aMSP, vtpnt([O_point2[0]+x-basic_x, abutmt_ganggz_insertpointy, 0]), f'φ{ganggz_D}开槽柱头截断线钢管桩2block', f'EX_φ{ganggz_D}')
                DM_Annotation_Mleader_Potdict2['制动桩'].append([(O_point2[0]+x-basic_x, ganggz_insertpointy-4500, 0), '钢管桩G2', (1, -1)]) # 引线点, 文本内容, 4500是钢管桩图块长7000, 固定值
            except:
                print('未匹配制动墩钢管桩直径, 生成钢管桩简图')
                abutmt_ganggz_ptlst1 = [[O_point2[0]+x-basic_x+nx*abutmt_ganggz_D/2, abutmt_ganggz_insertpointy+ny*abutmt_ganggz_length, 0] for nx, ny in [(-1,0),(-1,-1),(1,-1),(1,0),(-1,0)]]
                flat_points = pointlist_extend(abutmt_ganggz_ptlst1)
                make_polyline(aMSP, flat_points, '1粗实线')
                abutmt_ganggz_ptlst2 = [[O_point2[0]+x-basic_x+nx*abutmt_ganggz_D/2-abutmt_ganggz_d, abutmt_ganggz_insertpointy+ny*abutmt_ganggz_length, 0] for nx, ny in [(-1,0),(-1,-1),(1,-1),(1,0),(-1,0)]]
                flat_points = pointlist_extend(abutmt_ganggz_ptlst2)
                make_polyline(aMSP, flat_points, '2细实线')
                make_line(aMSP, [(O_point2[0]+x-basic_x, abutmt_ganggz_insertpointy, 0), (O_point2[0]+x-basic_x, abutmt_ganggz_insertpointy-abutmt_ganggz_length, 0)], '3中心线')
                DM_Annotation_Mleader_Potdict2['制动桩'].append([(O_point2[0]+x-basic_x, abutmt_ganggz_insertpointy-abutmt_ganggz_length/2, 0), '钢管桩G2', (1, -1)]) # 引线点, 文本内容, 标注方向(X和Y方向, 1表示正, -1表示负)
                if_bz_abutmt_ganggz = False
        DM_Annotation_Dimension_Xdict2['制动桩'].extend([O_point2[0]+x-basic_x for x in abutmt_ganggz_xlst]) # 制动桩X标注
        if if_bz_abutmt_ganggz:
            DM_Annotation_Dimension_Ydict2['制动桩'].extend([abutmt_ganggz_insertpointy, abutmt_ganggz_insertpointy-7000]) # 制动桩Y标注
        else:
            DM_Annotation_Dimension_Ydict2['制动桩'].extend([abutmt_ganggz_insertpointy, abutmt_ganggz_insertpointy-abutmt_ganggz_length]) # 制动桩Y标注
        Material_dict['钢管桩']['规格'].append(f'∅{abutmt_ganggz_D}×{abutmt_ganggz_d}×{int(abutmt_ganggz_length)}')
        # 联结系  
        ljx_xlst = abutmt_ganggz_xlst
        ljx_ylst = [abutmt_ganggz_insertpointy-1000, abutmt_ganggz_insertpointy-3000]
        for i in range(len(ljx_xlst)-1):
            ptx1, ptx2 = O_point2[0]+ljx_xlst[i]-basic_x, O_point2[0]+ljx_xlst[i+1]-basic_x
            pty1, pty2 = ljx_ylst
            try:
                ljx_length = int(abs(ptx1-ptx2)/1000)
                insert_block_incad_with_block_name(aMSP, vtpnt([(ptx1+ptx2)/2, abutmt_ganggz_insertpointy-2000, 0]), f'φ{abutmt_ganggz_D}型钢{ljx_length}m联结系block', f'EX_φ{abutmt_ganggz_D}')
            except:
                print('未匹配已有联结系长度, 生成联结系简图')
                make_line(aMSP, [(ptx1, pty1, 0), (ptx2, pty1, 0)], '3中心线')
                make_line(aMSP, [(ptx1, pty2, 0), (ptx2, pty2, 0)], '3中心线')
                make_line(aMSP, [(ptx1, pty1, 0), (ptx2, pty2, 0)], '3中心线')
                make_line(aMSP, [(ptx1, pty2, 0), (ptx2, pty1, 0)], '3中心线')
        DM_Annotation_Dimension_Ydict2['联结系'].extend(ljx_ylst) # 联结系Y标注
    # print(DM_Annotation_Dimension_Ydict2)


    
def draw_Trestle_DM_Annotation_incad(aCAD, aMSP, Global_Scale, XY_range_dict, if_ZD, ganggz_single_nodes, abutmt_ganggz_single_nodes, 
                                     DM_Annotation_Dimension_Xdict1, DM_Annotation_Dimension_Ydict1, DM_Annotation_Mleader_Potdict1,
                                     DM_Annotation_Dimension_Xdict2, DM_Annotation_Dimension_Ydict2, DM_Annotation_Mleader_Potdict2,  
                                    ):
    # 非制动墩
    # 尺寸标注
    X1_keylst = ['桥面板', '贝雷']
    X2_keylst = ['分配梁', '钢管桩']
    DM_Annotation_Dimension_X1 = sorted(set([x for key,value in DM_Annotation_Dimension_Xdict1.items() if key in X1_keylst for x in value]))
    DM_Annotation_Dimension_X2 = sorted(set([x for key,value in DM_Annotation_Dimension_Xdict1.items() if key in X2_keylst for x in value])) 
    DM_Annotation_Dimension_Y1 = sorted(set([y for value in DM_Annotation_Dimension_Ydict1.values() for y in value]), reverse = True)
    Dimension_O_X1, Dimension_O_X2 = min(DM_Annotation_Dimension_X1+DM_Annotation_Dimension_X2), max(DM_Annotation_Dimension_X1+DM_Annotation_Dimension_X2) # 标注的定位X
    Dimension_O_Y1, Dimension_O_Y2 = min(DM_Annotation_Dimension_Y1), max(DM_Annotation_Dimension_Y1) # 标注的定位Y
    ganggz_length = abs(ganggz_single_nodes[1][0][0][2] - ganggz_single_nodes[1][0][-1][2]) # 钢管桩桩长
    # 上标注
    for i in range(len(DM_Annotation_Dimension_X1)-1):
        X1 = DM_Annotation_Dimension_X1[i]
        X2 = DM_Annotation_Dimension_X1[i+1]
        Y0 = Dimension_O_Y2
        Add_Annotation_Linear(aMSP, (X1,Y0,0), (X2,Y0,0), (X1,Y0+6*Global_Scale,0), 0) # 两点之间标注
    # 下标注
    for i in range(len(DM_Annotation_Dimension_X2)-1):
        X1 = DM_Annotation_Dimension_X2[i]
        X2 = DM_Annotation_Dimension_X2[i+1]
        Y0 = Dimension_O_Y1
        Add_Annotation_Linear(aMSP, (X1,Y0,0), (X2,Y0,0), (X1,Y0-6*Global_Scale,0), 0) # 两点之间标注
    # 下总标注
    X1 = DM_Annotation_Dimension_X2[0]
    X2 = DM_Annotation_Dimension_X2[-1]
    Y0 = Dimension_O_Y1
    Add_Annotation_Linear(aMSP, (X1,Y0,0), (X2,Y0,0), (X1,Y0-12*Global_Scale,0), 0) # 两点之间标注
    # 左标注
    for i in range(len(DM_Annotation_Dimension_Y1)-1):
        Y1 = DM_Annotation_Dimension_Y1[i]
        Y2 = DM_Annotation_Dimension_Y1[i+1]
        X0 = Dimension_O_X1
        if i != len(DM_Annotation_Dimension_Y1)-2:
            Add_Annotation_Linear(aMSP, (X0,Y1,0), (X0,Y2,0), (X0-6*Global_Scale,Y1,0), 90) # 两点之间标注
        else:
            Annotation_Linear_txt = ganggz_length - 3000
            Add_Annotation_Linear(aMSP, (X0,Y1,0), (X0,Y2,0), (X0-6*Global_Scale,Y1,0), 90, '6标注线', Annotation_Linear_txt)
    # 总左标注
    Y1 = DM_Annotation_Dimension_Y1[0]
    Y2 = DM_Annotation_Dimension_Y1[-1]
    X0 = Dimension_O_X1
    Annotation_Linear_txt = abs(Y1-Y2) + (ganggz_length - 7000)
    Add_Annotation_Linear(aMSP, (X0,Y1,0), (X0,Y2,0), (X0-12*Global_Scale,Y1,0), 90, '6标注线', Annotation_Linear_txt) # 两点之间标注
    # 标题
    Header_Point = [(DM_Annotation_Dimension_X1[0]+DM_Annotation_Dimension_X1[-1])/2, Dimension_O_Y2+24*Global_Scale, 0]
    Add_Header(aMSP, Global_Scale, Header_Point, '2---2', 'Song_07')
    # 制动墩
    if if_ZD == True:
        # 尺寸标注
        X1_keylst = ['桥面板', '贝雷']
        X2_keylst = ['分配梁1', '分配梁2', '制动桩']
        DM_Annotation_Dimension_X1 = sorted(set([x for key,value in DM_Annotation_Dimension_Xdict2.items() if key in X1_keylst for x in value]))
        DM_Annotation_Dimension_X2 = sorted(set([x for key,value in DM_Annotation_Dimension_Xdict2.items() if key in X2_keylst for x in value])) 
        DM_Annotation_Dimension_Y1 = sorted(set([y for value in DM_Annotation_Dimension_Ydict2.values() for y in value]), reverse = True)
        Dimension_O_X1, Dimension_O_X2 = min(DM_Annotation_Dimension_X1+DM_Annotation_Dimension_X2), max(DM_Annotation_Dimension_X1+DM_Annotation_Dimension_X2) # 标注的定位X
        Dimension_O_Y1, Dimension_O_Y2 = min(DM_Annotation_Dimension_Y1), max(DM_Annotation_Dimension_Y1) # 标注的定位Y
        abutmt_ganggz_length = abs(abutmt_ganggz_single_nodes[1][0][0][2] - abutmt_ganggz_single_nodes[1][0][-1][2]) # 制动墩钢管桩桩长
        # 上标注
        for i in range(len(DM_Annotation_Dimension_X1)-1):
            X1 = DM_Annotation_Dimension_X1[i]
            X2 = DM_Annotation_Dimension_X1[i+1]
            Y0 = Dimension_O_Y2
            Add_Annotation_Linear(aMSP, (X1,Y0,0), (X2,Y0,0), (X1,Y0+6*Global_Scale,0), 0) # 两点之间标注
        # 下标注
        for i in range(len(DM_Annotation_Dimension_X2)-1):
            X1 = DM_Annotation_Dimension_X2[i]
            X2 = DM_Annotation_Dimension_X2[i+1]
            Y0 = Dimension_O_Y1
            Add_Annotation_Linear(aMSP, (X1,Y0,0), (X2,Y0,0), (X1,Y0-6*Global_Scale,0), 0) # 两点之间标注
        # 下总标注
        X1 = DM_Annotation_Dimension_X2[0]
        X2 = DM_Annotation_Dimension_X2[-1]
        Y0 = Dimension_O_Y1
        Add_Annotation_Linear(aMSP, (X1,Y0,0), (X2,Y0,0), (X1,Y0-12*Global_Scale,0), 0) # 两点之间标注
        # 左标注
        if if_ZD == True:
            for i in range(len(DM_Annotation_Dimension_Y1)-1):
                Y1 = DM_Annotation_Dimension_Y1[i]
                Y2 = DM_Annotation_Dimension_Y1[i+1]
                X0 = Dimension_O_X1
                if i != len(DM_Annotation_Dimension_Y1)-2:
                    Add_Annotation_Linear(aMSP, (X0,Y1,0), (X0,Y2,0), (X0-6*Global_Scale,Y1,0), 90) # 两点之间标注
                else:
                    Annotation_Linear_txt = abutmt_ganggz_length - 3000
                    Add_Annotation_Linear(aMSP, (X0,Y1,0), (X0,Y2,0), (X0-6*Global_Scale,Y1,0), 90, '6标注线', Annotation_Linear_txt) # 钢管桩最后一个标注
        # 总左标注
        Y1 = DM_Annotation_Dimension_Y1[0]
        Y2 = DM_Annotation_Dimension_Y1[-1]
        X0 = Dimension_O_X1
        Annotation_Linear_txt = abs(Y1-Y2) + (abutmt_ganggz_length - 7000)
        Add_Annotation_Linear(aMSP, (X0,Y1,0), (X0,Y2,0), (X0-12*Global_Scale,Y1,0), 90, '6标注线', Annotation_Linear_txt) # 两点之间标注
        # 引线标注
        for value in DM_Annotation_Mleader_Potdict1.values():
            for lst in value:
                startpoint = lst[0]
                txt = lst[1]
                nx, ny = lst[2]
                endpoint = [startpoint[0]+3*nx*Global_Scale, startpoint[1]+6*ny*Global_Scale, startpoint[2]]
                Add_Annotation_MLeader(aCAD, aMSP, Global_Scale, startpoint, endpoint, '5文本', 'Song_07', txt)
        for value in DM_Annotation_Mleader_Potdict2.values():
            for lst in value:
                startpoint = lst[0]
                txt = lst[1]
                nx, ny = lst[2]
                endpoint = [startpoint[0]+3*nx*Global_Scale, startpoint[1]+6*ny*Global_Scale, startpoint[2]]
                Add_Annotation_MLeader(aCAD, aMSP, Global_Scale, startpoint, endpoint, '5文本', 'Song_07', txt)
        # 标题
        Header_Point = [(DM_Annotation_Dimension_X1[0]+DM_Annotation_Dimension_X1[-1])/2, Dimension_O_Y2+24*Global_Scale, 0]
        Add_Header(aMSP, Global_Scale, Header_Point, '3---3', 'Song_07')


def draw_frame_notes_in_cad(acadmsp, Global_Scale, LM_Annotation_Dimension_Xdict, LM_Annotation_Dimension_Ydict, PM_Annotation_Dimension_Xdict, PM_Annotation_Dimension_Ydict):
    # 图框
    Draw_LeftX = sorted([x for value in PM_Annotation_Dimension_Xdict.values() for x in value])[0] # 平面图最左边的点
    Draw_TopY, Draw_Buttom_Y = sorted(LM_Annotation_Dimension_Ydict['桥面板'])[1], sorted(PM_Annotation_Dimension_Ydict['分配梁'])[0] # 立面图和平面图的最高点和最低点
    Draw_Y, Draw_Middle_Y = Draw_TopY - Draw_Buttom_Y, (Draw_TopY + Draw_Buttom_Y)/2 # 立面图和平面图占图框中的总高度Y及中点
    frame_insertpoint = [Draw_LeftX-75*Global_Scale, Draw_Middle_Y-297*Global_Scale/2-12*Global_Scale,0]
    insert_block_incad_with_block_name(acadmsp, vtpnt(frame_insertpoint), '设计分公司图框（二分公司）', f'设计分公司图框（二分公司）', [Global_Scale, Global_Scale, 1], 0)
    # 附注
    notes_lst = ['附注：', '1、图中尺寸除高程以m计外，其余均以mm计。']
    notes_insertpoint = [Draw_LeftX, Draw_Buttom_Y-30*Global_Scale,0]
    Add_MText(acadmsp, notes_insertpoint, notes_lst, 2.5*Global_Scale, 'Song_07', '5文本')  



# 总函数
def All(app,acadapp):   
    print("开始绘制")
    components_tab = app.tab_components
    # 参数获取
    (
        mct_lst, xl_single_nodes, single_start_nodes, distribute_single_nodes, brake_f2_nodes, brake_f1_nodes, pile_group_nodes, brake_pile_group_nodes, xl_section_cad, xl_section_name,distribute_section_cad, distribute_section_name,gz_section_cad, gz_section_name,zd_f1_section_cad, zd_f1_section_name,zd_f2_section_cad, zd_f2_section_name,zd_section_cad, zd_section_name,section_weight_per_meter_lst,Feng_values,Feng_result_lst, Shui_values, Shui_lst,single_csv_nodes
    ) = build_mct_from_ui(app)
    print("参数已获取")
    
    # 小肋
    xl_single_nodes = xl_single_nodes
    # 贝雷梁
    print(single_start_nodes)
    bailey_single_nodes = single_start_nodes
    # 分配梁
    fenpei_single_nodes = distribute_single_nodes
    # 钢管桩
    ganggz_single_nodes = strip_node_ids(pile_group_nodes)
    print(ganggz_single_nodes)
    # 2排非制动墩钢管桩
    # 制动墩顶分配梁1(上面的)
    abutmt_fenpei1_single_nodes = brake_f1_nodes
    # abutmt_fenpei1_single_nodes = {  
    #     1: [(13, (8900.0, 500.0, -2000.0)), (16, (8900.0, 3200.0, -2000.0)), (17, (8900.0, 4100.0, -2000.0)), (18, (8900.0, 5000.0, -2000.0))], 
    #     # 2: [(13, (0.0, 500.0, -2000.0)), (16, (0.0, 3200.0, -2000.0)), (17, (0.0, 4100.0, -2000.0)), (18, (0.0, 5000.0, -2000.0))], 
    # } # 两个制动墩
    # abutmt_fenpei1_single_nodes = {}
    # 制动墩顶分配梁2(下面的)
    abutmt_fenpei2_single_nodes = brake_f2_nodes
    # abutmt_fenpei2_single_nodes = {
    #     1: [(1, (7600.0, 1250.0, -2500.0)), (4, (8000.0, 1250.0, -2500.0)), (5, (9000.0, 1250.0, -2500.0)), (6, (10200.0, 1250.0, -2500.0))],
    #     2: [(7, (7600.0, 4250.0, -2000.0)), (10, (8000.0, 4250.0, -2000.0)), (11, (9000.0, 4250.0, -2000.0)), (12, (10200.0, 4250.0, -2000.00))], 
    #     3: [(1, (-3000.0, 1250.0, -2500.0)), (4, (1000.0, 1250.0, -2500.0)), (5, (2000.0, 1250.0, -2500.0)), (6, (3000.0, 1250.0, -2500.0))],
    #     4: [(7, (-3000.0, 4250.0, -2000.0)), (10, (1000.0, 4250.0, -2000.0)), (11, (2000.0, 4250.0, -2000.0)), (12, (3000.0, 4250.0, -2000.00))], 
    # } # 两个制动墩
    # abutmt_fenpei2_single_nodes = {}
    brake_pile_group_nodes = strip_node_ids(brake_pile_group_nodes)
    abutmt_ganggz_single_nodes = merge_brake_pile_groups(brake_pile_group_nodes)
    print(abutmt_ganggz_single_nodes)
    # abutmt_ganggz_single_nodes = {}

    # 延米重
    # xl_section_weight_per_meter = 30
    # distribute_section_per_meter = 40
    # gz_section_weight_per_meter = 50
    # zd_f1_section_weight_per_meter = 60
    # zd_f2_section_weight_per_meter = 70
    # zd_section_weight_per_meter = 80

    # 钢管桩    
    # acadapp = win32com.client.Dispatch("AutoCAD.Application")
    acaddoc = acadapp.ActiveDocument
    acadmsp = acaddoc.ModelSpace
    print("请在cad中选取一点作为栈桥总布置图的起点。")
    CAD_O_point = acaddoc.Utility.Getpoint() # CAD相对定位原点 # 用户指定的基准点
    Global_Scale = float(acaddoc.Utility.GetString(1, '请输入图纸比例')) # 全局比例
    XY_range_dict = {'平面图':{'X':[], 'Y':[]}, '立面图':{'X':[], 'Y':[]}, '断面图':{'X':[], 'Y':[]}} # 每部分图的宽X和高Y
    print(CAD_O_point)
    print(Global_Scale)
    # 在cad中生成0.7宋体, 字体名称为 Song_07
    Load_Font(acaddoc, 'Song_07', 0.0, 0.7)
    # 在cad中生成对应比例的标注样式
    Load_Standrad_Style(acaddoc, Global_Scale, f'DimStyle_{int(Global_Scale)}', 'Song_07')
    Material_dict = {'桥面板':{'名称':['桥面板'] , '规格':[]               , '材质':['Q235B'], '数量':[1] , '单重':[]   , '总重':[], '备注':['花纹钢板']},
                     '小肋'  :{'名称':['小肋']   , '规格':[]               , '材质':['Q235B'], '数量':[]  , '单重':[]   , '总重':[], '备注':['']       }, 
                     '贝雷'  :{'名称':['3m贝雷梁'], '规格':['“1500×3000”型'], '材质':['16Mn'] , '数量':[] , '单重':[270], '总重':[], '备注':['']       }, 
                     '分配梁':{'名称':[]         , '规格':[]               , '材质':[]       , '数量':[]  , '单重':[]   , '总重':[], '备注':[]         }, 
                     '钢管桩':{'名称':[]         , '规格':[]               , '材质':[]       , '数量':[]  , '单重':[]   , '总重':[], '备注':[]         },
                    }
    left_support  = components_tab.left_select_var.get()
    right_support = components_tab.right_select_var.get()
    has_brake = ("制动墩" in {left_support, right_support})
    if has_brake:
        if_ZD = True
    else:
        if_ZD = False
    print("标注样式生成完成")
    # 构件延米重
    xl_section_weight_per_meter, distribute_section_per_meter, ljx_section_per_meter, zd_f1_section_weight_per_meter, zd_f2_section_weight_per_meter = section_weight_per_meter_lst
    print("构件延米重已获取")
    # 绘制立面图
    LM_Annotation_Dimension_Xdict = {'贝雷':[], '钢管桩':[]} # 立面图中尺寸标注的X坐标
    LM_Annotation_Dimension_Ydict = {'桥面板':[], '贝雷':[], '分配梁':[], '分配梁1':[], '分配梁2':[], '钢管桩':[], '制动桩':[], '联结系':[]} # 立面图中尺寸标注的Y坐标
    LM_Annotation_Mleader_Potdict = {'桥面板':[], '贝雷':[], '分配梁':[], '分配梁1':[], '分配梁2':[], '钢管桩':[], '制动桩':[]} # 立面图中引线标注的所有点坐标
    draw_Trestle_LM_incad(
        acadmsp, CAD_O_point, XY_range_dict, Global_Scale, if_ZD, xl_section_weight_per_meter, 
        xl_single_nodes, xl_section_cad, xl_section_name, 
        bailey_single_nodes, 
        fenpei_single_nodes, distribute_section_cad, distribute_section_name, 
        ganggz_single_nodes, gz_section_cad, gz_section_name,
        abutmt_fenpei1_single_nodes, zd_f1_section_cad, zd_f1_section_name, 
        abutmt_fenpei2_single_nodes, zd_f2_section_cad, zd_f2_section_name, 
        abutmt_ganggz_single_nodes, zd_section_cad, zd_section_name,
        LM_Annotation_Dimension_Xdict, LM_Annotation_Dimension_Ydict, LM_Annotation_Mleader_Potdict, Material_dict, 
    )
    #绘制立面图标注
    draw_Trestle_LM_Annotation_incad(
        acadapp, acadmsp, Global_Scale, XY_range_dict, if_ZD, ganggz_single_nodes, abutmt_ganggz_single_nodes,
        LM_Annotation_Dimension_Xdict, LM_Annotation_Dimension_Ydict, LM_Annotation_Mleader_Potdict,
    )
    # 绘制断面图
    DM_Annotation_Dimension_Xdict1 = {'桥面板':[], '贝雷':[], '分配梁':[], '钢管桩':[]} 
    DM_Annotation_Dimension_Ydict1 = {'桥面板':[], '贝雷':[], '分配梁':[], '钢管桩':[], '联结系':[]} 
    DM_Annotation_Mleader_Potdict1 = {'桥面板':[], '贝雷':[], '分配梁':[], '钢管桩':[], '联结系':[]}
    DM_Annotation_Dimension_Xdict2 = {'桥面板':[], '贝雷':[], '分配梁1':[], '分配梁2':[], '制动桩':[]} 
    DM_Annotation_Dimension_Ydict2 = {'桥面板':[], '贝雷':[], '分配梁1':[], '分配梁2':[], '制动桩':[], '联结系':[]} 
    DM_Annotation_Mleader_Potdict2 = {'桥面板':[], '贝雷':[], '分配梁1':[], '分配梁2':[], '制动桩':[], '联结系':[]}
    draw_Trestle_DM_incad(
        acadmsp, CAD_O_point, XY_range_dict, Global_Scale, if_ZD,
        xl_single_nodes, xl_section_cad, xl_section_name, 
        bailey_single_nodes, 
        fenpei_single_nodes, distribute_section_cad, distribute_section_name, 
        ganggz_single_nodes, gz_section_cad, gz_section_name,
        abutmt_fenpei1_single_nodes, zd_f1_section_cad, zd_f1_section_name, 
        abutmt_fenpei2_single_nodes, zd_f2_section_cad, zd_f2_section_name, 
        abutmt_ganggz_single_nodes, zd_section_cad, zd_section_name,    
        DM_Annotation_Dimension_Xdict1, DM_Annotation_Dimension_Ydict1, DM_Annotation_Mleader_Potdict1, 
        DM_Annotation_Dimension_Xdict2, DM_Annotation_Dimension_Ydict2, DM_Annotation_Mleader_Potdict2, Material_dict, 
    )
    # 绘制断面图标注
    draw_Trestle_DM_Annotation_incad(
        acadapp, acadmsp, Global_Scale, XY_range_dict, if_ZD, ganggz_single_nodes, abutmt_ganggz_single_nodes, 
        DM_Annotation_Dimension_Xdict1, DM_Annotation_Dimension_Ydict1, DM_Annotation_Mleader_Potdict1, 
        DM_Annotation_Dimension_Xdict2, DM_Annotation_Dimension_Ydict2, DM_Annotation_Mleader_Potdict2,   
    )
    # 绘制平面图
    PM_Annotation_Dimension_Xdict = {'贝雷':[], '分配梁':[], '分配梁1':[], '分配梁2':[], '钢管桩':[]}
    PM_Annotation_Dimension_Ydict = {'贝雷':[], '分配梁':[], '分配梁1':[], '分配梁2':[], '钢管桩':[], '制动桩':[]}
    PM_Annotation_Mleader_Potdict = {'贝雷':[], '分配梁':[], '分配梁1':[], '分配梁2':[], '钢管桩':[], '制动桩':[]}
    PM_Annotation_Dimension_Xdict['贝雷'].extend(LM_Annotation_Dimension_Xdict['贝雷'])
    PM_Annotation_Dimension_Xdict['钢管桩'].extend(LM_Annotation_Dimension_Xdict['钢管桩'])
    draw_Trestle_PM_incad(
        acadmsp, CAD_O_point, XY_range_dict, Global_Scale, if_ZD, distribute_section_per_meter, zd_f1_section_weight_per_meter, zd_f2_section_weight_per_meter, 
        xl_single_nodes, xl_section_cad, xl_section_name, 
        bailey_single_nodes, 
        fenpei_single_nodes, distribute_section_cad, distribute_section_name, 
        ganggz_single_nodes, gz_section_cad, gz_section_name,
        abutmt_fenpei1_single_nodes, zd_f1_section_cad, zd_f1_section_name, 
        abutmt_fenpei2_single_nodes, zd_f2_section_cad, zd_f2_section_name, 
        abutmt_ganggz_single_nodes, zd_section_cad, zd_section_name,     
        PM_Annotation_Dimension_Xdict, PM_Annotation_Dimension_Ydict, PM_Annotation_Mleader_Potdict, Material_dict, 
    )
    # 绘制平面图标注
    draw_Trestle_PM_Annotation_incad(
        acadapp, acadmsp, Global_Scale, XY_range_dict, if_ZD, 
        PM_Annotation_Dimension_Xdict, PM_Annotation_Dimension_Ydict, PM_Annotation_Mleader_Potdict, 
    )
    # 绘制材料表
    draw_Material_List_in_cad(
        acadmsp, CAD_O_point, Global_Scale, section_weight_per_meter_lst, 
        XY_range_dict, PM_Annotation_Dimension_Xdict, PM_Annotation_Dimension_Ydict, Material_dict,
    )  
    # 绘制图框和附注
    draw_frame_notes_in_cad(acadmsp, Global_Scale, LM_Annotation_Dimension_Xdict, LM_Annotation_Dimension_Ydict, PM_Annotation_Dimension_Xdict, PM_Annotation_Dimension_Ydict)

