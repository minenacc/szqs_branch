# 3. 本地模块
from General.DataUtils import SET_CLIP_STRING
from General.Midas     import MidasAPI, num_lst_to_num_group

#=====================================================================================================================================================================

# 根据选择获取结构单元号
def get_select_node_element():
    # 闭包传递API获取的select字典
    global shared_select_dict
    # 获取字典
    select_dict = MidasAPI("GET" , "/VIEW/SELECT")
    shared_select_dict = select_dict
    # 获取单元列表
    node_lst = select_dict['SELECT']['NODE_LIST']
    elem_lst = select_dict['SELECT']['ELEM_LIST']

    return [node_lst, elem_lst]
# print(get_select_node_element())

#=====================================================================================================================================================================

# 在对话框中显示所选择的节点和单元列表
def get_node_ele_str(ent_var, param):
    try:
        # 获取单元列表转化的字符串
        node_lst, elem_lst = get_select_node_element()
        # 编辑框赋值
        if param == "NODE":
            # 处理字符串 x to y 格式
            node_str = num_lst_to_num_group(node_lst)
            ent_var.set(node_str)
            # 将内容同时复制到剪切板
            SET_CLIP_STRING(node_str)
            print("已将节点信息复制到剪切板中")
        if param == "ELEM":
            # 处理字符串 x to y 格式
            elem_str = num_lst_to_num_group(elem_lst)
            ent_var.set(elem_str)
            # 将内容同时复制到剪切板
            SET_CLIP_STRING(elem_str)
            print("已将单元信息复制到剪切板中")
    except:
        if param == "NODE":
            ent_var.set("节点获取失败, 请检查选择是否正确或API是否连接")
        if param == "ELEM":
            ent_var.set("单元获取失败, 请检查选择是否正确或API是否连接")

#=====================================================================================================================================================================

# 根据单元查询节点获取节点列表
def get_nodedict_from_elemdict(node_dict, elem_dict, elem_lst):
    # 包含单元、节点、坐标信息的字典{单元1：{节点1：坐标1，节点2：坐标2}}
    elem_node_coor_dict = {'elem_node_coor':{}}
    for x in elem_lst:
        node_lst = elem_dict['ELEM'][str(x)]['NODE']
        # 第一个点
        node1 = node_lst[0]
        coor1_dict = node_dict['NODE'][str(node1)]
        # 第二个点
        node2 = node_lst[1]
        coor2_dict = node_dict['NODE'][str(node2)]
        # 构成词典
        elem_node_coor_dict['elem_node_coor'][str(x)] = {str(node1) : [coor1_dict['X'], coor1_dict['Y'], coor1_dict['Z']], str(node2) : [coor2_dict['X'], coor2_dict['Y'], coor2_dict['Z']]}
    return elem_node_coor_dict
# print(NODE_dict_res)
# print(ELEM_dict_res)
# print(get_nodedict_from_elemdict(NODE_dict_res, ELEM_dict_res, [1,2,3,10,11,12]))


# 根据坐标是否首尾相连，将相连的单元放进同一张表中
def continuous_elements(elem_node_coor_dict):
    ele_lst = []
    for key1, value1 in elem_node_coor_dict['elem_node_coor'].items():
        ele_lst1 = []
        elem_id1 = key1
        node1_1 = value1[list(value1.keys())[1]]
        ele_lst1.append(elem_id1)
        # 寻找单元node1和node2相同的单元
        for key2, value2 in elem_node_coor_dict['elem_node_coor'].items():
            elem_id2 = key2
            if elem_id2 != elem_id1:
                node2_1 = value2[list(value2.keys())[0]]
                if node2_1 == node1_1:
                    ele_lst1.append(elem_id2)
        ele_lst.append(ele_lst1)
    # ele_lst = [['1', '3'], ['3', '2'], ['2'], ['10', '12'], ['12', '11'], ['11'], ['30']]
    # 根据表长度可以判断1、3单元和3、2单元首尾相连，2单元是末尾的单元
    # 先取出所有的末尾单元并组成表
    last_elem_lst = [x for x in ele_lst if len(x) == 1]
    else_elem_lst = [x for x in ele_lst if len(x) == 2]
    # 根据末尾单元倒推
    sorted_elem_lst = []
    for x in last_elem_lst:
        # 末尾单元号
        a = x[0]
        # 单次循环的结果表
        result = [a]
        # 副本
        else_elem_lst_copy = else_elem_lst.copy()
        while True:
            found = None
            # 查找包含a的子表
            for i, sublist in enumerate(else_elem_lst_copy):
                if str(a) in sublist:
                    found = else_elem_lst_copy.pop(i)  # 移除找到的子表
                    result.extend(found)  # 添加到结果
                    a = found[0]  # 更新a为找到的子表的第一个元素
                    break
            if not found:  # 如果没有找到匹配的子表
                break
        sorted_elem_lst.append(list(set(result)))
    # 把str变为int
    sorted_elem_lst = list(map(lambda sublist: list(map(int, sublist)), sorted_elem_lst))
    return sorted_elem_lst

# a = {'elem_node_coor': {'4': {'2': [0, -5600, 0], '3': [0, -5550, 0]}, '1': {'1': [0, -6450, 0], '2': [0, -5600, 0]}, '2': {'3': [0, -5550, 0], '4': [0, -4750, 0]}, '20': {'10': [0, -2530.7, 0], '11': [0, -2266.1, 0]}, '12': {'11': [0, -2266.1, 0], '12': [0, -2250, 0]}, '11': {'12': [0, -2250, 0], '13': [0, -2106.8, 0]}, '30': {'45': [0, -1, 0], '46': [0, -2, 0]}}}
# print(continuous_elements(a))


# 对于钢管桩和分配梁，生成一系列的结构组
def structure_group_divided(node_dict, elem_dict, select_dict):
    # 单元号列表
    elem_lst = select_dict['SELECT']['ELEM_LIST']
    # 包含单元、节点、坐标信息的字典{单元1：{节点1：坐标1，节点2：坐标2}}
    elem_node_coor_dict = get_nodedict_from_elemdict(node_dict, elem_dict, elem_lst)
    # 根据前一个单元的坐标2和后一个单元的坐标1判断2单元是否首位相连，若相连则放入同一张表中
    continuous_elements_lst = continuous_elements(elem_node_coor_dict)

    return continuous_elements_lst
# print(structure_group_divided(NODE_dict_res, ELEM_dict_res))


# 结构组单次input data form函数
def Structure_Group_input_data_form(new_group_num, new_group_name, node_lst, elem_lst):
    new_group = {
    "Assign": {
        new_group_num: {
            "NAME": "",
            "P_TYPE": 0,
            "N_LIST": [],
            "E_LIST": [],
            }
        }
    }
    # 如果是新组名，则添加新结构组
    new_group["Assign"][new_group_num]["NAME"] = new_group_name
    new_group["Assign"][new_group_num]["N_LIST"] = node_lst
    new_group["Assign"][new_group_num]["E_LIST"] = elem_lst
    MidasAPI("POST", "/db/GRUP", new_group)
    print("结构组\"{}\"成功".format(new_group_name))
    # info = MidasAPI("POST", "/db/GRUP", new_group)
    # if info == {'error': {'message': 'Key Already Exist'}}:
    #     print(info)
    #     print("结构组编号已存在, 请指定新的编号")
    # else:
    #     print("添加结构组\"{}\"成功".format(new_group_name))


# 获取边界组信息
def Boundary_Group_Get():
    Boundary_Group_namelst = MidasAPI("GET", "/db/BNGR")
    return Boundary_Group_namelst


# 判断边界组是否含有指定组名
def if_Boundary_Group_name_exist(Boundary_Group_name):
    Boundary_Group_data = Boundary_Group_Get()
    # print(Boundary_Group_data)
    # 存在指定组名则返回True
    if any(item['NAME'] == Boundary_Group_name for item in Boundary_Group_data['BNGR'].values()):
        return True
    else:
        return False


# 边界组单次input data form函数
def Boundary_Group_input_data_form(new_group_num, new_group_name):
    new_group = {
    "Assign": {
        new_group_num: {
            "NAME": new_group_name,
            "AUTOTYPE": 0,
            }
        }
    }
    # 如果是新组名，则添加新结构组
    info = MidasAPI("POST", "/db/BNGR", new_group)
    if info == {'error': {'message': 'Key Already Exist'}}:
        print(info)
        print("边界组编号已存在, 请指定新的编号")
    else:
        print("添加边界组\"{}\"成功".format(new_group_name))


# 获取当前结构组信息并返回编号列表和结构组名称
def Structure_Group_Get():
    group_lst = MidasAPI("GET", "/db/GRUP")['GRUP']
    group_index_lst = [int(x) for x in group_lst.keys()]
    group_name_lst = [x['NAME'] for x in group_lst.values()]
    return group_index_lst, group_name_lst


# 根据下拉列表的内容输出结构组的组名
def midas_group_output(combo_var, entry_var, node_dict, elem_dict, param):
    # 闭包传递API获取的select字典
    global shared_select_dict
    if node_dict[0] == [] or elem_dict[0] == []:
        print("还未获取模型的节点和单元信息")
    else:
        try:
            select_dict = shared_select_dict
            # 所选的全部节点
            node_lst = select_dict['SELECT']['NODE_LIST']
            # 所选的全部单元
            elem_lst = select_dict['SELECT']['ELEM_LIST']
            # 输出自定义的组名
            if combo_var.get() == "自定义":
                new_group_name = entry_var.get()
            # 输出固定的组名
            else:
                new_group_name = combo_var.get()
            group_index_lst, group_name_lst = Structure_Group_Get()
            # 结构组编号
            new_group_num = str(max(group_index_lst)+1)
            # 结构组名
            if new_group_name in group_name_lst:
                print("结构组{}已存在".format(new_group_name))
            else:
                # 结构组
                if param == "Structure Group":
                    # 写入结构组
                    Structure_Group_input_data_form(new_group_num, new_group_name, node_lst, elem_lst)
                    # 如果需要额外分组
                    if new_group_name == "钢管桩" or new_group_name == "分配梁":
                        # 计算得到的单元分组
                        elem_lsts = structure_group_divided(node_dict[0], elem_dict[0], select_dict)
                        # 循环时的初始计数
                        loop_group_num = str(int(new_group_num) + 1)
                        for i in range(len(elem_lsts)):
                            node_lst = []
                            elem_lst = elem_lsts[i]
                            # 循环时的结构组名
                            loop_group_name = new_group_name + str(i+1)
                            Structure_Group_input_data_form(loop_group_num, loop_group_name, node_lst, elem_lst)
                            loop_group_num = str(int(loop_group_num) + 1)
                if param == "Boundary Group":
                    # 写入结构组
                    Structure_Group_input_data_form(new_group_num, new_group_name, node_lst, [])
                    print("基础组{}在结构组中查看".format(new_group_name))
        except NameError:
            if param == "Structure Group":
                print("未指定需添加为结构组的单元")
            if param == "Boundary Group":
                print("未指定需添加为边界组的节点")
        except Exception as e:
            print("发生未知错误:", e)
        

#=====================================================================================================================================================================
