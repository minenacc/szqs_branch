# 1. 标准库
import math

# 3. 本地模块
from General.Midas import MidasAPI, API_DIST_FORCE_Unit, Picture_Beamstress, Picture_Deformed, Value_BeamStress, Value_Deformed, Picture_BeamForce, Value_BeamForce, Picture_Reaction, Value_Reaction 


# 获取模型的所有节点、单元、截面、结构组的Jason格式
def get_node_elem_sect_grup_dict():

    # 设置单位默认为N和mm
    API_DIST_FORCE_Unit("N", 'MM')

    NODE_dict_res = MidasAPI("GET", "/db/NODE") # 节点信息
    ELEM_dict_res = MidasAPI("GET", "/db/ELEM") # 单元信息
    SECT_dict_res = MidasAPI("GET", "/db/SECT") # 截面信息
    get_structure_group_res = MidasAPI("GET", "/db/GRUP")# 组信息

    return [NODE_dict_res, ELEM_dict_res, SECT_dict_res, get_structure_group_res]


# 后处理——小肋
def Post_processing_xiaolei(folder_path):
    API_DIST_FORCE_Unit("N", 'MM')# 小肋应力和变形均为N和MM

    # 小肋应力截图
    Picture_Beamstress(folder_path, "小肋组合应力图.jpg", ["小肋"], "基本组合", "Combined")
    Picture_Beamstress(folder_path, "小肋剪应力图.jpg", ["小肋"], "基本组合", "Ssz")

    # 小肋变形截图
    Picture_Deformed(folder_path, "小肋竖向变形图.jpg", ["小肋"], "标准组合", "Dz")

    # 小肋应力计算
    xiaolei_stress_value_lst = Value_BeamStress("xiaolestress", "小肋", "基本组合(CB)")
    xiaole_max_cb_value = max(xiaolei_stress_value_lst['xiaolestress']['DATA'], key=lambda x: abs(float(x[3])))# 最大组合应力所在的结果表
    xiaole_max_shear_value = max(xiaolei_stress_value_lst['xiaolestress']['DATA'], key=lambda x: abs(float(x[2])))# 最大剪应力所在的结果表
    xiaole_max_cb = xiaole_max_cb_value[3]
    xiaole_max_shear = xiaole_max_shear_value[2]
    print(xiaole_max_cb_value)
    print(xiaole_max_shear_value)

    # 小肋变形计算
    xiaolei_deformed_value_lst = Value_Deformed("xiaole_displacement", "小肋", "标准组合(CB)")
    xiaole_max_Dz_value = max(xiaolei_deformed_value_lst['xiaole_displacement']['DATA'], key=lambda x: abs(float(x[2])))# 最大变形所在的结果表
    xiaole_max_Dz = xiaole_max_Dz_value[2]
    print(xiaole_max_Dz_value)

    return [xiaole_max_cb, xiaole_max_shear, xiaole_max_Dz]


# 后处理——贝雷梁
def Post_processing_bailey(folder_path):
    API_DIST_FORCE_Unit("KN", 'MM')# 贝雷内力为KN, 变形为mm

    # 贝雷内力截图
    Picture_BeamForce(folder_path, "贝雷弦杆轴力图.jpg", ["贝雷梁弦杆"], "标准组合", "Fx")
    Picture_BeamForce(folder_path, "贝雷竖杆轴力图.jpg", ["贝雷梁竖杆"], "标准组合", "Fx")
    Picture_BeamForce(folder_path, "贝雷斜杆轴力图.jpg", ["贝雷梁斜杆"], "标准组合", "Fx")

    # 贝雷变形截图
    Picture_Deformed(folder_path, "贝雷竖向变形图.jpg", ["贝雷梁弦杆", "贝雷梁竖杆", "贝雷梁斜杆"], "标准组合", "Dz")

    # 贝雷内力计算
    bailey_xiangan_zhouli_value_lst = Value_BeamForce("baileyforce", "贝雷梁弦杆", "标准组合(CB)")
    bailey_max_xiangan_zhouli_value = max(bailey_xiangan_zhouli_value_lst['baileyforce']['DATA'], key=lambda x: abs(float(x[2])))# 最大弦杆轴力所在的结果表
    bailey_max_xiangan_zhouli = bailey_max_xiangan_zhouli_value[2]
    print(bailey_max_xiangan_zhouli_value)
    
    bailey_shugan_zhouli_value_lst = Value_BeamForce("baileyforce", "贝雷梁竖杆", "标准组合(CB)")
    bailey_max_shugan_zhouli_value = max(bailey_shugan_zhouli_value_lst['baileyforce']['DATA'], key=lambda x: abs(float(x[2])))# 最大竖杆轴力所在的结果表
    bailey_max_shugan_zhouli = bailey_max_shugan_zhouli_value[2]
    print(bailey_max_shugan_zhouli_value)

    bailey_xiegan_zhouli_value_lst = Value_BeamForce("baileyforce", "贝雷梁斜杆", "标准组合(CB)")
    bailey_max_xiegan_zhouli_value = max(bailey_xiegan_zhouli_value_lst['baileyforce']['DATA'], key=lambda x: abs(float(x[2])))# 最大斜杆轴力所在的结果表
    bailey_max_xiegan_zhouli = bailey_max_xiegan_zhouli_value[2]
    print(bailey_max_xiegan_zhouli_value)

    # 贝雷变形计算
    bailey_deformed_value_lst = Value_Deformed("bailey_displacement", "贝雷梁弦杆", "标准组合(CB)")
    bailey_max_Dz_value = max(bailey_deformed_value_lst['bailey_displacement']['DATA'], key=lambda x: abs(float(x[2])))# 最大变形所在的结果表
    bailey_max_Dz = bailey_max_Dz_value[2]
    print(bailey_max_Dz_value)

    return [bailey_max_xiangan_zhouli, bailey_max_shugan_zhouli, bailey_max_xiegan_zhouli, bailey_max_Dz]


# 后处理——分配梁
def Post_processing_fenpeiliang(folder_path):
    API_DIST_FORCE_Unit("N", 'MM')

    # 分配梁应力截图
    Picture_Beamstress(folder_path, "分配梁组合应力图.jpg", ["分配梁"], "基本组合", "Combined")
    Picture_Beamstress(folder_path, "分配梁剪应力图.jpg", ["分配梁"], "基本组合", "Ssz")

    # 分配梁变形截图
    Picture_Deformed(folder_path, "分配梁竖向变形图.jpg", ["分配梁"], "标准组合", "Dz")

    # 分配梁应力计算
    fenpeiliang_stress_value_lst = Value_BeamStress("fenpeiliangstress", "分配梁", "基本组合(CB)")
    fenpeiliang_max_cb_value = max(fenpeiliang_stress_value_lst['fenpeiliangstress']['DATA'], key=lambda x: abs(float(x[3])))# 最大组合应力所在的结果表
    fenpeiliang_max_shear_value = max(fenpeiliang_stress_value_lst['fenpeiliangstress']['DATA'], key=lambda x: abs(float(x[2])))# 最大剪应力所在的结果表
    fenpeiliang_max_cb = fenpeiliang_max_cb_value[3]
    fenpeiliang_max_shear = fenpeiliang_max_shear_value[2]
    print(fenpeiliang_max_cb_value)
    print(fenpeiliang_max_shear_value)

    # 分配梁变形计算
    fenpeiliang_deformed_value_lst = Value_Deformed("fenpeiliang_displacement", "分配梁", "标准组合(CB)")
    fenpeiliang_max_Dz_value = max(fenpeiliang_deformed_value_lst['fenpeiliang_displacement']['DATA'], key=lambda x: abs(float(x[2])))# 最大变形所在的结果表
    fenpeiliang_max_Dz = fenpeiliang_max_Dz_value[2]
    print(fenpeiliang_max_Dz_value)

    return [fenpeiliang_max_cb, fenpeiliang_max_shear, fenpeiliang_max_Dz]


# 后处理——联结系
def Post_processing_lianjiexi(folder_path):
    API_DIST_FORCE_Unit("N", 'MM')

    # 联结系应力截图
    Picture_Beamstress(folder_path, "联结系组合应力图.jpg", ["联结系"], "基本组合", "Combined")
    Picture_Beamstress(folder_path, "联结系剪应力图.jpg", ["联结系"], "基本组合", "Ssz")

    # 联结系应力计算
    lianjiexi_stress_value_lst = Value_BeamStress("lianjiexistress", "联结系", "基本组合(CB)")
    lianjiexi_max_cb_value = max(lianjiexi_stress_value_lst['lianjiexistress']['DATA'], key=lambda x: abs(float(x[3])))# 最大组合应力所在的结果表
    lianjiexi_max_shear_value = max(lianjiexi_stress_value_lst['lianjiexistress']['DATA'], key=lambda x: abs(float(x[2])))# 最大剪应力所在的结果表
    lianjiexi_max_cb = lianjiexi_max_cb_value[3]
    lianjiexi_max_shear = lianjiexi_max_shear_value[2]
    print(lianjiexi_max_cb_value)
    print(lianjiexi_max_shear_value)

    return [lianjiexi_max_cb, lianjiexi_max_shear]


# 后处理——钢管桩
#查询钢管桩信息，包括直径、壁厚、桩长
def gangguan_info(group_dict, node_dict, elem_dict, sect_dict, elem):

    # #判断当前坐标系
    # unit_res = MidasAPI("GET", "/db/UNIT")
    # F = unit_res['UNIT']['1']['FORCE']
    # D = unit_res['UNIT']['1']['DIST']
    # print("力单位: ", F)
    # print("距离单位: ", D)
    # if F == 'KN' and D == 'M':
    #     Length = Length*1000

    #查询直径和壁厚
    sec_type_num = elem_dict['ELEM'][str(elem)]['SECT'] # 截面编号
    # sec_name = sect_dict['SECT'][str(sec_type_num)]['SECT_NAME'] # 钢管桩名字格式必须为钢管桩820x10
    # match = re.findall(r'\d+', sec_name)
    # if match:
    #     Diameter = match[0]  # 直径
    #     Thickness = match[1]  # 壁厚   
    sec_info = sect_dict['SECT'][str(sec_type_num)]['SECT_BEFORE']['SECT_I']['vSIZE']
    Diameter = round(sec_info[0], 3)
    Thickness = round(sec_info[1], 3)

    # print(sec_type_num, sec_name, match)

    #查询该单元号所在的结构组的全部单元
    elem_list = [member['E_LIST'] for member in group_dict if int(elem) in member['E_LIST']]

    #查询单元号对应的节点号
    node_list = []
    [node_list.append([elem_dict['ELEM'][str(member)]['NODE'][0],elem_dict['ELEM'][str(member)]['NODE'][1]]) for member in elem_list[0]]
    node_list = [item for sublist in node_list for item in sublist]
    
    #得到所有节点的z坐标
    node_z = []
    [node_z.append(node_dict['NODE'][str(member)]['Z']) for member in node_list]
    node_z.sort()
    Length = abs(node_z[0]-node_z[-1])
    
    #生成结果字典
    res_lst = {
        "LIST":{
            # "U": F+D,# 单位
            "D": Diameter,# 直径
            "T": Thickness,# 壁厚
            "L": Length,# 桩长
        }
    }
    #返回结果列表
    return res_lst


def Post_processing_gangguanzhuang(folder_path, NODE_dict_res, ELEM_dict_res, SECT_dict_res, get_structure_group_res):
    # 挑出结构组中所有的"钢管桩N"形式的结构组的字典
    structure_group_gangguan = [dict_res for dict_res in get_structure_group_res['GRUP'].values() if '钢管桩' in dict_res['NAME'] and dict_res['NAME'] != '钢管桩']
    # print(structure_group_gangguan)

    API_DIST_FORCE_Unit("KN", 'M')

    # 钢管桩内力截图
    Picture_BeamForce(folder_path, "钢管桩轴力图.jpg", ["钢管桩"], "基本组合", "Fx")
    Picture_BeamForce(folder_path, "钢管桩弯矩图.jpg", ["钢管桩"], "基本组合", "Mz")

    # 钢管桩内力计算, 读取的一定是KN和M, 见函数 Value_BeamForce 中的单位定义
    gangguan_zhouli_value_lst = Value_BeamForce("gangguanforce", "钢管桩", "基本组合(CB)")
    gangguan_max_zhouli_value = max(gangguan_zhouli_value_lst['gangguanforce']['DATA'], key=lambda x: abs(float(x[2])))# 最大弦杆轴力所在的结果表
    print(gangguan_max_zhouli_value)

    gangguan_max_zhouli_ele = gangguan_max_zhouli_value[1]# 最大轴力所在的单元
    gangguan_max_zhouli = gangguan_max_zhouli_value[2]# 最大轴力
    gangguan_max_zhouli_My = gangguan_max_zhouli_value[3]# 最大轴力所在单元的My弯矩
    gangguan_max_zhouli_Mz = gangguan_max_zhouli_value[4]# 最大轴力所在单元的Mz弯矩
    print(gangguan_max_zhouli_ele, gangguan_max_zhouli, gangguan_max_zhouli_My, gangguan_max_zhouli_Mz)

    # 读取到的截面参数是来自于 SECT_dict_res , 所以单位取决于执行获取 SECT_dict_res 命令时的单位, 和当前坐标系无关, 桩长同理, 读取的是节点字典 NODE_dict_res
    gangguan_sec_info_lst = gangguan_info(structure_group_gangguan, NODE_dict_res, ELEM_dict_res, SECT_dict_res, gangguan_max_zhouli_ele)
    Diameter = gangguan_sec_info_lst["LIST"]["D"]# 对应最大轴力所在的钢管的直径
    Thickness = gangguan_sec_info_lst["LIST"]["T"]# 对应最大轴力所在的钢管的壁厚
    Length = gangguan_sec_info_lst["LIST"]["L"]# 对应最大轴力所在的钢管的桩长
    print(Diameter, Thickness, Length)

    return [gangguan_max_zhouli_ele, gangguan_max_zhouli, gangguan_max_zhouli_My, gangguan_max_zhouli_Mz, Diameter, Thickness, Length]


# 后处理——反力
def Post_processing_reaction(folder_path):
    API_DIST_FORCE_Unit("KN", 'M')

    # 反力和弯矩图
    Picture_Reaction(folder_path, "基础竖向反力图.jpg", ["桩底反力"], "标准组合", "Fz")
    Picture_Reaction(folder_path, "基础弯矩图.jpg", ["桩底反力"], "标准组合", "Mxyz")

    # 反力计算
    reaction_value_lst = Value_Reaction("fanli", "桩底反力", "标准组合(CB)")
    # reaction_value_lst['fanli']['DATA'] 的值格式为: [index, nodeid, load_comb, Fx, Fy, Fz, Mx, My, Mz]

    # 最大竖向力
    reaction_max_Fz_value = max(reaction_value_lst['fanli']['DATA'], key=lambda x: abs(float(x[5])))# 最大竖向力所在的结果表
    print(reaction_max_Fz_value)
    # 最大水平力
    reaction_max_Fxy_value = max(reaction_value_lst['fanli']['DATA'], key=lambda x: abs(math.sqrt(float(x[3])**2 + float(x[4])**2)))# 最大水平力所在的结果表
    print(reaction_max_Fxy_value)
    # 最大弯矩Mx
    reaction_max_Mx_value = max(reaction_value_lst['fanli']['DATA'], key=lambda x: abs(float(x[6])))# 最大弯矩所在的结果表
    print(reaction_max_Mx_value)
    # 最大弯矩My
    reaction_max_My_value = max(reaction_value_lst['fanli']['DATA'], key=lambda x: abs(float(x[7])))# 最大弯矩所在的结果表
    print(reaction_max_My_value)

    return [reaction_max_Fz_value, reaction_max_Fxy_value, reaction_max_Mx_value, reaction_max_My_value]

