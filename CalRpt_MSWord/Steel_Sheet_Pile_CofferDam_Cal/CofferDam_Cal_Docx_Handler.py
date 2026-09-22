# 1. 标准库
import os
import re
import math

# 2. 第三方库
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle

from docx import Document
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage

# 3. 本地模块
from General.WordHandle import replace_placeholder_with_formula
from General.Geometry   import insert_zero_and_clip_negative, calculate_pressure_resultant
from General.Matplotlib import draw_insert_text, draw_arrow, draw_elastic_support, draw_level_mark
from General.DataUtils  import Divide_Solid_within_LayerThickness, intersect_segment_with_intervals, num_to_chinese_num, gamma_cal, gamma_sub_cal
from General.Formula    import cal_Ka, cal_Kp, cal_m, cal_Ks, cal_sigmak_solid, cal_sigmak_water, cal_pak_soild, cal_pak_water, cal_ppk_soild, cal_ppk_water, stability_coefficient
from General.Midas      import (
    MidasAPI, API_DIST_FORCE_Unit, Picture_Beamstress, Picture_Deformed, Value_BeamStress, Value_Deformed, Picture_BeamForce, Value_BeamForce, Value_BeamForce_Elem,
    Picture_Beamstress_Stage, Value_BeamStress_Stage, Value_Deformed_Stage, Picture_Deformed_Stage, Picture_BeamForce_Stage, Value_BeamForce_Stage, Value_BeamForce_Elem_Stage
)

from CalRpt_MSWord.Steel_Sheet_Pile_CofferDam_Cal.CofferDam_Cal_Standard_Formula import search_by_de, print_report_de, Cal_Ks_detail, JGJ_120_2012_4_2_1, JGJ_120_2012_4_2_2, JGJ_120_2012_4_2_4


def group_by_layer(items, matchstr):
    groups = {}
    for item in items:
        match = re.match(matchstr, item)
        if match:
            layer_num = int(match.group(1))
            groups.setdefault(layer_num, []).append(item)
    # 排序
    return [groups[num] for num in sorted(groups)]


# 土压力计算
def Active_Earth_Pressure(Solid_data_dict, Solid_Top_Level, Water_Top_Level, CofferDam_Bottom_Level, Ea_p0):
    '''
    Solid_data_dict Excel读取结果
    Solid_Top_Level 土层顶标高
    Water_Top_Level 水面顶标高
    CofferDam_Top_Level 围堰顶标高
    CofferDam_L 钢板桩桩长
    Ea_p0 基坑外初始附加应力
    '''
    # 土层的所有标高
    Solid_Dlst = [v['层厚'] for v in Solid_data_dict.values()] # 土层厚度 [1.3, 2.1, 16.2]
    Solid_Hlst = [Solid_Top_Level]
    [Solid_Hlst.append(round(Solid_Hlst[-1] - x,3)) for x in Solid_Dlst] # 土层标高 [0.0, -1.3, -3.4, -19.6]
    print('Solid_Hlst', Solid_Hlst)
    # 所有标高对应的土层信息
    Solid_dict = {}
    for i in range(len(Solid_Hlst[1:])):
        paramlst = Solid_data_dict[f'第{i+1}层土']
        soil_layer_name = paramlst['土层名称']
        unit_weight = paramlst['重度']
        cohesion = paramlst['黏聚力']
        friction = paramlst['内摩擦角']
        calc_method = paramlst['计算方法']
        Solid_dict[Solid_Hlst[1:][i]] = [unit_weight, cohesion, friction, calc_method, soil_layer_name]
    print('Solid_dict', Solid_dict)
    # Solid_dict = {-1.3: [18.84, 28.5, 15.4], -3.4: [18.05, 17.6, 14.3], -19.6: [16.38, 16.0, 14.0]}
    # 所有计算标高
    Active_Earth_Pressure_Hlst = sorted([x for x in Solid_Hlst + [CofferDam_Bottom_Level, Water_Top_Level] if x >= CofferDam_Bottom_Level], key = lambda x : x, reverse=True) # [0.0, -1.3, -1.5, -3.4, -8.8]
    print('Active_Earth_Pressure_Hlst:', Active_Earth_Pressure_Hlst)
    # 所有层的状态
    Solid_State = [] # 每一层土的含水状态
    Solid_Info = [] # 每一层的土的名称, 重度, 黏聚力, 内摩擦角
    for h in Active_Earth_Pressure_Hlst[1:]:
        # 判断是否是水层, 层底标高如果没有小于土层顶标高则全是水
        if h < Solid_Top_Level:
            # 对于土层判断是否含水
            layer_state = '无水' if h >= Water_Top_Level else '含水'
            Solid_State.append(layer_state)
            for key, value in Solid_dict.items():
                if h >= key:
                    Solid_Info.append(value)
                    break
        else:
            Solid_State.append('含水')
            Solid_Info.append([10, 0, 0, '/', '水'])
    print('Solid_State:', Solid_State) # 每一层土的含水状态 ['无水', '无水', '含水', '含水']
    print('Solid_Info:', Solid_Info) # 每一层的土的重度, 黏聚力, 内摩擦角 [[18.84, 28.5, 15.4], [18.05, 17.6, 14.3], [18.05, 17.6, 14.3], [16.38, 16.0, 14.0]]
    # 主动土压力计算
    soild_sigma0 = Ea_p0 # 初始竖向有效应力
    water_sigma0 = 0
    Solid_Sigma_lst = [soild_sigma0] # 土体有效竖向应力
    Water_Sigma_lst = [water_sigma0] # 水的竖向应力
    Solid_Pak_lst = [] # 土的水平力
    Water_Pak_lst = [] # 水的水平力
    for i in range(len(Active_Earth_Pressure_Hlst)-1):
        # 层顶标高
        Layer_Level1 = Active_Earth_Pressure_Hlst[i]
        # 层底标高
        Layer_Level2 = Active_Earth_Pressure_Hlst[i+1]
        # ------------------------------------------------------
        # 土层含水状态
        Layer_state = Solid_State[i]
        # 重度
        Solid_Layer_Gs = float(Solid_Info[i][0])
        # 黏聚力
        Solid_Layer_C = float(Solid_Info[i][1])
        # 内摩擦角
        Solid_Layer_Fai = float(Solid_Info[i][2])
        # 计算方法
        layer_method = Solid_Info[i][3]
        # 土层Ka
        Solid_Layer_Ka = cal_Ka(Solid_Layer_Fai)
        # ------------------------------------------------------
        if Layer_state == '含水':
            # 层顶层底竖向应力标准值——土
            solid_sigma1, solid_sigma2 = cal_sigmak_solid(soild_sigma0, Solid_Layer_Gs, Layer_Level1, Layer_Level2, 10)
            # 层顶层底竖向应力标准值——水
            water_sigma1, water_sigma2 = cal_sigmak_water(water_sigma0, Layer_Level1, Layer_Level2, 10)
        else:
            # 层顶层底竖向应力标准值——土
            solid_sigma1, solid_sigma2 = cal_sigmak_solid(soild_sigma0, Solid_Layer_Gs, Layer_Level1, Layer_Level2, 0) # 水重度取0
            # 层顶层底竖向应力标准值——水
            water_sigma1, water_sigma2 = [0, 0]
        soild_sigma0 = solid_sigma2 # 竖向应力更新
        water_sigma0 = water_sigma2 # 竖向应力更新
        # ------------------------------------------------------
        Solid_Sigma_lst.append(round(solid_sigma2,2))
        Water_Sigma_lst.append(round(water_sigma2,2))
        # 层顶层底主动土压力——土
        soild_pak1, soild_pak2 = cal_pak_soild(Solid_Layer_Ka, solid_sigma1, solid_sigma2, Solid_Layer_C) # 与水土是否分算无关
        # 层顶层底主动土压力——水
        water_pak1, water_pak2 = cal_pak_water(Solid_Layer_Ka, water_sigma1, water_sigma2, layer_method) # 已经考虑水土分算带来的影响
        
        Solid_Pak_lst.append([round(soild_pak1,2), round(soild_pak2,2)])
        Water_Pak_lst.append([round(water_pak1,2), round(water_pak2,2)])
    print(f'Solid_Sigma_lst:{Solid_Sigma_lst}')
    print(f'Water_Sigma_lst:{Water_Sigma_lst}')
    print(f'Solid_Pak_lst:{Solid_Pak_lst}')
    print(f'Water_Pak_lst:{Water_Pak_lst}')
    '''
    Active_Earth_Pressure_Hlst 土层加水位的绝对标高 [0.0, -1.3, -1.5, -3.4, -8.8]
    Solid_State 每一层的含水状态 ['无水', '无水', '含水', '含水']
    Solid_Info 每一层的土的信息 [['1-1', 18.84, 28.5, 15.4], ['2a-3-2', 18.05, 17.6, 14.3], ['2a-3-2', 18.05, 17.6, 14.3], ['3-1-1', 16.38, 16.0, 14.0]]
    Solid_Sigma_lst 每一个标高的土的竖向应力 [20.0, 44.49, 48.1, 63.4, 97.85]
    Water_Sigma_lst 每一个标高的水的竖向应力 [0, 0, 0, 19.0, 73.0]
    Solid_Pak_lst 每一层的层顶层底土压力 [[-31.8, -17.6], [-0.5, 1.7], [1.7, 10.94], [13.7, 34.7]]
    Water_Pak_lst 每一层的层顶层底水压力 [[0.0, 0.0], [0.0, 0.0], [0.0, 11.48], [11.6, 44.53]]
    对于合算，层顶层底主动土压力 = 土压力 和 水压力 的线性叠加
    '''
    return Active_Earth_Pressure_Hlst, Solid_State, Solid_Info, Solid_Sigma_lst, Water_Sigma_lst, Solid_Pak_lst, Water_Pak_lst


def Passive_Earth_Pressure(Solid_data_dict, Solid_Top_Level, Water_Top_Level, CofferDam_Bottom_Level, Ep_p0, Concrete_Bottom_Level, has_stage, excavate_stages, combo_dict):
    '''
    Excel_Data Excel读取结果
    Solid_Top_Level 土层顶标高
    Water_Top_Level 水面顶标高
    CofferDam_Top_Level 围堰顶标高
    CofferDam_L 钢板桩桩长
    Waler_Strut_Zlst 围檩和内支撑所在的绝对标高
    concrete_Bottom_Level 基坑底标高
    Ep_p0 基坑内初始附加应力
    Drawdown_height 安装围檩和内支撑时每次降水或开挖的高度
    Method_for_Pressures_var 土压力计算方法
    '''
    # 土层的所有标高
    Solid_Dlst = [v['层厚'] for v in Solid_data_dict.values()] # 土层厚度 [1.3, 2.1, 16.2]
    Solid_Hlst = [Solid_Top_Level]
    [Solid_Hlst.append(round(Solid_Hlst[-1] - x,3)) for x in Solid_Dlst] # 土层标高 [0.0, -1.3, -3.4, -19.6]
    print(f"土层标高: {Solid_Hlst}")
    Condition_Hlst = [max(vlst['基坑内土顶标高'], vlst['基坑内水面标高']) for vlst in excavate_stages] if has_stage else [Concrete_Bottom_Level]
    print(f"所有工况的计算标高: {Condition_Hlst}")
    # 所有标高对应的土层信息
    Solid_dict = {}
    for i in range(len(Solid_Hlst[1:])):
        paramlst = Solid_data_dict[f'第{i+1}层土']
        soil_layer_name = paramlst['土层名称']
        unit_weight = paramlst['重度']
        cohesion = paramlst['黏聚力']
        friction = paramlst['内摩擦角']
        calc_method = paramlst['计算方法']
        Solid_dict[Solid_Hlst[1:][i]] = [unit_weight, cohesion, friction, calc_method, soil_layer_name]
    print('Solid_dict', Solid_dict)
    # Solid_dict = {-1.3: [18.84, 28.5, 15.4], -3.4: [18.05, 17.6, 14.3], -19.6: [16.38, 16.0, 14.0]}
    # 所有计算标高
    # Passive_Earth_Pressure_Hlst_ALL = sorted([x for x in Solid_Hlst + [CofferDam_Bottom_Level, Water_Top_Level] + Condition_Hlst if x >= CofferDam_Bottom_Level], key = lambda x : x, reverse=True) # [0.0, -1.3, -1.5, -2.0, -3.4, -4.14, -8.8]
    # print(f"所有计算标高: {Passive_Earth_Pressure_Hlst_ALL}")
    # 每一种工况下的被动土压力计算
    Condition_Passive_Earth_Pressure_Hlst = [] # 所有工况的计算标高
    Condition_Solid_State = [] # 所有工况的土层含水状态
    Condition_Solid_Info = [] # 所有工况的土层计算信息
    Condition_Solid_Sigma_lst = [] # 所有工况的土体有效竖向应力
    Condition_Water_Sigma_lst = [] # 所有工况的水的竖向应力
    Condition_Solid_Ppk_lst = [] # 所有工况的土的水平力
    Condition_Water_Ppk_lst = [] # 所有工况的水的水平力
    for Hi, H in enumerate(Condition_Hlst):
        print('================')
        print(f'当z={H}m')
        if has_stage:  # 模型存在施工阶段
            current_solid_level = excavate_stages[Hi]['基坑内土顶标高']
            current_water_level = excavate_stages[Hi]['基坑内水面标高']
        else:  # 整体模型
            current_solid_level = H
            current_water_level = round(min(H - combo_dict['基坑底降水'], Water_Top_Level), 3)
        Passive_Earth_Pressure_Hlst = [x for x in sorted(set(Solid_Hlst + [CofferDam_Bottom_Level, H, current_solid_level, current_water_level]),reverse=True) if x <= H and x >= CofferDam_Bottom_Level] # [-2.0, -3.4, -4.14, -8.8]
        Condition_Passive_Earth_Pressure_Hlst.append(Passive_Earth_Pressure_Hlst)
        # 所有层的状态
        Solid_State = [] # 每一层土的含水状态
        Solid_Info = [] # 每一层的土的名称, 重度, 黏聚力, 内摩擦角
        for h in Passive_Earth_Pressure_Hlst[1:]:
            if h >= current_solid_level:
                # 层底标高高于土顶标高的部分，按水来计算
                Solid_State.append('含水')
                Solid_Info.append([10, 0, 0, '/', '水'])
            else: 
                # 对于土层判断是否含水
                layer_state = '无水' if h >= current_water_level else '含水'
                Solid_State.append(layer_state)
                for key, value in Solid_dict.items():
                    if h >= key:
                        Solid_Info.append(value)
                        break
                
        print('Passive_Earth_Pressure_Hlst:', Passive_Earth_Pressure_Hlst)
        print('当前土层顶标高', current_solid_level, ', 当前水面标高', current_water_level)
        print('Solid_State', Solid_State) # 每一层土的含水状态 ['含水', '含水', '含水']
        print('Solid_Info', Solid_Info) # 每一层的土的重度, 黏聚力, 内摩擦角 [[18.05, 17.6, 14.3], [16.38, 16.0, 14.0], [16.38, 16.0, 14.0]]
        Condition_Solid_State.append(Solid_State)
        Condition_Solid_Info.append(Solid_Info)
        # 被动土压力计算
        soild_sigma0 = Ep_p0 # 初始竖向有效应力
        water_sigma0 = 0
        Solid_Sigma_lst = [soild_sigma0] # 土体有效竖向应力
        Water_Sigma_lst = [water_sigma0] # 水的竖向应力
        Solid_Ppk_lst = [] # 土的水平力
        Water_Ppk_lst = [] # 水的水平力
        kp_lst = []
        for i in range(len(Passive_Earth_Pressure_Hlst)-1):
            # 层顶标高
            Layer_Level1 = Passive_Earth_Pressure_Hlst[i]
            # 层底标高
            Layer_Level2 = Passive_Earth_Pressure_Hlst[i+1]
            # ------------------------------------------------------
            # 土层含水状态
            Layer_state = Solid_State[i]
            # 重度
            Solid_Layer_Gs = float(Solid_Info[i][0])
            # 黏聚力
            Solid_Layer_C = float(Solid_Info[i][1])
            # 内摩擦角
            Solid_Layer_Fai = float(Solid_Info[i][2])
            # 计算方法
            layer_method = Solid_Info[i][3]
            # 土层Ka
            Solid_Layer_Kp = cal_Kp(Solid_Layer_Fai)
            # ------------------------------------------------------
            if Layer_state == '含水':
                # 层顶层底竖向应力标准值——土
                solid_sigma1, solid_sigma2 = cal_sigmak_solid(soild_sigma0, Solid_Layer_Gs, Layer_Level1, Layer_Level2, 10)
                # 层顶层底竖向应力标准值——水
                water_sigma1, water_sigma2 = cal_sigmak_water(water_sigma0, Layer_Level1, Layer_Level2, 10)
            else:
                # 层顶层底竖向应力标准值——土
                solid_sigma1, solid_sigma2 = cal_sigmak_solid(soild_sigma0, Solid_Layer_Gs, Layer_Level1, Layer_Level2, 0) # 水重度取0
                # 层顶层底竖向应力标准值——水
                water_sigma1, water_sigma2 = [0, 0]
            soild_sigma0 = solid_sigma2 # 竖向应力更新
            water_sigma0 = water_sigma2 # 竖向应力更新
            # ------------------------------------------------------
            Solid_Sigma_lst.append(round(solid_sigma2,2))
            Water_Sigma_lst.append(round(water_sigma2,2))
            # 层顶层底主动土压力——土
            soild_pak1, soild_pak2 = cal_ppk_soild(Solid_Layer_Kp, solid_sigma1, solid_sigma2, Solid_Layer_C) # 与水土是否分算无关
            # 层顶层底主动土压力——水
            water_pak1, water_pak2 = cal_ppk_water(Solid_Layer_Kp, water_sigma1, water_sigma2, layer_method) # 已经考虑水土分算带来的影响
            
            Solid_Ppk_lst.append([round(soild_pak1,2), round(soild_pak2,2)])
            Water_Ppk_lst.append([round(water_pak1,2), round(water_pak2,2)])
            kp_lst.append(Solid_Layer_Kp)
        print(f'Solid_Sigma_lst: {Solid_Sigma_lst}')
        print(f'Water_Sigma_lst: {Water_Sigma_lst}')
        print(f'Solid_Ppk_lst: {Solid_Ppk_lst}')
        print(f'Water_Ppk_lst: {Water_Ppk_lst}')
        Condition_Solid_Sigma_lst.append(Solid_Sigma_lst)
        Condition_Water_Sigma_lst.append(Water_Sigma_lst)
        Condition_Solid_Ppk_lst.append(Solid_Ppk_lst)
        Condition_Water_Ppk_lst.append(Water_Ppk_lst)
    print('================')
    '''
    Condition_Passive_Earth_Pressure_Hlst 每一种工况下土层加水位的绝对标高 [[-2.0, -3.4, -4.14, -8.8], [-4.14, -8.8]]
    Condition_Solid_State 每一种工况下每一层的含水状态 [['含水', '含水', '含水'], ['含水']]
    Condition_Solid_Info 每一种工况下每一层的土的信息 [[['2a-3-2', 18.05, 17.6, 14.3], ['3-1-1', 16.38, 16.0, 14.0], ['3-1-1', 16.38, 16.0, 14.0]], [['3-1-1', 16.38, 16.0, 14.0]]]
    Condition_Solid_Sigma_lst 每一种工况下每一个标高的土的竖向应力 [[0, 11.27, 15.99, 45.72], [0, 29.73]]
    Condition_Water_Sigma_lst 每一种工况下每一个标高的水的竖向应力 [[0, 14.0, 21.4, 68.0], [0, 46.6]]
    Condition_Solid_Ppk_lst 每一种工况下每一层的层顶层底土压力 [[[45.3, 63.96], [59.4, 67.15], [67.1, 115.85]], [[41.0, 89.65]]]
    Condition_Water_Ppk_lst 每一种工况下每一层的层顶层底水压力 [[[0.0, 23.18], [22.9, 35.05], [35.1, 111.38]], [[0.0, 76.33]]]
    每一种工况下，对于合算，层顶层底主动土压力 = 土压力 和 水压力 的线性叠加
    '''
    return Condition_Passive_Earth_Pressure_Hlst, Condition_Solid_State, Condition_Solid_Info, Condition_Solid_Sigma_lst, Condition_Water_Sigma_lst,  Condition_Solid_Ppk_lst, Condition_Water_Ppk_lst


def Matplotlib_Draw_Solid_Pressure(save_path, Pile_Length, CofferDam_Top_Level, CofferDam_Bottom_Level, H1, H2, Pa_Solid_depth, Pa_Solid_pressure, Pa_Water_depth, Pa_Water_pressure, Pp_Solid_depth, Pp_Solid_pressure, Pp_Water_depth, Pp_Water_pressure, Elastic_Support_depth, Waler_Strut_depth, Level_Mark_depth, Level_Mark_txt):
    # 数据：深度 (y) 和对应的压力值 (x)
    # 创建图形和轴
    plt.rcParams['font.sans-serif'] = ['SimHei'] # 使用黑体显示中文
    plt.rcParams['axes.unicode_minus'] = False # 解决负号 '-' 显示为方块的问题
    fig, ax = plt.subplots(figsize=(10, 5))
    # 隐藏所有坐标轴
    ax.axis('off')
    # 各图形缩放比例，以横坐标100，纵坐标10为基准创建的标准图形
    x_scale = (max(Pa_Solid_pressure)+max(Pp_Solid_pressure)+max(Pa_Water_pressure)+max(Pp_Water_pressure))/100
    y_scale = Pile_Length/10
    # 钢板桩矩形图块
    # ax.axvline(x=0, color='black', linewidth=2, label='挡墙') # 垂直线
    wall_thickness = 1*x_scale
    wall = Rectangle((0, CofferDam_Bottom_Level),          # 左下角坐标 (x, y)
                 wall_thickness,           # 宽度（墙厚）
                #  max(depth) - min(depth),  # 高度（墙高）
                 Pile_Length, # 高度（墙高）
                 linewidth=0.1,
                 edgecolor='black',
                 facecolor='lightgray',
                 label='钢板桩')
    ax.add_patch(wall)
    # 填充土压力区域
    # fill_betweenx 的 x 参数接受两个数组，分别表示填充的左右边界
    # 左边界设为 -pressure（因为压力在左侧），右边界为 0
    ax.fill_betweenx(Pa_Solid_depth, -np.array(Pa_Solid_pressure), 0, 
                    where=None, color='lightblue', alpha=0.6, label='主动土压力')
    ax.fill_betweenx(Pp_Solid_depth, +np.array(Pp_Solid_pressure), 0+wall_thickness, 
                    where=None, color='lightblue', alpha=0.6, label='被动土压力')
    # 绘制压力分布的外轮廓线
    ax.plot(-np.array(Pa_Solid_pressure), Pa_Solid_depth, color='blue', linewidth=1.5) # 主动土压力
    ax.plot(+np.array(Pp_Solid_pressure), Pp_Solid_depth, color='blue', linewidth=1.5) # 被动土压力
    # 添加标注（数字、箭头）
    # 标注压力值
    draw_insert_text(ax, x_scale, y_scale, Pa_Solid_pressure, Pa_Solid_depth, 'right')
    draw_insert_text(ax, x_scale, y_scale, Pp_Solid_pressure, Pp_Solid_depth, 'left')
    # for p, d in zip(Pp_Solid_pressure, Pp_Solid_depth):
    #     ax.text(+p, d, f'{p} kPa', ha='left', va='bottom', fontsize=7, bbox=dict(facecolor='white', edgecolor='none', alpha=0.7)) # 被动土压力, 左下对齐
    # 水压力
    if any(Pa_Water_pressure): # 有非0元素
        ax.fill_betweenx(Pa_Water_depth, -np.array(Pa_Water_pressure)-max(Pa_Solid_pressure)-9*x_scale, 0-max(Pa_Solid_pressure)-9*x_scale, where=None, color='lightblue', alpha=0.6, label='外侧水压力')
        ax.plot(-np.array(Pa_Water_pressure)-max(Pa_Solid_pressure)-9*x_scale, Pa_Water_depth, color='blue', linewidth=1.5) # 外侧水压力
        for p, d in zip(Pa_Water_pressure, Pa_Water_depth):
            ax.text(-p-max(Pa_Solid_pressure)-10*x_scale, d+0.2*y_scale, f'{p} kPa', ha='right', va='center', fontsize=7) # 外侧水压力, 右下对齐
    if any(Pp_Water_pressure): # 有非0元素
        ax.fill_betweenx(Pp_Water_depth, +np.array(Pp_Water_pressure)+max(Pp_Solid_pressure)+9*x_scale, 0+max(Pp_Solid_pressure)+9*x_scale, where=None, color='lightblue', alpha=0.6, label='基坑水压力')
        ax.plot(+np.array(Pp_Water_pressure)+max(Pp_Solid_pressure)+9*x_scale, Pp_Water_depth, color='blue', linewidth=1.5) # 基坑水压力
        for p, d in zip(Pp_Water_pressure, Pp_Water_depth):
            ax.text(+p+max(Pp_Solid_pressure)+10*x_scale, d+0.2*y_scale, f'{p} kPa', ha='left', va='center', fontsize=7) # 基坑水压力, 左下对齐
    # 标注方向箭头（例如在墙顶附近画一个水平箭头）
    draw_arrow(ax, -(max(Pa_Solid_pressure)+max(Pa_Water_pressure))/2, CofferDam_Top_Level + 0.5*y_scale, 'right', x_scale)
    ax.text(-(max(Pa_Solid_pressure)+max(Pa_Water_pressure))/2, CofferDam_Top_Level + 1*y_scale, '主动土压力', color='red', fontsize=7, ha='center')
    draw_arrow(ax, +(max(Pp_Solid_pressure)+max(Pp_Water_pressure))/2, CofferDam_Top_Level + 0.5*y_scale, 'left', x_scale)
    ax.text(+(max(Pp_Solid_pressure)+max(Pp_Water_pressure))/2, CofferDam_Top_Level + 1*y_scale, '被动土压力', color='red', fontsize=7, ha='center')
    # 地面线和基坑线
    ax.plot([wall_thickness, +max(Pp_Solid_pressure)+max(Pp_Water_pressure)+5*x_scale], [Pp_Solid_depth[0], Pp_Solid_depth[0]], color='brown', linestyle='--', linewidth=1, alpha=0.8) # 基坑线
    ax.plot([0, -max(Pa_Solid_pressure)-max(Pa_Water_pressure)-5*x_scale], [Pa_Solid_depth[0], Pa_Solid_depth[0]], color='gray', linestyle='--', linewidth=1, alpha=0.8) # 地面线
    ax.plot([0, -max(Pa_Solid_pressure)-max(Pa_Water_pressure)-5*x_scale], [Pa_Water_depth[0], Pa_Water_depth[0]], color='blue', linestyle='--', linewidth=1, alpha=0.8) # 水面线
    # 绘制弹性支座
    for h in Elastic_Support_depth:
        draw_elastic_support(ax, wall_thickness, h, x_scale, y_scale)
    # 绘制内支撑标志
    for h in Waler_Strut_depth:
        draw_support(ax, wall_thickness, h, x_scale, y_scale)
    # 绘制标高
    MAX_X = max(Pp_Solid_pressure)+max(Pp_Water_pressure)+20*x_scale
    for h,txt in zip(Level_Mark_depth, Level_Mark_txt):
        draw_level_mark(ax, MAX_X, h, txt, x_scale, y_scale)
    
    # plt.tight_layout()
    # plt.show() # show 和 savefig 不能同时激活
    '''
    dpi：分辨率，例如 dpi=300 用于高清打印
    bbox_inches='tight'：自动裁剪图形周围的空白区域，使保存的图片更紧凑
    pad_inches：与 bbox_inches='tight' 配合使用，设置边距
    transparent：是否设置透明背景（用于 PNG 等格式）
    '''
    plt.savefig(save_path, dpi=300, bbox_inches='tight', pad_inches=0.1)
    plt.close('all')  # 显式关闭所有图形，释放资源
    print('图片已在{}路径下保存'.format(save_path))






# 绘制内支撑图标
def draw_support(ax, start_x, start_y, x_scale, y_scale):
    # 三角形坐标
    x_lst1 = [0, 3, 3, 0]
    y_lst1 = [0, 0.3, -0.3, 0]
    x_lst1 = [i*0.3*x_scale+start_x for i in x_lst1]
    y_lst1 = [i*0.3*y_scale+start_y for i in y_lst1]
    ax.plot(x_lst1, y_lst1, linestyle='-', color='green', linewidth=0.7) # 三角形
    # 直线段坐标
    x_lst2 = [3, 6]
    y_lst2 = [0, 0]
    x_lst2 = [i*0.3*x_scale+start_x for i in x_lst2]
    y_lst2 = [i*0.3*y_scale+start_y for i in y_lst2]
    ax.plot(x_lst2, y_lst2, linestyle='-', color='green', linewidth=0.7) # 直线段





def Draw_Solid_Pressure(
        save_path, Solid_Top_Level, Water_Top_Level, CofferDam_Top_Level, CofferDam_L, concrete_Bottom_Level, Method_for_Pressures_var, Waler_Strut_Zlst, Condition_m_method_Hlst, 
        Active_Earth_Pressure_Hlst, Active_Earth_Pressure_Solid_Pak_lst, Active_Earth_Pressure_Water_Pak_lst, 
        Passive_Earth_Pressure_Hlst, Passive_Earth_Pressure_Solid_Ppk_lst, Passive_Earth_Pressure_Water_Ppk_lst
    ):
    Condition_num = -1
    CofferDam_Bottom_Level = CofferDam_Top_Level-CofferDam_L # 桩底标高
    Excavation_Depth = max(Solid_Top_Level, Water_Top_Level) - concrete_Bottom_Level # 基坑深度
    Exposure_Height = CofferDam_Top_Level - max(Solid_Top_Level, Water_Top_Level) # 钢板桩露出地面或水面高度
    Pa_Solid_depth = [[Active_Earth_Pressure_Hlst[i], Active_Earth_Pressure_Hlst[i+1]] for i in range(len(Active_Earth_Pressure_Hlst)-1)] # 各土层和水位所在层的主动土压力的标高
    Pa_Solid_depth = [y for x in Pa_Solid_depth for y in x]
    Pa_Water_depth = [Water_Top_Level, CofferDam_Bottom_Level] # 水深
    Pp_Solid_depth = [[Passive_Earth_Pressure_Hlst[Condition_num][i], Passive_Earth_Pressure_Hlst[Condition_num][i+1]] for i in range(len(Passive_Earth_Pressure_Hlst[Condition_num])-1)] # 各土层和水位所在层的被动土压力的标高
    Pp_Solid_depth = [y for x in Pp_Solid_depth for y in x]
    Pp_Water_depth = [Passive_Earth_Pressure_Hlst[Condition_num][0], Passive_Earth_Pressure_Hlst[Condition_num][-1]] # 水深
    m_method_Hlst = Condition_m_method_Hlst[Condition_num]
    Elastic_Support_depth = [(m_method_Hlst[i] + m_method_Hlst[i+1])/2 for i in range(len(m_method_Hlst)-1)] # 节点弹性支承高度
    Waler_Strut_depth = Waler_Strut_Zlst # 围檩高度
    Level_Mark_depth = sorted([CofferDam_Top_Level, CofferDam_Bottom_Level, concrete_Bottom_Level] + Waler_Strut_Zlst, reverse = True) # 需要标高的高度: 围堰顶标高、围檩高度、基坑底标高、围堰底标高
    Level_Mark_txt = [f'{x:.3f}' if x<0 else f'+{x:.3f}' for x in Level_Mark_depth] # 标高文本
    if Method_for_Pressures_var == '水土分算':
        Pa_Solid_pressure = [x for lst in Active_Earth_Pressure_Solid_Pak_lst for x in lst]
        Pa_Water_pressure = [Active_Earth_Pressure_Water_Pak_lst[0][0], Active_Earth_Pressure_Water_Pak_lst[-1][-1]]
        Pp_Solid_pressure = [x for lst in Passive_Earth_Pressure_Solid_Ppk_lst[Condition_num] for x in lst]
        Pp_Water_pressure = [Passive_Earth_Pressure_Water_Ppk_lst[Condition_num][0][0], Passive_Earth_Pressure_Water_Ppk_lst[Condition_num][-1][-1]]
    else:
        Pa_Solid_pressure = [round(x+y,2) for lst1, lst2 in zip(Active_Earth_Pressure_Solid_Pak_lst, Active_Earth_Pressure_Water_Pak_lst) for x,y in zip(lst1,lst2)]
        Pa_Water_pressure = [0, 0]
        Pp_Solid_pressure = [round(x+y,2) for lst1, lst2 in zip(Passive_Earth_Pressure_Solid_Ppk_lst[Condition_num], Passive_Earth_Pressure_Water_Ppk_lst[Condition_num]) for x,y in zip(lst1,lst2)]
        Pp_Water_pressure = [0, 0]
    Pa_Solid_depth, Pa_Solid_pressure = insert_zero_and_clip_negative(Pa_Solid_depth, Pa_Solid_pressure)
    Pp_Solid_depth, Pp_Solid_pressure = insert_zero_and_clip_negative(Pp_Solid_depth, Pp_Solid_pressure)
    Matplotlib_Draw_Solid_Pressure(save_path, CofferDam_L, CofferDam_Top_Level, CofferDam_Bottom_Level, Excavation_Depth, Exposure_Height, Pa_Solid_depth, Pa_Solid_pressure, Pa_Water_depth, Pa_Water_pressure, Pp_Solid_depth, Pp_Solid_pressure, Pp_Water_depth, Pp_Water_pressure, Elastic_Support_depth, Waler_Strut_depth, Level_Mark_depth, Level_Mark_txt)

def soild_pressure_cal(doc_path, save_path, Method_for_Pressures_var, 
                       Solid_Top_Level, Water_Top_Level, CofferDam_Bottom_Level, concrete_Bottom_Level, Ea_p0,
                       Active_Earth_Pressure_Hlst, Active_Earth_Pressure_Solid_Info, Active_Earth_Pressure_Solid_State, 
                       Active_Earth_Pressure_Solid_Sigma_lst, Active_Earth_Pressure_Solid_Pak_lst, Active_Earth_Pressure_Water_Sigma_lst, Active_Earth_Pressure_Water_Pak_lst, 
                       Passive_Earth_Pressure_Hlst, Passive_Earth_Pressure_Solid_Info, Passive_Earth_Pressure_Solid_State, 
                       Passive_Earth_Pressure_Solid_Sigma_lst, Passive_Earth_Pressure_Solid_Ppk_lst, Passive_Earth_Pressure_Water_Sigma_lst, Passive_Earth_Pressure_Water_Ppk_lst,
                       excavate_stages, combo_dict, acc = 3
                       ):
    # 路径处理
    template_path = os.path.join(doc_path, 'Steel_Sheet_Pile_CofferDam_cal_template_solid_pressure.docx')
    save_path = os.path.join(save_path, 'solid_pressure.docx')
    # 根据施工阶段或基坑底降水信息获取基坑底水位
    # excavate_stages = [{'工况序号': xx,'工况类型': xx,'基坑内土顶标高': xx,'基坑内水面标高': xx}, {...}]
    # combo_dict = {'超挖深度': xx,'基坑底降水': xx,'开挖面降水': xx}
    water_level_at_excavation_bottom = round(min(concrete_Bottom_Level, Water_Top_Level), acc)  # 默认值为基坑底标高和基坑外水位标高的较小值
    water_level_condition_lst = []
    if combo_dict:
        water_level_at_excavation_bottom = round(min(concrete_Bottom_Level - float(combo_dict['基坑底降水']),  Water_Top_Level), acc) # 整体模型时, 基坑底水位
        water_level_condition_lst.append(water_level_at_excavation_bottom)
    else:
        for valuelst in excavate_stages: # 施工阶段时, 基坑底水位
            water_level_condition_lst.append(valuelst['基坑内水面标高'])
            if abs(valuelst['基坑内土顶标高'] - concrete_Bottom_Level) < 0.0001:
                water_level_at_excavation_bottom = valuelst['基坑内水面标高']
    print('各工况下水位标高:', water_level_condition_lst)
    print('基坑底水位计算结果:', water_level_at_excavation_bottom)

    if Method_for_Pressures_var == '水土分算':
        # 主动土压力
        active_pressure = []
        formula_registry = {}
        # 基坑外水压力公式
        Water_H1 = round(Water_Top_Level-CofferDam_Bottom_Level,acc)
        Water_P1 = round(10 * Water_H1,acc)
        Water_Pressure1 = f'P_1={Water_H1}' + r'\times10' + f'={Water_P1}kPa'
        water_pressure1_key = '{Separate_Active_Water_Pressure}'
        formula_registry[water_pressure1_key] = Water_Pressure1
        # 初始化一个显示用的序号计数器
        Solid_number = 1
        for i in range(len(Active_Earth_Pressure_Hlst)-1):
            Solid_Layer_Method = Active_Earth_Pressure_Solid_Info[i][3]
            Solid_Layer_Name = Active_Earth_Pressure_Solid_Info[i][4]
            Solid_H1 = Active_Earth_Pressure_Hlst[i]
            Solid_H2 = Active_Earth_Pressure_Hlst[i+1]
            if Solid_Layer_Name == '水':
                continue
            Solid_D1 = round(Solid_H1 - Solid_H2, acc)
            Solid_if_Water = '水下' if Active_Earth_Pressure_Solid_State[i] == '含水' else '水上'
            active_pressure.append({
                'index':Solid_number,
                'Separate_Active_Solid_number':Solid_number,
                'Separate_Active_Solid_Layer_Name':Solid_Layer_Name,
                'Separate_Active_Solid_if_Water':Solid_if_Water,
                'Separate_Active_Solid_Method_for_Pressures_var':Solid_Layer_Method,
                'Separate_Active_Solid_H1':Solid_H1,
                'Separate_Active_Solid_H2':Solid_H2,
                'Separate_Active_Solid_D1':Solid_D1  
            })
            # 公式计算
            Gs = round(float(Active_Earth_Pressure_Solid_Info[i][0])-10, acc) if Solid_if_Water == '水下' else round(float(Active_Earth_Pressure_Solid_Info[i][0]), acc) # 重度
            C = round(float(Active_Earth_Pressure_Solid_Info[i][1]), acc) # 黏聚力
            Fai = round(float(Active_Earth_Pressure_Solid_Info[i][2]), acc) # 内摩擦角
            Ka = cal_Ka(Fai) # Ka
            sigma_soil_top = round(Active_Earth_Pressure_Solid_Sigma_lst[i], acc)
            sigma_soil_bot = round(Active_Earth_Pressure_Solid_Sigma_lst[i+1], acc)
            pak_soil_top = round(Active_Earth_Pressure_Solid_Pak_lst[i][0], acc)
            pak_soil_bot = round(Active_Earth_Pressure_Solid_Pak_lst[i][1], acc)
            # 公式填充
            sigma1 = round(sigma_soil_top, acc)
            sigma2 = round(sigma_soil_bot, acc)
            pak1 = round(pak_soil_top, acc)
            pak2 = round(pak_soil_bot, acc)
            Solid_Sigma1 = r'\sigma_{\mathrm{ak}}=' + f'{sigma1}KPa'
            Solid_Pressure1 = r'p_{\mathrm{ak}}=' + str(sigma1) + r'\times' + f'{Ka}-2' + r'\times' + str(C) + r'\times\sqrt{' + str(Ka) + r'}=' + f'{pak1}kPa'
            Solid_Sigma2 = r'\sigma_{\mathrm{ak}}=' + f'{sigma1}+' + str(Gs) + r'\times' + f'{Solid_D1}={sigma2}KPa'
            Solid_Pressure2 = r'p_{\mathrm{ak}}=' + str(sigma2) + r'\times' + f'{Ka}-2' + r'\times' + str(C) + r'\times\sqrt{' + str(Ka) + r'}=' + f'{pak2}kPa'
            sigma1_key = f"{{Separate_Active_Solid_Sigma1_{Solid_number}_}}"
            pressure1_key = f"{{Separate_Active_Solid_Pressure1_{Solid_number}_}}"
            sigma2_key = f"{{Separate_Active_Solid_Sigma2_{Solid_number}_}}"
            pressure2_key = f"{{Separate_Active_Solid_Pressure2_{Solid_number}_}}"
            # 公式存入
            formula_registry[sigma1_key] = Solid_Sigma1
            formula_registry[pressure1_key] = Solid_Pressure1
            formula_registry[sigma2_key] = Solid_Sigma2
            formula_registry[pressure2_key] = Solid_Pressure2
            # 计数器自增
            Solid_number += 1
        # 被动土压力
        condition_num = len(Passive_Earth_Pressure_Hlst) # 基坑内的工况数量
        print(condition_num)
        condition_passive_pressure = []
        condition_passive_water = []
        for j in range(condition_num):
            passive_pressure = []
            condition_index = j+1
            # 基坑内水压力公式
            Water_H2 = round(water_level_condition_lst[j]-CofferDam_Bottom_Level,acc)
            Water_P2 = round(10 * Water_H2,acc)
            Water_Pressure2 = f'P_1={Water_H2}' + r'\times10' + f'={Water_P2}kPa'
            water_pressure2_key = f'{{Separate_Passive_Water_Pressure_{condition_index}}}'
            formula_registry[water_pressure2_key] = Water_Pressure2
            condition_passive_water.append({
                'index':condition_index, 
                'Separate_Passive_Water_Condition_Name': '基坑底' if combo_dict else f'z={Passive_Earth_Pressure_Hlst[j][0]}m',
                'Separate_Passive_Water_Top_Level': water_level_condition_lst[j],
                'Separate_Passive_CofferDam_Bottom_Level': CofferDam_Bottom_Level,
                }
            )
            for i in range(len(Passive_Earth_Pressure_Hlst[j])-1):
                Solid_number = i+1
                Solid_Layer_Method = Passive_Earth_Pressure_Solid_Info[j][i][3]
                Solid_Layer_Name = Passive_Earth_Pressure_Solid_Info[j][i][4]
                Solid_H1 = Passive_Earth_Pressure_Hlst[j][i]
                Solid_H2 = Passive_Earth_Pressure_Hlst[j][i+1]
                if Solid_Layer_Name == '水':
                     continue
                Solid_D1 = round(Solid_H1 - Solid_H2, acc)
                Solid_if_Water = '水下' if Passive_Earth_Pressure_Solid_State[j][i] == '含水' else '水上'
                passive_pressure.append({
                    'index':Solid_number,
                    'Separate_Passive_Solid_number':Solid_number,
                    'Separate_Passive_Solid_Layer_Name':Solid_Layer_Name,
                    'Separate_Passive_Solid_if_Water':Solid_if_Water,
                    'Separate_Passive_Solid_Method_for_Pressures_var':Solid_Layer_Method,
                    'Separate_Passive_Solid_H1':Solid_H1,
                    'Separate_Passive_Solid_H2':Solid_H2,
                    'Separate_Passive_Solid_D1':Solid_D1    
                    })
                # 公式计算
                Gs = round(float(Passive_Earth_Pressure_Solid_Info[j][i][0]-10), acc) if Solid_if_Water == '水下' else round(float(Passive_Earth_Pressure_Solid_Info[j][i][0]), acc) # 重度
                C = round(float(Passive_Earth_Pressure_Solid_Info[j][i][1]), acc) # 黏聚力
                Fai = round(float(Passive_Earth_Pressure_Solid_Info[j][i][2]), acc) # 内摩擦角
                Kp = cal_Kp(Fai) # Ka
                sigma_soil_top = round(Passive_Earth_Pressure_Solid_Sigma_lst[j][i], acc)
                sigma_soil_bot = round(Passive_Earth_Pressure_Solid_Sigma_lst[j][i+1], acc)
                pak_soil_top = round(Passive_Earth_Pressure_Solid_Ppk_lst[j][i][0], acc)
                pak_soil_bot = round(Passive_Earth_Pressure_Solid_Ppk_lst[j][i][1], acc)
                # 公式填充
                sigma1 = round(sigma_soil_top, acc)
                sigma2 = round(sigma_soil_bot, acc)
                ppk1 = round(pak_soil_top, acc)
                ppk2 = round(pak_soil_bot, acc)
                Solid_Sigma1 = r'\sigma_{\mathrm{pk}}=' + f'{sigma1}KPa'
                Solid_Pressure1 = r'p_{\mathrm{pk}}=' + str(sigma1) + r'\times' + f'{Kp}+2' + r'\times' + str(C) + r'\times\sqrt{' + str(Kp) + r'}=' + f'{ppk1}kPa'
                Solid_Sigma2 = r'\sigma_{\mathrm{pk}}=' + f'{sigma1}+' + str(Gs) + r'\times' + f'{Solid_D1}={sigma2}KPa'
                Solid_Pressure2 = r'p_{\mathrm{pk}}=' + str(sigma2) + r'\times' + f'{Kp}+2' + r'\times' + str(C) + r'\times\sqrt{' + str(Kp) + r'}=' + f'{ppk2}kPa'
                sigma1_key = f"{{Separate_Passive_Solid_Sigma1_{condition_index}_{Solid_number}_}}"
                pressure1_key = f"{{Separate_Passive_Solid_Pressure1_{condition_index}_{Solid_number}_}}"
                sigma2_key = f"{{Separate_Passive_Solid_Sigma2_{condition_index}_{Solid_number}_}}"
                pressure2_key = f"{{Separate_Passive_Solid_Pressure2_{condition_index}_{Solid_number}_}}"
                # 公式存入
                formula_registry[sigma1_key] = Solid_Sigma1
                formula_registry[pressure1_key] = Solid_Pressure1
                formula_registry[sigma2_key] = Solid_Sigma2
                formula_registry[pressure2_key] = Solid_Pressure2
            condition_passive_pressure.append({
                # condition循环序号
                'index':condition_index, 
                # 单个工况下的被动土压力信息
                'passive_pressure':passive_pressure,
                # 被动土压力文字
                'Separate_Passive_Solid_Condition_Name':'基坑底' if combo_dict else f'z={Passive_Earth_Pressure_Hlst[j][0]}m',
                }
            )
        context = {
            'separate':True,
            'combine':False,
            # 主动水压力文字
            'Separate_Active_Water_Top_Level': Water_Top_Level,
            'Separate_Active_CofferDam_Bottom_Level': CofferDam_Bottom_Level,
            # 被动水压力文字
            'separate_passive_water_condition':condition_passive_water,
            # 主动土压力文字
            'Separate_Active_Ea_p0': Ea_p0,
            # 主动土压力
            'active_pressure':active_pressure,
            # 被动土压力
            'separate_passive_pressure_condition':condition_passive_pressure,
            
        }
    elif Method_for_Pressures_var == '水土合算':
        # 主动土压力
        active_pressure = []
        formula_registry = {}
        # 初始化一个显示用的序号计数器
        Solid_number = 1
        for i in range(len(Active_Earth_Pressure_Hlst)-1):
            Solid_Layer_Method = Active_Earth_Pressure_Solid_Info[i][3]
            Solid_Layer_Name = Active_Earth_Pressure_Solid_Info[i][4]
            Solid_H1 = Active_Earth_Pressure_Hlst[i]
            Solid_H2 = Active_Earth_Pressure_Hlst[i+1]
            Solid_D1 = round(Solid_H1 - Solid_H2, acc)
            Solid_if_Water = '水下' if Active_Earth_Pressure_Solid_State[i] == '含水' else '水上'
            not_all_water = True # 不是全都是水-默认True
            method_separate = False # 水土合算下的水土分算-默认False
            if Solid_Layer_Method == '/': # 当全部是水
                not_all_water = False # 土压力计算书不显示土压力计算
                method_separate = True # 显示水压力计算
            elif Solid_Layer_Method == '水土分算' and Solid_if_Water == '水下': # 当该土层为水土分算并且位于水下
                method_separate = True
            active_pressure.append({
                'index':Solid_number,
                'Combine_Active_Solid_number':Solid_number,
                'Combine_Active_Solid_Layer_Name':Solid_Layer_Name,
                'Combine_Active_Solid_if_Water':Solid_if_Water,
                'Combine_Active_Solid_Method_for_Pressures_var':Solid_Layer_Method,
                'Combine_Active_Solid_H1':Solid_H1,
                'Combine_Active_Solid_H2':Solid_H2,
                'Combine_Active_Solid_D1':Solid_D1,
                'not_all_water':not_all_water,
                'method_separate':method_separate,
            })
            # 公式计算
            Gs = round(float(Active_Earth_Pressure_Solid_Info[i][0]), acc) # 重度
            C = round(float(Active_Earth_Pressure_Solid_Info[i][1]), acc) # 黏聚力
            Fai = round(float(Active_Earth_Pressure_Solid_Info[i][2]), acc) # 内摩擦角
            Ka = cal_Ka(Fai) # Ka
            sigma_soil_top = round(Active_Earth_Pressure_Solid_Sigma_lst[i], acc)
            sigma_soil_bot = round(Active_Earth_Pressure_Solid_Sigma_lst[i+1], acc)
            pak_soil_top = round(Active_Earth_Pressure_Solid_Pak_lst[i][0], acc)
            pak_soil_bot = round(Active_Earth_Pressure_Solid_Pak_lst[i][1], acc)
            sigma_water_top = round(Active_Earth_Pressure_Water_Sigma_lst[i], acc)
            sigma_water_bot = round(Active_Earth_Pressure_Water_Sigma_lst[i+1], acc)
            pak_water_top = round(Active_Earth_Pressure_Water_Pak_lst[i][0], acc)
            pak_water_bot = round(Active_Earth_Pressure_Water_Pak_lst[i][1], acc)
            if not_all_water:
                # 公式填充
                sigma1 = round(sigma_soil_top, acc) if method_separate else round(sigma_soil_top + sigma_water_top, acc)
                sigma2 = round(sigma_soil_bot, acc) if method_separate else round(sigma_soil_bot + sigma_water_bot, acc)
                pak1 = round(pak_soil_top, acc) if method_separate else round(pak_soil_top + pak_water_top, acc)
                pak2 = round(pak_soil_bot, acc) if method_separate else round(pak_soil_bot + pak_water_bot, acc)
                sigma_gs = round(Gs-10, acc) if method_separate else Gs
                Solid_Sigma1 = r'\sigma_{\mathrm{ak}}=' + f'{sigma1}KPa'
                Solid_Pressure1 = r'p_{\mathrm{ak}}=' + str(sigma1) + r'\times' + f'{Ka}-2' + r'\times' + str(C) + r'\times\sqrt{' + str(Ka) + r'}=' + f'{pak1}kPa'
                Solid_Sigma2 = r'\sigma_{\mathrm{ak}}=' + f'{sigma1}+' + str(sigma_gs) + r'\times' + f'{Solid_D1}={sigma2}KPa'
                Solid_Pressure2 = r'p_{\mathrm{ak}}=' + str(sigma2) + r'\times' + f'{Ka}-2' + r'\times' + str(C) + r'\times\sqrt{' + str(Ka) + r'}=' + f'{pak2}kPa'
                sigma1_key = f"{{Combine_Active_Solid_Sigma1_{Solid_number}_}}"
                pressure1_key = f"{{Combine_Active_Solid_Pressure1_{Solid_number}_}}"
                sigma2_key = f"{{Combine_Active_Solid_Sigma2_{Solid_number}_}}"
                pressure2_key = f"{{Combine_Active_Solid_Pressure2_{Solid_number}_}}"
                # 公式存入
                formula_registry[sigma1_key] = Solid_Sigma1
                formula_registry[pressure1_key] = Solid_Pressure1
                formula_registry[sigma2_key] = Solid_Sigma2
                formula_registry[pressure2_key] = Solid_Pressure2
            if method_separate:
                # 公式填充
                sigma1 = round(sigma_water_top, acc)
                sigma2 = round(sigma_water_bot, acc)
                Solid_Sigma1 = r'\sigma_{\mathrm{ak}}=' + f'{sigma1}KPa'
                Solid_Sigma2 = r'\sigma_{\mathrm{ak}}=' + f'{sigma1}+' + str(10) + r'\times' + f'{Solid_D1}={sigma2}KPa'
                sigma1_key = f"{{Combine_Active_Water_Sigma1_{Solid_number}_}}"
                sigma2_key = f"{{Combine_Active_Water_Sigma2_{Solid_number}_}}"
                # 公式存入
                formula_registry[sigma1_key] = Solid_Sigma1
                formula_registry[sigma2_key] = Solid_Sigma2
            # 计数器自增
            Solid_number += 1
        # 被动土压力
        condition_num = len(Passive_Earth_Pressure_Hlst) # 基坑内的工况数量
        condition_passive_pressure = []
        for j in range(condition_num):
            passive_pressure = []
            condition_index = j+1
            for i in range(len(Passive_Earth_Pressure_Hlst[j])-1):
                Solid_number = i+1
                Solid_Layer_Method = Passive_Earth_Pressure_Solid_Info[j][i][3]
                Solid_Layer_Name = Passive_Earth_Pressure_Solid_Info[j][i][4]
                Solid_H1 = Passive_Earth_Pressure_Hlst[j][i]
                Solid_H2 = Passive_Earth_Pressure_Hlst[j][i+1]
                Solid_D1 = round(Solid_H1 - Solid_H2, acc)
                Solid_if_Water = '水下' if Passive_Earth_Pressure_Solid_State[j][i] == '含水' else '水上'
                not_all_water = True # 不是全都是水-默认True
                method_separate = False # 水土合算下的水土分算-默认False
                if Solid_Layer_Method == '/': # 当全部是水
                    not_all_water = False # 土压力计算书不显示土压力计算
                    method_separate = True # 显示水压力计算
                elif Solid_Layer_Method == '水土分算' and Solid_if_Water == '水下': # 当该土层为水土分算并且位于水下
                    method_separate = True
                passive_pressure.append({
                    'index':Solid_number,
                    'Combine_Passive_Solid_number':Solid_number,
                    'Combine_Passive_Solid_Layer_Name':Solid_Layer_Name,
                    'Combine_Passive_Solid_if_Water':Solid_if_Water,
                    'Combine_Passive_Solid_Method_for_Pressures_var':Solid_Layer_Method,
                    'Combine_Passive_Solid_H1':Solid_H1,
                    'Combine_Passive_Solid_H2':Solid_H2,
                    'Combine_Passive_Solid_D1':Solid_D1,
                    'not_all_water':not_all_water,
                    'method_separate':method_separate,
                    })
                # 公式计算
                Gs = round(float(Passive_Earth_Pressure_Solid_Info[j][i][0]), acc) # 重度
                C = round(float(Passive_Earth_Pressure_Solid_Info[j][i][1]), acc) # 黏聚力
                Fai = round(float(Passive_Earth_Pressure_Solid_Info[j][i][2]), acc) # 内摩擦角
                Kp = cal_Kp(Fai) # Ka
                sigma_soil_top = round(Passive_Earth_Pressure_Solid_Sigma_lst[j][i], acc)
                sigma_soil_bot = round(Passive_Earth_Pressure_Solid_Sigma_lst[j][i+1], acc)
                pak_soil_top = round(Passive_Earth_Pressure_Solid_Ppk_lst[j][i][0], acc)
                pak_soil_bot = round(Passive_Earth_Pressure_Solid_Ppk_lst[j][i][1], acc)
                sigma_water_top = round(Passive_Earth_Pressure_Water_Sigma_lst[j][i], acc)
                sigma_water_bot = round(Passive_Earth_Pressure_Water_Sigma_lst[j][i+1], acc)
                pak_water_top = round(Passive_Earth_Pressure_Water_Ppk_lst[j][i][0], acc)
                pak_water_bot = round(Passive_Earth_Pressure_Water_Ppk_lst[j][i][1], acc)
                if not_all_water:
                    # 公式填充
                    sigma1 = round(sigma_soil_top, acc) if method_separate else round(sigma_soil_top + sigma_water_top, acc)
                    sigma2 = round(sigma_soil_bot, acc) if method_separate else round(sigma_soil_bot + sigma_water_bot, acc)
                    ppk1 = round(pak_soil_top, acc) if method_separate else round(pak_soil_top + pak_water_top, acc)
                    ppk2 = round(pak_soil_bot, acc) if method_separate else round(pak_soil_bot + pak_water_bot, acc)
                    sigma_gs = round(Gs-10, acc) if method_separate else Gs
                    Solid_Sigma1 = r'\sigma_{\mathrm{pk}}=' + f'{sigma1}KPa'
                    Solid_Pressure1 = r'p_{\mathrm{pk}}=' + str(sigma1) + r'\times' + f'{Kp}+2' + r'\times' + str(C) + r'\times\sqrt{' + str(Kp) + r'}=' + f'{ppk1}kPa'
                    Solid_Sigma2 = r'\sigma_{\mathrm{pk}}=' + f'{sigma1}+' + str(sigma_gs) + r'\times' + f'{Solid_D1}={sigma2}KPa'
                    Solid_Pressure2 = r'p_{\mathrm{pk}}=' + str(sigma2) + r'\times' + f'{Kp}+2' + r'\times' + str(C) + r'\times\sqrt{' + str(Kp) + r'}=' + f'{ppk2}kPa'
                    sigma1_key = f"{{Combine_Passive_Solid_Sigma1_{condition_index}_{Solid_number}_}}"
                    pressure1_key = f"{{Combine_Passive_Solid_Pressure1_{condition_index}_{Solid_number}_}}"
                    sigma2_key = f"{{Combine_Passive_Solid_Sigma2_{condition_index}_{Solid_number}_}}"
                    pressure2_key = f"{{Combine_Passive_Solid_Pressure2_{condition_index}_{Solid_number}_}}"
                    # 公式存入
                    formula_registry[sigma1_key] = Solid_Sigma1
                    formula_registry[pressure1_key] = Solid_Pressure1
                    formula_registry[sigma2_key] = Solid_Sigma2
                    formula_registry[pressure2_key] = Solid_Pressure2
                if method_separate:
                    # 公式填充
                    sigma1 = round(sigma_water_top, acc)
                    sigma2 = round(sigma_water_bot, acc)
                    Solid_Sigma1 = r'\sigma_{\mathrm{ak}}=' + f'{sigma1}KPa'
                    Solid_Sigma2 = r'\sigma_{\mathrm{ak}}=' + f'{sigma1}+' + str(10) + r'\times' + f'{Solid_D1}={sigma2}KPa'
                    sigma1_key = f"{{Combine_Passive_Water_Sigma1_{condition_index}_{Solid_number}_}}"
                    sigma2_key = f"{{Combine_Passive_Water_Sigma2_{condition_index}_{Solid_number}_}}"
                    # 公式存入
                    formula_registry[sigma1_key] = Solid_Sigma1
                    formula_registry[sigma2_key] = Solid_Sigma2
            Passive_Water_Top_Level_contxt = str(water_level_condition_lst[j]) if water_level_condition_lst[j] >= CofferDam_Bottom_Level else f'低于{CofferDam_Bottom_Level}m'
            condition_passive_pressure.append({
                # condition循环序号
                'index':condition_index, 
                # 单个工况下的被动土压力信息
                'passive_pressure':passive_pressure,
                # 被动土压力文字
                'Combine_Passive_Solid_Condition_Name': '基坑底' if combo_dict else f'z={Passive_Earth_Pressure_Hlst[j][0]}m',
                'Combine_Passive_Water_Top_Level': Passive_Water_Top_Level_contxt,
                'Combine_Passive_CofferDam_Bottom_Level': CofferDam_Bottom_Level,
                }
            )
        Active_Water_Top_Level_contxt = str(Water_Top_Level) if Water_Top_Level >= CofferDam_Bottom_Level else f'低于{CofferDam_Bottom_Level}m'
        context = {
            'separate':False,
            'combine':True,
            # 主动土压力文字
            'Combine_Active_Water_Top_Level': Active_Water_Top_Level_contxt,
            'Combine_Active_CofferDam_Bottom_Level': CofferDam_Bottom_Level,
            'Combine_Active_Ea_p0': Ea_p0,
            # 主动土压力
            'active_pressure':active_pressure,
            # 被动土压力
            'combine_passive_pressure_condition':condition_passive_pressure,
        }
    tpl = DocxTemplate(template_path)
    tpl.render(context)
    tpl.save(save_path)
    # 公式替换  
    final_path = save_path.replace('.docx', '_final.docx')
    finaLdoc =Document(save_path)
    for placeholder, latex in formula_registry.items():
        replace_placeholder_with_formula(finaLdoc, placeholder, latex)
    finaLdoc.save(final_path)
    print(f"土压力部分计算书已保存至{final_path}")


# m法
def m_method(doc_path, save_path, Condition_Passive_Earth_Pressure_Hlst, Condition_Passive_Earth_Pressure_Solid_Info, a, b, acc = 3):
    # 路径处理
    template_path = os.path.join(doc_path, 'Steel_Sheet_Pile_CofferDam_cal_template_m_method.docx')
    save_path = os.path.join(save_path, 'm_method.docx')
    Condition = [num_to_chinese_num(i+1) for i in range(len(Condition_Passive_Earth_Pressure_Hlst))]
    Condition_m_method_Hlst = []
    conditions = []
    for i in range(len(Condition)):
        index = i+1
        cn_index = num_to_chinese_num(index) 
        data_rows = []
        Solid_Hlst_origin = Condition_Passive_Earth_Pressure_Hlst[i]
        Solid_Layer_Infolst = Condition_Passive_Earth_Pressure_Solid_Info[i]
        # 判断是否有纯水层, 有则去掉该层顶标高
        Solid_Hlst = [Solid_Hlst_origin[i] for i, lst in enumerate(Solid_Layer_Infolst) if lst[4] != '水'] + [Solid_Hlst_origin[-1]]
        # 对土层划分土条
        m_method_Hlst = Divide_Solid_within_LayerThickness(Solid_Hlst[0], Solid_Hlst[-1], a, a/2)
        Condition_m_method_Hlst.append(m_method_Hlst)
        for j in range(len(m_method_Hlst)-1):
            Level1 = m_method_Hlst[j]
            Level2 = m_method_Hlst[j+1]
            # 判断土条的分层情况
            Solid_Slince_lst = intersect_segment_with_intervals(Solid_Hlst, [Level1, Level2])
            # 判断土条的分层情况下对应的土层信息
            Solid_Slince_Infolst = [[Solid_Slince_lst[m][-1]] + Solid_Layer_Infolst[Solid_Slince_lst[m][0]-1] for m in range(len(Solid_Slince_lst))]
            # 计算m所需要的土层厚度、黏聚力和内摩擦角
            mlst = [[float(Solid_Slince_Infolst[m][0]), cal_m(float(Solid_Slince_Infolst[m][3]), float(Solid_Slince_Infolst[m][2]))] for m in range(len(Solid_Slince_Infolst))]
            # 计算加权平均m
            m_sum = sum(w * v for w, v in mlst)
            h_sum = sum(w for w, v in mlst)
            m = m_sum / h_sum
            # 计算Ks
            m_Level = round((Level1 + Level2)/2, acc) # 计算土层标高
            Z = round(abs(m_Level - Solid_Hlst[0]), acc) # 计算深度
            Ks = cal_Ks(m, Z, h_sum, b/1000)
            data_rows.append({
                'no': j + 1,
                'a': round(h_sum, acc),          
                'b': round(b/1000, acc),         
                'M': round(m, 1),             
                'z': round(Z, acc),             
                'Ks': round(Ks, 1),           
                'level': round(m_Level, acc)
            })
        conditions.append({
            'index': index,
            'cn_index': cn_index,
            'tuceng': data_rows
        })
    context = {
        'conditions': conditions
    }
    m_method_doc = DocxTemplate(template_path)
    m_method_doc.render(context)
    m_method_doc.save(save_path)
    return Condition_m_method_Hlst


# 围堰概况
def CofferDam_Info(
        SECtype, Cap_length, Cap_width, Cap_height, Cap_Bottom_Level, CofferDam_a, CofferDam_b, CofferDam_L, CofferDam_Top_Level, CofferDam_Bottom_Level, Water_Top_Level,Concrete_Bottom_Level, CofferDam_type, conc_type, conc_grade, conc_t, Waler_SECT_lst, Waler_MATL_lst, DC_MATL_lst, XC_MATL_lst, DC_SECT_lst, XC_SECT_lst, DC_group_namelst,XC_group_namelst, Ea_p0, SECmaterial=''
    ):
    '''
    根据Waler_SECT_lst：['I40a', 'HW400X400']
    Strut_SECT_lst：['I40a', '377X6钢管桩']
    XC_group_namelst：['第1层斜撑', '第2层斜撑']
    DC_group_namelst：[]
    总结support_system
    (目前默认同一层支撑都是同种截面)
    '''
    support_system = []
    current_no = 1
    for i in range(len(Waler_SECT_lst)):
        layer_num = i + 1
        waler_sect = Waler_SECT_lst[i]
        Waler_matl = Waler_MATL_lst[i]
        dc_matl = DC_MATL_lst[i]
        xc_matl = XC_MATL_lst[i]
        dc_sect = DC_SECT_lst[i] 
        xc_sect = XC_SECT_lst[i]
        # 处理围檩
        if waler_sect:
            support_system.append({
                'no': current_no,
                'name': f"第{layer_num}层围檩",
                'material': Waler_matl,
                'sect': waler_sect,
                'note': f"",
            })
            current_no += 1
        # 处理对撑
        if dc_sect != '/' and dc_matl != '/':
            for name in [n for n in DC_group_namelst if f"第{layer_num}层" in n]:
                support_system.append({
                    'no': current_no,
                    'name': name,
                    'material': dc_matl,
                    'sect': dc_sect,
                    'note': "",
                })
                current_no += 1
        # 处理斜撑
        if xc_sect != '/' and xc_matl != '/':
            for name in [n for n in XC_group_namelst if f"第{layer_num}层" in n]:
                support_system.append({
                    'no': current_no,
                    'name': name,
                    'material': xc_matl,
                    'sect': xc_sect,
                    'note': "",
                })
                current_no += 1
    next_no = current_no
    # 读取自重系数
    Self_Weight = MidasAPI("GET","/db/BODF")
    if Self_Weight == {'message': ''} or 'error' in Self_Weight:
        Self_Weight_factor = "No Self_Weight Defined"
    else:
        Self_Weight_factor = Self_Weight['BODF']['1']['FV'][2]
    Water_Level = '/' if float(Water_Top_Level) < float(CofferDam_Bottom_Level) else round(float(Water_Top_Level),3)
    info = {
        'chengtaichang': round(Cap_length,1),
        'chengtaikuan': round(Cap_width,1),
        'chengtaigao': round(Cap_height,1),
        'chengtaidingbiaogao': round(Cap_Bottom_Level + Cap_height,3),
        'weiyanchang': round(CofferDam_a,1),
        'weiyankuan': round(CofferDam_b,1),
        'weiyangao': round(CofferDam_L,1),
        'weiyanbiaogao': round(CofferDam_Top_Level,3),
        'conc_type': conc_type,
        'conc_grade': conc_grade,
        'conc_t': conc_t,
        'keng_wai_shui_wei_biao_gao': Water_Level,
        'keng_nei_shui_wei_biao_gao': round(Concrete_Bottom_Level,3),
        'SECtype': SECtype,
        'SECmaterial': (SECmaterial + ('bz' if SECtype == "钢板桩" else 'B')) if SECmaterial else ('Q295bz' if SECtype == "钢板桩" else 'Q235B'),
        'SECspecification': CofferDam_type,
        'pile_length': round(CofferDam_L,0),
        'support_system': support_system,
        'next_no': str(next_no),
        'Model_Gz':Self_Weight_factor,
        'fu_jia_he_zai':round(Ea_p0,2)
    }
    for k, v in info.items():
        print(k, v)
    print('')
    return info


# 地质参数表
def Di_Zhi_Can_Shu(Excel_Data, Solid_Level):
    soild_layers =[]
    first_row_raw = Excel_Data[1]
    accumulated_thickness = first_row_raw[1]
    # 第一行（数据行索引1）
    soild_layers.append({
        'name': first_row_raw[0],              # 土层名称
        'level': float(Solid_Level),           # 土层标高
        't': float(first_row_raw[1]),          # 层厚
        'gamma': float(first_row_raw[2]),      # 重度
        'c': float(first_row_raw[3]),          # 粘聚力
        'phi': float(first_row_raw[4]),        # 内摩擦角
        'ka': cal_Ka(float(first_row_raw[4])),        # Ka
        'kp': cal_Kp(float(first_row_raw[4])),        # Kp
        'if_combine': first_row_raw[5]         # 计算方法
    })
    # 从第二行开始（数据行索引2），累计减去前面的层厚
    for i in range(2, len(Excel_Data)):
        row = Excel_Data[i]
        current_level = Solid_Level - accumulated_thickness
        soild_layers.append({
            'name': row[0],
            'level': round(current_level, 3),
            't': float(row[1]),
            'gamma': float(row[2]),
            'c': float(row[3]),
            'phi': float(row[4]),
            'ka': cal_Ka(float(row[4])),
            'kp': cal_Kp(float(row[4])),
            'if_combine': row[5]
        })
        # 累加层厚供下一行使用
        accumulated_thickness += row[1]
    for x in soild_layers:
        print(x)
    print('')
    return soild_layers


def Ji_Shu_Can_Shu(Waler_section_dict, Strut_section_dict, Waler_SECT_lst, DC_SECT_lst, XC_SECT_lst, WL_group_namelst, XC_group_namelst, DC_group_namelst, gbz_name, ggz_name, ggz_A, ggz_I, ggz_W):
    # 钢板桩
    gbz_dict = {
    '拉森Ⅳ':{'type':"PU400*170", 'b':400, 'h':170, 't':15.5, 'm':76.1, 'A':242.5, 'W':2270, 'I':38600},
    '拉森Ⅵ':{'type':"PU600*210", 'b':600, 'h':210, 't':18  , 'm':106 , 'A':225.5, 'W':2700, 'I':56700},
    }
    lasen4_namelst = ['拉森四', '拉森Ⅳ', '拉森4', 'SP-Ⅳ' , 'PU400*170']
    lasen6_namelst = ['拉森六', '拉森Ⅵ', '拉森6', 'SP-ⅣW', 'PU600*210']
    gbz = None
    if gbz_name in lasen4_namelst:
        gbz = gbz_dict['拉森Ⅳ']
    elif gbz_name in lasen6_namelst:
        gbz = gbz_dict['拉森Ⅵ']
    # 钢管桩
    ggz = {'type':ggz_name, 'A':ggz_A, 'I':ggz_I, 'W':ggz_W}
    # 支撑体系
    support_system = []
    wl_name_pattern = r'^第\d+层围檩$'
    xc_name_pattern = r'^第\d+层斜撑$'
    dc_name_pattern = r'^第\d+层对撑$'
    for i in range(len(Waler_SECT_lst)):
        layer_num = i + 1
        waler_sect = Waler_SECT_lst[i]
        dc_sect = DC_SECT_lst[i]
        xc_sect = XC_SECT_lst[i]
        print(i)
        print(waler_sect)
        print(xc_sect)
        print(dc_sect)
        # 围檩技术参数 
        if waler_sect and str(waler_sect).strip().lower() not in ['/', 'none', '']:
            wl_name = [n for n in WL_group_namelst if re.match(wl_name_pattern, n) and f"第{layer_num}层" in n][0]
            if wl_name:
                params = Waler_section_dict[waler_sect]
                support_system.append({
                    'name': wl_name,
                    'sect': waler_sect,
                    'A' : params['A'] ,
                    'WX': params['Wx'],
                    'WY': params['Wy'],
                    'ix': params['ix'],
                    'iy': params['iy'],
                })
        # 斜撑截面技术参数
        if xc_sect and str(xc_sect).strip().lower() not in ['/', 'none', '']:
            xc_name = [n for n in XC_group_namelst if re.match(xc_name_pattern, n) and f"第{layer_num}层" in n][0]
            if xc_name:
                params = Strut_section_dict[xc_sect]
                support_system.append({
                    'name': xc_name, 
                    'sect': xc_sect,
                    'A' : params['A'] ,
                    'WX': params['Wx'],
                    'WY': params['Wy'],
                    'ix': params['ix'],
                    'iy': params['iy'],
                })
        # 对撑截面技术参数
        if dc_sect and str(dc_sect).strip().lower() not in ['/', 'none', '']:
            dc_name = [n for n in DC_group_namelst if re.match(dc_name_pattern, n) and f"第{layer_num}层" in n][0]
            if dc_name:
                params = Strut_section_dict[dc_sect]
                support_system.append({
                    'name': dc_name,
                    'sect': dc_sect,
                    'A' : params['A'] ,
                    'WX': params['Wx'],
                    'WY': params['Wy'],
                    'ix': params['ix'],
                    'iy': params['iy'],
                })
    jishucanshu = {
        'gbz': gbz,
        'ggz': ggz,
        'support_system': support_system,
    }

    for k, v in jishucanshu.items():
        print(k, v)
    return jishucanshu


def case_info(has_stage, Drawdown_height, Waler_SECT_lst, DC_SECT_lst, XC_SECT_lst, conc_type, conc_grade, conc_t):
    cases = []
    load_cases = []
    num_layers = len(Waler_SECT_lst)
    is_cantilever = (num_layers == 0)
    def check_has_strut(idx):
        # 检查指定索引层是否有有效支撑
        has_zc = idx < len(DC_SECT_lst) and str(DC_SECT_lst[idx]).strip() != '/'
        has_xc = idx < len(XC_SECT_lst) and str(XC_SECT_lst[idx]).strip() != '/'
        return has_zc or has_xc
    # 施工阶段建模
    if has_stage :
        Stage_res = MidasAPI("GET", "/db/STAG")
        stag_dict = Stage_res.get('STAG', {})
        keys = sorted(stag_dict.keys(),key=lambda x: int(x))
        stage_cases = [stag_dict[key]['NAME'] for key in keys]
        has_plug = any("封底" in name for name in stage_cases) # 是否有封底
        after_plug = False # 是否已封底
        for j, case_name in enumerate(stage_cases):
            case_num_cn = num_to_chinese_num(j + 1)
            # 处理工况名称
            num_match = re.findall(r'\d+', case_name)
            target_layer_idx = int(num_match[0]) - 1 if num_match else 0
            if "封底" in case_name:
                desc = f"工况{case_num_cn}：浇筑{conc_t}m厚{conc_grade}{conc_type}；"
                after_plug = True
            elif "基坑" in case_name:
                if has_plug:
                    desc = f"工况{case_num_cn}：围堰内降水至基坑底标高;"
                else:
                    desc = f"工况{case_num_cn}：围堰内开挖至基坑底标高，并浇筑{conc_t}m厚{conc_grade}{conc_type}；"
            elif "安装" in case_name and not is_cantilever:
                layer_num_cn = num_to_chinese_num(int(num_match[0]))
                strut_desc = "及内支撑" if check_has_strut(target_layer_idx) else ""
                desc = f"工况{case_num_cn}：安装第{layer_num_cn}层围檩{strut_desc}；"
            elif "开挖" in case_name and "层" in case_name:
                action_text = "降水" if after_plug else "围堰内继续开挖"
                layer_num_cn = num_to_chinese_num(int(num_match[0]))
                if j == 0:
                    # 起始工况描述
                    desc = f"工况{case_num_cn}：钢板桩围堰插打完毕后，围堰内开挖至第{layer_num_cn}层围檩设计标高下{Drawdown_height}m；"
                else:
                    # 非起始工况
                    prev_case = stage_cases[j-1]
                    if "安装" in prev_case:
                        prev_num = re.findall(r'\d+', prev_case)
                        prev_layer_idx = int(prev_num[0]) - 1 if prev_num else 0
                        prev_strut = "及内支撑" if check_has_strut(prev_layer_idx) else ""
                        condition_prefix = (f"待第{num_to_chinese_num(prev_layer_idx + 1)}层围檩{prev_strut}安装完毕后，")
                    elif "封底" in prev_case:
                        condition_prefix = "待封底混凝土达到设计强度后，"
                    else:
                        condition_prefix = ""
                    target_label = f"第{num_to_chinese_num(target_layer_idx + 1)}层围檩设计标高下{Drawdown_height}m"
                    desc = f"工况{case_num_cn}：{condition_prefix}{action_text}至{target_label}；"
            else:
                # 处理未定义情况
                desc = f"工况{case_num_cn}：{case_name}；"
            cases.append(desc)
            # load_cases.append(desc.split('：')[1].split('；')[0])
            load_cases = stage_cases
    # 整体建模
    else:
        if is_cantilever:
            cases.append(f"工况一：钢板桩围堰插打完毕后，围堰内开挖至基坑底标高。")
            load_cases.append("工况一")
        else:
            cases.append(f"工况一：钢板桩围堰插打完毕后，围堰内挖土至第一道围檩设计标高下{Drawdown_height}m；")
            load_cases.append("工况一")
            for i in range(num_layers - 1):
                current_layer_idx = i
                next_layer_idx = i + 1
                # 判断当前层的情况
                strut_desc = "及内支撑" if check_has_strut(current_layer_idx) else ""
                case_num_cn = num_to_chinese_num(i + 2) # 从工况二开始编号
                desc = (f"工况{case_num_cn}：待第{num_to_chinese_num(current_layer_idx + 1)}层围檩{strut_desc}安装完毕后，"
                        f"围堰内继续开挖至第{num_to_chinese_num(next_layer_idx + 1)}层围檩设计标高下{Drawdown_height}m；")
                cases.append(desc)
                load_cases.append(f"工况{case_num_cn}")
            # 安装最后一层，开挖至基底并浇筑封底/垫层
            last_layer_idx = num_layers - 1
            last_strut_desc = "及内支撑" if check_has_strut(last_layer_idx) else ""
            last_case_num = num_to_chinese_num(len(cases) + 1)
            final_desc = (f"工况{last_case_num}：待第{num_to_chinese_num(last_layer_idx + 1)}层围檩{last_strut_desc}安装完毕后，"
                        f"围堰内继续开挖至基坑底，并浇筑{conc_t}m厚{conc_grade}{conc_type}。")
            cases.append(final_desc)
            load_cases.append(f"工况{last_case_num}")    
    return cases, load_cases


def comp_info():
    res_params = {
    'sw_sc': 1.0, 'solid_sc': 1.0, 'water_sc': 1.0,  # 标准组合默认值
    'sw_fc': 1.25, 'solid_fc': 1.25, 'water_fc': 1.25 # 基本组合默认值
    }
    COMP_res = MidasAPI("GET", "/db/LCOM-GEN")
    print('COMP_res')
    for k, v in COMP_res.items():
        print(k)
        for k1, v1 in v.items():
            print(k1, v1)
    for item in COMP_res.values():
        name = item.get('NAME', '')
        # 标准组合sc，基本组合fc
        suffix = '_sc' if '标准' in name else '_fc' if '基本' in name else None
        if suffix:
            v_comb = item.get('vCOMB', [])
            has_water_load = any('水压力' in v['LCNAME'] for v in v_comb)
            for v in v_comb:
                lc = v['LCNAME']
                f = v['FACTOR']
                # 处理自重
                if '自重' in lc:
                    res_params[f'sw{suffix}'] = f
                # 处理水土压力
                if '主动土压力' in lc:
                    if has_water_load:
                        # 【分算情况】：有独立的水工况，所以这里只赋值给土
                        res_params[f'solid{suffix}'] = f
                    else:
                        # 【合算情况】：没有独立的水工况，土压力系数同时赋给土和水
                        res_params[f'solid{suffix}'] = f
                        res_params[f'water{suffix}'] = f
                # 处理独立的水压力（仅在分算情况下会被触发）
                elif '水压力' in lc:
                    res_params[f'water{suffix}'] = f
    comp_info = {
        'sw_sc': res_params['sw_sc'],
        'solid_sc': res_params['solid_sc'],
        'water_sc': res_params['water_sc'],
        'sw_fc': res_params['sw_fc'],
        'solid_fc': res_params['solid_fc'],
        'water_fc': res_params['water_fc'],
    }
    return comp_info


def Pile_cal(doc,pic_path,load_cases, pile_leixing, allowable_f_dict, gamma_0, SECtype, CofferDam_L):
    # 容许应力
    pile_allowable_f = allowable_f_dict['pile']['f']
    pile_allowable_fv = allowable_f_dict['pile']['fv']
    # 容许位移
    pile_allowable_d = round(float(CofferDam_L)*2.5, 1)
    pile_jisuan = []
    # 获取截图
    API_DIST_FORCE_Unit("N", 'MM') # 应力为N和MM
    # 确保切换后处理状态，先截图
    Picture_Beamstress(pic_path, "gbz_cb_pic.jpg", [SECtype], "基本组合", "Combined", 45, 30)
    # 初始化追踪变量
    max_cb_all_cases = -1.0
    max_d_all_cases = -1.0
    max_cb_case_name = ""
    max_d_case_name = ""
    for i, case in enumerate(load_cases):
        case_name = case
        pile_stress_value_lst = Value_BeamStress("pilestress", SECtype, "基本组合(CB)")
        pile_deformed_value_lst = Value_Deformed("pile_displacement", SECtype, "标准组合(CB)")
        # 读取最大值
        pile_max_cb_value = max(pile_stress_value_lst['pilestress']['DATA'], key=lambda x: abs(float(x[3])))# 最大组合应力所在的结果表
        pile_max_cb = float(pile_max_cb_value[3])*gamma_0
        pile_max_shear_value = max(pile_stress_value_lst['pilestress']['DATA'], key=lambda x: abs(float(x[2])))# 最大剪应力所在的结果表
        pile_max_shear = float(pile_max_shear_value[2])*gamma_0
        pile_max_Dx_value = max(pile_deformed_value_lst['pile_displacement']['DATA'], key=lambda x: abs(float(x[2])))# 最大变形所在的结果表
        pile_max_Dy_value = max(pile_deformed_value_lst['pile_displacement']['DATA'], key=lambda x: abs(float(x[3])))# 最大变形所在的结果表
        pile_max_Dx = float(pile_max_Dx_value[2])*gamma_0
        pile_max_Dy = float(pile_max_Dy_value[3])*gamma_0
        dx_abs = abs(pile_max_Dx)
        dy_abs = abs(pile_max_Dy)
        if dx_abs >= dy_abs:
            pile_max_d = dx_abs
            pile_d_direction = "Dx"
        else:
            pile_max_d = dy_abs
            pile_d_direction= "Dy"
        if abs(pile_max_cb) <= float(pile_allowable_f) and pile_max_d <= float(pile_allowable_d):
            if_satisfied = '满足'
        else:
            if_satisfied = '不满足'
        # 加入列表
        pile_jisuan.append({
            'case': case_name,                          # 工况
            'type': pile_leixing,                       # 桩类型/规格
            'max_cb': round(pile_max_cb, 2),            # 最大应力
            'limit_cb': pile_allowable_f,               # 容许应力
            'max_d': round(pile_max_d, 2),              # 最大位移
            'limit_d': pile_allowable_d,                # 容许位移
            'check': if_satisfied                       # 验算结果
        })
        # 最不利情况更新
        current_cb_abs = abs(pile_max_cb)
        if current_cb_abs > max_cb_all_cases:
            max_cb_all_cases = current_cb_abs
            max_cb_case_name = case_name
        current_d_abs = pile_max_d
        if current_d_abs > max_d_all_cases:
            max_d_all_cases = current_d_abs
            max_d_case_name = case_name
    # 根据计算结果获取变形截图
    API_DIST_FORCE_Unit("N", 'MM') # 变形为MM
    Picture_Deformed(pic_path, "gbz_d_pic.jpg", [SECtype], "标准组合", pile_d_direction, 45, 30)
    # 处理图片
    cb_path = os.path.join(pic_path, "gbz_cb_pic.jpg")
    # shear_path = os.path.join(pic_path, "gbz_shear_pic.jpg")
    d_path = os.path.join(pic_path, "gbz_d_pic.jpg")
    # 读取图片
    cb_img = InlineImage(doc, cb_path, width=Mm(160))
    # shear_img = InlineImage(doc, shear_path, width=Mm(100))
    d_img = InlineImage(doc,d_path, width=Mm(160))
    pile_data={
        "SECtype": SECtype,
        "pile_jisuan": pile_jisuan,
        "pile_cb_pic": cb_img,
        "pile_d_pic": d_img
    }
    # 处理结论
    daxiao_cb = "≤" if max_cb_all_cases <= float(pile_allowable_f) else ">"
    daxiao_d = "≤" if max_d_all_cases <= float(pile_allowable_d) else ">"
    pile_res = f"{SECtype}在{max_cb_case_name}下组合应力最大值为{max_cb_all_cases:.2f} MPa{daxiao_cb}{pile_allowable_f}MPa，在{max_d_case_name}下变形最大值为 {max_d_all_cases:.2f}mm{daxiao_d}{pile_allowable_d}mm，{if_satisfied}要求；"
    return pile_data, pile_res, if_satisfied, max_cb_all_cases, max_d_all_cases


def Pile_cal_Stage(doc, version, pic_path, load_cases, pile_leixing, allowable_f_dict, gamma_0, SECtype, CofferDam_L):
    if version == "2026":
        minmax = "最小/最大"
    else:
        minmax = "Min/Max"
    # 容许应力
    pile_allowable_f = allowable_f_dict['pile']['f']
    pile_allowable_fv = allowable_f_dict['pile']['fv']
    # 容许位移
    pile_allowable_d = round(float(CofferDam_L)*2.5, 1)
    pile_jisuan = []
    # 获取截图
    API_DIST_FORCE_Unit("N", 'MM') # 应力为N和MM
    # 确保切换后处理状态，先截图
    Picture_Beamstress_Stage(pic_path, "gbz_cb_pic_max.jpg", SECtype, minmax, "合计", "Combined", "max", 45, 30)
    # 初始化追踪变量
    max_cb_all_cases = -1.0
    max_d_all_cases = -1.0
    max_cb_case_name = ""
    max_d_case_name = ""
    for case in load_cases:
        # case_name = f"工况{num_to_chinese_num(i+1)}"
        case_name = case
        # 施工阶段建模
        case_input = f"{case}:001(最后)"
        pile_stress_value_lst = Value_BeamStress_Stage("pilestress", SECtype, case_input, "合计(CS)")
        pile_deformed_value_lst = Value_Deformed_Stage("pile_displacement", SECtype, case_input, "合计(CS)")
        # 读取最大值
        pile_max_cb_value = max(pile_stress_value_lst['pilestress']['DATA'], key=lambda x: abs(float(x[3])))# 最大组合应力所在的结果表
        pile_max_cb = float(pile_max_cb_value[3])*gamma_0*1.25
        pile_max_shear_value = max(pile_stress_value_lst['pilestress']['DATA'], key=lambda x: abs(float(x[2])))# 最大剪应力所在的结果表
        pile_max_shear = float(pile_max_shear_value[2])*gamma_0*1.25
        pile_max_Dx_value = max(pile_deformed_value_lst['pile_displacement']['DATA'], key=lambda x: abs(float(x[2])))# 最大变形所在的结果表
        pile_max_Dy_value = max(pile_deformed_value_lst['pile_displacement']['DATA'], key=lambda x: abs(float(x[3])))# 最大变形所在的结果表
        pile_max_Dx = float(pile_max_Dx_value[2])*gamma_0
        pile_max_Dy = float(pile_max_Dy_value[3])*gamma_0
        dx_abs = abs(pile_max_Dx)
        dy_abs = abs(pile_max_Dy)
        if dx_abs >= dy_abs:
            pile_max_d = dx_abs
            pile_d_direction = "Dx"
        else:
            pile_max_d = dy_abs
            pile_d_direction= "Dy"
        if abs(pile_max_cb) <= float(pile_allowable_f) and pile_max_d <= float(pile_allowable_d):
            if_satisfied = '满足'
        else:
            if_satisfied = '不满足'
        # 加入列表
        pile_jisuan.append({
            'case': case_name,                          # 工况
            'type': pile_leixing,                       # 桩类型/规格
            'max_cb': round(pile_max_cb,2),             # 最大应力
            'limit_cb': pile_allowable_f,               # 容许应力
            'max_d': round(pile_max_d,2),               # 最大位移
            'limit_d': pile_allowable_d,                # 容许位移
            'check': if_satisfied                       # 验算结果
        })
        # 最不利情况更新
        current_cb_abs = abs(pile_max_cb)
        if current_cb_abs > max_cb_all_cases:
            max_cb_all_cases = current_cb_abs
            max_cb_case_name = case_name
        current_d_abs = pile_max_d
        if current_d_abs > max_d_all_cases:
            max_d_all_cases = current_d_abs
            max_d_case_name = case_name
    # 根据计算结果获取变形截图
    API_DIST_FORCE_Unit("N", 'MM') # 变形为MM
    # 施工阶段钢板桩应力截图BUG(激活多余结构组)，再截一次
    Picture_Beamstress_Stage(pic_path, "gbz_cb_pic_max.jpg", SECtype, minmax, "合计", "Combined", "max", 45, 30)
    Picture_Beamstress_Stage(pic_path, "gbz_cb_pic_min.jpg", SECtype, minmax, "合计", "Combined", "min", 45, 30)
    Picture_Deformed_Stage(pic_path, "gbz_d_pic_max.jpg", SECtype, minmax, "合计", pile_d_direction, "max", 45, 30)
    Picture_Deformed_Stage(pic_path, "gbz_d_pic_min.jpg", SECtype, minmax, "合计", pile_d_direction, "min", 45, 30)
    # 处理图片
    cb_max_path = os.path.join(pic_path, "gbz_cb_pic_max.jpg")
    cb_min_path = os.path.join(pic_path, "gbz_cb_pic_min.jpg")
    d_max_path = os.path.join(pic_path, "gbz_d_pic_max.jpg")
    d_min_path = os.path.join(pic_path, "gbz_d_pic_min.jpg")
    # 读取图片
    cb_max_img = InlineImage(doc, cb_max_path, width=Mm(160))
    cb_min_img = InlineImage(doc, cb_min_path, width=Mm(160))
    d_max_img = InlineImage(doc, d_max_path, width=Mm(160))
    d_min_img = InlineImage(doc, d_min_path, width=Mm(160))
    pile_data={
        "SECtype": SECtype,
        "pile_jisuan": pile_jisuan,
        "pile_cb_max_pic": cb_max_img,
        "pile_cb_min_pic": cb_min_img,
        "pile_d_max_pic": d_max_img,
        "pile_d_min_pic": d_min_img
    }
    # 处理结论
    daxiao_cb = "≤" if max_cb_all_cases <= float(pile_allowable_f) else ">"
    daxiao_d = "≤" if max_d_all_cases <= float(pile_allowable_d) else ">"
    pile_res = f"{SECtype}在{max_cb_case_name}下组合应力最大值为{max_cb_all_cases:.2f} MPa{daxiao_cb}{pile_allowable_f}MPa，在{max_d_case_name}下变形最大值为 {max_d_all_cases:.2f}mm{daxiao_d}{pile_allowable_d}mm，{if_satisfied}要求；"
    return pile_data, pile_res, if_satisfied, max_cb_all_cases, max_d_all_cases


def Waler_cal(doc, pic_path, load_cases, has_waler, WL_group_namelst, Waler_SECT_lst, allowable_f_dict, gamma_0, pbar=None, svar=None, root=None, step_weight=8):
    # 容许应力
    waler_allowable_f_lst = [v['f'] for v in allowable_f_dict['waler']]
    waler_allowable_fv_lst = [v['fv'] for v in allowable_f_dict['waler']]
    # 容许变形
    waler_allowable_d = ''

    WL_layers_data = []
    WL_res = []
    wl_summary = {}
    if has_waler:
        num_layers = len(WL_group_namelst)
        increment = step_weight / num_layers if num_layers > 0 else 0
        # 按层读取替换
        for i, WL_group_name in enumerate(WL_group_namelst):
            # 获取当前层围檩的容许应力
            waler_allowable_f = waler_allowable_f_lst[i]
            waler_allowable_fv = waler_allowable_fv_lst[i]             
            idx = i+1
            # 进程
            if svar:
                svar.set(f"正在进行模型结果验算\n正在验算第 {i+1}/{num_layers} 层围檩...")
            if pbar and root:
                pbar['value'] += increment
                root.update()
            # 初始化追踪变量
            layer_max_cb = -1.0
            layer_max_shear = -1.0
            layer_max_d = -1.0
            layer_max_cb_case = ""
            # 建立列表
            WLjisuan = []
            # 计算单层围檩在多个工况下的结果
            WL_d_direction = "Dx" # 初始方向
            for j, case in enumerate(load_cases):
                case_name = case
                # 整体建模
                WL_stress_value_lst = Value_BeamStress("WLstress", WL_group_name, "基本组合(CB)")
                WL_deformed_value_lst = Value_Deformed("WL_displacement", WL_group_name, "标准组合(CB)")
                # 读取最大值
                WL_max_cb_value = max(WL_stress_value_lst['WLstress']['DATA'], key=lambda x: abs(float(x[3])))# 最大组合应力所在的结果表
                WL_max_shear_value = max(WL_stress_value_lst['WLstress']['DATA'], key=lambda x: abs(float(x[2])))# 最大剪应力所在的结果表
                WL_max_cb = float(WL_max_cb_value[3])*gamma_0
                WL_max_shear = float(WL_max_shear_value[2])*gamma_0
                WL_max_Dx_value = max(WL_deformed_value_lst['WL_displacement']['DATA'], key=lambda x: abs(float(x[2])))# 最大变形所在的结果表
                WL_max_Dy_value = max(WL_deformed_value_lst['WL_displacement']['DATA'], key=lambda x: abs(float(x[3])))# 最大变形所在的结果表
                WL_max_Dx = float(WL_max_Dx_value[2])*gamma_0
                WL_max_Dy = float(WL_max_Dy_value[3])*gamma_0
                dx_abs = abs(WL_max_Dx)
                dy_abs = abs(WL_max_Dy)
                if dx_abs >= dy_abs:
                    WL_max_d = dx_abs
                    WL_d_direction = "Dx"
                else:
                    WL_max_d = dy_abs
                    WL_d_direction = "Dy"
                # WLjisuan.append([case,str(Waler_SECT_lst[i]),str(WL_max_cb),str(WL_max_shear),str(WL_cb_limit),str(WL_shear_limit),'','',''])
                # 计算表格内容
                WLjisuan.append({
                    "case": case_name,                  # 工况/阶段
                    "sect": Waler_SECT_lst[i],          # 围檩截面类型
                    "max_cb": round(WL_max_cb,2),       # 最大组合应力
                    "max_shear": round(WL_max_shear,2), # 最大剪应力
                    "limit_cb": waler_allowable_f,      # 组合应力限值
                    "limit_shear": waler_allowable_fv,  # 剪应力限值
                    "max_d": round(WL_max_d,2),         # 最大绝对变形
                    "limit_d": waler_allowable_d,       # 变形限值
                    "check_result": "",                 # 是否满足                  
                })
                # 获取当前层在所有工况下的最大应力
                current_cb_abs = abs(WL_max_cb)
                current_max_shear_abs = abs(WL_max_shear)
                current_d_abs = WL_max_d
                if current_cb_abs > layer_max_cb:
                    layer_max_cb = current_cb_abs
                    layer_max_cb_case = case_name
                if current_max_shear_abs > layer_max_shear:
                    layer_max_shear = current_max_shear_abs
                    layer_max_shear_case = case_name
                if current_d_abs > layer_max_d:
                    layer_max_d = current_d_abs
            # 获取截图
            API_DIST_FORCE_Unit("N", 'MM') # 应力为N和MM
            Picture_Beamstress(pic_path, f"wl{i+1}_cb_pic.jpg", [WL_group_namelst[i]], "基本组合", "Combined", 0, 90)
            Picture_Beamstress(pic_path, f"wl{i+1}_shear_pic.jpg", [WL_group_namelst[i]], "基本组合", "Ssz", 0, 90)
            Picture_Deformed(pic_path, f"WL{i+1}_d_pic.jpg", [WL_group_namelst[i]], "标准组合", WL_d_direction, 0, 90)
            # 处理图片
            cb_path = os.path.join(pic_path, f"wl{idx}_cb_pic.jpg")
            shear_path = os.path.join(pic_path, f"wl{idx}_shear_pic.jpg")
            d_path = os.path.join(pic_path, f"wl{idx}_d_pic.jpg")
            try:
                cb_img = InlineImage(doc, cb_path, width=Mm(160))
                shear_img = InlineImage(doc, shear_path, width=Mm(160))
                d_img = InlineImage(doc, d_path, width=Mm(160))
            except Exception as e:
                print(f"第 {idx} 层图片读取失败: {e}")
                cb_img = shear_img = d_img = "图片缺失"
            # 内容汇总
            WL_layers_data.append({
                "index": idx,
                "cn_index": num_to_chinese_num(idx),
                "WL_type": Waler_SECT_lst[i],
                "WLjisuan": WLjisuan,
                "WL_cb_pic":cb_img,
                "WL_shear_pic":shear_img,
                "WL_d_pic":d_img
            })
            # 处理结论
            daxiao_cb = "≤" if layer_max_cb <= float(waler_allowable_f) else ">"
            daxiao_shear = "≤" if layer_max_shear <= float(waler_allowable_fv) else ">"
            layer_summary = f"第{idx}层围檩在{layer_max_cb_case}下组合应力最大值为{layer_max_cb:.2f}MPa{daxiao_cb}{waler_allowable_f}MPa，在{layer_max_shear_case}下剪应力最大值为{layer_max_shear:.2f}MPa{daxiao_shear}{waler_allowable_fv}MPa；"
            WL_res.append(layer_summary)
            wl_summary[idx] = {'max_cb': layer_max_cb, 'max_shear': layer_max_shear, 'max_d': layer_max_d}
    return WL_layers_data, WL_res, wl_summary


def Waler_cal_Stage(doc, version, pic_path, load_cases, has_waler, WL_group_namelst, Waler_SECT_lst, allowable_f_dict, gamma_0, pbar=None, svar=None, root=None, step_weight=8):
    if version == "2026":
        minmax = "最小/最大"
    else:
        minmax = "Min/Max"
    # 容许应力
    waler_allowable_f_lst = [v['f'] for v in allowable_f_dict['waler']]
    waler_allowable_fv_lst = [v['fv'] for v in allowable_f_dict['waler']]
    # 容许变形
    waler_allowable_d = ''

    WL_layers_data = []
    WL_res = []
    wl_summary = {}
    if has_waler:
        num_layers = len(WL_group_namelst)
        increment = step_weight / num_layers if num_layers > 0 else 0
        # 按层读取替换
        for i, WL_group_name in enumerate(WL_group_namelst):
            # 获取当前层围檩的容许应力
            waler_allowable_f = waler_allowable_f_lst[i]
            waler_allowable_fv = waler_allowable_fv_lst[i]
            idx = i+1
            # 进程
            if svar:
                svar.set(f"正在进行模型结果验算\n正在验算第 {i+1}/{num_layers} 层围檩...")
            if pbar and root:
                pbar['value'] += increment
                root.update()
            # 初始化追踪变量
            layer_max_cb = -1.0
            layer_max_shear = -1.0
            layer_max_d = -1.0
            layer_max_cb_case = ""
            # 建立列表
            WLjisuan = []
            # 计算单层围檩在多个工况下的结果
            WL_d_direction = "Dx" # 初始方向
            for j, case in enumerate(load_cases):
                case_name = f"工况{num_to_chinese_num(j+1)}"
                # 施工阶段建模
                # 判断是否为前工况
                case_nums = re.findall(r'\d+', case)
                case_stage_num = int(case_nums[0]) if case_nums else 999
                if "第" in case and "层" in case:
                    if case_stage_num < idx or (case_stage_num == idx and "开挖" in case):
                        continue
                case_input = f"{case}:001(最后)"
                WL_stress_value_lst = Value_BeamStress_Stage("WLstress", WL_group_name, case_input, "合计(CS)")
                WL_deformed_value_lst = Value_Deformed_Stage("WL_displacement", WL_group_name, case_input, "合计(CS)")
                if not WL_stress_value_lst.get('WLstress', {}).get('DATA'):
                    continue
                # 读取最大值
                WL_max_cb_value = max(WL_stress_value_lst['WLstress']['DATA'], key=lambda x: abs(float(x[3])))# 最大组合应力所在的结果表
                WL_max_shear_value = max(WL_stress_value_lst['WLstress']['DATA'], key=lambda x: abs(float(x[2])))# 最大剪应力所在的结果表
                WL_max_cb = float(WL_max_cb_value[3])*gamma_0*1.25
                WL_max_shear = float(WL_max_shear_value[2])*gamma_0*1.25
                WL_max_Dx_value = max(WL_deformed_value_lst['WL_displacement']['DATA'], key=lambda x: abs(float(x[2])))# 最大变形所在的结果表
                WL_max_Dy_value = max(WL_deformed_value_lst['WL_displacement']['DATA'], key=lambda x: abs(float(x[3])))# 最大变形所在的结果表
                WL_max_Dx = WL_max_Dx_value[2]
                WL_max_Dy = WL_max_Dy_value[3]
                dx_abs = abs(float(WL_max_Dx))
                dy_abs = abs(float(WL_max_Dy))
                if dx_abs >= dy_abs:
                    WL_max_d = dx_abs
                    WL_d_direction = "Dx"
                else:
                    WL_max_d = dy_abs
                    WL_d_direction = "Dy"     # 验算结果判断
                # WLjisuan.append([case,str(Waler_SECT_lst[i]),str(WL_max_cb),str(WL_max_shear),str(WL_cb_limit),str(WL_shear_limit),'','',''])
                # 计算表格内容
                WLjisuan.append({
                    "case": case_name,                  # 工况/阶段
                    "sect": Waler_SECT_lst[i],          # 围檩截面类型
                    "max_cb": round(WL_max_cb,2),       # 最大组合应力
                    "max_shear": round(WL_max_shear,2), # 最大剪应力"
                    "limit_cb": waler_allowable_f,      # 组合应力限值
                    "limit_shear": waler_allowable_fv,  # 剪应力限值
                    "max_d": round(WL_max_d,2),         # 最大相对变形
                    "limit_d": waler_allowable_d,       # 变形限值
                    "check_result": "",                 # 是否满足
                })
                # 获取当前层在所有工况下的最大应力
                current_cb_abs = abs(WL_max_cb)
                current_max_shear_abs = abs(WL_max_shear)
                current_d_abs = WL_max_d
                if current_cb_abs > layer_max_cb:
                    layer_max_cb = current_cb_abs
                    layer_max_cb_case = case_name
                if current_max_shear_abs > layer_max_shear:
                    layer_max_shear = current_max_shear_abs
                    layer_max_shear_case = case_name
                if current_d_abs > layer_max_d:
                    layer_max_d = current_d_abs

            # 获取截图
            API_DIST_FORCE_Unit("N", 'MM') # 应力为N和MM
            Picture_Beamstress_Stage(pic_path, f"wl{i+1}_cb_max_pic.jpg", [WL_group_namelst[i]],  minmax, "合计", "Combined", "max", 0, 90)
            Picture_Beamstress_Stage(pic_path, f"wl{i+1}_cb_min_pic.jpg", [WL_group_namelst[i]],  minmax, "合计", "Combined", "min", 0, 90)
            Picture_Beamstress_Stage(pic_path, f"wl{i+1}_shear_max_pic.jpg", [WL_group_namelst[i]],  minmax, "合计", "Ssz", "max", 0, 90)
            Picture_Beamstress_Stage(pic_path, f"wl{i+1}_shear_min_pic.jpg", [WL_group_namelst[i]],  minmax, "合计", "Ssz", "min", 0, 90)
            Picture_Deformed_Stage(pic_path, f"WL{i+1}_d_max_pic.jpg", [WL_group_namelst[i]],  minmax, "合计", WL_d_direction, "max", 0, 90)
            Picture_Deformed_Stage(pic_path, f"WL{i+1}_d_min_pic.jpg", [WL_group_namelst[i]],  minmax, "合计", WL_d_direction, "min", 0, 90)
            # 处理图片
            cb_max_path = os.path.join(pic_path, f"wl{i+1}_cb_max_pic.jpg")
            cb_min_path = os.path.join(pic_path, f"wl{i+1}_cb_min_pic.jpg")
            shear_max_path = os.path.join(pic_path, f"wl{i+1}_shear_max_pic.jpg")
            shear_min_path = os.path.join(pic_path, f"wl{i+1}_shear_min_pic.jpg")
            d_max_path = os.path.join(pic_path, f"wl{i+1}_d_max_pic.jpg")
            d_min_path = os.path.join(pic_path, f"wl{i+1}_d_min_pic.jpg")
            try:
                cb_max_img = InlineImage(doc, cb_max_path, width=Mm(160))
                cb_min_img = InlineImage(doc, cb_min_path, width=Mm(160))
                shear_max_img = InlineImage(doc, shear_max_path, width=Mm(160))
                shear_min_img = InlineImage(doc, shear_min_path, width=Mm(160))
                d_max_img = InlineImage(doc, d_max_path, width=Mm(160))
                d_min_img = InlineImage(doc, d_min_path, width=Mm(160))
            except Exception as e:
                print(f"第 {i+1} 层图片读取失败: {e}")
                cb_max_img = cb_min_img = shear_max_img = shear_min_img = d_max_img = d_min_img = "图片缺失"
            # 内容汇总 
            WL_layers_data.append({
                "index": idx,
                "cn_index": num_to_chinese_num(idx),
                "WL_type": Waler_SECT_lst[i],
                "WLjisuan": WLjisuan,
                "WL_cb_max_pic":cb_max_img,
                "WL_cb_min_pic":cb_min_img,
                "WL_shear_max_pic":shear_max_img,
                "WL_shear_min_pic":shear_min_img,
                "WL_d_max_pic":d_max_img,
                "WL_d_min_pic":d_min_img
            })
            # 处理结论
            daxiao_cb = "≤" if layer_max_cb <= float(waler_allowable_f) else ">"
            daxiao_shear = "≤" if layer_max_shear <= float(waler_allowable_fv) else ">"
            layer_summary = f"第{idx}层围檩在{layer_max_cb_case}下组合应力最大值为{layer_max_cb:.2f}MPa{daxiao_cb}{waler_allowable_f}MPa，在{layer_max_shear_case}下剪应力最大值为{layer_max_shear:.2f}MPa{daxiao_shear}{waler_allowable_fv}MPa；"
            WL_res.append(layer_summary)
            wl_summary[idx] = {'max_cb': layer_max_cb, 'max_shear': layer_max_shear, 'max_d': layer_max_d}
    return WL_layers_data, WL_res, wl_summary


# 截面参数查询函数, A, Ix, Iy, Wx, Wy, ix, iy
def Section_parameters(section):
    section_dict = {
    # 钢管桩
    '377X6钢管桩': [6993.2, 120350095.7, 120350095.7, 638462.0, 638462.0, 131.2, 131.2],
    '426X6钢管桩': [7916.8, 174601363.1, 174601363.1, 819724.7, 819724.7, 148.5, 148.5],
    '600X8钢管桩': [14878.6, 651919984.3, 651919984.3, 2173066.6, 2173066.6, 209.3, 209.3],
    '630X8钢管桩': [15632.6, 756123722.3, 756123722.3, 2400392.8, 2400392.8, 219.9, 219.9],
    '820X10钢管桩': [25446.9, 2087282013.0, 2087282013.0, 5090931.7, 5090931.7, 286.4, 286.4],
    '1000X10钢管桩': [31101.8, 3810744034.8, 3810744034.8, 7621488.1, 7621488.1, 350.0, 350.0],
    '1000X12钢管桩': [37246.7, 4545441027.1, 4545441027.1, 9090882.1, 9090882.1, 349.3, 349.3],
    # 型钢
    'I40a': [8610.0, 217000000.0, 6600000.0, 1085000.0, 92957.7, 158.7, 27.7],
    'HW400X400': [21869.0, 664550000.0, 224100000.0, 3322750.0, 1120500.0, 174.3, 101.2],
    'HM588X300': [18721.0, 1128300000.0, 90090000.0, 3837755.1, 600600.0, 245.5, 69.4],
    'HN700X300': [23154.0, 1936200000.0, 108140000.0, 5532000.0, 720933.3, 289.2, 68.3],
    'HN900X300': [30582.0, 3972400000.0, 126310000.0, 8827555.5, 842066.7, 360.4, 64.2],
    '2I20a': [7100.0, 47400000.0, 12800000.0, 474000.0, 128000.0, 81.6, 42.5],
    '2I22a': [8400.0, 68200000.0, 16800000.0, 620000.0, 152700.0, 90.0, 44.7],
    '2I25a': [10600.0, 100400000.0, 24400000.0, 803200.0, 210344.0, 101.7, 48.0],
    '2I32a': [15400.0, 222000000.0, 52400000.0, 1387500.0, 403076.0, 128.6, 58.3],
    '2I36a': [18400.0, 316000000.0, 76200000.0, 1755555.5, 560294.0, 143.8, 64.4],
    '2I40a': [21800.0, 434000000.0, 104000000.0, 2170000.0, 732394.0, 158.8, 69.1],
    '2I45a': [28600.0, 644000000.0, 162000000.0, 2862222.2, 1080000.0, 177.7, 75.3],
    '2I50a': [37200.0, 930000000.0, 238000000.0, 3720000.0, 1506329.0, 197.7, 80.0],
    '2I56a': [46800.0, 1312000000.0, 354000000.0, 4685714.2, 2132530.0, 220.4, 87.0],
    '2I63a': [59600.0, 188000000.0, 526000000.0, 5968254.0, 2988636.0, 246.3, 94.0],
    '2HW400X400': [43738.0, 1329100000.0, 2197811016.0, 6645500.0, 5494527.5, 174.3, 224.2],
    '2HM588X300': [37442.0, 2256600000.0, 1022630738.0, 7675510.0, 3408769.1, 245.5, 165.3],
    '2HN600X200': [30100.0, 1474980000.0, 102450000.0, 4916600.0, 512250.0, 236.6, 58.3],
    '2HN700X300': [46308.0, 3872400000.0, 1258242224.0, 11064000.0, 4194140.7, 289.2, 164.8],
    '2HN800X300': [52100.0, 5618400000.0, 142450000.0, 14046000.0, 474833.3, 326.5, 52.3],
    '2HN900X300': [61164.0, 7944800000.0, 1628858461.0, 17655111.1, 5429528.2, 360.4, 163.2],
    '2HN1000X300': [68200.0, 12528000000.0, 182450000.0, 25056000.0, 608166.7, 398.2, 51.7],
    }
    return section_dict[section]


# 型钢的计算参数
def Profile_Steel_stability_param(section_dict, section, L, N, My, Mz, fy = 235):
    # A, Ix, Iy, Wx, Wy, ix, iy = Section_parameters(section)
    A  = section_dict[section]['A' ]
    Wx = section_dict[section]['Wx']
    Wy = section_dict[section]['Wy']
    ix = section_dict[section]['ix']
    iy = section_dict[section]['iy']
    print(section, L, N, My, Mz)
    Phix = stability_coefficient("b", L, ix, 235, 206000)
    Phiy = stability_coefficient("b", L, iy, 235, 206000)
    Bmx = 1.0
    Bmy = 1.0
    Btx = 1.0
    Bty = 1.0
    gamax = 1.05
    gamay = 1.2
    lambdax = L/ix
    lambday = L/iy
    Nex = (math.pi**2)*206000*A/(1.1*(lambdax**2))
    Ney = (math.pi**2)*206000*A/(1.1*(lambday**2))
    ek = pow(235/fy, 0.5)
    Phibx = min(1.07-((lambday**2)/(44000*(ek**2))), 1)
    Phiby = 1.0
    n = 0.7 if ('2' in section and ('C' not in section or '[' not in section)) else 1.0
    print(Phix, Phiy, Nex, Ney, Phibx, Phiby, n)
    sigmax = round(N*1000/(Phix*A) + (Bmx*My*1000000)/(gamax*Wx*(1-0.8*(N*1000/Nex))) + n*(Bty*Mz*1000000)/(Phiby*Wy),1)
    sigmay = round(N*1000/(Phiy*A) + (Bmy*Mz*1000000)/(gamay*Wy*(1-0.8*(N*1000/Ney))) + n*(Btx*My*1000000)/(Phibx*Wx),1)
    print(sigmax, sigmay)
    return sigmax, sigmay


# 钢管桩的截面特性参数
def SteelPile_section_param(D, t, L):
    # 单位一定是N和MM, 见2.1小节单位修改函数
    D = round(float(D),1) # 直径
    t = round(float(t),1) # 壁厚
    L = round(float(L),1) # 桩长
    # 根据参数计算
    Area = math.pi*(D*D-(D-2*t)*(D-2*t))/4
    Ix = math.pi*(D**4-(D-2*t)**4)/64
    ix = math.sqrt(Ix/Area)
    Wx = Ix/(D/2)
    return Area, Ix, ix, Wx


# 钢管桩的稳定计算参数
def SteelPile_stability_param(Area, ix, Wx, L, Fx, My, Mz):
    # ele, Fx, My, Mz, D, t, L = model_post_dict["model_post_gangguanzhuang"]
    # 单位一定是KN和M, 见函数 Value_BeamForce 中的单位定义
    print(Area, ix, Wx, L, Fx, My, Mz)
    Fx = abs(round(float(Fx),1)) # Fx
    My = abs(round(float(My),1)) # My
    Mz = abs(round(float(Mz),1)) # Mz
    quote_lambda = L/ix
    Mm = math.sqrt(math.pow(My, 2) + math.pow(Mz, 2))
    Phi = stability_coefficient("b", L, ix, 235, 206000)
    Nex = (math.pi**2)*206000*Area/(1.1*(quote_lambda**2))
    sigma = ((Fx*1000/(Phi*Area))+((1*Mm*1000000)/(1.15*Wx*(1-0.8*(Fx*1000/Nex)))))
    print(quote_lambda, Mm, Phi, Nex, sigma)
    return Phi, Nex, round(sigma,1)


def Cal_strut_Stability(Strut_section_dict, DC_Value_dict, DC_Section_dict, DC_L_dict, XC_Value_dict, XC_Section_dict, XC_L_dict):
    DC_sigma_dict = {}
    DC_Force_dict = {}
    for key, value in DC_Value_dict.items():
        DC_section = DC_Section_dict[key]
        DC_L = DC_L_dict[key]*1000
        DC_Fxmax, DC_Mymax, DC_Mzmax = [abs(float(x)) for x in value]
        if '钢管桩' not in DC_section:
            sigmax, sigmay = Profile_Steel_stability_param(Strut_section_dict, DC_section, DC_L, DC_Fxmax, DC_Mymax, DC_Mzmax)
            DC_sigma_dict[key] = [sigmax, sigmay]
        else:
            match = re.search(r'(\d+)X(\d+)', DC_section)
            D, t = match.groups()
            Area, Ix, ix, Wx = SteelPile_section_param(D, t, DC_L)
            Phi, Nex, sigma = SteelPile_stability_param(Area, ix, Wx, DC_L, DC_Fxmax, DC_Mymax, DC_Mzmax)
            DC_sigma_dict[key] = [sigma]
        DC_Force_dict[key] = [DC_Fxmax, DC_Mymax, DC_Mzmax]
    XC_sigma_dict = {}
    XC_Force_dict = {}
    for key, value in XC_Value_dict.items():
        XC_section = XC_Section_dict[key]
        XC_L = XC_L_dict[key]*1000
        XC_Fxmax, XC_Mymax, XC_Mzmax = [abs(float(x)) for x in value]
        if '钢管桩' not in XC_section:
            sigmax, sigmay = Profile_Steel_stability_param(Strut_section_dict, XC_section, XC_L, XC_Fxmax, XC_Mymax, XC_Mzmax)
            XC_sigma_dict[key] = [sigmax, sigmay]
        else:
            match = re.search(r'(\d+)X(\d+)', XC_section)
            D, t = match.groups()
            Area, Ix, ix, Wx = SteelPile_section_param(D, t, XC_L)
            Phi, Nex, sigma = SteelPile_stability_param(Area, ix, Wx, XC_L, XC_Fxmax, XC_Mymax, XC_Mzmax)
            XC_sigma_dict[key] = [sigma]
        XC_Force_dict[key] = [XC_Fxmax, XC_Mymax, XC_Mzmax]
    # print('每一层对撑的内力值', DC_Force_dict)
    # print('每一层斜撑的内力值', XC_Force_dict)
    # print('每一层对撑的应力值', DC_sigma_dict)
    # print('每一层斜撑的应力值', XC_sigma_dict)
    return DC_Force_dict, XC_Force_dict, DC_sigma_dict, XC_sigma_dict


def Strut_Force_Picture(pic_path, Layer_lst, pbar=None, svar=None, root=None, step_weight=8):
    API_DIST_FORCE_Unit("KN", 'M')
    num_layers = len(Layer_lst)
    increment = step_weight / num_layers if num_layers > 0 else 0
    for i, lst in enumerate(Layer_lst):
        idx = i+1
        if svar:
            svar.set(f"正在进行模型结果验算\n正在验算第 {idx}/{num_layers} 层内支撑...")
        if pbar and root:
            pbar['value'] += increment
            root.update()
        Picture_BeamForce(pic_path, f"Fx轴力图{idx}.jpg", lst, "基本组合", "Fx", 30, 45)
        Picture_BeamForce(pic_path, f"My轴力图{idx}.jpg", lst, "基本组合", "My", 30, 45)
        Picture_BeamForce(pic_path, f"Mz轴力图{idx}.jpg", lst, "基本组合", "Mz", 30, 45)


def Strut_Force_Picture_Stage(pic_path, version, Layer_lst, pbar=None, svar=None, root=None, step_weight=8):
    if version== "2026":
        minmax = "最小/最大"
    else:
        minmax = "Min/Max"
    API_DIST_FORCE_Unit("KN", 'M')
    num_layers = len(Layer_lst)
    increment = step_weight / num_layers if num_layers > 0 else 0
    for i, lst in enumerate(Layer_lst):
        idx = i+1
        if svar:
            svar.set(f"正在进行模型结果验算\n正在验算第 {idx}/{num_layers} 层内支撑...")
        if pbar and root:
            pbar['value'] += increment
            root.update()
        Picture_BeamForce_Stage(pic_path, f"Fx轴力图{idx}_max.jpg", lst,  minmax, "合计", "Fx", "max", 30, 45)
        Picture_BeamForce_Stage(pic_path, f"Fx轴力图{idx}_min.jpg", lst,  minmax, "合计", "Fx", "min", 30, 45)
        Picture_BeamForce_Stage(pic_path, f"My轴力图{idx}_max.jpg", lst,  minmax, "合计", "My", "max", 30, 45)
        Picture_BeamForce_Stage(pic_path, f"My轴力图{idx}_min.jpg", lst,  minmax, "合计", "My", "min", 30, 45)
        Picture_BeamForce_Stage(pic_path, f"Mz轴力图{idx}_max.jpg", lst,  minmax, "合计", "Mz", "max", 30, 45)
        Picture_BeamForce_Stage(pic_path, f"Mz轴力图{idx}_min.jpg", lst,  minmax, "合计", "Mz", "min", 30, 45)


# 获取内支撑内力值
def Strut_Force_Value(DC_group_namelst, XC_group_namelst, gamma_0):
    # 对撑
    DC_Value_dict = {}
    if DC_group_namelst != []:
        for x in DC_group_namelst:
            DC_Fx_value_lst = Value_BeamForce("DCforce", x, "基本组合(CB)")
            DC_Fxmax_value = max(DC_Fx_value_lst['DCforce']['DATA'], key=lambda x: abs(float(x[2])))# 最大Fx所在的结果表
            DC_Mymax_value = max(DC_Fx_value_lst['DCforce']['DATA'], key=lambda x: abs(float(x[3])))# 最大My所在的结果表
            DC_Mzmax_value = max(DC_Fx_value_lst['DCforce']['DATA'], key=lambda x: abs(float(x[4])))# 最大Mz所在的结果表
            DC_Fxmax = float(DC_Fxmax_value[2])*gamma_0
            DC_Mymax = float(DC_Mymax_value[3])*gamma_0
            DC_Mzmax = float(DC_Mzmax_value[4])*gamma_0
            # print(x, 'Fxmax:', DC_Fxmax)
            # print(x, 'Mymax:', DC_Mymax)
            # print(x, 'Mzmax:', DC_Mzmax)
            DC_Value_dict[x] = [round(DC_Fxmax,2), round(DC_Mymax,2), round(DC_Mzmax,2)]
    else:
        # DC_Fxmax, DC_Mymax, DC_Mzmax = 0, 0, 0
        print('无对撑')
    # 斜撑
    XC_Value_dict = {}
    if XC_group_namelst != []:
        for x in XC_group_namelst:
            XC_Fx_value_lst = Value_BeamForce("XCforce", x, "基本组合(CB)")
            XC_Fxmax_value = max(XC_Fx_value_lst['XCforce']['DATA'], key=lambda x: abs(float(x[2])))# 最大Fx所在的结果表
            XC_Mymax_value = max(XC_Fx_value_lst['XCforce']['DATA'], key=lambda x: abs(float(x[3])))# 最大My所在的结果表
            XC_Mzmax_value = max(XC_Fx_value_lst['XCforce']['DATA'], key=lambda x: abs(float(x[4])))# 最大Mz所在的结果表
            XC_Fxmax = float(XC_Fxmax_value[2])*gamma_0
            XC_Mymax = float(XC_Mymax_value[3])*gamma_0
            XC_Mzmax = float(XC_Mzmax_value[4])*gamma_0
            # print(x, 'Fxmax:', XC_Fxmax)
            # print(x, 'Mymax:', XC_Mymax)
            # print(x, 'Mzmax:', XC_Mzmax)
            XC_Value_dict[x] = [round(XC_Fxmax,2), round(XC_Mymax,2), round(XC_Mzmax,2)]
    else:
        # XC_Fxmax, XC_Mymax, XC_Mzmax = 0, 0, 0
        print('无斜撑')
    print('每一层对撑内力值', DC_Value_dict)
    print('每一层斜撑内力值', XC_Value_dict)
    return DC_Value_dict, XC_Value_dict


def Strut_Force_Value_Stage(version, DC_group_namelst, XC_group_namelst, gamma_0):
    if version == "2026":
        minmax1 = "最小/最大:最大"
        minmax2 = "最小/最大:最小"
    else:
        minmax1 = "Min/Max:Max"
        minmax2 = "Min/Max:Min"
    # 对撑
    DC_Value_dict = {}
    if DC_group_namelst != []:
        for x in DC_group_namelst:
            DC_Fx_value_lst = Value_BeamForce_Stage("DCforce", x, [minmax1, minmax2], "合计(CS)")
            DC_Fxmax_value = max(DC_Fx_value_lst['DCforce']['DATA'], key=lambda x: abs(float(x[2])))# 最大Fx所在的结果表
            DC_Mymax_value = max(DC_Fx_value_lst['DCforce']['DATA'], key=lambda x: abs(float(x[3])))# 最大My所在的结果表
            DC_Mzmax_value = max(DC_Fx_value_lst['DCforce']['DATA'], key=lambda x: abs(float(x[4])))# 最大Mz所在的结果表
            DC_Fxmax = float(DC_Fxmax_value[2])*gamma_0*1.25
            DC_Mymax = float(DC_Mymax_value[3])*gamma_0*1.25
            DC_Mzmax = float(DC_Mzmax_value[4])*gamma_0*1.25
            # print(x, 'Fxmax:', DC_Fxmax)
            # print(x, 'Mymax:', DC_Mymax)
            # print(x, 'Mzmax:', DC_Mzmax)
            DC_Value_dict[x] = [round(DC_Fxmax,2), round(DC_Mymax,2), round(DC_Mzmax,2)]
    else:
        # DC_Fxmax, DC_Mymax, DC_Mzmax = 0, 0, 0
        print('无对撑')
    # 斜撑
    XC_Value_dict = {}
    if XC_group_namelst != []:
        for x in XC_group_namelst:
            XC_Fx_value_lst = Value_BeamForce_Stage("XCforce", x, [minmax1,minmax2], "合计(CS)")
            XC_Fxmax_value = max(XC_Fx_value_lst['XCforce']['DATA'], key=lambda x: abs(float(x[2])))# 最大Fx所在的结果表
            XC_Mymax_value = max(XC_Fx_value_lst['XCforce']['DATA'], key=lambda x: abs(float(x[3])))# 最大My所在的结果表
            XC_Mzmax_value = max(XC_Fx_value_lst['XCforce']['DATA'], key=lambda x: abs(float(x[4])))# 最大Mz所在的结果表
            XC_Fxmax = float(XC_Fxmax_value[2])*gamma_0*1.25
            XC_Mymax = float(XC_Mymax_value[3])*gamma_0*1.25
            XC_Mzmax = float(XC_Mzmax_value[4])*gamma_0*1.25
            # print(x, 'Fxmax:', XC_Fxmax)
            # print(x, 'Mymax:', XC_Mymax)
            # print(x, 'Mzmax:', XC_Mzmax)
            XC_Value_dict[x] = [round(XC_Fxmax,2), round(XC_Mymax,2), round(XC_Mzmax,2)]
    else:
        # XC_Fxmax, XC_Mymax, XC_Mzmax = 0, 0, 0
        print('无斜撑')
    print('每一层对撑内力值', DC_Value_dict)
    print('每一层斜撑内力值', XC_Value_dict)
    return DC_Value_dict, XC_Value_dict


def Strut_cal(doc, pic_path, has_strut, Strut_section_dict, XC_group_namelst,DC_group_namelst, DC_Section_dict, XC_Section_dict, DC_L_dict, XC_L_dict, allowable_f_dict, gamma_0, pbar=None, svar=None, root=None, step_weight=8):
    # 容许应力（按层号索引）
    strut_dc_allowable_f_dict = allowable_f_dict['strut_dc']
    strut_xc_allowable_f_dict = allowable_f_dict['strut_xc']
    # 容许变形
    strut_dc_allowable_d = ''
    strut_xc_allowable_d = ''
    # 预设参数
    has_pile = False
    has_profile = False
    strut_layers = []
    pile_stability = []
    profile_stability = []
    Strut_res = []
    if_satisfied = "满足"
    DC_sigma_dict = {}
    XC_sigma_dict = {}
    if has_strut:
        # 判断构成类型
        all_sects = list(DC_Section_dict.values()) + list(XC_Section_dict.values())
        print(all_sects)
        has_pile = any("钢管" in s for s in all_sects)
        has_profile = any("钢管" not in s for s in all_sects)
        print(has_pile, has_profile)
        # 将对撑和斜撑按层分组
        group_namelst = DC_group_namelst + XC_group_namelst
        Layer_lst = group_by_layer(group_namelst, r'第(\d+)层')
        print(Layer_lst)
        # 获取图片
        Strut_Force_Picture(pic_path, Layer_lst,pbar=pbar, svar=svar, root=root, step_weight=step_weight)
        # 获取内力值
        DC_Value_dict, XC_Value_dict = Strut_Force_Value(DC_group_namelst, XC_group_namelst, gamma_0)
        # 计算稳定值
        DC_Force_dict, XC_Force_dict, DC_sigma_dict, XC_sigma_dict = Cal_strut_Stability(Strut_section_dict, DC_Value_dict, DC_Section_dict, DC_L_dict, XC_Value_dict, XC_Section_dict, XC_L_dict)
        # 遍历每层组合文字及图片
        for i, layer_group in enumerate(Layer_lst):
            match = re.search(r'第(\d+)层', layer_group[0])
            layer_num = int(match.group(1)) if match else i + 1
            cn_index = num_to_chinese_num(layer_num)
            xc_type = next((XC_Section_dict[k] for k in layer_group if k in XC_Section_dict), None)
            zc_type = next((DC_Section_dict[k] for k in layer_group if k in DC_Section_dict), None)
            has_xc = bool(xc_type)
            has_zc = bool(zc_type)
            xc = has_xc and not has_zc
            zc = has_zc and not has_xc
            zh = has_xc and has_zc
            print(layer_num,xc, zc, zh)
            # 图片处理
            img_paths = {
                "fx": os.path.join(pic_path, f"fx轴力图{layer_num}.jpg"),
                "my": os.path.join(pic_path, f"my轴力图{layer_num}.jpg"),
                "mz": os.path.join(pic_path, f"mz轴力图{layer_num}.jpg")
            }
            try:
                fx_img = InlineImage(doc, img_paths["fx"], width=Mm(160))
                my_img = InlineImage(doc, img_paths["my"], width=Mm(160))
                mz_img = InlineImage(doc, img_paths["mz"], width=Mm(160))
            except:
                fx_img = my_img = mz_img = "图片未生成"
            layer = {
                'layer_num': layer_num,
                'cn_index': cn_index,
                'xc_type': xc_type,
                'zc_type': zc_type,
                'xc': xc,
                'zc': zc,
                'zh': zh,
                "Fx_pic": fx_img,
                "My_pic": my_img,
                "Mz_pic": mz_img
            }
            strut_layers.append(layer)
        # 获取表格参数
        for group_names, force_dict, sigma_dict, len_dict in [
            (XC_group_namelst, XC_Force_dict, XC_sigma_dict, XC_L_dict),
            (DC_group_namelst, DC_Force_dict, DC_sigma_dict, DC_L_dict)
        ]:
            for key in group_names:
                layer_match = re.search(r'第(\d+)层', key)
                layer_num = int(layer_match.group(1)) if layer_match else None
                if '对撑' in key:
                    Allowable = strut_dc_allowable_f_dict.get(layer_num, {}).get('f')
                elif '斜撑' in key:
                    Allowable = strut_xc_allowable_f_dict.get(layer_num, {}).get('f')
                else:
                    Allowable = 215.0
                if Allowable is None:
                    Allowable = 215.0
                sigma_lst = sigma_dict[key]
                # 解构内力值
                f_max, m_y, m_z = force_dict[key]
                length = len_dict[key]
                # 基础数据行
                row = {
                    "name": key,      # 支撑名称 
                    "Fx": f_max,      # 轴力 
                    "Mx": m_y,        # 模板中的 Mx 对应模型中的 My 
                    "My": m_z,        # 模板中的 My 对应模型中的 Mz 
                    "l0": length,     # 计算长度 
                    "limit": Allowable, # 容许应力
                }
                # 根据应力列表长度判断截面类型 
                if len(sigma_lst) == 2:            # 型钢（双向压弯）
                    sigmax, sigmay = sigma_lst
                    row["in_stability"] = sigmax   # 平面内应力 
                    row["out_stability"] = sigmay  # 平面外应力 
                    row["if_satisfied"] = '满足' if sigmax <= abs(Allowable) and sigmay <= Allowable else '不满足' 
                    profile_stability.append(row)
                elif len(sigma_lst) == 1:          # 钢管（圆管）
                    sigma = sigma_lst[0]
                    row["ylz"] = sigma             # 综合应力值 
                    row["if_satisfied"] = '满足' if sigma <= abs(Allowable) else '不满足' 
                    pile_stability.append(row)
                # 全局满足判定
                if row["if_satisfied"] == "不满足":
                    if_satisfied = "不满足"
        # 收集钢管支撑的最不利信息
        for row in pile_stability:
            row_allowable = row['limit']
            daxiao_ylz = "≤" if row['ylz'] <= abs(row_allowable) else ">"
            Strut_res.append(
                f"{row['name']}最不利工况下应力值为{row['ylz']:.2f}MPa{daxiao_ylz}{row_allowable:.2f}MPa，{row['if_satisfied']}要求；"
            )
        # 收集型钢支撑的最不利信息
        for row in profile_stability:
            row_allowable = row['limit']
            daxiao_in = "≤" if row['in_stability'] <= abs(row_allowable) else ">"
            daxiao_out = "≤" if row['out_stability'] <= abs(row_allowable) else ">"
            # 取平面内和平面外应力的较大值作为描述参考
            Strut_res.append(
                f"{row['name']}最不利工况下平面内稳定性应力值为{row['in_stability']:.2f}MPa{daxiao_in}{row_allowable:.2f}MPa，平面外稳定性应力值为{row['out_stability']:.2f}MPa{daxiao_out}{row_allowable:.2f}MPa，{row['if_satisfied']}要求；"
            )
    return has_pile, has_profile, strut_layers, pile_stability, profile_stability, Strut_res, if_satisfied, DC_sigma_dict, XC_sigma_dict


def Strut_cal_Stage(doc, version, pic_path, has_strut, Strut_section_dict, XC_group_namelst,DC_group_namelst, DC_Section_dict, XC_Section_dict, DC_L_dict, XC_L_dict, allowable_f_dict, gamma_0, pbar=None, svar=None, root=None, step_weight=8):
    # 容许应力（按层号索引）
    strut_dc_allowable_f_dict = allowable_f_dict['strut_dc']
    strut_xc_allowable_f_dict = allowable_f_dict['strut_xc']
    # 容许变形
    strut_dc_allowable_d = ''
    strut_xc_allowable_d = ''
    # 预设参数
    has_pile = False
    has_profile = False
    strut_layers = []
    pile_stability = []
    profile_stability = []
    Strut_res = []
    if_satisfied = "满足"
    if has_strut:
        all_sects = list(DC_Section_dict.values()) + list(XC_Section_dict.values())
        print(all_sects)
        # 判断构成类型
        has_pile = any("钢管" in s for s in all_sects)
        has_profile = any("钢管" not in s for s in all_sects)
        print(has_pile, has_profile)
        # 将对撑和斜撑按层分组
        group_namelst = DC_group_namelst + XC_group_namelst
        Layer_lst = group_by_layer(group_namelst, r'第(\d+)层')
        print(Layer_lst)
        # 获取图片
        Strut_Force_Picture_Stage(pic_path, version, Layer_lst,pbar=pbar, svar=svar, root=root, step_weight=step_weight)
        # 获取内力值
        DC_Value_dict, XC_Value_dict = Strut_Force_Value_Stage(version, DC_group_namelst, XC_group_namelst, gamma_0)
        # 计算稳定值
        DC_Force_dict, XC_Force_dict, DC_sigma_dict, XC_sigma_dict = Cal_strut_Stability(Strut_section_dict, DC_Value_dict, DC_Section_dict, DC_L_dict, XC_Value_dict, XC_Section_dict, XC_L_dict)
        # 遍历每层组合文字及图片
        for i, layer_group in enumerate(Layer_lst):
            match = re.search(r'第(\d+)层', layer_group[0])
            layer_num = int(match.group(1)) if match else i + 1
            cn_index = num_to_chinese_num(layer_num)
            xc_type = next((XC_Section_dict[k] for k in layer_group if k in XC_Section_dict), None)
            zc_type = next((DC_Section_dict[k] for k in layer_group if k in DC_Section_dict), None)
            has_xc = bool(xc_type)
            has_zc = bool(zc_type)
            xc = has_xc and not has_zc
            zc = has_zc and not has_xc
            zh = has_xc and has_zc
            print(layer_num,xc, zc, zh)
            # 图片处理
            img_paths = {
                "fx_max": os.path.join(pic_path, f"fx轴力图{layer_num}_max.jpg"),
                "fx_min": os.path.join(pic_path, f"fx轴力图{layer_num}_min.jpg"),
                "my_max": os.path.join(pic_path, f"my轴力图{layer_num}_max.jpg"),
                "my_min": os.path.join(pic_path, f"my轴力图{layer_num}_min.jpg"),
                "mz_max": os.path.join(pic_path, f"mz轴力图{layer_num}_max.jpg"),
                "mz_min": os.path.join(pic_path, f"mz轴力图{layer_num}_min.jpg")
            }
            try:
                fx_max_img = InlineImage(doc, img_paths["fx_max"], width=Mm(160))
                fx_min_img = InlineImage(doc, img_paths["fx_min"], width=Mm(160))
                my_max_img = InlineImage(doc, img_paths["my_max"], width=Mm(160))
                my_min_img = InlineImage(doc, img_paths["my_min"], width=Mm(160))
                mz_max_img = InlineImage(doc, img_paths["mz_max"], width=Mm(160))
                mz_min_img = InlineImage(doc, img_paths["mz_min"], width=Mm(160))
            except Exception as e:
                print(f"第 {layer_num} 层图片读取失败: {e}")
                fx_max_img = fx_min_img = my_max_img = my_min_img = mz_max_img = mz_min_img = "图片缺失"
            layer = {
                'layer_num': layer_num,
                'cn_index': cn_index,
                'xc_type': xc_type,
                'zc_type': zc_type,
                'xc': xc,
                'zc': zc,
                'zh': zh,
                "Fx_max_pic": fx_max_img,
                "Fx_min_pic": fx_min_img,
                "My_max_pic": my_max_img,
                "My_min_pic": my_min_img,
                "Mz_max_pic": mz_max_img,
                "Mz_min_pic": mz_min_img
            }
            strut_layers.append(layer)
        # 获取表格参数
        for group_names, force_dict, sigma_dict, len_dict in [
            (XC_group_namelst, XC_Force_dict, XC_sigma_dict, XC_L_dict),
            (DC_group_namelst, DC_Force_dict, DC_sigma_dict, DC_L_dict)
        ]:
            for key in group_names:
                layer_match = re.search(r'第(\d+)层', key)
                layer_num = int(layer_match.group(1)) if layer_match else None
                if '对撑' in key:
                    Allowable = strut_dc_allowable_f_dict.get(layer_num, {}).get('f')
                elif '斜撑' in key:
                    Allowable = strut_xc_allowable_f_dict.get(layer_num, {}).get('f')
                else:
                    Allowable = 215.0
                if Allowable is None:
                    Allowable = 215.0
                sigma_lst = sigma_dict[key]
                # 解构内力值
                f_max, m_y, m_z = force_dict[key]
                length = len_dict[key]
                # 基础数据行
                row = {
                    "name": key,      # 支撑名称 
                    "Fx": f_max,      # 轴力 
                    "Mx": m_y,        # 模板中的 Mx 对应模型中的 My 
                    "My": m_z,        # 模板中的 My 对应模型中的 Mz 
                    "l0": length,     # 计算长度 
                    "limit": Allowable, # 容许应力
                }
                # 根据应力列表长度判断截面类型 
                if len(sigma_lst) == 2:            # 型钢（双向压弯）
                    sigmax, sigmay = sigma_lst
                    row["in_stability"] = sigmax   # 平面内应力 
                    row["out_stability"] = sigmay  # 平面外应力 
                    row["if_satisfied"] = '满足' if sigmax <= abs(Allowable) and sigmay <= Allowable else '不满足' 
                    profile_stability.append(row)
                elif len(sigma_lst) == 1:          # 钢管（圆管）
                    sigma = sigma_lst[0]
                    row["ylz"] = sigma             # 综合应力值 
                    row["if_satisfied"] = '满足' if sigma <= abs(Allowable) else '不满足' 
                    pile_stability.append(row)
                # 全局满足判定
                if row["if_satisfied"] == "不满足":
                    if_satisfied = "不满足"
        # 收集钢管支撑的最不利信息
        for row in pile_stability:
            row_allowable = row['limit']
            daxiao_ylz = "≤" if row['ylz'] <= abs(row_allowable) else ">"
            Strut_res.append(
                f"{row['name']}最不利工况下应力值为{row['ylz']:.2f}MPa{daxiao_ylz}{row_allowable:.2f}MPa，{row['if_satisfied']}要求；"
            )
        # 收集型钢支撑的最不利信息
        for row in profile_stability:
            row_allowable = row['limit']
            daxiao_in = "≤" if row['in_stability'] <= abs(row_allowable) else ">"
            daxiao_out = "≤" if row['out_stability'] <= abs(row_allowable) else ">"
            # 取平面内和平面外应力的较大值作为描述参考
            Strut_res.append(
                f"{row['name']}最不利工况下平面内稳定性应力值为{row['in_stability']:.2f}MPa{daxiao_in}{row_allowable:.2f}MPa，平面内稳定性应力值为{row['out_stability']:.2f}MPa{daxiao_out}{row_allowable:.2f}MPa，{row['if_satisfied']}要求；"
            )
    return has_pile, has_profile, strut_layers, pile_stability, profile_stability, Strut_res, if_satisfied, DC_sigma_dict, XC_sigma_dict


def Model_cal_res(
        version, doc_path, save_path, pic_path, 
        has_stage, load_cases, has_waler, has_strut, Strut_section_dict,
        allowable_f_dict, CofferDam_L, pile_leixing, Waler_SECT_lst, WL_group_namelst, 
        XC_group_namelst, DC_group_namelst, DC_Section_dict, XC_Section_dict, DC_L_dict, XC_L_dict, 
        gamma_0, SECtype, pbar=None, svar=None, root=None
        ):
    # 收集路径
    stage_template_path = os.path.join(doc_path, 'Steel_Sheet_Pile_CofferDam_cal_template_model_cal_stage.docx')
    template_path = os.path.join(doc_path, 'Steel_Sheet_Pile_CofferDam_cal_template_model_cal.docx')
    save_path = os.path.join(save_path, 'model_cal.docx')
    if svar:
        svar.set(f"正在进行模型结果验算\n正在验算{SECtype}...")
    # 加载模板
    if has_stage:
        doc = DocxTemplate(stage_template_path)
        # 获取计算结果
        pile_data, pile_res, pile_if_satisfied, pile_max_cb_all, pile_max_d_all = Pile_cal_Stage(doc, version, pic_path, load_cases, pile_leixing, allowable_f_dict, gamma_0, SECtype, CofferDam_L)
        # 围檩验算
        WL_layers_data, Waler_res, wl_summary = Waler_cal_Stage(doc, version, pic_path, load_cases, has_waler, WL_group_namelst, Waler_SECT_lst, allowable_f_dict, gamma_0, pbar=pbar, svar=svar, root=root,step_weight=8)
        # 内支撑验算
        has_pile, has_profile, strut_layers, pile_stability, profile_stability, Strut_res, Strut_if_satisfied, DC_sigma_dict, XC_sigma_dict = Strut_cal_Stage(doc, version, pic_path, has_strut, Strut_section_dict, XC_group_namelst, DC_group_namelst, DC_Section_dict, XC_Section_dict, DC_L_dict, XC_L_dict, allowable_f_dict, gamma_0, pbar=pbar, svar=svar, root=root,step_weight=8)
    else:
        doc = DocxTemplate(template_path)
        # 工况处理
        load_cases = [load_cases[-1]] # 取开挖至基坑底工况
        print(f"当前为整体建模模式，仅验算最终工况：{load_cases}")
        # 获取计算结果
        pile_data, pile_res, pile_if_satisfied, pile_max_cb_all, pile_max_d_all = Pile_cal(doc, pic_path, load_cases, pile_leixing, allowable_f_dict, gamma_0, SECtype, CofferDam_L)
        # 围檩验算
        WL_layers_data, Waler_res, wl_summary = Waler_cal(doc, pic_path, load_cases, has_waler, WL_group_namelst, Waler_SECT_lst, allowable_f_dict, gamma_0, pbar=pbar, svar=svar, root=root,step_weight=8)
        # 内支撑验算
        has_pile, has_profile, strut_layers, pile_stability, profile_stability, Strut_res, Strut_if_satisfied, DC_sigma_dict, XC_sigma_dict = Strut_cal(doc, pic_path, has_strut, Strut_section_dict, XC_group_namelst, DC_group_namelst, DC_Section_dict, XC_Section_dict, DC_L_dict, XC_L_dict, allowable_f_dict, gamma_0, pbar=pbar, svar=svar, root=root,step_weight=8)
    # 渲染模板
    context = {"pile": pile_data,
               "has_waler": has_waler,
               "waler_layers": WL_layers_data,
               "has_strut": has_strut,
               "has_pile": has_pile,
               "has_profile": has_profile,
               "strut_layers": strut_layers,
               "pile_stability": pile_stability,
               "profile_stability": profile_stability,
            }
    doc.render(context)
    # 保存文件
    doc.save(save_path)
    print(f"模型计算部分计算书已成功生成到{save_path}")
    # 结论处理
    model_cal_satisfied = "满足要求" if pile_if_satisfied == "满足" and Strut_if_satisfied == "满足" else "不满足要求"
    batch_results = {
        'pile_max_cb': pile_max_cb_all,
        'pile_max_d': pile_max_d_all,
        'wl_summary': wl_summary,
        'DC_sigma_dict': DC_sigma_dict,
        'XC_sigma_dict': XC_sigma_dict,
    }
    if pbar and root:
        pbar['value'] += 4
        root.update()
    return pile_res, Waler_res, Strut_res, model_cal_satisfied, batch_results


def Ke_split_calc(Active_Hlst,Solid_Pak,Water_Pak,Waler_Strut_Z,Condition_Passive_Earth_Pressure_Hlst,Condition_Solid_Ppk_lst,Condition_Water_Ppk_lst):
    """ 计算分算抗倾覆稳定性 """
    Ke_split_lst = []
    # 主动土压力
    E_ak_s,a_a_s = calculate_pressure_resultant(Active_Hlst,Solid_Pak,Waler_Strut_Z)
    # 外侧水压力
    E_ak_w,a_a_w = calculate_pressure_resultant(Active_Hlst,Water_Pak,Waler_Strut_Z)
    min_ke = float('inf')
    E_pk_req_s = 0.0
    a_pk_req_s = 0.0
    E_pk_req_w = 0.0
    a_pk_req_w = 0.0
    for i in range(len(Condition_Passive_Earth_Pressure_Hlst)):
        # 提取当前工况的数据
        current_passive_H = Condition_Passive_Earth_Pressure_Hlst[i]
        current_solid_P = Condition_Solid_Ppk_lst[i]
        current_water_P = Condition_Water_Ppk_lst[i]
        # 被动土压力
        E_pk_s, a_p_s = calculate_pressure_resultant(
            current_passive_H, 
            current_solid_P, 
            Waler_Strut_Z
        )
        # 内侧水压力
        E_pk_w, a_p_w = calculate_pressure_resultant(
            current_passive_H, 
            current_water_P, 
            Waler_Strut_Z
        )
        # 抗倾覆稳定性
        Ke_split= (E_pk_s*a_p_s+E_pk_w*a_p_w)/(E_ak_s*a_a_s+E_ak_w*a_a_w)
        Ke_split_lst.append(round(Ke_split, 3))
        # 判断并记录最小值对应的参数
        if Ke_split < min_ke:
            min_ke = Ke_split
            E_pk_req_s= E_pk_s
            a_pk_req_s = a_p_s
            E_pk_req_w= E_pk_w
            a_pk_req_w = a_p_w
    ke_res = min_ke if min_ke != float('inf') else 0.0
    return round(ke_res,3),round(E_ak_s, 2),round(a_a_s, 2),round(E_ak_w, 2),round(a_a_w, 2),round(E_pk_req_s, 2),round(a_pk_req_s, 2),round(E_pk_req_w, 2),round(a_pk_req_w, 2)


def Ke_joint_calc(Active_Hlst,Solid_Pak,Water_Pak,Waler_Strut_Z,Condition_Passive_Earth_Pressure_Hlst,Condition_Solid_Ppk_lst,Condition_Water_Ppk_lst):
    """ 计算合算抗倾覆稳定性 """
    Ke_joint_lst = []
    # 主动土压力
    total_Pak = [
        [s[0] + w[0], s[1] + w[1]] 
        for s, w in zip(Solid_Pak, Water_Pak)
    ]
    E_ak,a_a = calculate_pressure_resultant(Active_Hlst,total_Pak,Waler_Strut_Z)
    # 初始化记录变量
    min_ke = float('inf')
    E_pk_req = 0.0
    a_pk_req = 0.0
    for i in range(len(Condition_Passive_Earth_Pressure_Hlst)):
        # 提取当前工况的数据
        current_passive_H = Condition_Passive_Earth_Pressure_Hlst[i]
        current_solid_P = Condition_Solid_Ppk_lst[i]
        current_water_P = Condition_Water_Ppk_lst[i]
        current_total_P = [
            [s[0] + w[0], s[1] + w[1]] 
            for s, w in zip(current_solid_P, current_water_P)
        ]
        # 被动土压力
        E_pk, a_p = calculate_pressure_resultant(
            current_passive_H, 
            current_total_P, 
            Waler_Strut_Z
        )
        # 抗倾覆稳定性
        Ke_joint = (E_pk*a_p)/(E_ak*a_a)
        Ke_joint_lst.append(round(Ke_joint, 3))
        if Ke_joint < min_ke:
            min_ke = Ke_joint
            E_pk_req= E_pk
            a_pk_req = a_p
    ke_res = min_ke if min_ke != float('inf') else 0.0
    return round(ke_res,3),round(E_ak, 2),round(a_a, 2),round(E_pk_req, 2),round(a_pk_req, 2)  

# ============================================================
# 入口（test 函数）
# ============================================================
def Slice_Method_for_Circular_Slip_Surface(
    stability_result, pit_bottom_elev, ground_elev, strut_bottom_elev, water_elev, surcharge, level, slice_width, solid_layer_lst, Ld, GAMMA_W = 10
):
    # print(pit_bottom_elev)
    # print(ground_elev)
    # print(strut_bottom_elev)
    # print(water_elev)
    # print(surcharge)
    # print(level)
    # print(slice_width)
    # print(solid_layer_lst)
    # print(Ld)
    # print(GAMMA_W)
    
    # ── 差分进化搜索 ────────────────────────────────────────
    best = search_by_de(
        ground_elev       = ground_elev,
        pit_bottom_elev   = pit_bottom_elev,
        strut_bottom_elev = strut_bottom_elev,
        water_elev        = water_elev,
        surcharge         = surcharge,
        level             = level,
        slice_width       = slice_width,
        solid_layer_lst   = solid_layer_lst,
        Ld                = Ld,
        gamma_w           = GAMMA_W,
    )
    print_report_de(best,
                    level             = level,
                    strut_bottom_elev = strut_bottom_elev,
                    pit_bottom_elev   = pit_bottom_elev)

    # ── 详细计算，获取公式分子分母 ──────────────────────────
    Ks, sum_T_all, sum_Ws_all = Cal_Ks_detail(
        best.cx, best.cy, best.R,
        level, slice_width,
        ground_elev, water_elev, pit_bottom_elev,
        solid_layer_lst, surcharge, GAMMA_W,
    )

    # ── 构造 LaTeX 公式字符串 ───────────────────────────────
    require_Ks_dict = {'一级': 1.35, '二级': 1.30, '三级': 1.25}
    require_Ks = require_Ks_dict[level]

    if Ks >= require_Ks:
        omath_str = (r'K_{s,i}=\frac{'
                     + str(sum_T_all)
                     + r'}{'
                     + str(sum_Ws_all)
                     + r'}='
                     + str(Ks)
                     + r'\geq'
                     + str(require_Ks))
        if_satisfied_str = '满足要求'
        daxiao_res = '≥'
    else:
        omath_str = (r'K_{s,i}=\frac{'
                     + str(sum_T_all)
                     + r'}{'
                     + str(sum_Ws_all)
                     + r'}='
                     + str(Ks)
                     + r'<'
                     + str(require_Ks))
        if_satisfied_str = '不满足要求'
        daxiao_res = '＜'
    # 结论处理
    overall_stability_res = f'圆弧稳定性系数计算结果为{round(Ks,2)}{daxiao_res}{require_Ks}, {if_satisfied_str}；'
    stability_result['圆弧滑动稳定性'] = [round(Ks,2), require_Ks]
    return Ks, omath_str, if_satisfied_str, overall_stability_res


def anti_overturn_stability_cal(stability_result, Method_for_Pressures_var, Active_Hlst, Solid_Pak_lst, Water_Pak_lst, Condition_Passive_Earth_Pressure_Hlst, Condition_Solid_Ppk_lst, Condition_Water_Ppk_lst, Waler_Strut_Zlst, CofferDam_Bottom_Level, count, Ld, h, safety_level):
    if count == 0:
        standard = "根据《建筑基坑支护技术规程》JGJ120-2012中4.2.1条文规定：对于悬臂支挡式结构，嵌固稳定性应满足下式："
        cal_level = CofferDam_Bottom_Level
        qiangu = round(0.8*h,2)
        if Ld < qiangu:
            daxiao2 = r"\le"
            is_satisfied2 = "不满足要求"
        else:
            daxiao2 = r"\geq"
            is_satisfied2 = "满足要求"
        qiangu_stability = str(round(Ld,3)) + 'm' + str(daxiao2) + r'0.8\times' + str(h) + '=' + str(qiangu)
    elif count == 1:
        standard = "根据《建筑基坑支护技术规程》JGJ120-2012中4.2.2条文规定：对于单层锚杆和单层支撑的支挡式结构，嵌固稳定性应满足下式："
        cal_level = Waler_Strut_Zlst[-1]
        qiangu = round(0.3*h,2)
        if Ld < qiangu:
            daxiao2 = r"\le"
            is_satisfied2 = "不满足要求"
        else:
            daxiao2 = r"\geq"
            is_satisfied2 = "满足要求"
        qiangu_stability = str(round(Ld,3)) + 'm' + str(daxiao2) + r'0.3\times' + str(h) + '=' + str(qiangu)
    else:
        standard = "根据《建筑地基基础设计规范》GB50007-2011中V.0.1条文规定：对于多层锚杆和多层支撑的支挡式结构，嵌固稳定性应满足下式："
        cal_level = Waler_Strut_Zlst[-1]
        qiangu = round(0.2*h,2)
        if Ld < qiangu:
            daxiao2 = r"\le"
            is_satisfied2 = "不满足要求"
        else:
            daxiao2 = r"\geq"
            is_satisfied2 = "满足要求"
        qiangu_stability = str(round(Ld,3)) + 'm' + str(daxiao2) + r'0.2\times' + str(h) + '=' + str(qiangu)

    # 确定抗倾覆安全系数 Ke
    if safety_level == '一级':
        ke_req = 1.25
    elif safety_level == '二级':
        ke_req = 1.2
    elif safety_level == '三级':
        ke_req = 1.15
    else:
        raise ValueError("安全等级必须为一级/二级/三级")
    
    # 嵌固稳定性公式: 悬臂用4.2.1(a1), 单/多支点用4.2.2(a2)
    std_formula = JGJ_120_2012_4_2_1 if count == 0 else JGJ_120_2012_4_2_2

    if Method_for_Pressures_var ==  "水土分算":
        # 计算水土压力、距支点距离、分算抗倾覆稳定性
        Ke_split,E_ak_s,a_a_s,E_ak_w,a_a_w,E_pk_req_s,a_pk_req_s,E_pk_req_w,a_pk_req_w = Ke_split_calc(Active_Hlst,Solid_Pak_lst,Water_Pak_lst,cal_level,Condition_Passive_Earth_Pressure_Hlst,Condition_Solid_Ppk_lst,Condition_Water_Ppk_lst)
        # 判断抗倾覆稳定性
        if Ke_split < ke_req:
            is_satisfied1 = "不满足要求"
            daxiao1 = r"\le"
            daxiao_res = '＜'
        else:
            is_satisfied1 = "满足要求"
            daxiao1 = r"\geq"
            daxiao_res = '≥'
        anti_overturn_stability = std_formula  + r'\frac{' + str(E_pk_req_s) +r'\times'+ str(a_pk_req_s) + '+' + str(E_pk_req_w) +r'\times'+ str(a_pk_req_w) + '}{' + str(E_ak_s) +r'\times'+ str(a_a_s) + '+' + str(E_ak_w) +r'\times'+ str(a_a_w) + '}' + '=' + str(Ke_split) + str(daxiao1) + str(ke_req)
        ke = Ke_split
        print(anti_overturn_stability)
    else:
        # 计算水土压力、距支点距离、合算抗倾覆稳定性
        Ke_joint,E_ak,a_a,E_pk_req,a_pk_req = Ke_joint_calc(Active_Hlst,Solid_Pak_lst,Water_Pak_lst,cal_level,Condition_Passive_Earth_Pressure_Hlst,Condition_Solid_Ppk_lst,Condition_Water_Ppk_lst)
        # 判断抗倾覆稳定性
        if Ke_joint < ke_req:
            is_satisfied1 = "不满足要求"
            daxiao1 = r"\le"
            daxiao_res = '＜'
        else:
            is_satisfied1 = "满足要求"
            daxiao1 = r"\geq"
            daxiao_res = '≥'
        anti_overturn_stability = std_formula + r'\frac{' + str(E_pk_req) +r'\times'+ str(a_pk_req) + '}{' + str(E_ak) +r'\times'+ str(a_a) + '}' + '=' + str(Ke_joint) + str(daxiao1) + str(ke_req)
        ke = Ke_joint
        print(anti_overturn_stability)
    # 嵌固构造验算
    print(qiangu_stability)
    if is_satisfied1 == "满足要求" and is_satisfied2 == "满足要求":
        is_satisfied = "满足要求"
    else:
        is_satisfied = "不满足要求"
    print(f"是否满足: {is_satisfied}")
    # 结论处理
    anti_overturn_stability_res = f'钢板桩围堰抗倾覆稳定性系数在最不利工况下为{round(ke,2)}{daxiao_res}{ke_req}，嵌固构造{is_satisfied2}，抗倾覆稳定性{is_satisfied}；'
    stability_result['抗倾覆'] = [round(ke,2), ke_req]
    return anti_overturn_stability, qiangu_stability, is_satisfied, standard, anti_overturn_stability_res


def heave_resistant_stability_cal(stability_result, Excel_Data, Ld, h, count, q0, safety_level,
                                  Solid_Top_Level=None, Water_Top_Level=None,
                                  Concrete_Bottom_Level=None, CofferDam_Bottom_Level=None,
                                  include_water_column=True):
    """
    根据 JGJ 120-2012 第4.2.4条验算基坑抗隆起稳定性。
    公式: (γm2 * ld * Nq + c * Nc) / (γm1 * (h + ld) + q0) >= Kb

    深度基准: gamma_cal 的 depth=0 为第一个土层顶面(地面 Solid_Top_Level)，
    因此 h、γm1、γm2 均须以地面起算，不能直接用水面起算的基坑深度。
    参数:
    Ld      (float): 挡土构件嵌固深度 (m)
    h       (float): 基坑深度 (m)，缺省回退值；给定标高时以 Solid_Top_Level-Concrete_Bottom_Level 为准
    q0      (float): 坑外地面均布荷载 (kPa)
    Solid_Top_Level / Water_Top_Level / Concrete_Bottom_Level / CofferDam_Bottom_Level:
                    用于把深度基准统一到地面，并把水面高于地面的水柱计入荷载
    include_water_column (bool): 是否把土层顶以上水柱按 γw*hw 并入 q0
    返回:
    (standard, cal_flag, formula, is_satisfied, res_str)
    """
    if count == 0:
        standard = "根据《建筑基坑支护技术规程》JGJ-120-2012中4.2.4条文规定，悬臂式支挡结构可不进行隆起稳定性验算。"
        heave_resistant_stability_res = "悬臂式支挡结构可不进行隆起稳定性验算；"
        cal_heave_resistant_stability = False
        heave_resistant_stability = ""
        is_satisfied = "满足要求"
        kb_cal = None
        stability_result['抗隆起'] = ['/', '/']
    else:
        standard = "根据《建筑基坑支护技术规程》JGJ-120-2012中4.2.4条文规定，支撑式支挡结构的嵌固深度应该符合基坑抗隆起稳定性要求。锚拉式支挡结构和支撑式支挡结构的嵌固深度应符合下列规定："
        # 深度基准统一到土层顶面(地面): gamma_cal 的 depth=0 即第一个土层顶面
        if Solid_Top_Level is not None and Concrete_Bottom_Level is not None and CofferDam_Bottom_Level is not None:
            h = round(Solid_Top_Level - Concrete_Bottom_Level, 3)          # 基坑深度(地面起算)
            Ld = round(Concrete_Bottom_Level - CofferDam_Bottom_Level, 3)  # 嵌固深度
            toe_depth = round(Solid_Top_Level - CofferDam_Bottom_Level, 3) # 桩底埋深(地面起算)
        else:
            toe_depth = round(h + Ld, 3)
        # γm1: 地面→桩底; γm2: 坑底→桩底
        gamma_1, _, _ = gamma_cal(Excel_Data, 0, toe_depth)
        gamma_2, _, _ = gamma_cal(Excel_Data, h, toe_depth)
        # c、φ 取桩底以下土层(桩底恰在层界时取下层)
        _, c, phi_deg = gamma_cal(Excel_Data, toe_depth, round(toe_depth + 1e-3, 4))
        # 土层顶以上水柱(水面高于地面部分)按荷载并入 q0
        water_above = 0.0
        if include_water_column and Water_Top_Level is not None and Solid_Top_Level is not None:
            water_above = max(0.0, round(Water_Top_Level - Solid_Top_Level, 3))
        q0_eff = round(q0 + 10 * water_above, 2)
        # 确定抗隆起安全系数 Kb
        if safety_level == '一级':
            kb_req = 1.8
        elif safety_level == '二级':
            kb_req = 1.6
        elif safety_level == '三级':
            kb_req = 1.4
        else:
            raise ValueError("安全等级必须为一级/二级/三级")
        # 将角度转换为弧度
        phi_rad = math.radians(phi_deg)
        # 计算承载力系数 Nq 和 Nc
        # Nq = tan^2(45° + φ/2) * e^(π * tanφ)
        # 公式中 45°+φ/2 需要转为弧度计算
        term_angle = math.radians(45) + (phi_rad / 2.0)
        tan_term = math.tan(term_angle)
        nq = (tan_term ** 2) * math.exp(math.pi * math.tan(phi_rad))
        # Nc = (Nq - 1) / tanφ
        if phi_deg == 0:
            raise ValueError("内摩擦角φ不能为0，需重新取值。")
        # 计算公式
        nc = (nq - 1.0) / math.tan(phi_rad)
        numerator = (gamma_2 * Ld * nq) + (c * nc)
        denominator = (gamma_1 * (h + Ld)) + q0_eff
        if denominator == 0:
            raise ValueError("分母为0，请检查输入参数 (h, ld, q0, gamma_m1)。")
        # 计算实际安全系数 K_calc
        kb_cal = numerator / denominator 
        # 判断是否满足要求
        is_pass = True if kb_cal >= kb_req else False
        is_satisfied = "满足要求" if is_pass else "不满足要求"
        # daxiao = "\geq" if is_pass else "\le"
        daxiao = r"\geq" if is_pass else r"\le"
        daxiao_res = '≥' if is_pass else '＜'
        # q0 项显示: 有水柱时展开为 q0+γw·hw
        if include_water_column and water_above > 0:
            q0_str = str(round(q0, 2)) + r'+' + str(round(10 * water_above, 2))
        else:
            q0_str = str(round(q0, 2))
        heave_resistant_stability = JGJ_120_2012_4_2_4 + r'=\frac{' + str(gamma_2) + r'\times' + str(round(Ld,2)) + r'\times' + str(round(nq,2)) + '+' + str(c) + r'\times' + str(round(nc,2)) +'}{' + str(gamma_1) + r'\times\left(' + str(round(h,2)) + '+' + str(round(Ld,2)) + r'\right)' + '+' + q0_str + '}' +'=' + str(round(kb_cal,2)) + str(daxiao) + str(kb_req)
        print(heave_resistant_stability)
        heave_resistant_stability_res = f'基坑抗隆起稳定性系数在最不利工况下为{round(kb_cal,2)}{daxiao_res}{kb_req}，{is_satisfied}；'
        cal_heave_resistant_stability = True
        stability_result['抗隆起'] = [round(kb_cal,2), kb_req]
    return standard, cal_heave_resistant_stability, heave_resistant_stability, is_satisfied, heave_resistant_stability_res


def seepage_stability_cal(seepage_vars, Solid_Top_Level, Water_Top_Level, Concrete_Bottom_Level, Ld, Excel_Data, if_confined, d, safety_level):
    """ 基坑渗透稳定性验算 — 返回 (seepage_items, context_add, seepage_res)
    gamma_cal / gamma_sub_cal 的深度均自第一个土层顶面(地面 Solid_Top_Level)起算，
    故需把标高换算为深度后再传入。"""
    seepage_items = []
    context_add = {}
    seepage_res = []
    # 地面起算的深度
    excav_depth = round(Solid_Top_Level - Concrete_Bottom_Level, 3)   # 基坑底埋深
    water_depth = round(Solid_Top_Level - Water_Top_Level, 3)         # 水位埋深(可为负)
    for seepage_var in seepage_vars:
        if seepage_var == "抗突涌稳定性验算":
            gamma, _, _ = gamma_cal(Excel_Data, excav_depth, excav_depth + d)
            delta_h = Water_Top_Level - Concrete_Bottom_Level
            hw = delta_h + d
            Kh = d*gamma / hw / 10
            if Kh >= 1.1:
                is_satisfied = "满足要求"
                daxiao = r"\geq"
                daxiao_res = '≥'
            else:
                is_satisfied = "不满足要求"
                daxiao = r"\le"
                daxiao_res = '＜'
            formula_str = (r'K_{抗突涌}=\frac{' + str(round(d, 2))
                           + r'\times' + str(round(gamma, 2))
                           + r'}{' + str(round(hw, 2)) + r'\times10}='
                           + str(round(Kh, 2)) + daxiao + r'1.1')
            seepage_items.append(SeepageItem("抗突涌稳定性验算",
                                             water_heave_resistant_stability=formula_str))
            context_add['water_heave_resistant_stability_satisfied'] = is_satisfied
            seepage_res.append(f'抗突涌稳定性系数计算结果为{round(Kh,2)}{daxiao_res}1.1，{is_satisfied}；')
        elif seepage_var == "流土稳定性验算":
            d1 = d if if_confined else Water_Top_Level - Concrete_Bottom_Level
            delta_h = Water_Top_Level - Concrete_Bottom_Level
            gamma = gamma_sub_cal(Excel_Data, water_depth, excav_depth, water_depth)
            Kf = (2 * Ld + 0.8 * d1) * gamma / delta_h / 10
            req_map = {'一级': 1.6, '二级': 1.5, '三级': 1.4}
            Kf_req = req_map.get(safety_level, 1.5)
            if Kf >= Kf_req:
                is_satisfied = "满足要求"
                daxiao = r"\geq"
                daxiao_res = '≥'
            else:
                is_satisfied = "不满足要求"
                daxiao = r"\le"
                daxiao_res = '＜'
            formula_str = (r'K_{流土}=\frac{\left(2\times' + str(round(Ld, 2))
                           + r'+0.8\times' + str(round(d1, 2))
                           + r'\right)\times' + str(round(gamma, 2))
                           + r'}{' + str(round(delta_h, 2)) + r'\times10}='
                           + str(round(Kf, 2)) + daxiao + str(Kf_req))
            print(f'')
            seepage_items.append(SeepageItem("流土稳定性验算",
                                             water_flow_resistant_stability=formula_str))
            context_add['water_flow_resistant_stability_satisfied'] = is_satisfied
            seepage_res.append(f'流土稳定性系数计算结果为{round(Kf,2)}{daxiao_res}{Kf_req}，{is_satisfied}；')
    return seepage_items, context_add, seepage_res


def stability_cal(doc_path, save_path, Excel_Data, Solid_Top_Level, Water_Top_Level, Method_for_Pressures_var, Active_Hlst, Solid_Pak_lst, Water_Pak_lst, Condition_Passive_Earth_Pressure_Hlst, Condition_Solid_Ppk_lst, Condition_Water_Ppk_lst, Waler_Strut_Zlst, Concrete_Bottom_Level, CofferDam_Bottom_Level, level, slice_width, solid_layer_lst, Ld, h, Ea_p0, count, seepage_vars=[], if_confined=False, d_confined=None, cal_config=None):
    stability_result = {}
    # 计算配置：未勾选的稳定性项跳过计算, 结果置为 ['/', '/'] 且不影响满足性结论
    cfg = cal_config if cal_config else {}
    overall_selected = cfg.get("滑动稳定性分析", True)
    anti_overturn_selected = cfg.get("嵌固稳定性分析", True)
    heave_resistant_selected = cfg.get("抗隆起稳定性分析", True)
    # 抗倾覆参数计算
    if anti_overturn_selected:
        anti_overturn_stability, qiangu_stability, anti_overturn_stability_satisfied, anti_overturn_standard, anti_overturn_stability_res = anti_overturn_stability_cal(stability_result, Method_for_Pressures_var, Active_Hlst, Solid_Pak_lst, Water_Pak_lst, Condition_Passive_Earth_Pressure_Hlst, Condition_Solid_Ppk_lst, Condition_Water_Ppk_lst, Waler_Strut_Zlst, CofferDam_Bottom_Level, count, Ld, h, level)
    else:
        stability_result['抗倾覆'] = ['/', '/']
        anti_overturn_stability, qiangu_stability = '/', '/'
        anti_overturn_stability_satisfied, anti_overturn_standard = "满足要求", '/'
        anti_overturn_stability_res = []
        print("嵌固稳定性分析(抗倾覆)未勾选, 已跳过计算")
    # 抗隆起参数计算
    if heave_resistant_selected:
        heave_resistant_standard, cal_heave_resistant_stability, heave_resistant_stability, heave_resistant_stability_satisfied, heave_resistant_stability_res = heave_resistant_stability_cal(
            stability_result, Excel_Data, Ld, h, count, Ea_p0, level,
            Solid_Top_Level=Solid_Top_Level, Water_Top_Level=Water_Top_Level,
            Concrete_Bottom_Level=Concrete_Bottom_Level, CofferDam_Bottom_Level=CofferDam_Bottom_Level,
            include_water_column=cfg.get("抗隆起计入水柱", True))
    else:
        stability_result['抗隆起'] = ['/', '/']
        heave_resistant_standard, cal_heave_resistant_stability, heave_resistant_stability = '/', '/', '/'
        heave_resistant_stability_satisfied = "满足要求"
        heave_resistant_stability_res = []
        print("抗隆起稳定性分析未勾选, 已跳过计算")
    # 圆弧稳定性参数计算
    if overall_selected:
        Ks, overall_stability, overall_stability_satisfied, overall_stability_res = Slice_Method_for_Circular_Slip_Surface(stability_result, Concrete_Bottom_Level, Solid_Top_Level, CofferDam_Bottom_Level, Water_Top_Level, Ea_p0, level, slice_width, solid_layer_lst, Ld)
    else:
        stability_result['圆弧滑动稳定性'] = ['/', '/']
        Ks, overall_stability = '/', '/'
        overall_stability_satisfied = "满足要求"
        overall_stability_res = []
        print("滑动稳定性分析(圆弧滑动)未勾选, 已跳过计算")
    # 渗透稳定性验算
    cal_seepage_stability = True if len(seepage_vars) > 0 else False
    print(f'是否进行渗透稳定性验算：{cal_seepage_stability}')
    seepage_items = []
    seepage_ctx = {}
    seepage_stability_res = []
    if cal_seepage_stability:
        seepage_items, seepage_ctx, seepage_stability_res = seepage_stability_cal(
            seepage_vars, Solid_Top_Level, Water_Top_Level, Concrete_Bottom_Level, Ld, Excel_Data,
            if_confined, d_confined, level)

    stability_doc = DocxTemplate(os.path.join(doc_path, 'Steel_Sheet_Pile_CofferDam_cal_template_stability.docx'))
    save_path = os.path.join(save_path, 'Stability_cal.docx')
    context = {
        'overall_stability_satisfied':overall_stability_satisfied,
        'anti_overturn_standard':anti_overturn_standard,
        'anti_overturn_stability_satisfied':anti_overturn_stability_satisfied,
        'heave_resistant_standard':heave_resistant_standard,
        'cal_heave_resistant_stability':cal_heave_resistant_stability,
        'heave_resistant_stability_satisfied':heave_resistant_stability_satisfied,
        'cal_seepage_stability':cal_seepage_stability,
        'seepage_vars': seepage_items,
        'overall_selected': overall_selected,
        'anti_overturn_selected': anti_overturn_selected,
        'heave_resistant_selected': heave_resistant_selected,
        **seepage_ctx,
    }
    stability_doc.render(context)
    stability_doc.save(save_path)
    # 替换公式
    final_path = save_path.replace('.docx', '_final.docx')
    finaLdoc = Document(save_path)
    replace_placeholder_with_formula(finaLdoc, '{overall_stability}', overall_stability)
    replace_placeholder_with_formula(finaLdoc, '{anti_overturn_stability}', anti_overturn_stability)
    replace_placeholder_with_formula(finaLdoc, '{qiangu_stability}', qiangu_stability)
    replace_placeholder_with_formula(finaLdoc, '{heave_resistant_stability}', heave_resistant_stability)
    # 渗透稳定性公式 (仅当对应验算项存在时, 模板循环已生成其段落)
    for item in seepage_items:
        if item.name == "抗突涌稳定性验算":
            replace_placeholder_with_formula(finaLdoc, r'{seepage_var.water_heave_resistant_stability}',
                                             item.water_heave_resistant_stability)
        elif item.name == "流土稳定性验算":
            replace_placeholder_with_formula(finaLdoc, r'{water_flow_resistant_stability}',
                                             item.water_flow_resistant_stability)
    finaLdoc.save(final_path)
    # 结论处理（是否满足要求）— 含渗透验算结论
    seepage_satisfied = all("满足要求" in r for r in seepage_stability_res)
    stability_satisfied = "满足要求" if (
        overall_stability_satisfied == "满足要求"
        and anti_overturn_stability_satisfied == "满足要求"
        and heave_resistant_stability_satisfied == "满足要求"
        and seepage_satisfied
    ) else "不满足要求"
    cal_seepage_stability = "{}。".format("".join(seepage_stability_res)) if seepage_stability_res else ""
    return overall_selected, overall_stability_res, anti_overturn_selected, anti_overturn_stability_res, heave_resistant_selected, cal_heave_resistant_stability, heave_resistant_stability_res, stability_satisfied, cal_seepage_stability, seepage_stability_res


class SeepageItem:
    """渗透稳定性验算项目 — 支持 docxtpl 模板中 == 比较和 . 属性访问"""
    def __init__(self, name, **data):
        self.name = name
        self.__dict__.update(data)
    def __eq__(self, other):
        return self.name == other


def transform_soil_parameters(raw_data):
    """
    将原始地质数据列表转换为计算用的格式：[层厚, 重度, 黏聚力, 内摩擦角]
    """
    # raw_data[1:] 用于跳过第一行的表头
    integrated_data = [
        [
            row[2],  # 层厚(m)
            row[3],  # 重度(kN/m3)
            row[4],  # 黏聚力c(kPa)
            row[5]   # 内摩擦角φ(°)
        ] 
        for row in raw_data[1:]
    ]
    return integrated_data


# 读取焊缝连接点两侧单元截面和受力（一次拉取结果）
def Weld_Connection_data(has_stage, weld_nodes_WL_Strut, weld_nodes_WL_WL, sec_db_s, sec_deb_w, SECT_res, Node_res, Elem_res, Structure_group_res, hf, weld_type='full'):
    he = 0.7 * hf
    force_map = {'Fxmax': 2, 'Fymax': 3, 'Fzmax': 4, 'Tmax': 5, 'Mymax': 6, 'Mzmax': 7}
    # 立单元号→结构组映射
    elem_to_group_name = {}
    Structure_group_data = Structure_group_res.get('GRUP', {})
    for g_info in Structure_group_data.values():
        g_name = g_info.get('NAME', 'Unknown')
        for eid in g_info.get('E_LIST', []):
            elem_to_group_name[str(eid)] = g_name

    # 内支撑与围檩连接点处理
    target_strut_eids = list({int(item['strut_elem']) for item in weld_nodes_WL_Strut})
    strut_force_db = {}
    if target_strut_eids:
        res = Value_BeamForce_Elem_Stage("Strutforce", target_strut_eids, "合计(CS)", ["Min/Max:max","Min/Max:min"]) if has_stage else \
              Value_BeamForce_Elem("Strutforce", target_strut_eids, "基本组合(CB)")
        for row in res.get("Strutforce", {}).get('DATA', []):
            eid = str(row[1])
            if eid not in strut_force_db: strut_force_db[eid] = []
            strut_force_db[eid].append(row)
    WL_Strut_data = []
    processed_struts = set() 
    for item in weld_nodes_WL_Strut:
        strut_e = str(item['strut_elem'])
        if strut_e in processed_struts: continue # 支撑单元去重
        nid = item['node_id']
        secname_s = SECT_res['SECT'][str(Elem_res['ELEM'][strut_e]['SECT'])]['SECT_NAME']
        prop = sec_db_s.get(secname_s)
        if not prop: continue
        # 计算焊缝长度
        if prop['type'] == 'PIPE':
            lw = math.pi * prop['D']
        elif prop['type'] == 'TYPE_BEAM':
            lw = (4 * prop['B'] + 2 * (prop['H'] - 2 * prop['tf'])) - (8 * 2 * hf)
        elif prop['type'] == 'DOUBLE_TYPE_BEAM':
            single_lw = (4 * (prop['B']/2) + 2 * (prop['H'] - 2 * prop['tf_up']))
            lw = (single_lw * 2) - (16 * 2 * hf)
        else: lw = 0
        data = strut_force_db.get(strut_e, [])
        if data:
            max_forces = {k: round(float(max(data, key=lambda x: abs(float(x[idx])))[idx]), 2) for k, idx in force_map.items()}
            WL_Strut_data.append({
                'node_id': nid, 'sect_name': secname_s, 'lw': round(lw, 2), 'Ix': round(prop['Ix'],2), 'Wx': round(prop['Wx'],2), 'A': round(prop['A'],2),
                **max_forces, 'group_name': elem_to_group_name.get(strut_e, "未定义结构组")
            })
            processed_struts.add(strut_e)

    # 围檩与围檩连接点处理：选取小截面
    temp_wl_targets = []
    for item in weld_nodes_WL_WL:
        w1_e, w2_e = str(item['wl_elem1']), str(item['wl_elem2'])
        p1 = sec_deb_w.get(SECT_res['SECT'][str(Elem_res['ELEM'][w1_e]['SECT'])]['SECT_NAME'])
        p2 = sec_deb_w.get(SECT_res['SECT'][str(Elem_res['ELEM'][w2_e]['SECT'])]['SECT_NAME'])
        if not p1 or not p2: continue
        # 选取面积小的截面
        if p1['A'] <= p2['A']:
            t_prop, t_name, t_eid = p1, SECT_res['SECT'][str(Elem_res['ELEM'][w1_e]['SECT'])]['SECT_NAME'], w1_e
        else:
            t_prop, t_name, t_eid = p2, SECT_res['SECT'][str(Elem_res['ELEM'][w2_e]['SECT'])]['SECT_NAME'], w2_e
        
        temp_wl_targets.append({'nid': item['node_id'], 'prop': t_prop, 'sname': t_name, 'eid': t_eid})
    # 批量拉取围檩内力
    target_waler_ids = list({int(x['eid']) for x in temp_wl_targets})
    waler_force_db = {}
    if target_waler_ids:
        res = Value_BeamForce_Elem_Stage("Walerforce", target_waler_ids, "合计(CS)", ["Min/Max:max","Min/Max:min"]) if has_stage else \
              Value_BeamForce_Elem("Walerforce", target_waler_ids, "基本组合(CB)")
        for row in res.get("Walerforce", {}).get('DATA', []):
            eid = str(row[1])
            if eid not in waler_force_db: waler_force_db[eid] = []
            waler_force_db[eid].append(row)
    # 整合结果
    WL_WL_data = []
    for wl in temp_wl_targets:
        prop = wl['prop']
        # 焊缝长度计算
        if prop['type'] == 'TYPE_BEAM':
            lw = (4 * prop['B'] + 2 * (prop['H'] - 2 * prop['tf'])) - (8 * 2 * hf)
        elif prop['type'] == 'DOUBLE_TYPE':
            single_lw = (4 * (prop['B']/2) + 2 * (prop['H'] - 2 * prop['tf_up']))
            lw = (single_lw * 2) - (16 * 2 * hf)
        else: lw = 0
        data = waler_force_db.get(wl['eid'], [])
        if data:
            max_forces = {k: round(float(max(data, key=lambda x: abs(float(x[idx])))[idx]), 2) for k, idx in force_map.items()}
            WL_WL_data.append({
                'node_id': wl['nid'], 'sect_name': wl['sname'], 'lw': round(lw, 2), 'Ix': round(prop['Ix'],2), 'Wx': round(prop['Wx'],2), 'A': round(prop['A'],2),
                **max_forces, 'group_name': elem_to_group_name.get(wl['eid'], "未定义结构组"),
                'target_eid': wl['eid']
            })
    print(f"处理完成：支撑点 {len(WL_Strut_data)} 个，围檩点 {len(WL_WL_data)} 个")
    return WL_Strut_data, WL_WL_data


def _allowable_from_section(sec_name, matl_name, Section_dict, Steel_dict, steel_standard):
    """根据截面名称和材质名称计算 (tf, tw, 容许应力f, 容许剪应力fv)。
    截面/材质为 '/' 或 None 时返回 (None, None, None, None)。"""
    if not sec_name or sec_name == '/' or not matl_name or matl_name == '/':
        return None, None, None, None
    sect_info = Section_dict[sec_name]['SEC']
    if '∠' in sec_name:
        tf = float(sect_info[2])
        tw = float(sect_info[2])
    elif '钢管桩' in sec_name:
        tf = float(sect_info[1])
        tw = float(sect_info[1])
    else:
        tf = float(sect_info[3])
        tw = float(sect_info[2])

    matl_info = Steel_dict[steel_standard]['牌号参数'][matl_name]
    f = 215.0
    fv = 125.0
    for k, v in matl_info['f'].items():
        if tf <= float(k):
            f = float(v)
            break
    for k, v in matl_info['fv'].items():
        if tw <= float(k):
            fv = float(v)
            break
    return tf, tw, f, fv


def allowable_f(SECtype, pile_material, ggz_t, steel_standard, Steel_dict, 
                Waler_section_dict, WL_group_namelst, Waler_SECT_lst, Waler_MATL_lst, 
                Strut_section_dict, DC_SECT_lst, XC_SECT_lst, DC_MATL_lst, XC_MATL_lst, 
                ):
    allowable_f_dict = {
        'pile': None,
        'waler': [],
        'strut_dc': {},
        'strut_xc': {},
    }
    
    if SECtype == '钢板桩':
        pile_f = 265.0
        pile_fv = 125.0
    elif SECtype == '锁扣钢管桩':
        pile_f = 215
        pile_fv = 125
        pile_matl_info = Steel_dict[steel_standard]['牌号参数'][pile_material]
        pile_matl_f_dict = pile_matl_info['f']
        pile_matl_fv_dict = pile_matl_info['fv']
        pile_tf = float(ggz_t)
        pile_tw = float(ggz_t)
        for k, v in pile_matl_f_dict.items():
            if pile_tf <= float(k):
                pile_f = float(v)
                break
        for k, v in pile_matl_fv_dict.items():
            if pile_tw <= float(k):
                pile_fv = float(v)
                break
    else:
        pile_f = 215.0
        pile_fv = 125.0

    allowable_f_dict['pile'] = {'f': pile_f, 'fv': pile_fv}

    print('支护桩')
    print(f'    支护桩类型: {SECtype}')
    print(f'    支护桩材质: {pile_material}')
    print(f'    支护桩容许应力: {pile_f}')
    print(f'    支护桩容许剪应力: {pile_fv}')
    print("="*(20))

    num_layers = max(len(Waler_SECT_lst), len(DC_SECT_lst), len(XC_SECT_lst))
    for i in range(num_layers):
        layer_num = i + 1
        print(f'第{layer_num}层支撑信息:')
        # 围檩、对撑、斜撑截面（不足的层按'/'处理）
        waler_sec = Waler_SECT_lst[i] if i < len(Waler_SECT_lst) else '/'
        strut_dc_sec = DC_SECT_lst[i] if i < len(DC_SECT_lst) else '/'
        strut_xc_sec = XC_SECT_lst[i] if i < len(XC_SECT_lst) else '/'
        
        # 围檩、对撑、斜撑材质
        waler_matl = Waler_MATL_lst[i] if i < len(Waler_MATL_lst) else '/'
        strut_dc_matl = DC_MATL_lst[i] if i < len(DC_MATL_lst) else '/'
        strut_xc_matl = XC_MATL_lst[i] if i < len(XC_MATL_lst) else '/'

        # 根据截面和材质信息计算各构件的容许应力
        waler_tf, waler_tw, waler_matl_f, waler_matl_fv = _allowable_from_section(
            waler_sec, waler_matl, Waler_section_dict, Steel_dict, steel_standard)
        dc_tf, dc_tw, dc_matl_f, dc_matl_fv = _allowable_from_section(
            strut_dc_sec, strut_dc_matl, Strut_section_dict, Steel_dict, steel_standard)
        xc_tf, xc_tw, xc_matl_f, xc_matl_fv = _allowable_from_section(
            strut_xc_sec, strut_xc_matl, Strut_section_dict, Steel_dict, steel_standard)

        allowable_f_dict['waler'].append({'f': waler_matl_f, 'fv': waler_matl_fv})
        allowable_f_dict['strut_dc'][layer_num] = {'f': dc_matl_f, 'fv': dc_matl_fv}
        allowable_f_dict['strut_xc'][layer_num] = {'f': xc_matl_f, 'fv': xc_matl_fv}

        print('围檩')
        print(f'    第{layer_num}层[围檩]截面：{waler_sec}')
        print(f'    第{layer_num}层[围檩]截面翼缘厚度：{waler_tf}')
        print(f'    第{layer_num}层[围檩]截面腹板厚度：{waler_tw}')
        print(f'    第{layer_num}层[围檩]材质：{waler_matl}')
        print(f'    第{layer_num}层[围檩]容许应力：{waler_matl_f}')
        print(f'    第{layer_num}层[围檩]容许剪应力：{waler_matl_fv}')
        
        print('对撑')
        print(f'    第{layer_num}层[对撑]截面：{strut_dc_sec}')
        print(f'    第{layer_num}层[对撑]材质：{strut_dc_matl}')
        print(f'    第{layer_num}层[对撑]截面翼缘厚度：{dc_tf}')
        print(f'    第{layer_num}层[对撑]截面腹板厚度：{dc_tw}')
        print(f'    第{layer_num}层[对撑]容许应力：{dc_matl_f}')
        print(f'    第{layer_num}层[对撑]容许剪应力：{dc_matl_fv}')

        print('斜撑')
        print(f'    第{layer_num}层[斜撑]截面：{strut_xc_sec}')
        print(f'    第{layer_num}层[斜撑]材质：{strut_xc_matl}')
        print(f'    第{layer_num}层[斜撑]截面翼缘厚度：{xc_tf}')
        print(f'    第{layer_num}层[斜撑]截面腹板厚度：{xc_tw}')
        print(f'    第{layer_num}层[斜撑]容许应力：{xc_matl_f}')
        print(f'    第{layer_num}层[斜撑]容许剪应力：{xc_matl_fv}')
        print("="*(20))

    return allowable_f_dict