# 1. 标准库
import re

# 3. 本地模块
from General.DataUtils  import num_to_chinese_num
from General.Midas      import API_DIST_FORCE_Unit, Value_BeamStress, Value_Deformed, Value_BeamStress_Stage, Value_Deformed_Stage

from CalRpt_MSWord.Steel_Sheet_Pile_CofferDam_Cal.CofferDam_Cal_Docx_Handler import (
    group_by_layer, Strut_Force_Value, Strut_Force_Value_Stage, Cal_strut_Stability, anti_overturn_stability_cal, heave_resistant_stability_cal, Slice_Method_for_Circular_Slip_Surface
)


def Pile_cal_Stage(version, pile_leixing, allowable_f_dict, gamma_0, SECtype, CofferDam_L):
    if version== "2026":
        minmax = "最小/最大"
    else:
        minmax = "Min/Max"
    pile_jisuan = []
    pile_allowable_f = allowable_f_dict['pile']['f']
    pile_allowable_d = round(float(CofferDam_L)*2.5, 1)
    # 获取截图
    API_DIST_FORCE_Unit("N", 'MM') # 应力为N和MM
    # 确保切换后处理状态，先截图
    # Picture_Beamstress_Stage(pic_path, "gbz_cb_pic_max.jpg", SECtype, minmax, "合计", "Combined", "max", 45, 30)
    # 初始化追踪变量
    max_cb_all_cases = -1.0
    max_d_all_cases = -1.0
    max_cb_case_name = ""
    max_d_case_name = ""
    case_name  = minmax
    case_input = [minmax + ':最大', minmax + ':最小'] if version== "2026" else [minmax + ':max', minmax + ':min']
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
    # # 根据计算结果获取变形截图
    # API_DIST_FORCE_Unit("N", 'MM') # 变形为MM
    # # 施工阶段钢板桩应力截图BUG(激活多余结构组)，再截一次
    # Picture_Beamstress_Stage(pic_path, "gbz_cb_pic_max.jpg", SECtype, minmax, "合计", "Combined", "max", 45, 30)
    # Picture_Beamstress_Stage(pic_path, "gbz_cb_pic_min.jpg", SECtype, minmax, "合计", "Combined", "min", 45, 30)
    # Picture_Deformed_Stage(pic_path, "gbz_d_pic_max.jpg", SECtype, minmax, "合计", pile_d_direction, "max", 45, 30)
    # Picture_Deformed_Stage(pic_path, "gbz_d_pic_min.jpg", SECtype, minmax, "合计", pile_d_direction, "min", 45, 30)
    # # 处理图片
    # cb_max_path = os.path.join(pic_path, "gbz_cb_pic_max.jpg")
    # cb_min_path = os.path.join(pic_path, "gbz_cb_pic_min.jpg")
    # d_max_path = os.path.join(pic_path, "gbz_d_pic_max.jpg")
    # d_min_path = os.path.join(pic_path, "gbz_d_pic_min.jpg")
    # # 读取图片
    # cb_max_img = InlineImage(doc, cb_max_path, width=Mm(160))
    # cb_min_img = InlineImage(doc, cb_min_path, width=Mm(160))
    # d_max_img = InlineImage(doc, d_max_path, width=Mm(160))
    # d_min_img = InlineImage(doc, d_min_path, width=Mm(160))
    pile_data={
        "SECtype": SECtype,
        "pile_jisuan": pile_jisuan,
        # "pile_cb_max_pic": cb_max_img,
        # "pile_cb_min_pic": cb_min_img,
        # "pile_d_max_pic": d_max_img,
        # "pile_d_min_pic": d_min_img
    }
    # 处理结论
    daxiao_cb = "≤" if max_cb_all_cases <= float(pile_allowable_f) else ">"
    daxiao_d = "≤" if max_d_all_cases <= float(pile_allowable_d) else ">"
    pile_res = f"{SECtype}在{max_cb_case_name}下组合应力最大值为{max_cb_all_cases:.2f} MPa{daxiao_cb}{pile_allowable_f}MPa，在{max_d_case_name}下变形最大值为 {max_d_all_cases:.2f}mm{daxiao_d}{pile_allowable_d}mm，{if_satisfied}要求；"
    return pile_data, pile_res, if_satisfied, max_cb_all_cases, max_d_all_cases


def Waler_cal_Stage(version, has_waler, WL_group_namelst, Waler_SECT_lst, allowable_f_dict, gamma_0, pbar=None, svar=None, root=None, step_weight=8):
    if version == "2026":
        minmax = "最小/最大"
    else:
        minmax = "Min/Max"
    WL_layers_data = []
    WL_res = []
    wl_summary = {}
    wl_cb_limit_lst = [v['f'] for v in allowable_f_dict['waler']]
    wl_shear_limit_lst = [v['fv'] for v in allowable_f_dict['waler']]
    waler_allowable_d = ''
    if has_waler:
        num_layers = len(WL_group_namelst)
        increment = step_weight / num_layers if num_layers > 0 else 0
        # 按层读取替换
        for i, WL_group_name in enumerate(WL_group_namelst):
            idx = i+1
            WL_cb_limit = wl_cb_limit_lst[i]
            WL_shear_limit = wl_shear_limit_lst[i]
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
            # 施工阶段建模
            # 判断是否为前工况
            case_name = minmax
            case_input = [minmax + ':最大', minmax + ':最小'] if version== "2026" else [minmax + ':max', minmax + ':min']
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
                "limit_cb": WL_cb_limit,            # 组合应力限值
                "limit_shear": WL_shear_limit,      # 剪应力限值
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

            # # 获取截图
            # API_DIST_FORCE_Unit("N", 'MM') # 应力为N和MM
            # Picture_Beamstress_Stage(pic_path, f"wl{i+1}_cb_max_pic.jpg", [WL_group_namelst[i]],  minmax, "合计", "Combined", "max", 0, 90)
            # Picture_Beamstress_Stage(pic_path, f"wl{i+1}_cb_min_pic.jpg", [WL_group_namelst[i]],  minmax, "合计", "Combined", "min", 0, 90)
            # Picture_Beamstress_Stage(pic_path, f"wl{i+1}_shear_max_pic.jpg", [WL_group_namelst[i]],  minmax, "合计", "Ssz", "max", 0, 90)
            # Picture_Beamstress_Stage(pic_path, f"wl{i+1}_shear_min_pic.jpg", [WL_group_namelst[i]],  minmax, "合计", "Ssz", "min", 0, 90)
            # Picture_Deformed_Stage(pic_path, f"WL{i+1}_d_max_pic.jpg", [WL_group_namelst[i]],  minmax, "合计", WL_d_direction, "max", 0, 90)
            # Picture_Deformed_Stage(pic_path, f"WL{i+1}_d_min_pic.jpg", [WL_group_namelst[i]],  minmax, "合计", WL_d_direction, "min", 0, 90)
            # # 处理图片
            # cb_max_path = os.path.join(pic_path, f"wl{i+1}_cb_max_pic.jpg")
            # cb_min_path = os.path.join(pic_path, f"wl{i+1}_cb_min_pic.jpg")
            # shear_max_path = os.path.join(pic_path, f"wl{i+1}_shear_max_pic.jpg")
            # shear_min_path = os.path.join(pic_path, f"wl{i+1}_shear_min_pic.jpg")
            # d_max_path = os.path.join(pic_path, f"wl{i+1}_d_max_pic.jpg")
            # d_min_path = os.path.join(pic_path, f"wl{i+1}_d_min_pic.jpg")
            # try:
            #     cb_max_img = InlineImage(doc, cb_max_path, width=Mm(160))
            #     cb_min_img = InlineImage(doc, cb_min_path, width=Mm(160))
            #     shear_max_img = InlineImage(doc, shear_max_path, width=Mm(160))
            #     shear_min_img = InlineImage(doc, shear_min_path, width=Mm(160))
            #     d_max_img = InlineImage(doc, d_max_path, width=Mm(160))
            #     d_min_img = InlineImage(doc, d_min_path, width=Mm(160))
            # except Exception as e:
            #     print(f"第 {i+1} 层图片读取失败: {e}")
            #     cb_max_img = cb_min_img = shear_max_img = shear_min_img = d_max_img = d_min_img = "图片缺失"
            # 内容汇总 
            WL_layers_data.append({
                "index": idx,
                "cn_index": num_to_chinese_num(idx),
                "WL_type": Waler_SECT_lst[i],
                "WLjisuan": WLjisuan,
                # "WL_cb_max_pic":cb_max_img,
                # "WL_cb_min_pic":cb_min_img,
                # "WL_shear_max_pic":shear_max_img,
                # "WL_shear_min_pic":shear_min_img,
                # "WL_d_max_pic":d_max_img,
                # "WL_d_min_pic":d_min_img
            })
            # 处理结论
            daxiao_cb = "≤" if layer_max_cb <= float(WL_cb_limit) else ">"
            daxiao_shear = "≤" if layer_max_shear <= float(WL_shear_limit) else ">"
            layer_summary = f"第{idx}层围檩在{layer_max_cb_case}下组合应力最大值为{layer_max_cb:.2f}MPa{daxiao_cb}{WL_cb_limit}MPa，在{layer_max_shear_case}下剪应力最大值为{layer_max_shear:.2f}MPa{daxiao_shear}{WL_shear_limit}MPa；"
            WL_res.append(layer_summary)
            wl_summary[idx] = {'max_cb': layer_max_cb, 'max_shear': layer_max_shear, 'max_d': layer_max_d}
    return WL_layers_data, WL_res, wl_summary


def Strut_cal_Stage(version, has_strut, Strut_section_dict, XC_group_namelst,DC_group_namelst, DC_Section_dict, XC_Section_dict, DC_L_dict, XC_L_dict, allowable_f_dict, gamma_0, pbar=None, svar=None, root=None, step_weight=8):
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
    # 容许应力（按层号索引）
    strut_dc_allowable_f_dict = allowable_f_dict['strut_dc']
    strut_xc_allowable_f_dict = allowable_f_dict['strut_xc']
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
        # Strut_Force_Picture_Stage(pic_path, version, Layer_lst,pbar=pbar, svar=svar, root=root, step_weight=step_weight)
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
            # img_paths = {
            #     "fx_max": os.path.join(pic_path, f"fx轴力图{layer_num}_max.jpg"),
            #     "fx_min": os.path.join(pic_path, f"fx轴力图{layer_num}_min.jpg"),
            #     "my_max": os.path.join(pic_path, f"my轴力图{layer_num}_max.jpg"),
            #     "my_min": os.path.join(pic_path, f"my轴力图{layer_num}_min.jpg"),
            #     "mz_max": os.path.join(pic_path, f"mz轴力图{layer_num}_max.jpg"),
            #     "mz_min": os.path.join(pic_path, f"mz轴力图{layer_num}_min.jpg")
            # }
            # try:
            #     fx_max_img = InlineImage(doc, img_paths["fx_max"], width=Mm(160))
            #     fx_min_img = InlineImage(doc, img_paths["fx_min"], width=Mm(160))
            #     my_max_img = InlineImage(doc, img_paths["my_max"], width=Mm(160))
            #     my_min_img = InlineImage(doc, img_paths["my_min"], width=Mm(160))
            #     mz_max_img = InlineImage(doc, img_paths["mz_max"], width=Mm(160))
            #     mz_min_img = InlineImage(doc, img_paths["mz_min"], width=Mm(160))
            # except Exception as e:
            #     print(f"第 {layer_num} 层图片读取失败: {e}")
            #     fx_max_img = fx_min_img = my_max_img = my_min_img = mz_max_img = mz_min_img = "图片缺失"
            layer = {
                'layer_num': layer_num,
                'cn_index': cn_index,
                'xc_type': xc_type,
                'zc_type': zc_type,
                'xc': xc,
                'zc': zc,
                'zh': zh,
                # "Fx_max_pic": fx_max_img,
                # "Fx_min_pic": fx_min_img,
                # "My_max_pic": my_max_img,
                # "My_min_pic": my_min_img,
                # "Mz_max_pic": mz_max_img,
                # "Mz_min_pic": mz_min_img
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


def Pile_cal(pile_leixing, allowable_f_dict, gamma_0, SECtype, CofferDam_L):
    pile_jisuan = []
    pile_allowable_f = allowable_f_dict['pile']['f']
    pile_allowable_d = round(float(CofferDam_L)*2.5, 1)
    # 获取截图
    API_DIST_FORCE_Unit("N", 'MM') # 应力为N和MM
    # 确保切换后处理状态，先截图
    # Picture_Beamstress(pic_path, "gbz_cb_pic.jpg", [SECtype], "基本组合", "Combined", 45, 30)
    # 初始化追踪变量
    max_cb_all_cases = -1.0
    max_d_all_cases = -1.0
    max_cb_case_name = ""
    max_d_case_name = ""

    case_name = '开挖至基坑底'
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
        'limit_cb': pile_allowable_f,                  # 容许应力
        'max_d': round(pile_max_d, 2),              # 最大位移
        'limit_d': pile_allowable_d,                    # 容许位移
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
    # API_DIST_FORCE_Unit("N", 'MM') # 变形为MM
    # Picture_Deformed(pic_path, "gbz_d_pic.jpg", [SECtype], "标准组合", pile_d_direction, 45, 30)
    # # 处理图片
    # cb_path = os.path.join(pic_path, "gbz_cb_pic.jpg")
    # # shear_path = os.path.join(pic_path, "gbz_shear_pic.jpg")
    # d_path = os.path.join(pic_path, "gbz_d_pic.jpg")
    # # 读取图片
    # cb_img = InlineImage(doc, cb_path, width=Mm(160))
    # # shear_img = InlineImage(doc, shear_path, width=Mm(100))
    # d_img = InlineImage(doc,d_path, width=Mm(160))
    pile_data={
        "SECtype": SECtype,
        "pile_jisuan": pile_jisuan,
        # "pile_cb_pic": cb_img,
        # "pile_d_pic": d_img
    }
    # 处理结论
    daxiao_cb = "≤" if max_cb_all_cases <= float(pile_allowable_f) else ">"
    daxiao_d = "≤" if max_d_all_cases <= float(pile_allowable_d) else ">"
    pile_res = f"{SECtype}在{max_cb_case_name}下组合应力最大值为{max_cb_all_cases:.2f} MPa{daxiao_cb}{pile_allowable_f}MPa，在{max_d_case_name}下变形最大值为 {max_d_all_cases:.2f}mm{daxiao_d}{pile_allowable_d}mm，{if_satisfied}要求；"
    return pile_data, pile_res, if_satisfied, max_cb_all_cases, max_d_all_cases


def Waler_cal(has_waler, WL_group_namelst, Waler_SECT_lst, allowable_f_dict, gamma_0, pbar=None, svar=None, root=None, step_weight=8):
    WL_layers_data = []
    WL_res = []
    wl_summary = {}
    wl_cb_limit_lst = [v['f'] for v in allowable_f_dict['waler']]
    wl_shear_limit_lst = [v['fv'] for v in allowable_f_dict['waler']]
    waler_allowable_d = ''
    if has_waler:
        num_layers = len(WL_group_namelst)
        increment = step_weight / num_layers if num_layers > 0 else 0
        # 按层读取替换
        for i, WL_group_name in enumerate(WL_group_namelst):
            idx = i+1
            WL_cb_limit = wl_cb_limit_lst[i]
            WL_shear_limit = wl_shear_limit_lst[i]
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

            case_name = '开挖至基坑底'
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
                "limit_cb": WL_cb_limit,            # 组合应力限值
                "limit_shear": WL_shear_limit,      # 剪应力限值
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
            # API_DIST_FORCE_Unit("N", 'MM') # 应力为N和MM
            # Picture_Beamstress(pic_path, f"wl{i+1}_cb_pic.jpg", [WL_group_namelst[i]], "基本组合", "Combined", 0, 90)
            # Picture_Beamstress(pic_path, f"wl{i+1}_shear_pic.jpg", [WL_group_namelst[i]], "基本组合", "Ssz", 0, 90)
            # Picture_Deformed(pic_path, f"WL{i+1}_d_pic.jpg", [WL_group_namelst[i]], "标准组合", WL_d_direction, 0, 90)
            # # 处理图片
            # cb_path = os.path.join(pic_path, f"wl{idx}_cb_pic.jpg")
            # shear_path = os.path.join(pic_path, f"wl{idx}_shear_pic.jpg")
            # d_path = os.path.join(pic_path, f"wl{idx}_d_pic.jpg")
            # try:
            #     cb_img = InlineImage(doc, cb_path, width=Mm(160))
            #     shear_img = InlineImage(doc, shear_path, width=Mm(160))
            #     d_img = InlineImage(doc, d_path, width=Mm(160))
            # except Exception as e:
            #     print(f"第 {idx} 层图片读取失败: {e}")
            #     cb_img = shear_img = d_img = "图片缺失"
            # 内容汇总
            WL_layers_data.append({
                "index": idx,
                "cn_index": num_to_chinese_num(idx),
                "WL_type": Waler_SECT_lst[i],
                "WLjisuan": WLjisuan,
                # "WL_cb_pic":cb_img,
                # "WL_shear_pic":shear_img,
                # "WL_d_pic":d_img
            })
            # 处理结论
            daxiao_cb = "≤" if layer_max_cb <= float(WL_cb_limit) else ">"
            daxiao_shear = "≤" if layer_max_shear <= float(WL_shear_limit) else ">"
            layer_summary = f"第{idx}层围檩在{layer_max_cb_case}下组合应力最大值为{layer_max_cb:.2f}MPa{daxiao_cb}{WL_cb_limit}MPa，在{layer_max_shear_case}下剪应力最大值为{layer_max_shear:.2f}MPa{daxiao_shear}{WL_shear_limit}MPa；"
            WL_res.append(layer_summary)
            wl_summary[idx] = {'max_cb': layer_max_cb, 'max_shear': layer_max_shear, 'max_d': layer_max_d}
    return WL_layers_data, WL_res, wl_summary


def Strut_cal(has_strut, Strut_section_dict, XC_group_namelst,DC_group_namelst, DC_Section_dict, XC_Section_dict, DC_L_dict, XC_L_dict, allowable_f_dict, gamma_0, pbar=None, svar=None, root=None, step_weight=8):
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
    # 容许应力（按层号索引）
    strut_dc_allowable_f_dict = allowable_f_dict['strut_dc']
    strut_xc_allowable_f_dict = allowable_f_dict['strut_xc']
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
        # Strut_Force_Picture(pic_path, Layer_lst,pbar=pbar, svar=svar, root=root, step_weight=step_weight)
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
            # img_paths = {
            #     "fx": os.path.join(pic_path, f"fx轴力图{layer_num}.jpg"),
            #     "my": os.path.join(pic_path, f"my轴力图{layer_num}.jpg"),
            #     "mz": os.path.join(pic_path, f"mz轴力图{layer_num}.jpg")
            # }
            # try:
            #     fx_img = InlineImage(doc, img_paths["fx"], width=Mm(160))
            #     my_img = InlineImage(doc, img_paths["my"], width=Mm(160))
            #     mz_img = InlineImage(doc, img_paths["mz"], width=Mm(160))
            # except:
            #     fx_img = my_img = mz_img = "图片未生成"
            layer = {
                'layer_num': layer_num,
                'cn_index': cn_index,
                'xc_type': xc_type,
                'zc_type': zc_type,
                'xc': xc,
                'zc': zc,
                'zh': zh,
                # "Fx_pic": fx_img,
                # "My_pic": my_img,
                # "Mz_pic": mz_img
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


def Model_cal_res(
        version, has_stage, load_cases, has_waler, has_strut, Strut_section_dict,
        pile_leixing, allowable_f_dict, Waler_SECT_lst, WL_group_namelst, 
        XC_group_namelst, DC_group_namelst, DC_Section_dict, XC_Section_dict, DC_L_dict, XC_L_dict, 
        gamma_0, SECtype, CofferDam_L, pbar=None, svar=None, root=None
        ):
    if svar:
        svar.set(f"正在进行模型结果验算\n正在验算{SECtype}...")
    # 加载模板
    if has_stage:
        # 获取计算结果
        pile_data, pile_res, pile_if_satisfied, pile_max_cb_all, pile_max_d_all = Pile_cal_Stage(version, pile_leixing, allowable_f_dict, gamma_0, SECtype, CofferDam_L)
        # 围檩验算
        WL_layers_data, Waler_res, wl_summary = Waler_cal_Stage(version, has_waler, WL_group_namelst, Waler_SECT_lst, allowable_f_dict, gamma_0, pbar=pbar, svar=svar, root=root,step_weight=8)
        # 内支撑验算
        has_pile, has_profile, strut_layers, pile_stability, profile_stability, Strut_res, Strut_if_satisfied, DC_sigma_dict, XC_sigma_dict = Strut_cal_Stage(version, has_strut, Strut_section_dict, XC_group_namelst, DC_group_namelst, DC_Section_dict, XC_Section_dict, DC_L_dict, XC_L_dict, allowable_f_dict, gamma_0, pbar=pbar, svar=svar, root=root,step_weight=8)
    else:
        # 工况处理
        load_cases = [load_cases[-1]] # 取开挖至基坑底工况
        print(f"当前为整体建模模式，仅验算最终工况：{load_cases}")
        # 获取计算结果
        pile_data, pile_res, pile_if_satisfied, pile_max_cb_all, pile_max_d_all = Pile_cal(pile_leixing, allowable_f_dict, gamma_0, SECtype, CofferDam_L)
        # 围檩验算
        WL_layers_data, Waler_res, wl_summary = Waler_cal(has_waler, WL_group_namelst, Waler_SECT_lst, allowable_f_dict, gamma_0, pbar=pbar, svar=svar, root=root,step_weight=8)
        # 内支撑验算
        has_pile, has_profile, strut_layers, pile_stability, profile_stability, Strut_res, Strut_if_satisfied, DC_sigma_dict, XC_sigma_dict = Strut_cal(has_strut, Strut_section_dict, XC_group_namelst, DC_group_namelst, DC_Section_dict, XC_Section_dict, DC_L_dict, XC_L_dict, allowable_f_dict, gamma_0, pbar=pbar, svar=svar, root=root,step_weight=8)
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


def stability_cal(Excel_Data, Solid_Top_Level, Water_Top_Level, Method_for_Pressures_var, Active_Hlst, Solid_Pak_lst, Water_Pak_lst, Condition_Passive_Earth_Pressure_Hlst, Condition_Solid_Ppk_lst, Condition_Water_Ppk_lst, Waler_Strut_Zlst, Concrete_Bottom_Level, CofferDam_Bottom_Level, level, slice_width, solid_layer_lst, Ld, h, Ea_p0, count, cal_config=None):
    # 计算配置：未勾选的稳定性项跳过计算, 结果置为 ['/', '/'] 且不影响满足性结论
    cfg = cal_config if cal_config else {}
    stability_result = {}
    # 抗倾覆参数计算
    if cfg.get("嵌固稳定性分析", True):
        anti_overturn_stability_cal(stability_result, Method_for_Pressures_var, Active_Hlst, Solid_Pak_lst, Water_Pak_lst, Condition_Passive_Earth_Pressure_Hlst, Condition_Solid_Ppk_lst, Condition_Water_Ppk_lst, Waler_Strut_Zlst, CofferDam_Bottom_Level, count, Ld, h, level)
    else:
        stability_result['抗倾覆'] = ['/', '/']
        print("嵌固稳定性分析(抗倾覆)未勾选, 已跳过计算")
    # 抗隆起参数计算
    if cfg.get("抗隆起稳定性分析", True):
        heave_resistant_stability_cal(
            stability_result, Excel_Data, Ld, h, count, Ea_p0, level,
            Solid_Top_Level=Solid_Top_Level, Water_Top_Level=Water_Top_Level,
            Concrete_Bottom_Level=Concrete_Bottom_Level, CofferDam_Bottom_Level=CofferDam_Bottom_Level,
            include_water_column=cfg.get("抗隆起计入水柱", True))
    else:
        stability_result['抗隆起'] = ['/', '/']
        print("抗隆起稳定性分析未勾选, 已跳过计算")
    # 圆弧稳定性参数计算
    if cfg.get("滑动稳定性分析", True):
        Slice_Method_for_Circular_Slip_Surface(stability_result, Concrete_Bottom_Level, Solid_Top_Level, CofferDam_Bottom_Level, Water_Top_Level, Ea_p0, level, slice_width, solid_layer_lst, Ld)
    else:
        stability_result['圆弧滑动稳定性'] = ['/', '/']
        print("滑动稳定性分析(圆弧滑动)未勾选, 已跳过计算")
    return stability_result
