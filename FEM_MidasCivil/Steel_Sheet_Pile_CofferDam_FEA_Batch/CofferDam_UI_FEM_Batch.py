# 1. 标准库
import re
import os
import sys
import openpyxl
import tkinter as tk
from pathlib import Path
from tkinter import messagebox

# 2. 第三方库
import ttkbootstrap as tb
from ttkbootstrap.constants import *

# 3. 本地模块
sys.path.append(str(Path(__file__).parent.parent.parent))
from General.UIHandle    import dialog_headr, on_exit, if_Reg
from General.ExcelHandle import ExcelApp_Dispatch, read_excel_to_dict_by_Sheet
from General.FilePath    import qucik_open_file_dialog, quick_save_folder_dialog

sys.path.append(str(Path(__file__).parent))
from FEM_MidasCivil.Steel_Sheet_Pile_CofferDam_FEA.CofferDam_Generate_FEM import Steel_Sheet_Pile_CofferDam_on_submit_MCT as Generate_MCT
from FEM_MidasCivil.Steel_Sheet_Pile_CofferDam_FEA.CofferDam_Excel_io_FEM import _find_sheet, load_section_library, load_material_library
from FEM_MidasCivil.Steel_Sheet_Pile_CofferDam_FEA.CofferDam_Tab_FEM      import _parse_assist_add_data, _parse_assist_bracket_data, _resolve_steel_std


# ===== 批量处理辅助 =====

def load_waler_strut_sections(excel_path):
    """从数据选项 sheet 按表头动态读取各列截面列表，
    并从 properties_parameter.xlsx 截面库补齐截面尺寸参数。

    读取逻辑：第1行为表头(key)，第2行起为数据。
    同一列非空、非"/"的值收集为 list[str]，以表头为 key 组成 dict。
    然后通过固定的表头 key 取出围檩和内支撑列表。

    Returns: (waler_list: list[str], strut_list: list[str],
              waler_dict: dict, strut_dict: dict)
    其中 dict 格式为 {name: {'SEC': [H,B,tw,tf], 'Ix':..., 'Wx':...}}
    """
    if not excel_path or not os.path.exists(excel_path):
        return [], [], {}, {}
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = _find_sheet(wb, ["数据选项"])
    if ws is None:
        wb.close()
        return [], [], {}, {}

    # 读取截面库
    section_path = os.path.join(os.path.dirname(excel_path), 'properties_parameter.xlsx')
    print(section_path)
    if not os.path.exists(section_path):
        # 回退：从注册表路径查找
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\ShuZhiQiaoShi')
            app, _ = winreg.QueryValueEx(key, 'Applocation')
            winreg.CloseKey(key)
            section_path = os.path.join(app, 'Support', 'basic_param', 'properties_parameter.xlsx')
        except:
            pass
    library = load_section_library(section_path)

    def _lookup_section(name):
        """在截面库中查找名称，返回完整 props 或降级解析"""
        clean = str(name).strip()
        if not clean:
            return None
        # 1. 直接匹配
        if clean in library:
            return library[clean]
        # 2. 钢管桩：去掉"钢管桩"后缀
        if clean.endswith('钢管桩'):
            base = clean.replace('钢管桩', '').strip()
            if base in library:
                return library[base]
        # 3. 去掉空格再试
        no_space = clean.replace(' ', '')
        if no_space in library:
            return library[no_space]
        return None

    # ---- 按表头动态读取各列 ----
    # 第1行：表头
    headers = {}
    for c in range(1, ws.max_column + 1):
        val = ws.cell(1, c).value
        if val is not None and str(val).strip():
            headers[c] = str(val).strip()

    # 第2行起：按列收集非空值
    columns_data = {h: [] for h in headers.values()}
    for r in range(2, ws.max_row + 1):
        for c, header_name in headers.items():
            val = ws.cell(r, c).value
            if val is not None:
                s = str(val).strip()
                if s and s != "/":
                    columns_data[header_name].append(s)

    # ---- 通过固定 key 取出围檩和内支撑列表 ----
    WALER_KEY = "围檩截面可选类型"
    STRUT_KEY = "内支撑截面可选类型"
    waler_secs = columns_data.get(WALER_KEY, [])
    strut_secs = columns_data.get(STRUT_KEY, [])

    waler_dict = {}
    for name in waler_secs:
        props = _lookup_section(name)
        waler_dict[name] = props if props else {}

    strut_dict = {}
    for name in strut_secs:
        props = _lookup_section(name)
        strut_dict[name] = props if props else {}

    wb.close()
    return waler_secs, strut_secs, waler_dict, strut_dict


def prepare_submit_variables(program_name, param_dict, proj_index_dict, soil_param_dict, pile_sec_param_dict, elink_param_lst, save_path, excel_path=None):
    """
    从 Batch_CofferDam_Excel_Data_Process 解析出的参数字典中，
    提取 Steel_Sheet_Pile_CofferDam_on_submit_MCT (Generate版) 所需的全部变量。

    Args:
        param_dict: 主表参数字典（单个墩号）
        soil_param_dict: 土层参数字典，key=sheet名, value={土层名: [层厚, 重度, ...]}
        save_path: mct 保存路径

    Returns: dict — key 为 Generate 函数的形参名, value 为对应的 Python 值
    """
    v = param_dict

    # 从主表 key 列表中提取 围檩/内支撑 相关 key
    all_keys = list(v.keys())
    Waler_ilst = [k for k in all_keys if '围檩' in k and ('间距' in k or '截面' in k)]
    Strut_ilst = [k for k in all_keys if '内支撑' in k and ('对撑' in k or '斜撑' in k)]
    # 按层号分组排序
    def _layer_num(s):
        m = re.search(r'(\d+)', s)
        return int(m.group(1)) if m else 0
    Waler_ilst.sort(key=_layer_num)
    Strut_ilst.sort(key=_layer_num)
    # print(Waler_ilst)
    # print(Strut_ilst)

    # 围檩层数 = 不同层号的个数
    waler_layer_nums = sorted(set(_layer_num(k) for k in Waler_ilst))
    # 内支撑 key 按层号分组
    strut_by_layer = {}
    for k in Strut_ilst:
        n = _layer_num(k)
        strut_by_layer.setdefault(n, []).append(k)
    # print(waler_layer_nums)
    # print(strut_by_layer)

    # 基本几何
    Cap_X            = float(v['承台长(m)'])
    Cap_Y            = float(v['承台宽(m)'])
    Cap_H            = float(v['承台高(m)'])
    X_offset         = float(v['承台长边预留边距(m)'])
    Y_offset         = float(v['承台宽边预留边距(m)'])
    Cap_Bottom_Level = float(v['承台底标高(m)'])

    # 标高
    Solid_Level         = float(v['初始地面标高(m)'])
    Water_Level         = v['初始水位标高(m)'] if v['初始水位标高(m)'] in ('/', 'None') else float(v['初始水位标高(m)'])
    CofferDam_Top_Level = float(v['围堰顶标高'])

    if_consider_solid_stress_path = True if v['考虑土的应力路径'] == '是' else False

    # 钢板桩
    CofferDam_L        = float(v['围堰高'])
    Sheet_Pile_typevar = v['支护桩类型']
    Sheet_Pile_namevar = [value['截面名称'] for value in pile_sec_param_dict.values() if int(value['截面号']) == int(v['支护桩截面编号'])][0]

    # 支护桩材质（默认：钢板桩 Q295 / 锁扣钢管桩 Q235）
    Sheet_Pile_materialvar = v.get('支护桩材质', '')
    if not Sheet_Pile_materialvar or Sheet_Pile_materialvar in ('/', 'None'):
        Sheet_Pile_materialvar = 'Q295' if Sheet_Pile_typevar == '钢板桩' else 'Q235'

    # 截面特性（从 pile_sec_param_dict 读取完整数据）
    SEC_KEY_ORDER = [
        "H(mm)", "B(mm)", "As(mm^2)", "Asy(mm^2)", "Asz(mm^2)",
        "Ixx(mm^4)", "Iyy(mm^4)", "Izz(mm^4)",
        "Cyp(mm)", "Cym(mm)", "Czp(mm)", "Czm(mm)",
        "Qyb(mm^2)", "Qzb(mm^2)", "Peri:O(mm)", "Peri:I(mm)",
        "Cent:y(mm)", "Cent:z(mm)",
        "y1(mm)", "z1(mm)", "y2(mm)", "z2(mm)", "y3(mm)", "z3(mm)", "y4(mm)", "z4(mm)",
        "Zyy(mm^3)", "Zzz(mm^3)",
    ]
    sec_info = pile_sec_param_dict.get(Sheet_Pile_namevar, {})
    sec_params = sec_info.get('params', {})
    if Sheet_Pile_typevar == '钢板桩':
        sec_vals = [sec_params.get(k, 0) for k in SEC_KEY_ORDER]
        Pile_SEC_info_dict = {'SEC': sec_vals if any(sec_vals) else [400, 170] + [0]*26}
        SKGGZ_D, SKGGZ_t, SKGGZ_Gap = 0, 0, 0
    elif Sheet_Pile_typevar == '锁扣钢管桩':
        Pile_SEC_info_dict = {'SEC': []}
        D = float(sec_params.get('D', 0.0))
        t = float(sec_params.get('t', 0.0))
        lock_width = float(sec_params.get('锁扣宽度', 35.0))
        sheet_count = int(sec_params.get('钢板桩数量', 1))
        sheet_width = float(sec_params.get('钢板桩宽度', 600.0))
        SKGGZ_D = D
        SKGGZ_t = t
        SKGGZ_Gap = D + 2*lock_width + sheet_width*sheet_count
    else:
        Pile_SEC_info_dict = {'SEC': []}
        SKGGZ_D, SKGGZ_t, SKGGZ_Gap = 0.0, 0.0, 0.0

    # 围檩/内支撑截面字典（从数据选项 Excel 读取，路径为 batch Excel 同目录下的 properties_parameter.xlsx）
    # section_excel_path = os.path.join(os.path.dirname(excel_path), 'properties_parameter.xlsx') if excel_path else get_excel_path()
    _, _, Waler_section_dict, Strut_section_dict = load_waler_strut_sections(excel_path)

    # 钢材材质库 + 钢结构规范（用于 material_dict 的 standard / elast）
    Steel_dict, _, _ = load_material_library(excel_path)
    steel_std = _resolve_steel_std(v.get('钢结构规范', ''), Steel_dict)

    # 垫层
    Concrete_Blinding_thickness = float(v['垫层厚度(m)']) if v['垫层厚度(m)'] != '/' else 0
    Concrete_Blinding_check     = Concrete_Blinding_thickness > 0

    # 封底
    Concrete_Plug_thickness = float(v['封底厚度(m)']) if v['封底厚度(m)'] != '/' else 0
    Concrete_Plug_check     = Concrete_Plug_thickness > 0

    # 牛腿数量
    corbel_long  = int(v['牛腿长边布置数量']) if v['牛腿长边布置数量'] != '/' else 2
    corbel_short = int(v['牛腿短边布置数量']) if v['牛腿短边布置数量'] != '/' else 2

    # 超挖/降水
    Drawdown_height     = float(v.get('超挖深度(m)', 0))
    Waterdown_height    = float(v.get('水面至基坑底距离(m)', 0))
    Excavation_face_dewater = float(v.get('开挖面降水高度(m)', 0))

    # 土弹簧分层厚度
    Spring_Thickness = float(v.get('土弹簧分层厚度(m)', 0))

    # 围檩/内支撑 Waler_rows / Strut_blocks
    # 主表 key 格式: '围檩1间距(m)', '围檩1长边截面', '围檩1宽边截面'
    #               '内支撑1对撑截面', '内支撑1斜撑截面', '内支撑1对撑长边布置(m)', ...
    Waler_rows = []
    Strut_blocks = {}
    for i, layer_num in enumerate(waler_layer_nums):
        # 围檩：从扁平 key 中提取
        spacing_key = f'围檩{layer_num}间距(m)'
        sec_long_key = f'围檩{layer_num}长边截面'
        sec_short_key = f'围檩{layer_num}宽边截面'
        spacing = v[spacing_key]
        if not spacing or spacing in ('None', '/'):
            pass
        else:
            sec_long = v[sec_long_key]
            sec_short = v[sec_short_key]
            Waler_rows.append({
                'entry'      : tb.StringVar(value = spacing),
                'combo_long' : tb.StringVar(value = sec_long),
                'combo_short': tb.StringVar(value = sec_short),
                'combo2'     : tb.StringVar(value = '/'),
                'combo3'     : tb.StringVar(value = '/'),
                'mat_long'   : tb.StringVar(value = v.get(f'围檩{layer_num}长边材质', '/')),
                'mat_short'  : tb.StringVar(value = v.get(f'围檩{layer_num}宽边材质', '/')),
                'mat_dc'     : tb.StringVar(value = '/'),
                'mat_xc'     : tb.StringVar(value = '/'),
            })
            Strut_blocks[i] = {
                'DC_X': tb.StringVar(value = '/'), 'DC_Y': tb.StringVar(value = '/'),
                'XC_X': tb.StringVar(value = '/'), 'XC_Y': tb.StringVar(value = '/'),
            }
            # 内支撑：从扁平 key 中提取
            if layer_num in strut_by_layer:
                dc_sec_key = f'内支撑{layer_num}对撑截面'
                xc_sec_key = f'内支撑{layer_num}斜撑截面'
                dc_x_key = f'内支撑{layer_num}对撑长边布置(m)'
                dc_y_key = f'内支撑{layer_num}对撑短边布置(m)'
                xc_x_key = f'内支撑{layer_num}斜撑长边布置(m)'
                xc_y_key = f'内支撑{layer_num}斜撑短边布置(m)'
                dc_sec = v.get(dc_sec_key, '/')
                xc_sec = v.get(xc_sec_key, '/')
                if dc_sec not in ('/', 'None'):
                    Waler_rows[i]['combo2'].set(dc_sec)
                    Waler_rows[i]['mat_dc'].set(v.get(f'内支撑{layer_num}对撑材质', '/'))
                    Strut_blocks[i]['DC_X'] = tb.StringVar(value = v.get(dc_x_key, '/'))
                    Strut_blocks[i]['DC_Y'] = tb.StringVar(value = v.get(dc_y_key, '/'))
                if xc_sec not in ('/', 'None'):
                    Waler_rows[i]['combo3'].set(xc_sec)
                    Waler_rows[i]['mat_xc'].set(v.get(f'内支撑{layer_num}斜撑材质', '/'))
                    Strut_blocks[i]['XC_X'] = tb.StringVar(value = v.get(xc_x_key, '/'))
                    Strut_blocks[i]['XC_Y'] = tb.StringVar(value = v.get(xc_y_key, '/'))

    # 土层信息（从 soil_param_dict 中获取）
    soil_ref = v.get('地质参数', '')
    soil_sheet_name = soil_ref.split(',')[1] if ',' in soil_ref else soil_ref
    soil_data_from_sheet = soil_param_dict.get(soil_sheet_name, {})
    # soil_data_from_sheet: {土层名: [层厚, 重度, 黏聚力, 内摩擦角, 计算方法, 渗透系数, 土层顶承压水头]}
    # 加权平均土层参数：加权列能解析出完整 γ,c,φ 时，用加权单层土代替分层土层（与 UI 版一致）
    def _parse_wavg(raw):
        if raw is None or str(raw).strip() in ('', '/', 'None'):
            return None
        parts = [p.strip() for p in str(raw).replace('，', ',').split(',')]
        if len(parts) >= 3:
            try:
                return (float(parts[0]), float(parts[1]), float(parts[2]))
            except (ValueError, TypeError):
                return None
        return None

    _wavg = _parse_wavg(v.get('加权平均重度+粘聚力+摩擦角', '/'))
    if _wavg is not None and _wavg[0] <= 0:
        # 加权数据异常（重度须>0），退回按分层土层计算
        print(f"[CofferDam_Batch] {program_name} 加权平均重度({_wavg[0]})不合法(须>0)，已退回按分层土层计算")
        _wavg = None
    if _wavg is not None:
        _wg, _wc_, _wf_ = _wavg
        _wtop = float(Solid_Level)
        _wbot = float(CofferDam_Top_Level) - float(CofferDam_L)  # 围堰底=围堰顶-围堰高
        _soil_excel = {
            '第1层土': {
                '土层名称': '加权平均土层',
                '层厚': round(_wtop - _wbot + 1, 3),
                '重度': round(_wg, 3),
                '黏聚力': round(_wc_, 3),
                '内摩擦角': round(_wf_, 3),
                '计算方法': '水土合算',
                '渗透系数(m/d)': 0,
                '土层顶承压水头(m)': 0,
            }
        }
        Method_for_Pressures_lst = ['水土合算']
    else:
        _soil_excel = {
            f'第{i+1}层土': {
                '土层名称':        name,
                '层厚':            float(vals[0]) ,
                '重度':            float(vals[1]) ,
                '黏聚力':          float(vals[2]) ,
                '内摩擦角':        float(vals[3]) ,
                '计算方法':        str(vals[4]) ,
                '渗透系数(m/d)':   float(vals[5]) ,
                '土层顶承压水头(m)': float(vals[6]) ,
            } for i, (name, vals) in enumerate(soil_data_from_sheet.items())
        }
        Method_for_Pressures_lst = [value[4] for value in soil_data_from_sheet.values()]
    Steel_Sheet_Pile_CofferDam_Load_dict = {'Excel': _soil_excel}

    # 荷载组合（根据计算方法自动构建）
    Method_for_Pressures = '水土合算' if any(m == '水土合算' for m in Method_for_Pressures_lst) else '水土分算'
    load_grup_lst = ['自重', '基坑外主动土压力', '基坑底初始土反力']
    if Method_for_Pressures == '水土分算':
        load_grup_lst.insert(2, '基坑外水压力')
        load_grup_lst.append('基坑底水压力')
    # 从主表读取组合系数
    jb_factor = v.get('基本组合分项系数', '1.25')
    jb_combo  = v.get('基本组合组合值系数', '1.0')
    bz_factor = v.get('标准组合分项系数', '1.0')
    bz_combo  = v.get('标准组合组合值系数', '1.0')
    Load_Combo_Line_Data = []
    for i, load in enumerate(load_grup_lst):
        Load_Combo_Line_Data.append([str(i+1), load, jb_factor, jb_combo, bz_factor, bz_combo])
    # for _ in range(15 - len(load_grup_lst)):
    #     Load_Combo_Line_Data.append(['', '', '', '', '', ''])

    # 施工阶段
    if v['是否生成施工阶段'] == '是':
        construction_stage_check = True
        none_construction_stage_check = False
    else:
        construction_stage_check = False
        none_construction_stage_check = True
    # 施工阶段
    stage_dict = {}
    if construction_stage_check:
        stage_name_lst = v['各工况名称'].split(',')
        stage_soil_level_lst = v['各工况基坑内土顶标高(m)'].split(',')
        stage_water_level_lst = v['各工况基坑内水面标高(m)'].split(',')
        stage_layer_change_lst = v['各工况支撑变动'].split(',')
        for i in range(len(stage_name_lst)):
           stage_dict[i+1] = {
               '工况类型': stage_name_lst[i], 
               '基坑内土顶标高': stage_soil_level_lst[i], 
               '基坑内水面标高': stage_water_level_lst[i], 
               '支撑层数': stage_layer_change_lst[i],
           } 

    # Construct_Stage_Line_Data: [[工况类型, 土顶标高, 水面标高, 支撑层数], ...]
    Construct_Stage_Line_Data = [list(value.values()) for value in stage_dict.values() ]

    # mct 保存路径
    project, index =  proj_index_dict[program_name]
    mct_savepath = os.path.join(save_path, f'{project}-{index}.mct')

    # recognize 字典（Generate 内部填充）
    recognize_Cap_result_dict   = {}
    recognize_Bracket_result_dict = {}
    recognize_Waler_result_dict = {}
    recognize_Strut_result_dict = {}
    recognize_Strut_Replace_result_dict = {}

    # 超载 sigma_k_dict（从主表超载字段构建）
    # 结构: {active_type, 均布附加荷载: {面号: {q0}}, 矩形局部附加荷载: {面号: {...}}, 条形局部附加荷载: {面号: {...}}}
    sigma_k_dict = {
        '均布附加荷载':{
            '1': {'q0': '/'}, 
            '2': {'q0': '/'}, 
            '3': {'q0': '/'}, 
            '4': {'q0': '/'}
        }, 
        '矩形局部附加荷载':{
            '1': {'p0': '/', 'angle': '/', 'b': '/', 'a': '/', 'd': '/', 'l': '/', 'p2': '/', 'c': '/'}, 
            '2': {'p0': '/', 'angle': '/', 'b': '/', 'a': '/', 'd': '/', 'l': '/', 'p2': '/', 'c': '/'}, 
            '3': {'p0': '/', 'angle': '/', 'b': '/', 'a': '/', 'd': '/', 'l': '/', 'p2': '/', 'c': '/'}, 
            '4': {'p0': '/', 'angle': '/', 'b': '/', 'a': '/', 'd': '/', 'l': '/', 'p2': '/', 'c': '/'}
        }, 
        '条形局部附加荷载':{
            '1': {'p0': '/', 'angle': '/', 'b': '/', 'a': '/', 'd': '/'},
            '2': {'p0': '/', 'angle': '/', 'b': '/', 'a': '/', 'd': '/'}, 
            '3': {'p0': '/', 'angle': '/', 'b': '/', 'a': '/', 'd': '/'}, 
            '4': {'p0': '/', 'angle': '/', 'b': '/', 'a': '/', 'd': '/'}
        }
    }
    junbu = v['均布附加荷载']
    juxing = v['矩形局部附加荷载']
    tiaoxing = v['条形局部附加荷载']
    all_face = ['1','2','3','4']

    if junbu != '/':
        face, p = junbu.split(', ')
        face = re.findall(r'-?\d+\.?\d*', face)
        for i in all_face:
            if i in face:
                sigma_k_dict['均布附加荷载'][i]['q0'] = p
    if juxing != '/':
        face, p = juxing.split(', ')
        face = re.findall(r'-?\d+\.?\d*', face)
        for i in all_face:
            if i in face:
                p0, angle, b, a, d, l, p2, c = p.split(',')
                sigma_k_dict['矩形局部附加荷载'][i]['p0'] = p0
                sigma_k_dict['矩形局部附加荷载'][i]['angle'] = angle
                sigma_k_dict['矩形局部附加荷载'][i]['b'] = b
                sigma_k_dict['矩形局部附加荷载'][i]['a'] = a
                sigma_k_dict['矩形局部附加荷载'][i]['d'] = d
                sigma_k_dict['矩形局部附加荷载'][i]['l'] = l
                sigma_k_dict['矩形局部附加荷载'][i]['p2'] = p2
                sigma_k_dict['矩形局部附加荷载'][i]['c'] = c
    if tiaoxing != '/':
        face, p = tiaoxing.split(', ')
        face = re.findall(r'-?\d+\.?\d*', face)
        for i in all_face:
            if i in face:
                p0, angle, b, a, d = p.split(',')
                sigma_k_dict['条形局部附加荷载'][i]['p0'] = p0
                sigma_k_dict['条形局部附加荷载'][i]['angle'] = angle
                sigma_k_dict['条形局部附加荷载'][i]['b'] = b
                sigma_k_dict['条形局部附加荷载'][i]['a'] = a
                sigma_k_dict['条形局部附加荷载'][i]['d'] = d

    # 弹性连接（从 elink_param_lst 构建 form_dict 和 stiffness_dict）
    # 连接定义：4个固定连接类型
    # 默认连接类型映射
    cofferdam_elasticlink_form_dict = {
        "牛腿与支护桩连接": v['牛腿与支护桩连接'],
        "牛腿与围檩连接"  : v['牛腿与围檩连接'],
        "围檩与支护桩连接": v['围檩与支护桩连接'],
        "围檩与内支撑连接": v['围檩与内支撑连接'],
    }

    # 构建 stiffness_dict：非"共节点"的连接类型需要刚度
    sheet_elasticlink_dict = {}
    for _, conn_type in cofferdam_elasticlink_form_dict.items():
        if conn_type and conn_type != "共节点":
            for e in elink_param_lst:
                if e['name'] == conn_type:
                    gen = {k.lower(): e.get(k, '/') for k in ['SDx', 'SDy', 'SDz', 'SRx', 'SRy', 'SRz']}
                    comp = {'nsdx': e.get('NSDx', '/')}
                    sheet_elasticlink_dict[conn_type] = {"GEN": gen, "COMP": comp}

    Pile_Bottom_Boundary = v.get('支护桩底边界', '001001')

    # 辅助加撑（从 "辅助加围檩i" + "辅助加内支撑i" 解析）
    waler_raw = v.get('辅助加围檩i', '/')
    strut_raw = v.get('辅助加内支撑i', '/')
    _assist_add_data = _parse_assist_add_data(waler_raw, strut_raw)
    assist_add_dict = {}
    for item in _assist_add_data:
        layer_no = item.get('layer_no', '')
        if not layer_no:
            continue
        assist_add_dict[layer_no] = {
            '标高': item.get('elevation', '/'),
            '围檩长边材质': item.get('mat_long', '/'),
            '围檩长边截面': item.get('sec_long', '/'),
            '围檩宽边材质': item.get('mat_short', '/'),
            '围檩宽边截面': item.get('sec_short', '/'),
            '是否与承台连接': item.get('connected', '否'),
            '对撑材质': item.get('mat_dc', '/'),
            '对撑截面': item.get('dc_sec', '/'),
            '斜撑材质': item.get('mat_xc', '/'),
            '斜撑截面': item.get('xc_sec', '/'),
            '对撑长边布置(m)': item.get('dc_x', '/'),
            '对撑短边布置(m)': item.get('dc_y', '/'),
            '斜撑长边布置(m)': item.get('xc_x', '/'),
            '斜撑短边布置(m)': item.get('xc_y', '/'),
            '承台长边投影(m)': item.get('proj_long', '/'),
            '承台短边投影(m)': item.get('proj_short', '/'),
            '承台长边高差(m)': item.get('diff_long', '/'),
            '承台短边高差(m)': item.get('diff_short', '/'),
        }

    # 辅助换撑（从 "辅助换撑" 解析）
    replace_raw = v.get('辅助换撑', '/')
    _assist_replace_data = _parse_assist_bracket_data(replace_raw)
    assist_replace_dict = {}
    for item in _assist_replace_data:
        name = item.get('name', '')
        if not name or '-' not in name:
            continue
        parts = name.split('-')
        try:
            layer_i = int(parts[0])
            form_j = int(parts[1])
        except (ValueError, IndexError):
            continue
        if layer_i not in assist_replace_dict:
            assist_replace_dict[layer_i] = {}
        assist_replace_dict[layer_i][form_j] = {
            '是否与承台连接': item.get('connected', '否'),
            '对撑材质': item.get('mat_dc', '/'),
            '对撑截面': item.get('dc_sec', '/'),
            '斜撑材质': item.get('mat_xc', '/'),
            '斜撑截面': item.get('xc_sec', '/'),
            '对撑长边布置(m)': item.get('dc_x', '/'),
            '对撑短边布置(m)': item.get('dc_y', '/'),
            '斜撑长边布置(m)': item.get('xc_x', '/'),
            '斜撑短边布置(m)': item.get('xc_y', '/'),
            '承台长边投影(m)': item.get('proj_long', '/'),
            '承台短边投影(m)': item.get('proj_short', '/'),
            '承台长边高差(m)': item.get('diff_long', '/'),
            '承台短边高差(m)': item.get('diff_short', '/'),
        }

    return {
        'Cap_X': Cap_X, 'Cap_Y': Cap_Y, 'Cap_H': Cap_H, 'X_offset': X_offset, 'Y_offset': Y_offset,
        'Spring_Thickness': Spring_Thickness,
        'Sheet_Pile_typevar': Sheet_Pile_typevar, 'CofferDam_L': CofferDam_L,
        'Sheet_Pile_namevar': Sheet_Pile_namevar,
        'Sheet_Pile_materialvar': Sheet_Pile_materialvar,
        'steel_dict': Steel_dict,
        'steel_std': steel_std,
        'Pile_SEC_info_dict': Pile_SEC_info_dict,
        'SKGGZ_D': SKGGZ_D, 'SKGGZ_t': SKGGZ_t, 'SKGGZ_Gap': SKGGZ_Gap,
        'Waler_rows': Waler_rows, 'Strut_blocks': Strut_blocks,
        'corbel_long': corbel_long, 'corbel_short': corbel_short,
        'Solid_Level': Solid_Level, 'Water_Level': Water_Level,
        'Cap_Bottom_Level': Cap_Bottom_Level, 'CofferDam_Top_Level': CofferDam_Top_Level,
        'Concrete_Blinding_check': Concrete_Blinding_check,
        'Concrete_Blinding_thickness': Concrete_Blinding_thickness,
        'Concrete_Plug_check': Concrete_Plug_check,
        'Concrete_Plug_thickness': Concrete_Plug_thickness,
        'Drawdown_height': Drawdown_height,
        'Waterdown_height': Waterdown_height,
        'Excavation_face_dewater': Excavation_face_dewater,
        'Waler_section_dict': Waler_section_dict,
        'Strut_section_dict': Strut_section_dict,
        'recognize_Cap_result_dict': recognize_Cap_result_dict,
        'recognize_Bracket_result_dict': recognize_Bracket_result_dict,
        'recognize_Waler_result_dict': recognize_Waler_result_dict,
        'recognize_Strut_result_dict': recognize_Strut_result_dict,
        'recognize_Strut_Replace_result_dict': recognize_Strut_Replace_result_dict,
        'Steel_Sheet_Pile_CofferDam_Load_dict': Steel_Sheet_Pile_CofferDam_Load_dict,
        'sigma_k_dict': sigma_k_dict,
        'Pile_Bottom_Boundary': Pile_Bottom_Boundary,
        'cofferdam_elasticlink_form_dict': cofferdam_elasticlink_form_dict,
        'sheet_elasticlink_dict': sheet_elasticlink_dict,
        'none_construction_stage_check': none_construction_stage_check,
        'construction_stage_check': construction_stage_check,
        'Load_Combo_Line_Data': Load_Combo_Line_Data,
        'Construct_Stage_Line_Data': Construct_Stage_Line_Data,
        'stage_dict': stage_dict,
        'if_consider_solid_stress_path': if_consider_solid_stress_path,
        'mct_savepath': mct_savepath,
        'assist_add_dict': assist_add_dict,
        'assist_replace_dict': assist_replace_dict,
    }


# 标签 + 编辑框 + 按钮
def dialog_boxline(root, label_txt, label_length, entry_var, entry_length, button_txt, button_command):
    container = tb.Frame(root)
    container.pack(fill=X, expand=YES, pady=10)
    # 标签
    Label = tb.Label(master=container, text=label_txt.title(), width=label_length)
    Label.pack(side=LEFT, padx=10)
    # 编辑框
    ent = tb.Entry(master=container, textvariable=entry_var, width=entry_length, state="disabled")
    ent.pack(side=LEFT, padx=20, fill=X, expand=YES)
    # 按钮
    but = tb.Button(master=container, text=button_txt, command = button_command, width = 15)
    but.pack(side=RIGHT, padx=10)


# 最后一行
def dialog_boxline_end(root, checked, button_txt1, button_command1, button_txt2, button_command2):
    container = tb.Frame(root)
    container.pack(fill=X, expand=YES, pady=10)
    # 创建复选框
    check_button = tb.Checkbutton(master=container, text="同时生成mcb模型", variable=checked, bootstyle="primary")
    check_button.pack(side=LEFT, padx=20, pady=10)
    # 退出按钮
    but = tb.Button(master=container, text=button_txt2, command = button_command2, width=6)
    but.pack(side=RIGHT, padx=10, pady=10)
    # 确认按钮
    sub_btn = tb.Button(master=container, text=button_txt1, command = button_command1, bootstyle=SUCCESS, width=6)
    sub_btn.pack(side=RIGHT, pady=10)


# 选择 Excel 文件
def Choose_Excel_File(entry_var, save_folder_var):
    # 选择路径
    openfile_path = qucik_open_file_dialog().replace("/", "\\")

    # 路径检测
    if "xlsx" not in openfile_path:
        # 文件无效
        messagebox.showinfo("提示", '未选择有效 xlsx 文件, 请重新选择。')
        entry_var.set('未选取任何文件')
    else:
        # 更新对应编辑框的值
        entry_var.set(openfile_path)
        # 自动设置保存文件夹为 Excel 文件所在目录
        save_folder_var.set(os.path.dirname(openfile_path))


# 计算书另存为得路径
def MCT_File_SaveAs(entry_var):
    # 选择路径
    savefile_path = quick_save_folder_dialog().replace("/", "\\")
    print('savefile_path')
    print( savefile_path )
    # open_csv_saveas_xlsx(savefile_path, Applocation)
    # 更新对应编辑框的值
    if savefile_path != '':
        entry_var.set(savefile_path)

#=====================================================================================================================================================================

def remove_all_parentheses(text):
    '''
    栈匹配删除所有括号对
    可用于字符串单位和括号删除
    s = "abc (def) ghi (jkl (mno) pqr) xyz"
    result = remove_all_parentheses(s)
    print(result)  # 输出: abc  ghi  xyz
    '''
    stack = []
    remove = [False] * len(text)
    for i, ch in enumerate(text):
        if ch == '(':
            stack.append(i)
        elif ch == ')' and stack:
            start = stack.pop()
            for j in range(start, i + 1):
                remove[j] = True
    return ''.join(ch for i, ch in enumerate(text) if not remove[i])


#=====================================================================================================================================================================

def Batch_CofferDam_Excel_Data_Process(excel_path):
    # 连接Excel
    excel = ExcelApp_Dispatch(False)
    WorkBook = excel.Workbooks.Open(os.path.abspath(excel_path))
    excel_data_lst = read_excel_to_dict_by_Sheet(WorkBook, 1) # 第一张Sheet
    excel_data_lst = excel_data_lst[1:]
    for x in excel_data_lst:
        print(x)
        print(len(x))
    print('='*80 + '\n')

    # 分离参数列和结果列
    # 按列顺序读取，参数列截止到 标准组合组合值系数（其后的围檩/支撑等列已在其中）
    param_last_str = '标准组合组合值系数'
    param_last_index = excel_data_lst[0].index(param_last_str)
    excel_parameter_data_lst = [paramlst[:param_last_index+1] for paramlst in excel_data_lst]
    for x in excel_parameter_data_lst:
        print(x)
        print(len(x))
    print('='*80 + '\n')

    # 先处理所有的空列
    None_ilst = []
    param_lst = []
    for sublst in excel_parameter_data_lst:
        param_lst.append([s for i, s in enumerate(sublst) if i not in None_ilst])
    for x in param_lst:
        print(x)
        print(len(x))
    print('='*80 + '\n')
    param_keylst = param_lst[0]
    param_valuelst = param_lst[1:]
    main_param_dict = {}
    proj_index_dict = {}
    for sublst in param_valuelst:
        key = f'{sublst[0]}{sublst[1]}' # 墩号
        proj_index_dict[key] = [sublst[0], sublst[1]]
        main_param_dict[key] = {}
        for i, k in enumerate(param_keylst[2:]):
            v_value = sublst[2:][i]
            main_param_dict[key][k] = str(v_value)
    for k, v in main_param_dict.items():
        print(k)
        print(v)
    print('='*80 + '\n')

    # 获取所有地质参数表中的数据
    soil_sheetname_lst = list(set([p['地质参数'].split(',')[1] for _, p in main_param_dict.items() if p.get('地质参数', '') and ',' in p['地质参数']]))
    print(f'土层sheet名列表: {soil_sheetname_lst}')
    soil_param_dict = {}
    SOIL_HDRS = ['土层名称', '层厚(m)', '重度(kN/m3)', '黏聚力c(kPa)', '内摩擦角φ(°)',
                 '计算方法', '渗透系数(m/d)', '土层顶承压水头(m)']
    for soil_sheetname in soil_sheetname_lst:
        soil_data_raw = read_excel_to_dict_by_Sheet(WorkBook, soil_sheetname)
        if soil_data_raw and len(soil_data_raw) > 1:
            soil_keylst = soil_data_raw[0]  # R1=表头
            soil_valuelst = soil_data_raw[1:]  # R2起=数据
            soil_data = {}
            for row in soil_valuelst:
                if not row or not row[0]:
                    continue
                name = str(row[0]).strip()
                if not name:
                    continue
                # 按表头顺序提取值（计算方法为字符串，其余为数值）
                vals = []
                for hdr in SOIL_HDRS[1:]:
                    try:
                        idx = soil_keylst.index(hdr)
                        v = row[idx] if idx < len(row) else None
                        if hdr == '计算方法':
                            vals.append(str(v).strip() if v is not None else '')
                        else:
                            vals.append(float(v) if v is not None else 0.0)
                    except (ValueError, IndexError):
                        vals.append(0.0)
                soil_data[name] = vals
            soil_param_dict[soil_sheetname] = soil_data
    for k, v in soil_param_dict.items():
        print(k)
        print(v)     
    print('='*80 + '\n')

    # 获取所有截面参数表中的数据（支护桩参数 sheet）
    pile_sec_param_dict = {}
    sec_data_raw = read_excel_to_dict_by_Sheet(WorkBook, '支护桩参数')
    if sec_data_raw and len(sec_data_raw) > 2:
        # R1=类别合并行, R2=表头行, R3起=数据
        sec_keylst = sec_data_raw[1]  # R2=表头
        sec_valuelst = sec_data_raw[2:]  # R3起=数据

        def _sec_get_val(row, hdr, default=None):
            """按表头名查找列号取值（列顺序无关）"""
            try:
                idx = sec_keylst.index(hdr)
                v = row[idx] if idx < len(row) else None
                return v if v is not None else default
            except (ValueError, IndexError):
                return default

        for row in sec_valuelst:
            if not row:
                continue
            name = str(_sec_get_val(row, '截面名称', '')).strip()
            if not name:
                continue
            # 基础字段
            seq_raw = _sec_get_val(row, '截面号')
            try:
                seq = int(seq_raw)
            except (TypeError, ValueError):
                seq = 0
            material = str(_sec_get_val(row, '材料标号', '')).strip()
            # 判断类型：有D(mm)列且非空则是锁扣钢管桩
            d_val = _sec_get_val(row, 'D(mm)')
            if d_val is not None and str(d_val).strip() not in ('', '/', '.'):
                # 锁扣钢管桩
                def _f_val(hdr, default=0.0):
                    v = _sec_get_val(row, hdr)
                    try:
                        return float(v) if v is not None else default
                    except (TypeError, ValueError):
                        return default
                def _i_val(hdr, default=0):
                    v = _sec_get_val(row, hdr)
                    try:
                        return int(float(v)) if v is not None else default
                    except (TypeError, ValueError):
                        return default
                pile_sec_param_dict[name] = {
                    'type': '锁扣钢管桩',
                    '截面号': seq,
                    '截面名称': name,
                    '材料标号': material,
                    'params': {
                        'D': _f_val('D(mm)'),
                        't': _f_val('t(mm)'),
                        '锁扣宽度': _f_val('锁扣宽度(mm)', 35),
                        '钢板桩数量': _i_val('钢板桩数量', 1),
                        '钢板桩宽度': _f_val('钢板桩宽度(mm)', 600),
                    }
                }
            else:
                # 钢板桩
                params = {}
                for hdr in sec_keylst:
                    if hdr and hdr not in ('截面号', '截面名称', '材料标号'):
                        v = _sec_get_val(row, hdr)
                        if v is not None and str(v).strip() not in ('', '/'):
                            try:
                                params[hdr] = float(v)
                            except (TypeError, ValueError):
                                pass
                pile_sec_param_dict[name] = {
                    'type': '钢板桩',
                    '截面号': seq,
                    '截面名称': name,
                    '材料标号': material,
                    'params': params,
                }
    for k, v in pile_sec_param_dict.items():
        print(k)
        print(v)
    print('='*80 + '\n')
    
    # 获取连接形式参数表中的数据（连接形式 sheet）
    elink_param_lst = []
    conn_data_raw = read_excel_to_dict_by_Sheet(WorkBook, '连接形式')
    if conn_data_raw and len(conn_data_raw) > 2:
        # R1=大分类, R2=表头, R3起=数据
        conn_keylst = conn_data_raw[1]  # R2=表头
        conn_valuelst = conn_data_raw[2:]  # R3起=数据
        for row in conn_valuelst:
            if not row or not row[0]:
                continue
            name = str(row[0]).strip()
            if not name:
                continue
            def _get_conn_val(hdr):
                try:
                    idx = conn_keylst.index(hdr)
                    v = row[idx] if idx < len(row) else None
                    return str(v).strip() if v is not None else ''
                except (ValueError, IndexError):
                    return ''
            elink_param_lst.append({
                'name': name,
                'SDx': _get_conn_val('SDx(N/mm)'),
                'SDy': _get_conn_val('SDy(N/mm)'),
                'SDz': _get_conn_val('SDz(N/mm)'),
                'SRx': _get_conn_val('SRx(N·mm/rad)'),
                'SRy': _get_conn_val('SRy(N·mm/rad)'),
                'SRz': _get_conn_val('SRz(N·mm/rad)'),
                'NSDx': _get_conn_val('NSDx(N/mm)'),
            })
    for lst in elink_param_lst:
        print(lst)
    print('='*80 + '\n')

    # 关闭Excel
    WorkBook.Close(False)
    print('WorkBook已关闭')


    return {
        'main_param_dict': main_param_dict,
        'soil_param_dict': soil_param_dict,
        'pile_sec_param_dict': pile_sec_param_dict,
        'elink_param_lst': elink_param_lst,
    }, proj_index_dict


def On_Submit(excel_path, save_path, mcb_checked, csv_checked, root, status_var, status_label, progress_bar):
    # 重置进度条
    progress_bar['value'] = 0
    progress_bar.configure(bootstyle=SUCCESS + STRIPED)
    status_var.set("正在初始化...")
    status_label.configure(bootstyle=INFO)
    root.update_idletasks()

    """批量运行入口：解析 Excel → 逐墩号调用 Generate_MCT 生成 mct 文件"""
    data, proj_index_dict = Batch_CofferDam_Excel_Data_Process(excel_path)
    main_param_dict = data['main_param_dict']
    soil_param_dict = data['soil_param_dict']
    pile_sec_param_dict = data['pile_sec_param_dict']
    elink_param_lst = data['elink_param_lst']
    print(proj_index_dict)

    print(f'\n{"="*80}')
    print(f'数据汇总:')
    print(f'  主表参数: {len(main_param_dict)} 个墩号, 名称: {list(main_param_dict.keys())}')
    print(f'  土层数据: {len(soil_param_dict)} 个sheet, 名称: {list(soil_param_dict.keys())}')
    print(f'  支护桩截面: {len(pile_sec_param_dict)} 个, 名称: {list(pile_sec_param_dict.keys())}')
    print(f'  连接形式: {len(elink_param_lst)} 条')
    print(f'{"="*80}\n')

    # 定义任务清单
    total_steps = len(list(main_param_dict.keys()))
    current_step = 0
    def update_p(desc):
        nonlocal current_step
        current_step += 1
        if progress_bar and status_var and root:
            progress_bar['value'] = (current_step / total_steps) * 100
            status_var.set(desc)
            root.update()

    success = 0
    fail    = 0
    success_lst = []
    fail_lst    = []

    for k, v in main_param_dict.items():
        # 墩号化为 int（无法化 int 的保持字符串）
        pier_id = k
        # print(f'\n===== 处理墩号: {pier_id} =====')
        update_p(f'正在处理墩号:{pier_id}')

        # 从参数字典中提取 Generate 函数所需的全部变量
        # 调用本地 Generate_MCT 函数（参数检查打印在函数内部）
        try:
            pv = prepare_submit_variables(pier_id, v, proj_index_dict, soil_param_dict, pile_sec_param_dict, elink_param_lst, save_path, excel_path)
            Generate_MCT(
                pv['Cap_X'], pv['Cap_Y'], pv['Cap_H'], pv['X_offset'], pv['Y_offset'], pv['Spring_Thickness'],
                pv['Sheet_Pile_typevar'], pv['Sheet_Pile_materialvar'], pv['CofferDam_L'], pv['Sheet_Pile_namevar'],
                pv['Pile_SEC_info_dict'], pv['SKGGZ_D'], pv['SKGGZ_t'], pv['SKGGZ_Gap'],
                pv['Waler_rows'], pv['Strut_blocks'],
                pv['corbel_long'], pv['corbel_short'],
                pv['Solid_Level'], pv['Water_Level'], pv['Cap_Bottom_Level'], pv['CofferDam_Top_Level'],
                pv['Concrete_Blinding_check'], pv['Concrete_Blinding_thickness'],
                pv['Concrete_Plug_check'], pv['Concrete_Plug_thickness'],
                pv['Drawdown_height'], pv['Waterdown_height'], pv['Excavation_face_dewater'],
                pv['Waler_section_dict'], pv['Strut_section_dict'],
                pv['recognize_Cap_result_dict'], pv['recognize_Bracket_result_dict'],
                pv['recognize_Waler_result_dict'], pv['recognize_Strut_result_dict'],
                pv['recognize_Strut_Replace_result_dict'],
                pv['Steel_Sheet_Pile_CofferDam_Load_dict'],
                pv['sigma_k_dict'],
                pv['Pile_Bottom_Boundary'],
                pv['cofferdam_elasticlink_form_dict'],
                pv['sheet_elasticlink_dict'],
                pv['none_construction_stage_check'], pv['construction_stage_check'],
                pv['Load_Combo_Line_Data'], pv['Construct_Stage_Line_Data'],
                pv['stage_dict'], pv['if_consider_solid_stress_path'],
                pv['mct_savepath'], mcb_checked, csv_checked,
                pv['assist_add_dict'],
                pv['assist_replace_dict'],
                pv['steel_dict'],
                pv['steel_std'],
            )
            print(f'墩号 {pier_id} 生成完成')
            success = success + 1
            success_lst.append(pier_id)
        except Exception as e:
            print(e)
            fail = fail + 1
            fail_lst.append(pier_id)

    progress_bar.configure(bootstyle=SUCCESS)
    status_var.set("批量生成完成！")
    status_label.configure(bootstyle=SUCCESS)

    print(f'生成成功:{success}个')
    for x in success_lst:
        print(x)
    print(f'生成失败:{fail}个')
    for x in fail_lst:
        print(x)

#=====================================================================================================================================================================
# 含主对话框全局变量的函数
# 对话框总函数
def creat_dialog(parent=None, on_close=None):
    # root
    if parent is None:
        root = tb.Window("矩形围堰批量建模exe")
    else:
        root = tb.Toplevel(parent)
        root.title("矩形围堰批量建模exe")

    # 对话框变量设置
    Set_File_entry_var1 = tb.StringVar(value = "未选取任何文件") # 选取模型
    Set_File_entry_var2 = tb.StringVar(value = "") # 另存为路径
    # 创建一个 BooleanVar 用于存储复选框状态（True/False）
    mcb_checked = tb.BooleanVar(value=True)   # 默认勾选
    csv_checked = tb.BooleanVar(value=False)  # 默认不勾选

    # 标题：文件操作
    dialog_headr(root, "文件操作")
    # 文件读取、结果保存 
    dialog_boxline(root, "文件读取", 7, Set_File_entry_var1, 60, "选择Excel批量表", lambda: Choose_Excel_File(Set_File_entry_var1, Set_File_entry_var2))
    dialog_boxline(root, "结果保存", 7, Set_File_entry_var2, 60, "结果文件另存为",  lambda: MCT_File_SaveAs  (Set_File_entry_var2))

    # --- 进度可视化区域 ---
    progress_frame = tb.Frame(root, padding=(20, 10))
    progress_frame.pack(fill=X, expand=YES)

    status_var = tb.StringVar(value="请完善上方文件配置...")
    status_label = tb.Label(progress_frame, textvariable=status_var, bootstyle=INFO)
    status_label.pack(side=TOP, anchor=W)

    progress_bar = tb.Progressbar(
        progress_frame,
        bootstyle=SUCCESS + STRIPED,
        maximum=100,
        value=0
    )
    progress_bar.pack(fill=X, pady=5)

    # 运行按钮调用
    dialog_boxline_end(root, mcb_checked, "运行", lambda: On_Submit(Set_File_entry_var1.get(), Set_File_entry_var2.get(), mcb_checked.get(), csv_checked.get(), root, status_var, status_label, progress_bar), "关闭", lambda: on_exit(root))

    # --- 状态校验 ---
    def check_ready(*args):
        excel = Set_File_entry_var1.get()
        if "xlsx" in excel:
            status_var.set("准备就绪")
            status_label.configure(bootstyle=SUCCESS)
        else:
            progress_bar['value'] = 0
            status_var.set("请完善上方文件配置...")
            status_label.configure(bootstyle=INFO)

    Set_File_entry_var1.trace_add("write", check_ready)
    check_ready()

    if parent is None:
        root.mainloop()
    elif on_close:
        def _on_root_destroy(event, _root=root):
            if event.widget is _root:
                on_close()
        root.bind('<Destroy>', _on_root_destroy)


# def Generate_MCT(
#         Cap_X, Cap_Y, X_offset, Y_offset, Spring_Thickness,
#         Sheet_Pile_typevar, Sheet_Pile_materialvar, CofferDam_L, Sheet_Pile_namevar,
#         Pile_SEC_info_dict, SKGGZ_D, SKGGZ_t, SKGGZ_Gap,
#         Waler_rows, Strut_blocks,
#         corbel_long_num, corbel_short_num,
#         Solid_Level, Water_Level, Cap_Bottom_Level, CofferDam_Top_Level,
#         Concrete_Blinding_check_var, Concrete_Blinding_thickness_var,
#         Concrete_Plug_check_var, Concrete_Plug_thickness_var,
#         Drawdown_height, Waterdown_height, Excavation_face_dewater,
#         Waler_section_dict, Strut_section_dict,
#         recognize_Cap_result_dict, recognize_Bracket_result_dict,
#         recognize_Waler_result_dict, recognize_Strut_result_dict,
#         recognize_Strut_Replace_result_dict,
#         Steel_Sheet_Pile_CofferDam_Load_dict,
#         sigma_k_dict,
#         Pile_Bottom_Boundary,
#         cofferdam_elasticlink_form_dict,
#         sheet_elasticlink_dict,
#         none_construction_stage_check_var, construction_stage_check_var,
#         Load_Combo_Line_Data, Construct_Stage_Line_Data,
#         stage_dict, if_consider_solid_stress_path,
#         mct_savepath, if_mcb,
#         assist_add_dict=None,
#         assist_replace_dict=None,
#         steel_dict=None,
#         steel_std=None,
#     ):
#     # ===== 参数检查打印 =====
#     print('=' * 60)
#     print('Generate_MCT 参数检查')
#     print('=' * 60)

#     print('--- 基本几何 ---')
#     print(f"  Cap_X={Cap_X} (type={type(Cap_X).__name__}), Cap_Y={Cap_Y} (type={type(Cap_Y).__name__})")
#     print(f"  X_offset={X_offset} (type={type(X_offset).__name__}), Y_offset={Y_offset} (type={type(Y_offset).__name__})")
#     print(f"  Spring_Thickness={Spring_Thickness} (type={type(Spring_Thickness).__name__})")

#     print('--- 支护桩 ---')
#     print(f"  Sheet_Pile_typevar={Sheet_Pile_typevar} (type={type(Sheet_Pile_typevar).__name__}), Sheet_Pile_materialvar={Sheet_Pile_materialvar} (type={type(Sheet_Pile_materialvar).__name__})")
#     print(f"  CofferDam_L={CofferDam_L} (type={type(CofferDam_L).__name__}), Sheet_Pile_namevar={Sheet_Pile_namevar} (type={type(Sheet_Pile_namevar).__name__})")
#     print(f"  Pile_SEC_info_dict={Pile_SEC_info_dict} (type={type(Pile_SEC_info_dict).__name__})")
#     print(f"  SKGGZ_D={SKGGZ_D} (type={type(SKGGZ_D).__name__}), SKGGZ_t={SKGGZ_t} (type={type(SKGGZ_t).__name__}), SKGGZ_Gap={SKGGZ_Gap} (type={type(SKGGZ_Gap).__name__})")

#     print('--- 标高 ---')
#     print(f"  Solid_Level={Solid_Level} (type={type(Solid_Level).__name__}), Water_Level={Water_Level} (type={type(Water_Level).__name__})")
#     print(f"  Cap_Bottom_Level={Cap_Bottom_Level} (type={type(Cap_Bottom_Level).__name__}), CofferDam_Top_Level={CofferDam_Top_Level} (type={type(CofferDam_Top_Level).__name__})")

#     print('--- 垫层/封底 ---')
#     print(f"  Concrete_Blinding_check_var={Concrete_Blinding_check_var} (type={type(Concrete_Blinding_check_var).__name__}), Concrete_Blinding_thickness_var={Concrete_Blinding_thickness_var} (type={type(Concrete_Blinding_thickness_var).__name__})")
#     print(f"  Concrete_Plug_check_var={Concrete_Plug_check_var} (type={type(Concrete_Plug_check_var).__name__}), Concrete_Plug_thickness_var={Concrete_Plug_thickness_var} (type={type(Concrete_Plug_thickness_var).__name__})")

#     print('--- 牛腿 ---')
#     print(f"  corbel_long_num={corbel_long_num} (type={type(corbel_long_num).__name__}), corbel_short_num={corbel_short_num} (type={type(corbel_short_num).__name__})")

#     print('--- 超挖/降水 ---')
#     print(f"  Drawdown_height={Drawdown_height} (type={type(Drawdown_height).__name__}), Waterdown_height={Waterdown_height} (type={type(Waterdown_height).__name__}), Excavation_face_dewater={Excavation_face_dewater} (type={type(Excavation_face_dewater).__name__})")

#     print('--- Waler_rows (list) ---')
#     print(f"  type={type(Waler_rows).__name__}, len={len(Waler_rows) if Waler_rows else 0}")
#     if Waler_rows:
#         for i, wr in enumerate(Waler_rows):
#             print(f"  [{i}] type={type(wr).__name__}")
#             if isinstance(wr, dict):
#                 for k, val in wr.items():
#                     v_get = val.get() if hasattr(val, 'get') else val
#                     print(f"       {k}={v_get} (val_type={type(val).__name__})")

#     print('--- Strut_blocks (dict) ---')
#     print(f"  type={type(Strut_blocks).__name__}, keys={list(Strut_blocks.keys()) if Strut_blocks else '[]'}")
#     if Strut_blocks:
#         for i, sb in Strut_blocks.items():
#             print(f"  [{i}] type={type(sb).__name__}")
#             if isinstance(sb, dict):
#                 for k, val in sb.items():
#                     v_get = val.get() if hasattr(val, 'get') else val
#                     print(f"       {k}={v_get} (val_type={type(val).__name__})")

#     print('--- 围檩/内支撑截面字典 ---')
#     print(f"  Waler_section_dict: type={type(Waler_section_dict).__name__}, keys={list(Waler_section_dict.keys())[:5]}... (共{len(Waler_section_dict)}项)")
#     print(f"  Strut_section_dict: type={type(Strut_section_dict).__name__}, keys={list(Strut_section_dict.keys())[:5]}... (共{len(Strut_section_dict)}项)")

#     # print('--- 计算结果字典 ---')
#     # print(f"  recognize_Cap_result_dict: type={type(recognize_Cap_result_dict).__name__}, len={len(recognize_Cap_result_dict)}")
#     # print(f"  recognize_Bracket_result_dict: type={type(recognize_Bracket_result_dict).__name__}, len={len(recognize_Bracket_result_dict)}")
#     # print(f"  recognize_Waler_result_dict: type={type(recognize_Waler_result_dict).__name__}, len={len(recognize_Waler_result_dict)}")
#     # print(f"  recognize_Strut_result_dict: type={type(recognize_Strut_result_dict).__name__}, len={len(recognize_Strut_result_dict)}")
#     # print(f"  recognize_Strut_Replace_result_dict: type={type(recognize_Strut_Replace_result_dict).__name__}, len={len(recognize_Strut_Replace_result_dict)}")

#     print('--- 土层荷载字典 ---')
#     print(f"  type={type(Steel_Sheet_Pile_CofferDam_Load_dict).__name__}")
#     if isinstance(Steel_Sheet_Pile_CofferDam_Load_dict, dict):
#         for k, v in Steel_Sheet_Pile_CofferDam_Load_dict.items():
#             if k == 'Excel':
#                 print(f"  Excel (dict, {len(v)}项):")
#                 if isinstance(v, dict):
#                     for lk, lv in v.items():
#                         print(f"    {lk}: {lv}")
#             else:
#                 print(f"  {k}={v}")

#     print('--- 超载/荷载参数 ---')
#     print(f"  sigma_k_dict: type={type(sigma_k_dict).__name__}, len={len(sigma_k_dict)}")
#     if isinstance(sigma_k_dict, dict):
#         for k, v in sigma_k_dict.items():
#             print(f"    {k}: {v}")

#     print('--- 边界/连接 ---')
#     print(f"  Pile_Bottom_Boundary={Pile_Bottom_Boundary} (type={type(Pile_Bottom_Boundary).__name__})")
#     print(f"  cofferdam_elasticlink_form_dict: type={type(cofferdam_elasticlink_form_dict).__name__}, len={len(cofferdam_elasticlink_form_dict)}")
#     if isinstance(cofferdam_elasticlink_form_dict, dict):
#         for k, v in cofferdam_elasticlink_form_dict.items():
#             print(f"    {k}: {v}")
#     print(f"  sheet_elasticlink_dict: type={type(sheet_elasticlink_dict).__name__}, len={len(sheet_elasticlink_dict)}")
#     if isinstance(sheet_elasticlink_dict, dict):
#         for k, v in sheet_elasticlink_dict.items():
#             print(f"    {k}: {v}")

#     print('--- 工况 ---')
#     print(f"  none_construction_stage_check_var={none_construction_stage_check_var} (type={type(none_construction_stage_check_var).__name__})")
#     print(f"  construction_stage_check_var={construction_stage_check_var} (type={type(construction_stage_check_var).__name__})")
#     print(f"  Load_Combo_Line_Data: type={type(Load_Combo_Line_Data).__name__}, len={len(Load_Combo_Line_Data)}")
#     if isinstance(Load_Combo_Line_Data, list):
#         for row in Load_Combo_Line_Data:
#             if any(str(x).strip() for x in row):
#                 print(f"    {row}")
#     print(f"  Construct_Stage_Line_Data: type={type(Construct_Stage_Line_Data).__name__}, len={len(Construct_Stage_Line_Data)}")
#     if isinstance(Construct_Stage_Line_Data, list):
#         for row in Construct_Stage_Line_Data:
#             print(f"    {row}")
#     print(f"  stage_dict: type={type(stage_dict).__name__}, keys={list(stage_dict.keys()) if isinstance(stage_dict, dict) else 'N/A'}")
#     if isinstance(stage_dict, dict):
#         for k, v in stage_dict.items():
#             print(f"    {k}: {v}")

#     print('--- 辅助加撑/换撑 ---')
#     print(f"  assist_add_dict: type={type(assist_add_dict).__name__}, len={len(assist_add_dict) if assist_add_dict else 0}")
#     if assist_add_dict:
#         for k, v in assist_add_dict.items():
#             print(f"    {k}: {v}")
#     print(f"  assist_replace_dict: type={type(assist_replace_dict).__name__}, len={len(assist_replace_dict) if assist_replace_dict else 0}")
#     if assist_replace_dict:
#         for k, v in assist_replace_dict.items():
#             print(f"    {k}: {v}")

#     print('--- 材质库/规范 ---')
#     print(f"  steel_dict: type={type(steel_dict).__name__}, keys={list(steel_dict.keys()) if steel_dict else []}")
#     if isinstance(steel_dict, dict):
#         for k, v in steel_dict.items():
#             print(f"    {k}: {list(v.keys()) if isinstance(v, dict) else v}")
#     print(f"  steel_std={steel_std} (type={type(steel_std).__name__})")

#     print('--- 其他 ---')
#     print(f"  if_consider_solid_stress_path={if_consider_solid_stress_path} (type={type(if_consider_solid_stress_path).__name__})")
#     print(f"  mct_savepath={mct_savepath} (type={type(mct_savepath).__name__})")
#     print(f"  if_mcb={if_mcb} (type={type(if_mcb).__name__})")
#     print('=' * 60)
#     print('参数检查完毕')
#     print('=' * 60)

#=====================================================================================================================================================================

def CofferDam_FEM_MidasCivil_Main_Batch(parent=None, on_close=None):
    if not if_Reg(parent):
        if on_close:
            on_close()
        return

    creat_dialog(parent, on_close)


# CofferDam_FEM_MidasCivil_Main_Batch()

