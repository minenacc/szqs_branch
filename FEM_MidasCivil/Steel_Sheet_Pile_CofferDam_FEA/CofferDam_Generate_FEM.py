# =============================================================================
# 术语表（全文件统一约定）
#   SKGGZ    : 锁扣钢管桩
#   Waler    : 围檩（水平向围护梁）
#   Strut    : 内支撑（对撑 DuiCheng / 斜撑 XieCheng）
#   Cap      : 承台；Bracket: 牛腿
#   MCT      : MIDAS Civil 命令文本；MCB: MIDAS 模型文件
#   工况类型 : 取土/抽水/加撑/拆撑/封底/垫层/辅助加撑/辅助换撑/辅助加圈梁
#   Ka / Ks / Pa / Ps0 : 主动土压力系数 / 土弹簧刚度 / 主动土压力 / 初始土反力
#   单位约定 : 坐标与截面尺寸 mm；标高、层厚 m；应力 kPa；刚度 N/mm；
#              MCT 文件单位系为 "N,MM,KJ,C"
# =============================================================================

# 1. 标准库
import re
import math
import copy
from tkinter import messagebox
from itertools import accumulate

# 3. 本地模块
from General.ExcelHandle import write_csv
from General.FilePath    import write_file_within_path, file_extension_Modified
from General.Midas       import MidasURL, importmct_to_mcb, group_arithmetic_sequences, midas_stage
from General.Geometry    import non_negative_elem_load, calculate_full_angle, find_circle_line_intersection, point_to_line_projection
from General.DataUtils   import SET_CLIP_STRING, Divide_Solid_within_LayerThickness, intersect_segment_with_intervals, elem_load_lst_trans, add_continuous_beam_elem_loads
from General.Formula     import cal_m, cal_Ks, cal_Ka, cal_sigmak_solid, cal_sigmak_water, cal_pak_soild, cal_pak_water, cal_ps0_soild, Localized_Rectangular_Load, Localized_Strip_Load

from FEM_MidasCivil.Steel_Sheet_Pile_CofferDam_FEA.CofferDam_Calculate_FEM  import Cap_Pile_PM, Waler_PM, Strut_PM, Bracket_PM


def Steel_Sheet_Pile_CofferDam_on_submit_MCT(
        Cap_X, Cap_Y, Cap_H, X_offset, Y_offset, Spring_Thickness,
        Sheet_Pile_typevar, Sheet_Pile_materialvar, CofferDam_L, Sheet_Pile_namevar,
        Pile_SEC_info_dict, SKGGZ_D, SKGGZ_t, SKGGZ_Gap,
        Waler_rows, Strut_blocks,
        corbel_long_num, corbel_short_num,
        Solid_Level, Water_Level, Cap_Bottom_Level, CofferDam_Top_Level,
        Concrete_Blinding_check_var, Concrete_Blinding_thickness_var,
        Concrete_Plug_check_var, Concrete_Plug_thickness_var,
        Drawdown_height, Waterdown_height, Excavation_face_dewater,
        Waler_section_dict, Strut_section_dict,
        recognize_Cap_result_dict, recognize_Bracket_result_dict,
        recognize_Waler_result_dict, recognize_Strut_result_dict,
        recognize_Strut_Replace_result_dict,
        Steel_Sheet_Pile_CofferDam_Load_dict,
        sigma_k_dict,
        Pile_Bottom_Boundary,
        cofferdam_elasticlink_form_dict,
        sheet_elasticlink_dict,
        none_construction_stage_check_var, construction_stage_check_var,
        Load_Combo_Line_Data, Construct_Stage_Line_Data,
        stage_dict, if_consider_solid_stress_path,
        mct_savepath, if_mcb, if_csv,
        assist_add_dict=None,
        assist_replace_dict=None,
        steel_dict=None,
        steel_std=None,
    ):
    """生成围堰 MCT/MCB/CSV 的总入口：整合参数 -> 计算几何与荷载 -> 输出模型。

    负责把 UI 传入的原始参数整理成内部字典（材料、围檩/内支撑选型），
    依次调用 Cap_Pile_PM、Waler_PM、Strut_PM、Bracket_PM 计算几何，
    再调用 mct_output 完成荷载/边界/施工阶段与文件输出。

    Args:
        Cap_X, Cap_Y, Cap_H (float): 承台长/宽/高（m）。
        X_offset, Y_offset (float): 承台长/宽方向预留边距（m）。
        Spring_Thickness (float): 土弹簧分层厚度（m）。
        Sheet_Pile_typevar (str): 支护类型，'钢板桩' 或 '锁扣钢管桩'。
        Sheet_Pile_materialvar (str): 支护桩材质牌号。
        CofferDam_L (float): 围堰高（m）。
        Sheet_Pile_namevar (str): 支护桩截面名称。
        Pile_SEC_info_dict (dict): 支护桩截面信息 {'SEC': list}。
        SKGGZ_D, SKGGZ_t, SKGGZ_Gap (float): 锁扣钢管桩直径/壁厚/中心间距（mm）。
        Waler_rows (list[dict]): 各层围檩 UI 控件行，含 entry/combo/mat_* 等控件。
        Strut_blocks (dict): 内支撑布置输入控件 {层号: {'DC_X': tkvar, ...}}。
        corbel_long_num, corbel_short_num: 长/短边牛腿数量。
        Solid_Level, Water_Level (float): 初始地面/水位标高（m）。
        Cap_Bottom_Level, CofferDam_Top_Level (float): 承台底/围堰顶标高（m）。
        Concrete_Blinding_check_var, Concrete_Blinding_thickness_var: 垫层开关/厚度（m）。
        Concrete_Plug_check_var, Concrete_Plug_thickness_var: 封底开关/厚度（m）。
        Drawdown_height, Waterdown_height, Excavation_face_dewater (float): 超挖深度、
            基坑底降水、开挖面降水（m）。
        Waler_section_dict, Strut_section_dict (dict): 围檩/内支撑截面库。
        recognize_Cap_result_dict 等 recognize_* (dict): 几何计算结果输出字典。
        Steel_Sheet_Pile_CofferDam_Load_dict (dict): 土层信息，结构形式
            {'Excel': {土层名称: {'层厚':..., '重度':..., '黏聚力':..., '内摩擦角':..., '计算方法':...}}}。
        sigma_k_dict (dict): 附加荷载信息 {'均布附加荷载':..., '矩形局部附加荷载':..., '条形局部附加荷载':...}。
        Pile_Bottom_Boundary (str): 桩底边界编码。
        cofferdam_elasticlink_form_dict (dict): 各连接位置的形式（如"共节点"）。
        sheet_elasticlink_dict (dict): 各连接的具体刚度。
        none_construction_stage_check_var, construction_stage_check_var: 整体/施工阶段模式开关。
        Load_Combo_Line_Data (list): 荷载组合表数据。
        Construct_Stage_Line_Data (list[list]): 施工阶段表数据。
        stage_dict (dict): 施工阶段字典。
        if_consider_solid_stress_path (bool): 是否考虑土的应力路径。
        mct_savepath (str): MCT 保存路径。
        if_mcb, if_csv (bool): 是否同时生成 MCB/CSV。
        assist_add_dict, assist_replace_dict (dict|None): 辅助加撑/换撑定义。
        steel_dict (dict|None), steel_std (str|None): 钢材材料库与所选规范。

    Returns:
        None: 副作用为计算并写出 MCT/MCB/CSV 文件，并把命令流写入剪贴板。
    """
    print('')
    print('土层信息检查')
    for k, v in Steel_Sheet_Pile_CofferDam_Load_dict['Excel'].items():
        print(k)
        print(v)
    print('超载信息检查')
    for k, v in sigma_k_dict.items():
        print(k)
        print(v)
    print('连接形式定义检查')
    for k, v in cofferdam_elasticlink_form_dict.items():
        print(f"{k}: {v}")
    print('连接具体刚度检查')
    for k, v in sheet_elasticlink_dict.items():
        print(f"{k}: {v}")
    print('支护桩截面参数检查')
    print(Pile_SEC_info_dict)
    print('支护桩材质检查')
    print(Sheet_Pile_materialvar)
    print('工况定义信息检查')
    if construction_stage_check_var:
        print(f"stage_dict:")
        for stage_key in sorted(stage_dict.keys()):
            print(f"{stage_key}: {stage_dict[stage_key]}")
    else:
        print(f"Load_Combo_Line_Data = {Load_Combo_Line_Data}")
    print('施工阶段定义检查')
    for x in Construct_Stage_Line_Data:
        print(x)
    print('辅助加撑数据检查')
    for k, v in assist_add_dict.items():
        print(k, v)
    print('辅助换撑数据检查')
    for k, v in assist_replace_dict.items():
        print(k, v)

    Steel_Sheet_Pile_CofferDam_Load_dict['Excel'] = {
        k: v for k, v in Steel_Sheet_Pile_CofferDam_Load_dict['Excel'].items()
        if v.get('土层名称') != ''}

    Strut_blocks_dict = {}
    strut_i = 0
    Walers_Strut_selected_dict = {}

    if Waler_rows:
        for key, value in Strut_blocks.items():
            if strut_i < len(Waler_rows):
                Strut_blocks_dict[key] = {}
                for k, v in value.items():
                    Strut_blocks_dict[key][k] = '/' if v.get() == '' else v.get()
            strut_i += 1

        Walers_Strut_selected_dict = {
            f'{i+1}行': {
                '间距': float(Waler_rows[i]['entry'].get()),
                '围檩长边截面': Waler_rows[i]['combo_long'].get(),
                '围檩宽边截面': Waler_rows[i]['combo_short'].get(),
                '围檩长边材质': Waler_rows[i]['mat_long'].get(),
                '围檩宽边材质': Waler_rows[i]['mat_short'].get(),
                '对撑截面': Waler_rows[i]['combo2'].get(),
                '斜撑截面': Waler_rows[i]['combo3'].get(),
                '对撑材质': Waler_rows[i]['mat_dc'].get(),
                '斜撑材质': Waler_rows[i]['mat_xc'].get(),
                'X对撑布置': Strut_blocks[i]['DC_X'].get(),
                'Y对撑布置': Strut_blocks[i]['DC_Y'].get(),
                'X斜撑布置': Strut_blocks[i]['XC_X'].get(),
                'Y斜撑布置': Strut_blocks[i]['XC_Y'].get(),
            }
            for i in range(len(Waler_rows))
        }

        print('围檩及内支撑材质检查')
        for k, v in Walers_Strut_selected_dict.items():
            print(k, '围檩长边材质:', v.get('围檩长边材质'),
                  '围檩宽边材质:', v.get('围檩宽边材质'),
                  '对撑材质:', v.get('对撑材质'),
                  '斜撑材质:', v.get('斜撑材质'))

    # 材料定义（支护桩 + 围檩 + 内支撑 + 辅助加撑 + 辅助换撑，仅钢材）
    material_dict = collect_steel_material_dict(
        Sheet_Pile_typevar, Sheet_Pile_materialvar, Walers_Strut_selected_dict,
        assist_add_dict, assist_replace_dict, steel_dict, steel_std)
    print('材料定义检查')
    for k, v in material_dict.items():
        print(k, v)

    # 水位标准化：对于无水围堰，将水位等效为桩底标高-1m
    Water_Level = round(float(Water_Level), 3) if Water_Level != '/' else round(float(CofferDam_Top_Level) - float(CofferDam_L) - 1, 3)
    # 计算围堰大小及支护桩位置
    Cap_Pile_PM(Cap_X, Cap_Y, X_offset, Y_offset, Sheet_Pile_typevar,
                Pile_SEC_info_dict, recognize_Cap_result_dict, SKGGZ_D, SKGGZ_Gap)
    # 计算围檩
    Waler_PM(CofferDam_Top_Level, Pile_SEC_info_dict, recognize_Cap_result_dict,
             recognize_Waler_result_dict, Walers_Strut_selected_dict, assist_add_dict,
             Waler_section_dict, Sheet_Pile_typevar, SKGGZ_D)
    # 计算内支撑
    Strut_PM(Strut_blocks_dict, assist_add_dict, assist_replace_dict, stage_dict, recognize_Waler_result_dict, recognize_Strut_result_dict, recognize_Strut_Replace_result_dict)
    # 计算牛腿
    Bracket_PM(recognize_Cap_result_dict, recognize_Bracket_result_dict, corbel_long_num, corbel_short_num)
    # 桩底约束
    cofferdam_cons = Pile_Bottom_Boundary
    # 生成mcb
    mct_output(
        Cap_X, Cap_Y, Cap_H, Spring_Thickness,
        Sheet_Pile_typevar, Sheet_Pile_materialvar, CofferDam_L, Sheet_Pile_namevar, Pile_SEC_info_dict,
        SKGGZ_D, SKGGZ_t, SKGGZ_Gap,
        Walers_Strut_selected_dict, assist_add_dict, assist_replace_dict,
        Solid_Level, Water_Level, Cap_Bottom_Level, CofferDam_Top_Level,
        Concrete_Blinding_check_var, Concrete_Blinding_thickness_var,
        Concrete_Plug_check_var, Concrete_Plug_thickness_var,
        Drawdown_height, Waterdown_height, Excavation_face_dewater,
        Waler_section_dict, Strut_section_dict,
        recognize_Cap_result_dict, recognize_Bracket_result_dict,
        recognize_Waler_result_dict, recognize_Strut_result_dict,
        recognize_Strut_Replace_result_dict,
        Steel_Sheet_Pile_CofferDam_Load_dict,
        sigma_k_dict, sheet_elasticlink_dict, cofferdam_elasticlink_form_dict,
        cofferdam_cons, none_construction_stage_check_var, construction_stage_check_var,
        Load_Combo_Line_Data, Construct_Stage_Line_Data,
        stage_dict, if_consider_solid_stress_path,
        mct_savepath, if_mcb, if_csv,
        material_dict,
    )


# 仅受压方向在mct文件中的形式
def SPRING_COMP_Direction(angle):
    """把"仅受压"方向角映射为 MIDAS 节点弹性支承的方向编码。

    Args:
        angle (int): 单元旋转角度，取值 90/180/270/360。

    Returns:
        int: MIDAS 方向编码，90->3(Dy-)、180->0(Dx+)、270->2(Dy+)、360->1(Dx-)。
    """
    Direction_Dict = {
        90 :3, # Dy(-)
        180:0, # Dx(+)
        270:2, # Dy(+)
        360:1, # Dx(-)
    }
    Direction = Direction_Dict[angle]
    return Direction


def mct_material_midas(material_dict):
    """将材料字典转换为 MIDAS MCT 的 *MATERIAL 命令行列表。

    Args:
        material_dict (dict): 结构形式
            {编号(int): {'type': 'CONC'/'STEEL', 'name': str, 'spheat': str,
                        'heatco': str, 'dampratio': str, 'standrad': str,
                        'elast': str}}。
            传入弹性模量 elast 的单位为 N/mm。

    Returns:
        list[str]: 每行一条 MCT 材料命令，顺序与 material_dict 的编号一致。

    Example:
        >>> MATERIAL({1: {'type': 'STEEL', 'name': 'Q235', 'spheat': '0.6',
        ...               'heatco': '0', 'dampratio': '0.02',
        ...               'standrad': 'GB50017-17(S)', 'elast': '206000'}})
        ['1, STEEL, Q235, 0.6, 0, , C, NO, 0.02, 1, GB50017-17(S), , Q235, NO, 206000']
    """
    MATERIAL_lst = []
    for k, v in material_dict.items():
        mat_index = k
        mat_type = v['type']
        mat_name = v['name']
        mat_spheat = v['spheat']
        mat_heatco = v['heatco']
        mat_dampratio = v['dampratio']
        mat_standrad = v['standrad']
        mat_elast = v['elast']
        MATERIAL_lst.append(
            '{}, {}, {}, {}, {}, , C, NO, {}, 1, {}, , {}, NO, {}'.format(mat_index, mat_type, mat_name, mat_spheat, mat_heatco, mat_dampratio, mat_standrad, mat_name, mat_elast)
        )
    return MATERIAL_lst


def collect_steel_material_dict(Sheet_Pile_typevar, Sheet_Pile_materialvar, Walers_Strut_selected_dict,
                                assist_add_dict, assist_replace_dict,
                                steel_dict=None, steel_std=None):
    """收集所有钢材材质（支护桩+围檩+内支撑+辅助加撑+辅助换撑），生成 material_dict。

    去重所有出现过的材质牌号，按需从钢材材料库取弹性模量与规范编号；
    钢板桩固定补充一个 Q295 材质并置于首位。

    Args:
        Sheet_Pile_typevar (str): 支护类型，'钢板桩' 时强制加入 Q295。
        Sheet_Pile_materialvar (str): 支护桩材质牌号。
        Walers_Strut_selected_dict (dict): 围檩/内支撑选型表（含各材质字段）。
        assist_add_dict (dict): 辅助加撑定义。
        assist_replace_dict (dict): 辅助换撑定义。
        steel_dict (dict|None): 钢材材料库，见 Excel_io.load_material_library。
        steel_std (str|None): 选用的钢结构规范名，用于取规范编号与牌号参数。

    Returns:
        dict: 结构形式
            {编号(int): {'type': str, 'name': str, 'spheat': str, 'heatco': str,
                        'dampratio': str, 'standrad': str, 'elast': str}}。
    """
    names = set()

    def _add(m):
        """将非空、非 "/" 的材质名加入 names 集合。

        Args:
            m (str|None): 材质名称。

        Returns:
            None
        """
        if m is None:
            return
        m = str(m).strip()
        if m and m != '/':
            names.add(m)

    # 支护桩
    _add(Sheet_Pile_materialvar)
    # 围檩 + 内支撑
    for v in (Walers_Strut_selected_dict or {}).values():
        for key in ('围檩长边材质', '围檩宽边材质', '对撑材质', '斜撑材质'):
            _add(v.get(key))
    # 辅助加撑（围檩 + 内支撑）
    for v in (assist_add_dict or {}).values():
        for key in ('围檩长边材质', '围檩宽边材质', '对撑材质', '斜撑材质'):
            _add(v.get(key))
    # 辅助换撑（内支撑）
    for layer_forms in (assist_replace_dict or {}).values():
        for v in layer_forms.values():
            for key in ('对撑材质', '斜撑材质'):
                _add(v.get(key))

    standard = ''
    grade_params = {}
    if steel_dict and steel_std and steel_std in steel_dict:
        standard = str(steel_dict[steel_std].get('Midas对应钢结构规范编号', ''))
        grade_params = steel_dict[steel_std].get('牌号参数', {})

    def _fmt_num(val):
        """数值格式化：整数去掉小数点，非法值原样转字符串。

        Args:
            val: 任意值。

        Returns:
            str: 格式化后的字符串；空值返回 ""。
        """
        if val in (None, ''):
            return ''
        try:
            f = float(val)
            return str(int(f)) if f == int(f) else str(f)
        except (ValueError, TypeError):
            return str(val)

    material_dict = {}
    idx = 1

    # 钢板桩：Q295 不在材料库中，固定为第一个材质
    if Sheet_Pile_typevar == '钢板桩':
        names.discard(Sheet_Pile_materialvar)
        names.discard('Q295')
        material_dict[idx] = {
            'type': 'STEEL',
            'name': 'Q295',
            'spheat': '0',
            'heatco': '0',
            'dampratio': '0.02',
            'standrad': 'JGJ(S)',
            'elast': '206000',
        }
        idx += 1

    for name in sorted(names):
        material_dict[idx] = {
            'type': 'STEEL',
            'name': name,
            'spheat': '0.6',
            'heatco': '0',
            'dampratio': '0.02',
            'standrad': standard,
            'elast': _fmt_num(grade_params.get(name, {}).get('E', '')),
        }
        idx += 1
    return material_dict


def CSV_Cofferdam(cofferdam_sizelst, waler_Hlst, strut_seclst, cap_sizelst, solid_layerlst):
    """组装围堰信息 CSV 的二维行列表（各行列数补齐到最大值）。

    Args:
        cofferdam_sizelst (list): 围堰尺寸 [长(mm), 宽(mm), 高(mm), 顶标高(m)]。
        waler_Hlst (list): 各层围檩高度（距承台底 mm）。
        strut_seclst (list): 内支撑截面 [直径(mm), 厚度(mm)]。
        cap_sizelst (list): 承台尺寸 [长, 宽, 高(mm), 底标高(m)]。
        solid_layerlst (list): 各土层标高（m）。

    Returns:
        list[list[str]]: 8 行 CSV 数据（含表头行），每行已补齐至相同列数。

    Example:
        >>> CSV_Cofferdam([39000, 18000, 22000, 0], [3000, 6500], [820, 10],
        ...               [34000, 16000, 3000, -10], [-0.5, -2, -10])
        [['围堰尺寸', '类型', '长(mm)', '宽(mm)', '高(mm)', '顶标高(m)'], ...]
    """
    line1 = ['围堰尺寸', '类型', '长(mm)', '宽(mm)', '高(mm)', '顶标高(m)']
    line2 = ['', '450mm钢板桩']
    line3 = ['围檩高度']
    line4 = ['内支撑', '直径(mm)', '厚度(mm)']
    line5 = ['']
    line6 = ['承台', '长(mm)', '宽(mm)', '高(mm)', '底标高(m)']
    line7 = ['']
    line8 = ['土层标高(m)']
    # 写入数据
    # line2
    # cofferdam_sizelst = [39000, 18000, 22000, 0]
    for x in cofferdam_sizelst:
        line2.append(str(x))
    # line3
    # waler_Hlst = [3000, 6500, 9000]
    for x in waler_Hlst:
        line3.append(str(x))
    # line5
    # strut_seclst = [820, 10]
    for x in strut_seclst:
        line5.append(str(x))
    # line7
    # cap_sizelst = [34000, 16000, 3000, -10]
    for x in cap_sizelst:
        line7.append(str(x))
    # line8
    # solid_layerlst = [-0.5, -2, -10, -20, -30]
    for x in solid_layerlst:
        line8.append(str(x))
    # 同化格式
    linelst = [line1, line2, line3, line4, line5, line6, line7, line8]
    csv_line_length = max(len(line1), len(line2), len(line3), len(line4), len(line5), len(line6), len(line7), len(line8))
    for lst in linelst:
        lsti = csv_line_length - len(lst)
        for i in range(lsti):
            lst.append('')
    # for x in linelst:
    #     print(x)
    return linelst


def SECTION(Sheet_Pile_typevar, Sheet_Pile_namevar, Waler_section_dict, Strut_section_dict, selected_Waler_sectionlst, selected_add_Waler_sectionlst, selected_Strut_DuiCheng_sectionlst, selected_add_Strut_DuiCheng_sectionlst, selected_Strut_XieCheng_sectionlst, selected_add_Strut_XieCheng_sectionlst, selected_replace_Strut_DuiCheng_sectionlst, selected_replace_Strut_XieCheng_sectionlst, Pile_SEC_info_dict, SKGGZ_D, SKGGZ_t):
    """汇总所有选中的围檩/内支撑/支护桩截面，生成带编号的 MCT 截面命令列表。

    从截面库取各截面的 MCT 文本并去重，再按支护类型追加支护桩截面
    （钢板桩按 B 与 B/2 两个宽度缩放，锁扣钢管桩按 D/t 生成）。

    Args:
        Sheet_Pile_typevar (str): 支护类型，'钢板桩' 或 '锁扣钢管桩'。
        Sheet_Pile_namevar (str): 支护桩截面名称（钢板桩用）。
        Waler_section_dict (dict): 围檩截面库，见 Resource.Waler_sections。
        Strut_section_dict (dict): 内支撑截面库，见 Resource.Strut_sections。
        selected_Waler_sectionlst (list[str]): 选用的围檩截面名。
        selected_add_Waler_sectionlst (list[str]): 辅助加撑选用的围檩截面名。
        selected_Strut_DuiCheng_sectionlst (list[str]): 对撑截面名。
        selected_add_Strut_DuiCheng_sectionlst (list[str]): 辅助加撑对撑截面名。
        selected_Strut_XieCheng_sectionlst (list[str]): 斜撑截面名。
        selected_add_Strut_XieCheng_sectionlst (list[str]): 辅助加撑斜撑截面名。
        selected_replace_Strut_DuiCheng_sectionlst (list[str]): 换撑对撑截面名。
        selected_replace_Strut_XieCheng_sectionlst (list[str]): 换撑斜撑截面名。
        Pile_SEC_info_dict (dict): 支护桩截面信息 {'SEC': list}（钢板桩用）。
        SKGGZ_D (float): 锁扣钢管桩直径（mm）。
        SKGGZ_t (float): 锁扣钢管桩壁厚（mm）。

    Returns:
        list[str]: 每项形如 "截面号, MCT命令文本"。

    Example:
        >>> SECTION('锁扣钢管桩', 'SKGGZ', w_sec, s_sec, [], [], ['377X6钢管桩'], ...)
        ['1, DBUSER, ...', '2, DBUSER, 锁扣钢管桩800X10, ...']
    """
    # xm宽拉森四的截面特性
    def LaSen_sectionvalue(x,H,B,A,Asy,Asz,Ixx,Iyy,Izz,Cyp,Cym,Czp,Czm,Qyb,Qzb,PeriO,PeriI,Centy,Centz,y1,z1,y2,z2,y3,z3,y4,z4,Zyy,Zzz):
        """按宽度比例系数 x/B 缩放拉森钢板桩截面特性，生成 VALUE 命令文本。

        Args:
            x (float): 目标宽度（mm），B 为基准宽度。
            H,B,A,Asy,Asz,Ixx,Iyy,Izz,Cyp,Cym,Czp,Czm,Qyb,Qzb,PeriO,PeriI,
                Centy,Centz,y1,z1,y2,z2,y3,z3,y4,z4,Zyy,Zzz (float): 基准截面特性。

        Returns:
            str: 缩放后的 MIDAS VALUE 截面命令文本。
        """
        scale = x/B # 宽度比例系数
        LaSen_sec_info = 'VALUE, {}钢板桩{}, CC, 0, 0, 0, 0, 0, 0, YES, NO, SB, BUILT, {}, {}, 0, 0, 0, 0, 0, 0, 0' \
                '\n {}, {}, {}, {}, {}, {}' \
                '\n {}, {}, {}, {}, {}, {}, {}, {}, {}, {}' \
                '\n {}, {}, {}, {}, {}, {}, {}, {}, {}, {}'.format(
                          Sheet_Pile_namevar, B*scale, H, B*scale, 
                          A*scale, Asy*scale, Asz*scale, Ixx*scale, Iyy*scale, Izz*pow(scale,2),
                          Cyp*scale, Cym*scale, Czp, Czm, Qyb, Qzb*scale, PeriO-2*B+2*B*scale, PeriI, Centy*scale, Centz,
                          y1*scale, y2*scale, y3*scale, y4*scale, z1, z2, z3, z4, Zyy, Zzz
                          )
        # print(LaSen_sec_info)
        return LaSen_sec_info
    def SKGGZ_section_value(SKGGZ_D, SKGGZ_t):
        """按直径 D 与壁厚 t 生成锁扣钢管桩的 MIDAS DBUSER 截面命令文本。

        Args:
            SKGGZ_D (float): 钢管桩直径（mm）。
            SKGGZ_t (float): 钢管桩壁厚（mm）。

        Returns:
            str: 锁扣钢管桩截面命令文本。
        """
        SKGGZ_sec_info = 'DBUSER, 锁扣钢管桩{}X{}, CC, 0, 0, 0, 0, 0, 0, YES, NO, P, 2, {}, {}, 0, 0, 0, 0, 0, 0, 0, 0'.format(int(SKGGZ_D), int(SKGGZ_t), SKGGZ_D, SKGGZ_t)
        return SKGGZ_sec_info
    selected_Waler_section_mctlst = []
    selected_Strut_section_mctlst = []
    selected_add_Waler_section_mctlst = []
    selected_add_Strut_section_mctlst = []
    # 围檩
    for sec1 in selected_Waler_sectionlst:
        selected_Waler_section_mctlst.append(Waler_section_dict[sec1]['MCT'])
    # 内支撑
    selected_Strut_sectionlst = selected_Strut_DuiCheng_sectionlst + selected_Strut_XieCheng_sectionlst
    selected_Strut_sectionlst = list(set(selected_Strut_sectionlst))
    for sec2 in selected_Strut_sectionlst:
        if sec2 != '/':
            selected_Strut_section_mctlst.append(Strut_section_dict[sec2]['MCT'])
    # 辅助加围檩
    for sec1 in selected_add_Waler_sectionlst:
        selected_add_Waler_section_mctlst.append(Waler_section_dict[sec1]['MCT'])

    # 辅助加内支撑
    selected_add_Strut_sectionlst = selected_add_Strut_DuiCheng_sectionlst + selected_add_Strut_XieCheng_sectionlst
    selected_add_Strut_sectionlst = list(set(selected_add_Strut_sectionlst))
    for sec2 in selected_add_Strut_sectionlst:
        if sec2 != '/':
            selected_add_Strut_section_mctlst.append(Strut_section_dict[sec2]['MCT'])

    # 辅助换撑内支撑
    selected_replace_Strut_section_mctlst = []
    selected_replace_Strut_sectionlst = selected_replace_Strut_DuiCheng_sectionlst + selected_replace_Strut_XieCheng_sectionlst
    selected_replace_Strut_sectionlst = list(set(selected_replace_Strut_sectionlst))
    for sec2 in selected_replace_Strut_sectionlst:
        if sec2 != '/':
            selected_replace_Strut_section_mctlst.append(Strut_section_dict[sec2]['MCT'])

    section_lst = list(set(selected_Waler_section_mctlst + selected_Strut_section_mctlst + selected_add_Waler_section_mctlst + selected_add_Strut_section_mctlst + selected_replace_Strut_section_mctlst))

    if Sheet_Pile_typevar == '钢板桩':
        H,B,A,Asy,Asz,Ixx,Iyy,Izz,Cyp,Cym,Czp,Czm,Qyb,Qzb,PeriO,PeriI,Centy,Centz,y1,z1,y2,z2,y3,z3,y4,z4,Zyy,Zzz = Pile_SEC_info_dict['SEC']
        LaSen_SEC_lst = [LaSen_sectionvalue(x,H,B,A,Asy,Asz,Ixx,Iyy,Izz,Cyp,Cym,Czp,Czm,Qyb,Qzb,PeriO,PeriI,Centy,Centz,y1,z1,y2,z2,y3,z3,y4,z4,Zyy,Zzz) for x in [B, B/2]]
        section_lst = section_lst + LaSen_SEC_lst
    elif Sheet_Pile_typevar == '锁扣钢管桩':
        GGZ_SEC_lst = [SKGGZ_section_value(SKGGZ_D, SKGGZ_t)]
        section_lst = section_lst + GGZ_SEC_lst
    section_mctlst = [f'{i+1}, ' + section_lst[i] for i in range(len(section_lst))]
    return section_mctlst


# 计算施工阶段
def Cal_Stage(waler_elevations, cap_bottom_level, bottom_extra_depth,
              overdig_depth, excavation_face_dewater,
              waterdown_at_bottom, cap_height=0,
              design_water_level=None,
              ground_level=None,
              has_plug=False, has_blinding=False):
    """自动生成施工阶段（stage_dict 格式）

    生成顺序:
      1. 开挖-支撑循环（每层围檩：取土 → 加撑）
      2. 最终取土（至基坑底）
      3. 封底/垫层
      4. 辅助步骤：一道辅助加圈梁（承台顶）→ 从下至上全部拆撑

    Args:
        waler_elevations: list[float] — 各层围檩标高（从高到低）
        cap_bottom_level: float — 承台底标高
        bottom_extra_depth: float — 封底/垫层厚度（基坑底=承台底-此值）
        overdig_depth: float — 超挖深度(m)
        excavation_face_dewater: float — 开挖面降水深度(m)
        waterdown_at_bottom: float — 基坑底降水深度(m)
        cap_height: float — 承台高度(m)，用于辅助加圈梁位置
        design_water_level: float or None — 设防水位标高
        ground_level: float or None — 初始土顶标高，土顶不高于此值
        has_plug: bool — 是否有封底
        has_blinding: bool — 是否有垫层

    Returns:
        {1: {'工况类型':..., '基坑内土顶标高':..., '基坑内水面标高':..., '支撑层数':...}, ...}
    """

    # print(waler_elevations)
    # print(cap_bottom_level)
    # print(bottom_extra_depth)
    # print(overdig_depth)
    # print(excavation_face_dewater)
    # print(waterdown_at_bottom)
    # print(cap_height)
    # print(design_water_level)
    # print(ground_level)
    # print(has_plug)
    # print(has_blinding)

    pit_bottom = round((cap_bottom_level - bottom_extra_depth), 3)  # 基坑底标高
    cap_top = round(cap_bottom_level + cap_height, 3)  # 承台顶标高

    def _clamp_water(water_level):
        """将水面标高限制为不高于设防水位并保留 3 位小数。

        Args:
            water_level (float): 待限制的水面标高（m）。

        Returns:
            float: 限制后的水面标高（m）。
        """
        if design_water_level is not None and water_level > design_water_level:
            return round(design_water_level, 3)
        return round(water_level, 3)

    def _clamp_soil(soil_level):
        """将土顶标高限制为不高于初始地面标高并保留 3 位小数。

        Args:
            soil_level (float): 待限制的土顶标高（m）。

        Returns:
            float: 限制后的土顶标高（m）。
        """
        if ground_level is not None and soil_level > ground_level:
            return round(ground_level, 3)
        return round(soil_level, 3)

    stages = {}
    stage_index = 0

    # if not waler_elevations:
    #     return stages

    # 1. 开挖-支撑循环
    previous_water_level = None
    for i, waler_elevation in enumerate(waler_elevations):
        # 取土：土顶=围檩标高-超挖深度，水面=土顶-开挖面降水（不高于设防水位）
        stage_index += 1
        soil_top_level = _clamp_soil(waler_elevation - overdig_depth)
        water_surface_level = _clamp_water(soil_top_level - excavation_face_dewater)
        previous_water_level = water_surface_level
        stages[stage_index] = {
            "工况类型": "取土",
            "基坑内土顶标高": str(soil_top_level),
            "基坑内水面标高": str(water_surface_level),
            "支撑层数": "/",
        }
        # 加撑：水面与前次开挖一致
        stage_index += 1
        stages[stage_index] = {
            "工况类型": "加撑",
            "基坑内土顶标高": "/",
            "基坑内水面标高": str(previous_water_level),
            "支撑层数": str(i + 1),
        }

    # 2. 最终取土（至基坑底，水面不高于设防水位）
    stage_index += 1
    final_water_level = _clamp_water(pit_bottom - waterdown_at_bottom)
    stages[stage_index] = {
        "工况类型": "取土",
        "基坑内土顶标高": str(pit_bottom),
        "基坑内水面标高": str(final_water_level),
        "支撑层数": "/",
    }
    previous_water_level = final_water_level

    # 3. 封底
    if has_plug:
        stage_index += 1
        stages[stage_index] = {
            "工况类型": "封底",
            "基坑内土顶标高": "/",
            "基坑内水面标高": str(previous_water_level),
            "支撑层数": "/",
        }

    # 4. 垫层
    if has_blinding:
        stage_index += 1
        stages[stage_index] = {
            "工况类型": "垫层",
            "基坑内土顶标高": "/",
            "基坑内水面标高": str(previous_water_level),
            "支撑层数": "/",
        }

    # 5. 辅助步骤：一道辅助加圈梁（承台顶）→ 从下至上全部拆撑
    stage_index += 1
    stages[stage_index] = {
        "工况类型": "辅助加圈梁",
        "基坑内土顶标高": str(cap_top),
        "基坑内水面标高": "/",
        "支撑层数": "/",
    }
    for i in range(len(waler_elevations) - 1, -1, -1):
        stage_index += 1
        stages[stage_index] = {
            "工况类型": "拆撑",
            "基坑内土顶标高": "/",
            "基坑内水面标高": "/",
            "支撑层数": str(i + 1),
        }

    return stages



def Construct_Stage_Line_Data_Check(Construct_Stage_Line_Data, Walers_Strut_selected_dict, Drawdown_height):
    """校验施工阶段顺序是否合理，返回错误提示列表。

    检查项：封底须在开挖至基坑底之后；无封底时开挖至基坑底须在最后；
    每层围檩的"开挖到第i层围檩下"须在"安装第i层围檩及内支撑"之前；
    降水至基坑底须在最后。

    Args:
        Construct_Stage_Line_Data (list[list]): 施工阶段表数据，
            每行首列为工况名称（如 '取土'/'封底'/'开挖至基坑底'）。
        Walers_Strut_selected_dict (dict): 围檩/内支撑选型表，用于取围檩层数。
        Drawdown_height (float): 围檩下超挖深度（m），用于拼接开挖工况名。

    Returns:
        list[str]: 错误/提示信息列表；为空表示校验通过。

    Example:
        >>> Construct_Stage_Line_Data_Check(data, sel, 0.5)
        ['封底工况应在开挖至基坑底工况后']
    """
    Construct_Stage_Line_Lst = [lst[0] for lst in Construct_Stage_Line_Data]
    Waler_rows = len(Walers_Strut_selected_dict.keys())
    Message_box = []

    if '封底' in Construct_Stage_Line_Lst:
        index0 = Construct_Stage_Line_Lst.index('封底')
        index1 = Construct_Stage_Line_Lst.index('开挖至基坑底')
        # 检查 封底 是否在 开挖至基坑底 之后
        if index0 < index1:
            Message_box.append('封底工况应在开挖至基坑底工况后')
    else:
        # 当无封底时, 检查 开挖至基坑底
        index0 = Construct_Stage_Line_Lst.index('开挖至基坑底')
        if index0 != len(Construct_Stage_Line_Lst) - 1:
            Message_box.append(f'开挖至基坑底工况应在最后')

    # 检查 围檩开挖 是否 在 安装 之前
    for i in range(Waler_rows):
        index0 = Construct_Stage_Line_Lst.index(f'开挖到第{i+1}层围檩下{Drawdown_height}m')
        index1 = Construct_Stage_Line_Lst.index(f'安装第{i+1}层围檩及内支撑')
        if index0 > index1:
            Message_box.append(f'安装第{i+1}层支撑工况应该对应开挖工况之后')

    # 检查 降水至基坑底是否在最后
    if '降水至基坑底' in Construct_Stage_Line_Lst:
        index0 = Construct_Stage_Line_Lst.index('降水至基坑底')
        if index0 != len(Construct_Stage_Line_Lst) - 1:
            Message_box.append(f'降水至基坑底工况应在最后')

    return Message_box


# mct文件
def mct_output(
        Cap_X, Cap_Y, Cap_H, Spring_Thickness,
        Sheet_Pile_typevar, Sheet_Pile_materialvar, CofferDam_L, Sheet_Pile_namevar, Pile_SEC_info_dict,
        SKGGZ_D, SKGGZ_t, SKGGZ_Gap,
        Walers_Strut_selected_dict, assist_add_dict, assist_replace_dict,
        Solid_Top_Level, Water_Top_Level, Cap_Bottom_Level, CofferDam_Top_Level,
        Concrete_Blinding_check_var, Concrete_Blinding_thickness_var,
        Concrete_Plug_check_var, Concrete_Plug_thickness_var,
        Drawdown_height, Waterdown_height, Excavation_face_dewater,
        Waler_section_dict, Strut_section_dict,
        recognize_Cap_result_dict, recognize_Bracket_result_dict,
        recognize_Waler_result_dict, recognize_Strut_result_dict,
        recognize_Strut_Replace_result_dict,
        Steel_Sheet_Pile_CofferDam_Load_dict,
        sigma_k_dict, sheet_elasticlink_dict, cofferdam_elasticlink_form_dict,
        cofferdam_cons, none_construction_stage_check_var, construction_stage_check_var,
        Load_Combo_Line_Data, Construct_Stage_Line_Data,
        stage_dict, if_consider_solid_stress_path,
        mct_savepath, if_mcb, if_csv,
        material_dict=None,
        ):
    """组装并输出围堰 MCT 命令流（可选 MCB/CSV），是模型生成的核心实现。

    流程：确定整体/施工阶段模式与荷载组合 -> 计算土压力/土反力/土弹簧 ->
    生成材料、截面、边界组 -> 调用 make_node_elem_load_boundary 生成
    节点/单元/约束/弹性连接/荷载/施工阶段 -> 拼装 MCT 行、写入文件、
    可选调用 MIDAS 生成 MCB、写剪贴板。

    Args:
        各几何与参数含义同 Steel_Sheet_Pile_CofferDam_on_submit_MCT；
        另有：
        Solid_Top_Level, Water_Top_Level (float): 土层顶/水面顶标高（m）。
        Method_for_Pressures_var 等由内部计算，不在此列出。
        mct_savepath (str): MCT 保存路径。
        if_mcb (bool): 是否调用 MIDAS 生成 MCB。
        if_csv (bool): 是否生成 Cofferdam.csv。
        material_dict (dict|None): 材料字典，见 collect_steel_material_dict。

    Returns:
        None: 副作用为写文件、调用 MIDAS、写剪贴板。
    """
    if construction_stage_check_var == True:
        # 整体模型
        combo_dict = {}
    else:
        # 施工阶段定义
        stage_dict = {}
        # 整体模型
        combo_dict = {
            '超挖深度'  : Drawdown_height,
            '基坑底降水': Waterdown_height,
            '开挖面降水': Excavation_face_dewater,
        }
    
    print(Construct_Stage_Line_Data)
    if construction_stage_check_var and Construct_Stage_Line_Data == [['/', '/', '/', '/']]:
        messagebox.showinfo("提示", '当前已选择生成施工阶段\n请点击工况分析中的自动生成按钮生成施工阶段或手动输入')
    else:
        # csv文件相关
        csv_dict = {'cofferdam_sizelst':[],
                    'waler_Hlst':[],
                    'strut_seclst':[],
                    'cap_sizelst':[],
                    'solid_layerlst':[],
                }
        csv_dict['cap_sizelst'].extend([Cap_X*1000, Cap_Y*1000, Cap_H*1000, Cap_Bottom_Level])
        cofferdam_position_line_lst = recognize_Cap_result_dict['position_line_lst']
        csv_dict['cofferdam_sizelst'].append(abs(cofferdam_position_line_lst[0][0]-cofferdam_position_line_lst[1][0]))
        csv_dict['cofferdam_sizelst'].append(abs(cofferdam_position_line_lst[0][1]-cofferdam_position_line_lst[2][1]))
        csv_dict['cofferdam_sizelst'].append(CofferDam_L*1000)
        csv_dict['cofferdam_sizelst'].append(CofferDam_Top_Level)
        csv_dict['strut_seclst'] = ['820', '10']
        
        # 和生成荷载及节点单元均相关的被多次引用的变量
        # Method_for_Pressures_var = Steel_Sheet_Pile_CofferDam_Load_dict['荷载计算方法'] # 水土分算还是水土合算
        Method_for_Pressures_var_lst = [value['计算方法'] for value in Steel_Sheet_Pile_CofferDam_Load_dict['Excel'].values()]
        # 当存在水土合算工况时, 水土分算的两个荷载在模型中也合并为一个荷载, 只有当全部是水土分算时, 土压力和水压力才分开生成
        # 该变量不用于水土压力计算, 用于判断荷载名称和荷载组，以及生成施工阶段时组合荷载组
        Method_for_Pressures_var = '水土合算' if '水土合算' in Method_for_Pressures_var_lst else '水土分算'
        
        Waler_Gap = [float(value['间距']) for value in list(Walers_Strut_selected_dict.values())] # 围檩z间距
        Waler_Distance = list(accumulate(Waler_Gap))
        Waler_Strut_Z = [round(CofferDam_Top_Level - dist, 3) for dist in Waler_Distance] # 初始定义的围檩z坐标
        add_waler_strut_z = [[int(k), float(v['标高'])] for k, v in assist_add_dict.items()] # 辅助加撑的 [层号, z坐标]
        waler_Hlst = []
        for z in Waler_Strut_Z:
            waler_Hlst.append(abs((z-Cap_Bottom_Level))*1000) # 围檩到承台底的距离
        csv_dict['waler_Hlst'] = sorted(waler_Hlst)

        # 计算主动被动土压力
        if Concrete_Blinding_check_var == True and Concrete_Plug_check_var == False:
            Concrete_Blinding_thinckness = Concrete_Blinding_thickness_var # 垫层厚度不为0
            Concrete_Plug_thinckness = 0 # 封底厚度为0
        elif Concrete_Blinding_check_var == False and Concrete_Plug_check_var == True:
            Concrete_Blinding_thinckness = 0 # 垫层厚度为0
            Concrete_Plug_thinckness = Concrete_Plug_thickness_var # 封底厚度不为0

        # Drawdown_height_sigma = 0
        CofferDam_Bottom_Level = round(CofferDam_Top_Level - CofferDam_L, 3) # 桩底标高 
        Concrete_Bottom_Level  = round(Cap_Bottom_Level-max(Concrete_Plug_thinckness, Concrete_Blinding_thinckness), 3) # 基坑底标高

        # 当存在封底时, 对封底单元按0.5m进行划分
        plug_zlst = []
        if Concrete_Plug_thinckness != 0:
            plug_zlst = Divide_Solid_within_LayerThickness(Cap_Bottom_Level, Concrete_Bottom_Level, 0.5, 0.25)
        print('封底标高计算结果')
        print(plug_zlst)

        # 计算圈梁标高
        circle_beam_zlst = [float(v['基坑内土顶标高']) for v in stage_dict.values() if v['工况类型'] == '辅助加圈梁']
        print('圈梁标高', circle_beam_zlst)

        # 土压力及土弹簧计算
        Pa_dicts, Ps0_dicts, Ks_dicts = Cal_Pa_Ps0_Ks(
            Steel_Sheet_Pile_CofferDam_Load_dict, Pile_SEC_info_dict, plug_zlst, circle_beam_zlst, Waler_Strut_Z,
            Sheet_Pile_typevar, SKGGZ_Gap, sigma_k_dict, Drawdown_height, Concrete_Plug_thinckness, Spring_Thickness, 
            CofferDam_Top_Level, CofferDam_Bottom_Level, Concrete_Bottom_Level, Solid_Top_Level, Water_Top_Level, 
            Cap_Bottom_Level, stage_dict, combo_dict, if_consider_solid_stress_path, add_waler_strut_z, csv_dict
        )
        
        # 荷载组相关的组名
        active_earth_pressure_namelst = ['基坑外主动土压力'] if Method_for_Pressures_var == '水土合算' else ['基坑外主动土压力', '基坑外水压力'] # 围堰外侧水土压力相关
        distributed_soil_reaction_keylst = list(Ps0_dicts['层标高'].keys())
        if Method_for_Pressures_var == '水土合算':
            distributed_soil_reaction_namelst = distributed_soil_reaction_keylst  
        else:
            distributed_soil_reaction_namelst = distributed_soil_reaction_keylst + [x.replace('初始土反力', '水压力') for x in distributed_soil_reaction_keylst] # 围堰内侧分布土反力相关
        Load_Group_namelst = active_earth_pressure_namelst + distributed_soil_reaction_namelst
        print('计算方法')
        print(Method_for_Pressures_var_lst)
        print(Method_for_Pressures_var)
        print('')
        print('荷载名称')
        for x in Load_Group_namelst:
            print(x)

        # 除自重外的荷载说明
        StldCase_namelst = [f'{x}, USER,' for x in Load_Group_namelst]

        # 围檩和内支撑的截面
        selected_Waler_sectionlst = []
        selected_Strut_DuiCheng_sectionlst = []
        selected_Strut_XieCheng_sectionlst = []
        for value in list(Walers_Strut_selected_dict.values()):
            selected_Waler_sectionlst.append(value['围檩长边截面'])
            selected_Strut_DuiCheng_sectionlst.append(value['对撑截面'])
            selected_Strut_XieCheng_sectionlst.append(value['斜撑截面'])
        selected_Waler_sectionlst = list(set(selected_Waler_sectionlst)) # 用户选择的围檩截面
        selected_Strut_DuiCheng_sectionlst = list(set(selected_Strut_DuiCheng_sectionlst))
        selected_Strut_XieCheng_sectionlst = list(set(selected_Strut_XieCheng_sectionlst))

        # 辅助加围檩和辅助加内支撑的截面
        selected_add_Waler_sectionlst = []
        selected_add_Strut_DuiCheng_sectionlst = []
        selected_add_Strut_XieCheng_sectionlst = []
        for value in list(assist_add_dict.values()):
            selected_add_Waler_sectionlst.append(value['围檩长边截面'])
            selected_add_Strut_DuiCheng_sectionlst.append(value['对撑截面'])
            selected_add_Strut_XieCheng_sectionlst.append(value['斜撑截面'])
        selected_add_Waler_sectionlst = list(set(selected_add_Waler_sectionlst)) # 用户选择的围檩截面
        selected_add_Strut_DuiCheng_sectionlst = list(set(selected_add_Strut_DuiCheng_sectionlst))
        selected_add_Strut_XieCheng_sectionlst = list(set(selected_add_Strut_XieCheng_sectionlst))

        # 辅助换撑内支撑的截面
        selected_replace_Strut_DuiCheng_sectionlst = []
        selected_replace_Strut_XieCheng_sectionlst = []
        for layer_forms in assist_replace_dict.values():
            for value in layer_forms.values():
                selected_replace_Strut_DuiCheng_sectionlst.append(value['对撑截面'])
                selected_replace_Strut_XieCheng_sectionlst.append(value['斜撑截面'])
        selected_replace_Strut_DuiCheng_sectionlst = list(set(selected_replace_Strut_DuiCheng_sectionlst))
        selected_replace_Strut_XieCheng_sectionlst = list(set(selected_replace_Strut_XieCheng_sectionlst))

        # 边界组
        # 桩底边界
        BNDR_GROUP_lst = ['桩底约束, 0']
        # 封底边界
        if Concrete_Plug_check_var == True:
            BNDR_GROUP_lst = BNDR_GROUP_lst + ['封底, 0']
        # 土弹簧边界
        BNDR_GROUP_lst = BNDR_GROUP_lst + [f'{x}, 0' for x in list(Ks_dicts.keys())]
        # 圈梁
        BNDR_GROUP_lst = BNDR_GROUP_lst + [f'z={x}m处圈梁, 0' for x in circle_beam_zlst]
        # 支护桩、围檩、牛腿相关
        for i in range(len(Waler_Strut_Z)):
            for k, v in cofferdam_elasticlink_form_dict.items():
                if v != '共节点':
                    BNDR_GROUP_lst.append(f'第{i+1}层{k}弹性连接, 0')
        # 辅助加撑相关
        for i, z in add_waler_strut_z:
            for k, v in cofferdam_elasticlink_form_dict.items():
                if v != '共节点':
                    BNDR_GROUP_lst.append(f'第{i}层{k}弹性连接, 0')
        print('')
        print('边界组名称')
        for x in BNDR_GROUP_lst:
            print(x)
        print('')

        # 材料
        if material_dict is None:
            material_dict = {}
        MATERIAL_lst = mct_material_midas(material_dict)

        # 截面
        SECTION_lst = SECTION(Sheet_Pile_typevar, Sheet_Pile_namevar, Waler_section_dict, Strut_section_dict, selected_Waler_sectionlst, selected_add_Waler_sectionlst, selected_Strut_DuiCheng_sectionlst, selected_add_Strut_DuiCheng_sectionlst, selected_Strut_XieCheng_sectionlst, selected_add_Strut_XieCheng_sectionlst, selected_replace_Strut_DuiCheng_sectionlst, selected_replace_Strut_XieCheng_sectionlst, Pile_SEC_info_dict, SKGGZ_D, SKGGZ_t)
        
        # 生成节点、单元、支承、边界、荷载
        node, element, group, cons, elink, loadgrup, spring, stage, loadcomb, BNDR_GROUP_lst = make_node_elem_load_boundary(
            Sheet_Pile_namevar, Pile_SEC_info_dict, CofferDam_Top_Level, CofferDam_Bottom_Level, Cap_Bottom_Level, Method_for_Pressures_var, Waler_Strut_Z, none_construction_stage_check_var, construction_stage_check_var,
            Waler_section_dict, Walers_Strut_selected_dict, recognize_Cap_result_dict, recognize_Bracket_result_dict, recognize_Waler_result_dict, recognize_Strut_result_dict, recognize_Strut_Replace_result_dict, Steel_Sheet_Pile_CofferDam_Load_dict,
            SECTION_lst, selected_Waler_sectionlst, Pa_dicts, Ps0_dicts, Ks_dicts, Drawdown_height, Concrete_Plug_thinckness, cofferdam_cons, Load_Group_namelst, Concrete_Bottom_Level, BNDR_GROUP_lst, stage_dict, 
            Load_Combo_Line_Data, Construct_Stage_Line_Data, plug_zlst, circle_beam_zlst, Sheet_Pile_typevar, SKGGZ_Gap, sheet_elasticlink_dict, cofferdam_elasticlink_form_dict, assist_add_dict, assist_replace_dict, add_waler_strut_z,
            material_dict, Sheet_Pile_materialvar,
            )
        if node != []:
            # 节点
            NODE_lst = node
            # 单元
            ELEMENT_lst = element
            # 结构组
            GROUP_lst = group
            # 弹性连接
            ELASTICLINK_lst = elink
            # 一般支承
            CONSTRAINT_lst = cons
            # 节点弹性支承
            SPRING_lst = spring
            # mct
            mct_lst = [
                    # 版本号，进行各个版本Midas的mct文件测试
                    # ["*VERSION", "8.6.5"],
                    # 单位系
                    ["*UNIT", "N,MM,KJ,C"],
                    # 材料
                    ["*MATERIAL"] + MATERIAL_lst,
                    # 截面
                    ["*SECTION"] + SECTION_lst,
                    # 节点
                    ["*NODE"] + NODE_lst,
                    # 单元
                    ["*ELEMENT"] + ELEMENT_lst,
                    # 结构组
                    ["*GROUP"] + GROUP_lst,
                    # 约束组
                    ["*BNDR-GROUP"] + BNDR_GROUP_lst,
                    # 弹性连接
                    # ["*ELASTICLINK"] + ELASTICLINK_lst,
                    # 一般支承
                    ["*CONSTRAINT"] + CONSTRAINT_lst,
                    # 节点弹性支承
                    ["*SPRING"] + SPRING_lst,
                    # 荷载组
                    ["*LOAD-GROUP", '自重'] + Load_Group_namelst,
                    # 荷载说明
                    ["*STLDCASE", "自重, USER,"] + StldCase_namelst,
                    # 自重
                    ["*USE-STLD, 自重", "*SELFWEIGHT", "0, 0, {}, 自重".format(-1)],
                ]
            # 除自重外的荷载
            load_appendlst = []
            for x in Load_Group_namelst:
                try:
                    load_lst = [f"*USE-STLD, {x}", "*BEAMLOAD"] + loadgrup[x]
                    load_appendlst.append(load_lst)
                except:
                    pass
            # 弹性连接
            elasticlink_lst = [["*ELASTICLINK"] + ELASTICLINK_lst] if ELASTICLINK_lst != [] else []
            # 施工阶段
            stage_lst = [['*STAGE'] + stage] if construction_stage_check_var == True else []
            #  荷载组合
            loadcomb_lst = [['*LOADCOMB'] + loadcomb] if none_construction_stage_check_var == True else []
            # 更新 mct_lst
            mct_lst = mct_lst + elasticlink_lst + load_appendlst + stage_lst + loadcomb_lst
            # 换行 
            mct_lst_tolines = "\n".join([item for sublist in mct_lst for item in sublist])
            # 复制粘贴
            SET_CLIP_STRING(mct_lst_tolines)
            print("mct命令流已粘贴到剪切板中")
            # 根据指定路径生成mct文件
            print(mct_savepath)
            write_file_within_path(mct_savepath, mct_lst_tolines, "gbk")
            if if_mcb: # if_mcb == True
                # 生成mcb模型
                MidasURL()
                mcb_savepath = file_extension_Modified(mct_savepath, '.mcb')
                importmct_to_mcb(mct_savepath, mcb_savepath)
            if if_csv:
                # 生成csv文件
                try:
                    cofferdam_sizelst, waler_Hlst, strut_seclst, cap_sizelst, solid_layerlst = [value for value in csv_dict.values()]
                    csv_lst = CSV_Cofferdam(cofferdam_sizelst, waler_Hlst, strut_seclst, cap_sizelst, solid_layerlst)
                    write_csv('Cofferdam.csv', csv_lst)
                except:
                    pass
            # # 将参数文件保存至mcb同目录下
            # if Applocation != None:
            #     sources = [os.path.join(Applocation, 'Support', 'basic_param', Basic_txt_filename), os.path.join(Applocation, 'Support', 'program_param', Program_txt_filename)]
            #     save_dir = Path(mct_savepath).parent  # 获取目录部分
            #     destinations = [str(save_dir / "基本参数文件.txt"), str(save_dir / "项目参数文件.txt")]
            #     copy_txt_files(sources, destinations, overwrite=True)
            print('全部阶段完成')
            # return mct_lst


def Soil_Spring_Stiffness(solid_layer_dict, Solid_Top_Level, Cap_Bottom_Level, Concrete_Bottom_Level, CofferDam_But_Level, spring_zdict, Pile_SEC_info_dict, LayerThickness, Concrete_Plug_thinckness, Sheet_Pile_typevar, SKGGZ_Gap, stage_dict):
    """按施工阶段计算各标高处的等代土弹簧刚度 Ks。

    对每个"取土至xm处"工况，将桩侧土层按分层厚度切分为土条，取每层中点，
    按各土层加权平均的 m 值与深度、土条厚度、桩计算宽度计算 Ks；封底范围内
    的弹簧标记为固结。

    Args:
        solid_layer_dict (dict): 土层信息，结构形式
            {土层名称: {'层底标高': float, '层顶标高': float, '重度':..., '黏聚力':..., '内摩擦角':...}}。
        Solid_Top_Level (float): 土层顶标高（m）。
        Cap_Bottom_Level (float): 承台底标高（m）。
        Concrete_Bottom_Level (float): 基坑底标高（m）。
        CofferDam_But_Level (float): 围堰底（桩底）标高（m）。
        spring_zdict (dict): 各工况的弹性支承点标高，结构形式
            {'取土至xm处土层弹性支承点标高': [z1, z2, ...], ...}。
        Pile_SEC_info_dict (dict): 支护桩截面信息 {'SEC': [...]}。
        LayerThickness (float): 土弹簧分层厚度（m）。
        Concrete_Plug_thinckness (float): 封底厚度（m），非 0 时该范围弹簧固结。
        Sheet_Pile_typevar (str): 支护类型，'钢板桩' 或 '锁扣钢管桩'。
        SKGGZ_Gap (float): 锁扣钢管桩中心间距（mm），作为计算宽度。
        stage_dict (dict): 施工阶段字典。

    Returns:
        dict: 结构形式
            {'取土至xm处土层弹性支承点': [[z, if_concretion, m, Ks], ...], ...}，
            其中 if_concretion 为是否固结(bool)，m 为加权平均 m 值，Ks 为刚度。

    Example:
        >>> Soil_Spring_Stiffness(soil, 0.5, -3.5, -5.0, -10.0, spring_z,
        ...                       pile_sec, 0.5, 0, '钢板桩', 0, {})
        {'取土至-5.0m处土层弹性支承点': [[-4.75, False, 5000.0, 1234.5], ...]}
    """
    Ks_dict = {}
    if Sheet_Pile_typevar == '钢板桩':
        Pile_B = Pile_SEC_info_dict['SEC'][1] # 钢板桩计算宽度
    elif Sheet_Pile_typevar == '锁扣钢管桩':
        Pile_B = SKGGZ_Gap

    Solid_Hlst = [Solid_Top_Level, Concrete_Bottom_Level, CofferDam_But_Level]
    for value in solid_layer_dict.values():
        Solid_Hlst.append(value['层底标高'])
    Solid_Hlst = sorted(Solid_Hlst, reverse=True)

    valuelst = [[float(value['层底标高']), float(value['黏聚力']), float(value['内摩擦角'])] for value in list(solid_layer_dict.values())] # [层底标高, 黏聚力, 内摩擦角]
    if stage_dict:
        # Zlst = [float(stage_value['基坑内土顶标高']) for stage_value in stage_dict.values() if '取土' in stage_value['工况类型'] ] + [Concrete_Bottom_Level]
        Zlst = [float(stage_value['基坑内土顶标高']) for stage_value in stage_dict.values() if '取土' in stage_value['工况类型'] or '抽水' in stage_value['工况类型']]
    else:
        Zlst = [Concrete_Bottom_Level]
    Zlst = sorted(set(Zlst), reverse=True) # 去除重复元素并按标高从高到低排序
    print('所有有效的土层起始标高')
    print(Zlst)

    Condition_Solid_Hlst = []
    for z in Zlst:
        # Solid_Start_Level = z if concrete_plug_after_waler_i >= zi+1 else Concrete_Bottom_Level # 当存在封底时, 封底之后的施工阶段的土顶标高从基坑底开始算起, 封底之前的不变
        lst = sorted(set([h for h in Solid_Hlst + [z] if h >= CofferDam_But_Level and h <= z]), reverse=True)
        lst = [round(x, 3) for x in lst if x <= Solid_Top_Level]
        Condition_Solid_Hlst.append(lst)
    Condition_Solid_Infolst = []
    for lst in Condition_Solid_Hlst:
        Solid_Infolst = []
        for i in range(len(lst)-1):
            Level2 = lst[i+1]
            for infolst in valuelst:
                if Level2>=infolst[0]:
                    Solid_Infolst.append([infolst[1], infolst[2]])
                    break
        Condition_Solid_Infolst.append(Solid_Infolst)
    # for i in range(len(Condition_Solid_Hlst)):
    #     print('Condition_Solid_Hlst_i:   ', Condition_Solid_Hlst[i])
    #     print('Condition_Solid_Infolst_i:', Condition_Solid_Infolst[i])

    # print('')
    for i, _key in enumerate(list(spring_zdict.keys())):
        key = _key.replace('标高', '') # key = '取土至xm处土层弹性支承点标高' -->  # key = '取土至xm处土层弹性支承点'
        Ks_dict[key] = []
        Solid_Layer_Hlst = Condition_Solid_Hlst[i]
        Solid_Layer_Infolst = Condition_Solid_Infolst[i]
        m_method_Hlst = Divide_Solid_within_LayerThickness(Solid_Layer_Hlst[0], Solid_Layer_Hlst[-1], LayerThickness, LayerThickness / 2)
        print(m_method_Hlst)
        for j in range(len(m_method_Hlst)-1):
            Level1 = m_method_Hlst[j]
            Level2 = m_method_Hlst[j+1]
            z = (Level1 + Level2) / 2
            if z < Solid_Top_Level: # 在土层以下
                if_concretion = False # 默认无封底, 非固结
                if z >= Concrete_Bottom_Level and z <= Cap_Bottom_Level and Concrete_Plug_thinckness != 0: # 有封底, 需要固结
                    if_concretion = True
                # 判断土条的分层情况
                Solid_Slince_lst = intersect_segment_with_intervals(Solid_Layer_Hlst, [Level1, Level2]) # 
                # 判断土条的分层情况下对应的土层信息
                Solid_Slince_Infolst = [[Solid_Slince_lst[m][-1]] + Solid_Layer_Infolst[Solid_Slince_lst[m][0]-1] for m in range(len(Solid_Slince_lst))]
                # 计算m所需要的土层厚度、黏聚力和内摩擦角
                mlst = [[Solid_Slince_Infolst[m][0], cal_m(Solid_Slince_Infolst[m][2], Solid_Slince_Infolst[m][1])] for m in range(len(Solid_Slince_Infolst))]
                # 计算加权平均m
                m_sum = sum(w * v for w, v in mlst)
                h_sum = sum(w for w, v in mlst)
                m = m_sum / h_sum
                # print('mlst信息', mlst)
                # print('m总和', m_sum)
                # print('h总和', h_sum)
                # 计算Ks
                m_Level = round((Level1 + Level2)/2, 3) # 计算土层标高
                Z = round(abs(m_Level - Solid_Layer_Hlst[0]), 3) # 计算深度
                Ks = cal_Ks(m, Z, h_sum, Pile_B/1000)
                # print('z坐标',m_Level, ' 深度',Z, ' 加权平均m',m, ' Ks',Ks)
                Ks_dict[key].append([round(z,3), if_concretion, m, Ks])
    print('')
    for k, v in Ks_dict.items():
        print(k)
        print(v)
        print('')
    return Ks_dict


def delta_sigma_k_j(sigma_k_dict):
    """将 UI 的超载（附加荷载）原始参数换算为各面/各处的附加应力。

    均布荷载直接取 q0；矩形/条形局部荷载调用 Localized_Rectangular_Load /
    Localized_Strip_Load 计算影响深度范围 zamin/zamax 与应力 sigmak。

    Args:
        sigma_k_dict (dict): 附加荷载信息，结构形式
            {'均布附加荷载': {面号: {'q0': str, ...}, ...},
             '矩形局部附加荷载': {面号: {'angle','p0','b','a','d','l','p2','c'}, ...},
             '条形局部附加荷载': {面号: {'angle','p0','b','a','d'}, ...}}，
            未设置的字段为 '/'。

    Returns:
        dict: 结构形式
            {'均布附加荷载': {面号: {'sigmak': float}},
             '矩形局部附加荷载': {面号: [{'zamin': float, 'zamax': float, 'sigmak': float}, ...]},
             '条形局部附加荷载': {面号: {'zamin': float, 'zamax': float, 'sigmak': float}}}。

    Example:
        >>> delta_sigma_k_j(sigma_k_dict)['均布附加荷载']
        {'1': {'sigmak': 20.0}}
    """
    sigma_k_cal_dict = {
        '均布附加荷载': {},
        '矩形局部附加荷载': {},
        '条形局部附加荷载': {},
    }

    def is_valid(params):
        """判断荷载参数字典是否至少有一个字段已设置（非全 '/'）。

        Args:
            params (dict): 荷载参数字典。

        Returns:
            bool: 存在任一值不为 '/' 时返回 True。
        """
        return any(val != '/' for val in params.values())

    for num, params in sigma_k_dict['均布附加荷载'].items():
        if params['q0'] != '/':
            sigma_k_cal_dict['均布附加荷载'][num] = {'sigmak': round(float(params['q0']), 3)}

    for num, params in sigma_k_dict['矩形局部附加荷载'].items():
        if is_valid(params):
            sigma_k_cal_dict['矩形局部附加荷载'][num] = []
            angle = math.radians(float(params['angle']))
            p0 = float(params['p0'])
            b = float(params['b'])
            a = float(params['a'])
            d = float(params['d'])
            l = float(params['l'])
            p2 = params['p2']
            c = params['c']
            zamin, zamax, sigmak = Localized_Rectangular_Load(angle, p0, b, a, d, l)
            sigma_k_cal_dict['矩形局部附加荷载'][num].append({
                'zamin': round(zamin, 3),
                'zamax': round(zamax, 3),
                'sigmak': round(sigmak, 3),
            })
            if p2 != '/' and c != '/':
                p2 = float(p2)
                c = float(c)
                zamin, zamax, sigmak = Localized_Rectangular_Load(angle, p2, b, (a+b+c), d, l)
                sigma_k_cal_dict['矩形局部附加荷载'][num].append({
                    'zamin': round(zamin, 3),
                    'zamax': round(zamax, 3),
                    'sigmak': round(sigmak, 3),
                })

    for num, params in sigma_k_dict['条形局部附加荷载'].items():
        if is_valid(params):
            angle = math.radians(float(params['angle']))
            p0 = float(params['p0'])
            b = float(params['b'])
            a = float(params['a'])
            d = float(params['d'])
            zamin, zamax, sigmak = Localized_Strip_Load(angle, p0, b, a, d)
            sigma_k_cal_dict['条形局部附加荷载'][num] = {
                'zamin': round(zamin, 3),
                'zamax': round(zamax, 3),
                'sigmak': round(sigmak, 3),
            }

    # for k, v in sigma_k_cal_dict.items():
    #     if v != {}:
    #         print(f'添加{k}')
    #         print(v)
    #     else:
    #         print(f'未添加{k}')
    return sigma_k_cal_dict


def Active_Earth_Pressure(solid_layer_dict, Solid_Top_Level, Water_Top_Level, CofferDam_But_Level, level_zdict, spring_zdict, Waler_Strut_Z, excavate_solid_levels, excavate_water_levels, plug_zlst, sigma_k_cal_dict, water_drawdown_height_below_excavation_bottom_lst, circle_beam_zlst, add_waler_strut_z):
    """计算围堰四面的基坑外主动土压力（含竖向应力与单元荷载）。

    汇总所有特征标高构建统一分层，逐层判断含水/无水并取对应土层参数，
    计算层顶/层底土与水竖向应力、主动土压力，并按水土分算/合算整理为
    单元荷载结构。

    Args:
        solid_layer_dict (dict): 土层信息，结构形式
            {土层名称: {'层底标高','层顶标高','重度','黏聚力','内摩擦角','计算方法'}}。
        Solid_Top_Level (float): 土层顶标高（m）。
        Water_Top_Level (float): 水面顶标高（m）。
        CofferDam_But_Level (float): 围堰底标高（m），低于此标高不参与计算。
        level_zdict (dict): 特征标高字典 {名称: 标高(m)}。
        spring_zdict (dict): 各工况弹性支承点标高。
        Waler_Strut_Z (list[float]): 各层围檩标高（m）。
        excavate_solid_levels (list[float]): 各取土工况土顶标高（去重降序）。
        excavate_water_levels (list[float]): 各取土工况水面标高（去重降序）。
        plug_zlst (list[float]): 封底单元划分标高（m）。
        sigma_k_cal_dict (dict): delta_sigma_k_j 的输出。
        water_drawdown_height_below_excavation_bottom_lst (list[float]): 基坑底降水标高。
        circle_beam_zlst (list[float]): 圈梁标高（m）。
        add_waler_strut_z (list[list]): 辅助加撑的 [层号, 标高]。

    Returns:
        dict: 结构形式
            {'1'~'4': {'层标高': {层名: {'层顶标高','层底标高'}},
                       '层参数': {层名: {'土层状态','重度','黏聚力','内摩擦角','计算方法'}},
                       '层荷载': {层名: {'竖向应力标准值','主动土压力','层标高','计算方法'}},
                       '单元荷载': {层名: [...]}}},
            数字键代表围堰四个面。

    Example:
        >>> Active_Earth_Pressure(...)['1']['层荷载']['第1层']['主动土压力']
        [[0.0, 12.5], [0.0, 0.0]]
    """
    # print(solid_layer_dict)

    # (1) 将超载信息的zamin/zamax（埋深，以0标高计）转换为实际z坐标
    extra_zlst = []
    for params in sigma_k_cal_dict['矩形局部附加荷载'].values():
        for load_dict in params:
            zamin_actual = round(Solid_Top_Level - load_dict['zamin'], 3)
            zamax_actual = round(Solid_Top_Level - load_dict['zamax'], 3)
            extra_zlst.extend([zamin_actual, zamax_actual])
    for params in sigma_k_cal_dict['条形局部附加荷载'].values():
        zamin_actual = round(Solid_Top_Level - params['zamin'], 3)
        zamax_actual = round(Solid_Top_Level - params['zamax'], 3)
        extra_zlst.extend([zamin_actual, zamax_actual])

    # (2) 将超载z坐标写入zlst
    zlst = sorted(set([round(x, 3) for x in plug_zlst] +
                      [round(x, 3) for x in extra_zlst] + 
                      [round(x, 3) for x in Waler_Strut_Z] +
                      [round(x, 3) for x in circle_beam_zlst] +
                      [round(x[1], 3) for x in add_waler_strut_z] +
                      [round(x, 3) for x in excavate_solid_levels] +
                      [round(x, 3) for x in excavate_water_levels] +
                      [round(x, 3) for x in list(level_zdict.values())[1:]] +
                      [round(x, 3) for v in list(spring_zdict.values()) for x in v] +
                      [round(x, 3) for x in water_drawdown_height_below_excavation_bottom_lst]
                      ), key=lambda x: x, reverse=True)
    
    zlst = [z for z in zlst if z >= CofferDam_But_Level]

    # 构建分层（各面共用同一个zlst）
    layer_levels = {}
    for i in range(len(zlst) - 1):
        layer_levels[f'第{i + 1}层'] = {'层顶标高': zlst[i], '层底标高': zlst[i + 1]}

    # 判断每层的土层参数（各面共用）
    layer_params = {}
    layer_levels_values = list(layer_levels.values())
    for k in range(len(layer_levels_values)):
        # 判断土层是否含水, 当该层底标高大于等于水面标高时, 为无水层
        if layer_levels_values[k]['层底标高'] >= Water_Top_Level:
            layer_state = '无水'
        else:
            layer_state = '含水'
        # 判断该层是纯水还是土层
        if layer_levels_values[k]['层底标高'] < Solid_Top_Level:
            for key, value in solid_layer_dict.items():
                if layer_levels_values[k]['层底标高'] >= float(value['层底标高']):
                    solid_layer = key
                    break
            layer_solid_Gs = float(solid_layer_dict[solid_layer]['重度'])
            layer_solid_c = float(solid_layer_dict[solid_layer]['黏聚力'])
            layer_solid_fai = float(solid_layer_dict[solid_layer]['内摩擦角'])
            layer_method = solid_layer_dict[solid_layer]['计算方法']
            layer_params[f'第{k + 1}层'] = {'土层状态': layer_state, '重度': layer_solid_Gs, '黏聚力': layer_solid_c, '内摩擦角': layer_solid_fai, '计算方法': layer_method}
        else:
            layer_params[f'第{k + 1}层'] = {'土层状态': '含水', '重度': 10, '黏聚力': 0, '内摩擦角': 0, '计算方法': '/'}

    # (3) 逐面（1~4）计算主动土压力
    result_by_face = {}
    for face_num in ['1', '2', '3', '4']:
        # 确定该面的均布附加荷载
        q0 = sigma_k_cal_dict.get('均布附加荷载', {}).get(face_num, {}).get('sigmak', 0)

        # 确定该面的局部附加荷载
        local_rectangular_loads = []
        if face_num in sigma_k_cal_dict.get('矩形局部附加荷载', {}):
            params = sigma_k_cal_dict['矩形局部附加荷载'][face_num]
            for load_dict in params:
                local_rectangular_loads.append({
                    'zamin_actual': round(Solid_Top_Level - load_dict['zamin'], 3),
                    'zamax_actual': round(Solid_Top_Level - load_dict['zamax'], 3),
                    'sigmak': load_dict['sigmak'],
                })
        local_strip_loads = []
        if face_num in sigma_k_cal_dict.get('条形局部附加荷载', {}):
            params = sigma_k_cal_dict['条形局部附加荷载'][face_num]
            local_strip_loads.append({
                'zamin_actual': round(Solid_Top_Level - params['zamin'], 3),
                'zamax_actual': round(Solid_Top_Level - params['zamax'], 3),
                'sigmak': params['sigmak'],
            })

        # print(f'第{face_num}面均布附加荷载')
        # print(q0)
        # print(f'第{face_num}面矩形局部附加荷载')
        # print(local_rectangular_loads)
        # print(f'第{face_num}面条形局部附加荷载')
        # print(local_strip_loads)

        # 计算每一层的主动土压力
        layer_loads = {}
        soild_sigma0 = q0
        water_sigma0 = 0
        for i, key in enumerate(list(layer_levels.keys())):
            layer_level1 = layer_levels[key]['层顶标高']
            layer_level2 = layer_levels[key]['层底标高']
            layer_state = layer_params[key]['土层状态']
            solid_layer_Gs = layer_params[key]['重度']
            solid_layer_c = layer_params[key]['黏聚力']
            solid_layer_fai = layer_params[key]['内摩擦角']
            layer_method = layer_params[key]['计算方法']
            solid_layer_Ka = cal_Ka(solid_layer_fai)

            # 判断当前层是否在局部超载影响范围内
            sigma_k_add1 = 0 # 层顶附加荷载
            sigma_k_add2 = 0 # 层底附加荷载
            for load in local_rectangular_loads:
                if layer_level1 <= load['zamin_actual'] and layer_level1 > load['zamax_actual']:
                    sigma_k_add1 += load['sigmak']
                if layer_level2 < load['zamin_actual'] and layer_level2 >= load['zamax_actual']:
                    sigma_k_add2 += load['sigmak']
            for load in local_strip_loads:
                if layer_level1 <= load['zamin_actual'] and layer_level1 > load['zamax_actual']:
                    sigma_k_add1 += load['sigmak']
                if layer_level2 < load['zamin_actual'] and layer_level2 >= load['zamax_actual']:
                    sigma_k_add2 += load['sigmak']
            # print(f'第{face_num}面第{i+1}层 层顶标高:{layer_level1}, 层底标高:{layer_level2}, 均布附加荷载:{q0}, 层顶局部附加荷载:{sigma_k_add1}, 层底局部附加荷载:{sigma_k_add2}')

            # 计算除局部荷载外的每一层的层顶层底竖向应力
            if layer_state == '含水':
                solid_sigma1, solid_sigma2 = cal_sigmak_solid(soild_sigma0, solid_layer_Gs, layer_level1, layer_level2, 10)
                water_sigma1, water_sigma2 = cal_sigmak_water(water_sigma0, layer_level1, layer_level2, 10)
            else:
                solid_sigma1, solid_sigma2 = cal_sigmak_solid(soild_sigma0, solid_layer_Gs, layer_level1, layer_level2, 0)
                water_sigma1, water_sigma2 = [0, 0]
            soild_sigma0 = solid_sigma2
            water_sigma0 = water_sigma2

            # 加上层顶层底的局部附加荷载
            solid_sigma1 = solid_sigma1 + sigma_k_add1
            solid_sigma2 = solid_sigma2 + sigma_k_add2
            # print(f'第{face_num}面第{i+1}层 层顶土竖向应力:{round(solid_sigma1,3)}, 层底土竖向应力:{round(solid_sigma2,3)}, 层顶水竖向应力:{round(water_sigma1,3)}, 层底水竖向应力:{round(water_sigma2,3)}')
            
            soild_pak1, soild_pak2 = cal_pak_soild(solid_layer_Ka, solid_sigma1, solid_sigma2, solid_layer_c)
            water_pak1, water_pak2 = cal_pak_water(solid_layer_Ka, water_sigma1, water_sigma2, layer_method)
            # print(f'第{face_num}面第{i+1}层 计算方法:{layer_method}, 层顶主动土压力:{round(soild_pak1,3)}, 层底主动土压力:{round(soild_pak2,3)}, 层顶水压力:{round(water_pak1,3)}, 层底水压力:{round(water_pak2,3)}')
            # print('')

            layer_loads[key] = {
                '竖向应力标准值': [[round(solid_sigma1, 3), round(solid_sigma2, 3)], [round(water_sigma1, 3), round(water_sigma2, 3)]],
                '主动土压力': [[round(soild_pak1, 3), round(soild_pak2, 3)], [round(water_pak1, 3), round(water_pak2, 3)]],
                '层标高': [layer_level1, layer_level2],
                '计算方法': layer_method,
            }
        
        elem_loads = {}
        for key, value in layer_loads.items():
            elem_loads[key] = []
            level_lst = value['层标高']
            layer_method = value['计算方法']
            pak_lst1, pak_lst2 = value['主动土压力']
            # 土
            if layer_method == '水土合算':
                pak1 = pak_lst1[0] + pak_lst2[0]
                pak2 = pak_lst1[1] + pak_lst2[1]
                elem_loads[key] = [non_negative_elem_load([pak1, pak2], level_lst)]
            elif layer_method == '水土分算':
                pak_soild = non_negative_elem_load(pak_lst1, level_lst)
                pak_water = pak_lst2
                elem_loads[key] = [pak_soild, [level_lst, [0, 1], pak_water]]
            # 纯水 layer_method == '/'
            else:
                pak_water = pak_lst2
                # 纯水层以 len=2 结构（土为0、水承载）输出，避免被生成mct时误归为"水土合算"而进入土压力组
                elem_loads[key] = [non_negative_elem_load([0, 0], level_lst), [level_lst, [0, 1], pak_water]]

        # for k, v in elem_loads.items():
        #     print(k,':',v)
        # print('')

        result_by_face[face_num] = {'层标高': layer_levels, '层参数': layer_params, '层荷载': layer_loads, '单元荷载': elem_loads}
    return result_by_face



# 计算分布土反力
# 当存在施工阶段时：根据取土/抽水工况的土顶标高计算初始土反力
#   - 陆地钢板桩围堰：key为"取土至xm处初始土反力"
#   - 深水钢板桩围堰：工况类型为"取土"时key为"取土至xm处初始土反力"，工况类型为"抽水"时key为"抽水至xm处初始土反力"
# 当不存在施工阶段时：以基坑底标高作为起始标高，水位取基坑底降水标高，key为"取土至xm处初始土反力"
def Distributed_Soil_Reaction(Active_Earth_Pressure_layer_levels_dict, Active_Earth_Pressure_layer_params_dict, Concrete_Bottom_Level, Solid_Top_Level, Water_Top_Level, stage_dict, excavate_stages, combo_dict, solid_layer_dict, if_consider_solid_stress_path):
    """计算围堰内侧各工况的分布土反力（初始土反力）单元荷载。

    按取土/抽水工况确定计算起始标高（水土最高点），逐层计算竖向应力与
    分布土反力；当考虑土的应力路径时叠加历史附加应力。整体模型（无施工
    阶段）则以基坑底标高为起始、基坑底降水为水位。

    Args:
        Active_Earth_Pressure_layer_levels_dict (dict): 主动土压力计算的层标高
            {层名: {'层顶标高','层底标高'}}。
        Active_Earth_Pressure_layer_params_dict (dict): 层参数
            {层名: {'土层状态','重度','黏聚力','内摩擦角','计算方法'}}。
        Concrete_Bottom_Level (float): 基坑底标高（m）。
        Solid_Top_Level (float): 土层顶标高（m）。
        Water_Top_Level (float): 水面顶标高（m）。
        stage_dict (dict): 施工阶段字典；为空表示整体模型。
        excavate_stages (list[dict]): 取土/抽水工况信息
            [{'工况序号','工况类型','基坑内土顶标高','基坑内水面标高'}, ...]。
        combo_dict (dict): 整体模型荷载组合参数
            {'超挖深度','基坑底降水','开挖面降水'}。
        solid_layer_dict (dict): 土层信息，见 Active_Earth_Pressure。
        if_consider_solid_stress_path (bool): 是否考虑土的应力路径。

    Returns:
        dict: 结构形式
            {'层标高': {工况名: {...}}, '层参数': {工况名: {...}},
             '层荷载': {工况名: {层名: {...}}}, '单元荷载': {工况名: {层名: [...]}}}，
            工况名形如 '取土至xm处初始土反力' 或 '抽水至xm处初始土反力'。

    Example:
        >>> Distributed_Soil_Reaction(...)['单元荷载'].keys()
        dict_keys(['取土至-5.0m处初始土反力'])
    """
    layer_levels = {} # 需要计算的层数
    layer_params = {} # 计算该层土压力需要的参数
    layer_loads = {} # 该层土压力计算的结果
    elem_loads = {} # 该层主动土压力水土合起来计算的结果，考虑水土分算或合算

    # 计算基坑内从原始水土标高到当前工况水土标高之间的附加应力（用于分布土反力计算）
    def calc_excavation_solid_sigma_k_add(Solid_Top_Level, Water_Top_Level, current_solid_level, solid_layer_dict, gamma_w = 10):
        """计算基坑内从原始土顶到当前工况土顶之间累计的历史竖向应力（分布土反力用）。

        逐土层取交集厚度，按水上用天然容重、水下用浮容重累加。

        Args:
            Solid_Top_Level (float): 原始土层顶标高（m）。
            Water_Top_Level (float): 原始水面顶标高（m）。
            current_solid_level (float): 当前工况土顶标高（m）。
            solid_layer_dict (dict): 土层信息，见 Active_Earth_Pressure。
            gamma_w (float): 水的重度（kN/m3），默认 10。

        Returns:
            float: 累计附加竖向应力（kPa）；当前土顶不低于原始土顶时返回 0。
        """
        if current_solid_level >= Solid_Top_Level:
            return 0
        sigma_add = 0
        # 逐土层计算
        for layer_info in solid_layer_dict.values():
            layer_top = float(layer_info.get('层顶标高'))
            layer_bottom = float(layer_info.get('层底标高'))
            gamma = float(layer_info['重度'])
            # 取交集：[Solid_Top_Level, current_solid_level] 与 [layer_top, layer_bottom]
            seg_top = min(Solid_Top_Level, layer_top)
            seg_bottom = max(current_solid_level, layer_bottom)
            if seg_top <= seg_bottom:
                continue
            thickness = seg_top - seg_bottom
            # 判断水上/水下
            if seg_bottom >= Water_Top_Level:
                # 全在水上，用天然容重
                sigma_add += gamma * thickness
            elif seg_top >= Water_Top_Level and seg_bottom <= Water_Top_Level:
                # 部分水上、部分水下
                above_water = seg_top - Water_Top_Level
                below_water = Water_Top_Level - seg_bottom
                sigma_add += gamma * above_water + (gamma - gamma_w) * below_water
            elif seg_top <= Water_Top_Level:
                # 全在水下，用浮容重
                sigma_add += (gamma - gamma_w) * thickness
            # print(f'[{seg_top}, {seg_bottom}]--> {round(sigma_add, 3)}')
        return round(sigma_add, 3)

    # 构建 zdict，key 根据工况类型命名
    # 取土/取土开挖：key="取土至{土顶标高}m处初始土反力"
    # 抽水：key="抽水至{水面标高}m处初始土反力"
    # value = max(土顶标高, 水面标高)，即计算起始标高取水土最高点
    # 土顶标高以上按水计算，土顶标高以下根据水面标高判断含水/无水
    zdict = {}
    # 存储每个工况的详细信息，用于后续判断土层状态
    stage_detail_map = {}
    if stage_dict:  # 存在施工阶段
        print(excavate_stages)
        # 获取指定工况的信息，去重后从大到小排序
        excavate_stage_info = []
        for stage in excavate_stages:
            excavate_stage_info.append({
                '土顶标高': stage['基坑内土顶标高'],
                '水面标高': stage['基坑内水面标高'],
                '工况类型': stage['工况类型']
            })
        # 去重（按最终生成的 stage_key，即真正决定 zdict 唯一键的组合）
        # 注意：抽水工况的土顶标高可能相同但水面标高不同，需按水面标高区分
        seen = set()
        unique_stages = []
        for info in excavate_stage_info:
            stage_type = info['工况类型']
            solid_level = info['土顶标高']
            water_level = info['水面标高']
            if stage_type in ['取土', '取土开挖']:
                stage_key = f'取土至{solid_level}m处初始土反力'
            elif stage_type == '抽水':
                stage_key = f'抽水至{water_level}m处初始土反力'
            else:
                continue
            if stage_key not in seen:
                seen.add(stage_key)
                unique_stages.append(info)
        # 从大到小排序（按计算起始标高，即水土最高点）
        unique_stages.sort(key=lambda x: max(x['土顶标高'], x['水面标高']), reverse=True)

        for info in unique_stages:
            stage_type = info['工况类型']
            solid_level = info['土顶标高']
            water_level = info['水面标高']
            # 计算起始标高取水土最高点
            calc_level = max(solid_level, water_level)
            if stage_type in ['取土', '取土开挖']:
                stage_key = f'取土至{solid_level}m处初始土反力'
            elif stage_type == '抽水':
                stage_key = f'抽水至{water_level}m处初始土反力'
            else:
                continue
            zdict[stage_key] = calc_level
            # 存储详细信息
            stage_detail_map[stage_key] = {'土顶标高': solid_level, '水面标高': water_level}
    else:  # 不存在施工阶段
        # 以基坑底标高作为起始标高，水位取基坑底降水标高
        excavate_level = Concrete_Bottom_Level
        drawdown_level = Concrete_Bottom_Level - combo_dict['基坑底降水']
        calc_level = max(excavate_level, drawdown_level)
        stage_key = f'取土至{excavate_level}m处初始土反力'
        zdict[stage_key] = calc_level
        stage_detail_map[stage_key] = {'土顶标高': excavate_level, '水面标高': drawdown_level}

    print('初始土反力计算工况信息:')
    for kz, z in zdict.items():
        detail = stage_detail_map[kz]
        print(f"  {kz}: 土顶标高={detail['土顶标高']}m, 水面标高={detail['水面标高']}m, 计算起始标高={z}m")

    print('初始土反力计算标高:')
    print(zdict)

    # 获取水位标高信息（stage_detail_map 已在上面构建）
    if not stage_dict:  # 不存在施工阶段
        drawdown_level = Concrete_Bottom_Level - combo_dict['基坑底降水']

    for kz, z in zdict.items():
        layer_levels_zi = dict(filter(lambda item: item[1]['层顶标高'] <= z, Active_Earth_Pressure_layer_levels_dict.items())) # 筛选出满足条件的土层:层顶标高<=z
        key_num = len(list(layer_levels_zi.keys())) # layer_levels_zi 中满足要求的key数
        layer_params_zi = dict(list(Active_Earth_Pressure_layer_params_dict.items())[-key_num:]) # 和 layer_levels_zi 保持一致的key数
        # print(layer_levels_zi)
        # print(layer_params_zi)
        # 获取当前工况对应的水位标高和土顶标高
        if stage_dict:  # 存在施工阶段
            current_water_level = stage_detail_map[kz]['水面标高']
            current_solid_level = stage_detail_map[kz]['土顶标高']
        else:  # 不存在施工阶段
            current_water_level = drawdown_level
            current_solid_level = Concrete_Bottom_Level
        # print(f'\n{kz}:')
        # print(f'  土顶标高={current_solid_level}m, 水面标高={current_water_level}m')
        # 判断土层状态
        for k, v in layer_levels_zi.items():
            if v['层底标高'] >= current_solid_level:
                # 层底标高高于土顶标高的部分，按水来计算
                layer_params_zi[k] = {'土层状态': '含水', '重度': 10, '黏聚力': 0, '内摩擦角': 0, '计算方法': '/'}
            elif v['层底标高'] >= current_water_level:
                # 土层范围内，层底标高高于水面标高的为无水层
                # 根据层底标高查找对应的土层
                for solid_value in solid_layer_dict.values():
                    if v['层底标高'] >= float(solid_value['层底标高']):
                        layer_params_zi[k] = {
                            '土层状态': '无水',
                            '重度': float(solid_value['重度']),
                            '黏聚力': float(solid_value['黏聚力']),
                            '内摩擦角': float(solid_value['内摩擦角']),
                            '计算方法': solid_value['计算方法']
                        }
                        break
            else:
                # 土层范围内，层底标高低于水面标高的为含水层
                # 根据层底标高查找对应的土层
                for solid_value in solid_layer_dict.values():
                    if v['层底标高'] >= float(solid_value['层底标高']):
                        layer_params_zi[k] = {
                            '土层状态': '含水',
                            '重度': float(solid_value['重度']),
                            '黏聚力': float(solid_value['黏聚力']),
                            '内摩擦角': float(solid_value['内摩擦角']),
                            '计算方法': solid_value['计算方法']
                        }
                        break

        stage_name = kz # 施工阶段名称
        layer_levels[stage_name] = layer_levels_zi # 该层围檩或承台下的层标高
        layer_params[stage_name] = layer_params_zi # 该层围檩或承台下的层参数
        layer_loads[stage_name] = {} # 该层围檩或承台下的层荷载
        elem_loads[stage_name] = {}

        # 判断当前工况下需要每层需要增加的应力路径
        if if_consider_solid_stress_path:
            # 计算土的附加应力
            excavation_solid_sigma_k_add = calc_excavation_solid_sigma_k_add(Solid_Top_Level, Water_Top_Level, current_solid_level, solid_layer_dict)
            # 计算水的附加应力
            excavation_water_sigma_k_add = round(10 * (Water_Top_Level - current_water_level),3)
        else:
            excavation_solid_sigma_k_add = 0
            excavation_water_sigma_k_add = 0

        # 计算每一层的分布土反力
        soild_sigma0 = 0
        water_sigma0 = 0
        for key in layer_levels_zi.keys():
            # 层顶标高
            layer_level1 = layer_levels_zi[key]['层顶标高']
            # 层底标高
            layer_level2 = layer_levels_zi[key]['层底标高']
            # 土层状态
            layer_state = layer_params_zi[key]['土层状态']
            # 重度
            solid_layer_Gs = layer_params_zi[key]['重度']
            # 内摩擦角
            solid_layer_fai = layer_params_zi[key]['内摩擦角']
            # 土层Ka
            solid_layer_Ka = cal_Ka(solid_layer_fai)
            # 计算方法
            layer_method = layer_params_zi[key]['计算方法']
            # 主动土压力计算
            if layer_state == '含水':
                # 层顶层底竖向应力标准值——土
                solid_sigma1, solid_sigma2 = cal_sigmak_solid(soild_sigma0, solid_layer_Gs, layer_level1, layer_level2, 10)
                # 层顶层底竖向应力标准值——水
                water_sigma1, water_sigma2 = cal_sigmak_water(water_sigma0, layer_level1, layer_level2, 10)
            else:
                # 层顶层底竖向应力标准值——土
                solid_sigma1, solid_sigma2 = cal_sigmak_solid(soild_sigma0, solid_layer_Gs, layer_level1, layer_level2, 0) # 水重度取0
                # 层顶层底竖向应力标准值——水
                water_sigma1, water_sigma2 = [0, 0]
            soild_sigma0 = solid_sigma2 # 竖向应力更新
            water_sigma0 = water_sigma2 # 竖向应力更新

            # 添加附加应力
            if layer_level2 >= current_solid_level:
                solid_sigma_k_add = 0
            else:
                solid_sigma_k_add = excavation_solid_sigma_k_add
            if layer_level2 >= current_water_level:
                water_sigma_k_add = 0
            else:
                water_sigma_k_add = excavation_water_sigma_k_add
            solid_sigma1 = solid_sigma1 + solid_sigma_k_add
            solid_sigma2 = solid_sigma2 + solid_sigma_k_add
            water_sigma1 = water_sigma1 + water_sigma_k_add
            water_sigma2 = water_sigma2 + water_sigma_k_add
            # 层顶层底主动土压力——土
            soild_pak1, soild_pak2 = cal_ps0_soild(solid_layer_Ka, solid_sigma1, solid_sigma2) # 与水土是否分算无关
            # 层顶层底主动土压力——水
            water_pak1, water_pak2 = cal_pak_water(solid_layer_Ka, water_sigma1, water_sigma2, layer_method) # 已经考虑水土分算带来的影响
            
            layer_loads[stage_name][key] = {'竖向应力标准值':[[round(solid_sigma1,3), round(solid_sigma2,3)], [round(water_sigma1,3), round(water_sigma2,3)]],
                                           '分布土反力':[[round(soild_pak1,3), round(soild_pak2,3)], [round(water_pak1,3), round(water_pak2,3)]],
                                           '层标高':[layer_level1, layer_level2],
            }
            # 土
            if layer_method == '水土合算':
                ps01 = round(soild_pak1,1) + round(water_pak1,1)
                ps02 = round(soild_pak2,1) + round(water_pak2,1)
                elem_loads[stage_name][key] = [non_negative_elem_load([ps01, ps02], [layer_level1, layer_level2])]
            elif layer_method == '水土分算':
                pak_soild = non_negative_elem_load([round(soild_pak1,1), round(soild_pak2,1)], [layer_level1, layer_level2]) # 土的非负主动土压力
                pak_water = [round(water_pak1,1), round(water_pak2,1)]
                elem_loads[stage_name][key] = [pak_soild, [[layer_level1, layer_level2], [0, 1], pak_water]]
            # 纯水 layer_method == '/'
            else:
                pak_water = [round(water_pak1,1), round(water_pak2,1)]
                # 纯水层以 len=2 结构（土为0、水承载）输出，避免被生成mct时误归为"水土合算"而进入土反力组
                elem_loads[stage_name][key] = [non_negative_elem_load([0, 0], [layer_level1, layer_level2]), [[layer_level1, layer_level2], [0, 1], pak_water]]

        # # 打印信息===================================================================================================================
        #     print(f'{kz}工况下{key} 层顶标高:{round(layer_level1,3)}, 层底标高:{round(layer_level2,3)}, 土层状态:{layer_state}')
        #     if if_consider_solid_stress_path:
        #         print(f'{kz}工况下{key} 土层历史竖向应力:{solid_sigma_k_add}, 水层历史竖向应力:{water_sigma_k_add}')
        #     print(f'{kz}工况下{key} 层顶土竖向应力:{round(solid_sigma1,3)}, 层底土竖向应力:{round(solid_sigma2,3)}, 层顶水竖向应力:{round(water_sigma1,3)}, 层底水竖向应力:{round(water_sigma2,3)}')
        #     print(f'{kz}工况下{key} 计算方法:{layer_method}, 层顶主动土压力:{round(soild_pak1,3)}, 层底主动土压力:{round(soild_pak2,3)}, 层顶水压力:{round(water_pak1,3)}, 层底水压力:{round(water_pak2,3)}')
        #     print('')
        # for k, v in elem_loads[stage_name].items():
        #     print(stage_name,k,':',v)
        # print('')
            
    return {'层标高':layer_levels, '层参数':layer_params, '层荷载':layer_loads, '单元荷载':elem_loads}


# 计算土压力和土弹簧 (土层信息表-围堰顶标高-围堰底标高-土层顶标高-水面顶标高-承台底标高-混凝土垫层厚度-混凝土封底厚度-主动土压力地面附加应力-被动土压力地面附加应力)
def Cal_Pa_Ps0_Ks(Steel_Sheet_Pile_CofferDam_Load_dict, Pile_SEC_info_dict, plug_zlst, circle_beam_zlst, Waler_Strut_Z, 
                  Sheet_Pile_typevar, SKGGZ_Gap, sigma_k_dict, Drawdown_height, Concrete_Plug_thinckness, LayerThickness, 
                  CofferDam_Top_Level, CofferDam_But_Level, Concrete_Bottom_Level, Solid_Top_Level, Water_Top_Level, 
                  Cap_Bottom_Level, stage_dict, combo_dict, if_consider_solid_stress_path, add_waler_strut_z, csv_dict
                ):
    """土压力/土反力/土弹簧计算的总调度：整理土层标高后依次调用三个计算函数。

    先根据层厚推算各土层顶底标高，再计算土层弹性支承点标高，然后依次调用
    delta_sigma_k_j、Active_Earth_Pressure、Distributed_Soil_Reaction、
    Soil_Spring_Stiffness，并把结果汇总返回。

    Args:
        Steel_Sheet_Pile_CofferDam_Load_dict (dict): 土层信息（含 'Excel' 键）。
        Pile_SEC_info_dict (dict): 支护桩截面信息。
        plug_zlst (list[float]): 封底单元划分标高（m）。
        circle_beam_zlst (list[float]): 圈梁标高（m）。
        Waler_Strut_Z (list[float]): 各层围檩标高（m）。
        Sheet_Pile_typevar (str): 支护类型。
        SKGGZ_Gap (float): 锁扣钢管桩中心间距（mm）。
        sigma_k_dict (dict): 附加荷载信息。
        Drawdown_height (float): 超挖深度（m）。
        Concrete_Plug_thinckness (float): 封底厚度（m）。
        LayerThickness (float): 土弹簧分层厚度（m）。
        CofferDam_Top_Level, CofferDam_But_Level (float): 围堰顶/底标高（m）。
        Concrete_Bottom_Level (float): 基坑底标高（m）。
        Solid_Top_Level, Water_Top_Level (float): 土层顶/水面顶标高（m）。
        Cap_Bottom_Level (float): 承台底标高（m）。
        stage_dict (dict): 施工阶段字典。
        combo_dict (dict): 整体模型荷载组合参数。
        if_consider_solid_stress_path (bool): 是否考虑土的应力路径。
        add_waler_strut_z (list[list]): 辅助加撑的 [层号, 标高]。
        csv_dict (dict): CSV 数据缓存，就地写入 'solid_layerlst'。

    Returns:
        tuple: (Pa_dicts, Ps0_dict, Ks_dicts)，分别为主动土压力、分布土反力、
            土弹簧刚度结果字典，结构见对应函数。

    Example:
        >>> Pa, Ps0, Ks = Cal_Pa_Ps0_Ks(load_dict, pile_sec, [], [], [-0.5], ...)
    """
    spring_zdict = {} # 弹性支承点标高字典
    solid_layer_dict = copy.deepcopy(Steel_Sheet_Pile_CofferDam_Load_dict['Excel']) # excel表格读取结果
    # 所有土层的顶底标高
    # 所有土层的顶底标高
    solid_height = [float(items['层厚']) for items in solid_layer_dict.values()]
    solid_level = [Solid_Top_Level] + [(Solid_Top_Level - sum(abs(x) for x in solid_height[:(i + 1)])) for i in range(len(solid_height))]
    for i, level in enumerate(solid_level[1:]): # 把层底标高更新进 solid_layer_dict 中
        solid_layer_dict[f'第{i+1}层土']['层底标高'] = round(level,3)
    for i, level in enumerate(solid_level[:-1]): # 把层顶标高更新进 solid_layer_dict 中
        solid_layer_dict[f'第{i+1}层土']['层顶标高'] = round(level, 3)
    csv_dict['solid_layerlst'] = solid_level[:-1]

    # 计算土层弹性支承点标高
    # 当存在施工阶段时：根据取土工况的土顶标高计算土层分层，key为"取土至xm处土层弹性支承点标高"
    #   - 陆地钢板桩围堰：筛选工况类型为"取土开挖/设置垫层/设置封底"
    #   - 深水钢板桩围堰：筛选工况类型为"取土/抽水/设置垫层/设置封底"
    #   - 获取所有取土工况的土顶标高，去重后从大到小排序，分别计算土层中点坐标
    # 当不存在施工阶段时：仅计算基坑底以下的土层分层，key为"基底以下土层弹性支承点标高"
    print('\n土层弹性支承点标高计算结果:')
    spring_zdict = {}
    excavate_stages = []
    excavate_solid_levels = []
    excavate_water_levels = []
    if stage_dict:  # 存在施工阶段
        # 筛选取土工况
        if Solid_Top_Level >= Water_Top_Level:
            # 陆地钢板桩围堰
            target_types = ['取土']
        else:
            # 深水钢板桩围堰
            target_types = ['取土', '抽水']
        print('取土工况信息:')
        for stage_num, stage_info in stage_dict.items():
            # 筛选取土工况
            if stage_info['工况类型'] in target_types:
                # 收集对应施工阶段的土顶标高和水面标高
                excavate_solid_levels.append(float(stage_info['基坑内土顶标高']))
                excavate_water_levels.append(float(stage_info['基坑内水面标高']))
                excavate_stages.append({
                    '工况序号': stage_num,
                    '工况类型': stage_info['工况类型'],
                    '基坑内土顶标高': float(stage_info['基坑内土顶标高']),
                    '基坑内水面标高': float(stage_info['基坑内水面标高'])
                })
                print(f"  工况{stage_num}({stage_info['工况类型']}): 土顶标高={stage_info['基坑内土顶标高']}m, 水面标高={stage_info['基坑内水面标高']}m")
        # 去重并从大到小排序
        excavate_solid_levels = sorted(set(excavate_solid_levels), reverse=True)
        excavate_water_levels = sorted(set(excavate_water_levels), reverse=True)
        print(f"\n施工阶段基坑内土顶标高(去重排序): {excavate_solid_levels}")
        print(f"施工阶段基坑内水面标高(去重排序): {excavate_water_levels}")
        # 获取指定工况下的基坑内土顶标高，去重后从大到小排序
        excavate_stage_solid_levels = sorted(set([stage['基坑内土顶标高'] for stage in excavate_stages]), reverse=True)
        # 根据每个取土标高计算土层分层
        for level in excavate_stage_solid_levels:
            Divide_Solid_within_LayerThickness_reslst = Divide_Solid_within_LayerThickness(level, CofferDam_But_Level, LayerThickness, LayerThickness / 2)
            Solid_Middlepoint_inLayerThickness_reslst = [round((Divide_Solid_within_LayerThickness_reslst[i] + Divide_Solid_within_LayerThickness_reslst[i + 1]) / 2, 3) for i in range(len(Divide_Solid_within_LayerThickness_reslst) - 1)]
            spring_zdict[f'取土至{level}m处土层弹性支承点标高'] = Solid_Middlepoint_inLayerThickness_reslst
            print(f"  取土至{level}m处土层划分: {Divide_Solid_within_LayerThickness_reslst}")
            print(f"  取土至{level}m处土层中点: {Solid_Middlepoint_inLayerThickness_reslst}")
    else:  # 不存在施工阶段，仅计算基坑底以下
        level = Concrete_Bottom_Level
        Divide_Solid_within_LayerThickness_reslst = Divide_Solid_within_LayerThickness(level, CofferDam_But_Level, LayerThickness, LayerThickness / 2)
        Solid_Middlepoint_inLayerThickness_reslst = [round((Divide_Solid_within_LayerThickness_reslst[i] + Divide_Solid_within_LayerThickness_reslst[i + 1]) / 2, 3) for i in range(len(Divide_Solid_within_LayerThickness_reslst) - 1)]
        spring_zdict[f'取土至{level}m处土层弹性支承点标高'] = Solid_Middlepoint_inLayerThickness_reslst
        print(f"  取土至{level}m处土层划分: {Divide_Solid_within_LayerThickness_reslst}")
        print(f"  取土至{level}m处土层中点: {Solid_Middlepoint_inLayerThickness_reslst}")

    # 计算所有围檩以下 Drawdown_height m以下的土层标高
    Waler_Drawdown_height_lst = [round(z-Drawdown_height,3) for z in Waler_Strut_Z]

    # 获取土层和结构对应z坐标
    level_zdict = {'围堰顶标高': CofferDam_Top_Level, '围堰底标高': CofferDam_But_Level, '承台底标高': Cap_Bottom_Level, '基坑底标高': Concrete_Bottom_Level, '水面顶标高': Water_Top_Level}
    level_zdict.update({f'第{i + 1}层土顶标高': round(elevation, 3) for i, elevation in enumerate(solid_level) if elevation > CofferDam_But_Level})
    level_zdict = dict(sorted(level_zdict.items(), key=lambda item: item[1], reverse=True))
    print('土层和结构标高计算结果')
    print(level_zdict)

    # 计算当不存在施工阶段时, 基坑底降水的标高
    water_drawdown_height_below_excavation_bottom_lst = []
    if combo_dict:
        water_drawdown_height_below_excavation_bottom_lst.append(Concrete_Bottom_Level - combo_dict['基坑底降水'])

    print('超载信息计算')
    sigma_k_cal_dict = delta_sigma_k_j(sigma_k_dict)
    print('超载信息计算成功')
    print('主动土压力计算')
    Pa_dicts = Active_Earth_Pressure(solid_layer_dict, Solid_Top_Level, Water_Top_Level, CofferDam_But_Level, level_zdict, spring_zdict, Waler_Strut_Z, excavate_solid_levels, excavate_water_levels, plug_zlst, sigma_k_cal_dict, water_drawdown_height_below_excavation_bottom_lst, circle_beam_zlst, add_waler_strut_z)
    print('主动土压力计算成功')
    print('初始土反力计算')
    Ps0_dict = Distributed_Soil_Reaction(Pa_dicts['1']['层标高'], Pa_dicts['1']['层参数'], Concrete_Bottom_Level, Solid_Top_Level, Water_Top_Level, stage_dict, excavate_stages, combo_dict, solid_layer_dict, if_consider_solid_stress_path)
    print('初始土反力计算成功')
    print('土弹簧计算')
    Ks_dicts = Soil_Spring_Stiffness(solid_layer_dict, Solid_Top_Level, Cap_Bottom_Level, Concrete_Bottom_Level, CofferDam_But_Level, spring_zdict, Pile_SEC_info_dict, LayerThickness, Concrete_Plug_thinckness, Sheet_Pile_typevar, SKGGZ_Gap, stage_dict)
    print('土弹簧计算成功')
    return Pa_dicts, Ps0_dict, Ks_dicts


def validate_replace_stage_tracking(
    Waler_position_line_lst,
    elemdict,
    grupdict_elem,
    layer_arrangements,
    replace_stage_tracking,
    assist_replace_dict,
    recognize_Strut_Replace_result_dict,
    Construct_Stage_Line_Data,
    sec_num_name_lst=None
) -> list:
    """验证辅助换撑参数，打印调试信息，检查是否存在混合变更（同时有新增和移除）。

    Args:
        Waler_position_line_lst (list[list[tuple]]): 各层围檩中心线角点。
        elemdict (dict): 单元字典 {结构组名: [单元信息, ...]}。
        grupdict_elem (dict): 结构组 -> 单元号列表。
        layer_arrangements (dict): 各层布置跟踪 {层号: [{'dc': [...], 'xc': [...]}, ...]}。
        replace_stage_tracking (dict): 换撑阶段元素分配跟踪
            {(层号, 换撑号): {'dc_add','dc_remove','xc_add','xc_remove', ...}}。
        assist_replace_dict (dict): 辅助换撑定义。
        recognize_Strut_Replace_result_dict (dict): 换撑几何计算结果。
        Construct_Stage_Line_Data (list[list]): 施工阶段表数据。
        sec_num_name_lst (list|None): 截面号与截面名列表 [[号, 名], ...]。

    Returns:
        list[str]: 错误信息列表；为空表示验证通过。存在同一体系内既有新增又有
            移除时追加"建议分两次换撑"的提示。
    """
    errors = []

    # ==================== 计算要点第1点：辅助换撑工况信息汇总 ====================
    print('\n' + '='*80)
    print('【计算要点第1点】辅助换撑工况信息汇总')
    print('='*80)

    # 1. 打印基本支撑信息（初始对撑和斜撑）
    print('\n--- 1. 基本支撑信息（初始对撑和斜撑） ---')
    for z in range(len(Waler_position_line_lst)):
        _layer_i = z + 1
        _dc_elems = elemdict.get(f'第{_layer_i}层对撑', [])
        _xc_elems = elemdict.get(f'第{_layer_i}层斜撑', [])
        print(f'第{_layer_i}层:')
        print(f'  对撑单元数: {len(_dc_elems)}')
        if _dc_elems:
            print(f'  对撑单元号: {[e[0] for e in _dc_elems]}')
        print(f'  斜撑单元数: {len(_xc_elems)}')
        if _xc_elems:
            print(f'  斜撑单元号: {[e[0] for e in _xc_elems]}')

    # 2. 打印辅助换撑布置信息
    print('\n--- 2. 辅助换撑布置信息（assist_replace_dict） ---')
    if assist_replace_dict:
        for layer_key, layer_forms in assist_replace_dict.items():
            print(f'第{layer_key}层换撑:')
            for form_key, form_data in layer_forms.items():
                print(f'  第{form_key}次换撑:')
                print(f'    对撑截面: {form_data.get("对撑截面", "/")}')
                print(f'    斜撑截面: {form_data.get("斜撑截面", "/")}')
                print(f'    对撑长边布置: {form_data.get("对撑长边布置(m)", "/")}')
                print(f'    对撑短边布置: {form_data.get("对撑短边布置(m)", "/")}')
                print(f'    斜撑长边布置: {form_data.get("斜撑长边布置(m)", "/")}')
                print(f'    斜撑短边布置: {form_data.get("斜撑短边布置(m)", "/")}')
    else:
        print('  无辅助换撑数据')

    # 3. 打印辅助换撑识别结果（位置信息）
    print('\n--- 3. 辅助换撑识别结果（recognize_Strut_Replace_result_dict） ---')
    if recognize_Strut_Replace_result_dict:
        for layer_key, layer_forms in recognize_Strut_Replace_result_dict.items():
            print(f'第{layer_key}层换撑:')
            for form_key, form_data in layer_forms.items():
                print(f'  第{form_key}次换撑:')
                _dc_pts = form_data.get('position_line_duicheng_lst', [])
                _xc_pts = form_data.get('position_line_xiecheng_lst', [])
                print(f'    对撑布置点位数: {len(_dc_pts)}')
                if _dc_pts:
                    print(f'    对撑布置点位: {_dc_pts}')
                print(f'    斜撑布置点位数: {len(_xc_pts)}')
                if _xc_pts:
                    print(f'    斜撑布置点位: {_xc_pts}')
    else:
        print('  无辅助换撑识别结果')

    # 4. 提取所有加撑及辅助换撑工况，并按施工阶段顺序排列
    print('\n--- 4. 所有加撑及辅助换撑工况（按施工阶段顺序排列） ---')

    # 直接构造施工阶段名称列表（不调用外部函数，避免函数未定义的问题）
    def _stage_line_to_midas_name(stage_line):
        """将单行施工阶段表数据转换为 MIDAS 施工阶段名称。

        Args:
            stage_line (list): 施工阶段数据行，首列为工况类型。

        Returns:
            str: 对应的 MIDAS 阶段名，未知类型原样返回。
        """
        stage_name = stage_line[0]
        if stage_name == '加撑':
            return f'施工第{stage_line[3]}道支撑'
        elif stage_name == '辅助换撑':
            index_lst = stage_line[3].split('-')
            return f'第{index_lst[0]}层内支撑体系转换{index_lst[1]}'
        elif stage_name == '拆撑':
            return f'拆除第{stage_line[3]}道支撑'
        elif stage_name == '取土':
            return f'取土至{stage_line[1]}m处'
        elif stage_name == '抽水':
            return f'抽水至{stage_line[2]}m处'
        elif stage_name == '封底':
            return '封底'
        elif stage_name == '辅助加圈梁':
            return f'在{stage_line[1]}m处设置圈梁'
        return stage_name

    # 构造施工阶段名称列表
    Midas_Stage_Name_lst_temp = [_stage_line_to_midas_name(sl) for sl in Construct_Stage_Line_Data]
    print(f'施工阶段总数: {len(Midas_Stage_Name_lst_temp)}')
    print(f'施工阶段列表: {Midas_Stage_Name_lst_temp}')

    # 提取加撑和辅助换撑工况
    add_strut_stages = []  # 加撑工况
    replace_strut_stages = []  # 辅助换撑工况

    for stage_idx, stage_line in enumerate(Construct_Stage_Line_Data):
        stage_name = stage_line[0]
        midas_name = Midas_Stage_Name_lst_temp[stage_idx] if stage_idx < len(Midas_Stage_Name_lst_temp) else ''

        if stage_name == '加撑':
            _num = stage_line[3] if len(stage_line) > 3 else '/'
            add_strut_stages.append({
                '施工阶段序号': stage_idx + 1,
                '工况类型': '加撑',
                '支撑层数': _num,
                'Midas名称': midas_name,
                '原始数据': stage_line,
            })
        elif stage_name == '辅助换撑':
            _layer_form = stage_line[3] if len(stage_line) > 3 else '/'
            parts = _layer_form.split('-') if '-' in str(_layer_form) else [_layer_form, '1']
            _layer_i = parts[0] if len(parts) > 0 else '/'
            _form_j = parts[1] if len(parts) > 1 else '1'
            replace_strut_stages.append({
                '施工阶段序号': stage_idx + 1,
                '工况类型': '辅助换撑',
                '支撑层数': _layer_i,
                '换撑编号': _form_j,
                'Midas名称': midas_name,
                '原始数据': stage_line,
            })

    print(f'\n【加撑工况】共{len(add_strut_stages)}个:')
    for s in add_strut_stages:
        print(f'  施工阶段{s["施工阶段序号"]}: {s["Midas名称"]} (第{s["支撑层数"]}层支撑)')

    print(f'\n【辅助换撑工况】共{len(replace_strut_stages)}个:')
    for s in replace_strut_stages:
        print(f'  施工阶段{s["施工阶段序号"]}: {s["Midas名称"]} (第{s["支撑层数"]}层第{s["换撑编号"]}次换撑)')

    # 5. 按施工阶段顺序排列所有加撑及辅助换撑工况
    print('\n--- 5. 按施工阶段顺序排列的所有加撑及辅助换撑工况 ---')
    all_strut_change_stages = add_strut_stages + replace_strut_stages
    all_strut_change_stages.sort(key=lambda x: x['施工阶段序号'])

    print(f'总工况数: {len(all_strut_change_stages)}')
    for idx, s in enumerate(all_strut_change_stages):
        if s['工况类型'] == '加撑':
            print(f'  {idx+1}. 施工阶段{s["施工阶段序号"]}: {s["工况类型"]} - 第{s["支撑层数"]}层支撑 ({s["Midas名称"]})')
        else:
            print(f'  {idx+1}. 施工阶段{s["施工阶段序号"]}: {s["工况类型"]} - 第{s["支撑层数"]}层第{s["换撑编号"]}次换撑 ({s["Midas名称"]})')

    print('\n' + '='*80)
    print('【计算要点第1点】信息汇总完成')
    print('='*80 + '\n')

    # ==================== 计算要点第2点(2)：各换撑工况的对撑和斜撑单元号统计 ====================
    print('\n' + '='*80)
    print('【计算要点第2点(2)】各换撑工况的对撑和斜撑单元号统计')
    print('='*80)

    # 打印各层布置跟踪
    print('\n--- 1. 各层布置跟踪 (layer_arrangements) ---')
    for _layer_i, _arrs in layer_arrangements.items():
        print(f'第{_layer_i}层:')
        for _arr_idx, _arr in enumerate(_arrs):
            _dc_elems = _arr.get('dc', [])
            _xc_elems = _arr.get('xc', [])
            if _arr_idx == 0:
                _label = '原始布置'
            else:
                _label = f'第{_arr_idx}次换撑后'
            print(f'  {_label}:')
            print(f'    对撑单元数: {len(_dc_elems)}')
            print(f'    对撑单元号: {_dc_elems}')
            print(f'    斜撑单元数: {len(_xc_elems)}')
            print(f'    斜撑单元号: {_xc_elems}')

    # 打印各换撑阶段的元素分配跟踪，并验证是否存在混合变更
    print('\n--- 2. 各换撑阶段元素分配跟踪 (replace_stage_tracking) ---')
    for (_layer_i, _form_j), _tracking in sorted(replace_stage_tracking.items()):
        print(f'第{_layer_i}层第{_form_j}次换撑:')
        _dc_add = _tracking.get('dc_add', [])
        _dc_remove = _tracking.get('dc_remove', [])
        _xc_add = _tracking.get('xc_add', [])
        _xc_remove = _tracking.get('xc_remove', [])
        _dc_deact = _tracking.get('dc_deactivate_original', False)
        _xc_deact = _tracking.get('xc_deactivate_original', False)
        _grp_dc = _tracking.get('grp_dc', '')
        _grp_xc = _tracking.get('grp_xc', '')
        print(f'  结构组名: 对撑="{_grp_dc}", 斜撑="{_grp_xc}"')
        print(f'  对撑:')
        print(f'    新增单元数: {len(_dc_add)}')
        print(f'    新增单元号: {_dc_add}')
        print(f'    移除单元数: {len(_dc_remove)}')
        print(f'    移除单元号: {_dc_remove}')
        print(f'    钝化原始组: {_dc_deact}')
        print(f'  斜撑:')
        print(f'    新增单元数: {len(_xc_add)}')
        print(f'    新增单元号: {_xc_add}')
        print(f'    移除单元数: {len(_xc_remove)}')
        print(f'    移除单元号: {_xc_remove}')
        print(f'    钝化原始组: {_xc_deact}')

        # 判断工况类型
        if _dc_deact or _xc_deact:
            _case_desc = '类型消失（截面为/且之前有元素）'
        elif _dc_add and not _dc_remove and _xc_add and not _xc_remove:
            _case_desc = '情况1（增加）：对撑和斜撑都增加'
        elif _dc_remove and not _dc_add and _xc_remove and not _xc_add:
            _case_desc = '情况2（减少）：对撑和斜撑都减少'
        elif not _dc_add and not _dc_remove and not _xc_add and not _xc_remove:
            _case_desc = '情况3（不变）：布置不变'
        elif (_dc_add and _dc_remove) or (_xc_add and _xc_remove):
            _case_desc = '情况4（全换）：既有新增又有减少（需分别处理对撑和斜撑）'
        elif _dc_add and not _dc_remove:
            _case_desc = '情况1变体：对撑增加，斜撑不变/减少'
        elif _dc_remove and not _dc_add:
            _case_desc = '情况2变体：对撑减少，斜撑不变/增加'
        else:
            _case_desc = '混合情况'
        print(f'  工况类型: {_case_desc}')

        # 验证：检查是否存在混合变更（同时有新增和移除）
        if _dc_add and _dc_remove:
            errors.append(f'第{_layer_i}层第{_form_j}次换撑：对撑存在不合理体系转换，建议：分两次换撑进行体系转换')
        if _xc_add and _xc_remove:
            errors.append(f'第{_layer_i}层第{_form_j}次换撑：斜撑存在不合理体系转换，建议：分两次换撑进行体系转换')

    # 打印结构组中的换撑组
    print('\n--- 3. 换撑结构组汇总 ---')
    for _grp_name, _elems in grupdict_elem.items():
        if '换撑' in _grp_name:
            print(f'  {_grp_name}: 单元数={len(_elems)}, 单元号={_elems}')

    print('='*80 + '\n')

    return errors


# 根据钢板桩角点、钢板桩类型、围檩截面、内支撑位置、内支撑截面生成节点和单元, 根据土压力和土弹簧生成荷载和边界
def make_node_elem_load_boundary(
        Sheet_Pile_namevar, Pile_SEC_info_dict, CofferDam_Top_Level, CofferDam_Bottom_Level, Cap_Bottom_Level, Method_for_Pressures_var, Waler_Strut_Z, none_construction_stage_check_var, construction_stage_check_var,
        Waler_section_dict, Walers_Strut_selected_dict, recognize_Cap_result_dict, recognize_Bracket_result_dict, recognize_Waler_result_dict, recognize_Strut_result_dict, recognize_Strut_Replace_result_dict, Steel_Sheet_Pile_CofferDam_Load_dict,
        SECTION_lst, selected_Waler_sectionlst, Pa_dicts, Ps0_dicts, Ks_dicts, Drawdown_height, Concrete_Plug_thinckness, cofferdam_cons, Load_Group_namelst, Concrete_Bottom_Level, BNDR_GROUP_lst, stage_dict,
        Load_Combo_Line_Data, Construct_Stage_Line_Data, plug_zlst, circle_beam_zlst, Sheet_Pile_typevar, SKGGZ_Gap, sheet_elasticlink_dict, cofferdam_elasticlink_form_dict, assist_add_dict, assist_replace_dict, add_waler_strut_z,
        material_dict=None, Sheet_Pile_materialvar=None,
        ):
    """生成围堰模型的节点、单元、结构组、约束、弹性连接、荷载与施工阶段。

    依据几何计算结果（支护桩点、围檩角点、内支撑端点、牛腿点）建立节点和
    梁单元；依据土压力/土反力/土弹簧生成梁单元荷载、节点弹性支承与边界组；
    依据施工阶段表生成各阶段的激活/钝化及辅助换撑的结构组。

    Args:
        各参数含义与 Steel_Sheet_Pile_CofferDam_on_submit_MCT / mct_output 一致；
        另有：
        SECTION_lst (list[str]): SECTION() 输出的带编号截面命令。
        selected_Waler_sectionlst (list[str]): 选用的围檩截面名。
        Pa_dicts, Ps0_dicts, Ks_dicts (dict): 主动土压力/分布土反力/土弹簧结果。
        Concrete_Plug_thinckness (float): 封底厚度（m）。
        cofferdam_cons (str): 桩底边界编码。
        Load_Group_namelst (list[str]): 荷载组名称列表。
        BNDR_GROUP_lst (list[str]): 边界组名称列表。
        plug_zlst, circle_beam_zlst (list[float]): 封底/圈梁标高。
        sheet_elasticlink_dict, cofferdam_elasticlink_form_dict (dict): 连接刚度/形式。
        assist_add_dict, assist_replace_dict (dict): 辅助加撑/换撑定义。
        add_waler_strut_z (list[list]): 辅助加撑 [层号, 标高]。
        material_dict (dict|None): 材料字典，用于材质名到材质号的映射。
        Sheet_Pile_materialvar (str|None): 支护桩材质名。

    Returns:
        tuple: (node_lst, elem_lst, grup_lst, cons_lst, elink_lst, load_grupdict,
                spring_lst, stage_lst, loadcomb_lst, BNDR_GROUP_lst)
            依次为节点、单元、结构组、约束、弹性连接、梁单元荷载、节点弹性
            支承、施工阶段、荷载组合、边界组的 MCT 命令行列表。
    """
    # 材质名称 → 材质号（material_dict 的 key）
    mat_name_to_key = {}
    if material_dict:
        mat_name_to_key = {str(v.get('name', '')).strip(): k for k, v in material_dict.items()}
    pile_mat_no = mat_name_to_key.get(str(Sheet_Pile_materialvar).strip(), 1) if Sheet_Pile_materialvar else 1

    # 所有的截面及截面号
    # print(SECTION_lst)
    sec_num_name_lst = [] # 截面号及对应的截面名称
    for section_mctstring in SECTION_lst:
        secmct_splitlst = section_mctstring.split(', ')
        sec_num_name_lst.append([secmct_splitlst[0], secmct_splitlst[2]])
    print(sec_num_name_lst)
    # 承台及钢板桩参数计算结果
    O_point = recognize_Cap_result_dict['O_Point']
    OX_point = (O_point[0]+1, O_point[1], O_point[2])
    Pile_ptlst = recognize_Cap_result_dict['Pile_ptlst']
    # 围檩参数计算结果
    Waler_position_line_lst = recognize_Waler_result_dict['position_line_lst']
    # 内支撑参数计算结果
    Strut_position_line_lst_duicheng = recognize_Strut_result_dict['position_line_duicheng_lst']
    Strut_position_line_lst_xiecheng = recognize_Strut_result_dict['position_line_xiecheng_lst']
    # 封底厚度z坐标范围
    concrete_plug_zlst = [Cap_Bottom_Level, Cap_Bottom_Level-Concrete_Plug_thinckness]
    # Z坐标上的点
    Pa_Z = []
    Padict_Z = [[value['层顶标高'], value['层底标高']] for value in list(Pa_dicts['1']['层标高'].values())] # 主动土压力相关的Z
    for x,y in Padict_Z:
       Pa_Z.extend((x,y)) 

    # 加撑及辅助加撑工况下的所有z坐标
    Waler_Strut_Z_dict = {}
    for i in range(len(Waler_Strut_Z)):
        z = Waler_Strut_Z[i]
        Waler_Strut_Z_dict[z] = {
            'index':i+1,
            'if_add':False,
        }
    for i, z in add_waler_strut_z:
        Waler_Strut_Z_dict[z] = {
            'index':i,
            'if_add':True,
        }
    print('加撑及辅助加撑工况下的所有z坐标信息汇总')
    for k, v in Waler_Strut_Z_dict.items():
        print(k, v)
    waler_strut_z = list(Waler_Strut_Z_dict.keys())
    print(waler_strut_z)
    
    # print('主动土压力相关土层z坐标:', Pa_Z)
    Z_Points_appendlst = list(sorted(set([CofferDam_Top_Level] + Pa_Z), key = lambda x : x, reverse = True))
    # Steel_Sheet_Pile_dict['ZPoints'] = Z_Points_appendlst
    # 生成钢板桩节点和单元
    node_i = 1
    Pile_i = 1
    nodedict = {} # 第 Pile_i 排钢板桩的节点
    corner_pile_num_lst = [] # 储存角桩（用于锁扣钢管桩的情况）
    elem_angledict = {} # 第 Pile_i 排钢板桩的单元旋转角度
    grupdict_elem = {} # 结构组单元号字典
    grupdict_node = {} # 结构组节点号字典
    load_grupdict = {} # 荷载组字典
    section_dict = {} # 第 Pile_i 排钢板桩的截面号
    Pile_elasticlink_dict = {} # 钢板桩与围檩的弹性连接表
    Pile_elasticlink_dict_with_bracket = {} # 钢板桩与围檩的弹性连接表中与牛腿共节点的点
    Constraint_lst = {'桩底约束':[], '封底':[], '圈梁':[]} # 节点一般支承表/弹性支承表
    Cap_Cons_dict = {}
    spring_dict = {} # 节点弹性支承表
    Pile_name = Sheet_Pile_typevar
    # a = '钢板桩'
    # print(Ks_dicts)
    # 节点
    grupdict_node[Pile_name] = []
    for ptlst_i in range(len(Pile_ptlst)):
        ptlst = Pile_ptlst[ptlst_i]
        start_angle = calculate_full_angle(O_point, OX_point, ptlst[0]) # 起始角度
        elem_angle = (int(start_angle/90)+1)*90 # 该条边对应的单元旋转角度
        Pile_elasticlink_dict[f'第{ptlst_i+1}条边'] = {}
        Pile_elasticlink_dict_with_bracket[f'第{ptlst_i+1}条边'] = {}
        for i in range(len(ptlst)):
            nodedict[f'第{Pile_i}排{Pile_name}'] = []
            elem_angledict[f'第{Pile_i}排{Pile_name}'] = []
            section_dict[f'第{Pile_i}排{Pile_name}'] = []
            Pile_elasticlink_dict[f'第{ptlst_i+1}条边'][f'第{Pile_i}排{Pile_name}'] = []
            # Pile_elasticlink_dict_with_bracket[f'第{ptlst_i+1}条边'][f'第{Pile_i}排{Pile_name}'] = []
            if Sheet_Pile_typevar == '钢板桩':
                # 钢板桩的截面尺寸
                Sheet_Pile_Width = Pile_SEC_info_dict['SEC'][1]
                if i != 0 and i != len(ptlst)-1:
                    pt = ptlst[i]
                    elem_width = Sheet_Pile_Width
                elif i == 0:
                    pt = find_circle_line_intersection(ptlst[i+1], ptlst[i], Sheet_Pile_Width/4, 'ON')# 以点1点2为直线以点1为圆心求圆与直线的角点
                    elem_width = Sheet_Pile_Width/2
                elif i == len(ptlst)-1:
                    pt = find_circle_line_intersection(ptlst[i-1], ptlst[i], Sheet_Pile_Width/4, 'ON')# 以点1点2为直线以点1为圆心求圆与直线的角点
                    elem_width = Sheet_Pile_Width/2
            elif Sheet_Pile_typevar == '锁扣钢管桩':
                if i != 0: # 为避免角桩上建立重复节点, 第i条边的第1根角桩不建立
                    pt = ptlst[i]
                    elem_width = SKGGZ_Gap
                else:
                    pt = None
            if pt != None:
                bracket_nodelst = []
                for pt_z in Z_Points_appendlst:
                    node = [node_i, round(pt[0], 3), round(pt[1], 3), round(pt_z*1000, 3)]
                    # 围檩的弹性连接点
                    if pt_z in waler_strut_z:
                        Pile_elasticlink_dict[f'第{ptlst_i+1}条边'][f'第{Pile_i}排{Pile_name}'].append(node)
                        # 判断是否是牛腿的点
                        bracket_ptlst = recognize_Bracket_result_dict['bracket_points_by_side'][ptlst_i]
                        if (pt[0], pt[1], 0) in bracket_ptlst:
                            bracket_nodelst.append(node)  
                    # 桩底约束的一般支承点
                    if pt_z == Z_Points_appendlst[-1]:
                        Constraint_lst['桩底约束'].append(node_i)
                    # 封底    
                    if pt_z in plug_zlst and Concrete_Plug_thinckness != 0:
                        Constraint_lst['封底'].append([node_i, SPRING_COMP_Direction(elem_angle)])
                    # 圈梁
                    if pt_z in circle_beam_zlst:
                        Constraint_lst['圈梁'].append([node_i, pt_z, SPRING_COMP_Direction(elem_angle)])
                    # 节点弹性支承
                    spinglst = [(k, sublist[-1]) for k, v in Ks_dicts.items() for sublist in v if sublist[0] == pt_z]
                    for key, spring_value in spinglst:
                        if key not in list(spring_dict.keys()):
                            spring_dict[key] = {}
                        if spring_value not in list(spring_dict[key].keys()):
                            spring_dict[key][spring_value] = []
                        spring_dict[key][spring_value].append(node_i)
                    nodedict[f'第{Pile_i}排{Pile_name}'].append(node)
                    grupdict_node[Pile_name].append(node_i)
                    node_i += 1
                elem_angledict[f'第{Pile_i}排{Pile_name}'].append(elem_angle)
                section_dict[f'第{Pile_i}排{Pile_name}'].append(elem_width)
                corner_pile_num_lst.append(f'第{Pile_i}排{Pile_name}') if i == len(ptlst)-1 else None 
                if bracket_nodelst != []:
                    Pile_elasticlink_dict_with_bracket[f'第{ptlst_i+1}条边'][f'第{Pile_i}排{Pile_name}'] = bracket_nodelst
                Pile_i += 1
    print(f'{Pile_name}节点成功生成')
    
    # 单元
    elem_i = 1
    elemdict = {}
    grupdict_elem[Pile_name] = [] # 钢板桩结构组单元号列表

    distributed_soil_reaction_dict = Ps0_dicts['单元荷载'] # 一根钢板桩上的分布土反力计算结果
    if Method_for_Pressures_var == '水土合算':
        load_grupdict['基坑外主动土压力'] = []
    else:
        load_grupdict['基坑外主动土压力'] = []
        load_grupdict['基坑外水压力'] = []
    for keyi in list(distributed_soil_reaction_dict.keys()):
        if Method_for_Pressures_var == '水土合算':
            load_grupdict[keyi] = []
        else:
            load_grupdict[keyi] = []
            load_grupdict[keyi.replace('初始土反力', '水压力')] = []
    load_ndict = {90:[1,-1],180:[1,1],270:[-1,1],360:[-1,-1],0:[1,-1]} # 荷载随单元旋转角度的作用方向
    face_ndict = {90:'1',180:'2',270:'3',360:'4',0:'4'} # 判断面数
    for key, value in nodedict.items():
        elemdict[key] = [] # key == 第i排钢板桩
        elem_width = section_dict[key][0] # 单元宽度
        elem_load_scale = elem_width/1000 # 单位宽度对应下的荷载比例系数
        elem_angle = elem_angledict[key][0] # 第i根桩的单元旋转角度

        face_num = face_ndict[elem_angle] # 第i根桩所在的面数
        load_nx, load_ny = load_ndict[elem_angle] # 荷载随单元旋转角度的作用方向
        active_earth_pressure_dict = Pa_dicts[face_num]['单元荷载'] # 根据面数判断第i面钢板桩上的主动土压力计算结果
        
        # 判断荷载方向(绝对坐标轴)
        if Sheet_Pile_typevar == '钢板桩':
            load_DIR_lst = ['GY'] if elem_angle == 90 or elem_angle == 270 else ['GX'] # 荷载方向
        elif Sheet_Pile_typevar == '锁扣钢管桩':
            if key in corner_pile_num_lst:
                load_DIR_lst = ['GX', 'GY'] # 角桩荷载方向
                elem_load_scale = elem_load_scale * 0.5 # 角桩荷载比例系数
            else:
                load_DIR_lst = ['GY'] if elem_angle == 90 or elem_angle == 270 else ['GX'] # 荷载方向
        # 判断截面号
        for sec_num, sec_name in sec_num_name_lst:
            if Sheet_Pile_namevar in sec_name and Sheet_Pile_typevar == '钢板桩':
                lasen_width = float(re.findall(r'\d+\.?\d*', sec_name)[0])
                if abs(lasen_width-elem_width) < 0.5:
                    elem_secnum = sec_num # 单元截面号
            elif Sheet_Pile_typevar == '锁扣钢管桩':
                if Sheet_Pile_typevar in sec_name:
                    elem_secnum = sec_num # 单元截面号
        # 单元mct信息
        for i in range(len(value)-1):
            elem = [elem_i, 'BEAM', str(pile_mat_no), elem_secnum, value[i], value[i+1], elem_angle, '0'] # 单元信息
            grupdict_elem[Pile_name].append(elem_i) # 结构组信息
            elemdict[key].append(elem)
            # 主动土压力
            try:
                active_earth_pressure_lst = active_earth_pressure_dict[f'第{i}层']
                if len(active_earth_pressure_lst) == 1: # 水土合算
                    if active_earth_pressure_lst[0][2] != [0, 0]: # 存在荷载
                        position_lst, load_lst = elem_load_lst_trans(active_earth_pressure_lst[0][1], active_earth_pressure_lst[0][2])
                        for load_dir in load_DIR_lst:
                            load_n = load_nx if 'X' in load_dir else load_ny
                            elem_load = '{}, LINE, UNILOAD, {}, NO , NO, LY, , , , {}, {}, {}, {}, {}, {}, {}, {}, {}, NO, 0, 0, NO,'.format(
                                elem_i, load_dir, position_lst[0], round(load_lst[0]*load_n*elem_load_scale,3), position_lst[1], round(load_lst[1]*load_n*elem_load_scale,3), position_lst[2], round(load_lst[2]*load_n*elem_load_scale,3), position_lst[3], round(load_lst[3]*load_n*elem_load_scale,3), '基坑外主动土压力'
                                ) # 荷载信息
                            load_grupdict['基坑外主动土压力'].append(elem_load)
                else: # 水土分算
                    soild_active_earth_pressure_lst, water_active_earth_pressure_lst = active_earth_pressure_lst
                    if Method_for_Pressures_var == '水土分算':
                        if soild_active_earth_pressure_lst[2] != [0, 0]: # 存在土荷载
                            position_lst, load_lst = elem_load_lst_trans(soild_active_earth_pressure_lst[1], soild_active_earth_pressure_lst[2])
                            for load_dir in load_DIR_lst:
                                load_n = load_nx if 'X' in load_dir else load_ny
                                elem_load = '{}, LINE, UNILOAD, {}, NO , NO, LY, , , , {}, {}, {}, {}, {}, {}, {}, {}, {}, NO, 0, 0, NO,'.format(
                                    elem_i, load_dir, position_lst[0], round(load_lst[0]*load_n*elem_load_scale,3), position_lst[1], round(load_lst[1]*load_n*elem_load_scale,3), position_lst[2], round(load_lst[2]*load_n*elem_load_scale,3), position_lst[3], round(load_lst[3]*load_n*elem_load_scale,3), '基坑外主动土压力'
                                    ) # 荷载信息
                                load_grupdict['基坑外主动土压力'].append(elem_load)
                        if water_active_earth_pressure_lst[2] != [0, 0]: # 存在水荷载
                            position_lst, load_lst = elem_load_lst_trans(water_active_earth_pressure_lst[1], water_active_earth_pressure_lst[2])
                            for load_dir in load_DIR_lst:
                                load_n = load_nx if 'X' in load_dir else load_ny
                                elem_load = '{}, LINE, UNILOAD, {}, NO , NO, LY, , , , {}, {}, {}, {}, {}, {}, {}, {}, {}, NO, 0, 0, NO,'.format(
                                        elem_i, load_dir, position_lst[0], round(load_lst[0]*load_n*elem_load_scale,3), position_lst[1], round(load_lst[1]*load_n*elem_load_scale,3), position_lst[2], round(load_lst[2]*load_n*elem_load_scale,3), position_lst[3], round(load_lst[3]*load_n*elem_load_scale,3), '基坑外水压力'
                                        ) # 荷载信息
                                load_grupdict['基坑外水压力'].append(elem_load)
                    elif Method_for_Pressures_var == '水土合算':
                        position_lst, load_lst = add_continuous_beam_elem_loads(elem_load_lst_trans(soild_active_earth_pressure_lst[1], soild_active_earth_pressure_lst[2]), elem_load_lst_trans(water_active_earth_pressure_lst[1], water_active_earth_pressure_lst[2]))
                        for load_dir in load_DIR_lst:
                            load_n = load_nx if 'X' in load_dir else load_ny
                            elem_load = '{}, LINE, UNILOAD, {}, NO , NO, LY, , , , {}, {}, {}, {}, {}, {}, {}, {}, {}, NO, 0, 0, NO,'.format(
                                elem_i, load_dir, position_lst[0], round(load_lst[0]*load_n*elem_load_scale,3), position_lst[1], round(load_lst[1]*load_n*elem_load_scale,3), position_lst[2], round(load_lst[2]*load_n*elem_load_scale,3), position_lst[3], round(load_lst[3]*load_n*elem_load_scale,3), '基坑外主动土压力'
                                ) # 荷载信息
                            load_grupdict['基坑外主动土压力'].append(elem_load)
            except:
                pass
            # 分布土压力
            for keyii, valueii in distributed_soil_reaction_dict.items():
                try:
                # if i == 23: # 测试用
                    distributed_soil_reaction_lst = valueii[f'第{i}层']
                    if len(distributed_soil_reaction_lst) == 1: # 水土合算
                        if distributed_soil_reaction_lst[0][2] != [0, 0]: # 存在荷载
                            position_lst, load_lst = elem_load_lst_trans(distributed_soil_reaction_lst[0][1], distributed_soil_reaction_lst[0][2])
                            for load_dir in load_DIR_lst:
                                load_n = load_nx if 'X' in load_dir else load_ny
                                elem_load = '{}, LINE, UNILOAD, {}, NO , NO, LY, , , , {}, {}, {}, {}, {}, {}, {}, {}, {}, NO, 0, 0, NO,'.format(
                                    elem_i, load_dir, position_lst[0], round(load_lst[0]*load_n*-1*elem_load_scale,3), position_lst[1], round(load_lst[1]*load_n*-1*elem_load_scale,3), position_lst[2], round(load_lst[2]*load_n*-1*elem_load_scale,3), position_lst[3], round(load_lst[3]*load_n*-1*elem_load_scale,3), keyii
                                    ) # 荷载信息
                                load_grupdict[keyii].append(elem_load)
                    else: # 水土分算
                        soild_distributed_soil_reaction_lst, water_distributed_soil_reaction_lst = distributed_soil_reaction_lst
                        if Method_for_Pressures_var == '水土分算':
                            if soild_distributed_soil_reaction_lst[2] != [0, 0]: # 存在土荷载
                                position_lst, load_lst = elem_load_lst_trans(soild_distributed_soil_reaction_lst[1], soild_distributed_soil_reaction_lst[2])
                                for load_dir in load_DIR_lst:
                                    load_n = load_nx if 'X' in load_dir else load_ny
                                    elem_load = '{}, LINE, UNILOAD, {}, NO , NO, LY, , , , {}, {}, {}, {}, {}, {}, {}, {}, {}, NO, 0, 0, NO,'.format(
                                        elem_i, load_dir, position_lst[0], round(load_lst[0]*load_n*-1*elem_load_scale,3), position_lst[1], round(load_lst[1]*load_n*-1*elem_load_scale,3), position_lst[2], round(load_lst[2]*load_n*-1*elem_load_scale,3), position_lst[3], round(load_lst[3]*load_n*-1*elem_load_scale,3), keyii
                                        ) # 荷载信息
                                    load_grupdict[keyii].append(elem_load)
                            if water_distributed_soil_reaction_lst[2] != [0, 0]: # 存在水荷载
                                position_lst, load_lst = elem_load_lst_trans(water_distributed_soil_reaction_lst[1], water_distributed_soil_reaction_lst[2])
                                for load_dir in load_DIR_lst:
                                    load_n = load_nx if 'X' in load_dir else load_ny
                                    _keyii = keyii.replace('初始土反力', '水压力')
                                    elem_load = '{}, LINE, UNILOAD, {}, NO , NO, LY, , , , {}, {}, {}, {}, {}, {}, {}, {}, {}, NO, 0, 0, NO,'.format(
                                            elem_i, load_dir, position_lst[0], round(load_lst[0]*load_n*-1*elem_load_scale,3), position_lst[1], round(load_lst[1]*load_n*-1*elem_load_scale,3), position_lst[2], round(load_lst[2]*load_n*-1*elem_load_scale,3), position_lst[3], round(load_lst[3]*load_n*-1*elem_load_scale,3), _keyii
                                            ) # 荷载信息
                                    load_grupdict[_keyii].append(elem_load)
                        elif Method_for_Pressures_var == '水土合算':
                            position_lst, load_lst = add_continuous_beam_elem_loads(elem_load_lst_trans(soild_distributed_soil_reaction_lst[1], soild_distributed_soil_reaction_lst[2]), elem_load_lst_trans(water_distributed_soil_reaction_lst[1], water_distributed_soil_reaction_lst[2]))
                            for load_dir in load_DIR_lst:
                                load_n = load_nx if 'X' in load_dir else load_ny
                                elem_load = '{}, LINE, UNILOAD, {}, NO , NO, LY, , , , {}, {}, {}, {}, {}, {}, {}, {}, {}, NO, 0, 0, NO,'.format(
                                    elem_i, load_dir, position_lst[0], round(load_lst[0]*load_n*-1*elem_load_scale,3), position_lst[1], round(load_lst[1]*load_n*-1*elem_load_scale,3), position_lst[2], round(load_lst[2]*load_n*-1*elem_load_scale,3), position_lst[3], round(load_lst[3]*load_n*-1*elem_load_scale,3), keyii
                                    )# 荷载信息
                                load_grupdict[keyii].append(elem_load)
                except:
                    pass
            elem_i += 1
    print(f'{Pile_name}单元成功生成')
    
    Waler_elasticlink_dict = {}
    Strut_node_dict = {}
    Strut_node_elem_dict = {}

    # 确定哪些辅助加撑层在施工阶段中已定义，未定义的跳过生成
    _active_assist_layers = set()
    _active_replace_stages = set()  # (layer_i, form_j)
    if Construct_Stage_Line_Data:
        for _lst in Construct_Stage_Line_Data:
            if _lst[0] == '辅助加撑':
                _active_assist_layers.add(int(_lst[3]))
            elif _lst[0] == '辅助换撑':
                _parts = str(_lst[3]).split('-')
                if len(_parts) == 2:
                    _active_replace_stages.add((int(_parts[0]), int(_parts[1])))
    _waler_strut_z_sorted = list(Waler_Strut_Z_dict.keys())  # 与 Waler_position_line_lst 顺序一致
    _skip_waler_indices = set()
    for _zi, _z in enumerate(_waler_strut_z_sorted):
        if _zi < len(Waler_position_line_lst) and Waler_Strut_Z_dict[_z].get('if_add', False):
            _layer_idx = Waler_Strut_Z_dict[_z]['index']
            if _layer_idx not in _active_assist_layers:
                _skip_waler_indices.add(_zi)
    print('辅助加撑跳过层索引(未在施工阶段定义):', _skip_waler_indices)

    if Waler_position_line_lst != []:
        # 生成围檩节点
        for z in range(len(Waler_position_line_lst)): # 第z层
            if z in _skip_waler_indices:
                continue  # 跳过未在施工阶段定义的辅助加撑层
            nodedict[f'第{z+1}层围檩'] = []
            grupdict_node[f'第{z+1}层围檩'] = []
            Pile_i = 1
            first_nodei = node_i
            for i in range(len(Waler_position_line_lst[z])-1): # 第i条边
                node_idict = Pile_elasticlink_dict[f'第{i+1}条边'] # 第i条边所有的钢板桩节点
                if i == 0:
                    project_pt_ilst = [] # 第i条边的投影点表
                else:
                    project_pt_ilst = [tuple(nodedict[f'第{z+1}层围檩'][-1][1:])] # 上一条边的终点
                if f'第{i+1}条边' not in Waler_elasticlink_dict.keys():
                    Waler_elasticlink_dict[f'第{i+1}条边'] = {}
                for j, value in enumerate(node_idict.values()): # 第i条边下的第j排钢板桩
                    pt1 = Waler_position_line_lst[z][i]
                    pt2 = Waler_position_line_lst[z][i+1]
                    if f'围檩与第{Pile_i}排{Pile_name}弹连点' not in Waler_elasticlink_dict[f'第{i+1}条边'].keys():
                        Waler_elasticlink_dict[f'第{i+1}条边'][f'围檩与第{Pile_i}排{Pile_name}弹连点'] = []
                    pt0 = value[z]
                    project_pt = point_to_line_projection(pt0[1:], pt1, pt2) # 第j排钢板桩的平面节点在第i条边上的投影点
                    if project_pt_ilst == []:
                        project_pt_ilst.append(project_pt)
                        first_node = project_pt
                        node = [node_i, round(project_pt[0], 3), round(project_pt[1], 3), round(project_pt[2], 3)]
                        nodedict[f'第{z+1}层围檩'].append(node)
                        grupdict_node[f'第{z+1}层围檩'].append(node_i)
                        node_i += 1
                    else:
                        if project_pt != project_pt_ilst[-1]:
                            if i == len(Waler_position_line_lst[z])-2 and project_pt == first_node: # 最后一个节点号等于第一个节点号
                                node = [first_nodei, round(project_pt[0], 3), round(project_pt[1], 3), round(project_pt[2], 3)]
                            else:
                                project_pt_ilst.append(project_pt)
                                node = [node_i, round(project_pt[0], 3), round(project_pt[1], 3), round(project_pt[2], 3)]
                                nodedict[f'第{z+1}层围檩'].append(node)
                                grupdict_node[f'第{z+1}层围檩'].append(node_i)
                                node_i += 1
                        else:
                            node = [node_i-1, round(project_pt[0], 3), round(project_pt[1], 3), round(project_pt[2], 3)]
                    Waler_elasticlink_dict[f'第{i+1}条边'][f'围檩与第{Pile_i}排{Pile_name}弹连点'].append(node)
                    Pile_i += 1
        # print(Waler_elasticlink_dict)
        print('围檩节点成功生成')
        # 生成内支撑节点
        for z in range(len(Waler_position_line_lst)): # 第z层
            if z in _skip_waler_indices:
                continue  # 跳过未在施工阶段定义的辅助加撑层
            nodedict[f'第{z+1}层内支撑'] = []
            if f'第{z+1}层对撑' not in Strut_node_dict.keys():
                Strut_node_dict[f'第{z+1}层对撑'] = []
            Strut_position_line_lst_duicheng_i = Strut_position_line_lst_duicheng[z]
            if Strut_position_line_lst_duicheng_i != []:
                for ptlst1 in Strut_position_line_lst_duicheng_i:
                    pt1, pt2 = ptlst1
                    ptlst = [(round(pt1[0], 3), round(pt1[1], 3), round(waler_strut_z[z]*1000 - pt1[2], 3)), (round(pt2[0], 3), round(pt2[1], 3), round(waler_strut_z[z]*1000 + pt2[2], 3))]
                    Strut_node_dict[f'第{z+1}层对撑'].append(ptlst)
            if f'第{z+1}层斜撑' not in Strut_node_dict.keys():
                Strut_node_dict[f'第{z+1}层斜撑'] = []
            Strut_position_line_lst_xiecheng_i = Strut_position_line_lst_xiecheng[z]
            if Strut_position_line_lst_xiecheng_i != []:
                for ptlst2 in Strut_position_line_lst_xiecheng_i:
                    pt1, pt2 = ptlst2
                    ptlst = [(round(pt1[0], 3), round(pt1[1], 3), round(waler_strut_z[z]*1000 - pt1[2], 3)), (round(pt2[0], 3), round(pt2[1], 3), round(waler_strut_z[z]*1000 - pt2[2], 3))]
                    Strut_node_dict[f'第{z+1}层斜撑'].append(ptlst)
        for k, v in Strut_node_dict.items():
            print(k)
            print(v)
        Strut_DuiCheng_seclst = [value1['对撑截面'] for value1 in list(Walers_Strut_selected_dict.values())] + [value2['对撑截面'] for value2 in list(assist_add_dict.values())]
        Strut_XieCheng_seclst = [value['斜撑截面'] for value in list(Walers_Strut_selected_dict.values())] + [value2['斜撑截面'] for value2 in list(assist_add_dict.values())]
        Strut_DuiCheng_mat_lst = [value1['对撑材质'] for value1 in list(Walers_Strut_selected_dict.values())] + [value2['对撑材质'] for value2 in list(assist_add_dict.values())]
        Strut_XieCheng_mat_lst = [value1['斜撑材质'] for value1 in list(Walers_Strut_selected_dict.values())] + [value2['斜撑材质'] for value2 in list(assist_add_dict.values())]
        # print(Strut_DuiCheng_seclst)
        # print(Strut_XieCheng_seclst)
        for z in range(len(Waler_position_line_lst)): # 第z层
            if z in _skip_waler_indices:
                continue  # 跳过未在施工阶段定义的辅助加撑层
            Strut_node_lst1 = Strut_node_dict[f'第{z+1}层对撑'] # 第z层对撑坐标
            Strut_node_lst2 = Strut_node_dict[f'第{z+1}层斜撑'] # 第z层斜撑坐标
            Waler_z_nodelst = [sublst[1:] for sublst in nodedict[f'第{z+1}层围檩']] # 第z层围檩的所有节点
            Waler_z_nodeilst = [sublst[0] for sublst in nodedict[f'第{z+1}层围檩']] # 第z层围檩的所有节点号
            Strut_node_elem_dict[f'第{z+1}层对撑'] = []
            Strut_node_elem_dict[f'第{z+1}层斜撑'] = []
            grupdict_node[f'第{z+1}层对撑'] = []
            grupdict_node[f'第{z+1}层斜撑'] = []
            # 对撑
            DuiCheng_SEC_Name = Strut_DuiCheng_seclst[z]  
            if DuiCheng_SEC_Name != '/':
                for ptlst1 in Strut_node_lst1:
                    elem_nodei_lst = []
                    for sec_num, sec_name in sec_num_name_lst:
                        if DuiCheng_SEC_Name == sec_name:
                            Strut_secnum = sec_num
                    for pt in ptlst1:
                        try: # 查询内支撑节点是否为已有围檩节点, 是的话将围檩节点记录作为连接内支撑单元的依据
                            lst_index = Waler_z_nodelst.index(list(pt))
                            elem_nodei_lst.append([Waler_z_nodeilst[lst_index], pt])
                            # Constraint_lst.append(Waler_z_nodeilst[lst_index]) # 内支撑一般支撑点
                        except:
                            node = [node_i, round(pt[0], 3), round(pt[1], 3), round(pt[2], 3)]
                            elem_nodei_lst.append([node_i, pt])
                            nodedict[f'第{z+1}层内支撑'].append(node)
                            # Constraint_lst.append(node_i) # 内支撑一般支撑点
                            node_i += 1
                    DuiCheng_mat_no = mat_name_to_key.get(str(Strut_DuiCheng_mat_lst[z]).strip(), 2)
                    Strut_node_elem_dict[f'第{z+1}层对撑'].append([Strut_secnum, DuiCheng_mat_no, elem_nodei_lst])
            # 斜撑
            XieCheng_SEC_Name = Strut_XieCheng_seclst[z]
            if XieCheng_SEC_Name != '/':
                for ptlst2 in Strut_node_lst2:
                    elem_nodei_lst = []
                    for sec_num, sec_name in sec_num_name_lst:
                        if Strut_XieCheng_seclst[z] == sec_name:
                            Strut_secnum = sec_num
                    for pt in ptlst2:
                        try: # 查询内支撑节点是否为已有围檩节点, 是的话将围檩节点记录作为连接内支撑单元的依据
                            lst_index = Waler_z_nodelst.index(list(pt))
                            elem_nodei_lst.append([Waler_z_nodeilst[lst_index], pt])
                            # Constraint_lst.append(Waler_z_nodeilst[lst_index]) # 内支撑一般支撑点
                        except:
                            node = [node_i, round(pt[0], 3), round(pt[1], 3), round(pt[2], 3)]
                            elem_nodei_lst.append([node_i, pt])
                            nodedict[f'第{z+1}层内支撑'].append(node)
                            # Constraint_lst.append(node_i) # 内支撑一般支撑点
                            node_i += 1
                    XieCheng_mat_no = mat_name_to_key.get(str(Strut_XieCheng_mat_lst[z]).strip(), 2)
                    Strut_node_elem_dict[f'第{z+1}层斜撑'].append([Strut_secnum, XieCheng_mat_no, elem_nodei_lst])
        print('内支撑节点成功生成')
        # 生成内支撑单元
        for key, value in Strut_node_elem_dict.items():
            elemdict[key] = [] # 单元
            grupdict_elem[key] = [] # 结构组
            for sublist in value:
                nodeid1, nodeid2 = sublist[2][0][0], sublist[2][1][0]
                elem = [elem_i, 'BEAM', str(sublist[1]), sublist[0], nodeid1, nodeid2, '90', '0'] # 单元信息
                grupdict_elem[key].append(elem_i) # 结构组
                grupdict_node[key].extend([nodeid1, nodeid2])
                elemdict[key].append(elem)
                elem_i += 1
        print('内支撑单元成功生成')

        # ==================== 计算要点第2点(1)验证 ====================
        print('\n' + '='*80)
        print('【计算要点第2点(1)】基本对撑和斜撑结构组验证')
        print('='*80)
        for z in range(len(Waler_position_line_lst)):
            _layer_i = z + 1
            _dc_key = f'第{_layer_i}层对撑'
            _xc_key = f'第{_layer_i}层斜撑'
            _dc_elems = grupdict_elem.get(_dc_key, [])
            _xc_elems = grupdict_elem.get(_xc_key, [])
            _dc_nodes = grupdict_node.get(_dc_key, [])
            _xc_nodes = grupdict_node.get(_xc_key, [])
            print(f'\n第{_layer_i}层:')
            print(f'  对撑结构组 "{_dc_key}":')
            print(f'    单元数: {len(_dc_elems)}')
            print(f'    单元号: {_dc_elems}')
            print(f'    节点数: {len(_dc_nodes)}')
            print(f'    节点号: {list(set(_dc_nodes))}')
            print(f'  斜撑结构组 "{_xc_key}":')
            print(f'    单元数: {len(_xc_elems)}')
            print(f'    单元号: {_xc_elems}')
            print(f'    节点数: {len(_xc_nodes)}')
            print(f'    节点号: {list(set(_xc_nodes))}')
        print('='*80 + '\n')

        # ==================== 辅助换撑 ====================
        # 计算要点第1点：读取辅助换撑信息，结合基本支撑信息，生成所有对撑和斜撑的节点和单元
        # 提取所有加撑及辅助换撑工况，并按施工阶段顺序排列

        # 预先生成所有换撑节点、单元和结构组（根据 assist_replace_dict 布置数据）

        # 全局节点坐标查找表（用于匹配已有节点）
        _all_node_coords = {}
        for _ndlist in nodedict.values():
            for _nd in _ndlist:
                _all_node_coords[_nd[0]] = tuple(_nd[1:])

        # 跟踪各层在每个换撑阶段的布置（layer_i → [{'dc': [...], 'xc': [...]}, ...]）
        # 索引0=原始布置，索引1=换撑1后的布置，索引2=换撑2后的布置，...
        layer_arrangements = {}
        # 跟踪各换撑阶段的元素分配（用于 stage 循环中的激活/钝化）
        # (layer_i, form_j) → {'dc_add': [...], 'dc_remove': [...], 'xc_add': [...], 'xc_remove': [...]}
        replace_stage_tracking = {}
        assit_replace_layer_in_stage_lst = [v['支撑层数'] for v in stage_dict.values() if v['工况类型'] == '辅助换撑'] # 确认施工阶段中实际用到的换撑信息，不计算全部换撑信息

        for z in range(len(Waler_position_line_lst)):
            _layer_i = z + 1
            # 原始布置
            _orig_dc = [e[0] for e in elemdict.get(f'第{_layer_i}层对撑', [])]
            _orig_xc = [e[0] for e in elemdict.get(f'第{_layer_i}层斜撑', [])]
            layer_arrangements[_layer_i] = [{'dc': list(_orig_dc), 'xc': list(_orig_xc)}]
            if _layer_i not in assist_replace_dict:
                continue

            print(f'第{z+1}层围檩中心线', Waler_position_line_lst[z])
            print(f'第{z+1}层原始对撑单元号', _orig_dc)
            print(f'第{z+1}层原始斜撑单元号', _orig_xc)
            # 构建原始元素的坐标→元素ID查找表
            _orig_dc_match = {}
            for _eid in _orig_dc:
                for _e in elemdict.get(f'第{_layer_i}层对撑', []):
                    if _e[0] == _eid:
                        _pts = set()
                        for _nid in [_e[4], _e[5]]:
                            if _nid in _all_node_coords:
                                _pts.add(_all_node_coords[_nid])
                        if len(_pts) >= 2:
                            _key = frozenset((round(p[0], 1), round(p[1], 1), round(p[2], 1)) for p in _pts)
                            _orig_dc_match[_key] = _eid
                        break
            print('原始对撑单元号配对端点坐标', _orig_dc_match)
            _orig_xc_match = {}
            for _eid in _orig_xc:
                for _e in elemdict.get(f'第{_layer_i}层斜撑', []):
                    if _e[0] == _eid:
                        _pts = set()
                        for _nid in [_e[4], _e[5]]:
                            if _nid in _all_node_coords:
                                _pts.add(_all_node_coords[_nid])
                        if len(_pts) >= 2:
                            _key = frozenset((round(p[0], 1), round(p[1], 1), round(p[2], 1)) for p in _pts)
                            _orig_xc_match[_key] = _eid
                        break
            print('原始斜撑单元号配对端点坐标', _orig_xc_match)
            
            _waler_nodelst = [sublst[1:] for sublst in nodedict[f'第{_layer_i}层围檩']]
            _waler_nodeilst = [sublst[0] for sublst in nodedict[f'第{_layer_i}层围檩']]

            # 遍历该层的每个换撑阶段
            for _form_j in sorted(assist_replace_dict[_layer_i].keys()):

                if f'{_layer_i}-{_form_j}' not in assit_replace_layer_in_stage_lst:
                    continue

                _form_data = assist_replace_dict[_layer_i][_form_j]
                _recog_data = recognize_Strut_Replace_result_dict.get(_layer_i, {}).get(_form_j, {})
                _dc_pts = _recog_data.get('position_line_duicheng_lst', [])
                _xc_pts = _recog_data.get('position_line_xiecheng_lst', [])
        
                print('')
                print(f'--第{_layer_i}层第{_form_j}次换撑形态')
                print( '--对应的所有对撑节点坐标', _dc_pts)
                print( '--对应的所有斜撑节点坐标', _xc_pts)

                # 截面号
                _dc_sec_name = _form_data.get('对撑截面', '/')
                _dc_sec_num = None
                if _dc_sec_name != '/':
                    for _sn, _sname in sec_num_name_lst:
                        if _dc_sec_name == _sname:
                            _dc_sec_num = _sn
                            break
                print('--辅助换撑对撑截面', _dc_sec_name, '--辅助换撑对撑截面号', _dc_sec_num)
                _xc_sec_name = _form_data.get('斜撑截面', '/')
                _xc_sec_num = None
                if _xc_sec_name != '/':
                    for _sn, _sname in sec_num_name_lst:
                        if _xc_sec_name == _sname:
                            _xc_sec_num = _sn
                            break
                print('--辅助换撑斜撑截面', _xc_sec_name, '--辅助换撑斜撑截面号', _xc_sec_num)

                # 材质号（仅纯增时生效；混合情况严禁出现，不考虑）
                _dc_mat_name = _form_data.get('对撑材质', '/')
                _dc_mat_no = mat_name_to_key.get(str(_dc_mat_name).strip(), 2)
                _xc_mat_name = _form_data.get('斜撑材质', '/')
                _xc_mat_no = mat_name_to_key.get(str(_xc_mat_name).strip(), 2)

                # 取上一阶段的布置
                _prev_arr = layer_arrangements[_layer_i][-1]
                print('--上一阶段布置', _prev_arr)

                # ---- DC 差集 ----
                _dc_remove = []
                _dc_add_pts = []
                # if _dc_sec_num is not None:
                _prev_dc_elems = _prev_arr.get('dc', [])
                print('--上一阶段对撑单元', _prev_dc_elems)
                # 重建 prev 匹配表（遍历所有可能的 elemdict 来源）
                _prev_dc_match = {}
                _search_keys = [f'第{_layer_i}层对撑', f'第{_layer_i}层内支撑'] + [f'第{_layer_i}层对撑换撑{_s}' for _s in range(1, _form_j)]
                for _eid in _prev_dc_elems:
                    _found = False
                    for _sk in _search_keys:
                        for _e in elemdict.get(_sk, []):
                            if _e[0] == _eid:
                                _pts = set()
                                for _nid in [_e[4], _e[5]]:
                                    if _nid in _all_node_coords:
                                        _pts.add(_all_node_coords[_nid])
                                if len(_pts) >= 2:
                                    _key = frozenset((round(p[0], 1), round(p[1], 1), round(p[2], 1)) for p in _pts)
                                    _prev_dc_match[_key] = _eid
                                _found = True
                                break
                        if _found:
                            break
                print('--上一阶段对撑单元号配对端点坐标', _prev_dc_match)
                # 与当前阶段点位比较
                _matched_keys = set()
                for _ptlst in _dc_pts:
                    _rounded = frozenset((round(p[0], 1), round(p[1], 1), round(p[2], 1)) for p in _ptlst)
                    if _rounded in _prev_dc_match:
                        _matched_keys.add(_rounded) # 单元在上一阶段存在-->当前阶段和上一阶段都有的单元
                    else:
                        _dc_add_pts.append(_ptlst) # 单元不在上一阶段-->需要在当前阶段新增的单元
                # 上一阶段有，当前阶段无 → 移除（_dc_pts为空时，所有prev元素都被移除）
                for _pk, _eid in _prev_dc_match.items():
                    if _pk not in _matched_keys:
                        _dc_remove.append(_eid) # 单元在上一阶段存在, 当前阶段没有-->移除该单元, 上一阶段的所有单元中, 除了 当前阶段和上一阶段都有的单元 都去除
                print('--相比于上一阶段，当前阶段对撑也有的端点坐标', _matched_keys)
                print('--相比于上一阶段，当前阶段对撑需要移除的单元号', _dc_remove)
                print('--相比于上一阶段，当前阶段对撑需要新增的点位', _dc_add_pts)

                # ---- XC 差集 ----
                _xc_remove = []
                _xc_add_pts = []
                # if _xc_sec_num is not None:
                _prev_xc_elems = _prev_arr.get('xc', [])
                print('--上一阶段斜撑单元', _prev_xc_elems)
                # 重建 prev 匹配表（遍历所有可能的 elemdict 来源）
                _prev_xc_match = {}
                _search_keys_xc = [f'第{_layer_i}层斜撑', f'第{_layer_i}层内支撑'] + [f'第{_layer_i}层斜撑换撑{_s}' for _s in range(1, _form_j)]
                for _eid in _prev_xc_elems:
                    _found = False
                    for _sk in _search_keys_xc:
                        for _e in elemdict.get(_sk, []):
                            if _e[0] == _eid:
                                _pts = set()
                                for _nid in [_e[4], _e[5]]:
                                    if _nid in _all_node_coords:
                                        _pts.add(_all_node_coords[_nid])
                                if len(_pts) >= 2:
                                    _key = frozenset((round(p[0], 1), round(p[1], 1), round(p[2], 1)) for p in _pts)
                                    _prev_xc_match[_key] = _eid
                                _found = True
                                break
                        if _found:
                            break
                print('--上一阶段斜撑单元号配对端点坐标', _prev_xc_match)
                # 与当前阶段点位比较
                _matched_keys_xc = set()
                for _ptlst in _xc_pts:
                    _rounded = frozenset((round(p[0], 1), round(p[1], 1), round(p[2], 1)) for p in _ptlst)
                    if _rounded in _prev_xc_match:
                        _matched_keys_xc.add(_rounded)
                    else:
                        _xc_add_pts.append(_ptlst)
                # 上一阶段有，当前阶段无 → 移除
                for _pk, _eid in _prev_xc_match.items():
                    if _pk not in _matched_keys_xc:
                        _xc_remove.append(_eid)
                print('--相比于上一阶段，当前阶段斜撑也有的端点坐标', _matched_keys_xc)
                print('--相比于上一阶段，当前阶段斜撑需要移除的单元号', _xc_remove)
                print('--相比于上一阶段，当前阶段斜撑需要新增的点位', _xc_add_pts)

                # 类型消失标志（截面为/且之前有元素）→ 不创建替换组，直接钝化原始组
                _dc_disappear = (_dc_sec_num is None and bool(_prev_arr.get('dc', [])))
                _xc_disappear = (_xc_sec_num is None and bool(_prev_arr.get('xc', [])))
                # 当有新增或去除单元, 并且不是消失, 支撑变化=True
                # _has_dc_delta = bool(_dc_remove or _dc_add_pts) and not _dc_disappear
                # _has_xc_delta = bool(_xc_remove or _xc_add_pts) and not _xc_disappear
                _has_dc_delta = bool(_dc_remove or _dc_add_pts)
                _has_xc_delta = bool(_xc_remove or _xc_add_pts)
                print('--对撑消失', _dc_disappear)
                print('--斜撑消失', _xc_disappear)
                print('--对撑变化', _has_dc_delta)
                print('--斜撑变化', _has_xc_delta)

                # 记录跟踪数据
                replace_stage_tracking[(_layer_i, _form_j)] = {
                    # 'dc_add': [], 'dc_remove': list(_dc_remove) if not _dc_disappear else [],
                    # 'xc_add': [], 'xc_remove': list(_xc_remove) if not _xc_disappear else [],
                    'dc_add': [], 'dc_remove': list(_dc_remove),
                    'xc_add': [], 'xc_remove': list(_xc_remove),
                    'grp_dc': f'第{_layer_i}层对撑换撑{_form_j}',
                    'grp_xc': f'第{_layer_i}层斜撑换撑{_form_j}',
                    'dc_deactivate_original': _dc_disappear,
                    'xc_deactivate_original': _xc_disappear,
                }

                if not _has_dc_delta and not _has_xc_delta and not _dc_disappear and not _xc_disappear:
                    # Case3: 布置不变 → 无换撑组，布置不变
                    layer_arrangements[_layer_i].append(dict(_prev_arr))
                    continue

                # ---- 创建换撑结构组（仅当有差集时创建，类型消失不创建）----
                # _has_dc_add = bool(_dc_add_pts) and not _dc_disappear
                # _has_dc_remove = bool(_dc_remove) and not _dc_disappear
                # _has_xc_add = bool(_xc_add_pts) and not _xc_disappear
                # _has_xc_remove = bool(_xc_remove) and not _xc_disappear
                _has_dc_add = bool(_dc_add_pts)
                _has_dc_remove = bool(_dc_remove)
                _has_xc_add = bool(_xc_add_pts)
                _has_xc_remove = bool(_xc_remove)
                print('--对撑增加', _has_dc_add)
                print('--对撑减少', _has_dc_remove)
                print('--斜撑增加', _has_xc_add)
                print('--斜撑减少', _has_xc_remove)

                # DC 组（仅当有差集时创建）
                # 情况1(增加): 替换组=新增元素 → 激活
                # 情况2(减少): 替换组=被移除元素 → 钝化
                _grp_dc = f'第{_layer_i}层对撑换撑{_form_j}'
                # print(elemdict.get(f'第{_layer_i}层对撑', []))
                # print(elemdict.get(f'第{_layer_i}层斜撑', []))

                if _has_dc_remove and not _has_dc_add:
                    # Case2(减少): 替换组=被移除元素
                    grupdict_elem[_grp_dc] = []
                    grupdict_node[_grp_dc] = []
                    for _eid in _dc_remove:
                        grupdict_elem[_grp_dc].append(_eid)
                        for _e in elemdict.get(f'第{_layer_i}层对撑', []):
                            if _e[0] == _eid:
                                grupdict_node[_grp_dc].extend([_e[4], _e[5]])
                                break
                elif _has_dc_add:
                    # Case1(增加)
                    grupdict_elem[_grp_dc] = []
                    grupdict_node[_grp_dc] = []
                    for _ptlst in _dc_add_pts:
                        _elem_nodei_lst = []
                        for _pt in _ptlst:
                            try:
                                _lst_index = _waler_nodelst.index(list(_pt))
                                _elem_nodei_lst.append([_waler_nodeilst[_lst_index], _pt])
                            except:
                                _node = [node_i, round(_pt[0], 3), round(_pt[1], 3), round(_pt[2], 3)]
                                _elem_nodei_lst.append([node_i, _pt])
                                nodedict[f'第{_layer_i}层内支撑'].append(_node)
                                _all_node_coords[node_i] = tuple(_pt)
                                node_i += 1
                        if len(_elem_nodei_lst) >= 2:
                            _nodeid1, _nodeid2 = _elem_nodei_lst[0][0], _elem_nodei_lst[1][0]
                            _elem = [elem_i, 'BEAM', str(_dc_mat_no), _dc_sec_num, _nodeid1, _nodeid2, '90', '0']
                            grupdict_elem[_grp_dc].append(elem_i)
                            grupdict_node[_grp_dc].extend([_nodeid1, _nodeid2])
                            # elemdict.setdefault(f'第{_layer_i}层内支撑', []).append(_elem)
                            elemdict.setdefault(f'第{_layer_i}层对撑', []).append(_elem)
                            replace_stage_tracking[(_layer_i, _form_j)]['dc_add'].append(elem_i)
                            elem_i += 1

                # XC 组（仅当有差集时创建）
                _grp_xc = f'第{_layer_i}层斜撑换撑{_form_j}'
                if _has_xc_remove and not _has_xc_add:
                    # Case2(减少): 替换组=被移除元素
                    grupdict_elem[_grp_xc] = []
                    grupdict_node[_grp_xc] = []
                    for _eid in _xc_remove:
                        grupdict_elem[_grp_xc].append(_eid)
                        for _e in elemdict.get(f'第{_layer_i}层斜撑', []):
                            if _e[0] == _eid:
                                grupdict_node[_grp_xc].extend([_e[4], _e[5]])
                                break
                elif _has_xc_add:
                    # Case1(增加)
                    grupdict_elem[_grp_xc] = []
                    grupdict_node[_grp_xc] = []
                    for _ptlst in _xc_add_pts:
                        _elem_nodei_lst = []
                        for _pt in _ptlst:
                            try:
                                _lst_index = _waler_nodelst.index(list(_pt))
                                _elem_nodei_lst.append([_waler_nodeilst[_lst_index], _pt])
                            except:
                                _node = [node_i, round(_pt[0], 3), round(_pt[1], 3), round(_pt[2], 3)]
                                _elem_nodei_lst.append([node_i, _pt])
                                nodedict[f'第{_layer_i}层内支撑'].append(_node)
                                _all_node_coords[node_i] = tuple(_pt)
                                node_i += 1
                        if len(_elem_nodei_lst) >= 2:
                            _nodeid1, _nodeid2 = _elem_nodei_lst[0][0], _elem_nodei_lst[1][0]
                            _elem = [elem_i, 'BEAM', str(_xc_mat_no), _xc_sec_num, _nodeid1, _nodeid2, '90', '0']
                            grupdict_elem[_grp_xc].append(elem_i)
                            grupdict_node[_grp_xc].extend([_nodeid1, _nodeid2])
                            # elemdict.setdefault(f'第{_layer_i}层内支撑', []).append(_elem)
                            elemdict.setdefault(f'第{_layer_i}层斜撑', []).append(_elem)
                            replace_stage_tracking[(_layer_i, _form_j)]['xc_add'].append(elem_i)
                            elem_i += 1

                # 更新布置（供下一个换撑阶段使用）
                _prev_dc = list(_prev_arr.get('dc', []))
                _prev_xc = list(_prev_arr.get('xc', []))
                if _dc_disappear:
                    _new_dc = []
                elif not _has_dc_remove and _has_dc_add:
                    _new_dc = _prev_dc + list(grupdict_elem.get(_grp_dc, []))
                elif _has_dc_remove and not _has_dc_add:
                    _new_dc = [e for e in _prev_dc if e not in _dc_remove]
                elif _has_dc_remove and _has_dc_add:
                    _new_dc = list(grupdict_elem.get(_grp_dc, []))
                else:
                    _new_dc = _prev_dc
                if _xc_disappear:
                    _new_xc = []
                elif not _has_xc_remove and _has_xc_add:
                    _new_xc = _prev_xc + list(grupdict_elem.get(_grp_xc, []))
                elif _has_xc_remove and not _has_xc_add:
                    _new_xc = [e for e in _prev_xc if e not in _xc_remove]
                elif _has_xc_remove and _has_xc_add:
                    _new_xc = list(grupdict_elem.get(_grp_xc, []))
                else:
                    _new_xc = _prev_xc
                layer_arrangements[_layer_i].append({'dc': _new_dc, 'xc': _new_xc})

                print(f'--换撑组预生成: {_grp_dc}={grupdict_elem.get(_grp_dc, [])}, {_grp_xc}={grupdict_elem.get(_grp_xc, [])}')
                
        print('================================================')
        for track_k0, track_v0 in replace_stage_tracking.items():
            print(track_k0)
            print(track_v0)
        print('辅助换撑节点和单元成功生成')
        print('================================================')
        print('')

        # ==================== 补充围檩节点（内支撑端点） ====================
        # 将坐标在围檩线上的内支撑节点补充到围檩节点列表和围檩结构组中
        # 使围檩单元在内支撑连接处被正确分割
        _WALER_TOL = 0.001  # 坐标匹配容差(mm)
        waler_add_nodedict = {}
        for z in range(len(Waler_position_line_lst)):
            if z in _skip_waler_indices:
                continue
            _waler_key = f'第{z+1}层围檩'
            waler_add_nodedict[_waler_key] = []
            _waler_corners = Waler_position_line_lst[z]  # 围檩角点(含闭合点)
            # 遍历该层所有内支撑节点，判断是否在围檩线上
            for _snd in nodedict.get(f'第{z+1}层内支撑', []):
                _sn_id, _sn_x, _sn_y, _sn_z = _snd[0], _snd[1], _snd[2], _snd[3]
                print(_snd)
                # 检查该节点是否在围檩的四条边上
                for _ei in range(len(_waler_corners) - 1):
                    _ex1, _ey1 = _waler_corners[_ei][0], _waler_corners[_ei][1]
                    _ex2, _ey2 = _waler_corners[_ei + 1][0], _waler_corners[_ei + 1][1]
                    # 叉积判断共线性
                    _cross = ((_sn_x - _ex1) * (_ey2 - _ey1) -
                              (_sn_y - _ey1) * (_ex2 - _ex1))
                    _edge_len = max(abs(_ex2 - _ex1), abs(_ey2 - _ey1), 1)
                    # 包围盒判断是否在线段范围内
                    if (abs(_cross) < _WALER_TOL * _edge_len and
                        min(_ex1, _ex2) - _WALER_TOL <= _sn_x <= max(_ex1, _ex2) + _WALER_TOL and
                        min(_ey1, _ey2) - _WALER_TOL <= _sn_y <= max(_ey1, _ey2) + _WALER_TOL):
                        # 该内支撑节点在围檩线上，写入围檩节点列表和结构组
                        if _sn_id not in grupdict_node[_waler_key]:
                            waler_add_nodedict[_waler_key].append(_snd)
                            grupdict_node[_waler_key].append(_sn_id)
                            print(f'  第{z+1}层围檩 <- 内支撑节点{_sn_id} ({_sn_x}, {_sn_y}, {_sn_z})')
                        break
        print('围檩节点补充完成（内支撑端点已纳入围檩）')

        # ==================== 调用验证函数 ====================
        validation_errors = validate_replace_stage_tracking(
            Waler_position_line_lst,
            elemdict,
            grupdict_elem,
            layer_arrangements,
            replace_stage_tracking,
            assist_replace_dict,
            recognize_Strut_Replace_result_dict,
            Construct_Stage_Line_Data,
            sec_num_name_lst
        )
        if validation_errors:
            error_msg = '辅助换撑参数验证失败：\n\n' + '\n'.join(validation_errors)
            print(error_msg)
            try:
                import tkinter.messagebox as messagebox
                messagebox.showerror('参数错误', error_msg)
            except:
                pass
            return [], [], [], [], [], {}, [], [], []
        # 生成围檩单元
        Waler_seclst = [value1['围檩长边截面'] for value1 in list(Walers_Strut_selected_dict.values())] + [value2['围檩长边截面'] for value2 in list(assist_add_dict.values())] # 第z层内支撑截面
        Waler_mat_lst = [value1['围檩长边材质'] for value1 in list(Walers_Strut_selected_dict.values())] + [value2['围檩长边材质'] for value2 in list(assist_add_dict.values())]
        print(Waler_seclst)
        for z in range(len(Waler_position_line_lst)): # 第z层
            if z in _skip_waler_indices:
                continue  # 跳过未在施工阶段定义的辅助加撑层
            Waler_mat_no = mat_name_to_key.get(str(Waler_mat_lst[z]).strip(), 2)
            waler_nodelst = nodedict[f'第{z+1}层围檩'] + waler_add_nodedict[f'第{z+1}层围檩']
            grupdict_elem[f'第{z+1}层围檩'] = [] # 结构组
            # waler_strut_nodelst = [[x[0], x[1:]] for x in waler_nodelst + strut_nodelst]
            waler_strut_nodelst = [[x[0], x[1:]] for x in waler_nodelst]
            Pts_Counterclockwise_lst = []
            for ptlst in waler_strut_nodelst:
                Pts_Counterclockwise_lst.append([calculate_full_angle(O_point, OX_point, ptlst[1]), ptlst])
            Pts_Counterclockwise_lst = sorted(Pts_Counterclockwise_lst, key = lambda x : x[0])
            for sec_num, sec_name in sec_num_name_lst:
                if Waler_seclst[z] == sec_name:
                    Waler_secnum = sec_num
            for i in range(len(Pts_Counterclockwise_lst)):
                if i == len(Pts_Counterclockwise_lst)-1: # 首尾相连
                    elem = [elem_i, 'BEAM', str(Waler_mat_no), Waler_secnum, Pts_Counterclockwise_lst[-1][1][0], Pts_Counterclockwise_lst[1][1][0], '90', '0'] # 单元信息
                    grupdict_elem[f'第{z+1}层围檩'].append(elem_i) # 结构组
                    elemdict[key].append(elem)
                    elem_i += 1
                else:
                    elem = [elem_i, 'BEAM', str(Waler_mat_no), Waler_secnum, Pts_Counterclockwise_lst[i][1][0], Pts_Counterclockwise_lst[i+1][1][0], '90', '0'] # 单元信息
                    grupdict_elem[f'第{z+1}层围檩'].append(elem_i) # 结构组
                    elemdict[key].append(elem)
                    elem_i += 1
        print('围檩单元成功生成')
    else:
        print('无围檩内支撑节点单元')

    # ==================== 围檩角点、对撑节点、不在围檩上的对撑节点 信息输出 ====================
    _WALER_TOL_DEBUG = 0.001  # 坐标匹配容差(mm)
    print('\n' + '='*80)
    print('【调试信息】围檩角点坐标、对撑节点、不在围檩上的对撑节点（按结构组细分）')
    print('='*80)
    # 构建全局节点坐标查找表（用于从节点ID获取坐标）
    _debug_node_coords = {}
    for _ndlist in nodedict.values():
        for _nd in _ndlist:
            _debug_node_coords[_nd[0]] = (_nd[1], _nd[2], _nd[3])
    # 定义检查节点是否在围檩上的函数
    def _check_nodes_on_waler(_node_ids, _waler_corners, _tol):
        """检查节点是否落在围檩角点连成的多边形边上（仅用于调试输出）。

        以叉积判断共线、以包围盒判断在线段范围内。

        Args:
            _node_ids (iterable[int]): 待检查的节点号。
            _waler_corners (list[tuple]): 围檩角点坐标（闭合多边形）。
            _tol (float): 坐标匹配容差（mm）。

        Returns:
            tuple[list[tuple], list[tuple]]: (在围檩上的节点, 不在围檩上的节点)，
                每项为 (节点号, x, y, z)。
        """
        _on_waler = []
        _not_on_waler = []
        for _nid in sorted(_node_ids):
            if _nid not in _debug_node_coords:
                continue
            _nx, _ny, _nz = _debug_node_coords[_nid]
            _is_on = False
            for _ei in range(len(_waler_corners) - 1):
                _ex1, _ey1 = _waler_corners[_ei][0], _waler_corners[_ei][1]
                _ex2, _ey2 = _waler_corners[_ei + 1][0], _waler_corners[_ei + 1][1]
                # 叉积判断共线性
                _cross = ((_nx - _ex1) * (_ey2 - _ey1) -
                          (_ny - _ey1) * (_ex2 - _ex1))
                _edge_len = max(abs(_ex2 - _ex1), abs(_ey2 - _ey1), 1)
                # 包围盒判断是否在线段范围内
                if (abs(_cross) < _tol * _edge_len and
                    min(_ex1, _ex2) - _tol <= _nx <= max(_ex1, _ex2) + _tol and
                    min(_ey1, _ey2) - _tol <= _ny <= max(_ey1, _ey2) + _tol):
                    _is_on = True
                    break
            if _is_on:
                _on_waler.append((_nid, _nx, _ny, _nz))
            else:
                _not_on_waler.append((_nid, _nx, _ny, _nz))
        return _on_waler, _not_on_waler
    for z in range(len(Waler_position_line_lst)):
        _layer_i = z + 1
        _z_coord = _waler_strut_z_sorted[z] if z < len(_waler_strut_z_sorted) else 'N/A'
        _z_info = Waler_Strut_Z_dict.get(_z_coord, {})
        _is_add = _z_info.get('if_add', False)
        _layer_idx = _z_info.get('index', _layer_i)
        # 判断层类型
        if not _is_add:
            _layer_type = '普通加撑层'
        else:
            _layer_type = '辅助加撑层'
        # 检查是否有辅助换撑（只取施工阶段中实际存在的阶段）
        _has_replace = _layer_i in assist_replace_dict if assist_replace_dict else False
        _replace_forms = []
        if _has_replace:
            for _form_j in sorted(assist_replace_dict[_layer_i].keys()):
                if (_layer_i, _form_j) in _active_replace_stages:
                    _replace_forms.append(_form_j)
        _has_replace = len(_replace_forms) > 0
        # 跳过未激活的辅助加撑层
        if z in _skip_waler_indices:
            print(f'\n--- 第{_layer_i}层围檩 (z={_z_coord}) [{_layer_type}] [未在施工阶段定义，已跳过] ---')
            continue
        print(f'\n--- 第{_layer_i}层围檩 (z={_z_coord}) [{_layer_type}] ---')
        # 1. 围檩角点坐标
        _waler_corners = Waler_position_line_lst[z]
        print(f'  围檩角点坐标 (共{len(_waler_corners)}个点，含闭合点):')
        for _ci, _cpt in enumerate(_waler_corners):
            print(f'    角点{_ci+1}: ({_cpt[0]}, {_cpt[1]}, {_cpt[2]})')
        # 2. 按结构组检查对撑节点
        # 2.1 基本对撑结构组（普通加撑层/辅助加撑层）
        _base_group_name = f'第{_layer_i}层对撑'
        _base_node_ids = grupdict_node.get(_base_group_name, [])
        print(f'\n  【结构组: {_base_group_name}】')
        print(_base_node_ids)
        if _base_node_ids:
            _on_waler, _not_on_waler = _check_nodes_on_waler(_base_node_ids, _waler_corners, _WALER_TOL_DEBUG)
            Cap_Cons_dict[_base_group_name] = [_nid for _nid, _nx, _ny, _nz in _not_on_waler]
            print(f'    对撑节点 (共{len(_base_node_ids)}个):')
            for _nid, _nx, _ny, _nz in _on_waler:
                print(f'      节点{_nid}: ({_nx}, {_ny}, {_nz}) [在围檩上]')
            for _nid, _nx, _ny, _nz in _not_on_waler:
                print(f'      节点{_nid}: ({_nx}, {_ny}, {_nz}) [不在围檩上]')
            if _not_on_waler:
                print(f'    >>> 不在围檩上的节点: {len(_not_on_waler)}个')
            else:
                print(f'    >>> 所有节点均在围檩上')
        else:
            print(f'    对撑节点: 无')
        # 2.2 辅助换撑阶段的结构组
        for _form_j in _replace_forms:
            _replace_group_name = f'第{_layer_i}层对撑换撑{_form_j}'
            _replace_node_ids = grupdict_node.get(_replace_group_name, [])
            print(f'\n  【结构组: {_replace_group_name}】')
            print(_replace_node_ids)
            if _replace_node_ids:
                _on_waler, _not_on_waler = _check_nodes_on_waler(_replace_node_ids, _waler_corners, _WALER_TOL_DEBUG)
                Cap_Cons_dict[_replace_group_name] = [_nid for _nid, _nx, _ny, _nz in _not_on_waler]
                print(f'    对撑节点 (共{len(_replace_node_ids)}个):')
                for _nid, _nx, _ny, _nz in _on_waler:
                    print(f'      节点{_nid}: ({_nx}, {_ny}, {_nz}) [在围檩上]')
                for _nid, _nx, _ny, _nz in _not_on_waler:
                    print(f'      节点{_nid}: ({_nx}, {_ny}, {_nz}) [不在围檩上]')
                if _not_on_waler:
                    print(f'    >>> 不在围檩上的节点: {len(_not_on_waler)}个')
                else:
                    print(f'    >>> 所有节点均在围檩上')
            else:
                print(f'    对撑节点: 无')
    print('='*80 + '\n')

    # 添加节点和单元
    node_lst = []
    elem_lst = []
    for key,value in nodedict.items():
        for sublist in value:
            node_lst.append(', '.join(str(num) for num in sublist))
    for key,value in elemdict.items():
        for sublist in value:
            elem_lst.append(', '.join(str(item[0]) if isinstance(item, list) else str(item) for item in sublist))
    # 添加结构组（在施工阶段循环之后生成，以包含换撑结构组）
    print('结构组成功生成')
    # 添加桩底节点支承
    pile_cons_lst = [group_arithmetic_sequences(Constraint_lst['桩底约束']) + f", {cofferdam_cons}, 桩底约束"]
    print('桩底约束成功生成')
    # 添加承台固结
    cap_cons = '111111'
    cap_cons_lst = []
    for k, v in Cap_Cons_dict.items():
        if v != []:
            cap_cons_lst.extend([group_arithmetic_sequences(v) + f", {cap_cons}, {k}"])
            BNDR_GROUP_lst.append(f'{k}, 0')
    print('承台约束成功生成')
    cons_lst = pile_cons_lst + cap_cons_lst

    # 添加封底
    NSDx = 1e7
    spring_lst = []
    spring_comp_lst = []
    spring_linear_lst = []
    spring_comp_dict = {}
    for i, a in Constraint_lst['封底']:
        if a not in spring_comp_dict.keys():
            spring_comp_dict[a] = []
        spring_comp_dict[a].append(i)
    if Concrete_Plug_thinckness != 0:
        for ka, va in spring_comp_dict.items():
            spring_comp_lst.append('{}, COMP, {}, 0, 0, 0, {}, 封底, 0, 0, 0, 0, 0'.format(group_arithmetic_sequences(va), ka, NSDx))
        print(f'封底边界成功生成')
    else:
        print('无封底')
    # 添加圈梁
    spring_comp_dict = {}
    for i, z, a in Constraint_lst['圈梁']:
        if z not in spring_comp_dict.keys():
            spring_comp_dict[z] = {}
        if a not in spring_comp_dict[z].keys():
            spring_comp_dict[z][a] = []
        spring_comp_dict[z][a].append(i)
    for ka, va in spring_comp_dict.items():
        for kb, vb in va.items():
            spring_comp_lst.append('{}, COMP, {}, 0, 0, 0, {}, z={}m处圈梁, 0, 0, 0, 0, 0'.format(group_arithmetic_sequences(vb), kb, NSDx, ka))
    print('圈梁边界成功生成')
    
    # 创建弹性连接
    elink_lst = [] # 所有的弹连表
    elink_i = 1
    Pile_Waler_elink_nodei_lst = [] # 钢板桩与围檩的弹连节点号表
    Bracket_Waler_elink_nodei_lst = [] # 牛腿与围檩的弹连节点号表
    if Waler_position_line_lst != []:
        # 围檩与支护桩
        # 确定每一组弹性连接节点号
        for key1, value1 in Pile_elasticlink_dict.items():
            Pile_Waler_elink_nodei = []
            for key2, value2 in value1.items():
                Pile_nodei_lst = [x[0] for x in value2]
                Waler_nodei_lst = [x[0] for x in Waler_elasticlink_dict[key1]['围檩与' + key2 + '弹连点']]
                i_lst = [i+1 for i in range(len(value2))]
                Pile_Waler_elink_nodei.append((Pile_nodei_lst, Waler_nodei_lst, i_lst))
            Pile_Waler_elink_nodei_lst.append(Pile_Waler_elink_nodei)
        # [([3, 4], [1857, 1969], [1, 2]), ([19, 20], [1858, 1970], [1, 2])] # 3, 4是同一排钢板桩节点, 1857, 1969是对应的围檩弹连点, 1, 2是第1, 2层围檩(自上而下)
        for x in Pile_Waler_elink_nodei_lst:
            print(x)
        # 确定弹性连接的形式
        elink_name = '围檩与支护桩连接'
        elink_form = cofferdam_elasticlink_form_dict[elink_name]
        if elink_form != '共节点':
            for elink_key, elink_value in sheet_elasticlink_dict[elink_form].items():
                value_lst = list(elink_value.values()) 
                # 有任何非0或者非/的数
                if any(isinstance(x, str) and x not in ('0', '/') for x in value_lst):
                    if elink_key == 'GEN':
                        sdx = elink_value['sdx']
                        sdy = elink_value['sdy']
                        sdz = elink_value['sdz']
                        srx = elink_value['srx']
                        sry = elink_value['sry']
                        srz = elink_value['srz']
                        sdx = sdx if sdx != '/' else '0'
                        sdy = sdy if sdy != '/' else '0'
                        sdz = sdz if sdz != '/' else '0'
                        srx = srx if srx != '/' else '0'
                        sry = sry if sry != '/' else '0'
                        srz = srz if srz != '/' else '0'
                        for elink_sublist in Pile_Waler_elink_nodei_lst:
                            for sublist in elink_sublist:
                                for x, y, i in zip(sublist[0], sublist[1], sublist[2]):
                                    mctstring = "{}, {}, {}, {}, 0, NO, NO, NO, NO, NO, NO, {}, {}, {}, {}, {}, {}, NO, 0.5, 0.5, 第{}层{}弹性连接".format(elink_i, x, y, elink_key, sdx, sdy, sdz, srx, sry, srz, i, elink_name)
                                    elink_lst.append(mctstring) # 输入单位是KN和m
                                    elink_i = elink_i + 1
                    elif elink_key == 'COMP':
                        nsdx = elink_value['nsdx']
                        nsdx = nsdx if nsdx != '/' else '0'
                        for elink_sublist in Pile_Waler_elink_nodei_lst:
                            for i, sublist in enumerate(elink_sublist):
                                # 判断是否为角桩
                                if Sheet_Pile_typevar == '钢板桩':
                                    elink_i_form_verify = (i == 0 or i == len(elink_sublist)-1)
                                elif Sheet_Pile_typevar == '锁扣钢管桩':
                                    elink_i_form_verify = (i == len(elink_sublist)-1)
                                else:
                                    elink_i_form_verify = False
                                # 角桩边界处理
                                if elink_i_form_verify:
                                    for x, y, i in zip(sublist[0], sublist[1], sublist[2]):
                                        mctstring = "{}, {}, {}, {}, 0, NO, NO, NO, NO, NO, NO, {}, {}, {}, {}, {}, {}, NO, 0.5, 0.5, 第{}层{}弹性连接".format(elink_i, x, y, 'GEN', '1e7', '1e6', '1e6', '0', '0', '0', i, elink_name)
                                        elink_lst.append(mctstring) # 输入单位是KN和m
                                        elink_i = elink_i + 1
                                else:
                                    for x, y, i in zip(sublist[0], sublist[1], sublist[2]):
                                        mctstring = "{}, {}, {}, {}, 0, {}, NO, 0.5, 0.5, 第{}层{}弹性连接".format(elink_i, x, y, elink_key, nsdx, i, elink_name)
                                        elink_lst.append(mctstring) # 输入单位是KN和m
                                        elink_i = elink_i + 1
        print('围檩与钢板桩弹性连接成功生成')
        # 牛腿与围檩
        # 确定每一组弹性连接节点号
        for key1, value1 in Pile_elasticlink_dict_with_bracket.items():
            for key2, value2 in value1.items():
                Pile_nodei_lst = [x[0] for x in value2]
                Waler_nodei_lst = [x[0] for x in Waler_elasticlink_dict[key1]['围檩与' + key2 + '弹连点']]
                i_lst = [i+1 for i in range(len(value2))]
                Bracket_Waler_elink_nodei_lst.append((Pile_nodei_lst, Waler_nodei_lst, i_lst))
        # 确定弹性连接的形式
        elink_name = '牛腿与围檩连接'
        elink_form = cofferdam_elasticlink_form_dict[elink_name]
        if elink_form != '共节点':
            for elink_key, elink_value in sheet_elasticlink_dict[elink_form].items():
                value_lst = list(elink_value.values()) 
                # 有任何非0或者非/的数
                if any(isinstance(x, str) and x not in ('0', '/') for x in value_lst):
                    if elink_key == 'GEN':
                        sdx = elink_value['sdx']
                        sdy = elink_value['sdy']
                        sdz = elink_value['sdz']
                        srx = elink_value['srx']
                        sry = elink_value['sry']
                        srz = elink_value['srz']
                        sdx = sdx if sdx != '/' else '0'
                        sdy = sdy if sdy != '/' else '0'
                        sdz = sdz if sdz != '/' else '0'
                        srx = srx if srx != '/' else '0'
                        sry = sry if sry != '/' else '0'
                        srz = srz if srz != '/' else '0'
                        for sublist in Bracket_Waler_elink_nodei_lst:
                            for x, y, i in zip(sublist[0], sublist[1], sublist[2]):
                                mctstring = "{}, {}, {}, {}, 0, NO, NO, NO, NO, NO, NO, {}, {}, {}, {}, {}, {}, NO, 0.5, 0.5, 第{}层{}弹性连接".format(elink_i, x, y, elink_key, sdx, sdy, sdz, srx, sry, srz, i, elink_name)
                                elink_lst.append(mctstring) # 输入单位是KN和m
                                elink_i = elink_i + 1
                    elif elink_key == 'COMP':
                        nsdx = elink_value['nsdx']
                        nsdx = nsdx if nsdx != '/' else '0'
                        for sublist in Bracket_Waler_elink_nodei_lst:
                            for x, y, i in zip(sublist[0], sublist[1], sublist[2]):
                                mctstring = "{}, {}, {}, {}, 0, {}, NO, 0.5, 0.5, 第{}层{}弹性连接".format(elink_i, x, y, elink_key, nsdx, i, elink_name)
                                elink_lst.append(mctstring) # 输入单位是KN和m
                                elink_i = elink_i + 1
        print('牛腿与围檩弹性连接成功生成')
    else:
        print('无围檩, 无弹性连接')
    # 添加荷载组
    if none_construction_stage_check_var == True and construction_stage_check_var == False: # 如果不生成施工阶段, 则删除围檩下土反力相关荷载
        load_grupdict = {k: v for k, v in load_grupdict.items() if '围檩' not in k}
    # 添加边界组
    for key1, value1 in spring_dict.items():
        for key2, value2 in value1.items():
            spring_node_lst = group_arithmetic_sequences(value2) # 将所有节点号转换成x to y by a
            spring_linear_lst.append('{}, LINEAR, NO, NO, NO, NO, NO, NO, {}, {}, 0, 0, 0, 0, NO, 0, 0, 0, 0, 0, 0, {}, 0, 0, 0, 0, 0'.format(spring_node_lst, key2, key2, key1))
    print('节点弹性支承点成功生成')
    spring_lst = spring_comp_lst + spring_linear_lst
    # 添加工况
    stage_lst = []
    loadcomb_lst = []
    if construction_stage_check_var == True: # 存在工况
        Midas_Stage_Name_lst = Construct_Stage_Line_Data_to_Midas_Stage_Name(Construct_Stage_Line_Data)
        print('各施工阶段名称', Midas_Stage_Name_lst)

        dig_soil_pump_water_index_dict = {} # 获取所有取土/抽水的index
        for stage_name in Midas_Stage_Name_lst:
            if '取土' in stage_name or '抽水' in stage_name:
                dig_soil_pump_water_index_dict[stage_name] = Midas_Stage_Name_lst.index(stage_name)
            
        print(dig_soil_pump_water_index_dict)
        print('')

        dig_soil_pump_water_index_lst = list(dig_soil_pump_water_index_dict.values())
        print(dig_soil_pump_water_index_lst)
        print('')

        # 不同施工阶段需要激活和钝化的结构组、荷载组、边界组
        for i, stage_name in enumerate(Midas_Stage_Name_lst): 
            NAME_lst  = [stage_name, '0', 'YES', 'NO', 'NO', '5'] 
            STEP_lst  = []
            DELEM_lst = []
            DLOAD_lst = []
            DBNDR_lst = []
            if i == 0:
                AELEM_lst = [(Pile_name, '0')] # 第一个施工阶段需要激活的结构组
                ALOAD_lst = [('自重', 'FIRST'), ('基坑外主动土压力', 'FIRST')] # 第一个施工阶段需要激活的荷载组
                if Method_for_Pressures_var == '水土分算':
                    ALOAD_lst = ALOAD_lst + [('基坑外水压力', 'FIRST')]
                ABNDR_lst = [('桩底约束', 'ORIGINAL')] # 第一个施工阶段需要激活的边界组
            else:
                AELEM_lst = []
                ALOAD_lst = []
                ABNDR_lst = []
            
            # 取土工况/抽水工况
            if '取土' in stage_name or '抽水' in stage_name:
                # 判断取土/抽水工况的位置
                stage_index = dig_soil_pump_water_index_dict[stage_name] # 该取土/抽水工况在全部  工况中的index --> [0,2,4]
                stage_type = stage_dict[stage_index+1]['工况类型']
                stage_soil_level = float(stage_dict[stage_index+1]['基坑内土顶标高']) # 该工况对应基坑内土顶标高
                stage_water_level = float(stage_dict[stage_index+1]['基坑内水面标高']) # 该工况对应基坑内水面标高
                print(f'当前工况名:{stage_name}, 工况类型:{stage_type}, 基坑内土顶标高:{stage_soil_level}, 基坑内水面标高:{stage_water_level}')
                _stage_index = dig_soil_pump_water_index_lst.index(stage_index) # 该取土/抽水工况在全部  取土/抽水工况中的index --> [0,1,2]
                if _stage_index != 0:
                    last_stage_name = Midas_Stage_Name_lst[dig_soil_pump_water_index_lst[_stage_index-1]] # 上一个取土/抽水工况 的工况名
                    last_stage_index = dig_soil_pump_water_index_dict[last_stage_name] # 上一个取土/抽水工况在全部  工况中的index --> [0,2,4]
                    last_stage_type = stage_dict[last_stage_index+1]['工况类型']
                    last_stage_soil_level = float(stage_dict[last_stage_index+1]['基坑内土顶标高']) # 上一个工况对应基坑内土顶标高
                    last_stage_water_level = float(stage_dict[last_stage_index+1]['基坑内水面标高']) # 上一个工况对应基坑内水面标高
                    print(f'当前工况的上一个工况名:{last_stage_name}, 工况类型:{last_stage_type}, 基坑内土顶标高:{last_stage_soil_level}, 基坑内水面标高:{last_stage_water_level}')
                print('')
                # 结构组
                AELEM_lst = AELEM_lst + [] # 取土/抽水工况结构组无激活
                DELEM_lst = DELEM_lst + [] # 取土/抽水工况结构组无钝化
                # 边界组
                if '取土' in stage_name:
                    ABNDR_lst = ABNDR_lst + [(f'取土至{stage_soil_level}m处土层弹性支承点', 'DEFORMED')] # 取土工况激活的土弹簧
                    if _stage_index == 0:
                        DBNDR_lst = DBNDR_lst + [] # 第一个取土工况不钝化
                    else:
                        DBNDR_lst = DBNDR_lst + [(f'取土至{last_stage_soil_level}m处土层弹性支承点', 'DEFORMED')] # 钝化上一个取土工况的土弹簧
                elif '抽水' in stage_name:
                    if i== 0:
                        ABNDR_lst = ABNDR_lst + [(f'取土至{stage_soil_level}m处土层弹性支承点', 'DEFORMED')] # 抽水工况如果是第一个, 边界组激活第一个土顶标高对应的土弹簧
                    else:
                        ABNDR_lst = ABNDR_lst + [] # 抽水工况边界组无激活
                    DBNDR_lst = DBNDR_lst + [] # 抽水工况边界组无钝化
                # 荷载组
                if '取土' in stage_name:
                    level = stage_soil_level
                elif '抽水' in stage_name:
                    level = stage_water_level
                ALOAD_lst = ALOAD_lst + [(f'{stage_type}至{level}m处初始土反力', 'FIRST')] # 取土/抽水工况激活的初始土反力
                if Method_for_Pressures_var == '水土分算':
                    ALOAD_lst = ALOAD_lst + [(f'{stage_type}至{level}m处水压力', 'FIRST')]
                if _stage_index == 0:
                    DLOAD_lst = DLOAD_lst + [] # 第一个取土工况不钝化
                else:
                    if '取土' in last_stage_name:
                        last_level = last_stage_soil_level
                    elif '抽水' in last_stage_name:
                        last_level = last_stage_water_level
                    DLOAD_lst = DLOAD_lst + [(f'{last_stage_type}至{last_level}m处初始土反力', 'FIRST')] # 钝化上一个取土/抽水工况的初始土反力
                    if Method_for_Pressures_var == '水土分算':
                        DLOAD_lst = DLOAD_lst + [(f'{last_stage_type}至{last_level}m处水压力', 'FIRST')]

            # 加撑工况
            if ('施工' in stage_name) and ('支撑' in stage_name):
                # 该加撑工况是施工第i层支撑
                num = int(re.findall(r'\d+\.?\d*', stage_name)[0])
                # 结构组
                AELEM_lst = AELEM_lst + [(k, '0') for k, v in grupdict_elem.items() if v != [] and str(num) in k and '换撑' not in k] # 激活对应围檩和内支撑结构组
                DELEM_lst = DELEM_lst + [] # 加撑工况结构组无钝化
                # 边界组
                ABNDR_lst = ABNDR_lst + [(s.split(', ')[0], 'DEFORMED') for s in BNDR_GROUP_lst if f'第{num}层' in s and '弹性连接' in s] # 激活围檩、内支撑、牛腿对应边界组
                DBNDR_lst = DBNDR_lst + [] # 加撑工况边界组无钝化
                # 荷载组
                ALOAD_lst = ALOAD_lst + [] # 加撑工况荷载组无激活
                DLOAD_lst = DLOAD_lst + [] # 加撑工况荷载组无钝化

            # 封底
            elif '封底' in stage_name:
                AELEM_lst = AELEM_lst + []
                DELEM_lst = DELEM_lst + []
                ABNDR_lst = ABNDR_lst + [('封底', 'DEFORMED')] 
                DBNDR_lst = DBNDR_lst + []
                ALOAD_lst = ALOAD_lst + []
                DLOAD_lst = DLOAD_lst + []

            # 设置圈梁
            elif '圈梁' in stage_name:
                z = float(re.findall(r'-?\d+\.?\d*', stage_name)[0])
                AELEM_lst = AELEM_lst + []
                DELEM_lst = DELEM_lst + []
                ABNDR_lst = ABNDR_lst + [(f'z={z}m处圈梁', 'DEFORMED')]
                DBNDR_lst = DBNDR_lst + []
                ALOAD_lst = ALOAD_lst + []
                DLOAD_lst = DLOAD_lst + []

            # 辅助加撑
            elif '辅助撑' in stage_name:
                # 该加撑工况是施工第i层支撑
                num = int(re.findall(r'\d+\.?\d*', stage_name)[0])
                # 结构组
                AELEM_lst = AELEM_lst + [(k, '0') for k, v in grupdict_elem.items() if v != [] and str(num) in k] # 激活对应围檩和内支撑结构组
                DELEM_lst = DELEM_lst + [] # 加撑工况结构组无钝化
                # 边界组
                ABNDR_lst = ABNDR_lst + [(s.split(', ')[0], 'DEFORMED') for s in BNDR_GROUP_lst if f'第{num}层' in s and '弹性连接' in s] # 激活围檩、内支撑、牛腿对应边界组
                ABNDR_lst = ABNDR_lst + [(s.split(', ')[0], 'DEFORMED') for s in BNDR_GROUP_lst if s == f'第{num}层对撑, 0'] # 激活和承台连接的边界组
                DBNDR_lst = DBNDR_lst + [] # 加撑工况边界组无钝化
                # 荷载组
                ALOAD_lst = ALOAD_lst + [] # 加撑工况荷载组无激活
                DLOAD_lst = DLOAD_lst + [] # 加撑工况荷载组无钝化

            # 辅助换撑（使用预生成的换撑组，只做激活/钝化）
            elif '内支撑体系转换' in stage_name:
                _nums = re.findall(r'\d+', stage_name)
                _layer_i = int(_nums[0])
                _form_j = int(_nums[1])

                _tracking = replace_stage_tracking.get((_layer_i, _form_j))
                if _tracking:
                    _grp_dc = _tracking['grp_dc']
                    _grp_xc = _tracking['grp_xc']
                    _dc_remove = _tracking['dc_remove']
                    _dc_add = _tracking['dc_add']
                    _xc_remove = _tracking['xc_remove']
                    _xc_add = _tracking['xc_add']
                    
                    _has_remove = bool(_dc_remove or _xc_remove)
                    _has_add = bool(_dc_add or _xc_add)

                    if _has_add and not _has_remove:
                        # 情况1（增加）：激活替换组（仅含新增单元）
                        if grupdict_elem.get(_grp_dc, []):
                            AELEM_lst.append((_grp_dc, '0'))
                        if grupdict_elem.get(_grp_xc, []):
                            AELEM_lst.append((_grp_xc, '0'))
                        ABNDR_lst = ABNDR_lst + [(s.split(', ')[0], 'DEFORMED') for s in BNDR_GROUP_lst if s == f'{_grp_dc}, 0'] # 辅助换撑工况激活承台连接

                    elif _has_remove and not _has_add:
                        # 情况2（减少）：钝化替换组（仅含被移除单元）
                        if grupdict_elem.get(_grp_dc, []):
                            DELEM_lst.append((_grp_dc, '0'))
                        if grupdict_elem.get(_grp_xc, []):
                            DELEM_lst.append((_grp_xc, '0'))
                        DBNDR_lst = DBNDR_lst + [(s.split(', ')[0], 'DEFORMED') for s in BNDR_GROUP_lst if s == f'{_grp_dc}, 0'] # 辅助换撑工况激活承台连接

                ALOAD_lst = ALOAD_lst + [] # 辅助换撑工况荷载组无激活
                DLOAD_lst = DLOAD_lst + [] # 辅助换撑工况荷载组无钝化

            # 拆除
            elif '拆除' in stage_name:
                # 该拆撑工况是拆除第i层支撑
                num = int(re.findall(r'\d+\.?\d*', stage_name)[0])
                # 结构组
                AELEM_lst = AELEM_lst + [] # 拆撑工况结构组无激活
                DELEM_lst = DELEM_lst + [(k, '0') for k, v in grupdict_elem.items() if v != [] and (re.findall(r'\d+\.?\d*', k) and int(re.findall(r'\d+\.?\d*', k)[0]) == num)] # 钝化对应围檩和内支撑结构组
                # 边界组
                ABNDR_lst = ABNDR_lst + [] # 拆撑工况边界组无激活
                DBNDR_lst = DBNDR_lst + [(s.split(', ')[0], 'DEFORMED') for s in BNDR_GROUP_lst if f'第{num}层' in s and '弹性连接' in s] # 钝化围檩、内支撑、牛腿对应边界组
                DBNDR_lst = DBNDR_lst + [(s.split(', ')[0], 'DEFORMED') for s in BNDR_GROUP_lst if f'第{num}层对撑' in s] # 钝化所有和承台连接的边界组
                # 荷载组
                ALOAD_lst = ALOAD_lst + [] # 拆撑工况荷载组无激活
                DLOAD_lst = DLOAD_lst + [] # 拆撑工况荷载组无钝化

            print('施工阶段名称', stage_name)
            print('结构组激活', AELEM_lst)
            print('结构组钝化', DELEM_lst)
            print('边界组激活', ABNDR_lst)
            print('边界组钝化', DBNDR_lst)
            print('荷载组激活', ALOAD_lst)
            print('荷载组钝化', DLOAD_lst)
            print('')

            stage_lst.extend(midas_stage(NAME_lst, STEP_lst, AELEM_lst, DELEM_lst, ABNDR_lst, DBNDR_lst, ALOAD_lst, DLOAD_lst))

        print('施工阶段成功生成')
        for x in stage_lst:
            print(x)
    else:
        print('无施工阶段，生成荷载组合')
        print(Load_Combo_Line_Data)
        print(Load_Group_namelst)
        for row in Load_Combo_Line_Data:
            row[1] = row[1].replace("基坑底", f"取土至{Concrete_Bottom_Level}m处")
        BZ_loads = ['ST, {}, {}'.format(lst[1], float(lst[4])*float(lst[5])) for lst in Load_Combo_Line_Data if lst[0] != '']
        JB_loads = ['ST, {}, {}'.format(lst[1], float(lst[2])*float(lst[3])) for lst in Load_Combo_Line_Data if lst[0] != '']
        # 组合所有内容
        loadcomb_lst = ['NAME=标准组合, GEN, ACTIVE, 0, 0, , 0, 0, 0, 1', ', '.join(BZ_loads), 'NAME=基本组合, GEN, ACTIVE, 0, 0, , 0, 0, 0, 1', ', '.join(JB_loads)]

    # 结构组（在施工阶段循环之后生成，以包含换撑阶段动态创建的换撑结构组）
    grup_lst = ["{}, {}, {}, 0".format(key, group_arithmetic_sequences(grupdict_node[key]), group_arithmetic_sequences(value)) for key, value in grupdict_elem.items() if value != []]
    print(f'[DEBUG] grup_lst 包含 {len(grup_lst)} 个组')
    for _g in grup_lst:
        if '换撑' in _g:
            print(f'  [DEBUG] 换撑组: {_g}')
    print(grup_lst)

    return node_lst, elem_lst, grup_lst, cons_lst, elink_lst, load_grupdict, spring_lst, stage_lst, loadcomb_lst, BNDR_GROUP_lst


def Construct_Stage_Line_Data_to_Midas_Stage_Name(Construct_Stage_Line_Data):
    """根据施工阶段表格数据生成 MIDAS 中的施工阶段名称列表。

    映射规则：
        取土 -> "取土至 {基坑内土顶标高}m处"；抽水 -> "抽水至 {水面标高}m处"；
        加撑 -> "施工第 {支撑层数}道支撑"；拆撑 -> "拆除第 {支撑层数}道支撑"；
        封底 -> "封底"；垫层 -> 跳过；辅助加撑 -> "施工第 {支撑层数}道辅助撑"；
        辅助换撑 -> "第 {层}层内支撑体系转换{编号}"；辅助加圈梁 -> "在 {标高}m处设置圈梁"。

    Args:
        Construct_Stage_Line_Data (list[list]): 施工阶段表数据，每行形如
            [工况类型, 土顶标高, 水面标高, 支撑层数/层号-换撑号]。

    Returns:
        list[str]: 与输入一一对应的 MIDAS 施工阶段名称（'垫层' 被跳过）。

    Example:
        >>> Construct_Stage_Line_Data_to_Midas_Stage_Name([['加撑', '/', '/', '1']])
        ['施工第1道支撑']
    """
    Midas_Stage_Name_lst = []
    for lst in Construct_Stage_Line_Data:
        stage_name = lst[0]
        if stage_name == '垫层':
            pass
        elif stage_name == '封底':
            Midas_Stage_Name_lst.append('封底')
        elif stage_name == '取土':
            Midas_Stage_Name_lst.append(f'取土至{lst[1]}m处')
        elif stage_name == '抽水':
            Midas_Stage_Name_lst.append(f'抽水至{lst[2]}m处')
        elif stage_name == '加撑':
            Midas_Stage_Name_lst.append(f'施工第{lst[3]}道支撑')
        elif stage_name == '拆撑':
            Midas_Stage_Name_lst.append(f'拆除第{lst[3]}道支撑')
        elif stage_name == '辅助加撑':
            Midas_Stage_Name_lst.append(f'施工第{lst[3]}道辅助撑')
        elif stage_name == '辅助换撑':
            index_lst = lst[3].split('-') 
            Midas_Stage_Name_lst.append(f'第{index_lst[0]}层内支撑体系转换{index_lst[1]}')
        elif stage_name == '辅助加圈梁':
            Midas_Stage_Name_lst.append(f'在{lst[1]}m处设置圈梁')
    return Midas_Stage_Name_lst