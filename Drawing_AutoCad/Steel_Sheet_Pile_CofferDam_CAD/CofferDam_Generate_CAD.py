# 钢板桩围堰 CAD 绘图引擎（优化版）
# 支持多层内支撑平面图 + 锁扣钢管桩 + 跨层编号合并

# 1. 标准库
import re
import sys
import math
import time
import pywintypes
from pathlib import Path

# 添加项目根目录到 sys.path
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

# 2. 本地模块 - General
from General.AutoCAD import (
    vtpnt, pointlist_extend, insert_block_incad_with_block_name,
    Load_Font, Load_Standrad_Style, Add_Elevation_Symbol,
    make_line, make_polyline, Add_Header, Add_Annotation_Linear,
    Add_Annotation_MLeader, Add_Text, Add_Hatch,
    ensure_acad_document,
)
from General.Geometry import (
    calculate_full_angle,
    Line_Offset, line_inter, is_point_on_segment,
)

# 3. 本地模块
from Drawing_AutoCad.Steel_Sheet_Pile_CofferDam_CAD.CofferDam_Resource_CAD  import (int_or_float, SEC_Weight_per_meter)
from Drawing_AutoCad.Steel_Sheet_Pile_CofferDam_CAD.CofferDam_Calculate_CAD import Cap_Pile_PM, Waler_PM, Strut_PM, calc_skggz_center_spacing


# RPC_E_CALL_REJECTED / RPC_E_SERVERCALL_RETRYLATER（CAD 正忙被拒绝的调用）
_RETRYABLE_COM = {-2147418111, -2147417846}


def com_retry(func, *args, retries=20, interval=0.3, **kwargs):
    """CAD COM 调用重试：遇到 'CAD正忙' 类错误时等待后重试，其他错误直接抛出"""
    last = None
    for _ in range(retries):
        try:
            return func(*args, **kwargs)
        except pywintypes.com_error as e:
            last = e
            if e.hresult not in _RETRYABLE_COM:
                raise
            time.sleep(interval)
    raise last


def _default_mat(mat, default='Q235'):
    """材质兜底：截面存在但材质为空/'/'时使用默认牌号"""
    mat = str(mat).strip() if mat is not None else ''
    return mat if mat not in ('', '/') else default


def _mat_label(mat, suffix='B'):
    """材料表显示：普通钢材补'B'，钢板桩补'bz'（若已带后缀则不重复）"""
    mat = str(mat).strip() if mat is not None else ''
    if mat in ('', '/'):
        return ''
    if not mat.endswith(suffix):
        mat += suffix
    return mat


def Steel_Sheet_Pile_CofferDam_on_submit_DWG(
        acadapp,
        Cap_X, Cap_Y, X_offset, Y_offset, Cap_H,
        Sheet_Pile_typevar, CofferDam_L, Sheet_Pile_namevar,
        Pile_SEC_info_dict, SKGGZ_D, SKGGZ_t,
        SKGGZ_LK_Width, SKGGZ_GZ_Count, SKGGZ_GZ_Width,
        Waler_rows, Strut_blocks,
        Solid_Level, Water_Level,
        Cap_Bottom_Level, CofferDam_Top_Level,
        Concrete_Blinding_check_var, Concrete_Blinding_grade_var, Concrete_Blinding_thickness_var,
        Concrete_Plug_check_var, Concrete_Plug_grade_var, Concrete_Plug_thickness_var,
        Waler_section_dict, Strut_section_dict,
        recognize_Cap_result_dict, recognize_Waler_result_dict,
        recognize_Strut_result_dict, Steel_Sheet_Pile_CofferDam_Load_dict,
        Sheet_Pile_materialvar='',
        CofferDam_No='',
        ):

    print(Cap_X, Cap_Y, X_offset, Y_offset, Cap_H)
    print(Sheet_Pile_typevar, CofferDam_L, Sheet_Pile_namevar)
    print(Pile_SEC_info_dict)
    print(SKGGZ_D, SKGGZ_t, SKGGZ_LK_Width, SKGGZ_GZ_Count, SKGGZ_GZ_Width)
    print(Solid_Level, Water_Level)
    print(Cap_Bottom_Level, CofferDam_Top_Level)
    print(Concrete_Blinding_check_var, Concrete_Blinding_grade_var, Concrete_Blinding_thickness_var)
    print(Concrete_Plug_check_var, Concrete_Plug_grade_var, Concrete_Plug_thickness_var)
    print(CofferDam_No)
    print('')

    
    Steel_Sheet_Pile_CofferDam_Load_dict['Excel'] = {
        k: v for k, v in Steel_Sheet_Pile_CofferDam_Load_dict['Excel'].items()
        if v.get('土层名称') != ''}

    for k, v in Steel_Sheet_Pile_CofferDam_Load_dict['Excel'].items():
        print(k, v)
    print('')

    Strut_blocks_dict = {}
    strut_i = 0
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

    for k, v in Walers_Strut_selected_dict.items():
        print(k, v)
    print('')

    # ===== 材质信息收集（支护桩 + 围檩长边 + 对撑 + 斜撑 + 封底/垫层混凝土）=====
    # 围檩/对撑/斜撑按层号记录，key=层号，值为[截面, 材质]；无截面时记录['/', '/']
    material_info_dict = {
        '支护桩': [],
        '围檩': {},
        '对撑': {},
        '斜撑': {},
    }

    def _append_mat(target_lst, m):
        m = str(m).strip() if m is not None else ''
        if m and m != '/':
            target_lst.append(m)

    def _sec_mat_pair(sec, mat):
        sec = str(sec).strip() if sec is not None else ''
        mat = str(mat).strip() if mat is not None else ''
        if sec in ('', '/'):
            return ['/', '/']
        if mat in ('', '/'):
            return [sec, '/']
        return [sec, mat]

    _append_mat(material_info_dict['支护桩'], Sheet_Pile_materialvar)
    for i, v in enumerate(Walers_Strut_selected_dict.values(), start=1):
        # 围檩：目前只考虑长边
        material_info_dict['围檩'][i] = _sec_mat_pair(v.get('围檩长边截面'), v.get('围檩长边材质'))
        # 内支撑：对撑 + 斜撑
        material_info_dict['对撑'][i] = _sec_mat_pair(v.get('对撑截面'), v.get('对撑材质'))
        material_info_dict['斜撑'][i] = _sec_mat_pair(v.get('斜撑截面'), v.get('斜撑材质'))
    if Concrete_Blinding_check_var:
        material_info_dict['垫层'] = [str(Concrete_Blinding_grade_var)]
    if Concrete_Plug_check_var:
        material_info_dict['封底'] = [str(Concrete_Plug_grade_var)]
    print('材质信息检查')
    for k, v in material_info_dict.items():
        print(k, v)
    print('')

    SKGGZ_Gap = calc_skggz_center_spacing(SKGGZ_D, SKGGZ_LK_Width, SKGGZ_GZ_Count, SKGGZ_GZ_Width)

    # 计算围堰大小及支护桩位置
    Cap_Pile_PM(Cap_X, Cap_Y, X_offset, Y_offset, Sheet_Pile_typevar,
                Pile_SEC_info_dict, recognize_Cap_result_dict, SKGGZ_D, SKGGZ_Gap)
    # 计算围檩
    Waler_PM(CofferDam_Top_Level, Pile_SEC_info_dict, recognize_Cap_result_dict,
             recognize_Waler_result_dict, Walers_Strut_selected_dict,
             Waler_section_dict, Sheet_Pile_typevar, SKGGZ_D)
    # 计算内支撑
    Strut_PM(Strut_blocks_dict, recognize_Waler_result_dict, recognize_Strut_result_dict)

    # 生成dwg文件
    dwg_output(
        acadapp, Cap_H, Sheet_Pile_typevar, Sheet_Pile_namevar,
        Solid_Level, Water_Level, Cap_Bottom_Level, CofferDam_Top_Level, CofferDam_L,
        Pile_SEC_info_dict, SKGGZ_D, SKGGZ_t,
        Concrete_Blinding_check_var, Concrete_Blinding_grade_var, Concrete_Blinding_thickness_var,
        Concrete_Plug_check_var, Concrete_Plug_grade_var, Concrete_Plug_thickness_var,
        Waler_section_dict, Strut_section_dict, Walers_Strut_selected_dict,
        recognize_Cap_result_dict, recognize_Waler_result_dict, recognize_Strut_result_dict,
        Steel_Sheet_Pile_CofferDam_Load_dict, 
        SKGGZ_LK_Width, SKGGZ_GZ_Count, SKGGZ_GZ_Width,
        Sheet_Pile_materialvar,
        CofferDam_No,
    )
    print('全部阶段完成')


# ============================================================
# 绘图输出主函数
# ============================================================

def dwg_output(acadapp, Cap_H, Sheet_Pile_typevar, Sheet_Pile_namevar,
               Solid_Level, Water_Level, Cap_Bottom_Level,
               CofferDam_Top_Level, CofferDam_L,
               Pile_SEC_info_dict, SKGGZ_D, SKGGZ_t,
               Concrete_Blinding_check_var, Concrete_Blinding_grade_var, Concrete_Blinding_thickness_var,
               Concrete_Plug_check_var, Concrete_Plug_grade_var, Concrete_Plug_thickness_var,
               Waler_section_dict, Strut_section_dict,
               Walers_Strut_selected_dict,
               recognize_Cap_result_dict, recognize_Waler_result_dict, recognize_Strut_result_dict,
               Steel_Sheet_Pile_CofferDam_Load_dict,
               SKGGZ_LK_Width, SKGGZ_GZ_Count, SKGGZ_GZ_Width,
               Sheet_Pile_materialvar='',
               CofferDam_No='',
            ):
    """
    绘图输出主函数 — 协调所有绘图步骤
    """
    # 确保存在活动文档与模型空间（CAD 停留在开始页/无图纸时自动新建）
    acaddoc, acadmsp = ensure_acad_document(acadapp)
    # 用户选取 CAD 定位点
    CAD_O_point = acaddoc.Utility.GetPoint(vtpnt((0, 0, 0)), '请选择图框定位点')
    CAD_O_point = tuple(CAD_O_point)
    print('CAD选取点坐标:', CAD_O_point)

    O_point = recognize_Cap_result_dict['O_Point']
    OX_point = (O_point[0] + 1, O_point[1], O_point[2])

    # 收集截面信息
    Waler_sectionlst = [v['围檩长边截面'] for v in Walers_Strut_selected_dict.values()]
    Strut_DC_sectionlst = [v['对撑截面'] for v in Walers_Strut_selected_dict.values()]
    Strut_XC_sectionlst = [v['斜撑截面'] for v in Walers_Strut_selected_dict.values()]
    n_layers = len(Walers_Strut_selected_dict)
    # 材料编号映射默认值（无围檩/内支撑时为空的扁平dict）
    waler_number_map = {}
    global_strut_numbering = {}

    # 桩截面高度
    if Sheet_Pile_typevar == '钢板桩':
        Sheet_Pile_Y = Pile_SEC_info_dict['SEC'][0]  # Sheet_Pile_Y
        Pile_H = Sheet_Pile_Y if Sheet_Pile_Y > 400 else Sheet_Pile_Y + 30 # 拉森四图块高度为370mm, 计算高度为340mm, 拉森六均为420mm
    elif Sheet_Pile_typevar == '锁扣钢管桩':
        Pile_H = SKGGZ_D
    else:
        Pile_H = 0

    # 用户输入比例
    Global_Scale = float(acaddoc.Utility.GetString(1, '请输入图纸比例'))

    # GetString 交互后 AutoCAD 可能仍处于忙碌状态，稍等让其消息循环处理完成
    time.sleep(0.5)

    # 重新获取文档对象（GetString 之后 COM 代理可能失效，带重试）
    acaddoc = com_retry(lambda: acadapp.ActiveDocument)
    acadmsp = com_retry(lambda: acaddoc.ModelSpace)

    # 加载字体和标注样式
    Load_Font(acaddoc, 'Song_07', 0.0, 0.7)
    Load_Standrad_Style(acaddoc, Global_Scale, f'DimStyle_{int(Global_Scale)}', 'Song_07')

    # 材料表字典
    Material_dict = {
        '支护桩': {'名称': [], '规格': [], '材质': [], '数量': [ ], '单重': [   ], '总重': [], '备注': [    ]},
        '围檩':   {'名称': [], '规格': [], '材质': [], '数量': [ ], '单重': [   ], '总重': [], '备注': [    ]},
        '内支撑': {'名称': [], '规格': [], '材质': [], '数量': [ ], '单重': [   ], '总重': [], '备注': [    ]},
        '封底':   {'名称': [], '规格': [], '材质': [], '数量': [1], '单重': ['/'], '总重': [], '备注': ['m³']},
        '垫层':   {'名称': [], '规格': [], '材质': [], '数量': [1], '单重': ['/'], '总重': [], '备注': ['m³']},
    }

    if Sheet_Pile_typevar == '钢板桩':
        Sheet_Pile_Wdith = Pile_SEC_info_dict['SEC'][1]
        Sheet_Pile_name = '拉森Ⅳ钢板桩' if Sheet_Pile_Wdith == 400 else '拉森Ⅵ钢板桩'
        Sheet_Pile_L = int_or_float(CofferDam_L, 1, 3)
        Material_dict['支护桩']['名称'] = ['拉森钢板桩']  # 保持不变
        Material_dict['支护桩']['规格'] = [f'{Sheet_Pile_namevar}，L={Sheet_Pile_L}m']
        _pile_mat = _mat_label(_default_mat(Sheet_Pile_materialvar, 'Q295'), 'bz')
        Material_dict['支护桩']['材质'] = [_pile_mat]
        weight_per_m = 76.1 if Sheet_Pile_Wdith == 400 else 106.2
        Material_dict['支护桩']['数量'] = [0]
        Material_dict['支护桩']['单重'] = [round(weight_per_m * Sheet_Pile_L, 2)]
        Material_dict['支护桩']['总重'] = [0]
        Material_dict['支护桩']['备注'] = ['']
    elif Sheet_Pile_typevar == '锁扣钢管桩':
        # 锁扣钢管桩
        D_val = SKGGZ_D
        t_val = SKGGZ_t
        weight_per_m = round(3.14159 * (D_val - t_val) * t_val * 7.85 / 1000, 2)
        Sheet_Pile_L = int_or_float(CofferDam_L, 1, 3)
        Sheet_Pile_name = ''
        # 判断两锁扣钢管桩之间钢板桩数量
        if SKGGZ_GZ_Count > 0:
            # 有钢板桩：材料表包含锁扣钢管桩和拉森钢板桩两项
            Sheet_Pile_name = '拉森Ⅳ钢板桩' if SKGGZ_GZ_Width == 400 else '拉森Ⅵ钢板桩'
            Material_dict['支护桩']['名称'] = ['锁扣钢管桩', '拉森钢板桩']
            Material_dict['支护桩']['规格'] = [f'φ{D_val}×{t_val}，L={Sheet_Pile_L}m', f'{Sheet_Pile_name}，L={Sheet_Pile_L}m']
            _pile_mat = _mat_label(_default_mat(Sheet_Pile_materialvar, 'Q235'))
            Material_dict['支护桩']['材质'] = [_pile_mat, 'Q295bz']
            Material_dict['支护桩']['数量'] = [0, 0]  # 数量会在后面计算填充
            gz_weight_per_m = 76.1 if SKGGZ_GZ_Width == 400 else 106.2
            Material_dict['支护桩']['单重'] = [round(weight_per_m * Sheet_Pile_L, 2), round(gz_weight_per_m * Sheet_Pile_L, 2)]
            Material_dict['支护桩']['总重'] = [0, 0]
            Material_dict['支护桩']['备注'] = ['', '']
        else:
            # 无钢板桩：材料表只包含锁扣钢管桩
            Material_dict['支护桩']['名称'] = ['锁扣钢管桩']
            Material_dict['支护桩']['规格'] = [f'φ{D_val}×{t_val}，L={Sheet_Pile_L}m']
            _pile_mat = _mat_label(_default_mat(Sheet_Pile_materialvar, 'Q235'))
            Material_dict['支护桩']['材质'] = [_pile_mat]
            Material_dict['支护桩']['数量'] = [0]
            Material_dict['支护桩']['单重'] = [round(weight_per_m * Sheet_Pile_L, 2)]
            Material_dict['支护桩']['总重'] = [0]
            Material_dict['支护桩']['备注'] = ['']

    if Concrete_Blinding_check_var:
        Material_dict['垫层']['名称'].append('混凝土垫层')
        Material_dict['垫层']['材质'].append(Concrete_Blinding_grade_var)
    if Concrete_Plug_check_var:
        Material_dict['封底']['名称'].append('混凝土封底')
        Material_dict['封底']['材质'].append(Concrete_Plug_grade_var)

    CofferDam_Bottom_Level = CofferDam_Top_Level - CofferDam_L

    # ---- 标注字典初始化 ----
    PM_Annotation_Dimension_Xdict = {'边界': [], '承台': [], '内支撑': []}
    PM_Annotation_Dimension_Ydict = {'边界': [], '承台': [], '内支撑': []}
    PM_Annotation_Mleader_Potdict = {'内支撑': [], '围檩': []}
    PM_Annotation_Mleader_Pot = {}

    # 承台角点X坐标
    Cap_corner_ptlst = recognize_Cap_result_dict['Cap_corner_ptlst']

    # ---- 图框优先：用户指定点即为图框左下角 ----
    frame_x = CAD_O_point[0]
    frame_y = CAD_O_point[1]
    # 立即插入第一个图框
    insert_block_incad_with_block_name(
        acadmsp, vtpnt([frame_x, frame_y, 0]),
        '设计分公司图框（二分公司）',
        '设计分公司图框（二分公司）', [Global_Scale, Global_Scale, 1], 0)
    # 内容定位原点：使内容居中于图框（与 drawing_center 对齐）
    CAD_O_point = (frame_x + 220 * Global_Scale,
                   frame_y + 148.5 * Global_Scale, 0)
    # 图框内可用区域（左留白30*s，右留白10*s，上下留白20*s）
    usable_left = frame_x + 30 * Global_Scale
    usable_right = frame_x + 420 * Global_Scale - 10 * Global_Scale
    usable_bottom = frame_y + 20 * Global_Scale
    usable_top = frame_y + 297 * Global_Scale - 20 * Global_Scale
    usable_w = usable_right - usable_left
    usable_h = usable_top - usable_bottom
    # 实际图纸中心线在 220*s 位置
    drawing_center_x = frame_x + 220 * Global_Scale
    drawing_center_y = frame_y + 20 * Global_Scale + usable_h / 2
    # 四个区块中心点（以220*s为中心线）
    half_w = usable_w / 2
    half_h = usable_h / 2
    quadrant_centers = [
        (drawing_center_x - half_w / 2, drawing_center_y + half_h / 2),  # 左上
        (drawing_center_x + half_w / 2, drawing_center_y + half_h / 2),  # 右上
        (drawing_center_x - half_w / 2, drawing_center_y - half_h / 2),  # 左下
        (drawing_center_x + half_w / 2, drawing_center_y - half_h / 2),  # 右下
    ]
    print('区块中心点')
    print(quadrant_centers)
    # 左下区块中心（单层平面图位置）
    bl_cx, bl_cy = quadrant_centers[2]

    if n_layers <= 1:
        # ---- 单层(0或1道围檩)：在原图框左下区块居中绘制平面图 ----
        plan_O_point = (bl_cx, bl_cy, 0)
        draw_Cap_Pile_PM_incad(
            acadmsp, O_point, OX_point, plan_O_point, Pile_H,
            Concrete_Blinding_thickness_var, Concrete_Plug_thickness_var,
            recognize_Cap_result_dict, Sheet_Pile_typevar, Sheet_Pile_name,
            SKGGZ_D, SKGGZ_t,
            Concrete_Blinding_check_var, Concrete_Plug_check_var,
            PM_Annotation_Dimension_Xdict, PM_Annotation_Dimension_Ydict, Material_dict,
            SKGGZ_LK_Width, SKGGZ_GZ_Count, SKGGZ_GZ_Width,
        )
        if n_layers == 1:
            _draw_single_layer_plan(
                acadmsp, plan_O_point, 0,
                Waler_sectionlst[0], Strut_DC_sectionlst[0], Strut_XC_sectionlst[0],
                recognize_Waler_result_dict, recognize_Strut_result_dict,
                Waler_section_dict, Strut_section_dict,
                PM_Annotation_Dimension_Xdict, PM_Annotation_Dimension_Ydict,
                PM_Annotation_Mleader_Potdict,
                Walers_Strut_selected_dict,
            )
        # 计算围檩W编号映射（与多层路径格式一致：扁平dict）
        waler_number_map = {}
        waler_num = 1
        for item in PM_Annotation_Mleader_Potdict.get('围檩', []):
            sec, mat, L = item[1]
            L_key = round(L, 1)
            if (sec, mat, L_key) not in waler_number_map:
                waler_number_map[(sec, mat, L_key)] = waler_num
                waler_num += 1
        # 计算内支撑N编号映射
        global_strut_numbering = {}
        strut_num = 1
        for item in PM_Annotation_Mleader_Potdict.get('内支撑', []):
            key = item[1]  # (type, sec, mat, length)
            if key not in global_strut_numbering:
                global_strut_numbering[key] = strut_num
                strut_num += 1
        draw_Dimension_Mleader_Header_Elevation_PM_incad(
            acadapp, acadmsp, plan_O_point, Pile_H, Global_Scale,
            PM_Annotation_Mleader_Pot,
            PM_Annotation_Dimension_Xdict, PM_Annotation_Dimension_Ydict,
            PM_Annotation_Mleader_Potdict, 
            title='2---2',
            waler_number_map=waler_number_map, global_strut_numbering=global_strut_numbering,
        )
    else:
        # ---- 多层：原图框不画平面图，在右侧新图框中绘制 ----
        # 新图框左下角 = 原图框左下角 + (s*420, 0, 0)
        new_frame_x = frame_x + Global_Scale * 420
        new_frame_y = frame_y
        insert_block_incad_with_block_name(
            acadmsp, vtpnt([new_frame_x, new_frame_y, 0]),
            '设计分公司图框（二分公司）',
            '设计分公司图框（二分公司）', [Global_Scale, Global_Scale, 1], 0)

        # 新图框内可用区域（与原图框相同的留白比例）
        nf_usable_left = new_frame_x + 30 * Global_Scale
        nf_usable_right = new_frame_x + 420 * Global_Scale - 10 * Global_Scale
        nf_usable_bottom = new_frame_y + 20 * Global_Scale
        nf_usable_top = new_frame_y + 297 * Global_Scale - 20 * Global_Scale
        nf_usable_w = nf_usable_right - nf_usable_left
        nf_usable_h = nf_usable_top - nf_usable_bottom
        gap_x = 10 * Global_Scale
        gap_y = 10 * Global_Scale
        cell_w = (nf_usable_w - gap_x) / 2
        cell_h = (nf_usable_h - gap_y) / 2

        # 新图框四个区块中心点：左上、右上、左下、右下
        nf_quadrant_centers = [
            (nf_usable_left + cell_w / 2,                      nf_usable_bottom + cell_h + gap_y + cell_h / 2),
            (nf_usable_left + cell_w + gap_x + cell_w / 2,     nf_usable_bottom + cell_h + gap_y + cell_h / 2),
            (nf_usable_left + cell_w / 2,                      nf_usable_bottom + cell_h / 2),
            (nf_usable_left + cell_w + gap_x + cell_w / 2,     nf_usable_bottom + cell_h / 2),
        ]

        # 预计算全局围檩和内支撑编号映射（用于平面图和材料表编号一致）
        global_strut_numbering = {}
        waler_number_map = {}  # (截面, 长度) → W编号
        _strut_num = 1
        _waler_num = 1
        for _li in range(n_layers):
            _sel = list(Walers_Strut_selected_dict.values())[_li]
            _duicheng = recognize_Strut_result_dict.get('position_line_duicheng_lst', [[]])
            _xiecheng = recognize_Strut_result_dict.get('position_line_xiecheng_lst', [[]])
            _dc = _duicheng[_li] if _li < len(_duicheng) else []
            _xc = _xiecheng[_li] if _li < len(_xiecheng) else []
            _dc_sec = Strut_DC_sectionlst[_li] if Strut_DC_sectionlst[_li] != '/' else None
            _xc_sec = Strut_XC_sectionlst[_li] if Strut_XC_sectionlst[_li] != '/' else None
            _dc_mat = _default_mat(_sel.get('对撑材质'))
            _xc_mat = _default_mat(_sel.get('斜撑材质'))
            _w_mat = _default_mat(_sel.get('围檩长边材质'))
            for _ptlst in _dc:
                _l = round(math.sqrt((_ptlst[0][0] - _ptlst[1][0]) ** 2 +
                                     (_ptlst[0][1] - _ptlst[1][1]) ** 2), 1) if _dc_sec else 0
                _key = ('对撑', _dc_sec, _dc_mat, _l)
                if _dc_sec and _key not in global_strut_numbering:
                    global_strut_numbering[_key] = _strut_num
                    _strut_num += 1
            for _ptlst in _xc:
                _l = round(math.sqrt((_ptlst[0][0] - _ptlst[1][0]) ** 2 +
                                     (_ptlst[0][1] - _ptlst[1][1]) ** 2), 1) if _xc_sec else 0
                _key = ('斜撑', _xc_sec, _xc_mat, _l)
                if _xc_sec and _key not in global_strut_numbering:
                    global_strut_numbering[_key] = _strut_num
                    _strut_num += 1
            _w_sec = _sel.get('围檩长边截面', '')
            if _w_sec:
                _waler_position = recognize_Waler_result_dict['position_line_lst'][_li]
                for _j in range(len(_waler_position) - 1):
                    _p1 = _waler_position[_j]
                    _p2 = _waler_position[_j + 1]
                    _wl = round(math.sqrt((_p1[0] - _p2[0]) ** 2 + (_p1[1] - _p2[1]) ** 2), 1)
                    _wkey = (_w_sec, _w_mat, _wl)
                    if _wkey not in waler_number_map:
                        waler_number_map[_wkey] = _waler_num
                        _waler_num += 1

        # 调试打印编号映射
        print("=" * 60)
        print("【围檩编号映射】waler_number_map:")
        print(f"  格式: (截面, 材质, 长度mm) → W编号")
        for (_ws, _wm, _wl), _wn in waler_number_map.items():
            print(f"  ({_ws}, {_wm}, {_wl}) → W{_wn}")
        print(f"  共 {len(waler_number_map)} 种围檩")
        print()
        print("【内支撑编号映射】global_strut_numbering:")
        print(f"  格式: (类型, 截面, 材质, 长度mm) → N编号")
        for (_st, _ss, _sm, _sl), _sn in global_strut_numbering.items():
            print(f"  ({_st}, {_ss}, {_sm}, {_sl}) → N{_sn}")
        print(f"  共 {len(global_strut_numbering)} 种内支撑")
        print("=" * 60)

        # 预计算每层围檩的W编号（用于立面图标注）
        # 正立面图(LM1)看短边围檩，侧立面图(LM2)看长边围檩
        waler_layer_nums_lm1 = []  # 正立面图：取每层围檩的短边分段W编号
        waler_layer_nums_lm2 = []  # 侧立面图：取每层围檩的长边分段W编号
        for _li in range(n_layers):
            _sel = list(Walers_Strut_selected_dict.values())[_li]
            _w_sec = _sel.get('围檩长边截面', '')
            _w_mat = _default_mat(_sel.get('围檩长边材质'))
            _wlst = recognize_Waler_result_dict.get('position_line_lst', [[]])
            _plst = _wlst[_li] if _li < len(_wlst) else []
            if _plst and _w_sec and len(_plst) >= 2:
                # 收集所有分段长度及对应W编号
                _seg_lengths = []
                for _j in range(len(_plst) - 1):
                    _p1, _p2 = _plst[_j], _plst[_j + 1]
                    _wl = round(math.sqrt((_p1[0] - _p2[0]) ** 2 + (_p1[1] - _p2[1]) ** 2), 1)
                    _wn = waler_number_map.get((_w_sec, _w_mat, _wl))
                    _seg_lengths.append((_wl, _wn))
                # 正立面图：取最短分段
                _short = min(_seg_lengths, key=lambda x: x[0])
                waler_layer_nums_lm1.append(_short[1] if _short[1] is not None else _li + 1)
                # 侧立面图：取最长分段
                _long = max(_seg_lengths, key=lambda x: x[0])
                waler_layer_nums_lm2.append(_long[1] if _long[1] is not None else _li + 1)
            else:
                waler_layer_nums_lm1.append(_li + 1)
                waler_layer_nums_lm2.append(_li + 1)

        # 打印每层围檩分段详情
        print(f"\n【每层围檩分段详情】")
        for _li in range(n_layers):
            _sel = list(Walers_Strut_selected_dict.values())[_li]
            _w_sec = _sel.get('围檩长边截面', '')
            _w_mat = _default_mat(_sel.get('围檩长边材质'))
            _wlst = recognize_Waler_result_dict.get('position_line_lst', [[]])
            _plst = _wlst[_li] if _li < len(_wlst) else []
            print(f"  第{_li+1}层: 截面={_w_sec}, 材质={_w_mat}, 分段数={max(0, len(_plst)-1) if _plst else 0}")
            if _plst:
                for _j in range(len(_plst) - 1):
                    _p1, _p2 = _plst[_j], _plst[_j + 1]
                    _wl = round(math.sqrt((_p1[0] - _p2[0]) ** 2 + (_p1[1] - _p2[1]) ** 2), 1)
                    _wn = waler_number_map.get((_w_sec, _w_mat, _wl), '?')
                    print(f"    分段{_j+1}: 长度={_wl} → W{_wn}")
        print(f"  正立面图waler_layer_nums_lm1(短边) = {waler_layer_nums_lm1}")
        print(f"  侧立面图waler_layer_nums_lm2(长边) = {waler_layer_nums_lm2}")
        print("=" * 60)

        for layer_idx in range(n_layers):
            cx, cy = nf_quadrant_centers[layer_idx % 4]
            layer_CAD_O_point = (cx, cy, 0)

            layer_PM_Xdict = {'边界': [], '承台': [], '内支撑': []}
            layer_PM_Ydict = {'边界': [], '承台': [], '内支撑': []}
            layer_PM_Mleader_Potdict = {'内支撑': [], '围檩': []}
            layer_PM_Mleader_Pot = {}

            draw_Cap_Pile_PM_incad(
                acadmsp, O_point, OX_point, layer_CAD_O_point, Pile_H,
                Concrete_Blinding_thickness_var, Concrete_Plug_thickness_var,
                recognize_Cap_result_dict, Sheet_Pile_typevar, Sheet_Pile_name,
                SKGGZ_D, SKGGZ_t,
                Concrete_Blinding_check_var, Concrete_Plug_check_var,
                layer_PM_Xdict, layer_PM_Ydict, Material_dict,
                SKGGZ_LK_Width, SKGGZ_GZ_Count, SKGGZ_GZ_Width,
            )
            _draw_single_layer_plan(
                acadmsp, layer_CAD_O_point, layer_idx,
                Waler_sectionlst[layer_idx], Strut_DC_sectionlst[layer_idx], Strut_XC_sectionlst[layer_idx],
                recognize_Waler_result_dict, recognize_Strut_result_dict,
                Waler_section_dict, Strut_section_dict,
                layer_PM_Xdict, layer_PM_Ydict,
                layer_PM_Mleader_Potdict,
                Walers_Strut_selected_dict,
            )
            layer_title = f'{layer_idx + 2}---{layer_idx + 2}'
            draw_Dimension_Mleader_Header_Elevation_PM_incad(
                acadapp, acadmsp, layer_CAD_O_point, Pile_H, Global_Scale,
                layer_PM_Mleader_Pot,
                layer_PM_Xdict, layer_PM_Ydict,
                layer_PM_Mleader_Potdict,
                title=layer_title,
                waler_number_map=waler_number_map, global_strut_numbering=global_strut_numbering,
            )
            for category, items in layer_PM_Mleader_Pot.items():
                if category not in PM_Annotation_Mleader_Pot:
                    PM_Annotation_Mleader_Pot[category] = {}
                for length_key, coords_lst in items.items():
                    if length_key not in PM_Annotation_Mleader_Pot[category]:
                        PM_Annotation_Mleader_Pot[category][length_key] = []
                    PM_Annotation_Mleader_Pot[category][length_key].extend(coords_lst)

        # 多层：从 position_line_lst 填充全局边界字典（立面图需要）
        position_line_lst = recognize_Cap_result_dict['position_line_lst']
        _plx = sorted(set([x[0] + CAD_O_point[0] for x in position_line_lst]))
        PM_Annotation_Dimension_Xdict['边界'].extend([_plx[0] + Pile_H / 2, _plx[1] - Pile_H / 2])
        _ply = sorted(set([x[1] + CAD_O_point[1] for x in position_line_lst]))
        PM_Annotation_Dimension_Ydict['边界'].extend([_ply[0] + Pile_H / 2, _ply[1] - Pile_H / 2])
        # 多层：填充承台坐标（立面图需要）
        Cap_corner_ptlst = recognize_Cap_result_dict['Cap_corner_ptlst']
        _first_layer_CAD_O = nf_quadrant_centers[0]
        PM_Annotation_Dimension_Xdict['承台'].extend(set([x[0] + _first_layer_CAD_O[0] for x in Cap_corner_ptlst]))
        PM_Annotation_Dimension_Ydict['承台'].extend(set([x[1] + _first_layer_CAD_O[1] for x in Cap_corner_ptlst]))

    # 用平面图实际围檩长度重建waler_number_map（确保材料表和立面图编号一致）
    waler_number_map = {}
    _waler_num = 1
    for (sec, mat, L) in PM_Annotation_Mleader_Pot.get('围檩', {}).keys():
        if (sec, mat, L) not in waler_number_map:
            waler_number_map[(sec, mat, L)] = _waler_num
            _waler_num += 1

    # 重建waler_layer_nums（正立面图短边，侧立面图长边）
    waler_layer_nums_lm1 = []
    waler_layer_nums_lm2 = []
    for _li in range(n_layers):
        _sel = list(Walers_Strut_selected_dict.values())[_li]
        _w_sec = _sel.get('围檩长边截面', '')
        _w_mat = _default_mat(_sel.get('围檩长边材质'))
        _wlst = recognize_Waler_result_dict.get('position_line_lst', [[]])
        _plst = _wlst[_li] if _li < len(_wlst) else []
        if _plst and _w_sec and len(_plst) >= 2:
            _seg_lengths = []
            for _j in range(len(_plst) - 1):
                _p1, _p2 = _plst[_j], _plst[_j + 1]
                _wl = round(math.sqrt((_p1[0] - _p2[0]) ** 2 + (_p1[1] - _p2[1]) ** 2), 1)
                _seg_lengths.append(_wl)
            _short_len = min(_seg_lengths)
            _long_len = max(_seg_lengths)
            # 用容差匹配waler_number_map中的键
            _short_wn = None
            _long_wn = None
            for (ws, wm, wl), wn in waler_number_map.items():
                if ws == _w_sec and wm == _w_mat and abs(wl - _short_len) < 50:
                    _short_wn = wn
                if ws == _w_sec and wm == _w_mat and abs(wl - _long_len) < 50:
                    _long_wn = wn
            waler_layer_nums_lm1.append(_short_wn if _short_wn else _li + 1)
            waler_layer_nums_lm2.append(_long_wn if _long_wn else _li + 1)
        else:
            waler_layer_nums_lm1.append(_li + 1)
            waler_layer_nums_lm2.append(_li + 1)

    # ---- 绘制立面图1（长边方向）----
    LM1_Annotation_Dimension_Xdict = {'边界': [], '承台': []}
    LM1_Annotation_Dimension_Ydict = {'边界': [], '承台': [], '围檩': [], '封底': [], '垫层': []}
    LM1_Annotation_Mleader_Potdict = {'围檩': [], '内支撑': [], '封底': [], '垫层': []}

    # 立面图定位点：单层用CAD_O_point，多层用区块中心（上移CofferDam_H/2使视图居中）
    CofferDam_H_for_center = abs(CofferDam_Top_Level - CofferDam_Bottom_Level) * 1000
    # 正立面图：左上区块中心，Y下移CofferDam_H/2使视图在区块内垂直居中
    lm1_ref_point = (quadrant_centers[0][0], quadrant_centers[0][1] - CofferDam_H_for_center / 2, 0)
    # 侧立面图：右上区块中心，Y下移CofferDam_H/2使视图在区块内垂直居中
    lm2_ref_point = (quadrant_centers[1][0], quadrant_centers[1][1] - CofferDam_H_for_center / 2, 0)

    CapX = recognize_Cap_result_dict['CapX']
    CapY = recognize_Cap_result_dict['CapY']
    X_offset_mm = recognize_Cap_result_dict['X_offset_mm']
    Y_offset_mm = recognize_Cap_result_dict['Y_offset_mm']
    lm1_boundary_x = [lm1_ref_point[0] - CapX / 2 - X_offset_mm, lm1_ref_point[0] + CapX / 2 + X_offset_mm]
    lm2_boundary_x = [lm2_ref_point[0] - CapY / 2 - Y_offset_mm, lm2_ref_point[0] + CapY / 2 + Y_offset_mm]
    draw_LM1_incad(
        acadmsp, CapX, Pile_H, Global_Scale,
        Walers_Strut_selected_dict, Cap_H,
        CofferDam_Top_Level, CofferDam_Bottom_Level, Cap_Bottom_Level, Solid_Level,
        Concrete_Blinding_check_var, Concrete_Blinding_thickness_var,
        Concrete_Plug_check_var, Concrete_Plug_thickness_var,
        Steel_Sheet_Pile_CofferDam_Load_dict,
        LM1_Annotation_Dimension_Xdict, LM1_Annotation_Dimension_Ydict,
        LM1_Annotation_Mleader_Potdict,
        ref_point=lm1_ref_point,
        waler_number_map=waler_number_map, 
        waler_layer_nums=waler_layer_nums_lm1,
        boundary_x=lm1_boundary_x,
    )
    draw_Dimension_Mleader_Header_Elevation_LM1_incad(
        acadapp, acadmsp, Pile_H, Global_Scale, Sheet_Pile_typevar,
        CofferDam_Bottom_Level, Solid_Level, Water_Level,
        Steel_Sheet_Pile_CofferDam_Load_dict,
        LM1_Annotation_Dimension_Xdict, LM1_Annotation_Dimension_Ydict,
        LM1_Annotation_Mleader_Potdict, CofferDam_No,
    )

    # ---- 绘制立面图2（短边方向）----
    LM2_Annotation_Dimension_Xdict = {'边界': [], '承台': []}
    LM2_Annotation_Dimension_Ydict = {'边界': [], '承台': [], '围檩': [], '封底': [], '垫层': []}
    LM2_Annotation_Mleader_Potdict = {'围檩': [], '内支撑': [], '封底': [], '垫层': []}

    draw_LM2_incad(
        acadmsp, CapY, Pile_H, 
        Walers_Strut_selected_dict, Cap_H,
        CofferDam_Top_Level, CofferDam_Bottom_Level, Cap_Bottom_Level,
        Concrete_Blinding_check_var, Concrete_Blinding_thickness_var,
        Concrete_Plug_check_var, Concrete_Plug_thickness_var,
        LM2_Annotation_Dimension_Xdict, LM2_Annotation_Dimension_Ydict,
        LM2_Annotation_Mleader_Potdict, 
        ref_point=lm2_ref_point,
        waler_number_map=waler_number_map, 
        waler_layer_nums=waler_layer_nums_lm2,
        boundary_x=lm2_boundary_x,
    )
    draw_Dimension_Mleader_Header_Elevation_LM2_incad(
        acadapp, acadmsp, Pile_H, Global_Scale,
        LM2_Annotation_Dimension_Xdict, LM2_Annotation_Dimension_Ydict,
        LM2_Annotation_Mleader_Potdict,
    )

    # ---- 材料表 ----
    if n_layers <= 1:
        # 单层：材料表在右下区块中心
        br_cx, br_cy = quadrant_centers[3]
        mat_O_point = (br_cx, br_cy + 10 * Global_Scale , 0)
    else:
        # 多层：材料表在左下区块（与附注互换）
        bl_cx, bl_cy = quadrant_centers[2]  # 左下区块中心
        mat_O_point = (bl_cx, bl_cy, 0)
    material_bottom_y = draw_Material_List_in_cad(
        acadmsp, mat_O_point, Global_Scale,
        PM_Annotation_Dimension_Xdict, PM_Annotation_Dimension_Ydict,
        PM_Annotation_Mleader_Pot, Material_dict,
        waler_number_map=waler_number_map, global_strut_numbering=global_strut_numbering,
    )

    # ---- 图框附注 ----
    if n_layers <= 1:
        # 单层：附注在平面图正下方
        # 定位点x = 平面图中心线x - 30*s
        # 定位点y = 平面图底部边界 - 5*s
        plan_bottom_y = sorted(PM_Annotation_Dimension_Ydict['边界'])[0]
        notes_x = bl_cx - 30 * Global_Scale
        notes_y = plan_bottom_y - 20 * Global_Scale
    else:
        # 多层：附注在右下区块（与材料表互换）
        # 定位点x = 区块中心x - 30*s
        # 定位点y = 区块中心y + 25*s
        br_cx, br_cy = quadrant_centers[3]  # 右下区块
        notes_x = br_cx - 30 * Global_Scale
        notes_y = br_cy + 25 * Global_Scale
    draw_frame_notes_in_cad(acadmsp, Global_Scale,
                            PM_Annotation_Dimension_Xdict, PM_Annotation_Dimension_Ydict,
                            notes_x, notes_y,
                            frame_insertpoint=(frame_x, frame_y, 0))

    print('绘图完成')


# ============================================================
# 绘制承台和支护桩平面图
# ============================================================

def draw_Cap_Pile_PM_incad(acadmsp, O_point, OX_point, CAD_O_point, Pile_H,
                            Concrete_Blinding_thickness_var, Concrete_Plug_thickness_var,
                            recognize_Cap_result_dict, Sheet_Pile_typevar, Sheet_Pile_name,
                            SKGGZ_D, SKGGZ_t,
                            Concrete_Blinding_check_var, Concrete_Plug_check_var,
                            PM_Annotation_Dimension_Xdict, PM_Annotation_Dimension_Ydict, Material_dict,
                            SKGGZ_LK_Width, SKGGZ_GZ_Count, SKGGZ_GZ_Width):
    """绘制承台轮廓、中心线和支护桩（平面图）"""
    # 绘制承台
    Cap_corner_ptlst = recognize_Cap_result_dict['Cap_corner_ptlst']
    PM_Annotation_Dimension_Xdict['承台'].extend(set([x[0] + CAD_O_point[0] for x in Cap_corner_ptlst]))
    PM_Annotation_Dimension_Ydict['承台'].extend(set([x[1] + CAD_O_point[1] for x in Cap_corner_ptlst]))
    PM_Annotation_Dimension_Xdict['承台'].append(O_point[0] + CAD_O_point[0])
    PM_Annotation_Dimension_Ydict['承台'].append(O_point[1] + CAD_O_point[1])
    flat_points = pointlist_extend(Cap_corner_ptlst, CAD_O_point)
    make_polyline(acadmsp, flat_points, '7边界线')

    # 绘制中心线
    position_line_lst = recognize_Cap_result_dict['position_line_lst']
    position_line_xlst = sorted(set([x[0] + CAD_O_point[0] for x in position_line_lst]))
    position_line_xlst = [position_line_xlst[0] + Pile_H / 2, position_line_xlst[1] - Pile_H / 2]
    position_line_ylst = sorted(set([x[1] + CAD_O_point[1] for x in position_line_lst]))
    position_line_ylst = [position_line_ylst[0] + Pile_H / 2, position_line_ylst[1] - Pile_H / 2]
    # 填充边界坐标（每层独立 dict 用于平面图标注，主 dict 用于立面图/材料表）
    PM_Annotation_Dimension_Xdict['边界'].extend(position_line_xlst)
    PM_Annotation_Dimension_Ydict['边界'].extend(position_line_ylst)
    # flat_points = pointlist_extend(position_line_lst, CAD_O_point)
    # make_polyline(acadmsp, flat_points, '3中心线')

    X = round(abs(position_line_xlst[0] - position_line_xlst[1]) / 1000, 1)
    Y = round(abs(position_line_ylst[0] - position_line_ylst[1]) / 1000, 1)
    if Concrete_Blinding_check_var:
        Material_dict['垫层']['规格'].append(f'{X}m×{Y}m×{Concrete_Blinding_thickness_var}m')
        Material_dict['垫层']['总重'].append(round(X * Y * Concrete_Blinding_thickness_var, 1))
    if Concrete_Plug_check_var:
        Material_dict['封底']['规格'].append(f'{X}m×{Y}m×{Concrete_Plug_thickness_var}m')
        Material_dict['封底']['总重'].append(round(X * Y * Concrete_Plug_thickness_var, 1))

    # 绘制支护桩
    Pile_ptlst = recognize_Cap_result_dict['Pile_ptlst']
    # print('Pile_ptlst')
    # for x in Pile_ptlst:
    #     print(x)
    total_pile_count = len([x for ptlst in Pile_ptlst for x in ptlst]) - 4  # 所有桩的总数（减去重复角点）

    if Sheet_Pile_typevar == '钢板桩':
        # 钢板桩模式：只有一项
        Material_dict['支护桩']['数量'] = [total_pile_count]
        Material_dict['支护桩']['总重'] = [round(total_pile_count * Material_dict['支护桩']['单重'][0], 1)]
    elif Sheet_Pile_typevar == '锁扣钢管桩':
        if SKGGZ_GZ_Count > 0:
            # 锁扣钢管桩 + 拉森钢板桩
            skggz_count = total_pile_count  # 钢管桩数量
            # 拉森钢板桩数量 = 每两根钢管桩之间的间隔数 × 每间隔的钢板桩数
            interval_count = sum(len(ptlst) - 1 for ptlst in Pile_ptlst)  # 所有边的间隔数之和
            gz_count = interval_count * SKGGZ_GZ_Count
            Material_dict['支护桩']['数量'] = [skggz_count, gz_count]
            Material_dict['支护桩']['总重'] = [
                round(skggz_count * Material_dict['支护桩']['单重'][0], 1),
                round(gz_count * Material_dict['支护桩']['单重'][1], 1)
            ]
        else:
            # 只有锁扣钢管桩
            Material_dict['支护桩']['数量'] = [total_pile_count]
            Material_dict['支护桩']['总重'] = [round(total_pile_count * Material_dict['支护桩']['单重'][0], 1)]

    if Sheet_Pile_typevar == '钢板桩':
        _draw_sheet_pile_blocks(acadmsp, O_point, OX_point, CAD_O_point,
                                Pile_ptlst, Sheet_Pile_name)
    elif Sheet_Pile_typevar == '锁扣钢管桩':
        _draw_interlocking_pipe_piles(acadmsp, O_point, OX_point, CAD_O_point,
                                       Pile_ptlst, SKGGZ_D, SKGGZ_t,
                                       SKGGZ_LK_Width, SKGGZ_GZ_Count, SKGGZ_GZ_Width)


def _draw_sheet_pile_blocks(acadmsp, O_point, OX_point, CAD_O_point, Pile_ptlst, Sheet_Pile_name):
    # print(Sheet_Pile_name)
    """钢板桩模式：插入钢板桩图块"""
    for ptlst in Pile_ptlst:
        pt_angle = calculate_full_angle(O_point, OX_point, ptlst[0])
        RO_Angle = (pt_angle // 90) * 90 / 180 * math.pi
        # print('RO_Angle:', RO_Angle)
        for i in range(len(ptlst)):
            flat_point = [c1 + c2 for c1, c2 in zip(ptlst[i], CAD_O_point)]
            if i != 0 and i != len(ptlst) - 1:  # 非角点
                if i % 2 == 0:
                    insert_block_incad_with_block_name(
                        acadmsp, vtpnt(flat_point),
                        f'{Sheet_Pile_name}平面图块1',
                        f'Steel_Sheet_Pile_CofferDam_{Sheet_Pile_name}', [1, 1, 1], RO_Angle)
                else:
                    insert_block_incad_with_block_name(
                        acadmsp, vtpnt(flat_point),
                        f'{Sheet_Pile_name}平面图块2',
                        f'Steel_Sheet_Pile_CofferDam_{Sheet_Pile_name}', [1, 1, 1], RO_Angle)
            else:  # 角点
                if ptlst[0][0] == ptlst[-1][0]:  # 半桩
                    if i == 0:
                        insert_block_incad_with_block_name(
                            acadmsp, vtpnt(flat_point),
                            f'{Sheet_Pile_name}平面图块6',
                            f'Steel_Sheet_Pile_CofferDam_{Sheet_Pile_name}', [1, 1, 1], RO_Angle)
                    elif i == len(ptlst) - 1:
                        insert_block_incad_with_block_name(
                            acadmsp, vtpnt(flat_point),
                            f'{Sheet_Pile_name}平面图块5',
                            f'Steel_Sheet_Pile_CofferDam_{Sheet_Pile_name}', [1, 1, 1], RO_Angle)
                else:  # 全桩
                    insert_block_incad_with_block_name(
                        acadmsp, vtpnt(flat_point),
                        f'{Sheet_Pile_name}平面图块1',
                        f'Steel_Sheet_Pile_CofferDam_{Sheet_Pile_name}', [1, 1, 1], RO_Angle)


def _draw_interlocking_pipe_piles(acadmsp, O_point, OX_point, CAD_O_point,
                                    Pile_ptlst, SKGGZ_D, SKGGZ_t,
                                    LK_Width, GZ_Count, GZ_Width):
    
    """锁扣钢管桩模式：绘制双圆 + 锁扣线段 + 中间钢板桩图块"""
    D = SKGGZ_D       # 钢管桩外径
    t = SKGGZ_t       # 钢管桩壁厚
    r_out = D / 2
    r_in = D / 2 - t

    # print('GZ_Count')
    # print(type(GZ_Count), '--->', GZ_Count)

    # 根据钢板桩宽度确定图块名称
    if abs(GZ_Width - 400) < 1:
        gz_pile_name = '拉森Ⅳ钢板桩'
    else:
        gz_pile_name = '拉森Ⅵ钢板桩'

    for ptlst in Pile_ptlst:
        pt_angle = calculate_full_angle(O_point, OX_point, ptlst[0])
        RO_Angle = (pt_angle // 90) * 90 / 180 * math.pi

        # 绘制每根钢管桩（双圆）
        for pt in ptlst:
            flat_point = [c1 + c2 for c1, c2 in zip(pt, CAD_O_point)]
            center = vtpnt(flat_point)
            # 外圆（粗实线）
            acadmsp.AddCircle(center, r_out).Layer = '1粗实线'
            # 内圆（虚线）
            circle_in = acadmsp.AddCircle(center, r_in)
            circle_in.Layer = '4虚线'

        # 绘制相邻钢管桩之间的锁扣线段和钢板桩图块
        for i in range(len(ptlst) - 1):
            pt1 = [c1 + c2 for c1, c2 in zip(ptlst[i], CAD_O_point)]
            pt2 = [c1 + c2 for c1, c2 in zip(ptlst[i + 1], CAD_O_point)]

            # 方向向量（从 pt1 指向 pt2）
            dx = pt2[0] - pt1[0]
            dy = pt2[1] - pt1[1]
            dist = math.sqrt(dx ** 2 + dy ** 2)
            if dist < 1:
                continue
            ux, uy = dx / dist, dy / dist  # 单位方向向量

            # 锁扣线段：从钢管桩外壁延伸 LK_Width
            lk1_start = (pt1[0] + ux * r_out, pt1[1] + uy * r_out, 0)
            lk1_end = (pt1[0] + ux * (r_out + LK_Width), pt1[1] + uy * (r_out + LK_Width), 0)
            lk2_start = (pt2[0] - ux * r_out, pt2[1] - uy * r_out, 0)
            lk2_end = (pt2[0] - ux * (r_out + LK_Width), pt2[1] - uy * (r_out + LK_Width), 0)
            make_line(acadmsp, [lk1_start, lk1_end], '1粗实线')
            make_line(acadmsp, [lk2_start, lk2_end], '1粗实线')

            # 钢板桩图块插入
            n_piles = GZ_Count
            if n_piles <= 0:
                continue

            # 钢板桩区域起点和终点（两根锁扣内侧之间）
            region_start_x = pt1[0] + ux * (r_out + LK_Width)
            region_start_y = pt1[1] + uy * (r_out + LK_Width)
            region_end_x = pt2[0] - ux * (r_out + LK_Width)
            region_end_y = pt2[1] - uy * (r_out + LK_Width)

            if n_piles == 1:
                # 单块钢板桩：插入图块1，Y方向镜像
                mid_x = (region_start_x + region_end_x) / 2
                mid_y = (region_start_y + region_end_y) / 2
                insert_block_incad_with_block_name(
                    acadmsp, vtpnt([mid_x, mid_y, 0]),
                    f'{gz_pile_name}平面图块1',
                    f'Steel_Sheet_Pile_CofferDam_{gz_pile_name}', [1, -1, 1], RO_Angle)
            else:
                # 多块钢板桩：等间距分布，交替镜像/正常
                for j in range(n_piles):
                    frac = (2*j + 1) / (2*n_piles)
                    px = region_start_x + (region_end_x - region_start_x) * frac
                    py = region_start_y + (region_end_y - region_start_y) * frac
                    # 奇数索引(0,2,4...)镜像，偶数索引(1,3,5...)正常
                    if j % 2 == 0:
                        scale = [1, -1, 1]
                    else:
                        scale = [1, 1, 1]
                    insert_block_incad_with_block_name(
                        acadmsp, vtpnt([px, py, 0]),
                        f'{gz_pile_name}平面图块1',
                        f'Steel_Sheet_Pile_CofferDam_{gz_pile_name}', scale, RO_Angle)


# ============================================================
# 绘制单层围檩/内支撑平面图
# ============================================================

def _draw_single_layer_plan(acadmsp, CAD_O_point, layer_idx,
                             Waler_namevar, Strut_DC_namevar, Strut_XC_namevar, 
                             recognize_Waler_result_dict,
                             recognize_Strut_result_dict,
                             Waler_section_dict, Strut_section_dict,
                             PM_Annotation_Dimension_Xdict, PM_Annotation_Dimension_Ydict,
                             PM_Annotation_Mleader_Potdict,
                             Walers_Strut_selected_dict=None):
    """绘制第 layer_idx 层的围檩和内支撑平面图"""
    # 本层材质（围檩长边/对撑/斜撑）
    _waler_mat = 'Q235'
    _dc_mat = 'Q235'
    _xc_mat = 'Q235'
    if Walers_Strut_selected_dict:
        _layer_values = list(Walers_Strut_selected_dict.values())
        if layer_idx < len(_layer_values):
            _waler_mat = _default_mat(_layer_values[layer_idx].get('围檩长边材质'))
            _dc_mat = _default_mat(_layer_values[layer_idx].get('对撑材质'))
            _xc_mat = _default_mat(_layer_values[layer_idx].get('斜撑材质'))
    # 获取截面参数
    section_value = Waler_section_dict[Waler_namevar]['SEC']
    H0 = float(section_value[0])  # 围檩截面高度
    tf = float(section_value[3])  # 翼缘宽度

    position_line_lst = recognize_Waler_result_dict['position_line_lst'][layer_idx]

    # 绘制围檩五条线
    # 围檩在桩内侧，偏移方向应往承台方向（内侧），不能用对角方向
    ptlst1, ptlst2, ptlst3, ptlst4, ptlst5 = [], [], [], [], []
    for i in range(len(position_line_lst)):
        pt = position_line_lst[i]
        nx = -1 if pt[0] > 0 else (1 if pt[0] < 0 else 0)
        ny = -1 if pt[1] > 0 else (1 if pt[1] < 0 else 0)
        ptlst1.extend([(x + nx*H0/2     , y + ny*H0/2     , 0) for x, y, z in [pt]])
        ptlst2.extend([(x + nx*(H0/2-tf), y + ny*(H0/2-tf), 0) for x, y, z in [pt]])
        ptlst3.extend([(x + nx*0        , y + ny*0        , 0) for x, y, z in [pt]])
        ptlst4.extend([(x - nx*(H0/2-tf), y - ny*(H0/2-tf), 0) for x, y, z in [pt]])
        ptlst5.extend([(x - nx*H0/2     , y - ny*H0/2     , 0) for x, y, z in [pt]])

    for ptlst, layer_name in [(ptlst1, '1粗实线'), (ptlst2, '2细实线'), (ptlst3, '3中心线'),
                               (ptlst4, '2细实线'), (ptlst5, '1粗实线')]:
        flat_points = pointlist_extend(ptlst, CAD_O_point)
        make_polyline(acadmsp, flat_points, layer_name)

    # 围檩中心线和标注
    Waler_centerline_ptlst = [(x[0] + CAD_O_point[0], x[1] + CAD_O_point[1], x[2] + CAD_O_point[2]) for x in ptlst3]
    for i in range(len(Waler_centerline_ptlst) - 1):
        pt1 = Waler_centerline_ptlst[i]
        pt2 = Waler_centerline_ptlst[i + 1]
        center_pt = ((pt1[0] + pt2[0]) / 2, (pt1[1] + pt2[1]) / 2, (pt1[2] + pt2[2]) / 2)
        Waler_Length = round(math.sqrt((pt1[0] - pt2[0]) ** 2 + (pt1[1] - pt2[1]) ** 2), 1)
        PM_Annotation_Mleader_Potdict['围檩'].append([center_pt, (Waler_namevar, _waler_mat, Waler_Length)])

    # 中心线
    centerline_ptlst = []
    for i in range(len(Waler_centerline_ptlst) - 1):
        pt1 = Waler_centerline_ptlst[i]
        pt2 = Waler_centerline_ptlst[i + 1]
        center_pt = ((pt1[0] + pt2[0]) / 2, (pt1[1] + pt2[1]) / 2, (pt1[2] + pt2[2]) / 2)
        centerline_ptlst.append(center_pt)
    if len(centerline_ptlst) >= 4:
        make_line(acadmsp, [centerline_ptlst[0], centerline_ptlst[2]], '3中心线')
        make_line(acadmsp, [centerline_ptlst[1], centerline_ptlst[3]], '3中心线')

    # 内支撑（对撑 + 斜撑）
    duicheng_lst = recognize_Strut_result_dict.get('position_line_duicheng_lst', [[]])
    xiecheng_lst = recognize_Strut_result_dict.get('position_line_xiecheng_lst', [[]])
    dc_ptlst = duicheng_lst[layer_idx] if layer_idx < len(duicheng_lst) else []
    xc_ptlst = xiecheng_lst[layer_idx] if layer_idx < len(xiecheng_lst) else []
    Strut_ptlst = dc_ptlst + xc_ptlst

    # 如果对撑和斜撑都没有，直接返回
    if not Strut_ptlst:
        return
    # 标注
    PM_Annotation_Dimension_Xdict['内支撑'].extend(
        set([x[0] + CAD_O_point[0] for ptlst in Strut_ptlst for x in ptlst]))
    PM_Annotation_Dimension_Ydict['内支撑'].extend(
        set([x[1] + CAD_O_point[1] for ptlst in Strut_ptlst for x in ptlst]))
    # 内支撑与围檩的相交线(围檩最内侧边线)
    Strut_Waler_Intersect_line = [
        (x[0] + CAD_O_point[0], x[1] + CAD_O_point[1], x[2] + CAD_O_point[2]) for x in ptlst1]
    
    # =========对撑===========
    # 获取对撑截面参数
    dc_sec_name = Strut_DC_namevar if Strut_DC_namevar != '/' else None
    if dc_sec_name:
        _dc_section_value_strut = Strut_section_dict[dc_sec_name]['SEC']
        if len(_dc_section_value_strut) == 2:  # 钢管桩
            DC_Strut_H0 = float(_dc_section_value_strut[0])
            DC_Strut_tf = float(_dc_section_value_strut[1])
            DC_H_line, DC_tf_line = '1粗实线', '4虚线'
        elif len(_dc_section_value_strut) == 4:  # 单工钢
            DC_Strut_H0 = float(_dc_section_value_strut[0])
            DC_Strut_tf = float(_dc_section_value_strut[3])
            DC_H_line, DC_tf_line = '1粗实线', '2细实线'
        else:  # 双工钢
            DC_Strut_H0 = float(_dc_section_value_strut[0])
            DC_Strut_tf = float(_dc_section_value_strut[3])
            DC_H_line, DC_tf_line = '1粗实线', '2细实线'
        for i_strut, ptlst in enumerate(dc_ptlst):
            line0 = [[c1 + c2 for c1, c2 in zip(x, CAD_O_point)] for x in ptlst]
            pt1, pt2 = line0[0], line0[1]
            center_pt = ((pt1[0] + pt2[0]) / 2, (pt1[1] + pt2[1]) / 2, (pt1[2] + pt2[2]) / 2)
            dc_length = round(math.sqrt((pt1[0] - pt2[0]) ** 2 + (pt1[1] - pt2[1]) ** 2), 1)
            PM_Annotation_Mleader_Potdict['内支撑'].append([center_pt, ('对撑', dc_sec_name, _dc_mat, dc_length)])
            make_line(acadmsp, line0, '3中心线')
            line1_lst = Line_Offset(line0, DC_Strut_H0 / 2)
            line2_lst = Line_Offset(line0, (DC_Strut_H0 - DC_Strut_tf) / 2)
            for x_line, y_line in zip(line1_lst, line2_lst):
                intersections1, intersections2 = [], []
                for j in range(len(Strut_Waler_Intersect_line) - 1):
                    segment = [Strut_Waler_Intersect_line[j], Strut_Waler_Intersect_line[j + 1]]
                    int1 = line_inter(x_line, segment)
                    int2 = line_inter(y_line, segment)
                    if int1 and is_point_on_segment(int1, segment, acc=0.5):
                        intersections1.append(int1)
                    if int2 and is_point_on_segment(int2, segment, acc=0.5):
                        intersections2.append(int2)
                if intersections1:
                    make_line(acadmsp, intersections1, DC_H_line)
                if intersections2:
                    make_line(acadmsp, intersections2, DC_tf_line)

    # =========斜撑===========
    # 获取斜撑截面参数
    xc_sec_name = Strut_XC_namevar if Strut_XC_namevar != '/' else None
    if xc_sec_name:
        _xc_section_value_strut = Strut_section_dict[xc_sec_name]['SEC']
        if len(_xc_section_value_strut) == 2:  # 钢管桩
            XC_Strut_H0 = float(_xc_section_value_strut[0])
            XC_Strut_tf = float(_xc_section_value_strut[1])
            XC_H_line, XC_tf_line = '1粗实线', '4虚线'
        elif len(_xc_section_value_strut) == 4:  # 单工钢
            XC_Strut_H0 = float(_xc_section_value_strut[0])
            XC_Strut_tf = float(_xc_section_value_strut[3])
            XC_H_line, XC_tf_line = '1粗实线', '2细实线'
        else:  # 双工钢
            XC_Strut_H0 = float(_xc_section_value_strut[0])
            XC_Strut_tf = float(_xc_section_value_strut[3])
            XC_H_line, XC_tf_line = '1粗实线', '2细实线'
        for i_strut, ptlst in enumerate(xc_ptlst):
            line0 = [[c1 + c2 for c1, c2 in zip(x, CAD_O_point)] for x in ptlst]
            pt1, pt2 = line0[0], line0[1]
            center_pt = ((pt1[0] + pt2[0]) / 2, (pt1[1] + pt2[1]) / 2, (pt1[2] + pt2[2]) / 2)
            xc_length = round(math.sqrt((pt1[0] - pt2[0]) ** 2 + (pt1[1] - pt2[1]) ** 2), 1)
            PM_Annotation_Mleader_Potdict['内支撑'].append([center_pt, ('斜撑', xc_sec_name, _xc_mat, xc_length)])
            make_line(acadmsp, line0, '3中心线')
            line1_lst = Line_Offset(line0, XC_Strut_H0 / 2)
            line2_lst = Line_Offset(line0, (XC_Strut_H0 - XC_Strut_tf) / 2)
            for x_line, y_line in zip(line1_lst, line2_lst):
                intersections1, intersections2 = [], []
                for j in range(len(Strut_Waler_Intersect_line) - 1):
                    segment = [Strut_Waler_Intersect_line[j], Strut_Waler_Intersect_line[j + 1]]
                    int1 = line_inter(x_line, segment)
                    int2 = line_inter(y_line, segment)
                    if int1 and is_point_on_segment(int1, segment, acc=0.5):
                        intersections1.append(int1)
                    if int2 and is_point_on_segment(int2, segment, acc=0.5):
                        intersections2.append(int2)
                if intersections1:
                    make_line(acadmsp, intersections1, XC_H_line)
                if intersections2:
                    make_line(acadmsp, intersections2, XC_tf_line)

# ============================================================
# 平面图标注
# ============================================================

def draw_Dimension_Mleader_Header_Elevation_PM_incad(acadapp, acadmsp, CAD_O_point, Pile_H, Global_Scale,
                                                      PM_Annotation_Mleader_Pot,
                                                      PM_Annotation_Dimension_Xdict, PM_Annotation_Dimension_Ydict,
                                                      PM_Annotation_Mleader_Potdict, 
                                                      title='2---2',
                                                      waler_number_map=None, global_strut_numbering=None):
    """绘制平面图的尺寸标注、引线标注和标题"""
    PM_Annotation_Dimension_X = [[], []]
    PM_Annotation_Dimension_Y = [[], []]
    type_mapping = {'边界': [0, 1], '承台': [0], '内支撑': [1]}
    for key, value in PM_Annotation_Dimension_Xdict.items():
        for index in type_mapping.get(key, []):
            PM_Annotation_Dimension_X[index].extend(value)
    for key, value in PM_Annotation_Dimension_Ydict.items():
        for index in type_mapping.get(key, []):
            PM_Annotation_Dimension_Y[index].extend(value)

    # 分类围檩和内支撑（新格式：围檩为 (sec, mat, length)，内支撑为 (type, sec, mat, length)）
    for category, items in PM_Annotation_Mleader_Potdict.items():
        PM_Annotation_Mleader_Pot[category] = {}
        for item in items:
            coords = item[0]
            if category == '围檩':
                sec, mat, L = item[1]
                length_key = (sec, mat, round(L, 1))
            else:
                strut_type, sec, mat, L = item[1]
                length_key = (strut_type, sec, mat, round(L, 1))
            if length_key not in PM_Annotation_Mleader_Pot[category]:
                PM_Annotation_Mleader_Pot[category][length_key] = []
            PM_Annotation_Mleader_Pot[category][length_key].append(coords)

    # 尺寸标注
    for i in range(len(PM_Annotation_Dimension_X)):
        coord_Xlst = sorted(PM_Annotation_Dimension_X[i])
        if coord_Xlst:
            Y = sorted(PM_Annotation_Dimension_Y[i])[i * (-1)] + ((-1) ** (i + 1) * Pile_H)
            for ii in range(len(coord_Xlst) - 1):
                Add_Annotation_Linear(acadmsp, (coord_Xlst[ii], Y, 0), (coord_Xlst[ii + 1], Y, 0),
                                      (coord_Xlst[ii], Y + 6 * Global_Scale * (-1) ** (i + 1), 0), 0)
            if i == 0:
                Add_Annotation_Linear(acadmsp, (coord_Xlst[0], Y, 0), (coord_Xlst[-1], Y, 0),
                                      (coord_Xlst[0], Y + 12 * Global_Scale * (-1) ** (i + 1), 0), 0)
    for i in range(len(PM_Annotation_Dimension_Y)):
        coord_Ylst = sorted(PM_Annotation_Dimension_Y[i])
        if coord_Ylst:
            X = sorted(PM_Annotation_Dimension_X[i])[i * (-1)] + ((-1) ** (i + 1) * Pile_H)
            for ii in range(len(coord_Ylst) - 1):
                Add_Annotation_Linear(acadmsp, (X, coord_Ylst[ii], 0), (X, coord_Ylst[ii + 1], 0),
                                      (X + 6 * Global_Scale * (-1) ** (i + 1), coord_Ylst[ii], 0), 90)
            if i == 0:
                Add_Annotation_Linear(acadmsp, (X, coord_Ylst[0], 0), (X, coord_Ylst[-1], 0),
                                      (X + 12 * Global_Scale * (-1) ** (i + 1), coord_Ylst[0], 0), 90)

    # 引线标注（使用全局编号）
    if waler_number_map is None:
        waler_number_map = {}
    if global_strut_numbering is None:
        global_strut_numbering = {}
    for key, value in PM_Annotation_Mleader_Pot.items():
        for length_key, ptlst in value.items():
            for pt in ptlst:
                startpoint = pt
                # nx = -2.5 if (startpoint[0] > CAD_O_point[0]) else 1
                nx = -1 if (startpoint[0] > CAD_O_point[0]) else 1
                ny = -1 if (startpoint[1] < CAD_O_point[1]) else 1
                endpoint = [startpoint[0] + 3 * nx * Global_Scale,
                            startpoint[1] - 6 * ny * Global_Scale, startpoint[2]]
                if key == '围檩':
                    sec, mat, L = length_key
                    num = 1
                    for (ws, wm, wl), wn in waler_number_map.items():
                        if ws == sec and wm == mat and abs(wl - L) < 50:
                            num = wn
                            break
                    label = f'围檩W{num}'
                else:
                    num = global_strut_numbering.get(length_key, 1)
                    label = f'内支撑N{num}'
                Add_Annotation_MLeader(acadapp, acadmsp, Global_Scale, startpoint, endpoint,
                                       '5文本', 'Song_07', label)

    # 标题
    Y1, Y2 = PM_Annotation_Dimension_Ydict['边界']
    deltaY = abs(Y2 - Y1)
    Header_Point = [CAD_O_point[0], CAD_O_point[1] + deltaY / 2 + 24 * Global_Scale, CAD_O_point[2]]
    Add_Header(acadmsp, Global_Scale, Header_Point, title, 'Song_07')


# ============================================================
# 立面图1（长边方向）
# ============================================================

def draw_LM1_incad(acadmsp, Cap_X, Pile_H, Global_Scale, Walers_Strut_selected_dict, Cap_H,
                    CofferDam_Top_Level, CofferDam_Bottom_Level, Cap_Bottom_Level, Solid_Level,
                    Concrete_Blinding_check_var, Concrete_Blinding_thickness_var,
                    Concrete_Plug_check_var, Concrete_Plug_thickness_var,
                    Steel_Sheet_Pile_CofferDam_Load_dict,
                    LM1_Annotation_Dimension_Xdict, LM1_Annotation_Dimension_Ydict,
                    LM1_Annotation_Mleader_Potdict, 
                    ref_point=None,
                    waler_number_map=None, 
                    waler_layer_nums=None,
                    boundary_x=None):
    """绘制立面图1（长边方向）"""
    CofferDam_H = abs(CofferDam_Top_Level - CofferDam_Bottom_Level) * 1000
    LM1_O_point = (ref_point[0], ref_point[1], 0)

    # 承台宽度和起始X
    Cap_startpointX = LM1_O_point[0] - Cap_X / 2
    # 垫层/封底宽度
    Concrete_D = abs(boundary_x[0] - boundary_x[1])

    # 钢板桩
    Pile_startptlst = [(x, LM1_O_point[1], 0) for x in boundary_x]
    Pile_ptlst = [[pt, (pt[0] + i * Pile_H, pt[1], pt[2]),
                    (pt[0] + i * Pile_H, pt[1] + CofferDam_H, pt[2]),
                    (pt[0], pt[1] + CofferDam_H, pt[2]), pt]
                   for pt, i in zip(Pile_startptlst, [-1, 1])]
    for ptlst in Pile_ptlst:
        flat_points = pointlist_extend(ptlst)
        make_polyline(acadmsp, flat_points, '1粗实线')
    LM1_Annotation_Dimension_Xdict['边界'].extend(boundary_x)
    LM1_Annotation_Dimension_Ydict['边界'].extend([Pile_ptlst[0][1][1], Pile_ptlst[0][2][1]])

    # 承台
    Cap_Y_mm = Cap_H * 1000
    Cap_startpointY = LM1_O_point[1] + abs(Cap_Bottom_Level - CofferDam_Bottom_Level) * 1000
    Cap_ptlst = [(Cap_startpointX + nx * Cap_X, Cap_startpointY + ny * Cap_Y_mm, 0)
                  for nx, ny in [(0, 0), (1, 0), (1, 1), (0, 1), (0, 0)]]
    flat_points = pointlist_extend(Cap_ptlst)
    make_polyline(acadmsp, flat_points, '7边界线')
    # 立面图承台X坐标：以LM1_O_point为中心，与立面图坐标系一致
    cap_corner_x = sorted([LM1_O_point[0] - Cap_X / 2, LM1_O_point[0] + Cap_X / 2])
    LM1_Annotation_Dimension_Xdict['承台'].extend(cap_corner_x)
    LM1_Annotation_Dimension_Ydict['承台'].extend([Cap_startpointY, Cap_startpointY + Cap_Y_mm])

    # 垫层/封底
    Concrete_H = Concrete_Blinding_thickness_var * 1000 if Concrete_Blinding_check_var else Concrete_Plug_thickness_var * 1000
    Concrete_startpointX = boundary_x[0]
    Concrete_startpointY = Cap_startpointY
    Concrete_ptlst = [(Concrete_startpointX + nx * Concrete_D, Concrete_startpointY + ny * Concrete_H, 0)
                       for nx, ny in [(0, 0), (1, 0), (1, -1), (0, -1), (0, 0)]]
    flat_points = pointlist_extend(Concrete_ptlst)
    Obj = make_polyline(acadmsp, flat_points, '6标注线')
    Add_Hatch(acadmsp, 'AR-CONC', Obj, '6标注线', 1.0)
    if Concrete_Blinding_check_var:
        LM1_Annotation_Dimension_Ydict['垫层'].extend([Concrete_startpointY - Concrete_H, Concrete_startpointY])
        LM1_Annotation_Mleader_Potdict['垫层'].append(
            [(Concrete_startpointX + Concrete_D / 2, Concrete_startpointY - Concrete_H, 0),
             f'{Concrete_Blinding_thickness_var}cm混凝土垫层'])
    if Concrete_Plug_check_var:
        LM1_Annotation_Dimension_Ydict['封底'].extend([Concrete_startpointY - Concrete_H, Concrete_startpointY])
        LM1_Annotation_Mleader_Potdict['封底'].append(
            [(Concrete_startpointX + Concrete_D / 2, Concrete_startpointY - Concrete_H, 0),
             f'{Concrete_Plug_thickness_var}m混凝土封底'])

    # 围檩截面（所有层）
    Waler_insertname = [v['围檩长边截面'] for v in Walers_Strut_selected_dict.values()]
    # 围檩Y坐标：从桩顶累加间距
    z_accum = CofferDam_Top_Level
    Waler_insertpointY = []
    for v in Walers_Strut_selected_dict.values():
        z_accum -= float(v['间距'])
        Waler_insertpointY.append(LM1_O_point[1] + (z_accum - CofferDam_Bottom_Level) * 1000)
    if waler_number_map is None:
        waler_number_map = {}
    for i in range(len(Waler_insertname)):
        Waler_name = Waler_insertname[i].replace('I', '工') if 'I' in Waler_insertname[i] else Waler_insertname[i]
        insert_ptlst = [(x, Waler_insertpointY[i], 0) for x in boundary_x]
        make_line(acadmsp, insert_ptlst, '3中心线')
        LM1_Annotation_Mleader_Potdict['内支撑'].append(
            [((insert_ptlst[0][0] + insert_ptlst[1][0]) / 2, (insert_ptlst[0][1] + insert_ptlst[1][1]) / 2, 0),
             f'第{i + 1}层内支撑'])
        # 获取该层围檩的W编号
        w_num = waler_layer_nums[i] if waler_layer_nums and i < len(waler_layer_nums) else (i + 1)
        for pt in insert_ptlst:
            RO_Angle = 270 if pt[0] < LM1_O_point[0] else 90
            RO_Angle = (RO_Angle // 90) * 90 / 180 * math.pi
            insert_block_incad_with_block_name(
                acadmsp, vtpnt(pt), f'{Waler_name}分配梁block',
                f'Steel_Sheet_Pile_CofferDam_{Waler_name}', [1, 1, 1], RO_Angle)
            LM1_Annotation_Mleader_Potdict['围檩'].append([pt, f'围檩W{w_num}'])
    LM1_Annotation_Dimension_Ydict['围檩'].extend(Waler_insertpointY)

    # 土层信息
    solid_layer_level_lst = [[v['土层名称'], v['层厚']]
                              for v in Steel_Sheet_Pile_CofferDam_Load_dict.get('Excel', {}).values()]
    solid_start_pointX = boundary_x[1] + 15 * Global_Scale
    solid_start_pointY = LM1_O_point[1] + (Solid_Level - CofferDam_Bottom_Level) * 1000
    insert_pointY0 = solid_start_pointY
    Hatch_lst = [['ANSI33', 0.5], ['CROSS', 0.5], ['GRAVEL', 0.25], ['ANSI31', 0.5], ['AR-SAND', 0.05], ['AR-B816', 0.01]]
    hatch_type = None
    for i in range(len(solid_layer_level_lst)):
        if insert_pointY0 > LM1_O_point[1]:
            from Drawing_AutoCad.Steel_Sheet_Pile_CofferDam_CAD.CofferDam_Resource_CAD import list_random_choice
            hatch_type = list_random_choice(Hatch_lst, 1, hatch_type)[0]
            layer_thickness = float(solid_layer_level_lst[i][1])
            insert_pointY1 = insert_pointY0 - layer_thickness * 1000
            ptX1, ptX2 = solid_start_pointX, solid_start_pointX + 5 * Global_Scale
            ptY1, ptY2 = insert_pointY0, max(LM1_O_point[1], insert_pointY1)
            solid_ptlst = [(ptX1, ptY1, 0), (ptX2, ptY1, 0), (ptX2, ptY2, 0), (ptX1, ptY2, 0), (ptX1, ptY1, 0)]
            flat_points = pointlist_extend(solid_ptlst)
            Obj = make_polyline(acadmsp, flat_points, '6标注线')
            Add_Hatch(acadmsp, hatch_type[0], Obj, '5文本', hatch_type[1] * Global_Scale)
            insert_pointY0 = insert_pointY1


def draw_Dimension_Mleader_Header_Elevation_LM1_incad(acadapp, acadmsp, Pile_H, Global_Scale, Sheet_Pile_typevar,
                                                        CofferDam_Bottom_Level, Solid_Level, Water_Level,
                                                        Steel_Sheet_Pile_CofferDam_Load_dict,
                                                        LM1_Annotation_Dimension_Xdict, LM1_Annotation_Dimension_Ydict,
                                                        LM1_Annotation_Mleader_Potdict, CofferDam_No,):
    """绘制立面图1标注"""
    LM_Annotation_Dimension_X = sorted(set([x for v in LM1_Annotation_Dimension_Xdict.values() for x in v]))
    LM_Annotation_Dimension_Y = sorted(set([y for v in LM1_Annotation_Dimension_Ydict.values() for y in v]))
    if not LM_Annotation_Dimension_X or not LM_Annotation_Dimension_Y:
        return
    Dimension_O_point = (LM_Annotation_Dimension_X[0], LM_Annotation_Dimension_Y[0], 0)
    for i in range(len(LM_Annotation_Dimension_X) - 1):
        Add_Annotation_Linear(acadmsp, (LM_Annotation_Dimension_X[i], Dimension_O_point[1], 0),
                              (LM_Annotation_Dimension_X[i + 1], Dimension_O_point[1], 0),
                              (LM_Annotation_Dimension_X[i], Dimension_O_point[1] - 6 * Global_Scale, 0), 0)
    for i in range(len(LM_Annotation_Dimension_Y) - 1):
        Add_Annotation_Linear(acadmsp, (Dimension_O_point[0] - Pile_H, LM_Annotation_Dimension_Y[i], 0),
                              (Dimension_O_point[0] - Pile_H, LM_Annotation_Dimension_Y[i + 1], 0),
                              (Dimension_O_point[0] - 6 * Global_Scale - Pile_H, LM_Annotation_Dimension_Y[i], 0), 90)
    X1, X2 = LM1_Annotation_Dimension_Xdict['边界']
    Y1, Y2 = LM1_Annotation_Dimension_Ydict['边界']
    Add_Annotation_Linear(acadmsp, (X1, Y1, 0), (X2, Y1, 0), (X1, Y1 - 12 * Global_Scale, 0), 0)
    Add_Annotation_Linear(acadmsp, (X1 - Pile_H, Y1, 0), (X1 - Pile_H, Y2, 0),
                          (X1 - Pile_H - 12 * Global_Scale, Y1, 0), 90)
    for value in LM1_Annotation_Mleader_Potdict.values():
        for valuelst in value:
            startpoint = valuelst[0]
            endpoint = [startpoint[0] + 3 * Global_Scale, startpoint[1] - 6 * Global_Scale, startpoint[2]]
            Add_Annotation_MLeader(acadapp, acadmsp, Global_Scale, startpoint, endpoint, '5文本', 'Song_07', valuelst[1])
    # 立面图1标题
    Header_Point = [(X1 + X2) / 2, Y2 + 18 * Global_Scale, 0]
    cofferdam_numberlst = re.findall(r'\d+', CofferDam_No)
    cofferdam_num = cofferdam_numberlst[0] if cofferdam_numberlst != [] else 'XX'
    Add_Header(acadmsp, Global_Scale, Header_Point, f'{cofferdam_num}#墩{Sheet_Pile_typevar}围堰立面布置图', 'Song_07')

    # ---- 结构类标高 ----
    struct_level_lst = []
    for key, value in LM1_Annotation_Dimension_Ydict.items():
        if key == '承台' and value:
            ylst, namelst = [value[0]], ['承台底']
        elif key == '边界' and value:
            ylst, namelst = value, ['围堰底', '围堰顶']
        elif key == '围檩' and value:
            ylst = value
            namelst = [f'第{i+1}层围檩' for i, y in enumerate(ylst)]
        elif key == '垫层' and value:
            ylst, namelst = [value[0]], ['垫层']
        elif key == '封底' and value:
            ylst, namelst = [value[0]], ['封底砼底']
        else:
            continue
        hlst = [round((y - Y1) / 1000 + CofferDam_Bottom_Level, 3) for y in ylst]
        struct_level_lst.extend([(y, name, h) for y, name, h in zip(ylst, namelst, hlst)])
    struct_level_lst = sorted(set(tuple(struct_level_lst)), key=lambda x: x[0], reverse=True)
    level_param1 = -1
    level_param2 = 1
    for i in range(len(struct_level_lst)):
        y, name, h = struct_level_lst[i]
        # 垫层标高朝下，其余朝上
        if '垫层' in name:
            level_param2 = -1
        insert_point = (Dimension_O_point[0] - 24 * Global_Scale, y, 0)
        level_txt = f"{name}标高:+{h}m" if h >= 0 else f"{name}标高:{h}m"
        Add_Elevation_Symbol(acadmsp, Global_Scale, insert_point, level_txt, 'Song_07', level_param1, level_param2)

    # ---- 土层相关标高 ----
    solid_layer_level_lst = [[v['土层名称'], v['层厚'], v['重度'], v['黏聚力'], v['内摩擦角']]
                              for v in Steel_Sheet_Pile_CofferDam_Load_dict.get('Excel', {}).values()]
    start_pointX = X2 + 24 * Global_Scale
    start_pointY = Y1 + (Solid_Level - CofferDam_Bottom_Level) * 1000
    insert_pointY0 = start_pointY
    level_param1, level_param2 = 1, 1
    for i in range(len(solid_layer_level_lst)):
        if insert_pointY0 > Y1:
            layer_thickness = float(solid_layer_level_lst[i][1]) * 1000
            insert_pointY1 = insert_pointY0 - layer_thickness
            next_layer_y = max(insert_pointY0 - layer_thickness, Y1)
            layer_thickness = insert_pointY0 - next_layer_y
            insert_point1 = [start_pointX, max(insert_pointY0, Y1), 0]
            insert_point2 = [start_pointX + 15 * Global_Scale, insert_pointY0 - layer_thickness / 2 + 5.4 * Global_Scale, 0]
            insert_point3 = [start_pointX + 15 * Global_Scale, insert_pointY0 - layer_thickness / 2 + 1.8 * Global_Scale, 0]
            insert_point4 = [start_pointX + 15 * Global_Scale, insert_pointY0 - layer_thickness / 2 - 1.8 * Global_Scale, 0]
            h = round((insert_pointY0 - Y1) / 1000 + CofferDam_Bottom_Level, 3)
            level_txt = f"标高:+{h}m" if h >= 0 else f"标高:{h}m"
            Add_Elevation_Symbol(acadmsp, Global_Scale, insert_point1, level_txt, 'Song_07', level_param1, level_param2)
            Add_Text(acadmsp, insert_point2, solid_layer_level_lst[i][0], 2.5 * Global_Scale, 'Song_07', '5文本', 9)
            Add_Text(acadmsp, insert_point3, f'r={solid_layer_level_lst[i][2]}kN/m³', 2.5 * Global_Scale, 'Song_07', '5文本', 9)
            Add_Text(acadmsp, insert_point4, f'c={solid_layer_level_lst[i][3]}kPa,φ={solid_layer_level_lst[i][4]}°', 2.5 * Global_Scale, 'Song_07', '5文本', 9)
            insert_pointY0 = insert_pointY1
    level_txt = f"标高:+{CofferDam_Bottom_Level}m" if CofferDam_Bottom_Level >= 0 else f"标高:{CofferDam_Bottom_Level}m"
    Add_Elevation_Symbol(acadmsp, Global_Scale, [start_pointX, Y1, 0], level_txt, 'Song_07', level_param1, level_param2)

    # ---- 水位相关标高 ----
    level_txt = f"水位:+{Water_Level}m" if Water_Level >= 0 else f"水位:{Water_Level}m"
    Add_Elevation_Symbol(acadmsp, Global_Scale, [start_pointX, Y1 + (Water_Level-CofferDam_Bottom_Level)*1000, 0], level_txt, 'Song_07', level_param1, level_param2)


# ============================================================
# 立面图2（短边方向）
# ============================================================

def draw_LM2_incad(acadmsp, Cap_Y, Pile_H, Walers_Strut_selected_dict, Cap_H,
                    CofferDam_Top_Level, CofferDam_Bottom_Level, Cap_Bottom_Level,
                    Concrete_Blinding_check_var, Concrete_Blinding_thickness_var,
                    Concrete_Plug_check_var, Concrete_Plug_thickness_var,
                    LM2_Annotation_Dimension_Xdict, LM2_Annotation_Dimension_Ydict,
                    LM2_Annotation_Mleader_Potdict, 
                    ref_point=None,
                    waler_number_map=None, 
                    waler_layer_nums=None,
                    boundary_x=None):
    """绘制立面图2（短边方向）"""
    CofferDam_H = abs(CofferDam_Top_Level - CofferDam_Bottom_Level) * 1000
    LM2_O_point = (ref_point[0], ref_point[1], 0)

    # 承台宽度和起始X
    Cap_startpointX = LM2_O_point[0] - Cap_Y / 2
    # 垫层/封底宽度
    Concrete_D = abs(boundary_x[0] - boundary_x[1])

    Pile_startptlst = [(LM2_O_point[0] - Concrete_D / 2, LM2_O_point[1], 0),
                        (LM2_O_point[0] + Concrete_D / 2, LM2_O_point[1], 0)]
    Pile_ptlst = [[pt, (pt[0] + i * Pile_H, pt[1], pt[2]),
                    (pt[0] + i * Pile_H, pt[1] + CofferDam_H, pt[2]),
                    (pt[0], pt[1] + CofferDam_H, pt[2]), pt]
                   for pt, i in zip(Pile_startptlst, [-1, 1])]
    for ptlst in Pile_ptlst:
        flat_points = pointlist_extend(ptlst)
        make_polyline(acadmsp, flat_points, '1粗实线')
    LM2_Annotation_Dimension_Xdict['边界'].extend([Pile_startptlst[0][0], Pile_startptlst[1][0]])
    LM2_Annotation_Dimension_Ydict['边界'].extend([Pile_ptlst[0][1][1], Pile_ptlst[0][2][1]])

    Cap_Y_mm = Cap_H * 1000
    Cap_startpointY = LM2_O_point[1] + abs(Cap_Bottom_Level - CofferDam_Bottom_Level) * 1000
    Cap_ptlst = [(Cap_startpointX + nx * Cap_Y, Cap_startpointY + ny * Cap_Y_mm, 0)
                  for nx, ny in [(0, 0), (1, 0), (1, 1), (0, 1), (0, 0)]]
    flat_points = pointlist_extend(Cap_ptlst)
    make_polyline(acadmsp, flat_points, '7边界线')
    LM2_Annotation_Dimension_Xdict['承台'].extend([Cap_startpointX, Cap_startpointX + Cap_Y])
    LM2_Annotation_Dimension_Ydict['承台'].extend([Cap_startpointY, Cap_startpointY + Cap_Y_mm])

    Concrete_H = Concrete_Blinding_thickness_var * 1000 if Concrete_Blinding_check_var else Concrete_Plug_thickness_var * 1000
    Concrete_startpointX = sorted(LM2_Annotation_Dimension_Xdict['边界'])[0]
    Concrete_startpointY = Cap_startpointY
    Concrete_ptlst = [(Concrete_startpointX + nx * Concrete_D, Concrete_startpointY + ny * Concrete_H, 0)
                       for nx, ny in [(0, 0), (1, 0), (1, -1), (0, -1), (0, 0)]]
    flat_points = pointlist_extend(Concrete_ptlst)
    Obj = make_polyline(acadmsp, flat_points, '6标注线')
    Add_Hatch(acadmsp, 'AR-CONC', Obj, '6标注线', 1.0)
    if Concrete_Blinding_check_var:
        LM2_Annotation_Dimension_Ydict['垫层'].extend([Concrete_startpointY - Concrete_H, Concrete_startpointY])
        LM2_Annotation_Mleader_Potdict['垫层'].append(
            [(Concrete_startpointX + Concrete_D / 2, Concrete_startpointY - Concrete_H, 0),
             f'{Concrete_Blinding_thickness_var}cm混凝土垫层'])
    if Concrete_Plug_check_var:
        LM2_Annotation_Dimension_Ydict['封底'].extend([Concrete_startpointY - Concrete_H, Concrete_startpointY])
        LM2_Annotation_Mleader_Potdict['封底'].append(
            [(Concrete_startpointX + Concrete_D / 2, Concrete_startpointY - Concrete_H, 0),
             f'{Concrete_Plug_thickness_var}m混凝土封底'])

    Waler_insertname = [v['围檩长边截面'] for v in Walers_Strut_selected_dict.values()]
    # 围檩Y坐标：从桩顶累加间距
    z_accum = CofferDam_Top_Level
    Waler_insertpointY = []
    for v in Walers_Strut_selected_dict.values():
        z_accum -= float(v['间距'])
        Waler_insertpointY.append(LM2_O_point[1] + (z_accum - CofferDam_Bottom_Level) * 1000)
    if waler_number_map is None:
        waler_number_map = {}
    for i in range(len(Waler_insertname)):
        Waler_name = Waler_insertname[i].replace('I', '工') if 'I' in Waler_insertname[i] else Waler_insertname[i]
        insert_ptlst = [(x, Waler_insertpointY[i], 0) for x in LM2_Annotation_Dimension_Xdict['边界']]
        make_line(acadmsp, insert_ptlst, '3中心线')
        LM2_Annotation_Mleader_Potdict['内支撑'].append(
            [((insert_ptlst[0][0] + insert_ptlst[1][0]) / 2, (insert_ptlst[0][1] + insert_ptlst[1][1]) / 2, 0),
             f'第{i + 1}层内支撑'])
        # 获取该层围檩的W编号
        w_num = waler_layer_nums[i] if waler_layer_nums and i < len(waler_layer_nums) else (i + 1)
        for pt in insert_ptlst:
            RO_Angle = 270 if pt[0] < LM2_O_point[0] else 90
            RO_Angle = (RO_Angle // 90) * 90 / 180 * math.pi
            insert_block_incad_with_block_name(
                acadmsp, vtpnt(pt), f'{Waler_name}分配梁block',
                f'Steel_Sheet_Pile_CofferDam_{Waler_name}', [1, 1, 1], RO_Angle)
            LM2_Annotation_Mleader_Potdict['围檩'].append([pt, f'围檩W{w_num}'])
    LM2_Annotation_Dimension_Ydict['围檩'].extend(Waler_insertpointY)


def draw_Dimension_Mleader_Header_Elevation_LM2_incad(acadapp, acadmsp, Pile_H, Global_Scale,
                                                        LM2_Annotation_Dimension_Xdict, LM2_Annotation_Dimension_Ydict,
                                                        LM2_Annotation_Mleader_Potdict):
    """绘制立面图2标注"""
    LM_Annotation_Dimension_X = sorted(set([x for v in LM2_Annotation_Dimension_Xdict.values() for x in v]))
    LM_Annotation_Dimension_Y = sorted(set([y for v in LM2_Annotation_Dimension_Ydict.values() for y in v]))
    if not LM_Annotation_Dimension_X or not LM_Annotation_Dimension_Y:
        return
    Dimension_O_point = (LM_Annotation_Dimension_X[0], LM_Annotation_Dimension_Y[0], 0)
    for i in range(len(LM_Annotation_Dimension_X) - 1):
        Add_Annotation_Linear(acadmsp, (LM_Annotation_Dimension_X[i], Dimension_O_point[1], 0),
                              (LM_Annotation_Dimension_X[i + 1], Dimension_O_point[1], 0),
                              (LM_Annotation_Dimension_X[i], Dimension_O_point[1] - 6 * Global_Scale, 0), 0)
    for i in range(len(LM_Annotation_Dimension_Y) - 1):
        Add_Annotation_Linear(acadmsp, (Dimension_O_point[0] - Pile_H, LM_Annotation_Dimension_Y[i], 0),
                              (Dimension_O_point[0] - Pile_H, LM_Annotation_Dimension_Y[i + 1], 0),
                              (Dimension_O_point[0] - 6 * Global_Scale - Pile_H, LM_Annotation_Dimension_Y[i], 0), 90)
    X1, X2 = LM2_Annotation_Dimension_Xdict['边界']
    Y1, Y2 = LM2_Annotation_Dimension_Ydict['边界']
    Add_Annotation_Linear(acadmsp, (X1, Y1, 0), (X2, Y1, 0), (X1, Y1 - 12 * Global_Scale, 0), 0)
    Add_Annotation_Linear(acadmsp, (X1 - Pile_H, Y1, 0), (X1 - Pile_H, Y2, 0),
                          (X1 - Pile_H - 12 * Global_Scale, Y1, 0), 90)
    for value in LM2_Annotation_Mleader_Potdict.values():
        for valuelst in value:
            startpoint = valuelst[0]
            endpoint = [startpoint[0] + 3 * Global_Scale, startpoint[1] - 6 * Global_Scale, startpoint[2]]
            Add_Annotation_MLeader(acadapp, acadmsp, Global_Scale, startpoint, endpoint, '5文本', 'Song_07', valuelst[1])
    Header_Point = [(X1 + X2) / 2, Y2 + 18 * Global_Scale, 0]
    Add_Header(acadmsp, Global_Scale, Header_Point, '1---1', 'Song_07')


# ============================================================
# 材料表（跨层合并编号）
# ============================================================

def draw_Material_List_in_cad(acadmsp, CAD_O_point, Global_Scale,
                               PM_Annotation_Dimension_Xdict, PM_Annotation_Dimension_Ydict,
                               PM_Annotation_Mleader_Pot, Material_dict,
                               waler_number_map=None, global_strut_numbering=None):
    """生成材料表 — 跨层合并相同截面+长度的构件"""
    position_line_xlst = PM_Annotation_Dimension_Xdict['边界']
    position_line_ylst = PM_Annotation_Dimension_Ydict['边界']
    X = round(abs(position_line_xlst[0] - position_line_xlst[1]), 1)
    Y = round(abs(position_line_ylst[0] - position_line_ylst[1]), 1)
    O_point = [CAD_O_point[0], CAD_O_point[1], CAD_O_point[2]]

    # 列宽
    col_widths = [10, 20, 35, 20, 10, 15, 15, 25]
    col_widths = [w * Global_Scale for w in col_widths]
    Width = sum(col_widths)
    row_count = len(col_widths)

    # 跨层合并围檩编号（新格式：PM_Annotation_Mleader_Pot['围檩'] 的 key 为 (sec, mat, length)）
    waler_merged = {}  # (截面, 材质, 长度) → 数量
    for (sec, mat, L), ptlst in PM_Annotation_Mleader_Pot.get('围檩', {}).items():
        key = (sec, mat, round(L, 1))
        waler_merged[key] = waler_merged.get(key, 0) + len(ptlst)

    if waler_number_map is None:
        waler_number_map = {}

    for (sec, mat, L), qty in waler_merged.items():
        weight_info = SEC_Weight_per_meter.get(sec, [0, 0])
        singlelet = round(weight_info[1] * L / 1000, 1) if weight_info[1] else 0
        w_num = 1
        for (ws, wm, wl), wn in waler_number_map.items():
            if ws == sec and wm == mat and abs(wl - L) < 50:
                w_num = wn
                break
        Material_dict['围檩']['名称'].append(f'围檩W{w_num}')
        Material_dict['围檩']['规格'].append(f'{sec}×{int(L)}')
        Material_dict['围檩']['材质'].append(_mat_label(mat))
        Material_dict['围檩']['数量'].append(qty)
        Material_dict['围檩']['单重'].append(singlelet)
        Material_dict['围檩']['总重'].append(round(qty * singlelet, 1))
        Material_dict['围檩']['备注'].append('')

    # 跨层合并内支撑编号（新格式：PM_Annotation_Mleader_Pot['内支撑'] 的 key 为 (type, sec, mat, length)）
    strut_merged = {}  # (type, sec, mat, length) → 数量
    for key_tuple, ptlst in PM_Annotation_Mleader_Pot.get('内支撑', {}).items():
        strut_type, sec, mat, L = key_tuple
        if sec == '/':
            continue
        merge_key = (strut_type, sec, mat, round(L, 1))
        strut_merged[merge_key] = strut_merged.get(merge_key, 0) + len(ptlst)

    if global_strut_numbering is None:
        global_strut_numbering = {}
    # 构建 (type, sec, mat, length) → N编号 的映射（含取整后的长度）
    strut_key_to_num = {}
    for key_tuple, num in global_strut_numbering.items():
        strut_type, sec, mat, L = key_tuple
        strut_key_to_num[(strut_type, sec, mat, round(L, 1))] = num

    for (strut_type, sec, mat, L), qty in strut_merged.items():
        weight_info = SEC_Weight_per_meter.get(sec, [0, 0])
        singlelet = round(weight_info[1] * L / 1000, 1) if weight_info[1] else 0
        n_num = strut_key_to_num.get((strut_type, sec, mat, L), 1)
        Material_dict['内支撑']['名称'].append(f'内支撑N{n_num}')
        if '钢管桩' in sec:
            m = re.match(r'(\d+)[Xx×](\d+)', sec)
            D, t = int(m.group(1)), int(m.group(2))
            Material_dict['内支撑']['规格'].append(f'φ{D}×{t}×{int(L)}')
        else:
            Material_dict['内支撑']['规格'].append(f'{sec}×{int(L)}')
        Material_dict['内支撑']['材质'].append(_mat_label(mat))
        Material_dict['内支撑']['数量'].append(qty)
        Material_dict['内支撑']['单重'].append(singlelet)
        Material_dict['内支撑']['总重'].append(round(qty * singlelet, 1))
        Material_dict['内支撑']['备注'].append('')

    # 汇总材料行：各分类列数应一致，缺失列补空串，避免索引越界
    material_rows = []
    for _m_item in Material_dict.values():
        _m_names = _m_item.get('名称', [])
        for _i in range(len(_m_names)):
            if not str(_m_names[_i]).strip():
                continue
            material_rows.append([
                _m_names[_i],
                _m_item['规格'][_i] if _i < len(_m_item.get('规格', [])) else '',
                _m_item['材质'][_i] if _i < len(_m_item.get('材质', [])) else '',
                _m_item['数量'][_i] if _i < len(_m_item.get('数量', [])) else '',
                _m_item['单重'][_i] if _i < len(_m_item.get('单重', [])) else '',
                _m_item['总重'][_i] if _i < len(_m_item.get('总重', [])) else '',
                _m_item['备注'][_i] if _i < len(_m_item.get('备注', [])) else '',
            ])

    # 绘制表格
    line_count = len(material_rows)
    Height = (line_count + 2) * 6 * Global_Scale

    Header_Point = [O_point[0], O_point[1] + Height / 2 + 10 * Global_Scale, 0]
    Add_Header(acadmsp, Global_Scale, Header_Point, '材料表', 'Song_07')

    ptlst = [(O_point[0] + nx * Width / 2, O_point[1] + ny * Height / 2, 0)
              for nx, ny in [(1, 1), (-1, 1), (-1, -1), (1, -1), (1, 1)]]
    flat_points = pointlist_extend(ptlst)
    make_polyline(acadmsp, flat_points, '1粗实线', 0.4 * Global_Scale)
    Add_Text(acadmsp, (ptlst[0][0], ptlst[0][1] + 1.5 * Global_Scale, ptlst[0][2]),
             '以单个墩围堰计', 2.5 * Global_Scale, 'Song_07', '5文本', 14)

    ptstartX, ptendX = O_point[0] - Width / 2, O_point[0] + Width / 2
    ptstartY, ptendY = O_point[1] + Height / 2, O_point[1] - Height / 2

    txt_insert_X, txt_insert_Y = [], []
    firstline_txt_insert_X, firstline_txt_insert_Y = [], []

    for i in range(line_count + 2):
        ptY = ptstartY - (i + 1) * 6 * Global_Scale
        if i != line_count + 1:
            make_line(acadmsp, [(ptstartX, ptY, 0), (ptendX, ptY, 0)], '2细实线')
        txt_insert_Y.append(ptY + 3 * Global_Scale)
        firstline_txt_insert_Y.append(ptY + 3 * Global_Scale)

    ptX = ptstartX
    for i in range(row_count):
        if i != 0:
            make_line(acadmsp, [(ptX, ptstartY, 0), (ptX, ptendY + 6 * Global_Scale, 0)], '2细实线')
        txt_insert_X.append(ptX + 2 * Global_Scale if i == 2 else ptX + col_widths[i] / 2)
        firstline_txt_insert_X.append(ptX + col_widths[i] / 2)
        ptX += col_widths[i]

    # 填写内容
    txt_dict = {0: ['编号', '名称', '规    格', '材质/牌号', '数量', '单重(kg)', '总重(kg)', '备注']}
    txt_line = 1
    print(Material_dict)
    for _row in material_rows:
        txt_dict[txt_line] = [txt_line] + _row
        txt_line += 1

    for iy in range(len(txt_insert_Y) - 1):
        for ix in range(len(txt_insert_X)):
            txt = txt_dict[iy][ix]
            pt = (firstline_txt_insert_X[ix], firstline_txt_insert_Y[iy], 0) if iy == 0 else (txt_insert_X[ix], txt_insert_Y[iy], 0)
            AlignNum = 9 if ix == 2 and iy > 0 else 4
            Add_Text(acadmsp, pt, txt, 2.5 * Global_Scale, 'Song_07', '5文本', AlignNum)

    Steel_all_weight = sum([v[6] for k, v in txt_dict.items() if k != 0 and '混凝土' not in str(v[1])])
    Concrete_all_weight = sum([v[6] for k, v in txt_dict.items() if '混凝土' in str(v[1])])
    all_weight_line_txt = f'合计钢材 {round(float(Steel_all_weight), 1)} kg，混凝土 {round(float(Concrete_all_weight), 1)} m³。'
    insert_pt = (ptstartX + sum(col_widths[:-1]), O_point[1] - Height / 2 + 3 * Global_Scale, 0)
    Add_Text(acadmsp, insert_pt, all_weight_line_txt, 2.5 * Global_Scale, 'Song_07', '5文本', 11)
    # 返回材料表底部Y坐标（用于附注定位）
    return O_point[1] - Height / 2


# ============================================================
# 图框附注
# ============================================================

def draw_frame_notes_in_cad(acadmsp, Global_Scale, PM_Annotation_Dimension_Xdict, PM_Annotation_Dimension_Ydict,
                            notes_x=None, notes_y=None, frame_insertpoint=None):
    """绘制附注（多行文本，左对齐，指定定位点）。图框已在主函数中插入。"""

    # 附注内容
    notes_lst = ['附注：', '1、图中尺寸除高程以m计外，其余均以mm计。']

    # 附注尺寸
    notes_width = 60 * Global_Scale
    notes_height = 10 * Global_Scale
    text_height = 2.5 * Global_Scale

    # 附注定位点（左下角）
    if notes_x is not None and notes_y is not None:
        insert_x = notes_x
        insert_y = notes_y
    elif frame_insertpoint is not None:
        insert_x = frame_insertpoint[0] + 5 * Global_Scale
        insert_y = frame_insertpoint[1] + 5 * Global_Scale
    else:
        boundary_x_min = sorted(PM_Annotation_Dimension_Xdict['边界'])[0]
        boundary_y_min = sorted(PM_Annotation_Dimension_Ydict['边界'])[0]
        insert_x = boundary_x_min - 90 * Global_Scale
        insert_y = boundary_y_min - 60 * Global_Scale

    # 使用 MText 多行文本，左对齐
    insertPnt = vtpnt((insert_x, insert_y, 0))
    text_string = '\\P'.join(notes_lst)
    MtextObj = acadmsp.AddMText(insertPnt, notes_width, text_string)
    MtextObj.Height = text_height
    MtextObj.AttachmentPoint = 1  # Top Left（左对齐）
    try:
        MtextObj.TextStyleName = 'Song_07'
    except:
        pass
    try:
        MtextObj.Layer = '5文本'
    except:
        MtextObj.Layer = '0'


if __name__ == "__main__":
    print("CofferDam_CAD_Draft_Generate.py - 钢板桩围堰 CAD 绘图引擎")
    print("支持: 钢板桩 / 锁扣钢管桩 / 多层内支撑")
