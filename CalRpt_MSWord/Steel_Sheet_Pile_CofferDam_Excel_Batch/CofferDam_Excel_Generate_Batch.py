# 1. 标准库
import os
import re
import sys
import time
import glob
from pathlib import Path

# 2. 第三方库
import openpyxl

# 3. 本地模块
sys.path.append(str(Path(__file__).parent.parent.parent))
from General.Midas import MidasAPI, get_midas_version, API_DIST_FORCE_Unit, Picture_Pre, MidasConnectionError

sys.path.append(str(Path(__file__).parent))
from CalRpt_MSWord.Steel_Sheet_Pile_CofferDam_Excel_Batch.CofferDam_Cal_Handler_Batch import Model_cal_res, stability_cal
from CalRpt_MSWord.Steel_Sheet_Pile_CofferDam_Cal.CofferDam_Midas_Model_Post          import Get_Waler_info_from_Model, Get_Strut_info_from_Model
from CalRpt_MSWord.Steel_Sheet_Pile_CofferDam_Cal.CofferDam_Cal_File_Manipulation     import load_cofferdam_params_for_cal, _find_sheet, load_section_library, load_material_library
from CalRpt_MSWord.Steel_Sheet_Pile_CofferDam_Cal.CofferDam_Cal_Docx_Handler          import case_info, comp_info, Active_Earth_Pressure, Passive_Earth_Pressure, allowable_f


# ======================================================================================
# 辅助函数：Excel 参数读取
# ======================================================================================
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


def _build_header_map(ws, header_row):
    """读取指定行，返回 {表头名: 列号} 映射"""
    hmap = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if v is not None:
            h = str(v).strip()
            if h: hmap[h] = c
    return hmap


def _write_cell(ws, row, hm, header_name, value):
    """按表头名查找列号并写入值，列不存在则跳过"""
    col = hm.get(header_name)
    if col:
        ws.cell(row, col, value)


def _check_ok(result):
    """检查所有计算结果是否满足限值要求，跳过 '/' 占位值"""
    def _stress_ok(d):
        for v in d.values():
            if len(v) >= 2 and v[0] != '/' and v[1] != '/' and v[0] >= v[1]:
                return False
        return True

    def _stability_ok(d):
        for v in d.values():
            if len(v) >= 2 and v[0] != '/' and v[1] != '/' and v[0] < v[1]:
                return False
        return True

    if not _stress_ok(result.get('pile', {})): return False
    for w in result.get('waler', {}).values():
        if not _stress_ok(w): return False
    for s in result.get('strut', {}).values():
        if not _stress_ok(s): return False
    if not _stability_ok(result.get('stability', {})): return False
    return True


def _write_result(ws, row, result, hm):
    """将单个围堰的计算结果写入指定行"""
    w = lambda header, val: _write_cell(ws, row, hm, header, val)
    v0 = lambda lst: lst[0] if lst else None  # 取列表首项，空则None

    # 规范列表 -> 逗号拼接
    std_list = result.get('standard', [])
    w('计算依据', ', '.join(str(s) for s in std_list) if std_list else None)

    # 桩
    pile = result.get('pile', {})
    w('支护桩组合应力(MPa)', v0(pile.get('cb')))
    w('支护桩变形(mm)', v0(pile.get('d')))

    # 围檩 1~5
    for i in range(1, 6):
        wl = result.get('waler', {}).get(str(i), {})
        w(f'围檩{i}组合应力(MPa)', v0(wl.get('cb')))
        w(f'围檩{i}剪应力(MPa)', v0(wl.get('ssz')))
        w(f'围檩{i}变形(mm)', v0(wl.get('d')))

    # 内支撑 1~5
    for i in range(1, 6):
        st = result.get('strut', {}).get(str(i), {})
        w(f'对撑{i}截面压弯稳定应力(MPa)', v0(st.get('sigmax')))
        w(f'斜撑{i}截面压弯稳定应力(MPa)', v0(st.get('sigmay')))

    # 稳定性
    stab = result.get('stability', {})
    w('整体滑动稳定性系数', v0(stab.get('overall')))
    w('抗倾覆稳定性系数', v0(stab.get('anti_overturn')))
    w('抗隆起稳定性系数', v0(stab.get('heave_resistant')))
    w('渗透稳定性系数', v0(stab.get('seepage')))

    # 结论
    w('结论', '满足要求' if _check_ok(result) else '不满足要求')


def batch_write_results(excel_path, result_list, output_path=None, header_row=2, data_start=3):
    wb = openpyxl.load_workbook(excel_path)
    ws = wb['矩形围堰设计参数'] if '矩形围堰设计参数' in wb.sheetnames else wb[wb.sheetnames[0]]
    hm = _build_header_map(ws, header_row)
    ok, fail, missing = 0, 0, []
    for item in result_list:
        if item == {}:
            continue
        pname = item.get('project_name', '')
        pidx = item.get('cofferdam_index', '')
        result = item.get('result', {})
        if not pname or not pidx:
            fail += 1
            continue
        found = False
        for r in range(data_start, ws.max_row + 1):
            rname = str(ws.cell(r, hm.get('项目名称', 0)).value or '').strip()
            ridx = str(ws.cell(r, hm.get('围堰编号', 0)).value or '').strip()
            if rname == pname and ridx == pidx:
                _write_result(ws, r, result, hm)
                ok += 1
                print(f'成功: {pname}-{pidx} -> R{r}')
                found = True
                break
        if not found:
            missing.append(f'{pname}-{pidx}')
            fail += 1
    if output_path is None:
        output_path = f'{os.path.splitext(excel_path)[0]}_结果.xlsx'
    wb.save(output_path)
    print(f'\n完成: 成功{ok}, 失败{fail}')
    if missing: print(f'未找到: {missing}')
    return output_path

# ======================================================================================
# 主函数：批量生成计算书
# ======================================================================================
def make_all_doc_batch(excel_path, model_folder=None, pbar=None, svar=None, root=None, acc = 3, cal_config=None):
    # 记录开始时间
    start_time = time.time()

    # 获取文件夹内的所有mcb模型的名称
    files = os.listdir(model_folder)
    mcb_lst = [f for f in files if f.endswith(".mcb")]
    print(mcb_lst)

    # 定义任务清单
    total_steps = 5*len(mcb_lst)
    current_step = 0
    def update_p(desc):
        nonlocal current_step
        current_step += 1
        if pbar and svar and root:
            pbar['value'] = (current_step / total_steps) * 100
            svar.set(desc)
            root.update()

    # 存放全部项目对应的表格结果
    cofferdam_param_table_proj_result_lst = []

    # 提前打开 Excel 工作簿和表头映射，供逐行写入使用
    _wb = openpyxl.load_workbook(excel_path)
    _ws = _wb['矩形围堰设计参数'] if '矩形围堰设计参数' in _wb.sheetnames else _wb[_wb.sheetnames[0]]
    _hm = _build_header_map(_ws, 2)  # 表头在第2行
    _data_start_row = 3              # 数据从第3行开始
    _api_disconnected = False        # 标记 API 是否断开

    for mcb in mcb_lst:
        try:
            mcb_name = mcb.split('.')[0]
            proj, index = mcb_name.split('-')
            proj_key = f'{proj}{index}'
            print('当前计算项目名:', proj_key)

            docx_save_path = model_folder

            # 指定计算模型，则图片也会存到模型所在的文件夹内
            Midas_path = os.path.join(model_folder, mcb) # 文件目录路径
            # 记录流程开始前目录中已有的图片文件，用于后续精确清理中间图片
            _pre_existing_imgs = set(glob.glob(os.path.join(docx_save_path, '*.jpg')) + glob.glob(os.path.join(docx_save_path, '*.png'))) if docx_save_path else set()
            Midas_name = mcb # 文件名
        
            # ====== 连接excel参数 ===================================================================================================================================================================================
            # ====== 连接excel参数 ===================================================================================================================================================================================
            
            update_p("正在匹配围堰数据库")
            params = load_cofferdam_params_for_cal(excel_path, proj_key)
            for k, v in params.items():
                print(k, v)
            _, _, Waler_section_dict, Strut_section_dict = load_waler_strut_sections(excel_path)
            Steel_dict, Concrete_dict, Rebar_dict = load_material_library(excel_path)

            # 基本参数
            SECtype = params["SECtype"]
            pile_material = params["pile_material"]
            Solid_Level = params["Solid_Level"]
            Water_Level = params["Water_Level"]
            Cap_X = params["Cap_X"]
            Cap_Y = params["Cap_Y"]
            Cap_H = params["Cap_H"]
            Cap_Bottom_Level = params["Cap_Bottom_Level"]
            CofferDam_Top_Level = params["CofferDam_Top_Level"]
            CofferDam_Bottom_Level = params["CofferDam_Bottom_Level"]
            CofferDam_L = params["CofferDam_L"]
            CofferDam_a = params["CofferDam_a"]
            CofferDam_b = params["CofferDam_b"]
            Ea_p0 = params["Ea_p0"]
            Ep_p0 = 0.0 # 基坑内附加应力（暂硬编码）
            CofferDam_Level = params["CofferDam_Level"]
            Drawdown_height = params["Drawdown_height"]
            Waterdown_height = params["Waterdown_height"]
            Excavation_face_dewater = params["Excavation_face_dewater"]
            conc_type = params["conc_type"]
            conc_grade = params["conc_grade"]
            conc_t = params["conc_t"]
            Basic_Solid_Layer_Thickness = params["Basic_Solid_Layer_Thickness"]
            gamma_0 = params["gamma_0"]
            Method_for_Pressures_var = params["Method_for_Pressures_var"]

            # 规范
            steel_standard = params["steel_standard"]

            # 截面参数
            gbz_name = params["gbz_name"]
            gbz_sec_info_dict = params["gbz_sec_info_dict"]
            pile_b = params["pile_b"]
            ggz_D = params["ggz_D"]
            ggz_t = params["ggz_t"]
            ggz_name = params["ggz_name"]
            ggz_A = params["ggz_A"]
            ggz_I = params["ggz_I"]
            ggz_W = params["ggz_W"]

            # 土层数据
            Solid_data_dict = params["Solid_data_dict"]
            Excel_Data = params["Excel_Data"]
            solid_layer_lst = params["solid_layer_lst"]

            # 施工阶段参数
            stage_dict = params["stage_dict"]

            # 嵌固深度
            Ld = params['Ld']
            # 圆弧滑动竖分土条厚度
            slice_width = params['slice_width']
            # 基坑深度
            Foundation_Pit_Depth = params['Foundation_Pit_Depth']
            # 基坑底标高
            Concrete_Bottom_Level = params['Concrete_Bottom_Level']
            print("Success: 围堰数据库信息 读取成功")
            print('')

            # ====== 获取模型基本信息 ===================================================================================================================================================================================
            # ====== 获取模型基本信息 ===================================================================================================================================================================================
            
            update_p("正在获取模型结果")

            # # response变量代表了用户是否需要读取模型的结果
            # response = False
            # if response:
            arguments = {"Argument" : Midas_path}
            # 打开
            MidasAPI("POST" , "/doc/open" , arguments)
            # 运行
            # 根据文件名生成结果文件名
            Analysis_result_outfile_name = Midas_name.replace(".mcb", ".OUT")
            Analysis_result_ca1file_name = Midas_name.replace(".mcb", ".CA1")
            print(Analysis_result_outfile_name)
            print(Analysis_result_ca1file_name)
            # 生成对应的路径
            Analysis_result_file_path1 = os.path.join(model_folder, Analysis_result_outfile_name)
            Analysis_result_file_path2 = os.path.join(model_folder, Analysis_result_ca1file_name)
            # 文件运行
            if os.path.exists(Analysis_result_file_path1) and os.path.exists(Analysis_result_file_path2):
                # 两个结果文件存在，跳过不运行
                print("检测到结果文件，正在读取...")
            else:
                # 两个结果文件不存在，运行
                print("未检测到结果文件，开始运行。")
                MidasAPI("POST", "/doc/Anal", {})


            # 读取midas版本号
            version = get_midas_version()
            print("Midas Version:", version)
            
            # 确保节点→标高读取值单位为M
            API_DIST_FORCE_Unit("N", 'MM')
            
            # 判断是否存在施工阶段
            has_stage = MidasAPI("GET", "/db/STAG") != {'message': ''}
            print('是否存在施工阶段:', has_stage)

            # 在前处理阶段获取所有信息
            Picture_Pre(docx_save_path, False)
            Matl_res = MidasAPI("GET", "/db/MATL")
            SECT_res = MidasAPI("GET", "/db/SECT")
            Node_res = MidasAPI("GET", "/db/NODE")
            Elem_res = MidasAPI("GET", "/db/ELEM")
            Structure_group_res = MidasAPI("GET", "/db/GRUP")# 组信息
            Structure_group_namelst = [value['NAME'] for value in Structure_group_res['GRUP'].values()]
            print('结构组名称:',Structure_group_namelst)
            print('')

            # 恢复后处理
            Picture_Pre(docx_save_path, False, "模型后处理.jpg", 'post')

            # 围檩及内支撑参数(从模型读取)
            # 围檩参数
            # 围檩参数
            WL_group_namelst, Waler_Strut_Zlst, Waler_SECT_lst, Waler_MATL_lst = Get_Waler_info_from_Model(Matl_res, SECT_res, Node_res, Elem_res, Structure_group_res, Structure_group_namelst)
            num_layers = len(Waler_SECT_lst)
            has_waler = len(WL_group_namelst) > 0
            # 内支撑参数
            DC_group_namelst, XC_group_namelst, DC_Section_dict, XC_Section_dict, DC_L_dict, XC_L_dict, DC_Material_dict, XC_Material_dict = Get_Strut_info_from_Model(Matl_res, SECT_res, Node_res, Elem_res, Structure_group_res, Structure_group_namelst)
            DC_SECT_lst = [DC_Section_dict.get(next((k for k in DC_Section_dict if f"第{i+1}层" in k), None), '/') for i in range(num_layers)]
            XC_SECT_lst = [XC_Section_dict.get(next((k for k in XC_Section_dict if f"第{i+1}层" in k), None), '/') for i in range(num_layers)]
            DC_MATL_lst = [DC_Material_dict.get(next((k for k in DC_Material_dict if f"第{i+1}层" in k), None), '/') for i in range(num_layers)]
            XC_MATL_lst = [XC_Material_dict.get(next((k for k in XC_Material_dict if f"第{i+1}层" in k), None), '/') for i in range(num_layers)]
            print('第i层对撑材质的/处理:', DC_MATL_lst)
            print('第i层斜撑材质的/处理:', XC_MATL_lst)
            print('第i层对撑截面的/处理:', DC_SECT_lst)
            print('第i层斜撑截面的/处理:', XC_SECT_lst)
            # 支撑层数：以围檩层数为准（每层围檩对应一层内支撑，且嵌固稳定性支点取围檩标高）
            has_strut = (len(XC_group_namelst) + len(DC_group_namelst)) > 0
            count = len(WL_group_namelst) if has_waler else 0
            print(f"有无围檩  : {has_waler}")
            print(f"有无内支撑: {has_strut}")
            print(f"支撑层数: {count}")
            print('')

            # 工况参数
            cases_info, load_cases = case_info(has_stage, str(Drawdown_height), Waler_SECT_lst , DC_SECT_lst, XC_SECT_lst, conc_type, conc_grade, conc_t)
            print(f'cases_info: {cases_info}')
            print(f'load_cases: {load_cases}')
            # 荷载组合参数
            if not has_stage: # 仅整体建模读荷载组合参数
                comp = comp_info()
            else:
                comp = None
            print('comp:', comp)
            print("Success: 模型结果 读取成功")
            print('')

            # ====== 土压力计算 ===================================================================================================================================================================================
            # ====== 土压力计算 ===================================================================================================================================================================================
            
            update_p("正在进行土压力计算")
            excavate_stages = []
            if has_stage == True: # 当模型中存在施工阶段时
                # 整体模型
                combo_dict = {}
            else: # 当模型中不存在施工阶段时
                # 施工阶段定义
                stage_dict = {}
                # 整体模型
                combo_dict = {
                    '超挖深度'  : Drawdown_height,
                    '基坑底降水': Waterdown_height,
                    '开挖面降水': Excavation_face_dewater,
                }
            if stage_dict:  # 存在施工阶段
                # 筛选取土/抽水工况
                target_types = ['取土' ]if Solid_Level >= Water_Level else ['取土', '抽水']
                print('取土工况信息:')
                for stage_num, stage_info in stage_dict.items():
                    # 筛选取土工况
                    if stage_info['工况类型'] in target_types:
                        # 收集对应施工阶段的土顶标高和水面标高
                        excavate_stages.append({
                            '工况序号': stage_num,
                            '工况类型': stage_info['工况类型'],
                            '基坑内土顶标高': float(stage_info['基坑内土顶标高']),
                            '基坑内水面标高': float(stage_info['基坑内水面标高'])
                        })
                        print(f"  工况{stage_num}({stage_info['工况类型']}): 土顶标高={stage_info['基坑内土顶标高']}m, 水面标高={stage_info['基坑内水面标高']}m")
                
            # 主动土压力——当前仅考虑均布附加荷载
            # 调用 Active_Earth_Pressure 和 Passive_Earth_Pressure 函数获取主动土压力和被动土压力的计算结果，包括标高、土层状态、土层信息、土体竖向有效应力、水体竖向有效应力、土压力、水压力
            Active_Earth_Pressure_Hlst, Active_Earth_Pressure_Solid_State, Active_Earth_Pressure_Solid_Info, Active_Earth_Pressure_Solid_Sigma_lst, Active_Earth_Pressure_Water_Sigma_lst, Active_Earth_Pressure_Solid_Pak_lst, Active_Earth_Pressure_Water_Pak_lst = Active_Earth_Pressure(Solid_data_dict, Solid_Level, Water_Level, CofferDam_Bottom_Level, Ea_p0)
            print("Success: 主动土压力计算完成")
            print('')
            # 被动土压力——当前不考虑土的应力路径
            Passive_Earth_Pressure_Hlst, Passive_Earth_Pressure_Solid_State, Passive_Earth_Pressure_Solid_Info, Passive_Earth_Pressure_Solid_Sigma_lst, Passive_Earth_Pressure_Water_Sigma_lst, Passive_Earth_Pressure_Solid_Ppk_lst, Passive_Earth_Pressure_Water_Ppk_lst  = Passive_Earth_Pressure(Solid_data_dict, Solid_Level, Water_Level, CofferDam_Bottom_Level, Ep_p0, Concrete_Bottom_Level, has_stage, excavate_stages, combo_dict)
            print("Success: 被动土压力计算完成")
            print('')

            # ====== 模型读取 ===================================================================================================================================================================================
            # ====== 模型读取 ===================================================================================================================================================================================
                    
            # 模型计算部分计算
            update_p("正在读取模型结果")
            # 根据支护桩、围檩、对撑、斜撑的材质及截面信息计算容许应力
            allowable_f_dict = allowable_f(
                SECtype, pile_material, ggz_t, steel_standard, Steel_dict,
                Waler_section_dict, WL_group_namelst, Waler_SECT_lst, Waler_MATL_lst,
                Strut_section_dict, DC_SECT_lst, XC_SECT_lst, DC_MATL_lst, XC_MATL_lst,
                )

            pile_res, Waler_res, Strut_res, model_cal_satisified, _batch_results = Model_cal_res(
                version, has_stage, load_cases, has_waler, has_strut, Strut_section_dict,
                gbz_name, allowable_f_dict, Waler_SECT_lst, WL_group_namelst,  
                XC_group_namelst,DC_group_namelst, DC_Section_dict, XC_Section_dict, DC_L_dict, XC_L_dict, 
                gamma_0, SECtype, CofferDam_L, pbar=pbar, svar=svar, root=root
                )
            print('')
            print('='*20 + f'{proj_key}' + '='*20)
            print(pile_res)
            print('='*20 + f'{proj_key}' + '='*20)
            print(Waler_res)
            print('='*20 + f'{proj_key}' + '='*20)
            print(Strut_res)
            print('='*20 + f'{proj_key}' + '='*20)
            print(model_cal_satisified)
            print('='*20 + f'{proj_key}' + '='*20)
            print(_batch_results)
            print('')
            print('Success: 模型结果读取成功')

            # ====== 稳定性计算 ===================================================================================================================================================================================
            # ====== 稳定性计算 ===================================================================================================================================================================================

            # 稳定性计算
            update_p("正在进行稳定性分析")
            stability_result = stability_cal(Excel_Data, Solid_Level, Water_Level, Method_for_Pressures_var, Active_Earth_Pressure_Hlst, Active_Earth_Pressure_Solid_Pak_lst, Active_Earth_Pressure_Water_Pak_lst, Passive_Earth_Pressure_Hlst, Passive_Earth_Pressure_Solid_Ppk_lst, Passive_Earth_Pressure_Water_Ppk_lst, Waler_Strut_Zlst,  Concrete_Bottom_Level, CofferDam_Bottom_Level, CofferDam_Level, slice_width, solid_layer_lst, Ld, Foundation_Pit_Depth, Ea_p0, count, cal_config=cal_config)
            print('Success: 稳定性计算成功')
        
            # 清除运行过程中生成的中间图片（仅删除本次新增的，不影响已有文件）
            _current_imgs = set(glob.glob(os.path.join(docx_save_path, '*.jpg')) + glob.glob(os.path.join(docx_save_path, '*.png')))
            for img_path in _current_imgs - _pre_existing_imgs:
                try:
                    os.remove(img_path)
                except Exception as e:
                    print(f"删除文件 {img_path} 时出错: {e}")
            print("中间文件已清理")

            result_i = {
                'standard':[],
                'pile':{
                    'cb':[], 
                    'd' :[],
                }, 
                'waler':{
                    '1':{'cb':[], 'ssz':[], 'd': []}, 
                    '2':{'cb':[], 'ssz':[], 'd': []}, 
                    '3':{'cb':[], 'ssz':[], 'd': []}, 
                    '4':{'cb':[], 'ssz':[], 'd': []}, 
                    '5':{'cb':[], 'ssz':[], 'd': []}, 
                }, 
                'strut':{
                    '1':{'sigmax':[], 'sigmay':[]},
                    '2':{'sigmax':[], 'sigmay':[]},
                    '3':{'sigmax':[], 'sigmay':[]},
                    '4':{'sigmax':[], 'sigmay':[]},
                    '5':{'sigmax':[], 'sigmay':[]},
                }, 
                'stability':{
                    'overall'        :[],
                    'anti_overturn'  :[],
                    'heave_resistant':[],
                    'seepage'        :[],
                },
            }

            # 计算依据
            result_i['standard'].append('《建筑基坑支护技术规程》JGJ120-2012')
            # 支护桩组合应力与变形
            result_i['pile']['cb'] = [round(float(_batch_results['pile_max_cb']), acc), round(float(allowable_f_dict['pile']['f']), acc)]
            result_i['pile']['d'] = [round(float(_batch_results['pile_max_d']), acc), round(float(CofferDam_L)*2.5, 1)]
            # 围檩组合应力、剪应力与变形
            waler_num_lst = list(_batch_results['wl_summary'].keys())
            print(waler_num_lst)
            waler_max_cb_lst = [v['max_cb'] for v in _batch_results['wl_summary'].values()]
            waler_max_shear_lst = [v['max_shear'] for v in _batch_results['wl_summary'].values()]
            waler_max_d_lst = [v['max_d'] for v in _batch_results['wl_summary'].values()]
            waler_allowable_f_lst = [v['f'] for v in allowable_f_dict['waler']]
            waler_allowable_fv_lst = [v['fv'] for v in allowable_f_dict['waler']]
            for i in range(5):
                if i <= len(waler_num_lst)-1:
                    result_i['waler'][str(i+1)]['cb'] = [round(float(waler_max_cb_lst[i]), acc), round(float(waler_allowable_f_lst[i]), acc)]
                    result_i['waler'][str(i+1)]['ssz'] = [round(float(waler_max_shear_lst[i]), acc), round(float(waler_allowable_fv_lst[i]), acc)]
                    result_i['waler'][str(i+1)]['d'] = [round(float(waler_max_d_lst[i]), acc), '/']
                else:
                    result_i['waler'][str(i+1)]['cb'] = ['/', '/']
                    result_i['waler'][str(i+1)]['ssz'] = ['/', '/']
                    result_i['waler'][str(i+1)]['d'] = ['/', '/']
            # 内支撑稳定应力
            # 对撑
            strut_dc_name = list(_batch_results['DC_sigma_dict'].keys())
            strut_dc_max_cb_dict = {}
            for i in range(5):
                if f'第{i+1}层对撑' in strut_dc_name:
                    strut_dc_max_cb_dict[i+1] = sorted(_batch_results['DC_sigma_dict'][f'第{i+1}层对撑'], key = lambda x : abs(x))[-1]
                else:
                    strut_dc_max_cb_dict[i+1] = 0
            # 斜撑
            strut_xc_name = list(_batch_results['XC_sigma_dict'].keys())
            strut_xc_max_cb_dict = {}
            for i in range(5):
                if f'第{i+1}层斜撑' in strut_xc_name:
                    strut_xc_max_cb_dict[i+1] = sorted(_batch_results['XC_sigma_dict'][f'第{i+1}层斜撑'], key = lambda x : abs(x))[-1]
                else:
                    strut_xc_max_cb_dict[i+1] = 0
            print('对撑', strut_dc_name)
            print('斜撑' ,strut_xc_name)
            print('对撑结果', strut_dc_max_cb_dict)
            print('斜撑结果', strut_xc_max_cb_dict)
            for i in range(5):
                dc_max_cb = strut_dc_max_cb_dict[i+1]
                xc_max_cb = strut_xc_max_cb_dict[i+1]
                dc_allowable_f = allowable_f_dict['strut_dc'].get(i+1, {}).get('f', 215.0)
                xc_allowable_f = allowable_f_dict['strut_xc'].get(i+1, {}).get('f', 215.0)
                print(f'第{i+1}层内支撑')
                print('最大对撑', dc_max_cb, '最大斜撑', xc_max_cb)
                if dc_max_cb != 0:
                    result_i['strut'][str(i+1)]['sigmax'] = [round(float(dc_max_cb), acc), round(float(dc_allowable_f), acc)]
                else:
                    result_i['strut'][str(i+1)]['sigmax'] = ['/', '/']
                if xc_max_cb != 0:
                    result_i['strut'][str(i+1)]['sigmay'] = [round(float(xc_max_cb), acc), round(float(xc_allowable_f), acc)]
                else:
                    result_i['strut'][str(i+1)]['sigmay'] = ['/', '/']

            # 稳定性
            Ks, req_Ks = stability_result['圆弧滑动稳定性']
            Ke, req_Ke = stability_result['抗倾覆']
            Kb, req_Kb = stability_result['抗隆起']
            
            result_i['stability']['overall'] = [round(float(Ks), acc) if Ks != '/' else Ks, round(float(req_Ks), acc) if req_Ks != '/' else req_Ks]
            result_i['stability']['anti_overturn'] = [round(float(Ke), acc) if Ke != '/' else Ke, round(float(req_Ke), acc) if req_Ke != '/' else req_Ke]
            result_i['stability']['heave_resistant'] = [round(float(Kb), acc) if Kb != '/' else Kb, round(float(req_Kb), acc) if req_Kb != '/' else req_Kb]
            result_i['stability']['seepage'] = ['/', '/']


            cofferdam_param_table_proj_result_lst.append(
                {'project_name':proj, 'cofferdam_index':index, 'result':result_i}
            )

            MidasAPI("POST" , "/doc/save" , {})

            # ---- 立即将本条结果写入 Excel 并保存 ----
            _written = False
            for _r in range(_data_start_row, _ws.max_row + 1):
                _rname = str(_ws.cell(_r, _hm.get('项目名称', 0)).value or '').strip()
                _ridx  = str(_ws.cell(_r, _hm.get('围堰编号', 0)).value or '').strip()
                if _rname == proj and _ridx == index:
                    _write_result(_ws, _r, result_i, _hm)
                    _wb.save(excel_path)
                    print(f'已写入 Excel: {proj}-{index} -> R{_r}')
                    _written = True
                    break
            if not _written:
                print(f'警告: Excel 中未找到 {proj}-{index} 对应行，结果未写入')

        except MidasConnectionError as e:
            # API 连接断开，保存已有结果后跳出循环
            print(f"\n[MidasConnectionError] API 连接断开: {e}")
            print("正在保存已有的 Excel 结果并退出循环...")
            _wb.save(excel_path)
            _api_disconnected = True
            break
        except Exception as e:
            print(f"计算 {mcb} 时出错: {e}")
            cofferdam_param_table_proj_result_lst.append({})

    # 关闭工作簿
    _wb.close()

    if _api_disconnected:
        print(f"\n警告: Midas API 在处理过程中断开，已完成 {len([x for x in cofferdam_param_table_proj_result_lst if x])} 个模型，结果已保存。")
    elif pbar and svar and root:
        pbar['value'] = 100 # 强制进度条拉满
        svar.set("计算完成！结果已保存。")
        root.update() # 立即刷新 UI

    end_time = time.time() # 记录结束时间

    print('')
    for lst in cofferdam_param_table_proj_result_lst:
        print(lst)
        print('')

    print(f"总耗时: {end_time - start_time:.2f} 秒")

