# 1. 标准库
import os
import re
import time
import glob

# 2. 第三方库
from docx.shared import Mm
from docxtpl.subdoc import SubdocComposer
from docxtpl import DocxTemplate, InlineImage

# 3. 本地模块
from General.Midas import get_midas_version, MidasAPI, API_DIST_FORCE_Unit, Picture_Pre

from CalRpt_MSWord.Steel_Sheet_Pile_CofferDam_Cal.CofferDam_Cal_File_Manipulation import load_cofferdam_params_for_cal, load_waler_strut_sections, get_excel_path, load_material_library
from CalRpt_MSWord.Steel_Sheet_Pile_CofferDam_Cal.CofferDam_Midas_Model_Post      import Get_Waler_info_from_Model, Get_Strut_info_from_Model
from CalRpt_MSWord.Steel_Sheet_Pile_CofferDam_Cal.CofferDam_Cal_Docx_Handler      import (
    Active_Earth_Pressure, Passive_Earth_Pressure, case_info, comp_info, CofferDam_Info, Di_Zhi_Can_Shu, 
    Ji_Shu_Can_Shu, soild_pressure_cal, m_method, Model_cal_res, stability_cal, Draw_Solid_Pressure, allowable_f
)


# ── 修复子文档合并时编号重启问题 ──────────────────────────────
# restart_first_numbering() 会强制给段落加 <w:ilvl w:val="0"/>
# 导致图/表标题的多级编号跳回最高级，并全文传播。
# 覆写为 no-op 即可，避免影响父类 Composer。
SubdocComposer.restart_first_numbering = lambda *args: None

#=====================================================================================================================================================================

# 总函数, 生成所有word(进度条版)
def make_all_doc(Applocation, openfile_path, saveas_path, docx_save_path = "", pbar=None, svar=None, root=None, all_res_var=None, seepage_params=None, proj_key=None, cal_config=None):
    # 记录开始时间
    start_time = time.time()
    # 定义任务清单
    total_steps = 9
    current_step = 0
    def update_p(desc):
        nonlocal current_step
        current_step += 1
        if pbar and svar and root:
            pbar['value'] = (current_step / total_steps) * 100
            svar.set(desc)
            root.update()

    # 指定计算模型，则图片也会存到模型所在的文件夹内
    Midas_path = os.path.dirname(openfile_path) # 文件目录路径
    # 记录流程开始前目录中已有的图片文件，用于后续精确清理中间图片
    _pre_existing_imgs = set(glob.glob(os.path.join(docx_save_path, '*.jpg')) + glob.glob(os.path.join(docx_save_path, '*.png'))) if docx_save_path else set()
    Midas_name = os.path.basename(openfile_path) # 文件名
    # 计算书保存路径
    if docx_save_path == "":
        docx_save_path = Midas_path
    print(docx_save_path)
    # 计算书打开路径
    docx_open_path = os.path.join(Applocation, 'Support\\templates')

    # ====== 连接excel参数 ===================================================================================================================================================================================
    # ====== 连接excel参数 ===================================================================================================================================================================================
    
    update_p("正在匹配围堰数据库")

    # 从 Excel 读取围堰参数（替代原 txt 读取）
    excel_path = get_excel_path()
    if not excel_path:
        raise FileNotFoundError("找不到 CofferDam_ParamTable.xlsx，请确认程序安装路径")
    if not proj_key:
        raise ValueError("未指定项目，请在界面中选择项目名称和围堰编号")
    params = load_cofferdam_params_for_cal(excel_path, proj_key)

    _, _, Waler_section_dict, Strut_section_dict = load_waler_strut_sections(excel_path)
    Steel_dict, Concrete_dict, Rebar_dict = load_material_library(excel_path)
    for k, v in params.items():
        print(k, v)

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
    #     # 对于需要运行的模型，没有结果文件，则先运行，有则不运行
    #     Midas_Analysis(Midas_name, Midas_path)
    # 根据文件名生成结果文件名
    Analysis_result_outfile_name = Midas_name.replace(".mcb", ".OUT")
    Analysis_result_ca1file_name = Midas_name.replace(".mcb", ".CA1")
    print(Analysis_result_outfile_name)
    print(Analysis_result_ca1file_name)
    # 生成对应的路径
    Analysis_result_file_path1 = os.path.join(Midas_path, Analysis_result_outfile_name)
    Analysis_result_file_path2 = os.path.join(Midas_path, Analysis_result_ca1file_name)
    print(Analysis_result_file_path1)
    print(Analysis_result_file_path2)
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
    WL_group_namelst, Waler_Strut_Zlst, Waler_SECT_lst, Waler_MATL_lst = Get_Waler_info_from_Model(Matl_res, SECT_res, Node_res, Elem_res, Structure_group_res, Structure_group_namelst)
    num_layers = len(Waler_SECT_lst)
    has_waler = len(WL_group_namelst) > 0
    # 内支撑参数
    DC_group_namelst, XC_group_namelst, DC_Section_dict, XC_Section_dict, DC_L_dict, XC_L_dict, DC_Material_dict, XC_Material_dict = Get_Strut_info_from_Model(Matl_res, SECT_res, Node_res, Elem_res, Structure_group_res, Structure_group_namelst)
    # 层数：取围檩、对撑、斜撑结构组的最大层号
    DC_SECT_lst = [DC_Section_dict.get(next((k for k in DC_Section_dict if f"第{i+1}层" in k), None), '/') for i in range(num_layers)]
    XC_SECT_lst = [XC_Section_dict.get(next((k for k in XC_Section_dict if f"第{i+1}层" in k), None), '/') for i in range(num_layers)]
    DC_MATL_lst = [DC_Material_dict.get(next((k for k in DC_Material_dict if f"第{i+1}层" in k), None), '/') for i in range(num_layers)]
    XC_MATL_lst = [XC_Material_dict.get(next((k for k in XC_Material_dict if f"第{i+1}层" in k), None), '/') for i in range(num_layers)]
    print('第i层对撑材质的/处理:', DC_MATL_lst)
    print('第i层斜撑材质的/处理:', XC_MATL_lst)
    print('第i层对撑截面的/处理:', DC_SECT_lst)
    print('第i层斜撑截面的/处理:', XC_SECT_lst)
    # 支撑层数
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

    # ====== 计算书渲染 ===================================================================================================================================================================================
    # ====== 计算书渲染 ===================================================================================================================================================================================
        
    # try:

    # 打开主模板
    main_doc = DocxTemplate(os.path.join(docx_open_path, 'Steel_Sheet_Pile_CofferDam_cal_template_main.docx'))
    # 获取主模板信息
    update_p("正在生成主模型信息")
    # 围堰信息
    info = CofferDam_Info(SECtype, Cap_X, Cap_Y, Cap_H, Cap_Bottom_Level, CofferDam_a, CofferDam_b, CofferDam_L, CofferDam_Top_Level, CofferDam_Bottom_Level, Water_Level, Concrete_Bottom_Level, gbz_name, conc_type, conc_grade, conc_t, Waler_SECT_lst, Waler_MATL_lst, DC_MATL_lst, XC_MATL_lst, DC_SECT_lst, XC_SECT_lst, DC_group_namelst, XC_group_namelst, Ea_p0, SECmaterial=pile_material)
    # 土层参数
    soild_layers = Di_Zhi_Can_Shu(Excel_Data, Solid_Level)
    # 技术参数
    jishucanshu = Ji_Shu_Can_Shu(Waler_section_dict, Strut_section_dict, Waler_SECT_lst, DC_SECT_lst, XC_SECT_lst, WL_group_namelst, XC_group_namelst, DC_group_namelst, gbz_name, ggz_name, ggz_A, ggz_I, ggz_W)
    print('Success: 主模型信息 成功生成')
    print('')

    # 渲染子计算书
    #     try:
    # 土压力计算书
    update_p("正在生成土压力计算模板信息")
    soild_pressure_cal(docx_open_path, docx_save_path, Method_for_Pressures_var, 
                       Solid_Level, Water_Level, CofferDam_Bottom_Level, Concrete_Bottom_Level, Ea_p0,
                       Active_Earth_Pressure_Hlst, Active_Earth_Pressure_Solid_Info, Active_Earth_Pressure_Solid_State, 
                       Active_Earth_Pressure_Solid_Sigma_lst, Active_Earth_Pressure_Solid_Pak_lst, Active_Earth_Pressure_Water_Sigma_lst, Active_Earth_Pressure_Water_Pak_lst, 
                       Passive_Earth_Pressure_Hlst, Passive_Earth_Pressure_Solid_Info, Passive_Earth_Pressure_Solid_State, 
                       Passive_Earth_Pressure_Solid_Sigma_lst, Passive_Earth_Pressure_Solid_Ppk_lst, Passive_Earth_Pressure_Water_Sigma_lst, Passive_Earth_Pressure_Water_Ppk_lst, 
                       excavate_stages, combo_dict,
                       )
    soild_pressure_doc = main_doc.new_subdoc(os.path.join(docx_save_path, 'solid_pressure_final.docx'))
    print('Success: 土压力计算模板信息 成功生成')
    print('')
    #     except Exception as e:
    #         print('Fail: 土压力计算模板信息 生成失败', e)

    #     try:
    # m法计算书
    update_p("正在生成m法表格模板信息")
    Condition_m_method_Hlst = m_method(docx_open_path, docx_save_path, Passive_Earth_Pressure_Hlst, Passive_Earth_Pressure_Solid_Info, Basic_Solid_Layer_Thickness, pile_b)
    m_method_doc = main_doc.new_subdoc(os.path.join(docx_save_path,'m_method.docx'))
    print('Success: m法表格模板信息 成功生成')
    print('')
    #     except Exception as e:
    #         print('Fail: m法表格模板信息 生成失败', e)

    #     try:
    # 模型计算部分计算书
    update_p("正在进行模型结果验算")
    allowable_f_dict = allowable_f(SECtype, pile_material, ggz_t, steel_standard, Steel_dict, 
                Waler_section_dict, WL_group_namelst, Waler_SECT_lst, Waler_MATL_lst, 
                Strut_section_dict, DC_SECT_lst, XC_SECT_lst, DC_MATL_lst, XC_MATL_lst, 
                )
    
    pile_res, Waler_res, Strut_res, model_cal_satisified, _batch_results = Model_cal_res(
        version, docx_open_path, docx_save_path, docx_save_path, 
        has_stage, load_cases, has_waler, has_strut, Strut_section_dict,
        allowable_f_dict, CofferDam_L, gbz_name, Waler_SECT_lst, WL_group_namelst, 
        XC_group_namelst,DC_group_namelst, DC_Section_dict, XC_Section_dict, DC_L_dict, XC_L_dict, 
        gamma_0, SECtype, pbar=pbar, svar=svar, root=root
        )
    print('Success: 模型计算 替换成功')
    model_cal_doc = main_doc.new_subdoc(os.path.join(docx_save_path,'model_cal.docx'))
    #     except Exception as e:
    #         print('Fail: 模型计算 替换失败', e)

    #     try:
    # 稳定性计算书
    update_p("正在进行稳定性分析")
    overall_selected, overall_stability_res, anti_overturn_selected, anti_overturn_stability_res, heave_resistant_selected, cal_heave_resistant_stability, heave_resistant_stability_res, stability_satisfied, cal_seepage_stability, seepage_stability_res = stability_cal(docx_open_path, docx_save_path, Excel_Data, Solid_Level, Water_Level, Method_for_Pressures_var, Active_Earth_Pressure_Hlst, Active_Earth_Pressure_Solid_Pak_lst, Active_Earth_Pressure_Water_Pak_lst, Passive_Earth_Pressure_Hlst, Passive_Earth_Pressure_Solid_Ppk_lst, Passive_Earth_Pressure_Water_Ppk_lst, Waler_Strut_Zlst,  Concrete_Bottom_Level, CofferDam_Bottom_Level, CofferDam_Level, slice_width, solid_layer_lst, Ld, Foundation_Pit_Depth, Ea_p0, count, cal_config=cal_config)

    print('Success: 稳定性计算 替换成功')
    stability_cal_doc = main_doc.new_subdoc(os.path.join(docx_save_path,'stability_cal_final.docx'))
    #     except Exception as e:
    #         print('Fail: 稳定性计算 替换失败', e)
    
    # 结论处理
    satisfied = "满足要求" if model_cal_satisified == "满足要求" and stability_satisfied == "满足要求" else "不满足要求"
    if all_res_var:
        # 格式化排版：使用双换行符分隔不同构件的结论
        report_parts = []
        report_parts.append(f"\n【结论】该{SECtype}围堰计算{satisfied}。")
        report_parts.append(f"\n【{SECtype}验算】\n{pile_res}")
        if has_waler:
            Waler_res_text = "\n".join(Waler_res)
            report_parts.append(f"\n【围檩验算】\n{Waler_res_text}")
        if has_strut:
            Strut_res_text = "\n".join(Strut_res)
            report_parts.append(f"\n【内支撑验算】\n{Strut_res_text}")
        if overall_selected:
            report_parts.append(f"\n【整体稳定性验算】\n{overall_stability_res}")
        if anti_overturn_selected:
            report_parts.append(f"\n【抗倾覆稳定性验算】\n{anti_overturn_stability_res}")
        if heave_resistant_selected and cal_heave_resistant_stability:
            report_parts.append(f"\n【抗隆起稳定性验算】\n{heave_resistant_stability_res}")
        if cal_seepage_stability:
            seepage_res_text = "\n".join(seepage_stability_res)
            report_parts.append(f"{seepage_res_text}")
        # 将所有结论合并成一个字符串
        final_report_text = "\n".join(report_parts)
        # 生成简易计算结果时间标记
        local_t = time.localtime(start_time)
        formatted_start_time = time.strftime("%y.%#m.%#d %H:%M", local_t)
        final_report_text += f"|TIME_STAMP|{formatted_start_time}"
        all_res_var.set(final_report_text)
        # 成功后更新状态文字
        svar.set("计算完成，正在渲染主模板")
        root.update() # 立即刷新 UI

    # 土压力图示
    solid_pressure_pic_path = os.path.join(docx_save_path, '土压力图示.png')
    Draw_Solid_Pressure(
                    solid_pressure_pic_path, Solid_Level, Water_Level, CofferDam_Top_Level, CofferDam_L, Concrete_Bottom_Level, Method_for_Pressures_var, Waler_Strut_Zlst, Condition_m_method_Hlst, 
                    Active_Earth_Pressure_Hlst, Active_Earth_Pressure_Solid_Pak_lst, Active_Earth_Pressure_Water_Pak_lst, 
                    Passive_Earth_Pressure_Hlst, Passive_Earth_Pressure_Solid_Ppk_lst, Passive_Earth_Pressure_Water_Ppk_lst
                )
    
    #     try:
    context = {
        'info': info,
        'tucengcanshu':soild_layers,
        'jishucanshu':jishucanshu,
        'has_stage':has_stage,
        'cases_info':cases_info,
        'comp':comp,
        'model_pre_picture1': InlineImage(main_doc, os.path.join(docx_save_path, '模型前处理.jpg'),  width=Mm(100)),
        'Solid_Water_Pressure_Drawing':InlineImage(main_doc,solid_pressure_pic_path, width=Mm(100)),
        'Solid_Pressure':soild_pressure_doc,
        'm_method':m_method_doc,
        'model_cal':model_cal_doc,
        'stability_cal':stability_cal_doc,
        'has_waler':has_waler,
        'has_strut':has_strut,
        'pile_res':pile_res,
        'Waler_res':Waler_res,
        'Strut_res':Strut_res,
        'overall_selected':overall_selected,
        'anti_overturn_selected':anti_overturn_selected,
        'heave_resistant_selected':heave_resistant_selected,
        'overall_stability_res':overall_stability_res,
        'anti_overturn_stability_res':anti_overturn_stability_res,
        'cal_heave_resistant_stability':cal_heave_resistant_stability,
        'heave_resistant_stability_res':heave_resistant_stability_res,
        'cal_seepage_stability':cal_seepage_stability,
        'seepage_stability_res':seepage_stability_res,
        'satisfied':satisfied,
        'SECtype':SECtype,
    }
    # 渲染主模板
    main_doc.render(context)
    # 保存主模板
    update_p("正在保存并清理")
    main_doc.save(saveas_path)
    print('Success: 主模板替换成功')
    # 清除计算文件
    intermediate_files = [
        os.path.join(docx_save_path, 'solid_pressure.docx'),
        os.path.join(docx_save_path, 'solid_pressure_final.docx'),
        os.path.join(docx_save_path, 'm_method.docx'),
        os.path.join(docx_save_path,'model_cal.docx'),
        os.path.join(docx_save_path, 'Stability_cal.docx'),
        os.path.join(docx_save_path, 'Stability_cal_final.docx'), 
    ]
    for file_path in intermediate_files:
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
        except Exception as e:
            print(f"删除文件 {file_path} 时出错: {e}")
    # 清除运行过程中生成的中间图片（仅删除本次新增的，不影响已有文件）
    _current_imgs = set(glob.glob(os.path.join(docx_save_path, '*.jpg')) + glob.glob(os.path.join(docx_save_path, '*.png')))
    for img_path in _current_imgs - _pre_existing_imgs:
        try:
            os.remove(img_path)
        except Exception as e:
            print(f"删除文件 {img_path} 时出错: {e}")
    print("中间文件已清理")
    print(f'钢板桩围堰计算书已保存至{saveas_path}')
    if pbar and svar and root:
        pbar['value'] = 100 # 强制进度条拉满
        svar.set("计算完成！结果已保存。")
        root.update() # 立即刷新 UI
    #     except Exception as e:
    #         print('Fail: 主模板渲染失败', e)
    #         if pbar and svar and root:
    #             # 将进度条颜色改为红色 (DANGER)
    #             pbar.configure(bootstyle="danger") 
    #             pbar['value'] = 100 
    #             if 'Permission denied' in str(e):
    #                 svar.set("出错啦！请关闭Word后重试。")
    #             else:
    #             # 修改状态文字，提示用户查看原因
    #                 svar.set(f"出错啦！{str(e)}...") 
    #             # 立即刷新界面
    #             root.update()


    # except Exception as e:
    #     if svar:
    #         svar.set(f"运行出错：{str(e)}")
    #     print('Fail: 运行出错', e)

    end_time = time.time() # 记录结束时间
    print(f"总耗时: {end_time - start_time:.2f} 秒")
