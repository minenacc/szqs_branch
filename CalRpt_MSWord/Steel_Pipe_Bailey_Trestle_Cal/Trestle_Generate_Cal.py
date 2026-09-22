"""
栈桥计算书 - 主流程模块
包含：模型后处理调度、计算书子流程编排、主入口 make_all_doc
"""

# 1. 标准库
import os
import sys
from pathlib import Path

# 2. 第三方库
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm
from docx import Document
from docxtpl.subdoc import SubdocComposer

# 3. 本地模块
sys.path.append(str(Path(__file__).parent.parent.parent))
from General.WordHandle import replace_placeholder_with_formula
from General.Midas import Picture_Pre
from Trestle_File_Cal import Midas_Analysis
from Trestle_Post_Cal import get_json_dict
from Trestle_Handle_Cal import if_Boundary_Group_name_exist, Post_processing_Midas_Model, get_component_info, get_model_params, get_foundation_material, environment_params, get_conditions_info, vehicle_subdoc_render, conditions_table_subdoc_render, plate_cal_params, model_cal_params, foundation_subdoc_render


# ── 修复子文档合并时编号重启问题 ──────────────────────────────
# restart_first_numbering() 会强制给段落加 <w:ilvl w:val="0"/>
# 导致图/表标题的多级编号跳回最高级，并全文传播。
# 覆写为 no-op 即可，避免影响父类 Composer。
SubdocComposer.restart_first_numbering = lambda *args: None

#=====================================================================================================================================================================


# 总函数, 生成所有word
def make_all_doc(response, Applocation, openfile_path, model_post_result, foundation_type, foundation_standard, foundation_value_dict, saveas_path, excel_params=None, docx_save_path = ""):
    # 指定计算模型，则图片也会存到模型所在的文件夹内
    Midas_path = os.path.dirname(openfile_path) # 文件目录路径
    Midas_name = os.path.basename(openfile_path) # 文件名
    output_dir = os.path.dirname(saveas_path)  # 最终计算书所在目录，子文档临时文件也存此处
    # 计算书保存路径
    if docx_save_path == "":
        docx_save_path = Midas_path
    # 模板打开路径
    docx_open_path = os.path.join(Applocation, 'Support\\templates')
    # 没有桩底反力边界组则不生成
    if not if_Boundary_Group_name_exist("桩底反力"):
        print("模型[组]中缺失[桩底反力]边界组, 请添加后重试")
    else:
        # 无论是否需要重新分析，都读取模型基本信息（供后续计算使用）
        print("开始读取模型基本信息...")
        NODE_dict_res, ELEM_dict_res, SECT_dict_res, THIK_dict_res, MATL_dict_res, GRUP_dict_res, COMP_dict_res = get_json_dict()
        print("模型基本信息读取完成")
        # 获取模型前处理图片
        Picture_Pre(Midas_path)

        # response=True 时运行分析并读取结果；否则使用已有的 model_post_result
        if response:
            print("开始运行模型...")
            Midas_Analysis(Midas_name, Midas_path)
            print("开始读取模型结果...")
            model_post_result[0] = Post_processing_Midas_Model(
                Midas_path, NODE_dict_res, ELEM_dict_res, SECT_dict_res, MATL_dict_res, GRUP_dict_res
            )
            print("模型结果读取完成")

        if model_post_result[0] == {}:
            print("未读取到模型结果, 无法生成计算书")
        else:
            # ── docxtpl 渲染 ──
            print("   ")
            print("开始生成计算书(docxtpl)...")
            # 0、创建模板
            tpl_path = os.path.join(docx_open_path, 'Trestle_cal_template_main.docx')
            tpl = DocxTemplate(tpl_path)

            # 1、读取各构件截面/材料信息
            component_info = get_component_info(
                GRUP_dict_res, NODE_dict_res, ELEM_dict_res,
                SECT_dict_res, MATL_dict_res, THIK_dict_res
            )
            print("   ")
            print("构件信息读取完成")
            print(f'    component_info: {component_info}')
            # 2、从模型获取计算参数
            steel_dict = excel_params.get('steel_dict', {})
            concrete_dict = excel_params.get('concrete_dict', {})
            rebar_dict = excel_params.get('rebar_dict', {})
            self_weight, beam_n, pile_d_lst, deck_type, material, midas_version = get_model_params(
                component_info, GRUP_dict_res, NODE_dict_res, MATL_dict_res,
                steel_dict=steel_dict, concrete_dict=concrete_dict
            )
            print("   ")
            print("计算参数读取完成")
            print(f'    self_weight: {self_weight}')
            print(f'    beam_n: {beam_n}')
            print(f'    pile_d_lst: {pile_d_lst}')
            print(f'    deck_type: {deck_type}')
            print(f'    material: {material}')
            print(f'    midas_version: {midas_version}')

            # 3、从 UI 获取基础材料参数，合并到 material
            foundation_material = get_foundation_material(
                foundation_type, foundation_value_dict, material,
                concrete_dict=concrete_dict, rebar_dict=rebar_dict
            )
            material.update(foundation_material)
            print("   ")
            print("基础材料参数读取完成")
            print(f'    material: {material}')
            
            # 获取纵梁类型和栈桥类型
            beam_type = component_info.get('beam_type', '贝雷梁')
            bridge_type = excel_params.get('beam_type', '上承式桁架梁栈桥')  # 从Excel读取栈桥类型
            
            # 4、从 Excel 获取环境参数
            wind, water, soil_param = environment_params(excel_params, beam_n, pile_d_lst, 
                                                          beam_type=beam_type, 
                                                          component_info=component_info,
                                                          bridge_type=bridge_type)
            print("   ")
            print("环境参数读取完成")
            print(f'    wind: {wind}')
            print(f'    water: {water}')
            print(f'    soil_param: {soil_param}')

            # 5、解析荷载工况
            conditions_table, conditions = get_conditions_info(COMP_dict_res)
            conditions_table_subdoc = conditions_table_subdoc_render(
                tpl, conditions_table, excel_params.get('vehicle_params_list', []), docx_open_path, output_dir
            )
            print("   ")
            print("荷载工况解析完成")
            print(f'    conditions_table: {conditions_table}')
            print(f'    conditions: {conditions}')  

            # 6、计算模板渲染所需的全部验算参数
            # 6.1、桥面板验算
            deck, deck_formula = plate_cal_params(
                component_info, excel_params, material,
                conditions, GRUP_dict_res, NODE_dict_res, deck_type
            )
            print("   ")
            print("桥面板验算完成")
            print(f'    deck: {deck}')
            print(f'    deck_formula: {deck_formula}')

            # 6.2、模型读取结果验算
            rib_long, rib_trans, dist, beam, pile, brace = model_cal_params(
                model_post_result[0], Midas_path, component_info,
                ELEM_dict_res, SECT_dict_res, MATL_dict_res,
                GRUP_dict_res, NODE_dict_res, deck_type, steel_dict=steel_dict
            )
            print("   ")
            print("模型读取结果验算完成")
            print(f'    rib_long: {rib_long}')
            print(f'    rib_trans: {rib_trans}')
            print(f'    dist: {dist}')
            print(f'    beam: {beam}')
            print(f'    pile: {pile}')
            print(f'    brace: {brace}')  

            # 7、渲染车辆设备参数子模板
            vehicle_subdoc = vehicle_subdoc_render(tpl, excel_params.get('vehicle_params_list', []), docx_open_path, output_dir)
            print("   ")
            print("车辆设备参数子模板渲染完成")  

            # 8、渲染基础验算子模板（需要土层参数）
            soil_param_index = excel_params.get('soil_param_index')
            if soil_param_index:
                foundation_subdoc, foundation_satisfied = foundation_subdoc_render(
                    foundation_type, foundation_standard, foundation_value_dict,
                    model_post_result[0],
                    excel_params, component_info,
                    docx_open_path, tpl,
                    pile_d_lst[0] if pile_d_lst else 0,
                    material,
                    output_dir,
                    Midas_path,
                    soil_param_index
                )
                print("基础验算子模板渲染完成")
            else:
                print("未找到地形土层参数，跳过基础验算")
                foundation_subdoc = ''
                foundation_satisfied = '满足'
            print("   ")  
            
            # 9、创建模板并将 dict 中的图片路径批量转为 InlineImage

            def _img(pic_path):
                """将图片完整路径转为 InlineImage；路径无效时返回空字符串"""
                if pic_path and os.path.exists(str(pic_path)):
                    return InlineImage(tpl, str(pic_path), width=Cm(16), height=Cm(10))
                return ''

            def _convert_pics(d):
                """将 dict 中所有 *_pic 键的路径字符串转为 InlineImage（原地替换）"""
                for key in list(d.keys()):
                    if key.endswith('_pic'):
                        d[key] = _img(d[key])

            for comp in (deck, rib_long, rib_trans, dist, beam, pile, brace):
                if isinstance(comp, dict):
                    _convert_pics(comp)

            # 10、组装 context
            context = {
                # ── 基础信息 ──
                'project_name':  excel_params.get('trestle_no', ''),
                'self_weight':   self_weight,
                'midas_version': midas_version,
                'deck_type':     deck_type,
                'beam_type':     component_info.get('beam_type', ''),
                'model_pic':     _img(os.path.join(Midas_path, '模型前处理.jpg')),
                'has_brace':     component_info.get('has_brace', False),
                'has_rib_long':  component_info.get('has_rib_long', False),

                # ── 环境参数 ──
                'wind':          wind,
                'water':         water,
                'soil_param':    soil_param,

                # ── 材料 ──
                'material':      material,

                # ── 车辆（子模板 subdoc 渲染结果） ──
                'vehicle_param': vehicle_subdoc,

                # ── 构件验算（dict 内含 InlineImage） ──

                'deck':          deck,
                'rib_long':      rib_long,
                'rib_trans':     rib_trans,
                'dist':          dist,
                'beam':          beam,
                'pile':          pile,
                'brace':         brace,

                # ── 荷载工况 ──
                'conditions':       conditions,
                'conditions_table':  conditions_table_subdoc,

                # ── 基础验算（子模板 subdoc 占位符） ──
                'foundation_cal':    foundation_subdoc,
                'foundation_standard': foundation_standard,
            }

            # ── 项目综合判定（全部验算汇总，任一不满足即为不满足） ──
            _all_satisfied = [
                deck.get('if_satisfied', ''),
                rib_long.get('if_satisfied', ''),
                rib_trans.get('if_satisfied', ''),
                dist.get('if_satisfied', ''),
                beam.get('chord_satisfied', ''),
                beam.get('vertical_satisfied', ''),
                beam.get('diagonal_satisfied', ''),
                pile.get('stability_satisfied', ''),
                brace.get('if_satisfied', ''),
                foundation_satisfied,
            ]
            if any('不满足' in s for s in _all_satisfied):
                project_satisfied = '不满足'
            elif all(s and '不满足' not in s for s in _all_satisfied):
                project_satisfied = '满足'
            else:
                project_satisfied = '满足'
            context['project_satisfied'] = project_satisfied

            # 11、渲染并保存
            tpl.render(context)
            tpl.save(saveas_path)
            print("   ")
            print(f"docxtpl渲染计算书已保存至: {saveas_path}")

            # 12、公式替换（docxtpl 渲染后，将 LaTeX 占位符转为 Word OMML 公式）
            formula_registry = {}
            # 桥面板公式
            if isinstance(deck_formula, dict):
                formula_registry.update(deck_formula)
            # 钢管桩稳定性公式
            if isinstance(pile, dict):
                formula_registry.update(pile.get('formula_registry', {}))
            # 风荷载公式
            formula_registry.update(wind.get('formula_registry', {}))
            # 水流力公式
            formula_registry.update(water.get('formula_registry', {}))

            if formula_registry:
                final_path = saveas_path
                final_doc = Document(saveas_path)
                for placeholder, latex in formula_registry.items():
                    if latex:
                        replace_placeholder_with_formula(final_doc, placeholder, latex)
                final_doc.save(final_path)
                print(f"公式替换完成: {final_path}")
            else:
                print("无需公式替换")
            print("   ")
            print(f"计算书已保存至: {saveas_path}")
            print("计算书模块已全部加载完成")

            # 13、清理中间文件
            # 清理临时子模板 docx
            for tmp_name in ('_vehicle_subdoc_tmp.docx', '_conditions_table_tmp.docx', '_foundation_subdoc_tmp.docx'):
                tmp_file = os.path.join(output_dir, tmp_name)
                if os.path.exists(tmp_file):
                    try:
                        os.remove(tmp_file)
                    except OSError:
                        pass
            # 清理 Midas 输出图片
            for f in os.listdir(Midas_path):
                if f.lower().endswith('.jpg'):
                    try:
                        os.remove(os.path.join(Midas_path, f))
                    except OSError:
                        pass
            print("中间文件清理完成")

