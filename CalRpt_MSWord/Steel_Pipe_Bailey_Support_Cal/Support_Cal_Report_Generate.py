# 1. 标准库
import os
import re 
import math

# 3. 本地模块
from General.Midas      import Picture_Pre
from General.TextHandle import read_txt_found_line, format_scientific_unicode
from General.FilePath   import file_extension_Modified
from General.WordHandle import WordApp_Dispatch, WordApp_Quit
from General.Formula    import (
    foundation_Gk, foundation_pk, cal_Eccentricity, foundation_pkmax_Large_Eccentricity, foundation_pkmax_pkmin, foundation_Vs, foundation_Allowable_Vs, foundation_Allowable_Fl, foundation_Fl, 
    foundation_pcz, foundation_pz, cal_M_param, cal_M, cal_As, min_p_As, cal_rebar_arrangement, JGJ94_Qsk_DrivenPile, JGJ94_Qpk_steelpile, JGJ94_Quk, JGJ94_Ra, JGJ94_Qpk_precastpile, rebar_f,
    Side_Resistance_Factor, Base_Resistance_Factor, JGJ94_Qsk_DrilledShaft, JGJ94_Qpk_DrilledShaft, Subgrade_Reaction_Coefficient, Displacement_Coefficient, JGJ94_Table_Vm_Vs, concrete_f,
    JGJ94_Rha1, JGJ94_Rha2, SteelPile_param, cal_p, cal_pmax, cal_pmin, Section_Modulus_of_Transformed_Section_at_Tension_Edge_of_Pile_Shaft, Transformed_Section_Area_of_Pile_Shaft
)

from CalRpt_MSWord.Steel_Pipe_Bailey_Support_Cal.Support_Cal_Standard_Formula  import *
from CalRpt_MSWord.Steel_Pipe_Bailey_Support_Cal.Support_Cal_File_Manipulation import Midas_Analysis
from CalRpt_MSWord.Steel_Pipe_Bailey_Support_Cal.Support_Midas_Model_Pre       import if_Boundary_Group_name_exist
from CalRpt_MSWord.Steel_Pipe_Bailey_Support_Cal.Support_Word_Processing       import (
    repl_docstr_with_omath, WordA_PasteTo_WordB_DeleteLines, Delete_Word_FindText_Paragraphs, repl_docstr_with_pic, repl_docstr_with_strlst, str_join_n, 
    FEA_To_Cal_txt_str_lookfor, Standard_Codename, open_doc_saveas_adoc, adoc_save_close
)
from CalRpt_MSWord.Steel_Pipe_Bailey_Support_Cal.Support_Midas_Model_Post      import (
    Post_processing_xiaolei, Post_processing_bailey, Post_processing_fenpeiliang, Post_processing_lianjiexi, Post_processing_gangguanzhuang, Post_processing_reaction, get_node_elem_sect_grup_dict
)

#=====================================================================================================================================================================

# 对Midas模型进行结果的读取并截图
def Post_processing_Midas_Model(folder_path, NODE_dict_res, ELEM_dict_res, SECT_dict_res, get_structure_group_res):
    print("正在获取模型计算结果...")
    # 获取模型前处理图片
    Picture_Pre(folder_path)
    # 后处理——小肋
    print("小肋信息读取中...")
    list1 = Post_processing_xiaolei(folder_path)
    # 后处理——贝雷梁
    print("贝雷梁信息读取中...")
    list2 = Post_processing_bailey(folder_path)
    # 后处理——分配梁
    print("分配梁信息读取中...")
    list3 = Post_processing_fenpeiliang(folder_path)
    # 后处理——联结系
    print("联结系信息读取中...")
    list4 = Post_processing_lianjiexi(folder_path)
    # 后处理——钢管桩
    print("钢管桩信息读取中...")
    list5 = Post_processing_gangguanzhuang(folder_path, NODE_dict_res, ELEM_dict_res, SECT_dict_res, get_structure_group_res)
    # 后处理——反力
    print("反力信息读取中...")
    list6 = Post_processing_reaction(folder_path)
    # 组成字典
    model_post_dict = {"model_post_xiaolei":list1,
                       "model_post_bailey":list2,
                       "model_post_fenpeiliang":list3,
                       "model_post_lianjiexi":list4,
                       "model_post_gangguanzhuang":list5,
                       "model_post_reaction":list6,
                    }
    print("模型信息读取完成")
    return model_post_dict

#=====================================================================================================================================================================

# 公式替换函数

# 扩大基础地基应力计算——轴心受压
def word_pk(word, aDoc, x, y, z, gc, h1, gs, Fk, fa, Ak, Gk, pk, if_satisfied_fa1):
    selection = word.Selection
    # latex_list1, latex_list2, latex_list3 =  foundation_pk_LaTeX(x, y, z, gc, h1, gs, Fk, fa, Ak, Gk, pk, if_satisfied_fa1) # LaTeX公式
    hs = max(round(abs(h1) - z, 1), 0) # h1是基础底标高, z是基础高度, 得到填土的高度, 当存在填土时, Gk公式需要加上填土的自重
    if hs != 0:
        # latex_list1 = f"G_k={x}×{y}×{z}×{gc}+{x}×{y}×{hs}×{gs}={Gk}kN" # UnicodeMath格式
        latex_list1 = 'G_k=' + str(x) + r'\times' + str(y) + r'\times' + str(z) + r'\times' + str(gc) + '+' + str(x) + r'\times' + str(y) + r'\times' + str(hs) + r'\times' + str(gs) + f'={Gk}kN' # LaTeX格式
    else:
        # latex_list1 = f"G_k={x}×{y}×{z}×{gc}={Gk}kN" # UnicodeMath格式
        latex_list1 = 'G_k=' + str(x) + r'\times' + str(y) + r'\times' + str(z) + r'\times' + str(gc) + f'={Gk}kN' # LaTeX格式
    # latex_list2 = f"p_k=(F_k+G_k)/A=({Fk}+{Gk})/{Ak}={pk}kPa" # UnicodeMath格式
    latex_list2 = GB50007_2011_5_2_2_1 + r'=\frac{\left(' + f'{Fk}+{Gk}' + r'\right)}{' + str(Ak) + r'}=' + f'{pk}kPa' # LaTeX格式
    if if_satisfied_fa1 == "":
        latex_list3 = []
    elif if_satisfied_fa1 == "满足":
        latex_list3 = f"p_k={pk}kPa<f_a={fa}kPa" # LaTeX格式
    else:
        latex_list3 = f"p_k={pk}kPa>f_a={fa}kPa" # LaTeX格式
    # 公式替换
    repl_docstr_with_omath(aDoc,selection,latex_list1,'{foundation_Gk}')
    repl_docstr_with_omath(aDoc,selection,latex_list2,'{foundation_pk}')
    repl_docstr_with_omath(aDoc,selection,latex_list3,'{if_satisfied_foundation_pk}')


# 扩大基础地基应力计算——偏心受压
def word_pkmax(word, aDoc, Fk, My, fa, Ak, Gk, e, a, l, b, Wx, pkmax, pkmin, if_satisfied_fa2):
    selection = word.Selection
    # LaTeX公式
    # latex_list1, latex_list2, latex_list3, latex_list4 = foundation_pkmax_LaTeX(Fk, My, fa, Ak, Gk, e, a, l, b, Wx, pkmax, pkmin, if_satisfied_fa2)
    if e >= b/6:
        # UnicodeMath格式
        # latex_list1 = f"e=M_k/(F_k+G_k)={My}/({Fk}+{Gk})={e} > {b}/6"
        # latex_list2 = f"p_kmax=(2(F_k+G_k))/3la=(2({Fk}+{Gk}))/(3×{l}×{a})={pkmax}kPa"
        # latex_list3 = f"p_kmin=0kPa"
        # LaTeX格式
        latex_list1 = GB50007_2011_Eccentricity + r'=\frac{' + str(My) + r'}{\left(' + f'{Fk}+{Gk}' + r'\right)}=' + str(e) + r'>\frac{' + str(b) + r'}{6}'
        latex_list2 = GB50007_2011_5_2_2_4 + r'=\frac{2\times\left(' + f'{Fk}+{Gk}' + r'\right)}{3\times' + str(l) + r'\times' + str(a) + r'}=' + str(pkmax) + 'kPa'
        latex_list3 = r'p_{kmin}=0kPa'
    else:
        # UnicodeMath格式
        # latex_list1 = f"e=M_k/(F_k+G_k)={My}/({Fk}+{Gk})={e} < {b}/6"
        # latex_list2 = f"p_kmax=(F_k+G_k)/A+M_k/W=({Fk}+{Gk})/{Ak}+{My}/{Wx}={pkmax}kPa"# Fk,Gk,Ak,My,Wx,pkmax
        # latex_list3 = f"p_kmin=(F_k+G_k)/A-M_k/W=({Fk}+{Gk})/{Ak}-{My}/{Wx}={pkmin}kPa"
        # LaTeX格式
        latex_list1 = GB50007_2011_Eccentricity + r'=\frac{' + str(My) + r'}{\left(' + f'{Fk}+{Gk}' + r'\right)}=' + str(e) + r'<\frac{' + str(b) + r'}{6}'
        latex_list2 = GB50007_2011_5_2_2_2 + r'=\frac{\left(' + f'{Fk}+{Gk}' + r'\right)}{' + str(Ak) + r'}+\frac{' + str(My) + '}{' + str(Wx) + '}=' + f'{pkmax}kPa'
        latex_list3 = GB50007_2011_5_2_2_3 + r'=\frac{\left(' + f'{Fk}+{Gk}' + r'\right)}{' + str(Ak) + r'}-\frac{' + str(My) + '}{' + str(Wx) + '}=' + f'{pkmin}kPa'
    if if_satisfied_fa2 == "":
        latex_list4 = ''
    elif if_satisfied_fa2 == "满足":
        # UnicodeMath格式
        # latex_list4 = f"p_kmax={pkmax}kPa<1.2f_a={1.2*fa}kPa"
        # LaTeX格式
        latex_list4 = r'p_{kmax}=' + f'{pkmax}kPa<1.2f_a={1.2*fa}kPa'
    else:
        # UnicodeMath格式
        # latex_list4 = [f"p_kmax={pkmax}kPa>1.2f_a={1.2*fa}kPa"]
        # LaTeX格式
        latex_list4 = r'p_{kmax}=' + f'{pkmax}kPa>1.2f_a={1.2*fa}kPa'
    # 公式替换
    repl_docstr_with_omath(aDoc,selection,latex_list1,'{foundation_Eccentricity}')
    repl_docstr_with_omath(aDoc,selection,latex_list2,'{foundation_pkmax}')
    repl_docstr_with_omath(aDoc,selection,latex_list3,'{foundation_pkmin}')
    repl_docstr_with_omath(aDoc,selection,latex_list4,'{if_satisfied_foundation_pkmax}')


# 扩大基础地基应力计算——软弱下卧层
def word_pz(word, aDoc, z, h1, h2, gs, faz, pcz, pc, pz, if_satisfied_faz):
    selection = word.Selection
    # LaTeX公式
    # latex_list1, latex_list2, latex_list3, latex_list4 = foundation_pz_LaTeX(z, h1, h2, gs, faz, pcz, pc, pz, if_satisfied_faz)
    # UnicodeMath格式
    # latex_list1 = f"p_cz=γh_cz={gs}×{abs(h2)}={pcz}kPa"# gs,h2,pcz
    # LaTeX格式
    latex_list1 = r'p_{cz}=\gamma h_{cz}=' + str(gs) + r'\times' + f'{abs(h2)}={pcz}kPa'
    # 土厚
    hs = max(round(abs(h1) - z, 1), 0) # h1是基础底标高, z是基础高度, 得到填土的高度, 当存在填土时, Gk公式需要加上填土的自重
    if hs != 0 :
        # UnicodeMath格式
        # latex_list2 = f"p_c=γh_c={gs}×{abs(hs)}={pc}kPa"# gs,h1,pc
        # LaTeX格式
        latex_list2 = r'p_c=\gamma h_c=' + str(gs) + r'\times' + f'{abs(hs)}={pc}kPa'
    else:
        # UnicodeMath格式
        # latex_list2 = f"p_c=γh_c=0kPa"# gs,h1,pc
        # LaTeX格式
        latex_list2 = r'p_c=\gamma\ h_c=0kPa'
    # latex_list3 = [f"p_z=(lb(p_k-p_c))/((b+2ztanθ)(l+2ztanθ))=({x}×{y}×({pk}-{abs(gs * h1)}))/(({x}+2×{round(abs(h2-h1),1)}×tan{angle})({y}+2×{round(abs(h2-h1),1)}×tan{angle}))={pz}kPa"]# x, y, pk, gs, h1, h2, angle
    # UnicodeMath格式
    # latex_list3 = f"p_z=(lb(p_k-p_c))/((b+2ztanθ)(l+2ztanθ))={pz}kPa"# x, y, pk, gs, h1, h2, angle
    # LaTeX格式
    latex_list3 = GB50007_2011_5_2_7_3 + f'={pz}kPa'
    if if_satisfied_faz == "":
        latex_list4 = ''
    elif if_satisfied_faz == "满足":
        # UnicodeMath格式
        # latex_list4 = f"p_z+p_cz={pz}+{pcz}={pz+pcz}kPa<faz={faz}kPa"
        # LaTeX格式
        latex_list4 = r'p_z+p_{cz}=' + f'{pz}+{pcz}={pz+pcz}' + r'kPa<f_{az}=' + f'{faz}kPa'
    else:
        # UnicodeMath格式
        # latex_list4 = f"p_z+p_cz={pz}+{pcz}={pz+pcz}kPa>faz={faz}kPa"
        # LaTeX格式
        latex_list4 = r'p_z+p_{cz}=' + f'{pz}+{pcz}={pz+pcz}' + r'kPa>f_{az}=' + f'{faz}kPa'
    # 公式替换
    repl_docstr_with_omath(aDoc,selection,latex_list1,'{foundation_pcz}')
    repl_docstr_with_omath(aDoc,selection,latex_list2,'{foundation_pc}')
    repl_docstr_with_omath(aDoc,selection,latex_list3,'{foundation_pz}')
    repl_docstr_with_omath(aDoc,selection,latex_list4,'{if_satisfied_foundation_pk_pz}')
    

# 扩大基础地基应力计算——基底净反力
def word_p_pmax_pmin(word, aDoc, l, b, Ksi, Gk, pk, pkmax, pkmin, p, pmax, pmin):
    selection = word.Selection
    # LaTeX公式
    # latex_list1, latex_list2, latex_list3 = foundation_p_pmax_pmin_LaTeX(l, b, Ksi, Gk, pk, pkmax, pkmin, p, pmax, pmin)
    # UnicodeMath格式
    # latex_list1 = f"p_j=K_s (p_k-G_k/A)={Ksi}×({pk}-{Gk}/({l}×{b}))={p}kpa"
    # latex_list2 = f"p_jmax=K_s (p_kmax-G_k/A)={Ksi}×({pkmax}-{Gk}/({l}×{b}))={pmax}kpa"
    # LaTeX格式
    latex_list1 = GB50007_2011_5_2_8_pj + '=' + str(Ksi) + r'\times\left(' + str(pk) + r'-\frac{' + str(Gk) + '}{' + str(l) + r'\times' + str(b) + r'}\right)=' f'{p}kPa'
    latex_list2 = GB50007_2011_5_2_11_pjmax + '=' + str(Ksi) + r'\times\left(' + str(pkmax) + r'-\frac{' + str(Gk) + '}{' + str(l) + r'\times' + str(b) + r'}\right)=' + f'{pmax}kPa'
    if pmin > 0:
        # UnicodeMath格式
        # latex_list3 = f"p_jmin=K_s (p_kmin-G_k/A)={Ksi}×({pkmin}-{Gk}/({l}×{b}))={pmin}kpa"
        # LaTeX格式
        latex_list3 = GB50007_2011_5_2_11_pjmin + '=' + str(Ksi) + r'\times\left(' + str(pkmin) + r'-\frac{' + str(Gk) + '}{' + str(l) + r'\times' + str(b) + r'}\right)=' + f'{pmin}kPa'
    else:
        # UnicodeMath格式
        # latex_list3 = f"p_jmin=0kpa"
        # LaTeX格式
        latex_list3 = r'p_{jmin}=0kPa'
    # 公式替换
    repl_docstr_with_omath(aDoc,selection,latex_list1,'{foundation_pj}')
    repl_docstr_with_omath(aDoc,selection,latex_list2,'{foundation_pjmax}')
    repl_docstr_with_omath(aDoc,selection,latex_list3,'{foundation_pjmin}')


# 桩基础——单桩竖向承载力
def word_Pile_Ra(word, aDoc, Qsk, Qpk, Quk, Ra, Pile_type, Pile_Nk):
    selection = word.Selection
    # LaTeX公式
    # LaTeX1, LaTeX2 = JGJ94_Quk_Ra_LaTeX(Qsk, Qpk, Quk, Ra, Pile_type, Pile_Nk)
    # word公式(LaTeX格式)
    # 参数计算
    if Pile_type == "钢管打入桩":
        LaTeX1 = JGJ94_2008_5_3_7_1 + '=' + str(Qsk) + '+' + str(Qpk) + '=' + str(Quk) + 'kN'
    if Pile_type == "混凝土预制桩":
        LaTeX1 = JGJ94_2008_5_3_8_1 + '=' + str(Qsk) + '+' + str(Qpk) + '=' + str(Quk) + 'kN'
    if Pile_type == "混凝土灌注桩":
        LaTeX1 = JGJ94_2008_5_3_6 + '=' + str(Qsk) + '+' + str(Qpk) + '=' + str(Quk) + 'kN'
    # 公式(LaTeX格式)
    if Ra >= Pile_Nk:
        LaTeX2 = JGJ94_2008_5_2_2 + r'=\frac{1}{2}\times' + str(Quk) + '=' + str(Ra) + 'kN>N_k=' + str(Pile_Nk) + 'kN'
    else:
        LaTeX2 = JGJ94_2008_5_2_2 + r'=\frac{1}{2}\times' + str(Quk) + '=' + str(Ra) + 'kN<N_k=' + str(Pile_Nk) + 'kN'
    # 公式替换
    doc_LaTeX = [LaTeX1, LaTeX2]
    repl_docstr = ["{JGJ94_Quk}", "{JGJ94_Ra}"]
    for x, y in zip(doc_LaTeX, repl_docstr):
        repl_docstr_with_omath(aDoc, selection, x, y)
    

# 桩基础——配筋率小于0.65%的灌注桩的水平承载力
def word_Pile_Rha1(word, aDoc, Pile_D, Pile_p, Pile_Hk, Pile_as, friction_n, Cohesion_n, Soil_vb, m_n, b0, EI, alpha, Rha, Es, Ec, An, W0):
    selection = word.Selection
    # LaTeX1, LaTeX2, LaTeX3, LaTeX4, LaTeX5 = JGJ94_Rha1_LaTeX(Pile_D, Pile_p, Pile_Hk, Pile_as, friction_n, Cohesion_n, Soil_vb, m_n, b0, EI, alpha, Rha, Es, Ec, An, W0)
    # 公式(LaTeX格式)
    # LaTeX格式文本
    LaTeX1 = JGJ_120_2012_4_1_6 + r'=\frac{0.2\times{' + str(friction_n) + r'}^2-' + str(friction_n) + '+' + str(Cohesion_n) + r'}{' + str(Soil_vb) + r'}\times1000=' + str(m_n) + 'kN/m^4'
    LaTeX2 = JGJ94_2008_5_7_5 + r'=\sqrt[5]{\frac{' + str(m_n) + r'\times' + str(b0) + r'}{' + str(EI) + r'}}=' + str(alpha) + r'\left(1/m\right)'
    LaTeX4 = JGJ94_2008_5_7_2_1_W0 + r'=\frac{' + str(Pile_D) + r'\pi}{32}\times\left[{' + str(Pile_D) + r'}^2+2\left(\frac{' + str(Es) + '}{' + str(Ec) + r'}-1\right)\times' + str(Pile_p/100) + r'\times\left(' + str(Pile_D-Pile_as) + r'\right)^2\right]=' + str(W0) + 'mm^3'
    LaTeX5 = JGJ94_2008_5_7_2_1_An + r'=\frac{{' + str(Pile_D) + r'}^2\pi}{4}\times\left[1+\left(\frac{' + str(Es) + '}{' + str(Ec) + r'}-1\right)\times' + str(Pile_p/100) + r'\right]=' + str(An) + 'mm^2'
    if Rha >= Pile_Hk:
        LaTeX3 = JGJ94_2008_5_7_2_1 + '=' + str(Rha) + 'kN>H_k=' + str(Pile_Hk) + 'kN'
    else:
        LaTeX3 = JGJ94_2008_5_7_2_1 + '=' + str(Rha) + 'kN<H_k=' + str(Pile_Hk) + 'kN'
    # 公式替换
    doc_LaTeX = [LaTeX1, LaTeX2, LaTeX3, LaTeX4, LaTeX5]
    repl_docstr = ["{JGJ94_Subgrade_Reaction_Coefficient}", "{JGJ94_Displacement_Coefficient}", "{JGJ94_Lateral_Capacity}", "{JGJ94_Transformed_Modulus}", "{JGJ94_Transformed_Area}"]
    for x, y in zip(doc_LaTeX, repl_docstr):
        repl_docstr_with_omath(aDoc, selection, x, y)


# 桩基础——钢桩、预制桩、配筋率大于等于0.65%的灌注桩的水平承载力
def word_Pile_Rha2(word, aDoc, Pile_Hk, friction_n, Cohesion_n, Soil_vb, m_n, b0, EI, alpha, Rha, Vx):
    selection = word.Selection
    # LaTeX公式
    # LaTeX1, LaTeX2, LaTeX3 = JGJ94_Rha2_LaTeX(Pile_Hk, friction_n, Cohesion_n, Soil_vb, m_n, b0, EI, alpha, Rha, Vx)
    # 公式(LaTeX格式)
    LaTeX1 = JGJ_120_2012_4_1_6 + r'=\frac{0.2\times{' + str(friction_n) + r'}^2-' + str(friction_n) + '+' + str(Cohesion_n) + r'}{' + str(Soil_vb) + r'}\times1000=' + str(m_n) + 'kN/m^4'
    LaTeX2 = JGJ94_2008_5_7_5 + r'=\sqrt[5]{\frac{' + str(m_n) + r'\times' + str(b0) + r'}{' + str(EI) + r'}}=' + str(alpha) + r'\left(1/m\right)'
    if Rha >= Pile_Hk:
        LaTeX3 = JGJ94_2008_5_7_2_2 + r'=0.75\times\frac{{' + str(alpha) + r'}^3\times' + str(EI) + r'}{' + str(Vx) + r'}\times' + str(round(Soil_vb/1000, 3)) + '=' + str(Rha) + 'kN>H_k=' + str(Pile_Hk) + 'kN'
    else:
        LaTeX3 = JGJ94_2008_5_7_2_2 + r'=0.75\times\frac{{' + str(alpha) + r'}^3\times' + str(EI) + r'}{' + str(Vx) + r'}\times' + str(round(Soil_vb/1000, 3)) + '=' + str(Rha) + 'kN<H_k=' + str(Pile_Hk) + 'kN'
    # 公式替换
    doc_LaTeX = [LaTeX1, LaTeX2, LaTeX3]
    repl_docstr = ["{JGJ94_Subgrade_Reaction_Coefficient}", "{JGJ94_Displacement_Coefficient}", "{JGJ94_Lateral_Capacity}"]
    for x, y in zip(doc_LaTeX, repl_docstr):
        repl_docstr_with_omath(aDoc, selection, x, y)


# 荷载参数公式替换=============================================================================================================================================================
def Load_Parameter_Fuc(word, aDoc, bDoc, Wind_Standard_Codename, Flow_Standard_Codename, Wind_param_lst, Flow_param_lst): 
    # 告诉word现在这个 word.Selection 是 bDoc 的, 否则可能因为在两个word之间互相检索导致word文件并非正确的检索文件
    search_range = bDoc.Content
    search_range.Select()
    # 替换公式的selection
    selection = word.Selection
    # 风荷载
    if Wind_Standard_Codename == "公路桥梁抗风设计规范":
        # 基本参数
        # U10, DB, kt, Z, L, Ud_Formula, Gv, kf, kh, kc, Ud, Us10, a0, Ug, Fg1, Fg2, FF, CH = Wind_param_lst
        U10, kt, Z, L, n, CH, DB, Ud_Formula, Gv1, Gv2, kf, kh, kc, Us10, Ud, a0, Ug1, Ug2, Fg1, Fg2, Fg3, A, FF = Wind_param_lst
        if Ud_Formula == "Ud = kf·kt·kh·U10":
            # 公式替换前处理计算书中的 Ud_Formula1_3360
            Ud_Formula_str = JTG_3360_01_2018_4_2_6_2 + f'={str(kf)}' + r'\times' + str(kt) + r'\times' + str(kh) + r'\times' + f'{U10}={Ud}m/s'
            repl_docstr_with_omath(bDoc, selection, Ud_Formula_str, '{Ud_Formula1_3360}')
            # 粘贴至前处理计算书中的 Ud_Formula_3360_Paste_Here
            WordA_PasteTo_WordB_DeleteLines(word, bDoc, bDoc, '{Ud_Formula1_Start_3360}', '{Ud_Formula1_End_3360}', '{Ud_Formula_3360_Paste_Here}')
        elif Ud_Formula == "Ud = kf·(Z/10)^α0·Us10":
            # 公式替换前处理计算书中的 Uc_formula2_3360
            Uc_Formula_str = JTG_3360_01_2018_4_2_4 + f'={kc}' + r'\times' + f'{U10}={Us10}m/s'
            repl_docstr_with_omath(bDoc, selection, Uc_Formula_str, '{Uc_formula2_3360}')
            # 公式替换前处理计算书中的 Ud_Formula2_3360
            Ud_Formula_str = JTG_3360_01_2018_4_2_6_1 + f'={kf}' +r'\times' + str(Z/10) + '^{' + str(a0) + r'}\times' + f'{Us10}={Ud}m/s'
            repl_docstr_with_omath(bDoc, selection, Ud_Formula_str, '{Ud_Formula2_3360}')
            # 粘贴至前处理计算书中的 Ud_Formula_3360_Paste_Here
            WordA_PasteTo_WordB_DeleteLines(word, bDoc, bDoc, '{Ud_Formula2_Start_3360}', '{Ud_Formula2_End_3360}', '{Ud_Formula_3360_Paste_Here}')
        # Ug_Formula_3360
        Ug_Formula_str = JTG_3360_01_2018_5_2_1 + f'={Gv1}' + r'\times' + f'{Ud}={Ug1}m/s'
        repl_docstr_with_omath(bDoc, selection, Ug_Formula_str, '{Ug_Formula_3360}')
        # Fg_Formula_3360_1
        Fg_Formula_str_1 = JTG_3360_01_2018_5_3_1 + r'=0.5\times1.25' + r'\times{' + str(Ug1) + r'}^2\times' + f'{CH}={Fg1}kPa'
        repl_docstr_with_omath(bDoc, selection, Fg_Formula_str_1, '{Fg_Formula_3360_1}')
        # Fg_Formula_3360_2
        Fg_Formula_str_2 = JTG_3360_01_2018_5_3_1 + r'=0.5\times1.25' + r'\times{' + str(Ug1) + r'}^2\times' + f'1.7={Fg2}kPa'
        repl_docstr_with_omath(bDoc, selection, Fg_Formula_str_2, '{Fg_Formula_3360_2}')
        # Bailey_A_3360
        Bailey_A_str = r'A=\frac{1-{0.66}^' + str(n) + r'}{1-0.66}\times0.3\times4.5=' + f'{A}m^2'
        repl_docstr_with_omath(bDoc, selection, Bailey_A_str, '{Bailey_A_3360}')
        # Bailey_Wind_Formula_3360
        FF_Formula_str = r'F=\frac{F_g}{D}A=' + str(Fg2) + r'\times3.6=' + f'{FF}kN'
        repl_docstr_with_omath(bDoc, selection, FF_Formula_str, '{Bailey_Wind_Formula_3360}')
        # 粘贴至总计算书 Wind_Load_Paste_Here
        WordA_PasteTo_WordB_DeleteLines(word, bDoc, aDoc, '{Wind_Load_Start_3360}', '{Wind_Load_End_3360}', '{Wind_Load_Paste_Here}')
    # 没有风荷载则删除
    elif Wind_Standard_Codename == "无":
        Delete_Word_FindText_Paragraphs(aDoc, '{Wind_Load_Paste_Here}')
    # 水荷载
    if Flow_Standard_Codename == "港口工程荷载规范":
        # 基本参数
        p, V, H, Fw, D = Flow_param_lst
        # 公式替换前处理计算书中的 Fw_Formula_144
        Fw_Formula_str = JTS_144_1_2010_13_0_1 + r'=0.5\times0.73\times' + str(p) + r'\times{' + str(V) + r'}^2\times' + str(D) + r'\times1.0=' + f'{Fw}kN'
        repl_docstr_with_omath(bDoc, selection, Fw_Formula_str, '{Fw_Formula_144}')
        # 粘贴至总计算书 Flow_Load_Paste_Here
        WordA_PasteTo_WordB_DeleteLines(word, bDoc, aDoc, '{Flow_Load_Start_144}', '{Flow_Load_End_144}', '{Flow_Load_Paste_Here}')
    # 没有水流力则删除
    elif Flow_Standard_Codename == "无":
        Delete_Word_FindText_Paragraphs(aDoc, '{Flow_Load_Paste_Here}')

#=====================================================================================================================================================================

# 图片替换函数

# 地面以上结构的相关图片
def Word_above_structure_picture(aDoc, folder_path):
    # 有以下图片：模型前处理、小肋组合应力图、小肋剪应力图、小肋竖向变形图、贝雷弦杆轴力图、贝雷竖杆轴力图、贝雷斜杆轴力图、贝雷竖向变形图、分配梁组合应力图、分配梁剪应力图、分配梁竖向变形图、钢管桩轴力图、钢管桩弯矩图、联结系组合应力图、联结系剪合应力图
    # 图片名
    pic_name_lst = ["模型前处理.jpg", 
                    "小肋组合应力图.jpg", "小肋剪应力图.jpg", "小肋竖向变形图.jpg",
                    "贝雷弦杆轴力图.jpg", "贝雷竖杆轴力图.jpg", "贝雷斜杆轴力图.jpg", "贝雷竖向变形图.jpg",
                    "分配梁组合应力图.jpg", "分配梁剪应力图.jpg", "分配梁竖向变形图.jpg",
                    "钢管桩轴力图.jpg", "钢管桩弯矩图.jpg", 
                    "联结系组合应力图.jpg", "联结系剪应力图.jpg",
                ]
    # 图片对应的特殊文本
    pic_repl_docstr_lst = ['{model_ picture1}',
                        '{xiaole_ picture1}', '{xiaole_ picture2}', '{xiaole_picture3}',
                        '{Bailey-xiangan-picture}', '{Bailey-shugan-picture}', '{Bailey-xiegan-picture}', '{bailey_shape_Dz_picture}',
                        '{fenpeiliang-picture1}', '{fenpeiliang-picture2}', '{fenpeiliang_picture3}',
                        '{steel_pile_force_picture1}', '{steel_pile_force_picture2}',
                        '{lianjiexi_picture1}', '{lianjiexi_picture2}',
                    ]
    search_range = aDoc.Content
    search_range.Select()
    for i in range(len(pic_name_lst)):
        # 图片的路径
        doc_pic_path = os.path.join(folder_path, pic_name_lst[i])
        # 替换的特殊文本
        pic_repl_docstr = pic_repl_docstr_lst[i]
        # 特殊文本替换为对应的图片
        repl_docstr_with_pic(aDoc, str(pic_repl_docstr), str(doc_pic_path))


# 扩大基础——轴心荷载作用图片
def word_Pad_Foundation_Fonly(aDoc, folder_path):
    # 图片替换
    pic_name_lst = ["基础竖向反力图.jpg"]
    pic_repl_docstr_lst = ['{Midas_reaction_picture_biaozhun_Fxyz}']
    for i in range(len(pic_name_lst)):
        # 图片的路径
        doc_pic_path = os.path.join(folder_path, pic_name_lst[i])
        # 替换的特殊文本
        pic_repl_docstr = pic_repl_docstr_lst[i]
        # 特殊文本替换为对应的图片
        repl_docstr_with_pic(aDoc, str(pic_repl_docstr), str(doc_pic_path))

    
# 扩大基础——偏心荷载作用图片
def word_Pad_Foundation_FandM(aDoc, folder_path):
    # 图片替换
    pic_name_lst = ["基础竖向反力图.jpg", "基础弯矩图.jpg"]
    pic_repl_docstr_lst = ['{Midas_reaction_picture_biaozhun_Fxyz}', '{Midas_reaction_picture_biaozhun_Mxyz}']
    for i in range(len(pic_name_lst)):
        # 图片的路径
        doc_pic_path = os.path.join(folder_path, pic_name_lst[i]) 
        # 替换的特殊文本
        pic_repl_docstr = pic_repl_docstr_lst[i]
        # 特殊文本替换为对应的图片
        repl_docstr_with_pic(aDoc, str(pic_repl_docstr), str(doc_pic_path))

#=====================================================================================================================================================================

# 文本替换函数

# 计算依据文本替换==============================================================================================================================================================
def Calculation_Basis_TxT(word, aDoc, bDoc, Wind_Standard, Flow_Standard, Foundation_Standard):
    # 如果风规和水规一样或是没有水规
    if Wind_Standard == Flow_Standard and Flow_Standard == "":
        # 删除总计算书中的 Selected_Wind_Flow_Standard_Paste_Here
        Delete_Word_FindText_Paragraphs(aDoc, '{Selected_Wind_Flow_Standard_Paste_Here}')
    elif Wind_Standard == Flow_Standard or Flow_Standard == "":
        # 在前处理计算书中替换 Selected_Wind_Standard_1 
        repl_docstr_with_strlst(bDoc, '{Selected_Wind_Standard_1}', Wind_Standard)
        # 复制粘贴进总计算书中
        WordA_PasteTo_WordB_DeleteLines(word, bDoc, aDoc, '{Selected_Wind_Flow_Standard_Start_situation1}', '{Selected_Wind_Flow_Standard_End_situation1}', '{Selected_Wind_Flow_Standard_Paste_Here}')
    else:
        # 在前处理计算书中替换 Selected_Wind_Standard_2
        # 在前处理计算书中替换 Selected_Flow_Standard_2
        repl_docstr_with_strlst(bDoc, '{Selected_Wind_Standard_2}', Wind_Standard)
        repl_docstr_with_strlst(bDoc, '{Selected_Flow_Standard_2}', Flow_Standard)
        # 复制粘贴进总计算书中
        WordA_PasteTo_WordB_DeleteLines(word, bDoc, aDoc, '{Selected_Wind_Flow_Standard_Start_situation2}', '{Selected_Wind_Flow_Standard_End_situation2}', '{Selected_Wind_Flow_Standard_Paste_Here}')
    # 在总计算书中替换 Selected_foundation_Standard
    if Foundation_Standard == "":
        # 删除总计算书中的 Selected_foundation_Standard_Paste_Here
        Delete_Word_FindText_Paragraphs(aDoc, '{Selected_foundation_Standard_Paste_Here}')
    else:
        # 替换前处理计算书中的 Selected_foundation_Standard
        repl_docstr_with_strlst(bDoc, '{Selected_foundation_Standard}', Foundation_Standard)
        # 总计算书中的 Selected_foundation_Standard_Paste_Here
        WordA_PasteTo_WordB_DeleteLines(word, bDoc, aDoc, '{Selected_foundation_Standard_Start}', '{Selected_foundation_Standard_End}', '{Selected_foundation_Standard_Paste_Here}')


 # 荷载参数文本替换=============================================================================================================================================================
def Load_Parameter_TxT(aDoc, bDoc, Model_Gz, Concrete_Gs, Formwork_Load, Construction_Load, Wind_param_lst, Flow_param_lst):
    # 自重, 在总计算书中替换 Model_Gz
    repl_docstr_with_strlst(aDoc, '{Model_Gz}', Model_Gz)
    # 混凝土荷载, 在总计算书中替换 Concrete_Gs
    repl_docstr_with_strlst(aDoc, '{Concrete_Gs}', Concrete_Gs)
    # 模板荷载, 在总计算书中替换 Formwork_Load
    repl_docstr_with_strlst(aDoc, '{Formwork_Load}', Formwork_Load)
    # 施工荷载, 在总计算书中替换 Construction_Load
    repl_docstr_with_strlst(aDoc, '{Construction_Load}', Construction_Load)
    # 风荷载
    # U10, DB, kt, Z, L, Ud_Formula, Gv, kf, kh, kc, Ud, Us10, a0, Ug, Fg1, Fg2, FF, CH = Wind_param_lst
    if Wind_param_lst != ['None']:
        U10, kt, Z, L, n, CH, DB, Ud_Formula, Gv1, Gv2, kf, kh, kc, Us10, Ud, a0, Ug1, Ug2, Fg1, Fg2, Fg3, A, FF = Wind_param_lst
        if Ud_Formula == "Ud = kf·kt·kh·U10":
            # 前处理计算书中需要替换的检索词
            repl_docstr = ['{Wind_U10_3360_1}', '{Wind_CH_3360}', '{Bailey_n_3360}', '{Ud_Formula1_DB_3360}', '{Ud_Formula1_DB_value_3360}', '{Ud_Formula1_kf_3360}', '{Ud_Formula1_kt_3360}', '{Ud_Formula1_kh_3360}']
            repl_valstr = [U10, CH, n, DB, Gv1, kf, kt, kh]
            [repl_docstr_with_strlst(bDoc, y, str(x)) for x, y in zip(repl_valstr, repl_docstr)]
        elif Ud_Formula == "Ud = kf·(Z/10)^α0·Us10":
            # 前处理计算书中需要替换的检索词
            repl_docstr = ['{Wind_U10_3360_1}', '{Wind_CH_3360}', '{Bailey_n_3360}', '{Ud_Formula2_DB_3360}', '{Ud_Formula2_DB_value_3360}', '{Ud_Formula2_kf_3360}', '{Ud_Formula2_kc_3360}']
            repl_valstr = [U10, CH, n, DB, Gv1, kf, kc]
            [repl_docstr_with_strlst(bDoc, y, str(x)) for x, y in zip(repl_valstr, repl_docstr)]
    # 水荷载
    if Flow_param_lst != ['None']:
        p, V, H, Fw, D = Flow_param_lst
        repl_docstr = ['{Flow_V_144}', '{Flow_Density_144}', '{Flow_Pile_D_144}']
        repl_valstr = [V, p, D]
        [repl_docstr_with_strlst(bDoc, y, str(x)) for x, y in zip(repl_valstr, repl_docstr)] 


# 材料特性文本替换=============================================================================================================================================================
def Material_Properties_TxT(word, aDoc, bDoc, Material_Properties_dict):
    # 混凝土材料特性表
    concrete_properties_lst = Material_Properties_dict["Concrete"]
    concrete_properties_txt = []
    if concrete_properties_lst != []:
        for lst in concrete_properties_lst:
            concrete_properties_txt.append([f"{lst[0]}：混凝土轴心抗压{lst[1]}Mpa，轴心抗拉{lst[2]}Mpa，弹性模量{format_scientific_unicode(float(lst[3]))}Mpa"])
    concrete_properties_txt = str_join_n(concrete_properties_txt)
    # print(concrete_properties_txt)
    if concrete_properties_txt != "":
        # 在前处理计算书中替换 Concrete_Material_Properties
        repl_docstr_with_strlst(bDoc, '{Concrete_Material_Properties}', concrete_properties_txt.replace('\n', '\r'))
        # 在前处理计算书中粘贴 粘贴在 Concrete_Material_Properties_Paste_Here
        WordA_PasteTo_WordB_DeleteLines(word, bDoc, bDoc, '{Concrete_Material_Properties_Start}', '{Concrete_Material_Properties_End}', '{Concrete_Material_Properties_Paste_Here}')
    else:
        # 删除改行
        Delete_Word_FindText_Paragraphs(bDoc, '{Concrete_Material_Properties_Paste_Here}')
    # 钢筋材料特性表
    rebar_properties_lst = Material_Properties_dict["Rebar"]
    rebar_properties_txt = []
    if rebar_properties_lst != []:
        for lst in rebar_properties_lst:
            rebar_properties_txt.append([f"{lst[0]}：抗拉压强度设计值{lst[1]}Mpa，弹性模量{format_scientific_unicode(float(lst[2]))}Mpa"])
    rebar_properties_txt = str_join_n(rebar_properties_txt)
    # print(rebar_properties_txt)
    if rebar_properties_txt != "":
        # 在前处理计算书中替换 Rebar_Material_Properties , word中换行符是'\r'
        repl_docstr_with_strlst(bDoc, '{Rebar_Material_Properties}', rebar_properties_txt.replace('\n', '\r'))
        # 在前处理计算书中粘贴 粘贴在 Rebar_Material_Properties_Paste_Here
        WordA_PasteTo_WordB_DeleteLines(word, bDoc, bDoc, '{Rebar_Material_Properties_Start}', '{Rebar_Material_Properties_End}', '{Rebar_Material_Properties_Paste_Here}')
    else:
        # 删除改行
        Delete_Word_FindText_Paragraphs(bDoc, '{Rebar_Material_Properties_Paste_Here}')
    # 将材料特性粘贴进总计算书
    if concrete_properties_txt == '' and rebar_properties_txt == '':
        # 删除改行
        Delete_Word_FindText_Paragraphs(aDoc, '{Material_Properties_Paste_Here}')
    else:
        WordA_PasteTo_WordB_DeleteLines(word, bDoc, aDoc, '{Material_Properties_Start}', '{Material_Properties_End}', '{Material_Properties_Paste_Here}')


 # 工况分析文本替换=============================================================================================================================================================
def Case_Analysis_TxT(aDoc, Flow_Standard):
    if Flow_Standard == "":
        case_analysis = "自重+施工荷载+模板荷载+混凝土荷载+风荷载"
    else:
        case_analysis = "自重+施工荷载+模板荷载+混凝土荷载+风荷载+水流力"
    print("工况分析:", case_analysis)
    # 在总计算书中替换 Selected_foundation_Standard
    repl_docstr_with_strlst(aDoc, '{Case_Analysis}', case_analysis)


# 荷载组合文本替换=============================================================================================================================================================
def Load_Combination_TxT(word, aDoc, bDoc, Flow_Standard):
    if Flow_Standard == "":
        # 复制粘贴进总计算书中
        WordA_PasteTo_WordB_DeleteLines(word, bDoc, aDoc, '{Load_Combination_Start_situation1}', '{Load_Combination_End_situation1}', '{Load_Combination_Paste_Here}')
    else:
        # 复制粘贴进总计算书中
        WordA_PasteTo_WordB_DeleteLines(word, bDoc, aDoc, '{Load_Combination_Start_situation2}', '{Load_Combination_End_situation2}', '{Load_Combination_Paste_Here}')


# 文本替换——地面以上结构的相关结果，从Midas中获取
def word_above_structure_TXT_Table(aDoc, model_post_dict, Fx, Mm, sigma):
    # 从字典中提取变量
    xiaole_max_cb, xiaole_max_shear, xiaole_max_Dz = model_post_dict["model_post_xiaolei"]
    bailey_max_xiangan_zhouli, bailey_max_shugan_zhouli, bailey_max_xiegan_zhouli, bailey_max_Dz = model_post_dict["model_post_bailey"]
    fenpeiliang_max_cb, fenpeiliang_max_shear, fenpeiliang_max_Dz = model_post_dict["model_post_fenpeiliang"]
    lianjiexi_max_cb, lianjiexi_max_shear = model_post_dict["model_post_lianjiexi"]
    # 表格数值
    value_name_lst = {"小肋组合应力" : xiaole_max_cb, 
                    "小肋剪应力" : xiaole_max_shear, 
                    "小肋竖向变形" : xiaole_max_Dz,
                    "贝雷弦杆内力" : bailey_max_xiangan_zhouli, 
                    "贝雷竖杆内力" : bailey_max_shugan_zhouli,
                    "贝雷斜杆内力" : bailey_max_xiegan_zhouli, 
                    "贝雷竖向变形" : bailey_max_Dz,
                    "分配梁组合应力" : fenpeiliang_max_cb, 
                    "分配梁剪应力" : fenpeiliang_max_shear, 
                    "分配梁竖向变形" : fenpeiliang_max_Dz,
                    "联结系组合应力" : lianjiexi_max_cb, 
                    "联结系剪应力" : lianjiexi_max_shear,
    }
    # 数值对应的特殊文本
    value_repl_docstr_lst = {"小肋组合应力" : '{xiaole1}', 
                            "小肋剪应力" : '{xiaole2}', 
                            "小肋竖向变形" : '{xiaole_Zmax}',
                            "贝雷弦杆内力" : '{Bailey-xiangan}', 
                            "贝雷竖杆内力" : '{Bailey-shugan}', 
                            "贝雷斜杆内力" : '{Bailey-xiegan}', 
                            "贝雷竖向变形" : '{bailey_Zmax}',
                            "分配梁组合应力" : '{fenpeiliang1}', 
                            "分配梁剪应力" : '{fenpeiliang2}',
                            "分配梁竖向变形" : '{fenpeiliang_Zmax}',
                            "联结系组合应力" : '{lianjiexi1}', 
                            "联结系剪应力" : '{lianjiexi2}',
                        }
    # 文本替换
    for key, value in value_name_lst.items():
        # 表格数值，转换成绝对值字符串
        doc_value = str(abs(float(value)))
        # 替换的特殊文本
        value_repl_docstr = value_repl_docstr_lst[key]
        # 特殊文本替换为对应的图片
        repl_docstr_with_strlst(aDoc, value_repl_docstr, doc_value)

    # 表格判断值
    satisfied_name_lst = {"小肋" : "满足",
                        "贝雷弦杆" : "满足",
                        "贝雷竖杆" : "满足",
                        "贝雷斜杆" : "满足",
                        "分配梁" : "满足",
                        "联结系" : "满足",
                    }
    satisfied_value_lst = {"小肋" : {"小肋组合应力" : 215, "小肋剪应力" : 125},
                        "贝雷弦杆" : {"贝雷弦杆内力" : 560},
                        "贝雷竖杆" : {"贝雷竖杆内力" : 210},
                        "贝雷斜杆" : {"贝雷斜杆内力" : 171.5},
                        "分配梁" : {"分配梁组合应力" : 205, "分配梁剪应力" : 120},
                        "联结系" : {"联结系组合应力" : 215, "联结系剪应力" : 125},
                    }
    # 判断值对应的特殊文本
    satisfied_repl_docstr_lst = {"小肋" : '{xiaole3}',
                                "贝雷弦杆" : '{bailey_satisfied1}', 
                                "贝雷竖杆" : '{bailey_satisfied2}', 
                                "贝雷斜杆" : '{bailey_satisfied3}',
                                "分配梁" : '{fenpeiliang3}',
                                "联结系" : '{lianjiexi3}',
                            }
    # 文本替换
    for key1, value1 in satisfied_value_lst.items():
        for key2, value2 in value1.items():
            if abs(float(value_name_lst[key2])) > satisfied_value_lst[key1][key2]:
                satisfied_name_lst[key1] = "不满足"
        # 表格判断值
        doc_value = satisfied_name_lst[key1]
        # 替换的特殊文本
        value_repl_docstr = satisfied_repl_docstr_lst[key1]
        # 特殊文本替换
        repl_docstr_with_strlst(aDoc, value_repl_docstr, doc_value)

    # 非表格的文本内容
    if sigma < 215:
        if_steelpile_stability_satisfied = "满足"
    else:
        if_steelpile_stability_satisfied = "不满足"
    value_name_lst = [round(abs(Fx),1), round(abs(Mm),1), if_steelpile_stability_satisfied]
    value_repl_docstr_lst = ['{steelpile_Fmax}', '{steelpile_Mmax}', '{yagan_satisfied}']
    for x, y in zip(value_name_lst, value_repl_docstr_lst):
        repl_docstr_with_strlst(aDoc, y, str(x))


# 文本替换——扩大基础无软弱下卧层轴心荷载作用
def word_pk_TXT_Table(aDoc, Fmax_Fz, foundation_x, foundation_y, foundation_z, Foundation_Base_Level, foundation_fa, if_satisfied_fa1):
    # 文本替换
    value_name_lst = [round(float(Fmax_Fz),1), foundation_x, foundation_y, foundation_z, abs(Foundation_Base_Level), foundation_fa, if_satisfied_fa1]
    # print(value_name_lst)
    # print(round(float(Fmax_Fz),1))
    value_repl_docstr_lst = ['{biaozhun_reaction_Fzmax}', '{foundation_x}', '{foundation_y}', '{foundation_z}', '{Foundation_Base_Level}', '{foundation_fa}','{if_satisfied_pk_fa}']
    for i in range(len(value_name_lst)):
        # 表格数值，转换成绝对值字符串
        if type(value_name_lst[i]) == str:
            doc_value = value_name_lst[i]
        elif type(value_name_lst[i]) == int or type(value_name_lst[i]) == float:
            doc_value = str(abs(float(value_name_lst[i])))
        # 替换的特殊文本
        value_repl_docstr = value_repl_docstr_lst[i]
        # 特殊文本替换
        repl_docstr_with_strlst(aDoc, value_repl_docstr, doc_value)


# 文本替换——扩大基础无软弱下卧层偏心荷载作用
def word_pk_pkmax_TXT_Table(aDoc, if_satisfied_fa1, if_satisfied_fa2, Fk1, Fk2, My, foundation_x, foundation_y, foundation_z, Foundation_Base_Level, foundation_fa, Eccentricity):
    # 文本替换
    if_satisfied_pk_pkmax_fa = "不满足"
    if if_satisfied_fa1 == "满足" and if_satisfied_fa2 == "满足":
        if_satisfied_pk_pkmax_fa = "满足"
    value_name_lst = [round(float(Fk1),1), round(float(My),1), round(float(Fk2),1), foundation_x, foundation_y, foundation_z, abs(Foundation_Base_Level), foundation_fa, Eccentricity, if_satisfied_pk_pkmax_fa]
    # print(value_name_lst)
    # print(round(float(Fmax_Fz),1))
    value_repl_docstr_lst = ['{biaozhun_reaction_Fzmax}', '{biaozhun_reaction_Mymax}', '{biaozhun_reaction_FFzmax}', '{foundation_x}', '{foundation_y}', '{foundation_z}', '{Foundation_Base_Level}', '{foundation_fa}', '{Eccentricity}', '{if_satisfied_pk_pkmax_fa}']
    for i in range(len(value_name_lst)):
        # 表格数值，转换成绝对值字符串
        if type(value_name_lst[i]) == str:
            doc_value = value_name_lst[i]
        elif type(value_name_lst[i]) == int or type(value_name_lst[i]) == float:
            doc_value = str(abs(float(value_name_lst[i])))
        # 替换的特殊文本
        value_repl_docstr = value_repl_docstr_lst[i]
        # 特殊文本替换
        repl_docstr_with_strlst(aDoc, value_repl_docstr, doc_value)


# 文本替换——扩大基础有软弱下卧层轴心荷载作用
def word_pk_pz_TXT_Table(aDoc, if_satisfied_fa1, if_satisfied_faz, Fmax_Fz, foundation_x, foundation_y, foundation_z, foundation_fa, Foundation_Base_Level, Weak_Layer_Horizon, foundation_faz, Dispersion_Angle):
    # 文本替换
    if_satisfied_pk_pz_fa_faz = "不满足"
    if if_satisfied_fa1 == "满足" and if_satisfied_faz == "满足":
        if_satisfied_pk_pz_fa_faz = "满足"
    value_name_lst = [round(float(Fmax_Fz),1), foundation_x, foundation_y, foundation_z, foundation_fa, abs(Foundation_Base_Level), abs(Weak_Layer_Horizon), foundation_faz, Dispersion_Angle, if_satisfied_pk_pz_fa_faz]
    # print(value_name_lst)
    # print(round(float(Fmax_Fz),1))
    value_repl_docstr_lst = ['{biaozhun_reaction_Fzmax}', '{foundation_x}', '{foundation_y}', '{foundation_z}', '{foundation_fa}', '{Foundation_Base_Level}', '{Weak_Layer_Horizon}', '{foundation_faz}', '{Dispersion_Angle}', '{if_satisfied_pk_pz_fa_faz}']
    for i in range(len(value_name_lst)):
        # 表格数值，转换成绝对值字符串
        if type(value_name_lst[i]) == str:
            doc_value = value_name_lst[i]
        elif type(value_name_lst[i]) == int or type(value_name_lst[i]) == float:
            doc_value = str(abs(float(value_name_lst[i])))
        # 替换的特殊文本
        value_repl_docstr = value_repl_docstr_lst[i]
        # 特殊文本替换
        repl_docstr_with_strlst(aDoc, value_repl_docstr, doc_value)


# 文本替换——扩大基础有软弱下卧层偏心荷载作用
def word_pk_pkmax_pz_TXT_Table(aDoc, if_satisfied_fa1, if_satisfied_fa2, if_satisfied_faz, Fk1, Fk2, My, foundation_x, foundation_y, foundation_z, foundation_fa, Foundation_Base_Level, Weak_Layer_Horizon, foundation_faz, Dispersion_Angle, Eccentricity):
    # 文本替换
    if_satisfied_pk_pkmax_pz_fa_faz = "不满足"
    if if_satisfied_fa1 == "满足" and if_satisfied_fa2 == "满足" and if_satisfied_faz == "满足":
        if_satisfied_pk_pkmax_pz_fa_faz = "满足"
    value_name_lst = [round(float(Fk1),1), round(float(My),1), round(float(Fk2),1), foundation_x, foundation_y, foundation_z, foundation_fa, abs(Foundation_Base_Level), abs(Weak_Layer_Horizon), foundation_faz, Dispersion_Angle, Eccentricity, if_satisfied_pk_pkmax_pz_fa_faz]
    # print(value_name_lst)
    # print(round(float(Fmax_Fz),1))
    value_repl_docstr_lst = ['{biaozhun_reaction_Fzmax}', '{biaozhun_reaction_Mymax}', '{biaozhun_reaction_FFzmax}', '{foundation_x}', '{foundation_y}', '{foundation_z}', '{foundation_fa}', '{Foundation_Base_Level}', '{Weak_Layer_Horizon}', '{foundation_faz}', '{Dispersion_Angle}', '{Eccentricity}', '{if_satisfied_pk_pkmax_pz_fa_faz}']
    for i in range(len(value_name_lst)):
        # 表格数值，转换成绝对值字符串
        if type(value_name_lst[i]) == str:
            doc_value = value_name_lst[i]
        elif type(value_name_lst[i]) == int or type(value_name_lst[i]) == float:
            doc_value = str(abs(float(value_name_lst[i])))
        # 替换的特殊文本
        value_repl_docstr = value_repl_docstr_lst[i]
        # 特殊文本替换
        repl_docstr_with_strlst(aDoc, value_repl_docstr, doc_value)


# 文本替换——基础抗弯抗剪抗冲切
def word_foundation_Vs_FL_M_TXT_Table(aDoc, concrete_grade, as0,
                                      JCV_sf1, JCV_sf2, JCF_sf1, JCF_sf2, JCM_sf1, JCM_sf2,
                                      JCAv1, JCV1, JCAv2, JCV2, Bhs, ft, JCA01, JCA02, Allowable_Vs1, Allowable_Vs2, 
                                      JCAl1, JCFl1, JCAl2, JCFl2, Bhp, am, h0, Allowable_Fl,
                                      JC_a1, JC_aa, JC_bb, JC_M1, MAs1, JC_M2, MAs2, JC_Md1, JC_Md2, JC_Ms1, JC_Ms2, 
                                      JC_MA1, JC_MA2, JC_Mp1, JC_Mp2, JCM_if1, JCM_if2
                                     ):
    # 文字及数字替换
    if all(item == "满足" for item in [JCV_sf1, JCV_sf2, JCF_sf1, JCF_sf2, JCM_sf1, JCM_sf2]):
        if_foundation_Vs_FL_M_satisfied = "满足"
    else:
        if_foundation_Vs_FL_M_satisfied = "不满足"
    # 要进行替换的文本内容    
    value_name_lst1 = [concrete_grade, as0, if_foundation_Vs_FL_M_satisfied]
    # 要进行替换的表格内容
    value_name_lst2 = [JCAv1, JCV1, Bhs, ft, JCA01, Allowable_Vs1, JCV_sf1]# 剪上/下
    value_name_lst3 = [JCAv2, JCV2, Bhs, ft, JCA02, Allowable_Vs2, JCV_sf2]# 剪左/右
    value_name_lst4 = [JCAl1, JCFl1, Bhp, ft, am, h0, Allowable_Fl, JCF_sf1]# 冲切上/下
    value_name_lst5 = [JCAl2, JCFl2, Bhp, ft, am, h0, Allowable_Fl, JCF_sf2]# 冲切左/右
    value_name_lst6 = [JC_a1, JC_aa, JC_bb, JC_M1, JC_M2]# 弯矩
    value_name_lst7 = [JC_M1, MAs1, f"E{JC_Md1}@{JC_Ms1}", JC_MA1, JC_Mp1, JCM_if1, JCM_sf1]# 配筋上/下
    value_name_lst8 = [JC_M2, MAs2, f"E{JC_Md2}@{JC_Ms2}", JC_MA2, JC_Mp2, JCM_if2, JCM_sf2]# 配筋左/右
    # word中被替换的文本内容
    value_repl_docstr_lst1 = ["{concrete_grade}", "{steel_protective_layer}", "{if_foundation_Vs_FL_M_satisfied}"]
    # word中被替换的表格内容
    value_repl_docstr_lst2 = ["{JCAv1}", "{JCV1}", "{BHS1}", "{JCV_ft1}", "{JCA01}", "{JCV1_V}", "{JCV_sf1}"]# 剪上/下
    value_repl_docstr_lst3 = ["{JCAv2}", "{JCV2}", "{BHS2}", "{JCV_ft2}", "{JCA02}", "{JCV2_V}", "{JCV_sf2}"]# 剪左/右
    value_repl_docstr_lst4 = ["{JCAl1}", "{JCFl1}", "{BHP1}", "{JCF_ft1}", "{JCF_am1}", "{JC_h01}", "{JCF1_F}", "{JCF_sf1}"]# 冲切上/下
    value_repl_docstr_lst5 = ["{JCAl2}", "{JCFl2}", "{BHP2}", "{JCF_ft2}", "{JCF_am2}", "{JC_h02}", "{JCF2_F}", "{JCF_sf2}"]# 冲切左/右
    value_repl_docstr_lst6 = ["{JC_a1}", "{JC_aa}", "{JC_bb}", "{JC_M1}", "{JC_M2}"]# 弯矩
    value_repl_docstr_lst7 = ["{JCM1}", "{MAs1}", "{JCGJ1}", "{JCAs1}", "{JCGJ_p1}", "{JCM_if1}", "{JCM_sf1}"]# 配筋上/下
    value_repl_docstr_lst8 = ["{JCM2}", "{MAs2}", "{JCGJ2}", "{JCAs2}", "{JCGJ_p2}", "{JCM_if2}", "{JCM_sf2}"]# 配筋左/右

    value_name_lst = [value_name_lst1, value_name_lst2, value_name_lst3, value_name_lst4, value_name_lst5, value_name_lst6, value_name_lst7, value_name_lst8]
    value_repl_docstr_lst = [value_repl_docstr_lst1, value_repl_docstr_lst2, value_repl_docstr_lst3, value_repl_docstr_lst4, value_repl_docstr_lst5, value_repl_docstr_lst6, value_repl_docstr_lst7, value_repl_docstr_lst8]
    value_name_lst = [y for x in value_name_lst for y in x]
    value_repl_docstr_lst = [y for x in value_repl_docstr_lst for y in x]
    # print(value_name_lst)
    # print(value_repl_docstr_lst)
    for i in range(len(value_name_lst)):
        # 表格数值，转换成绝对值字符串
        doc_value = value_name_lst[i]
        if type(doc_value) == str:
            pass
        else:
            doc_value = str(abs(float(doc_value)))
        # 替换的特殊文本
        value_repl_docstr = value_repl_docstr_lst[i]
        # 特殊文本替换
        repl_docstr_with_strlst(aDoc, value_repl_docstr, doc_value)
    

# 文本替换——钢管
def word_SteelPile_TXT_Table(aDoc, Pile_D, Pile_t, Pile_H, Pile_Nk, Pile_Hk, Soil_vb, alpha_h, Vm, Vx, Rha, Ra, b0):
    if Rha >= Pile_Hk and Ra > Pile_Nk:
        if_Ra_Rha_satisfied = "满足"
    else:
        if_Ra_Rha_satisfied = "不满足"
    if Pile_D/1000 <= 1:
        Pile_b0 = f"0.9(1.5d+0.5)={b0}"
    else:
        Pile_b0 = f"0.9(d+1)={b0}"
    doc_str = [Pile_D, Pile_t, Pile_H, Pile_Nk, Pile_Hk, Soil_vb, Pile_b0, alpha_h, Vm, Vx, if_Ra_Rha_satisfied]
    repl_docstr = ["{Pile_D}", "{Pile_t}", "{Pile_H}", "{Pile_Nk}", "{Pile_Hk}", "{Soil_vb}", "{Pile_b0}", "{alpha_h}", "{Rha_vm}", "{Rha_vx}", "{if_Ra_Rha_satisfied}"]
    for x,y in zip(doc_str, repl_docstr):
        repl_docstr_with_strlst(aDoc, y, str(x))


# 文本替换——混凝土预制桩
def word_ConcretePile_TXT_Table(aDoc, Pile_D, Pile_t, concrete_grade, Pile_H, Pile_Nk, Pile_Hk, Soil_vb, alpha_h, Vm, Vx, Rha, Ra, b0):
    if Rha >= Pile_Hk and Ra > Pile_Nk:
        if_Ra_Rha_satisfied = "满足"
    else:
        if_Ra_Rha_satisfied = "不满足"
    if Pile_D/1000 <= 1:
        Pile_b0 = f"0.9(1.5d+0.5)={b0}"
    else:
        Pile_b0 = f"0.9(d+1)={b0}"
    doc_str = [Pile_D, Pile_t, concrete_grade, Pile_H, Pile_Nk, Pile_Hk, Soil_vb, Pile_b0, alpha_h, Vm, Vx, if_Ra_Rha_satisfied]
    repl_docstr = ["{Pile_D}", "{Pile_t}", "{concrete_grade}", "{Pile_H}", "{Pile_Nk}", "{Pile_Hk}", "{Soil_vb}", "{Pile_b0}", "{alpha_h}", "{Rha_vm}", "{Rha_vx}", "{if_Ra_Rha_satisfied}"]
    for x,y in zip(doc_str, repl_docstr):
        repl_docstr_with_strlst(aDoc, y, str(x))


# 文本替换——灌注桩配筋大于等于0.65%及灌注桩配筋小于0.65%
def word_DrilledShaft_TXT_Table(aDoc, concrete_grade, Pile_D, Pile_as0, Pile_p, Pile_H, Pile_Nk, Pile_Hk, Soil_vb, alpha_h, Vm, Vx, Rha, Ra, b0):
    if Rha >= Pile_Hk and Ra > Pile_Nk:
        if_Ra_Rha_satisfied = "满足"
    else:
        if_Ra_Rha_satisfied = "不满足"
    if Pile_D/1000 <= 1:
        Pile_b0 = f"0.9(1.5d+0.5)={b0}"
    else:
        Pile_b0 = f"0.9(d+1)={b0}"
    doc_str = [concrete_grade, Pile_D, Pile_as0, Pile_p, Pile_H, Pile_Nk, Pile_Hk, Soil_vb, Pile_b0, alpha_h, Vm, Vx, if_Ra_Rha_satisfied]
    repl_docstr = ["{concrete_grade}", "{Pile_D}", "{Pile_as0}", "{Pile_pg}", "{Pile_H}", "{Pile_Nk}", "{Pile_Hk}", "{Soil_vb}", "{Pile_b0}", "{alpha_h}", "{Rha_vm}", "{Rha_vx}", "{if_Ra_Rha_satisfied}"]
    for x,y in zip(doc_str, repl_docstr):
        repl_docstr_with_strlst(aDoc, y, str(x))

#=====================================================================================================================================================================

# 根据模型及对话框参数生成计算书中所需要的计算结果

# xyz是基础平面尺寸, gc是基础容重, h1为基础底标高, h2为软弱下卧层顶标高, gs为土容重, angle为扩散角
def foundation_pressure(x, y, z, gc, Fk1, Fk2, Mk, h1, h2, gs, angle, fa, faz, type_M):
    # Ak, 基础底面面积
    # Gk, 基础自重
    # pk, 基础底面处的平均压力值
    # Wx, 基础抗弯截面模量
    # pkmax, 基础底面边缘的最大压力值
    # pcz, 软弱下卧层顶面处土的自重压力值
    # pc, 基础底面处土的自重压力值
    # Az, 基础底面扩散影响面积
    # pz, 软弱下卧层顶面处的附加压力值
    # 初始化
    Ak, Gk, pk, Mk_e, Mk_a, Eccentricity, Wk, pkmax, pcz, pc, pz, if_satisfied_fa1, if_satisfied_fa2, if_satisfied_faz = [0, 0, 0, 0, 0, "", 0, 0, 0, 0, 0, "", "", ""]
    # Gk
    Gk = foundation_Gk(x, y, z, h1, gc, gs)
    # pk
    Ak, pk = foundation_pk(x, y, Gk, Fk1)
    if pk < fa:
        if_satisfied_fa1 = "满足"
    else:
        if_satisfied_fa1 = "不满足"
    # pkmax
    if Mk != 0:
        # 偏心检算
        Mk_e, Mk_a, Eccentricity = cal_Eccentricity(Gk, Fk2, Mk, x, y, type_M)
        if Eccentricity == "大偏心":
            pkmax = foundation_pkmax_Large_Eccentricity(Gk, Fk2, x, y, Mk_a, type_M)
            pkmin = 0
        elif Eccentricity == "小偏心":
            Wk, pkmax, pkmin = foundation_pkmax_pkmin(x, y, Gk, Fk2, Mk, type_M)
        if pkmax < 1.2 * fa:
            if_satisfied_fa2 = "满足" 
        else:
            if_satisfied_fa2 = "不满足"
    # pcz/pz
    if h2 != 0:
        pcz = foundation_pcz(h2, gs)
        pc, pz = foundation_pz(x, y, z, pk, h1, h2, gs, angle)
        if pcz + pz < faz:
            if_satisfied_faz = "满足"
        else:
            if_satisfied_faz = "不满足"
            
    return [Ak, Gk, pk, Mk_e, Mk_a, Eccentricity, Wk, pkmax, pkmin, pcz, pc, pz, if_satisfied_fa1, if_satisfied_fa2, if_satisfied_faz, type_M]


# 输入参数计算基础抗剪V、抗冲切F、抗弯M承载力及配筋
# 剪切计算主函数
def Vs_Verification(x, y, z, as0, at, bt, D, pmax, ft):
    # 是否需要抗剪验算
    Vs_lst = foundation_Vs(x, y, at, bt, D, pmax)
    # 抗剪切力
    Allowable_Vs_lst = foundation_Allowable_Vs(x, y, z, as0, ft)
    # 需要抗剪验算
    if Vs_lst != None:
        # print("基础底面短边尺寸小于等于柱宽加两倍基础有效高度，需要抗剪计算")
        # 剪切力
        As1, As2, Vs1, Vs2 = Vs_lst
        # 抗剪切力
        Bhs, A01, A02, Allowable_Vs1, Allowable_Vs2 = Allowable_Vs_lst
        print("剪切力: ", Vs1, Vs2)
        print("抗剪力: ", Allowable_Vs1, Allowable_Vs2)
        # 是否包络
        if Vs1 <= Allowable_Vs1:
            if_Vs_satisfied1 = "满足"
        else:
            if_Vs_satisfied1 = "不满足"
        if Vs2 <= Allowable_Vs2:
            if_Vs_satisfied2 = "满足"
        else:
            if_Vs_satisfied2 = "不满足"
    # else:
    #     print("基础底面短边尺寸大于柱宽加两倍基础有效高度，不需要抗剪计算")
    return [As1, As2, Vs1, Vs2, Bhs, A01, A02, Allowable_Vs1, Allowable_Vs2, if_Vs_satisfied1, if_Vs_satisfied2]


# 冲切计算主函数
def Fl_Verification(x, y, z, as0, at, bt, D, pmax, ft):
    # 抗冲切力
    h0, Bhp, am, Allowable_Fl = foundation_Allowable_Fl(ft, at, bt, z, D, as0)
    print("抗冲切力: ", Allowable_Fl)
    # 冲切力
    Al1, Al2, Fl1, Fl2 = foundation_Fl(pmax, at, bt, D, x, y, z, as0)
    print("冲切力: ", Fl1, Fl2)
    # 是否包络
    if Fl1 < Allowable_Fl:
        if_Fl_satisfied1 = "满足"
    else:
        if_Fl_satisfied1 = "不满足"
    if Fl2 < Allowable_Fl:
        if_Fl_satisfied2 = "满足"
    else:
        if_Fl_satisfied2 = "不满足"
    return [Al1, Al2, Fl1, Fl2, h0, Bhp, am, Allowable_Fl, if_Fl_satisfied1, if_Fl_satisfied2]


# 弯矩及配筋计算主函数
def M_Verification(x, y, z, as0, at, bt, D, p, pmax, pmin, fy, steel_s1, steel_s2, steel_d1, steel_d2, type_M):
    # 弯矩计算
    a1, aa, bb = cal_M_param(x, y, at, bt, D, type_M)
    print(a1, aa, bb)
    M1, M2 = cal_M(x, y, a1, aa, bb, p, pmax, pmin, type_M)
    print(M1, M2)
    # 配筋面积计算
    As1 = max([cal_As(M1, fy, x, y, z, as0, type_M), min_p_As(z, as0)])
    As2 = max([cal_As(M2, fy, x, y, z, as0, type_M), min_p_As(z, as0)])
    print("弯矩作用方向最小配筋面积: ", As1)# 弯矩方向
    print("垂直弯矩方向最小配筋面积: ", As2)# 垂直弯矩方向

    # 钢筋直径和间距计算
    rebar_arrangement_lst1 = cal_rebar_arrangement(z, as0, steel_s1, steel_s2, steel_d1, steel_d2, As1)
    rebar_arrangement_lst2 = cal_rebar_arrangement(z, as0, steel_s1, steel_s2, steel_d1, steel_d2, As2)
    if rebar_arrangement_lst1 != None: 
        s1, d1, p1, A1 = rebar_arrangement_lst1
        print(f"弯矩作用方向配筋规格: {d1}@{s1}mm", ", ", f"配筋率{round(p1*100,3)}% > 0.15%", ", ", f"钢筋截面面积{round(A1,1)}mm2 > {As1}mm2")
    else:
        print(f"弯矩作用方向配筋未成功")
    if rebar_arrangement_lst2 != None: 
        s2, d2, p2, A2 = rebar_arrangement_lst2
        print(f"垂直弯矩方向配筋规格: {d2}@{s2}mm", ", ", f"配筋率{round(p2*100,3)}% > 0.15%", ", ", f"钢筋截面面积{round(A2,1)}mm2 > {As2}mm2")
    else:
        print(f"垂直弯矩方向配筋未成功")
    
    if A1 >= As1:
        if_As_satisfied1 = "满足"
    else: 
        if_As_satisfied1 = "不满足"
    if s1 == 200:
        if_gouzao1 = "是"
    else:
        if_gouzao1 = "否"
    if A2 >= As2:
        if_As_satisfied2 = "满足"
    else: 
        if_As_satisfied2 = "不满足"
    if s2 == 200:
        if_gouzao2 = "是"
    else:
        if_gouzao2 = "否"

    return [a1, aa, bb, M1, M2, As1, As2, s1, d1, round(p1*100,3), round(A1,1), s2, d2, round(p2*100,3), round(A2,1), if_As_satisfied1, if_As_satisfied2, if_gouzao1, if_gouzao2]


# 钢管桩的竖向承载力, 返回所有对应公式需要的值
def JGJ94_Vertical_Capacity_steelpile(d, s_qsk, s_psk, s_layer_h):
    Qsk = JGJ94_Qsk_DrivenPile(d/1000, s_qsk, s_layer_h)
    Qpk = JGJ94_Qpk_steelpile(d/1000, s_psk[-1], s_layer_h[-1])
    Quk = JGJ94_Quk(Qsk, Qpk)
    Ra = JGJ94_Ra(Quk)
    return [Qsk, Qpk, Quk, Ra]


# 预制桩的竖向承载力, 返回所有对应公式需要的值
def JGJ94_Vertical_Capacity_precastpile(d, t, s_qsk, s_psk, s_layer_h):
    Qsk = JGJ94_Qsk_DrivenPile(d/1000, s_qsk, s_layer_h)
    Qpk = JGJ94_Qpk_precastpile(d/1000, t/1000, s_psk[-1], s_layer_h[-1])
    Quk = JGJ94_Quk(Qsk, Qpk)
    Ra = JGJ94_Ra(Quk)
    return [Qsk, Qpk, Quk, Ra]


# 混凝土灌注桩单桩竖向承载力总函数，返回所有对应公式需要的值
def JGJ94_Vertical_Capacity_concretepile(d, s_qsk, s_psk, s_layer_h, s_type):
    param_alst = Side_Resistance_Factor(s_type, d/1000) # 侧阻效应系数
    param_blst = Base_Resistance_Factor(s_type, d/1000) # 端阻效应系数
    Qsk = JGJ94_Qsk_DrilledShaft(d/1000, s_qsk, s_layer_h, param_alst)
    Qpk = JGJ94_Qpk_DrilledShaft(d/1000, s_psk[-1], param_blst)
    Quk = JGJ94_Quk(Qsk, Qpk)
    Ra = JGJ94_Ra(Quk)
    return [Qsk, Qpk, Quk, Ra]


# 钢桩的水平承载力总函数，返回所有对应公式需要的值
def JGJ94_Lateral_Capacity_steelpile(d, t, h, friction, Cohesion, vb):
    m_n = 1
    E = 206*(10**6) # 单位kPa
    I = math.pi*((d/1000)**4-((d/1000)-2*(t/1000))**4) / 64 # 单位m
    if d/1000 > 1:
        b0 = 0.9*(d/1000 + 1)
    else:
        b0 = 0.9*(1.5*d/1000 + 0.5)
    m_lst = Subgrade_Reaction_Coefficient(friction, Cohesion, vb) # 所有土层的m
    m = m_lst[m_n-1] # 计算土层的m
    phi = friction[m_n-1] # 计算土层的内摩擦角
    c = Cohesion[m_n-1] # 计算土层的黏聚力
    alpha = Displacement_Coefficient(m, b0, E, I) # 单位 1/m
    Vm, Vx, alpha_h = JGJ94_Table_Vm_Vs(alpha, h) # 系数及换算长度, 单位m
    Rha = JGJ94_Rha2(alpha, E, I, Vx, vb) # 单位kN
    return [phi, c, m, b0, round(E*I, 1), alpha, alpha_h, Vm, Vx, Rha]


# 预制桩的水平承载力总函数，返回所有对应公式需要的值
def JGJ94_Lateral_Capacity_precastpile(concrete_grade, d, t, h, friction, Cohesion, vb):
    m_n = 1
    Ec = concrete_f(concrete_grade)[2] * 1000 # 规范单位为Mpa
    E = 0.85 * Ec # 单位kPa
    I = math.pi*((d/1000)**4-((d/1000)-2*(t/1000))**4) / 64 # 单位m
    if d/1000 > 1:
        b0 = 0.9*(d/1000 + 1)
    else:
        b0 = 0.9*(1.5*d/1000 + 0.5)
    m_lst = Subgrade_Reaction_Coefficient(friction, Cohesion, vb) # 所有土层的m
    m = m_lst[m_n-1] # 计算土层的m
    phi = friction[m_n-1] # 计算土层的内摩擦角
    c = Cohesion[m_n-1] # 计算土层的黏聚力
    alpha = Displacement_Coefficient(m, b0, E, I) # 单位 1/m
    Vm, Vx, alpha_h = JGJ94_Table_Vm_Vs(alpha, h) # 系数及换算长度, 单位m
    Rha = JGJ94_Rha2(alpha, E, I, Vx, vb) # 单位kN
    return [phi, c, m, b0, round(E*I, 1), alpha, alpha_h, Vm, Vx, Rha]


# 钢筋混凝土灌注桩的水平承载力总函数，返回所有对应公式需要的值
def JGJ94_Lateral_Capacity_concretepile(concrete_grade, rebar_grade, d, pg, d_as, N, h, friction, Cohesion, vb):
    m_n = 1
    fc, ft, Ec = concrete_f(concrete_grade)
    fy, Es = rebar_f(rebar_grade)

    if d/1000 > 1:
        b0 = 0.9*(d/1000 + 1) # 计算宽度
    else:
        b0 = 0.9*(1.5*d/1000 + 0.5) # 计算宽度

    m_lst = Subgrade_Reaction_Coefficient(friction, Cohesion, vb) # 所有土层的m
    m = m_lst[m_n-1] # 计算土层的m
    phi = friction[m_n-1] # 计算土层的内摩擦角
    c = Cohesion[m_n-1] # 计算土层的黏聚力

    W0 = Section_Modulus_of_Transformed_Section_at_Tension_Edge_of_Pile_Shaft(d, Ec, Es, pg, d_as) # 单位mm3
    An = Transformed_Section_Area_of_Pile_Shaft(d, Ec, Es, pg) # 单位mm2
    E = 0.85 * Ec * 1000 # 单位kPa
    I = (W0*10**-9) * ((d-2*d_as)/1000) / 2 # 单位m
    alpha = Displacement_Coefficient(m, b0, E, I) # 单位 1/m
    Vm, Vx, alpha_h = JGJ94_Table_Vm_Vs(alpha, h) # 系数及换算长度, 单位m

    if pg < 0.65:
        Rha = JGJ94_Rha1(alpha, 2, ft, W0, Vm, pg, N, An) # 单位kN
    else:
        Rha = JGJ94_Rha2(alpha, E, I, Vx, vb) # 单位kN
    
    return [phi, c, m, b0, round(E*I, 1), alpha, alpha_h, Vm, Vx, Es, Ec, W0, An, Rha]

#=====================================================================================================================================================================

# 公式替换主函数

# 根据参数计算结果调用对应函数替换word中的公式
# 钢管桩
def cal_SteelPile_stability(word, aDoc, Fx, Mm, L, Area, Ix, ix, Phi, Wx, quote_lambda, Nex, sigma):
    
    # 替换word公式
    # word_steelpile(word, aDoc, Fx, round(Phi,3), round(Area,1), round(Mm,1), round(Wx,1), round(Nex,1), round(quote_lambda,1), round(sigma,1))
    search_range = aDoc.Content
    search_range.Select()
    selection = word.Selection
    # LaTeX公式
    # latex_list1, latex_list2 = SteelPile_stability_LaTeX(Fx, round(Phi,3), round(Area,1), round(Mm,1), round(Wx,1), round(Nex,1), round(quote_lambda,1), round(sigma,1))
    latex_list1 = GB50017_2017_8_2_1_2 + r'=\frac{\pi^2\times206\times{10}^3\times' + str(round(Area,1)) + r'}{1.1\times{' + str(round(quote_lambda,1)) + r'}^2}=' + str(round(Nex,1)) + 'N'
    latex_list2 = GB50017_2017_8_2_1_1 + r'=\frac{' + str(Fx) + r'\times{10}^3}{' + str(round(Phi,3)) + r'\times' + str(round(Area,1)) + r'}+\frac{1\times' + str(round(Mm,1)) + r'\times{10}^6}{1.15\times' + str(round(Wx,1)) + r'\times\left(1-0.8\times\frac{' + str(Fx) + r'\times{10}^3}{' + str(round(Nex,1)) + r'}\right)}=' + str(round(sigma,1)) + 'MPa'
    # 公式替换
    repl_docstr_with_omath(aDoc,selection,latex_list1,'[Nex]')
    repl_docstr_with_omath(aDoc,selection,latex_list2,'[SigmaNex]')
    # print("参数", Area, Ix, ix, L)
    print("面积(mm2): ", round(Area, 3))
    print("惯性矩(mm4): ", round(Ix, 3))
    print("回转半径(mm): ", round(ix, 3))
    print("计算长度(mm)", round(L, 3))
    print("长细比: ", round(quote_lambda, 3))


# 扩大基础
# xyz是基础平面尺寸, gc是基础容重, h1为基础底标高, h2为软弱下卧层顶标高, gs为土容重, angle为扩散角   
def cal_foundation_contact_stress(word, aDoc, x, y, z, gc, Fk1, Fk2, My, h1, h2, gs, angle, fa, faz, Ak, Gk, pk, My_e, My_a, Ecc, Wx, pkmax, pkmin, pcz, pc, pz, if_satisfied_fa1, if_satisfied_fa2, if_satisfied_faz):
    # Ak, Gk, pk, My_e, My_a, Eccentricity, Wx, pkmax, pcz, pc, pz, if_satisfied_fa1, if_satisfied_fa2, if_satisfied_faz = foundation_pressure(x, y, z, gc, Fk, My, h1, h2, gs, angle, fa, faz)
    # 无软弱下卧层
    if if_satisfied_faz == "":
        # 轴心受压
        if if_satisfied_fa2 == "":
            # print("轴心荷载作用，无软弱下卧层")
            # 调用轴心受压的单独计算书模板
            word_pk(word, aDoc, x, y, z, gc, h1, gs, Fk1, fa, Ak, Gk, pk, if_satisfied_fa1)
        # 偏心受压
        else:
            # print("偏心荷载作用，无软弱下卧层")
            # 调用偏心受压的单独计算书模板
            word_pk(word, aDoc, x, y, z, gc, h1, gs, Fk1, fa, Ak, Gk, pk, if_satisfied_fa1)
            word_pkmax(word, aDoc, Fk2, My, fa, Ak, Gk, My_e, My_a, x, y, Wx, pkmax, pkmin, if_satisfied_fa2)
    # 有软弱下卧层
    else:
        if if_satisfied_fa2 == "":
            # print("轴心荷载作用，有软弱下卧层")
            # 调用轴心受压且有软弱下卧层的单独计算书模板
            word_pk(word, aDoc, x, y, z, gc, h1, gs, Fk1, fa, Ak, Gk, pk, if_satisfied_fa1)
            word_pz(word, aDoc, z, h1, h2, gs, faz, pcz, pc, pz, if_satisfied_faz)
        else:  
            # print("偏心荷载作用，有软弱下卧层")
            # 调用偏心且有软弱下卧层的单独计算书模板
            word_pk(word, aDoc, x, y, z, gc, h1, gs, Fk1, fa, Ak, Gk, pk, if_satisfied_fa1)
            word_pkmax(word, aDoc, Fk2, My, fa, Ak, Gk, My_e, My_a, x, y, Wx, pkmax, pkmin, if_satisfied_fa2)
            word_pz(word, aDoc, z, h1, h2, gs, faz, pcz, pc, pz, if_satisfied_faz)

#=====================================================================================================================================================================

# 计算书前处理
def Pre_JiSuanShu(word, aDoc, bDoc, txt_path, foundation_standard, foundation_value_dict):
    # 从txt文本中读取的信息
    result_lst = read_txt_found_line(txt_path, FEA_To_Cal_txt_str_lookfor())
    # 规范代号对照库
    Standard_Codename_dict = Standard_Codename()
    # 风荷载规范
    Wind_Standard_Codename = result_lst[0][0]
    Wind_Standard = Standard_Codename_dict[Wind_Standard_Codename]
    # 水流力规范
    Flow_Standard_Codename = result_lst[1][0]
    Flow_Standard = Standard_Codename_dict[Flow_Standard_Codename]
    # 基础规范(计算书UI获取)
    Foundation_Standard_Codename = foundation_standard
    Foundation_Standard = Standard_Codename_dict[Foundation_Standard_Codename]
    # 自重参数
    Model_Gz = result_lst[2][0]
    # 混凝土荷载
    Concrete_Gs = result_lst[3][0]
    # 模板荷载
    Formwork_Load = result_lst[4][0]
    # 施工荷载
    Construction_Load = result_lst[5][0]
    # 风荷载参数
    Wind_param_lst = result_lst[6]
    if Wind_param_lst != ['None']:
        Wind_param_lst[6] = Wind_param_lst[6].split(":")[0]
    # 水流力参数
    Flow_param_lst = result_lst[7]
    # 材料特性参数
    Material_Properties_dict = {
        "Bailey"   : [], # 贝雷
        "Steel"    : [], # 钢材
        "Concrete" : [], # 混凝土
        "Rebar"    : [], # 钢筋
        "Bolt"     : [], # 螺栓
        "Weld"     : [], # 焊缝
        "ZhuJB"    : [], # 竹胶板
        "Fangmu"   : [], # 方木
    }
    # print(Material_Properties_dict.keys())
    Material_txt_lst = result_lst[8:16]
    # 合并模型中定义的和计算书中定义的混凝土和钢筋的信息
    Material_in_CalUI_lst_1 = foundation_value_dict['浅基础'] # 浅基础信息
    Material_in_CalUI_lst_2 = foundation_value_dict['深基础'] # 浅基础信息
    if Material_in_CalUI_lst_1 != []:
        Concrete_Material_1 = Material_in_CalUI_lst_1[12]
        Rebar_Material_1 = Material_in_CalUI_lst_1[13]
    else:
        Concrete_Material_1 = 'None'
        Rebar_Material_1 = 'None'
    if Material_in_CalUI_lst_2 != []:
        Concrete_Material_2 = Material_in_CalUI_lst_2[1]
        Rebar_Material_2 = Material_in_CalUI_lst_2[2]
        if type(Concrete_Material_2) != str:
            Concrete_Material_2 = 'None'
        if type(Rebar_Material_2) != str:
            Rebar_Material_2 = 'None'
    else:
        Concrete_Material_2 = 'None'
        Rebar_Material_2 = 'None'
    Material_txt_lst[2] = list(set(tuple([item for item in Material_txt_lst[2] + [Concrete_Material_1, Concrete_Material_2] if type(item) == str and item != 'None'])))
    Material_txt_lst[3] = list(set(tuple([item for item in Material_txt_lst[3] + [Rebar_Material_1, Rebar_Material_2] if type(item) == str and item != 'None'])))
    # 对 Material_Properties_dict 字典赋值
    print(Material_txt_lst)
    for i, key in enumerate(Material_Properties_dict.keys()):
        if Material_txt_lst[i] == [] or Material_txt_lst[i] == ['None']:
            Material_Properties_dict[key] = []
        else:
            for Material in Material_txt_lst[i]:
                if key == "Concrete":
                    Material_Properties_dict[key].append([Material] + concrete_f(Material))
                elif key == "Rebar":
                    Material_Properties_dict[key].append([Material] + rebar_f(Material))
                else:
                    # 其余的还没做材料特性定义函数
                    pass
    # 计算依据文本替换
    Calculation_Basis_TxT(word, aDoc, bDoc, Wind_Standard, Flow_Standard, Foundation_Standard)
    # 设计参数文本替换
    Load_Parameter_TxT(aDoc, bDoc, Model_Gz, Concrete_Gs, Formwork_Load, Construction_Load, Wind_param_lst, Flow_param_lst)
    # 设计参数公式替换
    Load_Parameter_Fuc(word, aDoc, bDoc, Wind_Standard_Codename, Flow_Standard_Codename, Wind_param_lst, Flow_param_lst)
    # 材料特性文本替换
    Material_Properties_TxT(word, aDoc, bDoc, Material_Properties_dict)
    # 工况分析文本替换
    Case_Analysis_TxT(aDoc, Flow_Standard)
    # 荷载组合文本替换
    Load_Combination_TxT(word, aDoc, bDoc, Flow_Standard)
    print("cal_doc_pre 成功生成")


def Soleplate_JiSuanShu(wdApp, aDoc, docx_open_path, docx_save_path, soleplate_values):
    cDoc = open_doc_saveas_adoc(wdApp, 'falsework_soleplate_template.docx', 'falsework_soleplate_template_saveas.docx', docx_open_path, docx_save_path) # 支架底模板
    # 计算参数
    ZhuJiaoBan_thickness, FangMu_SEC, Concrete_Cal_Height1, Concrete_Cal_Height2, FangMu_Gap1, FangMu_Gap2, Rib_SEC, Rib_Gap, Concrete_Unit, Construction_Load, Vibration_Load, Formwork_Load = soleplate_values['模板']
    FangMu_D, FangMu_H = [int(x) for x in  re.findall(r'\d+', FangMu_SEC)] # 方木长宽
    # 文本替换
    # 竹胶板1
    repl_docstr = ['{ZhuJiaoBan1_thickness}', '{ZhuJiaoBan1_FangMu_Gap}', '{ZhuJiaoBan1_FangMu_Sec}', '{ZhuJiaoBan1_FangMu_JGap}', '{ZhuJiaoBan1_Concrete_Cal_Height}']
    repl_valstr = [ZhuJiaoBan_thickness, FangMu_Gap1, FangMu_SEC, FangMu_Gap1-FangMu_D, Concrete_Cal_Height1]
    [repl_docstr_with_strlst(cDoc, y, str(x)) for x, y in zip(repl_valstr, repl_docstr)]
    # 方木1
    repl_docstr = ['{FangMu1_Sec}', '{FangMu1_Gap}', '{FangMu1_DaLe_Gap}', '{FangMu1_Concrete_Cal_Height}']
    repl_valstr = [FangMu_SEC, FangMu_Gap1, Rib_Gap, Concrete_Cal_Height1]
    [repl_docstr_with_strlst(cDoc, y, str(x)) for x, y in zip(repl_valstr, repl_docstr)]
    # 底模板1
    repl_docstr = ['{DM1_ZhuJiaoBan_thickness}', '{DM1_FangMu_Sec}', '{DM1_FangMu_Gap}', '{DM1_DaLe_SEC}', '{DM1_DaLe_Gap}']
    repl_valstr = [ZhuJiaoBan_thickness, FangMu_SEC, FangMu_Gap1, Rib_SEC, Rib_Gap]
    [repl_docstr_with_strlst(cDoc, y, str(x)) for x, y in zip(repl_valstr, repl_docstr)]
    if FangMu_Gap1 != FangMu_Gap2:
        # 竹胶板2
        repl_docstr = ['{ZhuJiaoBan2_thickness}', '{ZhuJiaoBan2_FangMu_Gap}', '{ZhuJiaoBan2_FangMu_Sec}', '{ZhuJiaoBan2_FangMu_JGap}', '{ZhuJiaoBan2_Concrete_Cal_Height}']
        repl_valstr = [ZhuJiaoBan_thickness, FangMu_Gap2, FangMu_SEC, FangMu_Gap2-FangMu_D, Concrete_Cal_Height2]
        [repl_docstr_with_strlst(cDoc, y, str(x)) for x, y in zip(repl_valstr, repl_docstr)]
        # 方木2
        repl_docstr = ['{FangMu2_Sec}', '{FangMu2_Gap}', '{FangMu2_DaLe_Gap}', '{FangMu2_Concrete_Cal_Height}']
        repl_valstr = [FangMu_SEC, FangMu_Gap2, Rib_Gap, Concrete_Cal_Height2]
        [repl_docstr_with_strlst(cDoc, y, str(x)) for x, y in zip(repl_valstr, repl_docstr)]
        # 底模板2
        repl_docstr = ['{DM2_ZhuJiaoBan_thickness}', '{DM2_FangMu_Sec}', '{DM2_FangMu_Gap1}', '{DM2_FangMu_Gap2}', '{DM2_DaLe_SEC}', '{DM2_DaLe_Gap}']
        repl_valstr = [ZhuJiaoBan_thickness, FangMu_SEC, FangMu_Gap1, FangMu_Gap2, Rib_SEC, Rib_Gap]
        [repl_docstr_with_strlst(cDoc, y, str(x)) for x, y in zip(repl_valstr, repl_docstr)]
    # 公式替换
    selection = wdApp.Selection
    # 公式
    latex_I     = r'I=\frac{bh^3}{12}'
    latex_W     = r'W=\frac{bh^2}{6}'
    latex_M     = r'M=0.1ql^2'    
    latex_sigma = r'\sigma=\frac{M}{W}'
    latex_f     = r'f=0.677\times\frac{ql^4}{100EI}'
    # 计算结果
    zjb_I1, zjb_I2 = scientific_notation((1000*pow(ZhuJiaoBan_thickness,3))/12) # 竹胶板的I
    zjb_W1, zjb_W2 = scientific_notation((1000*pow(ZhuJiaoBan_thickness,2))/6) # 竹胶板的W
    zjb_jbq1 = round(1.4*Concrete_Unit*Concrete_Cal_Height1+1.2*Formwork_Load+1.4*0.7*(Construction_Load+Vibration_Load),1) # 腹板基本组合下的q
    zjb_jbM1 = round(0.1*zjb_jbq1*pow(FangMu_Gap1/1000,2),2) # 腹板基本组合下的M
    zjb_jbsigma1 = round((zjb_jbM1*1000000)/(zjb_W1*pow(10,zjb_W2)),1) # 腹板基本组合下的应力
    zjb_bzq1 = round(Concrete_Unit*Concrete_Cal_Height1+Formwork_Load+0.7*(Construction_Load+Vibration_Load),1) # 腹板标准组合下的q
    zjb_bzf1 = round(0.677*(zjb_bzq1*pow((FangMu_Gap1-FangMu_D),4))/(100*5000*zjb_I1*pow(10,zjb_I2)),2) # 腹板标准组合下的f
    zjb_bzf1_rx = round((FangMu_Gap1-FangMu_D)/400,2) # 变形容许值
    # 竹胶板1
    fm_I1, fm_I2 = scientific_notation((FangMu_D*pow(FangMu_H,3))/12) # 方木的I
    fm_W1, fm_W2 = scientific_notation((FangMu_D*pow(FangMu_H,2))/6) # 方木的W
    fm_jbq1 = round(zjb_jbq1*FangMu_Gap1/1000,1) # 腹板基本组合下的q
    fm_jbM1 = round(0.1*fm_jbq1*pow(Rib_Gap/1000,2),2) # 腹板基本组合下的M
    fm_jbsigma1 = round((fm_jbM1*1000000)/(fm_W1*pow(10,fm_W2)),1) # 腹板基本组合下的应力
    fm_bzq1 = round(zjb_bzq1*FangMu_Gap1/1000,1) # 腹板标准组合下的q
    fm_bzf1 = round(0.677*(fm_bzq1*pow(Rib_Gap,4))/(100*9000*fm_I1*pow(10,fm_I2)),2) # 腹板标准组合下的f
    fm_bzf1_rx = round(Rib_Gap/400,2) # 变形容许值
    # 替换
    latex1 = latex_I + r'=\frac{1000\times' + str(ZhuJiaoBan_thickness) + r'^3}{12}=' + str(zjb_I1) + r'\times{10}^' + str(zjb_I2) + 'mm^4'
    repl_docstr_with_omath(cDoc, selection, latex1, '{ZhuJiaoBan1_I}')  
    latex2 = latex_W + r'=\frac{1000\times' + str(ZhuJiaoBan_thickness) + r'^2}{6}=' + str(zjb_W1) + r'\times{10}^' + str(zjb_W2) + r'{\rm mm}^3'
    repl_docstr_with_omath(cDoc, selection, latex2, '{ZhuJiaoBan1_W}')
    latex3 = r'q=1.4\times' + str(Concrete_Unit) + r'\times' + str(Concrete_Cal_Height1) + r'+1.2\times' + str(Formwork_Load) + r'+1.4\times0.7\times\left(' + str(Construction_Load) + '+' + str(Vibration_Load) + r'\right)=' + str(zjb_jbq1) + 'kN/m' 
    repl_docstr_with_omath(cDoc, selection, latex3, '{ZhuJiaoBan1_JBq}')
    latex4 = latex_M + r'=0.1\times' + str(zjb_jbq1) + r'\times{' + str(round(FangMu_Gap1/1000,1)) + r'}^2=' + str(zjb_jbM1) + r'kN\bullet m' 
    repl_docstr_with_omath(cDoc, selection, latex4, '{ZhuJiaoBan1_JBM}')
    latex5 = latex_sigma + r'=\frac{' + str(zjb_jbM1) + r'\times{10}^6}{' + str(zjb_W1) + r'\times{10}^' + str(zjb_W2) + '}=' + str(zjb_jbsigma1) + r'Mpa<f_m=30\mathrm{MPa}' 
    repl_docstr_with_omath(cDoc, selection, latex5, '{ZhuJiaoBan1_JBsigma}')
    latex6 = 'q=' + str(Concrete_Unit) + r'\times' + str(Concrete_Cal_Height1) + '+' + str(Formwork_Load) + r'+0.7\times\left(' + str(Construction_Load) + '+' + str(Vibration_Load) + r'\right)=' + str(zjb_bzq1) + 'kN/m' 
    repl_docstr_with_omath(cDoc, selection, latex6, '{ZhuJiaoBan1_BZq}')
    latex7 = latex_f + r'=0.677\times\frac{\mathrm{' + str(zjb_bzq1) + r'}\ \ \times{' + str(FangMu_Gap1-FangMu_D) + r'}^4}{100\times5\times{10}^3\times' + str(zjb_I1) + r'\times{10}^' + str(zjb_I2) + '}=' + str(zjb_bzf1) + r'\mathrm{\mathrm{mm<}}\frac{\mathrm{' + str(FangMu_Gap1-FangMu_D) + r'}}{\mathrm{400}}\mathrm{=' + str(zjb_bzf1_rx) + 'mm}' 
    repl_docstr_with_omath(cDoc, selection, latex7, '{ZhuJiaoBan1_BZf}')
    # 方木1
    latex1 = latex_I + r'=\frac{' + str(FangMu_D) + r'\times' + str(FangMu_H) + r'^3}{12}=' + str(fm_I1) + r'\times{10}^' + str(fm_I2) + 'mm^4'
    repl_docstr_with_omath(cDoc, selection, latex1, '{FangMu1_I}')
    latex2 = latex_W + r'=\frac{' + str(FangMu_D) + r'\times' + str(FangMu_H) + r'^2}{6}=' + str(fm_W1) + r'\times{10}^' + str(fm_W2) + r'{\rm mm}^3'
    repl_docstr_with_omath(cDoc, selection, latex2, '{FangMu1_W}')
    latex3 = r'q=\left(1.4\times' + str(Concrete_Unit) + r'\times' + str(Concrete_Cal_Height1) + r'+1.2\times' + str(Formwork_Load) + r'+1.4\times0.7\times\left(' + str(Construction_Load)  + '+' + str(Vibration_Load) + r'\right)\right)\times' + str(round(FangMu_Gap1/1000,1)) + '=' + str(fm_jbq1) + 'kN/m'
    repl_docstr_with_omath(cDoc, selection, latex3, '{FangMu1_JBq}')
    latex4 =latex_M + r'=0.1\times' + str(fm_jbq1) + r'\times{' + str(round(Rib_Gap/1000,1)) + r'}^2=' + str(fm_jbM1) + r'kN\bullet m' 
    repl_docstr_with_omath(cDoc, selection, latex4, '{FangMu1_JBM}')
    latex5 = latex_sigma + r'=\frac{' + str(fm_jbM1) + r'\times{10}^6}{' + str(fm_W1) + r'\times{10}^' + str(fm_W2) + '}=' + str(fm_jbsigma1) + r'Mpa<f_m=10\mathrm{MPa}' 
    repl_docstr_with_omath(cDoc, selection, latex5, '{FangMu1_JBsigma}')
    latex6 = r'q=\left(' + str(Concrete_Unit) + r'\times' + str(Concrete_Cal_Height1) + '+' + str(Formwork_Load) + r'+0.7\times\left(' + str(Construction_Load) + '+' + str(Vibration_Load) + r'\right)\right)\times' + str(round(FangMu_Gap1/1000,1)) + '=' + str(fm_bzq1) + 'kN/m'
    repl_docstr_with_omath(cDoc, selection, latex6, '{FangMu1_BZq}')
    latex7 = latex_f + r'=0.677\times\frac{\mathrm{' + str(fm_bzq1) + r'}\ \ \times{' + str(Rib_Gap) + r'}^4}{100\times5\times{10}^3\times' + str(fm_I1) + r'\times{10}^' + str(fm_I2) + '}=' + str(fm_bzf1) + r'\mathrm{\mathrm{mm<}}\frac{\mathrm{' + str(Rib_Gap) + r'}}{\mathrm{400}}\mathrm{=' + str(fm_bzf1_rx) + 'mm}'
    repl_docstr_with_omath(cDoc, selection, latex7, '{FangMu1_BZf}')
    if FangMu_Gap1 != FangMu_Gap2:
        # 竹胶板
        zjb_jbq2 = round(1.4*Concrete_Unit*Concrete_Cal_Height2+1.2*Formwork_Load+1.4*0.7*(Construction_Load+Vibration_Load),1) # 顶底板基本组合下的q
        zjb_jbM2 = round(0.1*zjb_jbq2*pow(FangMu_Gap2/1000,2),2) # 顶底板基本组合下的M
        zjb_jbsigma2 = round((zjb_jbM2*1000000)/(zjb_W1*pow(10,zjb_W2)),1) # 腹板基本组合下的应力
        zjb_bzq2 = round(Concrete_Unit*Concrete_Cal_Height2+Formwork_Load+0.7*(Construction_Load+Vibration_Load),1) # 顶底板标准组合下的q
        zjb_bzf2 = round(0.677*(zjb_bzq2*pow((FangMu_Gap2-FangMu_D),4))/(100*5000*zjb_I1*pow(10,zjb_I2)),2) # 腹板标准组合下的f
        zjb_bzf2_rx = round((FangMu_Gap2-FangMu_D)/400,2) # 变形容许值
        # 方木
        fm_jbq2 = round(zjb_jbq2*FangMu_Gap2/1000,1)# 顶底板基本组合下的q
        fm_jbM2 = round(0.1*fm_jbq2*pow(Rib_Gap/1000,2),2) # 顶底板基本组合下的M
        fm_jbsigma2 = round((fm_jbM2*1000000)/(fm_W1*pow(10,fm_W2)),1) # 腹板基本组合下的应力
        fm_bzq2 = round(zjb_bzq2*FangMu_Gap2/1000,1) # 顶底板标准组合下的q
        fm_bzf2 = round(0.677*(fm_bzq2*pow(Rib_Gap,4))/(100*9000*fm_I1*pow(10,fm_I2)),2) # 腹板标准组合下的f
        fm_bzf2_rx = round(Rib_Gap/400,2) # 变形容许值
        # 竹胶板2
        latex1 = latex_I + r'=\frac{1000\times' + str(ZhuJiaoBan_thickness) + r'^3}{12}=' + str(zjb_I1) + r'\times{10}^' + str(zjb_I2) + 'mm^4'
        repl_docstr_with_omath(cDoc, selection, latex1, '{ZhuJiaoBan2_I}')
        latex2 = latex_W + r'=\frac{1000\times' + str(ZhuJiaoBan_thickness) + r'^2}{6}=' + str(zjb_W1) + r'\times{10}^' + str(zjb_W2) + r'{\rm mm}^3'
        repl_docstr_with_omath(cDoc, selection, latex2, '{ZhuJiaoBan2_W}')
        latex3 = r'q=1.4\times' + str(Concrete_Unit) + r'\times' + str(Concrete_Cal_Height2) + r'+1.2\times' + str(Formwork_Load) + r'+1.4\times0.7\times\left(' + str(Construction_Load) + '+' + str(Vibration_Load) + r'\right)=' + str(zjb_jbq2) + 'kN/m'
        repl_docstr_with_omath(cDoc, selection, latex3, '{ZhuJiaoBan2_JBq}')
        latex4 = latex_M + r'=0.1\times' + str(zjb_jbq2) + r'\times{' + str(round(FangMu_Gap2/1000,1)) + r'}^2=' + str(zjb_jbM2) + r'kN\bullet m' 
        repl_docstr_with_omath(cDoc, selection, latex4, '{ZhuJiaoBan2_JBM}')
        latex5 = latex_sigma + r'=\frac{' + str(zjb_jbM2) + r'\times{10}^6}{' + str(zjb_W1) + r'\times{10}^' + str(zjb_W2) + '}=' + str(zjb_jbsigma2) + r'Mpa<f_m=30\mathrm{MPa}'
        repl_docstr_with_omath(cDoc, selection, latex5, '{ZhuJiaoBan2_JBsigma}')
        latex6 = 'q=' + str(Concrete_Unit) + r'\times' + str(Concrete_Cal_Height2) + '+' + str(Formwork_Load) + r'+0.7\times\left(' + str(Construction_Load) + '+' + str(Vibration_Load) + r'\right)=' + str(zjb_bzq2) + 'kN/m' 
        repl_docstr_with_omath(cDoc, selection, latex6, '{ZhuJiaoBan2_BZq}')
        latex7 = latex_f + r'=0.677\times\frac{\mathrm{' + str(zjb_bzq2) + r'}\ \ \times{' + str(FangMu_Gap2-FangMu_D) + r'}^4}{100\times5\times{10}^3\times' + str(zjb_I1) + r'\times{10}^' + str(zjb_I2) + '}=' + str(zjb_bzf2) + r'\mathrm{\mathrm{mm<}}\frac{\mathrm{' + str(FangMu_Gap2-FangMu_D) + r'}}{\mathrm{400}}\mathrm{=' + str(zjb_bzf2_rx) + 'mm}' 
        repl_docstr_with_omath(cDoc, selection, latex7, '{ZhuJiaoBan2_BZf}')
        # 方木2
        latex1 = latex_I + r'=\frac{' + str(FangMu_D) + r'\times' + str(FangMu_H) + r'^3}{12}=' + str(fm_I1) + r'\times{10}^' + str(fm_I2) + 'mm^4'
        repl_docstr_with_omath(cDoc, selection, latex1, '{FangMu2_I}')
        latex2 = latex_W + r'=\frac{100\times100^2}{6}=1.67\times{10}^5{\rm mm}^3'
        repl_docstr_with_omath(cDoc, selection, latex2, '{FangMu2_W}')
        latex3 = r'q=\left(1.4\times' + str(Concrete_Unit) + r'\times' + str(Concrete_Cal_Height2) + r'+1.2\times' + str(Formwork_Load) + r'+1.4\times0.7\times\left(' + str(Construction_Load)  + '+' + str(Vibration_Load) + r'\right)\right)\times' + str(round(FangMu_Gap2/1000,1)) + '=' + str(fm_jbq2) + 'kN/m'
        repl_docstr_with_omath(cDoc, selection, latex3, '{FangMu2_JBq}')
        latex4 = latex_M + r'=0.1\times' + str(fm_jbq2) + r'\times{' + str(round(Rib_Gap/1000,1)) + r'}^2=' + str(fm_jbM2) + r'kN\bullet m' 
        repl_docstr_with_omath(cDoc, selection, latex4, '{FangMu2_JBM}')
        latex5 = latex_sigma + r'=\frac{' + str(fm_jbM2) + r'\times{10}^6}{' + str(fm_W1) + r'\times{10}^' + str(fm_W2) + '}=' + str(fm_jbsigma2) + r'Mpa<f_m=10\mathrm{MPa}' 
        repl_docstr_with_omath(cDoc, selection, latex5, '{FangMu2_JBsigma}')
        latex6 = r'q=\left(' + str(Concrete_Unit) + r'\times' + str(Concrete_Cal_Height2) + '+' + str(Formwork_Load) + r'+0.7\times\left(' + str(Construction_Load) + '+' + str(Vibration_Load) + r'\right)\right)\times' + str(round(FangMu_Gap2/1000,1)) + '=' + str(fm_bzq2) + 'kN/m'
        repl_docstr_with_omath(cDoc, selection, latex6, '{FangMu2_BZq}')
        latex7 = latex_f + r'=0.677\times\frac{\mathrm{' + str(fm_bzq2) + r'}\ \ \times{' + str(Rib_Gap) + r'}^4}{100\times5\times{10}^3\times' + str(fm_I1) + r'\times{10}^' + str(fm_I2) + '}=' + str(fm_bzf2) + r'\mathrm{\mathrm{mm<}}\frac{\mathrm{' + str(Rib_Gap) + r'}}{\mathrm{400}}\mathrm{=' + str(fm_bzf2_rx) + 'mm}'
        repl_docstr_with_omath(cDoc, selection, latex7, '{FangMu2_BZf}')  
    if FangMu_Gap1 == FangMu_Gap2:
        WordA_PasteTo_WordB_DeleteLines(wdApp, cDoc, cDoc, '{ZhuJiaoBan1_Start}', '{ZhuJiaoBan1_End}', '{DM1_ZhuJiaoBan1_Paste_Here}')
        WordA_PasteTo_WordB_DeleteLines(wdApp, cDoc, cDoc, '{FangMu1_Start}', '{FangMu1_End}', '{DM1_FangMu1_Paste_Here}')
    else:
        WordA_PasteTo_WordB_DeleteLines(wdApp, cDoc, cDoc, '{ZhuJiaoBan1_Start}', '{ZhuJiaoBan1_End}', '{DM2_ZhuJiaoBan1_Paste_Here}')
        WordA_PasteTo_WordB_DeleteLines(wdApp, cDoc, cDoc, '{ZhuJiaoBan2_Start}', '{ZhuJiaoBan2_End}', '{DM2_ZhuJiaoBan2_Paste_Here}')
        WordA_PasteTo_WordB_DeleteLines(wdApp, cDoc, cDoc, '{FangMu1_Start}', '{FangMu1_End}', '{DM2_FangMu1_Paste_Here}')
        WordA_PasteTo_WordB_DeleteLines(wdApp, cDoc, cDoc, '{FangMu2_Start}', '{FangMu2_End}', '{DM2_FangMu2_Paste_Here}')
    # 粘贴到总计算书
    if FangMu_Gap1 == FangMu_Gap2:
        WordA_PasteTo_WordB_DeleteLines(wdApp, cDoc, aDoc, '{DiMoBan1_Start}', '{DiMoBan1_End}', '{DMB_Paste_Here}')
    else:
        WordA_PasteTo_WordB_DeleteLines(wdApp, cDoc, aDoc, '{DiMoBan2_Start}', '{DiMoBan2_End}', '{DMB_Paste_Here}')
    adoc_save_close(cDoc) # 保存
    print('falsework_soleplate_template_saveas 成功生成')
    # 删除计算书
    os.remove(os.path.join(docx_save_path, "falsework_soleplate_template_saveas.docx"))

    
# 支架模型计算书
def Midas_Model_JiSuanShu(wdApp, aDoc, folder_save_path, model_post_result):
    # 替换word中的公式和文本所需要的钢管桩的参数
    Fx, Mm, L, Area, Ix, ix, Phi, Wx, quote_lambda, Nex, sigma = SteelPile_param(model_post_result[0])
    Word_above_structure_picture(aDoc, folder_save_path) # word图片替换
    word_above_structure_TXT_Table(aDoc, model_post_result[0], Fx, Mm, sigma) # word文本替换
    cal_SteelPile_stability(wdApp, aDoc, Fx, Mm, L, Area, Ix, ix, Phi, Wx, quote_lambda, Nex, sigma) # word公式替换
    

# 扩大基础
def KuoDaJiChu_JiSuanShu(wdApp, aDoc, folder_open_path, folder_save_path, Fzmax_lst, Mxmax_lst, Mymax_lst, foundation_value, if_Error):
    # 基础计算以下内容：基底应力检算、基础强度检算

    # 基础参数
    foundation_x, foundation_y, foundation_z, Pile_D, concrete_gc, Foundation_Base_Level, soild_gs, foundation_fa, Weak_Layer_Horizon, Dispersion_Angle, foundation_faz, ksi, concrete_grade, rebar_grade, as0, steel_s1, steel_s2, steel_d1, steel_d2 = foundation_value
    # 荷载参数
    Fzmax_Fx, Fzmax_Fy, Fzmax_Fz, Fzmax_Mx, Fzmax_My, Fzmax_Mz = Fzmax_lst
    Mxmax_Fx, Mxmax_Fy, Mxmax_Fz, Mxmax_Mx, Mxmax_My, Mxmax_Mz = Mxmax_lst
    Mymax_Fx, Mymax_Fy, Mymax_Fz, Mymax_Mx, Mymax_My, Mymax_Mz = Mymax_lst

    # 基底应力检算
    # 计算计算书的文本和公式需要的参数
    if Fzmax_Mx >= Fzmax_My:
        # 最大Fz和对应的Mx
        F1, M1 = [Fzmax_Fz, Fzmax_Mx]
        pkmax_F_lst = foundation_pressure(foundation_x, foundation_y, foundation_z, concrete_gc, F1, F1, M1, Foundation_Base_Level, Weak_Layer_Horizon, soild_gs, Dispersion_Angle, foundation_fa, foundation_faz, "Mx")
        pkmax_F_lst.append([F1, F1, M1])
    else:
        # 最大Fz和对应的My
        F1, M1 = [Fzmax_Fz, Fzmax_My]
        pkmax_F_lst = foundation_pressure(foundation_x, foundation_y, foundation_z, concrete_gc, F1, F1, M1, Foundation_Base_Level, Weak_Layer_Horizon, soild_gs, Dispersion_Angle, foundation_fa, foundation_faz, "My")
        pkmax_F_lst.append([F1, F1, M1])
    if Mxmax_Mx >= Mymax_My:
        # 最大Mx和对应的Fz
        F2, M2 = [Mxmax_Fz, Mxmax_Mx]
        pkmax_M_lst = foundation_pressure(foundation_x, foundation_y, foundation_z, concrete_gc, F1, F2, M2, Foundation_Base_Level, Weak_Layer_Horizon, soild_gs, Dispersion_Angle, foundation_fa, foundation_faz, "Mx")
        pkmax_M_lst.append([F1, F2, M2])
    else:
        # 最大My和对应的Fz
        F2, M2 = [Mymax_Fz, Mymax_My]
        pkmax_M_lst = foundation_pressure(foundation_x, foundation_y, foundation_z, concrete_gc, F1, F2, M2, Foundation_Base_Level, Weak_Layer_Horizon, soild_gs, Dispersion_Angle, foundation_fa, foundation_faz, "My")
        pkmax_M_lst.append([F1, F2, M2])

    # 判断以上四种情况pkmax最大的一项, 对于软弱下卧层和pk检算, 取最大Fz
    pkmax_lst = max([pkmax_F_lst, pkmax_M_lst], key=lambda x: x[7])
    Ak, Gk, pk, Mk_e, Mk_a, Eccentricity, Wk, pkmax, pkmin, pcz, pc, pz, if_satisfied_fa1, if_satisfied_fa2, if_satisfied_faz, type_M, [Fk1,Fk2,Mk] = pkmax_lst
    
    # 无软弱下卧层
    if if_satisfied_faz == "":
        # 轴心受压
        if if_satisfied_fa2 == "":
            print("轴心荷载作用，无软弱下卧层")
            print("开始生成 cal_doc_pk ...")
            cDoc = open_doc_saveas_adoc(wdApp, "calculate_sample_pk.docx", "cal_doc_pk.docx", folder_open_path, folder_save_path) # 另存为word
            cDoc_name = "cal_doc_pk.docx"
            try:
                word_Pad_Foundation_Fonly(cDoc, folder_save_path) # word图片替换
                word_pk_TXT_Table(cDoc, Fk1, foundation_x, foundation_y, foundation_z, Foundation_Base_Level, foundation_fa, if_satisfied_fa1) # word文本替换
                cal_foundation_contact_stress(wdApp, cDoc, foundation_x, foundation_y, foundation_z, concrete_gc, Fk1, Fk2, Mk, Foundation_Base_Level, Weak_Layer_Horizon, soild_gs, Dispersion_Angle, foundation_fa, foundation_faz, Ak, Gk, pk, Mk_e, Mk_a, Eccentricity, Wk, pkmax, pkmin, pcz, pc, pz, if_satisfied_fa1, if_satisfied_fa2, if_satisfied_faz) # word公式替换
                print("cal_doc_pk 成功生成")
            except:
                print("cal_doc_pk 生成失败")
                if_Error[0] = True
        # 偏心受压
        else:
            print("偏心荷载作用，无软弱下卧层")
            print("开始生成 cal_doc_pk_pkmax ...")
            cDoc = open_doc_saveas_adoc(wdApp, "calculate_sample_pk_pkmax.docx", "cal_doc_pk_pkmax.docx", folder_open_path, folder_save_path) # 另存为word
            cDoc_name = "cal_doc_pk_pkmax.docx"
            try:
                word_Pad_Foundation_FandM(cDoc, folder_save_path) # word图片替换
                word_pk_pkmax_TXT_Table(cDoc, if_satisfied_fa1, if_satisfied_fa2, Fk1, Fk2, Mk, foundation_x, foundation_y, foundation_z, Foundation_Base_Level, foundation_fa, Eccentricity) # word文本替换
                cal_foundation_contact_stress(wdApp, cDoc, foundation_x, foundation_y, foundation_z, concrete_gc, Fk1, Fk2, Mk, Foundation_Base_Level, Weak_Layer_Horizon, soild_gs, Dispersion_Angle, foundation_fa, foundation_faz, Ak, Gk, pk, Mk_e, Mk_a, Eccentricity, Wk, pkmax, pkmin, pcz, pc, pz, if_satisfied_fa1, if_satisfied_fa2, if_satisfied_faz) # word公式替换
                print("cal_doc_pk_pkmax 成功生成")
            except:
                print("cal_doc_pk_pkmax 生成失败")
                if_Error[0] = True
    # 有软弱下卧层
    else:
        # 轴心受压
        if if_satisfied_fa2 == "":
            print("轴心荷载作用，有软弱下卧层")
            print("开始生成 cal_doc_pk_pz ...")
            cDoc = open_doc_saveas_adoc(wdApp, "calculate_sample_pk_pz.docx", "cal_doc_pk_pz.docx", folder_open_path, folder_save_path) # 另存为word
            cDoc_name = "cal_doc_pk_pz.docx"
            try:
                word_Pad_Foundation_Fonly(cDoc, folder_save_path) # word图片替换
                word_pk_pz_TXT_Table(cDoc, if_satisfied_fa1, if_satisfied_faz, Fk1, foundation_x, foundation_y, foundation_z, foundation_fa, Foundation_Base_Level, Weak_Layer_Horizon, foundation_faz, Dispersion_Angle) # word文本替换
                cal_foundation_contact_stress(wdApp, cDoc, foundation_x, foundation_y, foundation_z, concrete_gc, Fk1, Fk2, Mk, Foundation_Base_Level, Weak_Layer_Horizon, soild_gs, Dispersion_Angle, foundation_fa, foundation_faz, Ak, Gk, pk, Mk_e, Mk_a, Eccentricity, Wk, pkmax, pkmin, pcz, pc, pz, if_satisfied_fa1, if_satisfied_fa2, if_satisfied_faz) # word公式替换
                print("cal_doc_pk_pz 成功生成")
            except:
                print("cal_doc_pk_pz 生成失败")
                if_Error[0] = True
        # 偏心受压
        else:
            print("偏心荷载作用，有软弱下卧层")
            print("开始生成 cal_doc_pk_pkmax_pz ...")
            cDoc = open_doc_saveas_adoc(wdApp, "calculate_sample_pk_pkmax_pz.docx", "cal_doc_pk_pkmax_pz.docx", folder_open_path, folder_save_path) # 另存为word
            cDoc_name = "cal_doc_pk_pkmax_pz.docx"
            try:
                word_Pad_Foundation_FandM(cDoc, folder_save_path) # word图片替换
                word_pk_pkmax_pz_TXT_Table(cDoc, if_satisfied_fa1, if_satisfied_fa2, if_satisfied_faz, Fk1, Fk2, Mk, foundation_x, foundation_y, foundation_z, foundation_fa, Foundation_Base_Level, Weak_Layer_Horizon, foundation_faz, Dispersion_Angle, Eccentricity) # word文本替换
                cal_foundation_contact_stress(wdApp, cDoc, foundation_x, foundation_y, foundation_z, concrete_gc, Fk1, Fk2, Mk, Foundation_Base_Level, Weak_Layer_Horizon, soild_gs, Dispersion_Angle, foundation_fa, foundation_faz, Ak, Gk, pk, Mk_e, Mk_a, Eccentricity, Wk, pkmax, pkmin, pcz, pc, pz, if_satisfied_fa1, if_satisfied_fa2, if_satisfied_faz) # word公式替换
                print("cal_doc_pk_pkmax_pz 成功生成")
            except:
                print("cal_doc_pk_pkmax_pz 生成失败")
                if_Error[0] = True

    # 基础强度检算
    # 计算基底净反力
    p = cal_p(ksi, pk, foundation_x, foundation_y, Gk)
    pmax = cal_pmax(ksi, pkmax, foundation_x, foundation_y, Gk)
    pmin = cal_pmin(ksi, pkmin, foundation_x, foundation_y, Gk)
    print("基底净反力: ", p, pmax, pmin)

    ft = concrete_f(concrete_grade)[1]
    fy = rebar_f(rebar_grade)[0]
    at = 0 # 立柱与基础x边对应的柱宽
    bt = 0 # 立柱与基础y边对应的柱宽

    print("开始生成 cal_doc_Vs_Fl_M ...")
    dDoc = open_doc_saveas_adoc(wdApp, "calculate_sample_Vs_Fl_M.docx", "cal_doc_Vs_Fl_M.docx", folder_open_path, folder_save_path) # 另存为word
    try:
        JCAv1, JCAv2, JCV1, JCV2, Bhs, JCA01, JCA02, Allowable_Vs1, Allowable_Vs2, JCV_sf1, JCV_sf2 = Vs_Verification(foundation_x, foundation_y, foundation_z, as0, at, bt, Pile_D, pmax, ft) # 抗剪计算
        JCAl1, JCAl2, JCFl1, JCFl2, h0, Bhp, am, Allowable_Fl, JCF_sf1, JCF_sf2 = Fl_Verification(foundation_x, foundation_y, foundation_z, as0, at, bt, Pile_D, pmax, ft) # 抗冲切计算
        JC_a1, JC_aa, JC_bb, JC_M1, JC_M2, MAs1, MAs2, JC_Ms1, JC_Md1, JC_Mp1, JC_MA1, JC_Ms2, JC_Md2, JC_Mp2, JC_MA2, JCM_sf1, JCM_sf2, JCM_if1, JCM_if2 = M_Verification(foundation_x, foundation_y, foundation_z, as0, at, bt, Pile_D, p, pmax, pmin, fy, steel_s1, steel_s2, steel_d1, steel_d2, type_M) # 抗弯及配筋计算
        word_foundation_Vs_FL_M_TXT_Table(dDoc, concrete_grade, as0,
                                            JCV_sf1, JCV_sf2, JCF_sf1, JCF_sf2, JCM_sf1, JCM_sf2,
                                            JCAv1, JCV1, JCAv2, JCV2, Bhs, ft, JCA01, JCA02, Allowable_Vs1, Allowable_Vs2, 
                                            JCAl1, JCFl1, JCAl2, JCFl2, Bhp, am, h0, Allowable_Fl,
                                            JC_a1, JC_aa, JC_bb, JC_M1, MAs1, JC_M2, MAs2, JC_Md1, JC_Md2, JC_Ms1, JC_Ms2, 
                                            JC_MA1, JC_MA2, JC_Mp1, JC_Mp2, JCM_if1, JCM_if2
                                    ) # word文本替换
        word_p_pmax_pmin(wdApp, dDoc, foundation_x, foundation_y, ksi, Gk, pk, pkmax, pkmin, p, pmax, pmin) # word公式替换
        print("cal_doc_Vs_Fl_M 成功生成")
    except:
        print("cal_doc_Vs_Fl_M 生成失败")
        if_Error[0] = True
    # 合并计算书
    WordA_PasteTo_WordB_DeleteLines(wdApp, dDoc, cDoc, "基础强度计算书开始", "基础强度计算书结束", "基础强度计算书粘贴")
    WordA_PasteTo_WordB_DeleteLines(wdApp, cDoc, aDoc, "扩大基础计算书开始", "扩大基础计算书结束", "基础计算书粘贴")
    # 保存
    adoc_save_close(cDoc)
    adoc_save_close(dDoc)
    # 删除计算书
    os.remove(os.path.join(folder_save_path, cDoc_name))
    os.remove(os.path.join(folder_save_path, "cal_doc_Vs_Fl_M.docx"))

#=====================================================================================================================================================================

# 钢管打入桩
def GangGuanZhuang_JiSuanShu(wdApp, aDoc, folder_open_path, folder_save_path, Fzmax_lst, Fxymax_lst, Pile_type, foundation_value, if_Error):
    # 基础参数
    # Soil_Layer_Parameter_Table = {"土层分类":[], "土层厚度":[], "黏聚力":[], "内摩擦角":[], "侧摩阻":[], "端阻":[]}
    # Pile_D, Pile_t, Soil_vb, Soil_Layer_n, Soil_Layer_Parameter_Table = foundation_value
    Pile_D, Pile_t, Soil_vb, Soil_Layer_Parameter_Table = foundation_value
    print(foundation_value)
    print(Soil_Layer_Parameter_Table)
    Soil_type_lst = Soil_Layer_Parameter_Table['土层分类']
    Soil_Layer_Hlst = Soil_Layer_Parameter_Table['土层厚度']
    Soil_Cohesion_lst = Soil_Layer_Parameter_Table['黏聚力']
    Soil_friction_lst = Soil_Layer_Parameter_Table['内摩擦角']
    Soil_qsk_lst = Soil_Layer_Parameter_Table['侧摩阻']
    Soil_psk_lst = Soil_Layer_Parameter_Table['端阻']
    Soil_Layer_Quantity = len(Soil_qsk_lst) # 参与受力的所有土层数
    Pile_H = sum(Soil_Layer_Hlst) # 钢管桩打入深度，单位m
    # 荷载参数
    Pile_Nk = Fzmax_lst[2] # 竖向力，单位KN
    Pile_Hx = Fxymax_lst[0] # 水平力，单位KN
    Pile_Hy = Fxymax_lst[1] # 水平力，单位KN
    Pile_Hk = round(math.sqrt(Pile_Hx**2+Pile_Hy**2), 1)

    # 根据基础参数和荷载参数计算生成计算书所需要的参数
    Qsk, Qpk, Quk, Ra = JGJ94_Vertical_Capacity_steelpile(Pile_D, Soil_qsk_lst, Soil_psk_lst, Soil_Layer_Hlst) # 竖向力
    # friction_n, Cohesion_n, m_n, b0, EI, alpha, alpha_h, Vm, Vx, Rha = JGJ94_Lateral_Capacity_steelpile(Pile_D, Pile_t, Pile_H, Soil_Layer_n, Soil_friction_lst, Soil_Cohesion_lst, Soil_vb) # 水平力
    friction_n, Cohesion_n, m_n, b0, EI, alpha, alpha_h, Vm, Vx, Rha = JGJ94_Lateral_Capacity_steelpile(Pile_D, Pile_t, Pile_H, Soil_friction_lst, Soil_Cohesion_lst, Soil_vb) # 水平力

    # 生成计算书
    print("开始生成 cal_DrivenPile ...")
    cDoc = open_doc_saveas_adoc(wdApp, "calculate_sample_DrivenPile.docx", "cal_DrivenPile.docx", folder_open_path, folder_save_path) # 另存为word
    try:
        # 文本替换
        word_SteelPile_TXT_Table(cDoc, Pile_D, Pile_t, Pile_H, Pile_Nk, Pile_Hk, Soil_vb, alpha_h, Vm, Vx, Rha, Ra, b0)
        # 公式替换
        word_Pile_Ra(wdApp, cDoc, Qsk, Qpk, Quk, Ra, Pile_type, Pile_Nk) # 竖向力
        word_Pile_Rha2(wdApp, cDoc, Pile_Hk, friction_n, Cohesion_n, Soil_vb, m_n, b0, EI, alpha, Rha, Vx) # 水平力
        # 对于桩基础，将桩基础计算书粘贴到总计算书中
        WordA_PasteTo_WordB_DeleteLines(wdApp, cDoc, aDoc, "桩基础计算书开始", "桩基础计算书结束", "基础计算书粘贴")
        adoc_save_close(cDoc) # 保存
        print("cal_DrivenPile 成功生成")
    except:
        print("cal_DrivenPile 生成失败")
        if_Error[0] = True
    # 删除计算书
    os.remove(os.path.join(folder_save_path, "cal_DrivenPile.docx"))

#=====================================================================================================================================================================

# 混凝土预制桩
def YuZhiZhuang_JiSuanShu(wdApp, aDoc, folder_open_path, folder_save_path, Fzmax_lst, Fxymax_lst, Pile_type, foundation_value, if_Error):
    # 基础参数
    # Soil_Layer_Parameter_Table = {"土层分类":[], "土层厚度":[], "黏聚力":[], "内摩擦角":[], "侧摩阻":[], "端阻":[]}
    # Pile_D, Pile_t, concrete_grade, Soil_vb, Soil_Layer_n, Soil_Layer_Parameter_Table = foundation_value
    Pile_D, Pile_t, concrete_grade, Soil_vb, Soil_Layer_Parameter_Table = foundation_value
    Soil_type_lst = Soil_Layer_Parameter_Table['土层分类']
    Soil_Layer_Hlst = Soil_Layer_Parameter_Table['土层厚度']
    Soil_Cohesion_lst = Soil_Layer_Parameter_Table['黏聚力']
    Soil_friction_lst = Soil_Layer_Parameter_Table['内摩擦角']
    Soil_qsk_lst = Soil_Layer_Parameter_Table['侧摩阻']
    Soil_psk_lst = Soil_Layer_Parameter_Table['端阻']
    Soil_Layer_Quantity = len(Soil_qsk_lst) # 参与受力的所有土层数
    Pile_H = sum(Soil_Layer_Hlst) # 钢管桩打入深度，单位m
    # 荷载参数
    Pile_Nk = Fzmax_lst[2] # 竖向力，单位KN
    Pile_Hx = Fxymax_lst[0] # 水平力，单位KN
    Pile_Hy = Fxymax_lst[1] # 水平力，单位KN
    Pile_Hk = round(math.sqrt(Pile_Hx**2+Pile_Hy**2), 1)

    # 根据基础参数和荷载参数计算生成计算书所需要的参数
    Qsk, Qpk, Quk, Ra = JGJ94_Vertical_Capacity_precastpile(Pile_D, Pile_t, Soil_qsk_lst, Soil_psk_lst, Soil_Layer_Hlst) # 竖向力
    # friction_n, Cohesion_n, m_n, b0, EI, alpha, alpha_h, Vm, Vx, Rha = JGJ94_Lateral_Capacity_precastpile(concrete_grade, Pile_D, Pile_t, Pile_H, Soil_Layer_n, Soil_friction_lst, Soil_Cohesion_lst, Soil_vb) # 水平力
    friction_n, Cohesion_n, m_n, b0, EI, alpha, alpha_h, Vm, Vx, Rha = JGJ94_Lateral_Capacity_precastpile(concrete_grade, Pile_D, Pile_t, Pile_H, Soil_friction_lst, Soil_Cohesion_lst, Soil_vb) # 水平力
    
    # 生成计算书
    print("开始生成 cal_DrivenPile2 ...")
    cDoc = open_doc_saveas_adoc(wdApp, "calculate_sample_DrivenPile2.docx", "cal_DrivenPile2.docx", folder_open_path, folder_save_path) # 另存为word
    try:
        # 文本替换
        word_ConcretePile_TXT_Table(cDoc, Pile_D, Pile_t, concrete_grade, Pile_H, Pile_Nk, Pile_Hk, Soil_vb, alpha_h, Vm, Vx, Rha, Ra, b0)
        # 公式替换
        word_Pile_Ra(wdApp, cDoc, Qsk, Qpk, Quk, Ra, Pile_type, Pile_Nk) # 竖向力
        word_Pile_Rha2(wdApp, cDoc, Pile_Hk, friction_n, Cohesion_n, Soil_vb, m_n, b0, EI, alpha, Rha, Vx) # 水平力
        # 对于桩基础，将桩基础计算书粘贴到总计算书中
        WordA_PasteTo_WordB_DeleteLines(wdApp, cDoc, aDoc, "桩基础计算书开始", "桩基础计算书结束", "基础计算书粘贴")
        adoc_save_close(cDoc) # 保存
        print("cal_DrivenPile2 成功生成")
    except:
        print("cal_DrivenPile2 生成失败")
        if_Error[0] = True
    # 删除计算书
    os.remove(os.path.join(folder_save_path, "cal_DrivenPile2.docx"))

#=====================================================================================================================================================================

# 钢筋混凝土灌注桩
def GuanZhuZhuang_JiSuanShu(wdApp, aDoc, folder_open_path, folder_save_path, Fzmax_lst, Fxymax_lst, Pile_type, foundation_value, if_Error):
    # 基础参数
    # Soil_Layer_Parameter_Table = {"土层分类":[], "土层厚度":[], "黏聚力":[], "内摩擦角":[], "侧摩阻":[], "端阻":[]}
    # Pile_D, concrete_grade, rebar_grade, Pile_p, Pile_as, Soil_vb, Soil_Layer_n, Soil_Layer_Parameter_Table = foundation_value
    Pile_D, concrete_grade, rebar_grade, Pile_p, Pile_as, Soil_vb, Soil_Layer_Parameter_Table = foundation_value
    Soil_type_lst = Soil_Layer_Parameter_Table['土层分类']
    Soil_Layer_Hlst = Soil_Layer_Parameter_Table['土层厚度']
    Soil_Cohesion_lst = Soil_Layer_Parameter_Table['黏聚力']
    Soil_friction_lst = Soil_Layer_Parameter_Table['内摩擦角']
    Soil_qsk_lst = Soil_Layer_Parameter_Table['侧摩阻']
    Soil_psk_lst = Soil_Layer_Parameter_Table['端阻']
    Soil_Layer_Quantity = len(Soil_qsk_lst) # 参与受力的所有土层数
    Pile_H = sum(Soil_Layer_Hlst) # 钢管桩打入深度，单位m
    # 荷载参数
    Pile_Nk = Fzmax_lst[2] # 竖向力，单位KN
    Pile_Hx = Fxymax_lst[0] # 水平力，单位KN
    Pile_Hy = Fxymax_lst[1] # 水平力，单位KN
    Pile_Hk = round(math.sqrt(Pile_Hx**2+Pile_Hy**2), 1)

    # 根据基础参数和荷载参数计算生成计算书所需要的参数
    Qsk, Qpk, Quk, Ra = JGJ94_Vertical_Capacity_concretepile(Pile_D, Soil_qsk_lst, Soil_psk_lst, Soil_Layer_Hlst, Soil_type_lst) # 竖向力
    # friction_n, Cohesion_n, m_n, b0, EI, alpha, alpha_h, Vm, Vx, Es, Ec, W0, An, Rha = JGJ94_Lateral_Capacity_concretepile(concrete_grade, rebar_grade, Pile_D, Pile_p, Pile_as, Pile_Nk, Pile_H, Soil_Layer_n, Soil_friction_lst, Soil_Cohesion_lst, Soil_vb) # 水平力
    friction_n, Cohesion_n, m_n, b0, EI, alpha, alpha_h, Vm, Vx, Es, Ec, W0, An, Rha = JGJ94_Lateral_Capacity_concretepile(concrete_grade, rebar_grade, Pile_D, Pile_p, Pile_as, Pile_Nk, Pile_H, Soil_friction_lst, Soil_Cohesion_lst, Soil_vb) # 水平力
    
    if Pile_p < 0.65:
        print("开始生成 cal_DrilledShaft ...")
        cDoc = open_doc_saveas_adoc(wdApp, "calculate_sample_DrilledShaft.docx", "cal_DrilledShaft.docx", folder_open_path, folder_save_path) # 另存为word
        cDoc_name = "cal_DrilledShaft.docx"
        try:
            # 文本替换
            word_DrilledShaft_TXT_Table(cDoc, concrete_grade, Pile_D, Pile_as, Pile_p, Pile_H, Pile_Nk, Pile_Hk, Soil_vb, alpha_h, Vm, Vx, Rha, Ra, b0)
            # 公式替换
            word_Pile_Ra(wdApp, cDoc, Qsk, Qpk, Quk, Ra, Pile_type, Pile_Nk) # 竖向力
            word_Pile_Rha1(wdApp, cDoc, Pile_D, Pile_p, Pile_Hk, Pile_as, friction_n, Cohesion_n, Soil_vb, m_n, b0, EI, alpha, Rha, Es, Ec, An, W0) # 水平力
            print("cal_DrilledShaft 成功生成")
        except:
            print("cal_DrilledShaft 生成失败")
            if_Error[0] = True
    else:
        print("开始生成 cal_DrilledShaft2 ...")
        cDoc = open_doc_saveas_adoc(wdApp, "calculate_sample_DrilledShaft2.docx", "cal_DrilledShaft2.docx", folder_open_path, folder_save_path) # 另存为word
        cDoc_name = "cal_DrilledShaft2.docx"
        try:
            # 文本替换
            word_DrilledShaft_TXT_Table(cDoc, concrete_grade, Pile_D, Pile_as, Pile_p, Pile_H, Pile_Nk, Pile_Hk, Soil_vb, alpha_h, Vm, Vx, Rha, Ra, b0)
            # 公式替换
            word_Pile_Ra(wdApp, cDoc, Qsk, Qpk, Quk, Ra, Pile_type, Pile_Nk) # 竖向力
            word_Pile_Rha2(wdApp, cDoc, Pile_Hk, friction_n, Cohesion_n, Soil_vb, m_n, b0, EI, alpha, Rha, Vx) # 水平力
            print("cal_DrilledShaft2 成功生成")
        except:
            print("cal_DrilledShaft2 生成失败")
            if_Error[0] = True
    # 对于桩基础，将桩基础计算书粘贴到总计算书中
    WordA_PasteTo_WordB_DeleteLines(wdApp, cDoc, aDoc, "桩基础计算书开始", "桩基础计算书结束", "基础计算书粘贴")
    adoc_save_close(cDoc) # 保存
    # 删除计算书
    os.remove(os.path.join(folder_save_path, cDoc_name))

#=====================================================================================================================================================================

# 总函数, 生成所有word
def make_all_doc(response, Applocation, openfile_path, model_post_result, soleplate_values, foundation_name, foundation_standard, foundation_value_dict, saveas_path, docx_save_path = "", file_extension = ".txt"):
    # 指定计算模型，则图片也会存到模型所在的文件夹内
    Midas_path = os.path.dirname(openfile_path) # 文件目录路径
    Midas_name = os.path.basename(openfile_path) # 文件名
    # 计算书保存路径
    if docx_save_path == "":
        docx_save_path = Midas_path
    print(docx_save_path)
    # 计算书打开路径
    docx_open_path = os.path.join(Applocation, 'Support\\templates')
    txt_path = file_extension_Modified(openfile_path, file_extension)
    # 任意一个计算书出错则变成True
    if_Error = [False]
    # 没有桩底反力边界组则不生成
    if not if_Boundary_Group_name_exist("桩底反力"):
        print("模型[组]中缺失[桩底反力]边界组, 请添加后重试")
    else:
        # response变量代表了用户是否需要读取模型的结果
        if response:
            # 对于需要运行的模型，没有结果文件，则先运行，有则不运行
            Midas_Analysis(Midas_name, Midas_path)
            # 获取模型基本信息
            NODE_dict_res, ELEM_dict_res, SECT_dict_res, get_structure_group_res = get_node_elem_sect_grup_dict()
            # 对Midas模型进行结果的读取并截图
            model_post_result[0] = Post_processing_Midas_Model(Midas_path, NODE_dict_res, ELEM_dict_res, SECT_dict_res, get_structure_group_res)
        if model_post_result[0] == {}:
            print("未读取到模型结果, 无法生成计算书")
        else:
            # 与word建立连接
            wdApp = WordApp_Dispatch()
            try:
                # 打开总计算书和前处理计算书并另存为副本(必须的)
                print("开始生成 cal_doc ...")
                aDoc = open_doc_saveas_adoc(wdApp, 'calculate_sample.docx', 'cal_doc.docx', docx_open_path, docx_save_path) # 返回 cal_doc 对应的doc变量
                print("开始生成 cal_doc_pre ...")
                bDoc = open_doc_saveas_adoc(wdApp, 'calculate_sample_pre.docx', 'cal_doc_pre.docx', docx_open_path, docx_save_path) # 返回 cal_doc_pre 对应的doc变量
                # 现浇支架前处理部分(计算依据、设计参数、材料特性、工况分析、荷载组合)
                try:
                    Pre_JiSuanShu(wdApp, aDoc, bDoc, txt_path, foundation_standard, foundation_value_dict)
                    Soleplate_JiSuanShu(wdApp, aDoc, docx_open_path, docx_save_path, soleplate_values)
                except:
                    print("计算书前处理失败")
                    if_Error[0] = True
                # 现浇支架模型计算书部分 
                try:
                    # 根据 calculate_sample.docx 生成 cal_doc.docx
                    Midas_Model_JiSuanShu(wdApp, aDoc, docx_save_path, model_post_result)
                except:
                    print("计算书模型结果处理失败")
                    if_Error[0] = True
                # 基础计算书部分
                # if foundation_value_dict != {'浅基础': [], '深基础': []}:
                if foundation_name != "根据模型信息":
                    # 模型中反力结果的最大竖向力、水平力及两个方向的最大弯矩
                    foundation_loads = model_post_result[0]["model_post_reaction"]
                    Fzmax_lst = [round(abs(float(x)),2) for x in foundation_loads[0][3:]]
                    Fxymax_lst = [round(abs(float(x)),2) for x in foundation_loads[1][3:]]
                    Mxmax_lst = [round(abs(float(x)),2) for x in foundation_loads[2][3:]]
                    Mymax_lst = [round(abs(float(x)),2) for x in foundation_loads[3][3:]]
                    print("最大竖向力:", Fzmax_lst)
                    print("最大水平力:", Fxymax_lst)
                    print("最大弯矩Mx:", Mxmax_lst)
                    print("最大弯矩My:", Mymax_lst)
                    if foundation_name == "扩大基础":
                        foundation_value = foundation_value_dict["浅基础"]
                        KuoDaJiChu_JiSuanShu(wdApp, aDoc, docx_open_path, docx_save_path, Fzmax_lst, Mxmax_lst, Mymax_lst, foundation_value, if_Error)
                    elif foundation_name == "钢管打入桩":
                        foundation_value = foundation_value_dict["深基础"]
                        GangGuanZhuang_JiSuanShu(wdApp, aDoc, docx_open_path, docx_save_path, Fzmax_lst, Fxymax_lst, foundation_name, foundation_value, if_Error)
                    elif foundation_name == "混凝土预制桩":
                        foundation_value = foundation_value_dict["深基础"]
                        YuZhiZhuang_JiSuanShu(wdApp, aDoc, docx_open_path, docx_save_path, Fzmax_lst, Fxymax_lst, foundation_name, foundation_value, if_Error)
                    elif foundation_name == "混凝土灌注桩":
                        foundation_value = foundation_value_dict["深基础"]
                        GuanZhuZhuang_JiSuanShu(wdApp, aDoc, docx_open_path, docx_save_path, Fzmax_lst, Fxymax_lst, foundation_name, foundation_value, if_Error)
                else:
                    # print("未检测到基础相关参数，仅生成模型对应的计算书")
                    print("基础形式无, 取消生成基础计算书")
                    # 删除改行
                    Delete_Word_FindText_Paragraphs(aDoc, '基础计算书粘贴')
                    Delete_Word_FindText_Paragraphs(aDoc, '{Selected_foundation_Standard}')
                # 另存为最终路径
                try:
                    aDoc.SaveAs(saveas_path)
                    print(f'计算书在指定路径{saveas_path}下成功生成')
                except:
                    print("指定路径计算书另存为失败")
                # 保存并关闭另存为的文件
                adoc_save_close(aDoc)
                adoc_save_close(bDoc)
                # 删除模板另存为的计算书
                os.remove(os.path.join(docx_save_path, 'cal_doc.docx'))
                os.remove(os.path.join(docx_save_path, 'cal_doc_pre.docx'))
            except:
                if_Error[0] = True
            finally:
                # 退出和word的连接
                WordApp_Quit(wdApp)
                # 所有计算书均未出错
                if not if_Error[0]:
                    # print("cal_doc 成功生成")
                    print("计算书模块已全部加载完成")
                else:
                    print("有部分计算书生成失败, 参数错误或 Word 处于打开状态")
            

def scientific_notation(result, decimal_places=2):
    """
    计算 1000*(150^3)/12 并以科学计数法表示
    参数:
    decimal_places: a保留的小数位数
    返回:
    [a, b] 其中 a * 10^b
    """
    # 转换为科学计数法形式
    # 使用字符串格式化和分割来获取系数和指数
    sci_notation = f"{result:.{decimal_places}e}"
    parts = sci_notation.split('e')
    a = float(parts[0])
    b = int(parts[1])
    return [a, b]

#=====================================================================================================================================================================

