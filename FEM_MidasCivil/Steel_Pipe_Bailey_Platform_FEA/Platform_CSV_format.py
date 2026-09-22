# 3. 本地模块
from General.ExcelHandle import write_csv
from General.DataUtils   import midasdisttolst2
from General.Geometry    import get_distance_from_xlst

from FEM_MidasCivil.Steel_Pipe_Bailey_Platform_FEA.Platform_MCT_format_FEM import build_mct_from_ui


# 各个csv文件的名字
csv_filename1 = 'Wharf.csv' # 栈桥
csv_filename2 = 'DrillingPlatform.csv' # 平台
csv_filename3 = 'Bracket.csv' # 支架
csv_filename4 = 'Cofferdam.csv' # 围堰


def CSV_Wharf(app):
    bailey_Cantilever_length, rib_sec_name,rib_along_layout_gaplst, bailey_along_layout_namelst, bailey_along_layout_gaplst, bailey_cross_layout_gaplst,  fenpeiliang_sec_name, ganggz_along_layout_gaplst, ganggz_cross_layout_gaplst = PlatformUI_CSV(app)
    '''
    栈桥的csv文件格式
    '''
    line1 = ['贝雷梁', '伸出长度：']
    line2 = ['', '类型：']
    line3 = ['', '纵向：']
    line4 = ['', '横向：']
    line5 = ['大肋', '名称：']
    line6 = ['', '间距：']
    line7 = ['分配梁', '名称：']
    line8 = ['钢管桩', '纵向：']
    line9 = ['', '横向：']
    line10 = ['横向联结系', '竖向：']
    line11 = ['纵向联结系', '竖向：']
    # 写入数据
    # line1
    # bailey_Cantilever_length = 1500 # 贝雷悬出长度
    line1.append(bailey_Cantilever_length)
    # line2
    # bailey_along_layout_namelst = ['1.5m贝雷梁', '贝雷梁']
    for x in bailey_along_layout_namelst:
        line2.extend([str(x), '', '', ''])
    # line3
    # bailey_along_layout_gaplst  = [[1500], [3000, 3000, 3000, 3000, 3000, 3000, 3000, 3000]]
    for x in bailey_along_layout_gaplst:
        line3.extend([str(x[0]), '@', str(len(x)), ','])
    line3[-1] = '/'
    # line4
    # bailey_cross_layout_gaplst  = [[900, 900], [450, 450, 450], [900, 900], [450]]
    for x in bailey_cross_layout_gaplst:
        line4.extend([str(x[0]), '@', str(len(x)), ','])
    line4[-1] = '/'
    # line5
    # rib_sec_name = 'I10'
    line5.append(rib_sec_name)
    # line6
    # rib_along_layout_gaplst  = [[1000, 1000, 1000, 1000, 1000], [500, 500, 500, 500, 500]]
    for x in rib_along_layout_gaplst:
        line6.extend([str(x[0]), '@', str(len(x)), ','])
    line6[-1] = '/'
    # line7
    # fenpeiliang_sec_name = '2I32a'
    line7.append(fenpeiliang_sec_name)
    # line8
    # ganggz_along_layout_gaplst  = [[12000, 12000], [15000, 15000, 15000]]
    for x in ganggz_along_layout_gaplst:
        line8.extend([str(x[0]), '@', str(len(x)), ','])
    line8[-1] = '/'
    # line9
    # ganggz_cross_layout_gaplst  = [[3000, 3000]]
    for x in ganggz_cross_layout_gaplst:
        line9.extend([str(x[0]), '@', str(len(x)), ','])
    line9[-1] = '/'
    # line10
    line10.extend(['2000', '@', '1'])
    # line11
    line11.extend(['2000', '@', '1'])
    # 同化格式
    linelst = [line1, line2, line3, line4, line5, line6, line7, line8, line9, line10, line11]
    csv_line_length = max(len(line1), len(line2), len(line3), len(line4), len(line5), len(line6), len(line7), len(line8), len(line9), len(line10), len(line11))
    for lst in linelst:
        lsti = csv_line_length - len(lst)
        for i in range(lsti):
            lst.append('')
    for x in linelst:
        print(x)
    return linelst


def CSV_DrillingPlatform(app):
    '''
    平台的csv文件格式
    '''
    # line1 = ['顺桥向桩间距', '间隔个数']
    # line2 = ['横桥向桩间距', '间隔个数']
    # # 写入数据
    # # line1
    # # ganggz_along_layout_spacing  = 6000
    # # ganggz_along_layout_quantity = 3
    # line1.insert( 1, ganggz_along_layout_spacing )
    # line1.insert(-1, ganggz_along_layout_quantity)
    # # line2
    # # ganggz_cross_layout_spacing  = 9000
    # # ganggz_cross_layout_quantity = 3
    # line2.insert( 1, ganggz_cross_layout_spacing )
    # line2.insert(-1, ganggz_cross_layout_quantity)
    # # 同化格式
    # linelst = [line1, line2]
    # csv_line_length = max(len(line1), len(line2))
    # for lst in linelst:
    #     lsti = csv_line_length - len(lst)
    #     for i in range(lsti):
    #         lst.append('')
    # # for x in linelst:
    # #     print(x)
    # return linelst
    bailey_Cantilever_length, rib_sec_name,rib_along_layout_gaplst, bailey_along_layout_namelst, bailey_along_layout_gaplst, bailey_cross_layout_gaplst, fenpeiliang_sec_name, ganggz_along_layout_gaplst, ganggz_cross_layout_gaplst = PlatformUI_CSV(app)
    
    line1 = ['贝雷梁', '伸出长度：']
    line2 = ['', '类型：']
    line3 = ['', '纵向：']
    line4 = ['', '横向：']
    line5 = ['大肋', '名称：']
    line6 = ['', '间距：']
    line7 = ['分配梁', '名称：']
    line8 = ['钢管桩', '纵向：']
    line9 = ['', '横向：']
    line10 = ['横向联结系', '竖向：']
    line11 = ['纵向联结系', '竖向：']
    # 写入数据
    # line1
    # bailey_Cantilever_length = 1500 # 贝雷悬出长度
    line1.append(bailey_Cantilever_length)
    # line2
    # bailey_along_layout_namelst = ['1.5m贝雷梁', '贝雷梁']
    for x in bailey_along_layout_namelst:
        line2.extend([str(x), '', '', ''])
    # line3
    # bailey_along_layout_gaplst  = [[1500], [3000, 3000, 3000, 3000, 3000, 3000, 3000, 3000]]
    for x in bailey_along_layout_gaplst:
        line3.extend([str(x[0]), '@', str(len(x)), ','])
    line3[-1] = '/'
    # line4
    # bailey_cross_layout_gaplst  = [[900, 900], [450, 450, 450], [900, 900], [450]]
    for x in bailey_cross_layout_gaplst:
        line4.extend([str(x[0]), '@', str(len(x)), ','])
    line4[-1] = '/'
    # line5
    # rib_sec_name = 'I10'
    line5.append(rib_sec_name)
    # line6
    # rib_along_layout_gaplst  = [[1000, 1000, 1000, 1000, 1000], [500, 500, 500, 500, 500]]
    for x in rib_along_layout_gaplst:
        line6.extend([str(x[0]), '@', str(len(x)), ','])
    line6[-1] = '/'
    # line7
    # fenpeiliang_sec_name = '2I32a'
    line7.append(fenpeiliang_sec_name)
    # line8
    # ganggz_along_layout_gaplst  = [[12000, 12000], [15000, 15000, 15000]]
    for x in ganggz_along_layout_gaplst:
        line8.extend([str(x[0]), '@', str(len(x)), ','])
    line8[-1] = '/'
    # line9
    # ganggz_cross_layout_gaplst  = [[3000, 3000]]
    for x in ganggz_cross_layout_gaplst:
        line9.extend([str(x[0]), '@', str(len(x)), ','])
    line9[-1] = '/'
    # line10
    line10.extend(['2000', '@', '1'])
    # line11
    line11.extend(['2000', '@', '1'])
    # 同化格式
    linelst = [line1, line2, line3, line4, line5, line6, line7, line8, line9, line10, line11]
    csv_line_length = max(len(line1), len(line2), len(line3), len(line4), len(line5), len(line6), len(line7), len(line8), len(line9), len(line10), len(line11))
    for lst in linelst:
        lsti = csv_line_length - len(lst)
        for i in range(lsti):
            lst.append('')
    # for x in linelst:
    #     print(x)
    return linelst


def CSV_Bracket(bailey_Cantilever_length, bailey_along_layout_namelst, bailey_along_layout_gaplst, bailey_cross_layout_gaplst,
                rib_sec_name, rib_along_layout_gaplst, fenpeiliang_sec_name, ganggz_along_layout_gaplst, ganggz_cross_layout_gaplst, 
               ):
    '''
    支架的csv文件格式
    '''
    line1 = ['贝雷梁', '伸出长度：']
    line2 = ['', '类型：']
    line3 = ['', '纵向：']
    line4 = ['', '横向：']
    line5 = ['大肋', '名称：']
    line6 = ['', '间距：']
    line7 = ['分配梁', '名称：']
    line8 = ['钢管桩', '纵向：']
    line9 = ['', '横向：']
    line10 = ['横向联结系', '竖向：']
    line11 = ['纵向联结系', '竖向：']
    # 写入数据
    # line1
    # bailey_Cantilever_length = 1500 # 贝雷悬出长度
    line1.append(bailey_Cantilever_length)
    # line2
    # bailey_along_layout_namelst = ['1.5m贝雷梁', '贝雷梁']
    for x in bailey_along_layout_namelst:
        line2.extend([str(x), '', '', ''])
    # line3
    # bailey_along_layout_gaplst  = [[1500], [3000, 3000, 3000, 3000, 3000, 3000, 3000, 3000]]
    for x in bailey_along_layout_gaplst:
        line3.extend([str(x[0]), '@', str(len(x)), ','])
    line3[-1] = '/'
    # line4
    # bailey_cross_layout_gaplst  = [[900, 900], [450, 450, 450], [900, 900], [450]]
    for x in bailey_cross_layout_gaplst:
        line4.extend([str(x[0]), '@', str(len(x)), ','])
    line4[-1] = '/'
    # line5
    # rib_sec_name = 'I10'
    line5.append(rib_sec_name)
    # line6
    # rib_along_layout_gaplst  = [[1000, 1000, 1000, 1000, 1000], [500, 500, 500, 500, 500]]
    for x in rib_along_layout_gaplst:
        line6.extend([str(x[0]), '@', str(len(x)), ','])
    line6[-1] = '/'
    # line7
    # fenpeiliang_sec_name = '2I32a'
    line7.append(fenpeiliang_sec_name)
    # line8
    # ganggz_along_layout_gaplst  = [[12000, 12000], [15000, 15000, 15000]]
    for x in ganggz_along_layout_gaplst:
        line8.extend([str(x[0]), '@', str(len(x)), ','])
    line8[-1] = '/'
    # line9
    # ganggz_cross_layout_gaplst  = [[3000, 3000]]
    for x in ganggz_cross_layout_gaplst:
        line9.extend([str(x[0]), '@', str(len(x)), ','])
    line9[-1] = '/'
    # line10
    line10.extend(['2000', '@', '1'])
    # line11
    line11.extend(['2000', '@', '1'])
    # 同化格式
    linelst = [line1, line2, line3, line4, line5, line6, line7, line8, line9, line10, line11]
    csv_line_length = max(len(line1), len(line2), len(line3), len(line4), len(line5), len(line6), len(line7), len(line8), len(line9), len(line10), len(line11))
    for lst in linelst:
        lsti = csv_line_length - len(lst)
        for i in range(lsti):
            lst.append('')
    # for x in linelst:
    #     print(x)
    return linelst


def CSV_Cofferdam(cofferdam_sizelst, waler_Hlst, strut_seclst, cap_sizelst, solid_layerlst):
    '''
    围堰的csv文件格式
    '''
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

# 对一连串的数据进行整理，碰到相同的则合并
def group_consecutive_elements(lst):
    if not lst:
        return []
    
    result = []
    current_group = [lst[0]]
    
    for i in range(1, len(lst)):
        if lst[i] == lst[i-1]:
            current_group.append(lst[i])
        else:
            result.append(current_group)
            current_group = [lst[i]]
    
    result.append(current_group)
    return result


def get_bailey_along_layout_namelst(dist_lst):
    """
    输入：例如 [1500,1500,1500,3000,3000,1500,3000]
    输出：例如 ['1.5m贝雷梁','贝雷梁','1.5m贝雷梁','贝雷梁']
    """
    if not dist_lst:
        return []

    namelst = []
    prev = None

    for d in dist_lst:
        if abs(d - 3000) < 1e-6:
            name = "贝雷梁"
        elif abs(d - 1500) < 1e-6:
            name = "1.5m贝雷梁"
        else:
            raise ValueError(f"检测到非法贝雷梁间距 {d} mm，仅允许 1500 或 3000")
        if name != prev:
            namelst.append(name)
            prev = name
    return namelst

def int_or_float(Dividend, Divisor, Accuracy):
    if Dividend % Divisor != 0:
        return round(Dividend / Divisor, Accuracy)
    else:
        return int(Dividend / Divisor)
    
def SEC_trans(SEC_CAD, SEC_NAME):
    '''
    SEC_CAD: 截面信息表
    SEC_NAME: 截面名称
    '''
    H = SEC_CAD[SEC_NAME]['H'] # 截面高度
    B = SEC_CAD[SEC_NAME]['B'] # 截面宽度
    tf = SEC_CAD[SEC_NAME]['tf'] # 翼缘高度
    if SEC_NAME == '工字钢截面':
        format_str = int_or_float(H, 10, 1)
        if H >  180: Block_name = f'I{format_str}a'
        if H <= 180: Block_name = f'I{format_str}'
        print(Block_name)
    elif SEC_NAME == '槽钢截面':
        format_str = int_or_float(H, 10, 1)
        if H >= 140: Block_name = f'C{format_str}a'
        if H <  140: Block_name = f'C{format_str}'
        print(Block_name)
    elif SEC_NAME == 'HM截面' or SEC_NAME == 'HN截面':
        Block_name = f'{SEC_NAME.replace('截面', '')}{H}X{B}'
        print(Block_name)
    elif SEC_NAME == '双拼工字钢截面':
        format_str = int_or_float(H, 10, 1)
        if H >  180: Block_name = f'2I{format_str}a'
        if H <= 180: Block_name = f'2I{format_str}'
        print(Block_name)
    elif SEC_NAME == '双拼槽钢截面':
        format_str = int_or_float(H, 10, 1)
        if H >= 140: Block_name = f'2C{format_str}a'
        if H <  140: Block_name = f'2C{format_str}'
        print(Block_name)
    elif SEC_NAME == '双拼HM截面' or SEC_NAME == '双拼HN截面':
        C = SEC_CAD[SEC_NAME]['C'] 
        format_str_1 = int(H)
        format_str_2 = int(C)
        Block_name = f'2{SEC_NAME[2:4]}{format_str_1}X{format_str_2}'
        print(Block_name)
    return Block_name

# 钻孔平台参数汇总
def PlatformUI_CSV(app):
    # 贝雷梁伸出
    component_tab = app.tab_components
    bailey_Cantilever_length = float(component_tab.entry_pile_cantilever.get())

    # 小肋
    rib_spacing = component_tab.entry_smallrib.get()
    rib_x_expr = midasdisttolst2(0, rib_spacing)
    rib_along_layout_gaplst = group_consecutive_elements(get_distance_from_xlst(rib_x_expr))

    # 贝雷梁
    mct_lst, xl_single_nodes, single_start_nodes, distribute_single_nodes,pile_group_nodes,xl_section_cad,xl_section_name,distribute_section_cad, distribute_section_name,gz_section_cad, gz_section_name,section_weight_per_meter_lst,Feng_values,Feng_result_lst, Shui_values, Shui_lst,single_csv_nodes = build_mct_from_ui (app)
    print(single_csv_nodes)

    row1 = single_csv_nodes[1]
    z_min = min(pt[2] for pt in row1)
    bailey_x_lst = [pt[0] for pt in row1 if abs(pt[2] - z_min) < 1e-6]
    bailey_y_lst = [pts[0][1] for pts in single_csv_nodes.values()]
    print(bailey_x_lst)

    bailey_x_append_lst = get_distance_from_xlst(bailey_x_lst)
    bailey_y_append_lst = get_distance_from_xlst(bailey_y_lst)

    bailey_along_layout_namelst = get_bailey_along_layout_namelst(bailey_x_append_lst)
    bailey_along_layout_gaplst = group_consecutive_elements(bailey_x_append_lst)
    bailey_cross_layout_gaplst = group_consecutive_elements(bailey_y_append_lst)
    print(bailey_along_layout_namelst)
    print(bailey_along_layout_gaplst)

    # 钢管桩
    pile_x = component_tab.entry_pile_x.get()
    pile_y = component_tab.entry_pile_y.get()
    pile_x_expr = midasdisttolst2(0, pile_x)
    pile_y_expr = midasdisttolst2(0, pile_y)
    ganggz_along_layout_gaplst = group_consecutive_elements(get_distance_from_xlst(pile_x_expr))
    ganggz_cross_layout_gaplst = group_consecutive_elements(get_distance_from_xlst(pile_y_expr))

    #截面
    rib_sec_name = SEC_trans(xl_section_cad, xl_section_name)
    distribute_sec_name = SEC_trans(distribute_section_cad, distribute_section_name)

    return bailey_Cantilever_length, rib_sec_name,rib_along_layout_gaplst, bailey_along_layout_namelst, bailey_along_layout_gaplst, bailey_cross_layout_gaplst,  distribute_sec_name, ganggz_along_layout_gaplst, ganggz_cross_layout_gaplst

def CSV_format(app):
    CSV_lines = CSV_DrillingPlatform(app)
    write_csv(csv_filename2, CSV_lines)