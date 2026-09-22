"""
mct_format.py
生成完整 MCT 文本段，用于 Midas Civil 导入。
依赖：mct_test.py 的 build_all_sections_mct / build_materials / 节点生成函数等。
"""

# 3. 本地模块
from FEM_MidasCivil.Steel_Pipe_Bailey_Platform_FEA.Platform_MCT_pre_FEM import (
    build_materials, build_all_sections_mct, collect_all_section_cad, build_small_rib_nodes, build_bailey_top_nodes, build_bailey_bottom_nodes,
    build_bailey_diaver_nodes, collect_all_bailey_nodes, build_distribute_nodes, build_pile_nodes, build_connect_x_nodes, build_casing_nodes,
    build_small_rib_elements, build_bailey_top_elements, build_bailey_bottom_elements, build_bailey_vertical_elements, group_bailey_xz_panels_v2, 
    build_bailey_diagonal_elements, build_distribute_elements, build_pile_elements, build_connect_x_elements, build_casing_elements, build_plate_elements,
    build_structure_groups, build_elastic_links, build_FRAME_lst, build_constraint_lines, get_drill_coords, build_stldcase, build_wind_load, build_water_force_load, 
    build_pnloadtype, build_planeload
)


def build_mct_from_ui(app):
    """
    从 UI 界面获取数据，生成 MCT 格式字符串。
    """
    # -----------材料-----------
    def build_MATERIAL_lst():
        """
        读取 mct_test.build_materials() 返回值，
        格式化为符合 MCT 语法的材料定义行。
        返回：
            MATERIAL_lst = [
                "1, STEEL, Q235, 0, 0, , C, NO, 0.02, 1, GB03(S), , Q235, NO, 206",
                "2, STEEL, 贝雷梁, 0, 0, , C, NO, 0.02, 1, JTJ(S), , 16Mn, NO, 210",
            ]
        """
        # 从 mct_test 读取材料定义（返回列表，每项是一行字符串）
        materials = build_materials()

        # 这里可以根据你的版本号或导出习惯替换参数
        MATERIAL_lst = []
        for i, line in enumerate(materials, start=1):
            MATERIAL_lst.append(line.strip())
        return MATERIAL_lst
    MATERIAL_lst = build_MATERIAL_lst()

    # -----------截面-----------
    def build_SECTION_lst(all_sections: dict):
        """
        读取 mct_test.build_all_sections_mct() 的输出，
        直接生成 SECTION_lst，用于后续 mct 集成。
        """
        section_lines, name_to_id = build_all_sections_mct(all_sections)
        SECTION_lst = section_lines   
        return SECTION_lst
    all_sections =  app.tab_section.get_all_section_data()
    SECTION_lst = build_SECTION_lst(all_sections)
    _, name_to_id = build_all_sections_mct(all_sections)

    names, cad_params, weights = collect_all_section_cad(all_sections)
    def convert_section_data_for_cad(names, cad_params, weights):
        """
        将 collect_all_section_cad() 的输出（7个name、7个param、5个weight）
        按顺序展开为 CAD 绘图函数可直接解包的形式。
        """
        def safe_get(seq, idx, default):
            return seq[idx] if idx < len(seq) else default

        xl_section_cad         = safe_get(cad_params, 0, {})
        xl_section_name        = safe_get(names, 0, None)
        distribute_section_cad = safe_get(cad_params, 1, {})
        distribute_section_name= safe_get(names, 1, None)
        gz_section_cad         = safe_get(cad_params, 3, {})
        gz_section_name        = safe_get(names, 3, None)

        section_weight_per_meter_lst = list(weights) 
        return (
            xl_section_cad, xl_section_name,
            distribute_section_cad, distribute_section_name,
            gz_section_cad, gz_section_name,
            section_weight_per_meter_lst
        )
    (
    xl_section_cad, xl_section_name,
    distribute_section_cad, distribute_section_name,
    gz_section_cad, gz_section_name,
    section_weight_per_meter_lst
    ) = convert_section_data_for_cad(names, cad_params, weights)
    print(xl_section_name)
    print(xl_section_cad)
    print(distribute_section_name)
    print(distribute_section_cad)
    print(gz_section_name)
    print(gz_section_cad)
    print(section_weight_per_meter_lst)
    
    # -----------板厚-----------
    plate_thickness = app.tab_components.entry_deck.get() or 10
    THICKNESS_line = f"1,VALUE,1,YES,{plate_thickness},0,NO,0,0"

    print("板厚生成完成")
    # -----------节点-----------
    node_id = 1
    # 获取全部节点数据
    # 小肋
    components_tab = app.tab_components
    section_data = app.tab_section.get_all_section_data()   
    (xl_nodes, xl_single_nodes, xl_bailey_elink_nodes, xl_length, node_id               
        ) = build_small_rib_nodes(components_tab)
    print("小肋节点生成完成")
    # 贝雷梁
    (
        bailey_top_nodes,            # 全部节点
        bailey_top_elink_nodes,      # 与小肋弹连
        bailey_top_vertical_nodes,   # 竖杆共节点
        bailey_top_diagonal_nodes,   # 斜杆共节点
        bailey_top_release_nodes,    # 释放约束节点
        bailey_top_start_nodes,      # 起始端点
        bailey_top_end_nodes,        # 终止端点
        bailey_top_single_nodes,     # 每根贝雷梁节点归类
        z_top,
        node_id
    ) = build_bailey_top_nodes(
        components_tab,
        section_data,
        xl_bailey_elink_nodes,
        node_id
        )
    
    (
        bailey_bottom_nodes,
        bailey_bottom_elink_nodes,       # 与分配梁弹连节点
        bailey_bottom_vertical_nodes,    # 竖杆共节点
        bailey_bottom_diagonal_nodes,    # 斜杆共节点
        bailey_bottom_release_nodes,     # 释放约束节点
        bailey_bottom_single_nodes,      # 每根贝雷梁节点归类
        bailey_bottom_start_nodes,       # 起始端点
        bailey_bottom_end_nodes,         # 终止端点
        z_bottom,
        node_id
    ) = build_bailey_bottom_nodes(
        components_tab,
        bailey_top_vertical_nodes,
        bailey_top_diagonal_nodes,
        bailey_top_release_nodes,
        z_top,
        node_id
        )
    
    bailey_diaver_nodes, bailey_diaver_single_nodes, node_id = build_bailey_diaver_nodes(bailey_top_vertical_nodes, z_top, node_id)
    
    (
        all_nodes_bailey,
        single_nodes,
        release_nodes,
        start_nodes,
        single_start_nodes,
        vertical_nodes,
        diagonal_nodes,
        single_csv_nodes,
    ) =collect_all_bailey_nodes(
        bailey_top_nodes,
        bailey_bottom_nodes,
        bailey_diaver_nodes,
        bailey_top_single_nodes,
        bailey_bottom_single_nodes,
        bailey_diaver_single_nodes,
        bailey_top_release_nodes,
        bailey_bottom_release_nodes,
        bailey_top_diagonal_nodes,
        bailey_bottom_diagonal_nodes,
        bailey_top_start_nodes,
        bailey_bottom_start_nodes,
        bailey_top_end_nodes,
        bailey_bottom_end_nodes
    )

    # 分配梁
    (
        distribute_nodes,
        distribute_single_nodes,
        distribute_bottom_elink_nodes,
        distribute_pile_elink_nodes,
        node_id
    ) = build_distribute_nodes(
        components_tab,
        bailey_bottom_elink_nodes,
        xl_length,
        z_bottom,
        node_id,
        section_data=section_data
    )

    # 钢管桩
    (
        pile_top_nodes, 
        pile_connect_nodes,
        pile_bottom_nodes, 
        pile_group_nodes,
        water_force_points_by_y, node_id
    ) = build_pile_nodes(
        components_tab,
        distribute_pile_elink_nodes, 
        node_id,
        section_data,
        load_tab=app.tab_loads
        )

    # 联结系
    connect_x_nodes, connect_groups, node_id = build_connect_x_nodes(
        components_tab,
        pile_group_nodes,
        pile_connect_nodes,        
        node_id
        )

    # 钢护筒
    casing_nodes, casing_bottom_nodes, node_id = build_casing_nodes(app.tab_components, app.tab_drillarea, node_id)


    # 节点数据汇总
    def _flatten_group_nodes(group_dict: dict) -> dict:
        """{组号: [(nid,coord), ...]} -> {nid: coord}"""
        flat = {}
        if not group_dict:
            return flat
        for _, lst in group_dict.items():
            for nid, coord in lst:
                flat[nid] = coord
        return flat

    pile_nodes_flat        = _flatten_group_nodes(pile_group_nodes)
    connect_nodes_flat     = _flatten_group_nodes(connect_groups)

    all_nodes = (
    xl_nodes
    | all_nodes_bailey
    | distribute_nodes
    | pile_nodes_flat
    | connect_nodes_flat
    | casing_nodes
    )

    def format_nodes_mct(nodes: dict) -> str:
        """
        输入：dict{id:(x,y,z)}
        输出：MCT格式节点段
        """
        lines = []
        for node_id in sorted(nodes.keys()):
            x, y, z = nodes[node_id]
            lines.append(f"{node_id}, {x}, {y}, {z}")
        return lines
    NODE_lst = format_nodes_mct(all_nodes)
    print("节点生成完成")

    # -----------单元-----------
    elem_id = 1
    # 小肋
    rib_elems, elem_id = build_small_rib_elements(
        components_tab, section_data, name_to_id, xl_single_nodes, elem_id
    )
    # 贝雷梁上弦杆
    top_elems, elem_id = build_bailey_top_elements(
        section_data, name_to_id, bailey_top_single_nodes, elem_id
    )
    # 下弦杆
    bottom_elems, elem_id = build_bailey_bottom_elements(
        section_data, name_to_id, bailey_bottom_single_nodes, elem_id
    )
    # 竖杆
    vertical_elements, elem_id = build_bailey_vertical_elements(
        bailey_top_vertical_nodes, bailey_bottom_vertical_nodes,
        section_data, name_to_id, elem_id
    )
    # 斜杆（先分组再成单元）
    groups = group_bailey_xz_panels_v2(
        bailey_top_diagonal_nodes, bailey_bottom_diagonal_nodes, bailey_diaver_nodes
    )
    diagonal_elements, elem_id = build_bailey_diagonal_elements(
        groups, section_data, name_to_id, elem_id
    )
    # 分配梁（含制动墩 F1/F2）
    distribute_elements, elem_id = build_distribute_elements(
        components_tab, section_data, name_to_id,
        distribute_single_nodes, elem_id
    )
    # 钢管桩
    pile_elements, elem_id = build_pile_elements(
        section_data, name_to_id,
        pile_group_nodes, elem_id
    )

    # 联结系
    connect_elements, elem_id = build_connect_x_elements(
        section_data, name_to_id, connect_groups,  elem_id
    )

    casing_elements, elem_id = build_casing_elements(
        section_data, name_to_id, casing_nodes, elem_id
    )

    # 板单元
    plate_elements = build_plate_elements(
        xl_nodes, elem_id
    )
    print(plate_elements)


    def build_ELEMENT_lst(
        rib_elems,
        top_elems,
        bottom_elems,
        vertical_elements,
        diagonal_elements,
        distribute_elements,
        pile_elements,
        connect_elements,
        casing_elements,
        plate_elements
    ):
        """
        输入：
            各构件 build_xxx_elements() 返回的列表（每个元素均为完整 MCT 行字符串，如 "1, BEAM, 1, 2, 1, 2, 0.0, 0"）

        功能：
            合并全部单元行 → 按单元号排序 → 返回 ELEMENT_lst（不含表头）
        """
        all_elems = (
            rib_elems
            + top_elems
            + bottom_elems
            + vertical_elements
            + diagonal_elements
            + distribute_elements
            + pile_elements
            + connect_elements
            + casing_elements
            + plate_elements
        )
        def extract_id(line: str):
            try:
                return int(line.split(",")[0].strip())
            except Exception:
                return 0

        sorted_elems = sorted(all_elems, key=extract_id)
        ELEMENT_lst = sorted_elems

        return ELEMENT_lst
    
    ELEMENT_lst = build_ELEMENT_lst(
        rib_elems,
        top_elems,
        bottom_elems,
        vertical_elements,
        diagonal_elements,
        distribute_elements,
        pile_elements,
        connect_elements,
        casing_elements,
        plate_elements
    )   
    print("单元生成完成")

    # -----------组-----------
    group_lines = build_structure_groups(
        xl_nodes,
        all_nodes,
        bailey_top_nodes,
        bailey_bottom_nodes,
        vertical_nodes,
        diagonal_nodes,
        distribute_nodes,
        pile_group_nodes,
        connect_groups,
        rib_elems,
        top_elems,
        bottom_elems,
        vertical_elements,
        diagonal_elements,
        distribute_elements,
        pile_elements,
        connect_elements,
        )   
    GROUP_lst = group_lines
    print("结构组生成完成")

    # -----------约束-----------
    # 弹性连接
    elink_lines = build_elastic_links( 
        components_tab,
        # 1) 小肋-上弦
        xl_bailey_elink_nodes,           # {rib_idx: [nid, ...]}
        bailey_top_elink_nodes,      # {nid: (x,y,z)}
        # 2) 下弦-分配梁
        bailey_bottom_elink_nodes,    # {nid: (x,y,z)}
        distribute_bottom_elink_nodes,# {nid: (x,y,z)}
        # 3) 分配梁-钢管桩
        distribute_pile_elink_nodes,    # {nid: (x,y,z)}
        pile_top_nodes,                 # {nid: (x,y,z)}
        start_link_id = 1
        )
    ELASTICLINK_lst = elink_lines
    print("弹性连接生成完成")
    
    # 释放梁端约束
    found_in_I = found_in_J = 0
    for line in ELEMENT_lst:
        if "BEAM" not in line:
            continue
        parts = [x.strip() for x in line.split(",")]
        node1, node2 = int(parts[4]), int(parts[5])
        for rn in release_nodes:
            if rn == node1:
                found_in_I += 1
            if rn == node2:
                found_in_J += 1

    FRAME_lst = build_FRAME_lst(ELEMENT_lst, release_nodes)
    print("释放梁端约束生成完成")
    
    # 桩底固结
    CONSTRAINT_lst = build_constraint_lines(pile_bottom_nodes,casing_bottom_nodes)
    print("桩底固结生成完成")
    print("约束生成完成")

    # -----------荷载-----------
    # 荷载说明
    load_tab = app.tab_loads
    pile_coords = get_drill_coords(app.tab_components, app.tab_drillarea)
    STLD_lst = build_stldcase(pile_coords)
    print("荷载组生成完成")

    # 风荷载
    print("[调试] 桩顶节点号:", list(pile_top_nodes.keys()))
    WIND_GROUP_LINES,Feng_values,Feng_result_lst = build_wind_load(load_tab, components_tab,pile_top_nodes=pile_top_nodes)
    print("风荷载生成完成")

    # 水流力
    WATER_FORCE_LINE, Shui_values, Shui_lst= build_water_force_load(load_tab, section_data, components_tab, water_force_points_by_y)
    print("水流力生成完成")

    # 荷载组合
    # LOADCOMB_LINES = build_loadcomb_lines(load_tab)s
    # print("荷载组合生成完成")
    
    # 分配面荷载类型
    PNLOADTYPE = build_pnloadtype(load_tab)

    # 分配面荷载
    PLANELOAD = build_planeload (load_tab, pile_coords)


    mct_lst = [
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
                ["*BNDR-GROUP", "弹性连接, 0", "桩底反力, 0", "释放梁端约束, 0"],
                # 弹性连接
                ["*ELASTICLINK"] + ELASTICLINK_lst,
                # 释放梁端约束
                ["*FRAME-RLS"] + FRAME_lst,
                # 桩底固结
                ["*CONSTRAINT"] + CONSTRAINT_lst,
                # 荷载说明
                ["*STLDCASE", "自重, USER,", "风荷载, USER,", "水流力, USER,"] + STLD_lst,
                # 自重
                ["*USE-STLD, 自重", "*SELFWEIGHT", " 0, 0, -1"],
                # 风荷载
                WIND_GROUP_LINES,
                # 水流力
                WATER_FORCE_LINE,
                # 荷载组合
                # LOADCOMB_LINES,
                # 板厚
                ["*THICKNESS",THICKNESS_line],
                # 分配面荷载类型
                PNLOADTYPE,
                # 分配面荷载
                ["*PLANELOAD"] + PLANELOAD,
            ]

    return mct_lst, xl_single_nodes, single_start_nodes, distribute_single_nodes,pile_group_nodes,xl_section_cad,xl_section_name,distribute_section_cad, distribute_section_name,gz_section_cad, gz_section_name,section_weight_per_meter_lst,Feng_values,Feng_result_lst, Shui_values, Shui_lst,single_csv_nodes


def build_TXT_from_mct(app):
    mct_lst, xl_single_nodes, single_start_nodes, distribute_single_nodes,pile_group_nodes,xl_section_cad,xl_section_name,distribute_section_cad, distribute_section_name,gz_section_cad, gz_section_name,section_weight_per_meter_lst,Feng_values,Feng_result_lst, Shui_values, Shui_lst,single_csv_nodes= build_mct_from_ui(app)

    print("TXT生成开始")
    print("调试信息")
    print(Feng_values)
    print(Feng_result_lst)
    print(Shui_values)
    print(Shui_lst) 
    

    TXT_lst =[
        ["*风荷载规范","无"],
        ["*水流力规范","无"],
        ["*自重参数","-1"],
        ["*混凝土荷载参数","26.5"],
        ["*模板荷载参数","2.5"],
        ["*施工荷载参数","2.0"],
        ['*风荷载参数', ', '.join(str(y) for x in [Feng_values, Feng_result_lst] for y in x )],
        ['*水流力参数', ', '.join(str(y) for x in [Shui_values, Shui_lst] for y in x )],
        ["*贝雷梁材料特性参数","None"],
        ["*钢材材料特性参数","None"],
        ["*混凝土材料特性参数","None"],
        ["*钢筋材料特性参数","None"],
        ["*螺栓材料特性参数","None"],
        ["*焊缝材料特性参数","None"],
        ["*竹胶板材料特性参数","None"],
        ["*方木材料特性参数","None"]
    ]

    return TXT_lst