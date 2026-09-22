# 3. 本地模块
from General.Midas import num_lst_to_num_group
from General.Formula import cal_FengYa_JTG_T_3360_01_2018
from General.DataUtils import midasdisttolst, midasdisttolst2, inear_interpolation


# ==========生成材料信息==========
def build_materials():
    """
    生成 MCT 材料定义段（MATERIAL 部分）
    默认包含两种材料：
        1 - 16Mn 钢（JTJ 规范）
        2 - Q235 钢（GB 50917-13）
    输出格式完全符合 MCT 语法。
    """
    materials = []

    materials.append(
        "1, STEEL, 16Mn              , 0, 0, , F, NO, 0.02, 1, JTJ(S)     ,            , 16Mn          , NO, 210000"
    )

    materials.append(
        "2, STEEL, Q235              , 0.056872, 0, , F, NO, 0.02, 1, GB 50917-13(S),            , Q235          , NO, 206000"
    )

    return materials


# ==========生成截面信息==========
def build_all_sections_mct(all_sections: dict, start_index: int = 1):
    """
    根据UI收集的截面数据，生成完整的MCT *DGN-SECT命令行。
    通过截面名称自动识别双拼截面，追加扩展参数段。
    输出全部为整数形式（无小数点）。
    自动追加贝雷梁弦杆、竖杆、斜杆三种固定截面。
    """
    section_lines = []
    sec_id = start_index
    name_to_id = {}

    mapping = {
        "工字钢截面": ("H", ["H(mm)", "B(mm)", "tw(mm)", "tf(mm)", None, None,  "r(mm)", None, None, None]),
        "槽钢截面": ("C", ["H(mm)", "B(mm)", "tw(mm)", "tf(mm)", "B(mm)", "tf(mm)", None, None, None, None]),
        "HM截面": ("H", ["H(mm)", "B(mm)", "tw(mm)", "tf(mm)", None, None, None, "r(mm)", None, None, None]),
        "HN截面": ("H", ["H(mm)", "B(mm)", "tw(mm)", "tf(mm)", None, None, None, "r(mm)", None, None, None]),
        "双拼工字钢截面": ("B", ["H(mm)", "B(mm)", "tw(mm)", "tf(mm)", "C(mm)", None, None, None, None, None]),
        "双拼槽钢截面": ("B", ["H(mm)", "B(mm)", "tw(mm)", "tf(mm)", "C(mm)", None, None, None, None, None]),
        "双拼HM截面": ("B", ["H(mm)", "B(mm)", "tw(mm)", "tf(mm)", "C(mm)", None, None, None, None, None]),
        "双拼HN截面": ("B", ["H(mm)", "B(mm)", "tw(mm)", "tf(mm)",  "C(mm)", None, None, None, None, None]),
        "管型截面": ("P", ["D(mm)", "d(mm)", None, None, None, None, None, None, None, None]),
    }

    # -------- 1. 读取 UI 中定义的截面 --------
    for comp_name, sec_info in all_sections.items():
        sec_type = sec_info.get("section_type", "")
        params = sec_info.get("params", {})

        rule = mapping.get(sec_type)
        if not rule:
            print(f"未识别截面类型：{sec_type}")
            continue
        shape_code, key_order = rule

        values = []
        for key in key_order:
            if key is None:
                values.append("0")
            else:
                try:
                    v = float(params.get(key, 0))
                    formatted = f"{v:.1f}"
                    values.append(formatted)
                except ValueError:
                    values.append("0")

        base_prefix = f"{sec_id}, DBUSER, {comp_name}, CC, 0, 0, 0, 0, 0, 0, YES, NO, {shape_code}, 2, "
        param_part = ", ".join(values)

        # 是否双拼
        if "双拼" in sec_type:
            extra_part = ", NO, 0, NO, 0, 0"
            mct_line = base_prefix + param_part + extra_part
        else:
            mct_line = base_prefix + param_part

        section_lines.append(mct_line)
        name_to_id[comp_name] = sec_id 
        sec_id += 1

    # -------- 2. 自动补充贝雷梁固定截面 --------
    # 参数字符串直接按MCT格式
    fixed_sections = {
        "贝雷梁弦杆": ", CC, 0, 0, 0, 0, 0, 0, YES, NO, 2C, 2, 100, 48, 5, 8.5, 80, 0, 0, 0, 0, 0",
        "贝雷梁竖杆": ", CC, 0, 0, 0, 0, 0, 0, YES, NO, H, 2, 80, 50, 6.5, 4.5, 0, 0, 0, 0, 0, 0",
        "贝雷梁斜杆": ", CC, 0, 0, 0, 0, 0, 0, YES, NO, H, 2, 80, 50, 6.5, 4.5, 0, 0, 0, 0, 0, 0",
    }

    for sec_name, param in fixed_sections.items():
        mct_line = f"{sec_id}, DBUSER, {sec_name}{param}"
        section_lines.append(mct_line)
        name_to_id[sec_name] = sec_id
        sec_id += 1

    # -------- 3. 清理输出 --------
    fixed_section_lines = []
    for item in section_lines:
        if isinstance(item, list):
            fixed_section_lines.append(", ".join(str(x) for x in item))
        else:
            fixed_section_lines.append(str(item))

    if not isinstance(name_to_id, dict):
        name_to_id = dict(name_to_id)
    return section_lines, name_to_id

def get_section_param_dict(all_sections: dict, section_params_template: dict):
    """
    根据 UI 收集的 all_sections（含参数字典），
    生成 CAD 可用的简化结构：
    {
        "小肋": {"工字钢截面": {"H": 500, "B": 200, ...}},
        "分配梁": {"槽钢截面": {...}},
        ...
    }
    """
    section_cad_dict = {}

    # 建立通用键名映射：去掉(mm)
    def normalize_key(k):
        return k.replace("(mm)", "")

    for comp_name, sec_info in all_sections.items():
        sec_type = sec_info.get("section_type", "")
        params = sec_info.get("params", {})

        # 取出模板中该类型允许的字段
        template_keys = section_params_template.get(sec_type, [])
        valid_keys = [normalize_key(k) for k in template_keys]

        # 转换：中文键 → 简短键名，跳过空值
        result = {}
        for k, v in params.items():
            nk = normalize_key(k)
            if nk in valid_keys and v != "":
                try:
                    result[nk] = float(v)
                except ValueError:
                    result[nk] = v

        section_cad_dict[comp_name] = {sec_type: result}

    return section_cad_dict


def collect_all_section_cad(all_sections: dict):
    """
    一步生成 CAD 可用的截面参数字典：
    输出结构：
      names: tuple[str]
      cad_params: tuple[dict]，形如 {'工字钢截面': {'H':120, 'B':64, 'tw':4.8, 'tf':7.3, 'r':7.5}}
      weights: tuple[float]，仅含前五项（不含钢管桩类）
    """

    comps = [
        "小肋", "分配梁", "联结系",
        "钢管桩", "制动墩钢管桩",
        "制动墩分配梁F1", "制动墩分配梁F2"
    ]

    # 输出容器
    names, cad_params, weights = [], [], []

    for comp in comps:
        sec_info = all_sections.get(comp, {})
        sec_type = sec_info.get("section_type")
        params = sec_info.get("params", {})

        if sec_type:
            clean_dict = {}
            for k, v in params.items():
                # 跳过无效或非数值型参数
                if any(ignore in k for ignore in ("每延")) or not v:
                    continue
                key = k.replace("(mm)", "").replace(" ", "").strip()
                try:
                    clean_dict[key] = float(v)
                except ValueError:
                    clean_dict[key] = v

            cad_params.append({sec_type: clean_dict})
            names.append(sec_type)
        else:
            cad_params.append({})
            names.append(None)

        try:
            w = float(params.get("每延米重 (kg/m)", 0.0))
        except ValueError:
            w = 0.0
        weights.append(w)

    # 去掉钢管桩类重量
    weights = [w for i, w in enumerate(weights) if i not in (3, 4)]
    while len(weights) < 5:
        weights.append(0)

    return tuple(names), tuple(cad_params), tuple(weights)
 
# ------------------------------------------------
# 节点参数
# ------------------------------------------------

# ==========生成小肋节点信息==========
def build_small_rib_nodes(components_tab):
    """
    从UI读取参数生成小肋节点。
    小肋：每根沿y方向布置（x相同）。节点包括两端、贝雷梁交点

    返回：
        xl_nodes                —— {nid: (x, y, z)}                   # 全部小肋节点
        xl_single_nodes         —— {rib_idx: [(nid, (x, y, z)), ...]} # 每根小肋节点
        xl_bailey_elink_nodes   —— {nid: (x, y, z)}                   # 小肋-贝雷梁交点
        xl_length               —— 单根小肋总长度
        node_id                 —— 下一个可用节点号
    """
    # -------- 1. 读取UI参数 --------
    smallrib_expr = components_tab.entry_smallrib.get().strip()         # 例如 "29@400"
    bailey_expr = components_tab.entry_bailey_spacing.get("1.0", "end").strip()
    xl_length = float(components_tab.entry_smallrib_length.get() or 13100.0)  # 小肋长度
    cantilever_len = float(components_tab.entry_bailey_init_spacing.get() or 550.0)  # 悬臂长度

    # -------- 2. 计算布置 --------
    xl_xs = midasdisttolst2(0.0, smallrib_expr)            # 每根小肋的x坐标
    bailey_ys = midasdisttolst2(cantilever_len, bailey_expr)  # 贝雷梁y坐标

    # -------- 3. 初始化容器 --------
    xl_nodes = {}                # 全部节点
    xl_single_nodes = {}         # 每根小肋的节点
    xl_bailey_elink_nodes = {}   # 小肋-贝雷梁弹连点

    node_id = 1

    # -------- 4. 生成每根小肋 --------
    for rib_idx, x in enumerate(xl_xs, start=1):
        xl_single_nodes[rib_idx] = []

        y_endpoints = [0.0, xl_length]
        y_mid = xl_length / 2
        y_all = sorted(set(y_endpoints + bailey_ys + [y_mid]))

        for y in y_all:
            coord = (x, y, 0.0)
            xl_nodes[node_id] = coord
            xl_single_nodes[rib_idx].append((node_id, coord))
            # 贝雷梁弹连点
            if any(abs(y - yy) < 1e-6 for yy in bailey_ys):
                xl_bailey_elink_nodes[node_id] = coord

            node_id += 1

    return xl_nodes, xl_single_nodes, xl_bailey_elink_nodes, xl_length, node_id

# ==========生成贝雷梁上弦杆节点信息==========
def build_bailey_top_nodes(
    components_tab,
    section_data: dict,
    xl_bailey_elink_nodes: dict,
    start_id: int = 1
):
    # --------------------------------------
    # 1. 基本参数读取
    # --------------------------------------
    bailey_span_expr = components_tab.entry_bailey_span.get().strip()
    bailey_span_steps_m = midasdisttolst(bailey_span_expr)
    bailey_span_steps = [d * 1000.0 for d in bailey_span_steps_m] # 单位mm
    bailey_span_total = sum(bailey_span_steps) 
    
    init_spacing = float(components_tab.entry_smallrib_init_spacing.get() or 0.0)

    x0 = 0.0 - init_spacing
    x_end = x0 + bailey_span_total
    panel_len = 3000.0
    n_panels = int((x_end - x0) // panel_len) + 1

    # --------------------------------------
    # 2. z 坐标定义
    # --------------------------------------
    H_smallrib = float(section_data.get("小肋", {}).get("params", {}).get("H(mm)", 120.0))
    H_bailey = float(section_data.get("贝雷梁", {}).get("params", {}).get("H(mm)", 100))
    z_top = 0.0 - 0.5 * H_smallrib - 0.5 * H_bailey
    print(z_top)

    # --------------------------------------
    # 3. 从小肋节点提取弹连坐标
    # --------------------------------------
    elink_x_list = sorted({coord[0] for coord in xl_bailey_elink_nodes.values()})
    elink_y_list = sorted({coord[1] for coord in xl_bailey_elink_nodes.values()})

    # --------------------------------------
    # 4. 生成各类别 X 坐标模板
    # --------------------------------------
    # (1) 两端端点节点
    end_x_list = [x0, x_end]

    # (2) 与竖杆共节点
    vertical_x_list = []
    first_panel_offsets_vertical = [90.0, 1500.0, 2910.0]
    later_panel_offsets_vertical = [90.0, 1500.0, 2910.0]
    for p in range(n_panels):
        base = x0 + p * panel_len
        if base > x_end + 1e-6:
            break
        offsets = first_panel_offsets_vertical if p == 0 else later_panel_offsets_vertical
        for off in offsets:
            xx = base + off
            if xx <= x_end + 1e-6:
                vertical_x_list.append(xx)

    # (3) 与斜杆共节点
    diagonal_x_list = []
    first_panel_offsets_diagonal = [795.0, 2205.0]
    later_panel_offsets_diagonal = [795.0, 2205.0]
    for p in range(n_panels):
        base = x0 + p * panel_len
        if base > x_end + 1e-6:
            break
        offsets = first_panel_offsets_diagonal if p == 0 else later_panel_offsets_diagonal
        for off in offsets:
            xx = base + off
            if xx <= x_end + 1e-6:
                diagonal_x_list.append(xx)

    # (4) 片间释放节点
    release_x_list = []
    k = 1
    while x0 + 3000 * k < x_end - 1e-6:
        release_x_list.append(x0 + 3000 * k)
        k += 1
    print(release_x_list)


    # 合并模板去重
    end_x_set = set(end_x_list)
    elink_x_set = set(elink_x_list)
    vertical_x_set = set(vertical_x_list)
    diagonal_x_set = set(diagonal_x_list)
    release_x_set = set(release_x_list)

    x_union_template = sorted(
        set().union(end_x_set, elink_x_set, vertical_x_set, diagonal_x_set, release_x_set)
    )

    coords_all_set = set()
    coords_end_set = set()
    coords_elink_set = set()
    coords_vertical_set = set()
    coords_diagonal_set = set()
    coords_release_set = set()

    for y in elink_y_list:
        for x in x_union_template:
            coords_all_set.add((x, y, z_top))
        for x in end_x_set:
            coords_end_set.add((x, y, z_top))
        for x in elink_x_set:
            coords_elink_set.add((x, y, z_top))
        for x in vertical_x_set:
            coords_vertical_set.add((x, y, z_top))
        for x in diagonal_x_set:
            coords_diagonal_set.add((x, y, z_top))
        for x in release_x_set:
            coords_release_set.add((x, y, z_top))

    bailey_top_nodes = {}
    bailey_top_start_nodes = {}
    bailey_top_end_nodes = {}
    bailey_top_elink_nodes = {}
    bailey_top_vertical_nodes = {}
    bailey_top_diagonal_nodes = {}
    bailey_top_release_nodes = {}
    bailey_top_single_nodes = {}   

    node_id = start_id
    for j, y in enumerate(elink_y_list, start=1):
        bailey_top_single_nodes[j] = []  # 每根贝雷梁独立记录
        for x in x_union_template:
            coord = (x, y, z_top)
            bailey_top_nodes[node_id] = coord
            bailey_top_single_nodes[j].append((node_id, coord))  # ← 归入单根
            if abs(x - x0) < 1e-6:     # ← 左端起点
                bailey_top_start_nodes[node_id] = coord
            if abs(x - x_end) < 1e-6:  # ← 右端终点
                bailey_top_end_nodes[node_id] = coord
            if x in elink_x_list:
                bailey_top_elink_nodes[node_id] = coord
            if x in vertical_x_list:
                bailey_top_vertical_nodes[node_id] = coord
            if x in diagonal_x_list:
                bailey_top_diagonal_nodes[node_id] = coord
            if x in release_x_list:
                bailey_top_release_nodes[node_id] = coord
            node_id += 1

    # --------------------------------------
    # 6. 返回
    # --------------------------------------
    return (
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
    )

# ==========生成贝雷梁下弦杆节点信息==========
def build_bailey_bottom_nodes(
    components_tab,
    bailey_top_vertical_nodes: dict,
    bailey_top_diagonal_nodes: dict,
    bailey_top_release_nodes: dict,
    z_top: float,
    start_id: int = 1,
    ):
    """
    生成贝雷梁下弦杆节点坐标。
    分类：
        bailey_bottom_elink_nodes: 类型 1（与分配梁弹连点）
        bailey_bottom_vertical_nodes: 类型 2（竖杆共节点）
        bailey_bottom_diagonal_nodes: 类型 3（斜杆共节点）
        bailey_bottom_release_nodes: 类型 4（片段连接释放点）
    """

    # --------------------------------------
    # 1. 基本参数读取
    # --------------------------------------
    bailey_span_expr = components_tab.entry_bailey_span.get().strip()
    bailey_span_steps_m = midasdisttolst(bailey_span_expr)
    bailey_span_steps = [d * 1000.0 for d in bailey_span_steps_m] # 单位mm
    bailey_span_total = sum(bailey_span_steps) 
    
    init_spacing = float(components_tab.entry_smallrib_init_spacing.get() or 0.0)
    print(bailey_span_total)

    x0 = 0.0 - init_spacing
    x_end = x0 + bailey_span_total
    panel_len = 3000.0
    n_panels = int((x_end - x0) // panel_len) + 1


    pile_x_expr = components_tab.entry_pile_x.get().strip()
    pile_x_lst = midasdisttolst(pile_x_expr)

    pile_cantilever = float(components_tab.entry_pile_cantilever.get() or 3000)
    start_x = x0 + pile_cantilever

    # -------------------------------
    # 2. z 坐标上弦杆z坐标 - 1.4m
    # ------------------------------
    z_bottom = z_top - 1400

    # -------------------------------
    # 3. x坐标生成
    # ------------------------------

    # 类型1. 端点节点
    bailey_bottom_side_x = [x0, x_end]

    # 类型2. 与分配梁弹连点
    bailey_bottom_elink_x = [start_x]
    current = start_x
    for d in pile_x_lst:
        current += d
        bailey_bottom_elink_x.append(current)
   
    # 类型 345. 复用 XY 
    def shift_z(node_dict):
        return {nid: (x, y, z_bottom) for nid, (x, y, _) in node_dict.items()}

    bailey_bottom_vertical_nodes = shift_z(bailey_top_vertical_nodes)
    bailey_bottom_diagonal_nodes = shift_z(bailey_top_diagonal_nodes)
    bailey_bottom_release_nodes = shift_z(bailey_top_release_nodes)

    # 汇总节点x坐标
    x_all = sorted(
        set(
            bailey_bottom_side_x
            + bailey_bottom_elink_x
            + [coord[0] for coord in bailey_bottom_vertical_nodes.values()]
            + [coord[0] for coord in bailey_bottom_diagonal_nodes.values()]
            + [coord[0] for coord in bailey_bottom_release_nodes.values()]
        )
    )

    # --------------------------------------
    # 4. 汇总节点
    # --------------------------------------
    bailey_bottom_nodes = {}
    bailey_bottom_elink_nodes = {}
    bailey_bottom_single_nodes = {}
    bailey_bottom_start_nodes = {}
    bailey_bottom_end_nodes = {}

    # 从上弦杆任一类节点取 y 集合
    y_list = sorted({coord[1] for coord in bailey_top_vertical_nodes.values()})

    node_id = start_id
    for j, y in enumerate(y_list, start=1):
        bailey_bottom_single_nodes[j] = []
        for x in x_all:
            coord = (x, y, z_bottom)
            bailey_bottom_nodes[node_id] = coord
            bailey_bottom_single_nodes[j].append((node_id, coord))
            if x in bailey_bottom_elink_x:
                bailey_bottom_elink_nodes[node_id] = coord
            if x in [c[0] for c in bailey_bottom_vertical_nodes.values()]:
                bailey_bottom_vertical_nodes[node_id] = coord
            if x in [c[0] for c in bailey_bottom_diagonal_nodes.values()]:
                bailey_bottom_diagonal_nodes[node_id] = coord
            if x in [c[0] for c in bailey_bottom_release_nodes.values()]:
                bailey_bottom_release_nodes[node_id] = coord
            if abs(x - x0) < 1e-6:
                bailey_bottom_start_nodes[node_id] = coord
            if abs(x - x_end) < 1e-6:
                bailey_bottom_end_nodes[node_id] = coord
            node_id += 1

    # --------------------------------------
    # 5. 返回结果
    # --------------------------------------
    return (
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
    )


# ==========生成贝雷梁贝雷梁斜杆竖杆共节点信息==========
def build_bailey_diaver_nodes(
    bailey_top_vertical_nodes: dict,
    z_top: float,
    start_id: int = 1
):
    """
    生成斜杆竖杆共节点（diaver）。
    直接复制上弦杆竖杆共节点的 x,y 坐标，
    z = z_top - drop_mm
    同时整理每根贝雷梁（同一 y）对应的节点合集。
    """
    z_diaver = z_top - 700

    bailey_diaver_nodes = {}
    bailey_diaver_single_nodes = {}

    node_id = start_id
    # 先按 y 分组上弦杆竖杆共节点
    coords_by_y = {}
    for nid, (x, y, _) in bailey_top_vertical_nodes.items():
        y_key = round(y, 6)
        if y_key not in coords_by_y:
            coords_by_y[y_key] = []
        coords_by_y[y_key].append(x)

    # 排序每根贝雷梁内节点
    y_sorted = sorted(coords_by_y.keys())
    for j, y_val in enumerate(y_sorted, start=1):
        bailey_diaver_single_nodes[j] = []
        for x in sorted(coords_by_y[y_val]):
            coord = (x, y_val, z_diaver)
            bailey_diaver_nodes[node_id] = coord
            bailey_diaver_single_nodes[j].append((node_id, coord))
            node_id += 1

    return bailey_diaver_nodes, bailey_diaver_single_nodes, node_id

# ==========斜杆节点分组==========
def group_bailey_xz_panels_v2(
    bailey_top_diagonal_nodes: dict,
    bailey_bottom_diagonal_nodes: dict,
    bailey_diaver_nodes: dict
):
    """
    在同一 y 坐标内生成贝雷梁菱形组（上1-中2-下1）:
    顺序：[上弦杆diagonal，左diaver，下弦杆diagonal，右diaver]
    每组由上下同x的diagonal节点及左右最近的diaver节点组成。
    """
    from collections import defaultdict

    # 按 y 聚类
    def group_by_y(nodes):
        grouped = defaultdict(list)
        for nid, (x, y, z) in nodes.items():
            grouped[round(y, 6)].append((x, nid))
        for y in grouped:
            grouped[y].sort(key=lambda t: t[0])
        return grouped

    top_by_y = group_by_y(bailey_top_diagonal_nodes)
    bottom_by_y = group_by_y(bailey_bottom_diagonal_nodes)
    diaver_by_y = group_by_y(bailey_diaver_nodes)

    groups = {}
    group_index = 1

    # 遍历每一根贝雷梁（同一 y）
    for y in sorted(set(top_by_y.keys()) & set(bottom_by_y.keys()) & set(diaver_by_y.keys())):
        top_nodes = top_by_y[y]
        bottom_nodes = bottom_by_y[y]
        diaver_nodes = diaver_by_y[y]

        diaver_xs = [x for x, _ in diaver_nodes]

        # 建立 x→节点号 索引方便查找
        diaver_map = {round(x, 6): nid for x, nid in diaver_nodes}
        top_map = {round(x, 6): nid for x, nid in top_nodes}
        bottom_map = {round(x, 6): nid for x, nid in bottom_nodes}

        # 按上弦杆节点为基准，寻找同x下弦杆及左右最近的diaver
        for x_top, top_id in top_nodes:
            # 下弦杆是否有同x节点
            bottom_id = bottom_map.get(round(x_top, 6))
            if not bottom_id:
                continue

            # 找左、右最近 diaver
            left_candidates = [x for x in diaver_xs if x < x_top]
            right_candidates = [x for x in diaver_xs if x > x_top]
            if not left_candidates or not right_candidates:
                continue

            x_left = max(left_candidates)
            x_right = min(right_candidates)

            left_id = diaver_map.get(round(x_left, 6))
            right_id = diaver_map.get(round(x_right, 6))

            if not left_id or not right_id:
                continue

            # 形成组
            groups[group_index] = {
                "y": y,
                "x_center": x_top,
                "nodes": [top_id, left_id, bottom_id, right_id],
            }
            group_index += 1

    return groups

# ==========整理贝雷梁全部需使用节点信息==========
def collect_all_bailey_nodes(
    bailey_top_nodes: dict,
    bailey_bottom_nodes: dict,
    bailey_diaver_nodes: dict,
    bailey_top_single_nodes: dict,
    bailey_bottom_single_nodes: dict,
    bailey_diaver_single_nodes: dict,
    bailey_top_release_nodes: dict,
    bailey_bottom_release_nodes: dict,
    bailey_top_diagonal_nodes: dict,
    bailey_bottom_diagonal_nodes: dict,
    bailey_top_start_nodes: dict,
    bailey_bottom_start_nodes: dict,
    bailey_top_end_nodes: dict,
    bailey_bottom_end_nodes: dict,
):
    """
    整理贝雷梁全部节点信息：
    1. all_nodes：全部贝雷梁节点合集（上弦杆 + 下弦杆 + 斜杆竖杆共节点）
    2. single_nodes：每根贝雷梁节点合集（按 y 分组）
    3. release_nodes：全部贝雷梁单片接头节点（上、下弦杆释放点合并）
    4. single_release_nodes：每根贝雷梁单片接头节点合集（按 y 分组）
    5. vertical_nodes：全部竖杆节点合集（上弦杆竖杆共节点 + 下弦杆竖杆共节点）
    6  diagonal_nodes：全部斜杆节点合集（上弦杆斜杆共节点 + 下弦杆斜杆共节点）
    """
    # --------------------------------------------------
    # 1. 全部节点
    # --------------------------------------------------
    all_nodes = {}
    all_nodes.update(bailey_top_nodes)
    all_nodes.update(bailey_bottom_nodes)
    all_nodes.update(bailey_diaver_nodes)

    # --------------------------------------------------
    # 2. 每根贝雷梁节点
    # --------------------------------------------------
    single_nodes = {}
    single_nodes = {}
    for j in sorted(set(
        list(bailey_top_single_nodes.keys()) +
        list(bailey_bottom_single_nodes.keys()) +
        list(bailey_diaver_single_nodes.keys())
    )):
        single_nodes[j] = []
        single_nodes[j].extend(bailey_top_single_nodes.get(j, []))
        single_nodes[j].extend(bailey_bottom_single_nodes.get(j, []))
        single_nodes[j].extend(bailey_diaver_single_nodes.get(j, []))

    # --------------------------------------------------
    # 3. 全部贝雷梁释放节点
    # --------------------------------------------------
    release_nodes = {}
    release_nodes.update(bailey_top_release_nodes)
    release_nodes.update(bailey_bottom_release_nodes)

    # --------------------------------------------------
    # 4. 全部贝雷梁起始节点
    # --------------------------------------------------
    start_nodes ={}
    start_nodes.update(bailey_top_start_nodes)
    start_nodes.update(bailey_bottom_start_nodes)

    # --------------------------------------------------
    # 5. 全部贝雷梁结束节点
    # --------------------------------------------------
    end_nodes ={}
    end_nodes.update(bailey_top_end_nodes)
    end_nodes.update(bailey_bottom_end_nodes)

    # --------------------------------------------------
    # 5. 每根贝雷梁单片起始节点
    # --------------------------------------------------
    single_start_nodes = {}
    # 代表性 y（取该组第一个节点的 y 值）
    bailey_y_coords = {
        j: single_nodes[j][0][1][1]  # [(nid,(x,y,z))] -> y
        for j in single_nodes.keys() if single_nodes[j]
    }
    for j, y_val in bailey_y_coords.items():
        coords = []
        # 起点坐标
        coords.extend([
            coord for coord in start_nodes.values()
            if abs(coord[1] - y_val) < 1e-6
        ])
        
        single_start_nodes[j] = coords
        coords.extend([
            coord for coord in release_nodes.values()
            if abs(coord[1] - y_val) < 1e-6
        ])

        coords = list(dict.fromkeys(coords))
        single_start_nodes[j] = coords

    # --------------------------------------------------
    # 6. 每根贝雷梁单片起始节点+结束节点
    # --------------------------------------------------
    single_csv_nodes = {}
    # 代表性 y（取该组第一个节点的 y 值）
    bailey_y_coords = {
        j: single_nodes[j][0][1][1]  # [(nid,(x,y,z))] -> y
        for j in single_nodes.keys() if single_nodes[j]
    }
    for j, y_val in bailey_y_coords.items():
        coords = []
        # 起点坐标
        coords.extend([
            coord for coord in start_nodes.values()
            if abs(coord[1] - y_val) < 1e-6
        ])
        coords.extend([
            coord for coord in end_nodes.values()
            if abs(coord[1] - y_val) < 1e-6
        ])
        coords.extend([
            coord for coord in release_nodes.values()
            if abs(coord[1] - y_val) < 1e-6
        ])
        coords = sorted(list(dict.fromkeys(coords)), key=lambda p: p[0])
        single_csv_nodes[j] = coords

    # --------------------------------------------------
    # 7. 所有竖杆节点
    # --------------------------------------------------
    # 从 top / bottom 中提取竖杆共节点
    bailey_top_vertical_nodes = {
        nid: coord for nid, coord in bailey_top_nodes.items()
        if any(abs(coord[0] - v[0]) < 1e-6 for v in bailey_diaver_nodes.values())
    }
    bailey_bottom_vertical_nodes = {
        nid: coord for nid, coord in bailey_bottom_nodes.items()
        if any(abs(coord[0] - v[0]) < 1e-6 for v in bailey_diaver_nodes.values())
    }
    vertical_nodes = {}
    vertical_nodes.update(bailey_top_vertical_nodes)
    vertical_nodes.update(bailey_bottom_vertical_nodes)
    # --------------------------------------------------
    # 8. 所有斜杆节点
    # --------------------------------------------------
    diagonal_nodes = {}
    diagonal_nodes.update(bailey_top_diagonal_nodes)
    diagonal_nodes.update(bailey_bottom_diagonal_nodes)
    diagonal_nodes.update(bailey_diaver_nodes)

    return (
        all_nodes,
        single_nodes,
        release_nodes,
        start_nodes,
        single_start_nodes,
        vertical_nodes,
        diagonal_nodes,
        single_csv_nodes
    )
    
# ==========生成分配梁节点信息==========
def build_distribute_nodes(
    components_tab,
    bailey_bottom_elink_nodes: dict,        # 贝雷梁下弦杆与分配梁弹连节点
    xl_length: float,
    z_bottom: float = None,                 # 贝雷梁下弦杆 z 坐标
    start_id: int = 1,
    section_data: dict = None               # 截面数据
):

    # ============================================================
    # 普通分配梁
    # ============================================================
    distribute_length =float(components_tab.entry_distribute_length.get() or 1000.0) # 分配梁长度
    distribute_cantilever_bailey = float(components_tab.entry_distribute_cantilever.get() or 300.0) # 分配梁相对贝雷梁悬臂长度
    distribute_cantilever_pile = float(components_tab.entry_distribute_cantilever2.get() or 750.0) # 分配梁相对桩悬臂长度
    y_list = components_tab.entry_pile_y or 3000+3000 # 钢管桩纵向布置

    # 截面参数
    H_distributed = float(section_data.get("分配梁", {}).get("params", {}).get("H(mm)", 450))
    z_distribute = z_bottom - H_distributed / 2 - 50
    distribute_cantilever = float(components_tab.entry_distribute_length.get() or 1000.0)

    # 获取贝雷梁下弦弹连点坐标
    x_coords = sorted({coord[0] for coord in bailey_bottom_elink_nodes.values()})
    y_coords = sorted({coord[1] for coord in bailey_bottom_elink_nodes.values()})

    y_start = min(y_coords) - distribute_cantilever_bailey
    y_end =y_start + distribute_length

    # 桩排布
    pile_y_list = midasdisttolst(y_list.get())
    y_start_pile = y_start + distribute_cantilever_pile
    pile_y_coords = [y_start_pile]
    for dist in pile_y_list:
        pile_y_coords.append(pile_y_coords[-1] + dist)
    y_list = sorted(set([y_start, y_end, *pile_y_coords, *y_coords]))

    # 分配梁节点生成
    distribute_nodes = {}
    distribute_single_nodes = {}
    distribute_bottom_elink_nodes = {}
    distribute_pile_elink_nodes = {}

    node_id = start_id
    for i, x_val in enumerate(x_coords, start=1):
        distribute_single_nodes[i] = []
        for y in y_list:
            coord = (x_val, y, z_distribute)
            distribute_nodes[node_id] = coord
            distribute_single_nodes[i].append((node_id, coord))

            # 与贝雷下弦弹连
            if any(abs(y - yy) < 1e-6 for yy in y_coords):
                distribute_bottom_elink_nodes[node_id] = coord
            # 与普通桩弹连
            if any(abs(y - yp) < 1e-6 for yp in pile_y_coords):
                distribute_pile_elink_nodes[node_id] = coord
            node_id += 1

    return (
        distribute_nodes,
        distribute_single_nodes,
        distribute_bottom_elink_nodes,
        distribute_pile_elink_nodes,
        node_id
    )

    
# ==========生成钢管桩节点信息==========
def build_pile_nodes(
    components_tab,
    distribute_pile_elink_nodes: dict,
    start_id: int = 1,
    section_data: dict = None,
    load_tab=None
):
    """
    构建钢管桩节点信息（含桩顶、联结系、桩底、合力点）
    并按 y 排（桩排方向）汇总水流力合力点。
    """

    # ============================================================
    # 1、参数读取
    # ============================================================
    connect_zd = float(components_tab.entry_connect_zd.get() or 1000.0)
    connect_z = float(components_tab.entry_connect_z.get() or 2000.0)
    pile_length = float(components_tab.entry_pile_length.get()  or 15.0) * 1000.0

    # ============================================================
    # 2、初始化
    # ============================================================
    pile_group_nodes = {}
    pile_top_nodes, pile_connect_nodes, pile_bottom_nodes = {}, {}, {}
    node_id = start_id

    # 新增：按 y 排（而非 x 列）分类的水流力作用点
    water_force_points_by_y = {}

    # ============================================================
    # 3、荷载参数（单位：m）
    # ============================================================
    if load_tab:
        water_level = float(load_tab.water_level_entry.get() or 0.0)
        water_depth = float(load_tab.water_depth_entry.get() or 0.0)
    else:
        water_level = +10.0
        water_depth = 5.0

    deck_elev      = float(components_tab.entry_deck_height.get() or 0.0)
    deck_thickness = float(components_tab.entry_deck.get() or 0.0)

    # ============================================================
    # 4、截面参数（mm→m）
    # ============================================================
    def H(name):
        return float(section_data.get(name, {}).get("params", {}).get("H(mm)", 0.0))

    H_small, H_bailey, H_dist, = map(H, ["小肋", "贝雷梁", "分配梁"])

    # ============================================================
    # 5、钢管桩节点生成
    # ============================================================
    pile_by_x = {}
    for _, (x, y, z) in distribute_pile_elink_nodes.items():
        z_top = z - H_dist / 2
        pile_by_x.setdefault(x, []).append((x, y, z_top))

    group_id = 1
    ztop_elev = deck_elev - deck_thickness/1000 - H_small/2000 - H_bailey/1000 - 1.4 - H_dist/1000

    for x_val, arr in sorted(pile_by_x.items()):
        arr.sort(key=lambda t: t[1])  # 按 y 排序
        pile_group_nodes[group_id] = []

        for x, y, z_top in arr:
            z_conn1 = z_top - connect_zd
            z_conn2 = z_conn1 - connect_z
            z_bottom = z_top - pile_length
            z_force = z_top - (ztop_elev - water_level + 0.3 * water_depth) * 1000
            z_force = round(z_force, 1)

            cands = [
                ("桩顶", z_top),
                ("合力点", z_force),
                ("联结系上", z_conn1),
                ("联结系下", z_conn2),
                ("桩底", z_bottom)
            ]
            cands.sort(key=lambda t: t[1], reverse=True)

            for tag, z_val in cands:
                coord = (x, y, z_val)
                if tag == "桩顶":
                    pile_top_nodes[node_id] = coord
                elif "联结系" in tag:
                    pile_connect_nodes[node_id] = coord
                elif tag == "桩底":
                    pile_bottom_nodes[node_id] = coord
                elif tag == "合力点":
                    water_force_points_by_y.setdefault(y, []).append((node_id, coord))

                pile_group_nodes[group_id].append((node_id, coord))
                node_id += 1

        group_id += 1

    # ============================================================
    # 7、返回
    # ============================================================
    return (
        pile_top_nodes, pile_connect_nodes, pile_bottom_nodes, pile_group_nodes,
        water_force_points_by_y, node_id
    )

# ==========生成联结系节点信息==========
def build_connect_x_nodes(
    components_tab,
    pile_group_nodes,                 # {组号: [(nid,(x,y,z)), ...]}
    pile_connect_nodes,               # {nid: (x,y,z)}  ← 新增直接使用
    start_id: int = 1
):
    """
    从已知的“联结系节点字典”中抽取每根桩的上/下联结点，
    计算相邻两根桩（横向）或相邻两列桩（纵向）之间的 X 形中心节点。
    """

    connect_x_nodes = {}
    connect_groups = {}

    node_id = start_id
    gid_out = 1

    # ---------------------------
    # 工具：从“该组全部节点”中过滤得到“每根桩(y)的上下联结系节点”
    # 输入：group_nodes = [(nid,(x,y,z)), ...]； connect_dict = {nid:(x,y,z)}
    # 输出：pile_map_y = { y : {"x":x, "up":(nid,(x,y,z)), "dn":(nid,(x,y,z))} }
    # ---------------------------
    def build_pile_map_by_y(group_nodes, connect_dict):
        # 仅保留属于“联结系”的节点
        connect_only = [(nid, coord) for (nid, coord) in group_nodes if nid in connect_dict]
        # 按 y 归并
        piles = {}
        for nid, (x, y, z) in connect_only:
            key = y
            piles.setdefault(key, []).append((nid, (x, y, z)))
        # 每个 y 应有两个联结点：上、下（按 z 排序）
        pile_map_y = {}
        for y, lst in piles.items():
            if len(lst) < 2:
                # 该根桩联结点不足两处，忽略
                continue
            # 按 z 从高到低：第一个当作上联结，第二个当作下联结
            lst.sort(key=lambda t: t[1][2], reverse=True)
            up = lst[0]
            dn = lst[1]
            pile_map_y[y] = {"x": up[1][0], "up": up, "dn": dn}
        return pile_map_y

    for _, group_nodes in sorted(pile_group_nodes.items()):
        pile_map_y = build_pile_map_by_y(group_nodes, pile_connect_nodes)
        if len(pile_map_y) < 2:
            continue

        y_sorted = sorted(pile_map_y.keys())
        for i in range(len(y_sorted) - 1):
            y1, y2 = y_sorted[i], y_sorted[i + 1]
            p1, p2 = pile_map_y[y1], pile_map_y[y2]
            # 同一列 x
            x = p1["x"]
            # 中点坐标
            y_mid = (y1 + y2) / 2.0
            z_mid = (p1["up"][1][2] + p1["dn"][1][2] + p2["up"][1][2] + p2["dn"][1][2]) / 4.0
            mid_coord = (x, y_mid, z_mid)

            connect_x_nodes[gid_out] = (node_id, mid_coord)
            connect_groups[gid_out] = [p1["up"], p1["dn"], (node_id, mid_coord), p2["up"], p2["dn"]]

            node_id += 1
            gid_out += 1

    return connect_x_nodes, connect_groups, node_id

# ==========获取钻孔桩位==========
def get_drill_coords(tab_components, tab_drillarea):
    """
    根据 UI 输入 + 小肋长度，返回整体坐标系下的桩位坐标
    """

    # 1. 承台中心相对小肋中点的偏移
    dx_cap = float(tab_drillarea.x_entry.get() or 0.0)
    dy_cap = float(tab_drillarea.y_entry.get() or 0.0)

    # 2. 小肋中点在整体坐标系的位置 
    xl_length = float(tab_components.entry_smallrib_length.get() or 13100.0)  # 小肋长度
    rib_center_y = xl_length / 2.0

    pile_coords = []
    for item in tab_drillarea.pile_list:
        px = float(item["x"].get() or 0.0)
        py = float(item["y"].get() or 0.0)

        # 3. 绝对坐标转换
        X_abs = px + dx_cap
        Y_abs = py + dy_cap + rib_center_y
        Z_abs = 0.0

        pile_coords.append((X_abs, Y_abs, Z_abs))

    return pile_coords

# ==========钢护筒节点==========
# 钢护筒不参与结构计算，仅作位置示意
def build_casing_nodes(tab_components, tab_drillarea, start_id=1):
    """
    钢护筒节点生成（含桩底节点提取）
    返回:
        casing_nodes: {nid: (x,y,z)}
        casing_bottom_nodes: {nid: (x,y,z)}   ★ 已按你的要求调整
        next_node_id
    """

    L = float(tab_components.entry_pile_length.get() or 16.0) * 1000
    pile_coords = get_drill_coords(tab_components, tab_drillarea)

    casing_nodes = {}
    tmp_nodes = []
    node_id = start_id

    # ===== 1. 生成每根护筒的上下节点 =====
    for idx, item in enumerate(tab_drillarea.pile_list):
        X_abs, Y_abs, Z_base = pile_coords[idx]

        Z_top = 500.0
        Z_bottom = 500.0 - L - 5000

        nid_top = node_id
        nid_bottom = node_id + 1
        node_id += 2

        tmp_nodes.append((nid_top,    (X_abs, Y_abs, Z_top)))
        tmp_nodes.append((nid_bottom, (X_abs, Y_abs, Z_bottom)))

    # 按 (x,y) 分组，z 降序排序
    tmp_nodes.sort(
        key=lambda rec: (
            round(rec[1][0], 6),
            round(rec[1][1], 6),
            -round(rec[1][2], 6)
        )
    )

    for nid, coord in tmp_nodes:
        casing_nodes[nid] = coord

    # ===== 2. 提取桩底节点（你的要求格式：nid: (x,y,z)） =====
    from collections import defaultdict
    piles_by_xy = defaultdict(list)

    for nid, (x, y, z) in casing_nodes.items():
        xy_key = (round(x, 6), round(y, 6))
        piles_by_xy[xy_key].append((nid, z, x, y))

    casing_bottom_nodes = {}

    for xy_key, data in piles_by_xy.items():
        # z 最小即桩底
        data.sort(key=lambda t: t[1])  # t = (nid, z, x, y)
        nid_bottom, z_b, x_b, y_b = data[0]

        casing_bottom_nodes[nid_bottom] = (x_b, y_b, z_b)

    return casing_nodes, casing_bottom_nodes, node_id


# ------------------------------------------------
# 单元生成
# ------------------------------------------------

# ==========生成梁单元通用函数==========
def build_beam_elements(
    group_nodes: dict,
    section_name: str,
    section_data: dict,
    name_to_id: dict,
    start_elem_id=1
):
    """
    通用梁单元生成函数（支持按根生成）
    """
    section_id = name_to_id.get(section_name, 1)
    params = section_data.get(section_name, {}).get("params", {})

    # β角度提取
    beta_angle = 0
    for k, v in params.items():
        if "β" in k or "角度" in k or "beta" in k.lower():
            try:
                beta_angle = float(v)
            except:
                beta_angle = 0
            break

    # 材料号判断
    if any(key in section_name for key in ["贝雷", "Bailey", "bailey"]):
        mat_id = 1   # 贝雷梁类 → 16Mn
    else:
        mat_id = 2   # 其他构件 → Q235

    elem_id = start_elem_id
    beam_elements = []

    # 遍历每根构件
    for gid, node_list in sorted(group_nodes.items()):
        if not node_list or len(node_list) < 2:
            continue
        node_ids = [nid for nid, _ in node_list]
        for i in range(len(node_ids) - 1):
            n1, n2 = node_ids[i], node_ids[i + 1]
            beam_elements.append(
                f"{elem_id}, BEAM, {mat_id}, {section_id}, {n1}, {n2}, {beta_angle}, 0"
            )
            elem_id += 1

    next_elem_id = elem_id
    return beam_elements, next_elem_id


# ==========生成小肋单元信息==========
def build_small_rib_elements(
    components_tab,
    section_data: dict,
    name_to_id: dict,
    xl_single_nodes: dict,
    start_elem_id=1
):
    """小肋单元生成封装函数"""
    rib_elems, next_elem_id = build_beam_elements(
        xl_single_nodes, "小肋", section_data, name_to_id, start_elem_id
    )
    return rib_elems, next_elem_id


# ==========生成贝雷梁上弦杆单元信息==========
def build_bailey_top_elements(
    section_data: dict,
    name_to_id: dict,
    bailey_top_single_nodes: dict,
    start_elem_id=1
):
    """贝雷梁上弦 BEAM 单元生成函数"""
    top_elems, next_elem_id = build_beam_elements(
        bailey_top_single_nodes, "贝雷梁弦杆", section_data, name_to_id, start_elem_id
    )
    return top_elems, next_elem_id

# ==========生成贝雷梁下弦杆单元信息==========
def build_bailey_bottom_elements(
    section_data: dict,
    name_to_id: dict,
    bailey_bottom_single_nodes: dict,
    start_elem_id=1
):
    """贝雷梁下弦 BEAM 单元生成函数"""
    bottom_elems, next_elem_id = build_beam_elements(
        bailey_bottom_single_nodes, "贝雷梁弦杆", section_data, name_to_id, start_elem_id
    )
    return bottom_elems, next_elem_id



# ==========生成贝雷梁竖杆单元信息==========
def build_bailey_vertical_elements(
    bailey_top_vertical_nodes: dict,
    bailey_bottom_vertical_nodes: dict,
    section_data: dict,
    name_to_id: dict,
    start_elem_id=1
):
    """生成贝雷梁竖杆 BEAM 单元"""
    section_name = "贝雷梁竖杆"
    section_id = name_to_id.get(section_name, 1)
    beta_angle = 90
    mat_id = 1
    elem_id = start_elem_id
    vertical_elements = []

    bottom_coord_map = {
        (round(x, 3), round(y, 3)): (nid, z)
        for nid, (x, y, z) in bailey_bottom_vertical_nodes.items()
    }

    for top_nid, (x, y, z_top) in bailey_top_vertical_nodes.items():
        key = (round(x, 3), round(y, 3))
        if key in bottom_coord_map:
            bottom_nid, z_bot = bottom_coord_map[key]
            vertical_elements.append(
                f"{elem_id}, BEAM, {mat_id}, {section_id}, {top_nid}, {bottom_nid}, {beta_angle}, 0"
            )
            elem_id += 1

    next_elem_id = elem_id
    return vertical_elements, next_elem_id


# ==========生成贝雷梁斜杆单元信息==========
def build_bailey_diagonal_elements(
    groups: dict,
    section_data: dict,
    name_to_id: dict,
    start_elem_id=1
):
    """生成贝雷梁斜杆 BEAM 单元"""
    section_name = "贝雷梁斜杆"
    section_id = name_to_id.get(section_name, 1)
    beta_angle = 90
    mat_id = 1
    elem_id = start_elem_id
    diagonal_elements = []

    for gid, info in sorted(groups.items()):
        node_ids = info.get("nodes", [])
        if len(node_ids) != 4:
            continue

        top_id, left_id, bottom_id, right_id = node_ids
        elem_pairs = [
            (top_id, left_id),
            (left_id, bottom_id),
            (bottom_id, right_id),
            (right_id, top_id),
        ]

        for n1, n2 in elem_pairs:
            diagonal_elements.append(
                f"{elem_id}, BEAM, {mat_id}, {section_id}, {n1}, {n2}, {beta_angle}, 0"
            )
            elem_id += 1

    next_elem_id = elem_id
    return diagonal_elements, next_elem_id


# ==========生成分配梁单元信息==========
def build_distribute_elements(
    components_tab,
    section_data: dict,
    name_to_id: dict,
    distribute_single_nodes: dict,
    start_elem_id=1
):
    """分配梁单元生成函数"""
    elem_id = start_elem_id
    distribute_elements = []

    # 普通分配梁
    distribute_elems, elem_id = build_beam_elements(
        distribute_single_nodes, "分配梁", section_data, name_to_id, elem_id
    )
    distribute_elements.extend(distribute_elems)

    next_elem_id = elem_id
    return distribute_elements, next_elem_id

# ==========生成钢管桩单元信息==========
def build_pile_elements(
    section_data: dict,
    name_to_id: dict,
    pile_group_nodes: dict,
    start_elem_id=1
):
    """
    生成钢管桩 BEAM 单元。
    逻辑：
        - 钢管桩按 x 分组（同 x 值为一列）；
        - 每组内节点按 (x,y) 相同，z 坐标排序连接；
        - 材料号 2（Q235），β角 = 0。
    参数：
        section_data            : 截面参数全集
        name_to_id              : 截面名 → ID 映射
        pile_group_nodes        : 钢管桩 {x组号: [(nid, (x,y,z)), ...]}
        start_elem_id           : 起始单元号
    返回：
        pile_elements : list[str] —— MCT 行
        next_elem_id  : int —— 下一可用单元号
    """
    section_name = "钢管桩"
    section_id = name_to_id.get(section_name, 1)
    beta_angle = 0
    mat_id = 2  # Q235
    elem_id = start_elem_id
    pile_elements = []

    for gid, nodes in sorted(pile_group_nodes.items()):
        # 按 (y,z) 排序，以保证竖向连接
        sorted_nodes = sorted(nodes, key=lambda n: (round(n[1][1], 6), n[1][2]))
        # 分组内再按 y 聚合，每根桩竖向连通
        from collections import defaultdict
        piles_by_y = defaultdict(list)
        for nid, (x, y, z) in sorted_nodes:
            piles_by_y[round(y, 6)].append((nid, z))

        for y_key, nid_z_list in piles_by_y.items():
            nid_z_list.sort(key=lambda nz: nz[1])  # 按 z 坐标升序（由上到下）
            for i in range(len(nid_z_list) - 1):
                n1, _ = nid_z_list[i]
                n2, _ = nid_z_list[i + 1]
                pile_elements.append(
                    f"{elem_id}, BEAM, {mat_id}, {section_id}, {n1}, {n2}, {beta_angle}, 0"
                )
                elem_id += 1

    next_elem_id = elem_id
    return pile_elements, next_elem_id

# ==========生成联结系单元信息==========
def build_connect_x_elements(
    section_data: dict,
    name_to_id: dict,
    connect_groups: dict,
    start_elem_id=1
):
    """
    生成联结系 BEAM 单元。
    逻辑：
        - 每组结构固定为 [左上, 左下, 中心, 右上, 右下]
        - 连接顺序：
            1. 左上 → 右上
            2. 左下 → 右下
            3. 左上 → 中心
            4. 中心 → 右下
            5. 左下 → 中心
            6. 中心 → 右上
        - 材料号固定为 2（Q235），β角 0。
    参数：
        connect_groups        : 普通联结系 {group_id: [(nid, coord), ...]}
        section_data          : 截面参数全集
        name_to_id            : 截面名 → ID映射
        start_elem_id         : 起始单元号
    返回：
        connect_elements : list[str] —— 联结系单元 MCT 行
        next_elem_id     : int —— 下一可用单元号
    """

    section_name = "联结系"
    section_id = name_to_id.get(section_name, 1)
    mat_id = 2
    beta_angle = 0
    elem_id = start_elem_id
    connect_elements = []

    for gid, nodes in sorted(connect_groups.items()):
        if not nodes or len(nodes) < 5:
            continue
        up_left, down_left, mid, up_right, down_right = nodes

        connect_pairs = [
            (up_left[0], up_right[0]),    # 左上 → 右上
            (down_left[0], down_right[0]),# 左下 → 右下
            (up_left[0], mid[0]),         # 左上 → 中心
            (mid[0], down_right[0]),      # 中心 → 右下
            (down_left[0], mid[0]),       # 左下 → 中心
            (mid[0], up_right[0]),        # 中心 → 右上
        ]

        for n1, n2 in connect_pairs:
            connect_elements.append(
                f"{elem_id}, BEAM, {mat_id}, {section_id}, {n1}, {n2}, {beta_angle}, 0"
            )
            elem_id += 1

    next_elem_id = elem_id
    return connect_elements, next_elem_id

def build_casing_elements(
    section_data: dict,
    name_to_id: dict,
    casing_nodes: dict,
    start_elem_id=1):
    section_name = "钢护筒"
    section_id = name_to_id.get(section_name, 1)
    mat_id = 2
    beta_angle = 0
    elem_id = start_elem_id
    casing_elements = []
    from collections import defaultdict

    piles_by_xy = defaultdict(list)

    for nid, (x, y, z) in casing_nodes.items():
        xy_key = (round(x, 6), round(y, 6))
        piles_by_xy[xy_key].append((nid, z))


    for xy_key, nid_z_list in sorted(piles_by_xy.items()):
        nid_z_list.sort(key=lambda nz: -nz[1])

        # 按顺序两两连接
        for i in range(len(nid_z_list) - 1):
            n1, z1 = nid_z_list[i]
            n2, z2 = nid_z_list[i + 1]

            casing_elements.append(
                f"{elem_id}, BEAM, {mat_id}, {section_id}, {n1}, {n2}, {beta_angle}, 0"
            )
            elem_id += 1

    return casing_elements, elem_id 

# ==========生成板单元信息==========
def build_plate_elements(xl_nodes: dict, start_eid=1):
    """
    根据 xl_nodes（所有小肋节点）生成板单元。
    xl_nodes: {nid: (x, y, z)}

    小肋按 y 分排，每排 x 方向节点形成横向条带。
    两排之间按相邻 x 构成矩形板单元。
    """

    # ——1. 按 y 将节点分类为 rows：
    rows_dict = {}  # {y_value: [(nid, (x,y,z)), ...]}

    for nid, (x, y, z) in xl_nodes.items():
        rows_dict.setdefault(y, []).append((nid, (x, y, z)))

    # ——2. 对每一排按 x 排序：
    for y in rows_dict:
        rows_dict[y].sort(key=lambda item: item[1][0])  # 按 x 排序

    # ——3. 按 y 排序
    sorted_y = sorted(rows_dict.keys())
    rows = [rows_dict[y] for y in sorted_y]

    plate_elems = []
    eid = start_eid

    # ——4. 邻排形成板单元
    for k in range(len(rows) - 1):
        row_lower = rows[k]      # y = y_k
        row_upper = rows[k + 1]  # y = y_{k+1}

        # 只能形成 N-1 个矩形
        for j in range(len(row_lower) - 1):

            n_ld = row_lower[j][0]       # 左下
            n_rd = row_lower[j + 1][0]   # 右下
            n_ru = row_upper[j + 1][0]   # 右上
            n_lu = row_upper[j][0]       # 左上

            line = f"{eid}, PLATE, 1, 1, {n_ld}, {n_rd}, {n_ru}, {n_lu}, 2, 0"
            plate_elems.append(line)

            eid += 1

    return plate_elems

# ==========================================================
#   建立结构组
# ==========================================================
def build_structure_groups(
    xl_nodes: dict,
    all_nodes: dict,
    bailey_top_nodes: dict,
    bailey_bottom_nodes: dict,
    vertical_nodes: dict,
    diagonal_nodes: dict,
    distribute_nodes: dict,
    pile_group_nodes: dict,
    connect_groups: dict,
    rib_elems,
    top_elems,
    bottom_elems,
    vertical_elements,
    diagonal_elements,
    distribute_elements,
    pile_elements,
    connect_elements,
):
    """
    生成结构组 (*GROUP) 信息。
    ELEM_LIST 输出为“单元号压缩串”（如：1 to 10, 15, 18 to 20），
    而不是整条 MCT 的 ELEM 行。
    返回:
        group_lines : list[str]
        stats : dict{name: {"node_count": n, "elem_count": m}}
    """
    import re

    # 各种形式的“单元列表”统一提取为 int 的单元号列表
    def _extract_elem_ids(maybe_mct_list):
        """
        支持：
          - ["1, BEAM, ...", "ELEM, 2, BEAM, ...", ...]
          - [1, 2, 3]
          - 混合
        """
        ids = []
        if not maybe_mct_list:
            return ids

        for item in maybe_mct_list:
            # 如果是 int，直接加入
            if isinstance(item, int):
                ids.append(item)
                continue

            # 如果是字符串，取第一段数字
            if isinstance(item, str):
                parts = [p.strip() for p in item.split(",") if p.strip()]
                if not parts:
                    continue
                try:
                    # 第一列为单元号
                    first = parts[0]
                    if first.isdigit():
                        ids.append(int(first))
                except Exception:
                    continue

        return ids
    
    # 生成 *GROUP 文本，节点与单元号段用空格分隔，同段之间逗号分隔
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
    # -----------------------------
    # 1) 汇总所有节点（保持你现在的汇总写法）
    #    *注意*：这里默认你传入的是“已扁平”的 nid->coord 字典；
    #    如果传的是分组结构 {gid: [(nid,coord),...]}，需要在外面先拍平。
    # -----------------------------
    all_node_dicts = {
        "小肋": xl_nodes,
        "贝雷梁弦杆": {**(bailey_top_nodes or {}), **(bailey_bottom_nodes or {})},
        "贝雷梁竖杆": vertical_nodes or {},
        "贝雷梁斜杆": diagonal_nodes or {},
        "分配梁": distribute_nodes or {},
        "钢管桩": pile_nodes_flat or {},
        "联结系": connect_nodes_flat or {}
    }

    # -----------------------------
    # 2) 汇总所有单元并“抽取ID列表”
    # -----------------------------
    all_elem_dicts = {
        "小肋": _extract_elem_ids(rib_elems),
        "贝雷梁弦杆": (
            _extract_elem_ids(top_elems)
            + _extract_elem_ids(bottom_elems)),
        "贝雷梁竖杆": _extract_elem_ids(vertical_elements),
        "贝雷梁斜杆": _extract_elem_ids(diagonal_elements),
        "分配梁": _extract_elem_ids(distribute_elements),
        "钢管桩": _extract_elem_ids(pile_elements),
        "联结系": _extract_elem_ids(connect_elements),
    }

    # -----------------------------
    # 3) 生成结构组文本（ELEM_LIST 也用 num_lst_to_num_group 压缩）
    # -----------------------------
    group_lines = []
    plane_type = 0

    plane_type = 0

    for name, node_dict in all_node_dicts.items():
        if not node_dict:
            continue

        # 节点部分
        node_ids = sorted(node_dict.keys())
        node_range_str = num_lst_to_num_group(node_ids) if node_ids else ""

        # 单元部分
        elem_ids = sorted(set(all_elem_dicts.get(name, []) or []))
        elem_range_str = num_lst_to_num_group(elem_ids) if elem_ids else ""

        # 用空格连接节点号序列，节点和单元之间用逗号分隔
        group_line = f"{name}, {node_range_str.replace(',', '')}, {elem_range_str.replace(',', '')}, {plane_type}"
        group_lines.append(group_line)

    return group_lines

# ==========================================================
#   生成弹性连接节点
# ==========================================================
def build_elastic_links(
    components_tab,
    # 1) 小肋-上弦
    xl_bailey_elink_nodes: dict,            # {rib_idx: [nid, ...]}
    bailey_top_elink_nodes: dict,      # {nid: (x,y,z)}
    # 2) 下弦-分配梁
    bailey_bottom_elink_nodes: dict,    # {nid: (x,y,z)}
    distribute_bottom_elink_nodes: dict,# {nid: (x,y,z)}
    # 3) 分配梁-钢管桩
    distribute_pile_elink_nodes: dict,    # {nid: (x,y,z)}
    pile_top_nodes: dict,                 # {nid: (x,y,z)}
    start_link_id: int = 1
    ):
        """
        生成 *ELASTICLINK 段文本。
        规则：两两取 (x,y) 相同的点，z 较大的连接到 z 较小的。
        返回：
            elink_lines: [str, ...]
            next_link_id: int
        """
        def _flatten_dict_nodes(group_dict: dict) -> dict:
            """将 {组号:[nid,...]} 展平为 {nid:(x,y,z)}，若原输入已为 dict 则直接返回"""
            flat = {}
            if not group_dict:
                return flat
            # 如果是 {nid:(x,y,z)}，直接返回
            first_val = next(iter(group_dict.values()))
            if isinstance(first_val, tuple) and len(first_val) == 3:
                return group_dict
            # 否则说明是 {组:[nid,...]}，不含坐标
            # 需要用全局坐标表 bailey_top_elink_nodes 获取坐标（针对小肋）
            for arr in group_dict.values():
                for nid in arr:
                    if nid in bailey_top_elink_nodes:
                        flat[nid] = bailey_top_elink_nodes[nid]
            return flat

        def _pair_nodes(dict_a: dict, dict_b: dict, link_id: int, label: str):
            """匹配两组节点，按 (x,y) 相同且 z 大→z 小连接"""
            lines = []
            matched = 0

            # 构造快速查找表
            b_map = {(round(x, 6), round(y, 6)): (nid, z) for nid, (x, y, z) in dict_b.items()}

            for nid_a, (xa, ya, za) in dict_a.items():
                key = (round(xa, 6), round(ya, 6))
                if key not in b_map:
                    continue
                nid_b, zb = b_map[key]
                n1, n2 = (nid_a, nid_b) if za >= zb else (nid_b, nid_a)
                lines.append(
                    f" {link_id}, {n1}, {n2}, GEN, 0, NO, NO, NO, NO, NO, NO, "
                    "1e+07, 1e+07, 1e+07, 0.01, 0.01, 0.01, NO, 0.5, 0.5, 弹性连接"
                )
                link_id += 1
                matched += 1

            return lines, link_id

        # ------------------- 主逻辑 -------------------
        elink_lines = []
        lid = start_link_id

        # 1. 小肋 ↔ 贝雷梁上弦
        xl_nodes_flat = _flatten_dict_nodes(xl_bailey_elink_nodes)
        lines, lid = _pair_nodes(xl_nodes_flat, bailey_top_elink_nodes, lid, "小肋-上弦弹连")
        elink_lines.extend(lines)

        # 2. 下弦 ↔ 分配梁
        lines, lid = _pair_nodes(bailey_bottom_elink_nodes, distribute_bottom_elink_nodes, lid, "下弦-分配梁弹连")
        elink_lines.extend(lines)

        # 3. 分配梁 ↔ 钢管桩
        lines, lid = _pair_nodes(distribute_pile_elink_nodes, pile_top_nodes, lid, "分配梁-普通桩弹连")
        elink_lines.extend(lines)

        return elink_lines

# ==========================================================
#   释放梁端约束
# ==========================================================
def build_FRAME_lst(ELEMENT_lst: list[str], release_nodes: list[int]):
    """
    根据 release_nodes（右端节点号）从所有 BEAM 单元中匹配对应单元，
    并生成 *FRAME-RLS 段文本。

    参数：
        ELEMENT_lst : list[str]
            所有 BEAM 单元行，如：
            "42658, BEAM, 1, 1, 200, 201, 0.0, 0"
        release_nodes : list[int]
            已整理好的右端节点号（J 端）

    返回：
        FRAME_lst : list[str]
            可直接写入 MCT 的 *FRAME-RLS 段文本
    """

    FRAME_lst = []
    if not ELEMENT_lst or not release_nodes:
        print("[提示] 缺少 ELEMENT 或 release_nodes，跳过 *FRAME-RLS 段。")
        return FRAME_lst

    release_set = set(release_nodes)
    release_elems = []

    # 遍历单元行，识别右端节点属于 release_nodes 的单元
    for line in ELEMENT_lst:
        parts = [x.strip() for x in line.split(",")]
        if len(parts) < 6 or "BEAM" not in line:
            continue
        try:
            elem_id = int(parts[0])
            node2 = int(parts[5])  # J端
        except ValueError:
            continue
        if node2 in release_set:
            release_elems.append(elem_id)

    if not release_elems:
        print("[错误] 未找到与 release_nodes 对应的 BEAM 单元（右端节点未匹配）。")
        return FRAME_lst

    # 输出 MCT 段文本
    for eid in sorted(release_elems):
        FRAME_lst.append(f"{eid}, NO, 000000, 0, 0, 0, 0, 0, 0")
        FRAME_lst.append(f"          000010, 0, 0, 0, 0, 0, 0, 释放梁端约束")

    print(f"[提示] 成功生成 {len(release_elems)} 个梁端释放约束单元。")
    return FRAME_lst


# ==========================================================
#   桩底固结
# ==========================================================
def build_constraint_lines(pile_bottom_nodes: dict,casing_bottom_nodes: dict):
    """
    生成桩底固结 (*CONSTRAINT) 段。
    参数:
        pile_bottom_nodes: dict → 桩底节点 {nid:(x,y,z), ...}
    返回:
        constraint_lines: list[str] → 可直接写入 MCT 的行
    """
    constraint_lines = []

    # 合并两个节点集合
    all_nodes = []
    if pile_bottom_nodes:
        all_nodes += sorted(pile_bottom_nodes.keys())
    if casing_bottom_nodes:
        all_nodes += sorted(casing_bottom_nodes.keys())
    if not all_nodes:
        print("[提示] 未检测到桩底节点，跳过 *CONSTRAINT 段。")
        return constraint_lines

    node_groups_str = num_lst_to_num_group(all_nodes)
    node_groups_str = node_groups_str.replace(",", " ")
    constraint_lines.append(f"{node_groups_str}, 111111, 桩底反力")

    return constraint_lines

# ==========分配面荷载定义(根据钻机重量及履带参数)==========
def build_pnloadtype(load_tab):
    #需要输入load_tab = app.tab_loads
    """
    返回 *PNLOADTYPE 文本
    *PNLOADTYPE
    NAME=钻机名称, AREA,
    DATA=NO,NO,0, 0,荷载,履带长, 0,荷载,履带长, 履带宽, 荷载,0, 履带宽, 荷载
    """
    drill_name = load_tab.drill_name_entry.get()  
    weight = float(load_tab.drill_weight_entry.get()) or 71.0 #单位t
    L = float(load_tab.ld_length_entry.get()) or 4675.0
    B = float(load_tab.ld_width_entry.get()) or 800.0
    Pa =  - weight * 1000 * 9.8 / 2 / (L*B)

    line1 = "*PNLOADTYPE"
    line2 = f"NAME={drill_name}, AREA,"
    
    # 主体 DATA 行
    line3 = (
        "DATA=NO,NO,"
        f"0, 0,{Pa},"
        f"{L}, 0,{Pa},"
        f"{L}, {B}, {Pa},"
        f"0, {B}, {Pa}"
    )

    return [line1, line2, line3]


# ==========================================================
#   荷载组及荷载说明
# ==========================================================
def build_stldcase(pile_coords):
    stld_lines = []
    for i in range(len(pile_coords)):
        idx = i + 1
        stld_lines.append(f"桩{idx}-正钻, USER,")
        stld_lines.append(f"桩{idx}-侧钻, USER,")
    return stld_lines

def build_drilling_loadcase_names(pile_coords):
    """
    返回纯工况名列表：["桩1-正钻", "桩1-侧钻", "桩2-正钻", "桩2-侧钻", ...]
    """
    names = []
    for i in range(len(pile_coords)):
        idx = i + 1
        names.append(f"桩{idx}-正钻")
        names.append(f"桩{idx}-侧钻")
    return names

# ==========================================================
#  钻机荷载(分配面荷载)
# ==========================================================
def build_planeload(load_tab, pile_coords):
    """
    生成分配面荷载 (*PNLOAD / PLATE) 命令行，每个桩两种工况：正钻、侧钻
    返回列表，每行即最终写入MCT的字符串
    """

    # 获取参数
    ld_length = float(load_tab.ld_length_entry.get()) or 4675.0
    ld_width = float(load_tab.ld_width_entry.get()) or 800.0
    ld_axis = float(load_tab.ld_axis_entry.get().strip()) or 3600.0
    drill_name = load_tab.drill_name_entry.get()
    drill_radius = float(load_tab.drill_radius_entry.get()) or 100.0

    lines = []
    idx_case = 0  # stldcase_names 顺序：桩1正钻, 桩1侧钻, 桩2正钻, 桩2侧钻...

    for x0, y0, z0 in pile_coords:
        # ----------- 正钻 -----------
        # 偏移量
        dx = -(drill_radius + ld_length / 2)
        dy_up = + ld_axis / 2 - ld_width / 2
        dy_dn = - ld_axis / 2 - ld_width / 2

        # 两个加载点
        pts = [
            (x0 + dx, y0 + dy_up),   # 左上
            (x0 + dx, y0 + dy_dn),   # 左下
        ]

        # 对应工况名
        case_names = build_drilling_loadcase_names(pile_coords)
        case_name = case_names[idx_case]
        idx_case += 1

        for (xx, yy) in pts:
            lines.append(
                f"{case_name}, {drill_name}, PLATE,\n"
                f"   LPLANE, , 1, NLP, NO,\n"
                f"   {xx},{yy},0, {xx+1},{yy},0, {xx},{yy+1},0, 1, NO"
            )

        # ----------- 侧钻 -----------
        dx2 = -(ld_length / 2)
        dy_up2 = - drill_radius + ld_axis/2 - ld_width/2
        dy_dn2 = - drill_radius - ld_axis/2 - ld_width/2

        pts2 = [
            (x0 + dx2, y0 + dy_up2),    # 上
            (x0 + dx2, y0 + dy_dn2),    # 更下
        ]

        case_names = build_drilling_loadcase_names(pile_coords)
        case_name = case_names[idx_case]
        idx_case += 1

        for (xx, yy) in pts2:
            lines.append(
                f"{case_name}, {drill_name}, PLATE,\n"
                f"   LPLANE, , 1, NLP, NO,\n"
                f"   {xx},{yy},0, {xx+1},{yy},0, {xx},{yy+1},0, 1, NO"
            )

    return lines



# ==========================================================
#  风荷载
# ==========================================================
def build_wind_load(load_tab, components_tab ,pile_top_nodes: dict):
    """
    生成风荷载（节点荷载）MCT命令行。
    
    参数：
        load_tab: LoadTab 实例（包含风速输入框）
        components_tab: ComponentsTab 实例（用于读取贝雷梁片数）
        pile_top_nodes: dict {nid: (x, y, z)} 桩顶节点

    返回：
        wind_lines: list[str]  # 风荷载节点行
    """
    Feng_values = []
    Feng_result_lst = []

    # 提取参数
    try:
        u10_gz = float(load_tab.gongzuo_wind_entry.get())
        dbfl   = load_tab.wind_combo1.get()
        kt     = float(load_tab.kt_entry.get())
        Z      = float(load_tab.H_entry.get())
        CH     = float(load_tab.kh_entry.get())
        formula = load_tab.wind_combo2.get()
    except Exception as e:
        print("[警告] 风荷载参数读取失败：", e)
        return []

    # 从 ComponentsTab 读取贝雷梁片数
    try:
        bailey_count = components_tab.bailey_count.get()
    except Exception:
        bailey_count = 6  # 默认6片

    # 计算水平加载长度
    bailey_span_expr = components_tab.entry_bailey_span.get().strip()
    bailey_span_steps_m = midasdisttolst(bailey_span_expr)
    bailey_span_steps = [d * 1000.0 for d in bailey_span_steps_m] # 单位mm
    L = sum(bailey_span_steps) # 单位mm
    Feng_values = [u10_gz, dbfl, kt, Z, L, bailey_count, CH, formula]

    # 计算风压
    Gv1, Gv2, kf, kh, kc, Us10, Ud, a0, Ug1, Ug2, Pa1, Pa2, Pa3, A, FF = cal_FengYa_JTG_T_3360_01_2018(u10_gz, kt, Z, L, bailey_count, CH, dbfl, formula)
    F_gz = Pa1 * Z * 1000
    Feng_result_lst = [ Gv1, Gv2, kf, kh, kc, Us10, Ud, a0, Ug1, Ug2, Pa1, Pa2, Pa3, A, FF]
    # === 4. 整理桩顶节点 ===
    node_ids = list(pile_top_nodes.keys())

    # === 5. 平均分配到每个桩顶节点 ===
    Fy_per_node_gz = -F_gz / len(node_ids)
    # Fy_per_node_fgz = -F_fgz / len(node_ids)

    
    # === 6. 生成工作状态与非工作状态两组荷载 ===
    gz_state_lines = [f"{nid}, 0, {Fy_per_node_gz:.2f}, 0, 0, 0, 0, , " for nid in sorted(node_ids)]
    # fgz_state_lines = [f"{nid}, 0, {Fy_per_node_fgz:.2f}, 0, 0, 0, 0, , " for nid in sorted(node_ids)]

    # === 7. 拼接为完整 MCT 段 ===
    WIND_GROUP_LINES = [
        "*USE-STLD, 风荷载",
        "*CONLOAD",
        *gz_state_lines,
    ]

    return WIND_GROUP_LINES, Feng_values, Feng_result_lst

# ==========================================================
#  水流力
# ==========================================================
def build_water_force_load(load_tab, section_data, components_tab, 
                           water_force_points_by_y: dict, ):
    """
    生成水流力（节点荷载）MCT命令行。

    参数：
        load_tab: LoadTab 实例（包含水流参数输入框）
        section_data: dict，截面数据（需包含钢管桩直径信息）
        water_force_points_by_y: dict[y] = [(node_id, (x,y,z)), ...]，水流力加载点

    返回：
        WATER_FORCE_LINE: list[str]  # MCT文本行
    """
    Shui_values = []
    Shui_lst =[]

    # === 1. 读取 UI 输入参数 ===
    try:
        v = float(load_tab.flow_velocity_entry.get() or 0.0)     # 设计流速 (m/s)
        h = float(load_tab.water_level_entry.get() or 0.0)       # 入水深度 (m)
        s = float(load_tab.water_depth_entry.get() or 0.0)       # 入水深度 (m)
        Shui_values = [v, h, s]

    except Exception as e:
        print("[警告] 水流力参数读取失败：", e)
        return []

    # === 2. 读取钢管桩直径 d (m) ===
    try:
        d = float(section_data.get("钢管桩", {}).get("params", {}).get("D(mm)", 630.0))
        pile_y_expr = components_tab.entry_pile_y.get().strip()
        pile_y_list = midasdisttolst(pile_y_expr)
        if not pile_y_list:
            pile_y_list = [float(pile_y_expr or 3000.0)]
    except Exception:
        d = 630  # 
        pile_y_list = [3000.0]

    # === 3. 计算水流阻力系数 ===    
    def get_m1_from_lambda(lam):
        """
        根据 lambda = (L/D) 通过表格插值得到折减系数 m1
        """

        # 表格数据点（按 L/D 排序）
        LD_points = [1, 2, 3, 4, 6, 8, 12, 16, 18, 20]
        m1_points = [-0.38, 0.25, 0.54, 0.66, 0.78, 0.82, 0.86, 0.88, 0.90, 1.00]

        # --- 边界情况处理 ---
        if lam <= LD_points[0]:
            return m1_points[0]
        if lam >= LD_points[-1]:
            return m1_points[-1]

        # --- 区间查找并插值 ---
        for i in range(len(LD_points) - 1):
            x1, x2 = LD_points[i], LD_points[i+1]
            if x1 <= lam <= x2:
                y1, y2 = m1_points[i], m1_points[i+1]
                return inear_interpolation(x1, x2, y1, y2, lam)
        return m1_points[-1]

    def Cw_cal(pile_y_list,d,Cw_first):
        Cw_dict = {}
        # 第一排固定
        prev_Cw = Cw_first
        Cw_dict["Cw_1"] = prev_Cw

        for i, L in enumerate(pile_y_list):
            lam = (L - d) / d     
            m1 = get_m1_from_lambda(lam)
            new_Cw = round(prev_Cw * m1, 2) 
            Cw_dict[f"Cw_{i+2}"] = new_Cw  
            prev_Cw = new_Cw
        return Cw_dict

    # === 4. 定义参数 ===
    rho = 1000.0  # 水密度 (kg/m³)
    Cw_dict = Cw_cal(pile_y_list,d,0.73)
    Cw_list = [ Cw_dict[f"Cw_{i+1}"] for i in range(len(Cw_dict)) ]

    # === 5. 按 y 坐标分排 ===
    # y 较大 → 前排
    sorted_y = sorted(water_force_points_by_y.keys(), reverse=True)
    if not sorted_y:
        print("[提示] 未找到普通桩水流力加载点。")
        return []
    
    # === 6. 计算水流力 ===
    def calc_force(Cw):
        Fw = round(0.5 * rho * Cw * v*v * d * h / 1000 , 2)
        return Fw  # N

    Cw_front = Cw_list[0] 
    Fw_front = calc_force(Cw_front)
    
    # === 7. 生成普通钢管桩的水流力行 ===
    WATER_FORCE_LINE = [
        "*USE-STLD, 水流力",
        "*CONLOAD",
    ]

    # 遍历每一排 y，并配对应 Cw
    for idx, y in enumerate(sorted_y):

        # 若排数多于 Cw 数，则最后一排统一采用最后一个 Cw
        if idx < len(Cw_list):
            Cw = Cw_list[idx]
        else:
            Cw = Cw_list[-1]

        Fw = calc_force(Cw)

        # 输出这一排的所有节点
        for nid, coord in water_force_points_by_y[y]:
            line = f"{nid}, 0, -{Fw:.2f}, 0, 0, 0, 0, , "
            WATER_FORCE_LINE.append(line)

    # === 7. 整理输出 ===
    Shui_lst = [d, Fw_front]
    return WATER_FORCE_LINE, Shui_values, Shui_lst



# 荷载组合
def build_loadcomb_lines(load_tab):
    """
    生成 *LOADCOMB 段（车辆、标准、基本、非工作标准、非工作基本组合）

    参数：
        load_tab: LoadTab 实例（用于读取移动荷载工况名称）

    返回：
        list[str]: MCT命令行列表
    """
    loadcomb_lines = ["*LOADCOMB"]

    # === 1. 读取所有移动荷载工况名 ===
    move_case_names = []
    if hasattr(load_tab, "case_list"):
        for case in load_tab.case_list:
            name = case["name"].get().strip()
            if name:
                move_case_names.append(name)

    # === 2. 车辆荷载组合 ===
    loadcomb_lines.append("   NAME=车辆荷载, GEN, ACTIVE, 0, 1, , 0, 0, 0, 1")
    if move_case_names:
        # 每个工况都以 MV 开头
        mv_parts = []
        for n in move_case_names:
            mv_parts.append(f"MV, {n}, 1")
        mv_line = "        " + ", ".join(mv_parts)
        loadcomb_lines.append(mv_line)

    # === 3. 标准组合 ===
    loadcomb_lines.append("   NAME=标准组合, GEN, ACTIVE, 0, 0, , 0, 0, 0, 1")
    loadcomb_lines.append("        ST, 自重, 1, ST, 风荷载, 0.7, ST, 水流力, 0.7, CB, 车辆荷载, 1")

    # === 4. 基本组合 ===
    loadcomb_lines.append("   NAME=基本组合, GEN, ACTIVE, 0, 0, , 0, 0, 0, 1")
    loadcomb_lines.append("        ST, 自重, 1.2, ST, 风荷载, 0.98, ST, 水流力, 0.98, CB, 车辆荷载, 1.4")

    return loadcomb_lines