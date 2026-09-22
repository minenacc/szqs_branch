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
        "虚拟梁": ", CC, 0, 0, 0, 0, 0, 0, YES, NO, SB, 2, 0.01, 0.01, 0, 0, 0, 0, 0, 0, 0, 0",
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

    return tuple(names), tuple(cad_params), tuple(weights)
 
# ------------------------------------------------
# 节点参数
# ------------------------------------------------

# ==========生成小肋节点信息==========
def build_small_rib_nodes(components_tab):
    """
    从UI读取参数生成小肋节点。
    小肋：每根沿y方向布置（x相同）。节点包括两端、贝雷梁交点、中点（虚拟梁交点）。

    返回：
        xl_nodes                —— {nid: (x, y, z)}                   # 全部小肋节点
        xl_single_nodes         —— {rib_idx: [(nid, (x, y, z)), ...]} # 每根小肋节点
        xl_bailey_elink_nodes   —— {nid: (x, y, z)}                   # 小肋-贝雷梁交点（含虚拟梁）
        xl_virtual_elink_nodes  —— {nid: (x, y, z)}                   # 小肋-虚拟梁交点（中点）
        xl_length               —— 单根小肋总长度
        node_id                 —— 下一个可用节点号
    """
    # -------- 1. 读取UI参数 --------
    smallrib_expr = components_tab.entry_smallrib.get().strip()         # 例如 "29@400"
    bailey_expr = components_tab.entry_bailey_spacing.get().strip()     # 例如 "3@900,1500,3@900"
    xl_length = float(components_tab.entry_smallrib_length.get() or 8000.0)  # 悬臂长度
    cantilever_len = float(components_tab.entry_bailey_init_spacing.get() or 550.0)  # 悬臂长度

    # -------- 2. 计算布置 --------
    xl_xs = midasdisttolst2(0.0, smallrib_expr)            # 每根小肋的x坐标
    bailey_ys = midasdisttolst2(cantilever_len, bailey_expr)  # 贝雷梁y坐标
    
    # -------- 3. 初始化容器 --------
    xl_nodes = {}                # 全部节点
    xl_single_nodes = {}         # 每根小肋的节点
    xl_bailey_elink_nodes = {}   # 小肋-贝雷梁弹连点
    xl_virtual_elink_nodes = {}  # 小肋-虚拟梁弹连点

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

            # 虚拟梁弹连点（中点）
            if abs(y - y_mid) < 1e-6:
                xl_virtual_elink_nodes[node_id] = coord
            # 贝雷梁弹连点
            if any(abs(y - yy) < 1e-6 for yy in bailey_ys):
                xl_bailey_elink_nodes[node_id] = coord

            node_id += 1

    return xl_nodes, xl_single_nodes, xl_bailey_elink_nodes, xl_virtual_elink_nodes, xl_length, node_id

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
    print(bailey_span_total)

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

    # 合并模板（去重）
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


def to_float_safe(val):
    """将字符串转换为 float，失败则返回 0.0"""
    try:
        return float(val)
    except:
        return 0.0
    
# ==========生成贝雷梁下弦杆节点信息==========
def build_bailey_bottom_nodes(
    components_tab,
    bailey_top_vertical_nodes: dict,
    bailey_top_diagonal_nodes: dict,
    bailey_top_release_nodes: dict,
    z_top: float,
    start_id: int = 1,
    section_data: dict = None,
    ):
    """
    生成贝雷梁下弦杆节点坐标。
    分类：
        bailey_bottom_elink_nodes: 类型 1（与分配梁弹连点）
        bailey_bottom_vertical_nodes: 类型 2（竖杆共节点）
        bailey_bottom_diagonal_nodes: 类型 3（斜杆共节点）
        bailey_bottom_release_nodes: 类型 4（片段连接释放点）
        bailey_bottom_abutment_nodes: 类型 5（桥台固结）
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

    left_support = components_tab.left_select_var.get() 
    right_support = components_tab.right_select_var.get()

    # -------------------------------
    # 2. z 坐标上弦杆z坐标 - 1.4m
    # ------------------------------
    z_bottom = z_top - 1400

    # -------------------------------
    # 3. x坐标生成
    # ------------------------------
    # 端点坐标
    x_base = [x0, x_end]

    # 每跨的交界 x 坐标(如输入单跨则根据逻辑不生成)
    x_span_joints = []
    acc = x0
    for L in bailey_span_steps[:-1]:
        acc += L
        x_span_joints.append(acc)

    # 与钢管桩弹连/制动墩弹连/桥台固结判断
    bailey_bottom_elink_x = []          # 钢管桩 / 分配梁弹连
    bailey_bottom_brake_elink_x = []    # 制动墩分配梁 F1 弹连
    

    # ---- 左端 ----
    if left_support == "钢管桩":
        left_canti_raw = components_tab.left_entry_ggz_cantilever.get().strip()
        left_ggz_cantilever = to_float_safe(left_canti_raw)
        left_ggz_cantilever_x = x0 + left_ggz_cantilever
        bailey_bottom_elink_x.append(left_ggz_cantilever_x)

    elif left_support == "制动墩":
        B_f1 = float(section_data.get("制动墩分配梁F1", {}).get("params", {}).get("B(mm)", 0) or 600)
        x_brake_left = x0 - 60.0 + B_f1 / 2.0
        bailey_bottom_brake_elink_x.append(x_brake_left)

    elif left_support == "桥台":
        # 完全不读取悬臂
        pass

    # ---- 右端 ----
    if right_support == "钢管桩":
        right_canti_raw = components_tab.right_entry_ggz_cantilever.get().strip()
        right_ggz_cantilever = to_float_safe(right_canti_raw)
        right_ggz_cantilever_x = x_end - right_ggz_cantilever
        bailey_bottom_elink_x.append(right_ggz_cantilever_x)

    elif right_support == "制动墩":
        B_f1 = float(section_data.get("制动墩分配梁F1", {}).get("params", {}).get("B(mm)", 0) or 600)
        x_brake_right = x_end + 60.0 - B_f1 / 2.0
        bailey_bottom_brake_elink_x.append(x_brake_right)

    elif right_support == "桥台":
        # 完全不读取悬臂
        pass

    # ---- 中间跨的分界点始终为弹连点 ----
    bailey_bottom_elink_x.extend(x_span_joints)

    # 去重排序
    bailey_bottom_elink_x = sorted(set(bailey_bottom_elink_x))
    bailey_bottom_brake_elink_x = sorted(set(bailey_bottom_brake_elink_x))

    # 类型 234. 复用 XY 
    def shift_z(node_dict):
        return {nid: (x, y, z_bottom) for nid, (x, y, _) in node_dict.items()}

    bailey_bottom_vertical_nodes = shift_z(bailey_top_vertical_nodes)
    bailey_bottom_diagonal_nodes = shift_z(bailey_top_diagonal_nodes)
    bailey_bottom_release_nodes = shift_z(bailey_top_release_nodes)

    # 用于桥台固结点判定（最左 / 最右竖杆）
    vertical_xs = sorted({coord[0] for coord in bailey_bottom_vertical_nodes.values()})
    x_abutment_left = vertical_xs[0] if vertical_xs else None
    x_abutment_right = vertical_xs[-1] if vertical_xs else None

    # 汇总节点x坐标
    x_all = sorted(
        set(
            x_base
            + bailey_bottom_elink_x + bailey_bottom_brake_elink_x
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
    bailey_bottom_brake_elink_nodes = {}
    bailey_bottom_abutment_nodes = {}
    bailey_bottom_single_nodes = {}
    bailey_bottom_start_nodes = {}
    bailey_bottom_end_nodes = {}

    # 从上弦杆任一类节点取 y 集合
    y_list = sorted({coord[1] for coord in bailey_top_vertical_nodes.values()})
    vertical_x_set = set(vertical_xs)
    diagonal_x_set = {coord[0] for coord in bailey_bottom_diagonal_nodes.values()}
    release_x_set = {coord[0] for coord in bailey_bottom_release_nodes.values()}
    eps = 1e-6

    node_id = start_id
    for j, y in enumerate(y_list, start=1):
        bailey_bottom_single_nodes[j] = []
        for x in x_all:
            coord = (x, y, z_bottom)
            bailey_bottom_nodes[node_id] = coord
            bailey_bottom_single_nodes[j].append((node_id, coord))
            # 弹连点分类
            if x in bailey_bottom_elink_x:
                bailey_bottom_elink_nodes[node_id] = coord
            if x in bailey_bottom_brake_elink_x:
                bailey_bottom_brake_elink_nodes[node_id] = coord

            # 竖杆 / 斜杆 / 释放
            if x in vertical_x_set:
                bailey_bottom_vertical_nodes[node_id] = coord
            if x in diagonal_x_set:
                bailey_bottom_diagonal_nodes[node_id] = coord
            if x in release_x_set:
                bailey_bottom_release_nodes[node_id] = coord

            # 桥台固结点（最左 / 最右竖杆坐标）
            if left_support == "桥台" and x_abutment_left is not None and abs(x - x_abutment_left) < eps:
                bailey_bottom_abutment_nodes[node_id] = coord
            if right_support == "桥台" and x_abutment_right is not None and abs(x - x_abutment_right) < eps:
                bailey_bottom_abutment_nodes[node_id] = coord

            # 起 / 终端点定位（用于建单元）
            if abs(x - x0) < eps:
                bailey_bottom_start_nodes[node_id] = coord
            if abs(x - x_end) < eps:
                bailey_bottom_end_nodes[node_id] = coord
            node_id += 1

    # --------------------------------------
    # 5. 返回结果
    # --------------------------------------
    return (
        bailey_bottom_nodes,
        bailey_bottom_elink_nodes,       # 与分配梁弹连节点
        bailey_bottom_brake_elink_nodes, # 与制动墩分配梁F1弹连节点
        bailey_bottom_abutment_nodes,    # 桥台固结节点
        bailey_bottom_vertical_nodes,    # 竖杆共节点
        bailey_bottom_diagonal_nodes,    # 斜杆共节点
        bailey_bottom_release_nodes,     # 释放约束节点
        bailey_bottom_single_nodes,      # 每根贝雷梁节点归类
        bailey_bottom_start_nodes,       # 起始端点
        bailey_bottom_end_nodes,         # 终止端点
        z_bottom,
        x_end,
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
    bailey_bottom_elink_nodes: dict,        # 普通贝雷梁下弦杆弹连节点
    bailey_bottom_brake_elink_nodes: dict,  # 制动墩段下弦杆弹连节点
    x_end: float,
    xl_length: float,
    z_bottom: float = None,                 # 贝雷梁下弦杆 z 坐标
    start_id: int = 1,
    section_data: dict = None               # 截面数据
):
    """
    构建分配梁节点（普通 + 制动墩两部分独立生成）
    返回:
        distribute_nodes, distribute_single_nodes, distribute_bottom_elink_nodes, distribute_pile_elink_nodes,
        brake_f2_nodes, brake_f2_bottom_elink_nodes, brake_f2_f1_elink_nodes,
        brake_f1_nodes, brake_f1_f2_elink_nodes, brake_f1_pile_elink_nodes, node_id
    """

    # ====================== 读取参数 ======================
    # 是否设置制动墩
    left_support = components_tab.left_select_var.get()    # "钢管桩" / "制动墩" / "桥台"
    right_support = components_tab.right_select_var.get()  # "钢管桩" / "制动墩" / "桥台"
    need_left_brake = (left_support == "制动墩")
    need_right_brake = (right_support == "制动墩")

    both_brake = (left_support == "制动墩" and right_support == "制动墩")
    if both_brake:
        entry_zdd_y = components_tab.left_entry_zdd_y
        entry_zdd_x = components_tab.left_entry_zdd_x
        entry_f1_length    = components_tab.left_entry_f1_length
        entry_f1_cantilever1 = components_tab.left_entry_f1_cantilever1
        entry_f1_cantilever2 = components_tab.left_entry_f1_cantilever2
        entry_f2_length    = components_tab.left_entry_f2_length
        entry_f2_cantilever1 = components_tab.left_entry_f2_cantilever1
        entry_f2_cantilever2 = components_tab.left_entry_f2_cantilever2
    elif left_support == "制动墩":
        entry_zdd_y = components_tab.left_entry_zdd_y
        entry_zdd_x = components_tab.left_entry_zdd_x
        entry_f1_length    = components_tab.left_entry_f1_length
        entry_f1_cantilever1 = components_tab.left_entry_f1_cantilever1
        entry_f1_cantilever2 = components_tab.left_entry_f1_cantilever2
        entry_f2_length    = components_tab.left_entry_f2_length
        entry_f2_cantilever1 = components_tab.left_entry_f2_cantilever1
        entry_f2_cantilever2 = components_tab.left_entry_f2_cantilever2
    elif right_support == "制动墩":
        entry_zdd_y = components_tab.right_entry_zdd_y
        entry_zdd_x = components_tab.right_entry_zdd_x
        entry_f1_length    = components_tab.right_entry_f1_length
        entry_f1_cantilever1 = components_tab.right_entry_f1_cantilever1
        entry_f1_cantilever2 = components_tab.right_entry_f1_cantilever2
        entry_f2_length    = components_tab.right_entry_f2_length
        entry_f2_cantilever1 = components_tab.right_entry_f2_cantilever1
        entry_f2_cantilever2 = components_tab.right_entry_f2_cantilever2
    else:
        entry_zdd_y = entry_zdd_x = entry_f1_length = entry_f1_cantilever = entry_f2_length = entry_f2_cantilever1 = entry_f2_cantilever2 = None

    # 分配梁节点需要参数
    distribute_length =float(components_tab.entry_distribute_length.get() or 7500.0) # 分配梁长度
    distribute_cantilever_bailey = float(components_tab.entry_distribute_cantilever.get() or 300.0) # 分配梁相对贝雷梁悬臂长度
    distribute_cantilever_pile = float(components_tab.entry_distribute_cantilever2.get() or 750.0) # 分配梁相对桩悬臂长度
    y_list = components_tab.entry_pile_y or 3000+3000 # 钢管桩纵向布置
    # 截面参数
    H_distributed = float(section_data.get("分配梁", {}).get("params", {}).get("H(mm)", 450))
    z_distribute = z_bottom - H_distributed / 2 - 50
    
    # ====================== 坐标分类计算 ======================
    # 获取普通贝雷梁下弦坐标
    x_coords = sorted({coord[0] for coord in bailey_bottom_elink_nodes.values()})
    y_coords = sorted({coord[1] for coord in bailey_bottom_elink_nodes.values()})

    # 普通分配梁端点
    y_start = min(y_coords) - distribute_cantilever_bailey
    y_end =y_start + distribute_length

    # 桩排布
    pile_y_list = midasdisttolst(y_list.get())
    y_start_pile = y_start + distribute_cantilever_pile
    pile_y_coords = [y_start_pile]
    for dist in pile_y_list:
        pile_y_coords.append(pile_y_coords[-1] + dist)

    # 普通分配梁节点y坐标整合
    y_list = sorted(set([y_start, y_end, *pile_y_coords, *y_coords]))

    # ====================== 初始化 ====================== 
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

    # 若没有任何制动墩 → 无需生成 F1/F2
    if not (need_left_brake or need_right_brake):
        return (
            distribute_nodes,
            distribute_single_nodes,
            distribute_bottom_elink_nodes,
            distribute_pile_elink_nodes,
            {}, {}, {},
            {}, {}, {},
            node_id
        )

    # ============================================================
    # 制动墩 F1 / F2 梁（共享参数）
    # ============================================================
    # 共享参数
    H_f1 = float(section_data.get("制动墩分配梁F1", {}).get("params", {}).get("H(mm)", 450))
    H_f2 = float(section_data.get("制动墩分配梁F2", {}).get("params", {}).get("H(mm)", 450))
    z_f1 = z_bottom - H_f1 / 2 - 50
    z_f2 = z_f1 - H_f1 / 2 - H_f2 / 2


    f1_length = float(entry_f1_length.get() or 7500.0)
    f1_cantilever_bailey = float(entry_f1_cantilever1.get() or 300.0)
    f1_cantilever_f2 = float(entry_f1_cantilever2.get() or 750.0)

    f2_length = float(entry_f2_length.get() or 3000.0)
    f2_cantilever_f1 = float(entry_f2_cantilever1.get() or 1640.0)
    f2_cantilever_pile = float(entry_f2_cantilever2.get() or 500.0)

    zdd_x = entry_zdd_x.get() or "3000"
    zdd_y = entry_zdd_y.get() or "2000+2000"
    print(zdd_x, zdd_y)
    brake_pile_x_list = midasdisttolst(zdd_x)
    brake_pile_y_list = midasdisttolst(zdd_y)
    print(brake_pile_x_list, brake_pile_y_list)

    # 初始化
    brake_f1_nodes = {}
    brake_f1_bottom_elink_nodes = {}
    brake_f1_f2_elink_nodes = {}
    brake_f2_nodes = {}
    brake_f2_f1_elink_nodes = {}
    brake_f2_pile_elink_nodes = {}


    # ============== F1 ==============
    # 制动墩段的下弦弹连坐标
    x_f1_elink = sorted({coord[0] for coord in bailey_bottom_brake_elink_nodes.values()})
    y_f1_elink = sorted({coord[1] for coord in bailey_bottom_brake_elink_nodes.values()})
    # y坐标
    y_f1_start = min(y_f1_elink) - f1_cantilever_bailey
    y_f1_end = y_f1_start + f1_length

    y_brake_pile_start = y_f1_start + f1_cantilever_f2
    brake_pile_y_coords = [y_brake_pile_start]
    for dist in brake_pile_y_list:
        brake_pile_y_coords.append(brake_pile_y_coords[-1] + dist)

    y_list_f1 = sorted(set([
        y_f1_start,
        y_f1_end,
        *brake_pile_y_coords,
        *y_f1_elink,
    ]))
    for i, x in enumerate(x_f1_elink, start=1):
        brake_f1_nodes[i] = []
        for y in y_list_f1:
            coord = (x, y, z_f1)
            brake_f1_nodes[i].append((node_id, coord))
            if any(abs(y - yy) < 1e-6 for yy in y_f1_elink):
                brake_f1_bottom_elink_nodes[node_id] = coord
            if any(abs(y - yc) < 1e-6 for yc in brake_pile_y_coords):
                brake_f1_f2_elink_nodes[node_id] = coord
            node_id += 1

    # ============== F2 ==============
    x_f1_sorted = sorted(x_f1_elink)
    eps = 1e-6

    # F2按组编号，从 1 开始连续编号，左右两侧共用同一个字典
    f2_group_idx = 1

    # ---------- 左侧制动墩 ----------
    if need_left_brake and x_f1_sorted:
        x_f1_left = x_f1_sorted[0]

        # F2 左端点与右端点
        x_start_f2_left = x_f1_left - f2_cantilever_f1
        x_end_f2_left = x_start_f2_left + f2_length

        # X 方向桩坐标
        x_start_f2_pile_left = x_start_f2_left + f2_cantilever_pile
        pile_x_coords_left = [x_start_f2_pile_left]
        for dx in brake_pile_x_list:
            pile_x_coords_left.append(pile_x_coords_left[-1] + dx)

        # Y 方向桩坐标
        for i, y in enumerate(brake_pile_y_coords, start=1):
            group_id = f2_group_idx
            f2_group_idx += 1
            brake_f2_nodes[group_id] = []

            # 这一排所有 X 坐标
            x_list = sorted(set(
                [x_start_f2_left] +
                pile_x_coords_left +     
                [x_f1_left, x_end_f2_left]
            ))

            # 生成节点
            for x in x_list:
                coord = (x, y, z_f2)
                brake_f2_nodes[group_id].append((node_id, coord))

                # F1 弹连
                if abs(x - x_f1_left) < eps:
                    brake_f2_f1_elink_nodes[node_id] = coord

                # 桩弹连：遍历全部 pile_x_coords_left
                if any(abs(x - px) < eps for px in pile_x_coords_left):
                    brake_f2_pile_elink_nodes[node_id] = coord

                node_id += 1

    # ---------- 右侧制动墩 ----------
    if need_right_brake and x_f1_sorted:
        x_f1_right = x_f1_sorted[-1]

        # F2 端点
        x_start_f2_right = x_f1_right - f2_cantilever_f1
        x_end_f2_right = x_start_f2_right + f2_length

        # 桩弹连 X 坐标（支持多列桩）
        x_start_f2_pile_right = x_start_f2_right + f2_cantilever_pile
        pile_x_coords_right = [x_start_f2_pile_right]
        for dx in brake_pile_x_list:
            pile_x_coords_right.append(pile_x_coords_right[-1] + dx)

        # 多排桩 → 多根 F2
        for i, y in enumerate(brake_pile_y_coords, start=1):
            group_id = f2_group_idx
            f2_group_idx += 1
            brake_f2_nodes[group_id] = []

            # 一根 F2 梁的所有 X 坐标
            x_list = sorted(set(
                [x_start_f2_right] +
                pile_x_coords_right +   # 多列桩全部加入
                [x_f1_right, x_end_f2_right]
            ))

            # 生成节点
            for x in x_list:
                coord = (x, y, z_f2)
                brake_f2_nodes[group_id].append((node_id, coord))

                # 与 F1 弹连
                if abs(x - x_f1_right) < eps:
                    brake_f2_f1_elink_nodes[node_id] = coord

                # 与制动墩桩弹连（多列桩任意匹配）
                if any(abs(x - px) < eps for px in pile_x_coords_right):
                    brake_f2_pile_elink_nodes[node_id] = coord

                node_id += 1

    # =========================================================
    # 返回
    # =========================================================
    return (
        distribute_nodes,
        distribute_single_nodes,
        distribute_bottom_elink_nodes,
        distribute_pile_elink_nodes,
        brake_f1_nodes,
        brake_f1_bottom_elink_nodes,
        brake_f1_f2_elink_nodes,
        brake_f2_nodes,
        brake_f2_f1_elink_nodes,
        brake_f2_pile_elink_nodes,
        node_id,
    )
    
# ==========生成钢管桩节点信息==========
def build_pile_nodes(
    components_tab,
    distribute_pile_elink_nodes: dict,
    brake_f2_pile_elink_nodes: dict,
    start_id: int = 1,
    section_data: dict = None,
    load_tab=None
):
    """
    构建普通钢管桩与制动墩钢管桩节点信息（含桩顶、联结系、桩底、合力点）
    并按 y 排（桩排方向）汇总水流力合力点。

    返回：
        pile_top_nodes, pile_connect_nodes, pile_bottom_nodes, pile_group_nodes,
        brake_pile_top_nodes, brake_pile_connect_nodes, brake_pile_bottom_nodes, brake_pile_group_nodes,
        water_force_points_by_y, water_force_points_brake_by_y,
        node_id
    """

    # 参数读取
    connect_zd = float(components_tab.entry_connect_zd.get() or 1000.0)
    connect_z  = float(components_tab.entry_connect_z.get() or 2000.0)
    pile_length        = float(components_tab.entry_pile_length.get()  or 15.0) * 1000.0

    # 制动墩钢管桩参数
    left_support  = components_tab.left_select_var.get()
    right_support = components_tab.right_select_var.get()

    left_has_brake  = (left_support  == "制动墩")
    right_has_brake = (right_support == "制动墩")

    # 默认值（单位 m）
    left_len  = float(components_tab.left_entry_zdd_length.get()  or 15.0)
    right_len = float(components_tab.right_entry_zdd_length.get() or 15.0)

    if left_has_brake and right_has_brake:
        # 两端制动墩 → 右端视为相同设置
        brake_pile_length = left_len * 1000.0
    elif left_has_brake:
        brake_pile_length = left_len * 1000.0
    elif right_has_brake:
        brake_pile_length = right_len * 1000.0
    else:
        # 没有制动墩时不生成
        brake_pile_length = None

    # 初始化
    pile_group_nodes, brake_pile_group_nodes = {}, {}
    pile_top_nodes, pile_connect_nodes, pile_bottom_nodes = {}, {}, {}
    brake_pile_top_nodes, brake_pile_connect_nodes, brake_pile_bottom_nodes = {}, {}, {}
    node_id = start_id

    # 按 y 排分类的水流力作用点
    water_force_points_by_y = {}
    water_force_points_brake_by_y = {}

    # 荷载参数
    if load_tab:
        water_level = float(load_tab.water_level_entry.get() or 0.0)
        water_depth = float(load_tab.water_depth_entry.get() or 0.0)
    else:
        water_level = +10.0
        water_depth = 5.0

    deck_elev      = float(components_tab.entry_deck_height.get() or 0.0)
    deck_thickness = float(components_tab.entry_deck.get() or 0.0)

    # 截面参数（mm→m）
    def H(name):
        return float(section_data.get(name, {}).get("params", {}).get("H(mm)", 0.0))

    H_small, H_bailey, H_dist, H_f1, H_f2 = map(H, ["小肋", "贝雷梁", "分配梁", "制动墩分配梁F1", "制动墩分配梁F2"])

    # 1.普通钢管桩
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

    if brake_pile_length is None:
        return (
            pile_top_nodes, pile_connect_nodes, pile_bottom_nodes, pile_group_nodes,
            brake_pile_top_nodes, brake_pile_connect_nodes, brake_pile_bottom_nodes, brake_pile_group_nodes,
            water_force_points_by_y, water_force_points_brake_by_y,
            node_id
        )
    
    # 2.制动墩钢管桩
    ztop_elev_b = deck_elev - deck_thickness/1000 - H_small/2000 - H_bailey/1000 - 1.4 - H_f1/1000 - H_f2/1000
    pile_by_xb = {}
    seen_coords = set()
    for _, (x, y, z) in brake_f2_pile_elink_nodes.items():
        zt = z - H_f2 / 2
        key = (round(x, 3), round(y, 3), round(zt, 3))
        if key in seen_coords:
            continue
        seen_coords.add(key)
        pile_by_xb.setdefault(x, []).append((x, y, zt))

    gid = 1
    for x_val, arr in sorted(pile_by_xb.items()):
        arr.sort(key=lambda t: t[1])
        brake_pile_group_nodes[gid] = []

        for x, y, zt in arr:
            zc1 = zt - connect_zd
            zc2 = zc1 - connect_z
            zbot = zt - brake_pile_length
            zf = zt - (ztop_elev_b - water_level + 0.3 * water_depth) * 1000
            zf = round(zf, 1)

            cands = [
                ("桩顶", zt),
                ("合力点", zf),
                ("联结系上", zc1),
                ("联结系下", zc2),
                ("桩底", zbot),
            ]
            cands.sort(key=lambda t: t[1], reverse=True)

            for tag, zv in cands:
                coord = (x, y, zv)
                if tag == "桩顶":
                    brake_pile_top_nodes[node_id] = coord
                elif "联结系" in tag:
                    brake_pile_connect_nodes[node_id] = coord
                elif tag == "桩底":
                    brake_pile_bottom_nodes[node_id] = coord
                elif tag == "合力点":
                    water_force_points_brake_by_y.setdefault(y, []).append((node_id, coord))
                brake_pile_group_nodes[gid].append((node_id, coord))
                node_id += 1
        gid += 1

    # 返回
    return (
        pile_top_nodes, pile_connect_nodes, pile_bottom_nodes, pile_group_nodes,
        brake_pile_top_nodes, brake_pile_connect_nodes, brake_pile_bottom_nodes, brake_pile_group_nodes,
        water_force_points_by_y, water_force_points_brake_by_y,
        node_id
    )

# ==========生成联结系节点信息==========
def build_connect_x_nodes(
    components_tab,
    pile_group_nodes,                 # {组号: [(nid,(x,y,z)), ...]}
    pile_connect_nodes,               # {nid: (x,y,z)}  ← 新增直接使用
    brake_pile_group_nodes,           # {组号: [(nid,(x,y,z)), ...]}
    brake_pile_connect_nodes,         # {nid: (x,y,z)}  ← 新增直接使用
    start_id: int = 1
):
    """
    从已知的“联结系节点字典”中抽取每根桩的上/下联结点，
    计算相邻两根桩（横向）或相邻两列桩（纵向，制动墩）之间的 X 形中心节点。

    返回：
        connect_x_nodes:        {组号: (mid_nid, (x, y, z))}
        connect_groups:         {组号: [(左上),(左下),(中点),(右上),(右下)]}
        brake_connect_nodes:    {组号: (mid_nid, (x, y, z))}
        brake_connect_groups:   {组号: [(上1),(下1),(中点),(上2),(下2)]}
        node_id:                最后编号
    """
    tol = 1e-6
    connect_x_nodes = {}
    connect_groups = {}
    brake_connect_nodes = {}
    brake_connect_groups = {}

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

    # 1.普通钢管桩：横向 X 形（同 x、邻 y）
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

    # 记录普通桩结束后的编号
    start_id_after_normal = node_id

    # 2.制动墩钢管桩：横向 + 纵向
    left_support  = components_tab.left_select_var.get()
    right_support = components_tab.right_select_var.get()
    has_brake = ("制动墩" in {left_support, right_support})

    if has_brake:
        node_id = start_id_after_normal
        gid_out = 1

        # --- 按每根桩 (x,y) 收集 up/dn ---
        all_piles = []
        for _, group_nodes in sorted(brake_pile_group_nodes.items()):
            pile_map_y = build_pile_map_by_y(group_nodes, brake_pile_connect_nodes)
            for y, rec in pile_map_y.items():
                all_piles.append({"x": rec["x"], "y": y, "up": rec["up"], "dn": rec["dn"]})

        if not all_piles:
            return connect_x_nodes, connect_groups, brake_connect_nodes, brake_connect_groups, node_id

        # --- 将制动墩按左右分区 ---
        xs = sorted({p["x"] for p in all_piles})
        mid_x = (xs[0] + xs[-1]) / 2.0       # 用中点区分左右
        left_piles  = [p for p in all_piles if p["x"] <= mid_x]
        right_piles = [p for p in all_piles if p["x"] >  mid_x]

        def build_one_side(pile_list):
            nonlocal node_id, gid_out
            if not pile_list:
                return

            # 横向（同 x、邻 y）
            piles_by_x = {}
            for p in pile_list:
                piles_by_x.setdefault(p["x"], []).append(p)
            for x, lst in piles_by_x.items():
                lst.sort(key=lambda r: r["y"])
                for i in range(len(lst) - 1):
                    p1, p2 = lst[i], lst[i + 1]
                    y_mid = (p1["y"] + p2["y"]) / 2.0
                    z_mid = (p1["up"][1][2] + p1["dn"][1][2] + p2["up"][1][2] + p2["dn"][1][2]) / 4.0
                    mid_coord = (x, y_mid, z_mid)
                    brake_connect_nodes[gid_out] = (node_id, mid_coord)
                    brake_connect_groups[gid_out] = [p1["up"], p1["dn"], (node_id, mid_coord), p2["up"], p2["dn"]]
                    node_id += 1
                    gid_out += 1

            # 纵向（同 y、邻 x） —— 仅在该侧内部
            piles_by_y = {}
            for p in pile_list:
                piles_by_y.setdefault(p["y"], []).append(p)
            for y, lst in piles_by_y.items():
                lst.sort(key=lambda r: r["x"])
                for i in range(len(lst) - 1):
                    p1, p2 = lst[i], lst[i + 1]
                    x_mid = (p1["x"] + p2["x"]) / 2.0
                    z_mid = (p1["up"][1][2] + p1["dn"][1][2] + p2["up"][1][2] + p2["dn"][1][2]) / 4.0
                    mid_coord = (x_mid, y, z_mid)
                    brake_connect_nodes[gid_out] = (node_id, mid_coord)
                    brake_connect_groups[gid_out] = [p1["up"], p1["dn"], (node_id, mid_coord), p2["up"], p2["dn"]]
                    node_id += 1
                    gid_out += 1

        # 左、右分别执行（互不跨连接）
        build_one_side(left_piles)
        build_one_side(right_piles)

    return connect_x_nodes, connect_groups, brake_connect_nodes, brake_connect_groups, node_id

# ==========生成虚拟梁节点信息==========
def build_virtual_beam_nodes(
    xl_virtual_elink_nodes: dict,
    section_data: dict,
    start_id: int = 1
):
    """
    生成虚拟梁节点。
    逻辑：
        - 来源：小肋虚拟弹连节点 xl_virtual_elink_nodes；
        - 沿 x 方向布置，仅一根；
        - 修改 z 坐标：z_virtual = 0 + 小肋截面高度 / 2；
        - 按 x 从小到大排序；
        - 返回:
            virtual_nodes: {nid: (x, y, z)}
            virtual_single_nodes: {1: [(nid, (x,y,z)), ...]} （单组）
            next_node_id: 下一可用节点号
    """
    # 获取小肋高度
    H_smallrib = float(section_data.get("小肋", {}).get("params", {}).get("H(mm)", 120.0))
    z_virtual = 0.0 + H_smallrib / 2.0

    # 按 x 排序
    sorted_nodes = sorted(xl_virtual_elink_nodes.items(), key=lambda kv: kv[1][0])

    virtual_nodes = {}
    virtual_single_nodes = {1: []}
    node_id = start_id

    for _, (x, y, z) in sorted_nodes:
        coord = (x, y, z_virtual)
        virtual_nodes[node_id] = coord
        virtual_single_nodes[1].append((node_id, coord))
        node_id += 1

    next_node_id = node_id
    return virtual_nodes, virtual_single_nodes, node_id


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
    brake_f1_nodes: dict,
    brake_f2_nodes: dict,
    start_elem_id=1
):
    """分配梁单元生成函数"""
    elem_id = start_elem_id
    all_distribute_elements = []

    # 普通分配梁
    distribute_elems, elem_id = build_beam_elements(
        distribute_single_nodes, "分配梁", section_data, name_to_id, elem_id
    )
    all_distribute_elements.extend(distribute_elems)

    left_support  = components_tab.left_select_var.get()
    right_support = components_tab.right_select_var.get()
    has_brake = ("制动墩" in {left_support, right_support})

    # 若启用制动墩
    if has_brake:
        # F1 分配梁
        f1_elems, elem_id = build_beam_elements(
            brake_f1_nodes, "制动墩分配梁F1", section_data, name_to_id, elem_id
        )
        all_distribute_elements.extend(f1_elems)

        # F2 分配梁
        f2_elems, elem_id = build_beam_elements(
            brake_f2_nodes, "制动墩分配梁F2", section_data, name_to_id, elem_id
        )
        all_distribute_elements.extend(f2_elems)

    next_elem_id = elem_id
    return all_distribute_elements, next_elem_id

# ==========生成钢管桩单元信息==========
def build_pile_elements(
    section_data: dict,
    name_to_id: dict,
    pile_group_nodes: dict,
    brake_pile_group_nodes: dict,
    start_elem_id=1
):
    """
    生成钢管桩 BEAM 单元。
    逻辑：
        - 普通钢管桩按 x 分组（同 x 值为一列）；
        - 每组内节点按 (x,y) 相同，z 坐标排序连接；
        - 制动墩钢管桩所有为一组，同理；
        - 材料号 2（Q235），β角 = 0。
    参数：
        section_data            : 截面参数全集
        name_to_id              : 截面名 → ID 映射
        pile_group_nodes        : 普通钢管桩 {x组号: [(nid, (x,y,z)), ...]}
        brake_pile_group_nodes  : 制动墩钢管桩 {1: [(nid, (x,y,z)), ...]}
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

    # ========== 1. 普通钢管桩 ==========
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
    
    # ========== 2. 制动墩钢管桩 ==========
    if brake_pile_group_nodes:
        section_name = "制动墩钢管桩"
        section_id = name_to_id.get(section_name, 1)
        all_brake_nodes = []
        for _, nodes in brake_pile_group_nodes.items():
            all_brake_nodes.extend(nodes)

        if all_brake_nodes:
            from collections import defaultdict

            # 分组逻辑：按 (x, y) 精确分类
            piles_by_xy = defaultdict(list)
            for nid, (x, y, z) in all_brake_nodes:
                key = (round(x, 3), round(y, 3))
                piles_by_xy[key].append((nid, z))

            # 每根桩竖向排序并连线
            for (x_key, y_key), nid_z_list in piles_by_xy.items():
                nid_z_list.sort(key=lambda nz: nz[1])  # 按 z 由上到下排序
                for i in range(len(nid_z_list) - 1):
                    n1, _ = nid_z_list[i]
                    n2, _ = nid_z_list[i + 1]
                    pile_elements.append(
                        f"{elem_id}, BEAM, {mat_id}, {section_id}, {n1}, {n2}, {beta_angle:.1f}, 0"
                    )
                    elem_id += 1

    next_elem_id = elem_id
    return pile_elements, next_elem_id

# ==========生成联结系单元信息==========
def build_connect_x_elements(
    section_data: dict,
    name_to_id: dict,
    connect_groups: dict,
    brake_connect_groups: dict,
    start_elem_id=1
):
    """
    生成普通与制动墩联结系 BEAM 单元。
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
        brake_connect_groups  : 制动墩联结系 {group_id: [(nid, coord), ...]}
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

    # ========== 1. 普通联结系 ==========
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

    # ========== 2. 制动墩联结系 ==========
    for gid, nodes in sorted(brake_connect_groups.items()):
        if not nodes or len(nodes) < 5:
            continue
        up_left, down_left, mid, up_right, down_right = nodes

        connect_pairs = [
            (up_left[0], up_right[0]),
            (down_left[0], down_right[0]),
            (up_left[0], mid[0]),
            (mid[0], down_right[0]),
            (down_left[0], mid[0]),
            (mid[0], up_right[0]),
        ]

        for n1, n2 in connect_pairs:
            connect_elements.append(
                f"{elem_id}, BEAM, {mat_id}, {section_id}, {n1}, {n2}, {beta_angle}, 0"
            )
            elem_id += 1

    next_elem_id = elem_id
    return connect_elements, next_elem_id

# ==========生成虚拟梁单元信息==========
def build_virtual_beam_elements(
    virtual_single_nodes: dict,
    section_data: dict,
    name_to_id: dict,
    start_elem_id=1
):
    """
    生成虚拟梁 BEAM 单元。
    逻辑：
        - 仅一根虚拟梁；
        - 按 x 顺序相邻节点连接；
        - 材料号 2 (Q235)，β角 0；
        - 截面名称："虚拟梁"。
    """
    section_name = "虚拟梁"
    section_id = name_to_id.get(section_name, 1)
    mat_id = 2
    beta_angle = 0
    elem_id = start_elem_id
    virtual_elements = []

    # 遍历单根虚拟梁（形式上仍保留 group 结构）
    for gid, node_list in sorted(virtual_single_nodes.items()):
        if not node_list or len(node_list) < 2:
            continue
        node_ids = [nid for nid, _ in node_list]
        for i in range(len(node_ids) - 1):
            n1, n2 = node_ids[i], node_ids[i + 1]
            virtual_elements.append(
                f" {elem_id}, BEAM, {mat_id}, {section_id}, {n1}, {n2}, {beta_angle}, 0"
            )
            elem_id += 1

    next_elem_id = elem_id
    return virtual_elements, next_elem_id

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
    brake_pile_group_nodes: dict,
    connect_groups: dict,
    brake_connect_groups: dict,
    brake_f1_nodes: dict,
    brake_f2_nodes: dict,
    rib_elems,
    top_elems,
    bottom_elems,
    vertical_elements,
    diagonal_elements,
    all_distribute_elements,
    pile_elements,
    connect_elements,
    virtual_elements
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
    brake_pile_nodes_flat  = _flatten_group_nodes(brake_pile_group_nodes)
    connect_nodes_flat     = _flatten_group_nodes(connect_groups)
    brake_connect_nodes_flat = _flatten_group_nodes(brake_connect_groups)
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
        "分配梁": {**(distribute_nodes or {}), **(brake_f1_nodes or {}), **(brake_f2_nodes or {})},
        "钢管桩": {**(pile_nodes_flat or {}), **(brake_pile_nodes_flat or {})},
        "联结系": {**(connect_nodes_flat or {}), **(brake_connect_nodes_flat or {})},
        "横向联系梁": {**(xl_nodes or {}), **(all_nodes or {})},
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
        "分配梁": _extract_elem_ids(all_distribute_elements),
        "钢管桩": _extract_elem_ids(pile_elements),
        "联结系": _extract_elem_ids(connect_elements),
        "虚拟梁": _extract_elem_ids(virtual_elements),
        "横向联系梁": (
            _extract_elem_ids(rib_elems)
            + _extract_elem_ids(top_elems)
            + _extract_elem_ids(bottom_elems)
            + _extract_elem_ids(vertical_elements)
            + _extract_elem_ids(diagonal_elements)
        ),
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
    # 2) 下弦-普通分配梁
    bailey_bottom_elink_nodes: dict,    # {nid: (x,y,z)}
    distribute_bottom_elink_nodes: dict,# {nid: (x,y,z)}
    # 3) 制动墩相关（可选）
    bailey_bottom_brake_elink_nodes: dict, # {nid: (x,y,z)}
    brake_f1_bottom_elink_nodes: dict,     # {nid: (x,y,z)}
    brake_f1_f2_elink_nodes: dict,        # {nid: (x,y,z)}
    brake_f2_f1_elink_nodes: dict,        # {nid: (x,y,z)}
    brake_f2_pile_elink_nodes: dict,      # {nid: (x,y,z)}
    # 4) 分配梁-普通桩
    distribute_pile_elink_nodes: dict,    # {nid: (x,y,z)}
    pile_top_nodes: dict,                 # {nid: (x,y,z)}
    # 5) 制动墩 F1-制动墩桩
    brake_pile_top_nodes: dict,           # {nid: (x,y,z)}
    # 6. 虚拟梁与小肋
    xl_virtual_elink_nodes: dict,
    virtual_nodes: dict,
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
        
        left_support  = components_tab.left_select_var.get()
        right_support = components_tab.right_select_var.get()
        has_brake = ("制动墩" in {left_support, right_support})

        # 1. 小肋 ↔ 贝雷梁上弦
        xl_nodes_flat = _flatten_dict_nodes(xl_bailey_elink_nodes)
        lines, lid = _pair_nodes(xl_nodes_flat, bailey_top_elink_nodes, lid, "小肋-上弦弹连")
        elink_lines.extend(lines)

        # 2. 下弦 ↔ 普通分配梁
        lines, lid = _pair_nodes(bailey_bottom_elink_nodes, distribute_bottom_elink_nodes, lid, "下弦-分配梁弹连")
        elink_lines.extend(lines)

        # 3. 若有制动墩
        if has_brake:
            # 下弦 ↔ F1
            lines, lid = _pair_nodes(bailey_bottom_brake_elink_nodes, brake_f1_bottom_elink_nodes, lid, "下弦-F1弹连")
            elink_lines.extend(lines)

            # F1 ↔ F2
            lines, lid = _pair_nodes(brake_f1_f2_elink_nodes, brake_f2_f1_elink_nodes, lid, "F1-F2弹连")
            elink_lines.extend(lines)

            # F2 ↔ 制动墩桩
            lines, lid = _pair_nodes(brake_f2_pile_elink_nodes, brake_pile_top_nodes, lid, "F2-制动墩桩弹连")
            elink_lines.extend(lines)

        # 4. 普通分配梁 ↔ 普通钢管桩
        lines, lid = _pair_nodes(distribute_pile_elink_nodes, pile_top_nodes, lid, "分配梁-普通桩弹连")
        elink_lines.extend(lines)

        # 5. 虚拟梁 ↔ 小肋
        lines, lid = _pair_nodes(xl_virtual_elink_nodes, virtual_nodes, lid, "虚拟梁-小肋弹连")
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
def build_constraint_lines(pile_bottom_nodes: dict, brake_pile_bottom_nodes: dict = None,bailey_bottom_abutment_nodes: dict = None):
    """
    生成桩底固结 (*CONSTRAINT) 段。
    参数:
        pile_bottom_nodes: dict → 普通桩底节点 {nid:(x,y,z), ...}
        brake_pile_bottom_nodes: dict → 制动墩桩底节点（可选）
    返回:
        constraint_lines: list[str] → 可直接写入 MCT 的行
    """
    constraint_lines = []

    # 合并两个节点集合
    all_nodes = []
    if pile_bottom_nodes:
        all_nodes += sorted(pile_bottom_nodes.keys())
    if brake_pile_bottom_nodes:
        all_nodes += sorted(brake_pile_bottom_nodes.keys())
    if not all_nodes:
        print("[提示] 未检测到桩底节点，跳过 *CONSTRAINT 段。")
        return constraint_lines

    node_groups_str = num_lst_to_num_group(all_nodes)
    node_groups_str = node_groups_str.replace(",", " ")
    constraint_lines.append(f"{node_groups_str}, 111111, 桩底反力")
    if bailey_bottom_abutment_nodes:
        abutment_nodes = []
        abutment_nodes += sorted(bailey_bottom_abutment_nodes.keys())
        constraint_lines.append(f"{num_lst_to_num_group(abutment_nodes)}, 111000,")
    return constraint_lines

# ==========================================================
#   荷载组及荷载说明
# ==========================================================
def build_move_load_and_stldcase(load_tab):
    """
    从荷载界面 load_tab 中提取用户定义的工况名，
    生成需追加到 *LOAD-GROUP 与 *STLDCASE 的部分。

    参数：
        load_tab: 荷载界面实例（含 case_list）
    返回：
        move_load_lst: 用户定义工况名称列表
        move_STLDCASE_lst: 对应的 '名称, USER,' 列表
    """
    # === 用户输入的移动荷载工况 ===
    move_load_lst = []
    if hasattr(load_tab, "case_list"):
        for case in load_tab.case_list:
            name_entry = case.get("name")
            if name_entry:
                name_val = name_entry.get().strip()
                if name_val:
                    move_load_lst.append(name_val)

    # === 生成 *STLDCASE 的追加部分 ===
    move_STLDCASE_lst = [f"{name}, USER," for name in move_load_lst]

    return move_load_lst, move_STLDCASE_lst

# ==========================================================
#  风荷载
# ==========================================================
def build_wind_load(load_tab, components_tab,pile_top_nodes: dict, brake_pile_top_nodes: dict = None,):
    """
    生成风荷载（节点荷载）MCT命令行。
    
    参数：
        load_tab: LoadTab 实例（包含风速输入框）
        components_tab: ComponentsTab 实例（用于读取贝雷梁片数）
        pile_top_nodes: dict {nid: (x, y, z)} 普通桩顶节点
        brake_pile_top_nodes: dict {nid: (x, y, z)} 制动墩桩顶节点（可选）

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

    # 水平加载长度/贝雷片数
    bailey_span_expr = components_tab.entry_bailey_span.get().strip()
    bailey_span_steps_m = midasdisttolst(bailey_span_expr)
    bailey_span_steps = [d * 1000.0 for d in bailey_span_steps_m] # 单位mm
    bailey_count = len(bailey_span_steps)
    L = sum(bailey_span_steps) 
    Feng_values = [u10_gz, dbfl, kt, Z, L, bailey_count, CH, formula]

    # 计算风压
    Gv1, Gv2, kf, kh, kc, Us10, Ud, a0, Ug1, Ug2, Pa1, Pa2, Pa3, A, FF = cal_FengYa_JTG_T_3360_01_2018(u10_gz, kt, Z, L, bailey_count, CH, dbfl, formula)
    F_gz = Pa1 * Z * 1000
    Feng_result_lst = [ Gv1, Gv2, kf, kh, kc, Us10, Ud, a0, Ug1, Ug2, Pa1, Pa2, Pa3, A, FF]

    # === 4. 整理桩顶节点 ===
    node_ids = list(pile_top_nodes.keys())
    if brake_pile_top_nodes:
        node_ids.extend(list(brake_pile_top_nodes.keys()))
    if len(node_ids) == 0:
        print("[信息] 未检测到钢管桩，风荷载不加载。")
        return [], [], []
    
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
                           water_force_points_by_y: dict, 
                           water_force_points_brake_by_y: dict = None):
    """
    生成水流力（节点荷载）MCT命令行。

    参数：
        load_tab: LoadTab 实例（包含水流参数输入框）
        section_data: dict，截面数据（需包含钢管桩直径信息）
        water_force_points_by_y: dict[y] = [(node_id, (x,y,z)), ...]，普通桩水流力加载点
        water_force_points_brake_by_y: dict[y] = [(node_id, (x,y,z)), ...]，制动墩桩水流力加载点（可选）

    返回：
        WATER_FORCE_LINE: list[str]  # MCT文本行
    """
    Shui_values = []
    Shui_lst =[]

    # === 1. 读取 UI 输入参数 ===
    if (not water_force_points_by_y) and (not water_force_points_brake_by_y):
        print("[信息] 未检测到钢管桩，水流力不加载。")
        return [], [], []
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

    # === 7. 如果有制动墩钢管桩 ===
    if water_force_points_brake_by_y:
        d_b = float(section_data.get("制动墩钢管桩", {}).get("params", {}).get("D(mm)", 630.0)) 
        # 取制动墩 y 排序（从大到小 → 上游 → 下游）
        sorted_y_b = sorted(water_force_points_brake_by_y.keys(), reverse=True)
        
        y_front_b = sorted_y_b[0]   # 上游第一排
        Fw_front_b = calc_force(Cw_front)   # 制动墩第一排水流力

        for idx, y in enumerate(sorted_y_b):
            if idx < len(Cw_list):
                Cw_b = Cw_list[idx]
            else:
                Cw_b = Cw_list[-1]

            Fw_b = calc_force(Cw_b)
            # 写入本排所有节点
            for nid, coord in water_force_points_brake_by_y[y]:
                line = f"{nid}, 0 , -{Fw_b:.2f}, 0, 0, 0, 0, , "
                WATER_FORCE_LINE.append(line)

    # === 8. 整理输出 ===
    if water_force_points_brake_by_y:
        Fw_front_b_value = Fw_front_b
        d_b_value = d_b
    else:
        Fw_front_b_value = None
        d_b_value = None
    Shui_lst = [d, Fw_front,d_b_value,Fw_front_b_value]
    return WATER_FORCE_LINE, Shui_values, Shui_lst

# ==========================================================
# 移动荷载
# ==========================================================

# 车道
def build_line_lanes(load_tab, virtual_beam_elems):
    """
    按需生成 *LINELANE(CH) 段
    若用户未定义任何车道，则返回空列表 []
    """

    # 若没有 lane_list，直接返回空
    if not getattr(load_tab, "lane_list", []):
        return []

    lane_content_lines = []     # 保存真实内容（不含标题）

    # === 方向映射 ===
    moving_map = {"往返": "BOTH", "向前": "FORWARD", "向后": "BACKWARD"}

    # === 遍历用户定义的车道 ===
    for lane in load_tab.lane_list:

        # 没有 name_entry → 跳过
        name_entry = lane.get("name_entry")
        if not name_entry:
            continue

        try:
            name = name_entry.get().strip()
            direc = lane.get("direc").get() if lane.get("direc") else "往返"
            dist = float(lane.get("dist").get() or 0) * 1000     # ECC1 (mm)
            ws = float(lane.get("space").get() or 3) * 1000      # WS (mm)

            moving = moving_map.get(direc, "BOTH")

            # === 车道定义 header ===
            lane_header = (
                f"   NAME={name}, CROSS, 横向联系梁, 0, 0, "
                f"{moving}, {ws:.0f}, 3000, NO, 3000"
            )
            lane_content_lines.append(lane_header)

            # === 元素行 ===
            elem_lines = []
            for i, elem_line in enumerate(virtual_beam_elems):
                elem_id = int(elem_line.split(",")[0].strip())
                bstart = "YES" if i == 0 else "NO"
                elem_lines.append(f"{elem_id}, {dist:.0f}, 1500, {bstart}, 1")

            # === 每6项换一行 ===
            for i in range(0, len(elem_lines), 6):
                lane_content_lines.append(
                    "        " + ", ".join(elem_lines[i:i+6])
                )

        except Exception as e:
            print(f"[警告] 生成车道段出错: {e}")
            continue

    # 若没有任何有效车道 → 返回空（不生成 *LINELANE）
    if not lane_content_lines:
        return []

    # 返回完整段（包含标题）
    return ["*LINELANE(CH)"] + lane_content_lines

# 车辆
def build_vehicle_lines(load_tab):
    """
    生成 *VEHICLE 段
    参数:
        load_tab: LoadTab 实例，包含 vehicle_list

    返回:
        list[str] - MCT格式的 *VEHICLE 段行
    """

    vehicle_lines = []

    # 若没有 vehicle_list → 不生成
    if not getattr(load_tab, "vehicle_list", []):
        return []

    temp_lines = []   # 临时保存真实车辆内容行

    for v in load_tab.vehicle_list:
        vtype = v.get("type").get() if v.get("type") else ""
        content = v.get("content")

        # ------------ 履带车 ------------
        if vtype == "旧公路履带车荷载类型":
            try:
                name_entry = getattr(content, "name_entry", None)
                dw1 = float(content.dw1_entry.get())
                dd1 = float(content.dd1_entry.get()) * 1000

                name = name_entry.get().strip() if name_entry else "履带吊"
                temp_lines.append(
                    f"   NAME={name}, 2, TRUCK, 4, {dw1:.3f}, {dd1:.3f}"
                )
            except:
                pass

        # ------------ 公路车道荷载 ------------
        elif vtype == "城市桥梁车道荷载类型":
            try:
                name = content.name_entry.get().strip()
                std = content.std_type.get()
                if std == "CH-CL":
                    temp_lines.append(f"   NAME={name}, 1, CH-CL, JTGB01-2014")
                else:
                    temp_lines.append(f"   NAME={name}, 1, CL-CD, JTGB01-2014")
            except:
                pass

        # ------------ 用户车辆 ------------
        else:
            try:
                name_entry = getattr(content, "name_entry", None)
                name = name_entry.get().strip() if name_entry else "用户车辆"
                
                # 收集荷载和间距数据
                load_dist_pairs = []
                
                row_list = getattr(content, "row_list", [])
                for row_data in row_list:
                    try:
                        # 获取荷载(kN)并转换为N，间距(mm)保持不变
                        load_val = float(row_data["load"].get()) * 1000  # kN转换为N
                        dist_val = float(row_data["dist"].get())
                        load_dist_pairs.extend([int(load_val), int(dist_val)])
                    except (ValueError, KeyError):
                        continue
                if not load_dist_pairs:
                    continue
                # 构建荷载数据字符串
                load_data_str = ", ".join(str(x) for x in load_dist_pairs)
                
                # 生成命令格式：NAME=名称, 2, TRUCK, 1, 0, 0, 0
                # 荷载数据
                temp_lines.append(f"   NAME={name}, 2, TRUCK, 1, 0, 0, 0")
                temp_lines.append(f"        {load_data_str}")
                
            except Exception as e:
                print(f"处理用户车辆数据时出错: {e}")
                pass

    # 若无内容 → 返回空列表
    if not temp_lines:
        return []

    # 否则补上标题
    return ["*VEHICLE"] + temp_lines

# 移动荷载工况
def build_mvldcase_lines(load_tab):
    """
    生成移动荷载工况组合 *MVLDCASE(CH) 段
    若无有效工况则返回空列表 []
    """

    case_list = getattr(load_tab, "case_list", [])
    if not case_list:
        return []   # 没有工况 → 不输出任何段落

    mvldcase_content = []   # 暂存工况内容（不含标题）

    for case in case_list:
        try:
            # 工况名与车辆名
            case_name = case["name"].get().strip()
            vehicle_name = case["vehicle"].get().strip()

            # 若工况名或车辆为空 → 跳过此工况
            if not case_name or not vehicle_name:
                continue

            # === 工况头 ===
            mvldcase_content.append(f"   NAME={case_name}, , NO, 0, 1, 0")

            # === 固定系数模板 ===
            mvldcase_content.extend([
                "        1, 1, 0.8, 0.67, 0.6, 0.55, 0.55, 0.55",
                "        1, 1, 0.78, 0.67, 0.6, 0.55, 0.52, 0.5",
                "        1.2, 1, 0.78, 0.67, 0.6, 0.55, 0.52, 0.5"
            ])

            # === 分配的车道 ===
            assign_vars = case.get("assign", [])
            for lane_name, var in assign_vars:
                if var.get():   # 仅勾选的车道
                    mvldcase_content.append(
                        f"        VL, {vehicle_name}, 1, 1, 1, {lane_name}"
                    )

        except Exception as e:
            print(f"[警告] 生成移动荷载工况出错: {e}")
            continue

    # 若没有任何有效内容 → 不输出段落
    if not mvldcase_content:
        return []

    # 返回完整段（标题 + 内容）
    return ["*MVLDCASE(CH)"] + mvldcase_content

# 荷载组合
def build_loadcomb_lines(load_tab):
    """
    生成 *LOADCOMB 段
    仅当用户输入移动荷载工况时，才生成车辆组合并参与标准/基本组合
    """

    loadcomb_content = []   # 暂存组合内容，不含标题

    # 读取所有移动荷载工况名
    move_case_names = []

    if hasattr(load_tab, "case_list"):
        for case in load_tab.case_list:
            name = case["name"].get().strip()
            if name:
                move_case_names.append(name)

    has_vehicle_load = bool(move_case_names)   # 是否有移动荷载工况

    # 若有车辆工况
    if has_vehicle_load:
        loadcomb_content.append("   NAME=车辆荷载, GEN, ACTIVE, 0, 1, , 0, 0, 0, 1")

        mv_parts = [f"MV, {n}, 1" for n in move_case_names]
        mv_line = "        " + ", ".join(mv_parts)
        loadcomb_content.append(mv_line)

    # 标准组合
    loadcomb_content.append("   NAME=标准组合, GEN, ACTIVE, 0, 0, , 0, 0, 0, 1")

    if has_vehicle_load:
        # 有车辆工况 → 组合车辆荷载
        loadcomb_content.append(
            "        ST, 自重, 1, ST, 风荷载, 0.7, ST, 水流力, 0.7, CB, 车辆荷载, 1"
        )
    else:
        # 无车辆工况 → 不组合车辆荷载
        loadcomb_content.append(
            "        ST, 自重, 1, ST, 风荷载, 0.7, ST, 水流力, 0.7"
        )

    # 基本组合
    loadcomb_content.append("   NAME=基本组合, GEN, ACTIVE, 0, 0, , 0, 0, 0, 1")

    if has_vehicle_load:
        loadcomb_content.append(
            "        ST, 自重, 1.2, ST, 风荷载, 0.98, ST, 水流力, 0.98, CB, 车辆荷载, 1.4"
        )
    else:
        loadcomb_content.append(
            "        ST, 自重, 1.2, ST, 风荷载, 0.98, ST, 水流力, 0.98"
        )


    # 返回完整段（带标题）
    return ["*LOADCOMB"] + loadcomb_content