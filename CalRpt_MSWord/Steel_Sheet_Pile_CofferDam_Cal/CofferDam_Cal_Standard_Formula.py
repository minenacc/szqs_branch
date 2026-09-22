# 1. 标准库
from typing import List, Tuple
from dataclasses import dataclass

# 2. 第三方库
import numpy as np
from scipy.optimize import differential_evolution


#=====================================================================================================================================================================
# 钢结构设计标准GB50017_2017 公式 8.2.1-1/2
GB50017_2017_8_2_1_1 = r'\sigma=\frac{N}{\varphi A}+\frac{\beta M}{\gamma_mW(1-0.8\frac{N}{N_{Ex}^\prime})}' # NEX
GB50017_2017_8_2_1_2 = r'N_{EX}^\prime=\frac{\pi^2EA}{1.1\ \lambda^2}' 
# 港口工程荷载规范
JTS_144_1_2010_13_0_1 = r'F_\omega=C_\omega\frac{\rho}{2}V^2A'
# 建筑基坑支护技术规程JGJ 120-2012 4.2.4 抗隆起稳定性验算
JGJ_120_2012_4_2_4 = r'\frac{\gamma_{m2}l_dN_q+cN_c}{\gamma_{m1}\left(h+l_d\right)+q_0}'
# 建筑基坑支护技术规程JGJ 120-2012 4.2.1 悬臂式支挡结构嵌固稳定性验算
JGJ_120_2012_4_2_1 = r'\frac{E_{\mathrm{pk}}\cdot a_{\mathrm{p1}}}{E_{\mathrm{ak}}\cdot a_{\mathrm{a1}}}'
# 建筑基坑支护技术规程JGJ 120-2012 4.2.2 单层锚杆和单层支撑的支挡式结构嵌固稳定性验算
JGJ_120_2012_4_2_2 = r'\frac{E_{\mathrm{pk}}\cdot a_{\mathrm{p2}}}{E_{\mathrm{ak}}\cdot a_{\mathrm{a2}}}'
# 建筑基坑支护技术规范JGJ 120-2012 C.0.1 抗突涌稳定性验算
JGJ_120_2012_C_0_1 = r'\frac{D\gamma}{h_w\gamma_w}'
# 建筑基坑支护技术规范JGJ 120-2012 C.0.2 流土稳定性验算
JGJ_120_2012_C_0_2 = r'\frac{\left(2l_d+0.8D_1\right)\gamma^\prime}{\Delta h\gamma_w}'


# ============================================================
# 圆弧滑动计算函数
# ============================================================
def elev_to_y(elev, pit_bottom_elev):
    return elev - pit_bottom_elev


def get_soil_params(y, layers):
    for lyr in layers:
        if lyr['bot_y'] <= y <= lyr['top_y']:
            return lyr['c'], lyr['phi'], lyr['gamma']
    if y < layers[-1]['bot_y']:
        lyr = layers[-1]
        return lyr['c'], lyr['phi'], lyr['gamma']
    return layers[0]['c'], layers[0]['phi'], layers[0]['gamma']


def circle_x_at_y(cx, cy, R, y_line):
    disc = R**2 - (y_line - cy)**2
    if disc < 0:
        return None
    sd = np.sqrt(disc)
    return (cx - sd, cx + sd)


def get_pore_pressure(y_bot, water_y, gamma_w):
    if y_bot < water_y:
        return gamma_w * (water_y - y_bot)
    return 0.0


def calc_slice(x_start, b, cx, cy, R, surcharge, ground_y, layers, water_y, gamma_w):
    x_mid = x_start + b / 2.0
    disc = R**2 - (x_mid - cx)**2
    if disc < 0:
        return None
    y_bottom = cy - np.sqrt(disc)
    y_top = ground_y if x_mid > 0 else 0.0
    if y_bottom >= y_top - 1e-9:
        return None
    sin_alpha = (x_mid - cx) / R
    cos_alpha = np.sqrt(max(1.0 - sin_alpha**2, 0.0))
    alpha_deg = np.degrees(np.arcsin(np.clip(sin_alpha, -1.0, 1.0)))
    l_i = b / cos_alpha if cos_alpha > 1e-6 else b
    W = 0.0
    for lyr in layers:
        y_lyr_top = min(lyr['top_y'], y_top)
        y_lyr_bot = max(lyr['bot_y'], y_bottom)
        if y_lyr_top <= y_lyr_bot:
            continue
        dh = y_lyr_top - y_lyr_bot
        W += b * dh * lyr['gamma']
    last_lyr = layers[-1]
    if y_bottom < last_lyr['bot_y']:
        dh_extra = last_lyr['bot_y'] - y_bottom
        W += b * dh_extra * last_lyr['gamma']
    W_q     = surcharge * b if x_mid > 0 else 0.0
    W_total = W + W_q
    u_i     = get_pore_pressure(y_bottom, water_y, gamma_w)
    c_i, phi_i, _ = get_soil_params(y_bottom, layers)
    phi_rad = np.radians(phi_i)
    N_eff   = W_total * cos_alpha - u_i * l_i
    T_resist = c_i * l_i + max(N_eff, 0.0) * np.tan(phi_rad)
    W_sin   = W_total * sin_alpha
    side    = '外侧' if x_mid > 0 else '内侧'
    ws_flag = '↑促滑' if W_sin > 0 else '↓抗滑'
    return {
        'x_start': x_start, 'x_mid': x_mid, 'b': b, 'side': side,
        'y_bottom': y_bottom, 'y_top': y_top,
        'alpha_deg': alpha_deg, 'sin_alpha': sin_alpha, 'cos_alpha': cos_alpha,
        'l_i': l_i, 'W_soil': W, 'W_q': W_q, 'W_total': W_total,
        'u_i': u_i, 'N_eff': N_eff, 'c_i': c_i, 'phi_i': phi_i,
        'T_resist': T_resist, 'W_sin': W_sin, 'ws_flag': ws_flag,
    }


def cal_Ksi(cx, cy, R,
           level, slice_width,
           ground_elev, water_elev, pit_bottom_elev,
           solid_layer_lst, surcharge, gamma_w=10):
    ground_y = elev_to_y(ground_elev, pit_bottom_elev)
    water_y  = elev_to_y(water_elev,  pit_bottom_elev)
    layers = []
    current_elev = ground_elev
    for i, sublist in enumerate(solid_layer_lst):
        thick, gamma, c, phi = sublist[:4]
        top_elev = current_elev
        bot_elev = current_elev - thick
        layers.append({
            'id': i+1, 'thick': thick, 'gamma': gamma, 'c': c, 'phi': phi,
            'top_elev': top_elev, 'bot_elev': bot_elev,
            'top_y': elev_to_y(top_elev, pit_bottom_elev),
            'bot_y': elev_to_y(bot_elev, pit_bottom_elev),
        })
        current_elev = bot_elev
    pts_pit = circle_x_at_y(cx, cy, R, 0.0)
    if pts_pit is None:
        raise ValueError("圆弧与坑底(y=0)无交点")
    x_left = pts_pit[0]
    pts_ground = circle_x_at_y(cx, cy, R, ground_y)
    if cy > ground_y:
        if pts_ground is None:
            raise ValueError("圆弧与地面无交点")
        x_right = pts_ground[1]
    else:
        x_right = cx + R
    x_arr  = np.arange(x_left, x_right, slice_width)
    slices = []
    for x_s in x_arr:
        b_act = min(slice_width, x_right - x_s)
        if b_act < 1e-6:
            continue
        res = calc_slice(x_s, b_act, cx, cy, R, surcharge,
                         ground_y, layers, water_y, gamma_w)
        if res is not None:
            slices.append(res)
    slices_out = [s for s in slices if s['side'] == '外侧']
    slices_in  = [s for s in slices if s['side'] == '内侧']
    sum_T_all     = sum(s['T_resist'] for s in slices)
    sum_Ws_in_neg = sum(s['W_sin'] for s in slices_in if s['W_sin'] < 0)
    sum_Ws_in_pos = sum(s['W_sin'] for s in slices_in if s['W_sin'] > 0)
    sum_Ws_out    = sum(s['W_sin'] for s in slices_out)
    sum_Ws_in     = sum(s['W_sin'] for s in slices_in)
    sum_Ws_all    = sum_Ws_out + sum_Ws_in
    Ks = sum_T_all / sum_Ws_all if abs(sum_Ws_all) > 1e-9 else float('inf')
    M_R = (sum_T_all + abs(sum_Ws_in_neg)) * R
    M_S = (sum_Ws_out + sum_Ws_in_pos) * R
    Ks_moment = M_R / M_S if abs(M_S) > 1e-9 else float('inf')
    return round(Ks, 3)


def Cal_Ks_detail(cx, cy, R,
                  level, slice_width,
                  ground_elev, water_elev, pit_bottom_elev,
                  solid_layer_lst, surcharge, gamma_w=10):
    """
    与 cal_Ks() 完全相同的计算逻辑，
    额外返回 (Ks, sum_T_all, sum_Ws_all)
    """
    ground_y = elev_to_y(ground_elev, pit_bottom_elev)
    water_y  = elev_to_y(water_elev,  pit_bottom_elev)
    layers = []
    current_elev = ground_elev
    # solid_layer_lst = [[2.922, 17.0, 0.0, 10.0, '水土分算', '0.5', '0.0'], [...]]
    for i, sublist in enumerate(solid_layer_lst):
        thick, gamma, c, phi, method, shentou,shuitou = sublist
        top_elev = current_elev
        bot_elev = current_elev - thick
        layers.append({
            'id': i+1, 'thick': thick, 'gamma': gamma, 'c': c, 'phi': phi,
            'top_elev': top_elev, 'bot_elev': bot_elev,
            'top_y': elev_to_y(top_elev, pit_bottom_elev),
            'bot_y': elev_to_y(bot_elev, pit_bottom_elev),
        })
        current_elev = bot_elev
    pts_pit = circle_x_at_y(cx, cy, R, 0.0)
    if pts_pit is None:
        raise ValueError("圆弧与坑底(y=0)无交点")
    x_left = pts_pit[0]
    pts_ground = circle_x_at_y(cx, cy, R, ground_y)
    if cy > ground_y:
        if pts_ground is None:
            raise ValueError("圆弧与地面无交点")
        x_right = pts_ground[1]
    else:
        x_right = cx + R
    x_arr  = np.arange(x_left, x_right, slice_width)
    slices = []
    for x_s in x_arr:
        b_act = min(slice_width, x_right - x_s)
        if b_act < 1e-6:
            continue
        res = calc_slice(x_s, b_act, cx, cy, R, surcharge,
                         ground_y, layers, water_y, gamma_w)
        if res is not None:
            slices.append(res)
    slices_out = [s for s in slices if s['side'] == '外侧']
    slices_in  = [s for s in slices if s['side'] == '内侧']
    sum_T_all  = sum(s['T_resist'] for s in slices)
    sum_Ws_out = sum(s['W_sin']    for s in slices_out)
    sum_Ws_in  = sum(s['W_sin']    for s in slices_in)
    sum_Ws_all = sum_Ws_out + sum_Ws_in
    Ks = sum_T_all / sum_Ws_all if abs(sum_Ws_all) > 1e-9 else float('inf')
    return round(Ks, 3), round(sum_T_all, 1), round(sum_Ws_all, 1)


# ============================================================
# 工程有效性检验
# ============================================================
def is_valid_circle(cx: float, cy: float, R: float,
                    ground_y: float, pile_bot_y: float) -> bool:
    pts_pit = circle_x_at_y(cx, cy, R, 0.0)
    if pts_pit is None:
        return False
    x_left = pts_pit[0]
    if x_left >= -1e-6:
        return False
    if cy > ground_y:
        pts_ground = circle_x_at_y(cx, cy, R, ground_y)
        if pts_ground is None:
            return False
        x_right = pts_ground[1]
    else:
        x_right = cx + R
    if x_right <= 1e-6:
        return False
    dist_pile = np.hypot(cx, cy - pile_bot_y)
    if dist_pile > R + 1e-6:
        return False
    if (x_right - x_left) < 1.0:
        return False
    return True


def passes_pile_bottom(cx: float, cy: float, R: float,
                       pile_bot_y: float, tol: float = 0.15) -> bool:
    dist = np.hypot(cx, cy - pile_bot_y)
    return abs(dist - R) < tol


# ============================================================
# 数据类
# ============================================================
@dataclass
class CircleResult:
    cx: float
    cy: float
    R:  float
    Ks: float
    pass_pile_bottom: bool
    mode: str


@dataclass
class SearchDomain:
    cx_min: float; cx_max: float
    cy_min: float; cy_max: float
    nx: int = 30
    ny: int = 40


# ============================================================
# 搜索域  ← 补齐 4 个参数
# ============================================================
def build_search_domain(ground_elev: float,
                        pit_bottom_elev: float,
                        strut_bottom_elev: float,
                        Ld: float) -> SearchDomain:
    ground_y   = elev_to_y(ground_elev,       pit_bottom_elev)
    pile_bot_y = elev_to_y(strut_bottom_elev, pit_bottom_elev)

    cx_min = -(Ld * 2.0)
    cx_max = 0.0
    cy_min = 0.0
    # 与 build_bounds 保持一致：cy 不能太高，否则 R 必须极大才能与 y=0 相交
    cy_max = min(ground_y * 8.0, 60.0)

    dom = SearchDomain(cx_min=cx_min, cx_max=cx_max,
                       cy_min=cy_min, cy_max=cy_max)

    print(f"\n{'='*65}")
    print(f"[搜索域]")
    print(f"  ground_y={ground_y:.3f}m  pile_bot_y={pile_bot_y:.3f}m  Ld={Ld:.3f}m")
    print(f"  cx: [{cx_min:.3f}, {cx_max:.3f}]  网格数 nx={dom.nx}")
    print(f"  cy: [{cy_min:.3f}, {cy_max:.3f}]  网格数 ny={dom.ny}")
    print(f"  总圆心数: {dom.nx*dom.ny}")
    print(f"{'='*65}\n")
    return dom


# ============================================================
# 单圆心扫描  ← 补齐全部计算参数
# ============================================================
def scan_one_center_all(cx: float, cy: float,
                        ground_y: float, pile_bot_y: float,
                        # ↓ 以下为新增透传参数
                        level: str,
                        slice_width: float,
                        ground_elev: float,
                        water_elev: float,
                        pit_bottom_elev: float,
                        solid_layer_lst: list,
                        surcharge: float,
                        gamma_w: float,
                        nr_free: int = 12
                        ) -> List[CircleResult]:
    results: List[CircleResult] = []
    dist_pile = np.hypot(cx, cy - pile_bot_y)

    # 模式A：R = dist，圆恰好过桩底
    R_anchor = dist_pile
    if R_anchor > 0.1:
        if is_valid_circle(cx, cy, R_anchor, ground_y, pile_bot_y):
            try:
                Ks = cal_Ksi(cx, cy, R_anchor,
                            level, slice_width,
                            ground_elev, water_elev, pit_bottom_elev,
                            solid_layer_lst, surcharge, gamma_w)
                if 0.1 <= Ks <= 20.0:
                    results.append(CircleResult(
                        cx=cx, cy=cy, R=R_anchor, Ks=Ks,
                        pass_pile_bottom=True, mode='anchored'))
            except (ValueError, ZeroDivisionError, RuntimeError):
                pass

    # 模式B：自由搜索，R > dist
    step_free  = max(0.5, dist_pile * 0.06)
    r_min_free = dist_pile + step_free
    r_max_free = r_min_free + nr_free * step_free

    for R in np.linspace(r_min_free, r_max_free, nr_free + 1):
        if not is_valid_circle(cx, cy, R, ground_y, pile_bot_y):
            continue
        try:
            Ks = cal_Ksi(cx, cy, R,
                        level, slice_width,
                        ground_elev, water_elev, pit_bottom_elev,
                        solid_layer_lst, surcharge, gamma_w)
        except (ValueError, ZeroDivisionError, RuntimeError):
            continue
        if not (0.1 <= Ks <= 20.0):
            continue
        ppb = passes_pile_bottom(cx, cy, R, pile_bot_y)
        results.append(CircleResult(
            cx=cx, cy=cy, R=R, Ks=Ks,
            pass_pile_bottom=ppb, mode='free'))

    return results


# ============================================================
# 主搜索：双阶段  ← 补齐全部参数
# ============================================================
def search_critical_circle(ground_elev: float,
                           pit_bottom_elev: float,
                           strut_bottom_elev: float,
                           water_elev: float,
                           surcharge: float,
                           level: str,
                           slice_width: float,
                           solid_layer_lst: list,
                           Ld: float,
                           gamma_w: float
                           ) -> Tuple[CircleResult, List[CircleResult]]:

    dom        = build_search_domain(ground_elev, pit_bottom_elev,
                                     strut_bottom_elev, Ld)
    ground_y   = elev_to_y(ground_elev,       pit_bottom_elev)
    pile_bot_y = elev_to_y(strut_bottom_elev, pit_bottom_elev)

    cx_arr = np.linspace(dom.cx_min, dom.cx_max, dom.nx)
    cy_arr = np.linspace(dom.cy_min, dom.cy_max, dom.ny)

    total_centers = dom.nx * dom.ny
    print(f"[阶段1] 粗搜索  {dom.nx}×{dom.ny}={total_centers} 个圆心")
    print(f"  每圆心: 1个锚定圆(模式A) + 13个自由圆(模式B)")
    print(f"  理论上限: {total_centers * 14} 次计算\n")

    all_results: List[CircleResult] = []
    n_anchor = 0; n_free_valid = 0; n_center_empty = 0

    for cx in cx_arr:
        for cy in cy_arr:
            circle_list = scan_one_center_all(
                cx, cy, ground_y, pile_bot_y,
                level, slice_width,
                ground_elev, water_elev, pit_bottom_elev,
                solid_layer_lst, surcharge, gamma_w,
                nr_free=12
            )
            if not circle_list:
                n_center_empty += 1
                continue
            all_results.extend(circle_list)
            n_anchor     += sum(1 for r in circle_list if r.mode == 'anchored')
            n_free_valid += sum(1 for r in circle_list if r.mode == 'free')

    if not all_results:
        raise RuntimeError("粗搜索无有效结果，请检查参数")

    all_results.sort(key=lambda r: r.Ks)
    best_c = all_results[0]

    n_pass = sum(1 for r in all_results if r.pass_pile_bottom)
    print(f"  粗搜总圆数: {len(all_results)}")
    print(f"    模式A(锚定,过桩底): {n_anchor}")
    print(f"    模式B(自由搜索):    {n_free_valid}")
    print(f"    其中判定经过桩底:   {n_pass}  "
          f"({100*n_pass/max(len(all_results),1):.1f}%)")
    print(f"    无有效圆的圆心:     {n_center_empty}")
    print(f"  粗搜最优: cx={best_c.cx:.3f} cy={best_c.cy:.3f} "
          f"R={best_c.R:.3f} Ks={best_c.Ks:.3f} 模式={best_c.mode}\n")

    # 阶段2：局部精细化
    dcx = (dom.cx_max - dom.cx_min) / (dom.nx - 1)
    dcy = (dom.cy_max - dom.cy_min) / (dom.ny - 1)

    cx_fine = np.linspace(
        max(best_c.cx - 2*dcx, dom.cx_min),
        min(best_c.cx + 2*dcx, dom.cx_max), 15)
    cy_fine = np.linspace(
        max(best_c.cy - 2*dcy, dom.cy_min),
        min(best_c.cy + 2*dcy, dom.cy_max), 15)
    print(f"[阶段2] 精细搜索  15×15=225 圆心")

    fine_results: List[CircleResult] = []
    for cx in cx_fine:
        for cy in cy_fine:
            circle_list = scan_one_center_all(
                cx, cy, ground_y, pile_bot_y,
                level, slice_width,
                ground_elev, water_elev, pit_bottom_elev,
                solid_layer_lst, surcharge, gamma_w,
                nr_free=20
            )
            fine_results.extend(circle_list)

    if fine_results:
        best_final = min(fine_results, key=lambda r: r.Ks)
        combined   = all_results + fine_results
        combined.sort(key=lambda r: r.Ks)
    else:
        best_final = best_c
        combined   = all_results

    dist_final = np.hypot(best_final.cx, best_final.cy - pile_bot_y)
    print(f"  精细最优: cx={best_final.cx:.4f} cy={best_final.cy:.4f} "
          f"R={best_final.R:.4f} Ks={best_final.Ks:.3f}")
    print(f"  dist(圆心→桩底)={dist_final:.4f}m  R={best_final.R:.4f}m  "
          f"{'✓绕过桩底' if dist_final<=best_final.R+0.01 else '✗未绕过'}")

    return best_final, combined


# ============================================================
# 结果输出  ← 补齐 3 个参数
# ============================================================
def print_report(best: CircleResult,
                 all_results: List[CircleResult],
                 level: str,
                 strut_bottom_elev: float,
                 pit_bottom_elev: float):
    ks_require = {'一级': 1.35, '二级': 1.30, '三级': 1.25}
    Ks_req     = ks_require.get(level, 1.30)
    pile_bot_y = elev_to_y(strut_bottom_elev, pit_bottom_elev)

    Ks_arr = np.array([r.Ks for r in all_results])
    n_pass = sum(1 for r in all_results if r.pass_pile_bottom)
    n_anch = sum(1 for r in all_results if r.mode == 'anchored')
    n_free = sum(1 for r in all_results if r.mode == 'free')
    dist_p = np.hypot(best.cx, best.cy - pile_bot_y)

    print(f"\n{'='*65}")
    print(f"  整体圆弧滑动稳定性（JGJ120-2012 §4.2.3）")
    print(f"{'='*65}")
    print(f"  最危险圆弧：")
    print(f"    圆心 cx = {best.cx:>10.4f} m")
    print(f"    圆心 cy = {best.cy:>10.4f} m")
    print(f"    半径  R = {best.R:>10.4f} m")
    print(f"    搜索模式: {best.mode}")
    print(f"{'─'*65}")
    print(f"  工程约束验证：")
    print(f"    桩底(0, {pile_bot_y:.3f}m) → 圆心距离 = {dist_p:.4f}m")
    print(f"    R={best.R:.4f}m {'≥' if best.R>=dist_p-0.01 else '<'} "
          f"dist={dist_p:.4f}m "
          f"→ {'✓圆弧绕过桩底' if best.R>=dist_p-0.01 else '✗未绕过桩底'}")
    print(f"    经过桩底附近: {'是' if best.pass_pile_bottom else '否'}")
    print(f"{'─'*65}")
    print(f"    最小安全系数 Ks  = {best.Ks:.3f}")
    print(f"    规范要求  [Ks]  = {Ks_req}  ({level}基坑)")
    print(f"    判定: {'✓ 满足' if best.Ks >= Ks_req else '✗ 不满足'}")
    print(f"{'─'*65}")
    print(f"  搜索统计：")
    print(f"    有效圆总数:        {len(all_results)}")
    print(f"    模式A(锚定过桩底): {n_anch}  ({100*n_anch/max(len(all_results),1):.1f}%)")
    print(f"    模式B(自由搜索):   {n_free}  ({100*n_free/max(len(all_results),1):.1f}%)")
    print(f"    判定经过桩底:      {n_pass}  ({100*n_pass/max(len(all_results),1):.1f}%)")
    print(f"    Ks 范围:           [{Ks_arr.min():.3f}, {Ks_arr.max():.3f}]")
    print(f"{'='*65}")


# ============================================================
# 目标函数  ← 改为接收显式参数，用闭包工厂包装
# ============================================================
def make_objective(ground_elev: float,
                   pit_bottom_elev: float,
                   strut_bottom_elev: float,
                   water_elev: float,
                   surcharge: float,
                   level: str,
                   slice_width: float,
                   solid_layer_lst: list,
                   gamma_w: float):
    """
    返回一个仅接受 params=[cx,cy,R] 的目标函数（供 differential_evolution 使用）。
    所有问题参数通过闭包捕获，不依赖全局变量。
    """
    ground_y   = elev_to_y(ground_elev,       pit_bottom_elev)
    pile_bot_y = elev_to_y(strut_bottom_elev, pit_bottom_elev)

    def objective(params):
        cx, cy, R = params

        # ── 连续惩罚：对不满足几何约束的解给出梯度信息 ──
        penalty = 0.0

        # 圆弧与 y=0 相交要求 R >= cy
        if R < cy:
            penalty += (cy - R) * 50.0

        # 圆弧需绕过桩底: dist(圆心, 桩底) <= R
        dist_pile = np.hypot(cx, cy - pile_bot_y)
        if dist_pile > R:
            penalty += (dist_pile - R) * 50.0

        # 圆弧与 y=0 交点需在 x<0 侧（坑外）
        disc_pit = R**2 - cy**2
        if disc_pit > 0:
            x_left = cx - np.sqrt(disc_pit)
            if x_left >= 0:
                penalty += abs(x_left) * 30.0
        else:
            penalty += abs(disc_pit) * 10.0

        # 有明显惩罚则直接返回，不进入条分计算
        if penalty > 1.0:
            return 999.0 + penalty

        try:
            Ks = cal_Ksi(cx, cy, R,
                        level, slice_width,
                        ground_elev, water_elev, pit_bottom_elev,
                        solid_layer_lst, surcharge, gamma_w)
            if not (0.1 <= Ks <= 20.0):
                return 999.0
            return Ks
        except Exception:
            return 999.0

    return objective


# ============================================================
# 搜索边界  ← 补齐 4 个参数
# ============================================================
def build_bounds(pit_bottom_elev: float,
                 strut_bottom_elev: float,
                 ground_elev: float,
                 Ld: float):
    pile_bot_y = elev_to_y(strut_bottom_elev, pit_bottom_elev)
    ground_y   = elev_to_y(ground_elev, pit_bottom_elev)

    cx_min = -(Ld * 2.0)
    cx_max =  0.0
    cy_min =  0.0
    # cy 不能太高，否则 R 必须极大才能与 y=0 相交（约束 R >= cy）
    # 限制 cy 在合理范围内，使搜索空间中可行域占比足够大
    cy_max =  min(ground_y * 8.0, 60.0)
    # R 下界放开，由目标函数中的连续惩罚引导 R >= cy 和 R >= dist_pile
    R_min  =  1.0
    R_max  =  np.hypot(cx_min, cy_max - pile_bot_y) * 1.1

    bounds = [(cx_min, cx_max), (cy_min, cy_max), (R_min, R_max)]

    print(f"\n{'='*65}")
    print(f"[差分进化] 搜索边界")
    print(f"  cx : [{cx_min:.3f}, {cx_max:.3f}] m")
    print(f"  cy : [{cy_min:.3f}, {cy_max:.3f}] m")
    print(f"  R  : [{R_min:.3f}, {R_max:.3f}] m")
    print(f"  pile_bot_y = {pile_bot_y:.3f} m  ground_y = {ground_y:.3f} m")
    print(f"  R >= cy 由目标函数连续惩罚保证")
    print(f"{'='*65}\n")
    return bounds


# ============================================================
# 回调  ← objective 通过构造时注入，不依赖全局
# ============================================================
class DECallback:
    def __init__(self, objective_fn):
        self.objective   = objective_fn   # ← 注入目标函数
        self.iter        = 0
        self.best_Ks     = float('inf')
        self.best_params = None

    def __call__(self, xk, convergence):
        self.iter += 1
        Ks = self.objective(xk)
        if Ks < self.best_Ks:
            self.best_Ks     = Ks
            self.best_params = xk.copy()
        if self.iter % 5 == 0 or Ks < self.best_Ks + 0.01:
            print(f"  第{self.iter:>3d}代 | "
                  f"cx={xk[0]:>8.3f} cy={xk[1]:>8.3f} R={xk[2]:>8.3f} | "
                  f"Ks={Ks:.4f} | 收敛度={convergence:.4f}")
        return False


# ============================================================
# 差分进化主搜索  ← 补齐全部参数
# ============================================================
def search_by_de(ground_elev: float,
                 pit_bottom_elev: float,
                 strut_bottom_elev: float,
                 water_elev: float,
                 surcharge: float,
                 level: str,
                 slice_width: float,
                 solid_layer_lst: list,
                 Ld: float,
                 gamma_w: float) -> CircleResult:

    bounds    = build_bounds(pit_bottom_elev, strut_bottom_elev,
                             ground_elev, Ld)
    objective = make_objective(ground_elev, pit_bottom_elev,
                               strut_bottom_elev, water_elev,
                               surcharge, level, slice_width,
                               solid_layer_lst, gamma_w)
    callback  = DECallback(objective)

    print(f"[差分进化] 开始搜索...")
    print(f"  策略: best1bin | 种群大小: popsize=18 | 最大迭代: maxiter=200")
    print(f"  预计最大计算次数: 18×3×200 = {18*3*200} 次\n")

    result = differential_evolution(
        func          = objective,
        bounds        = bounds,
        strategy      = 'best1bin',
        maxiter       = 200,
        popsize       = 18,
        tol           = 1e-5,
        mutation      = (0.5, 1.0),
        recombination = 0.9,
        seed          = 42,
        callback      = callback,
        polish        = True,
        init          = 'latinhypercube',
        workers       = 1,
        disp          = False,
    )

    cx, cy, R  = result.x
    Ks         = result.fun
    pile_bot_y = elev_to_y(strut_bottom_elev, pit_bottom_elev)
    ground_y   = elev_to_y(ground_elev, pit_bottom_elev)
    dist_pile  = np.hypot(cx, cy - pile_bot_y)
    ppb        = passes_pile_bottom(cx, cy, R, pile_bot_y)

    print(f"\n[差分进化] 搜索完成")
    print(f"  收敛状态: {result.message}")
    print(f"  实际迭代: {result.nit} 代")
    print(f"  实际计算: {result.nfev} 次")
    print(f"  polish后: cx={cx:.4f} cy={cy:.4f} R={R:.4f} Ks={Ks:.4f}")
    print(f"  dist(圆心→桩底)={dist_pile:.4f}m R={R:.4f}m "
          f"{'✓绕过桩底' if R >= dist_pile - 0.01 else '✗未绕过'}")

    # ── 检查 DE 结果是否有效 ──
    if not is_valid_circle(cx, cy, R, ground_y, pile_bot_y) or Ks >= 999.0:
        print(f"\n  ⚠ DE 未找到有效圆弧，从最终种群中筛选...")
        # 遍历 DE 最终种群，找所有有效圆
        population = result.population
        pop_ks     = result.population_energies
        best_from_pop = None
        for i in np.argsort(pop_ks):
            pcx, pcy, pR = population[i]
            pKs = pop_ks[i]
            if pKs >= 999.0:
                continue
            if is_valid_circle(pcx, pcy, pR, ground_y, pile_bot_y):
                best_from_pop = (pcx, pcy, pR, pKs)
                break
        if best_from_pop is not None:
            cx, cy, R, Ks = best_from_pop
            dist_pile = np.hypot(cx, cy - pile_bot_y)
            ppb = passes_pile_bottom(cx, cy, R, pile_bot_y)
            print(f"  ✓ 从种群中找到有效圆: cx={cx:.4f} cy={cy:.4f} R={R:.4f} Ks={Ks:.4f}")
        else:
            # 种群中也没有 → 回退到网格搜索
            print(f"  ⚠ 种群中无有效圆，回退到网格搜索...")
            grid_best, _ = search_critical_circle(
                ground_elev, pit_bottom_elev, strut_bottom_elev,
                water_elev, surcharge, level, slice_width,
                solid_layer_lst, Ld, gamma_w)
            cx = grid_best.cx; cy = grid_best.cy; R = grid_best.R
            Ks = grid_best.Ks; ppb = grid_best.pass_pile_bottom
            dist_pile = np.hypot(cx, cy - pile_bot_y)
            print(f"  ✓ 网格搜索找到: cx={cx:.4f} cy={cy:.4f} R={R:.4f} Ks={Ks:.4f}")

    return CircleResult(cx=cx, cy=cy, R=R, Ks=Ks,
                        pass_pile_bottom=ppb,
                        mode='differential_evolution')


# ============================================================
# 结果输出  ← 补齐 3 个参数
# ============================================================
def print_report_de(best: CircleResult,
                    level: str,
                    strut_bottom_elev: float,
                    pit_bottom_elev: float):
    ks_require = {'一级': 1.35, '二级': 1.30, '三级': 1.25}
    Ks_req     = ks_require.get(level, 1.30)
    pile_bot_y = elev_to_y(strut_bottom_elev, pit_bottom_elev)
    dist_p     = np.hypot(best.cx, best.cy - pile_bot_y)

    print(f"\n{'='*65}")
    print(f"  整体圆弧滑动稳定性（JGJ120-2012 §4.2.3）")
    print(f"{'='*65}")
    print(f"  最危险圆弧：")
    print(f"    圆心 cx = {best.cx:>10.4f} m")
    print(f"    圆心 cy = {best.cy:>10.4f} m")
    print(f"    半径  R = {best.R:>10.4f} m")
    print(f"    搜索模式: {best.mode}")
    print(f"{'─'*65}")
    print(f"  工程约束验证：")
    print(f"    桩底(0, {pile_bot_y:.3f}m) → 圆心距离 = {dist_p:.4f}m")
    print(f"    R={best.R:.4f}m {'≥' if best.R >= dist_p - 0.01 else '<'} "
          f"dist={dist_p:.4f}m "
          f"→ {'✓圆弧绕过桩底' if best.R >= dist_p - 0.01 else '✗未绕过桩底'}")
    print(f"    经过桩底附近: {'是' if best.pass_pile_bottom else '否'}")
    print(f"{'─'*65}")
    print(f"    最小安全系数 Ks  = {best.Ks:.3f}")
    print(f"    规范要求  [Ks]  = {Ks_req}  ({level}基坑)")
    print(f"    判定: {'✓ 满足' if best.Ks >= Ks_req else '✗ 不满足'}")
    print(f"{'='*65}")


