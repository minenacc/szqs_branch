"""栈桥组件模块 — 包含UI图示及二级窗口等组件。
"""

# 1. 标准库
import math
import re
import os
import tkinter as tk
from tkinter import ttk

# 2. 第三方库
import ttkbootstrap as tb
from ttkbootstrap.constants import *
import matplotlib.patches as patches
import matplotlib.lines as mlines
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

# 3. 本地模块
from General.DataUtils import midasdisttolst, _parse_simple_spacing
from General.UIHandle import create_section_selector, SplitEntry
from Trestle_Excel_io_FEM import get_properties_path, load_section_library, _parse_material_value, _sec_dim


trestle_excel_dict: dict = {}


def _beam_dim(data, detail, name_key, param_key, default_val):
    """从 _detail.params 取截面尺寸，为空时按截面名从截面库查找"""
    tp = detail.get("params", {})
    v = tp.get(param_key)
    if v not in (None, ""):
        return float(v)
    sn = data.get(name_key, "")
    if sn and sn != "/":
        lib = trestle_excel_dict.get("section_library_db", {})
        rec = lib.get(sn) or lib.get(sn.upper())
        if rec and isinstance(rec, dict):
            return float(rec.get("params", {}).get(param_key, default_val))
    return float(default_val)


# ═══════════════════════════════════════════════════════════════
# 模块级常量
# ═══════════════════════════════════════════════════════════════

# 绘图配色
DRAW_LINE_COLOR = "#2b6fd4"
DRAW_TEXT_COLOR = "#2c3e50"
DRAW_TEXT_HIGHLIGHT = "#dc3243"


def _highlight_color(base, is_highlighted):
    return '#ffacb4' if is_highlighted else base


# ═══════════════════════════════════════════════════════════════
# 通用绘图工具函数
# ═══════════════════════════════════════════════════════════════

def draw_dim_annotation(ax, x1, x2, y, label, is_highlighted=False,
                         line_color="#2b6fd4", text_color="#2c3e50",
                         tick_h=0.05, hl_color="#dc3243", lw=1.2, zorder=10):
    """在 ax 上绘制尺寸标注：水平线 + 两端竖刻度 + 居中文本。"""
    ax.plot([x1, x2], [y, y], color=line_color, lw=lw, zorder=zorder)
    ax.plot([x1, x1], [y - tick_h, y + tick_h], color=line_color, lw=lw, zorder=zorder)
    ax.plot([x2, x2], [y - tick_h, y + tick_h], color=line_color, lw=lw, zorder=zorder)
    tc = hl_color if is_highlighted else text_color
    ax.text((x1 + x2) / 2, y + 0.05, label, fontsize=8, color=tc,
            ha='center', va='bottom',
            weight='bold' if is_highlighted else 'normal', zorder=zorder)


def draw_level_mark(ax, x, y, label, above=True):
    """绘制标高"""
    dy = 0.2 if above else -0.2
    text_offset = 0.5 if above else -1.0
    ax.plot([x, x + 0.4], [y + dy, y + dy], color='green', lw=0.5)
    ax.plot([x, x + 0.2], [y + dy, y], color='green', lw=0.5)
    ax.plot([x + 0.4, x + 0.2], [y + dy, y], color='green', lw=0.5)
    ax.plot([x + 0.4, x + 1.5], [y + dy, y + dy], color='green', lw=0.5)
    ax.text(x + 0.5, y + text_offset, label, fontsize=7, color='green')


def midasdisttolst_segments(txt):
    """解析间距字符串为标注段列表（保留分组元信息，用于标注）。

    返回 [(label, total_mm), ...] 格式：
      "0,2@750+50@(795+705+705+795)+2@750" →
        [("0", 0), ("2x750", 1500), ("50x(795+705+705+795)", 160000), ("2x750", 1500)]
    """
    segments = []
    tokens = []
    depth = 0
    current = ""
    for ch in txt:
        if ch == '(':
            depth += 1; current += ch
        elif ch == ')':
            depth -= 1; current += ch
        elif ch in ('+', ',', '，') and depth == 0:
            if current.strip(): tokens.append(current.strip())
            current = ""
        else:
            current += ch
    if current.strip():
        tokens.append(current.strip())

    for token in tokens:
        m = re.match(r'^(\d+)\s*@\s*\((.+)\)$', token)
        if m:
            count = int(m.group(1))
            inner = m.group(2)
            inner_vals = _parse_simple_spacing(inner)
            total = int(sum(inner_vals) * count)
            label = f"{count}x({inner})"
            segments.append((label, total))
        else:
            parts = re.findall("[^×@*]+", token)
            if len(parts) == 1:
                val = float(parts[0])
                display = int(val) if val == int(val) else val
                segments.append((str(display), int(val) if val == int(val) else val))
            else:
                count = int(float(parts[0]))
                val = float(parts[-1])
                display_val = int(val) if val == int(val) else val
                count_str = f"{count}x{display_val}"
                segments.append((count_str, count * display_val))
    # 合并相邻相同值的段，如 [("30x750",750), ("30x750",750)] → [("60x750",45000)]
    if segments:
        merged = [segments[0]]
        for lbl, tot in segments[1:]:
            prev_lbl, prev_tot = merged[-1]
            # 从 prev_lbl 提取基准值
            m_prev = re.match(r'^(\d+)x(.+)$', prev_lbl)
            if m_prev:
                prev_base_val = int(m_prev.group(2)) if m_prev.group(2).isdigit() else None
                prev_count = int(m_prev.group(1))
            elif prev_lbl.isdigit():
                prev_base_val = int(prev_lbl)
                prev_count = 1
            else:
                prev_base_val = None
                prev_count = 0

            # 当前段的基准值
            m_cur = re.match(r'^(\d+)x(.+)$', lbl)
            if m_cur:
                cur_base_val = int(m_cur.group(2)) if m_cur.group(2).isdigit() else None
                cur_count = int(m_cur.group(1))
            elif lbl.isdigit():
                cur_base_val = int(lbl)
                cur_count = 1
            else:
                cur_base_val = None
                cur_count = 0

            if prev_base_val is not None and cur_base_val is not None and prev_base_val == cur_base_val:
                merged[-1] = (f"{prev_count + cur_count}x{prev_base_val}", (prev_count + cur_count) * prev_base_val)
            else:
                merged.append((lbl, tot))
        segments = merged
    return segments


def merge_spacing_labels(vals):
    """合并连续相等间距 → [('label', total_mm)]

    如 [750,750,750,600] → [('3x750', 2250), ('600', 600)]
    """
    if not vals:
        return []
    r = []
    cur_cnt = 1
    cur_val = vals[0]
    for v in vals[1:]:
        if abs(v - cur_val) < 1e-6:
            cur_cnt += 1
        else:
            r.append((f"{cur_cnt}x{int(cur_val) if cur_val == int(cur_val) else cur_val}" if cur_cnt > 1 else str(int(cur_val) if cur_val == int(cur_val) else cur_val), cur_cnt * cur_val))
            cur_val = v
            cur_cnt = 1
    r.append((f"{cur_cnt}x{int(cur_val) if cur_val == int(cur_val) else cur_val}" if cur_cnt > 1 else str(int(cur_val) if cur_val == int(cur_val) else cur_val), cur_cnt * cur_val))
    return r


def parse_pile_diameter(sec_text):
    """从截面文本解析桩径(m), 如 'φ820X10' -> 0.82"""
    if not sec_text or sec_text in ('/', '请选择截面'):
        return 0.5
    m = re.search(r'φ?\s*([\d.]+)\s*[*x]', sec_text)
    if m:
        return float(m.group(1)) / 1000.0
    m = re.search(r'φ?\s*([\d.]+)', sec_text)
    if m:
        return float(m.group(1)) / 1000.0
    return 0.5


# ═══════════════════════════════════════════════════════════════
# SupportSetupDialog 横桥向视图 — 完整迁移
# ═══════════════════════════════════════════════════════════════


def draw_support_transverse_view(ax, data, idx, highlight_key, sup_type):
    """绘制横桥向构造示意图"""
    ax.set_aspect('equal')
    ax.axis('off')
    title = f"{idx}# 横桥向下部结构示意图"
    ax.set_title(title, fontsize=10, pad=8)

    # 从 data 读取参数
    space_raw = data.get("pile_trans_space", "2@3000")
    try: spacing_mm = midasdisttolst(space_raw) if space_raw and space_raw.strip() else []
    except: spacing_mm = []
    pile_sec_text = data.get("pile_sec", "820X10")
    pile_w = parse_pile_diameter(pile_sec_text)

    # 颜色规定
    line_color = DRAW_LINE_COLOR; TEXT_COLOR = DRAW_TEXT_COLOR; tick_h = 0.08
    is_pile_hl  = highlight_key in ("pile", "pile_trans_space", "pile_long_space")
    is_brace_hl = highlight_key in ("brace", "brace_space", "brace_height")
    is_trans_hl = highlight_key in ("dist_trans", "dist_trans_len")

    # 桩横向布置：前缀+间距（无间距时仅画一根桩在前缀位置）
    pile_prefix_mm = spacing_mm[0] if len(spacing_mm) > 1 else (spacing_mm[0] if len(spacing_mm) == 1 else 0)
    pile_gaps_mm  = spacing_mm[1:] if len(spacing_mm) > 1 else []
    rel_piles = [pile_prefix_mm / 1000.0]
    for s in pile_gaps_mm:
        rel_piles.append(rel_piles[-1] + s / 1000.0)
    total_w = rel_piles[-1] if rel_piles else 0.0
    pile_bottom_y = -2.0
    all_x = []

    dist_beam_y = 3.0 if sup_type == "单排桩" else 2.8
    pile_top_y = dist_beam_y

    # 分配梁左端
    try: dist_trans_offset_m = float(data.get("dist_trans_offset", "0")) / 1000.0
    except: dist_trans_offset_m = 0.0
    left_edge = - dist_trans_offset_m

    # ---- 1. 钢管桩 ----
    amp = 0.08; pw = pile_w
    piles_x = [left_edge + rp for rp in rel_piles]
    for px in piles_x:
        fc = _highlight_color('#bdc3c7', is_pile_hl)
        ax.add_patch(patches.Rectangle((px - pw/2, pile_bottom_y + amp), pw,
            pile_top_y - pile_bottom_y - amp, facecolor=fc,  zorder=1))
        pts = []
        for step in range(21):
            t = step / 20.0
            pts.append((px - pw/2 + t * pw, pile_bottom_y - amp * math.sin(t * 2 * math.pi)))
        pts.append((px + pw/2, pile_bottom_y + amp + 0.1))
        pts.append((px - pw/2, pile_bottom_y + amp + 0.1))
        ax.add_patch(patches.Polygon(pts, facecolor=fc, zorder=2))
        pts_in = []
        for step in range(11):
            t = 0.5 + step / 10.0 * 0.5; y_top = pile_bottom_y - amp * math.sin(t * 2 * math.pi)
            pts_in.append((px - pw/2 + t * pw, y_top))
        for step in range(10, -1, -1):
            t = 0.5 + step / 10.0 * 0.5; y_bot = pile_bottom_y + amp * math.sin(t * 2 * math.pi)
            pts_in.append((px - pw/2 + t * pw, y_bot))
        ax.add_patch(patches.Polygon(pts_in, facecolor='#95a5a6',  zorder=1))       

    # ---- 2. 分配梁 ----
    # 横向分配梁参数
    trans_detail = data.get("dist_trans_sec_detail", {})
    h_trans = _beam_dim(data, trans_detail, "dist_trans_sec", "H", 400) / 1000.0
    trans_len = data.get("dist_trans_len", "7600")
    try: trans_len_m = float(trans_len) / 1000.0
    except: trans_len_m = max(total_w * 1.1, 4.0)
    right_edge = left_edge + trans_len_m

    trans_color = _highlight_color("#506171", is_trans_hl)

    # 单排桩-仅横向分配梁
    if sup_type == "单排桩":
        dist_trans_y = dist_beam_y
        # 绘制横向分配梁
        ax.add_patch(patches.Rectangle((left_edge, dist_trans_y), trans_len_m, h_trans,
                                        facecolor=trans_color, zorder=3))
    # 制动墩-横纵向分配梁
    else:
        # 纵向分配梁参数
        long_detail = data.get("dist_long_sec_detail", {})
        h_long = _beam_dim(data, long_detail, "dist_long_sec", "H", 400) / 1000.0
        w_long = _beam_dim(data, long_detail, "dist_long_sec", "B", 200) / 1000.0
        is_long_hl  = highlight_key in ("dist_long", "dist_long_len")
        long_color = _highlight_color('#8899aa', is_long_hl)
        # 绘制纵向分配梁
        dist_long_y = dist_beam_y
        for px in piles_x:
            ax.add_patch(patches.Rectangle((px - w_long/2, dist_long_y), w_long, h_long,
                        facecolor=long_color, lw=1, zorder=4))
        # 横向分配梁搭接于纵向分配梁顶
        dist_trans_y = dist_beam_y + h_long
        # 绘制横向分配梁
        ax.add_patch(patches.Rectangle((left_edge, dist_trans_y), trans_len_m, h_trans,
                                        facecolor=trans_color, zorder=3))
        
    # ---- 4. 联结系（支持等高多层） ----
    brace_tops = []; brace_bots = []
    if data.get("has_brace", True) and len(piles_x) >= 2:
        bs_raw = data.get("brace_space", "620")
        bh_raw = data.get("brace_height", "500")
        # "/" 和 0 均表示无下一层，过滤掉使列表只含有效值
        bs_filtered = ",".join(x for x in bs_raw.split(",") if x.strip() not in ("", "/", "0"))
        bspacing = midasdisttolst(bs_filtered) if bs_filtered.strip() else [0]
        try: bheight_m = float(bh_raw) / 1000.0
        except: bheight_m = 0.5
        # 计算各层 y（视觉压缩：标注用实际数据）
        n = len(bspacing)
        ZONE_H = 5.0        # 联结系图示总区域高度(m)
        total_actual_m = sum(bspacing) / 1000.0 + n * bheight_m
        ratio = ZONE_H / total_actual_m if total_actual_m > ZONE_H else 1.0
        for i, sp in enumerate(bspacing):
            gap = sp / 1000.0 * ratio
            if i == 0:
                t = pile_top_y - gap  # 第1层：距桩顶
            else:
                t = brace_bots[-1] - gap  # 第2+层：距上一层下弦杆
            b = t - bheight_m * ratio
            brace_tops.append(t); brace_bots.append(b)
        # 绘制
        brace_color = _highlight_color("#bdc3c7", is_brace_hl)
        brace_form = data.get("brace_form", "X")
        for k in range(len(brace_tops)):
            yt, yb = brace_tops[k], brace_bots[k]
            for j in range(len(piles_x) - 1):
                x1, x2 = piles_x[j], piles_x[j+1]
                if brace_form == "-":
                    # 单根：一条水平杆
                    ax.plot([x1, x2], [brace_tops[k], brace_tops[k]], color=brace_color, lw=1.5)
                else:
                    ax.plot([x1, x2], [yt, yt], color=brace_color, lw=1)
                    ax.plot([x1, x2], [yb, yb], color=brace_color, lw=1)
                if brace_form == "X":
                    ax.plot([x1, x2], [yt, yb], color=brace_color, ls="--", lw=1)
                    ax.plot([x1, x2], [yb, yt], color=brace_color, ls="--", lw=1)
                elif brace_form == "X|X":
                    mid = (x1 + x2) / 2.0
                    ax.plot([x1, mid], [yt, yb], color=brace_color, ls="--", lw=1)
                    ax.plot([x1, mid], [yb, yt], color=brace_color, ls="--", lw=1)
                    ax.plot([mid, x2], [yt, yb], color=brace_color, ls="--", lw=1)
                    ax.plot([mid, x2], [yb, yt], color=brace_color, ls="--", lw=1)
                    ax.plot([mid, mid], [yb, yt], color=brace_color, lw=1)  # 竖杆
                elif brace_form == "Z":
                    if j % 2 == 0:
                        ax.plot([x1, x2], [yb, yt], color=brace_color, ls="--", lw=1)  # / 形
                    else:
                        ax.plot([x1, x2], [yt, yb], color=brace_color, ls="--", lw=1)  # \ 形
        # 左侧标注（间距+层高交替，从桩顶向下）
        is_br_space_prefix_hl = (highlight_key in ("brace_space", "brace_space_left"))
        is_br_space_gap_hl = (highlight_key in ("brace_space", "brace_space_right"))
        is_br_height_hl = (highlight_key == "brace_height")
        lx = left_edge - 0.2
        all_x += [lx]
        bheight_mm = int(bheight_m * 1000)
        prev_y = pile_top_y
        for i in range(len(brace_tops)):
            # 间距段：prev_y → brace_tops[i]
            gap_label = bspacing[i]
            cy = brace_tops[i]; my = (prev_y + cy) / 2.0
            ax.plot([lx, lx], [prev_y, cy], color=line_color, lw=1.2)
            ax.plot([lx-0.05, lx+0.05], [prev_y, prev_y], color=line_color, lw=1.2)
            ax.plot([lx-0.05, lx+0.05], [cy, cy], color=line_color, lw=1.2)
            if i == 0:
                c = DRAW_TEXT_HIGHLIGHT if is_br_space_prefix_hl else TEXT_COLOR
                weight = "bold" if is_br_space_prefix_hl else "normal"
            else:
                c = DRAW_TEXT_HIGHLIGHT if is_br_space_gap_hl else TEXT_COLOR
                weight = "bold" if is_br_space_gap_hl else "normal"
            ax.text(lx-0.1, my, f"{gap_label}", fontsize=9, color=c, ha="right", va="center", weight=weight)
            prev_y = cy
            # 层高段：brace_tops[i] → brace_bots[i]
            if bheight_m > 0 and brace_form != "-":
                yb = brace_bots[i]; my = (cy + yb) / 2.0
                ax.plot([lx, lx], [cy, yb], color=line_color, lw=1.2)
                ax.plot([lx-0.05, lx+0.05], [cy, cy], color=line_color, lw=1.2)
                ax.plot([lx-0.05, lx+0.05], [yb, yb], color=line_color, lw=1.2)
                c = DRAW_TEXT_HIGHLIGHT if is_br_height_hl else TEXT_COLOR
                ax.text(lx-0.1, my, f"{bheight_mm}", fontsize=9, color=c, ha="right", va="center", weight="bold" if is_br_height_hl else "normal")
                prev_y = yb
    # ---- 5. 标注 ----
    # 桩横向间距标注（含前缀：分配梁左端至第一桩）
    is_trans_space_prefix_hl = (highlight_key in ("pile_trans_space", "pile_trans_space_left"))
    is_trans_space_gap_hl = (highlight_key in ("pile_trans_space", "pile_trans_space_right"))
    pile_dim_y = pile_bottom_y - 0.8
    if piles_x:
        L, R = left_edge, piles_x[-1]
        ax.plot([L, R], [pile_dim_y, pile_dim_y], color=line_color, lw=1.2)
        # 分配梁左端刻度
        ax.plot([left_edge, left_edge], [pile_dim_y - tick_h, pile_dim_y + 1.5*tick_h], color=line_color, lw=1.2)
        # 桩位刻度
        for px in piles_x:
            ax.plot([px, px], [pile_dim_y - tick_h, pile_dim_y + 1.5*tick_h], color=line_color, lw=1.2)
        # 前缀标注（分配梁左端 → 第一桩）= entry1
        prefix_tc = DRAW_TEXT_HIGHLIGHT if is_trans_space_prefix_hl else TEXT_COLOR
        ax.text((left_edge + piles_x[0]) / 2, pile_dim_y + 0.05, f"{int((piles_x[0] - left_edge) * 1000)}",
                fontsize=9, color=prefix_tc, ha="center", va="bottom",
                weight="bold" if is_trans_space_prefix_hl else "normal")
        # 桩间间距标注 = entry2
        gap_tc = DRAW_TEXT_HIGHLIGHT if is_trans_space_gap_hl else TEXT_COLOR
        for i in range(len(piles_x) - 1):
            x1, x2 = piles_x[i], piles_x[i+1]; mid_x = (x1 + x2) / 2.0
            ax.text(mid_x, pile_dim_y + 0.05, f"{int((x2 - x1) * 1000)}",
                    fontsize=9, color=gap_tc, ha="center", va="bottom",
                    weight="bold" if is_trans_space_gap_hl else "normal")
            
    # 横向分配梁标注
    is_trans_len_hl  = (highlight_key == "dist_trans_len")
    is_trans_cant_hl = (highlight_key == "dist_trans_cant")
    trans_cant_y = dist_trans_y + h_trans  + 0.2
    trans_leng_y = trans_cant_y + 0.6
    trans_center_x = (left_edge + right_edge) / 2.0
    # 长度
    ax.plot([left_edge, right_edge], [trans_leng_y, trans_leng_y], color=line_color, lw=1.2)
    ax.plot([left_edge, left_edge], [trans_leng_y - tick_h, trans_leng_y + tick_h], color=line_color, lw=1.2)
    ax.plot([right_edge, right_edge], [trans_leng_y - tick_h, trans_leng_y + tick_h], color=line_color, lw=1.2)
    len_text_color = DRAW_TEXT_HIGHLIGHT if is_trans_len_hl else TEXT_COLOR
    ax.text(trans_center_x, trans_leng_y + 0.05, f"{trans_len}",
            fontsize=9, color=len_text_color, ha='center', va='bottom',
            weight='bold' if is_trans_len_hl else 'normal')
    # 中心线 x=0
    ax.axvline(0, color='red', ls='--', lw=1.0, alpha=0.5, zorder=10)
    ax.text(0, pile_dim_y - 0.5, "线路中心线", fontsize=8, color="red",
            ha="center", va="bottom", alpha=0.6)
    # 分配梁左端偏移标注（中心线 → 分配梁左端）
    offset_y = trans_cant_y
    if dist_trans_offset_m != 0:
        ax.plot([0, left_edge], [offset_y, offset_y], color=line_color, lw=1.2)
        ax.plot([0, 0], [offset_y - tick_h, offset_y + tick_h], color=line_color, lw=1.2)
        ax.plot([left_edge, left_edge], [offset_y - tick_h, offset_y + tick_h], color=line_color, lw=1.2)
        ax.text((0 + left_edge) / 2, offset_y + 0.05, f"{dist_trans_offset_m*1000:.0f}",
                fontsize=9, color=DRAW_TEXT_HIGHLIGHT if is_trans_cant_hl else TEXT_COLOR,
                ha='center', va='bottom',
                weight='bold' if is_trans_cant_hl else 'normal')
    all_x = [left_edge, right_edge] + piles_x
    top_y_limit = trans_cant_y + 1.0

    margin = 0.8
    ax.set_xlim(min(all_x) - margin, max(all_x) + margin)
    ax.set_ylim(pile_dim_y - 0.8, top_y_limit + 0.5)

# ==================== 纵桥向视图（仅制动墩）====================


# ═══════════════════════════════════════════════════════════════
# SupportSetupDialog 纵桥向视图 — 完整迁移
# ═══════════════════════════════════════════════════════════════


def draw_support_longitudinal_view(ax, data, idx, highlight_key):
    """绘制纵桥向构造示意图"""
    ax.set_aspect('equal')
    ax.axis('off')
    ax.set_title(f"{idx}# 纵桥向下部结构示意图", fontsize=10, pad=8)

    space_raw = data.get("pile_long_space", "2@3000")
    try: spacing_mm = midasdisttolst(space_raw) if space_raw and space_raw.strip() else []
    except: spacing_mm = []
    pile_sec_text = data.get("pile_sec", "820X10")
    pile_w = parse_pile_diameter(pile_sec_text)

    # 颜色规定
    TEXT_COLOR = DRAW_TEXT_COLOR; line_color = DRAW_LINE_COLOR; tick_h = 0.08
    is_pile_hl  = highlight_key in ("pile", "pile_trans_space", "pile_long_space")
    is_trans_hl = highlight_key in ("dist_trans", "dist_trans_len")
    is_long_hl  = highlight_key in ("dist_long", "dist_long_len")
    is_brace_hl = highlight_key in ("brace", "brace_space", "brace_height")

    # 桩纵向布置：前缀+间距（无间距时仅画一根桩在前缀位置）
    pile_prefix_mm = spacing_mm[0] if len(spacing_mm) > 1 else (spacing_mm[0] if len(spacing_mm) == 1 else 0)
    pile_gaps_mm  = spacing_mm[1:] if len(spacing_mm) > 1 else []
    # 分配梁(纵)左端相对联中心线偏移
    try: dist_long_center_offset_m = float(data.get("dist_long_center_offset", "0")) / 1000.0
    except: dist_long_center_offset_m = 0.0
    left_edge = -dist_long_center_offset_m  # 纵向分配梁左端 = 联中心线 - 偏移值
    # 桩位从分配梁左端起算
    piles_x = [left_edge + pile_prefix_mm / 1000.0]
    for s in pile_gaps_mm:
        piles_x.append(piles_x[-1] + s / 1000.0)

    pile_bottom_y = -2.0; dist_beam_y = 2.8
    pile_top_y = dist_beam_y
    pile_dim_y = pile_bottom_y - 0.8

    # ---- 1. 钢管桩 ----
    amp = 0.08; pw = pile_w
    all_x = [0.0]
    for px in piles_x:
        fc = _highlight_color('#bdc3c7', is_pile_hl)
        ax.add_patch(patches.Rectangle((px - pw/2, pile_bottom_y + amp), pw,
            pile_top_y - pile_bottom_y - amp, facecolor=fc,  zorder=1))
        pts = []
        for step in range(21):
            t = step / 20.0
            pts.append((px - pw/2 + t * pw, pile_bottom_y - amp * math.sin(t * 2 * math.pi)))
        pts.append((px + pw/2, pile_bottom_y + amp + 0.1))
        pts.append((px - pw/2, pile_bottom_y + amp + 0.1))
        ax.add_patch(patches.Polygon(pts, facecolor=fc, zorder=2))
        pts_in = []
        for step in range(11):
            t = 0.5 + step / 10.0 * 0.5; y_top = pile_bottom_y - amp * math.sin(t * 2 * math.pi)
            pts_in.append((px - pw/2 + t * pw, y_top))
        for step in range(10, -1, -1):
            t = 0.5 + step / 10.0 * 0.5; y_bot = pile_bottom_y + amp * math.sin(t * 2 * math.pi)
            pts_in.append((px - pw/2 + t * pw, y_bot))
        ax.add_patch(patches.Polygon(pts_in, facecolor='#95a5a6',  zorder=1))

    # ---- 2. 纵向分配梁 ----
    # 纵向分配梁参数
    long_detail = data.get("dist_long_sec_detail", {})
    h_long = _beam_dim(data, long_detail, "dist_long_sec", "H", 400) / 1000.0
    w_long = _beam_dim(data, long_detail, "dist_long_sec", "B", 200) / 1000.0
    long_len = data.get("dist_long_len", "4500")
    try: long_len_m = float(long_len) / 1000.0
    except: long_len_m = 0.45
    exp_len = data.get("expansion_len", "100")
    exp_m = float(exp_len) / 1000.0
    # left_edge 由 dist_long_center_offset 确定
    right_edge = left_edge + long_len_m
    all_x += [left_edge, right_edge]
    # 纵向分配梁绘制
    long_color = _highlight_color("#8899aa", is_long_hl)
    dist_long_y = dist_beam_y
    ax.add_patch(patches.Rectangle((left_edge, dist_long_y), long_len_m, h_long,
                                    facecolor=long_color, zorder=3))

    # ---- 3. 横向分配梁（以联中心线 x=0 对称） ----
    # 横向分配梁参数
    trans_detail = data.get("dist_trans_sec_detail", {})
    h_trans = _beam_dim(data, trans_detail, "dist_trans_sec", "H", 400) / 1000.0
    w_trans = _beam_dim(data, trans_detail, "dist_trans_sec", "B", 200) / 1000.0
    left_trans_cx = -exp_m / 2.0 - 0.09  # 左侧横向分配梁（对称于联中心线 x=0）
    right_trans_cx = exp_m / 2.0 + 0.09  # 右侧横向分配梁，与左侧间距 = exp_m
    all_x += [left_trans_cx, right_trans_cx]
    # 横向分配梁（搭接于纵向分配梁顶）
    trans_color = _highlight_color("#506171", is_trans_hl)
    dist_trans_y = dist_long_y + h_long
    for cx in (left_trans_cx, right_trans_cx):
        ax.add_patch(patches.Rectangle((cx - w_trans/2, dist_trans_y), w_trans, h_trans,
                    facecolor=trans_color,  lw=1, zorder=4))
    # 联中心线（两联伸缩缝中心）
    ax.axvline(0, color="red", ls="--", lw=1.0, alpha=0.5, zorder=10)
    ax.text(0, pile_dim_y - 0.5, "联中心线", fontsize=8, color="red",
            ha="center", va="bottom", alpha=0.6)

    # ---- 4. 联结系（支持等高多层） ----
    brace_tops = []; brace_bots = []
    if data.get("has_brace", True) and len(piles_x) >= 2:
        bs_raw = data.get("brace_space", "620")
        bh_raw = data.get("brace_height", "500")
        # "/" 和 0 均表示无下一层，过滤掉使列表只含有效值
        bs_filtered = ",".join(x for x in bs_raw.split(",") if x.strip() not in ("", "/", "0"))
        bspacing = midasdisttolst(bs_filtered) if bs_filtered.strip() else [0]
        try: bheight_m = float(bh_raw) / 1000.0
        except: bheight_m = 0.5
        # 计算各层 y（视觉压缩：标注用实际数据）
        n = len(bspacing)
        ZONE_H = 5.0        # 联结系图示总区域高度(m)
        total_actual_m = sum(bspacing) / 1000.0 + n * bheight_m
        ratio = ZONE_H / total_actual_m if total_actual_m > ZONE_H else 1.0
        for i, sp in enumerate(bspacing):
            gap = sp / 1000.0 * ratio
            if i == 0:
                t = pile_top_y - gap  # 第1层：距桩顶
            else:
                t = brace_bots[-1] - gap  # 第2+层：距上一层下弦杆
            b = t - bheight_m * ratio
            brace_tops.append(t); brace_bots.append(b)
        # 绘制
        brace_color = _highlight_color("#bdc3c7", is_brace_hl)
        brace_form = data.get("brace_form", "X")
        for k in range(len(brace_tops)):
            yt, yb = brace_tops[k], brace_bots[k]
            for j in range(len(piles_x) - 1):
                x1, x2 = piles_x[j], piles_x[j+1]
                if brace_form == "-":
                    # 单根：一条水平杆
                    ax.plot([x1, x2], [brace_tops[k], brace_tops[k]], color=brace_color, lw=1.5)
                else:
                    ax.plot([x1, x2], [yt, yt], color=brace_color, lw=1)
                    ax.plot([x1, x2], [yb, yb], color=brace_color, lw=1)
                if brace_form == "X":
                    ax.plot([x1, x2], [yt, yb], color=brace_color, ls="--", lw=1)
                    ax.plot([x1, x2], [yb, yt], color=brace_color, ls="--", lw=1)
                elif brace_form == "X|X":
                    mid = (x1 + x2) / 2.0
                    ax.plot([x1, mid], [yt, yb], color=brace_color, ls="--", lw=1)
                    ax.plot([x1, mid], [yb, yt], color=brace_color, ls="--", lw=1)
                    ax.plot([mid, x2], [yt, yb], color=brace_color, ls="--", lw=1)
                    ax.plot([mid, x2], [yb, yt], color=brace_color, ls="--", lw=1)
                    ax.plot([mid, mid], [yb, yt], color=brace_color, lw=1)  # 竖杆
                elif brace_form == "Z":
                    if j % 2 == 0:
                        ax.plot([x1, x2], [yb, yt], color=brace_color, ls="--", lw=1)  # / 形
                    else:
                        ax.plot([x1, x2], [yt, yb], color=brace_color, ls="--", lw=1)  # \ 形
        # 左侧标注（间距+层高交替，从桩顶向下）
        is_br_space_prefix_hl = (highlight_key in ("brace_space", "brace_space_left"))
        is_br_space_gap_hl = (highlight_key in ("brace_space", "brace_space_right"))
        is_br_height_hl = (highlight_key == "brace_height")
        lx = left_edge - 0.2
        all_x += [lx]
        bheight_mm = int(bheight_m * 1000)
        prev_y = pile_top_y
        for i in range(len(brace_tops)):
            # 间距段：prev_y → brace_tops[i]
            gap_label = bspacing[i]
            cy = brace_tops[i]; my = (prev_y + cy) / 2.0
            ax.plot([lx, lx], [prev_y, cy], color=line_color, lw=1.2)
            ax.plot([lx-0.05, lx+0.05], [prev_y, prev_y], color=line_color, lw=1.2)
            ax.plot([lx-0.05, lx+0.05], [cy, cy], color=line_color, lw=1.2)
            if i == 0:
                c = DRAW_TEXT_HIGHLIGHT if is_br_space_prefix_hl else TEXT_COLOR
                weight = "bold" if is_br_space_prefix_hl else "normal"
            else:
                c = DRAW_TEXT_HIGHLIGHT if is_br_space_gap_hl else TEXT_COLOR
                weight = "bold" if is_br_space_gap_hl else "normal"
            ax.text(lx-0.1, my, f"{gap_label}", fontsize=9, color=c, ha="right", va="center", weight=weight)
            prev_y = cy
            # 层高段：brace_tops[i] → brace_bots[i]
            if bheight_m > 0 and brace_form != "-":
                yb = brace_bots[i]; my = (cy + yb) / 2.0
                ax.plot([lx, lx], [cy, yb], color=line_color, lw=1.2)
                ax.plot([lx-0.05, lx+0.05], [cy, cy], color=line_color, lw=1.2)
                ax.plot([lx-0.05, lx+0.05], [yb, yb], color=line_color, lw=1.2)
                c = DRAW_TEXT_HIGHLIGHT if is_br_height_hl else TEXT_COLOR
                ax.text(lx-0.1, my, f"{bheight_mm}", fontsize=9, color=c, ha="right", va="center", weight="bold" if is_br_height_hl else "normal")
                prev_y = yb
    # ---- 4. 标注 ----
    # 桩纵向间距标注（含前缀：分配梁左端至第一桩）
    is_long_space_prefix_hl = (highlight_key in ("pile_long_space", "pile_long_space_left"))
    is_long_space_gap_hl = (highlight_key in ("pile_long_space", "pile_long_space_right"))
    if piles_x:
        L, R = left_edge, piles_x[-1]
        ax.plot([L, R], [pile_dim_y, pile_dim_y], color=line_color, lw=1.2)
        # 分配梁左端刻度
        ax.plot([left_edge, left_edge], [pile_dim_y - tick_h, pile_dim_y + 1.5*tick_h], color=line_color, lw=1.2)
        # 桩位刻度
        for px in piles_x:
            ax.plot([px, px], [pile_dim_y - tick_h, pile_dim_y + 1.5*tick_h], color=line_color, lw=1.2)
        # 前缀标注（分配梁左端 → 第一桩）= entry1
        prefix_tc = DRAW_TEXT_HIGHLIGHT if is_long_space_prefix_hl else TEXT_COLOR
        ax.text((left_edge + piles_x[0]) / 2, pile_dim_y + 0.05, f"{int((piles_x[0] - left_edge) * 1000)}",
                fontsize=9, color=prefix_tc, ha="center", va="bottom",
                weight="bold" if is_long_space_prefix_hl else "normal")
        # 桩间间距标注 = entry2
        gap_tc = DRAW_TEXT_HIGHLIGHT if is_long_space_gap_hl else TEXT_COLOR
        for i in range(len(piles_x) - 1):
            x1, x2 = piles_x[i], piles_x[i+1]; mid_x = (x1 + x2) / 2.0
            ax.text(mid_x, pile_dim_y + 0.05, f"{int((x2 - x1) * 1000)}",
                    fontsize=9, color=gap_tc, ha="center", va="bottom",
                    weight="bold" if is_long_space_gap_hl else "normal")
    
    # 纵向分配梁标注
    # 长度
    is_long_len_hl = (highlight_key == "dist_long_len")
    long_len_y = dist_trans_y + 1.2
    long_center_x = (left_edge + right_edge) / 2.0
    ax.plot([left_edge, right_edge], [long_len_y, long_len_y], color=line_color, lw=1.2)
    ax.plot([left_edge, left_edge], [long_len_y - tick_h, long_len_y + tick_h], color=line_color, lw=1.2)
    ax.plot([right_edge, right_edge], [long_len_y - tick_h, long_len_y + tick_h], color=line_color, lw=1.2)
    len_text_color = DRAW_TEXT_HIGHLIGHT if is_long_len_hl else TEXT_COLOR
    ax.text(long_center_x, long_len_y + 0.05, long_len,
            fontsize=9, color=len_text_color, ha='center', va='bottom',
            weight='bold' if is_long_len_hl else 'normal')
    # 联中心线偏移标注（联中心线 x=0 → 分配梁左端）
    is_center_offset_hl = (highlight_key == "dist_long_center_offset")
    center_offset_y = dist_trans_y + 0.6
    
    if left_edge != 0:
        ax.plot([0, left_edge], [center_offset_y, center_offset_y], color=line_color, lw=1.2)
        ax.plot([0, 0], [center_offset_y - tick_h, center_offset_y + tick_h], color=line_color, lw=1.2)
        ax.plot([left_edge, left_edge], [center_offset_y - tick_h, center_offset_y + tick_h], color=line_color, lw=1.2)
        offset_text_color = DRAW_TEXT_HIGHLIGHT if is_center_offset_hl else TEXT_COLOR
        ax.text((0 + left_edge)/2, center_offset_y + 0.05, f"{int(dist_long_center_offset_m * 1000)}",
                fontsize=9, color=offset_text_color, ha='center', va='bottom',
                weight='bold' if is_center_offset_hl else 'normal')
    top_y_limit = center_offset_y + 1.0

    all_x = [left_edge, right_edge, 0] + piles_x
    ax.set_xlim(min(all_x) - 0.8, max(all_x) + 0.8)
    ax.set_ylim(pile_dim_y - 0.9, top_y_limit + 0.5)


# ═══════════════════════════════════════════════════════════════
# SupportsTab 立面图 — 完整迁移
# ═══════════════════════════════════════════════════════════════


def draw_support_elevation(tab, highlight_idx=-1, trestle_excel_dict=None):
    """立面示意图：跨径长度、管桩长度、标高均使用实际数据"""
    if trestle_excel_dict is None:
        trestle_excel_dict = {}
    tab.ax.clear()
    spans = getattr(tab, 'spans_data', [])
    if not spans or len(tab.support_data) < len(spans) + 1:
        tab.canvas.draw()
        return

    # --- 实际标高基准 ---
    try:
        deck_y = float(tab.deck_level.get().replace('+', ''))
    except (ValueError, AttributeError):
        deck_y = 12.5

    gap_val = 0.3  # 伸缩缝宽度（示意）
    current_x = 0
    offset_accum = 0

    # --- 第一遍：计算各桩位 x 坐标和桩底标高 ---
    pile_positions = []  # [(draw_x, pile_bot, sup_type), ...]
    tmp_x = 0; tmp_accum = 0
    for i in range(len(spans) + 1):
        st = tab.support_data[i]["type"]
        if st == '制动墩':
            dx_draw = tmp_x + tmp_accum + (gap_val / 2)
        else:
            dx_draw = tmp_x + tmp_accum
        if st in ('重力式桥台', '简易桥台'):
            pile_positions.append((dx_draw, deck_y - 2 if st == '重力式桥台' else deck_y - 1, st))
        else:
            try:
                pl = float(tab.support_data[i].get("pile_len", "15"))
            except (ValueError, TypeError):
                pl = 15.0
            pile_positions.append((dx_draw, deck_y - pl, st))
        if i < len(spans):
            if st == '制动墩':
                tmp_accum += gap_val
            tmp_x += spans[i]

    label_y = min(pb for _, pb, _ in pile_positions) - 1.0

    all_y_min = [deck_y]

    for i in range(len(spans) + 1):
        sup_type = tab.support_data[i]["type"]
        color = '#e74c3c' if i == highlight_idx else '#95a5a6'

        # --- A. 计算当前支点绘图位置 ---
        if sup_type == '制动墩':
            draw_x = current_x + offset_accum + (gap_val / 2)
        else:
            draw_x = current_x + offset_accum

        # --- B. 绘制下部支撑 ---
        if sup_type == '制动墩':
            # 读取实际桩长
            try:
                pile_len = float(tab.support_data[i].get("pile_len", "15"))
            except (ValueError, TypeError):
                pile_len = 15.0
            pile_bot = deck_y - pile_len
            all_y_min.append(pile_bot)
            f2_w, f2_h = 2.8, 0.6
            f1_w, f1_h = 0.5, 0.6
            spacing = 1.0
            # 钢管桩（实际桩长，与单排桩同样式）
            for dx in [-spacing / 2, spacing / 2]:
                tab.ax.plot([draw_x + dx, draw_x + dx],
                             [pile_bot, deck_y - f2_h - f1_h],
                             color=color, lw=4)
            # F2 分配梁
            tab.ax.add_patch(patches.Rectangle(
                (draw_x - f2_w / 2, deck_y - f2_h - f1_h),
                f2_w, f2_h, facecolor=color))
            # F1 分配梁
            for dx in [-0.5, 0.5]:
                tab.ax.add_patch(patches.Rectangle(
                    (draw_x + dx - f1_w / 2, deck_y - f1_h),
                    f1_w, f1_h, facecolor=color))

        elif sup_type == '重力式桥台':
            direction = 1 if i == 0 else -1
            pts = [(-1.5 * direction, 0), (0.5 * direction, 0),
                   (0.5 * direction, -1.5), (0.8 * direction, -1.5),
                   (0.8 * direction, -2), (-1.0 * direction, -2),
                   (-1.0 * direction, -1.5), (-0.5 * direction, -1.5),
                   (-0.5 * direction, -1.0), (-1.5 * direction, -0.4)]
            actual_pts = [(current_x + dx, deck_y + dy) for dx, dy in pts]
            tab.ax.add_patch(patches.Polygon(
                actual_pts, facecolor=color, alpha=0.9, edgecolor='none'))
            all_y_min.append(deck_y - 2)

        elif sup_type == '简易桥台':
            tab.ax.add_patch(patches.Rectangle(
                (current_x - 1, deck_y - 1), 2, 1,
                facecolor=color, alpha=0.7))
            all_y_min.append(deck_y - 1)

        else:
            # 单排桩 / 制动墩以外的桩 — 按实际桩长绘制
            try:
                pile_len = float(tab.support_data[i].get("pile_len", "15"))
            except (ValueError, TypeError):
                pile_len = 15.0
            pile_bot = deck_y - pile_len
            all_y_min.append(pile_bot)
            tab.ax.plot([draw_x, draw_x], [pile_bot, deck_y - 0.5],
                         color=color, lw=4)

        # 支点编号标注（统一在最下方）
        tab.ax.text(draw_x, label_y, f"{i}#", ha='center', fontsize=8,
                     color='black' if i != highlight_idx else 'red',
                     fontweight='bold')

        # --- C. 绘制桥面系和跨径标注 ---
        if i < len(spans):
            span_len = spans[i]
            p_start_x = draw_x
            if sup_type == '制动墩':
                offset_accum += gap_val
                p_start_x = current_x + offset_accum
            next_draw_x = (current_x + span_len) + offset_accum
            tab.ax.plot([p_start_x, next_draw_x], [deck_y, deck_y],
                         color='#34495e', lw=6, solid_capstyle='butt')
            tab.ax.text(draw_x + span_len / 2, deck_y + 0.6,
                         f"{span_len}m", ha='center', fontsize=8, color='blue')
            current_x += span_len

    # --- 标高标注 ---
    mark_x = -10.0
    draw_level_mark(tab.ax,mark_x, deck_y, f"桥面高程({deck_y:.1f}m)")

    # 设防水位
    try:
        water_val = float(tab.water_level.get().replace('+', ''))
        tab.ax.axhline(y=water_val, color='cyan', linestyle='--',
                        linewidth=1, alpha=0.8, zorder=0)
        draw_level_mark(tab.ax,mark_x, water_val, f"设防水位({water_val:.1f}m)", above=False)
        all_y_min.append(water_val)
    except (ValueError, AttributeError):
        pass

    # --- 地面线（裁剪至栈桥起止桩位 x 范围） ---
    gl = trestle_excel_dict.get("ground_line", [])
    if gl:
        x_origin = float(trestle_excel_dict.get("basic", {}).get("x_origin", 0) or 0)
        x_start = x_origin
        x_end = x_origin + current_x + offset_accum
        clipped = [(x, y) for x, y in gl if x_start <= x <= x_end]
        if clipped:
            xs_gl = [x - x_origin for x, _ in clipped]
            ys_gl = [y for _, y in clipped]
            tab.ax.plot(xs_gl, ys_gl, color="#8B4513", lw=1.0, alpha=0.8, zorder=1)
            all_y_min.extend(ys_gl)

    # --- 视口范围 ---
    y_min = min(all_y_min) - 2 if all_y_min else deck_y - 20
    tab.ax.set_xlim(mark_x - 1, current_x + offset_accum + 3)
    tab.ax.set_ylim(y_min, deck_y + 3)
    tab.ax.axis('off')
    tab.fig.tight_layout(pad=0)
    tab.canvas.draw()


# ═══════════════════════════════════════════════════════════════
# ComponentsTab 横桥向 — 完整迁移
# ═══════════════════════════════════════════════════════════════

def draw_components_transverse_diagram(tab):
    """横桥向上部结构图示"""
    if not hasattr(tab, 'ax_trans'):
        return
    tab.ax_trans.clear()
    hl_key = getattr(tab, '_hl_key', None)

    line_color = DRAW_LINE_COLOR; TEXT_COLOR = DRAW_TEXT_COLOR; TEXT_HIGHLIGHT = DRAW_TEXT_HIGHLIGHT; tick_h = 0.05

    # ==== 1、参数获取 ====
    # 概念中心线
    trestle_mid = 0.0
    # 桥面参数（从 SupportsTab 读取）
    try:
        bw = tab.parent_ui.tab_supports.bridge_width.get() if tab.parent_ui and hasattr(tab.parent_ui, 'tab_supports') else "6"
        bridge_width = float(bw)
    except:
        bridge_width = 6.0
    try: deck_t_m = max(float(tab.deck_t.get()) / 1000.0, 0.04)
    except: deck_t_m = 0.04
    try: deck_trans_ecc = float(tab.deck_trans_ecc.get()) / 1000.0
    except: deck_trans_ecc = 2.7
    # 桥面系类型
    deck_type = tab.deck_type_var.get() if hasattr(tab, 'deck_type_var') else "单层钢面板"
    is_double = (deck_type == "双层钢面板")
    is_concrete = (deck_type == "混凝土桥面板")

    # 横肋参数（混凝土面板无横肋）
    rib_trans_left = 0 - deck_trans_ecc
    rib_trans_right = rib_trans_left + bridge_width
    x_all = [rib_trans_right, rib_trans_left]
    rib_long_space_raw = ""
    h_rib_trans = 0.0
    if is_concrete:
        rib_positions = []
        h_rib_long = 0.0; b_rib_long = 0.0; rib_trans_d2_m = 0.0; rib_long_space = []
    else:
        rt = tab.rib_trans_sec.get_section_data()
        rtp = rt.get("params", {}) if rt else {}
        h_rib_trans = float(rtp.get("H", 250)) / 1000.0
        if not rtp and rt and rt.get("section_name"):
            _db = tab._get_section_db()
            sn = rt["section_name"]
            if sn in _db:
                h_rib_trans = _sec_dim(_db[sn], "H", 250) / 1000.0

    # 纵肋参数（仅双层桥面系，含 DB 兜底确保初始绘图正确）
    if not is_concrete and is_double and hasattr(tab, 'rib_long_sec'):
        try:
            rl = tab.rib_long_sec.get_section_data()
            rlp = rl.get("params", {}) if rl else {}
        except:
            rlp = {}
        h_rib_long = float(rlp.get("H", 140)) / 1000.0
        b_rib_long = float(rlp.get("B", 80)) / 1000.0
        if not rlp and rl and rl.get("section_name"):
            sec_path2 = get_properties_path()
            if sec_path2:
                _db = load_section_library(sec_path2)
            else:
                _db = {}
            sn = rl["section_name"]
            if sn in _db:
                h_rib_long = _sec_dim(_db[sn], "H", 140) / 1000.0
                b_rib_long = _sec_dim(_db[sn], "B", 80) / 1000.0
        # rib_long_space 格式 "a,b@x"，a=横肋悬臂d2(已合并)，b@x=纵肋横向布置
        try:
            rib_long_space_raw = tab.rib_long_space.get()
            _rls_raw = rib_long_space_raw.strip() if rib_long_space_raw else ""
            _rls_parts = _rls_raw.split(",", 1)
            rib_trans_d2_m = float(_rls_parts[0].strip()) / 1000.0 if _rls_parts[0].strip() else 0.0
            # 有逗号时取第二段；无逗号说明 entry2 为空，纵肋布置取空
            rib_long_space = midasdisttolst(_rls_parts[1].strip()) if len(_rls_parts) > 1 and _rls_parts[1].strip() else []
        except:
            rib_trans_d2_m = 0.0; rib_long_space = []
        # 纵肋 x 坐标：最左侧 = 横肋左端 + a(前缀)，向右按 rib_long_space 排布
        leftmost_rib_x = rib_trans_left + rib_trans_d2_m
        rib_positions = [leftmost_rib_x]
        cur_x = leftmost_rib_x
        for sp in rib_long_space:
            cur_x += sp / 1000.0
            if cur_x > rib_trans_right:
                break
            rib_positions.append(cur_x)

    # 纵梁参数
    try:
        beam_space = midasdisttolst(tab.beam_space.get())
    except:
        beam_space = midasdisttolst("5@900")
    try: beam_first_space_m = float(tab.beam_first_space.get()) / 1000.0
    except: beam_first_space_m = 0.0
    # 纵梁截面尺寸：型钢按实际 B/H，贝雷梁用固定默认值
    beam_h, beam_w = tab._get_beam_dims()

    # 纵梁 x 坐标：以线路中心为基准，beam_first_space 为第一排间距
    leftmost_beam_x = - beam_first_space_m
    beam_pos = [leftmost_beam_x]
    cur_x = leftmost_beam_x
    for sp in beam_space:
        cur_x += sp / 1000.0
        if cur_x > rib_trans_right:
            break
        beam_pos.append(cur_x)

    # 分配梁参数（从 SupportsTab 取最短，首次读取后缓存）
    if tab._dist_beam_cache is None and tab.parent_ui and hasattr(tab.parent_ui, 'tab_supports'):
        tab._dist_beam_cache = tab.parent_ui.tab_supports.get_transverse_beam_data()
    if tab._dist_beam_cache and tab._dist_beam_cache.get("length_mm") is not None:
        dist_len_m = tab._dist_beam_cache["length_mm"] / 1000.0
        dist_h_m = tab._dist_beam_cache["height_mm"] / 1000.0
        dist_offset_m = tab._dist_beam_cache.get("offset_mm", 3750) / 1000.0
    else:
        dist_len_m = 8.0; dist_h_m = 0.4; dist_offset_m = 3.75
    dist_left = -dist_offset_m
    dist_right = dist_left + dist_len_m

    # ==== 2、标高定义 ====
    deck_top = 0.0                              # 桥面顶
    deck_bottom = deck_top - deck_t_m           # 桥面底
    if is_concrete:
        rib_trans_y = deck_bottom               # 混凝土面板无横肋，从桥面底开始
        rib_long_y = deck_bottom
    elif is_double:
        rib_long_y = deck_bottom - h_rib_long   # 纵肋底（双层）
        rib_trans_y = rib_long_y - h_rib_trans  # 横肋底（双层）
    else:
        rib_trans_y = deck_bottom - h_rib_trans # 横肋底
    beam_y = rib_trans_y - beam_h           # 纵梁底
    dist_y = beam_y - dist_h_m                # 分配梁底

    # ==== 3、绘图逻辑 ====
    # 桥面板
    is_deck_hl = (hl_key == "deck")
    deck_color = _highlight_color('#bdc3c7', is_deck_hl)
    tab.ax_trans.add_patch(patches.Rectangle((rib_trans_left, deck_bottom), bridge_width, deck_t_m,
                            facecolor=deck_color, zorder=1))
    # 桥面系（混凝土面板无横肋/纵肋）
    if not is_concrete:
        is_rib_trans_hl = hl_key in ("rib_trans", "rib_trans_space", "rib_trans_space_left", "rib_trans_space_right")
        rib_trans_color = _highlight_color('#506171', is_rib_trans_hl)
        if is_double:
            # 纵肋（矩形阵列）
            is_rib_long_hl = hl_key in ("rib_long", "rib_long_space")
            rib_long_color = _highlight_color('#7a8fa0', is_rib_long_hl)
            if hasattr(tab, 'rib_long_sec'):
                for rx in rib_positions:
                    tab.ax_trans.add_patch(patches.Rectangle(
                        (rx - b_rib_long/2, rib_long_y), b_rib_long, h_rib_long,
                        facecolor=rib_long_color, zorder=2))
            # 横肋
            tab.ax_trans.add_patch(patches.Rectangle((rib_trans_left, rib_trans_y), bridge_width, h_rib_trans,
                                    facecolor=rib_trans_color, zorder=3))
        else:
            # 横肋
            tab.ax_trans.add_patch(patches.Rectangle((rib_trans_left, rib_trans_y), bridge_width, h_rib_trans,
                                    facecolor=rib_trans_color, zorder=3))
    # 纵梁
    is_beam_hl = (hl_key == "beam_space")
    beam_color = _highlight_color("#3D4D5A", is_beam_hl)
    for bx in beam_pos:
        tab.ax_trans.add_patch(patches.Rectangle((bx - beam_w/2, beam_y), beam_w, beam_h,
                                facecolor=beam_color, zorder=4))
    # 分配梁
    tab.ax_trans.add_patch(patches.Rectangle((dist_left, dist_y), dist_len_m, dist_h_m,
                            facecolor="#bdc3c7", zorder=5))
    # 轴线
    tab.ax_trans.axvline(trestle_mid, color='red', ls='--', lw=1, alpha=0.4, zorder=10)
    # ==== 4、标注 ====
    # 标注标高设置
    rib_width_label_y = deck_bottom + 0.2
    deck_ecc_label_y = rib_trans_y - 0.3
    if is_double:
        rib_long_space_y = rib_width_label_y + 0.3
    beam_space_y = dist_y - 0.3
    beam_first_space_y = beam_space_y - 0.4

    # 高亮逻辑
    is_bridge_width_hl = (hl_key == "bridge_width")
    is_deck_ecc_hl = (hl_key == "deck_trans_ecc")
    is_rib_long_space_left_hl = (hl_key in ("rib_long_space", "rib_long_space_left"))
    is_rib_long_space_right_hl = (hl_key in ("rib_long_space", "rib_long_space_right"))
    is_beam_space_hl = (hl_key == "beam_space")
    is_beam_first_space_hl = (hl_key == "beam_first_space")

    # 桥面宽度标注
    draw_dim_annotation(tab.ax_trans, rib_trans_left, rib_trans_right, rib_width_label_y,
                         f"{bridge_width*1000:.0f}", is_highlighted=is_bridge_width_hl,
                         line_color=line_color, text_color=TEXT_COLOR, tick_h=tick_h, hl_color=TEXT_HIGHLIGHT)
    # 桥面左端间距标注（deck_trans_ecc）
    draw_dim_annotation(tab.ax_trans, rib_trans_left, 0, deck_ecc_label_y,
                         f"{deck_trans_ecc*1000:.0f}", is_highlighted=is_deck_ecc_hl,
                         line_color=line_color, text_color=TEXT_COLOR, tick_h=tick_h, hl_color=TEXT_HIGHLIGHT)
    if is_double:
        # 纵肋起始间距标注（rib_long_space 前缀 a 值 = 原横肋悬臂d2）
        draw_dim_annotation(tab.ax_trans, rib_trans_left, leftmost_rib_x, rib_long_space_y,
                             f"{rib_trans_d2_m*1000:.0f}", is_highlighted=is_rib_long_space_left_hl,
                             line_color=line_color, text_color=TEXT_COLOR, tick_h=tick_h, hl_color=TEXT_HIGHLIGHT)
        # 纵肋横向布置（等间距段仅绘两侧竖刻度）
        if len(rib_positions) > 1:
            segs = merge_spacing_labels(rib_long_space)
            left_x = rib_positions[0]; right_x = rib_positions[-1]
            tab.ax_trans.plot([left_x, right_x], [rib_long_space_y, rib_long_space_y], color=line_color, lw=1.2, zorder=10)
            cur = left_x
            for i, (lbl, tot_mm) in enumerate(segs):
                tot_m = tot_mm / 1000.0
                end = cur + tot_m
                tab.ax_trans.plot([cur, cur], [rib_long_space_y - tick_h, rib_long_space_y + tick_h], color=line_color, lw=1.2, zorder=10)
                tc = TEXT_HIGHLIGHT if is_rib_long_space_right_hl else TEXT_COLOR
                tab.ax_trans.text((cur + end) / 2, rib_long_space_y + 0.05, lbl, fontsize=8, color=tc, ha='center', va='bottom', weight='bold' if is_rib_long_space_right_hl else 'normal', zorder=10)
                cur = end
            tab.ax_trans.plot([right_x, right_x], [rib_long_space_y - tick_h, rib_long_space_y + tick_h], color=line_color, lw=1.2, zorder=10)
    # 纵梁横向布置（等间距段仅绘两侧竖刻度）
    if len(beam_pos) > 1:
        segs = merge_spacing_labels(beam_space)
        left_x = beam_pos[0]; right_x = beam_pos[-1]
        tab.ax_trans.plot([left_x, right_x], [beam_space_y, beam_space_y], color=line_color, lw=1.2, zorder=10)
        cur = left_x
        for i, (lbl, tot_mm) in enumerate(segs):
            tot_m = tot_mm / 1000.0
            end = cur + tot_m
            tab.ax_trans.plot([cur, cur], [beam_space_y - tick_h, beam_space_y + tick_h], color=line_color, lw=1.2, zorder=10)
            tc = TEXT_HIGHLIGHT if is_beam_space_hl else TEXT_COLOR
            tab.ax_trans.text((cur + end) / 2, beam_space_y + 0.05, lbl, fontsize=8, color=tc, ha='center', va='bottom', weight='bold' if is_beam_space_hl else 'normal', zorder=10)
            cur = end
        tab.ax_trans.plot([right_x, right_x], [beam_space_y - tick_h, beam_space_y + tick_h], color=line_color, lw=1.2, zorder=10)
    # 第一根纵梁相对线路中心线间距(beam_first_space)
    if len(beam_pos) > 0:
        draw_dim_annotation(tab.ax_trans, leftmost_beam_x, 0, beam_first_space_y,
                             f"{beam_first_space_m*1000:.0f}", is_highlighted=is_beam_first_space_hl,
                             line_color=line_color, text_color=TEXT_COLOR, tick_h=tick_h, hl_color=TEXT_HIGHLIGHT)
    # 中心线标注
    center_text_y = beam_first_space_y - 0.4
    tab.ax_trans.text(0, center_text_y, "线路中心线", fontsize=8, color="red",
            ha="center", va="bottom", alpha=0.6)
    
    # 视口（收集所有构件 x 极值）
    x_all.extend([dist_left, dist_right])
    x_all.extend(beam_pos)
    if is_double and hasattr(tab, 'rib_long_sec'):
        x_all.extend(rib_positions)
    margin = 0.5
    tab.ax_trans.set_xlim(min(x_all) - margin, max(x_all) + margin)
    top_y = deck_top + margin
    bottom_y = center_text_y
    tab.ax_trans.set_ylim(bottom_y, top_y)
    tab.ax_trans.axis('off')
    tab.fig_trans.tight_layout(pad=0)
    tab.canvas_trans.draw()
    # 缓存几何数据供 MoveLoadTab2 车辆加载图复用
    _rib_w = bridge_width
    _beam_y = beam_y
    _beam_h = beam_h
    _beam_w = beam_w
    _dist_h = dist_h_m
    _pos = list(beam_pos) if beam_pos else []
    _rib_xs = list(rib_positions) if is_double and hasattr(tab, 'rib_long_sec') else []
    tab._diagram_geometry = {
        "deck_bottom": deck_bottom, "deck_top": deck_top, "deck_width": bridge_width,
        "is_double": is_double,
        "rib_trans_y": rib_trans_y, "h_rib_trans": h_rib_trans, "rib_trans_w": _rib_w,
        "rib_trans_left": rib_trans_left, "rib_trans_right": rib_trans_right,
        "rib_long_y": rib_long_y if is_double else None,
        "h_rib_long": h_rib_long if is_double else None,
        "b_rib_long": b_rib_long if is_double else None,
        "rib_positions": _rib_xs,
        "beam_y": _beam_y, "beam_h": _beam_h, "beam_w": _beam_w,
        "beam_pos": _pos,
        "dist_y": dist_y, "dist_h": _dist_h, "dist_left": dist_left, "dist_right": dist_right,
    }

# ==================== 平面上部结构图示 ====================


# ═══════════════════════════════════════════════════════════════
# ComponentsTab 平面图 — 完整迁移
# ═══════════════════════════════════════════════════════════════

def draw_components_plan_diagram(tab):
    """平面示意图（纵桥向= X轴，横桥向= Y轴；纵梁 + 横肋）"""
    if not hasattr(tab, 'ax_plan'):
        return
    tab.ax_plan.clear()
    hl_key = getattr(tab, '_hl_key', None)
    line_color = DRAW_LINE_COLOR; TEXT_COLOR = DRAW_TEXT_COLOR; TEXT_HIGHLIGHT = DRAW_TEXT_HIGHLIGHT
    tick_h = 0.15
    # 纵梁宽度（平面图只体现翼缘宽 B）：型钢按实际截面 B，贝雷梁用固定默认值
    beam_w = tab._get_beam_dims()[1]

    # ==== 参数 ====
    try: beam_space = midasdisttolst(tab.beam_space.get())
    except: beam_space = [900, 900, 900, 900, 900]
    try:
        beam_cant = tab.beam_cant.get()
        beam_cant_parts = beam_cant.split(",")
        beam_cant_left_m = float(beam_cant_parts[0].strip()) / 1000.0 if beam_cant_parts[0].strip() else 0.0
        beam_cant_right_m = float(beam_cant_parts[1].strip()) / 1000.0 if len(beam_cant_parts) > 1 and beam_cant_parts[1].strip() else 0.0
    except:
        beam_cant_left_m = 0.0; beam_cant_right_m = 0.0
    try:
        bw = tab.parent_ui.tab_supports.bridge_width.get() if tab.parent_ui and hasattr(tab.parent_ui, 'tab_supports') else "6"
        bridge_width = float(bw)
    except:
        bridge_width = 6.0
    _rts_raw = ""
    if hasattr(tab, 'rib_trans_space'):
        try: _rts_raw = tab.rib_trans_space.get().strip()
        except: _rts_raw = ""
    try: rib_trans_space = midasdisttolst(_rts_raw) if _rts_raw else []
    except: rib_trans_space = []
    try: deck_trans_ecc = float(tab.deck_trans_ecc.get()) / 1000.0
    except: deck_trans_ecc = 0.0
    try: beam_first_space = float(tab.beam_first_space.get()) / 1000.0
    except: beam_first_space = 0.0
    deck_type = tab.deck_type_var.get() if hasattr(tab, 'deck_type_var') else "单层钢面板"
    is_concrete = (deck_type == "混凝土桥面板")
    if is_concrete:
        b_rib_trans = 0.0
    else:
        rt = tab.rib_trans_sec.get_section_data()
        rtp = rt.get("params", {}) if rt else {}
        b_rib_trans = float(rtp.get("B", 116)) / 1000.0

    # 跨径（从 SupportsTab读取），以 0# 桩为 x 基准点
    spans_mm = []
    if tab.parent_ui and hasattr(tab.parent_ui, 'tab_supports'):
        st = tab.parent_ui.tab_supports
        try: spans_mm = getattr(st, 'spans_data', None) or midasdisttolst(st.span.get())
        except: pass
    if not spans_mm: spans_mm = [6000]
    span_len_m = sum(spans_mm)
    x_origin = 0.0  # 0# 桩 x 坐标

    # 纵梁参数
    beam_top_y = beam_first_space
    beam_ys = [beam_top_y]; cur_y = beam_top_y
    for s in beam_space:
        cur_y += - s / 1000.0
        beam_ys.append(cur_y)
    beam_bottom_y = cur_y
    beam_left = x_origin - beam_cant_left_m
    beam_right = x_origin + span_len_m + beam_cant_right_m
    beam_len = beam_right - beam_left

    # 横肋参数：第一根横肋 = beam_left + prefix_value
    _rts_raw2 = ""
    if hasattr(tab, 'rib_trans_space'):
        try: _rts_raw2 = tab.rib_trans_space.get().strip()
        except: _rts_raw2 = ""
    try: rib_trans_space = midasdisttolst(_rts_raw2) if _rts_raw2 else []
    except: rib_trans_space = []
    rib_xs = []
    cur_x = beam_left  # 从纵梁左端开始
    for s in rib_trans_space:
        cur_x += s / 1000.0
        if cur_x > beam_right: break
        rib_xs.append(cur_x)
    # 混凝土面板无横肋，rib_xs 保持空（不加兜底）

    # ==== 高亮 ====
    is_beam_hl = hl_key == "beam_space"
    is_beam_cant_left_hl = (hl_key == "beam_cant_left")
    is_beam_cant_right_hl = (hl_key == "beam_cant_right")
    is_rib_trans_hl = hl_key in ("rib_trans", "rib_trans_space", "rib_trans_space_left", "rib_trans_space_right")
    is_deck_rib_layout_hl = (hl_key == "rib_trans_space")
    is_deck_ext_left_hl = (hl_key == "deck_end_ext_left")
    is_deck_ext_right_hl = (hl_key == "deck_end_ext_right")

    # 纵梁
    for by in beam_ys:
        tab.ax_plan.add_patch(patches.Rectangle((beam_left, by - beam_w / 2), beam_len, beam_w, facecolor=_highlight_color('#bdc3c7', is_beam_hl), zorder=2))

    # 横肋
    for rx in rib_xs:
        tab.ax_plan.add_patch(patches.Rectangle((rx - b_rib_trans / 2, deck_trans_ecc), b_rib_trans, -bridge_width, facecolor=_highlight_color('#506171', is_rib_trans_hl), zorder=3))

    # 横桥向中心线
    tab.ax_plan.axhline(0, color='blue', ls='--', lw=1.0, alpha=0.6)
    # 起始桩（0#）和终点桩（n#）竖向虚线
    x_start = x_origin
    x_end = x_origin + span_len_m
    tab.ax_plan.axvline(x_start, color='green', ls='--', lw=1.2, alpha=0.7)
    tab.ax_plan.axvline(x_end, color='green', ls='--', lw=1.2, alpha=0.7)
    # 各跨交界线
    bx = x_start
    for s in spans_mm[0:len(spans_mm) - 1]:
        bx += s
        tab.ax_plan.axvline(bx, color='blue', ls=':', lw=1.0, alpha=0.6)

    # ==== 标注 ====
    cant_y = beam_top_y + 1.5
    space_y = beam_bottom_y - 1.5  # 横肋纵向布置标注位置
    # 左侧伸出距离（0# 侧）
    if beam_cant_left_m > 0:
        x_left_ext = x_start - beam_cant_left_m
        draw_dim_annotation(tab.ax_plan, x_left_ext, x_start, cant_y,
                             f"{beam_cant_left_m*1000:.0f}", is_highlighted=is_beam_cant_left_hl,
                             line_color=line_color, text_color=TEXT_COLOR, tick_h=tick_h, hl_color=TEXT_HIGHLIGHT)
    # 右侧伸出距离（n# 侧）
    if beam_cant_right_m > 0:
        x_right_ext = x_end + beam_cant_right_m
        draw_dim_annotation(tab.ax_plan, x_end, x_right_ext, cant_y,
                             f"{beam_cant_right_m*1000:.0f}", is_highlighted=is_beam_cant_right_hl,
                             line_color=line_color, text_color=TEXT_COLOR, tick_h=tick_h, hl_color=TEXT_HIGHLIGHT)
    # 横肋纵向布置（使用原始格式标注，保留 N@(pattern) 合并显示）
    if rib_xs:
        try: segs = midasdisttolst_segments(tab.rib_trans_space.get())
        except: segs = merge_spacing_labels(rib_trans_space)
        tab.ax_plan.plot([beam_left, rib_xs[-1]], [space_y, space_y], color=line_color, lw=1.2)
        cur = beam_left
        for idx, (lbl, tot_mm) in enumerate(segs):
            if tot_mm <= 0:
                continue  # 0 段：无宽度，跳过竖线和标签
            tot_m = tot_mm / 1000.0; end = cur + tot_m
            tab.ax_plan.plot([cur, cur], [space_y - tick_h, space_y + tick_h], color=line_color, lw=1.2)
            tc = TEXT_HIGHLIGHT if is_deck_rib_layout_hl else TEXT_COLOR
            weight = "bold" if is_deck_rib_layout_hl else "normal"
            tab.ax_plan.text((cur + end) / 2, space_y + 0.05, lbl, fontsize=8, color=tc, ha='center', va='bottom', weight=weight)
            cur = end
        tab.ax_plan.plot([rib_xs[-1], rib_xs[-1]], [space_y - tick_h, space_y + tick_h], color=line_color, lw=1.2)

    # ==== 桥面板示意（zorder 最底层，20% 透明度灰色）====
    try:
        deck_end_ext_str = tab.deck_end_ext.get()
        ext_parts = deck_end_ext_str.split(",")
        deck_ext_left = float(ext_parts[0].strip()) / 1000.0 if ext_parts[0].strip() else 0.0
        deck_ext_right = float(ext_parts[1].strip()) / 1000.0 if len(ext_parts) > 1 and ext_parts[1].strip() else 0.0
    except:
        deck_ext_left = 0.0; deck_ext_right = 0.0

    # 桥面板基准：钢面板=横肋范围，混凝土面板=纵梁范围
    if is_concrete:
        deck_ref_left = beam_left
        deck_ref_right = beam_right
    else:
        deck_ref_left = min(rib_xs) if rib_xs else beam_left
        deck_ref_right = max(rib_xs) if rib_xs else beam_right

    deck_x_min = deck_ref_left - deck_ext_left
    deck_x_max = deck_ref_right + deck_ext_right
    deck_x_len = deck_x_max - deck_x_min
    deck_y_top = deck_trans_ecc
    deck_y_bot = deck_trans_ecc - bridge_width
    tab.ax_plan.add_patch(patches.Rectangle(
        (deck_x_min, deck_y_top), deck_x_len, deck_y_bot - deck_y_top,
        facecolor='#aab7c4', alpha=0.1, edgecolor='#7f8c8d', linewidth=0.5, zorder=1))
    # 伸出距离标注（正值=向外伸出，负值=向内间距，标注绝对值从参考端向内）
    ann_y = space_y
    if deck_ext_left > 0:
        draw_dim_annotation(tab.ax_plan, deck_x_min, deck_ref_left, ann_y,
                             f"{deck_ext_left*1000:.0f}", is_highlighted=is_deck_ext_left_hl,
                             line_color=line_color, text_color=TEXT_COLOR, tick_h=tick_h,
                             hl_color=TEXT_HIGHLIGHT, lw=1.0, zorder=1)
    elif deck_ext_left < 0:
        draw_dim_annotation(tab.ax_plan, deck_ref_left, deck_x_min, ann_y,
                             f"{abs(deck_ext_left)*1000:.0f}", is_highlighted=is_deck_ext_left_hl,
                             line_color=line_color, text_color=TEXT_COLOR, tick_h=tick_h,
                             hl_color=TEXT_HIGHLIGHT, lw=1.0, zorder=1)
    if deck_ext_right > 0:
        draw_dim_annotation(tab.ax_plan, deck_ref_right, deck_x_max, ann_y,
                             f"{deck_ext_right*1000:.0f}", is_highlighted=is_deck_ext_right_hl,
                             line_color=line_color, text_color=TEXT_COLOR, tick_h=tick_h,
                             hl_color=TEXT_HIGHLIGHT, lw=1.0, zorder=1)
    elif deck_ext_right < 0:
        draw_dim_annotation(tab.ax_plan, deck_x_max, deck_ref_right, ann_y,
                             f"{abs(deck_ext_right)*1000:.0f}", is_highlighted=is_deck_ext_right_hl,
                             line_color=line_color, text_color=TEXT_COLOR, tick_h=tick_h,
                             hl_color=TEXT_HIGHLIGHT, lw=1.0, zorder=1)

    # 视口
    margin = 1.0
    x_all_p = [beam_left, beam_left + beam_len, deck_x_min, deck_x_max]
    y_max = max(cant_y + 0.5, bridge_width /2, beam_ys[0] if beam_ys else 0)
    y_min = min(space_y - 0.5, -bridge_width / 2, beam_ys[-1] if beam_ys else 0)
    tab.ax_plan.set_xlim(min(x_all_p) - margin, max(x_all_p) + margin)
    tab.ax_plan.set_ylim(y_min - margin, y_max + margin)
    tab.ax_plan.set_aspect('equal')
    tab.ax_plan.axis('off')
    tab.fig_plan.tight_layout(pad=0)
    tab.canvas_plan.draw()


# ═══════════════════════════════════════════════════════════════
# MoveLoadTab 加载图 — 完整迁移
# ═══════════════════════════════════════════════════════════════

def draw_move_load_diagram(tab, ax, lt, params, lane_p, hl_key=None):
    """加载示意图 — 横桥向构件 + 车道/车轮 + 标注"""
    geo = {}
    if tab.parent_ui and hasattr(tab.parent_ui, 'tab_components'):
        geo = tab.parent_ui.tab_components.get_veh_diagram_geometry()

    TEXT_COLOR = DRAW_TEXT_COLOR; LINE_COLOR = DRAW_LINE_COLOR; HIGHLIGHT_COLOR = DRAW_TEXT_HIGHLIGHT; TICK_HEIGHT = 0.05

    # ---- 构件几何 ----
    deck_bottom = geo.get("deck_bottom", -0.2)
    deck_top = geo.get("deck_top", 0.0)
    deck_width = geo.get("deck_width", 6.0)
    is_double = geo.get("is_double", False)

    rib_trans_y = geo.get("rib_trans_y", -0.6)
    h_rib_trans = geo.get("h_rib_trans", 0.15)
    rib_trans_width = geo.get("rib_trans_w", 6.0)
    rib_trans_left = geo.get("rib_trans_left", -3.0)
    rib_trans_right = geo.get("rib_trans_right", 3.0)

    rib_long_y = geo.get("rib_long_y")
    h_rib_long = geo.get("h_rib_long")
    b_rib_long = geo.get("b_rib_long")
    rib_positions = geo.get("rib_positions", [])

    beam_y = geo.get("beam_y", -1.5)
    beam_height = geo.get("beam_h", 1.5)
    beam_width = geo.get("beam_w", 0.08)
    beam_positions = geo.get("beam_pos", [])

    distribution_y = geo.get("dist_y", -2.55)
    distribution_height = geo.get("dist_h", 0.4)
    distribution_left = geo.get("dist_left", -4.0)
    distribution_right = geo.get("dist_right", 4.0)

    # ---- 构件示意 ----
    ax.add_patch(patches.Rectangle((rib_trans_left, deck_bottom), deck_width, deck_top - deck_bottom,
                                   facecolor="#8f99a8", edgecolor="none", zorder=2))
    ax.add_patch(patches.Rectangle((rib_trans_left, rib_trans_y), rib_trans_width, h_rib_trans,
                                   facecolor="#4d5a68", edgecolor="none", zorder=1))
    if is_double and rib_long_y is not None:
        for rx in rib_positions:
            ax.add_patch(patches.Rectangle((rx - b_rib_long/2, rib_long_y), b_rib_long, h_rib_long,
                                           facecolor="#4d5a68", edgecolor="none", zorder=3))
    for bx in beam_positions:
        ax.add_patch(patches.Rectangle((bx - beam_width/2, beam_y), beam_width, beam_height,
                                       facecolor="#34495e", edgecolor="none", zorder=4))
    ax.add_patch(patches.Rectangle((distribution_left, distribution_y),
                                   distribution_right - distribution_left, distribution_height,
                                   facecolor="#bdc3c7", edgecolor="none", zorder=5))

    # ---- 车道/车轮 ----
    try: lane_offset = float(lane_p[0])
    except: lane_offset = 0.0
    try: wheel_spacing = float(lane_p[1])
    except: wheel_spacing = 1.8
    try: track_width = float(lane_p[2])
    except:
        track_width = 0.5
        if lt in ["吊装设备", "自定义-吊装设备"] and len(params) >= 6:
            track_width = float(params[5])
        elif lt in ["钻孔设备", "自定义-钻孔设备"] and len(params) >= 5:
            track_width = float(params[4])

    lane_band_height = 0.15
    lane_y_position = deck_top
    half_wheel_spacing = wheel_spacing / 2
    lane_center_x = lane_offset

    for wheel_edge_x in (lane_center_x - half_wheel_spacing, lane_center_x + half_wheel_spacing):
        ax.add_patch(patches.Rectangle((wheel_edge_x - track_width/2, lane_y_position),
                                       track_width, lane_band_height,
                                       facecolor="#0F5F90", edgecolor="#001b47", lw=1, zorder=7, alpha=0.5))

    # ---- 车道标注 ----
    is_wheel_spacing_highlighted = (hl_key == "lane_d")
    is_eccentricity_highlighted = (hl_key == "lane_ecc")
    annotation_base_y = lane_y_position + lane_band_height + 0.15

    ax.plot([lane_center_x - half_wheel_spacing, lane_center_x + half_wheel_spacing],
            [annotation_base_y, annotation_base_y], color=LINE_COLOR, lw=1.2, zorder=10)
    ax.plot([lane_center_x - half_wheel_spacing, lane_center_x - half_wheel_spacing],
            [annotation_base_y - TICK_HEIGHT, annotation_base_y + TICK_HEIGHT], color=LINE_COLOR, lw=1.2, zorder=10)
    ax.plot([lane_center_x + half_wheel_spacing, lane_center_x + half_wheel_spacing],
            [annotation_base_y - TICK_HEIGHT, annotation_base_y + TICK_HEIGHT], color=LINE_COLOR, lw=1.2, zorder=10)
    dim_color = HIGHLIGHT_COLOR if is_wheel_spacing_highlighted else TEXT_COLOR
    ax.text(lane_center_x, annotation_base_y + 0.1, f"轮距:{wheel_spacing}m", fontsize=9, color=dim_color,
            ha="center", va="bottom", weight="bold" if is_wheel_spacing_highlighted else "normal", zorder=10)

    ecc_color = HIGHLIGHT_COLOR if is_eccentricity_highlighted else TEXT_COLOR
    ecc_label_y = annotation_base_y + 0.4
    if lane_center_x > 0:
        ax.plot([0, lane_center_x], [ecc_label_y, ecc_label_y], color=LINE_COLOR, lw=1.2, zorder=10)
        ax.plot([0, 0], [ecc_label_y - TICK_HEIGHT, ecc_label_y + TICK_HEIGHT], color=LINE_COLOR, lw=1.2, zorder=10)
        ax.plot([lane_center_x, lane_center_x], [ecc_label_y - TICK_HEIGHT, ecc_label_y + TICK_HEIGHT], color=LINE_COLOR, lw=1.2, zorder=10)
    elif lane_center_x < 0:
        ax.plot([lane_center_x, 0], [ecc_label_y, ecc_label_y], color=LINE_COLOR, lw=1.2, zorder=10)
        ax.plot([0, 0], [ecc_label_y - TICK_HEIGHT, ecc_label_y + TICK_HEIGHT], color=LINE_COLOR, lw=1.2, zorder=10)
        ax.plot([lane_center_x, lane_center_x], [ecc_label_y - TICK_HEIGHT, ecc_label_y + TICK_HEIGHT], color=LINE_COLOR, lw=1.2, zorder=10)
    ax.text(lane_center_x, ecc_label_y + 0.1, f"偏心:{lane_offset}m", fontsize=8, color=ecc_color,
            ha="center", va="bottom", weight="bold" if is_eccentricity_highlighted else "normal", zorder=10)

    # ---- 荷载复核 ----
    total_load_kN = 0.0
    load_text = ""
    try:
        if lt in ["一般车辆", "自定义-一般车辆"]:
            if len(params) >= 2:
                for ld in params[1].split(","):
                    ld = ld.strip()
                    if ld:
                        total_load_kN += float(ld)
            if total_load_kN > 0:
                load_text = f"当前车辆总重：{total_load_kN:.1f} kN"
        elif lt in ["吊装设备", "自定义-吊装设备", "钻孔设备", "自定义-钻孔设备"]:
            total_load_kN = float(params[1]) if len(params) > 1 and str(params[1]).strip() else 0.0
            if total_load_kN > 0:
                load_text = f"当前设备总重：{total_load_kN:.1f} kN"
        elif lt == "公路标准荷载":
            std_type = params[1] if len(params) > 1 else "未指定"
            load_text = f"公路标准荷载：{std_type}"
    except Exception as e:
        print(f"荷载复核计算异常: {e}")

    load_label_y = ecc_label_y
    if load_text:
        load_label_y = ecc_label_y + 0.5
        ax.text(0, load_label_y, load_text, fontsize=9, ha="center", va="bottom", zorder=10)

    # ---- 中轴线 ----
    y_bottom = distribution_y - distribution_height
    ax.vlines(x=0, ymin=y_bottom - 0.1, ymax=load_label_y - 0.5,
              colors="red", linestyles="--", lw=1, alpha=0.4, zorder=10)
    ax.text(0, y_bottom - 0.15, "线路中心线", fontsize=8, color="red",
            ha="center", va="top", alpha=0.6, zorder=10)
    ax.text(0.98, 0.98, "※ 车道偏心(m)为相对线路中心线偏心距离",
            fontsize=7, color="#888", ha="right", va="top",
            transform=ax.transAxes, zorder=10)

    # ---- 视口（居中对称） ----
    view_margin = 0.5
    x_coords = [rib_trans_left, rib_trans_right, distribution_left, distribution_right] + beam_positions
    x_min = min(x_coords); x_max = max(x_coords)
    lane_extent = abs(lane_offset) + half_wheel_spacing + 0.5
    max_abs_x = max(abs(min(x_min, -lane_extent)), abs(max(x_max, lane_extent)))
    ax.set_xlim(-max_abs_x - view_margin, max_abs_x + view_margin)
    ax.set_ylim(y_bottom - view_margin, load_label_y + view_margin)
    ax.axis("off")


# ═══════════════════════════════════════════════════════════════
# StaticLoadTab 加载图 — 完整迁移
# ═══════════════════════════════════════════════════════════════

def draw_static_load_diagram(tab):
    """重绘动态元素(方块+车辆名)，固定元素已在init中画好"""
    if not hasattr(tab, '_load_ax'):
        return
    ax = tab._load_ax

    # 移除上一次的动态元素
    for elem in tab._load_drawn_elements:
        try:
            elem.remove()
        except Exception:
            pass
    tab._load_drawn_elements = []

    pile_count = tab._pile_count
    if pile_count < 2:
        return

    elems = tab._load_drawn_elements
    center_y = 0.5
    block_h = 0.08
    block_w = 0.15
    offset_y = 0.25

    # 动态图示显示判断
    car_name = tab._current_car_name or ""
    # 1. 检查是否勾选了“工况输出”
    output_on = tab.car_output_check.get(car_name, True) if car_name else False
    # 2. 检查是否有任意“荷载类型”被勾选
    has_active_type = any(v.get() for v in tab.load_type_vars.values())
    # 3. 检查是否有任意“跨间加载”或“桩顶加载”位置被勾选
    has_active_span = any(v.get() for row in tab._span_chk_vars for v in row)
    has_active_pile = any(v.get() for v in tab._pile_chk_vars)
    has_active_pos = has_active_span or has_active_pile
    # 显示判断
    show_vehicle = output_on and has_active_type and has_active_pos

    # 车辆名
    if car_name and show_vehicle:
        elems.append(ax.text(0.02, 0.95, car_name, transform=ax.transAxes,
                fontsize=9, fontweight="bold", color="#2c3e50",
                verticalalignment="top",
                bbox=dict(boxstyle="round,pad=0.3", facecolor="#e3f2fd", edgecolor="#2b6fd4", alpha=0.9)))

    # 加载示意块
    if show_vehicle:
        # 跨间加载方块
        for i, row_vars in enumerate(tab._span_chk_vars):
            if i >= pile_count - 1:
                break
            gap = 0.12
            positions = [i + gap, (i + i + 1) / 2, i + 1 - gap]
            for col_idx, var in enumerate(row_vars):
                if var.get():
                    px = positions[col_idx]
                    elems.append(ax.add_patch(patches.Rectangle(
                        (px - block_w / 2, center_y + offset_y - block_h / 2),
                        block_w, block_h,
                        facecolor="#2b6fd4", edgecolor="#1a4f9e", linewidth=0.5, zorder=5)))
                    elems.append(ax.add_patch(patches.Rectangle(
                        (px - block_w / 2, center_y - offset_y - block_h / 2),
                        block_w, block_h,
                        facecolor="#2b6fd4", edgecolor="#1a4f9e", linewidth=0.5, zorder=5)))

        # 桩顶加载方块 (橙黄色)
        for i, var in enumerate(tab._pile_chk_vars):
            if var.get() and i < pile_count:
                elems.append(ax.add_patch(patches.Rectangle(
                    (i - block_w / 2, center_y + offset_y - block_h / 2),
                    block_w, block_h,
                    facecolor="#FFA500", edgecolor="#CC8400", linewidth=0.5, zorder=5)))
                elems.append(ax.add_patch(patches.Rectangle(
                    (i - block_w / 2, center_y - offset_y - block_h / 2),
                    block_w, block_h,
                    facecolor="#FFA500", edgecolor="#CC8400", linewidth=0.5, zorder=5)))

    tab._load_canvas.draw_idle()


# ═══════════════════════════════════════════════════════════════
# 下部结构设置二级窗口
# ═══════════════════════════════════════════════════════════════

class SupportSetupDialog(tb.Toplevel):
    def __init__(self, parent, idx, current_data, on_select_section=None):
        super().__init__(parent)
        self.title(f"{idx}# 下部结构参数详细设置")

        self.result = None
        self._idx = idx
        self._on_select_section = on_select_section
        self.data = current_data.copy()
        sup_type = self.data.get("type", "")
        
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        geo_w = min(1400, int(sw * 0.85))
        geo_h = min(1000 if sup_type == "制动墩" else 800, int(sh * 0.9))
        geometry_size = f"{geo_w}x{geo_h}"
        self.geometry(f"{geometry_size}+{parent.winfo_rootx()+50}+{parent.winfo_rooty()+50}")
        self.minsize(1200, 600)  # 设置最小尺寸
        self.resizable(True, True)  # 允许调整大小
        self.transient(parent)
        self.grab_set()

        self.vars = {}
        self._sec_widgets = {}
        self._sec_btns = {}
        self._edit_widgets = {}
        self._mat_widgets = {}
        self._last_focused_key = None
        self.TEXT_HIGHLIGHT = DRAW_TEXT_HIGHLIGHT

        self.png_path = getattr(parent, 'png_path', os.path.join(os.getcwd(), "Support", "png_picture"))

        self.setup_ui()

    def _load_conn_form_options(self):
        return ["-", "X", "X|X", "Z"]

    def setup_ui(self):
        # --- 1. 底部按钮区 ---
        bottom_frame = ttk.Frame(self, padding=10)
        bottom_frame.pack(side="bottom", fill="x")
        tb.Button(bottom_frame, text="保存并关闭", bootstyle="success", command=self.on_save).pack(side="right", padx=5)
        tb.Button(bottom_frame, text="取消", bootstyle="secondary", command=self.destroy).pack(side="right", padx=5)

        # --- 2. 滚动区域 ---
        container = ttk.Frame(self)
        container.pack(side="top", fill="both", expand=True)

        self.canvas = tk.Canvas(container, borderwidth=0, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.bind('<Configure>', lambda e: self.canvas.itemconfig(self.canvas_window, width=e.width))
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        def _on_mousewheel(e):
            self.canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        def _on_enter(_):
            self.canvas.bind_all("<MouseWheel>", _on_mousewheel)
        def _on_leave(_):
            self.canvas.unbind_all("<MouseWheel>")
        self.canvas.bind("<Enter>", _on_enter)
        self.canvas.bind("<Leave>", _on_leave)

        self.scrollable_frame.grid_columnconfigure(0, weight=1) 
        self.scrollable_frame.grid_columnconfigure(1, weight=1) 

        # --- 3. 左侧参数模块 ---
        sup_type = self.data.get("type", "")

        # 从 steel_dict 构建材质选项
        steel_dict = trestle_excel_dict.get("steel_dict", {})
        _steel_brands = []
        _steel_spec_id = ""
        for _sd in steel_dict.values():
            if not _steel_spec_id:
                _steel_spec_id = _sd.get("钢结构规范编号", "")
            _steel_brands.extend(_sd.get("牌号参数", {}).keys())
        if not _steel_brands:
            _steel_brands = ["Q235", "Q345"]
        self._steel_brands = _steel_brands
        self._steel_spec_id = _steel_spec_id

        # (1) 桩布置
        lf_pile = ttk.LabelFrame(self.scrollable_frame, text="桩布置", padding=(15, 15, 15, 5))
        lf_pile.grid(row=0, column=0, sticky="nsew", padx=15, pady=(10, 2))
        # self.add_edit_row(lf_pile, "类型:", "pile_type", 0, readonly=True)
        self._add_mat_combo(lf_pile, "桩材质:", "pile_material", 0)
        self._sec_btns["pile_sec"] = create_section_selector(
            lf_pile, "桩截面:", 1, initial_data=self.data.get("pile_sec_detail") or self.data.get("pile_sec", ""),
            families=self.FAM_PIPE, label_width=22, entry_width=30, padx=(0, 5), pady=5,
            on_change=lambda r, k="pile_sec": (self.data.__setitem__(k + "_detail", r), self.data.__setitem__(k, r.get("section_name", "")), self._on_edit_focus(k, False), self.update_diagram_preview(self.data.get("type", ""), idx=self._idx)),
            on_select=lambda r: self._on_select_section and self._on_select_section(r),
            on_focus=lambda btn, k="pile_sec": self._on_edit_focus(k, True))
        self._sec_btns["pile_sec"].bind("<FocusOut>", lambda e, k="pile_sec": (self._on_edit_focus(k, False), self.update_diagram_preview(self.data.get("type", ""), idx=self._idx)))
        self._sec_widgets["pile_sec"] = self._sec_btns["pile_sec"]
        self.add_edit_row(lf_pile, "桩长(m):", "pile_len", 2)
        
        # 桩横向布置：分离高亮
        ttk.Label(lf_pile, text="桩横向布置(mm):", width=22, bootstyle=PRIMARY).grid(row=3, column=0, sticky="w", padx=(0, 5), pady=5)
        def _on_pts_diagram(side, is_focus):
            self._on_edit_focus(f"pile_trans_space_{side}", is_focus)
        sp_trans = SplitEntry(lf_pile, 3, default_val=self.data.get("pile_trans_space", ""), entry_width=29, on_focus=_on_pts_diagram)
        self._edit_widgets["pile_trans_space"] = sp_trans
        sp_trans.bind("<FocusOut>", lambda e: (self.data.__setitem__("pile_trans_space", sp_trans.get()), self.update_diagram_preview(self.data.get("type", ""), idx=self._idx)))

        if sup_type == "制动墩":
            ttk.Label(lf_pile, text="桩纵向布置(mm):", width=22, bootstyle=PRIMARY).grid(row=4, column=0, sticky="w", padx=(0, 5), pady=5)
            def _on_pls_diagram(side, is_focus):
                self._on_edit_focus(f"pile_long_space_{side}", is_focus)
            sp_long = SplitEntry(lf_pile, 4, default_val=self.data.get("pile_long_space", ""), entry_width=29, on_focus=_on_pls_diagram)
            self._edit_widgets["pile_long_space"] = sp_long
            sp_long.bind("<FocusOut>", lambda e: (self.data.__setitem__("pile_long_space", sp_long.get()), self.update_diagram_preview(self.data.get("type", ""), idx=self._idx)))
            self.add_edit_row(lf_pile, "伸缩缝长度(mm):", "expansion_len", 5, readonly=True)

        # (2) 分配梁定义
        lf_dist = ttk.LabelFrame(self.scrollable_frame, text="分配梁定义", padding=(15, 15, 15, 5))
        lf_dist.grid(row=1, column=0, sticky="nsew", padx=15, pady=2)
        self._add_mat_combo(lf_dist, "分配梁(横)材质:", "dist_trans_material", 0)
        self._sec_btns["dist_trans_sec"] = create_section_selector(
            lf_dist, "分配梁(横)截面:", 1, initial_data=self.data.get("dist_trans_sec_detail") or self.data.get("dist_trans_sec", ""),
            label_width=22, entry_width=28, padx=(0, 5), pady=5,
            on_change=lambda r, k="dist_trans_sec": (self.data.__setitem__(k + "_detail", r), self.data.__setitem__(k, r.get("section_name", "")), self._on_edit_focus(k, False), self.update_diagram_preview(self.data.get("type", ""), idx=self._idx)),
            on_select=lambda r: self._on_select_section and self._on_select_section(r),
            on_focus=lambda btn, k="dist_trans_sec": self._on_edit_focus(k, True))
        self._sec_btns["dist_trans_sec"].bind("<FocusOut>", lambda e, k="dist_trans_sec": (self._on_edit_focus(k, False), self.update_diagram_preview(self.data.get("type", ""), idx=self._idx)))
        self._sec_widgets["dist_trans_sec"] = self._sec_btns["dist_trans_sec"]
        self.add_edit_row(lf_dist, "分配梁(横)长(mm):", "dist_trans_len", 2)
        self.add_edit_row(lf_dist, "分配梁(横)左端距中心线(mm):", "dist_trans_offset", 3)
        
        if sup_type == "制动墩":
            self._add_mat_combo(lf_dist, "分配梁(纵)材质:", "dist_long_material", 4)
            self._sec_btns["dist_long_sec"] = create_section_selector(
                lf_dist, "分配梁(纵)截面:", 5, initial_data=self.data.get("dist_long_sec_detail") or self.data.get("dist_long_sec", ""),
                label_width=22, entry_width=30, padx=(0, 5), pady=5,
                on_change=lambda r, k="dist_long_sec": (self.data.__setitem__(k + "_detail", r), self.data.__setitem__(k, r.get("section_name", "")), self._on_edit_focus(k, False), self.update_diagram_preview(self.data.get("type", ""), idx=self._idx)),
                on_select=lambda r: self._on_select_section and self._on_select_section(r),
                on_focus=lambda btn, k="dist_long_sec": self._on_edit_focus(k, True))
            self._sec_btns["dist_long_sec"].bind("<FocusOut>", lambda e, k="dist_long_sec": (self._on_edit_focus(k, False), self.update_diagram_preview(self.data.get("type", ""), idx=self._idx)))
            self._sec_widgets["dist_long_sec"] = self._sec_btns["dist_long_sec"]
            self.add_edit_row(lf_dist, "分配梁(纵)长(mm):", "dist_long_len", 6)
            self.add_edit_row(lf_dist, "分配梁(纵)左端距中心线(mm):", "dist_long_center_offset", 7)
            ttk.Label(lf_dist, text="  ※ 联中心线在两联伸缩缝中心处", font=("微软雅黑", 8), foreground="#888").grid(row=8, column=0, columnspan=2, sticky="w", pady=(0, 3))

        # (3) 联结系定义
        lf_brace = ttk.LabelFrame(self.scrollable_frame, text="联结系定义", padding=(15, 15, 15, 5))
        lf_brace.grid(row=2, column=0, sticky="nsew", padx=15, pady=(2, 10))

        self.vars["has_brace"] = tk.BooleanVar(value=self.data.get("brace_form") not in ("", "/"))
        ttk.Label(lf_brace, text="设置联结系:", width=22, bootstyle=PRIMARY).grid(row=0, column=0, sticky="w", padx=(0,5), pady=5)
        tb.Checkbutton(lf_brace, variable=self.vars["has_brace"], bootstyle="round-toggle", command=self._toggle_brace).grid(row=0, column=1, sticky="w", padx=(0,5), pady=5)

        conn_options = self._load_conn_form_options()
        self.brace_form_var = tk.StringVar(value=self.data.get("brace_form", conn_options[0] if conn_options else "-"))
        ttk.Label(lf_brace, text="联结系形式:", width=22, bootstyle=PRIMARY).grid(row=1, column=0, sticky="w", padx=(0,5), pady=5)
        self.brace_form_combo = ttk.Combobox(lf_brace, textvariable=self.brace_form_var, values=conn_options, state="readonly", width=18)
        self.brace_form_combo.grid(row=1, column=1, sticky="w", padx=(0,5), pady=5)
        self.brace_form_combo.bind("<<ComboboxSelected>>", lambda e: (self.data.__setitem__("brace_form", self.brace_form_var.get()), self.update_diagram_preview(self.data.get("type", ""), idx=self._idx, highlight_key="brace_space")))

        self._add_mat_combo(lf_brace, "联结系材质:", "brace_material", 2)

        self._sec_btns["brace_sec"] = create_section_selector(
            lf_brace, "联结系截面:", 3, initial_data=self.data.get("brace_sec_detail") or self.data.get("brace_sec", ""),
            families=self.FAM_ALL, label_width=22, entry_width=30, padx=(0, 5), pady=5,
            on_change=lambda r, k="brace_sec": (self.data.__setitem__(k + "_detail", r), self.data.__setitem__(k, r.get("section_name", "")), self._on_edit_focus(k, False), self.update_diagram_preview(self.data.get("type", ""), idx=self._idx)),
            on_select=lambda r: self._on_select_section and self._on_select_section(r),
            on_focus=lambda btn, k="brace_sec": self._on_edit_focus(k, True))
        self._sec_btns["brace_sec"].bind("<FocusOut>", lambda e, k="brace_sec": (self._on_edit_focus(k, False), self.update_diagram_preview(self.data.get("type", ""), idx=self._idx)))
        self._sec_widgets["brace_sec"] = self._sec_btns["brace_sec"]

        self._sec_btns["brace_sec_tilt"] = create_section_selector(
            lf_brace, "联结系斜杆/竖杆:", 4, initial_data=self.data.get("brace_sec_tilt_detail") or self.data.get("brace_sec_tilt", ""),
            families=self.FAM_ALL, label_width=22, entry_width=30, padx=(0, 5), pady=5,
            on_change=lambda r, k="brace_sec_tilt": (self.data.__setitem__(k + "_detail", r), self.data.__setitem__(k, r.get("section_name", "")), self._on_edit_focus(k, False), self.update_diagram_preview(self.data.get("type", ""), idx=self._idx)),
            on_select=lambda r: self._on_select_section and self._on_select_section(r),
            on_focus=lambda btn, k="brace_sec_tilt": self._on_edit_focus(k, True))
        self._sec_btns["brace_sec_tilt"].bind("<FocusOut>", lambda e, k="brace_sec_tilt": (self._on_edit_focus(k, False), self.update_diagram_preview(self.data.get("type", ""), idx=self._idx)))
        self._sec_widgets["brace_sec_tilt"] = self._sec_btns["brace_sec_tilt"]

        # 间距：分离高亮
        ttk.Label(lf_brace, text="间距(mm):", width=22, bootstyle=PRIMARY).grid(row=5, column=0, sticky="w", padx=(0, 5), pady=5)
        def _on_bs_diagram(side, is_focus):
            self._on_edit_focus(f"brace_space_{side}", is_focus)
        sp_brace = SplitEntry(lf_brace, 5, default_val=self.data.get("brace_space", ""), entry_width=29, on_focus=_on_bs_diagram)
        self._edit_widgets["brace_space"] = sp_brace
        sp_brace.bind("<FocusOut>", lambda e: (self.data.__setitem__("brace_space", sp_brace.get()), self.update_diagram_preview(self.data.get("type", ""), idx=self._idx)))

        self.add_edit_row(lf_brace, "层高(mm):", "brace_height", 6)

        def _on_brace_form_change(*_):
            brace_on = self.vars["has_brace"].get()
            is_single = (self.brace_form_var.get() == "-")
            # 联结系形式为 "-" 时，斜杆/竖杆截面和层高应禁用
            tilt_state = "disabled" if (not brace_on or is_single) else "normal"
            if h_ent := self._edit_widgets.get("brace_height"): h_ent.configure(state=tilt_state)
            if w := self._sec_widgets.get("brace_sec_tilt"): w.configure(state=tilt_state)
        self._on_brace_form_change = _on_brace_form_change
        self.brace_form_var.trace_add("write", _on_brace_form_change)

        # --- 4. 右侧示意图区 ---
        self.preview_frame = ttk.LabelFrame(self.scrollable_frame, text="下部结构构造图示", padding=10)
        self.preview_frame.grid(row=0, column=1, rowspan=3, sticky="nsew", padx=(0, 20), pady=10)

        if sup_type == "制动墩":
            self.preview_fig = Figure(figsize=(3.8, 6.5), dpi=90)
            self.preview_ax_top = self.preview_fig.add_subplot(211)
            self.preview_ax_bottom = self.preview_fig.add_subplot(212)
            self.preview_ax = self.preview_ax_top
        else:
            self.preview_fig = Figure(figsize=(3.8, 6), dpi=90)
            self.preview_ax = self.preview_fig.add_subplot(111)
            
        self.preview_canvas = FigureCanvasTkAgg(self.preview_fig, master=self.preview_frame)
        self.preview_canvas.get_tk_widget().pack(expand=True, fill="both")

        self._toggle_brace()
        self.update_diagram_preview(sup_type, idx=self._idx)

    # ==================== 构建精简行组件 ====================
    def add_edit_row(self, parent, label_text, key, row, readonly=False):
        ttk.Label(parent, text=label_text, width=22, bootstyle=PRIMARY).grid(row=row, column=0, sticky="w", padx=(0, 5), pady=5)
        var = tk.StringVar(value=self.data.get(key, ""))
        self.vars[key] = var
        var.trace_add("write", lambda *_: self.data.__setitem__(key, var.get()))
        
        ent = ttk.Entry(parent, textvariable=var, width=30, state="readonly" if readonly else "normal")
        ent.grid(row=row, column=1, sticky="w", padx=(0, 5), pady=5)
        self._edit_widgets[key] = ent
        
        ent.bind("<FocusIn>", lambda e: self._on_edit_focus(key, True))
        ent.bind("<FocusOut>", lambda e: (self.data.__setitem__(key, var.get()), self._on_edit_focus(key, False), self.update_diagram_preview(self.data.get("type", ""), idx=self._idx)))

    FAM_PIPE = ["O"]                                                 
    FAM_BEAM = ["I","2I","HM","2HM","HN","2HN","HW","2HW","C","2C"]  
    FAM_ALL  = ["I","2I","HM","2HM","HN","2HN","HW","2HW","C","2C","∠","O"]  
    SEC_FAMILIES = {"pile_sec": FAM_PIPE, "dist_trans_sec": FAM_BEAM, "dist_long_sec": FAM_BEAM, "brace_sec": FAM_ALL, "brace_sec_tilt": FAM_ALL}

    # ==================== 视图更新与截面逻辑 ====================

    def update_diagram_preview(self, sup_type, idx=0, highlight_key=None):
        if sup_type == "制动墩":
            self.preview_ax_top.clear()
            self.preview_ax_bottom.clear()
            draw_support_transverse_view(self.preview_ax_top, self.data, idx, highlight_key, sup_type)
            draw_support_longitudinal_view(self.preview_ax_bottom, self.data, idx, highlight_key)
            self.preview_fig.tight_layout(pad=0)
            self.preview_fig.subplots_adjust(hspace=0.4)
            self.preview_fig.add_artist(mlines.Line2D([0, 1], [0.5, 0.5], transform=self.preview_fig.transFigure, color="#2c3e50", lw=0.5, clip_on=False))
        else:
            self.preview_ax.clear()
            draw_support_transverse_view(self.preview_ax, self.data, idx, highlight_key, sup_type)
            self.preview_fig.tight_layout(pad=0)
        self.preview_canvas.draw_idle()

    def _map_highlight(self, key):
        hl_map = {
            "pile_trans_space": "pile_trans_space", "pile_long_space": "pile_long_space",
            "pile_type": "pile", "pile_sec": "pile", "pile_len": "pile",
            "dist_trans_sec": "dist_trans", "dist_trans_len": "dist_trans_len", "dist_trans_offset": "dist_trans_cant",
            "dist_long_sec": "dist_long", "dist_long_len": "dist_long_len", "dist_long_center_offset": "dist_long_center_offset",
            "brace_sec": "brace", "brace_sec_tilt": "brace", "brace_space": "brace_space", "brace_height": "brace_height"
        }
        # split keys: 所有分割字段保留 _left/_right 用于分离高亮
        if key in hl_map:
            return hl_map[key]
        for prefix in ("pile_trans_space", "pile_long_space", "brace_space", "rib_trans_space", "rib_long_space"):
            if key.startswith(prefix):
                return key  # 保留完整 key（如 pile_trans_space_left）
        return None

    def _on_edit_focus(self, key, is_focus):
        self._last_focused_key = key if is_focus else None
        self.update_diagram_preview(self.data.get("type",""), idx=self._idx, highlight_key=self._map_highlight(key) if is_focus else None)

    def _toggle_brace(self):
          enabled = self.vars["has_brace"].get()
          self.data["has_brace"] = enabled
          # 主动更新联结系形式：不设置→"/"，设置→恢复默认（"-" 或 "X"）
          if not enabled:
              self.brace_form_var.set("/")
              self.data["brace_form"] = "/"
          elif self.brace_form_var.get() == "/":
              self.brace_form_var.set("-")
              self.data["brace_form"] = "-"
          state = "normal" if enabled else "disabled"
          combo_state = "readonly" if enabled else "disabled"
          # 联结系形式
          if hasattr(self, 'brace_form_combo'):
              self.brace_form_combo.configure(state=combo_state)
          # 材质
          if w := self._mat_widgets.get("brace_material"):
              w.configure(state=combo_state)
          # 截面
          for key in ("brace_sec", "brace_sec_tilt"):
              if w := self._sec_widgets.get(key): w.configure(state=state)
          # 间距、层高
          for key in ("brace_space", "brace_height"):
              if w := self._edit_widgets.get(key):
                  if hasattr(w, 'configure'):
                      w.configure(state=state)
          # 重新应用 brace_form 对斜杆/竖杆、层高的控制（checkbox 切换时 trace 不触发）
          if hasattr(self, '_on_brace_form_change'):
              self._on_brace_form_change()
          self.update_diagram_preview(self.data.get("type",""), idx=self._idx)

    def _add_mat_combo(self, parent, label, key, row):
        """在 parent frame 中添加一行材质 Combobox"""
        ttk.Label(parent, text=label, width=22, bootstyle=PRIMARY).grid(
            row=row, column=0, sticky="w", padx=(0, 5), pady=5)
        combo = ttk.Combobox(parent, values=self._steel_brands, state="readonly", width=18)
        combo.grid(row=row, column=1, sticky="w", padx=(0, 5), pady=5)
        raw = self.data.get(key, "")
        _, brand = _parse_material_value(raw)
        if brand and brand in self._steel_brands:
            combo.set(brand)
        elif self._steel_brands:
            combo.set(self._steel_brands[0])
            # 默认情况下也写入数据，保证保存时正常写入材质
            sid = self._steel_spec_id
            self.data[key] = f"{sid},{self._steel_brands[0]}" if sid and self._steel_brands[0] else self._steel_brands[0]
        def _on_select(e, k=key, c=combo):
            brand = c.get()
            sid = self._steel_spec_id
            self.data[k] = f"{sid},{brand}" if sid and brand else brand
        combo.bind("<<ComboboxSelected>>", _on_select)
        self._mat_widgets[key] = combo
        return combo

    def on_save(self):
        # 收集所有变量值
        for k, var in self.vars.items(): self.data[k] = var.get()
        for k, btn in self._sec_btns.items():
            if sd := btn.get_section_data():
                self.data[k] = sd.get("section_name", "")
                self.data[k + "_detail"] = sd
        if hasattr(self, 'brace_form_var'): self.data["brace_form"] = self.brace_form_var.get()
        # 联结系未启用时，清空所有联结系字段
        if not self.data.get("has_brace", True):
            for k in ("brace_form", "brace_sec", "brace_sec_tilt", "brace_space", "brace_height", "brace_material"):
                self.data[k] = "/"
        # 联结系形式为 "-" 时，斜杆/竖杆截面强制设为 "/"
        elif self.data.get("brace_form") == "-":
            self.data["brace_sec_tilt"] = "/"

        # 统一联结系间距数据结构：格式 "{第一项},{第二项}"; 第二项为空时统一为 "/"
        if self.data.get("has_brace", True):
            bs = str(self.data.get("brace_space", "")).strip()
            if "," in bs:
                v1, v2 = bs.split(",", 1)
                self.data["brace_space"] = f"{v1.strip()},{v2.strip()}" if v2.strip() else f"{v1.strip()},/"
            elif bs:
                self.data["brace_space"] = f"{bs},/"
            else:
                self.data["brace_space"] = ",/"

        # ---- 验证 ----
        errors = []
        has_brace = self.data.get("has_brace", True)
        # 数值字段
        for key, label in [("pile_len", "桩长"), ("dist_trans_offset", "分配梁(横)左端距中心线"),
                            ("dist_trans_len", "分配梁(横)长"), ("dist_long_len", "分配梁(纵)长"),
                            ("dist_long_center_offset", "分配梁(纵)左端距中心线")]:
            val = str(self.data.get(key, "")).strip()
            if val and val != "/":
                try: float(val)
                except (ValueError, TypeError): errors.append(f"{label} 输入错误: {val}")
        # SplitEntry：桩横向/纵向布置、联结系间距/层高
        # e1（第一项）= 数值；e2（第二项）= 间距格式（支持 2@3000 等）
        for key, label, allow_slash in [("pile_trans_space", "桩横向布置", False),
                                         ("pile_long_space", "桩纵向布置", False),
                                         ("brace_space", "联结系间距", True),
                                         ("brace_height", "联结系层高", False)]:
            # 未设置联结系时跳过联结系相关验证
            if not has_brace and key.startswith("brace"):
                continue
            sp = self._edit_widgets.get(key)
            if not sp or not hasattr(sp, 'e1'):
                continue
            from General.DataUtils import midasdisttolst
            for i, ent in enumerate([sp.e1, sp.e2]):
                val = ent.get().strip()
                if not val:
                    continue
                if allow_slash and i == 1 and val == "/":
                    continue
                try:
                    if i == 0:
                        # 第一项：纯数值
                        float(val)
                    else:
                        # 第二项：间距格式（支持括号/2@3000/750,68@750 等）
                        vals = midasdisttolst(val)
                        if not vals:
                            errors.append(f"{label} 第{i+1}项输入错误: {val}")
                except Exception:
                    errors.append(f"{label} 第{i+1}项输入错误: {val}")
        # 截面完整性
        brace_form = self.data.get("brace_form", "")
        for key, label in [("pile_sec", "桩截面"), ("dist_trans_sec", "分配梁(横)截面"),
                            ("dist_long_sec", "分配梁(纵)截面"), ("brace_sec", "联结系截面"),
                            ("brace_sec_tilt", "联结系斜杆截面")]:
            # 未设置联结系时跳过联结系相关截面验证
            if not has_brace and key.startswith("brace"):
                continue
            # brace_form == "-" 时，斜杆截面强制为 "/"，跳过验证
            if key == "brace_sec_tilt" and brace_form == "-":
                continue
            sd = self._sec_btns.get(key)
            if sd:
                sec_data = sd.get_section_data()
                if sec_data:
                    params = sec_data.get("params", {})
                    if not params:
                        errors.append(f"{label} 截面参数为空")
                    else:
                        for pname, pval in params.items():
                            if not pval or not str(pval).strip():
                                errors.append(f"{label} 参数 {pname} 输入值不能为空")
                            else:
                                try: float(pval)
                                except (ValueError, TypeError): errors.append(f"{label} 参数 {pname} 输入错误: {pval}")
        if errors:
            from tkinter import messagebox
            messagebox.showwarning("参数错误", "\n".join(errors), parent=self)
            return
        self.result = self.data
        self.destroy()
