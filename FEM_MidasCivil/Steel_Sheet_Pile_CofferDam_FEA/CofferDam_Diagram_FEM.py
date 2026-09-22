# =============================================================================
# 术语表（全文件统一约定）
#   SKGGZ  : 锁扣钢管桩
#   Waler  : 围檩；Strut: 内支撑（对撑 DC / 斜撑 XC）
#   Cap    : 承台；Corbel/Bracket: 牛腿
#   rtd    : RealTimeDiagram 实例，实时联动立面/平面示意图
#   高亮组 : 控件聚焦时对应的图示元素分组名（如 "waler_0_dc"）
#   单位约定 : 输入参数多为 m；截面 SEC 为 mm，绘制时 /1000 转为 m；
#              立面图 Y 方向有 1.5 倍视觉放大（v_scale），仅用于显示
# =============================================================================

# 1. 标准库
import re
import sys
import math
from pathlib import Path
import tkinter as tk
from tkinter import ttk

# 2. 第三方库
import matplotlib
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.patches import Polygon, Rectangle, Circle

sys.path.append(str(Path(__file__).parent.parent.parent))
from General.Matplotlib import draw_elevation_marker

matplotlib.use('TkAgg')
matplotlib.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei']
matplotlib.rcParams['axes.unicode_minus'] = False


class RealTimeDiagram:
    """围堰立面/平面实时联动示意图。

    左侧为长边立面图（土层、水位、支护桩、承台、垫层/封底、围檩、内支撑、
    中心线、标高与尺寸标注），右侧为平面图（承台、支护桩、围檩、对撑/斜撑、
    尺寸标注）。控件聚焦时对应元素高亮，参数变化时自动重算重绘。

    Attributes:
        var_dict (dict): 所有 tkinter Variable 引用 {变量名: tk.Variable}。
        soil_load_dict (dict): 土层信息，含 'Excel' 键。
        waler_rows (list[dict]): 各层围檩控件行。
        strut_blocks (dict): 内支撑布置控件 {层号: {'DC_X': tkvar, ...}}。
        waler_section_dict (dict): 围檩截面库。
        strut_section_dict (dict): 内支撑截面库。
        concrete_blinding_check, concrete_plug_check (tk.BooleanVar): 垫层/封底开关。
        g (dict): _calc_geometry 计算的绘图坐标缓存。
    """

    # 配色常量
    PILE_COLOR = "#4a5056"
    PILE_HIGHLIGHT = "#dc3243"
    CAP_FILL = "#f5f6f8"
    WALER_COLOR = "#4a5056"
    STRUT_COLOR = "#4a5056"
    CONCRETE_COLOR = "#777d83"
    SOIL_SURFACE_COLOR = "#4a5056"
    SOIL_LAYER_COLORS = ["#e8e9eb", "#dde0e3", "#d0d4d9", "#c5c9cf", "#bcc0c6"]
    SOIL_HATCHES = ["//", "\\\\", "xx", "++", ".."]
    WATER_COLOR = "#5dade2"
    CENTERLINE_COLOR = "#dc3243"
    TEXT_COLOR = "#2c3e50"
    DIM_COLOR = "#2c3e50"
    LEVEL_LINE_COLOR = "#2c3e50"

    def __init__(self, parent, var_dict, soil_load_dict, waler_rows,
                 concrete_blinding_check, concrete_blinding_thickness,
                 concrete_plug_check, concrete_plug_thickness,
                 strut_blocks=None, waler_section_dict=None, strut_section_dict=None):
        """初始化示意图：创建双轴画布、绑定容器缩放事件并首次绘制。

        Args:
            parent (tk.Widget): 承载画布的父容器。
            var_dict (dict): 所有 tkinter Variable 引用 {变量名: tk.Variable}。
            soil_load_dict (dict): 土层信息（含 'Excel' 键）。
            waler_rows (list[dict]): 各层围檩控件行。
            concrete_blinding_check, concrete_blinding_thickness: 垫层开关/厚度控件。
            concrete_plug_check, concrete_plug_thickness: 封底开关/厚度控件。
            strut_blocks (dict|None): 内支撑布置控件。
            waler_section_dict (dict|None): 围檩截面库。
            strut_section_dict (dict|None): 内支撑截面库。

        Returns:
            None
        """
        self.parent = parent
        self.var_dict = var_dict
        self.soil_load_dict = soil_load_dict
        self.waler_rows = waler_rows
        self.strut_blocks = strut_blocks if strut_blocks is not None else {}
        self.waler_section_dict = waler_section_dict if waler_section_dict is not None else {}
        self.strut_section_dict = strut_section_dict if strut_section_dict is not None else {}
        self.concrete_blinding_check = concrete_blinding_check
        self.concrete_blinding_thickness = concrete_blinding_thickness
        self.concrete_plug_check = concrete_plug_check
        self.concrete_plug_thickness = concrete_plug_thickness

        self.cached = {}
        self.need_recalc = True
        self.highlight_groups = {}
        self.widget_to_group = {}
        self._last_focused_group = ""

        self.fig, (self.ax, self.ax_plan) = plt.subplots(
            1, 2, figsize=(7.5, 3.2), dpi=100,
            facecolor='white',
            gridspec_kw={'width_ratios': [6, 4]})
        self.ax_plan.set_aspect('equal')

        self.canvas = FigureCanvasTkAgg(self.fig, master=parent)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        self.fig.subplots_adjust(left=0.02, right=0.98, top=0.92, bottom=0.04, wspace=0.12)

        # 容器尺寸变化时自动适配 figsize
        self._resize_job = None
        parent.bind("<Configure>", self._on_parent_resize)

        self.draw()

    def _on_parent_resize(self, event):
        """容器尺寸变化时自适应 figsize，防止图示重叠"""
        w, h = event.width, event.height
        if w < 100 or h < 80:
            return
        if self._resize_job is not None:
            self.parent.after_cancel(self._resize_job)
        def _apply():
            """按容器像素尺寸更新 figure 尺寸并重绘（防抖后执行）。"""
            dpi = self.fig.get_dpi()
            self.fig.set_size_inches(w / dpi, h / dpi)
            self.need_recalc = True
            self.draw()
            self._resize_job = None
        self._resize_job = self.parent.after(100, _apply)

    # ========== 控件注册 ==========

    def register_group(self, group_name, widgets):
        """注册一组控件到指定高亮组，并绑定焦点事件。

        Args:
            group_name (str): 高亮组名（如 "waler_0_dc"）。
            widgets (tk.Widget|list[tk.Widget]): 单个控件或控件列表。

        Returns:
            None
        """
        if not isinstance(widgets, (list, tuple)):
            widgets = [widgets]
        self.highlight_groups[group_name] = list(widgets)
        for w in widgets:
            self._bind_widget(w, group_name)

    def _bind_widget(self, w, group_name):
        """把单个控件的焦点/点击事件绑定到高亮组。

        Args:
            w (tk.Widget): 控件对象。
            group_name (str): 高亮组名。

        Returns:
            None
        """
        self.widget_to_group[str(w)] = group_name
        w.bind("<FocusIn>", lambda e, g=group_name: self._on_focus_in(g))
        w.bind("<FocusOut>", self._on_focus_out)
        if isinstance(w, ttk.Combobox) or "combobox" in str(w).lower():
            # 鼠标刚一点击（下拉框展开前）立刻触发高亮
            w.bind("<Button-1>", lambda e, g=group_name: self._on_focus_in(g))
            # 选完内容后也确保维持高亮
            w.bind("<<ComboboxSelected>>", lambda e, g=group_name: self._on_focus_in(g))

    def rebind_waler_widgets(self):
        """重新绑定围檩控件 (add/del row 后调用)"""
        # 清除旧的 waler 绑定组
        keys_to_del = [k for k in self.highlight_groups.keys() if k.startswith("waler_")]
        for k in keys_to_del:
            del self.highlight_groups[k]
        # 重新绑定新的 waler 绑定组
        for i, row in enumerate(self.waler_rows):
            # 间距 -> 高亮围檩+内支撑
            if "entry_w" in row and row["entry_w"].winfo_exists():
                self.register_group(f"waler_{i}_spacing", row["entry_w"])
            # 围檩长边截面 -> 高亮围檩(立面+平面长边)
            if "combo1_w" in row and row["combo1_w"].winfo_exists():
                self.register_group(f"waler_{i}_long", row["combo1_w"])
            # 围檩短边截面 -> 高亮围檩(立面+平面短边)
            if "combo_short_w" in row and row["combo_short_w"].winfo_exists():
                self.register_group(f"waler_{i}_short", row["combo_short_w"])
            # 对撑截面 → 仅高亮对撑
            if "combo2_w" in row and row["combo2_w"].winfo_exists():
                self.register_group(f"waler_{i}_dc", row["combo2_w"])
            # 斜撑截面 → 仅高亮斜撑
            if "combo3_w" in row and row["combo3_w"].winfo_exists():
                self.register_group(f"waler_{i}_xc", row["combo3_w"])
            # 按钮 → 两者都高亮
            if "btn_w" in row and row["btn_w"].winfo_exists():
                self.register_group(f"waler_{i}_btn", row["btn_w"])

    def _on_focus_in(self, group):
        """控件获得焦点：记录当前高亮组并仅重绘高亮。

        Args:
            group (str): 高亮组名。

        Returns:
            None
        """
        self._last_focused_group = group
        self.draw(only_highlight=True)

    def _on_focus_out(self, event=None):
        """控件失去焦点：延迟确认焦点未转移到同组控件后清除高亮。

        Args:
            event: Tk 事件对象（可选）。

        Returns:
            None
        """
        def do_clear():
            """确认焦点确实离开当前组后清空高亮组。"""
            try:
                focused = self.parent.focus_get()
            except (KeyError, tk.TclError):
                focused = None
            if focused and str(focused) in self.widget_to_group:
                if self.widget_to_group[str(focused)] == getattr(self, '_last_focused_group', ''):
                    return
        self._last_focused_group = ""
        self.need_recalc = True
        self.parent.after(10, do_clear)

    # ========== 数据读取 ==========

    def _v(self, name):
        """读取 var_dict 中的 DoubleVar/StringVar 值并转为 float, 容错"""
        v = self.var_dict.get(name)
        if v is None:
            return 0.0
        try:
            return float(v.get())
        except (ValueError, tk.TclError):
            return 0.0

    def _sf(self, value, default=0.0):
        """土层表数值转 float, 空串/None/'/'等非法值回退 default"""
        if value is None:
            return default
        try:
            s = str(value).strip()
            return default if s in ("", "/") else float(s)
        except (ValueError, TypeError):
            return default

    def _read_all(self):
        """从 var_dict 及各开关读取全部输入参数，缓存到 self 供几何计算使用。

        读取内容：土顶/水位/承台/围堰几何、SKGGZ 参数、垫层/封底开关与厚度、
        以及土层分层列表。

        Returns:
            None
        """
        self.solid_level = self._v("Solid_Level")
        self.water_level = self._v("Water_Level")
        self.cap_x = self._v("Cap_X")
        self.cap_y = self._v("Cap_Y")
        self.cap_bottom = self._v("Cap_Bottom_Level")
        self.x_offset = self._v("X_offset")
        self.y_offset = self._v("Y_offset")
        self.cap_h = self._v("Cap_H")
        self.pile_top = self._v("CofferDam_Top_Level")
        self.pile_length = self._v("CofferDam_L")
        self.pile_bottom = self.pile_top - self.pile_length
        self.pile_type = self.var_dict.get("Sheet_Pile_namevar")
        if self.pile_type:
            self.pile_type = self.pile_type.get()

        # SKGGZ 模式检测
        _st = self.var_dict.get("Sheet_Pile_typevar")
        self.is_skggz = _st is not None and (hasattr(_st, 'get') and _st.get() == "锁扣钢管桩")
        if self.is_skggz:
            self.skggz_D = self._v("SKGGZ_D")
            self.skggz_t = self._v("SKGGZ_t")
            self.skggz_Gap = self._v("SKGGZ_Gap")
        else:
            self.skggz_D = self.skggz_t = self.skggz_Gap = 0.0

        # 垫层／封底 (垫层 cm→m)
        self.has_blinding = self.concrete_blinding_check.get()
        self.blinding_thickness = (self._v("Concrete_Blinding_thickness_var") if self.has_blinding else 0.0)
        self.has_plug = self.concrete_plug_check.get()
        self.plug_thickness = (self._v("Concrete_Plug_thickness_var")
                               if self.has_plug else 0.0)

        # 土层
        soil_excel = self.soil_load_dict.get('Excel', {})
        self.soil_layers = []
        for i in range(1, 30):
            key = f"第{i}层土"
            if key in soil_excel:
                d = soil_excel[key]
                self.soil_layers.append({
                    'name': d.get('土层名称', ''),
                    'thickness': self._sf(d.get('层厚')),
                    'gamma': self._sf(d.get('重度')),
                    'cohesion': self._sf(d.get('黏聚力')),
                    'friction': self._sf(d.get('内摩擦角')),
                })
            else:
                break

    # ========== 几何计算 ==========

    def _calc_geometry(self):
        """由输入参数计算所有绘图坐标 (单位: m, 原点在地面中心)"""
        # X 方向关键位置 — 根据桩类型分别计算
        cap_half = self.cap_x / 2.0
        if self.is_skggz:
            D_m = self.skggz_D / 1000.0
            pile_half_w = D_m / 2.0
            pile_center = cap_half + self.x_offset + pile_half_w
        else:
            pile_half_w = 0.25
            pile_center = cap_half + self.x_offset
        left_outer = -pile_center - pile_half_w
        left_inner = -pile_center + pile_half_w
        left_anno = left_outer - 10

        self.g = {
            'pile_left_inner': left_inner,
            'pile_left_outer': left_outer,
            'pile_right_inner': pile_center - pile_half_w,
            'pile_right_outer': pile_center + pile_half_w,
            'cap_left': -cap_half,
            'cap_right': cap_half,
            'pile_center': pile_center,
            'left_anno': left_anno,
            'right_anno': pile_center + pile_half_w + 1.5,
            'pile_half_w': pile_half_w,
        }

        # Y 方向关键标高
        self.g['pile_top_y'] = self.pile_top
        self.g['pile_bottom_y'] = self.pile_bottom
        self.g['ground_y'] = self.solid_level
        self.g['water_y'] = self.water_level
        self.g['cap_top_y'] = self.cap_bottom + self.cap_h
        self.g['cap_bottom_y'] = self.cap_bottom

        # 垫层／封底
        concrete_bottom = self.cap_bottom
        if self.has_blinding:
            concrete_bottom -= self.blinding_thickness
        elif self.has_plug:
            concrete_bottom -= self.plug_thickness
            
        self.g['concrete_bottom'] = concrete_bottom

        # 围檩标高 (从桩顶向下累加间距)
        self.g['waler_levels'] = []
        self.g['valid_waler_indices'] = []
        
        y = self.pile_top
        for i, row in enumerate(self.waler_rows):
            try:
                spacing_str = row['entry'].get().strip()
                spacing = float(spacing_str)
                y -= spacing
                self.g['waler_levels'].append(y)
                self.g['valid_waler_indices'].append(i)
            except (ValueError, tk.TclError, AttributeError):
                # 只要未输入或输入了非法字符，直接放入 None，不再向下累加
                self.g['waler_levels'].append(None)

        # 土层界面标高 (从土顶向下)
        self.g['soil_interfaces'] = []
        y = self.solid_level
        for layer in self.soil_layers:
            y -= layer['thickness']
            self.g['soil_interfaces'].append(y)

        # 全图 Y 范围
        soil_draw_bottom = self.g['pile_bottom_y'] - 0.8
        real_ifaces = [self.g['ground_y']] + self.g['soil_interfaces']
        draw_ifaces = []
        for y_val in real_ifaces:
            if y_val >= soil_draw_bottom:
                draw_ifaces.append(y_val)
            else:
                draw_ifaces.append(soil_draw_bottom)
                break
        self.g['draw_soil_blocks'] = []
        if self.soil_layers:
            for i in range(len(draw_ifaces) - 1):
                top_y = draw_ifaces[i]
                bot_y = draw_ifaces[i+1]
                # 匹配对应的土层属性
                layer_idx = i if i < len(self.soil_layers) else len(self.soil_layers) - 1
                self.g['draw_soil_blocks'].append({
                    'top': top_y,
                    'bot': bot_y,
                    'layer_idx': layer_idx,
                    'name': self.soil_layers[layer_idx]['name']
                })
        all_ys = [self.g['pile_top_y'], self.g['ground_y'], self.g['water_y'],
                  self.g['cap_top_y'], self.g['cap_bottom_y'], self.g['concrete_bottom'],
                  soil_draw_bottom]
                  
        if self.g['waler_levels']:
            valid_waler_ys = [y for y in self.g['waler_levels'] if y is not None]
            all_ys.extend(valid_waler_ys)
            
        self.g['y_min'] = min(all_ys) - 3.5
        self.g['y_max'] = max(all_ys) + 2.0
        v_scale = 1.5
        self.g['pile_top_y'] *= v_scale
        self.g['pile_bottom_y'] *= v_scale
        self.g['ground_y'] *= v_scale
        self.g['water_y'] *= v_scale
        self.g['cap_top_y'] *= v_scale
        self.g['cap_bottom_y'] *= v_scale
        self.g['concrete_bottom'] *= v_scale
        
        self.g['waler_levels'] = [y * v_scale if y is not None else None for y in self.g['waler_levels']]
        self.g['soil_interfaces'] = [y * v_scale for y in self.g['soil_interfaces']]
        
        for block in self.g['draw_soil_blocks']:
            block['top'] *= v_scale
            block['bot'] *= v_scale
            
        self.g['y_min'] = min(all_ys) * v_scale - 0.5
        self.g['y_max'] = max(all_ys) * v_scale + 2.0

        # ---- 平面图几何 ----
        self.g['cap_x_half'] = self.cap_x / 2.0
        self.g['cap_y_half'] = self.cap_y / 2.0
        if self.is_skggz:
            D_m = self.skggz_D / 1000.0
            self.g['pile_center_x'] = self.g['cap_x_half'] + self.x_offset + D_m / 2.0
            self.g['pile_center_y'] = self.g['cap_y_half'] + self.y_offset + D_m / 2.0
            self.g['pile_eff_w'] = max(self.skggz_Gap / 1000.0, 0.001)
            self.g['pile_depth'] = D_m
        else:
            self.g['pile_center_x'] = self.g['cap_x_half'] + self.x_offset
            self.g['pile_center_y'] = self.g['cap_y_half'] + self.y_offset
            self.g['pile_eff_w'] = 0.8
            self.g['pile_depth'] = 0.371
        self.g['plan_max_y'] = max(self.g['pile_center_y'], self.g['pile_center_x'])
        # 长边单侧桩数 (两边各接半个)
        side_len = self.g['pile_center_x'] * 2
        n_long = max(int(side_len / self.g['pile_eff_w']), 1)
        self.g['n_long'] = n_long
        # 短边单侧桩数
        side_len_y = self.g['pile_center_y'] * 2
        n_short = max(int(side_len_y / self.g['pile_eff_w']), 1)
        self.g['n_short'] = n_short
        self.need_recalc = False

    # ==========================================================================
    #                            图示绘制函数
    # ==========================================================================

    def draw(self, only_highlight=False):
        """重绘立面图与平面图，并统一计算两个子图的比例与坐标范围。

        Args:
            only_highlight (bool): True 时若数据未变化则仅刷新高亮，不重算几何。

        Returns:
            None
        """
        if not only_highlight or self.need_recalc:
            self._read_all()
            self._calc_geometry()

        g = self.g
        self.ax.clear()
        self.ax.set_facecolor('white')
        self.ax_plan.clear()
        self.ax_plan.set_facecolor('white')
        self.ax_plan.set_aspect('equal')
        self.fig.texts.clear()

        # 获取当前高亮组
        highlight = getattr(self, '_last_focused_group', '')

        # 立面图
        self._draw_elev(highlight)

        # 平面图
        self._draw_plan(highlight)

        pcx = g['pile_center_x']
        pcy = g['pile_center_y']
        pd = g['pile_depth']

        # 子图物理尺寸参数
        fig_w, fig_h = self.fig.get_size_inches()
        total_w_frac = 0.98 - 0.02 - 0.12   # right - left - wspace = 0.84
        elev_w_frac = 0.6 * total_w_frac     # 0.504
        plan_w_frac = 0.4 * total_w_frac     # 0.336
        h_frac = 0.92 - 0.04                 # 0.88
        box_plan = (plan_w_frac * fig_w) / (h_frac * fig_h)
        R = elev_w_frac / plan_w_frac        # 1.5 — 立面宽高比 = 平面宽高比 × R

        # [立面图] 所需的极限数据范围 (含土层/标高标注)
        elev_x_min = g['left_anno'] - 10.0
        elev_x_max = g['right_anno'] + 2.5
        elev_y_max = g['y_max']
        elev_y_min = g['y_min']

        # [平面图] 所需的极限数据范围
        plan_x_min = -pcx - pd/2.0 - 1.0
        plan_x_max = pcx + pd/2.0 + 1.0

        # [平面图] 所需的极限数据范围 (Y 方向，纯平面坐标)
        plan_y_max = pcy + pd/2.0 + 2.5
        plan_y_min = -pcy - pd/2.0 - 2.5

        # 平面图紧密包裹 → 立面图按比例 R 扩展 → 共享 y_range
        plan_x_range = (plan_x_max - plan_x_min) * 1.03
        elev_x_range = plan_x_range * R
        elev_content_range = (elev_x_max - elev_x_min) * 1.03
        if elev_content_range > elev_x_range:
            elev_x_range = elev_content_range
            plan_x_range = elev_x_range / R

        y_range = plan_x_range / box_plan
        elev_y_needed = (elev_y_max - elev_y_min) * 1.05
        plan_y_needed = (plan_y_max - plan_y_min) * 1.05
        needed_y = max(elev_y_needed, plan_y_needed)
        if needed_y > y_range:
            y_range = needed_y
            plan_x_range = y_range * box_plan
            elev_x_range = plan_x_range * R

        # 各自数据独立居中
        elev_y_mid = (elev_y_max + elev_y_min) / 2.0
        elev_ylim_top = elev_y_mid + y_range / 2.0
        elev_ylim_bottom = elev_y_mid - y_range / 2.0

        plan_y_mid = (plan_y_max + plan_y_min) / 2.0
        plan_ylim_top = plan_y_mid + y_range / 2.0
        plan_ylim_bottom = plan_y_mid - y_range / 2.0

        elev_x_mid = (elev_x_min + elev_x_max) / 2.0
        plan_x_mid = (plan_x_min + plan_x_max) / 2.0

        self.ax.set_xlim(elev_x_mid - elev_x_range / 2.0, elev_x_mid + elev_x_range / 2.0)
        self.ax.set_ylim(elev_ylim_bottom, elev_ylim_top)
        self.ax.set_aspect('equal')
        self.ax.set_anchor('N')
        self.ax.axis('off')

        self.ax_plan.set_xlim(plan_x_mid - plan_x_range / 2.0, plan_x_mid + plan_x_range / 2.0)
        self.ax_plan.set_ylim(plan_ylim_bottom, plan_ylim_top)
        self.ax_plan.set_aspect('equal')
        self.ax_plan.set_anchor('N')
        self.ax_plan.axis('off')

        # 标题对齐到中心线 (x=0)
        e_center_ax = (0 - (elev_x_mid - elev_x_range / 2.0)) / elev_x_range
        self.ax.set_title("立面示意图(长边)", fontsize=10,
                          color="#2b6fd4", fontweight="bold", fontfamily="Microsoft YaHei")
        self.ax.title.set_position((e_center_ax, 1.0))

        _pl = getattr(self, '_plan_current_layer', 0)
        p_center_ax = (0 - (plan_x_mid - plan_x_range / 2.0)) / plan_x_range
        self.ax_plan.set_title(f"平面示意图 (第{_pl+1}层支撑)", fontsize=10,
                               color="#2b6fd4", fontweight="bold", fontfamily="Microsoft YaHei")
        self.ax_plan.title.set_position((p_center_ax, 1.0))

        self.canvas.draw_idle()

    def _draw_elev(self, highlight):
        """集成所有立面图元素的绘制"""
        self._draw_soil_layers(highlight)
        self._draw_soil_surface(highlight)
        self._draw_water_level(highlight)
        self._draw_piles(highlight)
        self._draw_cap(highlight)
        self._draw_concrete(highlight)
        self._draw_walers(highlight)
        self._draw_struts(highlight)
        self._draw_centerline(highlight)
        self._draw_level_marks(highlight)
        self._draw_dimensions(highlight)

    # ==========================================================================
    #                            立面图绘制函数
    # ==========================================================================

    def _is_highlight(self, group_name):
        """判断指定高亮组当前是否处于高亮状态。

        Args:
            group_name (str): 高亮组名。

        Returns:
            bool: 当前聚焦组与 group_name 相同时返回 True。
        """
        return getattr(self, '_last_focused_group', '') == group_name

    def _hl_color(self, group_name, default):
        """按高亮状态返回高亮色或默认色。

        Args:
            group_name (str): 高亮组名。
            default (str): 非高亮时的颜色。

        Returns:
            str: 颜色字符串。
        """
        return self.PILE_HIGHLIGHT if self._is_highlight(group_name) else default

    # --- 土顶 ---
    def _draw_soil_surface(self, highlight):
        """绘制土顶标高线（右侧延伸 + 左侧三角标注基线）。

        Args:
            highlight (str): 当前高亮组名（本函数内部按 "ground_level" 判断）。

        Returns:
            None
        """
        g = self.g
        hl = self._is_highlight("ground_level")
        color = self.PILE_HIGHLIGHT if hl else self.SOIL_SURFACE_COLOR
        lw = 1.5 if hl else 1.0
        y = g['ground_y']
        x_right = g['pile_right_outer']
        self.ax.plot([x_right, x_right + 3.0], [y, y],
                     color=color, lw=lw, solid_capstyle='round', zorder=10)
        # 左侧延伸至土顶标高三角
        anno_x_mid = g['left_anno']-0.5
        self.ax.plot([g['pile_left_outer'], anno_x_mid], [y, y],
                     color=color, lw=lw, solid_capstyle='round', zorder=10)

    # --- 水位 ---
    def _draw_water_level(self, highlight):
        """绘制水位虚线。

        Args:
            highlight (str): 当前高亮组名（按 "water_level" 判断）。

        Returns:
            None
        """
        g = self.g
        hl = self._is_highlight("water_level")
        color = self.PILE_HIGHLIGHT if hl else self.WATER_COLOR
        y = g['water_y']
        x_left = g['left_anno'] - 0.5
        x_right = g['pile_right_outer'] + 1.5
        self.ax.plot([x_left, x_right], [y, y], color=color,
                     linestyle='--', lw=1.5 if hl else 1.0, alpha=0.9)

    # --- 土层 ---
    def _draw_soil_layers(self, highlight):
        """绘制桩外侧的土层色块及纹样。

        Args:
            highlight (str): 当前高亮组名（按 "soil" 判断）。

        Returns:
            None
        """
        g = self.g
        hl = self._is_highlight("soil")
        edge_color = self.PILE_HIGHLIGHT if hl else "#a0a5aa"
        lw = 1.8 if hl else 0.5
        xl_outer = g['pile_left_outer']

        # 直接遍历计算好的土层区块进行绘制
        for block in g['draw_soil_blocks']:
            y_top = block['top']
            y_bot = block['bot']
            layer_idx = block['layer_idx']
            hatch = self.SOIL_HATCHES[layer_idx % len(self.SOIL_HATCHES)]
            color = self.SOIL_LAYER_COLORS[layer_idx % len(self.SOIL_LAYER_COLORS)]

            # 左外侧区域
            self.ax.fill_between([g['left_anno'] + 0.5, xl_outer],
                                 y_bot, y_top,
                                 facecolor=color, edgecolor=edge_color, lw=lw, zorder=1)
            # 纹样
            self.ax.fill_between([g['left_anno'] + 0.5, xl_outer],
                                 y_bot, y_top,
                                 facecolor='none', hatch=hatch,
                                 edgecolor="#b0b5ba", lw=0, alpha=0.4, zorder=2)

    # --- 钢板桩 ---
    def _draw_piles(self, highlight):
        """绘制两侧支护桩矩形。

        Args:
            highlight (str): 当前高亮组名（按桩顶/桩长/SKGGZ 参数判断）。

        Returns:
            None
        """
        g = self.g
        hl_top = self._is_highlight("pile_top")
        hl_len = self._is_highlight("pile_length")
        hl = (hl_top or hl_len or
              self._is_highlight("skggz_d") or
              self._is_highlight("skggz_t") or
              self._is_highlight("skggz_gap"))
        color = self.PILE_HIGHLIGHT if hl else self.PILE_COLOR
        lw = 2.5 if hl else 1.0
        alpha = 1.0 if hl else 0.9

        for x_inner, x_outer in [(g['pile_left_inner'], g['pile_left_outer']),
                                  (g['pile_right_inner'], g['pile_right_outer'])]:
            width = abs(x_outer - x_inner)
            height = g['pile_top_y'] - g['pile_bottom_y']
            start_x = min(x_inner, x_outer)
            self.ax.add_patch(Rectangle(
                (start_x, g['pile_bottom_y']), width, height,
                facecolor=color, lw=lw, alpha=alpha, zorder=8
            ))

    # --- 承台  ---
    def _draw_cap(self, highlight):
        """绘制承台矩形。

        Args:
            highlight (str): 当前高亮组名（按 "cap_x"/"cap_y"/"cap_h" 判断）。

        Returns:
            None
        """
        g = self.g
        hl = self._is_highlight("cap_x") or self._is_highlight("cap_y") or self._is_highlight("cap_h")
        
        face_color = "#f5c6cb" if hl else "#d5d8dc" 
        lw = 2.0 if hl else 1.5
        cap_h_scaled = g['cap_top_y'] - g['cap_bottom_y']
        
        self.ax.add_patch(Rectangle(
            (g['cap_left'], g['cap_bottom_y']), self.cap_x, cap_h_scaled,
            facecolor=face_color, lw=lw, zorder=9
        ))


    # --- 垫层／封底 ---
    def _draw_concrete(self, highlight):
        """绘制垫层或封底混凝土块（二者互斥，垫层优先）。

        Args:
            highlight (str): 当前高亮组名（按 "concrete_blinding"/"concrete_plug" 判断）。

        Returns:
            None
        """
        g = self.g
        hl_b = self._is_highlight("concrete_blinding")
        hl_p = self._is_highlight("concrete_plug")

        inner_left = g['pile_left_inner']
        inner_width = g['pile_right_inner'] - g['pile_left_inner']
        
        top_y = g['cap_bottom_y']
        concrete_h_scaled = g['concrete_bottom'] - g['cap_bottom_y']
        
        # 垫层 
        if self.has_blinding:
            color = self.PILE_HIGHLIGHT if hl_b else self.CONCRETE_COLOR
            lw = 1.5 if hl_b else 0.8
            self.ax.add_patch(Rectangle(
                (inner_left, top_y),
                inner_width, concrete_h_scaled,  # 使用缩放后的高度
                facecolor=color, lw=lw, zorder=7, alpha=0.9
            ))
        # 封底
        elif self.has_plug:
            color = self.PILE_HIGHLIGHT if hl_p else self.CONCRETE_COLOR
            lw = 1.5 if hl_p else 0.8
            self.ax.add_patch(Rectangle(
                (inner_left, top_y),
                inner_width, concrete_h_scaled,  # 使用缩放后的高度    
                facecolor=color,  lw=lw, zorder=7, alpha=0.85
            ))

    # --- 围檩 ---
    def _draw_walers(self, highlight):
        """绘制各层围檩（左右两侧梯形截面）。

        Args:
            highlight (str): 当前高亮组名（按 "waler_i_spacing/long/short" 判断）。

        Returns:
            None
        """
        g = self.g
        left_pit_face = -g['pile_center'] + g['pile_half_w']
        right_pit_face = g['pile_center'] - g['pile_half_w']
        for i, y in enumerate(g['waler_levels']):
            if y is None: continue
            hl = (self._is_highlight(f"waler_{i}_spacing") or
                  self._is_highlight(f"waler_{i}_long") or
                  self._is_highlight(f"waler_{i}_short"))
            color = self.PILE_HIGHLIGHT if hl else self.WALER_COLOR
            lw = 1.0 if hl else 0.5

            # 读取截面尺寸（取同层长边/短边最大高度）
            h, sw = 0.3, 0.35  # 默认
            try:
                if i < len(self.waler_rows):
                    for wk in ['combo1_w', 'combo_short_w']:
                        wn = self.waler_rows[i].get(wk, ttk.Combobox()).get()
                        if wn and wn != '/' and wn in self.waler_section_dict:
                            sec = self.waler_section_dict[wn]['SEC']
                            h_candidate = float(sec[0]) / 1000.0
                            if h_candidate > h:
                                h = h_candidate
                                sw = float(sec[1]) / 1000.0 if len(sec) > 1 else sw
            except: pass
            lw_edge = sw + 0.10  # 长边 = 短边 + 固定值

            # 左侧
            pts_left = [
                (left_pit_face,y + sw / 2),
                (left_pit_face + h, y + lw_edge / 2),
                (left_pit_face + h, y - lw_edge / 2),
                (left_pit_face, y - sw / 2)
            ]
            self.ax.add_patch(Polygon(pts_left, facecolor=color, lw=lw, zorder=9, alpha=0.9))
            # 右侧
            pts_right = [
                (right_pit_face, y + sw / 2),
                (right_pit_face - h, y + lw_edge / 2),
                (right_pit_face - h, y - lw_edge / 2),
                (right_pit_face, y - sw / 2)
            ]
            self.ax.add_patch(Polygon(pts_right, facecolor=color, lw=lw, zorder=9, alpha=0.9))

    # --- 内支撑 ---
    def _draw_struts(self, highlight):
        """绘制各层内支撑（围檩之间的水平矩形）。

        Args:
            highlight (str): 当前高亮组名（按 "waler_i_dc/xc/spacing" 判断）。

        Returns:
            None
        """
        g = self.g
        left_pit_face = -g['pile_center'] + g['pile_half_w']
        right_pit_face = g['pile_center'] - g['pile_half_w']
        for i, y in enumerate(g['waler_levels']):
            if y is None: continue
            hl = (self._is_highlight(f"waler_{i}_spacing") or
                  self._is_highlight(f"waler_{i}_dc") or
                  self._is_highlight(f"waler_{i}_xc"))
            color = self.PILE_HIGHLIGHT if hl else self.STRUT_COLOR
            lw = 1.0 if hl else 0.5

            # 读取内支撑截面宽度
            sw = 0.3  # 默认
            try:
                if i < len(self.waler_rows):
                    sw_values = []
                    for key in ['combo2_w', 'combo3_w']:
                        name = self.waler_rows[i].get(key, ttk.Combobox()).get()
                        if name and name != '/':
                            sw_values.append(self._get_strut_elev_width(name))
                    sw = max(sw_values) if sw_values else 0.5
            except: pass

            # 围檩高度 (取同层长边/短边最大截面高度)
            waler_h = 0.3
            try:
                if i < len(self.waler_rows):
                    heights = []
                    for wk in ['combo1_w', 'combo_short_w']:
                        wn = self.waler_rows[i].get(wk, ttk.Combobox()).get()
                        if wn and wn != '/' and wn in self.waler_section_dict:
                            heights.append(float(self.waler_section_dict[wn]['SEC'][0]) / 1000.0)
                    if heights:
                        waler_h = max(heights)
            except: pass

            strut_left = left_pit_face + waler_h
            strut_right = right_pit_face - waler_h

            self.ax.add_patch(Rectangle(
                (strut_left, y - sw / 2),
                strut_right - strut_left, sw,
                facecolor=color, lw=lw, zorder=9, alpha=0.9
            ))

    def _get_strut_elev_width(self, name):
        """读取内支撑截面在立面图中的宽度（mm 转 m）。

        H/I 型取 SEC[1]（宽度），钢管（名称含 'X'）取 SEC[0]（直径）。

        Args:
            name (str): 内支撑截面名称。

        Returns:
            float: 截面宽度（m）；无效或未知时返回默认 0.3。
        """
        if not name or name == '/':
            return 0.3
        if name in self.strut_section_dict:
            s = self.strut_section_dict[name].get('SEC', [])
            nu = name.upper()
            try:
                # H/I型: SEC[1] = 宽度
                if 'H' in nu or 'I' in nu:
                    if len(s) >= 2: return float(s[1]) / 1000.0
                # 钢管桩: SEC[0] = 直径
                elif 'X' in nu:
                    if len(s) >= 1: return float(s[0]) / 1000.0
            except: pass
        return 0.3

    # --- 中心线 ---
    def _draw_centerline(self, highlight):
        """绘制立面图竖直中心线（洋红虚线）。

        Args:
            highlight (str): 当前高亮组名（按 "centerline" 判断）。

        Returns:
            None
        """
        g = self.g
        hl = self._is_highlight("centerline")
        color = self.PILE_HIGHLIGHT if hl else self.CENTERLINE_COLOR
        lw = 1.8 if hl else 1.0
        
        y_bot = g['pile_bottom_y'] - 0.5
        y_top = g['pile_top_y'] + 0.5
        self.ax.plot([0, 0], [y_bot, y_top], color=color,
                     linestyle='--', lw=lw, alpha=0.85, zorder=30)

    # --- 标高标注 ---
    def _draw_level_marks(self, highlight):
        """绘制左右两侧的标高与土层名称标注（带防重叠偏移）。

        Args:
            highlight (str): 当前高亮组名（按各标高组名判断）。

        Returns:
            None
        """
        g = self.g
        anno_x_left = g['left_anno'] + 1.2
        anno_x_right = g['pile_right_outer'] + 1.0

        def split_text(text, max_len=8):
            """将长文本按 max_len 折行。

            Args:
                text (str): 原始文本。
                max_len (int): 每行最大字符数。

            Returns:
                str: 折行后的多行字符串。
            """
            if not text: return ""
            return "\n".join([text[i:i+max_len] for i in range(0, len(text), max_len)])

        # ===== 右侧标高标注 =====
        cap_top = self.cap_bottom + self.cap_h
        r_items = [
            ("cap_top", g['cap_top_y'], f"承台顶标高 {'+' if cap_top>=0 else ''}{cap_top:.2f}", 9),
            ("pile_top", g['pile_top_y'], f"围堰顶标高 {'+' if self.pile_top>=0 else ''}{self.pile_top:.2f}", 9),
            ("pile_bottom", g['pile_bottom_y'], f"围堰底标高 {'+' if self.pile_bottom>=0 else ''}{self.pile_bottom:.2f}", 9),
            ("cap_bottom", g['cap_bottom_y'], f"承台底标高 {'+' if self.cap_bottom>=0 else ''}{self.cap_bottom:.2f}", 9),
        ]
        r_offsets = [0.0] * len(r_items)
        for i in range(len(r_items) - 1):
            gap = r_items[i][1] - r_items[i+1][1]
            if gap < 1.0:
                r_offsets[i] += 0.4
                r_offsets[i+1] -= 0.4
        for (group, y, label, fs), off in zip(r_items, r_offsets):
            hl = self._is_highlight(group)
            color = self.PILE_HIGHLIGHT if hl else self.LEVEL_LINE_COLOR
            draw_elevation_marker(
                self.ax, anno_x_right, y, label,
                text_side='right',
                scale=1.25,
                color=color, fontsize=fs,
                fontweight='bold' if hl else 'normal',
                facecolor='white' if not hl else '#fdecea',
                edgewidth=1.5 if hl else 1.0,
                text_voffset=off,
            )

        # ===== 左侧土层名称标注 (带防重叠偏移) =====
        l_items = []
        for block in g['draw_soil_blocks']:
            y = block['bot']
            name = split_text(block['name'], 9)
            if name.strip():
                l_items.append((y, name))
        l_items.sort(key=lambda x: -x[0])

        l_offsets = [0.0] * len(l_items)
        for i in range(len(l_items) - 1):
            gap = l_items[i][0] - l_items[i+1][0]
            if gap < 0.8:
                l_offsets[i] += 0.3
                l_offsets[i+1] -= 0.3

        hl_soil = self._is_highlight("soil")
        soil_color = self.PILE_HIGHLIGHT if hl_soil else self.LEVEL_LINE_COLOR
        for (y, name), off in zip(l_items, l_offsets):
            draw_elevation_marker(
                self.ax, anno_x_left, y, name,
                text_side='right',
                scale=0.75,
                color=soil_color, fontsize=6,
                fontweight='bold' if hl_soil else 'normal',
                facecolor='white' if not hl_soil else '#fdecea',
                edgewidth=0.8, alpha=0.7,
                text_voffset=off - 0.225,
            )

        # ===== 左侧土顶/水位标高 (三角在右, 文字在左, 防重叠) =====
        anno_x_mid = g['left_anno']
        m_items = [
            ("ground_level", g['ground_y'], f"土顶标高 {'+' if self.solid_level>=0 else ''}{self.solid_level:.2f}", 9),
            ("water_level", g['water_y'], f"水位标高 {'+' if self.water_level>=0 else ''}{self.water_level:.2f}", 9),
        ]
        m_items.sort(key=lambda x: -x[1])
        MIN_GAP = 1.2
        m_offsets = [0.0] * len(m_items)
        for i in range(len(m_items) - 1):
            gap = m_items[i][1] - m_items[i+1][1]
            if gap < MIN_GAP:
                push = (MIN_GAP - gap) / 2 + 0.1
                m_offsets[i] += push
                m_offsets[i+1] -= push
        for idx, ((group, y, label, fs), off) in enumerate(zip(m_items, m_offsets)):
            # m_items 已按 y 从大到小（上→下）排序
            # 上面的 item 用 above（三角朝下，文字在上），下面的用 below（三角朝上，文字在下）
            marker_side = 'above' if idx == 0 else 'below'
            hl = self._is_highlight(group)
            color = self.PILE_HIGHLIGHT if hl else self.LEVEL_LINE_COLOR
            draw_elevation_marker(
                self.ax, anno_x_mid, y, label,
                text_side='left',
                marker_side=marker_side,
                scale=1.0,
                color=color, fontsize=fs,
                fontweight='bold' if hl else 'normal',
                facecolor='white' if not hl else '#fdecea',
                edgewidth=1.5 if hl else 1.0,
                text_voffset=off,
            )

    # --- 尺寸标注 ---
    def _draw_dimensions(self, highlight):
        """绘制立面图底部承台长与两侧预留边距的尺寸标注。

        Args:
            highlight (str): 当前高亮组名（按 "cap_x"/"x_offset" 判断）。

        Returns:
            None
        """
        g = self.g
        # 分别获取承台长和长边边距的高亮状态
        hl_cap = self._is_highlight("cap_x")
        hl_offset = self._is_highlight("x_offset")

        line_color = "#2b6fd4" 
        dim_y = g['concrete_bottom'] - 2.0
        tick_h = 0.25

        # ================= 承台长 =================
        left_pile = -g['pile_center'] + g['pile_half_w']
        cap_l = g['cap_left']
        cap_r = g['cap_right']
        right_pile = g['pile_center'] - g['pile_half_w']
        
        self.ax.plot([left_pile, right_pile], [dim_y, dim_y], color=line_color, lw=1.2)
        ticks_x = [left_pile, cap_l, cap_r, right_pile]
        for x in ticks_x:
            self.ax.plot([x, x], [dim_y - tick_h, dim_y + tick_h], color=line_color, lw=1.2)
        # 左边距
        text_color2 = self.PILE_HIGHLIGHT if hl_offset else self.TEXT_COLOR
        gap_l = cap_l - left_pile
        self.ax.text(left_pile + gap_l * 0.65, dim_y + 0.15, f"{self.x_offset*1000:.0f}",
                     fontsize=7, color=text_color2, ha='center', va='bottom',
                     weight='bold' if hl_offset else 'normal')
        # 承台长
        text_color1 = self.PILE_HIGHLIGHT if hl_cap else self.TEXT_COLOR
        self.ax.text(0, dim_y + 0.15, f"{self.cap_x*1000:.0f}",
                     fontsize=7, color=text_color1, ha='center', va='bottom',
                     weight='bold' if hl_cap else 'normal')
        # 右边距
        gap_r = right_pile - cap_r
        self.ax.text(cap_r + gap_r * 0.35, dim_y + 0.15, f"{self.x_offset*1000:.0f}",
                     fontsize=7, color=text_color2, ha='center', va='bottom',
                     weight='bold' if hl_offset else 'normal')



    # ==========================================================================
    #                            平面图绘制函数
    # ==========================================================================
    def _get_plan_section_width(self, name, sec_dict):
        """从截面字典获取截面宽度（mm 转 m）。

        Args:
            name (str): 截面名称。
            sec_dict (dict): 截面库 {截面名: {'SEC': [...]}}。

        Returns:
            float: 截面宽度（m）；无效或未知时返回默认 0.3。
        """
        if not name or name == '/':
            return 0.0
        if name in sec_dict:
            sec = sec_dict[name]['SEC']
            if sec:
                try:
                    return float(sec[0]) / 1000.0
                except (ValueError, IndexError):
                    pass
        return 0.3

    def _get_plan_layer(self):
        """从当前高亮组名中解析选中的围檩层号。

        Returns:
            int: 层号（0 基）；非围檩高亮时返回 -1。
        """
        group = getattr(self, '_last_focused_group', '')
        if group.startswith('waler_') and '_' in group:
            parts = group.split('_')
            if len(parts) >= 2:
                try:
                    return int(parts[1])
                except ValueError:
                    pass
        # 非围檩高亮时返回-1, 由调用方决定默认层
        return -1
    
    def _draw_plan(self, highlight):
        """绘制围堰平面图：读取当前层截面/间距后依次调用各平面元素绘制函数。

        Args:
            highlight (str): 当前高亮组名。

        Returns:
            None
        """
        g = self.g
        pcx = g['pile_center_x']
        pcy = g['pile_center_y']
        cap_x = self.cap_x
        cap_y = self.cap_y
        cap_h = self.cap_h
        pw = g['pile_eff_w']
        pd = g['pile_depth']

        hl_xo = self._is_highlight("x_offset")
        hl_yo = self._is_highlight("y_offset")

        # ---- 当前层检测 ----
        plan_ri = self._get_plan_layer()
        if plan_ri >= 0 and (plan_ri >= len(g['waler_levels']) or g['waler_levels'][plan_ri] is None):
            # 如果当前选中的层非法，不生成围檩和支撑
            show_waler_and_strut = False
        else:
            show_waler_and_strut = True
        # 如果没有聚焦高亮到特定层，默认寻找第一行合法的数据展示
        if plan_ri < 0:
            if g['valid_waler_indices']:
                plan_ri = g['valid_waler_indices'][0]
                show_waler_and_strut = True
            else:
                plan_ri = 0
                show_waler_and_strut = False
                
        self._plan_current_layer = plan_ri

        hl_dc = self._is_highlight(f"waler_{plan_ri}_dc") or self._is_highlight(f"waler_{plan_ri}_spacing")
        hl_xc = self._is_highlight(f"waler_{plan_ri}_xc") or self._is_highlight(f"waler_{plan_ri}_spacing")
        hl_long = self._is_highlight(f"waler_{plan_ri}_spacing") or self._is_highlight(f"waler_{plan_ri}_long")
        hl_short = self._is_highlight(f"waler_{plan_ri}_spacing") or self._is_highlight(f"waler_{plan_ri}_short")
        hl_corbel_long = self._is_highlight("corbel_long")
        hl_corbel_short = self._is_highlight("corbel_short")

        # ---- 读取围檩截面（长边+短边） ----
        waler_sec_long, waler_sec_short = '', ''
        try:
            if plan_ri < len(self.waler_rows):
                waler_sec_long = self.waler_rows[plan_ri].get('combo1_w', ttk.Combobox()).get()
                waler_sec_short = self.waler_rows[plan_ri].get('combo_short_w', ttk.Combobox()).get()
        except: pass
        ww_long = max(self._get_plan_section_width(waler_sec_long, self.waler_section_dict), 0.3)
        ww_short = max(self._get_plan_section_width(waler_sec_short, self.waler_section_dict), 0.3)
        ww_max = max(ww_long, ww_short)

        # ---- 读取支撑截面 ----
        dc_sec, xc_sec = '', ''
        try:
            if plan_ri < len(self.waler_rows):
                dc_sec = self.waler_rows[plan_ri].get('combo2_w', ttk.Combobox()).get().strip()
                xc_sec = self.waler_rows[plan_ri].get('combo3_w', ttk.Combobox()).get().strip()
        except: pass

        def _strut_w(sec):
            """从内支撑截面库读取截面宽度（mm 转 m）。

            Args:
                sec (str): 内支撑截面名称。

            Returns:
                float: 截面宽度（m）；无效时返回默认 0.4。
            """
            if not sec or sec == '/': return 0.0
            if sec in self.strut_section_dict:
                s = self.strut_section_dict[sec].get('SEC', [])
                try:
                    return float(s[0]) / 1000.0
                except: pass
            return 0.4

        dc_w = _strut_w(dc_sec)
        xc_w = _strut_w(xc_sec)

        # ---- 读取支撑间距文本 ----
        dc_vars = self.strut_blocks.get(plan_ri, {})
        dc_x_s = (dc_vars.get('DC_X').get() or '').strip() if dc_vars.get('DC_X') else ''
        dc_y_s = (dc_vars.get('DC_Y').get() or '').strip() if dc_vars.get('DC_Y') else ''
        xc_x_s = (dc_vars.get('XC_X').get() or '').strip() if dc_vars.get('XC_X') else ''
        xc_y_s = (dc_vars.get('XC_Y').get() or '').strip() if dc_vars.get('XC_Y') else ''

        def _parse(txt):
            """解析支撑间距字符串为绝对距离列表（m）。

            '1.0+1.0' 按累积处理 -> [1.0, 2.0]；
            '1.0,2.0' 按绝对距离处理 -> [1.0, 2.0]；含 '/' 或 '—' 返回 []。

            Args:
                txt (str): 间距字符串。

            Returns:
                list[float]: 绝对距离列表（m）。
            """
            if not txt or '/' in txt or '—' in txt: return []
            try:
                # 优先按 + 号拆分 (累积间距)
                parts = [p.strip() for p in txt.split('+') if p.strip()]
                if len(parts) > 1:
                    vals = [float(p) for p in parts]
                    accum = 0.0
                    result = []
                    for v in vals:
                        accum += v
                        result.append(accum)
                    return result
                # 再按逗号/空格拆分 (绝对距离)
                parts = [p.strip() for p in re.split(r'[,，\s]+', txt) if p.strip()]
                if len(parts) > 1:
                    return [float(p) for p in parts]
                # 单个值
                return [float(txt.strip())]
            except: return []

        dc_x_list = _parse(dc_x_s)
        dc_y_list = _parse(dc_y_s)
        xc_x_list = _parse(xc_x_s)
        xc_y_list = _parse(xc_y_s)

        # ===== 绘制各元素 =====
        self._draw_plan_cap(cap_x, cap_y)
        if self.is_skggz:
            hl_skggz = (self._is_highlight("skggz_d") or
                        self._is_highlight("skggz_t") or
                        self._is_highlight("skggz_gap"))
            pc_plan = self.PILE_HIGHLIGHT if hl_skggz else self.PILE_COLOR
            self._draw_plan_pile_skggz(self.ax_plan, pcx, pcy, pw, self.skggz_D / 1000.0, pc_plan)
        else:
            self._draw_plan_pile(self.ax_plan, pcx, pcy, w=pw, d=pd, color=self.PILE_COLOR)
        self._draw_plan_centerlines(pcx, pcy, pd)
        if show_waler_and_strut:
            self._draw_plan_walers(pcx, pcy, pd, ww_long, ww_short, hl_long, hl_short, hl_corbel_long, hl_corbel_short)
            self._draw_plan_struts(pcx, pcy, pd, ww_long, ww_short, dc_sec, dc_w, xc_sec, xc_w,
                                dc_x_list, dc_y_list, xc_x_list, xc_y_list,
                                hl_dc, hl_xc)
            self._draw_plan_dims(pcx, pcy, cap_x, cap_y, dc_x_list, dc_y_list,
                                xc_x_list, xc_y_list, hl_xo, hl_yo, ww_max, pd)
        else:
            self._draw_plan_dims(pcx, pcy, cap_x, cap_y, [], [], [], [], hl_xo, hl_yo, ww_max, pd)
        self._draw_plan_title(plan_ri)

    def _draw_plan_cap(self, cap_x, cap_y):
        """绘制平面图中的承台矩形。

        Args:
            cap_x (float): 承台长（m）。
            cap_y (float): 承台宽（m）。

        Returns:
            None
        """
        ax = self.ax_plan
        hc = self._is_highlight("cap_x") or self._is_highlight("cap_y") or self._is_highlight("cap_h")
        ax.add_patch(Rectangle(
            (-cap_x/2, -cap_y/2), cap_x, cap_y,
            facecolor="#d5d8dc" if not hc else '#f5c6cb', lw=1.5, zorder=5))
    
    def _draw_plan_centerlines(self, pcx, pcy, pd):
        """绘制平面图 X/Y 双向中心线（洋红虚线）。

        Args:
            pcx (float): 桩中心 X 坐标（m）。
            pcy (float): 桩中心 Y 坐标（m）。
            pd (float): 桩截面深度/直径（m）。

        Returns:
            None
        """
        ax = self.ax_plan
        cl = self.CENTERLINE_COLOR

        v_top = pcy + 1.2 + 0.5 
        v_bot = -pcy - (pd / 2.0) - 0.5
        h_left = -pcx - (pd / 2.0) - 0.5
        h_right = pcx + 1.5 + 0.5

        # 垂直中心线 (X=0)
        ax.plot([0, 0], [v_bot, v_top], color=cl, linestyle='--', lw=1.2, alpha=0.75, zorder=20)
        # 水平中心线 (Y=0)
        ax.plot([h_left, h_right], [0, 0], color=cl, linestyle='--', lw=1.2, alpha=0.75, zorder=20)

    def _draw_plan_pile(self, ax, pcx, pcy, w=0.8, d=0.371, t=0.06, color='#4a5056'):
        """绘制钢板桩平面图（沿矩形周边拼接波浪形 U 型桩轮廓）。

        Args:
            ax (matplotlib.axes.Axes): 目标坐标轴。
            pcx (float): 桩中心 X 坐标（m）。
            pcy (float): 桩中心 Y 坐标（m）。
            w (float): 单桩有效宽度（m）。
            d (float): 桩截面深度（m）。
            t (float): 视觉壁厚（m），用于生成内外圈。
            color (str): 填充色。

        Returns:
            None
        """
        Nx = max(int(round(2 * pcx / w)), 1)
        Ny = max(int(round(2 * pcy / w)), 1)
        wx = 2 * pcx / Nx
        wy = 2 * pcy / Ny

        def build_profile(num_full, add_start_half=False, add_end_half=False, wave_w=0.8, t_off=0):
            """生成一条边上的 U 型钢板桩波浪轮廓点列。

            Args:
                num_full (int): 完整波形数量。
                add_start_half (bool): 是否在起点补半个波形。
                add_end_half (bool): 是否在终点补半个波形。
                wave_w (float): 单个波形宽度（m）。
                t_off (float): 截面偏移（m），用于生成内圈。

            Returns:
                list[tuple]: 轮廓点 [(x, y), ...]（局部坐标）。
            """
            pts = []
            v_low = -d/2 + t_off
            v_high = d/2 - t_off

            if add_start_half:
                pts.extend([
                    (-0.5*wave_w, v_high),
                    (-0.35*wave_w, v_high),
                    (-0.15*wave_w, v_low),
                    (0, v_low)
                ])

            for i in range(num_full):
                u0 = i * wave_w
                if not pts or pts[-1] != (u0, v_low):
                    pts.append((u0, v_low))
                pts.extend([
                    (u0 + 0.15*wave_w, v_low),
                    (u0 + 0.35*wave_w, v_high),
                    (u0 + 0.65*wave_w, v_high),
                    (u0 + 0.85*wave_w, v_low),
                    (u0 + wave_w, v_low)
                ])

            if add_end_half:
                u0 = num_full * wave_w
                if not pts or pts[-1] != (u0, v_low):
                    pts.append((u0, v_low))
                pts.extend([
                    (u0 + 0.15*wave_w, v_low),
                    (u0 + 0.35*wave_w, v_high),
                    (u0 + 0.5*wave_w, v_high)
                ])
            return pts

        # ================= 1. 生成外圈路径 =================
        pts_outer = []
        for u, v in build_profile(Nx, False, False, wx, 0): pts_outer.append((-pcx + u, pcy + v))
        for u, v in build_profile(Ny, False, False, wy, 0): pts_outer.append((pcx + v, pcy - u))
        for u, v in build_profile(Nx, False, False, wx, 0): pts_outer.append((pcx - u, -pcy - v))
        for u, v in build_profile(Ny, False, False, wy, 0): pts_outer.append((-pcx - v, -pcy + u))

        # ================= 2. 生成内圈路径 =================
        pts_inner = []
        for u, v in build_profile(Nx, False, False, wx, t): pts_inner.append((-pcx + u, pcy + v))
        for u, v in build_profile(Ny, False, False, wy, t): pts_inner.append((pcx + v, pcy - u))
        for u, v in build_profile(Nx, False, False, wx, t): pts_inner.append((pcx - u, -pcy - v))
        for u, v in build_profile(Ny, False, False, wy, t): pts_inner.append((-pcx - v, -pcy + u))

        # ================= 3. 缝合拼接逻辑
        poly = Polygon(pts_outer + pts_inner[::-1], facecolor=color, edgecolor='#2c3035', lw=0.5, zorder=8)
        ax.add_patch(poly)

    # --- 锁扣钢管桩平面图 ---
    def _draw_plan_pile_skggz(self, ax, pcx, pcy, gap_w, D_m, color='#4a5056'):
        """绘制锁扣钢管桩平面图：沿矩形周边画管环。

        Args:
            ax (matplotlib.axes.Axes): 目标坐标轴。
            pcx (float): 桩中心 X 坐标（m）。
            pcy (float): 桩中心 Y 坐标（m）。
            gap_w (float): 桩中心间距（m）。
            D_m (float): 钢管桩直径（m）。
            color (str): 填充色。

        Returns:
            None
        """
        if D_m <= 0.001:
            return
        R = D_m / 2.0
        t_m = self.skggz_t / 1000.0
        t_vis = max(t_m, R * 0.15)
        r_vis = R - t_vis
        if r_vis <= 0:
            r_vis = R * 0.5

        hl = (self._is_highlight("skggz_d") or
              self._is_highlight("skggz_t") or
              self._is_highlight("skggz_gap"))
        ec = self.PILE_HIGHLIGHT if hl else '#2c3035'
        fc = self.PILE_HIGHLIGHT if hl else self.PILE_COLOR

        n_long = max(int(round(2 * pcx / gap_w)), 1)
        n_short = max(int(round(2 * pcy / gap_w)), 1)
        ax_gap_x = 2 * pcx / n_long if n_long > 0 else gap_w
        ax_gap_y = 2 * pcy / n_short if n_short > 0 else gap_w

        pipe_positions = []
        for i in range(n_long + 1):
            pipe_positions.append((-pcx + i * ax_gap_x, pcy, 'top'))
        for i in range(1, n_short + 1):
            pipe_positions.append((pcx, pcy - i * ax_gap_y, 'right'))
        for i in range(1, n_long + 1):
            pipe_positions.append((pcx - i * ax_gap_x, -pcy, 'bottom'))
        for i in range(1, n_short):
            pipe_positions.append((-pcx, -pcy + i * ax_gap_y, 'left'))

        for px, py, side in pipe_positions:
            ax.add_patch(Circle((px, py), R, facecolor=fc, edgecolor=ec, lw=1.0, zorder=8))
            ax.add_patch(Circle((px, py), r_vis, facecolor="white", edgecolor="none", zorder=9))

    # def _draw_skggz_connector(self, ax, x1, y1, x2, y2, R, gap, color, ec):
    #     """两钢管桩外壁间的波形锁扣连接器"""
    #     if gap <= 0.001:
    #         return
            
    #     dx, dy = x2 - x1, y2 - y1
    #     L = (dx**2 + dy**2)**0.5
    #     if L < 0.001:
    #         return
            
    #     # ux, uy 为沿两桩圆心连线的单位向量
    #     ux, uy = dx / L, dy / L
    #     # nx, ny 为垂直于连线的法向单位向量 (控制波浪的凸出方向)
    #     nx, ny = -uy, ux 
        
    #     # 计算连接器在两圆边界上的起止点
    #     sx1, sy1 = x1 + ux * R, y1 + uy * R
    #     sx2, sy2 = x2 - ux * R, y2 - uy * R
        
    #     # --- 锁扣波浪形核心参数 ---
    #     wave_h = 0.18  # 波浪凸出高度 (m)，对标普通钢板桩深度的一半
    #     t = 0.03       # 锁扣钢材本身的视觉厚度 (m)
    #     w = gap        # 波浪总宽度刚好等于间距
        
    #     overlap = 0.05
    #     sx1_adj = sx1 - ux * overlap
    #     sy1_adj = sy1 - uy * overlap
    #     sx2_adj = sx2 + ux * overlap
    #     sy2_adj = sy2 + uy * overlap
    #     w_adj = w + 2 * overlap
        
    #     # ====== 构建外侧轮廓 ======
    #     # 1. 起点
    #     o1 = (sx1_adj, sy1_adj)
    #     # 2. 平缓起步 (走总宽的 20%)
    #     o2 = (sx1_adj + ux * w_adj * 0.2, sy1_adj + uy * w_adj * 0.2)
    #     # 3. 升起波峰 (在 35% 处达到波峰高度)
    #     o3 = (sx1_adj + ux * w_adj * 0.35 + nx * wave_h, sy1_adj + uy * w_adj * 0.35 + ny * wave_h)
    #     # 4. 维持波峰 (平移到 65% 处)
    #     o4 = (sx1_adj + ux * w_adj * 0.65 + nx * wave_h, sy1_adj + uy * w_adj * 0.65 + ny * wave_h)
    #     # 5. 降落 (在 80% 处回到轴线)
    #     o5 = (sx1_adj + ux * w_adj * 0.8, sy1_adj + uy * w_adj * 0.8)
    #     # 6. 终点
    #     o6 = (sx2_adj, sy2_adj)

    #     i1 = (sx2_adj - nx * t, sy2_adj - ny * t)
    #     i2 = (sx1_adj + ux * w_adj * 0.8 - nx * t, sy1_adj + uy * w_adj * 0.8 - ny * t)
    #     i3 = (sx1_adj + ux * w_adj * 0.65 + nx * (wave_h - t), sy1_adj + uy * w_adj * 0.65 + ny * (wave_h - t))
    #     i4 = (sx1_adj + ux * w_adj * 0.35 + nx * (wave_h - t), sy1_adj + uy * w_adj * 0.35 + ny * (wave_h - t))
    #     i5 = (sx1_adj + ux * w_adj * 0.2 - nx * t, sy1_adj + uy * w_adj * 0.2 - ny * t)
    #     i6 = (sx1_adj - nx * t, sy1_adj - ny * t)

    #     pts = [o1, o2, o3, o4, o5, o6, i1, i2, i3, i4, i5, i6]
        
    #     ax.add_patch(Polygon(pts, facecolor=color, edgecolor=ec, lw=0.5, zorder=7))

    def _draw_plan_walers(self, pcx, pcy, pd, ww_long, ww_short,
                          hl_long=False, hl_short=False,
                          hl_corbel_long=False, hl_corbel_short=False):
        """绘制平面图中当前层的围檩（长边/短边矩形 + 中心线）。

        Args:
            pcx, pcy (float): 桩中心 X/Y 坐标（m）。
            pd (float): 桩截面深度/直径（m）。
            ww_long, ww_short (float): 围檩长边/短边宽度（m）。
            hl_long, hl_short (bool): 长边/短边是否高亮。
            hl_corbel_long, hl_corbel_short (bool): 牛腿相关高亮。

        Returns:
            None
        """
        ax = self.ax_plan
        default_color = '#4a5056'
        long_color = self.PILE_HIGHLIGHT if (hl_long or hl_corbel_long) else default_color
        short_color = self.PILE_HIGHLIGHT if (hl_short or hl_corbel_short) else default_color
        iwx = pcx - pd/2.0
        iwy = pcy - pd/2.0
        wh_long = ww_long / 2.0
        wh_short = ww_short / 2.0

        # 实体矩形 — 长边（上下，用 ww_long）
        ax.add_patch(Rectangle((-iwx, iwy - ww_long), 2*iwx, ww_long,
            facecolor=long_color, lw=0.8, zorder=9, alpha=0.85))
        ax.add_patch(Rectangle((-iwx, -iwy), 2*iwx, ww_long,
            facecolor=long_color, lw=0.8, zorder=9, alpha=0.85))
        # 实体矩形 — 短边（左右，用 ww_short）
        ax.add_patch(Rectangle((-iwx, -iwy + ww_long), ww_short, 2*iwy - 2*ww_long,
            facecolor=short_color, lw=0.8, zorder=9, alpha=0.85))
        ax.add_patch(Rectangle((iwx - ww_short, -iwy + ww_long), ww_short, 2*iwy - 2*ww_long,
            facecolor=short_color, lw=0.8, zorder=9, alpha=0.85))

        # 中心线
        cl = "#ffffff"
        long_cl = self.PILE_HIGHLIGHT if (hl_long or hl_corbel_long) else cl
        short_cl = self.PILE_HIGHLIGHT if (hl_short or hl_corbel_short) else cl
        # 长边
        ax.plot([-iwx, iwx], [iwy - wh_long, iwy - wh_long], color=long_cl, ls='-.', lw=0.8, zorder=13, alpha=0.85)
        ax.plot([-iwx, iwx], [-iwy + wh_long, -iwy + wh_long], color=long_cl, ls='-.', lw=0.8, zorder=13, alpha=0.85)
        # 短边
        ax.plot([-iwx + wh_short, -iwx + wh_short], [-iwy + ww_long, iwy - ww_long], color=short_cl, ls='-.', lw=0.8, zorder=13, alpha=0.85)
        ax.plot([iwx - wh_short, iwx - wh_short], [-iwy + ww_long, iwy - ww_long], color=short_cl, ls='-.', lw=0.8, zorder=13, alpha=0.85)

    def _draw_plan_struts(self, pcx, pcy, pd, ww_long, ww_short, dc_sec, dc_w, xc_sec, xc_w,
                          dc_x_list, dc_y_list, xc_x_list, xc_y_list, hl_dc, hl_xc):
        """绘制平面图中的对撑与斜撑。

        Args:
            pcx, pcy (float): 桩中心 X/Y 坐标（m）。
            pd (float): 桩截面深度/直径（m）。
            ww_long, ww_short (float): 围檩长边/短边宽度（m）。
            dc_sec, xc_sec (str): 对撑/斜撑截面名。
            dc_w, xc_w (float): 对撑/斜撑截面宽度（m）。
            dc_x_list, dc_y_list (list[float]): 对撑 X/Y 向绝对距离（m）。
            xc_x_list, xc_y_list (list[float]): 斜撑 X/Y 向绝对距离（m）。
            hl_dc, hl_xc (bool): 对撑/斜撑是否高亮。

        Returns:
            None
        """
        ax = self.ax_plan
        iwx = pcx - pd/2.0
        iwy = pcy - pd/2.0
        wh_long = ww_long / 2.0
        wh_short = ww_short / 2.0
        wc_top_y = iwy - wh_long
        wc_bot_y = -iwy + wh_long
        wc_left_x = -iwx + wh_short
        wc_right_x = iwx - wh_short

        dc_sc = self.PILE_HIGHLIGHT if hl_dc else '#5b6168'
        xc_sc = self.PILE_HIGHLIGHT if hl_xc else '#5b6168'
        cl = "#ffffff"

        # X方向对撑（长边之间，用 ww_long）
        if dc_sec != '/' and dc_w > 0.001:
            for dx in dc_x_list:
                locs = [dx] if abs(dx) < 0.001 else [-dx, dx]
                for xl in locs:
                    if abs(xl) < iwx:
                        ax.add_patch(Rectangle((xl-dc_w/2, -iwy+ww_long), dc_w, 2*iwy-2*ww_long,
                            facecolor=dc_sc, lw=0.5, zorder=11, alpha=0.85))
                        ax.plot([xl, xl], [wc_bot_y, wc_top_y],
                                color=cl, ls='-.', lw=0.8, zorder=14, alpha=0.85)

        # Y方向对撑（短边之间，用 ww_short）
        if dc_sec != '/' and dc_w > 0.001:
            for dy in dc_y_list:
                locs = [dy] if abs(dy) < 0.001 else [-dy, dy]
                for yl in locs:
                    if abs(yl) < iwy:
                        ax.add_patch(Rectangle((-iwx+ww_short, yl-dc_w/2), 2*iwx-2*ww_short, dc_w,
                            facecolor=dc_sc, lw=0.5, zorder=11, alpha=0.85))
                        ax.plot([wc_left_x, wc_right_x], [yl, yl],
                                color=cl, ls='-.', lw=0.8, zorder=14, alpha=0.85)

        # 斜撑 (支持多道: 每对 xc_x[i], xc_y[i] 定义一道)
        n_xc = min(len(xc_x_list), len(xc_y_list))
        if xc_sec != '/' and xc_w > 0.001 and n_xc > 0:
            for k in range(n_xc):
                xc_x = xc_x_list[k]
                xc_y = xc_y_list[k]
                if xc_x < 0.001 or xc_y < 0.001:
                    continue
                for sx in [-1, 1]:
                    for sy in [-1, 1]:
                        x1 = wc_right_x if sx > 0 else wc_left_x
                        y1 = sy * xc_y
                        x2 = sx * xc_x
                        y2 = wc_top_y if sy > 0 else wc_bot_y

                        dx, dy = x2 - x1, y2 - y1
                        length = (dx**2 + dy**2)**0.5
                        if length > 0.01:
                            nx = -dy / length * (xc_w/2)
                            ny =  dx / length * (xc_w/2)
                            pts = [(x1+nx, y1+ny), (x2+nx, y2+ny), (x2-nx, y2-ny), (x1-nx, y1-ny)]
                            ax.add_patch(Polygon(pts, facecolor=xc_sc, lw=0.5, zorder=11, alpha=0.85))
                            ax.plot([x1, x2], [y1, y2], color=cl, ls='-.', lw=0.8, zorder=14, alpha=0.85)

    def _draw_plan_dims(self, pcx, pcy, cap_x, cap_y, dc_x_list, dc_y_list,
                        xc_x_list, xc_y_list, hl_xo, hl_yo, ww, pd):
        """绘制平面图尺寸标注（承台长宽、预留边距、对撑/斜撑间距）。

        Args:
            pcx, pcy (float): 桩中心 X/Y 坐标（m）。
            cap_x, cap_y (float): 承台长/宽（m）。
            dc_x_list, dc_y_list (list[float]): 对撑 X/Y 向绝对距离（m）。
            xc_x_list, xc_y_list (list[float]): 斜撑 X/Y 向绝对距离（m）。
            hl_xo, hl_yo (bool): X/Y 向预留边距是否高亮。
            ww (float): 围檩宽度（m）。
            pd (float): 桩截面深度/直径（m）。

        Returns:
            None
        """
        ax = self.ax_plan
        ac, zc, fs = '#2b6fd4', self.TEXT_COLOR, 7
        hl_cx = self._is_highlight("cap_x")
        hl_cy = self._is_highlight("cap_y")
        pile_hw = self.g.get('pile_half_w', 0.25)

        # 桩内侧边 (与立面图一致: 标注距桩内侧边而非中心线)
        lx_in = -pcx + pile_hw; rx_in = pcx - pile_hw
        by_in = -pcy + pile_hw; ty_in = pcy - pile_hw

        # 底部连续标注
        ay = -pcy - 2.5; th = 0.2
        rx = [lx_in, -cap_x/2, cap_x/2, rx_in]
        ax.plot([lx_in, rx_in], [ay, ay], color=ac, lw=1.0)
        for x in rx: ax.plot([x, x], [ay-th, ay+th], color=ac, lw=1.0)
        c1 = self.PILE_HIGHLIGHT if hl_xo else zc
        c2 = self.PILE_HIGHLIGHT if hl_cx else zc
        ax.text((lx_in-cap_x/2)/2, ay+0.2, f"{self.x_offset*1000:.0f}", fontsize=fs, color=c1, ha='center', va='bottom', weight='bold' if hl_xo else 'normal')
        ax.text(0, ay+0.2, f"{self.cap_x*1000:.0f}", fontsize=fs, color=c2, ha='center', va='bottom', weight='bold' if hl_cx else 'normal')
        ax.text((cap_x/2+rx_in)/2, ay+0.2, f"{self.x_offset*1000:.0f}", fontsize=fs, color=c1, ha='center', va='bottom', weight='bold' if hl_xo else 'normal')

        # 左侧连续标注
        ax2 = -pcx - 1.4
        ry = [by_in, -cap_y/2, cap_y/2, ty_in]
        ax.plot([ax2, ax2], [by_in, ty_in], color=ac, lw=1.0)
        for y in ry: ax.plot([ax2-th, ax2+th], [y, y], color=ac, lw=1.0)
        c3 = self.PILE_HIGHLIGHT if hl_yo else zc
        c4 = self.PILE_HIGHLIGHT if hl_cy else zc
        ax.text(ax2-0.3, (by_in-cap_y/2)/2, f"{self.y_offset*1000:.0f}", fontsize=fs, color=c3, ha='right', va='center', weight='bold' if hl_yo else 'normal')
        ax.text(ax2-0.3, 0, f"{self.cap_y*1000:.0f}", fontsize=fs, color=c4, ha='right', va='center', weight='bold' if hl_cy else 'normal')
        ax.text(ax2-0.3, (cap_y/2+ty_in)/2, f"{self.y_offset*1000:.0f}", fontsize=fs, color=c3, ha='right', va='center', weight='bold' if hl_yo else 'normal')

        # ---- 对撑标注 (仅"支撑布置"按钮高亮文字, 线不变色) ----
        hl_dc_dim = self._is_highlight(f"waler_{self._plan_current_layer}_btn")
        dc_tc = self.PILE_HIGHLIGHT if hl_dc_dim else self.TEXT_COLOR

        if dc_x_list and any(abs(v) > 0.001 for v in dc_x_list):
            xs_dc = sorted(set([0.0] + [abs(v) for v in dc_x_list if abs(v) > 0.001]))
            x_lim_dc = xs_dc[-1]
            dcy = 0.5  # 中心上方, 连续标注线高度
            ax.plot([0, x_lim_dc], [dcy, dcy], color=ac, lw=1.0, zorder=25)
            for xv in xs_dc:
                ax.plot([xv, xv], [dcy - th, dcy + th], color=ac, lw=1.0, zorder=25)
            for k in range(len(xs_dc)):
                if xs_dc[k] > 0.001:
                    prev = xs_dc[k-1] if k > 0 else 0
                    mid = (prev + xs_dc[k]) / 2
                    ax.text(mid, dcy + 0.3, f"{(xs_dc[k]-prev)*1000:.0f}",
                            fontsize=fs, color=dc_tc, ha='center', va='bottom',
                            weight='bold' if hl_dc_dim else 'normal', zorder=25)

        if dc_y_list and any(abs(v) > 0.001 for v in dc_y_list):
            ys_dc = sorted(set([0.0] + [abs(v) for v in dc_y_list if abs(v) > 0.001]))
            y_lim_dc = ys_dc[-1]
            dcx = 0.5
            ax.plot([dcx, dcx], [0, y_lim_dc], color=ac, lw=1.0, zorder=25)
            for yv in ys_dc:
                ax.plot([dcx - th, dcx + th], [yv, yv], color=ac, lw=1.0, zorder=25)
            for k in range(len(ys_dc)):
                if ys_dc[k] > 0.001:
                    prev = ys_dc[k-1] if k > 0 else 0
                    mid = (prev + ys_dc[k]) / 2
                    ax.text(dcx - 0.3, mid, f"{(ys_dc[k] - prev)*1000:.0f}",
                            fontsize=fs, color=dc_tc, ha='right', va='center',
                            weight='bold' if hl_dc_dim else 'normal', zorder=25)

        # ---- 斜撑标注 (仅"支撑布置"按钮高亮文字, 线不变色) ----
        hl_xc_dim = self._is_highlight(f"waler_{self._plan_current_layer}_btn")
        xc_tc = self.PILE_HIGHLIGHT if hl_xc_dim else self.TEXT_COLOR

        if xc_x_list and xc_y_list and any(abs(v) > 0.001 for v in xc_x_list + xc_y_list):
            xs = sorted(set([0.0] + [abs(v) for v in xc_x_list if abs(v) > 0.001]))
            ys = sorted(set([0.0] + [abs(v) for v in xc_y_list if abs(v) > 0.001]))
            x_lim = xs[-1] if xs else 0.0
            y_lim = ys[-1] if ys else 0.0
            if x_lim < 0.001: x_lim = max(ys) * 0.8 if ys else 1.0
            if y_lim < 0.001: y_lim = max(xs) * 0.8 if xs else 1.0

            # ---- 右侧: X间距连续标注 (竖直) ----
            rx_anno = pcx + 1.5
            ax.plot([rx_anno, rx_anno], [0, y_lim], color=ac, lw=1.0)
            for y in ys:
                ax.plot([rx_anno - th, rx_anno + th], [y, y], color=ac, lw=1.0)
            for k in range(len(ys)):
                if ys[k] > 0.001:
                    prev = ys[k-1] if k > 0 else 0.0
                    mid = (prev + ys[k]) / 2
                    ax.text(rx_anno + 0.3, mid, f"{(ys[k] - prev)*1000:.0f}",
                            fontsize=fs, color=xc_tc, ha='left', va='center',
                            weight='bold' if hl_xc_dim else 'normal')

            # ---- 顶部: Y间距连续标注 (水平) ----
            ty_anno = pcy + 1.2
            ax.plot([0, x_lim], [ty_anno, ty_anno], color=ac, lw=1.0)
            for x in xs:
                ax.plot([x, x], [ty_anno - th, ty_anno + th], color=ac, lw=1.0)
            for k in range(len(xs)):
                if xs[k] > 0.001:
                    prev = xs[k-1] if k > 0 else 0.0
                    mid = (prev + xs[k]) / 2
                    ax.text(mid, ty_anno + 0.2, f"{(xs[k] - prev)*1000:.0f}",
                            fontsize=fs, color=xc_tc, ha='center', va='bottom',
                            weight='bold' if hl_xc_dim else 'normal')
                    
    def _draw_plan_title(self, plan_ri):
        """占位方法：平面图标题统一在 draw() 中 xlim 设定之后绘制以对齐中心线。

        Args:
            plan_ri (int): 当前层号。

        Returns:
            None
        """
        pass  # 标题移至 draw() 中 xlim 设定之后, 以对齐中心线
    # ========== 公开接口 ==========
    def clear_highlight(self):
        """外部调用：清空当前高亮组并重绘图示。

        Returns:
            None
        """
        self._last_focused_group = ""
        self.draw(only_highlight=True)

    def refresh(self):
        """外部强制刷新（土层表格变更后调用），强制重算几何并重绘。

        Returns:
            None
        """
        self.need_recalc = True
        self.draw()

    def destroy(self):
        """清理 matplotlib 资源并关闭 figure。

        Returns:
            None
        """
        self.fig.clear()
        plt.close(self.fig)


class SKGGZCrossSectionDiagram:
    """锁扣钢管桩截面俯视图 — 实时联动"""

    PILE_COLOR = "#4a5056"
    CENTERLINE_COLOR = "#dc3243"
    TEXT_COLOR = "#2c3e50"
    DIM_COLOR = "#1565C0"
    HL_COLOR = "#dc3243"

    def __init__(self, parent, D_var, t_var, Gap_var):
        """初始化截面俯视图画布并首次绘制。

        Args:
            parent (tk.Widget): 承载画布的父容器。
            D_var (tk.Variable): 钢管桩直径（mm）。
            t_var (tk.Variable): 钢管桩壁厚（mm）。
            Gap_var (tk.Variable): 桩中心间距（mm）。

        Returns:
            None
        """
        self.parent = parent
        self.D_var = D_var
        self.t_var = t_var
        self.Gap_var = Gap_var
        self._last_focused_group = ""
        self._focus_callback = None
        self.highlight_groups = {}
        self.widget_to_group = {}

        self.fig = Figure(figsize=(4.5, 2.5), dpi=100)
        self.fig.set_facecolor("white")
        self.ax = self.fig.add_subplot(111)
        self.ax.set_aspect("equal")
        self.ax.axis("off")
        self.fig.subplots_adjust(left=0.05, right=0.95, top=0.95, bottom=0.05)

        self.canvas = FigureCanvasTkAgg(self.fig, master=parent)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        # 容器尺寸变化时自动适配 figsize
        self._resize_job = None
        parent.bind("<Configure>", self._on_parent_resize)
        self.draw()
    
    # ---- 高亮联动 ----
    def _on_parent_resize(self, event):
        """容器尺寸变化时自适应 figsize"""
        w, h = event.width, event.height
        if w < 100 or h < 80:
            return
        if self._resize_job is not None:
            self.parent.after_cancel(self._resize_job)
        def _apply():
            """按容器像素尺寸更新 figure 尺寸并重绘（防抖后执行）。"""
            dpi = self.fig.get_dpi()
            self.fig.set_size_inches(w / dpi, h / dpi)
            self.draw()
            self._resize_job = None
        self._resize_job = self.parent.after(100, _apply)

    def set_focus_callback(self, cb):
        """设置焦点变化回调，用于同步主示意图高亮。

        Args:
            cb (callable): 回调函数，接收高亮组名参数。

        Returns:
            None
        """
        self._focus_callback = cb
    def register_group(self, group_name, widgets):
        """注册一组控件到指定高亮组，并绑定焦点事件。

        Args:
            group_name (str): 高亮组名。
            widgets (tk.Widget|list[tk.Widget]): 单个控件或控件列表。

        Returns:
            None
        """
        if not isinstance(widgets, (list, tuple)):
            widgets = [widgets]
        self.highlight_groups[group_name] = list(widgets)
        for w in widgets:
            self._bind_widget(w, group_name)

    def _bind_widget(self, w, group_name):
        """把单个控件的焦点事件绑定到高亮组。

        Args:
            w (tk.Widget): 控件对象。
            group_name (str): 高亮组名。

        Returns:
            None
        """
        self.widget_to_group[str(w)] = group_name
        w.bind("<FocusIn>", lambda e, g=group_name: self._on_focus_in(g))
        w.bind("<FocusOut>", self._on_focus_out)

    def _on_focus_in(self, group):
        """控件获得焦点：记录高亮组、触发回调并重绘。

        Args:
            group (str): 高亮组名。

        Returns:
            None
        """
        self._last_focused_group = group
        if self._focus_callback:
            self._focus_callback(group)
        self.draw()

    def _on_focus_out(self, event=None):
        """控件失去焦点：延迟确认后清除高亮并重绘。

        Args:
            event: Tk 事件对象（可选）。

        Returns:
            None
        """
        def do_clear():
            """确认焦点确实离开当前组后清空高亮。"""
            focused = self.parent.focus_get()
            if focused and str(focused) in self.widget_to_group:
                if self.widget_to_group[str(focused)] == self._last_focused_group:
                    return
            self._last_focused_group = ""
            if self._focus_callback:
                self._focus_callback("")
            self.draw()
        self.parent.after(10, do_clear)

    # ---- 绘图 ----
    def draw(self):
        """绘制锁扣钢管桩截面俯视图：两个管环 + 中心线 + 间距/直径/壁厚标注。

        D/t/Gap 任一非正数时仅刷新画布不重绘。

        Returns:
            None
        """
        self.ax.clear()
        self.ax.set_aspect("equal")
        self.ax.axis("off")

        try:
            D = float(self.D_var.get())
            t = float(self.t_var.get())
            gap = float(self.Gap_var.get())
        except (tk.TclError, ValueError):
            self.canvas.draw_idle()
            return

        if D <= 0 or t <= 0 or gap <= 0:
            self.canvas.draw_idle()
            return

        R = D / 2.0
        # 管壁视觉厚度: 至少 2% 半径保证可见，同时跟随 t 值变化
        t_eff = max(1.5*t, R * 0.02)
        r = R - t_eff
        if r <= 0:
            r = R * 0.5
        hl = self._last_focused_group

        pile_hl = hl in ("skggz_d","skggz_t")
        pile_ec = self.HL_COLOR if pile_hl else self.PILE_COLOR
        pile_fc = self.HL_COLOR if pile_hl else self.PILE_COLOR

        # 钢管桩 — 外圆(深灰填充+描边) + 内圆(白色填充,无边线) = 管壁环效果
        for cx in (0, gap):
            self.ax.add_patch(Circle((cx, 0), R, facecolor=pile_fc, edgecolor=pile_ec, lw=1.5, zorder=2))
            self.ax.add_patch(Circle((cx, 0), r, facecolor="white", edgecolor="none", zorder=3))

        # 中心线 — 最顶层
        ext = R * 0.5
        for cx in (0, gap):
            self.ax.plot([cx, cx], [-R - ext, R + ext], color=self.CENTERLINE_COLOR, ls="--", lw=0.8, zorder=20)
            self.ax.plot([cx - R - ext, cx + R + ext], [0, 0], color=self.CENTERLINE_COLOR, ls="--", lw=0.8, zorder=20)

        # 间距标注 (下方)
        gap_tc = self.HL_COLOR if hl == "skggz_gap" else self.TEXT_COLOR
        y_dim = -R - ext - R * 0.25
        self.ax.annotate("", xy=(gap, y_dim), xytext=(0, y_dim),
                         arrowprops=dict(arrowstyle="<->", color=self.DIM_COLOR, lw=1.2), zorder=25)
        self.ax.text(gap / 2, y_dim - R * 0.12, f"G={gap:.0f}mm",
                     ha="center", va="top", fontsize=8, color=gap_tc, fontweight="bold", zorder=25)

        # 直径标注 (右圆右侧)
        d_tc = self.HL_COLOR if hl == "skggz_d" else self.TEXT_COLOR
        x_d = gap + R + R * 0.5
        self.ax.annotate("", xy=(x_d, R), xytext=(x_d, -R),
                         arrowprops=dict(arrowstyle="<->", color=self.DIM_COLOR, lw=1.2), zorder=25)
        self.ax.text(x_d + R * 0.08, 0, f"D={D:.0f}", ha="left", va="center",
                     fontsize=8, color=d_tc, fontweight="bold", rotation=90, zorder=25)

        # 壁厚标注 (右圆左上角, 引线指向管壁)
        t_tc = self.HL_COLOR if hl == "skggz_t" else self.TEXT_COLOR
        ang = 135  # 左上角
        ang_rad = math.radians(ang)
        x_mid = (R - t / 2) * math.cos(ang_rad) + gap
        y_mid = (R - t / 2) * math.sin(ang_rad)
        x_out = (R + R * 0.55) * math.cos(ang_rad) + gap
        y_out = (R + R * 0.55) * math.sin(ang_rad)
        self.ax.plot([x_mid, x_out], [y_mid, y_out], color=self.DIM_COLOR, lw=0.8, zorder=25)
        self.ax.text(x_out - R * 0.05, y_out, f"t={t:.0f}", ha="right", va="center",
                     fontsize=7, color=t_tc, fontweight="bold", zorder=25)

        # 视图范围
        x_min = -R - R * 0.6
        x_max = gap + R + R * 0.7
        y_min = y_dim - R * 0.2
        y_max = R + ext + R * 0.5
        self.ax.set_xlim(x_min, x_max)
        self.ax.set_ylim(y_min, y_max)
        self.fig.tight_layout(pad=0.5)
        self.canvas.draw_idle()


class SurchargeDiagram:
    """围堰加载（附加荷载）示意图：绘制均布/矩形局部/条形局部附加荷载。"""

    def __init__(self, parent):
        """初始化加载示意图：标题、分隔线与画布。

        Args:
            parent (tk.Widget): 承载画布的父容器。

        Returns:
            None
        """
        self.parent = parent

        # 分隔线 + 标题
        top_sep = ttk.Separator(parent, orient="horizontal")
        top_sep.pack(fill="x", padx=10, pady=(5, 0))
        self.title_lbl = tk.Label(parent, text="加载示意图",
                                  font=("Microsoft YaHei UI", 10, "bold"),
                                  bg="white")
        self.title_lbl.pack(anchor="center", pady=(5, 3))
        sep_line = ttk.Separator(parent, orient="horizontal")
        sep_line.pack(fill="x", padx=20, pady=(0, 0))

        self.fig = Figure(figsize=(4, 3), dpi=100)
        self.fig.subplots_adjust(left=0.02, right=0.98, bottom=0.02, top=0.98)
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=parent)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)
        # 禁止图形交互（拖动/缩放）
        self.ax.set_navigate(False)
        self.fig.canvas.toolbar = None
        # Canvas 尺寸确定后重绘
        self.canvas.get_tk_widget().bind("<Configure>", lambda e: self._redraw_on_resize(e), add="+")

    def _redraw_on_resize(self, event):
        """画布尺寸变化时同步 figure 尺寸并重绘，消除空白。

        Args:
            event: Tk Configure 事件，含 width/height。

        Returns:
            None
        """
        if event.width > 50:
            # 根据 canvas 实际尺寸动态调整 figure 大小，消除下方空白
            w_inch = event.width / self.fig.dpi
            h_inch = event.height / self.fig.dpi
            if w_inch > 0 and h_inch > 0:
                self.fig.set_size_inches(w_inch, h_inch, forward=True)
            self.canvas.draw_idle()

    def _draw_standard_dim(self, p1, p2, dim_line_pos, text, orientation="h", text_side="right", highlight=False):
        """绘制一段标准尺寸标注（尺寸线 + 两端刻度 + 文字）。

        Args:
            p1, p2 (float): 标注起止坐标（水平标注为 x，竖直标注为 y）。
            dim_line_pos (float): 尺寸线位置（水平标注为 y，竖直标注为 x）。
            text (str|float): 标注文字。
            orientation (str): 'h' 水平 / 'v' 竖直。
            text_side (str): 竖直标注时文字在 'left'/'right'。
            highlight (bool): 是否高亮。

        Returns:
            None
        """
        line_color = "#2b6fd4"
        text_color = "#dc3243" if highlight else "#000000"
        text_weight = "bold" if highlight else "normal"
        tick_h = 0.25

        if orientation == "h":
            self.ax.plot([p1, p2], [dim_line_pos, dim_line_pos], color=line_color, lw=1.2)
            self.ax.plot([p1, p1], [dim_line_pos - tick_h, dim_line_pos + tick_h], color=line_color, lw=1.2)
            self.ax.plot([p2, p2], [dim_line_pos - tick_h, dim_line_pos + tick_h], color=line_color, lw=1.2)
            self.ax.text((p1 + p2) / 2, dim_line_pos + 0.15, text,
                         fontsize=8, color=text_color, ha="center", va="bottom", weight=text_weight)
        else:
            if text_side == "left":
                x_text, ha_text = dim_line_pos - 0.15, "right"
            else:
                x_text, ha_text = dim_line_pos + 0.15, "left"
            self.ax.plot([dim_line_pos, dim_line_pos], [p1, p2], color=line_color, lw=1.2)
            self.ax.plot([dim_line_pos - tick_h, dim_line_pos + tick_h], [p1, p1], color=line_color, lw=1.2)
            self.ax.plot([dim_line_pos - tick_h, dim_line_pos + tick_h], [p2, p2], color=line_color, lw=1.2)
            self.ax.text(x_text, (p1 + p2) / 2, text,
                         fontsize=8, color=text_color, ha=ha_text, va="center", weight=text_weight)

    def draw(self, sigma_k_dict, rtd=None, highlight_key=None):
        """绘制附加荷载示意图：围堰轮廓、四个面标号与各面荷载及尺寸标注。

        Args:
            sigma_k_dict (dict): 附加荷载信息，结构形式
                {'active_type': '均布附加荷载'/'矩形局部附加荷载'/'条形局部附加荷载',
                 '均布附加荷载': {面号: {'q0': str}},
                 '矩形局部附加荷载': {面号: {'p0','a','b','l','p2','c'}},
                 '条形局部附加荷载': {面号: {'p0','a','b','l'}}}。
            rtd (RealTimeDiagram|None): 主示意图实例，用于复用围堰轮廓几何。
            highlight_key (tuple|None): 高亮项，形如 (面号, 字段名)。

        Returns:
            None
        """
        from matplotlib.patches import Rectangle as MplRect
        self.ax.clear()
        self.ax.set_aspect("equal", adjustable="box")
        self.ax.axis("off")

        # 1. 围堰轮廓
        if rtd and hasattr(rtd, "_read_all"):
            rtd._read_all(); rtd._calc_geometry(); g = rtd.g
            pcx, pcy = g["pile_center_x"], g["pile_center_y"]
            half_L, half_B = pcx + g["pile_half_w"], pcy + g["pile_half_w"]
            if rtd.is_skggz: rtd._draw_plan_pile_skggz(self.ax, pcx, pcy, g["pile_eff_w"], g["pile_depth"])
            else: rtd._draw_plan_pile(self.ax, pcx, pcy, w=g["pile_eff_w"], d=g["pile_depth"])
        else:
            half_L, half_B = 10.0, 5.0
            self.ax.add_patch(MplRect((-half_L, -half_B), half_L*2, half_B*2, fill=False, edgecolor="#2c3e50", lw=2.5))

        # 2. 面标号与中心线
        self.ax.plot([-half_L-1.5, half_L+1.5], [0, 0], color="#dc3243", linestyle="-.", lw=1, zorder=1)
        self.ax.plot([0, 0], [-half_B-1.5, half_B+1.5], color="#dc3243", linestyle="-.", lw=1, zorder=1)
        for num, x, y, ha, va in [("①", 0, half_B-1.5, "center", "top"),
                                   ("②", -half_L+1.5, 0, "left", "center"),
                                   ("③", 0, -half_B+1.5, "center", "bottom"),
                                   ("④", half_L-1.5, 0, "right", "center")]:
            self.ax.text(x, y, num, ha=ha, va=va, fontsize=13, fontweight="bold")

        # 3. 荷载绘制
        active_type = sigma_k_dict.get("active_type", "")
        if not active_type or active_type not in sigma_k_dict:
            self.canvas.draw_idle(); return
        _title_map = {"均布附加荷载": "均布附加荷载示意图",
                      "矩形局部附加荷载": "矩形局部附加荷载示意图",
                      "条形局部附加荷载": "条形局部附加荷载示意图"}
        self.title_lbl.config(text=_title_map.get(active_type, "加载示意图"))


        def _f(val, default=0):
            """将值安全转为 float，失败返回 default。

            Args:
                val: 任意值。
                default (float): 转换失败时的返回值。

            Returns:
                float: 转换结果或 default。
            """
            try: return float(val)
            except: return default

        OUTSET = 2.5  # 标注距围堰边的偏移量，统一调整此值

        # ab 标注对齐参考 l 标注：右侧与 4 面 l 同 x，下侧与 3 面 l 同 y
        ref_ab4 = _f(sigma_k_dict[active_type].get("4", {}).get("a", 0)) + _f(sigma_k_dict[active_type].get("4", {}).get("b", 0))
        ref_ab3 = _f(sigma_k_dict[active_type].get("3", {}).get("a", 0)) + _f(sigma_k_dict[active_type].get("3", {}).get("b", 0))
        if ref_ab4 == 0: ref_ab4 = 2
        if ref_ab3 == 0: ref_ab3 = 2
        label_x_ab = half_L + ref_ab4 + OUTSET   # 与面④ l 同 x
        label_y_ab = -half_B - ref_ab3 - OUTSET  # 与面③ l 同 y

        for f_idx in range(1, 5):
            fd = sigma_k_dict[active_type].get(str(f_idx), {})
            p_key = "q0" if active_type == "均布附加荷载" else "p0"
            p = str(fd.get(p_key, "/")).strip()
            if p in ("/", "", "None"): continue
            p_val = _f(p); a = _f(fd.get("a")); b = _f(fd.get("b")); l = _f(fd.get("l"))
            p2_raw = str(fd.get("p2", "/")).strip()
            c_raw = str(fd.get("c", "/")).strip()
            has_dual = (p2_raw not in ("/", "", "0") and c_raw not in ("/", "", "0"))
            p2_val = _f(p2_raw) if has_dual else 0
            c_val = _f(c_raw) if has_dual else 0
            sign_y = 1 if f_idx == 1 else (-1 if f_idx == 3 else 0)
            sign_x = -1 if f_idx == 2 else (1 if f_idx == 4 else 0)

            # 高亮判定
            hl_p = (highlight_key == (f_idx, "p"))
            hl_p2 = (highlight_key == (f_idx, "p2"))

            if f_idx in (1, 3):  # 上下边
                cy = sign_y * half_B
                if active_type == "均布附加荷载":
                    p_color = "#dc3243" if hl_p else "#e74c3c"
                    p_weight = "bold" if hl_p else "normal"
                    self.ax.text(0, cy + sign_y * 1.5, f"{p_val} KPa", ha="center", va="center",
                                 color=p_color, fontsize=8, fontweight=p_weight)
                elif active_type == "矩形局部附加荷载":
                    y_s = cy + sign_y * a
                    y_e = y_s + sign_y * b
                    # 矩形1（p1）
                    self.ax.add_patch(MplRect((-l/2, min(y_s, y_e)), l, b, color="#e67e22", alpha=0.4))
                    # p1 文字
                    p1_color = "#dc3243" if hl_p else "#000000"
                    p1_weight = "bold" if hl_p else "normal"
                    self.ax.text(0, (y_s + y_e) / 2, f"{p_val}KPa", ha="center", va="center",
                                 fontsize=8, color=p1_color, fontweight=p1_weight)
                    if has_dual:
                        # 矩形2（p2）：距矩形1外侧 c 处
                        y2_s = y_e + sign_y * c_val
                        y2_e = y2_s + sign_y * b
                        self.ax.add_patch(MplRect((-l/2, min(y2_s, y2_e)), l, b, color="#e67e22", alpha=0.4))
                        # p2 文字
                        p2_color = "#dc3243" if hl_p2 else "#000000"
                        p2_weight = "bold" if hl_p2 else "normal"
                        self.ax.text(0, (y2_s + y2_e) / 2, f"{p2_val}KPa", ha="center", va="center",
                                     fontsize=8, color=p2_color, fontweight=p2_weight)
                        # 标注：a → b → c → b → l
                        total_out = a + b + c_val + b
                        self._draw_standard_dim(-l/2, l/2, cy + sign_y*(total_out+OUTSET), l, "h",
                                                 highlight=(highlight_key == (f_idx, "l")))
                        self._draw_standard_dim(y_s, y_e, label_x_ab, b, "v",
                                                 highlight=(highlight_key == (f_idx, "b")))
                        self._draw_standard_dim(y_e, y2_s, label_x_ab, c_val, "v",
                                                 highlight=(highlight_key == (f_idx, "c")))
                        self._draw_standard_dim(y2_s, y2_e, label_x_ab, b, "v",
                                                 highlight=(highlight_key == (f_idx, "b")))
                        self._draw_standard_dim(cy, y_s, label_x_ab, a, "v",
                                                 highlight=(highlight_key == (f_idx, "a")))
                    else:
                        # 单矩形标注：a → b → l
                        self._draw_standard_dim(-l/2, l/2, cy + sign_y*(a+b+OUTSET), l, "h",
                                                 highlight=(highlight_key == (f_idx, "l")))
                        self._draw_standard_dim(y_s, y_e, label_x_ab, b, "v",
                                                 highlight=(highlight_key == (f_idx, "b")))
                        self._draw_standard_dim(cy, y_s, label_x_ab, a, "v",
                                                 highlight=(highlight_key == (f_idx, "a")))
                elif active_type == "条形局部附加荷载":
                    y_s = cy + sign_y * a; y_e = y_s + sign_y * b
                    self.ax.add_patch(MplRect((-half_L, min(y_s, y_e)), half_L*2, b, color="#27ae60", alpha=0.4))
                    self.ax.plot([-half_L, -half_L], [min(y_s, y_e)-0.5, max(y_s, y_e)+0.5], color="#1e8449", lw=1.5, linestyle="--")
                    self.ax.plot([half_L, half_L], [min(y_s, y_e)-0.5, max(y_s, y_e)+0.5], color="#1e8449", lw=1.5, linestyle="--")
                    p1_color = "#dc3243" if hl_p else "#000000"
                    p1_weight = "bold" if hl_p else "normal"
                    self.ax.text(0, (y_s + y_e) / 2, f"{p_val}KPa", ha="center", va="center",
                                 fontsize=8, color=p1_color, fontweight=p1_weight)
                    self._draw_standard_dim(y_s, y_e, label_x_ab, b, "v",
                                             highlight=(highlight_key == (f_idx, "b")))
                    self._draw_standard_dim(cy, y_s, label_x_ab, a, "v",
                                             highlight=(highlight_key == (f_idx, "a")))
            else:  # 左右边
                cx = sign_x * half_L
                if active_type == "均布附加荷载":
                    p_color = "#dc3243" if hl_p else "#e74c3c"
                    p_weight = "bold" if hl_p else "normal"
                    self.ax.text(cx + sign_x * 1.5, 0.5, f"{p_val} KPa", ha="center", va="center",
                                 color=p_color, fontsize=8, fontweight=p_weight)
                elif active_type == "矩形局部附加荷载":
                    x_s = cx + sign_x * a
                    x_e = x_s + sign_x * b
                    # 矩形1（p1）
                    self.ax.add_patch(MplRect((min(x_s, x_e), -l/2), b, l, color="#e67e22", alpha=0.4))
                    p1_color = "#dc3243" if hl_p else "#000000"
                    p1_weight = "bold" if hl_p else "normal"
                    self.ax.text((x_s + x_e) / 2, 0, f"{p_val}KPa", ha="center", va="center",
                                 fontsize=8, color=p1_color, fontweight=p1_weight)
                    if has_dual:
                        x2_s = x_e + sign_x * c_val
                        x2_e = x2_s + sign_x * b
                        self.ax.add_patch(MplRect((min(x2_s, x2_e), -l/2), b, l, color="#e67e22", alpha=0.4))
                        p2_color = "#dc3243" if hl_p2 else "#000000"
                        p2_weight = "bold" if hl_p2 else "normal"
                        self.ax.text((x2_s + x2_e) / 2, 0, f"{p2_val}KPa", ha="center", va="center",
                                     fontsize=8, color=p2_color, fontweight=p2_weight)
                        total_out = a + b + c_val + b
                        l_text_side = "left" if f_idx == 2 else "right"
                        self._draw_standard_dim(-l/2, l/2, cx + sign_x*(total_out+OUTSET), l, "v",
                                                 text_side=l_text_side, highlight=(highlight_key == (f_idx, "l")))
                        self._draw_standard_dim(x_s, x_e, label_y_ab, b, "h",
                                                 highlight=(highlight_key == (f_idx, "b")))
                        self._draw_standard_dim(x_e, x2_s, label_y_ab, c_val, "h",
                                                 highlight=(highlight_key == (f_idx, "c")))
                        self._draw_standard_dim(x2_s, x2_e, label_y_ab, b, "h",
                                                 highlight=(highlight_key == (f_idx, "b")))
                        self._draw_standard_dim(cx, x_s, label_y_ab, a, "h",
                                                 highlight=(highlight_key == (f_idx, "a")))
                    else:
                        l_text_side = "left" if f_idx == 2 else "right"
                        self._draw_standard_dim(-l/2, l/2, cx + sign_x*(a+b+OUTSET), l, "v",
                                                 text_side=l_text_side, highlight=(highlight_key == (f_idx, "l")))
                        self._draw_standard_dim(x_s, x_e, label_y_ab, b, "h",
                                                 highlight=(highlight_key == (f_idx, "b")))
                        self._draw_standard_dim(cx, x_s, label_y_ab, a, "h",
                                                 highlight=(highlight_key == (f_idx, "a")))
                elif active_type == "条形局部附加荷载":
                    x_s = cx + sign_x * a; x_e = x_s + sign_x * b
                    self.ax.add_patch(MplRect((min(x_s, x_e), -half_B), b, half_B*2, color="#27ae60", alpha=0.4))
                    self.ax.plot([min(x_s, x_e)-0.5, max(x_s, x_e)+0.5], [-half_B, -half_B], color="#1e8449", lw=1.5, linestyle="--")
                    self.ax.plot([min(x_s, x_e)-0.5, max(x_s, x_e)+0.5], [half_B, half_B], color="#1e8449", lw=1.5, linestyle="--")
                    p1_color = "#dc3243" if hl_p else "#000000"
                    p1_weight = "bold" if hl_p else "normal"
                    self.ax.text((x_s + x_e) / 2, 0, f"{p_val}KPa", ha="center", va="center",
                                 fontsize=8, color=p1_color, fontweight=p1_weight)
                    self._draw_standard_dim(x_s, x_e, label_y_ab, b, "h",
                                             highlight=(highlight_key == (f_idx, "b")))
                    self._draw_standard_dim(cx, x_s, label_y_ab, a, "h",
                                             highlight=(highlight_key == (f_idx, "a")))

        self.ax.set_xlim(-half_L, half_L)
        self.ax.set_ylim(-half_B, half_B)
        self.ax.autoscale(enable=True, tight=False)
        self.ax.margins(0.15)
        self.canvas.draw_idle()


# ═══════════════════════════════════════════════════════════════
# 施工阶段图示
# ═══════════════════════════════════════════════════════════════

class StageDiagram:
    """施工阶段示意图：读取前 3 个 Tab 缓存绘制立面底图，并叠加当前工况内容。

    通过从第 1 个工况逐行累积状态（土顶/水位/支撑/垫层/封底/圈梁等），
    在底图上叠加绘制选中工况的形态与标注。
    """

    PILE_COLOR = "#4a5056"
    GROUND_COLOR = "#8B6914"
    WATER_COLOR = "#2980b9"
    SOIL_TOP_COLOR = "#8B4513"
    WALER_COLOR = "#7f8c8d"
    STRUT_COLOR = "#95a5a6"
    FILL_COLOR = "#3498db"
    SOIL_LAYER_COLORS = ["#f5deb3", "#d2b48c", "#c4a882", "#b89a7a", "#deb887"]
    SOIL_HATCHES = ['', '..', '...', '....', '.....']

    def __init__(self, parent, sheet_refs, stage_controller,
                 waler_section_dict=None, strut_section_dict=None,
                 title_label=None):
        """初始化施工阶段示意图。

        Args:
            parent (tk.Widget): 承载画布的父容器。
            sheet_refs (dict): 前 3 个 Tab 的缓存引用，结构形式
                {'basic_cache': {...}, 'substructure_cache': {...},
                 'soil_frame': <tksheet 控件>, 'assist_replace_data': [...]}。
            stage_controller (object|None): 提供 get_stage_dict() 的阶段控制器。
            waler_section_dict (dict|None): 围檩截面库。
            strut_section_dict (dict|None): 内支撑截面库。
            title_label (tk.Label|None): 用于显示标题的 Label。

        Returns:
            None
        """
        self.parent = parent
        self.sheet_refs = sheet_refs
        self.stage_controller = stage_controller
        self.waler_section_dict = waler_section_dict or {}
        self.strut_section_dict = strut_section_dict or {}
        self.title_label = title_label

        # 累积状态
        self.current_state = {}
        self._last_stage_idx = -1

        self.fig = Figure(figsize=(5.5, 3.0), dpi=100, facecolor='white')
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=parent)
        self.canvas_widget = self.canvas.get_tk_widget()
        # 用 place 精确撑满，避免 pack 的尺寸协商问题
        self.canvas_widget.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.fig.subplots_adjust(left=0.02, right=0.98, top=0.96, bottom=0.04)

    def resize_to_parent(self):
        """由外部调用，将 figsize 同步到父容器像素尺寸（不触发重绘）。

        Returns:
            None
        """
        try:
            w = self.parent.winfo_width()
            h = self.parent.winfo_height()
            if w > 80 and h > 40:
                dpi = self.fig.get_dpi()
                cur_w, cur_h = self.fig.get_size_inches()
                new_w, new_h = w / dpi, h / dpi
                if abs(cur_w - new_w) > 0.1 or abs(cur_h - new_h) > 0.1:
                    self.fig.set_size_inches(new_w, new_h)
        except:
            pass

    # ==================== 公开方法 ====================

    def draw(self, stage_idx=0):
        """主入口：累积到选中行状态并绘制底图、工况叠加与标注。

        Args:
            stage_idx (int): 施工阶段表选中行索引（0 基）。

        Returns:
            None
        """
        if self.stage_controller is None:
            return
        stage_dict = self.stage_controller.get_stage_dict()
        if not stage_dict or stage_idx < 0:
            self.ax.clear()
            self.ax.text(0.5, 0.5, "请选择施工阶段", ha='center', va='center',
                         transform=self.ax.transAxes, fontsize=10, color='gray')
            if self.title_label is not None:
                try: self.title_label.config(text="施工阶段图示")
                except: pass
            self.canvas.draw_idle()
            return

        # 获取当前行数据
        row_data = stage_dict.get(stage_idx + 1, {})
        stage_type = row_data.get("工况类型", "")

        # 更新累积状态（从1到当前行）
        self._update_state(stage_dict, stage_idx)

        # 绘制
        self.ax.clear()
        self._draw_base()
        self._draw_state(stage_type)
        self._draw_labels(row_data)

        # 更新标题（tkinter Label）
        title_text = f"施工阶段{stage_idx + 1}图示：{stage_type}"
        if self.title_label is not None:
            try: self.title_label.config(text=title_text)
            except: pass

        # 坐标系（按实际尺寸等比例显示）
        g = self._get_geometry()
        y_margin = (g['pile_top'] - g['pile_bottom']) * 0.06
        self.ax.set_xlim(-g['x_range'], g['x_range'])
        self.ax.set_ylim(g['pile_bottom'] - y_margin, g['pile_top'] + y_margin + 1.0)
        self.ax.set_aspect('equal', adjustable='box')
        self.ax.axis('off')
        self.canvas.draw_idle()

    # ==================== 状态累积 ====================

    def _update_state(self, stage_dict, target_idx):
        """从第 1 个工况开始逐个应用到 target_idx，累积 current_state。

        Args:
            stage_dict (dict): 施工阶段字典 {序号: {'工况类型','基坑内土顶标高',
                '基坑内水面标高','支撑层数'}}。
            target_idx (int): 目标行索引（0 基）。

        Returns:
            None: 结果写入 self.current_state。
        """
        state = {
            "soil_top": None,
            "water": None,
            "supports": {},
            "added_supports": {},
            "replaced_names": set(),
            "has_blinding": False,
            "has_plug": False,
            "fill_rect": None,
        }

        for i in range(1, target_idx + 2):
            sd = stage_dict.get(i, {})
            stype = sd.get("工况类型", "")
            if not stype:
                continue

            soil_top = sd.get("基坑内土顶标高", "/")
            water = sd.get("基坑内水面标高", "/")
            support_val = sd.get("支撑层数", "/")

            try:
                state["soil_top"] = float(soil_top) if soil_top not in ("", "/") else state.get("soil_top")
            except: pass
            try:
                state["water"] = float(water) if water not in ("", "/") else state.get("water")
            except: pass

            if stype == "加撑":
                layer = support_val
                if layer not in ("", "/"):
                    sec, elev = self._get_support_section(layer)
                    state["supports"][layer] = {"sec": sec, "elev": elev, "visible": True}

            elif stype == "封底":
                state["has_plug"] = True

            elif stype in ("垫层", "围檩"):
                state["has_blinding"] = True

            elif stype == "辅助加圈梁":
                cap_bottom = self._get_cache_float("basic_cache", "cap_bottom_level", 0.0)
                state["fill_rect"] = {"top": cap_bottom, "bottom": state.get("soil_top", cap_bottom)}

            elif stype == "辅助拆撑":
                layer = support_val
                if layer in state["supports"]:
                    state["supports"][layer]["visible"] = False

            elif stype == "辅助换撑":
                name = support_val
                if name not in ("", "/"):
                    state["replaced_names"].add(name)

            elif stype == "辅助加撑":
                layer = support_val
                if layer not in ("", "/"):
                    sec, elev = self._get_support_section(layer)
                    state["added_supports"][layer] = {"sec": sec, "elev": elev}

        self.current_state = state
        self._last_stage_idx = target_idx

    # ==================== 基础底图 ====================

    def _draw_base(self):
        """绘制施工阶段示意图底图：中心线、桩、土顶线、承台底线与土层色块。

        Returns:
            None
        """
        ax = self.ax
        g = self._get_geometry()

        # 中心线（洋红虚线，置顶显示）
        ax.axvline(0, color='magenta', ls='--', lw=0.8, alpha=0.6, zorder=999)

        # 左侧/右侧桩（矩形）
        for side in ['left', 'right']:
            x = g[f'pile_{side}_inner']
            w = abs(g[f'pile_{side}_outer'] - g[f'pile_{side}_inner'])
            x0 = min(x, g[f'pile_{side}_outer'])
            ax.add_patch(Rectangle(
                (x0, g['pile_bottom']), w, g['pile_top'] - g['pile_bottom'],
                facecolor=self.PILE_COLOR, lw=0.8, zorder=5, alpha=0.85))

        # 土顶线（三段：桩外左/桩内/桩外右）
        s = self.current_state
        has_excavated = s.get("soil_top") is not None
        li, ri = g['pile_left_inner'], g['pile_right_inner']
        lo, ro = g['pile_left_outer'], g['pile_right_outer']
        # 土顶线不延伸至 x_range 尽头，留出标注空间
        gl_end = ro + 1.5
        if has_excavated:
            ax.plot([-gl_end, lo], [g['ground'], g['ground']],
                    color=self.GROUND_COLOR, lw=1.5, zorder=6)
            ax.plot([ro, gl_end], [g['ground'], g['ground']],
                    color=self.GROUND_COLOR, lw=1.5, zorder=6)
        else:
            ax.plot([-gl_end, lo], [g['ground'], g['ground']],
                    color=self.GROUND_COLOR, lw=1.5, zorder=6)
            ax.plot([li, ri], [g['ground'], g['ground']],
                    color=self.GROUND_COLOR, lw=1.5, zorder=6)
            ax.plot([ro, gl_end], [g['ground'], g['ground']],
                    color=self.GROUND_COLOR, lw=1.5, zorder=6)

        # 承台底土层线（桩内侧通长，常态始终显示）
        ax.plot([li, ri], [g['cap_bottom'], g['cap_bottom']],
                color=self.GROUND_COLOR, lw=1.0, ls='--', alpha=0.7, zorder=6)

        # 土层色块（窄条，延至桩底）
        soil_info = self._get_soil_info()
        y = g['ground']
        for i, layer in enumerate(soil_info):
            thickness = layer['thickness']
            y_bot = y - thickness
            color = self.SOIL_LAYER_COLORS[i % len(self.SOIL_LAYER_COLORS)]
            hatch = self.SOIL_HATCHES[i % len(self.SOIL_HATCHES)]
            ax.fill_between([g['pile_left_outer'] - 1.2, g['pile_left_outer']],
                            y_bot, y, facecolor=color, edgecolor='#a0a5aa', lw=0.3, zorder=2)
            ax.fill_between([g['pile_left_outer'] - 1.2, g['pile_left_outer']],
                            y_bot, y, facecolor='none', hatch=hatch,
                            edgecolor="#b0b5ba", lw=0, alpha=0.3, zorder=2)
            y = y_bot

    # ==================== 工况叠加 ====================

    def _draw_state(self, current_stage_type=""):
        """叠加绘制当前累积状态的工况内容（水位、土顶、承台、垫层/封底、支撑、圈梁）。

        Args:
            current_stage_type (str): 当前工况类型，用于高亮封底/垫层。

        Returns:
            None
        """
        ax = self.ax
        s = self.current_state
        g = self._get_geometry()
        if not s:
            return

        li, ri = g['pile_left_inner'], g['pile_right_inner']

        # 基坑内水面线（仅桩内侧）
        if s.get("water") is not None:
            ax.plot([li, ri], [s["water"], s["water"]],
                    color=self.WATER_COLOR, lw=1.2, ls='--',
                    alpha=0.7, zorder=7)

        # 基坑内土顶线（仅桩内侧，虚线）
        if s.get("soil_top") is not None:
            ax.plot([li, ri], [s["soil_top"], s["soil_top"]],
                    color=self.SOIL_TOP_COLOR, lw=1.2, ls='-.',
                    alpha=0.7, zorder=7)

        # 承台（居中于承台长度，与详细立面图一致）
        ax.add_patch(Rectangle(
            (g['cap_left'], g['cap_bottom']),
            g['cap_right'] - g['cap_left'],
            g['cap_top'] - g['cap_bottom'],
            facecolor='#d5d8dc', lw=1.0, alpha=0.85, zorder=8))

        # 垫层/封底（桩内侧通长，底在基坑底，高=垫层/封底厚）
        # 当前工况为封底/垫层时红色高亮，后续工况灰色
        if s.get("has_blinding") or s.get("has_plug"):
            concrete_bottom = g.get('concrete_bottom', g['cap_bottom'])
            is_current_plug = current_stage_type in ("封底", "垫层")
            ax.add_patch(Rectangle(
                (li, concrete_bottom), ri - li,
                g['cap_bottom'] - concrete_bottom,
                facecolor='#e74c3c' if is_current_plug else '#aab7b8',
                lw=2.0 if is_current_plug else 0.8,
                alpha=0.8 if is_current_plug else 0.7, zorder=8))

        # 支撑
        for layer, info in s.get("supports", {}).items():
            if info.get("visible"):
                self._draw_support_beam(layer, info)

        # 辅助加撑
        for layer, info in s.get("added_supports", {}).items():
            self._draw_support_beam(layer, info, color='#2980b9')

        # 辅助加圈梁填充区
        fill = s.get("fill_rect")
        if fill and fill.get("top") is not None and fill.get("bottom") is not None:
            g = self._get_geometry()
            w = g['pile_right_inner'] - g['pile_left_inner']
            ax.add_patch(Rectangle(
                (g['pile_left_inner'], fill['bottom']), w, fill['top'] - fill['bottom'],
                facecolor=self.FILL_COLOR, alpha=0.15, lw=0, zorder=3))

    def _draw_support_beam(self, layer, info, color=None):
        """绘制单道支撑（截面矩形 + 中心线）。

        Args:
            layer (str|int): 支撑层号。
            info (dict): 支撑信息 {'sec': list, 'elev': float}。
            color (str|None): 指定颜色，默认使用围檩色。

        Returns:
            None
        """
        ax = self.ax
        g = self._get_geometry()
        elev = info.get("elev", 0)
        sec = info.get("sec", [])
        h = float(sec[0]) / 1000.0 if sec and len(sec) > 0 else 0.3
        c = color or self.WALER_COLOR
        left = g['pile_left_inner']
        right = g['pile_right_inner']
        ax.add_patch(Rectangle(
            (left, elev - h / 2), right - left, h,
            facecolor=c, lw=0.8, alpha=0.8, zorder=9))
        ax.plot([left, right], [elev, elev], color='white', ls='-.', lw=0.6, zorder=10, alpha=0.7)

    def _draw_replaced_highlight(self, name):
        """高亮辅助换撑层（红色矩形 + 右侧层号标注）。

        Args:
            name (str): 换撑名称，如 "1-1"（层号-换撑号）。

        Returns:
            None
        """
        ax = self.ax
        if not self.sheet_refs:
            return
        g = self._get_geometry()
        gl_end = g['pile_right_outer'] + 1.5
        label_x = gl_end + 1.2
        for item in self.sheet_refs.get("assist_replace_data", []):
            if item.get("name") == name:
                # 从 name 解析围檩层号：name格式如 "1-1"、"2-1" → wl=1,2
                wl_str = item.get("waler_layer", "")
                if not wl_str:
                    # 从 name 字段解析：取 "-" 前的数字
                    wl_str = name.split("-")[0] if "-" in name else ""
                try:
                    wl = int(wl_str)
                except: continue
                if str(wl) in self.current_state.get("supports", {}):
                    info = self.current_state["supports"][str(wl)]
                    elev = info.get("elev", 0)
                    sec = info.get("sec", [])
                    h = float(sec[0]) / 1000.0 if sec else 0.3
                    left, right = g['pile_left_inner'], g['pile_right_inner']
                    ax.add_patch(Rectangle(
                        (left, elev - h / 2), right - left, h,
                        facecolor='#e74c3c', lw=2, alpha=0.6, zorder=12))
                    ax.annotate(f"换撑 {name}",
                                xy=(right, elev), xytext=(label_x, elev),
                                fontsize=7, color='#e74c3c', fontweight='bold',
                                va='center', ha='left',
                                arrowprops=dict(arrowstyle='->', color='#e74c3c',
                                                lw=0.5, alpha=0.7))

    # ==================== 标注 ====================

    @staticmethod
    def _fmt(val, unit="m", decimals=3):
        """格式化标高：保留 decimals 位小数，正值加 + 号并附单位。

        Args:
            val: 标高校值。
            unit (str): 单位后缀。
            decimals (int): 小数位数。

        Returns:
            str: 格式化后的字符串；非法值原样转字符串。
        """
        try:
            f = float(val)
            s = f"{f:+.{decimals}f}" if f > 0 else f"{f:.{decimals}f}"
            return f"{s}{unit}"
        except:
            return str(val)

    def _draw_labels(self, row_data):
        """仅标注当前选中工况（土顶/水位/支撑/圈梁/换撑等）。

        Args:
            row_data (dict): 当前工况行数据 {'工况类型','支撑层数',...}。

        Returns:
            None
        """
        stage_type = row_data.get("工况类型", "")
        ax = self.ax
        g = self._get_geometry()
        # 标注 x 坐标：三角/文字全部在地面线右侧
        gl_end = g['pile_right_outer'] + 1.5          # 地面线右端
        marker_x = gl_end + 0.3                       # 三角符号位置
        label_x = gl_end + 1.2                        # 文字位置

        if stage_type == "取土开挖":
            soil_top = self.current_state.get("soil_top")
            if soil_top is not None:
                try:
                    from General.Matplotlib import draw_elevation_marker as dem
                    dem(ax, marker_x, soil_top,
                        f"基坑内土顶标高 {self._fmt(soil_top)}",
                        text_side='right', marker_side='above',
                        scale=0.7, fontsize=7, color=self.SOIL_TOP_COLOR)
                except: pass
            water = self.current_state.get("water")
            if water is not None:
                try:
                    from General.Matplotlib import draw_elevation_marker as dem
                    dem(ax, marker_x, water,
                        f"基坑内水面标高 {self._fmt(water)}",
                        text_side='right', marker_side='below',
                        scale=0.7, fontsize=7, color=self.WATER_COLOR)
                except: pass

        elif stage_type == "加撑":
            layer = row_data.get("支撑层数", "")
            if layer not in ("", "/") and layer in self.current_state.get("supports", {}):
                info = self.current_state["supports"][layer]
                elev = info.get("elev", 0)
                ax.annotate(f"第{layer}层支撑  {self._fmt(elev)}",
                            xy=(g['pile_right_inner'], elev),
                            xytext=(label_x, elev),
                            fontsize=7, color=self.WALER_COLOR,
                            va='center', ha='left',
                            arrowprops=dict(arrowstyle='->', color=self.WALER_COLOR,
                                            lw=0.5, alpha=0.7))

        elif stage_type == "降水":
            water = self.current_state.get("water")
            if water is not None:
                try:
                    from General.Matplotlib import draw_elevation_marker as dem
                    dem(ax, marker_x, water,
                        f"降水后水位 {self._fmt(water)}",
                        text_side='right', marker_side='below',
                        scale=0.7, fontsize=7, color=self.WATER_COLOR)
                except: pass

        elif stage_type == "辅助加圈梁":
            soil_top = self.current_state.get("soil_top")
            if soil_top is not None:
                ax.annotate(f"基坑内土顶标高 {self._fmt(soil_top)}",
                            xy=(g['pile_right_inner'], soil_top),
                            xytext=(label_x, soil_top),
                            fontsize=7, color=self.FILL_COLOR,
                            va='center', ha='left',
                            arrowprops=dict(arrowstyle='->', color=self.FILL_COLOR,
                                            lw=0.5, alpha=0.7))

        elif stage_type == "辅助加撑":
            layer = row_data.get("支撑层数", "")
            if layer not in ("", "/") and layer in self.current_state.get("added_supports", {}):
                info = self.current_state["added_supports"][layer]
                elev = info.get("elev", 0)
                ax.annotate(f"辅助加撑第{layer}层  {self._fmt(elev)}",
                            xy=(g['pile_right_inner'], elev),
                            xytext=(label_x, elev),
                            fontsize=7, color='#2980b9',
                            va='center', ha='left',
                            arrowprops=dict(arrowstyle='->', color='#2980b9',
                                            lw=0.5, alpha=0.7))

        elif stage_type == "辅助换撑":
            name = row_data.get("支撑层数", "")
            if name not in ("", "/"):
                self._draw_replaced_highlight(name)

    # ==================== 辅助方法 ====================

    def _get_geometry(self):
        """计算底图几何坐标（与详细立面图 RealTimeDiagram 一致）。

        Returns:
            dict: 结构形式
                {'pile_center','pile_half_w','pile_left_inner','pile_left_outer',
                 'pile_right_inner','pile_right_outer','pile_top','pile_bottom',
                 'ground','cap_bottom','cap_top','cap_left','cap_right',
                 'concrete_bottom','x_range'}，单位 m。
        """
        # 从 basic_cache 读取
        bc = self.sheet_refs.get("basic_cache", {}) if self.sheet_refs else {}
        def _cf(key, default=0.0):
            """从 basic_cache 读取浮点值，失败返回 default。"""
            try: return float(str(bc.get(key, str(default))).strip() or str(default))
            except: return default

        cap_length = _cf("cap_length", 9.5)           # 承台长
        cap_offset_long = _cf("cap_offset_long", 0.5) # 承台长边预留边距
        cap_bottom = _cf("cap_bottom_level", -3.5)    # 承台底标高
        cap_height = _cf("cap_height", 2.0)           # 承台高
        pile_top = _cf("pile_top_level", 1.0)         # 桩顶标高
        pile_length = _cf("pile_length", 9.0)         # 桩长
        ground = _cf("ground_level", 0.5)             # 地面标高

        cap_half = cap_length / 2.0
        pile_half_w = 0.25
        pile_center = cap_half + cap_offset_long
        left_inner = -pile_center + pile_half_w
        left_outer = -pile_center - pile_half_w
        right_inner = pile_center - pile_half_w
        right_outer = pile_center + pile_half_w
        pile_bottom = pile_top - pile_length
        cap_top = cap_bottom + cap_height

        # 垫层/封底（读取 substructure_cache）
        sub = self.sheet_refs.get("substructure_cache", {}) if self.sheet_refs else {}
        try:
            bt = float(str(sub.get("blinding_thickness", "0")).strip() or "0")
        except: bt = 0.0
        try:
            pt = float(str(sub.get("plug_thickness", "0")).strip() or "0")
        except: pt = 0.0
        concrete_bottom = cap_bottom - (bt if bt > 0 else pt if pt > 0 else 0.0)

        x_range = max(abs(left_outer), abs(right_outer)) + 3.0
        return {
            'pile_center': pile_center, 'pile_half_w': pile_half_w,
            'pile_left_inner': left_inner, 'pile_left_outer': left_outer,
            'pile_right_inner': right_inner, 'pile_right_outer': right_outer,
            'pile_top': pile_top, 'pile_bottom': pile_bottom,
            'ground': ground,
            'cap_bottom': cap_bottom, 'cap_top': cap_top,
            'cap_left': -cap_half, 'cap_right': cap_half,
            'concrete_bottom': concrete_bottom,
            'x_range': x_range,
        }

    def _get_soil_info(self):
        """从 tksheet 实时读取土层数据（非缓存快照）。

        Returns:
            list[dict]: [{'name': str, 'thickness': float}, ...]；
                无有效数据时返回默认单层 [{'name': '土', 'thickness': 10.0}]。
        """
        layers = []
        if self.sheet_refs:
            soil_frame = self.sheet_refs.get("soil_frame")
            if soil_frame and hasattr(soil_frame, "sheet"):
                try:
                    raw = soil_frame.sheet.get_sheet_data()
                    for row in raw:
                        if not row or not row[0]:
                            continue
                        name = str(row[0]).strip()
                        if not name:
                            continue
                        # 层厚在 index 1
                        try:
                            t = float(str(row[1]).strip() or "0")
                            if t > 0:
                                layers.append({'name': name, 'thickness': t})
                        except:
                            pass
                except Exception:
                    pass
        if not layers:
            layers = [{'name': '土', 'thickness': 10.0}]
        return layers

    def _get_support_section(self, layer):
        """获取指定支撑层的截面信息与标高（从缓存的 waler_rows 中查找）。

        Args:
            layer (str|int): 支撑层号（1 基）。

        Returns:
            tuple[list, float]: (截面 SEC 列表, 标高 m)；无效时返回 ([], 0.0)。
        """
        if not self.sheet_refs:
            return [], 0.0
        sub = self.sheet_refs.get("substructure_cache", {})
        waler_rows = sub.get("waler_rows", [])
        try:
            idx = int(layer) - 1
        except: return [], 0.0
        if 0 <= idx < len(waler_rows):
            try:
                # 间距累减计算标高
                pile_top = self._get_geometry()['pile_top']
                current_elev = pile_top
                for i in range(idx + 1):
                    sp = float(str(waler_rows[i].get("spacing", "0")).strip() or "0")
                    current_elev -= sp
                elev = current_elev
                sec_name = str(waler_rows[idx].get("sec_long", ""))
                if sec_name in self.waler_section_dict:
                    sec = self.waler_section_dict[sec_name].get('SEC', [])
                    return sec, elev
                return [], elev
            except: pass
        return [], 0.0

    def _get_cache_float(self, cache_key, field, default=0.0):
        """从 sheet_refs 的指定缓存字典中读取浮点值。

        Args:
            cache_key (str): 缓存键（如 'basic_cache'）。
            field (str): 字段名。
            default (float): 读取失败时的默认值。

        Returns:
            float: 读取结果或 default。
        """
        if not self.sheet_refs:
            return default
        cache = self.sheet_refs.get(cache_key, {})
        val = cache.get(field, str(default))
        try:
            return float(str(val).strip())
        except:
            return default