# 1. 标准库
import os
import sys
import copy
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from pathlib import Path

# 2. 第三方库
from tksheet import Sheet
import openpyxl
import ttkbootstrap as tb
from ttkbootstrap.constants import *
import matplotlib
matplotlib.use('TkAgg')
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

# 3. 本地模块
sys.path.append(str(Path(__file__).parent.parent.parent))
from General.UIHandle import (PileBoundaryFrame, ConnectionSettingFrame, BeamReleaseFrame,
                               create_label_entry, SplitEntry, make_card_frame,
                               create_scrollable_frame, create_section_selector)

from General.DataUtils import num_to_chinese_num
from General.DataUtils import midasdisttolst
from General.Geometry import polyline_y_at_x
from Trestle_Excel_io_FEM import (get_default_support_data,
                                   get_properties_path, load_section_library, _parse_material_value,
                                   _sec_dim, _norm_sec_name)
from Trestle_Generate_FEM import TRUSS_BEAM_TYPES
from Trestle_Widgets_FEM import SupportSetupDialog
import Trestle_Geology_FEM



trestle_excel_dict: dict = {}

from Trestle_Widgets_FEM import (draw_support_transverse_view, 
                                  draw_support_longitudinal_view,
                                  draw_support_elevation, draw_components_transverse_diagram,
                                  draw_components_plan_diagram, draw_move_load_diagram,
                                  draw_static_load_diagram)

# 指定中文字体
matplotlib.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei']
matplotlib.rcParams['axes.unicode_minus'] = False


# 绘图配色
DRAW_LINE_COLOR = "#2b6fd4"
DRAW_TEXT_COLOR = "#2c3e50"
DRAW_TEXT_HIGHLIGHT = "#dc3243"


# ═══════════════════════════════════════════════════════════════
# Tab 页面（包含所有 9 个设置 Tab + 截面选择 + 下部结构弹窗）
# ═══════════════════════════════════════════════════════════════

class SupportsTab(ttk.Frame):
    def __init__(self, parent, parent_ui=None):
        super().__init__(parent)
        self.parent_ui = parent_ui
        # 统一样式参数 
        self.padx = 5
        self.pady = 15
        self.lbl_w = 15  
        self.ent_w = 40

        # 变量设置
        self.support_data = []
        self.table_cells = []
        self.current_row = None
        self._is_updating_ui = False
        
        # 按顺序初始化各个 UI 区域
        self.basic_setting()
        self.plot()
        self.table()

        self.substructure_cache = {}   # 下部结构缓存 {idx: (type, [params], [soil])}
        self.basic_cache = {}          # 基本参数缓存 {span, water_level, deck_level}
        self.soil_data_cache = {}    # 缓存编辑过的新土层数据: {"项目编号-桩号": {...}}
        self.soil_pile_mapping = {}  # 缓存桩号对应的标签映射: {"1": "1-0"}（无 # 后缀）
        self.generate_topology()

    def reload_from_dict(self, new_dict):
        """从新项目数据刷新 UI（不重建 widget）"""
        basic = new_dict.get("basic", {})
        for ent, key, alt_key in [(self.span, "span", None), (self.bridge_width, "bridge_width", None),
                                   (self.water_level, "water_level", None), (self.deck_level, "deck_level", None)]:
            ent.delete(0, tk.END)
            val = basic.get(key, "")
            if not val and alt_key:
                val = basic.get(alt_key, "")
            ent.insert(0, val)
        self.substructure_cache.clear()
        self.basic_cache.clear()
        self.soil_data_cache.clear()
        self.soil_pile_mapping.clear()
        # fresh=True：全部从新项目数据重建，不复用旧行（否则旧项目行数据污染新项目前 N 行）
        self.generate_topology(fresh=True)

    def basic_setting(self):
        """ 栈桥基本设置参数"""
        basic_card, basic_content = make_card_frame(self, "栈桥基本设置")

        # 计算跨径和计算宽度
        ttk.Label(basic_content, text="计算跨径(m):", width=15, bootstyle=PRIMARY).grid(row=0, column=0, sticky="w", padx=self.padx, pady=self.pady)
        self.span = ttk.Entry(basic_content, width=20)
        self.span.grid(row=0, column=1, sticky="w", padx=(0, 20), pady=self.pady)
        self.span.insert(0, trestle_excel_dict["basic"]["span"])
        self.span._last_valid = self.span.get() or "9"
        self.span.bind("<Return>", lambda e: self.generate_topology())
        self.span.bind("<FocusOut>", lambda e: self._on_span_width_change())
        ttk.Label(basic_content, text="计算宽度(m):", width=18, bootstyle=PRIMARY).grid(row=0, column=2, sticky="e", padx=(10, 5), pady=self.pady)
        self.bridge_width = ttk.Entry(basic_content, width=20)
        self.bridge_width.grid(row=0, column=3, sticky="w", padx=(0, self.padx), pady=self.pady)
        self.bridge_width.insert(0, trestle_excel_dict["basic"]["bridge_width"])
        self.bridge_width._last_valid = self.bridge_width.get() or "8"
        self.bridge_width.bind("<FocusOut>", lambda e: self._on_span_width_change())

        # 设防水位和桥面高程
        ttk.Label(basic_content, text="设防水位(m):", width=15, bootstyle=PRIMARY).grid(row=1, column=0, sticky="w", padx=self.padx, pady=self.pady)
        self.water_level = ttk.Entry(basic_content, width=20)
        self.water_level.grid(row=1, column=1, sticky="w", padx=(0, 20), pady=self.pady)
        self.water_level.insert(0, trestle_excel_dict["basic"]["water_level"])
        self.water_level.bind("<KeyRelease>", self.on_water_change)
        self.water_level.bind("<FocusOut>", lambda e, s=self: s.refresh_cache())

        ttk.Label(basic_content, text="桥面高程(m):", width=18, bootstyle=PRIMARY).grid(row=1, column=2, sticky="e", padx=(10, 5), pady=self.pady)
        self.deck_level = ttk.Entry(basic_content, width=20)
        self.deck_level.grid(row=1, column=3, sticky="w", padx=(0, self.padx), pady=self.pady)
        self.deck_level.insert(0, trestle_excel_dict["basic"]["deck_level"])
        self.deck_level.bind("<FocusOut>", lambda e, s=self: s.refresh_cache())

        basic_content.columnconfigure(0, weight=0)
        basic_content.columnconfigure(1, weight=0)
        basic_content.columnconfigure(2, weight=1)
        basic_content.columnconfigure(3, weight=0)

    def plot(self):
        """ 图示区 (居中标题+边框) """
        plot_container = tk.Frame(self, bg="white", highlightbackground="#ccd1d8", highlightthickness=1)
        plot_container.pack(anchor="w", padx=10, pady=5, fill="x")
        title_lbl = tk.Label(plot_container, text="栈桥立面示意图", font=("Microsoft YaHei UI", 9, "bold"),
                             bg="white", fg="#000000")
        title_lbl.pack(anchor="center", pady=(10, 3))
        sep_line = ttk.Separator(plot_container, orient=HORIZONTAL)
        sep_line.pack(fill="x", padx=20, pady=(0, 5))
        plot_inner = ttk.Frame(plot_container)
        plot_inner.pack(fill="both", expand=True, padx=5, pady=5)
        self.fig = Figure(figsize=(6.5, 2), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=plot_inner)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

    def table(self):
        """ 表格区 (单表视图 + 底部固定操作栏) """
        table_container = ttk.Frame(self)
        table_container.pack(fill="both", expand=True, padx=10, pady=5)
        ttk.Label(table_container, text="下部结构参数表", font=("Microsoft YaHei UI", 10, "bold"),
                  bootstyle=PRIMARY).pack(anchor=W, padx=15, pady=(10, 5))
        table_inner_frame = ttk.Frame(table_container)
        table_inner_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # ====== 新增：底部固定操作栏 ======
        self.action_frame = ttk.Frame(table_container)
        self.action_frame.pack(side="bottom", fill="x", pady=(10, 0))

        self.center_frame = ttk.Frame(self.action_frame)
        self.center_frame.pack(expand=True, pady=5) 

        # 动态按钮文本变量
        self.btn_param_text = tk.StringVar(value="请先在上方选中支撑...")
        self.btn_soil_text = tk.StringVar(value="请先在上方选中支撑...")

        # 结构设置按钮
        self.btn_param = ttk.Button(self.center_frame, textvariable=self.btn_param_text, bootstyle="primary", width=25, state="normal", command=lambda: self.open_support_setup(self.current_row) if self.current_row is not None else None)
        self.btn_param.pack(side="left", padx=(0, 20))

        # ====== 常规的滚动条与单画布 ======
        self.v_scroll = ttk.Scrollbar(table_inner_frame, orient="vertical")
        self.v_scroll.pack(side="right", fill="y")
        self.h_scroll = ttk.Scrollbar(table_inner_frame, orient="horizontal")
        self.h_scroll.pack(side="bottom", fill="x")

        self.table_canvas = tk.Canvas(table_inner_frame, borderwidth=0, highlightthickness=0, height=180)
        self.table_canvas.pack(side="left", fill="both", expand=True)

        def _clamped_scrollbar_set(*args):
            lo, hi = float(args[0]), float(args[1])
            if lo < 0: lo = 0
            if hi > 1: hi = 1
            self.v_scroll.set(lo, hi)
        self.table_canvas.configure(yscrollcommand=_clamped_scrollbar_set, xscrollcommand=self.h_scroll.set)
        self.v_scroll.config(command=self.table_canvas.yview)
        self.h_scroll.config(command=self.table_canvas.xview)

        self.table_inner = tk.Frame(self.table_canvas)
        self.table_inner_id = self.table_canvas.create_window((0, 0), window=self.table_inner, anchor="nw")
        self.table_inner.bind("<Configure>", lambda e: self.table_canvas.configure(scrollregion=self.table_canvas.bbox("all")))

    def generate_topology(self, event=None, fresh=False):
        """重建下部结构表格

        fresh=True（切换项目时）：全部从项目数据重建，不复用旧行 ——
        否则新项目行数 ≥ 旧项目时前 N 行会被旧项目数据污染（B 覆盖 A bug）。
        fresh=False（用户改跨径/改类型时）：保留已有行数据，仅新增/删除行。
        """
        span_str = self.span.get()
        self.spans_data = midasdisttolst(span_str)
        if not self.spans_data: return
        support_count = len(self.spans_data) + 1

        old_data = list(self.support_data)
        old_count = len(old_data)

        # 清空旧表格
        for widget in self.table_inner.winfo_children(): widget.destroy()
        self.support_data.clear()
        self.table_cells.clear()

        self.columns_keys = ["pile_type",  "pile_sec", "pile_material", "pile_len", "pile_trans_space","pile_long_space", "expansion_len", "dist_trans_sec", "dist_trans_material", "dist_trans_len", "dist_long_sec", "dist_long_material", "dist_long_len", "brace_form", "brace_sec", "brace_material", "brace_sec_tilt", "brace_space", "brace_height"]

        headers = ["下部结构编号", "结构类型", "桩类型", "桩截面", "桩材质", "桩长(m)","桩横向布置(mm)", "桩纵向布置(mm)", "伸缩缝长度(mm)", "分配梁(横)截面", "分配梁(横)材质", "分配梁(横)长(mm)", "分配梁(纵)截面", "分配梁(纵)材质", "分配梁(纵)长(mm)", "联结系形式", "联结系截面", "联结系材质", "联结系斜杆截面", "间距(mm)", "层高(mm)"]
        
        # 绘制表头
        for col, text in enumerate(headers):
            header_bg = "#e8f4fd" if col < 2 else "#f2f2f2"
            lbl = tk.Label(self.table_inner, text=text, font=("微软雅黑", 9, "bold"), bg=header_bg, relief="solid", borderwidth=1, padx=8, pady=5)
            lbl.grid(row=0, column=col, sticky="nsew")

        # 填充行数据
        for i in range(support_count):
            row_idx = i + 1
            if fresh:
                # 切项目：全部从当前项目数据（cache 格式）重建，不复用旧行
                init_type, row_data = self._row_from_project_data(i, support_count)
            elif i < old_count:
                row_data = old_data[i].copy()
                init_type = row_data["type"]
                # 如果原先是桥台（端点），但由于跨径增加变成了中间墩，需要强制变更为单排桩并初始化默认参数
                is_end = (i == 0 or i == support_count - 1)
                if init_type in ["重力式桥台", "简易桥台"] and not is_end:
                    init_type = "单排桩"
                    row_data = self._default_support(init_type)
            else:
                # 新增支撑：优先取项目数据中已存在的行，否则按位置取默认模板
                init_type, row_data = self._row_from_project_data(i, support_count)
                
            self.support_data.append(row_data)
            self._auto_register_sections(row_data)

            lbl_text = f"{i}# (左端):" if i == 0 else f"{i}# (右端):" if i == support_count - 1 else f"{i}#:"
            row_cells = {}

            # 编号
            lbl_no = tk.Label(self.table_inner, text=lbl_text, relief="solid", borderwidth=1, padx=5, bg="white")
            lbl_no.grid(row=row_idx, column=0, sticky="nsew")
            lbl_no.bind("<Button-1>", lambda e, idx=i: self.select_row(idx))
            
            # 类型下拉框
            type_frame = tk.Frame(self.table_inner, relief="solid", borderwidth=1, bg="white")
            type_frame.grid(row=row_idx, column=1, sticky="nsew")
            options = ["单排桩", "制动墩", "重力式桥台", "简易桥台"] if (i==0 or i==support_count-1) else ["单排桩", "制动墩"]
            cb_var = tk.StringVar(value=init_type)
            cb = ttk.Combobox(type_frame, textvariable=cb_var, values=options, state="readonly", width=10)
            cb.pack(padx=2, pady=4, fill="both", expand=True)
            cb.bind("<<ComboboxSelected>>", lambda e, idx=i, v=cb_var: self.on_type_change(idx, v.get()))
            type_frame.bind("<Button-1>", lambda e, idx=i: self.select_row(idx))
            cb.bind("<Button-1>", lambda e, idx=i: self.select_row(idx), add="+")
            cb.bind("<FocusIn>", lambda e, idx=i: self.select_row(idx), add="+")
            
            # 数据列
            for col_offset, key in enumerate(self.columns_keys):
                val = self.support_data[i][key]
                lbl = tk.Label(self.table_inner, text=val, relief="solid", borderwidth=1, padx=5, bg="#fcfcfc", fg="#666666")
                lbl.grid(row=row_idx, column=col_offset + 2, sticky="nsew")
                lbl.bind("<Button-1>", lambda e, idx=i: self.select_row(idx))
                row_cells[key] = lbl
                
            self.table_cells.append(row_cells)

        self.select_row(0)
        self.after_idle(lambda: draw_support_elevation(self, 0, trestle_excel_dict=trestle_excel_dict))
        self.refresh_cache()

    def _default_support(self, sup_type):
        """创建默认支撑数据：get_default_support_data + 填充默认材质（避免材质为空）

        返回的支撑数据材质字段会从 steel_dict 取默认值（规范编号,牌号），
        保证用户不点开二级窗口修改时，材质也不为空。
        """
        data = get_default_support_data(sup_type).copy()
        steel_dict = trestle_excel_dict.get("steel_dict", {})
        if steel_dict:
            first_spec = next(iter(steel_dict.values()))
            spec_id = first_spec.get("钢结构规范编号", "")
            brands = list(first_spec.get("牌号参数", {}).keys())
            default_brand = brands[0] if brands else "Q235"
            default_val = f"{spec_id},{default_brand}" if spec_id else default_brand
            for k in ("pile_material", "dist_trans_material", "dist_long_material", "brace_material"):
                if data.get(k) in (None, ""):
                    data[k] = default_val
        return data

    def _row_from_project_data(self, idx, support_count):
        """从当前项目数据取第 idx 根支撑的完整行（切项目重建用，cache 格式）

        substructure 为逐支撑展开的 {str(i): {type, ...}}，直接取整行；
        无对应行（如新增支撑）时回退到纯净默认模板。
        """
        sub = trestle_excel_dict.get("substructure", {})
        if str(idx) in sub:
            row = dict(sub[str(idx)])
            init_type = row.get("type", "单排桩")
            return init_type, row
        # fallback：纯净默认模板（与旧行为一致，端部桥台由用户在类型下拉中选择）
        return "单排桩", self._default_support("单排桩")

    def on_type_change(self, idx, new_type):
        """表格中更改类型时触发：使用纯净默认值，不复用 Excel"""
        old_type = self.support_data[idx].get("type", "")
        self.support_data[idx] = self._default_support(new_type)

        self._auto_register_sections(self.support_data[idx])

        # 刷新表格单元格显示
        for key in self.columns_keys:
            self.table_cells[idx][key].config(text=self.support_data[idx][key])

        # 制动墩增删时自动重排跨径格式
        if old_type == "制动墩" or new_type == "制动墩":
            self._reformat_span_by_brake()

        # 如果改的正是当前选中的行，立刻联动刷新下方输入框
        if self.current_row == idx:
            self.select_row(idx)

        self.refresh_cache()
        draw_support_elevation(self, trestle_excel_dict=trestle_excel_dict, highlight_idx=idx)

    def _reformat_span_by_brake(self):
        """根据当前制动墩位置自动重排跨径字符串。
        有制动墩 → (组1)+(组2)+... 格式；无制动墩 → 连续 @ 格式。"""
        spans = self.spans_data if hasattr(self, 'spans_data') else []
        if not spans:
            return
        # 收集制动墩位置（不包含两端桥台）
        brake_indices = []
        for i, sd in enumerate(self.support_data):
            if sd.get("type") == "制动墩":
                brake_indices.append(i)
        # 格式化一组跨径为紧凑 @ 字符串
        def _fmt_group(vals):
            if not vals:
                return ""
            parts = []
            i = 0
            while i < len(vals):
                v = vals[i]
                count = 1
                while i + count < len(vals) and vals[i + count] == v:
                    count += 1
                if count == 1:
                    parts.append(str(int(v)) if v == int(v) else str(v))
                else:
                    parts.append(f"{count}@{int(v) if v == int(v) else v}")
                i += count
            return "+".join(parts)

        if not brake_indices:
            # 无制动墩：整体合并为连续 @ 格式
            new_span_str = _fmt_group(spans)
        else:
            # 按制动墩将跨径分组
            groups = []
            start = 0
            for bi in brake_indices:
                if bi > start:
                    groups.append(spans[start:bi])
                start = bi
            if start < len(spans):
                groups.append(spans[start:])
            new_span_parts = []
            for g in groups:
                if g:
                    new_span_parts.append(f"({_fmt_group(g)})")
            new_span_str = "+".join(new_span_parts)

        # 更新 Entry 并触发重绘
        self.span.delete(0, 'end')
        self.span.insert(0, new_span_str)
        self.generate_topology()

    def select_row(self, idx):
        """点击行 -> 高亮 -> 刷新绘图"""
        self.current_row = idx
        # 高亮逻辑 (重置所有行为白底，选中行为浅蓝)
        for r_idx, row_dict in enumerate(self.table_cells):
            bg_color = "#e3f2fd" if r_idx == idx else "white"
            for widget in self.table_inner.grid_slaves(row=r_idx + 1):
                try: 
                    widget.config(bg=bg_color) 
                except: pass
                if isinstance(widget, tk.Frame):
                    for sub in widget.winfo_children():
                        if isinstance(sub, tk.Label):
                            sub.config(bg=bg_color)

        sup_type = self.support_data[idx]["type"]
        if sup_type in ["重力式桥台", "简易桥台"]:
            self.btn_param.config(state="normal")
            self.btn_param_text.set(f"桥台无需下部结构设置")
        else:
            self.btn_param.config(state="normal")
            self.btn_param_text.set(f"设置 {idx}# 下部结构参数")
        # self.btn_soil.config(state="normal")
        # self.btn_soil_text.set(f"设置 {idx}# 地形土层")
        # 刷新图示区
        draw_support_elevation(self, trestle_excel_dict=trestle_excel_dict, highlight_idx=idx)

    def refresh_cache(self):
        """重建下部结构缓存 + basic_cache
        {"i": {"type":..., "pile_type":..., ..., "soil_label":..., "soil_data":[...]}}
        """
        self.substructure_cache.clear()
        for i, row in enumerate(self.support_data):
            entry = {"type": row.get("type", "")}
            for k in self.columns_keys:
                entry[k] = row.get(k, "/")
            # 额外字段：不在表格列中但需要保存（供 MCT 生成 + Excel 写入）
            for extra in ("dist_trans_offset", "dist_long_center_offset"):
                entry[extra] = row.get(extra, "/")
            # 截面 detail dict：UI 初始化必需（传递 SectionSetupFrame 完整数据）
            for detail_key in ("pile_sec_detail", "dist_trans_sec_detail", "dist_long_sec_detail",
                               "brace_sec_detail", "brace_sec_tilt_detail"):
                v = row.get(detail_key)
                if isinstance(v, dict) and v:
                    entry[detail_key] = copy.deepcopy(v)
            # 联结系形式为 "-" 时，斜杆/竖杆截面强制设为 "/"
            if entry.get("brace_form") == "-":
                entry["brace_sec_tilt"] = "/"
            # 土层数据：先透传行内已有值（切项目保留），SoilParamTab 收集到新数据时
            # 由 combine_all_summary 覆盖；收集失败时不丢原值
            entry["soil_data"] = row.get("soil_data") or {}
            self.substructure_cache[str(i)] = entry
        self.basic_cache = {
            "span": self.span.get(),
            "bridge_width": self.bridge_width.get() if hasattr(self, 'bridge_width') else "",
            "water_level": self.water_level.get(),
            "deck_level": self.deck_level.get(),
            "x_origin": "",  # x_origin 已移至 SoilParamTab，在 combine_all_summary 中读取
        }

    # soil_label 和 soil_data 已由 SoilParamTab tksheet 实时收集
    # _read_soil_for_pile() 和 get_soil_for_pile() 已废弃并移除

    def open_support_setup(self, idx):
        """点击结构参数设置按钮触发"""
        data = self.support_data[idx]
        if data["type"] in ["重力式桥台", "简易桥台"]:
            return
            
        # 弹出二级参数配置窗口
        dialog = SupportSetupDialog(self.winfo_toplevel(), idx, data, on_select_section=self.parent_ui.record_section)
        self.wait_window(dialog)
        
        if dialog.result:
            # 拿到保存结果，更新主字典
            self.support_data[idx].update(dialog.result)
            self.refresh_cache()

            # 刷新表格前端的文本显示
            for key in self.columns_keys:
                if key in dialog.result:
                    self.table_cells[idx][key].config(text=dialog.result[key])

            # 保存后立即刷新立面示意图
            draw_support_elevation(self, self.current_row if self.current_row is not None else -1, trestle_excel_dict=trestle_excel_dict)

    def on_water_change(self, event=None):
        """设防水位变更：单向同步桥面高程 + 重绘图示"""
        raw = self.water_level.get().replace('+', '')
        try:
            w_val = float(raw)
            self._last_water = w_val
        except (ValueError, TypeError):
            messagebox.showwarning("输入错误", f"设防水位 输入错误: {self.water_level.get()}", parent=self.winfo_toplevel())
            return
        # deck_level = water_level + 0.5
        d_val = w_val + 0.5
        d_str = f"+{d_val:.2f}" if d_val >= 0 else f"{d_val:.2f}"
        self.deck_level.delete(0, 'end')
        self.deck_level.insert(0, d_str)
        draw_support_elevation(self, self.current_row if self.current_row is not None else -1, trestle_excel_dict=trestle_excel_dict)

    def _on_span_width_change(self):
        """计算跨径/计算宽度失焦时即时验证 + 回退 _last_valid + 刷新缓存"""
        from General.DataUtils import midasdisttolst
        errors = []
        span_val = self.span.get().strip()
        if span_val:
            try:
                vals = midasdisttolst(span_val)
                if not vals:
                    errors.append(f"计算跨径 输入错误: {span_val}")
                else:
                    self.span._last_valid = span_val
            except Exception:
                errors.append(f"计算跨径 输入错误: {span_val}")
        else:
            errors.append("计算跨径 输入值不能为空")
        bw_val = self.bridge_width.get().strip()
        if bw_val:
            try:
                float(bw_val)
                self.bridge_width._last_valid = bw_val
            except (ValueError, TypeError):
                errors.append(f"计算宽度 输入错误: {bw_val}")
        else:
            errors.append("计算宽度 输入值不能为空")
        if errors:
            if any("跨径" in e for e in errors):
                prev = getattr(self.span, '_last_valid', "9")
                self.span.delete(0, 'end')
                self.span.insert(0, prev)
            if any("宽度" in e for e in errors):
                prev = getattr(self.bridge_width, '_last_valid', "8")
                self.bridge_width.delete(0, 'end')
                self.bridge_width.insert(0, prev)
            messagebox.showwarning("输入错误", "\n".join(errors), parent=self.winfo_toplevel())
        self.refresh_cache()
        self.generate_topology()

    def import_dwg(self):
        """导入地面线DWG文件"""
        messagebox.showinfo("功能开发中", "地面线DWG导入功能暂未开发，当前无需导入。", parent=self.winfo_toplevel())

    
    def get_transverse_beam_data(self):
        """返回横向分配梁的最短长度(mm)、截面高度(mm)和左端距中心线偏移(mm)"""
        min_len = None; sec_name = None; offset_mm = 3750
        for row in self.support_data:
            try:
                l = float(row.get("dist_trans_len", 0))
                if l > 0 and (min_len is None or l < min_len):
                    min_len = l; sec_name = row.get("dist_trans_sec", "")
                    raw_off = row.get("dist_trans_offset", "/")
                    offset_mm = float(raw_off) if raw_off and raw_off != "/" else 3750
            except: pass
        h_mm = 400
        if sec_name and sec_name != "/":
            sec_path = get_properties_path() or os.path.join(self.parent_ui.Applocation, "Support\\basic_param", "properties_parameter.xlsx")
            db = load_section_library(sec_path)
            if sec_name in db:
                h_mm = int(db[sec_name]["params"].get("H", db[sec_name]["params"].get("D", 400)))
        return {"length_mm": min_len, "height_mm": h_mm, "offset_mm": offset_mm}

    def _auto_register_sections(self, row_data):
        """切换下部结构类型未手动设置时自动注册截面"""
        record_func = getattr(self, '_on_select_section', None)
        if not record_func:
            import gc
            for obj in gc.get_referrers(self.winfo_toplevel()):
                if hasattr(obj, 'record_section'):
                    record_func = obj.record_section
                    break  
        if not record_func:
            return
        HARDCODED_SECTIONS = {
            "630X8":     {"type": "O", "name": "630X8", "params": {"D": "630", "d": "8"}},
            "2I40a":     {"type": "2I", "name": "2I40a", "params": {"H": "400", "B": "142", "tw": "10.5", "tf1": "16.5", "C": "100", "tf2": "16.5"}},
            "2HM588X300":{"type": "2HM","name": "2HM588X300", "params": {"H": "588", "B": "600", "tw": "12", "tf1": "20", "C": "300", "tf2": "20"}},
            "377X6":     {"type": "O", "name": "377X6", "params": {"D": "326", "d": "6"}},
            "219X6":     {"type": "O", "name": "219X6", "params": {"D": "219", "d": "6"}},
        }
        if row_data.get("type") == "制动墩" or row_data.get("type") == "单排桩" :
            for sec_info in HARDCODED_SECTIONS.values():
                record_func(sec_info, allow_update=False)

    def _on_page_shown(self):
        """侧边栏切换到此页时：刷新缓存 + 重绘立面图（含地面线）"""
        self.refresh_cache()
        self.after_idle(lambda: draw_support_elevation(self, self.current_row if self.current_row is not None else -1, trestle_excel_dict=trestle_excel_dict))


# ======================== 上部结构 Tab 页 ========================
class ComponentsTab(ttk.Frame):
    def __init__(self, parent, parent_ui=None):
        super().__init__(parent)
        self.parent_ui = parent_ui
        self._section_db = None  # 截面 DB 缓存（只加载一次）
        self._dist_beam_cache = None  # 分配梁数据缓存
        self.components_cache = {}     # 上部结构参数缓存

        # 统一样式参数
        self.padx = 5
        self.pady = 10
        self.lbl_w = 30
        self.ent_w = 40
        self._current_view = "trans"

        # ======================== 栈桥类型标题 + 下拉框 ========================
        header_row = ttk.Frame(self)
        header_row.pack(fill="x", padx=12, pady=(12, 2))
        tk.Label(header_row, text="栈桥类型", font=("Microsoft YaHei UI", 12, "bold"),
                 fg="#2b6fd4").pack(side="left", padx=(0, 10))
        self.bridge_type_var = tk.StringVar(value=trestle_excel_dict.get("bridge_type", "上承式桁架梁栈桥"))
        self._current_bridge_type = self.bridge_type_var.get()
        _BRIDGE_TYPE_OPTIONS = ["上承式桁架梁栈桥", "型钢栈桥"]
        bridge_cb = ttk.Combobox(header_row, textvariable=self.bridge_type_var,
                                  values=_BRIDGE_TYPE_OPTIONS, state="readonly", width=35)
        bridge_cb.pack(side="left")
        bridge_cb.bind("<<ComboboxSelected>>", lambda _: self._on_bridge_type_change())
        ttk.Separator(self, orient=HORIZONTAL).pack(fill="x", padx=10, pady=(6, 0))

        # ---- 内容控制框架 ----
        self.content_frame = ttk.Frame(self)
        self.content_frame.pack(fill="both", expand=True)

        # # ======================== 顶部：分页按钮 ========================
        self.btn_frame = ttk.Frame(self.content_frame)
        self.btn_frame.pack(fill="x", pady=(15,5), padx=10)
        self.tab_names = ["桥面系定义", "纵梁定义"]
        self.tab_btns = {}
        for name in self.tab_names:
            btn = ttk.Button(self.btn_frame, text=name, command=lambda n=name: self.switch_page(n))
            btn.pack(side="left", padx=2, fill="x", expand=True)
            self.tab_btns[name] = btn

        # ======================== 上方：滚动区域 ========================
        self.scroll_container = ttk.Frame(self.content_frame)
        self.scroll_container.pack(fill="both", expand=True)

        self.scroll_container.grid_rowconfigure(0, weight=1)
        self.scroll_container.grid_columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(self.scroll_container, borderwidth=0, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.scroll_container, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")

        self.inner = ttk.Frame(self.canvas)
        self.inner_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        def on_frame_configure(event):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.inner.bind("<Configure>", on_frame_configure)

        def on_canvas_resize(event):
            self.canvas.itemconfig(self.inner_id, width=event.width)
        self.canvas.bind("<Configure>", on_canvas_resize)

        def _on_mousewheel(event):
            widget = event.widget
            depth = 0
            while widget is not None and depth < 20:
                if isinstance(widget, tk.Canvas):
                    try:
                        widget.yview_scroll(int(-1 * (event.delta / 120)), "units")
                    except Exception:
                        pass
                    return
                try:
                    widget = widget.master
                except AttributeError:
                    return
                depth += 1
        self.winfo_toplevel().bind_all("<MouseWheel>", _on_mousewheel)
        self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", _on_mousewheel))
        self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))

        # ======================== 页面容器分配 ========================


        self.pages = {
            "桥面系定义": ttk.Frame(self.inner),
            "纵梁定义": ttk.Frame(self.inner),
            # "分配梁定义": ttk.Frame(self.inner),
        }

        self.create_deck_system()
        self.create_beam_section()

        self.switch_page("桥面系定义")

        # ===== 底部固定：上部结构示意图 =====
        sep = ttk.Separator(self.content_frame, orient=HORIZONTAL)
        sep.pack(fill="x", padx=10, pady=(5, 0))
        view_btn_frame = ttk.Frame(self.content_frame)
        view_btn_frame.pack(fill="x", padx=10, pady=(5, 0))
        self.btn_trans_view = ttk.Button(view_btn_frame, text="横桥向示意图",
            command=lambda: self._switch_view("trans"))
        self.btn_trans_view.pack(side="left", padx=2)
        self.btn_plan_view = ttk.Button(view_btn_frame, text="平面示意图",
            command=lambda: self._switch_view("plan"))
        self.btn_plan_view.pack(side="left", padx=2)
        self.title_lbl = tk.Label(self.content_frame, text="上部结构横桥向示意图", font=("Microsoft YaHei UI", 9, "bold"),
                                  bg="white", fg="#000000")
        self.title_lbl.pack(anchor="center", pady=(10, 3))
        sep_line = ttk.Separator(self.content_frame, orient=HORIZONTAL)
        sep_line.pack(fill="x", padx=20, pady=(0, 5))
        # 横桥向图
        self.fig_trans = Figure(figsize=(6, 2.5), dpi=100)
        self.ax_trans = self.fig_trans.add_subplot(111)
        self.canvas_trans = FigureCanvasTkAgg(self.fig_trans, master=self.content_frame)
        self.canvas_trans.get_tk_widget().pack(fill="x", padx=10, pady=5)
        # 平面图（同高，初始隐藏）
        self.fig_plan = Figure(figsize=(6, 2.5), dpi=100)
        self.ax_plan = self.fig_plan.add_subplot(111)
        self.canvas_plan = FigureCanvasTkAgg(self.fig_plan, master=self.content_frame)
        self._bind_highlight()
        self._redraw_current_view()
        # 初始化当前桥型缓存（widget 全部就绪）
        self.refresh_cache()

    def reload_from_dict(self, new_dict):
        """数据绑定更新：从 new_dict 刷新上部结构参数，不重建外层 widget"""
        # 更新栈桥类型
        self.bridge_type_var.set(new_dict.get("bridge_type", "上承式桁架梁栈桥"))
        self._current_bridge_type = self.bridge_type_var.get()

        # 切换项目：重置本会话桥型缓存（新项目数据由 create_* + refresh_cache 重建）
        self._dist_beam_cache = None
        self.components_cache.clear()

        # 重建桥面系定义子页面
        for w in self.pages["桥面系定义"].winfo_children():
            w.destroy()
        self.create_deck_system()

        # 重建纵梁定义子页面
        for w in self.pages["纵梁定义"].winfo_children():
            w.destroy()
        self.create_beam_section()

        # 默认切回桥面系定义页面
        self.switch_page("桥面系定义")

        # 重绘示意图
        self._bind_highlight()
        self._redraw_current_view()

        # 恢复其它桥型的会话缓存（当前桥型已由下方 refresh_cache 从 components 重建）
        saved_cache = new_dict.get("_components_cache")
        if isinstance(saved_cache, dict):
            _cur = self.bridge_type_var.get()
            for bt, comp in saved_cache.items():
                if bt != _cur and isinstance(comp, dict):
                    self.components_cache[bt] = copy.deepcopy(comp)

        # 重建 widget 后立即刷新缓存（否则旧 widget 引用在下次 refresh_cache 时 TclError）
        self.refresh_cache()

    def _on_bridge_type_change(self):
        """栈桥类型切换：保留各桥型会话内编辑，重建纵梁定义子页面"""
        new_type = self.bridge_type_var.get()
        if new_type == self._current_bridge_type:
            return
        # 1) 切走前先快照旧桥型当前 widget 值（此时 widget 未销毁，缓存 key 用旧桥型）
        if self._current_bridge_type:
            self.refresh_cache(self._current_bridge_type)
        # 2) 重建纵梁定义子页面（新桥型默认值优先取缓存中的会话编辑，见 _comp_data）
        for w in self.pages["纵梁定义"].winfo_children():
            w.destroy()
        self.create_beam_section()
        self.switch_page("纵梁定义")
        # 3) 刷新新桥型缓存并重绘示意图
        self._current_bridge_type = new_type
        self.refresh_cache()
        self._redraw_current_view()

    # ======================== 切换逻辑 ========================
    def switch_page(self, page_name):
        for name, frame in self.pages.items():
            if name == page_name:
                frame.pack(fill="x", anchor="n") 
            else:
                frame.pack_forget()
        
        self.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    # ======================== 结构定义布置 ========================
    def _comp_data(self):
        """当前桥型参数：优先取本会话缓存（含未保存的桥型切换后编辑），否则取项目数据（按桥型匹配）"""
        bt = self.bridge_type_var.get()
        if bt in self.components_cache and isinstance(self.components_cache[bt], dict):
            return self.components_cache[bt]
        comps = trestle_excel_dict.get("components", {})
        if isinstance(comps, dict):
            for k, v in comps.items():
                if isinstance(v, dict) and k == bt:
                    return v
            for v in comps.values():
                if isinstance(v, dict):
                    return v
        return {}

    # 桥面系定义
    def create_deck_system(self):
        lf_card, lf = make_card_frame(self.pages["桥面系定义"], "桥面系定义")

        # 桥面系类型下拉列表
        # _DECK_TYPES = ["单层钢面板", "双层钢面板", "混凝土桥面板"]
        _DECK_TYPES = ["单层钢面板", "双层钢面板"]
        _init_type = self._comp_data().get("deck_type", "单层钢面板")
        if _init_type not in _DECK_TYPES:
            _init_type = "单层钢面板"
        self.deck_type_var = tk.StringVar(value=_init_type)
        _dt_frame = ttk.Frame(lf)
        _dt_frame.pack(anchor="w", padx=10, pady=(5, 10))
        ttk.Label(_dt_frame, text="桥面系类型:", bootstyle=PRIMARY).pack(side="left")
        self.deck_type_combo = ttk.Combobox(_dt_frame, textvariable=self.deck_type_var,
                                            values=_DECK_TYPES, state="readonly", width=18)
        self.deck_type_combo.pack(side="left", padx=(5, 0))
        self.deck_type_combo.bind("<<ComboboxSelected>>", lambda _: self.render_ribs())

        # 桥面板定义
        self.lf_deck = ttk.LabelFrame(lf, text="桥面板定义", bootstyle=SECONDARY, padding=8)
        self.lf_deck.pack(fill="x", padx=10, pady=self.pady)
        lf_deck = self.lf_deck

        # Row 0: 材质（所有桥面板类型通用）
        self._add_deck_material_row(lf_deck, 0, is_concrete=(_init_type == "混凝土桥面板"))
        self.deck_t = create_label_entry(lf_deck, "桥面板厚(mm):", 1, default_val=self._comp_data().get("deck_thickness", ""), label_width=self.lbl_w, entry_width=self.ent_w, padx=self.padx, pady=self.pady)
        self.deck_t._last_valid = self.deck_t.get() or "10"
        self.deck_t.bind("<FocusOut>", lambda _: self._on_deck_param_change(self.deck_t, "数值", "10", "桥面板厚"))
        self.deck_trans_ecc = create_label_entry(lf_deck, "桥面左端相对线路中心横向间距(mm):", 2, default_val=self._comp_data().get("deck_trans_ecc", ""), label_width=self.lbl_w, entry_width=self.ent_w, padx=self.padx, pady=self.pady)
        self.deck_trans_ecc._last_valid = self.deck_trans_ecc.get() or "2700"
        self.deck_trans_ecc.bind("<FocusOut>", lambda _: self._on_deck_param_change(self.deck_trans_ecc, "数值", "2700", "桥面左端横向间距"))
        ttk.Label(lf_deck, text="桥面板两端较首尾横肋悬臂距离(mm):", width=self.lbl_w, bootstyle=PRIMARY).grid(row=3, column=0, sticky="w", padx=self.padx, pady=self.pady)
        def _on_de_diagram(side, is_focus):
            self._on_diagram_focus(f"deck_end_ext_{side}", is_focus)
        self.deck_end_ext = SplitEntry(lf_deck, row=3,
                                       default_val=self._comp_data().get("deck_end_ext", ""),
                                       placeholders=("左端伸出", "右端伸出"),
                                       entry_width=self.ent_w - 1,
                                       padx=self.padx, pady=3, on_focus=_on_de_diagram)
        self.deck_end_ext.configure(state="normal")
        self.deck_end_ext.bind("<FocusOut>", lambda _: (self.refresh_cache(), self._redraw_current_view()))

        # 动态肋容器
        self.ribs_container = ttk.Frame(lf)
        self.ribs_container.pack(fill="x")
        self.render_ribs()

    def render_ribs(self):
        """根据桥面系类型动态重绘横肋/纵肋/混凝土参数"""
        _BEAM = ["I","2I","HM","2HM","HN","2HN","HW","2HW","C","2C"]
        # 清理动态区域
        for widget in self.ribs_container.winfo_children():
            widget.destroy()
        # 清理桥面板定义 frame 中 row >= 4 的动态控件
        for widget in self.lf_deck.grid_slaves():
            info = widget.grid_info()
            if int(info.get("row", 0)) >= 4:
                widget.destroy()
        # 更新 row 0 材质 Combobox（钢材/混凝土选项随桥面板类型切换）
        self._init_material_options()
        deck_type = self.deck_type_var.get()
        is_concrete = (deck_type == "混凝土桥面板")
        options = self._concrete_grades if is_concrete else self._steel_brands
        spec = self._concrete_spec_id if is_concrete else self._steel_spec_id
        self.deck_material_combo.config(values=options)
        if self.deck_material_combo.get() not in options:
            self.deck_material_combo.set(options[0] if options else "")
        self._deck_material_spec = spec
        if deck_type == "混凝土桥面板":
            # 混凝土桥面板参数直接放入桥面板定义 frame
            _ds = self._comp_data()
            self.panel_length = create_label_entry(self.lf_deck, "单块面板长度(m):", 4,
                                                   default_val=_ds.get("panel_length", ""),
                                                   label_width=self.lbl_w, entry_width=self.ent_w,
                                                   padx=self.padx, pady=self.pady)
            self.panel_length.bind("<FocusOut>", lambda _: self.refresh_cache())
        elif deck_type == "双层钢面板":
            # ── 桥面板分节段（放入桥面板定义 frame）──
            _ds = self._comp_data()
            _init_panel_len = _ds.get("panel_length", "")
            _seg_enabled = bool(_init_panel_len and _init_panel_len != "/")
            self.seg_enabled_var = tk.BooleanVar(value=_seg_enabled)
            self.seg_enabled_cb = ttk.Checkbutton(self.lf_deck, text="桥面板分节段建模", variable=self.seg_enabled_var,
                                                    bootstyle=PRIMARY, command=lambda: (self._on_seg_toggle(), self.refresh_cache()))
            self.seg_enabled_cb.grid(row=4, column=0, sticky="w", padx=self.padx, pady=self.pady)
            self._seg_len_label = ttk.Label(self.lf_deck, text="节段长度(m):", width=self.lbl_w, bootstyle=PRIMARY)
            self._seg_len_label.grid(row=5, column=0, sticky="w", padx=self.padx, pady=self.pady)
            self.panel_length_entry = ttk.Entry(self.lf_deck, width=self.ent_w)
            self.panel_length_entry.grid(row=5, column=1, padx=self.padx, sticky="w")
            self.panel_length_entry.insert(0, _init_panel_len if _seg_enabled else "")
            self.panel_length_entry.bind("<FocusOut>", lambda _: self.refresh_cache())
            self._on_seg_toggle()
            # 纵肋
            lf_rib_long = ttk.LabelFrame(self.ribs_container, text="纵肋定义", bootstyle=SECONDARY, padding=8)
            lf_rib_long.pack(fill="x", padx=10, pady=self.pady)
            self._add_rib_long_material_row(lf_rib_long, 0)
            self.rib_long_sec = create_section_selector(lf_rib_long, "纵肋截面:", 1, initial_data=self._comp_data().get("rib_long_sec_detail") or self._comp_data().get("rib_long_sec", "I25a"), families=_BEAM, label_width=self.lbl_w, entry_width=self.ent_w-3, padx=self.padx, pady=3,
                           on_change=lambda _: self._redraw_current_view(),
                           on_select=lambda r: self.parent_ui and self.parent_ui.record_section(r))
            ttk.Label(lf_rib_long, text="纵肋横向布置(mm):", width=self.lbl_w, bootstyle=PRIMARY).grid(row=2, column=0, sticky="w", padx=self.padx, pady=self.pady)
            def _on_rls_diagram(side, is_focus):
                self._on_diagram_focus(f"rib_long_space_{side}", is_focus)
            self.rib_long_space = SplitEntry(lf_rib_long, row=2,
                                             default_val=self._comp_data().get("rib_long_space", "25@300"),
                                             placeholders=("起始间距", "布置间距"),
                                             entry_width=self.ent_w-1,
                                             padx=self.padx, pady=3, on_focus=_on_rls_diagram)
            self.rib_long_space.bind("<FocusOut>", lambda _: self.refresh_cache())
            # 横肋
            lf_rib_trans = ttk.LabelFrame(self.ribs_container, text="横肋定义", bootstyle=SECONDARY, padding=8)
            lf_rib_trans.pack(fill="x", padx=10, pady=self.pady)
            self._add_rib_trans_material_row(lf_rib_trans, 0)
            self.rib_trans_sec = create_section_selector(lf_rib_trans, "横肋截面:", 1, initial_data=self._comp_data().get("rib_trans_sec_detail") or self._comp_data().get("rib_trans_sec", "I14"), families=_BEAM, label_width=self.lbl_w, entry_width=self.ent_w-3, padx=self.padx, pady=3,
                           on_change=lambda _: self._redraw_current_view(),
                           on_select=lambda r: self.parent_ui and self.parent_ui.record_section(r))
            ttk.Label(lf_rib_trans, text="横肋纵向布置(mm):", width=self.lbl_w, bootstyle=PRIMARY).grid(row=2, column=0, sticky="w", padx=self.padx, pady=self.pady)
            self.rib_trans_space = create_label_entry(lf_rib_trans, "", 2,
                                                       default_val=self._comp_data().get("rib_trans_space", ""),
                                                       label_width=0, entry_width=self.ent_w, padx=self.padx, pady=self.pady)
            self.rib_trans_space._last_valid = self.rib_trans_space.get() or "750,68@750"
            self.rib_trans_space.bind("<FocusIn>", lambda e: self._on_diagram_focus("rib_trans_space", True))
            self.rib_trans_space.bind("<FocusOut>", lambda e: (self._on_diagram_focus("rib_trans_space", False), self._on_deck_param_change(self.rib_trans_space, "间距格式", "750,68@750", "横肋纵向布置")))
        else:
            # 单层钢面板
            lf_rib_trans = ttk.LabelFrame(self.ribs_container, text="横肋定义", bootstyle=SECONDARY, padding=8)
            lf_rib_trans.pack(fill="x", padx=10, pady=self.pady)
            self._add_rib_trans_material_row(lf_rib_trans, 0)
            self.rib_trans_sec = create_section_selector(lf_rib_trans, "横肋截面:", 1, initial_data=self._comp_data().get("rib_trans_sec_detail") or self._comp_data().get("rib_trans_sec", "I14"), families=_BEAM, label_width=self.lbl_w, entry_width=self.ent_w-3, padx=self.padx, pady=3,
                           on_change=lambda _: self._redraw_current_view(),
                           on_select=lambda r: self.parent_ui and self.parent_ui.record_section(r))
            ttk.Label(lf_rib_trans, text="横肋纵向布置(mm):", width=self.lbl_w, bootstyle=PRIMARY).grid(row=2, column=0, sticky="w", padx=self.padx, pady=self.pady)
            self.rib_trans_space = create_label_entry(lf_rib_trans, "", 2,
                                                       default_val=self._comp_data().get("rib_trans_space", ""),
                                                       label_width=0, entry_width=self.ent_w, padx=self.padx, pady=self.pady)
            self.rib_trans_space._last_valid = self.rib_trans_space.get() or "750,68@750"
            self.rib_trans_space.bind("<FocusIn>", lambda e: self._on_diagram_focus("rib_trans_space", True))
            self.rib_trans_space.bind("<FocusOut>", lambda e: (self._on_diagram_focus("rib_trans_space", False), self._on_deck_param_change(self.rib_trans_space, "间距格式", "750,68@750", "横肋纵向布置")))
        self._rebind_highlight()

    # 纵梁定义
    def create_beam_section(self):
        lf = ttk.LabelFrame(self.pages["纵梁定义"], text="纵梁定义", bootstyle=SECONDARY, padding=10)
        lf.pack(fill="x", padx=12, pady=6)
        _is_steel = ("型钢" in self.bridge_type_var.get())
        _comp = self._comp_data()
        _row = 0

        if _is_steel:
            # === 型钢栈桥：纵梁材质 + 纵梁截面 + 纵梁横向布置 + 伸出距离 + 第一排间距 ---
            # row 0: 纵梁材质
            self._init_material_options()
            ttk.Label(lf, text="纵梁材质:", width=self.lbl_w, bootstyle=PRIMARY).grid(row=_row, column=0, sticky="w", padx=self.padx, pady=self.pady)
            self.beam_material_combo = ttk.Combobox(lf, values=self._steel_brands, state="readonly", width=10)
            self.beam_material_combo.grid(row=_row, column=1, sticky="w", padx=self.padx, pady=2)
            _, _bm_brand = _parse_material_value(_comp.get("beam_material", ""))
            if _bm_brand and _bm_brand in self._steel_brands:
                self.beam_material_combo.set(_bm_brand)
            elif self._steel_brands:
                self.beam_material_combo.set(self._steel_brands[0])
            self.beam_material_combo.bind("<<ComboboxSelected>>", lambda _: self.refresh_cache())
            self._beam_material_spec = self._steel_spec_id
            _row += 1

            # row 1: 纵梁截面
            _STEEL_FAMILIES = ["I", "2I", "HM", "2HM", "HN", "2HN", "HW", "2HW", "C", "2C"]
            self.beam_sec = create_section_selector(lf, "纵梁截面:", _row,
                initial_data=_comp.get("beam_sec_detail") or _comp.get("beam_sec", ""),
                families=_STEEL_FAMILIES, label_width=self.lbl_w, entry_width=self.ent_w-3,
                padx=self.padx, pady=3,
                on_change=lambda _: (self.refresh_cache(), self._redraw_current_view()),
                on_select=lambda r: self.parent_ui and self.parent_ui.record_section(r))
            _row += 1
        else:
            # === 桁架梁栈桥：纵梁类型 + 纵梁横向布置 + 伸出距离 + 第一排间距 ---
            # row 0: 纵梁类型
            ttk.Label(lf, text="纵梁类型:", bootstyle=PRIMARY, width=self.lbl_w).grid(row=_row, column=0, padx=self.padx, pady=self.pady)
            self.truss_beam_type = ttk.Combobox(lf, values=list(TRUSS_BEAM_TYPES.keys()), state="readonly", width=38)
            self.truss_beam_type.grid(row=_row, column=1, padx=self.padx, pady=self.pady)
            _truss_val = _comp.get("truss_beam_type", "")
            if not _truss_val or _truss_val == "/":
                _truss_val = list(TRUSS_BEAM_TYPES.keys())[0] if TRUSS_BEAM_TYPES else "321型贝雷梁"
            self.truss_beam_type.set(_truss_val)
            _row += 1

        # === 以下为两种栈桥类型共有的行 ===
        # 纵梁横向布置
        self.beam_space = create_label_entry(lf, "纵梁横向布置(mm):", _row, default_val=_comp.get("beam_space", ""), label_width=self.lbl_w, entry_width=self.ent_w, padx=self.padx, pady=self.pady)
        self.beam_space._last_valid = self.beam_space.get() or "3@900+1500+3@900"
        self.beam_space.bind("<FocusOut>", lambda _: self._on_deck_param_change(self.beam_space, "间距格式", "3@900+1500+3@900", "纵梁横向布置"))
        _row += 1

        # 纵梁两端伸出距离
        def _on_bc_diagram(side, is_focus):
            self._on_diagram_focus(f"beam_cant_{side}", is_focus)
        ttk.Label(lf, text="纵梁两端较边墩悬臂距离(mm):", width=self.lbl_w, bootstyle=PRIMARY).grid(row=_row, column=0, sticky="w", padx=self.padx, pady=self.pady)
        self.beam_cant = SplitEntry(lf, row=_row,
                                       default_val=_comp.get("beam_cant", ""),
                                       placeholders=("左端伸出", "右端伸出"),
                                       entry_width=self.ent_w-1,
                                       padx=self.padx, pady=3, on_focus=_on_bc_diagram)
        self.beam_cant.configure(state="normal")
        self.beam_cant.bind("<FocusOut>", lambda _: self._on_beam_cant_change())
        _row += 1

        # 第一排相对线路中心线间距
        self.beam_first_space = create_label_entry(lf, "第一排相对线路中心线间距(mm):", _row, default_val=_comp.get("beam_first_space", ""), label_width=self.lbl_w, entry_width=self.ent_w, padx=self.padx, pady=self.pady)
        self.beam_first_space._last_valid = self.beam_first_space.get() or "0"
        self.beam_first_space.bind("<FocusOut>", lambda _: self._on_deck_param_change(self.beam_first_space, "数值", "0", "第一排间距"))

    def _init_material_options(self):
        """初始化材质选项列表（从 steel_dict/concrete_dict）"""
        if hasattr(self, '_material_options_ready'):
            return
        self._material_options_ready = True
        steel_dict = trestle_excel_dict.get("steel_dict", {})
        concrete_dict = trestle_excel_dict.get("concrete_dict", {})
        self._steel_brands = []
        self._steel_spec_id = ""
        for spec_data in steel_dict.values():
            if not self._steel_spec_id:
                self._steel_spec_id = spec_data.get("钢结构规范编号", "")
            self._steel_brands.extend(spec_data.get("牌号参数", {}).keys())
        if not self._steel_brands:
            self._steel_brands = ["Q235", "Q345"]
        self._concrete_grades = []
        self._concrete_spec_id = ""
        for spec_data in concrete_dict.values():
            if not self._concrete_spec_id:
                self._concrete_spec_id = spec_data.get("混凝土规范编号", "")
            self._concrete_grades.extend(spec_data.get("强度等级参数", {}).keys())
        if not self._concrete_grades:
            self._concrete_grades = ["C30", "C35", "C40"]

    def _add_deck_material_row(self, parent_frame, row, is_concrete=False):
        """桥面板材质 Combobox"""
        self._init_material_options()
        _ds = self._comp_data()
        options = self._concrete_grades if is_concrete else self._steel_brands
        spec = self._concrete_spec_id if is_concrete else self._steel_spec_id
        ttk.Label(parent_frame, text="桥面板材质:", width=self.lbl_w, bootstyle=PRIMARY).grid(
            row=row, column=0, sticky="w", padx=self.padx, pady=self.pady)
        self.deck_material_combo = ttk.Combobox(parent_frame, values=options, state="readonly", width=10)
        self.deck_material_combo.grid(row=row, column=1, padx=self.padx, sticky="w")
        _, brand = _parse_material_value(_ds.get("deck_material", ""))
        if brand and brand in options:
            self.deck_material_combo.set(brand)
        elif options:
            self.deck_material_combo.set(options[0])
        self.deck_material_combo.bind("<<ComboboxSelected>>", lambda _: self.refresh_cache())
        self._deck_material_spec = spec

    def _add_rib_trans_material_row(self, parent_frame, row):
        """横肋材质 Combobox（横肋定义 frame 内第一行）"""
        self._init_material_options()
        _ds = self._comp_data()
        ttk.Label(parent_frame, text="横肋材质:", width=self.lbl_w, bootstyle=PRIMARY).grid(
            row=row, column=0, sticky="w", padx=self.padx, pady=self.pady)
        self.rib_trans_material_combo = ttk.Combobox(parent_frame, values=self._steel_brands, state="readonly", width=10)
        self.rib_trans_material_combo.grid(row=row, column=1, padx=self.padx, sticky="w")
        _, brand = _parse_material_value(_ds.get("rib_trans_material", ""))
        if brand and brand in self._steel_brands:
            self.rib_trans_material_combo.set(brand)
        elif self._steel_brands:
            self.rib_trans_material_combo.set(self._steel_brands[0])
        self.rib_trans_material_combo.bind("<<ComboboxSelected>>", lambda _: self.refresh_cache())
        self._rib_trans_material_spec = self._steel_spec_id

    def _add_rib_long_material_row(self, parent_frame, row):
        """纵肋材质 Combobox（纵肋定义 frame 内第一行）"""
        self._init_material_options()
        _ds = self._comp_data()
        ttk.Label(parent_frame, text="纵肋材质:", width=self.lbl_w, bootstyle=PRIMARY).grid(
            row=row, column=0, sticky="w", padx=self.padx, pady=self.pady)
        self.rib_long_material_combo = ttk.Combobox(parent_frame, values=self._steel_brands, state="readonly", width=10)
        self.rib_long_material_combo.grid(row=row, column=1, padx=self.padx, sticky="w")
        _, brand = _parse_material_value(_ds.get("rib_long_material", ""))
        if brand and brand in self._steel_brands:
            self.rib_long_material_combo.set(brand)
        elif self._steel_brands:
            self.rib_long_material_combo.set(self._steel_brands[0])
        self.rib_long_material_combo.bind("<<ComboboxSelected>>", lambda _: self.refresh_cache())
        self._rib_long_material_spec = self._steel_spec_id

    def refresh_cache(self, bridge_type=None):
        """重建指定桥型缓存（默认当前桥型）；缓存条目在构建完成后才写入，中途异常不会清空既有数据

        bridge_type 为 None 时取当前下拉值；切换桥型时需显式传入旧桥型以快照其 widget 值。
        """
        try:
            return self._refresh_cache_inner(bridge_type)
        except tk.TclError:
            pass  # widget 已销毁（切换项目时旧 tab 被销毁），跳过本次缓存刷新

    @staticmethod
    def _sec_detail(widget):
        """从 create_section_selector 按钮取完整截面 dict（UI 初始化用）"""
        if widget is None:
            return {}
        try:
            data = widget.get_section_data()
            if isinstance(data, dict):
                return copy.deepcopy(data)
        except Exception:
            pass
        return {}

    def _refresh_cache_inner(self, bridge_type=None):
        if bridge_type is None:
            bridge_type = self.bridge_type_var.get()
        deck_type = self.deck_type_var.get()
        cache = {
            "deck_type": deck_type,
            "deck_thickness": self.deck_t.get(),
        }
        # 桥面板特有字段
        if deck_type == "混凝土桥面板":
            cache["panel_length"] = self.panel_length.get() if hasattr(self, 'panel_length') else ""
        elif deck_type == "双层钢面板":
            _seg_on = self.seg_enabled_var.get() if hasattr(self, 'seg_enabled_var') else False
            cache["seg_enabled"] = _seg_on
            cache["panel_length"] = self.panel_length_entry.get() if _seg_on and hasattr(self, 'panel_length_entry') else ""

        # 横肋/纵肋：始终保存（切换 deck_type 时保留数据）
        def _safe_get(widget, fallback=""):
            try: return widget.get()
            except: return fallback
        old_cache = self.components_cache.get(bridge_type, {})
        cache["rib_trans_sec"] = self._sec_summary(getattr(self, 'rib_trans_sec', None))
        cache["rib_trans_sec_detail"] = self._sec_detail(getattr(self, 'rib_trans_sec', None))
        cache["rib_trans_space"] = _safe_get(getattr(self, 'rib_trans_space', None), old_cache.get("rib_trans_space", ""))
        cache["rib_long_sec"] = self._sec_summary(getattr(self, 'rib_long_sec', None))
        cache["rib_long_sec_detail"] = self._sec_detail(getattr(self, 'rib_long_sec', None))
        cache["rib_long_space"] = _safe_get(getattr(self, 'rib_long_space', None), old_cache.get("rib_long_space", ""))

        # 材质字段（带规范编号前缀）
        def _mat_val(combo, spec_attr):
            if not hasattr(self, combo):
                return ""
            try:
                brand = getattr(self, combo).get()
            except tk.TclError:
                return ""
            spec = getattr(self, spec_attr, "")
            return f"{spec},{brand}" if spec and brand else brand

        cache["deck_material"] = _mat_val("deck_material_combo", "_deck_material_spec")
        cache["rib_trans_material"] = _mat_val("rib_trans_material_combo", "_rib_trans_material_spec")
        cache["rib_long_material"] = _mat_val("rib_long_material_combo", "_rib_long_material_spec")
        # 混凝土标号从桥面板材质中提取（格式 "规范编号,标号"）
        if deck_type == "混凝土桥面板":
            _dm = cache.get("deck_material", "")
            cache["concrete_grade"] = _dm.split(",", 1)[1] if "," in _dm else _dm

        # 纵梁参数（由栈桥类型驱动；beam_type 仅桁架类型读取）
        _is_steel = ("型钢" in bridge_type)
        if _is_steel:
            cache["truss_beam_type"] = ""
            cache["beam_material"] = _mat_val("beam_material_combo", "_beam_material_spec")
            cache["beam_sec"] = self._sec_summary(getattr(self, 'beam_sec', None))
            cache["beam_sec_detail"] = self._sec_detail(getattr(self, 'beam_sec', None))
        else:
            cache["truss_beam_type"] = self.truss_beam_type.get() if hasattr(self, 'truss_beam_type') else ""
            cache["beam_material"] = ""
            cache["beam_sec"] = ""
            cache["beam_sec_detail"] = {}

        cache["beam_space"] = self.beam_space.get() if hasattr(self, 'beam_space') else ""
        cache["beam_cant"] = self.beam_cant.get() if hasattr(self, 'beam_cant') else "0,0"
        cache["beam_first_space"] = self.beam_first_space.get() if hasattr(self, 'beam_first_space') else ""
        cache["deck_trans_ecc"] = self.deck_trans_ecc.get() if hasattr(self, 'deck_trans_ecc') else ""
        cache["deck_end_ext"] = self.deck_end_ext.get() if hasattr(self, 'deck_end_ext') else "0,0"
        # 原子写入：构建成功后才更新缓存，中途异常不会清空既有数据
        self.components_cache[bridge_type] = cache

    @staticmethod
    def _sec_summary(widget):
        """截面对应串"""
        if widget is None:
            return ""
        try:
            data = widget.get_section_data()
            if data:
                return data.get("section_name", "")
        except Exception:
            pass
        return ""

    def _get_beam_dims(self):
        """纵梁截面尺寸(m)：(beam_h, beam_w)

        型钢栈桥按实际截面 B/H（优先取 params，其次截面库兜底）；贝雷梁用固定默认值。
        供横桥向示意图、平面示意图与 get_veh_diagram_geometry 共用。
        """
        _is_steel = ("型钢" in self.bridge_type_var.get())
        if _is_steel and hasattr(self, 'beam_sec'):
            _sd = self.beam_sec.get_section_data()
            if _sd and "params" in _sd:
                return (float(_sd["params"].get("H", 300)) / 1000.0,
                        float(_sd["params"].get("B", 150)) / 1000.0)
            if _sd and _sd.get("section_name"):
                db = self._get_section_db()
                sn = _sd["section_name"]
                if sn in db:
                    return (_sec_dim(db[sn], "H", 300) / 1000.0,
                            _sec_dim(db[sn], "B", 150) / 1000.0)
            return 0.3, 0.15
        return 1.5, 0.08

    def _on_deck_param_change(self, entry, check_type, fallback, label=""):
        """Entry 失焦验证：数值/间距格式，错误弹窗回退 _last_valid"""
        val = entry.get().strip()
        if not val:
            entry.delete(0, 'end')
            entry.insert(0, getattr(entry, '_last_valid', fallback))
            messagebox.showwarning("输入错误", f"{label} 输入值不能为空", parent=self.winfo_toplevel())
            self.refresh_cache()
            self._redraw_current_view()
            return
        ok = False
        if check_type == "数值":
            try:
                float(val)
                ok = True
            except (ValueError, TypeError):
                pass
        elif check_type == "间距格式":
            try:
                from General.DataUtils import midasdisttolst
                result = midasdisttolst(val)
                ok = bool(result)
            except Exception:
                pass
        if ok:
            entry._last_valid = val
            self.refresh_cache()
            self._redraw_current_view()
        else:
            prev = getattr(entry, '_last_valid', fallback)
            entry.delete(0, 'end')
            entry.insert(0, prev)
            messagebox.showwarning("输入错误", f"{label} 输入错误: {val}", parent=self.winfo_toplevel())
            self.refresh_cache()
            self._redraw_current_view()

    def _on_beam_cant_change(self):
        """beam_cant (SplitEntry) 失焦验证：逗号分隔两项均数值"""
        e1_val = self.beam_cant.e1.get().strip() if hasattr(self.beam_cant, 'e1') else ""
        e2_val = self.beam_cant.e2.get().strip() if hasattr(self.beam_cant, 'e2') else ""
        errors = []
        for i, v in enumerate([e1_val, e2_val]):
            if v:
                try:
                    float(v)
                except (ValueError, TypeError):
                    errors.append(f"纵梁两端伸出距离 第{i+1}项输入错误: {v}")
        if errors:
            prev = getattr(self.beam_cant, '_last_valid', "0,0")
            parts = prev.split(",")
            if hasattr(self.beam_cant, 'e1'):
                self.beam_cant.e1.delete(0, 'end')
                self.beam_cant.e1.insert(0, parts[0].strip() if len(parts) > 0 else "0")
            if hasattr(self.beam_cant, 'e2'):
                self.beam_cant.e2.delete(0, 'end')
                self.beam_cant.e2.insert(0, parts[1].strip() if len(parts) > 1 else "0")
            messagebox.showwarning("输入错误", "\n".join(errors), parent=self.winfo_toplevel())
        else:
            self.beam_cant._last_valid = f"{e1_val},{e2_val}"
        self.refresh_cache()

    def _redraw_current_view(self):
        """根据当前视图重绘对应图示"""
        if self._current_view == "plan":
            draw_components_plan_diagram(self)
        else:
            draw_components_transverse_diagram(self)

    def _switch_view(self, view, redraw=True):
        """切换横桥向 / 平面视图"""
        if self._current_view == view:
            return
        if view == "trans":
            self.canvas_plan.get_tk_widget().pack_forget()
            self.canvas_trans.get_tk_widget().pack(fill="x", padx=10, pady=5)
            self.title_lbl.config(text="上部结构横桥向示意图")
            if redraw: draw_components_transverse_diagram(self)
        else:
            self.canvas_trans.get_tk_widget().pack_forget()
            self.canvas_plan.get_tk_widget().pack(fill="x", padx=10, pady=5)
            self.title_lbl.config(text="上部结构平面示意图(仅示意纵梁、横肋、桥面板)")
            if redraw: draw_components_plan_diagram(self)
        self._current_view = view


    def get_veh_diagram_geometry(self):
        """
        供 MoveLoadTab 及其他页面复用：主动读取当前的 UI 输入，计算并返回最新的上部结构几何数据。
        彻底解决“未切换过上部结构页面导致拿不到缓存”而使用错误默认值的 Bug。
        """
        # --- 1. 桥面板及横肋长度参数（从 SupportsTab 读取，保持联动） ---
        try:
            bw = self.parent_ui.tab_supports.bridge_width.get() if self.parent_ui and hasattr(self.parent_ui, 'tab_supports') else "6"
            bridge_width = float(bw)
        except:
            bridge_width = 6.0
        try: deck_t_m = max(float(self.deck_t.get()) / 1000.0, 0.04)
        except: deck_t_m = 0.04

        is_double = (self.deck_type_var.get() == "双层钢面板") if hasattr(self, 'deck_type_var') else False
        try: deck_trans_ecc = float(self.deck_trans_ecc.get()) / 1000.0
        except: deck_trans_ecc = 2.7
        rib_trans_left = 0 - deck_trans_ecc
        rib_trans_right = rib_trans_left + bridge_width

        # --- 2. 横肋截面高度 ---
        h_rib_trans = 0.25
        if hasattr(self, 'rib_trans_sec'):
            rt = self.rib_trans_sec.get_section_data()
            if rt and "params" in rt and "H" in rt["params"]:
                h_rib_trans = float(rt["params"]["H"]) / 1000.0
            elif rt and "section_name" in rt:
                db = self._get_section_db()
                if rt["section_name"] in db:
                    h_rib_trans = _sec_dim(db[rt["section_name"]], "H", 250) / 1000.0

        # --- 3. 纵肋参数 (仅双层桥面系) ---
        h_rib_long = 0.14
        b_rib_long = 0.08
        rib_positions = []
        if is_double and hasattr(self, 'rib_long_sec'):
            rl = self.rib_long_sec.get_section_data()
            if rl and "params" in rl and "H" in rl["params"]:
                h_rib_long = float(rl["params"]["H"]) / 1000.0
                b_rib_long = float(rl["params"]["B"]) / 1000.0
            elif rl and "section_name" in rl:
                db = self._get_section_db()
                if rl["section_name"] in db:
                    h_rib_long = _sec_dim(db[rl["section_name"]], "H", 140) / 1000.0
                    b_rib_long = _sec_dim(db[rl["section_name"]], "B", 80) / 1000.0
            # rib_long_space 格式 "a,b@x"，a=前缀(原横肋悬臂d2)
            try:
                _rls_raw = self.rib_long_space.get()
                _rls_parts = _rls_raw.split(",", 1)
                rib_trans_d2_m = float(_rls_parts[0]) / 1000.0 if len(_rls_parts) > 1 else 0.0
                rib_long_space = midasdisttolst(_rls_parts[-1])
            except:
                rib_trans_d2_m = 0.0; rib_long_space = []
            leftmost_rib_x = rib_trans_left + rib_trans_d2_m
            rib_positions = [leftmost_rib_x]
            cur_x = leftmost_rib_x
            for sp in rib_long_space:
                cur_x += sp / 1000.0
                if cur_x > rib_trans_right: break
                rib_positions.append(cur_x)

        # --- 4. 纵梁参数（以线路中心为基准，beam_first_space 为第一排间距）---
        try: beam_space = midasdisttolst(self.beam_space.get())
        except: beam_space = midasdisttolst("5@900")
        try: beam_first_space_m = float(self.beam_first_space.get()) / 1000.0
        except: beam_first_space_m = 0.0
        # 纵梁截面尺寸：型钢按实际 B/H，贝雷梁用固定默认值
        beam_h, beam_w = self._get_beam_dims()
        leftmost_beam_x = - beam_first_space_m
        beam_pos = [leftmost_beam_x]
        cur_x = leftmost_beam_x
        for sp in beam_space:
            cur_x += sp / 1000.0
            if cur_x > rib_trans_right: break
            beam_pos.append(cur_x)

        # --- 5. 分配梁参数 ---
        if self._dist_beam_cache is None and self.parent_ui and hasattr(self.parent_ui, 'tab_supports'):
            self._dist_beam_cache = self.parent_ui.tab_supports.get_transverse_beam_data()
        if self._dist_beam_cache and self._dist_beam_cache.get("length_mm") is not None:
            dist_len_m = self._dist_beam_cache["length_mm"] / 1000.0
            dist_h_m = self._dist_beam_cache["height_mm"] / 1000.0
            dist_offset_m = self._dist_beam_cache.get("offset_mm", 3750) / 1000.0
        else:
            dist_len_m = 8.0; dist_h_m = 0.4; dist_offset_m = 3.75
        dist_left = -dist_offset_m
        dist_right = dist_left + dist_len_m

        # --- 6. 核心标高计算 ---
        deck_top = 0.0
        deck_bottom = deck_top - deck_t_m
        if is_double:
            rib_long_y = deck_bottom - h_rib_long
            rib_trans_y = rib_long_y - h_rib_trans
        else:
            rib_long_y = None
            rib_trans_y = deck_bottom - h_rib_trans
            
        beam_y = rib_trans_y - beam_h
        dist_y = beam_y - dist_h_m

        # 返回完全契合画图脚本的字典
        return {
            "deck_bottom": deck_bottom, "deck_top": deck_top, "deck_width": bridge_width,
            "is_double": is_double,
            "rib_trans_y": rib_trans_y, "h_rib_trans": h_rib_trans, "rib_trans_w": bridge_width,
            "rib_trans_left": rib_trans_left, "rib_trans_right": rib_trans_right,
            "rib_long_y": rib_long_y, "h_rib_long": h_rib_long, "b_rib_long": b_rib_long,
            "rib_positions": rib_positions,
            "beam_y": beam_y, "beam_h": beam_h, "beam_w": beam_w,
            "beam_pos": beam_pos,
            "dist_y": dist_y, "dist_h": dist_h_m, "dist_left": dist_left, "dist_right": dist_right,
        }

    def _on_page_shown(self):
        """侧边栏切换到此页时：清分配梁缓存 + 刷新参数缓存 + 重绘"""
        self._dist_beam_cache = None
        self.refresh_cache()
        if self._current_view == "plan":
            draw_components_plan_diagram(self)
        else:
            draw_components_transverse_diagram(self)

    def _bind_highlight(self):
        """为可聚焦控件绑定焦点事件（跳过已销毁控件）"""
        for w, key in self._hl_map().items():
            if w is None: continue
            if getattr(w, '_hl_bound', False): continue
            try:
                # 如果是按钮（如截面选择器 btn），绑定鼠标点击事件
                if isinstance(w, (tk.Button, ttk.Button)):
                    w.bind("<Button-1>", lambda e, k=key: self._on_diagram_focus(k, True), add="+")
                # 如果是普通的输入框(Entry)，绑定焦点获取与失去事件
                else:
                    w.bind("<FocusIn>", lambda e, k=key: self._on_diagram_focus(k, True), add="+")
                    w.bind("<FocusOut>", lambda e, k=key: self._on_diagram_focus(k, False), add="+")
            except tk.TclError:
                pass

    def _get_section_db(self):
        """截面 DB 缓存（避免每次绘图都读 Excel）"""
        if self._section_db is None:
            sec_path = get_properties_path()
            if sec_path:
                self._section_db = load_section_library(sec_path)
            else:
                self._section_db = {}
        return self._section_db

    def _on_seg_toggle(self, *_):
        """桥面板分节段建模 checkbox 联动：显示/隐藏节段长度行"""
        enabled = self.seg_enabled_var.get()
        if hasattr(self, '_seg_len_label'):
            self._seg_len_label.grid() if enabled else self._seg_len_label.grid_remove()
        if hasattr(self, 'panel_length_entry'):
            if enabled:
                self.panel_length_entry.grid()
            else:
                self.panel_length_entry.delete(0, 'end')
                self.panel_length_entry.grid_remove()

    def _rebind_highlight(self):
        """render_ribs 后重新绑定"""
        self._bind_highlight()
        self._redraw_current_view()

    def _hl_map(self):
        """widget → highlight_key 映射（entry1/entry2 分离高亮）"""
        m = {}
        if hasattr(self, 'deck_t'):           m[self.deck_t]           = "deck"
        if hasattr(self, 'bridge_width'):     m[self.bridge_width]     = "bridge_width"
        if hasattr(self, 'deck_trans_ecc'):   m[self.deck_trans_ecc]   = "deck_trans_ecc"
        # SplitEntry 双字段：entry1 → left, entry2 → right
        if hasattr(self, 'deck_end_ext'):
            m[self.deck_end_ext.entry1] = "deck_end_ext_left"
            m[self.deck_end_ext.entry2] = "deck_end_ext_right"
        if hasattr(self, 'rib_trans_sec'):    m[self.rib_trans_sec]    = "rib_trans"
        if hasattr(self, 'rib_trans_space'):
            m[self.rib_trans_space] = "rib_trans_space"
        if hasattr(self, 'rib_long_sec'):     m[self.rib_long_sec]     = "rib_long"
        if hasattr(self, 'rib_long_space'):
            m[self.rib_long_space.entry1] = "rib_long_space_left"
            m[self.rib_long_space.entry2] = "rib_long_space_right"
        if hasattr(self, 'beam_space'):     m[self.beam_space]     = "beam_space"
        if hasattr(self, 'beam_cant'):
            m[self.beam_cant.entry1] = "beam_cant_left"
            m[self.beam_cant.entry2] = "beam_cant_right"
        if hasattr(self, 'beam_first_space'): m[self.beam_first_space] = "beam_first_space"
        return m

    def _on_diagram_focus(self, key, is_focus):
        self._hl_key = key if is_focus else None
        # 纵桥向相关参数 → 切换至平面图（支持 split keys）
        _plan_keys = ("rib_trans_space", "beam_cant_left", "beam_cant_right",
                       "deck_end_ext_left", "deck_end_ext_right")
        _plan_prefix = ("beam_cant", "deck_end_ext")
        if key in _plan_keys or any(key.startswith(p) for p in _plan_prefix):
            if is_focus and self._current_view != "plan":
                self._switch_view("plan", redraw=False)
            elif not is_focus and self._current_view == "plan":
                self._switch_view("trans", redraw=False)
        # 横桥向相关参数 → 切换至横截面图（支持 split keys）
        elif (key in ("bridge_width", "deck_trans_ecc", "rib_trans", "rib_long",
                      "deck", "beam_first_space", "beam_space") or
              key.startswith("rib_long_space")):
            if is_focus and self._current_view != "trans":
                self._switch_view("trans", redraw=False)
        self._redraw_current_view()


class SoilParamTab(ttk.Frame):
    """地形土层 Tab：调用地质插件 → 读取 Excel 数据源"""

    def __init__(self, parent, parent_ui=None):
        super().__init__(parent)
        self.parent_ui = parent_ui
        self._active_sheet = ""

        # ================= 1. 顶部地形参数区 =================
        card_geo, geology_param = make_card_frame(self, "地形参数")
        row1 = ttk.Frame(geology_param)
        row1.pack(fill="x", pady=(6, 2))
        self._plugin_btn = ttk.Button(row1, text="打开地形设置插件", bootstyle=PRIMARY, command=self._open_plugin, width=37)
        self._plugin_btn.pack(side="left")
        # self._status_lbl = ttk.Label(row1, text="", font=("微软雅黑", 9), foreground="#888")
        # self._status_lbl.pack(side="right")
        
        row2 = ttk.Frame(geology_param)
        row2.pack(fill="x", pady=(2, 6))
        ttk.Label(row2, text="栈桥起始定位线(0#)x坐标(m):", bootstyle=PRIMARY).pack(side="left", padx=(0, 6), pady=5)
        self._x_origin_var = tk.StringVar()
        x_init = trestle_excel_dict.get("basic", {}).get("x_origin", "")
        self._x_origin_var.set(str(x_init) if x_init else "")
        self._x_origin_entry = ttk.Entry(row2, textvariable=self._x_origin_var, width=14)
        self._x_origin_entry.pack(side="left", padx=(5, 0), pady=5)

        # ================= 2. 核心状态追踪器 =================
        self._soil_cache = {}          # 缓存各桩号数据
        self._is_edited = set()        # 记录被手动修改过的桩号
        self._soil_top_cache = {}      # 各桩号土顶标高(编辑值): {"0": "2.00", ...}
        self._current_pile_no = None   # 记录当前真实显示的桩号 (防串线核心)

        # ================= 3. 下方土层参数区 =================
        card_soil, self._soil_content = make_card_frame(self, "土层参数")
        soil_top_row = ttk.Frame(self._soil_content)
        
        ttk.Label(soil_top_row, text="选择桩号:", bootstyle=PRIMARY).pack(side="left", padx=(0, 5),pady=5)
        self._pile_combo = ttk.Combobox(soil_top_row, values=[], state="readonly", width=15)
        self._pile_combo.pack(side="left", padx=(0, 5),pady=5)
        self._pile_combo.bind("<<ComboboxSelected>>", lambda e: self._on_pile_change())
        
        soil_top_row.pack(fill="x", pady=(0, 5))
        self._soil_reset_btn = ttk.Button(soil_top_row, text="重置为地形图读取值",
                                          bootstyle=SUCCESS, command=self._on_reset_soil,width=18)
        self._soil_reset_btn.pack(side="right", padx=(0, 5),pady=5)
        # sep_line = tb.Separator(master=soil_top_row, orient=HORIZONTAL)
        # sep_line.pack(fill=X,side="bottom")
        
        pile_row = ttk.Frame(self._soil_content)
        pile_row.pack(fill="x", pady=(0, 5))
        ttk.Label(pile_row, text="土顶标高(m):", bootstyle=PRIMARY).pack(side="left", padx=(0, 5))
        self._soil_top_entry = ttk.Entry(pile_row, width=14, justify="left")
        self._soil_top_entry.pack(side="left", padx=(5, 5))
        self._soil_top_var = tk.StringVar(value="")
        self._soil_top_entry.config(textvariable=self._soil_top_var)
        self._soil_top_entry.bind("<FocusOut>", lambda e: self._on_soil_top_edit())

        self._pile_x_lbl = ttk.Label(pile_row, text="", font=("微软雅黑", 9), bootstyle="secondary")
        self._pile_x_lbl.pack(side="right", padx=(5, 5))
        ttk.Label(pile_row, text="X坐标:", bootstyle=PRIMARY).pack(side="right", padx=(5, 5))

        ttk.Separator(self._soil_content, orient=HORIZONTAL).pack(fill="x", pady=5)
        self._sheet_frame = ttk.Frame(self._soil_content)
        self._sheet_frame.pack(fill="both", expand=True)
        
        # ----------------- tksheet 最简搭建 -----------------
        self._sheet = Sheet(self._sheet_frame,
                            headers=["土层名称", "层厚(m)", "重度(kN/m3)",
                                     "黏聚力c(kPa)", "内摩擦角φ(°)", "极限侧阻力标准值qsik(kPa)", "极限端阻力标准值qpk(kPa)"],
                            row_height=35,
                            header_height=35)
        self._sheet.pack(fill="both", expand=True)
        
        # 开启必备功能：单选、行选、右键菜单、编辑、复制粘贴
        self._sheet.enable_bindings(
            "single_select", "row_select", "copy", "cut", "paste",
            "delete", "undo", "edit_cell", "column_width_resize",
            "rc_select", "rc_insert_row", "rc_delete_row", "rc_popup_menu"
        )
        
        # UI 样式配置 + 右键菜单/快捷键功能中文化（tksheet set_options *_label 项）
        self._sheet.set_options(
            header_bg="#e8f4fd",
            header_font=("微软雅黑", 9, "bold"),
            edit_cell_label="编辑单元格",
            cut_label="剪切",
            copy_label="复制",
            copy_plain_label="复制文本",
            paste_label="粘贴",
            delete_label="删除",
            clear_contents_label="清除内容",
            delete_rows_label="删除行",
            insert_row_label="插入行",
            insert_rows_above_label="上方插入行",
            insert_rows_below_label="下方插入行",
            select_all_label="全选",
            undo_label="撤销",
            redo_label="重做",
        )
        
        # 绑定手动编辑行为，触发保护
        self._sheet.extra_bindings([
            ("end_edit_cell", self._mark_edited), ("rc_insert_row", self._mark_edited),
            ("rc_delete_row", self._mark_edited), ("paste", self._mark_edited),
            ("delete", self._mark_edited), ("undo", self._mark_edited), ("cut", self._mark_edited)
        ])

        # tksheet快捷键功能说明tip
        self.tksheet_tip = ttk.Label(self._soil_content, text="* 可在表格中右键点击行序号，使用表格编辑快捷键。", font=("微软雅黑", 9), bootstyle="secondary")
        self.tksheet_tip.pack( fill="x", pady=10)

        # ================= 4. 初始化动作 =================
        def _on_x_origin_change(*_):
            self._is_edited.clear()  # X坐标变动，重置保护
            self._soil_cache.clear()
            self._refresh_pile_list()
            try: trestle_excel_dict.setdefault("basic", {})["x_origin"] = self._x_origin_var.get().strip()
            except: pass
        self._x_origin_var.trace_add("write", _on_x_origin_change)

        self._read_ground_line()
        if getattr(self, "_ground_line", []):
            trestle_excel_dict["ground_line"] = self._ground_line

    def reload_from_dict(self, new_dict):
        """数据绑定更新：从 new_dict 刷新土层参数，不重建 widget"""
        # 清空缓存状态
        self._soil_cache.clear()
        self._is_edited.clear()
        self._soil_top_cache.clear()
        self._current_pile_no = None
        self._ground_line = []
        if hasattr(self, "_soil_top_var"):
            self._soil_top_var.set("")

        # C2：先加载土层参数显示数据源（内存态 soil_data 优先 → {label}土层表 覆盖表 → 地形插值兜底），
        # 必须早于 _refresh_pile_list，否则地形插值先填充 tksheet 会在后续 _save_current_soil 时覆盖覆盖表
        self._load_pile_soil_sheet(new_dict)

        # 更新 x_origin（trace 回调会自动刷新桩号列表和 trestle_excel_dict）
        new_x = new_dict.get("basic", {}).get("x_origin", "")
        cur_x = self._x_origin_var.get().strip()
        if str(new_x).strip() != cur_x:
            self._x_origin_var.set(str(new_x) if new_x else "")
        else:
            # x 未变但支撑数据可能变了，手动刷新桩号列表
            self._refresh_pile_list()

        # 重新读取地面线
        self._read_ground_line()
        if getattr(self, "_ground_line", []):
            trestle_excel_dict["ground_line"] = self._ground_line

    # ================= 工具与状态方法 =================
    def _mark_edited(self, event=None):
        """一旦表格发生任何改动，标记当前桩号已被手动编辑"""
        if self._current_pile_no is not None:
            self._is_edited.add(self._current_pile_no)

    def _open_plugin(self):
        # 1. 获取参数
        excel_path = trestle_excel_dict.get("_excel_path", "")
        if not excel_path:
            messagebox.showwarning("提示", "找不到项目 Excel")
            return
        sheet = self._get_project_soil_sheet()
        water_level = trestle_excel_dict.get("basic", {}).get("water_level", "")
        
        # 2. 导入插件模块 (请确保模块名与实际文件名一致，不带 .py)
        root = self.winfo_toplevel()
        
        try:
            # 3. 实例化插件窗口 (将其作为当前 root 的子窗口)
            plugin_window = Trestle_Geology_FEM.MainWindow(
                parent=root,
                excel_path=excel_path,
                sheet_name=(sheet + "地形表") if sheet else "",
                water_level=water_level
            )
            
            # 4. 设置模态特性（阻塞主窗口交互，取代你之前的 root.attributes("-disabled", True)）
            plugin_window.transient(root)  # 依附于主窗口，最小化时同步
            plugin_window.grab_set()       # 捕获所有事件，主窗口变为不可交互
            
            # 5. 挂起当前流程，等待插件窗口销毁
            root.wait_window(plugin_window)
            
            # 6. 插件窗口关闭后，自动继续执行以下清理和刷新逻辑（原 _poll_plugin 的 else 分支）
            self._is_edited.clear()  
            self._soil_cache.clear()
            try:
                ep = trestle_excel_dict.get("_excel_path", "")
                if ep:
                    import openpyxl
                    wb = openpyxl.load_workbook(ep, data_only=True)
                    trestle_excel_dict["_wb_sheets"] = list(wb.sheetnames)
                    wb.close()
            except Exception: 
                pass
            trestle_excel_dict.pop("ground_line", None)
            self._on_page_shown()
            
        except Exception as e:
            messagebox.showerror("启动失败", str(e))

    def _poll_plugin(self, proc, root):
        if proc.poll() is None:
            self.after(200, lambda: self._poll_plugin(proc, root))
        else:
            root.attributes("-disabled", False)
            self._is_edited.clear()  # 重新导入地质后，重置保护
            self._soil_cache.clear()
            try:
                ep = trestle_excel_dict.get("_excel_path", "")
                if ep:
                    import openpyxl
                    wb = openpyxl.load_workbook(ep, data_only=True)
                    trestle_excel_dict["_wb_sheets"] = list(wb.sheetnames)
                    wb.close()
            except: pass
            trestle_excel_dict.pop("ground_line", None)
            self._on_page_shown()

    def _get_project_soil_sheet(self):
        # cache 格式：从项目级别取土层标签
        lbl = str(trestle_excel_dict.get("soil_label", "") or "").strip()
        return lbl if lbl and lbl != "/" else ""

    # ================= 核心渲染与数据逻辑 =================
    def _on_page_shown(self):
        self._read_ground_line()
        gl = getattr(self, "_ground_line", [])
        if gl: trestle_excel_dict["ground_line"] = gl
        self._refresh_pile_list()    
        # 延迟执行列宽重置，确保框架大小已定
        self.after(50, self._apply_soil_col_widths)

    def _refresh_pile_list(self):
        if not self.parent_ui or not hasattr(self.parent_ui, "tab_supports"): return
        s = self.parent_ui.tab_supports
        spans = getattr(s, "spans_data", [])
        supports = getattr(s, "support_data", [])
        if not spans: return
        
        try: x0 = float(self._x_origin_var.get() or 0)
        except: x0 = 0.0
        self._pile_xs = [x0]
        for i, span in enumerate(spans):
            dx = span
            if i < len(supports) and supports[i].get("type") == "制动墩": dx += 0.2
            self._pile_xs.append(self._pile_xs[-1] + dx)
            
        values = [f"{i}#" for i in range(len(self._pile_xs))]
        self._pile_combo.configure(values=values)
        if values:
            self._pile_combo.set(values[0])
            self._on_pile_change()

    def _save_current_soil(self):
        """绝对安全的保存逻辑：只认追踪器，不认 UI 状态

        去"层高"列后，各层 elevation 由 土顶标高 − 累计层厚 推导（土顶标高来自 _soil_top_cache/entry）。
        """
        if self._current_pile_no is None:
            return
        try:
            data = self._sheet.get_sheet_data()
            top_str = self._soil_top_var.get().strip() or self._soil_top_cache.get(self._current_pile_no, "")
            try:
                top = float(top_str)
            except (ValueError, TypeError):
                top = None
            layers = []
            cum = 0.0
            for row in data:
                name = str(row[0]).strip() if row[0] not in (None, "") else ""
                vals = [str(v).strip() if v not in (None, "") else "" for v in row[1:]]
                if name and any(v not in ("",) for v in vals):
                    try:
                        thickness = float(vals[0])
                    except (ValueError, TypeError):
                        thickness = 0.0
                    elevation = str(round(top - cum, 2)) if top is not None else ""
                    cum += thickness
                    layers.append({"name": name, "elevation": elevation, "thickness": vals[0],
                                   "unit_weight": vals[1], "cohesion": vals[2],
                                   "friction_angle": vals[3], "qsik": vals[4], "qpk": vals[5]})
            if layers:
                self._soil_cache[self._current_pile_no] = layers
        except Exception as e:
            print(f"[soil] 缓存保存失败: {e}")

    def _on_soil_top_edit(self):
        """土顶标高编辑：记录到该桩号并标记已编辑，重算各层 elevation 保存缓存"""
        if self._current_pile_no is None:
            return
        v = self._soil_top_var.get().strip()
        if v:
            try:
                float(v)
            except (ValueError, TypeError):
                return
            self._soil_top_cache[self._current_pile_no] = v
            self._is_edited.add(self._current_pile_no)
            self._save_current_soil()

    def _on_reset_soil(self):
        """重置为地形值：确认后清除全部桩号 per-pile 编辑（土顶标高 + tksheet），从地形重新推导"""
        from tkinter import messagebox
        if not messagebox.askyesno(
                "确认", "重置为地形值将清除本项目全部桩号手动编辑的土层参数（含土顶标高），\n"
                        "按地形参数重新推导显示。是否继续？",
                parent=self.winfo_toplevel()):
            return
        self._save_current_soil()
        self._soil_cache.clear()
        self._is_edited.clear()
        self._soil_top_cache.clear()
        self._soil_top_var.set("")
        sel = self._pile_combo.get() if hasattr(self, "_pile_combo") else ""
        self._refresh_pile_list()
        if sel and hasattr(self, "_pile_combo") and sel in self._pile_combo['values']:
            self._pile_combo.set(sel)
            self._on_pile_change()

    def _get_soil_cache(self, all_pile_nos=None):
        self._save_current_soil()
        # 确保 _pile_xs 已初始化（可能从未访问过土层 tab）
        if not hasattr(self, "_pile_xs") or not self._pile_xs:
            self._refresh_pile_list()
        if all_pile_nos:
            current_sel = self._pile_combo.get()
            values = list(self._pile_combo['values']) if hasattr(self, '_pile_combo') else []
            for pno in all_pile_nos:
                pile_tag = f"{pno}#"
                if str(pno) not in self._soil_cache and pile_tag in values:
                    self._pile_combo.set(pile_tag)
                    self._refresh_soil_table()
                    self._save_current_soil()
            if current_sel and current_sel in values:
                self._pile_combo.set(current_sel)
                self._refresh_soil_table()
        return dict(self._soil_cache)

    def _on_pile_change(self):
        self._save_current_soil()
        sel = self._pile_combo.get()
        if sel and sel.rstrip("#").isdigit():
            idx = int(sel.rstrip("#"))
            if hasattr(self, "_pile_xs") and idx < len(self._pile_xs):
                self._pile_x_lbl.config(text=f"{self._pile_xs[idx]:.2f}m")
        self._refresh_soil_table()

    def _apply_soil_col_widths(self):
        """固定列宽，配合横向滚动条查看完整数据"""
        widths = [120, 110, 110, 110, 110, 200, 200]
        try:
            self._sheet.set_column_widths(column_widths=widths)
            self._sheet.redraw()
        except Exception:
            pass

    def _refresh_soil_table(self):
        sel = self._pile_combo.get()
        if not sel or not sel.rstrip("#").isdigit(): return
        idx = int(sel.rstrip("#"))
        idx_str = str(idx)
        if not hasattr(self, "_pile_xs") or idx >= len(self._pile_xs): return
            
        # ==== 功能三：从缓存加载已被用户编辑过的数据 ====
        if idx_str in self._is_edited and idx_str in self._soil_cache:
            rows = []
            for p in self._soil_cache[idx_str]:
                rows.append([
                    p.get("name", ""), p.get("thickness", 0),
                    p.get("unit_weight", 0), p.get("cohesion", 0), p.get("friction_angle", 0),
                    p.get("qsik", 0), p.get("qpk", 0)
                ])
            # 不足 10 行时补空行至 10 行
            while len(rows) < 10:
                rows.append(["", "", "", "", "", "", ""])
            self._sheet.set_sheet_data(rows)
            self._current_pile_no = idx_str
            self._apply_soil_col_widths()
            # 土顶标高：编辑值优先，否则取首层 elevation
            _first = self._soil_cache[idx_str][0] if self._soil_cache[idx_str] else {}
            self._soil_top_var.set(self._soil_top_cache.get(idx_str, _first.get("elevation", "")))
            return

        # ==== 未被编辑：从 Excel 解析图纸拓扑并插值 ====
        pile_x = self._pile_xs[idx]
        segments, y_bottom = self._get_segments()
        if not segments:
            self._current_pile_no = idx_str
            # 无地形数据：显示 10 空行提示用户可手动输入
            self._sheet.set_sheet_data([["", "", "", "", "", "", ""] for _ in range(10)])
            self._apply_soil_col_widths()
            self._soil_top_var.set("")
            return
            
        import openpyxl
        layers = []
        for sid, coords in segments:
            y = self._interp_at_x(coords, pile_x)
            if y is not None: layers.append((sid, y))
        layers.sort(key=lambda x: -x[1])
        
        lib = {}
        try:
            excel_path = trestle_excel_dict.get("_excel_path", "")
            if excel_path:
                wb = openpyxl.load_workbook(excel_path, data_only=True)
                if "土层参数库" in wb.sheetnames:
                    ws_lib = wb["土层参数库"]
                    for r in range(2, ws_lib.max_row + 1):
                        lid = str(ws_lib.cell(r, 1).value or "").strip()
                        if lid:
                            lib[lid] = {
                                "soil_name": str(ws_lib.cell(r, 2).value or "").strip(),
                                "unit_weight": float(ws_lib.cell(r, 3).value or 0),
                                "cohesion": float(ws_lib.cell(r, 4).value or 0),
                                "friction_angle": float(ws_lib.cell(r, 5).value or 0),
                                "qsik": float(ws_lib.cell(r, 6).value or 0),
                                "qpk": float(ws_lib.cell(r, 7).value or 0),
                            }
                wb.close()
        except: pass
        
        rows = []
        soil_layers = []
        for i, (sid, top_y) in enumerate(layers):
            bot_y = layers[i + 1][1] if i + 1 < len(layers) else (y_bottom if y_bottom is not None else 0)
            thickness = top_y - bot_y
            p = lib.get(sid, {})
            rows.append([
                p.get("soil_name", f"层{sid}"),
                round(thickness, 2) if thickness > 0 else 0,
                p.get("unit_weight", 0), p.get("cohesion", 0),
                p.get("friction_angle", 0), p.get("qsik", 0), p.get("qpk", 0)
            ])
            soil_layers.append({
                "name": p.get("soil_name", f"层{sid}"),
                "elevation": str(round(top_y, 2)),
                "thickness": str(round(thickness, 2) if thickness > 0 else 0),
                "unit_weight": str(p.get("unit_weight", 0)),
                "cohesion": str(p.get("cohesion", 0)),
                "friction_angle": str(p.get("friction_angle", 0)),
                "qsik": str(p.get("qsik", 0)),
                "qpk": str(p.get("qpk", 0)),
            })

        # 不足 10 行时补空行至 10 行
        while len(rows) < 10:
            rows.append(["", "", "", "", "", "", ""])
        self._sheet.set_sheet_data(rows)
        self._current_pile_no = idx_str
        self._apply_soil_col_widths()
        # 土顶标高 = 首层顶标高（未编辑，显示地形推导值，不写入 _soil_top_cache）
        self._soil_top_var.set(str(round(layers[0][1], 2)) if layers else "")
        # 直接写入 cache，不依赖 _save_current_soil 从 UI 读取
        if soil_layers:
            self._soil_cache[idx_str] = soil_layers

    def _load_pile_soil_sheet(self, new_dict=None):
        """土层参数显示数据源（C2）：
        1) 项目内存态 substructure[*].soil_data 非空 → 优先采用（会话编辑/上次保存）；
        2) 否则读 {soil_label}土层表 覆盖表；
        3) 均无 → 保持空，由地形插值兜底。
        加载结果标记 _is_edited（视为 per-pile 覆盖数据），土顶标高写入 _soil_top_cache。
        """
        import openpyxl
        sub = (new_dict or {}).get("substructure", {})
        loaded = 0
        for sk in sorted(sub, key=lambda k: int(k) if k.isdigit() else 0):
            sd = sub.get(sk, {}).get("soil_data") or {}
            if not sd:
                continue
            layers = []
            for nm, v in sd.items():
                vv = [str(x) for x in (v or [])]
                layers.append({
                    "name": nm,
                    "thickness": vv[0] if len(vv) > 0 else "",
                    "unit_weight": vv[1] if len(vv) > 1 else "",
                    "cohesion": vv[2] if len(vv) > 2 else "",
                    "friction_angle": vv[3] if len(vv) > 3 else "",
                    "qsik": vv[4] if len(vv) > 4 else "",
                    "qpk": vv[5] if len(vv) > 5 else "",
                })
            if layers:
                self._soil_cache[sk] = layers
                self._is_edited.add(sk)
                ge = sub.get(sk, {}).get("ground_elevation")
                if ge not in (None, "", "/"):
                    self._soil_top_cache[sk] = str(ge)
                loaded += 1
        if loaded:
            return
        # 从 {soil_label}土层表 覆盖表读取
        excel_path = trestle_excel_dict.get("_excel_path", "")
        sheet_name = self._get_project_soil_sheet()
        if not excel_path or not sheet_name:
            return
        target = sheet_name + "土层表"
        try:
            wb = openpyxl.load_workbook(excel_path, data_only=True)
            if target not in wb.sheetnames:
                wb.close()
                return
            ws = wb[target]
            hm = {}
            for c in ws[1]:
                if c.value:
                    hm[str(c.value).strip()] = c.column

            def _split(v):
                return [x.strip() for x in str(v or "").split(",") if x.strip()]

            for r in range(2, ws.max_row + 1):
                pile_tag = str(ws.cell(r, hm.get("桩号", 1)).value or "").strip()
                if not pile_tag or not pile_tag.rstrip("#").isdigit():
                    continue
                idx = pile_tag.rstrip("#")
                names = _split(ws.cell(r, hm.get("土层名称", 3)).value)
                thicks = _split(ws.cell(r, hm.get("土层厚度(m)", 4)).value)
                weights = _split(ws.cell(r, hm.get("土层重度(kN/m3)", 5)).value)
                cs = _split(ws.cell(r, hm.get("黏聚力c(kPa)", 6)).value)
                phis = _split(ws.cell(r, hm.get("内摩擦角φ(°)", 7)).value)
                qsiks = _split(ws.cell(r, hm.get("桩极限侧阻力标准值qsik(kPa)", 8)).value)
                qpks = _split(ws.cell(r, hm.get("桩极限端阻力标准值qpk(kPa)", 9)).value)
                layers = []
                for i, nm in enumerate(names):
                    layers.append({
                        "name": nm,
                        "thickness": thicks[i] if i < len(thicks) else "",
                        "unit_weight": weights[i] if i < len(weights) else "",
                        "cohesion": cs[i] if i < len(cs) else "",
                        "friction_angle": phis[i] if i < len(phis) else "",
                        "qsik": qsiks[i] if i < len(qsiks) else "",
                        "qpk": qpks[i] if i < len(qpks) else "",
                    })
                if layers:
                    self._soil_cache[idx] = layers
                    self._is_edited.add(idx)
                    top = str(ws.cell(r, hm.get("土顶标高(m)", 2)).value or "").strip()
                    if top:
                        self._soil_top_cache[idx] = top
            wb.close()
        except Exception as e:
            print(f"[soil] 读取土层表覆盖表失败: {e}")

    # ================= 地面线插值与基础读取 =================
    def _get_segments(self):
        import openpyxl, re
        segments, y_bottom = [], None
        excel_path, sheet_name = trestle_excel_dict.get("_excel_path", ""), self._get_project_soil_sheet()
        if not excel_path or not sheet_name: return segments, y_bottom
        try:
            wb = openpyxl.load_workbook(excel_path, data_only=True)
            sheet_name = sheet_name + "地形表"
            if sheet_name not in wb.sheetnames: return segments, y_bottom
            ws = wb[sheet_name]
            for r in range(2, ws.max_row + 1):
                raw_a = ws.cell(r, 1).value
                if str(raw_a or "").strip() == "计算底边线y坐标":
                    bv = ws.cell(r, 2).value
                    if bv is not None:
                        try: y_bottom = float(bv)
                        except ValueError: pass
                    continue
                sid, coords = str(raw_a or "").strip(), []
                for c in range(2, ws.max_column + 1):
                    v = ws.cell(r, c).value
                    if v is None or str(v).strip() in ("", "/"): continue
                    m = re.match(r"\(\s*([\d.-]+)\s*,\s*([\d.-]+)\s*\)", str(v))
                    if m: coords.append((float(m.group(1)), float(m.group(2))))
                    else: break
                if len(coords) >= 2: segments.append((sid, coords))
            wb.close()
        except: pass
        return segments, y_bottom

    def _interp_at_x(self, coords, x_val):
        """折线插值：复用 General.Geometry.polyline_y_at_x（纯几何）"""
        return polyline_y_at_x(x_val, coords)

    def _read_ground_line(self):
        import openpyxl, re
        cached = trestle_excel_dict.get("ground_line", [])
        if cached:
            self._ground_line = cached
            return
            
        sheet_name = self._get_project_soil_sheet()
        excel_path = trestle_excel_dict.get("_excel_path", "")
        if not sheet_name or not excel_path:
            self._ground_line = []
            return
        
        try:
            wb = openpyxl.load_workbook(excel_path, data_only=True)
            sheet_name = sheet_name + "地形表"
            if sheet_name not in wb.sheetnames: return
            ws, boundaries = wb[sheet_name], []
            for r in range(2, ws.max_row + 1):
                if str(ws.cell(r, 1).value or "").strip() == "计算底边线y坐标":
                    continue
                coords = []
                for c in range(2, ws.max_column + 1):
                    v = ws.cell(r, c).value
                    if v is None or str(v).strip() in ('', '/'): continue
                    m = re.match(r'\(\s*([\d.-]+)\s*,\s*([\d.-]+)\s*\)', str(v))
                    if m: coords.append((float(m.group(1)), float(m.group(2))))
                if len(coords) >= 2: boundaries.append(coords)
            wb.close()
            if not boundaries: return

            all_x = sorted(set(p[0] for pts in boundaries for p in pts))
            ground = []
            for x in all_x:
                valid_ys = [y for y in [self._interp_at_x(pts, x) for pts in boundaries] if y is not None]
                if valid_ys: ground.append((x, max(valid_ys)))
                
            self._ground_line = ground
            trestle_excel_dict["ground_line"] = ground
        except: pass

# ======================== 连接形式 Tab 页 ========================
class ConnectionTab(ttk.Frame):
    """连接形式参数设置

    四个弹性连接切换按钮 + ConnectionSettingFrame 嵌入子 frame
    桩底边界 + 释放梁端约束 + 土弹簧预留
    """

    CONN_ITEMS = [
        ("桥面系横肋与纵梁", "conn_deck_beam"),
        ("纵梁与支撑架", "conn_beam_brace"),
        ("纵梁与横向分配梁", "conn_beam_dist"),
        ("桩顶分配梁与桩顶", "conn_dist_pile"),
    ]


    def __init__(self, parent, parent_ui=None):
        super().__init__(parent)
        self.parent_ui = parent_ui
        self.connection_cache = {}
        self._conn_frames = {}
        self._current_conn_key = None
        self._conn_stiffness_db = trestle_excel_dict.get("connection_stiffness", {})
        self._build_ui()
        self.refresh_cache()

    def reload_from_dict(self, new_dict):
        """从新项目数据刷新 UI（统一 cache 格式）"""
        conn = new_dict.get("connection", {})
        # 刷新连接刚度库：__init__ 仅在创建 tab 时加载一次，切项目后需跟随新项目数据更新
        self._conn_stiffness_db = new_dict.get("connection_stiffness", {}) or {}
        # 更新弹性连接类型（值为 dict：{"type": 名称, 刚度键...}）
        for label, key in self.CONN_ITEMS:
            frame = self._conn_frames.get(key)
            if frame:
                val = conn.get(label, "自定义")
                if isinstance(val, dict):
                    new_type = val.get("type", "自定义")
                    stiffness = {k: v for k, v in val.items() if k != "type"}
                    frame.set_type(new_type)
                    if stiffness:
                        frame.set_stiffness(stiffness)
                else:
                    frame.set_type(val)
        # 更新桩底边界
        pile_bnd = conn.get("桩底", "111111")
        self._pile_bnd_var.set(pile_bnd if pile_bnd and pile_bnd != "/" else "111111")
        # 更新释放梁端约束
        release = conn.get("贝雷梁释放梁端约束", "0000000,0000000")
        self._release_bnd_var.set(release if release and release != "/" else "0000000,0000000")
        # 更新桩边界计算方法
        soil_thickness = conn.get("土弹簧分层厚度(m)", "/")
        self._bnd_method_var.set("土弹簧" if soil_thickness and soil_thickness not in ("/", "") else "桩底边界")
        self._switch_boundary_method()
        self.refresh_cache()

    def _load_conn_options(self):
        """从连接形式 Sheet 读取可选连接名称"""
        options = ["自定义"]
        try:
            for name in self._conn_stiffness_db.keys():
                if name and name not in options:
                    options.append(name)
        except:
            pass
        return options

    def _build_ui(self):
        """构建连接形式设置界面"""
        # 提前初始化所有 StringVar（set_type → refresh_cache 回调需要它们已存在）
        pile_bnd_val = trestle_excel_dict.get("connection", {}).get("桩底", "111111")
        self._pile_bnd_var = tk.StringVar(value=pile_bnd_val if pile_bnd_val and pile_bnd_val != "/" else "111111")
        release_val = trestle_excel_dict.get("connection", {}).get("贝雷梁释放梁端约束", "0000000,0000000")
        self._release_bnd_var = tk.StringVar(value=release_val if release_val and release_val != "/" else "0000000,0000000")

        # ── Card 1：连接设置（弹性连接 + 释放梁端约束） ──
        card_conn, content_conn = make_card_frame(self, "连接设置")

        # 弹性连接（LabelFrame）
        conn_lf = ttk.LabelFrame(content_conn, text="弹性连接设置", padding=10)
        conn_lf.pack(fill="x", pady=(0, 8))
        self._tab_frame = ttk.Frame(conn_lf)
        self._tab_frame.pack(fill="x", pady=(0, 8))
        self._conn_btns = {}
        for label, key in self.CONN_ITEMS:
            btn = ttk.Button(self._tab_frame, text=label, bootstyle="primary",
                             command=lambda k=key: self._on_conn_btn(k))
            if key == "conn_beam_brace":
                btn.config(state="normal")
            btn.pack(side="left", padx=2, fill="x", expand=True)
            self._conn_btns[key] = btn
        # 刚度参数区：Canvas + 横向滚动条（参照 CofferDam 模式）
        self._conn_canvas = tk.Canvas(conn_lf, borderwidth=0, highlightthickness=0)
        self._conn_hscroll = ttk.Scrollbar(conn_lf, orient="horizontal",
                                            command=self._conn_canvas.xview)
        self._conn_canvas.configure(xscrollcommand=self._conn_hscroll.set)
        self._conn_canvas.pack(fill="x", expand=True)

        self._conn_detail_frame = ttk.Frame(self._conn_canvas)
        self._conn_window_id = self._conn_canvas.create_window(
            (0, 0), window=self._conn_detail_frame, anchor="nw")

        def _on_detail_cfg(e):
            self._conn_canvas.configure(scrollregion=self._conn_canvas.bbox("all"))
        self._conn_detail_frame.bind("<Configure>", _on_detail_cfg)

        def _on_canvas_cfg(e):
            fw = self._conn_detail_frame.winfo_reqwidth()
            self._conn_canvas.itemconfig(self._conn_window_id, width=max(e.width, fw))
        self._conn_canvas.bind("<Configure>", _on_canvas_cfg)

        conn_options = self._load_conn_options()
        conns_list = []
        for name, stiff in self._conn_stiffness_db.items():
            entry = {"name": name}
            entry.update(stiff)
            conns_list.append(entry)
        for label, key in self.CONN_ITEMS:
            preset_type = trestle_excel_dict.get("connection", {}).get(label, "自定义")
            preset_stiffness = self._conn_stiffness_db.get(preset_type, {})
            frame = ConnectionSettingFrame(self._conn_detail_frame,
                                           preset_type=preset_type,
                                           stiffness=preset_stiffness,
                                           conn_options=conn_options,
                                           on_change=self.refresh_cache)
            frame.set_conns_data(conns_list)
            if preset_type and preset_type != "自定义":
                matched = self._conn_stiffness_db.get(preset_type, {})
                if matched:
                    frame.set_stiffness(matched)
            frame.set_type(preset_type)
            self._conn_frames[key] = frame
        self._on_conn_btn(self.CONN_ITEMS[0][1])

        # 释放梁端约束（LabelFrame）
        release_lf = ttk.LabelFrame(content_conn, text="贝雷梁释放梁端约束", padding=10)
        release_lf.pack(fill="x", pady=(10, 0))
        self._release_frame = BeamReleaseFrame(release_lf, release_var=self._release_bnd_var,
                                                on_change=self.refresh_cache)
        self._release_frame.pack(fill="x", expand=True)

        # ── Card 2：桩边界设置 ──
        card_bnd, content_bnd = make_card_frame(self, "桩边界设置")

        # 从 Excel 读取桩底和土弹簧设置（互斥）
        pile_bnd_val = trestle_excel_dict.get("connection", {}).get("桩底", "/")
        soil_thickness_val = trestle_excel_dict.get("connection", {}).get("土弹簧分层厚度(m)", "/")

        # 初始计算方法：根据哪个列有值来决定
        if soil_thickness_val and soil_thickness_val not in ("/", ""):
            init_method = "土弹簧"
        else:
            init_method = "桩底边界"

        # 下拉行
        bnd_options = ["桩底边界", "土弹簧"]
        self._bnd_method_var = tk.StringVar(value=init_method)
        method_row = ttk.Frame(content_bnd)
        method_row.pack(fill="x", pady=(0, 8))
        ttk.Label(method_row, text="计算方法:", width=14, bootstyle=PRIMARY).pack(side="left")
        bnd_combo = ttk.Combobox(method_row, textvariable=self._bnd_method_var,
                                 values=bnd_options, state="readonly", width=18)
        bnd_combo.pack(side="left")
        bnd_combo.bind("<<ComboboxSelected>>", lambda e: self._switch_boundary_method())

        # 桩底边界（LabelFrame）
        self._bnd_pile_frame = ttk.LabelFrame(content_bnd, text="桩底边界", padding=10)
        self._bnd_pile_frame.pack(fill="x")
        self._pile_boundary = PileBoundaryFrame(self._bnd_pile_frame, boundary_var=self._pile_bnd_var,
                                                on_change=self.refresh_cache)
        self._pile_boundary.pack(fill="x")

        # 土弹簧（LabelFrame）
        self._bnd_soil_frame = ttk.LabelFrame(content_bnd, text="土弹簧", padding=10)
        soil_init = soil_thickness_val if soil_thickness_val and soil_thickness_val not in ("/", "") else "1.0"
        self._soil_spring_thickness = create_label_entry(
            self._bnd_soil_frame, "土弹簧分层厚度(m):", 0,
            default_val=soil_init, label_width=16, entry_width=12)

        self._switch_boundary_method()

    def _on_conn_btn(self, key):
        """切换弹性连接显示"""
        # 隐藏所有 frame
        for k, frame in self._conn_frames.items():
            frame.pack_forget()
        # 显示选中
        if key in self._conn_frames:
            self._conn_frames[key].pack(fill="x", expand=True)
        self._current_conn_key = key
        # 延迟刷新画布尺寸和滚动区域
        def _adjust():
            self._conn_canvas.update_idletasks()
            fw = self._conn_detail_frame.winfo_reqheight()
            self._conn_canvas.configure(height=fw + 4)
            cw = self._conn_canvas.winfo_width()
            self._conn_canvas.itemconfig(
                self._conn_window_id,
                width=max(cw, self._conn_detail_frame.winfo_reqwidth()))
            self._conn_canvas.configure(scrollregion=self._conn_canvas.bbox("all"))
            # 内容超宽时显示滚动条
            if self._conn_detail_frame.winfo_reqwidth() > cw:
                self._conn_hscroll.pack(side="bottom", fill="x")
            else:
                self._conn_hscroll.pack_forget()
        self._conn_canvas.after(10, _adjust)

    def _switch_boundary_method(self):
        """切换桩边界计算方法（显示/隐藏对应 frame）"""
        method = self._bnd_method_var.get()
        show_pile = (method == "桩底边界")
        for f_name in ('_bnd_pile_frame', '_bnd_soil_frame'):
            f = getattr(self, f_name, None)
            if f:
                f.pack_forget()
        if show_pile and hasattr(self, '_bnd_pile_frame'):
            self._bnd_pile_frame.pack(fill="x")
        elif hasattr(self, '_bnd_soil_frame'):
            self._bnd_soil_frame.pack(fill="x")
        self.refresh_cache()

    def refresh_cache(self):
        """重建连接形式缓存（保存类型名 + 完整刚度值）"""
        self.connection_cache = {}
        for label, key in self.CONN_ITEMS:
            frame = self._conn_frames.get(key)
            if frame:
                conn_type = frame.get_type()
                stiffness = frame.get_stiffness()
                self.connection_cache[label] = {"type": conn_type}
                self.connection_cache[label].update(stiffness)
        # 按选中的桩边界计算方法写入缓存（桩底和土弹簧互斥）
        bnd_method = self._bnd_method_var.get() if hasattr(self, '_bnd_method_var') else "桩底边界"
        if bnd_method == "桩底边界":
            self.connection_cache["桩底"] = self._pile_bnd_var.get()
            self.connection_cache["土弹簧分层厚度(m)"] = "/"
        else:
            self.connection_cache["桩底"] = "/"
            thickness = "1.0"
            if hasattr(self, '_soil_spring_thickness'):
                thickness = self._soil_spring_thickness.get() or "1.0"
            self.connection_cache["土弹簧分层厚度(m)"] = thickness
            self.connection_cache["土弹簧"] = {"thickness": thickness}  # 供 MCT 生成使用
        self.connection_cache["贝雷梁释放梁端约束"] = self._release_bnd_var.get()


# ======================== 环境荷载 Tab 页 ========================
class EnvironmentLoadTab(ttk.Frame):
    def __init__(self, parent, parent_ui=None):
        super().__init__(parent)
        
        # 定义全局间距属性
        self.padx = 5
        self.pady = 10
        self.label_width = 20

        # ======================== 1. 顶部：模块切换按钮区 ========================
        self.btn_frame = ttk.Frame(self)
        self.btn_frame.pack(fill="x", pady=(15,5), padx=10)

        self.tab_names = ["风荷载", "水流力","波浪力"]
        self.tab_btns = {}
        for name in self.tab_names:
            btn = ttk.Button(self.btn_frame, text=name, command=lambda n=name: self.switch_page(n))
            btn.pack(side="left", padx=2, fill="x", expand=True)
            self.tab_btns[name] = btn

        self.inner = ttk.Frame(self)
        self.inner.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.pages = {
            "风荷载": ttk.Frame(self.inner),
            "水流力": ttk.Frame(self.inner),
            "波浪力": ttk.Frame(self.inner)
        }

        # 调用方法创建各页面内容
        self.create_wind_page()
        self.create_water_page()
        self.create_wave_page()

        # 初始化环境缓存
        self.environment_cache = {}
        self.refresh_cache()
        # 初始化默认显示第一页
        self.switch_page("风荷载")

    def reload_from_dict(self, new_dict):
        """从新项目数据刷新 UI（统一 cache 格式：environment.wind/water/wave）"""
        env = new_dict.get("environment", {})
        wind = env.get("wind", {})
        water = env.get("water", {})
        wave = env.get("wave", {})

        # 风荷载
        _wind_spec = wind.get("spec", "")
        self.wind_fomula_combo.set(_wind_spec if _wind_spec else "《公路桥梁抗风设计规范》")
        self.design_wind_entry.delete(0, tk.END)
        self.design_wind_entry.insert(0, wind.get("design_wind_speed", ""))
        _gc = wind.get("ground_category", "")
        if _gc:
            try: self.ground_category.current(self.ground_category["values"].index(_gc))
            except ValueError: self.ground_category.current(0)
        self.kt_entry.delete(0, tk.END)
        self.kt_entry.insert(0, wind.get("terrain_factor", ""))
        self.H_entry.delete(0, tk.END)
        self.H_entry.insert(0, wind.get("girder_height", ""))
        self.kh_entry.delete(0, tk.END)
        self.kh_entry.insert(0, wind.get("transverse_coeff", ""))
        _wf = wind.get("formula", "")
        if _wf:
            try: self.wind_formula.current(self.wind_formula["values"].index(_wf))
            except ValueError: self.wind_formula.current(0)
        # 水流力
        _water_spec = water.get("spec", "")
        self.water_fomula_combo.set(_water_spec if _water_spec else "《港口工程荷载规范》")
        self.flow_velocity_entry.delete(0, tk.END)
        self.flow_velocity_entry.insert(0, water.get("flow_velocity", ""))
        # 波浪力
        _wave_spec = wave.get("spec", "")
        self.wave_fomula_combo.set(_wave_spec if _wave_spec else "《海港水文规范》")
        # 波浪力参数获取
        for _attr, _key in (("wave_height_entry", "wave_height"),
                            ("wave_period_entry", "wave_period"),
                            ("wave_length_entry", "wave_length")):
            _entry = getattr(self, _attr, None)
            if _entry is None:
                continue
            if str(_entry.cget("state")) == "disabled":
                continue
            _val = wave.get(_key, "")
            if _val in (None, "/"):
                _val = ""
            _entry.delete(0, tk.END)
            _entry.insert(0, str(_val))
        self.refresh_cache()

    def refresh_cache(self):
        """重建环境缓存 — 数值类型"""
        def _f(widget, default=None):
            try:
                v = widget.get()
                return float(v) if v.strip() else default
            except: return default
        def _s(widget, default=""):
            try: return widget.get()
            except: return default
        try:
            self.environment_cache = {
                "wind": {
                    "spec": self.wind_fomula_var.get(),
                    "design_wind_speed": _f(self.design_wind_entry) if hasattr(self, 'design_wind_entry') else None,
                    "ground_category": _s(self.ground_category) if hasattr(self, 'ground_category') else "",
                    "terrain_factor": _f(self.kt_entry) if hasattr(self, 'kt_entry') else None,
                    "girder_height": _f(self.H_entry) if hasattr(self, 'H_entry') else None,
                    "transverse_coeff": _f(self.kh_entry) if hasattr(self, 'kh_entry') else None,
                    "formula": _s(self.wind_formula) if hasattr(self, 'wind_formula') else "",
                },
                "water": {
                    "spec": self.water_fomula_var.get(),
                    "flow_velocity": _f(self.flow_velocity_entry) if hasattr(self, 'flow_velocity_entry') else None,
                },
                "wave": {
                    "spec": self.wave_fomula_var.get(),
                    "wave_height": _s(self.wave_height_entry) if (hasattr(self, 'wave_height_entry') and self.wave_calc_var.get()) else "/",
                    "wave_period": _s(self.wave_period_entry) if (hasattr(self, 'wave_period_entry') and self.wave_calc_var.get()) else "/",
                    "wave_length": _s(self.wave_length_entry) if (hasattr(self, 'wave_length_entry') and self.wave_calc_var.get()) else "/",
                }
            }
        except Exception as e:
            print(f"environment_cache 刷新异常: {e}")

    # ======================== 切换逻辑 ========================
    def switch_page(self, page_name):
        """控制下方 Frame 切换，实现紧贴上方对齐"""
        for name, frame in self.pages.items():
            if name == page_name:
                frame.pack(fill="x", anchor="n") 
            else:
                frame.pack_forget()
        

    # ======================== 页面内容构建函数 ========================
    def create_wind_page(self):
        page_frame = self.pages["风荷载"]
        wc, self.wind_frame = make_card_frame(page_frame, "风荷载")

        # 选择规范
        fomula_frame = ttk.Frame(self.wind_frame)
        fomula_frame.pack(fill=X, padx=10, pady=(0, 10))
        ttk.Label(fomula_frame, text="选择设计规范:", width=15).pack(side=LEFT)
        self.wind_fomula_var = tk.StringVar()
        self.wind_fomula_combo = ttk.Combobox(fomula_frame, textvariable=self.wind_fomula_var,
            values=["《公路桥梁抗风设计规范》"],
            state="readonly", width=50)
        self.wind_fomula_combo.pack(side=LEFT, padx=5)
        _wind_spec = trestle_excel_dict.get("environment", {}).get("wind", {}).get("spec", "")
        self.wind_fomula_combo.set(_wind_spec if _wind_spec else "《公路桥梁抗风设计规范》")

        # 动态内容容器
        self.wind_params = tk.Frame(self.wind_frame, bg="white",
                                    highlightbackground="#ccd1d8", highlightthickness=1)
        self.wind_params.pack(fill=X, padx=15, pady=5)

        self.highway_wind_frame = ttk.Frame(self.wind_params)
        self.highway_wind_frame.pack(fill=X, padx=5, pady=5)
        DBFL_lst = ["A:海面、海岸、开阔水面", "B:田野、乡村、丛林、平坦开阔地",
                    "C:树木及地层建筑密集区、平缓丘陵地", "D:中高层建筑密集区、起伏较大的丘陵地"]
        formula_lst = ["Ud = kf·kt·kh·U10", "Ud = kf·(Z/10)^α0·Us10"]

        _wind_cache = trestle_excel_dict.get("environment", {}).get("wind", {})
        self.design_wind_entry = create_label_entry(self.highway_wind_frame, "设计风速(m/s):", 0,
                                    default_val=_wind_cache.get("design_wind_speed", ""), label_width=self.label_width, entry_width=25, padx=self.padx, pady=self.pady)
        self.design_wind_entry.bind("<FocusOut>", lambda _: self.refresh_cache())
        ttk.Label(self.highway_wind_frame, text="地表分类:", width=self.label_width, bootstyle=PRIMARY).grid(row=1, column=0, sticky="w", padx=5, pady=3)
        self.ground_category = ttk.Combobox(self.highway_wind_frame, values=DBFL_lst, width=30, state="readonly")
        # 从 Excel 读取地面分类，匹配则选中，否则 default 0
        _gc_val = _wind_cache.get("ground_category", "")
        _gc_idx = 0
        if _gc_val:
            try: _gc_idx = DBFL_lst.index(_gc_val)
            except ValueError: _gc_idx = 0
        self.ground_category.current(_gc_idx)
        self.ground_category.grid(row=1, column=1, sticky="w", padx=5, pady=3)
        self.kt_entry = create_label_entry(self.highway_wind_frame, "地形条件系数:", 2, default_val=_wind_cache.get("terrain_factor", ""), label_width=self.label_width, entry_width=25, padx=self.padx, pady=self.pady)
        self.kt_entry.bind("<FocusOut>", lambda _: self.refresh_cache())
        self.H_entry = create_label_entry(self.highway_wind_frame, "主梁基准高度(m):", 3, default_val=_wind_cache.get("girder_height", ""), label_width=self.label_width, entry_width=25, padx=self.padx, pady=self.pady)
        self.H_entry.bind("<FocusOut>", lambda _: self.refresh_cache())
        self.kh_entry = create_label_entry(self.highway_wind_frame, "主梁横向力系数:", 4, default_val=_wind_cache.get("transverse_coeff", ""), label_width=self.label_width, entry_width=25, padx=self.padx, pady=self.pady)
        self.kh_entry.bind("<FocusOut>", lambda _: self.refresh_cache())
        ttk.Label(self.highway_wind_frame, text="设计基准风速计算公式:", width=self.label_width, bootstyle=PRIMARY).grid(row=5, column=0, sticky="w", padx=5, pady=3)
        self.wind_formula = ttk.Combobox(self.highway_wind_frame, values=formula_lst, width=30, state="readonly")
        _wf_val = _wind_cache.get("formula", "")
        _wf_idx = 0
        if _wf_val:
            try: _wf_idx = formula_lst.index(_wf_val)
            except ValueError: _wf_idx = 0
        self.wind_formula.current(_wf_idx)
        self.wind_formula.grid(row=5, column=1, sticky="w", padx=5, pady=3)

        self.custom_wind_frame = ttk.Frame(self.wind_params)
        ttk.Label(self.custom_wind_frame, text="待补充...", foreground="gray", font=("微软雅黑", 10, "italic")).pack(anchor="w", pady=20, padx=10)

        def on_fomula_change(event=None):
            fomula = self.wind_fomula_var.get()
            self.highway_wind_frame.pack_forget()
            self.custom_wind_frame.pack_forget()
            if fomula == "自定义":
                self.custom_wind_frame.pack(fill=X, padx=5, pady=5)
            else:
                self.highway_wind_frame.pack(fill=X, padx=5, pady=5)
        self.wind_fomula_combo.bind("<<ComboboxSelected>>", on_fomula_change)
        on_fomula_change()

    def create_water_page(self):
        page_frame = self.pages["水流力"]
        wc2, self.water_frame = make_card_frame(page_frame, "水流力")

        fomula_frame = ttk.Frame(self.water_frame)
        fomula_frame.pack(fill=X, padx=10, pady=(0, 10))
        ttk.Label(fomula_frame, text="选择设计规范:", width=15).pack(side=LEFT)
        self.water_fomula_var = tk.StringVar()
        self.water_fomula_combo = ttk.Combobox(fomula_frame, textvariable=self.water_fomula_var,
            values=["《港口工程荷载规范》"],
            state="readonly", width=50)
        self.water_fomula_combo.pack(side=LEFT, padx=5)
        _water_spec = trestle_excel_dict.get("environment", {}).get("water", {}).get("spec", "")
        self.water_fomula_combo.set(_water_spec if _water_spec else "《港口工程荷载规范》")

        # 动态内容容器 (带边框)
        water_params = tk.Frame(self.water_frame, bg="white",
                                highlightbackground="#ccd1d8", highlightthickness=1)
        water_params.pack(fill=X, padx=15, pady=5)

        self.port_water_frame = ttk.Frame(water_params)
        self.port_water_frame.pack(fill=X, padx=5, pady=5)
        self.flow_velocity_entry = create_label_entry(self.port_water_frame, "设计流速(m/s):", 0, default_val=trestle_excel_dict.get("environment", {}).get("water", {}).get("flow_velocity", ""), label_width=self.label_width, entry_width=25, padx=self.padx, pady=self.pady)
        self.flow_velocity_entry.bind("<FocusOut>", lambda _: self.refresh_cache())

        self.custom_water_frame = ttk.Frame(water_params)
        ttk.Label(self.custom_water_frame, text="待补充...", foreground="gray", font=("微软雅黑", 10, "italic")).pack(anchor="w", pady=20, padx=10)

        def on_fomula_change(event=None):
            fomula = self.water_fomula_var.get()
            self.port_water_frame.pack_forget()
            self.custom_water_frame.pack_forget()
            if fomula == "自定义":
                self.custom_water_frame.pack(fill=X, padx=5, pady=5)
            else:
                self.port_water_frame.pack(fill=X, padx=5, pady=5)
        self.water_fomula_combo.bind("<<ComboboxSelected>>", on_fomula_change)
        on_fomula_change()
        

    def create_wave_page(self):
        page_frame = self.pages["波浪力"]
        wc3, self.wave_frame = make_card_frame(page_frame, "波浪力")

        # 计算波浪力复选框
        _wave_cache = trestle_excel_dict.get("environment", {}).get("wave", {})
        wave_height_val = _wave_cache.get("wave_height", "")
        wave_period_val = _wave_cache.get("wave_period", "")
        wave_length_val = _wave_cache.get("wave_length", "")
        
        # 判断波浪力参数是否有效
        def _is_valid_wave_param(val):
            """检查波浪力参数是否有效"""
            if not val or val == "/":
                return False
            parts = [x.strip() for x in str(val).split(",")]
            if len(parts) != 3:
                return False
            try:
                for p in parts:
                    float(p)
                return True
            except (ValueError, TypeError):
                return False
        
        # 检查波浪力参数是否有效
        wave_params_valid = (_is_valid_wave_param(wave_height_val) and 
                           _is_valid_wave_param(wave_period_val) and 
                           _is_valid_wave_param(wave_length_val))
        
        # 根据参数有效性决定是否开启波浪力计算
        self.wave_calc_var = tk.BooleanVar(value=wave_params_valid)
        calc_frame = ttk.Frame(self.wave_frame)
        calc_frame.pack(fill=X, padx=10, pady=(5, 0))
        self.wave_calc_cb = ttk.Checkbutton(calc_frame, text="计算波浪力",
                                              variable=self.wave_calc_var,
                                              bootstyle="success-round-toggle",
                                              state="disabled",
                                              command=self._on_wave_calc_toggle)
        self.wave_calc_cb.pack(side=LEFT)
        self.wave_calc_cb.pack(side=LEFT)

        fomula_frame = ttk.Frame(self.wave_frame)
        fomula_frame.pack(fill=X, padx=10, pady=(0, 10))
        ttk.Label(fomula_frame, text="选择设计规范:", width=15).pack(side=LEFT)
        self.wave_fomula_var = tk.StringVar()
        self.wave_fomula_combo = ttk.Combobox(fomula_frame, textvariable=self.wave_fomula_var,
            values=["《海港水文规范》"],
            state="readonly", width=50)
        self.wave_fomula_combo.pack(side=LEFT, padx=5)
        _wave_spec = _wave_cache.get("spec", "")
        self.wave_fomula_combo.set(_wave_spec if _wave_spec else "《海港水文规范》")

        # 动态内容容器 (带边框)
        wave_params = tk.Frame(self.wave_frame, bg="white",
                                highlightbackground="#ccd1d8", highlightthickness=1)
        wave_params.pack(fill=X, padx=15, pady=5)

        self.port_wave_frame = ttk.Frame(wave_params)
        self.port_wave_frame.pack(fill=X, padx=5, pady=5)
        self.wave_period_entry = create_label_entry(self.port_wave_frame, "波浪周期(m):", 0, default_val=wave_height_val if wave_height_val and wave_height_val != "/" else "", label_width=self.label_width, entry_width=25, padx=self.padx, pady=self.pady)
        self.wave_period_entry.bind("<FocusOut>", lambda _: self.refresh_cache())
        self.wave_length_entry = create_label_entry(self.port_wave_frame, "设计波长(s):", 1, default_val=wave_period_val if wave_period_val and wave_period_val != "/" else "", label_width=self.label_width, entry_width=25, padx=self.padx, pady=self.pady)
        self.wave_length_entry.bind("<FocusOut>", lambda _: self.refresh_cache())
        self.wave_height_entry = create_label_entry(self.port_wave_frame, "设计波高(m):", 2, default_val=wave_length_val if wave_length_val and wave_length_val != "/" else "", label_width=self.label_width, entry_width=25, padx=self.padx, pady=self.pady)
        self.wave_height_entry.bind("<FocusOut>", lambda _: self.refresh_cache())

        # 初始化 disabled 状态
        self._on_wave_calc_toggle()
        self.custom_wave_frame = ttk.Frame(wave_params)
        ttk.Label(self.custom_wave_frame, text="待补充...", foreground="gray", font=("微软雅黑", 10, "italic")).pack(anchor="w", pady=20, padx=10)

        def on_fomula_change(event=None):
            fomula = self.wave_fomula_var.get()
            self.port_wave_frame.pack_forget()
            self.custom_wave_frame.pack_forget()
            if fomula == "自定义":
                self.custom_wave_frame.pack(fill=X, padx=5, pady=5)
            else:
                self.port_wave_frame.pack(fill=X, padx=5, pady=5)
        self.wave_fomula_combo.bind("<<ComboboxSelected>>", on_fomula_change)
        on_fomula_change()

    def _on_wave_calc_toggle(self):
        """波浪力复选框切换：激活时允许输入，不激活时 disabled"""
        enabled = self.wave_calc_var.get()
        state = "normal" if enabled else "disabled"
        # create_label_entry 直接返回 ttk.Entry
        if hasattr(self, 'wave_height_entry'):
            self.wave_height_entry.configure(state=state)
        if hasattr(self, 'wave_period_entry'):
            self.wave_period_entry.configure(state=state)
        if hasattr(self, 'wave_length_entry'):
            self.wave_length_entry.configure(state=state)
        self.wave_fomula_combo.configure(state="readonly" if enabled else "disabled")


# ======================== 车辆参数 Tab 页 ========================
class VehicleTab(ttk.Frame):
    """车辆/设备参数tab: 上表(车辆定义) + 下面板(详细参数)"""
    def __init__(self, parent, parent_ui=None):
        super().__init__(parent)
        self.parent_ui = parent_ui

        # --- 从 vehicle_parameter.xlsx 读取数据 ---
        self.VEHICLE_NAMES, self.vehicle_info = self.load_vehicle_xlsx()

        # --- 追加自定义项 ---
        for cust, lt in [("自定义-吊装设备", "吊装设备"), ("自定义-钻孔设备", "钻孔设备"), ("自定义-一般车辆", "一般车辆")]:
            self.VEHICLE_NAMES.append(cust)
            self.vehicle_info[cust] = {"load_type": lt, "params": [], "params": []}

        # 容错
        if not self.VEHICLE_NAMES:
            self.VEHICLE_NAMES = ["请检查数据文件"]
            self.vehicle_info = {"请检查数据文件": {"load_type": "未知", "params": []}}

        self.rows_data = []
        self.current_row = -1
        self.label_width = 18

        # 车辆/设备信息缓存 {vehicle_name: (load_type, [params])}
        self.vehicle_cache = {}

        # ========================================================
        # 快捷设置 + 表格区 + 详细参数区
        # ========================================================
        self.grid_rowconfigure(0, weight=0)                       # 快捷设置（固定高度）
        self.grid_rowconfigure(1, weight=1, uniform="split_50")   # 表格区 50%
        self.grid_rowconfigure(2, weight=1, uniform="split_50")   # 详细参数区 50%
        self.grid_columnconfigure(0, weight=1)

        # ==========================================
        # 0. 快捷设置区
        # ==========================================
        _TRAFFIC_LABELS = ["手动配置"] + [n for n in self.VEHICLE_NAMES if not n.startswith("自定义")]
        _STATIC_LABELS = ["手动配置"] + [n for n in self.VEHICLE_NAMES
                                         if not n.startswith("自定义")
                                         and self.vehicle_info.get(n, {}).get("load_type", "") in ("吊装设备", "钻孔设备")]
        _qd = trestle_excel_dict.get("quick_defaults", {})
        lf_quick = ttk.LabelFrame(self, text="车辆荷载快捷设置", bootstyle=PRIMARY, padding=8)
        lf_quick.grid(row=0, column=0, sticky="ew", padx=15, pady=5)
        _qf = ttk.Frame(lf_quick)
        _qf.pack(fill="x")
        ttk.Label(_qf, text="通行荷载:", bootstyle=PRIMARY).pack(side="left", padx=(5, 2))
        self.quick_traffic_var = tk.StringVar(value=_qd.get("traffic", "手动配置"))
        self.quick_traffic_combo = ttk.Combobox(_qf, textvariable=self.quick_traffic_var,
                                                values=_TRAFFIC_LABELS, state="readonly", width=25)
        self.quick_traffic_combo.pack(side="left", padx=(0, 15))
        ttk.Label(_qf, text="施工静载:", bootstyle=PRIMARY).pack(side="left", padx=(5, 2))
        self.quick_static_var = tk.StringVar(value=_qd.get("static", "手动配置"))
        self.quick_static_combo = ttk.Combobox(_qf, textvariable=self.quick_static_var,
                                               values=_STATIC_LABELS, state="readonly", width=25)
        self.quick_static_combo.pack(side="left", padx=(0, 15))
        self.btn_apply_quick = tb.Button(_qf, text="应用快捷设置", bootstyle=PRIMARY,
                                         command=self._apply_quick_setup)
        self.btn_apply_quick.pack(side="right", padx=10)
        # Combobox 变更时联动按钮状态
        self.quick_traffic_combo.bind("<<ComboboxSelected>>", lambda _: self._update_quick_btn_state())
        self.quick_static_combo.bind("<<ComboboxSelected>>", lambda _: self._update_quick_btn_state())
        self._update_quick_btn_state()

        # ==========================================
        # 1. 表格区
        # ==========================================
        table_container = ttk.Frame(self)
        table_container.grid(row=1, column=0, sticky="nsew", padx=5, pady=(5, 0))
        
        ttk.Label(table_container, text="车辆/设备定义表", font=("Microsoft YaHei UI", 10, "bold"), bootstyle=PRIMARY).pack(anchor=W, padx=10, pady=(5, 5))
        
        btn_bar = ttk.Frame(table_container)
        btn_bar.pack(fill="x", padx=5, pady=10)
        tb.Button(btn_bar, text="+ 添加车辆", bootstyle="info-outline", command=self.add_row).pack(side="left", padx=5)
        tb.Button(btn_bar, text="- 删除选中", bootstyle="danger-outline", command=self.delete_selected).pack(side="left", padx=5)
        tb.Button(btn_bar, text="保存车辆参数", bootstyle="success-outline", command=self.save_vehicle_params).pack(side="left", padx=5)
        ttk.Label(btn_bar, text="*新增自定义车辆及修改标准车辆时需点击按钮保存", font=("微软雅黑", 8), bootstyle="secondary").pack(side="left", padx=5)

        self.table_canvas, self.table_inner = create_scrollable_frame(table_container, pack_kwargs={"padx": 8, "pady": 5})

        headers = ["序号", "车辆/设备", "车辆荷载类型", "车辆/设备名称"]
        weights = [1, 3, 3, 3]  # 各列宽度比例
        for c, t in enumerate(headers):
            tk.Label(self.table_inner, text=t, font=("微软雅黑", 9, "bold"),
                     bg="#e8f4fd" if c < 2 else "#f2f2f2",
                     relief="solid", borderwidth=1, padx=8, pady=5).grid(row=0, column=c, sticky="nsew")
            # 锁定列表比例
            self.table_inner.grid_columnconfigure(c, weight=weights[c], uniform="locked_cols")

        # ==========================================
        # 2. 详细参数区 
        # ==========================================
        self.detail_frame = ttk.LabelFrame(self, text="详细参数", bootstyle=SECONDARY, padding=10)
        self.detail_frame.grid(row=2, column=0, sticky="nsew", padx=5, pady=(5, 5))

        self.refresh_cache()
        # 预填充所有车辆行（从 cache 格式 vehicle dict 读取）；
        # 无引用车辆时保持空表，不自动添加默认行
        veh_cache = trestle_excel_dict.get("vehicle", {})
        for vname, vdata in veh_cache.items():
            self.add_row(prefill_data={
                "name": vdata.get("vehicle_var", vname),
                "load_type": vdata.get("vehicle_type", ""),
                "params": vdata.get("vehicle_params", []),
                "vehicle_name": vname,
                "vehicle_tag": vdata.get("vehicle_tag", ""),
            })

    def reload_from_dict(self, new_dict):
        """数据绑定更新：从 new_dict（cache 格式）刷新车辆表格"""
        # 清空表格数据行（保留 header row 0）
        for w in list(self.table_inner.grid_slaves()):
            g = w.grid_info()
            if g and g.get("row") and int(g["row"]) > 0:
                w.destroy()
        self.rows_data.clear()
        self.current_row = -1
        self.vehicle_cache.clear()

        # 清空详细参数区
        for w in self.detail_frame.winfo_children():
            w.destroy()

        # cache 格式：{vname: {vehicle_var, vehicle_tag, vehicle_type, vehicle_params}}
        veh_cache = new_dict.get("vehicle", {})
        for vname, vdata in veh_cache.items():
            self.add_row(prefill_data={
                "name": vdata.get("vehicle_var", vname),
                "load_type": vdata.get("vehicle_type", ""),
                "params": vdata.get("vehicle_params", []),
                "vehicle_name": vname,
                "vehicle_tag": vdata.get("vehicle_tag", ""),
            })

        # 无引用车辆时保持空表（不自动添加默认行）
        self.refresh_cache()

    def load_vehicle_xlsx(self):
        """从 vehicle_parameter.xlsx 读取车辆/设备数据"""
        names, dct = [], {}
        try:
            fp = os.path.join(self.parent_ui.Applocation, "Support/basic_param", "vehicle_parameter.xlsx")
        except:
            fp = ""
        if not fp or not os.path.exists(fp):
            print(f"警告: 找不到 vehicle_parameter.xlsx")
            return names, dct

        wb = openpyxl.load_workbook(fp, data_only=True)
        COL_B, COL_C = 1, 2  # B=车辆/设备, C=车辆荷载类型

        for ws in wb.worksheets:
            rows = list(ws.iter_rows(min_row=2, values_only=True))
            for row in rows:
                name = (row[COL_B] or "").strip()
                lt = (row[COL_C] or "").strip()
                if not name or not lt:
                    continue
                names.append(name)

                raw = [str(v).strip() for v in row[3:] if v is not None and str(v).strip() not in ("", "/")]

                if lt in ("一般车辆", "自定义-一般车辆"):
                    # params = [车轮间距/轴距(m), 轴重(kN)CSV, 距前轴(m)CSV]
                    veh_params = [raw[0], raw[1], raw[2]] if len(raw) >= 3 else ["", "", ""]
                else:
                    veh_params = raw

                dct[name] = {"load_type": lt, "params": veh_params}
        return names, dct
    
    def refresh_cache(self):
        """重建车辆/设备缓存 {vehicle_name: {vehicle_var, vehicle_tag, vehicle_type, vehicle_params}}"""
        self.vehicle_cache.clear()
        for row in self.rows_data:
            vname = row.get("vehicle_name", row["name"])
            veh_data = row.get("data", {})
            self.vehicle_cache[vname] = {
                "vehicle_var": vname,
                "vehicle_tag": veh_data.get("vehicle_tag") or self._infer_vehicle_tag(vname, veh_data),
                "vehicle_type": veh_data.get("load_type", ""),
                "vehicle_params": veh_data.get("params", []),
            }

    def _infer_vehicle_tag(self, vname, veh_data):
        """推断车辆标签：标准库有同名 → 用标准标签；否则按类型加前缀"""
        if vname in self.vehicle_info:
            return vname  # 标准库标签即名称
        lt = veh_data.get("load_type", "")
        if "吊装" in lt:
            return "自定义-吊装设备"
        elif "钻孔" in lt:
            return "自定义-钻孔设备"
        else:
            return "自定义-一般车辆"

    def save_vehicle_params(self):
        """保存当前选中车辆的详细参数到 vehicle_cache（仅缓存，不写 Excel）"""
        if self.current_row < 0 or self.current_row >= len(self.rows_data):
            messagebox.showwarning("提示", "请先选择一行车辆")
            return

        row = self.rows_data[self.current_row]
        vname = row.get("vehicle_name", "")
        if not vname:
            messagebox.showwarning("提示", "请先输入车辆/设备名称")
            return

        veh_data = row.get("data", {})
        lt = veh_data.get("load_type", "")
        params = veh_data.get("params", [])

        # 查重：标准库
        if vname in self.vehicle_info:
            std_params = self.vehicle_info[vname].get("params", [])
            if params != std_params:
                messagebox.showwarning("名称冲突",
                    f"标准库已存在同名车辆 [{vname}]，参数不同，请重命名")
                return

        # 查重：vehicle_cache
        if vname in self.vehicle_cache:
            cached = self.vehicle_cache[vname]
            cached_params = cached.get("vehicle_params", [])
            if params != cached_params:
                messagebox.showwarning("名称冲突",
                    f"已存在同名车辆 [{vname}]，参数不同，请重命名")
                return

        # 生成 vehicle_tag
        if vname in self.vehicle_info:
            vehicle_tag = vname  # 标准库标签
        elif "吊装" in lt:
            vehicle_tag = "自定义-吊装设备"
        elif "钻孔" in lt:
            vehicle_tag = "自定义-钻孔设备"
        else:
            vehicle_tag = "自定义-一般车辆"

        # 写入 cache
        self.vehicle_cache[vname] = {
            "vehicle_var": vname,
            "vehicle_tag": vehicle_tag,
            "vehicle_type": lt,
            "vehicle_params": params,
        }

        # 同步 rows_data 中的 vehicle_tag
        row["data"]["vehicle_tag"] = vehicle_tag

        # 刷新第2列"车辆/设备"显示为新的标签
        r = self.current_row + 1
        for w in self.table_inner.grid_slaves(row=r, column=1):
            if isinstance(w, tk.Frame):
                for sub in w.winfo_children():
                    if isinstance(sub, ttk.Combobox):
                        sub.set(vehicle_tag)
                        break

        messagebox.showinfo("保存成功", f"车辆 [{vname}] 参数已保存")

    def _update_quick_btn_state(self):
        """快捷设置：两个 Combobox 均非"手动配置"时启用应用按钮"""
        t = self.quick_traffic_var.get()
        s = self.quick_static_var.get()
        if t == "手动配置" and s == "手动配置":
            self.btn_apply_quick.configure(state="normal")
        else:
            self.btn_apply_quick.configure(state="normal")

    def _rebuild_single_vehicle_row(self, ii, row_data):
        """从 row_data 重建单行 UI 控件"""
        r = ii + 1
        dn = row_data["name"]
        lt = row_data.get("load_type", "")
        vehicle_name = row_data.get("vehicle_name", dn)
        # 第2列"车辆/设备"显示标签（tag）：标准车=名称，自定义=“自定义-xxx”，重命名后按类型推断
        veh_data = row_data.get("data", {})
        dn = veh_data.get("vehicle_tag") or self._infer_vehicle_tag(vehicle_name, veh_data) or dn
        lbl_num = tk.Label(self.table_inner, text=str(r), relief="solid", borderwidth=1, padx=5,
                           bg="white", font=("微软雅黑", 9))
        lbl_num.grid(row=r, column=0, sticky="nsew")
        type_frame = tk.Frame(self.table_inner, relief="solid", borderwidth=1, bg="white")
        type_frame.grid(row=r, column=1, sticky="nsew")
        combo_var = tk.StringVar(value=dn)
        cb = ttk.Combobox(type_frame, textvariable=combo_var, values=self.VEHICLE_NAMES, state="readonly")
        cb.pack(padx=2, pady=4, fill="both", expand=True)
        type_label = tk.Label(self.table_inner, text=lt, relief="solid", borderwidth=1, padx=5,
                              bg="#fcfcfc", fg="#666", font=("微软雅黑", 9))
        type_label.grid(row=r, column=2, sticky="nsew")
        name_label = tk.Label(self.table_inner, text=vehicle_name, relief="solid", borderwidth=1, padx=5,
                              bg="white", font=("微软雅黑", 9))
        name_label.grid(row=r, column=3, sticky="nsew")
        name_label.name = "vehicle_name"
        for cell in [lbl_num, type_frame, type_label, name_label]:
            cell.bind("<Button-1>", lambda e, idx=ii: self.select(idx), add="+")
            for sub in cell.winfo_children():
                sub.bind("<Button-1>", lambda e, idx=ii: self.select(idx), add="+")
        cb.bind("<<ComboboxSelected>>", lambda e, idx=ii, _cv=combo_var: self.on_change(idx, _cv))
        cb.bind("<FocusIn>", lambda e, idx=ii: self.select(idx), add="+")

    def _apply_quick_setup(self):
        """应用快捷设置：确定目标车辆集 → 同步车辆表 → 配下游 tab"""
        traffic_name = self.quick_traffic_var.get()
        static_name = self.quick_static_var.get()
        has_traffic = (traffic_name != "手动配置")
        has_static = (static_name != "手动配置")

        print(f"\n[quick_setup] traffic={traffic_name} static={static_name}")

        # ── 1. 计算目标车辆集 ──
        if has_traffic and has_static:
            target = list(dict.fromkeys([traffic_name, static_name]))
        elif has_traffic:
            # 通行车辆 + 静载预设车辆（从 cache 格式 static_load_cases 读取）
            target = [traffic_name]
            for _entry in trestle_excel_dict.get("static_load_cases", {}).values():
                v = _entry.get("vehicle", "")
                if v and v not in target:
                    target.append(v)
        elif has_static:
            # 移动荷载引用的车辆 + 新施工车辆
            target = []
            if self.parent_ui and hasattr(self.parent_ui, 'tab_moveloads'):
                for cd in self.parent_ui.tab_moveloads.moveload_cache.values():
                    v = cd.get("vehicle", "")
                    if v and v not in target:
                        target.append(v)
            if static_name not in target:
                target.append(static_name)
        else:
            target = None

        print(f"[quick_setup] target vehicles: {target}")
        if target:
            veh_defaults = []
            for vname in target:
                if vname in self.vehicle_info:
                    vi = self.vehicle_info[vname]
                    veh_defaults.append({"name": vname, "load_type": vi["load_type"], "params": vi["params"]})
                else:
                    veh_defaults.append({"name": vname, "load_type": "", "params": []})
            print(f"[quick_setup] vehicle_defaults: {veh_defaults}")

        # ── 2. 同步 VehicleTab ──
        if target is not None:
            target_set = set(target)
            # 从 rows_data 移除不在目标集中的车辆
            self.rows_data = [r for r in self.rows_data
                              if r.get("vehicle_name", r["name"]) in target_set]
            # 清空 UI 并从 rows_data 重建
            for w in list(self.table_inner.grid_slaves()):
                g = w.grid_info()
                if g and g.get("row") and int(g["row"]) > 0:
                    w.destroy()
            for ii, row_data in enumerate(self.rows_data):
                self._rebuild_single_vehicle_row(ii, row_data)
            # 追加缺失的车辆
            existing = {r.get("vehicle_name", r["name"]) for r in self.rows_data}
            for vname in target:
                if vname not in existing:
                    if vname in self.vehicle_info:
                        vi = self.vehicle_info[vname]
                        self.add_row(prefill_data={"name": vname, "load_type": vi["load_type"], "params": vi["params"]})
                    else:
                        self.add_row(prefill_data={"name": vname, "load_type": "", "params": []})
            self.table_canvas.configure(scrollregion=self.table_canvas.bbox("all"))
            self.refresh_cache()
            if self.rows_data:
                self.select(0, force=True)

        # ── 3. 配置移动荷载 tab（仅通行选择时覆盖）──
        if has_traffic:
            case_name = f"{traffic_name}走行"
            print(f"[quick_setup] move_load_cases_defaults: [{{name: {case_name}, vehicle: {traffic_name}, lane_ecc: 0}}]")
        if has_traffic and self.parent_ui and hasattr(self.parent_ui, 'tab_moveloads'):
            mlt = self.parent_ui.tab_moveloads
            for w in list(mlt.table_inner.grid_slaves()):
                g = w.grid_info()
                if g and g.get("row") and int(g["row"]) > 0:
                    w.destroy()
            mlt.rows_data.clear()
            case_name = mlt._gen_case_name(traffic_name, "0")
            mlt.rows_data.append({
                "name": case_name, "vehicle": traffic_name,
                "lane_ecc": "0", "direction": "往返", "factor": "1.0",
            })
            vnames = mlt.get_vehicle_names()
            mlt._insert_row_ui(0, 1, traffic_name, "0", case_name, vnames)
            mlt.table_canvas.configure(scrollregion=mlt.table_canvas.bbox("all"))
            mlt.refresh_moveload_cache()
            mlt.select(0, force=True)

        # ── 4. 配置静载 tab（仅施工选择时覆盖）──
        if has_static:
            print(f"[quick_setup] static_load: vehicle={static_name}, types=[正吊(正向), 侧吊], span_mid=all, pile=all")
        if has_static and self.parent_ui and hasattr(self.parent_ui, 'tab_staticloads'):
            slt = self.parent_ui.tab_staticloads
            slt._on_page_shown()
            for vname in slt.car_dict:
                slt.car_output_check[vname] = (vname == static_name)
            if static_name in slt.car_dict:
                data = slt.car_dict[static_name]
                for lt in data.get("load_types", {}):
                    data["load_types"][lt] = (lt in ("正吊(正向)", "侧吊"))
                for row in data.get("span_checks", []):
                    if len(row) >= 3:
                        row[0] = False
                        row[1] = True
                        row[2] = False
                for j in range(len(data.get("pile_checks", []))):
                    data["pile_checks"][j] = True
                slt._rebuild_loading_tables()
                slt._current_car_name = static_name
                for lt, var in slt.load_type_vars.items():
                    var.set(lt in ("正吊(正向)", "侧吊"))
                for row_vars in slt._span_chk_vars:
                    if len(row_vars) >= 3:
                        row_vars[0].set(False)
                        row_vars[1].set(True)
                        row_vars[2].set(False)
                for var in slt._pile_chk_vars:
                    var.set(True)
                slt._save_current_state()
            slt.refresh_cache()
            draw_static_load_diagram(slt)

        # ── 5. 荷载组合 ──
        if (has_traffic or has_static) and self.parent_ui and hasattr(self.parent_ui, 'tab_loadcombs'):
            self.parent_ui.tab_loadcombs._auto_combine()

    def add_row(self, prefill_data=None):
        i = len(self.rows_data)
        if prefill_data:
            dn = prefill_data["name"]
            veh_data = {
                "load_type": prefill_data.get("load_type", ""),
                "params": prefill_data.get("params", []),
                "vehicle_tag": prefill_data.get("vehicle_tag", ""),
            }
        else:
            dn = self.VEHICLE_NAMES[0]
            veh_data = self.vehicle_info.get(dn, {"load_type": "", "params": []})
        lt = veh_data.get("load_type", "")

        # 自定义车辆：名称留空
        if dn.startswith("自定义-"):
            vehicle_name = prefill_data.get("vehicle_name", "") if prefill_data else ""
        else:
            vehicle_name = prefill_data.get("vehicle_name", dn) if prefill_data else dn

        self.rows_data.append({"name": dn, "load_type": lt, "data": veh_data, "vehicle_name": vehicle_name})

        # 统一由行构建函数创建该行控件（与 _apply_quick_setup / delete_selected 一致）
        self._rebuild_single_vehicle_row(i, self.rows_data[i])

        self.table_canvas.configure(scrollregion=self.table_canvas.bbox("all"))
        self.refresh_cache()
        self.select(i, force=True)

    def on_change(self, idx, cv):
        nm = cv.get()
        veh_data = self.vehicle_info.get(nm, {"load_type": "", "params": []})
        lt = veh_data.get("load_type", "")

        old = self.rows_data[idx]
        old_vehicle_name = old.get("vehicle_name", "")
        old_vehicle_name_internal = old["name"]

        # 自定义车辆：名称留空，参数全空
        if nm.startswith("自定义-"):
            new_vehicle_name = ""
            new_data = {"load_type": lt, "params": [], "vehicle_tag": nm}
        else:
            # 若用户未手动修改过名（与下拉值相同），则跟随新选择；否则保留自定义名
            if not old_vehicle_name or old_vehicle_name == old_vehicle_name_internal:
                new_vehicle_name = nm
            else:
                new_vehicle_name = old_vehicle_name
            # 保留用户已编辑的 params：同名切换不覆盖，异名切换用库默认值（深拷贝，避免污染标准库）
            if nm == old["name"]:
                new_data = old["data"]  # 同一辆车，保留所有编辑
            else:
                new_data = {"load_type": lt, "params": list(veh_data.get("params", []))}  # 深拷贝 params

        self.rows_data[idx] = {"name": nm, "load_type": lt, "data": new_data, "vehicle_name": new_vehicle_name}
        
        for w in self.table_inner.grid_slaves(row=idx+1, column=2):
            if isinstance(w, tk.Label):
                w.config(text=lt)
                break
        # 同步车辆/设备名称列显示
        for w in self.table_inner.grid_slaves(row=idx+1, column=3):
            if isinstance(w, tk.Label):
                w.config(text=new_vehicle_name)
                break

        # 仅在更改下拉框时强制刷新
        self.refresh_cache()
        self.select(idx, force=True)

    def delete_selected(self):
        if not self.rows_data: return
        idx = self.current_row
        if idx >= len(self.rows_data): return

        deleted_row = self.rows_data[idx]
        deleted_names = [deleted_row.get("vehicle_name", deleted_row["name"]), deleted_row["name"]]
        self.rows_data.pop(idx)

        # 完全销毁所有行控件，再从数据重建
        for w in list(self.table_inner.grid_slaves()):
            g = w.grid_info()
            if g and g.get("row") and int(g["row"]) > 0:
                w.destroy()

        for ii, row_data in enumerate(self.rows_data):
            self._rebuild_single_vehicle_row(ii, row_data)

        self.refresh_cache()
        self.table_canvas.configure(scrollregion=self.table_canvas.bbox("all"))

        new_idx = min(idx, len(self.rows_data) - 1) if self.rows_data else -1
        if new_idx >= 0:
            self.select(new_idx, force=True)
        else:
            for w in self.detail_frame.winfo_children(): w.destroy()

        # 联动：同步清理移动荷载、静载、荷载组合中涉及该车辆的内容
        if self.parent_ui:
            if hasattr(self.parent_ui, "tab_moveloads"):
                self.parent_ui.tab_moveloads.on_vehicle_deleted(deleted_names)
            if hasattr(self.parent_ui, "tab_staticloads"):
                self.parent_ui.tab_staticloads.on_vehicle_deleted(deleted_names)
            if hasattr(self.parent_ui, "tab_loadcombs"):
                self.parent_ui.tab_loadcombs.on_vehicle_deleted(deleted_names)
    def select(self, idx, force=False):
        if not self.rows_data or idx < 0 or idx >= len(self.rows_data): return
        
        need_refresh = force or getattr(self, "current_row", -1) != idx
        self.current_row = idx

        # 高亮颜色刷新
        for r in range(1, len(self.rows_data)+1):
            bg = "#e3f2fd" if r-1 == idx else "white"
            for w in self.table_inner.grid_slaves(row=r):
                try: w.config(bg=bg) 
                except: pass
                
                if isinstance(w, tk.Frame):
                    for sub in w.winfo_children():
                        if isinstance(sub, tk.Label):
                            sub.config(bg=bg) 
                            txt = sub.cget("text")
                            if txt in ["是", "否"]:
                                sub.config(fg="#2b6fd4" if txt == "是" else "#95a5a6")
                                
        if need_refresh:
            self.show_detail(idx)

    def show_detail(self, idx):
        # 1. 清空详情面板
        for w in self.detail_frame.winfo_children(): w.destroy()
        
        if not self.rows_data or idx >= len(self.rows_data): return
        row = self.rows_data[idx]
        nm = row["name"]
        veh_data = row["data"]
        lt = veh_data.get("load_type", "")
        params = veh_data.get("params", [])

        # ========================================================
        # 为包含表格的车辆动态嵌套滚动画布
        # ========================================================
        if "一般车辆" in lt:
            v_canvas = tk.Canvas(self.detail_frame, borderwidth=0, highlightthickness=0, height=180)
            v_scroll = ttk.Scrollbar(self.detail_frame, orient="vertical", command=v_canvas.yview)
            v_inner = ttk.Frame(v_canvas)

            v_canvas.pack(side="left", fill="both", expand=True)
            v_scroll.pack(side="right", fill="y")
            v_canvas.configure(yscrollcommand=v_scroll.set)

            inner_id = v_canvas.create_window((0, 0), window=v_inner, anchor="nw")

            v_inner.bind("<Configure>", lambda e: v_canvas.configure(scrollregion=v_canvas.bbox("all")))
            v_canvas.bind("<Configure>", lambda e: v_canvas.itemconfig(inner_id, width=e.width))

            def on_v_mousewheel(event):
                v_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            v_canvas.bind("<Enter>", lambda e: v_canvas.bind_all("<MouseWheel>", on_v_mousewheel))
            v_canvas.bind("<Leave>", lambda e: v_canvas.unbind_all("<MouseWheel>"))

            target_parent = v_inner
        else:
            target_parent = self.detail_frame

        # ========================================================
        # 核心布局：建立左面板、中轴线、右面板
        # ========================================================
        left_panel = ttk.Frame(target_parent)
        left_panel.pack(side="left", fill="y", padx=(10, 20), pady=10)

        # 垂直灰色分隔线
        sep = ttk.Separator(target_parent, orient="vertical")
        sep.pack(side="left", fill="y", padx=10, pady=10)

        right_panel = ttk.Frame(target_parent)
        right_panel.pack(side="left", fill="both", expand=True, padx=(20, 10), pady=10)

        def bind_sync(widget, p_idx):
            def save(e):
                while len(params) <= p_idx: params.append("")
                params[p_idx] = widget.get()
                row["data"]["params"] = params
                # 修改后立即更新自身缓存 + 通知 MoveLoadTab 重绘
                self.refresh_cache()
                if self.parent_ui and hasattr(self.parent_ui, 'tab_moveloads'):
                    mlt = self.parent_ui.tab_moveloads
                    if mlt.current_row >= 0:
                        mlt.redraw_diagram(mlt.current_row)
            widget.bind("<FocusOut>", save, add="+")
            widget.bind("<Return>", save, add="+")

        # ========== 1. 左侧通用参数渲染 (所有车辆都有) ==========
        stored_name = row.get("vehicle_name", nm)
        vehicle_name = create_label_entry(left_panel, "车辆/设备名称:", 0, default_val=stored_name, label_width=16, entry_width=20, padx=5, pady=12)
        def sync_name(event=None):
            old_name = self.rows_data[idx].get("name", "")
            new_vehicle_name = vehicle_name.get()
            new_name = new_vehicle_name
            # 更新界面 Label
            for w in self.table_inner.grid_slaves(row=idx+1, column=3):
                if isinstance(w, tk.Label):
                    w.config(text=new_vehicle_name)
                    break
            self.rows_data[idx]["vehicle_name"] = new_vehicle_name
            self.rows_data[idx]["name"] = new_name
            self.refresh_cache()
            # 级联通知所有依赖 tab
            if self.parent_ui:
                if hasattr(self.parent_ui, "tab_moveloads"):
                    self.parent_ui.tab_moveloads.on_vehicle_renamed(old_name, new_name)
                if hasattr(self.parent_ui, "tab_staticloads"):
                    self.parent_ui.tab_staticloads.on_vehicle_renamed(old_name, new_name)
                if hasattr(self.parent_ui, "tab_loadcombs"):
                    self.parent_ui.tab_loadcombs.on_vehicle_renamed(old_name, new_name)
        vehicle_name.bind("<FocusOut>", sync_name, add="+")
        vehicle_name.bind("<Return>", sync_name, add="+")

        vehicle_dd = create_label_entry(left_panel, "轴距(m):", 1, default_val=params[0] if len(params)>0 else "", label_width=16, entry_width=20, padx=5, pady=12)
        bind_sync(vehicle_dd, 0) # 绑定到 params[0]
        
        # ========== 2. 右侧详细参数渲染 (根据类型动态生成) ==========
        # --- 公路标准荷载 ---
        if lt == "公路标准荷载":
            ttk.Label(right_panel, text="标准荷载类型:", width=16, bootstyle=PRIMARY).grid(row=0, column=0, sticky="w", padx=5, pady=12)
            st = ttk.Combobox(right_panel, values=["CH-CL","CH-CD"], state="normal", width=18)
            st.grid(row=0, column=1, sticky="w", padx=5)
            st.set(params[1] if len(params)>1 else "")

        # --- 吊装设备 & 自定义吊装设备 ---
        elif lt in ["吊装设备", "自定义-吊装设备"]:
            w1 = create_label_entry(right_panel, "自重(kN):", 0, default_val=params[1] if len(params)>1 else "", label_width=16, entry_width=20, padx=5, pady=12)
            bind_sync(w1, 1)
            w2 = create_label_entry(right_panel, "最大吊重(kN):", 1, default_val=params[2] if len(params)>2 else "", label_width=16, entry_width=20, padx=5, pady=12)
            bind_sync(w2, 2)
            w3 = create_label_entry(right_panel, "吊臂(m):", 2, default_val=params[3] if len(params)>3 else "", label_width=16, entry_width=20, padx=5, pady=12)
            bind_sync(w3, 3)
            w4 = create_label_entry(right_panel, "接地长度(m):", 3, default_val=params[4] if len(params)>4 else "", label_width=16, entry_width=20, padx=5, pady=12)
            bind_sync(w4, 4)
            w5 = create_label_entry(right_panel, "接地宽度(m):", 4, default_val=params[5] if len(params)>5 else "", label_width=16, entry_width=20, padx=5, pady=12)
            bind_sync(w5, 5)

        # --- 钻孔设备 & 自定义钻孔设备 ---
        elif lt in ["钻孔设备", "自定义-钻孔设备"]:
            w1 = create_label_entry(right_panel, "自重(kN):", 0, default_val=params[1] if len(params)>1 else "", label_width=16, entry_width=20, padx=5, pady=12)
            bind_sync(w1, 1)
            w2 = create_label_entry(right_panel, "扭矩(kN·m):", 1, default_val=params[2] if len(params)>2 else "", label_width=16, entry_width=20, padx=5, pady=12)
            bind_sync(w2, 2)
            w3 = create_label_entry(right_panel, "接地长度(m):", 2, default_val=params[3] if len(params)>3 else "", label_width=16, entry_width=20, padx=5, pady=12)
            bind_sync(w3, 3)
            w4 = create_label_entry(right_panel, "接地宽度(m):", 3, default_val=params[4] if len(params)>4 else "", label_width=16, entry_width=20, padx=5, pady=12)
            bind_sync(w4, 4)

        # --- 一般车辆 & 自定义一般车辆 ---
        elif lt in ["一般车辆", "自定义-一般车辆"]:
            # 表头
            hf = ttk.Frame(right_panel)
            hf.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=(0, 10))
            for c, t in enumerate(["编号", "轴重(kN)", "距前轴(m)", "操作"]):
                ttk.Label(hf, text=t, width=10, anchor="center", font=("微软雅黑", 9, "bold")).grid(row=0, column=c, padx=2)
            
            rl = []
            def _ar(ld="", dd=""):
                if len(rl) >= 10: return  # 上限10行
                ri = len(rl) + 1
                fr = ttk.Frame(right_panel)
                fr.grid(row=ri, column=0, columnspan=2, sticky="ew", pady=1, padx=5)
                ttk.Label(fr, text=str(ri), width=8, anchor="center").grid(row=0, column=0, padx=2)
                le = ttk.Entry(fr, width=10, justify="center"); le.insert(0, str(ld)); le.grid(row=0, column=1, padx=5)
                de = ttk.Entry(fr, width=10, justify="center"); de.insert(0, str(dd)); de.grid(row=0, column=2, padx=5)
                tb.Button(fr, text="删除", bootstyle="danger-outline", width=5, command=lambda: (_dr(fr), None)).grid(row=0, column=3, padx=3)
                rl.append({"f": fr, "l": le, "d": de})
                
            def _dr(fr):
                fr.destroy()
                rl[:] = [x for x in rl if x["f"] != fr]
                for i, x in enumerate(rl, 1):
                    for w in x["f"].winfo_children():
                        if isinstance(w, ttk.Label): w.config(text=str(i)); break

            # 从 params 加载已有数据：params = [轴距, 轴重CSV, 距前轴CSV]
            if len(params) >= 3:
                loads = [x.strip() for x in params[1].split(",") if x.strip()]
                dists = [x.strip() for x in params[2].split(",") if x.strip()]
                for ld, dd in zip(loads, dists):
                    _ar(ld, dd)
                
            if len(rl) == 0: _ar() # 如果没有数据，默认加一行空行

            # 新增按钮放在列表下方
            tb.Button(right_panel, text="新增", bootstyle="info-outline", width=10, command=lambda: _ar()).grid(row=999, column=0, columnspan=2, pady=10)
            
            # 车辆荷载示意图
            try:
                import os
                from PIL import Image, ImageTk
                img_path = os.path.join(self.parent_ui.Applocation, "Support", "png_picture", "truck_load.png")
                if os.path.exists(img_path):
                    img = Image.open(img_path)
                    self._photo = ImageTk.PhotoImage(img.resize((450, 160), Image.Resampling.LANCZOS))
                    ttk.Label(right_panel, image=self._photo).grid(row=1000, column=0, columnspan=2, pady=(5,0))
            except Exception as e: 
                print("加载图片失败:", e)
        
        else:
            ttk.Label(right_panel, text=f"暂无当前选择类型的详细参数面板。\n类型: {lt}", font=("微软雅黑",9), foreground="gray").grid(row=0, column=0, pady=20, padx=10, sticky="w")


# ======================== 移动荷载 Tab 页 ========================
class MoveLoadTab(ttk.Frame):
    """移动荷载工况tab: 上表(工况定义) + 下图(加载示意图)"""
    def __init__(self, parent, parent_ui=None):
        super().__init__(parent)
        self.parent_ui = parent_ui

        self.rows_data = []
        self.current_row = -1

        # 移动荷载工况缓存 {case_name: [vehicle, lane_ecc, direction, factor]}
        self.moveload_cache = {}

        # 页面布局
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # ==================== 1. 工况表 ====================
        table_container = ttk.Frame(self)
        table_container.grid(row=0, column=0, sticky="nsew", padx=5, pady=(5, 0))

        ttk.Label(table_container, text="移动荷载工况表", font=("Microsoft YaHei UI", 10, "bold"),
                  bootstyle=PRIMARY).pack(anchor=W, padx=10, pady=(5, 5))

        btn_bar = ttk.Frame(table_container)
        btn_bar.pack(fill="x", padx=5, pady=10)
        tb.Button(btn_bar, text="+ 添加工况", bootstyle="info-outline", command=self.add_row).pack(side="left", padx=5)
        tb.Button(btn_bar, text="- 删除选中", bootstyle="danger-outline", command=self.delete_selected).pack(side="left", padx=5)

        self.table_canvas, self.table_inner = create_scrollable_frame(table_container, pack_kwargs={"padx": 8, "pady": 5})

        headers = ["序号", "工况名称", "车辆/设备", "车道偏心(m)"]
        weights = [1, 3, 3, 2]
        for c, (t, w) in enumerate(zip(headers, weights)):
            tk.Label(self.table_inner, text=t, font=("微软雅黑", 9, "bold"),
                     bg="#e8f4fd" if c < 2 else "#f2f2f2",
                     relief="solid", borderwidth=1, padx=8, pady=5).grid(row=0, column=c, sticky="nsew")
            self.table_inner.grid_columnconfigure(c, weight=w, uniform="mv_cols")

        # ==================== 2. 示意图区 ====================
        diagram_frame = ttk.Frame(self)
        diagram_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=(5, 5))
        # 上方分隔线
        top_sep = ttk.Separator(diagram_frame, orient=HORIZONTAL)
        top_sep.pack(fill="x", padx=10, pady=(5, 0))
        # 标题 + 标题下方的分隔线
        self.title_lbl = tk.Label(diagram_frame, text="", font=("Microsoft YaHei UI", 9, "bold"), bg="white", fg="#000000")
        self.title_lbl.pack(anchor="center", pady=(5, 3))
        sep_line = ttk.Separator(diagram_frame, orient=HORIZONTAL)
        sep_line.pack(fill="x", padx=20, pady=(0, 0))

        self.fig = Figure(figsize=(6.5, 3), dpi=100, facecolor="white")
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=diagram_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        # 预填充所有移动荷载工况行（cache 格式 move_load_cases dict）
        ml_cache = trestle_excel_dict.get("move_load_cases", {})
        for case_name, case_data in ml_cache.items():
            self.add_row({**case_data, "name": case_name})

    def reload_from_dict(self, new_dict):
        """数据绑定更新：从 new_dict（cache 格式）刷新移动荷载"""
        # 清空表格数据行（保留 header row 0）
        for w in list(self.table_inner.grid_slaves()):
            g = w.grid_info()
            if g and g.get("row") and int(g["row"]) > 0:
                w.destroy()
        self.rows_data.clear()
        self.current_row = -1
        self.moveload_cache.clear()

        # cache 格式：{case_name: {vehicle, lane_ecc, direction, factor}}
        # 空表时不预填默认行
        ml_cache = new_dict.get("move_load_cases", {})
        for case_name, case_data in ml_cache.items():
            self.add_row({**case_data, "name": case_name})

    def get_vehicle_names(self):
        """从 VehicleTab.vehicle_cache 获取车辆/设备名称列表"""
        if self.parent_ui and hasattr(self.parent_ui, 'tab_vehicles'):
            return list(self.parent_ui.tab_vehicles.vehicle_cache.keys())
        return []

    def _insert_row_ui(self, i, r, veh, ecc, name, vnames):
        """创建一行 UI 控件（不操作 rows_data）"""
        # ---- 序号 ----
        lbl_num = tk.Label(self.table_inner, text=str(r), relief="solid", borderwidth=1, padx=5,
                           bg="white", font=("微软雅黑", 9))
        lbl_num.grid(row=r, column=0, sticky="nsew")

        # ---- 工况名称（只读，自动生成）----
        name_lbl = tk.Label(self.table_inner, text=name, relief="solid", borderwidth=1, padx=5,
                           bg="#f0f0f0", font=("微软雅黑", 9), fg="#333")
        name_lbl.grid(row=r, column=1, sticky="nsew")

        # ---- 车辆/设备 ----
        veh_frame = tk.Frame(self.table_inner, relief="solid", borderwidth=1, bg="white")
        veh_frame.grid(row=r, column=2, sticky="nsew")
        combo_var = tk.StringVar(value=veh)
        cb = ttk.Combobox(veh_frame, textvariable=combo_var, values=vnames, state="readonly")
        cb.pack(padx=2, pady=4, fill="both", expand=True)
        cb.is_veh_cb = True

        # ---- 车道偏心 ----
        ecc_entry = tk.Entry(self.table_inner, relief="solid", borderwidth=1, bg="white", font=("微软雅黑", 9), justify="center")
        ecc_entry.grid(row=r, column=3, sticky="nsew")
        ecc_entry.insert(0, ecc)

        # 绑定点击与高亮事件
        cells = [lbl_num, name_lbl, veh_frame, ecc_entry]
        for cell in cells:
            cell.bind("<Button-1>", lambda e, idx=i: self.select(idx), add="+")
            for sub in cell.winfo_children():
                sub.bind("<Button-1>", lambda e, idx=i: self.select(idx), add="+")
                if isinstance(sub, ttk.Combobox):
                    sub.bind("<FocusIn>", lambda e, idx=i: self.select(idx), add="+")
                elif isinstance(sub, tk.Entry):
                    sub.bind("<FocusIn>", lambda e, idx=i: self.select(idx), add="+")

        cb.bind("<<ComboboxSelected>>", lambda e, idx=i, _cv=combo_var: self.on_veh_change(idx, _cv))
        ecc_entry.bind("<KeyRelease>", lambda e, idx=i: self.on_param_change(idx))
    def add_row(self, row_data=None):
        """添加一行移动荷载工况。

        row_data 提供时（reload 场景，cache 格式）按给定数据填行；
        否则按 trestle_excel_dict 的 move_load_cases 顺序预填。
        """
        i = len(self.rows_data)
        r = i + 1
        vnames = self.get_vehicle_names()
        default_veh = vnames[0] if vnames else "未定义"
        if row_data:
            veh = row_data.get("vehicle", default_veh)
            ecc = str(row_data.get("lane_ecc", "0.0"))
            name = row_data.get("name", self._gen_case_name(veh, ecc))
            direction = row_data.get("direction", "往返")
            factor = row_data.get("factor", "1.0")
        else:
            _ml_cache = list(trestle_excel_dict.get("move_load_cases", {}).values())
            _ml_idx = i if i < len(_ml_cache) else None
            veh = _ml_cache[i].get("vehicle", default_veh) if _ml_idx is not None else default_veh
            ecc = str(_ml_cache[i].get("lane_ecc", "0.0")) if _ml_idx is not None else "0.0"
            name = self._gen_case_name(veh, ecc)
            direction = "往返"
            factor = "1.0"
        self.rows_data.append({
            "name": name,
            "vehicle": veh,
            "lane_ecc": ecc,
            "direction": direction,
            "factor": factor,
        })
        self._insert_row_ui(i, r, veh, ecc, name, vnames)
        self.table_canvas.configure(scrollregion=self.table_canvas.bbox("all"))
        self.refresh_moveload_cache()
        self.select(i, force=True)
    def _gen_case_name(self, vehicle, lane_ecc):
        """根据车辆名称和车道偏心自动生成工况名称"""
        try: ecc_f = float(lane_ecc)
        except: ecc_f = 0.0
        if abs(ecc_f) > 0.001:
            return f"{vehicle}{ecc_f}m偏心走行"
        else:
            return f"{vehicle}走行"

    def on_vehicle_renamed(self, old_name, new_name):
        """VehicleTab 车辆改名后的联动：更新 vehicle 字段 + 重新生成工况名 + 刷新 UI"""
        renamed = []
        for row in self.rows_data:
            if row.get("vehicle", "") == old_name:
                row["vehicle"] = new_name
                # 重新生成工况名（使用新名 + 原有偏心）
                ecc = row.get("lane_ecc", "0.0")
                row["name"] = self._gen_case_name(new_name, ecc)
                renamed.append(True)
            else:
                renamed.append(False)
        if not any(renamed):
            return
        # 完全重绘所有行
        for w in list(self.table_inner.grid_slaves()):
            g = w.grid_info()
            if g and g.get("row") and int(g["row"]) > 0:
                w.destroy()
        vnames = self.get_vehicle_names()
        for ii, row_data in enumerate(self.rows_data):
            self._insert_row_ui(ii, ii + 1,
                row_data.get("vehicle", ""),
                row_data.get("lane_ecc", "0.0"),
                row_data.get("name", ""), vnames)
        self.refresh_moveload_cache()
        self.table_canvas.configure(scrollregion=self.table_canvas.bbox("all"))
        if self.rows_data:
            new_idx = min(getattr(self, "current_row", 0), len(self.rows_data) - 1)
            self.select(new_idx, force=True)

    def on_veh_change(self, idx, cv):
        nm = cv.get()
        self.rows_data[idx]["vehicle"] = nm
        ecc = self.rows_data[idx].get("lane_ecc", "0.0")
        new_name = self._gen_case_name(nm, ecc)
        self.rows_data[idx]["name"] = new_name
        # 更新界面上该行的名称标签
        for w in self.table_inner.grid_slaves(row=idx+1):
            if isinstance(w, tk.Label) and w.grid_info().get("column") == 1:
                w.config(text=new_name)
                break
        self.refresh_moveload_cache()
        self.redraw_diagram(idx)

    def on_param_change(self, idx):
        if idx < 0 or idx >= len(self.rows_data): return
        # 从界面控件读取当前值
        ecc_entry = None
        name_lbl = None
        for w in self.table_inner.grid_slaves(row=idx+1):
            if isinstance(w, tk.Entry) and w.grid_info().get("column") == 3:
                ecc_entry = w
            elif isinstance(w, tk.Label) and w.grid_info().get("column") == 1:
                name_lbl = w
        if ecc_entry:
            ecc = ecc_entry.get().strip()
            self.rows_data[idx]["lane_ecc"] = ecc
            veh = self.rows_data[idx].get("vehicle", "")
            new_name = self._gen_case_name(veh, ecc)
            self.rows_data[idx]["name"] = new_name
            if name_lbl:
                name_lbl.config(text=new_name)
        self.refresh_moveload_cache()
        self.redraw_diagram(idx)

    def on_dir_change(self, idx, dv):
        self.rows_data[idx]["direction"] = dv.get()
        self.refresh_moveload_cache()
        self.select(idx, force=True)

    def delete_selected(self):
        if not self.rows_data: return
        idx = self.current_row
        if idx >= len(self.rows_data): return

        self.rows_data.pop(idx)

        # 完全销毁所有行控件，再从数据重建
        for w in list(self.table_inner.grid_slaves()):
            g = w.grid_info()
            if g and g.get("row") and int(g["row"]) > 0:
                w.destroy()

        vnames = self.get_vehicle_names()
        for ii, row_data in enumerate(self.rows_data):
            self._insert_row_ui(ii, ii + 1,
                row_data.get("vehicle", ""),
                row_data.get("lane_ecc", "0.0"),
                row_data.get("name", ""), vnames)

        self.refresh_moveload_cache()
        self.table_canvas.configure(scrollregion=self.table_canvas.bbox("all"))

        new_idx = min(idx, len(self.rows_data) - 1) if self.rows_data else -1
        if new_idx >= 0:
            self.select(new_idx, force=True)
    def select(self, idx, force=False):
        if not self.rows_data or idx < 0 or idx >= len(self.rows_data): return

        need_redraw = force or getattr(self, "current_row", -1) != idx
        self.current_row = idx

        for r in range(1, len(self.rows_data)+1):
            bg = "#e3f2fd" if r-1 == idx else "white"
            for w in self.table_inner.grid_slaves(row=r):
                try: 
                    w.config(bg=bg)
                except: 
                    pass
                if isinstance(w, tk.Frame):
                    for sub in w.winfo_children():
                        if isinstance(sub, (ttk.Label, tk.Label)):
                            try: 
                                sub.config(bg=bg)
                            except: 
                                pass
        if need_redraw:
            self.redraw_diagram(idx)

    def get_vehicle_params(self, vehicle_name):
        """从 VehicleTab.vehicle_cache 获取车辆参数 (vehicle_dd, dd_w)"""
        cache = {}
        if self.parent_ui and hasattr(self.parent_ui, 'tab_vehicles'):
            cache = self.parent_ui.tab_vehicles.vehicle_cache
        entry = cache.get(vehicle_name)
        if entry is None:
            return None, None
        lt = entry.get("vehicle_type", "")
        params = entry.get("vehicle_params", [])
        vehicle_dd = params[0] if len(params) > 0 else ""
        dd_w = (params[5] if len(params) > 5 else "0.5") if lt in ["吊装设备", "自定义-吊装设备"] else (
            params[4] if len(params) > 4 else "0.5") if lt in ["钻孔设备", "自定义-钻孔设备"] else "0.5"
        return vehicle_dd, dd_w

    def redraw_diagram(self, idx):
        self.ax.clear()
        row = self.rows_data[idx]
        case_name = row.get("name", "")
        self.title_lbl.config(text=f"{case_name}-加载示意图" if case_name else "移动荷载加载示意图")
        if idx < 0 or not self.rows_data or idx >= len(self.rows_data):
            self.ax.text(0.5, 0.5, "请选择工况", ha="center", va="center", transform=self.ax.transAxes,
                        fontsize=12, color="gray")
            self.canvas.draw()
            return

        row = self.rows_data[idx]
        veh_name = row["vehicle"]
        lane_ecc = row["lane_ecc"]

        # 读 vehicle_cache
        entry = {}
        if self.parent_ui and hasattr(self.parent_ui, 'tab_vehicles'):
            entry = self.parent_ui.tab_vehicles.vehicle_cache.get(veh_name, {})
        if entry:
            lt = entry.get("vehicle_type", "一般车辆")
            params = entry.get("vehicle_params", [])
        else:
            lt, params = "一般车辆", []

        # 轴距 / 接地宽度
        vehicle_dd, dd_w = self.get_vehicle_params(veh_name)
        try: wd = float(vehicle_dd)
        except: wd = 1.8
        try: tw = float(dd_w)
        except: tw = 0.5

        lane_p = [lane_ecc, wd, tw]
        draw_move_load_diagram(self, self.ax, lt, params, lane_p, hl_key="lane_ecc")
        self.canvas.draw()

    def _on_page_shown(self):
        """切换到本页时刷新下拉框"""
        self.refresh_combobox_values()

    def refresh_combobox_values(self):
        """仅刷新'车辆/设备'列的下拉框"""
        vnames = self.get_vehicle_names()
        for r in range(1, len(self.rows_data)+1):
            for w in self.table_inner.grid_slaves(row=r):
                if isinstance(w, tk.Frame):
                    for sub in w.winfo_children():
                        # 【关键修复】只刷新带有 is_veh_cb=True 标记的下拉框
                        if isinstance(sub, ttk.Combobox) and getattr(sub, 'is_veh_cb', False):
                            sub.configure(values=vnames)
                            # 保持当前选中值不变（不在列表中则等下由级联删除逻辑处理）

    def refresh_moveload_cache(self):
        """重建移动荷载工况缓存 {case_name: {vehicle, lane_ecc}}"""
        self.moveload_cache.clear()
        for row in self.rows_data:
            try:
                ecc_f = float(row.get("lane_ecc", 0.0) or 0.0)
            except (ValueError, TypeError):
                ecc_f = 0.0  # 主表脏数据（如 "/"）兜底为 0，不阻断缓存刷新
            self.moveload_cache[row.get("name", "")] = {
                "vehicle": row.get("vehicle", ""),
                "lane_ecc": ecc_f,
                "direction": "往返",
                "factor": 1.0,
            }

    def on_vehicle_deleted(self, deleted_names):
        """VehicleTab 删除车辆后的联动处理"""
        # 1. 先过滤数据（必须在 refresh_combobox_values 之前，否则会被覆写）
        self.rows_data = [r for r in self.rows_data if r.get("vehicle", "") not in deleted_names]
        # 2. 再刷新下拉框值列表
        self.refresh_combobox_values()

        # 3. 完全重绘剩余行
        for w in list(self.table_inner.grid_slaves()):
            g = w.grid_info()
            if g and g.get("row") and int(g["row"]) > 0:
                w.destroy()

        vnames = self.get_vehicle_names()
        for ii, row_data in enumerate(self.rows_data):
            self._insert_row_ui(ii, ii + 1,
                row_data.get("vehicle", ""),
                row_data.get("lane_ecc", "0.0"),
                row_data.get("name", ""), vnames)

        self.refresh_moveload_cache()
        self.table_canvas.configure(scrollregion=self.table_canvas.bbox("all"))
        if self.rows_data:
            new_idx = min(getattr(self, "current_row", 0), len(self.rows_data) - 1)
            self.select(new_idx, force=True)
        else:
            self.current_row = -1
            self.ax.clear()
            self.ax.text(0.5, 0.5, "请添加工况", ha="center", va="center", transform=self.ax.transAxes,
                        fontsize=12, color="gray")
            self.canvas.draw()


# ======================== 车辆静载 Tab 页 ========================

class StaticLoadTab(ttk.Frame):
    def __init__(self, parent, parent_ui=None):
        super().__init__(parent)
        self.parent_ui = parent_ui

        self.padx = 10
        self.pady = 15
        self.label_width = 20

        self.car_dict = {}
        self.staticload_cache = {}
        self.car_output_check = {}
        self._current_car_name = ""
        self.pile_dict = {}
        self._pile_count = 5

        # 整体使用 grid 权重布局，使加载示意图获得合适空间
        self.grid_rowconfigure(0, weight=0)  # 车辆信息（auto）
        self.grid_rowconfigure(1, weight=0)  # 荷载类型（auto）
        self.grid_rowconfigure(2, weight=1)  # 跨间+桩顶（弹性）
        self.grid_rowconfigure(3, weight=2)  # 加载示意图（多占）
        self.grid_columnconfigure(0, weight=1)

        # ===== 车辆信息 LabelFrame (自动高度，内容超出可滚动) =====
        car_lf = ttk.LabelFrame(self, text="车辆信息", bootstyle=SECONDARY, padding=10)
        car_lf.grid(row=0, column=0, sticky="nsew", padx=12, pady=(10, 6))

        table_outer = ttk.Frame(car_lf, height=180)
        table_outer.pack(fill="x", padx=5, pady=5)
        table_outer.pack_propagate(False)

        self._table_canvas, self._table_frame = create_scrollable_frame(table_outer)

        headers = ["编号", "车辆", "工况输出"]
        table_weights = [1, 4, 2]
        for c, text in enumerate(headers):
            tk.Label(self._table_frame, text=text, font=("微软雅黑", 9, "bold"),
                     bg="#e8f4fd" if c < 2 else "#f2f2f2",
                     relief="solid", borderwidth=1, padx=8, pady=8).grid(row=0, column=c, sticky="nsew")
            self._table_frame.grid_columnconfigure(c, weight=table_weights[c], uniform="locked_cols")

        self._table_rows = []
        self._chk_vars = {}
        self.current_row = -1

        # ===== 荷载类型 LabelFrame =====
        self._load_lf = load_lf = ttk.LabelFrame(self, text="荷载类型", bootstyle=SECONDARY, padding=10)
        load_lf.grid(row=1, column=0, sticky="nsew", padx=12, pady=(6, 10))

        load_options = ["走行", "正吊(正向)", "正吊(反向)", "侧吊"]
        self.load_type_vars = {}
        radio_frame = ttk.Frame(load_lf)
        radio_frame.pack(fill="x", padx=5, pady=5)
        for c, opt in enumerate(load_options):
            var = tk.BooleanVar(value=(opt == "走行"))
            self.load_type_vars[opt] = var
            radio_frame.grid_columnconfigure(c, weight=1, uniform="radio_col")
            ttk.Checkbutton(radio_frame, text=opt, variable=var,
                            bootstyle="primary",
                            command=lambda: (self._save_current_state(), self.refresh_cache(), self._update_ui_state(), draw_static_load_diagram(self))).grid(row=0, column=c, sticky="w", padx=10)

        # ===== 跨间加载 & 桩顶加载 (左右并排，弹性高度) =====
        self._dual_frame = dual_frame = ttk.Frame(self)
        dual_frame.grid(row=2, column=0, sticky="nsew", padx=12, pady=(6, 10))
        dual_frame.grid_rowconfigure(0, weight=1)
        dual_frame.grid_columnconfigure(0, weight=3)
        dual_frame.grid_columnconfigure(1, weight=2)

        # ---------- 左：跨间加载 ----------
        span_lf = ttk.LabelFrame(dual_frame, text="跨间加载", bootstyle=SECONDARY, padding=10)
        span_lf.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        span_outer = ttk.Frame(span_lf)
        span_outer.pack(fill="both", expand=True)
        self._span_canvas, self._span_table = create_scrollable_frame(span_outer)

        span_headers = ["桥跨编号", "左侧", "跨中", "右侧"]
        span_weights = [2, 2, 2, 2]
        for c, text in enumerate(span_headers):
            tk.Label(self._span_table, text=text, font=("微软雅黑", 9, "bold"),
                     bg="#e8f4fd" if c < 1 else "#f2f2f2",
                     relief="solid", borderwidth=1, padx=8, pady=8).grid(row=0, column=c, sticky="nsew")
            self._span_table.grid_columnconfigure(c, weight=span_weights[c], uniform="span_cols")

        span_count = max(self._pile_count - 1, 0)
        self._span_chk_vars = []
        self._span_row_widgets = []
        self._span_current_row = -1
        for i in range(span_count):
            r = i + 1
            lbl_no = tk.Label(self._span_table, text=str(i+1), relief="solid",
                              borderwidth=1, padx=8, pady=8, bg="white", font=("微软雅黑", 9))
            lbl_no.grid(row=r, column=0, sticky="nsew")
            row_widgets = [lbl_no]
            row_vars = []
            # 默认: 第1行-左侧, 第2行-跨中, 第3行-右侧 勾选
            defaults = {0: [True, False, False], 1: [False, True, False], 2: [False, False, True]}
            for col_idx, col in enumerate(range(1, 4)):
                f = tk.Frame(self._span_table, relief="solid", borderwidth=1, bg="white", pady=8)
                f.grid(row=r, column=col, sticky="nsew")
                default_val = defaults.get(i, [False, False, False])[col_idx]
                var = tk.BooleanVar(value=default_val)
                ttk.Checkbutton(f, text="", variable=var, bootstyle="primary",
                                command=lambda: (self._save_current_state(), self.refresh_cache(), draw_static_load_diagram(self))).pack(expand=True)
                row_vars.append(var)
                row_widgets.append(f)
            self._span_chk_vars.append(row_vars)
            self._span_row_widgets.append(row_widgets)
            # 绑定点击选中
            for cell in row_widgets:
                cell.bind("<Button-1>", lambda e, idx=i: self._select_span_row(idx), add="+")
                for sub in cell.winfo_children():
                    sub.bind("<Button-1>", lambda e, idx=i: self._select_span_row(idx), add="+")

        # ---------- 右：桩顶加载 ----------
        pile_lf = ttk.LabelFrame(dual_frame, text="桩顶加载", bootstyle=SECONDARY, padding=10)
        pile_lf.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        pile_outer = ttk.Frame(pile_lf)
        pile_outer.pack(fill="both", expand=True)
        self._pile_canvas, self._pile_table = create_scrollable_frame(pile_outer)

        pile_headers = ["桩编号", "桩顶"]
        pile_weights = [2, 3]
        for c, text in enumerate(pile_headers):
            tk.Label(self._pile_table, text=text, font=("微软雅黑", 9, "bold"),
                     bg="#e8f4fd" if c < 1 else "#f2f2f2",
                     relief="solid", borderwidth=1, padx=8, pady=8).grid(row=0, column=c, sticky="nsew")
            self._pile_table.grid_columnconfigure(c, weight=pile_weights[c], uniform="pile_cols")

        self._pile_chk_vars = []
        self._pile_row_widgets = []
        self._pile_current_row = -1
        for i in range(self._pile_count):
            r = i + 1
            lbl_no = tk.Label(self._pile_table, text=str(i), relief="solid",
                              borderwidth=1, padx=8, pady=8, bg="white", font=("微软雅黑", 9))
            lbl_no.grid(row=r, column=0, sticky="nsew")
            f = tk.Frame(self._pile_table, relief="solid", borderwidth=1, bg="white", pady=8)
            f.grid(row=r, column=1, sticky="nsew")
            var = tk.BooleanVar(value=(i == 1))  # 默认勾选 1 号桩（与 _rebuild_loading_tables/_restore_vehicle_state 一致）
            ttk.Checkbutton(f, text="", variable=var, bootstyle="primary",
                            command=lambda: draw_static_load_diagram(self)).pack(expand=True)
            self._pile_chk_vars.append(var)
            self._pile_row_widgets.append((lbl_no, f))
            # 绑定点击选中
            for cell in (lbl_no, f):
                cell.bind("<Button-1>", lambda e, idx=i: self._select_pile_row(idx), add="+")
                for sub in cell.winfo_children():
                    sub.bind("<Button-1>", lambda e, idx=i: self._select_pile_row(idx), add="+")

        # ===== 车辆静载布置示意图 =====
        fig_w = max(self._pile_count * 1.2, 6)

        self._fig_frame = fig_frame = ttk.LabelFrame(self, text="车辆静载布置示意图", bootstyle=SECONDARY, padding=10)
        fig_frame.grid(row=3, column=0, sticky="nsew", padx=12, pady=(6, 10))
        self._load_fig = Figure(figsize=(fig_w, 2.5), dpi=100)
        self._load_ax = self._load_fig.add_subplot(111)
        self._load_canvas = FigureCanvasTkAgg(self._load_fig, master=fig_frame)
        tk_widget = self._load_canvas.get_tk_widget()
        tk_widget.pack()
        # 禁用画布滚轮和拖动
        tk_widget.bind("<MouseWheel>", lambda e: "break")
        tk_widget.bind("<Button-4>", lambda e: "break")
        tk_widget.bind("<Button-5>", lambda e: "break")
        tk_widget.bind("<B1-Motion>", lambda e: "break")
        tk_widget.bind("<ButtonPress-1>", lambda e: "break")

        # ===== 绘制固定元素 =====
        ax = self._load_ax
        center_y = 0.5
        x_max = self._pile_count - 1

        for i in range(self._pile_count):
            ax.axvline(x=i, color="#8B8B00", linestyle="--", linewidth=1.0, alpha=0.7)
        ax.axhline(y=center_y, color="#FF00FF", linestyle="-.", linewidth=1.2, alpha=0.8)

        ax.set_xlim(-0.5, x_max + 0.5)
        ax.set_ylim(-0.1, 1.1)
        ax.set_xticks(range(self._pile_count))
        ax.set_xticklabels([f"{i}#" for i in range(self._pile_count)], fontsize=9)
        ax.set_yticks([])
        for spine in ax.spines.values():
            spine.set_visible(False)
        ax.tick_params(axis="x", length=0)

        self._load_fig.subplots_adjust(left=0.05, right=0.97, top=0.88, bottom=0.15)
        self._load_drawn_elements = []

    def reload_from_dict(self, new_dict):
        """数据绑定更新：从 new_dict 全量刷新静载参数（统一 reload 契约）

        依赖 _page_defs 顺序契约：本 tab 必须排在 SupportsTab（substructure_cache）
        与 VehicleTab（vehicle_cache）之后 reload，保证下面重建读到新项目数据。
        """
        self.staticload_cache.clear()
        self.car_output_check.clear()
        self._current_car_name = ""
        self.pile_dict = {}
        # car_dict 必须同步清空：此前只清 car_output_check 不清 car_dict，
        # 切项目后 _on_page_shown 对"已在 car_dict 中"的车辆跳过重建，
        # 导致 car_output_check 缺键 → _on_page_shown 读取时 KeyError
        self.car_dict.clear()

        # 同步桩数并重建表格
        new_pile_count = self._get_pile_count_from_substructure()
        if new_pile_count != self._pile_count:
            self._pile_count = new_pile_count
            self.pile_dict = {i: [] for i in range(new_pile_count)}
            self._rebuild_loading_tables()
            self._rebuild_diagram_axes()

        # 全量重建 car_dict + 车辆表格 + staticload_cache：切项目后 UI 与缓存
        # 立即呈现新项目数据（不再依赖用户切走再切回触发 _on_page_shown）。
        # car_dict 已清空，_on_page_shown 的合并分支（name in car_dict 保留）全部走新建。
        # 顺带保证：切走时 combine 读到的 staticload_cache 非空 → 兜底不触发 → 不污染。
        self._on_page_shown()

    def _get_pile_count_from_substructure(self):
        """从 substructure_cache 读取桩数，无数据时返回默认 5"""
        if self.parent_ui and hasattr(self.parent_ui, 'tab_supports'):
            sc = getattr(self.parent_ui.tab_supports, 'substructure_cache', {})
            if sc:
                return len(sc)
        return 5

    def _rebuild_loading_tables(self):
        """清空并重建跨间加载和桩顶加载表格"""
        # --- 跨间加载 ---
        for r in range(1, len(self._span_chk_vars) + 1):
            for w in self._span_table.grid_slaves(row=r):
                w.destroy()
        self._span_chk_vars.clear()
        self._span_row_widgets.clear()
        self._span_current_row = -1

        span_count = max(self._pile_count - 1, 0)
        for i in range(span_count):
            r = i + 1
            lbl_no = tk.Label(self._span_table, text=str(i+1), relief="solid",
                              borderwidth=1, padx=8, pady=8, bg="white", font=("微软雅黑", 9))
            lbl_no.grid(row=r, column=0, sticky="nsew")
            row_widgets = [lbl_no]
            row_vars = []
            defaults = {0: [True, False, False], 1: [False, True, False], 2: [False, False, True]}
            for col_idx, col in enumerate(range(1, 4)):
                f = tk.Frame(self._span_table, relief="solid", borderwidth=1, bg="white", pady=8)
                f.grid(row=r, column=col, sticky="nsew")
                default_val = defaults.get(i, [False, False, False])[col_idx]
                var = tk.BooleanVar(value=default_val)
                ttk.Checkbutton(f, text="", variable=var, bootstyle="primary",
                                command=lambda: (self._save_current_state(), self.refresh_cache(), draw_static_load_diagram(self))).pack(expand=True)
                row_vars.append(var)
                row_widgets.append(f)
            self._span_chk_vars.append(row_vars)
            self._span_row_widgets.append(row_widgets)
            for cell in row_widgets:
                cell.bind("<Button-1>", lambda e, idx=i: self._select_span_row(idx), add="+")
                for sub in cell.winfo_children():
                    sub.bind("<Button-1>", lambda e, idx=i: self._select_span_row(idx), add="+")

        # --- 桩顶加载 ---
        for r in range(1, len(self._pile_chk_vars) + 1):
            for w in self._pile_table.grid_slaves(row=r):
                w.destroy()
        self._pile_chk_vars.clear()
        self._pile_row_widgets.clear()
        self._pile_current_row = -1

        for i in range(self._pile_count):
            r = i + 1
            lbl_no = tk.Label(self._pile_table, text=str(i), relief="solid",
                              borderwidth=1, padx=8, pady=8, bg="white", font=("微软雅黑", 9))
            lbl_no.grid(row=r, column=0, sticky="nsew")
            f = tk.Frame(self._pile_table, relief="solid", borderwidth=1, bg="white", pady=8)
            f.grid(row=r, column=1, sticky="nsew")
            var = tk.BooleanVar(value=(i == 1))
            ttk.Checkbutton(f, text="", variable=var, bootstyle="primary",
                            command=lambda: draw_static_load_diagram(self)).pack(expand=True)
            self._pile_chk_vars.append(var)
            self._pile_row_widgets.append((lbl_no, f))
            for cell in (lbl_no, f):
                cell.bind("<Button-1>", lambda e, idx=i: self._select_pile_row(idx), add="+")
                for sub in cell.winfo_children():
                    sub.bind("<Button-1>", lambda e, idx=i: self._select_pile_row(idx), add="+")

    def _rebuild_diagram_axes(self):
        """重建示意图的固定元素（桩位线、刻度）"""
        if not hasattr(self, '_load_ax'):
            return
        ax = self._load_ax
        ax.clear()
        x_max = self._pile_count - 1
        for i in range(self._pile_count):
            ax.axvline(x=i, color="#8B8B00", linestyle="--", linewidth=1.0, alpha=0.7)
        ax.axhline(y=0.5, color="#FF00FF", linestyle="-.", linewidth=1.2, alpha=0.8)
        ax.set_xlim(-0.5, x_max + 0.5)
        ax.set_ylim(-0.1, 1.1)
        ax.set_xticks(range(self._pile_count))
        ax.set_xticklabels([f"{i}#" for i in range(self._pile_count)], fontsize=9)
        ax.set_yticks([])
        for spine in ax.spines.values():
            spine.set_visible(False)
        ax.tick_params(axis="x", length=0)
        self._load_fig.subplots_adjust(left=0.05, right=0.97, top=0.88, bottom=0.15)
        self._load_canvas.draw_idle()

    def refresh_cache(self):
        """重建 staticload_cache
        {工况名: ("跨间加载", [[bool],...], "桩顶加载", [bool,...])}
        仅 output_enabled=True、有激活荷载类型、有激活加载位置时生成
        """
        self.staticload_cache.clear()
        for vname, data in self.car_dict.items():
            if not self.car_output_check.get(vname, False):
                continue  # 工况输出未激活 → 跳过

            lt = data.get("load_types", {})
            active_types = [opt for opt, v in lt.items() if v]
            if not active_types:
                continue  # 无激活荷载类型 → 跳过

            span_checks = [list(row) for row in data.get("span_checks", [])]
            pile_checks = list(data.get("pile_checks", []))
            has_span = any(any(row) for row in span_checks)
            has_pile = any(pile_checks)
            if not has_span and not has_pile:
                continue  # 无激活加载位置 → 跳过

            for load_type in active_types:
                key_prefix = f"{vname}{load_type}"
                entry_base = {"vehicle": vname, "type": load_type}
                # 跨间各位置 → 逐个展开为独立工况名
                for col, pos_label, suffix_fmt in [
                    (0, "跨间左侧", "跨{}左侧"),
                    (1, "跨中",     "跨{}跨中"),
                    (2, "跨间右侧", "跨{}右侧"),
                ]:
                    active = [i + 1 for i, row in enumerate(span_checks) if col < len(row) and row[col]]
                    for span_no in active:
                        case_name = f"{key_prefix}{suffix_fmt.format(span_no)}"
                        self.staticload_cache[case_name] = {
                            **entry_base, pos_label: [span_no]
                        }
                # 桩顶 → 逐个展开
                active_piles = [i for i, v in enumerate(pile_checks) if v]
                for pile_no in active_piles:
                    case_name = f"{key_prefix}桩{pile_no}顶"
                    self.staticload_cache[case_name] = {
                        **entry_base, "桩顶": [pile_no]
                    }

    @staticmethod
    def _set_widgets_state(frame, state):
        """递归设置 frame 内所有控件的启用/禁用状态"""
        for child in frame.winfo_children():
            try:
                if isinstance(child, ttk.Frame):
                    pass  # ttk.Frame 无 state 参数，递归处理其子控件
                elif isinstance(child, (ttk.Checkbutton, ttk.LabelFrame, ttk.Combobox, ttk.Button, ttk.Entry, ttk.Label)):
                    child.configure(state=state)
                elif isinstance(child, (tk.Frame, tk.Label, tk.Entry, tk.Canvas)):
                    try:
                        child.configure(state=state)
                    except:
                        pass
            except:
                pass
            if child.winfo_children():
                StaticLoadTab._set_widgets_state(child, state)

    def _update_ui_state(self):
        """根据车辆状态启用/禁用下方控件（不禁用示意图画布）"""
        car_name = self._current_car_name
        output_on = self.car_output_check.get(car_name, False) if car_name else False
        has_active_type = any(v.get() for v in self.load_type_vars.values())

        # 荷载类型区：output_on → 启用，否则禁用
        self._set_widgets_state(self._load_lf, "normal" if output_on else "disabled")
        if output_on:
            self._load_lf.configure(text="荷载类型")
        else:
            self._load_lf.configure(text="荷载类型 (请先激活上方工况输出)")

        # 跨间/桩顶区：output_on + has_active_type → 启用
        dual_state = "normal" if (output_on and has_active_type) else "disabled"
        self._set_widgets_state(self._dual_frame, dual_state)

        # 示意图始终可见，但禁用交互
        for child in self._fig_frame.winfo_children():
            if isinstance(child, tk.Canvas):
                # canvas 是绘图区域，不禁用
                pass
            else:
                try:
                    child.configure(state="normal" if output_on else "disabled")
                except:
                    pass

    def _save_current_state(self):
        """将当前界面的荷载类型/跨间/桩顶状态保存到 car_dict 中当前选中车辆"""
        if not self._current_car_name or self._current_car_name not in self.car_dict:
            return
        data = self.car_dict[self._current_car_name]
        lt = data.setdefault("load_types", {})
        for opt, var in self.load_type_vars.items():
            lt[opt] = var.get()
        sc = data.setdefault("span_checks", [])
        sc.clear()
        for row_vars in self._span_chk_vars:
            sc.append([v.get() for v in row_vars])
        pc = data.setdefault("pile_checks", [])
        pc.clear()
        for var in self._pile_chk_vars:
            pc.append(var.get())

    def _restore_vehicle_state(self, car_name):
        """从 car_dict 恢复指定车辆的状态到界面控件"""
        if car_name not in self.car_dict:
            return
        data = self.car_dict[car_name]
        lt = data.get("load_types", {})
        for opt, var in self.load_type_vars.items():
            var.set(lt.get(opt, opt == "走行"))
        sc = data.get("span_checks", [])
        for i, row_vars in enumerate(self._span_chk_vars):
            if i < len(sc):
                for j, v in enumerate(row_vars):
                    v.set(sc[i][j] if j < len(sc[i]) else False)
            else:
                for v in row_vars:
                    v.set(False)
        pc = data.get("pile_checks", [])
        for i, var in enumerate(self._pile_chk_vars):
            var.set(pc[i] if i < len(pc) else (i == 1))
        self.car_output_check[car_name] = self.car_output_check.get(car_name, True)
        self._update_ui_state()
        draw_static_load_diagram(self)



    def _on_page_shown(self):
        """当用户切换到此页面时调用，刷新车辆列表"""
        # 0. 从 substructure_cache 同步桩数
        new_pile_count = self._get_pile_count_from_substructure()
        if new_pile_count != self._pile_count:
            self._pile_count = new_pile_count
            self.pile_dict = {i: [] for i in range(new_pile_count)}
            self._rebuild_loading_tables()
            self._rebuild_diagram_axes()

        # 1. 从 VehicleTab 获取最新车辆列表，同步 car_dict
        vnames = []
        if self.parent_ui and hasattr(self.parent_ui, 'tab_vehicles'):
            vnames = list(self.parent_ui.tab_vehicles.vehicle_cache.keys())
        pile_count = self._pile_count
        span_count = max(pile_count - 1, 0)

        # 2. 从当前项目数据读取静载配置（cache 格式），按车辆名合并
        #    static_load_cases: {工况名: {vehicle, type, 跨间左侧/跨中/跨间右侧/桩顶}}
        sl_vehicle_map = {}
        for _entry in trestle_excel_dict.get("static_load_cases", {}).values():
            vname = _entry.get("vehicle", "")
            if not vname:
                continue
            if vname not in sl_vehicle_map:
                sl_vehicle_map[vname] = {
                    "types": [], "span_left": [], "span_mid": [], "span_right": [], "pile": [],
                }
            _t = _entry.get("type", "")
            if _t and _t not in sl_vehicle_map[vname]["types"]:
                sl_vehicle_map[vname]["types"].append(_t)
            for _src, _dst in (("跨间左侧", "span_left"), ("跨中", "span_mid"), ("跨间右侧", "span_right")):
                for _n in _entry.get(_src, []):
                    if _n not in sl_vehicle_map[vname][_dst]:
                        sl_vehicle_map[vname][_dst].append(_n)
            # 桩号为 0-based（与 UI 桩顶表显示编号一致）
            for _p in _entry.get("桩顶", []):
                if _p not in sl_vehicle_map[vname]["pile"]:
                    sl_vehicle_map[vname]["pile"].append(_p)

        # 3. 初始化每辆车 — 仅保留吊装设备/钻孔设备（其余无静载面荷载分布）
        crane_names = []
        vcache = self.parent_ui.tab_vehicles.vehicle_cache if self.parent_ui and hasattr(self.parent_ui, 'tab_vehicles') else {}
        for n in vnames:
            info = vcache.get(n, {})
            if info.get("vehicle_type", "") in ("吊装设备", "钻孔设备", "自定义-吊装设备", "自定义-钻孔设备"):
                crane_names.append(n)
        # 移除已存在的非吊装车辆
        for k in list(self.car_dict.keys()):
            if k not in crane_names:
                del self.car_dict[k]
                self.car_output_check.pop(k, None)
        for name in crane_names:
            if name not in self.car_dict:
                if name in sl_vehicle_map:
                    sv = sl_vehicle_map[name]
                    base_types = {"走行": False, "正吊(正向)": False, "正吊(反向)": False, "侧吊": False}
                    for t in sv.get("types", []):
                        if t in base_types:
                            base_types[t] = True
                    default_span = []
                    for i in range(span_count):
                        row = [False, False, False]
                        span_no = i + 1
                        if span_no in sv.get("span_left", []):
                            row[0] = True
                        if span_no in sv.get("span_mid", []):
                            row[1] = True
                        if span_no in sv.get("span_right", []):
                            row[2] = True
                        default_span.append(row)
                    pile_checks = [False] * pile_count
                    for pi in sv.get("pile", []):
                        # pi 为 0-based 桩号，与 UI 桩顶表显示编号一致，直接映射数组下标
                        if 0 <= pi < pile_count:
                            pile_checks[pi] = True
                    self.car_dict[name] = {
                        "load_types": base_types,
                        "span_checks": default_span,
                        "pile_checks": pile_checks,
                    }
                    self.car_output_check[name] = True
                else:
                    self.car_dict[name] = {
                        "load_types": {"走行": False, "正吊(正向)": False, "正吊(反向)": False, "侧吊": False},
                        "span_checks": [[False, False, False] for _ in range(span_count)],
                        "pile_checks": [False] * pile_count,
                    }
                    self.car_output_check[name] = False
        # 删除已不存在的车辆
        for k in list(self.car_dict.keys()):
            if k not in vnames:
                del self.car_dict[k]

        car_lst = list(self.car_dict.keys())

        # 清空旧行
        for widgets in self._table_rows:
            for w in widgets:
                w.destroy()
        self._table_rows.clear()
        self._chk_vars.clear()

        # car_output_check 已在每辆车初始化时设置

        # 填充表格行
        for i, car_name in enumerate(car_lst):
            row_idx = i + 1
            lbl_no = tk.Label(self._table_frame, text=str(i+1), relief="solid",
                              borderwidth=1, padx=8, pady=8, bg="white", font=("微软雅黑", 9))
            lbl_no.grid(row=row_idx, column=0, sticky="nsew")
            lbl_car = tk.Label(self._table_frame, text=car_name, relief="solid",
                               borderwidth=1, padx=8, pady=8, bg="#fcfcfc", fg="#666",
                               font=("微软雅黑", 9))
            lbl_car.grid(row=row_idx, column=1, sticky="nsew")
            chk_frame = tk.Frame(self._table_frame, relief="solid", borderwidth=1, bg="white", pady=8)
            chk_frame.grid(row=row_idx, column=2, sticky="nsew")
            var = tk.BooleanVar(value=self.car_output_check[car_name])
            chk = ttk.Checkbutton(chk_frame, text="", variable=var, bootstyle="primary",
                                  command=lambda n=car_name, v=var: self._on_check_change(n, v))
            chk.pack(expand=True, padx=5, pady=2)
            self._chk_vars[car_name] = var
            self._table_rows.append((lbl_no, lbl_car, chk_frame, chk))
            cells = [lbl_no, lbl_car, chk_frame]
            for cell in cells:
                cell.bind("<Button-1>", lambda e, idx=i: self._select(idx), add="+")
                for sub in cell.winfo_children():
                    sub.bind("<Button-1>", lambda e, idx=i: self._select(idx), add="+")

        # 默认选中第一行，恢复其状态
        # 重建缓存
        self.refresh_cache()
        if self._table_rows:
            self._select(0, force=True)

    def _on_check_change(self, car_name, var):
        """复选框（工况输出）状态变更时同步到字典并刷新图表"""
        self.car_output_check[car_name] = var.get()
        self.refresh_cache()
        if self._current_car_name == car_name:
            self._update_ui_state()
            draw_static_load_diagram(self)

    def _select(self, idx, force=False):
        """选中行：保存当前车辆状态 → 恢复新车辆状态 → 高亮"""
        if not self._table_rows or idx < 0 or idx >= len(self._table_rows):
            return
        car_lst = list(self.car_dict.keys())
        if idx >= len(car_lst):
            return

        # 保存当前车辆状态
        self._save_current_state()

        new_car = car_lst[idx]
        self.current_row = idx
        self._current_car_name = new_car

        # 恢复新选中车辆的状态
        self._restore_vehicle_state(new_car)
        self.refresh_cache()

        # 高亮颜色刷新
        for r in range(len(self._table_rows)):
            bg = "#e3f2fd" if r == idx else "white"
            for w in self._table_frame.grid_slaves(row=r + 1):
                try:
                    w.config(bg=bg)
                except:
                    pass
                if isinstance(w, tk.Frame):
                    for sub in w.winfo_children():
                        try:
                            sub.config(bg=bg)
                        except:
                            pass

    def _select_span_row(self, idx):
        """跨间加载表格行选中高亮"""
        self._span_current_row = idx
        draw_static_load_diagram(self)
        for r, widgets in enumerate(self._span_row_widgets):
            bg = "#e3f2fd" if r == idx else "white"
            for w in widgets:
                try:
                    w.config(bg=bg)
                except:
                    pass
                if isinstance(w, tk.Frame):
                    for sub in w.winfo_children():
                        try:
                            sub.config(bg=bg)
                        except:
                            pass

    def _select_pile_row(self, idx):
        """桩顶加载表格行选中高亮"""
        self._pile_current_row = idx
        draw_static_load_diagram(self)
        for r, (lbl_no, f) in enumerate(self._pile_row_widgets):
            bg = "#e3f2fd" if r == idx else "white"
            for w in (lbl_no, f):
                try:
                    w.config(bg=bg)
                except:
                    pass
                if isinstance(w, tk.Frame):
                    for sub in w.winfo_children():
                        try:
                            sub.config(bg=bg)
                        except:
                            pass

    def get_car_output_check(self):
        """返回当前车辆工况输出状态字典 {'车辆名': True/False}"""
        return self.car_output_check.copy()

    def StaticLoad_return(self):
        """返回当前参数（供外部调用）"""
        self._save_current_state()
        self.refresh_cache()
        return {
            "staticload_cache": self.staticload_cache, 
            "pile_dict": self.pile_dict,
        }

    def on_vehicle_deleted(self, deleted_names):
        """VehicleTab 删除车辆后的联动处理"""
        changed = False
        for vname in list(self.car_dict.keys()):
            if vname in deleted_names:
                del self.car_dict[vname]
                self.car_output_check.pop(vname, None)
                changed = True
        if changed:
            self._on_page_shown()
            self.refresh_cache()

    def on_vehicle_renamed(self, old_name, new_name):
        """VehicleTab 车辆改名后的联动：更新 car_dict 中的 key"""
        if old_name in self.car_dict and old_name != new_name:
            self.car_dict[new_name] = self.car_dict.pop(old_name)
            self.car_output_check[new_name] = self.car_output_check.pop(old_name, False)
            self._on_page_shown()
            self.refresh_cache()


# ======================== 荷载组合 Tab 页 ========================
class LoadCombinationTab(ttk.Frame):
    """荷载组合 Tab: 上表(组合定义) + 下表(包含的工况及系数)"""
    def __init__(self, parent, parent_ui=None):
        super().__init__(parent)
        self.parent_ui = parent_ui

        # 核心数据结构:
        # [ {"name": "cCB1", "type": "相加", "cases": [ {"case_name": "恒载", "gamma": "1.0", "psi": "1.2"}, ... ] }, ... ]
        self.comb_data = []
        self.current_comb_idx = -1
        self.current_case_idx = -1

        self.available_load_cases = []
        self.load_combination_cache = {}  # {comb_name: {"type": str, "cases": [...]}}
        self.collect_basic_cases()

        # ========================================================
        # 布局比例：上下各占 50% (weight=5)
        # ========================================================
        self.grid_rowconfigure(0, weight=5, uniform="comb_split")
        self.grid_rowconfigure(1, weight=5, uniform="comb_split")
        self.grid_columnconfigure(0, weight=1)

        # ==================== 1. 上半部分：荷载组合定义表 ====================
        top_container = ttk.Frame(self)
        top_container.grid(row=0, column=0, sticky="nsew", padx=5, pady=(5, 0))

        ttk.Label(top_container, text="荷载组合定义表", font=("Microsoft YaHei UI", 10, "bold"),
                  bootstyle=PRIMARY).pack(anchor="w", padx=10, pady=(5, 5))

        top_btn_bar = ttk.Frame(top_container)
        top_btn_bar.pack(fill="x", padx=5, pady=5)
        tb.Button(top_btn_bar, text="+ 添加", bootstyle="info-outline", command=self.add_comb_row).pack(side="left", padx=5)
        tb.Button(top_btn_bar, text="- 删除", bootstyle="danger-outline", command=self.delete_comb_row).pack(side="left", padx=5)
        tb.Button(top_btn_bar, text="复制", bootstyle="primary-outline", command=self.copy_comb_row).pack(side="left", padx=5)
        tb.Button(top_btn_bar, text="自动组合", bootstyle=PRIMARY, command=self._auto_combine).pack(side="right", padx=5)

        self.top_canvas, self.top_inner = create_scrollable_frame(top_container, pack_kwargs={"padx": 8, "pady": 10})

        top_headers = ["序号", "荷载组合名称", "类型"]
        top_weights = [1, 5, 4]
        for c, (t, w) in enumerate(zip(top_headers, top_weights)):
            tk.Label(self.top_inner, text=t, font=("微软雅黑", 9, "bold"),
                     bg="#e8f4fd" if c < 2 else "#f2f2f2",
                     relief="solid", borderwidth=1, padx=8, pady=5).grid(row=0, column=c, sticky="nsew")
            self.top_inner.grid_columnconfigure(c, weight=w, uniform="top_cols")

        # ==================== 2. 下半部分：对应荷载工况和系数 ====================
        bottom_container = ttk.Frame(self)
        bottom_container.grid(row=1, column=0, sticky="nsew", padx=5, pady=(5, 5))
        
        # 顶部分隔线
        ttk.Separator(bottom_container, orient="horizontal").pack(fill="x", padx=10, pady=(0, 5))

        self.bottom_title = ttk.Label(bottom_container, text="荷载工况和系数 (请先选择上方组合)", font=("Microsoft YaHei UI", 10, "bold"), bootstyle=PRIMARY)
        self.bottom_title.pack(anchor="w", padx=10, pady=(5, 5))

        bot_btn_bar = ttk.Frame(bottom_container)
        bot_btn_bar.pack(fill="x", padx=5, pady=5)
        self.btn_add_case = tb.Button(bot_btn_bar, text="+ 添加工况", bootstyle="info-outline", state="normal", command=self.add_case_row)
        self.btn_add_case.pack(side="left", padx=5)
        self.btn_del_case = tb.Button(bot_btn_bar, text="- 删除工况", bootstyle="danger-outline", state="normal", command=self.delete_case_row)
        self.btn_del_case.pack(side="left", padx=5)

        self.bot_canvas, self.bot_inner = create_scrollable_frame(bottom_container, pack_kwargs={"padx": 8, "pady": 5})

        bot_headers = ["序号", "荷载工况", "分项系数", "组合值系数"]
        bot_weights = [1, 5, 3, 3]
        for c, (t, w) in enumerate(zip(bot_headers, bot_weights)):
            tk.Label(self.bot_inner, text=t, font=("微软雅黑", 9, "bold"),
                     bg="#e8f4fd" if c < 2 else "#f2f2f2",
                     relief="solid", borderwidth=1, padx=8, pady=5).grid(row=0, column=c, sticky="nsew")
            self.bot_inner.grid_columnconfigure(c, weight=w, uniform="bot_cols")

        # 从 trestle_excel_dict 预填充所有荷载组合行（cache 格式，含工况/系数）
        comb_cache = trestle_excel_dict.get("load_combination", {})
        for comb_name, comb_info in comb_cache.items():
            cases_list = []
            for case_name, case_coeff in comb_info.get("cases", {}).items():
                cases_list.append({
                    "case_name": case_name,
                    "gamma": case_coeff.get("gamma", "1.0"),
                    "psi": case_coeff.get("psi", "1.0"),
                })
            self.add_comb_row(name=comb_name, ctype=comb_info.get("type", "相加"), cases=cases_list)

    def reload_from_dict(self, new_dict):
        """数据绑定更新：从 new_dict（cache 格式）刷新荷载组合"""
        # 清空数据
        self.comb_data.clear()
        self.current_comb_idx = -1
        self.current_case_idx = -1
        self.load_combination_cache.clear()

        # 删除顶表所有数据行（保留 header row 0）
        for w in list(self.top_inner.grid_slaves()):
            g = w.grid_info()
            if g and g.get("row") and int(g["row"]) > 0:
                w.destroy()

        # 删除底表所有数据行（保留 header row 0）
        for w in list(self.bot_inner.grid_slaves()):
            g = w.grid_info()
            if g and g.get("row") and int(g["row"]) > 0:
                w.destroy()

        # 重置底表按钮状态
        self.btn_add_case.config(state="normal")
        self.btn_del_case.config(state="normal")

        # 重新收集基本工况
        self.collect_basic_cases()

        # cache 格式：{comb_name: {"type": str, "cases": {case_name: {"gamma", "psi"}}}}
        comb_cache = new_dict.get("load_combination", {})
        for comb_name, comb_info in comb_cache.items():
            cases_list = []
            for case_name, case_coeff in comb_info.get("cases", {}).items():
                cases_list.append({
                    "case_name": case_name,
                    "gamma": case_coeff.get("gamma", "1.0"),
                    "psi": case_coeff.get("psi", "1.0"),
                })
            self.add_comb_row(name=comb_name, ctype=comb_info.get("type", "相加"), cases=cases_list)

    def _auto_combine(self):
        """自动生成荷载组合（标准+基本 各一套）"""
        from itertools import product as iter_product

        # ── 0. 收集移动荷载和静载工况 ──
        # 未点开过对应 tab 时 cache 为空 → 回退项目数据（cache 格式，结构一致）
        move_cases = {}
        static_cases = {}
        if self.parent_ui:
            if hasattr(self.parent_ui, 'tab_moveloads'):
                move_cases = dict(self.parent_ui.tab_moveloads.moveload_cache)
            if hasattr(self.parent_ui, 'tab_staticloads'):
                static_cases = dict(self.parent_ui.tab_staticloads.staticload_cache)
            if not move_cases:
                move_cases = dict(self.parent_ui.trestle_excel_dict.get("move_load_cases", {}))
            if not static_cases:
                static_cases = dict(self.parent_ui.trestle_excel_dict.get("static_load_cases", {}))

        # ── 1. 清空现有组合 ──
        self.comb_data.clear()
        for w in list(self.top_inner.grid_slaves()):
            g = w.grid_info()
            if g and g.get("row") and int(g["row"]) > 0:
                w.destroy()

        # 固定工况系数（标准 / 基本）
        # 格式: {工况名: (分项系数, 组合值系数)}
        FIXED_STANDARD = {"自重": ("1.0", "1.0"), "风荷载": ("1.0", "0.6"), "水流力": ("1.0", "0.7")}
        FIXED_BASIC    = {"自重": ("1.0", "1.2"), "风荷载": ("1.4", "0.6"), "水流力": ("1.4", "0.7")}
        MOVE_STANDARD  = ("1.0", "1.0")   # 移动荷载标准组合 (分项, 组合值)
        MOVE_BASIC     = ("1.0", "1.4")   # 移动荷载基本组合 (分项, 组合值)
        STATIC_STANDARD = ("1.0", "1.0")  # 静载标准组合 (分项, 组合值)
        STATIC_BASIC    = ("1.0", "1.4")  # 静载基本组合 (分项, 组合值)

        comb_idx = 0
        _std_names = []  # 收集所有标准组合名称，用于最后包络
        _basic_names = []  # 收集所有基本组合名称，用于最后包络

        # ── 辅助：生成一组组合（标准+基本，共享工况号） ──
        def _add_pair(name_prefix, variable_cases, std_factor, basic_factor):
            """生成标准组合和基本组合各一个，共享同一工况号
            std_factor/basic_factor: (分项系数, 组合值系数) 元组
            """
            nonlocal comb_idx
            comb_idx += 1
            for suffix, fixed_map, var_factor, name_list in [
                ("标准组合", FIXED_STANDARD, std_factor, _std_names),
                ("基本组合", FIXED_BASIC,    basic_factor, _basic_names),
            ]:
                cases_list = [{"case_name": k, "gamma": v[0], "psi": v[1]} for k, v in fixed_map.items()]
                for c in variable_cases:
                    cases_list.append({"case_name": c, "gamma": var_factor[0], "psi": var_factor[1]})
                comb_name = f"工况{num_to_chinese_num(comb_idx)}{suffix}"
                name_list.append(comb_name)
                self.add_comb_row(
                    name=comb_name,
                    ctype="相加",
                    cases=cases_list
                )

        # ── 2. 移动荷载组合 ──
        vcache = {}
        if self.parent_ui and hasattr(self.parent_ui, 'tab_vehicles'):
            vcache = self.parent_ui.tab_vehicles.vehicle_cache or {}

        # 按车辆名分组
        veh_cases = {}  # {vehicle_name: [(case_name, lane_ecc), ...]}
        for cname, cdata in move_cases.items():
            veh = cdata.get("vehicle", "")
            if veh:
                ecc = cdata.get("lane_ecc", "0")
                veh_cases.setdefault(veh, []).append((cname, str(ecc)))

        # 按车辆类型分类
        crane_vehicles = {}   # 吊装设备
        drill_vehicles = {}   # 钻孔设备
        general_vehicles = {} # 一般车辆
        highway_vehicles = {} # 公路标准荷载

        for veh, case_list in veh_cases.items():
            vinfo = vcache.get(veh, {})
            vtype = vinfo.get("vehicle_type", "")
            if vtype in ("吊装设备", "自定义-吊装设备"):
                crane_vehicles[veh] = case_list
            elif vtype in ("钻孔设备", "自定义-钻孔设备"):
                drill_vehicles[veh] = case_list
            elif vtype == "一般车辆":
                general_vehicles[veh] = case_list
            elif vtype == "公路标准荷载":
                highway_vehicles[veh] = case_list

        def _pick_heaviest(veh_dict):
            """从车辆字典中取自重最大的一辆车"""
            best_veh, best_cases, best_weight = None, [], -1
            for v, cl in veh_dict.items():
                w = self._get_vehicle_weight(vcache.get(v, {}))
                if w > best_weight:
                    best_veh, best_cases, best_weight = v, cl, w
            return best_veh, best_cases

        def _pick_heaviest_axle_sum(veh_dict):
            """从一般车辆字典中取轴重之和最大的一辆车（轴重CSV位于 vehicle_params[1]）"""
            best_veh, best_cases, best_sum = None, [], -1
            for v, cl in veh_dict.items():
                params = vcache.get(v, {}).get("vehicle_params", [])
                s = 0
                if len(params) > 1:
                    for x in str(params[1]).split(","):
                        try:
                            s += float(x)
                        except (ValueError, TypeError):
                            pass
                if s > best_sum:
                    best_veh, best_cases, best_sum = v, cl, s
            return best_veh, best_cases

        # 吊装/钻孔：各取自重最大的一辆
        for veh_dict in (crane_vehicles, drill_vehicles):
            if not veh_dict:
                continue
            veh, case_list = _pick_heaviest(veh_dict)
            if veh and case_list:
                for cname, ecc in case_list:
                    label = f"{veh}(偏心{ecc}m)" if ecc and ecc not in ("0", "0.0") else veh
                    _add_pair(f"{label}走行", [cname], MOVE_STANDARD, MOVE_BASIC)

        # 一般车辆：取轴重之和最大的一辆，每个移动荷载工况（不同偏心）各自生成组合
        if general_vehicles:
            veh, case_list = _pick_heaviest_axle_sum(general_vehicles)
            if veh and case_list:
                for cname, ecc in case_list:
                    label = f"{veh}(偏心{ecc}m)" if ecc and ecc not in ("0", "0.0") else veh
                    _add_pair(f"{label}走行", [cname], MOVE_STANDARD, MOVE_BASIC)

        # 公路荷载：取第一个车辆，每个工况（不同偏心）各自生成
        if highway_vehicles:
            veh, case_list = next(iter(highway_vehicles.items()))
            for cname, ecc in case_list:
                label = f"{veh}(偏心{ecc}m)" if ecc and ecc not in ("0", "0.0") else veh
                _add_pair(f"{label}走行", [cname], MOVE_STANDARD, MOVE_BASIC)

        # ── 3. 静载组合：按车辆分组 → 邻跨约束排列 ──
        if static_cases:
            veh_static = {}
            for cname, cdata in static_cases.items():
                veh = cdata.get("vehicle", "")
                if veh:
                    veh_static.setdefault(veh, []).append(cname)

            for veh, cnames in veh_static.items():
                case_info = []
                for cname in cnames:
                    cdata = static_cases[cname]
                    span_no = pos_type = None
                    if "桩顶" in cdata:
                        span_no = cdata["桩顶"][0] + 1
                        pos_type = "pile"
                    elif "跨间左侧" in cdata:
                        span_no = cdata["跨间左侧"][0]
                        pos_type = "left"
                    elif "跨中" in cdata:
                        span_no = cdata["跨中"][0]
                        pos_type = "mid"
                    elif "跨间右侧" in cdata:
                        span_no = cdata["跨间右侧"][0]
                        pos_type = "right"
                    if span_no is not None:
                        case_info.append((cname, span_no, pos_type))

                if not case_info:
                    continue

                span_cases = {}
                for cname, sno, ptype in case_info:
                    span_cases.setdefault(sno, []).append((cname, ptype))

                all_spans = sorted(span_cases.keys())
                # 根据总跨数自动计算邻跨间隔：跨数越多间隔越大，保证约4组交替覆盖全桥
                n_gap = max(1, (len(all_spans) - 1) // 4)
                valid_groups = self._generate_valid_span_groups(all_spans, n_gap)

                static_idx = 0
                for group in valid_groups:
                    options_per_span = [ [c[0] for c in span_cases[sno]] for sno in group ]
                    for combo in iter_product(*options_per_span):
                        static_idx += 1
                        _add_pair(
                            f"{veh}静载组合{static_idx}#",
                            list(combo),
                            STATIC_STANDARD, STATIC_BASIC
                        )

        # ── 4. 包络组合：汇总所有标准/基本组合 ──
        if _std_names:
            envelope_std_cases = [{"case_name": n, "gamma": "1.0", "psi": "1.0"} for n in _std_names]
            self.add_comb_row(name="标准组合", ctype="包络", cases=envelope_std_cases)
        if _basic_names:
            envelope_basic_cases = [{"case_name": n, "gamma": "1.0", "psi": "1.0"} for n in _basic_names]
            self.add_comb_row(name="基本组合", ctype="包络", cases=envelope_basic_cases)

        self.refresh_combination_cache()

    @staticmethod
    def _generate_valid_span_groups(all_spans, n_gap):
        """生成满足间隔约束的最大跨组合（跳过可合并的子集）

        规则：如果若干跨可同时出现在一个有效组合中，只保留最大组合，
        不生成其子集。孤立的跨（无法与其他跨搭配）仍单独保留。

        示例（3跨 {1,2,3}, n=1）：
          有效组合: {1,3}, {1}, {2}, {3}
          {1}⊂{1,3} 且 {3}⊂{1,3} → 跳过
          {2} 无法与任何跨搭配 → 保留
          最终: [{1,3}, {2}]
        """
        spans = sorted(all_spans)
        n = len(spans)

        # 第一步：枚举所有有效组合
        all_valid = []

        def _backtrack(start, current):
            if current:
                all_valid.append(frozenset(current))
            for i in range(start, n):
                if not current or spans[i] - current[-1] > n_gap:
                    current.append(spans[i])
                    _backtrack(i + 1, current)
                    current.pop()

        _backtrack(0, [])

        if not all_valid:
            return []

        # 第二步：仅保留最大组合（不被其他组合包含的）
        maximal = []
        for group in all_valid:
            is_subset = False
            for other in all_valid:
                if group < other:  # 真子集
                    is_subset = True
                    break
            if not is_subset:
                maximal.append(sorted(group))

        return maximal

    @staticmethod
    def _get_vehicle_weight(vinfo):
        """从车辆信息中提取自重（第一个数值参数）"""
        params = vinfo.get("vehicle_params", [])
        for p in params:
            try:
                return float(p)
            except (ValueError, TypeError):
                continue
        return 0.0

    # ========================== 主表 (荷载组合) 操作 ==========================
    def add_comb_row(self, name=None, ctype=None, cases=None):
        i = len(self.comb_data)
        r = i + 1
        # cache 格式：{comb_name: {"type": str, "cases": {...}}}，按插入序号取默认名/类型
        _combs = list(trestle_excel_dict.get("load_combination", {}).items())
        comb_name = name if name else (_combs[i][0] if i < len(_combs) else f"CB{r}")
        if ctype is None:
            idx = len(self.comb_data)
            ctype = _combs[idx][1].get("type", "相加") if idx < len(_combs) else "相加"
        comb_cases = copy.deepcopy(cases) if cases is not None else []
        
        self.comb_data.append({
            "name": comb_name,
            "type": ctype,
            "cases": comb_cases
        })

        lbl_num = tk.Label(self.top_inner, text=str(r), relief="solid", borderwidth=1, padx=5, bg="white", font=("微软雅黑", 9))
        lbl_num.grid(row=r, column=0, sticky="nsew")

        name_entry = tk.Entry(self.top_inner, relief="solid", borderwidth=1, bg="white", font=("微软雅黑", 9), justify="center")
        name_entry.grid(row=r, column=1, sticky="nsew")
        name_entry.insert(0, comb_name)

        type_frame = tk.Frame(self.top_inner, relief="solid", borderwidth=1, bg="white")
        type_frame.grid(row=r, column=2, sticky="nsew")
        combo_var = tk.StringVar(value=ctype)
        cb = ttk.Combobox(type_frame, textvariable=combo_var, values=["相加", "包络"], state="readonly")
        cb.pack(padx=2, pady=4, fill="both", expand=True)

        # 绑定事件
        cells = [lbl_num, name_entry, type_frame]
        for cell in cells:
            cell.bind("<Button-1>", lambda e, idx=i: self.select_comb(idx), add="+")
            for sub in cell.winfo_children():
                sub.bind("<Button-1>", lambda e, idx=i: self.select_comb(idx), add="+")
                if isinstance(sub, ttk.Combobox):
                    sub.bind("<FocusIn>", lambda e, idx=i: self.select_comb(idx), add="+")

        name_entry.bind("<FocusOut>", lambda e, idx=i: self.sync_comb_data(idx))
        name_entry.bind("<Return>", lambda e, idx=i: self.sync_comb_data(idx))
        cb.bind("<<ComboboxSelected>>", lambda e, idx=i, _cv=combo_var: self.on_comb_type_change(idx, _cv))

        self.top_canvas.configure(scrollregion=self.top_canvas.bbox("all"))
        self.refresh_combination_cache()
        self.select_comb(i, force=True)

    def sync_comb_data(self, idx):
        if idx >= len(self.comb_data): return
        row_data = self.comb_data[idx]
        for w in self.top_inner.grid_slaves(row=idx+1):
            g = w.grid_info()
            col = int(g["column"])
            if col == 1 and isinstance(w, tk.Entry):
                new_name = w.get()
                row_data["name"] = new_name
                # 同步下方标题
                if self.current_comb_idx == idx:
                    self.bottom_title.config(text=f"荷载工况和系数 [{new_name}]")
        self.refresh_combination_cache()

    def on_comb_type_change(self, idx, cv):
        self.comb_data[idx]["type"] = cv.get()
        self.refresh_combination_cache()
        self.select_comb(idx, force=True)

    def delete_comb_row(self):
        if not self.comb_data: return
        idx = self.current_comb_idx
        if idx < 0 or idx >= len(self.comb_data): return

        # 记录被删除的组合名，用于清理其他组合中的引用
        deleted_name = self.comb_data[idx].get("name", "")

        # 从数据中删除
        self.comb_data.pop(idx)

        # 清理其他组合中引用到该组合的工况项
        for comb in self.comb_data:
            comb["cases"] = [c for c in comb.get("cases", []) if c.get("case_name") != deleted_name]

        # 保存数据，清空后通过 add_comb_row 重建（含 UI）
        saved_data = list(self.comb_data)
        self.comb_data.clear()

        for w in list(self.top_inner.grid_slaves()):
            g = w.grid_info()
            if g and g.get("row") and int(g["row"]) > 0:
                w.destroy()

        for entry in saved_data:
            self.add_comb_row(name=entry.get("name"),
                              ctype=entry.get("type"),
                              cases=entry.get("cases", []))

        self.refresh_combination_cache()
        self.top_canvas.configure(scrollregion=self.top_canvas.bbox("all"))

        new_idx = min(idx, len(self.comb_data) - 1) if self.comb_data else -1
        if new_idx >= 0:
            self.select_comb(new_idx, force=True)
        else:
            self.current_comb_idx = -1
            self.render_bottom_table()
    def copy_comb_row(self):
        """复制选中行：包含名字(-1)、类型、及包含的所有工况深度复制"""
        if not self.comb_data: return
        idx = self.current_comb_idx
        if idx < 0 or idx >= len(self.comb_data): return

        target_data = self.comb_data[idx]
        new_name = target_data["name"] + "-1"
        new_type = target_data["type"]
        new_cases = target_data["cases"] # 会在 add_comb_row 时被 deepcopy

        self.add_comb_row(name=new_name, ctype=new_type, cases=new_cases)

    def select_comb(self, idx, force=False):
        if not self.comb_data or idx < 0 or idx >= len(self.comb_data): return

        need_refresh_bottom = force or self.current_comb_idx != idx
        self.current_comb_idx = idx

        # 高亮主表
        for r in range(1, len(self.comb_data)+1):
            bg = "#e3f2fd" if r-1 == idx else "white"
            for w in self.top_inner.grid_slaves(row=r):
                if isinstance(w, (tk.Label, tk.Entry)):
                    try: w.config(bg=bg)
                    except: pass
                elif isinstance(w, tk.Frame):
                    try: w.config(bg=bg)
                    except: pass
                    for sub in w.winfo_children():
                        if isinstance(sub, tk.Label):
                            try: sub.config(bg=bg)
                            except: pass

        if need_refresh_bottom:
            comb_name = self.comb_data[idx]["name"]
            self.bottom_title.config(text=f"荷载工况和系数 [{comb_name}]")
            self.btn_add_case.config(state="normal")
            self.btn_del_case.config(state="normal")
            self.render_bottom_table()

    # ========================== 从表 (工况和系数) 操作 ==========================
    def render_bottom_table(self):
        """根据当前组合索引重新渲染整个下方表格"""
        # 清理旧数据
        for w in self.bot_inner.grid_slaves():
            g = w.grid_info()
            if g and g.get("row") and int(g["row"]) > 0:
                w.destroy()

        self.current_case_idx = -1

        if self.current_comb_idx < 0:
            self.btn_add_case.config(state="normal")
            self.btn_del_case.config(state="normal")
            self.bottom_title.config(text="荷载工况和系数 (请先选择上方组合)")
            return

        cases = self.comb_data[self.current_comb_idx]["cases"]
        for i, c_data in enumerate(cases):
            gamma = c_data.get("gamma", "1.0") if isinstance(c_data, dict) else "1.0"
            psi = c_data.get("psi", "1.0") if isinstance(c_data, dict) else "1.0"
            # 兼容旧格式：如果只有 factor 没有 gamma/psi
            if "gamma" not in c_data and "factor" in c_data:
                gamma = c_data["factor"]
            if "psi" not in c_data and "factor" in c_data:
                psi = c_data["factor"]
            self.insert_case_row_ui(i, c_data["case_name"], gamma, psi)

        self.bot_canvas.configure(scrollregion=self.bot_canvas.bbox("all"))
        if cases:
            self.select_case(0, force=True)

    def insert_case_row_ui(self, i, case_name, gamma, psi):
        r = i + 1

        lbl_num = tk.Label(self.bot_inner, text=str(r), relief="solid", borderwidth=1, padx=5, bg="white", font=("微软雅黑", 9))
        lbl_num.grid(row=r, column=0, sticky="nsew")

        case_frame = tk.Frame(self.bot_inner, relief="solid", borderwidth=1, bg="white")
        case_frame.grid(row=r, column=1, sticky="nsew")
        combo_var = tk.StringVar(value=case_name)
        cb = ttk.Combobox(case_frame, textvariable=combo_var, values=self.get_load_cases_for_comb(self.current_comb_idx), state="readonly")
        cb.pack(padx=2, pady=4, fill="both", expand=True)

        gamma_entry = tk.Entry(self.bot_inner, relief="solid", borderwidth=1, bg="white", font=("微软雅黑", 9), justify="center")
        gamma_entry.grid(row=r, column=2, sticky="nsew")
        gamma_entry.insert(0, gamma)

        psi_entry = tk.Entry(self.bot_inner, relief="solid", borderwidth=1, bg="white", font=("微软雅黑", 9), justify="center")
        psi_entry.grid(row=r, column=3, sticky="nsew")
        psi_entry.insert(0, psi)

        cells = [lbl_num, case_frame, gamma_entry, psi_entry]
        for cell in cells:
            cell.bind("<Button-1>", lambda e, idx=i: self.select_case(idx), add="+")
            for sub in cell.winfo_children():
                sub.bind("<Button-1>", lambda e, idx=i: self.select_case(idx), add="+")
                if isinstance(sub, ttk.Combobox):
                    sub.bind("<FocusIn>", lambda e, idx=i: self.select_case(idx), add="+")

        cb.bind("<<ComboboxSelected>>", lambda e, idx=i, _cv=combo_var: self.sync_case_data(idx, _cv.get(), None, None))
        gamma_entry.bind("<FocusOut>", lambda e, idx=i, _w=gamma_entry: self.sync_case_data(idx, None, _w.get(), None))
        gamma_entry.bind("<Return>", lambda e, idx=i, _w=gamma_entry: self.sync_case_data(idx, None, _w.get(), None))
        psi_entry.bind("<FocusOut>", lambda e, idx=i, _w=psi_entry: self.sync_case_data(idx, None, None, _w.get()))
        psi_entry.bind("<Return>", lambda e, idx=i, _w=psi_entry: self.sync_case_data(idx, None, None, _w.get()))

    def sync_case_data(self, idx, case_name, gamma, psi):
        if self.current_comb_idx < 0: return
        cases = self.comb_data[self.current_comb_idx]["cases"]
        if idx < 0 or idx >= len(cases): return
        if case_name is not None:
            cases[idx]["case_name"] = case_name
        if gamma is not None:
            try: gamma = str(float(gamma))
            except: pass
            cases[idx]["gamma"] = gamma
        if psi is not None:
            try: psi = str(float(psi))
            except: pass
            cases[idx]["psi"] = psi
        self.refresh_combination_cache()
    def on_vehicle_deleted(self, deleted_names):
        """VehicleTab 删除车辆后：清理涉及该车辆的工况项"""
        changed = False
        for comb in self.comb_data:
            cases = comb.get("cases", [])
            comb["cases"] = [c for c in cases
                if not any(v in c.get("case_name", "") for v in deleted_names)]
            if len(comb["cases"]) != len(cases):
                changed = True
        if changed:
            self.refresh_combination_cache()
            if self.current_comb_idx >= 0 and self.current_comb_idx < len(self.comb_data):
                self.select_comb(self.current_comb_idx, force=True)
            else:
                self.render_bottom_table()

    def on_vehicle_renamed(self, old_name, new_name):
        """VehicleTab 车辆改名后：替换所有组合 case_name 中的旧车辆名"""
        if old_name == new_name:
            return
        changed = False
        for comb in self.comb_data:
            for c in comb.get("cases", []):
                old_case = c.get("case_name", "")
                if old_name in old_case:
                    c["case_name"] = old_case.replace(old_name, new_name)
                    changed = True
        if changed:
            self.refresh_combination_cache()
            if self.current_comb_idx >= 0 and self.current_comb_idx < len(self.comb_data):
                self.select_comb(self.current_comb_idx, force=True)
            else:
                self.render_bottom_table()

    def add_case_row(self):
        if self.current_comb_idx < 0: return
        cases = self.comb_data[self.current_comb_idx]["cases"]

        available = self.get_load_cases_for_comb(self.current_comb_idx)
        default_case = available[0] if available else "未知"
        cases.append({"case_name": default_case, "gamma": "1.0", "psi": "1.0"})

        i = len(cases) - 1
        self.insert_case_row_ui(i, default_case, "1.0", "1.0")
        self.refresh_combination_cache()

        self.bot_canvas.configure(scrollregion=self.bot_canvas.bbox("all"))
        self.select_case(i, force=True)

    def delete_case_row(self):
        if self.current_comb_idx < 0: return
        cases = self.comb_data[self.current_comb_idx]["cases"]
        idx = self.current_case_idx
        if idx < 0 or idx >= len(cases): return

        # 完全重绘下方表以保证编号和绑定的绝对安全
        cases.pop(idx)
        self.refresh_combination_cache()
        self.render_bottom_table()

        new_idx = min(idx, len(cases)-1)
        if new_idx >= 0:
            self.select_case(new_idx, force=True)

    def select_case(self, idx, force=False):
        if self.current_comb_idx < 0: return
        cases = self.comb_data[self.current_comb_idx]["cases"]
        if not cases or idx < 0 or idx >= len(cases): return

        self.current_case_idx = idx

        # 高亮从表
        for r in range(1, len(cases)+1):
            bg = "#e3f2fd" if r-1 == idx else "white"
            for w in self.bot_inner.grid_slaves(row=r):
                if isinstance(w, (tk.Label, tk.Entry)):
                    try: w.config(bg=bg)
                    except: pass
                elif isinstance(w, tk.Frame):
                    try: w.config(bg=bg)
                    except: pass
                    for sub in w.winfo_children():
                        if isinstance(sub, tk.Label):
                            try: sub.config(bg=bg)
                            except: pass

    # ========================== 基础工况收集 ==========================
    def collect_basic_cases(self):
        """收集基础工况（当前为静态列表，后续可从环境荷载 Tab 动态补充）"""
        basic = ["自重", "风荷载", "水流力"]
        # TODO: 可从 Excel 读基础工况列表
        self.available_load_cases = list(basic)

    def collect_ui_cases(self):
        """收集 MoveLoadTab / StaticLoadTab / 已定义组合 中的实时工况"""
        # 未点开过对应 tab 时 cache 为空 → 回退项目数据（cache 格式，结构一致）
        if self.parent_ui:
            move_cases = dict(getattr(self.parent_ui.tab_moveloads, 'moveload_cache', {})) if hasattr(self.parent_ui, 'tab_moveloads') else {}
            static_cases = dict(getattr(self.parent_ui.tab_staticloads, 'staticload_cache', {})) if hasattr(self.parent_ui, 'tab_staticloads') else {}
            if not move_cases:
                move_cases = dict(self.parent_ui.trestle_excel_dict.get("move_load_cases", {}))
            if not static_cases:
                static_cases = dict(self.parent_ui.trestle_excel_dict.get("static_load_cases", {}))
            # 移动荷载工况
            for name in move_cases:
                if name not in self.available_load_cases:
                    self.available_load_cases.append(name)
            # 静载工况
            for name in static_cases:
                if name not in self.available_load_cases:
                    self.available_load_cases.append(name)
        # 已定义的荷载组合不放入全局列表（由 get_load_cases_for_comb 分层处理）

    def get_load_cases_for_comb(self, comb_idx):
        """获取指定组合可选的工况列表（基础 + 移动 + 静载 + 序号更小的组合）"""
        cases = list(self.available_load_cases)  # 基础 + 移动 + 静载
        for i in range(comb_idx):
            if i < len(self.comb_data):
                name = self.comb_data[i]["name"]
                if name not in cases:
                    cases.append(name)
        return cases

    # ========================== 组合缓存 ==========================
    def refresh_combination_cache(self):
        """重建 load_combination_cache 并刷新 available_load_cases"""
        self.load_combination_cache.clear()
        for row in self.comb_data:
            cases_dict = {}
            for c in row.get("cases", []):
                try: g = float(c.get("gamma", 1.0))
                except: g = 1.0
                try: p = float(c.get("psi", 1.0))
                except: p = 1.0
                cases_dict[c.get("case_name", "")] = {"gamma": g, "psi": p}
            self.load_combination_cache[row["name"]] = {
                "type": row.get("type", "相加"),
                "cases": cases_dict
            }
        # 同步更新可选工况列表（基础 + 移动 + 静载 + 已定义组合）
        self.collect_basic_cases()
        self.collect_ui_cases()

    # ========================== 页面切换 ==========================
    def _on_page_shown(self):
        """切换到本页时：刷新实时工况列表 + 底部下拉框"""
        self.collect_basic_cases()
        self.collect_ui_cases()
        # 底部工况表下拉框刷新（重建时调用 get_load_cases_for_comb 分层获取）
        if self.current_comb_idx >= 0 and self.comb_data:
            self.render_bottom_table()