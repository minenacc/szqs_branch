# 1. 标准库
import os
import sys
import copy
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path

# 2. 第三方库
import ttkbootstrap as tb
from ttkbootstrap.constants import *

# 3. 本地模块
sys.path.append(str(Path(__file__).parent.parent.parent))
from General.UIHandle import if_Reg
from General.DataUtils import read_Register, midasdisttolst, _parse_simple_spacing

sys.path.append(str(Path(__file__).parent))
from Trestle_Excel_io_FEM import (get_projects, save_trestle_to_excel, PROJECT_TEMPLATE,
                                   _find_sheet, _build_header_map,
                                   _cell, _safe_str, _norm_sec_name,
                                   get_properties_path, load_section_library, load_material_library,
                                   get_excel_path, make_soil_label, copy_soil_sheet,
                                   delete_project_row, load_project_section_library)
from Trestle_Generate_FEM import mcb_generate
from Trestle_Tab_FEM import (
    SupportsTab, ComponentsTab, SoilParamTab, ConnectionTab,
    EnvironmentLoadTab, VehicleTab, MoveLoadTab, StaticLoadTab,
    LoadCombinationTab
)

trestle_excel_dict: dict = {}


def _first_soil_label(proj_data):
    """从 cache 格式项目数据取土层标签（无则返回 ""）"""
    lbl = str(proj_data.get("soil_label", "") or "").strip()
    return lbl if lbl and lbl != "/" else ""

# ===== 统一配色 =====
SIDEBAR_BG = "#1a1d23"
SIDEBAR_ACTIVE = "#2b6fd4"
SIDEBAR_HOVER = "#1a1d23"
SIDEBAR_IDLE_BG = "#e0ecf8"
SIDEBAR_IDLE_FG = "#343a43"

class TrestleUI:
    """主界面 (侧边栏导航)"""
    def __init__(self, root, Applocation):
        self.root = root
        self.Applocation = Applocation
        self.root.title("栈桥建模自动化程序")
        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        # 按屏幕比例计算窗口大小，带合理上下限
        MIN_W, MAX_W = 900, 1200
        MIN_H, MAX_H = 1000, 1300
        win_w = max(MIN_W, min(MAX_W, int(sw * 0.6)))
        win_h = max(MIN_H, min(MAX_H, int(sh * 0.85)))
        self.root.geometry(f"{win_w}x{win_h}+{max(0, (sw - win_w) // 2)}+{max(0, (sh - win_h) // 2)}")
        self.root.minsize(MIN_W, MIN_H)
        self.root.resizable(True, True)

        # ===== 从 Excel 读取全部项目参数 =====
        # 清除模块级缓存，确保每次打开窗口都从 Excel 重新加载
        # （_switch_project 会原地修改 projects_dict，污染缓存）
        import Trestle_Excel_io_FEM as _excel_mod
        _excel_mod.ALL_PROJECTS = None
        self.excel_path = get_excel_path()
        _proj_result = get_projects(excel_path=self.excel_path)
        self.projects_dict = _proj_result["projects_dict"]
        self.project_groups = _proj_result["project_groups"]
        self._current_proj_key = [_proj_result["last_proj_key"]]
        self._switching_flag = [False]
        self._PROJECT_TEMPLATE = PROJECT_TEMPLATE
        # 数据源只有 projects_dict：Excel 读取时已直接转为 cache 格式，
        # 切换/保存后被 combine_all_summary 更新（全程序仅此一种格式）

        # 当前项目的参数 dict（各 tab 从此读取）
        self.trestle_excel_dict = self.projects_dict[self._current_proj_key[0]]
        self.trestle_excel_dict["_excel_path"] = self.excel_path

        # ===== 预计算（在 tab 创建前注入数据）=====
        self._precompute_ground_line()
        self._precompute_quick_setup()

        # 同步到各模块级变量，供 Tab 类直接访问
        import Trestle_Tab_FEM as _tab_mod
        _tab_mod.trestle_excel_dict = self.trestle_excel_dict
        sys.modules[__name__].trestle_excel_dict = self.trestle_excel_dict
        import Trestle_Widgets_FEM as _widgets_mod
        _widgets_mod.trestle_excel_dict = self.trestle_excel_dict

        # ===== 加载 properties_parameter.xlsx（材质库 + 截面库）=====
        _prop_path = get_properties_path()
        if _prop_path:
            self.trestle_excel_dict["section_library_db"] = load_section_library(_prop_path)
            _steel, _concrete, _rebar = load_material_library(_prop_path)
            self.trestle_excel_dict["steel_dict"] = _steel
            self.trestle_excel_dict["concrete_dict"] = _concrete
            self.trestle_excel_dict["rebar_dict"] = _rebar
            print(f"[properties] 截面库={len(self.trestle_excel_dict['section_library_db'])}条, "
                  f"钢材规范={list(_steel.keys())}, 混凝土规范={list(_concrete.keys())}")
        else:
            self.trestle_excel_dict["section_library_db"] = {}
            self.trestle_excel_dict["steel_dict"] = {}
            self.trestle_excel_dict["concrete_dict"] = {}
            self.trestle_excel_dict["rebar_dict"] = {}
            print("[properties] 警告: 找不到 properties_parameter.xlsx")
        # ===== 预载参数项 =====
        # 截面字典（初始 + UI 操作累积，只增不减）
        self.section_dict = dict(self.trestle_excel_dict.get("section", {}))
        
        # ====== 底部按钮栏 ======
        self.create_bottom_bar()

        # ── 项目管理顶部栏（全宽，位于最上方）──
        top_bar = ttk.Frame(self.root)
        top_bar.pack(side=TOP, fill=X)
        proj_inner = ttk.Frame(top_bar, padding=(15, 10))
        proj_inner.pack(fill=X)
        ttk.Label(proj_inner, text="项目管理", font=("Microsoft YaHei UI", 10, "bold")).pack(side=LEFT, padx=(0, 10))
        self._proj_name_var = tk.StringVar()
        self.proj_name_combo = ttk.Combobox(proj_inner, textvariable=self._proj_name_var,
                                            state="readonly", width=15)
        self.proj_name_combo.pack(side=LEFT, padx=(0, 5))
        ttk.Label(proj_inner, text="栈桥编号", font=("Microsoft YaHei UI", 9)).pack(side=LEFT, padx=(0, 5))
        self._proj_no_var = tk.StringVar()
        self.proj_no_combo = ttk.Combobox(proj_inner, textvariable=self._proj_no_var,
                                          state="readonly")
        self.proj_no_combo.pack(side=LEFT, fill=X, expand=True, padx=(0, 10))
        # 按钮区
        proj_btn_frame = ttk.Frame(proj_inner)
        proj_btn_frame.pack(side=RIGHT, padx=5)
        tb.Button(proj_btn_frame, text="新增", bootstyle=(SUCCESS, OUTLINE), width=10,
                  command=self._add_project).pack(side=LEFT, padx=2)
        tb.Button(proj_btn_frame, text="另存为", bootstyle=(INFO, OUTLINE), width=10,
                  command=self._save_as_project).pack(side=LEFT, padx=2)
        tb.Button(proj_btn_frame, text="保存参数", bootstyle=(PRIMARY, OUTLINE), width=10,
                  command=self.save_params).pack(side=LEFT, padx=2)
        tb.Button(proj_btn_frame, text="删除", bootstyle=(DANGER, OUTLINE), width=10,
                  command=self._delete_project).pack(side=LEFT, padx=2)
        # 下拉事件绑定
        self.proj_name_combo.bind("<<ComboboxSelected>>", lambda _: self._on_project_name_change())
        self.proj_no_combo.bind("<<ComboboxSelected>>", lambda _: self._on_project_no_change())
        # 分隔线
        ttk.Separator(top_bar).pack(fill=X)

        # ====== 整体框架：侧边栏+内容 ======
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill="both", expand=True)

        # 侧边栏
        sidebar = tk.Frame(main_frame, bg=SIDEBAR_BG, width=170)
        sidebar.pack(side=LEFT, fill=Y)
        sidebar.pack_propagate(False)
        sep = ttk.Separator(main_frame, orient=tk.VERTICAL)
        sep.pack(side=LEFT, fill=Y)
        header_lbl = ttk.Label(sidebar, text="参数设置", font=("Microsoft YaHei UI", 10, "bold"),
                               bootstyle=DARK, anchor=W)
        header_lbl.pack(fill=X, padx=16, pady=(18, 8))

        # 内容区
        content_outer = ttk.Frame(main_frame)
        content_outer.pack(side=LEFT, fill=BOTH, expand=True)
        self.content_stack = ttk.Frame(content_outer)
        self.content_stack.pack(fill=BOTH, expand=True, padx=5, pady=5)

        # 页面定义（第一个 tab 同步创建，其余延后分批创建以加速窗口显示）
        self.pages_frames = {}
        self._page_defs = [
            ("🏗️", "下部结构", SupportsTab, "下部结构参数定义"),
            ("🌉", "上部结构", ComponentsTab, "上部结构构件定义"),
            ("🪨", "地形土层", SoilParamTab, "地形导入与土层设置"),
            ("🔗", "连接形式", ConnectionTab, "连接形式参数设置"),
            ("🌊", "环境荷载", EnvironmentLoadTab, "风/水/波浪荷载参数"),
            ("🚛", "车辆参数", VehicleTab, "车辆/设备参数定义"),
            ("↔️", "移动荷载", MoveLoadTab, "移动荷载工况定义"),
            ("⚓", "静载工况", StaticLoadTab, "车辆/设备静载工况定义"),
            ("🧮", "荷载组合", LoadCombinationTab, "荷载组合定义"),
        ]
        # 一次性创建所有 tab
        for _, text, cls, _ in self._page_defs:
            frame = cls(self.content_stack, parent_ui=self)
            self.pages_frames[text] = frame
            self._assign_tab_ref(cls, frame)

        # 侧边栏按钮
        self.nav_items = []
        self.nav_indicators = []
        
        for i, (icon, text, cls, _) in enumerate(self._page_defs):
            btn_frame = tk.Frame(sidebar, bg=SIDEBAR_IDLE_BG, height=52, cursor="hand2")
            btn_frame.pack(fill=X, pady=1)
            btn_frame.pack_propagate(False)
            
            indicator = tk.Frame(btn_frame, bg=SIDEBAR_IDLE_BG, width=4)
            indicator.pack(side=LEFT, fill=Y)
            
            # 【核心核心】：单独给 Emoji 设置 width=3 的死宽度，它怎么变都不会影响右边
            icon_lbl = tk.Label(btn_frame, text=icon, font=("Segoe UI Emoji", 12),
                                bg=SIDEBAR_IDLE_BG, fg=SIDEBAR_IDLE_FG, width=3, anchor="center")
            icon_lbl.pack(side=LEFT, fill=Y, padx=(2, 0))
            
            # 文字 Label
            text_lbl = tk.Label(btn_frame, text=text, font=("Microsoft YaHei UI", 11, "bold"),
                                bg=SIDEBAR_IDLE_BG, fg=SIDEBAR_IDLE_FG, anchor="w")
            text_lbl.pack(side=LEFT, fill=BOTH, expand=True)

            # 绑定事件
            for widget in (btn_frame, icon_lbl, text_lbl):
                widget.bind("<Enter>", lambda e, idx=i: self.on_nav_hover(idx, True))
                widget.bind("<Leave>", lambda e, idx=i: self.on_nav_hover(idx, False))
                widget.bind("<Button-1>", lambda e, idx=i: self.switch_page(idx))
                
            self.nav_items.append((btn_frame, icon_lbl, text_lbl))
            self.nav_indicators.append(indicator)

        # 初始化：快捷设置 + _on_page_shown + 切换到首页
        _qd = self.trestle_excel_dict.get("quick_defaults", {})
        if _qd.get("static"):
            self._apply_static_quick()
        for name, frame in self.pages_frames.items():
            if hasattr(frame, '_on_page_shown'):
                frame._on_page_shown()
        self.switch_page(0)
        self._init_project_selection()
        # 以 combine_all_summary 结果为基线（与关闭时的比对数据源一致）
        _cur_key = self._current_proj_key[0]
        self._force_refresh_all_caches()
        self.projects_dict[_cur_key] = self.combine_all_summary()
        self._capture_saved_snapshot()

        # 全局禁用 Combobox 的滚轮切换（展开下拉列表后的滚轮由 Tcl 内部处理，不受影响）
        self.root.bind_class("TCombobox", "<MouseWheel>", lambda e: "break")

    def _assign_tab_ref(self, cls, frame):
        """将 tab frame 赋值到对应的实例属性"""
        frame.pack_forget()
        if cls == SupportsTab:   self.tab_supports = frame
        elif cls == ComponentsTab: self.tab_components = frame
        elif cls == ConnectionTab: self.tab_connections = frame
        elif cls == SoilParamTab: self.tab_soilparams = frame
        elif cls == EnvironmentLoadTab: self.tab_enloads = frame
        elif cls == VehicleTab: self.tab_vehicles = frame
        elif cls == MoveLoadTab: self.tab_moveloads = frame
        elif cls == StaticLoadTab: self.tab_staticloads = frame
        elif cls == LoadCombinationTab: self.tab_loadcombs = frame


    def on_nav_hover(self, idx, is_hover):
        if getattr(self, "current_page_idx", -1) == idx:
            return  
        # 悬停时：深色背景 + 白色文字；离开时：浅色背景 + 原本深色文字
        bg_color = SIDEBAR_HOVER if is_hover else SIDEBAR_IDLE_BG
        fg_color = "white" if is_hover else SIDEBAR_IDLE_FG
        btn_frame, icon_lbl, text_lbl = self.nav_items[idx]
        btn_frame.config(bg=bg_color)
        icon_lbl.config(bg=bg_color, fg=fg_color)
        text_lbl.config(bg=bg_color, fg=fg_color)

    def switch_page(self, idx):
        self.current_page_idx = idx
        for i, (key, frame) in enumerate(self.pages_frames.items()):
            frame.pack_forget() if i != idx else frame.pack(fill=BOTH, expand=True)
            
        active = self.pages_frames[list(self.pages_frames.keys())[idx]]
        if hasattr(active, '_on_page_shown'):
            active._on_page_shown()
            
        for j, ((btn_frame, icon_lbl, text_lbl), ind) in enumerate(zip(self.nav_items, self.nav_indicators)):
            if j == idx:
                btn_frame.config(bg=SIDEBAR_ACTIVE)
                icon_lbl.config(bg=SIDEBAR_ACTIVE, fg="white")
                text_lbl.config(bg=SIDEBAR_ACTIVE, fg="white")
                ind.config(bg="#60a5fa")
            else:
                btn_frame.config(bg=SIDEBAR_IDLE_BG)
                icon_lbl.config(bg=SIDEBAR_IDLE_BG, fg=SIDEBAR_IDLE_FG)
                text_lbl.config(bg=SIDEBAR_IDLE_BG, fg=SIDEBAR_IDLE_FG)
                ind.config(bg=SIDEBAR_IDLE_BG)

    # ======================== 项目管理 ========================

    def _precompute_ground_line(self):
        """在 tab 创建前，从土层 sheet 读取地面线数据（若"地面线"sheet 无数据）"""
        if self.trestle_excel_dict.get("ground_line"):
            return  # 已有数据（来自"地面线"sheet），无需重复读取
        import re
        sheet_name = _first_soil_label(self.trestle_excel_dict)
        if not sheet_name:
            return
        excel_path = self.excel_path
        if not excel_path or not os.path.exists(excel_path):
            return
        try:
            import openpyxl
            wb = openpyxl.load_workbook(excel_path, data_only=True)
            sheet_name = sheet_name + "地形表"
            if sheet_name not in wb.sheetnames:
                wb.close()
                return
            ws = wb[sheet_name]
            boundaries = []
            for r in range(2, ws.max_row + 1):
                if str(ws.cell(r, 1).value or "").strip() == "计算底边线y坐标":
                    continue
                coords = []
                for c in range(2, ws.max_column + 1):
                    v = ws.cell(r, c).value
                    if v is None or str(v).strip() in ('', '/'):
                        continue
                    m = re.match(r'\(\s*([\d.-]+)\s*,\s*([\d.-]+)\s*\)', str(v))
                    if m:
                        coords.append((float(m.group(1)), float(m.group(2))))
                if len(coords) >= 2:
                    boundaries.append(coords)
            wb.close()
            if not boundaries:
                return
            # 对每个 x，取所有边界线中的最大 y
            all_x = sorted(set(p[0] for pts in boundaries for p in pts))
            ground = []
            for x in all_x:
                valid_ys = []
                for pts in boundaries:
                    sp = sorted(pts, key=lambda p: p[0])
                    xs = [p[0] for p in sp]
                    ys = [p[1] for p in sp]
                    if x < xs[0] or x > xs[-1]:
                        continue
                    for i in range(len(xs) - 1):
                        if xs[i] <= x <= xs[i + 1]:
                            if xs[i] == xs[i + 1]:
                                valid_ys.append(ys[i])
                            else:
                                valid_ys.append(ys[i] + (ys[i + 1] - ys[i]) * (x - xs[i]) / (xs[i + 1] - xs[i]))
                            break
                if valid_ys:
                    ground.append((x, max(valid_ys)))
            if ground:
                self.trestle_excel_dict["ground_line"] = ground
                print(f"[预计算] 地面线: {len(ground)} 个点 (来源: {sheet_name})")
        except Exception as e:
            print(f"[预计算] 地面线读取失败: {e}")

    def _precompute_quick_setup(self):
        """在 tab 创建前，根据快捷设置预计算并注入 trestle_excel_dict"""
        qd = self.trestle_excel_dict.get("quick_defaults", {})
        traffic_name = qd.get("traffic", "")
        static_name = qd.get("static", "")
        if not traffic_name and not static_name:
            return
        # 该项目车辆/荷载数据已存在（快捷设置已生成过或被手动编辑过）→ 跳过再生，防止覆盖手动修改。
        # 快捷模板生成必然先写 vehicle，故 vehicle 非空即可判定无需再生；重新应用走 UI 快捷按钮。
        if self.trestle_excel_dict.get("vehicle"):
            return

        print(f"\n[预计算] 快捷设置: 通行荷载={traffic_name}  施工静载={static_name}")

        # 加载车辆标准库
        from Trestle_Tab_FEM import VehicleTab as _VT
        import Trestle_Tab_FEM as _tab_mod
        _saved = _tab_mod.trestle_excel_dict
        _tab_mod.trestle_excel_dict = self.trestle_excel_dict
        _, veh_info = _VT.load_vehicle_xlsx(None)
        _tab_mod.trestle_excel_dict = _saved

        has_traffic = bool(traffic_name)
        has_static = bool(static_name)

        # ── 1. 目标车辆集 ──
        if has_traffic and has_static:
            target = list(dict.fromkeys([traffic_name, static_name]))
        elif has_traffic:
            target = [traffic_name]
            for _entry in self.trestle_excel_dict.get("static_load_cases", {}).values():
                v = _entry.get("vehicle", "")
                if v and v not in target:
                    target.append(v)
        else:
            target = [static_name]

        print(f"[预计算] 目标车辆: {target}")

        # ── 2. 注入 vehicle（cache 格式 dict）──
        veh_cache = {}
        for vname in target:
            if vname in veh_info:
                vi = veh_info[vname]
                veh_cache[vname] = {"vehicle_var": vname, "vehicle_type": vi["load_type"], "vehicle_params": vi["params"]}
            else:
                veh_cache[vname] = {"vehicle_var": vname, "vehicle_type": "", "vehicle_params": []}
        self.trestle_excel_dict["vehicle"] = veh_cache
        print(f"[预计算] vehicle: {list(veh_cache.keys())}")

        # ── 3. 注入 move_load_cases（cache 格式 dict）──
        if has_traffic:
            case_name = f"{traffic_name}走行"
            self.trestle_excel_dict["move_load_cases"] = {
                case_name: {"vehicle": traffic_name, "lane_ecc": 0.0, "direction": "往返", "factor": 1.0}
            }
            print(f"[预计算] move_load_cases: [{case_name}]")

        # ── 4. 注入 static_load_cases（cache 格式 dict，span/pile 在 _on_page_shown 时完善）──
        if has_static:
            static_cache = {}
            for _t in ("正吊(正向)", "侧吊"):
                static_cache[f"{static_name}{_t}"] = {
                    "vehicle": static_name, "type": _t,
                    "跨间左侧": [], "跨中": [], "跨间右侧": [], "桩顶": [],
                }
            self.trestle_excel_dict["static_load_cases"] = static_cache
            print(f"[预计算] static_load_cases: {list(static_cache.keys())}")

    def _apply_static_quick(self):
        """静载 tab 快捷设置（延后执行，依赖 SupportsTab 的 pile_count）"""
        qd = self.trestle_excel_dict.get("quick_defaults", {})
        static_name = qd.get("static", "")
        if not static_name or not hasattr(self, 'tab_staticloads'):
            return
        slt = self.tab_staticloads
        slt._on_page_shown()
        for vname in slt.car_dict:
            slt.car_output_check[vname] = (vname == static_name)
        if static_name in slt.car_dict:
            data = slt.car_dict[static_name]
            for lt in data.get("load_types", {}):
                data["load_types"][lt] = (lt in ("正吊(正向)", "侧吊"))
            for row in data.get("span_checks", []):
                if len(row) >= 3:
                    row[0] = False; row[1] = True; row[2] = False
            for j in range(len(data.get("pile_checks", []))):
                data["pile_checks"][j] = True
            slt._rebuild_loading_tables()
            slt._current_car_name = static_name
            for lt, var in slt.load_type_vars.items():
                var.set(lt in ("正吊(正向)", "侧吊"))
            for row_vars in slt._span_chk_vars:
                if len(row_vars) >= 3:
                    row_vars[0].set(False); row_vars[1].set(True); row_vars[2].set(False)
            for var in slt._pile_chk_vars:
                var.set(True)
            slt._save_current_state()
        slt.refresh_cache()
        # 荷载组合
        if hasattr(self, 'tab_loadcombs'):
            self.tab_loadcombs._auto_combine()

    def _init_project_selection(self):
        """初始化项目下拉列表，选中 last_proj_key"""
        self._refresh_project_combos()
        key = self._current_proj_key[0]
        if key in self.projects_dict:
            pname = self.projects_dict[key].get("project_name", "")
            pno = self.projects_dict[key].get("project_no", "")
            if pname:
                self._proj_name_var.set(pname)
                self._refresh_no_combo(pname)
            if pno:
                self._proj_no_var.set(pno)

    def _refresh_project_combos(self):
        """刷新项目名称下拉列表"""
        names = sorted(self.project_groups.keys())
        self.proj_name_combo["values"] = names

    def _refresh_no_combo(self, pname):
        """刷新编号下拉列表（按项目名称过滤）"""
        nos = self.project_groups.get(pname, [])
        self.proj_no_combo["values"] = nos

    def _on_project_name_change(self):
        """项目名称下拉变更：刷新编号列表 → 切换到第一个编号"""
        pname = self._proj_name_var.get()
        if not pname:
            return
        self._refresh_no_combo(pname)
        nos = self.project_groups.get(pname, [])
        if nos:
            self._proj_no_var.set(nos[0])
            self._switch_project(pname + nos[0])

    def _on_project_no_change(self):
        """编号下拉变更：切换到对应项目"""
        pname = self._proj_name_var.get()
        pno = self._proj_no_var.get()
        if pname and pno:
            self._switch_project(pname + pno)

    def _switch_project(self, proj_key, skip_save=False):
        """切换项目：保存旧项目 → 加载新项目到 UI"""
        if self._switching_flag[0]:
            return
        if proj_key not in self.projects_dict:
            return
        if proj_key == self._current_proj_key[0]:
            return

        self._switching_flag[0] = True
        try:
            old_key = self._current_proj_key[0]
            # 保存旧项目的 UI 状态（cache 格式，供 Excel/MCT 使用）；删除项目后跳过保存
            if not skip_save and old_key in self.projects_dict:
                self._force_refresh_all_caches()
                old_summary = self.combine_all_summary()
                self.projects_dict[old_key] = old_summary

            # 切换到新项目
            self._current_proj_key[0] = proj_key
            # 数据源为 projects_dict（统一 cache 格式）：切换/保存过的项目含会话内修改。
            # 不能用 Excel 快照：否则会话内修改（静载/组合等）在切回时全部丢失
            proj_data = self.projects_dict[proj_key]

            # 更新模块级 trestle_excel_dict（供各 tab 的 __init__/_on_page_shown 读取）
            # 保留 properties 数据（截面库/材质库），这些不随项目变化
            _PROPS_KEYS = ("section_library_db", "steel_dict", "concrete_dict", "rebar_dict")
            old_props = {k: self.trestle_excel_dict.get(k) for k in _PROPS_KEYS}
            self.trestle_excel_dict = proj_data
            self.trestle_excel_dict["_excel_path"] = self.excel_path
            for k, v in old_props.items():
                if v is not None and k not in proj_data:
                    proj_data[k] = v
            import Trestle_Tab_FEM as _tab_mod
            _tab_mod.trestle_excel_dict = proj_data
            sys.modules[__name__].trestle_excel_dict = proj_data
            import Trestle_Widgets_FEM as _widgets_mod
            _widgets_mod.trestle_excel_dict = proj_data

            # 更新截面字典
            self.section_dict = dict(proj_data.get("section", {}))

            # 预计算（在 tab 刷新前注入数据）
            self._precompute_ground_line()
            # 无土层数据的项目不应有地面线（防止 TEMPLATE 项目被意外污染）
            if not _first_soil_label(proj_data):
                proj_data.pop("ground_line", None)
            self._precompute_quick_setup()

            # 数据绑定更新：逐个 tab 调用 reload_from_dict（传入当前项目数据）
            for tab in self.pages_frames.values():
                if hasattr(tab, "reload_from_dict"):
                    tab.reload_from_dict(proj_data)

            # 更新下拉列表选中项
            pname = proj_data.get("project_name", "")
            pno = proj_data.get("project_no", "")
            self._proj_name_var.set(pname)
            self._refresh_no_combo(pname)
            self._proj_no_var.set(pno)

            print(f"[项目切换] {old_key} → {proj_key}")
        finally:
            self._switching_flag[0] = False

    def _rebuild_all_tabs(self):
        """销毁并重建所有 tab（用于项目切换后刷新 UI）"""
        # 保存当前页面索引
        current_idx = getattr(self, 'current_page_idx', 0)

        # 销毁所有 tab
        for frame in self.pages_frames.values():
            frame.destroy()
        self.pages_frames.clear()

        # 重建所有 tab
        for _, text, cls, _ in self._page_defs:
            frame = cls(self.content_stack, parent_ui=self)
            self.pages_frames[text] = frame
            self._assign_tab_ref(cls, frame)

        # 静载快捷设置（依赖 SupportsTab 的 pile_count，需同步执行）
        _qd = self.trestle_excel_dict.get("quick_defaults", {})
        if _qd.get("static"):
            self._apply_static_quick()

        # 恢复到当前页面
        self.switch_page(min(current_idx, len(self.pages_frames) - 1))

    def _add_project(self, template_data=None):
        """新增项目 / 另存为（template_data=None 为新增，否则为另存为）"""
        is_save_as = template_data is not None
        dialog = tb.Toplevel(self.root)
        dialog.title("另存为" if is_save_as else "新增项目")
        dialog.geometry("380x250+%d+%d" % (self.root.winfo_rootx()+80, self.root.winfo_rooty()+80))
        dialog.transient(self.root)
        dialog.grab_set()

        # 项目名称：可编辑 Combobox（列出已有名称供选择）
        tb.Label(dialog, text="项目名称:", width=12).pack(padx=10, pady=(15, 5), anchor=W)
        existing_names = sorted(self.project_groups.keys())
        current_name = self._proj_name_var.get()
        preset_name = template_data.get("project_name", current_name) if is_save_as else current_name
        name_var = tb.StringVar(value=preset_name)
        name_combo = tb.Combobox(dialog, textvariable=name_var, values=existing_names, width=28)
        name_combo.pack(padx=10, pady=5, fill=X)

        # 编号：手动输入
        tb.Label(dialog, text="栈桥编号:", width=12).pack(padx=10, pady=5, anchor=W)
        current_no = self._proj_no_var.get()
        preset_no = template_data.get("project_no", current_no) if is_save_as else ""
        no_var = tb.StringVar(value=preset_no)
        no_entry = tb.Entry(dialog, textvariable=no_var, width=28)
        no_entry.pack(padx=10, pady=5, fill=X)

        def confirm():
            raw_name = name_var.get().strip()
            raw_no = no_var.get().strip()
            if not raw_name:
                messagebox.showwarning("提示", "项目名称不能为空", parent=dialog)
                return
            if not raw_no:
                messagebox.showwarning("提示", "栈桥编号不能为空", parent=dialog)
                return
            combined_key = raw_name + raw_no
            if combined_key in self.projects_dict:
                messagebox.showwarning("提示", f"项目 '{combined_key}' 已存在", parent=dialog)
                return
            import copy
            if is_save_as:
                # 从当前项目最新数据复制（统一 cache 格式，含本会话修改）
                old_key = self._current_proj_key[0]
                new_proj = copy.deepcopy(self.projects_dict.get(old_key, self._PROJECT_TEMPLATE))
            else:
                new_proj = copy.deepcopy(self._PROJECT_TEMPLATE)
            new_proj["project_name"] = raw_name
            new_proj["project_no"] = raw_no
            # 每个项目唯一 soil_label（命名空间前缀 {项目名}{编号}，派生 {label}地形表 / {label}土层表）；新增/另存为即写
            new_proj["soil_label"] = make_soil_label(raw_name, raw_no)
            # 另存为：复制源项目地形表 + 土层表到新 label
            if is_save_as:
                old_label = self.projects_dict[old_key].get("soil_label", "")
                if old_label:
                    copy_soil_sheet(self.excel_path, old_label, new_proj["soil_label"])
            # 新增项目：从 properties 填充空材质字段和截面详情
            if not is_save_as:
                self._fill_default_materials(new_proj)
                self._fill_default_sections(new_proj)
            self.projects_dict[combined_key] = new_proj
            self.project_groups.setdefault(raw_name, [])
            if raw_no not in self.project_groups[raw_name]:
                self.project_groups[raw_name].append(raw_no)
            self._refresh_project_combos()
            self._switch_project(combined_key)
            # 立即写入 Excel（新增/另存为项目实时持久化）
            try:
                mct_path = self.save_path_var.get() if hasattr(self, 'save_path_var') else ""
                save_path, template_ok, _ = save_trestle_to_excel(self.projects_dict, mct_path)
                if not save_path:
                    messagebox.showwarning("提示", "项目已创建，但参数表写入失败（文件可能被占用），请关闭后重试。", parent=dialog)
                elif not template_ok:
                    messagebox.showwarning("提示", "项目已创建，但模板文件被占用，未能同步更新，请关闭模板文件后重试。", parent=dialog)
            except Exception:
                messagebox.showwarning("提示", "项目已创建，但参数表写入异常，请关闭后重试。", parent=dialog)
            dialog.destroy()

        btn_frame = tb.Frame(dialog)
        btn_frame.pack(fill=X, padx=10, pady=10)
        tb.Button(btn_frame, text="确定", bootstyle=SUCCESS, command=confirm).pack(side=RIGHT, padx=5)
        tb.Button(btn_frame, text="取消", bootstyle=SECONDARY, command=dialog.destroy).pack(side=RIGHT, padx=5)
        name_combo.focus()
        dialog.wait_window()

    def _fill_default_materials(self, proj):
        """从 properties_parameter.xlsx 填充空材质字段（规范编号,牌号）"""
        steel_dict = trestle_excel_dict.get("steel_dict", {})
        if not steel_dict:
            return
        # 取第一本规范的编号 + 第一个牌号
        first_spec = next(iter(steel_dict.values()))
        spec_id = first_spec.get("钢结构规范编号", "")
        brands = list(first_spec.get("牌号参数", {}).keys())
        default_brand = brands[0] if brands else "Q235"
        default_val = f"{spec_id},{default_brand}" if spec_id else default_brand

        def _fill(val):
            return default_val if not val or val.strip() == "" else val

        # 下部结构
        for sup in proj.get("substructure", {}).values():
            for k in ("pile_material", "dist_trans_material", "dist_long_material", "brace_material"):
                if k in sup:
                    sup[k] = _fill(sup[k])
        # 上部结构
        for comp in proj.get("components", {}).values():
            for k in ("deck_material", "rib_trans_material", "rib_long_material", "beam_material"):
                if k in comp:
                    comp[k] = _fill(comp[k])

    def _fill_default_sections(self, proj):
        """从截面库填充截面详情（section_type, params）"""
        section_lib = trestle_excel_dict.get("section_library_db", {})
        if not section_lib:
            return

        def _find_section(name):
            """从截面库查找截面详情"""
            if not name or name in ("", "/"):
                return {}
            norm = _norm_sec_name(name)
            for lib_name, lib_info in section_lib.items():
                if _norm_sec_name(lib_name) == norm:
                    return {
                        "section_type": lib_info.get("type", ""),
                        "section_name": lib_name,
                        "params": lib_info.get("params", {}),
                    }
            return {}

        # 下部结构
        for sup in proj.get("substructure", {}).values():
            for k in ("pile_sec", "dist_trans_sec", "dist_long_sec", "brace_sec", "brace_sec_tilt"):
                if k in sup and sup[k] and sup[k] != "/":
                    detail = _find_section(sup[k])
                    if detail:
                        sup[k + "_detail"] = detail
        # 上部结构
        for comp in proj.get("components", {}).values():
            for k in ("rib_trans_sec", "rib_long_sec", "beam_sec"):
                if k in comp and comp[k] and comp[k] != "/":
                    detail = _find_section(comp[k])
                    if detail:
                        comp[k + "_detail"] = detail

    def _save_as_project(self):
        """另存为：以当前项目为模板新增一份副本"""
        old_key = self._current_proj_key[0]
        if not old_key or old_key not in self.projects_dict:
            return
        self._add_project(template_data=self.projects_dict[old_key])

    def _delete_project(self):
        """删除当前项目"""
        if len(self.projects_dict) <= 1:
            messagebox.showwarning("提示", "至少保留一个项目", parent=self.root)
            return
        key = self._current_proj_key[0]
        if not messagebox.askyesno("确认删除", f"确定删除项目 '{key}' 吗？", parent=self.root):
            return

        pname = self.projects_dict[key].get("project_name", "")
        pno = self.projects_dict[key].get("project_no", "")

        # 记录删除项在项目列表中的位置（用于跳转到上一条）
        keys = list(self.projects_dict.keys())
        try:
            del_idx = keys.index(key)
        except ValueError:
            del_idx = 0

        del self.projects_dict[key]
        if pname in self.project_groups and pno in self.project_groups[pname]:
            self.project_groups[pname].remove(pno)
            if not self.project_groups[pname]:
                del self.project_groups[pname]

        # 立即从 Excel 删除对应行
        delete_project_row(pname, pno, self.excel_path)

        # 切换到删除项的上一条项目（如果删除的是第一条，则切换到新的第一条）
        remaining_keys = list(self.projects_dict.keys())
        if del_idx > 0 and del_idx - 1 < len(remaining_keys):
            next_key = remaining_keys[del_idx - 1]
        else:
            next_key = remaining_keys[0]
        self._refresh_project_combos()
        self._switch_project(next_key, skip_save=True)

    def choosesave_path(self):
        from tkinter import filedialog
        from General.FilePath import suppress_dialog_stderr
        with suppress_dialog_stderr():
            filename = filedialog.asksaveasfilename(
                title="保存文件",
                initialdir=os.getcwd(),
                defaultextension=".mct",
                filetypes=[("MCT文件", "*.mct")]
            )
        if filename:
            self.save_path_var.set(filename)
            print("保存路径更新为：", filename)

    def _force_refresh_all_caches(self):
        """在生成 MCT 或保存前，强制刷新所有 Tab 的 cache"""
        for attr in ('tab_supports', 'tab_components', 'tab_soilparams',
                      'tab_connections', 'tab_enloads', 'tab_vehicles',
                      'tab_moveloads', 'tab_staticloads', 'tab_loadcombs'):
            tab = getattr(self, attr, None)
            if tab and hasattr(tab, 'refresh_cache'):
                tab.refresh_cache()

    def _capture_saved_snapshot(self):
        """记录当前 projects_dict 的深拷贝，作为「上次保存」基线"""
        import copy
        self._saved_snapshot = copy.deepcopy(self.projects_dict)

    @staticmethod
    def _dicts_equal(a, b):
        """递归比较两个 dict/list/primitive 结构是否完全相等"""
        if type(a) is not type(b):
            return False
        if isinstance(a, dict):
            if set(a.keys()) != set(b.keys()):
                return False
            return all(TrestleUI._dicts_equal(a[k], b[k]) for k in a)
        if isinstance(a, (list, tuple)):
            if len(a) != len(b):
                return False
            return all(TrestleUI._dicts_equal(x, y) for x, y in zip(a, b))
        try:
            return a == b
        except Exception:
            return a is b

    def _on_closing(self):
        """关闭确认：对比当前数据与上次保存基线，未保存则弹窗提示"""
        self._force_refresh_all_caches()
        current = self.combine_all_summary()
        self.projects_dict[self._current_proj_key[0]] = current
        if not self._dicts_equal(self.projects_dict, self._saved_snapshot):
            result = messagebox.askyesnocancel(
                "未保存的修改",
                "当前项目有未保存的修改，是否保存后再关闭？",
                icon="warning",
                parent=self.root,
            )
            if result is True:
                if self.save_params():
                    self.root.destroy()
                # save_params 失败时不销毁（已弹错误提示，用户可重试或手动关闭）
            elif result is False:
                self.root.destroy()
            else:
                return
        else:
            self.root.destroy()

    # ── 数据验证辅助函数 ──

    @staticmethod
    def _is_numeric(val) -> bool:
        if val is None:
            return False
        try:
            float(val)
            return True
        except (ValueError, TypeError):
            return False

    @staticmethod
    def _is_dist_format(val) -> bool:
        if not val or not isinstance(val, str) or not val.strip():
            return False
        try:
            from General.DataUtils import midasdisttolst
            result = midasdisttolst(val.strip())
            return bool(result)
        except Exception:
            return False

    def _validate_project(self, proj: dict) -> list[str]:
        errors = []
        errors += self._validate_basic(proj.get("basic", {}))
        errors += self._validate_components(proj.get("components", {}))
        errors += self._validate_environment(proj.get("environment", {}))
        errors += self._validate_vehicles(proj.get("vehicle", {}))
        errors += self._validate_sections(proj.get("section", {}))
        errors += self._validate_load_cases(proj)
        errors += self._validate_loadcombs(proj)
        errors += self._validate_span_for_beam_type(proj)
        return errors

    @staticmethod
    def _validate_span_for_beam_type(proj: dict) -> list[str]:
        """桁架栈桥（321型贝雷梁等）每联总长须为3m倍数"""
        errors = []
        span_str = proj.get("basic", {}).get("span", "")
        if not span_str:
            return errors
        # 查 truss_beam_type
        truss_beam_type = ""
        for comp in proj.get("components", {}).values():
            if isinstance(comp, dict):
                truss_beam_type = comp.get("truss_beam_type", "")
                break
        if not truss_beam_type:
            return errors
        try:
            from General.DataUtils import midasdisttolst
            vals = midasdisttolst(span_str.strip())
        except Exception:
            return errors
        if not vals:
            return errors
        # 解析伸出距离（mm → m）
        beam_cant_str = "0,0"
        for comp in proj.get("components", {}).values():
            if isinstance(comp, dict):
                beam_cant_str = comp.get("beam_cant", "0,0")
                break
        cant_parts = [x.strip() for x in str(beam_cant_str).split(",")]
        try:
            cant_left_m = float(cant_parts[0]) / 1000.0 if cant_parts[0] else 0.0
        except (ValueError, TypeError):
            cant_left_m = 0.0
        try:
            cant_right_m = float(cant_parts[1]) / 1000.0 if len(cant_parts) > 1 and cant_parts[1] else 0.0
        except (ValueError, TypeError):
            cant_right_m = 0.0
        # 计算每联总长（含伸出）
        # 首联：第一跨 + 左伸出
        # 尾联：最后一跨 + 右伸出
        # 中间联：跨径
        sub = proj.get("substructure", {})
        sub_types = [sub[k].get("type", "") for k in sorted(sub.keys(), key=lambda k: int(k) if k.isdigit() else 0)]
        # 制动墩位置将联分成多段
        zdd_indices = [i for i, t in enumerate(sub_types) if t == "制动墩"]
        # 构建各联的跨径列表
        spans_per_span = []
        if not zdd_indices:
            # 无制动墩：所有跨为一联
            spans_per_span.append(vals)
        else:
            # 有制动墩：按制动墩分联
            prev_idx = 0
            for zdd_i in zdd_indices:
                spans_per_span.append(vals[prev_idx:zdd_i])
                prev_idx = zdd_i
            spans_per_span.append(vals[prev_idx:])
        # 验证每联总长
        for i, span_group in enumerate(spans_per_span):
            if not span_group:
                continue
            total = sum(span_group)
            # 首联加左伸出，尾联加右伸出
            if i == 0:
                total += cant_left_m
            if i == len(spans_per_span) - 1:
                total += cant_right_m
            if total > 0 and abs(round(total, 6) % 3) > 1e-4:
                errors.append(f"[基本参数] 第{i+1}联总长{total:.3f}m，须为3m倍数（贝雷梁每片3m）")
        return errors

    @staticmethod
    def _validate_sections(section: dict) -> list[str]:
        errors = []
        for sec_name, sec_data in section.items():
            if not isinstance(sec_data, dict):
                continue
            sec_type = sec_data.get("type", "")
            params = sec_data.get("params", {})
            if not sec_type:
                errors.append(f"[截面] {sec_name}: 截面类型为空")
                continue
            if not params:
                errors.append(f"[截面] {sec_name}: 截面参数为空")
                continue
            for pname, pval in params.items():
                # 跳过不适用的参数（值为"/"表示该参数不适用于此截面类型）
                if str(pval).strip() == "/":
                    continue
                if pval is None or str(pval).strip() == "":
                    errors.append(f"[截面] {sec_name}: 参数 {pname} 为空")
                elif not TrestleUI._is_numeric(pval):
                    errors.append(f"[截面] {sec_name}: 参数 {pname} 非数值: {pval}")
        return errors

    @staticmethod
    def _validate_basic(basic: dict) -> list[str]:
        errors = []
        span = basic.get("span", "")
        if not span or not isinstance(span, str) or not span.strip():
            errors.append("[基本参数] 计算跨径不能为空")
        else:
            try:
                from General.DataUtils import midasdisttolst
                vals = midasdisttolst(span.strip())
                if not vals:
                    errors.append(f"[基本参数] 计算跨径格式无法解析: {span}")
            except Exception:
                errors.append(f"[基本参数] 计算跨径格式错误: {span}")
        bw = basic.get("bridge_width", "")
        if not bw and bw != 0:
            errors.append("[基本参数] 计算宽度不能为空")
        elif bw and not TrestleUI._is_numeric(bw):
            errors.append(f"[基本参数] 计算宽度必须为数值: {bw}")
        wl = basic.get("water_level", "")
        if wl is not None and wl != "":
            if not TrestleUI._is_numeric(wl):
                errors.append(f"[基本参数] 设防水位必须为数值: {wl}")
        dl = basic.get("deck_level", "")
        if dl is not None and dl != "":
            if not TrestleUI._is_numeric(dl):
                errors.append(f"[基本参数] 桥面高程必须为数值: {dl}")
        if wl is not None and wl != "" and dl is not None and dl != "":
            if TrestleUI._is_numeric(wl) and TrestleUI._is_numeric(dl):
                if float(dl) <= float(wl):
                    errors.append(f"[基本参数] 桥面高程({dl})须高于设防水位({wl})")
        return errors

    def _validate_components(self, components: dict) -> list[str]:
        errors = []
        for bridge_type, comp in components.items():
            if not isinstance(comp, dict):
                continue
            deck_type = comp.get("deck_type", "单层钢面板")

            # 横肋间距 + 纵梁横向布置（所有桥面类型都需要）
            for key, label in (("rib_trans_space", "横肋间距"), ("beam_space", "纵梁横向布置")):
                val = comp.get(key, "")
                if not val or not isinstance(val, str) or not val.strip():
                    errors.append(f"[上部结构] {label}不能为空")
                elif not self._is_dist_format(val):
                    errors.append(f"[上部结构] {label}格式无法解析: {val}")

            # 数值字段
            for key, label in (("deck_thickness", "桥面板厚"), ("beam_first_space", "第一排间距")):
                val = comp.get(key, "")
                if not val and val != 0:
                    errors.append(f"[上部结构] {label}不能为空")
                elif val and not self._is_numeric(val):
                    errors.append(f"[上部结构] {label}必须为数值: {val}")

            # 双层钢面板：纵肋截面 + 纵肋横向布置
            if deck_type == "双层钢面板":
                rib_long_sec = comp.get("rib_long_sec", "")
                rib_long_sec_detail = comp.get("rib_long_sec_detail", {})
                if not rib_long_sec or rib_long_sec == "/" or not rib_long_sec_detail:
                    errors.append("[上部结构] 双层钢面板纵肋截面不能为空，请设置截面")
                rib_long_space = comp.get("rib_long_space", "")
                if not rib_long_space or not isinstance(rib_long_space, str) or not rib_long_space.strip():
                    errors.append("[上部结构] 双层钢面板纵肋横向布置不能为空")
                elif not self._is_dist_format(rib_long_space):
                    errors.append(f"[上部结构] 双层钢面板纵肋横向布置格式无法解析: {rib_long_space}")

            # 桥面板分节段
            if comp.get("seg_enabled") and not comp.get("panel_length"):
                errors.append("[上部结构] 桥面板分节段启用时，节段长度不能为空")

            # 伸出距离
            for key, label in (("beam_cant", "纵梁两端伸出距离"), ("deck_end_ext", "桥面纵向伸出距离")):
                val = comp.get(key, "")
                if val:
                    parts = [x.strip() for x in str(val).split(",") if x.strip()]
                    if len(parts) != 2:
                        errors.append(f"[上部结构] {label}需两个数值(逗号分隔)，当前: {val}")
                    else:
                        for i, p in enumerate(parts):
                            if not self._is_numeric(p):
                                errors.append(f"[上部结构] {label}第{i+1}项非数值: {p}")

            # 材质字段：验证格式完整性（规范,牌号），按桥型跳过不相关字段
            is_truss = ("桁架" in bridge_type)
            mat_checks = [("deck_material", "桥面板材质"), ("rib_trans_material", "横肋材质")]
            if not is_truss:
                mat_checks.append(("beam_material", "纵梁材质"))
            if deck_type == "双层钢面板":
                mat_checks.append(("rib_long_material", "纵肋材质"))
            for mat_key, label in mat_checks:
                mat_val = comp.get(mat_key, "")
                if mat_val:
                    parts = [x.strip() for x in mat_val.split(",", 1)]
                    if len(parts) < 2 or not parts[1]:
                        errors.append(f"[上部结构] {label}格式不完整(需'规范,牌号'): {mat_val}")

            # 型钢栈桥：验证纵梁截面不为空
            if not is_truss:
                beam_sec = comp.get("beam_sec", "")
                beam_sec_detail = comp.get("beam_sec_detail", {})
                if not beam_sec or beam_sec == "/" or not beam_sec_detail:
                    errors.append("[上部结构] 型钢栈桥纵梁截面不能为空，请设置截面")
        return errors

    @staticmethod
    def _validate_environment(env: dict) -> list[str]:
        errors = []
        wind = env.get("wind", {})
        for key, label in (("design_wind_speed", "设计风速"), ("terrain_factor", "地形条件系数"),
                            ("girder_height", "主梁基准高度"), ("transverse_coeff", "横向力系数")):
            val = wind.get(key)
            if val is None or val == "":
                errors.append(f"[环境荷载] {label}不能为空")
            elif not TrestleUI._is_numeric(val):
                errors.append(f"[环境荷载] {label}必须为数值: {val}")
        water = env.get("water", {})
        val = water.get("flow_velocity")
        if val is None or val == "":
            errors.append("[环境荷载] 设计流速不能为空")
        elif not TrestleUI._is_numeric(val):
            errors.append(f"[环境荷载] 设计流速必须为数值: {val}")
        # 波浪力：激活时验证所有参数均为数值
        wave = env.get("wave", {})
        wave_params = ["wave_period", "wave_length", "wave_height"]
        wave_labels = {"wave_period": "波浪周期", "wave_length": "设计波长", "wave_height": "设计波高"}
        # 判断波浪力是否激活（任一参数不为空且不为"/"）
        wave_active = any(
            wave.get(k) and str(wave.get(k)).strip() not in ("", "/")
            for k in wave_params
        )
        if wave_active:
            for k in wave_params:
                val = wave.get(k)
                if val is None or str(val).strip() == "" or str(val).strip() == "/":
                    errors.append(f"[环境荷载] 波浪力已激活，{wave_labels[k]}不能为空")
                elif not TrestleUI._is_numeric(val):
                    errors.append(f"[环境荷载] {wave_labels[k]}必须为数值: {val}")
        return errors

    @staticmethod
    def _validate_vehicles(vehicle: dict) -> list[str]:
        errors = []
        seen = {}
        for veh_name, veh_data in vehicle.items():
            if not isinstance(veh_data, dict):
                continue
            params = veh_data.get("vehicle_params", [])
            vtype = veh_data.get("vehicle_type", "")
            # 一般车辆：vehicle_params[1]=轴重CSV，[2]=距前轴CSV, 需验证逗号分隔+数量对应
            if "一般车辆" in vtype:
                if len(params) >= 3:
                    axle_dist = str(params[2] or "")
                    axle_weight = str(params[1] or "")
                    dist_vals = [x.strip() for x in axle_dist.split(",") if x.strip()]
                    weight_vals = [x.strip() for x in axle_weight.split(",") if x.strip()]
                    for i, v in enumerate(dist_vals):
                        if not TrestleUI._is_numeric(v):
                            errors.append(f"[车辆参数] {veh_name} 距前轴第{i+1}项无效: {v}")
                    for i, v in enumerate(weight_vals):
                        if not TrestleUI._is_numeric(v):
                            errors.append(f"[车辆参数] {veh_name} 轴重第{i+1}项无效: {v}")
                    if dist_vals and weight_vals and len(dist_vals) != len(weight_vals):
                        errors.append(f"[车辆参数] {veh_name} 距前轴({len(dist_vals)}项)与轴重({len(weight_vals)}项)数量不对应")
                    if dist_vals:
                        try:
                            last_dist = float(dist_vals[-1])
                            if last_dist != 0:
                                errors.append(f"[车辆参数] {veh_name} 距前轴最后一项应为0，当前: {dist_vals[-1]}")
                        except (ValueError, TypeError):
                            pass  # 非数值已在上面报错
            else:
                # 非一般车辆：逐项验证数值
                for i, p in enumerate(params):
                    if p is not None and p != "" and not TrestleUI._is_numeric(p):
                        errors.append(f"[车辆参数] {veh_name} 第{i+1}项参数无效: {p}")
            # 同名车辆参数一致性检查
            tag = veh_data.get("vehicle_tag", "")
            key = (veh_name, tag)
            if key in seen:
                if seen[key] != params:
                    errors.append(f"[车辆参数] 存在同名但参数不同的车辆: {veh_name}")
            else:
                seen[key] = params
        return errors

    def _validate_load_cases(self, proj: dict) -> list[str]:
        errors = []
        defined_vehicles = set(proj.get("vehicle", {}).keys())
        for case_name, case_data in proj.get("move_load_cases", {}).items():
            if not isinstance(case_data, dict):
                continue
            veh = case_data.get("vehicle", "")
            if veh and veh not in defined_vehicles:
                errors.append(f"[移动荷载] 工况 {case_name} 引用了不存在的车辆: {veh}")
        for case_name, case_data in proj.get("static_load_cases", {}).items():
            if not isinstance(case_data, dict):
                continue
            veh = case_data.get("vehicle", "")
            if veh and veh not in defined_vehicles:
                errors.append(f"[静载工况] 工况 {case_name} 引用了不存在的车辆: {veh}")
        return errors

    def _validate_loadcombs(self, proj: dict) -> list[str]:
        errors = []
        loadcombs = proj.get("load_combination", {})
        for comb_name, comb_data in loadcombs.items():
            if not isinstance(comb_data, dict):
                continue
            for case_name, coeffs in comb_data.get("cases", {}).items():
                if not isinstance(coeffs, dict):
                    continue
                # 系数有效性（gamma/psi 必须为 float）
                for coeff_key in ("gamma", "psi"):
                    val = coeffs.get(coeff_key)
                    if val is not None and val != "":
                        try:
                            float(val)
                        except (ValueError, TypeError):
                            errors.append(f"[荷载组合] {comb_name} 中工况 {case_name} 的 {coeff_key} 无效: {val}")
        return errors

    def combine_all_summary(self):
        """汇总所有 Tab 缓存为 trestle_ui_dict"""
        summary = {}
        # 项目身份信息（用于保存时定位 Excel 行）
        cur_proj = self.projects_dict.get(self._current_proj_key[0], {})
        summary["project_name"] = cur_proj.get("project_name", "")
        summary["project_no"] = cur_proj.get("project_no", "")
        # 透传元数据键（非 UI 编辑数据，供土层/地面线等 tab 切换后继续读取）
        for meta_key in ("_wb_sheets", "ground_line"):
            if meta_key in cur_proj:
                summary[meta_key] = cur_proj[meta_key]
        # 连接刚度库（共享 sheet，但随项目缓存整体替换，需透传防丢）
        summary["connection_stiffness"] = cur_proj.get("connection_stiffness", {})
        # 桥面系编号引用（保存时由 deck_no_map 重推，但需随项目保留完整缓存格式）
        summary["bridge_deck_type_ref"] = cur_proj.get("bridge_deck_type_ref", "1")
        # 土层标签（项目级别，由地形插件管理）
        summary["soil_label"] = cur_proj.get("soil_label", "")

        if hasattr(self, 'tab_supports'):
            s = self.tab_supports
            summary["basic"] = dict(getattr(s, 'basic_cache', {}))
            summary["substructure"] = copy.deepcopy(getattr(s, 'substructure_cache', {}))
            # 合并 SoilParamTab 的各桩土层数据（tksheet 实时收集）+ x_origin
            if hasattr(self, 'tab_soilparams'):
                soil = self.tab_soilparams
                # 收集所有桩号数据（自动补齐未访问桩）
                # substructure key 是 1-based，soil_cache key 是 0-based
                sub_keys = list(summary.get("substructure", {}).keys())
                soil._get_soil_cache(all_pile_nos=list(range(len(sub_keys))))
                soil_cache = soil._soil_cache  # 直接读 cache，key 是 0-based
                for idx, sk in enumerate(sub_keys):
                    pile_layers = soil_cache.get(str(idx), [])
                    if pile_layers:
                        soil_dict = {}
                        for layer in pile_layers:
                            soil_dict[layer["name"]] = [
                                layer["thickness"], layer["unit_weight"],
                                layer["cohesion"], layer["friction_angle"],
                                layer["qsik"], layer["qpk"],
                            ]
                        summary["substructure"][sk]["soil_data"] = soil_dict
                # x_origin 在 SoilParamTab 中，不在 SupportsTab
                xo = soil._x_origin_var.get().strip()
                if xo:
                    summary.setdefault("basic", {})["x_origin"] = xo

        # 计算各支撑处地面标高（土顶与桩交点标高）
        gl = trestle_excel_dict.get("ground_line", [])
        if gl and summary.get("substructure"):
            x_origin = float(summary.get("basic", {}).get("x_origin", 0) or 0)
            span_str = summary.get("basic", {}).get("span", "")
            spans = midasdisttolst(span_str) if span_str else []
            sub_keys = sorted(summary["substructure"].keys(), key=lambda k: int(k) if k.isdigit() else 0)
            # 计算各桩位 x（绝对坐标，与 ground_line 同系）
            pile_xs = [x_origin]
            for i, sp in enumerate(spans):
                dx = sp  # midasdisttolst 返回值已为米
                if i < len(sub_keys) and summary["substructure"].get(sub_keys[i], {}).get("type") == "制动墩":
                    dx += 0.2
                pile_xs.append(pile_xs[-1] + dx)
            # 地面线按 x 排序后线性插值
            gl_sorted = sorted(gl, key=lambda p: p[0])
            for idx, sk in enumerate(sub_keys):
                if idx >= len(pile_xs):
                    break
                px = pile_xs[idx]
                # 线性插值
                elev = gl_sorted[0][1]  # 默认取首点
                for j in range(len(gl_sorted) - 1):
                    x0, y0 = gl_sorted[j]
                    x1, y1 = gl_sorted[j + 1]
                    if x0 <= px <= x1:
                        if abs(x1 - x0) < 1e-9:
                            elev = y0
                        else:
                            elev = y0 + (y1 - y0) * (px - x0) / (x1 - x0)
                        break
                else:
                    elev = gl_sorted[-1][1]  # 超出范围取末点
                summary["substructure"][sk]["ground_elevation"] = round(elev, 3)

        # B2：用户编辑的土顶标高优先于地形插值（供 MCT 土弹簧使用 ground_elevation）
        _soil_top_cache = getattr(getattr(self, 'tab_soilparams', None), '_soil_top_cache', {})
        for _i, _sk in enumerate(sorted(summary.get("substructure", {}).keys(),
                                        key=lambda k: int(k) if k.isdigit() else 0)):
            _st = _soil_top_cache.get(str(_i), "")
            if _st:
                try:
                    summary["substructure"][_sk]["ground_elevation"] = round(float(_st), 3)
                except (ValueError, TypeError):
                    pass

        # bridge_width：统一取 basic_cache（tab_supports 存在时 basic_cache 必含该键，无需项目数据兜底）
        summary["bridge_width"] = summary.get("basic", {}).get("bridge_width", "")

        if hasattr(self, 'tab_components'):
            c = self.tab_components
            c.refresh_cache()
            # components_cache 为按桥型保留的会话缓存（可能含多个桥型），导出时只取当前桥型，
            # 保证下游"components 首键=当前桥型"的假设成立
            _bt = c.bridge_type_var.get() if hasattr(c, 'bridge_type_var') else ""
            _cc = getattr(c, 'components_cache', {})
            if _bt in _cc:
                _comp_data = copy.deepcopy(_cc[_bt])
                # deck_trans_ecc 和 deck_end_ext 从控件读取，放入 components 中
                try:
                    if hasattr(c, 'deck_trans_ecc'):
                        _comp_data["deck_trans_ecc"] = c.deck_trans_ecc.get()
                except tk.TclError:
                    pass
                try:
                    if hasattr(c, 'deck_end_ext'):
                        _comp_data["deck_end_ext"] = c.deck_end_ext.get()
                except tk.TclError:
                    pass
                summary["components"] = {_bt: _comp_data}
                summary["bridge_type"] = _bt
            else:
                summary["components"] = {}
                summary["bridge_type"] = _bt
            # 会话级：完整按桥型缓存（含全部桥型的未保存编辑），跨项目切换持久化用；
            # Excel/MCT 均只消费单桥型的 components，不读此键
            summary["_components_cache"] = copy.deepcopy(_cc)

        if hasattr(self, 'tab_connections'):
            conn = self.tab_connections
            conn.refresh_cache()
            summary["connection"] = dict(getattr(conn, 'connection_cache', {}))

        if hasattr(self, 'tab_vehicles'):
            summary["vehicle"] = dict(getattr(self.tab_vehicles, 'vehicle_cache', {}))
            # 快捷设置
            tv = getattr(self.tab_vehicles, 'quick_traffic_var', None)
            sv = getattr(self.tab_vehicles, 'quick_static_var', None)
            if tv and sv:
                qd = {}
                t_val = tv.get()
                s_val = sv.get()
                if t_val and t_val != "手动配置":
                    qd["traffic"] = t_val
                if s_val and s_val != "手动配置":
                    qd["static"] = s_val
                summary["quick_defaults"] = qd

        if hasattr(self, 'tab_enloads'):
            self.tab_enloads.refresh_cache()  # 确保环境参数最新
            summary["environment"] = dict(getattr(self.tab_enloads, 'environment_cache', {}))

        if hasattr(self, 'tab_moveloads'):
            summary["move_load_cases"] = dict(getattr(self.tab_moveloads, 'moveload_cache', {}))
            # 未点开过移动荷载 tab 时 cache 为空 → 复制项目持久化数据
            if not summary["move_load_cases"]:
                summary["move_load_cases"] = copy.deepcopy(self.trestle_excel_dict.get("move_load_cases", {}))

        if hasattr(self, 'tab_staticloads'):
            sl = self.tab_staticloads
            sl._save_current_state()  # 确保当前车辆 UI 状态保存到 car_dict
            sl.refresh_cache()
            summary["static_load_cases"] = dict(getattr(sl, 'staticload_cache', {}))
            # 未点开过静载 tab 时 cache 为空 → 复制项目持久化数据
            if not summary["static_load_cases"]:
                summary["static_load_cases"] = copy.deepcopy(self.trestle_excel_dict.get("static_load_cases", {}))

        if hasattr(self, 'tab_loadcombs'):
            summary["load_combination"] = dict(getattr(self.tab_loadcombs, 'load_combination_cache', {}))

        # ── 截面处理 ──
        sec_dict = getattr(self, 'section_dict', {})

        def _sec_name_of_ref(v):
            """引用字段 → 归一化名称"""
            v = str(v).strip()
            if not v or v == "/":
                return None
            return _norm_sec_name(v)

        # 1. 扫描 ref 字段，收集实际使用的截面归一化名
        used_norm = set()
        sec_ref_keys = ["pile_sec", "dist_trans_sec", "dist_long_sec", "brace_sec", "brace_sec_tilt"]

        for entry in summary.get("substructure", {}).values():
            for k in sec_ref_keys:
                name = _sec_name_of_ref(entry.get(k, "/"))
                if name:
                    used_norm.add(name)

        for entry in summary.get("components", {}).values():
            for k in ["rib_trans_sec", "rib_long_sec", "beam_sec"]:
                name = _sec_name_of_ref(entry.get(k, ""))
                if name:
                    used_norm.add(name)

        # 1.5 自动补注册：UI 引用了但 section_dict 中缺失的截面
        # 优先从项目 Excel 截面库查找，再从 properties_parameter.xlsx 查找
        existing_norm = {_norm_sec_name(r.get("name", "")) for r in sec_dict.values()}
        missing = used_norm - existing_norm
        if missing:
            # 先从当前项目 Excel 截面库补充
            proj_sec = load_project_section_library(self.excel_path)
            for name, rec in proj_sec.items():
                norm = _norm_sec_name(name)
                if norm in missing and norm not in existing_norm:
                    sec_dict[norm] = {"type": rec.get("type", ""), "name": norm, "params": rec.get("params", {})}
                    existing_norm.add(norm)
                    missing.discard(norm)
                    print(f"[补注册] 从项目截面库注册: {norm}")
            # 再从 properties_parameter.xlsx 补充
            if missing:
                lib = self.trestle_excel_dict.get("section_library_db") or {}
                print(f"[补注册] missing={missing}, properties库条目数={len(lib)}")
                for lib_name, lib_info in lib.items():
                    norm = _norm_sec_name(lib_name)
                    if norm in missing:
                        s_type = lib_info.get("type", "")
                        params = lib_info.get("params", {})
                        entry = {"type": s_type, "name": norm, "params": params}
                        self.record_section(entry, allow_update=False)
                        missing.discard(norm)
                        print(f"[补注册] 从properties注册: {norm}, type={s_type}")
                if missing:
                    print(f"[补注册] 未在任何截面库中找到: {missing}")

        # 2. 删除未被引用的截面，去重
        kept = []
        seen_names = set()  # 去重池：防止归一化同名的截面被错误收录多次
        for rec in sec_dict.values():
            norm_name = _norm_sec_name(rec.get("name", ""))
            if norm_name in used_norm and norm_name not in seen_names:
                # 强制覆盖为归一化名称，保证生成的 dict 里数据绝对干净
                rec["name"] = norm_name
                kept.append(rec)
                seen_names.add(norm_name)

        # 截面字典统一用归一化名称作 key（与 load_all_params 一致）：
        # MCT 生成由 _get_sec_id 按名称查插入位置编号，Excel 写入由 _write_section_sheet 按名称查重
        summary["section"] = {rec["name"]: rec for rec in kept}

        return summary

    def record_section(self, section_data, allow_update=True):
        """截面被选中时调用：按 (type, name) 查重，参数变化则更新，否则新增

        allow_update: True（用户手动设置）→ 已有同名截面则更新参数
                      False（自动注册）→ 已有同名截面则跳过
        """
        type_ = section_data.get("type", "")

        raw_name = section_data.get("custom_label", "") if section_data.get("name", "") == "自定义" else section_data.get("name", "")
        # 强制归一化
        norm_name = _norm_sec_name(raw_name)
        params = section_data.get("params", {})

        # 按 (type, 归一化名称) 查重
        for rec in self.section_dict.values():
            if (rec.get("type") == type_
                    and _norm_sec_name(rec.get("name", "")) == norm_name):
                if allow_update:
                    rec["params"] = params
                return

        # 新增：统一以归一化名称作 key（与 combine/_raw_to_cache 一致）
        self.section_dict[norm_name] = {"type": type_, "name": norm_name, "params": params}

    def save_params(self) -> bool:
        """保存参数 — 收集当前项目 → 验证 → 写入所有项目到 Excel。返回是否成功。"""
        self._force_refresh_all_caches()
        current_summary = self.combine_all_summary()
        self.projects_dict[self._current_proj_key[0]] = current_summary
        if not hasattr(self, 'tab_supports'):
            return False
        # 数据验证：遍历所有项目
        errors = []
        for key, proj in self.projects_dict.items():
            pname = proj.get("project_name", "")
            pno = proj.get("project_no", "")
            proj_label = f"{pname}{pno}" if pno else pname
            proj_errors = self._validate_project(proj)
            if proj_label:
                proj_errors = [f"[{proj_label}] {e}" for e in proj_errors]
            errors.extend(proj_errors)
        if errors:
            error_text = "\n".join(f"• {e}" for e in errors[:20])
            if len(errors) > 20:
                error_text += f"\n... 共 {len(errors)} 个问题"
            messagebox.showwarning("参数验证失败", f"以下参数存在问题：\n\n{error_text}\n\n请修正后重试。", parent=self.root)
            return False
        self._print_save_summary(current_summary)
        mct_path = self.save_path_var.get()
        save_path, template_ok, _ = save_trestle_to_excel(self.projects_dict, mct_path)
        if save_path:
            if template_ok:
                self._capture_saved_snapshot()
                messagebox.showinfo("完成", f"参数表已保存至 {save_path}")
                return True
            else:
                messagebox.showwarning("模板被占用",
                    f"参数表已保存至 {save_path}\n\n"
                    "但模板文件被占用，未能同步更新，请关闭模板文件后重试。")
                return False
        else:
            messagebox.showerror("失败", "Excel 文件被占用，请关闭后重试")
            return False

    @staticmethod
    def _print_save_summary(d):
        """简洁打印待保存的 trestle_ui_dict 结构"""
        print(f"\n{'='*60}")
        print(f"[保存] trestle_ui_dict 摘要")
        basic = d.get("basic", {})
        print(f"  基本: 跨径={basic.get('span','')}  宽度={d.get('bridge_width','')}  水位={basic.get('water_level','')}  桥面={basic.get('deck_level','')}")
        if basic.get("x_origin"):
            print(f"        x0={basic.get('x_origin','')}")

        sub = d.get("substructure", {})
        # 下部结构逐桩打印（键为 0 基桩号，非下部结构库编号）
        if sub:
            print(f"  下部结构({len(sub)}个):")
            for k in sorted(sub, key=lambda x: int(x) if str(x).isdigit() else 0):
                s = sub[k]
                print(f"    {k}#: {s.get('type',''):6s}  桩={s.get('pile_sec','/')}"
                      f"  桩横={s.get('pile_trans_space','/')}  桩纵={s.get('pile_long_space','/')}"
                      f"  brace={s.get('brace_form','-')}  offset={s.get('dist_trans_offset','/')}")

        comp = d.get("components", {})
        for bt, v in comp.items():
            print(f"  桥型: {bt}")
            print(f"    纵梁: type={v.get('beam_type','')} space={v.get('beam_space','')} first={v.get('beam_first_space','')} cant={v.get('beam_cant','')}")
            print(f"    桥面系: deck_t={v.get('deck_thickness','')} rib_sec={v.get('rib_trans_sec','')} rib_space={v.get('rib_trans_space','')}")

        conn = d.get("connection", {})
        if conn:
            print(f"  连接形式({len(conn)}项):")
            for k, v in conn.items():
                if isinstance(v, dict):
                    print(f"    {k}: type={v.get('type','')}")
                else:
                    print(f"    {k}: {v}")

        env = d.get("environment", {})
        wind = env.get("wind", {})
        water = env.get("water", {})
        wave = env.get("wave", {})
        print(f"  环境: 风={wind.get('design_wind_speed','')} 水={water.get('flow_velocity','')} 波={wave.get('wave_height','')}/{wave.get('wave_period','')}/{wave.get('wave_length','')}")

        veh = d.get("vehicle", {})
        lc = d.get("load_combination", {})
        print(f"  车辆：{len(veh)}辆")
        print(f"  组合：{len(lc)}个  ")
        print(f"\n{'='*60}")
        

    def _run_mct(self, mct_savepath=None):
        if mct_savepath is None:
            mct_savepath = self.save_path_var.get()
        self._force_refresh_all_caches()
        summary = self.combine_all_summary()
        self.projects_dict[self._current_proj_key[0]] = summary
        # 数据验证：当前项目
        errors = self._validate_project(summary)
        if errors:
            error_text = "\n".join(f"• {e}" for e in errors[:20])
            if len(errors) > 20:
                error_text += f"\n... 共 {len(errors)} 个问题"
            messagebox.showwarning("参数验证失败", f"以下参数存在问题：\n\n{error_text}\n\n请修正后重试。", parent=self.root)
            return
        import json
        print()
        print("========== trestle_ui_dict (MCT生成) ==========")
        print(json.dumps(summary, ensure_ascii=False, indent=2, default=str))
        print("==============================================")
        print()
        mcb_generate(summary, mct_savepath)
        

    def create_bottom_bar(self):
        # ---- 参考围堰UI: 分隔线 + 路径行 + 按钮行 ----
        bottom_bar = ttk.Frame(self.root)
        bottom_bar.pack(side="bottom", fill="x", padx=0, pady=0)
        ttk.Separator(bottom_bar, orient="horizontal").pack(fill="x")

        path_inner = ttk.Frame(bottom_bar)
        path_inner.pack(fill=X, padx=15, pady=8)
        bottom_inner = ttk.Frame(bottom_bar)
        bottom_inner.pack(fill=X, padx=15, pady=8)

        install_path = read_Register("Software\\ShuZhiQiaoShi", "Applocation")
        if not install_path:
            install_path = os.getcwd()
        default_mct_name = "Trestle_Untitled.mct"
        default_mct_path = os.path.join(install_path, "Custom", default_mct_name)
        self.save_path_var = tk.StringVar(value=default_mct_path)
        # mct_save_path = self.save_path_var.get()

        save_frame = ttk.Frame(path_inner)
        save_frame.pack(side=LEFT, fill=X, expand=True)
        ttk.Label(save_frame, text="保存路径:", bootstyle=PRIMARY, width=10).pack(side=LEFT)
        path_entry = ttk.Entry(save_frame, textvariable=self.save_path_var, state="readonly", width=60)
        path_entry.pack(side=LEFT, fill=X, expand=True, padx=5)
        path_btn = ttk.Button(save_frame, text="更改路径", command=self.choosesave_path,
                              bootstyle=(OUTLINE, PRIMARY), width=10)
        path_btn.pack(side=LEFT, padx=5)

        btn_frame = ttk.Frame(bottom_inner)
        btn_frame.pack(side=RIGHT)
        ttk.Button(btn_frame, text="退出", bootstyle=DANGER, width=8,
                   command=self._on_closing).pack(side=RIGHT, padx=4)
        
        self.qt_mct_btn = ttk.Button(btn_frame, text="导入桥通", bootstyle=PRIMARY, width=10)
        self.qt_mct_btn.pack(side=RIGHT, padx=4)
        # self.qt_mct_btn = ttk.Button(btn_frame, text="导入桥通", bootstyle=PRIMARY, width=10, command=self._run_mct)
        # self.qt_mct_btn.pack(side=RIGHT, padx=4)

        self.midas_mct_btn = ttk.Button(btn_frame, text="导入Midas", bootstyle=PRIMARY, width=10, command=self._run_mct)
        self.midas_mct_btn.pack(side=RIGHT, padx=4)

def Trestle_FEM_MidasCivil_Main(parent=None, on_close=None):
    if not if_Reg(parent):
        if on_close:
            on_close()
        return
    Applocation = read_Register('Software\\ShuZhiQiaoShi', 'Applocation') # 软件安装路径
    if parent == None:
        root = tb.Window(themename="litera")
        app = TrestleUI(root, Applocation)
    else:
        root = tb.Toplevel(parent)
        root.title("栈桥建模自动化程序")
        app = TrestleUI(root, Applocation)
    if parent == None:
        root.protocol("WM_DELETE_WINDOW", app._on_closing)
        root.mainloop()
    elif on_close:
        def _on_root_destroy(event, _root=root):
            if event.widget is _root:
                on_close()
        root.bind('<Destroy>', _on_root_destroy)

if __name__ == "__main__":
    Trestle_FEM_MidasCivil_Main()