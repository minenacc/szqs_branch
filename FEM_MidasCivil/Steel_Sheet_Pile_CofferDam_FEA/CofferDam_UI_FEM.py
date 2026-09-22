# =============================================================================
# 术语表（全文件统一约定）
#   SKGGZ  : 锁扣钢管桩
#   Waler  : 围檩；Strut: 内支撑（对撑 DC / 斜撑 XC）
#   Cap    : 承台；Corbel/Bracket: 牛腿
#   proj / proj_key : 单个项目参数字典 / 项目键（"项目名称+围堰编号"）
#   cache  : 各 Tab 的数据缓存，挂在 sheet_refs 中（basic_cache 等）
#   工况类型 : 取土/抽水/加撑/拆撑/封底/垫层/辅助加撑/辅助换撑/辅助加圈梁
#   单位约定 : 界面输入长度/标高为 m，截面参数 mm；土层表按列单位（见 Excel_io）
# =============================================================================

# 1. 标准库
import os
import math
import sys
import copy
import winreg
import tkinter as tk
from pathlib import Path
from tkinter import messagebox

# 2. 第三方库
import ttkbootstrap as tb
from ttkbootstrap.constants import *

# 3. 本地模块
sys.path.append(str(Path(__file__).parent.parent.parent))
from General.Midas          import Open_Midas_civil
from General.DataUtils      import read_Register
from General.ProgramMonitor import Program_State_Monitoring
from General.UIHandle       import if_Reg, createToolTip, on_exit
from General.FilePath       import qucik_save_file_dialog

sys.path.append(str(Path(__file__).parent))
from FEM_MidasCivil.Steel_Sheet_Pile_CofferDam_FEA.CofferDam_Diagram_FEM   import RealTimeDiagram
from FEM_MidasCivil.Steel_Sheet_Pile_CofferDam_FEA.CofferDam_Calculate_FEM import calc_skggz_center_spacing
from FEM_MidasCivil.Steel_Sheet_Pile_CofferDam_FEA.CofferDam_Generate_FEM  import Steel_Sheet_Pile_CofferDam_on_submit_MCT
from FEM_MidasCivil.Steel_Sheet_Pile_CofferDam_FEA.CofferDam_Excel_io_FEM  import load_all_params, save_to_excel, get_excel_path, PROJECT_TEMPLATE, global_connections, load_waler_strut_sections, load_material_library, serialize_stage_data, parse_stage_data


# ===== 主题配色 =====
SIDEBAR_BG = "#1a1d23"
SIDEBAR_ACTIVE = "#2b6fd4"
SIDEBAR_HOVER = "#262a33"
SIDEBAR_IDLE_BG = "#e0ecf8"
SIDEBAR_IDLE_FG = "#343a43"
CONTENT_BG = "#f5f6fa"
CARD_BG = "#ffffff"
TEXT_PRIMARY = "#2c3e50"
TEXT_SECONDARY = "#6b7280"
BORDER_COLOR = "#e5e7eb"

# ===== 工具函数 =====

def toggle_control(check_var1, check_var2, tb_lst1, tb_lst2):
    """互斥切换两组控件：勾选 check_var1 并启用 tb_lst1，取消 check_var2 并禁用 tb_lst2。

    Args:
        check_var1, check_var2 (tk.BooleanVar): 两个互斥的勾选变量。
        tb_lst1, tb_lst2 (list[tk.Widget]): 分别对应启用/禁用的控件列表。

    Returns:
        None
    """
    check_var1.set(True)
    check_var2.set(False)
    for x in tb_lst1:
        x.config(state="normal")
    for x in tb_lst2:
        x.config(state="disabled")


def mct_path_SaveAs(entry_var):
    """弹出保存对话框选择 .mct 计算书保存路径，并回填到 UI 输入框。

    副作用：弹出文件对话框，并就地修改 entry_var（Tk 变量）。

    Args:
        entry_var (tk.StringVar): 绑定到"计算书另存为路径"的 Tkinter 字符串变量。

    Returns:
        None
    """
    # 选择路径
    savefile_path = qucik_save_file_dialog('.mct')
    # 更新对应编辑框的值
    entry_var.set(savefile_path)


# ===== UI 辅助函数 =====

def make_form_row(parent, label_text, widget, label_width=18, tooltip=None, row_pady=4):
    """创建统一的"标签 + 控件"行，返回行容器。

    Args:
        parent (tk.Widget): 父容器。
        label_text (str): 标签文本。
        widget (tk.Widget): 放置在标签右侧的控件。
        label_width (int): 标签宽度（字符）。
        tooltip (str|None): 提示文本，非空时绑定 tooltip。
        row_pady (int): 行垂直间距。

    Returns:
        tk.Frame: 创建的行容器。
    """
    row_frame = tb.Frame(parent)
    row_frame.pack(fill=X, pady=row_pady, padx=5)
    lbl = tb.Label(row_frame, text=label_text, width=label_width, bootstyle=SECONDARY)
    lbl.pack(side=LEFT, padx=(10, 5))
    widget.pack(side=LEFT, padx=5)
    if tooltip:
        createToolTip(widget, tooltip)
    return row_frame


def make_section_title(parent, text):
    """创建带分隔线的分区标题，返回容器。

    Args:
        parent (tk.Widget): 父容器。
        text (str): 标题文本。

    Returns:
        tk.Frame: 标题容器。
    """
    style = tb.Style()
    style.configure("Section.TLabel", font=("Microsoft YaHei UI", 12, "bold"), foreground=TEXT_PRIMARY)
    container = tb.Frame(parent)
    container.pack(fill=X, pady=(15, 5))
    lbl = tb.Label(container, text=text, style="Section.TLabel")
    lbl.pack(side=LEFT, padx=15)
    sep = tb.Separator(container, orient=HORIZONTAL)
    sep.pack(side=LEFT, fill=X, expand=True, padx=10)
    return container

def make_card_frame(parent, title=None, **pack_kwargs):
    """创建卡片式容器，返回 (外框, 内容区)。

    Args:
        parent (tk.Widget): 父容器。
        title (str|None): 卡片标题。
        **pack_kwargs: 传给 outer.pack 的额外参数。

    Returns:
        tuple[tk.Frame, tk.Frame]: (outer 外框, content 内容区)。
    """
    outer = tb.Frame(parent)
    outer.pack(fill=X, padx=12, pady=5, **pack_kwargs)
    if title:
        lbl = tb.Label(outer, text=title, font=("Microsoft YaHei UI", 10, "bold"),
                        bootstyle=PRIMARY)
        lbl.pack(anchor=W, padx=15, pady=5)
    content = tb.Frame(outer)
    content.pack(fill=X, padx=5)
    return outer, content

def create_label_entry(parent, label_text, row, default_val="", placeholder="", 
                       label_width=15, entry_width=10, padx=5, pady=5, side="left", focus_callback=None):
    """创建"标签 + 输入框"，支持占位符与焦点回调。

    Args:
        parent (tk.Widget): 父容器。
        label_text (str): 标签文本。
        row (int): 行号（供调用方布局使用）。
        default_val (str): 默认值。
        placeholder (str): 占位符，聚焦时清除、失焦且为空时恢复。
        label_width, entry_width (int): 标签/输入框宽度。
        padx, pady (int): 间距。
        side (str): 布局方向。
        focus_callback (callable|None): 输入框聚焦时回调。

    Returns:
        tb.Entry: 创建的输入框控件。
    """
    tb.Label(parent, text=label_text, width=label_width).pack(side=side, padx=(10, padx), pady=pady)
    entry = tb.Entry(parent, width=entry_width)
    entry.pack(side=side, padx=(10, padx), pady=pady)
    
    if placeholder:
        entry.insert(0, placeholder)
        entry.config(foreground="gray")
        def on_focus_in(event):
            """聚焦时清除占位符并触发回调。"""
            if entry.get() == placeholder:
                entry.delete(0, tk.END)
                entry.config(foreground="black")
            if focus_callback: focus_callback()
        def on_focus_out(event):
            """失焦且为空时恢复占位符。"""
            if not entry.get():
                entry.insert(0, placeholder)
                entry.config(foreground="gray")
        entry.bind("<FocusIn>", on_focus_in)
        entry.bind("<FocusOut>", on_focus_out)
    else:
        if default_val: entry.insert(0, default_val)
        if focus_callback: entry.bind("<FocusIn>", lambda e: focus_callback())
    return entry


# ===== 主对话框 =====

def Steel_Pile_CofferDam_Dialog(root, Applocation):
    """构建钢板桩围堰建模主对话框（项目管理、侧边栏、7 个 Tab、图示与底部操作栏）。

    负责创建全部 Tk 变量、加载 Excel 参数与截面/材料库、装配各 Tab、
    实现输入校验、参数保存、MCT 生成与项目增删切换等交互逻辑。

    Args:
        root (tk.Tk|tk.Toplevel): 主窗口。
        Applocation (str): 程序安装目录，用于拼接默认保存路径。

    Returns:
        None: 副作用为构建 UI 并绑定事件，进入消息循环由调用方负责。
    """
    def _as_string(value, default=""):
        """安全转字符串并去空白，None 返回 default。"""
        if value is None: return default
        return str(value).strip()
    def _as_float(value, default=0.0):
        """安全转 float，失败返回 default。"""
        if value is None: return default
        try: return float(value)
        except: return default
    # 局部导入 Tab 函数（避免循环依赖）
    from CofferDam_Tab_FEM import (
        Basic_Setting_GUI, Waler_Strut_Concrete_Setting_GUI,
        Soil_Info_Setting_GUI, Boundary_Setting_GUI,
        LoadSetting_GUI, Condition_Analysis,
        _load_soil_into_frame,
        _parse_assist_bracket_data, _serialize_assist_bracket_data,
        _parse_assist_add_data, _serialize_assist_add_data,
        _serialize_assist_add_waler, _serialize_assist_add_strut,
        _resolve_steel_std,
    )

    # 获取屏幕的逻辑分辨率
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # 1. 设定动态初始比例
    target_width = int(screen_width * 0.7)
    target_height = int(screen_height * 0.8)

    # 2. 设定绝对最小值和最大值 
    MIN_WIDTH, MIN_HEIGHT = 1000, 1100   # 最小值设定得小一点，确保在轻薄本上也能正常缩放
    MAX_WIDTH, MAX_HEIGHT = 1250, 1350  # 宽屏/带鱼屏下的最大初始尺寸，避免铺得太大导致视线分散

    # 3. 将目标尺寸钳制 (Clamp) 在最小值和最大值之间
    win_width = max(MIN_WIDTH, min(target_width, MAX_WIDTH))
    win_height = max(MIN_HEIGHT, min(target_height, MAX_HEIGHT))

    # 4. 极端情况防御：如果物理屏幕比我们设定的最小值还小，就强行贴合屏幕大小
    win_width = min(win_width, screen_width)
    win_height = min(win_height, screen_height)

    # 5. 居中坐标计算 (y 轴稍微偏上 30 像素，给 Windows 底部任务栏留出视觉空间)
    x_cordinate = int((screen_width - win_width) / 2)
    y_cordinate = max(0, int((screen_height - win_height) / 2) - 30)

    # 6. 应用设置
    root.geometry(f"{win_width}x{win_height}+{x_cordinate}+{y_cordinate}")
    root.minsize(MIN_WIDTH, MIN_HEIGHT) # 使用更合理的最小值
    root.resizable(True, True)

    # 容许 Combobox popdown 销毁时的 KeyError
    def _report_callback_exception(exc, val, tb):
        """Tk 回调异常处理：忽略 Combobox popdown 销毁引发的 KeyError，其余打印堆栈。

        Args:
            exc (type): 异常类型。
            val (BaseException): 异常实例。
            tb (traceback): 异常回溯。

        Returns:
            None
        """
        if isinstance(val, KeyError) and 'popdown' in str(val):
            return
        import traceback
        traceback.print_exception(exc, val, tb)
    root.report_callback_exception = _report_callback_exception

    # 获取 Midas Civil NX 安装路径
    try:
        reg_path2 = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"SOFTWARE\MIDAS\CVLwNX_CH\PATH")
        civilnx_path = winreg.QueryValueEx(reg_path2, "Installed Path")[0]
        winreg.CloseKey(reg_path2)
    except Exception:
        civilnx_path = ""
    print('civilnx_path:', civilnx_path)

    # ---- 全局变量初始化（默认值，项目切换会覆盖） ----
    Solid_Level = tb.DoubleVar(value=0.0)
    Water_Level = tb.StringVar(value="0.0")
    CofferDam_Top_Level = tb.DoubleVar(value=0.0)
    CofferDam_L = tb.DoubleVar(value=0.0)

    Sheet_Pile_typelst = ['钢板桩', '锁扣钢管桩']
    Sheet_Pile_typevar = tb.StringVar(value='钢板桩')
    SKGGZ_D = tb.DoubleVar(value=0)
    SKGGZ_t = tb.DoubleVar(value=0)
    SKGGZ_Gap = tb.DoubleVar(value=0)
    SKGGZ_LockWidth = tb.DoubleVar(value=35)    # 锁扣宽度(mm)
    SKGGZ_SheetCount = tb.IntVar(value=1)    # 钢板桩数量
    SKGGZ_SheetWidth = tb.DoubleVar(value=600)  # 钢板桩宽度(mm)

    Pile_sec_values = {
        "拉森Ⅳ": [400, 170, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        "拉森Ⅵ": [600, 210, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    }
    Pile_SEC_info_dict = {'SEC': Pile_sec_values["拉森Ⅳ"]}

    Cap_X = tb.DoubleVar(value=0.0)
    Cap_Y = tb.DoubleVar(value=0.0)
    Cap_Bottom_Level = tb.DoubleVar(value=0.0)
    X_offset = tb.DoubleVar(value=1.0)
    Y_offset = tb.DoubleVar(value=1.0)
    Cap_H = tb.DoubleVar(value=0.0)

    Concrete_Blinding_check_var = tb.BooleanVar(value=False)
    Concrete_Blinding_grade_lst = ["C20", "C25", "C30", "C35", "C40"]
    Concrete_Blinding_grade_var = tb.StringVar(value="C20")
    Concrete_Blinding_thickness_var = tb.DoubleVar(value=0.0)

    Concrete_Plug_check_var = tb.BooleanVar(value=False)
    Concrete_Plug_grade_lst = ["C20", "C25", "C30", "C35", "C40", "C45", "C50", "C55", "C60"]
    Concrete_Plug_grade_var = tb.StringVar(value="C30")
    Concrete_Plug_thickness_var = tb.DoubleVar(value=0.0)
    Plug_support_d_var = tb.DoubleVar(value=2)
    Plug_support_long_var = tb.DoubleVar(value=0.0)
    Plug_support_short_var = tb.DoubleVar(value=0.0)
    Plug_bond_var = tb.DoubleVar(value=150)
    Plug_casing_count_var = tb.DoubleVar(value=8)
    Plug_casing_diam_var = tb.DoubleVar(value=1.2)

    Sheet_Pile_namelst = ["自定义", "SP-Ⅳ", "SP-ⅣW"]
    Sheet_Pile_namevar = tb.StringVar(value="自定义")
    Pile_Coords = tb.StringVar(value="")

    SDx = tb.StringVar(value="1e7")
    SDy = tb.StringVar(value="1e6")
    SDz = tb.StringVar(value="1e6")
    SRx = tb.StringVar(value="1")
    SRy = tb.StringVar(value="1")
    SRz = tb.StringVar(value="1")
    NSDx = tb.StringVar(value="")
    _boundary_conn_names = []
    corbel_long_num = tb.StringVar(value="2")
    corbel_short_num = tb.StringVar(value="2")
    Pile_Bottom_Boundary = tb.StringVar(value="001001")

    Walers_Strut_selected_dict = {}
    _boundary_conn_data = []   # 连接形式数据（从Excel读取）
    # 土层保存缓存
    _cofferdam_soil_data_cache = {}   # {proj_key: {layer_name: [7 vals]}} — 新建/修改的土层数据
    _cofferdam_soil_owner = {}        # {proj_key: True/False} — 项目是否拥有自己的sheet
    _cofferdam_soil_original = {}     # {proj_key: {layer_name: [7 vals]}} — 加载时的原始数据（用于检测修改）
    recognize_Cap_result_dict = {}
    recognize_Bracket_result_dict = {}
    recognize_Waler_result_dict = {}
    recognize_Strut_result_dict = {}
    recognize_Strut_Replace_result_dict = {}
    Steel_Sheet_Pile_CofferDam_Load_dict = {}
    Steel_Sheet_Pile_CofferDam_Load_dict['荷载计算规范'] = "建筑基坑支护技术规程"
    Steel_Sheet_Pile_CofferDam_Load_dict['围堰安全等级'] = "二级"
    Steel_Sheet_Pile_CofferDam_Load_dict['荷载自重系数'] = 1.0
    Steel_Sheet_Pile_CofferDam_Load_dict['围堰外初始附加应力'] = 0.0
    Steel_Sheet_Pile_CofferDam_Load_dict['基坑底初始附加应力'] = 0.0
    Steel_Sheet_Pile_CofferDam_Load_dict['荷载计算方法'] = ""
    Steel_Sheet_Pile_CofferDam_Load_dict['Excel'] = {}

    none_construction_stage_check_var = tb.BooleanVar(value=False)
    construction_stage_check_var = tb.BooleanVar(value=True)
    Drawdown_height = tb.DoubleVar(value=0.0)
    Waterdown_height = tb.DoubleVar(value=0.0)
    Excavation_face_dewater = tb.DoubleVar(value=0.0)  # 开挖面降水深度
    Spring_Thickness = tb.DoubleVar(value=0.5)
    Consider_Stress_Path = tb.StringVar(value="否")

    Construct_Stage_Line_Data = [['请点击自动生成按钮']]
    Load_Combo_Line_Data = []

    # ---- 构建主布局 ----
    # ---- 底部操作栏 ----
    bottom_bar = tb.Frame(root)
    bottom_bar.pack(side=BOTTOM, fill=X, padx=0, pady=0)

    separator = tb.Separator(bottom_bar, orient=HORIZONTAL)
    separator.pack(fill=X)

    path_inner = tb.Frame(bottom_bar)
    path_inner.pack(fill=X, padx=15, pady=8)
    bottom_inner = tb.Frame(bottom_bar)
    bottom_inner.pack(fill=X, padx=15, pady=8)

    # ---- 顶部项目栏 ----
    top_bar = tb.Frame(root, bootstyle=LIGHT)
    top_bar.pack(side=TOP, fill=X, pady=(0, 10))
    proj_inner = tb.Frame(top_bar, padding=(15, 10))
    proj_inner.pack(fill=X)

    # ---- 项目名称 + 围堰编号 双下拉 ----
    tb.Label(proj_inner, text="项目管理", font=("Microsoft YaHei UI", 10, "bold")).pack(side=LEFT, padx=(0, 10))
    project_name_combo = tb.Combobox(proj_inner, state="readonly", width=15)
    project_name_combo.pack(side=LEFT, padx=(0, 5))
    tb.Label(proj_inner, text="围堰编号", font=("Microsoft YaHei UI", 9)).pack(side=LEFT, padx=(0, 5))

    # ---- 加载数据 ----
    # 从Excel数据选项sheet读取围檩/内支撑截面列表(通用库读取)
    _excel_path = get_excel_path()
    _excel_waler_secs, _excel_strut_secs, Waler_section_dict, Strut_section_dict = load_waler_strut_sections(_excel_path)
    # print(f"围檩截面列表({_excel_waler_secs[:5]}...共{len(_excel_waler_secs)}项)")
    # print(f"内支撑截面列表({_excel_strut_secs[:5]}...共{len(_excel_strut_secs)}项)")
    Steel_dict, Concrete_dict, Rebar_dict = load_material_library(_excel_path)
    # print('钢结构材质库')
    # print(Steel_dict)
    # print('混凝土材质库')
    # print(Concrete_dict)
    # print('钢筋材质库')
    # print(Rebar_dict)
    # 项目参数主表读取
    _load_result = load_all_params(get_excel_path())
    projects_dict = _load_result["projects_dict"]
    project_groups = _load_result["project_groups"]
    last_proj_key = _load_result["last_proj_key"]
    _proj_keys_in_order = list(projects_dict.keys())  # 保持 Excel 行顺序
    # for k, v in projects_dict.items():
    #     print(k)
    #     print(v)
    #     print('')
    # 测试打印：加载后的全部项目数据
    # import json
    # print("=" * 50)
    # print("projects_dict 加载数据:")
    # print(json.dumps(projects_dict, ensure_ascii=False, indent=2, default=str))
    # print("=" * 50)

    # ---- 按钮区：新增 / 另存为 / 保存参数 / 删除 ----
    proj_btn_frame = tb.Frame(proj_inner)
    proj_btn_frame.pack(side=RIGHT, padx=5)
    proj_add_btn = tb.Button(proj_btn_frame, text="新增", bootstyle=(SUCCESS, OUTLINE), width=10)
    proj_add_btn.pack(side=LEFT, padx=2)
    proj_saveas_btn = tb.Button(proj_btn_frame, text="另存为", bootstyle=(INFO, OUTLINE), width=10)
    proj_saveas_btn.pack(side=LEFT, padx=2)
    proj_save_btn = tb.Button(proj_btn_frame, text="保存参数", bootstyle=(PRIMARY, OUTLINE), width=10)
    proj_save_btn.pack(side=LEFT, padx=2)
    proj_del_btn = tb.Button(proj_btn_frame, text="删除", bootstyle=(DANGER, OUTLINE), width=10)
    proj_del_btn.pack(side=LEFT, padx=2)

    project_no_combo = tb.Combobox(proj_inner, state="readonly")
    project_no_combo.pack(side=LEFT, fill=X, expand=True, padx=(0, 10))

    # ---- 当前项目 key（可变闭包变量）----
    _current_proj_key = [last_proj_key]

    tb.Separator(top_bar, bootstyle=SECONDARY).pack(fill=X)

    # 主内容区: 图示区固定高度，不可拖拽
    workspace = tk.PanedWindow(root, orient=VERTICAL, sashwidth=0, bg="#e5e7eb")
    workspace.pack(side=TOP, fill=BOTH, expand=True)

    # 上半部容器：侧边栏+内容区域
    top_half = tb.Frame(workspace)
    workspace.add(top_half, stretch="always")

    # 侧边栏
    sidebar = tk.Frame(top_half, bg=SIDEBAR_BG, width=170)
    sidebar.pack(side=LEFT, fill=Y)
    sidebar.pack_propagate(False)
    # 侧边栏右边界分隔线
    sidebar_sep = tb.Separator(top_half, orient=tk.VERTICAL)
    sidebar_sep.pack(side=LEFT, fill=Y)

    # 内容区域 (增加 Canvas 滚动条，防止小屏幕下输入框显示不全)
    content_outer = tb.Frame(top_half)
    content_outer.pack(side=LEFT, fill=BOTH, expand=True)

    content_canvas = tk.Canvas(content_outer, highlightthickness=0, bg=CONTENT_BG)
    content_scrollbar = tb.Scrollbar(content_outer, orient=VERTICAL)
    content_container = tb.Frame(content_canvas)

    def _clamped_scrollbar_set(*args):
        """防空白区域：限制滚动条不能超出内容范围"""
        lo, hi = float(args[0]), float(args[1])
        if lo < 0:
            lo = 0
        if hi > 1:
            hi = 1
        content_scrollbar.set(lo, hi)

    content_canvas.configure(yscrollcommand=_clamped_scrollbar_set)
    content_scrollbar.configure(command=content_canvas.yview)
    content_scrollbar.pack(side=RIGHT, fill=Y)
    content_canvas.pack(side=LEFT, fill=BOTH, expand=True)

    content_window = content_canvas.create_window((0, 0), window=content_container, anchor="nw")
    # 工况分析页：强制容器高度=视口高度。退出时由 switch_page 恢复自然尺寸
    def _fix_viewport_height(event=None):
        """在工况分析页（idx=5）强制内容容器高度等于视口高度。

        Args:
            event: Tk Configure 事件（可选）。

        Returns:
            None
        """
        if _current_page_idx[0] != 5:
            return
        try:
            vh = content_canvas.winfo_height()
            if vh > 50:
                content_canvas.itemconfig(content_window, height=vh)
        except:
            pass
    content_canvas.bind("<Configure>", _fix_viewport_height, add="+")

    _current_page_idx = [0]

    def _on_mousewheel(event):
        """内容区鼠标滚轮滚动（工况分析页不滚动，到边界不越界）。

        Args:
            event: 含 delta 的滚轮事件。

        Returns:
            None
        """
        if _current_page_idx[0] == 5:
            return
        if event.delta > 0 and content_canvas.yview()[0] <= 0:
            return
        if event.delta < 0 and content_canvas.yview()[1] >= 1:
            return
        content_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def on_configure_container(event):
        """内容容器尺寸变化时更新滚动区域。

        Args:
            event: Tk Configure 事件。

        Returns:
            None
        """
        content_canvas.configure(scrollregion=content_canvas.bbox("all"))

    def on_enter_content(_):
        """鼠标进入内容区时绑定全局滚轮事件。

        Args:
            _: Tk Enter 事件。

        Returns:
            None
        """
        content_canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def on_leave_content(_):
        """鼠标离开内容区时解绑全局滚轮事件。

        Args:
            _: Tk Leave 事件。

        Returns:
            None
        """
        content_canvas.unbind_all("<MouseWheel>")

    content_canvas.bind("<Enter>", on_enter_content)
    content_canvas.bind("<Leave>", on_leave_content)
    content_container.bind("<Configure>", on_configure_container)

    def on_configure_canvas(event):
        """画布宽度变化时同步内容窗口宽度，避免横向留白。

        Args:
            event: Tk Configure 事件。

        Returns:
            None
        """
        content_canvas.itemconfig(content_window, width=event.width)
    content_canvas.bind("<Configure>", on_configure_canvas)

    # 图示区域 (放入 PanedWindow 的下半部分)
    image_container = tk.Frame(workspace, bg="white", highlightbackground="#bdbfc3", highlightthickness=1)
    workspace.add(image_container, stretch="never")
    IMG_MINSIZE = max(int(win_height * 0.35), 200)
    workspace.paneconfig(image_container, minsize=IMG_MINSIZE)

    # ---- 控件注册表 (用于图示联动) ----
    widget_map = {}
    Waler_rows = []
    Strut_blocks = {}
    _waler_actions = {}

    # ---- 立面示意图 (需先于页面创建, 以便回调可用) ----
    var_dict = {
        "Solid_Level": Solid_Level, "Water_Level": Water_Level,
        "Cap_X": Cap_X, "Cap_Y": Cap_Y,
        "Cap_Bottom_Level": Cap_Bottom_Level,
        "X_offset": X_offset, "Y_offset": Y_offset,
        "Cap_H": Cap_H,
        "CofferDam_Top_Level": CofferDam_Top_Level,
        "CofferDam_L": CofferDam_L,
        "Sheet_Pile_namevar": Sheet_Pile_namevar,
        "Concrete_Blinding_thickness_var": Concrete_Blinding_thickness_var,
        "Concrete_Plug_thickness_var": Concrete_Plug_thickness_var,
        "Sheet_Pile_typevar": Sheet_Pile_typevar,
        "SKGGZ_D": SKGGZ_D,
        "SKGGZ_t": SKGGZ_t,
        "SKGGZ_Gap": SKGGZ_Gap,
    }
    diagram = RealTimeDiagram(image_container, var_dict,
                              Steel_Sheet_Pile_CofferDam_Load_dict, Waler_rows,
                              Concrete_Blinding_check_var, Concrete_Blinding_thickness_var,
                              Concrete_Plug_check_var, Concrete_Plug_thickness_var,
                              strut_blocks=Strut_blocks,
                              waler_section_dict=Waler_section_dict,
                              strut_section_dict=Strut_section_dict)

    # ---- 创建内容页面 ----
    pages = {}
    sheet_refs = {
        "_current_proj_key": _current_proj_key,
        "_cofferdam_soil_original": _cofferdam_soil_original,
        "waler_section_list": _excel_waler_secs,
        "strut_section_list": _excel_strut_secs,
        "waler_sec_dict": Waler_section_dict,
        "strut_sec_dict": Strut_section_dict,
    }
    _all_sections = {}
    for pk, pv in projects_dict.items():
        if pv.get("sections"):
            _all_sections = pv["sections"]
            break

    # 页面1: 基本设置 + 支护桩布置（合并）
    frame_env_cap = tb.Frame(content_container)
    Basic_Setting_GUI(frame_env_cap, Applocation, Solid_Level, Water_Level, Cap_X, Cap_Y,
                           X_offset, Y_offset, Cap_Bottom_Level, Cap_H, widget_map,
                           CofferDam_Top_Level=CofferDam_Top_Level, CofferDam_L=CofferDam_L,
                           Sheet_Pile_namevar=Sheet_Pile_namevar, Sheet_Pile_namelst=Sheet_Pile_namelst,
                           Pile_SEC_info_dict=Pile_SEC_info_dict, Pile_sec_values=Pile_sec_values,
                           root=root, Sheet_Pile_typevar=Sheet_Pile_typevar,
                           Sheet_Pile_typelst=Sheet_Pile_typelst, SKGGZ_D=SKGGZ_D, SKGGZ_t=SKGGZ_t,
                           SKGGZ_Gap=SKGGZ_Gap, applocation_sec=Applocation,
                            sheet_refs=sheet_refs, diagram=diagram, sections=_all_sections, mode="all",
                            SKGGZ_LockWidth=SKGGZ_LockWidth, SKGGZ_SheetCount=SKGGZ_SheetCount,
                            SKGGZ_SheetWidth=SKGGZ_SheetWidth,
                            steel_dict=Steel_dict, projects_dict=projects_dict,
                            current_proj_key_ref=_current_proj_key)
    pages[0] = frame_env_cap

    # 页面2: 支撑及封底
    frame_waler_strut = tb.Frame(content_container)
    Waler_Strut_Concrete_Setting_GUI(frame_waler_strut, Applocation, Waler_rows, Strut_blocks,
                                    Waler_section_dict, Strut_section_dict, Walers_Strut_selected_dict,
                                     Concrete_Plug_check_var, Concrete_Plug_thickness_var,
                                     Concrete_Plug_grade_var, Concrete_Plug_grade_lst,
                                     Concrete_Blinding_check_var, Concrete_Blinding_thickness_var,
                                     Concrete_Blinding_grade_var, Concrete_Blinding_grade_lst,
                                     widget_map, on_waler_change=lambda: [
                                         diagram.rebind_waler_widgets(),
                                         diagram.refresh(),
                                         sheet_refs.get("_update_waler_dependent_state", lambda *_: None)(
                                             len(Waler_rows),
                                             sum(1 for r in Waler_rows
                                                 if r.get("combo2", tb.StringVar()).get() not in ("", "/")
                                                 or r.get("combo3", tb.StringVar()).get() not in ("", "/"))
                                         )],
                                     _waler_actions=_waler_actions,
                                     plug_support_d_var=Plug_support_d_var,
                                     plug_support_long_var=Plug_support_long_var,
                                     plug_support_short_var=Plug_support_short_var,
                                     plug_bond_var=Plug_bond_var, plug_casing_count_var=Plug_casing_count_var,
                                     plug_casing_diam_var=Plug_casing_diam_var,
                                     pile_coords_var=Pile_Coords, sheet_refs=sheet_refs,
                                     steel_dict=Steel_dict, projects_dict=projects_dict,
                                     current_proj_key_ref=_current_proj_key,
                                     concrete_dict=Concrete_dict)
    pages[1] = frame_waler_strut

    # 页面4: 土层及边界
    frame_boundary = tb.Frame(content_container)
    # 页面4: 土层信息（Tab 4）
    frame_soil_info = tb.Frame(content_container)
    Soil_Info_Setting_GUI(frame_soil_info, Steel_Sheet_Pile_CofferDam_Load_dict, 
                         on_soil_change=lambda: diagram.refresh(),
                         sheet_refs=sheet_refs,
                         spring_thickness_var=Spring_Thickness,
                         consider_stress_path_var=Consider_Stress_Path)
    pages[2] = frame_soil_info

    # 页面5: 边界定义（Tab 5）
    frame_boundary = tb.Frame(content_container)
    Boundary_Setting_GUI(frame_boundary,
                         pile_boundary_var=Pile_Bottom_Boundary,
                         conn_data=_boundary_conn_data,
                         corbel_long_var=corbel_long_num, corbel_short_var=corbel_short_num,
                         skip_template_count=1,
                         sheet_refs=sheet_refs)

    # 页面6: 荷载参数（Tab 6）
    frame_load = tb.Frame(content_container)
    sigma_k_dict = {}
    LoadSetting_GUI(frame_load, sigma_k_dict,
                             sheet_refs=sheet_refs,
                             diagram=diagram)
    pages[3] = frame_boundary
    pages[4] = frame_load

    # 页面7: 工况分析（Tab 7）
    frame_analysis = tb.Frame(content_container)
    Condition_Analysis(frame_analysis, none_construction_stage_check_var, construction_stage_check_var,
                       Drawdown_height, Waterdown_height, Excavation_face_dewater,
                       Steel_Sheet_Pile_CofferDam_Load_dict, Load_Combo_Line_Data,
                       Construct_Stage_Line_Data, Waler_rows, Concrete_Plug_check_var,
                       Concrete_Blinding_check_var,
                       Solid_Level, Water_Level, sheet_refs)
    pages[5] = frame_analysis

    # ---- 输入校验 ----
    def validate_all_inputs():
        """检查所有标签页输入完整性，收集全部错误一次性返回。

        Returns:
            tuple[bool, list[tuple[str, str]]]: (是否通过, 错误列表)，
                错误项为 (Tab 名, 错误信息)。
        """
        errors = []

        def _is_num(value_str, is_level=False):
            """判断值是否为合法数字（允许前导 +/−）。

            Args:
                value_str: 待校验值。
                is_level (bool): 是否为标高类（仅用于语义，校验逻辑相同）。

            Returns:
                bool: 合法数字返回 True。
            """
            if isinstance(value_str, (int, float)):
                return True
            if not value_str or not str(value_str).strip():
                return False
            trimmed = str(value_str).strip()
            if trimmed.startswith('+') or trimmed.startswith('-'):
                trimmed = trimmed[1:]
            try:
                float(trimmed)
                return True
            except ValueError:
                return False

        # ================= 1. 基本设置 =================
        for name, var, is_lvl in [
            ("土顶标高", Solid_Level, True), ("水面标高", Water_Level, True),
            ("承台长", Cap_X, False), ("承台宽", Cap_Y, False),
            ("承台底标高", Cap_Bottom_Level, True), ("承台长边边距", X_offset, False),
            ("承台短边边距", Y_offset, False), ("承台高度", Cap_H, False),
        ]:
            try:
                val = var.get()
                # 水面标高允许输入 "/" 表示无水位
                if var is Water_Level and str(val).strip() == "/":
                    # 检查土层计算方法：存在水土分算时不允许设置为 "/"
                    method_lst = [v.get('计算方法', '') for v in Steel_Sheet_Pile_CofferDam_Load_dict.get('Excel', {}).values()]
                    if '水土分算' in method_lst:
                        errors.append(("基本设置", "土层计算方法中存在水土分算时不允许水面标高设置为'/'"))
                    continue
                if not _is_num(val, is_lvl):
                    errors.append(("基本设置", f"{name} 输入格式不正确"))
            except (tk.TclError, ValueError):
                errors.append(("基本设置", f"{name} 未输入或输入格式不正确"))

        # ================= 2. 支护桩布置 =================
        for name, var, is_lvl in [("围堰顶标高", CofferDam_Top_Level, True), ("围堰高", CofferDam_L, False)]:
            try:
                val = var.get()
                if not _is_num(val, is_lvl):
                    errors.append(("支护桩布置", f"{name} 输入格式不正确"))
            except (tk.TclError, ValueError):
                errors.append(("支护桩布置", f"{name} 未输入或输入格式不正确"))

        # 截面特性表值校验 —— 从六列 tksheet 读取 (列1=左半值, 列4=右半值)
        try:
            current_model = Sheet_Pile_namevar.get()
            s = sheet_refs.get("sec_sheet")
            if s:
                for ri, row_data in enumerate(s.get_sheet_data()):
                    for col_idx in (1, 4):
                        if col_idx < len(row_data):
                            var_symbol = row_data[col_idx - 1]
                            v_str = str(row_data[col_idx]).strip()
                            if not _is_num(v_str, False):
                                errors.append(("支护桩布置", f"型号【{current_model}】的截面参数 【{var_symbol}】 输入格式不正确"))
        except Exception as table_err:
            errors.append(("支护桩布置", f"截面特性表读取异常: {table_err}"))

        # ================= 3. 支撑及封底 =================
        # 3.1 围檩层间距检查 (直接读 Entry 当前值)
        for ri, row in enumerate(Waler_rows):
            try:
                sp = row['entry'].get().strip()
                if not _is_num(sp, False):
                    errors.append(("支撑及封底", f"间距{ri+1}  输入格式不正确"))
            except (tk.TclError, ValueError):
                errors.append(("支撑及封底", f"间距{ri+1}  未输入或输入格式不正确"))

        # 3.2 支撑布置检查 —— 仅允许数字、+、-、.、/
        def _is_strut_val(value_str):
            """判断支撑布置字符串是否合法（数字、+、-、.、,、/、空格）。

            Args:
                value_str (str): 待校验字符串。

            Returns:
                bool: 合法返回 True。
            """
            value_str = value_str.strip()
            if not value_str or value_str == '/':
                return True
            return all(c in '0123456789.+- ' for c in value_str) or \
                   all(c in '0123456789.+-, ' for c in value_str)

        for ri, row in enumerate(Waler_rows):
            dc_sec = row['combo2'].get()
            xc_sec = row['combo3'].get()
            strut_block = Strut_blocks[ri]
            for key, sec_name in [('DC_X', '对撑长边'), ('DC_Y', '对撑短边')]:
                if dc_sec != '/':
                    val = strut_block[key].get().strip()
                    if val and not _is_strut_val(val):
                        errors.append(("支撑及封底", f"第{ri+1}层{sec_name}布置 输入格式不正确"))
            for key, sec_name in [('XC_X', '斜撑长边'), ('XC_Y', '斜撑短边')]:
                if xc_sec != '/':
                    val = strut_block[key].get().strip()
                    if val and not _is_strut_val(val):
                        errors.append(("支撑及封底", f"第{ri+1}层{sec_name}布置 输入格式不正确"))

        # 3.3 牛腿数量检查：有围檩(任一层长边/短边围檩截面非空)时不允许为"/"
        has_waler = any(r.get('combo_long', tb.StringVar()).get().strip() not in ('', '/')
                        or r.get('combo_short', tb.StringVar()).get().strip() not in ('', '/')
                        for r in Waler_rows)
        if has_waler:
            for name, var in [("牛腿长边布置数量", corbel_long_num), ("牛腿短边布置数量", corbel_short_num)]:
                val = var.get()
                if not val or str(val).strip() in ("", "/"):
                    errors.append(("支撑及封底", f"已设置围檩，{name} 不能为空或/"))

        # 3.4 垫层和封底检查：
        if Concrete_Blinding_check_var.get():
            try:
                t_val = Concrete_Blinding_thickness_var.get()
                if not _is_num(t_val, False):
                    errors.append(("支撑及封底", f"启用了垫层，输入格式不正确"))
            except (tk.TclError, ValueError):
                errors.append(("支撑及封底", f"启用了垫层，垫层厚度 未输入或输入格式不正确"))

        if Concrete_Plug_check_var.get():
            try:
                t_val = Concrete_Plug_thickness_var.get()
                if not _is_num(t_val, False):
                    errors.append(("支撑及封底", f"启用了封底，输入格式不正确"))
            except (tk.TclError, ValueError):
                errors.append(("支撑及封底", f"启用了封底，封底厚度 未输入或输入格式不正确"))

        # ================= 4. 土层及边界 =================
        for name, var in [("SDx", SDx), ("SDy", SDy), ("SDz", SDz),
                          ("SRx", SRx), ("SRy", SRy), ("SRz", SRz)]:
            try:
                val = var.get()
                if not _is_num(val, False):
                    errors.append(("土层及边界", f"边界刚度 {name} 格式不正确"))
            except (tk.TclError, ValueError):
                errors.append(("土层及边界", f"边界刚度 {name} 未输入或输入格式不正确"))

        # 4.3 加权平均土层参数检查：勾选时加权重度必须为有效值且 >0（c/φ 允许为 0）
        try:
            _wadopt, _wg, _wc_, _wf_, _wok, _wmsg = _weighted_ui_values()
            if _wadopt and not _wok:
                errors.append(("土层及边界", _wmsg))
        except Exception:
            pass

        # ================= 5. 工况分析 =================
        # 5.1 超挖深度检查
        try:
            drawdown_value = Drawdown_height.get()
            if not _is_num(drawdown_value, False):
                errors.append(("工况分析", f"超挖深度 格式不正确"))
        except (tk.TclError, ValueError):
            errors.append(("工况分析", "超挖深度 未输入或输入格式不正确"))

        # 5.2 荷载组合表格 —— 直接从 tksheet 实时读取
        if none_construction_stage_check_var.get():
            try:
                sheet = sheet_refs.get("load_combo")
                if sheet:
                    all_rows = sheet.get_sheet_data()
                    # 前两行是伪表头, 从第3行起是数据
                    data_rows = all_rows[2:] if len(all_rows) > 2 else []
                    for ri, row in enumerate(data_rows):
                        for idx, fname in [(2, "分项系数"), (3, "组合值"), (4, "标准分项系数"), (5, "标准组合值")]:
                            if len(row) > idx:
                                val_str = str(row[idx]).strip()
                                if not _is_num(val_str, False):
                                    errors.append(("工况分析", f"荷载组合表格第 {ri+1} 行 【{fname}】 未输入或输入格式不正确"))
            except Exception:
                pass

        # # 5.3 施工阶段约束校验（仅施工阶段模式）
        if construction_stage_check_var.get():
            try:
                stage_ctrl = sheet_refs.get("construct_stage")
                if stage_ctrl and hasattr(stage_ctrl, "get_stage_dict"):
                    sd = stage_ctrl.get_stage_dict()
                    if sd:
                        ctx = _build_stage_check_context()
                        stage_msgs = _validate_stages(sd, ctx)
                        for msg in stage_msgs:
                            errors.append(("工况分析", msg))
            except Exception as e:
                errors.append(("工况分析", f"施工阶段校验异常: {e}"))

        # 完美返回最终的漏洞收集汇总
        if errors:
            return False, errors
        return True, []

    def _validate_all_projects():
        """对 projects_dict 中所有项目做全局数据校验（含施工阶段引用）。

        Returns:
            list[tuple[str, str]]: 错误列表，每项为 (项目标签, 错误信息)。
        """
        from CofferDam_Excel_io_FEM import parse_stage_data
        all_errors = []
        for pk, pv in projects_dict.items():
            pname = pv.get("项目名称", pk)
            pno = pv.get("围堰编号", "")
            tag = f"{pname}{pno}" if pno else pname

            for field in ["初始地面标高(m)", "初始水位标高(m)", "承台长(m)", "承台宽(m)",
                          "承台底标高(m)", "围堰顶标高", "围堰高"]:
                val = _as_string(pv.get(field, ""))
                try:
                    if val and val not in ("", "/"): float(val)
                except:
                    all_errors.append((tag, f"{field} '{val}' 格式不正确"))

            # 水面标高为"/"时，检查土层计算方法是否存在水土分算
            wl_val = _as_string(pv.get("初始水位标高(m)", ""))
            if wl_val == "/":
                soil = pv.get("soil_data", {})
                for layer_vals in soil.values():
                    if isinstance(layer_vals, (list, tuple)) and len(layer_vals) > 4 and layer_vals[4] == "水土分算":
                        all_errors.append((tag, "土层计算方法中存在水土分算时不允许水面标高设置为'/'"))
                        break

            if pv.get("是否生成施工阶段") == "是":
                stage_name = _as_string(pv.get("各工况名称", ""))
                stage_soil = _as_string(pv.get("各工况基坑内土顶标高(m)", ""))
                stage_water = _as_string(pv.get("各工况基坑内水面标高(m)", ""))
                stage_support = _as_string(pv.get("各工况支撑变动", ""))
                if stage_name:
                    sd = parse_stage_data(stage_name, stage_soil, stage_water, stage_support)
                    if sd:
                        ctx = {
                            'ground_level': _as_float(pv.get("初始地面标高(m)", 0)),
                            'water_level': _as_float(pv.get("初始水位标高(m)", 0)),
                            'water_above_ground': False,
                            'pit_bottom': _as_float(pv.get("围堰底标高", 0)),
                            'waler_count': 0, 'waler_elevations': [],
                            'has_plug': _as_float(pv.get("封底厚度(m)", 0)) > 0,
                            'has_blinding': _as_float(pv.get("垫层厚度(m)", 0)) > 0,
                            'assist_replace_names': [],
                            'assist_add_layers': [d.get("layer_no", "") for d in
                                _parse_assist_add_data(
                                    _as_string(pv.get("辅助加围檩i", "")),
                                    _as_string(pv.get("辅助加内支撑i", "")))],
                        }
                        ctx['water_above_ground'] = ctx['water_level'] > ctx['ground_level']
                        pile_top = _as_float(pv.get("围堰顶标高", 0))
                        current_elev = pile_top
                        for i in range(1, 6):
                            sp = _as_string(pv.get(f"围檩{i}间距(m)", ""))
                            if not sp or sp in ("", "/"): continue
                            if not _as_string(pv.get(f"围檩{i}长边截面", "")): continue
                            ctx['waler_count'] = i
                            try:
                                current_elev -= float(sp)
                                ctx['waler_elevations'].append(current_elev)
                            except: pass
                        stage_msgs = _validate_stages(sd, ctx)
                        for msg in stage_msgs:
                            all_errors.append((tag, msg))
        return all_errors

    # =========================================================================
    #                    施工阶段约束校验
    # =========================================================================

    def _build_stage_check_context():
        """构建施工阶段校验所需的上下文数据。

        Returns:
            dict: 结构形式
                {'ground_level': float, 'water_level': float|'/',
                 'water_above_ground': bool, 'pit_bottom': float,
                 'waler_count': int, 'waler_elevations': list[float],
                 'has_plug': bool, 'has_blinding': bool,
                 'assist_replace_names': list[str], 'assist_add_layers': list[str]}。
        """
        ctx = {}
        try: ctx['ground_level'] = float(Solid_Level.get())
        except: ctx['ground_level'] = 0.0
        try:
            _wl = Water_Level.get()
            ctx['water_level'] = "/" if str(_wl).strip() == "/" else float(_wl)
        except: ctx['water_level'] = "/"
        ctx['water_above_ground'] = ctx['water_level'] != "/" and ctx['water_level'] > ctx['ground_level']

        # 基坑底标高
        try: cap_bottom = float(Cap_Bottom_Level.get())
        except: cap_bottom = 0.0
        try: blinding = float(Concrete_Blinding_thickness_var.get())
        except: blinding = 0.0
        try: plug = float(Concrete_Plug_thickness_var.get())
        except: plug = 0.0
        extra = blinding if Concrete_Blinding_check_var.get() else plug if Concrete_Plug_check_var.get() else 0.0
        ctx['pit_bottom'] = round(cap_bottom - extra, 3)

        # 围檩数据
        ctx['waler_count'] = len(Waler_rows) if Waler_rows else 0
        try: pile_top = float(CofferDam_Top_Level.get())
        except: pile_top = 0.0
        ctx['waler_elevations'] = []
        current_elev = pile_top
        for wr in (Waler_rows or []):
            try: sp = float(wr['entry'].get().strip())
            except: sp = 0.0
            current_elev -= sp
            ctx['waler_elevations'].append(current_elev)

        ctx['has_plug'] = Concrete_Plug_check_var.get()
        ctx['has_blinding'] = Concrete_Blinding_check_var.get()

        if sheet_refs:
            ctx['assist_replace_names'] = [d.get("name", "") for d in sheet_refs.get("assist_replace_data", [])]
            ctx['assist_add_layers'] = [d.get("layer_no", "") for d in sheet_refs.get("assist_add_data", [])]
        else:
            ctx['assist_replace_names'] = []
            ctx['assist_add_layers'] = []
        return ctx

    def _validate_stages(stage_dict, ctx):
        """按工况类型逐条校验施工阶段约束（标高关系、层号、顺序等）。

        Args:
            stage_dict (dict): 施工阶段字典 {序号: {'工况类型','基坑内土顶标高',
                '基坑内水面标高','支撑层数'}}。
            ctx (dict): _build_stage_check_context 的输出。

        Returns:
            list[str]: 错误消息列表；为空表示通过。
        """
        msgs = []
        keys = sorted(stage_dict.keys())
        if not keys:
            msgs.append("施工阶段列表为空，请添加工况")
            return msgs

        g = ctx['ground_level']
        pit = ctx['pit_bottom']
        waler_els = ctx['waler_elevations']
        waler_cnt = ctx['waler_count']
        valid_layers = [str(i+1) for i in range(waler_cnt)]
        prev_water = None
        prev_soil = None
        prev_type = None
        # 记录已添加的加撑层号顺序（用于层高递增判断）
        added_support_layers = []  # [(layer, elev), ...]
        # 记录拆撑顺序（用于最后一道无后续判断）
        demolition_order = []
        # 用于"后次"检查的索引
        last_idx = keys[-1]

        for idx, k in enumerate(keys):
            sd = stage_dict[k]
            stype = sd.get("工况类型", "")
            soil_top_str = sd.get("基坑内土顶标高", "/")
            water_str = sd.get("基坑内水面标高", "/")
            support_str = sd.get("支撑层数", "/")
            is_last = (k == last_idx)

            def _f(val):
                """安全转 float，失败返回 None。"""
                try: return float(val)
                except: return None

            soil_top = _f(soil_top_str)
            water = _f(water_str)

            # ---- 全局规则 ----
            if idx == 0 and stype not in ("取土", "抽水"):
                msgs.append(f"工况{idx+1}: 首条工况必须为取土或抽水")

            # ---- 取土 ----
            if stype == "取土":
                if soil_top is not None:
                    if soil_top > g:
                        msgs.append(f"工况{k}: 基坑内土顶标高({soil_top:.3f})不得高于地面标高({g:.3f})")
                    if soil_top < pit:
                        msgs.append(f"工况{k}: 基坑内土顶标高({soil_top:.3f})不得低于基坑底标高({pit:.3f})")
                if water is not None and soil_top is not None:
                    # 水中围堰：不限制水面高于土顶
                    if not ctx.get("water_above_ground"):
                        if water > soil_top:
                            msgs.append(f"工况{k}: 基坑内水面标高({water:.3f})不得高于土顶标高({soil_top:.3f})")

            # ---- 加撑 ----
            elif stype == "加撑":
                if water is not None and prev_water is not None:
                    if water > prev_water:
                        msgs.append(f"工况{k}: 基坑内水面标高({water:.3f})不得高于前次水面标高({prev_water:.3f})")
                if support_str not in valid_layers:
                    msgs.append(f'工况{k}: 支撑层数 "{support_str}" 不在原始支撑层号范围内({",".join(valid_layers)})')
                # 层高递增检查
                if support_str in valid_layers:
                    s_idx = valid_layers.index(support_str)
                    if s_idx < len(waler_els):
                        curr_elev = waler_els[s_idx]
                        if added_support_layers:
                            last_layer, last_elev = added_support_layers[-1]
                            if curr_elev > last_elev:
                                msgs.append(f"工况{k}: 第{support_str}层围檩标高({curr_elev:.3f})高于上一层(第{last_layer}层{last_elev:.3f})，须从高到低")
                        added_support_layers.append((support_str, curr_elev))

            # ---- 抽水 ----
            elif stype == "抽水":
                if not ctx['water_above_ground']:
                    msgs.append(f"工况{k}: 抽水工况仅在水位高于地面时可用")
                if water is not None and soil_top is not None:
                    if soil_top > water:
                        msgs.append(f"工况{k}: 抽水阶段土顶标高({soil_top:.3f})不得高于水面标高({water:.3f})")
                if water is not None:
                    # 后次检查：当前抽水的水面不得低于后次工况的土顶/水面
                    for j in range(idx + 1, len(keys)):
                        next_k = keys[j]
                        next_sd = stage_dict[next_k]
                        next_soil = _f(next_sd.get("基坑内土顶标高", "/"))
                        next_water = _f(next_sd.get("基坑内水面标高", "/"))
                        if next_soil is not None and water < next_soil:
                            msgs.append(f"工况{k}: 抽水后水面标高({water:.3f})不得低于后次(工况{next_k})的土顶标高({next_soil:.3f})")
                        if next_water is not None and water < next_water:
                            msgs.append(f"工况{k}: 抽水后水面标高({water:.3f})不得低于后次(工况{next_k})的水面标高({next_water:.3f})")

            # ---- 封底 ----
            elif stype == "封底":
                if not ctx['has_plug']:
                    msgs.append(f"工况{k} 封底工况仅当支撑及垫层Tab勾选封底时可用")
                if water is not None and prev_water is not None:
                    if water > prev_water:
                        msgs.append(f"工况{k}: 封底水面标高({water:.3f})不得高于前次水面标高({prev_water:.3f})")

            # ---- 垫层 ----
            elif stype == "垫层":
                if not ctx['has_blinding']:
                    msgs.append(f"工况{k}: 垫层工况仅当支撑及垫层Tab勾选垫层时可用")
                if water is not None and prev_water is not None:
                    if water > prev_water:
                        msgs.append(f"工况{k}: 垫层水面标高({water:.3f})不得高于前次水面标高({prev_water:.3f})")

            # ---- 拆撑（仅首个检查前次，后续连续拆撑跳过）----
            elif stype == "拆撑":
                is_first_demolition = True
                # 检查当前是否为连续拆撑
                for j in range(idx - 1, -1, -1):
                    if stage_dict[keys[j]].get("工况类型") != "拆撑":
                        break
                    is_first_demolition = False
                if is_first_demolition and prev_type not in ("辅助加圈梁", "辅助换撑", "辅助加撑"):
                    msgs.append(f"工况{k}: 首个拆撐前次工况必须为辅助加圈梁/辅助换撑/辅助加撑（当前前次: {prev_type}）")
                # 拆撑可拆除基础支撑+辅助加撑
                assist_add = ctx.get('assist_add_layers', [])
                demo_layers = valid_layers + [l for l in assist_add if l not in valid_layers]
                if support_str not in demo_layers:
                    msgs.append(f'工况{k}: 支撑层数 "{support_str}" 不在支撑层号范围内({",".join(demo_layers)})')
                elif support_str in valid_layers:
                    s_idx = valid_layers.index(support_str)
                    if s_idx < len(waler_els):
                        curr_elev = waler_els[s_idx]
                        # 后次拆撑的层高检查
                        for j in range(idx + 1, len(keys)):
                            next_k = keys[j]
                            next_sd = stage_dict[next_k]
                            if next_sd.get("工况类型") == "拆撑":
                                next_sup = next_sd.get("支撑层数", "/")
                                if next_sup in valid_layers:
                                    n_idx = valid_layers.index(next_sup)
                                    if n_idx < len(waler_els):
                                        next_elev = waler_els[n_idx]
                                        if curr_elev > next_elev:
                                            msgs.append(f"工况{k}: 拆撑层标高({curr_elev:.3f})不得高于后次拆撑(工况{next_k})标高({next_elev:.3f})")
                                break

            # ---- 辅助加圈梁 ----
            elif stype == "辅助加圈梁":
                if soil_top is not None:
                    # 检查是否有后次拆撑
                    for j in range(idx + 1, len(keys)):
                        next_k = keys[j]
                        next_sd = stage_dict[next_k]
                        if next_sd.get("工况类型") == "拆撑":
                            next_sup = next_sd.get("支撑层数", "/")
                            if next_sup in valid_layers:
                                s_idx = valid_layers.index(next_sup)
                                if s_idx < len(waler_els):
                                    sup_elev = waler_els[s_idx]
                                    if soil_top > sup_elev:
                                        msgs.append(f"工况{k}: 土顶标高({soil_top:.3f})不得高于后次拆撑(工况{next_k})的支撑层标高({sup_elev:.3f})")
                            break

            # 辅助换撑/加撑数据检查在二级窗口保存时进行

            # 更新前次值
            if water is not None:
                prev_water = water
            if soil_top is not None:
                prev_soil = soil_top
            prev_type = stype

        return msgs

    # =========================================================================
    #                            校验
    # =========================================================================
    def _read_sheets_to_data():
        """从 tksheet 读取荷载组合与施工阶段最新数据到全局列表。

        就地更新 Load_Combo_Line_Data 与 Construct_Stage_Line_Data。

        Returns:
            None
        """
        _load_combo_sheet = sheet_refs.get("load_combo")
        if _load_combo_sheet:
            _data = _load_combo_sheet.get_sheet_data()
            if len(_data) > 2:
                Load_Combo_Line_Data[:] = [row[:6] for row in _data[2:]]
        _stage_sheet = sheet_refs.get("construct_stage")
        if _stage_sheet:
            _data = _stage_sheet.get_sheet_data()
            if len(_data) > 1:
                Construct_Stage_Line_Data[:] = [[row[0]] for row in _data[1:]]

    def _parse_surcharge_into_dict_save(proj, sk_dict):
        """将附加荷载字典序列化回项目字典的 Excel 字段。

        每类荷载编码为 "(启用面号,...), 值1,值2,..."；无启用面则写 "/"。

        Args:
            proj (dict): 单个项目参数字典（就地修改）。
            sk_dict (dict): 附加荷载信息 {类型: {面号: {字段: 值}}}。

        Returns:
            None
        """
        excel_map = {"均布附加荷载": "均布附加荷载", "矩形局部附加荷载": "矩形局部附加荷载", "条形局部附加荷载": "条形局部附加荷载"}
        fields_map = {"均布附加荷载": ["q0"], "矩形局部附加荷载": ["p0","angle","b","a","d","l","p2","c"], "条形局部附加荷载": ["p0","angle","b","a","d"]}
        for type_name, excel_key in excel_map.items():
            fields = fields_map[type_name]
            data = sk_dict.get(type_name, {})
            active_faces = []
            values_tokens = {}
            for fi in ["1","2","3","4"]:
                entry = data.get(fi, {})
                if entry.get("enabled", True) is False:
                    continue
                has_value = any(entry.get(f, "/") not in ("", "/") for f in fields)
                if not has_value:
                    continue
                active_faces.append(fi)
                tokens = [entry.get(f, "/") for f in fields]
                values_tokens[fi] = tokens
            if not active_faces:
                proj[excel_key] = "/"
                continue
            face_str = ",".join(active_faces)
            first_tokens = values_tokens[active_faces[0]]
            val_str = ",".join(first_tokens)
            proj[excel_key] = f"({face_str}), {val_str}"

    def _weighted_ui_values():
        """读取当前 UI 加权平均土层参数勾选态与三值。

        返回 (adopted, wg, wc, wf, valid, msg)：
          adopted —— 是否勾选"采用加权平均土层参数"（以 UI 当前选择为准）
          valid  —— 勾选时加权重度是否为有效数值且 >0（黏聚力/内摩擦角允许为 0）
        """
        if not sheet_refs:
            return False, 0.0, 0.0, 0.0, True, ""
        w_check = sheet_refs.get("_weighted_check")
        if not (w_check and w_check.get()):
            return False, 0.0, 0.0, 0.0, True, ""
        try:
            w_zd = sheet_refs.get("_weighted_重度")
            w_nj = sheet_refs.get("_weighted_黏聚力")
            w_nm = sheet_refs.get("_weighted_内摩擦角")
            wg = float(w_zd.get()) if w_zd else 0.0
            wc = float(w_nj.get()) if w_nj else 0.0
            wf = float(w_nm.get()) if w_nm else 0.0
        except (tk.TclError, ValueError, TypeError):
            return True, 0.0, 0.0, 0.0, False, "已勾选采用加权平均土层参数，重度/黏聚力/内摩擦角必须为有效数字"
        if wg <= 0:
            return True, wg, wc, wf, False, "已勾选采用加权平均土层参数，加权平均重度必须大于0"
        return True, wg, wc, wf, True, ""

    def _regular_soil_status():
        """检查当前 UI 常规土层表（非加权）的有效性。

        返回 (code, msg)：
          code 为 None   —— 土层信息完整有效；
          code == 'empty'   —— 表中没有任何土层名称行，msg='土层信息为空'；
          code == 'invalid' —— 存在土层行但数值列有空白或'/'或非数字，
                                msg='土层信息存在空值或无效值'。
        """
        try:
            sf = sheet_refs.get("soil_frame")
            if sf is None:
                return "empty", "土层信息为空"
            raw = sf.sheet.get_sheet_data() if hasattr(sf, "sheet") else sf.get_sheet_data()
        except Exception:
            return "empty", "土层信息为空"
        # 收集带土层名称的行
        rows = []
        for row in raw:
            if row and str(row[0]).strip():
                rows.append(row)
        if not rows:
            return "empty", "土层信息为空"
        # 列: 0名称 1层厚 2重度 3黏聚力 4内摩擦角 5计算方法（均不允许为空或'/'）
        _label_lst = [(1, "层厚"), (2, "重度"), (3, "黏聚力"), (4, "内摩擦角"), (5, "计算方法")]
        for row in rows:
            for _i, _label in _label_lst:
                if len(row) <= _i or str(row[_i]).strip() in ("", "/"):
                    return "invalid", "土层信息存在空值或无效值"
            for _col in (1, 2, 3, 4):
                try:
                    float(str(row[_col]).strip())
                except ValueError:
                    return "invalid", "土层信息存在空值或无效值"
            if str(row[5]).strip() not in ("水土合算", "水土分算"):
                return "invalid", "土层信息存在空值或无效值"
        return None, ""

    def combine_cofferdam_dict():
        """读取当前 UI 全部数据，写入 projects_dict[当前项目键]。

        读取流程：先调用各 Tab 的 refresh_cache 更新 cache，再从各 cache
        提取数据写入 projects_dict（中文表头键格式）。

        Returns:
            None: 就地修改 projects_dict[_current_proj_key[0]]。
        """
        key = _current_proj_key[0]
        if not key or key not in projects_dict:
            return
        proj = projects_dict[key]

        # ---- 1. 调用各 Tab refresh_cache ----
        for refresh_key in ["refresh_basic", "refresh_substructure", "refresh_boundary",
                            "refresh_surcharge", "refresh_condition"]:
            fn = sheet_refs.get(refresh_key)
            if fn:
                fn()

        # ---- 2. 从 cache 提取 → 写入 proj（中文表头名 key）----
        # 基本参数（来自 basic_cache + 直接读取）
        basic = sheet_refs.get("basic_cache", {})
        proj["项目名称"] = project_name_combo.get()
        proj["围堰编号"] = project_no_combo.get()
        proj["初始地面标高(m)"] = basic.get("ground_level", "")
        proj["初始水位标高(m)"] = basic.get("water_level", "")
        proj["支护桩类型"] = basic.get("pile_type", "")
        proj["围堰顶标高"] = basic.get("pile_top_level", "")
        proj["围堰高"] = basic.get("pile_length", "")
        # 围堰底标高 = 围堰顶标高 - 围堰高（自动计算，只读），使校验以当前 UI 为准
        try:
            _pile_top = float(str(proj["围堰顶标高"]).strip() or "0")
            _pile_len = float(str(proj["围堰高"]).strip() or "0")
            proj["围堰底标高"] = f"{_pile_top - _pile_len:.3f}"
        except (ValueError, TypeError):
            proj["围堰底标高"] = ""
        proj["承台长(m)"] = basic.get("cap_length", "")
        proj["承台宽(m)"] = basic.get("cap_width", "")
        proj["承台高(m)"] = basic.get("cap_height", "")
        proj["承台底标高(m)"] = basic.get("cap_bottom_level", "")
        proj["承台长边预留边距(m)"] = basic.get("cap_offset_long", "")
        proj["承台宽边预留边距(m)"] = basic.get("cap_offset_short", "")
        # 围堰内边长/边宽 (从 sheet_refs 的 StringVar 读取)
        if sheet_refs:
            _il = sheet_refs.get("_inner_length_var")
            _iw = sheet_refs.get("_inner_width_var")
            proj["围堰内边长(m)"] = str(_il.get()) if _il else "0"
            proj["围堰内边宽(m)"] = str(_iw.get()) if _iw else "0"
        # 加权平均土层参数：勾选 → 写 "γ,c,φ"(最多3位小数)；不勾选 → 写 "/"
        if sheet_refs:
            def _w3(v):
                """把数值格式化为最多 3 位小数、且至少保留 1 位小数的字符串。"""
                s = f"{round(float(v), 3):.3f}".rstrip("0").rstrip(".")
                return s if "." in s else s + ".0"
            _wd = sheet_refs.get("_weighted_重度")
            _wc = sheet_refs.get("_weighted_黏聚力")
            _wf = sheet_refs.get("_weighted_内摩擦角")
            _weighted_check = sheet_refs.get("_weighted_check") # 是否采用加权平均值 按钮控件 布尔值
            try:
                adopted = bool(_weighted_check and _weighted_check.get() and _wd and _wc and _wf)
                if adopted:
                    proj["加权平均重度+粘聚力+摩擦角"] = f"{_w3(_wd.get())},{_w3(_wc.get())},{_w3(_wf.get())}"
                else:
                    proj["加权平均重度+粘聚力+摩擦角"] = '/'
            except (tk.TclError, ValueError):
                proj["加权平均重度+粘聚力+摩擦角"] = '/'
        proj["钢护筒坐标(m)"] = str(Pile_Coords.get())
        proj["规范"] = basic.get("code_standard", "")
        proj["围堰安全等级"] = basic.get("safety_grade", "")
        proj["钢结构规范"] = "《钢结构设计标准》 （GB 50017-2017）"
        proj["混凝土规范"] = "《混凝土结构设计标准》（GB/T 50010-2010）"
        # 截面编号：直接存截面名，save_to_excel 会自动转为数字编号
        proj["支护桩截面编号"] = basic.get("section_ref", "")
        # 支护桩材质：从详细设置保存的 sheet_refs 读取，保证保存参数时写回参数表对应列
        _pile_type = basic.get("pile_type", "")
        _mat_key = "_last_gbz_material" if _pile_type == "钢板桩" else "_last_skggz_material"
        _saved_mat = sheet_refs.get(_mat_key, "") if sheet_refs else ""
        if _saved_mat:
            proj["支护桩材质"] = _saved_mat
        # 承台顶标高 = 承台底标高 + 承台高
        try:
            cap_bottom = float(str(basic.get("cap_bottom_level", 0)).strip() or "0")
            cap_height = float(str(basic.get("cap_height", 0)).strip() or "0")
            proj["承台顶标高(m)"] = str(round(cap_bottom + cap_height, 3))
        except (ValueError, TypeError):
            proj["承台顶标高(m)"] = ""

        # 支撑及封底（来自 substructure_cache）
        sub = sheet_refs.get("substructure_cache", {})
        # 垫层与封底互斥：仅保存勾选方的厚度与混凝土等级，未勾选方厚度与等级写/
        if Concrete_Blinding_check_var.get():
            proj["垫层厚度(m)"] = sub.get("blinding_thickness", "")
            proj["垫层混凝土等级"] = sub.get("blinding_grade", "")
            proj["封底厚度(m)"] = "/"
            proj["封底混凝土等级"] = "/"
        elif Concrete_Plug_check_var.get():
            proj["封底厚度(m)"] = sub.get("plug_thickness", "")
            proj["封底混凝土等级"] = sub.get("plug_grade", "")
            proj["垫层厚度(m)"] = "/"
            proj["垫层混凝土等级"] = "/"
        else:
            proj["垫层厚度(m)"] = "/"
            proj["封底厚度(m)"] = "/"
            proj["垫层混凝土等级"] = "/"
            proj["封底混凝土等级"] = "/"
        proj["钢与混凝土黏聚力(kPa)"] = sub.get("plug_bond", "")
        proj["护筒数量"] = sub.get("plug_casing_count", "")
        proj["护筒直径(m)"] = sub.get("plug_casing_diam", "")
        plug_long = sub.get("plug_support_long", "0")
        plug_short = sub.get("plug_support_short", "0")
        proj["封底支撑间距(m)"] = f"{plug_long},{plug_short}"
        for i, wd in enumerate(sub.get("waler_rows", [])):
            n = i + 1
            proj[f"围檩{n}间距(m)"] = wd["spacing"]
            proj[f"围檩{n}长边截面"] = wd["sec_long"]
            proj[f"围檩{n}宽边截面"] = wd["sec_short"]
            proj[f"围檩{n}长边材质"] = wd["mat_long"]
            proj[f"围檩{n}宽边材质"] = wd["mat_short"]
            proj[f"内支撑{n}对撑截面"] = wd["dc_sec"]
            proj[f"内支撑{n}斜撑截面"] = wd["xc_sec"]
            proj[f"内支撑{n}对撑材质"] = wd["mat_dc"]
            proj[f"内支撑{n}斜撑材质"] = wd["mat_xc"]
            proj[f"内支撑{n}对撑长边布置(m)"] = wd["dc_x"]
            proj[f"内支撑{n}对撑短边布置(m)"] = wd["dc_y"]
            proj[f"内支撑{n}斜撑长边布置(m)"] = wd["xc_x"]
            proj[f"内支撑{n}斜撑短边布置(m)"] = wd["xc_y"]
        # 清除未使用的围檩/支撑层数据
        for n in range(len(sub.get("waler_rows", [])) + 1, 6):
            for suffix in ["间距(m)", "长边截面", "宽边截面", "长边材质", "宽边材质"]:
                proj.pop(f"围檩{n}{suffix}", None)
            for suffix in ["对撑截面", "斜撑截面", "对撑材质", "斜撑材质",
                          "对撑长边布置(m)", "对撑短边布置(m)", "斜撑长边布置(m)", "斜撑短边布置(m)"]:
                proj.pop(f"内支撑{n}{suffix}", None)
                proj.pop(f"{n}{suffix}", None)
        # 土层数据（来自 soil_frame_cache）
        soil_frame = sheet_refs.get("soil_frame")
        if soil_frame:
            # 1. 从 tksheet 结构中读取当前用户修改后的土层数据
            raw_sheet_data = soil_frame.sheet.get_sheet_data()
            updated_soil = {}
            for row in raw_sheet_data:
                if row[0]: # 土层名称不为空
                    updated_soil[row[0]] = [row[1], row[2], row[3], row[4], row[5], row[6], row[7]]
            proj["soil_data"] = updated_soil

            key = _current_proj_key[0]

            def _normalize_soil(d):
                """把土层字典的值统一为去空白字符串，便于比较是否被修改。

                Args:
                    d (dict): {土层名称: [值, ...]}。

                Returns:
                    dict: 归一化后的同结构字典；d 为空时返回 {}。
                """
                return {k: [str(v).strip() if v not in (None, "") else "" for v in vals]
                        for k, vals in d.items()} if d else {}

            # 从 sheet_refs 读取当前引用标签（与 _on_existing_soil_selected 同步）
            current_ref_tag = sheet_refs.get("cofferdam_soil_ref", "") if sheet_refs else "" 

            if current_ref_tag and current_ref_tag in projects_dict and current_ref_tag != key:
                # 场景A：用户选择了别的围堰作为参考。基准数据就是那个被参考的围堰数据
                baseline_data = projects_dict[current_ref_tag].get("soil_data", {})
            else:
                # 场景B：用户没有参考别人，或者是修改自己的数据。基准就是自己初始的 original
                baseline_data = _cofferdam_soil_original.get(key, {})

            # 2. 判断当前表格数据是否被修改（与基准对比）
            is_modified = (_normalize_soil(updated_soil) != _normalize_soil(baseline_data))

            # 3. 判断是否有有效数据
            has_valid_soil = any(
                str(v).strip() not in ("", "/")
                for vals in updated_soil.values()
                for v in vals[:5]
            ) if updated_soil else False

            # ==========================================
            # 4. 执行标签与存储逻辑
            # ==========================================
            if has_valid_soil:
                # --- 调试代码 ---
                n_updated = _normalize_soil(updated_soil)
                n_baseline = _normalize_soil(baseline_data)
                
                if n_updated != n_baseline:
                    print("【调试】土层数据对比不一致！")
                    print("Updated keys:", list(n_updated.keys()))
                    print("Baseline keys:", list(n_baseline.keys()))
                    # 遍历查找具体的不同点
                    for k in set(n_updated.keys()) | set(n_baseline.keys()):
                        if n_updated.get(k) != n_baseline.get(k):
                            print(f"键 '{k}' 不匹配:")
                            print(f"  界面值: {n_updated.get(k)}")
                            print(f"  基准值: {n_baseline.get(k)}")
                # ----------------
                if is_modified:
                    # 必须作为独立的 sheet 存储，打上自己的标签
                    _cofferdam_soil_data_cache[key] = updated_soil
                    _cofferdam_soil_owner[key] = True
                    proj["地质参数"] = f"1,{key}" if key else ""
                else:
                    proj["地质参数"] = f"1,{current_ref_tag}" if current_ref_tag else ""
            else:
                proj["地质参数"] = f"1,{current_ref_tag}" if current_ref_tag else ""
            
        # 边界定义（来自 boundary_cache）
        boundary = sheet_refs.get("boundary_cache", {})
        proj["支护桩底边界"] = boundary.get("pile_bottom_boundary", "")
        proj["牛腿长边布置数量"] = boundary.get("corbel_long_num", "/")
        proj["牛腿短边布置数量"] = boundary.get("corbel_short_num", "/")
        for conn in boundary.get("connections", []):
            proj[conn["key"]] = conn["type"]

        # 工况分析（来自 condition_cache）
        condition = sheet_refs.get("condition_cache", {})
        proj["水面至基坑底距离(m)"] = condition.get("drawdown", "")
        proj["超挖深度(m)"] = condition.get("overdig_depth", "")
        proj["开挖面降水高度(m)"] = condition.get("excavation_face_dewater", "")
        proj["是否生成施工阶段"] = "是" if construction_stage_check_var.get() else "否"
        # 辅助换撑/加撑序列化
        if sheet_refs is not None:
            proj["辅助换撑"] = _serialize_assist_bracket_data(sheet_refs.get("assist_replace_data", []))
            proj["辅助加围檩i"] = _serialize_assist_add_waler(sheet_refs.get("assist_add_data", []))
            proj["辅助加内支撑i"] = _serialize_assist_add_strut(sheet_refs.get("assist_add_data", []))
        # 施工阶段表格序列化
        stage_controller = sheet_refs.get("construct_stage") if sheet_refs else None
        if stage_controller and hasattr(stage_controller, "get_stage_dict"):
            stage_dict = stage_controller.get_stage_dict()
            stage_name, stage_soil_top, stage_water, stage_support = serialize_stage_data(stage_dict)
            proj["各工况名称"] = stage_name
            proj["各工况基坑内土顶标高(m)"] = stage_soil_top
            proj["各工况基坑内水面标高(m)"] = stage_water
            proj["各工况支撑变动"] = stage_support
        proj["土弹簧分层厚度(m)"] = str(Spring_Thickness.get())
        proj["考虑土的应力路径"] = Consider_Stress_Path.get()
        if len(Load_Combo_Line_Data) > 0:
            first_row = Load_Combo_Line_Data[0]
            if len(first_row) >= 6:
                proj["基本组合分项系数"] = first_row[2]
                proj["基本组合组合值系数"] = first_row[3]
                proj["标准组合分项系数"] = first_row[4]
                proj["标准组合组合值系数"] = first_row[5]
        # 荷载参数（来自 surcharge_cache）
        surcharge = sheet_refs.get("surcharge_cache", {})
        if surcharge:
            _parse_surcharge_into_dict_save(proj, surcharge)

    def _validate_soil_for_save(proj):
        """土层保存前验证：非空、完整、且土层底标高低于桩底标高。

        Args:
            proj (dict): 单个项目参数字典，需含 'soil_data' 等字段。

        Returns:
            tuple[bool, str]: (是否通过, 错误信息)；通过时错误信息为 ""。
        """
        soil = proj.get("soil_data", {})
        if not soil:
            return False, "土层参数不能为空，请先设置土层信息"
        # 检查完整性：至少1条土层，且关键字段非空
        valid_layers = 0
        for name, vals in soil.items():
            if not name or not str(name).strip():
                continue
            # 前5个字段（层厚、重度、黏聚力、内摩擦角、计算方法）需有值
            has_data = any(str(v).strip() not in ("", "/") for v in vals[:5])
            if has_data:
                valid_layers += 1
        if valid_layers == 0:
            return False, "土层参数不完整，请确保至少有一条完整土层数据"
        # 检查最底层底标高 < 桩底标高
        try:
            total_thickness = 0.0
            for name, vals in soil.items():
                if not name or not str(name).strip():
                    continue
                try:
                    thickness = float(str(vals[0]).strip()) if vals[0] not in (None, "", "/") else 0.0
                    total_thickness += thickness
                except (ValueError, IndexError):
                    pass
            top_level = _as_float(proj.get("初始地面标高(m)", 0))
            bottom_level = top_level - total_thickness
            pile_bottom = _as_float(proj.get("围堰底标高", -8))
            if bottom_level >= pile_bottom:
                return False, f"土层底标高({bottom_level:.3f}m)应低于桩底标高({pile_bottom:.3f}m)"
        except Exception:
            pass  # 标高校验失败不阻止保存
        return True, ""

    def param_save():
        """保存参数：校验通过后更新内存并写入 Excel。

        校验加权平均土层、常规土层、全局项目数据与土层底标高，任一不通过
        弹窗提示并中止。

        Returns:
            None: 副作用为写 Excel 与弹窗提示。
        """
        key = _current_proj_key[0]
        if not key:
            return
        # 加权平均土层参数校验：勾选时加权重度必须 >0
        _wadopt, _wg, _wc_, _wf_, _wok, _wmsg = _weighted_ui_values()
        if _wadopt and not _wok:
            messagebox.showwarning("加权平均土层参数", _wmsg)
            return
        # 先收集当前项目数据
        combine_cofferdam_dict()
        # print(projects_dict[key])

        # 未采用加权平均土层时：常规土层表不得为空或含空值/无效值
        if not _wadopt:
            _soil_code, _soil_msg = _regular_soil_status()
            if _soil_code:
                messagebox.showwarning(_soil_msg, _soil_msg)
                return

        # 全局项目校验（从 projects_dict 读取）
        cross_errors = _validate_all_projects()
        if cross_errors:
            msg_lines = [f"[{tag}] {msg}" for tag, msg in cross_errors]
            messagebox.showwarning("参数检查未通过", "\n".join(msg_lines))
            return
        #（调试信息统一在 Steel_Sheet_Pile_CofferDam_on_submit_MCT 中输出）

        # 土层保存验证（采用加权平均土层时，常规土层可为空，跳过校验）
        proj = projects_dict.get(key, {})
        if not _wadopt:
            ok, msg = _validate_soil_for_save(proj)
            if not ok:
                messagebox.showwarning("土层参数检查", msg)
                return
        excel_path = get_excel_path()
        if excel_path:
            save_to_excel(excel_path, projects_dict, project_keys_in_order=_proj_keys_in_order,
                          soil_owner=_cofferdam_soil_owner,
                          soil_data_cache=_cofferdam_soil_data_cache)
            print(f"[CofferDam] 参数已保存: {key}")
            messagebox.showinfo("保存成功", f"所有项目参数均已保存")

    def _mct_generate():
        """生成 MCT：同步土层、刷新缓存、校验输入后调用建模主入口。

        校验不通过时弹窗提示并中止；通过则组装辅助加撑/换撑、材质与规范，
        调用 Steel_Sheet_Pile_CofferDam_on_submit_MCT 生成模型文件。

        Returns:
            None: 副作用为写 MCT/MCB 文件与弹窗提示。
        """
        # 同步土层 tksheet → Load_dict（确保未提交编辑也被保存）
        try:
            sf = sheet_refs.get("soil_frame")
            if sf and hasattr(sf, "sheet"):
                raw = sf.sheet.get_sheet_data()
                new_dict = {}
                soil_idx = 0
                for row in raw:
                    if not row or not row[0]: continue
                    name = str(row[0]).strip()
                    if not name: continue
                    soil_idx += 1
                    new_dict[f"第{soil_idx}层土"] = {"土层名称": name,
                        "层厚": str(row[1]) if len(row) > 1 else "/",
                        "重度": str(row[2]) if len(row) > 2 else "/",
                        "黏聚力": str(row[3]) if len(row) > 3 else "/",
                        "内摩擦角": str(row[4]) if len(row) > 4 else "/",
                        "计算方法": str(row[5]) if len(row) > 5 else "/",
                        "渗透系数(m/d)": str(row[6]) if len(row) > 6 else "/",
                        "土层顶承压水头(m)": str(row[7]) if len(row) > 7 else "/"}
                Steel_Sheet_Pile_CofferDam_Load_dict["Excel"] = new_dict
        except: pass
        # 再刷新全部缓存
        for refresh_key in ["refresh_basic", "refresh_substructure", "refresh_boundary",
                            "refresh_surcharge", "refresh_condition"]:
            fn = sheet_refs.get(refresh_key)
            if fn:
                fn()
        ok, error_list = validate_all_inputs()
        if not ok:
            msg_lines = []
            current_tab = ""
            for tab, msg in error_list:
                if tab != current_tab:
                    current_tab = tab
                    msg_lines.append(f"\n【{current_tab}】")
                msg_lines.append(f"{msg}")
            messagebox.showwarning("建模数据检查未通过", "\n".join(msg_lines))
            return
        # 土层保存验证（采用加权平均土层时，常规土层可为空，跳过校验）
        _wadopt_mct, *_ = _weighted_ui_values()
        if not _wadopt_mct:
            combine_cofferdam_dict()
            _proj_mct = projects_dict.get(_current_proj_key[0], {})
            _soil_ok, _soil_msg = _validate_soil_for_save(_proj_mct)
            if not _soil_ok:
                messagebox.showwarning("土层参数检查", _soil_msg)
                return
        _read_sheets_to_data()
        # if civilnx_path:
        #     state = Program_State_Monitoring(civilnx_path, "MIDAS CIVIL NX")
        #     if state == "未安装":
        #         messagebox.showerror("错误", "未检测到 Midas Civil NX 安装")
        #         return
        #     elif state == "未打开":
        #         Open_Midas_civil(civilnx_path, tb.StringVar(value=state))
        # 从 boundary_cache 读取弹性连接数据
        boundary = sheet_refs.get("boundary_cache", {})
        form_dict = boundary.get("form_dict", {})
        stiffness_dict = boundary.get("stiffness_dict", {})
        # 收集辅助加撑数据并转换格式
        # 原始格式: [{'layer_no': '3', 'elevation': '1.906', 'sec_long': '2I45a', ...}]
        # 目标格式: {'3': {'标高': '1.906', '围檩长边截面': '2I45a', ...}}
        assist_add_data = sheet_refs.get("assist_add_data", [])
        assist_add_dict = {}
        for item in assist_add_data:
            layer_no = item.get("layer_no", "")
            if not layer_no:
                continue
            assist_add_dict[layer_no] = {
                '标高': item.get("elevation", "/"),
                '围檩长边材质': item.get("mat_long", "/"),
                '围檩长边截面': item.get("sec_long", "/"),
                '围檩宽边材质': item.get("mat_short", "/"),
                '围檩宽边截面': item.get("sec_short", "/"),
                '是否与承台连接': item.get("connected", "否"),
                '对撑材质': item.get("mat_dc", "/"),
                '对撑截面': item.get("dc_sec", "/"),
                '斜撑材质': item.get("mat_xc", "/"),
                '斜撑截面': item.get("xc_sec", "/"),
                '对撑长边布置(m)': item.get("dc_x", "/"),
                '对撑短边布置(m)': item.get("dc_y", "/"),
                '斜撑长边布置(m)': item.get("xc_x", "/"),
                '斜撑短边布置(m)': item.get("xc_y", "/"),
                '承台长边投影(m)': item.get("proj_long", "/"),
                '承台短边投影(m)': item.get("proj_short", "/"),
                '承台长边高差(m)': item.get("diff_long", "/"),
                '承台短边高差(m)': item.get("diff_short", "/"),
            }
        # 收集辅助换撑数据并转换格式
        # 原始格式: [{'name': '1-1', 'waler_layer': '第1层', 'connected': '否', ...}]
        # 目标格式: {1: {1: {'是否与承台连接': '否', ...}}, 2: {1: {...}, 2: {...}}}
        assist_replace_data = sheet_refs.get("assist_replace_data", [])
        assist_replace_dict = {}
        for item in assist_replace_data:
            name = item.get("name", "")
            if not name or "-" not in name:
                continue
            parts = name.split("-")
            try:
                layer_i = int(parts[0])
                form_j = int(parts[1])
            except (ValueError, IndexError):
                continue
            if layer_i not in assist_replace_dict:
                assist_replace_dict[layer_i] = {}
            assist_replace_dict[layer_i][form_j] = {
                '是否与承台连接': item.get("connected", "否"),
                '对撑材质': item.get("mat_dc", "/"),
                '对撑截面': item.get("dc_sec", "/"),
                '斜撑材质': item.get("mat_xc", "/"),
                '斜撑截面': item.get("xc_sec", "/"),
                '对撑长边布置(m)': item.get("dc_x", "/"),
                '对撑短边布置(m)': item.get("dc_y", "/"),
                '斜撑长边布置(m)': item.get("xc_x", "/"),
                '斜撑短边布置(m)': item.get("xc_y", "/"),
                '承台长边投影(m)': item.get("proj_long", "/"),
                '承台短边投影(m)': item.get("proj_short", "/"),
                '承台长边高差(m)': item.get("diff_long", "/"),
                '承台短边高差(m)': item.get("diff_short", "/"),
            }
        # ---- 施工阶段引用到的加撑/换撑条目一致性校验（仅生成MCT、仅施工阶段模式）----
        # 只校验“施工阶段表格中实际引用到”的条目；未被引用的加撑/换撑不校验、不修改。
        if construction_stage_check_var.get():
            _assist_ref_errors = []
            _stage_ctrl_cur = sheet_refs.get("construct_stage") if sheet_refs else None
            if _stage_ctrl_cur and hasattr(_stage_ctrl_cur, "get_stage_dict"):
                _stage_dict_cur = _stage_ctrl_cur.get_stage_dict() or {}
                _n_waler = len(Waler_rows) if Waler_rows else 0
                _assist_add_layer_set = set(assist_add_dict.keys())
                for _stage_cur in _stage_dict_cur.values():
                    if not _stage_cur:
                        continue
                    _stype_cur = str(_stage_cur.get("工况类型", "")).strip()
                    _sup_cur = str(_stage_cur.get("支撑层数", "/")).strip()
                    if _sup_cur in ("", "/"):
                        continue
                    if _stype_cur == "辅助换撑":
                        _parts_cur = _sup_cur.split("-")
                        if len(_parts_cur) >= 2 and _parts_cur[0].strip().isdigit() \
                                and _parts_cur[1].strip().isdigit():
                            _li_cur = int(_parts_cur[0])
                            _ji_cur = int(_parts_cur[1])
                            if _li_cur not in assist_replace_dict \
                                    or _ji_cur not in assist_replace_dict.get(_li_cur, {}):
                                _assist_ref_errors.append(
                                    f'施工阶段辅助换撑引用了未定义的换撑数据“{_sup_cur}”')
                            # 换撑锚定的围檩层必须真实存在（原始层或辅助加撑层）
                            if _li_cur > _n_waler and str(_li_cur) not in _assist_add_layer_set:
                                _assist_ref_errors.append(
                                    f'施工阶段辅助换撑“{_sup_cur}”锚定的第{_li_cur}层围檩不存在'
                                    f'（当前原始围檩{_n_waler}层）')
                    elif _stype_cur == "辅助加撑":
                        if _sup_cur not in assist_add_dict:
                            _assist_ref_errors.append(
                                f'施工阶段辅助加撑引用了不存在的辅助加撑层“{_sup_cur}”')
                    elif _stype_cur == "拆撑":
                        if _sup_cur.isdigit() and int(_sup_cur) > _n_waler \
                                and _sup_cur not in _assist_add_layer_set:
                            _assist_ref_errors.append(
                                f'施工阶段拆撑引用了不存在的辅助加撑层“{_sup_cur}”')
            if _assist_ref_errors:
                messagebox.showwarning(
                    "辅助加撑/换撑数据检查",
                    "以下施工阶段引用存在问题：\n" + "\n".join(_assist_ref_errors))
                return
        # 支护桩材质：直接取对应材质下拉框当前值（无需点击“保存”即可生效）
        _pile_type = Sheet_Pile_typevar.get()
        if _pile_type == '钢板桩':
            _mat_var = sheet_refs.get("_gbz_material_var") if sheet_refs else None
            _sheet_pile_material = _mat_var.get() if _mat_var else "Q295"
        elif _pile_type == '锁扣钢管桩':
            _mat_var = sheet_refs.get("_skggz_material_var") if sheet_refs else None
            if _mat_var and _mat_var.get():
                _sheet_pile_material = _mat_var.get()
            else:
                _cur_proj = projects_dict.get(_current_proj_key[0], {}) if projects_dict and _current_proj_key[0] in projects_dict else {}
                _sheet_pile_material = _cur_proj.get("支护桩材质", "") or "Q235"
        else:
            _sheet_pile_material = "Q235"
        # 钢结构规范解析（用于材料定义中 standard / elast 的取值）
        _steel_std = None
        if projects_dict and _current_proj_key[0] in projects_dict:
            _steel_std = _resolve_steel_std(
                projects_dict[_current_proj_key[0]].get("钢结构规范", ""), Steel_dict)
        # 生成前：若勾选“采用加权平均土层参数”，以加权单层土代替分层土层表（仅本次生成生效，
        # 以当前 UI 选择为准，不改写原土层表 / Excel；未勾选则原样传入）
        _wadopt, _wg, _wc_, _wf_, _wok, _wmsg = _weighted_ui_values()
        # 未采用加权平均土层时：常规土层表不得为空或含空值/无效值
        if not _wadopt:
            _soil_code, _soil_msg = _regular_soil_status()
            if _soil_code:
                messagebox.showwarning(_soil_msg, _soil_msg)
                return
        if _wadopt and _wok:
            try:
                _wtop = float(Solid_Level.get())
                _wbot = float(CofferDam_Top_Level.get()) - float(CofferDam_L.get())  # 围堰底=围堰顶-围堰高
                _wthick = round(_wtop - _wbot + 1, 3)
            except (ValueError, TypeError, tk.TclError):
                _wthick = 0.0
            load_dict_for_gen = dict(Steel_Sheet_Pile_CofferDam_Load_dict)
            load_dict_for_gen['Excel'] = {
                '第1层土': {
                    '土层名称': '加权平均土层',
                    '层厚': _wthick,
                    '重度': round(_wg, 3),
                    '黏聚力': round(_wc_, 3),
                    '内摩擦角': round(_wf_, 3),
                    '计算方法': '水土合算',
                    '渗透系数(m/d)': 0,
                    '土层顶承压水头(m)': 0,
                }
            }
        else:
            load_dict_for_gen = Steel_Sheet_Pile_CofferDam_Load_dict
        Steel_Sheet_Pile_CofferDam_on_submit_MCT(
            Cap_X.get(), Cap_Y.get(), Cap_H.get(), X_offset.get(), Y_offset.get(), Spring_Thickness.get(),
            Sheet_Pile_typevar.get(), _sheet_pile_material, CofferDam_L.get(), Sheet_Pile_namevar.get(),
            Pile_SEC_info_dict, SKGGZ_D.get(), SKGGZ_t.get(), SKGGZ_Gap.get(),
            Waler_rows, Strut_blocks,
            corbel_long_num.get(), corbel_short_num.get(),
            Solid_Level.get(), Water_Level.get(),
            Cap_Bottom_Level.get(), CofferDam_Top_Level.get(),
            Concrete_Blinding_check_var.get(),
            Concrete_Blinding_thickness_var.get(),
            Concrete_Plug_check_var.get(),
            Concrete_Plug_thickness_var.get(),
            Drawdown_height.get(), Waterdown_height.get(), Excavation_face_dewater.get(),
            Waler_section_dict, Strut_section_dict,
            recognize_Cap_result_dict, recognize_Bracket_result_dict,
            recognize_Waler_result_dict, recognize_Strut_result_dict,
            recognize_Strut_Replace_result_dict,
            load_dict_for_gen,
            sigma_k_dict,
            Pile_Bottom_Boundary.get(),
            form_dict,
            stiffness_dict,
            none_construction_stage_check_var.get(), construction_stage_check_var.get(),
            Load_Combo_Line_Data, Construct_Stage_Line_Data,
            sheet_refs.get("construct_stage").get_stage_dict() if sheet_refs and sheet_refs.get("construct_stage") and hasattr(sheet_refs.get("construct_stage"), "get_stage_dict") else {},
            Consider_Stress_Path.get() == "是",
            mct_savepath.get(), gen_mcb_var.get(), True,
            assist_add_dict,
            assist_replace_dict,
            Steel_dict,
            _steel_std,
        )

    # ===== 项目管理 =====
    def _parse_surcharge_into_dict(proj, sk_dict):
        """从项目字典的附加荷载字段解析回 sk_dict（与序列化互为逆操作）。

        Args:
            proj (dict): 单个项目参数字典。
            sk_dict (dict): 附加荷载信息字典，函数内先清空再写入。

        Returns:
            None: 就地修改 sk_dict。
        """
        import re
        saved_active = sk_dict.get("active_type", "均布附加荷载")
        sk_dict.clear()
        sk_dict["active_type"] = saved_active
        fields_map = {"均布附加荷载": ["q0"], "矩形局部附加荷载": ["p0","angle","b","a","d","l","p2","c"], "条形局部附加荷载": ["p0","angle","b","a","d"]}
        excel_map = {"均布附加荷载": "均布附加荷载", "矩形局部附加荷载": "矩形局部附加荷载", "条形局部附加荷载": "条形局部附加荷载"}
        for type_name, excel_key in excel_map.items():
            fields = fields_map[type_name]
            sk_dict[type_name] = {str(i): {f: "/" for f in fields} for i in range(1, 5)}
            raw = str(proj.get(excel_key, "")).strip()
            if not raw or raw == "/": continue
            m = re.match(r"\s*\((.+?)\)\s*,?(.*)", raw)
            if not m: continue
            faces = [int(x.strip()) for x in m.group(1).split(",") if x.strip().isdigit()]
            rest = m.group(2).strip()
            if not rest: continue
            tokens = [x.strip() for x in rest.split(",")]
            for fi in faces:
                if fi < 1 or fi > 4: continue
                entry = {}
                for j, fk in enumerate(fields):
                    entry[fk] = tokens[j] if j < len(tokens) else "/"
                sk_dict[type_name][str(fi)] = entry

    _switching_flag = [False]  # 防递归

    def _switch_project(proj_key):
        """切换项目：先保存当前项目到内存，再把新项目数据加载到各 UI 控件。

        Args:
            proj_key (str): 目标项目键。

        Returns:
            None: 副作用为刷新大量 Tk 变量与图示。
        """
        if _switching_flag[0] or proj_key not in projects_dict:
            return
        _switching_flag[0] = True

        # ① 保存当前项目（读取 UI → 写入 projects_dict[旧key]）
        old_key = _current_proj_key[0]
        if old_key and old_key != proj_key and old_key in projects_dict:
            old_proj = projects_dict[old_key]
            old_name = old_proj.get("项目名称", "")
            old_no = old_proj.get("围堰编号", "")
            combine_cofferdam_dict()
            # 修正被 combo 改写过的名称/编号
            proj_saved = projects_dict[old_key]
            proj_saved["项目名称"] = old_name
            proj_saved["围堰编号"] = old_no

        # ② 切换到新项目
        _current_proj_key[0] = proj_key
        proj = projects_dict[proj_key]
        # 基坑参数
        Solid_Level.set(_as_float(proj.get("初始地面标高(m)", 0)))
        Water_Level.set(_as_string(proj.get("初始水位标高(m)", "0.0")))
        Cap_X.set(_as_float(proj.get("承台长(m)", 0)))
        Cap_Y.set(_as_float(proj.get("承台宽(m)", 0)))
        Cap_Bottom_Level.set(_as_float(proj.get("承台底标高(m)", 0)))
        Cap_H.set(_as_float(proj.get("承台高(m)", 0)))
        Pile_Coords.set(_as_string(proj.get("钢护筒坐标(m)", "")))
        # 基本规定
        code_var = widget_map.get("code_standard")
        if code_var:
            _std = proj.get("规范", "建筑基坑支护技术规程")
            code_var.set("《建筑基坑支护技术规程》（JGJ 120-2012）" if _std == "建筑基坑支护技术规程" else _std)
        safety_var = widget_map.get("safety_grade")
        if safety_var: safety_var.set(proj.get("围堰安全等级", "二级"))
        # 支护桩参数
        pile_type = proj.get("支护桩类型", "")
        if pile_type in ("钢板桩", "锁扣钢管桩"):
            Sheet_Pile_typevar.set(pile_type)
        # 截面编号(数字) → 截面名称
        seqmap = proj.get("section_seq_to_name", {})
        sec_ref = _as_string(proj.get("支护桩截面编号", ""))
        sec_name = ""
        try: sec_name = seqmap.get(int(sec_ref), "")
        except: sec_name = sec_ref
        secs = proj.get("sections", {})
        # 锁扣钢管桩参数
        if pile_type == "锁扣钢管桩":
            section_dict = secs.get(sec_name, {})
            section_params = section_dict.get("params", {})
            d_val = section_params.get("D") or 0
            t_val = section_params.get("t") or 0
            lock_val = section_params.get("锁扣宽度", 35)
            cnt_val = section_params.get("钢板桩数量", 1)
            w_val = section_params.get("钢板桩宽度", 600)
            SKGGZ_D.set(d_val)
            SKGGZ_t.set(t_val)
            SKGGZ_LockWidth.set(lock_val)
            SKGGZ_SheetCount.set(cnt_val)
            SKGGZ_SheetWidth.set(w_val)
            SKGGZ_Gap.set(calc_skggz_center_spacing(d_val, lock_val, cnt_val, w_val))
        elif pile_type == "钢板桩" and secs and sec_name in secs:
            section_params = secs[sec_name].get("params", {})
            excel_key_order = [
                "H(mm)", "B(mm)", "As(mm^2)", "Asy(mm^2)", "Asz(mm^2)",
                "Ixx(mm^4)", "Iyy(mm^4)", "Izz(mm^4)",
                "Cyp(mm)", "Cym(mm)", "Czp(mm)", "Czm(mm)",
                "Qyb(mm^2)", "Qzb(mm^2)", "Peri:O(mm)", "Peri:I(mm)",
                "Cent:y(mm)", "Cent:z(mm)", 
                "y1(mm)", "z1(mm)", "y2(mm)", "z2(mm)", "y3(mm)", "z3(mm)", "y4(mm)", "z4(mm)",
                "Zyy(mm^3)", "Zzz(mm^3)",
            ]
            sec_vals = [section_params.get(k, 0) for k in excel_key_order]
            if any(sec_vals):
                Pile_SEC_info_dict["SEC"] = sec_vals
        pile_change = sheet_refs.get("_on_section_frame_change")
        if pile_change: pile_change()
        Sheet_Pile_namevar.set(sec_name if sec_name else "自定义")
        # 记录当前截面名称，供详细设置对话框使用
        if sheet_refs and sec_name:
            sheet_refs["_last_gbz_section"] = sec_name if pile_type == "钢板桩" else ""
            sheet_refs["_last_skggz_section"] = sec_name if pile_type == "锁扣钢管桩" else ""
        # 承台边距加载后自动计算
        X_offset.set(_as_float(proj.get("承台长边预留边距(m)", 0.8)))
        Y_offset.set(_as_float(proj.get("承台宽边预留边距(m)", 0.8)))
        auto_calculate_pile_offsets(Cap_X, Cap_Y, X_offset, Y_offset,
                                     Sheet_Pile_typevar, Pile_SEC_info_dict,
                                     SKGGZ_D, SKGGZ_Gap)
        # 更新下拉选项
        if sheet_refs:
            gbz_combo = sheet_refs.get("_combo_gbz")
            if gbz_combo:
                gbz_combo.configure(values=sheet_refs.get("_gbz_names", ["自定义"]))
            skggz_combo = sheet_refs.get("_combo_skggz")
            if skggz_combo:
                skggz_opts = sheet_refs.get("_skggz_names", ["自定义"])
                skggz_combo.configure(values=skggz_opts)
                if pile_type == "锁扣钢管桩" and sec_name in skggz_opts:
                    skggz_combo.set(sec_name)
        CofferDam_Top_Level.set(_as_float(proj.get("围堰顶标高", 0)))
        CofferDam_L.set(_as_float(proj.get("围堰高", 0)))
        # 围檩/支撑
        num_waler = 0
        for i in range(1, 6):
            waler_spacing = _as_string(proj.get(f"围檩{i}间距(m)"))
            if not waler_spacing or waler_spacing in ("", "/"): continue
            waler_sec_long = _as_string(proj.get(f"围檩{i}长边截面"))
            if not waler_sec_long or waler_sec_long in ("", "/"): continue
            num_waler = i
        cur_cnt = len(Waler_rows)
        if num_waler > cur_cnt:
            for _ in range(num_waler - cur_cnt): _waler_actions['add']()
        elif num_waler < cur_cnt:
            for _ in range(cur_cnt - num_waler): _waler_actions['del']()
        for i in range(num_waler):
            n = i + 1
            Waler_rows[i]['entry'].set(_as_string(proj.get(f"围檩{n}间距(m)", "")))
            Waler_rows[i]['combo_long'].set(_as_string(proj.get(f"围檩{n}长边截面", "")))
            Waler_rows[i]['combo_short'].set(_as_string(proj.get(f"围檩{n}宽边截面", "")))
            Waler_rows[i]['mat_long'].set(_as_string(proj.get(f"围檩{n}长边材质", "")))
            Waler_rows[i]['mat_short'].set(_as_string(proj.get(f"围檩{n}宽边材质", "")))
            Waler_rows[i]['combo2'].set(_as_string(proj.get(f"内支撑{n}对撑截面", "/")))
            Waler_rows[i]['combo3'].set(_as_string(proj.get(f"内支撑{n}斜撑截面", "/")))
            Waler_rows[i]['mat_dc'].set(_as_string(proj.get(f"内支撑{n}对撑材质", "")))
            Waler_rows[i]['mat_xc'].set(_as_string(proj.get(f"内支撑{n}斜撑材质", "")))
            # 支撑布置坐标
            if i in Strut_blocks:
                Strut_blocks[i]['DC_X'].set(_as_string(proj.get(f"内支撑{n}对撑长边布置(m)", "/")))
                Strut_blocks[i]['DC_Y'].set(_as_string(proj.get(f"内支撑{n}对撑短边布置(m)", "/")))
                Strut_blocks[i]['XC_X'].set(_as_string(proj.get(f"内支撑{n}斜撑长边布置(m)", "/")))
                Strut_blocks[i]['XC_Y'].set(_as_string(proj.get(f"内支撑{n}斜撑短边布置(m)", "/")))
        # 切换项目后刷新围檩/内支撑材质下拉选项（随项目钢结构规范变化）
        _refresh_mat_opts = sheet_refs.get("_refresh_waler_material_options")
        if _refresh_mat_opts:
            _refresh_mat_opts()
        # 切换项目后刷新垫层/封底混凝土等级下拉选项（随项目混凝土规范变化）
        _refresh_grade_opts = sheet_refs.get("_refresh_concrete_grade_options")
        if _refresh_grade_opts:
            _refresh_grade_opts()
        # 更新牛腿+连接定义的启用状态
        update_waler_state = sheet_refs.get("_update_waler_dependent_state")
        if update_waler_state:
            num_strut = sum(1 for r in Waler_rows[:num_waler]
                            if r.get("combo2", tb.StringVar()).get() not in ("", "/")
                            or r.get("combo3", tb.StringVar()).get() not in ("", "/"))
            update_waler_state(num_waler, num_strut)
        # 封底/垫层
        blinding_t = _as_float(proj.get("垫层厚度(m)", 0))
        plug_t = _as_float(proj.get("封底厚度(m)", 0))
        Concrete_Blinding_thickness_var.set(blinding_t)
        Concrete_Plug_thickness_var.set(plug_t)
        Concrete_Blinding_grade_var.set(_as_string(proj.get("垫层混凝土等级", "")))
        Concrete_Plug_grade_var.set(_as_string(proj.get("封底混凝土等级", "")))
        Concrete_Blinding_check_var.set(blinding_t > 0)
        Concrete_Plug_check_var.set(plug_t > 0)
        blinding_widgets = widget_map.get("concrete_blinding", [])
        plug_widgets = widget_map.get("concrete_plug", [])
        for w in blinding_widgets:
            try: w.configure(state="normal" if blinding_t > 0 else "disabled")
            except: pass
        for w in plug_widgets:
            try: w.configure(state="normal" if plug_t > 0 else "disabled")
            except: pass
        plug_d_raw = _as_string(proj.get("封底支撑间距(m)", "0,0"))
        try:
            parts = plug_d_raw.replace("，", ",").split(",")
            Plug_support_long_var.set(float(parts[0].strip()) if parts[0].strip() else 0.0)
            if len(parts) > 1:
                Plug_support_short_var.set(float(parts[1].strip()) if parts[1].strip() else 0.0)
        except:
            Plug_support_long_var.set(0.0)
            Plug_support_short_var.set(0.0)
        Plug_support_d_var.set(Plug_support_long_var.get())
        Plug_bond_var.set(_as_float(proj.get("钢与混凝土黏聚力(kPa)", 150)))
        Plug_casing_count_var.set(_as_float(proj.get("护筒数量", 8)))
        Plug_casing_diam_var.set(_as_float(proj.get("护筒直径(m)", 1.2)))
        # 加权平均土层参数加载：字段为有效的 "γ,c,φ" 表示采用；为 "/" 或空表示不采用(0.0,0.0,0.0)
        if sheet_refs:
            w_zd = sheet_refs.get("_weighted_重度")
            w_nj = sheet_refs.get("_weighted_黏聚力")
            w_nm = sheet_refs.get("_weighted_内摩擦角")
            w_check = sheet_refs.get("_weighted_check")
            wavg_raw = _as_string(proj.get("加权平均重度+粘聚力+摩擦角", "/"))
            adopted = False
            wg = wc = wf = 0.0
            if wavg_raw and wavg_raw not in ("", "/"):
                parts = [p.strip() for p in wavg_raw.replace("，", ",").split(",")]
                if len(parts) >= 3:
                    try:
                        wg, wc, wf = float(parts[0]), float(parts[1]), float(parts[2])
                        adopted = True
                    except ValueError:
                        adopted = False
            if w_check: w_check.set(adopted)
            if w_zd: w_zd.set(wg)
            if w_nj: w_nj.set(wc)
            if w_nm: w_nm.set(wf)
        #
        # 连接定义：从Excel读取连接形式列表，更新到Boundary_Setting_GUI
        conns = proj.get("connections", [])
        _boundary_conn_data.clear()
        _boundary_conn_data.extend(conns)
        _boundary_conn_names.clear()
        _boundary_conn_names.extend([c.get("name", "") for c in conns if c.get("name")])
        # 读取主表中该项目的4个连接类型（牛腿与支护桩/牛腿与围檩/围檩与支护桩/围檩与内支撑）
        # 加载第一个连接的刚度值作为默认
        if conns:
            c = conns[0]
            SDx.set(_as_string(c.get("SDx", "1e7"))); SDy.set(_as_string(c.get("SDy", "1e6")))
            SDz.set(_as_string(c.get("SDz", "1e6"))); SRx.set(_as_string(c.get("SRx", "1e-6")))
            SRy.set(_as_string(c.get("SRy", "1e-6"))); SRz.set(_as_string(c.get("SRz", "1e-6")))
            NSDx.set(_as_string(c.get("NSDx", "")))
        # 更新连接定义 UI（4个 ConnectionSettingFrame）
        conn_frames = sheet_refs.get("conn_frames", [])
        conn_defs = sheet_refs.get("conn_defs", [])
        conn_options = [c.get("name", "") for c in conns[1:] if c.get("name")]  # 跳过第一项模板值
        for i, cdef in enumerate(conn_defs):
            if i >= len(conn_frames):
                break
            frame = conn_frames[i]
            frame.set_conn_options(conn_options)
            frame.set_conns_data(conns)
            conn_name = _as_string(proj.get(cdef["key"], ""))
            if conn_name:
                frame.set_type(conn_name)
            elif cdef["node_only"]:
                frame.set_type("共节点")
            else:
                frame.set_type("自定义")
            matched = next((c for c in conns if c.get("name") == conn_name), None)
            if matched:
                frame.set_stiffness(matched)
        # 仅当项目已有土层数据时加载，新建项目不填充
        soil_ref = _as_string(proj.get("地质参数", ""))
        existing_soil = proj.get("soil_data")
        _cofferdam_soil_original[proj_key] = copy.deepcopy(existing_soil)
        if existing_soil:
            # 剥离 "1," 前缀，避免保存时重复拼接
            if "," in soil_ref:
                parts = soil_ref.split(",", 1)
                soil_ref_clean = parts[1].strip()
            else:
                soil_ref_clean = soil_ref
            sheet_refs["cofferdam_soil_ref"] = soil_ref_clean
        else:
            # 无土层数据（新建项目），清空土层区域
            proj["soil_data"] = {}
            sheet_refs["cofferdam_soil_ref"] = ""
            Steel_Sheet_Pile_CofferDam_Load_dict['Excel'] = {}
            soil_frame = sheet_refs.get("soil_frame")
            if soil_frame:
                soil_frame.load_data(None)
        if existing_soil:
            Steel_Sheet_Pile_CofferDam_Load_dict['Excel'] = {
                f'第{i+1}层土': {
                    '土层名称': k,
                    '层厚': float(v[0]) if len(v) > 0 and _as_string(v[0]) not in ("", "/") else 0.0,
                    '重度': float(v[1]) if len(v) > 1 and _as_string(v[1]) not in ("", "/") else 0.0,
                    '黏聚力': float(v[2]) if len(v) > 2 and _as_string(v[2]) not in ("", "/") else 0.0,
                    '内摩擦角': float(v[3]) if len(v) > 3 and _as_string(v[3]) not in ("", "/") else 0.0,
                    '计算方法': v[4] if len(v) > 4 and _as_string(v[4]) not in ("", "/") else "",
                    '渗透系数(m/d)': v[5] if len(v) > 5 and _as_string(v[5]) not in ("", "/") else "",
                    '土层顶承压水头(m)': v[6] if len(v) > 6 and _as_string(v[6]) not in ("", "/") else "",
                } for i, (k, v) in enumerate(existing_soil.items())
            }
            soil_frame = sheet_refs.get("soil_frame")
            if soil_frame:
                _load_soil_into_frame(soil_frame, Steel_Sheet_Pile_CofferDam_Load_dict)
        # 判断土层sheet归属：sheet名==项目key则为自有，否则为引用
        if soil_ref and soil_ref == _current_proj_key[0]:
            _cofferdam_soil_owner[_current_proj_key[0]] = True
        else:
            _cofferdam_soil_owner[_current_proj_key[0]] = False
        _cofferdam_soil_ref = soil_ref
        # 记录原始加载数据（用于检测用户是否修改）
        _cofferdam_soil_original[_current_proj_key[0]] = dict(existing_soil) if existing_soil else {}
        # 刷新图示，使土层信息立即显示
        diagram.refresh()
        # 工况
        Drawdown_height.set(_as_float(proj.get("超挖深度(m)", 0)))
        Waterdown_height.set(_as_float(proj.get("水面至基坑底距离(m)", 0)))
        Excavation_face_dewater.set(_as_float(proj.get("开挖面降水高度(m)", 0)))
        # 辅助换撑/加撑数据加载（需先于施工阶段表格，供 dropdown 选项使用）
        if sheet_refs is not None:
            sheet_refs["assist_replace_data"] = _parse_assist_bracket_data(_as_string(proj.get("辅助换撑", "")))
            sheet_refs["assist_add_data"] = _parse_assist_add_data(
                _as_string(proj.get("辅助加围檩i", "")),
                _as_string(proj.get("辅助加内支撑i", "")))
        # 施工阶段表格数据加载
        stage_controller = sheet_refs.get("construct_stage") if sheet_refs else None
        if stage_controller and hasattr(stage_controller, "load_from_stage_dict"):
            stage_name = _as_string(proj.get("各工况名称", ""))
            stage_soil_top = _as_string(proj.get("各工况基坑内土顶标高(m)", ""))
            stage_water = _as_string(proj.get("各工况基坑内水面标高(m)", ""))
            stage_support = _as_string(proj.get("各工况支撑变动", ""))
            if stage_name:
                stage_dict = parse_stage_data(stage_name, stage_soil_top, stage_water, stage_support)
                stage_controller.load_from_stage_dict(stage_dict)
            else:
                stage_controller.load_from_stage_dict({})
        Spring_Thickness.set(_as_float(proj.get("土弹簧分层厚度(m)", 0.5)))
        Consider_Stress_Path.set(_as_string(proj.get("考虑土的应力路径", "否")))
        # 荷载参数数据加载
        sk_dict = sheet_refs.get("sigma_k_dict")
        if sk_dict is not None:
            _parse_surcharge_into_dict(proj, sk_dict)
        # 连接定义名称提取
        _boundary_conn_names.clear()
        _boundary_conn_names.extend([c.get("name", "") for c in conns if c.get("name")])
        # 刷新荷载参数 tab
        surcharge_refresh = sheet_refs.get("_surcharge_refresh")
        if surcharge_refresh:
            surcharge_refresh()
        # 工况分析：根据 是否生成施工阶段 切换整体模型/施工阶段
        has_stages = _as_string(proj.get("是否生成施工阶段", "是"))
        if has_stages == "否":
            none_construction_stage_check_var.set(True)
            construction_stage_check_var.set(False)
        else:
            none_construction_stage_check_var.set(False)
            construction_stage_check_var.set(True)
        # 记录项目分项系数，供生成荷载组合表使用
        if sheet_refs is not None:
            sheet_refs['combo_factors'] = {
            'basic_partial': proj.get("基本组合分项系数", "1.25"),
            'basic_combo': proj.get("基本组合组合值系数", "1.0"),
            'standard_partial': proj.get("标准组合分项系数", "1.0"),
            'standard_combo': proj.get("标准组合组合值系数", "1.0"),
        }
        # 触发 checkbox 的切换逻辑
        stf = sheet_refs.get("switch_table_from_project")
        if stf:
            stf()
        if sheet_refs.get("refresh_load_combo"):
            sheet_refs.get("refresh_load_combo")(force=True)
        corbel_long_num.set(_as_string(proj.get("牛腿长边布置数量","2")))
        corbel_short_num.set(_as_string(proj.get("牛腿短边布置数量","2")))
        Pile_Bottom_Boundary.set(_as_string(proj.get("支护桩底边界", "111000")))
        # 切项目后刷新缓存+施工阶段图示
        for _rk in ["refresh_basic", "refresh_substructure"]:
            _rf = sheet_refs.get(_rk)
            if _rf: _rf()
        sd = sheet_refs.get("stage_diagram") if sheet_refs else None
        if sd is not None and hasattr(sd, "draw"):
            try: sd.draw(0)
            except: pass

        _switching_flag[0] = False

    # ---- 下拉刷新 ----
    def _refresh_project_names():
        """刷新项目名称下拉选项。

        Returns:
            list[str]: 项目名称列表。
        """
        names = list(project_groups.keys())
        project_name_combo['values'] = names
        return names

    def _refresh_project_nos(pname):
        """根据项目名称刷新编号下拉选项并默认选中第一个。

        Args:
            pname (str): 项目名称。

        Returns:
            list[str]: 该名称下的围堰编号列表。
        """
        nos = project_groups.get(pname, [])
        project_no_combo['values'] = nos
        if nos:
            project_no_combo.set(nos[0])
        else:
            project_no_combo.set("")
        return nos

    def _on_project_name_change(*_):
        """项目名称下拉变化时刷新编号并切换到对应项目。

        Args:
            *_: 事件参数（忽略）。

        Returns:
            None
        """
        if _switching_flag[0]: return
        pname = project_name_combo.get()
        if not pname:
            return
        nos = _refresh_project_nos(pname)
        if nos:
            proj_key = f"{pname}{nos[0]}"
        else:
            proj_key = pname
        _switch_project(proj_key)

    def _on_project_no_change(*_):
        """围堰编号下拉变化时切换到对应项目。

        Args:
            *_: 事件参数（忽略）。

        Returns:
            None
        """
        pname = project_name_combo.get()
        pno = project_no_combo.get()
        if not pname:
            return
        proj_key = f"{pname}{pno}" if pno else pname
        if proj_key in projects_dict:
            _switch_project(proj_key)

    def _init_project_selection():
        """初始化项目下拉：选中 last_proj_key 对应项目并加载数据。

        Returns:
            None
        """
        if not _current_proj_key[0]:
            return
        proj = projects_dict.get(_current_proj_key[0], {})
        pname = proj.get("项目名称", "")
        pno = proj.get("围堰编号", "")
        # 刷新名称下拉并选中
        _refresh_project_names()
        if pname in project_groups:
            project_name_combo.set(pname)
            _refresh_project_nos(pname)
            if pno:
                project_no_combo.set(pno)
        # 显式触发加载（ComboboxSelected 事件在程序设置值时不触发）
        _switch_project(_current_proj_key[0])

    def _add_project(template_data=None):
        """弹出对话框新增项目或另存为副本，并切换到新项目。

        Args:
            template_data (dict|None): None 表示新增（用 PROJECT_TEMPLATE）；
                否则以此为模板复制生成新项目。

        Returns:
            None
        """
        is_save_as = template_data is not None
        dialog = tb.Toplevel(root)
        dialog.title("另存为" if is_save_as else "新增项目")
        dialog.geometry("380x250+%d+%d" % (root.winfo_rootx()+80, root.winfo_rooty()+80))
        dialog.transient(root); dialog.grab_set()

        # 项目名称：可编辑 Combobox
        tb.Label(dialog, text="项目名称:", width=12).pack(padx=10, pady=(15, 5), anchor=W)
        existing_names = list(project_groups.keys())
        current_name = project_name_combo.get() if project_name_combo.get() else ""
        preset_name = template_data.get("项目名称", current_name) if is_save_as else current_name
        name_var = tb.StringVar(value=preset_name)
        name_combo = tb.Combobox(dialog, textvariable=name_var, values=existing_names, width=28)
        name_combo.pack(padx=10, pady=5, fill=X)

        # 编号：手动输入
        tb.Label(dialog, text="编号:", width=12).pack(padx=10, pady=5, anchor=W)
        no_var = tb.StringVar()
        no_entry = tb.Entry(dialog, textvariable=no_var, width=28)
        no_entry.pack(padx=10, pady=5, fill=X)

        def confirm():
            """校验名称/编号并创建或复制项目，随后刷新下拉并切换。

            Returns:
                None
            """
            raw_name = name_var.get().strip()
            raw_no = no_var.get().strip()
            if not raw_name:
                messagebox.showwarning("提示", "项目名称不能为空", parent=dialog); return
            if not raw_no:
                messagebox.showwarning("提示", "围堰编号不能为空", parent=dialog); return
            combined_key = f"{raw_name}{raw_no}"
            if combined_key in projects_dict:
                messagebox.showwarning("提示", f"项目 '{combined_key}' 已存在", parent=dialog); return
            if is_save_as:
                new_proj = copy.deepcopy(template_data)
                new_proj.update({"项目名称": raw_name, "围堰编号": raw_no})
            else:
                new_proj = dict(PROJECT_TEMPLATE)
                new_proj.update({"项目名称": raw_name, "围堰编号": raw_no})
                # 新建项目复制全局连接形式列表
                if global_connections:
                    new_proj["connections"] = copy.deepcopy(global_connections)
                # 新项目继承截面库与编号映射，便于把模板默认编号解析为标准截面（如1->SP-Ⅳ）
                if projects_dict:
                    _src_proj = next(iter(projects_dict.values()))
                    new_proj["sections"] = dict(_src_proj.get("sections", {}))
                    new_proj["section_seq_to_name"] = dict(_src_proj.get("section_seq_to_name", {}))
            projects_dict[combined_key] = new_proj
            _proj_keys_in_order.append(combined_key)
            # 更新 project_groups
            if raw_name not in project_groups:
                project_groups[raw_name] = []
            if raw_no and raw_no not in project_groups[raw_name]:
                project_groups[raw_name].append(raw_no)
            _current_proj_key[0] = combined_key
            _refresh_project_names()
            project_name_combo.set(raw_name)
            _refresh_project_nos(raw_name)
            if raw_no:
                project_no_combo.set(raw_no)
            _switch_project(combined_key)
            dialog.destroy()

        btn_frame = tb.Frame(dialog)
        btn_frame.pack(fill=X, padx=10, pady=10)
        tb.Button(btn_frame, text="确定", bootstyle=SUCCESS, command=confirm).pack(side=RIGHT, padx=5)
        tb.Button(btn_frame, text="取消", bootstyle=SECONDARY, command=dialog.destroy).pack(side=RIGHT, padx=5)
        name_combo.focus()
        dialog.wait_window()

    def _save_as_project():
        """以当前项目为模板另存为一份新副本。

        Returns:
            None
        """
        old_key = _current_proj_key[0]
        if not old_key or old_key not in projects_dict:
            return
        _add_project(template_data=projects_dict[old_key])

    def _delete_project():
        """删除当前项目（含下拉与分组），并把剩余项目立即写回 Excel。

        Returns:
            None
        """
        key = _current_proj_key[0]
        if not key:
            return
        if not messagebox.askyesno("确认删除", f"确定删除项目 '{key}' 吗？", parent=root):
            return
        proj = projects_dict.pop(key)
        if key in _proj_keys_in_order:
            _proj_keys_in_order.remove(key)
        del_name = proj.get("项目名称", "")
        del_no = proj.get("围堰编号", "")
        # 从 project_groups 中移除
        if del_name in project_groups:
            if del_no in project_groups[del_name]:
                project_groups[del_name].remove(del_no)
            if not project_groups[del_name]:
                del project_groups[del_name]
        # 刷新下拉
        _refresh_project_names()
        if project_groups:
            first_name = list(project_groups.keys())[0]
            project_name_combo.set(first_name)
            _refresh_project_nos(first_name)
            first_no = project_no_combo.get()
            proj_key = f"{first_name}{first_no}" if first_no else first_name
            _current_proj_key[0] = proj_key
            _switch_project(proj_key)
        else:
            _current_proj_key[0] = None
        # 立即写入 Excel（删除行 + 重写剩余项目）
        excel_path = get_excel_path()
        if excel_path:
            save_to_excel(excel_path, projects_dict,
                          project_keys_in_order=_proj_keys_in_order,
                          soil_owner=_cofferdam_soil_owner,
                          soil_data_cache=_cofferdam_soil_data_cache)

    # ---- 绑定 ----
    project_name_combo.bind("<<ComboboxSelected>>", _on_project_name_change)
    project_no_combo.bind("<<ComboboxSelected>>", _on_project_no_change)
    proj_add_btn.config(command=_add_project)
    proj_saveas_btn.config(command=_save_as_project)
    proj_save_btn.config(command=param_save)
    proj_del_btn.config(command=_delete_project)
    _init_project_selection()

    # ---- 注册控件高亮组 ----
    for group, widgets in widget_map.items():
        if group in ("_waler_entries", "code_standard", "safety_grade"):
            continue
        elif isinstance(widgets, (tk.StringVar, tk.BooleanVar, tk.DoubleVar)):
            continue
        elif isinstance(widgets, list):
            diagram.register_group(group, widgets)
        else:
            diagram.register_group(group, widgets)

    # 牛腿长边/短边高亮联动
    if "corbel_long_entry" in sheet_refs:
        diagram.register_group("corbel_long", sheet_refs["corbel_long_entry"])
    if "corbel_short_entry" in sheet_refs:
        diagram.register_group("corbel_short", sheet_refs["corbel_short_entry"])

    # 混凝土垫层/封底勾选联动图示
    Concrete_Blinding_check_var.trace_add('write', lambda *_: diagram.refresh())
    Concrete_Plug_check_var.trace_add('write', lambda *_: diagram.refresh())

    # ---- 侧边栏导航按钮 ----
    nav_items = [
        ("◪  基本设置", "设置地面标高、承台尺寸、支护桩等基本参数"),
        ("⧓  支撑及封底", "配置围檩、内支撑、垫层及封底"),
        ("▦  土层信息", "编辑土层信息与土弹簧分层厚度"),
        ("▥  边界定义", "设置牛腿与围堰构件之间的弹性连接参数"),
        ("⇊  荷载参数", "设置均布/矩形/条形附加荷载"),
        ("▧  工况分析", "定义施工阶段与荷载组合"),
    ]

    nav_btns = []
    nav_indicators = []

    for i, (label, tooltip) in enumerate(nav_items):
        btn_frame = tk.Frame(sidebar, bg=SIDEBAR_IDLE_BG, height=52)
        btn_frame.pack(fill=X, pady=1)
        btn_frame.pack_propagate(False)

        indicator = tk.Frame(btn_frame, bg=SIDEBAR_IDLE_BG, width=4)
        indicator.pack(side=LEFT, fill=Y)

        btn = tk.Button(btn_frame, text=label,
                        font=("Microsoft YaHei UI", 11, "bold"),
                        bg=SIDEBAR_IDLE_BG, fg=SIDEBAR_IDLE_FG,
                        activebackground=SIDEBAR_HOVER, activeforeground="white",
                        bd=0, anchor=W, padx=14, pady=14,
                        cursor="hand2", compound=LEFT)
        btn.pack(side=LEFT, fill=BOTH, expand=True)
        createToolTip(btn, tooltip)
        createToolTip(btn_frame, tooltip)
        nav_btns.append(btn)
        nav_indicators.append(indicator)

    def switch_page(idx):
        """切换到指定页面，并同步滚动条、图示显隐与导航高亮。

        Args:
            idx (int): 页面索引（0~5，5 为工况分析页）。

        Returns:
            None
        """
        _current_page_idx[0] = idx

        # 仅在确实有高亮时才清空并重绘（避免无谓的 matplotlib 重绘卡顿）
        has_highlight = bool(getattr(diagram, '_last_focused_group', ''))
        skggz_diagram = sheet_refs.get("skggz_diagram")
        if skggz_diagram and getattr(skggz_diagram, '_last_focused_group', ''):
            has_highlight = True

        if has_highlight:
            diagram.clear_highlight()
            if skggz_diagram:
                skggz_diagram._last_focused_group = ""
                skggz_diagram.draw()

        # 滚动条：工况分析页隐藏，内容容器锁视口；其他页恢复自然尺寸
        if idx == 5:
            content_scrollbar.pack_forget()
        else:
            if not content_scrollbar.winfo_ismapped():
                content_scrollbar.pack(side=RIGHT, fill=Y)
            content_canvas.itemconfig(content_window, width=content_canvas.winfo_width(), height=0)
            content_canvas.configure(scrollregion=content_canvas.bbox("all"))

        # 图示显隐 (针对 PanedWindow 的特殊处理)
        current_panes = [str(p) for p in workspace.panes()]
        was_hidden = (str(image_container) not in current_panes)
        if idx in (4, 5):
            if not was_hidden:
                workspace.forget(image_container)
        else:
            if was_hidden:
                workspace.add(image_container, stretch="never")
                workspace.paneconfig(image_container, minsize=IMG_MINSIZE)
                workspace.paneconfig(top_half, minsize=int(win_height * 0.35))
                workspace.update_idletasks()
                total_h = workspace.winfo_height()
                try:
                    workspace.sash_place(0, 0, total_h - IMG_MINSIZE)
                except tk.TclError:
                    pass
            if was_hidden or has_highlight:
                diagram.canvas.draw_idle()

        # 切换页面
        for pg in pages.values():
            pg.pack_forget()
        pages[idx].pack(fill=BOTH, expand=True)
        # 荷载参数页：延迟重绘加载示意图
        if idx == 4:
            _rf = sheet_refs.get("_surcharge_refresh")
            if _rf:
                root.after(200, _rf)
        # 工况分析页：切回时自动智能刷新荷载组合 + 固定视口高度
        if idx == 5:
            _refresh_combo = sheet_refs.get("refresh_load_combo")
            if _refresh_combo:
                _refresh_combo()
            _fix_viewport_height()
            root.after(200, _fix_viewport_height)
        # 更新导航样式
        for j, (btn, ind) in enumerate(zip(nav_btns, nav_indicators)):
            btn_frame = btn.master
            if j == idx:
                btn_frame.config(bg=SIDEBAR_HOVER)
                btn.config(bg=SIDEBAR_ACTIVE, fg="white")
                ind.config(bg="#60a5fa")
            else:
                btn_frame.config(bg=SIDEBAR_IDLE_BG)
                btn.config(bg=SIDEBAR_IDLE_BG, fg=SIDEBAR_IDLE_FG)
                ind.config(bg=SIDEBAR_IDLE_BG)
            # 切页后关闭所有导航按钮的悬停提示
            if hasattr(btn, '_tooltip'):
                btn._tooltip.hidetip()
            if hasattr(btn_frame, '_tooltip'):
                btn_frame._tooltip.hidetip()

    for i in range(len(nav_items)):
        nav_btns[i].config(command=lambda idx=i: switch_page(idx))
    switch_page(0)
    # 初始图示定位到固定大小
    def _init_img():
        """初始将图示区定位为固定高度（通过调整分隔条位置）。

        Returns:
            None
        """
        workspace.update_idletasks()
        total_h = workspace.winfo_height()
        if total_h > IMG_MINSIZE:
            try:
                workspace.sash_place(0, 0, total_h - IMG_MINSIZE)
            except tk.TclError:
                pass
    root.after(340, _init_img)

    # 文件保存路径
    mct_savepath = tb.StringVar(value=os.path.join(Applocation, r"Custom\CofferDam_Untitled.mct"))
    save_frame = tb.Frame(path_inner)
    save_frame.pack(side=LEFT, fill=X, expand=True)
    tb.Label(save_frame, text="保存路径:", bootstyle=SECONDARY, width=10).pack(side=LEFT)
    path_entry = tb.Entry(save_frame, textvariable=mct_savepath, state="readonly", width=60)
    path_entry.pack(side=LEFT, fill=X, expand=True, padx=5)
    path_btn = tb.Button(save_frame, text="更改路径", command=lambda: mct_path_SaveAs(mct_savepath),
                          bootstyle=(OUTLINE, PRIMARY), width=10)
    path_btn.pack(side=LEFT, padx=5)

    # 操作按钮（右侧）
    btn_frame = tb.Frame(bottom_inner)
    btn_frame.pack(side=RIGHT)

    # 模型类型复选按钮（左侧，独立可多选）
    gen_mcb_var = tb.BooleanVar(value=True)
    gen_qiaotong_var = tb.BooleanVar(value=False)
    check_frame = tb.Frame(bottom_inner)
    check_frame.pack(side=LEFT)
    cb_mcb = tb.Checkbutton(check_frame, text='生成mcb模型', variable=gen_mcb_var, bootstyle=PRIMARY,
                            command=lambda: gen_mcb_var.set(gen_mcb_var.get()))
    cb_mcb.pack(side=LEFT, padx=10)
    cb_qiaotong = tb.Checkbutton(check_frame, text='生成桥通模型', variable=gen_qiaotong_var, bootstyle=PRIMARY,
                                command=lambda: gen_qiaotong_var.set(gen_qiaotong_var.get()))
    cb_qiaotong.pack(side=LEFT, padx=10)
    # invoke() 强制 toggle → 触发主题视觉刷新 → 延迟恢复原始值
    _orig_mcb = gen_mcb_var.get()
    _orig_qt = gen_qiaotong_var.get()
    cb_mcb.invoke()
    cb_qiaotong.invoke()
    root.after(50, lambda: gen_mcb_var.set(_orig_mcb))
    root.after(50, lambda: gen_qiaotong_var.set(_orig_qt))

    mct_btn = tb.Button(btn_frame, text='生成MCT', command=_mct_generate, bootstyle=SUCCESS, width=10)
    mct_btn.pack(side=LEFT, padx=10)

    exit_btn = tb.Button(btn_frame, text='退出', command=lambda: on_exit(root), bootstyle=DANGER, width=8)
    exit_btn.pack(side=RIGHT, padx=10)


def auto_calculate_pile_offsets(Cap_X, Cap_Y, X_offset, Y_offset,
                                 Sheet_Pile_typevar, Pile_SEC_info_dict,
                                 SKGGZ_D, SKGGZ_Gap):
    """根据承台尺寸和支护桩截面参数自动计算承台预留边距并回写 X/Y_offset。

    承台边距 = (围堰轮廓尺寸 - 承台尺寸) / 2；围堰轮廓尺寸按桩数量取整：
    钢板桩长边奇数桩、短边偶数桩；锁扣钢管桩按中心间距与直径计算。
    结果不大于 0.1m 时不更新。

    Args:
        Cap_X, Cap_Y (tk.DoubleVar): 承台长/宽（m）。
        X_offset, Y_offset (tk.DoubleVar): 长/短边预留边距（m），就地修改。
        Sheet_Pile_typevar (tk.StringVar): 支护类型。
        Pile_SEC_info_dict (dict): 支护桩截面信息 {'SEC': [...]}。
        SKGGZ_D (tk.DoubleVar): 锁扣钢管桩直径（mm）。
        SKGGZ_Gap (tk.DoubleVar): 锁扣钢管桩中心间距（mm）。

    Returns:
        None: 副作用为设置 X_offset/Y_offset。
    """
    try:
        cap_x_m = Cap_X.get()  # m
        cap_y_m = Cap_Y.get()  # m
        if cap_x_m <= 0 or cap_y_m <= 0:
            return
        cap_x_mm = cap_x_m * 1000
        cap_y_mm = cap_y_m * 1000
        min_clearance = 2000  # 两侧各1m最小施工净空(mm)

        pile_type = Sheet_Pile_typevar.get()
        if pile_type == "钢板桩":
            sec = Pile_SEC_info_dict.get("SEC", [])
            if len(sec) < 2 or sec[1] <= 0:
                return
            pile_width = sec[1]  # B (mm) — 钢板桩宽度

            # 长边（X）：奇数桩数量
            l_n = (cap_x_mm + min_clearance) / pile_width
            l_n = math.ceil(l_n)
            if l_n % 2 == 0:
                l_n += 1
            new_offset_x = round(((l_n * pile_width) - cap_x_mm) / 2 / 1000, 3)

            # 短边（Y）：偶数桩数量
            b_n = (cap_y_mm + min_clearance) / pile_width
            b_n = math.ceil(b_n)
            if b_n % 2 != 0:
                b_n += 1
            new_offset_y = round(((b_n * pile_width) - cap_y_mm) / 2 / 1000, 3)

        elif pile_type == "锁扣钢管桩":
            d = SKGGZ_D.get()       # 直径(mm)
            gap = SKGGZ_Gap.get()   # 中心间距(mm)
            if d <= 0 or gap <= 0:
                return

            # 长边（X）：围堰轮廓周长 = cap_x_mm + d + 2*min_clearance
            l_n = (cap_x_mm + d + min_clearance) / gap
            l_n = math.ceil(l_n)
            new_offset_x = round(((l_n * gap) - cap_x_mm - d) / 2 / 1000, 3)

            # 短边（Y）
            b_n = (cap_y_mm + d + min_clearance) / gap
            b_n = math.ceil(b_n)
            new_offset_y = round(((b_n * gap) - cap_y_mm - d) / 2 / 1000, 3)
        else:
            return

        if new_offset_x > 0.1:
            X_offset.set(new_offset_x)
        if new_offset_y > 0.1:
            Y_offset.set(new_offset_y)
    except (tk.TclError, ValueError, TypeError):
        pass


def CofferDam_FEM_MidasCivil_Main(parent=None, on_close=None):
    """钢板桩围堰建模程序入口：注册检查、创建主窗口并启动界面。

    Args:
        parent (tk.Widget|None): 父窗口；None 时创建独立 Tk 主窗口。
        on_close (callable|None): 窗口关闭回调。

    Returns:
        None
    """
    if not if_Reg(parent):
        if on_close:
            on_close()
        return

    FileName = 'ShuZhiQiaoShi'
    # Excel_xlsx_filename = 'CofferDam_Solid_Table.xlsx'

    Applocation = read_Register('Software\\{}'.format(FileName), 'Applocation')
    if parent == None:
        root = tb.Window("钢板桩围堰自动化建模程序", themename="litera")
    else:
        root = tb.Toplevel(parent)
        root.title("钢板桩围堰自动化建模程序")
    Steel_Pile_CofferDam_Dialog(root, Applocation)
    if parent == None:
        root.mainloop()
    elif on_close:
        def _on_root_destroy(event, _root=root):
            """窗口销毁时触发 on_close 回调（仅限根窗口自身）。

            Args:
                event: Tk Destroy 事件。
                _root: 根窗口引用（默认参数绑定）。

            Returns:
                None
            """
            if event.widget is _root:
                on_close()
        root.bind('<Destroy>', _on_root_destroy)


# if __name__ == "__main__":
#     CofferDam_FEM_MidasCivil_Main()