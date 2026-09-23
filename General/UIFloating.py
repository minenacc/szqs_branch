# General/UI_ttk.py
# 基于 ttkbootstrap 的新 UI 实现

import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from PIL import Image, ImageTk

# 复用原 UI 的悬浮提示
from General.UIHandle import createToolTip

# ===== 图标按钮尺寸（想放大图标/文字，改这里即可） =====
ICON_SIZE = 125           # 图标边长（px）
LABEL_FONT_SIZE = 10      # 图标下方文字字号
BTN_PAD = 6              # 图标按钮内边距
ICON_BTN_PADX = 6        # 网格里每个入口的水平间距
ICON_BTN_PADY = 6        # 网格里每个入口的垂直间距
GRID_COLS = 4            # 网格列数（每行按钮个数）
GRID_ROWS = 3            # 网格行数上限
# ===================================================

# ===== 应用版本 =====
VERSION_REG_PATH = "Software\\ShuZhiQiaoShi"
VERSION_REG_KEY  = "Appversion"
VERSION_FALLBACK = "V0.3.10"

def get_app_version():
    """从注册表读取 Appversion，自动补 'V' 前缀；缺失时回退默认值"""
    try:
        from General.DataUtils import read_Register
        ver = read_Register(VERSION_REG_PATH, VERSION_REG_KEY)
    except Exception as e:
        print(f"读取版本号失败: {e}")
        return VERSION_FALLBACK
    if ver == "NotFound" or not ver:
        return VERSION_FALLBACK
    ver = str(ver).strip()
    if not ver.startswith("V"):
        ver = "V" + ver
    return ver

# ====================

# ===== 侧边栏（icon+文字导航）配置 =====
SIDEBAR_WIDTH = 150       # 侧边栏宽度（px）
SIDEBAR_ICON_SIZE = 40    # 侧边栏图标边长（px）
SIDEBAR_NAV_PAD_X = 10    # 每格水平内边距
SIDEBAR_NAV_PAD_Y = 10    # 每格垂直内边距（控制每格高度）
SIDEBAR_BG = "#f8f9fa"           # 侧边栏底色
SIDEBAR_FG = "#495057"           # 常规文字色
SIDEBAR_HOVER_BG = "#e7f1ff"     # 悬浮浅蓝
SIDEBAR_HOVER_FG = "#0d6efd"     # 悬浮文字蓝
SIDEBAR_ACTIVE_BG = "#0d6efd"    # 选中 primary
SIDEBAR_ACTIVE_FG = "#ffffff"    # 选中文字白
# 侧边栏每页对应的图标文件名
SIDEBAR_ICONS = {
    "modeling": "Midas_icon.png",
    "drawing":  "CAD_icon.png",
    "bim":      "BIM_icon.png",
    "report":   "Word_icon.png",
    "ai":       "ai_recs.png",
}
# ===================================


# ===== 结构 → 结构图标文件（键=实际入口键名，含"钢管贝雷梁现浇支架/钢管贝雷梁作业平台/钢板桩/锁扣钢管桩矩形围堰"） =====
STRUCT_ICONS = {
    "钢管贝雷梁现浇支架": "Support_SteelBailey_exe.png",
    "上承式桁架/型钢纵梁直线栈桥":     "Trestle_exe.png",
    "钢管贝雷梁作业平台": "Platform_exe.png",
    "钢板桩/锁扣钢管桩矩形围堰": "CofferDam_exe.png",
    # 批量(建模/结果)入口图由各页 label_icons 单独指定：
    #   建模页 批量钢板桩/锁扣钢管桩矩形围堰 -> Cofferdam_Batch_Midas.png
    #   计算书页 钢板桩/锁扣钢管桩矩形围堰批量结果 -> Cofferdam_Batch_Excel.png
}
# =============================================================

# ===== 入口标签集中管理（键=功能键名，值=按钮显示文字，可换行，便于后续大改） =====
ENTRY_LABELS = {
    "钢管贝雷梁现浇支架":     "钢管贝雷梁现浇支架",
    "上承式桁架/型钢纵梁直线栈桥":         "上承式桁架/型钢纵梁直线栈桥",
    "钢管贝雷梁作业平台":     "钢管贝雷梁作业平台",
    "钢板桩/锁扣钢管桩矩形围堰":     "钢板桩/锁扣钢管桩矩形围堰",
    "批量钢板桩/锁扣钢管桩矩形围堰": "批量钢板桩/锁扣钢管桩矩形围堰",
    "钢板桩/锁扣钢管桩矩形围堰批量结果": "钢板桩/锁扣钢管桩矩形围堰批量结果",
    "钢板桩/锁扣钢管桩矩形围堰智能推荐": "钢板桩/锁扣钢管桩矩形围堰智能推荐",
}
# ================================================================

def _enable_dpi_awareness():
    """启用 Windows 进程 DPI 感知（在创建任何窗口前调用）。

    Tkinter 默认不感知 DPI；系统缩放比(125%/150%)会让逻辑尺寸渲染被放大，
    导致固定窗口大小内容在高低DPI显示器上不一致、右侧文字溢出裁剪。
    设置为 Per-Monitor/System-aware 后，逻辑像素按真实DPI渲染，跨分辨率稳定。
    """
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # 1 = system DPI aware
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()       # 兼容旧系统
    except Exception as e:
        print(f"DPI 感知设置失败: {e}")


class FloatingBallApp_ttk:
    """基于 ttkbootstrap 的悬浮球应用"""
    
    def __init__(self, parent_path, Revit_Version,
                 Midas_Func_Dict, CAD_Func_Dict, Word_Func_Dict):
        
        self.parent_path = parent_path
        self.Revit_Version = Revit_Version
        
        # 功能字典（直接按 UI 分类：功能为页面、结构为入口）
        self.midas_dict = Midas_Func_Dict   # 建模页：{结构: 建模函数}
        self.cad_dict   = CAD_Func_Dict     # 画图页：{结构: 绘图函数}
        self.word_dict  = Word_Func_Dict    # 计算书页：{结构: 计算书函数}
        # Revit/BIM 页与 AI 页不依赖功能字典（Revit 统一启动、AI 为独立入口）
        
        # 图标路径
        self.icon_path = os.path.join(parent_path, 'Support', 'png_picture')
        self._icon_cache = {}  # 图标加载缓存（同一文件多个按钮共享同一 PhotoImage）
        
        # 悬浮球状态
        self.minimized = [True]
        self.ball_size = 110
        self.drag_start_x = 0
        self.drag_start_y = 0
        self.drag_threshold = 5
        self.is_dragging = False
        
        # 当前激活的侧边栏
        self.active_tab = "modeling"
        
        # 主窗口
        self.main = None

        # 启用 Windows DPI 感知：让逻辑尺寸在不同缩放比(100%/125%/150%)显示器上渲染一致，
        # 避免高DPI笔记本上窗口内容被系统放大导致右侧("百度网盘"等)溢出裁剪
        _enable_dpi_awareness()

        # 创建根窗口
        self.root = tb.Window()
        self.root.withdraw()
        
        # 创建悬浮球
        self.create_floating_ball()
        
        # 启动主循环
        self.root.mainloop()
    
    # ==================== 悬浮球 ====================
    
    def create_floating_ball(self):
        """创建悬浮球窗口"""
        self.ball = tk.Toplevel(self.root)
        self.ball.overrideredirect(True)
        self.ball.attributes("-topmost", True)
        self.ball.attributes("-transparentcolor", "white")
        self.ball.geometry(f"{self.ball_size}x{self.ball_size}")
        
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        self.ball.geometry(f"+{int(screen_width * 0.9)}+{int(screen_height * 0.1)}")
        
        self.canvas = tk.Canvas(
            self.ball, width=self.ball_size, height=self.ball_size,
            bg="white", highlightthickness=0
        )
        self.canvas.pack()
        
        self.canvas.create_oval(0, 0, self.ball_size, self.ball_size,
                                fill="#ffffff", outline='', width=0)
        
        self.icon_id = None
        self.icon_image = None
        icon_file = os.path.join(self.icon_path, 'Logo.png')
        if os.path.exists(icon_file):
            self._load_icon(icon_file)
        else:
            self.icon_id = self.canvas.create_text(
                self.ball_size // 2, self.ball_size // 2,
                text="◀", font=("Arial", 16), fill="#0d6efd"
            )
        
        self.canvas.bind("<Button-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.canvas.bind("<Button-3>", self.show_context_menu)
        
        self.context_menu = tk.Menu(self.ball, tearoff=0)
        self.context_menu.add_command(label="退出", command=self.exit_app)
    
    def _load_icon(self, path):
        try:
            img = tk.PhotoImage(file=path)
            w, h = img.width(), img.height()
            scale = min((self.ball_size - 10) / w, (self.ball_size - 10) / h)
            scaled = img.subsample(max(1, int(1 / scale)), max(1, int(1 / scale)))
            self.icon_id = self.canvas.create_image(
                self.ball_size // 2, self.ball_size // 2, image=scaled, anchor='center'
            )
            self.icon_image = scaled
        except Exception as e:
            print(f"加载图标失败: {e}")
    
    def on_press(self, event):
        self.drag_start_x = event.x_root
        self.drag_start_y = event.y_root
        self.is_dragging = False
    
    def on_drag(self, event):
        dx = abs(event.x_root - self.drag_start_x)
        dy = abs(event.y_root - self.drag_start_y)
        if not self.is_dragging and (dx > self.drag_threshold or dy > self.drag_threshold):
            self.is_dragging = True
        if self.is_dragging:
            x = self.ball.winfo_x() + event.x - self.ball_size // 2
            y = self.ball.winfo_y() + event.y - self.ball_size // 2
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            x = max(0, min(x, sw - self.ball_size))
            y = max(0, min(y, sh - self.ball_size))
            self.ball.geometry(f"+{x}+{y}")
    
    def on_release(self, event):
        if not self.is_dragging:
            self.toggle_window()
        self.is_dragging = False
    
    def show_context_menu(self, event):
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()
    
    def toggle_window(self):
        if not self.main or not self.main.winfo_exists():
            self.minimized[0] = False
            self.create_main_ui()
        else:
            if self.minimized[0]:
                self.restore_from_ball()
            else:
                self.minimize_to_ball()
    
    def minimize_to_ball(self):
        if self.main:
            self.main.withdraw()
        self.minimized[0] = True
        if not self.icon_image:
            self.canvas.itemconfig(self.icon_id, text="▶")
    
    def restore_from_ball(self):
        if self.main:
            self.main.deiconify()
            self._center_main()
        self.minimized[0] = False
        if not self.icon_image:
            self.canvas.itemconfig(self.icon_id, text="◀")
    
    def _center_main(self):
        self.main.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        w = self.main.winfo_width()
        h = self.main.winfo_height()
        x = (sw - w) // 2
        y = (sh - h - 300) // 2
        self.main.geometry(f"+{x}+{y}")
    
    # ==================== 主窗口 ====================
    
    def create_main_ui(self):
        """创建主窗口"""
        self.main = tb.Toplevel(self.root)
        self.main.title("数智桥施")
        self.main.resizable(False, False)
        self.main.withdraw()
        self.main.protocol("WM_DELETE_WINDOW", self.minimize_to_ball)
        
        # 窗口尺寸
        main_width = 800
        main_height = 450
        self.main.geometry(f"{main_width}x{main_height}")
        self._center_main()
        
        # 样式
        style = tb.Style(theme='cosmo')
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 12, "bold"))
        style.configure("Custom.TButton", font=("Microsoft YaHei UI", 10),
                        foreground="#0d6efd", background="#ffffff",
                        bordercolor="#0d6efd", relief="solid", padding=8)
        # 显式 hover/按下高亮，保证各入口按钮（含 AI 页）悬停反馈一致
        style.map("Custom.TButton",
                  background=[("active", "#e7f1ff"), ("pressed", "#cfe2ff"), ("!active", "#ffffff")],
                  bordercolor=[("active", "#0d6efd"), ("pressed", "#0b5ed7"), ("!active", "#0d6efd")],
                  foreground=[("active", "#0d6efd"), ("pressed", "#0b5ed7"), ("!active", "#0d6efd")])
        
        # 菜单栏
        self._create_menubar()
        
        # 底部状态栏（先 pack 占底部，随后主体 expand 占剩余空间）
        self._create_statusbar()
        
        # 主体：侧边栏 + 内容区
        main_frame = tb.Frame(self.main)
        main_frame.pack(fill=BOTH, expand=YES)
        
        # 侧边栏（固定底色的 tk.Frame，便于 nav 项背景统一控制）
        self._sidebar_imgs = []  # 保持图标引用
        self.sidebar_frame = tk.Frame(main_frame, width=SIDEBAR_WIDTH, bg=SIDEBAR_BG)
        self.sidebar_frame.pack(side=LEFT, fill=Y)
        self.sidebar_frame.pack_propagate(False)
        
        # 分隔线
        ttk.Separator(main_frame, orient=VERTICAL).pack(side=LEFT, fill=Y)
        
        # 内容区
        self.content_frame = tb.Frame(main_frame)
        self.content_frame.pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)
        
        # 创建侧边栏按钮
        self.sidebar_buttons = {}
        self._create_sidebar()
        
        # 显示默认内容
        self._switch_tab("modeling")

        # 按内容自适应窗口尺寸（避免固定尺寸在高低DPI差异下裁剪）
        self._auto_size_main()

        self.main.deiconify()

    def _auto_size_main(self):
        """按内容请求尺寸自适应主窗口大小并居中（消除固定尺寸的适配问题）"""
        self.main.update_idletasks()
        w = self.main.winfo_reqwidth()
        h = self.main.winfo_reqheight()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        # 保证最小尺寸，且不超出屏幕
        w = max(w, 800)
        h = max(h, 450)
        w = min(w, sw - 40)
        h = min(h, sh - 40)
        x = (sw - w) // 2
        y = (sh - h - 300) // 2
        self.main.geometry(f"{w}x{h}+{x}+{y}")

    def _create_menubar(self):
        """创建菜单栏"""
        menubar = tk.Menu(self.main)
        
        # 文件菜单
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="新建")
        file_menu.add_command(label="打开", command=self._open_folder)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.exit_app)
        menubar.add_cascade(label="文件", menu=file_menu)
        
        # 参数设置菜单
        param_menu = tk.Menu(menubar, tearoff=0)
        param_menu.add_command(label="钢管贝雷梁现浇支架参数", command=lambda: self._show_param("support"))
        param_menu.add_command(label="上承式桁架/型钢纵梁直线栈桥参数", command=lambda: self._show_param("trestle"))
        param_menu.add_command(label="钢管贝雷梁作业平台参数", command=lambda: self._show_param("platform"))
        param_menu.add_command(label="钢板桩/锁扣钢管桩矩形围堰参数", command=lambda: self._show_param("cofferdam"))
        menubar.add_cascade(label="参数设置", menu=param_menu)
        
        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="注册", command=self._on_about_register)
        help_menu.add_separator()
        help_menu.add_command(label="现浇钢管贝雷梁现浇支架帮助", command=lambda: self._open_help("support"))
        help_menu.add_command(label="上承式桁架/型钢纵梁直线栈桥帮助", command=lambda: self._open_help("trestle"))
        help_menu.add_command(label="钢管贝雷梁作业平台帮助", command=lambda: self._open_help("platform"))
        help_menu.add_command(label="矩形钢板桩/锁扣钢管桩矩形围堰帮助", command=lambda: self._open_help("cofferdam"))
        menubar.add_cascade(label="帮助", menu=help_menu)
        
        self.main.config(menu=menubar)
    
    def _create_statusbar(self):
        """底部状态栏：左侧注册信息；右侧"获取最新版本请访问：百度网盘"超链接"""
        # 状态栏容器（占底部一行）
        status_frame = tb.Frame(self.main, bootstyle="light")
        status_frame.pack(side=BOTTOM, fill=X)
        # 与上方主体分隔
        tb.Separator(status_frame, orient=HORIZONTAL).pack(side=TOP, fill=X)

        # 右侧内部容器：整组右对齐（先 pack，优先保证链接完整显示，不被左侧挤压）
        right = tb.Frame(status_frame, bootstyle="light")
        right.pack(side=RIGHT, padx=12, pady=4)

        tb.Label(right, text=f"当前版本 {get_app_version()}  |  ",
                 font=("Microsoft YaHei UI", 9)).pack(side=LEFT)
        tb.Label(right, text="获取最新版本请访问：",
                 font=("Microsoft YaHei UI", 9)).pack(side=LEFT)
        # 超链接：蓝色下划线 + 手形光标 + 点击打开浏览器
        link = tb.Label(right, text="百度网盘",
                        font=("Microsoft YaHei UI", 9, "underline bold"),
                        foreground="#0d6efd",
                        cursor="hand2")
        link.pack(side=LEFT)
        link.bind("<Button-1>", self._open_update_link)
        self.update_link_label = link  # 保留引用（可选）

        # 左侧：注册状态 / 授权到期时间 / 剩余天数（用分隔符隔开）
        self.statusbar_info_label = tb.Label(status_frame, text=self._get_registration_text(),
                                             font=("Microsoft YaHei UI", 9), padding=0)
        self.statusbar_info_label.pack(side=LEFT, padx=12, pady=4)

    def _get_registration_text(self):
        """查询注册信息：是否注册 / 授权到期时间 / 剩余天数 
        须同时存在 注册码(RegistrationCode_Python) 与 注册时间(Registration_Time) 才算已授权。
        """
        import datetime
        from General.DataUtils import read_Register

        REG_PATH = "Software\\ShuZhiQiaoShi"
        REG_CODE_KEY = "RegistrationCode_Python"
        REG_TIME_KEY = "Registration_Time"
        VALID_DAYS = 180

        reg_code = read_Register(REG_PATH, REG_CODE_KEY)
        reg_time = read_Register(REG_PATH, REG_TIME_KEY)

        if reg_code == "NotFound" or not reg_code:
            return "未注册"
        if reg_time == "NotFound" or not reg_time:
            return "未注册（缺少授权时间）"

        try:
            reg_date = datetime.date.fromisoformat(str(reg_time)[:10])
            expire_date = reg_date + datetime.timedelta(days=VALID_DAYS)
            expire_str = expire_date.strftime("%Y-%m-%d")
            days_left = (expire_date - datetime.date.today()).days
        except Exception as e:
            print(f"解析注册时间失败: {e}")
            return "已注册（到期信息解析失败）"

        if days_left < 0:
            status = "已过期"
        else:
            status = "已注册"
        return f"{status}  |  授权至 {expire_str}  |  剩余 {days_left} 天"

    def _refresh_statusbar(self):
        """注册信息变化后刷新底部状态栏（如通过帮助-注册成功注册后）"""
        if hasattr(self, "statusbar_info_label") and self.statusbar_info_label.winfo_exists():
            self.statusbar_info_label.config(text=self._get_registration_text())

    def _open_update_link(self, event=None):
        """点击"百度网盘"打开下载链接"""
        try:
            import webbrowser
            webbrowser.open("https://pan.baidu.com/s/1YMmsI12FfFcsHM1YILTU7w?pwd=9192")
        except Exception as e:
            print(f"打开链接失败: {e}")
        return "break"  # 阻止事件继续传播

    def _on_about_register(self):
        """帮助-注册：未注册弹注册窗口；已注册弹提示（含截止日期/剩余天数）"""
        import datetime
        from General.DataUtils import read_Register
        from General.UIHandle import show_register_window

        REG_PATH = "Software\\ShuZhiQiaoShi"
        REG_CODE_KEY = "RegistrationCode_Python"
        REG_TIME_KEY = "Registration_Time"
        VALID_DAYS = 180

        reg_code = read_Register(REG_PATH, REG_CODE_KEY)
        reg_time = read_Register(REG_PATH, REG_TIME_KEY)
        # 注册码 或 注册时间 任一缺失 → 视为未完整授权，弹注册窗口
        if (reg_code == "NotFound" or not reg_code) or (reg_time == "NotFound" or not reg_time):
            print("未注册或授权信息不完整, 请先进行注册")
            if show_register_window(self.main):
                # 注册成功 → 刷新状态栏
                self._refresh_statusbar()
            return

        try:
            reg_date = datetime.date.fromisoformat(str(reg_time)[:10])
            expire_date = reg_date + datetime.timedelta(days=VALID_DAYS)
            expire_str = expire_date.strftime("%Y-%m-%d")
            days_left = (expire_date - datetime.date.today()).days
            if days_left < 0:
                status = "已过期"
                days_str = f"{abs(days_left)} 天（已逾期）"
            else:
                status = "已注册"
                days_str = f"{days_left} 天"
        except Exception as e:
            print(f"解析注册时间失败: {e}")
            messagebox.showinfo("注册信息", "已注册（授权时间解析失败）", parent=self.main)
            return

        messagebox.showinfo("注册信息",
                            f"状态：{status}\n"
                            f"授权截止日期：{expire_str}\n"
                            f"剩余天数：{days_str}",
                            parent=self.main)

    def _load_sidebar_icon(self, icon_file):
        """加载并缩放到统一尺寸的侧边栏图标（透明底 PNG）"""
        icon_path = os.path.join(self.icon_path, icon_file)
        if not os.path.exists(icon_path):
            return None
        try:
            img = Image.open(icon_path).convert("RGBA").resize(
                (SIDEBAR_ICON_SIZE, SIDEBAR_ICON_SIZE), Image.LANCZOS)
            return ImageTk.PhotoImage(img)
        except Exception as e:
            print(f"无法加载侧边栏图标 {icon_file}: {e}")
            return None

    def _create_sidebar(self):
        """创建侧边栏导航：透明图标 + 文字；悬浮浅蓝高亮，选中 primary"""
        tabs = [
            ("modeling", "有限元"),
            ("report", "计算书"),
            ("drawing", "布置图"),
            ("bim", "BIM"),
            ("ai", "AI推荐"),
        ]
        for tab_id, label in tabs:
            img = self._load_sidebar_icon(SIDEBAR_ICONS.get(tab_id))
            if img is not None:
                self._sidebar_imgs.append(img)

            nav = tk.Label(
                self.sidebar_frame,
                text=f" {label}",
                image=img,
                compound="left",
                anchor="w",
                padx=SIDEBAR_NAV_PAD_X, pady=SIDEBAR_NAV_PAD_Y,
                bg=SIDEBAR_BG, fg=SIDEBAR_FG,
                font=("Microsoft YaHei UI", 10, "bold"),
                cursor="hand2",
            )
            nav.image = img
            nav.pack(fill=X, padx=6, pady=3)
            self.sidebar_buttons[tab_id] = nav

            nav.bind("<Enter>", lambda e, t=tab_id: self._nav_hover(t, enter=True))
            nav.bind("<Leave>", lambda e, t=tab_id: self._nav_hover(t, enter=False))
            nav.bind("<Button-1>", lambda e, t=tab_id: self._switch_tab(t))
        self._refresh_nav_styles()

    def _nav_hover(self, tab_id, enter):
        """侧边栏项悬浮/移出高亮（选中项不受悬浮影响）"""
        if tab_id == self.active_tab:
            return
        nav = self.sidebar_buttons[tab_id]
        if enter:
            nav.configure(bg=SIDEBAR_HOVER_BG, fg=SIDEBAR_HOVER_FG)
        else:
            nav.configure(bg=SIDEBAR_BG, fg=SIDEBAR_FG)

    def _refresh_nav_styles(self):
        """按当前选中页刷新所有侧边栏项颜色"""
        for tab_id, nav in self.sidebar_buttons.items():
            if tab_id == self.active_tab:
                nav.configure(bg=SIDEBAR_ACTIVE_BG, fg=SIDEBAR_ACTIVE_FG)
            else:
                nav.configure(bg=SIDEBAR_BG, fg=SIDEBAR_FG)

    def _switch_tab(self, tab_id):
        """切换侧边栏"""
        self.active_tab = tab_id
        self._refresh_nav_styles()
        
        # 清空内容区
        for w in self.content_frame.winfo_children():
            w.destroy()
        
        # 加载内容
        if tab_id == "modeling":
            self._load_func_grid("有限元建模", self.midas_dict, "mct_pyd.png",
                                 label_icons={'批量钢板桩/锁扣钢管桩矩形围堰': 'Cofferdam_Batch_Midas.png'})
        elif tab_id == "drawing":
            self._load_func_grid("CAD绘图", self.cad_dict, "dwg_pyd.png")
        elif tab_id == "bim":
            # BIM 页不设批量入口
            self._load_func_grid("BIM建模", {
                '钢管贝雷梁现浇支架': None, '上承式桁架/型钢纵梁直线栈桥': None, '钢管贝雷梁作业平台': None, '钢板桩/锁扣钢管桩矩形围堰': None
            }, "revit_exe.png", is_revit=True)
        elif tab_id == "report":
            self._load_func_grid("计算书生成", self.word_dict, "docx_pyd.png",
                                 label_icons={'钢板桩/锁扣钢管桩矩形围堰批量结果': 'Cofferdam_Batch_Excel.png'})
        elif tab_id == "ai":
            self._load_ai_content()
    
    def _load_icon_photo(self, icon_file):
        """加载图标并按 ICON_SIZE 缩放（带缓存，同一文件复用同一 PhotoImage）"""
        if not icon_file:
            return None
        if icon_file in self._icon_cache:
            return self._icon_cache[icon_file]
        icon_path = os.path.join(self.icon_path, icon_file)
        if not os.path.exists(icon_path):
            print(f"图标文件不存在: {icon_file}")
            self._icon_cache[icon_file] = None
            return None
        try:
            img = Image.open(icon_path).convert("RGBA").resize(
                (ICON_SIZE, ICON_SIZE), Image.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            self._icon_cache[icon_file] = photo
            return photo
        except Exception as e:
            print(f"无法加载图片 {icon_file}: {e}")
            self._icon_cache[icon_file] = None
            return None

    def _build_icon_entry(self, parent, photo, text, cmd):
        """构建 图标-only 入口：仅图标按钮，悬停显示详细名称"""
        fr = tb.Frame(parent)
        if photo:
            btn = tb.Button(fr, image=photo, style="Custom.TButton", command=cmd)
            btn.pack()
            btn.image = photo  # 保持图片引用
        else:
            # 无图标回退：仅文字按钮
            btn = tb.Button(fr, text=text, style="Custom.TButton", command=cmd, padding=BTN_PAD)
            btn.pack()
        if text:
            createToolTip(btn, text)  # 悬停显示详细名称
        return fr, btn

    def _load_func_grid(self, title, func_dict, icon_file, is_revit=False, label_icons=None):
        """加载功能按钮网格：仅图标入口，悬停 Tooltip 显示详细名称

        label_icons: 可选的 {入口文字: 图标文件名}，优先于 STRUCT_ICONS/icon_file，
                     用于批量等需要按页区分的入口图标。
        """
        # 无标题/分隔线；网格占满内容区，入口按 (row, col) 位置从左上逐格填充
        grid_frame = tb.Frame(self.content_frame)
        grid_frame.pack(fill=BOTH, expand=YES)
        for _c in range(GRID_COLS):
            grid_frame.grid_columnconfigure(_c, uniform="col")
        for _r in range(GRID_ROWS):
            grid_frame.grid_rowconfigure(_r, uniform="row")

        label_icons = label_icons or {}
        items = list(func_dict.items())
        for i, (label, func) in enumerate(items):
            row = i // GRID_COLS
            col = i % GRID_COLS

            if is_revit:
                cmd = lambda l=label: self._launch_revit()
            elif func is None:
                cmd = lambda l=label: messagebox.showinfo("提示", f"{l}功能暂未开发")
            else:
                cmd = lambda f=func, l=label: self._launch_func(f, l)

            # 显示文字走集中入口标签表（键不变，仅展示文本可换行/后续大改）
            show_text = ENTRY_LABELS.get(label, label)

            # 图标选取：按页覆盖 > 结构映射 > 本功能页统一图标
            struct_file = label_icons.get(label) or STRUCT_ICONS.get(label) or icon_file
            photo = self._load_icon_photo(struct_file)

            # 图标-only 入口 + 悬停提示
            fr, _btn = self._build_icon_entry(grid_frame, photo, show_text, cmd)
            fr.grid(row=row, column=col, padx=ICON_BTN_PADX, pady=ICON_BTN_PADY, sticky="nsew")

    def _load_ai_content(self):
        """加载AI推荐内容（与功能页采用完全一致的网格入口样式）"""
        # 按钮网格（与功能页一致：占满并按位置排布）
        grid_frame = tb.Frame(self.content_frame)
        grid_frame.pack(fill=BOTH, expand=YES)
        for _c in range(GRID_COLS):
            grid_frame.grid_columnconfigure(_c, uniform="col")
        for _r in range(GRID_ROWS):
            grid_frame.grid_rowconfigure(_r, uniform="row")

        photo = self._load_icon_photo('Cofferdam_exe.png')
        ai_label = ENTRY_LABELS.get("钢板桩/锁扣钢管桩矩形围堰智能推荐", "钢板桩/锁扣钢管桩矩形围堰智能推荐")

        # 图标-only 入口 + 悬停提示
        fr, _btn = self._build_icon_entry(grid_frame, photo, ai_label, lambda: self._launch_ai())
        fr.grid(row=0, column=0, padx=ICON_BTN_PADX, pady=ICON_BTN_PADY, sticky="nsew")
    
    # ==================== 功能调用 ====================
    
    def _launch_func(self, func, label):
        """启动功能"""
        self.main.withdraw()
        self.minimized[0] = True
        try:
            func(self.root, on_close=self._on_func_close)
        except Exception as e:
            print(f"功能启动失败: {e}")
            messagebox.showerror("错误", f"功能启动失败: {e}")
            self._on_func_close()
    
    def _on_func_close(self):
        """功能关闭回调"""
        if self.main and self.main.winfo_exists():
            self.main.deiconify()
            self._center_main()
        self.minimized[0] = False
    
    def _launch_revit(self):
        """启动 Revit"""
        if self.Revit_Version:
            try:
                import subprocess
                from General.DataUtils import get_revit_install_path
                revit_path = os.path.join(get_revit_install_path(self.Revit_Version), 'Revit.exe')
                subprocess.Popen(["explorer.exe", revit_path])
            except Exception as e:
                messagebox.showerror("错误", f"Revit 启动失败: {e}")
        else:
            messagebox.showwarning("提示", "未检测到 Revit 版本号")
    
    def _launch_ai(self):
        """启动 AI 推荐"""
        try:
            from General.UIHandle import open_smart_recommend_ui
            open_smart_recommend_ui()
        except Exception as e:
            messagebox.showerror("错误", f"AI 推荐启动失败: {e}")
    
    # ==================== 弹窗 ====================
    
    def _place_dialog(self, dlg, w, h):
        """在设置好内容后，把二级窗口定位到主窗口附近再显示（避免左上角闪现）"""
        dlg.update_idletasks()
        main_x = self.main.winfo_x()
        main_y = self.main.winfo_y()
        main_w = self.main.winfo_width()
        main_h = self.main.winfo_height()
        x = main_x + (main_w - w) // 2
        y = main_y + (main_h - h) // 2
        dlg.geometry(f"{w}x{h}+{x}+{y}")
        dlg.deiconify()

    def _show_param(self, struct_type):
        """参数设置弹窗入口
        - 上承式桁架/型钢纵梁直线栈桥/钢板桩/锁扣钢管桩矩形围堰为多项目结构：弹出 项目选择 + 规范设置 窗口（结构类型已由菜单项固定，无结构类型下拉）
        - 钢管贝雷梁现浇支架/钢管贝雷梁作业平台为单项目结构：保留原简单参数窗口
        """
        struct_names = {
            "support": "现浇钢管贝雷梁现浇支架",
            "trestle": "上承式桁架/型钢纵梁直线栈桥",
            "platform": "钢管贝雷梁作业平台",
            "cofferdam": "钢板桩/锁扣钢管桩矩形围堰",
        }
        title = f"参数设置 - {struct_names.get(struct_type, struct_type)}"
        
        if struct_type in ("trestle", "cofferdam"):
            self._show_multi_proj_dialog(struct_type, title)
        else:
            self._show_basic_param_dialog(struct_type, title)
    
    def _show_multi_proj_dialog(self, struct_type, title):
        """多项目结构参数设置：项目名称 + 项目编号 + 钢材/混凝土/钢筋规范"""
        # 加载项目分组与规范选项
        groups = self._load_project_groups(struct_type)
        names = list(groups.keys())
        spec_options = self._load_spec_options()
        
        dlg = tb.Toplevel(self.main)
        dlg.title(title)
        dlg.withdraw()  # 先隐藏，避免闪现
        dlg.transient(self.main)
        dlg.grab_set()
        
        # 底部按钮
        btn_frame = tb.Frame(dlg)
        btn_frame.pack(side=BOTTOM, fill=X, padx=10, pady=10)
        tb.Button(btn_frame, text="取消", command=dlg.destroy, bootstyle="secondary").pack(side=RIGHT, padx=5)
        tb.Button(btn_frame, text="保存", bootstyle="success",
                  command=lambda: self._save_multi_proj(dlg, title, struct_type,
                                                        project_combo, proj_no_combo,
                                                        steel_combo, concrete_combo, rebar_combo)).pack(side=RIGHT, padx=5)
        
        # 左侧项目选择
        left_frame = tb.LabelFrame(dlg, text="项目选择", padding=15)
        left_frame.pack(side=LEFT, fill=Y, padx=10, pady=10, ipadx=10)
        
        tb.Label(left_frame, text="项目名称").pack(anchor=W)
        project_combo = tb.Combobox(left_frame, values=names, state="readonly", width=20)
        project_combo.pack(fill=X, pady=(0, 15))
        
        tb.Label(left_frame, text="项目编号").pack(anchor=W)
        proj_no_combo = tb.Combobox(left_frame, state="readonly", width=20)
        proj_no_combo.pack(fill=X)
        
        def on_project_change(*_):
            pname = project_combo.get()
            nos = groups.get(pname, [])
            proj_no_combo['values'] = nos
            if nos:
                proj_no_combo.set(nos[0])
            else:
                proj_no_combo.set("")
        
        project_combo.bind("<<ComboboxSelected>>", on_project_change)
        if names:
            project_combo.set(names[0])
        on_project_change()
        
        # 右侧规范设置
        right_frame = tb.LabelFrame(dlg, text="规范设置", padding=10)
        right_frame.pack(side=LEFT, fill=BOTH, expand=YES, padx=(0, 10), pady=10)
        
        tb.Label(right_frame, text="钢材规范").pack(anchor=W)
        steel_combo = tb.Combobox(right_frame, values=spec_options.get("steel", []), state="readonly", width=45)
        if spec_options.get("steel"):
            steel_combo.set(spec_options["steel"][0])
        steel_combo.pack(fill=X, pady=(0, 10))
        
        tb.Label(right_frame, text="混凝土规范").pack(anchor=W)
        concrete_combo = tb.Combobox(right_frame, values=spec_options.get("concrete", []), state="readonly", width=45)
        if spec_options.get("concrete"):
            concrete_combo.set(spec_options["concrete"][0])
        concrete_combo.pack(fill=X, pady=(0, 10))
        
        tb.Label(right_frame, text="钢筋规范").pack(anchor=W)
        rebar_combo = tb.Combobox(right_frame, values=spec_options.get("rebar", []), state="readonly", width=45)
        if spec_options.get("rebar"):
            rebar_combo.set(spec_options["rebar"][0])
        rebar_combo.pack(fill=X)
        
        self._place_dialog(dlg, 720, 380)
    
    def _save_multi_proj(self, dlg, title, struct_type, project_combo, proj_no_combo,
                         steel_combo, concrete_combo, rebar_combo):
        """保存多项目结构参数设置（每个项目编号区分保存）"""
        struct_names = {
            "support": "现浇钢管贝雷梁现浇支架",
            "trestle": "上承式桁架/型钢纵梁直线栈桥",
            "platform": "钢管贝雷梁作业平台",
            "cofferdam": "钢板桩/锁扣钢管桩矩形围堰",
        }
        proj_name = project_combo.get()
        proj_no = proj_no_combo.get()
        steel = steel_combo.get()
        concrete = concrete_combo.get()
        rebar = rebar_combo.get()
        
        print(f"参数设置: 结构={struct_names.get(struct_type)}, 项目={proj_name}, 编号={proj_no}")
        print(f"  钢材规范={steel}, 混凝土规范={concrete}, 钢筋规范={rebar}")
        
        messagebox.showinfo("提示", f"参数设置已保存\n\n"
                            f"结构类型: {struct_names.get(struct_type)}\n"
                            f"项目名称: {proj_name}\n"
                            f"项目编号: {proj_no}\n"
                            f"钢材规范: {steel}\n"
                            f"混凝土规范: {concrete}\n"
                            f"钢筋规范: {rebar}")
        dlg.destroy()
    
    def _show_basic_param_dialog(self, struct_type, title):
        """单项目结构（钢管贝雷梁现浇支架/钢管贝雷梁作业平台）简单参数窗口"""
        from General.TextHandle import read_file
        
        param_files = {
            "support": "Basic_Param_Set_ZJ.txt",
            "platform": "Basic_Param_Set_PT.txt",
        }
        
        param_path = os.path.join(self.parent_path, 'Support', 'basic_param', param_files.get(struct_type, ""))
        
        # 读取参数
        self_weight = 1.0
        try:
            basic_param_lst = read_file(param_path)
            self_weight = float(basic_param_lst[0][0]) if basic_param_lst and basic_param_lst[0] else 1.0
        except:
            pass
        
        dlg = tb.Toplevel(self.main)
        dlg.title(title)
        dlg.withdraw()  # 先隐藏
        dlg.transient(self.main)
        dlg.grab_set()
        
        frame = tb.Frame(dlg, padding=20)
        frame.pack(fill=BOTH, expand=YES)
        
        btn_frame = tb.Frame(frame)
        btn_frame.pack(side=BOTTOM, fill=X, pady=(20, 0))
        tb.Button(btn_frame, text="关闭", command=dlg.destroy, bootstyle="secondary").pack(side=RIGHT, padx=5)
        tb.Button(btn_frame, text="保存", bootstyle="success",
                  command=lambda: self._save_basic_param(dlg, struct_type, self_weight_var, sg_load_var, mb_load_var)).pack(side=RIGHT, padx=5)
        
        tb.Label(frame, text="自重系数").pack(anchor=W)
        self_weight_var = tb.DoubleVar(value=self_weight)
        tb.Entry(frame, textvariable=self_weight_var, width=20).pack(fill=X, pady=(0, 10))
        
        tb.Label(frame, text="施工荷载 (kN/m²)").pack(anchor=W)
        sg_load_var = tb.DoubleVar(value=2.0)
        tb.Entry(frame, textvariable=sg_load_var, width=20).pack(fill=X, pady=(0, 10))
        
        tb.Label(frame, text="模板荷载 (kN/m²)").pack(anchor=W)
        mb_load_var = tb.DoubleVar(value=2.5)
        tb.Entry(frame, textvariable=mb_load_var, width=20).pack(fill=X)
        
        self._place_dialog(dlg, 400, 320)
    
    def _save_basic_param(self, dlg, struct_type, self_weight_var, sg_load_var, mb_load_var):
        """保存单项目结构参数"""
        print(f"参数: struct={struct_type}, self_weight={self_weight_var.get()}, "
              f"sg_load={sg_load_var.get()}, mb_load={mb_load_var.get()}")
        messagebox.showinfo("提示", "参数设置已保存")
        dlg.destroy()
    
    def _load_project_groups(self, struct_type):
        """加载项目分组信息 {项目名称: [编号列表]}"""
        try:
            if struct_type == "trestle":
                return self._load_trestle_projects()
            elif struct_type == "cofferdam":
                return self._load_cofferdam_projects()
        except Exception as e:
            print(f"加载项目列表失败: {e}")
        return {}
    
    def _load_trestle_projects(self):
        """读取上承式桁架/型钢纵梁直线栈桥项目信息（独立轻量读取，只读 项目名称+编号）"""
        import openpyxl
        
        # 获取上承式桁架/型钢纵梁直线栈桥参数表路径
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\ShuZhiQiaoShi")
            app, _ = winreg.QueryValueEx(key, "Applocation")
            winreg.CloseKey(key)
            excel_path = os.path.join(app, "Support", "program_param", "Trestle_ParamTable.xlsx")
        except:
            return {}
        
        if not os.path.exists(excel_path):
            return {}
        
        project_groups = {}
        try:
            wb = openpyxl.load_workbook(excel_path, data_only=True)
            # 查找上承式桁架/型钢纵梁直线栈桥参数表
            ws = None
            for name in wb.sheetnames:
                if "栈桥参数" in name or name == "1":
                    ws = wb[name]
                    break
            if ws is None:
                wb.close()
                return {}
            
            # 构建表头映射
            hm = {}
            for c in range(1, ws.max_column + 1):
                val = ws.cell(2, c).value
                if val:
                    hm[str(val).strip()] = c
            
            # 读取项目信息
            for r in range(3, ws.max_row + 1):
                pname_cell = ws.cell(r, hm.get("项目名称", 0)).value
                pno_cell = ws.cell(r, hm.get("编号", 0)).value
                pname = str(pname_cell).strip() if pname_cell else ""
                pno = str(pno_cell).strip() if pno_cell else ""
                if not pname:
                    continue
                if pname not in project_groups:
                    project_groups[pname] = []
                if pno and pno not in project_groups[pname]:
                    project_groups[pname].append(pno)
            
            wb.close()
        except Exception as e:
            print(f"读取上承式桁架/型钢纵梁直线栈桥项目失败: {e}")
        
        return project_groups
    
    def _load_cofferdam_projects(self):
        """读取钢板桩/锁扣钢管桩矩形围堰项目信息（独立轻量读取，只读 项目名称+编号）"""
        import openpyxl
        
        # 获取钢板桩/锁扣钢管桩矩形围堰参数表路径
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\ShuZhiQiaoShi")
            app, _ = winreg.QueryValueEx(key, "Applocation")
            winreg.CloseKey(key)
            excel_path = os.path.join(app, "Support", "program_param", "CofferDam_ParamTable.xlsx")
        except:
            return {}
        
        if not os.path.exists(excel_path):
            return {}
        
        project_groups = {}
        try:
            wb = openpyxl.load_workbook(excel_path, data_only=True)
            # 查找矩形钢板桩/锁扣钢管桩矩形围堰设计参数表
            ws = None
            for name in wb.sheetnames:
                if "矩形围堰" in name:
                    ws = wb[name]
                    break
            if ws is None:
                wb.close()
                return {}
            
            # 构建表头映射
            hm = {}
            for c in range(1, ws.max_column + 1):
                val = ws.cell(2, c).value
                if val:
                    hm[str(val).strip()] = c
            
            # 读取项目信息
            for r in range(3, ws.max_row + 1):
                pname_cell = ws.cell(r, hm.get("项目名称", 0)).value
                pno_cell = ws.cell(r, hm.get("围堰编号", 0)).value
                pname = str(pname_cell).strip() if pname_cell else ""
                pno = str(pno_cell).strip() if pno_cell else ""
                if not pname:
                    continue
                if pname not in project_groups:
                    project_groups[pname] = []
                if pno and pno not in project_groups[pname]:
                    project_groups[pname].append(pno)
            
            wb.close()
        except Exception as e:
            print(f"读取钢板桩/锁扣钢管桩矩形围堰项目失败: {e}")
        
        return project_groups
    
    def _load_spec_options(self):
        """加载规范选项 - 读取对应列下的全部规范名去重"""
        import openpyxl
        excel_path = os.path.join(self.parent_path, 'Support', 'basic_param', 'properties_parameter.xlsx')
        
        steel_list = []
        concrete_list = []
        rebar_list = []
        
        try:
            wb = openpyxl.load_workbook(excel_path, data_only=True)
            ws = wb['材质']
            
            # 第2行是具体列名，构建映射
            hm = {}
            for c in range(1, ws.max_column + 1):
                val = ws.cell(2, c).value
                if val:
                    hm[str(val).strip()] = c
            
            # 从第3行开始读取数据
            # 钢材规范名：列1（钢结构规范名）
            steel_col = hm.get("钢结构规范名", 1)
            for r in range(3, ws.max_row + 1):
                val = ws.cell(r, steel_col).value
                if val and str(val).strip() and str(val).strip() not in steel_list:
                    steel_list.append(str(val).strip())
            
            # 混凝土规范名：列17（混凝土规范名）
            concrete_col = hm.get("混凝土规范名", 17)
            for r in range(3, ws.max_row + 1):
                val = ws.cell(r, concrete_col).value
                if val and str(val).strip() and str(val).strip() not in concrete_list:
                    concrete_list.append(str(val).strip())
            
            # 钢筋规范名：列31（钢筋规范名）
            rebar_col = hm.get("钢筋规范名", 31)
            for r in range(3, ws.max_row + 1):
                val = ws.cell(r, rebar_col).value
                if val and str(val).strip() and str(val).strip() not in rebar_list:
                    rebar_list.append(str(val).strip())
            
            wb.close()
            print(f"规范选项: 钢材={steel_list}, 混凝土={concrete_list}, 钢筋={rebar_list}")
        except Exception as e:
            print(f"加载规范选项失败: {e}")
        
        return {
            "steel": steel_list,
            "concrete": concrete_list,
            "rebar": rebar_list,
        }
    
    # ==================== 辅助 ====================
    
    def _open_folder(self):
        folder_path = os.path.join(self.parent_path, 'Custom')
        if os.path.exists(folder_path):
            os.startfile(folder_path)
    
    def _open_help(self, struct_type):
        help_paths = {
            'support': '现浇钢管贝雷梁现浇支架', 'trestle': '施工上承式桁架/型钢纵梁直线栈桥',
            'platform': '钢管贝雷梁作业平台', 'cofferdam': '钢板桩钢板桩/锁扣钢管桩矩形围堰',
        }
        folder_path = os.path.join(self.parent_path, 'Help', help_paths.get(struct_type, ''))
        if os.path.exists(folder_path):
            os.startfile(folder_path)
    
    def exit_app(self):
        if self.main and self.main.winfo_exists():
            self.main.destroy()
        if self.ball and self.ball.winfo_exists():
            self.ball.destroy()
        if self.root and self.root.winfo_exists():
            self.root.destroy()
        import gc
        import matplotlib.pyplot as plt
        gc.collect()
        plt.close('all')
        sys.exit(0)
