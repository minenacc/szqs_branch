# 1. 标准库
import os
import winreg
import tkinter as tk
from PIL import Image, ImageTk 

# 2. 第三方库
import ttkbootstrap as tb
from ttkbootstrap.constants import *

# 3. 本地模块
from General.DataUtils import read_Register


# 选择保存文件的路径
def dialog_savepath(parent, label_txt, entry_var, button_txt, button_command):
    container = tb.Frame(parent)
    container.pack(fill="x", pady=5)

    # 标签
    tb.Label(container, text=label_txt).pack(side="left", padx=5)

    # 显示路径的 Entry（不可编辑）
    ent = tb.Entry(container, textvariable=entry_var, state="disabled")
    ent.pack(side="left", padx=10, fill="x", expand=True)

    # 按钮
    tb.Button(container, text=button_txt, command=button_command).pack(side="right", padx=10)

    return container

# ============================================构件设置 Tab 页========================================
class ComponentsTab(tb.Frame):
    def __init__(self, parent):
        super().__init__(parent)

        # ============ 滚动设置 ============
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.scrollbar = tb.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")

        # 内层 Frame
        self.inner = tb.Frame(self.canvas)
        self.inner_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        # 确保滚动区域随内容更新
        def on_frame_configure(event):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            # 同时保证内部 frame 高度至少和 canvas 高度一样
            if self.inner.winfo_reqheight() < self.canvas.winfo_height():
                self.canvas.itemconfig(self.inner_id, height=self.canvas.winfo_height())
        self.inner.bind("<Configure>", on_frame_configure)

        # 保证内部 Frame 宽度跟 Canvas 一样
        def on_canvas_resize(event):
            self.canvas.itemconfig(self.inner_id, width=event.width)
        self.canvas.bind("<Configure>", on_canvas_resize)

        # 鼠标滚轮控件
        def _on_mousewheel(event):
            widget = event.widget
            while widget is not None:
                if isinstance(widget, tk.Canvas):
                    widget.yview_scroll(int(-1 * (event.delta / 120)), "units")
                    return
                widget = widget.master
        self.winfo_toplevel().bind_all("<MouseWheel>", _on_mousewheel)
        self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", _on_mousewheel))
        self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))

        # 统一样式参数
        self.padx = 8
        self.pady = 5
        self.label_width = 22
        self.entry_width = 40

        self._tips()

        # ======================== 桥面系定义 ========================
        self._create_deck_system()

        # ======================== 贝雷梁定义 ========================
        self._create_bailey_section()

        # ======================== 分配梁定义 ========================
        self._create_distribute_section()

        # ======================== 下部结构定义 ========================
        self._create_under_section()

    # ======================== 工具函数 ========================
    def _add_placeholder(self, entry, text):
        """给 Entry 添加占位符效果"""
        entry.insert(0, text)
        entry.config(foreground="gray")

        def on_focus_in(event):
            if entry.get() == text:
                entry.delete(0, tk.END)
                entry.config(foreground="black")

        def on_focus_out(event):
            if not entry.get():
                entry.insert(0, text)
                entry.config(foreground="gray")

        entry.bind("<FocusIn>", on_focus_in)
        entry.bind("<FocusOut>", on_focus_out)

    def toggle_widgets(self, var, widgets):
        """根据复选框变量控制控件显示/隐藏"""
        if var.get():
            for w in widgets:
                w.grid_remove()
        else:
            for w in widgets:
                w.grid()

        # 刷新滚动区域
        self.update_idletasks()
        try:
            # 使用保存的 canvas（更可靠）
            if hasattr(self, "canvas"):
                self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        except Exception:
            pass
                
    # ======================== 各区块定义 ========================
    def _tips(self):
        lf = tb.LabelFrame(self.inner, text="说明", padding=10)
        lf.pack(fill="x", padx=10, pady=5)

        tip_text = (
            "以下构件布置输入均以栈桥平面图中左下角端作为起始端，在midas建模中以小肋起始端为(0,0,0)点，"
            "向x轴(纵向)、y轴(横向)正方向及z轴负方向构建模型。"
        )

        label = tb.Label(
            lf,
            text=tip_text,
            foreground="gray",
            wraplength=1,
            justify="left"
        )
        label.grid(row=0, column=0, sticky="w", padx=self.padx, pady=self.pady)

        # 绑定宽度变化事件，动态更新 wraplength
        def _update_wrap(event):
            # event.width 是 LabelFrame 的内容区宽度
            label.configure(wraplength=event.width - 20)

        lf.bind("<Configure>", _update_wrap)

    def _create_deck_system(self):
        lf = tb.LabelFrame(self.inner, text="桥面系定义", padding=10)
        lf.pack(fill="x", padx=10, pady=5)

        # ========== 桥面板 ==========
        lf_deck= tb.LabelFrame(lf, text="桥面板定义", padding=10)
        lf_deck.pack(fill="x", padx=10, pady=self.pady)
        tb.Label(lf_deck, text="桥面板厚(mm):", width = self.label_width).grid(row=0, column=0, sticky="w", padx=self.padx, pady=self.pady)
        self.entry_deck = tb.Entry(lf_deck, width=self.entry_width)
        self.entry_deck.insert(0, "10")
        self.entry_deck.grid(row=0, column=1, padx=self.padx, pady=self.pady, sticky="w")
        tb.Label(lf_deck, text="桥面设计高程(m):", width = self.label_width).grid(row=1, column=0, sticky="w", padx=self.padx, pady=self.pady)
        self.entry_deck_height = tb.Entry(lf_deck, width=self.entry_width)
        self.entry_deck_height.insert(0, "+19.0")
        self.entry_deck_height.grid(row=1, column=1, padx=self.padx, pady=self.pady, sticky="w")

        # ========== 小肋 ==========
        lf_smallrib= tb.LabelFrame(lf, text="小肋定义", padding=10)
        lf_smallrib.pack(fill="x", padx=10, pady=self.pady)

        tb.Label(lf_smallrib, text="小肋纵向布置(mm):", width = self.label_width).grid(row=0, column=0, sticky="w", padx=self.padx, pady=self.pady)
        self.entry_smallrib = tb.Entry(lf_smallrib, width=self.entry_width)
        self.entry_smallrib.insert(0, "79@300")
        self.entry_smallrib.grid(row=0, column=1, padx=self.padx, pady=self.pady, sticky="w")
        
        tb.Label(lf_smallrib, text="纵向起始端间距(mm):", width = self.label_width).grid(row=1, column=0, sticky="w", padx=self.padx, pady=self.pady)
        self.entry_smallrib_init_spacing = tb.Entry(lf_smallrib, width=self.entry_width,foreground="gray")
        self.entry_smallrib_init_spacing.grid(row=1, column=1, padx=self.padx, pady=self.pady, sticky="w")
        self._add_placeholder(self.entry_smallrib_init_spacing, "请输入第一根小肋相对贝雷梁起始端的纵向起始间距")
        
        tb.Label(lf_smallrib, text="小肋长度(mm):", width = self.label_width).grid(row=2, column=0, sticky="w", padx=self.padx, pady=self.pady)
        self.entry_smallrib_length = tb.Entry(lf_smallrib, width=self.entry_width)
        self.entry_smallrib_length.insert(0, "19000")
        self.entry_smallrib_length.grid(row=2, column=1, padx=self.padx, pady=self.pady, sticky="w")

    def _create_bailey_section(self):
        lf = tb.LabelFrame(self.inner, text="贝雷梁定义", padding=10)
        lf.pack(fill="x", padx=10, pady=self.pady)

        tb.Label(lf, text="贝雷梁横向布置(mm):", width = self.label_width).grid(row=0, column=0, sticky="w", padx=self.padx, pady=self.pady)
        self.entry_bailey_spacing = tk.Text(
            lf,
            width=40,         # 输入框宽度
            height=3,         # 默认显示 3 行
            wrap="word"       # 自动换行
        )
        self.entry_bailey_spacing.insert("1.0", "900,450,4@225,450,3200,450,5@225,450,2850,450,5@225,450,3200,450,4@225,450")
        self.entry_bailey_spacing.grid(
            row=0, column=1,
            padx=self.padx, pady=self.pady,
            sticky="w"
        )

        tb.Label(lf, text="贝雷梁纵向布置(m):", width = self.label_width).grid(row=1, column=0, sticky="w", padx=self.padx, pady=self.pady)
        self.entry_bailey_span = tb.Entry(lf, width=self.entry_width)
        self.entry_bailey_span.insert(0, "2@12")
        self.entry_bailey_span.grid(row=1, column=1, padx=self.padx, pady=self.pady, sticky="w")

        tb.Label(lf, text="横向起始端间距(mm):", width = self.label_width).grid(row=2, column=0, sticky="w", padx=self.padx, pady=self.pady)
        self.entry_bailey_init_spacing = tb.Entry(lf, width=self.entry_width,foreground="gray")
        self.entry_bailey_init_spacing.grid(row=2, column=1, padx=self.padx, pady=self.pady, sticky="w")
        self._add_placeholder(self.entry_bailey_init_spacing, "请输入第一片贝雷梁相对小肋起始端的横向起始间距")

    def _create_distribute_section(self):
        lf = tb.LabelFrame(self.inner, text="分配梁定义", padding=10)
        lf.pack(fill="x", padx=10, pady=self.pady)

        tb.Label(lf, text="分配梁长度(mm):", width = self.label_width).grid(row=0, column=0, sticky="w", padx=self.padx, pady=self.pady)
        self.entry_distribute_length = tb.Entry(lf, width=self.entry_width)
        self.entry_distribute_length.insert(0, "19000")
        self.entry_distribute_length.grid(row=0, column=1, padx=self.padx, pady=self.pady, sticky="w")

        tb.Label(lf, text="起始端悬臂长度-贝雷梁(mm):", width = self.label_width).grid(row=1, column=0, sticky="w", padx=self.padx, pady=self.pady)
        self.entry_distribute_cantilever = tb.Entry(lf, width=self.entry_width, foreground="gray")
        self.entry_distribute_cantilever.grid(row=1, column=1, padx=self.padx, pady=self.pady, sticky="w")
        self._add_placeholder(self.entry_distribute_cantilever, "请输入分配梁起始端相对第一片贝雷梁的悬臂长度")

        tb.Label(lf, text="起始端悬臂长度-钢管桩(mm):", width = self.label_width).grid(row=2, column=0, sticky="w", padx=self.padx, pady=self.pady)
        self.entry_distribute_cantilever2 = tb.Entry(lf, width=self.entry_width, foreground="gray")
        self.entry_distribute_cantilever2.grid(row=2, column=1, padx=self.padx, pady=self.pady, sticky="w")
        self._add_placeholder(self.entry_distribute_cantilever2, "请输入分配梁起始端相对钢管桩的悬臂长度")


    def _create_under_section(self):
        lf = tb.LabelFrame(self.inner, text="下部结构定义", padding=10)
        lf.pack(fill="x", padx=10, pady=5)
        # lf.pack_propagate(False)

        # ========== 钢管桩 ==========
        lf_pile = tb.LabelFrame(lf, text="钢管桩设置", padding=10)
        lf_pile.pack(fill="x", padx=5, pady=5)
        # lf_pile.pack_propagate(False)
        # lf_pile.config(width=680, height=60)

        # 横向间距
        tb.Label(lf_pile, text="钢管桩横向间距(mm):",width = self.label_width).grid(row=0, column=0, sticky="w", padx=self.padx, pady=self.pady)
        self.entry_pile_y = tb.Entry(lf_pile, width=40)
        self.entry_pile_y.insert(0, "5500+5000+5500")
        self.entry_pile_y.grid(row=0, column=1, padx=self.padx, pady=self.pady, sticky="w")

        # 纵向间距
        tb.Label(lf_pile, text="钢管桩纵向间距(mm):",width = self.label_width).grid(row=1, column=0, sticky="w", padx=self.padx, pady=self.pady)
        self.entry_pile_x = tb.Entry(lf_pile, width=40)
        self.entry_pile_x.insert(0, "12000+9000")
        self.entry_pile_x.grid(row=1, column=1, padx=self.padx, pady=self.pady, sticky="w")

        # 悬臂长度
        tb.Label(lf_pile, text="悬臂长度(mm):",width = self.label_width).grid(row=2, column=0, sticky="w", padx=self.padx, pady=self.pady)
        self.entry_pile_cantilever = tb.Entry(lf_pile, width=40)
        self.entry_pile_cantilever.grid(row=2, column=1, padx=self.padx, pady=self.pady, sticky="w")
        self._add_placeholder(self.entry_pile_cantilever, "请输入贝雷梁起始端相对钢管桩的悬臂长度")

        # 桩长
        tb.Label(lf_pile, text="桩长(m):",width = self.label_width).grid(row=3, column=0, sticky="w", padx=self.padx, pady=self.pady)
        self.entry_pile_length = tb.Entry(lf_pile, width=40)
        self.entry_pile_length.insert(0, "31.76")
        self.entry_pile_length.grid(row=3, column=1, padx=self.padx, pady=self.pady, sticky="w")

        # ========== 联结系 ==========
        lf_connect = tb.LabelFrame(lf, text="联结系设置", padding=10)
        lf_connect.pack(fill="x", padx=5, pady=5)
        # lf_connect.pack_propagate(False)
        # lf_connect.config(width=680, height=60)

        self.label_connect_zd = tb.Label(lf_connect, text="联结系距桩顶间距(mm):")
        self.label_connect_zd.grid(row=0, column=1, sticky="w", padx=5, pady=2)
        self.entry_connect_zd = tb.Entry(lf_connect, width=10)
        self.entry_connect_zd.insert(0,"616")
        self.entry_connect_zd.grid(row=0, column=2, padx=5)

        self.label_connect_z = tb.Label(lf_connect, text="联结系竖向间距(mm):")
        self.label_connect_z.grid(row=0, column=3, sticky="w", padx=5, pady=2)
        self.entry_connect_z = tb.Entry(lf_connect, width=10)
        self.entry_connect_z.insert(0,"2500")
        self.entry_connect_z.grid(row=0, column=4, padx=5)

        

# ============================================截面设置 Tab 页========================================
class SectionTab(tb.Frame):
    def __init__(self, parent):
        super().__init__(parent)
    
        # 行列配置
        for i in range(10):
            self.grid_rowconfigure(i, weight=0)
        self.grid_columnconfigure(0, weight=1)

        # 读取软件安装路径（注册表）
        self.install_path = read_Register("Software\\ShuZhiQiaoShi", "Applocation")

        # 防止 None 错误：如果注册表不存在，则使用当前目录
        if not self.install_path:
            self.install_path = os.getcwd()

        # 设置 PNG 图片文件夹路径
        self.png_path = os.path.join(self.install_path, "Support", "png_picture")

        self.brake_enabled = True

        # 选择构件
        component_frame = tb.Frame(self)
        component_frame.grid(row=0, column=0, sticky="w", padx=15, pady=(10, 5))
        component_frame.columnconfigure((0,1,2,3,4,5), weight=1)

        tb.Label(component_frame, text="选择构件：").grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.component_list = ["小肋", "分配梁", "钢管桩", "联结系","钢护筒"]
        self.component_params = {}     # 存储每个构件的参数
        self.current_component = None  # 当前选中构件

        self.component_buttons = {}
        self.param_entries = {} 
        for i, name in enumerate(self.component_list):
            btn = tk.Button(
            component_frame,
            text=name,
            command=lambda n=name: self.select_component(n),
            activeforeground="white",
            relief="flat",
            bd=0
        )
            btn.grid(row=1, column=i, padx=5, pady=5, sticky="ew")
            self.component_buttons[name] = btn

        # 构件参数存放参数：self.component_params
        # 选择截面类型
        section_select_frame = tb.Frame(self)
        section_select_frame.grid(row=1, column=0, sticky="w", padx=15, pady=(0, 10))
        tb.Label(section_select_frame, text="选择截面类型：").grid(row=0, column=0, sticky="w", padx=(0, 5))

        self.section_var = tk.StringVar()
        self.section_params = {}
        self.combo = tb.Combobox(
            section_select_frame,
            values=["工字钢截面", "槽钢截面",  "HM截面","HN截面","管型截面", "双拼工字钢截面", "双拼槽钢截面", "双拼HM截面","双拼HN截面"],
            textvariable=self.section_var,
            width=15,
            state="readonly"
        )
        self.combo.grid(row=0, column=1, sticky="w")
        self.combo.bind("<<ComboboxSelected>>", self.on_section_change)
        
        # 主体部分（参数 + 示意图）
        main_frame = tb.Frame(self)
        main_frame.grid(row=2, column=0, sticky="nsew", padx=20, pady=10)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)

        # 左侧：参数设置
        self.param_frame = tb.LabelFrame(main_frame, text="参数设置", padding=10)
        self.param_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 20))

        # 右侧：截面示意图
        self.preview_frame = tb.LabelFrame(main_frame, text="截面示意图", padding=10)
        self.preview_frame.grid(row=0, column=1, sticky="nsew")

        self.preview_label = tb.Label(
            self.preview_frame,
            text="（此处预留示意图）",
            anchor="center"
        )
        self.preview_label.grid(row=0, column=0, sticky="nsew", padx=10, pady=30)
        self.preview_frame.grid_rowconfigure(0, weight=1)
        self.preview_frame.grid_columnconfigure(0, weight=1)

        # 参数定义
        self.section_params = {
            "工字钢截面": ["H(mm)", "B(mm)", "tw(mm)", "tf(mm)", "r(mm)"],
            "槽钢截面": ["H(mm)", "B(mm)", "tw(mm)", "tf(mm)", "r1(mm)", "r2(mm)"],
            "HM截面": ["H(mm)", "B(mm)", "tw(mm)", "tf(mm)", "r(mm)"],
            "HN截面": ["H(mm)", "B(mm)", "tw(mm)", "tf(mm)", "r(mm)"],
            "管型截面": ["D(mm)", "d(mm)"],
            "双拼工字钢截面": ["H(mm)", "B(mm)", "tw(mm)", "tf(mm)", "C(mm)"],
            "双拼槽钢截面": ["H(mm)", "B(mm)", "tw(mm)", "tf(mm)", "C(mm)"],
            "双拼HM截面": ["H(mm)", "B(mm)", "tw(mm)", "tf(mm)", "C(mm)"],
            "双拼HN截面": ["H(mm)", "B(mm)", "tw(mm)", "tf(mm)", "C(mm)"]
        }

        # === 截面示意图映射表 ===
        self.section_image_map = {
            "工字钢截面":      os.path.join(self.png_path, "H_beam.png"),
            "双拼工字钢截面":  os.path.join(self.png_path, "box_section.png"),

            "槽钢截面":        os.path.join(self.png_path, "C_beam.png"),
            "双拼槽钢截面":    os.path.join(self.png_path, "double_C_beam.png"),

            "HM截面":         os.path.join(self.png_path, "H_beam.png"),
            "双拼HM截面":     os.path.join(self.png_path, "box_section.png"),

            "HN截面":         os.path.join(self.png_path, "H_beam.png"),
            "双拼HN截面":     os.path.join(self.png_path, "box_section.png"),

            "管型截面":        os.path.join(self.png_path, "chs.png"),
        }


        # 初始化默认显示
        self.section_var.set("工字钢截面")
        self.create_param_entries("工字钢截面")
        self.select_component(self.component_list[0])
        
        self.section_settings = {
        comp: {
            "section_type": data.get("section_type", ""),
            "params": data.get("params", {})
        }
        for comp, data in self.component_params.items()
        }
        
        self.on_section_change(self.section_var.get())

    # ============ 工具函数 ============
    # 调用对应输入框
    def create_param_entries(self, section_type):
        # 清空旧输入框
        for widget in self.param_frame.winfo_children():
            widget.destroy()
        params = self.section_params.get(section_type, [])
        self.param_entries = {}

        # 获取当前保存值
        saved_params = {}
        if self.current_component:
            comp_data = self.component_params.get(self.current_component)
            if comp_data and comp_data.get("section_type") == section_type:
                saved_params = comp_data.get("params", {})

        # 创建参数输入框
        for i, name in enumerate(params):
            tb.Label(self.param_frame, text=f"{name}", width=15).grid(row=i, column=0, sticky="e", pady=3, padx=(0, 5))
            entry = tb.Entry(self.param_frame, width=10)
            entry.grid(row=i, column=1, sticky="w", pady=3)
            if name in saved_params and saved_params[name] != "":
                entry.insert(0, saved_params[name])
            elif "角度" in name:
                entry.insert(0, "0")
            self.param_entries[name] = entry

        # 若非管型截面，额外添加“每延米重”
        if section_type != "管型截面" :
            row_idx = len(params)
            tb.Label(self.param_frame, text="每延米重(kg/m)：", width=15).grid(row=row_idx, column=0, sticky="e", pady=3)
            weight_entry = tb.Entry(self.param_frame, width=10)
            weight_entry.grid(row=row_idx, column=1, sticky="w", pady=3)
            if "每延米重 (kg/m)" in saved_params:
                weight_entry.insert(0, saved_params["每延米重 (kg/m)"])
            self.param_entries["每延米重 (kg/m)"] = weight_entry

        # 回车跳转
        entries_list = list(self.param_entries.values())
        for i, entry in enumerate(entries_list):
            if i < len(entries_list) - 1:
                entry.bind("<Return>", lambda e, nxt=entries_list[i+1]: nxt.focus_set())
            else:
                entry.bind("<Return>", lambda e: entries_list[0].focus_set())

    def select_component(self, name):
        # 保存上一个构件参数
        if self.current_component is not None:
            self.save_current_params()

        self.current_component = name

        # 高亮当前按钮
        for comp, btn in self.component_buttons.items():
            btn.configure(background="#28a745" if comp == name else "gray")

        # 钢管桩类：仅允许“管型截面”，禁用下拉框
        if name in ["钢管桩", "钢护筒"]:
            self.section_var.set("管型截面")
            self.combo.set("管型截面")
            self.combo.configure(state="disabled")
        else:
            self.combo.configure(state="readonly")

        # 加载截面类型
        comp_data = self.component_params.get(name)
        if comp_data:
            section_type = comp_data.get("section_type", "工字钢截面")
        else:
            section_type = self.section_var.get() if self.combo["state"] != "disabled" else "管型截面"

        self.section_var.set(section_type)
        self.create_param_entries(section_type)
        self.on_section_change()

    def save_current_params(self):
        # 保存当前构件的参数与截面类型
        if not self.current_component:
            return
        section_type = self.section_var.get()
        current_values = {}
        for key, entry in self.param_entries.items():
            val = entry.get().strip()
            if val != "":
                current_values[key] = val

        # 存储
        self.component_params[self.current_component] = {
            "section_type": section_type,
            "params": current_values
        }

    def on_section_change(self, event=None):
        section_type = self.section_var.get()
        if self.current_component:
            self.save_current_params()
        self.create_param_entries(section_type)

    def on_section_change(self, event=None):
        section_type = self.section_var.get()
        # 切换示意图
        img_path = self.section_image_map.get(section_type)
        if img_path:
            try:
                self.update_preview_image(img_path)
            except Exception as e:
                print("图片载入失败：", e)

        # 切换参数
        if self.current_component:
            self.save_current_params()
        self.create_param_entries(section_type)

    def get_all_section_data(self):
        """返回所有构件的截面信息，供生成 MCT 使用"""
        # 先保存当前正在编辑的构件，防止切换时丢失输入
        self.save_current_params()

        all_sections = {}
        for comp, data in self.component_params.items():
            section_type = data.get("section_type", "")
            params = data.get("params", {})
            all_sections[comp] = {
                "section_type": section_type,
                "params": params
            }

        return all_sections
    
    def update_preview_image(self, img_path):
        """在右侧预览框加载图片"""
        # 1. 读取图片
        img = Image.open(img_path)

        # 2. 自适应缩放（可选）
        img = img.resize((350, 300), Image.LANCZOS)

        # 3. 转换成可用于 Tkinter 的图像
        self.preview_imgtk = ImageTk.PhotoImage(img)

        # 4. 放到 Label
        self.preview_label.config(image=self.preview_imgtk, text="")  # 清除文字               
# ============================================荷载参数 Tab 页========================================     
class LoadTab(tb.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        # ============ 滚动设置 ============
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.scrollbar = tb.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")

        # 内层 Frame
        self.inner = tb.Frame(self.canvas)
        self.inner_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        # 确保滚动区域随内容更新
        def on_frame_configure(event):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            # 同时保证内部 frame 高度至少和 canvas 高度一样
            # if self.inner.winfo_reqheight() < self.canvas.winfo_height():
            #     self.canvas.itemconfig(self.inner_id, height=self.canvas.winfo_height())
        self.inner.bind("<Configure>", on_frame_configure)

        # 保证内部 Frame 宽度跟 Canvas 一样
        def on_canvas_resize(event):
            self.canvas.itemconfig(self.inner_id, width=event.width)
        self.canvas.bind("<Configure>", on_canvas_resize)

        # 鼠标滚轮控件
        def _on_mousewheel(event):
            widget = event.widget
            while widget is not None:
                if isinstance(widget, tk.Canvas):
                    widget.yview_scroll(int(-1 * (event.delta / 120)), "units")
                    return
                widget = widget.master
        self.winfo_toplevel().bind_all("<MouseWheel>", _on_mousewheel)
        self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", _on_mousewheel))
        self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))

        # 定义间距属性
        self.padx = 10
        self.pady = 5
        self.label_width = 20

        # ============ 主体设置 ============
        main_frame = tb.Frame(self.inner)
        main_frame.grid(row=0, column=0, sticky="nsew", padx=15, pady=10)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)

        # ========= 风荷载 =========
        self.wind_frame = tb.LabelFrame(main_frame, text="风荷载", padding=10)
        self.wind_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=5)

        self.wind_params = tb.Frame(self.wind_frame)
        self.wind_params.grid(row=0, column=0, sticky="nsew", padx=10, pady=5)

        # 下拉列表
        DBFL_lst = ["A:海面、海岸、开阔水面", "B:田野、乡村、丛林、平坦开阔地", "C:树木及地层建筑密集区、平缓丘陵地", "D:中高层建筑密集区、起伏较大的丘陵地"]
        formula_lst = ["Ud = kf·kt·kh·U10", "Ud = kf·(Z/10)^α0·Us10"]
        # 参数输入框
        tb.Label(self.wind_params, text="基本风速(m/s):", width = self.label_width).grid(row=0, column=0, sticky="e", padx=(0,5), pady=3)
        self.gongzuo_wind_entry = tb.Entry(self.wind_params, width=10)
        self.gongzuo_wind_entry.insert(0, "20.70")
        self.gongzuo_wind_entry.grid(row=0, column=1, sticky="w", pady=3)
        tb.Label(self.wind_params, text="地表分类:", width = self.label_width).grid(row=1, column=0, sticky="e", padx=(0,5), pady=3)
        self.wind_combo1 = tb.Combobox(self.wind_params, values=DBFL_lst, width=25, state="readonly")
        self.wind_combo1.current(0)
        self.wind_combo1.grid(row=1, column=1, sticky="w", padx=(0,10), pady=3)
        tb.Label(self.wind_params, text="地形条件系数(m):", width = self.label_width).grid(row=2, column=0, sticky="e", padx=(0,5), pady=3)
        self.kt_entry = tb.Entry(self.wind_params, width=10)
        self.kt_entry.insert(0, "1.0")
        self.kt_entry.grid(row=2, column=1, sticky="w", pady=3)
        tb.Label(self.wind_params, text="主梁基准高度(m):", width = self.label_width).grid(row=3, column=0, sticky="e", padx=(0,5), pady=3)
        self.H_entry = tb.Entry(self.wind_params, width=10)
        self.H_entry.insert(0, "6.4")
        self.H_entry.grid(row=3, column=1, sticky="w", pady=3)
        tb.Label(self.wind_params, text="主梁横向力系数:", width = self.label_width).grid(row=4, column=0, sticky="e", padx=(0,5), pady=3)
        self.kh_entry = tb.Entry(self.wind_params, width=10)
        self.kh_entry.insert(0, "1.17")
        self.kh_entry.grid(row=4, column=1, sticky="w", pady=3)
        tb.Label(self.wind_params, text="设计基准风速计算公式:", width = self.label_width).grid(
        row=5, column=0, sticky="e", padx=(0,5), pady=3)
        self.wind_combo2 = tb.Combobox(self.wind_params, values=formula_lst, width=14, state="readonly")
        self.wind_combo2.current(0)
        self.wind_combo2.grid(row=5, column=1, sticky="w", padx=(0,10), pady=3)

        tb.Label(
            self.wind_frame,
            text="风荷载按照《公路桥梁抗风设计规范》（JTG/T 3360-01—2018）计算。",
            foreground="gray",
        ).grid(row=1, column=0, columnspan=2, sticky="w", padx=self.padx, pady=self.pady)

        # ========= 水流力 =========
        self.water_frame = tb.LabelFrame(main_frame, text="水流力", padding=10)
        self.water_frame.grid(row=1,column=0, sticky="nsew", padx=10, pady=5)

        water_params = tb.Frame(self.water_frame)
        water_params.grid(row=0, column=0, sticky="nsew", padx=10, pady=5)

        # 参数输入框
        tb.Label(water_params, text="设计流速(m/s):", width = self.label_width).grid(row=0, column=0, sticky="e", padx=(0,5), pady=3)
        self.flow_velocity_entry = tb.Entry(water_params, width=10)
        self.flow_velocity_entry.insert(0, "2.7")
        self.flow_velocity_entry.grid(row=0, column=1, sticky="w", pady=3)
        tb.Label(water_params, text="设防水位(m):", width = self.label_width).grid(row=1, column=0, sticky="e", padx=(0,5), pady=3)
        self.water_level_entry = tb.Entry(water_params, width=10)
        self.water_level_entry.insert(0, "+12.0")
        self.water_level_entry.grid(row=1, column=1, sticky="w", pady=3)
        tb.Label(water_params, text="入水深度(m):", width = self.label_width).grid(row=2, column=0, sticky="e", padx=(0,5), pady=3)
        self.water_depth_entry = tb.Entry(water_params, width=10)
        self.water_depth_entry.insert(0, "30.16")
        self.water_depth_entry.grid(row=2, column=1, sticky="w", pady=3)
        tb.Label(
            self.water_frame,
            text="水流力按照《港口工程荷载规范》（JTS 144-1—2010）计算。",
            foreground="gray",
        ).grid(row=1, column=0, columnspan=2, sticky="w", padx=self.padx, pady=self.pady)
        
        self.wind_frame.grid_columnconfigure(0, minsize=self.label_width)
        self.water_frame.grid_columnconfigure(0, minsize=self.label_width)


        # =========钻机荷载 =========
        self.drill_frame = tb.LabelFrame(main_frame, text="钻机参数设置", padding=10)
        self.drill_frame.grid(row=2, column=0, columnspan=2, sticky="nsew", padx=10, pady=5)

        tb.Label(self.drill_frame, text="钻机名称:",width = self.label_width).grid(row=0, column=0, sticky="e", padx=(0,5), pady=3)
        self.drill_name_entry = tb.Entry(self.drill_frame, width=15)
        self.drill_name_entry.insert(0, "XR550D")
        self.drill_name_entry.grid(row=0, column=1, sticky="w", pady=3)

        tb.Label(self.drill_frame, text="钻机工作重量(t):",width = self.label_width).grid(row=1, column=0, sticky="e", padx=(0,5), pady=3)
        self.drill_weight_entry = tb.Entry(self.drill_frame, width=15)
        self.drill_weight_entry.insert(0, "180")
        self.drill_weight_entry.grid(row=1, column=1, sticky="w", pady=3)
        
        tb.Label(self.drill_frame, text="履带长度(mm):",width = self.label_width).grid(row=2, column=0, sticky="e", padx=(0,5), pady=3)
        self.ld_length_entry = tb.Entry(self.drill_frame, width=15)
        self.ld_length_entry.insert(0, "6870")
        self.ld_length_entry.grid(row=2, column=1, sticky="w", pady=3)

        tb.Label(self.drill_frame, text="履带宽度(mm):",width = self.label_width).grid(row=3, column=0, sticky="e", padx=(0,5), pady=3)
        self.ld_width_entry = tb.Entry(self.drill_frame, width=15)
        self.ld_width_entry.insert(0, "1000")
        self.ld_width_entry.grid(row=3, column=1, sticky="w", pady=3)

        tb.Label(self.drill_frame, text="履带轴距(mm):",width = self.label_width).grid(row=4, column=0, sticky="e", padx=(0,5), pady=3)
        self.ld_axis_entry = tb.Entry(self.drill_frame, width=15)
        self.ld_axis_entry.insert(0, "5000")
        self.ld_axis_entry.grid(row=4, column=1, sticky="w", pady=3)

        tb.Label(self.drill_frame, text="钻孔半径(mm):",width = self.label_width).grid(row=5, column=0, sticky="e", padx=(0,5), pady=3)
        self.drill_radius_entry = tb.Entry(self.drill_frame, width=15)
        self.drill_radius_entry.insert(0, "5250")
        self.drill_radius_entry.grid(row=5, column=1, sticky="w", pady=3)
    
# ============================================钻孔区域设置 Tab 页========================================     
class drillAreaTab(tb.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        # ============ 滚动设置 ============
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.scrollbar = tb.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")

        # 内层 Frame
        self.inner = tb.Frame(self.canvas)
        self.inner_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        # 确保滚动区域随内容更新
        def on_frame_configure(event):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.inner.bind("<Configure>", on_frame_configure)

        # 保证内部 Frame 宽度跟 Canvas 一样
        def on_canvas_resize(event):
            self.canvas.itemconfig(self.inner_id, width=event.width)
        self.canvas.bind("<Configure>", on_canvas_resize)

        # 鼠标滚轮控件
        def _on_mousewheel(event):
            widget = event.widget
            while widget is not None:
                if isinstance(widget, tk.Canvas):
                    widget.yview_scroll(int(-1 * (event.delta / 120)), "units")
                    return
                widget = widget.master
        self.winfo_toplevel().bind_all("<MouseWheel>", _on_mousewheel)
        self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", _on_mousewheel))
        self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))

        # 定义间距属性
        self.padx = 10
        self.pady = 5

        # ============ 主体设置 ============
        self.main_frame = tb.Frame(self.inner)
        self.main_frame.grid(row=0, column=0, sticky="nsew", padx=15, pady=10)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(1, weight=1)

        # ========= 桩位确定 =========
        self.location_frame = tb.LabelFrame(
            self.main_frame,
            text="承台中心定位",
            padding=10
        )
        self.location_frame.grid(row=0, column=0, sticky="we", padx=10, pady=10)

        # 标题行
        header_location = tb.Frame(self.location_frame)
        header_location.grid(row=0, column=0, sticky="w", pady=(0, 5))
        
        tb.Label(
            header_location, 
            text="请设置承台中心相对起始端小肋中点的平面内坐标", 
            font=("微软雅黑", 10, "bold")
        ).grid(row=0, column=0, sticky="w")

        coord_frame = tb.Frame(self.location_frame)
        coord_frame.grid(row=1, column=0, sticky="w") 

        tb.Label(coord_frame, text="X 坐标(mm):")\
            .grid(row=0, column=0, sticky="e", padx=3, pady=3)
        self.x_entry = tb.Entry(coord_frame, width=10)
        self.x_entry.grid(row=0, column=1, padx=3, pady=3, sticky="w")
        self.x_entry.insert(0, "11800")

        tb.Label(coord_frame, text="Y 坐标(mm):")\
            .grid(row=0, column=2, sticky="e", padx=3, pady=3)
        self.y_entry = tb.Entry(coord_frame, width=10)
        self.y_entry.grid(row=0, column=3, padx=3, pady=3, sticky="w")
        self.y_entry.insert(0, "0")

 
        # ========= 桩位定义 =========
        self.pile_frame = tb.LabelFrame(
            self.main_frame,
            text="钻孔桩位平面布置",
            padding=10
        )
        self.pile_frame.grid(row=1, column=0, sticky="we", padx=10, pady=10)

        # 标题行
        header_pile = tb.Frame(self.pile_frame)
        header_pile.grid(row=0, column=0, sticky="w", pady=(0, 5))

        tb.Label(
            header_pile, 
            text="钻孔桩位定义（请以承台中心为基点设置平面内坐标）", 
            font=("微软雅黑", 10, "bold")
        ).grid(row=0, column=0, sticky="w")

        add_pile_btn = tb.Button(header_pile, text="+ New", bootstyle="info-outline")
        add_pile_btn.grid(row=0, column=1, sticky="e", padx=10)

        # 容器
        self.pile_list = []

        def add_pile():
            idx = len(self.pile_list) + 1
            pile = tb.LabelFrame(self.pile_frame, text=f"桩位{idx}", padding=10)
            pile.grid(row=idx, column=0, sticky="we", pady=5)

            # X 坐标
            tb.Label(pile, text="X 坐标(mm):").grid(row=0, column=0, sticky="e", padx=3, pady=3)
            x_entry = tb.Entry(pile, width=10)
            x_entry.grid(row=0, column=1, padx=3, pady=3)

            # Y 坐标
            tb.Label(pile, text="Y 坐标(mm):").grid(row=0, column=2, sticky="e", padx=3, pady=3)
            y_entry = tb.Entry(pile, width=10)
            y_entry.grid(row=0, column=3, padx=3, pady=3)

            # 初始值
            dx, dy = default_coords[idx-1] if idx <= len(default_coords) else (0, 0)
            x_entry.insert(0, str(dx))
            y_entry.insert(0, str(dy))

            # 删除按钮
            del_btn = tb.Button(
                pile, text="删除",
                bootstyle="danger-outline",
                command=lambda: delete_pile(pile)
            )
            del_btn.grid(row=0, column=4, rowspan=1, padx=10)

            self.pile_list.append({
                "frame": pile,
                "x": x_entry,
                "y": y_entry
            })
            
            # 更新滚动区
            self.inner.update_idletasks()
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        def delete_pile(frame):
            frame.destroy()
            # 从 list 中删除对应项
            for i in self.pile_list:
                if i["frame"] == frame:
                    self.pile_list.remove(i)
                    break

            # 更新滚动区
            self.inner.update_idletasks()
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        add_pile_btn.config(command=add_pile)

        # 添加 9 个桩位
        default_coords = [
            (-5000, 5000),
            (0, 5000),
            (5000, 5000),
            (-5000, 0),
            (0,0),
            (5000,0),
            (-5000,-5000),
            (0, -5000),
            (5000, -5000),
        ]
        for _ in range(9):
            add_pile()


# ------------- 主界面类不变，仅 ComponentsTab 内部边框固定 -------------
class PlatformUI:
    """主界面"""

    def __init__(self, root):
        self.root = root
        self.root.title("钻孔平台自动化程序")
        self.root.geometry("850x1000")

        # ====== 整体框架：上部tab滚动区域 + 下部固定操作栏 ======
        main_frame = tb.Frame(self.root)
        main_frame.pack(fill="both", expand=True)

        # ====== Notebook 直接放在主框架中（不滚动） ======
        notebook = tb.Notebook(main_frame)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # 各标签页
        self.tab_components = ComponentsTab(notebook)
        notebook.add(self.tab_components, text="构件设置")

        self.tab_section = SectionTab(notebook)
        notebook.add(self.tab_section, text="截面设置")

        self.tab_loads = LoadTab(notebook)
        notebook.add(self.tab_loads, text="荷载参数")

        self.tab_drillarea = drillAreaTab(notebook)
        notebook.add(self.tab_drillarea, text="钻孔区域设置")

        self._create_bottom_bar()

    def choose_save_path(self):
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

    def _create_bottom_bar(self):
        bottom_frame = tb.Frame(self.root)
        bottom_frame.pack(fill="x", pady=5)

        path_frame = tb.Frame(bottom_frame)
        path_frame.pack(fill="x", pady=3)

        # -------- 读取注册表默认路径 --------
        install_path = read_Register("Software\\ShuZhiQiaoShi", "Applocation")
        if not install_path:
            install_path = os.getcwd()     # 若注册表不存在，则用当前目录

        default_mct_name = "Platform_Untitled.mct"
        default_mct_path = os.path.join(install_path, "Custom", default_mct_name)

        # -------- 保存路径 StringVar --------
        self.save_path_var = tk.StringVar(value=default_mct_path)

        # -------- 路径选择行（新的 UI）--------
        dialog_savepath(
            parent=bottom_frame,
            label_txt="MCT 文件保存路径：",
            entry_var=self.save_path_var,
            button_txt="更改路径",
            button_command=self.choose_save_path
        )


        btn_frame = tb.Frame(bottom_frame)
        btn_frame.pack(fill="x", pady=3)

        self.run_mct_btn = tb.Button(btn_frame, text="生成MCT", bootstyle=SUCCESS)
        self.run_mct_btn.pack(side="right", padx=10)

        tb.Button(btn_frame, text="退出", bootstyle=DANGER, command=self.root.destroy).pack(side="right", padx=10)

