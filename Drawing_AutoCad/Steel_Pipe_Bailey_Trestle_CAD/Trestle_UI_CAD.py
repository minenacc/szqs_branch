# 1. 标准库
import os
import sys
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import messagebox

# 2. 第三方库
import win32com.client
import ttkbootstrap as tb
from PIL import Image, ImageTk 
from ttkbootstrap.constants import *

# 3. 本地模块
sys.path.append(str(Path(__file__).parent.parent.parent))
from General.UIHandle       import if_Reg
from General.DataUtils      import read_Register
from General.AutoCAD        import Import_Blocks, find_installed_autocad_versions, AutoCAD_VersionR_tans_to_COM_ProgID

sys.path.append(str(Path(__file__).parent))
from Drawing_AutoCad.Steel_Pipe_Bailey_Trestle_CAD.Trestle_CAD_draw import All


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
        self.padx = 5
        self.pady = 5
        self.label_width = 22
        self.entry_width = 40

        # ======================== 提示 ========================
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
        self.entry_smallrib.insert(0, "68@750")
        self.entry_smallrib.grid(row=0, column=1, padx=self.padx, pady=self.pady, sticky="w")
        
        tb.Label(lf_smallrib, text="纵向起始端间距(mm):", width = self.label_width).grid(row=1, column=0, sticky="w", padx=self.padx, pady=self.pady)
        self.entry_smallrib_init_spacing = tb.Entry(lf_smallrib, width=self.entry_width,foreground="gray")
        self.entry_smallrib_init_spacing.grid(row=1, column=1, padx=self.padx, pady=self.pady, sticky="w")
        self._add_placeholder(self.entry_smallrib_init_spacing, "请输入第一根小肋相对贝雷梁起始端的纵向起始间距")
        
        tb.Label(lf_smallrib, text="小肋长度(mm):", width = self.label_width).grid(row=2, column=0, sticky="w", padx=self.padx, pady=self.pady)
        self.entry_smallrib_length = tb.Entry(lf_smallrib, width=self.entry_width)
        self.entry_smallrib_length.insert(0, "9000")
        self.entry_smallrib_length.grid(row=2, column=1, padx=self.padx, pady=self.pady, sticky="w")

    def _create_bailey_section(self):
        lf = tb.LabelFrame(self.inner, text="贝雷梁定义", padding=10)
        lf.pack(fill="x", padx=10, pady=self.pady)

        tb.Label(lf, text="贝雷梁横向布置(mm):", width = self.label_width).grid(row=0, column=0, sticky="w", padx=self.padx, pady=self.pady)
        self.entry_bailey_spacing = tb.Entry(lf, width=self.entry_width)
        self.entry_bailey_spacing.insert(0, "3@450,5@900,3@450")
        self.entry_bailey_spacing.grid(row=0, column=1, padx=self.padx, pady=self.pady, sticky="w")

        tb.Label(lf, text="贝雷梁纵向布置(m):", width = self.label_width).grid(row=1, column=0, sticky="w", padx=self.padx, pady=self.pady)
        self.entry_bailey_span = tb.Entry(lf, width=self.entry_width)
        self.entry_bailey_span.insert(0, "3+4@12")
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
        self.entry_distribute_length.insert(0, "8000")
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
        # ================= 下部结构定义 =================
        lf = tb.LabelFrame(self.inner, text="下部结构定义", padding=10)
        lf.pack(fill="x", padx=10, pady=5) 

        # ========== 栈桥端下部结构设置 ==========
        lf_side = tb.LabelFrame(lf, text="栈桥端下部结构设置", padding=10)
        lf_side.grid(row=0, column=0, sticky="w", padx=5, pady=5)

        # 左端
        self.left_frame = tb.LabelFrame(lf_side, text="左端设置", padding=10)
        self.left_frame.grid(row=0, column=0, sticky="w", padx=5, pady=5)

        self.left_select_var = tk.StringVar(value="钢管桩")
        self.left_select = tb.Combobox(
            self.left_frame,
            values=["钢管桩", "制动墩", "桥台"],
            textvariable=self.left_select_var,
            width=15,
            state="readonly"
        )
        self.left_select.grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.left_select.bind("<<ComboboxSelected>>", self.on_left_select)

        # ====== 三类内容 ======
        # 钢管桩
        self.left_frame_ggz = tb.Frame(self.left_frame)
        tb.Label(self.left_frame_ggz, text="悬臂长度(mm):").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.left_entry_ggz_cantilever = tb.Entry(self.left_frame_ggz, width=40, foreground="gray")
        self.left_entry_ggz_cantilever.grid(row=0, column=1, padx=5,sticky="w")
        self._add_placeholder(self.left_entry_ggz_cantilever, "请输入贝雷梁起始端相对左端钢管桩的悬臂长度")

        # 制动墩
        self.left_frame_zdd = tb.Frame(self.left_frame)

        tb.Label(self.left_frame_zdd, text="分配梁F1(上层)长度(mm):", width = 28).grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.left_entry_f1_length = tb.Entry(self.left_frame_zdd, width=40)
        self.left_entry_f1_length.insert(0, "7600")
        self.left_entry_f1_length.grid(row=0, column=1, padx=5,sticky="w")
        tb.Label(self.left_frame_zdd, text="分配梁F1(上层)悬臂长度-贝雷梁(mm):", width = 28).grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.left_entry_f1_cantilever1 = tb.Entry(self.left_frame_zdd, width=40, foreground="gray")
        self.left_entry_f1_cantilever1.grid(row=1, column=1,padx=self.padx, pady=self.pady, sticky="w")
        self._add_placeholder(self.left_entry_f1_cantilever1, "请输入上层分配梁起始端相对第一片贝雷梁的悬臂长度")
        tb.Label(self.left_frame_zdd, text="分配梁F1(上层)悬臂长度-F2(mm):", width = 28).grid(row=2, column=0, sticky="w", padx=5, pady=2)
        self.left_entry_f1_cantilever2 = tb.Entry(self.left_frame_zdd, width=40, foreground="gray")
        self.left_entry_f1_cantilever2.grid(row=2, column=1,   padx=self.padx, pady=self.pady, sticky="w")
        self._add_placeholder(self.left_entry_f1_cantilever2, "请输入上层分配梁起始端相对下层分配梁的悬臂长度")

        tb.Label(self.left_frame_zdd, text="分配梁F2(下层)长度(mm):", width = 28).grid(row=3, column=0, sticky="w", padx=5, pady=2)
        self.left_entry_f2_length = tb.Entry(self.left_frame_zdd, width=40)
        self.left_entry_f2_length.insert(0, "4600")
        self.left_entry_f2_length.grid(row=3, column=1, padx=5,sticky="w")
        tb.Label(self.left_frame_zdd, text="分配梁F2(下层)悬臂长度-F1(mm):", width = 28).grid(row=4, column=0, sticky="w", padx=5, pady=2)
        self.left_entry_f2_cantilever1 = tb.Entry(self.left_frame_zdd, width=40, foreground="gray")
        self.left_entry_f2_cantilever1.grid(row=4, column=1, padx=self.padx, pady=self.pady, sticky="w")
        self._add_placeholder(self.left_entry_f2_cantilever1, "请输入下层分配梁起始端相对上层分配梁的悬臂长度")
        tb.Label(self.left_frame_zdd, text="分配梁F2(下层)悬臂长度-钢管桩(mm):", width = 28).grid(row=5, column=0, sticky="w", padx=5, pady=2)
        self.left_entry_f2_cantilever2 = tb.Entry(self.left_frame_zdd, width=40, foreground="gray")
        self.left_entry_f2_cantilever2.grid(row=5, column=1,   padx=self.padx, pady=self.pady, sticky="w")
        self._add_placeholder(self.left_entry_f2_cantilever2, "请输入下层分配梁起始端相对制动墩钢管桩的悬臂长度")

        tb.Label(self.left_frame_zdd, text="制动墩横向间距(mm):", width = 28).grid(row=6, column=0, sticky="w", padx=5, pady=2)
        self.left_entry_zdd_y = tb.Entry(self.left_frame_zdd, width=40)
        self.left_entry_zdd_y.insert(0, "3000+3000")
        self.left_entry_zdd_y.grid(row=6, column=1, padx=5,sticky="w")
        tb.Label(self.left_frame_zdd, text="制动墩纵向间距(mm):", width = 28).grid(row=7, column=0, sticky="w", padx=5, pady=2)
        self.left_entry_zdd_x = tb.Entry(self.left_frame_zdd, width=40)
        self.left_entry_zdd_x.insert(0, "3000")
        self.left_entry_zdd_x.grid(row=7, column=1, padx=5,sticky="w") 
        tb.Label(self.left_frame_zdd, text="制动墩钢管桩长度(m):", width = 28).grid(row=8, column=0, sticky="w", padx=5, pady=2)
        self.left_entry_zdd_length = tb.Entry(self.left_frame_zdd, width=40)
        self.left_entry_zdd_length.insert(0, "330")
        self.left_entry_zdd_length.grid(row=8, column=1, padx=5,sticky="w")

        # 桥台
        self.left_frame_abutment = tb.Frame(self.left_frame)

        # 布局
        self.left_frame_ggz.grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.left_frame_zdd.grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.left_frame_abutment.grid(row=1, column=0, sticky="w", padx=5, pady=5)
        
        # 默认显示钢管桩
        self.left_frame_zdd.grid_remove()
        self.left_frame_abutment.grid_remove()
    

        # ====================== 右端设置 ======================
        self.right_frame = tb.LabelFrame(lf_side, text="右端设置", padding=10)
        self.right_frame.grid(row=1, column=0, sticky="w", padx=5, pady=5)

        self.right_select_var = tk.StringVar(value="钢管桩")
        self.right_select = tb.Combobox(
            self.right_frame,
            values=["钢管桩", "制动墩", "桥台"],
            textvariable=self.right_select_var,
            width=15,
            state="readonly"
        )
        self.right_select.grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.right_select.bind("<<ComboboxSelected>>", self.on_right_select)

        # ====== 三类内容 ======
        # 钢管桩
        self.right_frame_ggz = tb.Frame(self.right_frame)
        tb.Label(self.right_frame_ggz, text="悬臂长度(mm):").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.right_entry_ggz_cantilever = tb.Entry(self.right_frame_ggz, width=40, foreground="gray")
        self.right_entry_ggz_cantilever.grid(row=0, column=1, padx=5,sticky="w")
        self._add_placeholder(self.right_entry_ggz_cantilever, "请输入贝雷梁起始端相对右端钢管桩的悬臂长度")

        # 制动墩
        self.right_frame_zdd = tb.Frame(self.right_frame)

        tb.Label(self.right_frame_zdd, text="分配梁F1(上层)长度(mm):", width = 28).grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.right_entry_f1_length = tb.Entry(self.right_frame_zdd, width=40)
        self.right_entry_f1_length.insert(0, "7600")
        self.right_entry_f1_length.grid(row=0, column=1, padx=5,sticky="w")
        tb.Label(self.right_frame_zdd, text="分配梁F1(上层)悬臂长度-贝雷梁(mm):", width = 28).grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.right_entry_f1_cantilever1 = tb.Entry(self.right_frame_zdd, width=40, foreground="gray")
        self.right_entry_f1_cantilever1.grid(row=1, column=1,padx=self.padx, pady=self.pady, sticky="w")
        self._add_placeholder(self.right_entry_f1_cantilever1, "请输入上层分配梁起始端相对第一片贝雷梁的悬臂长度")
        tb.Label(self.right_frame_zdd, text="分配梁F1(上层)悬臂长度-F2(mm):", width = 28).grid(row=2, column=0, sticky="w", padx=5, pady=2)
        self.right_entry_f1_cantilever2 = tb.Entry(self.right_frame_zdd, width=40, foreground="gray")
        self.right_entry_f1_cantilever2.grid(row=2, column=1, padx=self.padx, pady=self.pady, sticky="w")
        self._add_placeholder(self.right_entry_f1_cantilever2, "请输入上层分配梁起始端相对下层分配梁的悬臂长度")

        tb.Label(self.right_frame_zdd, text="分配梁F2(下层)长度(mm):", width = 28).grid(row=3, column=0, sticky="w", padx=5, pady=2)
        self.right_entry_f2_length = tb.Entry(self.right_frame_zdd, width=40)
        self.right_entry_f2_length.insert(0, "4600")
        self.right_entry_f2_length.grid(row=3, column=1, padx=5,sticky="w")
        tb.Label(self.right_frame_zdd, text="分配梁F2(下层)悬臂长度-F1(mm):", width = 28).grid(row=4, column=0, sticky="w", padx=5, pady=2)
        self.right_entry_f2_cantilever1 = tb.Entry(self.right_frame_zdd, width=40, foreground="gray")
        self.right_entry_f2_cantilever1.grid(row=4, column=1,   padx=self.padx, pady=self.pady, sticky="w")
        self._add_placeholder(self.right_entry_f2_cantilever1, "请输入下层分配梁起始端相对上层分配梁的悬臂长度")
        tb.Label(self.right_frame_zdd, text="分配梁F2(下层)悬臂长度-钢管桩(mm):", width = 28).grid(row=5, column=0, sticky="w", padx=5, pady=2)
        self.right_entry_f2_cantilever2 = tb.Entry(self.right_frame_zdd, width=40, foreground="gray")
        self.right_entry_f2_cantilever2.grid(row=5, column=1,   padx=self.padx, pady=self.pady, sticky="w")
        self._add_placeholder(self.right_entry_f2_cantilever2, "请输入下层分配梁起始端相对制动墩钢管桩的悬臂长度")

        tb.Label(self.right_frame_zdd, text="制动墩横向间距(mm):", width = 28).grid(row=6, column=0, sticky="w", padx=5, pady=2)
        self.right_entry_zdd_y = tb.Entry(self.right_frame_zdd, width=40)
        self.right_entry_zdd_y.insert(0, "3000+3000")
        self.right_entry_zdd_y.grid(row=6, column=1, padx=5,sticky="w")
        tb.Label(self.right_frame_zdd, text="制动墩纵向间距(mm):", width = 28).grid(row=7, column=0, sticky="w", padx=5, pady=2)
        self.right_entry_zdd_x = tb.Entry(self.right_frame_zdd, width=40)
        self.right_entry_zdd_x.insert(0, "3000")
        self.right_entry_zdd_x.grid(row=7, column=1, padx=5,sticky="w") 
        tb.Label(self.right_frame_zdd, text="制动墩钢管桩长度(m):", width = 28).grid(row=8, column=0, sticky="w", padx=5, pady=2)
        self.right_entry_zdd_length = tb.Entry(self.right_frame_zdd, width=40)
        self.right_entry_zdd_length.insert(0, "33.0")
        self.right_entry_zdd_length.grid(row=8, column=1, padx=5,sticky="w")

        # 桥台
        self.right_frame_abutment = tb.Frame(self.right_frame)

        # 布局注册（同一位置，三选一显示）
        self.right_frame_ggz.grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.right_frame_zdd.grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.right_frame_abutment.grid(row=1, column=0, sticky="w", padx=5, pady=5)

        # 默认显示钢管桩
        self.right_frame_zdd.grid_remove()
        self.right_frame_abutment.grid_remove()
        # 右端提示文字（默认隐藏）
        self.right_zdd_tip = tb.Label(
            self.right_frame,
            text="请在左端设置制动墩参数，右端无需重复设置",
            foreground="gray"
        )
        self.right_zdd_tip.grid(row=2, column=0, sticky="w", padx=5, pady=3)
        self.right_zdd_tip.grid_remove()
        
        # ========== 钢管桩设置 ==========
        lf_pile = tb.LabelFrame(lf, text="钢管桩设置", padding=10)
        lf_pile.grid(row=1, column=0, sticky="w", padx=5, pady=5)

        self.label_pile_y = tb.Label(lf_pile, text="钢管桩横向间距(mm):", width = self.label_width)
        self.label_pile_y.grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.entry_pile_y = tb.Entry(lf_pile, width=15)
        self.entry_pile_y.insert(0, "3000+3000")
        self.entry_pile_y.grid(row=0, column=1, padx=5)

        self.label_pile_length = tb.Label(lf_pile, text="桩长(m):", width = self.label_width)
        self.label_pile_length.grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.entry_pile_length = tb.Entry(lf_pile, width=15)
        self.entry_pile_length.insert(0, "33.0")
        self.entry_pile_length.grid(row=1, column=1, padx=5)

        # ========== 联结系 ==========
        lf_connect = tb.LabelFrame(lf, text="联结系设置", padding=10)
        lf_connect.grid(row=2, column=0, sticky="w", padx=5, pady=5)

        self.label_connect_zd = tb.Label(lf_connect, text="联结系距桩顶间距(mm):", width = self.label_width)
        self.label_connect_zd.grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.entry_connect_zd = tb.Entry(lf_connect, width=15)
        self.entry_connect_zd.insert(0, "620")
        self.entry_connect_zd.grid(row=0, column=1, padx=5)

        self.label_connect_z = tb.Label(lf_connect, text="联结系竖向间距(mm):", width = self.label_width)
        self.label_connect_z.grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.entry_connect_z = tb.Entry(lf_connect, width=15)
        self.entry_connect_z.insert(0, "2500")
        self.entry_connect_z.grid(row=1, column=1, padx=5)

    # ============= 选项改变回调 =============
    def on_left_select(self, event=None):
        self.refresh_under_section()
        if hasattr(self, "on_support_changed"):
            has_brake = (self.left_select_var.get() == "制动墩") or \
                    (self.right_select_var.get() == "制动墩")
            self.on_support_changed(has_brake)
            
    def on_right_select(self, event=None):
        self.refresh_under_section()
        if hasattr(self, "on_support_changed"):        
            has_brake = (self.left_select_var.get() == "制动墩") or \
                    (self.right_select_var.get() == "制动墩")
            self.on_support_changed(has_brake)
            
    def refresh_under_section(self):
        sel_left  = self.left_select_var.get()
        sel_right = self.right_select_var.get()

        # 先隐藏全部
        self.left_frame_ggz.grid_remove()
        self.left_frame_zdd.grid_remove()
        self.left_frame_abutment.grid_remove()

        self.right_frame_ggz.grid_remove()
        self.right_frame_zdd.grid_remove()
        self.right_frame_abutment.grid_remove()
        self.right_zdd_tip.grid_remove()

        # 左端始终正常显示
        if sel_left == "钢管桩":
            self.left_frame_ggz.grid()
        elif sel_left == "制动墩":
            self.left_frame_zdd.grid()
        elif sel_left == "桥台":
            self.left_frame_abutment.grid()

        # 右端逻辑
        if sel_right == "制动墩":
            if sel_left == "制动墩":
                # 两端都制动墩 → 固定左端可填 → 隐藏右端参数 + 提示
                self.right_zdd_tip.grid()
            else:
                self.right_frame_zdd.grid()
        else:
            # 非制动墩 → 正常显示
            if sel_right == "钢管桩":
                self.right_frame_ggz.grid()
            elif sel_right == "桥台":
                self.right_frame_abutment.grid()
    


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

        # 默认无钢管桩
        self.brake_enabled = True

        # 选择构件
        component_frame = tb.Frame(self)
        component_frame.grid(row=0, column=0, sticky="w", padx=15, pady=(10, 5))
        component_frame.columnconfigure((0,1,2,3,4,5), weight=1)

        tb.Label(component_frame, text="选择构件：").grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.component_list = ["小肋", "分配梁", "钢管桩", "联结系", "制动墩钢管桩", "制动墩分配梁F1", "制动墩分配梁F2"]
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
        if section_type != "管型截面":
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
        if name in ["钢管桩","制动墩钢管桩"]:
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
    
    def set_brake_enabled(self, has_brake_pier: bool):
        """
        根据是否存在制动墩，动态显示/隐藏制动墩相关构件按钮。
        当 ComponentsTab 勾选“不设置制动墩”时会自动触发此函数。
        """
        self.has_brake_pier = has_brake_pier

        brake_components = ["制动墩钢管桩", "制动墩分配梁F1", "制动墩分配梁F2"]

        for name in brake_components:
            btn = self.component_buttons.get(name)
            if btn:
                if has_brake_pier:
                    btn.grid() 
                else:
                    btn.grid_remove() 

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


        # ========= 移动荷载 =========
        self.move_frame = tb.LabelFrame(main_frame, text="移动荷载", padding=10)
        self.move_frame.grid(row=2, column=0, columnspan=2, sticky="nsew", padx=10, pady=5)


        # ==设置车道==
        self.chedao_frame = tb.Frame(self.move_frame)
        self.chedao_frame.grid(row=0, column=0, sticky="w", padx=self.padx, pady=3)
        header1 = tb.Frame(self.chedao_frame)
        header1.grid(row=0, column=0, sticky="w", pady=(0, 5))
        tb.Label(header1, text="车道定义", font=("微软雅黑", 10, "bold")).grid(row=0, column=0, sticky="w")
        add_lane_btn = tb.Button(header1, text="+ New", bootstyle="info-outline")
        add_lane_btn.grid(row=0, column=1, sticky="e", padx=10)

        self.lane_list = []

        def add_lane():
            idx = len(self.lane_list) + 1
            lane = tb.LabelFrame(self.chedao_frame, text=f"车道{idx}", padding=10)
            lane.grid(row=idx, column=0, sticky="w", pady=5)
            tb.Label(lane, text="车道名称:",width = self.label_width).grid(row=0, column=0, sticky="w", padx=3, pady=3)
            name = tb.Entry(lane); name.grid(row=0, column=1, padx=3, pady=3)
            tb.Label(lane, text="距桥面中心线距离(m):",width = self.label_width).grid(row=1, column=0, sticky="w", padx=3, pady=3)
            dist = tb.Entry(lane); dist.grid(row=1, column=1, padx=3, pady=3)
            tb.Label(lane, text="车轮间距(m):",width = self.label_width).grid(row=2, column=0, sticky="w", padx=3, pady=3)
            space = tb.Entry(lane); space.grid(row=2, column=1, padx=3, pady=3)
            tb.Label(lane, text="移动方向:",width = self.label_width).grid(row=3, column=0, sticky="w", padx=3, pady=3)
            direc = tb.Combobox(lane, values=["往返","向前","向后"], width=10)
            direc.set("往返")
            direc.grid(row=3, column=1, padx=3, pady=3)
            del_btn = tb.Button(lane, text="删除车道", bootstyle="danger-outline",
                                command=lambda: delete_lane(lane))
            del_btn.grid(row=0, column=2, padx=15, sticky="e")
            self.lane_list.append({"frame": lane, "name_entry": name, "dist": dist, "space": space, "direc": direc})
            self.inner.update_idletasks()
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        def delete_lane(frame):
            frame.destroy()
            for i in self.lane_list:
                if i["frame"] == frame:
                    self.lane_list.remove(i)
                    break
            self.inner.update_idletasks()
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        add_lane_btn.config(command=add_lane)
        add_lane()


        # ==设置车辆==
        self.vehicle_frame = tb.Frame(self.move_frame)
        self.vehicle_frame.grid(row=1, column=0, sticky="w", padx=self.padx, pady=3)
        header2 = tb.Frame(self.vehicle_frame)
        header2.grid(row=1, column=0, sticky="w", pady=(5, 2))
        tb.Label(header2, text="车辆定义", font=("微软雅黑", 10, "bold")).grid(row=0, column=0, sticky="w")
        add_vehicle_btn = tb.Button(header2, text="+ New", bootstyle="info-outline")
        add_vehicle_btn.grid(row=0, column=1, sticky="e", padx=10)

        self.vehicle_list = []

        def add_vehicle():
            idx = len(self.vehicle_list) + 1
            vehicle = tb.LabelFrame(self.vehicle_frame, text=f"车辆{idx}", padding=10)
            vehicle.grid(row=idx + 1, column=0, sticky="we", pady=5)

            # ===== 类型选择 =====
            tb.Label(vehicle, text="车辆荷载类型:",width = self.label_width).grid(row=0, column=0, sticky="e", padx=3, pady=3)
            type_var = tk.StringVar(value="城市桥梁车道荷载类型")
            type_combo = tb.Combobox(
                vehicle,
                textvariable=type_var,
                values=["城市桥梁车道荷载类型", "旧公路履带车荷载类型","用户定义车辆"],
                state="readonly",
                width=22
            )
            type_combo.grid(row=0, column=1, sticky="w", padx=3, pady=3)

            del_vehicle_btn = tb.Button(vehicle, text="删除车辆", bootstyle="danger-outline", width=8,
                                        command=lambda: delete_vehicle(vehicle))
            del_vehicle_btn.grid(row=0, column=5, padx=15, sticky="e")

            content_frame = tb.Frame(vehicle)
            content_frame.grid(row=1, column=0, columnspan=6, sticky="we", pady=(5, 0))

            # ===== 用户定义车辆 =====
            def normal_content(parent):
                tb.Label(parent, text="车辆荷载名称:",width = self.label_width).grid(row=0, column=0, sticky="e", padx=3, pady=3)
                name_entry = tb.Entry(parent, width=20)
                name_entry.insert(0, f"50t混凝土罐车")
                name_entry.grid(row=0, column=1, padx=3, pady=3)
                parent.name_entry = name_entry

                header_frame = tb.Frame(parent)
                header_frame.grid(row=2, column=0, columnspan=4, sticky="w", pady=(0,5))
                tb.Label(header_frame, text="编号", width=8, anchor="center").grid(row=0, column=0)
                tb.Label(header_frame, text="荷载(kN)", width=12, anchor="center").grid(row=0, column=1)
                tb.Label(header_frame, text="间距(mm)", width=12, anchor="center").grid(row=0, column=2)
                tb.Label(header_frame, text="操作", width=10, anchor="center").grid(row=0, column=3)

                row_list = []
                parent.row_list = row_list

                def add_vehicle_row( load_val="",dist_val="",):
                    ridx = len(row_list) + 1
                    row = tb.Frame(parent)
                    row.grid(row=3 + ridx, column=0, columnspan=4, sticky="w", pady=2)
                    tb.Label(row, text=str(ridx), width=8, anchor="center").grid(row=0, column=0, padx=2)
                    load_entry = tb.Entry(row, width=10, justify="center")
                    load_entry.insert(0, str(load_val))
                    load_entry.grid(row=0, column=1, padx=5)
                    dist_entry = tb.Entry(row, width=10, justify="center")
                    dist_entry.insert(0, str(dist_val))
                    dist_entry.grid(row=0, column=2, padx=5)
                  
                    del_btn = tb.Button(row, text="删除", bootstyle="danger-outline", width=5,
                                        command=lambda: delete_vehicle_row(row))
                    del_btn.grid(row=0, column=3, padx=3, sticky="e")
                    row_list.append({"frame": row,  "load": load_entry,"dist": dist_entry})
                    self.inner.update_idletasks()
                    self.canvas.configure(scrollregion=self.canvas.bbox("all"))

                def delete_vehicle_row(row):
                    row.destroy()
                    for r in row_list:
                        if r["frame"] == row:
                            row_list.remove(r)
                            break
                    for i, r in enumerate(row_list, start=1):
                        for widget in r["frame"].winfo_children():
                            if isinstance(widget, tb.Label):
                                widget.config(text=str(i))
                                break
                    self.inner.update_idletasks()
                    self.canvas.configure(scrollregion=self.canvas.bbox("all"))

                default_data = [
                    (80, 1700),
                    (80, 3175),
                    (170, 1350),
                    (170, 0),
                ]
                for load,dist in default_data:
                    add_vehicle_row(load,dist)

                tb.Button(parent, text="新增", bootstyle="info-outline", width=8,
                          command=add_vehicle_row).grid(row=999, column=0, columnspan=4, pady=(10, 5))
                
                self.install_path = read_Register("Software\\ShuZhiQiaoShi", "Applocation")
                if not self.install_path:
                    self.install_path = os.getcwd()

                self.png_path = os.path.join(self.install_path, "Support", "png_picture")
                img_path = os.path.join(self.png_path, "truck_load.png")
                img = Image.open(img_path)
                img_resized = img.resize((500, 200), Image.Resampling.LANCZOS)
                self.photo_img = ImageTk.PhotoImage(img_resized)
                img_label = tb.Label(content_frame, image=self.photo_img)
                img_label.grid(row=1000, column=0, columnspan=4, pady=(10, 5))
    
            def create_city_abutment_content(parent):
                """城市桥梁车道荷载类型 - 仅允许选择 CH-CL / CL-CD"""
                tb.Label(parent, text="车辆荷载名称:",width = self.label_width).grid(row=0, column=0, sticky="e", padx=3, pady=3)
                name_entry = tb.Entry(parent, width=20)
                name_entry.insert(0, "公路一级荷载")
                name_entry.grid(row=0, column=1, padx=3, pady=3)
                parent.name_entry = name_entry

                # ===== 标准荷载选择 =====
                tb.Label(parent, text="标准荷载类型:",width = self.label_width).grid(
                    row=1, column=0, sticky="e", padx=3, pady=(10, 5)
                )
                std_type = tb.Combobox(
                    parent,
                    values=["CH-CL", "CL-CD"],
                    state="readonly",
                    width=10
                )
                std_type.set("CH-CD")
                std_type.grid(row=1, column=1, sticky="w", padx=3, pady=(10, 5))
                parent.std_type = std_type

            # ===== 旧公路履带车荷载 =====
            def create_crawler_content(parent):
                tb.Label(parent, text="车辆荷载名称:",width = self.label_width).grid(row=0, column=0, sticky="e", padx=3, pady=3)
                name_entry = tb.Entry(parent, width=30)
                name_entry.insert(0, f"150t履带吊")
                name_entry.grid(row=0, column=1, padx=3, pady=3, sticky="w")
                parent.name_entry = name_entry

                tb.Label(parent, text="dW1(kN/m):",width = self.label_width).grid(row=1, column=0, sticky="e", padx=3, pady=3)
                dw1_entry = tb.Entry(parent, width=15, justify="center")
                dw1_entry.insert(0, "208.86")
                dw1_entry.grid(row=1, column=1, sticky="w", padx=3, pady=3)

                tb.Label(parent, text="dD1(m):",width = self.label_width).grid(row=2, column=0, sticky="e", padx=3, pady=3)
                dd1_entry = tb.Entry(parent, width=15, justify="center")
                dd1_entry.insert(0, "7.182")
                dd1_entry.grid(row=2, column=1, sticky="w", padx=3, pady=3)

                parent.dw1_entry = dw1_entry
                parent.dd1_entry = dd1_entry
                
            def load_vehicle_type(*_):
                for widget in content_frame.winfo_children():
                    widget.destroy()
                if type_var.get() == "城市桥梁车道荷载类型":
                    create_city_abutment_content(content_frame)
                elif type_var.get() == "旧公路履带车荷载类型":
                    create_crawler_content(content_frame)
                else:
                    normal_content(content_frame)
                self.inner.update_idletasks()
                self.canvas.configure(scrollregion=self.canvas.bbox("all"))

            type_combo.bind("<<ComboboxSelected>>", load_vehicle_type)
            load_vehicle_type()

            self.vehicle_list.append({"frame": vehicle, "type": type_var, "content": content_frame})
            self.inner.update_idletasks()
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        def delete_vehicle(frame):
            frame.destroy()
            for v in self.vehicle_list:
                if v["frame"] == frame:
                    self.vehicle_list.remove(v)
                    break
            self.inner.update_idletasks()
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        add_vehicle_btn.config(command=add_vehicle)
        add_vehicle()


        # ==设置工况==
        self.gongkuang_frame = tb.Frame(self.move_frame)
        self.gongkuang_frame.grid(row=3, column=0, sticky="w", padx=self.padx, pady=3)
        header3 = tb.Frame(self.gongkuang_frame)
        header3.grid(row=0, column=0, sticky="we", pady=(5, 2))
        tb.Label(header3, text="移动荷载工况组合", font=("微软雅黑", 10, "bold")).grid(row=0, column=0, sticky="w")
        add_case_btn = tb.Button(header3, text="+ New", bootstyle="info-outline")
        add_case_btn.grid(row=0, column=1, sticky="e", padx=10)

        self.case_list = []

        def add_case():
            idx = len(self.case_list) + 1
            case = tb.LabelFrame(self.gongkuang_frame, text=f"工况{idx}", padding=10)
            case.grid(row=idx + 1, column=0, sticky="we", pady=5)

            # 动态读取车辆与车道名称
            vehicle_names = []
            for j, v in enumerate(self.vehicle_list):
                name = getattr(v["content"], "name_entry", None)
                vehicle_names.append(name.get() if name else f"车辆{j+1}")

            lane_names = []
            for j, l in enumerate(self.lane_list):
                name = getattr(l["name_entry"], "get", lambda: f"车道{j+1}")()
                lane_names.append(name)

            tb.Label(case, text="荷载工况名称:",width = self.label_width).grid(row=0, column=0, sticky="e", padx=3, pady=3)
            name_entry = tb.Entry(case, width=30)
            name_entry.insert(0, f"工况{idx}")
            name_entry.grid(row=0, column=1, sticky="w", padx=3, pady=3)

            tb.Label(case, text="车辆组:",width = self.label_width).grid(row=1, column=0, sticky="e", padx=3, pady=3)
            vehicle_combo = tb.Combobox(case, values=vehicle_names or ["无车辆定义"], width=25, state="readonly")
            if vehicle_names:
                vehicle_combo.set(vehicle_names[0])
            vehicle_combo.grid(row=1, column=1, sticky="w", padx=3, pady=3)

            tb.Label(case, text="系数:",width = self.label_width).grid(row=2, column=0, sticky="e", padx=3, pady=3)
            factor_entry = tb.Entry(case, width=10, justify="center")
            factor_entry.insert(0, "1.000")
            factor_entry.grid(row=2, column=1, sticky="w", padx=3, pady=3)

            tb.Label(case, text="分配车道:",width = self.label_width).grid(row=3, column=0, sticky="ne", padx=3, pady=3)
            assign_frame = tb.Frame(case)
            assign_frame.grid(row=3, column=1, sticky="w", padx=3, pady=3)

            assign_vars = []
            if lane_names:
                for lane in lane_names:
                    var = tk.BooleanVar(value=False)
                    tb.Checkbutton(assign_frame, text=lane, variable=var, bootstyle="info-round-toggle").pack(anchor="w", pady=2)
                    assign_vars.append((lane, var))
            else:
                tb.Label(assign_frame, text="（尚未定义车道）", foreground="gray").pack(anchor="w")

            del_case_btn = tb.Button(case, text="删除工况", bootstyle="danger-outline", width=8,
                                    command=lambda: delete_case(case))
            del_case_btn.grid(row=0, column=5, padx=10, sticky="e")

            self.case_list.append({
                "frame": case,
                "name": name_entry,
                "vehicle": vehicle_combo,
                "factor": factor_entry,
                "assign": assign_vars
            })
            self.inner.update_idletasks()
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        def delete_case(frame):
            frame.destroy()
            for i in self.case_list:
                if i["frame"] == frame:
                    self.case_list.remove(i)
                    break
            self.inner.update_idletasks()
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        add_case_btn.config(command=add_case)


    # ============ 工具函数 ============
    def _toggle_wind_visibility(self):
        checked = self.wind_var.get()
        if checked:
            self.jixian_wind_spin.set("")                
            self.jixian_wind_spin.config(state="disabled")
            self.jixian_wind_label2.grid(row=1, column=2, sticky="e", padx=self.padx, pady=(0, self.pady))
            self.jixian_wind_entry2.grid(row=1, column=3, sticky="w", padx=self.padx, pady=(0, self.pady))
        else:
            self.jixian_wind_label2.grid_remove()
            self.jixian_wind_entry2.grid_remove()
            self.jixian_wind_spin.set(10)
            self.jixian_wind_spin.config(state="readonly")

    def add_to_selected(self):
        selected_items = [self.move_listbox.get(i) for i in self.move_listbox.curselection()]
        # 避免重复添加
        for item in selected_items:
            if item not in self.selected_loads:
                self.selected_loads.append(item)
        self.update_selected()

    def update_selected(self):
        self.selected_output.config(text=", ".join(self.selected_loads) if self.selected_loads else "无")
    
    
    
    


# ------------- 主界面类不变，仅 ComponentsTab 内部边框固定 -------------
class TrestleUI_v2:
    """主界面"""

    def __init__(self, root):
        self.root = root
        self.root.title("栈桥自动化绘图程序")
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

        self.tab_components.on_support_changed = self.tab_section.set_brake_enabled
        # 初始同步默认不设置制动墩
        self.tab_section.set_brake_enabled(False)
    
        # self.tab_loads = LoadTab(notebook)
        # notebook.add(self.tab_loads, text="荷载参数")

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

        default_mct_name = "Trestle_Untitled.mct"
        default_mct_path = os.path.join(install_path, "Custom", default_mct_name)

        # -------- 保存路径 StringVar --------
        self.save_path_var = tk.StringVar(value=default_mct_path)

        # -------- 路径选择行（新的 UI）--------
        # dialog_savepath(
        #     parent=bottom_frame,
        #     label_txt="MCT 文件保存路径：",
        #     entry_var=self.save_path_var,
        #     button_txt="更改路径",
        #     button_command=self.choose_save_path
        # )
        btn_frame = tb.Frame(bottom_frame)
        btn_frame.pack(fill="x", pady=3)
        self.run_cad_btn = tb.Button(btn_frame, text="绘制CAD", bootstyle=SUCCESS)
        self.run_cad_btn.pack(side="right", padx=10)

        tb.Button(btn_frame, text="退出", bootstyle=DANGER, command=self.root.destroy).pack(side="right", padx=10)


def flatten_to_str_lines(nested):
    """递归展开任意层嵌套列表，并转为字符串"""
    lines = []
    for item in nested:
        if isinstance(item, (list, tuple)):
            lines.extend(flatten_to_str_lines(item))
        else:
            lines.append(str(item))
    return lines


def run_cad(app, acadapp):
    print("CAD 绘图")
    try:
        All(app, acadapp)
    except Exception as e:
        messagebox.showerror("错误", f"CAD 绘图失败：{e}")
        traceback.print_exc()
        

def Trestle_Drawing_AutoCad_Main(parent=None, on_close=None):
    if not if_Reg(parent):
        if on_close:
            on_close()
        return
    # if_AutoCAD_True_False, if_AutoCAD_error_massage = if_AutoCAD()
    FileName = 'ShuZhiQiaoShi'
    Applocation = read_Register('Software\\{}'.format(FileName), 'Applocation')
    print("正在搜索已安装的AutoCAD版本...")
    VersionR_lst = find_installed_autocad_versions()
    versionR_lst = sorted(VersionR_lst, key = lambda x : x.replace('R', ''), reverse = True)
    print(versionR_lst)

    try:
        for Rx in versionR_lst:
            try:
                ProgID = AutoCAD_VersionR_tans_to_COM_ProgID(Rx)
                print(ProgID)
                acadapp = win32com.client.Dispatch(ProgID)
                acadapp.Visible = True
                print(f'AutoCAD\'{Rx}\'连接成功')
                break
            except:
                pass
        Import_Blocks(Applocation, acadapp)
        if parent == None:
            root = tb.Window(themename="cosmo")
        else:
            root = tb.Toplevel(parent)
        app = TrestleUI_v2(root)
        app.run_cad_btn.config(command= lambda:run_cad(app, acadapp))
        if parent == None:
            root.mainloop()
        elif on_close:
            def _on_root_destroy(event, _root=root):
                if event.widget is _root:
                    on_close()
            root.bind('<Destroy>', _on_root_destroy)
    except:
        print('AutoCAD版本号检索失败, 尝试直接启动连接')
        try:
            acadapp = win32com.client.Dispatch('AutoCAD.Application')
            acadapp.Visible = True
            print('AutoCAD连接成功')
            Applocation = read_Register('Software\\ShuZhiQiaoShi', 'Applocation') # 软件安装路径
            Import_Blocks(Applocation, acadapp)
            if parent == None:
                root = tb.Window(themename="cosmo")
            else:
                root = tb.Toplevel(parent)
            app = TrestleUI_v2(root)
            app.run_cad_btn.config(command= lambda:run_cad(app, acadapp))
            if parent == None:
                root.mainloop()
            elif on_close:
                def _on_root_destroy2(event, _root=root):
                    if event.widget is _root:
                        on_close()
                root.bind('<Destroy>', _on_root_destroy2)
        except:
            print('AutoCAD连接失败')
            print('请手动启动AutoCAD')
            if on_close:
                on_close()

# Trestle_Drawing_AutoCad_Main()