# 1. 标准库
import os
import sys
import traceback
from pathlib import Path

# 2. 第三方库
import ttkbootstrap as tb
from tkinter import messagebox

# 3. 本地模块
sys.path.append(str(Path(__file__).parent.parent.parent))
from General.UIHandle import if_Reg
from General.DataUtils import SET_CLIP_STRING
from General.FilePath import file_extension_Modified, write_file_within_path

sys.path.append(str(Path(__file__).parent))
from FEM_MidasCivil.Steel_Pipe_Bailey_Platform_FEA.Platform_FEM_UI         import PlatformUI
from FEM_MidasCivil.Steel_Pipe_Bailey_Platform_FEA.Platform_CSV_format     import CSV_format
from FEM_MidasCivil.Steel_Pipe_Bailey_Platform_FEA.Platform_MCT_format_FEM import build_mct_from_ui, build_TXT_from_mct


# 展开嵌套列表
def flatten_to_str_lines(nested):
    """递归展开任意层嵌套列表，并转为字符串"""
    lines = []
    for item in nested:
        if isinstance(item, (list, tuple)):
            lines.extend(flatten_to_str_lines(item))
        else:
            lines.append(str(item))
    return lines


def run_mct(app):
    print("MCT 生成")
    (
        mct_lst, xl_single_nodes, single_start_nodes, distribute_single_nodes,  pile_group_nodes,  xl_section_cad, xl_section_name,distribute_section_cad, distribute_section_name,gz_section_cad, gz_section_name,section_weight_per_meter_lst,Feng_values,Feng_result_lst, Shui_values, Shui_lst,single_csv_nodes
    ) = build_mct_from_ui(app)
    mct_lst_tolines = "\n".join([item for sublist in mct_lst for item in sublist])
    SET_CLIP_STRING(mct_lst_tolines)
    messagebox.showinfo("成功", f"mct命令流已粘贴到剪切板中")

    save_dir = app.save_path_var.get()
    if not save_dir:
        messagebox.showwarning("提示", "请先选择保存路径！")
        return

    save_path_mct = os.path.join(save_dir)
    with open(save_path_mct, "w", encoding="gbk") as f:
        f.write(mct_lst_tolines)
    print(f"MCT 文件已生成：{save_path_mct}")
    messagebox.showinfo("成功", f"MCT 文件已生成：\n{save_path_mct}")

    print("TXT 生成")
    mct_path = app.save_path_var.get()
    txt_path = file_extension_Modified(mct_path, ".txt")
    try:
        TXT_lst = build_TXT_from_mct(app)
        txt_lst_tolines = "\n".join([item for sublist in TXT_lst for item in sublist])
        write_file_within_path(txt_path, txt_lst_tolines, "utf-8")
        print(f"TXT 文件已生成：{txt_path}")
    except Exception as e:
        messagebox.showerror("错误", f"TXT 格式化失败：{e}")
        traceback.print_exc()

    print("CSV 生成")
    try:
        CSV_format(app)
    except Exception as e:
        messagebox.showerror("错误", f"CSV 格式化失败：{e}")
        traceback.print_exc()


def PlatForm_FEM_MidasCivil_Main(parent=None, on_close=None):
    if not if_Reg(parent):
        if on_close:
            on_close()
        return
    if parent == None:
        root = tb.Window(themename="cosmo")
    else:
        root = tb.Toplevel(parent)
    app = PlatformUI(root)
    app.run_mct_btn.config(command= lambda: run_mct(app))
    if parent == None:
        root.mainloop()
    elif on_close:
        def _on_root_destroy(event, _root=root):
            if event.widget is _root:
                on_close()
        root.bind('<Destroy>', _on_root_destroy)

# PlatForm_FEM_MidasCivil_Main()