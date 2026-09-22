# 1. 标准库
import shutil
from pathlib import Path

# 3. 本地模块
from General.FilePath import qucik_open_file_dialog


# 读取文件返回值
def read_txt_file(path):
    lst = []
    with open(path, 'r', encoding='utf-8') as file:
        for line in file:
            str_line = line.strip()
            str_part = str_line.split(", ")
            lst.append(str_part)
    return lst


# 将值写入文件更新
def update_txt_file(path, content, join_str = ', '):
    with open(path, 'w', encoding='utf-8') as f:
        for item in content:
            line = join_str.join([str(x) for x in item])
            f.write(line + '\n')


# 获取txt文本内的指定内容
def read_txt_found_line(file_path, line_lookfor_lst):
    # 初始化
    found_txt = ""
    result = []
    # 读取指定文件
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()
    except:
        print("读取参数文件默认路径失败, 文件可能已经被移动, 请重新指定文件位置")
        file_path = qucik_open_file_dialog()
        with open(file_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()
    # 找到指定行的下一行
    for str_lookfor in line_lookfor_lst:
        for i, line in enumerate(lines):
            if line.startswith(str_lookfor):
                found_txt = lines[i+1]
        # 将结果转变为列表
        txt_to_lst = [string.strip() for string in found_txt.split(',')]
        result.append(txt_to_lst)
    return result


# 读取文件返回值
def read_file(path):
    lst = []
    with open(path, 'r', encoding='utf-8') as file:
        for line in file:
            str_line = line.strip()
            str_part = str_line.split(", ")
            lst.append(str_part)
    return lst


# 将值写入文件更新
def update_file(path, content, join_str = ', '):
    with open(path, 'w', encoding='utf-8') as f:
        for item in content:
            line = join_str.join([str(x) for x in item])
            f.write(line + '\n')


def copy_txt_files(src_files: list[str], dst_files: list[str], overwrite: bool = False) -> list[str]:
    """
    批量复制文件到指定路径（支持重命名）。

    Args:
        src_files: 源文件完整路径列表
        dst_files: 目标文件完整路径列表（与 src_files 一一对应）
        overwrite: 是否覆盖已存在的同名文件，默认 False

    Returns:
        成功复制的文件路径列表

    Raises:
        ValueError: 两个列表长度不一致时抛出
    """
    if len(src_files) != len(dst_files):
        raise ValueError(f"列表长度不一致: src_files({len(src_files)}) != dst_files({len(dst_files)})")

    all_copied = []

    for src_file, dst_file in zip(src_files, dst_files):
        src = Path(src_file)
        dst = Path(dst_file)

        # print(f"\n{'='*50}")
        # print(f"源文件:   {src.name}")
        # print(f"目标文件: {dst.name}")

        if not src.is_file():
            print(f"⚠️源文件不存在，跳过: {src_file}")
            continue

        # 自动创建目标目录
        dst.parent.mkdir(parents=True, exist_ok=True)

        if dst.exists() and not overwrite:
            print(f"跳过（已存在）: {dst.name}")
            continue

        shutil.copy2(src, dst)
        all_copied.append(str(dst))
        print(f"已复制: {src.name} → {dst.name}")

    # print(f"\n{'='*50}")
    # print(f"全部完成，共复制 {len(all_copied)} 个文件")
    return all_copied

# ===========================================================================
# 以下为新增函数/类 260814
# ===========================================================================

# 科学计数法
def format_scientific_unicode(number, decimals=1):
    # 计算科学计数法的系数和指数
    coeff, exponent = f"{number:.{decimals}e}".split("e")
    exponent = int(exponent)
    coeff = coeff.rstrip(".0")
    
    # Unicode 上标数字映射
    superscript_digits = str.maketrans("0123456789-", "⁰¹²³⁴⁵⁶⁷⁸⁹⁻")
    exponent_str = f"10{str(exponent).translate(superscript_digits)}"
    
    return f"{coeff}×{exponent_str}"


def format_scientific_latex(number, decimals=1):
    """科学计数法 → LaTeX 格式，如 '2.1\\times{10}^{5}'"""
    coeff, exponent = f"{number:.{decimals}e}".split("e")
    exponent = int(exponent)
    return f"{coeff}\\times{{10}}^{{{exponent}}}"