# 1. 标准库
import re
import base64
import winreg
import datetime

# 2. 第三方库
import wmi
import rsa
import win32clipboard

# 捕获当前时间
def time_str():
    current_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    return current_time


# 获取当前时间
def get_current_time():
    """返回当前时间的 ISO 格式字符串（如 '2023-10-01T12:00:00'）"""
    current_time = datetime.datetime.now().isoformat()
    date = datetime.datetime.fromisoformat(current_time)
    date_part = date.strftime("%Y-%m-%d")
    return date_part


# 判断两个日期相减的天数
def Days_apart(day1, day2):
    date1 = datetime.datetime.strptime(day1, "%Y-%m-%d").date()
    date2 = datetime.datetime.strptime(day2, "%Y-%m-%d").date()
    Days = (date2 - date1).days
    return Days


class Hardware:
    @staticmethod
    def get_cpu_sn():
        """
        获取CPU序列号
        :return: CPU序列号
        """
        c = wmi.WMI()
        for cpu in c.Win32_Processor():
            return cpu.ProcessorId.strip()

    @staticmethod
    def get_baseboard_sn():
        """
        获取主板序列号
        :return: 主板序列号
        """
        c = wmi.WMI()
        for board_id in c.Win32_BaseBoard():
            return board_id.SerialNumber

    @staticmethod
    def get_bios_sn():
        """
        获取BIOS序列号
        :return: BIOS序列号
        """
        c = wmi.WMI()
        for bios_id in c.Win32_BIOS():
            return bios_id.SerialNumber.strip()

    @staticmethod
    def get_disk_sn():
        """
        获取硬盘序列号
        :return: 硬盘序列号列表
        """
        c = wmi.WMI()

        disk_sn_list = []
        for physical_disk in c.Win32_DiskDrive():
            disk_sn_list.append(physical_disk.SerialNumber.replace(" ", ""))
        return disk_sn_list
    

def get_unique_id():
    # 组合信息生成唯一ID
    cpu_sn = Hardware.get_cpu_sn()
    register_str = f"{cpu_sn}"
    return register_str


# 将信息写入注册表
def write_Register(path, key, value):
    # 打开或创建注册表项
    registry_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, path)
    # 设置键值
    winreg.SetValueEx(registry_key, key, 0, winreg.REG_SZ, value)
    # 关闭注册表项
    winreg.CloseKey(registry_key)


# 读取注册表中的信息
def read_Register(path, key_handle):
    try:
        # 打开或创建注册表项
        registry_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, path)
        # 设置键值
        value, value_name = winreg.QueryValueEx(registry_key, key_handle)
        # 关闭注册表项
        winreg.CloseKey(registry_key)
        # 返回值
        return value
    except FileNotFoundError:
        return "NotFound"


# 删除主键的值
def delete_registry_value(path, key_handle):
    try:
        # 打开或创建注册表项
        registry_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, path)
        # 删除键值
        winreg.DeleteValue(registry_key, key_handle)
        # 关闭注册表项
        winreg.CloseKey(registry_key)
        return True
    except FileNotFoundError:
        return "NotFound"


# 私钥还原注册码
def Recover_code(registration_code, privkey):
    # 解密注册码
    decrypted_info = rsa.decrypt(base64.b64decode(registration_code), privkey).decode()
    return decrypted_info


# 是否过期
def if_Expired(valid_days = 180):
    Reg_path = "Software\\ShuZhiQiaoShi"
    Reg_time_key = "Registration_Time"
    error_massage = ''
    try:
        Reg_time_value = read_Register(Reg_path, Reg_time_key) # 其实直接调用 Reg_time_value 变量也是一样的
        # 天数差
        current_time = get_current_time()
        passed_days = Days_apart(Reg_time_value, current_time)
        # print(passed_days)
    except:
        error_massage = "错误, 未在注册表中读取到注册日期"
        print(error_massage)
        return [False, error_massage]
    # 有效期180天
    if passed_days > valid_days:
        error_massage = "软件已过期"
        print("error_massage")
        return [False, error_massage]
    else:
        remaining_days = valid_days - passed_days
        print(f"软件剩余使用天数: {remaining_days}")
        return [True, error_massage]


# 获取revit注册表的安装地址
def get_revit_install_path(Revit_Version):
    """获取 Revit 的安装路径"""
    try:
        # 打开 Revit 的注册表项
        key_path = f"SOFTWARE\\Autodesk\\Revit\\{Revit_Version}\\REVIT-05:0804"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
            # 读取安装路径
            install_path, _ = winreg.QueryValueEx(key, "InstallationLocation")
            return install_path
    except FileNotFoundError:
        # 当前版本未找到，尝试下一个版本
        print('版本未找到')
    except Exception as e:
        print(f"查询 Revit {Revit_Version} 时出错: {str(e)}")
    raise FileNotFoundError("未找到 Revit 安装路径。请确保 Revit 已安装")


# 字符串剪切板
def SET_CLIP_STRING(string):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardText(string)
    win32clipboard.CloseClipboard()


def format_value(value, boundary = 10000):
    """格式化数值，如果大于 boundary 则使用科学计数法"""
    try:
        num = float(value)
        if abs(num) > boundary:
            return "{:.4e}".format(num)
        else:
            return str(num)
    except (ValueError, TypeError):
        return str(value) if value is not None else ""
    

def StringRegExpS(pat, string, key):
    matches = []
    pattern = re.compile(pat)
    if key:
        flags = re.I
    else:
        flags = 0
    result = re.findall(pattern, string, flags)
    for match in result:
        matches.append(match)
    return matches

#正则表达式, 字符串转化为间距
# def midasdisttolst(txt):
#     temp1 = StringRegExpS("[^,，+ ]+", txt, "")
#     temp5 = []
#     for x in temp1:
#         temp2 = StringRegExpS("[^×@*]+", x, "")
#         if len(temp2) == 1:
#             temp2 = ["1"] + temp2
#         temp3 = int(temp2[0])
#         temp4 = int(temp2[-1])
#         temp5 += [temp4] * temp3
#     return temp5


#正则表达式, 指定起点后字符串转化为所有x坐标（间距累加）
def midasdisttolst2(basecoord, txt):
    a = basecoord
    temp = []
    for x in midasdisttolst(txt):
        sum = a + x
        a = sum
        temp.append(sum)
    return [basecoord] + temp


# 线性内插，x1默认在左，x2默认在右
def inear_interpolation(x1, x2, h1, h2, p):
    # 斜率
    k = (h2 - h1) / (x2 -x1)
    # 高度
    q = h1 + k * (p - x1)
    # 返回值
    return q


#  单元荷载去除负值
def elem_load_lst_trans(position_lst, load_lst, param = 4):
    '''
    position_lst: 荷载的相对位置表[0, 0.556, 1]
    load_lst: 荷载的值表[-3.6, 0, 4.5]
    返回: [0, 0.556, 1, 0], [0, 0, 4.5, 0]
    '''
    new_position_lst = []
    new_load_lst = []
    for i in range(param):
        try:
            new_position_lst.append(position_lst[i])
            new_load_lst.append(max(load_lst[i], 0))
        except:
            new_position_lst.append(0)
            new_load_lst.append(0)
    return new_position_lst, new_load_lst


def add_continuous_beam_elem_loads(load1, load2):
    """
    叠加两个分段线性荷载。
    规则：
      - 以第一个荷载的位置列表为输出位置（保持原顺序）。
      - 找到第一个荷载中最大值（通常为 1）的位置，该位置及之前为有效区间，之后的所有点均视为结束标记，对应值强制为 0。
      - 对第二个荷载进行线性插值（仅在其有效定义域内），并与第一个荷载的有效值相加。
    """
    pos1, val1 = load1
    pos2, val2 = load2

    # 聚合第二个荷载的重复位置（累加相同位置的值）
    d2 = {}
    for p, v in zip(pos2, val2):
        d2[p] = d2.get(p, 0.0) + v
    sorted_pos = sorted(d2.keys())
    sorted_val = [d2[p] for p in sorted_pos]

    # 线性插值函数（外推取最近端点值）
    def interp(x):
        if x <= sorted_pos[0]:
            return sorted_val[0]
        if x >= sorted_pos[-1]:
            return sorted_val[-1]
        for i in range(len(sorted_pos) - 1):
            if sorted_pos[i] <= x <= sorted_pos[i + 1]:
                return sorted_val[i] + (sorted_val[i + 1] - sorted_val[i]) * (x - sorted_pos[i]) / (sorted_pos[i + 1] - sorted_pos[i])
        return 0.0

    # 确定第一个荷载中的最大值（一般为1）及其首次出现位置
    max_pos = max(pos1)
    max_index = pos1.index(max_pos)  # 第一个最大值的位置

    # 构建输出值列表
    new_vals = []
    for i, (p, v1) in enumerate(zip(pos1, val1)):
        if i > max_index:   # 最大值之后的所有点均为结束标记
            new_vals.append(0.0)
        else:
            new_vals.append(v1 + interp(p))
    return (pos1, new_vals)


# 双参数表查询程序,table1是高度或者水平加载长度表，table2是table1在对应地表类别下对应的参数表,param是基准高度,digits是返回值的精度
def double_param_table_index(table1, table2, param, digits):
    # 参数是否是现成的，是的话直接查询结果不需要线性内插
    if param in table1:
        param_index = table1.index(param)
        value = table2[param_index]
    # 参数没在表中找到，开始线性内插
    else:
        # 基准高度<5
        if param < table1[0]:
            value = table2[0]
        # 基准高度>450
        elif param > table1[-1]:
            value = table2[-1]
        else:
            # 基准高度在5和450之间
            # 查询位置和对应的value
            for i, num in enumerate(table1):
                if num > param:
                    left_index = i - 1
                    right_index = i
                    break
            # 得到param在表中左右的高度和对应的value值
            left_param = table1[left_index]
            right_param = table1[right_index]
            left_value = table2[left_index]
            right_value = table2[right_index]
            # 线性内插
            value = round(inear_interpolation(left_param, right_param, left_value, right_value, param), digits)
    return value


#(600 600 400 400 600 600)→(600 600 500 400 500 600 600)小肋承担混凝土荷载宽度表
def ribdistancelist_trans (lst):
    new_lst = []
    for x in range(len(lst)+1):
        if x == 0 :
            new_lst.append(lst[x])
        elif x == len(lst):
            new_lst.append(lst[-1])
        else:
            new_lst.append((lst[x-1]+lst[x])/2)
    return new_lst


#去除表头表尾的一个0或者添加首尾0
def trim0(s,prep):
    #复制源列表,不改变源列表
    t = s.copy()
    if prep == 0:#等于0 去掉首尾0
        if len(t) > 0 and t[0] == 0:
            t = t[1:]
        if len(t) > 0 and t[-1] == 0:
            t = t[:-1]
    if prep == 1:#等于1 添加首尾0
        t.insert(0,0)
        t.append(0)
    return t


#对列表第n项进行判断，如果有则替换，没有则新建
def update_or_append(lst, n, new_value):
    if n < len(lst):
        lst[n] = new_value
    else:
        lst.append(new_value)


# 判断一个点是否在一个点表中有近似点
def Bailey_pt_close_to_ptinlst(pt, lst, angle):
    for x in lst:
        if abs(pt[0] - x[1][0]) < 1 and abs(pt[1] - x[1][1]) < 1 and angle == x[2]:
            return x
    return False


# 移除一个表中的近似的点
def remove_nearby_points(points, threshold=1):
    unique_points = []
    for point in points:
        if not any(all(abs(p1 - p2) < threshold for p1, p2 in zip(point, unique_point)) for unique_point in unique_points):
            unique_points.append(point)
    return unique_points


def pointlist_extend(pointlist, reference_point=[0,0,0]):
    """
    将[(1,2,3),(4,5,6)]转化成CAD可识别的安全数组(1,2,3,4,5,6), 存在定位点时则将相对坐标值转移
    """
    flat_points = []
    for point in pointlist:
        flat_points.extend([coord1+coord2 for coord1, coord2 in zip(point, reference_point)])
    return flat_points


def num_to_chinese_num(num):
    """将正整数转换为中文数字（支持 1~9999）"""
    chinese_digits = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九']
    units = ['', '十', '百', '千']
    if num == 0:
        return '零'
    # 将数字拆分为各位
    digits = list(map(int, str(num)))
    length = len(digits)
    result = []
    for i, d in enumerate(digits):
        if d != 0:
            # 添加数字
            result.append(chinese_digits[d])
            # 添加单位（十、百、千），注意个位没有单位
            if length - i - 1 > 0:
                result.append(units[length - i - 1])
        else:
            # 处理连续的零：只在非零位前加一个“零”
            if i > 0 and digits[i-1] != 0 and i != length-1:
                result.append('零')
    # 处理特殊情况：如 10 应输出“十”而非“一十”
    if num >= 10 and num < 20 and result[0] == '一':
        result.pop(0)  # 去掉开头的“一”
    return ''.join(result)


# 根据土层标高划分土条
def Divide_Solid_within_LayerThickness(Solid_Top_Level, Solid_Bottom_Level, LayerThickness, Param = 0.5):
    '''
    Solid_Top_Level 土层顶标高 -2.0m
    Solid_Bottom_Level 土层底标高 -8.8m
    LayerThickness 土条划分厚度 1m
    Param 余数的判定值
    '''
    Solid_H = abs(Solid_Top_Level-Solid_Bottom_Level)
    remainder = round(Solid_H % LayerThickness,3) # 余数
    Layer = int(Solid_H / LayerThickness) if remainder < Param else int(round(Solid_H / LayerThickness, 0)) # 层数
    Hlst = [round(Solid_Top_Level - i*LayerThickness,3) for i in range(Layer)]
    Hlst.append(Solid_Bottom_Level)
    '''
    Hlst 划分土条后的标高 [-2.0, -3.0, -4.0, -5.0, -6.0, -7.0, -8.0, -8.8]
    '''
    return Hlst


def intersect_segment_with_intervals(intervals, segment):
    '''
    intervals 连续线段上的坐标点 [1, 2, 3, 4, 5, 6, 7, 8, 9]
    segment 线段[2.6, 4.6]
    '''
    # intervals = sorted(intervals)
    intervals_lst = [sorted([intervals[i], intervals[i+1]]) for i in range(len(intervals)-1)]
    # print(intervals_lst)
    segment_pt1, segment_pt2 = sorted(segment)
    # print(segment_pt1, segment_pt2)
    segment_pt1_ilst, segment_pt2_ilst = [], [] # 考虑边界问题,将所有可能值放进表中
    # 定位 point1 point2 所在的区间
    for i, lst in enumerate(intervals_lst):
        if segment_pt1 >= lst[0] and segment_pt1 <= lst[-1]:
            segment_pt1_ilst.append(i+1)
        if segment_pt2 >= lst[0] and segment_pt2 <= lst[-1]:
            segment_pt2_ilst.append(i+1)
    segment_pt1_i = max(segment_pt1_ilst)
    segment_pt2_i = min(segment_pt2_ilst)
    # print(segment_pt1_i, segment_pt2_i)
    # 计算不同区间的长度
    reslst = []
    if segment_pt1_i == segment_pt2_i:
        reslst.append([segment_pt1_i, round(abs(segment[0]-segment[-1]), 3)])
    else:
        for j in list(range(min(segment_pt1_i, segment_pt2_i), max(segment_pt1_i, segment_pt2_i)+1)):
            if j == segment_pt1_i:
                reslst.append([segment_pt1_i, round(abs(segment_pt1 - sorted(intervals_lst[segment_pt1_i-1])[-1]), 3)])
            elif j == segment_pt2_i:
                reslst.append([segment_pt2_i, round(abs(segment_pt2 - sorted(intervals_lst[segment_pt2_i-1])[0]), 3)])
            else:
                reslst.append([j, round(abs(intervals_lst[j-1][0] - intervals_lst[j-1][1]),3)])
    reslst = sorted(reslst, key = lambda x : x[0])
    '''
    reslst 返回表, 子表的元素分别为[区间号(区间号从1开始计), segment 在该区间的线段长度] [[2, 0.4], [3, 1], [4, 0.6]]
    '''
    return reslst


def gamma_cal(Excel_Data, start_depth, end_depth):
    """计算深度区间内的加权平均天然重度，并返回桩底处土层的 c/φ"""
    start_depth, end_depth = min(start_depth, end_depth), max(start_depth, end_depth)
    data = Excel_Data[1:]
    total_moment = 0.0
    calc_length = end_depth - start_depth
    if calc_length == 0:
        return 0.0, 0.0, 0.0
    target_c, target_phi = 0.0, 0.0
    current_layer_top = 0.0
    for row in data:
        # row = ['1-1', 2.922, 17.0, 0.0, 10.0, '水土分算', '0.5', '0.0']
        layer_thick = float(row[1])
        gamma_nat = float(row[2])
        current_layer_bottom = current_layer_top + layer_thick
        if current_layer_top < end_depth <= current_layer_bottom:
            target_c = float(row[3])
            target_phi = float(row[4])
        # 寻找当前土层与计算区间的交集并累加
        intersect_top = max(current_layer_top, start_depth)
        intersect_bottom = min(current_layer_bottom, end_depth)
        if intersect_top < intersect_bottom:
            h = intersect_bottom - intersect_top
            total_moment += h * gamma_nat
        # 更新下一层顶标高
        current_layer_top = current_layer_bottom
        if current_layer_top >= end_depth:
            break
    else:
        last = data[-1]
        target_c, target_phi = float(last[3]), float(last[4])
    return round(total_moment / calc_length, 2), round(target_c, 2), round(target_phi, 2)


def gamma_sub_cal(Excel_Data, start_depth, end_depth, water_level, gamma_w=10):
    """
    计算深度区间 [start_depth, end_depth] 内各土层的加权平均浮重度。
    深度自第一个土层顶面(地面)起算；水位 water_level(自地面起算，可为负值)以上取天然重度，
    以下取浮重度(天然重度 - gamma_w)。仅按实际存在的土层长度加权。
    """
    start_depth, end_depth = min(start_depth, end_depth), max(start_depth, end_depth)
    data = Excel_Data[1:]
    total_moment = 0.0
    total_length = 0.0
    current_layer_top = 0.0
    for row in data:
        layer_thick = float(row[1])   # 层厚
        gamma_nat = float(row[2])     # 重度
        current_layer_bottom = current_layer_top + layer_thick
        # 当前土层与计算区间的交集
        seg_top = max(current_layer_top, start_depth)
        seg_bot = min(current_layer_bottom, end_depth)
        if seg_top < seg_bot:
            # 水位以上部分: 天然重度
            above_bot = min(seg_bot, water_level)
            if seg_top < above_bot:
                total_moment += (above_bot - seg_top) * gamma_nat
                total_length += above_bot - seg_top
            # 水位以下部分: 浮重度 (天然 - gamma_w)
            below_top = max(seg_top, water_level)
            if below_top < seg_bot:
                total_moment += (seg_bot - below_top) * (gamma_nat - gamma_w)
                total_length += seg_bot - below_top
        current_layer_top = current_layer_bottom
        if current_layer_top >= end_depth:
            break
    if total_length <= 0:
        return 0.0
    return round(total_moment / total_length, 2)

# ===========================================================================
# 以下为新增函数/类 260814
# ===========================================================================
def _parse_simple_spacing(s):
    """解析无括号的简单间距格式 → [数值列表]（浮点兼容）"""
    temp1 = re.findall("[^,，+ ]+", s)
    temp5 = []
    for x in temp1:
        temp2 = re.findall("[^×@*]+", x)
        if len(temp2) == 1:
            val = float(temp2[0])
            temp5.append(int(val) if val == int(val) else val)
        else:
            temp3 = int(float(temp2[0]))
            temp4 = float(temp2[-1])
            temp5 += [int(temp4) if temp4 == int(temp4) else temp4] * temp3
    return temp5

# ===========================================================================
# 以下为修改函数/类 260814
# ===========================================================================
def midasdisttolst(txt):
    """解析间距字符串为数值列表（展开括号分组）。

    支持格式：
      "2@3000"         → [3000, 3000]
      "3×3000"         → [3000, 3000, 3000]  （旧格式兼容）
      "1.5+10@15+1.5"  → [1.5, 15, 15, ..., 1.5]
      "0,2@750+50@(795+705+705+795)+2@750" → [0, 750, 750, 795, 705, 705, 795 × 50, 750, 750]
    """
    if not txt:
        return []
    def _expand(s):
        result = s
        while '(' in result:
            m = re.search(r'(\d+)\s*@\s*\(([^)]+)\)', result)
            if m:
                count = int(m.group(1))
                inner_vals = _parse_simple_spacing(m.group(2))
                result = result[:m.start()] + ",".join(str(v) for v in inner_vals * count) + result[m.end():]
            else:
                m2 = re.search(r'\(([^)]+)\)', result)
                if m2:
                    inner_vals = _parse_simple_spacing(m2.group(1))
                    result = result[:m2.start()] + ",".join(str(v) for v in inner_vals) + result[m2.end():]
                else:
                    break
        return result
    return _parse_simple_spacing(_expand(txt))
