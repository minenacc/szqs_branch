# 1. 标准库
import random
import math

# 2. 第三方库
import win32com.client

# 3. 本地模块
from General.ExcelHandle import read_excel_to_dict
from General.FilePath import qucik_save_file_dialog


# ============================================================
# Block 格式解析（从 FEA 移植）
# ============================================================

def parse_blocks_from_lines(lines):
    """
    解析 *BlockName 格式的文本文件
    参数: lines: list of str，文本文件的每一行（末尾换行符已去除）
    返回: dict: { block_name: [list_of_data_lines] }
    """
    blocks = {}
    current_key = None
    current_data = []
    for line in lines:
        line = line.strip()
        # 空行作为块结束分隔符
        if line == '':
            if current_key is not None:
                blocks[current_key] = current_data
                current_key = None
                current_data = []
            continue
        # 标题行：以 '*' 开头
        if line.startswith('*'):
            if current_key is not None:
                blocks[current_key] = current_data
            current_key = line.lstrip('*').strip()
            current_data = []
        else:
            if current_key is not None:
                current_data.append(line)
    # 处理最后可能没有结尾空行的情况
    if current_key is not None:
        blocks[current_key] = current_data
    return blocks


# ============================================================
# Excel 数据读取
# ============================================================

def get_excel_data(Steel_Sheet_Pile_CofferDam_Load_dict, excel_path):
    """
    从 Excel 文件读取土层数据，存入 Load_dict['Excel']
    """
    exApp = win32com.client.Dispatch("Excel.Application")
    exApp.Visible = False
    exApp.DisplayAlerts = False
    excel_data = read_excel_to_dict(exApp, excel_path)
    exApp.Quit()
    excel_dict = {}
    layer_data = excel_data[1:]  # 去掉第一行（表头）
    for i in range(len(layer_data)):
        excel_dict[f'第{i+1}层土'] = {
            '土层名称': layer_data[i][0],
            '土层性质': layer_data[i][1],
            '层厚': layer_data[i][2],
            '重度': layer_data[i][3],
            '黏聚力': layer_data[i][4],
            '内摩擦角': layer_data[i][5],
        }
    Steel_Sheet_Pile_CofferDam_Load_dict['Excel'] = excel_dict
    print('Excel 土层数据已读取')


# ============================================================
# 几何计算工具
# ============================================================

def coordinate_plus_minus_by_angle(angle, param=1):
    """
    根据角度返回坐标缩放系数. param=-1 时反向
    """
    if angle == 0 or angle == 360:
        x, y = [1 * param, 0]
    elif angle == 90:
        x, y = [0, 1 * param]
    elif angle == 180:
        x, y = [-1 * param, 0]
    elif angle == 270:
        x, y = [0, -1 * param]
    elif 0 < angle < 90:
        x, y = [1 * param, 1 * param]
    elif 90 < angle < 180:
        x, y = [-1 * param, 1 * param]
    elif 180 < angle < 270:
        x, y = [-1 * param, -1 * param]
    elif 270 < angle < 360:
        x, y = [1 * param, -1 * param]
    return x, y


def new_length_divisible(length, param1, param2):
    """
    调整 length 使其能被 param1 整除，且商满足 param2 的奇偶约束
    param2=2: 商为偶数; param2=1: 商为奇数
    """
    divisor = length // param1
    if param2 == 2:
        if divisor % 2 != 0:
            divisor += 1
    if param2 == 1:
        if divisor % 2 == 0:
            divisor += 1
    new_length = param1 * divisor
    return new_length


def new_length_divisible2(length, param1):
    """
    将 length 向上取整到 param1 的倍数（无奇偶约束）
    用于锁扣钢管桩间距计算
    """
    remainder = length % param1
    if remainder != 0:
        length += param1 - remainder
    return length


def get_equidistant_points(pointA, pointB, spacing):
    """
    计算线段 AB 上间距为 spacing 的等分点坐标（含端点）
    """
    x1, y1 = pointA[0:2]
    x2, y2 = pointB[0:2]
    length = math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)
    n = int(round(length / spacing))
    num_points = n + 1
    dx = (x2 - x1) / length
    dy = (y2 - y1) / length
    points = []
    for i in range(num_points):
        distance = i * spacing
        x = x1 + distance * dx
        y = y1 + distance * dy
        points.append((x, y, 0))
    return points


# ============================================================
# 截面文件解析
# ============================================================

def Waler_sections(Waler_section_filepath):
    """
    读取围檩截面 .sec 文件，返回截面信息字典
    """
    Waler_section_dict = {}
    with open(Waler_section_filepath, 'r', encoding='utf-8') as txtfile:
        for line in txtfile:
            line_lst = [x for x in line.strip().split()]
            if line_lst[0][0] == '2':  # 双拼
                mctstring = ", ".join([
                    'DBUSER', line_lst[0],
                    'CC, 0, 0, 0, 0, 0, 0, YES, NO, B, 2',
                    line_lst[1], line_lst[2], line_lst[3],
                    line_lst[4], line_lst[5], line_lst[6],
                    '0, 0, 0, 0, NO, 0, NO, 0, 0'
                ])
                section_value = line_lst[1:6]
                Waler_section_dict[line_lst[0]] = {
                    'SEC': section_value, 'MCT': mctstring,
                    'Ix': line_lst[7], 'Wx': line_lst[8], 'ix': line_lst[9],
                }
            else:  # 单工钢/单H型钢
                mctstring = ", ".join([
                    'DBUSER', line_lst[0],
                    'CC, 0, 0, 0, 0, 0, 0, YES, NO, H, 2',
                    line_lst[1], line_lst[2], line_lst[3], line_lst[4],
                    '0, 0', line_lst[5], '0, 0, 0'
                ])
                section_value = line_lst[1:5]
                Waler_section_dict[line_lst[0]] = {
                    'SEC': section_value, 'MCT': mctstring,
                    'Ix': line_lst[6], 'Wx': line_lst[7], 'ix': line_lst[8],
                }
    return Waler_section_dict


def Strut_sections(Strut_section_filepath):
    """
    读取内支撑截面 .sec 文件，返回截面信息字典
    """
    Strut_section_dict = {}
    with open(Strut_section_filepath, 'r', encoding='utf-8') as txtfile:
        for line in txtfile:
            line_lst = [x for x in line.strip().split()]
            if line_lst[0][0] == '2':  # 双拼
                mctstring = ", ".join([
                    'DBUSER', line_lst[0],
                    'CC, 0, 0, 0, 0, 0, 0, YES, NO, B, 2',
                    line_lst[1], line_lst[2], line_lst[3],
                    line_lst[4], line_lst[5], line_lst[6],
                    '0, 0, 0, 0, NO, 0, NO, 0, 0'
                ])
                section_value = line_lst[1:6]
                Strut_section_dict[line_lst[0]] = {
                    'SEC': section_value, 'MCT': mctstring,
                    'Ix': line_lst[7], 'Wx': line_lst[8], 'ix': line_lst[9],
                }
            else:
                if '钢管桩' in line_lst[0]:  # 钢管桩
                    mctstring = ", ".join([
                        'DBUSER', line_lst[0],
                        'CC, 0, 0, 0, 0, 0, 0, YES, NO, P, 2',
                        line_lst[1], line_lst[2],
                        '0, 0, 0, 0, 0, 0, 0, 0'
                    ])
                    section_value = line_lst[1:3]
                    Strut_section_dict[line_lst[0]] = {
                        'SEC': section_value, 'MCT': mctstring,
                        'Ix': line_lst[3], 'Wx': line_lst[4], 'ix': line_lst[5],
                    }
                else:  # 单工钢/单H型钢
                    mctstring = ", ".join([
                        'DBUSER', line_lst[0],
                        'CC, 0, 0, 0, 0, 0, 0, YES, NO, H, 2',
                        line_lst[1], line_lst[2], line_lst[3], line_lst[4],
                        '0, 0', line_lst[5], '0, 0, 0'
                    ])
                    section_value = line_lst[1:5]
                    Strut_section_dict[line_lst[0]] = {
                        'SEC': section_value, 'MCT': mctstring,
                        'Ix': line_lst[6], 'Wx': line_lst[7], 'ix': line_lst[8],
                    }
    return Strut_section_dict


# ============================================================
# 其他工具函数
# ============================================================

def list_random_choice(lst, ki, exclude=None):
    """
    从列表中随机选取 ki 个元素，可排除指定元素
    """
    alst = lst.copy()
    if exclude is not None:
        try:
            alst.remove(exclude)
        except ValueError:
            print('指定元素未找到')
    choice = random.choices(alst, k=ki)
    return choice


def int_or_float(Dividend, Divisor, Accuracy):
    """
    整除返回 int，否则返回保留 Accuracy 位小数的 float
    """
    if Dividend % Divisor != 0:
        return round(Dividend / Divisor, Accuracy)
    else:
        return int(Dividend / Divisor)


def calculate_full_angle(O_point, OX_point, target_point):
    """
    计算 OX 向量到 target 向量的角度（0~360度）
    """
    dx1 = OX_point[0] - O_point[0]
    dy1 = OX_point[1] - O_point[1]
    dx2 = target_point[0] - O_point[0]
    dy2 = target_point[1] - O_point[1]
    angle1 = math.atan2(dy1, dx1)
    angle2 = math.atan2(dy2, dx2)
    angle = angle2 - angle1
    if angle < 0:
        angle += 2 * math.pi
    return round(math.degrees(angle), 2)


# ============================================================
# 每米重量表（用于材料表计算）
# ============================================================
SEC_Weight_per_meter = {
    # 围檩截面: (截面面积 mm², 每米重量 kg/m)
    'HW400X400': (21869, 171.7),
    'HM588X300': (18721, 147.0),
    'HN700X300': (23154, 181.7),
    'HN900X300': (30582, 240.1),
    '2I20a': (7120, 55.8),
    '2I22a': (8420, 66.0),
    '2I25a': (9700, 76.2),
    '2I28a': (11080, 87.0),
    '2I32a': (13420, 105.4),
    '2I36a': (15280, 120.0),
    '2I40a': (17220, 135.2),
    '2I45a': (20400, 160.2),
    '2I50a': (23800, 186.8),
    '2I56a': (27000, 212.0),
    '2I63a': (31000, 243.4),
    '2HW400X400': (43738, 343.4),
    '2HM588X300': (37442, 294.0),
    '2HN600X200': (26342, 206.8),
    '2HN700X300': (46308, 363.4),
    '2HN800X300': (52700, 413.6),
    '2HN900X300': (61164, 480.2),
    '2HN1000X300': (79020, 620.2),
    # 内支撑截面
    '1000X12钢管桩': (37247, 292.4),
    '1000X10钢管桩': (31102, 244.1),
    '820X10钢管桩': (25447, 199.8),
    '630X8钢管桩': (15632, 122.7),
    '600X8钢管桩': (14879, 116.8),
    '426X6钢管桩': (7917, 62.1),
    '377X6钢管桩': (6984, 54.8),
    '325X6钢管桩': (5984, 47.0),
    '273X6钢管桩': (5004, 39.3),
    '219X6钢管桩': (3984, 31.3),
    # 拉森钢板桩
    '拉森Ⅲ': (0, 76.1),   # 每延米重量
    '拉森Ⅳ': (0, 76.1),
    '拉森Ⅵ': (0, 106.2),
}


# if __name__ == "__main__":
#     # 测试 parse_blocks_from_lines
#     test_lines = [
#         '*Param1',
#         '3.406 1.107 10.7 8.7 -1.319 1.25 1.25 2.5',
#         '',
#         '*Param2',
#         '3.906 12.0 拉森Ⅳ 820.0 14.0 1490.0',
#         '',
#         '*Waler',
#         '-1.0 2HM588X300 630X8钢管桩 / 2800 1800 / /',
#         '-2.5 2HM588X300 630X8钢管桩 / 2800 1800 / /',
#     ]
#     blocks = parse_blocks_from_lines(test_lines)
#     for name, data in blocks.items():
#         print(f'*{name}: {data}')

#     # 测试 new_length_divisible2
#     print(f'\nnew_length_divisible2(12370, 1490) = {new_length_divisible2(12370, 1490)}')
#     print(f'new_length_divisible(12300, 400, 1) = {new_length_divisible(12300, 400, 1)}')
