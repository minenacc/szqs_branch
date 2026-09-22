# 1. 标准库
import re 
import os
import winreg
import subprocess
from itertools import chain

# 2. 第三方库
import win32gui
import requests


# ---------- 自定义异常：Midas API 连接失败 ----------
class MidasConnectionError(Exception):
    """Midas API 连接超时或断开，多次重试后仍失败时抛出"""
    pass


# 从注册表获取 Url 和 Mapi-key
def MidasURL():
    # 定义全局变量
    try:
        global base_url, reg_key
        reg_path = winreg.OpenKey(winreg.HKEY_CURRENT_USER,r"SOFTWARE\MIDAS\CVLwNX_CH\CONNECTION")
        reg_uri = winreg.QueryValueEx(reg_path, "URI")[0]  # 获取URI
        reg_port = winreg.QueryValueEx(reg_path, "PORT")[0]  # 获取PORT
        reg_key = winreg.QueryValueEx(reg_path, "Key")[0]  # 获取Mapi-key
        base_url = "https://" + reg_uri + ":" + reg_port + "/civil"  # 定义基础URL
        # print("基础URL: ", base_url)
    except:
        print('错误, API未连接, URL获取失败')


# 定义MidasAPI函数，用于调用Midas API
# timeout  : 单次请求超时秒数，默认 300s（5分钟）；None 表示不设超时（兼容旧行为）
# max_retries: 超时/连接失败后的重试次数，默认 0（不重试，直接抛异常）
def MidasAPI(method, command, body=None, timeout=600, max_retries=0):
    # 拼接URL
    url = base_url + command
    # 定义请求头
    headers = {
        "Content-Type": "application/json",
        "MAPI-Key": reg_key
    }

    for attempt in range(max_retries + 1):
        try:
            # 根据请求方法，发送请求
            if method == "POST":
                response = requests.post(url=url, headers=headers, json=body, timeout=timeout)
            elif method == "PUT":
                response = requests.put(url=url, headers=headers, json=body, timeout=timeout)
            elif method == "GET":
                response = requests.get(url=url, headers=headers, timeout=timeout)
            elif method == "DELETE":
                response = requests.delete(url=url, headers=headers, timeout=timeout)
            # 打印请求方法、URL和状态码
            print(method, command, response.status_code)
            # 返回响应的json数据
            return response.json()
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as e:
            if attempt < max_retries:
                print(f"[MidasAPI] 第{attempt+1}次请求失败({type(e).__name__}), 正在重试...")
            else:
                print(f"[MidasAPI] 已重试{max_retries}次仍失败, 抛出 MidasConnectionError")
                raise MidasConnectionError(
                    f"Midas API {method} {command} 失败: {type(e).__name__}: {e}"
                ) from e


# 打开Midas并修改编辑框对应的值 
def Open_Midas_civil(civilnx_path, entry_var):
    # print(entry_var.get())
    # 检查exe文件是否存在,如果存在则启动exe
    if os.path.exists(civilnx_path) and entry_var.get() != "已打开":
        process = subprocess.Popen(['explorer', civilnx_path])
        # 更新对应编辑框的值
        entry_var.set("已打开")
        print("Midas civil NX 正在启动。")
    
    if entry_var.get() == "未安装":
        # print("civilnx未安装或路径不正确。")
        print("Midas civil NX 未成功启动, 或未安装或路径不正确")

    if entry_var.get() == "已打开":
        print("Midas civil NX 已打开")


# 检测当前打开文件的Midas文件
def get_current_midas_file():
    """尝试获取当前打开的 Midas Civil NX 文件路径"""
    detected_path = []
    def enum_windows_callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            # 匹配格式：MIDAS CIVIL NX ... - [路径] - [MIDAS CIVIL NX]
            # 正则解释：找第一个 [ 之后的内容，直到遇到 ] 结束
            match = re.search(r'-\s+\[(.:\\.*?)\]\s+-\s+\[MIDAS CIVIL NX\]', title)
            if match:
                raw_path = match.group(1).strip()
                raw_path = raw_path.split('*')[0].strip()
                if not raw_path.lower().endswith('.mcb'):
                    full_path = raw_path + ".mcb"
                else:
                    full_path = raw_path
                if os.path.exists(full_path):
                    detected_path.append(full_path)
    win32gui.EnumWindows(enum_windows_callback, None)
    return detected_path[0] if detected_path else None


# 获取版本号
def get_midas_version():
    """获取当前 Midas Civil NX 的版本号（如 2024）"""
    version = "Unknown"
    def enum_windows_callback(hwnd, _):
        nonlocal version # 允许在内部函数修改外部变量
        if version != "Unknown":
            return True
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            match = re.search(r'MIDAS CIVIL NX\s+(\d+)\s+-', title, re.I)
            if match:
                version = match.group(1)
        return True

    win32gui.EnumWindows(enum_windows_callback, None)
    return version


# 在Midas已经启动的前提下打开选定的文件
def Open_mcb_within_API(openfile_path):
    if "mcb" in openfile_path:
        arguments = {"Argument" : openfile_path}
        MidasAPI("POST" , "/doc/open" , arguments)
    else:
        print("未识别到有效mcb文件, 请重新选择")


def importmct_to_mcb(mct_path, mcb_path):
    mct = {
        "Argument" : mct_path
    }
    mcb = {
        "Argument" : mcb_path
    }
    MidasAPI("POST", "/DOC/NEW", {}) # 新建
    MidasAPI("POST", "/DOC/IMPORTMXT", mct) # 导入mct
    MidasAPI("POST", "/DOC/SAVEAS", mcb) # 导入mct


# 将单元号列表变为x1 to y1, x2 to y2, ...的形式, 计算简单，适用于连续加1比较多的数组
def num_lst_to_num_group(numbers):
    numbers = sorted(numbers)  # 确保列表有序（如果输入可能无序）
    ranges = []
    start = numbers[0]
    for i in range(1, len(numbers)):
        if numbers[i] != numbers[i-1] + 1:
            if start == numbers[i-1]:
                ranges.append(str(start))
            else:
                ranges.append(f"{start} to {numbers[i-1]}")
            start = numbers[i]
    # 处理最后一个序列
    if start == numbers[-1]:
        ranges.append(str(start))
    else:
        ranges.append(f"{start} to {numbers[-1]}")
    return ", ".join(ranges)


# 将单元号列表变为x1 to y1, x2 to y2, ...的形式, 计算较复杂，适用于存在连续加n的数组
def group_arithmetic_sequences(numbers):
    '''
    [1, 3, 5, 8, 13, 16, 19, 20, 24, 28, 32, 40, 60, 80]
    转换成
    1to5by2 8 13to19by3 20to32by4 40to80by20
    '''
    if not numbers:
        return ""
    result = []
    i = 0
    n = len(numbers)
    while i < n:
        # 查找当前序列的结束位置
        j = i
        # 尝试找到至少3个元素的等差数列
        if j + 2 < n:
            # 计算前两个差值
            diff1 = numbers[j+1] - numbers[j]
            diff2 = numbers[j+2] - numbers[j+1]
            # 如果前两个差值相等，则继续查找更多相同差值的元素
            if diff1 == diff2:
                step = diff1
                k = j + 2
                while k + 1 < n and numbers[k+1] - numbers[k] == step:
                    k += 1
                # 添加找到的序列
                result.append(f"{numbers[i]}to{numbers[k]}by{step}")
                i = k + 1
                continue
        # 如果没有找到合适的序列，添加单个数字
        result.append(str(numbers[i]))
        i += 1
    return " ".join(result)


# 将模型中单位变为KN和M
# unit_force = "KN"   # KN, N, KGF, TONF, LBF, KIPS
# unit_dist = "M"     # M, CM, MM, IN, FT
def API_DIST_FORCE_Unit(unit_force, unit_dist):
    # 调整模型单位
    # Create Unit Bodyy
    unit_dict = {"Assign": {
        "1": {
            "DIST": unit_dist,
            "FORCE": unit_force
        }
    }
    }
    MidasAPI("PUT", "/db/UNIT", unit_dict)


# 图片处理函数——获取模型前处理图片
def Picture_Pre(picture_path , SET_HIDDEN = False, file_name = "模型前处理.jpg", set_mode = 'pre', ANGLE_HORIZONTAL = 15, ANGLE_VERTICAL = 15, BOUNDARY_SUPPORT = False):
    model_file_path = os.path.join(picture_path, file_name)
    FEA_model_dict = {
        "Argument": {
            "SET_MODE": set_mode,
            "EXPORT_PATH": model_file_path,
            "ANGLE": {
                        "HORIZONTAL": ANGLE_HORIZONTAL,
                        "VERTICAL": ANGLE_VERTICAL
            },
            "ACTIVE": {
                "ACTIVE_MODE": "All"
            },
            "SET_HIDDEN": SET_HIDDEN,
            "BOUNDARY": {
                "SUPPORT": BOUNDARY_SUPPORT
            }
        }
    }
    MidasAPI("POST", "/VIEW/CAPTURE", FEA_model_dict)


# picture_path：图片储存路径, str
# picture_name：图片名, str
# active_identity：激活的结构组名称，可以有多个结构组, list
# load_case_comb：所需的荷载组合名，如钢管桩失稳检算时需要“基本组合”作用下的内力图, str
# components_comp：需要查询的内力方向，Fx/My/Mz, str
# 图片处理函数——梁单元内力图
def Picture_BeamForce(picture_path, picture_name, active_identity, load_case_comb, components_comp, HORIZONTAL=45, VERTICAL=30):
    # 图片储存路径
    picture_file_path = os.path.join(picture_path, picture_name)
    Beam_force_dict = {
        "Argument": {
            "SET_MODE": "post",
            "EXPORT_PATH": picture_file_path,
            "ANGLE": {
                        "HORIZONTAL": HORIZONTAL,
                        "VERTICAL": VERTICAL
            },
            "ACTIVE": {
                # "ACTIVE_MODE": "All"
                "ACTIVE_MODE": "Identity",
                "IDENTITY_TYPE": "Group",
                "IDENTITY_LIST": active_identity
            },
            "RESULT_GRAPHIC": {
                "CURRENT_MODE": "Beam Diagrams",
                "LOAD_CASE_COMB": {
                    "TYPE": "CB",
                    "NAME": load_case_comb,
                    "STEP_INDEX": 1
                },
                "COMPONENTS": {
                    "PART": "Total",
                    "COMP": components_comp,
                    # "COMP_SUB": "Maximum"
                },
                "DISPLAY_OPTIONS": {
                    "FIDELITY": "5 Points",
                    "FILL": "line fill",
                    "SCALE": 1.0
                },
                "TYPE_OF_DISPLAY": {
                    "CONTOUR": {
                        "OPT_CHECK": True
                    },
                    "VALUES": {
                        "OPT_CHECK": True,
                        "DECIMAL_PT": 1,
                        "MINMAX_ONLY": {
                            "MAXMIN": "absmax",
                            "LIMIT_SCALE": 0
                        }
                    },
                    "LEGEND": {
                        "OPT_CHECK": True,
                        "VALUE_EXP": False,
                        "DECIMAL_PT": 1
                    },
                    "DEFORM": {
                        "OPT_CHECK": False
                    }
                },
                "OUTPUT_SECT_LOCATION": {
                    "OPT_MAX_MINMAX_ALL": "Min/Max"
                }
            }
        }
    }
    # API调用
    MidasAPI("POST", "/VIEW/CAPTURE", Beam_force_dict)


# picture_path：图片储存路径, str
# picture_name：图片名, str
# active_identity：激活的结构组名称，可以有多个结构组, list
# load_case_comb：所需的荷载组合名, str
# components_comp：需要查询的应力方向，Combined/Ssz, str
# 图片处理函数——梁单元应力图
def Picture_Beamstress(picture_path, picture_name, active_identity, load_case_comb, components_comp, HORIZONTAL=45, VERTICAL=30):
    # 图片储存路径
    picture_file_path = os.path.join(picture_path, picture_name)
    Beam_Stresses_dict = {
        "Argument": {
            "SET_MODE": "post",
            "EXPORT_PATH": picture_file_path,
            "ANGLE": {
                        "HORIZONTAL": HORIZONTAL,
                        "VERTICAL": VERTICAL
            },
            "ACTIVE": {
                # "ACTIVE_MODE": "All"
                "ACTIVE_MODE": "Identity",
                "IDENTITY_TYPE": "Group",
                "IDENTITY_LIST": active_identity
            },
            "RESULT_GRAPHIC": {
                "CURRENT_MODE": "Beam Stresses Diagram",
                "LOAD_CASE_COMB": {
                    "TYPE": "CB",
                    "NAME": load_case_comb,
                    "STEP_INDEX": 1
                },
                "COMPONENTS": {
                    "PART": "Total",
                    "COMP": components_comp,
                    "COMP_SUB": "Maximum"
                },
                "DISPLAY_OPTIONS": {
                    "FIDELITY": "5 Points",
                    "FILL": "line fill",
                    "SCALE": 1.0
                },
                "TYPE_OF_DISPLAY": {
                    "CONTOUR": {
                        "OPT_CHECK": True
                    },
                    "VALUES": {
                        "OPT_CHECK": True,
                        "DECIMAL_PT": 1,
                        "MINMAX_ONLY": {
                            "MAXMIN": "absmax",
                            "LIMIT_SCALE": 0
                        }
                    },
                    "LEGEND": {
                        "OPT_CHECK": True,
                        "VALUE_EXP": False,
                        "DECIMAL_PT": 1
                    },
                    "DEFORM": {
                        "OPT_CHECK": False
                    }
                },
                "OUTPUT_SECT_LOCATION": {
                    "OPT_MAX_MINMAX_ALL": "Abs Max"
                }
            }
        }
    }
    # API调用
    MidasAPI("POST", "/VIEW/CAPTURE", Beam_Stresses_dict)


# picture_path：图片储存路径, str
# picture_name：图片名, str
# active_identity：激活的结构组名称，可以有多个结构组, list
# load_case_comb：所需的荷载组合名, str
# components_comp：需要查询的变形方向，Dz, str
# 图片处理函数——结构变形图
def Picture_Deformed(picture_path, picture_name, active_identity, load_case_comb, components_comp, HORIZONTAL=45, VERTICAL=30):
    # 图片储存路径
    picture_file_path = os.path.join(picture_path, picture_name)
    Beam_Deformed_dict = {
        "Argument": {
            "SET_MODE": "post",
            "EXPORT_PATH": picture_file_path,
            "ANGLE": {
                        "HORIZONTAL": HORIZONTAL,
                        "VERTICAL": VERTICAL
            },
            "ACTIVE": {
                # "ACTIVE_MODE": "All"
                "ACTIVE_MODE": "Identity",
                "IDENTITY_TYPE": "Group",
                "IDENTITY_LIST": active_identity
            },
            "RESULT_GRAPHIC": {
                "CURRENT_MODE": "Deformed Shape",
                "LOAD_CASE_COMB": {
                    "TYPE": "CB",
                    "MINMAX": "max",
                    "NAME": load_case_comb,
                    "STEP_INDEX": 1,
                    "TH_OPTION": "Displacement"
                },
                "COMPONENTS": {
                    "COMP": components_comp,
                    # "OPT_LOCAL_CHECK": True,
                    "OPT_LOCAL_CHECK": False,
                },
                "TYPE_OF_DISPLAY": {
                    "CONTOUR": {
                        "OPT_CHECK": True,
                    },
                    "VALUES": {
                        "OPT_CHECK": True,
                        # "DECIMAL_PT": 1
                        "MINMAX_ONLY": {
                            "MAXMIN": "absmax",
                            "LIMIT_SCALE": 0
                        }
                    },
                    "LEGEND": {
                        "OPT_CHECK": True,
                        # "VALUE_EXP": False,
                        # "DECIMAL_PT": 1
                    },
                    "MIRRORED": {
                        "OPT_CHECK": True,
                    },
                    "DEFORM": {
                        "OPT_CHECK": True,
                        # "OPT_CHECK": False,
                    },
                }
            }
        }
    }
    # API调用
    MidasAPI("POST", "/VIEW/CAPTURE", Beam_Deformed_dict)


# 图片处理函数——反力
def Picture_Reaction(picture_path, picture_name, active_identity, load_case_comb, components_comp):
    # 图片储存路径
    picture_file_path = os.path.join(picture_path, picture_name)
    Reaction_dict = {
        "Argument": {
            "SET_MODE": "post",
            "EXPORT_PATH": picture_file_path,
            "ANGLE": {
                "HORIZONTAL":0,
                "VERTICAL":90
            },
            "ACTIVE": {
                "ACTIVE_MODE": "Identity",
                "IDENTITY_TYPE": "Boundary Group",
                "IDENTITY_LIST": active_identity
            },
            "RESULT_GRAPHIC": {
                "CURRENT_MODE": "reaction forces/moments",
                "LOAD_CASE_COMB": {
                    "TYPE": "CB",
                    "MINMAX": "max",
                    "NAME": load_case_comb,
                    "STEP_INDEX": 1,
                },
                "COMPONENTS": {
                    "COMP": components_comp,
                    "OPT_LOCAL_CHECK": False,
                },
                "TYPE_OF_DISPLAY": {
                    "VALUES": {
                        "OPT_CHECK": True,
                    },
                    "LEGEND": {
                        "OPT_CHECK": True,
                    },
                    "ARROW_SCALE_FACTOR": 1.0
                }
            }
        }
    }
    # API调用
    MidasAPI("POST", "/VIEW/CAPTURE", Reaction_dict)


# table_name：内力命名
# structure_group_name：需要查询的结构组名称, str
# load_case_name：所需的荷载组合名, "基本组合(CB)"/"标准组合(CB)",str
# 计算结果处理函数——内力
def Value_BeamForce(table_name, structure_group_name, load_case_name):
    Beam_force_dict = {
        "Argument": {
            "TABLE_NAME": table_name,
            "TABLE_TYPE": "BEAMFORCE",
            # "EXPORT_PATH": "D:\\Output.JSON",
            "UNIT": {
                "FORCE": "kN",
                "DIST": "M"
            },
            "STYLES": {
                "FORMAT": "Fixed",
                "PLACE": 1
            },
            "NODE_ELEMS": {
                "STRUCTURE_GROUP_NAME": structure_group_name
            },
            "LOAD_CASE_NAMES": load_case_name,
            # "PARTS": [
            #     "Part I",
            #     "Part J"
            # ],
            "COMPONENTS": [
                "Elem",
                "Axial",
                "Moment-y",
                "Moment-z",
            ]
        }
    }
    post_table_res = MidasAPI("POST", "/post/table", Beam_force_dict)
    return post_table_res


# table_name：应力命名
# structure_group_name：需要查询的结构组名称, str
# load_case_name：所需的荷载组合名, "基本组合(CB)"/"标准组合(CB)",str
# 计算结果处理函数——应力
def Value_BeamStress(table_name, structure_group_name, load_case_name):
    Beam_stress_dict = {
        "Argument": {
            "TABLE_NAME": table_name,
            "TABLE_TYPE": "BEAMSTRESS",
            # "EXPORT_PATH": "D:\\Output.JSON",
            "UNIT": {
                "FORCE": "N",
                "DIST": "MM"
            },
            "STYLES": {
                "FORMAT": "Fixed",
                "PLACE": 1
            },
            "NODE_ELEMS": {
                "STRUCTURE_GROUP_NAME": structure_group_name
            },
            "LOAD_CASE_NAMES": load_case_name,
            # "PARTS": [
            #     "Part I",
            #     "Part J"
            # ],
            "COMPONENTS": [
                "Elem",
                # "Part",
                "Cb(min/max)",
                "Shear-z",
                # "Shear-z"
            ]
        }
    }
    post_table_res = MidasAPI("POST", "/post/table", Beam_stress_dict)
    return post_table_res


# 计算结果处理函数——变形
# 用组名读取结构的变形
def Value_Deformed(table_name, structure_group_name, load_case_name):
    Beam_deformed_dict = {
        "Argument": {
            "TABLE_NAME": table_name,
            "TABLE_TYPE": "DISPLACEMENTG",
            "UNIT": {
                "FORCE": "N",
                "DIST": "MM"
            },
            "STYLES": {
                "FORMAT": "Fixed",
                "PLACE": 1
            },
            "NODE_ELEMS": {
                # "KEYS": structure_group_node_lst
                "STRUCTURE_GROUP_NAME": structure_group_name
            },
            "LOAD_CASE_NAMES": load_case_name,
            "COMPONENTS": [
                "Node",
                "DX",
                "DY",
                "DZ",
            ]
        }
    }
    post_table_res = MidasAPI("POST", "/post/table", Beam_deformed_dict)
    return post_table_res


# 计算结果处理函数——反力
def Value_Reaction(table_name, boundary_group_name, load_case_name):
    Reaction_dict = {
       "Argument": {
        "TABLE_NAME": table_name,
        "TABLE_TYPE": "REACTIONG",
        # "EXPORT_PATH": "C:\\MIDAS\\Result\\Output.JSON",
        "UNIT": {
            "FORCE": "kN",
            "DIST": "m"
        },
        "STYLES": {
            "FORMAT": "Fixed",
            "PLACE": 12
        },
        "COMPONENTS": [
            "Node",
            "Load",
            "FX",
            "FY",
            "FZ",
            "MX",
            "MY",
            "MZ",
            "Mb"
        ],
        "NODE_ELEMS": {
            # "STRUCTURE_GROUP_NAME": "桩底固结",
            "BOUNDARY_GROUP_NAME": boundary_group_name,
            # "KEYS": [],
        },
        "LOAD_CASE_NAMES": load_case_name
        } 
    }
    post_table_res = MidasAPI("POST", "/post/table", Reaction_dict)
    return post_table_res


# 图片处理函数——梁单元内力图（施工阶段）
def Picture_BeamForce_Stage(picture_path, picture_name, active_identity, stage_name, load_case_comb, components_comp, minmax, HORIZONTAL=45, VERTICAL=30):
    # 图片储存路径
    picture_file_path = os.path.join(picture_path, picture_name)
    Stage_Beam_force_dict = {
        "Argument": {
            "SET_MODE": "post",
            "EXPORT_PATH": picture_file_path,
            "ACTIVE": {
                # "ACTIVE_MODE": "All"
                "ACTIVE_MODE": "Identity",
                "IDENTITY_TYPE": "Group",
                "IDENTITY_LIST": active_identity
            },
            "STAGE_NAME": stage_name,
            "ANGLE": {
                        "HORIZONTAL": HORIZONTAL,
                        "VERTICAL": VERTICAL
            },
            "RESULT_GRAPHIC": {
                "CURRENT_MODE": "Beam Diagrams",
                "LOAD_CASE_COMB": {
                    "TYPE": "CS",
                    "NAME": load_case_comb,
                    "MINMAX": minmax,
                    "STEP_INDEX": 1
                },
                "COMPONENTS": {
                    "PART": "Total",
                    "COMP": components_comp,
                    # "COMP_SUB": "Maximum"
                },
                "DISPLAY_OPTIONS": {
                    "FIDELITY": "5 Points",
                    "FILL": "line fill",
                    "SCALE": 1.0
                },
                "TYPE_OF_DISPLAY": {
                    "CONTOUR": {
                        "OPT_CHECK": True
                    },
                    "VALUES": {
                        "OPT_CHECK": True,
                        "DECIMAL_PT": 1,
                        "MINMAX_ONLY": {
                            "MAXMIN": "absmax",
                            "LIMIT_SCALE": 0
                        }
                    },
                    "LEGEND": {
                        "OPT_CHECK": True,
                        "VALUE_EXP": False,
                        "DECIMAL_PT": 1
                    },
                    "DEFORM": {
                        "OPT_CHECK": False
                    }
                },
                "OUTPUT_SECT_LOCATION": {
                    "OPT_MAX_MINMAX_ALL": "Min/Max"
                }
            }
        }
    }
    # API调用
    MidasAPI("POST", "/VIEW/CAPTURE", Stage_Beam_force_dict)


# 图片处理函数——梁单元应力图（施工阶段）
def Picture_Beamstress_Stage(picture_path, picture_name, active_identity, stage_name, load_case_comb, components_comp, minmax, angle1, angle2):
    # 图片储存路径
    picture_file_path = os.path.join(picture_path, picture_name)
    Stage_Beam_Stresses_dict = {    
        "Argument": {
            "SET_MODE": "post",
            "EXPORT_PATH": picture_file_path,
            "STAGE_NAME": stage_name,
            "ACTIVE": {
                    # "ACTIVE_MODE": "All"
                    "ACTIVE_MODE": "Identity",
                    "IDENTITY_TYPE": "Group",
                    "IDENTITY_LIST": active_identity
                },
            "ANGLE": {
                        "HORIZONTAL": angle1,
                        "VERTICAL": angle2
            },
            "RESULT_GRAPHIC": {
                "CURRENT_MODE": "Beam Stresses Diagram",
                "LOAD_CASE_COMB": {
                    "TYPE": "CS",
                    "MINMAX": minmax,
                    "NAME": load_case_comb,
                    "STEP_INDEX": 1
                },
                "ACTIVE": {
                    # "ACTIVE_MODE": "All"
                    "ACTIVE_MODE": "Identity",
                    "IDENTITY_TYPE": "Group",
                    "IDENTITY_LIST": active_identity
                },
                "COMPONENTS": {
                    "PART": "Total",
                    "COMP": components_comp,
                    "COMP_SUB": "Maximum"
                },
                "DISPLAY_OPTIONS": {
                    "FIDELITY": "5 Points",
                    "FILL": "line fill",
                    "SCALE": 1.0
                },
                "TYPE_OF_DISPLAY": {
                    "CONTOUR": {
                        "OPT_CHECK": True
                    },
                    "VALUES": {
                        "OPT_CHECK": True,
                        "DECIMAL_PT": 1,
                        "MINMAX_ONLY": {
                            "MAXMIN": "absmax",
                            "LIMIT_SCALE": 0
                        }
                    },
                    "LEGEND": {
                        "OPT_CHECK": True,
                        "VALUE_EXP": False,
                        "DECIMAL_PT": 1
                    },
                    "DEFORM": {
                        "OPT_CHECK": False
                    }
                },
                "OUTPUT_SECT_LOCATION": {
                    "OPT_MAX_MINMAX_ALL": "Abs Max"
                }
            }
        }
    }
    # API调用
    MidasAPI("POST", "/VIEW/CAPTURE", Stage_Beam_Stresses_dict)


# 图片处理函数——结构变形图（施工阶段）
def Picture_Deformed_Stage(picture_path, picture_name, active_identity, stage_name, load_case_comb, components_comp, minmax, angle1,angle2):
    # 图片储存路径
    picture_file_path = os.path.join(picture_path, picture_name)
    Stage_Beam_Deformed_dict = {
        "Argument": {
            "SET_MODE": "post",
            "EXPORT_PATH": picture_file_path,
            "ACTIVE": {
                # "ACTIVE_MODE": "All"
                "ACTIVE_MODE": "Identity",
                "IDENTITY_TYPE": "Group",
                "IDENTITY_LIST": active_identity
            },
            "STAGE_NAME": stage_name,
            "ANGLE": {
                        "HORIZONTAL": angle1,
                        "VERTICAL": angle2
            },
            "RESULT_GRAPHIC": {
                "CURRENT_MODE": "Deformed Shape",
                "LOAD_CASE_COMB": {
                    "TYPE": "CS",
                    "MINMAX": minmax,
                    "NAME": load_case_comb,
                    "STEP_INDEX": 1,
                    "TH_OPTION": "Displacement"
                },
                "COMPONENTS": {
                    "COMP": components_comp,
                    # "OPT_LOCAL_CHECK": True,
                    "OPT_LOCAL_CHECK": False,
                },
                "TYPE_OF_DISPLAY": {
                    "CONTOUR": {
                        "OPT_CHECK": True,
                    },
                    "VALUES": {
                        "OPT_CHECK": True,
                        # "DECIMAL_PT": 1
                        "MINMAX_ONLY": {
                            "MAXMIN": "absmax",
                            "LIMIT_SCALE": 0
                        }
                    },                     
                    "LEGEND": {
                        "OPT_CHECK": True,
                        # "VALUE_EXP": False,                        
                        # "DECIMAL_PT": 1
                    },
                    "MIRRORED": {
                        "OPT_CHECK": True,
                    },
                    "DEFORM": {
                        "OPT_CHECK": True,
                        # "OPT_CHECK": False,
                    },
                }
            }
        }
    }
    # API调用
    MidasAPI("POST", "/VIEW/CAPTURE", Stage_Beam_Deformed_dict)


# 图片处理函数——反力（施工阶段）
def Picture_Reaction_Stage(picture_path, picture_name, active_identity, stage_name, load_case_comb, components_comp, minmax):
    # 图片储存路径
    picture_file_path = os.path.join(picture_path, picture_name)
    Stage_Reaction_dict = {
        "Argument": {
            "SET_MODE": "post",
            "EXPORT_PATH": picture_file_path,
            "ACTIVE": {
                "ACTIVE_MODE": "Identity",
                "IDENTITY_TYPE": "Boundary Group",
                "IDENTITY_LIST": active_identity
            },
            "STAGE_NAME": stage_name,
            "ANGLE": {
                "HORIZONTAL":0,
                "VERTICAL":90
            },
            "RESULT_GRAPHIC": {
                "CURRENT_MODE": "reaction forces/moments",
                "LOAD_CASE_COMB": {
                    "TYPE": "CS",
                    "MINMAX": minmax,
                    "NAME": load_case_comb,
                    "STEP_INDEX": 1,
                },
                "COMPONENTS": {
                    "COMP": components_comp,
                    "OPT_LOCAL_CHECK": False,
                },
                "TYPE_OF_DISPLAY": {
                    "VALUES": {
                        "OPT_CHECK": True,
                    },
                    "LEGEND": {
                        "OPT_CHECK": True,
                    },
                    "ARROW_SCALE_FACTOR": 1.0
                }
            }
        }
    }
    # API调用
    MidasAPI("POST", "/VIEW/CAPTURE", Stage_Reaction_dict)


# 计算结果处理函数——内力（施工阶段）
def Value_BeamForce_Stage(table_name, structure_group_name, stage_step, load_case_name):
    Stage_Beam_force_dict = {
        "Argument": {
            "TABLE_NAME": table_name,
            "TABLE_TYPE": "BEAMFORCE",
            # "EXPORT_PATH": "C:\\MIDAS\\Result\\Output.JSON",
            "UNIT": {
                "FORCE": "kN",
                "DIST": "m"
                },
            "STYLES": {
                "FORMAT": "Fixed",
                "PLACE": 12
                },
            "COMPONENTS": [
                "Elem",
                "Axial",
                "Moment-y",
                "Moment-z",
            ],
            "NODE_ELEMS": {
                "STRUCTURE_GROUP_NAME": structure_group_name
            },
            "LOAD_CASE_NAMES": load_case_name,
            "OPT_CS": True,
            "STAGE_STEP": stage_step
            }
        }
    post_table_res = MidasAPI("POST", "/post/table", Stage_Beam_force_dict)
    return post_table_res


# 计算结果处理函数——应力（施工阶段）
def Value_BeamStress_Stage(table_name, structure_group_name, stage_step, load_case_name):
    Stage_BeamStress_dict = {
        "Argument": {
            "TABLE_NAME": table_name,
            "TABLE_TYPE": "BEAMSTRESS",
            "UNIT": {
                "FORCE": "N",
                "DIST": "mm"
            },
            "STYLES": {
                "FORMAT": "Fixed",
                "PLACE": 12
                },
            "COMPONENTS": [
                "Elem",
                "Cb(min/max)",
                "Shear-z",
            ],
            "NODE_ELEMS": {
                "STRUCTURE_GROUP_NAME": structure_group_name
            },
            "LOAD_CASE_NAMES": load_case_name,
            # "PARTS": ["Part I","Part J"],
            "OPT_CS": True,
            "STAGE_STEP": stage_step
            }
        }
    post_table_res = MidasAPI("POST", "/post/table", Stage_BeamStress_dict)
    return post_table_res


# 计算结果处理函数——变形（施工阶段）
def Value_Deformed_Stage(table_name, structure_group_name, stage_step, load_case_name):
    Stage_Beam_deformed_dict = {
        "Argument": {
            "TABLE_NAME": table_name,
            "TABLE_TYPE": "DISPLACEMENTG",
            "UNIT": {
                "FORCE": "N",
                "DIST": "mm"
            },
            "STYLES": {
                "FORMAT": "Fixed",
                "PLACE": 1
            },
            "COMPONENTS": [
                "Node",
                "DX",
                "DY",
                "DZ",
            ],
            "NODE_ELEMS": {
                # "KEYS": structure_group_node_lst
                "STRUCTURE_GROUP_NAME": structure_group_name
            },
            "LOAD_CASE_NAMES": load_case_name,
            "OPT_CS": True,
            "STAGE_STEP": stage_step
        }
    }
    post_table_res = MidasAPI("POST", "/post/table", Stage_Beam_deformed_dict)
    return post_table_res


# 计算结果处理函数——反力（施工阶段）
def Value_Reaction_Stage(table_name, boundary_group_name, stage_step, load_case_name):
    Stage_Reaction_dict = {
        "Argument": {
            "TABLE_NAME": table_name,
            "TABLE_TYPE": "REACTIONG",
            # "EXPORT_PATH": "C:\\MIDAS\\Result\\Output.JSON",
            "UNIT": {
                "FORCE": "kN",
                "DIST": "m"
            },
            "STYLES": {
                "FORMAT": "Fixed",
                "PLACE": 12
            },
            "COMPONENTS": [
                "Node",
                "Load",
                "FX",
                "FY",
                "FZ",
                "MX",
                "MY",
                "MZ",
                "Mb"
            ],
            "NODE_ELEMS": {
                # "STRUCTURE_GROUP_NAME": "桩底固结",
                "BOUNDARY_GROUP_NAME": boundary_group_name,
                # "KEYS": [],
            },
            "LOAD_CASE_NAMES": load_case_name,
            "OPT_CS": True,
            "STAGE_STEP": stage_step
        }
    }
    post_table_res = MidasAPI("POST", "/post/table", Stage_Reaction_dict)
    return post_table_res


#=====================================================================================================================================================================

# 计算结果处理函数——内力(单个单元)
def Value_BeamForce_Elem(table_name, Elem_eid, load_case_name):
    Beam_force_dict = {
        "Argument": {
            "TABLE_NAME": table_name,
            "TABLE_TYPE": "BEAMFORCE",
            # "EXPORT_PATH": "C:\\MIDAS\\Result\\Output.JSON",
            "UNIT": {
                "FORCE": "kN",
                "DIST": "m"
                },
            "STYLES": {
                "FORMAT": "Fixed",
                "PLACE": 12
                },
            "COMPONENTS": [
                "Elem",
                "Axial",
                "Shear-y",
                "Shear-z",
                "Torsion",
                "Moment-y",
                "Moment-z",
            ],
            "LOAD_CASE_NAMES": load_case_name,
            "NODE_ELEMS": {
                "KEYS": Elem_eid
            },
            }
        }
    post_table_res = MidasAPI("POST", "/post/table", Beam_force_dict)
    return post_table_res


# 计算结果处理函数——内力(单个单元-施工阶段)
def Value_BeamForce_Elem_Stage(table_name, Elem_eid, load_case_name, stage_step):
    Beam_force_dict = {
        "Argument": {
            "TABLE_NAME": table_name,
            "TABLE_TYPE": "BEAMFORCE",
            # "EXPORT_PATH": "C:\\MIDAS\\Result\\Output.JSON",
            "UNIT": {
                "FORCE": "kN",
                "DIST": "m"
                },
            "STYLES": {
                "FORMAT": "Fixed",
                "PLACE": 12
                },
            "COMPONENTS": [
                "Elem",
                "Axial",
                "Shear-y",
                "Shear-z",
                "Torsion",
                "Moment-y",
                "Moment-z",
            ],
            "NODE_ELEMS": {
                "KEYS": Elem_eid
            },
            "LOAD_CASE_NAMES": load_case_name,
            "OPT_CS": True,
            "STAGE_STEP": stage_step
            }
        }
    post_table_res = MidasAPI("POST", "/post/table", Beam_force_dict)
    return post_table_res


# 单个施工阶段分析
def midas_stage(NAME_lst, STEP_lst, AELEM_lst, DELEM_lst, ABNDR_lst, DBNDR_lst, ALOAD_lst, DLOAD_lst):
    NAME_line_lst  = ['NAME=', ', '.join([str(x) for x in NAME_lst])] # [插打钢管桩, 0, YES, NO, NO, 5'] ➡ '插打钢管桩, 0, YES, NO, NO, 5'
    STEP_line_lst  = ['STEP=', ', '.join([str(x) for x in STEP_lst])] # [1, 2, 3, 4, 5] ➡ 'DAY1, DAY2, ..'
    AELEM_line_lst = ['AELEM=', ', '.join(chain.from_iterable(AELEM_lst))] # [(钢管桩, '0'), (第一层围檩内支撑, '0')] ➡ '钢管桩, 0, 第一层围檩内支撑, 0'
    DELEM_line_lst = ['DELEM=', ', '.join(chain.from_iterable(DELEM_lst))] # [(钢管桩, '0'), (第一层围檩内支撑, '0')] ➡ '钢管桩, 0, 第一层围檩内支撑, 0'
    ABNDR_line_lst = ['ABNDR=', ', '.join(chain.from_iterable(ABNDR_lst))] # [(第一层围檩下土弹簧, DEFORMED)] ➡ '第一层围檩下土弹簧, DEFORMED'
    DBNDR_line_lst = ['DBNDR=', ', '.join([group for group, age in DBNDR_lst])] # [(第一层围檩下土弹簧, DEFORMED)] ➡ '第一层围檩下土弹簧'
    ALOAD_line_lst = ['ALOAD=', ', '.join(chain.from_iterable(ALOAD_lst))] # [(自重, FIRST), (外部土压力, LAST)] ➡ '自重, FIRST, 外部土压力, LAST'
    DLOAD_line_lst = ['DLOAD=', ', '.join(chain.from_iterable(DLOAD_lst))] # [(初始内部水压力, FIRST)] ➡ '初始内部水压力, FIRST'
    STAGE_line_lst = [x[0] + x[1] for x in [NAME_line_lst, STEP_line_lst, AELEM_line_lst, DELEM_line_lst, ABNDR_line_lst, DBNDR_line_lst, ALOAD_line_lst, DLOAD_line_lst] if x[1] != '']
    return STAGE_line_lst


def Value_Deformed_ByNodes(table_name, node_keys, load_case_name):
    """按节点号查询变形数据（结构与 Value_Deformed 对齐）。

    table_name:      表名（任意字符串，用于结果字典的键）
    node_keys:       节点号列表，如 ['101', '102', '201']
    load_case_name:  荷载组合名，如 '标准组合(CB:all)'

    返回格式与 Value_Deformed 一致:
        {table_name: {DATA: [[荷载组合, 节点号, DX, DY, DZ], ...]}}
    """
    Node_deformed_dict = {
        "Argument": {
            "TABLE_NAME": table_name,
            "TABLE_TYPE": "DISPLACEMENTG",
            "UNIT": {
                "FORCE": "N",
                "DIST": "MM"
            },
            "STYLES": {
                "FORMAT": "Fixed",
                "PLACE": 1
            },
            "NODE_ELEMS": {
                "KEYS": [node_keys]
            },
            "LOAD_CASE_NAMES": load_case_name,
            "COMPONENTS": [
                "Node",
                "DX",
                "DY",
                "DZ",
            ]
        }
    }
    post_table_res = MidasAPI("POST", "/post/table", Node_deformed_dict)
    return post_table_res