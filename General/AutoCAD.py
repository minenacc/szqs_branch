# 1. 标准库
import os
import time
import math
import winreg

# 2. 第三方库
import pythoncom
import win32com.client

# 3. 本地模块
from General.DataUtils import pointlist_extend
from General.Geometry import midpoint, is_point_inside_polygon, check_polyline_intersection, isPolylineWithinPolyline, Centroid,  midptcoordonarc


def _com_retry(func, *args, retries=3, delay=0.1, **kwargs):
    """
    COM 调用重试包装器。
    当 AutoCAD 忙碌时 (RPC_E_CALL_REJECTED)，自动等待并重试。
    """
    for attempt in range(retries):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            if attempt < retries - 1:
                time.sleep(delay)
            else:
                raise


# ===========================================================================
# COM 消息过滤器：解决 AutoCAD 忙时 "被呼叫方拒绝"(RPC_E_CALL_REJECTED)
# 本环境的 pywin32 未暴露 CoRegisterMessageFilter，故用 ctypes 直接实现 IMessageFilter。
# ===========================================================================
import ctypes


class _GUID(ctypes.Structure):
    _fields_ = [("Data1", ctypes.c_uint32), ("Data2", ctypes.c_uint16),
                ("Data3", ctypes.c_uint16), ("Data4", ctypes.c_ubyte * 8)]


def _make_guid(d1, d2, d3, d4):
    return _GUID(d1, d2, d3, (ctypes.c_ubyte * 8)(*d4))


_IID_IUnknown = _make_guid(0x00000000, 0, 0, (0xC0, 0, 0, 0, 0, 0, 0, 0x46))
_IID_IMessageFilter = _make_guid(0x00000016, 0, 0, (0xC0, 0, 0, 0, 0, 0, 0, 0x46))

_HRESULT = ctypes.c_long
_QI_t = ctypes.WINFUNCTYPE(_HRESULT, ctypes.c_void_p, ctypes.POINTER(_GUID), ctypes.POINTER(ctypes.c_void_p))
_AddRef_t = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)
_Release_t = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)
_Handle_t = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p, ctypes.c_ulong, ctypes.c_void_p, ctypes.c_ulong, ctypes.c_void_p)
_Retry_t = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_ulong, ctypes.c_ulong)
_Pending_t = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_ulong, ctypes.c_ulong)


class _IMessageFilterVtbl(ctypes.Structure):
    _fields_ = [("QueryInterface", _QI_t), ("AddRef", _AddRef_t), ("Release", _Release_t),
                ("HandleInComingCall", _Handle_t), ("RetryRejectedCall", _Retry_t),
                ("MessagePending", _Pending_t)]


class _IMessageFilterObj(ctypes.Structure):
    _fields_ = [("lpVtbl", ctypes.POINTER(_IMessageFilterVtbl))]


_acad_message_filter_refs = None


def register_acad_message_filter(retry_delay_ms=100):
    """
    注册 COM 消息过滤器：当 AutoCAD 正忙拒绝调用（被呼叫方拒绝 RPC_E_CALL_REJECTED /
    SERVERCALL_RETRYLATER）时，自动等待 retry_delay_ms 后重试，而不是抛错。
    幂等；需在发起 COM 调用的线程调用。返回 True/False。
    """
    global _acad_message_filter_refs
    if _acad_message_filter_refs is not None:
        return True
    try:
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass

        def _is_iid(riid, iid):
            return ctypes.string_at(ctypes.cast(riid, ctypes.c_void_p), 16) == bytes(iid)

        def _qi(this, riid, ppv):
            if _is_iid(riid, _IID_IUnknown) or _is_iid(riid, _IID_IMessageFilter):
                ppv[0] = this
                return 0
            ppv[0] = None
            return -2147467262  # E_NOINTERFACE

        def _addref(this):
            return 1

        def _release(this):
            return 0

        def _handle(this, dwCallType, htaskCaller, dwTickCount, lpInterfaceInfo):
            return 0  # SERVERCALL_ISHANDLED

        def _retry(this, htaskCallee, dwTickCount, dwRejectType):
            # 1=SERVERCALL_REJECTED, 2=SERVERCALL_RETRYLATER
            if dwRejectType in (1, 2):
                return retry_delay_ms
            return 0xFFFFFFFF  # -1 取消

        def _pending(this, htaskCallee, dwTickCount, dwPendingType):
            return 2  # PENDINGMSG_WAITDEFPROCESS

        vtbl = _IMessageFilterVtbl(
            _QI_t(_qi), _AddRef_t(_addref), _Release_t(_release),
            _Handle_t(_handle), _Retry_t(_retry), _Pending_t(_pending))
        obj = _IMessageFilterObj(ctypes.pointer(vtbl))
        ptr = ctypes.cast(ctypes.pointer(obj), ctypes.c_void_p)

        ole32 = ctypes.windll.ole32
        ole32.CoRegisterMessageFilter.restype = _HRESULT
        ole32.CoRegisterMessageFilter.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
        hr = ole32.CoRegisterMessageFilter(ptr, None)
        if hr != 0:
            print("注册 COM 消息过滤器失败, hr =", hr)
            return False
        _acad_message_filter_refs = (vtbl, obj, ptr)  # 保持引用，防止被 GC
        return True
    except Exception as e:
        print("注册 COM 消息过滤器异常:", e)
        return False


def find_installed_autocad_versions():
    """
    通过查询Windows注册表来查找已安装的AutoCAD版本
    返回一个包含版本号和安装路径的字典
    """
    # autocad_versions = {}
    version_key_name_lst = []
    # AutoCAD通常在注册表的这两个位置记录信息
    registry_paths = [
        r"SOFTWARE\Autodesk\AutoCAD",  # 64位系统下的主要路径
        r"SOFTWARE\WOW6432Node\Autodesk\AutoCAD"  # 32位软件在64位系统下的路径
    ]
    for registry_path in registry_paths:
        try:
            # 尝试从HKEY_LOCAL_MACHINE打开注册表项
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, registry_path)
            # 获取所有子键（这些通常对应不同版本）
            for i in range(winreg.QueryInfoKey(key)[0]):
                version_key_name = winreg.EnumKey(key, i)
                version_key_name_lst.append(version_key_name)
            winreg.CloseKey(key)
        except FileNotFoundError:
            # 如果注册表路径不存在，继续尝试下一个
            continue
        except PermissionError:
            print(f"权限不足，无法访问注册表路径: {registry_path}")
            print("请以管理员身份运行")
            break
        except Exception as e:
            print(f"访问注册表时发生错误: {e}")
            continue
    return version_key_name_lst


# 每个版本的cad对应的ProgID
def AutoCAD_VersionR_tans_to_COM_ProgID(VersionR):
    Version_ProgID_dict = {
        'R17.0': 'AutoCAD.Application.17.0', # 2007
        'R17.1': 'AutoCAD.Application.17.1', # 2008
        'R17.2': 'AutoCAD.Application.17.2', # 2009
        'R18.0': 'AutoCAD.Application.18.0', # 2010
        'R18.1': 'AutoCAD.Application.18.1', # 2011
        'R18.2': 'AutoCAD.Application.18.2', # 2012
        'R19.0': 'AutoCAD.Application.19.0', # 2013
        'R19.1': 'AutoCAD.Application.19.1', # 2014
        'R20.0': 'AutoCAD.Application.20.0', # 2015
        'R20.1': 'AutoCAD.Application.20.1', # 2016
        'R21.0': 'AutoCAD.Application.21.0', # 2017
        'R22.0': 'AutoCAD.Application.22.0', # 2018
        'R23.0': 'AutoCAD.Application.23.0', # 2019
        'R23.1': 'AutoCAD.Application.23.1', # 2020
        'R24.0': 'AutoCAD.Application.24.0', # 2021
        'R24.1': 'AutoCAD.Application.24.1', # 2022
        'R24.2': 'AutoCAD.Application.24.2', # 2023
        'R24.3': 'AutoCAD.Application.24.3', # 2024
    }
    ProgID = Version_ProgID_dict[VersionR]
    return ProgID


def get_acad_application(progid='AutoCAD.Application', visible=True):
    """
    获取当前可用的 AutoCAD 应用程序实例。
    优先从运行对象表(ROT)取活动实例，失败再 Dispatch(连接到已运行实例或新建)。
    用户关闭/切换 CAD 后旧代理会失效，重新调用本函数即可重新挂接。
    """
    app = None
    try:
        app = win32com.client.GetActiveObject(progid)
    except Exception:
        app = None
    if app is None:
        app = win32com.client.Dispatch(progid)
    if visible:
        try:
            app.Visible = True
        except Exception:
            pass
    return app


def ensure_acad_document(app, create_if_none=True):
    """
    确保 AutoCAD 存在活动文档与模型空间。
    当 CAD 停留在“开始页”或未打开任何图纸时 ActiveDocument 为 None，
    此时自动新建一张空白图纸(create_if_none=True)。返回 (doc, msp)。
    """
    doc = None
    try:
        doc = app.ActiveDocument
    except Exception:
        doc = None
    if doc is None:
        if not create_if_none:
            raise RuntimeError("AutoCAD 未打开任何图纸")
        doc = app.Documents.Add()   # 停留在开始页/无图纸：新建空白图纸
    msp = doc.ModelSpace            # 访问一次，确认文档代理有效
    _ = doc.Name
    return doc, msp


def prepare_acad(progid='AutoCAD.Application', visible=True, create_if_none=True):
    """
    一步获取可用的 (app, doc, msp)：重新挂接活动实例 + 确保有图纸/模型空间。
    用于每次绘图前，解决“关闭并切换 CAD”“停留在开始页无模型空间”等问题。
    """
    app = get_acad_application(progid, visible)
    doc, msp = ensure_acad_document(app, create_if_none)
    return app, doc, msp


# 检测当前CAD文件中是否含有指定块名的块
def if_block_exist(doc, block_namestring):
    # 遍历当前CAD文件中所有的块名
    blocks = doc.Blocks
    for block in blocks:
        if block.Name == block_namestring:  
            return True
    print("未检测到已导入预定义块信息，开始尝试导入预定义块。")
    return False


def Import_Blocks(Applocation, acad):
    # 等待 COM 连接就绪（刚连接到 AutoCAD 时 ActiveDocument 可能还未完全初始化）
    time.sleep(0.5)
    for _attempt in range(3):
        try:
            doc = acad.ActiveDocument
            msp = doc.ModelSpace
            # 验证 doc 可用
            _ = doc.Name
            break
        except Exception:
            time.sleep(0.5)
    else:
        print("Import_Blocks: 无法获取 ActiveDocument，跳过块导入")
        return
    dwg_name = "MBS_Suppot_Drawing"
    path = os.path.join(Applocation, f"Support\\templates\\{dwg_name}.dwg")
    # 检查文件是否存在
    if not os.path.exists(path):
        print(f"错误：文件不存在 - {path}")
        return
    # 检查块是否已存在
    if not if_block_exist(doc, dwg_name):
        point = [0, 0, 0]
        insert_point = vtpnt(point)
        # 保存当前系统变量
        osmode = doc.GetVariable("OSMODE")
        filedia = doc.GetVariable("FILEDIA")
        # 关闭捕捉和文件对话框
        doc.SetVariable("OSMODE", 0)
        doc.SetVariable("FILEDIA", 0)
        try:
            # 插入外部 DWG 作为块引用
            dwg_block = msp.InsertBlock(insert_point, path, 1, 1, 1, 0)
            time.sleep(0.5)
            # 检查返回对象是否有效
            if dwg_block is None:
                raise Exception("InsertBlock 返回空对象，插入失败")
            # 尝试删除块引用
            if hasattr(dwg_block, "Delete"):
                dwg_block.Delete()
            elif hasattr(dwg_block, "Erase"):
                dwg_block.Erase()
            else:
                # 最后手段：发送 AutoCAD 命令删除最后一个对象
                doc.SendCommand("_.ERASE L ")
            print("已成功导入预定义块。")
        except Exception as e:
            print(f"操作过程中发生异常: {e}")
            # 可在此处添加额外调试：检查 AutoCAD 是否有对话框弹出
        finally:
            # 恢复系统变量前稍作等待，确保 AutoCAD 已就绪
            time.sleep(0.5)
            # 恢复文件对话框设置
            try:
                doc.SetVariable("FILEDIA", filedia)
            except Exception as e:
                print(f"恢复 FILEDIA 失败: {e}")
            # 恢复捕捉模式
            try:
                doc.SetVariable("OSMODE", osmode)
            except Exception as e:
                print(f"恢复 OSMODE 失败: {e}")
    else:
        print("检测到已导入预定义块信息。")


# 求autocad中*line在polygon中的长度
# xline、polygon为autocad中的object对象
def length_of_line_inside_Polygon(xline, polygon, accuracy):
    inters = polygon.IntersectWith(xline , 0)
    crosspts= [[inters[i], inters[i+1], inters[i+2]] for i in range(0, len(inters), 3)]
    # print(crosspts)
    polygon_pts = getptsfromplwithcricle(polygon,1)
    if len(crosspts) <= 1:
        sumlen = 0
    elif len(crosspts) == 2:
        midpt1 = midpoint(crosspts[0], crosspts[1])
        if is_point_inside_polygon(midpt1, polygon_pts):
            len1 = round(abs(crosspts[1][1] - crosspts[0][1]),accuracy)
        else:
            len1 = 0
        sumlen = len1
    elif len(crosspts) == 3:
        midpt1 = midpoint(crosspts[0], crosspts[1])
        midpt2 = midpoint(crosspts[1], crosspts[2])
        if is_point_inside_polygon(midpt1, polygon_pts):
            len1 = round(abs(crosspts[1][1] - crosspts[0][1]),accuracy)
        else:
            len1 = 0
        if is_point_inside_polygon(midpt2, polygon_pts):
            len2 = round(abs(crosspts[2][1] - crosspts[1][1]),accuracy)
        else:
            len2 = 0
        sumlen = len1 + len2
    elif len(crosspts) == 4:
        midpt1 = midpoint(crosspts[0], crosspts[1])
        midpt2 = midpoint(crosspts[1], crosspts[2])
        midpt3 = midpoint(crosspts[2], crosspts[3])
        if is_point_inside_polygon(midpt1, polygon_pts):
            len1 = round(abs(crosspts[1][1] - crosspts[0][1]),accuracy)
        else:
            len1 = 0
        if is_point_inside_polygon(midpt2, polygon_pts):
            len2 = round(abs(crosspts[2][1] - crosspts[1][1]),accuracy)
        else:
            len2 = 0
        if is_point_inside_polygon(midpt3, polygon_pts):
            len3 = round(abs(crosspts[3][1] - crosspts[2][1]),accuracy)
        else:
            len3 = 0
        sumlen = len1 + len2 + len3
    elif len(crosspts) == 5:
        midpt1 = midpoint(crosspts[0], crosspts[1])
        midpt2 = midpoint(crosspts[1], crosspts[2])
        midpt3 = midpoint(crosspts[2], crosspts[3])
        midpt4 = midpoint(crosspts[3], crosspts[4])
        if is_point_inside_polygon(midpt1, polygon_pts):
            len1 = abs(crosspts[1][1] - crosspts[0][1])
        else:
            len1 = 0
        if is_point_inside_polygon(midpt2, polygon_pts):
            len2 = abs(crosspts[2][1] - crosspts[1][1])
        else:
            len2 = 0
        if is_point_inside_polygon(midpt3, polygon_pts):
            len3 = abs(crosspts[3][1] - crosspts[2][1])
        else:
            len3 = 0
        if is_point_inside_polygon(midpt4, polygon_pts):
            len4 = abs(crosspts[4][1] - crosspts[3][1])
        else:
            len4 = 0
        sumlen = round(len1 + len2 + len3 + len4,accuracy)
    return sumlen


def Group_Select_Sections(acadapp, doc):    
    ss = SelectOnScreen(acadapp, "请选择截面")[2]
    polyline_ss = ss_to_lst(ss)
    # 截面选择集分组词典
    dict_ss = {}
    # 记录截面个数
    section_num = 0
    # 记录外轮廓的句柄
    handle_outline = []
    for i in range(len(polyline_ss)):
        # 第i个图形
        ssi = polyline_ss[i]
        # 全部的图形 去掉 当前图形
        lsti = polyline_ss.copy()
        lsti.remove(ssi)
        obji = doc.HandleToObject(ssi)
        relation_lst = []
        # 当前图形 和 其余图形 比较
        for j in range(len(lsti)):
            objj = doc.HandleToObject(lsti[j])
            # 如果两多段线没有交点
            if not check_polyline_intersection(obji, objj):
                # 判断两多段线的关系，在不相交的前提下，是包含还是相离
                relation = (check_relationship(obji, objj))
                relation_lst.append(relation)
                if relation == "1包含2":
                    # 1是外轮廓，2是内轮廓，说明此时ssi是一个截面的外轮廓，lsti[j]是该截面的内轮廓
                    if ssi not in handle_outline:
                        section_num = section_num + 1
                    # if f"第{section_num}个截面外轮廓" not in dict_ss:
                    #     dict_ss[f"第{section_num}个截面外轮廓"] = []
                    if f"第{section_num}个截面内轮廓" not in dict_ss:
                        dict_ss[f"第{section_num}个截面内轮廓"] = []
                    dict_ss[f"第{section_num}个截面外轮廓"] = [ssi]
                    dict_ss[f"第{section_num}个截面内轮廓"].append(lsti[j])
                    handle_outline.append(ssi)
            else:
                # relation_lst.append("相交")
                print("出现多段线相交的情况或是未识别到多段线，请检查截面及布局。")
                return None
        # 全部相离说明是一个单截面
        if all(x == "相离" for x in relation_lst):
            section_num = section_num + 1
            # if f"第{section_num}个截面外轮廓" not in dict_ss:
            #     dict_ss[f"第{section_num}个截面外轮廓"] = []
            if f"第{section_num}个截面内轮廓" not in dict_ss:
                dict_ss[f"第{section_num}个截面内轮廓"] = []
            dict_ss[f"第{section_num}个截面外轮廓"] = [ssi]
    # print(dict_ss)
    return section_num, dict_ss


# 从CAD中获取截面的相关信息
def get_SECinfo_fromcad(acadapp, acaddoc, acadmsp):
    #根据句柄对应的图元的面积排序，entlst表中包含的元素是每个图元的句柄，HandleToObject可以根据图元的句柄获取CAD中的对象
    #a是ss_to_lst(slt)生成的表中的元素，slt是选择集，ss_to_lst(slt)返回选择集中所有所需图元的句柄
    section_group_reslst = Group_Select_Sections(acadapp, acaddoc)
    section_num = section_group_reslst[0]
    section_dict =  section_group_reslst[1]
    # print(section_num)
    # print(section_dict)
    height_dict = {}
    vertexxcoords = []
    # 词典中储存的内轮廓和外轮廓为图元句柄
    for ii in range(section_num):
        outline = section_dict[f"第{ii+1}个截面外轮廓"][0]
        outlinepts = getptsfromplwithcricle(acaddoc.HandleToObject(outline),1)
        holes = []
        vertexpts = outlinepts
        for ent in section_dict[f"第{ii+1}个截面内轮廓"]:
            holepts =getptsfromplwithcricle(acaddoc.HandleToObject(ent),1)
            vertexpts = vertexpts + holepts
            if isPolylineWithinPolyline(holepts,outlinepts):
                holes.append(ent)
        holes = sorted(holes, key=lambda a: Centroid(getptsfromplwithcricle(acaddoc.HandleToObject(a),1)), reverse=True)
        # print(holes)
        # 得到多段线端点和圆弧中点对应的排序后的x坐标
        vertexxcoord = list(sorted (set (map(lambda x: round(x[0],1), vertexpts))))
        vertexxcoords = vertexxcoords + vertexxcoord
        # print(vertexxcoord)
        # 生成射线
        xlinexcoord = []
        for i in range(len(vertexxcoord)):
            xlinexcoord.append(vertexxcoord[i] - 0.1)
            xlinexcoord.append(vertexxcoord[i] + 0.1)
        #生成射线，addxline根据两点生成射线
        xlinelst = []
        for xcoord in xlinexcoord:
            xline = acadmsp.AddXLine(vtpnt2([xcoord,0,0]),vtpnt2([x + y for x, y in zip([xcoord,0,0] , [0,10,0])]))
            xlinelst.append(xline) 
        # print(xlinelst)
        #得到每条射线的外轮廓高度
        outlinedata = []
        for xline in xlinelst:
            outlinedata.append(length_of_line_inside_Polygon(xline, acaddoc.HandleToObject(outline), 1))
        # print(outlinedata)
        #得到每条射线的内轮廓高度及差值
        if holes != []:
            holesdata = []
            for hole in holes:
                holedata = []
                for xline in xlinelst:
                    holedata.append(length_of_line_inside_Polygon(xline, acaddoc.HandleToObject(hole), 0))
                # print(holedata)
                if holesdata == []:
                    holesdata = holedata
                elif holesdata != []:
                    holesdata = [x + y for x, y in zip(holesdata, holedata)]
            # print(holesdata)
            hightlst = [x - y for x, y in zip(outlinedata, holesdata)]
        else:
            hightlst = outlinedata
        for xline in xlinelst:
            xline.Delete()
        hightlst = [round(h,1) for h in hightlst] 
        height_dict[f"第{ii+1}个截面CAD横坐标及截面高度"] = [xlinexcoord, hightlst]
    # 根据截面横坐标和对应的高度将词典重新排序
    # 提取值并排序
    sorted_height_dict_values = sorted(height_dict.values(), key=lambda x: x[0][0])
    # 创建一个辅助字典来存储排序后的值
    sorted_height_dict_data = {key: value for key, value in zip(height_dict.keys(), sorted_height_dict_values)}
    # 将排序后的值映射回原来的键
    sorted_height_dict = {key: sorted_height_dict_data[key] for key in height_dict.keys()}
    # print(sorted_height_dict)
    all_xlinexcoord = [round(x, 1) for key, value in sorted_height_dict.items() for x in value[0]]
    all_heightlst = [round(x, 1) for key, value in sorted_height_dict.items() for x in value[1]]
    # print(all_xlinexcoord)
    # print(all_heightlst)
    vertexxcoords.sort()
    # print(vertexxcoords)
    if all(x == 0 for x in all_heightlst):
        print("Warning: 截面计算高度结果为0, 请检查多段线是否在Z=0平面上")
    entmake_sec_load(acaddoc, acadmsp, all_xlinexcoord, all_heightlst)
    # 返回活动文档和模型空间、多段线x坐标
    return acaddoc, acadmsp, vertexxcoords, all_xlinexcoord, all_heightlst


def entmake_sec_load(doc, msp, xlinexcoord, heightlst):
    # 询问是否生成截面，只有当输入y时，才会开始生成
    user_select = doc.Utility.getstring (1, "是否生成截面线荷载图示，选择Y后选取生成点(Y/N)")
    # print(user_select)
    if user_select.lower() == 'y':
        # 根据指定位置插入截面的线荷载图
        pl_startpoint = doc.Utility.Getpoint()
        py_xlst = [x - xlinexcoord[0]  for x in xlinexcoord]
        # 绘制底部多段线
        pointlst1 = [[pl_startpoint[0] + x, pl_startpoint[1], pl_startpoint[2]] for x in py_xlst]
        #需要将所有坐标去括号变成[1,2,0,2,2,0,3,2,0]的形式
        pl_pointlst1 = [y for x in pointlst1 for y in x]
        pl_var_ptlst1 = vtfloat(pl_pointlst1)
        polyline1 = msp.AddPolyline(pl_var_ptlst1)
        # 绘制顶部多段线
        pointlst2 = []
        for i in range(len(heightlst)):
            pt = pointlst1[i]
            pointlst2.append([pt[0], pt[1] + heightlst[i], pt[2]])
        pl_pointlst2 = [y for x in pointlst2 for y in x]
        pl_var_ptlst2 = vtfloat(pl_pointlst2)
        polyline2 = msp.AddPolyline(pl_var_ptlst2)
        # 绘制线段
        for i in range(len(pointlst1)):
            pt1 = pointlst1[i]
            pt2 = pointlst2[i]
            var_pt1 = vtpnt(pt1)
            var_pt2 = vtpnt(pt2)
            line_i = msp.AddLine(var_pt1, var_pt2)
        print("成功生成截面线荷载图示")
    elif user_select.lower() == 'n':
        print("已取消截面线荷载图示生成")
    else:
        print("用户取消截面线荷载图生成")


# 识别单根分配梁的长度
def get_lines_from_block(block_reference, layer_name):
    lst = []
    # 分解原块引用
    exploded_entities = block_reference.Explode()
    # print(exploded_entities)
    # 筛选出与指定图层相关的直线
    lines = [e for e in exploded_entities if e.EntityName == "AcDbLine" and e.Layer == layer_name]
    for line in lines:
        lst.append([line.StartPoint, line.EndPoint, line.Length])
    max_length = max(item[2] for item in lst)
    Line_Startpoint = [item[0] for item in lst if item[2] == max_length]
    Line_Endpoint = [item[1] for item in lst if item[2] == max_length]
    other_lines = [e for e in exploded_entities if e not in lines]
    for line in other_lines:
        line.Delete()
    for line in lines:
        line.Delete()
    # print(max_length, Line_Startpoint, Line_Endpoint)
    # 返回lines
    return max_length, Line_Startpoint[0], Line_Endpoint[0]


def SelectOnScreen (acadapp, txt):
    acaddoc = acadapp.ActiveDocument
    acadmsp = acaddoc.ModelSpace
    acadulty = acaddoc.Utility
    #在CAD中删除名为ss1的选择集，如果失败的话，打印信息，然后创建一个名为ss1的选择集并赋值给slt
    try:
        acaddoc.SelectionSets.Item("SS1").Delete()
    except:
        print("Delete selection failed")
    slt = acaddoc.SelectionSets.Add("SS1")#创建选择集
    #在CAD中会出现该提示行
    acadulty.Prompt(txt)#提示对话框
    #执行操作后用户可以在CAD中进行对象捕捉
    slt.SelectOnScreen()
    return acaddoc,acadmsp,slt


#对ss选择集中的图元分别获取图元的句柄
def ss_to_lst(ss):
    lst = []
    if ss:
        for i in range(ss.Count):
            lst.append(ss.Item(i).handle)#得到选择集中每个图元的句柄
    return lst


# 坐标
def vtpnt(point):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, point)


# 浮点
def vtfloat(val):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, val)


# 整型
def vtint(val):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, val)


# 变体
def vtvariant(var):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, var)


# 坐标
def vtpnt2(pt):
    #坐标点转换为浮点数
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (pt[0], pt[1], pt[2]))


# 对象
def vtobj(obj):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, obj)


# 在cad中插入文字
def insert_TXT_incad(msp, insert_point, text_string, text_height):
    text_str = str(text_string)
    try:
        text_obj = msp.AddText(text_str, insert_point, text_height)
        text_obj.Layer = "Defpoints"
        return text_obj
    except Exception as e:
        print(f"插入文字时发生错误: {e}")
        return None
    

# 查询对应块的扩展属性
def get_block_reference_xdata(Block):
    # Block_Entity = Block[0]
    try:
        xdata_app_name = "SZQS"
        xdata = Block.GetXData(xdata_app_name)
        if xdata is not None:
            return xdata
        else:
            return None
    except Exception as e:
        print(f"查询扩展属性时发生错误: {e}")
        return None
    

#从多段线中获取端点，对于圆弧加一个中点
def getptsfromplwithcricle(curve_object, accuracy): 
    vertices = []
    a = curve_object.Coordinates#得到多段线的所有端点和中点
    elevation = round(curve_object.Elevation,accuracy)#round函数用于将一个浮点数四舍五入为最接近的整数，或将其精确到指定的小数位数。
    coords = [[round(a[x],accuracy),round(a[x+1],accuracy),elevation] for x in range(0, len(a), 2)]#每个点的x,y,z坐标
    vertices_num = len(coords)
    # 遍历多段线的顶点
    for j in range(vertices_num):
        vertei0 = coords[j]
        if j == vertices_num - 1:
            vertei1 = coords[0]
        else:
            vertei1 = coords[j+1]
        bulge = curve_object.GetBulge(j)
        # 判断是否为圆弧段
        if bulge != 0.0:
            # 获取圆弧的中点坐标
            arc_center = midptcoordonarc(vertei0 , vertei1 , bulge , accuracy)
            vertices = vertices + [vertei0] + [arc_center + [elevation]]
        else:
            vertices = vertices + [vertei0]
    return vertices


def check_relationship(polyline1, polyline2):
    # 获取多段线的顶点
    vertices1 = getptsfromplwithcricle(polyline1, 1)
    vertices2 = getptsfromplwithcricle(polyline2, 1)
    # 检查第一个多段线的顶点是否在第二个多段线内部
    inside_count1 = isPolylineWithinPolyline(vertices1, vertices2)
    # print(inside_count1)
    inside_count2 = isPolylineWithinPolyline(vertices2, vertices1)
    # print(inside_count2)
    if inside_count1 == True and inside_count2 == False:
        # print("图元2包含图元1")
        # 说明1是内轮廓，2是外轮廓
        return "2包含1"
    elif inside_count1 == False and inside_count2 == True:
        # print("图元1包含图元2")
        # 说明2是内轮廓，1是外轮廓
        return "1包含2"
    elif inside_count1 == False and inside_count2 == False:
        # print("两图元为相离关系")
        # 说明1和2相离
        return "相离"
    

# 插入图块
def insert_block_incad_with_block_name(msp, insert_point, block_name, ex_data, Scale_Factor = [1,1,1], RO_Angle = 0):
    # 比例因子（x, y, z）
    # Scale_Factor = [1, 1, 1]
    # 创建对应截面类型的块
    block_ref = msp.InsertBlock(insert_point, block_name, Scale_Factor[0], Scale_Factor[1], Scale_Factor[2], RO_Angle)
    # 扩展属性的对应应用程序名
    app_name = "SZQS"  
    # 扩展属性组码，1001代表应用名称，1000代表字符串数据，数据类型均为变体
    data_type = vtint([1001, 1000])
    data_lst = vtvariant([app_name, ex_data])
    # 添加扩展属性
    block_ref.SetXData(data_type, data_lst)
    # 检查扩展属性
    # xdata = get_block_reference_xdata(block_ref)
    # print(xdata)


def Load_Font(aDoc, Font_Name, Font_Height, Font_Width):
    '''
    在 aDoc 中加载或更新字体名 Font_Name, 字体高度 Font_Height, 字体宽度 Font_Width
    '''
    # 获取文本样式集合
    text_styles = aDoc.TextStyles
    # 检查是否已存在字体样式
    style_exists = False
    for style in text_styles:
        if style.Name == Font_Name:
            style_exists = True
            existing_style = style
            break
    if style_exists:
        # 如果样式已存在，则修改它
        existing_style.SetFont("宋体", False, False, 0, 0)
        existing_style.Height = Font_Height
        existing_style.Width = Font_Width
        print('已更新现有字体样式{}'.format(Font_Name))
    else:
        # 创建新的文本样式
        new_style = text_styles.Add(Font_Name)
        # 设置字体为宋体，参数分别为：字体名、粗体、斜体、字符集、字体内码
        new_style.SetFont("宋体", False, False, 0, 0)
        new_style.Height = Font_Height  # 设置字体高度
        new_style.Width = Font_Width   # 设置宽度因子
        print('已创建新字体样式{}'.format(Font_Name))
    # 设置为当前文本样式
    aDoc.ActiveTextStyle = new_style if not style_exists else existing_style


def Load_Standrad_Style(aDoc, DIMSCALE, DimStyle_name, DimStyle_font):
    """
    DIMSCALE 全局比例; DimStyle_name 标注样式名称; DimStyle_font 标注的文字样式
    """
    # 创建标注样式
    DimStyleObj = aDoc.DimStyles.Add(DimStyle_name)  
    print(DimStyleObj.Name)
    aDoc.ActiveDimStyle = aDoc.DimStyles.Item(DimStyle_name) # 置为当前
    # 设置标注变量
    aDoc.SetVariable("DIMSCALE", DIMSCALE) # 先设置全局比例因子为1
    aDoc.SetVariable("DIMDLE", 0)  # 当箭头使用倾斜、建筑标记、积分和无标记时尺寸线超过尺寸界线的距离
    aDoc.SetVariable("DIMDLI", 3.75) # 设定尺寸线之间的距离
    aDoc.SetVariable("DIMEXE", 1.25) # 指定尺寸界线超出尺寸线的距离
    aDoc.SetVariable("DIMEXO", 0.625) # 设定自图形中定义标注的点到尺寸界线的偏移距离。
    aDoc.SetVariable("DIMBLK", "")  # 尺寸箭头类型, 实心闭合箭头
    aDoc.SetVariable("DIMBLK1", "") 
    aDoc.SetVariable("DIMBLK2", "")
    aDoc.SetVariable("DIMLDRBLK", "") # 引线箭头类型, ""表示实心闭合
    aDoc.SetVariable("DIMTSZ", 0) # 设置为0，表示使用箭头
    aDoc.SetVariable("DIMASZ", 1.5) # 箭头大小
    aDoc.SetVariable("DIMTXSTY", DimStyle_font) # 文字样式
    aDoc.SetVariable("DIMTXT", 2) # 文字高度
    aDoc.SetVariable("DIMDEC", 0) # 小数位数
    aDoc.SetVariable("DIMTAD", 1) # 垂直位置——上
    aDoc.SetVariable("DIMTIH", 0) # 文字在尺寸线内的文字对齐方式为与尺寸线对齐
    # aDoc.SetVariable("DIMTOH", 1) # 文字在尺寸线外的文字对齐方式为水平
    aDoc.SetVariable("DIMTOH", 0) # 文字在尺寸线外的文字对齐方式为与尺寸线对齐
    aDoc.SetVariable("DIMCEN", 2) # 圆或圆弧圆心标记以及中心线的绘制大小, 0表示不绘制, <0为绘制中心线
    aDoc.SetVariable("DIMTMOVE", 1) # 在移动标注文字时添加一条引线
    aDoc.SetVariable("DIMCLRD", 256) # 尺寸线颜色 ByLayer
    aDoc.SetVariable("DIMLTYPE", 'BYLAYER') # 尺寸线线型 ByLayer
    aDoc.SetVariable("DIMLWD", -1) # 尺寸线线宽 ByLayer
    aDoc.SetVariable("DIMCLRE", 256) # 尺寸界线颜色 ByLayer
    aDoc.SetVariable("DIMLTEX1", 'BYLAYER') # 尺寸界线线型 ByLayer
    aDoc.SetVariable("DIMLTEX2", 'BYLAYER') # 尺寸界线线型 ByLayer
    aDoc.SetVariable("DIMCLRT", 4) # 文字颜色
    aDoc.SetVariable("DIMGAP", 1) # 从尺寸线偏移
    # 保存当前设置到标注样式
    try:
        aDoc.ActiveDimStyle.CopyFrom(aDoc)
        print(f"标注样式{DimStyle_name}设置已保存")
    except Exception as e:
        print(f"保存标注样式时出错: {e}")


# 插入直线段
def make_line(mspace, ptlst, layer):
    # 绘制多段线
    line = _com_retry(mspace.AddLine, vtpnt(ptlst[0]), vtpnt(ptlst[1]))
    # 设置多段线属性
    try:
        _com_retry(setattr, line, 'Layer', layer)
    except:
        _com_retry(setattr, line, 'Layer', '0')
    return line # 返回Obj


# 插入多段线
def make_polyline(mspace, ptlst, layer, width = 0):
    # 绘制多段线
    polyline = _com_retry(mspace.AddPolyline, vtpnt(ptlst))
    # 设置多段线属性
    try:
        _com_retry(setattr, polyline, 'Layer', layer)
    except:
        _com_retry(setattr, polyline, 'Layer', '0')
    _com_retry(setattr, polyline, 'ConstantWidth', width)
    return polyline # 返回Obj


def Add_Annotation_Linear(aMSp, pt1, pt2, pt3, angle, layer = '6标注线'):
    # 线性标注
    XLine1Point = vtpnt(pt1) # 起点
    XLine2Point = vtpnt(pt2) # 终点
    DimLineLocation = vtpnt(pt3) # 尺寸线经过该点
    RotationAngle = math.radians(angle) # 角度 0 水平 90 竖直
    dimRotObj = aMSp.AddDimRotated(XLine1Point, XLine2Point, DimLineLocation, RotationAngle)
    try:
        dimRotObj.Layer = layer
    except:
        dimRotObj.Layer = '0'


def Add_Annotation_Alignment(aMSp, pt1, pt2, pt3, layer = '6标注线'):
    # 对齐标注
    ExtLine1Point = vtpnt(pt1) # 起点
    ExtLine2Point = vtpnt(pt2) # 终点
    TextPosition = vtpnt(pt3) # 尺寸线经过该点
    dimAliObj = aMSp.AddDimAligned(ExtLine1Point, ExtLine2Point, TextPosition)
    try:
        dimAliObj.Layer = layer
    except:
        dimAliObj.Layer = '0' 


def Add_Annotation_Angle(aMSp, pt1, pt2, pt3, pt4, layer = '6标注线'):
    # 角度标注  
    AngleVertex = vtpnt(pt1) # 顶点(圆心)
    FirstEndPoint = vtpnt(pt2) # 起点
    SecondEndPoint = vtpnt(pt3) # 终点
    TextPoint = vtpnt(pt4) # 尺寸线经过该点
    dimAngObj = aMSp.AddDimAngular(AngleVertex, FirstEndPoint, SecondEndPoint, TextPoint)
    try:
        dimAngObj.Layer = layer
    except:
        dimAngObj.Layer = '0' 


def Add_Annotation_Arc(aMSp, pt1, pt2, pt3, pt4, layer = '6标注线'):
    # 弧长标注
    ArcCenter = vtpnt(pt1) # 顶点(圆心)
    FirstEndPoint = vtpnt(pt2) # 起点
    SecondEndPoint = vtpnt(pt3) # 终点
    ArcPoint = vtpnt(pt4) # 尺寸线经过该点
    dimArcObj = aMSp.AddDimArc(ArcCenter, FirstEndPoint, SecondEndPoint, ArcPoint)
    try:
        dimArcObj.Layer = layer
    except:
        dimArcObj.Layer = '0' 


def Add_Annotation_Diameter(aMSp, pt1, pt2, length, layer = '6标注线'):
    # 直径标注
    ChordPoint = vtpnt(pt1) # 圆上任一直径的端点
    FarChordPoint = vtpnt(pt2) # 直径的另一个端点
    LeaderLength = length # 引线长度, 为点ChordPoint到标准文字定位夹点的距离
    dimDiaObj = aMSp.AddDimDiametric(ChordPoint, FarChordPoint, LeaderLength)
    try:
        dimDiaObj.Layer = layer
    except:
        dimDiaObj.Layer = '0' 


def Add_Annotation_Radius(aMSp, pt1, pt2, length, layer = '6标注线'):
    # 半径标注
    Center = vtpnt(pt1) # 圆心
    ChordPoint = vtpnt(pt2) # 圆上一点
    LeaderLength = length # 引线长度, 为点ChordPoint到标准文字定位夹点的距离
    dimRadObj = aMSp.AddDimRadial(Center, ChordPoint, LeaderLength)
    try:
        dimRadObj.Layer = layer
    except:
        dimRadObj.Layer = '0' 


# def Add_Annotation_MLeader(aCAD, aMSp, ScaleFactor, start_point, end_point, line_layer_name, Font_name, text_content):
#     """
#     aMSp 模型空间; ScaleFactor 全局比例; start_point 引线起点; end_point 引线终点; line_layer_name 标注图层; Font_name 文本字体样式; text_content 文本内容
#     """
#     # 创建点数组
#     points_list = [start_point[0], start_point[1], start_point[2], end_point[0], end_point[1], end_point[2]]
#     # 转换为VARIANT格式
#     points_var = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, points_list)
#     # 创建多重引线 - 检查返回值类型
#     MLeaderObj = aMSp.AddMLeader(points_var, 0)[0] # 返回一个元组, (<MLeaderObj>, 0)
#     # 设置引线属性
#     try:
#         MLeaderObj.Layer = line_layer_name
#     except:
#         MLeaderObj.Layer = '0'
#     MLeaderObj.ScaleFactor = ScaleFactor # 比例
#     MLeaderObj.TextStyleName = Font_name # 文字样式
#     MLeaderObj.TextString = text_content # 文字内容
#     MLeaderObj.TextHeight = 2 # 文字高度
#     MLeaderObj.TextLeftAttachmentType = 7 # 引线连接——左
#     MLeaderObj.TextRightAttachmentType = 7 # 引线连接——右
#     MLeaderObj.ArrowheadSize = 1.5 # 箭头大小
#     MLeaderObj.DoglegLength = 0 # 基线长度
#     MLeaderObj.LandingGap = 3 # 基线间隙
#     version = aCAD.Version[:2]  # 当前CAD的版本号
#     color = aCAD.GetInterfaceObject("AutoCAD.AcCmColor.%s" % version) # 真彩色和CAD版本号有关
#     color.SetRGB(0, 255, 0) # 绿色
#     MLeaderObj.LeaderLineColor = color # 线的颜色
def Add_Annotation_MLeader(aCAD, aMSp, ScaleFactor, start_point, end_point, line_layer_name, Font_name, text_content):
    """
    aMSp 模型空间; ScaleFactor 全局比例; start_point 引线起点; end_point 引线终点; line_layer_name 标注图层; Font_name 文本字体样式; text_content 文本内容
    """
    # 判断方向
    dx = end_point[0] - start_point[0]
    # 创建点数组（始终保持起点在前、终点在后，箭头指向起点）
    points_list = [start_point[0], start_point[1], start_point[2], end_point[0], end_point[1], end_point[2]]
    # 转换为VARIANT格式
    points_var = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, points_list)
    # 创建多重引线 - 检查返回值类型
    MLeaderObj = aMSp.AddMLeader(points_var, 0)[0] # 返回一个元组, (<MLeaderObj>, 0)
    # 设置引线属性
    try:
        MLeaderObj.Layer = line_layer_name
    except:
        MLeaderObj.Layer = '0'
    MLeaderObj.ScaleFactor = ScaleFactor # 比例
    MLeaderObj.TextStyleName = Font_name # 文字样式
    MLeaderObj.TextString = text_content # 文字内容
    MLeaderObj.TextHeight = 2 # 文字高度
    MLeaderObj.TextLeftAttachmentType = 7 # 引线连接——左
    MLeaderObj.TextRightAttachmentType = 7 # 引线连接——右
    MLeaderObj.ArrowheadSize = 1.5 # 箭头大小
    MLeaderObj.DoglegLength = 0 # 基线长度
    MLeaderObj.LandingGap = 3 # 基线间隙

    # 当终点在起点左侧时，翻转dogleg方向，再恢复引线终点位置
    if dx < 0:
        direction = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [-1.0, 0.0, 0.0])
        MLeaderObj.SetDoglegDirection(0, direction)
        # SetDoglegDirection会改变引线终点，需要恢复
        verts = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,
            [start_point[0], start_point[1], start_point[2], end_point[0], end_point[1], end_point[2]])
        MLeaderObj.SetLeaderLineVertices(0, verts)
    version = aCAD.Version[:2]  # 当前CAD的版本号
    color = aCAD.GetInterfaceObject("AutoCAD.AcCmColor.%s" % version) # 真彩色和CAD版本号有关
    color.SetRGB(0, 255, 0) # 绿色
    MLeaderObj.LeaderLineColor = color # 线的颜色


def get_text_length_height(text_obj):
    """获取文本对象的长度"""
    point1, point2 = text_obj.GetBoundingBox()
    # 计算长度和高度
    length = round(abs(point1[0] - point2[0]),1)
    height = round(abs(point1[1] - point2[1]),1)
    return length, height


def Add_Hatch(aMSp, ptnName, Obj, layer, scale, ptnType = 0, bAss = True):
    '''
    ptnName 填充类型, Obj 填充对象, scale 填充比例
    '''
    # ptnName, ptnType, bAss = "ANSI31", 0, True
    outerLoop = []
    outerLoop.append(Obj)
    outerLoop = vtobj(outerLoop)
    hatchObj = aMSp.AddHatch(ptnType, ptnName, bAss)
    hatchObj.AppendOuterLoop(outerLoop)
    hatchObj.Evaluate()  # 进行填充计算，使图案吻合于边界
    hatchObj.PatternScale = scale  # 设置填充图案比例, 1.0
    try:
        hatchObj.Layer = layer
    except:
        hatchObj.Layer = '0'


def Add_Text(aMSp, insert_point, text_string, height, font_name, layer, AlignNum = 12):
    '''
    insert_point 文本插入点, text_string 文本内容, height 文本高度, AlignNum 对齐方式
    AlignNum = 6 # 左上为定位点
    AlignNum = 7 # 中上为定位点
    AlignNum = 8 # 右上为定位点

    AlignNum = 3 # 左下为定位点
    AlignNum = 1 # 中下为定位点
    AlignNum = 2 # 右下为定位点

    AlignNum = 9 # 左中为定位点
    AlignNum = 4 # 居中为定位点
    AlignNum = 10 # 居中为定位点, 好像效果和4完全一样?
    AlignNum = 11 # 右中为定位点

    AlignNum = 12 # 左底为定位点, 字体相对于定位点会抬升一点
    AlignNum = 13 # 中底为定位点, 字体相对于定位点会抬升一点
    AlignNum = 14 # 右底为定位点, 字体相对于定位点会抬升一点
    '''
    # 创建文本
    insertPnt = vtpnt(insert_point)
    textObj = _com_retry(aMSp.AddText, text_string, insertPnt, height)
    try:
        _com_retry(setattr, textObj, 'TextStyleName', font_name)
    except:
        pass
    try:
        _com_retry(setattr, textObj, 'Layer', layer)
    except:
        _com_retry(setattr, textObj, 'Layer', '0')
    # 文本对齐方式
    _com_retry(setattr, textObj, 'Alignment', AlignNum)
    _com_retry(setattr, textObj, 'TextAlignmentPoint', insertPnt)

    return textObj


def Add_MText(aMSp, insert_point, text_lst, text_height, font_name, layer):
    '''
    插入多行文本
    '''
    # 创建文本
    insertPnt = vtpnt(insert_point)
    text_string = '\n'.join(text_lst)
    text_width = 0 # 定义文本宽度 (设置为0表示不限制宽度，按一行显示；设置一个正数用于控制换行宽度)
    MtextObj = _com_retry(aMSp.AddMText, insertPnt, text_width, text_string)
    _com_retry(setattr, MtextObj, 'Height', text_height)
    try:
        _com_retry(setattr, MtextObj, 'TextStyleName', font_name)
    except:
        pass
    try:
        _com_retry(setattr, MtextObj, 'Layer', layer)
    except:
        _com_retry(setattr, MtextObj, 'Layer', '0')
    return MtextObj


def Add_Header(aMSp, Global_Scale, insert_point, text_string, font_name):
    """
    aMSp 模型空间; Global_Scale 全局比例; insert_point 插入点; text_string 文本内容; font_name 文本字体样式
    """
    # 添加文本
    textObj = Add_Text(aMSp, insert_point, text_string, 4.0*Global_Scale, font_name, '1粗实线', 4)
    # 获取文本长度
    textLength, textHeight = get_text_length_height(textObj)
    # 添加下划线
    ptlst1 = [(insert_point[0]-textLength/2-Global_Scale, insert_point[1]-textHeight/2-Global_Scale, 0), (insert_point[0]+textLength/2+Global_Scale, insert_point[1]-textHeight/2-Global_Scale, 0)]
    ptlst2 = [(x, y-Global_Scale, z) for x,y,z in ptlst1]
    make_line(aMSp, ptlst1, '1粗实线')
    make_line(aMSp, ptlst2, '2细实线')


def Add_Elevation_Symbol(aMSp, Global_Scale, insert_point, text_string, font_name, param1 = 1, param2 = 1):
    """
    aMSp 模型空间; Global_Scale 全局比例; insert_point 插入点; text_string 文本内容; font_name 文本字体样式; param 向量[1, 1], 分别表示左右和上下
    添加标高
    """
    x,y,z = insert_point
    nx, ny = param1, param2
    # 添加下划线
    ptlst1 = [(x-2.5*Global_Scale*nx, y, 0), (x+4.0*Global_Scale*nx, y, 0)]
    make_line(aMSp, ptlst1, '6标注线')
    # 添加文本
    
    text_justfy = {'1,1':[12, 0.5], '-1,1':[14, 0.5], '-1,-1':[8, 1.0], '1,-1':[6, 1.0]} # 对应不同 param 时文本的对齐方式以及文本额外移动的距离比例
    AlignNum, i = text_justfy[f'{param1},{param2}']
    text_insert_point = (x-2.5*Global_Scale*nx, y+(2.5+i)*Global_Scale*ny, 0)
    textObj = Add_Text(aMSp, text_insert_point, text_string, 2.5*Global_Scale, font_name, '5文本', AlignNum)
    # 获取文本长度
    textLength, textHeight = get_text_length_height(textObj)
    # 添加多段线
    ptlst2 = [(x+2.5*Global_Scale*nx, y+2.5*Global_Scale*ny, 0), (x, y, 0), (x-2.5*Global_Scale*nx, y+2.5*Global_Scale*ny, 0), (x+(textLength-2.5*Global_Scale)*nx, y+2.5*Global_Scale*ny, 0, 0)]
    ptlst2 = pointlist_extend(ptlst2)
    make_polyline(aMSp, ptlst2, '6标注线')