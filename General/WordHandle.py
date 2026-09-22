# 1. 标准库
import os
from copy import deepcopy

# 2. 第三方库
from lxml import etree
import win32com.client
import latex2mathml.converter


# ── Office Math 命名空间 ──────────────────────────────────────
OMML_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
# 全局注册，让 lxml 序列化时自动输出 m: 前缀
etree.register_namespace('m', OMML_NS)


# 与word建立连接
def WordApp_Dispatch():
    # 建立连接
    # wdApp = EnsureDispatch("Word.Application") # 创建Word对象
    wdApp = win32com.client.Dispatch("Word.Application")
    # print(sys.modules[wdApp.__module__].__file__) # 显示COM缓存文件路径
    return wdApp


# 退出与word的连接
def WordApp_Quit(wdApp):
    wdApp.Quit()


# 公式替换
# 将 LaTeX 字符串转换为 Word OMML节点
def latex_to_word_omml(latex_input):
    # LaTeX -> MathML
    mathml = latex2mathml.converter.convert(latex_input)
    tree = etree.fromstring(mathml)

    # MathML -> OMML (需要 Office 自带的转换表)
    xslt_path = r'C:/Program Files/Microsoft Office/root/Office16/MML2OMML.XSL'
    if not os.path.exists(xslt_path):
        raise FileNotFoundError(f"未找到转换表，请检查路径: {xslt_path}")

    xslt = etree.parse(xslt_path)
    transform = etree.XSLT(xslt)
    new_dom = transform(tree)
    root = new_dom.getroot()

    # ── 关键修复 ────────────────────────────────────────────────
    # XSLT 输出的根元素可能将 m: 命名空间声明挂在文档层，
    # getroot() 后该声明不会自动跟随到节点本身。
    # 序列化再重新解析，强制把 xmlns:m="..." 写到元素自身，
    # 这样 append 到 python-docx 段落时命名空间不会丢失。
    xml_bytes = etree.tostring(root, encoding='utf-8', xml_declaration=False)
    return etree.fromstring(xml_bytes)


# 在指定占位符处插入公式
def replace_placeholder_with_formula(doc, placeholder, latex_str):
    """
    doc: Document 对象
    placeholder: 模板中的占位符文本，如 "{KLQ_GS}"
    latex_str: LaTeX 公式字符串
    """
    for paragraph in doc.paragraphs:
        if placeholder in paragraph.text:
            # 获取转换后的 Word 原生公式节点
            omml_node = latex_to_word_omml(latex_str)
            # 清除占位符文本
            # 保留段落原有的样式
            paragraph.text = paragraph.text.replace(placeholder, "")
            # ── 关键修复 ────────────────────────────────────────
            # 1. deepcopy 防止 lxml 的 append 从源树"搬移"节点时
            #    剥离已挂载的命名空间声明。
            # 2. 清除占位符后可能残留的空 run，避免 Word 渲染异常。
            paragraph._p.append(deepcopy(omml_node))
            # 移除占位符清除后留下的空 run（如果有）
            _remove_empty_runs(paragraph)
            print(f" {placeholder} 已成功替换为公式")
            break # 找到并替换后退出循环


def _remove_empty_runs(paragraph):
    """移除段落中不含可见文本的空 run，保留包含绘图/图片等的 run。"""
    w_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    for run in paragraph._p.findall(f'{{{w_ns}}}r'):
        # 如果 run 下没有 <w:t> 或 <w:t> 为空，且没有绘图子元素，则移除
        t_elem = run.find(f'{{{w_ns}}}t')
        drawing = run.find(f'{{{w_ns}}}drawing')
        pict = run.find(f'{{{w_ns}}}pict')
        if (t_elem is None or not t_elem.text) and drawing is None and pict is None:
            # 确保不是包含 rPr（格式）的空 run
            rpr = run.find(f'{{{w_ns}}}rPr')
            if rpr is None:
                paragraph._p.remove(run)
