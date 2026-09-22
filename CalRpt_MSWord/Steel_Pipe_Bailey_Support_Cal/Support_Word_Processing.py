# 1. 标准库
import os

# 2. 第三方库
import win32com.client


# 打开计算书样板并另存为
def open_doc_saveas_adoc(wdApp, doc_name, new_doc_name, doc_open_path, doc_save_path):
    # 打开文件
    wdApp.Visible = 0 # 后台
    # wdApp.Visible = 1 # 前台
    doc_file_path = os.path.join(doc_open_path, doc_name)
    Doc = wdApp.Documents.Open(doc_file_path) # 打开Word文件
    # Doc.Activate() # 激活到最前面
    # 另存为
    new_doc_file_path = os.path.join(doc_save_path, new_doc_name)
    new_doc_file_path = new_doc_file_path.replace('/', '\\')
    print(new_doc_file_path)
    Doc.SaveAs(new_doc_file_path)
    Doc.Close()
    # 打开另存为的文件
    aDoc = wdApp.Documents.Open(new_doc_file_path) # 打开Word文件
    # aDoc.Activate() # 激活到最前面
    return aDoc


# 关闭并保存word文件
def adoc_save_close(aDoc):
    aDoc.Save()
    aDoc.Close()


def repl_docstr_with_strlst(doc, orign, new):
    # search_range = doc.Content
    # search_range.Find.Execute(FindText=orign, ReplaceWith=new)
    try:
        result = doc.Content.Find.Execute(
                FindText=orign,        # 要查找的文本
                ReplaceWith=new,     # 替换为的文本
                Replace=2,                # wdReplaceAll = 2 (全部替换)
                Forward=True,             # 向前搜索
                Wrap=1,                   # wdFindContinue = 1 (继续搜索)
                Format=False,             # 不匹配格式
                MatchCase=False,          # 不区分大小写
                MatchWholeWord=False,     # 不匹配整个单词
                MatchWildcards=False,     # 不使用通配符
                MatchSoundsLike=False,    # 不匹配发音
                MatchAllWordForms=False   # 不匹配所有词形
            )
    except:
        print(f"文本{orign}未替换")


# 特殊文本替换为对应的公式
def repl_docstr_with_omath(doc, select, latex, txt):
    try:
        rng = doc.Range()
        rng.Find.IgnoreSpace = True
        ret = rng.Find.Execute(txt)
        formula = select.OMaths.Add(Range=rng)
        rng.Text = latex
        rng.OMaths(1).ConvertToMathText()
        rng.OMaths.BuildUp()
    except:
        print(f"公式{txt}未替换")


# 特殊文本替换为对应的图片(无需额外执行删除文字命令)
def repl_docstr_with_pic(doc,pic_str,pic_path): 
    try:
        search_range = doc.Content
        search_range.Find.Execute(pic_str)
        doc.Application.Selection.SetRange(search_range.Start, search_range.End)
        doc.Application.Selection.InlineShapes.AddPicture(pic_path)
    except:
        print(f"图片{pic_str}未替换")


# 将计算书的一部分粘贴至总计算书中
def WordA_PasteTo_WordB(word, file_path1, file_path2, FindText1, FindText2, FindText3):
    # 打开源文档
    doc = word.Documents.Open(file_path1)

    # 获取源文档中要复制的内容范围
    search_range = doc.Content
    search_range.Find.Execute(FindText=FindText1)
    search_range.Select()
    start = word.Selection.End

    search_range = doc.Content
    search_range.Find.Execute(FindText=FindText2)
    search_range.Select()
    end = word.Selection.Start

    # 复制选定内容
    doc.Range(start, end).Copy()

    # 打开目标文档
    doc_new = word.Documents.Open(file_path2)

    # 在目标文档中查找"粘贴在这里"的位置
    search_range = doc_new.Content
    if search_range.Find.Execute(FindText=FindText3):
        # 选中找到的内容
        search_range.Select()
        # 删除"粘贴在这里"这个标记文本
        word.Selection.Text = ""
        # 在当前位置粘贴
        # word.Selection.Paste()
        word.Selection.PasteSpecial(
                Link=False, 
                DataType=win32com.client.constants.wdPasteRTF  # 保留所有格式
            )
    else:
        print(f"未找到{FindText3}标记")

    # 关闭两个文件并保存更改
    doc_new.Close(SaveChanges=True)
    doc.Close(SaveChanges=False)


# 检索关键词复制粘贴(去空白行)
def WordA_PasteTo_WordB_DeleteLines(word, doc1, doc2, FindText1, FindText2, FindText3):
    try:
        # 打开源文档
        # doc = word.Documents.Open(file_path1)
        # 获取源文档中要复制的内容的起点
        search_range = doc1.Content
        search_range.Find.Execute(FindText=FindText1)
        search_range.Select()
        start = word.Selection.End
        # 去除换行符
        if doc1.Range(start, start+1).Text == "\r":
                start += 1
        # 获取源文档中要复制的内容的终点
        search_range = doc1.Content
        search_range.Find.Execute(FindText=FindText2)
        search_range.Select()
        end = word.Selection.Start
        # 复制选定内容
        doc1.Range(start, end).Copy()
        # 打开目标文档
        # doc_new = word.Documents.Open(file_path2)
        # 在目标文档中查找"粘贴在这里"的位置
        search_range = doc2.Content
        if search_range.Find.Execute(FindText=FindText3):
            # 选中找到的内容
            search_range.Select()
            # 删除"粘贴在这里"这个标记文本所在行
            Delete_Word_FindText_Paragraphs(doc2, FindText3)
            # 在当前位置粘贴
            word.Selection.Paste()
        else:
            print(f"未找到{FindText3}标记")
    except:
        print(f"计算书模板{FindText3}未粘贴")


# 根据检索字符串'：'在指定位置换行
def str_join_n(data):
    result = []
    for item in data:
        # 提取字符串并分割标题和内容
        text = item[0]
        title, content = text.split('：', 1)  # 只分割第一个冒号
        # 格式化输出
        formatted = f"{title}：\n{content}"
        result.append(formatted)

    # 合并结果并用换行分隔
    output = '\n'.join(result)
    # print(output)
    return output


"""删除包含指定字符串的整行"""
def Delete_Word_FindText_Paragraphs(doc, search_text):
    try:
        # 遍历所有段落
        for para in doc.Paragraphs:
            # 检查段落是否包含目标字符串
            if search_text in para.Range.Text:
                # 删除整个段落（包括换行符）
                para.Range.Delete()
                break  # 只需删除第一个匹配项，使用break
    except:
        print(f"{search_text}未删除")


# 规范代号对照库, 防止中文识别出现问题, 用数字代号替代
def Standard_Codename():
    Standard_Codename_dict = {
        "自定义": "", 
        "公路桥梁抗风设计规范": "《公路桥梁抗风设计规范》（JTG／T-3360-01-2018）", 
        "工程结构通用规范": "《工程结构通用规范》（GB 55001-2021）", 
        "港口工程荷载规范": "《港口工程荷载规范》（JTS 144-1-2010）", 
        "建筑桩基技术规范": "《建筑桩基技术规范》（JGJ 94-2008）", 
        "建筑地基基础设计规范": "《建筑地基基础设计规范》（GB 50007-2011）", 
    }
    return Standard_Codename_dict


# 从 Midas UI 中获取的用于生成计算书的参数 标题
def FEA_To_Cal_txt_str_lookfor():
    str_lookfor_lst = ['*风荷载规范',
                        '*水流力规范', 
                        '*自重参数', 
                        '*混凝土荷载参数', 
                        '*模板荷载参数', 
                        '*施工荷载参数', 
                        '*风荷载参数', 
                        '*水流力参数', 
                        '*贝雷梁材料特性参数', 
                        '*钢材材料特性参数', 
                        '*混凝土材料特性参数', 
                        '*钢筋材料特性参数', 
                        '*螺栓材料特性参数', 
                        '*焊缝材料特性参数', 
                        '*竹胶板材料特性参数', 
                        '*方木材料特性参数', 
                    ]
    return str_lookfor_lst

#=====================================================================================================================================================================

