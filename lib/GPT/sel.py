import win32com.client as win32
from win32com.client import constants
import os
import time
import pygetwindow as gw
import pyautogui

# Get the current directory of this script
current_directory = os.path.dirname(os.path.abspath(__file__))
def seg(title="mor"):
    # 打开Word应用程序
    doc_app = win32.gencache.EnsureDispatch('Word.Application')
    doc_app.Visible = True

    # 打开一个已经存在的文档
    file_path = os.path.join(current_directory, 'output\debug\output.docx')
    doc = doc_app.Documents.Open(file_path)

    # 添加延迟，确保Word处理完毕
    time.sleep(3)

    # 将Word置于前台
    word_windows = gw.getWindowsWithTitle('')
    for win in word_windows:
        if "Word" in win.title or ".docx" in win.title:
            win.activate()
            break
    time.sleep(2)
    # 模拟按下回车键来关闭可能的弹窗
    pyautogui.click()
    time.sleep(1)
    pyautogui.press('enter')
    # 假设您已经知道在哪里插入分节符来创建第二部分和第三部分
    time.sleep(2)



    # 为第二部分设置页码
    footer2 = doc.Sections(2).Footers(constants.wdHeaderFooterPrimary)
    footer2.LinkToPrevious = False
    footer2.PageNumbers.Add(PageNumberAlignment=constants.wdAlignPageNumberCenter, FirstPage=True)
    footer2.PageNumbers.RestartNumberingAtSection = True
    footer2.PageNumbers.StartingNumber = 1
    # 设置页码为罗马数字
    footer2.PageNumbers.NumberStyle = constants.wdPageNumberStyleUppercaseRoman

    # 为第三部分设置页码
    footer3 = doc.Sections(3).Footers(constants.wdHeaderFooterPrimary)
    footer3.LinkToPrevious = False
    footer3.PageNumbers.Add(PageNumberAlignment=constants.wdAlignPageNumberCenter, FirstPage=True)
    footer3.PageNumbers.RestartNumberingAtSection = True
    footer3.PageNumbers.StartingNumber = 1

    # 定位到整个文档的第三页开始位置
    doc_app.Selection.GoTo(What=constants.wdGoToPage, Which=constants.wdGoToAbsolute, Count=3)
    start_of_page = doc_app.Selection.Start

    # 找到“目录”标题的结束位置
    doc_app.Selection.Find.Execute(FindText="目录")
    end_of_title = doc_app.Selection.End

    # 将目录插入到“目录”标题之后
    range_for_toc = doc.Range(Start=end_of_title, End=end_of_title)

    # 在标题后插入目录
    toc = doc.TablesOfContents.Add(range_for_toc,
                                   RightAlignPageNumbers=True,
                                   UseHeadingStyles=True,
                                   UpperHeadingLevel=1,
                                   LowerHeadingLevel=3,
                                   IncludePageNumbers=True,
                                   AddedStyles="")

    # 更新目录，以确保它反映了文档中的所有当前标题
    toc.Update()

    # 调整目录内容的段落间距
    toc_range = toc.Range
    toc_range.ParagraphFormat.SpaceBefore = 0  # 设置段前间距为0磅
    toc_range.ParagraphFormat.SpaceAfter = 0  # 设置段后间距为0磅

    # 保存文档
    # doc.Save()

    # 保存文档为title.docx
    doc.SaveAs(os.path.join(os.getcwd(), r"output\\"+title+".docx"))
    time.sleep(1)
    doc.Close()
    doc_app.Quit()
if __name__ =="__main__":
    seg()