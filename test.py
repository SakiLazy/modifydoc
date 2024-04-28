import win32com.client as win32


def add_footer_with_auto_numbering(doc_path):
    # 启动 Word 并打开文档
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(doc_path)
    word.Visible = True  # 设置为 True 可以看到操作过程，便于调试
    # 查找第一个“标题 1”并插入分节符
    found = False
    for paragraph in doc.Paragraphs:
        if paragraph.Style.NameLocal == '标题 1':  # 根据你的Word版本调整样式名称
            found = True
            # 在“标题 1”起始处插入分节符
            range = paragraph.Range
            range.Collapse(Direction=1)  # Collapse the range to its start
            range.InsertBreak(Type=win32.constants.wdSectionBreakNextPage)
            break
    if found:
        # 获取最新创建的分节
        section = doc.Sections.Last
        # 设置新分节的页脚
        footer = section.Footers(win32.constants.wdHeaderFooterPrimary)
        footer.LinkToPrevious = False  # 不与前一节的页脚链接
        # 插入格式化的页码字段，并确保破折号正确插入
        footer_text = footer.Range
        footer_text.Text = " - "  # 首先插入前破折号
        footer_text.Collapse(Direction=1)  # Collapse range to the end
        footer_text.Fields.Add(footer_text, win32.constants.wdFieldEmpty, r'PAGE \* Arabic \* MERGEFORMAT', True)
        footer_text.InsertAfter(" - ")  # 在页码后添加第二个破折号
        footer.Range.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphCenter
        # 保存更改
        doc.Save()
    else:
        print("没有找到 '标题 1'。")
    # 关闭文档和 Word 应用
    doc.Close()
    word.Quit()
# 替换为你的文档路径
add_footer_with_auto_numbering(r"C:\Users\sakil\Desktop\ddl\modifydoc\cyx.docx")
