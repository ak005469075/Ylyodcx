def To_docx(doc_path):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # 无界面模式

    try:
        # 1. 用Word打开.doc文件并另存为.docx到内存
        doc = word.Documents.Open(os.path.abspath(doc_path))
        stream = io.BytesIO()
        doc.SaveAs(stream, FileFormat=16)  # 16 = wdFormatDOCX
        doc.Close()
        
        # 2. 从内存加载到docx.Document对象
        docx_obj = Document(stream)
        

        #remove_konghang(docx_obj)
        
        return docx_obj  # 直接返回Document对象

    except Exception as e:
        raise RuntimeError(f"转换失败: {e}")
    finally:
        word.Quit()


def remove_konghang(doc):
    paras_move=[]
    for para in doc.paragraphs:
        if not para.text.strip():
            paras_move.append(para)
    for para in reversed(paras_move):
        p=para._element
        p.getparent().remove(p)
        p._p=p._element=None
    return doc