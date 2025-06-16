import os
from core import Pic_Handle as pHto
from core import Para_Handle as pGto
from config import eazy_setting as Es
from tips import utils
import io


if __name__=="__main__":

    #docx的标题初始位置确认
    Myconfig=Es.load_config()
    '''
    print("原docx文件标题在第",Myconfig["head_pos"]+1,"行")
    print("原docx文件标题共",Myconfig["head_pos_end"],"行")
    print("模板docx文件标题在第",Myconfig["head_new_pos"]+1,"行")
    '''
    
    docx_path=input("输入处理的工单: ").strip('"\'')
    
    if os.path.exists(docx_path):
        print("====找到文件====")
##        if docx_path.lower().endswith('.doc'):
##            data=utils.To_docx(docx_path)
##        else:
        with open(docx_path, 'rb') as f:
            data = f.read()
                
        doc,new_doc,work_order_number=pGto.handle_docx(io.BytesIO(data),Myconfig)
        pHto.Pics_handle(doc,io.BytesIO(data),new_doc)    
        new_doc.save(f"{work_order_number}.docx")
    else:
        print("文件不存在")
    
    
    
