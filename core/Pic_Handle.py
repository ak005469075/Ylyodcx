import os,io
import re
from docx import Document
from docx.oxml.shared import qn
from docx.oxml import OxmlElement
from docx.shared import Inches,Emu
from docx.parts.image import ImagePart
import zipfile
from lxml import etree
from collections import OrderedDict

#By yLy
def has_pics(data):
    try:
        with zipfile.ZipFile(data) as z:
            # 快速检查media目录是否存在图片文件
            if any(f.startswith('word/media/') for f in z.namelist()):
                print("√检测到有图片!")
                return True
    except FileNotFoundError:
        print("文件错误!")
        return False

def get_image_size_from_raw_file(data):
   
    image_sizes = OrderedDict()
    
    with zipfile.ZipFile(data) as z:
        # 1. 读取document.xml
        with z.open('word/document.xml') as f:
            tree = etree.parse(f)
            ns = {
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
            }
            
            # 2. 查找所有包含图片引用的drawing元素
            drawings = tree.xpath('//w:drawing[.//a:blip/@r:embed]', namespaces=ns)
            print(f"找到 {len(drawings)} 个包含图片的drawing元素")

            for i, drawing in enumerate(drawings, 1):
                # 获取图片引用ID
                rId = drawing.xpath('.//a:blip/@r:embed', namespaces=ns)[0]
                
                # 提取尺寸（优先wp:extent，备选a:ext）
                extent = drawing.xpath('.//wp:extent', namespaces=ns)
                if extent:
                    width = int(extent[0].get('cx'))
                    height = int(extent[0].get('cy'))
                    image_sizes[rId] = (width, height)
                    #print(f"[{i}] rId={rId}: {width}x{height} EMU (来自wp:extent)")
                    continue
                
                ext = drawing.xpath('.//a:ext', namespaces=ns)
                if ext:
                    width = int(ext[0].get('cx'))
                    height = int(ext[0].get('cy'))
                    image_sizes[rId] = (width, height)
                    #print(f"[{i}] rId={rId}: {width}x{height} EMU (来自a:ext)")
    
    return image_sizes

def insert_pics(source_doc,data,target_doc):
  
    # 1. 获取所有图片尺寸映射（rId -> (width, height)）
    image_sizes = get_image_size_from_raw_file(data)
    
  
    for rId, (width, height) in image_sizes.items():
        # 2. 查找对应的图片关系
        rel = next(
            (r for r in source_doc.part.rels.values() 
             if getattr(r, 'rId', None) == rId),
            None
        )
        
        if rel is None or not isinstance(getattr(rel, 'target_part', None), ImagePart):
            print(f"⚠️ 跳过 {rId}（关系不存在或不是图片）")
            continue
            

        #print(image_sizes)
        #print(width,height)
            
        # 3. 插入图片
        try:
            target_doc.add_picture(
                        io.BytesIO(rel.target_part.blob),
                        width=Emu(width),
                        height=Emu(height)
                    )
            print(f"已插入图片 {rel.rId}，尺寸: {width}x{height} EMU")
        except Exception as e:
            print(f"⚠️ 插入图片 {rel.rId} 失败（使用默认尺寸）: {str(e)}")
            target_doc.add_picture(io.BytesIO(rel.target_part.blob), width=Inches(2.0))
        

def Pics_handle(doc,data,new_doc):
    if(has_pics(data)):
        insert_pics(doc,data,new_doc)
    
