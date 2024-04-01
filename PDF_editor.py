from os.path import split
import pdfplumber
from tkinter import *
from tkinter import ttk,filedialog
from PIL import Image,ImageTk

"""本文件为PDF_editor的主文件"""

def open_pdf_file():
    """弹窗选取PDF文件，返回1.png图片列表；2.文件所在绝对路径；3.文件名"""
    #设置不要弹出tk窗口
    root = Tk()
    root.withdraw()
    #弹窗选取文件，获取文件绝对路径
    file_path_str = filedialog.askopenfilename()
    #拆分文件所在的路径与文件名称
    folder_path_str,file_name_str = split(file_path_str)
    #适用pdfplumber转换PDF文件为image对象
    with pdfplumber.open(file_path_str) as pdf:
        images = list ()
        for index,page in enumerate(pdf.pages):
            #转换为图片，dpi设置为200
            image = page.to_image(resolution=200)
            images.append(image)
            # image.save('pictures'+'_'+ str(index)+'.png')
        
    return(images,folder_path_str,file_name_str)


if __name__ == '__main__':
    images, folder_path,file_name = open_pdf_file()
    print (len(images))
    print (folder_path)
    print (file_name)