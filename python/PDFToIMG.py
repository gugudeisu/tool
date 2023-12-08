import os
import sys
import fitz
from reportlab.lib.pagesizes import portrait
from reportlab.pdfgen import canvas
from PIL import Image


def pdf2img(filename=r'./document.pdf'):
    #  打开PDF文件，生成一个对象
    doc = fitz.open(filename)
    print("共",doc.pageCount,"页")
    for pg in range(doc.pageCount):
        print("\r转换为图片",pg+1,"/",doc.pageCount,end="")
        page = doc[pg]
        rotate = int(0)
        # 每个尺寸的缩放系数为8，这将为我们生成分辨率提高64倍的图像。
        zoom_x = 8.0
        zoom_y = 8.0
        trans = fitz.Matrix(zoom_x, zoom_y).preRotate(rotate)
        pm = page.getPixmap(matrix=trans, alpha=False)
        pm.writePNG(r'./tu'+'{:02}.png' .format(pg))
    print()






def getimgfile(input_paths, file_type=".png"):
    pathDir =  os.listdir(input_paths)
    imglist = list()
    for i in pathDir:
        if file_type in i:
            imglist.append(i)
    return imglist

def imgtopdf(input_paths="./", outputpath="./docimg.pdf", file_type=".png"):
    imglist = getimgfile(input_paths, file_type)
    (maxw, maxh) = Image.open(imglist[0]).size
    c = canvas.Canvas(outputpath, pagesize=portrait((maxw, maxh)))
    for i in range(len(imglist)):
        print("\r转换为PDF",i+1,"/",len(imglist),end="")
        c.drawImage(imglist[i], 0, 0, maxw, maxh)
        c.showPage()
        os.remove(imglist[i])
    c.save()