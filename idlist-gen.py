#coding: utf-8
import datetime as dt
import pandas as pd
import numpy as np
import os

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont  
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

from reportlab.lib.units import inch
from reportlab.lib import fonts,colors
from reportlab.pdfbase.ttfonts import TTFont

from Tkinter import Tk
from tkFileDialog import askopenfilename, askdirectory

cwd = os.getcwd()

date_today = dt.date.today().strftime("%m-%d-%Y")
Tk().withdraw()
print "==============选取发货清单==============="
print "请在跳出的窗口中选中excel表格"
print "请用英文文件名和路径名，扩展名必须为xlsx"
raw_input("点击回车开始选择excel表格...")
inputXLSX = askopenfilename()
print "你选中的表格是："
print(inputXLSX)

outputPDF = inputXLSX[:-5] + "_usatocn2013_YU_身份证.pdf"


#print "==============选取照片文件夹================"
#print "请在下面跳出的窗口中选中身份证照片所在的文件夹，请用英文路径"
#raw_input("点击回车开始选择照片文件夹...")
#photodir = askdirectory()
#print "你选中的照片文件夹是"
#print(photodir)


df_shipList = pd.read_excel(inputXLSX)
df_shipList = df_shipList[[u"制单时间",u"运单号",u"收件人"]]
n0 = len(df_shipList)
print "%s一共有%s个包裹"%(date_today,n0)
print df_shipList
raw_input("正确的话请点回车继续，错误的话请重启程序重新选择...")

#当同一收件人有多个包裹时，只保留第一个包裹单号
df_shipList = df_shipList.drop_duplicates(subset = u"收件人", keep = "first")
print "%s我们一共送了%s个包裹"%(date_today,n0)
print "当同一收件人有多个包裹时，身份证清单中只保留第一个包裹"
print "移除重名后，身份证清单中一共有%s个包裹"%(len(df_shipList))
print df_shipList

#pdfmetrics.registerFont(TTFont("haha", "simsun.ttc"))

c3 = SimpleDocTemplate(outputPDF,pagesize=letter, rightMargin=72,leftMargin=72, topMargin=72,bottomMargin=18)
#c3.fontName   = "haha"
#c3.fontSize   = 12

idList = []
#missingID = []
missingN = 0

styles=getSampleStyleSheet()
#photoidpath = "./photo-id/"

for n in range(len(df_shipList)):
    # print tracking number
    tracking = str(df_shipList[n:n+1][u"运单号"].iloc[0])
    
    
    # print name
    #name = str(df_shipList[n:n+1]["name"].iloc[0].encode('utf-8'))
    #idList.append(Paragraph(u"刘女士",styles["Normal"]))
    
    # print photo-id
    try:
        id_0 = "./photo-id/" + df_shipList[n:n+1][u"收件人"].iloc[0].encode('utf-8') + ".jpg"
        im_0 = Image(id_0, 4*inch, 6*inch, kind='proportional')
        idList.append(Paragraph(tracking,styles["Normal"]))
        idList.append(im_0)
        idList.append(PageBreak())
    except IOError:
        try:
            id_1 = "./photo-id/" + df_shipList[n:n+1][u"收件人"].iloc[0].encode('utf-8') + "1.jpg"
            im_1 = Image(id_1, 4*inch, 4*inch, kind='proportional')
            im_1True = True
        except IOError:
#            missingID.append(id_1)
            im_1True = False
#            idList.append(Paragraph("Photo ID not available", styles["Normal"]))
            print id_1[11:-5] + '\t' + "正面" + ' ' + "身份证没有找到"
            missingN += 1
            
        if im_1True is True:
            idList.append(Paragraph(tracking,styles["Normal"]))
            idList.append(im_1)
            idList.append(Spacer(1,12))
            try:
                id_2 = "./photo-id/" + df_shipList[n:n+1][u"收件人"].iloc[0].encode('utf-8') + "2.jpg"
                im_2 = Image(id_2, 4*inch, 4*inch, kind='proportional')
                idList.append(im_2)
                im_2True = True
            except IOError:
                im_2True = False
                print id_1[11:-5] + '\t' + "背面" + ' ' + "身份证没有找到"
                missingN += 1
            idList.append(PageBreak())
print
print "一共%s个运单，其中%s个没有身份证"%(n+1,missingN)
print "生成文件已经保存到 " + cwd + " 目录下面"
print "文件名为 " + outputPDF
            
print
print "PS:"
print "请把缺失的身份证添加到图片目录下面"
print "如果正反面在同一张图片上，命名为 姓名.jpg"
print "如果正反面分两张图片，命名为 姓名1.jpg 和 姓名2.jpg"

    

c3.build(idList)
