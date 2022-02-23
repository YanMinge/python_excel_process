# -*- coding: utf-8 -*-

# 字体下载 http://www.font5.com.cn/fontlist/fontlist_1_1.html

import xlrd
import cv2
import numpy
from datetime import datetime,date
from PIL import Image, ImageDraw, ImageFont

#月份查表
tup = ('JAN', 'FEB', 'MAY', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC')

#居中对齐书写文字
#img：操作的图片数组
#text：文字内容
#left：左侧坐标
#top：上部坐标
#width：书写区域的宽度
#high：书写区域的高度
#textColor：字体颜色设置
#textSize：字体大小设置
def cv2ImgAddTextCenter(img, text, left, top, width, high, textColor=(0, 255, 0), textSize = 20):
    #判断是否OpenCV图片类型
    if (isinstance(img, numpy.ndarray)):
        img = Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))
    draw = ImageDraw.Draw(img)

    #设定字体
    fontText = ImageFont.truetype("font/Cinzel-Regular.otf", textSize, encoding="utf-8")
    font_width, font_height = draw.textsize(text, fontText)
    print("font_width:%d, font_height:%d" %(font_width, font_height))

    #字符长度过长，缩小字体大小
    if(font_width > width):
        textSize = (int)(textSize * (width / font_width))
        fontText = ImageFont.truetype("font/Cinzel-Regular.otf", textSize, encoding="utf-8")
        font_width, font_height = draw.textsize(text, fontText)
        print("now font_width:%d, font_height:%d, textSize:%d" %(font_width, font_height, textSize))
    draw.text((left + (width - font_width-fontText.getoffset(text)[0]) / 2, top + (high - font_height-fontText.getoffset(text)[1]) / 2), text, textColor, font=fontText)
    return cv2.cvtColor(numpy.asarray(img), cv2.COLOR_RGB2BGR)

#指定位置书写文字
#img：操作的图片数组
#text：文字内容
#left：左侧坐标
#top：上部坐标
#textColor：字体颜色设置
#textSize：字体大小设置
def cv2ImgAddText(img, text, left, top, textColor=(0, 255, 0), textSize=20):
    #判断是否OpenCV图片类型
    if (isinstance(img, numpy.ndarray)):
        img = Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))
    draw = ImageDraw.Draw(img)

    #设定字体
    fontText = ImageFont.truetype("font/Cinzel-Regular.otf", textSize, encoding="utf-8")
    # fontText = ImageFont.truetype("font/华文行楷体简体/华文行楷.ttf", textSize, encoding="utf-8")
    font_width, font_height = draw.textsize(text, fontText)
    draw.text((left, top), text, textColor, font=fontText)
    return cv2.cvtColor(numpy.asarray(img), cv2.COLOR_RGB2BGR)

data = xlrd.open_workbook(r'MCE.xls')
bk_img = cv2.imread("bk_img.jpg")

#索引的方式，从0开始
# sheet = data.sheet_by_index(0)

#索引的方式, 名字
sheet = data.sheet_by_name('Advanced (2)')

nrows = sheet.nrows    #行
ncols = sheet.ncols    #列
print("nrows:%d, ncols:%d" %(nrows,ncols))

for i in range(nrows-13):
    print("value:%s" %(sheet.cell(i + 13, 3).value))
    img = cv2ImgAddTextCenter(bk_img, str(sheet.cell(i + 13, 3).value).upper(), 316, 500, 920, 60, (6, 4, 10), 60)
    print(xlrd.xldate_as_tuple(sheet.cell(i + 13, 0).value, data.datemode))
    img = cv2ImgAddTextCenter(img, 'MATATALAB CODING ADVANCED COURSE', 240, 660, 1076, 50, (6, 4, 10), 40)
    print(xlrd.xldate_as_tuple(sheet.cell(i + 13, 0).value, data.datemode))
    date_value = xlrd.xldate_as_tuple(sheet.cell(i + 13, 0).value, data.datemode)
    datastr = date(*date_value[:3]).strftime('%Y/%m/%d')
    img = cv2ImgAddText(img, datastr, 470, 883, (6, 4, 10), 15)
    datastr = date(*date_value[:3]).strftime('%Y')
    img = cv2ImgAddText(img, datastr, 1244, 871, (6, 4, 10), 15)
    datastr = date(*date_value[:3]).strftime('%Y')
    datastr = date(*date_value[:3]).strftime('%m')
    img = cv2ImgAddText(img, tup[int(datastr)-1], 1245, 775, (6, 4, 10), 15)
    cv2.imencode('.jpg', img)[1].tofile('advance/' + str(i+1+13) + '_' + str(sheet.cell(i + 13, 3).value) + '_CODING_ADVANCED' + '.jpg')


#索引的方式, 名字
sheet = data.sheet_by_name('Pro (2)')

nrows = sheet.nrows    #行
ncols = sheet.ncols    #列
print("nrows:%d, ncols:%d" %(nrows,ncols))

for i in range(nrows-33):
    print("value:%s" %(sheet.cell(i + 33, 3).value))
    img = cv2ImgAddTextCenter(bk_img, str(sheet.cell(i + 33, 3).value).upper(), 316, 500, 920, 60, (6, 4, 10), 60)
    print(xlrd.xldate_as_tuple(sheet.cell(i + 33, 0).value, data.datemode))
    img = cv2ImgAddTextCenter(img, 'MATATALAB CODING/PRO SET COURSE', 240, 660, 1076, 50, (6, 4, 10), 40)
    print(xlrd.xldate_as_tuple(sheet.cell(i + 33, 0).value, data.datemode))
    date_value = xlrd.xldate_as_tuple(sheet.cell(i + 33, 0).value, data.datemode)
    datastr = date(*date_value[:3]).strftime('%Y/%m/%d')
    img = cv2ImgAddText(img, datastr, 470, 883, (6, 4, 10), 15)
    datastr = date(*date_value[:3]).strftime('%Y')
    img = cv2ImgAddText(img, datastr, 1244, 871, (6, 4, 10), 15)
    datastr = date(*date_value[:3]).strftime('%Y')
    datastr = date(*date_value[:3]).strftime('%m')
    img = cv2ImgAddText(img, tup[int(datastr)-1], 1245, 775, (6, 4, 10), 15)
    cv2.imencode('.jpg', img)[1].tofile('pro/' + str(i+1+33) + '_' + str(sheet.cell(i + 33, 3).value) + '_CODING_PRO_SET' + '.jpg')

sheet = data.sheet_by_name('Lite (2)')

nrows = sheet.nrows    #行
ncols = sheet.ncols    #列
print("nrows:%d, ncols:%d" %(nrows,ncols))

for i in range(nrows - 16):
    print("value:%s" %(sheet.cell(i + 16, 3).value))
    img = cv2ImgAddTextCenter(bk_img, str(sheet.cell(i + 16, 3).value).upper(), 316, 500, 920, 60, (6, 4, 10), 60)
    print(xlrd.xldate_as_tuple(sheet.cell(i + 16, 0).value, data.datemode))
    img = cv2ImgAddTextCenter(img, 'MATATALAB LITE COURSE', 240, 660, 1076, 50, (6, 4, 10), 40)
    print(xlrd.xldate_as_tuple(sheet.cell(i + 16, 0).value, data.datemode))
    date_value = xlrd.xldate_as_tuple(sheet.cell(i + 16, 0).value, data.datemode)
    datastr = date(*date_value[:3]).strftime('%Y/%m/%d')
    img = cv2ImgAddText(img, datastr, 470, 883, (6, 4, 10), 15)
    datastr = date(*date_value[:3]).strftime('%Y')
    img = cv2ImgAddText(img, datastr, 1244, 871, (6, 4, 10), 15)
    datastr = date(*date_value[:3]).strftime('%Y')
    datastr = date(*date_value[:3]).strftime('%m')
    img = cv2ImgAddText(img, tup[int(datastr)-1], 1245, 775, (6, 4, 10), 15)
    # 对特殊字符进行异常处理
    if('/' in str(sheet.cell(i + 16, 3).value)):
        cv2.imencode('.jpg', img)[1].tofile('lite/' + str(i+1+16) + '_' + '00000000' + '_LITE' + '.jpg')
    else:
        cv2.imencode('.jpg', img)[1].tofile('lite/' + str(i+1+16) + '_' + str(sheet.cell(i + 16, 3).value) + '_LITE' + '.jpg')

cv2.waitKey()
cv2.destroyAllWindows()