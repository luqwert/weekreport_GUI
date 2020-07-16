#!/usr/bin/env python
# _*_ coding:utf-8 _*_
# @Author  : Administrator.DESKTOP-4V3P1KOheng
import datetime
import os

from wand.image import Image
from PIL import Image as mage
import pytesseract
import io
import re

# def get_text():
#     req_image = []
#     final_text = ''
#     with Image(filename=('C:\\Users\\Administrator.DESKTOP-4V3P1KO\Desktop\\周报材料\\' + '电解锰周评' + '.pdf'), resolution=400) as img:
#         with img.convert('jpeg') as converted:
#             # converted.save(filename='image.jpeg')
#             for img in converted.sequence:
#                 img_page = Image(image=img)
#                 req_image.append(img_page.make_blob('jpeg'))
#
#         #为每个图像运行OCR，识别图像中的文本
#         for img in req_image[-2:-1]:
#             text = pytesseract.image_to_string(mage.open(io.BytesIO(img)), lang='chi_sim')
#             final_text = final_text + text
#         # text = pytesseract.image_to_string(mage.open('image-0.jpeg'),lang='chi_sim')
#         final_text = final_text.replace(' ', '').replace('\n', '')
#         print(final_text)
#         mengpian_text = re.search(r'(?<=分析预测)(.+?)(?=\(原创)', final_text).group()
#         print(mengpian_text)
#
#         f_mysteel = open('C:\\Users\\Administrator.DESKTOP-4V3P1KO\\Desktop\\周报材料\\华诚金属.txt', 'w', encoding='utf-8')
#         f_mysteel.write(mengpian_text)
#         f_mysteel.close()

def get_text():
    filename = 'C:\\Users\\Administrator.DESKTOP-4V3P1KO\\Desktop\\周报材料\\' + '电解锰周评' + '.pdf'
    req_image = []
    final_text = ''
    with Image(filename=filename, resolution=300) as img:
        with img.convert('jpeg') as converted:
            n = len(converted.sequence)
            converted.save(filename='image.jpeg')
            # for img2 in converted.sequence:
            #     img_page = Image(image=img2)
            #     req_image.append(img_page.make_blob('jpeg'))

    # 为每个图像运行OCR，识别图像中的文本
    # for img in req_image:
    #     text = pytesseract.image_to_string(mage.open(io.BytesIO(img)), lang='chi_sim')
    #     final_text = final_text + text
    for i in range(n):
        text = pytesseract.image_to_string(mage.open('image-%d.jpeg' % i), lang='chi_sim')
        final_text = final_text + text
        os.remove('image-%d.jpeg' % i)
    final_text = final_text.replace(' ', '').replace('\n', '')
    print(final_text)
    # result = final_text
    mengpian_text = re.search(r'(?<=分析预测)(.+?)(?=\(原创)', final_text).group()
    print(mengpian_text)

    f_mysteel = open('C:\\Users\\Administrator.DESKTOP-4V3P1KO\\Desktop\\周报材料\\华诚金属.txt', 'w', encoding='utf-8')
    f_mysteel.write(mengpian_text)
    f_mysteel.close()

# get_text()