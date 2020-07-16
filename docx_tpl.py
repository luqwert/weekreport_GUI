#!/usr/bin/env python  
# _*_ coding:utf-8 _*_  
# @Author  : Administrator.DESKTOP-4V3P1KOheng


from docx import shared
from docxtpl import DocxTemplate, RichText, InlineImage
import platts
import mysteel
import cnfeol
import asiametal
import huachengjinshu
import globalmap as gl
import os
import weekreportsend


def make_report(cookies):
    tpl = DocxTemplate('C:\\Users\\Administrator.DESKTOP-4V3P1KO\\Desktop\\周报材料\\周报模板.docx')
    # gl._init()

    cnfeol.cnfeol('sinometal', '20091228',cookies)
    platts.platts(1, 2)
    mysteel.mysteel('hnxgscb', 'xg8659291')
    mysteel.steelhome('xmx', 'xiemx')
    # asiametal.asiametal('sinometal', '50808266')
    huachengjinshu.huachengjinshu()
    if os.path.exists('C:\\Users\\Administrator.DESKTOP-4V3P1KO\\Desktop\\各产品周市场分析.docx'):
        os.remove('C:\\Users\\Administrator.DESKTOP-4V3P1KO\\Desktop\\各产品周市场分析.docx')

    # rt = RichText('an exemple of ')
    # rt.add('a rich text', style='')
    # rt.add('some violet', color='#ff00ff')
    image1 = InlineImage(tpl,'C:\\Users\\Administrator.DESKTOP-4V3P1KO\\Desktop\\周报材料\\普氏指数.jpeg',width=shared.Cm(16))
    image2 = InlineImage(tpl,'C:\\Users\\Administrator.DESKTOP-4V3P1KO\\Desktop\\周报材料\\废钢指数近一年变化.png',width=shared.Cm(16))
    image3 = InlineImage(tpl,'C:\\Users\\Administrator.DESKTOP-4V3P1KO\\Desktop\\周报材料\\各地废钢市场价格.png',width=shared.Cm(16))
    image4 = InlineImage(tpl,'C:\\Users\\Administrator.DESKTOP-4V3P1KO\\Desktop\\周报材料\\锰片价格变化.jpeg',width=shared.Cm(15))

    date1 = str(gl.get_value('date1'))[:10]
    date1_62 = gl.get_value('date1_62')
    date1_58 = gl.get_value('date1_58')
    date2 = str(gl.get_value('date2'))[:10]
    date2_62 = gl.get_value('date2_62')
    date2_58 = gl.get_value('date2_58')
    kucun_text = gl.get_value('kucun_text')
    kaigong_text = gl.get_value('kaigong_text')
    fenxi_text = gl.get_value('fenxi_text')
    haiyunfei_text = gl.get_value('haiyunfei_text')
    feigang_text = gl.get_value('feigang_text')
    menggkuang_kucun = gl.get_value('menggkuang_kucun')
    updown_mengkuang = gl.get_value('updown_mengkuang')
    diff_mengkuang = gl.get_value('diff_mengkuang')
    mengkuang_text = gl.get_value('mengkuang_text')
    guimeng_text = gl.get_value('guimeng_text')
    # stock = gl.get_value('stock')
    mengkuang_price = gl.get_value('mengkuang_price')
    stock2 = gl.get_value('stock2')
    guimeng_price = gl.get_value('guimeng_price')
    mengpian_text = gl.get_value('mengpian_text')
    # mengpian_text2 = gl.get_value('mengpian_text2')

    if date1_62 - date2_62 >= 0:
        updown_62 = '上涨'
        diff_62 = round(date1_62 - date2_62,2)
    else:
        updown_62 = '下跌'
        diff_62 = round(-(date1_62 - date2_62),2)

    if date1_58 - date2_58 >= 0:
        updown_58 = '上涨'
        diff_58 = round(date1_58 - date2_58,2)
    else:
        updown_58 = '下跌'
        diff_58 = round(-(date1_58 - date2_58),2)

    #需要传入的数据
    context = {
        # 文本和数字
        'date1': date1,
        'date2': date2,
        'date1_62': date1_62,
        'updown_62': updown_62,
        'diff_62': diff_62,
        'date1_58': date1_58,
        'updown_58': updown_58,
        'diff_58': diff_58,
        'kucun_text': kucun_text,
        'kaigong_text': kaigong_text,
        'fenxi_text': fenxi_text,
        'haiyunfei_text': haiyunfei_text,
        'feigang_text': feigang_text,
        'menggkuang_kucun': menggkuang_kucun,
        'updown_mengkuang': updown_mengkuang,
        'diff_mengkuang': diff_mengkuang,
        'mengkuang_text': mengkuang_text,
        'guimeng_text': guimeng_text,
        'mengpian_text': mengpian_text,
        # 'mengpian_text2': mengpian_text2,
        #表格
        # 'stock': stock, #表格，传入参数为list嵌套字典
        'stock2': stock2,
        'mengkuang_price': mengkuang_price,
        'guimeng_price': guimeng_price,

        #图片
        'image1': image1,
        'image2': image2,
        'image3': image3,
        'image4': image4,
    }

    tpl.render(context)
    tpl.save('C:\\Users\\Administrator.DESKTOP-4V3P1KO\\Desktop\\各产品周市场分析.docx')
    result = "周报已生成"
    gl.set_value('result',result)
    print('周报已生成完毕')
    # weekreportsend.send_report()
