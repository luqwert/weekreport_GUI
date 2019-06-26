#!/usr/bin/env python  
# _*_ coding:utf-8 _*_  
# @Author  : lusheng

from tkinter import Label, Tk, Entry, Button
import weekreportsend
import mysteel
import test
import cnfeol
import platts
import huachengjinshu
import make_report


def get_cnfeol():
    cookies = entry.get()
    if cookies == '':
        cookies = 'bdshare_firstime=1465884234301; ls_token=047ada6b-ce18-47ca-8ead-4ca01a9632f2; evercookie_png=047ada6b-ce18-47ca-8ead-4ca01a9632f2; evercookie_etag=047ada6b-ce18-47ca-8ead-4ca01a9632f2; evercookie_cache=047ada6b-ce18-47ca-8ead-4ca01a9632f2; ASP.NET_SessionId=xkmjlrqoqlaenk45kuycg3eq; cnfeol_mn=; cnfeol_smn=; Hm_lvt_7d36fb642594f3133d486f18ce21e9fd=1558314083; cnfeol_mn=D4DF29036A2136C443B2919FB212EE01; cnfeol_smn=F6AD58D36D91CDBC5C420224F8B16EEF; mobileNotRedict=1; cnfeol_mkoc=false; cnfeol_mt_2015=40AE7CC5E559216E884991B0798EFF24EE6FE5E85F21D8467196F4694031AC4BF51A64B1469B8295834D9B6949562F5A041D903706701B8A3DC6C601CBB25A2084BAD9A65E4BBF459B2623B51FD2A95E148E4ED6622BC2DF3B4F13B6E9C54A89FC2E6C6D803932317EF0B4AF394DC4BFFC02C20ECA737DF227005A03E5E16FDB3316D1549D93F0644E060826F5A268246B4E35258445A97312CF0D8DAC8FEDDA4D71647F90E3AE07BFF62F59668992572543B1027F9302C4DC9ED839C9804A73ED3731C359D722A9DBF0D116179483AFA626375995EE6A626B9EDCEE75113206E6BFD9F18EF2BC727B7B6CAF2ECC73986DAD1A180AE4886A2DD19995B6BD0D2BD081A0E620056202234C3A584B6DA622E103B78D365CB5A289D40B310828D9F6941936234B35A33C6DB970098A148E8FDCEB0CF2F45E315D35E73A85610B9BC67B58711BB61BFB6D567393568BCF039DE35AC84FBC0D2D33C9D791AEE86F55DE9450B69BC16ED6DF2EF49A90899574102575A6696C076079B8635852B722B90C; cnfeol_mt=bC2tT9Z9wi/dBNNEng0GvySuhBY74ZKx0x+ZcgnhqsYmfrB+GGMCeSV6zOMXz6OhJz2F4hbWFkoSNfOz/8GTcb8ut1Zou7B7h4H5xnDN2XrNauw1N5H9NRpGrgSAJBZbBVAEVtPh4/ADG7knE+psX3rT7MoXzG0KHdzfC4ydJYo/GZ64hTvSs5/C0rT14Q8ftXMZN4ygXeWe2YMG1KKgQuXiHAwxd74Ofug/KWsXE7qAioWYe94EE1Vl3yNY8XdG6wUPxT+SYKZVBXPDVDYISPSK9DKRMPabprEWpJfgU0xnpXJg+zSwlBAdeW0EL22sNW8g5i+4deQ=; Hm_lpvt_7d36fb642594f3133d486f18ce21e9fd=1560822484'  # 没有备案过的浏览器，需要手机验证，无法手机验证，因此手动输入更新的cooki
    try:
        cnfeol.cnfeol('sinometal', '20091228', cookies)
        label11 = Label(root, text='铁合金在线数据下载成功', font=('楷体', 15), fg='red')
        label11.grid(row=1, column=1)
    except:
        label11 = Label(root, text='铁合金在线数据下载失败', font=('楷体', 15), fg='red')
        label11.grid(row=1, column=1)

def get_mysteel():
    try:
        mysteel.mysteel('hnxgscb', 'xg8659291')
        mysteel.steelhome('xmx', 'xiemx')
        label21 = Label(root, text='铁矿石数据下载成功', font=('楷体', 15), fg='red')
        label21.grid(row=2, column=1)
    except:
        label21 = Label(root, text='铁矿石数据下载失败', font=('楷体', 15), fg='red')
        label21.grid(row=2, column=1)

def get_platts():
    try:
        platts.platts('','')
        label31 = Label(root, text='platts数据下载成功', font=('楷体', 15), fg='red')
        label31.grid(row=3, column=1)
    except:
        label31 = Label(root, text='platts数据下载失败', font=('楷体', 15), fg='red')
        label31.grid(row=3, column=1)

def get_huachengjinshu():
    try:
        huachengjinshu.huachengjinshu()
        label41 = Label(root, text='华诚金属网数据下载成功', font=('楷体', 15), fg='red')
        label41.grid(row=4, column=1)
    except:
        label41 = Label(root, text='华诚金属网数据下载失败', font=('楷体', 15), fg='red')
        label41.grid(row=4, column=1)

def get_test():
    try:
        test.get_text()
        label51 = Label(root, text='华诚周报转换成功', font=('楷体', 15), fg='red')
        label51.grid(row=5, column=1)
    except:
        label51 = Label(root, text='华诚周报转换失败', font=('楷体', 15), fg='red')
        label51.grid(row=5, column=1)

def makereport():
    try:
        make_report.make_report()
        label61 = Label(root, text='周报生成成功', font=('楷体', 15), fg='red')
        label61.grid(row=6, column=1)
    except:
        label61 = Label(root, text='周报生成失败', font=('楷体', 15), fg='red')
        label61.grid(row=6, column=1)

def send_report():
    try:
        weekreportsend.send_report()
        label71 = Label(root, text='报告发送成功', font=('楷体', 15), fg='red')
        label71.grid(row=7, column=1)
    except:
        label71 = Label(root, text='报告发送失败', font=('楷体', 15), fg='red')
        label71.grid(row=7, column=1)

def onekey():
    get_cnfeol()
    get_mysteel()
    get_platts()
    get_huachengjinshu()
    get_test()
    makereport()


root = Tk()
root.title('周报自动工具')
root.geometry('600x320')
root.geometry('+400+200')
#文本输入框前的提示文本
label = Label(root, text='铁合金在线headers：',font=('楷体', 15),fg='blue')
label.grid()
#文本输入框
entry = Entry(root, font=('微软雅黑', 15),width=30)
entry.grid(row=0, column=1)
#铁合金在线按钮
button10 = Button(root,text='铁合金在线', font=('楷体',15),fg='red',command=get_cnfeol)
button10.grid(row=1, column=0)
#铁矿石按钮
button20 = Button(root,text='铁矿石数据', font=('楷体',15),fg='red',command=get_mysteel)
button20.grid(row=2, column=0)
#platts按钮
button30 = Button(root,text='platts数据', font=('楷体',15),fg='red',command=get_platts)
button30.grid(row=3, column=0)
#华诚金属按钮
button40 = Button(root,text='华诚金属网', font=('楷体',15),fg='red',command=get_huachengjinshu)
button40.grid(row=4, column=0)
#周报转换按钮
button50 = Button(root,text='周报转文本', font=('楷体',15),fg='red',command=get_test)
button50.grid(row=5, column=0)
#生成周报按钮
button60 = Button(root,text=' 生成周报 ', font=('楷体',15),fg='red',command=makereport)
button60.grid(row=6, column=0)
#发送周报按钮
button70 = Button(root,text=' 发送周报 ', font=('楷体',15),fg='red',command=send_report)
button70.grid(row=7, column=0)
#发送周报按钮
button80 = Button(root,text='                    一键生成                  ', font=('幼圆',15),fg='purple',command=onekey)
button80.grid(row=8, column=0, columnspan=2)


root.mainloop()

