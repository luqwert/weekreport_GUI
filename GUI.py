#!/usr/bin/env python  
# _*_ coding:utf-8 _*_  
# @Author  : lusheng

from tkinter import Label, Tk, Entry, Button
import weekreportsend
import docx_tpl
import globalmap as gl
gl._init()

def callback():
    print("你好")


def make_report():
    cookies = entry.get()
    if cookies == '':
        cookies = 'callback=jQuery111103071598438499954_1550644039396&c=40AE7CC5E559216E884991B0798EFF24EE6FE5E85F21D84632DDD8263819AC75D68D37137F89CDBABE9FFCDFAFA4DBA8F6B8A7B392465292A1D95694DB71BC9B21223EF5BEF9D7FB12AE73B2CFEC807D6095F9E129F5AEBE08CA2F7478122695C5DABA6D11523140878E08FADF0ACEC63F5C0C4DEB895AD0FCFE122DAF7907248AD2C1E3A147AE78A66995420E0C0FB572E8F317E0C9832AC40664C4D11AC642CB4C63A83AC151EEE983A071C49E92BAA28AFC5553F960A7D9542BAAD6EDBAAA7E89E2DF14C13569ABFB8373FEBF2B62350F023E8CEEDA852370D93637EC10DE79EB1AAFC7E077AFDAC2C308116C576917D7B9D14BC47FBC8B26C54BC727324547C2B2C2511E20E046ABCCE8016069AB3547C06255DBC2AAAAA5274D455CFBE68BAFF4406012DC815F8F2BC1593E8E024EB1FA3C2C973C88423310C8269761C7768866B43496EDFAE6A5F06DF787CE328DA465F119A3CB7C3D522C4214963E3EDF3742ABA3D7E3DA3ECF94046C2F47A5F6D1B5FA23F6AFC6&_=1550644039397'
    docx_tpl.make_report(cookies)
    # result = '报告生成成功'
    result = gl.get_value('result')
    label1 = Label(root, text=result, font=('楷体', 20), fg='red')
    label1.grid(row=1, column=1)

def send_report():
    weekreportsend.send_report()
    # result2 = '报告发送成功'
    result2 = gl.get_value('result2')
    label2 = Label(root, text=result2, font=('楷体', 20), fg='red')
    label2.grid(row=2, column=1)


root = Tk()
root.title('周报自动工具')
root.geometry('600x200')
root.geometry('+400+200')
#文本输入框前的提示文本
label = Label(root, text='铁合金在线headers：',font=('楷体', 20),fg='blue')
label.grid()
#文本输入框
entry = Entry(root, font=('微软雅黑', 12),width=35)
entry.grid(row=0, column=1)
#生成周报按钮
button1 = Button(root,text='生成周报', font=('楷体',20),fg='red',command=make_report)
button1.grid(row=1, column=0)
#发送周报按钮
button2 = Button(root,text='发送周报', font=('楷体',20),fg='red',command=send_report)
button2.grid(row=2, column=0)


root.mainloop()
