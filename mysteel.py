#!/usr/bin/env python
# _*_ coding:utf-8 _*_
# @Author  : lusheng


import requests
from selenium import webdriver
import time
import re
import globalmap as gl
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#mysteel
def mysteel(username, password):
    options = webdriver.ChromeOptions()
    # headers = '_last_loginuname=hnxgscb; _login_psd=f0884e703e554e127cdb7c9cde02cc810; _rememberStatus=true; 2f66e82bf907187fff7eef1623bc51ad=35%3D5%2636%3D5%2633%3D5%2634%3D5%2613%3D5%2637%3D5%2638%3D5%262%3D5%261%3D5%2642%3D5%2632%3D5%2641%3D5%264%3D10%2631%3D5%26catalog%3D020105%2C020206%2C040803%2C010205%2C010202%2C040805%2C040818%2C040801%2C0222%2C0223%2C0205; f2fb5a8705d66699a774f5b21b31e01a=35%3D5%2636%3D5%2633%3D5%2634%3D5%2613%3D5%2637%3D5%2638%3D5%262%3D5%261%3D5%2642%3D5%2632%3D5%2641%3D5%264%3D10%2631%3D5%26catalog%3D020105%2C020206%2C040803%2C010205%2C010202%2C040805%2C040818%2C040801%2C0222%2C0223%2C0205; Hm_lvt_1c4432afacfa2301369a5625795031b8=1557990427; _last_ch_r_t=1557990377532; Hm_lpvt_1c4432afacfa2301369a5625795031b8=1559526055'  # 没有备案过的浏览器，需要手机验证，无法手机验证，因此手动输入更新的cooki
    # options.add_argument(headers)
    prefs = {"profile.managed_default_content_settings.images": 2}
    options.add_experimental_option("prefs", prefs)
    # options.add_argument('--headless')
    # options.add_argument('--no-sandbox')
    browser = webdriver.Chrome(options=options)
    #登录mysteel铁矿石分站
    url = 'https://tks.mysteel.com/'
    browser.get(url)
    print(browser)
    time.sleep(4)
    # html = browser.find_element_by_xpath("//*").get_attribute("outerHTML")
    # print(html)

    # 库存情况
    kucunlink = browser.find_element_by_xpath('/html/body/ul[1]/li[8]/p/a[4]').get_attribute('href')
    # kucunlink = browser.find_element_by_css_selector('body > ul.navBox > li:nth-child(8) > p > a:nth-child(4)').get_attribute('href')
    kaigonglink = browser.find_element_by_xpath('/html/body/ul[1]/li[8]/p/a[6]').get_attribute('href')
    zhoupinglink = browser.find_element_by_xpath('/html/body/ul[1]/li[9]/p/a[4]').get_attribute('href')
    print(kucunlink,kaigonglink)
    #<a href="http://list1.mysteel.com/article/p-4260-------------1.html" target="_blank">港口库存</a>
    browser.get(kucunlink)
    kucunlink2 = browser.find_element_by_xpath('//*[@id="list"]/li[1]/h3/a').get_attribute('href')
    print(kucunlink2)
    #<a href="//tks.mysteel.com/18/1228/08/D46F179FC72070F8.html" title="12月28日进口矿港口库存统计与分析" target="_blank" class="ellipsis">12月28日进口矿港口库存统计与分析</a>
    browser.get(kucunlink2)
    # try:
    #     kucuntext = browser.find_element_by_xpath('//*[@id="text"]/p[1]').text
    #     gl.set_value('kucun_text', kucuntext)
    #     print(kucuntext)
    # except:
    #     print('需要重新登录')
    browser.find_element_by_xpath('//*[@id="userName"]').send_keys(username)
    browser.find_element_by_xpath('//*[@id="login-dialog"]/div[1]/div[2]/input[1]').send_keys(password)
    browser.find_element_by_xpath('//*[@id="login-dialog"]/div[1]/div[4]/button').click()
    time.sleep(3)
    kucuntext = browser.find_element_by_xpath('//*[@id="text"]/p[1]').text
    print(kucuntext)


    #高炉开工率
    browser.get(kaigonglink)
    kaigonglink2 = browser.find_element_by_xpath('//*[@id="list"]/li[1]/h3/a').get_attribute('href')
    print(kaigonglink2)
    #<a href="//tks.mysteel.com/18/1228/08/D46F179FC72070F8.html" title="12月28日进口矿港口库存统计与分析" target="_blank" class="ellipsis">12月28日进口矿港口库存统计与分析</a>
    browser.get(kaigonglink2)
    kaigongtext = browser.find_element_by_xpath('//*[@id="text"]/p').text
    print(kaigongtext)

    #铁矿石周评
    browser.get(zhoupinglink)
    zhoupinglink2 = browser.find_elements_by_tag_name('a')
    for i in zhoupinglink2:
        if '进口铁矿石' in i.get_attribute('title'):
            zhoupinglink2 = i.get_attribute('href')
            break
    print(zhoupinglink2)
    #获取周评文本内容
    browser.get(zhoupinglink2)
    zhoupinglist = browser.find_elements_by_xpath('//*[@id="text"]//p')
    print(len(zhoupinglist))
    zhoupingtext = ''
    for li in zhoupinglist:
        zhoupingtext = zhoupingtext + li.text
    zhoupingtext = zhoupingtext.replace('\n', '').replace(' ','')
    print(zhoupingtext)


    zhoupingtext1 = re.search('(.*?)(?=一、)',zhoupingtext).group() + '\r\n' + re.search('(?<=下周市场分析|节后市场预判|下周市场展望|下周市场预判|本周市场预判|上周市场预判)(.*?)(?=（文章|免责声明|（负责人)',zhoupingtext).group()
    # zhoupingtext1 = re.search('(?<=下周市场分析|下周市场展望|下周市场预判)(.*?)(?=免责声明)',zhoupingtext).group()
    print(zhoupingtext1)



    #废钢
    browser.get('https://feigang.mysteel.com/')
    browser.find_element_by_xpath('/html/body/div[4]/div[2]/div[3]/h4/a[2]').click()
    report_list = browser.find_elements_by_tag_name('a')
    # print(report_list)
    report_link = ''
    for report in report_list:
        if '全国废钢周评' in report.get_attribute('title'):
            report_link = report.get_attribute('href')
            report_title = report.get_attribute('title')
            print(report_link, report_title)
            break
    browser.get(report_link)
    feigangtext = browser.find_element_by_xpath('//*[@id="text"]/p[1]').text
    # try:
    #     feigangpic1_title = browser.find_element_by_xpath('//*[@id="text"]/p[2]').text
    # except:
    #     feigangpic1_title = browser.find_element_by_xpath('//*[@id="text"]/p[2]').text
    # try:
    #     feigangpic1 = browser.find_element_by_xpath('//*[@id="text"]/p[3]/img').get_attribute('src')
    # except:
    #     feigangpic1 = browser.find_element_by_xpath('//*[@id="text"]/p[3]/img').get_attribute('src')
    try:
        feigangpic2_title = browser.find_element_by_xpath('//*[@id="text"]/p[4]').text
    except:
        feigangpic2_title = browser.find_element_by_xpath('//*[@id="text"]/p[5]').text
    try:
        feigangpic2 = browser.find_element_by_xpath('//*[@id="text"]/p[6]/img').get_attribute('src')
    except:
        feigangpic2 = browser.find_element_by_xpath('//*[@id="text"]/p[5]/img').get_attribute('src')
    print(feigangtext)
    # print(feigangpic1_title, feigangpic1)
    print(feigangpic2_title, feigangpic2)

    # response = requests.get(feigangpic1)
    # print(response)
    # with open('C:\\Users\\Administrator.DESKTOP-4V3P1KO\\Desktop\\周报材料\\废钢指数近一年变化.png', 'wb') as f:
    #     f.write(response.content)
    #     f.close()

    response = requests.get(feigangpic2)
    print(response)
    with open('C:\\Users\\Administrator.DESKTOP-4V3P1KO\\Desktop\\周报材料\\各地废钢市场价格.png', 'wb') as f:
        f.write(response.content)
        f.close()
    f_mysteel = open('C:\\Users\\Administrator.DESKTOP-4V3P1KO\\Desktop\\周报材料\\mysteel.txt', 'w', encoding='utf-8')
    f_mysteel.write(kucuntext + '\n\n')
    f_mysteel.write(kaigongtext + '\n\n')
    f_mysteel.write(zhoupingtext1 + '\n\n')
    f_mysteel.write(feigangtext + '\n\n')
    f_mysteel.close()
    browser.quit()


#钢之家
def steelhome(username,password):
    options = webdriver.ChromeOptions()
    prefs = {"profile.managed_default_content_settings.images": 2}
    options.add_experimental_option("prefs", prefs)
    options.add_argument('--headless')
    browser = webdriver.Chrome(options=options)
    browser.get('http://news2.steelhome.cn/ll/col-058/var-c13/')
    report_list = browser.find_elements_by_xpath('/html/body/center/div[7]/div[2]/form/div/div/a')
    # print(report_list)
    report_link = ''
    for report in report_list:
        if '铁矿石市场一周综述' in report.get_attribute('title'):
            report_link = report.get_attribute('href')
            report_title = report.get_attribute('title')
            print(report_link,report_title)
            break
    browser.get(report_link)
    browser.find_element_by_xpath('//*[@id="username"]').send_keys(username)
    browser.find_element_by_xpath('//*[@id="password"]').send_keys(password)
    browser.find_element_by_xpath('//*[@id="sth_content"]/center/div/div/div/input[4]').click()
    time.sleep(2)
    gangzhijiatext = browser.find_element_by_xpath('//*[@id="sth_content"]').text
    # print(gangzhijiatext)
    gangzhijiatext = re.search('海运费：(.+?)美元。',gangzhijiatext).group()
    print(gangzhijiatext)

#保存获得的内容
    f_mysteel = open('C:\\Users\\Administrator.DESKTOP-4V3P1KO\\Desktop\\周报材料\\mysteel.txt', 'a+', encoding='utf-8')
    f_mysteel.write(gangzhijiatext)
    #关闭文件
    f_mysteel.close()
    browser.quit()


mysteel('hnxgscb', 'xg8659291')
steelhome('xmx', 'xiemx')
print('mysteel数据下载完成')
