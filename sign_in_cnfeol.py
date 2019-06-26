#!/usr/bin/env python  
# _*_ coding:utf-8 _*_  
# @Author  : lusheng

import requests
from http import cookiejar

url1 = 'http://www.cnfeol.com/member/membersigninfrombox.aspx'
headers1 = {
    'Host': 'www.cnfeol.com',
    'Proxy-Connection': 'keep-alive',
    'Content-Length': '161',
    'Cache-Control': 'max-age=0',
    'Origin': 'http://www.cnfeol.com',
    'Upgrade-Insecure-Requests': '1',
    'Content-Type': 'application/x-www-form-urlencoded',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.75 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
    'Referer': 'http://www.cnfeol.com/',
    'Accept-Encoding': 'gzip, deflate',
    'Accept-Language': 'zh-CN,zh;q=0.9',
    'Cookie': 'bdshare_firstime=1465884234301; ls_token=047ada6b-ce18-47ca-8ead-4ca01a9632f2; evercookie_png=047ada6b-ce18-47ca-8ead-4ca01a9632f2; evercookie_etag=047ada6b-ce18-47ca-8ead-4ca01a9632f2; evercookie_cache=047ada6b-ce18-47ca-8ead-4ca01a9632f2; cnfeol_mn=D4DF29036A2136C443B2919FB212EE01; cnfeol_smn=F6AD58D36D91CDBC5C420224F8B16EEF; ASP.NET_SessionId=xkmjlrqoqlaenk45kuycg3eq; cnfeol_mn=; cnfeol_smn=; Hm_lvt_7d36fb642594f3133d486f18ce21e9fd=1550711229,1550804664,1552267374; cnfeol_mt_2015=; cnfeol_mkoc=; cnfeol_mt=; Hm_lpvt_7d36fb642594f3133d486f18ce21e9fd=1553220406',
}
data1 = {
    'return_url': 'http://www.cnfeol.com/',
    'signin_check': '650C943E117B8FB2BBE3F246DC448FD9',
    'Signin_MemberName': 'sinometal',
    'Signin_MemberPassword': '20091228',
    'Signin_Submit': '',
}


req1 = requests.post(url1,headers=headers1,data=data1)
print(req1)
print(req1.headers)
print(req1.cookies)

url2 = 'https://member.cnfeol.com/passport/login.aspx'
headers2 = {
    'Host': 'member.cnfeol.com',
    'Connection': 'keep-alive',
    'Content-Length': '314',
    'Cache-Control': 'max-age=0',
    'Origin': 'http://www.cnfeol.com',
    'Upgrade-Insecure-Requests': '1',
    'Content-Type': 'application/x-www-form-urlencoded',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.75 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
    'Referer': 'http://www.cnfeol.com/member/membersigninfrombox.aspx',
    'Accept-Encoding': 'gzip, deflate, br',
    'Accept-Language': 'zh-CN,zh;q=0.9',
    'Cookies': 'ls_token=047ada6b-ce18-47ca-8ead-4ca01a9632f2; evercookie_png=047ada6b-ce18-47ca-8ead-4ca01a9632f2; evercookie_etag=047ada6b-ce18-47ca-8ead-4ca01a9632f2; evercookie_cache=047ada6b-ce18-47ca-8ead-4ca01a9632f2; cnfeol_mn=D4DF29036A2136C443B2919FB212EE01; cnfeol_smn=F6AD58D36D91CDBC5C420224F8B16EEF; ASP.NET_SessionId=45iaviia3dgo4n55ehpzgs45; Hm_lvt_7d36fb642594f3133d486f18ce21e9fd=1550711229,1550804664,1552267374; cnfeol_mt_2015=; cnfeol_mkoc=; cnfeol_mt=; Hm_lpvt_7d36fb642594f3133d486f18ce21e9fd=1553220406'

}
# cookies2 = 'ls_token=047ada6b-ce18-47ca-8ead-4ca01a9632f2; evercookie_png=047ada6b-ce18-47ca-8ead-4ca01a9632f2; evercookie_etag=047ada6b-ce18-47ca-8ead-4ca01a9632f2; evercookie_cache=047ada6b-ce18-47ca-8ead-4ca01a9632f2; cnfeol_mn=D4DF29036A2136C443B2919FB212EE01; cnfeol_smn=F6AD58D36D91CDBC5C420224F8B16EEF; ASP.NET_SessionId=45iaviia3dgo4n55ehpzgs45; Hm_lvt_7d36fb642594f3133d486f18ce21e9fd=1550711229,1550804664,1552267374; cnfeol_mt_2015=; cnfeol_mkoc=; cnfeol_mt=; Hm_lpvt_7d36fb642594f3133d486f18ce21e9fd=1553220406'


data2 = {
    'sourceurl': 'http://www.cnfeol.com/',
    'un': 'sinometal',
    'pwd': 'a4944917e5c9c5dc0c37a73eb2b9b5ad2b81d18ed2a481c3b9f6e926d27dda336cb93cfa6d9beaf722b600b371f32372529ce3bedf2f2acb22c5704a34823f256a526c9e900a45079f62a361db3e00d0eb60ef0431e0e63e2cce901b40620e55f8105f345f30ef1821a22be3295795284ba8b5eea119a3a3ac02d93a3fa5da27',
}

req2 = requests.post(url2,headers=headers2,data=data2)

# cookies = req2.cookies
# cookies = requests.cookie.RequestsCookieJar()
# response = requests.request('post', url2
#                             , data=data2
#                             , headers=headers2
#                             , cookies=cookies2)  # 传递cookie
#
# self.cookies.update(response.cookies)
print(req2.status_code)
print(req2.headers)
print(req2.cookies)
# result = self.session.get(url, allow_redirects=False)#allow_redirects=False的意义为拒绝默认的301/302重定向从而可以通过html.headers[‘Location’]拿到重定向的URL
#                 self.resume_url = result.headers['location']  #拿到重定向后的url,这里面生成了我们需要的t和k值
