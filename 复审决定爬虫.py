
#作者：泊头子 微信公众号：专利方舟

#专利复审委复审决定提取源代码

from lxml import etree
import requests
import os
import time
import random
import socket
socket.setdefaulttimeout(20)      #设置默认超时时间
head={'User-Agent': 'Mozilla/5.0'}# 设置User-Agent浏览器信息
i1=range(170000,170100) #示例性提取复审决定号为170000—170100的100个复审决定
for juedinghao in i1:
    try:              #异常处理,url为某个复审决定网址
        url = "http://app.sipo-reexam.gov.cn/reexam_out1110/searchdoc/decidedetail.jsp?jdh="+str(juedinghao)+"&lx=fs"
        response = requests.get(url=url,headers=head) #获取响应状态码，返回某个复审决定网页源代码
    except:
        continue #异常后进入下一个循环
    response.encoding = response.apparent_encoding #获得真实编码,解决乱码
    root = etree.HTML(response.content)  #利用lxml.html的xpath对html进行分析，获取抓取信息
    response.close()                     #断开连接
    try:
        dizhi=root.xpath('//a/@href')[0] #a标签下选取名为href的所有属性，网页中只存在一个，实际获得复审决定WORD文件网址
    except:
        continue #异常后进入下一个循环
    try:
        r = requests.get(dizhi,headers=head,timeout=10) #获取响应状态码，返回WORD文件内容
        filename = os.path.basename(dizhi)              #从复审决定WORD文档网址中抽取文件名
        with open(filename, "wb") as code:              #打开文件
            code.write(r.content) # 写入内容
        r.close()                 #断开连接
        code.close()              #关闭文件,可以没有
        time.sleep(random.uniform(3,7)) #按给定的秒数暂停执行，时间自己定，本代码在3-7随机暂停，可以删除
    except:
        continue
    print(dizhi,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),juedinghao) #屏幕输出复审决定网址、当前时间及决定号
