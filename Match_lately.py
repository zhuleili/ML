# -*- coding:utf-8 -*-

import urllib
import urllib.request
import gzip
import http.cookiejar
import io
import sys
import xlsxwriter
import time
from datetime import datetime
import xlrd
import logging
import numpy
from xlrd import open_workbook
from xlutils.copy import copy
import socket

from apscheduler.schedulers.background import BackgroundScheduler
import os

# gb18030
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# 请求头值
header = {
    'Connection': 'Keep-alive',
    'Accept': 'text/html, */*',
    'Accept-Language': 'zh-CN,zh;q=0.8,en;q=0.6',
    'Accept-Encoding': 'gzip,deflate',
    'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36',
    'Host': 'http://www.okooo.com/jingcai/'
}
url = 'http://www.okooo.com/jingcai/'
title = ['borderLeft trbghui feedbackObj', 'trbghui feedbackObj', 'trbghui feedbackObj',
         'borderLeft feedbackObj', 'feedbackObj', 'bright  feedbackObj', 'feedbackObj', 'feedbackObj',
         'bright feedbackObj', 'borderRight borderLeft feedbackObj']
# 神笔马良竞彩 WJ04140806

def mkdir(path):
    # 去除首位空格
    path = path.strip()
    # 去除尾部 \ 符号
    path = path.rstrip("\\")

    # 判断路径是否存在
    # 存在     True
    # 不存在   False
    isExists = os.path.exists(path)

    # 判断结果
    if not isExists:
        # 如果不存在则创建目录
        # 创建目录操作函数
        os.makedirs(path)
        return True
    else:
        # 如果目录存在则不创建，并提示目录已存在
        return False


# 解压
def ungzip(data):
    try:
        data = gzip.decompress(data)
    except:
        logger.info('未经压缩，无需解压')
    return data.decode('gbk')


# 封装好请求头
def getOpener(head):
    # deal with the Cookies
    cj = http.cookiejar.CookieJar()
    pro = urllib.request.HTTPCookieProcessor(cj)
    opener = urllib.request.build_opener(pro)
    header = []
    for key, value in head.items():
        elem = (key, value)
        header.append(elem)
    opener.addheaders = header
    return opener


def getAddress():
    nowTime = time.strftime("%Y-%m-%d", time.localtime())
    header_dict = {'Accept': 'application/json, text/javascript, */*; q=0.01',
                   # 'Accept-Encoding': 'gzip, deflate',
                   'Accept-Language': 'zh-CN,zh;q=0.8',
                   'Connection': 'keep-alive',
                   'Cookie': '__utma=56961525.1153961841.1556420889.1556420889.1556420889.1; __utmc=56961525; __utmz=56961525.1556420889.1.1.utmcsr=(direct)|utmccn=(direct)|utmcmd=(none); LastUrl=; FirstOKURL=http%3A//www.okooo.com/jingcai/; First_Source=www.okooo.com; PHPSESSID=a9f58190114d92d8048554a7c86ede58cce33af4; pm=; Hm_lvt_5ffc07c2ca2eda4cc1c4d8e50804c94b=1556420890; Hm_lpvt_5ffc07c2ca2eda4cc1c4d8e50804c94b=1556421066; __utmb=56961525.10.8.1556421066409; IMUserID=30482642; IMUserName=%E4%BC%A6%E6%96%AF433832; OKSID=a9f58190114d92d8048554a7c86ede58cce33af4; M_UserName=%22ok_122874431803%22; M_UserID=30482642; M_Ukey=4f30ff7136a6ea67b30013378b10785d; DRUPAL_LOGGED_IN=Y; isInvitePurview=0',
                   'Host': 'www.okooo.com',
                   'Referer': 'http://www.okooo.com/jingcai/%s' % nowTime,
                   'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36)',
                   'X-Requested-With': 'XMLHttpRequest'
                   }
    a_url_arr = []
    info = []
    try:
        req = urllib.request.Request(url='http://www.okooo.com/jingcai/%s' % nowTime, headers=header_dict)
        html = urllib.request.urlopen(req).read().decode('gbk')
        # 获取html中第一个'<div class="cont "'的下标,今天的数据
        today = html.find('<span class="float_l">%s' % nowTime)
        html = html[today + 20:html.find('<span class="float_l">', today + 20)]
        #从今天开始-昨天截取

        xulie = html.find('xulie_h')
        while xulie != -1:
            html = html[xulie + 8:]
            ind = html.find('</span>')
            info.append(html[ind - 3: ind])
            html = html[html.find('/soccer/match/') + 10:]
            a_start = html.find('/soccer/match/')
            # while a_start != -1:
            a_end = html.find('history/', a_start)
            if a_end != -1:
                # 根据a_start和a_end截取字符串,获得地址
                a_url = 'http://www.okooo.com' + html[a_start:a_end + 8]
                # 加入数组
                a_url_arr.append(a_url)
                # 继续查找
                xulie = html.find('xulie_h')
        # logger.info(a_url_arr)
        for i in info:
            if info.count(i) > 1:
                del a_url_arr[info.index(i)]
                # logger.info(info.index(i))
                info.remove(i)
        return a_url_arr, info
    except Exception as e:
        logger.info("连接失败")
        return a_url_arr, info

def getDate(res, comp):
    arr2 = []
    c = 0
    res = res[res.find(comp):]
    while c < 20:
        arr = []
        res = res[res.find('data-lt="'):]
        res = res[res.find('trbghui blacktxt borderLeft'):]
        str = res[res.find('_blank') + 7:res.find('_blank') + 20]
        res = res[res.find('_blank') + 7:]
        if str.find('span') != -1:
            left = res[res.find('\'>') + 2:res.find('</span>')]
        else:
            left = res[res.find('>') + 1:res.find('</a>')]
        arr.append(left)

        res = res[res.find('trbghui'):]
        ro = res[res.find('attr') + 6:res.find('">')]
        if ro == '':
            break
        if ro == '-':
            ro = '0-0'
        a = ro.split('-')
        if a[0] != '':
            arr.append(int(a[0]))
        else:
            arr.append(0)
        if a[1] != '':
            arr.append(int(a[1]))
        else:
            arr.append(0)

        res = res[res.find('trbghui blacktxt bright'):]
        str = res[res.find('_blank') + 7:res.find('_blank') + 20]
        res = res[res.find('_blank') + 7:]
        if str.find('span') != -1:
            rigth = res[res.find('\'>') + 2:res.find('</span>')]
        else:
            rigth = res[res.find('>') + 1:res.find('</a>')]
        arr.append(rigth)
        arr2.append(arr)
        c += 1
        res = res[res.find('data-lt="'):]
    return arr2

def getResult(res):
    res = res[res.find('qpai_zi'):]
    zhu = res[res.find('qpai_zi') + 9:res.find('</div>')]
    vs = ''
    str = res[res.find('class="vs"'):res.find('class="vs"') + 50]
    if str.find('<span') != -1:
        res = res[res.find('</span>') - 1:]
        vs = res[:res.find('</span>')]
        res = res[res.find('</span>') + 5:]
        res = res[res.find('</span>') - 1:]
        vs = vs + '-' + res[:res.find('</span>')]

    res = res[res.find('qpai_zi_1'):]
    ke = res[res.find('qpai_zi_1') + 11:res.find('</div>')]

    arr_zhu = getDate(res, 'homecomp')
    arr_ke = getDate(res, 'awaycomp')
    return arr_zhu, arr_ke, zhu, ke, vs


def getLog():
    # 创建一个logger
    logger = logging.getLogger('mylogger')
    logger.setLevel(logging.DEBUG)

    # 创建一个handler，用于写入日志文件
    fh = logging.FileHandler('D:/投注信息/日志.log')
    fh.setLevel(logging.DEBUG)

    # 再创建一个handler，用于输出到控制台
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)

    # 定义handler的输出格式
    formatter = logging.Formatter('[%(asctime)s][line: %(lineno)d] ## %(message)s')
    fh.setFormatter(formatter)
    ch.setFormatter(formatter)

    # 给logger添加handler
    logger.addHandler(fh)
    logger.addHandler(ch)
    return logger
    # 记录一条日志


def createTable():
        nowFormatTime = time.strftime("%Y%m%d", time.localtime())
        if not os.path.exists('D:/投注信息/历史/%s.xls' % nowFormatTime):
            mkdir('D:/投注信息/历史/')
            workbook = xlsxwriter.Workbook('D:/投注信息/历史/%s.xls' % nowFormatTime)
            # for i in info:
            #     workbook.add_worksheet(i)
            workbook.close()


logger = getLog()


def startGetDate(url_arr, nowTime, info):
    try:
        logger.info('抓取时间:%s' % time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
        logger.info('=====开始抓取,今天共' + str(len(url_arr)) + '场比赛=====')
        fail = ''
        for i in range(len(url_arr)):
            # logger.info('--------------------------------------------------------')
            logger.info('正在抓取第' + str(i + 1) + '场,地址:' + url_arr[i])
            res = getOnes(url_arr[i])
            if res.find('lsdata') != -1:
                arr_zhu, arr_ke, zhu, ke, vs = getResult(res)  # 抓取结果(一场比赛)
                if len(arr_zhu) == 20 and len(arr_ke) == 20:
                    insertExcel(i, url_arr[i], nowTime, arr_zhu, arr_ke, zhu, ke, info, vs)
                # logger.info('第' + str(i + 1) + '场比赛数据已写入表格')
            else:
                # logger.info('第' + str(i + 1) + '场比赛抓取失败，进行下一场')
                fail = fail + str(i + 1) + ','
                pass
        if len(fail) > 0:
            logger.info('=====完成,失败场次：%s 文件路径:D:/投注信息/历史/%s.xls)=====' % (fail, nowTime))
            logger.info('')
        else:
            logger.info('=====全部完成,文件路径:D:/投注信息/历史/%s.xls)=====' % nowTime)
            logger.info('')
    except Exception as e:
        if str(e).find('404') != -1 or str(e).find('NoneType') != -1:
            logger.info(e)
            logger.info('未抓到数据，连接中断，请复制保存之前的数据，重新运行程序')
        elif str(e).find('Permission denied') != -1:
            logger.info('Excel文件被其他应用打开，请关闭！')
        else:
            logger.info(e)
            logger.info('')


def getOnes(arr):
    header_dict = {'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8 ',
                   #'Accept-Encoding': 'gzip, deflate',
                   'Accept-Language': 'zh-CN,zh;q=0.8',
                   'Connection': 'keep-alive',
                   'Cookie': 'bookMakerCustomData=a%3A1%3A%7Bs%3A14%3A%22AuthoriteBooks%22%3Ba%3A2%3A%7Bi%3A0%3Ba%3A14%3A%7Bi%3A0%3Bi%3A2%3Bi%3A1%3Bi%3A84%3Bi%3A2%3Bi%3A14%3Bi%3A3%3Bi%3A82%3Bi%3A4%3Bi%3A27%3Bi%3A5%3Bi%3A43%3Bi%3A6%3Bi%3A94%3Bi%3A7%3Bi%3A35%3Bi%3A8%3Bi%3A250%3Bi%3A9%3Bi%3A280%3Bi%3A10%3Bi%3A131%3Bi%3A11%3Bi%3A65%3Bi%3A12%3Bi%3A19%3Bi%3A13%3Bi%3A406%3B%7Ds%3A12%3A%22bookTypeName%22%3Bs%3A12%3A%22%C8%A8%CD%FE%B2%A9%B2%CA%B9%AB%CB%BE%22%3B%7D%7D; __utmz=56961525.1510387948.1.1.utmcsr=(direct)|utmccn=(direct)|utmcmd=(none); FirstOKURL=http%3A//www.okooo.com/jingcai/; First_Source=www.okooo.com; Last_Source=http%3A//www.okooo.com/jingcai/; zq_companytype_odds=AuthoriteBooks; OkAutoUuid=2babff49aecfed7b171579a4a967d1f0; OkMsIndex=6; LastUrl=; IMUserID=23449487; IMUserName=%E7%A5%9E%E7%AC%94%E9%A9%AC%E8%89%AF%E7%AB%9E%E5%BD%A9; UWord=d461d8cd982f00b204e9800998ecf84527e; __utma=56961525.325478833.1510387948.1521796942.1521878334.105; __utmc=56961525; Hm_lvt_5ffc07c2ca2eda4cc1c4d8e50804c94b=1521796942,1521878334; PHPSESSID=4370a77f88a53cc7bd5b70d3b0889c68f8459e01; pm=; LStatus=N; LoginStr=%7B%22welcome%22%3A%22%u60A8%u597D%uFF0C%u6B22%u8FCE%u60A8%22%2C%22login%22%3A%22%u767B%u5F55%22%2C%22register%22%3A%22%u6CE8%u518C%22%2C%22TrustLoginArr%22%3A%7B%22alipay%22%3A%7B%22LoginCn%22%3A%22%u652F%u4ED8%u5B9D%22%7D%2C%22tenpay%22%3A%7B%22LoginCn%22%3A%22%u8D22%u4ED8%u901A%22%7D%2C%22qq%22%3A%7B%22LoginCn%22%3A%22QQ%u767B%u5F55%22%7D%2C%22weibo%22%3A%7B%22LoginCn%22%3A%22%u65B0%u6D6A%u5FAE%u535A%22%7D%2C%22renren%22%3A%7B%22LoginCn%22%3A%22%u4EBA%u4EBA%u7F51%22%7D%2C%22baidu%22%3A%7B%22LoginCn%22%3A%22%u767E%u5EA6%22%7D%2C%22weixin%22%3A%7B%22LoginCn%22%3A%22%u5FAE%u4FE1%u767B%u5F55%22%7D%2C%22snda%22%3A%7B%22LoginCn%22%3A%22%u76DB%u5927%u767B%u5F55%22%7D%7D%2C%22userlevel%22%3A%22%22%2C%22flog%22%3A%22hidden%22%2C%22UserInfo%22%3A%22%22%2C%22loginSession%22%3A%22___GlobalSession%22%7D; Hm_lpvt_5ffc07c2ca2eda4cc1c4d8e50804c94b=1521878884; __utmb=56961525.23.7.1521878883204',
                   'Host': 'www.okooo.com',
                   'Referer': 'http://www.okooo.com/jingcai/%s' % nowTime,
                   'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36)'
                   }
    res = ''
    try:
        req = urllib.request.Request(url=arr, headers=header_dict)
        res = urllib.request.urlopen(req)
        res = res.read().decode('gbk')
    except Exception as e:
        logger.info(e)
        return res
    return res


def insertExcel(cell, url, nowTime, arr_zhu, arr_ke, zhu, ke, info, vs):
    rb = open_workbook('D:/投注信息/历史/%s.xls' % nowTime)
    wb = copy(rb)
    ws = wb.get_sheet(0)
    count = cell * 9
    ws.write(count, 0, url)
    ws.write(count, 1, '主队')
    ws.write(count, 2, zhu)
    ws.write(count, 4, '客队')
    ws.write(count, 5, ke)
    ws.write(count, 6, info[cell])
    ws.write(count, 7, vs)
    ws.write(count + 1, 1, '进球')
    ws.write(count + 1, 2, '失球')
    ws.write(count + 1, 4, '进球')
    ws.write(count + 1, 5, '失球')
    zhu_jin6 = 0
    zhu_shi6 = 0
    ke_jin6 = 0
    ke_shi6 = 0
    zhu_jin20 = 0
    zhu_shi20 = 0
    ke_jin20 = 0
    ke_shi20 = 0
    for i in range(len(arr_zhu)):
        if i < 6:
            if arr_zhu[i][0] == zhu:
                zhu_jin6 += arr_zhu[i][1]
                zhu_jin20 += arr_zhu[i][1]
                zhu_shi6 += arr_zhu[i][2]
                zhu_shi20 += arr_zhu[i][2]
            elif arr_zhu[i][3] == zhu:
                zhu_jin6 += arr_zhu[i][2]
                zhu_jin20 += arr_zhu[i][2]
                zhu_shi6 += arr_zhu[i][1]
                zhu_shi20 += arr_zhu[i][1]
        else:
            if arr_zhu[i][0] == zhu:
                zhu_jin20 += arr_zhu[i][1]
                zhu_shi20 += arr_zhu[i][2]
            elif arr_zhu[i][3] == zhu:
                zhu_jin20 += arr_zhu[i][2]
                zhu_shi20 += arr_zhu[i][1]
    for i in range(len(arr_ke)):
        if i < 6:
            if arr_ke[i][0] == ke:
                ke_jin6 += arr_ke[i][1]
                ke_jin20 += arr_ke[i][1]
                ke_shi6 += arr_ke[i][2]
                ke_shi20 += arr_ke[i][2]
            elif arr_ke[i][3] == ke:
                ke_jin6 += arr_ke[i][2]
                ke_jin20 += arr_ke[i][2]
                ke_shi6 += arr_ke[i][1]
                ke_shi20 += arr_ke[i][1]
        else:
            if arr_ke[i][0] == ke:
                ke_jin20 += arr_ke[i][1]
                ke_shi20 += arr_ke[i][2]
            elif arr_ke[i][3] == ke:
                ke_jin20 += arr_ke[i][1]
                ke_shi20 += arr_ke[i][2]

    ws.write(count + 2, 0, '6场')
    ws.write(count + 2, 1, zhu_jin6)
    ws.write(count + 2, 2, zhu_shi6)
    ws.write(count + 2, 4, ke_jin6)
    ws.write(count + 2, 5, ke_shi6)

    ws.write(count + 3, 0, '20场')
    ws.write(count + 3, 1, zhu_jin20)
    ws.write(count + 3, 2, zhu_shi20)
    ws.write(count + 3, 4, ke_jin20)
    ws.write(count + 3, 5, ke_shi20)

    ws.write(count + 4, 0, '6场平均')
    ws.write(count + 4, 1, '%.4f' % (zhu_jin6 / 6))
    ws.write(count + 4, 2, '%.4f' % (zhu_shi6 / 6))
    ws.write(count + 4, 4, '%.4f' % (ke_jin6 / 6))
    ws.write(count + 4, 5, '%.4f' % (ke_shi6 / 6))

    ws.write(count + 5, 0, '20场平均')
    ws.write(count + 5, 1, '%.4f' % (zhu_jin20 / 20))
    ws.write(count + 5, 2, '%.4f' % (zhu_shi20 / 20))
    ws.write(count + 5, 4, '%.4f' % (ke_jin20 / 20))
    ws.write(count + 5, 5, '%.4f' % (ke_shi20 / 20))

    ws.write(count + 6, 0, '6场平均/20场平均')
    zhu1 = 0
    zhu2 = 0
    ke1 = 0
    ke2 = 0
    if zhu_jin20 != 0:
        zhu1 = (zhu_jin6 / 6)/(zhu_jin20 / 20)
    if zhu_shi20 != 0:
        zhu2 = (zhu_shi6 / 6)/(zhu_shi20 / 20)
    if ke_jin20 != 0:
        ke1 = (ke_jin6 / 6)/(ke_jin20 / 20)
    if ke_shi20 != 0:
        ke2 = (ke_shi6 / 6)/(ke_shi20 / 20)
    ws.write(count + 6, 1, '%.4f' % zhu1)
    ws.write(count + 6, 2, '%.4f' % zhu2)
    ws.write(count + 6, 4, '%.4f' % ke1)
    ws.write(count + 6, 5, '%.4f' % ke2)

    ws.write(count + 7, 0, '')
    ws.write(count + 7, 1, '%.4f' % (zhu1 * ke2))
    ws.write(count + 7, 4, '%.4f' % (ke1 * zhu2))
    wb.save('D:/投注信息/历史/%s.xls' % nowTime)


if __name__ == '__main__':
    socket.setdefaulttimeout(30)
    nowTime = time.strftime("%Y%m%d", time.localtime())
    url_arr, info = getAddress()
    c = 0
    while c < 10:
        if len(url_arr) > 0:
            logger.info('连接成功!')
            createTable()
            # count = get_row(nowTime)
            startGetDate(url_arr, nowTime, info)
            break
        else:
            c += 1
            logger.info('1秒后重新连接，第%s次' % c)
            time.sleep(1)
            url_arr, info = getAddress()
    else:
        logger.info('无法连接，请稍后重试！')
