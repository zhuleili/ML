# -*- coding:utf-8 -*-

import urllib
import urllib.request
import gzip
import http.cookiejar
import io
import sys
import xlsxwriter
import time
import datetime
import xlrd
import logging
import numpy
from xlrd import open_workbook
from xlutils.copy import copy
import socket
import xlwt

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
baijia = ('竞彩官方', '澳门彩票', '威廉.希尔', '立博', 'Bet365', 'Interwetten', 'bwin', '易胜博',
          '皇冠(Singbet)', '利记(sbobet)', '沙巴(IBCBET)', '香港马会', '伟德国际', '必发')
title = ['borderLeft trbghui feedbackObj', 'trbghui feedbackObj', 'trbghui feedbackObj',
         'borderLeft feedbackObj', 'feedbackObj', 'bright  feedbackObj', 'feedbackObj', 'feedbackObj',
         'bright feedbackObj', 'borderRight borderLeft feedbackObj']
# 使得发达省份 WJ04140806

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


def getAddress(yesterday):
    header_dict = {'Accept': 'application/json, text/javascript, */*; q=0.01',
                   # 'Accept-Encoding': 'gzip, deflate',
                   'Accept-Language': 'zh-CN,zh;q=0.8',
                   'Connection': 'keep-alive',
                   'Cookie': 'FirstOKURL=http%3A//www.okooo.com/jingcai/; First_Source=www.okooo.com; bookMakerCustomData=a%3A1%3A%7Bs%3A14%3A%22AuthoriteBooks%22%3Ba%3A2%3A%7Bi%3A0%3Ba%3A12%3A%7Bi%3A0%3Bi%3A14%3Bi%3A1%3Bi%3A82%3Bi%3A2%3Bi%3A27%3Bi%3A3%3Bi%3A94%3Bi%3A4%3Bi%3A84%3Bi%3A5%3Bi%3A24%3Bi%3A6%3Bi%3A2%3Bi%3A7%3Bi%3A15%3Bi%3A8%3Bi%3A43%3Bi%3A9%3Bi%3A65%3Bi%3A10%3Bi%3A36%3Bi%3A11%3Bi%3A37%3B%7Ds%3A12%3A%22bookTypeName%22%3Bs%3A12%3A%22%C8%A8%CD%FE%B2%A9%B2%CA%B9%AB%CB%BE%22%3B%7D%7D; Last_Source=http%3A//www.okooo.com/jingcai/; OkAutoUuid=9f0c4882c8db1421dc83a7b62371442e; OkMsIndex=2; PHPSESSID=395e85dec11e9f5d47fee8825aad825703ecc7ab; DRUPAL_LOGGED_IN=Y; IMUserID=23960264; IMUserName=zll735508686; isInvitePurview=0; UWord=d441d8cd980f00b204e9800998ecf84827e; Hm_lvt_5ffc07c2ca2eda4cc1c4d8e50804c94b=1513660142,1513750149,1513829012; Hm_lpvt_5ffc07c2ca2eda4cc1c4d8e50804c94b=1513844227; __utma=56961525.1965621462.1513660140.1513841832.1513844150.9; __utmb=56961525.7.8.1513844227209; __utmc=56961525; __utmz=56961525.1513660140.1.1.utmcsr=(direct)|utmccn=(direct)|utmcmd=(none); LastUrl=',
                   'Host': 'www.okooo.com',
                   'Referer': 'http://www.okooo.com/jingcai/%s' % yesterday,
                   'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36)',
                   'X-Requested-With': 'XMLHttpRequest'
                   }
    info = []
    result = []
    try:
        req = urllib.request.Request(url='http://www.okooo.com/jingcai/%s' % yesterday, headers=header_dict)
        html = urllib.request.urlopen(req).read().decode('gbk')
        # 获取html中第一个'<div class="cont "'的下标,今天的数据
        # nowTime = time.strftime("%Y-%m-%d", time.localtime())
        yester = html.find('<span class="float_l">%s' % yesterday)
        # 截取昨天的部分
        html = html[yester + 20:html.find('<span class="float_l">', yester + 20)]
        xulie = html.find('xulie_h')
        while xulie != -1:
            html = html[xulie + 8:]
            ind = html.find('</span>')
            info.append(html[ind - 3: ind])
            html = html[html.find('more_bg hover1_2 sel1_2'):]
            start = html.find('class="p1">') + 11
            end = html.find('</p>')
            result.append(html[start:end])
            html = html[end + 5:]
            # 继续查找
            xulie = html.find('xulie_h')
        # logger.info(a_url_arr)
        for i in info:
            if info.count(i) > 1:
                del result[info.index(i)]
                # logger.info(info.index(i))
                info.remove(i)
        return result, info
    except Exception as e:
        logger.info("连接失败")
        return result, info


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

logger = getLog()


def getYesterday():
    today = datetime.date.today()
    return today - datetime.timedelta(days=1)


if __name__ == '__main__':
    socket.setdefaulttimeout(30)
    yesterday = str(getYesterday())
    result, info = getAddress(yesterday)
    c = 0
    font0 = xlwt.Font()
    font0.name = '5'
    font0.colour_index = 2  # 2红 3绿
    font0.bold = True
    style0 = xlwt.XFStyle()
    style0.font = font0

    while c < 10:
        if len(result) > 0:
            try:
                logger.info('连接成功!')
                lastTime = str(getYesterday()).replace('-', '')
                rb = open_workbook('D:/投注信息/14家/%s.xls' % lastTime, formatting_info=True)
                wb = copy(rb)
                for i in info:
                    ws = wb.get_sheet(i)
                    ws.write(0, 12, result[info.index(i)], style0)
                    wb.save('D:/投注信息/14家/%s.xls' % lastTime)
                logger.info("比赛结果已全部写入14家的表格")

                rb = open_workbook('D:/投注信息/2家/%s_2.xls' % lastTime, formatting_info=True)
                wb = copy(rb)
                for i in info:
                    ws = wb.get_sheet(i)
                    ws.write(0, 12, result[info.index(i)], style0)
                    wb.save('D:/投注信息/2家/%s_2.xls' % lastTime)
                logger.info("比赛结果已全部写入2家的表格")

                # rb = open_workbook('D:/投注信息/胜负/%s.xls' % str(getYesterday()), formatting_info=True)
                # wb = copy(rb)
                # ws = wb.get_sheet(0)
                # for i in range(len(result)):
                #     ws.write(i + 1, 7, result[i], style0)
                #     wb.save('D:/投注信息/胜负/%s.xls' % str(getYesterday()))
                # logger.info("比赛结果已全部写入胜负的表格")
                break
            except Exception as e:
                if str(e).find('Permission denied') != -1:
                    logger.info('Excel文件被其他应用打开，请关闭！')
                elif str(e).find('No such file or directory') != -1:
                    # logger.info('未找到昨天的文件！')
                    break
                else:
                    logger.info(e)
                    logger.info('')
        else:
            c += 1
            logger.info('1秒后重新连接，第%s次' % c)
            time.sleep(1)
            result, info = getAddress(yesterday)
    else:
        logger.info('无法连接，请稍后重试！')
