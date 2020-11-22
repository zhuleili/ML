# -*- coding:utf-8 -*-

import urllib
import urllib.request
import io
import sys
import xlsxwriter
import time
import logging
import numpy
from xlrd import open_workbook
from xlutils.copy import copy
import socket

from apscheduler.schedulers.background import BackgroundScheduler
import os

# gb18030
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


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


def getOnes(index):
    week = time.strftime("%w", time.localtime())
    if week == '0':
        week = '7'
    now_time = time.strftime("%Y-%m-%d", time.localtime())
    header_dict = {'Accept': 'text/html, */*; q=0.01',
                   #'Accept-Encoding': 'gzip, deflate',
                   'Accept-Language': 'zh-CN,zh;q=0.8',
                   'Connection': 'keep-alive',
                   'Cookie': 'FirstOKURL=http%3A//www.okooo.com/jingcai/; First_Source=www.okooo.com; bookMakerCustomData=a%3A1%3A%7Bs%3A14%3A%22AuthoriteBooks%22%3Ba%3A2%3A%7Bi%3A0%3Ba%3A12%3A%7Bi%3A0%3Bi%3A14%3Bi%3A1%3Bi%3A82%3Bi%3A2%3Bi%3A27%3Bi%3A3%3Bi%3A94%3Bi%3A4%3Bi%3A84%3Bi%3A5%3Bi%3A24%3Bi%3A6%3Bi%3A2%3Bi%3A7%3Bi%3A15%3Bi%3A8%3Bi%3A43%3Bi%3A9%3Bi%3A65%3Bi%3A10%3Bi%3A36%3Bi%3A11%3Bi%3A37%3B%7Ds%3A12%3A%22bookTypeName%22%3Bs%3A12%3A%22%C8%A8%CD%FE%B2%A9%B2%CA%B9%AB%CB%BE%22%3B%7D%7D; Last_Source=http%3A//www.okooo.com/jingcai/; OkAutoUuid=9f0c4882c8db1421dc83a7b62371442e; OkMsIndex=2; PHPSESSID=0bdc10a873fa56726f0e853a523246def47a35fe; DRUPAL_LOGGED_IN=Y; IMUserID=23960264; IMUserName=zll735508686; isInvitePurview=0; UWord=d401d8cd988f00b204e9800998ecf84427e; LastUrl=; Hm_lvt_5ffc07c2ca2eda4cc1c4d8e50804c94b=1513660142,1513750149; Hm_lpvt_5ffc07c2ca2eda4cc1c4d8e50804c94b=1513751484; __utma=56961525.1965621462.1513660140.1513660140.1513750149.2; __utmb=56961525.9.8.1513751484071; __utmc=56961525; __utmz=56961525.1513660140.1.1.utmcsr=(direct)|utmccn=(direct)|utmcmd=(none)',
                   'Host': 'www.okooo.com',
                   'If-Modified-Since': 'Wed, 20 Dec 2017 06:22:13 GMT',
                   'Referer': 'http://www.okooo.com/jingcai/',
                   'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36)',
                   'X-Requested-With': 'XMLHttpRequest'
                   }
    res = ''
    try:
        req = urllib.request.Request(url='http://www.okooo.com/jingcai/?action=more&LotteryNo=%s&MatchOrder=%s%s' % (now_time, week, index), headers=header_dict)
        res = urllib.request.urlopen(req)
        res = res.read().decode()
    except Exception as e:
        logger.info(e)
        return res
    return res


def getAddress(nowTime):
    header_dict = {'Accept': 'application/json, text/javascript, */*; q=0.01',
                   # 'Accept-Encoding': 'gzip, deflate',
                   'Accept-Language': 'zh-CN,zh;q=0.8',
                   'Connection': 'keep-alive',
                   'Cookie': 'FirstOKURL=http%3A//www.okooo.com/jingcai/; First_Source=www.okooo.com; bookMakerCustomData=a%3A1%3A%7Bs%3A14%3A%22AuthoriteBooks%22%3Ba%3A2%3A%7Bi%3A0%3Ba%3A12%3A%7Bi%3A0%3Bi%3A14%3Bi%3A1%3Bi%3A82%3Bi%3A2%3Bi%3A27%3Bi%3A3%3Bi%3A94%3Bi%3A4%3Bi%3A84%3Bi%3A5%3Bi%3A24%3Bi%3A6%3Bi%3A2%3Bi%3A7%3Bi%3A15%3Bi%3A8%3Bi%3A43%3Bi%3A9%3Bi%3A65%3Bi%3A10%3Bi%3A36%3Bi%3A11%3Bi%3A37%3B%7Ds%3A12%3A%22bookTypeName%22%3Bs%3A12%3A%22%C8%A8%CD%FE%B2%A9%B2%CA%B9%AB%CB%BE%22%3B%7D%7D; Last_Source=http%3A//www.okooo.com/jingcai/; OkAutoUuid=9f0c4882c8db1421dc83a7b62371442e; OkMsIndex=2; PHPSESSID=395e85dec11e9f5d47fee8825aad825703ecc7ab; DRUPAL_LOGGED_IN=Y; IMUserID=23960264; IMUserName=zll735508686; isInvitePurview=0; UWord=d441d8cd980f00b204e9800998ecf84827e; Hm_lvt_5ffc07c2ca2eda4cc1c4d8e50804c94b=1513660142,1513750149,1513829012; Hm_lpvt_5ffc07c2ca2eda4cc1c4d8e50804c94b=1513844227; __utma=56961525.1965621462.1513660140.1513841832.1513844150.9; __utmb=56961525.7.8.1513844227209; __utmc=56961525; __utmz=56961525.1513660140.1.1.utmcsr=(direct)|utmccn=(direct)|utmcmd=(none); LastUrl=',
                   'Host': 'www.okooo.com',
                   'Referer': 'http://www.okooo.com/jingcai/',
                   'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36)',
                   'X-Requested-With': 'XMLHttpRequest'
                   }
    result = []
    try:
        req = urllib.request.Request(url='http://www.okooo.com/jingcai/', headers=header_dict)
        html = urllib.request.urlopen(req).read().decode('gbk')
        # 获取html中第一个'<div class="cont "'的下标,今天的数据
        today = html.find('<span class="float_l">%s' % nowTime)
        html = html[today + 20:html.find('<span class="float_l">', today + 20)]
        #从今天开始-昨天截取
        xulie = html.find('xulie_h')
        while xulie != -1:
            html = html[xulie + 8:]
            ind = html.find('</span>')
            index = html[ind - 3: ind]  # 找到序列
            result.append(index)
            valuse = []
            valuse.append(index)
            html = html[html.find('class="shenpf'):]
            html = html[html.find('class="zhu'):]
            html = html[html.find('class="peilv fff hui_colo red_colo"'):]
            sheng = html[html.find('">') + 2:html.find('</div>')].replace('\r\n', '').replace(' ', '')  # 找到胜率
            valuse.append(sheng)
            html = html[html.find('class="ping'):]
            html = html[html.find('class="peilv fff hui_colo red_colo"'):]
            ping = html[html.find('">') + 2:html.find('</div>')].replace('\r\n', '').replace(' ', '')  # 找到平率
            valuse.append(ping)
            html = html[html.find('class="fu'):]
            html = html[html.find('class="peilv fff hui_colo red_colo"'):]
            fu = html[html.find('">') + 2:html.find('</div>')].replace('\r\n', '').replace(' ', '')  # 找到负率
            valuse.append(fu)

            count = 0
            data = getOnes(index)  # 获取详细信息
            while count < 10:
                if data != '':
                    break
                else:
                    count += 1
                    logger.info('%s比赛数据获取失败，重新获取第%s次' % (index, count))
                    time.sleep(1)
                    data = getOnes(index)  # 获取详细信息

            if float(sheng) <= float(fu):  # 胜/胜（A）和平/胜(B)的赔率，计算一下（A*B）/（A+B）
                if float(sheng) == 0:
                    valuse.append('0.00')
                    valuse.append('0.00')
                    valuse.append('0.00')
                else:
                    valuse.append(sheng)
                    data = data[data.find('胜/胜'):]
                    shengsheng = data[data.find('">') + 2:data.find('</div>', 10)].replace('\r\n', '').replace(' ', '')
                    data = data[data.find('平/胜'):]
                    pingsheng = data[data.find('">') + 2:data.find('</div>', 10)].replace('\r\n', '').replace(' ', '')
                    jisuan = (float(shengsheng) * float(pingsheng)) / (float(shengsheng) + float(pingsheng))
                    valuse.append('%.2f' % jisuan)
                    valuse.append('%.2f' % (jisuan / float(sheng)))
            elif float(sheng) > float(fu):  # 负负和平负
                valuse.append(fu)
                data = data[data.find('平/负'):]
                pingfu = data[data.find('">') + 2:data.find('</div>', 10)].replace('\r\n', '').replace(' ', '')
                data = data[data.find('负/负'):]
                fufu = data[data.find('">') + 2:data.find('</div>', 10)].replace('\r\n', '').replace(' ', '')
                jisuan = (float(fufu) * float(pingfu)) / (float(fufu) + float(pingfu))
                valuse.append('%.2f' % jisuan)
                valuse.append('%.2f' % (jisuan / float(fu)))
            result[result.index(index)] = valuse
            logger.info('%s比赛数据已获取' % index)
            xulie = html.find('xulie_h')
        return result
    except Exception as e:
        logger.info("连接失败")
        return result



def createTable(nowTime):
        if not os.path.exists('D:/投注信息/胜负/%s.xls' % nowTime):
            mkdir('D:/投注信息/胜负/')
            workbook = xlsxwriter.Workbook('D:/投注信息/胜负/%s.xls' % nowTime)
            workbook.add_worksheet("1")
            workbook.close()


def get_cell(table):
    try:
        if table.row(0)[0].value == "":
            return 0
    except Exception as e:
        if str(e).find('list index out of range') != -1:
            return 0
    try:
        for i in range(1000):
            table.row(0)[i].value
    except Exception as e:
        if str(e).find('list index out of range') != -1:
            return i


def insertExcel(result, nowTime):
    rb = open_workbook('D:/投注信息/胜负/%s.xls' % nowTime)
    wb = copy(rb)
    ws = wb.get_sheet(0)
    table = rb.sheet_by_index(0)
    cell = get_cell(table)
    if cell == 0:
        ws.write(0, 0, '编号')
        ws.write(0, 1, '胜')
        ws.write(0, 2, '平')
        ws.write(0, 3, '负')
        ws.write(0, 4, '最低值')
        ws.write(0, 5, '计算值')
        ws.write(0, 6, '计算值/最低值')
        ws.write(0, 7, '比赛结果')

        for i in range(len(result)):
            for j in range(len(result[i])):
                ws.write(i + 1, j, result[i][j])
        wb.save('D:/投注信息/胜负/%s.xls' % nowTime)
    else:
        ws.write(0, cell, '计算值/最低值')
        for i in range(len(result)):
            for j in range(len(result[i])):
                ws.write(i + 1, cell, result[i][j])
        wb.save('D:/投注信息/胜负/%s.xls' % nowTime)

if __name__ == '__main__':
    socket.setdefaulttimeout(30)
    nowTime = time.strftime("%Y-%m-%d", time.localtime())
    result = getAddress(nowTime)
    c = 0
    while c < 100:
        if len(result) > 0:
            createTable(nowTime)
            insertExcel(result, nowTime)
            logger.info("over")
            break
        else:
            c += 1
            logger.info('1秒后重新连接，第%s次' % c)
            time.sleep(1)
            result = getAddress(nowTime)
    else:
        logger.info('无法连接，请稍后重试！')