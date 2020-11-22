# -*- coding:utf-8 -*-

import urllib
import urllib.request
import gzip
import http.cookiejar
import io
import sys
import xlsxwriter
import time
import logging
import datetime
from xlrd import open_workbook
from xlutils.copy import copy
import os
import json
from datetime import date, timedelta

# gb18030
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


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


def getLog():
    # 创建一个logger
    logger = logging.getLogger('mylogger')
    logger.setLevel(logging.DEBUG)

    open("D:/投注信息/日志.log", 'w')

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


def createTable(nowTime):
    if not os.path.exists('D:/投注信息/凯利值/%s.xls' % nowTime):
        mkdir('D:/投注信息/凯利值/')
        workbook = xlsxwriter.Workbook('D:/投注信息/凯利值/%s.xls' % nowTime)
        workbook.close()


def get_cell(nowTime):
    rb = open_workbook('D:/投注信息/凯利值/%s.xls' % nowTime)
    table = rb.sheet_by_index(0)
    cell = 0
    try:
        if table.row(1)[0].value == "":
            return 0
    except Exception as e:
        if str(e).find('list index out of range') != -1:
            return 0
    try:
        for i in range(500):
            table.cell(1, cell).value
            cell += 5
    except Exception as e:
        if str(e).find('array index out of range') != -1:
            return cell


logger = getLog()


# -------------------------------------------------------------
def getAddress(nowTime):
    # request_time = (date.today() + timedelta(days = -1)).strftime("%Y-%m-%d")
    header_dict = {'Accept': 'application/json, text/javascript, */*; q=0.01',
                   # 'Accept-Encoding': 'gzip, deflate',
                   'Accept-Language': 'zh-CN,zh;q=0.8',
                   'Connection': 'keep-alive',
                   'Cookie': 'bookMakerCustomData=a%3A1%3A%7Bs%3A14%3A%22AuthoriteBooks%22%3Ba%3A2%3A%7Bi%3A0%3Ba%3A14%3A%7Bi%3A0%3Bi%3A2%3Bi%3A1%3Bi%3A84%3Bi%3A2%3Bi%3A14%3Bi%3A3%3Bi%3A82%3Bi%3A4%3Bi%3A27%3Bi%3A5%3Bi%3A43%3Bi%3A6%3Bi%3A94%3Bi%3A7%3Bi%3A35%3Bi%3A8%3Bi%3A250%3Bi%3A9%3Bi%3A280%3Bi%3A10%3Bi%3A131%3Bi%3A11%3Bi%3A65%3Bi%3A12%3Bi%3A19%3Bi%3A13%3Bi%3A406%3B%7Ds%3A12%3A%22bookTypeName%22%3Bs%3A12%3A%22%C8%A8%CD%FE%B2%A9%B2%CA%B9%AB%CB%BE%22%3B%7D%7D; FirstOKURL=http%3A//www.okooo.com/jingcai/; First_Source=www.okooo.com; Last_Source=http%3A//www.okooo.com/jingcai/; zq_companytype_odds=AuthoriteBooks; PHPSESSID=58777b9e0f758bdf7a4cf02dee4acddb40e477ef; DRUPAL_LOGGED_IN=Y; IMUserID=23449487; IMUserName=%E7%A5%9E%E7%AC%94%E9%A9%AC%E8%89%AF%E7%AB%9E%E5%BD%A9; OkAutoUuid=2babff49aecfed7b171579a4a967d1f0; OkMsIndex=6; isInvitePurview=0; UWord=9147842a7503615b7e7738eecc9d0263084; LastUrl=; Hm_lvt_5ffc07c2ca2eda4cc1c4d8e50804c94b=1520217259,1520591379; Hm_lpvt_5ffc07c2ca2eda4cc1c4d8e50804c94b=1520591711; __utma=56961525.325478833.1510387948.1520242133.1520591379.5; __utmb=56961525.38.9.1520591725480; __utmc=56961525; __utmz=56961525.1510387948.1.1.utmcsr=(direct)|utmccn=(direct)|utmcmd=(none)',
                   'Host': 'www.okooo.com',
                   'Referer': 'http://www.okooo.com/jingcai/%s' % nowTime,
                   'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36)',
                   'X-Requested-With': 'XMLHttpRequest'
                   }
    a_url_arr = []
    info = []
    try:
        req = urllib.request.Request(url='http://www.okooo.com/jingcai/%s' % nowTime, headers=header_dict)
        # req = urllib.request.Request(url='http://www.okooo.com/', headers=header_dict)
        html = urllib.request.urlopen(req).read().decode('gb18030')
        # 获取html中第一个'<div class="cont "'的下标,今天的数据
        today = html.find('<span class="float_l">%s' % nowTime)
        xulie = html.find('xulie_h')
        if (today == -1 or xulie == -1):
            req = urllib.request.Request(url='http://www.okooo.com/jingcai/', headers=header_dict)
            html = urllib.request.urlopen(req).read().decode('gb18030')
            today = html.find('<span class="float_l">%s' % nowTime)
        html = html[today + 20:html.find('<span class="float_l">', today + 20)]
        # 从今天开始-昨天截取

        xulie = html.find('xulie_h')
        while xulie != -1:
            html = html[xulie + 8:]
            ind = html.find('</span>')
            info.append(html[ind - 3: ind])
            a_start = html.find('/soccer/match/')
            # while a_start != -1:
            a_end = html.find('odds', a_start)
            if a_end != -1:
                # 根据a_start和a_end截取字符串,获得地址
                a_url = 'http://www.okooo.com' + html[a_start:a_end + 5]
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
        logger.info(e)
        logger.info("连接失败")
        return a_url_arr, info

def getcokie():
    f = open('D:/投注信息/凯利值/cookie.txt')
    cookie = f.read()
    f.close()
    return cookie

def getResult(url_arr, info):
    results = {}
    cookie = getcokie()
    for i in range(len(url_arr)):
        logger.info('正在抓取第' + str(i + 1) + '场,地址:' + url_arr[i])
        res = getOldOnes(url_arr[i], cookie)
        result = getdatas(res)
        results[info[i]] = result
    return results


def createTable(nowTime):
    if not os.path.exists('D:/投注信息/凯利值/%s.xls' % nowTime):
        mkdir('D:/投注信息/凯利值/')
        workbook = xlsxwriter.Workbook('D:/投注信息/凯利值/%s.xls' % nowTime)
        workbook.close()


def getdatas(res):
    result = []
    name_start = res.find('data-pname')
    while name_start != -1:
        data = []
        res = res[name_start + 12:]  # 名字开始的位置
        name_end = res.find('">')  # 用来记录名字结束的位置
        name = res[0: name_end]
        if name == '99家平均' or name == 'Pinnacle' or name == '威廉.希尔' or name == '必发':
            data.append(name)
            zuidi = res[res.find('borderLeft trbghui feedbackObj') + 20:]
            res = res[res.find('bright  feedbackObj') + 10:]
            for j in range(3):
                res = res[res.find('feedbackObj') + 10:]
                in1 = res.find('</span>')  # span的下标
                in1_in = res[in1 - 4: in1]
                data.append(in1_in)

            if name == '99家平均':
                zui = {"胜": 0, "平": 0, "负": 0}
                for j in range(3):
                    in1 = zuidi.find('</span>')  # span的下标
                    in1_in = zuidi[in1 - 4: in1]
                    zuidi = zuidi[zuidi.find('</span>') + 5:]
                    if j == 0:
                        zui['胜'] = in1_in
                    elif j == 1:
                        zui['平'] = in1_in
                    elif j == 2:
                        zui['负'] = in1_in
                # 计算最低
                max_v = min([zui['胜'], zui['平'], zui['负']])
                name = ""
                if zui['胜'] == max_v:
                    name = "99家-胜"
                elif zui['平'] == max_v:
                    name = "99家-平"
                elif zui['负'] == max_v:
                    name = "99家-负"

                data.append(name)
                data.append(max_v)
            result.append(data)
        name_start = res.find('data-pname')
    return result

# 使得发达省份 WJ04140806
def getOldOnes(arr, cookie):
    header_dict = {'Accept': 'text/html, */*',
                   # 'Accept-Encoding': 'gzip, deflate',
                   'Accept-Language': 'zh-CN,zh;q=0.8',
                   'Connection': 'keep-alive',
                   'Cookie': cookie,
                   'Host': 'www.okooo.com',
                   'Referer': arr,
                   'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36)',
                   'X-Requested-With': 'XMLHttpRequest'
                   }
    res = ''
    try:
        # BaijiaBooks  AuthoriteBooks
        req = urllib.request.Request(url='%sajax/?companytype=MyOddsBooks&page=0&trnum=14&type=0' % arr,
                                     headers=header_dict)
        res = urllib.request.urlopen(req)
        res = res.read().decode()
    except Exception as e:
        logger.info(e)
        return res
    return res


# -------------------------------------------------------------


def insertExcel(nowTime, result, cell):
    rb = open_workbook('D:/投注信息/凯利值/%s.xls' % nowTime)
    wb = copy(rb)
    ws = wb.get_sheet(0)

    row = 0
    for i in result:
        if cell == 0:
            ws.write(row, cell, i)
            ws.write(row, cell + 1, '')
            ws.write(row, cell + 2, '胜')
            ws.write(row, cell + 3, '平')
            ws.write(row, cell + 4, '负')

            if len(result[i]) > 0:
                ws.write(row + 1, cell, result[i][0][0])
                if result[i][0][0] == "99家平均" and len(result[i][0]) > 5:
                    ws.write(row + 1, cell + 1, result[i][0][4])
                    ws.write(row + 2, cell + 1, result[i][0][5])
                ws.write(row + 1, cell + 2, result[i][0][1])
                ws.write(row + 1, cell + 3, result[i][0][2])
                ws.write(row + 1, cell + 4, result[i][0][3])

            if len(result[i]) > 1:
                ws.write(row + 2, cell, result[i][1][0])
                ws.write(row + 2, cell + 2, result[i][1][1])
                ws.write(row + 2, cell + 3, result[i][1][2])
                ws.write(row + 2, cell + 4, result[i][1][3])

            if len(result[i]) > 2:
                ws.write(row + 3, cell, result[i][2][0])
                ws.write(row + 3, cell + 1, '')
                ws.write(row + 3, cell + 2, result[i][2][1])
                ws.write(row + 3, cell + 3, result[i][2][2])
                ws.write(row + 3, cell + 4, result[i][2][3])

            if len(result[i]) > 3:
                ws.write(row + 4, cell, result[i][3][0])
                ws.write(row + 4, cell + 1, '')
                ws.write(row + 4, cell + 2, result[i][3][1])
                ws.write(row + 4, cell + 3, result[i][3][2])
                ws.write(row + 4, cell + 4, result[i][3][3])
        else:
            ws.write(row, cell + 1, '胜')
            ws.write(row, cell + 2, '平')
            ws.write(row, cell + 3, '负')

            if len(result[i]) > 0:
                if result[i][0][0] == "99家平均" and len(result[i][0]) > 5:
                    ws.write(row + 1, cell, result[i][0][4])
                    ws.write(row + 2, cell, result[i][0][5])
                ws.write(row + 1, cell + 1, result[i][0][1])
                ws.write(row + 1, cell + 2, result[i][0][2])
                ws.write(row + 1, cell + 3, result[i][0][3])
            if len(result[i]) > 1:
                ws.write(row + 2, cell + 1, result[i][1][1])
                ws.write(row + 2, cell + 2, result[i][1][2])
                ws.write(row + 2, cell + 3, result[i][1][3])
            if len(result[i]) > 2:
                ws.write(row + 3, cell + 1, result[i][2][1])
                ws.write(row + 3, cell + 2, result[i][2][2])
                ws.write(row + 3, cell + 3, result[i][2][3])
            if len(result[i]) > 3:
                ws.write(row + 4, cell + 1, result[i][3][1])
                ws.write(row + 4, cell + 2, result[i][3][2])
                ws.write(row + 4, cell + 3, result[i][3][3])
        row += 6
    wb.save('D:/投注信息/凯利值/%s.xls' % nowTime)


if __name__ == '__main__':
    logger.info('开始')
    nowTime = time.strftime("%Y-%m-%d", time.localtime())
    # 从澳客网获取 99家平 和 必发
    url_arr, info = getAddress(nowTime)
    results = getResult(url_arr, info)
    createTable(nowTime)
    cell = get_cell(nowTime)
    logger.info('处理成功，开始写入数据......')
    insertExcel(nowTime, results, cell)
    logger.info('完成，文件地址：D:/投注信息/凯利值/%s.xls' % nowTime)
