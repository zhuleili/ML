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
    header_dict = {'Accept': '*/*',
                   # 'Accept-Encoding': 'gzip, deflate',
                   'Accept-Language': 'zh-CN,zh;q=0.9',
                   'Connection': 'keep-alive',
                   'Cookie': 'win007LotteryCookie=null; jcWin007IsCleared=1; UM_distinctid=1727e4fe2302a5-069f538449ac49-3a36550e-15f900-1727e4fe2318ea; CNZZDATA1276099572=570986820-1591253853-%7C1591259421',
                   'Host': 'jc.win007.com',
                   'Referer': 'http://jc.win007.com/index.aspx',
                   'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.122 Safari/537.36'
                   }
    info = []
    try:
        req = urllib.request.Request(url='http://jc.win007.com/xml/odds_jc.txt?1591259474000', headers=header_dict)
        res = urllib.request.urlopen(req).read().decode('gb18030')
        arr = res.split('!')
        for i in arr:
            info.append(i.split('^')[0])
        return info
    except Exception as e:
        logger.info(e)
        logger.info("连接失败")
        return info


def getResult(info):
    results = {}
    for i in range(len(info)):
        logger.info('正在抓取第' + str(i + 1) + '场')
        res = getIds(info[i])
        result = getdatas(res)
        results[info[i]] = result
    return results


def createTable(nowTime):
    if not os.path.exists('D:/投注信息/凯利值/%s.xls' % nowTime):
        mkdir('D:/投注信息/凯利值/')
        workbook = xlsxwriter.Workbook('D:/投注信息/凯利值/%s.xls' % nowTime)
        workbook.close()


def getdatas(res):
    return ''

def getIds(id):
    header_dict = {'Accept': '*/*',
                   # 'Accept-Encoding': 'gzip, deflate',
                   'Accept-Language': 'zh-CN,zh;q=0.9',
                   'Connection': 'keep-alive',
                   'Cookie': 'UM_distinctid=1727e4fe2302a5-069f538449ac49-3a36550e-15f900-1727e4fe2318ea',
                   'Host': '1x2d.win007.com',
                   'Referer': 'http://op1.win007.com/oddslist/%s.htm' % id,
                   'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.122 Safari/537.36',
                   }
    res = ''
    try:
        req = urllib.request.Request(url='http://1x2d.win007.com/%s.js?r=007132357333948113867' % id, headers=header_dict)
        res = urllib.request.urlopen(req).read().decode('gb18030')

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
                if result[i][0][0] == "竞彩官方" and len(result[i][0]) > 5:
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
        else:
            ws.write(row, cell + 1, '胜')
            ws.write(row, cell + 2, '平')
            ws.write(row, cell + 3, '负')

            if len(result[i]) > 0:
                if result[i][0][0] == "竞彩官方" and len(result[i][0]) > 5:
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
        row += 5
    wb.save('D:/投注信息/凯利值/%s.xls' % nowTime)


if __name__ == '__main__':
    logger.info('开始')
    nowTime = time.strftime("%Y-%m-%d", time.localtime())
    #
    info = getAddress(nowTime)
    results = getResult(info)
    createTable(nowTime)
    cell = get_cell(nowTime)
    logger.info('处理成功，开始写入数据......')
    insertExcel(nowTime, results, cell)
    logger.info('完成')
