# -*- coding:utf-8 -*-

import urllib
import urllib.request
import io
import sys
import xlsxwriter
import time
import logging
from xlrd import open_workbook
from xlutils.copy import copy
import socket
import os
import requests
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


# 调用高德地图接口，输入地址，获取经纬度
def geocodeG(address):
    par = {'address': address, 'key': 'cb649a25c1f81c1451adbeca73623251'}
    base = 'http://restapi.amap.com/v3/geocode/geo'
    response = requests.get(base, par)
    answer = response.json()
    GPS = []
    try:
        GPS = answer['geocodes'][0]['location'].split(",")
    except Exception as e:
        logger.info(e)
        logger.info("经纬度获取失败")
        return GPS
    return GPS

# 模拟发送请求，获取网页数据
def getAddress(i):
    header_dict = {'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
                   # 'Accept-Encoding': 'gzip, deflate',
                   'Accept-Language': 'zh-CN,zh;q=0.9',
                   'Cache-Control': 'max-age=0',
                   'Connection': 'keep-alive',
                   'Cookie': '_lxsdk_cuid=16228ac61a0c8-0d9458cabac90e-4448052d-100200-16228ac61a0c8; _lxsdk=16228ac61a0c8-0d9458cabac90e-4448052d-100200-16228ac61a0c8; _hc.v=adfe5fd7-3106-c2c1-7501-563eb3b77c93.1521100809; s_ViewType=10',
                   'Host': 'www.dianping.com',
                   'Upgrade-Insecure-Requests': '1',
                   'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.162 Safari/537.36'
                   }
    one_page = []
    try:
        # 发送请求
        req = urllib.request.Request(url='http://www.dianping.com/xian/ch10/p%s' % str(i), headers=header_dict)
        # 解析返回结果为字符串
        html = urllib.request.urlopen(req).read().decode()
        # 找到商家信息的模块，从此处截取，只保留后面的
        html = html[html.find('shop-all-list'):]
        while html.find('class="tit"') != -1:
            result = []
            # 获取店名
            html = html[html.find('class="tit"'):]
            html = html[html.find('<h4>'):]
            name = html[html.find('<h4>') + 4:html.find('</h4>')]
            # 获取评价数
            html = html[html.find('class="comment"'):]
            review = html[html.find('<b>') + 3:html.find('</b>')]
            # 获取店名
            html = html[html.find('</b>')+5:]
            price = html[html.find('<b>') + 3:html.find('</b>')]
            # 获取地名
            html = html[html.find('class="tag-addr"'):]
            html = html[html.find('class="addr"'):]
            addr = html[html.find('class="addr"') + 13:html.find('</span>')]
            result.append(name)
            result.append(price)
            result.append(review)
            result.append(addr)
            # 调用高德地图接口获取经纬度
            GPS = geocodeG('陕西省西安市%s' % addr)
            if len(GPS) > 1:
                result.append(GPS[0])
                result.append(GPS[1])
            else:
                result.append('0')
                result.append('0')
            one_page.append(result)
            # 继续查找下一个店面的信息
            html = html[html.find('class="tit"'):]
        return one_page
    except Exception as e:
        logger.info(e)
        logger.info("连接失败")
        return one_page


def getLog():
    # 创建一个logger
    logger = logging.getLogger('mylogger')
    logger.setLevel(logging.DEBUG)

    # 创建一个handler，用于写入日志文件
    fh = logging.FileHandler('D:/大众数据/日志.log')
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

# 创建文件夹及文件
def createTable(nowTime):
        if not os.path.exists('D:/大众数据/%s.xls' % nowTime):
            mkdir('D:/大众数据/')
            workbook = xlsxwriter.Workbook('D:/大众数据/%s.xls' % nowTime)
            workbook.close()


# 初始化日志
logger = getLog()


# 将获取到的结果保存到Excel
def insertExcel(nowTime, arr):
    try:
        rb = open_workbook('D:/大众数据/%s.xls' % nowTime)
        wb = copy(rb)
        ws = wb.get_sheet(0)
        ws.write(0, 0, '店名')
        ws.write(0, 1, '人均价')
        ws.write(0, 2, '评论人数')
        ws.write(0, 3, '位置')
        ws.write(0, 4, '经度')
        ws.write(0, 5, '纬度')
        for i in range(len(arr)):
            ws.write(i + 1, 0, arr[i][0])
            ws.write(i + 1, 1, arr[i][1])
            ws.write(i + 1, 2, arr[i][2])
            ws.write(i + 1, 3, arr[i][3])
            ws.write(i + 1, 4, arr[i][4])
            ws.write(i + 1, 5, arr[i][5])
        wb.save('D:/大众数据/%s.xls' % nowTime)
        logger.info('写入数据完成，D:/大众数据/%s.xls' % nowTime)
    except Exception as e:
        if str(e).find('Permission denied') != -1:
            logger.info('Excel文件被其他应用打开，请关闭后重新运行！')
        else:
            logger.info(e)
            logger.info("连接失败")


# 程序入口
if __name__ == '__main__':
    socket.setdefaulttimeout(30)  # 超时时间30秒
    page = 50  # 查找50页
    all_page = []  # 用来保存所有店家的数据
    nowTime = time.strftime("%Y-%m-%d", time.localtime())  # 获取当前的时间
    createTable(nowTime)  # 创建Excel文件
    # 循环查找50个页面
    for i in range(page):
        logger.info('正在抓取第' + str(i + 1) + '页数据')
        one_page = getAddress(i + 1)  # 一个页面的所有店家
        all_page = all_page + one_page  # 50个页面的所有店家
    logger.info('抓取完成，正在往excel写入数据，D:/大众数据/%s.xls' % nowTime)
    insertExcel(nowTime, all_page)
