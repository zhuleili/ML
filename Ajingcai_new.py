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
    if not os.path.exists('D:/投注信息/新网站/%s.xls' % nowTime):
        mkdir('D:/投注信息/新网站/')
        workbook = xlsxwriter.Workbook('D:/投注信息/新网站/%s.xls' % nowTime)
        workbook.close()


logger = getLog()

def get_cell(nowTime):
    rb = open_workbook('D:/投注信息/新网站/%s.xls' % nowTime)
    table = rb.sheet_by_index(0)
    cell = 3
    try:
        if table.row(1)[0].value == "":
            return 0
    except Exception as e:
        if str(e).find('list index out of range') != -1:
            return 0
    try:
        for i in range(500):
            val = table.cell(1, cell).value
            # val = table.row(1)[4].value
            cell += 8
    except Exception as e:
        if str(e).find('array index out of range') != -1:
            return cell - 1


def getOnes(page, nowTime):
    header_dict = {'Accept': '*/*',
                   #'Accept-Encoding': 'gzip, deflate',
                   'Accept-Language': 'zh-CN,zh;q=0.9',
                   'Connection': 'keep-alive',
                   'Cookie': 'Hm_lvt_860f3361e3ed9c994816101d37900758=1552466770; BIGipServerPool_apache_web=1375797514.20480.0000; Hm_lpvt_860f3361e3ed9c994816101d37900758=1552467132; sajssdk_2015_cross_new_user=1; sensorsdata2015jssdkcross=%7B%22distinct_id%22%3A%2216976407f47df5-09acbca2d28717-551f3c12-1044480-16976407f49147%22%2C%22%24device_id%22%3A%2216976407f47df5-09acbca2d28717-551f3c12-1044480-16976407f49147%22%2C%22props%22%3A%7B%22%24latest_traffic_source_type%22%3A%22%E9%90%A9%E5%AD%98%E5%B8%B4%E5%A8%B4%E4%BE%80%E5%99%BA%22%2C%22%24latest_referrer%22%3A%22%22%2C%22%24latest_referrer_host%22%3A%22%22%2C%22%24latest_search_keyword%22%3A%22%E9%8F%88%EE%81%84%E5%BD%87%E9%8D%92%E6%9D%BF%E2%82%AC%E7%B3%AD%E9%90%A9%E5%AD%98%E5%B8%B4%E9%8E%B5%E6%92%B3%E7%B4%91%22%2C%22platForm%22%3A%22information%22%2C%22%24ip%22%3A%22222.71.16.161%22%2C%22source%22%3A%22pc%E7%AB%AF%22%2C%22browser_name%22%3A%22chrome%22%2C%22browser_version%22%3A%2269.0.3497.81%22%2C%22user_gent%22%3A%22Mozilla%2F5.0%20(Windows%20NT%206.1%3B%20Win64%3B%20x64)%20AppleWebKit%2F537.36%20(KHTML%2C%20like%20Gecko)%20Chrome%2F69.0.3497.81%20Safari%2F537.36%22%2C%22cname%22%3A%22%E4%B8%8A%E6%B5%B7%E5%B8%82%E6%99%AE%E9%99%80%E5%8C%BA%22%7D%2C%22data_from%22%3A%22information%22%7D',
                   'Host': 'i.sporttery.cn',
                   'Referer': 'https://info.sporttery.cn/basketball/vote/fb_vote.php',
                   'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.81 Safari/537.36'
                   }
    encode_json = ''
    try:
        req = urllib.request.Request(url='https://i.sporttery.cn/api/fb_match_info/get_lottery_vote/?p_code=had&f_callback=supportData&b_date=%s&l_id=0&page=%s&_=1552467342743' % (nowTime, page), headers=header_dict)
        res = urllib.request.urlopen(req)
        res = res.read().decode()
        res = res[12: len(res) - 2]
        encode_json = json.loads(res)
    except Exception as e:
        logger.info(e)
        return encode_json
    return encode_json

# 获取让球结果
def getRangOnes(page, nowTime):
    header_dict = {'Accept': '*/*',
                   #'Accept-Encoding': 'gzip, deflate',
                   'Accept-Language': 'zh-CN,zh;q=0.9',
                   'Connection': 'keep-alive',
                   'Cookie': 'Hm_lvt_860f3361e3ed9c994816101d37900758=1552466770; BIGipServerPool_apache_web=1375797514.20480.0000; Hm_lpvt_860f3361e3ed9c994816101d37900758=1552467132; sajssdk_2015_cross_new_user=1; sensorsdata2015jssdkcross=%7B%22distinct_id%22%3A%2216976407f47df5-09acbca2d28717-551f3c12-1044480-16976407f49147%22%2C%22%24device_id%22%3A%2216976407f47df5-09acbca2d28717-551f3c12-1044480-16976407f49147%22%2C%22props%22%3A%7B%22%24latest_traffic_source_type%22%3A%22%E9%90%A9%E5%AD%98%E5%B8%B4%E5%A8%B4%E4%BE%80%E5%99%BA%22%2C%22%24latest_referrer%22%3A%22%22%2C%22%24latest_referrer_host%22%3A%22%22%2C%22%24latest_search_keyword%22%3A%22%E9%8F%88%EE%81%84%E5%BD%87%E9%8D%92%E6%9D%BF%E2%82%AC%E7%B3%AD%E9%90%A9%E5%AD%98%E5%B8%B4%E9%8E%B5%E6%92%B3%E7%B4%91%22%2C%22platForm%22%3A%22information%22%2C%22%24ip%22%3A%22222.71.16.161%22%2C%22source%22%3A%22pc%E7%AB%AF%22%2C%22browser_name%22%3A%22chrome%22%2C%22browser_version%22%3A%2269.0.3497.81%22%2C%22user_gent%22%3A%22Mozilla%2F5.0%20(Windows%20NT%206.1%3B%20Win64%3B%20x64)%20AppleWebKit%2F537.36%20(KHTML%2C%20like%20Gecko)%20Chrome%2F69.0.3497.81%20Safari%2F537.36%22%2C%22cname%22%3A%22%E4%B8%8A%E6%B5%B7%E5%B8%82%E6%99%AE%E9%99%80%E5%8C%BA%22%7D%2C%22data_from%22%3A%22information%22%7D',
                   'Host': 'i.sporttery.cn',
                   'Referer': 'https://info.sporttery.cn/basketball/vote/fb_vote_hhad.php',
                   'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.81 Safari/537.36'
                   }
    encode_json = ''
    try:
        req = urllib.request.Request(url='https://i.sporttery.cn/api/fb_match_info/get_lottery_vote/?p_code=hhad&f_callback=supportData&b_date=%s&l_id=0&page=%s&_=1552467342743' % (nowTime, page), headers=header_dict)
        res = urllib.request.urlopen(req)
        res = res.read().decode()
        res = res[12: len(res) - 2]
        encode_json = json.loads(res)
    except Exception as e:
        logger.info(e)
        return encode_json
    return encode_json

# -------------------------------------------------------------
title = ['borderLeft feedbackObj', 'feedbackObj', 'bright  feedbackObj']
def getAddress(nowTime):
    request_time = (date.today() + timedelta(days = -1)).strftime("%Y-%m-%d")
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
        html = urllib.request.urlopen(req).read().decode('gb18030')
        # 获取html中第一个'<div class="cont "'的下标,今天的数据
        today = html.find('<span class="float_l">%s' % nowTime)
        if (today == -1):
            req = urllib.request.Request(url='http://www.okooo.com/jingcai/', headers=header_dict)
            html = urllib.request.urlopen(req).read().decode('gb18030')
            today = html.find('<span class="float_l">%s' % nowTime)
        html = html[today + 20:html.find('<span class="float_l">', today + 20)]
        #从今天开始-昨天截取

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

def getResult(url_arr, info):
    results = {}
    for i in range(len(url_arr)):
        logger.info('正在抓取第' + str(i + 1) + '场,地址:' + url_arr[i])
        res = getOldOnes(url_arr[i], 0)
        result = getdatas(res)
        pageNo = 0
        while len(result) < 3 and pageNo < 10:
            pageNo += 1
            res = getOldOnes(url_arr[i], pageNo)
            result_temp = getdatas(res)
            if(len(result_temp) > 0):
                for j in result_temp:
                    result.append(j)
                break
        results[info[i]] = result
    return results


def getdatas(res):
    result = []
    name_start = res.find('data-pname')
    while name_start != -1:
        data = []
        res = res[name_start + 12:]  # 名字开始的位置
        name_end = res.find('">')  # 用来记录名字结束的位置
        name = res[0: name_end]
        if name == '威廉.希尔' or name == 'Pinnacle' or name == '99家平均':
            data.append(name)
            res = res[res.find('trbghui bright'):]
            for j in range(3):
                res = res[res.find(title[j]):]  # 从td的下标+20截取
                in1 = res.find('<span class="')  # span的下标
                if in1 != -1:
                    res = res[in1 + 13:]  # 从span的下标+15截取,就是截掉span
                    red = res.find('bredtxt')
                    in_end = res.find('</span>')
                    if red == 0:
                        in1_in = res[9:in_end]  # 截取span里面的数据
                    elif in_end > 7:
                        in1_in = res[17:in_end]  # 截取span里面的数据
                    else:
                        in1_in = res[2:in_end]  # 截取span里面的数据
                    data.append(in1_in)
            result.append(data)
        name_start = res.find('data-pname')
    return result

def getOldOnes(arr, page):
    header_dict = {'Accept': 'text/html, */*',
                   #'Accept-Encoding': 'gzip, deflate',
                   'Accept-Language': 'zh-CN,zh;q=0.8',
                   'Connection': 'keep-alive',
                   'Cookie': 'bookMakerCustomData=a%3A1%3A%7Bs%3A14%3A%22AuthoriteBooks%22%3Ba%3A2%3A%7Bi%3A0%3Ba%3A14%3A%7Bi%3A0%3Bi%3A2%3Bi%3A1%3Bi%3A84%3Bi%3A2%3Bi%3A14%3Bi%3A3%3Bi%3A82%3Bi%3A4%3Bi%3A27%3Bi%3A5%3Bi%3A43%3Bi%3A6%3Bi%3A94%3Bi%3A7%3Bi%3A35%3Bi%3A8%3Bi%3A250%3Bi%3A9%3Bi%3A280%3Bi%3A10%3Bi%3A131%3Bi%3A11%3Bi%3A65%3Bi%3A12%3Bi%3A19%3Bi%3A13%3Bi%3A406%3B%7Ds%3A12%3A%22bookTypeName%22%3Bs%3A12%3A%22%C8%A8%CD%FE%B2%A9%B2%CA%B9%AB%CB%BE%22%3B%7D%7D; __utmz=56961525.1510387948.1.1.utmcsr=(direct)|utmccn=(direct)|utmcmd=(none); FirstOKURL=http%3A//www.okooo.com/jingcai/; First_Source=www.okooo.com; Last_Source=http%3A//www.okooo.com/jingcai/; zq_companytype_odds=AuthoriteBooks; __utma=56961525.1242083174.1521878451.1521445582.1521445582.82; Hm_lpvt_5ffc07c2ca2eda4cc1c4d8e50804c94b=1521445582; LastUrl=; PHPSESSID=9827ff99687eece3a3cfa4a75a51f6306da41961; DRUPAL_LOGGED_IN=Y; IMUserID=23449487; IMUserName=%E7%A5%9E%E7%AC%94%E9%A9%AC%E8%89%AF%E7%AB%9E%E5%BD%A9; OkAutoUuid=f3815e97d7dd7e187d70a53fe4b8a41f; OkMsIndex=5; isInvitePurview=0; UWord=9127842a7583615b7e7738eecc9d0263084; __utmc=56961525; Hm_lvt_5ffc07c2ca2eda4cc1c4d8e50804c94b=1521974578,1521291376,1520392383,1522222472',
                   'Host': 'www.okooo.com',
                   'Referer': arr,
                   'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36)',
                   'X-Requested-With': 'XMLHttpRequest'
                   }
    res = ''
    try:
        # BaijiaBooks  AuthoriteBooks
        req = urllib.request.Request(url='%sajax/?companytype=BaijiaBooks&page=%d&trnum=30&type=0' % (arr, page), headers=header_dict)
        res = urllib.request.urlopen(req)
        res = res.read().decode()
    except Exception as e:
        logger.info(e)
        return res
    return res
# -------------------------------------------------------------


def insertExcel(nowTime, result, cell):
    rb = open_workbook('D:/投注信息/新网站/%s.xls' % nowTime)
    wb = copy(rb)
    ws = wb.get_sheet(0)

    row = 1
    for i in result:
        ws.write(row + 2, 0, i['score'])

        if cell == 0:
            ws.write(row, cell, i['num'])
            ws.write(row + 1, cell, i['h_cn'] + ' VS ' + i['a_cn'])
            ws.write(row, cell + 1, '胜')
            ws.write(row + 1, cell + 1, '平')
            ws.write(row + 2, cell + 1, '负')

            ws.write(row, cell + 2, i['win_per'])
            ws.write(row + 1, cell + 2, i['draw_per'])
            ws.write(row + 2, cell + 2, i['lose_per'])

            ws.write(row, cell + 3, i['win_odds'])
            ws.write(row + 1, cell + 3, i['draw_odds'])
            ws.write(row + 2, cell + 3, i['lose_odds'])

            ws.write(row, cell + 4, i['win_err'])
            ws.write(row + 1, cell + 4, i['draw_err'])
            ws.write(row + 2, cell + 4, i['lose_err'])

            ws.write(row, cell + 5, i['win_99'])
            ws.write(row + 1, cell + 5, i['draw_99'])
            ws.write(row + 2, cell + 5, i['lose_99'])

            ws.write(row, cell + 6, i['win_bifa'])
            ws.write(row + 1, cell + 6, i['draw_bifa'])
            ws.write(row + 2, cell + 6, i['lose_bifa'])

            ws.write(row, cell + 7, i['win_guanfang'])
            ws.write(row + 1, cell + 7, i['draw_guanfang'])
            ws.write(row + 2, cell + 7, i['lose_guanfang'])

            if i['win_99'] != '' and i['win_bifa'] != '':
                if float(i['win_99']) >= float(i['lose_99']):
                    ws.write(row + 2, cell + 8, '%.2f' % (float(i['win_bifa']) - float(i['win_99'])))
                else:
                    ws.write(row + 2, cell + 8, '%.2f' % (float(i['lose_bifa']) - float(i['lose_99'])))

            if i['win_99'] != '' and i['win_guanfang'] != '':
                if float(i['win_99']) >= float(i['lose_99']):
                    ws.write(row + 1, cell + 8, '%.2f' % (float(i['win_guanfang']) - float(i['win_99'])))
                else:
                    ws.write(row + 1, cell + 8, '%.2f' % (float(i['lose_guanfang']) - float(i['lose_99'])))

            # 让球计算结果
            ws.write(row, cell + 8, i['result'])

            ws.write(0, cell + 5, '平均欧赔')
            ws.write(0, cell + 6, '威廉.希尔')
            ws.write(0, cell + 7, 'Pinnacle')
            ws.write(0, cell + 8, '让球计算')
        else:
            ws.write(row, cell, i['win_per'])
            ws.write(row + 1, cell, i['draw_per'])
            ws.write(row + 2, cell, i['lose_per'])

            ws.write(row, cell + 1, i['win_odds'])
            ws.write(row + 1, cell + 1, i['draw_odds'])
            ws.write(row + 2, cell + 1, i['lose_odds'])

            ws.write(row, cell + 2, i['win_err'])
            ws.write(row + 1, cell + 2, i['draw_err'])
            ws.write(row + 2, cell + 2, i['lose_err'])

            ws.write(row, cell + 3, i['win_99'])
            ws.write(row + 1, cell + 3, i['draw_99'])
            ws.write(row + 2, cell + 3, i['lose_99'])

            ws.write(row, cell + 4, i['win_bifa'])
            ws.write(row + 1, cell + 4, i['draw_bifa'])
            ws.write(row + 2, cell + 4, i['lose_bifa'])

            ws.write(row, cell + 5, i['win_guanfang'])
            ws.write(row + 1, cell + 5, i['draw_guanfang'])
            ws.write(row + 2, cell + 5, i['lose_guanfang'])

            if i['win_99'] != '' and i['win_bifa'] != '':
                if float(i['win_99']) >= float(i['lose_99']):
                    ws.write(row + 2, cell + 6, '%.2f' % (float(i['win_bifa']) - float(i['win_99'])))
                else:
                    ws.write(row + 2, cell + 6, '%.2f' % (float(i['lose_bifa']) - float(i['lose_99'])))

            if i['win_99'] != '' and i['win_guanfang'] != '':
                if float(i['win_99']) >= float(i['lose_99']):
                    ws.write(row + 1, cell + 6, '%.2f' % (float(i['win_guanfang']) - float(i['win_99'])))
                else:
                    ws.write(row + 1, cell + 6, '%.2f' % (float(i['lose_guanfang']) - float(i['lose_99'])))

            # 让球计算结果
            ws.write(row, cell + 6, i['result'])

            ws.write(0, cell + 3, '平均欧赔')
            ws.write(0, cell + 4, '威廉.希尔')
            ws.write(0, cell + 5, 'Pinnacle')
            ws.write(0, cell + 6, '让球计算')
        row += 4
    wb.save('D:/投注信息/新网站/%s.xls' % nowTime)


# 001的比赛，让球为（-1、-2的比赛）计算 胜+让胜+让平 的得票数得到A， 计算平+负+让负的得票数得到B，
# 用A/(A+B)得到C，用C减去胜对应赛果概率得到一个数值，将数值C记录在表格中
# -------------------
# 如果让球胜平负为（+1或者+2），计算负+让负+让平的得票数得到A， 计算胜+平+让胜的得票数得到B，
# 用A/(A+B)得到C，用C减去负对应赛果概率得到一个数值，将数值C记录在表格中
def calculation(votes, rang_votes, week):
    for i in range(len(votes)):
        if(votes[i]['num'][0:2] == week):
            votes[i]['result'] = ''
            rang_vote = get_rang_data(votes[i]['num'], week, rang_votes)
            if (rang_vote is not None):
                if(rang_vote['goalline'][:1] == '-'):
                    a = int(votes[i]['win_num']) + int(rang_vote['win_num']) + int(rang_vote['draw_num'])
                    b = int(votes[i]['draw_num']) + int(votes[i]['lose_num']) + int(rang_vote['lose_num'])
                    c = a / (a + b)
                    votes[i]['result'] = '%.1f' % ((c - int(votes[i]['win_ratio'][:len(votes[i]['win_ratio']) - 1]) / 100) * 100) + '%'
                if (rang_vote['goalline'][:1] == '+'):
                    a2 = int(votes[i]['lose_num']) + int(rang_vote['lose_num']) + int(rang_vote['draw_num'])
                    b2 = int(votes[i]['win_num']) + int(votes[i]['draw_num']) + int(rang_vote['win_num'])
                    c2 = a2 / (a2 + b2)
                    votes[i]['result'] = '%.1f' % ((c2 - int(votes[i]['lose_ratio'][:len(votes[i]['lose_ratio']) - 1]) / 100) * 100) + '%'


def get_rang_data(num, week, rang_votes):
    for i in rang_votes:
        if (i['num'] == num and i['num'][0:2] == week):
            return i


if __name__ == '__main__':
    logger.info('开始')
    nowTime = time.strftime("%Y-%m-%d", time.localtime())
    # 从澳客网获取 99家平 和 必发
    url_arr, info = getAddress(nowTime)
    results = getResult(url_arr, info)

    # 从新竞彩网获取数据
    d = datetime.datetime.now()
    weeks = {'0': '周一', '1': '周二', '2': '周三', '3': '周四', '4': '周五', '5': '周六', '6': '周日'}
    week = weeks[str(d.weekday())]
    createTable(nowTime)
    result = []
    all_votes = []
    all_rang_votes = []
    try:
        logger.info('开始获取竞彩网数据......')
        for i in range(20):
            data = getOnes(i + 1, nowTime)
            for j in data['result']['votes']:
                all_votes.append(j)

            # 获取让球结果
            rang_data = getRangOnes(i + 1, nowTime)
            for j in rang_data['result']['votes']:
                all_rang_votes.append(j)

        logger.info('获取成功，开始计算处理数据......')
        calculation(all_votes, all_rang_votes, week)

        for j in all_votes:
            if j['num'][0:2] == week:
                index = j['num'][2:5]
                j['win_99'] = ''
                j['draw_99'] = ''
                j['lose_99'] = ''
                j['win_guanfang'] = ''
                j['draw_guanfang'] = ''
                j['lose_guanfang'] = ''
                j['win_bifa'] = ''
                j['draw_bifa'] = ''
                j['lose_bifa'] = ''
                for k in results[index]:
                    if k[0] == '99家平均':
                        j['win_99'] = k[1]
                        j['draw_99'] = k[2]
                        j['lose_99'] = k[3]
                    elif k[0] == 'Pinnacle':
                        j['win_guanfang'] = k[1]
                        j['draw_guanfang'] = k[2]
                        j['lose_guanfang'] = k[3]
                    elif k[0] == '威廉.希尔':
                        j['win_bifa'] = k[1]
                        j['draw_bifa'] = k[2]
                        j['lose_bifa'] = k[3]
                result.append(j)
    except Exception as e:
        logger.info(e)
    cell = get_cell(nowTime)
    logger.info('处理成功，开始写入数据......')
    insertExcel(nowTime, result, cell)
    logger.info('完成')
