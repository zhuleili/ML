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
from xlrd import open_workbook
from xlutils.copy import copy
import socket

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
# zll735508686 WJ04140806

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
    header_dict = {'Accept': 'application/json, text/javascript, */*; q=0.01',
                   # 'Accept-Encoding': 'gzip, deflate',
                   'Accept-Language': 'zh-CN,zh;q=0.8',
                   'Connection': 'keep-alive',
                   'Cookie': 'bookMakerCustomData=a%3A1%3A%7Bs%3A14%3A%22AuthoriteBooks%22%3Ba%3A2%3A%7Bi%3A0%3Ba%3A14%3A%7Bi%3A0%3Bi%3A2%3Bi%3A1%3Bi%3A84%3Bi%3A2%3Bi%3A14%3Bi%3A3%3Bi%3A82%3Bi%3A4%3Bi%3A27%3Bi%3A5%3Bi%3A43%3Bi%3A6%3Bi%3A94%3Bi%3A7%3Bi%3A35%3Bi%3A8%3Bi%3A250%3Bi%3A9%3Bi%3A280%3Bi%3A10%3Bi%3A131%3Bi%3A11%3Bi%3A65%3Bi%3A12%3Bi%3A19%3Bi%3A13%3Bi%3A406%3B%7Ds%3A12%3A%22bookTypeName%22%3Bs%3A12%3A%22%C8%A8%CD%FE%B2%A9%B2%CA%B9%AB%CB%BE%22%3B%7D%7D; FirstOKURL=http%3A//www.okooo.com/jingcai/; First_Source=www.okooo.com; PHPSESSID=5397c55438a8970d0999a254ea68c46fb22cea3d; Last_Source=http%3A//www.okooo.com/jingcai/; zq_companytype_odds=AuthoriteBooks; DRUPAL_LOGGED_IN=Y; IMUserID=23960264; IMUserName=zll735508686; OkAutoUuid=deee31c6a2342121bf490b99d26ca39d; OkMsIndex=6; isInvitePurview=0; UWord=5c4ef7418f370ef8ab80183d01d55e118eb; LastUrl=; Hm_lvt_5ffc07c2ca2eda4cc1c4d8e50804c94b=1520217259; Hm_lpvt_5ffc07c2ca2eda4cc1c4d8e50804c94b=1520233747; __utma=56961525.325478833.1510387948.1520217258.1520231584.3; __utmb=56961525.16.8.1520233746877; __utmc=56961525; __utmz=56961525.1510387948.1.1.utmcsr=(direct)|utmccn=(direct)|utmcmd=(none)',
                   'Host': 'www.okooo.com',
                   'Referer': 'http://www.okooo.com/jingcai/',
                   'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36)',
                   'X-Requested-With': 'XMLHttpRequest'
                   }
    a_url_arr = []
    info = []
    try:
        req = urllib.request.Request(url='http://www.okooo.com/jingcai/', headers=header_dict)
        html = urllib.request.urlopen(req).read().decode('gbk')
        # 获取html中第一个'<div class="cont "'的下标,今天的数据
        nowTime = time.strftime("%Y-%m-%d", time.localtime())
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
        logger.info("连接失败")
        return a_url_arr, info

def getResult(res):
    result = []
    name_start = res.find('data-pname')
    re = ('Bet365','bwin')
    while name_start != -1:
        data = []
        res = res[name_start + 12:]  # 名字开始的位置
        name_end = res.find('">')  # 用来记录名字结束的位置
        name = res[0: name_end]
        if name in re:
            data.append(name)
            for i in range(10):
                res = res[res.find(title[i]):]  # 从td的下标+20截取
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


def createTable(info):
        nowFormatTime = time.strftime("%Y%m%d", time.localtime())
        if not os.path.exists('D:/投注信息/2家/%s_2.xls' % nowFormatTime):
            mkdir('D:/投注信息/2家/')
            workbook = xlsxwriter.Workbook('D:/投注信息/2家/%s_2.xls' % nowFormatTime)
            for i in info:
                workbook.add_worksheet(i)
            workbook.close()


logger = getLog()

def get_row(nowTime):
    rb = open_workbook('D:/投注信息/2家/%s_2.xls' % nowTime)
    table = rb.sheet_by_index(0)
    row = 4
    try:
        if table.row(0)[0].value == "":
            return 0
    except Exception as e:
        if str(e).find('list index out of range') != -1:
            return 0
    try:
        for i in range(24):
            # val = table.cell(4, row).value
            val = table.row(row)[4].value
            row += 1
    except Exception as e:
        if str(e).find('list index out of range') != -1:
            return row + 1

def startGetDate(url_arr, count, nowTime):
    try:
        logger.info('抓取时间:%s' % time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
        logger.info('=====开始抓取,今天共' + str(len(url_arr)) + '场比赛=====')
        fail = ''
        for i in range(len(url_arr)):
            # logger.info('--------------------------------------------------------')
            logger.info('正在抓取第' + str(i + 1) + '场,地址:' + url_arr[i])
            res = getOnes(url_arr[i])
            if res.find('data-pname') != -1:
                result = getResult(res)  # 抓取结果(一场比赛)
                insertExcel(i, url_arr[i], nowTime, result, count)
                # logger.info('第' + str(i + 1) + '场比赛数据已写入表格')
            else:
                # logger.info('第' + str(i + 1) + '场比赛抓取失败，进行下一场')
                fail = fail + str(i + 1) + ','
                pass
        if len(fail) > 0:
            logger.info('=====完成,失败场次：%s 文件路径:D:/投注信息/2家/%s_2.xls)=====' % (fail, nowTime))
            logger.info('')
        else:
            logger.info('=====全部完成,文件路径:D:/投注信息/2家/%s_2.xls)=====' % nowTime)
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
    header_dict = {'Accept': 'text/html, */*',
                   #'Accept-Encoding': 'gzip, deflate',
                   'Accept-Language': 'zh-CN,zh;q=0.8',
                   'Connection': 'keep-alive',
                   'Cookie': 'bookMakerCustomData=a%3A1%3A%7Bs%3A14%3A%22AuthoriteBooks%22%3Ba%3A2%3A%7Bi%3A0%3Ba%3A14%3A%7Bi%3A0%3Bi%3A2%3Bi%3A1%3Bi%3A84%3Bi%3A2%3Bi%3A14%3Bi%3A3%3Bi%3A82%3Bi%3A4%3Bi%3A27%3Bi%3A5%3Bi%3A43%3Bi%3A6%3Bi%3A94%3Bi%3A7%3Bi%3A35%3Bi%3A8%3Bi%3A250%3Bi%3A9%3Bi%3A280%3Bi%3A10%3Bi%3A131%3Bi%3A11%3Bi%3A65%3Bi%3A12%3Bi%3A19%3Bi%3A13%3Bi%3A406%3B%7Ds%3A12%3A%22bookTypeName%22%3Bs%3A12%3A%22%C8%A8%CD%FE%B2%A9%B2%CA%B9%AB%CB%BE%22%3B%7D%7D; FirstOKURL=http%3A//www.okooo.com/jingcai/; First_Source=www.okooo.com; PHPSESSID=5397c55438a8970d0999a254ea68c46fb22cea3d; Last_Source=http%3A//www.okooo.com/jingcai/; zq_companytype_odds=AuthoriteBooks; DRUPAL_LOGGED_IN=Y; IMUserID=23960264; IMUserName=zll735508686; OkAutoUuid=deee31c6a2342121bf490b99d26ca39d; OkMsIndex=6; isInvitePurview=0; UWord=5c4ef7418f370ef8ab80183d01d55e118eb; LastUrl=; Hm_lvt_5ffc07c2ca2eda4cc1c4d8e50804c94b=1520217259; Hm_lpvt_5ffc07c2ca2eda4cc1c4d8e50804c94b=1520232754; __utma=56961525.325478833.1510387948.1520217258.1520231584.3; __utmb=56961525.12.9.1520232763874; __utmc=56961525; __utmz=56961525.1510387948.1.1.utmcsr=(direct)|utmccn=(direct)|utmcmd=(none)',
                   'Host': 'www.okooo.com',
                   'Referer': arr,
                   'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36)',
                   'X-Requested-With': 'XMLHttpRequest'
                   }
    res = ''
    try:
        req = urllib.request.Request(url='%sajax/?companytype=AuthoriteBooks&page=0&trnum=6&type=0' % arr, headers=header_dict)
        res = urllib.request.urlopen(req)
        res = res.read().decode()
    except Exception as e:
        logger.info(e)
        return res
    return res


def insertExcel(sheet, url, nowTime, result, count):
    rb = open_workbook('D:/投注信息/2家/%s_2.xls' % nowTime)
    wb = copy(rb)
    ws = wb.get_sheet(sheet)
    ws.write(count, 0, url)
    ws.write(count, 7, time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    ws.write(count + 1, 0, '公司名')
    ws.write(count + 1, 1, '主')
    ws.write(count + 1, 2, '平')
    ws.write(count + 1, 3, '客')
    ws.write(count + 1, 4, '主')
    ws.write(count + 1, 5, '平')
    ws.write(count + 1, 6, '客')
    ws.write(count + 1, 7, '主')
    ws.write(count + 1, 8, '平')
    ws.write(count + 1, 9, '客')
    ws.write(count + 1, 10, '值')
    # arr1 = []
    # arr2 = []
    # arr3 = []
    if len(result) > 0:
        for j in range(len(result)):
            for k in range(len(result[j])):
                # if k == 4:
                #     arr1.append(float(result[j][k]))
                # elif k == 5:
                #     arr2.append(float(result[j][k]))
                # elif k == 6:
                #     arr3.append(float(result[j][k]))
                ws.write(j + count + 2, k, result[j][k])
    # arr1.remove(min(arr1))
    # arr2.remove(min(arr2))
    # arr3.remove(min(arr3))
    # arr1.remove(max(arr1))
    # arr2.remove(max(arr2))
    # arr3.remove(max(arr3))
    # re1 = numpy.std(arr1, ddof = 1) / (sum(arr1) / len(arr1)) * 100
    # re2 = numpy.std(arr2, ddof = 1) / (sum(arr2) / len(arr2)) * 100
    # re3 = numpy.std(arr3, ddof = 1) / (sum(arr3) / len(arr3)) * 100
    # ws.write(count + 4, 4, '%.2f' % re1 + '%')
    # ws.write(count + 4, 5, '%.2f' % re2 + '%')
    # ws.write(count + 4, 6, '%.2f' % re3 + '%')
    wb.save('D:/投注信息/2家/%s_2.xls' % nowTime)


if __name__ == '__main__':
    socket.setdefaulttimeout(30)
    url_arr, info = getAddress()
    c = 0
    while c < 100:
        if len(url_arr) > 0:
            logger.info('连接成功!')
            nowTime = time.strftime("%Y%m%d", time.localtime())
            createTable(info)
            count = get_row(nowTime)
            startGetDate(url_arr, count, nowTime)
            break
        else:
            c += 1
            logger.info('1秒后重新连接，第%s次' % c)
            time.sleep(1)
            url_arr, info = getAddress()
    else:
        logger.info('无法连接，请稍后重试！')
