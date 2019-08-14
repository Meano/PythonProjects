#! /usr/bin/env python3
# -*- coding: utf-8 -*-
import json
import requests
import time
import datetime
import xlwt
import re

xmanSession = requests.Session()
xmanTicket = None
xmanSign = None
xmanCheckList = None
xmanDiffSheetList = None

def LoginXman():
    loginData = {
        "username":"username",
        "password":"password",
        "redirectUrl":"https://redirectUrl",
        "setTicketUrl":"https://setTicketUrl"
    }
    loginUrl = "https://loginUrl"
    requestHeader = {
        "Accept" : "*/*",
        "Accept-Encoding" : "gzip, deflate, br",
        "Accept-Language" : "zh-CN",
        "Cache-Control" : "no-cache",
        "Connection" : "Keep-Alive",
        "Content-Length" : "196",
        'DNT':'1',
        "content-type": "application/json; charset=UTF-8",
        "Host": "usercenter-api.blibee.com",
        "Origin": "https://originDomain",
        "Referer": "https://RefererDomain",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36 Edge/16.16299"
    }

    result = xmanSession.post(
        loginUrl,
        json = loginData,
        headers = requestHeader
    )

    logindata = json.loads(result.content)
    if logindata['ret'] != True:
        print(("登陆失败"))
        return False
    else:
        print(("登陆成功"))
        global xmanSign, xmanTicket
        xmanTicket = logindata['data']['ticket']
        xmanSign = logindata['data']['sign']
        return True

def GetCheckList(CheckDate, Day):
    try:
        CheckDateTo = datetime.date.fromtimestamp(time.mktime(time.strptime(CheckDate, "%Y-%m-%d")))
        CheckDateFrom = CheckDateTo - datetime.timedelta(days = Day)
    except Exception as e:
        print("日期转换有误")
        print(e)
        return None
    print("将查找" + CheckDateFrom.isoformat() + "到" + CheckDateTo.isoformat() + "的盘点差异！")
    result = xmanSession.get( \
        "https://SSOLoginUrl" + \
        xmanTicket + \
        "SSOLoginUrlPara" + \
        xmanSign
    )
    CheckListUrl = "https://CheckListUrl"
    CheckParam = {
        "checkName":"周盘",
        "checkTopic":"WeekCheck",
        "orderStatus":["Finished"],
        "checkDateFrom": CheckDateFrom.isoformat(),
        "checkDateTo": CheckDateTo.isoformat(),
        "page": {
            "pageNo":1,
            "pageSize":200
        }
    }
    result = xmanSession.post(CheckListUrl, json = CheckParam)
    checkList = json.loads(result.content)
    return checkList

xmanWorkBook = None
xlStyle = None
xlLinkStyle = None
checkListSheet = None
xmanScheduleList = {}

def WriteRow(sheet, rowindex, row):
    for index in range(0, len(row)):
        if isinstance(row[index], str) and row[index].startswith('='):
            sheet.write(rowindex, index, xlwt.Formula(row[index][1:]), xlStyle)
        elif isinstance(row[index], xlwt.ExcelFormula.Formula):
            sheet.write(rowindex, index, row[index], xlLinkStyle)
        else:
            sheet.write(rowindex, index, row[index], xlStyle)

def InitializeWorkBook():
    global xmanWorkBook, xlStyle, xlLinkStyle, checkListSheet, xmanDiffSheetList
    xlStyle = xlwt.XFStyle()
    xlLinkStyle = xlwt.XFStyle()
    xlFont = xlwt.Font()
    xlFont.name = '微软雅黑'
    xlFont.blod = False
    xlFont.colour_index = 8
    xlFont.height = 200
    xlStyle.font = xlFont
    xlLinkFont = xlwt.Font()
    xlLinkFont.name = '微软雅黑'
    xlLinkFont.blod = False
    xlLinkFont.colour_index = 30
    xlLinkFont.height = 200
    xlLinkFont.underline = 11
    xlLinkStyle.font = xlLinkFont
    xmanWorkBook = xlwt.Workbook()
    checkListSheet = xmanWorkBook.add_sheet('周盘差异汇总', cell_overwrite_ok=True)
    checkListSheetHeader = ['门店编码', '门店名称', '盘盈数量', '盘盈销售额', '盘亏数量', '盘亏销售额', '盘点差异数', '盘点差异金额', '差异原因']
    WriteRow(checkListSheet, 0, checkListSheetHeader)
    xmanDiffSheetList = {}

def GetScheduleList(scheduleId):
    global xmanScheduleList
    scheduleId = str(scheduleId)
    if not (scheduleId in xmanScheduleList):
        scheduleUrl = "https://scheduleUrl" + scheduleId + "scheduleUrlPara" + str(int(time.time()*1000))
        result = xmanSession.get(scheduleUrl)
        scheduleDetail = json.loads(result.content)['data']['data']
        xmanScheduleList[scheduleId] = scheduleDetail
    return xmanScheduleList[scheduleId]

def FindShopSchedule(shopCode, scheduleData):
    for shopSchedule in scheduleData:
        if shopSchedule['shopCode'] == shopCode:
            return shopSchedule
    return {
        "gainQty" : -1,
        "gainSaleTotalPrice" : {
            "amount" : -1,
            "currency" : "CNY" 
            },
        "loseQty" : -1,
        "loseSaleTotalPrice" : {
            "amount" : -1,
            "currency" : "CNY" 
            }
        }

def GetDiffList(orderId):
    diffUrl = "https://diffUrl"
    diffParam = {
        "orderId" : orderId,
        "page" : {
            "pageNo" : 1,
            "pageSize" : 100
            }
        }
    result = xmanSession.post(diffUrl, json = diffParam)
    diffDetail = json.loads(result.content)
    return diffDetail

if __name__ == '__main__':
    print("盘点整理工具V1.1")
    if not LoginXman():
        print("将在10秒后自动关闭！")
        time.sleep(10)
        exit(1)
    xmanCheckList = None
    while(xmanCheckList == None):
        dateinput = input("请输入检查日期（格式为YYYY-mm-dd 例如：" + datetime.date.today().isoformat() + "），不输入则默认为当天日期：")
        dateinput = dateinput if dateinput else datetime.date.today().isoformat()
        dayinput = input("请输入间隔天数(1~9)不输入则默认为7天：")
        dayinput = int(dayinput) if re.match("[1-9]", dayinput) else 7
        xmanCheckList = GetCheckList(dateinput, dayinput)

    if not xmanCheckList and xmanCheckList['status'] == 0:
        print("获取盘点列表失败！")
        exit(1)

    InitializeWorkBook()
    shopCount = 0
    for shop in xmanCheckList['data']['data'][:] :
        if not (shop['shopCode'].startswith('100') or shop['shopCode'].startswith('123')):
            xmanCheckList['data']['data'].remove(shop)
        else :
            shopCount += 1
            shopSchedule = FindShopSchedule(shop['shopCode'], GetScheduleList(shop['scheduleId']))
            shopName = ''
            TotalDiffQty = shopSchedule['gainSaleTotalPrice']['amount'] + shopSchedule['loseSaleTotalPrice']['amount']
            if TotalDiffQty > 0:
                print("--此门店存在差异，正在整理差异列表")
                diffList = GetDiffList(shop['orderId'])['data']['data']
                sheetCount = 1
                sheetName = shop['shopName'] + "-" + str(sheetCount)
                while sheetName in xmanDiffSheetList:
                    sheetCount += 1
                    sheetName = shop['shopName'] + "-" + str(sheetCount)
                xmanDiffSheetList[sheetName] = {}
                xmanDiffSheetList[sheetName]['checkDate'] = shop['checkDate']
                xmanDiffSheetList[sheetName]['sheet'] = xmanWorkBook.add_sheet(sheetName, cell_overwrite_ok=True)
                xmanDiffSheetList[sheetName]['sheet'].write_merge(0, 0, 0, 8, shop['shopName'] + " (盘点日期:" + shop['checkDate'] + ")", xlStyle)
                diffSheetHeader = ['商品编码', '商品条码', '商品名称', '应盘数', '实盘数', '差异值', '售价', '盘点差异金额', '原因']
                WriteRow(xmanDiffSheetList[sheetName]['sheet'], 1, diffSheetHeader)
                diffRowCount = 1
                for diffProduct in diffList:
                    diffRowCount += 1
                    diffProductRow = [
                        diffProduct['productCode'],
                        diffProduct['barcode'],
                        diffProduct['productName'],
                        diffProduct['inventoryQtyDecimal'],
                        diffProduct['actualCheckQty'],
                        diffProduct['diffQtyDecimal'],
                        diffProduct['saleUnitPrice']['amount'],
                        "=F" + str(diffRowCount + 1) + "*G" + str(diffRowCount + 1)
                    ]
                    WriteRow(xmanDiffSheetList[sheetName]['sheet'], diffRowCount, diffProductRow)
                shopName = xlwt.Formula('HYPERLINK("#\'' + sheetName + '\'!A1";"' + shop['shopName'] + '")')
            else:
                shopName = shop['shopName']
            shopRow = [shop['shopCode'], shopName, shopSchedule['gainQty'], shopSchedule['gainSaleTotalPrice']['amount'], shopSchedule['loseQty'], shopSchedule['loseSaleTotalPrice']['amount']]

            WriteRow(checkListSheet, shopCount, shopRow)
            print("正在处理店铺: " + shop['shopCode'] + ' ' + shop['shopName'])
    ResultFileName = '盘点整理-' + time.strftime("%Y%m%d-%H%M%S") + '.xls'
    xmanWorkBook.save(ResultFileName)
    print('已经保存整理列表到:' + ResultFileName + '！')
    print('10秒后自动退出！')
    time.sleep(10)

