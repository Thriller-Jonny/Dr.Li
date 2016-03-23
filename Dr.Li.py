#__author__ = 'Jonny'
# coding=utf-8

import xdrlib, sys
import xlrd
import xlwt
import urllib
import json

DATA_SOURCE_URL = "http://www.gffunds.com.cn/apistore/JsonService?method=Fund&op=queryFundByGFCategory&service=BaseInfo"
FILTER_FUND_CODE = ["000117", "000529", "000167", "000477","270025","000567","000550","001468","001763","000747","270021","270007","000942","000968","001133","001064","001180","001469","001460","000826","000992","001189","270004","270014","000475"]

#创建Excel表格
def create_excel():
    try:
        data = xlwt.Workbook()
        return data
    except Exception(e):
        print(str(e))

#向表格内写数据
def excel_table_write(columns, content):
    data = create_excel()
    sheet = data.add_sheet('Fund from guangfa', cell_overwrite_ok = True)
    for col in columns:
        sheet.write(0,columns.index(col),col.decode('utf-8'))

    for rownum in range(0, len(content)):
        for colnum in range(0, len(content[rownum])):
            sheet.write(rownum+1, colnum, content[rownum][colnum])

    data.save('Dr.Li.xls')


#解析网站返回的html数据
def parseHtml(html):
    html = html.decode('utf-8')
    htmlJson = json.loads(html)
    if htmlJson["errormsg"] != "Success!":
        return

    dataTuple = htmlJson["data"]
    listData = []
    for dataLine in dataTuple:
        code = dataLine["FUNDCODE"]
        for condition in FILTER_FUND_CODE:
            if condition == code:
                fund = [dataLine["WEBPRODUCTNAME"],code,dataLine["WEBNAVUNIT"],dataLine["WEBNAVACCUMULATED"],dataLine["DAYINCREMENTRATE"],dataLine["YIELDTHISY"],dataLine["YIELDLASTY"]]
                listData.append(fund)
                break

    return listData

#获取网站html
def getHtml():
    page = urllib.urlopen(DATA_SOURCE_URL)
    html = page.read()
    return html

#加载数据
def loadData():
    htmlContent = getHtml()
    return parseHtml(htmlContent)

# 开始运行
def run():
    cols = ["基金名称", "基金代码", "基金净值", "累计净值", "日涨跌", "今年以来", "过去一年"]
    fundData = loadData()
    excel_table_write(cols,fundData)

if __name__=="__main__":
    run()
