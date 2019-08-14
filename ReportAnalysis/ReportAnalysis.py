#! /usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import re
import time
import xlwt

default_ana_dir = './TestReport'

def WriteRow(sheet, rowindex, row):
    for index in range(0, len(row)):
        if isinstance(row[index], str) and row[index].startswith('='):
            sheet.write(rowindex, index, xlwt.Formula(row[index][1:]), xlStyle)
        elif isinstance(row[index], xlwt.ExcelFormula.Formula):
            sheet.write(rowindex, index, row[index], xlLinkStyle)
        else:
            sheet.write(rowindex, index, row[index], xlStyle)


def WriteCol(sheet, rowstart, colstart, data):
    for index in range(0, len(data)):
        if isinstance(data[index], dict):
            coladd = 0
            for (key,value) in data[index].items():
                sheet.write_merge(
                    rowstart + index, rowstart + index,
                    colstart + coladd, colstart + coladd + value - 1,
                    key,
                    xlStyle
                )
                coladd = coladd + value
        elif isinstance(data[index], list):
            for cindex in range(0, len(data[index])):
                sheet.write(rowstart + index, colstart + cindex, data[index][cindex], xlStyle)
        elif isinstance(data[index], str):
            sheet.write(rowstart + index, colstart, data[index], xlStyle)

def InitializeWorkBook():
    global reportWorkBook, xlStyle, xlLinkStyle, reportSheet
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
    reportWorkBook = xlwt.Workbook()
    reportSheet = reportWorkBook.add_sheet('SDRAM Test Report', cell_overwrite_ok=True)

if __name__ == '__main__':
    Reports = {}
    if not os.path.isdir(default_ana_dir):
        ana_dir = './'
    else:
        ana_dir = default_ana_dir + ("" if default_ana_dir.endswith("/") else "/")
    print('将检查 %s 下所有的报告文件！' % ana_dir)
    ana_files = os.listdir(ana_dir)
    for ana_file in ana_files:
        if os.path.isfile(ana_dir + ana_file) and ana_file.endswith('.txt'):
            print('正在分析 %s ...' % ana_file)
            Reports[ana_file] = {}
            Reports[ana_file]["Performance"] = {}
            Reports[ana_file]["PerformanceCount"] = 0
            ana_performance = ""
            ana_order = "main"
            ana_f = open(ana_dir + ana_file)
            for line in iter(ana_f):
                if ana_order == "main":
                    clock_re = re.match(r'(\S+?)\s*Clock:\s*([0-9]* MHz)', line)
                    if clock_re:
                        Reports[ana_file][clock_re.group(1) + " Freq"] = clock_re.group(2)
                    if "SDRAM Test Start" in line:
                        ana_order = "func"
                elif ana_order == "func":
                    error_re = re.match(r'.*?Erro Count: ([0-9]*)', line)
                    big_re = re.match(r'Big-Endian Test Result: ([0-9]{8})', line)
                    little_re = re.match(r'Little-Endian Test Result: ([0-9]{8})', line)
                    if error_re:
                        Reports[ana_file]["Erro Count"] = error_re.group(1)
                    if big_re:
                        Reports[ana_file]["Big-Endian"] = "Pass" if big_re.group(1) == "34127856" else "Fail"
                    if little_re:
                        Reports[ana_file]["Little-Endian"] = "Pass" if little_re.group(1) == "78563412" else "Fail"
                        ana_order = "perform"
                elif ana_order == "perform":
                    if ana_performance == "":
                        performance_re = re.match(r'^.*Performance Tests (.*?)==', line)
                        if performance_re:
                            ana_performance = performance_re.group(1)
                            Reports[ana_file]["Performance"][ana_performance] = {}
                            Reports[ana_file]["PerformanceCount"] = Reports[ana_file]["PerformanceCount"] + 1
                            print("正在分析 %s 性能测试" % ana_performance)
                    else:
                        write_time_re = re.match(r'(.*) Write Time: ([0-9]* ms)', line)
                        write_speed_re = re.match(r'(.*) Write Speed: ([0-9.]* MB/s)', line)
                        read_time_re = re.match(r'(.*) Read Time: ([0-9]* ms)', line)
                        read_speed_re = re.match(r'(.*) Read Speed: ([0-9.]* MB/s)', line)
                        performance_end_re = re.match(r'^.*Performance Test End', line)
                        if performance_end_re:
                            ana_performance = ""
                        if write_time_re:
                            Reports[ana_file]["Performance"][ana_performance][write_time_re.group(1)] = {}
                            Reports[ana_file]["Performance"][ana_performance][write_time_re.group(1)]["Write Time"] = write_time_re.group(2)
                        if write_speed_re:
                            Reports[ana_file]["Performance"][ana_performance][write_speed_re.group(1)]["Write Speed"] = write_speed_re.group(2)
                        if read_time_re:
                            # Reports[ana_file]["Performance"][ana_performance][read_time_re.group(1)] = {}
                            Reports[ana_file]["Performance"][ana_performance][read_time_re.group(1)]["Read Time"] = read_time_re.group(2)
                        if read_speed_re:
                            Reports[ana_file]["Performance"][ana_performance][read_speed_re.group(1)]["Read Speed"] = read_speed_re.group(2)
    InitializeWorkBook()
    MainHeader = [
        "Test File Name",
        "PLL Freq",
        "CPU Freq",
        "HCLK Freq",
        "PCLK Freq",
    ]
    FunctionHeader = [
        "Erro Count",
        "Big-Endian",
        "Little-Endian"
    ]
    PerformanceHeader = [
        "Test Name",
        "RoW"
    ]
    WriteCol(reportSheet, 0, 0, MainHeader)
    WriteCol(reportSheet, len(MainHeader), 0, FunctionHeader)
    colstart = 1
    for (testfile,content) in Reports.items():
        writecol = []
        writecolCount = content["PerformanceCount"] * 2
        writecol.append({ testfile : writecolCount })
        for index in range(1, len(MainHeader)):
            writecol.append({content[MainHeader[index]] : writecolCount})
        for index in range(0, len(FunctionHeader)):
            writecol.append({ content[FunctionHeader[index]] : writecolCount})
        pnamecol = {}
        prwheader = []
        pdatas = []
        pdatas.append(pnamecol)
        pdatas.append(prwheader)
        for (pname,pcontent) in content["Performance"].items():
            pnamecol[pname] = 2
            prwheader.append("Write")
            prwheader.append("Read")
            for (dname,dcontent) in pcontent.items():
                try:
                    if PerformanceHeader.index(dname) >= len(pdatas):
                        pdatas.append([])
                except ValueError:
                    PerformanceHeader.append(dname)
                    pdatas.append([])
                pdatas[PerformanceHeader.index(dname)].append(dcontent["Write Speed"])
                pdatas[PerformanceHeader.index(dname)].append(dcontent["Read Speed"])
        for pdata in pdatas:
            writecol.append(pdata)
        WriteCol(reportSheet, 0, colstart, writecol)
        colstart = colstart + writecolCount
    WriteCol(reportSheet, len(MainHeader) + len(FunctionHeader), 0, PerformanceHeader)
    reportWorkBook.save("SDRAM Report-" + time.strftime("%Y%m%d-%H%M%S") + ".xls")