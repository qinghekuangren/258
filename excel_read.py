"""

"""

import os
import xlrd
import xlwt
import pandas as pd
from datetime import datetime


def solve_report(input_file=None, id=0, col_num=1):

    wb = xlrd.open_workbook(input_file)
    print(wb)
    sheet_name_list = wb.sheet_names()
    print(sheet_name_list)
    all_data = []
    index = 0
    for s in sheet_name_list:
        print(index)

        sh = wb.sheet_by_name(s)
        tmp = []
        for rownum in range(sh.nrows):
            tmp.append(sh.row_values(rownum))
        for i in range(len(tmp)):
            print(i, tmp[i])
        all_data.append(tmp)
        print('='*50)
        index += 1
    # for i in range(len(all_data)):
    #     print(i, all_data[i])
    res = {}
    res['序号'] = col_num
    res['试验编号'] = str(int(all_data[0][4][1]))
    res['检验单位'] = all_data[0][2][0]
    res['工程名称'] = all_data[0][9][1]
    res['样品名称'] = all_data[0][6][1]
    res['检验项目'] = all_data[0][7][1]
    res['样品数量'] = all_data[0][8][5]
    res['最大玻璃尺寸长'] = 1210
    res['最大玻璃尺寸宽'] = 1210
    res['最大玻璃尺寸厚'] = 1210
    res['五金件状况'] = '良好'
    res["送检单位"] = 'P =2.5P ='
    res["镶嵌方式"] = '湿法'
    res['框扇密封材料'] = all_data[0][17][5]
    res['玻璃密封材料'] = all_data[0][16][5]
    res['气密工程设计值'] = '/'
    res['检验类别'] = 2523
    res["样品规格"] = '/'
    res["样品型号"] = '/'

    if id == 0:
        res['样品编号'] = str(int(all_data[1][3][11]))
        res['生产单位'] = all_data[0][8][1]
        res['委托单位'] = all_data[0][5][1]
        res['玻璃厚度'] = int(all_data[0][30][0])
        res['玻璃品种'] = '白色平玻璃'
        res["玻璃镶嵌材料"] = ''
        res["正工程渗漏量f"] = ''
        res["正工程渗漏量z"] = ''
        res["气体温度"] = int(all_data[1][26][6])
        res['室内气压'] = int(all_data[1][25][6])

        sh = wb.sheet_by_name('报告')
        da = xlrd.xldate_as_tuple(sh.cell_value(13, 5), wb.datemode)
        da1 = datetime(*da)
        strda = da1.strftime('%Y%m%d')
        # print(strda)
        res['检验日期'] = strda
        res['试件面积'] = all_data[0][15][2]
        res['淋水量'] = 120
        res['测点间距1'] = int(all_data[0][30][5])
        res['测点间距2'] = 50000
        res['测点间距'] = "测点间距"
        res['开启缝长'] = all_data[0][15][5]
        res['开启形式'] = '平开'
        res['加压方式'] = ''

        res['雨水渗漏加压方式'] = '稳定加压'
        res['挠度最大值'] = 5.87
        res['工程设计值'] = 0
        res['雨水工程设计值'] = 200

        res['雨水渗漏'] = ""
        res['负工程渗漏量f'] = ""
        res['负工程渗漏量z'] = ""
        res["气密结论"] = ""
        res["气密结论-"] = ""
        res["水密结论"] = ""
        res["风压结论"] = ""
        res["200PaB"] = ""
        res["400PaB"] = ""
        res["600PaB"] = ""
        res["800PaB"] = ""
        res["1000PaB"] = ""
        res["1200PaB"] = ""
        res["1400PaB"] = ""
        res["1600PaB"] = ""
        res["1800PaB"] = ""
        res["2000PaB"] = ""
        res["-200PaB"] = ""
        res["-400PaB"] = ""
        res["-600PaB"] = ""
        res["-800PaB"] = ""

        res["-1000PaB"] = ""
        res["-1200PaB"] = ""
        res["-1400PaB"] = ""
        res["-1600PaB"] = ""
        res["-1800PaB"] = ""
        res["-2000PaB"] = ""

        res["200PaA"] = all_data[10][7][3]
        res["400PaA"] = all_data[10][8][3]
        res["600PaA"] = all_data[10][9][3]
        res["800PaA"] = all_data[10][10][3]
        res["1000PaA"] = all_data[10][11][3]
        res["1200PaA"] = all_data[10][12][3]
        res["1400PaA"] = all_data[10][13][3]
        res["1600PaA"] = all_data[10][14][3]
        res["1800PaA"] = all_data[10][15][3]
        res["2000PaA"] = all_data[10][16][3]

        res["B200Pa"] = all_data[10][7][5]
        res["B400Pa"] = all_data[10][8][5]
        res["B600Pa"] = all_data[10][9][5]
        res["B800Pa"] = all_data[10][10][5]
        res["B1000Pa"] = all_data[10][11][5]
        res["B1200Pa"] = all_data[10][12][5]
        res["B1400Pa"] = all_data[10][13][5]
        res["B1600Pa"] = all_data[10][14][5]
        res["B1800Pa"] = all_data[10][15][5]
        res["B2000Pa"] = all_data[10][16][5]

        res["-B200Pa"] = all_data[10][7][16]
        res["-B400Pa"] = all_data[10][8][16]
        res["-B600Pa"] = all_data[10][9][16]
        res["-B800Pa"] = all_data[10][10][16]
        res["-B1000Pa"] = all_data[10][11][16]
        res["-B1200Pa"] = all_data[10][12][16]
        res["-B1400Pa"] = all_data[10][13][16]
        res["-B1600Pa"] = all_data[10][14][16]
        res["-B1800Pa"] = all_data[10][15][16]
        res["-B2000Pa"] = all_data[10][16][16]

        res["C200Pa"] = all_data[10][7][7]
        res["C400Pa"] = all_data[10][8][7]
        res["C600Pa"] = all_data[10][9][7]
        res["C800Pa"] = all_data[10][10][7]
        res["C1000Pa"] = all_data[10][11][7]
        res["C1200Pa"] = all_data[10][12][7]
        res["C1400Pa"] = all_data[10][13][7]
        res["C1600Pa"] = all_data[10][14][7]
        res["C1800Pa"] = all_data[10][15][7]
        res["C2000Pa"] = all_data[10][16][7]

        res["-C200Pa"] = all_data[10][7][18]
        res["-C400Pa"] = all_data[10][8][18]
        res["-C600Pa"] = all_data[10][9][18]
        res["-C800Pa"] = all_data[10][10][18]
        res["-C1000Pa"] = all_data[10][11][18]
        res["-C1200Pa"] = all_data[10][12][18]
        res["-C1400Pa"] = all_data[10][13][18]
        res["-C1600Pa"] = all_data[10][14][18]
        res["-C1800Pa"] = all_data[10][15][18]
        res["-C2000Pa"] = all_data[10][16][18]

        res["P1"] = 3960
        res["-P1"] = -2746

        res['P3残余变形'] = all_data[10][26][4]
        res['Pmax残余变形'] = all_data[10][26][10]

        res["持续时间1"] = 600
        res["持续时间2"] = 300
        res["持续时间3"] = 300
        res["持续时间4"] = 300
        res["持续时间5"] = 122
        res["持续时间6"] = 0
        res["持续时间7"] = 0
        res["持续时间8"] = 0
        res["持续时间9"] = 0
        res["持续时间10"] = 0
        res["持续时间11"] = 0
        res["持续时间12"] = 0

        res['稳定渗漏1'] = all_data[7][7][15]
        res['稳定渗漏2'] = all_data[7][8][15]
        res['稳定渗漏3'] = all_data[7][9][15]
        res['稳定渗漏4'] = all_data[7][10][15]
        res['稳定渗漏5'] = all_data[7][11][15]
        res['稳定渗漏6'] = all_data[7][12][15]
        res['稳定渗漏7'] = all_data[7][13][15]
        res['稳定渗漏8'] = all_data[7][14][15]
        res['稳定渗漏9'] = all_data[7][15][15]
        res['稳定渗漏10'] = all_data[7][16][15]
        res['稳定渗漏11'] = ""
        res['稳定渗漏12'] = '分级等级'
        res['分级等级'] = ''

        res['升压流量10F'] = all_data[1][9][2]
        res['升压流量30F'] = all_data[1][10][2]
        res['升压流量50F'] = all_data[1][11][2]
        res['升压流量70F'] = all_data[1][12][2]
        res['升压流量100F'] = all_data[1][13][2]
        res['升压流量150F'] = all_data[1][14][2]

        res['降压流量10F'] = all_data[1][9][4]
        res['降压流量30F'] = all_data[1][10][4]
        res['降压流量50F'] = all_data[1][11][4]
        res['降压流量70F'] = all_data[1][12][4]
        res['降压流量100F'] = all_data[1][13][4]
        res['降压流量150F'] = all_data[1][14][4]

        res['升压流量10z'] = all_data[1][9][8]
        res['升压流量30z'] = all_data[1][10][8]
        res['升压流量50z'] = all_data[1][11][8]
        res['升压流量70z'] = all_data[1][12][8]
        res['升压流量100z'] = all_data[1][13][8]
        res['升压流量150z'] = all_data[1][14][8]

        res['降压流量10z'] = all_data[1][9][10]
        res['降压流量30z'] = all_data[1][10][10]
        res['降压流量50z'] = all_data[1][11][10]
        res['降压流量70z'] = all_data[1][12][10]
        res['降压流量100z'] = all_data[1][13][10]
        res['降压流量150z'] = all_data[1][14][10]

        res['q1'] = 0
        res['q2'] = 0
        res['-q1'] = 0.1
        res['-q2'] = 0.2

        res['正压破坏压力差值'] = ""
        res['综合等级'] = ""

        res["-200PaA"] = all_data[11][7][14]
        res["-400PaA"] = all_data[11][8][14]
        res["-600PaA"] = all_data[11][9][14]
        res["-800PaA"] = all_data[11][10][14]
        res["-1000PaA"] = all_data[11][11][14]
        res["-1200PaA"] = all_data[11][12][14]
        res["-1400PaA"] = all_data[11][13][14]
        res["-1600PaA"] = all_data[11][14][14]
        res["-1800PaA"] = all_data[11][15][14]
        res["-2000PaA"] = all_data[11][16][14]

        res['p2'] = 3000
        res['-p2'] = -3000
        res['p3'] = 5000
        res['-p3'] = -5000

        res['抗风压记录1'] = '请记录'
        res['抗风压记录2'] = '请记录'
        res['p1a'] = ''
        res['p2a'] = ''
        res['p3a'] = ''
        res['负压破坏压力差值'] = ''
        res['单锁点检测标志'] = ''
        res['抗风压检测方式'] = '定级检测'
        res['检测方式'] = '定级检测'
        res['渗漏标志'] = '检测压力差值(Pa):'
        res['严重渗漏'] = '250'
        res['分级指标'] = '200'
        res['渗透综合等级+'] = ''
        res['渗透综合等级-'] = ''

        res['负升压流量10F'] = all_data[2][9][2]
        res['负升压流量30F'] = all_data[2][10][2]
        res['负升压流量50F'] = all_data[2][11][2]
        res['负升压流量70F'] = all_data[2][12][2]
        res['负升压流量100F'] = all_data[2][13][2]
        res['负升压流量150F'] = all_data[2][14][2]

        res['负降压流量10F'] = all_data[2][9][4]
        res['负降压流量30F'] = all_data[2][10][4]
        res['负降压流量50F'] = all_data[2][11][4]
        res['负降压流量70F'] = all_data[2][12][4]
        res['负降压流量100F'] = all_data[2][13][4]
        res['负降压流量150F'] = all_data[2][14][4]

        res['负降压流量10z'] = all_data[2][9][10]
        res['负降压流量30z'] = all_data[2][10][10]
        res['负降压流量50z'] = all_data[2][11][10]
        res['负降压流量70z'] = all_data[2][12][10]
        res['负降压流量100z'] = all_data[2][13][10]
        res['负降压流量150z'] = all_data[2][14][10]

        res['负升压流量10z'] = all_data[2][9][8]
        res['负升压流量30z'] = all_data[2][10][8]
        res['负升压流量50z'] = all_data[2][11][8]
        res['负升压流量70z'] = all_data[2][12][8]
        res['负升压流量100z'] = all_data[2][13][8]
        res['负升压流量150z'] = all_data[2][14][8]

    elif id == 1:
        res['样品编号'] = str(int(all_data[3][3][11]))
        res['生产单位'] = all_data[0][8][1]
        res['委托单位'] = all_data[0][5][1]
        res['玻璃厚度'] = int(all_data[0][30][0])
        res['玻璃品种'] = '白色平玻璃'
        res["玻璃镶嵌材料"] = ''
        res["正工程渗漏量f"] = ''
        res["正工程渗漏量z"] = ''
        res["气体温度"] = int(all_data[3][26][6])
        res['室内气压'] = int(all_data[3][25][6])

        sh = wb.sheet_by_name('报告')
        da = xlrd.xldate_as_tuple(sh.cell_value(13, 5), wb.datemode)
        da1 = datetime(*da)
        strda = da1.strftime('%Y%m%d')
        # print(strda)
        res['检验日期'] = strda
        res['试件面积'] = all_data[0][15][2]
        res['淋水量'] = 120
        res['测点间距1'] = int(all_data[0][30][5])
        res['测点间距2'] = 50000
        res['测点间距'] = "测点间距"
        res['开启缝长'] = all_data[0][15][5]
        res['开启形式'] = '平开'
        res['加压方式'] = ''

        res['雨水渗漏加压方式'] = '稳定加压'
        res['挠度最大值'] = 5.87
        res['工程设计值'] = 0
        res['雨水工程设计值'] = 200

        res['雨水渗漏'] = ""
        res['负工程渗漏量f'] = ""
        res['负工程渗漏量z'] = ""
        res["气密结论"] = ""
        res["气密结论-"] = ""
        res["水密结论"] = ""
        res["风压结论"] = ""
        res["200PaB"] = ""
        res["400PaB"] = ""
        res["600PaB"] = ""
        res["800PaB"] = ""
        res["1000PaB"] = ""
        res["1200PaB"] = ""
        res["1400PaB"] = ""
        res["1600PaB"] = ""
        res["1800PaB"] = ""
        res["2000PaB"] = ""
        res["-200PaB"] = ""
        res["-400PaB"] = ""
        res["-600PaB"] = ""
        res["-800PaB"] = ""

        res["-1000PaB"] = ""
        res["-1200PaB"] = ""
        res["-1400PaB"] = ""
        res["-1600PaB"] = ""
        res["-1800PaB"] = ""
        res["-2000PaB"] = ""

        res["200PaA"] = all_data[11][7][3]
        res["400PaA"] = all_data[11][8][3]
        res["600PaA"] = all_data[11][9][3]
        res["800PaA"] = all_data[11][10][3]
        res["1000PaA"] = all_data[11][11][3]
        res["1200PaA"] = all_data[11][12][3]
        res["1400PaA"] = all_data[11][13][3]
        res["1600PaA"] = all_data[11][14][3]
        res["1800PaA"] = all_data[11][15][3]
        res["2000PaA"] = all_data[11][16][3]

        res["B200Pa"] = all_data[11][7][5]
        res["B400Pa"] = all_data[11][8][5]
        res["B600Pa"] = all_data[11][9][5]
        res["B800Pa"] = all_data[11][10][5]
        res["B1000Pa"] = all_data[11][11][5]
        res["B1200Pa"] = all_data[11][12][5]
        res["B1400Pa"] = all_data[11][13][5]
        res["B1600Pa"] = all_data[11][14][5]
        res["B1800Pa"] = all_data[11][15][5]
        res["B2000Pa"] = all_data[11][16][5]

        res["-B200Pa"] = all_data[11][7][16]
        res["-B400Pa"] = all_data[11][8][16]
        res["-B600Pa"] = all_data[11][9][16]
        res["-B800Pa"] = all_data[11][10][16]
        res["-B1000Pa"] = all_data[11][11][16]
        res["-B1200Pa"] = all_data[11][12][16]
        res["-B1400Pa"] = all_data[11][13][16]
        res["-B1600Pa"] = all_data[11][14][16]
        res["-B1800Pa"] = all_data[11][15][16]
        res["-B2000Pa"] = all_data[11][16][16]

        res["C200Pa"] = all_data[11][7][7]
        res["C400Pa"] = all_data[11][8][7]
        res["C600Pa"] = all_data[11][9][7]
        res["C800Pa"] = all_data[11][10][7]
        res["C1000Pa"] = all_data[11][11][7]
        res["C1200Pa"] = all_data[11][12][7]
        res["C1400Pa"] = all_data[11][13][7]
        res["C1600Pa"] = all_data[11][14][7]
        res["C1800Pa"] = all_data[11][15][7]
        res["C2000Pa"] = all_data[11][16][7]

        res["-C200Pa"] = all_data[11][7][18]
        res["-C400Pa"] = all_data[11][8][18]
        res["-C600Pa"] = all_data[11][9][18]
        res["-C800Pa"] = all_data[11][10][18]
        res["-C1000Pa"] = all_data[11][11][18]
        res["-C1200Pa"] = all_data[11][12][18]
        res["-C1400Pa"] = all_data[11][13][18]
        res["-C1600Pa"] = all_data[11][14][18]
        res["-C1800Pa"] = all_data[11][15][18]
        res["-C2000Pa"] = all_data[11][16][18]

        res["P1"] = 3960
        res["-P1"] = -2746

        res['P3残余变形'] = all_data[11][26][4]
        res['Pmax残余变形'] = all_data[11][26][10]

        res["持续时间1"] = 600
        res["持续时间2"] = 300
        res["持续时间3"] = 300
        res["持续时间4"] = 300
        res["持续时间5"] = 122
        res["持续时间6"] = 0
        res["持续时间7"] = 0
        res["持续时间8"] = 0
        res["持续时间9"] = 0
        res["持续时间10"] = 0
        res["持续时间11"] = 0
        res["持续时间12"] = 0

        res['稳定渗漏1'] = all_data[8][7][15]
        res['稳定渗漏2'] = all_data[8][8][15]
        res['稳定渗漏3'] = all_data[8][9][15]
        res['稳定渗漏4'] = all_data[8][10][15]
        res['稳定渗漏5'] = all_data[8][11][15]
        res['稳定渗漏6'] = all_data[8][12][15]
        res['稳定渗漏7'] = all_data[8][13][15]
        res['稳定渗漏8'] = all_data[8][14][15]
        res['稳定渗漏9'] = all_data[8][15][15]
        res['稳定渗漏10'] = all_data[8][16][15]
        res['稳定渗漏11'] = ""
        res['稳定渗漏12'] = '分级等级'
        res['分级等级'] = ''

        res['升压流量10F'] = all_data[3][9][2]
        res['升压流量30F'] = all_data[3][10][2]
        res['升压流量50F'] = all_data[3][11][2]
        res['升压流量70F'] = all_data[3][12][2]
        res['升压流量100F'] = all_data[3][13][2]
        res['升压流量150F'] = all_data[3][14][2]

        res['降压流量10F'] = all_data[3][9][4]
        res['降压流量30F'] = all_data[3][10][4]
        res['降压流量50F'] = all_data[3][11][4]
        res['降压流量70F'] = all_data[3][12][4]
        res['降压流量100F'] = all_data[3][13][4]
        res['降压流量150F'] = all_data[3][14][4]

        res['升压流量10z'] = all_data[3][9][8]
        res['升压流量30z'] = all_data[3][10][8]
        res['升压流量50z'] = all_data[3][11][8]
        res['升压流量70z'] = all_data[3][12][8]
        res['升压流量100z'] = all_data[3][13][8]
        res['升压流量150z'] = all_data[3][14][8]

        res['降压流量10z'] = all_data[3][9][10]
        res['降压流量30z'] = all_data[3][10][10]
        res['降压流量50z'] = all_data[3][11][10]
        res['降压流量70z'] = all_data[3][12][10]
        res['降压流量100z'] = all_data[3][13][10]
        res['降压流量150z'] = all_data[3][14][10]

        res['q1'] = 0
        res['q2'] = 0
        res['-q1'] = 0.1
        res['-q2'] = 0.2

        res['正压破坏压力差值'] = ""
        res['综合等级'] = ""

        res["-200PaA"] = all_data[11][7][14]
        res["-400PaA"] = all_data[11][8][14]
        res["-600PaA"] = all_data[11][9][14]
        res["-800PaA"] = all_data[11][10][14]
        res["-1000PaA"] = all_data[11][11][14]
        res["-1200PaA"] = all_data[11][12][14]
        res["-1400PaA"] = all_data[11][13][14]
        res["-1600PaA"] = all_data[11][14][14]
        res["-1800PaA"] = all_data[11][15][14]
        res["-2000PaA"] = all_data[11][16][14]

        res['p2'] = 3000
        res['-p2'] = -3000
        res['p3'] = 5000
        res['-p3'] = -5000

        res['抗风压记录1'] = '请记录'
        res['抗风压记录2'] = '请记录'
        res['p1a'] = ''
        res['p2a'] = ''
        res['p3a'] = ''
        res['负压破坏压力差值'] = ''
        res['单锁点检测标志'] = ''
        res['抗风压检测方式'] = '定级检测'
        res['检测方式'] = '定级检测'
        res['渗漏标志'] = '检测压力差值(Pa):'
        res['严重渗漏'] = '250'
        res['分级指标'] = '200'
        res['渗透综合等级+'] = ''
        res['渗透综合等级-'] = ''

        res['负升压流量10F'] = all_data[4][9][2]
        res['负升压流量30F'] = all_data[4][10][2]
        res['负升压流量50F'] = all_data[4][11][2]
        res['负升压流量70F'] = all_data[4][12][2]
        res['负升压流量100F'] = all_data[4][13][2]
        res['负升压流量150F'] = all_data[4][14][2]


        res['负降压流量10F'] = all_data[4][9][4]
        res['负降压流量30F'] = all_data[4][10][4]
        res['负降压流量50F'] = all_data[4][11][4]
        res['负降压流量70F'] = all_data[4][12][4]
        res['负降压流量100F'] = all_data[4][13][4]
        res['负降压流量150F'] = all_data[4][14][4]

        res['负降压流量10z'] = all_data[4][9][10]
        res['负降压流量30z'] = all_data[4][10][10]
        res['负降压流量50z'] = all_data[4][11][10]
        res['负降压流量70z'] = all_data[4][12][10]
        res['负降压流量100z'] = all_data[4][13][10]
        res['负降压流量150z'] = all_data[4][14][10]

        res['负升压流量10z'] = all_data[4][9][8]
        res['负升压流量30z'] = all_data[4][10][8]
        res['负升压流量50z'] = all_data[4][11][8]
        res['负升压流量70z'] = all_data[4][12][8]
        res['负升压流量100z'] = all_data[4][13][8]
        res['负升压流量150z'] = all_data[4][14][8]

    elif id == 2:
        res['样品编号'] = str(int(all_data[5][3][11]))
        res['生产单位'] = all_data[0][8][1]
        res['委托单位'] = all_data[0][5][1]
        res['玻璃厚度'] = int(all_data[0][30][0])
        res['玻璃品种'] = '白色平玻璃'
        res["玻璃镶嵌材料"] = ''
        res["正工程渗漏量f"] = ''
        res["正工程渗漏量z"] = ''
        res["气体温度"] = int(all_data[5][26][6])
        res['室内气压'] = int(all_data[5][25][6])

        sh = wb.sheet_by_name('报告')
        da = xlrd.xldate_as_tuple(sh.cell_value(13, 5), wb.datemode)
        da1 = datetime(*da)
        strda = da1.strftime('%Y%m%d')
        # print(strda)
        res['检验日期'] = strda
        res['试件面积'] = all_data[0][15][2]
        res['淋水量'] = 120
        res['测点间距1'] = int(all_data[0][30][5])
        res['测点间距2'] = 50000
        res['测点间距'] = "测点间距"
        res['开启缝长'] = all_data[0][15][5]
        res['开启形式'] = '平开'
        res['加压方式'] = ''

        res['雨水渗漏加压方式'] = '稳定加压'
        res['挠度最大值'] = 5.87
        res['工程设计值'] = 0
        res['雨水工程设计值'] = 200

        res['雨水渗漏'] = ""
        res['负工程渗漏量f'] = ""
        res['负工程渗漏量z'] = ""
        res["气密结论"] = ""
        res["气密结论-"] = ""
        res["水密结论"] = ""
        res["风压结论"] = ""
        res["200PaB"] = ""
        res["400PaB"] = ""
        res["600PaB"] = ""
        res["800PaB"] = ""
        res["1000PaB"] = ""
        res["1200PaB"] = ""
        res["1400PaB"] = ""
        res["1600PaB"] = ""
        res["1800PaB"] = ""
        res["2000PaB"] = ""
        res["-200PaB"] = ""
        res["-400PaB"] = ""
        res["-600PaB"] = ""
        res["-800PaB"] = ""

        res["-1000PaB"] = ""
        res["-1200PaB"] = ""
        res["-1400PaB"] = ""
        res["-1600PaB"] = ""
        res["-1800PaB"] = ""
        res["-2000PaB"] = ""

        res["200PaA"] = all_data[12][7][3]
        res["400PaA"] = all_data[12][8][3]
        res["600PaA"] = all_data[12][9][3]
        res["800PaA"] = all_data[12][10][3]
        res["1000PaA"] = all_data[12][11][3]
        res["1200PaA"] = all_data[12][12][3]
        res["1400PaA"] = all_data[12][13][3]
        res["1600PaA"] = all_data[12][14][3]
        res["1800PaA"] = all_data[12][15][3]
        res["2000PaA"] = all_data[12][16][3]

        res["B200Pa"] = all_data[12][7][5]
        res["B400Pa"] = all_data[12][8][5]
        res["B600Pa"] = all_data[12][9][5]
        res["B800Pa"] = all_data[12][10][5]
        res["B1000Pa"] = all_data[12][11][5]
        res["B1200Pa"] = all_data[12][12][5]
        res["B1400Pa"] = all_data[12][13][5]
        res["B1600Pa"] = all_data[12][14][5]
        res["B1800Pa"] = all_data[12][15][5]
        res["B2000Pa"] = all_data[12][16][5]

        res["-B200Pa"] = all_data[12][7][16]
        res["-B400Pa"] = all_data[12][8][16]
        res["-B600Pa"] = all_data[12][9][16]
        res["-B800Pa"] = all_data[12][10][16]
        res["-B1000Pa"] = all_data[12][11][16]
        res["-B1200Pa"] = all_data[12][12][16]
        res["-B1400Pa"] = all_data[12][13][16]
        res["-B1600Pa"] = all_data[12][14][16]
        res["-B1800Pa"] = all_data[12][15][16]
        res["-B2000Pa"] = all_data[12][16][16]

        res["C200Pa"] = all_data[12][7][7]
        res["C400Pa"] = all_data[12][8][7]
        res["C600Pa"] = all_data[12][9][7]
        res["C800Pa"] = all_data[12][10][7]
        res["C1000Pa"] = all_data[12][11][7]
        res["C1200Pa"] = all_data[12][12][7]
        res["C1400Pa"] = all_data[12][13][7]
        res["C1600Pa"] = all_data[12][14][7]
        res["C1800Pa"] = all_data[12][15][7]
        res["C2000Pa"] = all_data[12][16][7]

        res["-C200Pa"] = all_data[12][7][18]
        res["-C400Pa"] = all_data[12][8][18]
        res["-C600Pa"] = all_data[12][9][18]
        res["-C800Pa"] = all_data[12][10][18]
        res["-C1000Pa"] = all_data[12][11][18]
        res["-C1200Pa"] = all_data[12][12][18]
        res["-C1400Pa"] = all_data[12][13][18]
        res["-C1600Pa"] = all_data[12][14][18]
        res["-C1800Pa"] = all_data[12][15][18]
        res["-C2000Pa"] = all_data[12][16][18]

        res["P1"] = 3960
        res["-P1"] = -2746

        res['P3残余变形'] = all_data[12][26][4]
        res['Pmax残余变形'] = all_data[12][26][10]

        res["持续时间1"] = 600
        res["持续时间2"] = 300
        res["持续时间3"] = 300
        res["持续时间4"] = 300
        res["持续时间5"] = 122
        res["持续时间6"] = 0
        res["持续时间7"] = 0
        res["持续时间8"] = 0
        res["持续时间9"] = 0
        res["持续时间10"] = 0
        res["持续时间11"] = 0
        res["持续时间12"] = 0

        res['稳定渗漏1'] = all_data[9][7][15]
        res['稳定渗漏2'] = all_data[9][8][15]
        res['稳定渗漏3'] = all_data[9][9][15]
        res['稳定渗漏4'] = all_data[9][10][15]
        res['稳定渗漏5'] = all_data[9][11][15]
        res['稳定渗漏6'] = all_data[9][12][15]
        res['稳定渗漏7'] = all_data[9][13][15]
        res['稳定渗漏8'] = all_data[9][14][15]
        res['稳定渗漏9'] = all_data[9][15][15]
        res['稳定渗漏10'] = all_data[9][16][15]
        res['稳定渗漏11'] = ""
        res['稳定渗漏12'] = '分级等级'
        res['分级等级'] = ''

        res['升压流量10F'] = all_data[5][9][2]
        res['升压流量30F'] = all_data[5][10][2]
        res['升压流量50F'] = all_data[5][11][2]
        res['升压流量70F'] = all_data[5][12][2]
        res['升压流量100F'] = all_data[5][13][2]
        res['升压流量150F'] = all_data[5][14][2]

        res['降压流量10F'] = all_data[5][9][4]
        res['降压流量30F'] = all_data[5][10][4]
        res['降压流量50F'] = all_data[5][11][4]
        res['降压流量70F'] = all_data[5][12][4]
        res['降压流量100F'] = all_data[5][13][4]
        res['降压流量150F'] = all_data[5][14][4]

        res['升压流量10z'] = all_data[5][9][8]
        res['升压流量30z'] = all_data[5][10][8]
        res['升压流量50z'] = all_data[5][11][8]
        res['升压流量70z'] = all_data[5][12][8]
        res['升压流量100z'] = all_data[5][13][8]
        res['升压流量150z'] = all_data[5][14][8]

        res['降压流量10z'] = all_data[5][9][10]
        res['降压流量30z'] = all_data[5][10][10]
        res['降压流量50z'] = all_data[5][11][10]
        res['降压流量70z'] = all_data[5][12][10]
        res['降压流量100z'] = all_data[5][13][10]
        res['降压流量150z'] = all_data[5][14][10]

        res['q1'] = 0
        res['q2'] = 0
        res['-q1'] = 0.1
        res['-q2'] = 0.2

        res['正压破坏压力差值'] = ""
        res['综合等级'] = ""

        res["-200PaA"] = all_data[12][7][14]
        res["-400PaA"] = all_data[12][8][14]
        res["-600PaA"] = all_data[12][9][14]
        res["-800PaA"] = all_data[12][10][14]
        res["-1000PaA"] = all_data[12][11][14]
        res["-1200PaA"] = all_data[12][12][14]
        res["-1400PaA"] = all_data[12][13][14]
        res["-1600PaA"] = all_data[12][14][14]
        res["-1800PaA"] = all_data[12][15][14]
        res["-2000PaA"] = all_data[12][16][14]

        res['p2'] = 3000
        res['-p2'] = -3000
        res['p3'] = 5000
        res['-p3'] = -5000

        res['抗风压记录1'] = '请记录'
        res['抗风压记录2'] = '请记录'
        res['p1a'] = ''
        res['p2a'] = ''
        res['p3a'] = ''
        res['负压破坏压力差值'] = ''
        res['单锁点检测标志'] = ''
        res['抗风压检测方式'] = '定级检测'
        res['检测方式'] = '定级检测'
        res['渗漏标志'] = '检测压力差值(Pa):'
        res['严重渗漏'] = '250'
        res['分级指标'] = '200'
        res['渗透综合等级+'] = ''
        res['渗透综合等级-'] = ''

        res['负升压流量10F'] = all_data[6][9][2]
        res['负升压流量30F'] = all_data[6][10][2]
        res['负升压流量50F'] = all_data[6][11][2]
        res['负升压流量70F'] = all_data[6][12][2]
        res['负升压流量100F'] = all_data[6][13][2]
        res['负升压流量150F'] = all_data[6][14][2]

        res['负降压流量10F'] = all_data[6][9][4]
        res['负降压流量30F'] = all_data[6][10][4]
        res['负降压流量50F'] = all_data[6][11][4]
        res['负降压流量70F'] = all_data[6][12][4]
        res['负降压流量100F'] = all_data[6][13][4]
        res['负降压流量150F'] = all_data[6][14][4]

        res['负降压流量10z'] = all_data[6][9][10]
        res['负降压流量30z'] = all_data[6][10][10]
        res['负降压流量50z'] = all_data[6][11][10]
        res['负降压流量70z'] = all_data[6][12][10]
        res['负降压流量100z'] = all_data[6][13][10]
        res['负降压流量150z'] = all_data[6][14][10]

        res['负升压流量10z'] = all_data[6][9][8]
        res['负升压流量30z'] = all_data[6][10][8]
        res['负升压流量50z'] = all_data[6][11][8]
        res['负升压流量70z'] = all_data[6][12][8]
        res['负升压流量100z'] = all_data[6][13][8]
        res['负升压流量150z'] = all_data[6][14][8]

    res['检验依据'] = ''
    res['稳定渗漏说明1'] = '试件未出现渗漏'
    res['稳定渗漏说明2'] = '试件未出现渗漏'
    res['稳定渗漏说明3'] = '试件未出现渗漏'
    res['稳定渗漏说明4'] = '试件未出现渗漏'
    res['稳定渗漏说明5'] = '试件未出现渗漏'
    res['稳定渗漏说明6'] = ''
    res['稳定渗漏说明7'] = ''
    res['稳定渗漏说明8'] = ''
    res['稳定渗漏说明9'] = ''
    res['稳定渗漏说明10'] = ''
    res['稳定渗漏说明11'] = ''
    res['稳定渗漏说明12'] = ''
    res['波动渗漏说明1'] = ''
    res['波动渗漏说明2'] = ''
    res['波动渗漏说明3'] = ''
    res['波动渗漏说明4'] = ''
    res['波动渗漏说明5'] = ''
    res['波动渗漏说明6'] = ''
    res['波动渗漏说明7'] = ''
    res['波动渗漏说明8'] = ''
    res['波动渗漏说明9'] = ''
    res['波动渗漏说明10'] = ''
    res['波动渗漏说明11'] = ''
    res['波动渗漏说明12'] = ''
    return res


if __name__ == '__main__':
    input_dir = 'data'
    col = 1
    res_ = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = res_.add_sheet('res', cell_overwrite_ok=True)
    for file_name in os.listdir(input_dir):
        print(file_name)
        for i in range(3):
            res = solve_report(input_file=input_dir + '/' + file_name, id=i, col_num=col)
            col_list = list(res.keys())
            col_value = list(res.values())
            if col == 1:
                for j in range(len(col_list)):
                    sheet.write(0, j, col_list[j])
            for j in range(len(col_list)):
                sheet.write(col, j, col_value[j])
            col += 1
    savepath = 'res/res1.xls'
    res_.save(savepath)





