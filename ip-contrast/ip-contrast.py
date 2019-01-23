# -*- coding: utf-8 -*-

import configparser
import os
import xlrd
import time
import openpyxl
import shutil

config = configparser.ConfigParser()
configFileName = "config.ini"
t0 = 0


# ip 转换： int 转 str
def int2ip(x):
    return ".".join([str(int(x / (256 ** i) % 256)) for i in range(3, -1, -1)])


# ip 转换： str 转 int
def ip2int(x):
    return sum([256 ** j * int(i) for j, i in enumerate(x.split(".")[::-1])])


# ip 字符串 10.10.10.10/32 转成数组: ip/mask/ip_start/ip_end
def ipParse(ipStr):
    if ipStr == "":
        return [None, None, None, None]
    maskLen = 32
    if "/" in ipStr:
        ipStr, maskLen = ipStr.split("/")
    # print(ipStr, maskLen)
    mask = 0xFFFFFFFF << (32 - int(maskLen))
    # print("mask:", mask)
    ipInt = ip2int(ipStr)
    ipStart = ipInt & mask & 0xFFFFFFFF
    ipEnd = ipInt | ~mask & 0xFFFFFFFF
    return [ipInt, mask, ipStart, ipEnd]


# ip, mask 转成 字符串 10.10.10.10/32
def ipImport(ipInt, mask):
    pass


# 检查配置格式是否正确
def checkConfig(configFileName):
    global config
    config.read(configFileName)
    sections = config.sections()
    sectionLens = len(sections)
    if sectionLens < 4:
        return False
    # 每个 section 中的 options 的个数必须是相同的
    for i in range(1, sectionLens):
        if i < sectionLens - 1:
            n1 = len(config.options(sections[i]))
            n2 = len(config.options(sections[i + 1]))
            if n1 == 0 | n1 != n2:
                return False
    # 可进一步判断 options 名是否一致
    pass
    return True


# 释放默认配置
def writeConfig(configFileName):
    global config
    config.add_section("对比文件名")
    config.set("对比文件名", "省内资管", "IP地址.")
    config.set("对比文件名", "集团", "-IP地址-")
    config.set("对比文件名", "工信部备案", "fpxxList")
    config.add_section("省内资管")
    config.set("省内资管", "before", "1")
    config.set("省内资管", "ip", "IP地址")
    config.set("省内资管", "field2", "联系人姓名(客户侧)")
    config.set("省内资管", "field3", "联系电话(客户侧)")
    config.set("省内资管", "field4", "分配使用时间")
    config.set("省内资管", "field5", "单位详细地址")
    config.set("省内资管", "field6", "联系人邮箱(客户侧)")
    config.set("省内资管", "field7", "单位名称/具体业务信息")
    config.add_section("集团")
    config.set("集团", "before", "3")
    config.set("集团", "ip", "网段名称")
    config.set("集团", "field2", "联系人姓名(客户侧)")
    config.set("集团", "field3", "联系人电话(客户侧)")
    config.set("集团", "field4", "分配使用时间")
    config.set("集团", "field5", "单位详细地址")
    config.set("集团", "field6", "联系人邮箱(客户侧)")
    config.set("集团", "field7", "单位名称/具体业务信息")
    config.add_section("工信部备案")
    config.set("工信部备案", "before", "1")
    config.set("工信部备案", "ip", "起始IP;终止IP")
    config.set("工信部备案", "field2", "联系人姓名")
    config.set("工信部备案", "field3", "联系人电话")
    config.set("工信部备案", "field4", "分配日期")
    config.set("工信部备案", "field5", "单位详细地址")
    config.set("工信部备案", "field6", "联系人电子邮件")
    config.set("工信部备案", "field7", "使用单位名称")
    with open(configFileName, "w") as configFile:
        configFile.write(
            "# Author: Xianda\n\n# 本配置为一致性检查工具的配置\n# 如删除本配置文件，重新生成的配置文件即为默认配置\n# 如需修改配置，修改本文件后直接保存即可\n\n######\n\n# before 表示数据的起始行（列名占用的行数）\n# 若 IP 地址字段分为起始IP、终止IP的，`ip` 字段中用“;”(英文分号)隔开\n# 程序会依次对 field 字段进行对比并输出对比结果\n# field 字段有变化可直接在此增删改，满足一一对应即可\n\n\n\n"
        )
        config.write(configFile)


# 初始化配置
def initConfig():
    global configFileName
    if os.path.exists(configFileName):
        if checkConfig(configFileName):
            # 配置检查通过，开始对比数据
            return contrast()
        else:
            print('Error:', "配置中 field 字段的个数不一致，请核对！")
            os.system("pause")
    else:
        # 释放默认配置，开始对比数据
        writeConfig(configFileName)
        return contrast()


# 读取配置
def readConfig(section, option):
    global config
    global configFileName
    config.read(configFileName)
    value = config.get(section, option)
    return value


# 匹配导出数据文件名
def matchedFileName():
    global config
    global configFileName
    fileName = {}
    config.read(configFileName)
    for option in config.options("对比文件名"):
        value = config.get("对比文件名", option)
        for f in os.listdir():
            if value in f:
                fileName[option] = f
        if option not in fileName:
            print('Warning:', '未找到', option, '的导出数据，将不进行该数据的一致性检查。')
    return fileName


# 生成中间文件，转换导出数据为每一行一个 ip
def generateTemp(fileName):
    if fileName == {}:
        print('Error:', '未识别到excel文件。')
        exit()
    global config
    # print(list(fileName.keys())[0])
    # print(config.options(list(fileName.keys())[0]))

    # 基于 ip 列 生成中间文件，并补充 ipStart、ipEnd 列

    # 智能识别、生成配置、并输出中间文件，在一次 fileName 的 for 循环中完成
    options = {}
    colNames = {}
    ipNames = {}
    sheet = {}
    for k, v in fileName.items():
        xls = xlrd.open_workbook(v)
        # 获取最后一个 sheet
        sheet[k] = xls.sheet_by_index(len(xls.sheet_names()) - 1)
        # 默认将第一行做为列名所在的行
        Row0 = sheet[k].row_values(0)
        options[k] = {}
        options[k]['fieldCols'] = {}
        options[k]['ipCols'] = []
        # options[k]['output'] = []
        colNames[k] = []
        ipNames[k] = []
        fields = config.options(k)

        # 遍历配置文件中要对比的列名 记录要对比的字段所在的列
        for field in fields:
            configValue = config.get(k, field)
            if 'before' == field:
                # 数据起始行（标题所占的行数）默认为 1
                options[k]['before'] = configValue
            else:
                colFounded = False
                # 遍历导出数据的第一行单元格
                for i in range(0, len(Row0)):
                    cellValue = str(Row0[i].strip())
                    if cellValue in configValue:
                        colFounded = True
                        if 'ip' == field:
                            # 记录 ip 所在的列
                            options[k]['ipCols'].append(i)
                            ipNames[k].append(cellValue)
                        elif 'field' in field:
                            # 记录要对比的字段所在的列
                            options[k]['fieldCols'][field] = i
                            colNames[k].append(cellValue)
                        # 此处不能加 break，否则匹配到`起始IP`即跳出 for 循环，无法匹配`结束IP`
                # 若根据配置未找到对应的列：给出提示，结束运行
                if not colFounded:
                    print('在', k, '中，未找到数据列：', field +
                          '【', configValue, '】，工具即将退出。')
                    exit()
        # colNames[k] = ['ipStart', 'ipEnd'] + colNames[k]
        colNames[k].extend(['ipStart', 'ipEnd'])
        colNames[k].extend(ipNames[k])

    # print('Info:', "options", options, '\n')
    # print('Info:', 'colNames', colNames, '\n')
    # os.mkdir('\\_temp')
    configPath = os.getcwd() + '\\Xianda\\ipContrast\\'
    # tempdir=>C:\\ProgramData
    # tempdir = os.environ.get('ALLUSERSPROFILE')
    # tempdir=>%USERPROFILE%/Local/Temp, C:/Users/XXXX/AppData/Local/Temp
    tempdir = os.environ.get('TEMP')
    if os.path.exists(tempdir):
        configPath = tempdir + '\\Xianda\\ipContrast\\'
    if not os.path.exists(configPath):
        os.makedirs(configPath)
    # 创建中间文件
    note = '对比结果在第一个sheet页最后 '+ str(len(colNames)-1) +' 列'
    author = '--Xianda'
    # wb = openpyxl.Workbook(write_only=True)
    # ws0 = wb.create_sheet('说明')
    # from openpyxl.worksheet.write_only import WriteOnlyCell
    # from openpyxl.styles import Font
    # cell = WriteOnlyCell(ws0, value=note)
    # cell.font = Font(size=18, color='FF0000')
    # ws0.append([None])
    # ws0.append([None, cell])
    # ws0.append([None])
    # cell = WriteOnlyCell(ws0, value=author)
    # cell.font = Font(size=11, color='8060ee', bold=True)
    # ws0.append([None, None, cell])
    # cell = WriteOnlyCell(ws0, value=str(options))
    # cell.font=Font(size=11, color='666600')
    # ws0.append([None, cell])
    # cell=WriteOnlyCell(ws0, value=str(colNames))
    # cell.font=Font(size=11, color='666600')
    # ws0.append([None, cell])
    wb = openpyxl.Workbook()
    ws0 = wb.active
    ws0.title = '说明'
    ws0.column_dimensions['B'].width = 120
    ws0['B2'] = note
    ws0['C3'] = author
    line = 6
    count = len(fileName) + 1
    for i in fileName:
        ws0['A'+str(line)] = i
        ws0['B'+str(line)] = str(options[i])
        ws0['A'+str(line+count)] = i
        ws0['B'+str(line+count)] = str(colNames[i])
        line += 1
    # ws.merge_cells('A2:G3')
    currDate = time.strftime('%Y-%m-%d', time.localtime())
    wrName = configPath + currDate + '-temp-' + str(time.time()) + '.xlsx'
    isFirstSheet = True
    for k, v in fileName.items():
        print('Info:','Temp File is being Generated for:', k, v)  # debug
        # 为每个对比文件创建中间文件的一个sheet
        ws = wb.create_sheet(k)
        if isFirstSheet:
            extendTitle = ['与' + x + '一致' for x in fileName.keys()]
            extendTitle.pop(0)
            title = colNames[k] +\
                ['预留1', '预留2', 'concatenate'] + extendTitle
        else:
            title = colNames[k]
        ws.append(title)
        # 遍历每一行
        for row in range(int(options[k]['before']), sheet[k].nrows):
            tempRow = []
            rowValues = sheet[k].row_values(row)
            fieldCols = options[k]['fieldCols']
            if not isFirstSheet:
                if rowValues[int(fieldCols['field5'])] == '':
                    continue
            currentRow = str(ws.max_row+1)
            strings = '=CONCATENATE(A'+currentRow+',"-",B'+currentRow+',"-",C' + \
                currentRow+',"-",D'+currentRow+',"-",E'+currentRow+',"-",F'+currentRow+')'
            for i in fieldCols:
                tempRow.append(rowValues[int(fieldCols[i])])
            ipCols = options[k]['ipCols']
            if len(ipCols) == 1:
                ipStr = rowValues[ipCols[0]]
                tempRow.extend([ipParse(ipStr)[2], ipParse(
                    ipStr)[3], ipStr, None, None, strings])
            elif len(ipCols) == 2:
                # 导出数据自带起始 IP 结束IP
                ip1Str = rowValues[ipCols[0]]
                ip2Str = rowValues[ipCols[1]]
                ip1 = ip2int(ip1Str)
                ip2 = ip2int(ip2Str)
                if ip1 < ip2:
                    tempRow.extend([ip1, ip2, ip1Str, ip2Str, None, strings])
                else:
                    tempRow.extend([ip2, ip1, ip2Str, ip1Str, None, strings])
            else:
                print('Warning:','IP 列识别错误', k, currentRow, '行')
                return
            if isFirstSheet:
                formula = [x for x in fileName.keys()]
                formula.pop(0)
                tempRow.extend(['=VLOOKUP(G' + currentRow + ',' +
                                x + '!G:L,6,0)=L' + currentRow for x in formula])
            # print('tempRow', tempRow) # debug
            # 过滤客户地址为空的数据
            # if tempRow[3] == '':
            #     continue
            # else:
            ws.append(tempRow)
        isFirstSheet = False
        ws.auto_filter.ref = "A1:N"+str(ws.max_row)
        # ws.auto_filter.add_sort_condition('G2:G'+str(ws.max_row))
    wb.save(wrName)
    return wrName


def copyfile(srcfile, dstfile):
    if not os.path.isfile(srcfile):
        print("%s not exist!" % (srcfile))
    else:
        fpath, fname = os.path.split(dstfile)  # 分离文件名和路径
        if not os.path.exists(fpath):
            os.makedirs(fpath)  # 创建路径
        shutil.copyfile(srcfile, dstfile)  # 复制文件
        # print "copy %s -> %s"%( srcfile,dstfile)


def contrast():
    # 获取导出文件的文件名
    fileName = matchedFileName()
    # print('Info:',"fileName", fileName, '\n')
    tempFile = generateTemp(fileName)
    print('Info:','tempFile', tempFile)
    return tempFile

# @测试 configparser
def _test_configparser():
    global config
    config.read("config.ini")
    sections = config.sections()
    for x in sections:
        options = config.options(x)
        for y in options:
            print(config.get(x, y))


if __name__ == "__main__":
    t0 = time.time()
    result = initConfig()
    t = time.time()-t0
    print('\n本工具执行用时：%2.4f s' % t)
    print('\n- 即将打开对比结果 -\n')
    os.system('pause')
    os.system('explorer /select, '+result)
