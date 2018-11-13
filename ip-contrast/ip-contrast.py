# -*- coding: utf-8 -*-

import configparser
import os
import xlrd


config = configparser.ConfigParser()

# 检查配置格式是否正确
def checkConfig(configFileName):
    global config
    config.read(configFileName)
    sections = config.sections()
    sectionLens = len(sections)
    # 每个 section 中的 options 的个数必须是相同的
    for i in range(0, sectionLens):
        if i < sectionLens-1:
            if len(config.options(sections[i])) != len(config.options(sections[i + 1])):
                return False
    # 可进一步判断 options 名是否一致
    pass
    return True


# 释放默认配置
def writeConfig(configFileName):
    global config
    config.add_section("省内资管")
    config.set("省内资管","filed1", "IP地址")
    config.set("省内资管","filed2", "联系人姓名(客户侧)")
    config.set("省内资管","filed3", "联系电话(客户侧)")
    config.set("省内资管","filed4", "分配使用时间")
    config.set("省内资管","filed5", "单位详细地址")
    config.set("省内资管","filed6", "联系人邮箱(客户侧)")
    config.set("省内资管","filed7", "单位名称/具体业务信息")
    config.add_section("集团")
    config.set("集团","filed1", "网段名称")
    config.set("集团","filed2", "联系人姓名(客户侧)")
    config.set("集团","filed3", "联系电话(客户侧)")
    config.set("集团","filed4", "分配使用时间")
    config.set("集团","filed5", "单位详细地址")
    config.set("集团","filed6", "联系人邮箱(客户侧)")
    config.set("集团","filed7", "单位名称/具体业务信息")
    config.add_section("工信部备案")
    config.set("工信部备案","filed1", "起始IP;终止IP")
    config.set("工信部备案","filed2", "联系人姓名")
    config.set("工信部备案","filed3", "联系电话")
    config.set("工信部备案","filed4", "分配使用时间")
    config.set("工信部备案","filed5", "单位详细地址")
    config.set("工信部备案","filed6", "联系人邮箱")
    config.set("工信部备案","filed7", "单位名称")
    with open(configFileName, "w") as configFile:
        configFile.write('# Author: Xianda\n\n# 本配置为一致性检查工具的配置\n# 如需恢复默认请删除本文件，重新生成的配置即为默认配置\n# 如需修改配置，修改本文件后直接保存即可\n\n######\n\n# 若 IP 地址字段分为起始IP、终止IP的，`filed1` 字段中用“;”(英文分号)隔开\n# 程序会依次对 filed 字段进行对比并输出对比结果\n# filed 字段有变化可直接在此增删改，满足一一对应即可\n\n\n\n')
        config.write(configFile)


# 初始化配置
def initConfig():
    configFileName = "config.ini"
    if os.path.exists(configFileName):
        if checkConfig(configFileName):
            # 配置检查通过，开始对比数据
            contrast()
        else:
            print("配置中 filed 字段的个数不一致，请核对！")
            os.system("pause")
    else:
        # 释放默认配置，开始对比数据
        writeConfig(configFileName)
        contrast()


# 对比数据
def contrast():
    # 定义模板字段，根据第一个字段为查询索引，对比其余字段是否一致
    # 打开文件
    zgg = xlrd.open_workbook("demo/ip-zg.xlsx")
    # jt = xlrd.open_workbook(r'ip-jt.xlsx')
    # print zgg.name
    print(zgg.sheet_names())
    sheet2 = zgg.sheet_by_index(0)  # sheet索引从0开始
    # sheet2 = zg.sheet_by_name('sheet2')
    # sheet的名称，行数，列数
    print(sheet2.name, sheet2.nrows, sheet2.ncols)
    # 获取整行和整列的值（数组）
    rows = sheet2.row_values(3)  # 获取第四行内容
    # cols = sheet2.col_values(2) # 获取第三列内容
    print(rows)
    # print cols
    # 获取单元格内容
    print(str(sheet2.cell(1, 0).value))
    print(int(sheet2.cell_value(1, 0)))
    print(sheet2.row(1)[0].value)

def ttt():
    global config
    config.read("config.ini")
    sections = config.sections()
    for x in sections:
        options = config.options(x)
        for y in options:
            print(config.get(x, y))

if __name__ == "__main__":
    initConfig()
    ttt()

