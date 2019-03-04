#!/usr/bin/env python
# -*- coding: utf-8 -*-

from datetime import datetime
import configparser
import os
import platform
import sys
import ctypes
import numpy
import pandas
import traceback


class ip_contrast(object):
    """ IP 信息对比"""

    def __init__(self):
        self.t0 = datetime.now()
        self.__version__ = '0.9.0'
        self.config = configparser.ConfigParser()
        self.configFileName = 'config_%s.ini' % self.__version__
        self.DEBUG_FILE = 'debug_log.txt'

        print('****欢迎使用一致性检查工具 %s\n' % __version__)
        print(self.now(), '**开始运行')
        """初始化配置"""
        if os.path.exists(self.configFileName):
            if not self.checkConfig():
                print('Warning:', '配置有误，已恢复默认', self.configFileName)
                self.writeConfig()
        else:
            print('Info:', '配置已更新', self.configFileName)
            # 释放默认配置
            self.writeConfig()
        print('当前加载的配置：', self.configFileName)
        self._matchFileNames()

    def now(self):
        """当前时间的字符串"""
        return datetime.strftime(datetime.now(), '%Y-%m-%d %H:%M:%S')

    def checkConfig(self):
        """检查配置格式是否正确"""
        self.config.read(self.configFileName)
        sections = self.config.sections()
        sectionLens = len(sections)
        if sectionLens < 5:
            return False
        # 每个 section 中的 options 的个数必须是相同的
        SETTING_LENS = 2
        for i in range(SETTING_LENS, sectionLens):
            if i < sectionLens - SETTING_LENS:
                n1 = len(self.config.options(sections[i]))
                n2 = len(self.config.options(sections[i + 1]))
                if n1 == 0 | n1 != n2:
                    return False
        # 可进一步判断 options 名是否一致
        pass
        return True

    def writeConfig(self):
        config = self.config
        config['common'] = {
            'version': self.__version__
        }
        config['对比文件名'] = {
            '省内资管': 'IP地址.',
            '集团': '-IP地址-',
            '工信部备案': 'fpxxList',
        }
        config['省内资管'] = {
            'tag1': '所属地市',
            'ip': 'IP地址',
            'starttime': '分配使用时间',
            'field3': '联系人姓名(客户侧)',
            'field4': '联系电话(客户侧)',
            'field5': '单位详细地址',
            'field6': '联系人邮箱(客户侧)',
            'field7': '单位名称/具体业务信息',
        }
        config['集团'] = {
            'tag1': '所属地市',
            'ip': '网段名称',
            'starttime': '分配使用时间',
            'field3': '联系人姓名(客户侧)',
            'field4': '联系人电话(客户侧)',
            'field5': '单位详细地址',
            'field6': '联系人邮箱(客户侧)',
            'field7': '单位名称/具体业务信息',
        }
        config['工信部备案'] = {
            'tag1': '所属地',
            'ip': '起始IP;终止IP',
            'starttime': '分配日期',
            'field3': '联系人姓名',
            'field4': '联系人电话',
            'field5': '单位详细地址',
            'field6': '联系人电子邮件',
            'field7': '使用单位名称',
        }
        with open(self.configFileName, 'w') as configFile:

            title = '''# Author: Xianda

    # 本配置为一致性检查工具的配置
    # 如删除本配置文件，重新生成的配置文件即为默认配置
    # 如需修改配置，修改本文件后直接保存即可

    ######

    # before 表示数据的起始行（列名占用的行数）。现已自动识别（配置无效）
    # 若 IP 地址字段分为起始IP、终止IP的，`ip` 字段中用“;”(英文分号)隔开
    # 程序会依次对 field 字段进行对比并输出对比结果
    # field 字段有变化可直接在此增删改，满足一一对应即可



    '''
            configFile.write(title)
            config.write(configFile)

    def _matchFileNames(self):
        """匹配导出数据文件名"""
        fileName = {}
        self.config.read(self.configFileName)
        for option in self.config.options('对比文件名'):
            value = self.config.get('对比文件名', option)
            fileName[option] = []
            for f in os.listdir():
                if value in f and not '~' in f:
                    fileName[option].append(f)
            if not fileName[option]:
                fileName.pop(option)
                print('Warning:', '未找到', option, '的导出数据，将不进行该数据的一致性检查。')
        self.fileName = fileName

    def recognizeOptions(self, k, xls):
        """识别配置"""
        fields = self.config.options(k)
        options = {}
        options['fieldCols'] = {}
        options['ipCols'] = []
        colNames = []
        ipNames = []
        # 默认获取第一个 sheet 的第一行做为列名所在的行
        Row0 = xls.sheet_by_index(0).row_values(0)
        # 遍历配置文件中要对比的列名 记录要对比的字段所在的列
        for field in fields:
            configValue = self.config.get(k, field)
            colFounded = False
            # 遍历导出数据的第一行单元格
            for i in range(len(Row0)):
                cellValue = str(Row0[i].strip().strip('*'))
                if cellValue in configValue:  # 为了满足： `起始IP` in `起始IP;终止IP`
                    colFounded = True
                    if 'ip' == field:
                        # 记录 ip 所在的列
                        options['ipCols'].append(i)
                        ipNames.append(Row0[i])
                    else:
                        # 记录要对比的字段所在的列
                        options['fieldCols'][field] = i
                        colNames.append(Row0[i])
                    # 此处不能加 break，否则匹配到`起始IP`即跳出 for 循环，无法匹配`结束IP`
            # 若根据配置未找到对应的列：给出提示，结束运行
            if not colFounded:
                if not field in ['tag1']:
                    print('在%s中，未找到数据列：【%s】，工具即将退出。' % (k, configValue))
                    pause()
                    exit()
        self.options, self.colNames, self.ipNames = options, colNames, ipNames
        return options, colNames, ipNames

    def run(self):
        """执行对比"""
        # self.recognizeOptions()


if 'Windows' in platform.system():
    __stdInputHandle = -10
    __stdOutputHandle = -11
    __stdErrorHandle = -12
    __foreGroundBLUE = 0x09
    __foreGroundGREEN = 0x0a
    __foreGroundRED = 0x0c
    __foreGroundYELLOW = 0x0e
    stdOutHandle = ctypes.windll.kernel32.GetStdHandle(__stdOutputHandle)

    def setCmdColor(color, handle=stdOutHandle):
        return ctypes.windll.kernel32.SetConsoleTextAttribute(handle, color)

    def resetCmdColor():
        setCmdColor(__foreGroundRED | __foreGroundGREEN | __foreGroundBLUE)

    def printBlue(msg):
        setCmdColor(__foreGroundBLUE)
        sys.stdout.write(msg + '\n')
        resetCmdColor()

    def printGreen(msg):
        setCmdColor(__foreGroundGREEN)
        sys.stdout.write(msg + '\n')
        resetCmdColor()

    def printRed(msg):
        setCmdColor(__foreGroundRED)
        # sys.stdout.write(msg + '\n')
        print(msg)
        resetCmdColor()

    def printYellow(msg):
        setCmdColor(__foreGroundYELLOW)
        sys.stdout.write(msg + '\n')
        resetCmdColor()

    def pause():
        os.system('pause')

    def locateFile(file):
        print('\n- 即将跳转到输出结果 -\n')
        pause()
        os.system('explorer /select, ' + file)

else:
    STYLE = {
        'fore': {
            'red': 31,
            'green': 32,
            'yellow': 33,
            'blue': 34,
        }
    }

    def UseStyle(msg, mode='', fore='', back='40'):
        fore = '%s' % STYLE['fore'][fore] if fore in STYLE['fore'] else ''
        style = ';'.join([s for s in [mode, fore, back] if s])
        style = '\033[%sm' % style if style else ''
        end = '\033[%sm' % 0 if style else ''
        return '%s%s%s' % (style, msg, end)

    def printRed(msg):
        print(UseStyle(msg, fore='red'))

    def printGreen(msg):
        print(UseStyle(msg, fore='green'))

    def printYellow(msg):
        print(UseStyle(msg, fore='yellow'))

    def printBlue(msg):
        print(UseStyle(msg, fore='blue'))

    def pause():
        pass

    def locateFile(file):
        print('\n- 结果文件保存为：%s' % file)
        print('刷新页面可到 `__output` 文件夹中下载查看。')

if __name__ == '__main__':
    ip_tables = ip_contrast()
    try:
        ip_tables.run()
    except Exception as err:
        printRed('Error:')
        traceback.print_exc(file=open(DEBUG_FILE, 'w'))
        printRed(err)
        printBlue('出错啦。请反馈目录中的 debug_log.txt 文件内容')

    t = datetime.now() - t0
    print('\n执行用时：%2.4f s' % t.total_seconds())
    locateFile(result)
