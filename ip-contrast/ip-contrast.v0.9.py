#!/usr/bin/env python
# -*- coding: utf-8 -*-

import configparser
import pandas


__version__ = '0.9.0'
configFileName = 'config_%s.ini' % __version__
DEBUG_FILE = 'debug_log.txt'

t0 = datetime.now()


class ip_contrast(object):
    """
    IP 信息对比
    """

    def __init__(self):
        self.config = configparser.ConfigParser()
        pass

    # 检查配置格式是否正确
    def checkConfig(self, configFileName):
        self.config.read(configFileName)
        # _version = config.get('common', 'version', fallback=None)
        # if not _version:
        #     return False
        # _version = [int(i) for i in _version.split('.')]
        # _current_ver = [int(i) for i in __version__.split('.')]
        # if _version < _current_ver:
        #     return False
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


if __name__ == '__main__':
    ip_tables = ip_contrast()
