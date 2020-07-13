#!/usr/bin/env python
# -*-coding:utf-8-*-

' 模块说明：对config.xlsx文件的操作 '
import pathlib

from printAll import PrintAll, getPrinter

__author__ = '袁可庆'
import os
import win32api
import win32con
import sys
import configparser
import time
from robotLibrary.common import xlsx as x
from robotLibrary.common.log import ILog
from robotLibrary.common import jsonOperator as j

logger = ILog(__file__)


class PrinterNotExistError(Exception) :
    pass

class Config :
    '''
    前期费读取及修改配置config.xlsx
    '''

    def __init__(self) :
        logger.info('初始化config.xlsx')
        self._config_path = os.path.dirname(__file__) + "\\config.xlsx"  # config.xlsx的绝对地址
        self._config_file = x.loadWorkBook(self._config_path)  # 载入config.xlsx

    def config(self) :
        '''
        config.xlsx的sheet1配置文件表
        得到一个config_mapping 组键值的映射关系，可以按照self.config_mapping['组名']['键名']的方式得到 值
        :return:
        '''
        logger.info('载入sheet1配置文件表')
        self._sheet1 = x.getSheet(self._config_file, 'Sheet1')
        # 获取组键值的列号
        group_col = 0
        key_col = 0
        val_col = 0
        for i in range(20) :
            val = x.getCellData(self._sheet1, 1, i + 1)
            if val.strip() == '组' :
                group_col = i + 1
            elif val.strip() == '键' :
                key_col = i + 1
            elif val.strip() == '值' :
                val_col = i + 1

        # 获取最大行
        row_num = 1
        key_list = []  # 键列
        while True :
            val = x.getCellData(self._sheet1, row_num, key_col)
            if val != 'None' :
                key_list.append(val)
                row_num += 1
            else :
                break
        max_row = len(key_list)
        # logger.info('最大行%s' % len(key_list))

        # 组 和 本组所包含的行号 的对应关系
        group = []  # 有哪些组，组中跨多行的包含None
        pure_group = []  # 有哪些组
        for i in range(2, max_row + 1) :
            gro = x.getCellData(self._sheet1, i, group_col).strip()
            group.append(gro)
            if gro != 'None' :
                pure_group.append(gro)

        # 记录每个组所包含的行数，长度等于组的个数
        include_row = [[] for i in range(len(pure_group))]
        for i in range(len(pure_group)) :
            for j in range(len(group)) :
                if group[j] == pure_group[i] or group[j] == 'None' :
                    include_row[i].append(j + 2)
                    group[j] = 0
                elif group[j] == 0 :
                    continue
                else :
                    break

        # 初始化一个config字典映射：形式如 config = {'system':{'app_path':'1111', 'client':'330'}, 'pathConfig':{'excel_path':'333', 'errorPath': '444'}}
        self.config_mapping = {i : {} for i in pure_group}
        # 给config赋值 表达出 组 键 值 的映射关系
        for i in range(len(include_row)) :
            for row in include_row[i] :
                gro = pure_group[i]
                key = x.getCellData(self._sheet1, row, key_col).strip()
                val = x.getCellData(self._sheet1, row, val_col).strip()
                self.config_mapping[gro][key] = val
        # logger.info('组-键-值对应关系：%s' % self.config_mapping)

    def sheet1(self) :
        '''
        根据config_mapping 初始化excel中的具体值
        :return:
        '''
        # system
        self.printer = self.config_mapping['printConfig']['printer']
        self.get_printer_flag = self.config_mapping['printConfig']['getPrinter']
        self.file_path = self.config_mapping['printConfig']['filePath']
        self.print_flag = self.config_mapping['printAll']['print']

        # 判断路径是否存在 若不存在给弹窗提示
        if not os.path.exists(self.file_path) :
            error = '文件路径错误'
            win32api.MessageBox(0, error, '配置文件', win32con.MB_ICONWARNING)
            logger.error(error)
            raise Exception(error)

        all_printers = getPrinter()
        if self.printer not in all_printers :
            error_info = "电脑未连接此打印机【%s】！" % self.printer
            logger.error(error_info)
            win32api.MessageBox(0, error_info, '配置文件', win32con.MB_ICONWARNING)
            raise PrinterNotExistError(error_info)

# 使用单例模式易于引用
c = Config()
c.config()
c.sheet1()
del Config

if __name__ == '__main__' :
    # c = Config()
    # c.config()
    # c.sheet1()
    # c.constantConfig()
    # c.statusConfig()
    # c.proConfig()
    # c.excelInit()
    pass
