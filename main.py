#!/usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 2020-07-10 13:02
# @Author  : Jarno_Y
# @Email   : ykq12313@gmail.com
# @File    : main.py
# @Software: PyCharm
'''
Module Introduction
'''
import time
from robotLibrary.common.file import getAllFormatFile, getAllDirectories
from robotLibrary.common.log import ILog

from printAll import PrintAll, getPrinter
from config import c

# 1. HP LaserJet 1020
# 2. HP Color LaserJet M750 PCL 6
logger = ILog(__file__)

if __name__ == '__main__':
    if c.get_printer_flag == "True":
        info = getPrinter()
        for i in range(100):
            time.sleep(1)
    if c.print_flag == "True":
        p = PrintAll(c.printer)
        p.printFileList(c.file_path, 'g')


    # a = getAllDirectories(file_path)
    # print(a)