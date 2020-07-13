#!/usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 2020-07-10 10:35
# @Author  : Jarno_Y
# @Email   : ykq12313@gmail.com
# @File    : printAll.py
# @Software: PyCharm
'''
Module Introduction
'''
from robotLibrary.common.file import getAllFormatFile
from robotLibrary.common.log import ILog

import win32api
import win32con
import win32print
import win32ui
from PIL import Image, ImageWin

logger = ILog(__file__)




class PrintAll :
    def __init__(self, printer_name) :
        self.printer_name = printer_name

        # Set printer
        logger.info("正在设置默认打印机为： %s" % self.printer_name)
        win32print.SetDefaultPrinter(self.printer_name)
        self.printer = win32print.GetDefaultPrinter()

    def printFile(self, file_path) :
        '''
        打印单个文件
        :param file_path:
        :return:
        '''
        # Constants for GetDeviceCaps
        # HORZRES / VERTRES = printable area
        HORZRES = 8
        VERTRES = 10
        # LOGPIXELS = dots per inch
        #
        LOGPIXELSX = 88
        LOGPIXELSY = 90
        #
        # PHYSICALWIDTH/HEIGHT = total area
        #
        PHYSICALWIDTH = 110
        PHYSICALHEIGHT = 111
        #
        # PHYSICALOFFSETX/Y = left / top margin
        #
        PHYSICALOFFSETX = 112
        PHYSICALOFFSETY = 113
        # You can only write a Device-independent bitmap
        #  directly to a Windows device context; therefore
        #  we need (for ease) to use the Python Imaging
        #  Library to manipulate the image.
        #
        # Create a device context from a named printer
        #  and assess the printable size of the paper.
        #
        hDC = win32ui.CreateDC()
        hDC.CreatePrinterDC(self.printer)
        printable_area = hDC.GetDeviceCaps(HORZRES), hDC.GetDeviceCaps(VERTRES)
        printer_size = hDC.GetDeviceCaps(PHYSICALWIDTH), hDC.GetDeviceCaps(PHYSICALHEIGHT)
        printer_margins = hDC.GetDeviceCaps(PHYSICALOFFSETX), hDC.GetDeviceCaps(PHYSICALOFFSETY)

        #
        # Open the image, rotate it if it's wider than
        #  it is high, and work out how much to multiply
        #  each pixel by to get it as big as possible on
        #  the page without distorting.
        #
        bmp = Image.open(file_path)
        if bmp.size[0] > bmp.size[1] :
            bmp = bmp.rotate(90)

        ratios = [1.0 * printable_area[0] / bmp.size[0], 1.0 * printable_area[1] / bmp.size[1]]
        scale = min(ratios)

        # Start the print job, and draw the bitmap to
        #  the printer device at the scaled size.

        hDC.StartDoc(file_path)
        hDC.StartPage()

        dib = ImageWin.Dib(bmp)
        scaled_width, scaled_height = [int(scale * i) for i in bmp.size]
        x1 = int((printer_size[0] - scaled_width) / 2)
        y1 = int((printer_size[1] - scaled_height) / 2)
        x2 = x1 + scaled_width
        y2 = y1 + scaled_height
        dib.draw(hDC.GetHandleOutput(), (x1, y1, x2, y2))

        hDC.EndPage()
        hDC.EndDoc()
        hDC.DeleteDC()

    def printFileList(self, path, file_type) :
        '''
        打印列表中的所有文件
        :param path:
        :param file_type:
        :return:
        '''
        file_list = getAllFormatFile(path, file_type)
        logger.info("开始打印，共【%s】个文件" % len(file_list))
        for file in file_list :
            logger.info("正在打印： %s" % file)
            self.printFile(file)
        logger.info("打印结束")

def getPrinter() :
    default_printer = win32print.GetDefaultPrinter()
    printer_lst = list(win32print.EnumPrinters(2))
    all_printers = [one_tuple[1].split(',')[0] for one_tuple in printer_lst]
    info = "本机连接的打印机：%s\n" \
           "当前默认打印机：%s" % (all_printers, default_printer)
    logger.info(info)
    return (all_printers, default_printer)

if __name__ == '__main__' :
    file_path = r'D:\RPA\a'
    file_list = getAllFormatFile(file_path, 'g')
    pass
