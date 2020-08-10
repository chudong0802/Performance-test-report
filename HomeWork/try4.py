import matplotlib.pyplot as plt
import pandas as pd
import matplotlib.dates as mdate
from pylab import *
import numpy as np
import matplotlib.ticker as ticker
from docx import Document
from docx.shared import Inches
import os

# 字体和日期格式的处理
plt.rcParams['font.sans-serif'] = ['SimHei']
dateparse = lambda dates: pd.datetime.strptime(dates, '%H:%M:%S')
# 获取excel文件下的所有sheetname
name = pd.ExcelFile('analysis.xlsx')
sheetname = name.sheet_names
doc = Document()


def fun():
    for i in range(1, len(sheetname)):
        data = pd.read_excel('analysis.xlsx', sheet_name=sheetname[i], arse_dates=True, date_parser=dateparse)
        x = data['TIME']
        y1 = data['DALVIK']
        y2 = data['NATIVE']
        y3 = data['CPU']
        # 画三行一列图：横坐标均为‘TIME’，1-3列纵坐标分别为‘DALVIK’，‘NATIVE’，‘CPU’
        fig = plt.figure(figsize=(15, 15))
        ax1 = fig.add_subplot(311)
        ax1.plot(x, y1, label='DALVIK', linewidth=3.0, ms=10)
        ax1.set_ylabel('DALVIK', size=20, color='b')
        # 隐藏横坐标的标记
        plt.setp(ax1.get_xticklabels(), fontsize=20, visible=False)
        # 画图标以及位置和大小
        plt.legend(loc=4, fontsize=20)

        ax2 = fig.add_subplot(312, sharex=ax1)
        ax2.plot(x, y2, label='NATIVE', linewidth=3.0, ms=10, color='green')
        ax2.set_ylabel('NATIVE', size=20, color='b')
        plt.setp(ax2.get_xticklabels(), visible=False)
        plt.legend(loc=4, fontsize=20)

        ax3 = fig.add_subplot(313, sharex=ax1)
        ax3.plot(x, y3, label='CPU', linewidth=3.0, ms=10, color='y')
        ax3.set_xlabel('TIME', size=20, color='b')
        ax3.set_ylabel('CPU', size=20, color='b')

        plt.legend(loc=0, fontsize=20, framealpha=0.2)
        # 设置横坐标的显示个数，每60个，标记一个坐标。
        plt.gca().xaxis.set_major_locator(ticker.MultipleLocator(60))
        plt.yticks(size=20)
        # 横坐标的标记旋转60度
        plt.xticks(size=20, rotation=60)
        plt.setp(ax3.get_xticklabels(), fontsize=20)
        # 图片保存到指定文件夹下
        plt.savefig('D:/untitled/HomeWork/Pic/TIME_DALVIK_NATIVE_CPU{}.png'.format(i))
        # dic=dict(zip(data['TIME'],data['DALVIK']))

        # 找出pid发生改变的时间点
        list = []
        for j in range(len(data['TIME']) - 1):
            if data['PID'][j] != 0:
                if data['PID'][j + 1] != data['PID'][j]:
                    list.append(data['TIME'][j + 1])
        # 依据list的结果，设置要写入的文本
        if list == []:
            string = '{}没有出现crash'.format(sheetname[i])
            images = 'D:/untitled/HomeWork/Pic/TIME_DALVIK_NATIVE_CPU{}.png'.format(i)  # 保存在本地的图片
            # 使用python-docx模块，将文本以及对应的图片桉顺序写入指定word文本
            doc.add_paragraph(string)  # 添加文字
            doc.add_picture(images, height=Inches(5), width=Inches(5))  # 添加图, 设置宽度
            doc.add_page_break()
            doc.save('word.docx')  # 保存路径
        else:
            # print(sheetname[i])
            string = '{}共crush{}次:分别在\n {}几个时间点'.format(sheetname[i], len(list), list)
            images = 'D:/untitled/HomeWork/Pic/TIME_DALVIK_NATIVE_CPU{}.png'.format(i)  # 保存在本地的图片
            doc.add_paragraph(string)  # 添加文字
            doc.add_picture(images, height=Inches(5), width=Inches(5))  # 添加图, 设置宽度
            doc.add_page_break()
            doc.save('word.docx')


if __name__ == '__main__':
    fun()