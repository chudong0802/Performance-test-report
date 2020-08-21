import os
import pandas as pd
import matplotlib.pyplot as plt
import sys

#将指定文件筛选出来并转换为csv文件写到指定路径
#相对路径
os.system('mkdir chart')
os.system('mkdir csvfile')
path = '.\\analysis_module\\analysis_module\\'
out_path = './csvfile/'
all_file = os.listdir(path)
dealed_path = []
file_name = []

#处理xlsx文件
def deal_xlsx():
    for name in all_file:
        if name.startswith('com.') and os.path.splitext(name)[1] == '.xlsx':
            new_name = os.path.splitext(name)[0]
            data_xlsx = pd.read_excel(path+name)
            data_xlsx.to_csv(out_path+new_name+'.csv',encoding='utf-8')

#处理csv文件
def deal_csv():
    for needed_name in os.listdir(out_path):
        needed_file = os.path.join(out_path,needed_name)
        dealed_path.append(needed_file)
        filename = os.path.splitext(needed_name)[0]
        file_name.append(filename)

#遍历读取文件并画图
def deal_file_draw():
    for i in range(len(dealed_path)):
        dealed_file = pd.read_csv(dealed_path[i])
        dealed_filename = './chart/'+file_name[i]
        x = dealed_file['Folder']
        all_y = [dealed_file['dalvik max'], dealed_file['dalvik min'], dealed_file['dalvik avg']]
        filedir = os.makedirs(name=dealed_filename)
        plt.figure(figsize=(20, 10))
        # 坐标轴重合时，设置倾斜度
        # plt.xticks(rotation=45)
        plt.plot(x, all_y[0], color='orange', marker='s', label='max')
        plt.plot(x, all_y[1], color='green', marker='o', label='min')
        plt.plot(x, all_y[2], color='blue', marker='*', label='avg')
        plt.legend(bbox_to_anchor=(0.9, 1), framealpha=0.2)
        plt.savefig(dealed_filename + '/dalvik.png')

        # bluetooth_native
        plt.figure(figsize=(20, 10))
        all_y1 = [dealed_file['native max'], dealed_file['native min'], dealed_file['native avg']]
        plt.plot(x, all_y1[0], color='orange', marker='s', label='max')
        plt.plot(x, all_y1[1], color='green', marker='o', label='min')
        plt.plot(x, all_y1[2], color='blue', marker='*', label='avg')
        plt.legend(bbox_to_anchor=(0.9, 1), framealpha=0.2)
        plt.savefig(dealed_filename + '/native.png')

        # bluetooth_cpu
        plt.figure(figsize=(20, 10))
        all_y2 = [dealed_file['cpu max'], dealed_file['cpu min'], dealed_file['cpu avg']]
        plt.plot(x, all_y2[0], color='orange', marker='s', label='max')
        plt.plot(x, all_y2[1], color='green', marker='o', label='min')
        plt.plot(x, all_y2[2], color='blue', marker='*', label='avg')
        plt.legend(bbox_to_anchor=(0.9, 1), framealpha=0.2)
        plt.savefig(dealed_filename + '/cpu.png')
#获取当前路径
def fileDir():
    path = './csvfile/'
    print(path)
    #判断为脚本文件还是编译后文件，如果是脚本文件则返回脚本目录，否则返回编译后的文件路径
    if os.path.isdir(path):
        return path
    elif os.path.isfile(path):
        return os.path.dirname(path)

#获取文件后缀名
def suffix(file,*suffixName):
    array = map(file.endswith,suffixName)
    if True in array:
        return True
    else:
        return False

def deleteFile():
    targetDir = fileDir()
    for file in os.listdir(targetDir):
        targetFile = os.path.join(targetDir,file)
        if suffix(file,'csv'):
            os.remove(targetFile)


if __name__ == '__main__':
    deal_xlsx()
    deal_csv()
    deal_file_draw()
    deleteFile()

