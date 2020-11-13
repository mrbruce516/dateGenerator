import datetime
import os
from datetime import timedelta
import random
import xlrd
from xlutils.copy import copy

# 设定开始日期并格式化
print("输入1自行输入日期，输入其他字符自动选择上周一")
# 定义全局数组
dateList = []
timeList = []
dateGenerator = []
# 定义生成日期时间函数
def useDate():
    choise = input()
    if choise == str(1):
        dateStart = input("请输入周一的日期:")
        dt = datetime.datetime.strptime(dateStart,"%Y-%m-%d")
        dateEnd = (dt + datetime.timedelta(days=4)).strftime("%Y-%m-%d")
    
        dateStart=datetime.datetime.strptime(dateStart,'%Y-%m-%d')
        dateEnd=datetime.datetime.strptime(dateEnd,'%Y-%m-%d')
        date_list = dateList
        date_list.append(dateStart.strftime('%Y/%m/%d'))
        while dateStart<dateEnd:
            dateStart += datetime.timedelta(days=+1)
            date_list.append(dateStart.strftime('%Y/%m/%d'))
    else:
        now = datetime.datetime.now()
        dateStart = now - timedelta(days = now.weekday() + 7)
        dateStart = dateStart.strftime("%Y-%m-%d")
        dt = datetime.datetime.strptime(dateStart,"%Y-%m-%d")
        dateEnd = (dt + datetime.timedelta(days=4)).strftime("%Y-%m-%d")
    
        dateStart=datetime.datetime.strptime(dateStart,'%Y-%m-%d')
        dateEnd=datetime.datetime.strptime(dateEnd,'%Y-%m-%d')
        date_list = dateList
        date_list.append(dateStart.strftime('%Y/%m/%d'))
        while dateStart<dateEnd:
            dateStart += datetime.timedelta(days=+1)
            date_list.append(dateStart.strftime('%Y/%m/%d'))

    # 创建时间
    a = 0
    time_list = timeList
    while a<10:
        moring = 8
        eveing = random.randint(17,18)
        mminute = random.randint(45,59)
        eminute = random.randint(30,59)
        second = random.randint(10,59)
        mtime = str(0) + str(moring) + ":" + str(mminute) + ":" + str(second)
        etime = str(eveing) + ":" + str(eminute) + ":" + str(second)
        a = a+1
        time_list.append(mtime)
        time_list.append(etime)
    
# 拼接日期与时间函数
def splice():
    date_generator = dateGenerator
    # 工作日
    dateFreq = 0
    # 早上和晚上的时间
    timeFreq = 0
    # 一周工作日5天
    while dateFreq < 5:
        date_generator.append(dateList[dateFreq] + " " + timeList[timeFreq])
        timeFreq += 1
        date_generator.append(dateList[dateFreq] + " " + timeList[timeFreq])
        dateFreq += 1
        timeFreq += 1

# 创建函数完毕并调用
useDate()
# 打开excel文件
rb = xlrd.open_workbook("考勤大法好.xls")
wb = copy(rb)

sh1 = wb.get_sheet(0)

people = int(input("请输入需要生成的打卡人数"))
# 生成时间库
timeLib = 0
while timeLib < people:
    splice()
    timeLib += 1

genFreq = people * 10
xlsLine = 1
while xlsLine < genFreq+1:
    sh1.write(xlsLine,4,dateGenerator[xlsLine-1])
    xlsLine += 1

wb.save('考勤大法好.xls')

# 校验文件是否存在
check = os.path.isfile("考勤大法好.xls")
if check == True:
    print("已成功完成任务")
else:
    print("遇到迷之问题。。")
