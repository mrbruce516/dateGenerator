import datetime
import os
from datetime import timedelta
import random
import xlwt

# 设定开始日期并格式化
print("输入1自行输入日期，输入其他字符自动选择上周一")
now = datetime.datetime.now()
choise = input()
if choise == str(1):
    dateStart = input("请输入周一的日期:")
    dt = datetime.datetime.strptime(dateStart,"%Y-%m-%d")
    dateEnd = (dt + datetime.timedelta(days=4)).strftime("%Y-%m-%d")

    dateStart=datetime.datetime.strptime(dateStart,'%Y-%m-%d')
    dateEnd=datetime.datetime.strptime(dateEnd,'%Y-%m-%d')
    date_list = []
    date_list.append(dateStart.strftime('%Y/%m/%d'))
    while dateStart<dateEnd:
        dateStart += datetime.timedelta(days=+1)
        date_list.append(dateStart.strftime('%Y/%m/%d'))
else:
    dateStart = now - timedelta(days = now.weekday() + 7)
    dateStart = dateStart.strftime("%Y-%m-%d")
    dt = datetime.datetime.strptime(dateStart,"%Y-%m-%d")
    dateEnd = (dt + datetime.timedelta(days=4)).strftime("%Y-%m-%d")

    dateStart=datetime.datetime.strptime(dateStart,'%Y-%m-%d')
    dateEnd=datetime.datetime.strptime(dateEnd,'%Y-%m-%d')
    date_list = []
    date_list.append(dateStart.strftime('%Y/%m/%d'))
    while dateStart<dateEnd:
        dateStart += datetime.timedelta(days=+1)
        date_list.append(dateStart.strftime('%Y/%m/%d'))

# 创建时间
a = 0
time_list = []
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

# 拼接日期与时间
dateGenerator = []
dateGenerator.append(date_list[0] + " " + time_list[0])
dateGenerator.append(date_list[0] + " " + time_list[1])
dateGenerator.append(date_list[1] + " " + time_list[2])
dateGenerator.append(date_list[1] + " " + time_list[3])
dateGenerator.append(date_list[2] + " " + time_list[4])
dateGenerator.append(date_list[2] + " " + time_list[5])
dateGenerator.append(date_list[3] + " " + time_list[6])
dateGenerator.append(date_list[3] + " " + time_list[7])
dateGenerator.append(date_list[4] + " " + time_list[8])
dateGenerator.append(date_list[4] + " " + time_list[9])

# 创建xls文件
wb = xlwt.Workbook()

sh1 = wb.add_sheet('日期')

sh1.write(0,0,'日期')
sh1.write(1,0,dateGenerator[0])
sh1.write(2,0,dateGenerator[1])
sh1.write(3,0,dateGenerator[2])
sh1.write(4,0,dateGenerator[3])
sh1.write(5,0,dateGenerator[4])
sh1.write(6,0,dateGenerator[5])
sh1.write(7,0,dateGenerator[6])
sh1.write(8,0,dateGenerator[7])
sh1.write(9,0,dateGenerator[8])
sh1.write(10,0,dateGenerator[9])

wb.save('日期.xls')

# 校验文件是否存在
check = os.path.isfile("日期.xls")
if check == True:
    print("已成功完成任务")
else:
    print("遇到迷之问题。。")
