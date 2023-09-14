import pandas as pd
from os import rename
import datetime


print("""
————————————————————课表导入系统————————————————————
    您需要进入my.cqu.edu.cn，
    登陆您的账号，进入选课界面，点击上方“查看课表”按钮，
    然后点击右上角的Excel下载按钮
    将下载好的"课表.xlsx"拖入此页面
      """)
path = input("请拖入课表: ")
path = path[:-8]
start = input("请输入学期起始日期（2023学年上学期为“20230828”）: ")
timetable = [["083000", "091500"], 
             ["092500", "101000"], 
             ["103000", "111500"], 
             ["112500", "121000"],
             ["133000", "141500"], 
             ["142500", "151000"], 
             ["152000", "160500"], 
             ["162500", "171000"], 
             ["172000", "180500"],
             ["190000", "194500"],
             ["195500", "204000"], 
             ["205000", "213500"],
             ["214500", "223000"]]
timezone = "+8"


numstr = ["0","1","2","3","4","5","6","7","8","9"]
day0str = ["一","二","三","四","五","六"]
start = datetime.datetime.strptime(start, "%Y%m%d")
for i in range(13):
    timetable[i][0] = datetime.datetime.strptime(timetable[i][0], "%H%M%S")
    timetable[i][1] = datetime.datetime.strptime(timetable[i][1], "%H%M%S")
fm=pd.read_excel(path+"课表.xlsx",sheet_name="Sheet0", skiprows=1)
event = """BEGIN:VCALENDAR
"""
uid = 0
print("""
-.-.-.-.-.-.-.-.-.-.正在创建课表.-.-.-.-.-.-.-.-.-.-.-
      
      """)
for i in range(0,fm["课程名称"].count()):
    time = (fm["上课时间"])[i]
    week0, else0 = time.split("周")
    week = []
    if "," in week0:
        week1 = week0.split(",")
        for week2 in week1:
            if "-" in week2:
                week2 = week2.split("-")
                week2 = list(range(int(week2[0]), int(week2[-1])+1))
            else:
                week2 = [int(week2)]
            week += week2
    elif "-" in week0:
        week0 = week0.split("-")
        week0 = list(range(int(week0[0]), int(week0[-1])+1))
        week += week0
    else:
        week += [int(week0)]
    if len(else0) != 0:
        day0 = day0str.index(else0[2:3]) + 1
        classnum = else0[3:-1]
        if "-" in classnum:
            classnum = classnum.split("-")
            classnum = [int(classnum[0]), int(classnum[-1])]
        else:
            classnum = [int(classnum)]
        for j in week:
            event += """
BEGIN:VEVENT
UID:{0}
DTSTART:{1}T{2}Z
DTEND:{3}T{4}Z
SUMMARY:{5}
DESCRIPTION:{6}
LOCATION:{7}
END:VEVENT
""".format(uid, 
           (start+datetime.timedelta(days=j*7+day0-8)).strftime("%Y%m%d"), 
           (timetable[classnum[0]-1][0]+datetime.timedelta(hours=-int(timezone))).strftime("%H%M%S"), 
           (start+datetime.timedelta(days=j*7+day0-8)).strftime("%Y%m%d"), 
           (timetable[classnum[1]-1][1]+datetime.timedelta(hours=-int(timezone))).strftime("%H%M%S"), 
           (fm["课程名称"])[i], 
           (fm["上课教师"])[i], 
           (fm["上课地点"])[i])
            uid += 1
            print("已导入{}节课程".format(uid))
    else:
        for j in week:
            for k in range(7):
                event += """
BEGIN:VEVENT
UID:{0}
DTSTART:{1}T{2}Z
DTEND:{3}T{4}Z
SUMMARY:{5}
DESCRIPTION:{6}
END:VEVENT
""".format(uid, (start+datetime.timedelta(days=j*7+k-7)).strftime("%Y%m%d"), 
           (timetable[0][0]+datetime.timedelta(hours=-int(timezone))).strftime("%H%M%S"), 
           (start+datetime.timedelta(days=j*7+k-7)).strftime("%Y%m%d"), 
           (timetable[-1][1]+datetime.timedelta(hours=-int(timezone))).strftime("%H%M%S"), 
           (fm["课程名称"])[i], 
           (fm["上课教师"])[i])
                uid += 1
                print("已导入{}节课程".format(uid))
event += """
END:VCALENDAR"""
with open(path+"ClassTable.txt", "w") as file:
    file.write(event)
rename(path+"ClassTable.txt", path+"ClassTable.ics")
print("""
    课表已成功创建
    请将原文件夹下的“ClassTable.ics”文件发送到手机端
    使用默认日历按照提示操作即可导入成功
""")