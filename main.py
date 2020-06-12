# -*- coding:utf-8 -*-
#developer:倪小白
#blog:www.nixiaobai.com
import copy
import json
import os
import time
import xlwings as xw
import re
from pycurriculum import Course, Curriculum

#新建一个班级课表数据的模板
kebiao = {
    "classname": "",
    "dateinfo": "",
    "Monday":[],
    "Tuesday":[],
    "Wednesday":[],
    "Thursday":[],
    "Friday":[],
    "Saturday":[],
    "Sunday":[]
}

#整周枚举
week_name = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]

#单节课数据模板
course = {
    "course_name": "", #课程名称
    "class_room":"",   #教学地点
    "section":"",      #节次（1~6）
    "teacher":"",      #任课教师
    "week":"",         #总周次缩写
    "week_array":[]    #周次数组（字符）
}

#建立 当周-节次 关系表
#指代的是excel单元格
relation=[
    ["B","C","D","E","F","G"],
    ["H","I","J","K","L","M"],
    ["N","O","P","Q","R","S"],
    ["T","U","V","W","X","Y"],
    ["Z","AA","AB","AC","AD","AE"],
    ["AF","AG","AH","AI","AJ","AK"],
    ["AL","AM","AN","AO","AP","AQ"]
]

#获取俩字符之间字符串(事实上是从小括号到末尾)
def GetMidstrring(txt, start, end):
    index_1 = txt.find(start) + 1
    return txt[index_1:-1]

#分析是否是教学地点
def EndAnalyze(strr):
    if ("教室" not in strr and "楼" not in strr and "博明楼" not in strr and "工艺室" not in strr and "食B" not in strr and "食A" not in strr and "实验室" not in strr and "场地" not in strr and "逸" not in strr and "博文" not in strr and "西泳" not in strr and "土木2" not in strr and "检测室" not in strr and "工艺室" not in strr and "国际1" not in strr and "国际2" not in strr and "国际3" not in strr and "国际4" not in strr and "国际5" not in strr and "国际6" not in strr and "检测站" not in strr):
        return True
    return False

#将教师信息与周次信息分离
def TeacherAnalyze(strr):
    if("(" in strr and ")" in strr):
        weeks = GetMidstrring(strr,"(",")")
        return (strr.replace("({})".format(weeks),""),weeks)
    return (strr,"0")

#周次分析
def WeekAnalyze(weekstrr):
    strr = weekstrr
    #是否为空
    if strr == None or strr.strip() == "" : return ["0"]
    #是否包含数字
    if not bool(re.search(r'\d', strr)) : return ["0"]

    #默认单双周为True
    single = True
    double = True
    if("单" in strr) : double = False
    elif("双" in strr): single = False

    #过滤非法字符
    strr = strr.strip().replace(",周", "").replace("、", ",").replace("=", ",").replace("周", "").replace("单", "").replace("双", "").replace("[", "").replace("]", "").replace("(", "").replace(")", "")
    #按,分割成数组
    week_array = strr.split(",")
    #周次数字列表
    week_list = []
    for item in week_array :
        if("-" in item):
            w_array = item.split("-")
            prev = int(float(w_array[0].strip()))
            next = int(float(w_array[1].strip()))

            if prev==next:
                week_list.append(str(prev))
            else:
                for i in range(prev,next + 1):
                     if(single and double):week_list.append(str(i))
                     elif(single and i%2 == 1):week_list.append(str(i))
                     elif(double and i%2 == 0):week_list.append(str(i))
        else:
            if (single and int(float(item.strip()))%2 == 1):week_list.append(str(item))
            elif (double and int(float(item.strip()))%2 == 0):week_list.append(str(item))

    return week_list

#ics周次分析
def IcsWeek(weekstrr):
    strr = weekstrr
    #是否为空
    if strr == None or strr.strip() == "" : return ["0"]
    #是否包含数字
    if not bool(re.search(r'\d', strr)) : return ["0"]

    #默认单双周为True
    single = True
    double = True
    if("单" in strr) : double = False
    elif("双" in strr): single = False

    #过滤非法字符
    strr = strr.strip().replace(",周", "").replace("、", ",").replace("=", ",").replace("周", "").replace("单", "").replace("双", "").replace("[", "").replace("]", "").replace("(", "").replace(")", "")
    #按,分割成数组
    week_array = strr.split(",")
    #周次数字列表
    week_list = []
    for item in week_array :
        if("-" in item):

            if(single and double):week_list.append(item)
            elif(single):week_list.append("{}-1".format(item))
            elif(double):week_list.append("{}-2".format(item))
        else:
            week_list.append("{}-{}".format(item,item))

    return week_list

#节数转换
def sectionToNum(var):
	return {
			'1': "1-2",
			'2': "3-4",
            '3': "5-6",
            '4': "7-8",
            '5': "9-10",
            '6': "11-12"
	}.get(var,'1-2')

#记录解析时间
##kebiao["dateinfo"] = str(time.strftime('%Y-%m-%d',time.localtime(time.time())))
kebiao["dateinfo"] = "2020-09-01"
def GenerateJson():
    xls_name = input("请手动输入xls文件名称：")
    xls_min_row = int(float(input("请手动输入xls班级名称起始行：").strip()))
    xls_max_row = int(float(input("请手动输入xls最大行：").strip()))
    if os.path.exists(xls_name):
        try:
            #引用xls数据
            wb = xw.App(visible=False, add_book=False).books.open(xls_name)
            sht = wb.sheets[0]
            print("读取 " + xls_name + " 文件成功")

            #获取A[min]:A[max]原始数据(默认A4是起始行)
            class_row_raw = sht.range("A{}:A{}".format(xls_min_row,xls_max_row)).value
            #A[min]:A[max]中，每一个班级的起始行号
            class_row_list = []
            # A[min]:A[max]中，每一个班级的名称
            class_name_list = []

            #通过for循环将有效数据提取
            for i in range(0,len(class_row_raw)):
                if class_row_raw[i] != None:
                    class_name_list.append(class_row_raw[i])
                    class_row_list.append(i + xls_min_row)

            #阶段日志
            print("\n总共检录 {} 个班级".format(len(class_name_list)))

            #解析计时
            start_time = time.clock()

            #在脚本根目录新建json文件夹
            if not os.path.exists("json"):
                os.mkdir("json")
                print("\n成功新建 json 文件夹")

            #最大索引值
            class_index = len(class_row_list)
            for i in range(0,class_index):
                #当前课程开始行号
                this = class_row_list[i]
                #当前课程结束行号
                if i != class_index - 1 :
                    next = class_row_list[i + 1] - 1
                else:
                    next = xls_max_row

                #引用区间
                this_range = "A{}:A{}".format(this,next)
                #print(this_range)

                #获取班级名称
                classname = sht.range("A{}".format(this)).value
                #print(classname)

                kb = None
                kb =  copy.deepcopy(kebiao)
                kb["classname"] = classname


                for week in range(1,8):
                    weekname = week_name[week-1]
                    for section in range(1,7):
                        try:
                            #单节次单元格
                            self_range = this_range.replace("A",relation[week - 1][section - 1])
                            this_sect = sht.range(self_range).value
                            count = 0
                            for j in range(0,len(this_sect)):
                                item = this_sect[j]
                                if this_sect[j] == None : break
                                count += 1
                                if(count == 1):
                                    kb[weekname].append(copy.deepcopy(course))
                                    #下面这种写法十分不推荐,lazy,效率差
                                    kb[weekname][len(kb[weekname]) - 1]["course_name"] = item
                                    kb[weekname][len(kb[weekname]) - 1]["section"] = str(section)
                                elif(count == 2):
                                    kb[weekname][len(kb[weekname]) - 1]["teacher"],kb[weekname][len(kb[weekname]) - 1]["week"] = TeacherAnalyze(item)
                                    kb[weekname][len(kb[weekname]) - 1]["week_array"] = WeekAnalyze(kb[weekname][len(kb[weekname]) - 1]["week"])
                                    if(j + 1 < len(this_sect) ):
                                        next = this_sect[j + 1]
                                        if(next != None and EndAnalyze(next)):count = 0
                                elif(count == 3):
                                    kb[weekname][len(kb[weekname]) - 1]["class_room"] = item
                                    count = 0
                                else:
                                    count = 0
                        except Exception as e:
                            print("\n解析错误：\n班级名称：{}\n周{}-第{}节\n{}".format(classname,week,section,e))

                print("解析 {} 结束".format(classname))
                #生成json文件
                f = open("json/{}.json".format(classname),"w",encoding="utf-8")
                f.write(json.dumps(kb,ensure_ascii=False).strip())
                f.close()
                print("生成 {}.json 文件\n".format(classname))

            print("总共解析 {} 个班级".format(len(os.listdir("json"))))
        except:
            print("解析 {} 失败\n".format(xls_name))
    else:
        print("未在脚本根目录检测到 " + xls_name)

#将一些重复的课程去掉，比如体育、英语选修，保留一节即可
def JsonModify():
    json_bool = input("输入y/Y开始对json进行去重操作：")
    if ("y" in json_bool or "Y" in json_bool):
        json_list=os.listdir("json")
        for j in json_list:
            try:
                f = open("json/{}".format(j),"r",encoding="utf-8")
                data = json.load(f)
                classname = data["classname"]
                print("读取 {} 成功".format(classname))

                data2 = data.copy()
                for day in week_name:
                    del_list = []
                    count = 0
                    for item in data[day]:
                        count += 1
                        for it in range(count,len(data[day])):
                            if(item["section"]==data[day][it]["section"] and item["week_array"]==data[day][it]["week_array"] and  item["teacher"]!=data[day][it]["teacher"] ) : del_list.append(it)
                    del_set = list(set(del_list))
                    if(len(del_set) > 0):
                        del_set.sort(reverse=True)
                        for m in del_set:
                            del data2[day][int(m)]

                 # 在脚本根目录新建json_modify文件夹
                if not os.path.exists("json_modify"):
                    os.mkdir("json_modify")
                    print("\n成功新建 json_modify 文件夹")

                # 生成json文件
                f = open("json_modify/{}.json".format(classname), "w", encoding="utf-8")
                f.write(json.dumps(data, ensure_ascii=False).strip())
                f.close()
                print("去重 {}.json 完成\n".format(classname))

            except Exception as e:
                print("去重 {} 错误\n{}".format(j,e) )

#一种课模板
one_ics={
    "name":"",
    "teacher":"",
    "detail":[]
}

def GenerateIcs():
    json_dir = input("请输入你要解析成ics的json文件夹:")
    # 在脚本根目录新建ics文件夹
    start_time = time.clock()
    if not os.path.exists("ics"):
        os.mkdir("ics")
        print("\n成功新建 ics 文件夹")

    json_list = os.listdir(json_dir)
    for j in json_list:
        #第一周第一天时间
        tc = Curriculum('2020-02-24')
        f = open("{}/{}".format(json_dir,j), "r", encoding="utf-8")
        data = json.load(f)
        classname = data["classname"]
        print("读取 {} 成功".format(classname))
        some_course = {}
        count = 0
        for day in week_name:
            count += 1
            for item in data[day]:
                name = item["course_name"]
                if name in some_course.keys():
                    w = IcsWeek(item["week"])
                    for it in w:
                        some_course[name]["detail"].append([item["class_room"],"{}-{}".format(count,sectionToNum(item["section"])),it])
                else:
                    some_course[name] = copy.deepcopy(one_ics)
                    some_course[name]["name"] = name
                    some_course[name]["teacher"] = item["teacher"]
                    w = IcsWeek(item["week"])
                    for it in w:
                        some_course[name]["detail"].append([item["class_room"], "{}-{}".format(count, sectionToNum(item["section"])), it])

        for one in some_course.values():
            tc.add(Course(one["name"],one["teacher"], one["detail"]))

        tc.to_ics(classname)
        print("生成 {}.ics 文件\n".format(classname))

if __name__ == "__main__":
    while True:
        num = input("""
        欢迎使用 林大课表 数据生成器 version 1.0
        请输入数字代码：
        【1】生成课表json
        【2】去重课表json
        【3】生成课表ics
        【4】开发者信息
        """)

        if ("1" in num):
            GenerateJson()
        elif ("2" in num):
            JsonModify()
        elif ("3" in num):
            GenerateIcs()
        elif ("4" in num):
            print("开发者：倪小白\n博客：www.nixiaobai.com")
        else:
            print("输入错误")