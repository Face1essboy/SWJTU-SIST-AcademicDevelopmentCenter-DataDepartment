# -*- coding: utf-8 -*-
import csv

root_path = r"D:/synchronize folder/OneDrive - my.swjtu.edu.cn/文件/学业发展中心/源数据/"  # 源数据的根目录
output_path = r"D:/synchronize folder/OneDrive - my.swjtu.edu.cn/文件/教务信息/统计分析/"  # 输出数据路径
output_file_name = r"通信+计算机+物联网+轨道 区间长度1"  # 输出文件名
folder_and_file = [
    r"18 通信2014~2018/", [
        r"通信2014-01班", r"通信2014-02班", r"通信2014-03班", r"通信2014-01班", r"通信2014-02班", r"通信2014-03班",
        r"通信2015-01班", r"通信2015-02班", r"通信2015-03班",
        r"通信2016-01班", r"通信2016-02班",
        r"通信2017-01班", r"通信2017-02班", r"通信2017-03班",
        r"通信2018-01班", r"通信2018-02班", r"通信2018-03班"],

    # r"04 其他/", [
    #     r"茅院(信息)2014-01班", r"通信2015-05班[国]",
    #     r"通信2016-03班[国]", r"智能(茅班)2019-01班"],

    r"15 计算机2014~2018/", [
        r"计算机2014-01班", r"计算机2014-02班", r"计算机2014-03班", r"计算机2014-04班",
        r"计算机2015-01班", r"计算机2015-02班", r"计算机2015-03班",
        r"计算机2016-01班", r"计算机2016-02班",
        r"计算机2017-01班", r"计算机2017-02班", r"计算机2017-03班",
        r"计算机2018-01班", r"计算机2018-02班", r"计算机2018-03班"],

    r"07 物联网2014~2017/", [
        r"物联网2014-01班", r"物联网2014-02班",
        r"物联网2015-01班", r"物联网2015-02班",
        r"物联网2016-01班", r"物联网2016-02班",
        r"物联网2017-01班"],

    r"18 轨道2014~2019/", [
        r"轨道2014-01班", r"轨道2014-02班", r"轨道2014-03班",
        r"轨道2015-01班", r"轨道2015-02班", r"轨道2015-03班",
        r"轨道2016-01班", r"轨道2016-02班", r"轨道2016-03班", r"轨道2016-04班",
        r"轨道2017-01班", r"轨道2017-02班", r"轨道2017-03班",
        r"轨道2018-01班", r"轨道2018-02班", r"轨道2018-03班",
        r"轨道2019-01班", r"轨道2019-02班", r"轨道2019-03班"]
]
file_format = r".csv"  # 源数据的文件格式


rang_leng = 1
offset = 4
offset_dynamic = 100 // rang_leng

output_header = [
    r"课程代码",
    r"课程名称",
    r"参加人数",
    r"第一次正考所有参加者成绩之和"]

str1 = ["总成绩", "期末成绩", "平时成绩"]
for i in range(3):
    for j in range(offset_dynamic):
        output_header.append(
            str1[i] + '[' + str(j * rang_leng) + ',' + str((j+1) * rang_leng) + (')' if j != offset_dynamic - 1 else ']'))

output_header += [
    r"期末成绩占比100%人数",
    r"平时成绩占比100%人数",
    r"第1次正考未通过人数(正考总成绩 < 60)",
    r"第1次考试补考人数",
    r"第1次考试补考通过人数",
    r"重修人次"]


# 获得一个成绩所处的区间
def getRank(grade):
    rank = int(grade // rang_leng)
    return rank if rank != offset_dynamic else rank - 1


def getStudentId(item):
    return item[2]


def getCourseId(item):
    return item[4]


def getCourseName(item):
    return item[5]


def getTotalGrade(item):
    return float(item[8])


def getFinalGrade(item):
    try:
        grade = float(item[9])
    except ValueError as e:
        print('ValueError:', e)
        print(item)
        grade = getTotalGrade(item)
    finally:
        return grade


def getUsualGrade(item):
    try:
        grade = float(item[10])
    except ValueError as e:
        print('ValueError:', e)
        print(item)
        grade = getTotalGrade(item)
    finally:
        return grade


def isFinalGrade100(item):
    try:
        if int(item[11]) == 100:
            return 1
        else:
            return 0
    except ValueError as e:
        print('ValueError:', e)
        print(item)
        return 0


def isUsualGrade100(item):
    try:
        if int(item[12]) == 100:
            return 1
        else:
            return 0
    except ValueError as e:
        print('ValueError:', e)
        print(item)
        return 0


def isMakeup(item):
    if item[14] == "补考":
        return 1
    else:
        return 0


def getTerm(item):
    return int(item[20]) * 10 + int(item[21])


# 依照学号进行一次分割,分割出从begin开始的同学号的数据, 返回下一个学号的位置
def partitionByStudentID(source_list, begin):
    id = getStudentId(source_list[begin])
    end = begin
    while source_list[end] is not None:
        if getStudentId(source_list[end]) != id:
            break
        else:
            end += 1
    return end


# 根据分割好的数据进行统计,不包含end
def doStatisticAboutStudent(source_list, ans_dict, begin, end):
    # 依照课程代码进行一次分割,分割出从course_begin开始到的同课程id的数据,返回下一个课程的位置
    def partitionByCourseId(course_begin):
        course_end = course_begin
        Course_id_begin = getCourseId(source_list[course_begin])
        while source_list[course_end] is not None:
            if getCourseId(source_list[course_end]) == Course_id_begin:
                course_end += 1
            else:
                break
        return course_end

    # 统计一门课程的数据
    def doStatisticAboutCourse(course_begin, course_end):
        # 获得第一次正考的位置,第一次补考的位置(若无则为None),重修次数
        def getDetail():
            detail = {"earliest": getTerm(
                source_list[course_begin]), "formal": None, "makeup": None}
            count_makeup = 0
            for i in range(course_begin, course_end):
                if isMakeup(source_list[i]):
                    count_makeup += 1
                if detail["earliest"] >= getTerm(source_list[i]):
                    detail["earliest"] = getTerm(source_list[i])
                    if isMakeup(source_list[i]):
                        detail["makeup"] = i
                    else:
                        detail["formal"] = i
                        if detail["makeup"] is not None:
                            if getTerm(source_list[detail["makeup"]]) > detail["earliest"]:
                                detail["makeup"] = None
            if detail["formal"] is None and detail["makeup"] is not None:
                detail["formal"] = detail["makeup"]
                detail["makeup"] = None
                count_makeup = 0
            return detail["formal"], detail["makeup"], course_end - course_begin - count_makeup - 1

        # 输出格式的成绩区间相对于首列的偏移量

        total_seat, formal_seat, makeup_seat = offset, offset + \
            offset_dynamic, offset + offset_dynamic * 2
        # 课程代码
        course_id = getCourseId(source_list[course_begin])
        # 若为新添加的课程,初始化dict中的一行
        if course_id not in ans_dict:
            ans_dict[course_id] = [course_id,
                                   getCourseName(source_list[course_begin])] + [0] * (len(output_header) - 2)
        # 获得该课程第一次正考与其补考位置,该人重修该课程次数
        formal_i, makeup_i, count_makeup = getDetail()
        # 参加人数+1
        ans_dict[course_id][2] += 1
        # 成绩求和
        ans_dict[course_id][3] \
            += getTotalGrade(source_list[formal_i])
        # 总成绩相应区间人数+1
        ans_dict[course_id][total_seat +
                            getRank(getTotalGrade(source_list[formal_i]))] += 1
        # 期末成绩相应区间人数+1
        ans_dict[course_id][formal_seat +
                            getRank(getFinalGrade(source_list[formal_i]))] += 1
        # 平时成绩相应区间人数+1
        ans_dict[course_id][makeup_seat +
                            getRank(getUsualGrade(source_list[formal_i]))] += 1
        # 期末成绩占比100%人数
        ans_dict[course_id][offset + offset_dynamic * 3] \
            += isFinalGrade100(source_list[formal_i])
        # 平时成绩占比100%人数
        ans_dict[course_id][offset + offset_dynamic * 3 + 1] \
            += isUsualGrade100(source_list[formal_i])
        # 第一次正考未通过人数
        ans_dict[course_id][offset + offset_dynamic * 3 + 2] \
            += 1 if getTotalGrade(source_list[formal_i]) < 60 else 0
        if makeup_i is not None:
            # 第一次考试补考人数
            ans_dict[course_id][offset + offset_dynamic * 3 + 3] += 1
            # 第一次考试补考通过人数
            ans_dict[course_id][offset + offset_dynamic * 3 + 4] \
                += 1 if getTotalGrade(source_list[makeup_i]) >= 60 else 0
        # 重修人次
        ans_dict[course_id][offset + offset_dynamic * 3 + 5] += count_makeup

    course_begin, course_end = begin, end
    while True:
        course_end = partitionByCourseId(course_begin)
        doStatisticAboutCourse(course_begin, course_end)
        if course_end >= end:
            return
        else:
            course_begin = course_end


# 对统计结果字典进行排序与处理,返回一个可以直接用于输出的list
def formatResults(ans_dict, ans_header, save_path, file_name):
    content = sorted(ans_dict.values(),
                     key=lambda value: value[2], reverse=True)
    with open(save_path + file_name + '.csv', "w+", newline='') as output_file:
        writer = csv.writer(output_file)
        writer.writerow(ans_header)
        writer.writerows(content)
    return content


ans_dict = dict()
for i in range(len(folder_and_file) // 2):
    folder_name = folder_and_file[i*2]
    for file_name in folder_and_file[i*2+1]:
        with open(
                root_path + folder_name + file_name + file_format,
                newline="",
                encoding="utf-8"
        ) as csv_file:
            source_list = list(csv.reader(csv_file))[3:]
            source_list.append(None)  # 用于判断是否到达结尾
            begin, end = 0, 0

            while source_list[begin] is not None:
                end = partitionByStudentID(source_list, begin)
                doStatisticAboutStudent(source_list, ans_dict, begin, end)
                begin = end
output_list = formatResults(
    ans_dict,
    output_header,
    output_path,
    output_file_name
)

# xls = csv.reader(
#     r"D:/synchronize folder/OneDrive - my.swjtu.edu.cn/文件/学业发展中心/源数据/04 其他/智能(茅班)2019-01班.csv", header=None)
# print(xls)
