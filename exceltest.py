import json
import os

import xlwt


def getDesktop():
    return os.path.join(os.path.expanduser("~"), 'Desktop')


week = 1

book = xlwt.Workbook(encoding="ascii")
sheet = book.add_sheet("全部课程表")
sheet.write_merge(0, 0, 0, 7, "本学期全部课程表")

style = xlwt.easyxf('font:height 360;')  # 18pt,类型小初的字号
row = sheet.row(0)
row.set_style(style)

hindiNums = dict(
    {1: "一", 2: "二", 3: "三", 4: "四", 5: "五",
     6: "六", 7: "七", 8: "八", 9: "九", 10: "十",
     11: "十一", 12: "十二", })

sheet.write(1, 1, "周日")
for i in range(1, 7):
    sheet.write(1, i + 1, "周{}".format(hindiNums.get(i)))

for i in range(1, 13):
    sheet.write(i + 1, 0, "第{}节".format(hindiNums.get(i)))

listSch = json.load(open(file="schedule/ScheduleList.json", encoding="utf-8", mode="r"))
listCourse = json.load(open(file="schedule/CourseList.json", encoding="utf-8", mode="r"))

'''
sheetMatrix = [["" * 8] * 14]
for item in listSch:
    if item["startWeek"] <= week < item["endWeek"] or week == 0:
        title = item["name"] + " " + item["room"]
        sheetMatrix[item["beginTime"] + 2][item["day"] + 1] = sheetMatrix[item["beginTime"] + 2][
            item["day"]] + title

for i in range(0, len(sheetMatrix)):
    for j in range(0, len(sheetMatrix[i])):
        if sheetMatrix[i][j] != "":
            sheet.write(i, j, sheetMatrix[i][j])
'''
for item in listSch:
    if item["startWeek"] <= week < item["endWeek"] or week == 0:
        sheet.write(item["beginTime"] + 1, item["day"] + 1, item["name"] + " " + item["room"])
book.save(getDesktop() + "\\test.xls")
