import json
import os
import re
import sys
from time import sleep

import selenium.common.exceptions
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

config = json.load(open(r"config.json"))
browsers = json.load(open(r"params\\Browsers.json"))
drivers = json.load(open(r"params\\Drivers.json"))


def getDesktop():
    return os.path.join(os.path.expanduser("~"), 'Desktop')


def getOption(browser):
    if browser == 'edge':
        return webdriver.EdgeOptions()
    elif browser == 'firefox':
        return webdriver.FirefoxOptions()
    elif browser == 'chrome':
        return webdriver.ChromeOptions()


def getDriver(browser, options):
    driverPath = "drivers/" + drivers[config["browser"]]
    service = Service(driverPath)
    options.add_experimental_option('excludeSwitches', ['enable-logging'])

    if browser == 'edge':
        return webdriver.Edge(service=service, options=options)
    elif browser == 'firefox':
        return webdriver.Firefox(service=service, options=options)
    elif browser == 'chrome':
        return webdriver.Chrome(service=service, options=options)


def exportToExcel(week):
    import xlwt
    book = xlwt.Workbook(encoding="ascii")
    sheet = book.add_sheet("全部课程表" if week == "0" else "第{}周课程表".format(week))
    sheet.write_merge(0, 0, 0, 7, "本学期全部课程表" if week == "0" else "第{}周课程表".format(week))

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

    book.save(getDesktop()+"\\test.xls")


def exportToImage(week):
    pass


def afterDownloading():
    print("请选择接下来的操作")
    print("1. 导出Excel")
    print("2. 导出图片")
    print("3. 退出")
    sel = input("请选择接下来的操作")

    if sel == "3":
        exit(0)
    else:
        sel1 = input("请输入要导出第几周的课表，输入0导出全部课表")
        if sel == "1":
            exportToExcel(sel1)
        else:
            exportToImage(sel1)
        exit(0)


def main():
    sel = input("是否要更新课表？0-否，1-是")
    if sel == "0":
        afterDownloading()
    browserName = config["browser"]
    browserFile = browsers[config["browser"]]
    binPath = config["binPath"] + "\\" + browserFile

    options = getOption(browserName)
    options.binary_location = binPath
    options.add_argument("--headless")

    driver = getDriver(browserName, options)

    driver.get(config["defaultPage"])

    if config["stuID"] == "":
        stuID = input("请输入你的学号: \n")
    else:
        stuID = config["stuID"]

    if config["pwd"] == "":
        pwd = input("请输入密码: \n")
    else:
        pwd = config["pwd"]

    idBox = driver.find_element(By.ID, "un")
    idBox.clear()
    idBox.send_keys(stuID)

    pwdBox = driver.find_element(By.ID, "pd")
    pwdBox.clear()
    pwdBox.send_keys(pwd)

    driver.find_element(By.ID, "index_login_btn").click()

    try:
        errormsg = driver.find_element(By.ID, "errormsg")
        print(errormsg.text)
        pass
    except selenium.common.exceptions.NoSuchElementException:
        print("登陆成功")

    print("正在获取课程表")
    sleep(2)
    driver.find_elements(By.CLASS_NAME, "p_1")[16].click()

    sleep(2)
    print("课程表获取成功")
    tbSch = driver.find_element(By.ID, "manualArrangeCourseTable")
    tbCourse = driver.find_element(By.ID, "tasklesson")

    # spSch = BeautifulSoup(tbSch.get_attribute('innerHTML'), features=config["feature"])
    spSch = BeautifulSoup(open("test.html",encoding="UTF-8",mode="r"), features=config["feature"])
    spCourse = BeautifulSoup(tbCourse.get_attribute('innerHTML'), features=config["feature"])

    driver.close()

    print("开始下载课程表数据")
    # Now save the course list
    trCourse = spCourse.select("table.gridtable > tbody > tr")

    courseList = []
    for tr in trCourse:
        tds = tr.select('td')

        courseList.append(dict({
            "ID": int(tds[0].get_text()),
            "name": tds[2].get_text().strip(),
            "point": float(tds[3].get_text()),
            "code": tds[4].select("a")[0].get_text(),
            "teacher": tds[5].get_text(),
            "memo": tds[7].get_text()
        }))

    json.dump(courseList, open("schedule/courseList.json", "w", encoding="UTF-8"), ensure_ascii=False)
    print("课程列表加载完成")

    # Now save the schedule
    trSch = spSch.select("tbody > tr")

    ptnCode = re.compile(r'(A\d{6})')
    ptnCampus = re.compile(r'\([\u4e00-\u9fa5]{0,4}\)')
    ptnSigDbl = re.compile(r'[\u4e00-\u9fa5]{1}')

    listSch = []

    print("正在下载课程表")
    prog = 0
    target = len(trSch) * 7 - 1

    print("\r", end="")
    print("进度: {}/{}".format(prog, target), end="")
    sys.stdout.flush()

    row = 1
    for tr in trSch:
        tds = tr.select("td")

        col = 0
        for td in tds:
            if col == 0:
                col += 1
                prog += 1
                continue
            else:
                if td.get_text() == "":
                    prog += 1
                    continue
                content = td["title"]
                splCtt = content.split(";")
                if splCtt[1] == "":
                    content = content.replace(";;;", ";")
                    splCtt = content.split(";")

                for i in range(0, (len(splCtt) - 1), 2):
                    splEven = splCtt[i + 1].split(",")
                    weeks = splEven[0].replace("(", "").replace(")", "").split(" ")
                    if len(splEven) == 2:
                        splEven[1] = splEven[1].replace("))", ")")
                        room = ptnCampus.sub("", splEven[1])
                        campus = ptnCampus.findall(splEven[1])[0]
                    else:
                        room = ""
                        campus = ""

                    for weekGroup in weeks:
                        weekNum = weekGroup.split("-")
                        if len(weekNum) == 1:
                            weekNum.append(weekNum[0])

                        if ptnSigDbl.search(weekNum[1]) is not None:
                            sigdbl = ptnSigDbl.findall(weekNum[1])[0]
                            weekNum[1] = ptnSigDbl.sub("", weekNum[1])
                        else:
                            sigdbl = "无"
                        listSch.append(dict({
                            "code": ptnCode.findall(splCtt[i])[0],
                            "name": splCtt[0].split("(")[0],
                            "day": col,
                            "beginTime": row,
                            "duration": td["rowspan"],
                            "startWeek": int(weekNum[0]),
                            "endWeek": int(weekNum[1]),
                            "sigdbl": ("" if sigdbl is None else sigdbl),
                            "room": room,
                            "campus": campus.replace("(", "").replace(")", "")
                        }))

                        prog += 1
                        print("\r", end="")
                        print("进度: {}/{}".format(prog, target), end="")
                        sys.stdout.flush()

            col += 1
        row += 1

    json.dump(listSch, open("schedule/scheduleList.json", "w", encoding="UTF-8"), ensure_ascii=False)


if __name__ == "__main__":
    main()
