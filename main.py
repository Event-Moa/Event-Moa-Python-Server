from pyhwpx import Hwp
from openpyxl import Workbook
import datetime
import json

def getJson(): # json 읽기
    with open("data.json", "r") as f:
        data = json.load(f)


def matchDay(day): # 요일 한글 적용
    if(day == 0):
        return " \n(일)"
    elif (day == 1):
        return " \n(월)"
    elif (day == 2):
        return " \n(화)"
    elif (day == 3):
        return " \n(수)"
    elif (day == 4):
        return " \n(목)"
    elif (day == 5):
        return " \n(금)"
    elif (day == 6):
        return " \n(토)"

def insertDate(startDate):
    sixDays = datetime.timedelta(days = 6)
    date1 = datetime.datetime.strptime(startDate, '%Y-%m-%d')
    date2 = date1 + sixDays

    totalDate = str(date1.strftime('%Y. %m. %d')) + ' ~ ' + str(date2.strftime('%m. %d.'))

    return totalDate


def changeEvent(event): # 행사 데이터 한글 작성 방식으로 변환
    if(len(event) == 8):
        event[0] = ""
        event.pop()
    else:
        date = datetime.datetime.strptime(event[0], '%Y-%m-%d')
        event[0] = str(date.strftime('%m. %d.')) + matchDay(int(str(date.strftime('%w'))))

    event[4] = str(event[4]) + "명"
    return event

def checkDate(preDate, curDate): # 빈 날짜 채우기
    date1 = datetime.datetime.strptime(preDate, '%Y-%m-%d')
    date2 = datetime.datetime.strptime(curDate, '%Y-%m-%d')
    diff = int((str(date2 - date1))[0])

    if(diff == 0):
        return False
    elif (diff == 1):
        return True
    else:
        blankDates = []
        blankDate = date1
        for i in range(diff - 1):
            oneDay = datetime.timedelta(days = 1)
            blankDate = blankDate + oneDay
            tmpDate = str(blankDate.strftime('%m. %d'))

            tmpDate = tmpDate + matchDay(int(str(blankDate.strftime('%w'))))
            blankDates.append([tmpDate, "", "", "", "", "", ""])

        return blankDates


filePath = "..\\주간행사표샘플.hwp"
hwp = Hwp()
hwp.open(filePath)

event = [["2024-05-30", "09:00", "중앙재난안정대책본부 영상회의(의료계 파업)", "재난안전상황실", 4, "재난안전과", ""],
         ["2024-05-30", "09:00", "중앙재난안정대책본부 영상회의(의료계 파업)", "재난안전상황실", 4, "재난안전과", ""],
         ["2024-06-02", "09:00", "중앙재난안정대책본부 영상회의(의료계 파업)", "재난안전상황실", 4, "재난안전과", ""]]

preDate = ""
row = 1
eventRow = 0

try:
    while(True):
        if (preDate == ""):
            preDate = event[eventRow][0]
            hwp.put_field_text(str(row), changeEvent(event[eventRow]))
            hwp.put_field_text("DATE", insertDate(preDate))
            row = row + 1
            eventRow = eventRow + 1

        else:
            blank = checkDate(preDate, event[eventRow][0])
            if(blank == True):
                hwp.put_field_text(str(row), changeEvent(event[eventRow]))
                row = row + 1
                eventRow = eventRow + 1

            elif(blank == False):
                event[eventRow].append("")
                hwp.put_field_text(str(row), changeEvent(event[eventRow]))
                row = row + 1
                eventRow = eventRow + 1

            else:
                for i in range(0, len(blank)):
                    hwp.put_field_text(str(row),blank[i])
                    row = row + 1

                hwp.put_field_text(str(row), changeEvent(event[eventRow]))
                row = row + 1
                eventRow = eventRow + 1



except IndexError:
    # 파일 저장 및 닫기
    hwp.SaveAs("..\\output.hwp")
    hwp.Quit()
