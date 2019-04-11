import os.path
import datetime

def Print(str):
    return None
    #print(str)

def PrintDate(date):
    str = date.strftime("%Y-%m-%d")
    Print(str)

def PrintDateList(date_list):
    for date in date_list:
        PrintDate(date)

def PrintTime(time):
    Print(time.strftime("%H:%M"))

def PrintTimeList(time_list):
    for time in time_list:
        PrintTime(time)

def IntToStr(n):
    return "%d" % n

def TimeToStr(time):
    if isinstance(time, str):
        return time
    return time.strftime("%H:%M")

def DateToStr(date):
    return date.strftime("%Y-%m-%d")

def GetOutputPathFromInputPath(original_path):
    full_list = os.path.split(original_path)
    file_path = full_list[len(full_list) - 1]
    file_list = os.path.splitext(file_path)
    save_path = file_list[len(file_list) - 2] + "_统计结果" + file_list[len(file_list) - 1]
    save_path = os.path.join(os.path.dirname(original_path), save_path)
    save_path = os.path.normpath(os.path.normcase(save_path))
    return save_path

# 9.1-9.9.xlsx
def GetDatesFromPath(path):
    full_list = os.path.split(path)
    file_path = full_list[len(full_list) - 1]
    file_list = os.path.splitext(file_path)
    date_str = file_list[len(file_list) - 2]
    return date_str

def GetDateListFromPath(path):
    date_str = GetDatesFromPath(path)
    date_list = date_str.split("-")
    return date_list

def IsAllDigit(list):
    for val in list:
        if not val.isdigit():
            return False
    return True

def CharListToDigitList(char_list):
    digit_list = []
    for val in char_list:
        n = int(val)
        digit_list.append(n)
    return digit_list

def ValidDate(date1, date2):
    b1 = date1[0] != date2[0]
    b2 = date1[0] < 1
    b3 = date1[0] > 12
    b4 = date1[1] < 1
    b5 = date1[1] > 31
    b6 = date2[1] < 1
    b7 = date2[1] > 31
    if b1 or b2 or b3 or b4 or b5 or b6 or b7:
        return False
    return True

# 9.1
def StrIsValidDate(str):
    try:
        d = datetime.datetime.strptime(str, "%m.%d")
        return True
    except:
        return False
    return True

def StrToDate(str):
    d = datetime.datetime.strptime(str, "%m.%d")
    return d