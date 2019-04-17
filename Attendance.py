import os.path
import datetime
import os
import threading

import babel
import babel.numbers

import utils
import excel_handler

#from time import sleep

class GoalItem:
    def __init__(self, earliest, last):
        self.earliest = earliest
        self.last = last
        
class TimeItem:
    def __init__(self, date, time):
        str = date + " " + time
        self.datetime = datetime.datetime.strptime(str, "%Y-%m-%d %H:%M")
        self.test = 0
    def DateTime(self):
        return self.datetime
    def Date(self):
        return self.datetime.date()
    def Time(self):
        return self.datetime.time()

class UserItem:
    def __init__(self, id, name, department):
        self.id = id;
        self.name = name
        self.department = department
        self.times = {}
        self.goals = {}
    def AddTime(self, time):
        d = time.Date()
        t = time.Time()
        if d not in self.times:
            self.times[d] = []
        self.times[d].append(t)
    def Times(self):
        return self.times
    def AddGoal(self, date, goal):
        self.goals[date] = goal
    def Goals(self):
        return self.goals

def ParseToUserItems(table):
    ulist = {}
    for row in range(table.max_row):
        actual_row = row + 1
        if actual_row == 1:
            continue
        ti = TimeItem(table.cell(actual_row, 7).value, table.cell(actual_row, 8).value)
        id = table.cell(actual_row, 2).value
        nid = int(id)
        if nid not in ulist:
            name = table.cell(actual_row, 3).value
            department = table.cell(actual_row, 4).value
            ulist[nid] = UserItem(id, name, department)
        ulist[nid].AddTime(ti)
    return ulist

def GetAttendance(users, date_list, ulist):
    # 单人，单天分析
    for id in users:
        utils.Print("用户: " + utils.IntToStr(id))
        if id not in ulist:
            # 1、始终未打卡
            continue
        user_info = ulist[id]
        user_times = user_info.Times()
        for date in date_list:
            if date not in user_times:
                # 2.1、当天没有打卡记录
                continue
            # 当天打卡记录
            time_list = list(user_times[date])
            # 2.2、空列表代表当天没有打卡记录
            list_len = len(time_list)
            if list_len == 0:
                continue
            # 排序计算出最早最晚
            time_list.sort()
            # 最早
            earliest = time_list[0]
            # 最晚
            last = time_list[list_len - 1]
            # TODO: 如果出现早退，查看是否在第二天打卡
            #utils.Print("早: " + utils.TimeToStr(earliest) + "    晚: " + utils.TimeToStr(last))
            goal_item = GoalItem(earliest, last)
            user_info.AddGoal(date, goal_item)
    return None

def ParseFile(path, dates):
    table = excel_handler.LoadTable(path)
    users = excel_handler.GetUsers(table)
    user_items = ParseToUserItems(table)
    date_list = list(dates)
    date_list.sort()
    #utils.PrintDateList(date_list)
    GetAttendance(users, date_list, user_items)
    work_book = excel_handler.SaveToWorkbook(users, date_list, user_items)
    excel_handler.SaveToExcel(work_book, path)
    return None



# Dialog
try:
    import tkinter as tk
    from tkinter import filedialog
    from tkinter import ttk
    from tkinter import Grid
    from tkinter import messagebox
except ImportError:
    import Tkinter as tk
    import ttk

from tkcalendar import Calendar, DateEntry

def ParseAndSaveToFile(path, date_list):
    #sleep(10)
    # parse file
    ParseFile(path, date_list)
    # save file
    dir_name = os.path.dirname(path)
    os.startfile(dir_name)
    # Backup UI
    process.stop()
    process.grid_remove()
    btn_choose_file.grid()
    btn_parse.grid()

def OnChooseFile():
        path = filedialog.askopenfilename(initialdir = "/", 
                                          title = "选择Excel文件", 
                                          filetypes = (("Excel files", "*.xlsx"), ("All files", "*.*"))
                                          )
        cal.calevent_remove("all")
        entry.delete(0, len(entry.get()))

        entry.insert(0, path)

        date_list = utils.GetDateListFromPath(path)
        if (len(date_list) == 2):
            date_start = None
            date_end = None
            start = date_list[0]
            end = date_list[1]
            ss = start.split(".")
            ee = end.split(".")
            start = "%d." % (cal.datetime.today().year) + start
            end = "%d." % (cal.datetime.today().year) + end
            if len(ss) == 2 and utils.IsAllDigit(ss) and len(ee) == 2 and utils.IsAllDigit(ee):
                ssn = utils.CharListToDigitList(ss);
                een = utils.CharListToDigitList(ee);
                if not utils.ValidDate(ssn, een):
                    messagebox.showinfo("提示", "日期起止点必须在一个月哦")
                    return None
                date_start = datetime.datetime.strptime(start, "%Y.%m.%d").date()
                date_end = datetime.datetime.strptime(end, "%Y.%m.%d").date()
                cal.selection_set(date_start)
                flag = 0;
                for i in range((date_end - date_start).days + 1):
                    day = date_start + datetime.timedelta(days=i)
                    if day.weekday() >= 5:
                        continue
                    if flag == 0:
                        flag+=1
                        cal.selection_set(day)
                    cal.calevent_create(day, "Hello", 'reminder')
            else:
                messagebox.showinfo("提示", "选中的文件需要类似'9.1-9.9.xlsx'的日期格式，才能确定考勤日期，日期好像没有办法解析哦")
                entry.delete(0, len(entry.get()))
        else:
            messagebox.showinfo("提示", "选中的文件需要类似'9.1-9.9.xlsx'的日期格式，才能确定考勤日期哦")
            entry.delete(0, len(entry.get()))
        return None

def OnParse():
    # get date list
    date_list = []
    ev_ids = cal.get_calevents(tag = "reminder")
    if len(ev_ids) == 0:
        messagebox.showinfo("提示", "需要选中考勤的日期哦")
        return None
    for ev_id in ev_ids:
        date = cal.calevent_cget(ev_id, 'date')
        date_list.append(date)
        utils.Print(date)

    # get file path
    path = entry.get()
    if len(path) == 0:
        messagebox.showinfo("提示", "好像没有指定考勤文件的路径哦")
        return None

    # check whether the file name is valid
    path_date_list = utils.GetDateListFromPath(path)
    if len(path_date_list) != 2:
        messagebox.showinfo("提示", "文件名称的日期格式不正确哦(9.1-9.9.xlsx)")
        return None
    for dt in path_date_list:
        if not utils.StrIsValidDate(dt):
            messagebox.showinfo("提示", "文件名称的日期格式不正确哦(9.1-9.9.xlsx)")
            return None
    date1 = utils.StrToDate(path_date_list[0])
    date2 = utils.StrToDate(path_date_list[1])
    if date1.month != date2.month:
        messagebox.showinfo("提示", "起止日期必须在同一个月份哦")
        return None
    if date1.day > date2.day:
        messagebox.showinfo("提示", "结束日期必须大于起始日期哦")
        return None

    # show process
    process.grid()
    btn_parse.grid_remove()
    #entry.grid_remove()
    btn_choose_file.grid_remove()
    process.start()

    # New thread to parse file
    th_parse = threading.Thread(target = ParseAndSaveToFile, 
                                kwargs = {"path": path, "date_list": date_list})
    th_parse.setDaemon(True)
    th_parse.start()
    #td_parse.join()

    
    return None

def OnSelected(event):
    date = cal.selection_get()
    ev_ids = cal.get_calevents(date = date, tag = "reminder")
    if len(ev_ids) == 0:
        cal.calevent_create(date = date, text = 'Hello', tags = ['reminder'])
    else:
        cal.calevent_remove(ev_ids[0])
    return None

def OnMonthChanged(event):
    cal.calevent_remove("all")
    return None

root = tk.Tk()

Grid.rowconfigure(root, 0, weight = 1)
Grid.columnconfigure(root, 0, weight = 1)

cal = Calendar(root, 
               font="Arial 25", 
               locale='en_US', 
               showweeknumbers = False, 
               showothermonthdays = False,
               )
entry = ttk.Entry(root)
btn_choose_file = ttk.Button(root, text = "Choose file", command = OnChooseFile)
btn_parse = ttk.Button(root, text = "Parse", command = OnParse)
process = ttk.Progressbar(root, orient = "horizontal", mode = "indeterminate")

cal.grid(row = 0, column = 0, rowspan = 8, columnspan = 8, sticky = tk.N + tk.S + tk.W + tk.E)
entry.grid(row = 8, column = 0, columnspan = 7, sticky = tk.W + tk.E)
btn_choose_file.grid(row = 8, column = 7)
btn_parse.grid(row = 9, column = 0, rowspan = 1, columnspan = 8)
process.grid(row = 9, column = 0, rowspan = 1, columnspan = 8, sticky = tk.W + tk.E)
process.grid_remove()

cal.bind("<<CalendarSelected>>", OnSelected)
cal.bind("<<CalendarMonthChanged>>", OnMonthChanged)

cal.tag_config('reminder', background='deep pink', foreground='white')#background='royal blue'

root.mainloop()