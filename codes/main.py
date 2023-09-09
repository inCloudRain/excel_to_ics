import datetime

import openpyxl
from ics import Calendar, Event

begin_date = datetime.date(2023, 8, 28)  # 填写第一周的周一日期
wb = openpyxl.load_workbook('../files/example.xlsx')  # 课程表文件，只支持xlsx
time_delta = '+08:00'  # 与UTC的时差

delta_dict = {'B': 0,
              'C': 1,
              'D': 2,
              'E': 3,
              'F': 4}

sheet = wb.active

cal = Calendar()


def add_event():
    e = Event()
    e.name = c_name + c_teacher
    e.begin = f'{date}T{start_time}{time_delta}'
    e.end = f'{date}T{end_time}{time_delta}'
    e.location = location
    cal.events.add(e)


for i in ['B', 'C', 'D', 'E', 'F']:
    for j in [str(i) for i in range(2, 8)]:
        v: str = sheet[i + j].value
        if v != None:
            vl = v.split('\n')
            cs = [vl[2 * i] + '\n' + vl[2 * i + 1] for i in range(int(len(vl) / 2))]
            for c in cs:
                cl = c.split('\n')
                l1 = cl[0].split(' ')
                l2 = cl[1][1:-1].split(' ')
                c_name = l1[0]
                c_teacher = l1[2]
                week_l = l2[0].split(',')
                location = l2[1]
                time = sheet['A' + j].value
                start_time = time.split('~ ')[0]
                end_time = time.split('~ ')[1]
                for week in week_l:
                    if '-' not in week:
                        date = begin_date + datetime.timedelta(weeks=int(week) - 1, days=delta_dict[i])
                        add_event()
                    else:
                        if '单' in week:
                            week = week.replace('单', '')
                            wl = week.split('-')
                            start_w = wl[0]
                            end_w = wl[1]
                            cl = [w for w in range(int(start_w), int(end_w) + 1) if w % 2 == 1]

                        elif '双' in week:
                            week = week.replace('双', '')
                            wl = week.split('-')
                            start_w = wl[0]
                            end_w = wl[1]
                            cl = [w for w in range(int(start_w), int(end_w) + 1) if w % 2 == 0]
                        else:
                            wl = week.split('-')
                            start_w = wl[0]
                            end_w = wl[1]
                            cl = range(int(start_w), int(end_w) + 1)
                        for w in cl:
                            date = begin_date + datetime.timedelta(weeks=w - 1, days=delta_dict[i])
                            add_event()
open('../files/out.ics', mode='w', encoding='u8').writelines(cal)
