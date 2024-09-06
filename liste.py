import xlsxwriter
import datetime



# input: tuple(år, måned, dag)
def make_interval (start, slutt):
    #start = datetime.datetime.strptime("21-06-2014", "%d-%m-%Y")
    #end = datetime.datetime.strptime("07-07-2014", "%d-%m-%Y")
    start_date = make_date(start, this_year)
    end_date = make_date(slutt, this_year)
    start = datetime.datetime.strptime(start_date, "%d-%m-%Y")
    while start.weekday() not in [1,3]:
        start = start + datetime.timedelta(days=1)
    end = datetime.datetime.strptime(end_date, "%d-%m-%Y")
    date_generated = [start + datetime.timedelta(days=x) for x in range(0, (end-start).days+1)]

    return date_generated

def make_month_intervals (dates, positions) -> dict:
    months_dates = {}
    months_pos = {}
    for date, pos in zip(dates, positions):
        try:
            months_dates[int_to_month(date.month)].append(date)
            months_pos[int_to_month(date.month)].append(pos)
        except:
            months_dates[int_to_month(date.month)] = [date]
            months_pos[int_to_month(date.month)] = [pos]

    return months_dates, months_pos

def make_positions (left, right):
    left_pos = []
    right_pos = []

    #tirsdag
    if left[0].weekday() == 1:
        motsatt = True
    #torsdag
    elif left[0].weekday() == 3:
        motsatt = False

    current_pos = 1
    count = 0

    for date in left:
        left_pos.append(stilling(current_pos))
        count += 1
        if motsatt:
            if count == 2:
                motsatt = False
                current_pos += 1
                if current_pos > 3:
                    current_pos = 1
                count = 0
        else:
            if count == 3:
                motsatt = True
                current_pos += 1
                if current_pos > 3:
                    current_pos = 1
                count = 0

    current_pos = 1
    count = 0

    #tirsdag
    if right[0].weekday() == 1:
        motsatt = True
    #torsdag
    if right[0].weekday() == 3:
        motsatt = False

    for date in right:
        right_pos.append(stilling(current_pos))
        count += 1
        if motsatt:
            if count == 2:
                motsatt = False
                current_pos += 1
                if current_pos > 3:
                    current_pos = 1
                count = 0
        else:
            if count == 3:
                motsatt = True
                current_pos += 1
                if current_pos > 3:
                    current_pos = 1
                count = 0

    return left_pos, right_pos


def print_interval (interval):
    for date in interval:
        print(date.strftime("%d-%m-%Y"))

def print_interval_days (interval):
    for date in interval:
        print(date.strftime("%d-%m-%Y"), day(date.weekday()))

def print_interval_days_pos (interval, pos):
    assert len(interval) == len(pos)
    for date, pos in zip(interval, pos):
        print(date.strftime("%d-%m-%Y"), day(date.weekday()), pos)
#fjerne søndager og mandager
def prune_interval (interval):
    pruned = []
    for date in interval:
        if date.weekday() not in [0,6]:
            pruned.append(date)
    return pruned


def split (interval):
    left = []
    right = []

    count_left = 0
    count_right = 0
    venstre = True
    tirsdag = False
    #hopper til første tirsdag eller torsdag
    #while interval[0].weekday() not in [1,3]:
    #    del interval[0]
    #tirsdag
    if interval[0].weekday() == 1:
        tirsdag = True
    #torsdag
    elif interval[0].weekday() == 3:
        tirsdag = False

    for date in interval:
        if venstre:
            left.append(date)
            count_left += 1
            if tirsdag:
                if count_left == 2:
                    tirsdag = False
                    venstre = False
                    count_left = 0
            else:
                if count_left == 3:
                    tirsdag = True
                    venstre = True
                    count_left = 0
        else:
            right.append(date)
            count_right += 1
            if tirsdag:
                if count_right == 2:
                    tirsdag = False
                    venstre = True
                    count_right = 0
            else:
                if count_right == 3:
                    tirsdag = True
                    venstre = False
                    count_right = 0

    return left, right


def split2 (interval):
    left = []
    left_pos = []
    right = []
    right_pos = []

    count_pos_left = 1
    count_pos_right = 1


    count1 = 0
    count2 = 0
    venstre = True
    motsatt = False
    for date in interval:
        if venstre:
            left.append(date)
            left_pos.append(stilling(count_pos_left))
            count1 += 1
            if motsatt:
                if count1 == 2:
                    motsatt = False
                    venstre = False
                    count1 = 0

                    count_pos_left += 1
                    if count_pos_left > 3:
                        count_pos_left = 1
            else:
                if count1 == 3:
                    venstre = False
                    motsatt = True
                    count1 = 0

                    count_pos_left += 1
                    if count_pos_left > 3:
                        count_pos_left = 1
        else:
            right.append(date)
            right_pos.append(stilling(count_pos_right))
            count2 += 1
            if motsatt:
                if count2 == 2:
                    motsatt = True
                    venstre = True
                    count2 = 0

                    count_pos_right += 1
                    if count_pos_right > 3:
                        count_pos_right = 1
            else:
                if count2 == 3:
                    motsatt = False
                    venstre = True
                    count2 = 0

                    count_pos_right += 1
                    if count_pos_right > 3:
                        count_pos_right = 1

    return left, right, left_pos, right_pos

def make_date (dato):
    make_date(dato, this_year)

def make_date (dato, year):
    dag = dato[0]
    if dag // 10 == 0:
        dag = "0" + str(dag)
    måned = dato[1]
    if måned // 10 == 0:
        måned = "0" + str(måned)
    år = year
    string = f"{dag}-{måned}-{år}"
    return string

def stilling (number):
    if number == 1:
        return (1,2,3)
    elif number == 2:
        return (3,1,2)
    elif number == 3:
        return (2,3,1)
    else:
        raise Exception("Ugyldig input, stillinger: 1-3")

def day (number):
    if number == 0:
        return "mandag"
    elif number == 1:
        return "tirsdag"
    elif number == 2:
        return "onsdag"
    elif number == 3:
        return "torsdag"
    elif number == 4:
        return "fredag"
    elif number == 5:
        return "lørdag"
    elif number == 6:
        return "søndag"
    else:
        raise Exception("Ugyldig input, dager: 0-6")

def int_to_month (number):
    if number == 1:
        return "januar"
    elif number == 2:
        return "februar"
    elif number == 3:
        return "mars"
    elif number == 4:
        return "april"
    elif number == 5:
        return "mai"
    elif number == 6:
        return "juni"
    elif number == 7:
        return "juli"
    elif number == 8:
        return "august"
    elif number == 9:
        return "september"
    elif number == 10:
        return "oktober"
    elif number == 11:
        return "november"
    elif number == 12:
        return "desember"
    else:
        raise Exception("Ugyldig input, måneder: 1-12")

def write_month_string_left (månad, worksheet, row, col, format):
    #måned = month(number)
    for letter in månad:
        worksheet.write(row+1, col, letter, format)
        row += 1

def write_month_string_right (månad, worksheet, row, col, format):
    #måned = month(number)
    for letter in månad:
        worksheet.write(row+1, col + 6, letter, format)
        row += 1

def write_month_left (int, pos, worksheet, row, col):
    måned = int_to_month(int[0].month)

    for i, (date, pos) in enumerate(zip(int, pos)):
        if (i == len(int)-1) and i > len(måned):
            worksheet.write(row, col, date.day, dato_nederst)
            worksheet.write(row, col + 1, "", nederst)
            worksheet.write(row, col + 2, pos[0], venstre_nederst)
            worksheet.write(row, col + 3, pos[1], venstre_nederst)
            worksheet.write(row, col + 4, pos[2], venstre_nederst)

        else:
            worksheet.write(row, col, date.day, dato)
            worksheet.write(row, col + 2, pos[0], venstre)
            worksheet.write(row, col + 3, pos[1], venstre)
            worksheet.write(row, col + 4, pos[2], venstre)

            row += 1

            if (i == len(int)-1) and i <= len(måned):
                for i in range(len(måned) - i) :
                    worksheet.write(row, col, "", dato)
                    worksheet.write(row, col + 2, "", venstre)
                    worksheet.write(row, col + 3, "", venstre)
                    worksheet.write(row, col + 4, "", venstre)

                    row += 1
                worksheet.write(row, col, "", dato_nederst)
                worksheet.write(row, col + 1, "", nederst)
                worksheet.write(row, col + 2, "", venstre_nederst)
                worksheet.write(row, col + 3, "", venstre_nederst)
                worksheet.write(row, col + 4, "", venstre_nederst)

def write_month_right (int, pos, worksheet, row, col):
    måned = int_to_month(int[0].month)

    for i, (date, pos) in enumerate(zip(int, pos)):
        if (i == len(int)-1) and i > len(måned):
            worksheet.write(row, col + 6, date.day, dato_nederst)
            worksheet.write(row, col + 7, "", nederst)
            worksheet.write(row, col + 8, pos[0], høyre_nederst)
            worksheet.write(row, col + 9, pos[1], høyre_nederst)
            worksheet.write(row, col + 10, pos[2], høyre_nederst)
        else:
            worksheet.write(row, col + 6, date.day, dato)
            worksheet.write(row, col + 8, pos[0], høyre)
            worksheet.write(row, col + 9, pos[1], høyre)
            worksheet.write(row, col + 10, pos[2], høyre)

            row += 1

            if (i == len(int)-1) and i <= len(måned):
                for i in range(len(måned) - i):
                    worksheet.write(row, col + 6, "", dato)
                    worksheet.write(row, col + 8, "", høyre)
                    worksheet.write(row, col + 9, "", høyre)
                    worksheet.write(row, col + 10, "", høyre)

                    row += 1
                worksheet.write(row, col + 6, "", dato_nederst)
                worksheet.write(row, col + 7, "", nederst)
                worksheet.write(row, col + 8, "", høyre_nederst)
                worksheet.write(row, col + 9, "", høyre_nederst)
                worksheet.write(row, col + 10, "", høyre_nederst)






start_date = (29,1)
end_date = (31,5)
this_year = 2019

interval = []
igjen = "ja"
this_year = int(input("årstall: "))
while igjen == "ja":
    start_input = input("start-dato: ").split("-")
    start_date = (int(start_input[0]), int(start_input[1]))
    end_input = input("slutt-dato: ").split("-")
    end_date = (int(end_input[0]), int(end_input[1]))
    interval.extend(make_interval(start_date, end_date))
    igjen = input("Vil du legge til flere datoer? ")
#interval = make_interval(start_date, end_date)
interval = prune_interval(interval)
#print_interval_days(interval)

i1, i2 = split(interval)
pos1, pos2 = make_positions(i1,i2)
#januar = make_month_intervals(interval)["januar"]
#i1, i2, pos1, pos2 = split(januar)

left_months_dates, left_months_pos = make_month_intervals(i1, pos1)
right_months_dates, right_months_pos = make_month_intervals(i2, pos2)

#print_interval_days(i1)
#print("---------------")
#print_interval_days(i2)
print_interval_days_pos(i1, pos1)
print("---------------")
print_interval_days_pos(i2, pos2)

m1 = int_to_month(interval[0].month)
m2 = int_to_month(interval[-1].month)
workbook = xlsxwriter.Workbook(f'etasjefordeling {m1}-{m2} {this_year}.xlsx')
worksheet = workbook.add_worksheet()

etasje = workbook.add_format()
etasje.set_align('center')
etasje.set_bold()
etasje.set_border(1)
etasje.set_top(2)
etasje.set_bottom(2)
etasje.set_font_name("Footlight MT Light")

år = workbook.add_format()
år.set_bold()
år.set_align('center')
år.set_font_name("Footlight MT Light")

høyre = workbook.add_format()
høyre.set_align('center')
høyre.set_border()
høyre.set_font_name("Footlight MT Light")

høyre_nederst = workbook.add_format()
høyre_nederst.set_align('center')
høyre_nederst.set_border()
høyre_nederst.set_font_name("Footlight MT Light")
høyre_nederst.set_bottom(2)

venstre = workbook.add_format()
venstre.set_align('center')
venstre.set_border(1)
venstre.set_pattern(1)
venstre.set_bg_color("#DCE6F1")
venstre.set_font_name("Footlight MT Light")

venstre_nederst = workbook.add_format()
venstre_nederst.set_align('center')
venstre_nederst.set_border(1)
venstre_nederst.set_pattern(1)
venstre_nederst.set_bg_color("#DCE6F1")
venstre_nederst.set_font_name("Footlight MT Light")
venstre_nederst.set_bottom(2)

dato = workbook.add_format()
dato.set_border(1)
dato.set_font_name("Footlight MT Light")
dato.set_align("left")

dato_nederst = workbook.add_format()
dato_nederst.set_border(1)
dato_nederst.set_font_name("Footlight MT Light")
dato_nederst.set_align("left")
dato_nederst.set_bottom(2)

måned = workbook.add_format()
måned.set_left(1)
måned.set_right(1)
måned.set_font_name("Footlight MT Light")
måned.set_align('center')
måned.set_bold()

lag1 = workbook.add_format()
lag1.set_pattern(1)
lag1.set_bg_color("#DCE6F1")
lag1.set_bottom(2)
lag1.set_font_name("Footlight MT Light")
lag1.set_bold()

lag2 = workbook.add_format()
lag2.set_pattern(1)
lag2.set_bg_color("#FFFFFF")
lag2.set_bottom(2)
lag2.set_font_name("Footlight MT Light")
lag2.set_bold()

nederst = workbook.add_format()
nederst.set_bottom(2)

worksheet.write(0,2,"O", etasje)
worksheet.write(0,3,"P", etasje)
worksheet.write(0,4,"B", etasje)
worksheet.write(0,5, this_year, år)
worksheet.write(0,8,"O", etasje)
worksheet.write(0,9,"P", etasje)
worksheet.write(0,10,"B", etasje)

"""row = 1
col = 0
for date, pos in zip(i1, pos1):
    worksheet.write(row, col, date.day, dato)
    worksheet.write(row, col + 2, pos[0], venstre)
    worksheet.write(row, col + 3, pos[1], venstre)
    worksheet.write(row, col + 4, pos[2], venstre)
    row += 1

row = 1
col = 6
for date, pos in zip(i2, pos2):
    worksheet.write(row, col, date.day, dato)
    worksheet.write(row, col + 2, pos[0], høyre)
    worksheet.write(row, col + 3, pos[1], høyre)
    worksheet.write(row, col + 4, pos[2], høyre)
    row += 1"""

#write_month_string(1, worksheet, 1, 1, måned)

#write_month(i1, pos1, i2, pos2, worksheet, 1, 0)

row = 1
col = 0
for month, interval in left_months_dates.items():
    write_month_string_left(month, worksheet, row, col + 1, måned)
    pos = left_months_pos[month]
    write_month_left(interval, pos, worksheet, row, col)
    row += max(len(interval), len(month)+2)

row = 1
for month, interval in right_months_dates.items():
    write_month_string_right(month, worksheet, row, col + 1, måned)
    pos = right_months_pos[month]
    write_month_right(interval, pos, worksheet, row, col)
    row += max(len(interval), len(month)+2)

#endrer cellebredde
worksheet.set_column(0,4,6)
worksheet.set_column(6,10,6)

row += 5
for i in range(3):
    for j in range(5):
        if j == 1:
            worksheet.write(row+i, col+j, f"{i+1}:", lag1)
            worksheet.write(row+i, col+j+6, f"{i+1}:", lag2)
        else:
            worksheet.write(row+i, col+j, "", lag1)
            worksheet.write(row+i, col+j+6, "", lag2)

workbook.close()
