import gspread
import time

sa = gspread.service_account(filename=r"C:\Users\Angel\Downloads\bionic-medley-428717-m7-0c4210d09c81.json")
sh = sa.open("Botpress Appointments")


wks1 = sh.worksheet("Booked Appointments")
wks2 = sh.worksheet("Available Appointments")

date_rows = {
    wks2.cell(2,1).value : 2,
    wks2.cell(3,1).value : 3,
    wks2.cell(4,1).value : 4
}

time_cols = {
    "10:00 AM" : 2,
    "11:00 AM" : 3,
    "12:00 PM" : 4,
    "1:00 PM" : 5,
    "2:00 PM" : 6,
    "3:00 PM" : 7
}

updated_booked_rows = 2

while True:
    i = updated_booked_rows
    date_i = wks1.cell(i,2).value
    time_i = wks1.cell(i,3).value
    while not(str(date_i.__class__)=="<class 'NoneType'>") and not(str(time_i.__class__)=="<class 'NoneType'>"):
        date_i = wks1.cell(i,2).value
        time_i = wks1.cell(i,3).value
        if str(date_i.__class__)=="<class 'NoneType'>" or str(time_i.__class__)=="<class 'NoneType'>":
            break
        elif wks2.cell(date_rows[date_i], time_cols[time_i]).value != "Not Available":
            wks2.update_cell(date_rows[date_i], time_cols[time_i], "Not Available")
        i += 1
    
    updated_booked_rows = i
    print(updated_booked_rows)
    time.sleep(15)