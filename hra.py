import xlrd
import xlwt
from xlwt import Workbook
import sys
import re
from datetime import datetime
import copy
import datetime
sys.path
sys.executable

location = "excel_sheet.xlsx"
wb  = xlrd.open_workbook(location)
sheet_no = len(wb.sheet_names())
sheet = wb.sheet_by_index(0)
single_room_occupancy=["occupied","OCCUPIED","Occupied","occupied ","OCCUPIED ","Occupied "]
single_room_identification=["SINGLE","single","Single"]
single_room_vacancy=["blocked","cancelled","BLOCKED","CANCELLED","","Blocked","Cancelled"]
btech_degree=["B.Tech","B.TECH","b.tech","B.tech","B.Tech ","B.TECH ","b.tech ","B.tech "]
mtech_degree=["M.TECH","M.Tech","m.tech","M.tech","M.TECH ","M.Tech ","m.tech ","M.tech "]
mba_degree =["MBA","M.B.A","mba","Mba","m.b.a","M.b.a","MBA ","M.B.A ","mba ","Mba ","m.b.a ","M.b.a "]
msc_degree=["MSC","msc","Msc","M.s.c","M.S.C","m.s.c","MSC ","msc ","Msc ","M.s.c ","M.S.C ","m.s.c "]
msctech_degree = ["msc(tech) ","MSC(tech) ", "Msc(tech) ","MSC(TECH) ","msc(tech)","MSC(tech)", "Msc(tech)","MSC(TECH)","MSC(TEch)","MSC (TEch)","MSc(Tech)",]
mca_degree =["MCA","mca","Mca","M.C.A","m.c.a","MCA ","mca ","Mca ","M.C.A ","m.c.a "]
total_no_of_students=[]
double_room_occupancy =["occupied","OCCUPIED","Occupied","occupied ","OCCUPIED ","Occupied "]
double_room_identification=["DOUBLE","Double","double"]
single_rooms_in_Ablock=[10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,43,44,45,46,47,48,49,50,51,52,53,60]
double_rooms_in_Bblock=[1,2,3,4,5,6,7,8,9,10,11,33,34,35,36,37,38,39,40,41,42,43,44,56,57,58,59,60,61]
double_rooms_in_Ablock=[1,2,3,4,5,6,7,8,9,31,32,33,34,35,36,37,38,39,40,41,42,54,55,56,57,58,59]
single_rooms_in_Bblock=[12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,44,45,46,47,48,49,50,51,52,53,54,55]


def single_room_vacancies():
    count_of_single_rooms=[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
    floor=0
    for sheets in range(0,sheet_no-1):
        worksheet1 = wb.sheet_by_index(sheets)
        count_of_eachfloor = 0
        print("-----------------------------------------------------------")
        print(worksheet1.name)
        for j in range(worksheet1.nrows):
            if worksheet1.cell_value(j,3) in single_room_identification:
                if not worksheet1.cell_value(j,4) in single_room_occupancy:
                    if not worksheet1.cell_value(j,2) == '':
                        print(worksheet1.cell_value(j,2))
                    count_of_eachfloor +=1
        count_of_single_rooms[floor]= count_of_eachfloor
        floor += 1
    print(" ")
    print("__________________________________________________________________________")
    print("----  floor ---------- |||------- vacencies -------|| ----Filled --------'")
    print("-----------------------------A-Block -------------------------------------")
    print("--------------------------------------------------------------------------")
    for floor in range(10):
        print("    A"+str(floor)+"                         "+str(count_of_single_rooms[floor])+"          "+str(32-count_of_single_rooms[floor]))
        print("----------------------------------------------------")
    print("-------------------- B-Block ----------------------")
    print("---------------------------------------------------")
    for floor in range(10,20):
        print("    B"+str(floor-10)+"                      "+str(count_of_single_rooms[floor])+"                "+str(32-count_of_single_rooms[floor]))
        print("----------------------------------------------------")



def double_room_vacancies():
    double_rooms_list=[]
    for sheets in range(0,sheet_no-2):
        worksheet1 = wb.sheet_by_index(sheets)
        print("-----------------------------------------------------------")
        print(worksheet1.name)
        for j in range(worksheet1.nrows):
            if worksheet1.cell_value(j,3) in double_room_identification:
                if not worksheet1.cell_value(j,4) in double_room_occupancy:
                    if not worksheet1.cell_value(j,2) == '':
                        double_rooms_list.append(worksheet1.cell_value(j,2))
    single_double_room=[]
    double_double_room=[]
    for i in range(0,len(double_rooms_list)-1):
        if not double_rooms_list[i] == double_rooms_list[i+1]:
            single_double_room.append(double_rooms_list[i])
        else:
            double_double_room.append(double_rooms_list[i])
            i=i+1
    single=[]
    for i in double_rooms_list:
        if not (i in double_double_room):
            single.append(i)
    print(single)
    print("------------------------------------------------ ----------------------")
    print("Singles in double room vacancies is ")
    for single_vacant_room in single:
        print(single_vacant_room)
    print("__________________________________________________________________________")
    print("Full double rooms vacant are:")
    for double_room in double_double_room:
        print(double_room)
    print("________________________________________________________________________")
    print("total number of single vacancies in double room is "+str(len(single)))
    print("total full double vacant rooms where:"+str(len(double_double_room)))
    print("total rooms vacancies in double rooms are  :"+str((len(single))+(len(double_double_room) *2)))

def number_of_days(date):
    today = datetime.date.today()
    date = date.split("/")
    someday = datetime.date(int(date[2]), int(date[0]), int( date[1]))
    diff = today - someday
    return diff.days

def filled_rooms():
    print("filled")
    print(" ")
    print(" ")
    print("   1. A0-GF          11. B0-GF")
    print("   2. A1             12. B1")
    print("   3. A2             13. B2")
    print("   4. A3             14. B3")
    print("   5. A4             15. B4")
    print("   6. A5             16. B5")
    print("   7. A6             17. B6")
    print("   8. A7             18. B7")
    print("   9. A8             19. B8")
    print("   10. A9            20. B9")
    print(" ")
    try:
        while True:
            option = input("Enter your choice\n")
            if option.isdigit():
                if int(option) >= 1 and int(option) <= 20:
                    break
                else:
                    print("enter valid choice")
    except NameError as e:
        print(e)
    worksheet1= wb.sheet_by_index(int(option)-1)
    print("Filled single rooms are :")
    for j in range(worksheet1.nrows):
        if worksheet1.cell_value(j,3) in single_room_identification:
            if  worksheet1.cell_value(j,4) in single_room_occupancy:
                if not worksheet1.cell_value(j,2) == '':
                    for details in range(0,worksheet1.ncols):
                        print(worksheet1.cell_value(j,details), end=' ')
                    print(" ")
                    print(" ")
    print(" ")
    print("______________________________________________________________________________")
    print("Filled double  rooms are: ")
    worksheet1= wb.sheet_by_index(int(option)-1)
    for j in range(worksheet1.nrows):
        if worksheet1.cell_value(j,3) in double_room_identification:
            if  worksheet1.cell_value(j,4) in double_room_occupancy:
                if not worksheet1.cell_value(j,2) == '':
                    print(worksheet1.cell_value(j,2))
                    for details in range(0,worksheet1.ncols):
                        print(worksheet1.cell_value(j,details), end=' ')
                    print(" ")
                    print(" ")

def search_detail():
    print("search")
    find1= str(input("Enter the name that you are looking for\n"))
    print(" ")
    for sheets in range(0,int(sheet_no)-2):
        worksheet1= wb.sheet_by_index(sheets)
        for rows in  range(0,worksheet1.nrows):
            if re.search(find1.lower(),worksheet1.cell_value(rows, 5).lower()):
                list1=[]
                for details in range(0,worksheet1.ncols-1):
                    list1.append(worksheet1.cell_value(rows, details))
                days = number_of_days(worksheet1.cell_value(rows,details+1))
                list1.append(days)
                display_student_details(list1)
                print("  ")
                print("  ")





def display_all_rooms():
    print("display")
    sheet1=0
    print(" ")
    print(" ")
    print("   1. A0-GF          11. B0-GF")
    print("   2. A1             12. B1")
    print("   3. A2             13. B2")
    print("   4. A3             14. B3")
    print("   5. A4             15. B4")
    print("   6. A5             16. B5")
    print("   7. A6             17. B6")
    print("   8. A7             18. B7")
    print("   9. A8             19. B8")
    print("   10. A9            20. B9")
    print(" ")
    try:
        while True:
            option = input("Enter your choice\n")
            if option.isdigit():
                if int(option) >= 1 and int(option) <= 20:
                    break
                else:
                    print("enter valid choice")
    except NameError as e:
        print(e)
    worksheet1 = wb.sheet_by_index(int(option)-1)
    for j in range(0,worksheet1.nrows):
        if j==0:
            pass
        list1=[]
        for i in range(worksheet1.ncols-2):
            list1.append(worksheet1.cell_value(j, i))
        if not (worksheet1.cell_value(j,i+1) == 'Date of joining' or worksheet1.cell_value(j,i+1) == 'N/A' or  worksheet1.cell_value(j,i+1) == ''):
            days = number_of_days(str(worksheet1.cell_value(j,i+1)))
            list1.append(days)
            display_student_details(list1)
            print(" ")
        print(" ")

def degree_pursuing():
    print("--------------------------------------")
    print("1. B.Tech")
    print("2. M.Tech")
    print("3. Msc")
    print("4. MCA")
    print("5. Msc(Tech)")
    print("6. MBA")
    try:
        while True:
            option = input("Enter your choice\n")
            if option.isdigit():
                if int(option) >= 1 and int(option) <= 6:
                    break
                else:
                    print("enter valid choice")
    except NameError as e:
        print(e)
    if int(option) ==1:
        course("btech")
    elif int(option) ==2:
        course("mtech")
    elif int(option) ==3:
        course("msc")
    elif int(option) ==4:
        course("mca")
    elif int(option) ==5:
        course("msc_tech")
    elif int(option) ==6:
        course("mba")
    else:
        print("enter correct option")


def course(degree):
    degree1=''
    #total_no_of_students=[]
    if degree == "btech":
        degree1 = btech_degree
    elif degree == "mtech":
        degree1 = mtech_degree
    elif degree == "mca":
        degree1 = mca_degree
    elif degree == "msc":
        degree1 = msc_degree
    elif degree == "msc_tech":
        degree1 = msctech_degree
    elif degree == "mba":
        degree1 = mba_degree
    count_of_all_btech=0
    for sheets in range(0,sheet_no-2):
        worksheet1 = wb.sheet_by_index(sheets)
        count_of_eachfloor = 0
        print("------------------------------------------------------------------------------")
        print(worksheet1.name)
        for j in range(worksheet1.nrows):
            if worksheet1.cell_value(j,7) in degree1:
                count_of_eachfloor +=1
                count_of_all_btech +=1
                for details in range(worksheet1.ncols):
                    print(worksheet1.cell_value(j,details), end='  ')
                print(" ")
                print(" ")
        print("total number of "+degree+" students in "+str(worksheet1.name)+" is "+str(count_of_eachfloor))
    total_no_of_students.append(int(count_of_all_btech))
    print("--------------------------------------------------------------------")
    print("Total number of "+degree+" students are "+str(count_of_all_btech))
    print("    ")
    print("    ")
    print("    ")

def total_number_of_students():
    list_of_all_branches=["btech","mtech","msc","mba","msc_tech","mca"]
    for each_degree in list_of_all_branches:
        course(each_degree)
    print("______________________________________________________________________")
    print("----------------------------------------------------------------------")
    print("-------  Degree -------------|-------- No of students ----------------")
    print(" ")
    print("       BTech                 :          "+str(total_no_of_students[0]))
    print("       MTech                 :          "+str(total_no_of_students[1]))
    print("       Msc                   :          "+str(total_no_of_students[2]))
    print("       Mba                   :          "+str(total_no_of_students[3]))
    print("       Msc-Tech              :          "+str(total_no_of_students[4]))
    print("       Mca                   :          "+str(total_no_of_students[5]))
    print("______________________________________________________________")
    print("   Total number of students   = "+str(sum(total_no_of_students)))
    for i in range(0,6):
        total_no_of_students.pop()
    print(" ")
    print("  ")


def search_by_room_number():
    find1= str(input("Enter the room number\n"))
    print(" ")
    for sheets in range(0,int(sheet_no)-2):
        worksheet1= wb.sheet_by_index(sheets)
        for rows in  range(0,worksheet1.nrows-1):
            if re.search(find1,worksheet1.cell_value(rows, 2)):
                list1=[]
                for details in range(0,worksheet1.ncols-2):
                    list1.append(worksheet1.cell_value(rows, details))
                days = number_of_days(str(worksheet1.cell_value(rows, details+1)))
                list1.append(days)
                display_student_details(list1)
                print("  ")
                print("  ")

def display_student_details(list1):
    print("------------------------------------------------------")
    print(" ")
    print("Room number            : "+list1[2])
    print("Room type              : "+list1[3])
    print("Name                   : "+list1[5])
    print("Roll Number            : "+str(list1[6]))
    print("Course                 : "+list1[7])
    print("Specialization         : "+list1[8])
    print("Department             : "+list1[9])
    print("Student mobile number  : "+str((list1[10])))
    print("Student Adhar Number   : "+str(list1[11]))
    print("Student Mail Id        : "+str(list1[12]))
    print("Father Mobile Numer    : "+str(list1[13]))
    print("Mother Mobile Number   : "+str((list1[14])))
    print("Mess Card Number       : "+str((list1[15])))
    print("I Collect              : "+str(list1[16]))
    print("Amount                 : "+str((list1[17])))
    print("Remarks                : "+str(list1[18]))
    print("Number_of days stayed  : "+str(list1[19]))
    print(" ")
    print("------------------------------------------------------")


print("__________________________________________________________")
print("----------------------------------------------------------")
print(" ")
print(" ")
while(1):
    print("Welcome to Hostel Room Allotment")
    print(" ")
    print("1. Single Room Vacancies")
    print("2. Double Room Vacancies")
    print("3. Filled Rooms")
    print("4. Search ")
    print("5. Details of each floor")
    print("6. Degree_pursuing")
    print("7. Total number of students")
    print("8. Search by room number")
    print("9. Exit")
    try:
        while True:
            option = input("Enter your choice\n")
            if option.isdigit():
                if int(option) >= 1 and int(option) <= 9:
                    break
                else:
                    print("enter valid choice")
    except NameError as e:
        print(e)
    if int(option) == 1:
        single_room_vacancies()
    elif int(option) == 2:
        double_room_vacancies()
    elif int(option) == 3:
        filled_rooms()
    elif int(option) == 4:
        search_detail()
    elif int(option) == 5:
        display_all_rooms()
    elif int(option) ==6:
        degree_pursuing()
    elif int(option) ==7:
        total_number_of_students()
    elif int(option) ==8:
        search_by_room_number()
    elif int(option) == 9:
        sys.exit(3)
    else:
        print("enter valid number")
        print(" ")
        print(" ")














