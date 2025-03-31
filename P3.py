#Import Libraries
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font

#Import the poorly organized data
dataWorkbook = openpyxl.load_workbook('Poorly_Organized_Data_1.xlsx')
sheet = dataWorkbook.active

#Create a new workbook
myWorkbook = Workbook()
myWorkbook.remove(myWorkbook.active)

#Gathers student info
class Student:
    def __init__(self, fname, lname, id, grade):
        self.fname = fname
        self.lname = lname
        self.id = id
        self.grade = grade

#Creates objects according to the subject and stores a list of students taking said subject
class Class:
    def __init__(self, className):
        self.className = className
        self.studentList = []

for row in sheet.iter_rows(min_row=2):
    if row[0].value not in myWorkbook.sheetnames:
        myWorkbook.create_sheet(row[0].value)

for sheet_name in myWorkbook.sheetnames: 
    currSheet = myWorkbook[sheet_name]  
    currSheet["A1"] = "Last Name"
    currSheet["B1"] = "First Name"
    currSheet["C1"] = "Student ID"
    currSheet["D1"] = "Grade"

for row in sheet.iter_rows(min_row=2, values_only= True):
    subject = row[0]
    myWorkbook.active = myWorkbook[subject]
    currSheet = myWorkbook.active
    values = row[1].split('_')
    grade = row[2]
    values.append(grade)
    currSheet.append(values)

for sheet_name in myWorkbook.sheetnames:
    last_row = myWorkbook.active.max_row
    myWorkbook.active.auto_filter.ref = f"A1:D1{last_row}"

myWorkbook.save(filename="P3.xlsx")

myWorkbook.close()