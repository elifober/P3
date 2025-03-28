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

