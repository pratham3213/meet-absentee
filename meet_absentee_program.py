from os import *
from tkinter import *
import datetime
import openpyxl as xl
import pandas as pd
from tkinter import filedialog


path=getcwd()

def tkinter_window():
	root=Tk()
	root.title("Choose Your CSV File")
	root.filename=filedialog.askopenfilename(initialdir=path, title="Select A File")
	return root.filename


def csv():
	data = pd.read_csv(tkinter_window())

	names=data["Participants"].tolist()
	return names


def excel():

	wb = xl.load_workbook("12_B.xlsx")

	sheet = wb["Sheet1"]

	Present_Student=csv()

	for row in range(2, sheet.max_row+1):
		cell1 = sheet[f"B{row}"]
		date = sheet[f"C{1}"]
		date.value=datetime.date.today()
		cell2=sheet[f"C{row}"]
		
		for present in Present_Student:
			if present==cell1.value:
				cell2.value="Present"
				
	

	wb.save("11_B.xlsx")

excel()
