#!/usr/bin/python

'''
internal_use_pc_data_process.py - Automatically seperate the Excel file into difference sheets based on the Pcid
	usage: py internal_use_pc_data_process.py  (make sure internal_use_pc.xls is on the same foler)
	
	Date created : 2/10/2019
 	Date last modified : 2/19/2019

 	Author : Yu Li
	
Depends on: xlrd
			xlwt
'''

import xlrd
import xlwt
import datetime

new_workbook = xlwt.Workbook()


def open_file(path):
	workbook = xlrd.open_workbook(path)
	sheet = workbook.sheet_by_index(0)
	counter = 0
	laptop_counter = 0
	printer_counter = 0
	server_counter = 0
	desktop_counter = 0
	rowNum = 0
	
	new_sheet1 = new_workbook.add_sheet('Desktops')
	new_sheet2 = new_workbook.add_sheet('Laptops')
	new_sheet3 = new_workbook.add_sheet('Printers')
	new_sheet4 = new_workbook.add_sheet('Servers')
	
	# Get the number of row in the Excel file
	for row in range(sheet.nrows):
		rowNum = row
		
	while (counter < rowNum):
		data = [sheet.cell_value(counter, col) for col in range(sheet.ncols)]
		print(data)
		
		first_cell = sheet.cell_value(counter,0)
		if (first_cell[0])== 'L':
			for index, value in enumerate(data):
				new_sheet2.write(laptop_counter, index, value)
			laptop_counter += 1
			
		elif (first_cell[0])== 'P':
			for index, value in enumerate(data):
				new_sheet3.write(printer_counter, index, value)
			printer_counter += 1
		
		elif (first_cell[0])== 'S':
			for index, value in enumerate(data):
				new_sheet4.write(server_counter, index, value)
			server_counter += 1
			
		else:
			for index, value in enumerate(data):
				new_sheet1.write(desktop_counter, index, value)
			desktop_counter += 1
			
		counter += 1	
		
def save_file():
	date_object = datetime.date.today()
	new_workbook.save("Asset List " + str(date_object) + ".xls")

if __name__ == "__main__":
	file = "internal_use_pc.xls"
	open_file(file)
	save_file()