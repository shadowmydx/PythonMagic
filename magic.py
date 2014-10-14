# -*- coding:utf-8 -*-

from xlrd import open_workbook,cellname
import xlwt

wd = open_workbook(r'C:\Users\user\Desktop\DevelopMent\python\gongwuyuan\2015.xls')
professional   = (u'专业',u'法律',u'法学')
workExperience = (u'基层工作最低年限',u'无限制')
degree         = (u'学历',u'仅限本科')
mainCol        = [] # save what cols are p & w in.

tablehead   = [] # save all cols's names.
resultTable = [] # suitable res are in this table.

writeFile = xlwt.Workbook() # 

def getTableHead(sheet):
	for colIndex in range(sheet.ncols):
		tmp = sheet.cell(0,colIndex).value
		if tmp == professional[0]:
			mainCol.append(colIndex)
		if tmp == workExperience[0]:
			mainCol.append(colIndex)
		if tmp == degree[0]:
			mainCol.append(colIndex)
		tablehead.append(tmp)

def findSuitRow(sheet):
	startRow = 1
	for rowIndex in range(startRow,sheet.nrows):
		tar1 = sheet.cell(rowIndex,mainCol[0]).value
		tar2 = sheet.cell(rowIndex,mainCol[1]).value
		tar3 = sheet.cell(rowIndex,mainCol[2]).value
		if (tar1.find(professional[1]) != -1 or tar1.find(professional[2]) != -1) and tar3 == workExperience[1] and tar2.find(degree[1]) == -1:
			resultTable.append(rowIndex)
		
def printResult(sheet):
	tmpTable = writeFile.add_sheet(sheet.name)
	for inx,key in enumerate(tablehead):
		tmpTable.write(0,inx,key)
	startRow = 1
	for key in resultTable:
		for colIndex in range(sheet.ncols):
			tmpTable.write(startRow,colIndex,sheet.cell(key,colIndex).value)
		startRow += 1

		
def main():
	global mainCol
	global resultTable
	global tablehead
	for sheet in wd.sheets():
		getTableHead(sheet)
		findSuitRow(sheet)
		printResult(sheet)
		mainCol = []
		resultTable = []
		tablehead = []
		
	writeFile.save(r'C:\Users\user\Desktop\DevelopMent\python\gongwuyuan\output.xls')

main()