import xlrd 
import xlwt 
import xlsxwriter
import os 
import re 

def casesDV(): 

#winning % of cases alleging DV
def winning():	
#
def defendantAgree():

def defendantDisagree(): 

def petition1st():

def petition2nd(): 

def legalPlantiff(): 

def hospitalVisit(): 

def deniedDV(): 


if __name__ == '__main__':

	beforeCom = "../2014哈尔滨.xls"
	target1 = xlrd.open_workbook(beforeCom)
	open_sheet = target1.sheets()[0]
	
	final = xlwt.Workbook()
	final_sheet = final.add_sheet(u'processed', cell_overwrite_ok=True)
	arrange = xlwt.Workbook()
	arrange_sheet = arrange.add_sheet(u'open', cell_overwrite_ok=True)
	
	jufa = xlrd.open_workbook("../../Downloads/2015哈尔滨.xlsx")
	jufa_sheet = jufa.sheets()[0] 

	final.save("../2015哈尔滨comp.xls")	
	arrange.save("../arrangeOpen.xls")