# coding=utf8
import xlrd 
import xlwt 
import xlsxwriter
import os 
import re 
import requests 
'''
def callAPI(): 
	#需要解析的文本，是一个字符串，直接用你请求裁判文书网返回的值即可
	str_data="[http://wenshu.court.gov.cn/website/wenshu/181217BMTKHNT2W0/index.html?pageId=c0e3d8f04290ebe20efddfa2513fdda2&s21=%E5%AE%B6%E5%BA%AD%E6%9A%B4%E5%8A%9B&s16=%E6%B0%91%E4%BA%8B%E6%A1%88%E7%94%B1&s11=9000&s1=%E7%A6%BB%E5%A9%9A%E7%BA%A0%E7%BA%B7&s4=4&s8=03&s6=01&cprqStart=2015-01-01&cprqEnd=2015-12-31]"
	return_str=requests.post('http://www.ulaw.top:5677/crack',data={'a':str_data},timeout=20).text
	print(return_str)
'''
def casesDV(sheets): 
	pass 

#winning % of cases alleging DV 
def winning(sheets):	
	#the column is 25
	num_y = 0
	num_n = 0
	total = 0
	for out in sheets: 
		for sheet in out: 
			print(sheet.name)
			ncols = sheet.ncols
			nrows = sheet.nrows

			win = sheet.col_values(25, start_rowx=1, end_rowx=nrows)
			count = 1 
			for i in win: 
				if i == "y": 
					num_y = num_y + 1
					total = total + 1 
				elif i == "n": 
					num_n = num_n + 1
					total = total + 1
				elif i == "": 
					continue
				else: 
					print(count)
					print(i)
				count = count + 1 
	percent_y = round((num_y / total)*100,2)
	percent_n = round((num_n / total)*100,2)
	print("1. Winning", percent_y, "% of cases alleging DV ")
	return total 

def defendantOpin(sheets):
	#the column is 24 
	yy = 0 
	yn = 0 
	yna = 0
	ycy = 0
	totaly = 0
	totaln = 0

	totalyop=0
	totalnop=0 
	totalnaop=0

	total1 = 0 
	total2 = 0
	for out in sheets:
		for sheet in out: 
			ncols = sheet.ncols
			nrows = sheet.nrows

			opinion = sheet.col_values(24, start_rowx=1, end_rowx=nrows)
			rule = sheet.col_values(25, start_rowx=1, end_rowx=nrows)

			for i in range(0, len(opinion)): 
				
				if opinion[i] == "y": 
					totalyop = totalyop + 1
					if rule[i] == "y": 
						yy = yy + 1
				elif opinion[i] == "n": 
					totalnop = totalnop + 1
					if rule[i] == "y": 
						yn = yn + 1 
				elif opinion[i] == "n/a" or opinion[i] == "nm" or opinion[i] == "not mentioned" or opinion[i] == "not mentinoed": 
					totalnaop = totalnaop + 1 
					if rule[i] == "y":
						yna = yna + 1
				elif opinion[i] == "cy" or opinion[i][:2] == "cy":
					totalyop = totalyop + 1
					if rule[i] == "y": 
						ycy = ycy + 1 
				elif opinion[i] == "": 
					continue  
				else: 
					print(opinion[i])
				total1 = total1 + 1

	percent_yy = round(((yy+ycy)/totalyop)*100,2)
	percent_yn = round((yn/totalnop)*100,2)
	percent_yna = round((yna/totalnaop)*100,2) 
	#percent_ycy = (ycy/total)*100 
	print("2.",percent_yy, "% of cases alleging DV in which defendant agreed to divorce that got approved")
	print("3.",percent_yn, "% of cases alleging DV in which defendant disagreed to divorce that got approved")
	print("4.",percent_yna, "% of cases alleging DV in which defendant neither agreeded or disagreed to divorce that got approved")

	return total1 

def petitionInflu(sheets):
	totalfirst = 0 
	totalsecond = 0 
	rulefirst = 0
	rulesecond = 0 
	for out in sheets:
		for sheet in out: 
			ncols = sheet.ncols
			nrows = sheet.nrows

			first = sheet.col_values(17, start_rowx=1, end_rowx=nrows)
			second = sheet.col_values(18, start_rowx=1, end_rowx=nrows)
			rule = sheet.col_values(25, start_rowx=1, end_rowx=nrows)
			if len(first) != len(second): 
				print("DANGER!")
			if len(first) != len(rule): 
				print("DANGER!")
			for i in range(0, len(first)): 
				
				if first[i] == "y": 
					totalfirst = totalfirst + 1
					if rule[i] == "y": 
						rulefirst = rulefirst + 1 
				elif first[i] == "": 
					if second[i] != "": 
						print("something went wrong: ", i)

				if second[i] == "y": 
					totalsecond = totalsecond + 1 
					if rule[i] == "y": 
						rulesecond = rulesecond + 1
				
	percent_first = round((rulefirst/totalfirst)*100,2)
	percent_second = round((rulesecond/totalsecond)*100,2)
	print("5.",percent_first,"% of cases alleging DV in 1st petition that got approved")
	print("6.",percent_second,"% of cases alleging DV in 2nd petition or later that got approved")
	total = totalfirst+totalsecond 
	return total 

def legalPlantiff(sheets): 
	y = 0 
	n = 0
	total = 0
	for out in sheets: 
		for sheet in out: 
			ncols = sheet.ncols
			nrows = sheet.nrows

			legal_p = sheet.col_values(10, start_rowx=1, end_rowx=nrows)
			for i in legal_p: 
				if i == "y": 
					y = y + 1
				elif i == "n": 
					n = n + 1
				elif i == "": 
					continue 
				total = total + 1 

	percent_y = round((y/total)*100,2)
	print("7.",percent_y,"% of cases alleging DV in which plaintiff was legally represented")
	return total 

def hospitalVisit(sheets): 
	#evidence @31 
	#plantiff reasons @22 
	total_hos = 0 
	total = 0
	for out in sheets: 
		for sheet in out:
			ncols = sheet.ncols
			nrows = sheet.nrows

			evidence = sheet.col_values(31, start_rowx=1, end_rowx=nrows)
			reasons = sheet.col_values(26, start_rowx=1, end_rowx=nrows)
			for i in range(0, len(evidence)): 
				num0 = evidence[i].find("医院")
				num1 = evidence[i].find("病历")
				num2 = evidence[i].find("诊断")
				if num0!=-1 or num1 != -1 or num2 !=-1: 
					total_hos = total_hos + 1 
				else: 
					num3 = reasons[i].find("医院")
					num4 = reasons[i].find("病历")
					num5 = reasons[i].find("诊断")
					if num3 != -1 or num4 != -1 or num5 !=-1: 
						total_hos = total_hos + 1 
				if evidence[i] != "": 
					total = total + 1 
	percent_hos = round((total_hos/total)*100,2) 
	print("8.",percent_hos,"% of cases alleging DV in which plaintiff had at least one hospital visit")
	return total 

def deniedDV(sheets): 
	total_y = 0 
	total_na = 0 
	total = 0
	for out in sheets: 
		for sheet in out:
			ncols = sheet.ncols
			nrows = sheet.nrows

			denial = sheet.col_values(30, start_rowx=1, end_rowx=nrows)
			count = 0
			for d in denial:
				count = count + 1 
				if d == "": 
					continue 
				else: 
					total = total + 1 
					if d == "y" or d[:1] == "y": 
						total_y = total_y + 1
					elif d == "n/a" or d == "nm" or d.find("not") != -1:  
						total_na = total_na + 1
					elif d == "n" or d[:1] == "n": 
						continue 
					else:
						print(sheet.name)
						print(count) 
						print(d)

	percent_deny = round((total_y/total)*100,2)
	print("9.",percent_deny,"% of cases alleging DV in which defendant denied DV (yes or no)")
	return total 

if __name__ == '__main__':
	filename1 = "1.1 广州市all case.xlsx"
	filename2 = "3.2 重庆市case#0694-1116.xlsx"
	filename3 = "3.3 重庆市case#1117-1474.xlsx"
	str_data1 = "../"+filename1
	str_data2 = "../"+filename2
	str_data3 = "../"+filename3

	target1 = xlrd.open_workbook(str_data1)
	target2 = xlrd.open_workbook(str_data2)
	target3 = xlrd.open_workbook(str_data3)

	sheets1 = target1.sheets()
	sheets2 = target2.sheets()
	sheets3 = target3.sheets()

	sheets = [sheets1]

	num1 = winning(sheets)
	#num2 = defendantOpin(sheets)
	num3 = petitionInflu(sheets)
	num4 = legalPlantiff(sheets)
	num5 = hospitalVisit(sheets)
	num6 = deniedDV(sheets)
	print(num1,num3,num4,num5,num6)
'''
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
	arrange.save("../arrangeOpen.xls")'''