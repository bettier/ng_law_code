import xlrd 
import xlwt 
import xlsxwriter
import os 
import re #regularization

#read the sheet from openLaws 
def read_(target, write_sheet):
	sheets = target.sheets()

	#iterate all 
	for sheet in sheets: 
		print(sheet.name)
		ncols = sheet.ncols
		nrows = sheet.nrows
		print("the rows: ", nrows)
		
		for i in range(0,ncols):
			#the title for the column to process
			title = sheet.cell(0,i).value
			if title == '法院':#@2
				district = sheet.col_values(i, start_rowx=1, end_rowx=nrows)
				write_col(district, 2, write_sheet) 
			elif title == '标题': #@4 
				case_names = sheet.col_values(i, start_rowx=1, end_rowx=nrows)
				write_sheet.write(0, 4, "case name")
				write_col(case_names, 4, write_sheet)
			elif title == '案号': #@3
				case_nums = sheet.col_values(i, start_rowx=1, end_rowx=nrows)
				write_sheet.write(0, 3,"case num")
				write_col(case_nums, 3, write_sheet)
			elif title == '当事人': #@06-11 
				dangshiren = sheet.col_values(i, start_rowx=1, end_rowx=nrows)
				process_danshiren(dangshiren, write_sheet)
			elif title == '庭审程序说明': #@12-13
				court = sheet.col_values(i, start_rowx=1, end_rowx=nrows)
				court3 = sheet.col_values(21, start_rowx=1, end_rowx=nrows)
				process_court(court, court3, write_sheet)
			elif title == '庭审过程': 
				#print("THIS IS the======", i)
				court2 = sheet.col_values(i, start_rowx=1, end_rowx=nrows)
				process_court2(court2, write_sheet)

def write_same(stuff, colN, rowN, write_sheet): 
	for i in range(1, rowN): 
		write_sheet.write(i, colN, stuff)

def write_col(col, colN, write_sheet):
	for i in range(0, len(col)):
		write_sheet.write(i+1, colN, col[i])

def write_col_highlight(col, highlight, color, colN, write_sheet):
	#pattern highlight 
	pattern1 = xlwt.Pattern() 
	pattern1.pattern = xlwt.Pattern.SOLID_PATTERN
	pattern1.pattern_fore_colour = color #this is yello 
	style1 = xlwt.XFStyle()
	style1.pattern = pattern1 

	for i in range(0, len(col)):
		if highlight[i]: 
			write_sheet.write(i+1, colN, col[i], style1)
		else: 
			write_sheet.write(i+1, colN, col[i])

#process the dangshi ren informatin 
#petition gender 
def process_danshiren(dangshiren, write_sheet):
	yuanGao_gender = []
	beiGao_gender = []
	yuanGao_DOB = []
	beiGao_DOB = []
	yuanGao_legal = []
	beiGao_legal = []
	for i in range(0, len(dangshiren)): 
		str_ = dangshiren[i]
		str_ = str_.replace('、', '')

		#check if it is empty 
		if str_ == "":
			yuanGao_gender.append("n/a")
			beiGao_gender.append("n/a")
			yuanGao_DOB.append("n/a")
			beiGao_DOB.append("n/a")
			yuanGao_legal.append("n/a")
			beiGao_legal.append("n/a")
			continue  

		#found the Beigao 
		two = str_.split('被告')
		if len(two) < 2: 	
			print("DID NOT FOUND 被告 in DANGSHIREN!!!!")
			#print(str_)
			continue 
		yuanGao = two[0]
		beiGao = '被告' + two[1]
		
		#find the gender of the Yuangao and Beigao
		match = re.search("女|男", yuanGao)
		if match: 
			y_gender = yuanGao[match.start():match.end()]
			if y_gender == '女': 
				yuanGao_gender.append('female')
				beiGao_gender.append('male')
			elif y_gender == '男':
				yuanGao_gender.append('male')
				beiGao_gender.append('female')
			else: 
				yuanGao_gender.append('n/a')
				beiGao_gender.append('n/a')
		else: 
			yuanGao_gender.append('n/a')
			beiGao_gender.append('n/a')

		#find the DOB of the Yuangao 
		#1. 找到原告+字符串+逗号+性别+逗号+年份
		#2. 找到前两个逗号，然后提取年份
		#concern: you have to avoid to use the info from weituo representative
		pattern_y = r'原告[\u4e00-\u9fa5]*(\d)*.(女|男).(汉族，)*(生于)*[\d]{4}年'
		match_y = re.search(pattern_y, yuanGao)
		if match_y:
			str_y = yuanGao[match_y.start():match_y.end()]
			dob_y = r'[\d]{4}'
			yuanGao_year = re.search(dob_y, str_y)
			if yuanGao_year: 
				yuanGao_year = str_y[yuanGao_year.start():yuanGao_year.end()]
				yuanGao_DOB.append(int(yuanGao_year))
			else: 
				#print(yuanGao)
				yuanGao_DOB.append("n/a")
		else: 
			#print(yuanGao)
			yuanGao_DOB.append("n/a")

		#find the DOB of the 被告
		pattern_b = r'被告[\u4e00-\u9fa5]*.(女|男).(汉族，)*(生于)*[\d]{4}年'
		match_b = re.search(pattern_b, beiGao)
		if match_b:
			str_b = beiGao[match_b.start():match_b.end()]
			dob_b = r'[\d]{4}'
			beiGao_year = re.search(dob_b, str_b)
			if beiGao_year: 
				beiGao_year = str_b[beiGao_year.start():beiGao_year.end()]
				beiGao_DOB.append(int(beiGao_year))
			else: 
				beiGao_DOB.append("n/a")
		else: 
			beiGao_DOB.append("n/a")


		#decide whether the they have legal representatives 
		yuan_rep = re.search('法律|律师', yuanGao)
		bei_rep = re.search('法律|律师', beiGao)
		if yuan_rep: 
			yuanGao_legal.append('y')
		else: 
			yuanGao_legal.append('n')
		if bei_rep: 
			beiGao_legal.append('y')
		else: 
			beiGao_legal.append('n')

	#out of the for loop 
	write_col(yuanGao_gender, 6, write_sheet)
	write_col(yuanGao_DOB, 7, write_sheet)
	write_col(beiGao_gender, 8, write_sheet)
	write_col(beiGao_DOB, 9, write_sheet)
	write_sheet.write(0,10,"yuanGao_legal")
	write_sheet.write(0,11,"beiGao_legal")
	write_col(yuanGao_legal, 10, write_sheet)
	write_col(beiGao_legal, 11, write_sheet)

#summary trial, @12 
#public or not public @13
def process_court(court2, court3, rite_sheet): 
	summary_col = []
	public_col = []
	for i in range(0, len(court2)):
		court = court2[i]+court3[i]
		#summary trial or not 
		summary = re.search('简易程序|独任', court)
		if summary: 
			summary_col.append('y')
		else: 
			not_sum = re.search('合议庭|和议庭', court)
			if not_sum: 
				summary_col.append('n')
			else: 
				#num = i + 2
				#print("About the summary trial or not")
				#print(str(num)+": " + court)
				summary_col.append('n/a')

		#public or not 
		no_public = re.search('不公开', court)
		if no_public: 
			public_col.append('n')
		else: 
			public_ = re.search('公开', court)
			if public_: 
				public_col.append('y')
			else: 
				public_col.append('n/a')

	write_sheet.write(0,12, "summary trial")
	write_col(summary_col, 12, write_sheet)
	write_sheet.write(0,13,"publicOrNot")
	write_col(public_col, 13, write_sheet)	

#process the 审理 to decide the year and remarry part 
def prove_year_remarry(prove):
	year_return = 0 
	remarried_return = None 
	highlight = None 

	marriage = re.search(r'结婚|登记', prove)
	if marriage: 
		the_ = re.split("，", prove)
		for index, t in enumerate(the_): 
			year_ = re.search(r'结婚|登记', t) 
			if year_:
				mar_year = re.search(r'[0-9][0-9][0-9][0-9]', t)
				if mar_year: 
					year_return = t[mar_year.start():mar_year.end()]
					break 
				else:  
					#同年结婚
					tong = re.search(r'同年.*结婚', t)
					if tong: 
						'''f0 = index-1 
						f1 = True
						y_ = 0
						while(f1): 
							if f0 >= 0 and f0 < len(the_): 
								#b = re.search(r'[0-9][0-9][0-9][0-9]年(.*(相识|同居|恋爱|建立关系|认识))*', the_[f0])
								#if b: 
								the_y = re.findall(r'[0-9][0-9][0-9][0-9]', the_[f0])
								if the_y: 
									y_ = the_y[-1]
									y_ = int(y_)
									f1 = False 
								else:
									f0 = f0 -1 
							else: 
								f1 = False  
						year_return = y_'''
					 
						t_s = index 
						while(t_s >= 1): 
							new_t=the_[t_s-1]
							year_2 = re.findall(r'[0-9][0-9][0-9][0-9]', new_t) 
							if len(year_2) !=0 : 
								year_return = int(year_2[-1])
								break 
							else: 		
								year_return = 0
							t_s = t_s -1 
			if year_return != 0: 
				break 
	else:
		year_return = 0

	#decide whether it is remarried or not 
	remarried = re.search(r'再婚', prove)
	if remarried:
		remarried_return = 'y'
	else: 
		remarried_return = 'n'

	return year_return, remarried_return

def prove_petition(prove):
	
	sum_petition = 0
	str_petition = "("
	str_second = ""

	sentences = re.split('。', prove)
	for s in sentences: 
		#find the first petition, #fine the second petition 
		petition = re.findall(r'撤诉|撤回', s)
		for p in petition: #if found 
			years = re.findall(r'[0-9][0-9][0-9][0-9]年', s)
			if len(years) != 0:
				str_petition = str_petition+str(years[-1][0:4])+","
			str_petition = str_petition + "withdrawl."
			sum_petition = sum_petition + 1 

		petition2 = re.findall(r'不准离婚|驳回', s)
		for p in petition2: #if found 
			years = re.findall(r'[0-9][0-9][0-9][0-9]年', s)
			if len(years) != 0:
				str_petition = str_petition+str(years[-1][0:4])+","
			str_petition = str_petition + "no divorce."
			sum_petition = sum_petition + 1 

	str_petition = str_petition + ")"
	if sum_petition == 0: 
		str_petition = "y"
		str_second = "n"
	else: 
		str_petition = "n"+str_petition
		str_second = "y"

	return str_petition, str_second 

def prove_children(prove):
	sum_female = 0 
	sum_male = 0
	year_list = []
	str_year = ""
	str_gender = ""	 
	
	sentences = re.split('。', prove)
	for s in sentences: 
		comma = re.split("，",s)
		
		for c in comma: 
			#find the number of children 
			girls = re.findall(r'一女|生育长女|次女|生育女', c)
			boys = re.findall(r'一子|生育长子|次子|生育([\u4e00-\u9fa5]*)儿|生育男', c)
			sum_female = sum_female+len(girls)
			sum_male = sum_male+len(boys)
		
			if len(girls) != 0 or len(boys) != 0:
				years = re.findall(r'[0-9][0-9][0-9][0-9]年', c)
				for y in years: 
					year_list.append(int(str(y[0:4])))
	#process the chlidren problems 
	year_list.sort()
	for y in year_list:
		str_year = str(str_year) + str(y) + ","
	if len(year_list) == 0:
		str_year = "n/a"
	else: 
		str_year = str_year[:len(str_year)-1]
		str_year = str_year+"."

	if sum_female != 0: 
		str_gender = str(sum_female)+" female,"
	if sum_male != 0:
		str_gender = str_gender + str(sum_male)+" male."
	
	total = sum_female+sum_male
	if total == 0: 
		str_gender = "n/a"

	return total, str_year, str_gender

def yuanGao_reason(rest): 
	#find the yuan Gao 
	sep1 = re.search(r'被告([\u4e00-\u9fa5]*)未举示证据|辩称|被告([\u4e00-\u9fa5]*)辨称|被告([\u4e00-\u9fa5]*)(X)*([\u4e00-\u9fa5]*)辩称|被告([\u4e00-\u9fa5]*)未到|被告([\u4e00-\u9fa5]*)未出庭|被告辨称|辩称:|被告梁X全辩称|被告([\u4e00-\u9fa5]*)诉称|被告([\u4e00-\u9fa5]*)未应诉答辨|被告.*拒不到庭参加诉讼|未到庭|被告([\u4e00-\u9fa5]*)对原告主张|被告([\u4e00-\u9fa5]*)没有到庭|被告([\u4e00-\u9fa5]*)(×|\d)*辩称|被告([\u4e00-\u9fa5]*)答辩|被告([\u4e00-\u9fa5]*)未到庭|被告([\u4e00-\u9fa5]*)，*未到庭|被告([\u4e00-\u9fa5]*)，*未有答辩|被告([\u4e00-\u9fa5]*)辩称|被告([\u4e00-\u9fa5]*)，*无书面|被告([\u4e00-\u9fa5]*)，*未提交书面|被告[\u4e00-\u9fa5]*，[\u4e00-\u9fa5]*未向本院[\u4e00-\u9fa5]*|([\u4e00-\u9fa5]*)未向本院([\u4e00-\u9fa5]*)', rest)
	yuan = ""
	bei = ""
	if sep1:
		yuan = rest[:sep1.start()]
		bei = rest[sep1.start():]
		#delete evidence part from the yuangao 
		eviYuan = re.search("原告为证实其主张|原告为证明|向法庭提交|在本院开庭", yuan)
		if eviYuan: 
			yuan = yuan[:eviYuan.start()]
			bei = rest[eviYuan.end():]

		#process the beigao 
		dele = re.search(r'本院[\u4e00-\u9fa5]*，认证如下|原告为证实其主张|原告为证明|向法庭提交|在本院开庭', bei)
		if dele:
			bei = bei[:dele.start()]
	else: 
		print("DID NOT FOUND 被告")
		print(rest)
	return yuan, bei 

#decude whether it is disputed or not 
def defendant_dispute_op(bei, ab): 
	#if bei is empty or absent 
	return_ = ["",""]
	if bei=="" or ab: 
		return "n/a","n/a"
	dispute = re.search(r'没[\u4e00-\u9fa5]*家[\u4e00-\u9fa5]*暴|无[\u4e00-\u9fa5]*家[\u4e00-\u9fa5]*暴|家[\u4e00-\u9fa5]*暴[\u4e00-\u9fa5]*不是', bei)
	if dispute: 
		return_[0] = "y"
	else: 
		return_[0] = "n"

	opt = re.search(r'不同意([\u4e00-\u9fa5]*)离婚', bei)
	if opt: 
		return_[1] = 'n'
	elif re.search(r'同意([\u4e00-\u9fa5]*)离婚', bei):
		return_[1] = 'y'
	else: 
		return_[1] = 'n/a'

	return return_[0],return_[1]

def process_court2(court2, write_sheet): 
	both_col = []
	
	yuan_l = []
	bei_l =[]
	dispute_l = []
	opt_l = []

	year_l = [] 
	remarry_l = []
	str_petition_l = [] 
	str_second_l = []
	total_l = [] 
	str_year_l = []
	str_gender_l = [] 

	prove_l = []

	pattern1 = xlwt.Pattern() 
	pattern1.pattern = xlwt.Pattern.SOLID_PATTERN
	pattern1.pattern_fore_colour = 5 #this is yello 
	style1 = xlwt.XFStyle()
	style1.pattern = pattern1 

	for i in range(0, len(court2)): 
		num = i + 1 
		#if the line is empty 
		line = court2[i]
		if line == "": 
			both_col.append("")
			yuan_l.append("")
			bei_l.append("")
			dispute_l.append("")
			opt_l.append("")
			year_l.append("")
			remarry_l.append("")
			str_petition_l.append("")
			str_second_l.append("")
			total_l.append("")
			str_year_l.append("")
			str_gender_l.append("")
			continue 
		#remove the 。、and replace it with 。
		line = line.replace('。、', '。')
		line = line.replace('。；', '。')
		#find if the defendant is present 
		absence = re.search("依法缺席|未到庭|缺席|拒不到庭", line)
		if absence: 
			both_col.append("n(miss plantiff)")
		else: 
			both_col.append('y')

		#plantiff reason 
		sets = line.split("诉称", 1)
		second = None 
		if len(sets) >= 2: 
			second = sets[1]
		else: 
			sets2 = re.split("原告[\u4e00-\u9fa5]*称|原告[\u4e00-\u9fa5]*诉请", line)
			if len(sets2) >= 2: 
				second = sets2[1]
			else: 
				print("NO YUANGAO: "+line)
				both_col.append("")
				yuan_l.append("")
				bei_l.append("")
				dispute_l.append("")
				opt_l.append("")
				year_l.append("")
				remarry_l.append("")
				str_petition_l.append("")
				str_second_l.append("")
				total_l.append("")
				str_year_l.append("")
				str_gender_l.append("")
				continue 
		second = second[1:] 
		
		#there might be the case that the above are in front of the 被告
		prove = re.search(r'现查明以下事实|将法律事实确认如下|本案诉讼中，|确认如下事实|本院[\u4e00-\u9fa5]*事实|原告罗某某与被告谭某甲|本院[\u4e00-\u9fa5]*确认如下[\u4e00-\u9fa5]*事实|本院确认以下|本院认定以下|本案确认以下|本案确认|本案如下法律事实|本院确认事实如下|经审核认定|经庭审审核|审理查明|经审理查明|经庭审查明|本院确认如下事实|对本案事实..如下|根据当事人举证质证，对本案事实认定如下|确认以下事实', second)
		if prove: 
			#remove the last part if prove exists  
			rest = second[:prove.start()]
		else: 
			rest = second 

		yuan, bei = yuanGao_reason(rest)
		dispute, opt = defendant_dispute_op(bei, absence)

		yuan_l.append(yuan)
		bei_l.append(bei)
		dispute_l.append(dispute)
		opt_l.append(opt)

		#if there is prove, separate the evidences from itself 
		year=0
		remarry="n"
		str_petition="y" 
		str_second="n" 
		total=0 
		str_year="" 
		str_gender="" 
		highlight=False

		if prove:
			prove_ = second[prove.start():]
			found = None 
			found = re.search(r"家庭暴力|家暴|纠纷|打骂|争执|打闹|致伤|损伤|吵闹|矛盾|报警", prove_)
			evi = re.search("上述事实", prove_)
			if found: 
				if evi: 
					write_sheet.write(num,33, prove_[:evi.start()], style1)
				else:
					write_sheet.write(num,33, prove_, style1)
			else:
				if evi: 
					write_sheet.write(num,33, prove_[:evi.start()])
				else: 
					write_sheet.write(num,33, prove_)

			#evidences @31上述事实
			if evi: 
				write_sheet.write(num,31, prove_[evi.start():])	
			else: 
				sss = re.search(r'下列证据|原告([\u4e00-\u9fa5])*为证明自己的主张|原告([\u4e00-\u9fa5])*(为)*证明', yuan)
				if sss: 
					write_sheet.write(num, 31, yuan[sss.start():])
				else: 
					write_sheet.write(num,31,"n/a")

			year, remarry = prove_year_remarry(prove_)
			str_petition, str_second = prove_petition(prove_) 
			total,str_year, str_gender = prove_children(prove_)
			prove_l.append(prove)
			
		#if there is no proof, find the information from the yuangao reason 
		else: 
			year, remarry = prove_year_remarry(prove_)
			str_petition, str_second = prove_petition(prove_) 
			total,str_year, str_gender = prove_children(prove_)
			prove_l.append("n/a")

		year_l.append(year)
		remarry_l.append(remarry)
		str_petition_l.append(str_petition)
		str_second_l.append(str_second)
		total_l.append(total)
		str_year_l.append(str_year)
		str_gender_l.append(str_gender)

	#outside of each loop 
	write_sheet.write(0, 14, 'both parties')
	write_col(both_col, 14, write_sheet)
	write_sheet.write(0,15, "married year")
	write_col(year_l, 15, write_sheet)
	write_sheet.write(0,16, "remarried?")
	write_col(remarry_l, 16, write_sheet)
	write_sheet.write(0,17,"first time?")
	write_col(str_petition_l, 17, write_sheet)	
	write_sheet.write(0,18,"second time?")
	write_col(str_second_l, 18, write_sheet)
	write_sheet.write(0,19,"amount of children")
	write_col(total_l, 19, write_sheet)
	write_sheet.write(0,20,"birth")
	write_col(str_year_l, 20, write_sheet)
	write_sheet.write(0,21,"gender")
	write_col(str_gender_l, 21, write_sheet)
	write_sheet.write(0,22, 'plantiff')
	write_col(yuan_l, 22, write_sheet)	
	write_sheet.write(0, 23, "defendant reasons")
	write_col(bei_l, 23, write_sheet)
	write_sheet.write(0,24, "defendant agrees")
	write_col(opt_l, 24, write_sheet)
	write_sheet.write(0,30, "disputed?")
	write_col(dispute_l, 30, write_sheet)

def write_row(row, rowN, sheet, style1, style2, style3): 
	for i in range(0, len(row)): 
		found = None 
		if i == 33: #prove 
			found = re.search(r"家庭暴力|家暴|纠纷|打骂|争执|打闹|致伤|损伤|吵闹|矛盾|报警", row[i].value)
			if found: 
				sheet.write(rowN, i, row[i].value, style1)
			else: 
				sheet.write(rowN, i, row[i].value) 
		else:
			sheet.write(rowN, i, row[i].value)

def write_col_gen(col, colN, style, sheet):
	for i in range(0, len(col)):		
		found = None 
		if colN == 25: #result 
			found = re.search(r"y", col[i])
			if found: 
				sheet.write(i+1, colN, col[i], style)
			else: 
				sheet.write(i+1, colN, col[i])
		elif colN == 26: #本院认为
			found = re.search(r"第三十二条第二款", col[i])
			if found: 
				sheet.write(i+1, colN, col[i], style)
			else: 
				sheet.write(i+1, colN, col[i])
		else: 
			sheet.write(i+1, colN, col[i])

def compareToJufa(start, the_year, jufa, open_, final, arrange, original):
	sheets = original.sheets()
	origin = None 
	for sheet in sheets: 
		origin = sheet 
		if origin: 
			break 
	#YELLO 
	pattern1 = xlwt.Pattern() 
	pattern1.pattern = xlwt.Pattern.SOLID_PATTERN
	pattern1.pattern_fore_colour = 5 
	style1 = xlwt.XFStyle()
	style1.pattern = pattern1 

	#RED 
	pattern2 = xlwt.Pattern() 
	pattern2.pattern = xlwt.Pattern.SOLID_PATTERN
	pattern2.pattern_fore_colour = xlwt.Style.colour_map['pink'] 
	style2 = xlwt.XFStyle()
	style2.pattern = pattern2

	#BLUE 
	pattern3 = xlwt.Pattern() 
	pattern3.pattern = xlwt.Pattern.SOLID_PATTERN
	pattern3.pattern_fore_colour = xlwt.Style.colour_map['pale_blue']  
	style3 = xlwt.XFStyle()
	style3.pattern = pattern3

	jufa_not = []
	open_not = []
	jufa_yes = []
	open_yes = []

	ncols_jufa = jufa.ncols
	nrows_jufa = jufa.nrows
	ncols_open = open_.ncols
	nrows_open = open_.nrows
	
	jufa_case = jufa.col(10, start_rowx=1, end_rowx=nrows_jufa)
	open_case = open_.col(3, start_rowx=1, end_rowx=nrows_open)
	#for each of the case number, find the corresponding one from the new laaws 
	name_row = open_.row(0)
	jufa_v = []
	open_v = []
	for i in jufa_case: 
		jufa_v.append(i.value)
	for i in open_case: 
		open_v.append(i.value)
	#write the title to the final sheet 
	write_row(name_row, 0, final, style1, style2, style3)

	for i in range(0, len(jufa_case)): 
		num=i+1 
		case = jufa_v[i]
		if case in open_v:
			index = open_v.index(case)
			row = open_.row(index+1) 
			write_row(row,num,final,style1,style2,style3)

			ori_row = origin.row(index+1)
			write_row(ori_row,num,arrange,style1,style2,style3)

		#not in the openLaws sheet 
		else: 
			#self add in the information 
			the_row = jufa.row(num)
			name = the_row[1].value 
			court = the_row[6].value 
			final.write(num,3, case)
			final.write(num,4, name)
			final.write(num,2, court)

			#经审理@34, write again 
			prove = jufa.row(num)[12].value
			found = None 
			found = re.search(r"家庭暴力|家暴|纠纷|打骂|争执|打闹|致伤|损伤|吵闹|矛盾|报警", prove)
			evi = re.search("上述事实", prove)
			if found: 
				if evi: 
					final.write(num,33, prove[:evi.start()], style1)
				else:
					final.write(num,33, prove, style1)
			else:
				if evi: 
					final.write(num,33, prove[:evi.start()])
				else: 
					final.write(num,33, prove)

			#evidences @31上述事实	
			if evi: 
				final.write(num,31, prove[evi.start():])	
			else: 
				final.write(num,31,"n/a")

			if prove=="": 
				continue 		
			year, remarry = prove_year_remarry(prove)
			str_petition, str_second = prove_petition(prove)
			total,str_year, str_gender = prove_children(prove)
			final.write(num,15,int(year))
			final.write(num,16,remarry)
			final.write(num,17,str_petition)
			final.write(num,18,str_second)
			final.write(num,19,total)
			final.write(num,20,str_year)
			final.write(num,21,str_gender)

	#everything in the jufa 
	result_col = []
	custody_col = []
	results = jufa.col(14, start_rowx=1, end_rowx=nrows_jufa)
	for r in results: 
		no = re.search('不准|驳回|无效', r.value)
		if no: 
			result_col.append('n')
			custody_col.append('n/a')
		else: 
			yes = re.search(r'准予|准许|原告[\u4e00-\u9fa5]*离婚|某离婚',r.value)
			if yes: 
				result_col.append('y')
				ss = re.split("。|；", r.value)
				here = True 
				for s in ss:
					if re.search("抚养", s): 
						#print(s)
						custody_col.append(s)
						here = False 
						break
				if here: 
					custody_col.append('n/a')
			else: 
				print("This is about RULING: " + r.value)
				result_col.append('n/a')
				custody_col.append('n/a')

	#本院认为，@27 
	opin_col = []
	comment_col = []
	opinion = jufa.col(13, start_rowx=1, end_rowx=nrows_jufa)
	for o in opinion: 
		opin_col.append(o.value)
		#court comments @35
		value = o.value
		sent = re.split("。", value)
		comment = "n/a" 
		for s in sent:
			jiabao = re.search("家庭暴力|家暴|打|证据", s)
			if jiabao:
				comment = s + "。"
		comment_col.append(comment)	

	write_same(int(the_year), 1, nrows_jufa, final)
	final.write(0, 25, "divorce?")
	write_col_gen(result_col, 25, style2, final)
	final.write(0, 26, "opinion?")
	write_col_gen(opin_col, 26, style3, final)
	final.write(0, 27, "custody")
	write_col(custody_col, 27, final)
	final.write(0, 28, "was mentioned?")
	write_same("y", 28, nrows_jufa, final)
	final.write(0, 29, "who was accused?")
	write_same("defendant", 29, nrows_jufa, final)
	final.write(0, 31, "evidences?")
	final.write(0, 32, "court looks into?")
	write_same("n", 32, nrows_jufa, final)
	final.write(0, 33, "经审理")
	final.write(0, 34, "court comments?")
	write_col(comment_col, 34, final)

	#the rows of the number 
	for i in range(0, nrows_jufa): 
		final.write(i+1, 0, int(start))
		start = start + 1 

if __name__ == '__main__':
	write_to = xlwt.Workbook() 
	write_sheet = write_to.add_sheet(u'sheet1', cell_overwrite_ok=True)
	
	#get the targeted file
	target = xlrd.open_workbook("../../Downloads/open广州.xlsx")
	
	read_(target, write_sheet) 
	beforeCom = "../2014广州after.xls"
	write_to.save(beforeCom)	
	target1 = xlrd.open_workbook(beforeCom)
	open_sheet = target1.sheets()[0]
	
	final = xlwt.Workbook()
	final_sheet = final.add_sheet(u'processed', cell_overwrite_ok=True)
	arrange = xlwt.Workbook()
	arrange_sheet = arrange.add_sheet(u'open', cell_overwrite_ok=True)
	
	jufa = xlrd.open_workbook("../../Downloads/jufa广州2014.xlsx")
	jufa_sheet = jufa.sheets()[0] 

	year = 2015
	start = 1
	compareToJufa(start, year, jufa_sheet, open_sheet, final_sheet,arrange_sheet,target)

	final.save("../2014广州comp.xls")	
	arrange.save("../arrangeOpen.xls")
