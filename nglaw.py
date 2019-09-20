import xlrd #read 
import xlwt #write 
import xlsxwriter
import os 
import re #regularization


#TODO: group process many excels at the same time in the director 
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
			elif title == '当事人': 
				dangshiren = sheet.col_values(i, start_rowx=1, end_rowx=nrows)
				process_danshiren(dangshiren, write_sheet)
			elif title == '庭审程序说明': 
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

def process_danshiren(dangshiren, write_sheet):
	yuanGao_gender = []
	beiGao_gender = []
	yuanGao_DOB = []
	beiGao_DOB = []
	yuanGao_legal = []
	beiGao_legal = []
	for i in range(0, len(dangshiren)): 
		str_ = dangshiren[i]
		#cleaninng the data
		str_ = str_.replace('、', '')
		#separate 被告和原告
		two = str_.split('被告') 
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
			not_sum = re.search('合议庭|和议庭|审判员陈建平于2014年2月25日', court)
			if not_sum: 
				summary_col.append('n')
			else: 
				num = i + 2
				print(str(num)+": " + court)
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

def process_prove(prove, write_sheet):
	year_return = None 
	remarried_return = None 
	highlight = None 

	marriage = re.search(r'结婚', prove)
	if marriage: 
		the_ = re.split("，", prove)
		for index, t in enumerate(the_): 
			year_ = re.search(r'[0-9][0-9][0-9][0-9]年.*结婚', t) 
			if year_:
				all_ = t[year_.start():year_.end()]
				mar_year = re.search(r'[0-9][0-9][0-9][0-9]', all_)
				year_return = all_[mar_year.start():mar_year.end()]
				break
			else: 
				#同年结婚
				tong = re.search(r'同年.*结婚', t)
				if tong: 
					f0 = index-1 
					f1 = True
					y_ = 0
					while(f1): 
						if f0 >= 0 and f0 < len(the_): 
							b = re.search(r'[0-9][0-9][0-9][0-9]年(.*(相识|同居|恋爱|建立关系|认识))*', the_[f0])
							if b: 
								y_ = re.findall(r'[0-9][0-9][0-9][0-9]', the_[f0])[0]
								y_ = int(y_)
								f1 = False 
							else:
								f0 = f0 -1 
						else: 
							f1 = False  
					year_return = y_
				else: 
					year_return = 0
	else:
		year_return = 0
	#if year_return == 0: 
		#print(prove)	

	#print("ToFInd: ", year_return)

		#decide whether it is remarried or not 
	remarried = re.search(r'再婚', prove)
	if remarried:
		remarried_return = 'y'
	else: 
		remarried_return = 'n'

	found = re.search(r"家庭暴力|家暴|纠纷|打骂|争执|打闹|致伤|损伤|吵闹", prove)
	if found: 
		highlight = True 
	else: 
		highlight = False 
	#separate the sentences 
	sentences = re.split('。', prove)
	sum_petition = 0
	str_petition = "("
	str_second = ""

	sum_female = 0 
	sum_male = 0

	year_list = []
	str_year = ""
	str_gender = ""	 
	
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
		
		comma = re.split("，",s)
		
		for c in comma: 
			#find the number of children 
			girls = re.findall(r'一女|长女|次女|生育女|婚生女', c)
			boys = re.findall(r'一子|长子|次子|生育([\u4e00-\u9fa5]*)儿|生育男|婚生子', c)
			sum_female = sum_female+len(girls)
			sum_male = sum_male+len(boys)
		
			if len(girls) != 0 or len(boys) != 0:
				#new_s = re.search('结婚|登记', s)
				#ss = None
				#if new_s: 
				#	ss = s[new_s.end():]
				#else: 
				#	ss = s
				years = re.findall(r'[0-9][0-9][0-9][0-9]年', c)
		#if len(years) != 0:
		#		year_list.append(int(str(years[-1][0:4])))
		#	sum_female = sum_female + 1
				for y in years: 
					year_list.append(int(str(y[0:4])))

		#for b in boys: 
		#	years = re.findall(r'[0-9][0-9][0-9][0-9]年', s)
		#	if len(years) != 0:
		#		year_list.append(int(str(years[-1][0:4])))
		#	sum_male = sum_male + 1 

	#outside of the for loop 		
	str_petition = str_petition + ")"
	if sum_petition == 0: 
		str_petition = "y"
		str_second = "n"
	else: 
		str_petition = "n"+str_petition
		str_second = "y"

	#process the chlidren problems 
	year_list.sort()
	for y in year_list:
		str_year = str(str_year) + str(y) + ","
	if len(year_list) == 0:
		str_year = "n/a"
	else: 
		str_year = str_year[:len(str_year)-1]
		str_year = str_year+"."
	#if len(year_list)>=3: 
	#	print(prove)

	#print(str(sum_female)+", "+str(sum_male))
	if sum_female != 0: 
		str_gender = str(sum_female)+" female,"
		#print("There is female")
	if sum_male != 0:
		#print("WHY NOT")
		str_gender = str_gender + str(sum_male)+" male."
	
	total = sum_female+sum_male
	if total == 0: 
		str_gender = "n/a"
	#print(str(total)+", "+str_year+", "+str_gender)
	return year_return, remarried_return, str_petition, str_second, total,str_year, str_gender, highlight 

#both parties @14 
#plantiff reasons @22
def process_court2(court2, write_sheet): 
	both_col = []
	num = 2
	plantiff_reasons = []
	defendant_reasons = []
	proven = []
	defendant_agree = []
	year_list = []
	remarried_list = []

	first_l = []
	second_l = []
	total_child_l = []
	child_year_l = []
	child_gender_l = []
	highlight_l = []

	dispute_l = []

	pattern1 = xlwt.Pattern() 
	pattern1.pattern = xlwt.Pattern.SOLID_PATTERN
	pattern1.pattern_fore_colour = 5 #this is yello 
	style1 = xlwt.XFStyle()
	style1.pattern = pattern1 

	for i in range(0, len(court2)): 
		
		line = court2[i]
		#remove the 。、and replace it with 。
		line = line.replace('。、', '。')
		line = line.replace('。；', '。')

		#find if the defendant is present 
		absence = re.search("依法缺席", line)
		if absence: 
			#to decide which one is present and which one is not 
			present = re.search(r"到庭参加([\u4e00-\u9fa5]*)诉讼", line)
			
			if present: 
				pre_line = line[:present.start()]
				yuanG = re.search("原告", pre_line)
			else: 
				print("FOUND THE PROBLEM for both attendant")
				#print(i+1, line)
			if yuanG: 
				both_col.append("n(miss defendant)")
			elif re.search("被告", pre_line): 
				both_col.append("n(miss plantiff)")
		else: 
			both_col.append('y')

		#plantiff reason 
		#诉称
		sets = line.split("诉称", 1)
		second = sets[1]
		second = second[1:] #remove the first comma and others in rare cases 

		'''if num == 5: 
			try_ = re.search(r'被告([\u4e00-\u9fa5]*)，([\u4e00-\u9fa5]*)未向本院([\u4e00-\u9fa5]*)', second)
			if try_: 
				print("find")
				bei = second[try_.start():try_.end()]
				print(str(num)+": "+bei)
			else: 
				print("OH NO")'''
		#1. 被告xxx辩称
		#2. 被告xxx未提交书面答辩状。
		#whether it agrees to divorce	
		
		#there might be the case that the above are in front of the 被告
		prove = re.search(r'经审理查明|经庭审查明|本院确认如下事实|对本案事实..如下|根据当事人举证质证，对本案事实认定如下|确认以下事实', second)
		if prove: 
			#remove the last part 
			rest = second[:prove.start()]
			#contains the YuanGao, Beigao argument 
			sep1 = re.search(r'被告([\u4e00-\u9fa5]*)没有到庭|被告([\u4e00-\u9fa5]*)(×|\d)*辩称|被告([\u4e00-\u9fa5]*)答辩|被告([\u4e00-\u9fa5]*)未到庭|被告([\u4e00-\u9fa5]*)，*未到庭|被告([\u4e00-\u9fa5]*)，*未有答辩|被告([\u4e00-\u9fa5]*)辩称|被告([\u4e00-\u9fa5]*)，*无书面|被告([\u4e00-\u9fa5]*)，*未提交书面|被告[\u4e00-\u9fa5]*，[\u4e00-\u9fa5]*未向本院[\u4e00-\u9fa5]*|([\u4e00-\u9fa5]*)未向本院([\u4e00-\u9fa5]*)', rest)
			yuan = None 

			if sep1:
				yuan = rest[:sep1.start()]
				bei = rest[sep1.start():]
				
				plantiff_reasons.append(yuan)
				defendant_reasons.append(bei)

				#decude whether it is disputed or not 
				dispute = re.search(r'没[\u4e00-\u9fa5]*家[\u4e00-\u9fa5]*暴|无[\u4e00-\u9fa5]*家[\u4e00-\u9fa5]*暴|家[\u4e00-\u9fa5]*暴[\u4e00-\u9fa5]*不是', bei)
				if dispute: 
					dispute_l.append("y")
				elif absence: 
					dispute_l.append("n/a")
				else: 
					dispute_l.append("n")
			#print(str(num)+": "+yuan)
			#print(str(num)+": "+bei)


			#if prove: 
			#	bei_full = rest[:prove.start()]
			#	clarify = rest[prove.start():]
			#	defendant_reasons.append(bei_full)
			#	proven.append(clarify)
				#year, remarry, str_petition, str_second, total,str_year, str_gender, highlight = process_prove(clarify, write_sheet)
				#year_list.append(year)
				#remarried_list.append(remarry)
				#first_l.append(str_petition)
				#second_l.append(str_second)
				#total_child_l.append(total)
				#child_year_l.append(str_year)
				#child_gender_l.append(str_gender)
				#highlight_l.append(highlight)

				#decide whether the defendant agrees or not 
				opt = re.search(r'不同意([\u4e00-\u9fa5]*)离婚', bei)
				if opt: 
					defendant_agree.append('n')
				elif re.search(r'同意([\u4e00-\u9fa5]*)离婚', bei):
					defendant_agree.append('y')
				else: 
					defendant_agree.append('n/a')
			else: 
				print(str(num)+": DID NOT FOUND 被告ARGUMENT")
				print(rest)
				plantiff_reasons.append("N/A")
				defendant_agree.append('n/a')
				defendant_reasons.append("n/a")

			#经审理@33
			num = i + 1 
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
				#print(str(num)+": "+prove_)
				sss = re.search(r'下列证据|原告([\u4e00-\u9fa5])*为证明自己的主张|原告([\u4e00-\u9fa5])*(为)*证明', yuan)
				if sss: 
					write_sheet.write(num, 31, yuan[sss.start():])
					#print(str(num)+": "+yuan[sss.start():])
				else: 
					#print(str(num)+": "+yuan)
					#can continue to find this in the beigao section based on the modified keywords 
					write_sheet.write(num,31,"n/a")

			year, remarry, str_petition, str_second, total,str_year, str_gender, highlight = process_prove(prove_, write_sheet)
			#year married @17
			#print(year)
			write_sheet.write(num,15,int(year))
			#remarried yes, or not @18 
			write_sheet.write(num,16,remarry)
			#first petition @19
			write_sheet.write(num,17,str_petition)	
			#second petition @20
			write_sheet.write(num,18,str_second)
			#the number of children @21 
			write_sheet.write(num,19,total)
			#the age of the children @22
			write_sheet.write(num,20,str_year)
			#gender of the children @23
			write_sheet.write(num,21,str_gender)

		#when there is no parse 		
		else: 
			print(str(num)+": DID NOT FOUND 庭审证明")
			plantiff_reasons.append("N/A")
			print(second)
		num = num + 1

	#write both parties 
	write_sheet.write(0, 14, 'both parties')
	write_col(both_col, 14, write_sheet)
	
	#year married @15
	write_sheet.write(0,15,"year married")
	#remarried yes, or not @16 
	write_sheet.write(0,16,"remarrried")
	#first petition @19
	write_sheet.write(0,17,"first petition")	
	#second petition @20
	write_sheet.write(0,18,"second petition")
	#the number of children @21 
	write_sheet.write(0,19,"num of children")
	#the age of the children @22
	write_sheet.write(0,20,"age of children")
	#gender of the children @23
	write_sheet.write(0,21,"gender of children")

	write_sheet.write(0,22, 'plantiff')
	write_col(plantiff_reasons, 22, write_sheet)	

	#defendant reasons @23
	write_sheet.write(0, 23, "defendant reasons")
	write_col(defendant_reasons, 23, write_sheet)
	#defendant agree to divorce @24
	write_sheet.write(0,24, "defendant agrees")
	write_col(defendant_agree, 24, write_sheet)
	
	#dispute or not @31 
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

#案号10 col
def compareToJufa(start, the_year, jufa, open_, final):
	pattern1 = xlwt.Pattern() 
	pattern1.pattern = xlwt.Pattern.SOLID_PATTERN
	pattern1.pattern_fore_colour = 5 #this is yello 
	style1 = xlwt.XFStyle()
	style1.pattern = pattern1 

	pattern2 = xlwt.Pattern() 
	pattern2.pattern = xlwt.Pattern.SOLID_PATTERN
	pattern2.pattern_fore_colour = xlwt.Style.colour_map['pink'] #red
	style2 = xlwt.XFStyle()
	style2.pattern = pattern2

	pattern3 = xlwt.Pattern() 
	pattern3.pattern = xlwt.Pattern.SOLID_PATTERN
	pattern3.pattern_fore_colour = xlwt.Style.colour_map['pale_blue']  #blue
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
	#print(name_row)
	jufa_v = []
	open_v = []
	for i in jufa_case: 
		jufa_v.append(i.value)
	for i in open_case: 
		open_v.append(i.value)
	write_row(name_row, 0, final, style1, style2, style3)

	for i in range(0, len(jufa_case)): 
		num = i + 1 
		case = jufa_v[i]
		if case in open_v:
			open_yes.append(case)
			index = open_v.index(case)
			row = open_.row(index+1)
			write_row(row, num, final, style1, style2, style3)
		#not in the openLaws
		else: 
			jufa_not.append(case)	
			#self add in the information 
			the_row = jufa.row(num)
			name = the_row[1].value 
			final.write(num,3, case)
			final.write(num,4, name)

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

			year, remarry, str_petition, str_second, total,str_year, str_gender, highlight = process_prove(prove, final)
			#year married @17
			#print(year)
			final.write(num,15,int(year))
			#remarried yes, or not @18 
			final.write(num,16,remarry)
			#first petition @19
			final.write(num,17,str_petition)	
			#second petition @20
			final.write(num,18,str_second)
			#the number of children @21 
			final.write(num,19,total)
			#the age of the children @22
			final.write(num,20,str_year)
			#gender of the children @23
			final.write(num,21,str_gender)

	#result @26
	result_col = []
	custody_col = []
	results = jufa.col(14, start_rowx=1, end_rowx=nrows_jufa)
	for r in results: 
		no = re.search('不准|驳回|无效', r.value)
		if no: 
			result_col.append('n')
			custody_col.append('n/a')
		else: 
			yes = re.search('准予|准许',r.value)
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
				if re.search(r'某离婚|杨某甲离婚', r.value): 
					result_col.append('y')
				else:
					print(r.value)
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
			#zheng = re.search("证据", s)
			jiabao = re.search("家庭暴力|家暴|打", s)
			if jiabao:
				comment = s + "。"

		#print(comment)
		comment_col.append(comment)	

	#divorce? @25
	final.write(0, 25, "divorce?")
	write_col_gen(result_col, 25, style2, final)
	final.write(0, 26, "opinion?")
	write_col_gen(opin_col, 26, style3, final)
	final.write(0, 27, "custody")
	write_col(custody_col, 27, final)
	#35
	final.write(0, 35, "court comments?")
	#write_col(comment_col, 35, final)	
	#the one after is the custody @28 
	#was mentioned? @28 
	final.write(0, 28, "was mentioned?")
	write_same("y", 28, nrows_jufa, final)
	#who @30 
	final.write(0, 29, "who was accused?")
	write_same("defendant", 29, nrows_jufa, final)
	#@1, year 
	write_same(int(the_year), 1, nrows_jufa, final)
	
	#evidences @32上述事实
	final.write(0, 31, "evidences?")
	#did the one for the court @33 
	final.write(0, 32, "court looks into?")
	write_same("n", 32, nrows_jufa, final)

	#经审理@33
	final.write(0, 33, "经审理")

	#court comments, @35
	final.write(0, 35, "court comments?")
	write_col(comment_col, 35, final)
	#final.write_same(")

	for i in range(0, nrows_jufa): 
		final.write(i+1, 0, int(start))
		start = start + 1 

	open_not = list(set(open_v)-set(open_yes))
	return jufa_not, open_not 

if __name__ == '__main__':
	write_to = xlwt.Workbook() 
	write_sheet = write_to.add_sheet(u'sheet1', cell_overwrite_ok=True)
	
	#get the targeted file
	target = xlrd.open_workbook("../Downloads/2014奉节open.xlsx")
	
	read_(target, write_sheet) 
	write_to.save("../Downloads/2014奉节after.xls")	
	
	target1 = xlrd.open_workbook("../Downloads/2014奉节after.xls")
	open_sheet = target1.sheets()[0]
	final = xlwt.Workbook()
	final_sheet = final.add_sheet(u'final', cell_overwrite_ok=True)

	#adjust to the Jufa cases 
	jufa = xlrd.open_workbook("../Downloads/2014奉节.xlsx")
	jufa_sheet = jufa.sheets()[0] 

	year = 2015
	start = 596
	jufa_not, open_not = compareToJufa(start, year, jufa_sheet, open_sheet, final_sheet)

	final.save("../Downloads/2014奉节comp3.xls")	

