import os
import re 

''' What to achieve ''' 
#1. sort the files in order based on the number 
#2. name the file from the start num to the end num in the sorted order 

def rename(path): 
	#list all the files under this director 
	files = os.listdir(path)
	#print(files)
	#change the names first 
	#print(files[0])
	the_path = path+"/"
	for f in files: 
		new_name = None 
		set_ = re.split("-", f)
		if len(set_) == 1: 
			new_name = f
		else: 
			new_name = set_[0]+".docx"
		os.rename(the_path+f, the_path+new_name)
		
	#print(the_path)
	new_files = os.listdir(path)
	new_files.sort(key= lambda x:int(x[:-5]))
	

def move(path, target_path, start): 
	files = os.listdir(path)
	the_path = path+"/"
	for f in files: 
		new_name = str(start) + ".docx" 
		os.rename(the_path+f, target_path+new_name)
		start = start + 1
	return start

def excel_order(start, path): 
	files = os.listdir(path)
	files.sort(key= lambda x:int(x[:-5]))
	for f in files: 
		os.rename(path+"/"+f, path+"/"+str(start)+".docx")
		start = start + 1

if __name__ == '__main__':
	downloads = "../../Downloads/"
	targets = ["2015丰都"]
	start_num = 1286
	start = len(os.listdir(downloads+targets[0])) + 1 
	target_path = downloads+targets[0]+"/"
	count = 1
	for t in targets: 
		dire = downloads+t
		#files = os.listdir(dire)
		#for f in files: 
		#	sets = re.split("-", f)
	#		print(sets[0])
		#print(files)
		rename(dire)
		if count != 1: 
			start = move(dire, target_path, start)
		count = count + 1 

	final_path = downloads+targets[0]
	final = os.listdir(final_path)
	final.sort(key= lambda x:int(x[:-5]))
	
	excel_order(start_num, final_path)

