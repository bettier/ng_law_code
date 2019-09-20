import os
import re 

''' What to achieve ''' 
#1. sort the files in order based on the number 
#2. name the file from the start num to the end num in the sorted order 

def rename(path): 
	#list all the files under this director 
	files = os.listdir(path)
	#change the names first 
	#print(files[0])
	the_path = path+"/"
	for f in files: 
		new_name = f.split("-")[0]+".docx"
		os.rename(the_path+f, the_path+new_name)
	#print(the_path)
	new_files = os.listdir(path)
	new_files.sort(key= lambda x:int(x[:-5]))
	print(new_files)

if __name__ == '__main__':
	downloads = "../Downloads/"
	targets = ["巫溪2"]

	for t in targets: 
		dire = downloads+t
		rename(dire)