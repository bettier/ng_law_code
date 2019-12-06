from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import time 
import os 
import xlrd 
import xlwt 

def folder(driver, fnames, num,write_sheet):
    path = "file:///Users/bettyzhou/Desktop/"+fnames
    total = 0 
    for i in range(num): 
        n = i + 1 
        the_path = path + "/"+str(n)+".htm"
        driver.get(the_path)
        cases = driver.find_elements_by_class_name("ah")
        for c in cases: 
            write_sheet.write(total,0,c.text)
            total = total + 1 
if __name__ == '__main__':
    write_to = xlwt.Workbook() 
    write_sheet = write_to.add_sheet(u'2015C', cell_overwrite_ok=True)
    driver = webdriver.Chrome(ChromeDriverManager().install())
    folder(driver,"2015C",98,write_sheet)
    write_to.save("2015C.xls")
