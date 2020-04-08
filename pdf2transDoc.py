import subprocess, sys

def install(package):
	subprocess.check_call([sys.executable, "-m", "pip", "install", package])
toimport = ['pypdf2', 'selenium', 'pyautoit', 'wkhtmltopdf']
for pck in toimport:
	install(pck)

from PyPDF2 import PdfFileWriter, PdfFileReader, PdfFileMerger
import os, shutil
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import threading
import autoit
from glob import glob
import pdfkit

def split_pdf(filename, groupsof):
	# Import your PDF file:
	inputpdf = PdfFileReader(open(filename, "rb"))
	# Create splitting sequence
	numpages = inputpdf.getNumPages()
	splitseq = list(range(groupsof,numpages+groupsof,groupsof))
	if numpages%groupsof != 0:
		splitseq[-1] = splitseq[-2]+numpages%groupsof
	# Create a directory temporary directory
	os.system('rmdir /S /Q "{}"'.format('tempdir'))
	os.mkdir('tempdir')
	# Split PDF
	firstpage = 0
	for lastpage in splitseq:
		output = PdfFileWriter()
		for p in range(firstpage,lastpage):
			output.addPage(inputpdf.getPage(p))
		with open("./tempdir/documentpage{}to{}.pdf".format(firstpage+1, lastpage), "wb") as outputStream:
			output.write(outputStream)
		firstpage = lastpage;

def setup_driver():
	driver = webdriver.Chrome("chromedriver.exe")
	driver.set_page_load_timeout(5)
	return(driver)

def main_process(driver, cwd, file, fromlang, tolang):
	driver.get("https://translate.google.com/#view=home&op=docs&sl={}&tl={}".format(fromlang,tolang))
	driver.find_element_by_id("tlid-file-input").send_keys("{}/tempdir/{}".format(cwd, file))
	driver.find_element_by_class_name("tlid-translate-doc-button").click()
	currurl = driver.current_url
	path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
	config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
	autoit.send("^p")
	time.sleep(2)
	autoit.send("{ENTER}")
	time.sleep(3)
	autoit.win_active("Save Print Output As")
	time.sleep(2)
	autoit.send("temporaryout_{}".format(file))
	autoit.send("{ENTER}")
	time.sleep(5)

def disconnect(driver):
	driver.quit()

def join_pdfs(files):
	merger = PdfFileMerger()
	for file in files:
		merger.append("temporaryout_{}".format(file))
	merger.write("Your_Translated_Document.pdf")
	merger.close()

def clear_mess(files):
	for file in files:
		os.remove("temporaryout_{}".format(file))
	os.system('rmdir /S /Q "{}"'.format('tempdir'))

##############################################################################################################################################################################
def main():
	print("STEP 1: Splitting pdfs appropriately...")
	split_pdf('pdffile.pdf', 4)
	print("						   ...Completed")
	cwd = os.getcwd()
	files = os.listdir("{}/tempdir/".format(cwd))
	print("STEP 2: Connecting with translate...")
	gtrans = setup_driver()
	print("						   ...Completed")
	print("STEP 3: Starting main sequence...")
	for file in files:
		main_process(gtrans, cwd, file, 'en', 'de')
	print("						   ...Completed")
	print("STEP 4: Starting finishing sequence...")
	disconnect(gtrans)
	join_pdfs(files)
	clear_mess(files)
	print("						   	  ...Completed")
	print("ENTIRE PROCESS COMPLETED. Your_Translated_Document.pdf available.")

if __name__ == '__main__':
	main()

