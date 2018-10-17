# -*- coding: utf-8 -*-

import os
import sys
import win32com
from win32com import client as wc
from docx import Document
from docx.shared import Inches
import argparse
import config 
import re
import md5
reload(sys)
sys.setdefaultencoding('gb18030')

print "[*]For use info:python main.py -h,--help"

parser = argparse.ArgumentParser(description='Scan Mission Params Needed')

parser.add_argument('-f', help = "folders you want to watermark" , dest = "floder")
parser.add_argument('-c', help = "config about watermark product" , dest = "confpath")
parser.add_argument('-o', help = "new floder you want to save" ,  dest = "outpath")
parser.add_argument('-l', help = "log of relationship before file and watermark" , dest = "logpath")

args = parser.parse_args()
floder = args.floder
confpath = args.confpath
outpath = args.outpath
logpath = args.logpath
conf = config.get_conf(confpath)

def get_file(floder):
	word_list = []
	file_list = os.listdir(floder)
	for i in range(0,len(file_list)):
		file_path = os.path.join(floder,file_list[i])
		if os.path.isfile(file_path):
			if 'doc' in file_path.split('.')[-1]:
				word_list.append(file_path)
				#print file_path
	return word_list

def watermark(file):
	document = Document(file)
	document.add_picture('c:\\users\\aldin\\desktop\\scribles\\1.JPG', width=Inches(0.25))
	document.save(outpath + '\\' + file.split('\\')[-1].split('.')[0] + '.docx')
	print "[*]now watermark file name :" + file + "finished"               
	word = wc.Dispatch('Word.Application') 
	doc = word.Documents.Open(outpath + '\\' + file.split('\\')[-1].split('.')[0] + '.docx') 
	doc.SaveAs(outpath + '\\' + file.split('\\')[-1].split('.')[0] + '.html', 10) 
	doc.Close() 
	word.Quit()
	print "[*]now convert word to html :" + file + "finished"
	f = open(outpath + '\\' + file.split('\\')[-1].split('.')[0] + '.html')
	content = f.readlines()
	m1 = md5.new()
	m1.update(file)
	random_str = m1.hexdigest()
	watermark_content = conf['url'] + '://' + conf['host'] + '/' + conf['path'] + '/' + conf['filename'] + '/' + random_str + '/'
	print watermark_content
	for i in range(0,len(content)):
		if 'src="Test_files/image' in content[len(content)-1-i]:
			content[len(content)-1-i] = content[len(content)-1-i].replace('Test_files/',watermark_content)
			#print content[len(content)-1-i]
			break
	f.close()
	print "[*]now replace with url :" + file + "finished"
	with open(logpath,'w+') as f:
		f.write(file + '  ||  ' + watermark_content + '\n')
	f.close()
	print "[*]now write log :" + file + "finished"
	html_content = ""
	for i in range(0,len(content)):
		#print content[i]
		html_content = html_content + content[i]

	with open(outpath + '\\' + file.split('\\')[-1].split('.')[0] + '.html',"w") as f:
		f.write(html_content)
	f.close()

def savefile(file):
	#print outpath + '\\' + file.split('\\')[-1].split('.')[0] + '.doc'
	word = wc.Dispatch('Word.Application')
	
	doc = word.Documents.Open(outpath + '\\' + file.split('\\')[-1].split('.')[0] + '.html')
	
	doc.SaveAs(outpath + '\\' + file.split('\\')[-1].split('.')[0] + '.doc', FileFormat=0)
	
	doc.Close()
	word.Quit() 
	print "[*]now handle file :" + file + "finished"


doc_list = get_file(floder)
for i in range(0,len(doc_list)):
	watermark(doc_list[i])
	savefile(doc_list[i])
exit()

