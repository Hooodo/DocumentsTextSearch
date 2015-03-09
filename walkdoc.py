#!/usr/bin/python
# -*- coding: utf-8 -*-

import os
import re
from win32com import client as wc

mypath = 'E:\\Work\\E_\\xxx\\xx'
#mypath = 'E:\\Work\\qqqqqqqqqqq'
#mypath = 'E:\\PyPro'

pat1 = 'Accounting'
pattern1 = re.compile(pat1, re.I|re.M)
pat2 = 'RADIUS'
pattern2 = re.compile(pat2, re.I|re.M)
pat3 = 'ACCT-SRV'
pattern3 = re.compile(pat3, re.I|re.M)

def walkdir(path):
	files = os.listdir(path)
	for file in files:
		fullpath = path+'\\'+file
		if os.path.isdir(fullpath):
			walkdir(fullpath)
		else:
			sufix = os.path.splitext(fullpath)[1][1:]
			if sufix == 'doc':
				print fullpath
				processdoc(fullpath)
				
				
def processdoc(filename):
	#doc
	try:
		doc = word.Documents.Open(filename)
		doc.SaveAs('E://temp', 4)
		doc.Close()
	except:
		print 'process fail.'
		return
		doc.Close()
	#text
	flag = False
	file = open('E://temp.txt', 'r')
	buff = file.read()
	if pattern1.search(buff) != None:
		print 'has  ' + pat1
		flag = True
	if pattern2.search(buff) != None:
		print 'has  ' + pat2
		flag = True
	if pattern3.search(buff) != None:
		print 'has  ' + pat3
		flag = True
		
	file.close()
	if flag:
		logfile = open('E://log.txt', 'a')
		logfile.write('**'+filename)
		logfile.write('\r\n')
		logfile.close()
	
	
word = wc.Dispatch('Word.Application')
wc.Visible = False

walkdir(mypath)