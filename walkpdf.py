#!/usr/bin/python
# -*- coding: utf-8 -*-

import sys
import os
import re
from pdfminer.pdfinterp import PDFResourceManager,process_pdf
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams

mypath = 'E:\\Work\\E_\\xxx\\xx'
#mypath = 'E:\\Work\\qqqqqqqqqqq'
#mypath = 'E:\\PyPro'

pat1 = 'Accounting'
pattern1 = re.compile(pat1, re.I|re.M)
pat2 = 'RADIUS'
pattern2 = re.compile(pat2, re.I|re.M)
pat3 = 'ACCT-SRV'
pattern3 = re.compile(pat3, re.I|re.M)

caching = True
codec = 'utf-8'
laparams = LAParams()

def walkdir(path):
	files = os.listdir(path)
	for file in files:
		fullpath = path+'\\'+file
		if os.path.isdir(fullpath):
			walkdir(fullpath)
		else:
			sufix = os.path.splitext(fullpath)[1][1:]
			if sufix == 'pdf':
				print fullpath
				processpdf(fullpath)
				
				
def processpdf(filename):
	#pdf
	outfp = open('E:\\temp.txt', 'w')
	rsrcmgr = PDFResourceManager(caching=caching)
	device = TextConverter(rsrcmgr, outfp, codec=codec, laparams=laparams)
	fp = open(filename, 'rb')
	try:
		process_pdf(rsrcmgr, device, fp, pagenos=set(), maxpages=0, password='', check_extractable=True)
	except:
		print 'process fail.'
	device.close()
	fp.close()
	outfp.close()
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

walkdir(mypath)