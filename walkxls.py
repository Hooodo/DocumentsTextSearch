#!/usr/bin/python
# -*- coding: utf-8 -*-

import os
import re
import sys
from pyExcelerator import *

#mypath = 'E:\\Work\\E_\\xxx\\xx'
mypath = 'E:\\Work\\qqqqqqqqqqq'
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
			if sufix == 'xls':
				print fullpath
				processdoc(fullpath)
				
				
def processdoc(filename):
	#xls
	savef = open('E://temp.txt', 'w')
	try:
		for sheet_name, values in parse_xls(filename, 'cp1251'):
			buf = 'Sheet = "%s"' % sheet_name.encode('cp866', 'backslashreplace')
			buf += '\n'
			savef.write(buf)
			for row_idx, col_idx in sorted(values.keys()):
				v = values[(row_idx, col_idx)]
				if isinstance(v, unicode):
					v = v.encode('cp866', 'backslashreplace')
				else:
					v = `v`
				buf =  v+'\n'
				savef.write(buf)
	except:
		print 'process fail.'
	savef.close()
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