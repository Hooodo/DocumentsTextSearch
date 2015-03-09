#!/usr/bin/python
# -*- coding: utf-8 -*-

import os
import re
from win32com import client as wc
from pyExcelerator import *
from pdfminer.pdfinterp import PDFResourceManager,process_pdf
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams

DEBUG = False

class walkdoc:
	def __init__(self, dirpath, tempath='E:\\', logpath='E:\\', pat1='texttosearch', pat2=None, pat3=None, flagdoc=False, flagpdf=False, flagxls=False, flagtxt=False, flagalp=False, othertype=None):
		if DEBUG:
			print '*********Init class.*********'
		self.tempath = tempath
		self.dirpath = dirpath
		self.logpath = logpath
		self.pat1 = pat1
		self.pat2 = pat2
		self.pat3 = pat3
		self.flagdoc = flagdoc
		self.flagpdf = flagpdf
		self.flagxls = flagxls
		self.flagtxt = flagtxt
		self.flagalp = flagalp
		self.othertype = othertype
		
		if flagalp:
			if pat1:
				self.pattern1 = re.compile(pat1, re.M)	
			if pat2:
				self.pattern2 = re.compile(pat2, re.M)
			if pat3:
				self.pattern3 = re.compile(pat3, re.M)
		else:
			if pat1:
				self.pattern1 = re.compile(pat1, re.I|re.M)	
			if pat2:
				self.pattern2 = re.compile(pat2, re.I|re.M)
			if pat3:
				self.pattern3 = re.compile(pat3, re.I|re.M)
		
		if self.flagdoc:
			self.word = wc.Dispatch('Word.Application')
			wc.Visible = False
			wc.DisplayAlerts = False
			
		if self.flagpdf:
			caching = True
			codec = 'utf-8'
			laparams = LAParams()
			self.rsrcmgr = PDFResourceManager(caching)
			#self.device = TextConverter(self.rsrcmgr, outfp, codec=codec, laparams=laparams)
		
	def walkpath(self, path = None):
		if path:
			files = os.listdir(path)
			if DEBUG:
				print '*********In subpath.*********'
		else:
			files = os.listdir(self.dirpath)
			
		for file in files:
			if file.find('~') >= 0:
				continue
			if path:
				fullpath = path+'\\'+file
			else:
				fullpath = self.dirpath+'\\'+file
			
			if os.path.isdir(fullpath):
				if DEBUG:
					print 'Dir:%s' % fullpath
				self.walkpath(fullpath)			
			else:
				if DEBUG:
					print 'path:%s' % fullpath
				sufix = os.path.splitext(fullpath)[1][1:]
				if self.flagdoc and sufix == 'doc':
					print fullpath
					if self.processdoc(fullpath):
						self.processtemp(fullpath)
				elif self.flagxls and sufix == 'xls':
					print fullpath
					if self.processxls(fullpath):
						self.processtemp(fullpath)
				elif self.flagpdf and sufix == 'pdf':
					print fullpath
					if self.processpdf(fullpath):
						self.processtemp(fullpath)
				elif self.flagtxt and sufix == 'txt':
					print fullpath
					if self.processtxt(fullpath):
						self.processtemp(fullpath)
				elif self.othertype and sufix == self.othertype:
					print fullpath
					if self.processtxt(fullpath):
						self.processtemp(fullpath)
					
	def processdoc(self, path):
		if DEBUG:
			print '*********In processdoc.*********'
		try:
			doc = self.word.Documents.Open(path)
			doc.SaveAs(self.tempath+'temp', 4)
			doc.Close()
		except:
			print 'process fail.'
			return
			doc.Close()
			return False
		return True
			
	def processxls(self, path):
		if DEBUG:
			print '*********In processxls.*********'
		savef = open(self.tempath+'temp.txt', 'w')
		try:
			for sheet_name, values in parse_xls(path, 'cp1251'):
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
			return False
		savef.close()
		return True
		
	def processpdf(self, path):
		if DEBUG:
			print '*********In processpdf.*********'
		outfp = open(self.tempath+'temp.txt', 'w')
		fp = open(path, 'rb')
		device = TextConverter(self.rsrcmgr, outfp, codec='utf-8', laparams=LAParams())
		try:
			process_pdf(self.rsrcmgr, device, fp, pagenos=set(), maxpages=0, password='', check_extractable=True)
		except:
			print 'process fail.'
			return False
		device.close()
		fp.close()
		outfp.close()
		return True
		
	def processtxt(self, path):
		if DEBUG:
			print '*********In processtxt.*********'
		try:
			fp1 = open(path, 'r')
			fp2 = open(self.tempath+'temp.txt', 'w')
			fp2.writelines(fp1.read())
			fp1.close()
			fp2.close()
		except:
			print 'process fail'
			return False
		return True
		
	def processtemp(self, path):
		flag = False
		file = open(self.tempath+'temp.txt', 'r')
		buff = file.read()
		if self.pat1 and self.pattern1.search(buff):
			print '########## Has:  ' + self.pat1
			flag = True
		if self.pat2 and self.pattern2.search(buff):
			print '########## Has:  ' + self.pat2
			flag = True
		if self.pat3 and self.pattern3.search(buff):
			print '########## Has:  ' + self.pat3
			flag = True
		
		file.close()
		if flag:
			logfile = open(self.logpath+'log.txt', 'a')
			logfile.write('**'+path)
			logfile.write('\r\n')
			logfile.close()
					
def main(argv):
	if DEBUG:
		print '*********In main*********'
		#print argv
	import getopt
	def usage():
		print ('usage: %s [-w process doc] [-p process pdf] [-s process xls] [-x process txt] [-d path] [-t temp file] [-l log file]'
               '[-1 pattern1] [-2 pattern2] [-3 pattern3] [-o other type][-i]file ...' % argv[0])
		return 100
	try:
		(opts, args) = getopt.getopt(argv[1:], 'wpsxid:t:l:1:2:3:o:')
	except getopt.GetoptError:
		print 'error1'
		return usage()
		
	#print args
	#print opts
	#if not args: 
	#	print 'error2'
	#	return usage()
		
	flagdoc = False
	flagxls = False
	flagpdf = False
	flagtxt = False
	flagalp = False
	dirpath = None
	tempath = 'E:\\'
	logpath = 'E:\\'
	#pat1 = 'BST'
	pat1 = '192.168.12.140'
	pat2 = ''
	pat3 = ''
	#pat2 = None
	#pat3 = None
	othertype = None
	
	for (k, v) in opts:
		if k == '-w': flagdoc = True
		elif k == '-p': flagpdf = True
		elif k == '-s': flagxls = True
		elif k == '-x': flagtxt = True
		elif k == '-i': flagalp = True
		elif k == '-t': tempath = v
		elif k == '-l': logpath = v
		elif k == '-d': dirpath = v
		elif k == '-1': pat1 = v
		elif k == '-2': pat2 = v
		elif k == '-3': pat3 = v
		elif k == '-o': othertype = v
	
	ins = walkdoc(dirpath, tempath, logpath, pat1, pat2, pat3, flagdoc, flagpdf, flagxls, flagtxt, flagalp, othertype)
	ins.walkpath()
	
	return
	
if __name__ == '__main__': 
	if DEBUG:
		print '*********Start*********'
	sys.exit(main(sys.argv))
	