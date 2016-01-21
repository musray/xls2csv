# -*- coding: utf-8 -*-

import win32com.client as win32
from math import modf
import os, re

__version__ = 'v0.1.1, 21th Jan, 2015'

class Convert(object):
	'''
	xl_file :  the excel file name with absolute path.
	xl_name :  the excel file name without absolute path.
	csv_file:  the csv file name with absolute path. Only extion is different with xl_file  
	'''
	def __init__(self):
		self.handler = XL_Handler()

	def xlConvert(self, xl_file, xl_name, csv_file):
		'''
		look at the xls file name, determine which subfunction should be used to convert the xls to csv
		'''
		self.xl_file = xl_file
		self.xl_name = xl_name
		self.csv_file = csv_file

		if re.search('Opecom', self.xl_name,
				re.IGNORECASE):
			self._convert_opecom()

		elif re.search('Serial.*Send', self.xl_name,
				re.IGNORECASE):
			self._convert_serialsend()

		elif re.search('Sebus.*Send', self.xl_name,
				re.IGNORECASE):
			self._convert_sebussend()

		elif re.search('Recv', self.xl_name,
				re.IGNORECASE):
			self._convert_recv()

		elif re.search('Parameter', self.xl_name,
				re.IGNORECASE):
			self._convert_para()

		elif re.search('Drawing', self.xl_name,
				re.IGNORECASE):
			self._convert_drawing()

		elif re.search('AIO', self.xl_name,
				re.IGNORECASE):
			self._convert_AIO()

		elif re.search('DIO', self.xl_name,
				re.IGNORECASE):
			self._convert_DIO()

		elif re.search('PIF', self.xl_name,
				re.IGNORECASE):
			self._convert_PIF()


	def _convert_opecom(self):
		'''
		OAI, ODI files. Status: Done 
		'''
		self.handler.handler(self.xl_file)

		mess_rows = self.handler.ws1.Range("B2:D500").Value
		#mess_rows = xl.ws1.Range("B2:D500").Value
		#print mess_rows
		clear_rows = [] 
		for arow in mess_rows:
			if arow[0]: 
				clear_rows.append(arow)
		row_c = len(clear_rows)        # row_c means row_count
		self.handler.ws2.Range("A1:C%d" % row_c).Value = clear_rows
		self.handler.close(self.csv_file)

	def _convert_serialsend(self):
		'''
		serial send digital and analog files. Status: Done 
		'''
		self.handler.handler(self.xl_file)

		mess_rows = self.handler.ws1.Range("B4:AB1000").Value
		#print mess_rows
		clear_rows = [] 
		for arow in mess_rows:
			if arow[0]: 
				clear_rows.append(arow)
		row_c = len(clear_rows)        # row_c means row_count
		self.handler.ws2.Range("A1:AA%d" % row_c).Value = clear_rows
		self.handler.close(self.csv_file)

	def _convert_sebussend(self):
		'''
		W-NET send digital and analog files. Status: Done 
		'''
		self.handler.handler(self.xl_file)

		mess_rows = self.handler.ws1.Range("B4:P2000").Value
		#print mess_rows
		clear_rows = [] 
		for arow in mess_rows:
			if arow[0]: 
				clear_rows.append(arow)
		row_c = len(clear_rows)        # row_c means row_count
		self.handler.ws2.Range("A1:O%d" % row_c).Value = clear_rows
		self.handler.close(self.csv_file)

	def _convert_recv(self):
		'''
		serial send digital and analog files. Status: Not yet
		'''
		self.handler.handler(self.xl_file)

		mess_rows = self.handler.ws1.Range("B2:G1000").Value
		#print mess_rows
		clear_rows = [] 
		for arow in mess_rows:
			if arow[0]: 
				clear_rows.append(arow)
		row_c = len(clear_rows)        # row_c means row_count
		self.handler.ws2.Range("A1:F%d" % row_c).Value = clear_rows
		self.handler.close(self.csv_file)

	def _convert_para(self):
		'''
		prarameter list Status: Done
		'''
		self.handler.handler(self.xl_file)

		mess_rows = self.handler.ws1.Range("B2:F3500").Value
		#print mess_rows
		clear_rows = [] 
		for arow in mess_rows:
			if arow[0]: 
				clear_rows.append(arow)
		row_c = len(clear_rows)        # row_c means row_count
		self.handler.ws2.Range("A1:E%d" % row_c).Value = clear_rows
		self.handler.close(self.csv_file)

	def _convert_drawing(self):
		'''
		Drawing Management List. Status: Done
		'''
		self.handler.handler(self.xl_file)

		mess_rows = self.handler.ws1.Range("A1:K500").Value
		clear_rows = [] 
		for arow in mess_rows:
			if arow[1] and arow[7:9]!=(u'Not Changed',u'Not Changed'):
				clear_rows.append(arow)
		row_c = len(clear_rows)        # row_c means row_count
		self.handler.ws2.Range("A1:K%d" % row_c).Value = clear_rows
		self.handler.close(self.csv_file)

	def _convert_AIO(self):
		'''
		AIO. Status: Done 
		'''
		self.handler.handler(self.xl_file)

		mess_rows = self.handler.ws1.Range("B3:AO200").Value
		clear_rows = [] 
		for arow in mess_rows:
			if arow[0] and arow[2] and arow[3]:
				clear_rows.append(arow)
		row_c = len(clear_rows)        # row_c means row_count
		self.handler.ws2.Range("A1:AN%d" % row_c).Value = clear_rows
		self.handler.close(self.csv_file)

	def _convert_DIO(self):
		'''
		DIO. Status: Done 
		'''
		self.handler.handler(self.xl_file)

		mess_rows = self.handler.ws1.Range("B3:AC1500").Value
		clear_rows = [] 
		for arow in mess_rows:
			if arow[2] and arow[3]:
				clear_rows.append(arow)
		row_c = len(clear_rows)        # row_c means row_count
		self.handler.ws2.Range("A1:AB%d" % row_c).Value = clear_rows
		self.handler.close(self.csv_file)

	def _convert_PIF(self):
		'''
		PIF. Status: Done 
		'''
		self.handler.handler(self.xl_file)

		mess_rows = self.handler.ws1.Range("B3:AE7000").Value
		#print "mess_rows has %r elements" % len(mess_rows)
		STARTS = list(range(0, len(mess_rows),40))
		clear_rows = [] 
		for START in STARTS:
			relay = []
			COUNT = 0
			END = START + 40
			for arow in mess_rows[START:END]:
				relay.append(arow)
				if not arow[0]:
					COUNT += 1
			if COUNT < 40:
				for row in relay:
					clear_rows.append(row)
		row_c = len(clear_rows)        # row_c means row_count
		self.handler.ws2.Range("A1:AD%d" % row_c).Value = clear_rows
		self.handler.close(self.csv_file)
	
	def terminate(self):
		self.handler.quit()

class XL_Handler(object):
	def __init__(self): 
		#self.handle_file = handle_file 
		#self.content = []
		self.excel = win32.DispatchEx("Excel.Application")

	def handler(self, handle_file):
		#self.handle_file = handle_file
		self.wb1 = self.excel.Workbooks.Open(handle_file)
		#self.wb1.Visiable = 0
		self.ws1 = self.wb1.Worksheets(1)

		self.wb2 = self.excel.Workbooks.Add()
		#self.wb2.Visiable = 0
		self.ws2 = self.wb2.Worksheets(1)

	def close(self, csv_file):
		self.wb1.Close(SaveChanges=0)
		self.wb2.SaveAs(csv_file, 6)
		self.wb2.Close(SaveChanges=0)

	def quit(self):
		self.excel.Application.Quit()
