# -*- coding: utf-8 -*-

import wx
from convert import Convert
import os, re, time

__version__ = "xls2csv Robot\n*************\nVersion 0.2.0\nCreated by Chu Rui\nchurui@ctecdcs.net"
__data__ = '21th Jan, 2015'


def xlName(name):
	path, basename = os.path.split(name)
	base_deExt = os.path.splitext(basename)
	return path, base_deExt[0]

def fileList(folder):
	file_list = []
	pre_list = os.listdir(folder)
	for basename in pre_list:
		if re.search('.xls[x]?$', basename):
			full_name = os.path.join(folder, basename)
			file_list.append(full_name)
	return file_list

class FileDrop(wx.FileDropTarget):
	def __init__(self, window):
		wx.FileDropTarget.__init__(self)
		self.window = window
		self.prompt = u'欢迎使用xls转csv机器人。请将xls文件所在的文件夹拖拽到此处。\n'

		self.window.WriteText(self.prompt)
		#self.window.WriteText('Drag and dorp the xls(x) files on me!\n')

	def OnDropFiles(self, x, y, folders):
		
		for folder in folders:
			xl_list = fileList(folder)
			if len(xl_list) != 0:
				start_time = time.time()
				converter = Convert()
				for xl_file in xl_list:
					#print "type(xl_list) %s" % type(xl_file)
					xl_name = os.path.basename(xl_file)  # xl_name is xl_file without its path.
					self.window.WriteText('%s' % xl_name + '....')
					csv_file = re.sub('.xls[x]?$', '.csv', xl_file)
					converter.xlConvert(xl_file, xl_name, csv_file)
					#converter.xlConvert()
					self.window.WriteText('%s' % 'OK\n')

				converter.terminate()
				end_time = time.time() - start_time
				self.window.WriteText(u'转换完成，用时%r秒!\n' % round(end_time,1))
			else:
				self.window.WriteText('No excel file in the directory!\n')

class DropFile(wx.Frame):
	def __init__(self, parent, id, title):
		wx.Frame.__init__(self, parent, id, title, size = (450, 400))
		#self.bkg = wx.Panel()
		
		helpmenu = wx.Menu()

		menuAbout = helpmenu.Append(wx.ID_ABOUT, 'About', ' About this program')
		helpmenu.AppendSeparator()
		menuExit = helpmenu.Append(wx.ID_EXIT, 'Exit', ' Terminate the program')

		menuBar = wx.MenuBar()
		menuBar.Append(helpmenu, 'Help')
		self.SetMenuBar(menuBar)

		self.Bind(wx.EVT_MENU, self.OnAbout, menuAbout)
		self.Bind(wx.EVT_MENU, self.OnExit, menuExit)

		self.text = wx.TextCtrl(self, -1, style = wx.TE_MULTILINE | wx.TE_READONLY)
		dt = FileDrop(self.text)
		self.text.SetDropTarget(dt)
		#self.btn = wx.Button(self, label='All in one')
		self.Centre()
		self.Show(True)
	
	def OnAbout(self, e):
		dlg = wx.MessageDialog(self, __version__, 'About xls2csv Robot', wx.OK)
		dlg.ShowModal()
		dlg.Destroy()
	
	def OnExit(self, e):
		self.Close(True)


app = wx.App()
DropFile(None, -1, 'xls2csv robot')
app.MainLoop()
