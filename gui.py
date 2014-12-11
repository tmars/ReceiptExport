#coding=utf8
from Tkinter import *
from tkMessageBox import *
from tkFileDialog import *
import os
from excel import excel_open, excel_quit, get_products, save_products
import fileinput
import shutil

class window():
	def __init__(self):
		self.sourceFilename = ""
		self.plusFilename = ""

		self.root = Tk()
		self.root.geometry('{}x{}'.format(250, 80))
		self.root.resizable(width=FALSE, height=FALSE)

		self.sourceBut = Button(self.root)
		self.sourceBut.grid(row=0, column=0)
		self.sourceBut["text"] = "Куда добавить"
		self.sourceBut.bind("<Button-1>", self.openSource)

		self.sourceLbl = Label(self.root)
		self.sourceLbl.grid(row=0, column=1)

		self.plusBut = Button(self.root)
		self.plusBut.grid(row=1, column=0)
		self.plusBut["text"] = "Что добавить"
		self.plusBut.bind("<Button-1>", self.openPlus)

		self.plusLbl = Label(self.root)
		self.plusLbl.grid(row=1, column=1)

		self.execBut = Button(self.root)
		self.execBut.grid(row=2, columnspan=2)
		self.execBut["text"] = "Перенести товары"
		self.execBut.bind("<Button-1>", self.transfer)

	def start(self):
		self.root.mainloop() 

	def openSource(self, event):
		self.sourceFilename = askopenfilename(filetypes=(("Файлы Excel", "*.xls;*.xlsx"),))
		self.sourceLbl['text'] = os.path.basename(self.sourceFilename)
		

	def openPlus(self, event):
		self.plusFilename = askopenfilename(filetypes=(("Файлы Excel", "*.xls;*.xlsx"),))
		self.plusLbl['text'] = os.path.basename(self.plusFilename)
		
	def transfer(self, event):
		if self.sourceFilename == "": 
			showerror('Внимание', 'Не выбрано куда добавлять.')
			return 
			
		if self.plusFilename == "":
			showerror('Внимание', 'Не выбрано что добавлять.')
			return 
		
		newFilename = os.path.dirname(self.sourceFilename) + u'/КОПИЯ ' + os.path.basename(self.sourceFilename)
		print newFilename
		shutil.copyfile(self.sourceFilename, newFilename)

		excel_open()
		try:
			products = get_products(self.plusFilename)
			save_products(self.sourceFilename, products)
		except Exception as e:
			showerror('Ошибка:' + e.__class__.__name__, str(e))

		excel_quit()

		showinfo('Успех', 'Мы перенесли %d товаров.' % len(products))

w = window()
w.start()