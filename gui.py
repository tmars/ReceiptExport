#coding=utf8
from Tkinter import *
from tkMessageBox import *
from tkFileDialog import *
from excel import get_products, save_products
import fileinput

sourceFilename = ""
plusFilename = ""

def openSource(event):
	global sourceFilename
	sourceFilename = askopenfilename(filetypes=(("Файлы Excel", "*.xls;*.xlsx"),))
	print 'source', sourceFilename

def openPlus(event):
	global plusFilename
	plusFilename = askopenfilename(filetypes=(("Файлы Excel", "*.xls;*.xlsx"),))
	print 'plus', plusFilename

def transfer(event):
	global sourceFilename, plusFilename  
	if sourceFilename == "": 
		showerror('Внимание', 'Не выбрано куда добавлять.')
		return 
		
	if plusFilename == "":
		showerror('Внимание', 'Не выбрано что добавлять.')
		return 

	products = get_products(plusFilename)
	save_products(sourceFilename, products)
	showinfo('Успех', 'Мы перенесли %d товаров.' % len(products))

root = Tk()

sourceBut = Button(root)
sourceBut.grid(row=0, column=0)
sourceBut["text"] = "Куда добавить"
sourceBut.bind("<Button-1>",openSource)
#sourceBut.pack()

plusBut = Button(root)
plusBut.grid(row=0, column=1)
plusBut["text"] = "Что добавить"
plusBut.bind("<Button-1>",openPlus)
#plusBut.pack()

execBut = Button(root)
execBut.grid(row=1, columnspan=2)
execBut["text"] = "Перенести товары"
execBut.bind("<Button-1>",transfer)
#execBut.pack()


root.mainloop() 