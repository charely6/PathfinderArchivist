import gspread
import pprint
import locale
from ebooklib import epub
from dialog import Dialog
import sys, os
import tkinter as tk
from tkinter import filedialog
from tkinter import *
from oauth2client.service_account import ServiceAccountCredentials
import csv
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import timeit
import pprint
def hasNumbers(inputString):
    return any(char.isdigit() for char in inputString)


def title(val):
	return val.title
def titleX(val):
	return val.title()
def first(val):
	return val[0]
	
def SpellNameFix(Name,List):
	try:
		y = List.index(Name)
		return Name
	except:
		print('======'+Name+'======')
		x=process.extractOne(Name,List)
		return x[0]



root = Tk()
root.withdraw()

filepath =filedialog.askopenfilename( filetypes = ( ( "CSV", "*.csv"),("All files","*.*") ) )


print(filepath)

SpellList = []

with open(filepath,'r') as f:
	reader = csv.reader(f)
	for row in reader:
		for x in row:
			if x:
				SpellList.append(x.title())
for spell in SpellList:
	print(spell.title())
# use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('EpubSpellBook.json', scope)
client = gspread.authorize(creds)

# Find a workbook by name and open the first sheet
# Make sure you use the right name here.
sheet = client.open('PathfinderSpells').sheet1

# Extract and print all of the values
list_of_rows = sheet.get_all_values()
#list_of_rows.sort(key = first)
#list_of_spell_Names = []
list_of_spell_Names = sheet.col_values(1)
#for x in list_of_rows:
#	list_of_spell_Names.append(x[0])


book = epub.EpubBook()

# set metadata
book.set_identifier('id123456')
book.set_title('SpellBookTest')
book.set_language('en')

book.add_author('EpubSpellBook')

chapters = []
ClassSections = []
classes = ['Sorcerer', 'Wizard', 'Cleric', 'Druid', 'Ranger', 'Bard', 'Paladin', 'Alchemist', 'Summoner', 'Witch', 'Inquisitor', 'Oracle', 'Antipaladin', 'Magus', 'adept']

# define CSS style

style = '''
/* Please include proper OGL citation if redistributed. */
/* OGL Section 15 Notice at end of file. */

div {
padding-top: 0px;
padding-bottom: 0px;
padding-right: 0px;
padding-left: 30px;
}

div.heading {
padding-top: 0px;
padding-bottom: 0px;
padding-right: 0px;
padding-left: 0px;
background-color: #663300;
height: 30px;
}

h1 {
margin-bottom: 0;
margin-top: 0;
text-indent: -20px;
font-size: 22;
font-family: arial;
}

h2 {
margin-bottom: 0;
margin-top: 0;
text-indent: -20px;
font-size: 22;
font-family: arial;
text-align:left;
}

h3 {
margin-bottom: 0;
margin-top: 0;
margin-left: -20px;
text-indent: 0px;
font-size: 16;
font-weight: normal;
font-family: arial;
}

h4 {
margin-bottom: 0px;
margin-top: 0px;
margin-left: -20px;
text-indent: 0px;
text-align: justify;
font-size: 16;
font-weight: normal;
font-family: arial;
}

h5 {
margin-bottom: 0;
margin-top: 0;
text-indent: -20px;
font-size: 16;
font-weight: normal;
font-family: arial;
}

h6 {
margin-bottom: 0;
margin-top: 0;
margin-left: 30px;
text-indent: -20px;
font-size: 16;
font-weight: normal;
font-family: arial;
}

p {
margin-top: 0px;
margin-right: 0px;
margin-bottom: 1px;
margin-left: 0px;
padding-top: 0px;
padding-right: 0px;
padding-bottom: 1px;
padding-left: 0px;
}

p.alignleft {
margin-left: 0px;
text-indent: 10px;
float: left;
text-align:left;
font-weight: bold;
font-size: 22;
font-family: arial;
color: white;
}

p.alignright {
float: right;
text-align:right;
font-weight: bold;
font-size: 22;
font-family: arial;
color: white;
}

/* Paizo Stat Block Database. Copyright 2011 Mike Chopswil, d20pfsrd.com */


'''   
#test

default_css = epub.EpubItem(uid="style_default", file_name="style/default.css", media_type="text/css", content=style)
book.add_item(default_css)
print("test")
BookList = []
for y in SpellList:
	x=SpellNameFix(y,list_of_spell_Names)
	#x=process.extractOne(y,list_of_spell_Names)
	BookList.append(x)

for y in classes:
	ClassSections.append((epub.Section(y),[]))
print(ClassSections)
print("================")
for x in list_of_rows:
	if x[0]=="name":
		continue
	if x[0] not in BookList:
		continue
	cl = epub.EpubHtml(title = x[0], file_name=x[0]+'.xhtml', lang='en')
	cl.content=x[20]
	cl.add_item(default_css)
	#print(x[0])
	for idx, y in enumerate(ClassSections):
		if hasNumbers(x[idx+26]):
			y[1].append(epub.Link(x[0]+'.xhtml',x[0],x[0]))
			
	chapters.append(cl)
ClassSections = [x for x in ClassSections if  x[1]]
for x in ClassSections:
	print('|'+x[0].title+'|')
	for y in x[1]:
		print(y.title,end = "  ")
	print()

chapters.sort(key = title)

for x in chapters:
	book.add_item(x)

#TOCList.append(chapters)
#print(TOCList)
#print('Class Sections')
#for x in ClassSections:
	#print(x)
print("================")
print(ClassSections)
print("================")

book.toc = (epub.Link('nav.xhtml','TOC','TOC'),
			(epub.Section('A-Z'),chapters))
			
print(book.toc)
print("================")
x = (epub.Link('nav.xhtml','TOC','TOC'),
			(epub.Section('A-Z'),chapters),ClassSections)
print(x)
#for x in ClassSections:
#	book.toc.append(x)
 

# # define Table Of Contents
# book.toc = (epub.Link('chap_01.xhtml', 'Introduction', 'intro'),
             # (epub.Section('Simple book'),
             # (c1, ))
            # )

# add default NCX and Nav file
book.add_item(epub.EpubNcx())
book.add_item(epub.EpubNav())


nav_css = epub.EpubItem(uid="style_nav", file_name="style/nav.css", media_type="text/css", content=style)

# add CSS file
book.add_item(nav_css)

# basic spine
book.spine = ['nav']
book.spine.extend(chapters)
# write to the file
epub.write_epub('test.epub', book, {})