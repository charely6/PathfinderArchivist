import requests
import urllib.request
import locale
from ebooklib import epub
from ebooklib import plugins
from dialog import Dialog
import sys, os
import math
from tkinter import Tk
from tkinter import filedialog
import csv
from fuzzywuzzy import process
import pandas as pd
import uuid

def title(val):
	return val.title
def titleX(val):
	return val.title()

def first(val):
	return val[0]
	
def SpellNameFix(Name,List):
	try:
		List.index(Name)
		return Name
	except:
		x=process.extractOne(Name,List)
		return x[0]
	
def FileNameFix(Name):
	return 'OEBPS/'+Name.replace(' ','_')+'.xhtml'
	

#if not os.path.exists('spell_full.csv'):
#		urllib.request.urlretrieve('https://docs.google.com/spreadsheets/d/1cuwb3QSvWDD7GG5McdvyyRBpqycYuKMRsXgyrvxvLFI/pub?output=csv', 'spell_full.csv')
		
#pathtoCsv = r'https://docs.google.com/spreadsheets/d/1cuwb3QSvWDD7GG5McdvyyRBpqycYuKMRsXgyrvxvLFI/pub?output=csv'
pathtoCsv = r'spell_full.csv'
df = pd.read_csv(pathtoCsv, encoding = 'utf8')

list_of_rows = df.values.tolist()

list_of_spell_Names = df.loc[:,'name'].values.tolist()

book = epub.EpubBook()

# set metadata

uuidNum = uuid.uuid4()

book.set_identifier(str(uuidNum))
book.set_title('SpellBookTest')
book.set_language('en')

book.add_author('EpubSpellBook')
#book.FOLDER_NAME=
book.EPUB_VERSION=2.0

root = Tk()
root.withdraw()
root.filename =filedialog.askopenfilename( filetypes = ( ( "CSV", "*.csv"),("All files","*.*") ) )

filepath = root.filename

SpellList = []

with open(filepath,'r') as f:
	reader = csv.reader(f)
	for row in reader:
		for x in row:
			if x:
				SpellList.append(x.title())

chapters = []
ClassSections = []
classes = ['Sorcerer', 'Wizard', 'Cleric', 
		   'Druid', 'Ranger', 'Bard', 
		   'Paladin', 'Alchemist', 
		   'Summoner', 'Witch', 'Inquisitor', 
		   'Oracle', 'Antipaladin', 'Magus', 
		   'Adept']

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
font-family: arial;
}

h2 {
margin-bottom: 0;
margin-top: 0;
text-indent: -20px;
font-family: arial;
text-align:left;
}

h3 {
margin-bottom: 0;
margin-top: 0;
margin-left: -20px;
text-indent: 0px;
font-weight: normal;
font-family: arial;
}

h4 {
margin-bottom: 0px;
margin-top: 0px;
margin-left: -20px;
text-indent: 0px;
text-align: justify;
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
font-family: arial;
color: white;
}



'''
default_css = epub.EpubItem(uid="style_default", file_name="default.css", media_type="text/css", content=style)
book.add_item(default_css)
BookList = []
for y in SpellList:
	x=SpellNameFix(y,list_of_spell_Names)
	print(x)
	BookList.append(x)

for y in classes:
	ClassSections.append((epub.Section(y),[]))
#create files for each spell and add them to the class sections for the class lists they are on
	
for currentSpell in list_of_rows:
	if currentSpell[0]=="name":
		continue
	if currentSpell[0] not in BookList:
		continue
	cl = epub.EpubHtml(title = currentSpell[0], file_name=FileNameFix(currentSpell[0]), lang='en')
	cl.content=currentSpell[20].replace('h5','p')
	cl.add_item(default_css)
	for idx, y in enumerate(ClassSections):
		if not math.isnan(currentSpell[idx+26]):
			y[1].append(epub.Link(FileNameFix(currentSpell[0]),currentSpell[0],currentSpell[0].replace(' ','_')+'_'+classes[idx]))
	chapters.append(cl)

chapters.sort(key = title)

ClassSections = [s for s in ClassSections if len(s[1]) !=0]


for x in chapters:
	book.add_item(x)

TOCWClass = [epub.Link('nav.xhtml','TOC','TOC'),
			(epub.Section('A-Z'),chapters)]
#TOCWClass.extend(ClassSections)


book.toc = TOCWClass
#book.toc.extend(ClassSections)
			


# add default NCX and Nav file
book.add_item(epub.EpubNcx())
book.add_item(epub.EpubNav())


nav_css = epub.EpubItem(uid="style_nav", file_name="nav.css", media_type="text/css", content=style)

# add CSS file
book.add_item(nav_css)

# basic spine
book.spine = ['nav']
book.spine.extend(chapters)
book.guide = [{"href":"nav.xhtml","title":"Table of Contents", "type":"toc"}]
# write to the file
epub.write_epub('test.epub', book, {})



