# -*- coding: utf-8 -*-
"""
Created on Wed Oct  2 08:11:41 2019

@author: Spare1
"""
import ebooklib
from ebooklib import epub


book = epub.read_epub('Zokthra.epub')

bookFold = book.FOLDER_NAME
bookTitle = book.title
bookSpine = book.spine
bookDirection = book.direction
bookPages = book.pages
bookItems = book.items



print(book)







