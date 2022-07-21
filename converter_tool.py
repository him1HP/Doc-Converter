import sys
import os
import win32com.client
from pdf2docx import Converter

print('\t\t\t\t \033[1;31m ******* Welcome to the Python Document Converter tool!!*******\033[1;m \n')
print (40 * '-')
print('Select from the below choices:')
print (40 * '-')
print ('1. Doc/Docx to Pdf format')
print ('2. Html to Doc format')
print ('3. pdf to Doc/Docx format')

choice = input('Enter your choice [1-3] : ')
choice = int(choice)


if choice == 1:
    word_doc = input('Please Enter the absolute path of Word document you are looking for conversion :')
    pdf_doc = input('Please Enter the absolute path for Pdf document you want to generate :')
    pdf = pdf_doc.split('.')[1]
    if 'pdf' not in pdf_doc or 'doc' not in word_doc:
        print('This is the wrong extension of filename! Please enter extension as .doc/docx as input and .pdf as output file')
        exit()
    else:
        wdFormatPDF = 17
        word = win32com.client.Dispatch('Word.Application')
        doc = word.Documents.Open(word_doc)
        doc.SaveAs(pdf_doc, FileFormat=wdFormatPDF)
        print('The Word document is successfully converted to pdf document! Thanks for using the converter!!')
        doc.Close()
        word.Quit()

elif choice == 2:
    html_doc = input('Please Enter the absolute path of Html document you are looking for conversion :')
    word_doc = input('Please Enter the absolute path for doc document you want to generate :')

    if 'doc' not in word_doc or 'html' not in html_doc:
        print('This is the wrong extension of filename! Please enter extension as .html for source and .doc for output file')
        exit()
    else:
        word = win32com.client.Dispatch('Word.Application')
        doc = word.Documents.Add(html_doc)
        doc.SaveAs(word_doc, FileFormat=0)
        print('The Html document is successfully converted to Word document! Thanks for using the converter!!')
        doc.Close()
        word.Quit()

elif choice == 3:
    pdf_doc = input('Please Enter the absolute path of pdf document you are looking for conversion :')
    word_doc = input('Please Enter the absolute path for word document you want to generate :')

    if 'pdf' not in pdf_doc or 'doc' not in word_doc:
        print('This is the wrong extension of filename! Please enter extension as .pdf for input and .doc/.docx for output file')
        exit()
    else:
        pages_list = [0]
        cv = Converter(pdf_doc)
        cv.convert(word_doc, pages=pages_list)
        print('The pdf document is successfully converted to Word document! Thanks for using the converter!!')
        cv.close()
else:
    print('Sorry! Invalid choice!')
