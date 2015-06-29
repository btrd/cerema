#!/usr/bin/python
# -*- coding: utf8 -*-
import sys
import imp
import os

def check_requirements():
  res = True
  #check Python version
  # if sys.version_info[0] != 2:
  #     print("Cerema Script est uniquement compatible Python 2, vous utilisez Python " + str(sys.version_info[0]))
  #     return False

  #check instalation ezodf
  try:
    imp.find_module('ezodf')
  except ImportError:
    print("Vous avez besoin de la librairie ezodf pour exécuter Cerema Script")
    return False
  return res

def createDocument():
  from ezodf import newdoc, Paragraph, Heading, Sheet
  odt = newdoc(doctype='odt', filename='text.odt')
  odt.body += Heading("Chapter 1")
  odt.body += Paragraph("This is a paragraph.")
  odt.save()

  ods = newdoc(doctype='ods', filename='spreadsheet.ods')
  sheet = Sheet('SHEET', size=(10, 10))
  ods.sheets += sheet
  sheet['A1'].set_value("cell with text")
  sheet['B2'].set_value(3.141592)
  sheet['C3'].set_value(100, currency='USD')
  sheet['D4'].formula = "=SUM([.B2];[.C3])"
  pi = sheet[1, 1].value
  ods.save()

if __name__ == '__main__':
  if check_requirements():
    createDocument()
