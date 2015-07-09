#!/usr/bin/python
# -*- coding: utf8 -*-
import sys
import imp
from sys import exit

def check_requirements():
    #check Python version
    # if sys.version_info[0] != 2:
    #     print("Cerema Script est uniquement compatible Python 2, vous utilisez Python " + str(sys.version_info[0]))
    #     return False

    #check instalation ezodf
    try:
        imp.find_module('lxml')
    except ImportError:
        print("Vous avez besoin de la librairie lxml, cf. fichier README")
        exit(0)
    try:
        imp.find_module('ezodf')
    except ImportError:
        print("Vous avez besoin de la librairie ezodf, cf. fichier README")
        exit(0)

if __name__ == '__main__':
    check_requirements()

    # from spreadsheet import Spreadsheet
    # pathToBruit = "data/bruit.ods"
    # pathToTrafic = "data/trafic.ods"
    # pathToSortie = "data/sortie.ods"
    # Spreadsheet(5, pathToBruit, pathToTrafic, pathToSortie)

    from report import Report
    pathToReport = "data/report.odt"
    pathToParam = "data/param.ods"
    Report(pathToReport, pathToParam)
