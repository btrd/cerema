#!/usr/bin/python
# -*- coding: utf8 -*-
import sys
import imp
from sys import exit
from subprocess import call

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

    pathData = "data"

    from spreadsheet import Spreadsheet
    pathToBruit = pathData + "/bruit.ods"
    pathToTrafic = pathData + "/trafic.ods"
    pathToSortie = pathData + "/sortie.ods"
    # Spreadsheet(5, pathToBruit, pathToTrafic, pathToSortie)

    #converti ODS en CSV
    # cmdConversion = "soffice --headless --convert-to csv --outdir " + pathData + " " + pathToSortie
    # return_code = call(cmdConversion, shell=True)
    # if return_code == 1:
    #     print("Erreur pendant la conversion de " + pathToSortie + ", quittez LibreOffice et/ou OpenOffice")
    #     exit(0)

    from graph import Graph
    pathToImg = pathData + "/graph.png"
    pathToCsv = pathData + "/sortie.csv"
    Graph(pathToImg, pathToCsv)

    from report import Report
    pathToReport = pathData + "/report.odt"
    pathToParam = pathData + "/param.ods"
    # Report(pathToReport, pathToParam)
