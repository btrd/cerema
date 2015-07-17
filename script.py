#!/usr/bin/python
# -*- coding: utf8 -*-
import sys
import imp
from sys import exit
from subprocess import call

def check_requirements():
    # check Python version
    if sys.version_info[0] != 2:
        print("Cerema Script est uniquement compatible Python 2, vous utilisez Python " + str(sys.version_info[0]))
        return False

    #check instalation ezodf
    try:
        imp.find_module('lxml')
    except ImportError:
        print("Vous avez besoin de la librairie lxml, cf. fichier README")
        exit(1)
    try:
        imp.find_module('ezodf')
    except ImportError:
        print("Vous avez besoin de la librairie ezodf, cf. fichier README")
        exit(1)

def convert_file():
    #converti ODS en CSV
    cmdConversion = "soffice --headless --convert-to csv --outdir " + pathData + " " + pathToSortie
    return_code = call(cmdConversion, shell=True)
    if return_code == 1:
        print("Erreur pendant la conversion de " + pathToSortie + ", quittez LibreOffice et/ou OpenOffice")
        exit(1)

    cmdConversion = "cp " + pathToSortieCsv + " " + pathToSortieCsv2
    return_code = call(cmdConversion, shell=True)
    if return_code == 1:
        print("Erreur pendant la conversion de " + pathToSortieCsv + ", quittez LibreOffice et/ou OpenOffice")
        exit(1)

    cmdConversion = "soffice --headless --convert-to ods --outdir " + pathData + " " + pathToSortieCsv2
    return_code = call(cmdConversion, shell=True)
    if return_code == 1:
        print("Erreur pendant la conversion de " + pathToSortieCsv2 + ", quittez LibreOffice et/ou OpenOffice")
        exit(1)

if __name__ == '__main__':
    check_requirements()

    pathData = "data"

    pathToBruit = pathData + "/bruit.ods"
    pathToTrafic = pathData + "/trafic.ods"
    pathToSortie = pathData + "/sortie.ods"

    pathToSortieCsv = pathData + "/sortie.csv"
    pathToSortieCsv2 = pathData + "/sortie2.csv"
    pathToData = pathData + "/sortie2.ods"

    pathToImg = pathData + "/graph.png"

    pathToReport = pathData + "/report.odt"
    pathToParam = pathData + "/param.ods"

    from spreadsheet import Spreadsheet
    Spreadsheet(5, pathToBruit, pathToTrafic, pathToSortie)

    convert_file()

    from graph import Graph
    Graph(pathToImg, pathToData)

    from report import Report
    Report(pathToReport, pathToParam, pathToBruit, pathToData)
