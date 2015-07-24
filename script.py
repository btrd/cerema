#!/usr/bin/python
# -*- coding: utf8 -*-
import sys, glob, imp, os
from sys import exit
from subprocess import call, PIPE
from shutil import copyfile
import argparse

def check_requirements():
    # check Python version
    if sys.version_info[0] != 2:
        print("Cerema Script est uniquement compatible Python 2, vous utilisez Python " + str(sys.version_info[0]))
        exit(1)

    #vérifie si lxml, ezodf et lpod sont installés
    try:
        imp.find_module('lxml')
    except ImportError:
        print("Vous avez besoin de la librairie lxml")
        exit(1)
    try:
        imp.find_module('ezodf')
    except ImportError:
        print("Vous avez besoin de la librairie ezodf")
        exit(1)
    try:
        imp.find_module('lpod')
    except ImportError:
        print("Vous avez besoin de la librairie lpod")
        exit(1)

    #vérifie si tous les fichiers nécessaires sont présent
    fileNeeded(pathToBruit)
    fileNeeded(pathToTrafic)
    fileNeeded(pathToGraph2)
    fileNeeded(pathToParam)
    fileNeeded(pathToPic1)
    fileNeeded(pathToPic2)

def fileNeeded(path):
    if not os.path.isfile(path):
        print("Vous avez besoin de du fichier " + path)
        exit(1)


def convert_file():
    #converti ODS en CSV
    if pathData == "":
        cmdConversion = "soffice --headless --convert-to csv " + pathToSortie
    else:
        cmdConversion = "soffice --headless --convert-to csv --outdir data " + pathToSortie

    return_code = call(cmdConversion, shell=True, stdout=stdout, stderr=stderr)
    if return_code == 1:
        print("Erreur pendant la conversion de " + pathToSortie + ", quittez LibreOffice et/ou OpenOffice")
        exit(1)

    try:
        copyfile(pathToSortieCsv, pathToSortieCsv2)
    except Exception, e:
        print("Erreur pendant la conversion de " + pathToSortieCsv + ", quittez LibreOffice et/ou OpenOffice")
        exit(1)

    if pathData == "":
        cmdConversion = "soffice --headless --convert-to ods " + pathToSortieCsv2
    else:
        cmdConversion = "soffice --headless --convert-to ods --outdir data " + pathToSortieCsv2
    return_code = call(cmdConversion, shell=True, stdout=stdout, stderr=stderr)
    if return_code == 1:
        print("Erreur pendant la conversion de " + pathToSortieCsv2 + ", quittez LibreOffice et/ou OpenOffice")
        exit(1)

def clean():
    try:
        os.remove(pathToSortieCsv)
        os.remove(pathToSortieCsv2)
        os.remove(pathToData)
        os.remove("pic_merge.jpeg")

        filelist = glob.glob(pathData + "*.bak")
        for f in filelist:
            os.remove(f)
    except Exception, e:
        print("Erreur pendant le nettoyage du dossier, quittez LibreOffice et/ou OpenOffice")
        exit(1)

#si un fichier existe déjà on créé une nouvelle version
def getPathToReport(pathToReport):
    if os.path.isfile(pathToReport):
        notFound = True;
        name, ext = os.path.splitext(pathToReport)
        i = 1
        while notFound:
            if not os.path.isfile(name + str(i) + ext):
                pathToReport = name + str(i) + ext
                notFound = False
            i = i + 1
    
    return pathToReport

if __name__ == '__main__':

    parser = argparse.ArgumentParser()
    parser.add_argument("-d", "--dev", action="store_true", help="Dev environment")
    parser.add_argument("-c", "--clean", action="store_true", help="Only clean dev environment (require option --dev)")
    args = parser.parse_args()

    if args.dev:
        if args.clean:
            clean()
            exit(0)
        pathData = "data/"
        stderr = None
        stdout = None
    else:
        pathData = ""
        stderr = PIPE
        stdout = PIPE
    
    pathToBruit = pathData + "bruit.ods"
    pathToTrafic = pathData + "trafic.ods"
    pathToSortie = pathData + "sortie.ods"

    pathToSortieCsv = pathData + "sortie.csv"
    pathToSortieCsv2 = pathData + "sortie2.csv"
    pathToData = pathData + "sortie2.ods"

    pathToGraph1 = pathData + "graph.png"
    pathToGraph2 = pathData + "laeq.jpg"

    pathToParam = pathData + "param.ods"

    pathToPic1 = pathData + "pic1.jpg"
    pathToPic2 = pathData + "pic2.jpg"

    pathToReport = pathData + "report.odt"
    pathToReport = getPathToReport(pathToReport)

    check_requirements()
    
    from spreadsheet import Spreadsheet
    Spreadsheet(5, pathToBruit, pathToTrafic, pathToSortie)
    
    convert_file()

    from graph import Graph
    Graph(pathToGraph1, pathToData)

    from report import Report
    Report(pathToReport, pathToParam, pathToBruit, pathToData, pathToPic1, pathToPic2, pathToGraph1, pathToGraph2)

    # Clean data directory
    clean()

    print "Rapport créé sous " + pathToReport
