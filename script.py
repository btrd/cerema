#!/usr/bin/python
# -*- coding: utf8 -*-
import sys, glob, imp, os, locale
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


# converti pathToSortie en pathToData (pathToData équivalent à pathToSortie mais les formules sont remplacé par les valeurs)
# on utilise LibreOffice en ligne de commande pour faire la conversion vers du CSV puis de nouveau vers de l'ods
# soffice ne permet pas de préciser un nom de fichier de sortie... donc on copie le fichier CSV
def convert_file():

    #si on utilise le dossier local
    if pathData == "":
        cmdConversion = "soffice --headless --convert-to csv " + pathToSortie
    #dans le cas contraire
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

# supprime les fichiers temporaires
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

#si un fichier report existe déjà on créé une nouvelle version
def getPathToReport(pathToReport):
    # si le fichier existe on cherche un nom qui n'existe pas encore en concatenant 1 ou 2 ou 3, etc
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
    # set locale to fr
    try:
        #UNIX
        locale.setlocale(locale.LC_ALL, 'fr_fr')
    except Exception, e:
        try:
            #WINDOWS
            locale.setlocale(locale.LC_ALL, 'fra_fra')
        except Exception, e:
            raise e

    # on parse les arguments passé en ligne de commande
    parser = argparse.ArgumentParser()
    parser.add_argument("-d", "--dev", action="store_true", help="Dev environment")
    parser.add_argument("-c", "--clean", action="store_true", help="Only clean dev environment (require option --dev)")
    args = parser.parse_args()

    # si option dev on utilise le dossier data et on envoie les erreurs vers le terminal
    if args.dev:
        if args.clean:
            clean()
            exit(0)
        pathData = "data/"
        stderr = None
        stdout = None
    # sinon on utilise le dossier courant et on cache les erreurs
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

    pathToGraph1 = pathData + "recalage_trafic.png"
    pathToGraph2 = pathData + "laeq_15min.jpg"

    pathToParam = pathData + "param.ods"

    pathToPic1 = pathData + "vue_facade.jpg"
    pathToPic2 = pathData + "vue_dessus.jpg"

    if args.dev:
        pathToReport = pathData + "report.odt"
    else:
        pathToReport = getPathToReport(pathData + "report.odt")

    check_requirements()
    
    #on créé sortie.ods
    from spreadsheet import Spreadsheet
    Spreadsheet(5, pathToBruit, pathToTrafic, pathToSortie)
    
    convert_file()

    # on créé graph.png
    from graph import Graph
    Graph(pathToGraph1, pathToData)

    # on créé le raport
    from report import Report
    Report(pathToReport, pathToParam, pathToBruit, pathToData, pathToPic1, pathToPic2, pathToGraph1, pathToGraph2)

    # Clean data directory
    clean()

    print "Rapport enregistre sous " + pathToReport
