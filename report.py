#!/usr/bin/python
# -*- coding: utf8 -*-
import ezodf
from datetime import datetime, timedelta
from lpod.document import odf_new_document
from lpod.paragraph import odf_create_paragraph
from lpod.style import odf_create_style
from PIL import Image
import glob, os

class Report(object):
    def __init__(self, pathToReport, pathToParam, pathToBruit, pathToSortie, pathToPic1, pathToPic2):
        self.param = self.openSheet(pathToParam)
        self.bruit = self.openSheet(pathToBruit)

        sortieFile = self.openDoc(pathToSortie)
        self.sortie = sortieFile.sheets[0]

        self.reportFile = odf_new_document("text")
        self.report = self.reportFile.get_body()
        
        self.getDataBruit()
        
        self.formatSortie()
        self.getDataSortie()

        self.saveFileEzodf(sortieFile, pathToSortie)

        self.formatImages(pathToPic1, pathToPic2)

        self.addStyle()
        self.addContent()
        self.saveFileLpod(self.reportFile, pathToReport)

    # Try to save document, quit if error
    def saveFileEzodf(self, doc, pathToDoc):
        try:
            doc.save()
        except PermissionError:
            print("Le fichier " + pathToDoc + " est utilisé par un autre logiciel, impossible de le sauvegarder.")
            exit(1)

    def saveFileLpod(self, doc, pathToDoc):
        try:
            doc.save(target=pathToDoc)
        except PermissionError:
            print("Le fichier " + pathToDoc + " est utilisé par un autre logiciel, impossible de le sauvegarder.")
            exit(1)

    # Try to open sheet document, quit if error, returns sheet
    def openSheet(self, pathToSheet):
        try:
            sheet = ezodf.opendoc(pathToSheet)
        except:
            print("Fichier " + pathToSheet + " introuvable")
            exit(1)
        data = sheet.sheets[0]
        return data    # Try to open document, quit if error

    # Try to open sheet document, quit if error, returns doc
    def openDoc(self, pathToDoc):
        try:
            doc = ezodf.opendoc(pathToDoc)
        except:
            print("Fichier " + pathToDoc + " introuvable")
            exit(1)
        return doc

    def getDataBruit(self):
        self.startPeriod = self.getDate(self.bruit[2,1].value)
        self.endPeriod = self.getDate(self.bruit[3,1].value)
        self.lieu = str(self.bruit[4,1].value)
        self.weighting = str(self.bruit[5,1].value)
        self.dataType = str(self.bruit[6,1].value)
        self.unit = str(self.bruit[7,1].value)

    def formatSortie(self):
        # enlève les colonnes à ne pas afficher
        self.sortie.delete_columns(index=10, count=11)
        self.sortie.delete_columns(index=0, count=1)

        # enlève les infos à ne pas afficher
        emtptyCell1 = ezodf.Cell("")
        emtptyCell2 = ezodf.Cell("")
        emtptyCell3 = ezodf.Cell("")
        self.sortie[self.sortie.nrows()-1, 0] = emtptyCell1
        self.sortie[self.sortie.nrows()-1, 2] = emtptyCell2
        self.sortie[self.sortie.nrows()-1, 3] = emtptyCell3

        # change le nombre de décimale après la virgule
        for j in [5, 6]:
            for i in range(1, self.sortie.nrows()-4):
                self.roundCell(i, j, 1)

        # (6h - 22h)
        self.roundCell(self.sortie.nrows()-4, 1, 1)
        # (22h - 6h)
        self.roundCell(self.sortie.nrows()-3, 1, 1)
        # V/jour VL
        self.roundCell(self.sortie.nrows()-3, 7, 0)
        # V/jour PL
        self.roundCell(self.sortie.nrows()-3, 8, 0)
        # % PL
        self.roundCell(self.sortie.nrows()-1, 8, 1)


    def getDataSortie(self):
        self.laeq6_22 = self.sortie[self.sortie.nrows()-4, 1].value
        self.laeq22_6 = self.sortie[self.sortie.nrows()-3, 1].value

        self.lden = self.laeq6_22
        self.lnight = self.laeq22_6 - 3

        self.pourcPL = self.sortie[self.sortie.nrows()-1, 8].value
        self.nbrVehicule = self.sortie[self.sortie.nrows()-2, 8].value

    #arrondie la valeur de la cell [i, j] à n décimale après la virgule
    def roundCell(self, i, j, n):
        value = round(self.sortie[i, j].value, n)
        self.sortie[i, j].set_value(value)

    # Take a date string (format ISO 8601) and return a date
    def getDate(self, dateString):
        try:
            time = datetime.strptime(dateString, "%Y-%m-%dT%H:%M:%S")
        except ValueError:
            # If date string is midnight there is no hours
            try:
                time = datetime.strptime(dateString, "%Y-%m-%d")
            except ValueError:
                print("La première colonne doit uniquement contenir des dates")
                exit(1)
        return time

    def addStyle(self):
        self.reportFile.delete_styles()
        _style_bold = odf_create_style('text', name = u'bold', bold = True)
        self.reportFile.insert_style(_style_bold)

    def addText(self, text, style="", regex=".*"):
        paragraph = odf_create_paragraph()
        paragraph.append_plain_text(text)
        paragraph.set_span(style, regex=regex)
        self.report.append(paragraph)

    def formatImages(self, pathToPic1, pathToPic2):
        filename1, ext1 = os.path.splitext(pathToPic1)
        filename2, ext2 = os.path.splitext(pathToPic2)

        size = 595, 595

        im = Image.open(pathToPic1)
        im.thumbnail(size, Image.ANTIALIAS)
        im.save(filename1 + "_small" + ext1, "JPEG")

        im = Image.open(pathToPic2)
        im.thumbnail(size, Image.ANTIALIAS)
        im.save(filename2 + "_small" + ext2, "JPEG")

    def addContent(self):
        text = "Traffic du " + str(self.startPeriod.day) + " au " + self.endPeriod.strftime('%d/%m/%Y')
        self.addText(text, style="bold", regex = "Traffic")

        text = "Véh/j " + str(self.nbrVehicule) + " PL " + str(self.pourcPL) + "%"
        self.addText(text, style="bold")

        text = "Résultats des mesures"
        self.addText(text, style="bold")

        self.addText("Lieu \t\t\t" + self.lieu)
        self.addText("Type de données \t" + self.dataType)
        self.addText("Pondération \t\t" + self.weighting)
        self.addText("Unité \t\t\t" + self.unit)
        self.addText("Début \t\t\t" + self.endPeriod.strftime('%d/%m/%Y %H:%M'))
        self.addText("Fin \t\t\t" + self.startPeriod.strftime('%d/%m/%Y %H:%M'))

        self.addText("Lden \t\t\t" + str(self.lden))
        self.addText("Lnight \t\t\t" + str(self.lnight))

        self.addText("LAeq(6h-22h) \t" + str(self.laeq6_22))
        self.addText("LAeq(22h-6h) \t" + str(self.laeq22_6))
        