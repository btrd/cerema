#!/usr/bin/python
# -*- coding: utf8 -*-
import ezodf

class Report(object):
    def __init__(self, pathToReport, pathToParam, pathToBruit, pathSortie):
        self.param = self.openSheet(pathToParam)
        self.bruit = self.openSheet(pathToBruit)
        self.sortie = self.openSheet(pathSortie)

        self.report = self.createDoc(pathToReport)

        self.getDataBruit()
        
        self.formatSortie()
        self.getDataSortie()

        self.report.body.append(ezodf.Heading("A Simple Test Document"))
        self.report.body.append(ezodf.ezlist(['Point 1', 'Point 2', 'Point 3']))
        self.report.body.append(ezodf.SoftPageBreak())
        self.report.body.append(ezodf.Heading("A Simple Test Document"))
        self.report.body.append(self.sortie)

        self.saveDoc(self.report, pathToReport)

    # Try to create document, quit if error
    def createDoc(self, pathToDoc):
        try:
            doc = ezodf.newdoc(doctype="odt", filename=pathToDoc)
        except PermissionError:
            print("Le fichier " + pathToDoc + " est utilisé par un autre logiciel, impossible de le créer.")
            exit(1)
        return doc

    # Try to save document, quit if error
    def saveDoc(self, doc, pathToDoc):
        try:
            doc.save()
        except PermissionError:
            print("Le fichier " + pathToDoc + " est utilisé par un autre logiciel, impossible de le sauvegarder.")
            exit(1)

    # Try to open document, quit if error
    def openSheet(self, pathToSheet):
        try:
            sheet = ezodf.opendoc(pathToSheet)
        except:
            print("Fichier " + pathToSheet + " introuvable")
            exit(1)
        data = sheet.sheets[0]
        return data

    def getDataBruit(self):
        self.startPeriod = self.bruit[2,1].value
        self.endPeriod = self.bruit[3,1].value
        self.lieu = self.bruit[4,1].value
        self.weighting = self.bruit[5,1].value
        self.dataType = self.bruit[6,1].value
        self.unit = self.bruit[7,1].value

    def formatSortie(self):
        # enlève les colonnes à ne pas afficher
        self.sortie.delete_columns(index=10, count=11)
        self.sortie.delete_columns(index=0, count=1)

        # enlève les infos à ne pas afficher
        self.sortie[self.sortie.nrows()-1, 0].set_value("")
        self.sortie[self.sortie.nrows()-1, 2].set_value("")
        self.sortie[self.sortie.nrows()-1, 3].set_value("")

    def getDataSortie(self):
        self.laeq6_22 = self.sortie[self.sortie.nrows()-4, 1].value
        self.laeq22_6 = self.sortie[self.sortie.nrows()-3, 1].value
