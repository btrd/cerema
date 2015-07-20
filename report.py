#!/usr/bin/python
# -*- coding: utf8 -*-
import ezodf
import lpod

class Report(object):
    def __init__(self, pathToReport, pathToParam, pathToBruit, pathToSortie):
        self.param = self.openSheet(pathToParam)
        self.bruit = self.openSheet(pathToBruit)
        try:
            sortieFile = ezodf.opendoc(pathToSortie)
        except:
            print("Fichier " + pathToSortie + " introuvable")
            exit(1)

        self.sortie = sortieFile.sheets[0]

        self.report = self.createDoc(pathToReport)

        self.getDataBruit()
        
        self.formatSortie()
        self.getDataSortie()

        self.saveFile(sortieFile, pathToSortie)

        

        self.saveFile(self.report, pathToReport)

    # Try to create document, quit if error
    def createDoc(self, pathToDoc):
        try:
            doc = ezodf.newdoc(doctype="odt", filename=pathToDoc)
        except PermissionError:
            print("Le fichier " + pathToDoc + " est utilisé par un autre logiciel, impossible de le créer.")
            exit(1)
        return doc

    # Try to save document, quit if error
    def saveFile(self, doc, pathToDoc):
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

        self.roundCell(self.sortie.nrows()-4, 1, 1)
        self.roundCell(self.sortie.nrows()-3, 1, 1)
        self.roundCell(self.sortie.nrows()-3, 7, 1)
        self.roundCell(self.sortie.nrows()-3, 8, 1)
        self.roundCell(self.sortie.nrows()-1, 8, 0)


    def getDataSortie(self):
        self.laeq6_22 = self.sortie[self.sortie.nrows()-4, 1].value
        self.laeq22_6 = self.sortie[self.sortie.nrows()-3, 1].value

    #arrondie la valeur de la cell [i, j] à n décimale après la virgule
    def roundCell(self, i, j, n):
        value = round(self.sortie[i, j].value, n)
        self.sortie[i, j].set_value(value)
