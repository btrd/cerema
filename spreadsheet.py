#!/usr/bin/python
# -*- coding: utf8 -*-
import ezodf

class Spreadsheet(object):
    def __init__(self, pathToBruit, pathToRes):
        self.openSheet(pathToBruit)
        print(self.getData())
        self.addData()
        self.saveSheet(pathToRes)

    def openSheet(self, pathToBruit):
        try:
            self.bruits_doc = ezodf.opendoc(pathToBruit)
        except:
            print("Fichier " + pathToBruit + " introuvable")
            exit(0)
        self.data = self.bruits_doc.sheets[0]

    def saveSheet(self, pathToRes):
        res_sheet = ezodf.newdoc(doctype="ods", filename=pathToRes)
        res_sheet.sheets += self.data
        try:
            res_sheet.save()
        except PermissionError:
            print("Le fichier sorties.ods est utilisé par un autre logiciel, impossible de le sauvegarder.")
            exit(0)
        
    def getData(self):
        self.startPeriod = self.data[2,1].value
        self.endPeriod = self.data[3,1].value
        self.lieu = self.data[4,1].value
        self.weighting = self.data[5,1].value
        self.dataType = self.data[6,1].value
        self.unit = self.data[7,1].value
        self.data.delete_rows(0, 8)
        return self.lieu, self.dataType, self.weighting, self.unit, self.startPeriod, self.endPeriod

    def addData(self):
        self.data.append_columns(1)
        self.data[0, self.data.ncols()-1].set_value("Laeq Gauss")
        for x in range(2, self.data.nrows()):
            form = "=(E" + str(x) + "+D" + str(x) + ")/2+0.0175*(E" + str(x) + "-D" + str(x) + ")^2"
            self.data[x-1, self.data.ncols()-1].formula = form
