#!/usr/bin/python
# -*- coding: utf8 -*-
import ezodf

class Spreadsheet(object):
    def __init__(self, eqVLPL, pathToBruit, pathToTrafic, pathToRes):
        self.eqVLPL = eqVLPL
        self.bruits = self.openSheet(pathToBruit)
        self.trafics = self.openSheet(pathToTrafic)
        res_sheet = self.createSheet(pathToRes)

        #get first data and delete lines
        self.getData()

        #copy data from bruits in res_sheet
        res_sheet.sheets += self.bruits
        self.data = res_sheet.sheets[0]

        #add column data
        self.addGauss()
        self.addD()
        self.addVL()
        self.addPL()
        self.addQeq()
        self.addLaeqCalc()
        self.addTmpM()
        self.addTmpN()

        self.addEqVLPL()
        self.saveSheet(res_sheet)

    def openSheet(self, pathToSheet):
        try:
            sheet = ezodf.opendoc(pathToSheet)
        except:
            print("Fichier " + pathToSheet + " introuvable")
            exit(0)
        data = sheet.sheets[0]
        return data

    def createSheet(self, pathToSheet):
        try:
            sheet = ezodf.newdoc(doctype="ods", filename=pathToSheet)
        except PermissionError:
            print("Le fichier sorties.ods est utilisé par un autre logiciel, impossible de le sauvegarder.")
            exit(0)
        return sheet

    def saveSheet(self, sheet):
        try:
            sheet.save()
        except PermissionError:
            print("Le fichier sorties.ods est utilisé par un autre logiciel, impossible de le sauvegarder.")
            exit(0)
    
    def addColumnFormula(self, name, data):
        self.data.append_columns(1)
        self.data[0, self.data.ncols()-1].set_value(name)
        for x in range(1, self.data.nrows()):
            to = {'x': str(x+1)}
            form = data.format(**to)
            self.data[x, self.data.ncols()-1].formula = form

    def getData(self):
        self.startPeriod = self.bruits[2,1].value
        self.endPeriod = self.bruits[3,1].value
        self.lieu = self.bruits[4,1].value
        self.weighting = self.bruits[5,1].value
        self.bruitsType = self.bruits[6,1].value
        self.unit = self.bruits[7,1].value
        self.bruits.delete_rows(0, 8)
        self.bruits.delete_rows(self.bruits.nrows()-1, 1)
        return self.lieu, self.bruitsType, self.weighting, self.unit, self.startPeriod, self.endPeriod

    def addGauss(self):
        form = "=(E{x}+D{x})/2+0.0175*(E{x}-D{x})^2"
        self.addColumnFormula("Laeq Gauss", form)

    def addD(self):
        form = "=B{x}-F{x}"
        self.addColumnFormula("d", form)

    def addVL(self):
        self.data.append_columns(1)
        self.data[0, self.data.ncols()-1].set_value("VL")
        for x in range(1, self.data.nrows()):
            self.data[x, self.data.ncols()-1].set_value(self.trafics[x,1].value)

    def addPL(self):
        self.data.append_columns(1)
        self.data[0, self.data.ncols()-1].set_value("PL")
        for x in range(1, self.data.nrows()):
            self.data[x, self.data.ncols()-1].set_value(self.trafics[x,2].value)

    def addQeq(self):
        form = "=H{x}+B" + str(self.data.nrows()+2) + "*I{x}"
        self.addColumnFormula("Qeq", form)

    def addLaeqCalc(self):
        form = "0"
        self.addColumnFormula("Laeq calc", form)

    def addTmpM(self):
        form = "=(B{x}-K{x})"
        self.addColumnFormula("", form)

    def addTmpN(self):
        form = "=SI(A4>5;SI(A4<22;K4;0);0)"
        self.addColumnFormula("", form)

    def addEqVLPL(self):
        self.data.append_rows(2)
        self.data[self.data.nrows()-1, 0].set_value("Equivalence VL-PL")
        self.data[self.data.nrows()-1, 1].set_value(self.eqVLPL)
