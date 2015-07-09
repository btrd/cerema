#!/usr/bin/python
# -*- coding: utf8 -*-
import ezodf
from datetime import datetime, timedelta
import copy

class Spreadsheet(object):
    def __init__(self, eqVLPL, pathToBruit, pathToTrafic, pathToSortie):
        self.eqVLPL = eqVLPL
        self.bruits = self.openSheet(pathToBruit)
        self.trafic = self.openSheet(pathToTrafic)
        res_sheet = self.createSheet(pathToSortie)

        #supprimes les premières lignes inutiles
        self.deleteData()

        #copie les données de la feuille bruits vers la feuille de résultat
        res_sheet.sheets += copy.deepcopy(self.bruits)
        self.sortie = res_sheet.sheets[0]

        self.checkLine()

        #ajoute les colonnes contenant les formules
        self.addHours()
        self.addGauss()
        self.addD()
        self.addVL()
        self.addPL()
        self.addQeq()
        self.addLaeqCalc()
        self.addM()
        self.addN()
        self.addO()
        self.addSommeLaeqJour()
        self.addSommeLaeqNuit()
        self.addR()
        self.addS()
        self.addT()
        self.addU()

        #ajoute 4 lignes pour y ajouter les infos
        self.sortie.append_rows(4)

        #ajoutes les données de celulle
        self.addSommeU()
        self.addSommeT()
        self.addSommeS()
        self.addSommeR()
        self.addSommeO()
        self.addSommeN()
        self.addSommeIJ()
        self.addSommeLaeq()

        self.addEqVLPL()
        self.addNbrJour()

        self.saveSheet(res_sheet, pathToSortie)

    # Demande à l'utilisateur si la première ligne des deux tableaux coincide
    def checkLine(self):
        print("     _________________________________________________")
        print()
        date = self.getDate(self.sortie[1, 0].value)
        var = "    | " + str(date) + " | "
        for x in range(1, self.sortie.ncols()):
            var = var + str(self.sortie[1, x].value) + " | "
        print(var)
        print("     _________________________________________________")
        print()
        var = "    | "
        for x in range(0, self.trafic.ncols()):
            var = var + str(self.trafic[1, x].value) + " | "
        print(var)
        print("     _________________________________________________")
        print()
        var = input("    Les deux lignes coincide ? (Oui/non): ").capitalize()

        if var != "Oui" and var != "O" and var != "":
            print("Le fichier trafic.ods doit être modifié pour correspondre au fichier bruit.ods")
            exit(0)

    # Try to open document, quit if error
    def openSheet(self, pathToSheet):
        try:
            sheet = ezodf.opendoc(pathToSheet)
        except:
            print("Fichier " + pathToSheet + " introuvable")
            exit(0)
        data = sheet.sheets[0]
        return data

    # Try to create document, quit if error
    def createSheet(self, pathToSheet):
        try:
            sheet = ezodf.newdoc(doctype="ods", filename=pathToSheet)
        except PermissionError:
            print("Le fichier " + pathToSheet + " est utilisé par un autre logiciel, impossible de le sauvegarder.")
            exit(0)
        return sheet

    # Try to save document, quit if error
    def saveSheet(self, sheet, pathToSheet):
        try:
            sheet.save()
        except PermissionError:
            print("Le fichier " + pathToSheet + " est utilisé par un autre logiciel, impossible de le sauvegarder.")
            exit(0)
    
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
                exit(0)
        return time

    # Add a column with a header and a formula
    def addColumnFormula(self, name, data):
        self.sortie.append_columns(1)
        self.sortie[0, self.sortie.ncols()-1].set_value(name)
        for x in range(1, self.sortie.nrows()):
            to = {'x': str(x+1)}
            form = data.format(**to)
            self.sortie[x, self.sortie.ncols()-1].formula = form

    # Delete first 8 row of bruits
    def deleteData(self):
        # self.startPeriod = self.bruits[2,1].value
        # self.endPeriod = self.bruits[3,1].value
        # self.lieu = self.bruits[4,1].value
        # self.weighting = self.bruits[5,1].value
        # self.bruitsType = self.bruits[6,1].value
        # self.unit = self.bruits[7,1].value

        #suprimme les 8 premières lignes
        self.bruits.delete_rows(0, 8)
        #suprimme la dernière ligne
        self.bruits.delete_rows(self.bruits.nrows()-1, 1)

    # add first column hours
    def addHours(self):
        self.sortie.insert_columns(0, 1)
        for x in range(1, self.sortie.nrows()):
            cell = self.sortie[x, 1].value
            time = self.getDate(cell)
            self.sortie[x, 0].set_value(time.hour)

    def addGauss(self):
        name = "Laeq Gauss"
        form = "=(F{x}+E{x})/2+0.0175*(F{x}-E{x})^2"
        self.addColumnFormula(name, form)

    def addD(self):
        form = "=C{x}-G{x}"
        self.addColumnFormula("d", form)

    def addVL(self):
        self.sortie.append_columns(1)
        self.sortie[0, self.sortie.ncols()-1].set_value("VL")
        for x in range(1, self.sortie.nrows()):
            self.sortie[x, self.sortie.ncols()-1].set_value(self.trafic[x,1].value)

    def addPL(self):
        self.sortie.append_columns(1)
        self.sortie[0, self.sortie.ncols()-1].set_value("PL")
        for x in range(1, self.sortie.nrows()):
            self.sortie[x, self.sortie.ncols()-1].set_value(self.trafic[x,2].value)

    def addQeq(self):
        form = "=I{x}+B" + str(self.sortie.nrows()+4) + "*J{x}"
        self.addColumnFormula("Qeq", form)

    def addLaeqCalc(self):
        form = "0"
        self.addColumnFormula("Laeq calc", form)

    def addM(self):
        form = "=(C{x}-L{x})"
        self.addColumnFormula("", form)

    def addN(self):
        form = "=IF(A{x}>5;IF(A{x}<22;K{x};0);0)"
        self.addColumnFormula("", form)

    def addO(self):
        form = "=IF(A{x}>5;IF(A{x}<22;0;K{x});K{x})"
        self.addColumnFormula("", form)

    def addSommeLaeqJour(self):
        form = "=IF(A{x}>5;IF(A{x}<22;C{x};0);0)"
        self.addColumnFormula("Somme LAeq JOUR", form)

    def addSommeLaeqNuit(self):
        form = "=IF(A{x}>5;IF(A{x}<22;0;C{x});C{x})"
        self.addColumnFormula("Somme LAeq NUIT", form)

    def addR(self):
        form = "=IF(P{x}=0;0;10^(P{x}/10))"
        self.addColumnFormula("", form)

    def addS(self):
        form = "=IF(Q{x}=0;0;10^(Q{x}/10))"
        self.addColumnFormula("", form)

    def addT(self):
        form = "=IF(P{x}=0;0;1)"
        self.addColumnFormula("", form)

    def addU(self):
        form = "=IF(Q{x}=0;0;1)"
        self.addColumnFormula("", form)

    def addSommeU(self):
        form = "=SUM(U2:U" + str(self.sortie.nrows()-4) + ")"
        self.sortie[self.sortie.nrows()-4, 20].formula = form

    def addSommeT(self):
        form = "=SUM(T2:T" + str(self.sortie.nrows()-4) + ")"
        self.sortie[self.sortie.nrows()-4, 19].formula = form

    def addSommeS(self):
        form = "=SUM(S2:S" + str(self.sortie.nrows()-4) + ")/U" + str(self.sortie.nrows()-3)
        self.sortie[self.sortie.nrows()-4, 18].formula = form

    def addSommeR(self):
        form = "=SUM(R2:R" + str(self.sortie.nrows()-4) + ")/T" + str(self.sortie.nrows()-3)
        self.sortie[self.sortie.nrows()-4, 17].formula = form

    def addSommeO(self):
        form = "=SUM(O2:O" + str(self.sortie.nrows()-4) + ")/U" + str(self.sortie.nrows()-3)
        self.sortie[self.sortie.nrows()-4, 14].formula = form

    def addSommeN(self):
        form = "=SUM(N2:N" + str(self.sortie.nrows()-4) + ")/T" + str(self.sortie.nrows()-3)
        self.sortie[self.sortie.nrows()-4, 13].formula = form

    def addSommeIJ(self):
        self.sortie[self.sortie.nrows()-4, 7].set_value("Total")
        form = "=SUM(I2:I" + str(self.sortie.nrows()-4) + ")"
        self.sortie[self.sortie.nrows()-4, 8].formula = form
        form = "=SUM(J2:J" + str(self.sortie.nrows()-4) + ")"
        self.sortie[self.sortie.nrows()-4, 9].formula = form

        self.sortie[self.sortie.nrows()-3, 7].set_value("V/jour")
        form = "=I" + str(self.sortie.nrows()-3) + "/E" + str(self.sortie.nrows())
        self.sortie[self.sortie.nrows()-3, 8].formula = form
        form = "=J" + str(self.sortie.nrows()-3) + "/E" + str(self.sortie.nrows())
        self.sortie[self.sortie.nrows()-3, 9].formula = form

        self.sortie[self.sortie.nrows()-2, 8].set_value("TV/j")
        form = "=I" + str(self.sortie.nrows()-2) + "+J" + str(self.sortie.nrows()-2)
        self.sortie[self.sortie.nrows()-2, 9].formula = form

        self.sortie[self.sortie.nrows()-1, 8].set_value("%PL")
        form = "=100*J" + str(self.sortie.nrows()-2) + "/J" + str(self.sortie.nrows()-1)
        self.sortie[self.sortie.nrows()-1, 9].formula = form

    def addSommeLaeq(self):
        self.sortie[self.sortie.nrows()-4, 1].set_value("(6h - 22h)")
        form = "=10*LOG10(R" + str(self.sortie.nrows()-3) + ")"
        self.sortie[self.sortie.nrows()-4, 2].formula = form

        self.sortie[self.sortie.nrows()-3, 1].set_value("(22h - 6h)")
        form = "=10*LOG10(S" + str(self.sortie.nrows()-3) + ")"
        self.sortie[self.sortie.nrows()-3, 2].formula = form

    def addEqVLPL(self):
        self.sortie[self.sortie.nrows()-1, 0].set_value("Equivalence VL-PL")
        self.sortie[self.sortie.nrows()-1, 1].set_value(self.eqVLPL)

    def addNbrJour(self):
        self.sortie[self.sortie.nrows()-1, 3].set_value("Nbr jour")

        cell1 = self.sortie[1, 1].value
        d1 = self.getDate(cell1)

        cell2 = self.sortie[self.sortie.nrows()-5, 1].value
        #ajoute 1h car heure de début de la période et non de fin
        d2 = self.getDate(cell2) + timedelta(hours=1)

        nbrDays = (d2-d1).days
        self.sortie[self.sortie.nrows()-1, 4].set_value(nbrDays)
