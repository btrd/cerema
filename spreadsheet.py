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

        #ajoute 4 lignes pour y ajouter les infos
        self.sortie.append_rows(4)

        #ajoute les colonnes contenant les formules
        self.addHours()
        self.addGauss()
        self.addD()
        self.addVL()
        self.addPL()
        self.addQeq()
        self.addLaeqCalc()
        self.addLea()
        self.addQeqJour()
        self.addQeqNuit()
        self.addSommeLaeqJour()
        self.addSommeLaeqNuit()
        self.addPuissanceAcoustiqueJour()
        self.addPuissanceAcoustiqueNuit()
        self.addNbJour()
        self.addNbNuit()
        self.addPuissanceAcoustiqueSoir()
        self.addNbSoir()

        #ajoutes les données de celulle
        self.addSommeIJ()
        self.addSommeLaeq()

        self.addEqVLPL()
        self.addNbrJour()

        #enregistre le fichier
        self.saveSheet(res_sheet, pathToSortie)

    # Demande à l'utilisateur si la première ligne des deux tableaux coincide
    def checkLine(self):
        # Affiche les premières lignes
        print("     _________________________________________________")
        print("")
        date = self.getDate(self.sortie[1, 0].value)
        var = "    | " + str(date) + " | "
        for x in range(1, self.sortie.ncols()):
            var = var + str(self.sortie[1, x].value) + " | "
        print(var)
        print("     _________________________________________________")
        print("")
        var = "    | "
        for x in range(0, self.trafic.ncols()):
            var = var + str(self.trafic[1, x].value) + " | "
        print(var)
        print("     _________________________________________________")
        print("")

        var = raw_input("    Les deux lignes coincident ? (Oui/non): ").capitalize()
        print("")

        if var != "Oui" and var != "O" and var != "":
            print("Le fichier trafic.ods doit etre modifie pour correspondre au fichier bruit.ods")
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

    # Try to create document, quit if error
    def createSheet(self, pathToSheet):
        try:
            sheet = ezodf.newdoc(doctype="ods", filename=pathToSheet)
        except PermissionError:
            print("Le fichier " + pathToSheet + " est utilise par un autre logiciel, impossible de le sauvegarder.")
            exit(1)
        return sheet

    # Try to save document, quit if error
    def saveSheet(self, sheet, pathToSheet):
        try:
            sheet.save()
        except PermissionError:
            print("Le fichier " + pathToSheet + " est utilise par un autre logiciel, impossible de le sauvegarder.")
            exit(1)
    
    # Take a date string (format ISO 8601) and return a date
    def getDate(self, dateString):
        try:
            time = datetime.strptime(dateString, "%Y-%m-%dT%H:%M:%S")
        except ValueError:
            # If date string is midnight there is no hours
            try:
                time = datetime.strptime(dateString, "%Y-%m-%d")
            except ValueError:
                print("La premiere colonne doit uniquement contenir des dates")
                exit(1)
        return time

    # Add a column with a header and a formula
    def addColumnFormula(self, name, data):
        self.sortie.append_columns(1)
        self.sortie[0, self.sortie.ncols()-1].set_value(name)
        for i in range(1, self.sortie.nrows()-4):
            #si la formule comporte un {x} on remplace par i
            to = {'x': str(i+1)}
            form = data.format(**to)

            self.sortie[i, self.sortie.ncols()-1].formula = form

    # Delete first 8 row of bruits
    def deleteData(self):
        #suprimme les 8 premières lignes
        self.bruits.delete_rows(0, 8)
        #suprimme la dernière ligne
        self.bruits.delete_rows(self.bruits.nrows()-1, 1)

    # add first column hours
    def addHours(self):
        self.sortie.insert_columns(0, 1)
        for x in range(1, self.sortie.nrows()-4):
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

    #copie la colonne VL de traffic dans sortie
    def addVL(self):
        self.sortie.append_columns(1)
        self.sortie[0, self.sortie.ncols()-1].set_value("VL")
        for x in range(1, self.sortie.nrows()-4):
            self.sortie[x, self.sortie.ncols()-1].set_value(self.trafic[x,1].value)

    #copie la colonne PL de traffic dans sortie
    def addPL(self):
        self.sortie.append_columns(1)
        self.sortie[0, self.sortie.ncols()-1].set_value("PL")
        for x in range(1, self.sortie.nrows()-4):
            self.sortie[x, self.sortie.ncols()-1].set_value(self.trafic[x,2].value)

    def addQeq(self):
        form = "=I{x}+B" + str(self.sortie.nrows()) + "*J{x}"
        self.addColumnFormula("Qeq", form)

    def addLaeqCalc(self):
        lastR1 = str(self.sortie.nrows()-3)
        lastR2 = str(self.sortie.nrows()-2)
        form = "=IF(Q{x}=0;C" + lastR1 + "+10*LOG10(K{x}/N" + lastR1 + ");C" + lastR2 + "+10*LOG10(K{x}/O" + lastR1 + "))"
        self.addColumnFormula("Laeq calc", form)

    def addLea(self):
        form = "=(C{x}-L{x})"
        self.addColumnFormula("Lea mes-LAeq calc", form)

    def addQeqJour(self):
        form = "=IF(A{x}>5;IF(A{x}<22;K{x};0);0)"
        self.addColumnFormula("Qeq Jour", form)

        form = "=SUM(N2:N" + str(self.sortie.nrows()-4) + ")/T" + str(self.sortie.nrows()-3)
        self.sortie[self.sortie.nrows()-4, 13].formula = form

    def addQeqNuit(self):
        form = "=IF(A{x}>5;IF(A{x}<22;0;K{x});K{x})"
        self.addColumnFormula("Qeq Nuit", form)

        form = "=SUM(O2:O" + str(self.sortie.nrows()-4) + ")/U" + str(self.sortie.nrows()-3)
        self.sortie[self.sortie.nrows()-4, 14].formula = form

    def addSommeLaeqJour(self):
        form = "=IF(A{x}>5;IF(A{x}<22;C{x};0);0)"
        self.addColumnFormula("Somme LAeq Jour", form)

    def addSommeLaeqNuit(self):
        form = "=IF(A{x}>5;IF(A{x}<22;0;C{x});C{x})"
        self.addColumnFormula("Somme LAeq Nuit", form)

    def addPuissanceAcoustiqueJour(self):
        form = "=IF(P{x}=0;0;10^(P{x}/10))"
        self.addColumnFormula("Puissance acoustique Jour", form)

        form = "=SUM(R2:R" + str(self.sortie.nrows()-4) + ")/T" + str(self.sortie.nrows()-3)
        self.sortie[self.sortie.nrows()-4, 17].formula = form

    def addPuissanceAcoustiqueNuit(self):
        form = "=IF(Q{x}=0;0;10^(Q{x}/10))"
        self.addColumnFormula("Puissance acoustique Nuit", form)

        form = "=SUM(S2:S" + str(self.sortie.nrows()-4) + ")/U" + str(self.sortie.nrows()-3)
        self.sortie[self.sortie.nrows()-4, 18].formula = form

    def addNbJour(self):
        form = "=IF(P{x}=0;0;1)"
        self.addColumnFormula("Nb Nuit", form)

        form = "=SUM(T2:T" + str(self.sortie.nrows()-4) + ")"
        self.sortie[self.sortie.nrows()-4, 19].formula = form

    def addNbNuit(self):
        form = "=IF(Q{x}=0;0;1)"
        self.addColumnFormula("Nb Nuit", form)

        form = "=SUM(U2:U" + str(self.sortie.nrows()-4) + ")"
        self.sortie[self.sortie.nrows()-4, 20].formula = form

    def addPuissanceAcoustiqueSoir(self):
        form = "=IF(A{x}>17;IF(A{x}<22;10^((C{x})/10);0);0)"
        self.addColumnFormula("Puissance acoustique Soir (18h-22h)", form)

        form = "=SUM(V2:V" + str(self.sortie.nrows()-4) + ")/W" + str(self.sortie.nrows()-3)
        self.sortie[self.sortie.nrows()-4, 21].formula = form

    def addNbSoir(self):
        form = "=IF(V{x}=0;0;1)"
        self.addColumnFormula("Nb Soir (18h-22h)", form)

        form = "=SUM(W2:W" + str(self.sortie.nrows()-4) + ")"
        self.sortie[self.sortie.nrows()-4, 22].formula = form

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

        self.sortie[self.sortie.nrows()-2, 1].set_value("Lden")
        lastRow = str(self.sortie.nrows()-3)
        form = "=10*LOG10((16*R" + lastRow + "-4*V" + lastRow + ")/24+4/24*V" + lastRow + "*10^(5/10)+8/24*S" + lastRow + "*10^(10/10))-3"
        self.sortie[self.sortie.nrows()-2, 2].formula = form

    def addEqVLPL(self):
        self.sortie[self.sortie.nrows()-1, 0].set_value("Equivalence VL-PL")
        self.sortie[self.sortie.nrows()-1, 1].set_value(self.eqVLPL)

    def addNbrJour(self):
        self.sortie[self.sortie.nrows()-1, 3].set_value("Nbr jour")

        #date de début
        cell1 = self.sortie[1, 1].value
        d1 = self.getDate(cell1)

        #date de fin
        cell2 = self.sortie[self.sortie.nrows()-5, 1].value
        #ajoute 1h car cell2 est l'heure de début de la période et non l'heure de fin
        d2 = self.getDate(cell2) + timedelta(hours=1)

        nbrDays = (d2-d1).days
        self.sortie[self.sortie.nrows()-1, 4].set_value(nbrDays)
