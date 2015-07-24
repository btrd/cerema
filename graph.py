#!/usr/bin/python
# -*- coding: utf8 -*-
import csv
import ezodf
import matplotlib.pyplot as mpl

#create Graph from Spreadsheet
class Graph(object):
    def __init__(self, pathToImg, pathToData):

        self.data = self.openSheet(pathToData)

        hours = []
        laeq = []
        laeqCalc = []

        #cherche les infos dans le fichier ODS
        for i in range(1, self.data.nrows()-4):
            hours.append(int(self.data[i, 0].value))
            laeq.append(self.data[i, 2].value)
            laeqCalc.append(self.data[i, 11].value)

        #initialise le graph
        mpl.figure(figsize=(20, 7))

        #ajoute la courbe pour Laeq.mes
        mpl.plot(laeq, color="blue", linewidth=1.0, linestyle="-", label="Laeq.mes", marker='^', markersize=7)

        #ajoute la courbe pour Laeq.calc
        mpl.plot(laeqCalc, color="red", linewidth=1.0, linestyle="-", label="Laeq.calc", marker='s', markersize=5)

        #garde juste une heure sur deux pour le label
        for i in range(0, len(hours)):
            if i % 2 == 0:
                hours[i] = ''

        #affiche les heures comme label
        mpl.xticks(list(range(len(hours))), hours)
        mpl.yticks([30, 35, 40, 45, 50, 55, 60, 65, 70, 75, 80])

        #limite du graph
        mpl.xlim(-1, len(laeq))
        mpl.ylim(30, 80)

        #affiche des ligne horizontal
        mpl.axes().yaxis.grid(True)

        #affiche la légende en haut à gauche
        mpl.legend(loc='upper left')

        #enregistre le graph
        mpl.savefig(pathToImg, bbox_inches='tight')

    # Try to open document, quit if error
    def openSheet(self, pathToSheet):
        try:
            sheet = ezodf.opendoc(pathToSheet)
        except:
            print("Fichier " + pathToSheet + " introuvable")
            exit(1)
        data = sheet.sheets[0]
        return data
