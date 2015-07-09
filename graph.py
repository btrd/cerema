#!/usr/bin/python
# -*- coding: utf8 -*-
import csv
import pandas as pd
import numpy as np
import matplotlib.pyplot as mpl

class Graph(object):
    def __init__(self, pathToImg, pathToData):

        name = np.loadtxt(pathToData, dtype=str, delimiter=',', skiprows=1, usecols=(0,))

        X1 = [59.3, 59.2, 59.8, 59.6, 59.9, 59.6, 59.1, 58.1, 57.1, 56, 54.2, 52.3, 51.5, 50.4, 52.2, 53.1, 54.4, 58.2, 59, 60.6, 60.2, 59.2, 59.3, 59.1, 58.7, 59.3, 59, 59.9, 59.5, 59.5, 58.8, 58.2, 56.8, 55.7, 54.9, 53, 51.5, 51.9, 51.3, 52.4, 54.2, 56.8, 59.2, 60.1, 59.6, 58.7, 58.2, 59, 58.3, 58.4, 58.6, 58.8, 59, 59.5, 58.4, 57.4, 56.9, 55.2, 53.6, 52.7, 51.7, 51.1, 51.9, 52.1, 54.3, 56.6, 59.3, 59.9, 61.2, 60.3, 59.1, 58.8]
        X2 = [58.8445065189, 58.9916657136, 59.4873213067, 59.6831163423, 60.2841930794, 60.3593839031, 59.1765219356, 57.6017177029, 55.7519892395, 53.7749193778, 54.0718954537, 52.9261632423, 52.0389352772, 51.477297757, 52.5082787502, 53.3047551934, 54.4875977137, 56.4262712935, 57.4229950908, 60.1183133251, 60.3133272578, 59.562132563, 59.2613339683, 59.1596660275, 58.7527448756, 59.1668250103, 59.5312351212, 59.6352182457, 60.0073280762, 60.5055939639, 59.2125673915, 57.748657215, 55.905162373, 54.1247162645, 54.0935452733, 52.6999774862, 52.4202855411, 51.9301366264, 52.2288210328, 52.9795204938, 54.3207341344, 56.1942641343, 57.6345609006, 59.9415728326, 60.4264190669, 59.353829997, 58.9590941422, 59.2528267279, 58.7991460919, 58.971403933, 59.2728173032, 59.5439015639, 60.2108926927, 60.2145069611, 58.9338303286, 58.090390425, 55.9084094347, 54.2596410822, 54.3309696914, 53.0951923171, 52.4550672291, 52.0458618255, 52.5516156939, 53.0542337159, 54.4280227765, 56.2300633021, 57.655595664, 59.7826801565, 60.1383769346, 59.2307302912, 59.0910605672, 59.2307302912]

        hours = [11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 11]

        #initialise le graph
        mpl.figure(figsize=(20,10))

        #ajoute la courbe pour Laeq.mes
        mpl.plot(X1, color="blue", linewidth=1.0, linestyle="-", label="Laeq.mes", marker='^', markersize=7)

        #ajoute la courbe pour Laeq.calc
        mpl.plot(X2, color="red", linewidth=1.0, linestyle="-", label="Laeq.calc", marker='s', markersize=5)

        #garde juste une heure sur deux pour le label
        for i in range(0, len(hours)):
            if i % 2 == 0:
                hours[i] = ''

        #affiche les heures comme label
        mpl.xticks(list(range(len(hours))), hours)

        #limite du graph
        mpl.xlim(-1, len(X1))

        #affiche la légende en haut à gauche
        mpl.legend(loc='upper left')

        #enregistre le graph
        mpl.savefig(pathToImg, dpi=72)

    # Try to open document, quit if error
    def openSheet(self, pathToSheet):
        try:
            sheet = ezodf.opendoc(pathToSheet)
        except:
            print("Fichier " + pathToSheet + " introuvable")
            exit(0)
        data = sheet.sheets[0]
        return data
        