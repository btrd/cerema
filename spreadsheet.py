#!/usr/bin/python
# -*- coding: utf8 -*-
import ezodf
def createSpreadsheet(pathToBruit):
    try:
        bruits_doc = ezodf.opendoc(pathToBruit)
        res_doc = ezodf.newdoc(doctype="ods", filename="data/sortie.ods")
    except:
        print("Fichier " + pathToBruit + " introuvable")
        exit(0)
    data = bruits_doc.sheets[0]
    print(getData(data))
    addData(data)
    res_doc.sheets += data
    try:
        res_doc.save()
    except PermissionError:
        print("Le fichier sorties.ods est utilisé par un autre logiciel, impossible de le sauvegarder.")
        exit(0)
    
def getData(data):
    startPeriod = data[2,1].value
    endPeriod = data[3,1].value
    lieu = data[4,1].value
    weighting = data[5,1].value
    dataType = data[6,1].value
    unit = data[7,1].value
    data.delete_rows(0, 8)
    return lieu, dataType, weighting, unit, startPeriod, endPeriod

def addData(data):
    data.append_columns(1)
    data[0, data.ncols()-1].set_value("Laeq Gauss")
    for x in range(2, data.nrows()):
        form = "=(E" + str(x) + "+D" + str(x) + ")/2+0,0175*(E" + str(x) + "-D" + str(x) + ")^2"
        data[x-1, data.ncols()-1].formula = form
