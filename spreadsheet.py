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
    res_doc.sheets += bruits_doc.sheets[0]
    try:
        res_doc.save()
    except PermissionError:
        print("Le fichier sorties.ods est utilisé par un autre logiciel, impossible de le sauvegarder.")
        exit(0)
    