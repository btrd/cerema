#!/usr/bin/python
# -*- coding: utf8 -*-
import ezodf

class Report(object):
    def __init__(self, pathToReport, pathToParam, pathData, pathToBruit):
        self.param = self.openSheet(pathToParam)
        self.bruit = self.openSheet(pathToBruit)
        self.data = self.openSheet(pathData)

        self.report = self.createDoc(pathToReport)
        
        self.report.body.append(ezodf.Heading("A Simple Test Document"))
        self.report.body.append(ezodf.ezlist(['Point 1', 'Point 2', 'Point 3']))
        self.report.body.append(ezodf.SoftPageBreak())
        self.report.body.append(ezodf.Heading("A Simple Test Document"))

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