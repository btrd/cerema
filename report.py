#!/usr/bin/python
# -*- coding: utf8 -*-
import ezodf

class Report(object):
    def __init__(self, pathToReport, pathToParam):
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
            exit(0)
        return doc

    # Try to save document, quit if error
    def saveDoc(self, doc, pathToDoc):
        try:
            doc.save()
        except PermissionError:
            print("Le fichier " + pathToDoc + " est utilisé par un autre logiciel, impossible de le sauvegarder.")
            exit(0)