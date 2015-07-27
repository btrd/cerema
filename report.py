#!/usr/bin/python
# -*- coding: utf8 -*-
import ezodf
from datetime import datetime, timedelta
from lpod.document import odf_new_document, odf_get_document
from lpod.element import odf_create_element
from lpod.paragraph import odf_create_paragraph
from lpod.style import odf_create_style, odf_master_page, odf_create_table_cell_style
from lpod.image import odf_create_image
from lpod.frame import odf_create_image_frame
from lpod.table import odf_create_table, odf_create_row, odf_create_cell
from PIL import Image
import os

# Crée le raport
class Report(object):
    def __init__(self, pathToReport, pathToParam, pathToBruit, pathToSortie, pathToPic1, pathToPic2, pathToGraph1, pathToGraph2):
        self.param = self.openSheet(pathToParam)
        self.bruit = self.openSheet(pathToBruit)

        sortieFile = self.openDoc(pathToSortie)
        self.sortie = sortieFile.sheets[0]

        self.reportFile = odf_new_document("text")
        self.report = self.reportFile.get_body()

        self.getDataBruit()
        
        self.formatSortie()
        self.getDataSortie()

        self.saveFileEzodf(sortieFile, pathToSortie)

        self.formatPictures(pathToPic1, pathToPic2)
        self.graph1Url = self.reportFile.add_file(pathToGraph1)
        self.graph2Url = self.reportFile.add_file(pathToGraph2)

        self.addStyle()
        self.addContent()
        self.saveFileLpod(self.reportFile, pathToReport)

    # Try to save document, quit if error
    def saveFileEzodf(self, doc, pathToDoc):
        try:
            doc.save()
        except PermissionError:
            print("Le fichier " + pathToDoc + " est utilise par un autre logiciel, impossible de le sauvegarder.")
            exit(1)

    def saveFileLpod(self, doc, pathToDoc):
        try:
            doc.save(target=pathToDoc)
        except PermissionError:
            print("Le fichier " + pathToDoc + " est utilise par un autre logiciel, impossible de le sauvegarder.")
            exit(1)

    # Try to open sheet document, quit if error, returns sheet
    def openSheet(self, pathToSheet):
        try:
            sheet = ezodf.opendoc(pathToSheet)
        except:
            print("Fichier " + pathToSheet + " introuvable")
            exit(1)
        data = sheet.sheets[0]
        return data    # Try to open document, quit if error

    # Try to open sheet document, quit if error, returns doc
    def openDoc(self, pathToDoc):
        try:
            doc = ezodf.opendoc(pathToDoc)
        except:
            print("Fichier " + pathToDoc + " introuvable")
            exit(1)
        return doc

    # Récupère les informations nécessaires dans bruit
    def getDataBruit(self):
        self.startPeriod = self.getDate(self.bruit[2,1].value)
        self.endPeriod = self.getDate(self.bruit[3,1].value)
        self.lieu = str(self.bruit[4,1].value)
        self.weighting = str(self.bruit[5,1].value)
        self.dataType = str(self.bruit[6,1].value)
        self.unit = str(self.bruit[7,1].value)

    def formatSortie(self):
        # enlève les colonnes à ne pas afficher
        self.sortie.delete_columns(index=10, count=13)
        self.sortie.delete_columns(index=0, count=1)

        # enlève les infos à ne pas afficher
        emtptyCell1 = ezodf.Cell("")
        emtptyCell2 = ezodf.Cell("")
        emtptyCell3 = ezodf.Cell("")
        self.sortie[self.sortie.nrows()-1, 0] = emtptyCell1
        self.sortie[self.sortie.nrows()-1, 2] = emtptyCell2
        self.sortie[self.sortie.nrows()-1, 3] = emtptyCell3

        # change le nombre de décimale après la virgule
        # sur les colonnes 5 et 6
        for j in [5, 6]:
            for i in range(1, self.sortie.nrows()-4):
                self.roundCell(i, j, 1)

        # sur (6h - 22h)
        self.roundCell(self.sortie.nrows()-4, 1, 1)
        # sur (22h - 6h)
        self.roundCell(self.sortie.nrows()-3, 1, 1)
        # sur V/jour VL
        self.roundCell(self.sortie.nrows()-3, 7, 0)
        # sur V/jour PL
        self.roundCell(self.sortie.nrows()-3, 8, 0)
        # sur % PL
        self.roundCell(self.sortie.nrows()-1, 8, 1)
        # sur lden
        self.roundCell(self.sortie.nrows()-2, 1, 1)

    #arrondie la valeur de la cell [i, j] à n décimale après la virgule
    def roundCell(self, i, j, n):
        value = round(self.sortie[i, j].value, n)
        self.sortie[i, j].set_value(value)

    # Récupère les informations nécessaires dans sortie
    def getDataSortie(self):
        self.laeq6_22 = self.sortie[self.sortie.nrows()-4, 1].value
        self.laeq22_6 = self.sortie[self.sortie.nrows()-3, 1].value

        self.lden = self.sortie[self.sortie.nrows()-2, 1].value
        self.lnight = self.laeq22_6 - 3

        self.pourcPL = self.sortie[self.sortie.nrows()-1, 8].value
        self.nbrVehicule = self.sortie[self.sortie.nrows()-2, 8].value

    # Take a date string (format ISO 8601) and return a date
    def getDate(self, dateString):
        try:
            time = datetime.strptime(dateString, "%Y-%m-%dT%H:%M:%S")
        except ValueError:
            # If date string is midnight there is no hours
            try:
                time = datetime.strptime(dateString, "%Y-%m-%d")
            except ValueError:
                try:
                    time = datetime.strptime(dateString, "%d/%m/%Y %H:%M:%S")
                except ValueError:
                    print("La premiere colonne doit uniquement contenir des dates")
                    exit(1)
        return time

    # redimensionne les 2 photos et les fusionnent en une image
    def formatPictures(self, pathToPic1, pathToPic2):
        size = 350, 350

        # redimensionne pic1
        im1 = Image.open(pathToPic1)
        im1.thumbnail(size, Image.ANTIALIAS)

        # redimensionne pic2
        im2 = Image.open(pathToPic2)
        im2.thumbnail(size, Image.ANTIALIAS)

        s1 = im1.size
        s2 = im2.size
        #créé une image de la taille de pic1+pic2 en largeur et du max entre pic1, pic2 en hauteur
        new_im = Image.new('RGB', (s1[0]+s2[0], max(s1[1],s2[1])) )

        #colle pic1 et pic2 dans la nouvelle image
        new_im.paste(im1, (0, 0))
        new_im.paste(im2, (s1[0], 0))

        new_im.save("pic_merge.jpeg", "JPEG")

        self.picRatio = (new_im.size[1] + 0.0)/(new_im.size[0] + 0.0)
        self.picUrl = self.reportFile.add_file("pic_merge.jpeg")

    # rajoute les styles utilisé dans le document
    def addStyle(self):
        self.reportFile.delete_styles()
        _style_bold = odf_create_style('text', name = u'bold', bold = True)
        self.reportFile.insert_style(_style_bold)

        _style_red_bold = odf_create_style('text', name = u'red_bold', bold = True, color="#FF0000")
        self.reportFile.insert_style(_style_red_bold)

    # rajoute un paragraphe avec un style appliqué sur tout le paragraphe par défaut
    def addText(self, text, style="", regex=".*"):
        paragraph = odf_create_paragraph()
        paragraph.append_plain_text(text)
        regex = unicode(regex, 'utf-8')
        paragraph.set_span(style, regex=regex)
        self.report.append(paragraph)

    # rajoute une image à report 
    def addImage(self, url, name="", size=("17cm", "7cm")):
        image = odf_create_image_frame(url = url,
            anchor_type = "paragraph",
            name = name,
            size = size
        )
        p = odf_create_paragraph("")
        p.append(image)
        self.report.append(p)

    def addTable(self):
        table = odf_create_table(u"Table")
        self.report.append(table)

        _red_style = odf_create_table_cell_style(background_color = 'red')
        red_style = self.reportFile.insert_style(_red_style, automatic=True)

        _default_style = odf_create_table_cell_style(border = '0.03pt solid #000000')
        default_style = self.reportFile.insert_style(_default_style, automatic=True)
        
        _style_small = odf_create_element(u"""\
        <style:style style:name="small_txt" style:family="table-cell" style:class="text">
            <style:text-properties fo:font-size="10pt"/>
        </style:style>""")
        self.reportFile.insert_style(_style_small)

        _style_width = odf_create_element(u"""
            <style:style style:family="table-column" style:name="colDate">
                <style:table-column-properties style:column-width="3.4cm" style:rel-column-width="13107*"/>
            </style:style>
        """)
        self.reportFile.insert_style(_style_width, automatic=True)
        _style_width = odf_create_element(u"""
            <style:style style:family="table-column" style:name="colNotDate">
                <style:table-column-properties style:column-width="1.7cm" style:rel-column-width="6553.5*"/>
            </style:style>
        """)
        self.reportFile.insert_style(_style_width, automatic=True)

        for i in xrange(0, self.sortie.nrows()):
            row = odf_create_row()
            for j in xrange(0, self.sortie.ncols()):
                cell = odf_create_cell()
                cell.set_style("small_txt")
                cell.set_style(_default_style)
                if j == 0 and i < self.sortie.nrows()-4 and i > 0:
                    value = self.getDate(self.sortie[i, j].value).strftime('%d/%m/%Y %Hh')
                    cell.set_value(value)
                elif j == 6 and i > 0 and i < self.sortie.nrows()-4 and self.sortie[i, j].value > 1:
                    cell.set_style(red_style)
                    cell.set_value(self.sortie[i, j].value)
                elif (j == 7 or j == 8) and i > 0 and i < self.sortie.nrows()-2 :
                    cell.set_value(int(self.sortie[i, j].value))
                else:
                    cell.set_value(self.sortie[i, j].value)
                row.set_cell(j, cell)
            table.set_row(i, row)

        first = True
        for column in table.get_columns():
            if first:
                column.set_style("colDate")
                table.set_column(column.x, column)
                first = False
            else:
                column.set_style("colNotDate")
                table.set_column(column.x, column)

    # rajoute le contenu dans report
    def addContent(self):
        text = "Traffic du " + str(self.startPeriod.day) + " au " + self.endPeriod.strftime('%d/%m/%Y')
        self.addText(text, style="bold", regex = "Traffic")

        text = "Véh/j " + str(self.nbrVehicule) + " PL " + str(self.pourcPL) + "%"
        self.addText(text, style="bold")

        self.addText("Résultats des mesures", style="bold")

        self.addText("Lieu \t\t\t" + self.lieu)
        self.addText("Type de données \t" + self.dataType)
        self.addText("Pondération \t\t" + self.weighting)
        self.addText("Unité \t\t\t" + self.unit)
        self.addText("Début \t\t\t" + self.endPeriod.strftime('%d/%m/%Y %H:%M'))
        self.addText("Fin \t\t\t" + self.startPeriod.strftime('%d/%m/%Y %H:%M'))

        self.addText("Lden \t\t\t" + str(self.lden))
        self.addText("Lnight \t\t\t" + str(self.lnight))

        self.addText("LAeq(6h-22h) \t" + str(self.laeq6_22))
        self.addText("LAeq(22h-6h) \t" + str(self.laeq22_6))

        self.addImage(self.picUrl, "Photographie", ("17cm", str(17 * self.picRatio) + "cm"))

        self.addText("\tÉvolution temporelle en Laeq par pas de 15 minutes (Laeq élémentaire 1 seconde).\n", style="red_bold")
        
        self.addImage(self.graph2Url, "laeq 15min")

        self.addText("Remarque : \n", style="bold", regex="Remarque :")

        self.addText("La validation des résultats des mesurages fait l’objet de tests :")
        self.addText("\n\t•\tTest temporel : continuité du signal\n", style="bold", regex="Test temporel :")
        self.addText("Certains évènements particuliers ont été isolés par codage. Test réalisé par l’opérateur.")
        self.addText("\n\t•\tTest statistique : répartition gaussienne du bruit du trafic routier\n", style="bold", regex="Test statistique :")

        text = "Semaine du " + str(self.startPeriod.day) + " au " + self.endPeriod.strftime('%d %B %Y') + " de " + str(self.startPeriod.hour) + "h à " + str(self.endPeriod.hour) + "h"
        self.addText(text, style="bold")

        self.addImage(self.graph1Url, "recalage trafic")

        self.addText("•  Corrélation bruit/trafic :", style="bold")
        self.addText("\n\n")

        self.addTable()
