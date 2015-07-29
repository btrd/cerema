#!/usr/bin/python
# coding: utf8
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
    def __init__(self, pathToReport, pathToParam, pathToBruit, pathToSortie, pathToPic1, pathToPic2, pathToGraph1, pathToGraph2, code, realise, depouille):
        self.code = code
        self.realise = realise
        self.depouille = depouille

        # get files
        self.param = self.openParam(pathToParam)
        self.bruit = self.openSheet(pathToBruit)

        sortieFile = self.openDoc(pathToSortie)
        self.sortie = sortieFile.sheets[0]

        # create report
        self.reportFile = odf_new_document("text")
        self.report = self.reportFile.get_body()

        # gather data
        self.getDataBruit()
        self.getDataParam()
        
        self.formatSortie()
        self.getDataSortie()

        self.saveFileEzodf(sortieFile, pathToSortie)

        self.formatPictures(pathToPic1, pathToPic2)
        self.graph1Url = self.reportFile.add_file(pathToGraph1)
        self.graph2Url = self.reportFile.add_file(pathToGraph2)

        # add content to report
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
            sheetFile = ezodf.opendoc(pathToSheet)
        except:
            print("Fichier " + pathToSheet + " introuvable")
            exit(1)
        data = sheetFile.sheets[0]
        return data    # Try to open document, quit if error

    # Try to open sheet document, quit if error, returns doc
    def openDoc(self, pathToDoc):
        try:
            doc = ezodf.opendoc(pathToDoc)
        except:
            print("Fichier " + pathToDoc + " introuvable")
            exit(1)
        return doc

    def openParam(self, pathToParam):
        try:
            sheetFile = ezodf.opendoc(pathToParam)
        except:
            print("Fichier " + pathToParam + " introuvable")
            exit(1)

        i = 1
        for sheet in sheetFile.sheets:
            print str(i) + ": " + sheet.name
            i += 1

        select = raw_input("Quelle feuille de " + pathToParam + ": ")
        if select == "": # aucune valeur
            index = 0
        elif select.isdigit(): # index de la feuille
            index = int(select)-1
        else:  # nom de la feuille
            index = select
        return sheetFile.sheets[index]

    # Récupère les informations nécessaires dans bruit
    def getDataBruit(self):
        self.startPeriod = self.getDate(self.bruit[2,1].value)
        self.endPeriod = self.getDate(self.bruit[3,1].value)
        self.lieu = (self.bruit[4,1].value or "").encode('utf-8','replace')
        self.weighting = (self.bruit[5,1].value or "").encode('utf-8','replace')
        self.dataType = (self.bruit[6,1].value or "").encode('utf-8','replace')
        self.unit = (self.bruit[7,1].value or "").encode('utf-8','replace')

    def getDataParam(self):
        self.pointDe = (self.param[0,1].value or "").encode('utf-8','replace')
        self.sonometre = (self.param[0,3].value or "").encode('utf-8','replace')
        self.nom = (self.param[6,2].value or "").encode('utf-8','replace')
        self.adresse1 = (self.param[8,2].value or "").encode('utf-8','replace')
        self.adresse2 = str(int(self.param[10,2].value or 0)) + " " + (self.param[9,2].value or "").encode('utf-8','replace')
        self.distanceVoie = ""
        self.hauteurPriseSon = (self.param[16,0].value or "").encode('utf-8','replace') + " " + (self.param[15,2].value or "").encode('utf-8','replace')
        self.natureSol = (self.param[26,3].value or "").encode('utf-8','replace')

        self.nebulosite = (self.param[34,2].value or "").encode('utf-8','replace') + "/" + (self.param[35,2].value or "").encode('utf-8','replace')
        self.directVent = "" #TODO
        self.forceVent = "" #TODO
        self.tempDeb = str(self.param[34,3].value or "")
        self.tempFin = str(self.param[35,3].value or "")

        self.nbrVoies = (self.param[24,3].value or "").encode('utf-8','replace')
        self.profilTravers = (self.param[23,3].value or "").encode('utf-8','replace')

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
        _red_style = odf_create_table_cell_style(background_color = 'red')
        red_style = self.reportFile.insert_style(_red_style, automatic=True)

        _default_style = odf_create_table_cell_style(border='0.03pt solid #000000', padding="0.05cm")
        default_style = self.reportFile.insert_style(_default_style, automatic=True)

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

        table = odf_create_table(u"Table")
        self.report.append(table)

        for i in xrange(0, self.sortie.nrows()):
            row = odf_create_row()
            for j in xrange(0, self.sortie.ncols()):
                cell = odf_create_cell()
                cell.set_style(default_style)
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

    def createCell(self, value):
        if self.cptCell in [0, 2, 50, 48, 64]:
            style = self.cctl
        elif self.cptCell in [1, 3, 51, 49, 65]:
            style = self.cctr
        elif self.cptCell in [42, 40, 70, 68, 56]:
            style = self.ccbl
        elif self.cptCell in [43, 41, 71, 69, 57]:
            style = self.ccbr
        elif self.cptCell in [44, 45, 46, 47, 60, 61]:
            style = ""
        elif self.cptCell % 4 == 0:
            style = self.ccl
        elif self.cptCell % 4 == 1:
            style = self.ccr
        elif self.cptCell % 4 == 2:
            style = self.ccl
        elif self.cptCell % 4 == 3:
            style = self.ccr
        cell = odf_create_cell(value=value.decode("utf-8"), cell_type="string", style=style)
        self.row.set_cell(self.cptCell % 4, cell)
        self.cptCell += 1

    # rajoute le contenu dans report
    def addContent(self):
        _style_bold = odf_create_style('text', name = u'bold', bold = True, automatic=True)
        self.reportFile.insert_style(_style_bold)

        _style_red_bold = odf_create_style('text', name = u'red_bold', bold = True, color="#FF0000")
        self.reportFile.insert_style(_style_red_bold)

        border_style = '1pt solid #000000'
        padding_style = "0.05cm"
        #left border
        _ccl = odf_create_table_cell_style(border_left=border_style, padding=padding_style)
        self.ccl = self.reportFile.insert_style(_ccl, automatic=True)
        #right border
        _ccr = odf_create_table_cell_style(border_right=border_style, padding=padding_style)
        self.ccr = self.reportFile.insert_style(_ccr, automatic=True)
        #top right border
        _cctr = odf_create_table_cell_style(border_right=border_style, border_top=border_style, padding=padding_style)
        self.cctr = self.reportFile.insert_style(_cctr, automatic=True)
        #top left border
        _cctl = odf_create_table_cell_style(border_left=border_style, border_top=border_style, padding=padding_style)
        self.cctl = self.reportFile.insert_style(_cctl, automatic=True)
        #bottom right border
        _ccbr = odf_create_table_cell_style(border_right=border_style, border_bottom=border_style, padding=padding_style)
        self.ccbr = self.reportFile.insert_style(_ccbr, automatic=True)
        #bottom left border
        _ccbl = odf_create_table_cell_style(border_left=border_style, border_bottom=border_style, padding=padding_style)
        self.ccbl = self.reportFile.insert_style(_ccbl, automatic=True)

        paragraph = odf_create_paragraph(self.code, style="Title")
        self.report.append(paragraph)

        table = odf_create_table(u"Table")
        self.report.append(table)

        self.cptCell = 0
        self.row = odf_create_row()
        cell = self.createCell("Description du point de mesure")
        cell = self.createCell("")
        cell = self.createCell("Résultats des mesures")
        cell = self.createCell("")
        table.set_row(0, self.row)

        self.row = odf_create_row()
        cell = self.createCell("Point de")
        cell = self.createCell(self.pointDe)
        cell = self.createCell("Lieu")
        cell = self.createCell(self.lieu)
        table.set_row(1, self.row)

        self.row = odf_create_row()
        cell = self.createCell("Sonomètre utilisé")
        cell = self.createCell(self.sonometre)
        cell = self.createCell("Type de données")
        cell = self.createCell(self.dataType)
        table.set_row(2, self.row)

        self.row = odf_create_row()
        cell = self.createCell("Nom")
        cell = self.createCell(self.nom)
        cell = self.createCell("Pondération")
        cell = self.createCell(self.weighting)
        table.set_row(3, self.row)

        self.row = odf_create_row()
        cell = self.createCell("Adresse")
        cell = self.createCell(self.adresse1)
        cell = self.createCell("Unité")
        cell = self.createCell(self.unit)
        table.set_row(4, self.row)

        self.row = odf_create_row()
        cell = self.createCell("")
        cell = self.createCell(self.adresse2)
        cell = self.createCell("Début")
        cell = self.createCell(self.startPeriod.strftime('%d/%m/%Y %H:%M'))
        table.set_row(5, self.row)

        self.row = odf_create_row()
        cell = self.createCell("Distance voie")
        cell = self.createCell(self.distanceVoie)
        cell = self.createCell("Fin")
        cell = self.createCell(self.endPeriod.strftime('%d/%m/%Y %H:%M'))
        table.set_row(6, self.row)

        self.row = odf_create_row()
        cell = self.createCell("H. prise de son")
        cell = self.createCell(self.hauteurPriseSon)
        cell = self.createCell("Lden")
        cell = self.createCell(str(self.lden))
        table.set_row(7, self.row)

        self.row = odf_create_row()
        cell = self.createCell("Nature du sol")
        cell = self.createCell(self.natureSol)
        cell = self.createCell("Lnight")
        cell = self.createCell(str(self.lnight))
        table.set_row(8, self.row)

        self.row = odf_create_row()
        cell = self.createCell("Réalisée par")
        cell = self.createCell(self.realise)
        cell = self.createCell("LAeq (6h-22h)")
        cell = self.createCell(str(self.laeq6_22))
        table.set_row(9, self.row)

        self.row = odf_create_row()
        cell = self.createCell("Dépouillée par")
        cell = self.createCell(self.depouille)
        cell = self.createCell("LAeq (22h-6h)")
        cell = self.createCell(str(self.laeq22_6))
        table.set_row(10, self.row)

        self.row = odf_create_row()
        cell = self.createCell("")
        cell = self.createCell("")
        cell = self.createCell("")
        cell = self.createCell("")
        table.set_row(11, self.row)

        self.row = odf_create_row()
        cell = self.createCell("Caractéristique de la voie")
        cell = self.createCell("")
        cell = self.createCell("Météo")
        cell = self.createCell("")
        table.set_row(12, self.row)

        self.row = odf_create_row()
        cell = self.createCell("Nombre de voies")
        cell = self.createCell(self.nbrVoies)
        cell = self.createCell("Nébulosité")
        cell = self.createCell(self.nebulosite)
        table.set_row(13, self.row)

        self.row = odf_create_row()
        cell = self.createCell("Profil en travers")
        cell = self.createCell(self.profilTravers)
        cell = self.createCell("Direction vent")
        cell = self.createCell(self.directVent)
        table.set_row(14, self.row)

        self.row = odf_create_row()
        cell = self.createCell("")
        cell = self.createCell("")
        cell = self.createCell("Force vent")
        cell = self.createCell(self.forceVent)
        table.set_row(15, self.row)

        self.row = odf_create_row()
        cell = self.createCell("Traffic du " + str(self.startPeriod.day) + " au " + self.endPeriod.strftime('%d/%m/%Y'))
        cell = self.createCell("")
        cell = self.createCell("Température début")
        cell = self.createCell(self.tempDeb)
        table.set_row(16, self.row)

        self.row = odf_create_row()
        cell = self.createCell("Véh/j " + str(self.nbrVehicule))
        cell = self.createCell("PL " + str(self.pourcPL) + "%")
        cell = self.createCell("Température début")
        cell = self.createCell(self.tempFin)
        table.set_row(17, self.row)

        self.addImage(self.picUrl, "Photographie", ("17cm", str(17 * self.picRatio) + "cm"))

        self.addText("\tÉvolution temporelle en Laeq par pas de 15 minutes (Laeq élémentaire 1 seconde).\n", style="red_bold")
        
        self.addImage(self.graph2Url, "laeq 15min", ("15cm", "5.5cm"))

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
