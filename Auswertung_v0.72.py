"""


LEIBNIZ UNIVERSITÄT HANNOVER
Institut für Geotechnik(IGtH)

Skript für Automatisierung der K+S Auswertung v0.72


Coded by Eren Ahmet Isik


"""

import copy
import os
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch, cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PyPDF2 import PdfFileWriter, PdfFileReader
from pandas import read_excel
from numpy import nan
import io
import fnmatch
import matplotlib.pyplot as plt
import matplotlib as mpl


from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import openpyxl
import shutil


# Die Funktion hier arbeitet als Hauptfunktion.

# Funktionsparameter : Versuchsliste : Excel-Datei
# Layout: Datum des PDF-Layouts
# X_limit und x_limit_2 : Grenze von X-Achsen

def haupt_funktion(versuch_list,versuch_list_name, layout,x_limit,x_limit_2,block_ort_lfs,block_ort,y_limit,y_limit_2):
    i = block_ort #
    import os

    datei_typen = datei_suchen() #Die Funktion hier sucht nach Rohdaten.

    packet = io.BytesIO() #Ein Parameter für PDF Erstellen
    x = 21 * cm # A4 Breite
    y = 29.7 * cm # A4 Höhe
    seite_größe = (x, y)

    x1 = 11.36 * cm
    y1 = 23.87 * cm
    kanvas = canvas.Canvas(packet, pagesize=seite_größe) # Library um PDF zu erstellen

    file_name = versuch_list_name

    data_path = os.getcwd()

    copy_file_name = "Neu_" + versuch_list_name

    kontroll_datei = os.path.join(data_path, copy_file_name)

    if os.path.exists(kontroll_datei):
        versuch_list_excel = openpyxl.load_workbook(copy_file_name)
    else:

        shutil.copy(file_name, copy_file_name)
        versuch_list_excel = openpyxl.load_workbook(copy_file_name)
        print("Neuer Übersicht-Datei erstellt.")

    #Der Teil hier wiederholt sich im Allgemeinen so oft wie die Versuchnummer.

    for j in range(0, versuch_list.shape[0]):

        if versuch_list.to_numpy()[j][block_ort_lfs] == "x":
            proben_nummer = probennummer(versuch_list, i, j)  # Probennummer Funktion

            nummer_versuchblock = nummerierung_versuchblock(versuch_list, block_ort,int(proben_nummer))




            drehwinkel, spannung = grafik_datei(proben_nummer,
                                                datei_typen)  # Funktion für Drehwinkel und Spannung. Es liest die Rohdaten.


            try:
                tabelle_werten = grafik_darstellung(versuch_list_excel,versuch_list_name,drehwinkel, spannung, proben_nummer, x_limit,
                                                    x_limit_2,y_limit,y_limit_2)  # ein Funktion Erstellung von Grafiken Nach Probendaten.

            except IndexError:  # Wir ignorieren Fehler.
                pass

            try:

                ### In diesem Teil lesen wir im Allgemeinen die Daten in der Excel-Datei aus und bestimmen den Ort des Experiments.
                # Wichtige Parameterposition. z.B : Süd, 50m über Fuß Listenlänge 2, Plateau: Listenlänge 1, Das Programm wird für diese bedingten Bedingungen fortgesetzt.
                if len(probe_position(versuch_list, j)) > 1:
                    pdfmetrics.registerFont(TTFont('Arial', 'Fonts/arial.ttf'))
                    kanvas.setFont("Arial", 12)
                    # kanvas.setFont("Vera",12)

                    kanvas.drawString(8.54 * cm, y1,
                                      "REKAL Halde,")  # Diese Funktion wird verwendet, um Text in PDF zu schreiben. Text : REKAL Halde

                    position = probe_position(versuch_list, j)  # Es ruft die Funktion Position auf.

                    versuch_block = versuchblock(position)  # Es ruft die Funktion Versuchblock auf.
                    H = h_nummerierung_(versuch_list, i)  # Es ruft die H_nummerieung auf.
                    probeentnahme_termin = str(Probeentnahmetermin(versuch_list, i))  # Es ruft die Funktion Probeentnahme _ermin auf.
                    m = meter(versuch_list, i, j)  # Es ruft die Funktion Position als meter auf.
                    datum = entnahmedatum(versuch_list, i)  # Es ruft die Funktion Datum  auf.

                    sondierdatum = sondier_datum(versuch_list, i, j,
                                                 proben_nummer)  # Es ruft die Funktion Sondierdatum auf.

                    wasser_gehalt = wassergehalt(versuch_list, i, j)  # Es ruft die Funktion Wassergehalt auf.
                    feucht_dichte = feuchtdichte(versuch_list, i, j)  # Es ruft die Funktion Feuchtdicte auf.
                    trocken_dichte = trockendichte(versuch_list, i, j)  # Es ruft die Funktion Trockendichte auf.


                    # Wir leben die Werte, die wir in diesem Abschnitt auf dem pdf berechnet haben.

                    kanvas.drawString(x1, y1, position[0])  # Koordinat von Position
                    kanvas.drawString(9.8 * cm, 23.50 * cm, position[1])
                    kanvas.drawString(7.54 * cm, 23.0 * cm, "Ergebnisse der Laborflügelsondierungen")
                    kanvas.drawString(17.98 * cm, 25.83 * cm, H)
                    kanvas.drawString(16 * cm, 25.831 * cm, "Anhang :")
                    kanvas.drawString(18.8 * cm, 25.83 * cm, ".4")
                    kanvas.drawString(19.2 * cm, 25.83 * cm, versuch_block)
                    kanvas.drawString(19.6 * cm, 25.83 * cm, nummer_versuchblock)

                    kanvas.drawString(16 * cm, 25.49 * cm, f"zu Az.:   01/00-{probeentnahme_termin}")
                    kanvas.setFillColorRGB(1, 0, 0)
                    kanvas.drawString(2.5 * cm, 21.75 * cm, f"Probe :  {proben_nummer}, obere Probenseite ")
                    kanvas.setFillColorRGB(0, 0, 0)
                    kanvas.drawString(2.5 * cm, 21.23 * cm,
                                      f"Entnahmestelle : {position[0]}, rd. {m} m von der Böschungskante entfernt")
                    kanvas.drawString(2.5 * cm, 20 * cm, "Bodenart : Rekal + Asche")
                    kanvas.drawString(2.5 * cm, 19.5 * cm, f"Entnahmedatum :   {datum}")
                    kanvas.drawString(2.5 * cm, 19 * cm, f"Sondierdatum     :   {sondierdatum}   ")
                    kanvas.drawString(2.5 * cm, 18 * cm, f"Wassergehalt  :   {wasser_gehalt}   ")
                    kanvas.drawString(2.5 * cm, 17.5 * cm, f"Feuchtdichte   :   {feucht_dichte}   ")
                    kanvas.drawString(2.5 * cm, 17.0 * cm, f"Trockendichte :   {trocken_dichte}   ")
                    kanvas.drawString(2.5 * cm, 16.0 * cm, "Schergeschwindigkeit : ω  = 0,1°/s ")
                    kanvas.drawString(2.5 * cm, 15.5 * cm, "Flügelabmessungen : H/D = 25,4 mm / 12,7 mm ")

                    # Wir leben von dem pdf mit den Maximalwerten in der quadratischen Box, die wir in diesem Abschnitt berechnet haben.

                    try:

                        kanvas.drawString(12.80 * cm, 16.00 * cm, f"{tabelle_werten[0][0]}".replace(".",","))
                        kanvas.drawString(15.75 * cm, 16.00 * cm, f"{tabelle_werten[0][1]}".replace(".",","))

                        kanvas.setFillColorRGB(1, 0, 0)

                        kanvas.drawString(12.80 * cm, 15.50 * cm, f"{tabelle_werten[1][0]}".replace(".",","))
                        kanvas.drawString(15.75 * cm, 15.50 * cm, f"{tabelle_werten[1][1]}".replace(".",","))

                        kanvas.setFillColorRGB(0, 0, 1)

                        kanvas.drawString(12.80 * cm, 15.00 * cm, f"{tabelle_werten[2][0]}".replace(".",","))
                        kanvas.drawString(15.75 * cm, 15.00 * cm, f"{tabelle_werten[2][1]}".replace(".",","))

                        kanvas.setFillColorRGB(0, 1, 0)
                        kanvas.drawString(12.80 * cm, 14.57 * cm, f"{tabelle_werten[3][0]}".replace(".",","))
                        kanvas.drawString(15.75 * cm, 14.57 * cm, f"{tabelle_werten[3][1]}".replace(".",","))

                        kanvas.setFillColorRGB(0, 0, 0)


                    except IndexError:
                        pass

                    # Für die richtige Versuchsnummer suchen wir die zuvor erstellten Graphen in der Datei „aus Grafiken“.
                    try:

                        import os
                        data_path = os.getcwd()

                        kanvas.drawImage(data_path + "\\Grafiken\\" + f"{proben_nummer}" + ".png", 0.2 * cm, 0.2 * cm,
                                         width=20 * cm, height=13 * cm)
                    except OSError:
                        pass
                    kanvas.showPage()



                # Die zweite Bedingung ist, dass der Positionsname unterschiedlich ist.

                elif len(probe_position(versuch_list, j)) == 1 and probe_position(versuch_list, j)[0] != "nan":
                    pdfmetrics.registerFont(TTFont('Arial', 'Fonts/arial.ttf'))

                    kanvas.setFont("Arial", 12)
                    kanvas.drawString(8.54 * cm, y1, "REKAL Halde,")

                    position = probe_position(versuch_list, j)
                    versuch_block = versuchblock(position)
                    H = h_nummerierung_(versuch_list, i)
                    probeentnahme_termin = Probeentnahmetermin(versuch_list, i)
                    m = meter(versuch_list, i, j)
                    datum = entnahmedatum(versuch_list, i)
                    sondierdatum = sondier_datum(versuch_list, i, j, proben_nummer)
                    wasser_gehalt = wassergehalt(versuch_list, i, j)
                    feucht_dichte = feuchtdichte(versuch_list, i, j)
                    trocken_dichte = trockendichte(versuch_list, i, j)
                    nummer_versuchblock = nummerierung_versuchblock(versuch_list,block_ort,int(proben_nummer))
                    kanvas.drawString(x1, y1, position[0])
                    kanvas.drawString(6.3 * cm, 23.0 * cm, "Ergebnisse der Laborflügelsondierungen")

                    kanvas.drawString(17.98 * cm, 25.83 * cm, H)
                    kanvas.drawString(16 * cm, 25.83 * cm, "Anhang :")
                    kanvas.drawString(18.8 * cm, 25.83 * cm, ".4")
                    kanvas.drawString(19.2 * cm, 25.83 * cm, versuch_block)
                    kanvas.drawString(19.6 * cm, 25.83 * cm, nummer_versuchblock)

                    kanvas.drawString(16 * cm, 25.49 * cm, f"zu Az.:   01/00-{probeentnahme_termin}")
                    kanvas.setFillColorRGB(1, 0, 0)

                    kanvas.drawString(2.5 * cm, 21.75 * cm, f"Probe :  {proben_nummer}, obere Probenseite ")
                    kanvas.setFillColorRGB(0, 0, 0)

                    kanvas.drawString(2.5 * cm, 21.23 * cm,
                                      f"Entnahmestelle : {position[0]}, rd. {m} m von der Böschungskante entfernt")
                    kanvas.drawString(2.5 * cm, 20 * cm, "Bodenart : Rekal + Asche")
                    kanvas.drawString(2.5 * cm, 19.5 * cm, f"Entnahmedatum :   {datum}")
                    kanvas.drawString(2.5 * cm, 19 * cm, f"Sondierdatum     :   {sondierdatum}   ")
                    kanvas.drawString(2.5 * cm, 18 * cm, f"Wassergehalt  :   {wasser_gehalt}   ")
                    kanvas.drawString(2.5 * cm, 17.5 * cm, f"Feuchtdichte   :   {feucht_dichte}   ")
                    kanvas.drawString(2.5 * cm, 17.0 * cm, f"Trockendichte :   {trocken_dichte}   ")
                    kanvas.drawString(2.5 * cm, 16.0 * cm, "Schergeschwindigkeit : ω  = 0,1°/s ")
                    kanvas.drawString(2.5 * cm, 15.5 * cm, "Flügelabmesungen : H/D = 25/12,5 mm ")

                    try:

                        kanvas.drawString(12.80 * cm, 16.00 * cm, f"{tabelle_werten[0][0]}".replace(".",","))
                        kanvas.drawString(15.75 * cm, 16.00 * cm, f"{tabelle_werten[0][1]}".replace(".",","))

                        kanvas.setFillColorRGB(1, 0, 0)

                        kanvas.drawString(12.80 * cm, 15.50 * cm, f"{tabelle_werten[1][0]}".replace(".",","))
                        kanvas.drawString(15.75 * cm, 15.50 * cm, f"{tabelle_werten[1][1]}".replace(".",","))

                        kanvas.setFillColorRGB(0, 0, 1)

                        kanvas.drawString(12.80 * cm, 15.00 * cm, f"{tabelle_werten[2][0]}".replace(".",","))
                        kanvas.drawString(15.75 * cm, 15.00 * cm, f"{tabelle_werten[2][1]}".replace(".",","))

                        kanvas.setFillColorRGB(0, 1, 0)
                        kanvas.drawString(12.80 * cm, 14.50 * cm, f"{tabelle_werten[3][0]}".replace(".",","))
                        kanvas.drawString(15.75 * cm, 14.50 * cm, f"{tabelle_werten[3][1]}".replace(".",","))

                        kanvas.setFillColorRGB(0, 0, 0)




                    except IndexError:
                        pass

                    # kanvas.rect(12.7*cm,14.6*cm,6*cm,4*cm)
                    # kanvas.line(12.7*cm,16.6*cm,12.7*cm,18.7)

                    # rect = Rect(12.7*cm,16.7*cm,2*cm,2*cm)
                    # grafik = Image("E:\Projekt\IGTH/test.png")
                    # grafik._restrictSize(1 * cm, 2 * cm)
                    try:
                        import os
                        data_path = os.getcwd()

                        kanvas.drawImage(data_path + "\\Grafiken\\" + f"{proben_nummer}" + ".png", 0.2 * cm, 0.2 * cm,
                                         width=20 * cm, height=13 * cm)
                    except OSError:
                        pass
                    kanvas.showPage()
            except TypeError:
                j += 1


        else:

            pass
    # Der Name der zu speichernden Datei.

    # Der Dateiaufzeichnungsprozess, nachdem alle Prozesse in diesem Abschnitt abgeschlossen sind.
    kanvas.save()

    packet.seek(0)

    canvas_pdf = PdfFileReader(packet)

    canvas_pdf.getNumPages()

    outfile = PdfFileWriter()

    for z in range(0, canvas_pdf.getNumPages()):
        page = copy.copy(layout.getPage(0))
        page.mergePage(canvas_pdf.getPage(z))

        outfile.addPage(page)

    with open(datum + "_Auswertungen_LFS.pdf", "wb") as out:
        outfile.write(out)


# Diese Funktion ist die gleiche wie oben, außer dass sie auf der Grundlage einer einzelnen Versuchnummer funktioniert.





# Funktionsparameter : Versuchsliste : Excel-Datei
# Layout: Datum des PDF-Layouts
# Index : Die Indexnummer der gewünschten Probennummer
# block_ort : Blocknummer
# X_limit und x_limit_2 : Grenze von X-Achsen

def haupt_funktion2(versuch_list, layout, index, block_ort,x_limit,x_limit_2,y_limit,y_limit_2):
    seite = versuch_list.shape[0] // 19
    datei_typen = datei_suchen()

    packet = io.BytesIO()
    x = 21 * cm
    y = 29.7 * cm
    seite_größe = (x, y)

    x1 = 11.36 * cm
    y1 = 23.87 * cm
    kanvas = canvas.Canvas(packet, pagesize=seite_größe)
    # pdfmetrics.registerFont(TTFont('Vera', 'Vera.ttf'))
    # pdfmetrics.registerFont(TTFont('VeraBd', 'VeraBd.ttf'))
    # pdfmetrics.registerFont(TTFont('VeraIt', 'VeraIt.ttf'))
    # pdfmetrics.registerFont(TTFont('VeraBI', 'VeraBI.ttf'))"""
    # Registered fontfamily
    # pdfmetrics.registerFontFamily('Vera', normal='Vera', bold='VeraBd', italic='VeraIt', boldItalic='VeraBI')
    for a in range(1):

        for b in range(1):

            proben_nummer = probennummer(versuch_list, block_ort, index)

            drehwinkel, spannung = grafik_datei(proben_nummer, datei_typen)

            nummer_versuchblock = nummerierung_versuchblock(versuch_list, block_ort, int(proben_nummer))


            try:
                tabelle_werten = grafik_darstellung(drehwinkel, spannung, proben_nummer,x_limit,x_limit_2,y_limit,y_limit_2)
            except IndexError:
                pass
            try:
                if len(probe_position(versuch_list, index)) > 1:
                    pdfmetrics.registerFont(TTFont('Arial', 'Fonts/arial.ttf'))
                    kanvas.setFont("Arial", 12)
                    # kanvas.setFont("Vera",12)

                    kanvas.drawString(8.54 * cm, y1, "REKAL Halde,")
                    position = probe_position(versuch_list, index)
                    versuch_block = versuchblock(position)
                    H = h_nummerierung_(versuch_list, block_ort)
                    probeentnahme_termin = str(Probeentnahmetermin(versuch_list, block_ort))
                    m = meter(versuch_list, block_ort, index)
                    datum = entnahmedatum(versuch_list, block_ort)
                    sondierdatum = sondier_datum(versuch_list, block_ort, index, proben_nummer)
                    wasser_gehalt = wassergehalt(versuch_list, block_ort, index)
                    feucht_dichte = feuchtdichte(versuch_list, block_ort, index)
                    trocken_dichte = trockendichte(versuch_list, block_ort, index)
                    nummer_versuchblock = nummerierung_versuchblock(versuch_list,block_ort,int(proben_nummer))
                    kanvas.drawString(x1, y1, position[0])  # Koordinat von Position
                    kanvas.drawString(9.8 * cm, 23.50 * cm, position[1])
                    kanvas.drawString(7.54 * cm, 23.0 * cm, "Ergebnisse der Laborflügelsondierungen")
                    kanvas.drawString(17.98 * cm, 25.83 * cm, H)
                    kanvas.drawString(16 * cm, 25.831 * cm, "Anhang :")
                    kanvas.drawString(18.8 * cm, 25.83 * cm, ".4")
                    kanvas.drawString(19.2 * cm, 25.83 * cm, versuch_block)
                    kanvas.drawString(19.6 * cm, 25.83 * cm, nummer_versuchblock)

                    kanvas.drawString(16 * cm, 25.49 * cm, f"zu Az.:   01/00-{probeentnahme_termin}")
                    kanvas.setFillColorRGB(1, 0, 0)
                    kanvas.drawString(2.5 * cm, 21.75 * cm, f"Probe :  {proben_nummer}, obere Probenseite ")
                    kanvas.setFillColorRGB(0, 0, 0)
                    kanvas.drawString(2.5 * cm, 21.23 * cm,
                                      f"Entnahmestelle : {position[0]}, rd. {m} m von der Böschungskante entfernt")
                    kanvas.drawString(2.5 * cm, 20 * cm, "Bodenart : Rekal + Asche")
                    kanvas.drawString(2.5 * cm, 19.5 * cm, f"Entnahmedatum :   {datum}")
                    kanvas.drawString(2.5 * cm, 19 * cm, f"Sondierdatum     :   {sondierdatum}   ")
                    kanvas.drawString(2.5 * cm, 18 * cm, f"Wassergehalt  :   {wasser_gehalt}   ")
                    kanvas.drawString(2.5 * cm, 17.5 * cm, f"Feuchtdichte   :   {feucht_dichte}   ")
                    kanvas.drawString(2.5 * cm, 17.0 * cm, f"Trockendichte :   {trocken_dichte}   ")
                    kanvas.drawString(2.5 * cm, 16.0 * cm, "Schergeschwindigkeit : ω  = 0,1°/s ")
                    kanvas.drawString(2.5 * cm, 15.5 * cm, "Flügelabmessungen : H/D = 25,4 mm / 12,7 mm ")

                    try:

                        kanvas.drawString(12.80 * cm, 16.00 * cm, f"{tabelle_werten[0][0]}".replace(".",","))
                        kanvas.drawString(15.75 * cm, 16.00 * cm, f"{tabelle_werten[0][1]}".replace(".",","))

                        kanvas.setFillColorRGB(1, 0, 0)

                        kanvas.drawString(12.80 * cm, 15.50 * cm, f"{tabelle_werten[1][0]}".replace(".",","))
                        kanvas.drawString(15.75 * cm, 15.50 * cm, f"{tabelle_werten[1][1]}".replace(".",","))

                        kanvas.setFillColorRGB(0, 0, 1)

                        kanvas.drawString(12.80 * cm, 15.00 * cm, f"{tabelle_werten[2][0]}".replace(".",","))
                        kanvas.drawString(15.75 * cm, 15.00 * cm, f"{tabelle_werten[2][1]}".replace(".",","))

                        kanvas.setFillColorRGB(0, 1, 0)
                        kanvas.drawString(12.80 * cm, 14.57 * cm, f"{tabelle_werten[3][0]}".replace(".",","))
                        kanvas.drawString(15.75 * cm, 14.57 * cm, f"{tabelle_werten[3][1]}".replace(".",","))

                        kanvas.setFillColorRGB(0, 0, 0)


                    except IndexError:
                        pass
                    # kanvas.rect(12.7*cm,14.6*cm,6*cm,4*cm)
                    # kanvas.line(12.7*cm,16.6*cm,12.7*cm,18.7)

                    # rect = Rect(12.7*cm,16.7*cm,2*cm,2*cm)
                    try:
                        import os
                        data_path = os.getcwd()

                        kanvas.drawImage(data_path + "\\Grafiken\\" + f"{proben_nummer}" + ".png", 0.2 * cm, 0.2 * cm,
                                         width=20 * cm, height=13 * cm)
                    except OSError:
                        pass
                    kanvas.showPage()

                elif len(probe_position(versuch_list, index)) == 1 and probe_position(versuch_list, index)[0] != "nan":
                    pdfmetrics.registerFont(TTFont('Arial', 'Fonts/arial.ttf'))
                    # kanvas.setFont("Vera",12)
                    kanvas.setFont("Arial", 12)
                    kanvas.drawString(8.54 * cm, y1, "REKAL Halde,")

                    position = probe_position(versuch_list, index)
                    versuch_block = versuchblock(position)
                    H = h_nummerierung_(versuch_list, block_ort)
                    probeentnahme_termin = Probeentnahmetermin(versuch_list, block_ort)
                    m = meter(versuch_list, block_ort, index)
                    datum = entnahmedatum(versuch_list, block_ort)
                    sondierdatum = sondier_datum(versuch_list, block_ort, index, proben_nummer)
                    wasser_gehalt = wassergehalt(versuch_list, block_ort, index)
                    feucht_dichte = feuchtdichte(versuch_list, block_ort, index)
                    trocken_dichte = trockendichte(versuch_list, block_ort, index)
                    nummer_versuchblock = nummerierung_versuchblock(versuch_list,block_ort,int(proben_nummer))

                    kanvas.drawString(x1, y1, position[0])
                    kanvas.drawString(6.3 * cm, 23.0 * cm, "Ergebnisse der Laborflügelsondierungen")

                    kanvas.drawString(17.98 * cm, 25.83 * cm, H)
                    kanvas.drawString(16 * cm, 25.83 * cm, "Anhang :")
                    kanvas.drawString(18.8 * cm, 25.83 * cm, ".4")
                    kanvas.drawString(19.2 * cm, 25.83 * cm, versuch_block)
                    kanvas.drawString(19.6 * cm, 25.83 * cm, nummer_versuchblock)

                    kanvas.drawString(16 * cm, 25.49 * cm, f"zu Az.:   01/00-{probeentnahme_termin}")
                    kanvas.setFillColorRGB(1, 0, 0)

                    kanvas.drawString(2.5 * cm, 21.75 * cm, f"Probe :  {proben_nummer}, obere Probenseite ")
                    kanvas.setFillColorRGB(0, 0, 0)

                    kanvas.drawString(2.5 * cm, 21.23 * cm,
                                      f"Entnahmestelle : {position[0]}, rd. {m} m von der Böschungskante entfernt")
                    kanvas.drawString(2.5 * cm, 20 * cm, "Bodenart : Rekal + Asche")
                    kanvas.drawString(2.5 * cm, 19.5 * cm, f"Entnahmedatum :   {datum}")
                    kanvas.drawString(2.5 * cm, 19 * cm, f"Sondierdatum     :   {sondierdatum}   ")
                    kanvas.drawString(2.5 * cm, 18 * cm, f"Wassergehalt  :   {wasser_gehalt}   ")
                    kanvas.drawString(2.5 * cm, 17.5 * cm, f"Feuchtdichte   :   {feucht_dichte}   ")
                    kanvas.drawString(2.5 * cm, 17.0 * cm, f"Trockendichte :   {trocken_dichte}   ")
                    kanvas.drawString(2.5 * cm, 16.0 * cm, "Schergeschwindigkeit : ω  = 0,1°/s ")
                    kanvas.drawString(2.5 * cm, 15.5 * cm, "Flügelabmesungen : H/D = 25/12,5 mm ")

                    try:

                        kanvas.drawString(12.80 * cm, 16.00 * cm, f"{tabelle_werten[0][0]}".replace(".",","))
                        kanvas.drawString(15.75 * cm, 16.00 * cm, f"{tabelle_werten[0][1]}".replace(".",","))

                        kanvas.setFillColorRGB(1, 0, 0)

                        kanvas.drawString(12.80 * cm, 15.50 * cm, f"{tabelle_werten[1][0]}".replace(".",","))
                        kanvas.drawString(15.75 * cm, 15.50 * cm, f"{tabelle_werten[1][1]}".replace(".",","))

                        kanvas.setFillColorRGB(0, 0, 1)

                        kanvas.drawString(12.80 * cm, 15.00 * cm, f"{tabelle_werten[2][0]}".replace(".",","))
                        kanvas.drawString(15.75 * cm, 15.00 * cm, f"{tabelle_werten[2][1]}".replace(".",","))

                        kanvas.setFillColorRGB(0, 1, 0)
                        kanvas.drawString(12.80 * cm, 14.50 * cm, f"{tabelle_werten[3][0]}".replace(".",","))
                        kanvas.drawString(15.75 * cm, 14.50 * cm, f"{tabelle_werten[3][1]}".replace(".",","))

                        kanvas.setFillColorRGB(0, 0, 0)




                    except IndexError:
                        pass

                    # kanvas.rect(12.7*cm,14.6*cm,6*cm,4*cm)
                    # kanvas.line(12.7*cm,16.6*cm,12.7*cm,18.7)

                    # rect = Rect(12.7*cm,16.7*cm,2*cm,2*cm)
                    # grafik = Image("E:\Projekt\IGTH/test.png")
                    # grafik._restrictSize(1 * cm, 2 * cm)
                    try:
                        import os
                        data_path = os.getcwd()

                        kanvas.drawImage(data_path + "\\Grafiken\\" + f"{proben_nummer}" + ".png", 0.2 * cm, 0.2 * cm,
                                         width=20 * cm, height=13 * cm)
                    except OSError:
                        pass
                    kanvas.showPage()
            except TypeError:
                pass

    kanvas.save()

    packet.seek(0)

    canvas_pdf = PdfFileReader(packet)

    canvas_pdf.getNumPages()

    outfile = PdfFileWriter()

    for z in range(0, canvas_pdf.getNumPages()):
        page = copy.copy(layout.getPage(0))
        page.mergePage(canvas_pdf.getPage(z))

        outfile.addPage(page)

    with open(f"{proben_nummer}_LFS.pdf", "wb") as out:
        outfile.write(out)


def h_nummerierung_(versuch_list, i):
    versuch_list = versuch_list.to_numpy()
    a = 20 * i

    H = versuch_list[0, a]
    return H


def probe_position(versuch_list, j):
    a = 20 * j  # veri setini 19 luk parcaladik.
    b = 20 * j + 19  #
    position = versuch_list.iloc[:, a:b]
    position = str(versuch_list["Position"][j]).split(",")

    if len(position) == 1 and position[0] == "nan":

        return None



    elif len(position) == 1 and position[0] != "nan":

        return (position)


    else:
        return "Böschung " + str(position[0]), f"({position[1]})"


def versuchblock(position):
    if position[0] == "Plateau":
        return ".1"

    elif position[0] + position[1] == "Böschung Süd(50 m über Fuß)":
        return ".2"

    elif position[0] + position[1] == "Böschung Süd(direkt am Fuß)":
        return ".3"
    elif position[0] + position[1] == "Böschung Nord(50 m über Fuß)":
        return ".4"

    elif position[0] + position[1] == 'Böschung Nord(direkt am Fuß)':
        return ".5"


def Probeentnahmetermin(versuch_list, i):
    a = 20 * i  # veri setini 19 luk parcaladik.
    b = 20 * i + 19  #
    position = versuch_list.iloc[:, a:b]

    return position.to_numpy()[1][0]


def probennummer(versuch_list, i, j):
    a = 20 * i  # veri setini 19 luk parcaladik.
    b = 20 * i + 19  #
    versuch_list = versuch_list.to_numpy()
    versuch_list = versuch_list[:, a:b]

    if versuch_list[j][1] == "nan":
        return " "

    else:

        return str(int(versuch_list[j][1]))


def meter(versuch_list, i, j):
    a = 20 * i
    b = 20 * i + 19
    versuch_list = versuch_list.to_numpy()
    versuch_list = versuch_list[:, a:b]

    if versuch_list[j][4] == "nan":

        return " "

    else:

        return str(versuch_list[j][4]).replace(".",",")


def entnahmedatum(versuch_list, i):
    a = 20 * i

    if type(versuch_list.columns[a]) == type(" "):
        print("Datum ist gleich wie vorherige , bitte entnahme datum ändern")
    return (versuch_list.columns[a]).strftime("%d-%m-%Y")


def sondier_datum(versuch_list, i, j, probe):
    a = 20 * i
    b = 20 * i + 19
    versuch_list = versuch_list.replace()
    versuch_list = versuch_list.to_numpy()
    versuch_list = versuch_list[:, a:b]

    if type((versuch_list[j][16])) == ((type(" ") or "nan")):

        return ""  # input("Bitte die richtigen Datum für Proben nummer : "+ f"{probe}" +" eingeben " )

    else:

        return (versuch_list[j][16]).strftime("%d-%m-%Y")


def wassergehalt(versuch_list, i, j):
    a = 20 * i
    b = 20 * i + 19
    versuch_list = versuch_list.replace()
    versuch_list = versuch_list.to_numpy()
    versuch_list = versuch_list[:, a:b]

    if type((versuch_list[j][7])) == ((type(" ") or "nan")):

        return ""  # input("Bitte die richtigen Datum für Proben nummer : "+ f"{probe}" +" eingeben " )

    else:

        return str((round(versuch_list[j][7], 2))).replace(".", ",") + " %"


def trockendichte(versuch_list, i, j):
    a = 20 * i
    b = 20 * i + 19
    versuch_list = versuch_list.replace()
    versuch_list = versuch_list.to_numpy()
    versuch_list = versuch_list[:, a:b]

    if type((versuch_list[j][9])) == ((type(" ") or "nan")):

        return ""  # input("Bitte die richtigen Datum für Proben nummer : "+ f"{probe}" +" eingeben " )

    else:

        return str((round(versuch_list[j][9],2))).replace(".", ",") + " g/cm³"


def feuchtdichte(versuch_list, i, j):
    a = 20 * i
    b = 20 * i + 19
    versuch_list = versuch_list.replace()
    versuch_list = versuch_list.to_numpy()
    versuch_list = versuch_list[:, a:b]

    if type((versuch_list[j][8])) == ((type(" ") or "nan")):

        return ""  # input("Bitte die richtigen Datum für Proben nummer : "+ f"{probe}" +" eingeben " )

    else:

        return str((round(versuch_list[j][8],2))).replace(".", ",") + " g/cm³"


def datei_suchen():
    roh_data_path = os.getcwd() + "\\roh_LFS"

    with os.scandir(roh_data_path) as suchen:
        datei_list = []
        data_typen = []
        for datei_name in suchen:
            datei_list.append(datei_name.name)

        for datei in datei_list:

            if fnmatch.fnmatch(datei, "*.1LA") or (fnmatch.fnmatch(datei, "*.2LA")) or fnmatch.fnmatch(datei,
                                                                                                       "*.3LA") or fnmatch.fnmatch(
                    datei, "*.4LA"):
                data_typen.append(datei)

            # datei_list.append(datei_name)
            # if len(data_typen) == 4 :
            #     print("**** "
            #           ""
            #           "4 Grafikdateien gefunden, bitte überprüfen die Grafikdatei erneut!!"
            #           ""
            #           "****")
    return data_typen


def grafik_datei(proben_nummer, data_typen):
    drehwinkel = {}
    spannungen = {}
    roh_data_path = os.getcwd() + "\\roh_LFS"
    for i in data_typen:

        if i.startswith(f"{proben_nummer}") == True:
            datei = open(roh_data_path + "\\" + f"{i}", "r")

            roh_data = datei.readlines()
            grafik_datei = []
            grafik_datei_2 = []

            for j in range(len(roh_data)):
                grafik_datei.append((roh_data[j][3:10].replace(",", ".")))

                grafik_datei_2.append((roh_data[j][15:21].replace(",", ".")))
            drehwinkel[f"{i}"] = grafik_datei[11:]
            spannungen[f"{i}"] = grafik_datei_2[11:]

            datei.close()

    return drehwinkel, spannungen
def max_min_function(drehwinkel_float):

    # max_min = {}
    drehwinkel_float = drehwinkel_float[::-1]
    for i in range(len(drehwinkel_float)-1):


        fehler = abs(drehwinkel_float[i]-1800)

        if fehler < 50 :


            # max_min[f"{proben}"] = drehwinkel_float[i]
            # max_min["end_wert"] = i


            return len(drehwinkel_float)-i

def grafik_darstellung(versuch_list_excel,versuch_list_name,drehwinkel, spannung, proben,x_limit,x_limit2,y_limit,y_limit_2):
    keys = drehwinkel.keys()
    keys_2 = spannung.keys()
    farben = ["k", "r", "b", "g"]
    markern = ["x", "d", "s", "x"]
    mpl.rc('font', family='Arial',size=18)
    tabelle_werten_1 = []
    tabelle_werten_2 = []

    zähler = 0
    grafik_limit_werten = []

    for i in keys:
        drehwinkel_float_werte = [float(x) for x in drehwinkel[i]]
        spannung_float_werte = [float(x) for x in spannung[i]]
        drehwinkel_end = max_min_function(drehwinkel_float_werte)

        # grafik_werten = max_min_function(drehwinkel_float_werte, proben)
        # drehwinkel_end = grafik_werten["end_wert"]
        # grafik_limit_werten.append(grafik_werten[f"{proben}"])

        plt.rcParams["figure.figsize"] = (18, 8)

        # ---------------Plot-Max-Value-------------
        plt.subplot(1, 2, 1)

        max_spannung_1 = max(spannung_float_werte[0:110])
        max_spannung_index_1 = spannung_float_werte[0:110].index(max_spannung_1)

        plt.plot(drehwinkel_float_werte[0:110][max_spannung_index_1], spannung_float_werte[0:110][max_spannung_index_1],
                 color=farben[zähler], marker=markern[zähler], markersize=9, )
        plt.text(drehwinkel_float_werte[0:110][max_spannung_index_1], spannung_float_werte[0:110][max_spannung_index_1],
                 str(max_spannung_1).replace(".",","), fontsize=18)
        tabelle_werten_1.append(max_spannung_1)

        # ---------------Plot-1-------------

        plt.plot(drehwinkel_float_werte[0:110], spannung_float_werte[0:110], color=farben[zähler],
                 label=f"Teilversuch {zähler + 1}")

        plt.ylabel(r'$\tau_{FS} $' + "  " + "[kN/m\u00b2]", fontsize=18, fontname="Arial")
        plt.xlabel(r"$\varphi$" + " " + "[°]", fontsize=25, fontname="Arial")
        plt.title("Flügelscherfestigkeit", fontsize=22, fontname="Arial")
        plt.legend(fontsize=16)
        plt.ylim(y_limit, y_limit_2)
        plt.xlim(0,50)

        # ---------------Plot-Max-Value-------------

        plt.subplot(1, 2, 2)

        max_spannung_2 = max(spannung_float_werte[drehwinkel_end:])
        max_spannung_index_2 = spannung_float_werte[drehwinkel_end:].index(max_spannung_2)

        plt.plot(drehwinkel_float_werte[drehwinkel_end:][max_spannung_index_2],
                 spannung_float_werte[drehwinkel_end:][max_spannung_index_2],
                 color=farben[zähler], marker=markern[zähler], markersize=9)
        plt.text(drehwinkel_float_werte[drehwinkel_end:][max_spannung_index_2],
                 spannung_float_werte[drehwinkel_end:][max_spannung_index_2],
                 str(max_spannung_2).replace(".",","), fontsize=18, fontname="Arial")
        tabelle_werten_2.append(max_spannung_2)

        # ---------------Plot-2-------------
        plt.plot(drehwinkel_float_werte, spannung_float_werte, color=farben[zähler],
                 label=f"Teilversuch {zähler + 1}")
        plt.ylim(0, 25)


        plt.xlim(x_limit,x_limit2)
        # plt.xlim(min(grafik_limit_werten)-5,1950)

        plt.ylabel(r'$\tau_{FS} $' + "  " + "[kN/m\u00b2]", fontsize=18, fontname="Arial")
        plt.xlabel(r"$\varphi$" + " " + "[°]", fontsize=25, fontname="Arial")
        plt.title("Restscherfestigkeit", fontsize=22, fontname="Arial")
        plt.legend(fontsize=16)
        zähler += 1
    if len(drehwinkel) == 0 or len(spannung) == 0:
        plt.close()
    else:
        try:

            data_path = os.getcwd()
            os.mkdir(data_path + "\\Grafiken")

        except FileExistsError:
            pass
        plt.savefig(data_path + "\\Grafiken" + "\\" + f"{proben}" + ".png", dpi=500)  # dpi = 500
        plt.close()




    versuch_list = read_excel(versuch_list_name, sheet_name="Festigkeiten", header=None).replace(
        nan, "nan")
    festigkeit_proben_nummer = versuch_list[::][1].to_list()

    index_nummer = festigkeit_proben_nummer.index(int(proben))
    blatt = versuch_list_excel["Festigkeiten"]

    blatt["H"][index_nummer].value = tabelle_werten_1[0]
    blatt["I"][index_nummer].value = tabelle_werten_1[1]
    blatt["J"][index_nummer].value = tabelle_werten_1[2]
    blatt["L"][index_nummer].value = tabelle_werten_2[0]
    blatt["M"][index_nummer].value = tabelle_werten_2[1]
    blatt["N"][index_nummer].value = tabelle_werten_2[2]
    versuch_list_excel.save("Neu_" + versuch_list_name)

    return tuple(zip(tabelle_werten_1, tabelle_werten_2))


def nummerierung_versuchblock(versuch_list, block_ort, proben_nummer):
    if block_ort == 0:
        block = 3

    elif block_ort == 1:
        block = 23

    versuch_list = versuch_list.to_numpy()
    platau_versuchblock = []
    sho_versuchblock = []
    shu_versuchblock = []
    nho_versuchblock = []
    nhu_versuchblock = []

    for i in range(versuch_list.shape[0]):

        if ((versuch_list[i][block] == 'Plateau') and (versuch_list[i][block+12] == "x" )) ==  True:
            platau_versuchblock.append(versuch_list[i][block - 2])
            platau_versuchblock = [x for x in platau_versuchblock if x != 'nan']

        elif (versuch_list[i][block] == 'Süd,50 m über Fuß') and (versuch_list[i][block+12] == "x" ):

            sho_versuchblock.append(versuch_list[i][block - 2])
            sho_versuchblock = [x for x in sho_versuchblock if x != 'nan']

        elif (versuch_list[i][block] == 'Süd,direkt am Fuß' ) and (versuch_list[i][block+12] == "x" ):
            shu_versuchblock.append(versuch_list[i][block - 2])
            shu_versuchblock = [x for x in shu_versuchblock if x != 'nan']

        elif (versuch_list[i][block] == 'Nord,50 m über Fuß') and (versuch_list[i][block+12] == "x" ):
            nho_versuchblock.append(versuch_list[i][block - 2])
            nho_versuchblock = [x for x in nho_versuchblock if x != 'nan']
        elif (versuch_list[i][block] == 'Nord,direkt am Fuß') and (versuch_list[i][block+12] == "x" ):
            nhu_versuchblock.append(versuch_list[i][block - 2])
            nhu_versuchblock = [x for x in nhu_versuchblock if x != 'nan']

    for x in (platau_versuchblock, shu_versuchblock, nho_versuchblock, sho_versuchblock,
              nhu_versuchblock):
        for y in x:

            if y == proben_nummer:
                return ".0"+str(x.index(y) + 1)


print("""

Das Programm bietet zwei verschiedene Möglichkeiten, die Auswertungen der Testergebnisse zu berechnen.
Aus Versuchsname oder Alle Versuchsdaten erstellen.


!!!Hinweis: Wenn das Programm eine Fehlermeldung ausgibt, muss es neu gestartet werden.!!!

!!!Testergebnisse müssen in einem stabilen Excel-Format vorliegen. !!!


!!!Änderungen sollten in der beigefügten Excel-Datei vorgenommen werden. Neue Testergebnisse sollten auf derselben Datei vorbereitet werden.!!!

1 - ) Dateiname von Versuchlist eingeben

2- ) Geben bitte Auswahl ein !!!

     Einzeln PDF zu erstellen  : 1 eingeben 

     Alle Auswertungen von einer Liste zu erstellen : 2 eingeben 

3 - ) Drücken Sie 0, um das Programm zu schließen

4-) Um X-Achse Limit zu ändern , drücken bitte 3
""")

try:
    versuch_list_name = input("Bitte Exceldatei von Versucliste eingeben !!! ")
    versuch_list = read_excel(versuch_list_name, sheet_name="Uebersicht").replace(
        nan, "nan")

except FileNotFoundError:
    print("Excel-Datei nicht gefunden, überprüft mal bitte den Dateinamen."


          "Bitte starten das Programm neu"

          )

try:

    layout_path = os.getcwd()

    layout = PdfFileReader(f"{layout_path}" + "\\" + "Layout_LFS.pdf", "rb")


except FileNotFoundError:
    print("Layout-Datei nicht gefunden, überprüft mal bitte den Dateinamen."
          "Bitte starten das Programm neu"

          ""
          "")




x_limit = 1850
x_limit_2 = 1960
y_limit = 0
y_limit_2 = 100

while True:
    auswahl = int(input(" Auswahl :  "))

    if type(auswahl) == type(" ") or auswahl > 3 or auswahl < 0:

        print("Es wurde ein ungültiger Wert eingegeben, bitte versuchen Sie es erneut.")

    else:

        if auswahl == 1:

            block_ort = int(input("Bitte Blocknummer von Probennummer eingeben 1. oder 2. Block")) - 1

            if type(block_ort) == type(" ") or block_ort > 1 or block_ort < 0 or type(block_ort) == type(10.25):
                print("Es wurde ein ungültiger Wert eingegeben, bitte versuchen Sie es erneut.")
            else:
                if block_ort == 0:

                    try:
                        k = versuch_list["Probennummer"].to_list()

                        index_nummer = k.index(int(input("Bitte Probennummer eingeben !")))
                        haupt_funktion2(versuch_list, layout, index_nummer, block_ort,x_limit,x_limit_2,y_limit,y_limit_2)
                        print("Prozess erfolgreich abgeschlossen !!!")

                        print(""" * * Drücken Sie 0, um das Programm zu schließen! 

                                  * * einen neuen PDF zu erstellen 1 drücken 


                              """)
                    except (FileExistsError, ValueError):

                        print("ungültiger Prozess")

                else:
                    try:
                        k = versuch_list["Probennummer.1"].to_list()
                        index_nummer = k.index(int(input("Bitte Probennummer eingeben !")))
                        a = haupt_funktion2(versuch_list, layout, index_nummer, block_ort,x_limit,x_limit_2,y_limit,y_limit_2)

                        print("""   

                        *************
                        Prozess erfolgreich abgeschlossen !!!
                        *************

                        """
                              )

                        print(""" * * Drücken Sie 0, um das Programm zu schließen! 

                                  * * einen neuen PDF zu erstellen 1 drücken 


                              """)

                    except (FileExistsError, ValueError, NameError):

                        print("ungültiger Prozess")

        elif auswahl == 2:
            try:
                block_ort = int(input("Bitte Blocknummer von Probennummer eingeben 1. oder 2. Block")) - 1

                if block_ort == 0 :
                    block_ort_lfs = 15
                elif block_ort == 1 :
                    block_ort_lfs = 35

                haupt_funktion(versuch_list,versuch_list_name,layout,x_limit,x_limit_2,block_ort_lfs,block_ort,y_limit,y_limit_2)

                print("""   

                *************
                Prozess erfolgreich abgeschlossen !!!
                *************

                """
                      )

                print(""" * * Drücken Sie 0, um das Programm zu schließen! 

                          * * einen neuen PDF zu erstellen 1 drücken 


                      """)
            except (PermissionError,NameError):

                print("Bitte Excel-Datei schlisessen oder Dateiname kontrolieren")



        elif auswahl == 3 :

            try:

                x_limit = int(input("Bitte X-Achse Anfang-Limit eingeben (Vorgegebene Wert : 1850)"))
                x_limit_2 = int(input("Bitte X-Achse End-Limit eingeben (Vorgegebene Wert : 1950)"))
                y_limit = int(input("Bitte Y-Achse Anfang-Limit eingeben (Vorgegebene Wert : 0)"))
                y_limit2 = int(input("Bitte Y-Achse End-Limit eingeben (Vorgegebene Wert : 100)"))
            except NameError:

                print("Falsche Eingabe")

        elif auswahl == 0:

            print("Skript erfolgreich abgeschlossen !!!"
                  ""
                  "Tschüss")
            break





# h_nummerierung_(versuch_list, 1)
# f = datei_suchen()
#
# c ,d = grafik_datei(43101,f)
#
# h = grafik_darstellung(c,d,43101)

# versuch_list["Probennummer"].to_list()
# k.index(43090)

