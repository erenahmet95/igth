

import copy

import matplotlib.pyplot as plt
import numpy
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch, cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.platypus import Frame, Image
from pandas import read_excel
from numpy import nan
import io
import fnmatch
from reportlab.graphics.shapes import Rect,Line
from matplotlib.transforms import Affine2D
import matplotlib.pyplot as plt
import matplotlib.ticker as plticker


# dpl_liste = read_excel("E:\Projekt\IGTH\DPL\DPL_09-11-2021.xlsx", sheet_name="Tabelle1").replace(nan,"nan")
# dpl_liste_2 = dpl_liste.to_numpy()
# dpl_liste3 = read_excel("E:\Projekt/IGTH/DPL/Versuche_2021_az0100_40-41_H35-36_v03.xlsx", sheet_name="DPL").replace(nan,"nan")

# x_werte = []
# x_werte2 = []
# x_werte3 = []
#
# for i in dpl_liste_2[2::,4] :
#
#     if type(i) == type (" ") :
#         pass
#
#     else :
#
#         x_werte.append(i)
#
# for i in dpl_liste_2[2::,5] :
#
#     if type(i) == type (" ") :
#         pass
#
#     else :
#
#         x_werte2.append(i)
#
# for i in dpl_liste_2[2::,6] :
#
#     if type(i) == type (" ") :
#         pass
#
#     else :
#
#         x_werte3.append(i)
#
# y = range(0,len(x_werte)*10,10)
# plt.rcParams["figure.figsize"] = (5, 18)
# fig, ax = plt.subplots()
#
# ax.step(x_werte2,range(0,len(x_werte2)*10,10))
# ax.step(x_werte3,range(0,len(x_werte3)*10,10))
#
# ax.step(x_werte,y,linewidth = 2.5)
#
#
# secax = ax.secondary_xaxis('top')
# secax.set_xlabel('N_10 [-]')
# # loc = plticker.MultipleLocator(base=5) # this locator puts ticks at regular intervals
# # ax.xaxis.set_major_locator(loc)
#
# plt.yticks((range(0,200,20)))
# plt.ylim(0,200)
# plt.xlim(0,50)
# plt.gca().invert_yaxis()
# plt.gca().axes.get_xaxis().set_visible(False)
#
# plt.annotate("Abbruch nach 30 Schlägen bis 0,50 m",
#              xy=(31, 80), xycoords='data',
#              xytext=(31.5,32), textcoords='offset points', ha="center", fontsize=10,
#              arrowprops=dict(facecolor='black', shrink=0.01))
#
#
# plt.annotate("Abbruch nach 30 Schlägen bis 0,50 m",
#              xy=(31, 180), xycoords='data',
#              xytext=(31.5,31), textcoords='offset points', ha="center", fontsize=10,
#              arrowprops=dict(facecolor='black', shrink=0.01))
#
# plt.annotate("Abbruch nach 30 Schlägen bis 0,50 m",
#              xy=(30, 160), xycoords='data',
#              xytext=(31.5,31), textcoords='offset points', ha="center", fontsize=10,
#              arrowprops=dict(facecolor='black', shrink=0.01))
# # plt.show()
def grafik_darstellung(dpl_liste,block_ort):
    dpl_liste_array = dpl_liste.to_numpy()
        # dpl_liste = dpl_liste[::, 4:]
    plt.rcParams["figure.figsize"] = (18, 8)

    if block_ort == 1 :
        dpl_liste_array = dpl_liste_array[10:32:,::]
    elif block_ort == 2 :
        dpl_liste_array = dpl_liste_array[39::,::]


    for x in range(3) :

        if x == (0) :

            ort = "0" # "Plateau"
            dpl_block = dpl_liste_array

            dpl_werte_1 = dpl_block[2:,2]
            dpl_werte_2 = dpl_block[2:, 3]
            dpl_werte_3 = dpl_block[2:, 4]
            dpl_werte_4 = dpl_block[2:, 5]
            dpl_werte_5 = dpl_block[2:, 6]

            dpl_werte_1 = [x for x in dpl_werte_1 if x != 'nan']
            dpl_werte_1.append(dpl_werte_1[-1])
            dpl_werte_2 = [x for x in dpl_werte_2 if x != 'nan']
            dpl_werte_2.append(dpl_werte_2[-1])


            dpl_werte_3 = [x for x in dpl_werte_3 if x != 'nan']
            dpl_werte_3.append(dpl_werte_3[-1])

            dpl_werte_4 = [x for x in dpl_werte_4 if x != 'nan']

            dpl_werte_4.append(dpl_werte_4[-1])

            dpl_werte_5 = [x for x in dpl_werte_5 if x != 'nan']

            dpl_werte_5.append(dpl_werte_5[-1])

            fig, axes = plt.subplots(nrows=1, ncols=2)
            ax1, ax2 = axes.flatten()
            plt.setp(axes,xticks=range(0,40,5),yticks=range(0,210,10),xlim=(0,40),ylim=(0,230))
            # ax1.step(dpl_werte_1, dpl_block[2::,0])
            # ax1.step(dpl_werte_2, dpl_block[2::,0])
            # ax1.step(dpl_werte_3, dpl_block[2::,0])

            for i in dpl_werte_1 :#dpl_werte_1[:-4:-1]:

                if i >= 30 :
                    index_abbruch1 = dpl_werte_1.index(i)

                    dpl_werte_1[index_abbruch1] = 30

                    for x in range(len(dpl_werte_1[index_abbruch1::])):
                        dpl_werte_1[index_abbruch1 + x] = 30


            if (30 in dpl_werte_1) == True :

                ax1.step(dpl_werte_1, range(0, len(dpl_werte_1) * 10, 10), color="k", label="DPL-1")

                ax1.annotate("    Abbruch bei " f"{(index_abbruch1) / 10}".replace(".",",") + " m" + "\n        da " +"$N_{10}$" +"> 30",
                             xy=(0,220), xycoords='data',
                             xytext=(0, 0), textcoords='offset points', ha="left", fontsize=14, color="k"
                            )
            else :
                ax1.step(dpl_werte_1, range(0, len(dpl_werte_1) * 10, 10), color="k", label="DPL-1")



            for i in dpl_werte_2 :#dpl_werte_1[:-4:-1]:

                if i >= 30 :
                    index_abbruch2 = dpl_werte_2.index(i)
                    dpl_werte_2[index_abbruch2] = 30
                    for x in range(len(dpl_werte_2[index_abbruch2::])):
                        dpl_werte_2[index_abbruch2 + x] = 30

            if (30 in dpl_werte_2) == True :

                ax1.step(dpl_werte_2, range(0, len(dpl_werte_2) * 10, 10), color="g", label="DPL-2")


                ax1.annotate("    Abbruch bei " f"{(index_abbruch2) / 10}".replace(".",",") + " m" + "\n       da " +"$N_{10}$" +"> 30",
                             xy=(13,220), xycoords='data',
                             xytext=(0, 0), textcoords='offset points', ha="left", fontsize=14, color="g")
            else :
                ax1.step(dpl_werte_2, range(0, len(dpl_werte_2) * 10, 10), color="g", label="DPL-2")

            for i in dpl_werte_3 :#dpl_werte_1[:-4:-1]:

                if i >= 30 :
                    index_abbruch3 = dpl_werte_3.index(i)
                    dpl_werte_3[index_abbruch3] = 30
                    for x in range(len(dpl_werte_3[index_abbruch3::])):
                        dpl_werte_3[index_abbruch3 + x] = 30


            if (30 in dpl_werte_3) == True :
                ax1.step(dpl_werte_3, range(0, len(dpl_werte_3) * 10, 10), color="r", label="DPL-3")


                ax1.annotate("    Abbruch bei " f"{(index_abbruch3) / 10}".replace(".",",") + " m" + "\n        da " +"$N_{10}$" +"> 30",
                             xy=(26,220), xycoords='data',
                             xytext=(0, 0), textcoords='offset points', ha="left", fontsize=14, color="r")
            else :

                ax1.step(dpl_werte_3, range(0, len(dpl_werte_3) * 10, 10), color="r", label="DPL-3")




            ax1.legend(fontsize=20,framealpha=1)
            ax1.set_ylabel("t [cm]",fontsize=18)

            secax1 = ax1.secondary_xaxis('top')

            secax1.set_xlabel(r'$N_{10} [-]$' ,fontsize =18)
            secax1.set_xticks(range(0,50,5))
            secax1.set_xlim([0,50])
            # ax1.set_yticks((range(0, 210, 10)))
            # plt.yticks((range(0, 210, 10)))
            # plt.setp(ax1, xticks=range(0,50,5))
            # plt.yticks((range(0, 200, 20)))
            # plt.ylim(0, 220)
            # plt.xlim(0, 50)
            # plt.gca().invert_yaxis()
            ax1.invert_yaxis()
            # plt.gca().axes.get_xaxis().set_visible(False)
            ylim = ax1.get_ylim()
            for x in ax1.get_xticks():
                ax1.plot((x, x), ylim, ls="--", color="gray", lw=1,alpha=0.5)
            ax1.set_ylim(ylim)
            ax1.xaxis.set_visible(False) # Um X Achse zu entfernen


            for i in dpl_werte_4 :#dpl_werte_1[:-4:-1]:

                if i >= 30 :
                    index_abbruch4 = dpl_werte_4.index(i)

                    dpl_werte_4[index_abbruch4] = 30

                    for x in range(len(dpl_werte_4[index_abbruch4::])):
                        dpl_werte_4[index_abbruch4 + x] = 30


            if (30 in dpl_werte_4) == True :

                ax2.step(dpl_werte_4, range(0, len(dpl_werte_4) * 10, 10), color="k", label="DPL-4")

                ax2.annotate("    Abbruch bei " f"{(index_abbruch4) / 10}".replace(".",",") + " m" + "\n        da " +"$N_{10}$" +"> 30",
                             xy=(0,220), xycoords='data',
                             xytext=(0, 0), textcoords='offset points', ha="left", fontsize=14, color="k"
                            )
            else :
                ax2.step(dpl_werte_4, range(0, len(dpl_werte_4) * 10, 10), color="k", label="DPL-4")



            for i in dpl_werte_5 :#dpl_werte_1[:-4:-1]:

                if i >= 30 :
                    index_abbruch5 = dpl_werte_5.index(i)

                    dpl_werte_5[index_abbruch5] = 30

                    for x in range(len(dpl_werte_5[index_abbruch5::])):
                        dpl_werte_5[index_abbruch5 + x] = 30


            if (30 in dpl_werte_5) == True :

                ax2.step(dpl_werte_5, range(0, len(dpl_werte_5) * 10, 10), color="b", label="DPL-5")


                ax2.annotate("    Abbruch bei " f"{(index_abbruch5) / 10}".replace(".",",") + " m" + "\n        da " +"$N_{10}$" +"> 30",
                             xy=(15,220), xycoords='data',
                             xytext=(0, 0), textcoords='offset points', ha="left", fontsize=14, color="b"
                            )
            else :
                ax2.step(dpl_werte_5, range(0, len(dpl_werte_5) * 10, 10), color="b", label="DPL-5")


            secax2 = ax2.secondary_xaxis('top')
            secax2.set_xlabel(r'$N_{10} [-]$' ,fontsize =18)
            secax2.set_xticks(range(0, 50, 5))
            secax2.set_xlim([0,50])
            ax2.legend(fontsize=20,framealpha=1)
            ax2.set_ylabel("t [cm]",fontsize=18)
            # ax2.set_yticks((range(0, 210, 10)))

            # plt.yticks((range(0, 210, 10)))
            # plt.ylim(0, 220)
            # plt.xlim(0, 50)
            ax2.invert_yaxis()
            ylim = ax2.get_ylim()
            for x in ax2.get_xticks():
                ax2.plot((x, x), ylim, ls="--", color="gray", lw=1,alpha=0.5)
            ax2.set_ylim(ylim)
            # plt.gca().invert_yaxis()
            plt.gca().axes.get_xaxis().set_visible(False)
            plt.gca().axes.get_yaxis().set_visible(True)

            import os
            try:

                data_path = os.getcwd()
                os.mkdir(data_path + "\\Grafik")

            except FileExistsError:
                pass
            plt.savefig(data_path + "\\Grafik" + "\\" + ort + ".png",dpi = 500)  # dpi = 500

            plt.close()
            # plt.show()





        elif x == (1) :
            ort = "1" #"Südhang oben"
            dpl_block = dpl_liste_array

            dpl_werte_1 = dpl_block[2:, 7]
            dpl_werte_2 = dpl_block[2:, 8]
            dpl_werte_4 = dpl_block[2:, 9]
            dpl_werte_5 = dpl_block[2:, 10]

            dpl_werte_1 = [x for x in dpl_werte_1 if x != 'nan']
            dpl_werte_1.append(dpl_werte_1[-1])

            dpl_werte_2 = [x for x in dpl_werte_2 if x != 'nan']
            dpl_werte_2.append(dpl_werte_2[-1])


            dpl_werte_4 = [x for x in dpl_werte_4 if x != 'nan']
            dpl_werte_4.append(dpl_werte_4[-1])


            dpl_werte_5 = [x for x in dpl_werte_5 if x != 'nan']
            dpl_werte_5.append(dpl_werte_5[-1])



            fig, axes = plt.subplots(nrows=1, ncols=2)
            ax1, ax2 = axes.flatten()
            plt.setp(axes,xticks=range(0,40,5),yticks=range(0,210,10),xlim=(0,40),ylim=(0,230))
            # ax1.step(dpl_werte_1, dpl_block[2::,0])
            # ax1.step(dpl_werte_2, dpl_block[2::,0])
            # ax1.step(dpl_werte_3, dpl_block[2::,0])

            for i in dpl_werte_1:  # dpl_werte_1[:-4:-1]:

                if i >= 30:
                    index_abbruch1 = dpl_werte_1.index(i)

                    dpl_werte_1[index_abbruch1] = 30

                    for x in range(len(dpl_werte_1[index_abbruch1::])):
                        dpl_werte_1[index_abbruch1 + x] = 30


            if (30 in dpl_werte_1) == True:

                ax1.step(dpl_werte_1, range(0, len(dpl_werte_1) * 10, 10), color="k", label="DPL-6")

                ax1.annotate("    Abbruch bei " f"{(index_abbruch1) / 10}".replace(".",",") + " m" + "\n        da " +"$N_{10}$" +"> 30",
                             xy=(0,220), xycoords='data',
                             xytext=(0, 0), textcoords='offset points', ha="left", fontsize=14, color="k"
                            )
            else:
                ax1.step(dpl_werte_1, range(0, len(dpl_werte_1) * 10, 10), color="k", label="DPL-6")

            for i in dpl_werte_2:  # dpl_werte_1[:-4:-1]:

                if i >= 30:
                    index_abbruch2 = dpl_werte_2.index(i)
                    dpl_werte_2[index_abbruch2] = 30
                    for x in range(len(dpl_werte_2[index_abbruch2::])):
                        dpl_werte_2[index_abbruch2 + x] = 30


            if (30 in dpl_werte_2) == True:

                ax1.step(dpl_werte_2, range(0, len(dpl_werte_2) * 10, 10), color="g", label="DPL-7")

                ax1.annotate("    Abbruch bei " f"{(index_abbruch2) / 10}".replace(".",",") + " m" + "\n        da " +"$N_{10}$" +"> 30",
                             xy=(13,220), xycoords='data',
                             xytext=(0, 0), textcoords='offset points', ha="left", fontsize=14, color="g")
            else:
                ax1.step(dpl_werte_2, range(0, len(dpl_werte_2) * 10, 10), color="g", label="DPL-7")

            ax1.legend(fontsize=20,framealpha=1)
            ax1.set_ylabel("t [cm]",fontsize = 18)

            secax1 = ax1.secondary_xaxis('top')

            secax1.set_xlabel(r'$N_{10} [-]$' ,fontsize =18)
            secax1.set_xticks(range(0, 50, 5))
            secax1.set_xlim([0, 50])
            # ax1.set_yticks((range(0, 210, 10)))
            # plt.yticks((range(0, 210, 10)))
            # plt.setp(ax1, xticks=range(0,50,5))
            # plt.yticks((range(0, 200, 20)))
            # plt.ylim(0, 220)
            # plt.xlim(0, 50)
            # plt.gca().invert_yaxis()
            ax1.invert_yaxis()
            ylim = ax1.get_ylim()
            for x in ax1.get_xticks():
                ax1.plot((x, x), ylim, ls="--", color="gray", lw=1,alpha=0.5)
            ax1.set_ylim(ylim)
            # plt.gca().axes.get_xaxis().set_visible(False)
            ax1.xaxis.set_visible(False)  # Um X Achse zu entfernen


            for i in dpl_werte_4 :#dpl_werte_1[:-4:-1]:

                if i >= 30 :
                    index_abbruch4 = dpl_werte_4.index(i)

                    dpl_werte_4[index_abbruch4] = 30

                    for x in range(len(dpl_werte_4[index_abbruch4::])):
                        dpl_werte_4[index_abbruch4 + x] = 30


            if (30 in dpl_werte_4) == True :

                ax2.step(dpl_werte_4, range(0, len(dpl_werte_4) * 10, 10), color="k", label="DPL-8")

                ax2.annotate("    Abbruch bei " f"{(index_abbruch4) / 10}".replace(".",",") + " m" + "\n        da " +"$N_{10}$" +"> 30",
                             xy=(0,220), xycoords='data',
                             xytext=(0, 0), textcoords='offset points', ha="left", fontsize=14, color="k"
                            )
            else :
                ax2.step(dpl_werte_4, range(0, len(dpl_werte_4) * 10, 10), color="k", label="DPL-8")



            for i in dpl_werte_5 :#dpl_werte_1[:-4:-1]:

                if i >= 30 :
                    index_abbruch5 = dpl_werte_5.index(i)

                    dpl_werte_5[index_abbruch5] = 30

                    for x in range(len(dpl_werte_5[index_abbruch5::])):
                        dpl_werte_5[index_abbruch5 + x] = 30


            if (30 in dpl_werte_5) == True :

                ax2.step(dpl_werte_5, range(0, len(dpl_werte_5) * 10, 10), color="b", label="DPL-9")

                ax2.annotate("    Abbruch bei " f"{(index_abbruch5) / 10}".replace(".",",") + " m" + "\n        da " +"$N_{10}$" +"> 30",
                             xy=(15,220), xycoords='data',
                             xytext=(0, 0), textcoords='offset points', ha="left", fontsize=14, color="b"
                            )
            else :
                ax2.step(dpl_werte_5, range(0, len(dpl_werte_5) * 10, 10), color="b", label="DPL-9")


            secax2 = ax2.secondary_xaxis('top')
            secax2.set_xlabel(r'$N_{10} [-]$' ,fontsize =18)
            secax2.set_xticks(range(0, 50, 5))
            secax2.set_xlim([0,50])
            ax2.legend(fontsize=20,framealpha=1)
            ax2.set_ylabel("t [cm]",fontsize =18)
            # ax2.set_yticks((range(0, 210, 10)))

            # plt.yticks((range(0, 210, 10)))
            # plt.ylim(0, 220)
            # plt.xlim(0, 50)
            ax2.invert_yaxis()
            ylim = ax2.get_ylim()
            for x in ax2.get_xticks():
                ax2.plot((x, x), ylim, ls="--", color="gray", lw=1,alpha=0.5)
            ax2.set_ylim(ylim)
            # plt.gca().invert_yaxis()
            plt.gca().axes.get_xaxis().set_visible(False)
            plt.gca().axes.get_yaxis().set_visible(True)

            import os
            try:

                data_path = os.getcwd()
                os.mkdir(data_path + "\\Grafik")

            except FileExistsError:
                pass
            plt.savefig(data_path + "\\Grafik" + "\\" + ort + ".png",dpi=500)  # dpi = 500

            plt.close()

        elif x == (2) :

            ort = "2" #"Nordhang oben"
            dpl_block = dpl_liste_array
            dpl_werte_1 = dpl_block[2:, 11]
            dpl_werte_2 = dpl_block[2:, 12]
            dpl_werte_4 = dpl_block[2:, 13]
            dpl_werte_5 = dpl_block[2:, 14]

            dpl_werte_1 = [x for x in dpl_werte_1 if x != 'nan']
            dpl_werte_1.append(dpl_werte_1[-1])


            dpl_werte_2 = [x for x in dpl_werte_2 if x != 'nan']
            dpl_werte_2.append(dpl_werte_2[-1])


            dpl_werte_4 = [x for x in dpl_werte_4 if x != 'nan']

            dpl_werte_4.append(dpl_werte_4[-1])


            dpl_werte_5 = [x for x in dpl_werte_5 if x != 'nan']

            dpl_werte_5.append(dpl_werte_5[-1])


            fig, axes = plt.subplots(nrows=1, ncols=2)
            ax1, ax2 = axes.flatten()
            plt.setp(axes,xticks=range(0,40,5),yticks=range(0,210,10),xlim=(0,40),ylim=(0,230))
            # ax1.step(dpl_werte_1, dpl_block[2::,0])
            # ax1.step(dpl_werte_2, dpl_block[2::,0])
            # ax1.step(dpl_werte_3, dpl_block[2::,0])

            for i in dpl_werte_1:  # dpl_werte_1[:-4:-1]:

                if i >= 30:
                    index_abbruch1 = dpl_werte_1.index(i)

                    dpl_werte_1[index_abbruch1] = 30

                    for x in range(len(dpl_werte_1[index_abbruch1::])):
                        dpl_werte_1[index_abbruch1 + x] = 30

            if (30 in dpl_werte_1) == True:

                ax1.step(dpl_werte_1, range(0, len(dpl_werte_1) * 10, 10), color="k", label="DPL-10")

                ax1.annotate("    Abbruch bei " f"{(index_abbruch1) / 10}".replace(".",",") + " m" + "\n        da " +"$N_{10}$" +"> 30",
                             xy=(0,220), xycoords='data',
                             xytext=(0, 0), textcoords='offset points', ha="left", fontsize=14, color="k"
                            )
            else:
                ax1.step(dpl_werte_1, range(0, len(dpl_werte_1) * 10, 10), color="k", label="DPL-10")

            for i in dpl_werte_2:  # dpl_werte_1[:-4:-1]:

                if i >= 30:
                    index_abbruch2 = dpl_werte_2.index(i)
                    dpl_werte_2[index_abbruch2] = 30
                    for x in range(len(dpl_werte_2[index_abbruch2::])):
                        dpl_werte_2[index_abbruch2 + x] = 30



            if (30 in dpl_werte_2) == True:

                ax1.step(dpl_werte_2, range(0, len(dpl_werte_2) * 10, 10), color="g", label="DPL-11")


                ax1.annotate("    Abbruch bei " f"{(index_abbruch2) / 10}".replace(".",",") + " m" + "\n        da " +"$N_{10}$" +"> 30",
                             xy=(13,220), xycoords='data',
                             xytext=(0, 0), textcoords='offset points', ha="left", fontsize=14, color="g")
            else:
                ax1.step(dpl_werte_2, range(0, len(dpl_werte_2) * 10, 10), color="g", label="DPL-11")

            ax1.legend(fontsize=20,framealpha=1)
            ax1.set_ylabel("t [cm]",fontsize=18)

            secax1 = ax1.secondary_xaxis('top')

            secax1.set_xlabel(r'$N_{10} [-]$',fontsize=18)
            secax1.set_xticks(range(0, 50, 5))
            secax1.set_xlim([0, 50])
            # ax1.set_yticks((range(0, 210, 10)))
            # plt.yticks((range(0, 210, 10)))
            # plt.setp(ax1, xticks=range(0,50,5))
            # plt.yticks((range(0, 200, 20)))
            # plt.ylim(0, 220)
            # plt.xlim(0, 50)
            # plt.gca().invert_yaxis()
            ax1.invert_yaxis()

            ylim = ax1.get_ylim()
            for x in ax1.get_xticks():
                ax1.plot((x, x), ylim, ls="--", color="gray", lw=1,alpha=0.5)
            ax1.set_ylim(ylim)
            # plt.gca().axes.get_xaxis().set_visible(False)
            ax1.xaxis.set_visible(False)  # Um X Achse zu entfernen


            for i in dpl_werte_4 :#dpl_werte_1[:-4:-1]:

                if i >= 30 :
                    index_abbruch4 = dpl_werte_4.index(i)

                    dpl_werte_4[index_abbruch4] = 30

                    for x in range(len(dpl_werte_4[index_abbruch4::])):
                        dpl_werte_4[index_abbruch4 + x] = 30


            if (30 in dpl_werte_4) == True :

                ax2.step(dpl_werte_4, range(0, len(dpl_werte_4) * 10, 10), color="k", label="DPL-12")


                ax2.annotate("    Abbruch bei " f"{(index_abbruch4) / 10}".replace(".",",") + " m" + "\n        da " +"$N_{10}$" +"> 30",
                             xy=(0,220), xycoords='data',
                             xytext=(0, 0), textcoords='offset points', ha="left", fontsize=14, color="k")
            else :
                ax2.step(dpl_werte_4, range(0, len(dpl_werte_4) * 10, 10), color="k", label="DPL-12")



            for i in dpl_werte_5 :#dpl_werte_1[:-4:-1]:

                if i >= 30 :
                    index_abbruch5 = dpl_werte_5.index(i)

                    dpl_werte_5[index_abbruch5] = 30

                    for x in range(len(dpl_werte_5[index_abbruch5::])):
                        dpl_werte_5[index_abbruch5 + x] = 30


            if (30 in dpl_werte_5) == True :

                ax2.step(dpl_werte_5, range(0, len(dpl_werte_5) * 10, 10), color="b", label="DPL-13")


                ax2.annotate("    Abbruch bei " f"{(index_abbruch5) / 10}".replace(".",",") + " m" + "\n        da " +"$N_{10}$" +"> 30",
                             xy=(15,220), xycoords='data',
                             xytext=(0, 0), textcoords='offset points', ha="left", fontsize=14, color="b")
            else :
                ax2.step(dpl_werte_5, range(0, len(dpl_werte_5) * 10, 10), color="b", label="DPL-13")


            secax2 = ax2.secondary_xaxis('top')
            secax2.set_xlabel(r'$N_{10} [-]$' ,fontsize =18)
            secax2.set_xticks(range(0, 50, 5))
            secax2.set_xlim([0,50])
            ax2.legend(fontsize=20,framealpha=1)
            ax2.set_ylabel("t [cm]",fontsize = 18)
            # ax2.set_yticks((range(0, 210, 10)))

            # plt.yticks((range(0, 210, 10)))
            # plt.ylim(0, 220)
            # plt.xlim(0, 50)
            ax2.invert_yaxis()
            ylim = ax2.get_ylim()
            for x in ax2.get_xticks():
                ax2.plot((x, x), ylim, ls="--", color="gray", lw=1,alpha=0.5)
            ax2.set_ylim(ylim)
            # plt.gca().invert_yaxis()
            plt.gca().axes.get_xaxis().set_visible(False)
            plt.gca().axes.get_yaxis().set_visible(True)

            import os
            try:

                data_path = os.getcwd()
                os.mkdir(data_path + "\\Grafik")

            except FileExistsError:
                pass
            plt.savefig(data_path + "\\Grafik" + "\\" + str(ort) + ".png",dpi=500)  # dpi = 500

            plt.close()






def haupt_funktion2( dpl_list,block_ort,layout):

    grafik_darstellung(dpl_list,block_ort)
    dpl_list =dpl_list.to_numpy()

    packet = io.BytesIO()
    x = 21 * cm
    y = 29.7 * cm
    seite_größe = (x, y)

    x1 = 12.18 * cm
    y1 = 23.87 * cm
    kanvas = canvas.Canvas(packet, pagesize=seite_größe)
    for a in range(1):

        for b in range(3):

            # proben_nummer = probennummer(versuch_list, block_ort, index)

            # drehwinkel, spannung = grafik_datei(proben_nummer, datei_typen)
            #
            # try:
            #     tabelle_werten = grafik_darstellung(drehwinkel, spannung, proben_nummer,x_limit,x_limit_2)
            # except IndexError:
            #     pass

            H = h_nummerierung_(dpl_list, block_ort)
            versuch_block = versuchblock(b)[0]  # Es ruft die Funktion Versuchblock auf.
            position = versuchblock(b)[1]
            anhang_teil = h_anhanteil_az(dpl_list,block_ort)
            sondierdatum =sondier_datum(dpl_list,block_ort)
            try:
                if True == True:
                    pdfmetrics.registerFont(TTFont('Arial', 'Fonts/arial.ttf'))
                    kanvas.setFont("Arial", 12)
                    # kanvas.setFont("Vera",12)

                    kanvas.drawString(9.36 * cm, y1, "REKAL Halde,")

                    kanvas.drawString(x1, y1, position)  # Koordinat von Position
                    # kanvas.drawString(9.8 * cm, 23.50 * cm, position[1])
                    kanvas.drawString(7.54 * cm, 23.0 * cm, "Ergebnisse der leichten Rammsondierungen")
                    kanvas.drawString(17.91 * cm, 25.85 * cm, H)
                    kanvas.drawString(16 * cm, 25.85 * cm, "Anhang :")
                    kanvas.drawString(18.8 * cm, 25.85 * cm, ".7")

                    kanvas.drawString(19.2 * cm, 25.85 * cm, versuch_block)

                    kanvas.drawString(16 * cm, 25.49 * cm, f"zu Az.:     {anhang_teil}")
                    kanvas.setFillColorRGB(1, 0, 0)
                    # kanvas.drawString(2.5 * cm, 21.75 * cm, f"Probe :  {proben_nummer}, obere Probenseite ")
                    # kanvas.setFillColorRGB(0, 0, 0)
                    # kanvas.drawString(2.5 * cm, 21.23 * cm,
                    #                   f"Entnahmestelle : {position[0]}, rd. {m} m von der Böschungskante entfernt")
                    kanvas.drawString(2.5 * cm, 20 * cm, "Sondierdatum : "+f" {sondierdatum}")
                    kanvas.setFillColorRGB(0, 0, 0)

                    kanvas.drawString(2.5 * cm, 19.5 * cm, f"Sondierart:" + "   " + "DPL")
                    kanvas.drawString(2.5 * cm, 19 * cm, "Spitzenquerschnitt: 10 cm² ")
                    kanvas.drawString(2.5 * cm, 18.5 * cm, "Masse: 10 kg")
                    kanvas.drawString(2.5 * cm, 18 * cm, "Fallhöhe: 50 cm")
                    kanvas.drawString(2.5 * cm, 17.5 * cm, "Fallhöhe: 50 cm")
                    kanvas.drawString(2.5 * cm, 17 * cm, "Spitzenwinkel: 90°")




                    try:
                        import os
                        data_path = os.getcwd()

                        kanvas.drawImage(data_path + "\\Grafik\\" + str(b) + ".png", 0.1 * cm, 0.1 * cm,
                                         width=22 * cm, height=15 * cm)
                    except OSError:
                        pass
                    kanvas.showPage()
            except OSError:
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

    with open(f"{sondierdatum}_DPL.pdf", "wb") as out:
        outfile.write(out)

def h_nummerierung_(dpl_list,block_ort):


    if block_ort == 1 :

        return dpl_list[6][1]

    elif block_ort == 2 :

        return  dpl_list[35][1]





def h_anhanteil_az(dpl_list,block_ort):

    if block_ort == 1:

        return dpl_list[5][1]

    elif block_ort == 2:

        return dpl_list[34][1]


def versuchblock(b):

    if b == 0 :
        return ".1" , "Plateau"

    elif b  == 1 :
        return ".2","Böschung Süd"

    elif b == 2 :
        return ".3" , "Böschung Nord"

def sondier_datum(dpl_list,block_ort):

    if block_ort == 1:

        return dpl_list[4][1].strftime("%d-%m-%Y")

    elif block_ort == 2:

        return dpl_list[33][1].strftime("%d-%m-%Y")


print("""

Skript für Auswertungen von DPL

1 - ) Dateiname von Versuchlist eingeben

2- ) Geben bitte Auswahl ein !!!

     Einzeln PDF zu erstellen  : 1 eingeben 

3 - ) Drücken Sie 0, um das Programm zu schließen


""")

import os

try:

    versuch_list = read_excel(input("Bitte Exceldatei von Versucliste eingeben !!! "), sheet_name="DPL").replace(
        nan, "nan")  ## nan lari degistirdik.

except FileNotFoundError:
    print("Excel-Datei nicht gefunden, überprüft mal bitte den Dateinamen."


          "Bitte starten das Programm neu"

          )

try:

    layout_path = os.getcwd()

    layout = PdfFileReader(f"{layout_path}" + "\\" + "Layout_DPL.pdf", "rb")


except FileNotFoundError:
    print("Layout-Datei nicht gefunden, überprüft mal bitte den Dateinamen."
          "Bitte starten das Programm neu"

          ""
          "")


while True:
    auswahl = int(input(" Auswahl :  "))

    if type(auswahl) == type(" ") or auswahl > 3 or auswahl < 0:

        print("Es wurde ein ungültiger Wert eingegeben, bitte versuchen Sie es erneut.")

    else:

        if auswahl == 1:

            block_ort = int(input("Bitte Blocknummer von Probennummer eingeben 1. oder 2. Block"))

            if type(block_ort) == type(" ") or block_ort > 3 or block_ort < 0 or type(block_ort) == type(10.25):
                print("Es wurde ein ungültiger Wert eingegeben, bitte versuchen Sie es erneut.")
            else:
                if block_ort == 1:

                    try:

                        haupt_funktion2(versuch_list,1,layout)

                        print("Prozess erfolgreich abgeschlossen !!!")

                        print(""" * * Drücken Sie 0, um das Programm zu schließen!

                                  * * einen neuen PDF zu erstellen 1 drücken


                              """)
                    except (FileExistsError, ValueError):

                        print("ungültiger Prozess")

                elif block_ort == 2:
                    try:
                        haupt_funktion2(versuch_list,block_ort,layout)

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

# layout = PdfFileReader("E:\Projekt\IGTH\DPL/Layout.pdf", "rb")
# haupt_funktion2(dpl_liste3,1,layout)

# dpl_liste4[12:32:,::]
# dpl_liste4[41::,::]