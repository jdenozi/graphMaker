#!/usr/bin/env python3
import xlrd
import sys
import matplotlib as mpl
import matplotlib.pyplot as plt
from mpl_toolkits.axisartist.axislines import SubplotZero
import numpy as np


try:
    path=sys.argv[1]

    fichier=xlrd.open_workbook(str(path))
except:
    print("Nope")

def etiquette():
    try:
        print("Rentrer la valeur de la colonne de l'etiquette")
        col=input()
        global etiquette
        etiquette=sh.col_values(int(col))
    except:
        etiquette()

def axeX():
    try:
        print("Rentrer la colonne de la valeur des X")
        col1=input()
        global x
        x=sh.col_values(int(col1))
    except:
        axeX()
def axeY():
    try:
        print("Rentrer la colonne de la valeur des Y")
        col2=input()
        global y
        y=sh.col_values(int(col2))
    except:
        axeY()

def sheetinput():
    print("Rentrer le nom de la feuille")
    try:
        sheet=input()
        global sh
        sh=fichier.sheet_by_name(sheet)
    
    except:
        sheetinput()

def loopNameAxes():
    for value in x:
        if type(value)==str and value!="":
            x.remove(value)
            global xAxeName
            xAxeName=value
            
        if value == "":
            x.remove(value)
            loopNameAxes()
        else:
            continue

    for value2 in y:
        if type(value2)==str and value2!="":
            y.remove(value2)
            global yAxeName
            yAxeName=value2
            
        if value2 == "":
            y.remove(value2)
            loopNameAxes()
        else:
            continue
    for name in etiquette:
        if name=="":
            etiquette.remove(name)
            loopNameAxes()
        else:
            continue



def graph():
    sheetinput()
    etiquette()
    axeX()
    print(x)
    axeY()
    loopNameAxes() 
    print("Rentrer le nom du graphique")
    title=input()
    fig= plt.figure()
    ax=SubplotZero(fig, 111)
    fig.add_subplot(ax)
    plt.axhline(y=0.5, c='black')
    plt.axvline(x=0.5, c='black') 
    liste=[]
    space="\n"
    for i, txt in enumerate(etiquette):
        if x[i]+y[i] in liste:
            space=space+"\n"
            ax.annotate("\n"+txt+space, (x[i], y[i]),
                rotation=20, 
                arrowprops=dict(arrowstyle="->")).draggable()
        else:
            ax.annotate("\n"+txt+"\n", (x[i], y[i]),
                rotation=20, 
                arrowprops=dict(arrowstyle="->")).draggable()
        liste.append(x[i]+y[i])

    ax.scatter(x,y,c="black")
    ax.set_ylim(-max(y)+2,max(y)+2)
    ax.set_xlim(-max(x)+2,max(x)+2)
    ax.set_xlabel(xAxeName)
    ax.set_ylabel(yAxeName)
    print(yAxeName, xAxeName)
    ax.get_xaxis().set_minor_locator(mpl.ticker.AutoMinorLocator())
    ax.get_yaxis().set_minor_locator(mpl.ticker.AutoMinorLocator())
    ax.grid(b=True, which='major', color='grey', linewidth=1.0)
    ax.grid(b=True, which='minor', color='grey', linewidth=0.5)
    plt.title(title)
    plt.show()
graph()
