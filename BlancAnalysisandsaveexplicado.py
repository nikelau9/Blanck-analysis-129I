# -*- coding: utf-8 -*-
"""
Created on Thu Apr 21 12:01:55 2022

@author: nikel
"""

"""
AMS measurament data excel import and blanc analysis
"""

import datetime
from threading import Timer
import holoviews as hv
import hvplot.pandas
import panel as pn
import codecs
import webbrowser
import os
from tkinter import Tk, StringVar,DoubleVar, Entry, Y, W, X, PhotoImage, Toplevel, YES, END, Label,ANCHOR, ACTIVE, RIGHT,LEFT, Button, BOTH, TOP, BOTTOM, Frame,  messagebox, filedialog, ttk, Scrollbar, VERTICAL, HORIZONTAL, Checkbutton, Listbox, MULTIPLE
import xlsxwriter 
import pandas as pd
import math
import numpy as np
from pandas import DataFrame

###Canvas definition
frame1 = Tk()
frame1.config(bg='grey63')
frame1.geometry('650x200')
frame1.minsize(width=650, height=100)
frame1.title('AMS CNA')

###Data saving parameters
cabecero=[]
datosbl={}

###Data treatment functions

#Select data in the input file, between 2 points
def Cortardatos(fichero_input):
    fa=fichero_input.read()
    p1=fa.find('Analyses')
    p2=fa.find('# Batch')
    documento=fa[p1+9:p2-1]
    documento=documento.split("\n")
    k=1
    variables=[]
    recopilacion={}
    for lerro in documento:
        if k==1:
            cabecero=lerro.split("\t")
            for i in range(len(cabecero)):
                variables.append(cabecero[i])
            k=k+1      
        else:
            k=k+1
            datos=lerro.split("\t")
            if datos[2] not in recopilacion:
                recopilacion[datos[2]]=  [datos]#Crear funcion adquirirdatos
            else:
                listenlista=recopilacion[datos[2]]
                listenlista.append(datos)
                recopilacion[datos[2]]=listenlista
    fichero_input.close()
    return cabecero , recopilacion
#Data treatment functions
def Media(ld,index):
    batu=[]
    for zerren in ld:
        batu.append(zerren[index])
    return (np.mean(batu))
def Suma(ld,index):
    batu=0
    for zerren in ld:
        batu+=zerren[index]
    return batu  
def Desviacion(ld,index):
    batu=[]
    for zerren in ld:
        batu.append(zerren[index])
    st=np.std(batu)
    return st

#Adquire and management of diferent data (Template)
def Modificador(gakoak,ldtup):
    cabenew=['Analysis',	'Sample ID',	'Sample Description','Measurement',	'Start Time',	'Accumulated Block Time',	'Excluded Ratio',	'127I charge',	'127I time',	'127I current',	'127I particle current',	'127I rate',	'129I counts',	'129I time',	'129I rate',	'Error',	'129I/127I ratio',	'Background 129I rate','error',	'Corrected rate'	,'Error',	'Counts Corrected 129I ratio',	'Error',	'Final corrected ratio',	'Error',	'Average', 'Error',	'Average 129I counts',	'Sqrt average 129I counts',	'129I statistical error',	'129I/127I std deviation', '129I/127I rel std deviation']
    ldnew=[]
    ldtupnew={}
    sonnumeros=[]
    for elem in ldtup:
        ld=ldtup[elem] 
        for zerrenda in ld:
            zerrendanew=[None]*len(cabenew)
            for i in range(len(cabenew)):
                if i==0:
                    a=gakoak.index('Analysis')
                    zerrendanew[i]=zerrenda[a]
                    sonnumeros.append(i)
                elif i==1:
                    a=gakoak.index('Sample ID')
                    zerrendanew[i]=zerrenda[a]
                elif i==2:
                    a=gakoak.index('Sample Description')
                    zerrendanew[i]=zerrenda[a]
                elif i==3:
                    a=gakoak.index('Measurement')
                    zerrendanew[i]=zerrenda[a]
                    sonnumeros.append(i)
                elif i==4:
                    a=gakoak.index('Start Time')
                    zerrendanew[i]=zerrenda[a]
                elif i==5:
                    a=gakoak.index('Accumulated Block Time')
                    zerrendanew[i]=zerrenda[a]
                    sonnumeros.append(i)
                elif i==6:
                    a=gakoak.index('Excluded Ratio')
                    zerrendanew[i]=zerrenda[a]
                    sonnumeros.append(i)
                elif i==7:
                    a=gakoak.index('127I charge')
                    zerrendanew[i]=zerrenda[a]
                    sonnumeros.append(i)
                elif i==8:
                    a=gakoak.index('127I time')
                    zerrendanew[i]=zerrenda[a]
                    sonnumeros.append(i)
                elif i==9:
                    a=gakoak.index('127I current')
                    zerrendanew[i]=zerrenda[a]
                    sonnumeros.append(i)
                elif i==10:
                    a=gakoak.index('127I particle current')
                    zerrendanew[i]=zerrenda[a]
                    sonnumeros.append(i)
                elif i==11:
                    a=gakoak.index('127I rate')
                    zerrendanew[i]=zerrenda[a]
                    sonnumeros.append(i)
                elif i==12:
                    a=gakoak.index('129I counts')
                    zerrendanew[i]=zerrenda[a]
                    sonnumeros.append(i)
                elif i==13:
                    a=gakoak.index('129I time')
                    zerrendanew[i]=zerrenda[a]
                    sonnumeros.append(i)
                elif i==14:
                    a=gakoak.index('129I rate')
                    zerrendanew[i]=zerrenda[a]
                    sonnumeros.append(i)
                elif i==15: #Error
                    b=gakoak.index('129I counts')
                    c=gakoak.index('129I time')
                    zerrendanew[i]=math.sqrt(float(zerrenda[b]))/float(zerrenda[c])
                    sonnumeros.append(i)
                elif i==16:
                    a=gakoak.index('129I/127I ratio')
                    zerrendanew[i]=zerrenda[a]
                    sonnumeros.append(i)
                elif i==17: #Background 129I rate
                    zerrendanew[i]=float(Bck.get())
                    sonnumeros.append(i)
                elif i==18: #error
                    zerrendanew[i]=float(Bckerr.get())
                    sonnumeros.append(i)
                elif i==19: #Corrected rate
                    b=gakoak.index('129I rate')
                    zerrendanew[i]=float(zerrenda[b])-float(Bck.get())
                    sonnumeros.append(i)
                elif i==20: #Error
                    b=gakoak.index('129I counts')
                    c=gakoak.index('129I time')
                    cb=(math.sqrt(float(zerrenda[b])))/float(zerrenda[c])
                    aa=float(Bckerr.get())
                    zerrendanew[i]=math.sqrt((cb)**(2) + (aa)**(2))
                    sonnumeros.append(i)   
                elif i==21: #Counts Corrected 129I ratio
                    b=gakoak.index('129I rate')
                    a=gakoak.index('127I rate')
                    bb=float(zerrenda[b])-float(Bck.get())#(-Background 129I rate)
                    zerrendanew[i]=bb/float(zerrenda[a])
                    sonnumeros.append(i) 
                elif i==22: #Error
                    b=gakoak.index('129I counts')
                    c=gakoak.index('129I time')
                    a=gakoak.index('127I rate')
                    cb=(math.sqrt(float(zerrenda[b])))/float(zerrenda[c])
                    aa=float(Bckerr.get())
                    bb=math.sqrt((cb)**(2) + (aa)**(2))
                    zerrendanew[i]=bb/float(zerrenda[a])
                    sonnumeros.append(i)  
                elif i==23: #Final corrected ratio
                    b=gakoak.index('129I rate')
                    a=gakoak.index('127I rate')
                    bb=float(zerrenda[b])-float(Bck.get())#(-Background 129I rate)
                    zerrendanew[i]=bb/float(zerrenda[a])
                    sonnumeros.append(i) 
                elif i==24: #Error
                    b=gakoak.index('129I counts')
                    c=gakoak.index('129I time')
                    a=gakoak.index('127I rate')
                    cb=(math.sqrt(float(zerrenda[b])))/float(zerrenda[c])
                    aa=float(Bckerr.get())
                    bb=math.sqrt((cb)**(2) + (aa)**(2))
                    zerrendanew[i]=bb/float(zerrenda[a])
                    sonnumeros.append(i)  
                elif i==25: #Average
                    zerrendanew[i]=None
                    sonnumeros.append(i) 
                elif i==26: #Error
                    zerrendanew[i]=None
                    sonnumeros.append(i) 
                elif i==27: #Average 129I counts
                    b=gakoak.index('Average 129I counts')
                    zerrendanew[i]=float(zerrenda[b])
                    sonnumeros.append(i)
                elif i==28: #Sqrt average 129I counts
                    b=gakoak.index('Sqrt average 129I counts')
                    zerrendanew[i]=float(zerrenda[b])
                    sonnumeros.append(i)
                elif i==29: #129I statistical error
                    b=gakoak.index('129I statistical error')
                    zerrendanew[i]=float(zerrenda[b])
                    sonnumeros.append(i)
                elif i==30: #129I/127I std deviation
                    b=gakoak.index('129I/127I std deviation')
                    zerrendanew[i]=float(zerrenda[b])
                    sonnumeros.append(i)
                elif i==31: #129I/127I rel std deviation
                    b=gakoak.index('129I/127I rel std deviation')
                    zerrendanew[i]=float(zerrenda[b])
                    sonnumeros.append(i)
            ldnew.append(zerrendanew)
            Avercor=Media(ldnew,cabenew.index('Final corrected ratio'))
            statErrorsuma=len(ldnew)
            StatError=Desviacion(ldnew,cabenew.index('Final corrected ratio'))/(math.sqrt(statErrorsuma))
            ldnew2=[]
            for zerren in ldnew:
                zerrenew=zerren
                zerrenew[25]=Avercor
                zerrenew[26]=StatError
                ldnew2.append(zerrenew)
        ldtupnew[elem]=ldnew2
    return cabenew,ldtupnew,sonnumeros
#Create data
def datosadquisicion():
    datos_obtenidos = indicasel['text']

      
    try:
        fichero_input= open(datos_obtenidos,"r")
        cabecero, datos = Cortardatos(fichero_input)
             
    except ValueError:
        messagebox.showerror('Informacion', 'Formato incorrecto')
        return None
    
    except FileNotFoundError:
        messagebox.showerror('Informacion', 'No se encuentra \n el archivo está ')
        return None
    lili=[]
    for elem in datos:
        lili.append(elem) 
    datasort(lili)

###Buttons functions

##Select document data
def seleccionarchivo():
    indicasel['text']= filedialog.askopenfilename(initialdir ='/', 
                                            title='Selecione archivo', 
                                            filetype=(('ams files', '*.ams*'),('All files', '*.*')))

##Definition of samples selections

def closeSelected():
    ws.destroy()
def Selecteall():
    lb.select_set(0, END)
def Selvar():
    printeable = []
    listasel.clear()
    cname = lb.curselection()
    for i in cname:
        op = lb.get(i)
        printeable.append(op)
    for val in printeable:
        listasel.append(val)  
    
def datasort(cabec):
    ws.deiconify()
    x=cabec     
    for item in range(len(x)): 
    	lb.insert(END, x[item]) 
    	lb.itemconfig(item, bg="#bdc1d6") 

##Interface buttons
ws = Toplevel() 
ws.title('Selección de variables') 
ws.geometry('400x300')
ws.withdraw()
scrollbar = Scrollbar(ws)
scrollbar.pack( side = RIGHT, fill=Y)
show = Label(ws, text = "Seleccione las variables", font = ("Times", 14), padx = 10, pady = 10)
show.pack() 
lb = Listbox(ws, selectmode = "multiple",yscrollcommand = scrollbar.set)
lb.pack( side = RIGHT, fill = BOTH )
scrollbar.config( command = lb.yview )
lb.pack(padx = 10, pady = 10, expand = YES, fill = "both") 

Button(ws, text="Selecionar todos", command=Selecteall).pack(side = TOP,anchor=W, fill=X)
Button(ws, text="Realizar selección", bg='blue', command=Selvar).pack(side = TOP,anchor=W, fill=X)
Button(ws, text="Cerrar", bg='red',command=closeSelected).pack(side = TOP,anchor=W, fill=X)
listasel=[]

##Blanc data treatment
def calcublancos():
    if boton3['bg']=='yellow':
        indicacal['text']=True
        boton3['bg']='purple3'
        boton3['text']='No calcular blancos'
    else:
        indicacal['text']=False
        boton3['bg']='yellow'
        boton3['text']='Calculos blancos'
        
##Save document to Excel 
#Select the needed blanc 
def seleblanc(datos):
    newdatos={}
    for elem in datos:
        if elem in listasel:
            newdatos[elem]=datos[elem]
    return newdatos
#Save excel
def escribirexcel():
    datos_obtenidos = indicasel['text']

      
    try:
        fichero_input= open(datos_obtenidos,"r")
        cabecero, datos = Cortardatos(fichero_input)
             
    except ValueError:
        messagebox.showerror('Informacion', 'Formato incorrecto')
        return None
    
    except FileNotFoundError:
        messagebox.showerror('Informacion', 'No se encuentra \n el archivo está ')
        return None
    datosbl=seleblanc(datos)
    sonnumeros=[0,3,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20]
    if indicacal['text']==True:
        cabecero,datosbl,sonnumeros=Modificador(cabecero,datosbl)
    indicaex['text']=filedialog.asksaveasfilename(title='Selecione archivo con los números del carrusel', filetype=(('xlsx files', '*.xlsx*'),('All files', '*.*')))
    indicaex['text']=indicaex['text']+'.xlsx'
    workbook = xlsxwriter.Workbook(indicaex['text']) #LocalizaciÃ³n y nombre del Archivo que se quiere obtener xlsx
    worksheet = workbook.add_worksheet() 
    row=0
    column=0 
    for item in cabecero :   
        worksheet.write(row, column, item)    
        column += 1
    
    for laginizena in datosbl:
        archivodedatos=datosbl[laginizena]
        for lisra in archivodedatos:
            row+=1
            column=0
            for item in lisra:
                if column in sonnumeros:
                    if item==None:
                        worksheet.write(row, column, item)
                    else:
                        worksheet.write(row, column, float(item))
                else:
                    worksheet.write(row, column, item)
                column += 1
        row+=1
    workbook.close() 

###Principal canvas function definition
boton1 = Button(frame1, text= 'Seleccionar Archivo', bg='blue', command= seleccionarchivo)
boton1.grid(column = 0, row = 0, sticky='nsew', padx=10, pady=10)

boton2 = Button(frame1, text= 'Seleccionar muestras \n para cálculo de blancos', fg='white', bg='black', command= datosadquisicion)
boton2.grid(column = 1, row = 0, sticky='nsew', padx=10, pady=10)

boton3 = Button(frame1, text= 'Calculos blancos', fg='black', bg='yellow', command= calcublancos)
boton3.grid(column = 2, row = 0, sticky='nsew', padx=10, pady=10)

boton4 = Button(frame1, text= 'Guardar en excel', fg='black', bg='green2', command= escribirexcel)
boton4.grid(column = 3, row = 0, sticky='nsew', padx=10, pady=10)

Bck = DoubleVar(value=0.32116324)   
indicadist = Label(frame1, fg= 'black', bg='white', text= 'Background 129I rate' , font= ('Arial',10,'bold') )
indicadist.grid(column = 2, row = 1, sticky='nsew', padx=10, pady=10)
dist = Entry(frame1, textvariable=Bck, width=10)
dist.grid(column = 2, row = 2, sticky='nsew', padx=10, pady=10)

Bckerr = DoubleVar(value=0.050281916)   
indicadisterr = Label(frame1, fg= 'black', bg='white', text= 'Background 129I rate error' , font= ('Arial',10,'bold') )
indicadisterr.grid(column = 3, row = 1, sticky='nsew', padx=10, pady=10)
disterr = Entry(frame1, textvariable=Bckerr, width=10)
disterr.grid(column = 3, row = 2, sticky='nsew', padx=10, pady=10)

###Indicators, save some global information
indicasel = Label(frame1, fg= 'white', bg='gray26', text= '' , font= ('Arial',10,'bold') )
indicaex = Label(frame1, fg= 'white', bg='gray26', text= '' , font= ('Arial',10,'bold') )
indicacal = Label(frame1, fg= 'white', bg='gray26', text= False, font= ('Arial',10,'bold') )
frame1.mainloop()