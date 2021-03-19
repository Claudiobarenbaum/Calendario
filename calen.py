global encab
global hacer
global quien
global fecha
global pathos
encab=[]
hacer=[]
quien=[]
from datetime import datetime
dif=datetime.utcnow().hour-datetime.now().hour
 
patio=r"C:\Users\claudio\AppData\Local\Programs\Python\Python36-32"

import pandas as pd
from tkinter import filedialog
from tkinter import *
import tkinter as tk
import os
import zipfile

def limpia():
     return [],[],[]

def abre():
 ventana= Tk()
 ventana.withdraw()
 archi=filedialog.askopenfilename()
 return archi

def imprime():
    archtxt=open(patio+r"\Consulta.txt","w")
    for j in range(len(cons)):
        Boras=str(cons.iloc[j]['hora'])[0:2]
        minu=str(cons.iloc[j]['hora'])
        Boras=int(Boras)-dif
        Boras=str(Boras)
        minu=str(cons.iloc[j]['hora'])[2:4]
        archtxt.write("-------------------------------------------------------------------------------------\n")
        archtxt.write("!FECHA.........: " + fecha+'\t')
        archtxt.write("Hora....:  "+ Boras+":"+minu+'\t'+'\t')
        archtxt.write("Historia Clínica:"+str(cons.iloc[j]['HC1'])+'\n')
        archtxt.write('!'+'\n')
        archtxt.write("!NOMBRE........: "+str(cons.iloc[j]["Paciente"])+'\n')
        archtxt.write("!DNI...........: "+str(int(cons.iloc[j]['DNI']))+'\n')
        
        
        archtxt.write("!Teléfono......: "+str(cons.iloc[j]['tele'])+'\n')
        archtxt.write("!Teléfono......: "+str(cons.iloc[j]['telealt'])+'\n')
        archtxt.write('!'+'\n')
        archtxt.write('!'+'\n')
        archtxt.write("!Frecuencia ...: "+cons.iloc[j]["Frecuencia"]+'\n')
        archtxt.write("!----------------------------\n")
        archtxt.write("!----------------------------\n")
        archtxt.write("!Psiquiatra....: "+str(cons.iloc[j]['psiq'])+'\n')
        archtxt.write("!Psicologo.....: "+str(cons.iloc[j]['Psicolo'])+'\n')
        archtxt.write("!----------------------------\n")
        archtxt.write("!Plan Farma....: "+str(cons.iloc[j]['Plan'])+'\n')
        archtxt.write('!'+'\n')
        archtxt.write("!Diagnóstico...: "+str(cons.iloc[j]['Diag'])+'\n')
        archtxt.write("-------------------------------------------------------------------------------------\n")
        archtxt.write("=====================================================================================\n")
        archtxt.write('\n')
        archtxt.write('\n')
        
        
    archtxt.close()    
###Ventana donde ejecuta


 
###trae pacientes
archi=abre() #abre excel pacientes
pacientes=pd.read_excel(archi)
### trajo pacientes

#abre zip
archi=abre()
####FIN SELECCION ARCHIVO

#####DEZIPEADO
with zipfile.ZipFile(archi,'r')  as Todozip:
  listOfFileNames = Todozip.namelist()
   # Iterate over the file names
  for fileName in listOfFileNames:
       # Check filename endswith csv
       if fileName.endswith('.ics'):
           # Extract a single file from zip
           Todozip.extract(fileName)


    
encab,hacer,quien=limpia()
#VENTANA SOLICITA FECHA
vent=tk.Tk()
vent.geometry("380x300")
vent.title("Consultas x Dia")
etiq1=tk.Label(vent,text="Fecha : ", bg="black", fg="white")
etiq1.pack
vent.mainloop
##FIN VENTANA SOLITIA FECHA

fecha=input("Fecha de Consulta ")
while fecha!="":
 calen=open(fileName,encoding='UTF-8')
 encontro=False
 for x in calen:
    
    if   x.startswith("DTSTART:"+fecha): 
       x=x[17:21]
       encab.append(x)
       encontro=True  
       #print(x)
    if x.startswith("SUMMARY") and encontro:    
       historia=""
       primera=True
       posi=0
       for letra in x  : 
             posi=posi+1
             if letra.isdigit():
                  historia=historia+letra
                  if primera:
                     tipo=x[7:posi-3]
                  primera=False  
       if len(historia)==0:
           historia=0
       hacer.append(int(historia))            
       quien.append(tipo[0:10])
       encontro=False
         
 #print(encab)
 #print(hacer)
 #print(quien)
 consulta=pd.DataFrame({"quien": quien,"HC1":hacer,"hora": encab})
 print(consulta)
 
 cons=pd.merge(how="right",left=pacientes,right=consulta,left_on="HC",right_on="HC1")

 cons['DNI']=cons["DNI"].fillna(0)
 cons['HC']=cons["HC"].fillna(0)
 cons['Paciente']=cons["Paciente"].fillna("")
 cons['Psicolo']=cons["Psicolo"].fillna("")
 cons['Plan']=cons["Plan"].fillna("")
 cons['Frecuencia']=cons["Frecuencia"].fillna("")
 cons['psiq']=cons["psiq"].fillna("")
 cons['tele']=cons["tele"].fillna("")
 cons['telealt']=cons["telealt"].fillna(0)
 cons['telealt']=cons["telealt"].apply(lambda a: int(a))
 
 
 imprime()
 encab,hacer,quien=limpia()
 
 fecha=input("Fecha de Consulta ")
 #print(calen.read(5))
