global encab
global hacer
global quien
global fecha

encab=[]
hacer=[]
quien=[]

from datetime import datetime
dif=datetime.utcnow().hour-datetime.now().hour
 

import pandas as pd
from tkinter import filedialog
from tkinter import *
import tkinter as tk
import os
import zipfile

def genera():
 fecha=fec.get()
 encab,hacer,quien=limpia()
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
 #print(consulta)
 
 cons=pd.merge(how="right",left=pacientes,right=consulta,left_on="HC",right_on="HC1")

 #print(cons)

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
 
 #IMPRIME
 alli=r"c:\users\claudio\appdata\local\programs\python\python36-32"
 archtxt=open(alli+r"\Consulta.txt","w")
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

 encab,hacer,quien=limpia()
 ##FIN IMPRIME     
      
def salir():
    vent.destroy()


def limpia():
     return [],[],[]

def abre():
 ventana= Tk()
 ventana.withdraw()
 archi=filedialog.askopenfilename()
 return archi

def imprime():
     pass
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
##genera ventana
fecha=""
vent=tk.Tk()
vent.geometry("500x200+500+300")
vent.title("Consultas x Dia")

etiq1=tk.Label(vent,text="Fecha (AAAAMMDD): ")
etiq1.pack(padx=20,pady=20,side=LEFT)

fec=tk.Entry(vent)
fec.pack(padx=20,pady=20,side=LEFT)
fecha=fec.get()

boton1=tk.Button(vent,text="aceptar", command=genera)
                 
boton1.pack(side=LEFT)

boton2=tk.Button(vent,text="Salir", command=salir)
boton2.pack(padx=20,pady=20,side=LEFT)

vent.mainloop()
##fin ventana
