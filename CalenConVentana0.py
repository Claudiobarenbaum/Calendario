###Importa modulos
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import pandas as pd
from tkinter import filedialog
from tkinter import *
from tkinter import messagebox
from datetime import datetime
import tkinter as tk
import os
import zipfile

class impresor:
    def _init_(self):
        self.junta=True
    def verdad(self):
        self.junta=True
    def falsif(self):
        self.junta=False
    def valor(self):
       return self.junta

class ruta:
    def _init_(self):
        self.rut=""
    def set(self,donde):
        self.rut=donde
    def valor(self):
        return self.rut
 
###Define e incializa variables 
#global encab
#global hacer
#global quien
#global fecha


tipoxl=sys.argv[3]

if tipoxl=='xls':
   import xlrd
else:
   import openpyxl
#endif

ruta=ruta()

ruta.set(sys.argv[2])
         
Junta=impresor() 
Junta._init_()

encab=[]
hacer=[]
quien=[]


###HACE CALCULO PARA MODIFICAR HORAS UTC
dif=datetime.utcnow().hour-datetime.now().hour
if dif <0:   ##diferencia horas respecto de greenw para cambiar horas en calend
     dif=24+dif
#endif

#####DEFINICION DE FUNCIONES     
def forma():
    if elijo.get()==1:
       Junta._init_()
    #endif   
    if elijo.get()==2: 
       Junta.falsif()
    #endif   
    print(Junta.valor())
#fin forma    


def genera():
 if len(fec.get())!=8 or not fec.get().isdigit:
     messagebox.showinfo(message="Formato AAAAMMDD", title="CORRIJA")
     return
 #endif
   
 fecha=fec.get()
 encab,hacer,quien=limpia()
 calen=open(Archi_calen,encoding='UTF-8')
 encontro=False
 for x in calen:
    
    if   x.startswith("DTSTART:"+fecha): 
       x=x[17:21]
       encab.append(x)
       encontro=True  
       #print(x)
    #endif
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
                  #endif   
                  primera=False
             #endif     
       #endfor
       if len(historia)==0:
           historia=0
       #endif    
       hacer.append(int(historia))            
       quien.append(tipo[0:10])
       encontro=False
    #endif
 #endfor      
 consulta=pd.DataFrame({"quien": quien,"HC1":hacer,"hora": encab})
 
 cons=pd.merge(how="right",left=pacientes,right=consulta,left_on="HC",right_on="HC1")


 cons['DNI']=cons["DNI"].fillna(0)
 cons['HC']=cons["HC"].fillna(0)
 cons['Paciente']=cons["Paciente"].fillna("")
 cons['Psicólogo']=cons["Psicólogo"].fillna("")
 cons['Plan farmacológico']=cons["Plan farmacológico"].fillna("")
 cons['Frecuencia']=cons["Frecuencia"].fillna("")
 cons['Columna4']=cons["Columna4"].fillna("")
 cons['Teléfono']=cons["Teléfono"].fillna("")
 cons['Columna2']=cons["Columna2"].fillna(0)
 cons['Columna2']=cons["Columna2"].apply(lambda a: int(a))
 
 #Cabecera e Impresion
 print(Junta.valor())
 
 if Junta.valor():  #hoja Unica GENERA HOJA
  c = canvas.Canvas(ruta.valor()+"Consulta"+fecha+".pdf", pagesize=A4)
  #endif
 j=0 
 for j in range(len(cons)):
    if not Junta.valor():  #Hoja para cada consulta
      c = canvas.Canvas(ruta.valor()+"Consulta"+fecha+'-'+str(j)+".pdf", pagesize=A4)
    #endif
    impre(j,cons,fecha,c)
 #endfor
    
 messagebox.showinfo(message="El Dia tiene "+str(j+1)+" consultas", title="Fin del dia "+fecha[6:]+"/"+fecha[4:6]+"/"+fecha[0:4])
 encab,hacer,quien=limpia()
 ##FIN Cabecera e Impresion     
##Fin Genera
 
def salir():
    vent.destroy()
#Fin Salir

def limpia():
     return [],[],[]
#Fin limpia
              
def abre(titulo):
 ventana= Tk()
 ventana.withdraw()
 archi=filedialog.askopenfilename(title=titulo)
 ventana.destroy()
 #archi=askopenfilename()
 return archi
#Fin abre

def impre(cual,cons,fecha,c): 
  Horas=str(cons.iloc[cual]['hora'])[0:2]
  minu=str(cons.iloc[cual]['hora'])
  Horas=int(Horas)-int(dif)
  Horas=str(Horas)
  minu=str(cons.iloc[cual]['hora'])[2:4]
  w, h = A4
  x = 120
  y = h - 45
  paci=str(pacientes.iloc[cual]["Paciente"])
  text = c.beginText(50, h - 50)    
  text.setFont("Times-Roman", 12)
  text.textLine("FECHA..: "+ fecha[6:]+"/"+fecha[4:6]+"/"+fecha[0:4]+"              "+"    Historia Clínica: "+str(cons.iloc[cual]['HC1']))
  text.textLine("HORA....: "+ Horas+":"+minu)
  text.textLine("NOMBRE........: "+str(cons.iloc[cual]["Paciente"]))
  text.textLine("DNI..................: "+str(int(cons.iloc[cual]['DNI'])))
  text.textLine("TELÉFONO.....: "+str(cons.iloc[cual]['Teléfono']))
  text.textLine("EDAD..........:")
  text.textLine("SEXO...........:")
  text.textLine("CABA")
  text.textLine("PROVINCIA")
  text.textLine("")
  text.textLine("")
  text.textLine("                                                                     NP")
  text.textLine("ORIENTACIÓN")
  text.textLine("ADMISIÓN")
  text.textLine("PSICO INDIVIDUAL")
  text.textLine("PSIQ.CTL")
  text.textLine("INTERCONSULTA")
  text.textLine("CERTIFICADO S.M.")
  text.textLine("PSIQ FAMILIAR")
  text.textLine("PSICO FAMILIAR")
  text.textLine("TERAPIA OCUP.")
  text.textLine("TERAPIA OCUP.FLIAR")
  text.textLine("NUTRICIÓN IND")
  text.textLine("SUPERVISIÓN")
  text.textLine("TRABAJO SOCIAL")
  text.textLine("")
  text.textLine("")
  text.textLine("EGRESO")
  text.textLine("AUSENTE")
  text.textLine("")
  text.textLine("")
  text.textLine("FECHA PRÓX TURNO")

  text.setFont("Times-Roman", 5)
  text.textLine("Frecuencia....: "+str(cons.iloc[cual]['Frecuencia']))
  text.textLine("Psiquiatra....: "+str(cons.iloc[cual]['Columna4']))
  text.textLine("Psicólogo.....: "+str(cons.iloc[cual]['Psicólogo']))
  text.textLine("Plan Farma....: "+str(cons.iloc[cual]['Plan farmacológico'])[0:80])
  text.textLine("                 "+str(cons.iloc[cual]['Plan farmacológico'])[80:])
  
  text.textLine("Diagnóstico...: "+str(cons.iloc[cual]['Diagnostico presuntivo']))
  

  c.rect(10,h-550,360,700)
  c.drawText(text)
  xlist = [200, 250]
  ylist = [h-140,h-160,h-180]
  c.grid(xlist, ylist)

  xlist = [200,250, 300]
  ylist = [ h-198,h-212,h-226,h-240,h-254,h-268,h-282,h-296,h-310,h-324,h-338,h-352,h-366,h-380,h-394]
  c.grid(xlist, ylist)

  xlist = [200, 300]
  ylist = [h-427,h-441,h-455]
  c.grid(xlist, ylist)

  xlist = [200, 300]
  ylist = [h-489,h-503]
  c.grid(xlist, ylist)

  c.showPage()
  
  if Junta.valor():
      if cual==len(cons)-1:  #.si es ultimo del dia graba pdf
         c.save()
      #endif  
  else:
      c.save()               #Graba para cada paciente dentro del dia hoja diferente
  #endif   
   
##Fin Impre


 
###trae pacientes
archi=abre("Seleccione Archivo Excel Pacientes") #abre excel pacientes

pacientes=pd.read_excel(archi,engine="openpyxl")
       
### trajo pacientes

#abre zip
archi=abre("Seleccione Zip de Calendario")
####FIN SELECCION ARCHIVO

#####DEZIPEADO
Archi_calen=""

with zipfile.ZipFile(archi,'r')  as Todozip:
  listOfFileNames = Todozip.namelist()
   # Iterate over the file names
  for fileName in listOfFileNames:
       # Check filename endswith csv
       if fileName.endswith('.ics') and fileName.startswith('Hospi'):
           # Extract a single file from zip
           Archi_calen=fileName
           Todozip.extract(fileName)
       #endif
  #endfor

#print("el archi es "+Archi_calen)
#if not Archi_calen.startswith("Hospi"):
#    messagebox.showinfo(message="No hubo Calendario Hospital", title="CORRIJA")
#    break
    

    
encab,hacer,quien=limpia()



##genera ventana Principal

fecha=""
vent=tk.Tk()
vent.geometry("400x150+500+300")
vent.title("Consultas x Dia")

etiq1=tk.Label(vent,text="Fecha (AAAAMMDD): ")
etiq1.pack(padx=20,pady=20,side=tk.LEFT)

fec=tk.Entry(vent,width=8)
fec.pack(side=tk.LEFT)
fecha=fec.get()

boton1=tk.Button(vent,text="aceptar", command=genera) ##desdencadena verdadero proceso
                  
boton1.pack(padx=20,pady=40,side=tk.LEFT)

boton2=tk.Button(vent,text="Salir", command=salir)
boton2.pack(side=tk.LEFT)

elijo=IntVar()
elijo.set(value=1)
opc1=tk.Radiobutton(vent,text="Juntos",variable=elijo,value=1,command=forma).pack(side=tk.BOTTOM)
opc2=tk.Radiobutton(vent,text="Separado",variable=elijo,value=2,command=forma).pack(side=tk.BOTTOM)

vent.mainloop()
##Fin genera ventana Principal

