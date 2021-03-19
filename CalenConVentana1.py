###Importa modulos
from tkinter import messagebox
import tkinter as tk
from tkinter import *

###Define e incializa variables 
global fecha
global Junta
global donde


donde=r"C:\Users\claudio\Desktop"
Junta=True 
fecha=""



#####DEFINICION DE FUNCIONES     
def ComoImprimo():
    if elijo.get()==1:
       Junta=True
    #endif   
    if elijo.get()==2: 
       Junta=False
    #endif   
    print(Junta)
#fin ComoImprimo    


def genera():
 if len(fec.get())!=8 or not fec.get().isdigit:
     messagebox.showinfo(message="Formato AAAAMMDD", title="CORRIJA")
     return
 #endif
 
 cons=[1,2,3,4,5]  
 ##Cabecera e Impresion
 if Junta:  #hoja Unica GENERA HOJA
      c="creo canvas unico"
      print(c)
 #endif
  
 for j in range(len(cons)):
    print("Juntas ?"+str(Junta))
    if not Junta:  #Hoja para cada consulta
      print("si ..SEPARA") 
      c="creo canvas para cada uno "+str(j)
      print(c)
    #endif
    impre(j,cons,fecha,c)
 #endfor
    
 messagebox.showinfo(message="El Dia tiene "+str(len(cons))+" consultas", title="Fin del dia "+fecha[6:]+"/"+fecha[4:6]+"/"+fecha[0:4])
 ##FIN Cabecera e Impresion     
##Fin Genera
 
def salir():
    vent.destroy()
#Fin Salir

              

def impre(cual,cons,fecha,c):

  
  if Junta:
      if cual==len(cons)-1:  #.si es ultimo del dia graba pdf
         print("Grabo una para todos")
      #endif  
  else:
      print("Grabo cada una "+str(cual))               #Graba para cada paciente dentro del dia hoja diferente
  #endif   
   
##Fin Impre


 
#           

    



##genera ventana Principal

vent=tk.Tk()
vent.geometry("400x150+500+300")
vent.title("Consultas x Dia")

etiq1=tk.Label(vent,text="Fecha (AAAAMMDD): ")
etiq1.pack(padx=20,pady=20,side=tk.LEFT)

fec=tk.Entry(vent,width=8)
fec.pack(side=tk.LEFT)
fecha=fec.get()

boton1=tk.Button(vent,text="aceptar", command=genera)
boton1.pack(padx=20,pady=40,side=tk.LEFT)
boton2=tk.Button(vent,text="Salir", command=salir)
boton2.pack(side=tk.LEFT)

elijo=IntVar()
elijo.set(1)
opc1=tk.Radiobutton(vent,text="Juntos",variable=elijo,value=1,command=ComoImprimo).pack()
opc2=tk.Radiobutton(vent,text="Separado",variable=elijo,value=2,command=ComoImprimo).pack()

vent.mainloop()
##Fin genera ventana Principal

