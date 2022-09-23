## CARGA LIBRERIAS
import pandas as pd
import glob
from os import system
from tkinter import Tk
from tkinter.filedialog import askopenfilenames
import PySimpleGUI as sg


## FUNCIONES
def cargaArchivos(): #selecciono archivos dlt a convertir 
  
  Tk().withdraw() 
  filez = askopenfilenames ()

  return list(filez)



def obtenerPlanillas(): #convierto archivos dtl a csv y devuelvo la ruta donde los cree 
  
  easyConverter = r"C:\Users\tbermudez\Desktop\Efluentes\EasyConverter\Easyconverter" #ruta donde esta ubicada la aplicacion easyConverter
  listaArchivosDTL = cargaArchivos()
  ruta = ''

  print(f"Convirtiendo {len(listaArchivosDTL)} archivos .dtl a .cvs")

  for item in listaArchivosDTL:
    if item[-3:] == "dtl": #solo proceso archivos con extension DTL

      itemb = item.replace('/', '\\')
      origen = itemb
      destino = itemb[0:-3] + "csv"
      argumentos = f"\"{origen}\" \"{destino}\""
  
      system(easyConverter + " " + argumentos) #es recomendable usar la libreria SUBPROCCES pero no lo pude hacer funcionar

      ruta = item[0:-12]

  return ruta



def crearDF(ruta): #cargo un DataFrame de los csv en la ruta especificada
  listaCsv = glob.glob(ruta + "*.csv")
  df = pd.read_csv(listaCsv[0])

  for i in range(1, len(listaCsv)):
    df = pd.concat ([df, pd.read_csv(listaCsv[i])])

  return df


#columns=['Fecha', 'Hora', 'PH', 'Caudal', 'PH Max', 'PH Min', 'Caudal Vertedero', 'Caudal Totalizador', 'Caudal Totalizador Reseteable']


def resumenDF(df): #creo un DataFrame con promedios diarios

  dffinal = pd.DataFrame(columns=['Fecha', 'PH promedio', 'Caudal Promedio', 'Cantidad de valores']) #df vacio con las columnas de interes
  fechas = df['Fecha'].unique() #df con todas las fechas incluidas en el df original

  for item in fechas: #promedio de cada columna en cada fecha
    dffinal.loc[dffinal.shape[0]] = [item, round(df.loc[df['Fecha'] == item, 'PH'].mean(), 2),  round(df.loc[df['Fecha'] == item, 'Caudal'].mean(), 2), round(df.loc[df['Fecha'] == item, 'PH'].count(), 2)]

  return dffinal




#----------- SCRIPT --------------------------


window = sg.Window(title="Conversor DTL a XLSX", #Ventana TKInter
                   layout=[[sg.Text('Pulse para iniciar conversion')],
                           [sg.Button('Iniciar'), sg.Button('Cancelar')]],
                   margins=(200, 200))

while True:
    event, values = window.read()
    
    if event == "Iniciar":
      
      ruta = obtenerPlanillas()

      df = crearDF(ruta)
      with pd.ExcelWriter(ruta + 'final.xlsx') as writer: #creo el xlsx a partir de la funcion resumenDF()
        resumenDF(df).to_excel(writer, sheet_name="Hoja1")
      

      window = sg.Window(title="Conversor DTL a XLSX",
                   layout=[[sg.Text('Conversion finalizada')]],
                   margins=(100, 100))
      
    elif event == sg.WIN_CLOSED or event == 'Cancelar':
      break


window.close()



