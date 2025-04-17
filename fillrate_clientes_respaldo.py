# %%
#IMPORTACION DE LIBRERIAS
import pandas as pd
import datetime
import os
import numpy as np
import getpass
# hoy = datetime.datetime.today() dejar esta linea cuadno se haga el calculo real
hoy = datetime.datetime.today()
#LECTURA DE DFS
from pathlib import Path
usuario = getpass.getuser()
import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from datetime import datetime, timedelta

# Variables globales para almacenar las fechas formateadas
fecha_1 = None
fecha_2 = None

# Función que calcula el rango de fechas y cierra la ventana
def calcular_y_continuar():
    global fecha_1, fecha_2
    
    fecha_input = calendario.get_date()

    # Calcular el lunes de la semana anterior
    fecha_lunes_anterior = fecha_input
    
    # Calcular el domingo de la semana anterior
    fecha_domingo_anterior = fecha_lunes_anterior + timedelta(days=6)

    # Formatear las fechas como dd.mm.yyyy
    fecha_1 = fecha_lunes_anterior.strftime("%d.%m.%Y")
    fecha_2 = fecha_domingo_anterior.strftime("%d.%m.%Y")
    
    # Cerrar la ventana
    ventana.destroy()

# Crear la ventana principal
ventana = tk.Tk()
ventana.title("Selección de Fechas")
ventana.geometry("300x250")

# Etiqueta de instrucción
label_instruccion = tk.Label(ventana, text="Selecciona una fecha:")
label_instruccion.pack(pady=10)

# Calendario de selección de fecha
calendario = DateEntry(ventana, date_pattern='dd.mm.yyyy', background='darkblue', foreground='white', borderwidth=2)
calendario.pack(pady=10)

# Botón para calcular el rango y continuar
boton_ok = ttk.Button(ventana, text="OK", command=calcular_y_continuar)
boton_ok.pack(pady=10)

# Iniciar la aplicación
ventana.mainloop()

# Una vez que la ventana se cierra, las fechas ya están disponibles
print(f"Fechas disponibles para usar:\nLunes: {fecha_1}\nDomingo: {fecha_2}")

# Aquí puedes continuar con el resto del código
# Ejemplo: 
# print("Continuando con el proceso usando fecha_1 y fecha_2")


import getpass

usuario = getpass.getuser()
import win32com.client

# Initialize the SAP GUI scripting
SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

# Maximize the SAP window
session.findById("wnd[0]").maximize()

# Enter transaction code "va05" and execute
session.findById("wnd[0]/tbar[0]/okcd").text = "va05"
session.findById("wnd[0]").sendVKey(0)

session.findById("wnd[0]/tbar[1]/btn[33]").press()
session.findById("wnd[1]/usr/ctxtVBCOM-VKORG").text = "vi02"
session.findById("wnd[1]/usr/ctxtVBCOM-VKORG").caretPosition = 4
session.findById("wnd[1]").sendVKey(0)

# Set date range
session.findById("wnd[0]/usr/ctxtVBCOM-AUDAT").text = fecha_1
session.findById("wnd[0]/usr/ctxtVBCOM-AUDAT_BIS").text = fecha_2
session.findById("wnd[0]/usr/ctxtVBCOM-AUDAT_BIS").setFocus()
session.findById("wnd[0]/usr/ctxtVBCOM-AUDAT_BIS").caretPosition = 10

# Press button to continue
session.findById("wnd[0]/tbar[1]/btn[33]").press()
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[0]/tbar[1]/btn[20]").press()

# Select checkbox in the popup
session.findById("wnd[1]/usr/sub:SAPLKAB1:0400/chkRKAB1-XSUCH[4,0]").selected = True
session.findById("wnd[1]/usr/sub:SAPLKAB1:0400/chkRKAB1-XSUCH[4,0]").setFocus()
session.findById("wnd[1]/tbar[0]/btn[0]").press()

# Set some text in the next screen
session.findById("wnd[2]/usr/sub:SAPLKAB1:0410/ctxtRKAB1-VONSL[0,23]").text = "zrdp"
session.findById("wnd[2]/usr/sub:SAPLKAB1:0410/ctxtRKAB1-BISSL[0,44]").text = "zvdp"
session.findById("wnd[2]/usr/sub:SAPLKAB1:0410/ctxtRKAB1-BISSL[0,44]").setFocus()
session.findById("wnd[2]/usr/sub:SAPLKAB1:0410/ctxtRKAB1-BISSL[0,44]").caretPosition = 4
session.findById("wnd[2]/tbar[0]/btn[0]").press()

# Execute and download Excel file
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell(6, "ARKTX")
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu()
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem("&XXL")

# Define file path and save the file
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/usr/ctxtDY_PATH").text = f'C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Nivel de Servicio OEM/automatizacion'
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "VA05.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[11]").press()

# Close the windows
session.findById("wnd[0]").sendVKey(3)
session.findById("wnd[0]").sendVKey(3)

import pandas as pd
dtype = {'Documento comercial':'str',
'Posición (SD)':'str','Solicitante':'str',
'Sector':'str' }

import time

time.sleep(30)




va05 = pd.read_excel(f'C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Nivel de Servicio OEM/automatizacion/VA05.XLSX', dtype=dtype)

# %%
va05['Material'] = va05['Material'].astype('str')
va05_reducida = va05[['Documento comercial','Posición (SD)','Clase doc.ventas','Fecha documento','Nº pedido cliente','Solicitante','Cantidad de pedido','Material','Nombre 1','Unidad de medida', 'Valor neto','Sector']]

va05_reducida['AUX'] = va05_reducida['Documento comercial'] + va05_reducida['Posición (SD)']
va05_reducida = va05_reducida.sort_values(by='AUX', ascending=True)
va05_reducida['repetido'] = (va05_reducida['AUX'] == va05_reducida['AUX'].shift(1)).astype(int)
sectores_permitidos = ['01','02','03','04','05','06','07']

# Aplicar el filtro correcto
va05_base = va05_reducida[(va05_reducida['Sector'].isin(sectores_permitidos)) & (va05_reducida['repetido'] == 0)]



ruta = f'C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Bases Indicadores en CSV {hoy.year}-{hoy.month:02d}'
ruta_repo = Path(ruta)

columnas= ['Nro_pieza_fabricante_1',	'Cod_Actual_1']
ruta_cod = ruta_repo.joinpath('COD_ACTUAL.csv')

# Leer el archivo CSV en un DataFrame
cadena_de_remplazo = pd.read_csv(ruta_cod)
cadena_de_remplazo = cadena_de_remplazo[columnas]

va05_base = va05_base.merge(cadena_de_remplazo, left_on='Material', right_on='Nro_pieza_fabricante_1', how='left')
va05_base['Cod_Actual_1'] = va05_base['Cod_Actual_1'].fillna(va05_base['Material'])
va05_base = va05_base.drop('Nro_pieza_fabricante_1', axis=1)
ruta_mara = ruta_repo.joinpath('MARA_R3.csv')

# Leer el archivo CSV en un DataFrame
df_mara = pd.read_csv(ruta_mara, dtype={'Grupo_articulo':'str'})
sectores = pd.read_excel(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Gerenciamiento MOS/Panel PBI/Info Sectores.xlsx",dtype = {'Sector - material':'str'}, usecols=['Sector - material','Sector'])
sectores['Sector - material'] = sectores['Sector - material'].str.zfill(2)
va05_base = va05_base.merge(sectores, left_on = 'Sector',right_on='Sector - material', how='left')
va05_base.rename(columns={'Sector_x':'Sector','Sector_y':'Nombre Sector'}, inplace=True)
va05_base.drop('Sector - material', axis=1, inplace=True)
va05_base = va05_base.merge(df_mara[['Material_R3','Material_dsc']], left_on='Cod_Actual_1',right_on='Material_R3', how='left')

columnas = ['Material','País de origen (Material)']

# %%
#Cambiar y dejar 1 lectura de archivo
# ###################################################
# ruta_carpeta_base = "C:/Users/lravlic/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras Maestros/Base Planificable"
# carpeta_base = os.listdir(ruta_carpeta_base)
# for archivo in carpeta_base:
#     if  str(hoy.year) in archivo and f'{hoy.month:02d}' in archivo:
#         ruta = ruta_carpeta_base + '/' + archivo
#         ruta_2 = os.listdir(ruta)
#         for archivo in ruta_2:
#             if 'Base' in archivo:
#                 print(ruta + '/' + archivo )
#                 base = pd.read_excel(ruta + '/' + archivo, header=1, engine='openpyxl', sheet_name='BP Febrero')
base = pd.read_excel("C:/Users/lravlic/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras Maestros/Base Planificable/2025-02 Base Planificable/Base Febrero OEM-AXS.xlsx", header=1, engine='openpyxl', sheet_name='BP Febrero')

base_df =base[columnas]

base_df.drop_duplicates(subset='Material', inplace=True)

base_df_ue = pd.merge(base_df, cadena_de_remplazo, left_on="Material", right_on="Nro_pieza_fabricante_1", how="left")
base_df_ue['Cod_Actual_1'] = base_df_ue['Cod_Actual_1'].fillna(base_df_ue['Material'])

base_df_ue = base_df_ue[['Material','Cod_Actual_1','País de origen (Material)']]
df_nnss = va05_base.merge(base_df[['Material','País de origen (Material)']], left_on='Material', right_on='Material', how= 'left')
base_df_ue.drop_duplicates(subset=['Cod_Actual_1'], keep='first', inplace=True)
df_nnss = df_nnss.merge(base_df_ue[['Cod_Actual_1','País de origen (Material)']], left_on='Cod_Actual_1', right_on='Cod_Actual_1', how='left')
df_nnss['País de origen (Material)_x'] = df_nnss['País de origen (Material)_x'].fillna(df_nnss['País de origen (Material)_y'])
df_nnss.drop(['País de origen (Material)_y'], inplace = True, axis=1)
df_nnss = df_nnss.rename(columns = {'País de origen (Material)_x':'Origen'})
df_mara.drop_duplicates(subset='Material_R3', inplace=True)

df_mara_ue = pd.merge(df_mara, cadena_de_remplazo, left_on="Material_R3", right_on="Nro_pieza_fabricante_1", how="left")
df_mara_ue['Cod_Actual_1'] = df_mara_ue['Cod_Actual_1'].fillna(df_mara_ue['Material_R3'])

df_mara_ue = df_mara_ue[['Cod_Actual_1','Familia','Modelo']]
df_nnss = df_nnss.merge(df_mara[['Material_R3','Familia','Modelo']], left_on='Material', right_on='Material_R3', how= 'left')


df_mara_ue.drop_duplicates(subset=['Cod_Actual_1'], keep='first', inplace=True)
df_nnss = df_nnss.merge(df_mara_ue[['Cod_Actual_1','Familia','Modelo']], left_on='Cod_Actual_1', right_on='Cod_Actual_1', how='left')

df_nnss['Familia_x'] = df_nnss['Familia_x'].fillna(df_nnss['Familia_y'])
df_nnss['Modelo_x'] = df_nnss['Modelo_x'].fillna(df_nnss['Modelo_y'])
df_nnss.drop(columns = ['Familia_y','Modelo_y'], inplace = True, axis=1)
df_nnss = df_nnss.rename(columns = {'Familia_x':'Familia', 'Modelo_x':'Modelo'})
df_nnss.drop(columns={'Material_R3_x', 'Material_R3_y'}, inplace=True)

p_mantencion = pd.read_excel(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras Maestros/Mantención/SKU Mantenciones OEM 1.xlsx", sheet_name = 'Plan de mantencion Antiguo', dtype={'UE Junio':'str','Plan de mantención':'str'})

p_mantencion.rename(columns={'UE Junio':'UE','Plan de mantención':'Plan_mantencion'}, inplace=True)
p_mantencion_ue = pd.merge(p_mantencion, cadena_de_remplazo, left_on="UE", right_on="Nro_pieza_fabricante_1", how="left")
p_mantencion_ue['Cod_Actual_1'] = p_mantencion_ue['Cod_Actual_1'].fillna(p_mantencion_ue['UE'])
p_mantencion_ue['Plan_mantencion'].fillna(0, inplace=True)


p_mantencion_ue.drop_duplicates(subset=['Cod_Actual_1'], keep='first', inplace=True)
df_nnss = df_nnss.merge(p_mantencion_ue[['Cod_Actual_1','Plan_mantencion']], left_on='Cod_Actual_1', right_on='Cod_Actual_1', how='left')


df_nnss['Plan_mantencion'].fillna('0',inplace=True)
path = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Bases Indicadores en CSV 2025-02/Segmentacion WebApp"
dfs = []
carpeta = os.listdir(path)
for i in carpeta:
    print(path + '/'+ i)
    df = pd.read_csv(path + '/'+ i)
    dfs.append(df)

for df in dfs:
    df = df[['SKU ERP','Total Segmentation']]
df_consolidado= pd.concat(dfs)

# %%
df_consolidado = df_consolidado[['SKU ERP','Total Segmentation']]
df_consolidado['Total Segmentation'].value_counts()
segmentacion = df_consolidado
segmentacion.to_csv(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Bases Indicadores en CSV 2025-02/segmentacion_consolidado.csv")

segmentacion_ue = pd.merge(segmentacion, cadena_de_remplazo, left_on="SKU ERP", right_on="Nro_pieza_fabricante_1", how="left")
segmentacion_ue['Cod_Actual_1'] = segmentacion_ue['Cod_Actual_1'].fillna(segmentacion_ue['SKU ERP'])
segmentacion_ue = segmentacion_ue[['Cod_Actual_1','Total Segmentation']]
segmentacion_ue =segmentacion_ue.sort_values(by='Total Segmentation')
segmentacion_ue.drop_duplicates(subset='Cod_Actual_1',keep='first', inplace=True)

df_nnss = df_nnss.merge(segmentacion_ue, left_on='Cod_Actual_1', right_on='Cod_Actual_1', how='left')
df_nnss['Total Segmentation'] = df_nnss['Total Segmentation'].fillna('OO')


flag_ipl = pd.read_excel(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras Maestros/2025/2025-02/2025-02 Base Flag IPL OEM.xlsx", header=1, dtype={'Flag IPL': 'str'})
estrategico = pd.read_excel(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras Maestros/2025/2025-01/Estratégicos Chile - Cierre Diciembre 2024.xlsx", sheet_name='Estratégico')
flag_ipl = flag_ipl[['Último Eslabón R3','Flag IPL']]
estrategico = estrategico[['Último Eslabón R3', 'Tipo']]

flag_ipl_ue = pd.merge(flag_ipl, cadena_de_remplazo, left_on="Último Eslabón R3", right_on="Nro_pieza_fabricante_1", how="left")
flag_ipl_ue['Cod_Actual_1'] = flag_ipl_ue['Cod_Actual_1'].fillna(flag_ipl_ue['Último Eslabón R3'])
flag_ipl_ue.drop(columns = {'Nro_pieza_fabricante_1','Último Eslabón R3'}, inplace=True)
flag_ipl_ue.sort_values(by='Flag IPL', ascending=False, inplace=True)
flag_ipl_ue.drop_duplicates(subset='Cod_Actual_1', keep='first', inplace=True)
df_nnss = df_nnss.merge(flag_ipl_ue, left_on='Cod_Actual_1', right_on='Cod_Actual_1', how='left')
df_nnss['Flag IPL'] = df_nnss['Flag IPL'].fillna(0)

estrategico_ue = pd.merge(estrategico, cadena_de_remplazo, left_on="Último Eslabón R3", right_on="Nro_pieza_fabricante_1", how="left")
estrategico_ue['Cod_Actual_1'] = estrategico_ue['Cod_Actual_1'].fillna(estrategico_ue['Último Eslabón R3'])
estrategico_ue.drop(columns = {'Nro_pieza_fabricante_1','Último Eslabón R3'}, inplace=True)
estrategico_ue.drop_duplicates(subset='Cod_Actual_1', keep='first', inplace=True)
df_nnss = df_nnss.merge(estrategico_ue, left_on='Cod_Actual_1', right_on='Cod_Actual_1', how='left')
df_nnss['Tipo'] = df_nnss['Tipo'].fillna(0)
df_nnss['Tipo'] = df_nnss['Tipo'].replace('Estrategico', 1)

# Convert the 'Tipo' column to string type
df_nnss['Tipo'] = df_nnss['Tipo'].astype(str)
df_mara = df_mara[['Material_R3','Grupo_articulo']]
df_mara_ue_2 = df_mara.merge(cadena_de_remplazo, left_on='Material_R3', right_on='Nro_pieza_fabricante_1', how='left')
df_mara_ue_2 = df_mara_ue_2[['Cod_Actual_1','Grupo_articulo']]
cabecera = 3
hoja = 'Vigencias'
gr_art= pd.read_excel(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras Maestros/2024/2024-11/Vigencia Grupo Articulo MU Q4_2023.xlsx", header=cabecera,sheet_name = hoja)
gr_art = gr_art[['Grupo articulo','Chile']]
gr_art.rename(columns = {'Chile':'Apertura Parque'} , inplace=True)
gr_art.drop_duplicates(subset='Grupo articulo', keep='first',inplace=True)
df_mara_ue_vigencia = df_mara_ue_2.merge(gr_art, left_on='Grupo_articulo', right_on='Grupo articulo', how='left')
df_mara_ue_vigencia['Apertura Parque'].fillna('0', inplace=True)
df_mara_ue_vigencia = df_mara_ue_vigencia[df_mara_ue_vigencia['Apertura Parque'].isin(['No Vigente','Vigente','Nuevo'])]
df_mara_ue_vigencia.drop(columns={'Grupo_articulo','Grupo articulo'}, inplace=True)
df_mara_ue_vigencia['Parque'] = df_mara_ue_vigencia['Apertura Parque'].apply(lambda x: 0 if x == 'No Vigente' else 1)
df_mara_ue_vigencia.dropna(subset='Cod_Actual_1', inplace=True)
df_mara_ue_vigencia.sort_values(by='Parque', ascending=False, inplace=True)
df_mara_ue_vigencia.drop_duplicates(subset='Cod_Actual_1', keep='first', inplace=True)
df_nnss = df_nnss.merge(df_mara_ue_vigencia, left_on='Cod_Actual_1', right_on='Cod_Actual_1', how='left')
df_nnss['Parque'] = df_nnss['Parque'].fillna(0).astype('int')
df_nnss['Linea OK'] = df_nnss['Clase doc.ventas'].apply(lambda x: 1 if x=='ZRDP' else 0 )
df_nnss['Linea VOR'] = 1 - df_nnss['Linea OK']
df_nnss['Fecha documento'] = pd.to_datetime(df_nnss['Fecha documento'])
df_nnss['Semana'] = df_nnss['Fecha documento'].dt.isocalendar().week
df_nnss['Month'] = df_nnss['Fecha documento'].dt.month
df_nnss['Equipo'] = df_nnss['Nombre Sector'].apply(
    lambda x: 'Mainstream 1' if x in ["JAC Cars", "Changan", "Great Wall"] else 'Mainstream 2'
)


# %%
df_nnss['Moneda'] = 'CLP'
df_nnss['Material R3'] = df_nnss['Material']
df_nnss['Nombre Sector Anterior2'] = ""
df_nnss['Clase R (actual)'] = ""
df_nnss['Segmentación N-5'] =""
df_nnss['Planificable'] = df_nnss.apply(lambda row: 
    'Planificable' if (
        row['Total Segmentation'] in ["AA", "AB", "AC", "BA", "BB", "BC", "CA", "CB", "CC"] or 
        (row['Tipo'] == 1 and row['Apertura Parque'] in ["Vigente", "Nuevo"])
    ) else 'No Planificable', axis=1)
df_nnss['Planificable N-5']=""
df_nnss['Incidencia'] =""
df_nnss['Causa Raíz']=""
df_nnss['Línea Totales'] = 1
df_nnss['Qty Venta'] = df_nnss.apply(lambda row: row['Cantidad de pedido'] if row['Clase doc.ventas'] in ['ZVTA', 'ZRDP'] else 0, axis=1) 
df_nnss['Qty VOR'] = df_nnss.apply(lambda row: row['Cantidad de pedido'] if row['Clase doc.ventas'] in ['ZVOR', 'ZVDP'] else 0, axis=1) 
df_nnss['Qty Total'] = df_nnss['Qty Venta'] + df_nnss['Qty VOR']

df_nnss['Día'] = df_nnss['Fecha documento'].dt.day
df_nnss['Año'] = df_nnss['Fecha documento'].dt.year

df_nnss.rename(columns= {'Nº pedido cliente':'Referencia de cliente',
                        #fecha documento,
                        #clase documento venta,
                        'Documento comercial':'Documento de ventas',
                        'Posición (SD)':'Posición',
                        #aux,
                        #solicitante,
                        'Nombre 1':'Nombre Solicitante',
                        'Cantidad de pedido':'Cantidad de pedido (Posición)',
                        'Unidad de medida':'Un.medida venta',
                        'Valor neto':'Valor neto (posición)',
                        'Moneda':'Moneda del documento',
                        #Material,
                        'Material_dsc':'Descripción del material',
                        #MAterial_r3,
                        'Cod_Actual_1':'UE R3',
                        #nom_sec_anterior,
                        'Nombre Sector':'Nombre sector',
                        #origen,
                        #modelo,
                        #familia,
                        #parque
                        'Apertura Parque':'Apretura Parque',
                        'Plan_mantencion': 'Plan Mantención',
                        #clase r,
                        'Total Segmentation':'Segmentación',
                        'Tipo': 'Estratégicos',
                        'Planificable':'Planificables',
                        'Linea OK':'Línea OK',
                        'Linea VOR':'Línea VOR',
                        #Equipo,
                        #Flag IPL,
                        'Month':'Mes'}, inplace=True)
df_nnss = df_nnss[['Referencia de cliente', 'Fecha documento', 'Clase doc.ventas',
       'Documento de ventas', 'Posición', 'AUX', 'Solicitante',
       'Nombre Solicitante', 'Cantidad de pedido (Posición)',
       'Un.medida venta', 'Valor neto (posición)', 'Moneda del documento',
       'Material', 'Descripción del material', 'Material R3', 'UE R3',
       'Nombre Sector Anterior2', 'Nombre sector', 'Origen', 'Modelo',
       'Familia', 'Parque', 'Apretura Parque', 'Plan Mantención',
       'Clase R (actual)', 'Segmentación', 'Segmentación N-5', 'Estratégicos',
       'Planificables', 'Planificable N-5', 'Equipo', 'Flag IPL', 'Incidencia',
       'Causa Raíz', 'Línea OK', 'Línea VOR', 'Línea Totales', 'Qty Venta',
       'Qty VOR', 'Qty Total', 'Semana', 'Día', 'Mes', 'Año']]
df_nnss.drop(columns={'Moneda del documento'}, inplace=True)
df_nnss.to_csv(f"derco_sem6.csv")

# %%
import pandas as pd
from tkinter import Tk, filedialog

# Hide the main tkinter window
root = Tk()
root.withdraw()

# Open file dialog and get the file path
file_path = filedialog.askopenfilename(
    title="Selecciona FR Inchcape",
    filetypes=[("Excel files", "*.xls *.xlsx *.xlsm")]
)

# Print the selected file path
print(f"Selected file: {file_path}")

# Check if a file was selected
if file_path:
    try:
        # Read the Excel file into a pandas DataFrame
        inchcape = pd.read_excel(file_path, dtype={'Ce.': 'str', 'Material': 'str'}, sheet_name='Sheet1')

        # Print first few rows for verification
        print("Archivo cargado correctamente:")
        print(inchcape.head())

    except Exception as e:
        print(f"Error al leer el archivo: {e}")

# Destroy the Tk instance to free resources
root.destroy()

# %%
columns = [
    "Ce.", "ClVt", "Fecha doc-","Pedido","Material", "NPF", "Denominación", "Ctd.ped..1", 'Inventario'
]
inchcape = inchcape[columns]

# %%
maestro = pd.read_excel(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Nivel de Servicio OEM/automatizacion/Maestro.xlsx",sheet_name= 'Maestro 2.0', header=1, dtype={'Último Eslabón y Material SAP':'str'})

# %%
maestro = maestro[['Último Eslabón y Material SAP','Familia','Segmentación ','Modelo']]

# %%
maestro.rename(columns = {'Segmentación ': 'Segmentación'}, inplace=True)

# %%
maestro.drop_duplicates(subset='Último Eslabón y Material SAP', inplace=True)

# %%
inchcape = inchcape.merge(maestro, left_on='Material', right_on='Último Eslabón y Material SAP', how='left')

# %%
inchcape.drop(columns={'Último Eslabón y Material SAP'}, inplace=True)

# %%
#Marca
mappings = {
    '1335': "Subaru",
    '1305': "DFSK",
    '1344': "Geely"
}

# Apply the mapping to create the 'Marca' column
inchcape['Marca'] = inchcape['Ce.'].map(mappings).fillna(0)


# %%
inchcape['Líneas OK'] = inchcape['ClVt'].apply(lambda x: 1 if x == "Z300" else 0)
inchcape['Líneas VOR'] = inchcape['ClVt'].apply(lambda x: 1 if x == "ZVOR" else 0)


# %%
inchcape['Líneas Totales'] = inchcape['Líneas VOR']+inchcape['Líneas OK'] 

# %%
inchcape['Qty OK'] = inchcape['Líneas OK'] * inchcape['Ctd.ped..1']

# %%
inchcape['Qty VOR'] = inchcape['Líneas VOR'] * inchcape['Ctd.ped..1']

# %%
inchcape['Qty VOR'] = inchcape['Líneas VOR'] * inchcape['Ctd.ped..1']

# %%
inchcape.groupby('Marca').agg({'Líneas OK':'sum',
                               'Líneas VOR':'sum',
                               'Líneas Totales':'sum'}).reset_index()

# %%


# %%
meses = {
    "ene": 1,
    "feb": 2,
    "mar": 3,
    "abr": 4,
    "may": 5,
    "jun": 6,
    "jul": 7,
    "ago": 8,
    "sep": 9,
    "oct": 10,
    "nov": 11,
    "dic": 12
}

# Cambiar valores de la columna 'Mes' usando el diccionario
inchcape['Fecha doc-'] = pd.to_datetime(inchcape['Fecha doc-'] )
inchcape['Mes'] = inchcape['Fecha doc-'].dt.month
inchcape['Semana'] =   inchcape['Fecha doc-'].dt.isocalendar().week


# %%
# nnss['Centro'] = ''
# nnss['NPF'] = ""
# nnss['VOR'] = ""
# nnss['Considerar'] = ""

inchcape.rename(columns= {'Ce.':'Centro','ClVt':'Clase doc.ventas','Pedido':'Documento de ventas','Denominación':'Descripción del material','Ctd.ped..1':'Cantidad de pedido (Posición)','Marca':'Nombre sector','Líneas OK':'Línea OK', 'Líneas VOR':'Línea VOR', 'Líneas Totales':'Línea Totales','Qty OK':'Qty Venta' ,'Qty Totales':'Qty Total','Fecha doc-':'Fecha documento', 'Inventario':'Planificables'}, inplace=True)




# %%
inchcape.to_csv(f'C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Nivel de Servicio OEM/automatizacion/bases_semanales/inchcape.csv')

# %%
df_final = pd.concat([df_nnss, inchcape])

# %%
df_final.to_csv(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Nivel de Servicio OEM/automatizacion/consolidado_legacies/consolidado_sem_{datetime.today().isocalendar()[1]}.csv")

# %%

# %%



