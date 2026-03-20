#%% 1.Imports
import os
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from datetime import date
import shutil
import xlwings as xw
import calendar
import locale
#%% 2. Directorios
carpeta_ing_path = r'C:\Users\fbonvecchiato\OneDrive - LOGINTER SA\General - Ingeniería\Carga de Horas\Carga de Horas\Integrantes ING'
integrantes_info_path = r'C:\Users\fbonvecchiato\OneDrive - LOGINTER SA\General - Ingeniería\Carga de Horas\Carga de Horas\Horas ING.xlsm'


#Consolidado
dir_consolidado = r'C:\Users\fbonvecchiato\OneDrive - LOGINTER SA\General - Ingeniería\Carga de Horas\Carga de Horas\Horas Ingeniería (Histórico)\Consolidado Horas'


#Plantilla
folder_plantilla = r'C:\Users\fbonvecchiato\OneDrive - LOGINTER SA\General - Ingeniería\Carga de Horas\Carga de Horas\Complementos'
filename_plantilla = 'plantilla_usuario.xlsm'
path_p = os.path.join(folder_plantilla, filename_plantilla)

#Status Integrantes
status_int_path = r'C:\Users\fbonvecchiato\OneDrive - LOGINTER SA\General - Ingeniería\Carga de Horas\Carga de Horas\Complementos\Status Integrantes\Int - NC'

#Plantilla status
folder_plantilla_st = r'C:\Users\fbonvecchiato\OneDrive - LOGINTER SA\General - Ingeniería\Carga de Horas\Carga de Horas\Complementos\Status Integrantes'
filename_plantilla_st = 'plantilla_status.xlsx'
path_p_stat = os.path.join(folder_plantilla_st, filename_plantilla_st)




#%% 2.Funciones


def get_integrantes_dfs(carpeta_ing_path):
    
    
    integrantes_files = [file for file in os.listdir(carpeta_ing_path) if 'xlsm' in file]
    
    
    dfs = []
    for integrante_f in integrantes_files:
        
        path = os.path.join(carpeta_ing_path, integrante_f)
        integrante = integrante_f[:-5]
        
        df = pd.read_excel(path, sheet_name = integrante, header = 6)
        
        df = (df, integrante)
        
        
        dfs.append(df)
        

    #Obtengo la información de los integrantes
    integrantes_info = pd.read_excel(integrantes_info_path, sheet_name = "Integrantes", index_col= 0 )
    
    
    #Me quedo solo con las columnas que tienen valores
    
    filt_dfs = []
    for df, integrante in dfs:
        
        
        #Ahora multipli el porcentaje por las horas correspondientes del integrante
        horas_integrante = integrantes_info.loc[integrante, 'Horas semanales [h]']
        
        
        #Multiplico los porcentajes por las horas correspondientes a cada usuario
        semanas = list(df.columns[8:])
        for semana in semanas:
            df[semana] = df[semana] * horas_integrante
            
        #df = df.dropna(axis = 1, how = 'all')
        
        
        
        filt_dfs.append((df, integrante))
        
    dfs  =[df for df in filt_dfs if len(df[0])>0]
    
    
    return dfs

def get_tareas_info(x):
    
    #Armo Dataframe
    df_tarea = pd.DataFrame(columns = ['ID', 'Proyecto', 'Usuario', 'Área', 'UUNN',
                                 'CeCo', 'Tipo','Facturación','Semana', 'Horas'])
    
    
    
    
    #Divido la parte de tarea y semanas
    #Tarea
    tarea = x[:8]
    #A tarea le agrego una columna que diga semana y otra horas:
    tarea["Semana"] = np.nan
    tarea["Horas"] = np.nan
    
    #Semanas
    data_semanas = x[8:]
    #De la parte de semanas tengo que eliminar los NaN values
    data_semanas.dropna(inplace = True)
    
    
    i = 0
    
    if len(data_semanas) != 0:
        
        
        
    
        for semana, horas in data_semanas.items():
            #Armo una serie:
            tarea["Semana"] = semana
            tarea["Horas"] = horas
            
            #Ahora agrego esa a un df 
            
            df_tarea.loc[i] = list(tarea)
            
            i = i + 1
    
    return df_tarea


def filter_integrante_df(df, integrante):
    
    lista_dfs = []
    for row in df.iterrows():
        x = row[1]
        
        #Si la data de semanas no tiene todos los valores vacios
        if not x[8:].isna().all():
            df_tarea = get_tareas_info(x)
            if len(df_tarea) > 0:
                lista_dfs.append(df_tarea)
    
    
    try:
        
        general_df = pd.concat(lista_dfs)
        general_df['Integrante'] = integrante 
    
    except:
        general_df = 0 #No hace falta la info de este integrante porque no tiene nada
    
    return general_df

def get_integrantes_dfs_filtered(carpeta_ing_path):
    
    #Obtengo información de los integrantes
    integrantes_info = pd.read_excel(integrantes_info_path, sheet_name = "Integrantes", index_col= 0 )
    
    #Obtengo los dfs de cada integrante
    dfs = get_integrantes_dfs(carpeta_ing_path)
    
    #Ahora los adapto a otro formato:
    dfs = [filter_integrante_df(df, integrante) for df, integrante in dfs]
    
    
    new_dfs = []
    for df in dfs:
        if not isinstance(df, int):
            new_dfs.append(df)
        
    #Ahora que tengo todos los dfs, les agrego los legajos
    
    dfs_finales = []
    for df in new_dfs:
        
        try:    
            integrante = df['Integrante'].unique()[0]
            legajo = integrantes_info.loc[integrante, 'Legajo']
            df["Legajo"] = legajo
            
        except:
            continue
        
        
        dfs_finales.append(df)
        
    
    return dfs_finales


def get_inicio_semana(semana):
    '''Esta función obtiene la fecha de inicio de la semana indicada.'''
    
    #Seteo para obtener formato en español
    locale.setlocale(locale.LC_TIME, 'es_ES.utf8')
    
    s = pd.read_excel(path_p, usecols = "I:BI").iloc[0]
    
    d = dict(zip(list(s.values), list(s.index)))
    
    
    return d[semana].date()
    
    
    
def general_df(dfs_filtrados):
    '''Esta funció agarra todos los dfs filtrados de los usuarios, los consolida en uno solo y despues se les hace una 
    modificación a ese archivo consolidado'''
    
    df_general = pd.concat(dfs_filtrados)
    
    #Le cambio el tipo a la columna 
    df_general["Semana"] = df_general["Semana"].astype(int) 
    
    
    #Agrego el mes a cada línea
    df_general["Mes"] = df_general["Semana"].apply(lambda x: get_month(x, path_p))
    
    
    #Ahora a cada mes le pongo su nombre 
    
    df_general["Mes nombre"] = df_general["Mes"].apply(lambda x: get_mes_name(x))
    
    
    #A cada semana, le pongo su inicio de semana:
        
    df_general["Inicio Semana"] = df_general["Semana"].apply(lambda x: get_inicio_semana(x))
    
    
    #A cada linea le actualizo sus datos
    
    #Abro base de datos de proyectos y tareas
    df_projects = pd.read_excel(integrantes_info_path, index_col= 0)
    
    
    #Actualizo todas las líneas:
    #df_general = df_general.apply(lambda x: update_task_info(x, df_projects), axis = 1) #ERROR ACÁ
        
    
    return df_general
    

def get_month(x, path_p):
    
    year_info = get_year_info(path_p) 
    
    mes = year_info[year_info["Semana"] == x]["Mes"].iloc[0]
    
    return mes


def save_df(df_general, dir_consolidado):
    
    fecha_hoy = datetime.today().strftime('%d-%m-%Y')
    
    path = os.path.join(dir_consolidado, f'Consolidado Horas ({fecha_hoy}).xlsx')
    
    df_general.to_excel(path, index = False)
    
    
    print("Se ha guardado el consolidado de horas")
    

def get_year_info(path_p):
    '''Esta función obtiene un dataframe que relaciona la semana con el mes que pertenece '''
    
    
    df = pd.read_excel(path_p)
    cols = [col for col in df.columns if not isinstance(col, str)]
    s = df[cols].iloc[0]

    fechas = [fecha.date() for fecha in s.index]
    semanas = s.values


    df = pd.DataFrame({'Fechas': fechas,
                       'Semana': semanas})


    df['Mes'] = df['Fechas'].apply(lambda x: x.month)
    
    
    
    return df

def get_semanas_info(path_p):
    '''Esta función obtiene la información de las semanas y sus feriados del año y la devuelve en un dataframe'''
    
    
    def get_sem_per(x):
        
        if isinstance(x, str):
            
            x  = int(x[2:-1])
            
        return x
    
    
    
    
    df = pd.read_excel(path_p).transpose()
    
    df = df.iloc[2:,:3]
    
    df.reset_index(inplace = True)
    
    #Obtengo los nombres de las columnas
    cols = list(df.iloc[0])
    
    #Me quedo con la data que necesito
    
    df = df.iloc[6:]
    
    df.rename(columns = dict(zip(df.columns, cols)), inplace = True)
    
    
    #Obtengo el numero del porcentaje de la semana. Los nan significan semanas al 100%
    df["% Semana Correcto"] =  df["% Semana"].apply(lambda x: get_sem_per(x))
    df.drop(columns = "% Semana", inplace = True)
    
    
    
    #A los Nan value sen % Semana Correcto le pongo el 100 de semana completa
    df["% Semana Correcto"].fillna(100, inplace = True)
    
    
    df["Mes"] = df["Fecha Semana"].apply(lambda x: x.month)
    
    #Cambio de float a int y lo pongo como index
    df["Número Semana"] = df["Número Semana"].astype(int)
    
    
    df.set_index("Número Semana", inplace = True)
    
    

    
    return df


def get_integrantes_horas(integrantes_info_path):  
    
    #Obtengo información de los integrantes
    integrantes_info = pd.read_excel(integrantes_info_path, sheet_name = "Integrantes", index_col= 0 )
 
    
    d = dict(zip(list(integrantes_info.index), integrantes_info["Horas semanales [h]"]))
    
    
    return d


def get_month_results(general_df, month, path_p, integrantes_info_path):
    
    
    #Obtengo la info de las semanas
    semanas_info = get_semanas_info(path_p)
    
    
    d_horas = get_integrantes_horas(integrantes_info_path)
    
    
    #Me quedo solo con la info respecto al mes indicado
    df_mes = df_general[df_general["Mes"] == month]
    
    

    #Ahora quiero chequear cuanto cargo cada integrante de los que cargaron
    result = df_mes.groupby(["Integrante", "Semana"])["Horas"].sum().reset_index()
    
    
    #A cada semana le asigno su porcentaje si tienen feriados o no
    result["% Semana Correcto"] = result["Semana"].apply(lambda x: semanas_info.loc[x]["% Semana Correcto"])
        
    
    
    #Asigno las horas correspondientes a cada semana en condiciones normales
    result["Horas correspondientes normales"] = result["Integrante"].apply(lambda x: d_horas[x])
    
    #Estas horas dependen si hay feriado o no
    result["Horas correspondientes reales"] = result["Horas correspondientes normales"] * (result ["% Semana Correcto"]/100)
     
    
    #Calculo el porcentaje realizado por semana
    
    result["Estado"] = result["Horas"]/result["Horas correspondientes reales"] * 100
    
    result["EstadoB"] = result["Horas"]/result["Horas correspondientes normales"] * 100

    
    return result


def get_ultimatum_date(month, year):
    
    '''Esta función obtiene la fecha de ultimatum que tienen que estar todas
    las horas cargadas.
    
    Pre: month, year: type int. Tienen que ser los datos del mes al que le quiero saber 
    el ultimatum'''
    
    
    day_last = calendar.monthrange(year, month)[1] #Ultimo día del mes
    
    
    #Transformo a date time object el último día del mes
    fecha_last = datetime(year, month, day_last)
    
    #A la última fecha del mes le sumo una semana
    delta_t = timedelta(days = 8)
    
    
    #Armo la fecha ultimatum
    ultimatum = fecha_last + delta_t
    
    return ultimatum

def define_stat(percentage, fecha_hoy, ultimatum):
    
    '''Esta función obtiene el porcnetaje de completaod del mes, la fecha actual
    y la fecha de ultimatum'''
    
    
    
    days_left = (ultimatum - fecha_hoy).days
    
    
    if (days_left <= 2) and (percentage < 1):
        
        stat = "AVISAR MUY FUERTE"
        
    else:
        
         if (fecha_hoy.day >= 25) or (fecha_hoy.month == ultimatum.month):
            #Si estamos en el 25 tengo que hacer el aviso general a todos
            
            if (percentage) > 0 and (percentage < 0.6):
                
                #Acá tengo que hacer un aviso.
                #Al que le falta solo 1 semana anterior, le aviso leve. Esta carga
                #está entre un 0.5 y 0.6
                
                if (percentage >= 0.5):
                    
                    stat = "AVISAR LEVE"
                    
                else:
                    
                    stat = "AVISAR"
                
                
            else:
                
                if percentage == 0:
                    
                    stat = "AVISAR FUERTE" 
                    
                else:
                    
                    stat = "NO AVISAR"
         else: 
            stat = "NO AVISAR"
                    
    return stat


def get_integrantes_month_status(df_general, path_p, integrantes_info_path, fecha = False):
    
    #Obtengo la fecha de hoy y el ultimatum
    fecha_hoy = datetime.today()
    
    
    if fecha == False:
    
        
        ultimatum = get_ultimatum_date(fecha_hoy.month, fecha_hoy.year)
        
        month = fecha_hoy.month
        
    else:
        month = fecha[0]
        
        ultimatum = get_ultimatum_date(month, fecha[1])
    
    
    
    
    result = get_month_results(df_general, month, path_p, integrantes_info_path)
    

    #Obtengo la información de las semanas
    semanas_info = get_semanas_info(path_p)
    semanas_info = semanas_info[semanas_info["Mes"] == month]
    
    
    grouped = result.groupby(["Integrante"])
    d = {}
    
    
    
    for integrante, group_df in grouped:
        
        df = group_df.copy()[["Semana", "Estado"]]
        
        df.set_index("Semana", inplace = True)
        
        #Conversión a porcentaje
        df["Estado"] = df["Estado"] / 100
        
        for semana in semanas_info.index:
            
            #Si no se encuentra el numero de semana en el df, agregarlo
            if semana not in df.index:
                
                df.loc[semana] = [0]
        #Para cada semana agrego una columna fecha
        df.reset_index(inplace = True)
        df["Fecha Semana"] = df["Semana"].apply(lambda x: semanas_info.loc[x, "Fecha Semana"].date())
        
        #Ahora necesito hacer una columna que diga si la semana falta o no. Esto es
        #para hacer un warning el 25 de cada mes
        total = len(df["Semana"]) * 100
        estado_mensual = df["Estado"].sum() * 100
        percentage = estado_mensual / total
        
        
        df["Aviso"] = define_stat(percentage, fecha_hoy, ultimatum)
        
            
        
        df.set_index("Fecha Semana", inplace = True)
        
        
        d[integrante] = df.transpose().sort_index(axis=1)
        
        
    #Ahora tengo que agregar al diccionario los integrantes que no cargaron nada en todo el mes
    integrantes_info = pd.read_excel(integrantes_info_path, sheet_name = "Integrantes", index_col= 0 )
    
    peligrosos = [integrante for integrante in integrantes_info.index if integrante not in d.keys()]
    
    
    if len(peligrosos) > 0:
        
        #Los agrego, pero con vacías
        #Hago una copia de la tabla pero vacio algunos datos
        #y se las agrego a todos los preligrosos
        dfp = d[list(d.keys())[0]].copy()
        dfp.loc["Estado"] = 0
        
        
        dfp.loc["Aviso"] = define_stat(0, fecha_hoy, ultimatum)
        for peligroso in peligrosos:
            d[peligroso] = dfp.copy()
    
        
    return d

def delete_files_in_folder(folder_path):
    try:
        # Get the list of files in the folder
        files = os.listdir(folder_path)

        # Iterate through each file and delete it
        for file in files:
            file_path = os.path.join(folder_path, file)
            if os.path.isfile(file_path):
                os.remove(file_path)
                print(f"Deleted: {file_path}")
            else:
                print(f"Skipped: {file_path} (not a file)")

        print("Deletion complete.")
    except Exception as e:
        print(f"An error occurred: {e}")

def save_status(d):
    
    
    #De algun integrante saco el mes
    month = d[list(d.keys())[0]].columns[0].month
    
    
    #Convierto para status
    month = get_mes_name(month)
    
    
    #Primero borro los status anteriores
    delete_files_in_folder(status_int_path)
    
    app = xw.App()
    for integrante, df in d.items():
        
        dest_path = os.path.join(status_int_path, integrante + ".xlsx")
        
        shutil.copy(path_p_stat, dest_path)
        
        
        #Ahora abro archivo luego de hacer la copia
        wb = app.books.open(dest_path)
        ws = wb.sheets["plantilla_status"]
        
        #Cambio valores de la tabla y el mes
        ws.range("B2").options(index = False, header = True).value = df
        
        
        ws.range("A1").value = month
        
        
        #Guardo cambios
        wb.save(dest_path)
        
        wb.close()
    app.quit()


def get_mes_name(num):
    '''Devuelve el nombre del mes'''
    

    # List of month names in Spanish
    month_names_spanish = [
        "Enero", "Febrero", "Marzo", "Abril",
        "Mayo", "Junio", "Julio", "Agosto",
        "Septiembre", "Octubre", "Noviembre", "Diciembre"
    ]
    
    
    # Get the corresponding month name in Spanish
    month_name_spanish = month_names_spanish[num - 1]
    
    
    return month_name_spanish


def update_task_info(x, df_projects):
    
    '''Esta función actualiza los valores de filas de las tareas cargadas por los 
    usuarios. Muchas veces se hacen cambios en la hoja madre "HORAS_ING", pero los 
    usuarios no actualizan sus hojas. El consolidado se tiene que hacer con estos cambios.
    Lo único que permance igual son los ID de los proyectos.
    
    '''
    
    #Tomo el id del proyecto
    p_id = x["ID"]
    #Proyecto data
    s = df_projects.loc[p_id]
    
    
    #Los elementos en la lista columns, son los elementos que debo cambiar:
    #Defino elementos:
    columns = ['Proyecto', 'Usuario', 'Área', 'UUNN', 'CeCo', 'Tipo', 'Facturación']
    
    
    for col in columns:
        
        x[col] = s[col]
    
    
    return x



#%%Programa

dfs_filtrados = get_integrantes_dfs_filtered(carpeta_ing_path)


#%%
df_general = general_df(dfs_filtrados)


#%%
save_df(df_general, dir_consolidado)

#%%
status = get_integrantes_month_status(df_general, path_p, integrantes_info_path, (6, 2024))
save_status(status)







    
