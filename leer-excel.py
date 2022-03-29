from re import S
from statistics import mode
from numpy import empty
import pandas as pd
from pandas import ExcelWriter
import openpyxl 
import xlsxwriter
from seaborn import load_dataset
#leo el archivo con todas las hojas (se puede cambiar el archivo de origen)
archivo = r"C:\Users\trpblg\Downloads\inventario\slees.xlsx"

#leo el archivo donde voy a escribir y a la vez lo preparo para la escritura
archivo2= r"C:\Users\trpblg\Downloads\inventario\destino.xlsx"

#declaro arrays para poder usarlos mas adelante para buscar en las paginas y construir el dataframe que luego se insertara en el excel
#Modificar el nombre de columnas dependiendo del excel de entrada 
#la variable "paginas" contiene las paginas que va a leer del excel, deben ser las mismas que hay en el excel
#la variable "columna" contiene las columnas que se escribiran en el excel
#paginas=["kapacitor-1.5.9", "elasticsearch-7.10.2", "alerta.io-latest", "influxdb-1.8.6", "nginx-1.21.0", "postgresql-13.3", "grafana-latest", "logstash-7.10.2", "node.js-14.15", "debian-10.slim", "debian-11.slim"]
#columnas=["Aplicacion","kapacitor-1.5.9", "elasticsearch-7.10.2", "alerta.io-latest", "influxdb-1.8.6", "nginx-1.21.0", "postgresql-13.3", "grafana-latest", "logstash-7.10.2", "node.js-14.15", "debian-10.slim", "debian-11.slim" ]    
paginas=["bci-base", "os-rpm"]
columnas=["Aplicacion", "bci-base", "os-rpm"]

#concateno todas las paginas par que se quede en un solo dataframe y sea mas facil manejarlo
df= pd.concat(pd.read_excel(archivo, sheet_name=None, header=None), ignore_index=True)
#elimino las entradas vacias
df = df.dropna()

df2=pd.DataFrame( columns=list("AB"))

#itero sobre cada linea
for index, row in df.iterrows():
    flag = False
    #print(row[1])
    #Elimino del row el :amd64 y el .x86_64
    row[0]= row[0].replace(":amd64", "")
    row[0]=row[0].replace(".x86_64", "")
    #leo el dataframe que inicia como vacio, al no encontrar nada comienza a escribir en df2, posteriormente al ir añadiendose rows si que comienza a comprobar duplicados
    
    for index2, row2 in df2.iterrows():
        
        if row[0] == row2[0]:
            flag = True
            break
    
    if not flag :
        #creo un nuevo dataframe para insertar los datos de las columnas en cuestion, 0 y 1 de los datos no repetidos
        insertar= pd.DataFrame([[row[0], str(row[1])]], columns=list("AB"))
        df2 = pd.concat([df2, insertar], ignore_index=True)
        
#Ordeno el dataframe usando la columna A como criterio, se ordena alfabeticamente
df2 = df2.sort_values("A")
   

#vuelvo a leer el excel de origen con los datos de aplicaciones y versiones en bruto para poder realizar una busqueda        
dforigen= pd.read_excel(archivo, sheet_name=None, header= None)
dfvolcado= pd.DataFrame(columns=columnas)

#recorro cada linea del archivo modificado anteriormente, para compararla con el archivo de origen y buscar las coincidencias, asi podre identificar en que hoja coinciden y guardar la version para añadirla posteriormente a otro dataframe que volcare en un excel
for index2, row2 in df2.iterrows():
    #empiezo a construir el dataframe con cada row del  archivo previamente procesado
    datos={"Aplicacion":row2[0]}
    insertar2=pd.DataFrame([datos],columns=["Aplicacion"])

    for sheet in paginas:
        #limpio las lineas vacias del archivo de origen para poder comparar mejor
        dforigen[sheet] = dforigen[sheet].dropna()
        
        for index, linea in dforigen[sheet].iterrows():
            #elimino parte de los nombres que no me interesa para compararlos
            linea[0]= linea[0].replace(":amd64", "")
            linea[0]= linea[0].replace(".x86_64", "")
            #elinmino posibles .86_64 de las versiones
            linea[1]= linea[1].replace(".x86_64", "")
            linea[1]= linea[1].replace(".noarch", "")
            
            if row2[0] == linea[0]:
                #cuando encuentra coincidencia recogo el valor de linea[1] que es la segunda columna del archivo de origen e inserto una nueva columna en el dataframe que estoy construyendo junto con el dato que corresponde con la version
                insertar2[sheet] = str(linea[1])
                break
    #una vez construido el dataframe con todas las hojas lo concateno con el dataframe de volcado que luego insertare en el excel
    dfvolcado = pd.concat([dfvolcado, insertar2], ignore_index=True)            

print(dfvolcado)
#Escribo el dataframe procesado en el excel
with pd.ExcelWriter(archivo2, engine= "openpyxl", mode='a', if_sheet_exists='replace')as writer:
    dfvolcado.to_excel(writer,sheet_name="Sheet1", index=None )
    writer.save()     