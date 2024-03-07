#' @export
# Cafe_verde
# Cargar la biblioteca readxl

f_Cafe_verde_pergamino<-function(directorio,mes,anio){

  #Cargar librerias
  library(readxl)
  library(dplyr)
  library(openxlsx)
  library(zoo)

  #identificar la carpeta
  carpeta=nombre_carpeta(mes,anio)



# STOCKS cafe verde ------------------------------------------------------------------

  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Cafe","NOMBRE"]

  # Especifica la ruta del archivo de Excel
  Cafe_verde <- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/CAFÉ/",archivo),
                              sheet = "Inventarios")


  #En esa fila, reemplazar NA por el valor de la columna anterior
  #Identificar la fila donde esta la palabra totales Colombia
  n_fila=which(grepl("Totales Colombia",as.data.frame(t(Cafe_verde))), arr.ind=TRUE)
  Cafe_verde[n_fila, ] <- na.locf0(Cafe_verde[n_fila,])
  ###crear alerta de que cambia formato

  #si which es 0 entonces generar error o algo

  #identificar las columna donde dice total general y peso en pie
  columna1=which(grepl("Totales Colombia",Cafe_verde),arr.ind = TRUE)
  columna2=which(grepl("Verde",Cafe_verde),arr.ind = TRUE)
  columnaf <- intersect(columna1, columna2)


  tamaño=25+mes
  #Tomar el valor que nos interesa
  vector_existencias=as.data.frame(tail(Cafe_verde[,columnaf],tamaño))

  Valor_existencias=as.numeric(vector_existencias$...8)-as.numeric(lag(vector_existencias$...8))
  Valor_existencias=Valor_existencias[-1]


# Exportaciones cafe verde -----------------------------------------------------------

  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Cafe","NOMBRE"]

  # Especifica la ruta del archivo de Excel
  Cafe_verde <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/CAFÉ/",archivo),
                           sheet = "Info exportaciones",startRow = 6)
  Cafe_verde[,1]= as.Date(Cafe_verde[,1], origin = "1899-12-30")


  exportaciones <- Cafe_verde %>%
    filter(Cafe_verde[,1]>=paste0((anio-2),"-","01-01"))%>%
    as.data.frame()

  exportaciones=exportaciones %>%
       filter(grepl("verde" ,tolower(exportaciones[,2]))) %>%
       group_by(Mes.de.Embarque) %>%
       summarise(total_exportaciones=sum(as.numeric(`Sacos.de.60.kilos.-.café.verde.equivalente`))/1000)
  total_exportaciones=as.numeric(exportaciones$total_exportaciones)

# IMPORTACIONES cafe verde -----------------------------------------------------------
  archivos=list.files(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE"))
  elementos_seleccionados <- archivos[grepl("Expos e", archivos) ]

  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Importaciones","NOMBRE"]
  # Especifica la ruta del archivo de Excel
  Cafe_verde <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/",elementos_seleccionados,"/",archivo),
                           sheet = "PNK")
  n_fila=which(Cafe_verde== "230801",arr.ind = TRUE)[,"row"]
  n_col1=which(Cafe_verde== paste0((anio-2)," 01"),arr.ind = TRUE)[,"col"]
  n_col2=which(Cafe_verde== paste0(anio," ",mes_0[mes]),arr.ind = TRUE)[,"col"]


importaciones=(as.numeric(Cafe_verde[n_fila,n_col1:n_col2])/60)/1000




# Consumo interno cafe verde ---------------------------------------------------------

archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Cafe trillado","NOMBRE"]

# Especifica la ruta del archivo de Excel
Cafe_verde <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/CAFÉ/",archivo),
                        sheet = "Consumo Total ")
tabla=Cafe_verde %>%
      filter(CAFÉ.TRILLADO.CONSUMIDO.POR.LA.INDUSTRIA.TORREFACTORA.NACIONAL>=(anio-2))
tamaño=24+mes
fila1=which(tabla==(anio-2) ,arr.ind = TRUE)[, "row"]
n_columna=which(grepl("Cantidad Consumida",Cafe_verde),arr.ind = TRUE)

consumo_interno <- (as.numeric(tabla[fila1[1]:(fila1[1]+tamaño-1),n_columna])/60)/1000




# Produccion cafe verde --------------------------------------------------------------

valor_produccion=Valor_existencias+(total_exportaciones-importaciones)+consumo_interno

vector_valores_cafe_verde=cbind(Valor_existencias,total_exportaciones,importaciones,consumo_interno,valor_produccion)



# Produccion cafe pergamino -----------------------------------------------

valor_denominador=0.754

produccion_equivalente_pergamino=valor_produccion/valor_denominador


# Inventarios pergamino ---------------------------------------------------
archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Cafe","NOMBRE"]
# Especifica la ruta del archivo de Excel
Cafe_pergamino <- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/CAFÉ/",archivo),
                         sheet = "Inventarios")


#En esa fila, reemplazar NA por el valor de la columna anterior
#Identificar la fila donde esta la palabra totales Colombia
n_fila=which(grepl("Totales Colombia",as.data.frame(t(Cafe_pergamino))), arr.ind=TRUE)
Cafe_pergamino[n_fila, ] <- na.locf0(Cafe_pergamino[n_fila,])
###crear alerta de que cambia formato

#si which es 0 entonces generar error o algo

#identificar las columna donde dice total general y peso en pie
columna1=which(grepl("Totales Colombia",Cafe_pergamino),arr.ind = TRUE)
columna2=which(grepl("Pergamino",Cafe_pergamino),arr.ind = TRUE)
columnaf <- intersect(columna1, columna2)


tamaño=25+mes
#Tomar el valor que nos interesa
vector_existencias=as.data.frame(tail(Cafe_pergamino[,columnaf],tamaño))

Valor_existencias_pergamino=as.numeric(vector_existencias$...7)/valor_denominador-as.numeric(lag(vector_existencias$...7))/valor_denominador

Valor_existencias_pergamino=Valor_existencias_pergamino[-1]

# Produccion total pergamino ----------------------------------------------

produccion_total_pergamino=(produccion_equivalente_pergamino+Valor_existencias_pergamino)*1000



# datos finales -----------------------------------------------------------

valores_cafe_verde_pergamino=cbind(vector_valores_cafe_verde,produccion_equivalente_pergamino,Valor_existencias_pergamino,produccion_total_pergamino)

return(valores_cafe_verde_pergamino)
}
