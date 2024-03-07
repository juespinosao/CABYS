#' @export
# Cafetos
# Cargar la biblioteca readxl

f_Cafetos<-function(directorio,mes,anio){

  #Cargar librerias
  library(readxl)
  library(dplyr)
  library(openxlsx)
  library(zoo)

  #identificar la carpeta
  carpeta=nombre_carpeta(mes,anio)



  # STOCKS cafe verde ------------------------------------------------------------------
  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Cafetos","NOMBRE"]


  # Especifica la ruta del archivo de Excel
  Cafetos <- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/CAFÉ/",archivo),
                           sheet = "Renovaciones")


  #En esa fila, reemplazar NA por el valor de la columna anterior
  #Identificar la fila donde esta la palabra totales Colombia
  n_columna=which(grepl("Total de",Cafetos), arr.ind=TRUE)
  ###crear alerta de que cambia formato

  #si which es 0 entonces generar error o algo

  #identificar las columna donde dice total general y peso en pie
  fila1=which(Cafetos=="Enero",arr.ind = TRUE)[,"row"]




  #Tomar el valor que nos interesa
  vector_area=as.data.frame(Cafetos[fila1:(fila1+mes-1),n_columna])

  Valor_area=as.numeric(vector_area$...3)-as.numeric(lag(vector_area$...3))
  Valor_area[1]=as.numeric(vector_area[1,])

  return(Valor_area)
}
