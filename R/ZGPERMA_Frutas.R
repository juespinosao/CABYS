#' @export
# Frutas
# Cargar la biblioteca readxl

f_Frutas<-function(directorio,mes,anio){

  #Cargar librerias
  library(readxl)
  library(dplyr)
  library(openxlsx)
  library(zoo)

  #identificar la carpeta
  carpeta=nombre_carpeta(mes,anio)



  # Exportaciones ------------------------------------------------------------------
  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Exportaciones","NOMBRE"]


  # Especifica la ruta del archivo de Excel
  archivos=list.files(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE"))
  elementos_seleccionados <- archivos[grepl("Expos e", archivos) ]
  # Especifica la ruta del archivo de Excel
  Frutas <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/",elementos_seleccionados,"/",archivo),
                       sheet = "TOTAL EXPO_KTES")


  n_fila=which(Frutas == "010499" |Frutas == "010403",arr.ind = TRUE)[,"row"]
  n_fila=c(n_fila[[1]],n_fila[[2]])
  n_col_1=which(Frutas== paste0(anio," ",mes),arr.ind = TRUE)[,"col"]
  n_col_2=which(Frutas== paste0((anio-1)," ",mes),arr.ind = TRUE)[,"col"]


  #Tomar el valor que nos interesa
  Valor_exportaciones=sum(as.numeric(Frutas[n_fila,n_col_1[1]]))/sum(as.numeric(Frutas[n_fila,n_col_2[1]]))*100-100






  # Consumo interno ---------------------------------------------------------
archivo=nombre_archivos[nombre_archivos$PRODUCTO=="SIPSA","NOMBRE"]
  Frutas <- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/Datos_SIPSA/Base_EB_SIPSA.xlsx"))

  fila1=which(Frutas==anio,arr.ind = TRUE)[,"row"]
  fila2=which(Frutas==mes,arr.ind = TRUE)[,"row"]
  fila_f=intersect(fila1,fila2)
  Valor_Frutas=as.data.frame(na.omit(Frutas[1:fila_f,"Frutas"]))




  # Agrupar datos -----------------------------------------------------------


  return(list(variacion = Valor_exportaciones, vector = Valor_Frutas))
}
