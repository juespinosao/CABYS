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
  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
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
  n_col_2=which(Frutas== paste0((anio-1)," ",1),arr.ind = TRUE)[,"col"]



  Frutas2 <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/",elementos_seleccionados,"/",archivo),
                      sheet = "CTES FOBPES")
  n_fila_2=which(Frutas2 == "010499" |Frutas2 == "010403",arr.ind = TRUE)[,"row"]
  n_fila_2=c(n_fila[[1]],n_fila[[2]])
  n_col_1_2=which(Frutas2== paste0(anio," ",mes),arr.ind = TRUE)[,"col"]
  n_col_2_2=which(Frutas2== paste0((anio-1)," ",1),arr.ind = TRUE)[,"col"]

  Frutas3 <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/",elementos_seleccionados,"/",archivo),
                       sheet = "IP_EXPO")


   valor_exportaciones=as.data.frame(cbind(t(Frutas[n_fila,n_col_2[1]:n_col_1[1]]),t(Frutas2[n_fila_2,n_col_2_2[1]:n_col_1_2[1]]),t(Frutas3[n_fila_2,n_col_2_2[1]:n_col_1_2[1]])))
   for (i in 1:6) {
     valor_exportaciones[,i]=as.numeric(valor_exportaciones[,i])
   }






  # Consumo interno ---------------------------------------------------------
archivo=nombre_archivos[nombre_archivos$PRODUCTO=="SIPSA","NOMBRE"]
  Frutas <- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/Datos_SIPSA/",archivo))

  fila1=min(which(Frutas==2013,arr.ind = TRUE)[,"row"])
  fila2=min(which(Frutas==anio,arr.ind = TRUE)[,"row"])
  columna1=which(Frutas=="Frutas citricas retropolado",arr.ind = TRUE)[,"col"]
  columna2=which(Frutas=="Otras frutas retropolado",arr.ind = TRUE)[,"col"]

  Valor_Frutas=as.data.frame(na.omit(Frutas[fila1:fila2,c(columna1[1],columna2[1])]))




  # Agrupar datos -----------------------------------------------------------


  return(list(variacion = valor_exportaciones, vector = Valor_Frutas))
}
