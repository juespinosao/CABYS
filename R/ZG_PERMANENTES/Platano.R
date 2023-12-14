# Platano
# Cargar la biblioteca readxl

f_Platano<-function(directorio,mes,anio){

  #Cargar librerias
  library(readxl)
  library(dplyr)
  library(openxlsx)
  library(zoo)

  #identificar la carpeta
  carpeta=nombre_carpeta(mes,anio)



  # Exportaciones ------------------------------------------------------------------



  # Especifica la ruta del archivo de Excel
  archivos=list.files(paste0(directorio,"/",anio,"/",carpeta,"/consolidado_ISE"))
  elementos_seleccionados <- archivos[grepl("Expos e", archivos) ]
  # Especifica la ruta del archivo de Excel
  Platano <- read.xlsx(paste0(directorio,"/",anio,"/",carpeta,"/consolidado_ISE/",elementos_seleccionados,"/Resumen Exportaciones ",mes_0[mes],"-",anio," - copia.xlsx"),
                      sheet = "PNK")


  n_fila=which(Platano == "010402",arr.ind = TRUE)[,"row"]
  n_col1=which(Platano== paste0((anio-2)," 1"),arr.ind = TRUE)[,"col"]
  n_col2=which(Platano== paste0(anio," ",mes),arr.ind = TRUE)[,"col"]



  #Tomar el valor que nos interesa
  Valor_exportaciones=as.numeric(Platano[n_fila[1],n_col1[1]:n_col2[1]])/1000






  # Consumo interno ---------------------------------------------------------

  Platano <- read_excel(paste0(directorio,"/",anio,"/",carpeta,"/Datos_SIPSA/Base_EB_SIPSA.xlsx"))
  fila_i=which(Platano$year==(anio-2) & Platano$month==1,arr.ind = TRUE)
  fila1=which(Platano==anio,arr.ind = TRUE)[,"row"]
  fila2=which(Platano==mes,arr.ind = TRUE)[,"row"]
  fila_f=intersect(fila1,fila2)
  Valor_Platano=as.data.frame(na.omit(Platano[fila_i:fila_f,"Platanos"]))



  # Agrupar datos -----------------------------------------------------------


  return(list(exportaciones=Valor_exportaciones,consumo_interno=Valor_Platano))
}
