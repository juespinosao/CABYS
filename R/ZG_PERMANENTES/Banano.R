# Banano
# Cargar la biblioteca readxl
directorio="C:/Users/Asus/OneDrive - dane.gov.co/proyecto2/Automatizacion CABYS/Automatizacion/Formato_carpetas"
mes=7
anio=2023
f_Banano<-function(directorio,mes,anio){
  
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
  Banano <- read.xlsx(paste0(directorio,"/",anio,"/",carpeta,"/consolidado_ISE/",elementos_seleccionados,"/Resumen Exportaciones ",mes_0[mes],"-",anio," - copia.xlsx"), 
                          sheet = "TOTAL EXPO_KTES")
  
  
  n_fila=which(Banano == "010401",arr.ind = TRUE)[,"row"]
  n_col1=which(Banano== paste0(anio-2," 1"),arr.ind = TRUE)[,"col"]
  n_col2=which(Banano== paste0(anio," ",mes),arr.ind = TRUE)[,"col"]
  
  
  
  #Tomar el valor que nos interesa
  Valor_exportaciones=as.numeric(Banano[n_fila[1],(n_col1[1]:n_col2[1])])
  

  
  
  

# Consumo interno ---------------------------------------------------------

  Banano <- read_excel(paste0(directorio,"/",anio,"/",carpeta,"/Datos_SIPSA/Base_EB_SIPSA.xlsx"))
  fila_i=which(Banano$year==(anio-2) & Banano$month==1,arr.ind = TRUE)
  fila1=which(Banano==anio,arr.ind = TRUE)[,"row"]
  fila2=which(Banano==mes,arr.ind = TRUE)[,"row"]
  fila_f=intersect(fila1,fila2)
  Valor_interno=as.data.frame(na.omit(Banano[fila_i:fila_f,"Bananos"]))
# Agrupar datos -----------------------------------------------------------


  
return(list(exportaciones=Valor_exportaciones,consumo_interno=Valor_interno))
}
