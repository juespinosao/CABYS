#' @export
# Banano
# Cargar la biblioteca readxl

f_Banano<-function(directorio,mes,anio){

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
  Banano <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/",elementos_seleccionados,"/",archivo),
                          sheet = "TOTAL EXPO_KTES")


  n_fila=which(Banano == "010401",arr.ind = TRUE)[,"row"]
  n_col1=which(Banano== paste0(anio-2," 1"),arr.ind = TRUE)[,"col"]
  n_col2=which(Banano== paste0(anio," ",mes),arr.ind = TRUE)[,"col"]



  #Tomar el valor que nos interesa
  Valor_exportaciones=as.numeric(Banano[n_fila[1],(n_col1[1]:n_col2[1])])






# Consumo interno ---------------------------------------------------------
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="SIPSA","NOMBRE"]

  Banano <- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/Datos_SIPSA/",archivo))
  fila_i=which(Banano$year==(anio-2) & Banano$month==1,arr.ind = TRUE)
  fila1=which(Banano==anio,arr.ind = TRUE)[,"row"]
  fila2=which(Banano==mes,arr.ind = TRUE)[,"row"]
  fila_f=intersect(fila1,fila2)
  Valor_interno=as.data.frame(na.omit(Banano[fila_i:fila_f,"Bananos"]))
# Agrupar datos -----------------------------------------------------------



return(list(exportaciones=Valor_exportaciones,consumo_interno=Valor_interno))
}
