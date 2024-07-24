#' @export
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

  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Exportaciones","NOMBRE"]


  # Especifica la ruta del archivo de Excel
  archivos=list.files(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE"))
  elementos_seleccionados <- archivos[grepl("Expos e", archivos) ]
  # Especifica la ruta del archivo de Excel
  Platano <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/",elementos_seleccionados,"/",archivo),
                      sheet = "PNK")


  n_fila=which(Platano == "010402",arr.ind = TRUE)[,"row"]
  n_col1=which(Platano== paste0((anio-2)," 1"),arr.ind = TRUE)[,"col"]
  n_col2=which(Platano== paste0(anio," ",mes),arr.ind = TRUE)[,"col"]



  #Tomar el valor que nos interesa
  Valor_exportaciones=as.numeric(Platano[n_fila[1],n_col1[1]:n_col2[1]])/1000






  # Consumo interno ---------------------------------------------------------
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="SIPSA","NOMBRE"]

  Platano <- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/Datos_SIPSA/Microdatos desde 2013/",archivo))
  fila1=min(which(Platano[,"year"]==(anio-2),arr.ind = TRUE)[,"row"])
  fila2=min(which(Platano[,"year"]==anio,arr.ind = TRUE)[,"row"])


  Valor_Platano=as.data.frame(na.omit(Platano[fila1:(fila2+mes-1),"Platanos"]))

  # Agrupar datos -----------------------------------------------------------


  return(list(exportaciones=Valor_exportaciones,consumo_interno=Valor_Platano))
}
