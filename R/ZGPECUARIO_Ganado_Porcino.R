#' @export
# Ganado_porcino
# Cargar la biblioteca readxl

f_Porcino<-function(directorio,mes,anio){

#Cargar librerias
  library(readxl)
  library(dplyr)
  library(zoo)
#identificar la carpeta
  carpeta=nombre_carpeta(mes,anio)

  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="ESAG1","NOMBRE"]

  # Especifica la ruta del archivo de Excel
  Ganado_Porcino <- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/ESAG/",archivo),
                              sheet = "Cuadro_3")

#Identificar la fila donde esta la palabra Periodo
  n_fila=which(Ganado_Porcino == "Periodo",arr.ind = TRUE)[, "row"]

#En esa fila, reemplazar NA por el valor de la columna anterior
  Ganado_Porcino[n_fila, ] <- na.locf0(Ganado_Porcino[n_fila,])
  ###crear alerta de que cambia formato

#identificar las columna donde dice total general y peso en pie
  #identificar las columna donde dice total general y peso en pie
  columna1=which(grepl("Total general",Ganado_Porcino),arr.ind = TRUE)
  columna2=which(grepl("Peso en pie",Ganado_Porcino),arr.ind = TRUE)
  columnaf1 <- intersect(columna1, columna2)
  columna1=which(grepl("Machos",Ganado_Porcino),arr.ind = TRUE)
  columnaf2 <- intersect(columna1, columna2)
  columna1=which(grepl("Hembras",Ganado_Porcino),arr.ind = TRUE)
  columnaf3 <- intersect(columna1, columna2)
  fila=which(Ganado_Porcino=="Enero",arr.ind = TRUE)[,"row"]
  #Filtrar la fila del mes de interes



  #Tomar el valor que nos interesa
  Valor_Porcino=as.data.frame(Ganado_Porcino[fila:(fila+mes-1),c(columnaf1,columnaf2,columnaf3)])




 ###### modificar expo e impo




  # Exportaciones -----------------------------------------------------------

  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Exportaciones","NOMBRE"]

  # Especifica la ruta del archivo de Excel
  archivos=list.files(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE"))
  elementos_seleccionados <- archivos[grepl("Expos e", archivos) ]
  # Especifica la ruta del archivo de Excel
  Porcino <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/",elementos_seleccionados,"/",archivo),
                      sheet = "PNK")


  n_fila=which(Porcino == "020300",arr.ind = TRUE)[,"row"]
  n_col1=which(Porcino== paste0(anio," 1"),arr.ind = TRUE)[,"col"]
  n_col2=which(Porcino== paste0(anio," ",mes),arr.ind = TRUE)[,"col"]

  n_fila2=which(Porcino == "210102",arr.ind = TRUE)[,"row"]


  #Tomar el valor que nos interesa
  Valor_exportaciones=as.data.frame(t(Porcino[c(n_fila,n_fila2[[1]]),(n_col1[1]:n_col2[1])]))
  Valor_exportaciones=as.data.frame(lapply(Valor_exportaciones, as.numeric))




  # Importaciones -----------------------------------------------------------

  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Importaciones","NOMBRE"]

  # Especifica la ruta del archivo de Excel
  Impor <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/",elementos_seleccionados,"/",archivo),
                     sheet = "PNK")
  n_fila=which(Impor== "020300",arr.ind = TRUE)[,"row"]
  n_col1=which(Impor== paste0((anio)," 01"),arr.ind = TRUE)[,"col"]
  n_col2=which(Impor== paste0(anio," ",mes_0[mes]),arr.ind = TRUE)[,"col"]

  n_fila2=which(Impor == "210102"|Impor == "210104",arr.ind = TRUE)[,"row"]


  #Tomar el valor que nos interesa
  Valor_importaciones=as.data.frame(t(Impor[c(n_fila,n_fila2[[1]],n_fila2[[2]]),(n_col1[1]:n_col2[1])]))
  Valor_importaciones=as.data.frame(lapply(Valor_importaciones, as.numeric))
  Valor_importaciones=Valor_importaciones %>%
    mutate(suma=Valor_importaciones[,2]+Valor_importaciones[,3])




  Valor_Porcino=cbind(Valor_Porcino,Valor_importaciones[1],Valor_importaciones[,4],Valor_exportaciones[1],Valor_exportaciones[,2])
  colnames(Valor_Porcino)=c("TOTAL","MACHOS","HEMBRAS","IMPO_VIVO","IMPO_CARNE","EXPO_VIVO","EXPO_CARNE")
  Valor_Porcino=as.data.frame(lapply(Valor_Porcino, as.numeric))


  return(Valor_Porcino)

}
