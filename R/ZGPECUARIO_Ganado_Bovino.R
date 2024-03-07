#' @export
#'
# Ganado_Bovino_kilo en pie
# Cargar la biblioteca readxl

f_Bovino<-function(directorio,mes,anio){

#Cargar librerias
library(readxl)
library(dplyr)
library(zoo)

#identificar la carpeta
carpeta=nombre_carpeta(mes,anio)

nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
archivo=nombre_archivos[nombre_archivos$PRODUCTO=="ESAG1","NOMBRE"]
# Especifica la ruta del archivo de Excel
Ganado_Bovino <- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/ESAG/",archivo),
                            sheet = "Cuadro_1")


#Identificar la fila donde esta la palabra Periodo
n_fila=which(Ganado_Bovino == "Periodo",arr.ind = TRUE)[, "row"]

#En esa fila, reemplazar NA por el valor de la columna anterior
Ganado_Bovino[n_fila, ] <- na.locf0(Ganado_Bovino[n_fila,])


###crear alerta de que cambia formato

#si which es 0 entonces generar error o algo

#identificar las columna donde dice total general y peso en pie
columna1=which(grepl("Total general",Ganado_Bovino),arr.ind = TRUE)
columna2=which(grepl("Peso en pie",Ganado_Bovino),arr.ind = TRUE)
columnaf1 <- intersect(columna1, columna2)
columna1=which(grepl("Machos",Ganado_Bovino),arr.ind = TRUE)
columnaf2 <- intersect(columna1, columna2)
columna1=which(grepl("Hembras",Ganado_Bovino),arr.ind = TRUE)
columnaf3 <- intersect(columna1, columna2)
columna1=which(grepl("Terneros",Ganado_Bovino),arr.ind = TRUE)
columnaf4 <- intersect(columna1, columna2)
columna1=which(grepl("Exportación",Ganado_Bovino),arr.ind = TRUE)
columnaf5 <- intersect(columna1, columna2)

fila=which(Ganado_Bovino=="Enero",arr.ind = TRUE)[,"row"]
#Filtrar la fila del mes de interes



#Tomar el valor que nos interesa
Valor_Bovino=as.data.frame(Ganado_Bovino[fila:(fila+mes-1),c(columnaf1,columnaf2,columnaf3,columnaf4,columnaf5)])


# Exportaciones -----------------------------------------------------------

archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Exportaciones","NOMBRE"]

# Especifica la ruta del archivo de Excel
archivos=list.files(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE"))
elementos_seleccionados <- archivos[grepl("Expos e", archivos) ]
# Especifica la ruta del archivo de Excel
Bovino <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/",elementos_seleccionados,"/",archivo),
                    sheet = "PNK")


n_fila=which(Bovino == "020100",arr.ind = TRUE)[,"row"]
n_col1=which(Bovino== paste0(anio," 1"),arr.ind = TRUE)[,"col"]
n_col2=which(Bovino== paste0(anio," ",mes),arr.ind = TRUE)[,"col"]

n_fila2=which(Bovino == "210101"|Bovino == "210104",arr.ind = TRUE)[,"row"]


#Tomar el valor que nos interesa
Valor_exportaciones=as.data.frame(t(Bovino[c(n_fila,n_fila2[[1]],n_fila2[[2]]),(n_col1[1]:n_col2[1])]))
Valor_exportaciones=as.data.frame(lapply(Valor_exportaciones, as.numeric))
Valor_exportaciones=Valor_exportaciones %>%
                    mutate(suma=Valor_exportaciones[,2]+Valor_exportaciones[,3])



# Importaciones -----------------------------------------------------------

archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Importaciones","NOMBRE"]

dos_digitos <- anio %% 100
# Especifica la ruta del archivo de Excel
Impor <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/",elementos_seleccionados,"/",archivo),
                        sheet = "PNK")
n_fila=which(Impor== "020100",arr.ind = TRUE)[,"row"]
n_col1=which(Impor== paste0((anio)," 01"),arr.ind = TRUE)[,"col"]
n_col2=which(Impor== paste0(anio," ",mes_0[mes]),arr.ind = TRUE)[,"col"]

n_fila2=which(Impor == "210101"|Impor == "210104",arr.ind = TRUE)[,"row"]


#Tomar el valor que nos interesa
Valor_importaciones=as.data.frame(t(Impor[c(n_fila,n_fila2[[1]],n_fila2[[2]]),(n_col1[1]:n_col2[1])]))
Valor_importaciones=as.data.frame(lapply(Valor_importaciones, as.numeric))
Valor_importaciones=Valor_importaciones %>%
  mutate(suma=Valor_importaciones[,2]+Valor_importaciones[,3])




Valor_Bovino=cbind(Valor_Bovino,Valor_exportaciones[1],Valor_importaciones[1],Valor_exportaciones[,4],Valor_importaciones[,4])
colnames(Valor_Bovino)=c("TOTAL","MACHOS","HEMBRAS","TERNEROS","EXPORTACION_CANAL","EXPO_VIVO","IMPO_VIVO","EXPO_CARNE","IMPO_CARNE")
Valor_Bovino=as.data.frame(lapply(Valor_Bovino, as.numeric))

return(Valor_Bovino)
}
