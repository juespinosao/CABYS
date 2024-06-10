#' @export
# Pesca
# Cargar la biblioteca readxl

f_Silvicultura_complemento<-function(directorio,mes,anio){

  #Cargar librerias
  library(readxl)
  library(dplyr)
  library(openxlsx)
  library(zoo)

trimestre=f_trimestre(mes)

  carpeta=nombre_carpeta(mes,anio)
  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")

  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Silvicultura","NOMBRE"]
  hojas=excel_sheets(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/Silvicultura/",archivo))

  hoja_final <- hojas[grepl("BD", hojas) ]

  Silvicultura<- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/Silvicultura/",archivo),
                   sheet = hoja_final,startRow = 8)
Silvicultura$MES=toupper(Silvicultura$MES)
  Silvicultura_tabla=Silvicultura %>%
    group_by(AÑO,TRIMESTRE,CLASE.DE.PRODUCTO)%>%
    filter(AÑO>(anio-3)) %>%
    summarise(suma=sum(VOLUMEN.M3))%>%
    as.data.frame()

troncos=Silvicultura_tabla %>% filter(CLASE.DE.PRODUCTO=="Maderable")
forestales=Silvicultura_tabla %>% filter(CLASE.DE.PRODUCTO=="No Maderable")


#var_anual=troncos$suma/lag(troncos$suma,4)*100-100
#troncos_ant=lag(troncos$suma,4)
#tamaño=length(troncos$suma)
#Estado <- rep("",tamaño)
#
#for (i in seq(4, tamaño, by = 4)) {
#  Estado[i] <- (sum(troncos$suma[(i-3):i]) / sum(troncos_ant[(i-3):i]))*100-100  # Realiza la suma y división
#}
#Estado=as.numeric(Estado)
#
#cuadro_troncos=data.frame(var_anual[c(8+trimestre)],var_anual[c(12+trimestre)],Estado[c(8+trimestre)],
#                                Estado[c(12+trimestre)])
#
#
#var_anual=forestales$suma/lag(forestales$suma,4)*100-100
#forestales_ant=lag(forestales$suma,4)
#tamaño=length(forestales$suma)
#Estado <- rep("",tamaño)
#
#for (i in seq(4, tamaño, by = 4)) {
#  Estado[i] <- (sum(forestales$suma[(i-3):i]) / sum(forestales_ant[(i-3):i]))*100-100  # Realiza la suma y división
#}
#Estado=as.numeric(Estado)
#
#cuadro_forestales=data.frame(var_anual[c(8+trimestre)],var_anual[c(12+trimestre)],Estado[c(8+trimestre)],
#                          Estado[c(12+trimestre)])
#
#nuevos_datos=bind_rows(cuadro_troncos,cuadro_forestales)
#



# Agregar a las columnas --------------------------------------------------
library(openxlsx)


#Crear el nombre de las carpetas del mes anterior y el actual
carpeta_actual=nombre_carpeta(mes,anio)
if(mes==3){

  carpeta_anterior=nombre_carpeta(12,(anio-1))
  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio-1,"/",carpeta_anterior,"/Doc/Nombres_archivos_",nombres_meses[12],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Silvicultura2","NOMBRE"]
  entrada=paste0(directorio,"/ISE/",anio-1,"/",carpeta_anterior,"/Data/consolidado_ISE/Silvicultura/",archivo)

}else{
  carpeta_anterior=nombre_carpeta(mes-3,anio)
  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_anterior,"/Doc/Nombres_archivos_",nombres_meses[mes-3],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Silvicultura2","NOMBRE"]
  entrada=paste0(directorio,"/ISE/",anio,"/",carpeta_anterior,"/Data/consolidado_ISE/Silvicultura/",archivo)

}

carpeta_actual=nombre_carpeta(mes,anio)
nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Silvicultura2","NOMBRE"]


#Dirección de entrada del archivo ZG_pecuario del mes anterior y donde se va a guardar el siguiente
salida=paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Data/consolidado_ISE/Silvicultura/",archivo)

# Cargar el archivo de entrada
wb <- loadWorkbook(entrada)



# Silvicultura hoja -----------------------------------------------------------------

data <- read.xlsx(wb, sheet = "Silvicultura producción", colNames = TRUE,startRow = 3)
ultima_fila=nrow(data)
fila=which(data$Año==(anio-2))
fila_trim_ant=max(which(data$Trimestre==trimestre))
filas=tail(which(data$Trimestre==trimestre),3)
datos_nuevos=data.frame(periodo=anio,trim=trimestre)
writeData(wb, sheet = "Silvicultura producción", x = datos_nuevos,colNames = FALSE,startCol = "A", startRow = (ultima_fila[1]+6))
writeData(wb, sheet = "Silvicultura producción", x = troncos$suma,colNames = FALSE,startCol = "C", startRow = (fila[1]+5))
writeData(wb, sheet = "Silvicultura producción", x = (data[fila_trim_ant,"Leña"]*1.005),colNames = FALSE,startCol = "D", startRow = (ultima_fila[1]+6))
writeData(wb, sheet = "Silvicultura producción", x = forestales$suma,colNames = FALSE,startCol = "E", startRow = (fila[1]+5))
if(length(troncos$suma)-length(forestales$suma)>0){
  writeFormula(wb, sheet ="Silvicultura producción" , x = paste0("AVERAGE(E",(filas[1]+5),",E",(filas[2]+5),",E",(filas[3]+5),")") ,startCol = "E", startRow = (ultima_fila[1]+6))

}
writeFormula(wb, sheet = "Silvicultura producción", x = paste0("C",(ultima_fila[1]+6),"/C",(fila_trim_ant[1]+5),"*100-100"),startCol = "G", startRow = (ultima_fila[1]+6))
writeFormula(wb, sheet = "Silvicultura producción", x = paste0("D",(ultima_fila[1]+6),"/D",(fila_trim_ant[1]+5),"*100-100"),startCol = "H", startRow = (ultima_fila[1]+6))
writeFormula(wb, sheet = "Silvicultura producción", x = paste0("E",(ultima_fila[1]+6),"/E",(fila_trim_ant[1]+5),"*100-100"),startCol = "I", startRow = (ultima_fila[1]+6))

writeFormula(wb, sheet = "Silvicultura producción", x = paste0("C",(ultima_fila[1]+6),"/AVERAGE(C74:C77)*100"),startCol = "K", startRow = (ultima_fila[1]+6))
writeFormula(wb, sheet = "Silvicultura producción", x = paste0("D",(ultima_fila[1]+6),"/AVERAGE(D74:D77)*100"),startCol = "L", startRow = (ultima_fila[1]+6))
writeFormula(wb, sheet = "Silvicultura producción", x = paste0("E",(ultima_fila[1]+6),"/AVERAGE(E74:E77)*100"),startCol = "M", startRow = (ultima_fila[1]+6))


writeFormula(wb, sheet = "Silvicultura producción", x = paste0("(O2/4)*K",(ultima_fila[1]+6),"/100"),startCol = "O", startRow = (ultima_fila[1]+6))
writeFormula(wb, sheet = "Silvicultura producción", x = paste0("(P2/4)*L",(ultima_fila[1]+6),"/100"),startCol = "P", startRow = (ultima_fila[1]+6))
writeFormula(wb, sheet = "Silvicultura producción", x = paste0("(Q2/4)*M",(ultima_fila[1]+6),"/100"),startCol = "Q", startRow = (ultima_fila[1]+6))

writeFormula(wb, sheet = "Silvicultura producción", x = paste0("O",(ultima_fila[1]+6),"*S2"),startCol = "S", startRow = (ultima_fila[1]+6))
writeFormula(wb, sheet = "Silvicultura producción", x = paste0("P",(ultima_fila[1]+6),"*T2"),startCol = "T", startRow = (ultima_fila[1]+6))
writeFormula(wb, sheet = "Silvicultura producción", x = paste0("Q",(ultima_fila[1]+6),"*U2"),startCol = "U", startRow = (ultima_fila[1]+6))

writeFormula(wb, sheet = "Silvicultura producción", x = paste0("SUM(S",(ultima_fila[1]+6),":U",(ultima_fila[1]+6),")"),startCol = "AA", startRow = (ultima_fila[1]+6))
writeFormula(wb, sheet = "Silvicultura producción", x = paste0("AA",(ultima_fila[1]+6),"/AA",(fila_trim_ant[1]+5),"*100-100"),startCol = "AB", startRow = (ultima_fila[1]+6))

# Guardar el libro --------------------------------------------------------


if (!file.exists(salida)) {
  saveWorkbook(wb, file = salida)
} else {
  saveWorkbook(wb, file = salida,overwrite= TRUE)
}
}
