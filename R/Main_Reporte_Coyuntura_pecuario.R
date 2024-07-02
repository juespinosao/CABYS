#' @export


Reporte_coyuntura<-function(directorio,mes,anio){

#Cargar librerias
library(openxlsx)

#Crear el nombre de las carpetas del mes anterior y el actual
carpeta_actual=nombre_carpeta(mes,anio)
if(mes==1){
  carpeta_anterior=nombre_carpeta(12,(anio-1))
  entrada=paste0(directorio,"/ISE/",anio-1,"/",carpeta_anterior,"/Results/Reporte Coyuntura Pecuario ISE_",nombres_meses[12],"_",(anio-1),".xlsx")

}else{
  carpeta_anterior=nombre_carpeta(mes-1,anio)
  entrada=paste0(directorio,"/ISE/",anio,"/",carpeta_anterior,"/Results/Reporte Coyuntura Pecuario ISE_",nombres_meses[mes-1],"_",anio,".xlsx")

}

carpeta_actual=nombre_carpeta(mes,anio)

#Dirección de entrada del archivo ZG_pecuario del mes anterior y donde se va a guardar el siguiente
salida=paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Results/Reporte Coyuntura Pecuario ISE_",nombres_meses[mes],"_",anio,".xlsx")

# Cargar el archivo de entrada
wb <- loadWorkbook(entrada)



# Precios -----------------------------------------------------------------

data <- read.xlsx(wb, sheet = "Precios", colNames = TRUE,startRow = 5)
ultima_fila_precios=nrow(data)

data$Período <- as.Date(data$Período, origin = "1899-12-30")
data$Período=format(data$Período, "%Y-%m")
fila_enero_precios=which(data$Período== paste0(anio,"-01"))
fila_enero_ant_precios=which(data$Período== paste0((anio-1),"-01"))
fila_anterior_precios=which(data$Período== paste0((anio-1),"-",mes_0[mes]))

Precios=f_Precios(directorio,mes,anio)
IPP=Precios[[1]]
Porkcol=Precios[[2]]
USP=Precios[[3]]
IPC=Precios[[4]]
valor_fecha=as.integer(as.Date(paste0(1,"/",mes_0[mes],"/",anio), format = "%d/%m/%Y") - as.Date("1899-12-30"))
writeData(wb, sheet = "Precios", x = valor_fecha,colNames = FALSE,startCol = "A", startRow = (ultima_fila_precios+6))
writeData(wb, sheet = "Precios", x = IPP[,1:2],colNames = FALSE,startCol = "B", startRow = (ultima_fila_precios+5))
if (mes==1) {
  writeData(wb, sheet = "Precios", x = Porkcol,colNames = FALSE,startCol = "D", startRow = (ultima_fila_precios+6))
}else{
  writeData(wb, sheet = "Precios", x = Porkcol,colNames = FALSE,startCol = "D", startRow = (fila_enero_precios+5))
}

writeData(wb, sheet = "Precios", x = IPP[,3:4],colNames = FALSE,startCol = "F", startRow = (ultima_fila_precios+5))
if (mes==1) {
  writeData(wb, sheet = "Precios", x = USP,colNames = FALSE,startCol = "H", startRow = (ultima_fila_precios+6))
}else{
  writeData(wb, sheet = "Precios", x = USP,colNames = FALSE,startCol = "H", startRow = (fila_enero_precios+5))
}
writeData(wb, sheet = "Precios", x = IPP[,5:8],colNames = FALSE,startCol = "J", startRow = (ultima_fila_precios+5))
writeData(wb, sheet = "Precios", x = IPC,colNames = FALSE,startCol = "N", startRow = (ultima_fila_precios+6))
writeData(wb, sheet = "Precios", x = IPP[,9],colNames = FALSE,startCol = "P", startRow = (ultima_fila_precios+5))
if(mes-length(USP$Valor_sinbon)>0){
for (i in 1:(mes-length(USP$Valor_sinbon))) {

  writeFormula(wb, sheet ="Precios" , x = paste0("AVERAGE(H",(fila_enero_precios[1]+5),":H",ultima_fila_precios+6-i,")") ,startCol = "H", startRow = (ultima_fila_precios+7-i))

}
}
if(mes-length(USP$Valor_total)>0){
  for (i in 1:(mes-length(USP$Valor_total))) {

    writeFormula(wb, sheet ="Precios" , x = paste0("AVERAGE(I",(fila_enero_precios[1]+5),":I",ultima_fila_precios+6-i,")") ,startCol = "I", startRow = (ultima_fila_precios+7-i))

  }
}



tasa_anual <- c(paste('IFERROR(B',ultima_fila_precios+6,'/B',fila_anterior_precios+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(C',ultima_fila_precios+6,'/C',fila_anterior_precios+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(D',ultima_fila_precios+6,'/D',fila_anterior_precios+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(E',ultima_fila_precios+6,'/E',fila_anterior_precios+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(F',ultima_fila_precios+6,'/F',fila_anterior_precios+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(G',ultima_fila_precios+6,'/G',fila_anterior_precios+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(H',ultima_fila_precios+6,'/H',fila_anterior_precios+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(I',ultima_fila_precios+6,'/I',fila_anterior_precios+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(J',ultima_fila_precios+6,'/J',fila_anterior_precios+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(K',ultima_fila_precios+6,'/K',fila_anterior_precios+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(L',ultima_fila_precios+6,'/L',fila_anterior_precios+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(M',ultima_fila_precios+6,'/M',fila_anterior_precios+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(N',ultima_fila_precios+6,'/N',fila_anterior_precios+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(O',ultima_fila_precios+6,'/O',fila_anterior_precios+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(P',ultima_fila_precios+6,'/P',fila_anterior_precios+5,'*100-100,','"*")',sep = ""))


for (i in 28:42) {
  writeFormula(wb, sheet ="Precios" , x = tasa_anual[i-27] ,startCol = i, startRow = ultima_fila_precios+6)
}

if(mes==1){
  tasa_corrido <- c(paste('IFERROR(SUM(B',ultima_fila_precios+6,':B',ultima_fila_precios+6,')/SUM(B',fila_enero_ant_precios+5,':B',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(C',ultima_fila_precios+6,':C',ultima_fila_precios+6,')/SUM(C',fila_enero_ant_precios+5,':C',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(D',ultima_fila_precios+6,':D',ultima_fila_precios+6,')/SUM(D',fila_enero_ant_precios+5,':D',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(E',ultima_fila_precios+6,':E',ultima_fila_precios+6,')/SUM(E',fila_enero_ant_precios+5,':E',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(F',ultima_fila_precios+6,':F',ultima_fila_precios+6,')/SUM(F',fila_enero_ant_precios+5,':F',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(G',ultima_fila_precios+6,':G',ultima_fila_precios+6,')/SUM(G',fila_enero_ant_precios+5,':G',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(H',ultima_fila_precios+6,':H',ultima_fila_precios+6,')/SUM(H',fila_enero_ant_precios+5,':H',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(I',ultima_fila_precios+6,':I',ultima_fila_precios+6,')/SUM(I',fila_enero_ant_precios+5,':I',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(J',ultima_fila_precios+6,':J',ultima_fila_precios+6,')/SUM(J',fila_enero_ant_precios+5,':J',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(K',ultima_fila_precios+6,':K',ultima_fila_precios+6,')/SUM(K',fila_enero_ant_precios+5,':K',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(L',ultima_fila_precios+6,':L',ultima_fila_precios+6,')/SUM(L',fila_enero_ant_precios+5,':L',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(M',ultima_fila_precios+6,':M',ultima_fila_precios+6,')/SUM(M',fila_enero_ant_precios+5,':M',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(N',ultima_fila_precios+6,':N',ultima_fila_precios+6,')/SUM(N',fila_enero_ant_precios+5,':N',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(O',ultima_fila_precios+6,':O',ultima_fila_precios+6,')/SUM(O',fila_enero_ant_precios+5,':O',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(P',ultima_fila_precios+6,':P',ultima_fila_precios+6,')/SUM(P',fila_enero_ant_precios+5,':P',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = "")

  )

}else{
  tasa_corrido <- c(paste('IFERROR(SUM(B',fila_enero_precios+5,':B',ultima_fila_precios+6,')/SUM(B',fila_enero_ant_precios+5,':B',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(C',fila_enero_precios+5,':C',ultima_fila_precios+6,')/SUM(C',fila_enero_ant_precios+5,':C',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(D',fila_enero_precios+5,':D',ultima_fila_precios+6,')/SUM(D',fila_enero_ant_precios+5,':D',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(E',fila_enero_precios+5,':E',ultima_fila_precios+6,')/SUM(E',fila_enero_ant_precios+5,':E',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(F',fila_enero_precios+5,':F',ultima_fila_precios+6,')/SUM(F',fila_enero_ant_precios+5,':F',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(G',fila_enero_precios+5,':G',ultima_fila_precios+6,')/SUM(G',fila_enero_ant_precios+5,':G',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(H',fila_enero_precios+5,':H',ultima_fila_precios+6,')/SUM(H',fila_enero_ant_precios+5,':H',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(I',fila_enero_precios+5,':I',ultima_fila_precios+6,')/SUM(I',fila_enero_ant_precios+5,':I',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(J',fila_enero_precios+5,':J',ultima_fila_precios+6,')/SUM(J',fila_enero_ant_precios+5,':J',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(K',fila_enero_precios+5,':K',ultima_fila_precios+6,')/SUM(K',fila_enero_ant_precios+5,':K',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(L',fila_enero_precios+5,':L',ultima_fila_precios+6,')/SUM(L',fila_enero_ant_precios+5,':L',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(M',fila_enero_precios+5,':M',ultima_fila_precios+6,')/SUM(M',fila_enero_ant_precios+5,':M',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(N',fila_enero_precios+5,':N',ultima_fila_precios+6,')/SUM(N',fila_enero_ant_precios+5,':N',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(O',fila_enero_precios+5,':O',ultima_fila_precios+6,')/SUM(O',fila_enero_ant_precios+5,':O',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(P',fila_enero_precios+5,':P',ultima_fila_precios+6,')/SUM(P',fila_enero_ant_precios+5,':P',fila_enero_ant_precios+4+mes,')*100-100,','"*")',sep = "")

  )

}

for (i in 44:58) {
  writeFormula(wb, sheet ="Precios" , x = tasa_corrido[i-43] ,startCol = i, startRow = ultima_fila_precios+6)
}

if (mes %in% c(3,6,9,12)){
  tasa_trimestre <- c(paste('IFERROR(SUM(B',ultima_fila_precios+4,':B',ultima_fila_precios+6,')/SUM(B',fila_anterior_precios+3,':B',fila_anterior_precios+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(C',ultima_fila_precios+4,':C',ultima_fila_precios+6,')/SUM(C',fila_anterior_precios+3,':C',fila_anterior_precios+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(D',ultima_fila_precios+4,':D',ultima_fila_precios+6,')/SUM(D',fila_anterior_precios+3,':D',fila_anterior_precios+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(E',ultima_fila_precios+4,':E',ultima_fila_precios+6,')/SUM(E',fila_anterior_precios+3,':E',fila_anterior_precios+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(F',ultima_fila_precios+4,':F',ultima_fila_precios+6,')/SUM(F',fila_anterior_precios+3,':F',fila_anterior_precios+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(G',ultima_fila_precios+4,':G',ultima_fila_precios+6,')/SUM(G',fila_anterior_precios+3,':G',fila_anterior_precios+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(H',ultima_fila_precios+4,':H',ultima_fila_precios+6,')/SUM(H',fila_anterior_precios+3,':H',fila_anterior_precios+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(I',ultima_fila_precios+4,':I',ultima_fila_precios+6,')/SUM(I',fila_anterior_precios+3,':I',fila_anterior_precios+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(J',ultima_fila_precios+4,':J',ultima_fila_precios+6,')/SUM(J',fila_anterior_precios+3,':J',fila_anterior_precios+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(K',ultima_fila_precios+4,':K',ultima_fila_precios+6,')/SUM(K',fila_anterior_precios+3,':K',fila_anterior_precios+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(L',ultima_fila_precios+4,':L',ultima_fila_precios+6,')/SUM(L',fila_anterior_precios+3,':L',fila_anterior_precios+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(M',ultima_fila_precios+4,':M',ultima_fila_precios+6,')/SUM(M',fila_anterior_precios+3,':M',fila_anterior_precios+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(N',ultima_fila_precios+4,':N',ultima_fila_precios+6,')/SUM(N',fila_anterior_precios+3,':N',fila_anterior_precios+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(P',ultima_fila_precios+4,':P',ultima_fila_precios+6,')/SUM(P',fila_anterior_precios+3,':P',fila_anterior_precios+5,')*100-100,','"*")',sep = "")
  )

  for (i in 60:73) {
    writeFormula(wb, sheet ="Precios" , x = tasa_trimestre[i-59] ,startCol = i, startRow = ultima_fila_precios+6)
  }

}else{
  cat("Este mes no se actualiza trimestre")
}


addStyle(wb, sheet = "Precios",style=cbp,rows = (ultima_fila_precios+6),cols = 1)
addStyle(wb, sheet = "Precios",style=cbn4,rows = (ultima_fila_precios+6),cols = 2:9)
addStyle(wb, sheet = "Precios",style=rn5,rows = (ultima_fila_precios+6),cols = 10:16)
addStyle(wb, sheet = "Precios",style=rn4,rows = (ultima_fila_precios+6),cols = 28:73)


# Bovino_kilo_en_pie ------------------------------------------------------

data <- read.xlsx(wb, sheet = "Bovino kilo en pie", colNames = TRUE,startRow = 6)
ultima_fila=nrow(data)

data$Período <- as.Date(data$Período, origin = "1899-12-30")
data$Período=format(data$Período, "%Y-%m")
fila_enero=which(data$Período== paste0(anio,"-01"))
fila_enero_ant=which(data$Período== paste0((anio-1),"-01"))
fila_anterior=which(data$Período== paste0((anio-1),"-",mes_0[mes]))

valor_Bovino=f_Bovino(directorio,mes,anio)
valor_fecha=as.integer(as.Date(paste0(1,"/",mes_0[mes],"/",anio), format = "%d/%m/%Y") - as.Date("1899-12-30"))
writeData(wb, sheet = "Bovino kilo en pie", x = valor_fecha,colNames = FALSE,startCol = "A", startRow = (ultima_fila+7))
if(mes==1){
  writeData(wb, sheet = "Bovino kilo en pie", x = valor_Bovino,colNames = FALSE,startCol = "B", startRow = (ultima_fila+7))
}else{
  writeData(wb, sheet = "Bovino kilo en pie", x = valor_Bovino,colNames = FALSE,startCol = "B", startRow = (fila_enero[1]+6))
}

#Añadir formulas
participacion <- c(paste0("SUM(N",ultima_fila+7,":Q",ultima_fila+7,")"),paste0("C",ultima_fila+7,"/B",ultima_fila+7,"*100"),
       paste0("D",ultima_fila+7,"/B",ultima_fila+7,"*100"),paste0("E",ultima_fila+7,"/B",ultima_fila+7,"*100"),paste0("F",ultima_fila+7,"/B",ultima_fila+7,"*100")) ## skip header row

for (i in 13:17) {
  writeFormula(wb, sheet ="Bovino kilo en pie" , x = participacion[i-12] ,startCol = i, startRow = ultima_fila+7)
}

tasa_anual <- c(paste('IFERROR(B',ultima_fila+7,'/B',fila_anterior+6,'*100-100,','"*")',sep = ""),
                      paste('IFERROR(C',ultima_fila+7,'/C',fila_anterior+6,'*100-100,','"*")',sep = ""),
                      paste('IFERROR(D',ultima_fila+7,'/D',fila_anterior+6,'*100-100,','"*")',sep = ""),
                      paste('IFERROR(E',ultima_fila+7,'/E',fila_anterior+6,'*100-100,','"*")',sep = ""),
                      paste('IFERROR(F',ultima_fila+7,'/F',fila_anterior+6,'*100-100,','"*")',sep = ""),
                      paste('IFERROR(G',ultima_fila+7,'/G',fila_anterior+6,'*100-100,','"*")',sep = ""),
                      paste('IFERROR(H',ultima_fila+7,'/H',fila_anterior+6,'*100-100,','"*")',sep = ""),
                      paste('IFERROR(I',ultima_fila+7,'/I',fila_anterior+6,'*100-100,','"*")',sep = ""),
                      paste('IFERROR(J',ultima_fila+7,'/J',fila_anterior+6,'*100-100,','"*")',sep = ""))


for (i in 19:27) {
  writeFormula(wb, sheet ="Bovino kilo en pie" , x = tasa_anual[i-18] ,startCol = i, startRow = ultima_fila+7)
}

if(mes==1){
  tasa_corrido <- c(paste('IFERROR(SUM(B',ultima_fila+7,':B',ultima_fila+7,')/SUM(B',fila_enero_ant+6,':B',fila_enero_ant+5+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(C',ultima_fila+7,':C',ultima_fila+7,')/SUM(C',fila_enero_ant+6,':C',fila_enero_ant+5+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(D',ultima_fila+7,':D',ultima_fila+7,')/SUM(D',fila_enero_ant+6,':D',fila_enero_ant+5+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(E',ultima_fila+7,':E',ultima_fila+7,')/SUM(E',fila_enero_ant+6,':E',fila_enero_ant+5+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(F',ultima_fila+7,':F',ultima_fila+7,')/SUM(F',fila_enero_ant+6,':F',fila_enero_ant+5+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(G',ultima_fila+7,':G',ultima_fila+7,')/SUM(G',fila_enero_ant+6,':G',fila_enero_ant+5+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(H',ultima_fila+7,':H',ultima_fila+7,')/SUM(H',fila_enero_ant+6,':H',fila_enero_ant+5+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(I',ultima_fila+7,':I',ultima_fila+7,')/SUM(I',fila_enero_ant+6,':I',fila_enero_ant+5+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(J',ultima_fila+7,':J',ultima_fila+7,')/SUM(J',fila_enero_ant+6,':J',fila_enero_ant+5+mes,')*100-100,','"*")',sep = ""))

}else{
  tasa_corrido <- c(paste('IFERROR(SUM(B',fila_enero+6,':B',ultima_fila+7,')/SUM(B',fila_enero_ant+6,':B',fila_enero_ant+5+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(C',fila_enero+6,':C',ultima_fila+7,')/SUM(C',fila_enero_ant+6,':C',fila_enero_ant+5+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(D',fila_enero+6,':D',ultima_fila+7,')/SUM(D',fila_enero_ant+6,':D',fila_enero_ant+5+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(E',fila_enero+6,':E',ultima_fila+7,')/SUM(E',fila_enero_ant+6,':E',fila_enero_ant+5+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(F',fila_enero+6,':F',ultima_fila+7,')/SUM(F',fila_enero_ant+6,':F',fila_enero_ant+5+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(G',fila_enero+6,':G',ultima_fila+7,')/SUM(G',fila_enero_ant+6,':G',fila_enero_ant+5+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(H',fila_enero+6,':H',ultima_fila+7,')/SUM(H',fila_enero_ant+6,':H',fila_enero_ant+5+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(I',fila_enero+6,':I',ultima_fila+7,')/SUM(I',fila_enero_ant+6,':I',fila_enero_ant+5+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(J',fila_enero+6,':J',ultima_fila+7,')/SUM(J',fila_enero_ant+6,':J',fila_enero_ant+5+mes,')*100-100,','"*")',sep = ""))

}

for (i in 29:37) {
  writeFormula(wb, sheet ="Bovino kilo en pie" , x = tasa_corrido[i-28] ,startCol = i, startRow = ultima_fila+7)
}

if (mes %in% c(3,6,9,12)){
  tasa_trimestre <- c(paste('IFERROR(SUM(B',ultima_fila+5,':B',ultima_fila+7,')/SUM(B',fila_anterior+4,':B',fila_anterior+6,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(C',ultima_fila+5,':C',ultima_fila+7,')/SUM(C',fila_anterior+4,':C',fila_anterior+6,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(D',ultima_fila+5,':D',ultima_fila+7,')/SUM(D',fila_anterior+4,':D',fila_anterior+6,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(E',ultima_fila+5,':E',ultima_fila+7,')/SUM(E',fila_anterior+4,':E',fila_anterior+6,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(F',ultima_fila+5,':F',ultima_fila+7,')/SUM(F',fila_anterior+4,':F',fila_anterior+6,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(G',ultima_fila+5,':G',ultima_fila+7,')/SUM(G',fila_anterior+4,':G',fila_anterior+6,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(H',ultima_fila+5,':H',ultima_fila+7,')/SUM(H',fila_anterior+4,':H',fila_anterior+6,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(I',ultima_fila+5,':I',ultima_fila+7,')/SUM(I',fila_anterior+4,':I',fila_anterior+6,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(J',ultima_fila+5,':J',ultima_fila+7,')/SUM(J',fila_anterior+4,':J',fila_anterior+6,')*100-100,','"*")',sep = ""))

  for (i in 40:48) {
    writeFormula(wb, sheet ="Bovino kilo en pie" , x = tasa_trimestre[i-39] ,startCol = i, startRow = ultima_fila+7)
  }

}else{
  cat("Este mes no se actualiza trimestre")
}

contribucion <- c(paste0("SUM(AY",ultima_fila+7,":BB",ultima_fila+7,")"),
                  paste0("N",fila_anterior+6,"*T",ultima_fila+7,"/100"),
                  paste0("O",fila_anterior+6,"*U",ultima_fila+7,"/100"),
                  paste0("P",fila_anterior+6,"*V",ultima_fila+7,"/100"),
                  paste0("Q",fila_anterior+6,"*W",ultima_fila+7,"/100")) ## skip header row

for (i in 50:54) {
  writeFormula(wb, sheet ="Bovino kilo en pie" , x = contribucion[i-49] ,startCol = i, startRow = ultima_fila+7)
}




addStyle(wb, sheet = "Bovino kilo en pie",style=cbp,rows = (ultima_fila+7),cols = 1)
addStyle(wb, sheet = "Bovino kilo en pie",style=cbn,rows = (ultima_fila+7),cols = 2:10)
addStyle(wb, sheet = "Bovino kilo en pie",style=rn4,rows = (ultima_fila+7),cols = 13:17)
addStyle(wb, sheet = "Bovino kilo en pie",style=cbn2,rows = (ultima_fila+7),cols = 19:27)
addStyle(wb, sheet = "Bovino kilo en pie",style=rn4,rows = (ultima_fila+7),cols = 29:37)
addStyle(wb, sheet = "Bovino kilo en pie",style=rn4,rows = (ultima_fila+7),cols = 40:48)
addStyle(wb, sheet = "Bovino kilo en pie",style=rn4,rows = (ultima_fila+7),cols = 50:54)



# Bovino_cabezas ------------------------------------------------------

data <- read.xlsx(wb, sheet = "Bovino cabezas", colNames = TRUE,startRow = 5)
ultima_fila=nrow(data)

data$Período <- as.Date(data$Período, origin = "1899-12-30")
data$Período=format(data$Período, "%Y-%m")
fila_enero=which(data$Período== paste0(anio,"-01"))
fila_enero_ant=which(data$Período== paste0((anio-1),"-01"))
fila_anterior=which(data$Período== paste0((anio-1),"-",mes_0[mes]))

valor_Bovino=f_Bovino_cabezas(directorio,mes,anio)
valor_fecha=as.integer(as.Date(paste0(1,"/",mes_0[mes],"/",anio), format = "%d/%m/%Y") - as.Date("1899-12-30"))
writeData(wb, sheet = "Bovino cabezas", x = valor_fecha,colNames = FALSE,startCol = "A", startRow = (ultima_fila+6))

if(mes==1){
  writeData(wb, sheet = "Bovino cabezas", x = valor_Bovino,colNames = FALSE,startCol = "B", startRow = (ultima_fila+6))
}else{
  writeData(wb, sheet = "Bovino cabezas", x = valor_Bovino,colNames = FALSE,startCol = "B", startRow = (fila_enero[1]+5))
}

#Añadir formulas
participacion <- c(paste0("IFERROR(SUM(M",ultima_fila+6,":P",ultima_fila+6,"),0)"),
                   paste0("IFERROR(C",ultima_fila+6,"/B",ultima_fila+6,"*100,0)"),
                   paste0("IFERROR(D",ultima_fila+6,"/B",ultima_fila+6,"*100,0)"),
                   paste0("IFERROR(E",ultima_fila+6,"/B",ultima_fila+6,"*100,0)"),
                   paste0("IFERROR(F",ultima_fila+6,"/B",ultima_fila+6,"*100,0)")) ## skip header row

for (i in 12:16) {
  writeFormula(wb, sheet ="Bovino cabezas" , x = participacion[i-11] ,startCol = i, startRow = ultima_fila+6)
}

tasa_anual <- c(paste('IFERROR(B',ultima_fila+6,'/B',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(C',ultima_fila+6,'/C',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(D',ultima_fila+6,'/D',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(E',ultima_fila+6,'/E',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(F',ultima_fila+6,'/F',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(G',ultima_fila+6,'/G',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(H',ultima_fila+6,'/H',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(I',ultima_fila+6,'/I',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(J',ultima_fila+6,'/J',fila_anterior+5,'*100-100,','"*")',sep = ""))


for (i in 18:26) {
  writeFormula(wb, sheet ="Bovino cabezas" , x = tasa_anual[i-17] ,startCol = i, startRow = ultima_fila+6)
}

if(mes==1){
  tasa_corrido <- c(paste('IFERROR(SUM(B',ultima_fila+6,':B',ultima_fila+6,')/SUM(B',fila_enero_ant+5,':B',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(C',ultima_fila+6,':C',ultima_fila+6,')/SUM(C',fila_enero_ant+5,':C',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(D',ultima_fila+6,':D',ultima_fila+6,')/SUM(D',fila_enero_ant+5,':D',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(E',ultima_fila+6,':E',ultima_fila+6,')/SUM(E',fila_enero_ant+5,':E',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(F',ultima_fila+6,':F',ultima_fila+6,')/SUM(F',fila_enero_ant+5,':F',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(G',ultima_fila+6,':G',ultima_fila+6,')/SUM(G',fila_enero_ant+5,':G',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(H',ultima_fila+6,':H',ultima_fila+6,')/SUM(H',fila_enero_ant+5,':H',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(I',ultima_fila+6,':I',ultima_fila+6,')/SUM(I',fila_enero_ant+5,':I',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(J',ultima_fila+6,':J',ultima_fila+6,')/SUM(J',fila_enero_ant+5,':J',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""))

}else{
  tasa_corrido <- c(paste('IFERROR(SUM(B',fila_enero+5,':B',ultima_fila+6,')/SUM(B',fila_enero_ant+5,':B',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(C',fila_enero+5,':C',ultima_fila+6,')/SUM(C',fila_enero_ant+5,':C',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(D',fila_enero+5,':D',ultima_fila+6,')/SUM(D',fila_enero_ant+5,':D',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(E',fila_enero+5,':E',ultima_fila+6,')/SUM(E',fila_enero_ant+5,':E',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(F',fila_enero+5,':F',ultima_fila+6,')/SUM(F',fila_enero_ant+5,':F',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(G',fila_enero+5,':G',ultima_fila+6,')/SUM(G',fila_enero_ant+5,':G',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(H',fila_enero+5,':H',ultima_fila+6,')/SUM(H',fila_enero_ant+5,':H',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(I',fila_enero+5,':I',ultima_fila+6,')/SUM(I',fila_enero_ant+5,':I',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(J',fila_enero+5,':J',ultima_fila+6,')/SUM(J',fila_enero_ant+5,':J',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""))

}

for (i in 28:36) {
  writeFormula(wb, sheet ="Bovino cabezas" , x = tasa_corrido[i-27] ,startCol = i, startRow = ultima_fila+6)
}

if (mes %in% c(3,6,9,12)){
  tasa_trimestre <- c(paste('IFERROR(SUM(B',ultima_fila+4,':B',ultima_fila+6,')/SUM(B',fila_anterior+3,':B',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(C',ultima_fila+4,':C',ultima_fila+6,')/SUM(C',fila_anterior+3,':C',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(D',ultima_fila+4,':D',ultima_fila+6,')/SUM(D',fila_anterior+3,':D',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(E',ultima_fila+4,':E',ultima_fila+6,')/SUM(E',fila_anterior+3,':E',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(F',ultima_fila+4,':F',ultima_fila+6,')/SUM(F',fila_anterior+3,':F',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(G',ultima_fila+4,':G',ultima_fila+6,')/SUM(G',fila_anterior+3,':G',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(H',ultima_fila+4,':H',ultima_fila+6,')/SUM(H',fila_anterior+3,':H',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(I',ultima_fila+4,':I',ultima_fila+6,')/SUM(I',fila_anterior+3,':I',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(J',ultima_fila+4,':J',ultima_fila+6,')/SUM(J',fila_anterior+3,':J',fila_anterior+5,')*100-100,','"*")',sep = ""))

  for (i in 38:46) {
    writeFormula(wb, sheet ="Bovino cabezas" , x = tasa_trimestre[i-37] ,startCol = i, startRow = ultima_fila+6)
  }

}else{
  cat("Este mes no se actualiza trimestre")
}


addStyle(wb, sheet = "Bovino cabezas",style=cbp,rows = (ultima_fila+6),cols = 1)
addStyle(wb, sheet = "Bovino cabezas",style=cbn,rows = (ultima_fila+6),cols = 2:10)
addStyle(wb, sheet = "Bovino cabezas",style=rn4,rows = (ultima_fila+6),cols = 12:16)
addStyle(wb, sheet = "Bovino cabezas",style=rn4,rows = (ultima_fila+6),cols = 18:26)
addStyle(wb, sheet = "Bovino cabezas",style=rn4,rows = (ultima_fila+6),cols = 28:36)
addStyle(wb, sheet = "Bovino cabezas",style=cbn2,rows = (ultima_fila+6),cols = 38:46)




# CI CARNE ------------------------------------------------------

data <- read.xlsx(wb, sheet = "CI_Carne", colNames = TRUE,startRow = 5)
ultima_fila=nrow(data)

data$Período <- as.Date(data$Período, origin = "1899-12-30")
data$Período=format(data$Período, "%Y-%m")
fila_enero=which(data$Período== paste0(anio,"-01"))
fila_enero_ant=which(data$Período== paste0((anio-1),"-01"))
fila_anterior=which(data$Período== paste0((anio-1),"-",mes_0[mes]))

ganado_vacuno=f_Consumo_vacuno(directorio,mes,anio)
ganado_porcino=f_Consumo_porcino(directorio,mes,anio)
valor_fecha=as.integer(as.Date(paste0(1,"/",mes_0[mes],"/",anio), format = "%d/%m/%Y") - as.Date("1899-12-30"))
writeData(wb, sheet = "CI_Carne", x = valor_fecha,colNames = FALSE,startCol = "A", startRow = (ultima_fila+6))

if(mes==1){
  writeData(wb, sheet = "CI_Carne", x = ganado_vacuno,colNames = FALSE,startCol = "C", startRow = (ultima_fila+6))
  writeData(wb, sheet = "CI_Carne", x = ganado_porcino,colNames = FALSE,startCol = "I", startRow = (ultima_fila+6))

}else{
  writeData(wb, sheet = "CI_Carne", x = ganado_vacuno,colNames = FALSE,startCol = "C", startRow = (fila_enero[1]+5))
  writeData(wb, sheet = "CI_Carne", x = ganado_porcino,colNames = FALSE,startCol = "I", startRow = (fila_enero[1]+5))

}





#Añadir formulas

writeFormula(wb, sheet ="CI_Carne" , x = paste0("SUM(C",ultima_fila+6,":E",ultima_fila+6,")") ,startCol = "B", startRow = ultima_fila+6)
#writeFormula(wb, sheet ="CI_Carne" , x = makeHyperlinkString(sheet = "Bovino kilo en pie", row = ultima_fila+7, col = 10) ,startCol = "F", startRow = ultima_fila+6)
writeFormula(wb, sheet ="CI_Carne" , x = paste0("'Bovino kilo en pie'!J",ultima_fila+7) ,startCol = "F", startRow = ultima_fila+6)

writeFormula(wb, sheet ="CI_Carne" , x = paste0("SUM(I",ultima_fila+6,":K",ultima_fila+6,")") ,startCol = "H", startRow = ultima_fila+6)
#writeFormula(wb, sheet ="CI_Carne" , x = makeHyperlinkString(sheet = "Porcino kilo en pie", row = ultima_fila+6, col = 7) ,startCol = "L", startRow = ultima_fila+6)
writeFormula(wb, sheet ="CI_Carne" , x = paste0("'Porcino kilo en pie'!F",ultima_fila+6) ,startCol = "L", startRow = ultima_fila+6)


participacion <- c(paste0("IFERROR(SUM(O",ultima_fila+6,":Q",ultima_fila+6,"),0)"),
                   paste0("IFERROR(C",ultima_fila+6,"/B",ultima_fila+6,"*100,0)"),
                   paste0("IFERROR(D",ultima_fila+6,"/B",ultima_fila+6,"*100,0)"),
                   paste0("IFERROR(E",ultima_fila+6,"/B",ultima_fila+6,"*100,0)"),
                   "",
                   paste0("IFERROR(SUM(T",ultima_fila+6,":V",ultima_fila+6,"),0)"),
                   paste0("IFERROR(I",ultima_fila+6,"/H",ultima_fila+6,"*100,0)"),
                   paste0("IFERROR(J",ultima_fila+6,"/H",ultima_fila+6,"*100,0)"),
                   paste0("IFERROR(K",ultima_fila+6,"/H",ultima_fila+6,"*100,0)")) ## skip header row

for (i in 14:22) {
  writeFormula(wb, sheet ="CI_Carne" , x = participacion[i-13] ,startCol = i, startRow = ultima_fila+6)
}


tasa_anual <- c(paste('IFERROR(B',ultima_fila+6,'/B',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(C',ultima_fila+6,'/C',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(D',ultima_fila+6,'/D',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(E',ultima_fila+6,'/E',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(F',ultima_fila+6,'/F',fila_anterior+5,'*100-100,','"*")',sep = ""),
                "",
                paste('IFERROR(H',ultima_fila+6,'/H',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(I',ultima_fila+6,'/I',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(J',ultima_fila+6,'/J',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(K',ultima_fila+6,'/K',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(L',ultima_fila+6,'/L',fila_anterior+5,'*100-100,','"*")',sep = ""))


for (i in 24:34) {
  writeFormula(wb, sheet ="CI_Carne" , x = tasa_anual[i-23] ,startCol = i, startRow = ultima_fila+6)
}

if(mes==1){
  tasa_corrido <- c(paste('IFERROR(SUM(B',ultima_fila+6,':B',ultima_fila+6,')/SUM(B',fila_enero_ant+5,':B',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(C',ultima_fila+6,':C',ultima_fila+6,')/SUM(C',fila_enero_ant+5,':C',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(D',ultima_fila+6,':D',ultima_fila+6,')/SUM(D',fila_enero_ant+5,':D',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(E',ultima_fila+6,':E',ultima_fila+6,')/SUM(E',fila_enero_ant+5,':E',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(F',ultima_fila+6,':F',ultima_fila+6,')/SUM(F',fila_enero_ant+5,':F',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    "",
                    paste('IFERROR(SUM(H',ultima_fila+6,':H',ultima_fila+6,')/SUM(H',fila_enero_ant+5,':H',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(I',ultima_fila+6,':I',ultima_fila+6,')/SUM(I',fila_enero_ant+5,':I',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(J',ultima_fila+6,':J',ultima_fila+6,')/SUM(J',fila_enero_ant+5,':J',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(K',ultima_fila+6,':K',ultima_fila+6,')/SUM(K',fila_enero_ant+5,':K',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(L',ultima_fila+6,':L',ultima_fila+6,')/SUM(L',fila_enero_ant+5,':L',fila_enero_ant+4+mes,')*100-100,','"*")',sep = "")
  )

}else{
  tasa_corrido <- c(paste('IFERROR(SUM(B',fila_enero+5,':B',ultima_fila+6,')/SUM(B',fila_enero_ant+5,':B',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(C',fila_enero+5,':C',ultima_fila+6,')/SUM(C',fila_enero_ant+5,':C',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(D',fila_enero+5,':D',ultima_fila+6,')/SUM(D',fila_enero_ant+5,':D',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(E',fila_enero+5,':E',ultima_fila+6,')/SUM(E',fila_enero_ant+5,':E',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(F',fila_enero+5,':F',ultima_fila+6,')/SUM(F',fila_enero_ant+5,':F',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    "",
                    paste('IFERROR(SUM(H',fila_enero+5,':H',ultima_fila+6,')/SUM(H',fila_enero_ant+5,':H',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(I',fila_enero+5,':I',ultima_fila+6,')/SUM(I',fila_enero_ant+5,':I',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(J',fila_enero+5,':J',ultima_fila+6,')/SUM(J',fila_enero_ant+5,':J',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(K',fila_enero+5,':K',ultima_fila+6,')/SUM(K',fila_enero_ant+5,':K',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(L',fila_enero+5,':L',ultima_fila+6,')/SUM(L',fila_enero_ant+5,':L',fila_enero_ant+4+mes,')*100-100,','"*")',sep = "")
  )

}

for (i in 36:46) {
  writeFormula(wb, sheet ="CI_Carne" , x = tasa_corrido[i-35] ,startCol = i, startRow = ultima_fila+6)
}

if (mes %in% c(3,6,9,12)){
  tasa_trimestre <- c(paste('IFERROR(SUM(B',ultima_fila+4,':B',ultima_fila+6,')/SUM(B',fila_anterior+3,':B',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(C',ultima_fila+4,':C',ultima_fila+6,')/SUM(C',fila_anterior+3,':C',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(D',ultima_fila+4,':D',ultima_fila+6,')/SUM(D',fila_anterior+3,':D',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(E',ultima_fila+4,':E',ultima_fila+6,')/SUM(E',fila_anterior+3,':E',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(F',ultima_fila+4,':F',ultima_fila+6,')/SUM(F',fila_anterior+3,':F',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      "",
                      paste('IFERROR(SUM(H',ultima_fila+4,':H',ultima_fila+6,')/SUM(H',fila_anterior+3,':H',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(I',ultima_fila+4,':I',ultima_fila+6,')/SUM(I',fila_anterior+3,':I',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(J',ultima_fila+4,':J',ultima_fila+6,')/SUM(J',fila_anterior+3,':J',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(K',ultima_fila+4,':K',ultima_fila+6,')/SUM(K',fila_anterior+3,':K',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(L',ultima_fila+4,':L',ultima_fila+6,')/SUM(L',fila_anterior+3,':L',fila_anterior+5,')*100-100,','"*")',sep = "")
                      )

  for (i in 48:58) {
    writeFormula(wb, sheet ="CI_Carne" , x = tasa_trimestre[i-47] ,startCol = i, startRow = ultima_fila+6)
  }

}else{
  cat("Este mes no se actualiza trimestre")
}



addStyle(wb, sheet = "CI_Carne",style=cbp,rows = (ultima_fila+6),cols = 1)
addStyle(wb, sheet = "CI_Carne",style=cbn,rows = (ultima_fila+6),cols = c(2:6,8:12))
addStyle(wb, sheet = "CI_Carne",style=rn4,rows = (ultima_fila+6),cols = 14:58)


# Leche ------------------------------------------------------

data <- read.xlsx(wb, sheet = "Leche", colNames = TRUE,startRow = 5)
ultima_fila=nrow(data)

data$Período <- as.Date(data$Período, origin = "1899-12-30")
data$Período=format(data$Período, "%Y-%m")
fila_enero=which(data$Período== paste0(anio,"-01"))
fila_enero_ant=which(data$Período== paste0((anio-1),"-01"))
fila_anterior=which(data$Período== paste0((anio-1),"-",mes_0[mes]))

Leche_polvo=f_Leche_polvo(directorio,mes,anio)
Prod_leche=Leche_polvo[[1]]
Expo_impo=Leche_polvo[[2]]
valor_fecha=as.integer(as.Date(paste0(1,"/",mes_0[mes],"/",anio), format = "%d/%m/%Y") - as.Date("1899-12-30"))
writeData(wb, sheet = "Leche", x = valor_fecha,colNames = FALSE,startCol = "A", startRow = (ultima_fila+6))
writeData(wb, sheet = "Leche", x = Prod_leche,colNames = FALSE,startCol = "B", startRow = (fila_enero_ant[1]+5))


  writeData(wb, sheet = "Leche", x = Expo_impo$Importacion,colNames = FALSE,startCol = "G", startRow = (fila_enero_ant[1]+5))
  writeData(wb, sheet = "Leche", x = Expo_impo[,2:3],colNames = FALSE,startCol = "I", startRow = (fila_enero_ant[1]+5))

if(sum(is.na(Prod_leche[,2]))>0){
for (i in 1:sum(is.na(Prod_leche[,2]))) {

writeFormula(wb, sheet ="Leche" , x = paste0("AVERAGE(C",(fila_enero[1]+5),":C",ultima_fila+6-i,")") ,startCol = "C", startRow = (ultima_fila+7-i))

}
}else{}
  if(sum(is.na(Prod_leche[,3]))>0){
for (i in 1:sum(is.na(Prod_leche[,3]))) {

  writeFormula(wb, sheet ="Leche" , x = paste0("AVERAGE(D",(fila_enero[1]+5),":D",ultima_fila+6-i,")") ,startCol = "D", startRow = (ultima_fila+7-i))

}}else{}

#Añadir formulas

writeFormula(wb, sheet ="Leche" , x = paste0("SUM(F",ultima_fila+6,":G",ultima_fila+6,")") ,startCol = "E", startRow = ultima_fila+6)
writeFormula(wb, sheet ="Leche" , x = paste0("SUM(I",ultima_fila+6,":J",ultima_fila+6,")") ,startCol = "H", startRow = ultima_fila+6)




tasa_anual <- c(paste('IFERROR(B',ultima_fila+6,'/B',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(C',ultima_fila+6,'/C',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('D',ultima_fila+6,'-D',ultima_fila+5,sep = ""),
                paste('IFERROR(E',ultima_fila+6,'/E',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(F',ultima_fila+6,'/F',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(G',ultima_fila+6,'/G',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(H',ultima_fila+6,'/H',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(I',ultima_fila+6,'/I',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(J',ultima_fila+6,'/J',fila_anterior+5,'*100-100,','"*")',sep = ""))


for (i in 12:20) {
  writeFormula(wb, sheet ="Leche" , x = tasa_anual[i-11] ,startCol = i, startRow = ultima_fila+6)
}

if(mes==1){
  tasa_corrido <- c(paste('IFERROR(SUM(B',ultima_fila+6,':B',ultima_fila+6,')/SUM(B',fila_enero_ant+5,':B',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(C',ultima_fila+6,':C',ultima_fila+6,')/SUM(C',fila_enero_ant+5,':C',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(D',ultima_fila+6,':D',ultima_fila+6,')/SUM(D',fila_enero_ant+5,':D',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(E',ultima_fila+6,':E',ultima_fila+6,')/SUM(E',fila_enero_ant+5,':E',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(F',ultima_fila+6,':F',ultima_fila+6,')/SUM(F',fila_enero_ant+5,':F',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(G',ultima_fila+6,':G',ultima_fila+6,')/SUM(G',fila_enero_ant+5,':G',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(H',ultima_fila+6,':H',ultima_fila+6,')/SUM(H',fila_enero_ant+5,':H',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(I',ultima_fila+6,':I',ultima_fila+6,')/SUM(I',fila_enero_ant+5,':I',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(J',ultima_fila+6,':J',ultima_fila+6,')/SUM(J',fila_enero_ant+5,':J',fila_enero_ant+4+mes,')*100-100,','"*")',sep = "")
  )

}else{
  tasa_corrido <- c(paste('IFERROR(SUM(B',fila_enero+5,':B',ultima_fila+6,')/SUM(B',fila_enero_ant+5,':B',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(C',fila_enero+5,':C',ultima_fila+6,')/SUM(C',fila_enero_ant+5,':C',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(D',fila_enero+5,':D',ultima_fila+6,')/SUM(D',fila_enero_ant+5,':D',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(E',fila_enero+5,':E',ultima_fila+6,')/SUM(E',fila_enero_ant+5,':E',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(F',fila_enero+5,':F',ultima_fila+6,')/SUM(F',fila_enero_ant+5,':F',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(G',fila_enero+5,':G',ultima_fila+6,')/SUM(G',fila_enero_ant+5,':G',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(H',fila_enero+5,':H',ultima_fila+6,')/SUM(H',fila_enero_ant+5,':H',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(I',fila_enero+5,':I',ultima_fila+6,')/SUM(I',fila_enero_ant+5,':I',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(J',fila_enero+5,':J',ultima_fila+6,')/SUM(J',fila_enero_ant+5,':J',fila_enero_ant+4+mes,')*100-100,','"*")',sep = "")
  )

}

for (i in 22:30) {
  writeFormula(wb, sheet ="Leche" , x = tasa_corrido[i-21] ,startCol = i, startRow = ultima_fila+6)
}

if (mes %in% c(3,6,9,12)){
  tasa_trimestre <- c(paste('IFERROR(SUM(B',ultima_fila+4,':B',ultima_fila+6,')/SUM(B',fila_anterior+3,':B',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(C',ultima_fila+4,':C',ultima_fila+6,')/SUM(C',fila_anterior+3,':C',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(D',ultima_fila+4,':D',ultima_fila+6,')/SUM(D',fila_anterior+3,':D',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(E',ultima_fila+4,':E',ultima_fila+6,')/SUM(E',fila_anterior+3,':E',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(F',ultima_fila+4,':F',ultima_fila+6,')/SUM(F',fila_anterior+3,':F',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(G',ultima_fila+4,':G',ultima_fila+6,')/SUM(G',fila_anterior+3,':G',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(H',ultima_fila+4,':H',ultima_fila+6,')/SUM(H',fila_anterior+3,':H',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(I',ultima_fila+4,':I',ultima_fila+6,')/SUM(I',fila_anterior+3,':I',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(J',ultima_fila+4,':J',ultima_fila+6,')/SUM(J',fila_anterior+3,':J',fila_anterior+5,')*100-100,','"*")',sep = "")
  )

  for (i in 32:40) {
    writeFormula(wb, sheet ="Leche" , x = tasa_trimestre[i-31] ,startCol = i, startRow = ultima_fila+6)
  }

}else{
  cat("Este mes no se actualiza trimestre")
}


addStyle(wb, sheet = "Leche",style=cbp,rows = (ultima_fila+6),cols = 1)
addStyle(wb, sheet = "Leche",style=cbn3,rows = (ultima_fila+6),cols = 2)
addStyle(wb, sheet = "Leche",style=cbn,rows = (ultima_fila+6),cols = 3:10)
addStyle(wb, sheet = "Leche",style=rn4,rows = (ultima_fila+6),cols = 12:40)



# Porcino_kilo_en_pie ------------------------------------------------------

data <- read.xlsx(wb, sheet = "Porcino kilo en pie", colNames = TRUE,startRow = 5)
ultima_fila=nrow(data)

data$Período <- as.Date(data$Período, origin = "1899-12-30")
data$Período=format(data$Período, "%Y-%m")
fila_enero=which(data$Período== paste0(anio,"-01"))
fila_enero_ant=which(data$Período== paste0((anio-1),"-01"))
fila_anterior=which(data$Período== paste0((anio-1),"-",mes_0[mes]))

valor_Porcino=f_Porcino(directorio,mes,anio)
valor_fecha=as.integer(as.Date(paste0(1,"/",mes_0[mes],"/",anio), format = "%d/%m/%Y") - as.Date("1899-12-30"))
writeData(wb, sheet = "Porcino kilo en pie", x = valor_fecha,colNames = FALSE,startCol = "A", startRow = (ultima_fila+6))

if(mes==1){
  writeData(wb, sheet = "Porcino kilo en pie", x = valor_Porcino,colNames = FALSE,startCol = "B", startRow = (ultima_fila+6))
}else{
  writeData(wb, sheet = "Porcino kilo en pie", x = valor_Porcino,colNames = FALSE,startCol = "B", startRow = (fila_enero[1]+5))
}

#Añadir formulas
participacion <- c(paste0("SUM(K",ultima_fila+6,":L",ultima_fila+6,")"),paste0("C",ultima_fila+6,"/B",ultima_fila+6,"*100"),
                   paste0("D",ultima_fila+6,"/B",ultima_fila+6,"*100")) ## skip header row

for (i in 10:12) {
  writeFormula(wb, sheet ="Porcino kilo en pie" , x = participacion[i-9] ,startCol = i, startRow = ultima_fila+6)
}

tasa_anual <- c(paste('IFERROR(B',ultima_fila+6,'/B',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(C',ultima_fila+6,'/C',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(D',ultima_fila+6,'/D',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(E',ultima_fila+6,'/E',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(F',ultima_fila+6,'/F',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(G',ultima_fila+6,'/G',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(H',ultima_fila+6,'/H',fila_anterior+5,'*100-100,','"*")',sep = ""))


for (i in 14:20) {
  writeFormula(wb, sheet ="Porcino kilo en pie" , x = tasa_anual[i-13] ,startCol = i, startRow = ultima_fila+6)
}

if(mes==1){
  tasa_corrido <- c(paste('IFERROR(SUM(B',ultima_fila+6,':B',ultima_fila+6,')/SUM(B',fila_enero_ant+5,':B',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(C',ultima_fila+6,':C',ultima_fila+6,')/SUM(C',fila_enero_ant+5,':C',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(D',ultima_fila+6,':D',ultima_fila+6,')/SUM(D',fila_enero_ant+5,':D',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(E',ultima_fila+6,':E',ultima_fila+6,')/SUM(E',fila_enero_ant+5,':E',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(F',ultima_fila+6,':F',ultima_fila+6,')/SUM(F',fila_enero_ant+5,':F',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(G',ultima_fila+6,':G',ultima_fila+6,')/SUM(G',fila_enero_ant+5,':G',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(H',ultima_fila+6,':H',ultima_fila+6,')/SUM(H',fila_enero_ant+5,':H',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""))


}else{
  tasa_corrido <- c(paste('IFERROR(SUM(B',fila_enero+5,':B',ultima_fila+6,')/SUM(B',fila_enero_ant+5,':B',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(C',fila_enero+5,':C',ultima_fila+6,')/SUM(C',fila_enero_ant+5,':C',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(D',fila_enero+5,':D',ultima_fila+6,')/SUM(D',fila_enero_ant+5,':D',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(E',fila_enero+5,':E',ultima_fila+6,')/SUM(E',fila_enero_ant+5,':E',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(F',fila_enero+5,':F',ultima_fila+6,')/SUM(F',fila_enero_ant+5,':F',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(G',fila_enero+5,':G',ultima_fila+6,')/SUM(G',fila_enero_ant+5,':G',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(H',fila_enero+5,':H',ultima_fila+6,')/SUM(H',fila_enero_ant+5,':H',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""))


}
for (i in 22:28) {
  writeFormula(wb, sheet ="Porcino kilo en pie" , x = tasa_corrido[i-21] ,startCol = i, startRow = ultima_fila+6)
}

if (mes %in% c(3,6,9,12)){
  tasa_trimestre <- c(paste('IFERROR(SUM(B',ultima_fila+4,':B',ultima_fila+6,')/SUM(B',fila_anterior+4,':B',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(C',ultima_fila+4,':C',ultima_fila+6,')/SUM(C',fila_anterior+4,':C',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(D',ultima_fila+4,':D',ultima_fila+6,')/SUM(D',fila_anterior+4,':D',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(E',ultima_fila+4,':E',ultima_fila+6,')/SUM(E',fila_anterior+4,':E',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(F',ultima_fila+4,':F',ultima_fila+6,')/SUM(F',fila_anterior+4,':F',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(G',ultima_fila+4,':G',ultima_fila+6,')/SUM(G',fila_anterior+4,':G',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(H',ultima_fila+4,':H',ultima_fila+6,')/SUM(H',fila_anterior+4,':H',fila_anterior+5,')*100-100,','"*")',sep = ""))

  for (i in 30:36) {
    writeFormula(wb, sheet ="Porcino kilo en pie" , x = tasa_trimestre[i-29] ,startCol = i, startRow = ultima_fila+6)
  }

}else{
  cat("Este mes no se actualiza trimestre")
}

addStyle(wb, sheet = "Porcino kilo en pie",style=cbp,rows = (ultima_fila+6),cols = 1)
addStyle(wb, sheet = "Porcino kilo en pie",style=cbn,rows = (ultima_fila+6),cols = 2:9)
addStyle(wb, sheet = "Porcino kilo en pie",style=rn4,rows = (ultima_fila+6),cols = 10:12)
addStyle(wb, sheet = "Porcino kilo en pie",style=cbn2,rows = (ultima_fila+6),cols = 14:20)
addStyle(wb, sheet = "Porcino kilo en pie",style=rn4,rows = (ultima_fila+6),cols = 22:28)
addStyle(wb, sheet = "Porcino kilo en pie",style=rn4,rows = (ultima_fila+6),cols = 30:36)


# Porcino_cabezas ------------------------------------------------------

data <- read.xlsx(wb, sheet = "Porcino en cabezas" , colNames = TRUE,startRow = 5)
ultima_fila=nrow(data)

data$Período <- as.Date(data$Período, origin = "1899-12-30")
data$Período=format(data$Período, "%Y-%m")
fila_enero=which(data$Período== paste0(anio,"-01"))
fila_enero_ant=which(data$Período== paste0((anio-1),"-01"))
fila_anterior=which(data$Período== paste0((anio-1),"-",mes_0[mes]))

valor_Porcino=f_Porcino_cabezas(directorio,mes,anio)
valor_fecha=as.integer(as.Date(paste0(1,"/",mes_0[mes],"/",anio), format = "%d/%m/%Y") - as.Date("1899-12-30"))
writeData(wb, sheet = "Porcino en cabezas" , x = valor_fecha,colNames = FALSE,startCol = "A", startRow = (ultima_fila+6))
if(mes==1){
  writeData(wb, sheet = "Porcino en cabezas" , x = valor_Porcino,colNames = FALSE,startCol = "B", startRow = (ultima_fila+6))
}else{
  writeData(wb, sheet = "Porcino en cabezas" , x = valor_Porcino,colNames = FALSE,startCol = "B", startRow = (fila_enero[1]+5))

}

#Añadir formulas
participacion <- c(paste0("IFERROR(SUM(H",ultima_fila+6,":I",ultima_fila+6,"),0)"),
                   paste0("IFERROR(C",ultima_fila+6,"/B",ultima_fila+6,"*100,0)"),
                   paste0("IFERROR(D",ultima_fila+6,"/B",ultima_fila+6,"*100,0)")) ## skip header row

for (i in 7:9) {
  writeFormula(wb, sheet ="Porcino en cabezas"  , x = participacion[i-6] ,startCol = i, startRow = ultima_fila+6)
}

tasa_anual <- c(paste('IFERROR(B',ultima_fila+6,'/B',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(C',ultima_fila+6,'/C',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(D',ultima_fila+6,'/D',fila_anterior+5,'*100-100,','"*")',sep = ""))


for (i in 11:13) {
  writeFormula(wb, sheet ="Porcino en cabezas"  , x = tasa_anual[i-10] ,startCol = i, startRow = ultima_fila+6)
}

if(mes==1){
  tasa_corrido <- c(paste('IFERROR(SUM(B',ultima_fila+6,':B',ultima_fila+6,')/SUM(B',fila_enero_ant+5,':B',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(C',ultima_fila+6,':C',ultima_fila+6,')/SUM(C',fila_enero_ant+5,':C',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(D',ultima_fila+6,':D',ultima_fila+6,')/SUM(D',fila_enero_ant+5,':D',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(E',ultima_fila+6,':E',ultima_fila+6,')/SUM(E',fila_enero_ant+5,':E',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""))

}else{
  tasa_corrido <- c(paste('IFERROR(SUM(B',fila_enero+5,':B',ultima_fila+6,')/SUM(B',fila_enero_ant+5,':B',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(C',fila_enero+5,':C',ultima_fila+6,')/SUM(C',fila_enero_ant+5,':C',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(D',fila_enero+5,':D',ultima_fila+6,')/SUM(D',fila_enero_ant+5,':D',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(E',fila_enero+5,':E',ultima_fila+6,')/SUM(E',fila_enero_ant+5,':E',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""))

}

for (i in 16:19) {
  writeFormula(wb, sheet ="Porcino en cabezas"  , x = tasa_corrido[i-15] ,startCol = i, startRow = ultima_fila+6)
}

if (mes %in% c(3,6,9,12)){
  tasa_trimestre <- c(paste('IFERROR(SUM(B',ultima_fila+4,':B',ultima_fila+6,')/SUM(B',fila_anterior+3,':B',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(C',ultima_fila+4,':C',ultima_fila+6,')/SUM(C',fila_anterior+3,':C',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(D',ultima_fila+4,':D',ultima_fila+6,')/SUM(D',fila_anterior+3,':D',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(E',ultima_fila+4,':E',ultima_fila+6,')/SUM(E',fila_anterior+3,':E',fila_anterior+5,')*100-100,','"*")',sep = ""))

  for (i in 21:24) {
    writeFormula(wb, sheet ="Porcino en cabezas"  , x = tasa_trimestre[i-20] ,startCol = i, startRow = ultima_fila+6)
  }

}else{
  cat("Este mes no se actualiza trimestre")
}

addStyle(wb, sheet = "Porcino en cabezas" ,style=cbp,rows = (ultima_fila+6),cols = 1)
addStyle(wb, sheet = "Porcino en cabezas" ,style=cbn,rows = (ultima_fila+6),cols = 2:5)
addStyle(wb, sheet = "Porcino en cabezas" ,style=rn4,rows = (ultima_fila+6),cols = 7:9)
addStyle(wb, sheet = "Porcino en cabezas" ,style=rn4,rows = (ultima_fila+6),cols = 11:13)
addStyle(wb, sheet = "Porcino en cabezas" ,style=rn4,rows = (ultima_fila+6),cols = 16:19)
addStyle(wb, sheet = "Porcino en cabezas" ,style=cbn2,rows = (ultima_fila+6),cols = 21:24)



# Pollo_Huevo ------------------------------------------------------

data <- read.xlsx(wb, sheet = "Pollo_Huevo", colNames = TRUE,startRow = 5)
ultima_fila=nrow(data)

data$Período <- as.Date(data$Período, origin = "1899-12-30")
data$Período=format(data$Período, "%Y-%m")
fila_enero=which(data$Período== paste0(anio,"-01"))
fila_enero_ant=which(data$Período== paste0((anio-1),"-01"))
fila_anterior=which(data$Período== paste0((anio-1),"-",mes_0[mes]))

Valor_Huevos=f_Huevos(directorio,mes,anio)
Valor_Pollos=f_Pollos(directorio,mes,anio)
Valor_Fenavi=f_Fenavi(directorio,mes,anio)
valor_fecha=as.integer(as.Date(paste0(1,"/",mes_0[mes],"/",anio), format = "%d/%m/%Y") - as.Date("1899-12-30"))
writeData(wb, sheet = "Pollo_Huevo", x = valor_fecha,colNames = FALSE,startCol = "A", startRow = (ultima_fila+6))
writeData(wb, sheet = "Pollo_Huevo", x = Valor_Huevos,colNames = FALSE,startCol = "B", startRow = (fila_enero_ant[1]+5))
writeData(wb, sheet = "Pollo_Huevo", x = Valor_Pollos,colNames = FALSE,startCol = "C", startRow = (fila_enero_ant[1]+5))
writeData(wb, sheet = "Pollo_Huevo", x = Valor_Fenavi,colNames = FALSE,startCol = "D", startRow = (fila_enero_ant[1]+5))


#Añadir formulas
tasa_mensual <-    c(paste('IFERROR(B',ultima_fila+6,'/B',ultima_fila+5,'*100-100,','"*")',sep = ""),
                     paste('IFERROR(C',ultima_fila+6,'/C',ultima_fila+5,'*100-100,','"*")',sep = ""),
                     paste('IFERROR(D',ultima_fila+6,'/D',ultima_fila+5,'*100-100,','"*")',sep = ""),
                     paste('IFERROR(E',ultima_fila+6,'/E',ultima_fila+5,'*100-100,','"*")',sep = ""),
                     paste('IFERROR(F',ultima_fila+6,'/F',ultima_fila+5,'*100-100,','"*")',sep = ""),
                     paste('IFERROR(G',ultima_fila+6,'/G',ultima_fila+5,'*100-100,','"*")',sep = ""),
                     paste('IFERROR(H',ultima_fila+6,'/H',ultima_fila+5,'*100-100,','"*")',sep = ""),
                     paste('IFERROR(I',ultima_fila+6,'/I',ultima_fila+5,'*100-100,','"*")',sep = ""),
                     paste('IFERROR(L',ultima_fila+6,'/L',ultima_fila+5,'*100-100,','"*")',sep = ""))

for (i in 16:24) {
  writeFormula(wb, sheet ="Pollo_Huevo" , x = tasa_mensual[i-15] ,startCol = i, startRow = ultima_fila+6)
}

tasa_anual <- c(paste('IFERROR(B',ultima_fila+6,'/B',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(C',ultima_fila+6,'/C',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(D',ultima_fila+6,'/D',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(E',ultima_fila+6,'/E',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(F',ultima_fila+6,'/F',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(G',ultima_fila+6,'/G',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(H',ultima_fila+6,'/H',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(I',ultima_fila+6,'/I',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(J',ultima_fila+6,'/J',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(K',ultima_fila+6,'/K',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(L',ultima_fila+6,'/L',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(M',ultima_fila+6,'/M',fila_anterior+5,'*100-100,','"*")',sep = ""),
                paste('IFERROR(N',ultima_fila+6,'/N',fila_anterior+5,'*100-100,','"*")',sep = ""))


for (i in 26:38) {
  writeFormula(wb, sheet ="Pollo_Huevo" , x = tasa_anual[i-25] ,startCol = i, startRow = ultima_fila+6)
}

if(mes==1){
  tasa_corrido <- c(paste('IFERROR(SUM(B',ultima_fila+6,':B',ultima_fila+6,')/SUM(B',fila_enero_ant+5,':B',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(C',ultima_fila+6,':C',ultima_fila+6,')/SUM(C',fila_enero_ant+5,':C',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(D',ultima_fila+6,':D',ultima_fila+6,')/SUM(D',fila_enero_ant+5,':D',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(E',ultima_fila+6,':E',ultima_fila+6,')/SUM(E',fila_enero_ant+5,':E',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(F',ultima_fila+6,':F',ultima_fila+6,')/SUM(F',fila_enero_ant+5,':F',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(G',ultima_fila+6,':G',ultima_fila+6,')/SUM(G',fila_enero_ant+5,':G',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(H',ultima_fila+6,':H',ultima_fila+6,')/SUM(H',fila_enero_ant+5,':H',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(I',ultima_fila+6,':I',ultima_fila+6,')/SUM(I',fila_enero_ant+5,':I',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(J',ultima_fila+6,':J',ultima_fila+6,')/SUM(J',fila_enero_ant+5,':J',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(K',ultima_fila+6,':K',ultima_fila+6,')/SUM(K',fila_enero_ant+5,':K',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(L',ultima_fila+6,':L',ultima_fila+6,')/SUM(L',fila_enero_ant+5,':L',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(M',ultima_fila+6,':M',ultima_fila+6,')/SUM(M',fila_enero_ant+5,':M',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(N',ultima_fila+6,':N',ultima_fila+6,')/SUM(N',fila_enero_ant+5,':N',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""))


}else{
  tasa_corrido <- c(paste('IFERROR(SUM(B',fila_enero+5,':B',ultima_fila+6,')/SUM(B',fila_enero_ant+5,':B',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(C',fila_enero+5,':C',ultima_fila+6,')/SUM(C',fila_enero_ant+5,':C',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(D',fila_enero+5,':D',ultima_fila+6,')/SUM(D',fila_enero_ant+5,':D',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(E',fila_enero+5,':E',ultima_fila+6,')/SUM(E',fila_enero_ant+5,':E',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(F',fila_enero+5,':F',ultima_fila+6,')/SUM(F',fila_enero_ant+5,':F',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(G',fila_enero+5,':G',ultima_fila+6,')/SUM(G',fila_enero_ant+5,':G',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(H',fila_enero+5,':H',ultima_fila+6,')/SUM(H',fila_enero_ant+5,':H',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(I',fila_enero+5,':I',ultima_fila+6,')/SUM(I',fila_enero_ant+5,':I',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(J',fila_enero+5,':J',ultima_fila+6,')/SUM(J',fila_enero_ant+5,':J',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(K',fila_enero+5,':K',ultima_fila+6,')/SUM(K',fila_enero_ant+5,':K',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(L',fila_enero+5,':L',ultima_fila+6,')/SUM(L',fila_enero_ant+5,':L',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(M',fila_enero+5,':M',ultima_fila+6,')/SUM(M',fila_enero_ant+5,':M',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""),
                    paste('IFERROR(SUM(N',fila_enero+5,':N',ultima_fila+6,')/SUM(N',fila_enero_ant+5,':N',fila_enero_ant+4+mes,')*100-100,','"*")',sep = ""))


}

for (i in 40:52) {
  writeFormula(wb, sheet ="Pollo_Huevo" , x = tasa_corrido[i-39] ,startCol = i, startRow = ultima_fila+6)
}

if (mes %in% c(3,6,9,12)){
  tasa_trimestre <- c(paste('IFERROR(SUM(B',ultima_fila+4,':B',ultima_fila+6,')/SUM(B',fila_anterior+3,':B',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(C',ultima_fila+4,':C',ultima_fila+6,')/SUM(C',fila_anterior+3,':C',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(D',ultima_fila+4,':D',ultima_fila+6,')/SUM(D',fila_anterior+3,':D',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(E',ultima_fila+4,':E',ultima_fila+6,')/SUM(E',fila_anterior+3,':E',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(F',ultima_fila+4,':F',ultima_fila+6,')/SUM(F',fila_anterior+3,':F',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(G',ultima_fila+4,':G',ultima_fila+6,')/SUM(G',fila_anterior+3,':G',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(H',ultima_fila+4,':H',ultima_fila+6,')/SUM(H',fila_anterior+3,':H',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(I',ultima_fila+4,':I',ultima_fila+6,')/SUM(I',fila_anterior+3,':I',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(J',ultima_fila+4,':J',ultima_fila+6,')/SUM(J',fila_anterior+3,':J',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(K',ultima_fila+4,':K',ultima_fila+6,')/SUM(K',fila_anterior+3,':K',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(L',ultima_fila+4,':L',ultima_fila+6,')/SUM(L',fila_anterior+3,':L',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(M',ultima_fila+4,':M',ultima_fila+6,')/SUM(M',fila_anterior+3,':M',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(N',ultima_fila+4,':N',ultima_fila+6,')/SUM(N',fila_anterior+3,':N',fila_anterior+5,')*100-100,','"*")',sep = ""))

  for (i in 54:66) {
    writeFormula(wb, sheet ="Pollo_Huevo" , x = tasa_trimestre[i-53] ,startCol = i, startRow = ultima_fila+6)
  }

}else{
  cat("Este mes no se actualiza trimestre")
}


if (mes==12){
  tasa_anual_trimestre <- c(paste('IFERROR(SUM(B',ultima_fila-5,':B',ultima_fila+6,')/SUM(B',fila_anterior-6,':B',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(C',ultima_fila-5,':C',ultima_fila+6,')/SUM(C',fila_anterior-6,':C',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(D',ultima_fila-5,':D',ultima_fila+6,')/SUM(D',fila_anterior-6,':D',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(E',ultima_fila-5,':E',ultima_fila+6,')/SUM(E',fila_anterior-6,':E',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(F',ultima_fila-5,':F',ultima_fila+6,')/SUM(F',fila_anterior-6,':F',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(G',ultima_fila-5,':G',ultima_fila+6,')/SUM(G',fila_anterior-6,':G',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(H',ultima_fila-5,':H',ultima_fila+6,')/SUM(H',fila_anterior-6,':H',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(I',ultima_fila-5,':I',ultima_fila+6,')/SUM(I',fila_anterior-6,':I',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(J',ultima_fila-5,':J',ultima_fila+6,')/SUM(J',fila_anterior-6,':J',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(K',ultima_fila-5,':K',ultima_fila+6,')/SUM(K',fila_anterior-6,':K',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(L',ultima_fila-5,':L',ultima_fila+6,')/SUM(L',fila_anterior-6,':L',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(M',ultima_fila-5,':M',ultima_fila+6,')/SUM(M',fila_anterior-6,':M',fila_anterior+5,')*100-100,','"*")',sep = ""),
                      paste('IFERROR(SUM(N',ultima_fila-5,':N',ultima_fila+6,')/SUM(N',fila_anterior-6,':N',fila_anterior+5,')*100-100,','"*")',sep = ""))

  for (i in 68:80) {
    writeFormula(wb, sheet ="Pollo_Huevo" , x = tasa_anual_trimestre[i-67] ,startCol = i, startRow = ultima_fila+6)
  }

}else{
  cat("Este mes no se actualiza anual")
}


addStyle(wb, sheet = "Pollo_Huevo",style=cbp,rows = (ultima_fila+6),cols = 1)
addStyle(wb, sheet = "Pollo_Huevo",style=cbn,rows = (ultima_fila+6),cols = 2:14)
addStyle(wb, sheet = "Pollo_Huevo",style=rn4,rows = (ultima_fila+6),cols = 16:24)
addStyle(wb, sheet = "Pollo_Huevo",style=rn4,rows = (ultima_fila+6),cols = 26:38)
addStyle(wb, sheet = "Pollo_Huevo",style=rn4,rows = (ultima_fila+6),cols = 40:52)
addStyle(wb, sheet = "Pollo_Huevo",style=rn4,rows = (ultima_fila+6),cols = 54:66)
addStyle(wb, sheet = "Pollo_Huevo",style=rn4,rows = (ultima_fila+6),cols = 68:80)


# Cuadro Bovino -----------------------------------------------------------

trim_rom=f_trim_rom(mes)
writeData(wb, sheet = "CUADROS BOVINO", x = paste0(trim_rom," Trimestre ","Variación Anual "),colNames = FALSE,startCol = "H", startRow = 5)
writeData(wb, sheet = "CUADROS BOVINO", x = paste0(nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "D", startRow = 6)
writeData(wb, sheet = "CUADROS BOVINO", x = paste0(nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "E", startRow = 6)
writeData(wb, sheet = "CUADROS BOVINO", x = paste0(anio-1,"-",trim_rom," / ",anio-2,"-",trim_rom),colNames = FALSE,startCol = "H", startRow = 6)
writeData(wb, sheet = "CUADROS BOVINO", x = paste0(anio,"-",trim_rom," / ",anio-1,"-",trim_rom),colNames = FALSE,startCol = "I", startRow = 6)
writeData(wb, sheet = "CUADROS BOVINO", x = paste0(anio-1,"-",trim_rom," / ",anio-2,"-",trim_rom),colNames = FALSE,startCol = "H", startRow = 13)
writeData(wb, sheet = "CUADROS BOVINO", x = paste0(anio,"-",trim_rom," / ",anio-1,"-",trim_rom),colNames = FALSE,startCol = "I", startRow = 13)
writeData(wb, sheet = "CUADROS BOVINO", x = paste0(anio-1,"-",trim_rom," / ",anio-2,"-",trim_rom),colNames = FALSE,startCol = "H", startRow = 18)
writeData(wb, sheet = "CUADROS BOVINO", x = paste0(anio,"-",trim_rom," / ",anio-1,"-",trim_rom),colNames = FALSE,startCol = "I", startRow = 18)
writeData(wb, sheet = "CUADROS BOVINO", x = paste0(" Variación Anual ",nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "D", startRow = 18)
writeData(wb, sheet = "CUADROS BOVINO", x = paste0(" Variación Anual ",nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "E", startRow = 18)
writeData(wb, sheet = "CUADROS BOVINO", x = paste0("Variación Año corrido ",nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "F", startRow = 18)
writeData(wb, sheet = "CUADROS BOVINO", x = paste0("Variación Año corrido ",nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "G", startRow = 18)
writeData(wb, sheet = "CUADROS BOVINO", x = paste0(anio-1,"-",trim_rom," / ",anio-2,"-",trim_rom),colNames = FALSE,startCol = "H", startRow = 24)
writeData(wb, sheet = "CUADROS BOVINO", x = paste0(anio,"-",trim_rom," / ",anio-1,"-",trim_rom),colNames = FALSE,startCol = "I", startRow = 24)
writeData(wb, sheet = "CUADROS BOVINO", x = paste0(" Variación Anual ",nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "D", startRow = 24)
writeData(wb, sheet = "CUADROS BOVINO", x = paste0(" Variación Anual ",nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "E", startRow = 24)
writeData(wb, sheet = "CUADROS BOVINO", x = paste0("Variación Año corrido ",nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "F", startRow = 24)
writeData(wb, sheet = "CUADROS BOVINO", x = paste0("Variación Año corrido ",nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "G", startRow = 24)

Valores_totales<-matrix(c(paste0("'Bovino kilo en pie'!S",fila_anterior+6),
                          paste0("'Bovino kilo en pie'!T",fila_anterior+6),
                          paste0("'Bovino kilo en pie'!U",fila_anterior+6),
                          paste0("'Bovino kilo en pie'!V",fila_anterior+6),
                          paste0("'Bovino kilo en pie'!W",fila_anterior+6),
                          paste0("'Bovino cabezas'!R",fila_anterior+5),
                          paste0("'Bovino kilo en pie'!S",ultima_fila+7),
                          paste0("'Bovino kilo en pie'!T",ultima_fila+7),
                          paste0("'Bovino kilo en pie'!U",ultima_fila+7),
                          paste0("'Bovino kilo en pie'!V",ultima_fila+7),
                          paste0("'Bovino kilo en pie'!W",ultima_fila+7),
                          paste0("'Bovino cabezas'!R",ultima_fila+6),
                          "SUM(F8:F11)",
                          paste0("'Bovino kilo en pie'!N",fila_anterior+6),
                          paste0("'Bovino kilo en pie'!O",fila_anterior+6),
                          paste0("'Bovino kilo en pie'!P",fila_anterior+6),
                          paste0("'Bovino kilo en pie'!Q",fila_anterior+6),
                          "",
                          "SUM(G8:G11)",
                          "E8*F8/100",
                          "E9*F9/100",
                          "E10*F10/100",
                          "E11*F11/100",
                          "",
                          paste0("'Bovino kilo en pie'!AN",fila_anterior+6),
                          paste0("'Bovino kilo en pie'!AO",fila_anterior+6),
                          paste0("'Bovino kilo en pie'!AP",fila_anterior+6),
                          paste0("'Bovino kilo en pie'!AQ",fila_anterior+6),
                          paste0("'Bovino kilo en pie'!AR",fila_anterior+6),
                          paste0("'Bovino cabezas'!AL",fila_anterior+5),
                          paste0("'Bovino kilo en pie'!AN",ultima_fila+7),
                          paste0("'Bovino kilo en pie'!AO",ultima_fila+7),
                          paste0("'Bovino kilo en pie'!AP",ultima_fila+7),
                          paste0("'Bovino kilo en pie'!AQ",ultima_fila+7),
                          paste0("'Bovino kilo en pie'!AR",ultima_fila+7),
                          paste0("'Bovino cabezas'!AL",ultima_fila+6)
                          ),nrow=6,ncol=6,byrow = FALSE)

for (i in 1:6) {
for (j in 1:6) {
  writeFormula(wb, sheet ="CUADROS BOVINO" , x = Valores_totales[i,j] ,startCol = j+3, startRow = i+6)
}
}

Valores_consumo<-matrix(c(paste0("'CI_Carne'!X",fila_anterior+5),
                          paste0("'CI_Carne'!Y",fila_anterior+5),
                          paste0("'CI_Carne'!Z",fila_anterior+5),
                          paste0("'CI_Carne'!AA",fila_anterior+5),
                          paste0("'CI_Carne'!X",ultima_fila+6),
                          paste0("'CI_Carne'!Y",ultima_fila+6),
                          paste0("'CI_Carne'!Z",ultima_fila+6),
                          paste0("'CI_Carne'!AA",ultima_fila+6),
                          "SUM(F15:F17)",
                          paste0("'CI_Carne'!O",fila_anterior+5),
                          paste0("'CI_Carne'!P",fila_anterior+5),
                          paste0("'CI_Carne'!Q",fila_anterior+5),
                          "SUM(G15:G17)",
                          "E15*F15/100",
                          "E16*F16/100",
                          "E17*F17/100",
                          paste0("'CI_Carne'!AV",fila_anterior+5),
                          paste0("'CI_Carne'!AW",fila_anterior+5),
                          paste0("'CI_Carne'!AY",fila_anterior+5),
                          paste0("'CI_Carne'!AZ",fila_anterior+5),
                          paste0("'CI_Carne'!AV",ultima_fila+6),
                          paste0("'CI_Carne'!AW",ultima_fila+6),
                          paste0("'CI_Carne'!AY",ultima_fila+6),
                          paste0("'CI_Carne'!AZ",ultima_fila+6)
),nrow=4,ncol=6,byrow = FALSE)

for (i in 1:4) {
  for (j in 1:6) {
    writeFormula(wb, sheet ="CUADROS BOVINO" , x = Valores_consumo[i,j] ,startCol = j+3, startRow = i+13)
  }
}



Valores_precios<-matrix(c(paste0("'Precios'!AB",fila_anterior_precios+5),
                          paste0("'Precios'!AC",fila_anterior_precios+5),
                          paste0("'Precios'!AB",ultima_fila_precios+6),
                          paste0("'Precios'!AC",ultima_fila_precios+6),
                          paste0("'Precios'!AR",fila_anterior_precios+5),
                          paste0("'Precios'!AS",fila_anterior_precios+5),
                          paste0("'Precios'!AR",ultima_fila_precios+6),
                          paste0("'Precios'!AS",ultima_fila_precios+6),
                          paste0("'Precios'!BH",fila_anterior_precios+5),
                          paste0("'Precios'!BI",fila_anterior_precios+5),
                          paste0("'Precios'!BH",ultima_fila_precios+6),
                          paste0("'Precios'!BI",ultima_fila_precios+6)

),nrow=2,ncol=6,byrow = FALSE)

for (i in 1:2) {
  for (j in 1:6) {
    writeFormula(wb, sheet ="CUADROS BOVINO" , x = Valores_precios[i,j] ,startCol = j+3, startRow = i+18)
  }
}


Valores_exportaciones<-matrix(c(paste0("'Bovino kilo en pie'!X",fila_anterior+6),
                                paste0("'Bovino kilo en pie'!Z",fila_anterior+6),
                                paste0("'Bovino kilo en pie'!X",ultima_fila+7),
                                paste0("'Bovino kilo en pie'!Z",ultima_fila+7),
                                paste0("'Bovino kilo en pie'!AH",fila_anterior+6),
                                paste0("'Bovino kilo en pie'!AJ",fila_anterior+6),
                                paste0("'Bovino kilo en pie'!AH",ultima_fila+7),
                                paste0("'Bovino kilo en pie'!AJ",ultima_fila+7),
                                paste0("'Bovino kilo en pie'!AS",fila_anterior+6),
                                paste0("'Bovino kilo en pie'!AU",fila_anterior+6),
                                paste0("'Bovino kilo en pie'!AS",ultima_fila+7),
                                paste0("'Bovino kilo en pie'!AU",ultima_fila+7)

),nrow=2,ncol=6,byrow = FALSE)

for (i in 1:2) {
  for (j in 1:6) {
    writeFormula(wb, sheet ="CUADROS BOVINO" , x = Valores_exportaciones[i,j] ,startCol = j+3, startRow = i+24)
  }
}

writeFormula(wb, sheet ="CUADROS BOVINO" , x = paste0("'Bovino kilo en pie'!AA",fila_anterior+6) ,startCol = "D", startRow =29 )
writeFormula(wb, sheet ="CUADROS BOVINO" , x = paste0("'Bovino kilo en pie'!AA",ultima_fila+7) ,startCol = "E", startRow =29 )
writeFormula(wb, sheet ="CUADROS BOVINO" , x = paste0("'Bovino kilo en pie'!AK",fila_anterior+6) ,startCol = "F", startRow =29 )
writeFormula(wb, sheet ="CUADROS BOVINO" , x = paste0("'Bovino kilo en pie'!AK",ultima_fila+7) ,startCol = "G", startRow =29 )
writeFormula(wb, sheet ="CUADROS BOVINO" , x = paste0("'Bovino kilo en pie'!AV",fila_anterior+6) ,startCol = "H", startRow =29 )
writeFormula(wb, sheet ="CUADROS BOVINO" , x = paste0("'Bovino kilo en pie'!AV",ultima_fila+7) ,startCol = "I", startRow =29 )

setRowHeights(wb,sheet ="CUADROS BOVINO",rows = c(21,22,28),heights = 0)



# Cuadro LECHE -----------------------------------------------------------

writeData(wb, sheet = "CUADROS LECHE", x = paste0(trim_rom," Trimestre ","Variación Anual "),colNames = FALSE,startCol = "F", startRow = 5)
writeData(wb, sheet = "CUADROS LECHE", x = paste0(nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "D", startRow = 6)
writeData(wb, sheet = "CUADROS LECHE", x = paste0(nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "E", startRow = 6)
writeData(wb, sheet = "CUADROS LECHE", x = paste0(anio-1,"-",trim_rom," / ",anio-2,"-",trim_rom),colNames = FALSE,startCol = "F", startRow = 6)
writeData(wb, sheet = "CUADROS LECHE", x = paste0(anio,"-",trim_rom," / ",anio-1,"-",trim_rom),colNames = FALSE,startCol = "G", startRow = 6)


Valores_variaciones<-matrix(c(paste0("'Leche'!L",fila_anterior+5),
                              paste0("'Leche'!M",fila_anterior+5),
                              paste0("'Leche'!N",fila_anterior+5),
                              paste0("'Leche'!L",ultima_fila+6),
                              paste0("'Leche'!M",ultima_fila+6),
                              paste0("'Leche'!N",ultima_fila+6),
                              paste0("'Leche'!AF",fila_anterior+5),
                              paste0("'Leche'!AG",fila_anterior+5),
                              paste0("'Leche'!AH",fila_anterior+5),
                              paste0("'Leche'!AF",ultima_fila+6),
                              paste0("'Leche'!AG",ultima_fila+6),
                              paste0("'Leche'!AH",ultima_fila+6)

),nrow=3,ncol=4,byrow = FALSE)

for (i in 1:3) {
  for (j in 1:4) {
    writeFormula(wb, sheet ="CUADROS LECHE" , x = Valores_variaciones[i,j] ,startCol = j+3, startRow = i+6)
  }
}




Valores_precios<-matrix(c(paste0("'Precios'!AJ",fila_anterior_precios+5),
                          paste0("'Precios'!AI",fila_anterior_precios+5),
                          paste0("'Precios'!AJ",ultima_fila_precios+6),
                          paste0("'Precios'!AI",ultima_fila_precios+6),
                          paste0("'Precios'!BP",fila_anterior_precios+5),
                          paste0("'Precios'!BO",fila_anterior_precios+5),
                          paste0("'Precios'!BP",ultima_fila_precios+6),
                          paste0("'Precios'!BO",ultima_fila_precios+6)

),nrow=2,ncol=4,byrow = FALSE)

for (i in 1:2) {
  for (j in 1:4) {
    writeFormula(wb, sheet ="CUADROS LECHE" , x = Valores_precios[i,j] ,startCol = j+3, startRow = i+10)
  }
}



#IMPORTACIONES
writeFormula(wb, sheet ="CUADROS LECHE" , x = paste0("'Leche'!O",fila_anterior+5) ,startCol = "D", startRow =14 )
writeFormula(wb, sheet ="CUADROS LECHE" , x = paste0("'Leche'!O",ultima_fila+6) ,startCol = "E", startRow =14 )
writeFormula(wb, sheet ="CUADROS LECHE" , x = paste0("'Leche'!AK",fila_anterior+5) ,startCol = "F", startRow =14 )
writeFormula(wb, sheet ="CUADROS LECHE" , x = paste0("'Leche'!AK",ultima_fila+6) ,startCol = "G", startRow =14 )




#EXPORTACIONES
writeFormula(wb, sheet ="CUADROS LECHE" , x = paste0("'Leche'!R",fila_anterior+5) ,startCol = "D", startRow =16 )
writeFormula(wb, sheet ="CUADROS LECHE" , x = paste0("'Leche'!R",ultima_fila+6) ,startCol = "E", startRow =16 )
writeFormula(wb, sheet ="CUADROS LECHE" , x = paste0("'Leche'!AL",fila_anterior+5) ,startCol = "F", startRow =16 )
writeFormula(wb, sheet ="CUADROS LECHE" , x = paste0("'Leche'!AL",ultima_fila+6) ,startCol = "G", startRow =16 )



# Cuadro PORCINO -----------------------------------------------------------

writeData(wb, sheet = "CUADROS PORCINO", x = paste0(trim_rom," Trimestre ","Variación Anual "),colNames = FALSE,startCol = "H", startRow = 5)
writeData(wb, sheet = "CUADROS PORCINO", x = paste0(nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "D", startRow = 6)
writeData(wb, sheet = "CUADROS PORCINO", x = paste0(nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "E", startRow = 6)
writeData(wb, sheet = "CUADROS PORCINO", x = paste0(nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "D", startRow = 12)
writeData(wb, sheet = "CUADROS PORCINO", x = paste0(nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "E", startRow = 12)
writeData(wb, sheet = "CUADROS PORCINO", x = paste0(nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "D", startRow = 17)
writeData(wb, sheet = "CUADROS PORCINO", x = paste0(nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "E", startRow = 17)
writeData(wb, sheet = "CUADROS PORCINO", x = paste0(nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "D", startRow = 22)
writeData(wb, sheet = "CUADROS PORCINO", x = paste0(nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "E", startRow = 22)
writeData(wb, sheet = "CUADROS PORCINO", x = paste0(anio-1,"-",trim_rom," / ",anio-2,"-",trim_rom),colNames = FALSE,startCol = "H", startRow = 6)
writeData(wb, sheet = "CUADROS PORCINO", x = paste0(anio,"-",trim_rom," / ",anio-1,"-",trim_rom),colNames = FALSE,startCol = "I", startRow = 6)
writeData(wb, sheet = "CUADROS PORCINO", x = paste0(anio-1,"-",trim_rom," / ",anio-2,"-",trim_rom),colNames = FALSE,startCol = "H", startRow = 12)
writeData(wb, sheet = "CUADROS PORCINO", x = paste0(anio,"-",trim_rom," / ",anio-1,"-",trim_rom),colNames = FALSE,startCol = "I", startRow = 12)
writeData(wb, sheet = "CUADROS PORCINO", x = paste0(anio-1,"-",trim_rom," / ",anio-2,"-",trim_rom),colNames = FALSE,startCol = "H", startRow = 17)
writeData(wb, sheet = "CUADROS PORCINO", x = paste0(anio,"-",trim_rom," / ",anio-1,"-",trim_rom),colNames = FALSE,startCol = "I", startRow = 17)
writeData(wb, sheet = "CUADROS PORCINO", x = paste0(anio-1,"-",trim_rom," / ",anio-2,"-",trim_rom),colNames = FALSE,startCol = "H", startRow = 22)
writeData(wb, sheet = "CUADROS PORCINO", x = paste0(anio,"-",trim_rom," / ",anio-1,"-",trim_rom),colNames = FALSE,startCol = "I", startRow = 22)
writeData(wb, sheet = "CUADROS PORCINO", x = paste0("Variación Año corrido ",nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "F", startRow = 17)
writeData(wb, sheet = "CUADROS PORCINO", x = paste0("Variación Año corrido ",nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "G", startRow = 17)
writeData(wb, sheet = "CUADROS PORCINO", x = paste0("Variación Año corrido ",nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "F", startRow = 22)
writeData(wb, sheet = "CUADROS PORCINO", x = paste0("Variación Año corrido ",nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "G", startRow = 22)

Valores_totales<-matrix(c(paste0("'Porcino kilo en pie'!N",fila_anterior+5),
                          paste0("'Porcino kilo en pie'!O",fila_anterior+5),
                          paste0("'Porcino kilo en pie'!P",fila_anterior+5),
                          paste0("'Porcino en cabezas'!K",fila_anterior+5),
                          paste0("'Porcino kilo en pie'!N",ultima_fila+6),
                          paste0("'Porcino kilo en pie'!O",ultima_fila+6),
                          paste0("'Porcino kilo en pie'!P",ultima_fila+6),
                          paste0("'Porcino en cabezas'!K",ultima_fila+6),
                          "SUM(F8:F9)",
                          paste0("'Porcino kilo en pie'!K",fila_anterior+5),
                          paste0("'Porcino kilo en pie'!L",fila_anterior+5),
                          "",
                          "SUM(G8:G9)",
                          "E8*F8/100",
                          "E9*F9/100",
                          "",
                          paste0("'Porcino kilo en pie'!AD",fila_anterior+5),
                          paste0("'Porcino kilo en pie'!AE",fila_anterior+5),
                          paste0("'Porcino kilo en pie'!AF",fila_anterior+5),
                          paste0("'Porcino en cabezas'!U",fila_anterior+5),
                          paste0("'Porcino kilo en pie'!AD",ultima_fila+6),
                          paste0("'Porcino kilo en pie'!AE",ultima_fila+6),
                          paste0("'Porcino kilo en pie'!AF",ultima_fila+6),
                          paste0("'Porcino en cabezas'!U",ultima_fila+6)
),nrow=4,ncol=6,byrow = FALSE)

for (i in 1:4) {
  for (j in 1:6) {
    writeFormula(wb, sheet ="CUADROS PORCINO" , x = Valores_totales[i,j] ,startCol = j+3, startRow = i+6)
  }
}

Valores_consumo<-matrix(c(paste0("'CI_Carne'!AD",fila_anterior+5),
                          paste0("'CI_Carne'!AE",fila_anterior+5),
                          paste0("'CI_Carne'!AF",fila_anterior+5),
                          paste0("'CI_Carne'!AG",fila_anterior+5),
                          paste0("'CI_Carne'!AD",ultima_fila+6),
                          paste0("'CI_Carne'!AE",ultima_fila+6),
                          paste0("'CI_Carne'!AF",ultima_fila+6),
                          paste0("'CI_Carne'!AG",ultima_fila+6),
                          "SUM(F14:F16)",
                          paste0("'CI_Carne'!T",fila_anterior+5),
                          paste0("'CI_Carne'!U",fila_anterior+5),
                          paste0("'CI_Carne'!V",fila_anterior+5),
                          "SUM(G14:G16)",
                          "E14*F14/100",
                          "E15*F15/100",
                          "E16*F16/100",
                          paste0("'CI_Carne'!BB",fila_anterior+5),
                          paste0("'CI_Carne'!BC",fila_anterior+5),
                          paste0("'CI_Carne'!BD",fila_anterior+5),
                          paste0("'CI_Carne'!BE",fila_anterior+5),
                          paste0("'CI_Carne'!BB",ultima_fila+6),
                          paste0("'CI_Carne'!BC",ultima_fila+6),
                          paste0("'CI_Carne'!BD",ultima_fila+6),
                          paste0("'CI_Carne'!BE",ultima_fila+6)
),nrow=4,ncol=6,byrow = FALSE)

for (i in 1:4) {
  for (j in 1:6) {
    writeFormula(wb, sheet ="CUADROS PORCINO" , x = Valores_consumo[i,j] ,startCol = j+3, startRow = i+12)
  }
}


Valores_precios<-matrix(c(paste0("'Precios'!AF",fila_anterior_precios+5),
                          paste0("'Precios'!AG",fila_anterior_precios+5),
                          paste0("'Precios'!AD",fila_anterior_precios+5),
                          paste0("'Precios'!AE",fila_anterior_precios+5),
                          paste0("'Precios'!AF",ultima_fila_precios+6),
                          paste0("'Precios'!AG",ultima_fila_precios+6),
                          paste0("'Precios'!AD",ultima_fila_precios+6),
                          paste0("'Precios'!AE",ultima_fila_precios+6),
                          paste0("'Precios'!AV",fila_anterior_precios+5),
                          paste0("'Precios'!AW",fila_anterior_precios+5),
                          paste0("'Precios'!AT",fila_anterior_precios+5),
                          paste0("'Precios'!AU",fila_anterior_precios+5),
                          paste0("'Precios'!AV",ultima_fila_precios+6),
                          paste0("'Precios'!AW",ultima_fila_precios+6),
                          paste0("'Precios'!AT",ultima_fila_precios+6),
                          paste0("'Precios'!AU",ultima_fila_precios+6),
                          paste0("'Precios'!BL",fila_anterior_precios+5),
                          paste0("'Precios'!BM",fila_anterior_precios+5),
                          paste0("'Precios'!BJ",fila_anterior_precios+5),
                          paste0("'Precios'!BK",fila_anterior_precios+5),
                          paste0("'Precios'!BL",ultima_fila_precios+6),
                          paste0("'Precios'!BM",ultima_fila_precios+6),
                          paste0("'Precios'!BJ",ultima_fila_precios+6),
                          paste0("'Precios'!BK",ultima_fila_precios+6)

),nrow=4,ncol=6,byrow = FALSE)

for (i in 1:4) {
  for (j in 1:6) {
    writeFormula(wb, sheet ="CUADROS PORCINO" , x = Valores_precios[i,j] ,startCol = j+3, startRow = i+17)
  }
}


#IMPORTACIONES

writeFormula(wb, sheet ="CUADROS PORCINO" , x = paste0("'PORCINO kilo en pie'!R",fila_anterior+5) ,startCol = "D", startRow =27 )
writeFormula(wb, sheet ="CUADROS PORCINO" , x = paste0("'PORCINO kilo en pie'!R",ultima_fila+6) ,startCol = "E", startRow =27 )
writeFormula(wb, sheet ="CUADROS PORCINO" , x = paste0("'PORCINO kilo en pie'!Z",fila_anterior+5) ,startCol = "F", startRow =27 )
writeFormula(wb, sheet ="CUADROS PORCINO" , x = paste0("'PORCINO kilo en pie'!Z",ultima_fila+6) ,startCol = "G", startRow =27 )
writeFormula(wb, sheet ="CUADROS PORCINO" , x = paste0("'PORCINO kilo en pie'!AH",fila_anterior+5) ,startCol = "H", startRow =27 )
writeFormula(wb, sheet ="CUADROS PORCINO" , x = paste0("'PORCINO kilo en pie'!AH",ultima_fila+6) ,startCol = "I", startRow =27 )

setRowHeights(wb,sheet ="CUADROS PORCINO",rows = c(11,23:26),heights = 0)




# Cuadro AVICULTURA -----------------------------------------------------------

writeData(wb, sheet = "CUADROS AVICULTURA", x = paste0(trim_rom," Trimestre ","Variación Anual "),colNames = FALSE,startCol = "H", startRow = 5)
writeData(wb, sheet = "CUADROS AVICULTURA", x = paste0(nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "D", startRow = 6)
writeData(wb, sheet = "CUADROS AVICULTURA", x = paste0(nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "E", startRow = 6)
writeData(wb, sheet = "CUADROS AVICULTURA", x = paste0(nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "D", startRow = 11)
writeData(wb, sheet = "CUADROS AVICULTURA", x = paste0(nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "E", startRow = 11)
writeData(wb, sheet = "CUADROS AVICULTURA", x = paste0(nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "D", startRow = 16)
writeData(wb, sheet = "CUADROS AVICULTURA", x = paste0(nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "E", startRow = 16)
writeData(wb, sheet = "CUADROS AVICULTURA", x = paste0(anio-1,"-",trim_rom," / ",anio-2,"-",trim_rom),colNames = FALSE,startCol = "H", startRow = 6)
writeData(wb, sheet = "CUADROS AVICULTURA", x = paste0(anio,"-",trim_rom," / ",anio-1,"-",trim_rom),colNames = FALSE,startCol = "I", startRow = 6)
writeData(wb, sheet = "CUADROS AVICULTURA", x = paste0(anio-1,"-",trim_rom," / ",anio-2,"-",trim_rom),colNames = FALSE,startCol = "H", startRow = 11)
writeData(wb, sheet = "CUADROS AVICULTURA", x = paste0(anio,"-",trim_rom," / ",anio-1,"-",trim_rom),colNames = FALSE,startCol = "I", startRow = 11)
writeData(wb, sheet = "CUADROS AVICULTURA", x = paste0(anio-1,"-",trim_rom," / ",anio-2,"-",trim_rom),colNames = FALSE,startCol = "H", startRow = 16)
writeData(wb, sheet = "CUADROS AVICULTURA", x = paste0(anio,"-",trim_rom," / ",anio-1,"-",trim_rom),colNames = FALSE,startCol = "I", startRow = 16)
writeData(wb, sheet = "CUADROS AVICULTURA", x = paste0("Variación Año corrido ",nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "F", startRow = 6)
writeData(wb, sheet = "CUADROS AVICULTURA", x = paste0("Variación Año corrido ",nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "G", startRow = 6)
writeData(wb, sheet = "CUADROS AVICULTURA", x = paste0("Variación Año corrido ",nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "F", startRow = 11)
writeData(wb, sheet = "CUADROS AVICULTURA", x = paste0("Variación Año corrido ",nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "G", startRow = 11)
writeData(wb, sheet = "CUADROS AVICULTURA", x = paste0("Variación Año corrido ",nombres_meses[mes]," ",anio-1),colNames = FALSE,startCol = "F", startRow = 16)
writeData(wb, sheet = "CUADROS AVICULTURA", x = paste0("Variación Año corrido ",nombres_meses[mes]," ",anio),colNames = FALSE,startCol = "G", startRow = 16)


Valores_totales<-matrix(c(paste0("'Pollo_Huevo'!AA",fila_anterior+5),
                          paste0("'Pollo_Huevo'!Z",fila_anterior+5),
                          paste0("'Pollo_Huevo'!AB",fila_anterior+4),
                          paste0("'Pollo_Huevo'!AD",fila_anterior+5),
                          paste0("'Pollo_Huevo'!AA",ultima_fila+6),
                          paste0("'Pollo_Huevo'!Z",ultima_fila+6),
                          paste0("'Pollo_Huevo'!AB",ultima_fila+5),
                          paste0("'Pollo_Huevo'!AD",ultima_fila+6),
                          paste0("'Pollo_Huevo'!AO",fila_anterior+5),
                          paste0("'Pollo_Huevo'!AN",fila_anterior+5),
                          paste0("'Pollo_Huevo'!AP",fila_anterior+4),
                          paste0("'Pollo_Huevo'!AR",fila_anterior+5),
                          paste0("'Pollo_Huevo'!AO",ultima_fila+6),
                          paste0("'Pollo_Huevo'!AN",ultima_fila+6),
                          paste0("'Pollo_Huevo'!AP",ultima_fila+5),
                          paste0("'Pollo_Huevo'!AR",ultima_fila+6),
                          paste0("'Pollo_Huevo'!BC",fila_anterior+5),
                          paste0("'Pollo_Huevo'!BB",fila_anterior+5),
                          paste0("'Pollo_Huevo'!BD",fila_anterior+4),
                          paste0("'Pollo_Huevo'!BF",fila_anterior+5),
                          paste0("'Pollo_Huevo'!BC",ultima_fila+6),
                          paste0("'Pollo_Huevo'!BB",ultima_fila+6),
                          paste0("'Pollo_Huevo'!BD",ultima_fila+5),
                          paste0("'Pollo_Huevo'!BF",ultima_fila+6)
),nrow=4,ncol=6,byrow = FALSE)

for (i in 1:4) {
  for (j in 1:6) {
    writeFormula(wb, sheet ="CUADROS AVICULTURA" , x = Valores_totales[i,j] ,startCol = j+3, startRow = i+6)
  }
}

Valores_precios<-matrix(c(paste0("'Precios'!AK",fila_anterior_precios+5),
                          paste0("'Precios'!AM",fila_anterior_precios+5),
                          paste0("'Precios'!AP",fila_anterior_precios+5),
                          paste0("'Precios'!AN",fila_anterior_precios+5),
                          paste0("'Precios'!AK",ultima_fila_precios+6),
                          paste0("'Precios'!AM",ultima_fila_precios+6),
                          paste0("'Precios'!AP",ultima_fila_precios+6),
                          paste0("'Precios'!AN",ultima_fila_precios+6),
                          paste0("'Precios'!BA",fila_anterior_precios+5),
                          paste0("'Precios'!BC",fila_anterior_precios+5),
                          paste0("'Precios'!BF",fila_anterior_precios+5),
                          paste0("'Precios'!BD",fila_anterior_precios+5),
                          paste0("'Precios'!BA",ultima_fila_precios+6),
                          paste0("'Precios'!BC",ultima_fila_precios+6),
                          paste0("'Precios'!BF",ultima_fila_precios+6),
                          paste0("'Precios'!BD",ultima_fila_precios+6),
                          paste0("'Precios'!BQ",fila_anterior_precios+5),
                          paste0("'Precios'!BS",fila_anterior_precios+5),
                          paste0("'Precios'!BU",fila_anterior_precios+5),
                          paste0("'Precios'!BT",fila_anterior_precios+5),
                          paste0("'Precios'!BQ",ultima_fila_precios+6),
                          paste0("'Precios'!BS",ultima_fila_precios+6),
                          paste0("'Precios'!BU",ultima_fila_precios+6),
                          paste0("'Precios'!BT",ultima_fila_precios+6)

),nrow=4,ncol=6,byrow = FALSE)

for (i in 1:4) {
  for (j in 1:6) {
    writeFormula(wb, sheet ="CUADROS AVICULTURA" , x = Valores_precios[i,j] ,startCol = j+3, startRow = i+11)
  }
}


#IMPORTACIONES

Valores_exterior<-matrix(c(paste0("'Pollo_Huevo'!AG",fila_anterior+5),
                          paste0("'Pollo_Huevo'!AK",fila_anterior+5),
                          paste0("'Pollo_Huevo'!AJ",fila_anterior+5),
                          paste0("'Pollo_Huevo'!AG",ultima_fila+6),
                          paste0("'Pollo_Huevo'!AK",ultima_fila+6),
                          paste0("'Pollo_Huevo'!AJ",ultima_fila+6),
                          paste0("'Pollo_Huevo'!AU",fila_anterior+5),
                          paste0("'Pollo_Huevo'!AY",fila_anterior+5),
                          paste0("'Pollo_Huevo'!AZ",fila_anterior+5),
                          paste0("'Pollo_Huevo'!AU",ultima_fila+6),
                          paste0("'Pollo_Huevo'!AY",ultima_fila+6),
                          paste0("'Pollo_Huevo'!AZ",ultima_fila+6),
                          paste0("'Pollo_Huevo'!BI",fila_anterior+5),
                          paste0("'Pollo_Huevo'!BM",fila_anterior+5),
                          paste0("'Pollo_Huevo'!BL",fila_anterior+5),
                          paste0("'Pollo_Huevo'!BI",ultima_fila+6),
                          paste0("'Pollo_Huevo'!BM",ultima_fila+6),
                          paste0("'Pollo_Huevo'!BL",ultima_fila+6)
),nrow=3,ncol=6,byrow = FALSE)

for (i in 1:3) {
  for (j in 1:6) {
    writeFormula(wb, sheet ="CUADROS AVICULTURA" , x = Valores_exterior[i,j] ,startCol = j+3, startRow = i+16)
  }
}


setColWidths(wb,sheet ="CUADROS LECHE",cols = c(8,9),hidden = c(TRUE,TRUE))

if(mes %in% c(3,6,9,12)){
  setColWidths(wb,sheet ="CUADROS BOVINO",cols = c(8,9),widths = 8)
  setColWidths(wb,sheet ="CUADROS LECHE",cols = c(6,7),widths = 8)
  setColWidths(wb,sheet ="CUADROS PORCINO",cols = c(8,9),widths = 8)
  setColWidths(wb,sheet ="CUADROS AVICULTURA",cols = c(8,9),widths = 8)
}else{
  setColWidths(wb,sheet ="CUADROS BOVINO",cols = c(8,9),hidden = c(TRUE,TRUE))
  setColWidths(wb,sheet ="CUADROS LECHE",cols = c(6,7),hidden = c(TRUE,TRUE))
  setColWidths(wb,sheet ="CUADROS PORCINO",cols = c(8,9),hidden = c(TRUE,TRUE))
  setColWidths(wb,sheet ="CUADROS AVICULTURA",cols = c(8,9),hidden = c(TRUE,TRUE))
}



# Guardar el libro --------------------------------------------------------


if (!file.exists(salida)) {
  saveWorkbook(wb, file = salida)
} else {
  saveWorkbook(wb, file = salida,overwrite= TRUE)
}
}
