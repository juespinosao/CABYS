#' @export
# Cambios_Mes
# Cargar la biblioteca readxl

f_Cambios_Mes<-function(directorio,mes,anio){

  #Cargar librerias
  library(readxl)
  library(dplyr)
  library(openxlsx)
  library(zoo)




  # Actualizacion mensual ---------------------------------------------------

  if(mes==1){
    carpeta_anterior=nombre_carpeta(12,(anio-1))
    mes_ant_per=paste0(directorio,"/ISE/",anio-1,"/",carpeta_anterior,"/Results/ZG1_Permanentes_ISE_",nombres_meses[12],"_",anio-1,".xlsx")
    mes_ant_tran=paste0(directorio,"/ISE/",anio-1,"/",carpeta_anterior,"/Results/ZG1_Transitorios_ISE_",nombres_meses[12],"_",anio-1,".xlsx")
    mes_ant_pecu=paste0(directorio,"/ISE/",anio-1,"/",carpeta_anterior,"/Results/ZG2_Pecuario_ISE_",nombres_meses[12],"_",anio-1,".xlsx")

  }else{
    carpeta_anterior=nombre_carpeta(mes-1,anio)
    mes_ant_per=paste0(directorio,"/ISE/",anio,"/",carpeta_anterior,"/Results/ZG1_Permanentes_ISE_",nombres_meses[mes-1],"_",anio,".xlsx")
    mes_ant_tran=paste0(directorio,"/ISE/",anio,"/",carpeta_anterior,"/Results/ZG1_Transitorios_ISE_",nombres_meses[mes-1],"_",anio,".xlsx")
    mes_ant_pecu=paste0(directorio,"/ISE/",anio,"/",carpeta_anterior,"/Results/ZG2_Pecuario_ISE_",nombres_meses[mes-1],"_",anio,".xlsx")

  }

  carpeta_actual=nombre_carpeta(mes,anio)

  #Permanentes
  mes_act_per=paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Results/ZG1_Permanentes_ISE_",nombres_meses[mes],"_",anio,".xlsx")

  # Cargar el archivo de entrada
  wb_ant_per <- loadWorkbook(mes_ant_per)
  wb_act_per <- loadWorkbook(mes_act_per)

  #Transitorio
  mes_act_tran=paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Results/ZG1_Transitorios_ISE_",nombres_meses[mes],"_",anio,".xlsx")

  # Cargar el archivo de entrada
  wb_ant_tran <- loadWorkbook(mes_ant_tran)
  wb_act_tran <- loadWorkbook(mes_act_tran)


  #Pecuario
  mes_act_pecu=paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Results/ZG2_Pecuario_ISE_",nombres_meses[mes],"_",anio,".xlsx")

  # Cargar el archivo de entrada
  wb_ant_pecu <- loadWorkbook(mes_ant_pecu)
  wb_act_pecu <- loadWorkbook(mes_act_pecu)

  if(mes==1){
   mes=13
   anio=anio-1
  }else{
  }


#Otras Frutas


  data_ant <- read.xlsx(wb_ant_per, sheet = "Otras frutas.", colNames = TRUE,startRow = 2)
  data_act <- read.xlsx(wb_act_per, sheet = "Otras frutas.", colNames = TRUE,startRow = 2)
  fila=which(data_ant$Año==anio & data_ant$Mes==(mes-1))

  tabla=data.frame(Producto="Otras frutas",
                   Anterior=data_ant[fila,"Variacion.Indice.Retropolado"],
                   Actual=data_act[fila,"Variacion.Indice.Retropolado"]
  )

#Frutas citricas

  data_ant <- read.xlsx(wb_ant_per, sheet = "Frutas Citricas", colNames = TRUE,startRow = 2)
  data_act <- read.xlsx(wb_act_per, sheet = "Frutas Citricas", colNames = TRUE,startRow = 2)


    fila=which(data_ant$Año==anio & data_ant$Mes==(mes-1))
  tabla=rbind(tabla,c("Frutas cítricas",data_ant[fila,"Variacion.Indice.Retropolado"],data_act[fila,"Variacion.Indice.Retropolado"]))


  #Yuca
  data_ant <- read.xlsx(wb_ant_tran, sheet = "Yuca", colNames = TRUE,startRow = 10)
  data_act <- read.xlsx(wb_act_tran, sheet = "Yuca", colNames = TRUE,startRow = 10)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-1))

  tabla=rbind(tabla,c("Yuca",data_ant[fila,"Variación.Anual"],data_act[fila,"Variación.Anual"]))

  #Banano
  data_ant <- read.xlsx(wb_ant_per, sheet = "Banano Total(Expos+Interno)", colNames = TRUE,startRow = 11)
  data_act <- read.xlsx(wb_act_per, sheet = "Banano Total(Expos+Interno)", colNames = TRUE,startRow = 11)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-1))

  tabla=rbind(tabla,c("Banano",data_ant[fila,"Índice.de.producción.ponderado.Variación.Anual"],data_act[fila,"Índice.de.producción.ponderado.Variación.Anual"]))

  #Flores
  data_ant <- read.xlsx(wb_ant_per, sheet = "Flores", colNames = TRUE,startRow = 10)
  data_act <- read.xlsx(wb_act_per, sheet = "Flores", colNames = TRUE,startRow = 10)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-1))

  tabla=rbind(tabla,c("Flores",data_ant[fila,"Variacion.Anual"],data_act[fila,"Variacion.Anual"]))

  #Platano
  data_ant <- read.xlsx(wb_ant_per, sheet = "Plátano Total(Expos+Interno)", colNames = TRUE,startRow = 11)
  data_act <- read.xlsx(wb_act_per, sheet = "Plátano Total(Expos+Interno)", colNames = TRUE,startRow = 11)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-1))

  tabla=rbind(tabla,c("Platano",data_ant[fila,"Índice.de.producción.ponderado.Variación.Anual"],data_act[fila,"Índice.de.producción.ponderado.Variación.Anual"]))

  #Hortalizas
  data_ant <- read.xlsx(wb_ant_tran, sheet = "Hortalizas", colNames = TRUE,startRow = 10)
  data_act <- read.xlsx(wb_act_tran, sheet = "Hortalizas", colNames = TRUE,startRow = 10)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-1))

  tabla=rbind(tabla,c("Hortalizas",data_ant[fila,"Variación.Anual"],data_act[fila,"Variación.Anual"]))

  #Aceite de palma

data_ant <- read.xlsx(wb_ant_per, sheet = "Aceite de palma", colNames = TRUE,startRow = 9)
data_act <- read.xlsx(wb_act_per, sheet = "Aceite de palma", colNames = TRUE,startRow = 9)
fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-1))

tabla=rbind(tabla,c("Aceite de palma",data_ant[fila,"Variacion.Anual"],data_act[fila,"Variacion.Anual"]))

  #Papa
  data_ant <- read.xlsx(wb_ant_tran, sheet = "Papa", colNames = TRUE,startRow = 10)
  data_act <- read.xlsx(wb_act_tran, sheet = "Papa", colNames = TRUE,startRow = 10)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-1))

  tabla=rbind(tabla,c("Papa",data_ant[fila,"Variación.Anual"],data_act[fila,"Variación.Anual"]))

  #Caña de azucar
  data_ant <- read.xlsx(wb_ant_per, sheet = "Caña de Azúcar", colNames = TRUE,startRow = 9)
  data_act <- read.xlsx(wb_act_per, sheet = "Caña de Azúcar", colNames = TRUE,startRow = 9)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-1))

  tabla=rbind(tabla,c("Caña de azúcar",data_ant[fila,"Variación.Anual"],data_act[fila,"Variación.Anual"]))

  #Cacao
  data_ant <- read.xlsx(wb_ant_per, sheet = "Cacao", colNames = TRUE,startRow = 9)
  data_act <- read.xlsx(wb_act_per, sheet = "Cacao", colNames = TRUE,startRow = 9)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-1))

  tabla=rbind(tabla,c("Cacao",data_ant[fila,"Variacion.Anual"],data_act[fila,"Variacion.Anual"]))

  #Panela
  data_ant <- read.xlsx(wb_ant_per, sheet = "Panela", colNames = TRUE,startRow = 9)
  data_act <- read.xlsx(wb_act_per, sheet = "Panela", colNames = TRUE,startRow = 9)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-1))

  tabla=rbind(tabla,c("Caña panelera",data_ant[fila,"Variacion.Anual"],data_act[fila,"Variacion.Anual"]))

  #Fruto de palma
  data_ant <- read.xlsx(wb_ant_per, sheet = "Fruto de Palma", colNames = TRUE,startRow = 9)
  data_act <- read.xlsx(wb_act_per, sheet = "Fruto de Palma", colNames = TRUE,startRow = 9)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-1))

  tabla=rbind(tabla,c("Fruto de palma",data_ant[fila,"Variacion.Anual"],data_act[fila,"Variacion.Anual"]))

  #Maiz
  data_ant <- read.xlsx(wb_ant_tran, sheet = "Maíz", colNames = TRUE,startRow = 10)
  data_act <- read.xlsx(wb_act_tran, sheet = "Maíz", colNames = TRUE,startRow = 10)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-1))

  tabla=rbind(tabla,c("Maiz",data_ant[fila,"Variacion.Anual"],data_act[fila,"Variacion.Anual"]))

  #Legumbres
  data_ant <- read.xlsx(wb_ant_tran, sheet = "Legumbres", colNames = TRUE,startRow = 10)
  data_act <- read.xlsx(wb_act_tran, sheet = "Legumbres", colNames = TRUE,startRow = 10)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-1))

  tabla=rbind(tabla,c("Legumbres verdes y secas",data_ant[fila,"Variacion.Anual"],data_act[fila,"Variacion.Anual"]))

  #Arroz
  data_ant <- read.xlsx(wb_ant_tran, sheet = "Arroz", colNames = TRUE,startRow = 10)
  data_act <- read.xlsx(wb_act_tran, sheet = "Arroz", colNames = TRUE,startRow = 10)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-1))

  tabla=rbind(tabla,c("Arroz",data_ant[fila,"Variación.Anual"],data_act[fila,"Variación.Anual"]))

  #Cafetos
  data_ant <- read.xlsx(wb_ant_per, sheet = "Cafetos", colNames = TRUE,startRow = 9)
  data_act <- read.xlsx(wb_act_per, sheet = "Cafetos", colNames = TRUE,startRow = 9)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-1))

  tabla=rbind(tabla,c("Cafetos",data_ant[fila,"Variacion.Anual"],data_act[fila,"Variacion.Anual"]))

  #Cafe pergamino
  data_ant <- read.xlsx(wb_ant_per, sheet = "Cafe Pergamino", colNames = TRUE,startRow = 9)
  data_act <- read.xlsx(wb_act_per, sheet = "Cafe Pergamino", colNames = TRUE,startRow = 9)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-1))

  tabla=rbind(tabla,c("Café pergamino",data_ant[fila,"Variacion.Procuccion.pergamino"],data_act[fila,"Variacion.Procuccion.pergamino"]))

  #leche
  data_ant <- read.xlsx(wb_ant_pecu, sheet = "Leche", colNames = TRUE,startRow = 10)
  data_act <- read.xlsx(wb_act_pecu, sheet = "Leche", colNames = TRUE,startRow = 10)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-1))

  tabla=rbind(tabla,c("Leche",data_ant[fila,"Variación.Anual"],data_act[fila,"Variación.Anual"]))

  #Huevos
  data_ant <- read.xlsx(wb_ant_pecu, sheet = "Huevos", colNames = TRUE,startRow = 10)
  data_act <- read.xlsx(wb_act_pecu, sheet = "Huevos", colNames = TRUE,startRow = 10)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-1))

  tabla=rbind(tabla,c("Huevos",data_ant[fila,"Variacion.anual"],data_act[fila,"Variacion.anual"]))

  #Ganado bovino
  data_ant <- read.xlsx(wb_ant_pecu, sheet = "Ganado_Bovino", colNames = TRUE,startRow = 10)
  data_act <- read.xlsx(wb_act_pecu, sheet = "Ganado_Bovino", colNames = TRUE,startRow = 10)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-1))

  tabla=rbind(tabla,c("Ganado bovino",data_ant[fila,"Variacion.Anual"],data_act[fila,"Variacion.Anual"]))

  #Ganado porcino
  data_ant <- read.xlsx(wb_ant_pecu, sheet = "Porcino", colNames = TRUE,startRow = 10)
  data_act <- read.xlsx(wb_act_pecu, sheet = "Porcino", colNames = TRUE,startRow = 10)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-1))

  tabla=rbind(tabla,c("Ganado porcino",data_ant[fila,"Variacion.Anual"],data_act[fila,"Variacion.Anual"]))

  #Aves de corral
  data_ant <- read.xlsx(wb_ant_pecu, sheet = "Pollos", colNames = TRUE,startRow = 10)
  data_act <- read.xlsx(wb_act_pecu, sheet = "Pollos", colNames = TRUE,startRow = 10)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-1))

  tabla=rbind(tabla,c("Aves de corral",data_ant[fila,"Variación.Anual"],data_act[fila,"Variación.Anual"]))
  tabla$Anterior=as.numeric(tabla$Anterior)
  tabla$Actual=as.numeric(tabla$Actual)
  tabla$Diferencia=tabla[,"Actual"]-tabla[,"Anterior"]

return(tabla)
}
