#' @export
# Cambios_Mes
# Cargar la biblioteca readxl

f_Cambios_Trim<-function(directorio,mes,anio){

  #Cargar librerias
  library(readxl)
  library(dplyr)
  library(openxlsx)
  library(zoo)

  mes_ori=mes
  anio_ori=anio


  # Actualizacion trimestral ---------------------------------------------------
  if(mes==3){
    carpeta_anterior=nombre_carpeta(12,(anio-1))
    mes_ant_per=paste0(directorio,"/ISE/",anio-1,"/",carpeta_anterior,"/Results/ZG1_Permanentes_ISE_",nombres_meses[12],"_",anio-1,".xlsx")
    mes_ant_tran=paste0(directorio,"/ISE/",anio-1,"/",carpeta_anterior,"/Results/ZG1_Transitorios_ISE_",nombres_meses[12],"_",anio-1,".xlsx")
    mes_ant_pecu=paste0(directorio,"/ISE/",anio-1,"/",carpeta_anterior,"/Results/ZG2_Pecuario_ISE_",nombres_meses[12],"_",anio-1,".xlsx")

  }else{
    carpeta_anterior=nombre_carpeta(mes-3,anio)
    mes_ant_per=paste0(directorio,"/ISE/",anio,"/",carpeta_anterior,"/Results/ZG1_Permanentes_ISE_",nombres_meses[(mes-3)],"_",anio,".xlsx")
    mes_ant_tran=paste0(directorio,"/ISE/",anio,"/",carpeta_anterior,"/Results/ZG1_Transitorios_ISE_",nombres_meses[mes-3],"_",anio,".xlsx")
    mes_ant_pecu=paste0(directorio,"/ISE/",anio,"/",carpeta_anterior,"/Results/ZG2_Pecuario_ISE_",nombres_meses[mes-3],"_",anio,".xlsx")

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


  if(mes==3){
    mes=15
    anio=anio-1
   trimestre=5
  }else{
    trimestre=f_trimestre(mes)
  }


  #Otras Frutas


  data_ant <- read.xlsx(wb_ant_per, sheet = "Otras frutas.", colNames = TRUE,startRow = 2)
  data_act <- read.xlsx(wb_act_per, sheet = "Otras frutas.", colNames = TRUE,startRow = 2)
  fila=which(data_ant[,45]==anio )+trimestre-2

  tabla=data.frame(Producto="Otras frutas",
                   Anterior=data_ant[fila,49],
                   Actual=data_act[fila,49]
  )

  #Frutas citricas

  data_ant <- read.xlsx(wb_ant_per, sheet = "Frutas Citricas", colNames = TRUE,startRow = 2)
  data_act <- read.xlsx(wb_act_per, sheet = "Frutas Citricas", colNames = TRUE,startRow = 2)


  fila=which(data_ant[,45]==anio )+trimestre-2
  tabla=rbind(tabla,c("Frutas cítricas",data_ant[fila,49],data_act[fila,49]))



  #Legumbres
  data_ant <- read.xlsx(wb_ant_tran, sheet = "Legumbres", colNames = TRUE,startRow = 10)
  data_act <- read.xlsx(wb_act_tran, sheet = "Legumbres", colNames = TRUE,startRow = 10)

  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-3))
  tabla=rbind(tabla,c("Legumbres",data_ant[fila,"Estado"],data_act[fila,"Estado"]))

  #Flores
  data_ant <- read.xlsx(wb_ant_per, sheet = "Flores", colNames = TRUE,startRow = 10)
  data_act <- read.xlsx(wb_act_per, sheet = "Flores", colNames = TRUE,startRow = 10)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-3))

  tabla=rbind(tabla,c("Flores",data_ant[fila,"Estado"],data_act[fila,"Estado"]))

  #Caña panelera
  data_ant <- read.xlsx(wb_ant_per, sheet = "Panela", colNames = TRUE,startRow = 9)
  data_act <- read.xlsx(wb_act_per, sheet = "Panela", colNames = TRUE,startRow = 9)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-3))

  tabla=rbind(tabla,c("Caña panelera",data_ant[fila,"Estado"],data_act[fila,"Estado"]))

  #Hortalizas
  data_ant <- read.xlsx(wb_ant_tran, sheet = "Hortalizas", colNames = TRUE,startRow = 10)
  data_act <- read.xlsx(wb_act_tran, sheet = "Hortalizas", colNames = TRUE,startRow = 10)

  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-3))
  tabla=rbind(tabla,c("Hortalizas",data_ant[fila,"Estado"],data_act[fila,"Estado"]))

  #Banano
  data_ant <- read.xlsx(wb_ant_per, sheet = "Banano Total(Expos+Interno)", colNames = TRUE,startRow = 11)
  data_act <- read.xlsx(wb_act_per, sheet = "Banano Total(Expos+Interno)", colNames = TRUE,startRow = 11)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-3))

  tabla=rbind(tabla,c("Banano",data_ant[fila,"total.trimestral"],data_act[fila,"total.trimestral"]))

  #Fruto de palma
  data_ant <- read.xlsx(wb_ant_per, sheet = "Fruto de Palma", colNames = TRUE,startRow = 9)
  data_act <- read.xlsx(wb_act_per, sheet = "Fruto de Palma", colNames = TRUE,startRow = 9)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-3))

  tabla=rbind(tabla,c("Fruto de palma",data_ant[fila,"Estado"],data_act[fila,"Estado"]))

  #Papa
  data_ant <- read.xlsx(wb_ant_tran, sheet = "Papa", colNames = TRUE,startRow = 10)
  data_act <- read.xlsx(wb_act_tran, sheet = "Papa", colNames = TRUE,startRow = 10)

  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-3))
  tabla=rbind(tabla,c("Papa",data_ant[fila,"Estado"],data_act[fila,"Estado"]))

  #Platano
  data_ant <- read.xlsx(wb_ant_per, sheet = "Plátano Total(Expos+Interno)", colNames = TRUE,startRow = 11)
  data_act <- read.xlsx(wb_act_per, sheet = "Plátano Total(Expos+Interno)", colNames = TRUE,startRow = 11)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-3))

  tabla=rbind(tabla,c("Platano",data_ant[fila,"total.trimestral"],data_act[fila,"total.trimestral"]))

  #Aceite de palma
  data_ant <- read.xlsx(wb_ant_per, sheet = "Aceite de palma", colNames = TRUE,startRow = 9)
  data_act <- read.xlsx(wb_act_per, sheet = "Aceite de palma", colNames = TRUE,startRow = 9)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-3))

  tabla=rbind(tabla,c("Aceite de palma",data_ant[fila,"Estado"],data_act[fila,"Estado"]))

  #Maiz
  data_ant <- read.xlsx(wb_ant_tran, sheet = "Maíz", colNames = TRUE,startRow = 10)
  data_act <- read.xlsx(wb_act_tran, sheet = "Maíz", colNames = TRUE,startRow = 10)

  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-3))
  tabla=rbind(tabla,c("Maíz",data_ant[fila,"Estado"],data_act[fila,"Estado"]))

  #Caña de azúcar
  data_ant <- read.xlsx(wb_ant_per, sheet = "Caña de Azúcar", colNames = TRUE,startRow = 9)
  data_act <- read.xlsx(wb_act_per, sheet = "Caña de Azúcar", colNames = TRUE,startRow = 9)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-3))

  tabla=rbind(tabla,c("Caña de azúcar",data_ant[fila,"Estado"],data_act[fila,"Estado"]))

  #Arroz
  data_ant <- read.xlsx(wb_ant_tran, sheet = "Arroz", colNames = TRUE,startRow = 10)
  data_act <- read.xlsx(wb_act_tran, sheet = "Arroz", colNames = TRUE,startRow = 10)

  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-3))
  tabla=rbind(tabla,c("Arroz",data_ant[fila,"Estado"],data_act[fila,"Estado"]))

  #Cacao
  data_ant <- read.xlsx(wb_ant_per, sheet = "Cacao", colNames = TRUE,startRow = 9)
  data_act <- read.xlsx(wb_act_per, sheet = "Cacao", colNames = TRUE,startRow = 9)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-3))

  tabla=rbind(tabla,c("Cacao en grano",data_ant[fila,"Estado"],data_act[fila,"Estado"]))

  #Yuca
  data_ant <- read.xlsx(wb_ant_tran, sheet = "Yuca", colNames = TRUE,startRow = 10)
  data_act <- read.xlsx(wb_act_tran, sheet = "Yuca", colNames = TRUE,startRow = 10)

  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-3))
  tabla=rbind(tabla,c("Yuca",data_ant[fila,"Estado"],data_act[fila,"Estado"]))

  #Café pergamino
  data_ant <- read.xlsx(wb_ant_per, sheet = "Cafe Pergamino", colNames = TRUE,startRow = 9)
  data_act <- read.xlsx(wb_act_per, sheet = "Cafe Pergamino", colNames = TRUE,startRow = 9)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-3))

  tabla=rbind(tabla,c("Café pergamino",data_ant[fila,"Estado"],data_act[fila,"Estado"]))

  #Cafetos
  data_ant <- read.xlsx(wb_ant_per, sheet = "Cafetos", colNames = TRUE,startRow = 9)
  data_act <- read.xlsx(wb_act_per, sheet = "Cafetos", colNames = TRUE,startRow = 9)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-3))

  tabla=rbind(tabla,c("Cafetos",data_ant[fila,"Estado"],data_act[fila,"Estado"]))

  #Ganado bovino
  data_ant <- read.xlsx(wb_ant_pecu, sheet = "Ganado_Bovino", colNames = TRUE,startRow = 10)
  data_act <- read.xlsx(wb_act_pecu, sheet = "Ganado_Bovino", colNames = TRUE,startRow = 10)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-3))

  tabla=rbind(tabla,c("Ganado bovino",data_ant[fila,"Estado"],data_act[fila,"Estado"]))

  #Ganado porcino
  data_ant <- read.xlsx(wb_ant_pecu, sheet = "Porcino", colNames = TRUE,startRow = 10)
  data_act <- read.xlsx(wb_act_pecu, sheet = "Porcino", colNames = TRUE,startRow = 10)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-3))

  tabla=rbind(tabla,c("Ganado porcino",data_ant[fila,"Estado"],data_act[fila,"Estado"]))

  #Aves de corral
  data_ant <- read.xlsx(wb_ant_pecu, sheet = "Pollos", colNames = TRUE,startRow = 10)
  data_act <- read.xlsx(wb_act_pecu, sheet = "Pollos", colNames = TRUE,startRow = 10)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-3))

  tabla=rbind(tabla,c("Aves de corral",data_ant[fila,"Estado"],data_act[fila,"Estado"]))

  #Leche
  data_ant <- read.xlsx(wb_ant_pecu, sheet = "Leche", colNames = TRUE,startRow = 10)
  data_act <- read.xlsx(wb_act_pecu, sheet = "Leche", colNames = TRUE,startRow = 10)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-3))

  tabla=rbind(tabla,c("Leche",data_ant[fila,"Estado"],data_act[fila,"Estado"]))

  #Huevos
  data_ant <- read.xlsx(wb_ant_pecu, sheet = "Huevos", colNames = TRUE,startRow = 10)
  data_act <- read.xlsx(wb_act_pecu, sheet = "Huevos", colNames = TRUE,startRow = 10)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-3))

  tabla=rbind(tabla,c("Huevos",data_ant[fila,"Estado"],data_act[fila,"Estado"]))


  #Ovino caprino
  data_ant <- read.xlsx(wb_ant_pecu, sheet = "Ovino y Caprino trimestral", colNames = TRUE,startRow = 10)
  data_act <- read.xlsx(wb_act_pecu, sheet = "Ovino y Caprino trimestral", colNames = TRUE,startRow = 10)
  fila=which(data_ant$Año==anio & data_ant$Periodicidad==(trimestre-1))

  tabla=rbind(tabla,c("Ovino caprino",data_ant[fila,"Variación.Anual.Trimestral"],data_act[fila,"Variación.Anual.Trimestral"]))

  #Madera

anio=anio_ori
mes=mes_ori
  if(mes==3){
    nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio-1,"/",carpeta_anterior,"/Doc/Nombres_archivos_",nombres_meses[12],".xlsx"),sheet = "Nombres")
    archivo_ant=nombre_archivos[nombre_archivos$PRODUCTO=="EMMET","NOMBRE"]
  }else{
    nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_anterior,"/Doc/Nombres_archivos_",nombres_meses[mes_ori-3],".xlsx"),sheet = "Nombres")
    archivo_ant=nombre_archivos[nombre_archivos$PRODUCTO=="EMMET","NOMBRE"]
  }
  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Doc/Nombres_archivos_",nombres_meses[mes_ori],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="EMMET","NOMBRE"]
  carpeta=nombre_carpeta(mes_ori,anio)
  # Especifica la ruta del archivo de Excel
  if(mes==3){
    EMMET_ant <- read.xlsx(paste0(directorio,"/ISE/",(anio-1),"/",carpeta_anterior,"/Data/consolidado_ISE/EMMET/",archivo_ant),
                           sheet = "COMPLETO")
  }else{
    EMMET_ant <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_anterior,"/Data/consolidado_ISE/EMMET/",archivo),
                           sheet = "COMPLETO")
  }
  # Seleccionar solo las columnas que necesitas
  EMMET_tabla_ant <- EMMET_ant[, c("anio", "mes", "Clase_CIIU4", "ProduccionRealPond")]

  EMMET_act <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/EMMET/",archivo),
                          sheet = "COMPLETO")
  # Seleccionar solo las columnas que necesitas
  EMMET_tabla_act <- EMMET_act[, c("anio", "mes", "Clase_CIIU4", "ProduccionRealPond")]

  carpeta_actual=nombre_carpeta(mes,anio)
  if(mes==3){

    carpeta_anterior=nombre_carpeta(12,(anio-1))
    nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio-1,"/",carpeta_anterior,"/Doc/Nombres_archivos_",nombres_meses[12],".xlsx"),sheet = "Nombres")
    archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Silvicultura2","NOMBRE"]
    sil_ant=paste0(directorio,"/ISE/",anio-1,"/",carpeta_anterior,"/Data/consolidado_ISE/Silvicultura/",archivo)

  }else{
    carpeta_anterior=nombre_carpeta(mes-3,anio)
    nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_anterior,"/Doc/Nombres_archivos_",nombres_meses[mes-3],".xlsx"),sheet = "Nombres")
    archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Silvicultura2","NOMBRE"]
    sil_ant=paste0(directorio,"/ISE/",anio,"/",carpeta_anterior,"/Data/consolidado_ISE/Silvicultura/",archivo)

  }

  carpeta_actual=nombre_carpeta(mes,anio)
  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Silvicultura2","NOMBRE"]
  sil_act=paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Data/consolidado_ISE/Silvicultura/",archivo)
  wb_sil_ant=loadWorkbook(sil_ant)
  wb_sil_act=loadWorkbook(sil_act)

  data_ant <- read.xlsx(wb_sil_ant, sheet ="Silvicultura producción", colNames = TRUE,startRow = 3)
  data_act <- read.xlsx(wb_sil_act, sheet ="Silvicultura producción", colNames = TRUE,startRow = 3)

  if(mes==3){
    anio=(anio-1)
    trimestre=5
  }else{

  }
  fila=which(data_ant$Año==anio & data_ant$Trimestre==(trimestre-1))

  tabla=rbind(tabla,c("Madera",data_ant[fila,6],data_act[fila,6]))


  #forestales difetentes

  tabla=rbind(tabla,c("forestales",data_ant[fila,8],data_act[fila,8]))

  #Leña

  tabla=rbind(tabla,c("Leña",data_ant[fila,7],data_act[fila,7]))

  #Pesca
  if(mes_ori==3){
    vector=c(10,11,12)
    anio_act=anio-1
    anio_ant=anio-2
    mes=12
  }else{
    vector=c(mes_ori-5,mes_ori-4,mes_ori-3)
    anio_act=anio
    anio_ant=anio-1
    mes=mes_ori
  }

  Pesca_tabla_ant=EMMET_tabla_ant %>%
    filter(Clase_CIIU4==1012 )%>%
    group_by(anio,mes)%>%
    summarise(suma=sum(ProduccionRealPond))%>%
    as.data.frame()


  valor_ant=sum(Pesca_tabla_ant[Pesca_tabla_ant$anio==anio_act & Pesca_tabla_ant$mes %in% vector,"suma"])/
    sum(Pesca_tabla_ant[Pesca_tabla_ant$anio==anio_ant & Pesca_tabla_ant$mes %in% vector,"suma"])*100-100


  Pesca_tabla_act=EMMET_tabla_act %>%
    filter(Clase_CIIU4==1012 )%>%
    group_by(anio,mes)%>%
    summarise(suma=sum(ProduccionRealPond))%>%
    as.data.frame()

  valor_act=sum(Pesca_tabla_act[Pesca_tabla_act$anio==anio_act & Pesca_tabla_act$mes %in% vector,"suma"])/
    sum(Pesca_tabla_act[Pesca_tabla_act$anio==anio_ant & Pesca_tabla_act$mes %in% vector,"suma"])*100-100

  tabla=rbind(tabla,c("Pesca",valor_ant,valor_act))

  tabla$Anterior=as.numeric(tabla$Anterior)
  tabla$Actual=as.numeric(tabla$Actual)
  tabla$Diferencia=tabla[,"Actual"]-tabla[,"Anterior"]

  mes=mes_ori
  anio=anio_ori
return(tabla)
}
