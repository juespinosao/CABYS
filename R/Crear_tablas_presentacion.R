
Tablas_presentacion<-function(directorio,mes,anio){
library(gt)
library(dplyr)
library(gtExtras)
library(readxl)
library(openxlsx)



# Actualizacion mensual ---------------------------------------------------

if(mes==1){
  carpeta_anterior=nombre_carpeta(12,(anio-1))
  mes_ant_per=paste0(directorio,"/ISE/",anio-1,"/",carpeta_anterior,"/Results/ZG1_Permanentes_ISE_",nombres_meses[mes-1],"_",anio,".xlsx")
  mes_ant_tran=paste0(directorio,"/ISE/",anio-1,"/",carpeta_anterior,"/Results/ZG1_Transitorios_ISE_",nombres_meses[mes-1],"_",anio,".xlsx")
  mes_ant_pecu=paste0(directorio,"/ISE/",anio-1,"/",carpeta_anterior,"/Results/ZG2_Pecuario_ISE_",nombres_meses[mes-1],"_",anio,".xlsx")

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



#Frutas
data_ant <- read.xlsx(wb_ant_per, sheet = "Frutas Total(Expos+Interno)", colNames = TRUE,startRow = 11)
data_act <- read.xlsx(wb_act_per, sheet = "Frutas Total(Expos+Interno)", colNames = TRUE,startRow = 11)

fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-1))

tabla=data.frame(Producto="Otras frutas",
                    Anterior=data_ant[fila,"Variación.Anual"],
                    Actual=data_act[fila,"Variación.Anual"]
)

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

#Areas en desarrollo
data_ant <- read.xlsx(wb_ant_per, sheet = "Áreas en desarrollo", colNames = TRUE,startRow = 11)
data_act <- read.xlsx(wb_act_per, sheet = "Áreas en desarrollo", colNames = TRUE,startRow = 11)
fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-1))

tabla=rbind(tabla,c("Áreas en desarrollo",as.numeric(data_ant[fila,"Variacion.Anual"]),data_act[fila,"Variacion.Anual"]))

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
tabla$Nombre=c(rep("Agricultura y actividades de servicios conexas",15),rep("Productos de café",2),rep("Ganadería, caza y actividades de servicios conexas",5))
tabla$Diferencia=tabla[,"Actual"]-tabla[,"Anterior"]

options(gt.width = 15, gt.height = 8)

tabla %>%
  mutate_if(is.numeric, ~round(., 1)) %>%  group_by(Nombre) %>%
  gt(groupname_col = "Nombre")%>%
  cols_label(
    Nombre = "Agricultura, caza, silvicultura y pesca",
    Producto = "",
    Anterior = paste0("Publicación", "\n", nombres_meses[mes]),
    Actual = paste0(nombres_meses[mes-1], "\n", "Publicación", "\n", nombres_meses[mes])
  ) %>% tab_spanner(columns = c(Anterior, Actual,Diferencia),
                    label = "Tasa de crecimiento anual (%)") %>%
  tab_options(row_group.as_column = TRUE) %>%
  gt_highlight_rows(
    rows = c(16,17), #Cambiar el 4 por el número de filas real
    # Background color
    fill = "#F2F2F2",
    font_weight = "normal"
 ) %>%

  tab_style(
    ### Estilo para el label de Variación %
    style = list(
      cell_fill(color = "#B6004B"),
      cell_text(color = "white")
    ),
    locations = cells_column_spanners(spanners = everything())
  ) %>%
  tab_style(
    ### Estilo para el resto de labels %
    style = list(
      cell_fill(color = "#B6004B"),
      cell_text(color = "white")
    ),
    locations = cells_column_labels(columns = everything())
  ) %>%
  gtsave(filename = paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Results/Coyuntura","/prueba.png"),zoom=2,vwidth = 1900,vheight = 1800)

# Actualizacion trimestral ---------------------------------------------------
if(mes==3){
  carpeta_anterior=nombre_carpeta(12,(anio-1))
}else{
  carpeta_anterior=nombre_carpeta(mes-3,anio)
}
#Permanentes
mes_ant_per=paste0(directorio,"/ISE/",anio,"/",carpeta_anterior,"/Results/ZG1_Permanentes_ISE_",nombres_meses[mes-3],"_",anio,".xlsx")

# Cargar el archivo de entrada
wb_ant_per <- loadWorkbook(mes_ant_per)

#Transitorio
mes_ant_tran=paste0(directorio,"/ISE/",anio,"/",carpeta_anterior,"/Results/ZG1_Transitorios_ISE_",nombres_meses[mes-3],"_",anio,".xlsx")

# Cargar el archivo de entrada
wb_ant_tran <- loadWorkbook(mes_ant_tran)


#Pecuario
mes_ant_pecu=paste0(directorio,"/ISE/",anio,"/",carpeta_anterior,"/Results/ZG2_Pecuario_ISE_",nombres_meses[mes-3],"_",anio,".xlsx")

# Cargar el archivo de entrada
wb_ant_pecu <- loadWorkbook(mes_ant_pecu)


trimestre=f_trimestre(mes)

#Tabaco
data_ant <- read.xlsx(wb_ant_tran, sheet = "Tabaco trimestral", colNames = TRUE,startRow = 10)
data_act <- read.xlsx(wb_act_tran, sheet = "Tabaco trimestral", colNames = TRUE,startRow = 10)

fila=which(data_ant$Año==anio & data_ant$Periodicidad==(trimestre-1))

tabla=data.frame(Producto="Tabaco",
                 Anterior=data_ant[fila,"Variación.Anual.Trimestral"],
                 Actual=data_act[fila,"Variación.Anual.Trimestral"]
)

#Algodon
data_ant <- read.xlsx(wb_ant_per, sheet = "Algodón Trimestral", colNames = TRUE,startRow = 11)
data_act <- read.xlsx(wb_act_per, sheet = "Algodón Trimestral", colNames = TRUE,startRow = 11)
fila=which(data_ant$Año==anio & data_ant$Periodicidad==(trimestre-1))

tabla=rbind(tabla,c("Algodon",data_ant[fila,"Variación.Anual.Trimestral"],data_act[fila,"Variación.Anual.Trimestral"]))

#Trigo
data_ant <- read.xlsx(wb_ant_tran, sheet = "Trigo trimestral", colNames = TRUE,startRow = 10)
data_act <- read.xlsx(wb_act_tran, sheet = "Trigo trimestral", colNames = TRUE,startRow = 10)

fila=which(data_ant$Año==anio & data_ant$Periodicidad==(trimestre-1))
tabla=rbind(tabla,c("Trigo",data_ant[fila,"Variación.Anual.Trimestral"],data_act[fila,"Variación.Anual.Trimestral"]))

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

#Frutas
data_ant <- read.xlsx(wb_ant_per, sheet = "Frutas Total(Expos+Interno)", colNames = TRUE,startRow = 11)
data_act <- read.xlsx(wb_act_per, sheet = "Frutas Total(Expos+Interno)", colNames = TRUE,startRow = 11)
fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-3))

tabla=rbind(tabla,c("Otras frutas",data_ant[fila,"total.trimestral"],data_act[fila,"total.trimestral"]))

#"Áreas en desarrollo"
data_ant <- read.xlsx(wb_ant_per, sheet = "Áreas en desarrollo", colNames = TRUE,startRow = 11)
data_act <- read.xlsx(wb_act_per, sheet = "Áreas en desarrollo", colNames = TRUE,startRow = 11)
fila=which(data_ant$Año==anio & data_ant$Periodicidad==(mes-3))

tabla=rbind(tabla,c("Areas en desarrolo",data_ant[fila,"X15"],data_act[fila,"X15"]))

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


nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta_actual,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
archivo=nombre_archivos[nombre_archivos$PRODUCTO=="EMMET","NOMBRE"]


carpeta=nombre_carpeta(mes,anio)
# Especifica la ruta del archivo de Excel
EMMET_ant <- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta_anterior,"/Data/consolidado_ISE/EMMET/",archivo),
                        sheet = "COMPLETO")
# Seleccionar solo las columnas que necesitas
EMMET_tabla_ant <- EMMET_ant[, c("anio", "mes", "Clase_CIIU4", "ProduccionRealPond")]

EMMET_act <- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/EMMET/",archivo),
                        sheet = "COMPLETO")
# Seleccionar solo las columnas que necesitas
EMMET_tabla_act <- EMMET_act[, c("anio", "mes", "Clase_CIIU4", "ProduccionRealPond")]
if(mes==3){
  vector=c(10,11,12)
  anio_act=anio-1
  anio_ant=anio-2
}else{
  vector=c(mes-5,mes-4,mes-3)
  anio_act=anio
  anio_ant=anio-1
}

Madera_tabla_ant=EMMET_tabla_ant %>%
  filter(Clase_CIIU4==1610 )%>%
  group_by(anio,mes)%>%
  summarise(suma=sum(ProduccionRealPond))%>%
  as.data.frame()


valor_ant=sum(Madera_tabla_ant[Madera_tabla_ant$anio==anio_act & Madera_tabla_ant$mes %in% vector,"suma"])/
  sum(Madera_tabla_ant[Madera_tabla_ant$anio==anio_ant & Madera_tabla_ant$mes %in% vector,"suma"])*100-100


Madera_tabla_act=EMMET_tabla_act %>%
  filter(Clase_CIIU4==1610 )%>%
  group_by(anio,mes)%>%
  summarise(suma=sum(ProduccionRealPond))%>%
  as.data.frame()

valor_act=sum(Madera_tabla_act[Madera_tabla_act$anio==anio_act & Madera_tabla_act$mes %in% vector,"suma"])/
sum(Madera_tabla_act[Madera_tabla_act$anio==anio_ant & Madera_tabla_act$mes %in% vector,"suma"])*100-100

tabla=rbind(tabla,c("Madera",valor_ant,valor_act))


#papel y carton



#Leña



#Pesca


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


}
