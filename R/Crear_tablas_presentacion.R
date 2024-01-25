library(gt)
library(dplyr)
library(gtExtras)
a <- rnorm(21, mean = 112, sd =12)
colour <- rep(c("Blue", "Red", "Green"), 7)

data <- data.frame(colour, a) %>%
  group_by(colour) %>%
  summarise(mean = mean(a), sd = sd(a), n = n()) %>%
  mutate(grp = "colour") %>%
  rename(cat = colour)

b <- rnorm(21, mean = 60, sd =12)
day <- rep(c("2", "4", "6"), 7)

data2 <- data.frame(day, b) %>%
  group_by(day) %>%
  summarise(mean = mean(a), sd = sd(a), n = n()) %>%
  mutate(grp = "day") %>%
  rename(cat =  day)

bind_rows(data, data2) %>%
  group_by(grp) %>%
  gt(rowname_col = "cat")


bind_rows(data, data2) %>%
  group_by(grp) %>%
  gt() %>%
  tab_options(row_group.as_column = TRUE) %>%
  tab_style(style = list(cell_text(v_align = "bottom")),locations = cells_body(columns=grp))






tabla <- bind_rows(data, data2) %>%
  group_by(grp) %>%
  gt() %>%
  tab_options(row_group.as_column = TRUE) %>%
  tab_spanner(
    label = "grp",
    columns = vars(grp),
    align = "center"
  )





### Prueba

Periodo=c("Enero","Febrero","Marzo","Abril")
Ingresos=c(32.8,30.6,21.3,23.2)
personal=c(17.7,17.6,16.4,12.7)
salario=c(21.8,18.3,20.6,18.6)
tabla1=data.frame(Periodo,Ingresos,personal,salario)




tabla1 %>%
  dplyr::select(Periodo, Ingresos, personal, salario) %>%
  dplyr::mutate(Ingresos_scaled = Ingresos/max(Ingresos) * 100) %>%
  dplyr::mutate(personal_scaled = personal/max(personal) * 100) %>%
  dplyr::mutate(salario_scaled = salario/max(salario) * 100) %>%
  gt() %>%
  cols_move(Ingresos_scaled, Ingresos) %>% #mover esa columna, despues de
  cols_move(personal_scaled, personal) %>%
  cols_move(salario_scaled, salario) %>%
  gt::tab_spanner(
    label = "Variación %",
    columns = c(Ingresos, Ingresos_scaled, personal, personal_scaled, salario, salario_scaled),
  ) %>%
  gt::cols_label(
    Periodo = "Periodo que pasa si se pone un titulo grande y se mueve",
    Ingresos = "Ingresos Nominales",
    personal = "Personal Ocupado",
    salario = "Salario Nominal",
    Ingresos_scaled = "Ingresos Nominales Escalados",
    personal_scaled = "Personal Ocupado Escalado",
    salario_scaled = "Salario Nominal Escalado"
  ) %>%
  gt_plt_bar_pct(column = Ingresos_scaled, scaled = TRUE, fill = "#FF7BB0", background = "white") %>%
  gt_plt_bar_pct(column = personal_scaled, scaled = TRUE, fill = "#D9D9D9", background = "white") %>%
  gt_plt_bar_pct(column = salario_scaled, scaled = TRUE, fill = "#7F7F7F", background = "white") %>%
  cols_align("center", contains("scale")) %>%
  cols_width(everything() ~ px(120)) %>%

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
  gt_highlight_rows(
    rows = seq(1, 4, by = 2), #Cambiar el 4 por el número de filas real
    # Background color
    fill = "#F2F2F2",
    font_weight = "normal"
    # Select target column
    #target_col = c(Periodo, Ingresos, Ingresos_scaled, personal, personal_scaled, salario, salario_scaled)
  ) %>%
  gtsave(filename = "salidas/graficas/grafico_1.png", expand = 10)



# Actualizacion mensual ---------------------------------------------------

if(mes==1){
  carpeta_anterior=nombre_carpeta(12,(anio-1))
}else{
  carpeta_anterior=nombre_carpeta(mes-1,anio)
}

carpeta_actual=nombre_carpeta(mes,anio)

#Permanentes
mes_ant_per=paste0(directorio,"/",anio,"/",carpeta_anterior,"/consolidado_ISE/ZG1_Permanentes_ISE_",nombres_meses[mes-1],"_",anio,".xlsx")
mes_act_per=paste0(directorio,"/",anio,"/",carpeta_actual,"/consolidado_ISE/ZG1_Permanentes_ISE_",nombres_meses[mes],"_",anio,".xlsx")

# Cargar el archivo de entrada
wb_ant_per <- loadWorkbook(mes_ant_per)
wb_act_per <- loadWorkbook(mes_act_per)

#Transitorio
mes_ant_tran=paste0(directorio,"/",anio,"/",carpeta_anterior,"/consolidado_ISE/ZG1_Transitorios_ISE_",nombres_meses[mes-1],"_",anio,".xlsx")
mes_act_tran=paste0(directorio,"/",anio,"/",carpeta_actual,"/consolidado_ISE/ZG1_Transitorios_ISE_",nombres_meses[mes],"_",anio,".xlsx")

# Cargar el archivo de entrada
wb_ant_tran <- loadWorkbook(mes_ant_tran)
wb_act_tran <- loadWorkbook(mes_act_tran)


#Pecuario
mes_ant_pecu=paste0(directorio,"/",anio,"/",carpeta_anterior,"/consolidado_ISE/ZG2_Pecuario_ISE_",nombres_meses[mes-1],"_",anio,".xlsx")
mes_act_pecu=paste0(directorio,"/",anio,"/",carpeta_actual,"/consolidado_ISE/ZG2_Pecuario_ISE_",nombres_meses[mes],"_",anio,".xlsx")

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

tabla=rbind(tabla,c("Panela",data_ant[fila,"Variacion.Anual"],data_act[fila,"Variacion.Anual"]))

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
tabla$Diferencia=tabla[,"Anterior"]-tabla[,"Actual"]

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
  gtsave(filename = paste0(directorio,"/",anio,"/",carpeta_actual,"/Coyuntura ",nombres_meses[mes],"/prueba.png"), expand = 10)

# Actualizacion trimestral ---------------------------------------------------
if(mes==3){
  carpeta_anterior=nombre_carpeta(12,(anio-1))
}else{
  carpeta_anterior=nombre_carpeta(mes-3,anio)
}
#Permanentes
mes_ant_per=paste0(directorio,"/",anio,"/",carpeta_anterior,"/consolidado_ISE/ZG1_Permanentes_ISE_",nombres_meses[mes-3],"_",anio,".xlsx")

# Cargar el archivo de entrada
wb_ant_per <- loadWorkbook(mes_ant_per)

#Transitorio
mes_ant_tran=paste0(directorio,"/",anio,"/",carpeta_anterior,"/consolidado_ISE/ZG1_Transitorios_ISE_",nombres_meses[mes-3],"_",anio,".xlsx")

# Cargar el archivo de entrada
wb_ant_tran <- loadWorkbook(mes_ant_tran)


#Pecuario
mes_ant_pecu=paste0(directorio,"/",anio,"/",carpeta_anterior,"/consolidado_ISE/ZG2_Pecuario_ISE_",nombres_meses[mes-3],"_",anio,".xlsx")

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
tabla=rbind(tabla,c("Arroz",data_ant[fila,"Estado"],data_act[fila,"Estado"]))

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

