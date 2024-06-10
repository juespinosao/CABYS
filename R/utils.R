## utils
library(openxlsx)
library(dplyr)
#Vectores de meses
nombres_meses <- c("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
nombres_siglas <- c("Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic")
mes_0=c("01","02","03","04","05","06","07","08","09",10,11,12)
#Crear la funcion con el nombre de la carpeta segun año y mes
nombre_carpeta=function(mes,anio){
    carpeta=paste0(mes_0[mes],nombres_meses[mes])
}

vector_participaciones=data.frame(productos=c("Banano","Otras frutas","Plátano","Frutas cítricas","Flores",
                                              "Hortalizas","Yuca","Áreas en desarrollo","Legumbres verdes y secas",
                                              "Fruto de palma","Cacao","Caña panelera","Arroz","Papa","Caña de azúcar",
                                              "Maíz","Café pergamino","Cafetos","Ganado porcino","Ganado bovino",
                                              "Leche","Huevos","Aves de corral"))
# Asignar trimestre y semestre en base al número de mes
f_trimestre=function(mes){
trimestre <- case_when(
  mes %in% 1:3 ~ 1,
  mes %in% 4:6 ~ 2,
  mes %in% 7:9 ~ 3,
  mes %in% 10:12 ~ 4,
  TRUE ~ NA
)
return(trimestre)
}

f_trim_rom=function(mes){
  trimestre <- case_when(
    mes %in% 1:3 ~ "I",
    mes %in% 4:6 ~ "II",
    mes %in% 7:9 ~ "III",
    mes %in% 10:12 ~ "IV",
    TRUE ~ NA
  )
  return(trimestre)
}

f_trim_nombre=function(mes){
  trimestre <- case_when(
    mes %in% 1:3 ~ "Primer",
    mes %in% 4:6 ~ "Segundo",
    mes %in% 7:9 ~ "Tercero",
    mes %in% 10:12 ~ "Cuarto",
    TRUE ~ NA
  )
  return(trimestre)
}

f_semestre=function(mes){
semestre <- case_when(
  mes %in% 1:6 ~ 1,
  mes %in% 7:12 ~ 2,
  TRUE ~ NA
)
return(semestre)}

f_semestre_nombre=function(mes){
  semestre <- case_when(
    mes %in% 1:6 ~ "Primer",
    mes %in% 7:12 ~ "Segundo",
    TRUE ~ NA
  )
  return(semestre)}


# formatos de celdas pecuario
col1<- createStyle(
  fontName = "Arial Narrow",
  fontSize = 11,
  fontColour = "black",  # Puedes ajustar el color de fuente según tu preferencia
  border = "Top: thin, Left: thin, Right: thin",
  borderColour = "#A6A6A6",
  halign = "center"
)



col2 <- createStyle(
  fontName = "Arial Narrow",
  fontSize = 11,
  fontColour = "black",
  border = "Top: thin, Left: thin, Right: thin",
  borderColour = "#A6A6A6",
  numFmt = "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)"
)


col3 <- createStyle(
  fontName = "Arial Narrow",
  fontSize = 11,
  fontColour = "black",
  border = "Top: thin, Left: thin, Right: thin",
  borderColour = "#A6A6A6",
  numFmt = "0.0",
  halign = "center",
  valign = "center"
)

##un decimal sin centrar
col4 <- createStyle(
  fontName = "Arial Narrow",
  fontSize = 11,
  fontColour = "black",
  border = "Top: thin, Left: thin, Right: thin",
  borderColour = "#A6A6A6",
  numFmt = "0.0"
)


col5 <- createStyle(
  fontName = "Arial Narrow",
  fontSize = 11,
  fontColour = "black",
  border = "Top: thin, Left: thin, Right: thin",
  borderColour = "#A6A6A6",
  numFmt = "0.00",
  halign = "center",
  valign = "center"
)

#formatos de celda transitorios
#maiz

colmb <- createStyle(
  fontName = "Arial Narrow",
  fontSize = 11,
  fontColour = "black",
  border = "Top: thin, Left: thin, Right: thin",
  borderColour = "#A6A6A6",
  numFmt = "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)",
  fgFill = "#B7DEE8"
)

colma <- createStyle(
  fontName = "Arial Narrow",
  fontSize = 11,
  fontColour = "black",
  border = "Top: thin, Left: thin, Right: thin",
  borderColour = "#A6A6A6",
  numFmt = "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)",
  fgFill = "#CCC0DA"
)

##solo un decimal
col6 <- createStyle(
  fontName = "Arial Narrow",
  fontSize = 11,
  fontColour = "black",
  border = "Top: thin, Left: thin, Right: thin",
  borderColour = "#A6A6A6",
  numFmt = "_(* #,##0.0_);_(* \\(#,##0.0\\);_(* \"-\"??_);_(@_)"
)

##sin decimales, personalizado, sin centrar
col7 <- createStyle(
  fontName = "Arial Narrow",
  fontSize = 11,
  fontColour = "black",
  border = "Top: thin, Left: thin, Right: thin",
  borderColour = "#A6A6A6",
  numFmt = "_(* #,##0_);_(* \\(#,##0\\);_(* \"-\"??_);_(@_)"
)



## numero, sin decimales,centrado

col8 <- createStyle(
  fontName = "Arial Narrow",
  fontSize = 11,
  fontColour = "black",
  border = "Top: thin, Left: thin, Right: thin",
  borderColour = "#A6A6A6",
  numFmt = "0",
  halign = "center",
  valign = "center"
)

## numero, sin decimales,derecha
col9 <- createStyle(
  fontName = "Arial Narrow",
  fontSize = 11,
  fontColour = "black",
  border = "Top: thin, Left: thin, Right: thin",
  borderColour = "#A6A6A6",
  numFmt = "0"
)

col10 <- createStyle(
  fontName = "Arial Narrow",
  fontSize = 11,
  fontColour = "black",
  border = "Top: thin, Left: thin, Right: thin",
  borderColour = "#A6A6A6",
  numFmt = "_-* #,##0.0_-;-* #,##0.0_-;_-* \"-\"_-;_-@_-"
)


# Formatos Reporte coyuntura ----------------------------------------------


cbp <- createStyle(
  fontName = "Arial",
  fontSize = 11,
  fontColour = "black",
  border = "Top: thin, Left: thin, Right: thin",
  borderColour = "#A6A6A6",
  numFmt = "mmm-yy"
)
cbn <- createStyle(
  fontName = "Arial",
  fontSize = 11,
  fontColour = "black",
  numFmt = "_(* #,##0_);_(* \\(#,##0\\);_(* \"-\"??_);_(@_)"
)

##un decimal sin centrar
rn4 <- createStyle(
  fontName = "Arial",
  fontSize = 11,
  fontColour = "black",
  numFmt = "0.0"
)

##dos decimal sin centrar
rn5 <- createStyle(
  fontName = "Arial",
  fontSize = 11,
  fontColour = "black",
  numFmt = "0.00"
)

cbn2 <- createStyle(
  fontName = "Arial",
  fontSize = 11,
  fontColour = "black",
  numFmt = "_-* #,##0.0_-;-* #,##0.0_-;_-* \"-\"_-;_-@_-"
)

cbn3 <- createStyle(
  fontName = "Arial",
  fontSize = 11,
  fontColour = "black",
  numFmt = "_-* #,##0.00_-;-* #,##0.00_-;_-* \"-\"_-;_-@_-"
)

conta <- createStyle(
  fontName = "Arial",
  fontSize = 11,
  fontColour = "black",
  numFmt = "ACCOUNTING"
)


cbn4 <- createStyle(
  fontName = "Arial",
  fontSize = 11,
  fontColour = "black",
  numFmt = "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)"
)



