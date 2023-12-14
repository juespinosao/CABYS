

f_Leche_polvo<-function(directorio,mes,anio){


  library(readxl)
  library(dplyr)


  carpeta=nombre_carpeta(mes,anio)
  # Especifica la ruta del archivo de Excel

  # Leche_sipsa -------------------------------------------------------------


  # Especifica la ruta del archivo de Excel
  Leche <- read_excel(paste0(directorio,"/",anio,"/",carpeta,"/consolidado_ISE/Leche/SIPSA/LECHE_CRUDA_EST_",nombres_meses[mes],"_",anio,".xlsx"),
                      sheet = "LecheDANE")

  Valor_Leche=as.data.frame(Leche[Leche[,"Año"] == anio | Leche[,"Año"] == (anio-1),"PRODUCCION LECHE CRUDA DANE"])
  Valor_Leche=as.numeric(Valor_Leche$`PRODUCCION LECHE CRUDA DANE`)



# Leche_cruda -------------------------------------------------------------

  Leche_cruda <- read.xlsx(paste0(directorio,"/",anio,"/",carpeta,"/consolidado_ISE/Leche/USP/Volumen de Acopio Total - Res 0017 de 2012.xlsx"),sheet = "VOLUMEN NACIONAL")
  fila_fuente=which(grepl("Fuente",as.data.frame(t(Leche_cruda))))
  col_per=which(Leche_cruda=="Periodo",arr.ind = TRUE)[,"col"]
  col_vol=which(Leche_cruda=="Volumen (lt)",arr.ind = TRUE)[,"col"]
  Leche_cruda[,col_per[1]]=as.numeric(Leche_cruda[,col_per[1]])
  Leche_cruda[,col_per[1]] <- as.Date(Leche_cruda[,col_per[1]], origin = "1899-12-30")
  Leche_cruda[,col_per[1]]=format(Leche_cruda[,col_per[1]], "%Y-%m")
  fila1=which(Leche_cruda==paste0(anio-1,"-01"))
  Valor_cruda=as.numeric(Leche_cruda[fila1:(fila_fuente[length(fila_fuente)]-1),col_vol[1]])





# leche en polvo ----------------------------------------------------------



  Leche_polvo <- read.xlsx(paste0(directorio,"/",anio,"/",carpeta,"/consolidado_ISE/Leche/USP/Inventarios Cuadro principal.xlsx"),
                     sheet = "INVENTARIOS CUADRO PRINCIPAL",startRow = 4)

  fila_fuente=which(grepl("Fuente",as.data.frame(t(Leche_polvo))))

  fila1=which(grepl(anio-1,Leche_polvo$Fecha),arr.ind = TRUE)
  fila2=which(grepl("Enero",Leche_polvo$Fecha),arr.ind = TRUE)
  filaf <- intersect(fila1, fila2)
  columna=which(Leche_polvo=="Leche en Polvo Entera (Tn)",arr.ind = TRUE)[,"col"]
  Valor_polvo=as.numeric(Leche_polvo[filaf:(fila_fuente[1]-1),columna[1]])





  longitud_maxima <- max(length(Valor_Leche), length(Valor_cruda), length(Valor_polvo))
  Valor_Leche <- c(Valor_Leche, rep(NA, longitud_maxima - length(Valor_Leche)))
  Valor_cruda <- c(Valor_cruda, rep(NA, longitud_maxima - length(Valor_cruda)))
  Valor_polvo <- c(Valor_polvo, rep(NA, longitud_maxima - length(Valor_polvo)))

  leche_final=as.data.frame(cbind(Valor_Leche,Valor_cruda,Valor_polvo))





  # Importaciones -----------------------------------------------------------

  archivos=list.files(paste0(directorio,"/",anio,"/",carpeta,"/consolidado_ISE"))
  elementos_seleccionados <- archivos[grepl("Expos e", archivos) ]


  dos_digitos <- anio %% 100
  # Especifica la ruta del archivo de Excel
  Impor <- read.xlsx(paste0(directorio,"/",anio,"/",carpeta,"/consolidado_ISE/",elementos_seleccionados,"/Resumen Importaciones ",mes_0[mes],"_",dos_digitos," - copia.xlsx"),
                     sheet = "PNK")

  n_col1=which(Impor== paste0((anio)," 01"),arr.ind = TRUE)[,"col"]
  n_col2=which(Impor== paste0(anio," ",mes_0[mes]),arr.ind = TRUE)[,"col"]
  n_fila1=which(Impor== "220201",arr.ind = TRUE)[,"row"]


  #Tomar el valor que nos interesa
  Valor_importaciones=as.data.frame(t(Impor[n_fila1,(n_col1[1]:n_col2[1])]))
  Valor_importaciones=as.data.frame(lapply(Valor_importaciones, as.numeric))





  # Exportaciones -----------------------------------------------------------



  # Especifica la ruta del archivo de Excel
  # Especifica la ruta del archivo de Excel
  Expor <- read.xlsx(paste0(directorio,"/",anio,"/",carpeta,"/consolidado_ISE/",elementos_seleccionados,"/Resumen Exportaciones ",mes_0[mes],"-",anio," - copia.xlsx"),
                     sheet = "PNK")


  n_fila1=which(Expor == "220100",arr.ind = TRUE)[,"row"]
  n_col1=which(Expor== paste0(anio," 1"),arr.ind = TRUE)[,"col"]
  n_col2=which(Expor== paste0(anio," ",mes),arr.ind = TRUE)[,"col"]

  n_fila2=which(Expor == "220201",arr.ind = TRUE)[,"row"]



  #Tomar el valor que nos interesa
  Valor_exportaciones=as.data.frame(t(Expor[c(n_fila1,n_fila2),(n_col1[1]:n_col2[1])]))
  Valor_exportaciones=as.data.frame(lapply(Valor_exportaciones, as.numeric))

  Expo_impo=cbind(Valor_importaciones,Valor_exportaciones)
  colnames(Expo_impo)=c("Importacion","Pasterizada","Crema")
  return(list(leche_final,Expo_impo))
}
