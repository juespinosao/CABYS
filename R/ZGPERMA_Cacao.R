#' @export
# Cacao
# Cargar la biblioteca readxl

f_Cacao<-function(directorio,mes,anio){

  #Cargar librerias
  library(readxl)
  library(dplyr)
  library(openxlsx)
  library(zoo)

  #identificar la carpeta
  carpeta=nombre_carpeta(mes,anio)



  # STOCKS cafe verde ------------------------------------------------------------------
  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Cacao","NOMBRE"]


  # Especifica la ruta del archivo de Excel
  Cacao <- read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/Cacao/",archivo))


  fila_mes=which(Cacao == "MES",arr.ind = TRUE)[, "row"][1]

  #si which es 0 entonces generar error o algo

  #identificar las columna
  columna2a=which(grepl((anio-2),Cacao[fila_mes,]),arr.ind = TRUE)

  columna0a=which(grepl(anio,Cacao[fila_mes,]),arr.ind = TRUE)
  n_fila=which(Cacao=="Enero",arr.ind = TRUE)[,"row"][1]



  #Tomar el valor que nos interesa
  Valor_Cacao=Cacao[n_fila:(n_fila+11),columna2a:columna0a]
  Valor_Cacao=Valor_Cacao %>%
    mutate(across(everything(), as.numeric)) %>%
    as.data.frame()
  Valor_Cacao=c(Valor_Cacao[,1],Valor_Cacao[,2],Valor_Cacao[,3])
  return(Valor_Cacao)
}
