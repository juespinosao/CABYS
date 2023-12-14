# Leche

directorio="C:/Users/Asus/Desktop/Dane/proyecto2/Automatizacion CABYS/Automatizacion/Formato_carpetas"
mes=7
anio=2023
f_Leche<-function(directorio,mes,anio){
  
  
  library(readxl)
  library(dplyr)
  
  
  carpeta=nombre_carpeta(mes,anio)
  

# Leche_sipsa -------------------------------------------------------------

  
  # Especifica la ruta del archivo de Excel
  Leche <- read_excel(paste0(directorio,"/",anio,"/",carpeta,"/consolidado_ISE/Leche/SIPSA/LECHE_CRUDA_EST_",nombres_meses[mes],"_",anio,".xlsx"),
                       sheet = "LecheDANE")

  Valor_Leche=as.data.frame(Leche[Leche[,"Año"] == anio,"PRODUCCION LECHE CRUDA DANE"])

  
  
# Leche cruda  
  
  
  
  
    
  return(as.numeric(Valor_Leche[,1]))
}
