# -----------------------------------------------------------------------------
# (c) Observatorio del Mercado de Trabajo. Junta de Comunidades de Castilla-La Mancha
# (c) Isidro Hidalgo Arellano (ihidalgo@jccm.es)
# -----------------------------------------------------------------------------

# Borramos los objetos en memoria
rm(list = ls(all.names = TRUE))

# Cargamos las librerías necesarias
library(pxR)
library(openxlsx)

# Leemos el fichero
datos = read.px(file.choose())$DATA[[1]] # Tarda un rato...

# A partir de 2020 se incluye el campo "periodo", pero hay que revisar porque
# el INE graba los ficheros de forma diferente a años anteriores
names(datos)
names(datos) = c("periodo", "edad", "municipio", "sexo", "valor") # Modificar según
# estructura, porque cambian cada año...

# Cambiamos las Variables globales según los valores del año en el fichero:
anyo <- "2020"

levels(datos$edad)
etiquetaTotalEdades <- "Todas las edades" # Otros años lo especifican como "Total"
etiquetaTotalEdadesPrimera <- str_split_fixed(etiquetaTotalEdades, " ", Inf)[,1]
etiqueta100yMas <- "100" # si está como "100 y más años" no hace falta
anyos1664 <- paste(16:64, "años") # Si la etiqueta de años es de la forma "16 años"...
# Si no, cambiar como corresponda
anyos1665 <- paste(16:65, "años") # Si la etiqueta de años es de la forma "16 años"...
# Si no, cambiar como corresponda

levels(datos$sexo)
etiquetaAmbosSexos <- "Total" # Otros años lo especifican como "Ambos sexos"

# Nos quedamos sólo con los datos del último año
datos <- datos[datos$periodo == anyo,]
datos$periodo = NULL

# Extraemos las poblaciones totales de España
ESP1664 = sum(datos[datos$municipio == "Total Nacional" & datos$sexo == "Total" & datos$edad %in% anyos1664, "valor"])
ESP1665 = sum(datos[datos$municipio == "Total Nacional" & datos$sexo == "Total" & datos$edad %in% anyos1665, "valor"])

# Montamos el dataframe con los datos de España
# salidaESP = dcast(datos, sexo + municipio ~ edad, valor.var = "valor")

# Seleccionamos los datos de Castilla-La Mancha
CLM = datos$municipio[!is.na(str_extract(datos$municipio, '^02|^13|^16|^19|^45'))]
table(substr(CLM,1,2)) # Comprobamos que no hay registros de otras provincias
datosCLM = datos[datos$municipio %in% CLM,]
datosCLM$municipio = factor( datosCLM$municipio)
# levels(datosCLM$municipio)

# Guardamos los datos en bruto para usarlos en tablas dinámicas de Excel
brutoCLM = datosCLM
brutoCLM$edad = as.character(brutoCLM$edad)
brutoCLM$edad = str_split_fixed(brutoCLM$edad, " ", Inf)[,1]
brutoCLM= brutoCLM[brutoCLM$edad != etiquetaTotalEdadesPrimera,]
brutoCLM$edad = as.numeric(brutoCLM$edad)
brutoCLM$valor[is.na(brutoCLM$valor)] <-0
brutoCLM$grupoEdad[brutoCLM$edad < 16] = "0-15"
brutoCLM$grupoEdad[brutoCLM$edad > 15 & brutoCLM$edad < 45] = "16-44"
brutoCLM$grupoEdad[brutoCLM$edad > 44 & brutoCLM$edad < 65] = "45-64"
brutoCLM$grupoEdad[brutoCLM$edad == 65] = "65"
brutoCLM$grupoEdad[brutoCLM$edad > 65] = "66+"

brutoCLM$provincia = substr(brutoCLM$municipio, 1, 2)
brutoCLM$INE = substr(brutoCLM$municipio, 1, 5)
brutoCLM$municipio = substr(brutoCLM$municipio, 7, 200)
columnas = c("edad", "grupoEdad", "provincia", "INE", "municipio", "sexo", "valor")
brutoCLM = brutoCLM[, columnas]

# Extraemos las poblaciones totales de Castilla-La Mancha
CLM1664 = sum(brutoCLM[brutoCLM$sexo == etiquetaAmbosSexos & brutoCLM$edad %in% 16:64, "valor"])
CLM1665 = sum(brutoCLM[brutoCLM$sexo == etiquetaAmbosSexos & brutoCLM$edad %in% 16:65, "valor"])

# Montamos los dataframe con los datos de CLM
CLM = dcast(brutoCLM, sexo + provincia + municipio + INE ~ edad, valor.var = "valor")
factores <- sapply(CLM, is.factor)
CLM[factores] <- lapply(CLM[factores], as.character)
CLM$Total <- apply(CLM[,as.character(0:100)],1 , sum)
columnas = c("sexo", "INE", "provincia", "municipio", "Total", 0:100)
salidaCLM = CLM[order(CLM$sexo, CLM$INE), columnas]
names(salidaCLM) <- c("sexo", "INE", "provincia", "municipio", "Total", 0:99, "100 y más")

salidaAB = salidaCLM[!is.na(str_extract(salidaCLM$INE,'^02')), ]
salidaCR = salidaCLM[!is.na(str_extract(salidaCLM$INE,'^13')), ]
salidaCU = salidaCLM[!is.na(str_extract(salidaCLM$INE,'^16')), ]
salidaGU = salidaCLM[!is.na(str_extract(salidaCLM$INE,'^19')), ]
salidaTO = salidaCLM[!is.na(str_extract(salidaCLM$INE,'^45')), ]

salidaCLM166465 = data.frame(INE = as.numeric(salidaCLM$INE[salidaCLM$sexo == etiquetaAmbosSexos]),
                             Pobl1664 = rowSums(salidaCLM[salidaCLM$sexo == etiquetaAmbosSexos, colnames(salidaCLM) %in% as.character(16:64)]),
                             Pobl1665 = rowSums(salidaCLM[salidaCLM$sexo == etiquetaAmbosSexos, colnames(salidaCLM) %in% as.character(16:65)]))
salidaCLM166465 = salidaCLM166465[order(salidaCLM166465$INE), ]

salidaCLM166465Hombres = data.frame(INE = as.numeric(salidaCLM$INE[salidaCLM$sexo == "Hombres"]),
                             Pobl1664 = rowSums(salidaCLM[salidaCLM$sexo == "Hombres", colnames(salidaCLM) %in% as.character(16:64)]),
                             Pobl1665 = rowSums(salidaCLM[salidaCLM$sexo == "Hombres", colnames(salidaCLM) %in% as.character(16:65)]))
salidaCLM166465Hombres = salidaCLM166465Hombres[order(salidaCLM166465Hombres$INE), ]

salidaCLM166465Mujeres = data.frame(INE = as.numeric(salidaCLM$INE[salidaCLM$sexo == "Mujeres"]),
                             Pobl1664 = rowSums(salidaCLM[salidaCLM$sexo == "Mujeres", colnames(salidaCLM) %in% as.character(16:64)]),
                             Pobl1665 = rowSums(salidaCLM[salidaCLM$sexo == "Mujeres", colnames(salidaCLM) %in% as.character(16:65)]))
salidaCLM166465Mujeres = salidaCLM166465Mujeres[order(salidaCLM166465Mujeres$INE), ]

# Grabamos los datos en un fichero Excel
TOTALES = data.frame(Total = c("ESP1664", "ESP1665", "CLM1664", "CLM1665"),
                     Valor = c(ESP1664, ESP1665, CLM1664, CLM1665))
lista <- list("En bruto" = brutoCLM, "CLM" = salidaCLM, "AB" = salidaAB, "CR" = salidaCR, "CU" = salidaCU, "GU" = salidaGU, "TO" = salidaTO, "TOTALES" = TOTALES,
              "CLM166465" = salidaCLM166465, "CLM166465Hombres" = salidaCLM166465Hombres, "CLM166465mujeres" = salidaCLM166465Mujeres)
write.xlsx(lista, file = "//jclm.es/PROS/SC/OBSERVATORIOEMPLEO/_Extractores/Población.xlsx")
print("El fichero resultante está en: //jclm.es/PROS/SC/OBSERVATORIOEMPLEO/_Extractores/Población.xlsx")
