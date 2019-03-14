# Trabajo Final 

#install.packages("openxlsx")


library (openxlsx)

setwd ("C:/Users/emico/Dropbox/simulacion de modelos economicos/trabajo final") #windows

setwd ("/home/emiliano/Escritorio/Dropbox/simulacion de modelos economicos/trabajo final") #linux

WB= loadWorkbook("MIP2005.xlsx") # hoja con las MIP 2005

MIP_2005 <- as.matrix (read.xlsx( WB,sheet="Sintesis 40",rows=c(3:42),cols=c(3:42), colNames=F , rowNames = F)) #Importando la hoja

I_2005= diag(1,40) #matriz identidad para MIP 2005 

valor_bruto_2005= read.xlsx( WB,sheet="Sintesis 40",rows=c(53),cols=c(3:42), colNames=F , rowNames = F) # VB por sector


A_2005= MIP_2005 %*% diag ( 1/valor_bruto_2005)  #matriz de coeficientes tecnicos


X_2005= (solve (I_2005-A_2005) ) # matriz de requerimientos directos e indirectos

sectores_2005=  as.character (read.xlsx( WB ,sheet="Sintesis 40",rows=c(2),cols=c(3:42), colNames=F , rowNames = F ) )#nombre de los sectores


colnames(X_2005)= sectores_2005 #sector en columna matriz X
row.names(X_2005)= sectores_2005 #sector en fila  para matriz X

colnames (A_2005)= sectores_2005 # sector en columnsa para  matriz A
rownames (A_2005) = sectores_2005 #sector en columna para matriz B

WB_2= loadWorkbook("MIP_97_05.xlsx") #hoja de calculo donde esta registrada la MIP 1997

addWorksheet(WB_2, "2005_coef_tec") #agregando hoja adicional para la matriz  de coeficientes tecnicos del 2005

addWorksheet(WB_2, "2005_req_tec") #agregando hoja adicional para la matriz de req tecnicos



writeData(WB_2, sheet = "2005_req_tec", X_2005, colNames = T , rowNames = T, startCol = 2)#exportando matriz X
writeData(WB_2, sheet = "2005_coef_tec", A_2005, colNames =T  , rowNames = T, startRow =2 )

saveWorkbook(WB_2 , "MIP_97_05.xlsx", overwrite = T )



