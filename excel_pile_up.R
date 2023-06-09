##Enter the folder address with the files in .pdf on line 22
##replacing the single bars "" with inverted bars '/'
##Press ctrl+shift+enter to run the code

packages <- c("readxl", "stringr", "openxlsx", "dplyr")


install.packages(setdiff(packages, rownames(installed.packages())))


library(dplyr)
library(stringr)
library(openxlsx)
library(readxl)

## Insert the folder adress
getwd()
setwd("[folder adress]")
readx

## list all files in folder. 
file.list <- list.files(pattern = '.xlsx')
file.list
## pile it up
df.list <- lapply(file.list, read_xlsx)
names(df.list) <- file.list
unlist(df.list)
a <- bind_rows(df.list, .id = "column_label")



## Salvamento da base Empilhada

wb <- createWorkbook()
# 
addWorksheet(wb, "Files pilled up")
writeData(wb, "Files pilled up", a, rowNames = TRUE, sep = ".")
saveWorkbook(wb, "excel_piledup.xlsx", overwrite = TRUE)

