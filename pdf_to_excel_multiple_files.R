
##Enter the folder address with the files in .pdf on line 22
##replacing the single bars "" with inverted bars '/'
##Press ctrl+shift+enter to run the code


packages <- c("pdftools", "stringr", "openxlsx")


install.packages(setdiff(packages, rownames(installed.packages())))

setwd("[folder adress]")
library(pdftools)

library(stringr)
library(openxlsx)

## Enter the folder address with the files in .pdf on line 22

setwd("[folder adress]")
  
file.list <- list.files(pattern=c('*.pdf','*.PDF' ))

file.list

## see if the files are correcly listed

rm(base)
rm(base_nomes)
i = 1

for (i in 1:length(file.list)) {
  
  tx <- pdf_text(file.list[i])
  tx2 <- unlist(str_split(tx, "[\\r\\n]+"))
  tx3 <- str_split_fixed(str_trim(tx2), "\\s{2,}", 5)
  nomes <-  rep(file.list[i], nrow(tx3))
  if(i == 1){
    base <- tx3
    base_nomes <- as.data.frame(nomes)
  }
   if(i > 1) {
     
     base <- rbind(base, tx3)
     base_nomes <- rbind(base_nomes, as.data.frame(nomes))
   }

  
  }



library(openxlsx)

wb <- createWorkbook()
# 
addWorksheet(wb, "folder_infos")
writeData(wb, "folder_infos", base, rowNames = TRUE, sep = ',')
addWorksheet(wb, "folder_name")
writeData(wb, "folder_name", base_nomes, rowNames = TRUE, sep = ',')
saveWorkbook(wb, "pdftoexcel.xlsx", overwrite = TRUE)



 