#Set up stuff
library("xlsx")
library("stringr")
setwd("C:/Users/joychiu/Documents/Coding in R/Git example/")

#Function paredown for clearing extra text
paredown=function(string){
  pared=word(string, -1)
  return(as.character(pared))
}

#Function to get data from the Excel doc
get_data=function(columns, rows){
  df=read.xlsx(file="Neurodegeneration Gene Expression.xlsx",
               sheetName="Genes in ALS, FTD, PD",
               colIndex=columns,
               rowIndex=rows)
  return(df)
}

#Function that uses the "paredown" fcn to make a table
make_table=function(columns, rows){
  table=get_data(columns, rows) #gets the numerical data and column titles
  table=apply(table, c(1:2), paredown) #applies paredown to all cells
  return(table)
}

#DO THE PARSING!
newdata=make_table(1, 1:327)

#Export to new Excel file sheet, appended to same document
write.xlsx(newdata,
           file="Neurodegeneration Gene Expression.xlsx",
           sheetName="pared down list",
           append=TRUE
           )

#hello