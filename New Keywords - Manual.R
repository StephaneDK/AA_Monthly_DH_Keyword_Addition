.libPaths()

# Sys.getenv("PATH")
# Sys.getenv("HOME")

#Sys.setenv(PATH = paste("C:/rtools40/usr/bin", Sys.getenv("PATH"), sep=";"))
#write('PATH="${RTOOLS40_HOME}\\usr\\bin;${PATH}"', file = "~/.Renviron", append = TRUE)
#Sys.setenv(PATH = paste("C:/Rtools/bin", Sys.getenv("PATH"), sep=";"))



cat("Keyword addition start\n")

#oldw <- getOption("warn")
#options(warn = -1, rlib_downstream_check = FALSE)

library(tidyverse)
library(stringr)
library(reshape2)
library(ggthemes)
library(gridExtra)
library(forecast)
library(aTSA)
library(DescTools)
library(Rcpp)
library(plyr)
library(EnvStats)
library(qcc)
library(openxlsx)
library(magrittr)



#options(warn = oldw)

options(scipen=999, digits = 3 )

`%notin%` <- Negate(`%in%`)

convert.brackets <- function(x){
  if(grepl("\\(.*\\)", x)){
    paste0("-", gsub("\\(|\\)", "", x))
  } else {
    x
  }
}


#Setting the directory where all files will be used from for this project
#setwd("C:\\Users\\steph\\Documents\\DK\\Work\\Forecasting book sales and inventory\\Pipeline\\csv")
#setwd("C:\\Users\\snichanian\\Documents\\DK\\Work\\Forecasting book sales and inventory\\Auxilary tasks\\Amazon Advertisment\\Keyword Addition\\Monthly\\csv")


Volume <- read.csv("Volume.csv", header = T, stringsAsFactors = FALSE)
Impression <- read.csv("Impression.csv", header = T, stringsAsFactors = FALSE)
datahawk <- read.csv("datahawk.csv", header = T, stringsAsFactors = FALSE, encoding ="UTF-8")


#Importing and clean Sales Data
Sales_DE <- read.csv("DE_AMZ.csv ", header = T, skip = 1, stringsAsFactors = FALSE, encoding ="UTF-8")
Sales_DE$REGION <- "DE"
Sales_UK <- read.csv("UK_AMZ.csv", header = T, skip = 1, stringsAsFactors = FALSE, encoding ="UTF-8")
Sales_UK$REGION <- "UK"
Sales_US <- read.csv("US_AMZ.csv", header = T, skip = 1, stringsAsFactors = FALSE, encoding ="UTF-8")
Sales_US$REGION <- "US"

colnames(Sales_US) <- colnames(Sales_UK)

Sales <- rbind.data.frame(Sales_DE, Sales_UK)
Sales <- rbind.data.frame(Sales, Sales_US)

Sales <- Sales[,c(1,2,22)]
colnames(Sales) <- c("RANK", "KEYWORD", "REGION")

keywords_new <- read.csv("UK_Event.csv", header = T, stringsAsFactors = FALSE)
keywords_new$REGION <- "UK"
colnames(keywords_new)[2] <- "KEYWORD"
keywords_new <- filter(keywords_new, duplicated(keywords_new$KEYWORD) == FALSE)
keywords_new_UK <- merge(keywords_new, Sales, by = c("REGION","KEYWORD"), all.x=T)
keywords_new_UK <- merge(keywords_new_UK, Volume, by = c("REGION","RANK"), all.x=T)
keywords_new_UK <- merge(keywords_new_UK, Impression, by = c("REGION","KEYWORD"), all.x=T)
keywords_new_UK <- keywords_new_UK[,c(1:3,15,17,4:14)]
keywords_new_UK <- merge(keywords_new_UK, datahawk, by = c("KEYWORD"), all.x=T)


write.csv(keywords_new_UK, "new_keywords_UK.csv", row.names = F)

keywords_new <- read.csv("US_Event.csv", header = T, stringsAsFactors = FALSE)
keywords_new$REGION <- "US"
colnames(keywords_new)[2] <- "KEYWORD"
keywords_new <- filter(keywords_new, duplicated(keywords_new$KEYWORD) == FALSE)
keywords_new_US <- merge(keywords_new, Sales, by = c("REGION","KEYWORD"), all.x=T)
keywords_new_US <- merge(keywords_new_US, Volume, by = c("REGION","RANK"), all.x=T)
keywords_new_US <- merge(keywords_new_US, Impression, by = c("REGION","KEYWORD"), all.x=T)
keywords_new_US <- keywords_new_US[,c(1:3,15,17,4:14)]
keywords_new_US <- merge(keywords_new_US, datahawk, by = c("KEYWORD"), all.x=T)

write.csv(keywords_new_US, "new_keywords_US.csv", row.names = F)


keywords_new <- read.csv("DE_Event.csv", header = T, stringsAsFactors = FALSE)
keywords_new$REGION <- "DE"
colnames(keywords_new)[2] <- "KEYWORD"
keywords_new <- filter(keywords_new, duplicated(keywords_new$KEYWORD) == FALSE)
keywords_new_DE <- merge(keywords_new, Sales, by = c("REGION","KEYWORD"), all.x=T)
keywords_new_DE <- merge(keywords_new_DE, Volume, by = c("REGION","RANK"), all.x=T)
keywords_new_DE <- merge(keywords_new_DE, Impression, by = c("REGION","KEYWORD"), all.x=T)
keywords_new_DE <- keywords_new_DE[,c(1:3,15,17,4:14)]
keywords_new_DE <- merge(keywords_new_DE, datahawk, by = c("KEYWORD"), all.x=T)

write.csv(keywords_new_DE, "new_keywords_DE.csv", row.names = F)

