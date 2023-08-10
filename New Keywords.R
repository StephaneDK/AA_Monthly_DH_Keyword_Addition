cat("Keywords addition Start\n")
renv::activate()

#Declaring variables for directories change
path_var <- Sys.getenv("path_code")
path_var_csv <- gsub("Code","csv",path_var)
path_var_output <- gsub("Code","Output",path_var)


#Calling outside scripts
source(paste0(path_var,"OutsideBorders.R"))

options(warn = -1)

suppressMessages({
  
  library(tidyverse)
  library(stringr)
  library(reshape2)
  library(ggthemes)
  library(gridExtra)
  library(forecast)
  library(aTSA)
  library(DescTools)
  library(plyr)
  library(EnvStats)
  library(qcc)
  library(openxlsx)
  library(magrittr)
  library(ggrepel)
  
})

options(scipen=999, digits = 3 )

all_days <- c("Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday")

`%notin%` <- Negate(`%in%`)

convert.brackets <- function(x){
  if(grepl("\\(.*\\)", x)){
    paste0("-", gsub("\\(|\\)", "", x))
  } else {
    x
  }
}


#Setting the directory where all files will be used from for this project
setwd(path_var_csv)
#setwd("C:\\Users\\steph\\Documents\\DK\\Work\\Forecasting book sales and inventory\\Auxilary tasks\\Amazon Advertisment\\Keyword Addition\\Monthly\\csv")
#setwd("C:\\Users\\snichanian\\Documents\\DK\\Work\\Forecasting book sales and inventory\\Auxilary tasks\\Amazon Advertisment\\Keyword Addition\\Monthly\\csv")


#Importing keywords to be removed
Remove <- read.csv("Remove.csv", header = T, stringsAsFactors = FALSE)
Remove <- Remove[!duplicated(Remove),]
Remove$remove <- "Y"

#Importing and clean Sales Data
Sales <- read.csv("New Terms DH.csv", header = T, stringsAsFactors = FALSE)

# --------------------------------------------------------------------------------------------------------------
#                                       UK Analysis                                                            \
# --------------------------------------------------------------------------------------------------------------

keywords_new <- read.csv("UK_Event.csv", header = T, stringsAsFactors = FALSE)
keywords_new <- keywords_new %>% 
  dplyr::rename("KEYWORD" = "Target") %>% 
  mutate(REGION = "UK") %>% 
  filter(duplicated(KEYWORD) == FALSE) %>% 
  merge(Sales, by = c("REGION","KEYWORD"), all.x=T) %>% 
  merge(Remove, by = c("REGION","KEYWORD"), all.x=T) %>% 
  mutate(Datahawk = case_when(is.na(Datahawk) | Datahawk == "" ~ "N", TRUE ~ "Y"),
         remove = case_when(is.na(remove) | Datahawk == "" ~ "N" , TRUE ~ "Y") ) %>% 
  filter(remove == "N",
         Datahawk == "N") %>% 
  select(-c(Datahawk,remove))


# --------------------------------------------------------------------------------------------------------------
#                               Saving file + formatting UK                                                    \
# --------------------------------------------------------------------------------------------------------------


write.csv(keywords_new, "new_keywords_UK.csv", row.names = F)


# Adds after the second column
keywords_new <- add_column(keywords_new, new_col = NA, .after = 13)


colnames(keywords_new)[14] <- ""



keywords_new$AdGroup.Name <- as.character(keywords_new$AdGroup.Name)
keywords_new$SEARCH_FREQUENCY_RANK <- as.numeric(keywords_new$SEARCH_FREQUENCY_RANK)
keywords_new$VOLUME <- as.numeric(keywords_new$VOLUME)
keywords_new$IMPRESSIONS <- as.numeric(keywords_new$IMPRESSIONS)


#Highlighting columns with inventory issues
wb <- createWorkbook()
addWorksheet(wb, sheetName="UK")
writeData(wb, sheet="UK", x=keywords_new)


#adding filters
addFilter(wb, "UK", rows = 1, cols = 1:ncol(keywords_new))

#auto width for columns
width_vec <- suppressWarnings(apply(keywords_new, 2, function(x) max(nchar(as.character(x)) + 1, na.rm = TRUE)))
width_vec_header <- nchar(colnames(keywords_new))  + 3
max_vec_header <- pmax(width_vec, width_vec_header)
setColWidths(wb, "UK", cols = 1:ncol(keywords_new), widths = max_vec_header )
setColWidths(wb, "UK",  cols = 1, widths = 10)
setColWidths(wb, "UK",  cols = 2, widths = 45)
setColWidths(wb, "UK",  cols = 3, widths = 25)
setColWidths(wb, "UK",  cols = 4, widths = 15)
setColWidths(wb, "UK",  cols = 5, widths = 25)
setColWidths(wb, "UK",  cols = 6, widths = 15)
setColWidths(wb, "UK",  cols = 7, widths = 15)
setColWidths(wb, "UK",  cols = 8, widths = 45)
setColWidths(wb, "UK",  cols = 9, widths = 15)
setColWidths(wb, "UK",  cols = 10, widths = 15)
setColWidths(wb, "UK",  cols = 11, widths = 15)
setColWidths(wb, "UK",  cols = 12, widths = 15)
setColWidths(wb, "UK",  cols = 13, widths = 15)
setColWidths(wb, "UK",  cols = 14, widths = 15)
setColWidths(wb, "UK",  cols = 15, widths = 18)
setColWidths(wb, "UK",  cols = 16, widths = 20)
setColWidths(wb, "UK",  cols = 17, widths = 10)



#Centering cells
centerStyle <- createStyle(halign = "center")
addStyle(wb, "UK", style=centerStyle, rows = 2:nrow(keywords_new), cols = 15:ncol(keywords_new), 
         gridExpand = T, stack = TRUE)

# Adding borders
invisible(OutsideBorders(
  wb,
  sheet_ = "UK",
  rows_ = 1:nrow(keywords_new)+1,
  cols_ = 1:13
))

invisible(OutsideBorders(
  wb,
  sheet_ = "UK",
  rows_ = 1:nrow(keywords_new)+1,
  cols_ = 15:18
))


freezePane(
  wb,
  sheet = "UK",
  firstActiveRow = 2
)

pred_date <- format(Sys.Date()-30,format="%b %y")

saveWorkbook(wb, paste0("New Keywords UK - ",pred_date,".xlsx"), overwrite = T) 




# --------------------------------------------------------------------------------------------------------------
#                                       US Analysis                                                            \
# --------------------------------------------------------------------------------------------------------------


keywords_new <- read.csv("US_Event.csv", header = T, stringsAsFactors = FALSE)
keywords_new <- keywords_new %>% 
  dplyr::rename("KEYWORD" = "Target") %>% 
  mutate(REGION = "US") %>% 
  filter(duplicated(KEYWORD) == FALSE) %>% 
  merge(Sales, by = c("REGION","KEYWORD"), all.x=T) %>% 
  merge(Remove, by = c("REGION","KEYWORD"), all.x=T) %>% 
  mutate(Datahawk = case_when(is.na(Datahawk) | Datahawk == "" ~ "N", TRUE ~ "Y"),
         remove = case_when(is.na(remove) | Datahawk == "" ~ "N" , TRUE ~ "Y") ) %>% 
  filter(remove == "N",
         Datahawk == "N") %>% 
  select(-c(Datahawk,remove))



# --------------------------------------------------------------------------------------------------------------
#                               Saving file + formatting UK                                                    \
# --------------------------------------------------------------------------------------------------------------


write.csv(keywords_new, "new_keywords_US.csv", row.names = F)

# Adds after the second column
keywords_new <- add_column(keywords_new, new_col = NA, .after = 13)


colnames(keywords_new)[14] <- ""



keywords_new$AdGroup.Name <- as.character(keywords_new$AdGroup.Name)
keywords_new$SEARCH_FREQUENCY_RANK <- as.numeric(keywords_new$SEARCH_FREQUENCY_RANK)
keywords_new$VOLUME <- as.numeric(keywords_new$VOLUME)
keywords_new$IMPRESSIONS <- as.numeric(keywords_new$IMPRESSIONS)


#Highlighting columns with inventory issues
wb <- createWorkbook()
addWorksheet(wb, sheetName="US")
writeData(wb, sheet="US", x=keywords_new)


#adding filters
addFilter(wb, "US", rows = 1, cols = 1:ncol(keywords_new))

#auto width for columns
width_vec <- suppressWarnings(apply(keywords_new, 2, function(x) max(nchar(as.character(x)) + 1, na.rm = TRUE)))
width_vec_header <- nchar(colnames(keywords_new))  + 3
max_vec_header <- pmax(width_vec, width_vec_header)
setColWidths(wb, "US", cols = 1:ncol(keywords_new), widths = max_vec_header )
setColWidths(wb, "US",  cols = 1, widths = 10)
setColWidths(wb, "US",  cols = 2, widths = 45)
setColWidths(wb, "US",  cols = 3, widths = 25)
setColWidths(wb, "US",  cols = 4, widths = 15)
setColWidths(wb, "US",  cols = 5, widths = 25)
setColWidths(wb, "US",  cols = 6, widths = 15)
setColWidths(wb, "US",  cols = 7, widths = 15)
setColWidths(wb, "US",  cols = 8, widths = 45)
setColWidths(wb, "US",  cols = 9, widths = 15)
setColWidths(wb, "US",  cols = 10, widths = 15)
setColWidths(wb, "US",  cols = 11, widths = 15)
setColWidths(wb, "US",  cols = 12, widths = 15)
setColWidths(wb, "US",  cols = 13, widths = 15)
setColWidths(wb, "US",  cols = 14, widths = 15)
setColWidths(wb, "US",  cols = 15, widths = 18)
setColWidths(wb, "US",  cols = 16, widths = 20)
setColWidths(wb, "US",  cols = 17, widths = 10)



#Centering cells
centerStyle <- createStyle(halign = "center")
addStyle(wb, "US", style=centerStyle, rows = 2:nrow(keywords_new), cols = 15:ncol(keywords_new), 
         gridExpand = T, stack = TRUE)

# Adding borders
invisible(OutsideBorders(
  wb,
  sheet_ = "US",
  rows_ = 1:nrow(keywords_new)+1,
  cols_ = 1:13
))

invisible(OutsideBorders(
  wb,
  sheet_ = "US",
  rows_ = 1:nrow(keywords_new)+1,
  cols_ = 15:18
))


freezePane(
  wb,
  sheet = "US",
  firstActiveRow = 2
)

pred_date <- format(Sys.Date()-30,format="%b %y")

saveWorkbook(wb, paste0("New Keywords US - ",pred_date,".xlsx"), overwrite = T) 




# --------------------------------------------------------------------------------------------------------------
#                                       DE Analysis                                                            \
# --------------------------------------------------------------------------------------------------------------



keywords_new <- read.csv("DE_Event.csv", header = T, stringsAsFactors = FALSE)
keywords_new <- keywords_new %>% 
  dplyr::rename("KEYWORD" = "Target") %>% 
  mutate(REGION = "DE") %>% 
  filter(duplicated(KEYWORD) == FALSE) %>% 
  merge(Sales, by = c("REGION","KEYWORD"), all.x=T) %>% 
  merge(Remove, by = c("REGION","KEYWORD"), all.x=T) %>% 
  mutate(Datahawk = case_when(is.na(Datahawk) | Datahawk == "" ~ "N", TRUE ~ "Y"),
         remove = case_when(is.na(remove) | Datahawk == "" ~ "N" , TRUE ~ "Y") ) %>% 
  filter(remove == "N",
         Datahawk == "N") %>% 
  select(-c(Datahawk,remove))

write.csv(keywords_new, "new_keywords_DE.csv", row.names = F)



# --------------------------------------------------------------------------------------------------------------
#                               Saving file + formatting UK                                                    \
# --------------------------------------------------------------------------------------------------------------


write.csv(keywords_new, "new_keywords_DE.csv", row.names = F)

# Adds after the second column
keywords_new <- add_column(keywords_new, new_col = NA, .after = 13)


colnames(keywords_new)[14] <- ""



keywords_new$AdGroup.Name <- as.character(keywords_new$AdGroup.Name)
keywords_new$SEARCH_FREQUENCY_RANK <- as.numeric(keywords_new$SEARCH_FREQUENCY_RANK)
keywords_new$VOLUME <- as.numeric(keywords_new$VOLUME)
keywords_new$IMPRESSIONS <- as.numeric(keywords_new$IMPRESSIONS)


#Highlighting columns with inventory issues
wb <- createWorkbook()
addWorksheet(wb, sheetName="DE")
writeData(wb, sheet="DE", x=keywords_new)


#adding filters
addFilter(wb, "DE", rows = 1, cols = 1:ncol(keywords_new))

#auto width for columns
width_vec <- suppressWarnings(apply(keywords_new, 2, function(x) max(nchar(as.character(x)) + 1, na.rm = TRUE)))
width_vec_header <- nchar(colnames(keywords_new))  + 3
max_vec_header <- pmax(width_vec, width_vec_header)
setColWidths(wb, "DE", cols = 1:ncol(keywords_new), widths = max_vec_header )
setColWidths(wb, "DE",  cols = 1, widths = 10)
setColWidths(wb, "DE",  cols = 2, widths = 45)
setColWidths(wb, "DE",  cols = 3, widths = 25)
setColWidths(wb, "DE",  cols = 4, widths = 15)
setColWidths(wb, "DE",  cols = 5, widths = 25)
setColWidths(wb, "DE",  cols = 6, widths = 15)
setColWidths(wb, "DE",  cols = 7, widths = 15)
setColWidths(wb, "DE",  cols = 8, widths = 45)
setColWidths(wb, "DE",  cols = 9, widths = 15)
setColWidths(wb, "DE",  cols = 10, widths = 15)
setColWidths(wb, "DE",  cols = 11, widths = 15)
setColWidths(wb, "DE",  cols = 12, widths = 15)
setColWidths(wb, "DE",  cols = 13, widths = 15)
setColWidths(wb, "DE",  cols = 14, widths = 15)
setColWidths(wb, "DE",  cols = 15, widths = 18)
setColWidths(wb, "DE",  cols = 16, widths = 20)
setColWidths(wb, "DE",  cols = 17, widths = 10)



#Centering cells
centerStyle <- createStyle(halign = "center")
addStyle(wb, "DE", style=centerStyle, rows = 2:nrow(keywords_new), cols = 15:ncol(keywords_new), 
         gridExpand = T, stack = TRUE)

# Adding borders
invisible(OutsideBorders(
  wb,
  sheet_ = "DE",
  rows_ = 1:nrow(keywords_new)+1,
  cols_ = 1:13
))

invisible(OutsideBorders(
  wb,
  sheet_ = "DE",
  rows_ = 1:nrow(keywords_new)+1,
  cols_ = 15:18
))


freezePane(
  wb,
  sheet = "DE",
  firstActiveRow = 2
)

pred_date <- format(Sys.Date()-30,format="%b %y")

saveWorkbook(wb, paste0("New Keywords DE - ",pred_date,".xlsx"), overwrite = T) 

renv::deactivate()
cat("Keywords addition end\n\n")
