#!/bin/sh

#Keyword Addition path
path_local='C:\Users\steph\Documents\DK\Work\Forecasting book sales and inventory\Auxilary tasks\Amazon Advertisment\Keyword Addition\Monthly\Code\'
path_DK='C:\Users\snichanian\Documents\DK\Work\Forecasting book sales and inventory\Auxilary tasks\Amazon Advertisment\Keyword Addition\Monthly\Code\'


#Change this line for directory
path_code=$path_local

cd "$path_code"

export path_code

python "fetch_data_snfk_dly.py"

Rscript "New Keywords.R"











