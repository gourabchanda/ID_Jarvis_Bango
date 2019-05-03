################################################# Generalised template for the data preprocessing of Input files ##########################################################

################################################# Business Sources and the KPI list for the script ##########################################################################

#There are total 9 internal data source - Primary Sales,Secondary Sales,Secondary Stock,DRO,Penetration,TTS,BMI,Itrust,Sellout
# The base KPI list for Primary Sales   {GSV,NIV,Sales Quantity(In PCS)}
# The base KPI list for Secondary Sales {Sec sales value,Sec sales Qty(PC)}
# The base KPI list for Secondary Stock {Sec stock value,Sec stock Qty(PC)}
# The base KPI list for DROO {Orginal order qty,Final expected Order Qty,Dispatch qty}
# The base KPI list for Penetration {Penetration_Category,Penetration}
# The base KPI list for TTS {TTS,BBT,BBT-Place,BBT-Place on invoice,BBT-Place off invoice,BBT-Price,CPP on invoice,CPP off invoice,BBT-Product,BBT-Pack,BBT-Proposition,BBT-Promotion,EOT}
# The base KPI list for BMI {Brand & Marketing Investment,Brand & Marketing Investment Trade,Brand & Marketing Investment Consumer,Promotion Team Cost Trade,Promotion Team Cost Consumer}-set1
# The base KPI list for BMI {Promotion Team Cost	Promotion Packaging Material Cost Trade	Promotion Packaging Material Cost Consumer,Promotion Communication Material Cost Trade,Promotion Communication Material Cost Consumer}-set2
# The base KPI list for BMI {Promotion Communication Material Cost	,Promo Agency Remuneration Fees & Commissions Consumer,Promotional Expenses,Promotion Packaging Material Cost,Promotion Repacking Cost,Promotion Repacking Cost Trade}-set3
# The base KPI list for BMI {Promotion Repacking Cost Consumer,Promo Samples Gifts and Incentive Costs,Promo SamplesGifts and Incentive Costs,Consumer,Promo Agency Remun Fees & Commissions,Promo Agency Remuneration Fees & Commissions Trade} -set

################################################## Start point of the generalised template script ##########################################################################

# The output values should not in scientific notation use the format(,scientific=F)
# Initiate the script time and calculate the script time for the generalised template , display it back to the user
# load the required packages into the r script , if the package is not present in the r-env it will install aand load the package into the environment

start.time <- Sys.time()
options(scipen=999)
options(digits=22)
packages.required <- c("dplyr","readxl","writexl","reshape","zoo","tidyr","stringr","tibble","data.table","bit64","readr","hablar","magrittr")
packages.diff <- setdiff(packages.required,rownames(installed.packages()))
if(length(packages.diff>1)){
  install.packages(packages.diff,dependencies = T)
}
invisible(sapply(packages.required,library,character.only=T))

######## initiate the mapping path and the directory path for the script and load the input files ###############################

folder.psales.path <- "C:/Users/Gourab.Chanda/OneDrive - Unilever/ID_Jarvis_Bango-HICW10NA512SG4H/Extracted_Files/Primary Sales"
folder.ssales.path <- "C:/Users/Gourab.Chanda/OneDrive - Unilever/ID_Jarvis_Bango-HICW10NA512SG4H/Extracted_Files/Secondary Sales"
folder.sstock.path <- "C:/Users/Gourab.Chanda/OneDrive - Unilever/ID_Jarvis_Bango-HICW10NA512SG4H/Extracted_Files/Secondary Stock"
folder.tts.path <- "C:/Users/Gourab.Chanda/OneDrive - Unilever/ID_Jarvis_Bango-HICW10NA512SG4H/Extracted_Files/TTS"
folder.bmi.path <- "C:/Users/Gourab.Chanda/OneDrive - Unilever/ID_Jarvis_Bango-HICW10NA512SG4H/Extracted_Files/BMI"
folder.droo.path <- "C:/Users/Gourab.Chanda/OneDrive - Unilever/ID_Jarvis_Bango-HICW10NA512SG4H/Extracted_Files/DROO"
folder.itrust.path <- "C:/Users/Gourab.Chanda/OneDrive - Unilever/ID_Jarvis_Bango-HICW10NA512SG4H/Extracted_Files/Itrust"
folder.penetration.path <- "C:/Users/Gourab.Chanda/OneDrive - Unilever/ID_Jarvis_Bango-HICW10NA512SG4H/Extracted_Files/Penetration/Penetration_xlsx"
folder.sellout.path <- "C:/Users/Gourab.Chanda/OneDrive - Unilever/ID_Jarvis_Bango-HICW10NA512SG4H/Extracted_Files/Sellout"
consolidated.backup.path <- "C:/Users/Gourab.Chanda/OneDrive - Unilever/ID_Jarvis_Bango-HICW10NA512SG4H/Consolidated_Backup"
mapping.file.path <- "C:/Users/Gourab.Chanda/OneDrive - Unilever/ID_Jarvis_Bango-HICW10NA512SG4H/Mapping_File/Jarvis ID- Universal Mapping.xlsx"
unmapped.files.path <- "C:/Users/Gourab.Chanda/OneDrive - Unilever/ID_Jarvis_Bango-HICW10NA512SG4H/Unmapped_Files"


product.mapping <- tibble::as_tibble(read_xlsx(paste0(mapping.file.path),sheet = "Material",col_names = T,guess_max = 100))
customer.mapping <- tibble::as_tibble(read_xlsx(paste0(mapping.file.path),sheet = "Customer",col_names = T,guess_max = 100))
conversion.mapping <- tibble::as_tibble(read_xlsx(paste0(mapping.file.path),sheet ="Conversion",col_names = T,guess_max = 100 ))
banner.mapping <- tibble::as_tibble(read_xlsx(paste0(mapping.file.path),sheet = "Banner",col_names = T))
banner.mapping[which((banner.mapping$`Banner Desc`=="Not assigned")),"Banner code"] <- "#";rownames(banner.mapping) <- NULL
banner.mapping$`Banner code` <- as.character(banner.mapping$`Banner code`)
week.mapping <- tibble::as_tibble(read_xlsx(paste0(mapping.file.path),sheet = "Week",col_names = T,guess_max = 100))
month.mapping <- tibble::as_tibble(read_xlsx(paste0(mapping.file.path),sheet = "Month",col_names = T,guess_max = 100))
unmapped.tts.mapping <- tibble::as_tibble(read_xlsx(paste0(mapping.file.path),sheet = "Unmapped TTS",col_names = T,guess_max = 100))
unmapped.bmi.mapping <- tibble::as_tibble(read_xlsx(paste0(mapping.file.path),sheet = "Unmapped BMI",col_names = T,guess_max = 100))
penetration.customer.region.mapping <- tibble::as_tibble(read_xlsx(paste0(mapping.file.path),sheet = "Penetration Customer mapping",col_names = T,guess_max = 100))
psales.material.mt.mapping <- tibble::as_tibble(read_xlsx(paste0(mapping.file.path),sheet = "Material_PS_Sellout",col_names = T,guess_max = 100))
sellout.material.mapping <- tibble::as_tibble(read_xlsx(paste0(mapping.file.path),sheet = "Sellout Material",col_names = T,guess_max = 100))
sellout.customer.mapping <- tibble::as_tibble(read_xlsx(paste0(mapping.file.path),sheet = "Sell out Customer",col_names = T,guess_max = 100))
subbrand.mapping <- tibble::as_tibble(read_xlsx(paste0(mapping.file.path),sheet = "Sub brands",col_names = T,guess_max = 100))

###################################################################### Data Preprocessing User Defined Functions ##################################################

#write the function to convert the extracted input file into a dataframe and also print the structure and dimension of the table
input_clean <- function(extracted_df){
  extracted_df <- extracted_df[,5:ncol(extracted_df)]
  colnames(extracted_df) <- as.character(unlist(extracted_df[1,]))
  clean_df <- extracted_df[-1,]
  print(str(clean_df))
  dim(clean_df)
  return(clean_df)
}
#write the function to perform the customised mapping between the input files and the mapping file provided by the business
mapping <- function(input,map,map_key1,map_key2){
  mapped_bango_df <- merge.data.frame(x = input,y = map,by.x =map_key1,by.y = map_key2,all.x = T)
  return(mapped_bango_df)
}
#write the function which will transform the penetration input data into tidy dataset
penetration_transpose <- function(input_df){
  
  penetration.long <- tidyr::gather(input_df,Month_Year,penetration_value,2:ncol(input_df))
  penetration.long[,1] <- penetration.long[,2]
  penetration.long <- penetration.long[c(1,3)] 
  penetration <- penetration.long
  return(penetration)
}
#write a function which will return the required packgroup dataframe
packgroup_selectdf <- function(input_packgroup_df){
  #filter out the observations for which the requiered packgroup is !=0 and the DT
  output_packgroup_df <- input_packgroup_df%>%
    filter(`REQUIRED PACKGROUP`!=0)%>%
    filter(str_detect(`Account/Region`,"DT"))%>%
    filter(`Account/Region`!=0)%>%
    select(-c(`Material Desc.y`:`BASEPACK NO`,`PACKGROUP NAME`:`Sub channel`,National:UMREN,DIVCD:ncol(.)))
  return(output_packgroup_df)
}
#write a function which will perform the data preprocessing for the region : other Island (oi) penetration data
penetration_oi_format <-  function(input_penetration_oi){
  colnames(input_penetration_oi) <- input_penetration_oi[1,]
  input_penetration_oi <- input_penetration_oi[-1,]
  input_penetration_oi%<>%slice(1)
  input_penetration_oi$"Account/Region" <- "DT OI"
  input_penetration_oi <- input_penetration_oi[2:ncol(input_penetration_oi)]
  input_penetration_oi %<>% select(c(`Account/Region`,everything()))
  input_penetration_oi<- gather(input_penetration_oi,key = "Month", "Penetration_Region",2:ncol(input_penetration_oi))
  return(input_penetration_oi)
  
}
#write a function which will perform the data preprocessing for the 3 region : Central & East Java (CEJ) ; Sumatera ; West Jave (WJ)
penetration_3regions_format <- function(input_penetration_3region){
  mapped.3region <- merge.data.frame(input_penetration_3region,penetration.customer.region.mapping,by = "Local Sales Force 2(m.d.)\\Calendar Year/Month")
  mapped.3region <- mapped.3region[which(mapped.3region$`Account/Region`!=0),]
  mapped.3region%<>%select(-c(`Local Sales Force 2(m.d.)\\Calendar Year/Month`))
  mapped.3region%<>%select(c(`Account/Region`,everything()))
  mapped.3region <- gather(mapped.3region,key = "Month","Penetration_Region",2:ncol(mapped.3region))
  return(mapped.3region)
  
}
#Write a function which will calculate the penetration percentage (%) 
penetration_percentage <- function(penetration_final_df){
  penetration_final_df%<>%hablar::convert(num(Penetration_Region,Penetration_Category))
  penetration.region.final <- penetration_final_df%>%
    mutate(penetration_percentage = Penetration_Region/Penetration_Category*100)
  return(penetration.region.final)
  
}
#write a function which will return the final penetration output data frame 
penetration_packgroup_output <- function(penetration_final_df){
  packgroup.month.mapping <- month.mapping%>%select(c(`Sec stock_DROO_Penetration`,Month))
  penetration_final_df <- merge.data.frame(penetration_final_df,packgroup.month.mapping,by.x = "Month",by.y = "Sec stock_DROO_Penetration")
  penetration_final_df%<>%select(-Month)
  penetration_final_df <- penetration_final_df%>%
    select(c(Month.y,`Account/Region`:penetration_percentage))
  penetration_final_df <- dplyr::rename(penetration_final_df,Month=Month.y)
  return(penetration_final_df)
}


sellout_preprocessing <- function(input_sellout_data1,input_sellout_data2,mapping_set1=month.mapping,mapping_Set=sellout.customer.mapping){
  input_sellout_data1[3,]<-NA
  for(i in 1:ncol(input_sellout_data1)){
    input_sellout_data1[3,i]<-paste(input_sellout_data1[1,i], input_sellout_data1[2,i],sep="-")
  }
  input_sellout_data1 <- input_sellout_data1[3,]
  input_sellout_data2.split1 <- input_sellout_data2[,1:5]
  colnames(input_sellout_data2.split1) <- as.character(unlist(input_sellout_data2.split1[1,]))
  input_sellout_data2.split2 <- input_sellout_data2[6:ncol(input_sellout_data2)]
  colnames(input_sellout_data2.split2) <- input_sellout_data1
  output_sellout_data <- cbind(input_sellout_data2.split1,input_sellout_data2.split2)
  output_sellout_data <- output_sellout_data[-1,]
  output_sellout_data <- output_sellout_data[, -grep("NA", colnames(output_sellout_data))]
  output_sellout_data <- output_sellout_data%>%
    tidyr::gather(.,Sellout_Type,Sellout_Value,6:ncol(.))
  output_sellout_data <- output_sellout_data%>%
    filter(str_detect(Sellout_Type,"WITHOUT"))
  output_sellout_data <- output_sellout_data%>%
    tidyr::separate(.,col = Sellout_Type,into = c("Sellout_Type","Month"),sep = "-",remove = T)
  output_sellout_data$New_Date <-as.Date(as.numeric(output_sellout_data$Month),origin="1899-12-30")
  output_sellout_data <- output_sellout_data%>%
    hablar::convert(chr(New_Date))
  output_sellout_data$month <- as.integer(substr(output_sellout_data$New_Date,6,7))
  output_sellout_data$year <- as.integer(substr(output_sellout_data$New_Date,1,4))
  output_sellout_data$Sellout <- format(as.Date(paste0(output_sellout_data$month,"-1-",output_sellout_data$year), "%m-%d-%Y"),"%m.%Y")
  output_sellout_data <- output_sellout_data%>%
    select(-c(Month,New_Date:year))
  output_sellout_data_mapped <- merge.data.frame(output_sellout_data,month.mapping,by.x = "Sellout",by.y = "Sellout")
  output_sellout_data_mapped <- output_sellout_data_mapped%>%
    select(-c(Sellout,`Primary sales`:TTS_BMI))
  output_sellout_data_mapped%<>%select(c(Month,MAP:Sellout_Value))
  output_sellout_data_mapped <-merge.data.frame(x = output_sellout_data_mapped,y = sellout.customer.mapping,by.x = "Customer",by.y = "Customer")
  output_sellout_data_mapped <- output_sellout_data_mapped%>%
    select(-Customer)
  output_sellout_data_mapped <- output_sellout_data_mapped%>%
    select(c(Month,Customer = `Customer Mapped`,MAP:Sellout_Value))
  return(output_sellout_data_mapped)
  
}
################################################################ Data Preprocessing ( Input psales) ######################################################

#fetch the list of all psales files and apply the functions to convert it into tidy dataset

psales.files <- list.files(path = folder.psales.path,pattern = "^ID_Food_Primary Sales GSV(.*)xlsx|XLSX$")
psales.list <- invisible(lapply(paste0(folder.psales.path,"/",psales.files),FUN = read_xlsx,sheet=1,skip=13,guess_max=1000))
names(psales.list) <- basename(psales.files)

cleaned.psales.list <- invisible(lapply(psales.list, input_clean))
cleaned.psales.dfs.list <- lapply(cleaned.psales.list,data.frame)
input.psales.df <- bind_rows(cleaned.psales.dfs.list)

psales.col.headers <- c("Fiscal year/period","Sales Organization","Business Type","Category(m.d.) from SKU","Market(m.d.) from SKU",
                        "Sector(m.d.) from SKU","Brand(PH)(m.d.) from SKU","Material","Material Desc","Customer","Customer Desc",
                        "Gross Sales Value (GSV)","Sales Quantity","NIV")


colnames(input.psales.df) <- psales.col.headers
id.psales.header <- gsub("[\r\n]","",colnames(input.psales.df))
colnames(input.psales.df) <- id.psales.header


#convert the KPI data type from char-> numberic for the calculation purpose
input.psales.df%<>%convert(num(`Gross Sales Value (GSV)`,`Sales Quantity`,NIV))
#setnames for the default Material and customer columns names into  Material Code and Customer Code respectively
input.psales.df <- input.psales.df%<>%dplyr::rename(Material_Code=Material,Customer_Code=Customer)
#extract the 8 digit code from the exisiting customer code and store it in the customer 8 digit code variable 
input.psales.df <- separate(data = input.psales.df,"Customer_Code",c("a","b","c","Customer_Code_8D"),sep = "/",remove = F)
input.psales.df%<>%dplyr::select(-c("a","b","c"))

#filter out the observations for which the brand is "Bango"
brand.bango <- c("BANGO","Bango")
id.psales.bango <- input.psales.df%>%
  filter(`Brand(PH)(m.d.) from SKU`%in% brand.bango)
#replace the NA's in the KPI's with 0 as confirmed by business
psales.kpi.list <- c("Gross Sales Value (GSV)","Sales Quantity","NIV")
for (kpi in psales.kpi.list) {
  id.psales.bango[which(is.na(id.psales.bango[,kpi])),kpi] <- 0
  
}


#id.psales.bango[which(is.na(id.psales.bango$`Gross Sales Value (GSV)`)),"Gross Sales Value (GSV)"] <- 0
#id.psales.bango[which(is.na(id.psales.bango$`Sales Quantity`)),"Sales Quantity"] <- 0
#id.psales.bango[which(is.na(id.psales.bango$NIV)),"NIV"] <- 0
#call the product mapping , customer mapping and conversion mapping functions to find the mapped observations
bango.ps.pm <- mapping(id.psales.bango,product.mapping,map_key1 = "Material_Code",map_key2 = "Material Code")
bango.ps.cm <- mapping(id.psales.bango,customer.mapping,map_key1 = "Customer_Code_8D",map_key2 = "Customer Code")
bango.ps.convm <- mapping(id.psales.bango,conversion.mapping,map_key1 = "Material_Code",map_key2 = "MATNR")
#find out the unmapped observations for the product , customer and conversion mapping and store it in 
bango.ps.npm <- unique(bango.ps.pm[which(is.na(bango.ps.pm$`Mapped Product Name`)),"Material_Code"])
bango.ps.ncm <- unique(bango.ps.pm[which(is.na(bango.ps.cm$Channel)),c("Customer_Code","Customer Desc")])
bango.ps.nconvm <- unique(bango.ps.convm[which(is.na(bango.ps.convm$Conversion)),"Material_Code"])

if(length(bango.ps.npm)>0){
  write.csv(bango.ps.npm,file = paste0(unmapped.files.path,"/psmaterial_unmapped.csv"),row.names = F)
}
if(length(bango.ps.ncm)>0){
  write.csv(bango.ps.ncm,file = paste0(unmapped.files.path,"/pscustomer_unmapped.csv"),row.names = F)
}
if(length(bango.ps.nconvm)>0){
  write.csv(bango.ps.nconvm,file = paste0(unmapped.files.path,"/psconversion_unmapped.csv"),row.names = F)
}

psales_file_name <- psales.files
psales_file_name <- str_remove(psales_file_name,".xlsx")
x <- psales_file_name%>%
  str_match_all("[0-9]+") %>% unlist %>% as.numeric
file_name_split <- unlist(str_split(psales_file_name," "))
file_latest_time <- file_name_split[length(file_name_split)]
time_abb <- lubridate::dmy(paste("01-",file_latest_time,sep=""))
time_abb <- format(time_abb,format="%b %y")


############################################################## Data Preprocessing ( Input ssales) #####################################################
ssales.files <- list.files(path = folder.ssales.path,pattern = "^ID_Food_Secondary Sales IP(.*)xlsx|XLSX$")
ssales.list <- invisible(lapply(paste0(folder.ssales.path,"/",ssales.files),FUN = read_xlsx,sheet=1,skip=13,guess_max=1000))
names(ssales.list) <- basename(ssales.files)

cleaned.ssales.list <- invisible(lapply(ssales.list, input_clean))
cleaned.ssales.dfs.list <- lapply(cleaned.ssales.list,data.frame)
input.ssales.df <- bind_rows(cleaned.ssales.dfs.list)
ssales.col.headers <- c("Calendar Year/Month","Sales Organization","Business Type","Category(m.d.)","Market(m.d.)","Sector(m.d.)",
                        "Brand(PH)(m.d.)","Material","Material Desc","Distributor","Distributor Desc","Gross Sec Sales (TUR)","Sec Volume")
colnames(input.ssales.df) <- ssales.col.headers
id.ssales.header <- gsub("[\r\n]","",colnames(input.ssales.df))
colnames(input.ssales.df) <- id.ssales.header
#convert the KPI data type from char-> numberic for the calculation purpose
input.ssales.df%<>%convert(num(`Gross Sec Sales (TUR)`,`Sec Volume`))
#setnames for the default Material and customer columns names into  Material Code and Customer Code respectively
input.ssales.df <- input.ssales.df%<>%dplyr::rename(Material_Code=Material,Distributor_Code=Distributor)
#extract the 8 digit code from the exisiting customer code and store it in the customer 8 digit code variable 
input.ssales.df <- separate(data = input.ssales.df,"Distributor_Code",c("a","b","c","Distributor_Code_8D"),sep = "/",remove = F)
input.ssales.df%<>%dplyr::select(-c("a","b","c"))
#Filter out the observations for the brand Bango
brand.bango <- c("BANGO","Bango")
id.ssales.bango <- input.ssales.df%>%
  filter(`Brand(PH)(m.d.)`%in% brand.bango)
#store the KPI's in the KPI list
ssales.kpi.list <- c("Gross Sec Sales (TUR)","Sec Volume")
for (kpi in ssales.kpi.list) {
  id.ssales.bango[which(is.na(id.ssales.bango[,kpi])),kpi] <- 0
  
}
#call the product mapping , customer mapping and conversion mapping functions to find the mapped observations
bango.ssales.pm <- mapping(id.ssales.bango,product.mapping,map_key1 = "Material_Code",map_key2 = "Material Code")
bango.ssales.cm <- mapping(id.ssales.bango,customer.mapping,map_key1 = "Distributor_Code_8D",map_key2 = "Customer Code")
bango.ssales.convm <- mapping(id.ssales.bango,conversion.mapping,map_key1 = "Material_Code",map_key2 = "MATNR")

#find out the unmapped observations for the product , customer and conversion mapping and store it in 
bango.ss.npm <- unique(bango.ssales.pm[which(is.na(bango.ssales.pm$`Mapped Product Name`)),"Material_Code"])
bango.ss.ncm <- unique(bango.ssales.cm[which(is.na(bango.ssales.cm$Channel)),c("Distributor_Code","Distributor Desc")])
bango.ss.nconvm <- unique(bango.ssales.convm[which(is.na(bango.ssales.convm$Conversion)),"Material_Code"])

if(length(bango.ss.npm)>0){
  write.csv(bango.ss.npm,file = paste0(unmapped.files.path,"/ssmaterial_unmapped.csv"),row.names = F)
}
if(length(bango.ss.ncm)>0){
  write.csv(bango.ss.ncm,file = paste0(unmapped.files.path,"/sscustomer_unmapped.csv"),row.names = F)
}
if(length(bango.ss.nconvm)>0){
  write.csv(bango.ss.nconvm,file = paste0(unmapped.files.path,"/ssconversion_unmapped.csv"),row.names = F)
}

############################################################## Data Preprocessing  (Input SStock) #################################################################
sstock.files <- list.files(path = folder.sstock.path,pattern = "^ID_Food_Secondary Stock IP(.*)xlsx|XLSX$")
sstock.list <- invisible(lapply(paste0(folder.sstock.path,"/",sstock.files),FUN = read_xlsx,sheet=1,skip=13,guess_max=1000))
names(sstock.list) <- basename(sstock.files)

cleaned.sstock.list <- invisible(lapply(sstock.list, input_clean))
cleaned.sstock.dfs.list <- lapply(cleaned.sstock.list,data.frame)
input.sstock.df <- bind_rows(cleaned.sstock.dfs.list)
sstock.col.headers <- c("Calendar Year/Month","Sales Organization","Category(m.d.)","Market(m.d.)","Sector(m.d.)","Brand(PH)(m.d.)","Material","Material Desc","Sold-to party","STP Desc","Secondary Stock Volume","Secondary Stock Value [@ DT Rate.]")
colnames(input.sstock.df) <- sstock.col.headers
id.sstock.header <- gsub("[\r\n]","",colnames(input.sstock.df))
colnames(input.sstock.df) <- id.sstock.header
#convert the KPI data type from char-> numberic for the calculation purpose

#setnames for the default Material and customer columns names into  Material Code and Customer Code respectively
input.sstock.df <- input.sstock.df%<>%dplyr::rename(Material_Code=Material,STP_Code=`Sold-to party`)
#Filter out the observations for the brand Bango
brand.bango <- c("BANGO","Bango")
id.sstock.bango <- input.sstock.df%>%
  filter(`Brand(PH)(m.d.)`%in% brand.bango)
#store the KPI's in the KPI list
sstock.kpi.list <- c("Secondary Stock Volume","Secondary Stock Value [@ DT Rate.]")
for (kpi in sstock.kpi.list) {
  id.sstock.bango[which(is.na(id.sstock.bango[,kpi])),kpi] <- 0
  
}
#call the product mapping , customer mapping and conversion mapping functions to find the mapped observations
bango.sstock.pm <- mapping(id.sstock.bango,product.mapping,map_key1 = "Material_Code",map_key2 = "Material Code")
bango.sstock.cm <- mapping(id.sstock.bango,customer.mapping,map_key1 = "STP_Code",map_key2 = "Customer Code")
bango.sstock.convm <- mapping(id.sstock.bango,conversion.mapping,map_key1 = "Material_Code",map_key2 = "MATNR")
#find out the unmapped observations for the product , customer and conversion mapping and store it in 
bango.sstock.npm <- unique(bango.sstock.pm[which(is.na(bango.sstock.pm$`Mapped Product Name`)),"Material_Code"])
bango.sstock.ncm <- unique(bango.sstock.cm[which(is.na(bango.sstock.cm$Channel)),c("STP_Code","STP Desc")])
bango.sstock.nconvm <- unique(bango.sstock.convm[which(is.na(bango.sstock.convm$Conversion)),"Material_Code"])

if(length(bango.sstock.npm)>0){
  write.csv(bango.sstock.npm,file = paste0(unmapped.files.path,"/sstock_material_unmapped.csv"),row.names = F)
}
if(length(bango.sstock.ncm)>0){
  write.csv(bango.sstock.ncm,file = paste0(unmapped.files.path,"/sstock_customer_unmapped.csv"),row.names = F)
}
if(length(bango.sstock.nconvm)>0){
  write.csv(bango.sstock.nconvm,file = paste0(unmapped.files.path,"/sstock_conversion_unmapped.csv"),row.names = F)
}
#perform the mapping for the observations from the extracted files with the product , customer and conversion mapping 
sstock.bango <- merge.data.frame(x = id.sstock.bango,y =product.mapping,by.x ="Material_Code",by.y = "Material Code")%>%
  merge.data.frame(x = .,y = customer.mapping,by.x = "STP_Code",by.y ="Customer Code")%>%
  merge.data.frame(x = .,y = conversion.mapping,by.x = "Material_Code",by.y = "MATNR")
######################################################################## Data Preprocessing (Input DROO) ##########################################################333
droo.files <- list.files(path = folder.droo.path,pattern = "^ID_Food_DROO IP_(.*)xlsx|XLSX$")
droo.list <- invisible(lapply(paste0(folder.droo.path,"/",droo.files),FUN = read_xlsx,sheet=1,skip=13,guess_max=1000))
names(droo.list) <- basename(droo.files)

cleaned.droo.list <- invisible(lapply(droo.list, input_clean))
cleaned.droo.dfs.list <- lapply(cleaned.droo.list,data.frame)
input.droo.df <- bind_rows(cleaned.droo.dfs.list)
droo.col.headers <- c("Calendar Year/Month","Sales Organization","Business Type","Category(m.d.)","Market(m.d.)","Sector(m.d.)","Brand(PH)(m.d.)","Material","Material Desc","Sold-to party","STP Desc","OriginalOrder Qty","Final Customer Expected Order Qty","Dispatched Qty")
id.droo.header <- gsub("[\r\n]","",colnames(input.droo.df))
colnames(input.droo.df) <- id.droo.header
colnames(input.droo.df) <- droo.col.headers
#convert the KPI data type from char-> numberic for the calculation purpose
input.droo.df%<>%convert(num(`OriginalOrder Qty`,`Final Customer Expected Order Qty`,`Dispatched Qty`))
#setnames for the default Material and customer columns names into  Material Code and Customer Code respectively
input.droo.df <- input.droo.df%<>%dplyr::rename(Material_Code=Material,STP_Code=`Sold-to party`)
#Filter out the observations for the brand Bango
brand.bango <- c("BANGO","Bango")
id.droo.bango <- input.droo.df%>%
  filter(`Brand(PH)(m.d.)`%in% brand.bango)
#store the KPI's in the KPI list
droo.kpi.list <- c("OriginalOrder Qty","Final Customer Expected Order Qty","Dispatched Qty")
for (kpi in droo.kpi.list) {
  id.droo.bango[which(is.na(id.droo.bango[,kpi])),kpi] <- 0
  
}
#call the product mapping , customer mapping and conversion mapping functions to find the mapped observations
bango.droo.pm <- mapping(id.droo.bango,product.mapping,map_key1 = "Material_Code",map_key2 = "Material Code")
bango.droo.cm <- mapping(id.droo.bango,customer.mapping,map_key1 = "STP_Code",map_key2 = "Customer Code")
bango.droo.convm <- mapping(id.droo.bango,conversion.mapping,map_key1 = "Material_Code",map_key2 = "MATNR")
#find out the unmapped observations for the product , customer and conversion mapping and store it in 
bango.droo.npm <- unique(bango.droo.pm[which(is.na(bango.droo.pm$`Mapped Product Name`)),"Material_Code"])
bango.droo.ncm <- unique(bango.droo.cm[which(is.na(bango.droo.cm$Channel)),c("STP_Code","STP Desc")])
bango.droo.nconvm <- unique(bango.droo.convm[which(is.na(bango.droo.convm$Conversion)),"Material_Code"])

if(length(bango.droo.npm)>0){
  write.csv(bango.droo.npm,file = paste0(unmapped.files.path,"/droo_material_unmapped.csv"),row.names = F)
}
if(length(bango.sstock.ncm)>0){
  write.csv(bango.droo.ncm,file = paste0(unmapped.files.path,"/droo_customer_unmapped.csv"),row.names = F)
}
if(length(bango.droo.nconvm)>0){
  write.csv(bango.droo.nconvm,file = paste0(unmapped.files.path,"/droo_conversion_unmapped.csv"),row.names = F)
}
#perform the mapping for the observations from the extracted files with the product , customer and conversion mapping 
droo.bango <- merge.data.frame(x = id.droo.bango,y =product.mapping,by.x ="Material_Code",by.y = "Material Code")%>%
  merge.data.frame(x = .,y = customer.mapping,by.x = "STP_Code",by.y ="Customer Code")%>%
  merge.data.frame(x = .,y = conversion.mapping,by.x = "Material_Code",by.y = "MATNR")
############################################################# (Itrust) ######################################################
input.itrust.path <- paste0(folder.itrust.path,"/Itrust.xlsx")
itrust.input <- tibble::as_tibble(read_xlsx(input.itrust.path,sheet = 1,guess_max = 1000))
#if (!"YEAR" %in% colnames(itrust.input)) {
  #itrust.input[,"YEAR"] <- substr(basename(input.itrust.path),8,12)
  
#}

itrust.input[,"ITRUST_LINE"]<- NA
itrust.input[,"ITRUST_TOTAL"] <- NA

itrust.input[which(itrust.input$FINAL_CRR==0),"ITRUST_LINE"] <- 0
itrust.input$ITRUST_LINE <-ifelse(itrust.input$FINAL_CRR==0,0,ifelse(itrust.input$STOCK_CS>itrust.input$FINAL_CRR,1,0))
itrust.input$ITRUST_TOTAL <- ifelse(itrust.input$FINAL_CRR==0,0,1)

itrust.input <- itrust.input%>%
  tidyr::unite(col = "Week_Concat",c("WK","YEAR"),sep=".",remove=F)

itrust.input$Week_Concat <- stringr::str_trim(itrust.input$Week_Concat,side = "both")
#perform the product and customer mapping for the itrust observations
itrust.bango.pm <- mapping(itrust.input,product.mapping,map_key1 ="PROD_CODE",map_key2 = "Material Code")
itrust.bango.cm <- mapping(itrust.input,customer.mapping,map_key1 ="CUST_CODE",map_key2 = "Customer Code")
itrust.bango.npm <- unique(itrust.bango.pm[which(is.na(itrust.bango.pm$BRAND)),"PROD_CODE"])
itrust.bango.ncm <- unique(itrust.bango.cm[which(is.na(itrust.bango.cm$Channel)),"CUST_CODE"])
if(length(itrust.bango.npm)>0){
  write.csv(itrust.bango.npm,file = paste0(unmapped.files.path,"/itrustmaterial_unmapped.csv"),row.names = F)
}
if(length(itrust.bango.ncm)>0){
  write.csv(itrust.bango.ncm,file = paste0(unmapped.files.path,"/itrustcustomer_unmapped.csv"),row.names = F)
}

itrust.input$Week_Concat <- stringr::str_trim(itrust.input$Week_Concat)
#################################################################### TTS (Mapped)###############################################
tts.files <- list.files(path = folder.tts.path,pattern = "^ID_Food_Primary Sales TTS_(.*)xlsx|XLSX$")
tts.list <- invisible(lapply(paste0(folder.tts.path,"/",tts.files),FUN = read_xlsx,sheet=1,skip=13,guess_max=1000))
names(tts.list) <- basename(tts.files)

cleaned.tts.list <- invisible(lapply(tts.list, input_clean))
cleaned.tts.dfs.list <- lapply(cleaned.tts.list,data.frame)
input.tts.df <- bind_rows(cleaned.tts.dfs.list)
tts.col.headers <- c("Fiscal year/period","Sales Organization",NA,"Category(m.d.) from SKU","Sector(m.d.) from SKU","Market(m.d.) from SKU",
                     "Brand(m.d.)","Material",NA,"Banner(m.d.)",NA,"Banner",NA,"TTS","BBT","BBT - Place","BBT - Place on invoice","BBT - Place off invoice",
                     "BBT - Price","CPP on invoice","CPP off invoice","BBT - Product","BBT - Pack","BBT - Proposition","BBT - Promotion","EOT")
colnames(input.tts.df) <- tts.col.headers
id.tts.header <- gsub("[\r\n]","",colnames(input.tts.df))
colnames(input.tts.df) <- id.tts.header
input.tts.headers <- colnames(input.tts.df)
input.tts.headers[(which(is.na(input.tts.headers)))] <- c("Business_Type","Material_Description","Banner(m.d)_Description","Banner_Description")
colnames(input.tts.df) <- input.tts.headers
#convert the KPI list flrom character to numeric
mapped.tts.material <- mapping(input = input.tts.df,map = product.mapping,map_key1 ="Material",map_key2 = "Material Code" )
#filter out the observations for which the brand is BANGO
id.tts.mapped.bango <- mapped.tts.material%>%
  filter(BRAND =="BANGO")

#convert the KPI list flrom character to numeric
id.tts.mapped.bango[,14:26] <- sapply(id.tts.mapped.bango[,14:26],as.numeric)
#remove the unrequired columns from the data frame
id.tts.mapped.bango1 <-id.tts.mapped.bango%>%
  select(-c(27:32,34:ncol(id.tts.mapped.bango)))
#perform the banner mapping to fetch the channel for the mapped banners
id.tts.channel.mapping <- mapping(input=id.tts.mapped.bango1,map =banner.mapping, map_key1 = "Banner",map_key2 = "Banner code")

id.ttschannel.na <- unique(id.tts.channel.mapping[which(is.na(id.tts.channel.mapping$Channel)),"Banner"])

id.tts.channel.mapping <- id.tts.channel.mapping%>%
  select(1:13,27:32,14:26)

for (i in 20:ncol(id.tts.channel.mapping)){
  id.tts.channel.mapping[which(is.na(id.tts.channel.mapping[,i])),i] <- 0
  
}
########################################################################### TTS (Unmapped)########################################################
input.tts.unmapped.path <- paste0(folder.tts.path,"/TTS Bango-Unmapped.xlsx")
tts.unmapped.input <- as_tibble(read_xlsx(paste0(input.tts.unmapped.path),sheet = 1,guess_max = 1000,skip = 12))

unmapped.tts.input <- input_clean(tts.unmapped.input)
input.unmapped.tts.headers <- gsub("[\r\n]","",colnames(unmapped.tts.input))
colnames(unmapped.tts.input) <- input.unmapped.tts.headers

colnames(unmapped.tts.input)[which(is.na(colnames(unmapped.tts.input)))] <- "Unnammed_cols"

colnames(unmapped.tts.input)[which(colnames(unmapped.tts.input)=="Unnammed_cols")] <-c ("Business_Type","Material_Description","Banner_Name","TFL_Channel","LSF2_Description","LSF3_Description")

#perform the material mapping to fetch the Brands 

unmapped.material.tts.mapping <- merge.data.frame(x = unmapped.tts.input,y = product.mapping,by.x ="Material",by.y= "Material Code",all.y = T)
unmapped.material.tts.mapping1 <- unmapped.material.tts.mapping%>%
  select(-c(32:37,39:42))
unmapped.material.tts.mapping1 <- unmapped.material.tts.mapping1[,c(32,1:ncol(unmapped.material.tts.mapping1)-1)]
id.tts.unmapped.df <- unmapped.material.tts.mapping1
#filter out the rows for which the Brand_Name is Bango

#convert the NA's from the KPI columns to 0 as mentioned by business

#for (i in 20:ncol(id.tts.unmapped.df)) {
#id.tts.unmapped.df[which(is.na(id.tts.unmapped.df[,i])),i] <- 0}

#create the key for the channel mapping and map the channels for the same
id.tts.unmapped.df <- id.tts.unmapped.df%>%
  unite(col = "key",c("Banner","Trade Format Level 2","Local Sales Force 2(m.d.)","Local Sales Force 3(m.d.)","Key Customer Level3"),sep="",remove = F)
id.tts.unmapped.channel.mapping <- merge.data.frame(id.tts.unmapped.df,unmapped.tts.mapping,by.x = "key",by.y = "Key")

id.tts.unmapped.bango <- id.tts.unmapped.channel.mapping[,-c(1,37:41)]

x <- unique(id.tts.unmapped.bango[which(is.na(id.tts.unmapped.bango$Channel)),"Material"])
write.csv(x = x,file = paste0(unmapped.files.path,"/unmapped.tts.nochannel.csv"))
id.tts.unmapped.rna <- id.tts.unmapped.bango[which(!is.na(id.tts.unmapped.bango$Channel)),]

id.tts.unmapped.rna1 <-  id.tts.unmapped.rna[,c(3:9,2,10,1,11:19,33:35,20:32)]

id.tts.unmapped.rna1[23:35] <- sapply(id.tts.unmapped.rna1[23:35],as.numeric)
#################################################################################### BMI (Mapped)###########################################################
bmi.files <- list.files(path = folder.bmi.path,pattern = "^ID_Food_Primary Sales BMI(.*)xlsx|XLSX$")
bmi.list <- invisible(lapply(paste0(folder.bmi.path,"/",bmi.files),FUN = read_xlsx,sheet=1,skip=13,guess_max=1000))
names(bmi.list) <- basename(bmi.files)
cleaned.bmi.list <- invisible(lapply(bmi.list, input_clean))
cleaned.bmi.dfs.list <- lapply(cleaned.bmi.list,data.frame)
input.bmi.df <- bind_rows(cleaned.bmi.dfs.list)
id.bmi.mapped.df1 <- input.bmi.df[,1:20]
id.bmi.mapped.df2 <- input.bmi.df[,21:ncol(input.bmi.df)]
colnames(id.bmi.mapped.df1) <- as.character(unlist(id.bmi.mapped.df1[1,]))
id.bmi.mapped.df1 <- id.bmi.mapped.df1[-1,]
id.bmi.mapped.df2 <- id.bmi.mapped.df2[-1,]

id.bmi.mapped.df <- cbind(id.bmi.mapped.df1,id.bmi.mapped.df2)
rownames(id.bmi.mapped.df) <- NULL


id.bmi.mapped.df[,21:ncol(id.bmi.mapped.df)] <- sapply(id.bmi.mapped.df[,21:ncol(id.bmi.mapped.df)],as.numeric)
colnames(id.bmi.mapped.df)[which(is.na(colnames(id.bmi.mapped.df)))] <- "Unnammed_cols"
bmi.bango.mapped.headers <- gsub("[\r\n]","",colnames(id.bmi.mapped.df))
colnames(id.bmi.mapped.df) <- bmi.bango.mapped.headers 

for (i in 21:ncol(id.bmi.mapped.df)) {
  id.bmi.mapped.df[which(is.na(id.bmi.mapped.df[,i])),i] <- 0
  
}


#rename the headers as per the convention
colnames(id.bmi.mapped.df)[colnames(id.bmi.mapped.df)=="Brands"] <- "Brand_Code"
names(id.bmi.mapped.df)[16] <- "Brand"
bango.names <- c("Bango","BANGO")
#filter out the brand Bango rows from the input file 
bmi.mapped.bango <- subset(x = id.bmi.mapped.df,subset = Brand=="Bango")
bmi.bango.mapped <- merge.data.frame(x = bmi.mapped.bango,y =banner.mapping,by.x = "Banner",by.y = "Banner code",all.x = T )
#find out the Banner codes for which the Channel are missing i.e. "NA"
bmi.mapped.bango.channel.na <- unique(bmi.bango.mapped[which(is.na(bmi.bango.mapped$Channel)),"Banner"])
############################################################### BMI (Unmapped)##############################################################################
input.bmi.unmapped.path <- paste0(folder.bmi.path,"/BMI Bango-Unmapped.xlsx")
bmi.unmapped.input <- as_tibble(read_xlsx(paste0(input.bmi.unmapped.path),sheet = 1,guess_max = 1000,skip = 12))
#bmi.headers <- bmi.input[2,]
id.bmi.unmapped.df <- input_clean(bmi.unmapped.input)
bmi.unmapped.headers1 <- colnames(id.bmi.unmapped.df)
#id.bmi.unmapped.df1 <- id.bmi.unmapped.df[,c(1:20,42:43)]
#id.bmi.unmapped.df2 <- id.bmi.unmapped.df[,21:41]
bmi.unmapped.headers1 <- gsub("[\r\n]","",colnames(id.bmi.unmapped.df))


colnames(id.bmi.unmapped.df) <- as.character(unlist(id.bmi.unmapped.df[1,]))
id.bmi.unmapped.df <- id.bmi.unmapped.df[-1,]
#change the KPI headers by picking the headers from bmi.unmapped.header1
names(id.bmi.unmapped.df)[21:41] <- bmi.unmapped.headers1[21:41]
colnames(id.bmi.unmapped.df)[which(is.na(colnames(id.bmi.unmapped.df)))] <- "Unnammed_cols"

colnames(id.bmi.unmapped.df)[which(colnames(id.bmi.unmapped.df)=="Unnammed_cols")] <-c ("a","Brand_Name","b","TFL_Channel","c","d","e")

#filter out the rows for which the Brand_Name is Bango

id.bmi.unmapped.df <- subset.data.frame(x = id.bmi.unmapped.df,subset = Brand_Name=="Bango")

#convert the NA's from the KPI columns to 0 as mentioned by business

for (i in 21:41) {
  id.bmi.unmapped.df[which(is.na(id.bmi.unmapped.df[,i])),i] <- 0
  
}


#create the key to fetch the mapped channels from the table 

id.bmi.unmapped.df1 <- id.bmi.unmapped.df%>%
  unite(col = "key",c("Banner","Trade Format Level 2","Local Sales Force 2","Local Sales Force 3","Key Customer Level3","ConsumpOccassClass05","Local Sales Force 1"),sep="",remove = F)

id.bmi.unmapped.df1 <- id.bmi.unmapped.df1[,-43]
#perform the join on the key column and fetch the channels 

unmapped.bmi.mapping1 <- unmapped.bmi.mapping[,c(1,2)]

id.bmi.unmapped.join <- merge.data.frame(x = id.bmi.unmapped.df1,y = unmapped.bmi.mapping1,by.x = "key",by.y ="Key")

setnames(x = id.bmi.unmapped.join,old = "Channel.y",new = "Channel")
id.bmi.unmapped.join <- id.bmi.unmapped.join%>%
  select(-c(Channel.x))

id.bmi.unmapped.join[22:42] <- sapply(id.bmi.unmapped.join[22:42],as.numeric)

#calculate the script time for the benchmark purpose

end.time <- Sys.time()
script.time <- round(end.time-start.time)
print(script.time)



