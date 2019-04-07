start_time <- Sys.time()
working_output <- getwd()
working_output <- setwd("C:/Users/goura/OneDrive/Desktop/UL_R_Projects/ID_Jarvis_Bango_1.1.0/ID_Jarvis_Bango_1.1.0")
source(paste0(working_output,"/ID_Jarvis_Generalised_template_1.1.0.r"))

############################################# ID Psales (MT Model) ###############################################
psales.mt.input <- merge.data.frame(x = id.psales.bango,y =product.mapping,by.x ="Material_Code",by.y = "Material Code")%>%
  merge.data.frame(x = .,y = customer.mapping,by.x = "Customer_Code_8D",by.y ="Customer Code")%>%
  merge.data.frame(x = .,y = conversion.mapping,by.x = "Material_Code",by.y = "MATNR")

#select the required columns i.e. Required Basepack fromo the material mapping , Account/Region from the Customer Mapping and Conversion rate from the Conversion mapping 

psales.mt <- psales.mt.input%>%
  select(-c(DIVCD:ncol(.),National:UMREN))
#filter out the observations for which the the Account is "MT" and the required basepack is not 0

psales.mt.filtered <- psales.mt%>%
  filter(`REQUIRED BASEPACK`!=0)
psales.mt.filtered <- psales.mt.filtered%>%
  filter(str_detect(`Account/Region`,"MT"))
psales.mt.filtered <- setnames(psales.mt.filtered,c("Material Desc.x","Customer Desc.x"),c("Material Desc","Customer Desc"))
psales.mt.filtered%<>%mutate(psales_KG=`Sales Quantity`*Conversion,psales_Tonns=psales_KG/1000)
#group the observations and aggregate the KPI's based on {Month,Required Basepack,Account}
psales.mt.aggregate <- psales.mt.filtered%>%
  group_by(`Fiscal year/period`,`REQUIRED BASEPACK`,`Account/Region`)%>%
  summarise(`Gross Sales Value (GSV)`=sum(`Gross Sales Value (GSV)`),`Sales Quantity`=sum(`Sales Quantity`),NIV=sum(NIV),psales_KG=sum(psales_KG),psales_Tonns=sum(psales_Tonns))
#calculate the unit price- GSV/Sales_Quantity(Kg)
psales.mt.aggregate%<>%mutate(unit_price=`Gross Sales Value (GSV)`/psales_KG)
#perform the month mapping for the Fiscal Period column 
psales.mt.aggregate <- merge.data.frame(psales.mt.aggregate,month.mapping,by.x = "Fiscal year/period",by.y ="Primary sales" )
psales.mt.aggregate%<>%select(-c(`Sec stock_DROO_Penetration`,TTS_BMI,`Secondary sales`,`Fiscal year/period`,Sellout))
psales.mt.aggregate%<>%select(c(Month,`REQUIRED BASEPACK`,`Account/Region`,`Gross Sales Value (GSV)`:unit_price))



############################################# ID DROO (MT Model) ###############################################

droo.mt.input <- merge.data.frame(x = id.droo.bango,y =product.mapping,by.x ="Material_Code",by.y = "Material Code")%>%
  merge.data.frame(x = .,y = customer.mapping,by.x = "STP_Code",by.y ="Customer Code")%>%
  merge.data.frame(x = .,y = conversion.mapping,by.x = "Material_Code",by.y = "MATNR")
#select the required columns i.e. Required Basepack fromo the material mapping , Account/Region from the Customer Mapping and Conversion rate from the Conversion mapping 
droo.mt <- droo.mt.input%>%
  select(-c(DIVCD:ncol(.),National:UMREN))
##filter out the observations for which the the Account is "MT" and the required basepack is not 0

droo.mt.filtered <- droo.mt%>%
  filter(`REQUIRED BASEPACK`!=0)
droo.mt.filtered <- droo.mt.filtered%>%
  filter(str_detect(`Account/Region`,"MT"))
#perform the groupby and summarise the KPI's for the given observation
droo.mt.aggregate <- droo.mt.filtered%>%
  group_by(`Calendar Year/Month`,`REQUIRED BASEPACK`,`Account/Region`)%>%
  summarise(`OriginalOrder Qty`=sum(`OriginalOrder Qty`),`Final Customer Expected Order Qty`=sum(`Final Customer Expected Order Qty`),`Dispatched Qty`=sum(`Dispatched Qty`))
#calculate the DR and DROO for the aggregated months 
droo.mt.aggregate%<>%mutate(DR=(`Dispatched Qty`/`Final Customer Expected Order Qty`*100),DROO=(`Dispatched Qty`/`OriginalOrder Qty`*100))
droo.mt.aggregate <- merge.data.frame(droo.mt.aggregate,month.mapping,by.x = "Calendar Year/Month",by.y ="Sec stock_DROO_Penetration" )
droo.mt.aggregate%<>%select(-c(`Primary sales`,TTS_BMI,`Secondary sales`,`Calendar Year/Month`,Sellout))
droo.mt.aggregate[which(is.nan(droo.mt.aggregate$DR)),"DR"] <- 0
droo.mt.aggregate%<>%select(c(Month,`REQUIRED BASEPACK`:DROO))

####################################################### Sell out Data ( MT Model) ################################################################
sellout.mt.input.path <- "C:/Users/goura/OneDrive/Desktop/ID_Jarvis/Python ID KT/Input_files/Sellout"
sellout.files <- list.files(path = sellout.mt.input.path,pattern = "^Sell Out ID AF(.*)xlsx|XLSX$")
sellout.files <- paste0(sellout.mt.input.path,"/",sellout.files)
sellout.alfamart.df1 <- read_xlsx(sellout.files,sheet = "SAT",guess_max = 1000,n_max = 2)
sellout.alfamart.df1[3,]<-NA
for(i in 1:ncol(sellout.alfamart.df1)){
  sellout.alfamart.df1[3,i]<-paste(sellout.alfamart.df1[1,i], sellout.alfamart.df1[2,i],sep="-")
}






################################################### ID TTS Mapped (MT Model) ######################################################################
tts.mapped.mt.input <- id.tts.channel.mapping
tts.mapped.mt.df <- tts.mapped.mt.input%>%
  filter(str_detect(`Account/Region`,"MT")|`Account/Region`=="ALL")


tts.bango.mapped.aggregate <- id.tts.channel.mapping%>%
  group_by(`Fiscal year/period`,BRAND,Channel,National)%>%
  summarise(TTS=sum(TTS),BBT=sum(BBT),`BBT - Place`=sum(`BBT - Place`),`BBT - Place on invoice`=sum(`BBT - Place on invoice`),`BBT - Place off invoice`=sum(`BBT - Place off invoice`),`BBT - Price`=sum(`BBT - Price`),`CPP on invoice`=sum(`CPP on invoice`),`CPP off invoice`=sum(`CPP off invoice`),`BBT - Product`=sum(`BBT - Product`),`BBT - Pack`=sum(`BBT - Pack`),`BBT - Proposition`=sum(`BBT - Proposition`),`BBT - Promotion`=sum(`BBT - Promotion`),EOT=sum(EOT))

