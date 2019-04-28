
working_directory <- getwd()
working_directory <- setwd("C:/Users/goura/OneDrive/Desktop/Unilever_R_Projects/ID_Jarvis_Bango/ID_Jarvis_2.0.0")
source(paste0(working_directory,"/data_preprocessing_script.r"))

############################################# ID Psales (total Bango GT) ###############################################

#perform the mapping for the observations from the extracted files with the product , customer and conversion mapping 
psales.bango <- merge.data.frame(x = id.psales.bango,y =product.mapping,by.x ="Material_Code",by.y = "Material Code")%>%
  merge.data.frame(x = .,y = customer.mapping,by.x = "Customer_Code_8D",by.y ="Customer Code")%>%
  merge.data.frame(x = .,y = conversion.mapping,by.x = "Material_Code",by.y = "MATNR")

#select the required columns from the DF change the column names in the data frame for Material Description and the Customer Description 
psales.bango.df1  <- psales.bango%>%
  select(-c(2,16:21,23:27,32:34,36:ncol(.)))%>%
  select(c(2:8,1,15,9:11,16:19,12:14,20))
psales.bango.df1 <- setnames(psales.bango.df1,c("Material Desc.x","Customer Desc.x"),c("Material Desc","Customer Desc"))
#calculate the derived KPI's from the exisiting KPI's using the conversion factor and add the columns into the data frame
psales.bango.df1%<>%mutate(psales_KG=`Sales Quantity`*Conversion,psales_Tonns=psales_KG/1000)
#group by observations and summarise the KPI's based on the pkey {Month,Brand,Channel}
psales.bango.aggregate <- psales.bango.df1%>%
  group_by(`Fiscal year/period`,BRAND,Channel)%>%
  summarise(`Gross Sales Value (GSV)`=sum(`Gross Sales Value (GSV)`),`Sales Quantity`=sum(`Sales Quantity`),NIV=sum(NIV),psales_KG=sum(psales_KG),psales_Tonns=sum(psales_Tonns))
#perform the month mapping for the data frame values in the given period column in the psales.bango.aggregate data frame
psales.bango.aggregate <- merge.data.frame(psales.bango.aggregate,month.mapping,by.x = "Fiscal year/period",by.y ="Primary sales" )
psales.bango.aggregate%<>%select(-c(`Sec stock_DROO_Penetration`,TTS_BMI,`Secondary sales`,`Fiscal year/period`))
psales.bango.aggregate%<>%select(c(Month,BRAND,Channel,`Gross Sales Value (GSV)`,`Sales Quantity`,NIV,psales_KG,psales_Tonns))
#rename the column names as per the output file
psales.bango.aggregate <- setnames(psales.bango.aggregate,old = c("Gross Sales Value (GSV)","Sales Quantity","NIV","psales_KG","psales_Tonns"),new=c("GSV","Sales Qty","NIV","Sales Qty (Kg)","Sales Qty(Tonns)"))
#filter out the rows based on the channel froom the aggregate DS
channel.type <- c("DISTRIBUTIVE TRADE","MODERN TRADE","OTHER")
#grouped data for the DT/GT
psales.gt <- psales.bango.aggregate%>%
  filter(Channel==channel.type[1])
#grouped data for the MT
psales.mt <- psales.bango.aggregate%>%
  filter(Channel==channel.type[2])
#grouped data for the others
psales.others <- psales.bango.aggregate%>%
  filter(Channel==channel.type[3])
########################################################### ID Ssales (Total Bango GT) ########################################################
#perform the mapping for the observations from the extracted files with the product , customer and conversion mapping 
ssales.bango <- merge.data.frame(x = id.ssales.bango,y =product.mapping,by.x ="Material_Code",by.y = "Material Code")%>%
  merge.data.frame(x = .,y = customer.mapping,by.x = "Distributor_Code_8D",by.y ="Customer Code")%>%
  merge.data.frame(x = .,y = conversion.mapping,by.x = "Material_Code",by.y = "MATNR")

ssales.bango.df <- ssales.bango%>%
  select(-c(2,15:20,22:26,31:33,35:ncol(.)))%>%
  select(c(2:8,1,9,14,10:11,15:18,12:13,19))
ssales.bango.df1 <- ssales.bango.df
#calculate the derived KPI's from the exisiting KPI's using the conversion factor and add the columns into the data frame
ssales.bango.df1%<>%mutate(ssales_KG=`Sec Volume`*Conversion,ssales_Tonns=ssales_KG/1000)
ssales.bango.df1 <- setnames(ssales.bango.df1,c("Material Desc.x"),c("Material Desc"))
#group by observations and summarise the KPI's based on the pkey {Month,Brand,Channel}
ssales.bango.aggregate <- ssales.bango.df1%>%
  group_by(`Calendar Year/Month`,BRAND,Channel)%>%
  summarise(`Gross Sec Sales (TUR)`=sum(`Gross Sec Sales (TUR)`),`Sec Volume`=sum(`Sec Volume`),ssales_KG=sum(ssales_KG),ssales_Tonns=sum(ssales_Tonns))
ssales.bango.aggregate <- setnames(ssales.bango.aggregate,old = c("Gross Sec Sales (TUR)","Sec Volume","ssales_KG","ssales_Tonns"),new=c( "GT_Sec Sales Value","GT_Sec Sales Volume","GT_Sec Sales Volume(Kg)","Secondary Volume Qty (Tonns)"))
#perform the month mapping for the data frame values in the given period column in the psales.bango.aggregate data frame
ssales.bango.aggregate <- merge.data.frame(ssales.bango.aggregate,month.mapping,by.x ="Calendar Year/Month",by.y ="Secondary sales")
ssales.bango.aggregate%<>%select(-c(`Sec stock_DROO_Penetration`,TTS_BMI,`Primary sales`,`Calendar Year/Month`))
ssales.bango.aggregate%<>%select(c(Month,BRAND,Channel,`GT_Sec Sales Value`,`GT_Sec Sales Volume`,`GT_Sec Sales Volume(Kg)`,`Secondary Volume Qty (Tonns)`))

################################################################ ID SStock(Total Bango GT)#########################################################
sstock.bango.df <- sstock.bango%>%
  select(-c(13:18,20:24,29:31,33:ncol(.)))%>%
  select(c(3:8,1,9,13,2,10,14:17,11:12,18))

sstock.bango.df1 <- sstock.bango.df
#calculate the derived KPI's from the exisiting KPI's using the conversion factor and add the columns into the data frame
sstock.bango.df1%<>%mutate(sstock_KG=`Secondary Stock Volume`*Conversion,sstock_Tonns=sstock_KG/1000)
sstock.bango.df1 <- setnames(sstock.bango.df1,c("Material Desc.x"),c("Material Desc"))
#group by observations and summarise the KPI's based on the pkey {Month,Brand,Channel}
sstock.bango.aggregate <- sstock.bango.df1%>%
  group_by(`Calendar Year/Month`,BRAND,Channel)%>%
  summarise(`Secondary Stock Volume`=sum(`Secondary Stock Volume`),`Secondary Stock Value [@ DT Rate.]`=sum(`Secondary Stock Value [@ DT Rate.]`),sstock_KG=sum(sstock_KG),sstock_Tonns=sum(sstock_Tonns))
sstock.bango.aggregate <- setnames(sstock.bango.aggregate,old = c("Secondary Stock Volume","Secondary Stock Value [@ DT Rate.]","sstock_KG","sstock_Tonns"),new=c( "Secondary Stock Volume","Secondary Stock Value [@ DT Rate.]","Secondary Stock Volume(Kg)","Secondary Stock Volume(Tonns)"))
#perform the month mapping for the data frame values in the given period column in the psales.bango.aggregate data frame
sstock.bango.aggregate <- merge.data.frame(sstock.bango.aggregate,month.mapping,by.x ="Calendar Year/Month",by.y ="Sec stock_DROO_Penetration")
sstock.bango.aggregate%<>%select(-c(`Primary sales`,`Secondary sales`,TTS_BMI))
sstock.bango.aggregate%<>%select(c(Month,BRAND,Channel,`Secondary Stock Volume`,`Secondary Stock Value [@ DT Rate.]`,`Secondary Stock Volume(Kg)`,`Secondary Stock Volume(Tonns)`))
################################################################# ID DROO ( Total Bango GT)#################################################################

droo.bango <- merge.data.frame(x = id.droo.bango,y =product.mapping,by.x ="Material_Code",by.y = "Material Code")%>%
  merge.data.frame(x = .,y = customer.mapping,by.x = "STP_Code",by.y ="Customer Code")%>%
  merge.data.frame(x = .,y = conversion.mapping,by.x = "Material_Code",by.y = "MATNR")

droo.bango.df <- droo.bango%>%
  select(-c(15:20,22:26))%>%
  select(c(3:9,2,10,1,15,11,16:19,12:14))
#perform the group by and the aggregation of the KPI's
droo.aggregate <- droo.bango.df%>%
  group_by(`Calendar Year/Month`,BRAND,Channel)%>%
  summarise(`OriginalOrder Qty`=sum(`OriginalOrder Qty`),`Final Customer Expected Order Qty`=sum(`Final Customer Expected Order Qty`),`Dispatched Qty`=sum(`Dispatched Qty`))
#calculate the DR and DROO for the aggregated months 
droo.aggregate%<>%mutate(DR=(`Dispatched Qty`/`Final Customer Expected Order Qty`*100),DROO=(`Dispatched Qty`/`OriginalOrder Qty`*100))
#mapping the month mapping to get the period in <Month>-<Year>
droo.bango.aggregate <- mapping(input = droo.aggregate,map = month.mapping,map_key1 ="Calendar Year/Month",map_key2 ="Sec stock_DROO_Penetration")
droo.bango.aggregate <- select(droo.bango.aggregate,-c(1,9:11))
droo.bango.aggregate <- select(droo.bango.aggregate,c(8,1:7))
#store the channel in the channel.type variable
channel.type <- c("DISTRIBUTIVE TRADE","MODERN TRADE","OTHER")
#grouped data for GT
droo.gt <- droo.bango.aggregate%>%
  filter(Channel==channel.type[1])
#grouped data for the MT
droo.mt <- droo.bango.aggregate%>%
  filter(Channel==channel.type[2])
#grouped data for the others
droo.others <- droo.bango.aggregate%>%
  filter(Channel==channel.type[3])
######################################################## ID Penetration (Total Bango GT) ################################################
penetration.path <- paste0(input.path,"/GD ID PENETRATION BANGO.xlsx")
penetration.cat.path <- paste0(input.path,"/GD ID PENETRATION Total Category.xlsx")
input.penetration <- tibble::as_tibble(readxl::read_xlsx(path = paste0(penetration.path),sheet = 1,guess_max = 1000,skip = 13 ))
input.penetration.category <- tibble::as_tibble(readxl::read_xlsx(path = paste0(penetration.cat.path),sheet = 1,guess_max = 1000,skip = 13))
penetration.df <- input_clean(input.penetration)
penetration.cat.df <- input_clean(input.penetration.category)
colnames(penetration.df) <- penetration.df[1,]
penetration.df <- penetration.df[-1,]
colnames(penetration.cat.df) <- penetration.cat.df[1,]
penetration.cat.df <- penetration.cat.df[-1,]
penetration.df <- penetration.df%>%
  slice(1)
penetration.cat.df <- penetration.cat.df%>%
  slice(1)

penetration.df.tp <- penetration_transpose(penetration.df)
penetration.df.tp$penetration_value <- as.numeric(penetration.df.tp$penetration_value)
penetration.cat.df.tp <- penetration_transpose(penetration.cat.df)
setnames(x = penetration.cat.df.tp,old = "penetration_value",new = "Penetration_cat_value")
penetration.cat.df.tp$Penetration_cat_value <- as.numeric(penetration.cat.df.tp$Penetration_cat_value)


penetration.blend.df <- mapping(penetration.df.tp,penetration.cat.df.tp,map_key1 = "Calendar Year/Month",map_key2 = "Calendar Year/Month")
penetration.blend.df <- penetration.blend.df%>%
  mutate(penetration_percentage = penetration_value/Penetration_cat_value*100)
#mapping the month mapping to get the period in <Month>-<Year>
penetration.blend.df <- mapping(input = penetration.blend.df,map = month.mapping,map_key1 ="Calendar Year/Month",map_key2 ="Sec stock_DROO_Penetration")
penetration.blend.df <- select(penetration.blend.df,-c(1,5:7))
penetration.blend.df <- select(penetration.blend.df,c(4,1:3))
penetration.gt <- penetration.blend.df
penetration.gt <- setnames(penetration.gt,old =c("penetration_value","Penetration_cat_value","penetration_percentage") ,new =c("Penetration","Category_Penetration","Penetration%") )
penetration.gt$BRAND <- "BANGO"
penetration.gt$Channel <- "DISTRIBUTIVE TRADE"
penetration.gt <- penetration.gt%<>%select(c(1,5:6,2:4))
################################################################## ID Itrust ( Total Bango - GT) ###########################################

itrust.bango <- merge.data.frame(x = itrust.input,y = product.mapping,by.x = "PROD_CODE",by.y = "Material Code")%>%
  merge.data.frame(x=.,y = customer.mapping,by.x ="CUST_CODE",by.y = "Customer Code")%>%
  merge.data.frame(x=.,y=week.mapping,by.x="Week_Concat",by.y = "Concatenation")
itrust.bango.gt <- itrust.bango%>%
  select(-c(1,10:15,17:20,26:ncol(.)-1))%>%
  select(c(14,3,6,2,9,1,10:13,4:5,7:8))

colnames(itrust.bango)[colnames(itrust.bango)=="WK.x"] <- "WK"


itrust.bango.gt <- itrust.bango%>%
  group_by(Month,BRAND,Channel)%>%
  summarise(ITRUST_LINE=sum(ITRUST_LINE),ITRUST_TOTAL=sum(ITRUST_TOTAL))%>%
  mutate(ITRUST_PERCENTAGE=ITRUST_LINE/ITRUST_TOTAL*100)

itrust.bango.gt <- itrust.bango.gt
itrust.bango.gt <- setnames(itrust.bango.gt,old = c("ITRUST_LINE","ITRUST_TOTAL","ITRUST_PERCENTAGE") ,new=c("I TRUST LINE","I TRUST TOTAL","I Trust %"))
itrust.gt <- itrust.bango.gt
################################################################# ID TTS ( Total bango - GT) ##########################################################

#group the entire observations based on Period , Brand and Channel and aggregare the KPI's

tts.bango.mapped.aggregate <- id.tts.channel.mapping%>%
  group_by(`Fiscal year/period`,BRAND,Channel,National)%>%
  summarise(TTS=sum(TTS),BBT=sum(BBT),`BBT - Place`=sum(`BBT - Place`),`BBT - Place on invoice`=sum(`BBT - Place on invoice`),`BBT - Place off invoice`=sum(`BBT - Place off invoice`),`BBT - Price`=sum(`BBT - Price`),`CPP on invoice`=sum(`CPP on invoice`),`CPP off invoice`=sum(`CPP off invoice`),`BBT - Product`=sum(`BBT - Product`),`BBT - Pack`=sum(`BBT - Pack`),`BBT - Proposition`=sum(`BBT - Proposition`),`BBT - Promotion`=sum(`BBT - Promotion`),EOT=sum(EOT))
tts.bango.mapped.aggregate1 <- tts.bango.mapped.aggregate[which(!is.na(tts.bango.mapped.aggregate$Channel)),]
tts.bango.mapped.aggregate1 <- tts.bango.mapped.aggregate1%>%
  select(-c(National))
#aggregate the unmapped tts based on month ,brand,channel
id.tts.unmapped.aggregate <- id.tts.unmapped.rna1%>%
  group_by(`Fiscal year/period`,BRAND,Channel)%>%
  summarise(TTS=sum(TTS),BBT=sum(BBT),`BBT - Place`=sum(`BBT - Place`),`BBT - Place on invoice`=sum(`BBT - Place on invoice`),`BBT - Place off invoice`=sum(`BBT - Place off invoice`),`BBT - Price`=sum(`BBT - Price`),`CPP on invoice`=sum(`CPP on invoice`),`CPP off invoice`=sum(`CPP off invoice`),`BBT - Product`=sum(`BBT - Product`),`BBT - Pack`=sum(`BBT - Pack`),`BBT - Proposition`=sum(`BBT - Proposition`),`BBT - Promotion`=sum(`BBT - Promotion`),EOT=sum(EOT))

tts.bango.unmapped.aggregate1 <- id.tts.unmapped.aggregate
tts.bango.aggregate <- rbind(tts.bango.mapped.aggregate1,tts.bango.unmapped.aggregate1)
#perform the group by and the aggregate on the dataset

tts.bango.aggregate.final <- tts.bango.aggregate%>%
  group_by(`Fiscal year/period`,BRAND,Channel)%>%
  summarise(TTS=sum(TTS),BBT=sum(BBT),`BBT - Place`=sum(`BBT - Place`),`BBT - Place on invoice`=sum(`BBT - Place on invoice`),`BBT - Place off invoice`=sum(`BBT - Place off invoice`),`BBT - Price`=sum(`BBT - Price`),`CPP on invoice`=sum(`CPP on invoice`),`CPP off invoice`=sum(`CPP off invoice`),`BBT - Product`=sum(`BBT - Product`),`BBT - Pack`=sum(`BBT - Pack`),`BBT - Proposition`=sum(`BBT - Proposition`),`BBT - Promotion`=sum(`BBT - Promotion`),EOT=sum(EOT))
#perform the month mapping

tts.bango.aggregate.final <- merge.data.frame(tts.bango.aggregate.final,month.mapping,by.x = "Fiscal year/period",by.y = "TTS_BMI" )
tts.bango.aggregate.final <- tts.bango.aggregate.final%>%
  select(-c(`Primary sales`,`Secondary sales`,`Sec stock_DROO_Penetration`,`Fiscal year/period`))

tts.bango.aggregate.final <- tts.bango.aggregate.final%>%
  select(c(16,1:15))
#filter out the gt and mt channels observations

tts.bango.gt <- tts.bango.aggregate.final%>%
  filter(str_detect(Channel,"DISTRIBUTIVE")|Channel=="ALL")
tts.bango.mt <- tts.bango.aggregate.final%>%
  filter(str_detect(Channel,"MODERN TRADE")|Channel=="ALL")

########################################################################### ID BMI (Mapped)################################################################


bmi.bango.mapped.aggregate <- bmi.bango.mapped%>%
  group_by(`Fiscal year/period`,Brand,Channel)%>%
  summarise(X..Brand...Marketing.Investment=sum(X..Brand...Marketing.Investment),X..Brand...Marketing.Investment.Trade=sum(X..Brand...Marketing.Investment.Trade),X..Brand...Marketing.Investment.Consumer=sum(X..Brand...Marketing.Investment.Consumer),X..Promotional.Expenses=sum(X..Promotional.Expenses),X....Promotion.Packaging.Material.Cost=sum(X....Promotion.Packaging.Material.Cost),X......Promotion.Packaging.Material.Cost.Trade=sum(X......Promotion.Packaging.Material.Cost.Trade),X......Promotion.Packaging.Material.Cost.Consumer=sum(X......Promotion.Packaging.Material.Cost.Consumer),X....Promotion.Repacking.Cost=sum(X....Promotion.Repacking.Cost),X......Promotion.Repacking.Cost.Trade=sum(X....Promotion.Repacking.Cost),X......Promotion.Repacking.Cost.Trade=sum(X......Promotion.Repacking.Cost.Trade),X......Promotion.Repacking.Cost.Consumer=sum(X......Promotion.Repacking.Cost.Consumer),X....Promotion.Communication.Material.Cost=sum(X....Promotion.Communication.Material.Cost),X......Promotion.Communication.Material.Cost.Trade=sum(X......Promotion.Communication.Material.Cost.Trade),X......Promotion.Communication.Material.Cost.Consumer=sum(X......Promotion.Communication.Material.Cost.Consumer),X....Promo.Samples..Gifts.and.Incentive.Costs=sum(X....Promo.Samples..Gifts.and.Incentive.Costs),X......Promo.Samples..Gifts.and.Incentive.Costs.Consumer=sum(X......Promo.Samples..Gifts.and.Incentive.Costs.Consumer),X....Promotion.Team.Cost=sum(X....Promotion.Team.Cost),X......Promotion.Team.Cost.Trade=sum(X......Promotion.Team.Cost.Trade),X......Promotion.Team.Cost.Consumer=sum(X......Promotion.Team.Cost.Consumer),X....Promo.Agency.Remun.Fees...Commissions=sum(X....Promo.Agency.Remun.Fees...Commissions),X......Promo.Agency.Remuneration.Fees...Commissions.Trade=sum(X......Promo.Agency.Remuneration.Fees...Commissions.Trade),X......Promo.Agency.Remuneration.Fees...Commissions.Consumer=sum(X......Promo.Agency.Remuneration.Fees...Commissions.Consumer))
#remove the bmi observations for which the channel is NA 

bmi.bango.mapped.aggregate <- bmi.bango.mapped.aggregate[which(!is.na(bmi.bango.mapped.aggregate$Channel)),]


bmi.bango.mapped.aggregate1 <- bmi.bango.mapped.aggregate%>%
  dplyr::rename(`Brand & Marketing Investment`=X..Brand...Marketing.Investment,`Brand & Marketing Investment Trade`=X..Brand...Marketing.Investment.Trade,`Brand & Marketing Investment Consumer`=X..Brand...Marketing.Investment.Consumer,`Promotional Expenses`=X..Promotional.Expenses,`Promotion Packaging Material Cost`=X....Promotion.Packaging.Material.Cost,`Promotion Communication Material Cost Trade`=X......Promotion.Communication.Material.Cost.Trade,`Promotion Packaging Material Cost Consumer`=X......Promotion.Packaging.Material.Cost.Consumer,`Promotion Repacking Cost`=X....Promotion.Repacking.Cost,`Promotion Repacking Cost Trade`=X......Promotion.Repacking.Cost.Trade,`Promotion Repacking Cost Consumer`=X......Promotion.Repacking.Cost.Consumer,`Promotion Communication Material Cost`=X....Promotion.Communication.Material.Cost,`Promotion Communication Material Cost Consumer`=X......Promotion.Communication.Material.Cost.Consumer,`Promo Samples, Gifts and Incentive Costs` =X....Promo.Samples..Gifts.and.Incentive.Costs,`Promo Samples, Gifts and Incentive Costs Consumer`=X......Promo.Samples..Gifts.and.Incentive.Costs.Consumer,`Promotion Team Cost`=X....Promotion.Team.Cost,`Promotion Team Cost Trade`=X......Promotion.Team.Cost.Trade,`Promotion Team Cost Consumer`=X......Promotion.Team.Cost.Consumer,`Promo Agency Remun Fees & Commissions`=X....Promo.Agency.Remun.Fees...Commissions,`Promo Agency Remuneration Fees & Commissions Trade`=X......Promo.Agency.Remuneration.Fees...Commissions.Trade,`Promo Agency Remuneration Fees & Commissions Consumer`=X......Promo.Agency.Remuneration.Fees...Commissions.Consumer)

bmi.bango.mapped.aggregate1%<>%dplyr::rename(`Promotion Packaging Material Cost Trade`=X......Promotion.Packaging.Material.Cost.Trade)

############################################################################# ID BMI (Unmapped)###################################################################
bmi.bango.unmapped.aggregate <- id.bmi.unmapped.join%>%
  group_by(`Fiscal year/period`,Brand_Name,Channel)%>%
  summarise(`Brand & Marketing Investment`=sum(`Brand & Marketing Investment`),`Brand & Marketing Investment Trade`=sum(`Brand & Marketing Investment Trade`),`Brand & Marketing Investment Consumer`=sum(`Brand & Marketing Investment Consumer`),`Promotional Expenses`=sum(`Promotional Expenses`),`Promotion Packaging Material Cost`=sum(`Promotion Packaging Material Cost`),`Promotion Packaging Material Cost Trade`=sum(`Promotion Repacking Cost Trade`),`Promotion Packaging Material Cost Consumer`=sum(`Promotion Packaging Material Cost Consumer`),`Promotion Repacking Cost`=sum(`Promotion Repacking Cost`),`Promotion Repacking Cost Trade`=sum(`Promotion Repacking Cost Trade`),`Promotion Repacking Cost Consumer`=sum(`Promotion Repacking Cost Consumer`),`Promotion Communication Material Cost`=sum(`Promotion Communication Material Cost`),`Promotion Communication Material Cost Trade`=sum(`Promotion Communication Material Cost Trade`),`Promotion Communication Material Cost Consumer`=sum(`Promotion Communication Material Cost Consumer`),`Promo Samples, Gifts and Incentive Costs`=sum(`Promo Samples, Gifts and Incentive Costs`),`Promo Samples, Gifts and Incentive Costs Consumer`=sum(`Promo Samples, Gifts and Incentive Costs Consumer`),`Promotion Team Cost`=sum(`Promotion Team Cost`),`Promotion Team Cost Trade`=sum(`Promotion Team Cost Trade`),`Promotion Team Cost Consumer`=sum(`Promotion Team Cost Consumer`),`Promo Agency Remun Fees & Commissions`=sum(`Promo Agency Remun Fees & Commissions`),`Promo Agency Remuneration Fees & Commissions Trade`=sum(`Promo Agency Remuneration Fees & Commissions Trade`),`Promo Agency Remuneration Fees & Commissions Consumer`=sum(`Promo Agency Remuneration Fees & Commissions Consumer`))

setnames(bmi.bango.unmapped.aggregate,old = "Brand_Name",new = "Brand")

bmi.bango.aggregate <- rbind(bmi.bango.mapped.aggregate1,bmi.bango.unmapped.aggregate)
#perform the mapping for the months 

bmi.bango.aggregate1 <- mapping(input = bmi.bango.aggregate,map = month.mapping,map_key1 = "Fiscal year/period",map_key2 = "TTS_BMI" )

bmi.bango.aggregate1$`Fiscal year/period` <- bmi.bango.aggregate1$Month
setnames(bmi.bango.aggregate1,old = "Fiscal year/period",new = "Month")
bmi.bango.aggregate1 <- bmi.bango.aggregate1[,1:24]

bmi.bango.aggregate1$Brand <- "BANGO"


############################################################################ BMI Consolidation ( Mapped+Unmapped) ######################################3

#perform the group by and aggregation on the entire bmi.bango.aggregate data frame

bmi.bango.aggregate2 <- bmi.bango.aggregate1%>%
  group_by(Month,Brand,Channel)%>%
  summarise(`Brand & Marketing Investment`=sum(`Brand & Marketing Investment`),`Brand & Marketing Investment Trade`=sum(`Brand & Marketing Investment Trade`),`Brand & Marketing Investment Consumer`=sum(`Brand & Marketing Investment Consumer`),`Promotional Expenses`=sum(`Promotional Expenses`),`Promotion Packaging Material Cost`=sum(`Promotion Packaging Material Cost`),`Promotion Packaging Material Cost Trade`=sum(`Promotion Packaging Material Cost Trade`),`Promotion Packaging Material Cost Consumer`=sum(`Promotion Packaging Material Cost Consumer`),`Promotion Repacking Cost`=sum(`Promotion Repacking Cost`),`Promotion Repacking Cost Trade`=sum(`Promotion Repacking Cost Trade`),`Promotion Repacking Cost Consumer`=sum(`Promotion Repacking Cost Consumer`),`Promotion Communication Material Cost`=sum(`Promotion Communication Material Cost`),`Promotion Communication Material Cost Trade`=sum(`Promotion Communication Material Cost Trade`),`Promotion Communication Material Cost Consumer`=sum(`Promotion Communication Material Cost Consumer`),`Promo Samples, Gifts and Incentive Costs`=sum(`Promo Samples, Gifts and Incentive Costs`),`Promo Samples, Gifts and Incentive Costs Consumer`=sum(`Promo Samples, Gifts and Incentive Costs Consumer`),`Promotion Team Cost`=sum(`Promotion Team Cost`),`Promotion Team Cost Trade`=sum(`Promotion Team Cost Trade`),`Promotion Team Cost Consumer`=sum(`Promotion Team Cost Consumer`),`Promo Agency Remun Fees & Commissions`=sum(`Promo Agency Remun Fees & Commissions`),`Promo Agency Remuneration Fees & Commissions Trade`=sum(`Promo Agency Remuneration Fees & Commissions Trade`),`Promo Agency Remuneration Fees & Commissions Consumer`=sum(`Promo Agency Remuneration Fees & Commissions Consumer`))




bmi.bango.aggregate <- bmi.bango.aggregate2%>%
  select(c(Month,Brand,Channel,`Brand & Marketing Investment`,`Brand & Marketing Investment Trade`,`Brand & Marketing Investment Consumer`
           ,`Promotion Team Cost Trade`,`Promotion Team Cost Consumer`,`Promotion Team Cost`
           ,`Promotion Packaging Material Cost Trade`,`Promotion Packaging Material Cost Consumer`,`Promotion Communication Material Cost Trade`
           ,`Promotion Communication Material Cost Consumer`,`Promotion Communication Material Cost`,`Promo Agency Remuneration Fees & Commissions Consumer`
           ,`Promotional Expenses`,`Promotion Packaging Material Cost`,`Promotion Repacking Cost`
           ,`Promotion Repacking Cost Trade`,`Promotion Repacking Cost Consumer`,`Promotion Repacking Cost Consumer`,`Promo Samples, Gifts and Incentive Costs`
           ,`Promo Samples, Gifts and Incentive Costs Consumer`,`Promo Agency Remun Fees & Commissions`,`Promo Agency Remun Fees & Commissions`,`Promo Agency Remuneration Fees & Commissions Trade`))

#filter out the rows which has distributive string in the channel and also "All" in the channel
bmi.bango.gt <- bmi.bango.aggregate%>%
  filter(str_detect(Channel,"DISTRIBUTIVE")|Channel=="ALL")
bmi.bango.gt <- setnames(bmi.bango.gt,old = "Brand",new = "BRAND")

bmi.bango.mt <- bmi.bango.aggregate%>%
  filter(str_detect(Channel,"MODERN TRADE")|Channel=="ALL")


######################################################################## GT Consolidation #############################################

#consolidate the 8 - sources KPI's as per the output
psales.bango.gt <- psales.gt; ssales.bango.gt <- ssales.bango.aggregate
sstock.bango.gt <- sstock.bango.aggregate; droo.bango.gt <- droo.gt
tts.bango.gt <- tts.bango.gt;bmi.bango.gt <- bmi.bango.gt
penetration.bango.gt <- penetration.gt; itrsut.bango.gt <- itrust.gt

primary_key <- c("Month","BRAND","Channel")
#check for the penetration mapping as in penetration the Brand and Channel are missing 
consolidated.list <- plyr::join_all(list(psales.bango.gt,ssales.bango.gt,sstock.bango.gt,droo.bango.gt,itrust.bango.gt,penetration.bango.gt),by= primary_key,type = "left")
consolidated.list1 <- merge.data.frame(x = consolidated.list,y = tts.bango.gt,by = primary_key,all.y = T)
consolidated.list2 <- merge.data.frame(x=consolidated.list1,y = bmi.bango.gt,by=primary_key,all.y = T)

consolidated.list2[is.na(consolidated.list2)] <- ""
total.bango.gt <- consolidated.list2
total.bango.gt$`Sales Organization` <- as.character(9001)
total.bango.gt$Category <- "SAVOURY"

total.bango.gt$Sector <- "SEASONINGS"

total.bango.gt%<>%select(Month,`Sales Organization`,Category,Sector,BRAND,Channel,GSV,`Sales Qty`,NIV,`Sales Qty (Kg)`,`Sales Qty(Tonns)`,`GT_Sec Sales Value`,`GT_Sec Sales Volume`,`GT_Sec Sales Volume(Kg)`,`Secondary Volume Qty (Tonns)`,`Secondary Stock Volume`,`Secondary Stock Value [@ DT Rate.]`,`Secondary Stock Volume(Kg)`,`Secondary Stock Volume(Tonns)`,`OriginalOrder Qty`,`Final Customer Expected Order Qty`,`Dispatched Qty`,DR,DROO,TTS,BBT,`BBT - Place`,`BBT - Place on invoice`,`BBT - Place off invoice`,`BBT - Price`,`CPP on invoice`,`CPP off invoice`,`BBT - Product`,`BBT - Pack`,`BBT - Proposition`,`BBT - Promotion`,EOT,`Brand & Marketing Investment`,`Brand & Marketing Investment Trade`,`Brand & Marketing Investment Consumer`,`Promotion Team Cost Trade`,`Promotion Team Cost Consumer`,`Promotion Team Cost`,`Promotion Packaging Material Cost Trade`,`Promotion Packaging Material Cost Consumer`,`Promotion Communication Material Cost Trade`,`Promotion Communication Material Cost Consumer`,`Promotion Communication Material Cost`,`Promo Agency Remuneration Fees & Commissions Consumer`,`Promotional Expenses`,`Promotion Packaging Material Cost`,`Promotion Repacking Cost`,`Promotion Repacking Cost Trade`,`Promotion Repacking Cost Consumer`,`Promo Samples, Gifts and Incentive Costs`,`Promo Samples, Gifts and Incentive Costs Consumer`,`Promo Agency Remun Fees & Commissions`,`Promo Agency Remuneration Fees & Commissions Trade`,Penetration,Category_Penetration,`Penetration%`,`I TRUST LINE`,`I TRUST TOTAL`,`I Trust %`)
total.bango.gt.header.new <- c("Month",
                               "Sales Organization",
                               "Category",
                               "Sector",
                               "Brand",
                               "Channel",
                               "Bango GT_GSV",
                               "Bango GT_Sales Qty",
                               "Bango GT_NIV",
                               "Bango GT_Sales Qty(kg)",
                               "Bango GT_Sales Qty(Tonns)",
                               "Bango GT_Sec Sales Value",
                               "Bango GT_Sec Sales Volume",
                               "Bango GT_Sec Sales Volume(Kg)",
                               "Bango GT_Sec Sales Volume(Tonns)",
                               "Bango GT_Secondary Stock Volume",
                               "Bango GT_Secondary Stock Value [@ DT Rate.]",
                               "Bango GT_Secondary Stock Volume(Kg)",
                               "Bango GT_Secondary Stock Volume(Tonns)",
                               "Bango GT_OriginalOrder Qty.",
                               "Bango GT_Final Customer Expected Order Qty.",
                               "Bango GT_Dispatched Qty.",
                               "Bango GT_DR",
                               "Bango GT_DROO",
                               "Bango GT_TTS",
                               "Bango GT_BBT",
                               "Bango GT_BBT-Place",
                               "Bango GT_BBT-Place on invoice",
                               "Bango GT_BBT-Place off invoice",
                               "Bango GT_BBT-Price",
                               "Bango GT_CPP on invoice",
                               "Bango GT_CPP off invoice",
                               "Bango GT_BBT-Product",
                               "Bango GT_BBT-Pack",
                               "Bango GT_BBT-Proposition",
                               "Bango GT_BBT-Promotion",
                               "Bango GT_EOT",
                               "Bango GT_Brand & Marketing Investment",
                               "Bango GT_Brand & Marketing Investment Trade",
                               "Bango GT_Brand & Marketing Investment Consumer",
                               "Bango GT_Promotion Team Cost Trade",
                               "Bango GT_Promotion Team Cost Consumer",
                               "Bango GT_Promotion Team Cost",
                               "Bango GT_Promotion Packaging Material Cost Trade",
                               "Bango GT_Promotion Packaging Material Cost Consumer",
                               "Bango GT_Promotion Communication Material Cost Trade",
                               "Bango GT_Promotion Communication Material Cost Consumer",
                               "Bango GT_Promotion Communication Material Cost",
                               "Bango GT_Promo Agency Remuneration Fees & Commissions Consumer",
                               "Bango GT_Promotional Expenses",
                               "Bango GT_Promotion Packaging Material Cost",
                               "Bango GT_Promotion Repacking Cost",
                               "Bango GT_Promotion Repacking Cost Trade",
                               "Bango GT_Promotion Repacking Cost Consumer",
                               "Bango GT_Promo Samples, Gifts and Incentive Costs",
                               "Bango GT_Promo Samples, Gifts and Incentive Costs Consumer",
                               "Bango GT_Promo Agency Remun Fees & Commissions",
                               "Bango GT_Promo Agency Remuneration Fees & Commissions Trade",
                               "Bango_Penetration",
                               "Category_Penetration",
                               "Bango_Penetration%",
                               "Bango_I TRUST LINE",
                               "Bango_I TRUST TOTAL",
                               "Bango_I Trust %"
)

colnames(total.bango.gt) <- total.bango.gt.header.new

total.bango.gt.final <- total.bango.gt%>%
  arrange(Channel,Month)

total.bango.gt.final <- format(total.bango.gt.final,scientific = F) 

writexl::write_xlsx(total.bango.gt.final,paste0(directory.path,"/formated_total_bango_gt.xlsx"),format_headers = T)

######################################################### Bango MT Consolidation ######################################################
psales.bango.mt <- psales.mt; droo.bango.mt <- droo.mt
tts.bango.mt <- tts.bango.mt ; bmi.bango.mt <- bmi.bango.mt

primary_key <- c("Month","BRAND","Channel")
consolidated.list.mt <- merge.data.frame(x =psales.bango.mt,y = droo.bango.mt,by = primary_key )
consolidated.list.mt <- merge.data.frame(consolidated.list.mt,bmi.bango.mt,by = primary_key,all.y = T)
consolidated.list.mt <- merge.data.frame(consolidated.list.mt,tts.bango.mt,by = primary_key,all.x = T)
consolidated.list.mt[is.na(consolidated.list.mt)] <- ""
total.bango.mt <- consolidated.list.mt
total.bango.mt$`Sales Organization` <- as.character(9001)
total.bango.mt$Category <- "SAVOURY"
total.bango.mt$Sector <- "SEASONINGS"

#rearrange the columns along with the output column name as per the requirement
total.bango.mt.r <- total.bango.mt%>%
  select(c(1,48:50,2:3,4:13,35:47,14:34))
total.bango.mt.r1 <- total.bango.mt.r%>%
  arrange(Channel,Month)



total.bango.mt.header.new <- c("Month",
                               "Sales Organization",
                               "Category",
                               "Sector",
                               "Brand",
                               "Channel",
                               "Bango_GSV",
                               "Bango_Sales Qty",
                               "Bango_NIV",
                               "Sales Qty(Kg)",
                               "Sales Qty(Tonns)",
                               "Bango_OriginalOrder Qty.",
                               "Bango_Final Customer Expected Order Qty.",
                               "Bango_Dispatched Qty.",
                               "Bango_DR",
                               "Bango_DROO",
                               "Bango_TTS",
                               "Bango_BBT",
                               "Bango_BBT-Place",
                               "Bango_BBT-Place on invoice",
                               "Bango_BBT-Place off invoice",
                               "Bango_BBT-Price",
                               "Bango_CPP on invoice",
                               "Bango_CPP off invoice",
                               "Bango_BBT-Product",
                               "Bango_BBT-Pack",
                               "Bango_BBT-Proposition",
                               "Bango_BBT-Promotion",
                               "Bango_EOT",
                               "Bango_Brand & Marketing Investment",
                               "Bango_Brand & Marketing Investment Trade",
                               "Bango_Brand & Marketing Investment Consumer",
                               "Bango_Promotion Team Cost Trade",
                               "Bango_Promotion Team Cost Consumer",
                               "Bango_Promotion Team Cost",
                               "Bango_Promotion Packaging Material Cost Trade",
                               "Bango_Promotion Packaging Material Cost Consumer",
                               "Bango_Promotion Communication Material Cost Trade",
                               "Bango_Promotion Communication Material Cost Consumer",
                               "Bango_Promotion Communication Material Cost",
                               "Bango_Promo Agency Remuneration Fees & Commissions Consumer",
                               "Bango_Promotional Expenses",
                               "Bango_Promotion Packaging Material Cost",
                               "Bango_Promotion Repacking Cost",
                               "Bango_Promotion Repacking Cost Trade",
                               "Bango_Promotion Repacking Cost Consumer",
                               "Bango_Promo Samples, Gifts and Incentive Costs",
                               "Bango_Promo Samples, Gifts and Incentive Costs Consumer",
                               "Bango_Promo Agency Remun Fees & Commissions",
                               "Bango_Promo Agency Remuneration Fees & Commissions Trade")

colnames(total.bango.mt.r1) <- total.bango.mt.header.new

total.bango.mt.final <- total.bango.mt.r1

total.bango.mt.final <- format(total.bango.mt.final,scientific =F)

writexl::write_xlsx(total.bango.mt.final,paste0(directory.path,"/formated_total_bango_mt.xlsx"),format_headers = T)

############################################################## Total ID Bango (GT+MT)#########################################################
#perform the data blending for the psales source
psales.total.bango <-  psales.bango.df1
#select the national column from the psales total bango dataframe 

psales.total.bango%<>%mutate(psales_KG=`Sales Quantity`*Conversion,psales_Tonns=psales_KG/1000)
#group by observations and summarise the KPI's based on the pkey {Month,Brand,Channel}
psales.bango.aggregate <- psales.total.bango%>%
  group_by(`Fiscal year/period`,BRAND,National)%>%
  summarise(`Gross Sales Value (GSV)`=sum(`Gross Sales Value (GSV)`),`Sales Quantity`=sum(`Sales Quantity`),NIV=sum(NIV),psales_KG=sum(psales_KG),psales_Tonns=sum(psales_Tonns))
#perform the month mapping for the data frame values in the given period column in the psales.bango.aggregate data frame
psales.bango.aggregate <- merge.data.frame(psales.bango.aggregate,month.mapping,by.x = "Fiscal year/period",by.y ="Primary sales" )
psales.bango.aggregate%<>%select(-c(`Sec stock_DROO_Penetration`,TTS_BMI,`Secondary sales`,`Fiscal year/period`))
psales.bango.aggregate%<>%select(c(Month,BRAND,National,`Gross Sales Value (GSV)`,`Sales Quantity`,NIV,psales_KG,psales_Tonns))
#rename the column names as per the output file for the psales bango
psales.bango.aggregate <- setnames(psales.bango.aggregate,old = c("Gross Sales Value (GSV)","Sales Quantity","NIV","psales_KG","psales_Tonns"),new=c("GSV","Sales Qty","NIV","Sales Qty (Kg)","Sales Qty(Tonns)"))
psales.total.bango <- psales.bango.aggregate
#perform the data blending for the sstock blending 
sstock.total.bango <- sstock.bango.df1
sstock.total.bango <- sstock.total.bango%>%
  group_by(`Calendar Year/Month`,BRAND,National)%>%
  summarise(`Secondary Stock Volume`=sum(`Secondary Stock Volume`),`Secondary Stock Value [@ DT Rate.]`=sum(`Secondary Stock Value [@ DT Rate.]`),sstock_KG=sum(sstock_KG),sstock_Tonns=sum(sstock_Tonns))

#perform the month mapping for the data frame values in the given period column in the sstock.total.aggregate data frame
sstock.total.bango <- merge.data.frame(x = sstock.total.bango,y = month.mapping,by.x ="Calendar Year/Month",by.y = "Sec stock_DROO_Penetration")
sstock.total.bango%<>%select(-c(`Calendar Year/Month`,`Primary sales`,`Secondary sales`,TTS_BMI))
sstock.total.bango%<>%select(c(Month,BRAND:sstock_Tonns))
#perform the data blending for the penetration source
penetration.total.bango <- penetration.blend.df
penetration.total.bango$BRAND <- "BANGO"
penetration.total.bango$National <- "Total Indonesia"
penetration.total.bango%<>%select(Month,BRAND,National,Penetration:`Penetration%`)
#perform the data blending for the Itrust source 
itrust.total.bango <- itrust.bango
#perform the group by and summarisation of the Itrust KPI's
itrust.total.bango <- itrust.total.bango%>%
  group_by(Month,BRAND,National)%>%
  summarise(ITRUST_LINE=sum(ITRUST_LINE),ITRUST_TOTAL=sum(ITRUST_TOTAL))%>%
  mutate(ITRUST_PERCENTAGE=ITRUST_LINE/ITRUST_TOTAL*100)
#perform the data blending for the TTS source
tts.total.bango <- id.tts.channel.mapping
tts.total.bango.aggregate <- tts.total.bango%>%
  group_by(`Fiscal year/period`,BRAND,National)%>%
  summarise(TTS=sum(TTS),BBT=sum(BBT),`BBT - Place`=sum(`BBT - Place`),`BBT - Place on invoice`=sum(`BBT - Place on invoice`),`BBT - Place off invoice`=sum(`BBT - Place off invoice`),`BBT - Price`=sum(`BBT - Price`),`CPP on invoice`=sum(`CPP on invoice`),`CPP off invoice`=sum(`CPP off invoice`),`BBT - Product`=sum(`BBT - Product`),`BBT - Pack`=sum(`BBT - Pack`),`BBT - Proposition`=sum(`BBT - Proposition`),`BBT - Promotion`=sum(`BBT - Promotion`),EOT=sum(EOT))
#perform the month Mapping for the tts 

tts.total.bango.final <- merge.data.frame(x = tts.total.bango.aggregate,y = month.mapping,by.x = "Fiscal year/period" ,by.y = "TTS_BMI")
tts.total.bango.final %<>%select(-c(`Primary sales`,`Secondary sales`,`Sec stock_DROO_Penetration`,`Fiscal year/period`))
tts.total.bango.final %<>%select(c(Month,BRAND:EOT))

#perform the data blending for the BMI source
bmi.total.bango <- bmi.bango.mapped

bmi.total.bango.aggregate <- bmi.total.bango%>%
  group_by(`Fiscal year/period`,Brand,National)%>%
  summarise(X..Brand...Marketing.Investment=sum(X..Brand...Marketing.Investment),X..Brand...Marketing.Investment.Trade=sum(X..Brand...Marketing.Investment.Trade),X..Brand...Marketing.Investment.Consumer=sum(X..Brand...Marketing.Investment.Consumer),X..Promotional.Expenses=sum(X..Promotional.Expenses),X....Promotion.Packaging.Material.Cost=sum(X....Promotion.Packaging.Material.Cost),X......Promotion.Packaging.Material.Cost.Trade=sum(X......Promotion.Packaging.Material.Cost.Trade),X......Promotion.Packaging.Material.Cost.Consumer=sum(X......Promotion.Packaging.Material.Cost.Consumer),X....Promotion.Repacking.Cost=sum(X....Promotion.Repacking.Cost),X......Promotion.Repacking.Cost.Trade=sum(X....Promotion.Repacking.Cost),X......Promotion.Repacking.Cost.Trade=sum(X......Promotion.Repacking.Cost.Trade),X......Promotion.Repacking.Cost.Consumer=sum(X......Promotion.Repacking.Cost.Consumer),X....Promotion.Communication.Material.Cost=sum(X....Promotion.Communication.Material.Cost),X......Promotion.Communication.Material.Cost.Trade=sum(X......Promotion.Communication.Material.Cost.Trade),X......Promotion.Communication.Material.Cost.Consumer=sum(X......Promotion.Communication.Material.Cost.Consumer),X....Promo.Samples..Gifts.and.Incentive.Costs=sum(X....Promo.Samples..Gifts.and.Incentive.Costs),X......Promo.Samples..Gifts.and.Incentive.Costs.Consumer=sum(X......Promo.Samples..Gifts.and.Incentive.Costs.Consumer),X....Promotion.Team.Cost=sum(X....Promotion.Team.Cost),X......Promotion.Team.Cost.Trade=sum(X......Promotion.Team.Cost.Trade),X......Promotion.Team.Cost.Consumer=sum(X......Promotion.Team.Cost.Consumer),X....Promo.Agency.Remun.Fees...Commissions=sum(X....Promo.Agency.Remun.Fees...Commissions),X......Promo.Agency.Remuneration.Fees...Commissions.Trade=sum(X......Promo.Agency.Remuneration.Fees...Commissions.Trade),X......Promo.Agency.Remuneration.Fees...Commissions.Consumer=sum(X......Promo.Agency.Remuneration.Fees...Commissions.Consumer))


bmi.total.bango.aggregate1 <- bmi.total.bango.aggregate%>%
  dplyr::rename(`Brand & Marketing Investment`=X..Brand...Marketing.Investment,`Brand & Marketing Investment Trade`=X..Brand...Marketing.Investment.Trade,`Brand & Marketing Investment Consumer`=X..Brand...Marketing.Investment.Consumer,`Promotional Expenses`=X..Promotional.Expenses,`Promotion Packaging Material Cost`=X....Promotion.Packaging.Material.Cost,`Promotion Communication Material Cost Trade`=X......Promotion.Communication.Material.Cost.Trade,`Promotion Packaging Material Cost Consumer`=X......Promotion.Packaging.Material.Cost.Consumer,`Promotion Repacking Cost`=X....Promotion.Repacking.Cost,`Promotion Repacking Cost Trade`=X......Promotion.Repacking.Cost.Trade,`Promotion Repacking Cost Consumer`=X......Promotion.Repacking.Cost.Consumer,`Promotion Communication Material Cost`=X....Promotion.Communication.Material.Cost,`Promotion Communication Material Cost Consumer`=X......Promotion.Communication.Material.Cost.Consumer,`Promo Samples, Gifts and Incentive Costs` =X....Promo.Samples..Gifts.and.Incentive.Costs,`Promo Samples, Gifts and Incentive Costs Consumer`=X......Promo.Samples..Gifts.and.Incentive.Costs.Consumer,`Promotion Team Cost`=X....Promotion.Team.Cost,`Promotion Team Cost Trade`=X......Promotion.Team.Cost.Trade,`Promotion Team Cost Consumer`=X......Promotion.Team.Cost.Consumer,`Promo Agency Remun Fees & Commissions`=X....Promo.Agency.Remun.Fees...Commissions,`Promo Agency Remuneration Fees & Commissions Trade`=X......Promo.Agency.Remuneration.Fees...Commissions.Trade,`Promo Agency Remuneration Fees & Commissions Consumer`=X......Promo.Agency.Remuneration.Fees...Commissions.Consumer)

bmi.total.bango.aggregate1%<>%dplyr::rename(`Promotion Packaging Material Cost Trade`=X......Promotion.Packaging.Material.Cost.Trade)
#mapping the month for each observation 
bmi.total.bango.final <- merge.data.frame(x = bmi.total.bango.aggregate1,y = month.mapping,by.x ="Fiscal year/period",by.y = "TTS_BMI")
bmi.total.bango.final %<>%select(-c(`Primary sales`,`Secondary sales`,`Sec stock_DROO_Penetration`,`Fiscal year/period`))
bmi.total.bango.final %<>%select(c(Month,Brand:`Promo Agency Remuneration Fees & Commissions Consumer`))
bmi.total.bango.final <- setnames(bmi.total.bango.final,old = "Brand",new = "BRAND")
bmi.total.bango.final$BRAND <- "BANGO"

#perform the data blending for the DROO source 

droo.total.bango <- droo.bango.df
#perform the group by and summarization for the KPI's
droo.total.bango <- droo.total.bango%>%
  group_by(`Calendar Year/Month`,BRAND,National)%>%
  summarise(`OriginalOrder Qty`=sum(`OriginalOrder Qty`),`Final Customer Expected Order Qty`=sum(`Final Customer Expected Order Qty`),`Dispatched Qty`=sum(`Dispatched Qty`))
#perform the month mapping for the droo.total.bango df 

droo.total.bango.final <- merge.data.frame(x = droo.total.bango,y = month.mapping,by.x = "Calendar Year/Month",by.y="Sec stock_DROO_Penetration")
droo.total.bango.final %<>%select(-c(`Primary sales`,`Secondary sales`,TTS_BMI,`Calendar Year/Month`))
droo.total.bango.final%<>%select(c(Month,BRAND:`Dispatched Qty`))
droo.total.bango.final%<>%mutate(DR=(`Dispatched Qty`/`Final Customer Expected Order Qty`*100),DROO=(`Dispatched Qty`/`OriginalOrder Qty`*100))
#perform the data blending for the secondary sales for the total indonesia - Bango 
psales.total.bango.mt <- psales.bango.mt
ssales.total.bango.gt <- ssales.bango.gt
ssales.total.bango <- merge.data.frame(x = psales.total.bango.mt,y = ssales.total.bango.gt,by = c("Month","BRAND"))%>%
  select(-c("Channel.x","Channel.y"))
ssales.total.bango$National <- "Total Indonesia"
ssales.total.bango <- ssales.total.bango%>%
  mutate(`Secondary Sales Volume (PC)`=`Sales Qty`+`GT_Sec Sales Volume`,`Secondary Sales Volume (KG)`=`Sales Qty (Kg)`+`GT_Sec Sales Volume(Kg)`,`Secondary Volume Qty (Tonns)`=`Secondary Sales Volume (KG)`/1000,`Secondary sales Value`=`GT_Sec Sales Value`+NIV)
ssales.total.bango <- ssales.total.bango%>%
  dplyr::rename(`GT_Sec Sales Volume- From GT (PC)`=`GT_Sec Sales Volume`,`Primary Sales Qty- From MT (PC)`=`Sales Qty`,`GT_Sec Sales Volume(Kg)- From GT`=`GT_Sec Sales Volume(Kg)`,`Primary Sales Qty(Kg)- From MT`=`Sales Qty (Kg)`,`Primary Sales Qty(Kg)- From MT`=`Sales Qty (Kg)`,`Sec Sales Value- From GT`=`GT_Sec Sales Value`,`Primary NIV- From MT`=NIV)

ssales.total.bango <- ssales.total.bango[,-3]

ssales.total.bango.final <- ssales.total.bango%>%
  select(c(Month,BRAND,National,`GT_Sec Sales Volume- From GT (PC)`),`Primary Sales Qty- From MT (PC)`,`Secondary Sales Volume (PC)`,`GT_Sec Sales Volume(Kg)- From GT`,`Primary Sales Qty(Kg)- From MT`,`Secondary Sales Volume (KG)`,`Secondary Volume Qty (Tonns)`,`Sec Sales Value- From GT`,`Primary NIV- From MT`,`Secondary sales Value`)
#consolidate the KPI's across the sources

total.bango1 <- merge.data.frame(x = psales.total.bango,y = ssales.total.bango.final,by = c("Month","BRAND","National"))
total.bango2 <- merge.data.frame(x = total.bango1,y = sstock.total.bango,by=c("Month","BRAND","National"))
total.bango3 <- merge.data.frame(x=total.bango2,y = droo.total.bango.final,by = c("Month","BRAND","National"))
total.bango4 <- merge.data.frame(x = total.bango3,y = tts.total.bango.final,by = c("Month","BRAND","National"))
total.bango5 <- merge.data.frame(x = total.bango4,y = bmi.total.bango.final,by=c("Month","BRAND","National"))
total.bango6 <- merge.data.frame(x = total.bango5,y=itrust.total.bango,by=c("Month","BRAND","National"),all.x = T)
total.bango7 <- merge.data.frame(x = total.bango6,y=penetration.total.bango,by=c("Month","BRAND","National"))

total.bango7$`Sales Organization` <- as.character(9001)
total.bango7$Category <- "SAVOURY"
total.bango7$Sector <- "SEASONINGS"
total.bango7 <- total.bango7%>%
  select(c(Month,`Sales Organization`,Category,Sector,BRAND,National,GSV:`Penetration%`))

total.bango.headers <- c("Month",
                         "Sales Organization",
                         "Category",
                         "Sector",
                         "Brand",
                         "Market",
                         "Bango_GSV",
                         "Bango_Sales Qty",
                         "Bango_NIV",
                         "Bango_Sales Qty (Kg)",
                         "Bango_Sales Qty(Tonns)",
                         "Bango GT_Sec Sales Volume- From GT (PC)",
                         "Bango_Primary Sales Qty- From MT (PC)",
                         "Bango_Secondary Sales Volume (PC)",
                         "Bango GT_Sec Sales Volume(Kg)- From GT",
                         "Primary Sales Qty(Kg)- From MT",
                         "Bango_Secondary Sales Volume (KG)",
                         "Bango_Secondary Volume Qty (Tonns)",
                         "Bango GT_Sec Sales Value- From GT",
                         "Bango_Primary NIV- From MT",
                         "Bango_Secondary sales Value",
                         "Bango_Secondary Stock Volume",
                         "Bango_Secondary Stock Value [@ DT Rate.]",
                         "Bango_Secondary Stock Volume(Kg)",
                         "Bango_Secondary Stock Volume(Tonns)",
                         "Bango_OriginalOrder Qty.",
                         "Bango_Final Customer Expected Order Qty.",
                         "Bango_Dispatched Qty.",
                         "Bango_DR",
                         "Bango_DROO",
                         "Bango_TTS",
                         "Bango_BBT",
                         "Bango_BBT-Place",
                         "Bango_BBT-Place on invoice",
                         "Bango_BBT-Place off invoice",
                         "Bango_BBT-Price",
                         "Bango_CPP on invoice",
                         "Bango_CPP off invoice",
                         "Bango_BBT-Product",
                         "Bango_BBT-Pack",
                         "Bango_BBT-Proposition",
                         "Bango_BBT-Promotion",
                         "Bango_EOT",
                         "Bango_Brand & Marketing Investment",
                         "Bango_Brand & Marketing Investment Trade",
                         "Bango_Brand & Marketing Investment Consumer",
                         "Bango_Promotion Team Cost Trade",
                         "Bango_Promotion Team Cost Consumer",
                         "Bango_Promotion Team Cost",
                         "Bango_Promotion Packaging Material Cost Trade",
                         "Bango_Promotion Packaging Material Cost Consumer",
                         "Bango_Promotion Communication Material Cost Trade",
                         "Bango_Promotion Communication Material Cost Consumer",
                         "Bango_Promotion Communication Material Cost",
                         "Bango_Promo Agency Remuneration Fees & Commissions Consumer",
                         "Bango_Promotional Expenses",
                         "Bango_Promotion Packaging Material Cost",
                         "Bango_Promotion Repacking Cost",
                         "Bango_Promotion Repacking Cost Trade",
                         "Bango_Promotion Repacking Cost Consumer",
                         "Bango_Promo Samples, Gifts and Incentive Costs",
                         "Bango_Promo Samples, Gifts and Incentive Costs Consumer",
                         "Bango_Promo Agency Remun Fees & Commissions",
                         "Bango_Promo Agency Remuneration Fees & Commissions Trade",
                         "Bango_I TRUST LINE",
                         "Bango_I TRUST TOTAL",
                         "Bango_I Trust %",
                         "Bango_Penetration",
                         "Category_Penetration",
                         "Bango_Penetration%")
colnames(total.bango7) <- total.bango.headers
total.final.bango <- total.bango7
total.final.bango$Market <- toupper(total.final.bango$Market)
total.final.bango <- format(total.final.bango,scientific=F)
total.final.bango[is.na(total.final.bango)] <- ""
writexl::write_xlsx(total.final.bango,paste0(directory.path,"/ID_total_bango.xlsx"))


########################################### Model setup for Region #####################################################################
region.list <- c("DT SUMATERA","DT CENTRAL & EAST JAVA","DT OI","DT WEST JAVA")
############################################################ psales (Region- GT) ##############################################
psales.region.df <- psales.bango.df1
# filter `out the regions named (DT SUMATERA, DT WEST JAVA, DT CENTRAL & EAST JAVA, DT OI)
psales.region.data <- psales.region.df%>%
  filter(`Account/Region` %in% region.list)
#aggregate the data based on the brand , fiscal year and region 
psales.region.aggregate <- psales.region.data%>%
  group_by(`Fiscal year/period`,BRAND,`Account/Region`)%>%
  summarise(`Gross Sales Value (GSV)`=sum(`Gross Sales Value (GSV)`),`Sales Quantity`=sum(`Sales Quantity`),NIV=sum(NIV),psales_KG=sum(psales_KG),psales_Tonns=sum(psales_Tonns))
#perform the month mapping for the data frame values in the given period column in the psales.bango.aggregate data frame
psales.region.aggregate <- merge.data.frame(psales.region.aggregate,month.mapping,by.x = "Fiscal year/period",by.y ="Primary sales" )
psales.region.aggregate%<>%select(-c(`Sec stock_DROO_Penetration`,TTS_BMI,`Secondary sales`,`Fiscal year/period`))
psales.region.aggregate%<>%select(c(Month,BRAND,`Account/Region`,`Gross Sales Value (GSV)`,`Sales Quantity`,NIV,psales_KG,psales_Tonns))
#rename the column names as per the output file
psales.region.aggregate <- setnames(psales.region.aggregate,old = c("Gross Sales Value (GSV)","Sales Quantity","NIV","psales_KG","psales_Tonns"),new=c("GSV","Sales Qty","NIV","Sales Qty (Kg)","Sales Qty(Tonns)"))
########################################################## SSales (Region - GT) ###########################################################
ssales.region.df <- ssales.bango.df1
ssales.region.data <- ssales.region.df%>%
  filter(`Account/Region` %in% region.list)
ssales.region.aggregate <- ssales.region.data%>%
  group_by(`Calendar Year/Month`,BRAND,`Account/Region`)%>%
  summarise(`Gross Sec Sales (TUR)`=sum(`Gross Sec Sales (TUR)`),`Sec Volume`=sum(`Sec Volume`),ssales_KG=sum(ssales_KG),ssales_Tonns=sum(ssales_Tonns))
#perform the month mapping for the data frame values in the given period column in the psales.bango.aggregate data frame
ssales.region.aggregate <- merge.data.frame(ssales.region.aggregate,month.mapping,by.x = "Calendar Year/Month",by.y = "Secondary sales")
ssales.region.aggregate%<>%select(-c(`Sec stock_DROO_Penetration`,TTS_BMI,`Primary sales`,`Calendar Year/Month`))
#rearrange the column as per the requirement
ssales.region.aggregate%<>%select(c(Month,BRAND,`Account/Region`,`Gross Sec Sales (TUR)`,`Sec Volume`,ssales_KG,ssales_Tonns))
ssales.region.aggregate <- setnames(ssales.region.aggregate,old = c("Gross Sec Sales (TUR)","Sec Volume","ssales_KG","ssales_Tonns"),new=c( "GT_Sec Sales Value","GT_Sec Sales Volume","GT_Sec Sales Volume(Kg)","Secondary Volume Qty (Tonns)"))
###################################################### SStock (Region-GT)##################################################################
sstock.region.df <- sstock.bango.df1
sstock.region.data <- sstock.region.df%>%
  filter(`Account/Region` %in% region.list)
sstock.region.aggregate <- sstock.region.data%>%
  group_by(`Calendar Year/Month`,BRAND,`Account/Region`)%>%
  summarise(`Secondary Stock Volume`=sum(`Secondary Stock Volume`),`Secondary Stock Value [@ DT Rate.]`=sum(`Secondary Stock Value [@ DT Rate.]`),sstock_KG=sum(sstock_KG),sstock_Tonns=sum(sstock_Tonns))
#perform the month mapping for the data frame values in the given period column in the psales.bango.aggregate data frame
sstock.region.aggregate <- merge.data.frame(sstock.region.aggregate,month.mapping,by.x = "Calendar Year/Month",by.y = "Sec stock_DROO_Penetration")
sstock.region.aggregate%<>%select(-c(`Primary sales`,TTS_BMI,`Primary sales`,`Calendar Year/Month`,`Secondary sales`))
sstock.region.aggregate%<>%select(Month,BRAND,`Account/Region`,`Secondary Stock Volume`,`Secondary Stock Value [@ DT Rate.]`,sstock_KG,sstock_Tonns)
sstock.region.aggregate <- setnames(sstock.region.aggregate,old = c("Secondary Stock Volume","Secondary Stock Value [@ DT Rate.]","sstock_KG","sstock_Tonns"),new=c( "Secondary Stock Volume","Secondary Stock Value [@ DT Rate.]","Secondary Stock Volume(Kg)","Secondary Stock Volume(Tonns)"))
#################################################### DROO (Region-GT)#########################################################################
droo.region.df <- droo.bango.df
droo.region.data <- droo.region.df%>%
  filter(`Account/Region`%in% region.list)
droo.region.aggregate <- droo.region.data%>%
  group_by(`Calendar Year/Month`,BRAND,`Account/Region`)%>%
  summarise(`OriginalOrder Qty`=sum(`OriginalOrder Qty`),`Final Customer Expected Order Qty`=sum(`Final Customer Expected Order Qty`),`Dispatched Qty`=sum(`Dispatched Qty`))
#calculate the DR and DROO for the aggregated months 
droo.region.aggregate%<>%mutate(DR=(`Dispatched Qty`/`Final Customer Expected Order Qty`*100),DROO=(`Dispatched Qty`/`OriginalOrder Qty`*100))
#perform the month mapping for the data frame values in the given period column in the psales.bango.aggregate data frame
droo.region.aggregate <- merge.data.frame(droo.region.aggregate,month.mapping,by.x = "Calendar Year/Month",by.y = "Sec stock_DROO_Penetration")
droo.region.aggregate%<>%select(-c(`Primary sales`,TTS_BMI,`Primary sales`,`Calendar Year/Month`,`Secondary sales`))
droo.region.aggregate%<>%select(Month,BRAND,`Account/Region`,`OriginalOrder Qty`,`Final Customer Expected Order Qty`,`Dispatched Qty`,DR,DROO)
################################################## Itrust ( Region-GT) ##########################################################################
itrust.region.df <- itrust.bango
#perform the aggregation of the KPI's based on the key of {Month,Brand,Account/Region}
itrust.region.data <- itrust.region.df%>%
  filter(`Account/Region` %in% region.list)
itrust.region.data <- itrust.region.data%>%
  group_by(Month,BRAND,`Account/Region`)%>%
  summarise(ITRUST_LINE=sum(ITRUST_LINE),ITRUST_TOTAL=sum(ITRUST_TOTAL))%>%
  mutate(ITRUST_PERCENTAGE=ITRUST_LINE/ITRUST_TOTAL*100)
############################################### TTS ( Region - GT- Mapped+Unmapped) ###################################################
tts.mapped.region.df <- id.tts.channel.mapping
tts.mapped.region.data <- tts.mapped.region.df%>%
  filter(str_detect(`Account/Region`,"DT")|`Account/Region`=="ALL")
tts.mapped.region.aggregate <- tts.mapped.region.data%>%
  group_by(`Fiscal year/period`,BRAND,`Account/Region`)%>%
  summarise(TTS=sum(TTS),BBT=sum(BBT),`BBT - Place`=sum(`BBT - Place`),`BBT - Place on invoice`=sum(`BBT - Place on invoice`),`BBT - Place off invoice`=sum(`BBT - Place off invoice`),`BBT - Price`=sum(`BBT - Price`),`CPP on invoice`=sum(`CPP on invoice`),`CPP off invoice`=sum(`CPP off invoice`),`BBT - Product`=sum(`BBT - Product`),`BBT - Pack`=sum(`BBT - Pack`),`BBT - Proposition`=sum(`BBT - Proposition`),`BBT - Promotion`=sum(`BBT - Promotion`),EOT=sum(EOT))
tts.unmapped.region.df <- id.tts.unmapped.rna1
tts.unmapped.region.data <- tts.unmapped.region.df%>%
  filter(str_detect(`Account/Region`,"DT")|`Account/Region`=="ALL")
#aggregate the unmapped tts data based on {Month,Brand,Account/Region}
tts.unmapped.region.aggregate <- tts.unmapped.region.data%>%
  group_by(`Fiscal year/period`,BRAND,`Account/Region`)%>%
  summarise(TTS=sum(TTS),BBT=sum(BBT),`BBT - Place`=sum(`BBT - Place`),`BBT - Place on invoice`=sum(`BBT - Place on invoice`),`BBT - Place off invoice`=sum(`BBT - Place off invoice`),`BBT - Price`=sum(`BBT - Price`),`CPP on invoice`=sum(`CPP on invoice`),`CPP off invoice`=sum(`CPP off invoice`),`BBT - Product`=sum(`BBT - Product`),`BBT - Pack`=sum(`BBT - Pack`),`BBT - Proposition`=sum(`BBT - Proposition`),`BBT - Promotion`=sum(`BBT - Promotion`),EOT=sum(EOT))
#perform the rbind for the mapped and the unmapped tts data 

tts.region.aggregate <- rbind(tts.mapped.region.aggregate,tts.unmapped.region.aggregate)
#perform the grouping of the total aggregate data and aggregate the KPI's

tts.region.aggregate <- tts.region.aggregate%>%
  group_by(`Fiscal year/period`,BRAND,`Account/Region`)%>%
  summarise(TTS=sum(TTS),BBT=sum(BBT),`BBT - Place`=sum(`BBT - Place`),`BBT - Place on invoice`=sum(`BBT - Place on invoice`),`BBT - Place off invoice`=sum(`BBT - Place off invoice`),`BBT - Price`=sum(`BBT - Price`),`CPP on invoice`=sum(`CPP on invoice`),`CPP off invoice`=sum(`CPP off invoice`),`BBT - Product`=sum(`BBT - Product`),`BBT - Pack`=sum(`BBT - Pack`),`BBT - Proposition`=sum(`BBT - Proposition`),`BBT - Promotion`=sum(`BBT - Promotion`),EOT=sum(EOT))
#perform the month mapping for the region wise aggregate data
tts.region.aggregate <- merge.data.frame(x = tts.region.aggregate,y=month.mapping,by.x = "Fiscal year/period","TTS_BMI",all.x = T)
tts.region.aggregate%<>%select(-c(`Primary sales`,`Secondary sales`,`Sec stock_DROO_Penetration`,`Fiscal year/period`))
tts.region.aggregate <- tts.region.aggregate%>%
  select(c(Month,BRAND:EOT))
############################################################## BMI ( Region-GT Mapped+Unmapped)####################################################
bmi.bango.region.mapped.aggregate <- bmi.bango.mapped%>%
  group_by(`Fiscal year/period`,Brand,`Account/Region`)%>%
  summarise(X..Brand...Marketing.Investment=sum(X..Brand...Marketing.Investment),X..Brand...Marketing.Investment.Trade=sum(X..Brand...Marketing.Investment.Trade),X..Brand...Marketing.Investment.Consumer=sum(X..Brand...Marketing.Investment.Consumer),X..Promotional.Expenses=sum(X..Promotional.Expenses),X....Promotion.Packaging.Material.Cost=sum(X....Promotion.Packaging.Material.Cost),X......Promotion.Packaging.Material.Cost.Trade=sum(X......Promotion.Packaging.Material.Cost.Trade),X......Promotion.Packaging.Material.Cost.Consumer=sum(X......Promotion.Packaging.Material.Cost.Consumer),X....Promotion.Repacking.Cost=sum(X....Promotion.Repacking.Cost),X......Promotion.Repacking.Cost.Trade=sum(X....Promotion.Repacking.Cost),X......Promotion.Repacking.Cost.Trade=sum(X......Promotion.Repacking.Cost.Trade),X......Promotion.Repacking.Cost.Consumer=sum(X......Promotion.Repacking.Cost.Consumer),X....Promotion.Communication.Material.Cost=sum(X....Promotion.Communication.Material.Cost),X......Promotion.Communication.Material.Cost.Trade=sum(X......Promotion.Communication.Material.Cost.Trade),X......Promotion.Communication.Material.Cost.Consumer=sum(X......Promotion.Communication.Material.Cost.Consumer),X....Promo.Samples..Gifts.and.Incentive.Costs=sum(X....Promo.Samples..Gifts.and.Incentive.Costs),X......Promo.Samples..Gifts.and.Incentive.Costs.Consumer=sum(X......Promo.Samples..Gifts.and.Incentive.Costs.Consumer),X....Promotion.Team.Cost=sum(X....Promotion.Team.Cost),X......Promotion.Team.Cost.Trade=sum(X......Promotion.Team.Cost.Trade),X......Promotion.Team.Cost.Consumer=sum(X......Promotion.Team.Cost.Consumer),X....Promo.Agency.Remun.Fees...Commissions=sum(X....Promo.Agency.Remun.Fees...Commissions),X......Promo.Agency.Remuneration.Fees...Commissions.Trade=sum(X......Promo.Agency.Remuneration.Fees...Commissions.Trade),X......Promo.Agency.Remuneration.Fees...Commissions.Consumer=sum(X......Promo.Agency.Remuneration.Fees...Commissions.Consumer))

bmi.bango.region.mapped.aggregate <- bmi.bango.region.mapped.aggregate[which(!is.na(bmi.bango.region.mapped.aggregate$`Account/Region`)),]


bmi.bango.region.mapped.aggregate1 <- bmi.bango.region.mapped.aggregate%>%
  dplyr::rename(`Brand & Marketing Investment`=X..Brand...Marketing.Investment,`Brand & Marketing Investment Trade`=X..Brand...Marketing.Investment.Trade,`Brand & Marketing Investment Consumer`=X..Brand...Marketing.Investment.Consumer,`Promotional Expenses`=X..Promotional.Expenses,`Promotion Packaging Material Cost`=X....Promotion.Packaging.Material.Cost,`Promotion Communication Material Cost Trade`=X......Promotion.Communication.Material.Cost.Trade,`Promotion Packaging Material Cost Consumer`=X......Promotion.Packaging.Material.Cost.Consumer,`Promotion Repacking Cost`=X....Promotion.Repacking.Cost,`Promotion Repacking Cost Trade`=X......Promotion.Repacking.Cost.Trade,`Promotion Repacking Cost Consumer`=X......Promotion.Repacking.Cost.Consumer,`Promotion Communication Material Cost`=X....Promotion.Communication.Material.Cost,`Promotion Communication Material Cost Consumer`=X......Promotion.Communication.Material.Cost.Consumer,`Promo Samples, Gifts and Incentive Costs` =X....Promo.Samples..Gifts.and.Incentive.Costs,`Promo Samples, Gifts and Incentive Costs Consumer`=X......Promo.Samples..Gifts.and.Incentive.Costs.Consumer,`Promotion Team Cost`=X....Promotion.Team.Cost,`Promotion Team Cost Trade`=X......Promotion.Team.Cost.Trade,`Promotion Team Cost Consumer`=X......Promotion.Team.Cost.Consumer,`Promo Agency Remun Fees & Commissions`=X....Promo.Agency.Remun.Fees...Commissions,`Promo Agency Remuneration Fees & Commissions Trade`=X......Promo.Agency.Remuneration.Fees...Commissions.Trade,`Promo Agency Remuneration Fees & Commissions Consumer`=X......Promo.Agency.Remuneration.Fees...Commissions.Consumer)

bmi.bango.region.mapped.aggregate1%<>%dplyr::rename(`Promotion Packaging Material Cost Trade`=X......Promotion.Packaging.Material.Cost.Trade)
bmi.mapped.bango.final <- merge.data.frame(x = bmi.bango.region.mapped.aggregate1,y = month.mapping,by.x ="Fiscal year/period",by.y = "TTS_BMI")
bmi.mapped.bango.final %<>%select(-c(`Primary sales`,`Secondary sales`,`Sec stock_DROO_Penetration`,`Fiscal year/period`))
bmi.mapped.bango.final %<>%select(c(Month,Brand:`Promo Agency Remuneration Fees & Commissions Consumer`))
bmi.mapped.bango.final$BRAND <- "BANGO"
bmi.mapped.bango.final <- bmi.mapped.bango.final%>%
  select(c(Month,BRAND,`Account/Region`:`Promo Agency Remuneration Fees & Commissions Consumer`))
#perform the BMI unmmapped data blending 
bmi.bango.region.unmapped <- id.bmi.unmapped.df1
unmapped.bmi.mapping.region <- unmapped.bmi.mapping[,c(1,4)]
bmi.unmapped.join <- merge.data.frame(x = bmi.bango.region.unmapped,y = unmapped.bmi.mapping.region,by.x = "key",by.y ="Key")
bmi.unmapped.join[22:42] <- sapply(bmi.unmapped.join[22:42],as.numeric)


#perform the aggregate of the unmapped join

bmi.unmapped.region.aggregate <- bmi.unmapped.join%>%
  group_by(`Fiscal year/period`,Brand_Name,`Account/Region`)%>%
  summarise(`Brand & Marketing Investment`=sum(`Brand & Marketing Investment`),`Brand & Marketing Investment Trade`=sum(`Brand & Marketing Investment Trade`),`Brand & Marketing Investment Consumer`=sum(`Brand & Marketing Investment Consumer`),`Promotional Expenses`=sum(`Promotional Expenses`),`Promotion Packaging Material Cost`=sum(`Promotion Packaging Material Cost`),`Promotion Packaging Material Cost Trade`=sum(`Promotion Repacking Cost Trade`),`Promotion Packaging Material Cost Consumer`=sum(`Promotion Packaging Material Cost Consumer`),`Promotion Repacking Cost`=sum(`Promotion Repacking Cost`),`Promotion Repacking Cost Trade`=sum(`Promotion Repacking Cost Trade`),`Promotion Repacking Cost Consumer`=sum(`Promotion Repacking Cost Consumer`),`Promotion Communication Material Cost`=sum(`Promotion Communication Material Cost`),`Promotion Communication Material Cost Trade`=sum(`Promotion Communication Material Cost Trade`),`Promotion Communication Material Cost Consumer`=sum(`Promotion Communication Material Cost Consumer`),`Promo Samples, Gifts and Incentive Costs`=sum(`Promo Samples, Gifts and Incentive Costs`),`Promo Samples, Gifts and Incentive Costs Consumer`=sum(`Promo Samples, Gifts and Incentive Costs Consumer`),`Promotion Team Cost`=sum(`Promotion Team Cost`),`Promotion Team Cost Trade`=sum(`Promotion Team Cost Trade`),`Promotion Team Cost Consumer`=sum(`Promotion Team Cost Consumer`),`Promo Agency Remun Fees & Commissions`=sum(`Promo Agency Remun Fees & Commissions`),`Promo Agency Remuneration Fees & Commissions Trade`=sum(`Promo Agency Remuneration Fees & Commissions Trade`),`Promo Agency Remuneration Fees & Commissions Consumer`=sum(`Promo Agency Remuneration Fees & Commissions Consumer`))

bmi.unmapped.region.aggregate <- setnames(bmi.unmapped.region.aggregate,old = "Brand_Name",new = "BRAND")
#perform the mapping for the months
bmi.unmapped.region.aggregate1 <- mapping(input = bmi.unmapped.region.aggregate,map = month.mapping,map_key1 = "Fiscal year/period",map_key2 = "TTS_BMI" )
bmi.unmapped.region.aggregate1%<>%select(-c(`Primary sales`,`Secondary sales`,`Sec stock_DROO_Penetration`,`Fiscal year/period`))
bmi.unmapped.region.aggregate1%<>%select(c(Month,BRAND:`Promo Agency Remuneration Fees & Commissions Consumer`))
bmi.unmapped.region.aggregate1$BRAND <- "BANGO"
#perform the rbind for the unmapped and the mapped BMI data 
bmi.region.aggregate <- rbind(bmi.mapped.bango.final,bmi.unmapped.region.aggregate1)
#perform the group by and aggregate the data
bmi.region.aggregate1 <- bmi.region.aggregate%>%
  group_by(Month,BRAND,`Account/Region`)%>%
  summarise(`Brand & Marketing Investment`=sum(`Brand & Marketing Investment`),`Brand & Marketing Investment Trade`=sum(`Brand & Marketing Investment Trade`),`Brand & Marketing Investment Consumer`=sum(`Brand & Marketing Investment Consumer`),`Promotional Expenses`=sum(`Promotional Expenses`),`Promotion Packaging Material Cost`=sum(`Promotion Packaging Material Cost`),`Promotion Packaging Material Cost Trade`=sum(`Promotion Packaging Material Cost Trade`),`Promotion Packaging Material Cost Consumer`=sum(`Promotion Packaging Material Cost Consumer`),`Promotion Repacking Cost`=sum(`Promotion Repacking Cost`),`Promotion Repacking Cost Trade`=sum(`Promotion Repacking Cost Trade`),`Promotion Repacking Cost Consumer`=sum(`Promotion Repacking Cost Consumer`),`Promotion Communication Material Cost`=sum(`Promotion Communication Material Cost`),`Promotion Communication Material Cost Trade`=sum(`Promotion Communication Material Cost Trade`),`Promotion Communication Material Cost Consumer`=sum(`Promotion Communication Material Cost Consumer`),`Promo Samples, Gifts and Incentive Costs`=sum(`Promo Samples, Gifts and Incentive Costs`),`Promo Samples, Gifts and Incentive Costs Consumer`=sum(`Promo Samples, Gifts and Incentive Costs Consumer`),`Promotion Team Cost`=sum(`Promotion Team Cost`),`Promotion Team Cost Trade`=sum(`Promotion Team Cost Trade`),`Promotion Team Cost Consumer`=sum(`Promotion Team Cost Consumer`),`Promo Agency Remun Fees & Commissions`=sum(`Promo Agency Remun Fees & Commissions`),`Promo Agency Remuneration Fees & Commissions Trade`=sum(`Promo Agency Remuneration Fees & Commissions Trade`),`Promo Agency Remuneration Fees & Commissions Consumer`=sum(`Promo Agency Remuneration Fees & Commissions Consumer`))
#filter out the 4 Account/Region 
bmi.region <- bmi.region.aggregate1%>%
  filter(str_detect(`Account/Region`,"DT")|`Account/Region`=="ALL")

#################################################### Penetration ( 3 Region+1 Region)#####################################

penetration_folder_path <- "C:/Users/goura/OneDrive/Desktop/ID_Jarvis/Python ID KT/Input_files/Penetration_Region"
penetration.bango.oi.path <- paste0(penetration_folder_path,"/GD ID PENETRATION BANGO OI_Region.xlsx")
penetration.bango.region.path <-  paste0(penetration_folder_path,"/GD ID PENETRATION BANGO_Region.xlsx")
penetration.totalcat.oi.path <- paste0(penetration_folder_path,"/GD ID PENETRATION Total Category OI_Region.xlsx")
penetration.totalcat.region.path <- paste0(penetration_folder_path,"/GD ID PENETRATION Total Category_Region.xlsx")
penetration.bango.oi.df <- as_tibble(read_xlsx(path = paste0(penetration.bango.oi.path),sheet = 1,skip = 14))
penetration.totalcat.oi.df <- as_tibble(read_xlsx(path = paste0(penetration.totalcat.oi.path),sheet = 1,skip = 14))
penetration.bango.region.df <- as_tibble(read_xlsx(path = paste0(penetration.bango.region.path),sheet = 1,skip = 14))
penetration.totalcat.region.df <- as_tibble(read_xlsx(path = paste0(penetration.totalcat.region.path),sheet = 1,skip = 14))
#clean the dataframes using clean function
penetration.bango.oi.df <- input_clean(penetration.bango.oi.df)
penetration.bango.region.df <- input_clean(penetration.bango.region.df)
penetration.bango.region.df <- penetration.bango.region.df[complete.cases(penetration.bango.region.df),]
penetration.totalcat.oi.df <- input_clean(penetration.totalcat.oi.df)
penetration.totalcat.region.df <- input_clean(penetration.totalcat.region.df)
penetration.totalcat.region.df <- penetration.totalcat.region.df[complete.cases(penetration.totalcat.region.df),]
#perform the data blending for the input df's
colnames(penetration.bango.oi.df) <- penetration.bango.oi.df[1,]
penetration.bango.oi.df <- penetration.bango.oi.df[-1,]
penetration.bango.oi.df <- penetration.bango.oi.df%>%
  slice(1)
penetration.bango.oi.df$"Account/Region" <- "DT OI"
penetration.bango.oi.df <- penetration.bango.oi.df[2:ncol(penetration.bango.oi.df)]
penetration.bango.oi.df %<>% select(c(`Account/Region`,everything()))

colnames(penetration.totalcat.oi.df) <- penetration.totalcat.oi.df[1,]
penetration.totalcat.oi.df <- penetration.totalcat.oi.df[-1,]
penetration.totalcat.oi.df <- penetration.totalcat.oi.df%>%
  slice(1)
penetration.totalcat.oi.df$"Account/Region" <- "DT OI"
penetration.totalcat.oi.df <- penetration.totalcat.oi.df[2:ncol(penetration.totalcat.oi.df)]
penetration.totalcat.oi.df%<>%select(c(`Account/Region`,everything()))
penetration.bango.oi.df <- gather(penetration.bango.oi.df,key = "Month", "Penetration_Region",2:ncol(penetration.bango.oi.df))
penetration.totalcat.oi.df <- gather(penetration.totalcat.oi.df,key = "Month","Penetration_Category",2:ncol(penetration.totalcat.oi.df))
penetration.1region <- merge.data.frame(penetration.bango.oi.df,penetration.totalcat.oi.df,by = c("Account/Region","Month"))
#perform the mapping for the 3 regions {DT SUMATERA,DT WEST JAVA,DT CENTRAL & EAST JAVA} for penetration.bango.region.mapped
penetration.bango.region.mapped <- merge.data.frame(penetration.bango.region.df,penetration.customer.region.mapping,by = "Local Sales Force 2(m.d.)\\Calendar Year/Month")
penetration.bango.region.mapped <- penetration.bango.region.mapped[which(penetration.bango.region.mapped$`Account/Region`!=0),]
penetration.bango.region.mapped %<>%select(-c(`Local Sales Force 2(m.d.)\\Calendar Year/Month`))
penetration.bango.region.mapped%<>%select(c(`Account/Region`,everything()))
#transpose the penetration.bango.region.mapped df
penetration.bango.region.mapped <- gather(penetration.bango.region.mapped,key = "Month","Penetration_Region",2:ncol(penetration.bango.region.mapped))
#perform the mapping for the 3 regions {DT SUMATERA,DT WEST JAVA,DT CENTRAL & EAST JAVA} for penetration.bango.region.mapped
penetration.totalcat.region.mapped <- merge.data.frame(penetration.totalcat.region.df,penetration.customer.region.mapping,by = "Local Sales Force 2(m.d.)\\Calendar Year/Month")
penetration.totalcat.region.mapped <- penetration.totalcat.region.mapped[which(penetration.totalcat.region.mapped$`Account/Region`!=0),]
penetration.totalcat.region.mapped %<>%select(-c(`Local Sales Force 2(m.d.)\\Calendar Year/Month`))
penetration.totalcat.region.mapped%<>%select(c(`Account/Region`,everything()))
#transpose the penetration.totalcat.region.mapped
penetration.totalcat.region.mapped <- gather(penetration.totalcat.region.mapped,key = "Month",Penetration_Category,2:ncol(penetration.totalcat.region.mapped))
#merge the penetration.bango.region.mapped and penetration.totalcat.region.mapped
penetration.3regions <- merge.data.frame(x = penetration.bango.region.mapped,y =penetration.totalcat.region.mapped,by = c("Account/Region","Month") )
#perform the rbind for the total 4 regions
penetration.region <- rbind(penetration.3regions,penetration.1region)
#calculate the penetration percentage 
penetration.region%<>%hablar::convert(num(Penetration_Region,Penetration_Category))
penetration.region.final <- penetration.region%>%
  mutate(penetration_percentage = Penetration_Region/Penetration_Category*100)
penetration.region.final$penetration_percentage <- round(penetration.region.final$penetration_percentage,0)
#perform the month mapping <month>-<year> for the penetration.region.final data frame 
penetration.region.final1 <- merge.data.frame(x = penetration.region.final,y = month.mapping,by.x = "Month",by.y = "Sec stock_DROO_Penetration")
penetration.region.final1 %<>%select(-c(Month,`Primary sales`,`Secondary sales`,TTS_BMI))
penetration.region.final1 <- setnames(penetration.region.final1,old = "Month.y",new="Month")
penetration.region.final1%<>%select(c(Month,`Account/Region`:penetration_percentage))
penetration.region.final1$BRAND <- "BANGO"

#######################################Consolidation( Bango - Region)#########################################
psales.region.aggregate;ssales.region.aggregate
sstock.region.aggregate; droo.region.aggregate
itrust.region.data;tts.region.aggregate
bmi.total.region.aggregate <- bmi.region;penetration.region.aggregate <- penetration.region.final1
#create the BMI Primary Key 

primary_key <- c("Month","BRAND","Account/Region")
region.total.data1 <- merge.data.frame(psales.region.aggregate,ssales.region.aggregate,by = primary_key,all = T)
region.total.data2 <- merge.data.frame(region.total.data1,sstock.region.aggregate,by=primary_key,all=T)
region.total.data3 <- merge.data.frame(region.total.data2,droo.region.aggregate,by=primary_key,all = T)
region.total.data4 <- merge.data.frame(region.total.data3,tts.region.aggregate,by=primary_key,all = T)
region.total.data5 <- merge.data.frame(region.total.data4,bmi.total.region.aggregate,by=primary_key,all = T)
region.total.data6 <- merge.data.frame(region.total.data5,penetration.region.aggregate,by=primary_key,all=T)
region.total.data7 <- merge.data.frame(region.total.data6,itrust.region.data,by=primary_key,all.x=T)
region.total.data7[is.na(region.total.data7)] <- ""
bango.region.data <- region.total.data7
bango.region.data%<>%arrange(`Account/Region`)
bango.region.data <- setnames(bango.region.data,old ="Account/Region" ,new = "Customer")

bango.region.data$`Sales Organization` <- as.character(9001)
bango.region.data$Category <- "SAVOURY"
bango.region.data$Sector <- "SEASONINGS"
bango.region.data$Channel <- "DISTRIBUTIVE TRADE"
bango.region.data%<>%select(c(Month,`Sales Organization`,Category,Sector,BRAND,Channel,Customer,everything()))
bango.region.headers <- c("Month",
                          "Sales Organization",
                          "Category",
                          "Sector",
                          "Brand",
                          "Channel",
                          "Customer",
                          "Bango_GSV",
                          "Bango_Sales Qty",
                          "Bango_NIV",
                          "Bango_Sales Qty(Kg)",
                          "Bango_Sales Qty(Tonns)",
                          "Bango_Sec Sales Value",
                          "Bango_Sec Sales Volume",
                          "Bango_Sec sales Volume (Kg)",
                          "Bango_Sec sales Volume (Tonns)",
                          "Bango_Secondary Stock Volume",
                          "Bango_Secondary Stock Value [@ DT Rate.]",
                          "Bango_Secondary stock volume(Kg)",
                          "Bango_Secondary stock volume(Tonns)",
                          "Bango_Original Order Qty.",
                          "Bango_Final Customer Expected Order Qty.",
                          "Bango_Dispatched Qty.",
                          "Bango_DR",
                          "Bango_DROO",
                          "Bango_TTS",
                          "Bango_BBT",
                          "Bango_BBT-Place",
                          "Bango_BBT-Place on invoice",
                          "Bango_BBT-Place off invoice",
                          "Bango_BBT-Price",
                          "Bango_CPP on invoice",
                          "Bango_CPP off invoice",
                          "Bango_BBT-Product",
                          "Bango_BBT-Pack",
                          "Bango_BBT-Proposition",
                          "Bango_BBT-Promotion",
                          "Bango_EOT",
                          "Bango_Brand & Marketing Investment",
                          "Bango_Brand & Marketing Investment Trade",
                          "Bango_Brand & Marketing Investment Consumer",
                          "Bango_Promotion Team Cost Trade",
                          "Bango_Promotion Team Cost Consumer",
                          "Bango_Promotion Team Cost",
                          "Bango_Promotion Packaging Material Cost Trade",
                          "Bango_Promotion Packaging Material Cost Consumer",
                          "Bango_Promotion Communication Material Cost Trade",
                          "Bango_Promotion Communication Material Cost Consumer",
                          "Bango_Promotion Communication Material Cost",
                          "Bango_Promo Agency Remuneration Fees & Commissions Consumer",
                          "Bango_Promotional Expenses",
                          "Bango_Promotion Packaging Material Cost",
                          "Bango_Promotion Repacking Cost",
                          "Bango_Promotion Repacking Cost Trade",
                          "Bango_Promotion Repacking Cost Consumer",
                          "Bango_Promo Samples, Gifts and Incentive Costs",
                          "Bango_Promo Samples, Gifts and Incentive Costs Consumer",
                          "Bango_Promo Agency Remun Fees & Commissions",
                          "Bango_Promo Agency Remuneration Fees & Commissions Trade",
                          "Bango_Internal Penetration",
                          "Category_Internal Penetration",
                          "Bango_Internal penetration%",
                          "Bango_I trust Line",
                          "Bango_I trust Total",
                          "Bango_I trust %"
)

colnames(bango.region.data) <- bango.region.headers


#seggregate the data into 4 regions with All as the common component in Customer Variable
#{cej -> All Central & East Java & All ; oi -> All Oi & All ; sumtra -> All Sumatra & All ; wj ->  All West java & All }
bango.region.cej <- bango.region.data%>%
  filter(str_detect(Customer,"DT CENTRAL & EAST JAVA")|Customer=="ALL")
bango.region.cej <- format(bango.region.cej,scientific=F)
bango.region.oi <- bango.region.data%>%
  filter(str_detect(Customer,"DT OI")|Customer=="ALL")
bango.region.oi <- format(bango.region.oi,scientific=F)
bango.region.sumatra <- bango.region.data%>%
  filter(str_detect(Customer,"DT SUMATERA")|Customer=="ALL")
bango.region.sumatra <- format(bango.region.sumatra,scientific=F)
bango.region.wj <- bango.region.data%>%
  filter(str_detect(Customer,"DT WEST JAVA")|Customer=="ALL")
bango.region.wj <- format(bango.region.wj,scientific=F)
writexl::write_xlsx(bango.region.cej,paste0(output.path,"\\ID_Central&EastJava_Bango.xlsx"))
writexl::write_xlsx(bango.region.oi,paste0(output.path,"\\ID_OtherIsland_Bango.xlsx"))
writexl::write_xlsx(bango.region.sumatra,paste0(output.path,"\\ID_Sumatera_Bango.xlsx"))
writexl::write_xlsx(bango.region.wj,paste0(output.path,"\\ID_WestJava_Bango.xlsx"))
################################################################################## Pack group level(Psales)########################################################
#assign the required data to psales.packgroup.input variable
psales.packgroup.input <- psales.bango
#filter out the observations for which the requiered packgroup is !=0 and the channel has "DT" and select the required columns by calling packgroup_selectdf()
psales.packgroup.df <- packgroup_selectdf(psales.packgroup.input)
#calculate the derived KPI's for the Psales - Sales Quantity
psales.packgroup.df%<>%mutate(psales_KG=`Sales Quantity`*Conversion,psales_Tonns=psales_KG/1000)
#perform the grouping and aggregate of the KPI's
psales.packgroup.aggregate <- psales.packgroup.df%>%
  group_by(`Fiscal year/period`,`REQUIRED PACKGROUP`,`Account/Region`)%>%
  summarise(`Gross Sales Value (GSV)`=sum(`Gross Sales Value (GSV)`),`Sales Quantity`=sum(`Sales Quantity`),NIV=sum(NIV),psales_KG=sum(psales_KG),psales_Tonns=sum(psales_Tonns))
#perform the month mapping <month>-<year>
psales.packgroup.aggregate <- merge.data.frame(psales.packgroup.aggregate,month.mapping,by.x = "Fiscal year/period",by.y ="Primary sales" )
psales.packgroup.aggregate%<>%select(c(Month,`REQUIRED PACKGROUP`:psales_Tonns))
psales.packgroup.aggregate <- setnames(psales.packgroup.aggregate,old = c("REQUIRED PACKGROUP","Account/Region","Gross Sales Value (GSV)","Sales Quantity","NIV","psales_KG","psales_Tonns"),new=c("Product","Customer","GSV","Sales Qty","NIV","Sales Qty (Kg)","Sales Qty(Tonns)"))

####################################################################### pack group level (ssales)######################################################
#assign the required data to ssales.packgroup.input variable
ssales.packgroup.input <- ssales.bango
#filter out the observations for which the requiered packgroup is !=0 and the channel has "DT" and select the required columns by calling packgroup_selectdf()
ssales.packgroup.df <- packgroup_selectdf(ssales.packgroup.input)
#calculate the derived KPI's for the ssales -salesquantity
ssales.packgroup.df%<>%mutate(ssales_KG=`Sec Volume`*Conversion,ssales_Tonns=ssales_KG/1000)
#perform the grouping and aggregate of the KPI's
ssales.packgroup.aggregate <- ssales.packgroup.df%>%
  group_by(`Calendar Year/Month`,`REQUIRED PACKGROUP`,`Account/Region`)%>%
  summarise(`Gross Sec Sales (TUR)`=sum(`Gross Sec Sales (TUR)`),`Sec Volume`=sum(`Sec Volume`),ssales_KG=sum(ssales_KG),ssales_Tonns=sum(ssales_Tonns))
#perform the month mapping for the data frame values in the given period column in the ssales.packgroup.aggr
ssales.packgroup.aggregate <- merge.data.frame(ssales.packgroup.aggregate,month.mapping,by.x = "Calendar Year/Month",by.y ="Secondary sales" )
ssales.packgroup.aggregate%<>%select(c(Month,`REQUIRED PACKGROUP`:ssales_Tonns))
#setname for the columns as per the output
ssales.packgroup.aggregate <- setnames(ssales.packgroup.aggregate,old = c("REQUIRED PACKGROUP","Account/Region"),new=c("Product","Customer"))
####################################################################### pack group level (sstock)#######################################################
#assign the required data to sstock.packgroup,input variable
sstock.packgroup.input <- sstock.bango
#filter out the observations for which the requiered packgroup is !=0 and the channel has "DT" and select the required columns by calling packgroup_selectdf()
sstock.packgroup.df <- packgroup_selectdf(sstock.packgroup.input)
#calculate the derived KPI's for the sstock quantity
sstock.packgroup.df%<>%mutate(sstock_KG=`Secondary Stock Volume`*Conversion,sstock_Tonns=sstock_KG/1000)
sstock.packgroup.df <- setnames(sstock.packgroup.df,c("Material Desc.x"),c("Material Desc"))
#perform the grouping and aggregate of the KPI's
sstock.packgroup.aggregate <- sstock.packgroup.df%>%
  group_by(`Calendar Year/Month`,`REQUIRED PACKGROUP`,`Account/Region`)%>%
  summarise(`Secondary Stock Volume`=sum(`Secondary Stock Volume`),`Secondary Stock Value [@ DT Rate.]`=sum(`Secondary Stock Value [@ DT Rate.]`),sstock_KG=sum(sstock_KG),sstock_Tonns=sum(sstock_Tonns))
sstock.packgroup.aggregate <- setnames(sstock.packgroup.aggregate,old = c("Secondary Stock Volume","Secondary Stock Value [@ DT Rate.]","sstock_KG","sstock_Tonns"),new=c( "Secondary Stock Volume","Secondary Stock Value [@ DT Rate.]","Secondary Stock Volume(Kg)","Secondary Stock Volume(Tonns)"))
#perform the month mapping for the data frame values in the given period column in the sstock.packgroup.aggregate data frame
sstock.packgroup.aggregate <- merge.data.frame(sstock.packgroup.aggregate,month.mapping,by.x ="Calendar Year/Month",by.y ="Sec stock_DROO_Penetration")
sstock.packgroup.aggregate%<>%select(-c(`Primary sales`,`Secondary sales`,TTS_BMI,`Calendar Year/Month`))
sstock.packgroup.aggregate <- sstock.packgroup.aggregate%>%
  dplyr::rename(Product=`REQUIRED PACKGROUP`,Customer=`Account/Region`)
sstock.packgroup.aggregate%<>%select(c(Month,Product,Customer,`Secondary Stock Value [@ DT Rate.]`,`Secondary Stock Volume`,`Secondary Stock Volume(Kg)`,`Secondary Stock Volume(Tonns)`))
######################################################################### pack group (DROO)##############################################################
#assign the required data to droo.packgroup.input variable
droo.packgroup.input <- droo.bango
#filter out the observations for which the requiered packgroup is !=0 and the channel has "DT" and select the required columns by calling packgroup_selectdf()
droo.packgroup.df <- packgroup_selectdf(droo.packgroup.input)
#perform the grouping and aggregate of the KPI's
droo.packgroup.aggregate <- droo.packgroup.df%>%
  group_by(`Calendar Year/Month`,`REQUIRED PACKGROUP`,`Account/Region`)%>%
  summarise(`OriginalOrder Qty`=sum(`OriginalOrder Qty`),`Final Customer Expected Order Qty`=sum(`Final Customer Expected Order Qty`),`Dispatched Qty`=sum(`Dispatched Qty`))
#calculate the DR and DROO dervied KPI's
droo.packgroup.aggregate%<>%mutate(DR=(`Dispatched Qty`/`Final Customer Expected Order Qty`*100),DROO=(`Dispatched Qty`/`OriginalOrder Qty`*100))
#perform the month mapping for the data frame values in the given period column in the 
droo.packgroup.aggregate <- merge.data.frame(droo.packgroup.aggregate,month.mapping,by.x ="Calendar Year/Month",by.y ="Sec stock_DROO_Penetration")
droo.packgroup.aggregate%<>%select(-c(`Primary sales`,`Secondary sales`,TTS_BMI,`Calendar Year/Month`))
droo.packgroup.aggregate%<>%select(c(Month,`REQUIRED PACKGROUP`:DROO))
droo.packgroup.aggregate%<>%select(c(Month,`REQUIRED PACKGROUP`:`Dispatched Qty`,DROO,DR))
droo.packgroup.aggregate <- droo.packgroup.aggregate%>%
  dplyr::rename(Product=`REQUIRED PACKGROUP`,Customer=`Account/Region`)

######################################################################### packgroup (I trust)###########################################################
itrust.packgroup.input <- itrust.bango
itrust.packgroup.df <- itrust.packgroup.input%>%
  filter(`REQUIRED PACKGROUP`!=0)%>%
  filter(str_detect(`Account/Region`,"DT"))%>%
  filter(`Account/Region`!=0)

itrust.packgroup.aggregate <- itrust.packgroup.df%>%
  group_by(Month,`REQUIRED PACKGROUP`,`Account/Region`)%>%
  summarise(ITRUST_LINE=sum(ITRUST_LINE),ITRUST_TOTAL=sum(ITRUST_TOTAL))%>%
  mutate(ITRUST_PERCENTAGE=ITRUST_LINE/ITRUST_TOTAL*100)
itrust.packgroup.aggregate <- itrust.packgroup.aggregate%>%
  dplyr::rename(Product=`REQUIRED PACKGROUP`,Customer=`Account/Region`)
######################################################################## packgroup(TTS-mapped)#################################################################
#load the mapped tts data into tts.packgroup.mapped.input
tts.packgroup.mapped.input <- id.tts.mapped.bango
#select the required observations for which the packgroup !=0 
tts.packgroup.mapped.df <- tts.packgroup.mapped.input%>%
  filter(`REQUIRED PACKGROUP`!=0)
banner.packgroup.mapping <- banner.mapping%>%
  select(c(`Banner code`,`Account/Region`))
#perform the mapping for the banner
id.tts.packgroup.mapping <- mapping(input=tts.packgroup.mapped.df,map =banner.packgroup.mapping, map_key1 = "Banner",map_key2 = "Banner code")
#filter out the observations for which channel is not in region list
tts.packgroup.mapped.data <- id.tts.packgroup.mapping%>%
  filter(str_detect(`Account/Region`,"DT")|`Account/Region`=="ALL")
#perform the grouping and aggregation of the KPI's
tts.packgroup.mapped.aggregate <- tts.packgroup.mapped.data%>%
  group_by(`Fiscal year/period`,`REQUIRED PACKGROUP`,`Account/Region`)%>%
  summarise(TTS=sum(TTS),BBT=sum(BBT),`BBT - Place`=sum(`BBT - Place`),`BBT - Place on invoice`=sum(`BBT - Place on invoice`),`BBT - Place off invoice`=sum(`BBT - Place off invoice`),`BBT - Price`=sum(`BBT - Price`),`CPP on invoice`=sum(`CPP on invoice`),`CPP off invoice`=sum(`CPP off invoice`),`BBT - Product`=sum(`BBT - Product`),`BBT - Pack`=sum(`BBT - Pack`),`BBT - Proposition`=sum(`BBT - Proposition`),`BBT - Promotion`=sum(`BBT - Promotion`),EOT=sum(EOT))
##################################################################### Packgroup (TTS-Unmapped)#############################################################
#load the unmapped tts data into tts.packgroup.unmapped.input
tts.packgroup.unmapped.input <- unmapped.material.tts.mapping
tts.packgroup.unmapped.df <- tts.packgroup.unmapped.input%>%
  select(-c(`Material Desc`:`BASEPACK NO`,`PACKGROUP NAME`))
tts.packgroup.unmapped.df <- tts.packgroup.unmapped.df%>%
  filter(`REQUIRED PACKGROUP`!=0)
tts.packgroup.unmapped.df <- tts.packgroup.unmapped.df%>%
  unite(col = "key",c("Banner","Trade Format Level 2","Local Sales Force 2(m.d.)","Local Sales Force 3(m.d.)","Key Customer Level3"),sep="",remove = F)

unmapped.packgroup.tts.mapping <- unmapped.tts.mapping%>%
  select(c(Key,`Account/Region`))
tts.packgroup.unmapped.df <- merge.data.frame(tts.packgroup.unmapped.df,unmapped.packgroup.tts.mapping,by.x = "key",by.y = "Key")
#filter out the observations for which the account starts with "DT" and "All"
tts.packgroup.unmapped.df <- tts.packgroup.unmapped.df%>%
  filter(`Account/Region`!=0)
tts.packgroup.unmapped.df %<>%filter(str_detect(`Account/Region`,"DT") | `Account/Region` =="ALL")

tts.packgroup.unmapped.df%<>%convert(num(TTS:EOT))

tts.packgroup.unmapped.aggregate <- tts.packgroup.unmapped.df%>%
  group_by(`Fiscal year/period`,`REQUIRED PACKGROUP`,`Account/Region`)%>%
  summarise(TTS=sum(TTS),BBT=sum(BBT),`BBT - Place`=sum(`BBT - Place`),`BBT - Place on invoice`=sum(`BBT - Place on invoice`),`BBT - Place off invoice`=sum(`BBT - Place off invoice`),`BBT - Price`=sum(`BBT - Price`),`CPP on invoice`=sum(`CPP on invoice`),`CPP off invoice`=sum(`CPP off invoice`),`BBT - Product`=sum(`BBT - Product`),`BBT - Pack`=sum(`BBT - Pack`),`BBT - Proposition`=sum(`BBT - Proposition`),`BBT - Promotion`=sum(`BBT - Promotion`),EOT=sum(EOT))

####################################################################### packgroup ( TTS - Mapped + Unmapped) ###############################################

tts.packgroup.final <- rbind(tts.packgroup.mapped.aggregate,tts.packgroup.unmapped.aggregate)
tts.packgroup.aggregate <- tts.packgroup.final%>%
  group_by(`Fiscal year/period`,`REQUIRED PACKGROUP`,`Account/Region`)%>%
  summarise(TTS=sum(TTS),BBT=sum(BBT),`BBT - Place`=sum(`BBT - Place`),`BBT - Place on invoice`=sum(`BBT - Place on invoice`),`BBT - Place off invoice`=sum(`BBT - Place off invoice`),`BBT - Price`=sum(`BBT - Price`),`CPP on invoice`=sum(`CPP on invoice`),`CPP off invoice`=sum(`CPP off invoice`),`BBT - Product`=sum(`BBT - Product`),`BBT - Pack`=sum(`BBT - Pack`),`BBT - Proposition`=sum(`BBT - Proposition`),`BBT - Promotion`=sum(`BBT - Promotion`),EOT=sum(EOT))
tts.packgroup.aggregate.final <- tts.packgroup.aggregate%>%
  arrange(`Account/Region`)
#perform the <month>-<year> mapping for the final tts packgroup final df 
tts.packgroup.aggregate.final <- merge.data.frame(tts.packgroup.aggregate.final,month.mapping,by.x = "Fiscal year/period" ,by.y ="TTS_BMI" )
tts.packgroup.aggregate.final <- tts.packgroup.aggregate.final%>%
  select(-c(`Primary sales`:`Sec stock_DROO_Penetration`,`Fiscal year/period`))
tts.packgroup.aggregate.final%<>%select(c(Month,`REQUIRED PACKGROUP`:EOT))
tts.packgroup.aggregate.final <- tts.packgroup.aggregate.final%>%
  dplyr::rename(Product=`REQUIRED PACKGROUP`,Customer=`Account/Region`)

######################################################################### Packgroup ( BMI )################################################
bmi.packgroup.input <- bmi.region
packgroup.list <- unique(tts.packgroup.aggregate.final$`REQUIRED PACKGROUP`)
#create and replicate the data for the 6 pack group named w.r.t abbrevations are {bmbp - "Bango Manis Big Pouch; bmbs - "Bango Manis Big Sachet ; bmmb - Bango Manis Medium Bottle ; bmmp - Bango Manis Medium Pouch ; bmsb - "Bango Manis Small Bottle} ; bmsp - "bango Manis small pouch
bmi.packgroup.bmbp <- bmi.packgroup.input;bmi.packgroup.bmbs <- bmi.packgroup.input
bmi.packgroup.bmmb <- bmi.packgroup.input;bmi.packgroup.bmmp <- bmi.packgroup.input
bmi.packgroup.bmsb <- bmi.packgroup.input;bmi.packgroup.bmsp <- bmi.packgroup.input
#replicate across the packgroup for 4 regions - 6*4 =  24 output df
bmi.packgroup.bmbp$Required_Packgroup <- "Bango Manis Big Pouch"
bmi.packgroup.bmbp <-bmi.packgroup.bmbp%>%
  ungroup()%>%
  select(-BRAND)
bmi.packgroup.bmbs$Required_Packgroup <- "Bango Manis Big Sachet"
bmi.packgroup.bmbs <-bmi.packgroup.bmbs%>%
  ungroup()%>%
  select(-BRAND)

bmi.packgroup.bmmb$Required_Packgroup <- "Bango Manis Medium Bottle"
bmi.packgroup.bmmb <-bmi.packgroup.bmmb%>%
  ungroup()%>%
  select(-BRAND)

bmi.packgroup.bmmp$Required_Packgroup <- "Bango Manis Medium Pouch"
bmi.packgroup.bmmp <-bmi.packgroup.bmmp%>%
  ungroup()%>%
  select(-BRAND)

bmi.packgroup.bmsb$Required_Packgroup <- "Bango Manis Small Bottle"
bmi.packgroup.bmsb <-bmi.packgroup.bmsb%>%
  ungroup()%>%
  select(-BRAND)
bmi.packgroup.bmsp$Required_Packgroup <- "Bango Manis Small Pouch"
bmi.packgroup.bmsp <-bmi.packgroup.bmsp%>%
  ungroup()%>%
  select(-BRAND)



####################################################################### Packgroup ( penetration) ########################################################
#store the folder path for the penetration packgroup and get the list of penetration files 
penetration_packgroup_folder <- "C:/Users/goura/OneDrive/Desktop/ID_Jarvis/Python ID KT/Input_files/penetration_packgroup/penetration_packgroup_input"
file_names <- list.files(path = penetration_packgroup_folder)
#store the files data into variables with packgroup abbrevation 
penetration_bmbp_oi <- tibble::as_tibble(read_xlsx(paste0(penetration_packgroup_folder,"/",file_names[1]),sheet = 1,skip = 14))
penetration_bmbp <- tibble::as_tibble(read_xlsx(paste0(penetration_packgroup_folder,"/",file_names[2]),sheet = 1,skip = 14))
penetration_bmbs_oi <- tibble::as_tibble(read_xlsx(paste0(penetration_packgroup_folder,"/",file_names[3]),sheet = 1,skip = 14))
penetration_bmbs <- tibble::as_tibble(read_xlsx(paste0(penetration_packgroup_folder,"/",file_names[4]),sheet = 1,skip = 14))
penetration_bmmb_oi <- tibble::as_tibble(read_xlsx(paste0(penetration_packgroup_folder,"/",file_names[5]),sheet = 1,skip = 14))
penetration_bmmb <- tibble::as_tibble(read_xlsx(paste0(penetration_packgroup_folder,"/",file_names[6]),sheet = 1,skip = 14))
penetration_bmmp_oi <- tibble::as_tibble(read_xlsx(paste0(penetration_packgroup_folder,"/",file_names[7]),sheet = 1,skip = 14))
penetration_bmmp <- tibble::as_tibble(read_xlsx(paste0(penetration_packgroup_folder,"/",file_names[8]),sheet = 1,skip = 14))
penetration_bmsb_oi <- tibble::as_tibble(read_xlsx(paste0(penetration_packgroup_folder,"/",file_names[9]),sheet = 1,skip = 14))
penetration_bmsb <- tibble::as_tibble(read_xlsx(paste0(penetration_packgroup_folder,"/",file_names[10]),sheet = 1,skip = 14))
penetration_bmsp_oi <- tibble::as_tibble(read_xlsx(paste0(penetration_packgroup_folder,"/",file_names[11]),sheet = 1,skip = 14))
penetration_bmsp <- tibble::as_tibble(read_xlsx(paste0(penetration_packgroup_folder,"/",file_names[12]),sheet = 1,skip = 14))
#clean the dataframes using clean function
penetration_bmbp_oi <- input_clean(penetration_bmbp_oi)
penetration_bmbp <- input_clean(penetration_bmbp);penetration_bmbp<- penetration_bmbp[complete.cases(penetration_bmbp),]
penetration_bmbs_oi <- input_clean(penetration_bmbs_oi)
penetration_bmbs <- input_clean(penetration_bmbs);penetration_bmbs <- penetration_bmbs[complete.cases(penetration_bmbs),]
penetration_bmmb_oi <- input_clean(penetration_bmmb_oi)
penetration_bmmb <- input_clean(penetration_bmmb);penetration_bmmb <- penetration_bmmb[complete.cases(penetration_bmmb),]
penetration_bmmp_oi <- input_clean(penetration_bmmp_oi)
penetration_bmmp <- input_clean(penetration_bmmp);penetration_bmmp <- penetration_bmmp[complete.cases(penetration_bmmp),]
penetration_bmsb_oi <- input_clean(penetration_bmsb_oi)
penetration_bmsb <- input_clean(penetration_bmsb);penetration_bmsb <- penetration_bmsb[complete.cases(penetration_bmsb),]
penetration_bmsp_oi <- input_clean(penetration_bmsp_oi)
penetration_bmsp <- input_clean(penetration_bmsp);penetration_bmsp <- penetration_bmsp[complete.cases(penetration_bmsp),]
#perform the data cleaning for the df "other islands(oi) for the 6 packgroups
penetration_bmbp_oi <- penetration_oi_format(penetration_bmbp_oi); penetration_bmbs_oi <- penetration_oi_format(penetration_bmbs_oi)
penetration_bmmb_oi <- penetration_oi_format(penetration_bmmb_oi);penetration_bmmp_oi <- penetration_oi_format(penetration_bmmp_oi)
penetration_bmsb_oi <- penetration_oi_format(penetration_bmsb_oi);penetration_bmsp_oi <- penetration_oi_format(penetration_bmsp_oi)
penetration.totalcat.oi.df
#merge the dataframe based on the packgroup data frame for the "OI"
penetration_bmbp_oi <- merge.data.frame(penetration_bmbp_oi,penetration.totalcat.oi.df,by = c("Account/Region","Month"))
penetration_bmbs_oi <- merge.data.frame(penetration_bmbs_oi,penetration.totalcat.oi.df,by = c("Account/Region","Month"))
penetration_bmmb_oi <- merge.data.frame(penetration_bmmb_oi,penetration.totalcat.oi.df,by = c("Account/Region","Month"))
penetration_bmmp_oi <- merge.data.frame(penetration_bmmp_oi,penetration.totalcat.oi.df,by = c("Account/Region","Month"))
penetration_bmsb_oi <- merge.data.frame(penetration_bmsb_oi,penetration.totalcat.oi.df,by = c("Account/Region","Month"))
penetration_bmsp_oi <- merge.data.frame(penetration_bmsp_oi,penetration.totalcat.oi.df,by = c("Account/Region","Month"))
#perform the data preprocessing for the other 3 regions - at pack group level
penetration.totalcat.region.mapped
penetration_bmbp <-penetration_3regions_format(penetration_bmbp);penetration_bmbs <- penetration_3regions_format(penetration_bmbs)
penetration_bmmb <- penetration_3regions_format(penetration_bmmb);penetration_bmmp <- penetration_3regions_format(penetration_bmmp)
penetration_bmsb <- penetration_3regions_format(penetration_bmsb);penetration_bmsp <- penetration_3regions_format(penetration_bmsp)
#merge the dataframe based on the packgroup data frame for the 3 regions ( CEJ,WJ,Sumatra)
penetration_bmbp <- merge.data.frame(penetration_bmbp,penetration.totalcat.region.mapped,by=c("Account/Region","Month"))
penetration_bmbs <- merge.data.frame(penetration_bmbs,penetration.totalcat.region.mapped,by=c("Account/Region","Month"))
penetration_bmmp <- merge.data.frame(penetration_bmmp,penetration.totalcat.region.mapped,by=c("Account/Region","Month"))
penetration_bmmb <- merge.data.frame(penetration_bmmb,penetration.totalcat.region.mapped,by=c("Account/Region","Month"))
penetration_bmsb <- merge.data.frame(penetration_bmsb,penetration.totalcat.region.mapped,by=c("Account/Region","Month"))
penetration_bmsp <- merge.data.frame(penetration_bmsp,penetration.totalcat.region.mapped,by=c("Account/Region","Month"))
#perform the rbind for 3+1 region df across 6 packgroup's
penetration_bmbp_final <- rbind(penetration_bmbp,penetration_bmbp_oi)
penetration_bmbs_final <- rbind(penetration_bmbs,penetration_bmbs_oi)
penetration_bmmp_final <- rbind(penetration_bmmp,penetration_bmmp_oi)
penetration_bmmb_final <- rbind(penetration_bmmb,penetration_bmmb_oi)
penetration_bmsb_final <- rbind(penetration_bmsb,penetration_bmsb_oi)
penetration_bmsp_final <- rbind(penetration_bmsp,penetration_bmsp_oi)
#calculate the penetration% for the each packgroup data frame
penetration_bmbp_final <- penetration_percentage(penetration_bmbp_final);penetration_bmbs_final <- penetration_percentage(penetration_bmbs_final)
penetration_bmmp_final <- penetration_percentage(penetration_bmmp_final);penetration_bmmb_final <- penetration_percentage(penetration_bmmb_final)
penetration_bmsb_final <- penetration_percentage(penetration_bmsb_final);penetration_bmsp_final <- penetration_percentage(penetration_bmsp_final)
#perform the month mapping for the final packgroup data in <month>-<year>
penetration_bmbp_final <- penetration_packgroup_output(penetration_bmbp_final)
penetration_bmbs_final <- penetration_packgroup_output(penetration_bmbs_final)
penetration_bmmp_final <- penetration_packgroup_output(penetration_bmmp_final)
penetration_bmmb_final <- penetration_packgroup_output(penetration_bmmb_final)
penetration_bmsb_final <- penetration_packgroup_output(penetration_bmsb_final)
penetration_bmsp_final <- penetration_packgroup_output(penetration_bmsp_final)

############################################################### packgroup ( consolidation - first by packgroup then by region ) #################################
#collate the source dataframes into a single place for the consolidation (except the BMI )

psales.packgroup.aggregate ; ssales.packgroup.aggregate
sstock.packgroup.aggregate;  droo.packgroup.aggregate
tts.packgroup.aggregate.final;itrust.packgroup.aggregate

packgroup.consolidation <- merge.data.frame(psales.packgroup.aggregate,ssales.packgroup.aggregate,by = c("Month","Product","Customer"))
packgroup.consolidation1 <- merge.data.frame(packgroup.consolidation,sstock.packgroup.aggregate,by=c("Month","Product","Customer"),all = T)
packgroup.consolidation2 <- merge.data.frame(packgroup.consolidation1,droo.packgroup.aggregate,by=c("Month","Product","Customer"),all = T)
packgroup.consolidation3 <- merge.data.frame(packgroup.consolidation2,itrust.packgroup.aggregate,by=c("Month","Product","Customer"),all.x =T)
packgroup.consolidation4 <- merge.data.frame(packgroup.consolidation3,tts.packgroup.aggregate.final,by=c("Month","Product","Customer"),all = T)
#segregate the consolidated packgroup into 6 packgroup files 
packgroup.list <- unique(packgroup.consolidation4$Product)
packgroup.bmbp.df <- packgroup.consolidation4%>%
  filter(Product=="Bango Manis Big Pouch" )
packgroup.bmbs.df <- packgroup.consolidation4%>%
  filter(Product=="Bango Manis Big Sachet" )
packgroup.bmmb.df <- packgroup.consolidation4%>%
  filter(Product=="Bango Manis Medium Bottle" )
packgroup.bmmp.df <- packgroup.consolidation4%>%
  filter(Product=="Bango Manis Medium Pouch" )
packgroup.bmsb.df <- packgroup.consolidation4%>%
  filter(Product=="Bango Manis Small Bottle" )
packgroup.bmsp.df <- packgroup.consolidation4%>%
  filter(Product=="Bango Manis Small Pouch" )

#peform the merging for the BMI data with the packgroup data 
packgroup.bmbp.df1 <- merge.data.frame(packgroup.bmbp.df,bmi.packgroup.bmbp,by.x = c("Month","Product","Customer"),by.y = c("Month","Required_Packgroup","Account/Region"),all = T)
packgroup.bmbs.df1 <- merge.data.frame(packgroup.bmbs.df,bmi.packgroup.bmbs,by.x = c("Month","Product","Customer"),by.y = c("Month","Required_Packgroup","Account/Region"),all = T)
packgroup.bmmb.df1 <- merge.data.frame(packgroup.bmmb.df,bmi.packgroup.bmmb,by.x = c("Month","Product","Customer"),by.y = c("Month","Required_Packgroup","Account/Region"),all = T)
packgroup.bmmp.df1 <- merge.data.frame(packgroup.bmmp.df,bmi.packgroup.bmmp,by.x = c("Month","Product","Customer"),by.y = c("Month","Required_Packgroup","Account/Region"),all = T)
packgroup.bmsp.df1 <- merge.data.frame(packgroup.bmsp.df,bmi.packgroup.bmsp,by.x = c("Month","Product","Customer"),by.y = c("Month","Required_Packgroup","Account/Region"),all = T)
packgroup.bmsb.df1 <- merge.data.frame(packgroup.bmsb.df,bmi.packgroup.bmsb,by.x = c("Month","Product","Customer"),by.y = c("Month","Required_Packgroup","Account/Region"),all = T)

#map the penetration data into the respective packgroups 
packgroup.bmbp.df2 <- merge.data.frame(packgroup.bmbp.df1,penetration_bmbp_final,by.x = c("Month","Customer"),by.y=c("Month","Account/Region"),all.x =T)
packgroup.bmbs.df2 <- merge.data.frame(packgroup.bmbs.df1,penetration_bmbs_final,by.x = c("Month","Customer"),by.y=c("Month","Account/Region"),all.x =T)
packgroup.bmmp.df2 <- merge.data.frame(packgroup.bmmp.df1,penetration_bmmp_final,by.x = c("Month","Customer"),by.y=c("Month","Account/Region"),all.x =T)
packgroup.bmmb.df2 <- merge.data.frame(packgroup.bmmb.df1,penetration_bmmb_final,by.x = c("Month","Customer"),by.y=c("Month","Account/Region"),all.x =T)
packgroup.bmsb.df2 <- merge.data.frame(packgroup.bmsb.df1,penetration_bmsb_final,by.x = c("Month","Customer"),by.y=c("Month","Account/Region"),all.x =T)
packgroup.bmsp.df2 <- merge.data.frame(packgroup.bmsp.df1,penetration_bmsp_final,by.x = c("Month","Customer"),by.y=c("Month","Account/Region"),all.x =T)

# add the sales organisation , Brand and Channel to each 6 packgroup dataset
packgroup.bmbp.df2%<>%mutate(`Sales Organization`="9001",Category="SAVOURY",Sector="SEASONINGS",Brand="BANGO",Channel="DISTRIBUTIVE TRADE")
packgroup.bmbp.df2%<>%select(c(Month,`Sales Organization`:Brand,Product,Channel,Customer,GSV:penetration_percentage))
#perform the data transformation for the Bango manis Big Sachet
packgroup.bmbs.df2%<>%mutate(`Sales Organization`="9001",Category="SAVOURY",Sector="SEASONINGS",Brand="BANGO",Channel="DISTRIBUTIVE TRADE")
packgroup.bmbs.df2%<>%select(c(Month,`Sales Organization`:Brand,Product,Channel,Customer,GSV:penetration_percentage))
#perform the data transformation for the Bango manis Big Sachet
packgroup.bmmp.df2%<>%mutate(`Sales Organization`="9001",Category="SAVOURY",Sector="SEASONINGS",Brand="BANGO",Channel="DISTRIBUTIVE TRADE")
packgroup.bmmp.df2%<>%select(c(Month,`Sales Organization`:Brand,Product,Channel,Customer,GSV:penetration_percentage))
#perform the data transformation for the Bango manis Big Sachet
packgroup.bmmb.df2%<>%mutate(`Sales Organization`="9001",Category="SAVOURY",Sector="SEASONINGS",Brand="BANGO",Channel="DISTRIBUTIVE TRADE")
packgroup.bmmb.df2%<>%select(c(Month,`Sales Organization`:Brand,Product,Channel,Customer,GSV:penetration_percentage))
#perform the data transformation for the Bango manis Big Sachet
packgroup.bmsb.df2%<>%mutate(`Sales Organization`="9001",Category="SAVOURY",Sector="SEASONINGS",Brand="BANGO",Channel="DISTRIBUTIVE TRADE")
packgroup.bmsb.df2%<>%select(c(Month,`Sales Organization`:Brand,Product,Channel,Customer,GSV:penetration_percentage))
#perform the data transformation for the Bango manis Big Sachet
packgroup.bmsp.df2%<>%mutate(`Sales Organization`="9001",Category="SAVOURY",Sector="SEASONINGS",Brand="BANGO",Channel="DISTRIBUTIVE TRADE")
packgroup.bmsp.df2%<>%select(c(Month,`Sales Organization`:Brand,Product,Channel,Customer,GSV:penetration_percentage))

#rename the column headers as per the output for each packgroup {Bango Manis Big Pouch : bmbp}
packgroup.bmbp.headers <- c("Month",
                            "Sales Organization",
                            "Category",
                            "Sector",
                            "Brand",
                            "Product",
                            "Channel",
                            "Customer",
                            "Bango Manis Big Pouch_GSV",
                            "Bango Manis Big Pouch_Primary Sales Qty(PC)",
                            "Bango Manis Big Pouch_NIV",
                            "Bango Manis Big Pouch_Primary Sales Qty(Kg)",
                            "Bango Manis Big Pouch_Primary Sales Qty(Tonns)",
                            "Bango Manis Big Pouch_Sec Sales Value",
                            "Bango Manis Big Pouch_Sec Sales Volume(PC)",
                            "Bango Manis Big Pouch_Sec sales Volume (Kg)",
                            "Bango Manis Big Pouch_Sec sales Volume (Tonns)",
                            "Bango Manis Big Pouch_Secondary Stock Value [@ DT Rate.]",
                            "Bango Manis Big Pouch_Secondary Stock Volume(PC)",
                            "Bango Manis Big Pouch_Secondary stock volume(Kg)",
                            "Bango Manis Big Pouch_Secondary stock volume(Tonns)",
                            "Bango Manis Big Pouch_Original Order Qty.(PC)",
                            "Bango Manis Big Pouch_Final Customer Expected Order Qty.(PC)",
                            "Bango Manis Big Pouch_Dispatched Qty.(PC)",
                            "Bango Manis Big Pouch_DROO",
                            "Bango Manis Big Pouch_DR",
                            "Bango Manis Big Pouch_I trust Line",
                            "Bango Manis Big Pouch_I trust Total",
                            "Bango Manis Big Pouch_I trust %",
                            "Bango Manis Big Pouch_TTS",
                            "Bango Manis Big Pouch_BBT",
                            "Bango Manis Big Pouch_BBT-Place",
                            "Bango Manis Big Pouch_BBT-Place on invoice",
                            "Bango Manis Big Pouch_BBT-Place off invoice",
                            "Bango Manis Big Pouch_BBT-Price",
                            "Bango Manis Big Pouch_CPP on invoice",
                            "Bango Manis Big Pouch_CPP off invoice",
                            "Bango Manis Big Pouch_BBT-Product",
                            "Bango Manis Big Pouch_BBT-Pack",
                            "Bango Manis Big Pouch_BBT-Proposition",
                            "Bango Manis Big Pouch_BBT-Promotion",
                            "Bango Manis Big Pouch_EOT",
                            "Bango Manis Big Pouch_Brand & Marketing Investment",
                            "Bango Manis Big Pouch_Brand & Marketing Investment Trade",
                            "Bango Manis Big Pouch_Brand & Marketing Investment Consumer",
                            "Bango Manis Big Pouch_Promotion Team Cost Trade",
                            "Bango Manis Big Pouch_Promotion Team Cost Consumer",
                            "Bango Manis Big Pouch_Promotion Team Cost",
                            "Bango Manis Big Pouch_Promotion Packaging Material Cost Trade",
                            "Bango Manis Big Pouch_Promotion Packaging Material Cost Consumer",
                            "Bango Manis Big Pouch_Promotion Communication Material Cost Trade",
                            "Bango Manis Big Pouch_Promotion Communication Material Cost Consumer",
                            "Bango Manis Big Pouch_Promotion Communication Material Cost",
                            "Bango Manis Big Pouch_Promo Agency Remuneration Fees & Commissions Consumer",
                            "Bango Manis Big Pouch_Promotional Expenses",
                            "Bango Manis Big Pouch_Promotion Packaging Material Cost",
                            "Bango Manis Big Pouch_Promotion Repacking Cost",
                            "Bango Manis Big Pouch_Promotion Repacking Cost Trade",
                            "Bango Manis Big Pouch_Promotion Repacking Cost Consumer",
                            "Bango Manis Big Pouch_Promo Samples, Gifts and Incentive Costs",
                            "Bango Manis Big Pouch_Promo Samples, Gifts and Incentive Costs Consumer",
                            "Bango Manis Big Pouch_Promo Agency Remun Fees & Commissions",
                            "Bango Manis Big Pouch_Promo Agency Remuneration Fees & Commissions Trade",
                            "Bango Manis Big Pouch_Internal_Penetration",
                            "Bango Manis Big Pouch_Category_Penetration",
                            "Bango Manis Big Pouch_Internal_penetration%"
                            
)

colnames(packgroup.bmbp.df2) <- packgroup.bmbp.headers
#packgroup.bmbp.df2 <- data.frame(lapply(packgroup.bmbp.df2, gsub, pattern='NA', replacement=''))
#seggregate the packgroup bmbp into 4 regions {cej;oi;wj;sumatera}
packgroup.bmbp.cej <- packgroup.bmbp.df2%>%
  filter(str_detect(Customer,"DT CENTRAL & EAST JAVA")|Customer=="ALL")%>%
  arrange(Customer)
#packgroup.bmbp.cej <- format(packgroup.bmbp.cej,scientific=F)
#packgroup.bmbp.cej <- data.frame(lapply(packgroup.bmbp.cej, gsub, pattern='NA', replacement=''))
write_xlsx(packgroup.bmbp.cej,paste0(output.path,"/packgroup.bmbp.cej.xlsx"))

packgroup.bmbp.oi <- packgroup.bmbp.df2%>%
  filter(str_detect(Customer,"DT OI")|Customer=="ALL")%>%
  arrange(Customer)
#packgroup.bmbp.oi <- format(packgroup.bmbp.oi,scientific=F)
#packgroup.bmbp.oi <- data.frame(lapply(packgroup.bmbp.oi, gsub, pattern='NA', replacement=''))
write_xlsx(packgroup.bmbp.oi,paste0(output.path,"/packgroup.bmbp.oi.xlsx"))

packgroup.bmbp.sumatera <- packgroup.bmbp.df2%>%
  filter(str_detect(Customer,"DT SUMATERA")|Customer=="ALL")%>%
  arrange(Customer)
#packgroup.bmbp.sumatera <- format(packgroup.bmbp.sumatera,scientific=F)
#packgroup.bmbp.sumatera <- data.frame(lapply(packgroup.bmbp.sumatera, gsub, pattern='NA', replacement=''))
write_xlsx(packgroup.bmbp.sumatera,paste0(output.path,"/packgroup.bmbp.sumatera.xlsx"))

packgroup.bmbp.wj <- packgroup.bmbp.df2%>%
  filter(str_detect(Customer,"DT WEST JAVA")|Customer=="ALL")%>%
  arrange(Customer)
#packgroup.bmbp.wj <- format(packgroup.bmbp.wj,scientific=F)
#packgroup.bmbp.wj <- data.frame(lapply(packgroup.bmbp.wj, gsub, pattern='NA', replacement=''))
write_xlsx(packgroup.bmbp.wj,paste0(output.path,"/packgroup.bmbp.wj.xlsx"))


#rename the column headers as per the output for each packgroup {Bango Manis Big Sachet : bmbs}
packgroup.bmbs.headers <- c("Month",
                            "Sales Organization",
                            "Category",
                            "Sector",
                            "Brand",
                            "Product",
                            "Channel",
                            "Customer",
                            "Bango Manis Big Sachet_GSV",
                            "Bango Manis Big Sachet_Primary Sales Qty(PC)",
                            "Bango Manis Big Sachet_NIV",
                            "Bango Manis Big Sachet_Primary Sales Qty(Kg)",
                            "Bango Manis Big Sachet_Primary Sales Qty(Tonns)",
                            "Bango Manis Big Sachet_Sec Sales Value",
                            "Bango Manis Big Sachet_Sec Sales Volume(PC)",
                            "Bango Manis Big Sachet_Sec sales Volume (Kg)",
                            "Bango Manis Big Sachet_Sec sales Volume (Tonns)",
                            "Bango Manis Big Sachet_Secondary Stock Value [@ DT Rate.]",
                            "Bango Manis Big Sachet_Secondary Stock Volume(PC)",
                            "Bango Manis Big Sachet_Secondary stock volume(Kg)",
                            "Bango Manis Big Sachet_Secondary stock volume(Tonns)",
                            "Bango Manis Big Sachet_Original Order Qty.(PC)",
                            "Bango Manis Big Sachet_Final Customer Expected Order Qty.(PC)",
                            "Bango Manis Big Sachet_Dispatched Qty.(PC)",
                            "Bango Manis Big Sachet_DROO",
                            "Bango Manis Big Sachet_DR",
                            "Bango Manis Big Sachet_I trust Line",
                            "Bango Manis Big Sachet_I trust Total",
                            "Bango Manis Big Sachet_I trust %",
                            "Bango Manis Big Sachet_TTS",
                            "Bango Manis Big Sachet_BBT",
                            "Bango Manis Big Sachet_BBT-Place",
                            "Bango Manis Big Sachet_BBT-Place on invoice",
                            "Bango Manis Big Sachet_BBT-Place off invoice",
                            "Bango Manis Big Sachet_BBT-Price",
                            "Bango Manis Big Sachet_CPP on invoice",
                            "Bango Manis Big Sachet_CPP off invoice",
                            "Bango Manis Big Sachet_BBT-Product",
                            "Bango Manis Big Sachet_BBT-Pack",
                            "Bango Manis Big Sachet_BBT-Proposition",
                            "Bango Manis Big Sachet_BBT-Promotion",
                            "Bango Manis Big Sachet_EOT",
                            "Bango Manis Big Sachet_Brand & Marketing Investment",
                            "Bango Manis Big Sachet_Brand & Marketing Investment Trade",
                            "Bango Manis Big Sachet_Brand & Marketing Investment Consumer",
                            "Bango Manis Big Sachet_Promotion Team Cost Trade",
                            "Bango Manis Big Sachet_Promotion Team Cost Consumer",
                            "Bango Manis Big Sachet_Promotion Team Cost",
                            "Bango Manis Big Sachet_Promotion Packaging Material Cost Trade",
                            "Bango Manis Big Sachet_Promotion Packaging Material Cost Consumer",
                            "Bango Manis Big Sachet_Promotion Communication Material Cost Trade",
                            "Bango Manis Big Sachet_Promotion Communication Material Cost Consumer",
                            "Bango Manis Big Sachet_Promotion Communication Material Cost",
                            "Bango Manis Big Sachet_Promo Agency Remuneration Fees & Commissions Consumer",
                            "Bango Manis Big Sachet_Promotional Expenses",
                            "Bango Manis Big Sachet_Promotion Packaging Material Cost",
                            "Bango Manis Big Sachet_Promotion Repacking Cost",
                            "Bango Manis Big Sachet_Promotion Repacking Cost Trade",
                            "Bango Manis Big Sachet_Promotion Repacking Cost Consumer",
                            "Bango Manis Big Sachet_Promo Samples, Gifts and Incentive Costs",
                            "Bango Manis Big Sachet_Promo Samples, Gifts and Incentive Costs Consumer",
                            "Bango Manis Big Sachet_Promo Agency Remun Fees & Commissions",
                            "Bango Manis Big Sachet_Promo Agency Remuneration Fees & Commissions Trade",
                            "Bango Manis Big Sachet_Internal_Penetration",
                            "Bango Manis Big Sachet_Category_Penetration",
                            "Bango Manis Big Sachet_Internal_penetration%"
)                            

colnames(packgroup.bmbs.df2) <- packgroup.bmbs.headers

#seggregate the packgroup bmbp into 4 regions {cej;oi;wj;sumatera}
packgroup.bmbs.cej <- packgroup.bmbs.df2%>%
  filter(str_detect(Customer,"DT CENTRAL & EAST JAVA")|Customer=="ALL")%>%
  arrange(Customer)
#packgroup.bmbs.cej <- format(packgroup.bmbs.cej,scientific=F)
#packgroup.bmbs.cej <- data.frame(lapply(packgroup.bmbs.cej, gsub, pattern='NA', replacement=''))
write_xlsx(packgroup.bmbs.cej,paste0(output.path,"/packgroup.bmbs.cej.xlsx"))


packgroup.bmbs.oi <- packgroup.bmbs.df2%>%
  filter(str_detect(Customer,"DT OI")|Customer=="ALL")%>%
  arrange(Customer)
#packgroup.bmbp.oi <- format(packgroup.bmbs.oi,scientific=F)
#packgroup.bmbs.oi <- data.frame(lapply(packgroup.bmbs.oi, gsub, pattern='NA', replacement=''))
write_xlsx(packgroup.bmbs.oi,paste0(output.path,"/packgroup.bmbs.oi.xlsx"))


packgroup.bmbs.sumatera <- packgroup.bmbs.df2%>%
  filter(str_detect(Customer,"DT SUMATERA")|Customer=="ALL")%>%
  arrange(Customer)
#packgroup.bmbs.sumatera <- format(packgroup.bmbs.sumatera,scientific=F)
#packgroup.bmbs.sumatera <- data.frame(lapply(packgroup.bmbs.sumatera, gsub, pattern='NA', replacement=''))
write_xlsx(packgroup.bmbs.sumatera,paste0(output.path,"/packgroup.bmbs.sumatera.xlsx"))


packgroup.bmbs.wj <- packgroup.bmbs.df2%>%
  filter(str_detect(Customer,"DT WEST JAVA")|Customer=="ALL")%>%
  arrange(Customer)
#packgroup.bmbs.wj <- format(packgroup.bmbs.wj,scientific=F)
#packgroup.bmbs.wj <- data.frame(lapply(packgroup.bmbs.wj, gsub, pattern='NA', replacement=''))
write_xlsx(packgroup.bmbs.wj,paste0(output.path,"/packgroup.bmbs.wj.xlsx"))

#rename the column headers as per the output for each packgroup {Bango Manis Medium Pouch : bmmp}
packgroup.bmmp.headers <- c("Month",
                            "Sales Organization",
                            "Category",
                            "Sector",
                            "Brand",
                            "Product",
                            "Channel",
                            "Customer",
                            "Bango Manis Medium Pouch_GSV",
                            "Bango Manis Medium Pouch_Primary Sales Qty(PC)",
                            "Bango Manis Medium Pouch_NIV",
                            "Bango Manis Medium Pouch_Primary Sales Qty(Kg)",
                            "Bango Manis Medium Pouch_Primary Sales Qty(Tonns)",
                            "Bango Manis Medium Pouch_Sec Sales Value",
                            "Bango Manis Medium Pouch_Sec Sales Volume(PC)",
                            "Bango Manis Medium Pouch_Sec sales Volume (Kg)",
                            "Bango Manis Medium Pouch_Sec sales Volume (Tonns)",
                            "Bango Manis Medium Pouch_Secondary Stock Value [@ DT Rate.]",
                            "Bango Manis Medium Pouch_Secondary Stock Volume(PC)",
                            "Bango Manis Medium Pouch_Secondary stock volume(Kg)",
                            "Bango Manis Medium Pouch_Secondary stock volume(Tonns)",
                            "Bango Manis Medium Pouch_Original Order Qty.(PC)",
                            "Bango Manis Medium Pouch_Final Customer Expected Order Qty.(PC)",
                            "Bango Manis Medium Pouch_Dispatched Qty.(PC)",
                            "Bango Manis Medium Pouch_DROO",
                            "Bango Manis Medium Pouch_DR",
                            "Bango Manis Medium Pouch_I trust Line",
                            "Bango Manis Medium Pouch_I trust Total",
                            "Bango Manis Medium Pouch_I trust %",
                            "Bango Manis Medium Pouch_TTS",
                            "Bango Manis Medium Pouch_BBT",
                            "Bango Manis Medium Pouch_BBT-Place",
                            "Bango Manis Medium Pouch_BBT-Place on invoice",
                            "Bango Manis Medium Pouch_BBT-Place off invoice",
                            "Bango Manis Medium Pouch_BBT-Price",
                            "Bango Manis Medium Pouch_CPP on invoice",
                            "Bango Manis Medium Pouch_CPP off invoice",
                            "Bango Manis Medium Pouch_BBT-Product",
                            "Bango Manis Medium Pouch_BBT-Pack",
                            "Bango Manis Medium Pouch_BBT-Proposition",
                            "Bango Manis Medium Pouch_BBT-Promotion",
                            "Bango Manis Medium Pouch_EOT",
                            "Bango Manis Medium Pouch_Brand & Marketing Investment",
                            "Bango Manis Medium Pouch_Brand & Marketing Investment Trade",
                            "Bango Manis Medium Pouch_Brand & Marketing Investment Consumer",
                            "Bango Manis Medium Pouch_Promotion Team Cost Trade",
                            "Bango Manis Medium Pouch_Promotion Team Cost Consumer",
                            "Bango Manis Medium Pouch_Promotion Team Cost",
                            "Bango Manis Medium Pouch_Promotion Packaging Material Cost Trade",
                            "Bango Manis Medium Pouch_Promotion Packaging Material Cost Consumer",
                            "Bango Manis Medium Pouch_Promotion Communication Material Cost Trade",
                            "Bango Manis Medium Pouch_Promotion Communication Material Cost Consumer",
                            "Bango Manis Medium Pouch_Promotion Communication Material Cost",
                            "Bango Manis Medium Pouch_Promo Agency Remuneration Fees & Commissions Consumer",
                            "Bango Manis Medium Pouch_Promotional Expenses",
                            "Bango Manis Medium Pouch_Promotion Packaging Material Cost",
                            "Bango Manis Medium Pouch_Promotion Repacking Cost",
                            "Bango Manis Medium Pouch_Promotion Repacking Cost Trade",
                            "Bango Manis Medium Pouch_Promotion Repacking Cost Consumer",
                            "Bango Manis Medium Pouch_Promo Samples, Gifts and Incentive Costs",
                            "Bango Manis Medium Pouch_Promo Samples, Gifts and Incentive Costs Consumer",
                            "Bango Manis Medium Pouch_Promo Agency Remun Fees & Commissions",
                            "Bango Manis Medium Pouch_Promo Agency Remuneration Fees & Commissions Trade",
                            "Bango Manis Medium Pouch_Internal_Penetration",
                            "Bango Manis Medium Pouch_Category_Penetration",
                            "Bango Manis Medium Pouch_Internal_penetration%"
                            
)

colnames(packgroup.bmmp.df2) <- packgroup.bmmp.headers
#packgroup.bmmp.df2 <- data.frame(lapply(packgroup.bmmp.df2, gsub, pattern='NA', replacement=''))

#seggregate the packgroup bmbp into 4 regions

#seggregate the packgroup bmbp into 4 regions {cej;oi;wj;sumatera}
packgroup.bmmp.cej <- packgroup.bmmp.df2%>%
  filter(str_detect(Customer,"DT CENTRAL & EAST JAVA")|Customer=="ALL")%>%
  arrange(Customer)
#packgroup.bmmp.cej <- format(packgroup.bmmp.cej,scientific=F)
#packgroup.bmmp.cej <- data.frame(lapply(packgroup.bmmp.cej, gsub, pattern='NA', replacement=''))
write_xlsx(packgroup.bmmp.cej,paste0(output.path,"/packgroup.bmmp.cej.xlsx"))


packgroup.bmmp.oi <- packgroup.bmmp.df2%>%
  filter(str_detect(Customer,"DT OI")|Customer=="ALL")%>%
  arrange(Customer)
#packgroup.bmmp.oi <- format(packgroup.bmmp.oi,scientific=F)
#packgroup.bmmp.oi <- data.frame(lapply(packgroup.bmmp.oi, gsub, pattern='NA', replacement=''))
write_xlsx(packgroup.bmmp.oi,paste0(output.path,"/packgroup.bmmp.oi.xlsx"))


packgroup.bmmp.sumatera <- packgroup.bmmp.df2%>%
  filter(str_detect(Customer,"DT SUMATERA")|Customer=="ALL")%>%
  arrange(Customer)
#packgroup.bmmp.sumatera <- format(packgroup.bmmp.sumatera,scientific=F) 
#packgroup.bmmp.sumatera <- data.frame(lapply(packgroup.bmmp.sumatera, gsub, pattern='NA', replacement=''))
write_xlsx(packgroup.bmmp.sumatera,paste0(output.path,"/packgroup.bmmp.sumatera.xlsx"))


packgroup.bmmp.wj <- packgroup.bmmp.df2%>%
  filter(str_detect(Customer,"DT WEST JAVA")|Customer=="ALL")%>%
  arrange(Customer)
#packgroup.bmmp.wj <- format(packgroup.bmmp.wj,scientific=F) 
#packgroup.bmmp.wj <- data.frame(lapply(packgroup.bmmp.wj, gsub, pattern='NA', replacement=''))
write_xlsx(packgroup.bmmp.wj,paste0(output.path,"/packgroup.bmmp.wj.xlsx"))


#rename the column headers as per the output for each packgroup {Bango Manis Medium Bottle : bmmb}
packgroup.bmmb.headers <- c("Month",
                            "Sales Organization",
                            "Category",
                            "Sector",
                            "Brand",
                            "Product",
                            "Channel",
                            "Customer",
                            "Bango Manis Medium Bottle_GSV",
                            "Bango Manis Medium Bottle_Primary Sales Qty(PC)",
                            "Bango Manis Medium Bottle_NIV",
                            "Bango Manis Medium Bottle_Primary Sales Qty(Kg)",
                            "Bango Manis Medium Bottle_Primary Sales Qty(Tonns)",
                            "Bango Manis Medium Bottle_Sec Sales Value",
                            "Bango Manis Medium Bottle_Sec Sales Volume(PC)",
                            "Bango Manis Medium Bottle_Sec sales Volume (Kg)",
                            "Bango Manis Medium Bottle_Sec sales Volume (Tonns)",
                            "Bango Manis Medium Bottle_Secondary Stock Value [@ DT Rate.]",
                            "Bango Manis Medium Bottle_Secondary Stock Volume(PC)",
                            "Bango Manis Medium Bottle_Secondary stock volume(Kg)",
                            "Bango Manis Medium Bottle_Secondary stock volume(Tonns)",
                            "Bango Manis Medium Bottle_Original Order Qty.(PC)",
                            "Bango Manis Medium Bottle_Final Customer Expected Order Qty.(PC)",
                            "Bango Manis Medium Bottle_Dispatched Qty.(PC)",
                            "Bango Manis Medium Bottle_DROO",
                            "Bango Manis Medium Bottle_DR",
                            "Bango Manis Medium Bottle_I trust Line",
                            "Bango Manis Medium Bottle_I trust Total",
                            "Bango Manis Medium Bottle_I trust %",
                            "Bango Manis Medium Bottle_TTS",
                            "Bango Manis Medium Bottle_BBT",
                            "Bango Manis Medium Bottle_BBT-Place",
                            "Bango Manis Medium Bottle_BBT-Place on invoice",
                            "Bango Manis Medium Bottle_BBT-Place off invoice",
                            "Bango Manis Medium Bottle_BBT-Price",
                            "Bango Manis Medium Bottle_CPP on invoice",
                            "Bango Manis Medium Bottle_CPP off invoice",
                            "Bango Manis Medium Bottle_BBT-Product",
                            "Bango Manis Medium Bottle_BBT-Pack",
                            "Bango Manis Medium Bottle_BBT-Proposition",
                            "Bango Manis Medium Bottle_BBT-Promotion",
                            "Bango Manis Medium Bottle_EOT",
                            "Bango Manis Medium Bottle_Brand & Marketing Investment",
                            "Bango Manis Medium Bottle_Brand & Marketing Investment Trade",
                            "Bango Manis Medium Bottle_Brand & Marketing Investment Consumer",
                            "Bango Manis Medium Bottle_Promotion Team Cost Trade",
                            "Bango Manis Medium Bottle_Promotion Team Cost Consumer",
                            "Bango Manis Medium Bottle_Promotion Team Cost",
                            "Bango Manis Medium Bottle_Promotion Packaging Material Cost Trade",
                            "Bango Manis Medium Bottle_Promotion Packaging Material Cost Consumer",
                            "Bango Manis Medium Bottle_Promotion Communication Material Cost Trade",
                            "Bango Manis Medium Bottle_Promotion Communication Material Cost Consumer",
                            "Bango Manis Medium Bottle_Promotion Communication Material Cost",
                            "Bango Manis Medium Bottle_Promo Agency Remuneration Fees & Commissions Consumer",
                            "Bango Manis Medium Bottle_Promotional Expenses",
                            "Bango Manis Medium Bottle_Promotion Packaging Material Cost",
                            "Bango Manis Medium Bottle_Promotion Repacking Cost",
                            "Bango Manis Medium Bottle_Promotion Repacking Cost Trade",
                            "Bango Manis Medium Bottle_Promotion Repacking Cost Consumer",
                            "Bango Manis Medium Bottle_Promo Samples, Gifts and Incentive Costs",
                            "Bango Manis Medium Bottle_Promo Samples, Gifts and Incentive Costs Consumer",
                            "Bango Manis Medium Bottle_Promo Agency Remun Fees & Commissions",
                            "Bango Manis Medium Bottle_Promo Agency Remuneration Fees & Commissions Trade",
                            "Bango Manis Medium Bottle_Internal_Penetration",
                            "Bango Manis Medium Bottle_Category_Penetration",
                            "Bango Manis Medium Bottle_Internal_penetration%"
                            
)

colnames(packgroup.bmmb.df2) <- packgroup.bmmb.headers
#packgroup.bmmb.df2 <- data.frame(lapply(packgroup.bmmb.df2, gsub, pattern='NA', replacement=''))

#seggregate the packgroup data into regions

packgroup.bmmb.cej <- packgroup.bmmb.df2%>%
  filter(str_detect(Customer,"DT CENTRAL & EAST JAVA")|Customer=="ALL")%>%
  arrange(Customer)
#packgroup.bmmb.cej <- format(packgroup.bmmb.cej,scientific=F)
#packgroup.bmmb.cej <- data.frame(lapply(packgroup.bmmb.cej, gsub, pattern='NA', replacement=''))
write_xlsx(packgroup.bmmb.cej,paste0(output.path,"/packgroup.bmmb.cej.xlsx"))


packgroup.bmmb.oi <- packgroup.bmmb.df2%>%
  filter(str_detect(Customer,"DT OI")|Customer=="ALL")%>%
  arrange(Customer)
#packgroup.bmmb.oi <- format(packgroup.bmmb.oi,scientific=F) 
#packgroup.bmmb.oi <- data.frame(lapply(packgroup.bmmb.oi, gsub, pattern='NA', replacement=''))
write_xlsx(packgroup.bmmb.oi,paste0(output.path,"/packgroup.bmmb.oi.xlsx"))

packgroup.bmmb.sumatera <- packgroup.bmmb.df2%>%
  filter(str_detect(Customer,"DT SUMATERA")|Customer=="ALL")%>%
  arrange(Customer)
#packgroup.bmmb.sumatera <- format(packgroup.bmmb.sumatera,scientific=F)
#packgroup.bmmb.sumatera <- data.frame(lapply(packgroup.bmmb.sumatera, gsub, pattern='NA', replacement=''))
write_xlsx(packgroup.bmmb.sumatera,paste0(output.path,"/packgroup.bmmb.sumatera.xlsx"))

packgroup.bmmb.wj <- packgroup.bmmb.df2%>%
  filter(str_detect(Customer,"DT WEST JAVA")|Customer=="ALL")%>%
  arrange(Customer)
#packgroup.bmmb.wj <- format(packgroup.bmmb.wj,scientific=F) 
#packgroup.bmmb.wj <- data.frame(lapply(packgroup.bmmb.wj, gsub, pattern='NA', replacement=''))
write_xlsx(packgroup.bmmb.wj,paste0(output.path,"/packgroup.bmmb.wj.xlsx"))

#rename the column headers as per the output for each packgroup {Bango Manis small Bottle : bmsb}

packgroup.bmsb.headers <- c("Month",
                            "Sales Organization",
                            "Category",
                            "Sector",
                            "Brand",
                            "Product",
                            "Channel",
                            "Customer",
                            "Bango Manis Small Bottle_GSV",
                            "Bango Manis Small Bottle_Primary Sales Qty(PC)",
                            "Bango Manis Small Bottle_NIV",
                            "Bango Manis Small Bottle_Primary Sales Qty(Kg)",
                            "Bango Manis Small Bottle_Primary Sales Qty(Tonns)",
                            "Bango Manis Small Bottle_Sec Sales Value",
                            "Bango Manis Small Bottle_Sec Sales Volume(PC)",
                            "Bango Manis Small Bottle_Sec sales Volume (Kg)",
                            "Bango Manis Small Bottle_Sec sales Volume (Tonns)",
                            "Bango Manis Small Bottle_Secondary Stock Value [@ DT Rate.]",
                            "Bango Manis Small Bottle_Secondary Stock Volume(PC)",
                            "Bango Manis Small Bottle_Secondary stock volume(Kg)",
                            "Bango Manis Small Bottle_Secondary stock volume(Tonns)",
                            "Bango Manis Small Bottle_Original Order Qty.(PC)",
                            "Bango Manis Small Bottle_Final Customer Expected Order Qty.(PC)",
                            "Bango Manis Small Bottle_Dispatched Qty.(PC)",
                            "Bango Manis Small Bottle_DROO",
                            "Bango Manis Small Bottle_DR",
                            "Bango Manis Small Bottle_I trust Line",
                            "Bango Manis Small Bottle_I trust Total",
                            "Bango Manis Small Bottle_I trust %",
                            "Bango Manis Small Bottle_TTS",
                            "Bango Manis Small Bottle_BBT",
                            "Bango Manis Small Bottle_BBT-Place",
                            "Bango Manis Small Bottle_BBT-Place on invoice",
                            "Bango Manis Small Bottle_BBT-Place off invoice",
                            "Bango Manis Small Bottle_BBT-Price",
                            "Bango Manis Small Bottle_CPP on invoice",
                            "Bango Manis Small Bottle_CPP off invoice",
                            "Bango Manis Small Bottle_BBT-Product",
                            "Bango Manis Small Bottle_BBT-Pack",
                            "Bango Manis Small Bottle_BBT-Proposition",
                            "Bango Manis Small Bottle_BBT-Promotion",
                            "Bango Manis Small Bottle_EOT",
                            "Bango Manis Small Bottle_Brand & Marketing Investment",
                            "Bango Manis Small Bottle_Brand & Marketing Investment Trade",
                            "Bango Manis Small Bottle_Brand & Marketing Investment Consumer",
                            "Bango Manis Small Bottle_Promotion Team Cost Trade",
                            "Bango Manis Small Bottle_Promotion Team Cost Consumer",
                            "Bango Manis Small Bottle_Promotion Team Cost",
                            "Bango Manis Small Bottle_Promotion Packaging Material Cost Trade",
                            "Bango Manis Small Bottle_Promotion Packaging Material Cost Consumer",
                            "Bango Manis Small Bottle_Promotion Communication Material Cost Trade",
                            "Bango Manis Small Bottle_Promotion Communication Material Cost Consumer",
                            "Bango Manis Small Bottle_Promotion Communication Material Cost",
                            "Bango Manis Small Bottle_Promo Agency Remuneration Fees & Commissions Consumer",
                            "Bango Manis Small Bottle_Promotional Expenses",
                            "Bango Manis Small Bottle_Promotion Packaging Material Cost",
                            "Bango Manis Small Bottle_Promotion Repacking Cost",
                            "Bango Manis Small Bottle_Promotion Repacking Cost Trade",
                            "Bango Manis Small Bottle_Promotion Repacking Cost Consumer",
                            "Bango Manis Small Bottle_Promo Samples, Gifts and Incentive Costs",
                            "Bango Manis Small Bottle_Promo Samples, Gifts and Incentive Costs Consumer",
                            "Bango Manis Small Bottle_Promo Agency Remun Fees & Commissions",
                            "Bango Manis Small Bottle_Promo Agency Remuneration Fees & Commissions Trade",
                            "Bango Manis Small Bottle_Internal_Penetration",
                            "Bango Manis Small Bottle_Category_Penetration",
                            "Bango Manis Small Bottle_Internal_penetration%"
                            
)

colnames(packgroup.bmsb.df2) <- packgroup.bmsb.headers
#packgroup.bmsb.df2 <- data.frame(lapply(packgroup.bmsb.df2, gsub, pattern='NA', replacement=''))


packgroup.bmsb.cej <- packgroup.bmsb.df2%>%
  filter(str_detect(Customer,"DT CENTRAL & EAST JAVA")|Customer=="ALL")%>%
  arrange(Customer)
#packgroup.bmsb.cej <- format(packgroup.bmsb.cej,scientific=F)
#packgroup.bmsb.cej <- data.frame(lapply(packgroup.bmsb.cej, gsub, pattern='NA', replacement=''))
write_xlsx(packgroup.bmsb.cej,paste0(output.path,"/packgroup.bmsb.cej.xlsx"))

packgroup.bmsb.oi <- packgroup.bmsb.df2%>%
  filter(str_detect(Customer,"DT OI")|Customer=="ALL")%>%
  arrange(Customer)
#packgroup.bmsb.oi <- format(packgroup.bmsb.oi,scientific=F)
#packgroup.bmsb.cej <- data.frame(lapply(packgroup.bmsb.cej, gsub, pattern='NA', replacement=''))
write_xlsx(packgroup.bmsb.oi,paste0(output.path,"/packgroup.bmsb.oi.xlsx"))

packgroup.bmsb.sumatera <- packgroup.bmsb.df2%>%
  filter(str_detect(Customer,"DT SUMATERA")|Customer=="ALL")%>%
  arrange(Customer)
#packgroup.bmsb.sumatera <- format(packgroup.bmsb.sumatera,scientific=F)
#packgroup.bmsb.cej <- data.frame(lapply(packgroup.bmsb.cej, gsub, pattern='NA', replacement=''))
write_xlsx(packgroup.bmsb.sumatera,paste0(output.path,"/packgroup.bmsb.sumatera.xlsx"))

packgroup.bmsb.wj <- packgroup.bmsb.df2%>%
  filter(str_detect(Customer,"DT WEST JAVA")|Customer=="ALL")%>%
  arrange(Customer)
#packgroup.bmsb.wj <- format(packgroup.bmsb.wj,scientific=F) 
#packgroup.bmsb.cej <- data.frame(lapply(packgroup.bmsb.cej, gsub, pattern='NA', replacement=''))
write_xlsx(packgroup.bmsb.wj,paste0(output.path,"/packgroup.bmsb.wj.xlsx"))

#rename the column headers as per the output for each packgroup {Bango Manis Small Pouch : bmsp}

packgroup.bmsp.headers <- c("Month",
                            "Sales Organization",
                            "Category",
                            "Sector",
                            "Brand",
                            "Product",
                            "Channel",
                            "Customer",
                            "Bango Manis Small Pouch_GSV",
                            "Bango Manis Small Pouch_Primary Sales Qty(PC)",
                            "Bango Manis Small Pouch_NIV",
                            "Bango Manis Small Pouch_Primary Sales Qty(Kg)",
                            "Bango Manis Small Pouch_Primary Sales Qty(Tonns)",
                            "Bango Manis Small Pouch_Sec Sales Value",
                            "Bango Manis Small Pouch_Sec Sales Volume(PC)",
                            "Bango Manis Small Pouch_Sec sales Volume (Kg)",
                            "Bango Manis Small Pouch_Sec sales Volume (Tonns)",
                            "Bango Manis Small Pouch_Secondary Stock Value [@ DT Rate.]",
                            "Bango Manis Small Pouch_Secondary Stock Volume(PC)",
                            "Bango Manis Small Pouch_Secondary stock volume(Kg)",
                            "Bango Manis Small Pouch_Secondary stock volume(Tonns)",
                            "Bango Manis Small Pouch_Original Order Qty.(PC)",
                            "Bango Manis Small Pouch_Final Customer Expected Order Qty.(PC)",
                            "Bango Manis Small Pouch_Dispatched Qty.(PC)",
                            "Bango Manis Small Pouch_DROO",
                            "Bango Manis Small Pouch_DR",
                            "Bango Manis Small Pouch_I trust Line",
                            "Bango Manis Small Pouch_I trust Total",
                            "Bango Manis Small Pouch_I trust %",
                            "Bango Manis Small Pouch_TTS",
                            "Bango Manis Small Pouch_BBT",
                            "Bango Manis Small Pouch_BBT-Place",
                            "Bango Manis Small Pouch_BBT-Place on invoice",
                            "Bango Manis Small Pouch_BBT-Place off invoice",
                            "Bango Manis Small Pouch_BBT-Price",
                            "Bango Manis Small Pouch_CPP on invoice",
                            "Bango Manis Small Pouch_CPP off invoice",
                            "Bango Manis Small Pouch_BBT-Product",
                            "Bango Manis Small Pouch_BBT-Pack",
                            "Bango Manis Small Pouch_BBT-Proposition",
                            "Bango Manis Small Pouch_BBT-Promotion",
                            "Bango Manis Small Pouch_EOT",
                            "Bango Manis Small Pouch_Brand & Marketing Investment",
                            "Bango Manis Small Pouch_Brand & Marketing Investment Trade",
                            "Bango Manis Small Pouch_Brand & Marketing Investment Consumer",
                            "Bango Manis Small Pouch_Promotion Team Cost Trade",
                            "Bango Manis Small Pouch_Promotion Team Cost Consumer",
                            "Bango Manis Small Pouch_Promotion Team Cost",
                            "Bango Manis Small Pouch_Promotion Packaging Material Cost Trade",
                            "Bango Manis Small Pouch_Promotion Packaging Material Cost Consumer",
                            "Bango Manis Small Pouch_Promotion Communication Material Cost Trade",
                            "Bango Manis Small Pouch_Promotion Communication Material Cost Consumer",
                            "Bango Manis Small Pouch_Promotion Communication Material Cost",
                            "Bango Manis Small Pouch_Promo Agency Remuneration Fees & Commissions Consumer",
                            "Bango Manis Small Pouch_Promotional Expenses",
                            "Bango Manis Small Pouch_Promotion Packaging Material Cost",
                            "Bango Manis Small Pouch_Promotion Repacking Cost",
                            "Bango Manis Small Pouch_Promotion Repacking Cost Trade",
                            "Bango Manis Small Pouch_Promotion Repacking Cost Consumer",
                            "Bango Manis Small Pouch_Promo Samples, Gifts and Incentive Costs",
                            "Bango Manis Small Pouch_Promo Samples, Gifts and Incentive Costs Consumer",
                            "Bango Manis Small Pouch_Promo Agency Remun Fees & Commissions",
                            "Bango Manis Small Pouch_Promo Agency Remuneration Fees & Commissions Trade",
                            "Bango Manis Small Pouch_Internal_Penetration",
                            "Bango Manis Small Pouch_Category_Penetration",
                            "Bango Manis Small Pouch_Internal_penetration%"
                            
)
colnames(packgroup.bmsp.df2) <- packgroup.bmsp.headers

#packgroup.bmsp.df2 <- data.frame(lapply(packgroup.bmsp.df2, gsub, pattern='NA', replacement=''))


packgroup.bmsp.cej <- packgroup.bmsp.df2%>%
  filter(str_detect(Customer,"DT CENTRAL & EAST JAVA")|Customer=="ALL")%>%
  arrange(Customer)
#packgroup.bmsp.cej <- format(packgroup.bmsp.cej,scientific=F)
#packgroup.bmsp.cej <- data.frame(lapply(packgroup.bmsp.cej, gsub, pattern='NA', replacement=''))
write_xlsx(packgroup.bmsp.cej,paste0(output.path,"/packgroup.bmsp.cej.xlsx"))


packgroup.bmsp.oi <- packgroup.bmsp.df2%>%
  filter(str_detect(Customer,"DT OI")|Customer=="ALL")%>%
  arrange(Customer)
#packgroup.bmsp.oi <- format(packgroup.bmsp.oi,scientific=F) 
#packgroup.bmsp.oi <- data.frame(lapply(packgroup.bmsp.oi, gsub, pattern='NA', replacement=''))
write_xlsx(packgroup.bmsp.oi,paste0(output.path,"/packgroup.bmsp.oi.xlsx"))


packgroup.bmsp.sumatera <- packgroup.bmsp.df2%>%
  filter(str_detect(Customer,"DT SUMATERA")|Customer=="ALL")%>%
  arrange(Customer)
#packgroup.bmsp.sumatera <- format(packgroup.bmsp.sumatera,scientific=F)
#packgroup.bmsp.sumatera <- data.frame(lapply(packgroup.bmsp.sumatera, gsub, pattern='NA', replacement=''))
write_xlsx(packgroup.bmsp.sumatera,paste0(output.path,"/packgroup.bmsp.sumatera.xlsx"))

packgroup.bmsp.wj <- packgroup.bmsp.df2%>%
  filter(str_detect(Customer,"DT WEST JAVA")|Customer=="ALL")%>%
  arrange(Customer)
#packgroup.bmsp.wj <- format(packgroup.bmsp.wj,scientific=F)
#packgroup.bmsp.wj <- data.frame(lapply(packgroup.bmsp.wj, gsub, pattern='NA', replacement=''))
write_xlsx(packgroup.bmsp.wj,paste0(output.path,"/packgroup.bmsp.wj.xlsx"))


############################################# ID Psales (Bango-Account MT Model) ###################################################

#select the required columns i.e. Required Basepack from the material mapping , Account/Region from the Customer Mapping and Conversion rate from the Conversion mapping 

psales.mt.bp.mapping <- psales.material.mt.mapping%>%
  select(c(`Material Code`,`REQUIRED BASEPACK`))

customer.account.mapping <- customer.mapping%>%
  select(c(`Customer Code`,`Account/Region`))

conversion.rate.mapping <- conversion.mapping%>%
  select(c(MATNR,Conversion))


psales.mt.input <- merge.data.frame(x = id.psales.bango,y =psales.mt.bp.mapping,by.x ="Material_Code",by.y = "Material Code")%>%
  merge.data.frame(x = .,y = customer.account.mapping,by.x = "Customer_Code_8D",by.y ="Customer Code")%>%
  merge.data.frame(x = .,y = conversion.rate.mapping,by.x = "Material_Code",by.y = "MATNR")


psales.mt <- psales.mt.input
#filter out the observations for which the the Account is "MT" and the required basepack is not 0

psales.mt.filtered <- psales.mt%>%
  filter(`REQUIRED BASEPACK`!=0)
psales.mt.filtered <- psales.mt.filtered%>%
  filter(str_detect(`Account/Region`,"MT"))
psales.mt.filtered%<>%mutate(psales_KG=`Sales Quantity`*Conversion,psales_Tonns=psales_KG/1000)
#group the observations and aggregate the KPI's based on {Month,Required Basepack,Account}
psales.mt.aggregate <- psales.mt.filtered%>%
  group_by(`Fiscal year/period`,`REQUIRED BASEPACK`,`Account/Region`)%>%
  summarise(`Gross Sales Value (GSV)`=sum(`Gross Sales Value (GSV)`),`Sales Quantity`=sum(`Sales Quantity`),NIV=sum(NIV),psales_KG=sum(psales_KG),psales_Tonns=sum(psales_Tonns))
#calculate the unit price- GSV/Sales_Quantity(Kg)
psales.mt.aggregate%<>%mutate(unit_price= NIV/psales_KG)
#perform the month mapping for the Fiscal Period column 
psales.mt.aggregate <- merge.data.frame(psales.mt.aggregate,month.mapping,by.x = "Fiscal year/period",by.y ="Primary sales" )
psales.mt.aggregate%<>%select(-c(`Sec stock_DROO_Penetration`,TTS_BMI,`Secondary sales`,`Fiscal year/period`,Sellout))
psales.mt.aggregate%<>%select(c(Month,`REQUIRED BASEPACK`,`Account/Region`,`Gross Sales Value (GSV)`:unit_price))
#rename the account/region column to Customer
psales.mt.aggregate <- setnames(psales.mt.aggregate,old ="Account/Region" ,new ="Customer")
############################################# ID DROO (Bango-Account MT Model) #####################################################

#select the required columns i.e. Required Basepack from the material mapping , Account/Region from the Customer Mapping and Conversion rate from the Conversion mapping 

product.bp.mapping <- product.mapping%>%
  select(c(`Material Code`,`REQUIRED BASEPACK`))

droo.mt.input <- merge.data.frame(x = id.droo.bango,y =product.bp.mapping,by.x ="Material_Code",by.y = "Material Code")%>%
  merge.data.frame(x = .,y = customer.account.mapping,by.x = "STP_Code",by.y ="Customer Code")%>%
  merge.data.frame(x = .,y = conversion.rate.mapping,by.x = "Material_Code",by.y = "MATNR")

droo.mt <- droo.mt.input

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
droo.mt.aggregate%<>%select(c(Month,everything()))
droo.mt.aggregate <- setnames(droo.mt.aggregate,old ="Account/Region" ,new ="Customer")
####################################################### Sell out Data Pre processing ( Bango-Account MT Model) ################################################################

sellout.mt.input.path <- "C:/Users/goura/OneDrive/Desktop/ID_Jarvis/Python ID KT/Input_files/Sellout"
sellout.files <- list.files(path = sellout.mt.input.path,pattern = "^Sell Out ID AF(.*)xlsx|XLSX$")
sellout.files <- paste0(sellout.mt.input.path,"/",sellout.files)
#read the Alfamart and Indomart data for the preprocessing 
sellout.alfamart.df1 <- read_xlsx(sellout.files,sheet = "SAT",guess_max = 1000,n_max = 2)
sellout.alfamart.df <- read_xlsx(sellout.files,sheet = "SAT",guess_max = 1000,skip = 2)
sellout.indomart.df1 <- read_xlsx(sellout.files,sheet = "IDM",guess_max = 1000,n_max = 2)
sellout.indomart.df <- read_xlsx(sellout.files,sheet = "IDM",guess_max = 1000,skip = 2)
#perform the preprocessing of the Input datasets "Indomart" and "Alfamart and store it in sellout.alfamart.pp and sellout.indomart.pp variable where pp stands for pre processed
sellout.alfamart.pp <- sellout_preprocessing(sellout.alfamart.df1,sellout.alfamart.df)
sellout.indomart.pp <- sellout_preprocessing(sellout.indomart.df1,sellout.indomart.df)
sellout.account.data <- rbind(sellout.alfamart.pp,sellout.indomart.pp)
#perform the mapping at various level as per the business logic - map with the psales data to fetch the Unit price for each required basepack
#map the material description of sellout Alfamart and Indomart with the sellout.product.mapping
sellout.account.basepack.mapped <- merge.data.frame(x = sellout.account.data,y =sellout.material.mapping,by.x ="MAP",by.y = "Material Code",all.x = T  )
sellout.account.basepack.mapped%<>%select(-`Material Desc`)
#perform the Unit price mapping for the alfamart and indomart : Create the Key using the variables {Month,Basepack,Customer}
sellout.account <- sellout.account.basepack.mapped%>%
  unite(col = "Key",c("Month","REQUIRED BASEPACK","Customer"),sep="",remove = F)

psales.mt.aggregate.key <- psales.mt.aggregate%>%
  unite(col="Key",c("Month","REQUIRED BASEPACK","Customer"),sep="",remove = F)
psales.mt.aggregate.key <- psales.mt.aggregate.key%>%
  select(c(Key,unit_price))
#fetch the unit price from the psales.mt.aggregate df using the key {Month,Basepack,Customer}
sellout.account.df <- merge.data.frame(sellout.account,psales.mt.aggregate.key,by = "Key",all.x = T)
# Calculate the sellout volumne based on the following logic as mentioned below 
#If unit price is #NA , replace NA's with 0 and Sell out Volume (Kg)= 0 If unit price=0, then sellout volume=0 else Sellout value/Unit Price
sellout.account.df[is.na(sellout.account.df$unit_price),"unit_price"] <- 0
#calculate the sellout volume using the condition ifelse(unit_price==0,sellout_volume=0,sellout_volume=sellout_value/unit_price)
sellout.account.df$Sellout_Volume <- ifelse(sellout.account.df$unit_price==0,0,(sellout.account.df$Sellout_Value/sellout.account.df$unit_price))
#remove the Required Basepack : Bango MT 1-2-3 and NA
sellout.account.df <- sellout.account.df%>%
  filter(!is.na(`REQUIRED BASEPACK`))

########################################################## sellout ( Basepack to Account level) #########################################
exclude_bp.list <- c("Bango Mat 1","Bango Mat 2","Bango Mat 3")
sellout.bp_account.df <- sellout.account.df%>%
  filter(!`REQUIRED BASEPACK` %in% exclude_bp.list)
#perform the grouping and aggregate for the KPI's ( sell out Volume and Sell out Value : sum() and Unit_price : mean())
sellout.bp_account.aggregate <- sellout.bp_account.df%>%
  group_by(Month,Customer,`REQUIRED BASEPACK`)%>%
  summarise(Sellout_Value=sum(Sellout_Value),Sellout_Volume=sum(Sellout_Volume),unit_price=mean(unit_price))
sellout.bp_account.aggregate%<>%select(Month:Sellout_Value,unit_price,Sellout_Volume)
########################################################### sellout ( Bango to Account level) ############################################
sellout.bango_account.df <- sellout.account.df%>%
  filter(!`REQUIRED BASEPACK` %in% c("Bango Mat 3"))

sellout.bango_account.aggregate <- sellout.bango_account.df%>%
  group_by(Month,Customer)%>%
  summarise(Sellout_Value=sum(Sellout_Value),Sellout_Volume=sum(Sellout_Volume),unit_price=mean(unit_price))
sellout.bango_account.aggregate%<>%select(Month:Sellout_Value,unit_price,Sellout_Volume)
################################################################# ID TTS (Bango-Account  Mapped - MT Model) #############################################
tts.mapped.mt.input <- id.tts.mapped.bango
#perform the banner mapping to fetch the channel for the mapped banners
tts.account.mapping <- mapping(input=tts.mapped.mt.input,map =banner.mapping, map_key1 = "Banner",map_key2 = "Banner code")
#filter out the 9 basepacks and all the account for which it's "MT" 
tts.mt.mapped <- tts.account.mapping%>%
  filter(`REQUIRED BASEPACK`!=0)
tts.mt.mapped %<>%filter(str_detect(`Account/Region`,"MT"))
#perform the grouping of the tts.mt.mapped based on the {Fiscal Year/Period,Required Base Pack,Account/Region}
tts.mt.mapped.aggregate <- tts.mt.mapped%>%
  group_by(`Fiscal year/period`,`REQUIRED BASEPACK`,`Account/Region`)%>%
  summarise(TTS=sum(TTS),BBT=sum(BBT),`BBT - Place`=sum(`BBT - Place`),`BBT - Place on invoice`=sum(`BBT - Place on invoice`),`BBT - Place off invoice`=sum(`BBT - Place off invoice`),`BBT - Price`=sum(`BBT - Price`),`CPP on invoice`=sum(`CPP on invoice`),`CPP off invoice`=sum(`CPP off invoice`),`BBT - Product`=sum(`BBT - Product`),`BBT - Pack`=sum(`BBT - Pack`),`BBT - Proposition`=sum(`BBT - Proposition`),`BBT - Promotion`=sum(`BBT - Promotion`),EOT=sum(EOT))
#################################################### ID TTS UNMAPPED (Bango-Account  MT MODEL)##########################################################################
tts.unmapped.mt.input <- unmapped.material.tts.mapping
tts.unmapped.mt.df <- tts.unmapped.mt.input%>%
  unite(col = "key",c("Banner","Trade Format Level 2","Local Sales Force 2(m.d.)","Local Sales Force 3(m.d.)","Key Customer Level3"),sep="",remove = F)
tts.unmapped.account.mapping <- merge.data.frame(tts.unmapped.mt.df,unmapped.tts.mapping,by.x = "key",by.y = "Key")
tts.mt.unmapped <- tts.unmapped.account.mapping%>%
  filter(`REQUIRED BASEPACK`!=0)
tts.mt.unmapped%<>%filter(str_detect(`Account/Region`,"MT"))
tts.mt.unmapped%<>%convert(num(TTS:EOT))
tts.mt.unmapped.aggregate <- tts.mt.unmapped%>%
  group_by(`Fiscal year/period`,`REQUIRED BASEPACK`,`Account/Region`)%>%
  summarise(TTS=sum(TTS),BBT=sum(BBT),`BBT - Place`=sum(`BBT - Place`),`BBT - Place on invoice`=sum(`BBT - Place on invoice`),`BBT - Place off invoice`=sum(`BBT - Place off invoice`),`BBT - Price`=sum(`BBT - Price`),`CPP on invoice`=sum(`CPP on invoice`),`CPP off invoice`=sum(`CPP off invoice`),`BBT - Product`=sum(`BBT - Product`),`BBT - Pack`=sum(`BBT - Pack`),`BBT - Proposition`=sum(`BBT - Proposition`),`BBT - Promotion`=sum(`BBT - Promotion`),EOT=sum(EOT))
#################################################### TTS Consolidation (Bango-Account  MT Model) ######################################################################
tts.mt.aggregate <-  rbind(tts.mt.mapped.aggregate,tts.mt.unmapped.aggregate)
#group by the Fiscal {Period,Account,Required Basepack} and summarise the KPI's
tts.mt.aggregate.final <- tts.mt.aggregate%>%
  group_by(`Fiscal year/period`,`REQUIRED BASEPACK`,`Account/Region`)%>%
  summarise(TTS=sum(TTS),BBT=sum(BBT),`BBT - Place`=sum(`BBT - Place`),`BBT - Place on invoice`=sum(`BBT - Place on invoice`),`BBT - Place off invoice`=sum(`BBT - Place off invoice`),`BBT - Price`=sum(`BBT - Price`),`CPP on invoice`=sum(`CPP on invoice`),`CPP off invoice`=sum(`CPP off invoice`),`BBT - Product`=sum(`BBT - Product`),`BBT - Pack`=sum(`BBT - Pack`),`BBT - Proposition`=sum(`BBT - Proposition`),`BBT - Promotion`=sum(`BBT - Promotion`),EOT=sum(EOT))
#perform the Month mapping for the aggregate data and format the data into non scientific format 
tts.mt.aggregate.final.df <- merge.data.frame(x = tts.mt.aggregate.final,y = month.mapping,by.x = "Fiscal year/period" , by.y = "TTS_BMI")
tts.mt.aggregate.final.df%<>%select(-c(`Primary sales`:`Sec stock_DROO_Penetration`,Sellout,`Fiscal year/period`))
tts.mt.aggregate.final.df%<>%select(c(Month,everything()))
tts.mt.aggregate.final.df <- setnames(tts.mt.aggregate.final.df,old = "Account/Region" ,new ="Customer")
#################################################### BMI Mapped (Bango-Account MT Model) ##############################################################################
bmi.mapped.mt.input <- id.bmi.mapped.df
#map the sub brands and fetch the required basepacks 
sub_brand_basepack <- subbrand.mapping%>%
  select(c(`Sub brands`,Mapping_Basepack))
bmi.mapped.mt.data <- merge.data.frame(x = bmi.mapped.mt.input,y = sub_brand_basepack,by.x = "Sub Brand (PH)" ,by.y = "Sub brands" )
#fetch out the observations for which Mapping_Basepack is Bango
bmi.mapped.mt.data <- subset(x = bmi.mapped.mt.data ,subset = Mapping_Basepack=="Bango")
#mapped the customers from the banner mapping file 
bmi.mapped.mt.cust <- merge.data.frame(bmi.mapped.mt.data,y = banner.mapping,by.x = "Banner",by.y ="Banner code")
#filter out the observations for MT and All
bmi.mapped.mt.cust.filtered <- bmi.mapped.mt.cust%>%
  filter(str_detect(`Account/Region`,"MT"))
#aggregate the KPI's after grouping them
bmi.mapped.mt.filtered.aggregate <- bmi.mapped.mt.cust.filtered%>%
  group_by(`Fiscal year/period`,Mapping_Basepack,`Account/Region`)%>%
  summarise(X..Brand...Marketing.Investment=sum(X..Brand...Marketing.Investment),X..Brand...Marketing.Investment.Trade=sum(X..Brand...Marketing.Investment.Trade),X..Brand...Marketing.Investment.Consumer=sum(X..Brand...Marketing.Investment.Consumer),X..Promotional.Expenses=sum(X..Promotional.Expenses),X....Promotion.Packaging.Material.Cost=sum(X....Promotion.Packaging.Material.Cost),X......Promotion.Packaging.Material.Cost.Trade=sum(X......Promotion.Packaging.Material.Cost.Trade),X......Promotion.Packaging.Material.Cost.Consumer=sum(X......Promotion.Packaging.Material.Cost.Consumer),X....Promotion.Repacking.Cost=sum(X....Promotion.Repacking.Cost),X......Promotion.Repacking.Cost.Trade=sum(X....Promotion.Repacking.Cost),X......Promotion.Repacking.Cost.Trade=sum(X......Promotion.Repacking.Cost.Trade),X......Promotion.Repacking.Cost.Consumer=sum(X......Promotion.Repacking.Cost.Consumer),X....Promotion.Communication.Material.Cost=sum(X....Promotion.Communication.Material.Cost),X......Promotion.Communication.Material.Cost.Trade=sum(X......Promotion.Communication.Material.Cost.Trade),X......Promotion.Communication.Material.Cost.Consumer=sum(X......Promotion.Communication.Material.Cost.Consumer),X....Promo.Samples..Gifts.and.Incentive.Costs=sum(X....Promo.Samples..Gifts.and.Incentive.Costs),X......Promo.Samples..Gifts.and.Incentive.Costs.Consumer=sum(X......Promo.Samples..Gifts.and.Incentive.Costs.Consumer),X....Promotion.Team.Cost=sum(X....Promotion.Team.Cost),X......Promotion.Team.Cost.Trade=sum(X......Promotion.Team.Cost.Trade),X......Promotion.Team.Cost.Consumer=sum(X......Promotion.Team.Cost.Consumer),X....Promo.Agency.Remun.Fees...Commissions=sum(X....Promo.Agency.Remun.Fees...Commissions),X......Promo.Agency.Remuneration.Fees...Commissions.Trade=sum(X......Promo.Agency.Remuneration.Fees...Commissions.Trade),X......Promo.Agency.Remuneration.Fees...Commissions.Consumer=sum(X......Promo.Agency.Remuneration.Fees...Commissions.Consumer))
#rename the KPI's columns
bmi.mapped.mt.filtered.aggregate1 <- bmi.mapped.mt.filtered.aggregate%>%
  dplyr::rename(`Brand & Marketing Investment`=X..Brand...Marketing.Investment,`Brand & Marketing Investment Trade`=X..Brand...Marketing.Investment.Trade,`Brand & Marketing Investment Consumer`=X..Brand...Marketing.Investment.Consumer,`Promotional Expenses`=X..Promotional.Expenses,`Promotion Packaging Material Cost`=X....Promotion.Packaging.Material.Cost,`Promotion Communication Material Cost Trade`=X......Promotion.Communication.Material.Cost.Trade,`Promotion Packaging Material Cost Consumer`=X......Promotion.Packaging.Material.Cost.Consumer,`Promotion Repacking Cost`=X....Promotion.Repacking.Cost,`Promotion Repacking Cost Trade`=X......Promotion.Repacking.Cost.Trade,`Promotion Repacking Cost Consumer`=X......Promotion.Repacking.Cost.Consumer,`Promotion Communication Material Cost`=X....Promotion.Communication.Material.Cost,`Promotion Communication Material Cost Consumer`=X......Promotion.Communication.Material.Cost.Consumer,`Promo Samples, Gifts and Incentive Costs` =X....Promo.Samples..Gifts.and.Incentive.Costs,`Promo Samples, Gifts and Incentive Costs Consumer`=X......Promo.Samples..Gifts.and.Incentive.Costs.Consumer,`Promotion Team Cost`=X....Promotion.Team.Cost,`Promotion Team Cost Trade`=X......Promotion.Team.Cost.Trade,`Promotion Team Cost Consumer`=X......Promotion.Team.Cost.Consumer,`Promo Agency Remun Fees & Commissions`=X....Promo.Agency.Remun.Fees...Commissions,`Promo Agency Remuneration Fees & Commissions Trade`=X......Promo.Agency.Remuneration.Fees...Commissions.Trade,`Promo Agency Remuneration Fees & Commissions Consumer`=X......Promo.Agency.Remuneration.Fees...Commissions.Consumer)
bmi.mapped.mt.filtered.aggregate1%<>%dplyr::rename(`Promotion Packaging Material Cost Trade`=X......Promotion.Packaging.Material.Cost.Trade)

##################################################### BMI Unmapped (Bango-Account MT Model) #####################################################
bmi.unmapped.mt.input <- id.bmi.unmapped.df1
sub_brand_basepack <- subbrand.mapping%>%
  select(c(`Sub brands`,Mapping_Basepack))

bmi.unmapped.mt.data <- merge.data.frame(x = bmi.unmapped.mt.input,y = sub_brand_basepack,by.x = "Sub Brand (PH)" ,by.y = "Sub brands" )
#fetch out the observations for which Mapping_Basepack is Bango
bmi.unmapped.mt.data <- subset(x = bmi.unmapped.mt.data,subset = Mapping_Basepack=="BANGO")

#bmi unmapped data mapping for the key 
unmapped.bmi.key <- unmapped.bmi.mapping[,c(1,4)]
#merge the account/region based on the key 
bmi.unmapped.mt.data.mapped <- merge.data.frame(bmi.unmapped.mt.data,unmapped.bmi.key,by.x = "key",by.y = "Key")
#filter out the accounts which is contain "MT" and "All"
bmi.unmapped.mt.data.mapped.filtered <- bmi.unmapped.mt.data.mapped%>%
  filter(str_detect(`Account/Region`,"MT"))
#convert the KPI's from character to numeric dt
bmi.unmapped.mt.data.mapped.filtered %<>%convert(num(`Brand & Marketing Investment`:`Promo Agency Remuneration Fees & Commissions Consumer`))
#perfor the group by and summarise the KPI's for the BMI unmapped 
bmi.unmapped.mt.aggregate <- bmi.unmapped.mt.data.mapped.filtered%>%
  group_by(`Fiscal year/period`,Mapping_Basepack,`Account/Region`)%>%
  summarise(`Brand & Marketing Investment`=sum(`Brand & Marketing Investment`),`Brand & Marketing Investment Trade`=sum(`Brand & Marketing Investment Trade`),`Brand & Marketing Investment Consumer`=sum(`Brand & Marketing Investment Consumer`),`Promotional Expenses`=sum(`Promotional Expenses`),`Promotion Packaging Material Cost`=sum(`Promotion Packaging Material Cost`),`Promotion Packaging Material Cost Trade`=sum(`Promotion Repacking Cost Trade`),`Promotion Packaging Material Cost Consumer`=sum(`Promotion Packaging Material Cost Consumer`),`Promotion Repacking Cost`=sum(`Promotion Repacking Cost`),`Promotion Repacking Cost Trade`=sum(`Promotion Repacking Cost Trade`),`Promotion Repacking Cost Consumer`=sum(`Promotion Repacking Cost Consumer`),`Promotion Communication Material Cost`=sum(`Promotion Communication Material Cost`),`Promotion Communication Material Cost Trade`=sum(`Promotion Communication Material Cost Trade`),`Promotion Communication Material Cost Consumer`=sum(`Promotion Communication Material Cost Consumer`),`Promo Samples, Gifts and Incentive Costs`=sum(`Promo Samples, Gifts and Incentive Costs`),`Promo Samples, Gifts and Incentive Costs Consumer`=sum(`Promo Samples, Gifts and Incentive Costs Consumer`),`Promotion Team Cost`=sum(`Promotion Team Cost`),`Promotion Team Cost Trade`=sum(`Promotion Team Cost Trade`),`Promotion Team Cost Consumer`=sum(`Promotion Team Cost Consumer`),`Promo Agency Remun Fees & Commissions`=sum(`Promo Agency Remun Fees & Commissions`),`Promo Agency Remuneration Fees & Commissions Trade`=sum(`Promo Agency Remuneration Fees & Commissions Trade`),`Promo Agency Remuneration Fees & Commissions Consumer`=sum(`Promo Agency Remuneration Fees & Commissions Consumer`))

###################################################### BMI Consolidation ( Bango-Account Mapped + Unmapped) #####################################################
bmi.consolidated.data <- rbind(bmi.mapped.mt.filtered.aggregate1,bmi.unmapped.mt.aggregate)
bmi.consolidated.data1 <- bmi.consolidated.data%>%
  group_by(`Fiscal year/period`,Mapping_Basepack,`Account/Region`)%>%
  summarise(`Brand & Marketing Investment`=sum(`Brand & Marketing Investment`),`Brand & Marketing Investment Trade`=sum(`Brand & Marketing Investment Trade`),`Brand & Marketing Investment Consumer`=sum(`Brand & Marketing Investment Consumer`),`Promotional Expenses`=sum(`Promotional Expenses`),`Promotion Packaging Material Cost`=sum(`Promotion Packaging Material Cost`),`Promotion Packaging Material Cost Trade`=sum(`Promotion Repacking Cost Trade`),`Promotion Packaging Material Cost Consumer`=sum(`Promotion Packaging Material Cost Consumer`),`Promotion Repacking Cost`=sum(`Promotion Repacking Cost`),`Promotion Repacking Cost Trade`=sum(`Promotion Repacking Cost Trade`),`Promotion Repacking Cost Consumer`=sum(`Promotion Repacking Cost Consumer`),`Promotion Communication Material Cost`=sum(`Promotion Communication Material Cost`),`Promotion Communication Material Cost Trade`=sum(`Promotion Communication Material Cost Trade`),`Promotion Communication Material Cost Consumer`=sum(`Promotion Communication Material Cost Consumer`),`Promo Samples, Gifts and Incentive Costs`=sum(`Promo Samples, Gifts and Incentive Costs`),`Promo Samples, Gifts and Incentive Costs Consumer`=sum(`Promo Samples, Gifts and Incentive Costs Consumer`),`Promotion Team Cost`=sum(`Promotion Team Cost`),`Promotion Team Cost Trade`=sum(`Promotion Team Cost Trade`),`Promotion Team Cost Consumer`=sum(`Promotion Team Cost Consumer`),`Promo Agency Remun Fees & Commissions`=sum(`Promo Agency Remun Fees & Commissions`),`Promo Agency Remuneration Fees & Commissions Trade`=sum(`Promo Agency Remuneration Fees & Commissions Trade`),`Promo Agency Remuneration Fees & Commissions Consumer`=sum(`Promo Agency Remuneration Fees & Commissions Consumer`))
#perform the month mapping for the consolidated BMI list

bmi.consolidated.data1 <- merge.data.frame(bmi.consolidated.data1,month.mapping,by.x ="Fiscal year/period" ,by.y = "TTS_BMI")
bmi.consolidated.data1 <- bmi.consolidated.data1%>%
  select(-c(`Primary sales`,`Secondary sales`,`Sec stock_DROO_Penetration`,Sellout,`Fiscal year/period`))
bmi.consolidated.data1%<>%select(c(Month,everything()))

#replicate the Bango Data across the 9 bp's BANGO KECAP MANIS 620ML
#BANGO KECAP MANIS 275ML (bkm275)
#BANGO KECAP MANIS 220ML (bkm220)
#BANGO KECAP MANIS PEDAS 220ML (bkmp220)
#BANGO KECAP MANIS PEDAS 135ML (bkmp135)
#BANGO SOY/OYSTER/FISH SAUCE MANIS 60ML (bsofsm60)
#BANGO SOY/OYSTR/FISH SUCE MNIS SCHT 30ML(bsofsm30)
#BANGO KECAP MANIS 135ML(bkm135)
#BANGO SOY/OYSTER/FISH KECAP MANIS 580ML(bsofsm580)
#BANGO KECAP MANIS 620ML (bkm620)

bmi.bkm275 <- bmi.consolidated.data1 ;bmi.bkm220 <- bmi.consolidated.data1
bmi.bkmp220 <- bmi.consolidated.data1 ;bmi.bkmp135 <- bmi.consolidated.data1
bmi.bsofsm60 <- bmi.consolidated.data1 ; bmi.bsofsm30 <- bmi.consolidated.data1
bmi.bkm135 <-   bmi.consolidated.data1;bmi.bsofsm580 <- bmi.consolidated.data1
bmi.bkm620 <- bmi.consolidated.data1
#Add the product column for each basepacks
bmi.bkm275$Product <- "BANGO KECAP MANIS 275ML";bmi.bkm220$Product <- "BANGO KECAP MANIS 220ML"
bmi.bkmp220$Product <- "BANGO KECAP MANIS PEDAS 220ML" ; bmi.bkmp135$Product <- "BANGO KECAP MANIS PEDAS 135ML"
bmi.bsofsm60$Product <- "BANGO SOY/OYSTER/FISH SAUCE MANIS 60ML" ; bmi.bsofsm30$Product <- "BANGO SOY/OYSTR/FISH SUCE MNIS SCHT 30ML"
bmi.bkm135$Product <-  "BANGO KECAP MANIS 135ML"; bmi.bsofsm580$Product <- "BANGO SOY/OYSTER/FISH KECAP MANIS 580ML"
bmi.bkm620$Product <- "BANGO KECAP MANIS 620ML"
#perform the r bind to create the final consolidated bmi data frame

bmi.mt.final <- rbind(bmi.bkm275,bmi.bkm220,bmi.bkmp220,bmi.bkmp135,bmi.bsofsm60,bmi.bsofsm30,bmi.bkm135,bmi.bsofsm580,bmi.bkm620)
bmi.mt.final%<>%select(c(Month,Mapping_Basepack,Product,`Account/Region`,everything()))
bmi.mt.final <- setnames(bmi.mt.final,old = c("Account/Region","Product") ,new =c("Customer","REQUIRED BASEPACK"))

############################################################################# Consolidation Final( Bango-Account MT Model ) #########################################
psales.basepack.final <- psales.mt.aggregate ; psales.basepack.final%<>%select(-unit_price)
droo.basepack.final <- droo.mt.aggregate; sellout.basepack.final <- sellout.bp_account.aggregate
tts.basepack.final <- tts.mt.aggregate.final.df ; bmi.basepack.final <- bmi.mt.final%>%select(-Mapping_Basepack) 

#create the join key for the mapping  key : {Month,Required Basepack ,Customer}
pkey <- c("Month","REQUIRED BASEPACK","Customer")
consolidated.list1 <- merge.data.frame(x = psales.basepack.final,droo.basepack.final,by = pkey,all.y = T)
consolidated.list2 <- merge.data.frame(x=consolidated.list1,y = tts.basepack.final,all.y = T)
consolidated.list3 <- merge.data.frame(x = consolidated.list2,y = bmi.basepack.final,all.y = T)
consolidated.list4 <- merge.data.frame(x=consolidated.list3,y=sellout.basepack.final,all.x = T)

#seggrgate the consolidated list  into 3 regions : {MT "Alfamart" + All } , {MT "Indomart" + All } ,{MT "Carefourr"+All}

alfamart.list <- consolidated.list4%>%
  filter(str_detect(Customer,"MT Alfa"))
alfamart.list$`Sales Organization` <- "9001";alfamart.list$Category <- "SAVOURY"
alfamart.list$Sector <- "SEASONINGS"  ; alfamart.list$Channel <- "MODERN TRADE"
alfamart.list$Brand <- "BANGO"
alfamart.list <- setnames(alfamart.list,old = "REQUIRED BASEPACK",new = "Product")
#rearrange the columns as per the output
alfamart.final.list <- alfamart.list%>%
  select(c(Month,`Sales Organization`,Category,Sector,Brand,Product,Channel,Customer,everything()))
# rename the column as per the output 

alfamart.headers <- c("Month",
                      "Sales Organization",
                      "Category",
                      "Sector" ,
                      "Brand",
                      "Product",
                      "Channel",
                      "Customer",
                      "GSV",
                      "Primary Sales Qty(PC)",
                      "NIV",
                      "Primary Sales Qty(Kg)",
                      "Primary Sales Qty(Tonns)",
                      "Original Order Qty.",
                      "Final Customer Expected Order Qty.",
                      "Dispatched Qty.",
                      "DR",
                      "DROO",
                      "TTS",
                      "BBT",
                      "BBT-Place",
                      "BBT-Place on invoice",
                      "BBT-Place off invoice",
                      "BBT-Price",
                      "CPP on invoice",
                      "CPP off invoice",
                      "BBT-Product",
                      "BBT-Pack",
                      "BBT-Proposition",
                      "BBT-Promotion",
                      "EOT",
                      "Brand & Marketing Investment",
                      "Brand & Marketing Investment Trade",
                      "Brand & Marketing Investment Consumer",
                      "Promotional Expenses",
                      "Promotion Packaging Material Cost",
                      "Promotion Packaging Material Cost Trade",
                      "Promotion Packaging Material Cost Consumer",
                      "Promotion Repacking Cost",
                      "Promotion Repacking Cost Trade",
                      "Promotion Repacking Cost Consumer",
                      "Promotion Communication Material Cost",
                      "Promotion Communication Material Cost Trade",
                      "Promotion Communication Material Cost Consumer",
                      "Promo Samples, Gifts and Incentive Costs",
                      "Promo Samples, Gifts and Incentive Costs Consumer",
                      "Promotion Team Cost",
                      "Promotion Team Cost Trade",
                      "Promotion Team Cost Consumer",
                      "Promo Agency Remun Fees & Commissions",
                      "Promo Agency Remuneration Fees & Commissions Trade",
                      "Promo Agency Remuneration Fees & Commissions Consumer",
                      "Sell out Value (IDR)",
                      "Unit Price",
                      "Sell out Volume (KGs)"
)

colnames(alfamart.final.list) <- alfamart.headers

indomart.list <- consolidated.list4%>%
  filter(str_detect(Customer,"MT Indo"))
indomart.list$`Sales Organization` <- "9001";indomart.list$Category <- "SAVOURY"
indomart.list$Sector <- "SEASONINGS"  ; indomart.list$Channel <- "MODERN TRADE"
indomart.list$Brand <- "BANGO"
indomart.list <- setnames(indomart.list,old = "REQUIRED BASEPACK",new = "Product")

indomart.final.list <- indomart.list%>%
  select(c(Month,`Sales Organization`,Category,Sector,Brand,Product,Channel,Customer,everything()))

indomart.headers <- c("Month",
                      "Sales Organization",
                      "Category",
                      "Sector" ,
                      "Brand",
                      "Product",
                      "Channel",
                      "Customer",
                      "Bango_GSV",
                      "Primary Sales Qty(PC)",
                      "Bango_NIV",
                      "Primary Sales Qty(Kg)",
                      "Primary Sales Qty(Tonns)",
                      "Bango_OriginalOrder Qty.",
                      "Bango_Final Customer Expected Order Qty.",
                      "Bango_Dispatched Qty.",
                      "Bango_DR",
                      "Bango_DROO",
                      "Bango_TTS",
                      "Bango_BBT",
                      "Bango_BBT-Place",
                      "Bango_BBT-Place on invoice",
                      "Bango_BBT-Place off invoice",
                      "Bango_BBT-Price",
                      "Bango_CPP on invoice",
                      "Bango_CPP off invoice",
                      "Bango_BBT-Product",
                      "Bango_BBT-Pack",
                      "Bango_BBT-Proposition",
                      "Bango_BBT-Promotion",
                      "Bango_EOT",
                      "Bango_Brand & Marketing Investment",
                      "Bango_Brand & Marketing Investment Trade",
                      "Brand & Marketing Investment Consumer",
                      "Bango_Promotional Expenses",
                      "Bango_Promotion Packaging Material Cost",
                      "Bango_Promotion Packaging Material Cost Trade",
                      "Bango_Promotion Packaging Material Cost Consumer",
                      "Bango_Promotion Repacking Cost",
                      "Bango_Promotion Repacking Cost Trade",
                      "Bango_Promotion Repacking Cost Consumer",
                      "Bango_Promotion Communication Material Cost",
                      "Bango_Promotion Communication Material Cost Trade",
                      "Bango_Promotion Communication Material Cost Consumer",
                      "Bango_Promo Samples, Gifts and Incentive Costs",
                      "Bango_Promo Samples, Gifts and Incentive Costs Consumer",
                      "Bango_Promotion Team Cost",
                      "Bango_Promotion Team Cost Trade",
                      "Bango_Promotion Team Cost Consumer",
                      "Bango_Promo Agency Remun Fees & Commissions",
                      "Bango_Promo Agency Remuneration Fees & Commissions Trade",
                      "Bango_Promo Agency Remuneration Fees & Commissions Consumer",
                      "Sell out Value (IDR)",
                      "Unit Price",
                      "Sell out Volume (KGs)"
)

colnames(indomart.final.list) <- indomart.headers





Carrefour.list <- consolidated.list4%>%
  filter(str_detect(Customer,"MT Carrefour"))
Carrefour.list%<>%select(-c(Sellout_Value,Sellout_Volume,unit_price))
Carrefour.list$`Sales Organization` <- "9001";Carrefour.list$Category <- "SAVOURY"
Carrefour.list$Sector <- "SEASONINGS"  ; Carrefour.list$Channel <- "MODERN TRADE"
Carrefour.list$Brand <- "BANGO"
Carrefour.list <- setnames(Carrefour.list,old = "REQUIRED BASEPACK",new = "Product")
Carrefour.final.list <- Carrefour.list%>%
  select(c(Month,`Sales Organization`,Category,Sector,Brand,Product,Channel,Customer,everything()))

Carrefour.headers <- c("Month",
                       "Sales Organization",
                       "Category",
                       "Sector" ,
                       "Brand",
                       "Product",
                       "Channel",
                       "Customer",
                       "GSV",
                       "Primary Sales Qty(PC)",
                       "NIV",
                       "Primary Sales Qty(Kg)",
                       "Primary Sales Qty(Tonns)",
                       "Original Order Qty.",
                       "Final Customer Expected Order Qty.",
                       "Dispatched Qty.",
                       "DR",
                       "DROO",
                       "TTS",
                       "BBT",
                       "BBT-Place",
                       "BBT-Place on invoice",
                       "BBT-Place off invoice",
                       "BBT-Price",
                       "CPP on invoice",
                       "CPP off invoice",
                       "BBT-Product",
                       "BBT-Pack",
                       "BBT-Proposition",
                       "BBT-Promotion",
                       "EOT",
                       "Brand & Marketing Investment",
                       "Brand & Marketing Investment Trade",
                       "Brand & Marketing Investment Consumer",
                       "Promotional Expenses",
                       "Promotion Packaging Material Cost",
                       "Promotion Packaging Material Cost Trade",
                       "Promotion Packaging Material Cost Consumer",
                       "Promotion Repacking Cost",
                       "Promotion Repacking Cost Trade",
                       "Promotion Repacking Cost Consumer",
                       "Promotion Communication Material Cost",
                       "Promotion Communication Material Cost Trade",
                       "Promotion Communication Material Cost Consumer",
                       "Promo Samples, Gifts and Incentive Costs",
                       "Promo Samples, Gifts and Incentive Costs Consumer",
                       "Promotion Team Cost",
                       "Promotion Team Cost Trade",
                       "Promotion Team Cost Consumer",
                       "Promo Agency Remun Fees & Commissions",
                       "Promo Agency Remuneration Fees & Commissions Trade",
                       "Promo Agency Remuneration Fees & Commissions Consumer")
colnames(Carrefour.final.list) <- Carrefour.headers
#seggregate the Alfamart , Indomart and Carrefour list based on basepacks
basepack_list <- c("BANGO KECAP MANIS 275ML","BANGO KECAP MANIS 620ML","BANGO KECAP MANIS 220ML","BANGO KECAP MANIS PEDAS 220ML","BANGO KECAP MANIS PEDAS 135ML",
                   "BANGO SOY/OYSTER/FISH SAUCE MANIS 60ML","BANGO SOY/OYSTR/FISH SUCE MNIS SCHT 30ML","BANGO KECAP MANIS 135ML","BANGO SOY/OYSTER/FISH KECAP MANIS 580ML")


alfamart.bkm275 <- alfamart.final.list%>%
  filter(Product==basepack_list[1])
data.frame(lapply(alfamart.bkm275, gsub, pattern=NA, replacement=''))
write_xlsx(alfamart.bkm275,paste0(output.path,"/alfamart.bkm275.xlsx"))

alfamart.bkm620 <- alfamart.final.list%>%
  filter(Product==basepack_list[2])
data.frame(lapply(alfamart.bkm620, gsub, pattern=NA, replacement=''))
write_xlsx(alfamart.bkm620,paste0(output.path,"/alfamart.bkm620.xlsx"))


alfamart.bkm220 <- alfamart.final.list%>%
  filter(Product==basepack_list[3])
data.frame(lapply(alfamart.bkm220, gsub, pattern=NA, replacement=''))
write_xlsx(alfamart.bkm220,paste0(output.path,"/alfamart.bkm220.xlsx"))



alfamart.bkmp220 <- alfamart.final.list%>%
  filter(Product==basepack_list[4])
data.frame(lapply(alfamart.bkmp220, gsub, pattern=NA, replacement=''))
write_xlsx(alfamart.bkmp220,paste0(output.path,"/alfamart.bkmp220.xlsx"))

alfamart.bkmp135 <- alfamart.final.list%>%
  filter(Product==basepack_list[5])
data.frame(lapply(alfamart.bkmp135, gsub, pattern=NA, replacement=''))
write_xlsx(alfamart.bkmp135,paste0(output.path,"/alfamart.bkmp135.xlsx"))

alfamart.bsofs60 <- alfamart.final.list%>%
  filter(Product==basepack_list[6])
data.frame(lapply(alfamart.bsofs60, gsub, pattern=NA, replacement=''))
write_xlsx(alfamart.bsofs60,paste0(output.path,"/alfamart.bsofs60.xlsx"))

alfamart.bsofms30 <- alfamart.final.list%>%
  filter(Product==basepack_list[7])
data.frame(lapply(alfamart.bsofms30, gsub, pattern=NA, replacement=''))
write_xlsx(alfamart.bsofms30,paste0(output.path,"/alfamart.bsofms30.xlsx"))

alfamart.bkm135 <- alfamart.final.list%>%
  filter(Product==basepack_list[8])
data.frame(lapply(alfamart.bkm135, gsub, pattern=NA, replacement=''))
write_xlsx(alfamart.bkm135,paste0(output.path,"/alfamart.bkm135.xlsx"))

alfamart.bsof580 <- alfamart.final.list%>%
  filter(Product==basepack_list[9])
data.frame(lapply(alfamart.bsof580, gsub, pattern=NA, replacement=''))
write_xlsx(alfamart.bsof580,paste0(output.path,"/alfamart.bsof580.xlsx"))

#indomart 
indomart.bkm275 <- indomart.final.list%>%
  filter(Product==basepack_list[1])
data.frame(lapply(indomart.bkm275, gsub, pattern=NA, replacement=''))
write_xlsx(indomart.bkm275,paste0(output.path,"/indomart.bkm275.xlsx"))

indomart.bkm620 <- indomart.final.list%>%
  filter(Product==basepack_list[2])
data.frame(lapply(indomart.bkm620, gsub, pattern=NA, replacement=''))
write_xlsx(indomart.bkm620,paste0(output.path,"/indomart.bkm620.xlsx"))


indomart.bkm220 <- indomart.final.list%>%
  filter(Product==basepack_list[3])
data.frame(lapply(indomart.bkm220, gsub, pattern=NA, replacement=''))
write_xlsx(indomart.bkm220,paste0(output.path,"/indomart.bkm220.xlsx"))

indomart.bkmp220 <- indomart.final.list%>%
  filter(Product==basepack_list[4])
data.frame(lapply(indomart.bkmp220, gsub, pattern=NA, replacement=''))
write_xlsx(indomart.bkmp220,paste0(output.path,"/indomart.bkmp220.xlsx"))


indomart.bkmp135 <- indomart.final.list%>%
  filter(Product==basepack_list[5])
data.frame(lapply(indomart.bkmp135, gsub, pattern=NA, replacement=''))
write_xlsx(indomart.bkmp135,paste0(output.path,"/indomart.bkmp135.xlsx"))

indomart.bsofs60 <- indomart.final.list%>%
  filter(Product==basepack_list[6])
data.frame(lapply(indomart.bsofs60, gsub, pattern=NA, replacement=''))
write_xlsx(indomart.bsofs60,paste0(output.path,"/indomart.bsofs60.xlsx"))

indomart.bsofms30 <- indomart.final.list%>%
  filter(Product==basepack_list[7])
data.frame(lapply(indomart.bsofms30, gsub, pattern=NA, replacement=''))
write_xlsx(indomart.bsofms30,paste0(output.path,"/indomart.bsofms30.xlsx"))

indomart.bkm135 <- indomart.final.list%>%
  filter(Product==basepack_list[8])
data.frame(lapply(indomart.bkm135, gsub, pattern=NA, replacement=''))
write_xlsx(indomart.bkm135,paste0(output.path,"/indomart.bkm135.xlsx"))


indomart.bsof580 <- indomart.final.list%>%
  filter(Product==basepack_list[9])
data.frame(lapply(indomart.bsof580, gsub, pattern=NA, replacement=''))
write_xlsx(indomart.bsof580,paste0(output.path,"/indomart.bsof580.xlsx"))

#carrefour
Carrefour.bkm275 <- Carrefour.final.list%>%
  filter(Product==basepack_list[1])
data.frame(lapply(Carrefour.bkm275, gsub, pattern=NA, replacement=''))
write_xlsx(Carrefour.bkm275,paste0(output.path,"/Carrefour.bkm275.xlsx"))


Carrefour.bkm620 <- Carrefour.final.list%>%
  filter(Product==basepack_list[2])
data.frame(lapply(Carrefour.bkm620, gsub, pattern=NA, replacement=''))
write_xlsx(Carrefour.bkm620,paste0(output.path,"/Carrefour.bkm620.xlsx"))


Carrefour.bkm220 <- Carrefour.final.list%>%
  filter(Product==basepack_list[3])
data.frame(lapply(Carrefour.bkm220, gsub, pattern=NA, replacement=''))
write_xlsx(Carrefour.bkm220,paste0(output.path,"/Carrefour.bkm220.xlsx"))


Carrefour.bkmp220 <- Carrefour.final.list%>%
  filter(Product==basepack_list[4])
data.frame(lapply(Carrefour.bkmp220, gsub, pattern=NA, replacement=''))
write_xlsx(Carrefour.bkmp220,paste0(output.path,"/Carrefour.bkmp220.xlsx"))


Carrefour.bkmp135 <- Carrefour.final.list%>%
  filter(Product==basepack_list[5])
data.frame(lapply(Carrefour.bkmp135, gsub, pattern=NA, replacement=''))
write_xlsx(Carrefour.bkmp135,paste0(output.path,"/Carrefour.bkmp135.xlsx"))


Carrefour.bsofs60 <- Carrefour.final.list%>%
  filter(Product==basepack_list[6])
data.frame(lapply(Carrefour.bsofs60, gsub, pattern=NA, replacement=''))
write_xlsx(Carrefour.bsofs60,paste0(output.path,"/Carrefour.bsofs60.xlsx"))


Carrefour.bsofms30 <- Carrefour.final.list%>%
  filter(Product==basepack_list[7])
data.frame(lapply(Carrefour.bsofms30, gsub, pattern=NA, replacement=''))
write_xlsx(Carrefour.bsofms30,paste0(output.path,"/Carrefour.bsofms30.xlsx"))


Carrefour.bkm135 <- Carrefour.final.list%>%
  filter(Product==basepack_list[8])
data.frame(lapply(Carrefour.bkm135, gsub, pattern=NA, replacement=''))
write_xlsx(Carrefour.bkm135,paste0(output.path,"/Carrefour.bkm135.xlsx"))


Carrefour.bsof580 <- Carrefour.final.list%>%
  filter(Product==basepack_list[9])
data.frame(lapply(Carrefour.bsof580, gsub, pattern=NA, replacement=''))
write_xlsx(Carrefour.bsof580,paste0(output.path,"/Carrefour.bsof580.xlsx"))

###################################################### ID Psales (Bango-Sub channel MT Model) ##################################################

cust.sch.map <- customer.mapping%>%
  select(c(`Customer Code`,Channel,`Sub channel`))

conv.sch.map <- conversion.mapping%>%
  select(c(MATNR,Conversion))


psales.sch.input <- merge.data.frame(x = id.psales.bango,y =product.mapping,by.x ="Material_Code",by.y = "Material Code")%>%
  merge.data.frame(x = .,y = cust.sch.map,by.x = "Customer_Code_8D",by.y ="Customer Code")%>%
  merge.data.frame(x = .,y = conv.sch.map,by.x = "Material_Code",by.y = "MATNR")
#filter out the sub channel which is HSM and Mini
sub_channel.list <- c("HSM","Mini")
psales.sch.input%<>%filter(`Sub channel` %in% sub_channel.list)
#rename the column and calculate the psales_kg and psales_tonns KPI's
psales.sch.df <- setnames(psales.sch.input,c("Material Desc.x"),c("Material Desc"))
psales.sch.df%<>%mutate(psales_KG=`Sales Quantity`*Conversion,psales_Tonns=psales_KG/1000)

#perform the group by and summarise the KPI's
psales.sch.aggregate <- psales.sch.df%>%
  group_by(`Fiscal year/period`,BRAND,`Sub channel`)%>%
  summarise(`Gross Sales Value (GSV)`=sum(`Gross Sales Value (GSV)`),`Sales Quantity`=sum(`Sales Quantity`),NIV=sum(NIV),psales_KG=sum(psales_KG),psales_Tonns=sum(psales_Tonns))
#perform the month mapping for the Fiscal Period column 
psales.sch.aggregate <- merge.data.frame(psales.sch.aggregate,month.mapping,by.x = "Fiscal year/period",by.y ="Primary sales" )
psales.sch.aggregate%<>%select(-c(`Secondary sales`:TTS_BMI,Sellout,`Fiscal year/period`))
psales.sch.aggregate%<>%select(c(Month,everything()))

###################################################### ID DROO (Bango-Sub channel MT Model) ##################################################
droo.sch.input <- merge.data.frame(x = id.droo.bango,y =product.mapping,by.x ="Material_Code",by.y = "Material Code")%>%
  merge.data.frame(x = .,y = cust.sch.map,by.x = "STP_Code",by.y ="Customer Code")%>%
  merge.data.frame(x = .,y = conv.sch.map,by.x = "Material_Code",by.y = "MATNR")

#filter out the sub channel which is HSM and Mini
droo.sch.input%<>%filter(`Sub channel`%in% sub_channel.list)
droo.sch.df <- setnames(droo.sch.input,c("Material Desc.x"),c("Material Desc"))
#summarise and aggregate the KPI's 
droo.sch.aggregate <- droo.sch.df%>%
  group_by(`Calendar Year/Month`,BRAND,`Sub channel`)%>%
  summarise(`OriginalOrder Qty`=sum(`OriginalOrder Qty`),`Final Customer Expected Order Qty`=sum(`Final Customer Expected Order Qty`),`Dispatched Qty`=sum(`Dispatched Qty`))

droo.sch.aggregate%<>%mutate(DR=(`Dispatched Qty`/`Final Customer Expected Order Qty`*100),DROO=(`Dispatched Qty`/`OriginalOrder Qty`*100))
#perform the month-year mapping 
droo.sch.aggregate <- merge.data.frame(droo.sch.aggregate,month.mapping,by.x = "Calendar Year/Month",by.y ="Sec stock_DROO_Penetration" )
droo.sch.aggregate%<>%select(-c(`Primary sales`,TTS_BMI,`Secondary sales`,`Calendar Year/Month`,Sellout))
droo.sch.aggregate[which(is.nan(droo.sch.aggregate$DR)),"DR"] <- 0
droo.sch.aggregate%<>%select(c(Month,everything()))

###################################################### ID Mapped TTS (Bango-Sub channel MT Model) ##################################################
tts.sch.mapped.input <- tts.account.mapping
#filter out the observation for which Sub Channel is {HSM,Mini}
tts.sch.mapped.input%<>%filter(`Sub channel` %in% sub_channel.list)
#group-by and aggregate the KPI's based on the key {Month,Brand,Sub Channel}
tts.sch.mapped.aggregate <- tts.sch.mapped.input%>%
  group_by(`Fiscal year/period`,BRAND,`Sub channel`)%>%
  summarise(TTS=sum(TTS),BBT=sum(BBT),`BBT - Place`=sum(`BBT - Place`),`BBT - Place on invoice`=sum(`BBT - Place on invoice`),`BBT - Place off invoice`=sum(`BBT - Place off invoice`),`BBT - Price`=sum(`BBT - Price`),`CPP on invoice`=sum(`CPP on invoice`),`CPP off invoice`=sum(`CPP off invoice`),`BBT - Product`=sum(`BBT - Product`),`BBT - Pack`=sum(`BBT - Pack`),`BBT - Proposition`=sum(`BBT - Proposition`),`BBT - Promotion`=sum(`BBT - Promotion`),EOT=sum(EOT))

###################################################### ID Un Mapped TTS (Bango-Sub channel MT Model) ##################################################
tts.sch.unmapped.input <- tts.unmapped.account.mapping
#filter out the observation for which Sub Channel is {HSM,Mini}
tts.sch.unmapped.input%<>%filter(`Sub channel` %in% sub_channel.list)
#convert the KPI's data type from Char to Numeric
tts.sch.unmapped.input%<>%convert(num(TTS:EOT))
#group-by and aggregate the KPI's based on the key {Month,Brand,Sub Channel}
tts.sch.unmapped.aggregate <- tts.sch.unmapped.input%>%
  group_by(`Fiscal year/period`,BRAND,`Sub channel`)%>%
  summarise(TTS=sum(TTS),BBT=sum(BBT),`BBT - Place`=sum(`BBT - Place`),`BBT - Place on invoice`=sum(`BBT - Place on invoice`),`BBT - Place off invoice`=sum(`BBT - Place off invoice`),`BBT - Price`=sum(`BBT - Price`),`CPP on invoice`=sum(`CPP on invoice`),`CPP off invoice`=sum(`CPP off invoice`),`BBT - Product`=sum(`BBT - Product`),`BBT - Pack`=sum(`BBT - Pack`),`BBT - Proposition`=sum(`BBT - Proposition`),`BBT - Promotion`=sum(`BBT - Promotion`),EOT=sum(EOT))
###################################################### ID TTS consolidation (Bango-Sub channel MT Model) ##################################################
tts.sch.aggregate <- rbind(tts.sch.mapped.aggregate,tts.sch.unmapped.aggregate)
tts.sch.aggregate <- tts.sch.aggregate%>%
  group_by(`Fiscal year/period`,BRAND,`Sub channel`)%>%
  summarise(TTS=sum(TTS),BBT=sum(BBT),`BBT - Place`=sum(`BBT - Place`),`BBT - Place on invoice`=sum(`BBT - Place on invoice`),`BBT - Place off invoice`=sum(`BBT - Place off invoice`),`BBT - Price`=sum(`BBT - Price`),`CPP on invoice`=sum(`CPP on invoice`),`CPP off invoice`=sum(`CPP off invoice`),`BBT - Product`=sum(`BBT - Product`),`BBT - Pack`=sum(`BBT - Pack`),`BBT - Proposition`=sum(`BBT - Proposition`),`BBT - Promotion`=sum(`BBT - Promotion`),EOT=sum(EOT))
#perform the month-year mapping by using the month mapping df
tts.sch.aggregate.final <- merge.data.frame(x = tts.sch.aggregate,y = month.mapping,by.x = "Fiscal year/period" , by.y = "TTS_BMI")
tts.sch.aggregate.final%<>%select(-c(`Primary sales`:`Sec stock_DROO_Penetration`,Sellout,`Fiscal year/period`))
tts.sch.aggregate.final%<>%select(c(Month,everything()))
####################################################### BMI Mapped (Bango-Sub channel MT Model) ###################################################
bmi.sch.mapped.input <- bmi.mapped.bango
#perform the mapping to fetch the sub channels 
sch.mapped.bmi.key <- banner.mapping[,c(2,4)]
bmi.sch.mapped.input <- merge.data.frame(bmi.sch.mapped.input,sch.mapped.bmi.key,by.x = "Banner",by.y = "Banner code",all.x = T)

#filter out the observation for which the sub channel is {HSM.Mini}
bmi.sch.mapped.input%<>%filter(`Sub channel` %in% sub_channel.list)
#perform the grouping and the aggregation 
bmi.sch.mapped.aggregate <- bmi.sch.mapped.input%>%
  group_by(`Fiscal year/period`,Brand,`Sub channel`)%>%
  summarise(X..Brand...Marketing.Investment=sum(X..Brand...Marketing.Investment),X..Brand...Marketing.Investment.Trade=sum(X..Brand...Marketing.Investment.Trade),X..Brand...Marketing.Investment.Consumer=sum(X..Brand...Marketing.Investment.Consumer),X..Promotional.Expenses=sum(X..Promotional.Expenses),X....Promotion.Packaging.Material.Cost=sum(X....Promotion.Packaging.Material.Cost),X......Promotion.Packaging.Material.Cost.Trade=sum(X......Promotion.Packaging.Material.Cost.Trade),X......Promotion.Packaging.Material.Cost.Consumer=sum(X......Promotion.Packaging.Material.Cost.Consumer),X....Promotion.Repacking.Cost=sum(X....Promotion.Repacking.Cost),X......Promotion.Repacking.Cost.Trade=sum(X....Promotion.Repacking.Cost),X......Promotion.Repacking.Cost.Trade=sum(X......Promotion.Repacking.Cost.Trade),X......Promotion.Repacking.Cost.Consumer=sum(X......Promotion.Repacking.Cost.Consumer),X....Promotion.Communication.Material.Cost=sum(X....Promotion.Communication.Material.Cost),X......Promotion.Communication.Material.Cost.Trade=sum(X......Promotion.Communication.Material.Cost.Trade),X......Promotion.Communication.Material.Cost.Consumer=sum(X......Promotion.Communication.Material.Cost.Consumer),X....Promo.Samples..Gifts.and.Incentive.Costs=sum(X....Promo.Samples..Gifts.and.Incentive.Costs),X......Promo.Samples..Gifts.and.Incentive.Costs.Consumer=sum(X......Promo.Samples..Gifts.and.Incentive.Costs.Consumer),X....Promotion.Team.Cost=sum(X....Promotion.Team.Cost),X......Promotion.Team.Cost.Trade=sum(X......Promotion.Team.Cost.Trade),X......Promotion.Team.Cost.Consumer=sum(X......Promotion.Team.Cost.Consumer),X....Promo.Agency.Remun.Fees...Commissions=sum(X....Promo.Agency.Remun.Fees...Commissions),X......Promo.Agency.Remuneration.Fees...Commissions.Trade=sum(X......Promo.Agency.Remuneration.Fees...Commissions.Trade),X......Promo.Agency.Remuneration.Fees...Commissions.Consumer=sum(X......Promo.Agency.Remuneration.Fees...Commissions.Consumer))

#rename the KPI's columns
bmi.sch.mapped.aggregate <- bmi.sch.mapped.aggregate%>%
  dplyr::rename(`Brand & Marketing Investment`=X..Brand...Marketing.Investment,`Brand & Marketing Investment Trade`=X..Brand...Marketing.Investment.Trade,`Brand & Marketing Investment Consumer`=X..Brand...Marketing.Investment.Consumer,`Promotional Expenses`=X..Promotional.Expenses,`Promotion Packaging Material Cost`=X....Promotion.Packaging.Material.Cost,`Promotion Communication Material Cost Trade`=X......Promotion.Communication.Material.Cost.Trade,`Promotion Packaging Material Cost Consumer`=X......Promotion.Packaging.Material.Cost.Consumer,`Promotion Repacking Cost`=X....Promotion.Repacking.Cost,`Promotion Repacking Cost Trade`=X......Promotion.Repacking.Cost.Trade,`Promotion Repacking Cost Consumer`=X......Promotion.Repacking.Cost.Consumer,`Promotion Communication Material Cost`=X....Promotion.Communication.Material.Cost,`Promotion Communication Material Cost Consumer`=X......Promotion.Communication.Material.Cost.Consumer,`Promo Samples, Gifts and Incentive Costs` =X....Promo.Samples..Gifts.and.Incentive.Costs,`Promo Samples, Gifts and Incentive Costs Consumer`=X......Promo.Samples..Gifts.and.Incentive.Costs.Consumer,`Promotion Team Cost`=X....Promotion.Team.Cost,`Promotion Team Cost Trade`=X......Promotion.Team.Cost.Trade,`Promotion Team Cost Consumer`=X......Promotion.Team.Cost.Consumer,`Promo Agency Remun Fees & Commissions`=X....Promo.Agency.Remun.Fees...Commissions,`Promo Agency Remuneration Fees & Commissions Trade`=X......Promo.Agency.Remuneration.Fees...Commissions.Trade,`Promo Agency Remuneration Fees & Commissions Consumer`=X......Promo.Agency.Remuneration.Fees...Commissions.Consumer)
bmi.sch.mapped.aggregate%<>%dplyr::rename(`Promotion Packaging Material Cost Trade`=X......Promotion.Packaging.Material.Cost.Trade)

####################################################### BMI Un Mapped (Bango-Sub channel MT Model) ###################################################
bmi.sch.unmapped.input <- id.bmi.unmapped.df1
#bmi unmapped data mapping for the key 
sch.unmapped.bmi.key <- unmapped.bmi.mapping[,c(1,3)]
#merge the account/region based on the key 
bmi.unmapped.sch.data.mapped <- merge.data.frame(bmi.sch.unmapped.input,sch.unmapped.bmi.key,by.x = "key",by.y = "Key")
#filter out the observation for which the sub channel is {HSM.Mini} and convert the KPI's from Char to num
bmi.unmapped.sch.data.mapped%<>%filter(`Sub channel` %in% sub_channel.list)
bmi.unmapped.sch.data.mapped%<>%convert(num(`Brand & Marketing Investment`:`Promo Agency Remuneration Fees & Commissions Consumer`))
#perfor the group by and summarise the KPI's for the BMI unmapped 
bmi.sch.unmapped.aggregate <- bmi.unmapped.sch.data.mapped%>%
  group_by(`Fiscal year/period`,Brand_Name,`Sub channel`)%>%
  summarise(`Brand & Marketing Investment`=sum(`Brand & Marketing Investment`),`Brand & Marketing Investment Trade`=sum(`Brand & Marketing Investment Trade`),`Brand & Marketing Investment Consumer`=sum(`Brand & Marketing Investment Consumer`),`Promotional Expenses`=sum(`Promotional Expenses`),`Promotion Packaging Material Cost`=sum(`Promotion Packaging Material Cost`),`Promotion Packaging Material Cost Trade`=sum(`Promotion Repacking Cost Trade`),`Promotion Packaging Material Cost Consumer`=sum(`Promotion Packaging Material Cost Consumer`),`Promotion Repacking Cost`=sum(`Promotion Repacking Cost`),`Promotion Repacking Cost Trade`=sum(`Promotion Repacking Cost Trade`),`Promotion Repacking Cost Consumer`=sum(`Promotion Repacking Cost Consumer`),`Promotion Communication Material Cost`=sum(`Promotion Communication Material Cost`),`Promotion Communication Material Cost Trade`=sum(`Promotion Communication Material Cost Trade`),`Promotion Communication Material Cost Consumer`=sum(`Promotion Communication Material Cost Consumer`),`Promo Samples, Gifts and Incentive Costs`=sum(`Promo Samples, Gifts and Incentive Costs`),`Promo Samples, Gifts and Incentive Costs Consumer`=sum(`Promo Samples, Gifts and Incentive Costs Consumer`),`Promotion Team Cost`=sum(`Promotion Team Cost`),`Promotion Team Cost Trade`=sum(`Promotion Team Cost Trade`),`Promotion Team Cost Consumer`=sum(`Promotion Team Cost Consumer`),`Promo Agency Remun Fees & Commissions`=sum(`Promo Agency Remun Fees & Commissions`),`Promo Agency Remuneration Fees & Commissions Trade`=sum(`Promo Agency Remuneration Fees & Commissions Trade`),`Promo Agency Remuneration Fees & Commissions Consumer`=sum(`Promo Agency Remuneration Fees & Commissions Consumer`))

bmi.sch.unmapped.aggregate <- setnames(bmi.sch.unmapped.aggregate,old = "Brand_Name",new = "Brand")

####################################################### BMI Consolidation (Bango-Sub channel MT Model) ###################################################
bmi.sch.aggregate <- rbind(bmi.sch.mapped.aggregate,bmi.sch.unmapped.aggregate)
#perform the group by and the aggregation of the KPi's
bmi.sch.aggregate.final <- bmi.sch.aggregate%>%
  group_by(`Fiscal year/period`,Brand,`Sub channel`)%>%
  summarise(`Brand & Marketing Investment`=sum(`Brand & Marketing Investment`),`Brand & Marketing Investment Trade`=sum(`Brand & Marketing Investment Trade`),`Brand & Marketing Investment Consumer`=sum(`Brand & Marketing Investment Consumer`),`Promotional Expenses`=sum(`Promotional Expenses`),`Promotion Packaging Material Cost`=sum(`Promotion Packaging Material Cost`),`Promotion Packaging Material Cost Trade`=sum(`Promotion Repacking Cost Trade`),`Promotion Packaging Material Cost Consumer`=sum(`Promotion Packaging Material Cost Consumer`),`Promotion Repacking Cost`=sum(`Promotion Repacking Cost`),`Promotion Repacking Cost Trade`=sum(`Promotion Repacking Cost Trade`),`Promotion Repacking Cost Consumer`=sum(`Promotion Repacking Cost Consumer`),`Promotion Communication Material Cost`=sum(`Promotion Communication Material Cost`),`Promotion Communication Material Cost Trade`=sum(`Promotion Communication Material Cost Trade`),`Promotion Communication Material Cost Consumer`=sum(`Promotion Communication Material Cost Consumer`),`Promo Samples, Gifts and Incentive Costs`=sum(`Promo Samples, Gifts and Incentive Costs`),`Promo Samples, Gifts and Incentive Costs Consumer`=sum(`Promo Samples, Gifts and Incentive Costs Consumer`),`Promotion Team Cost`=sum(`Promotion Team Cost`),`Promotion Team Cost Trade`=sum(`Promotion Team Cost Trade`),`Promotion Team Cost Consumer`=sum(`Promotion Team Cost Consumer`),`Promo Agency Remun Fees & Commissions`=sum(`Promo Agency Remun Fees & Commissions`),`Promo Agency Remuneration Fees & Commissions Trade`=sum(`Promo Agency Remuneration Fees & Commissions Trade`),`Promo Agency Remuneration Fees & Commissions Consumer`=sum(`Promo Agency Remuneration Fees & Commissions Consumer`))

#perform the month-year mapping by using the month mapping df
bmi.sch.aggregate.final <- merge.data.frame(x = bmi.sch.aggregate.final,y = month.mapping,by.x = "Fiscal year/period" , by.y = "TTS_BMI")
bmi.sch.aggregate.final%<>%select(-c(`Primary sales`:`Sec stock_DROO_Penetration`,Sellout,`Fiscal year/period`))
bmi.sch.aggregate.final%<>%select(c(Month,everything()))
bmi.sch.aggregate.final$Brand <- toupper(bmi.sch.aggregate.final$Brand)
bmi.sch.aggregate.final <- setnames(bmi.sch.aggregate.final,old = "Brand",new = "BRAND")

########################################################### Final Consolidation (Bango-Sub channel MT Model) #######################################
psales.sch.aggregate.final <- psales.sch.aggregate;droo.sch.aggregate.final <- droo.sch.aggregate
tts.sch.aggregate.final;bmi.sch.aggregate.final
sch.pkey <- c("Month","BRAND","Sub channel")
sch.consolidated.list1 <- merge.data.frame(psales.sch.aggregate.final,droo.sch.aggregate.final,by = sch.pkey)
sch.consolidated.list2 <- merge.data.frame(sch.consolidated.list1,y = tts.sch.aggregate.final,by=sch.pkey)
sch.consolidated.list3 <- merge.data.frame(sch.consolidated.list2,bmi.sch.aggregate.final,by=sch.pkey)

sch.final.data <- sch.consolidated.list3
sch.final.data$`Sales Organization` <- "9001";sch.final.data$Category <- "SAVOURY"
sch.final.data$Sector <- "SEASONINGS" ; sch.final.data$Channel <- "MODERN TRADE"
sch.final.data <- setnames(sch.final.data,old = "Sub channel",new="Customer")

#arrange the KPI's and rename the KPI's name as per the output
sch.final.data%<>%select(c(Month,`Sales Organization`,Category,Sector,BRAND,Channel,Customer,everything()))
sch.column.headers <- c(
  "Month",
  "Sales Organization",
  "Category",
  "Sector",
  "Brand",
  "Channel",
  "Customer",
  "Bango_GSV",
  "Bango_Sales Qty",
  "Bango_NIV",
  "Bango_Sales Qty (KG)",
  "Bango_Sales Qty(Tonns)",
  "Bango_OriginalOrder Qty.",
  "Bango_Final Customer Expected Order Qty.",
  "Bango_Dispatched Qty.",
  "Bango_DR",
  "Bango_DROO",
  "Bango_TTS",
  "Bango_BBT",
  "Bango_BBT - Place",
  "Bango_BBT - Place on invoice",
  "Bango_BBT - Place off invoice",
  "Bango_BBT - Price",
  "Bango_CPP on invoice",
  "Bango_CPP off invoice",
  "Bango_BBT - Product",
  "Bango_BBT - Pack",
  "Bango_BBT - Proposition",
  "Bango_BBT - Promotion",
  "Bango_EOT",
  "Bango_Brand & Marketing Investment",
  "Bango_Brand & Marketing Investment Trade",
  "Bango_Brand & Marketing Investment Consumer",
  "Bango_Promotional Expenses",
  "Bango_Promotion Packaging Material Cost",
  "Bango_Promotion Packaging Material Cost Trade",
  "Bango_Promotion Packaging Material Cost Consumer",
  "Bango_Promotion Repacking Cost",
  "Bango_Promotion Repacking Cost Trade",
  "Bango_Promotion Repacking Cost Consumer",
  "Bango_Promotion Communication Material Cost",
  "Bango_Promotion Communication Material Cost Trade",
  "Bango_Promotion Communication Material Cost Consumer",
  "Bango_Promo Samples, Gifts and Incentive Costs",
  "Bango_Promo Samples, Gifts and Incentive Costs Consumer",
  "Bango_Promotion Team Cost",
  "Bango_Promotion Team Cost Trade",
  "Bango_Promotion Team Cost Consumer",
  "Bango_Promo Agency Remun Fees & Commissions",
  "Bango_Promo Agency Remuneration Fees & Commissions Trade",
  "Bango_Promo Agency Remuneration Fees & Commissions Consumer"
)

colnames(sch.final.data) <- sch.column.headers 

#seggregate the data into HSM and Mini Sub Channel

HSM.final.data <- sch.final.data%>%
  filter(Customer==sub_channel.list[1])
HSM.final.data <- format(HSM.final.data,scientific=F)

writexl::write_xlsx(HSM.final.data,paste0(output.path,"/HSM.final.data.xlsx"))

Mini.final.data <- sch.final.data%>%
  filter(Customer==sub_channel.list[2])
Mini.final.data <- format(Mini.final.data,scientific=F)

writexl::write_xlsx(Mini.final.data,paste0(output.path,"/Mini.final.data.xlsx"))

###################################################################  ID Psales (Bango-Account MT Model) ####################################
#get the Account/region from the customer mapping file 
cust.account.map <- customer.mapping%>%
  select(c(`Customer Code`,`Account/Region`))
#get the conversion factor from the conversion mapping file
conv.account.map <- conversion.mapping%>%
  select(c(MATNR,Conversion))
#perform the mapping and merge the  input data frame with the {product mapping , customer mapping and conversion mapping }
psales.account.input <- merge.data.frame(x = id.psales.bango,y =product.mapping,by.x ="Material_Code",by.y = "Material Code")%>%
  merge.data.frame(x = .,y = cust.account.map,by.x = "Customer_Code_8D",by.y ="Customer Code")%>%
  merge.data.frame(x = .,y = conv.account.map,by.x = "Material_Code",by.y = "MATNR")
#filter out the observation for which the account/region is in the account list {MT Indomart , MT Alfamart , MT Carrefour}
mt.account.list <- c("MT Alfamart","MT Carrefour","MT Indomart")
psales.account.df <- psales.account.input%>%
  filter(`Account/Region` %in% mt.account.list)
#calculate the dervied KPI's from the given KPI's and then group and summarise the KPI's 

#rename the column and calculate the derived KPI's psales_kg and psales_tonns KPI's
psales.account.df <- setnames(psales.account.df,c("Material Desc.x"),c("Material Desc"))
psales.account.df%<>%mutate(psales_KG=`Sales Quantity`*Conversion,psales_Tonns=psales_KG/1000)

#perform the group by and summarise the KPI's
psales.account.aggregate <- psales.account.df%>%
  group_by(`Fiscal year/period`,BRAND,`Account/Region`)%>%
  summarise(`Gross Sales Value (GSV)`=sum(`Gross Sales Value (GSV)`),`Sales Quantity`=sum(`Sales Quantity`),NIV=sum(NIV),psales_KG=sum(psales_KG),psales_Tonns=sum(psales_Tonns))
#perform the month mapping for the Fiscal Period column 
psales.account.aggregate <- merge.data.frame(psales.account.aggregate,month.mapping,by.x = "Fiscal year/period",by.y ="Primary sales" )
psales.account.aggregate%<>%select(-c(`Secondary sales`:TTS_BMI,Sellout,`Fiscal year/period`))
psales.account.aggregate%<>%select(c(Month,everything()))
###################################################################  ID sellout  (Bango-Account MT Model) ####################################
sellout.account <- sellout.bango_account.aggregate%>%
  mutate(BRAND="BANGO")
sellout.account%<>%select(c(Month,BRAND,Customer,everything()))
sellout.account <- setnames(sellout.account,old = "Customer",new = "Account/Region")
##################################################################  ID DROO (Bango-Account MT Model) #############################################
droo.account.input <- merge.data.frame(x = id.droo.bango,y =product.mapping,by.x ="Material_Code",by.y = "Material Code")%>%
  merge.data.frame(x = .,y = cust.account.map,by.x = "STP_Code",by.y ="Customer Code")%>%
  merge.data.frame(x = .,y = conv.account.map,by.x = "Material_Code",by.y = "MATNR")

#filter out the sub channel which is HSM and Mini
droo.account.input%<>%filter(`Account/Region`%in% mt.account.list)
droo.account.df <- setnames(droo.account.input,c("Material Desc.x"),c("Material Desc"))
#summarise and aggregate the KPI's 
droo.account.aggregate <- droo.account.df%>%
  group_by(`Calendar Year/Month`,BRAND,`Account/Region`)%>%
  summarise(`OriginalOrder Qty`=sum(`OriginalOrder Qty`),`Final Customer Expected Order Qty`=sum(`Final Customer Expected Order Qty`),`Dispatched Qty`=sum(`Dispatched Qty`))

droo.account.aggregate%<>%mutate(DR=(`Dispatched Qty`/`Final Customer Expected Order Qty`*100),DROO=(`Dispatched Qty`/`OriginalOrder Qty`*100))
#perform the month-year mapping <month>-<year>
droo.account.aggregate <- merge.data.frame(droo.account.aggregate,month.mapping,by.x = "Calendar Year/Month",by.y ="Sec stock_DROO_Penetration" )
droo.account.aggregate%<>%select(-c(`Primary sales`,TTS_BMI,`Secondary sales`,`Calendar Year/Month`,Sellout))
droo.account.aggregate[which(is.nan(droo.account.aggregate$DR)),"DR"] <- 0
droo.account.aggregate%<>%select(c(Month,everything()))

######################################################### ID TTS Mapped (Bango-Account MT Model) ###############################################

tts.account.mapped.input <- tts.account.mapping
#filter out the observation for which Sub Channel is {HSM,Mini}
tts.account.mapped.input%<>%filter(`Account/Region`%in% mt.account.list)
#group-by and aggregate the KPI's based on the key {Month,Brand,Sub Channel}
tts.account.mapped.aggregate <- tts.account.mapped.input%>%
  group_by(`Fiscal year/period`,BRAND,`Account/Region`)%>%
  summarise(TTS=sum(TTS),BBT=sum(BBT),`BBT - Place`=sum(`BBT - Place`),`BBT - Place on invoice`=sum(`BBT - Place on invoice`),`BBT - Place off invoice`=sum(`BBT - Place off invoice`),`BBT - Price`=sum(`BBT - Price`),`CPP on invoice`=sum(`CPP on invoice`),`CPP off invoice`=sum(`CPP off invoice`),`BBT - Product`=sum(`BBT - Product`),`BBT - Pack`=sum(`BBT - Pack`),`BBT - Proposition`=sum(`BBT - Proposition`),`BBT - Promotion`=sum(`BBT - Promotion`),EOT=sum(EOT))

######################################################### ID TTS unMapped (Bango-Account MT Model) ###############################################
tts.account.unmapped.input <- tts.unmapped.mt.input
#filter out the observations for which the Account initiates with "MT" 
tts.account.unmapped.df <- tts.account.unmapped.input%>%
  unite(col = "key",c("Banner","Trade Format Level 2","Local Sales Force 2(m.d.)","Local Sales Force 3(m.d.)","Key Customer Level3"),sep="",remove = F)
tts.unmapped.account.mapping <- merge.data.frame(tts.account.unmapped.df,unmapped.tts.mapping,by.x = "key",by.y = "Key")
#filter out the observations for which the Account starts with "MT"
tts.unmapped.account.df <- tts.unmapped.account.mapping%>%
  filter(str_detect(`Account/Region`,"MT"))

tts.unmapped.account.df%<>%convert(num(TTS:EOT))
#group-by and aggregate the KPI's based on the key {Month,Brand,Sub Channel}
tts.unmapped.account.aggregate <- tts.unmapped.account.df%>%
  group_by(`Fiscal year/period`,BRAND,`Account/Region`)%>%
  summarise(TTS=sum(TTS),BBT=sum(BBT),`BBT - Place`=sum(`BBT - Place`),`BBT - Place on invoice`=sum(`BBT - Place on invoice`),`BBT - Place off invoice`=sum(`BBT - Place off invoice`),`BBT - Price`=sum(`BBT - Price`),`CPP on invoice`=sum(`CPP on invoice`),`CPP off invoice`=sum(`CPP off invoice`),`BBT - Product`=sum(`BBT - Product`),`BBT - Pack`=sum(`BBT - Pack`),`BBT - Proposition`=sum(`BBT - Proposition`),`BBT - Promotion`=sum(`BBT - Promotion`),EOT=sum(EOT))
############################################################## ID TTS consolidation (Bango-Account MT Model) ###############################################
tts.account.aggregate <- rbind(tts.account.mapped.aggregate,tts.unmapped.account.aggregate)
#group and summarise the KPI's 

tts.account.aggregate%<>%
  group_by(`Fiscal year/period`,BRAND,`Account/Region`)%>%
  summarise(TTS=sum(TTS),BBT=sum(BBT),`BBT - Place`=sum(`BBT - Place`),`BBT - Place on invoice`=sum(`BBT - Place on invoice`),`BBT - Place off invoice`=sum(`BBT - Place off invoice`),`BBT - Price`=sum(`BBT - Price`),`CPP on invoice`=sum(`CPP on invoice`),`CPP off invoice`=sum(`CPP off invoice`),`BBT - Product`=sum(`BBT - Product`),`BBT - Pack`=sum(`BBT - Pack`),`BBT - Proposition`=sum(`BBT - Proposition`),`BBT - Promotion`=sum(`BBT - Promotion`),EOT=sum(EOT))
#perform the Month - year mapping and get the period 
tts.account.aggregate.final <- merge.data.frame(x = tts.account.aggregate,y = month.mapping,by.x = "Fiscal year/period" , by.y = "TTS_BMI")
tts.account.aggregate.final%<>%select(-c(`Primary sales`:`Sec stock_DROO_Penetration`,Sellout,`Fiscal year/period`))
tts.account.aggregate.final%<>%select(c(Month,everything()))

############################################################# BMI Mapped (Bango-Account MT Model) ############################################
bmi.account.mapped.input <- bmi.mapped.bango
account.mapped.bmi.key<- banner.mapping[,c(2,5)]

bmi.account.mapped.input <- merge.data.frame(bmi.account.mapped.input,account.mapped.bmi.key,by.x = "Banner",by.y = "Banner code",all.x = T)

#filter out the observation for which the sub channel is {HSM.Mini}
bmi.account.mapped.input%<>%filter(str_detect(`Account/Region`,"MT"))
#perform the grouping and the aggregation 
bmi.account.mapped.aggregate <- bmi.account.mapped.input%>%
  group_by(`Fiscal year/period`,Brand,`Account/Region`)%>%
  summarise(X..Brand...Marketing.Investment=sum(X..Brand...Marketing.Investment),X..Brand...Marketing.Investment.Trade=sum(X..Brand...Marketing.Investment.Trade),X..Brand...Marketing.Investment.Consumer=sum(X..Brand...Marketing.Investment.Consumer),X..Promotional.Expenses=sum(X..Promotional.Expenses),X....Promotion.Packaging.Material.Cost=sum(X....Promotion.Packaging.Material.Cost),X......Promotion.Packaging.Material.Cost.Trade=sum(X......Promotion.Packaging.Material.Cost.Trade),X......Promotion.Packaging.Material.Cost.Consumer=sum(X......Promotion.Packaging.Material.Cost.Consumer),X....Promotion.Repacking.Cost=sum(X....Promotion.Repacking.Cost),X......Promotion.Repacking.Cost.Trade=sum(X....Promotion.Repacking.Cost),X......Promotion.Repacking.Cost.Trade=sum(X......Promotion.Repacking.Cost.Trade),X......Promotion.Repacking.Cost.Consumer=sum(X......Promotion.Repacking.Cost.Consumer),X....Promotion.Communication.Material.Cost=sum(X....Promotion.Communication.Material.Cost),X......Promotion.Communication.Material.Cost.Trade=sum(X......Promotion.Communication.Material.Cost.Trade),X......Promotion.Communication.Material.Cost.Consumer=sum(X......Promotion.Communication.Material.Cost.Consumer),X....Promo.Samples..Gifts.and.Incentive.Costs=sum(X....Promo.Samples..Gifts.and.Incentive.Costs),X......Promo.Samples..Gifts.and.Incentive.Costs.Consumer=sum(X......Promo.Samples..Gifts.and.Incentive.Costs.Consumer),X....Promotion.Team.Cost=sum(X....Promotion.Team.Cost),X......Promotion.Team.Cost.Trade=sum(X......Promotion.Team.Cost.Trade),X......Promotion.Team.Cost.Consumer=sum(X......Promotion.Team.Cost.Consumer),X....Promo.Agency.Remun.Fees...Commissions=sum(X....Promo.Agency.Remun.Fees...Commissions),X......Promo.Agency.Remuneration.Fees...Commissions.Trade=sum(X......Promo.Agency.Remuneration.Fees...Commissions.Trade),X......Promo.Agency.Remuneration.Fees...Commissions.Consumer=sum(X......Promo.Agency.Remuneration.Fees...Commissions.Consumer))

#rename the KPI's columns
bmi.account.mapped.aggregate <- bmi.account.mapped.aggregate%>%
  dplyr::rename(`Brand & Marketing Investment`=X..Brand...Marketing.Investment,`Brand & Marketing Investment Trade`=X..Brand...Marketing.Investment.Trade,`Brand & Marketing Investment Consumer`=X..Brand...Marketing.Investment.Consumer,`Promotional Expenses`=X..Promotional.Expenses,`Promotion Packaging Material Cost`=X....Promotion.Packaging.Material.Cost,`Promotion Communication Material Cost Trade`=X......Promotion.Communication.Material.Cost.Trade,`Promotion Packaging Material Cost Consumer`=X......Promotion.Packaging.Material.Cost.Consumer,`Promotion Repacking Cost`=X....Promotion.Repacking.Cost,`Promotion Repacking Cost Trade`=X......Promotion.Repacking.Cost.Trade,`Promotion Repacking Cost Consumer`=X......Promotion.Repacking.Cost.Consumer,`Promotion Communication Material Cost`=X....Promotion.Communication.Material.Cost,`Promotion Communication Material Cost Consumer`=X......Promotion.Communication.Material.Cost.Consumer,`Promo Samples, Gifts and Incentive Costs` =X....Promo.Samples..Gifts.and.Incentive.Costs,`Promo Samples, Gifts and Incentive Costs Consumer`=X......Promo.Samples..Gifts.and.Incentive.Costs.Consumer,`Promotion Team Cost`=X....Promotion.Team.Cost,`Promotion Team Cost Trade`=X......Promotion.Team.Cost.Trade,`Promotion Team Cost Consumer`=X......Promotion.Team.Cost.Consumer,`Promo Agency Remun Fees & Commissions`=X....Promo.Agency.Remun.Fees...Commissions,`Promo Agency Remuneration Fees & Commissions Trade`=X......Promo.Agency.Remuneration.Fees...Commissions.Trade,`Promo Agency Remuneration Fees & Commissions Consumer`=X......Promo.Agency.Remuneration.Fees...Commissions.Consumer)
bmi.account.mapped.aggregate%<>%dplyr::rename(`Promotion Packaging Material Cost Trade`=X......Promotion.Packaging.Material.Cost.Trade)

############################################################### BMI unmapped (bango-Account MT model) #############################################
bmi.account.unmapped.input <- id.bmi.unmapped.df1
#bmi unmapped data mapping for the key 
account.unmapped.bmi.key <- unmapped.bmi.mapping[,c(1,4)]
#merge the account/region based on the key 
bmi.unmapped.account.data.mapped <- merge.data.frame(bmi.account.unmapped.input,account.unmapped.bmi.key,by.x = "key",by.y = "Key")
#filter out the observation for which the sub channel is {HSM.Mini} and convert the KPI's from Char to num
bmi.unmapped.account.data.mapped%<>%filter(str_detect(`Account/Region`,"MT"))
bmi.unmapped.account.data.mapped%<>%convert(num(`Brand & Marketing Investment`:`Promo Agency Remuneration Fees & Commissions Consumer`))
#perfor the group by and summarise the KPI's for the BMI unmapped 
bmi.account.unmapped.aggregate <- bmi.unmapped.account.data.mapped%>%
  group_by(`Fiscal year/period`,Brand_Name,`Account/Region`)%>%
  summarise(`Brand & Marketing Investment`=sum(`Brand & Marketing Investment`),`Brand & Marketing Investment Trade`=sum(`Brand & Marketing Investment Trade`),`Brand & Marketing Investment Consumer`=sum(`Brand & Marketing Investment Consumer`),`Promotional Expenses`=sum(`Promotional Expenses`),`Promotion Packaging Material Cost`=sum(`Promotion Packaging Material Cost`),`Promotion Packaging Material Cost Trade`=sum(`Promotion Repacking Cost Trade`),`Promotion Packaging Material Cost Consumer`=sum(`Promotion Packaging Material Cost Consumer`),`Promotion Repacking Cost`=sum(`Promotion Repacking Cost`),`Promotion Repacking Cost Trade`=sum(`Promotion Repacking Cost Trade`),`Promotion Repacking Cost Consumer`=sum(`Promotion Repacking Cost Consumer`),`Promotion Communication Material Cost`=sum(`Promotion Communication Material Cost`),`Promotion Communication Material Cost Trade`=sum(`Promotion Communication Material Cost Trade`),`Promotion Communication Material Cost Consumer`=sum(`Promotion Communication Material Cost Consumer`),`Promo Samples, Gifts and Incentive Costs`=sum(`Promo Samples, Gifts and Incentive Costs`),`Promo Samples, Gifts and Incentive Costs Consumer`=sum(`Promo Samples, Gifts and Incentive Costs Consumer`),`Promotion Team Cost`=sum(`Promotion Team Cost`),`Promotion Team Cost Trade`=sum(`Promotion Team Cost Trade`),`Promotion Team Cost Consumer`=sum(`Promotion Team Cost Consumer`),`Promo Agency Remun Fees & Commissions`=sum(`Promo Agency Remun Fees & Commissions`),`Promo Agency Remuneration Fees & Commissions Trade`=sum(`Promo Agency Remuneration Fees & Commissions Trade`),`Promo Agency Remuneration Fees & Commissions Consumer`=sum(`Promo Agency Remuneration Fees & Commissions Consumer`))

bmi.account.unmapped.aggregate <- setnames(bmi.account.unmapped.aggregate,old = "Brand_Name",new = "Brand")

############################################################### BMI consolidation (bango-Account MT model) #############################################
bmi.account.aggregate <- rbind(bmi.account.mapped.aggregate,bmi.account.unmapped.aggregate)
#perform the group by and the aggregation of the KPi's
bmi.account.aggregate.final <- bmi.account.aggregate%>%
  group_by(`Fiscal year/period`,Brand,`Account/Region`)%>%
  summarise(`Brand & Marketing Investment`=sum(`Brand & Marketing Investment`),`Brand & Marketing Investment Trade`=sum(`Brand & Marketing Investment Trade`),`Brand & Marketing Investment Consumer`=sum(`Brand & Marketing Investment Consumer`),`Promotional Expenses`=sum(`Promotional Expenses`),`Promotion Packaging Material Cost`=sum(`Promotion Packaging Material Cost`),`Promotion Packaging Material Cost Trade`=sum(`Promotion Repacking Cost Trade`),`Promotion Packaging Material Cost Consumer`=sum(`Promotion Packaging Material Cost Consumer`),`Promotion Repacking Cost`=sum(`Promotion Repacking Cost`),`Promotion Repacking Cost Trade`=sum(`Promotion Repacking Cost Trade`),`Promotion Repacking Cost Consumer`=sum(`Promotion Repacking Cost Consumer`),`Promotion Communication Material Cost`=sum(`Promotion Communication Material Cost`),`Promotion Communication Material Cost Trade`=sum(`Promotion Communication Material Cost Trade`),`Promotion Communication Material Cost Consumer`=sum(`Promotion Communication Material Cost Consumer`),`Promo Samples, Gifts and Incentive Costs`=sum(`Promo Samples, Gifts and Incentive Costs`),`Promo Samples, Gifts and Incentive Costs Consumer`=sum(`Promo Samples, Gifts and Incentive Costs Consumer`),`Promotion Team Cost`=sum(`Promotion Team Cost`),`Promotion Team Cost Trade`=sum(`Promotion Team Cost Trade`),`Promotion Team Cost Consumer`=sum(`Promotion Team Cost Consumer`),`Promo Agency Remun Fees & Commissions`=sum(`Promo Agency Remun Fees & Commissions`),`Promo Agency Remuneration Fees & Commissions Trade`=sum(`Promo Agency Remuneration Fees & Commissions Trade`),`Promo Agency Remuneration Fees & Commissions Consumer`=sum(`Promo Agency Remuneration Fees & Commissions Consumer`))

#perform the month-year mapping by using the month mapping df
bmi.account.aggregate.final <- merge.data.frame(x = bmi.account.aggregate.final,y = month.mapping,by.x = "Fiscal year/period" , by.y = "TTS_BMI")
bmi.account.aggregate.final%<>%select(-c(`Primary sales`:`Sec stock_DROO_Penetration`,Sellout,`Fiscal year/period`))
bmi.account.aggregate.final%<>%select(c(Month,everything()))
bmi.account.aggregate.final$Brand <- toupper(bmi.account.aggregate.final$Brand)
bmi.account.aggregate.final <- setnames(bmi.account.aggregate.final,old = "Brand",new = "BRAND")

############################################################### Final COnsolidation( bango-Account MT model)########################################
psales.account.final <- psales.account.aggregate ; sellout.account.final <- sellout.account
droo.account.final <- droo.account.aggregate ; tts.account.final <- tts.account.aggregate.final
bmi.account.final <- bmi.account.aggregate.final

account.pkey <- c("Month","BRAND","Account/Region")

consolidated.account.list1 <- merge.data.frame(psales.account.final,droo.account.final,by = account.pkey)
consolidated.account.list2 <- merge.data.frame(consolidated.account.list1,tts.account.final,by = account.pkey,all.y = T)
consolidated.account.list3 <- merge.data.frame(consolidated.account.list2,bmi.account.final,by = account.pkey,all.y = T)
consolidated.account.list4 <- merge.data.frame(consolidated.account.list3,sellout.account.final,by = account.pkey,all.x = T)

account.final.data <- consolidated.account.list4

account.final.data$`Sales Organization` <- "9001";account.final.data$Category <- "SAVOURY"
account.final.data$Sector <- "SEASONINGS" ; account.final.data$Channel <- "MODERN TRADE"
account.final.data %<>% select(c(Month,`Sales Organization`,Category,Sector,BRAND,Channel,`Account/Region`,everything()))
account.final.data <- setnames(account.final.data,old = "Account/Region" ,new = "Customer")

#select the column names as per the output 

account.col.headers <- c("Month",
                         "Sales Organization",
                         "Category",
                         "Sector",
                         "Brand",
                         "Channel",
                         "Customer",
                         "Bango_GSV",
                         "Bango_Sales Qty",
                         "Bango_NIV",
                         "Bango_Sales Qty(Kg)",
                         "Bango_Sales Qty(Tonns)",
                         "Bango_OriginalOrder Qty.",
                         "Bango_Final Customer Expected Order Qty.",
                         "Bango_Dispatched Qty.",
                         "Bango_DR",
                         "Bango_DROO",
                         "Bango_TTS",
                         "Bango_BBT",
                         "Bango_BBT - Place",
                         "Bango_BBT - Place on invoice",
                         "Bango_BBT - Place off invoice",
                         "Bango_BBT - Price",
                         "Bango_CPP on invoice",
                         "Bango_CPP off invoice",
                         "Bango_BBT - Product",
                         "Bango_BBT - Pack",
                         "Bango_BBT - Proposition",
                         "Bango_BBT - Promotion",
                         "Bango_EOT",
                         "Bango_Brand & Marketing Investment",
                         "Bango_Brand & Marketing Investment Trade",
                         "Bango_Brand & Marketing Investment Consumer",
                         "Bango_Promotional Expenses",
                         "Bango_Promotion Packaging Material Cost",
                         "Bango_Promotion Packaging Material Cost Trade",
                         "Bango_Promotion Packaging Material Cost Consumer",
                         "Bango_Promotion Repacking Cost",
                         "Bango_Promotion Repacking Cost Trade",
                         "Bango_Promotion Repacking Cost Consumer",
                         "Bango_Promotion Communication Material Cost",
                         "Bango_Promotion Communication Material Cost Trade",
                         "Bango_Promotion Communication Material Cost Consumer",
                         "Bango_Promo Samples, Gifts and Incentive Costs",
                         "Bango_Promo Samples, Gifts and Incentive Costs Consumer",
                         "Bango_Promotion Team Cost",
                         "Bango_Promotion Team Cost Trade",
                         "Bango_Promotion Team Cost Consumer",
                         "Bango_Promo Agency Remun Fees & Commissions",
                         "Bango_Promo Agency Remuneration Fees & Commissions Trade",
                         "Bango_Promo Agency Remuneration Fees & Commissions Consumer",
                         "Bango_Sell out Value (IDR)",
                         "Bango_Unit Price",
                         "Bango_Sell out Volume (KGs)"
                         
                         
)
colnames(account.final.data) <- account.col.headers


#seggregate the Data into 3 channels and write the data into excel

account.alfamart.data <- account.final.data%>%
  filter(str_detect(Customer,"MT Alfa"))
data.frame(lapply(account.alfamart.data, gsub, pattern="NA", replacement=''))

account.indomart.data <- account.final.data%>%
  filter(str_detect(Customer,"MT Indo"))
data.frame(lapply(account.indomart.data, gsub, pattern="NA", replacement=''))


account.carrefour.data <- account.final.data%>%
  filter(str_detect(Customer,"MT Carrefour"))%>%
  select(-c(`Bango_Sell out Value (IDR)`:`Bango_Sell out Volume (KGs)`))
#account.carrefour.data <- format(account.carrefour.data,scientific=F)
data.frame(lapply(account.carrefour.data, gsub, pattern="NA", replacement=''))



writexl::write_xlsx(account.alfamart.data,paste0(output.path,"/MT.alfamart.xlsx"))
writexl::write_xlsx(account.indomart.data,paste0(output.path,"/MT.Indomart.xlsx"))
writexl::write_xlsx(account.carrefour.data,paste0(output.path,"/MT.Carrefour.xlsx"))



