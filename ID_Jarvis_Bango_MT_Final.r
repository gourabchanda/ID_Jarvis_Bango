start_time <- Sys.time()
working_directory <- getwd()
working_directory <- setwd("C:/Users/Gourab.Chanda/Desktop/R_Codes/ID_Jarvis_Bango_1.1")
source(paste0(working_directory,"/ID_Jarvis_Generalised_template1.0.0.r"))

############################################# ID Psales (Bango-Account MT Model) ###################################################
psales.mt.input <- merge.data.frame(x = id.psales.bango,y =psales.material.mt.mapping,by.x ="Material_Code",by.y = "Material Code")%>%
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
psales.mt.aggregate%<>%mutate(unit_price= NIV/psales_KG)
#perform the month mapping for the Fiscal Period column 
psales.mt.aggregate <- merge.data.frame(psales.mt.aggregate,month.mapping,by.x = "Fiscal year/period",by.y ="Primary sales" )
psales.mt.aggregate%<>%select(-c(`Sec stock_DROO_Penetration`,TTS_BMI,`Secondary sales`,`Fiscal year/period`,Sellout))
psales.mt.aggregate%<>%select(c(Month,`REQUIRED BASEPACK`,`Account/Region`,`Gross Sales Value (GSV)`:unit_price))
#rename the account/region column to Customer
psales.mt.aggregate <- setnames(psales.mt.aggregate,old ="Account/Region" ,new ="Customer")
############################################# ID DROO (Bango-Account MT Model) #####################################################

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
droo.mt.aggregate%<>%select(c(Month,everything()))
droo.mt.aggregate <- setnames(droo.mt.aggregate,old ="Account/Region" ,new ="Customer")
####################################################### Sell out Data Pre processing ( Bango-Account MT Model) ################################################################

sellout.mt.input.path <- "C:/Users/Gourab.Chanda/Desktop/Yashmin/Extracted_Data/Extracted_xlsx/Sellout"
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
#filter out the 9 basepacks and all the account for which it's "MT" and "All"
tts.mt.mapped <- tts.account.mapping%>%
  filter(`REQUIRED BASEPACK`!=0)
tts.mt.mapped %<>%filter(str_detect(`Account/Region`,"MT")|`Account/Region`=="ALL")
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
tts.mt.unmapped%<>%filter(str_detect(`Account/Region`,"MT")|`Account/Region`=="ALL")
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
  filter(str_detect(`Account/Region`,"MT")|`Account/Region`=="ALL")
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
bmi.unmapped.mt.data <- subset(x = bmi.unmapped.mt.data,subset = Mapping_Basepack=="Bango")

#bmi unmapped data mapping for the key 
unmapped.bmi.key <- unmapped.bmi.mapping[,c(1,4)]
#merge the account/region based on the key 
bmi.unmapped.mt.data.mapped <- merge.data.frame(bmi.unmapped.mt.data,unmapped.bmi.key,by.x = "key",by.y = "Key")
#filter out the accounts which is contain "MT" and "All"
bmi.unmapped.mt.data.mapped.filtered <- bmi.unmapped.mt.data.mapped%>%
  filter(str_detect(`Account/Region`,"MT")|`Account/Region`=="ALL")
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
  filter(str_detect(Customer,"MT Alfa")|Customer=="ALL")
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
                      "OriginalOrder Qty.",
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
  filter(str_detect(Customer,"MT Indo")|Customer=="ALL")
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
  filter(str_detect(Customer,"MT Carrefour")|Customer=="ALL")
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
                       "OriginalOrder Qty.",
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
write_xlsx(alfamart.bkm275,paste0(directory.path,"/alfamart.bkm275.xlsx"))

alfamart.bkm620 <- alfamart.final.list%>%
  filter(Product==basepack_list[2])
data.frame(lapply(alfamart.bkm620, gsub, pattern=NA, replacement=''))
write_xlsx(alfamart.bkm620,paste0(directory.path,"/alfamart.bkm620.xlsx"))


alfamart.bkm220 <- alfamart.final.list%>%
  filter(Product==basepack_list[3])
data.frame(lapply(alfamart.bkm220, gsub, pattern=NA, replacement=''))
write_xlsx(alfamart.bkm220,paste0(directory.path,"/alfamart.bkm220.xlsx"))



alfamart.bkmp220 <- alfamart.final.list%>%
  filter(Product==basepack_list[4])
data.frame(lapply(alfamart.bkmp220, gsub, pattern=NA, replacement=''))
write_xlsx(alfamart.bkmp220,paste0(directory.path,"/alfamart.bkmp220.xlsx"))

alfamart.bkmp135 <- alfamart.final.list%>%
  filter(Product==basepack_list[5])
data.frame(lapply(alfamart.bkmp135, gsub, pattern=NA, replacement=''))
write_xlsx(alfamart.bkmp135,paste0(directory.path,"/alfamart.bkmp135.xlsx"))

alfamart.bsofs60 <- alfamart.final.list%>%
  filter(Product==basepack_list[6])
data.frame(lapply(alfamart.bsofs60, gsub, pattern=NA, replacement=''))
write_xlsx(alfamart.bsofs60,paste0(directory.path,"/alfamart.bsofs60.xlsx"))

alfamart.bsofms30 <- alfamart.final.list%>%
  filter(Product==basepack_list[7])
data.frame(lapply(alfamart.bsofms30, gsub, pattern=NA, replacement=''))
write_xlsx(alfamart.bsofms30,paste0(directory.path,"/alfamart.bsofms30.xlsx"))

alfamart.bkm135 <- alfamart.final.list%>%
  filter(Product==basepack_list[8])
data.frame(lapply(alfamart.bkm135, gsub, pattern=NA, replacement=''))
write_xlsx(alfamart.bkm135,paste0(directory.path,"/alfamart.bkm135.xlsx"))

alfamart.bsof580 <- alfamart.final.list%>%
  filter(Product==basepack_list[9])
data.frame(lapply(alfamart.bsof580, gsub, pattern=NA, replacement=''))
write_xlsx(alfamart.bsof580,paste0(directory.path,"/alfamart.bsof580.xlsx"))

#indomart 
indomart.bkm275 <- indomart.final.list%>%
  filter(Product==basepack_list[1])
data.frame(lapply(indomart.bkm275, gsub, pattern=NA, replacement=''))
write_xlsx(indomart.bkm275,paste0(directory.path,"/indomart.bkm275.xlsx"))

indomart.bkm620 <- indomart.final.list%>%
  filter(Product==basepack_list[2])
data.frame(lapply(indomart.bkm620, gsub, pattern=NA, replacement=''))
write_xlsx(indomart.bkm620,paste0(directory.path,"/indomart.bkm620.xlsx"))


indomart.bkm220 <- indomart.final.list%>%
  filter(Product==basepack_list[3])
data.frame(lapply(indomart.bkm220, gsub, pattern=NA, replacement=''))
write_xlsx(indomart.bkm220,paste0(directory.path,"/indomart.bkm220.xlsx"))

indomart.bkmp220 <- indomart.final.list%>%
  filter(Product==basepack_list[4])
data.frame(lapply(indomart.bkmp220, gsub, pattern=NA, replacement=''))
write_xlsx(indomart.bkmp220,paste0(directory.path,"/indomart.bkmp220.xlsx"))


indomart.bkmp135 <- indomart.final.list%>%
  filter(Product==basepack_list[5])
data.frame(lapply(indomart.bkmp135, gsub, pattern=NA, replacement=''))
write_xlsx(indomart.bkmp135,paste0(directory.path,"/indomart.bkmp135.xlsx"))

indomart.bsofs60 <- indomart.final.list%>%
  filter(Product==basepack_list[6])
data.frame(lapply(indomart.bsofs60, gsub, pattern=NA, replacement=''))
write_xlsx(indomart.bsofs60,paste0(directory.path,"/indomart.bsofs60.xlsx"))

indomart.bsofms30 <- indomart.final.list%>%
  filter(Product==basepack_list[7])
data.frame(lapply(indomart.bsofms30, gsub, pattern=NA, replacement=''))
write_xlsx(indomart.bsofms30,paste0(directory.path,"/indomart.bsofms30.xlsx"))

indomart.bkm135 <- indomart.final.list%>%
  filter(Product==basepack_list[8])
data.frame(lapply(indomart.bkm135, gsub, pattern=NA, replacement=''))
write_xlsx(indomart.bkm135,paste0(directory.path,"/indomart.bkm135.xlsx"))


indomart.bsof580 <- indomart.final.list%>%
  filter(Product==basepack_list[9])
data.frame(lapply(indomart.bsof580, gsub, pattern=NA, replacement=''))
write_xlsx(indomart.bsof580,paste0(directory.path,"/indomart.bsof580.xlsx"))

#carrefour
Carrefour.bkm275 <- Carrefour.final.list%>%
  filter(Product==basepack_list[1])
data.frame(lapply(Carrefour.bkm275, gsub, pattern=NA, replacement=''))
write_xlsx(Carrefour.bkm275,paste0(directory.path,"/Carrefour.bkm275.xlsx"))


Carrefour.bkm620 <- Carrefour.final.list%>%
  filter(Product==basepack_list[2])
data.frame(lapply(Carrefour.bkm620, gsub, pattern=NA, replacement=''))
write_xlsx(Carrefour.bkm620,paste0(directory.path,"/Carrefour.bkm620.xlsx"))


Carrefour.bkm220 <- Carrefour.final.list%>%
  filter(Product==basepack_list[3])
data.frame(lapply(Carrefour.bkm220, gsub, pattern=NA, replacement=''))
write_xlsx(Carrefour.bkm220,paste0(directory.path,"/Carrefour.bkm220.xlsx"))


Carrefour.bkmp220 <- Carrefour.final.list%>%
  filter(Product==basepack_list[4])
data.frame(lapply(Carrefour.bkmp220, gsub, pattern=NA, replacement=''))
write_xlsx(Carrefour.bkmp220,paste0(directory.path,"/Carrefour.bkmp220.xlsx"))


Carrefour.bkmp135 <- Carrefour.final.list%>%
  filter(Product==basepack_list[5])
data.frame(lapply(Carrefour.bkmp135, gsub, pattern=NA, replacement=''))
write_xlsx(Carrefour.bkmp135,paste0(directory.path,"/Carrefour.bkmp135.xlsx"))


Carrefour.bsofs60 <- Carrefour.final.list%>%
  filter(Product==basepack_list[6])
data.frame(lapply(Carrefour.bsofs60, gsub, pattern=NA, replacement=''))
write_xlsx(Carrefour.bsofs60,paste0(directory.path,"/Carrefour.bsofs60.xlsx"))


Carrefour.bsofms30 <- Carrefour.final.list%>%
  filter(Product==basepack_list[7])
data.frame(lapply(Carrefour.bsofms30, gsub, pattern=NA, replacement=''))
write_xlsx(Carrefour.bsofms30,paste0(directory.path,"/Carrefour.bsofms30.xlsx"))


Carrefour.bkm135 <- Carrefour.final.list%>%
  filter(Product==basepack_list[8])
data.frame(lapply(Carrefour.bkm135, gsub, pattern=NA, replacement=''))
write_xlsx(Carrefour.bkm135,paste0(directory.path,"/Carrefour.bkm135.xlsx"))


Carrefour.bsof580 <- Carrefour.final.list%>%
  filter(Product==basepack_list[9])
data.frame(lapply(Carrefour.bsof580, gsub, pattern=NA, replacement=''))
write_xlsx(Carrefour.bsof580,paste0(directory.path,"/Carrefour.bsof580.xlsx"))

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
  "Sector" ,
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
  "Bango_Promo Agency Remuneration Fees & Commissions Consumer"
)

colnames(sch.final.data) <- sch.column.headers 

#seggregate the data into HSM and Mini Sub Channel

HSM.final.data <- sch.final.data%>%
  filter(Customer==sub_channel.list[1])
HSM.final.data <- format(HSM.final.data,scientific=F)

writexl::write_xlsx(HSM.final.data,paste0(directory.path,"/HSM.final.data.xlsx"))

Mini.final.data <- sch.final.data%>%
  filter(Customer==sub_channel.list[2])
Mini.final.data <- format(Mini.final.data,scientific=F)

writexl::write_xlsx(Mini.final.data,paste0(directory.path,"/Mini.final.data.xlsx"))

################################################################### 


