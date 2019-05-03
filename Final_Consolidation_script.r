############################################ Final COnsolidation Script ##########################################################
options("scipen"=100, "digits"=4)
start_time <- Sys.time()
working_directory <- setwd("C:/Users/Gourab.Chanda/Desktop/R_Codes/ID_Jarvis_Final")
#source(paste0(working_directory,"/data_preprocessing_script.r"))
#source(paste0(working_directory,"/ID_MT_Final.r"))
source(paste0(working_directory,"/Data_Blending_script.r"))

psales_file_name <- psales.files
psales_file_name <- str_remove(psales_file_name,".xlsx")
x <- psales_file_name%>%
  str_match_all("[0-9]+") %>% unlist %>% as.numeric
file_name_split <- unlist(str_split(psales_file_name," "))
file_latest_time <- file_name_split[length(file_name_split)]
time_abb <- lubridate::dmy(paste("01-",file_latest_time,sep=""))
time_abb <- format(time_abb,format="%b %y")


output_file_name <- function(output_file_name,time=time_abb){
  file_name_split <- unlist(str_split(output_file_name,pattern = "_"))
  file_name <- paste0(file_name_split[1],"_",time,".xlsx")
  return(file_name)
}

############################################# Bango-GT Model ( 1 output file)  ######################################################################
bango_gt_output_path <- "C:/Users/Gourab.Chanda/OneDrive - Unilever/ID_Jarvis_Bango-HICW10NA512SG4H/Final_Files/Bango-GT"
bango_gt_output_files <- list.files(path = bango_gt_output_path ,pattern = "^GD ID GT -Bango(.*)xlsx|XLSX$")
bango_gt_output_file.df <- read_xlsx(paste0(bango_gt_output_path,"/",bango_gt_output_files),sheet = 1,guess_max = 100)
total.bango.gt.final%<>%convert(num(`Bango GT_GSV`:`Bango_I Trust %`))
#bango_gt_output_file.df%<>%convert(num(`Bango GT_GSV`:`Indofood_Have a more attractive packaging than other brands`))
bango.gt.df <- dplyr::bind_rows(bango_gt_output_file.df,total.bango.gt.final)
bango.gt.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.gt.df%<>%arrange(Month_ymd)
bango.gt.df <- bango.gt.df[order(bango.gt.df$Channel,decreasing=T),]
bango.gt.df %<>% mutate_if(is.numeric, as.character)
#bango.gt.df <- as.character(bango.gt.df)
#bango.gt.df <- data.frame(lapply(bango.gt.df, gsub, pattern='NA', replacement=''))
gt_fname <- output_file_name(bango_gt_output_files)
write_xlsx(bango.gt.df,paste0(bango_gt_output_path,"/",gt_fname))
############################################# Bango - MT Model ( 1 output file)  ###################################################################
bango_mt_output_path <- "C:/Users/Gourab.Chanda/OneDrive - Unilever/ID_Jarvis_Bango-HICW10NA512SG4H/Final_Files/Bango-MT"
bango_mt_output_files <- list.files(path = bango_mt_output_path ,pattern = "^GD ID MT -Bango(.*)xlsx|XLSX$")
bango_mt_output_file.df <- read_xlsx(paste0(bango_mt_output_path,"/",bango_mt_output_files),sheet = 1,guess_max = 100)
total.bango.mt.final%<>%convert(num(Bango_GSV:`Bango_Promo Agency Remuneration Fees & Commissions Trade`))
bango_mt_output_file.df%<>%convert(num(Bango_GSV:Indofood_TBCA))
bango.mt.df <- dplyr::bind_rows(bango_mt_output_file.df,total.bango.mt.final)
bango.mt.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.mt.df%<>%arrange(Month_ymd)
bango.mt.df <- bango.mt.df[order(bango.mt.df$Channel,decreasing=T),]
bango.mt.df%<>%mutate_if(is.numeric, as.character)
#bango.mt.df <- format.data.frame(bango.mt.df,scientific=F)
#bango.mt.df <- data.frame(lapply(bango.mt.df, gsub, pattern='NA', replacement=''))
mt_fname <- output_file_name(bango_mt_output_files)
write_xlsx(bango.mt.df,paste0(bango_mt_output_path,"/",mt_fname))
############################################ Bango - National Model ( 1 output file) ###############################################################
bango_national_output_path <- "C:/Users/Gourab.Chanda/OneDrive - Unilever/ID_Jarvis_Bango-HICW10NA512SG4H/Final_Files/Bango-National"
bango_national_output_files <- list.files(path = bango_national_output_path ,pattern = "^GD ID Indonesia-Bango(.*)xlsx|XLSX$")
bango_national_output_file.df <- read_xlsx(paste0(bango_national_output_path,"/",bango_national_output_files),sheet = 1,guess_max = 100)
bango_national_output_file.df%<>%convert(num(Bango_GSV:`Indofood_Have a more attractive packaging than other brands`))
total.final.bango%<>%convert(num(Bango_GSV:`Bango_Penetration%`))

bango.national.df <- dplyr::bind_rows(bango_national_output_file.df,total.final.bango)
bango.national.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.national.df%<>%arrange(Month_ymd)
bango.national.df <- bango.national.df[order(bango.national.df$Market),]
#bango.national.df <- format.data.frame(bango.national.df,scientific=F)
#bango.national.df <- data.frame(lapply(bango.national.df, gsub, pattern='NA', replacement=''))
national_fname <- output_file_name(bango_national_output_files)
bango.national.df%<>%mutate_if(is.numeric, as.character)

write_xlsx(bango.national.df,paste0(bango_national_output_path,"/",national_fname))
########################################### Bango - Region ( 4 output files) #########################################################################

bango_region_output_path <- "C:/Users/Gourab.Chanda/OneDrive - Unilever/ID_Jarvis_Bango-HICW10NA512SG4H/Final_Files/Bango-Region"
bango_region_output_files <- list.files(path=bango_region_output_path,pattern="^GD ID GT(.*)xlsx|XLSX$")
bango_cej_fname <- output_file_name(bango_region_output_files[1])
bango_oi_fname <- output_file_name(bango_region_output_files[2])
bango_sumatera_fname <- output_file_name(bango_region_output_files[3])
bango_wj_fname <- output_file_name(bango_region_output_files[4])


bango_region_cej_output <-  read_xlsx(paste0(bango_region_output_path,"/",bango_region_output_files[1]),sheet = 1,guess_max = 100)
bango_region_oi_output <-   read_xlsx(paste0(bango_region_output_path,"/",bango_region_output_files[2]),sheet = 1,guess_max = 100)
bango_region_sumatra_output <- read_xlsx(paste0(bango_region_output_path,"/",bango_region_output_files[3]),sheet = 1,guess_max = 100)
bango_region_wj_output <- read_xlsx(paste0(bango_region_output_path,"/",bango_region_output_files[4]),sheet = 1,guess_max = 100)

bango.region.cej%<>%convert(num(Bango_GSV:`Bango_I trust %`))
bango.region.oi%<>%convert(num(Bango_GSV:`Bango_I trust %`))
bango.region.sumatra%<>%convert(num(Bango_GSV:`Bango_I trust %`))
bango.region.wj%<>%convert(num(Bango_GSV:`Bango_I trust %`))


bango.cej.df <- dplyr::bind_rows(bango_region_cej_output,bango.region.cej)
bango.cej.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.cej.df%<>%arrange(Month_ymd)
bango.cej.df <- bango.cej.df[order(bango.cej.df$Customer,decreasing=T),]
bango.cej.df%<>%mutate_if(is.numeric, as.character)


bango.oi.df <- dplyr::bind_rows(bango_region_oi_output,bango.region.oi)
bango.oi.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.oi.df%<>%arrange(Month_ymd)
bango.oi.df <- bango.oi.df[order(bango.oi.df$Customer,decreasing=T),]
bango.oi.df%<>%mutate_if(is.numeric, as.character)

bango.sumatra.df <- dplyr::bind_rows(bango_region_sumatra_output,bango.region.sumatra)
bango.sumatra.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.sumatra.df%<>%arrange(Month_ymd)
bango.sumatra.df <- bango.sumatra.df[order(bango.sumatra.df$Customer,decreasing=T),]
bango.sumatra.df%<>%mutate_if(is.numeric, as.character)

bango.wj.df <- dplyr::bind_rows(bango_region_wj_output,bango.region.wj)
bango.wj.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.wj.df%<>%arrange(Month_ymd)
bango.wj.df <- bango.wj.df[order(bango.wj.df$Customer,decreasing=T),]
bango.wj.df%<>%mutate_if(is.numeric, as.character)


write_xlsx(bango.cej.df,paste0(bango_region_output_path,"/",bango_cej_fname))
write_xlsx(bango.oi.df,paste0(bango_region_output_path,"/",bango_oi_fname))
write_xlsx(bango.sumatra.df,paste0(bango_region_output_path,"/",bango_sumatera_fname))
write_xlsx(bango.wj.df,paste0(bango_region_output_path,"/",bango_wj_fname))



############################################################### Bango - Sub Channel (2 output files) #########################################

bango_subchannel_output_path <- "C:/Users/Gourab.Chanda/OneDrive - Unilever/ID_Jarvis_Bango-HICW10NA512SG4H/Final_Files/Bango-Sub channel"

bango_subchannel_output_files <- list.files(path=bango_subchannel_output_path,pattern="^GD ID MT(.*)xlsx|XLSX$")

bango_hsm_fname <- output_file_name(bango_subchannel_output_files[1])
bango_mini_fname <- output_file_name(bango_subchannel_output_files[2])   

bango_subchannel_hsm_output <-  read_xlsx(paste0(bango_subchannel_output_path,"/",bango_subchannel_output_files[1]),sheet = 1,guess_max = 100)
bango_subchannel_mini_output <- read_xlsx(paste0(bango_subchannel_output_path,"/",bango_subchannel_output_files[2]),sheet = 1,guess_max = 100)


HSM.final.data%<>%convert(num(Bango_GSV:`Bango_Promo Agency Remuneration Fees & Commissions Consumer`))
Mini.final.data%<>%convert(num(Bango_GSV:`Bango_Promo Agency Remuneration Fees & Commissions Consumer`))


bango.hsm.df <- dplyr::bind_rows(bango_subchannel_hsm_output,HSM.final.data)
bango.hsm.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.hsm.df%<>%arrange(Month_ymd)
bango.hsm.df <- bango.hsm.df[order(bango.hsm.df$Customer,decreasing=T),]
bango.hsm.df%<>%mutate_if(is.numeric, as.character)

bango.mini.df <- dplyr::bind_rows(bango_subchannel_mini_output,Mini.final.data)
bango.mini.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.mini.df%<>%arrange(Month_ymd)
bango.mini.df <- bango.mini.df[order(bango.mini.df$Customer,decreasing=T),]
bango.mini.df%<>%mutate_if(is.numeric, as.character)

#bango.hsm.df <- format.data.frame(bango.hsm.df,scientific=F)
#bango.mini.df <- format.data.frame(bango.mini.df,scientific=F)

write_xlsx(bango.hsm.df,paste0(bango_subchannel_output_path,"/",bango_hsm_fname))
write_xlsx(bango.mini.df,paste0(bango_subchannel_output_path,"/",bango_mini_fname))


############################################################### Bango - Account(3 output files) #########################################

bango_account_output_path <- "C:/Users/Gourab.Chanda/OneDrive - Unilever/ID_Jarvis_Bango-HICW10NA512SG4H/Final_Files/Bango-Account"
bango_account_output_files <- list.files(path=bango_account_output_path,pattern = "^GD ID MT(.*).xlsx|XLSX")
bango_account_alfamart_output <- read_xlsx(paste0(bango_account_output_path,"/",bango_account_output_files[1]),sheet = 1,guess_max = 100)
bango_account_careffour_output <- read_xlsx(paste0(bango_account_output_path,"/",bango_account_output_files[2]),sheet = 1,guess_max = 100)
bango_account_indomart_output <- read_xlsx(paste0(bango_account_output_path,"/",bango_account_output_files[3]),sheet = 1,guess_max = 100)

alfamart_fname <- output_file_name(bango_account_output_files[1])
carrefour_fname <- output_file_name(bango_account_output_files[2])
Indomart_fname <- output_file_name(bango_account_output_files[3])


account.alfamart.data%<>%convert(num(Bango_GSV:`Bango_Sell out Volume (KGs)`))
account.carrefour.data%<>%convert(num(Bango_GSV:`Bango_Promo Agency Remuneration Fees & Commissions Consumer`))
account.indomart.data%<>%convert(num(Bango_GSV:`Bango_Sell out Volume (KGs)`))


bango.alfamart.df <- dplyr::bind_rows(bango_account_alfamart_output,account.alfamart.data)
bango.alfamart.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.alfamart.df%<>%arrange(Month_ymd)
bango.alfamart.df <- bango.alfamart.df[order(bango.alfamart.df$Customer,decreasing=T),]
bango.alfamart.df%<>%mutate_if(is.numeric, as.character)

bango_account_careffour_output%<>%mutate_if(is.logical, as.character)
bango_account_careffour_output[,9:ncol(bango_account_careffour_output)-1] <- lapply(bango_account_careffour_output[,9:ncol(bango_account_careffour_output)-1], as.numeric)
bango.carrefour.df <- dplyr::bind_rows(bango_account_careffour_output,account.carrefour.data)
bango.carrefour.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.carrefour.df%<>%arrange(Month_ymd)
bango.carrefour.df <- bango.carrefour.df[order(bango.carrefour.df$Customer,decreasing=T),]
bango.carrefour.df%<>%mutate_if(is.numeric, as.character)

bango.indomart.df <- dplyr::bind_rows(bango_account_indomart_output,account.indomart.data)
bango.indomart.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.indomart.df%<>%arrange(Month_ymd)
bango.indomart.df <- bango.indomart.df[order(bango.indomart.df$Customer,decreasing=T),]
bango.indomart.df%<>%mutate_if(is.numeric, as.character)

#bango.alfamart.df <- format.data.frame(bango.alfamart.df,scientific=F)
#bango.carrefour.df <- format.data.frame(bango.carrefour.df,scientific=F)
#bango.indomart.df <- format.data.frame(bango.indomart.df,scientific=F)

write_xlsx(bango.alfamart.df,paste0(bango_account_output_path,"/",alfamart_fname))
write_xlsx(bango.carrefour.df,paste0(bango_account_output_path,"/",carrefour_fname))
write_xlsx(bango.indomart.df,paste0(bango_account_output_path,"/",Indomart_fname))

################################################################## Packgroup - Region ( 24 output files - 4*6 =24 ) ###########################################################

packgroup_region_dir_path <-  "C:/Users/Gourab.Chanda/OneDrive - Unilever/ID_Jarvis_Bango-HICW10NA512SG4H/Final_Files/Pack group- Region"

packgroup_region_files <- list.files(path = packgroup_region_dir_path,recursive = T)

bmbp_cej_fname <- output_file_name(packgroup_region_files[1])
bmbp_oi_fname <- output_file_name(packgroup_region_files[2])
bmbp_sumatera_fname <- output_file_name(packgroup_region_files[3])
bmbp_wj_fname <- output_file_name(packgroup_region_files[4])

bmbs_cej_fname <- output_file_name(packgroup_region_files[5])
bmbs_oi_fname <- output_file_name(packgroup_region_files[6])
bmbs_sumatera_fname <- output_file_name(packgroup_region_files[7])
bmbs_wj_fname <- output_file_name(packgroup_region_files[8])

bmmb_cej_fname <- output_file_name(packgroup_region_files[9])
bmmb_oi_fname <- output_file_name(packgroup_region_files[10])
bmmb_sumatera_fname <- output_file_name(packgroup_region_files[11])
bmmb_wj_fname <-output_file_name(packgroup_region_files[12])

bmmp_cej_fname <- output_file_name(packgroup_region_files[13])
bmmp_oi_fname <- output_file_name(packgroup_region_files[14])
bmmp_sumatera_fname <- output_file_name(packgroup_region_files[15])
bmmp_wj_fname <- output_file_name(packgroup_region_files[16])




bmsb_cej_fname <- output_file_name(packgroup_region_files[17])
bmsb_oi_fname <- output_file_name(packgroup_region_files[18])
bmsb_sumatera_fname <- output_file_name(packgroup_region_files[19])
bmsb_wj_fname <- output_file_name(packgroup_region_files[20])


bmsp_cej_fname <- output_file_name(packgroup_region_files[21])
bmsp_oi_fname <- output_file_name(packgroup_region_files[22])
bmsp_sumatera_fname <- output_file_name(packgroup_region_files[23])
bmsp_wj_fname <- output_file_name(packgroup_region_files[24])



packgroup_bmbp_cej_output <- read_xlsx(paste0(packgroup_region_dir_path,"/",packgroup_region_files[1]),sheet = 1,guess_max = 100)
packgroup_bmbp_oi_output <- read_xlsx(paste0(packgroup_region_dir_path,"/",packgroup_region_files[2]),sheet = 1,guess_max = 100)
packgroup_bmbp_sumatera_output <- read_xlsx(paste0(packgroup_region_dir_path,"/",packgroup_region_files[3]),sheet = 1,guess_max = 100)
packgroup_bmbp_wj_output <- read_xlsx(paste0(packgroup_region_dir_path,"/",packgroup_region_files[4]),sheet = 1,guess_max = 100)

packgroup_bmbs_cej_output <- read_xlsx(paste0(packgroup_region_dir_path,"/",packgroup_region_files[5]),sheet = 1,guess_max = 100)
packgroup_bmbs_oi_output <- read_xlsx(paste0(packgroup_region_dir_path,"/",packgroup_region_files[6]),sheet = 1,guess_max = 100)
packgroup_bmbs_sumatera_output <- read_xlsx(paste0(packgroup_region_dir_path,"/",packgroup_region_files[7]),sheet = 1,guess_max = 100)
packgroup_bmbs_wj_output <- read_xlsx(paste0(packgroup_region_dir_path,"/",packgroup_region_files[8]),sheet = 1,guess_max = 100)

packgroup_bmmb_cej_output <- read_xlsx(paste0(packgroup_region_dir_path,"/",packgroup_region_files[9]),sheet = 1,guess_max = 100)
packgroup_bmmb_oi_output <- read_xlsx(paste0(packgroup_region_dir_path,"/",packgroup_region_files[10]),sheet = 1,guess_max = 100)
packgroup_bmmb_sumatera_output <- read_xlsx(paste0(packgroup_region_dir_path,"/",packgroup_region_files[11]),sheet = 1,guess_max = 100)
packgroup_bmmb_wj_output <- read_xlsx(paste0(packgroup_region_dir_path,"/",packgroup_region_files[12]),sheet = 1,guess_max = 100)

packgroup_bmmp_cej_output <- read_xlsx(paste0(packgroup_region_dir_path,"/",packgroup_region_files[13]),sheet = 1,guess_max = 100)
packgroup_bmmp_oi_output <- read_xlsx(paste0(packgroup_region_dir_path,"/",packgroup_region_files[14]),sheet = 1,guess_max = 100)
packgroup_bmmp_sumatera_output <- read_xlsx(paste0(packgroup_region_dir_path,"/",packgroup_region_files[15]),sheet = 1,guess_max = 100)
packgroup_bmmp_wj_output <- read_xlsx(paste0(packgroup_region_dir_path,"/",packgroup_region_files[16]),sheet = 1,guess_max = 100)

packgroup_bmsb_cej_output <- read_xlsx(paste0(packgroup_region_dir_path,"/",packgroup_region_files[17]),sheet = 1,guess_max = 100)
packgroup_bmsb_oi_output <- read_xlsx(paste0(packgroup_region_dir_path,"/",packgroup_region_files[18]),sheet = 1,guess_max = 100)
packgroup_bmsb_sumatera_output <- read_xlsx(paste0(packgroup_region_dir_path,"/",packgroup_region_files[19]),sheet = 1,guess_max = 100)
packgroup_bmsb_wj_output <- read_xlsx(paste0(packgroup_region_dir_path,"/",packgroup_region_files[20]),sheet = 1,guess_max = 100)

packgroup_bmsp_cej_output <- read_xlsx(paste0(packgroup_region_dir_path,"/",packgroup_region_files[21]),sheet = 1,guess_max = 100)
packgroup_bmsp_oi_output <- read_xlsx(paste0(packgroup_region_dir_path,"/",packgroup_region_files[22]),sheet = 1,guess_max = 100)
packgroup_bmsp_sumatera_output <- read_xlsx(paste0(packgroup_region_dir_path,"/",packgroup_region_files[23]),sheet = 1,guess_max = 100)
packgroup_bmsp_wj_output <- read_xlsx(paste0(packgroup_region_dir_path,"/",packgroup_region_files[24]),sheet = 1,guess_max = 100)


#final consolidation for the Bango Manis Big Pouch : packgroup level
packgroup.bmbp.cej%<>%convert(num(`Bango Manis Big Pouch_GSV`:`Bango Manis Big Pouch_Internal_penetration%`))
packgroup.bmbp.cej.df <- dplyr::bind_rows(packgroup_bmbp_cej_output,packgroup.bmbp.cej)
packgroup.bmbp.cej.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmbp.cej.df%<>%arrange(Month_ymd)
packgroup.bmbp.cej.df <- packgroup.bmbp.cej.df[order(packgroup.bmbp.cej.df$Customer,decreasing=T),]
packgroup.bmbp.cej.df%<>%mutate_if(is.numeric, as.character)


packgroup.bmbp.oi%<>%convert(num(`Bango Manis Big Pouch_GSV`:`Bango Manis Big Pouch_Internal_penetration%`))
packgroup.bmbp.oi.df <- dplyr::bind_rows(packgroup_bmbp_oi_output,packgroup.bmbp.oi)
packgroup.bmbp.oi.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmbp.oi.df%<>%arrange(Month_ymd)
packgroup.bmbp.oi.df <- packgroup.bmbp.oi.df[order(packgroup.bmbp.oi.df$Customer,decreasing=T),]
packgroup.bmbp.oi.df%<>%mutate_if(is.numeric, as.character)

packgroup.bmbp.sumatera%<>%convert(num(`Bango Manis Big Pouch_GSV`:`Bango Manis Big Pouch_Internal_penetration%`))
packgroup.bmbp.sumatera.df <- dplyr::bind_rows(packgroup_bmbp_sumatera_output,packgroup.bmbp.sumatera)
packgroup.bmbp.sumatera.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmbp.sumatera.df%<>%arrange(Month_ymd)
packgroup.bmbp.sumatera.df <- packgroup.bmbp.sumatera.df[order(packgroup.bmbp.sumatera.df$Customer,decreasing=T),]
packgroup.bmbp.sumatera.df%<>%mutate_if(is.numeric, as.character)


packgroup.bmbp.wj%<>%convert(num(`Bango Manis Big Pouch_GSV`:`Bango Manis Big Pouch_Internal_penetration%`))
packgroup_bmbp_wj_headers <- c("Month",
                               "Sales Organization",
                               "Category",
                               "Sector",
                               "Brand",
                               "Product",
                               "Channel",
                               "Customer",
                               "Bango Big Pouch_GSV",
                               "Bango Big Pouch_Primary Sales Qty(PC)",
                               "Bango Big Pouch_NIV",
                               "Bango Big Pouch_Primary Sales Qty(Kg)",
                               "Bango Big Pouch_Primary Sales Qty(Tonns)",
                               "Bango Big Pouch_Sec Sales Value",
                               "Bango Big Pouch_Sec Sales Volume(PC)",
                               "Bango Big Pouch_Sec sales Volume (Kg)",
                               "Bango Big Pouch_Sec sales Volume (Tonns)",
                               "Bango Big Pouch_Secondary Stock Value [@ DT Rate.]",
                               "Bango Big Pouch_Secondary Stock Volume(PC)",
                               "Bango Big Pouch_Secondary stock volume(Kg)",
                               "Bango Big Pouch_Secondary stock volume(Tonns)",
                               "Bango Big Pouch_Original Order Qty.(PC)",
                               "Bango Big Pouch_Final Customer Expected Order Qty.(PC)",
                               "Bango Big Pouch_Dispatched Qty.(PC)",
                               "Bango Big Pouch_DROO",
                               "Bango Big Pouch_DR",
                               "Bango Big Pouch_I trust Line",
                               "Bango Big Pouch_I trust Total",
                               "Bango Big Pouch_I trust %",
                               "Bango Big Pouch_TTS",
                               "Bango Big Pouch_BBT",
                               "Bango Big Pouch_BBT-Place",
                               "Bango Big Pouch_BBT-Place on invoice",
                               "Bango Big Pouch_BBT-Place off invoice",
                               "Bango Big Pouch_BBT-Price",
                               "Bango Big Pouch_CPP on invoice",
                               "Bango Big Pouch_CPP off invoice",
                               "Bango Big Pouch_BBT-Product",
                               "Bango Big Pouch_BBT-Pack",
                               "Bango Big Pouch_BBT-Proposition",
                               "Bango Big Pouch_BBT-Promotion",
                               "Bango Big Pouch_EOT",
                               "Bango Big Pouch_Brand & Marketing Investment",
                               "Bango Big Pouch_Brand & Marketing Investment Trade",
                               "Bango Big Pouch_Brand & Marketing Investment Consumer",
                               "Bango Big Pouch_Promotion Team Cost Trade",
                               "Bango Big Pouch_Promotion Team Cost Consumer",
                               "Bango Big Pouch_Promotion Team Cost",
                               "Bango Big Pouch_Promotion Packaging Material Cost Trade",
                               "Bango Big Pouch_Promotion Packaging Material Cost Consumer",
                               "Bango Big Pouch_Promotion Communication Material Cost Trade",
                               "Bango Big Pouch_Promotion Communication Material Cost Consumer",
                               "Bango Big Pouch_Promotion Communication Material Cost",
                               "Bango Big Pouch_Promo Agency Remuneration Fees & Commissions Consumer",
                               "Bango Big Pouch_Promotional Expenses",
                               "Bango Big Pouch_Promotion Packaging Material Cost",
                               "Bango Big Pouch_Promotion Repacking Cost",
                               "Bango Big Pouch_Promotion Repacking Cost Trade",
                               "Bango Big Pouch_Promotion Repacking Cost Consumer",
                               "Bango Big Pouch_Promo Samples, Gifts and Incentive Costs",
                               "Bango Big Pouch_Promo Samples, Gifts and Incentive Costs Consumer",
                               "Bango Big Pouch_Promo Agency Remun Fees & Commissions",
                               "Bango Big Pouch_Promo Agency Remuneration Fees & Commissions Trade",
                               "Bango Big Pouch_Internal_Penetration",
                               "Bango Big Pouch_Category_Penetration",
                               "Bango Big Pouch_Internal_penetration%"
                               
)
packgroup.bmbp.wj1 <- packgroup.bmbp.wj
colnames(packgroup.bmbp.wj1) <- packgroup_bmbp_wj_headers
packgroup.bmbp.wj.df <- dplyr::bind_rows(packgroup_bmbp_wj_output,packgroup.bmbp.wj1)
packgroup.bmbp.wj.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmbp.wj.df%<>%arrange(Month_ymd)
packgroup.bmbp.wj.df <- packgroup.bmbp.wj.df[order(packgroup.bmbp.wj.df$Customer,decreasing=T),]
packgroup.bmbp.wj.df%<>%mutate_if(is.numeric, as.character)



bmbp_folder_path <- paste0(packgroup_region_dir_path)
bmbs_folder_path <- paste0(packgroup_region_dir_path)
bmmb_folder_path <- paste0(packgroup_region_dir_path)
bmmp_folder_path <- paste0(packgroup_region_dir_path)
bmsb_folder_path <- paste0(packgroup_region_dir_path)
bmsp_folder_path <- paste0(packgroup_region_dir_path)

#packgroup.bmbp.cej.df <- format.data.frame(packgroup.bmbp.cej.df,scientific=F)
#packgroup.bmbp.oi.df <- format.data.frame(packgroup.bmbp.oi.df,scientific=F)
#packgroup.bmbp.sumatera.df <- format.data.frame(packgroup.bmbp.sumatera.df,scientific=F)
#packgroup.bmbp.wj.df <- format.data.frame(packgroup.bmbp.wj.df,scientific=F)


write_xlsx(packgroup.bmbp.cej.df,paste0(bmbp_folder_path,"/",bmbp_cej_fname))
write_xlsx(packgroup.bmbp.oi.df,paste0(bmbp_folder_path,"/",bmbp_oi_fname))
write_xlsx(packgroup.bmbp.sumatera.df,paste0(bmbp_folder_path,"/",bmbp_sumatera_fname))
write_xlsx(packgroup.bmbp.wj.df,paste0(bmbp_folder_path,"/",bmbp_wj_fname))



#final consolidation for the Bango Manis Big sachet : packgroup level
packgroup_bmbs_cej_output%<>%convert(num(`Bango Manis Big Sachet_GSV`:`Indofood_Have a more attractive packaging than other brands`))
packgroup.bmbs.cej%<>%convert(num(`Bango Manis Big Sachet_GSV`:`Bango Manis Big Sachet_Internal_penetration%`))
packgroup.bmbs.cej.df <- dplyr::bind_rows(packgroup_bmbs_cej_output,packgroup.bmbs.cej)
packgroup.bmbp.cej.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmbp.cej.df%<>%arrange(Month_ymd)
packgroup.bmbp.cej.df <- packgroup.bmbp.cej.df[order(packgroup.bmbp.cej.df$Customer,decreasing=T),]
packgroup.bmbp.cej.df%<>%mutate_if(is.numeric, as.character)


packgroup.bmbs.oi%<>%convert(num(`Bango Manis Big Sachet_GSV`:`Bango Manis Big Sachet_Internal_penetration%`))
packgroup.bmbs.oi.df <- dplyr::bind_rows(packgroup_bmbs_oi_output,packgroup.bmbs.oi)
packgroup.bmbs.oi.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmbs.oi.df%<>%arrange(Month_ymd)
packgroup.bmbs.oi.df <- packgroup.bmbp.oi.df[order(packgroup.bmbp.oi.df$Customer,decreasing=T),]
packgroup.bmbs.oi.df%<>%mutate_if(is.numeric, as.character)

packgroup.bmbs.sumatera%<>%convert(num(`Bango Manis Big Sachet_GSV`:`Bango Manis Big Sachet_Internal_penetration%`))
packgroup.bmbs.sumatera.df <- dplyr::bind_rows(packgroup_bmbp_sumatera_output,packgroup.bmbs.sumatera)
packgroup.bmbs.sumatera.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmbs.sumatera.df%<>%arrange(Month_ymd)
packgroup.bmbs.sumatera.df <- packgroup.bmbp.oi.df[order(packgroup.bmbp.oi.df$Customer,decreasing=T),]
packgroup.bmbs.sumatera.df%<>%mutate_if(is.numeric, as.character)

packgroup.bmbs.wj%<>%convert(num(`Bango Manis Big Sachet_GSV`:`Bango Manis Big Sachet_Internal_penetration%`))
packgroup.bmbs.wj.df <- dplyr::bind_rows(packgroup_bmbs_wj_output,packgroup.bmbs.wj)
packgroup.bmbs.wj.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmbs.wj.df%<>%arrange(Month_ymd)
packgroup.bmbs.wj.df <- packgroup.bmbs.wj.df[order(packgroup.bmbs.wj.df$Customer,decreasing=T),]
packgroup.bmbs.wj.df%<>%mutate_if(is.numeric, as.character)



#packgroup.bmbs.cej.df <- format.data.frame(packgroup.bmbs.cej.df,scientific=F)
#packgroup.bmbs.oi.df <- format.data.frame(packgroup.bmbs.oi.df,scientific=F)
#packgroup.bmbs.sumatera.df <- format.data.frame(packgroup.bmbs.sumatera.df,scientific=F)
#packgroup.bmbs.wj.df <- format.data.frame(packgroup.bmbs.wj.df,scientific=F)



write_xlsx(packgroup.bmbs.cej.df,paste0(bmbs_folder_path,"/",bmbs_cej_fname))
write_xlsx(packgroup.bmbs.oi.df,paste0(bmbs_folder_path,"/",bmbs_oi_fname))
write_xlsx(packgroup.bmbs.sumatera.df,paste0(bmbs_folder_path,"/",bmbs_sumatera_fname))
write_xlsx(packgroup.bmbs.wj.df,paste0(bmbs_folder_path,"/",bmbs_wj_fname))

#final consolidation for the Bango Manis Medium Pouch : packgroup level

packgroup.bmmp.cej%<>%convert(num(`Bango Manis Medium Pouch_GSV`:`Bango Manis Medium Pouch_Internal_penetration%`))
packgroup.bmmp.cej.df <- dplyr::bind_rows(packgroup_bmmp_cej_output,packgroup.bmmp.cej)
packgroup.bmmp.cej.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmmp.cej.df%<>%arrange(Month_ymd)
packgroup.bmmp.cej.df <- packgroup.bmmp.cej.df[order(packgroup.bmmp.cej.df$Customer,decreasing=T),]
packgroup.bmmp.cej.df%<>%mutate_if(is.numeric, as.character)

packgroup.bmmp.oi%<>%convert(num(`Bango Manis Medium Pouch_GSV`:`Bango Manis Medium Pouch_Internal_penetration%`))
packgroup.bmmp.oi.df <- dplyr::bind_rows(packgroup_bmmp_oi_output,packgroup.bmmp.oi)
packgroup.bmmp.oi.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmmp.oi.df%<>%arrange(Month_ymd)
packgroup.bmmp.oi.df <- packgroup.bmmp.oi.df[order(packgroup.bmmp.oi.df$Customer,decreasing=T),]
packgroup.bmmp.oi.df%<>%mutate_if(is.numeric, as.character)

packgroup.bmmp.sumatera%<>%convert(num(`Bango Manis Medium Pouch_GSV`:`Bango Manis Medium Pouch_Internal_penetration%`))
packgroup.bmmp.sumatera.df <- dplyr::bind_rows(packgroup_bmmp_sumatera_output,packgroup.bmmp.sumatera)
packgroup.bmmp.sumatera.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmmp.sumatera.df%<>%arrange(Month_ymd)
packgroup.bmmp.sumatera.df <- packgroup.bmmp.sumatera.df[order(packgroup.bmmp.sumatera.df$Customer,decreasing=T),]
packgroup.bmmp.sumatera.df%<>%mutate_if(is.numeric, as.character)


packgroup.bmmp.wj%<>%convert(num(`Bango Manis Medium Pouch_GSV`:`Bango Manis Medium Pouch_Internal_penetration%`))
packgroup.bmmp.wj.df <- dplyr::bind_rows(packgroup_bmmp_wj_output,packgroup.bmmp.wj)
packgroup.bmmp.wj.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmmp.wj.df%<>%arrange(Month_ymd)
packgroup.bmmp.wj.df <- packgroup.bmmp.wj.df[order(packgroup.bmmp.wj.df$Customer,decreasing=T),]
packgroup.bmmp.wj.df%<>%mutate_if(is.numeric, as.character)

#packgroup.bmmp.cej.df <- format.data.frame(packgroup.bmmp.cej.df,scientific=F)
#packgroup.bmmp.oi.df <- format.data.frame(packgroup.bmmp.oi.df,scientific=F)
#packgroup.bmmp.sumatera.df <- format.data.frame(packgroup.bmmp.sumatera.df,scientific=F)
#packgroup.bmmp.wj.df <- format.data.frame(packgroup.bmmp.wj.df,scientific=F)


write_xlsx(packgroup.bmmp.cej.df,paste0(bmmp_folder_path,"/",bmmp_cej_fname))
write_xlsx(packgroup.bmmp.oi.df,paste0(bmmp_folder_path,"/",bmmp_oi_fname))
write_xlsx(packgroup.bmmp.sumatera.df,paste0(bmmp_folder_path,"/",bmmp_sumatera_fname))
write_xlsx(packgroup.bmmp.wj.df,paste0(bmmp_folder_path,"/",bmmp_wj_fname))

#final consolidation for the Bango Manis Medium Bottle : packgroup level

packgroup.bmmb.cej%<>%convert(num(`Bango Manis Medium Bottle_GSV`:`Bango Manis Medium Bottle_Internal_penetration%`))
packgroup.bmmb.cej.df <- dplyr::bind_rows(packgroup_bmmb_cej_output,packgroup.bmmb.cej)
packgroup.bmmb.cej.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmmb.cej.df%<>%arrange(Month_ymd)
packgroup.bmmb.cej.df <- packgroup.bmmb.cej.df[order(packgroup.bmmb.cej.df$Customer,decreasing=T),]
packgroup.bmmb.cej.df%<>%mutate_if(is.numeric, as.character)

packgroup.bmmb.oi%<>%convert(num(`Bango Manis Medium Bottle_GSV`:`Bango Manis Medium Bottle_Internal_penetration%`))
packgroup.bmmb.oi.df <- dplyr::bind_rows(packgroup_bmmb_oi_output,packgroup.bmmb.oi)
packgroup.bmmb.oi.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmmb.oi.df%<>%arrange(Month_ymd)
packgroup.bmmb.oi.df <- packgroup.bmmb.oi.df[order(packgroup.bmmb.oi.df$Customer,decreasing=T),]
packgroup.bmmb.oi.df%<>%mutate_if(is.numeric, as.character)


packgroup.bmmb.sumatera%<>%convert(num(`Bango Manis Medium Bottle_GSV`:`Bango Manis Medium Bottle_Internal_penetration%`))
packgroup.bmmb.sumatera.df <- dplyr::bind_rows(packgroup_bmmb_sumatera_output,packgroup.bmmb.sumatera)
packgroup.bmmb.sumatera.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmmb.sumatera.df%<>%arrange(Month_ymd)
packgroup.bmmb.sumatera.df <- packgroup.bmmb.sumatera.df[order(packgroup.bmmb.sumatera.df$Customer,decreasing=T),]
packgroup.bmmb.sumatera.df%<>%mutate_if(is.numeric, as.character)


packgroup.bmmb.wj%<>%convert(num(`Bango Manis Medium Bottle_GSV`:`Bango Manis Medium Bottle_Internal_penetration%`))
packgroup.bmmb.wj.df <- dplyr::bind_rows(packgroup_bmmb_wj_output,packgroup.bmmb.wj)
packgroup.bmmb.wj.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmmb.wj.df%<>%arrange(Month_ymd)
packgroup.bmmb.wj.df <- packgroup.bmmb.wj.df[order(packgroup.bmmb.wj.df$Customer,decreasing=T),]
packgroup.bmmb.wj.df%<>%mutate_if(is.numeric, as.character)

#packgroup.bmmb.cej.df <- format.data.frame(packgroup.bmmb.cej.df,scientific=F)
#packgroup.bmmb.oi.df <- format.data.frame(packgroup.bmmb.oi.df,scientific=F)
#packgroup.bmmb.sumatera.df <- format.data.frame(packgroup.bmmb.sumatera.df,scientific=F)
#packgroup.bmmb.wj.df <- format.data.frame(packgroup.bmmb.wj.df,scientific=F)


write_xlsx(packgroup.bmmb.cej.df,paste0(bmmb_folder_path,"/",bmmb_cej_fname))
write_xlsx(packgroup.bmmb.oi.df,paste0(bmmb_folder_path,"/",bmmb_oi_fname))
write_xlsx(packgroup.bmmb.sumatera.df,paste0(bmmb_folder_path,"/",bmmb_sumatera_fname))
write_xlsx(packgroup.bmmb.wj.df,paste0(bmmb_folder_path,"/",bmmb_wj_fname))

#final consolidation for the Bango Manis small Bottle : packgroup level

packgroup.bmsb.cej%<>%convert(num(`Bango Manis Small Bottle_GSV`:`Bango Manis Small Bottle_Internal_penetration%`))
packgroup.bmsb.cej.df <- dplyr::bind_rows(packgroup_bmsb_cej_output,packgroup.bmsb.cej)
packgroup.bmsb.cej.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmsb.cej.df%<>%arrange(Month_ymd)
packgroup.bmsb.cej.df <- packgroup.bmsb.cej.df[order(packgroup.bmsb.cej.df$Customer,decreasing=T),]
packgroup.bmsb.cej.df%<>%mutate_if(is.numeric, as.character)

packgroup.bmsb.oi%<>%convert(num(`Bango Manis Small Bottle_GSV`:`Bango Manis Small Bottle_Internal_penetration%`))
packgroup.bmsb.oi.df <- dplyr::bind_rows(packgroup_bmsb_oi_output,packgroup.bmsb.oi)
packgroup.bmsb.oi.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmsb.oi.df%<>%arrange(Month_ymd)
packgroup.bmsb.oi.df <- packgroup.bmsb.oi.df[order(packgroup.bmsb.oi.df$Customer,decreasing=T),]
packgroup.bmsb.oi.df%<>%mutate_if(is.numeric, as.character)

packgroup.bmsb.sumatera%<>%convert(num(`Bango Manis Small Bottle_GSV`:`Bango Manis Small Bottle_Internal_penetration%`))
packgroup.bmsb.sumatera.df <- dplyr::bind_rows(packgroup_bmsb_sumatera_output,packgroup.bmsb.sumatera)
packgroup.bmsb.sumatera.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmsb.sumatera.df%<>%arrange(Month_ymd)
packgroup.bmsb.sumatera.df <- packgroup.bmsb.sumatera.df[order(packgroup.bmsb.sumatera.df$Customer,decreasing=T),]
packgroup.bmsb.sumatera.df%<>%mutate_if(is.numeric, as.character)


packgroup.bmsb.wj%<>%convert(num(`Bango Manis Small Bottle_GSV`:`Bango Manis Small Bottle_Internal_penetration%`))
packgroup.bmsb.wj.df <- dplyr::bind_rows(packgroup_bmsb_wj_output,packgroup.bmsb.wj)
packgroup.bmsb.wj.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmsb.wj.df%<>%arrange(Month_ymd)
packgroup.bmsb.wj.df <- packgroup.bmsb.wj.df[order(packgroup.bmsb.wj.df$Customer,decreasing=T),]
packgroup.bmsb.wj.df%<>%mutate_if(is.numeric, as.character)
#packgroup.bmsb.cej.df <- format.data.frame(packgroup.bmsb.cej.df,scientific=F)
#packgroup.bmsb.oi.df <- format.data.frame(packgroup.bmsb.oi.df,scientific=F)
#packgroup.bmsb.sumatera.df <- format.data.frame(packgroup.bmsb.sumatera.df,scientific=F)
#packgroup.bmsb.wj.df <- format.data.frame(packgroup.bmsb.wj.df,scientific=F)


write_xlsx(packgroup.bmsb.cej.df,paste0(bmsb_folder_path,"/",bmsb_cej_fname))
write_xlsx(packgroup.bmsb.oi.df,paste0(bmsb_folder_path,"/",bmsb_oi_fname))
write_xlsx(packgroup.bmsb.sumatera.df,paste0(bmsb_folder_path,"/",bmsb_sumatera_fname))
write_xlsx(packgroup.bmsb.wj.df,paste0(bmsb_folder_path,"/",bmsb_wj_fname))

#final consolidation for the Bango Manis small Pouch : packgroup level

packgroup.bmsp.cej%<>%convert(num(`Bango Manis Small Pouch_GSV`:`Bango Manis Small Pouch_Internal_penetration%`))
packgroup.bmsp.cej.df <- dplyr::bind_rows(packgroup_bmsp_cej_output,packgroup.bmsp.cej)
packgroup.bmsp.cej.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmsp.cej.df%<>%arrange(Month_ymd)
packgroup.bmsp.cej.df <- packgroup.bmsp.cej.df[order(packgroup.bmsp.cej.df$Customer,decreasing=T),]
packgroup.bmsp.cej.df%<>%mutate_if(is.numeric, as.character)

packgroup.bmsp.oi%<>%convert(num(`Bango Manis Small Pouch_GSV`:`Bango Manis Small Pouch_Internal_penetration%`))
packgroup.bmsp.oi.df <- dplyr::bind_rows(packgroup_bmsp_oi_output,packgroup.bmsp.oi)
packgroup.bmsp.oi.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmsp.oi.df%<>%arrange(Month_ymd)
packgroup.bmsp.oi.df <- packgroup.bmsp.oi.df[order(packgroup.bmsp.oi.df$Customer,decreasing=T),]
packgroup.bmsp.oi.df%<>%mutate_if(is.numeric, as.character)

packgroup.bmsp.sumatera%<>%convert(num(`Bango Manis Small Pouch_GSV`:`Bango Manis Small Pouch_Internal_penetration%`))
packgroup.bmsp.sumatera.df <- dplyr::bind_rows(packgroup_bmsp_sumatera_output,packgroup.bmsp.sumatera)
packgroup.bmsp.sumatera.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmsp.sumatera.df%<>%arrange(Month_ymd)
packgroup.bmsp.sumatera.df <- packgroup.bmsp.sumatera.df[order(packgroup.bmsp.sumatera.df$Customer,decreasing=T),]
packgroup.bmsp.sumatera.df%<>%mutate_if(is.numeric, as.character)


packgroup.bmsp.wj%<>%convert(num(`Bango Manis Small Pouch_GSV`:`Bango Manis Small Pouch_Internal_penetration%`))
packgroup.bmsp.wj.df <- dplyr::bind_rows(packgroup_bmsp_wj_output,packgroup.bmsp.wj)
packgroup.bmsp.wj.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmsp.wj.df%<>%arrange(Month_ymd)
packgroup.bmsp.wj.df <- packgroup.bmsp.wj.df[order(packgroup.bmsp.wj.df$Customer,decreasing=T),]
packgroup.bmsp.wj.df%<>%mutate_if(is.numeric, as.character)

#packgroup.bmsp.cej.df <- format.data.frame(packgroup.bmsp.cej.df,scientific=F)
#packgroup.bmsp.oi.df <- format.data.frame(packgroup.bmsp.oi.df,scientific=F)
#packgroup.bmsp.sumatera.df <- format.data.frame(packgroup.bmsp.sumatera.df,scientific=F)
#packgroup.bmsp.wj.df <- format.data.frame(packgroup.bmsp.wj.df,scientific=F)


write_xlsx(packgroup.bmsp.cej.df,paste0(bmsp_folder_path,"/",bmsp_cej_fname))
write_xlsx(packgroup.bmsp.oi.df,paste0(bmsp_folder_path,"/",bmsp_oi_fname))
write_xlsx(packgroup.bmsp.sumatera.df,paste0(bmsp_folder_path,"/",bmsp_sumatera_fname))
write_xlsx(packgroup.bmsp.wj.df,paste0(bmsp_folder_path,"/",bmsp_wj_fname))

##################################################### Basepack - Account ( 27 output files - 3*9=27) ##############################################

#There are 27 output files generated by the Basepack Account where each Account i.e. indomart , Indomart and Carrefour has 9 basepack files which are as follows
#BANGO KECAP MANIS 275ML:bkm275 ,BANGO KECAP MANIS 620ML:bkm620 ,BANGO KECAP MANIS 220ML:bkm220 ,BANGO KECAP MANIS PEDAS 220ML:bkmp220 ,BANGO KECAP MANIS PEDAS 135ML: bkmp135,
#BANGO SOY/OYSTER/FISH SAUCE MANIS 60ML:bsofs60,BANGO SOY/OYSTR/FISH SUCE MNIS SCHT 30ML:bsofms30,BANGO KECAP MANIS 135ML:bkm135 ,BANGO SOY/OYSTER/FISH KECAP MANIS 580ML:bsof580

basepack_account_path <- "C:/Users/Gourab.Chanda/OneDrive - Unilever/ID_Jarvis_Bango-HICW10NA512SG4H/Final_Files/Basepack-Account"
basepack_bkm275 <- list.files(path =basepack_account_path ,pattern = "^GD ID.*BP5(.*)xlsx|XLSX$",recursive = T)
basepack_bkm620 <-  list.files(path =basepack_account_path ,pattern = "^GD ID.*BP8(.*)xlsx|XLSX$",recursive = T)
basepack_bkm220 <- list.files(path =basepack_account_path ,pattern = "^GD ID.*BP3(.*)xlsx|XLSX$",recursive = T)
basepack_bkmp220 <- list.files(path =basepack_account_path ,pattern = "^GD ID.*BP7(.*)xlsx|XLSX$",recursive = T)
basepack_bkmp135 <- list.files(path =basepack_account_path ,pattern = "^GD ID.*BP9(.*)xlsx|XLSX$",recursive = T)
basepack_bsofs60  <- list.files(path =basepack_account_path ,pattern = "^GD ID.*BP4(.*)xlsx|XLSX$",recursive = T)
basepack_bsofms30 <- list.files(path =basepack_account_path ,pattern = "^GD ID.*BP2(.*)xlsx|XLSX$",recursive = T)
basepack_bkm135 <- list.files(path =basepack_account_path ,pattern = "^GD ID.*BP6(.*)xlsx|XLSX$",recursive = T)
basepack_bsof580 <- list.files(path =basepack_account_path ,pattern = "^GD ID.*BP1(.*)xlsx|XLSX$",recursive = T)

#Alfamart 9 output files
alfamart_bkm275_output <- read_xlsx(paste0(basepack_account_path,"/",basepack_bkm275[1]),sheet = 1,guess_max = 100)
alfamart_bkm620_output <- read_xlsx(paste0(basepack_account_path,"/",basepack_bkm620[1]),sheet = 1,guess_max = 100)
alfamart_bkm220_output <- read_xlsx(paste0(basepack_account_path,"/",basepack_bkm220[1]),sheet = 1,guess_max = 100)
alfamart_bkmp220_output <- read_xlsx(paste0(basepack_account_path,"/",basepack_bkmp220[1]),sheet = 1,guess_max = 100)
alfamart_bkmp135_output <- read_xlsx(paste0(basepack_account_path,"/",basepack_bkmp135[1]),sheet = 1,guess_max = 100)
alfamart_bsofs60_output <- read_xlsx(paste0(basepack_account_path,"/",basepack_bsofs60[1]),sheet = 1,guess_max = 100)
alfamart_bsofms30_output <- read_xlsx(paste0(basepack_account_path,"/",basepack_bsofms30[1]),sheet = 1,guess_max = 100)
alfamart_bkm135_output <-  read_xlsx(paste0(basepack_account_path,"/",basepack_bkm135[1]),sheet = 1,guess_max = 100)
alfamart_bsof580_output <- read_xlsx(paste0(basepack_account_path,"/",basepack_bsof580[1]),sheet = 1,guess_max = 100)

#Indomart 9 output files 

indomart_bkm275_output <- read_xlsx(paste0(basepack_account_path,"/",basepack_bkm275[3]),sheet = 1,guess_max = 100)
indomart_bkm620_output <- read_xlsx(paste0(basepack_account_path,"/",basepack_bkm620[3]),sheet = 1,guess_max = 100)
indomart_bkm220_output <- read_xlsx(paste0(basepack_account_path,"/",basepack_bkm220[3]),sheet = 1,guess_max = 100)
indomart_bkmp220_output <- read_xlsx(paste0(basepack_account_path,"/",basepack_bkmp220[3]),sheet = 1,guess_max = 100)
indomart_bkmp135_output <- read_xlsx(paste0(basepack_account_path,"/",basepack_bkmp135[3]),sheet = 1,guess_max = 100)
indomart_bsofs60_output <- read_xlsx(paste0(basepack_account_path,"/",basepack_bsofs60[3]),sheet = 1,guess_max = 100)
indomart_bsofms30_output <- read_xlsx(paste0(basepack_account_path,"/",basepack_bsofms30[3]),sheet = 1,guess_max = 100)
indomart_bkm135_output <-  read_xlsx(paste0(basepack_account_path,"/",basepack_bkm135[3]),sheet = 1,guess_max = 100)
indomart_bsof580_output <- read_xlsx(paste0(basepack_account_path,"/",basepack_bsof580[3]),sheet = 1,guess_max = 100)

#carrefour 9 output files 

carrefour_bkm275_output <- read_xlsx(paste0(basepack_account_path,"/",basepack_bkm275[2]),sheet = 1,guess_max = 100)
carrefour_bkm620_output <- read_xlsx(paste0(basepack_account_path,"/",basepack_bkm620[2]),sheet = 1,guess_max = 100)
carrefour_bkm220_output <- read_xlsx(paste0(basepack_account_path,"/",basepack_bkm220[2]),sheet = 1,guess_max = 100)
carrefour_bkmp220_output <- read_xlsx(paste0(basepack_account_path,"/",basepack_bkmp220[2]),sheet = 1,guess_max = 100)
carrefour_bkmp135_output <- read_xlsx(paste0(basepack_account_path,"/",basepack_bkmp135[2]),sheet = 1,guess_max = 100)
carrefour_bsofs60_output <- read_xlsx(paste0(basepack_account_path,"/",basepack_bsofs60[2]),sheet = 1,guess_max = 100)
carrefour_bsofms30_output <- read_xlsx(paste0(basepack_account_path,"/",basepack_bsofms30[2]),sheet = 1,guess_max = 100)
carrefour_bkm135_output <-  read_xlsx(paste0(basepack_account_path,"/",basepack_bkm135[2]),sheet = 1,guess_max = 100)
carrefour_bsof580_output <- read_xlsx(paste0(basepack_account_path,"/",basepack_bsof580[2]),sheet = 1,guess_max = 100)

####################################################### Basepack - Account ( Alfamart) #################################################################

alfamart.bkm275%<>%convert(num(GSV:`Sell out Volume (KGs)`))
alfamart.bkm275.df <- bind_rows(alfamart_bkm275_output,alfamart.bkm275)
alfamart.bkm275.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
alfamart.bkm275.df%<>%arrange(Month_ymd)
alfamart.bkm275.df <- alfamart.bkm275.df[order(alfamart.bkm275.df$Customer,decreasing=T),]
alfamart.bkm275.df%<>%mutate_if(is.numeric, as.character)

alfamart.bkm620%<>%convert(num(GSV:`Sell out Volume (KGs)`))
alfamart.bkm620.df <- bind_rows(alfamart_bkm620_output,alfamart.bkm620)
alfamart.bkm620.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
alfamart.bkm620.df%<>%arrange(Month_ymd)
alfamart.bkm620.df <- alfamart.bkm620.df[order(alfamart.bkm620.df$Customer,decreasing=T),]
alfamart.bkm620.df%<>%mutate_if(is.numeric, as.character)

alfamart.bkm220%<>%convert(num(GSV:`Sell out Volume (KGs)`))
alfamart.bkm220.df <- bind_rows(alfamart_bkm220_output,alfamart.bkm220)
alfamart.bkm220.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
alfamart.bkm220.df%<>%arrange(Month_ymd)
alfamart.bkm220.df <- alfamart.bkm220.df[order(alfamart.bkm220.df$Customer,decreasing=T),]
alfamart.bkm220.df%<>%mutate_if(is.numeric, as.character)

alfamart.bkmp220%<>%convert(num(GSV:`Sell out Volume (KGs)`))
alfamart.bkmp220.df <- bind_rows(alfamart_bkmp220_output,alfamart.bkmp220)
alfamart.bkmp220.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
alfamart.bkmp220.df%<>%arrange(Month_ymd)
alfamart.bkmp220.df <- alfamart.bkmp220.df[order(alfamart.bkmp220.df$Customer,decreasing=T),]
alfamart.bkmp220.df%<>%mutate_if(is.numeric, as.character)



alfamart.bkmp135%<>%convert(num(GSV:`Sell out Volume (KGs)`))
alfamart.bkmp135.df <- bind_rows(alfamart_bkmp135_output,alfamart.bkmp135)
alfamart.bkmp135.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
alfamart.bkmp135.df%<>%arrange(Month_ymd)
alfamart.bkmp135.df <- alfamart.bkmp135.df[order(alfamart.bkmp135.df$Customer,decreasing=T),]
alfamart.bkmp135.df <- alfamart.bkmp135.df[order(alfamart.bkmp135.df$Customer,decreasing=T),]
alfamart.bkmp135.df%<>%mutate_if(is.numeric, as.character)


alfamart.bsofs60%<>%convert(num(GSV:`Sell out Volume (KGs)`))
alfamart.bsofs60.df <- bind_rows(alfamart_bsofs60_output,alfamart.bsofs60)
alfamart.bsofs60.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
alfamart.bsofs60.df%<>%arrange(Month_ymd)
alfamart.bsofs60.df <- alfamart.bsofs60.df[order(alfamart.bsofs60.df$Customer,decreasing=T),]
alfamart.bsofs60.df%<>%mutate_if(is.numeric, as.character)


alfamart.bsofms30%<>%convert(num(GSV:`Sell out Volume (KGs)`))
alfamart.bsofms30.df <- bind_rows(alfamart_bsofms30_output,alfamart.bsofms30)
alfamart.bsofms30.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
alfamart.bsofms30.df%<>%arrange(Month_ymd)
alfamart.bsofms30.df <- alfamart.bsofms30.df[order(alfamart.bsofms30.df$Customer,decreasing=T),]
alfamart.bsofms30.df%<>%mutate_if(is.numeric, as.character)

alfamart.bkm135%<>%convert(num(GSV:`Sell out Volume (KGs)`))
alfamart.bkm135.df <- bind_rows(alfamart_bkm135_output,alfamart.bkm135)
alfamart.bkm135.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
alfamart.bkm135.df%<>%arrange(Month_ymd)
alfamart.bkm135.df <- alfamart.bkm135.df[order(alfamart.bkm135.df$Customer,decreasing=T),]
alfamart.bkm135.df%<>%mutate_if(is.numeric, as.character)


alfamart.bsof580_updated <- setnames(alfamart.bsof580,new = "Sales Qty",old = "Primary Sales Qty(PC)")
alfamart.bsof580_updated%<>%convert(num(GSV:`Sell out Volume (KGs)`))
alfamart.bsof580.df <- bind_rows(alfamart_bsof580_output,alfamart.bsof580_updated)
alfamart.bsof580.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
alfamart.bsof580.df%<>%arrange(Month_ymd)
alfamart.bsof580.df <- alfamart.bsof580.df[order(alfamart.bsof580.df$Customer,decreasing=T),]
alfamart.bsof580.df%<>%mutate_if(is.numeric, as.character)


write_xlsx(alfamart.bkm275.df,paste0(basepack_account_path,"/alfamart.bkm275.df.xlsx"))
write_xlsx(alfamart.bkm620,paste0(basepack_account_path,"/alfamart.bkm620.xlsx"))
write_xlsx(alfamart.bkm220.df,paste0(basepack_account_path,"/alfamart.bkm220.df.xlsx"))
write_xlsx(alfamart.bkmp220.df,paste0(basepack_account_path,"/alfamart.bkmp220.df.xlsx"))
write_xlsx(alfamart.bkmp135.df,paste0(basepack_account_path,"/alfamart.bkmp135.df.xlsx"))
write_xlsx(alfamart.bsofs60.df,paste0(basepack_account_path,"/alfamart.bsofs60.df.xlsx"))
write_xlsx(alfamart.bsofms30.df,paste0(basepack_account_path,"/alfamart.bsofms30.df.xlsx"))
write_xlsx(alfamart.bkm135.df,paste0(basepack_account_path,"/alfamart.bkm135.df.xlsx"))
write_xlsx(alfamart.bsof580.df,paste0(basepack_account_path,"/alfamart.bsof580.df.xlsx"))

####################################################### Basepack - Account ( Carrefour) #################################################################


Carrefour.bkm275%<>%convert(num(GSV:`Promo Agency Remuneration Fees & Commissions Consumer`))
carrefour.bkm275.df <- bind_rows(carrefour_bkm275_output,Carrefour.bkm275)
carrefour.bkm275.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
carrefour.bkm275.df%<>%arrange(Month_ymd)
carrefour.bkm275.df <- carrefour.bkm275.df[order(carrefour.bkm275.df$Customer,decreasing=T),]
carrefour.bkm275.df%<>%mutate_if(is.numeric, as.character)

Carrefour.bkm620%<>%convert(num(GSV:`Promo Agency Remuneration Fees & Commissions Consumer`))
carrefour.bkm620.df <- bind_rows(carrefour_bkm620_output,Carrefour.bkm620)
carrefour.bkm620.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
carrefour.bkm620.df%<>%arrange(Month_ymd)
carrefour.bkm620.df <- carrefour.bkm620.df[order(carrefour.bkm620.df$Customer,decreasing=T),]
carrefour.bkm620.df%<>%mutate_if(is.numeric, as.character)

Carrefour.bkm220%<>%convert(num(GSV:`Promo Agency Remuneration Fees & Commissions Consumer`))
carrefour.bkm220.df <- bind_rows(carrefour_bkm220_output,Carrefour.bkm220)
carrefour.bkm220.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
carrefour.bkm220.df%<>%arrange(Month_ymd)
carrefour.bkm220.df <- carrefour.bkm220.df[order(carrefour.bkm220.df$Customer,decreasing=T),]
carrefour.bkm220.df%<>%mutate_if(is.numeric, as.character)

#column names varies in carrefour_bkmp220_output and Carrefour.bkmp220 (month column)
Carrefour.bkmp220%<>%convert(num(GSV:`Promo Agency Remuneration Fees & Commissions Consumer`))
carrefour.bkmp220.df <- bind_rows(carrefour_bkmp220_output,Carrefour.bkmp220)
carrefour.bkmp220.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
carrefour.bkmp220.df%<>%arrange(Month_ymd)
carrefour.bkmp220.df <- carrefour.bkmp220.df[order(carrefour.bkmp220.df$Customer,decreasing=T),]
carrefour.bkmp220.df%<>%mutate_if(is.numeric, as.character)


Carrefour.bkmp135%<>%convert(num(GSV:`Promo Agency Remuneration Fees & Commissions Consumer`))
carrefour.bkmp135.df <- bind_rows(carrefour_bkmp135_output,Carrefour.bkmp135)
carrefour.bkmp135.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
carrefour.bkmp135.df%<>%arrange(Month_ymd)
carrefour.bkmp135.df <- carrefour.bkmp135.df[order(carrefour.bkmp135.df$Customer,decreasing=T),]
carrefour.bkmp135.df%<>%mutate_if(is.numeric, as.character)

Carrefour.bsofs60%<>%convert(num(GSV:`Promo Agency Remuneration Fees & Commissions Consumer`))
carrefour.bsofs60.df <- bind_rows(carrefour_bsofs60_output,Carrefour.bsofs60)
carrefour.bsofs60.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
carrefour.bsofs60.df%<>%arrange(Month_ymd)
carrefour.bsofs60.df <- carrefour.bsofs60.df[order(carrefour.bsofs60.df$Customer,decreasing=T),]
carrefour.bsofs60.df%<>%mutate_if(is.numeric, as.character)

Carrefour.bsofms30%<>%convert(num(GSV:`Promo Agency Remuneration Fees & Commissions Consumer`))
carrefour.bsofms30.df <- bind_rows(carrefour_bsofms30_output,Carrefour.bsofms30)
carrefour.bsofms30.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
carrefour.bsofms30.df%<>%arrange(Month_ymd)
carrefour.bsofms30.df <- carrefour.bsofms30.df[order(carrefour.bsofms30.df$Customer,decreasing=T),]
carrefour.bsofms30.df%<>%mutate_if(is.numeric, as.character)

Carrefour.bkm135%<>%convert(num(GSV:`Promo Agency Remuneration Fees & Commissions Consumer`))
carrefour.bkm135.df <- bind_rows(carrefour_bkm135_output,Carrefour.bkm135)
carrefour.bkm135.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
carrefour.bkm135.df%<>%arrange(Month_ymd)
carrefour.bkm135.df <- carrefour.bkm135.df[order(carrefour.bkm135.df$Customer,decreasing=T),]
carrefour.bkm135.df%<>%mutate_if(is.numeric, as.character)

#**have to work on this 
#carrefour.bsof580_updated <- Carrefour.bsof580%>%dplyr::rename(`Primary Sales Qty(PC)`=`Sales Qty`,`Original Order Qty.`=`OriginalOrder Qty.`)
carrefour.bsof580_updated <- Carrefour.bsof580


#carrefour.bsof580_updated <- setnames(Carrefour.bsof580,new = "Sales Qty",old = "Primary Sales Qty(PC)")
carrefour.bsof580_updated%<>%convert(num(GSV:`Promo Agency Remuneration Fees & Commissions Consumer`))
carrefour.bsof580.df <- bind_rows(carrefour_bsof580_output,carrefour.bsof580_updated)
carrefour.bsof580.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
carrefour.bsof580.df%<>%arrange(Month_ymd)
carrefour.bsof580.df <- carrefour.bsof580.df[order(carrefour.bsof580.df$Customer,decreasing=T),]
carrefour.bsof580.df%<>%mutate_if(is.numeric, as.character)


write_xlsx(carrefour.bkm275.df,paste0(basepack_account_path,"/carrefour.bkm275.df.xlsx"))
write_xlsx(carrefour.bkm620.df,paste0(basepack_account_path,"/carrefour.bkm620.df.xlsx"))
write_xlsx(carrefour.bkm220.df,paste0(basepack_account_path,"/carrefour.bkm220.df.xlsx"))
write_xlsx(carrefour.bkmp220.df,paste0(basepack_account_path,"/carrefour.bkmp220.df.xlsx"))
write_xlsx(carrefour.bkmp135.df,paste0(basepack_account_path,"/carrefour.bkmp135.df.xlsx"))
write_xlsx(carrefour.bsofs60.df,paste0(basepack_account_path,"/carrefour.bsofs60.df.xlsx"))
write_xlsx(carrefour.bsofms30.df,paste0(basepack_account_path,"/carrefour.bsofms30.df.xlsx"))
write_xlsx(carrefour.bkm135.df,paste0(basepack_account_path,"/carrefour.bkm135.df.xlsx"))
write_xlsx(carrefour.bsof580.df,paste0(basepack_account_path,"/carrefour.bsof580.df.xlsx"))

###################################################### Basepack (Indomart) ##############################################

indomart.bkm275%<>%convert(num(Bango_GSV:`Sell out Volume (KGs)`))
indomart.bkm275.df <- bind_rows(indomart_bkm275_output,indomart.bkm275)
indomart.bkm275.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
indomart.bkm275.df%<>%arrange(Month_ymd)
indomart.bkm275.df <- indomart.bkm275.df[order(indomart.bkm275.df$Customer,decreasing=T),]
indomart.bkm275.df%<>%mutate_if(is.numeric, as.character)

indomart.bkm620%<>%dplyr::rename(`Sell out Value (IDR)...53`=`Sell out Value (IDR)`,`Unit Price...54`=`Unit Price`,`Sell out Value (IDR)...56`=`Sell out Volume (KGs)`)
indomart.bkm620%<>%convert(num(Bango_GSV:ncol(indomart.bkm620)))
indomart.bkm620.df <- bind_rows(indomart_bkm620_output,indomart.bkm620)
indomart.bkm620.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
indomart.bkm620.df%<>%arrange(Month_ymd)
indomart.bkm620.df <- indomart.bkm620.df[order(indomart.bkm620.df$Customer,decreasing=T),]
indomart.bkm620.df%<>%mutate_if(is.numeric, as.character)

indomart.bkm220%<>%convert(num(Bango_GSV:`Sell out Volume (KGs)`))
indomart.bkm220.df <- bind_rows(indomart_bkm220_output,indomart.bkm220)
indomart.bkm220.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
indomart.bkm220.df%<>%arrange(Month_ymd)
indomart.bkm220.df <- indomart.bkm220.df[order(indomart.bkm220.df$Customer,decreasing=T),]
indomart.bkm220.df%<>%mutate_if(is.numeric, as.character)

indomart.bkmp220%<>%dplyr::rename(`Bango_Primary Sales Qty(KG)`=`Bango_Primary Sales Qty(Kg)`)
indomart.bkmp220%<>%convert(num(Bango_GSV:`Sell out Volume (KGs)`))
indomart.bkmp220.df <- bind_rows(indomart_bkmp220_output,indomart.bkmp220)
indomart.bkmp220.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
indomart.bkmp220.df%<>%arrange(Month_ymd)
indomart.bkmp220.df <- indomart.bkmp220.df[order(indomart.bkmp220.df$Customer,decreasing=T),]
indomart.bkmp220.df%<>%mutate_if(is.numeric, as.character)


indomart.bkmp135%<>%dplyr::rename(`Bango_Primary Sales Qty(Tonn)`=`Bango_Primary Sales Qty(Tonns)`,`Bango_Primary Sales Qty(kg)`=`Bango_Primary Sales Qty(Kg)`)
indomart.bkmp135%<>%convert(num(Bango_GSV:`Sell out Volume (KGs)`))
indomart.bkmp135.df <- bind_rows(indomart_bkmp135_output,indomart.bkmp135)
indomart.bkmp135.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
indomart.bkmp135.df%<>%arrange(Month_ymd)
indomart.bkmp135.df <- indomart.bkmp135.df[order(indomart.bkmp135.df$Customer,decreasing=T),]
indomart.bkmp135.df <- indomart.bkmp135.df[order(indomart.bkmp135.df$Customer,decreasing=T),]
indomart.bkmp135.df%<>%mutate_if(is.numeric, as.character)


indomart.bsofs60%<>%convert(num(Bango_GSV:`Sell out Volume (KGs)`))
indomart.bsofs60.df <- bind_rows(indomart_bsofs60_output,indomart.bsofs60)
indomart.bsofs60.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
indomart.bsofs60.df%<>%arrange(Month_ymd)
indomart.bsofs60.df <- indomart.bsofs60.df[order(indomart.bsofs60.df$Customer,decreasing=T),]
indomart.bsofs60.df%<>%mutate_if(is.numeric, as.character)


indomart.bsofms30%<>%convert(num(Bango_GSV:`Sell out Volume (KGs)`))
indomart.bsofms30.df <- bind_rows(indomart_bsofms30_output,indomart.bsofms30)
indomart.bsofms30.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
indomart.bsofms30.df%<>%arrange(Month_ymd)
indomart.bsofms30.df <- indomart.bsofms30.df[order(indomart.bsofms30.df$Customer,decreasing=T),]
indomart.bsofms30.df%<>%mutate_if(is.numeric, as.character)

indomart.bkm135%<>%convert(num(Bango_GSV:`Sell out Volume (KGs)`))
indomart.bkm135.df <- bind_rows(indomart_bkm135_output,indomart.bkm135)
indomart.bkm135.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
indomart.bkm135.df%<>%arrange(Month_ymd)
indomart.bkm135.df <- indomart.bkm135.df[order(indomart.bkm135.df$Customer,decreasing=T),]
indomart.bkm135.df%<>%mutate_if(is.numeric, as.character)



indomart.bsof580%<>%dplyr::rename(`Primary Sales Qty(PC)`=`Bango_Primary Sales Qty(PC)`,`Primary Sales Qty(Kg)`=`Bango_Primary Sales Qty(Kg)`,`Primary Sales Qty(Tonns)`=`Bango_Primary Sales Qty(Tonns)`)
indomart.bsof580_updated <- indomart.bsof580
indomart.bsof580_updated%<>%convert(num(Bango_GSV:`Sell out Volume (KGs)`))
indomart.bsof580.df <- bind_rows(indomart_bsof580_output,indomart.bsof580_updated)
indomart.bsof580.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
indomart.bsof580.df%<>%arrange(Month_ymd)
indomart.bsof580.df <- indomart.bsof580.df[order(indomart.bsof580.df$Customer,decreasing=T),]
indomart.bsof580.df%<>%mutate_if(is.numeric, as.character)


write_xlsx(indomart.bkm275.df,paste0(basepack_account_path,"/indomart.bkm275.df.xlsx"))
write_xlsx(indomart.bkm620.df,paste0(basepack_account_path,"/indomart.bkm620.df.xlsx"))
write_xlsx(indomart.bkm220.df,paste0(basepack_account_path,"/indomart.bkm220.df.xlsx"))
write_xlsx(indomart.bkmp220.df,paste0(basepack_account_path,"/indomart.bkmp220.df.xlsx"))
write_xlsx(indomart.bkmp135.df,paste0(basepack_account_path,"/indomart.bkmp135.df.xlsx"))
write_xlsx(indomart.bsofs60.df,paste0(basepack_account_path,"/indomart.bsofs60.df.xlsx"))
write_xlsx(indomart.bsofms30.df,paste0(basepack_account_path,"/indomart.bsofms30.df.xlsx"))
write_xlsx(indomart.bkm135.df,paste0(basepack_account_path,"/indomart.bkm135.df.xlsx"))
write_xlsx(indomart.bsof580.df,paste0(basepack_account_path,"/indomart.bsof580.df.xlsx"))





end_time <- Sys.time()
script_time <- round(end_time-start_time,2)
print(script_time)
Print(paste0("The Final Consolidation time ",script_time))