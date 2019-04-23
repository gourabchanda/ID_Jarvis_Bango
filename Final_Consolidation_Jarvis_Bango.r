############################################ Final COnsolidation Script ##########################################################
start_time <- Sys.time()
working_directory <- setwd("C:/Users/goura/OneDrive/Desktop/Unilever_R_Projects/ID_Jarvis_Bango/ID_Jarvis_2.0.0")
#source(paste0(working_directory,"/data_preprocessing_script.r"))
#source(paste0(working_directory,"/ID_MT_Final.r"))
source(paste0(working_directory,"/ID_Jarvis_blending_script.r"))

############################################# Bango-GT Model ######################################################################
bango_gt_output_path <- "C:/Users/goura/OneDrive/Desktop/Unilever_Projects/ID_Jarvis_Bango/Final_Consolidated_Output_Files/Final_Files/Bango-GT"
bango_gt_output_files <- list.files(path = bango_gt_output_path ,pattern = "^GD ID GT -Bango(.*)xlsx|XLSX$")
bango_gt_output_file.df <- read_xlsx(paste0(bango_gt_output_path,"/",bango_gt_output_files),sheet = 1,guess_max = 100)
total.bango.gt.final%<>%convert(num(`Bango GT_GSV`:`Bango_I Trust %`))
bango.gt.df <- dplyr::bind_rows(bango_gt_output_file.df,total.bango.gt.final)
bango.gt.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.gt.df%<>%arrange(Month_ymd)
bango.gt.df <- bango.gt.df[order(bango.gt.df$Channel),]
write_xlsx(bango.gt.df,paste0(working_directory,"/bango.gt.df.xlsx"))
############################################# Bango - MT Model ###################################################################
bango_mt_output_path <- "C:/Users/goura/OneDrive/Desktop/Unilever_Projects/ID_Jarvis_Bango/Final_Consolidated_Output_Files/Final_Files/Bango-MT"
bango_mt_output_files <- list.files(path = bango_mt_output_path ,pattern = "^GD ID MT -Bango(.*)xlsx|XLSX$")
bango_mt_output_file.df <- read_xlsx(paste0(bango_mt_output_path,"/",bango_mt_output_files),sheet = 1,guess_max = 100)
total.bango.mt.final%<>%convert(num(Bango_GSV:`Bango_Promo Agency Remuneration Fees & Commissions Trade`))
bango.mt.df <- dplyr::bind_rows(bango_mt_output_file.df,total.bango.mt.final)
bango.mt.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.mt.df%<>%arrange(Month_ymd)
bango.mt.df <- bango.mt.df[order(bango.mt.df$Channel),]
write_xlsx(bango.mt.df,paste0(working_directory,"/bango.mt.df.xlsx"))
############################################ Bango - National Model ###############################################################
bango_national_output_path <- "C:/Users/goura/OneDrive/Desktop/Unilever_Projects/ID_Jarvis_Bango/Final_Consolidated_Output_Files/Final_Files/Bango-National"
bango_national_output_files <- list.files(path = bango_national_output_path ,pattern = "^GD ID Indonesia-Bango(.*)xlsx|XLSX$")
bango_national_output_file.df <- read_xlsx(paste0(bango_national_output_path,"/",bango_national_output_files),sheet = 1,guess_max = 100)
total.final.bango%<>%convert(num(Bango_GSV:`Bango_Penetration%`))
bango.national.df <- dplyr::bind_rows(bango_national_output_file.df,total.final.bango)
bango.national.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.national.df%<>%arrange(Month_ymd)
bango.national.df <- bango.mt.df[order(bango.national.df$Market),]
write_xlsx(bango.national.df,paste0(working_directory,"/bango.national.df.xlsx"))
########################################### Bango - Region #########################################################################

bango_region_output_path <- "C:/Users/goura/OneDrive/Desktop/Unilever_Projects/ID_Jarvis_Bango/Final_Consolidated_Output_Files/Final_Files/Bango-Region"
bango_region_output_files <- list.files(path=bango_region_output_path,pattern="^GD ID GT(.*)xlsx|XLSX$")


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
bango.cej.df <- bango.cej.df[order(bango.cej.df$Customer),]


bango.oi.df <- dplyr::bind_rows(bango_region_oi_output,bango.region.oi)
bango.oi.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.oi.df%<>%arrange(Month_ymd)
bango.oi.df <- bango.oi.df[order(bango.oi.df$Customer),]


bango.sumatra.df <- dplyr::bind_rows(bango_region_sumatra_output,bango.region.sumatra)
bango.sumatra.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.sumatra.df%<>%arrange(Month_ymd)
bango.sumatra.df <- bango.cej.df[order(bango.sumatra.df$Customer),]


bango.wj.df <- dplyr::bind_rows(bango_region_wj_output,bango.region.wj)
bango.wj.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.wj.df%<>%arrange(Month_ymd)
bango.wj.df <- bango.wj.df[order(bango.wj.df$Customer),]


write_xlsx(bango.cej.df,paste0(working_directory,"/bango.cej.df"))
write_xlsx(bango.oi.df,paste0(working_directory,"/bango.oi.df.xlsx"))
write_xlsx(bango.sumatra.df,paste0(working_directory,"/bango.sumatra.df.xlsx"))
write_xlsx(bango.wj.df,paste0(working_directory,"/bango.wj.df.xlsx"))



############################################################### Bango - Sub Channel #########################################

bango_subchannel_output_path <- "C:/Users/goura/OneDrive/Desktop/Unilever_Projects/ID_Jarvis_Bango/Final_Consolidated_Output_Files/Final_Files/Bango-Sub channel"

bango_subchannel_output_files <- list.files(path=bango_subchannel_output_path,pattern="^GD ID MT(.*)xlsx|XLSX$")


bango_subchannel_hsm_output <-  read_xlsx(paste0(bango_subchannel_output_path,"/",bango_subchannel_output_files[1]),sheet = 1,guess_max = 100)
bango_subchannel_mini_output <- read_xlsx(paste0(bango_subchannel_output_path,"/",bango_subchannel_output_files[2]),sheet = 1,guess_max = 100)


HSM.final.data%<>%convert(num(Bango_GSV:`Bango_Promo Agency Remuneration Fees & Commissions Consumer`))
Mini.final.data%<>%convert(num(Bango_GSV:`Bango_Promo Agency Remuneration Fees & Commissions Consumer`))


bango.hsm.df <- dplyr::bind_rows(bango_subchannel_hsm_output,HSM.final.data)
bango.hsm.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.hsm.df%<>%arrange(Month_ymd)
bango.hsm.df <- bango.cej.df[order(bango.cej.df$Customer),]

bango.mini.df <- dplyr::bind_rows(bango_subchannel_mini_output,Mini.final.data)
bango.mini.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.mini.df%<>%arrange(Month_ymd)
bango.mini.df <- bango.mini.df[order(bango.mini.df$Customer),]

write_xlsx(bango.hsm.df,paste0(working_directory,"/bango.hsm.df.xlsx"))
write_xlsx(bango.mini.df,paste0(working_directory,"/bango.mini.df.xlsx"))


############################################################### Bango - Account #########################################

bango_account_output_path <- "C:/Users/goura/OneDrive/Desktop/Unilever_Projects/ID_Jarvis_Bango/Final_Consolidated_Output_Files/Final_Files/Bango-Account"
bango_account_output_files <- list.files(path=bango_account_output_path,pattern = "^GD ID MT(.*).xlsx|XLSX")
bango_account_alfamart_output <- read_xlsx(paste0(bango_account_output_path,"/",bango_account_output_files[1]),sheet = 1,guess_max = 100)
bango_account_careffour_output <- read_xlsx(paste0(bango_account_output_path,"/",bango_account_output_files[2]),sheet = 1,guess_max = 100)
bango_account_indomart_output <- read_xlsx(paste0(bango_account_output_path,"/",bango_account_output_files[3]),sheet = 1,guess_max = 100)

account.alfamart.data%<>%convert(num(Bango_GSV:`Bango_Sell out Volume (KGs)`))
account.carrefour.data%<>%convert(num(Bango_GSV:`Bango_Promo Agency Remuneration Fees & Commissions Consumer`))
account.indomart.data%<>%convert(num(Bango_GSV:`Bango_Sell out Volume (KGs)`))


bango.alfamart.df <- dplyr::bind_rows(bango_account_alfamart_output,account.alfamart.data)
bango.alfamart.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.alfamart.df%<>%arrange(Month_ymd)
bango.alfamart.df <- bango.alfamart.df[order(bango.alfamart.df$Customer),]

bango.carrefour.df <- dplyr::bind_rows(bango_account_careffour_output,account.carrefour.data)
bango.carrefour.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.carrefour.df%<>%arrange(Month_ymd)
bango.carrefour.df <- bango.carrefour.df[order(bango.carrefour.df$Customer),]

bango.indomart.df <- dplyr::bind_rows(bango_account_indomart_output,account.indomart.data)
bango.indomart.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.indomart.df%<>%arrange(Month_ymd)
bango.indomart.df <- bango.indomart.df[order(bango.indomart.df$Customer),]

write_xlsx(bango.alfamart.df,paste0(working_directory,"/bango.alfamart.df.xlsx"))
write_xlsx(bango.carrefour.df,paste0(working_directory,"/bango.carrefour.df.xlsx"))
write_xlsx(bango.indomart.df,paste0(working_directory,"/bango.indomart.df.xlsx"))

################################################################## Packgroup - Region ###########################################################

packgroup_region_dir_path <-  "C:/Users/goura/OneDrive/Desktop/Unilever_Projects/ID_Jarvis_Bango/Final_Consolidated_Output_Files/Final_Files/Pack group- Region"

packgroup_region_files <- list.files(path = packgroup_region_dir_path,recursive = T)

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


###############column names need to be renamed in the file as it's not matching ( work mode) #########################

packgroup.bmbp.cej%<>%convert(num(`Bango Manis Big Pouch_GSV`:`Bango Manis Big Pouch_Internal_penetration%`))
packgroup.bmbp.cej.df <- dplyr::bind_rows(packgroup_bmbp_cej_output,packgroup.bmbp.cej)
packgroup.bmbp.cej.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmbp.cej.df%<>%arrange(Month_ymd)
packgroup.bmbp.cej.df <- packgroup.bmbp.cej.df[order(packgroup.bmbp.cej.df$Customer),]


packgroup.bmbp.oi%<>%convert(num(`Bango Manis Big Pouch_GSV`:`Bango Manis Big Pouch_Internal_penetration%`))
packgroup.bmbp.oi.df <- dplyr::bind_rows(packgroup_bmbp_oi_output,packgroup.bmbp.oi)
packgroup.bmbp.oi.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmbp.oi.df%<>%arrange(Month_ymd)
packgroup.bmbp.oi.df <- packgroup.bmbp.oi.df[order(packgroup.bmbp.oi.df$Customer),]


packgroup.bmbp.sumatera%<>%convert(num(`Bango Manis Big Pouch_GSV`:`Bango Manis Big Pouch_Internal_penetration%`))

packgroup.bmbp.sumatera.df <- dplyr::bind_rows(packgroup_bmbp_sumatera_output,packgroup.bmbp.sumatera)
packgroup.bmbp.sumatera.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmbp.sumatera.df%<>%arrange(Month_ymd)
packgroup.bmbp.sumatera.df <- packgroup.bmbp.sumatera.df[order(packgroup.bmbp.sumatera.df$Customer),]

packgroup.bmbp.wj%<>%convert(num(`Bango Manis Big Pouch_GSV`:`Bango Manis Big Pouch_Internal_penetration%`))

packgroup.bmbp.wj.df <- dplyr::bind_rows(packgroup_bmbp_wj_output,packgroup.bmbp.wj)
packgroup.bmbp.wj.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmbp.wj.df%<>%arrange(Month_ymd)
packgroup.bmbp.wj.df <- packgroup.bmbp.wj.df[order(packgroup.bmbp.wj.df$Customer),]








end_time <- Sys.time()
script_time <- round(end_time-start_time,2)
print(script_time)
