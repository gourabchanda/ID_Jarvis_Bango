############################################ Final COnsolidation Script ##########################################################
start_time <- Sys.time()
working_directory <- setwd("C:/Users/goura/OneDrive/Desktop/Unilever_R_Projects/ID_Jarvis_Bango/ID_Jarvis_2.0.0")
#source(paste0(working_directory,"/data_preprocessing_script.r"))
#source(paste0(working_directory,"/ID_MT_Final.r"))
source(paste0(working_directory,"/ID_Jarvis_blending_script.r"))
library(ggplot2)

############################################# Bango-GT Model ( 1 output file)  ######################################################################
bango_gt_output_path <- "C:/Users/goura/OneDrive/Desktop/Unilever_Projects/ID_Jarvis_Bango/Final_Consolidated_Output_Files/Final_Files/Bango-GT"
bango_gt_output_files <- list.files(path = bango_gt_output_path ,pattern = "^GD ID GT -Bango(.*)xlsx|XLSX$")
bango_gt_output_file.df <- read_xlsx(paste0(bango_gt_output_path,"/",bango_gt_output_files),sheet = 1,guess_max = 100)
total.bango.gt.final%<>%convert(num(`Bango GT_GSV`:`Bango_I Trust %`))
bango.gt.df <- dplyr::bind_rows(bango_gt_output_file.df,total.bango.gt.final)
bango.gt.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.gt.df%<>%arrange(Month_ymd)
bango.gt.df <- bango.gt.df[order(bango.gt.df$Channel,decreasing=T),]

write_xlsx(bango.gt.df,paste0(working_directory,"/bango.gt.df.xlsx"))
############################################# Bango - MT Model ( 1 output file)  ###################################################################
bango_mt_output_path <- "C:/Users/goura/OneDrive/Desktop/Unilever_Projects/ID_Jarvis_Bango/Final_Consolidated_Output_Files/Final_Files/Bango-MT"
bango_mt_output_files <- list.files(path = bango_mt_output_path ,pattern = "^GD ID MT -Bango(.*)xlsx|XLSX$")
bango_mt_output_file.df <- read_xlsx(paste0(bango_mt_output_path,"/",bango_mt_output_files),sheet = 1,guess_max = 100)
total.bango.mt.final%<>%convert(num(Bango_GSV:`Bango_Promo Agency Remuneration Fees & Commissions Trade`))
bango.mt.df <- dplyr::bind_rows(bango_mt_output_file.df,total.bango.mt.final)
bango.mt.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.mt.df%<>%arrange(Month_ymd)
bango.mt.df <- bango.mt.df[order(bango.mt.df$Channel,decreasing=T),]
write_xlsx(bango.mt.df,paste0(working_directory,"/bango.mt.df.xlsx"))
############################################ Bango - National Model ( 1 output file) ###############################################################
bango_national_output_path <- "C:/Users/goura/OneDrive/Desktop/Unilever_Projects/ID_Jarvis_Bango/Final_Consolidated_Output_Files/Final_Files/Bango-National"
bango_national_output_files <- list.files(path = bango_national_output_path ,pattern = "^GD ID Indonesia-Bango(.*)xlsx|XLSX$")
bango_national_output_file.df <- read_xlsx(paste0(bango_national_output_path,"/",bango_national_output_files),sheet = 1,guess_max = 100)
total.final.bango%<>%convert(num(Bango_GSV:`Bango_Penetration%`))
bango.national.df <- dplyr::bind_rows(bango_national_output_file.df,total.final.bango)
bango.national.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.national.df%<>%arrange(Month_ymd)
bango.national.df <- bango.mt.df[order(bango.national.df$Market),]
write_xlsx(bango.national.df,paste0(working_directory,"/bango.national.df.xlsx"))
########################################### Bango - Region ( 4 output files) #########################################################################

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
bango.cej.df <- bango.cej.df[order(bango.cej.df$Customer,decreasing=T),]


bango.oi.df <- dplyr::bind_rows(bango_region_oi_output,bango.region.oi)
bango.oi.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.oi.df%<>%arrange(Month_ymd)
bango.oi.df <- bango.oi.df[order(bango.oi.df$Customer,decreasing=T),]


bango.sumatra.df <- dplyr::bind_rows(bango_region_sumatra_output,bango.region.sumatra)
bango.sumatra.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.sumatra.df%<>%arrange(Month_ymd)
bango.sumatra.df <- bango.cej.df[order(bango.sumatra.df$Customer,decreasing=T),]

# The output column has a mismatch ( Need to work on it)
bango.wj.df <- dplyr::bind_rows(bango_region_wj_output,bango.region.wj)
bango.wj.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.wj.df%<>%arrange(Month_ymd)
bango.wj.df <- bango.wj.df[order(bango.wj.df$Customer,decreasing=T),]


write_xlsx(bango.cej.df,paste0(working_directory,"/bango.cej.df"))
write_xlsx(bango.oi.df,paste0(working_directory,"/bango.oi.df.xlsx"))
write_xlsx(bango.sumatra.df,paste0(working_directory,"/bango.sumatra.df.xlsx"))
write_xlsx(bango.wj.df,paste0(working_directory,"/bango.wj.df.xlsx"))



############################################################### Bango - Sub Channel (2 output files) #########################################

bango_subchannel_output_path <- "C:/Users/goura/OneDrive/Desktop/Unilever_Projects/ID_Jarvis_Bango/Final_Consolidated_Output_Files/Final_Files/Bango-Sub channel"

bango_subchannel_output_files <- list.files(path=bango_subchannel_output_path,pattern="^GD ID MT(.*)xlsx|XLSX$")


bango_subchannel_hsm_output <-  read_xlsx(paste0(bango_subchannel_output_path,"/",bango_subchannel_output_files[1]),sheet = 1,guess_max = 100)
bango_subchannel_mini_output <- read_xlsx(paste0(bango_subchannel_output_path,"/",bango_subchannel_output_files[2]),sheet = 1,guess_max = 100)


HSM.final.data%<>%convert(num(Bango_GSV:`Bango_Promo Agency Remuneration Fees & Commissions Consumer`))
Mini.final.data%<>%convert(num(Bango_GSV:`Bango_Promo Agency Remuneration Fees & Commissions Consumer`))


bango.hsm.df <- dplyr::bind_rows(bango_subchannel_hsm_output,HSM.final.data)
bango.hsm.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.hsm.df%<>%arrange(Month_ymd)
bango.hsm.df <- bango.hsm.df[order(bango.hsm.df$Customer,decreasing=T),]

bango.mini.df <- dplyr::bind_rows(bango_subchannel_mini_output,Mini.final.data)
bango.mini.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.mini.df%<>%arrange(Month_ymd)
bango.mini.df <- bango.mini.df[order(bango.mini.df$Customer,decreasing=T),]

write_xlsx(bango.hsm.df,paste0(working_directory,"/bango.hsm.df.xlsx"))
write_xlsx(bango.mini.df,paste0(working_directory,"/bango.mini.df.xlsx"))


############################################################### Bango - Account(3 output files) #########################################

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
bango.alfamart.df <- bango.alfamart.df[order(bango.alfamart.df$Customer,decreasing=T),]

bango.carrefour.df <- dplyr::bind_rows(bango_account_careffour_output,account.carrefour.data)
bango.carrefour.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.carrefour.df%<>%arrange(Month_ymd)
bango.carrefour.df <- bango.carrefour.df[order(bango.carrefour.df$Customer,decreasing=T),]

bango.indomart.df <- dplyr::bind_rows(bango_account_indomart_output,account.indomart.data)
bango.indomart.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
bango.indomart.df%<>%arrange(Month_ymd)
bango.indomart.df <- bango.indomart.df[order(bango.indomart.df$Customer,decreasing=T),]

write_xlsx(bango.alfamart.df,paste0(working_directory,"/bango.alfamart.df.xlsx"))
write_xlsx(bango.carrefour.df,paste0(working_directory,"/bango.carrefour.df.xlsx"))
write_xlsx(bango.indomart.df,paste0(working_directory,"/bango.indomart.df.xlsx"))

################################################################## Packgroup - Region ( 24 output files - 4*6 =24 ) ###########################################################

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


#final consolidation for the Bango Manis Big Pouch : packgroup level
packgroup.bmbp.cej%<>%convert(num(`Bango Manis Big Pouch_GSV`:`Bango Manis Big Pouch_Internal_penetration%`))
packgroup.bmbp.cej.df <- dplyr::bind_rows(packgroup_bmbp_cej_output,packgroup.bmbp.cej)
packgroup.bmbp.cej.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmbp.cej.df%<>%arrange(Month_ymd)
packgroup.bmbp.cej.df <- packgroup.bmbp.cej.df[order(packgroup.bmbp.cej.df$Customer,decreasing=T),]


packgroup.bmbp.oi%<>%convert(num(`Bango Manis Big Pouch_GSV`:`Bango Manis Big Pouch_Internal_penetration%`))
packgroup.bmbp.oi.df <- dplyr::bind_rows(packgroup_bmbp_oi_output,packgroup.bmbp.oi)
packgroup.bmbp.oi.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmbp.oi.df%<>%arrange(Month_ymd)
packgroup.bmbp.oi.df <- packgroup.bmbp.oi.df[order(packgroup.bmbp.oi.df$Customer,decreasing=T),]


packgroup.bmbp.sumatera%<>%convert(num(`Bango Manis Big Pouch_GSV`:`Bango Manis Big Pouch_Internal_penetration%`))
packgroup.bmbp.sumatera.df <- dplyr::bind_rows(packgroup_bmbp_sumatera_output,packgroup.bmbp.sumatera)
packgroup.bmbp.sumatera.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmbp.sumatera.df%<>%arrange(Month_ymd)
packgroup.bmbp.sumatera.df <- packgroup.bmbp.sumatera.df[order(packgroup.bmbp.sumatera.df$Customer,decreasing=T),]

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

write_xlsx(packgroup.bmbp.cej.df,paste0(working_directory,"/packgroup.bmbp.cej.df.xlsx"))
write_xlsx(packgroup.bmbp.oi.df,paste0(working_directory,"/packgroup.bmbp.oi.df.xlsx"))
write_xlsx(packgroup.bmbp.sumatera.df,paste0(working_directory,"/packgroup.bmbp.sumatera.df.xlsx"))
write_xlsx(packgroup.bmbp.wj.df,paste0(working_directory,"/packgroup.bmbp.wj.df.xlsx"))



#final consolidation for the Bango Manis Big sachet : packgroup level
packgroup_bmbs_cej_output%<>%convert(num(`Bango Manis Big Sachet_GSV`:`Indofood_Have a more attractive packaging than other brands`))
packgroup.bmbs.cej%<>%convert(num(`Bango Manis Big Sachet_GSV`:`Bango Manis Big Sachet_Internal_penetration%`))
packgroup.bmbs.cej.df <- dplyr::bind_rows(packgroup_bmbs_cej_output,packgroup.bmbs.cej)
packgroup.bmbp.cej.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmbp.cej.df%<>%arrange(Month_ymd)
packgroup.bmbp.cej.df <- packgroup.bmbp.cej.df[order(packgroup.bmbp.cej.df$Customer,decreasing=T),]



packgroup.bmbs.oi%<>%convert(num(`Bango Manis Big Sachet_GSV`:`Bango Manis Big Sachet_Internal_penetration%`))
packgroup.bmbs.oi.df <- dplyr::bind_rows(packgroup_bmbs_oi_output,packgroup.bmbs.oi)
packgroup.bmbs.oi.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmbs.oi.df%<>%arrange(Month_ymd)
packgroup.bmbs.oi.df <- packgroup.bmbp.oi.df[order(packgroup.bmbp.oi.df$Customer,decreasing=T),]


packgroup.bmbs.sumatera%<>%convert(num(`Bango Manis Big Sachet_GSV`:`Bango Manis Big Sachet_Internal_penetration%`))
packgroup.bmbs.sumatera.df <- dplyr::bind_rows(packgroup_bmbp_sumatera_output,packgroup.bmbs.sumatera)
packgroup.bmbs.sumatera.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmbs.sumatera.df%<>%arrange(Month_ymd)
packgroup.bmbs.sumatera.df <- packgroup.bmbp.oi.df[order(packgroup.bmbp.oi.df$Customer,decreasing=T),]


packgroup.bmbs.wj%<>%convert(num(`Bango Manis Big Sachet_GSV`:`Bango Manis Big Sachet_Internal_penetration%`))
packgroup.bmbs.wj.df <- dplyr::bind_rows(packgroup_bmbs_wj_output,packgroup.bmbs.wj)
packgroup.bmbs.wj.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmbs.wj.df%<>%arrange(Month_ymd)
packgroup.bmbs.wj.df <- packgroup.bmbs.wj.df[order(packgroup.bmbs.wj.df$Customer,decreasing=T),]


write_xlsx(packgroup.bmbs.cej.df,paste0(working_directory,"/packgroup.bmbs.cej.dsf.xlsx"))
write_xlsx(packgroup.bmbs.oi.df,paste0(working_directory,"/packgroup.bmbs.oi.df.xlsx"))
write_xlsx(packgroup.bmbs.sumatera.df,paste0(working_directory,"/packgroup.bmbs.sumatera.df.xlsx"))
write_xlsx(packgroup.bmbs.wj.df,paste0(working_directory,"/packgroup.bmbs.wj.df.xlsx"))

#final consolidation for the Bango Manis Medium Pouch : packgroup level

packgroup.bmmp.cej%<>%convert(num(`Bango Manis Medium Pouch_GSV`:`Bango Manis Medium Pouch_Internal_penetration%`))
packgroup.bmmp.cej.df <- dplyr::bind_rows(packgroup_bmmp_cej_output,packgroup.bmmp.cej)
packgroup.bmmp.cej.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmmp.cej.df%<>%arrange(Month_ymd)
packgroup.bmmp.cej.df <- packgroup.bmmp.cej.df[order(packgroup.bmmp.cej.df$Customer,decreasing=T),]


packgroup.bmmp.oi%<>%convert(num(`Bango Manis Medium Pouch_GSV`:`Bango Manis Medium Pouch_Internal_penetration%`))
packgroup.bmmp.oi.df <- dplyr::bind_rows(packgroup_bmmp_oi_output,packgroup.bmmp.oi)
packgroup.bmmp.oi.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmmp.oi.df%<>%arrange(Month_ymd)
packgroup.bmmp.oi.df <- packgroup.bmmp.oi.df[order(packgroup.bmmp.oi.df$Customer,decreasing=T),]


packgroup.bmmp.sumatera%<>%convert(num(`Bango Manis Medium Pouch_GSV`:`Bango Manis Medium Pouch_Internal_penetration%`))
packgroup.bmmp.sumatera.df <- dplyr::bind_rows(packgroup_bmmp_sumatera_output,packgroup.bmmp.sumatera)
packgroup.bmmp.sumatera.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmmp.sumatera.df%<>%arrange(Month_ymd)
packgroup.bmmp.sumatera.df <- packgroup.bmmp.sumatera.df[order(packgroup.bmmp.sumatera.df$Customer,decreasing=T),]

packgroup.bmmp.wj%<>%convert(num(`Bango Manis Medium Pouch_GSV`:`Bango Manis Medium Pouch_Internal_penetration%`))
packgroup.bmmp.wj.df <- dplyr::bind_rows(packgroup_bmmp_wj_output,packgroup.bmmp.wj)
packgroup.bmmp.wj.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmmp.wj.df%<>%arrange(Month_ymd)
packgroup.bmmp.wj.df <- packgroup.bmmp.wj.df[order(packgroup.bmmp.wj.df$Customer,decreasing=T),]

write_xlsx(packgroup.bmmp.cej.df,paste0(working_directory,"/packgroup.bmmp.cej.xlsx"))
write_xlsx(packgroup.bmmp.oi.df,paste0(working_directory,"/packgroup.bmmp.oi.df.xlsx"))
write_xlsx(packgroup.bmmp.sumatera.df,paste0(working_directory,"/packgroup.bmmp.sumatera.df.xlsx"))
write_xlsx(packgroup.bmmp.wj.df,paste0(working_directory,"/packgroup.bmmp.wj.df.xlsx"))

#final consolidation for the Bango Manis Medium Bottle : packgroup level

packgroup.bmmb.cej%<>%convert(num(`Bango Manis Medium Bottle_GSV`:`Bango Manis Medium Bottle_Internal_penetration%`))
packgroup.bmmb.cej.df <- dplyr::bind_rows(packgroup_bmmb_cej_output,packgroup.bmmb.cej)
packgroup.bmmb.cej.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmmb.cej.df%<>%arrange(Month_ymd)
packgroup.bmmb.cej.df <- packgroup.bmmb.cej.df[order(packgroup.bmmb.cej.df$Customer,decreasing=T),]


packgroup.bmmb.oi%<>%convert(num(`Bango Manis Medium Bottle_GSV`:`Bango Manis Medium Bottle_Internal_penetration%`))
packgroup.bmmb.oi.df <- dplyr::bind_rows(packgroup_bmmb_oi_output,packgroup.bmmb.oi)
packgroup.bmmb.oi.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmmb.oi.df%<>%arrange(Month_ymd)
packgroup.bmmb.oi.df <- packgroup.bmmb.oi.df[order(packgroup.bmmb.oi.df$Customer,decreasing=T),]


packgroup.bmmb.sumatera%<>%convert(num(`Bango Manis Medium Bottle_GSV`:`Bango Manis Medium Bottle_Internal_penetration%`))
packgroup.bmmb.sumatera.df <- dplyr::bind_rows(packgroup_bmmb_sumatera_output,packgroup.bmmb.sumatera)
packgroup.bmmb.sumatera.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmmb.sumatera.df%<>%arrange(Month_ymd)
packgroup.bmmb.sumatera.df <- packgroup.bmmb.sumatera.df[order(packgroup.bmmb.sumatera.df$Customer,decreasing=T),]

packgroup.bmmb.wj%<>%convert(num(`Bango Manis Medium Bottle_GSV`:`Bango Manis Medium Bottle_Internal_penetration%`))
packgroup.bmmb.wj.df <- dplyr::bind_rows(packgroup_bmmb_wj_output,packgroup.bmmb.wj)
packgroup.bmmb.wj.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmmb.wj.df%<>%arrange(Month_ymd)
packgroup.bmmb.wj.df <- packgroup.bmmb.wj.df[order(packgroup.bmmb.wj.df$Customer,decreasing=T),]

write_xlsx(packgroup.bmmb.cej.df,paste0(working_directory,"/packgroup.bmmb.cej.xlsx"))
write_xlsx(packgroup.bmmb.oi.df,paste0(working_directory,"/packgroup.bmmb.oi.df.xlsx"))
write_xlsx(packgroup.bmmb.sumatera.df,paste0(working_directory,"/packgroup.bmmb.sumatera.df.xlsx"))
write_xlsx(packgroup.bmmb.wj.df,paste0(working_directory,"/packgroup.bmmb.wj.df.xlsx"))

#final consolidation for the Bango Manis small Bottle : packgroup level

packgroup.bmsb.cej%<>%convert(num(`Bango Manis Small Bottle_GSV`:`Bango Manis Small Bottle_Internal_penetration%`))
packgroup.bmsb.cej.df <- dplyr::bind_rows(packgroup_bmsb_cej_output,packgroup.bmsb.cej)
packgroup.bmsb.cej.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmsb.cej.df%<>%arrange(Month_ymd)
packgroup.bmsb.cej.df <- packgroup.bmsb.cej.df[order(packgroup.bmsb.cej.df$Customer,decreasing=T),]


packgroup.bmsb.oi%<>%convert(num(`Bango Manis Small Bottle_GSV`:`Bango Manis Small Bottle_Internal_penetration%`))
packgroup.bmsb.oi.df <- dplyr::bind_rows(packgroup_bmsb_oi_output,packgroup.bmsb.oi)
packgroup.bmsb.oi.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmsb.oi.df%<>%arrange(Month_ymd)
packgroup.bmsb.oi.df <- packgroup.bmsb.oi.df[order(packgroup.bmsb.oi.df$Customer,decreasing=T),]


packgroup.bmsb.sumatera%<>%convert(num(`Bango Manis Small Bottle_GSV`:`Bango Manis Small Bottle_Internal_penetration%`))
packgroup.bmsb.sumatera.df <- dplyr::bind_rows(packgroup_bmsb_sumatera_output,packgroup.bmsb.sumatera)
packgroup.bmsb.sumatera.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmsb.sumatera.df%<>%arrange(Month_ymd)
packgroup.bmsb.sumatera.df <- packgroup.bmsb.sumatera.df[order(packgroup.bmsb.sumatera.df$Customer,decreasing=T),]

packgroup.bmsb.wj%<>%convert(num(`Bango Manis Small Bottle_GSV`:`Bango Manis Small Bottle_Internal_penetration%`))
packgroup.bmsb.wj.df <- dplyr::bind_rows(packgroup_bmsb_wj_output,packgroup.bmsb.wj)
packgroup.bmsb.wj.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmsb.wj.df%<>%arrange(Month_ymd)
packgroup.bmsb.wj.df <- packgroup.bmsb.wj.df[order(packgroup.bmsb.wj.df$Customer,decreasing=T),]

write_xlsx(packgroup.bmsb.cej.df,paste0(working_directory,"/packgroup.bmsb.cej.xlsx"))
write_xlsx(packgroup.bmsb.oi.df,paste0(working_directory,"/packgroup.bmsb.oi.df.xlsx"))
write_xlsx(packgroup.bmsb.sumatera.df,paste0(working_directory,"/packgroup.bmsb.sumatera.df.xlsx"))
write_xlsx(packgroup.bmsb.wj.df,paste0(working_directory,"/packgroup.bmsb.wj.df.xlsx"))

#final consolidation for the Bango Manis small Pouch : packgroup level

packgroup.bmsp.cej%<>%convert(num(`Bango Manis Small Pouch_GSV`:`Bango Manis Small Pouch_Internal_penetration%`))
packgroup.bmsp.cej.df <- dplyr::bind_rows(packgroup_bmsp_cej_output,packgroup.bmsp.cej)
packgroup.bmsp.cej.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmsp.cej.df%<>%arrange(Month_ymd)
packgroup.bmsp.cej.df <- packgroup.bmsp.cej.df[order(packgroup.bmsp.cej.df$Customer,decreasing=T),]


packgroup.bmsp.oi%<>%convert(num(`Bango Manis Small Pouch_GSV`:`Bango Manis Small Pouch_Internal_penetration%`))
packgroup.bmsp.oi.df <- dplyr::bind_rows(packgroup_bmsp_oi_output,packgroup.bmsp.oi)
packgroup.bmsp.oi.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmsp.oi.df%<>%arrange(Month_ymd)
packgroup.bmsp.oi.df <- packgroup.bmsp.oi.df[order(packgroup.bmsp.oi.df$Customer,decreasing=T),]


packgroup.bmsp.sumatera%<>%convert(num(`Bango Manis Small Pouch_GSV`:`Bango Manis Small Pouch_Internal_penetration%`))
packgroup.bmsp.sumatera.df <- dplyr::bind_rows(packgroup_bmsp_sumatera_output,packgroup.bmsp.sumatera)
packgroup.bmsp.sumatera.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmsp.sumatera.df%<>%arrange(Month_ymd)
packgroup.bmsp.sumatera.df <- packgroup.bmsp.sumatera.df[order(packgroup.bmsp.sumatera.df$Customer,decreasing=T),]

packgroup.bmsp.wj%<>%convert(num(`Bango Manis Small Pouch_GSV`:`Bango Manis Small Pouch_Internal_penetration%`))
packgroup.bmsp.wj.df <- dplyr::bind_rows(packgroup_bmsp_wj_output,packgroup.bmsp.wj)
packgroup.bmsp.wj.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
packgroup.bmsp.wj.df%<>%arrange(Month_ymd)
packgroup.bmsp.wj.df <- packgroup.bmsp.wj.df[order(packgroup.bmsp.wj.df$Customer,decreasing=T),]

write_xlsx(packgroup.bmsp.cej.df,paste0(working_directory,"/packgroup.bmsp.cej.xlsx"))
write_xlsx(packgroup.bmsp.oi.df,paste0(working_directory,"/packgroup.bmsp.oi.df.xlsx"))
write_xlsx(packgroup.bmsp.sumatera.df,paste0(working_directory,"/packgroup.bmsp.sumatera.df.xlsx"))
write_xlsx(packgroup.bmsp.wj.df,paste0(working_directory,"/packgroup.bmsp.wj.df.xlsx"))


##################################################### Basepack - Account ( 27 output files - 3*9=27) ##############################################

#There are 27 output files generated by the Basepack Account where each Account i.e. indomart , Indomart and Carrefour has 9 basepack files which are as follows
#BANGO KECAP MANIS 275ML:bkm275 ,BANGO KECAP MANIS 620ML:bkm620 ,BANGO KECAP MANIS 220ML:bkm220 ,BANGO KECAP MANIS PEDAS 220ML:bkmp220 ,BANGO KECAP MANIS PEDAS 135ML: bkmp135,
#BANGO SOY/OYSTER/FISH SAUCE MANIS 60ML:bsofs60,BANGO SOY/OYSTR/FISH SUCE MNIS SCHT 30ML:bsofms30,BANGO KECAP MANIS 135ML:bkm135 ,BANGO SOY/OYSTER/FISH KECAP MANIS 580ML:bsof580

basepack_account_path <- "C:/Users/goura/OneDrive/Desktop/Unilever_Projects/ID_Jarvis_Bango/Final_Consolidated_Output_Files/Final_Files/Basepack-Account"
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

####################################################### Basepack - Account ( Alfamart) work in progress #################################################################

alfamart.bkm275%<>%convert(num(GSV:`Sell out Volume (KGs)`))
alfamart.bkm275.df <- bind_rows(alfamart_bkm275_output,alfamart.bkm275)
alfamart.bkm275.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
alfamart.bkm275.df%<>%arrange(Month_ymd)
alfamart.bkm275.df <- alfamart.bkm275.df[order(alfamart.bkm275.df$Customer,decreasing=T),]

alfamart.bkm620%<>%convert(num(GSV:`Sell out Volume (KGs)`))
alfamart.bkm620.df <- bind_rows(alfamart_bkm620_output,alfamart.bkm620)
alfamart.bkm620.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
alfamart.bkm620.df%<>%arrange(Month_ymd)
alfamart.bkm620.df <- alfamart.bkm620.df[order(alfamart.bkm620.df$Customer,decreasing=T),]


alfamart.bkm220%<>%convert(num(GSV:`Sell out Volume (KGs)`))
alfamart.bkm220.df <- bind_rows(alfamart_bkm220_output,alfamart.bkm220)
alfamart.bkm220.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
alfamart.bkm220.df%<>%arrange(Month_ymd)
alfamart.bkm220.df <- alfamart.bkm220.df[order(alfamart.bkm220.df$Customer,decreasing=T),]


alfamart.bkmp220%<>%convert(num(GSV:`Sell out Volume (KGs)`))
alfamart.bkmp220.df <- bind_rows(alfamart_bkmp220_output,alfamart.bkmp220)
alfamart.bkmp220.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
alfamart.bkmp220.df%<>%arrange(Month_ymd)
alfamart.bkmp220.df <- alfamart.bkmp220.df[order(alfamart.bkmp220.df$Customer,decreasing=T),]



alfamart.bkmp135%<>%convert(num(GSV:`Sell out Volume (KGs)`))
alfamart.bkmp135.df <- bind_rows(alfamart_bkmp135_output,alfamart.bkmp135)
alfamart.bkmp135.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
alfamart.bkmp135.df%<>%arrange(Month_ymd)
alfamart.bkmp135.df <- alfamart.bkmp135.df[order(alfamart.bkmp135.df$Customer,decreasing=T),]

alfamart.bsofs60%<>%convert(num(GSV:`Sell out Volume (KGs)`))
alfamart.bsofs60.df <- bind_rows(alfamart_bsofs60_output,alfamart.bsofs60)
alfamart.bsofs60.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
alfamart.bsofs60.df%<>%arrange(Month_ymd)
alfamart.bsofs60.df <- alfamart.bsofs60.df[order(alfamart.bsofs60.df$Customer,decreasing=T),]


alfamart.bsofms30%<>%convert(num(GSV:`Sell out Volume (KGs)`))
alfamart.bsofms30.df <- bind_rows(alfamart_bsofms30_output,alfamart.bsofms30)
alfamart.bsofms30.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
alfamart.bsofms30.df%<>%arrange(Month_ymd)
alfamart.bsofms30.df <- alfamart.bsofms30.df[order(alfamart.bsofms30.df$Customer,decreasing=T),]


alfamart.bkm135%<>%convert(num(GSV:`Sell out Volume (KGs)`))
alfamart.bkm135.df <- bind_rows(alfamart_bkm135_output,alfamart.bkm135)
alfamart.bkm135.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
alfamart.bkm135.df%<>%arrange(Month_ymd)
alfamart.bkm135.df <- alfamart.bkm135.df[order(alfamart.bkm135.df$Customer,decreasing=T),]


alfamart.bsof580%<>%convert(num(GSV:`Sell out Volume (KGs)`))
alfamart.bsof580.df <- bind_rows(alfamart_bsof580_output,alfamart.bsof580)
alfamart.bsof580.df%<>%mutate(Month_ymd=lubridate::dmy(paste("01-",Month,sep = "")))
alfamart.bsof580.df%<>%arrange(Month_ymd)
alfamart.bsof580.df <- alfamart.bsof580.df[order(alfamart.bsof580.df$Customer,decreasing=T),]













end_time <- Sys.time()
script_time <- round(end_time-start_time,2)
print(script_time)
