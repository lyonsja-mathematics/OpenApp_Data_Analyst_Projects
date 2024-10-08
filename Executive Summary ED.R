library(rJava)
library(xlsxjars)
library(xlsx)
library(extrafont)

import_trolleygar_data = function(){
  
  trolleygar_file_name='C:/Users/jack.lyons/Downloads/TrolleyGAR.csv'
  reference_file='~/provider_codes_20210625.csv'
  
  d = read.csv(
    trolleygar_file_name,
    stringsAsFactors = F
    )
  
  names(d) = c('id', 'date', 't2', 't20', 't8', 'chart_date')
  
  d = d[, c('id', 'date', 't8')]
  
  d = merge(
    d, 
    read.csv(
      reference_file,
      stringsAsFactors = F
    ), 
    by.x = 'id', 
    by.y = 'biu_hospital_id', 
    all.x = T
  )
  
  tf = '%d/%m/%Y'
  
  d$target_month = as.numeric(
    format(
      as.Date(
        d$date, 
        format = tf
      ), 
      '%m'
    )
  )
  
  d$target_year = as.numeric(
    format(
      as.Date(
        d$date, 
        format = tf
      ), 
      '%Y'
    )
  )
  
  names(d)[names(d)=='hospital_group']='group'
  
  d = d[duplicated(d[,names(d)])==FALSE,]
  
  return(d)
}

import_biu_data = function(){
  
  source("C:/Users/jack.lyons/Downloads/Time.R")
  
  fls = list.files(
    path = "C:/Users/jack.lyons/Downloads/PET_Data_01-MMM-YYYY_to_DD-MMM-YYYY/",
    pattern = ".csv",
    full.names = T
  )
  
  df = data.frame(
    provider_code=NA,
    mrn=NA,
    reg = NA,
    age = NA,
    triage = NA,
    edclinician = NA,
    depart = NA,
    destination = NA
  )
  
  for(i in fls){
    
    df1 = read.csv(
      i,
      stringsAsFactors = F,
      header = F
    )
    
    names(df1)=c('provider_code',"mrn",'reg','age',"triage","edclinician",'depart','destination')
    
    for(c in c("reg","depart")){
      df1[, c] = standard_format(df1[, c])
    }
    
    if(nrow(subset(df1, is.na(reg))) != 0){
      View(subset(df1, is.na(reg)), i)
      cat(i, sep = "\n")
    }
    
    df = rbind(
      subset(
        df,
        complete.cases(mrn)
      ),
      df1
    )
  }
  
  
  df = df[
    duplicated(
      df[, 
         names(df)
         ]
    ) == F, 
    ]
  
  names(df) = c('id', 'mrn', 'reg', 'age', 'triage', 'clinician', 'depart', 'destination')
  
  providers = c(904,923,908,501,724,941,402,913,
                1049,992,403,209,910,940,300,108,
                202,201,203,607,601,102,600,605,
                500,800,802,809,726)
  
  df = subset(
    df,
    id %in% providers
  )
  
  # df = subset(
  #   df, 
  #   !(id %in% c(101,938,400))
  # )
  
  df = merge(
    df, 
    read.csv(
      "C:/Users/jack.lyons/Documents/provider_codes.csv",
      stringsAsFactors = F
    ), 
    by.x = 'id', 
    by.y = 'biu_hospital_id', 
    all.x = T
  )
  
  tf = "%Y-%m-%d %H:%M:%S"
  
  df$reg_month = as.numeric(
    format(
      df$reg, 
      '%m'
    )
  )
  
  df$reg_year = as.numeric(
    format(
      df$reg, 
      '%Y'
    )
  )
  
  df$depart_month = as.numeric(
    format(
      df$depart, 
      '%m'
    )
  )
  
  df$depart_year = as.numeric(
    format(
      df$depart, 
      '%Y'
    )
  )
  
  df$attends = 1
  
  df$departs = 1
  
  df$adm = ifelse(
    df$destination %in% c(
      "Admission Lounge",
      "Admitted to Ward",
      "Discharged to CDU",
      "Referred to AMU"
    ), 
    1, 
    0
  )
  
  df[
    is.na(df$age), 
    'age'
    ] = 200
  
  df[
    df$age > 110, 
    'age'
    ] = 200
  
  df$age_group = ifelse(
    df$age %in% 0:15, 
    'Paeds 0-15', 
    ifelse(
      df$age %in% 16:74, 
      'Adults 16-74', 
      ifelse(
        df$age %in% 75:110, 
        'Adults 75+', 
        'Unknown'
      )
    )
  )
  
  df$count = 1
  
  df$paed = ifelse(
    df$age %in% 0:15, 
    1, 
    0
  )
  
  df$over_75 = ifelse(
    df$age %in% 75:110, 
    1, 
    0
  )
  
  df$paed_adm = ifelse(
    df$age %in% 0:15 & df$adm == 1, 
    1, 
    0
  )
  
  df$over_75_adm = ifelse(
    df$age %in% 75:110 & df$adm == 1, 
    1, 
    0
  )
  
  for(
    i in c('reg', 'depart')
  ){
    df[, i] = as.POSIXct(
      as.character(df[, i]),
      tz = "GMT",
      format = tf
    )
  }
  
  df$pet = as.numeric(
    difftime(
      as.POSIXct(
        as.character(
          df$depart
          ),
        tz = "GMT"
      ),
      as.POSIXct(
        as.character(
          df$reg
          ),
        tz = "GMT"
      ),
      units = 'mins'
    )
  )
  
  df[
    is.na(df$pet) == F & df$pet < 0, 
    'pet'
    ] = NA
  
  df$pet_under_6 = ifelse(
    df$pet < 360, 
    1, 
    0
  )
  
  df$pet_under_9 = ifelse(
    df$pet < 540, 
    1,
    0
  )
  
  df$pet_under_6_adm = ifelse(
    df$pet < 360 & df$adm == 1, 
    1, 
    0
  )
  
  df$pet_under_9_adm = ifelse(
    df$pet < 540 & df$adm == 1, 
    1, 
    0
  )
  
  df$pet_under_6_75_plus = ifelse(
    df$pet < 360 & df$age %in% 75:110, 
    1, 
    0
  )
  
  df$pet_under_9_75_plus = ifelse(
    df$pet < 540 & df$age %in% 75:110, 
    1, 
    0
  )
  
  df$pet_under_6_75_plus_adm = ifelse(
    df$pet < 360 & df$age %in% 75:110 & df$adm == 1, 
    1, 
    0
  )
  
  df$pet_under_9_75_plus_adm = ifelse(
    df$pet < 540 & df$age %in% 75:110 & df$adm == 1, 
    1, 
    0
  )
  
  df$pet_under_6_paed = ifelse(
    df$pet < 360 & df$age %in% 0:15, 
    1, 
    0
  )
  
  df$pet_under_9_paed = ifelse(
    df$pet < 540 & df$age %in% 0:15, 
    1, 
    0
  )
  
  df$pet_under_6_paed_adm = ifelse(
    df$pet < 360 & df$age %in% 0:15 & df$adm == 1, 
    1, 
    0
  )
  
  df$pet_under_9_paed_adm = ifelse(
    df$pet < 540 & df$age %in% 0:15 & df$adm == 1, 
    1, 
    0
  )
  
  df$pet_over_24 = ifelse(
    df$pet >= 1440, 
    1, 
    0
  )
  
  df$pet_under_24 = ifelse(
    df$pet < 1440, 
    1, 
    0
  )
  
  df$pet_bin = floor(df$pet/60)
  
  df$depart_time = as.numeric(
    format(
      df$depart, 
      '%H'
    )
  )
  
  df$reg_time = as.numeric(
    format(
      df$reg, 
      '%H'
    )
  )
  
  names(df)[names(df)=='hospital_group']='group'
  
  
  return(df)
}

data = import_biu_data()
year_of_interest=max(data$reg_year)
month_of_interest=max(subset(data,reg_year==year_of_interest)$reg_month)
trolleygar_data = subset(import_trolleygar_data(), group != 'CHI')

run_ed_exec_summary = function(hse_group="NAT"){
  
  if(hse_group == "NAT"){
    hse_group = c("CHI", "DML", "IEHG", "RCSI", "SAOLTA", "SSW", "UL")
  }
  new_attends<-nrow(subset(data, reg_month ==month_of_interest & reg_year == year_of_interest & group %in% hse_group))
  new_attends_ytd<-nrow(subset(data, reg_month %in% c(1:month_of_interest) & reg_year == year_of_interest & group %in% hse_group))
  
  new_attends_national<-nrow(subset(data, reg_month ==month_of_interest & reg_year == year_of_interest))
  
  new_attends_ytd_national<-nrow(subset(data, reg_month %in% c(1:month_of_interest) & reg_year == year_of_interest))
  
  new_6hour_pet<-round(100*(nrow(subset(data, depart_month ==month_of_interest & depart_year == year_of_interest & is.na(pet) == F & pet_under_6 == 1 & group %in% hse_group)) / nrow(subset(data, depart_month ==month_of_interest & depart_year == year_of_interest & is.na(pet) == F & group %in% hse_group))),digits=1)
  
  new_6hour_pet_national<-round(100*(nrow(subset(data,depart_month == month_of_interest & depart_year == year_of_interest & pet_under_6 == 1 & is.na(pet) == F)) / nrow(subset(data, depart_month == month_of_interest & depart_year == year_of_interest & is.na(pet) == F))), digits=1)
  
  new_9hour_pet<-round(100*(nrow(subset(data, depart_month == month_of_interest & depart_year == year_of_interest & is.na(pet) == F & pet_under_9 == 1 & group %in% hse_group)) /nrow(subset(data, depart_month == month_of_interest & depart_year == year_of_interest & is.na(pet) == F & group %in% hse_group))),digits=1)
  
  new_9hour_pet_national<-round(100*(nrow(subset(data, depart_month == month_of_interest & depart_year == year_of_interest & is.na(pet) == F & pet_under_9 == 1)) /nrow(subset(data, depart_month == month_of_interest & depart_year == year_of_interest & is.na(pet) == F))), digits=1)
  
  new_pet_gt_24 = nrow(
    subset(
      data, 
      depart_month == month_of_interest & depart_year == year_of_interest & is.na(pet) == F & pet_over_24 == 1 & group %in% hse_group
    )
  )
  
  ## Value: national pet greater than 24 hours
  new_pet_gt_24_national<-nrow(subset(data, depart_month ==month_of_interest & depart_year == year_of_interest & is.na(pet) == F & pet_over_24 == 1))
  
  new_departs = nrow(subset(data, depart_month == month_of_interest & is.na(pet) == F & group %in% hse_group & depart_year == year_of_interest))
  
  new_departs_national = nrow(subset(data, depart_month == month_of_interest & depart_year == year_of_interest & is.na(pet) == F))
  
  
  ## Value: percentage of attends with PET greater than 24 hours
  new_pet_gt_24_percentage<-round(100*new_pet_gt_24/new_departs,1)
  
  
  ## Value: percentage of national attends with PETgreater than 24 hours
  new_pet_gt_24_percentage_national<-round(100*new_pet_gt_24_national/new_departs_national,1)
  
  new_8am_trolley<-sum(subset(trolleygar_data,target_month==month_of_interest & target_year==year_of_interest & group %in% hse_group)$t8)
  
  
  ## Value: 8 am trollies (national)
  new_8am_trolley_national<-sum(subset(trolleygar_data, target_year == year_of_interest  & target_month == month_of_interest)$t8)
  
  
  ## Value: 8 am trollies YTD
  new_8am_trolley_ytd<-sum(subset(trolleygar_data, target_month %in% 1:month_of_interest & target_year == year_of_interest & group %in% hse_group)$t8)
  
  
  ## 8 am trollies YTD (national)
  new_8am_trolley_ytd_national<-sum(subset(trolleygar_data,target_year==year_of_interest  & target_month %in% c(1:month_of_interest ))$t8)
  
  
  ## Calculate values for the previous year ##
  ## Monthly Attends
  prev_attends<-nrow(subset(data, reg_month ==month_of_interest & reg_year == year_of_interest-1 & group %in% hse_group))
  
  ## Value: Attends YTD
  prev_attends_ytd<-nrow(subset(data, reg_month %in% c(1:month_of_interest) & reg_year == year_of_interest-1 & group %in% hse_group))
  
  ## national monthly attends
  prev_attends_national<-nrow(subset(data, reg_month ==month_of_interest & reg_year == year_of_interest-1))
  
  ## national monthly attends
  prev_attends_ytd_national<-nrow(subset(data, reg_month %in% c(1:month_of_interest) & reg_year == year_of_interest-1))
  
  ## Value: National 6 hour PET
  prev_6hour_pet_national<-ifelse(month_of_interest != 1, round(100*(nrow(subset(data, depart_month ==month_of_interest-1 & depart_year == year_of_interest & is.na(pet) == F & pet_under_6 == 1)) /nrow(subset(data, depart_month ==month_of_interest-1 & depart_year == year_of_interest & is.na(pet) == F))),digits=1), round(100*(nrow(subset(data, depart_month == 12 & depart_year == year_of_interest-1 & is.na(pet) == F & pet_under_6 == 1)) /nrow(subset(data, depart_month == 12 & depart_year == year_of_interest-1 & is.na(pet) == F))),digits=1))
  
  ## Value: national 9 hour pet
  prev_9hour_pet_national<-ifelse(month_of_interest != 1, round(100*(nrow(subset(data, depart_month ==month_of_interest-1 & depart_year == year_of_interest & is.na(pet) == F & pet_under_9 == 1)) /nrow(subset(data, depart_month ==month_of_interest-1 & depart_year == year_of_interest & is.na(pet) == F))),digits=1), round(100*(nrow(subset(data, depart_month ==12 & depart_year == year_of_interest-1 & is.na(pet) == F & pet_under_9 == 1)) /nrow(subset(data, depart_month == 12 & depart_year == year_of_interest-1 & is.na(pet) == F))),digits=1))
  
  prev_8am_trolley<-sum(subset(trolleygar_data,target_month==month_of_interest & target_year==year_of_interest-1 & group %in% hse_group & is.na(t8)==F)$t8)
  
  
  ## Value: 8 am trollies (national)
  prev_8am_trolley_national<-sum(subset(trolleygar_data,target_year==year_of_interest-1  & target_month==month_of_interest & is.na(t8)==F)$t8)
  
  
  ## Value: 8 am trollies YTD
  prev_8am_trolley_ytd<-sum(subset(trolleygar_data,target_month %in% c(1:month_of_interest) & target_year==year_of_interest-1 & group %in% hse_group & is.na(t8)==F)$t8)
  
  
  ## 8 am trollies YTD (national)
  prev_8am_trolley_ytd_national<-sum(subset(trolleygar_data, target_year==year_of_interest-1  & target_month %in% c(1:month_of_interest) & is.na(t8)==F)$t8)
  
  
  
  ## Calculate the change from previous ##
  ## Change in monthly attends ##
  change_attends<-round((new_attends-prev_attends)*100/prev_attends,1)
  
  ## Change in attends YTD
  change_attends_ytd<-round((new_attends_ytd-prev_attends_ytd)*100/prev_attends_ytd,1)
  
  ## Change in national attends
  change_attends_national<-round((new_attends_national-prev_attends_national)*100/prev_attends_national,1)
  
  ## Change in national attends YTD
  change_attends_ytd_national<-round((new_attends_ytd_national-prev_attends_ytd_national)*100/prev_attends_ytd_national,1)
  
  ## Change in 8am trollies
  change_8am_trolley<-ifelse(prev_8am_trolley == 0, 0, round((new_8am_trolley-prev_8am_trolley)*100/prev_8am_trolley,1))
  
  
  ## Change in 8am trollies
  change_8am_trolley_national<-round((new_8am_trolley_national-prev_8am_trolley_national)*100/ prev_8am_trolley_national,1)
  
  
  ## Change in YTD 8am trollies
  change_8am_trolley_ytd<-ifelse(prev_8am_trolley_ytd == 0, 0, round((new_8am_trolley_ytd-prev_8am_trolley_ytd)*100/prev_8am_trolley_ytd,1))
  
  
  ## Change in YTD 8am trollies
  change_8am_trolley_ytd_national<-round((new_8am_trolley_ytd_national-prev_8am_trolley_ytd_national)*100/ prev_8am_trolley_ytd_national,1)
  
  
  ## Comments ##
  ## Comment: monthly attends
  
  if(change_attends>=0){sign=paste("+")} else {sign=""}
  comment_attends<-paste(sign, change_attends,"% compared to ", format(as.Date(paste(year_of_interest-1, month_of_interest, 1, sep = '-')), '%B %Y'))
  
  ## comment: Monthly attends (YTD)
  comment_attends_ytd<-paste(if(change_attends_ytd>=0){ "+" } else {""},change_attends_ytd,"% YTD")
  
  
  ## Comment: monthly attends (national)
  if(change_attends_national>=0){sign=paste("+")} else {sign=""}
  comment_attends_national<-paste(sign, change_attends_national,"% compared to ", format(as.Date(paste(year_of_interest-1, month_of_interest, 1, sep = '-')), '%B %Y'))
  
  ## Comment: Monthly attends (YTD)
  comment_attends_national_ytd<-paste(if(change_attends_ytd_national>=0){ "+" } else {""},change_attends_ytd_national,"% YTD")
  
  
  ## Comment: 6 hour pet
  comment_6hour_pet<-if(length(hse_group) > 1){
    paste("Prior Month: ",prev_6hour_pet_national,"%")
  }  else {
    paste("National 6-hour PET: ",new_6hour_pet_national," %")
  }
  
  
  ## Comment: 9 hour PET
  comment_9hour_pet<-if(length(hse_group) > 1){
    paste("Prior Month: ",prev_9hour_pet_national,"%")
  }  else {
    paste("National 9-hour PET: ",new_9hour_pet_national," %")
  }
  
  
  ## Comment: PET greater than 24 hours
  comment_pet_gt_24<-if(length(hse_group) > 1){
    paste(100-new_pet_gt_24_percentage,"% of all ED departures < 24 hours from registration.")
  } else  if (hse_group == "CHI"){
    paste("National Average:", new_pet_gt_24_percentage_national," %")
  } else {
    paste("Equivalent to ",new_pet_gt_24_percentage,"% of all ED departures.",
          "National average: ",new_pet_gt_24_percentage_national," %")
  }
  
  
  ## Comment: 8am trollies
  sign=if(change_8am_trolley>=0){paste("+")} else {paste("")}
  comment_8am_trolley<-paste(sign,change_8am_trolley,"% compared to", format(as.Date(paste(year_of_interest-1, month_of_interest, 1, sep = '-')), '%B %Y'))
  
  ## Comment: 8am trollies YTD
  sign=if(change_8am_trolley_ytd>=0){paste("+")} else {paste("")}
  comment_8am_trolley_ytd<-paste(sign,change_8am_trolley_ytd,"% compared to YTD", year_of_interest-1)
  
  
  
  ## Comment: 8am trollies (national)
  sign=if(change_8am_trolley_national>=0){paste("+")} else {paste("")}
  comment_8am_trolley_national<-paste(sign,change_8am_trolley_national,"% compared to", format(as.Date(paste(year_of_interest-1, month_of_interest, 1, sep = '-')), '%B %Y'))
  
  ## Comment: 8am trollies YTD
  sign=if(change_8am_trolley_ytd_national>=0){paste("+")} else {paste("")}
  comment_8am_trolley_ytd_national<-paste(sign,change_8am_trolley_ytd_national,"% compared to YTD", year_of_interest-1)
  
  
  
  ## Create the executive summary table ##
  ## Key performance indicator
  KPI=c("Monthly Attends",
        "    ",
        "6 Hour PET",
        "9 Hour PET",
        "PET >= 24 hours",
        "8 AM Trolleys (YTD)","")
  
  if(length(hse_group) == 1){
    if(hse_group == "CHI"){
      KPI=c("Monthly Attends",
        "    ",
        "6 Hour PET",
        "9 Hour PET",
        "PET >= 24 hours")
    }
  }
  
  
  ## Values
  Values<-if(length(hse_group) > 1){
    
    c(format(new_attends_national,big.mark = ","),
      " ",
      paste(new_6hour_pet_national,"%"),
      paste(new_9hour_pet_national,"%"),
      paste(format(new_pet_gt_24_national,big.mark = ","),""),
      paste(format(new_8am_trolley_national,big.mark = ",")),
      paste(format(new_8am_trolley_ytd_national,big.mark = ","),"(YTD)"))
    
  } else if(hse_group == "CHI") {
    
    c(paste(format(new_attends,big.mark = ",")),
      " ",
      paste(new_6hour_pet,"%"),
      paste(new_9hour_pet,"%"),
      paste(format(new_pet_gt_24,big.mark = ","), " (", new_pet_gt_24_percentage, "% of all ED departures)",sep=""))
    
  } else {
    c(paste(format(new_attends,big.mark = ",")),
      " ",
      paste(new_6hour_pet,"%"),
      paste(new_9hour_pet,"%"),
      paste(format(new_pet_gt_24,big.mark = ","),""),
      paste(format(new_8am_trolley,big.mark = ","), " (",
            format(as.Date(paste(year_of_interest-1, month_of_interest, 1, sep = '-')), '%B'),")",sep=""),
      paste(format(new_8am_trolley_ytd,big.mark = ","),"(YTD)"))
  }
  
  
  ## Comments
  Comments<-if(length(hse_group) > 1){
    c(comment_attends_national,
      comment_attends_national_ytd,
      comment_6hour_pet,
      comment_9hour_pet,
      comment_pet_gt_24,
      comment_8am_trolley_national,
      comment_8am_trolley_ytd_national)
    
  } else if(hse_group == "CHI"){
    
    c(comment_attends,
      comment_attends_ytd,
      comment_6hour_pet,
      comment_9hour_pet,
      comment_pet_gt_24)
  } else {
    
    c(comment_attends,
      comment_attends_ytd,
      comment_6hour_pet,
      comment_9hour_pet,
      comment_pet_gt_24,
      comment_8am_trolley,
      comment_8am_trolley_ytd)
  }
  
  ## Create the table, call it "summ_tab"
  summ_tab<-data.frame(KPI,Values,Comments)
  return(summ_tab)
  #if(length(hse_group) > 1){hse_group = 'NAT'}
  #write.csv(summ_tab, paste('~/public_folder/jl_csv/kpi_reviews/kpi_executive_summaries/ED Executive Summary ', hse_group, '.csv', sep = ''), row.names = F)
  #paste('ED Executive Summary ', hse_group, '.csv', sep = '')
  
}

write_ed_summary_to_xlsx = function(){
  
  wb = createWorkbook(type = "xlsx")
  
  header_style = CellStyle(wb) + Alignment(horizontal = "ALIGN_LEFT") + Font(wb, name = "Open Sans", heightInPoints = 14, color = "9") + Fill(foregroundColor = "#4F81BD")
  tbl_style = CellStyle(wb) + Alignment(horizontal = "ALIGN_LEFT") + Font(wb, name = "Open Sans", heightInPoints = 14)
  tbl_style_1 = tbl_style + Fill(foregroundColor = "#E9EDF4")
  tbl_style_2 = tbl_style + Fill(foregroundColor = "#D0D8E7")
  
  tbl_nat = run_ed_exec_summary('NAT')
  tbl_chg = run_ed_exec_summary('CHI')
  tbl_dml = run_ed_exec_summary('DML')
  tbl_iehg = run_ed_exec_summary('IEHG')
  tbl_rcsi = run_ed_exec_summary('RCSI')
  tbl_saolta = run_ed_exec_summary('SAOLTA')
  tbl_ssw = run_ed_exec_summary('SSW')
  tbl_ul = run_ed_exec_summary('UL')
  
  gc()
  .jcall("java/lang/System", method = "gc")
  
  sheet_nat = createSheet(wb, sheetName = 'National')
  addDataFrame(tbl_nat, sheet = sheet_nat, row.names = F, colnamesStyle = header_style)
  rows = getRows(sheet_nat)
  cells = getCells(rows)
  
  for(j in c(4:9, 16:18)){
    setCellStyle(cells[[j]], tbl_style_1)
  }
  for(j in c(10:15, 19:24)){
    setCellStyle(cells[[j]], tbl_style_2)
  }
  autoSizeColumn(sheet_nat, colIndex = 1:3)
  gc()
  .jcall("java/lang/System", method = "gc")
  
  sheet_chg = createSheet(wb, sheetName = 'CHI')
  addDataFrame(tbl_chg, sheet = sheet_chg, row.names = F, colnamesStyle = header_style)
  rows = getRows(sheet_chg)
  cells = getCells(rows)
  
  for(j in c(4:9, 16:18)){
    setCellStyle(cells[[j]], tbl_style_1)
  }
  for(j in 10:15){
    setCellStyle(cells[[j]], tbl_style_2)
  }
  autoSizeColumn(sheet_chg, colIndex = 1:3)
  gc()
  .jcall("java/lang/System", method = "gc")
  
  sheet_dml = createSheet(wb, sheetName = 'DML')
  addDataFrame(tbl_dml, sheet = sheet_dml, row.names = F, colnamesStyle = header_style)
  rows = getRows(sheet_dml)
  cells = getCells(rows)
  
  for(j in c(4:9, 16:18)){
    setCellStyle(cells[[j]], tbl_style_1)
  }
  for(j in c(10:15, 19:24)){
    setCellStyle(cells[[j]], tbl_style_2)
  }
  autoSizeColumn(sheet_dml, colIndex = 1:3)
  gc()
  .jcall("java/lang/System", method = "gc")
  
  sheet_iehg = createSheet(wb, sheetName = 'IEHG')
  addDataFrame(tbl_iehg, sheet = sheet_iehg, row.names = F, colnamesStyle = header_style)
  rows = getRows(sheet_iehg)
  cells = getCells(rows)
  
  for(j in c(4:9, 16:18)){
    setCellStyle(cells[[j]], tbl_style_1)
  }
  for(j in c(10:15, 19:24)){
    setCellStyle(cells[[j]], tbl_style_2)
  }
  autoSizeColumn(sheet_iehg, colIndex = 1:3)
  gc()
  .jcall("java/lang/System", method = "gc")
  
  sheet_rcsi = createSheet(wb, sheetName = 'RCSI')
  addDataFrame(tbl_rcsi, sheet = sheet_rcsi, row.names = F, colnamesStyle = header_style)
  rows = getRows(sheet_rcsi)
  cells = getCells(rows)
  
  for(j in c(4:9, 16:18)){
    setCellStyle(cells[[j]], tbl_style_1)
  }
  for(j in c(10:15, 19:24)){
    setCellStyle(cells[[j]], tbl_style_2)
  }
  autoSizeColumn(sheet_rcsi, colIndex = 1:3)
  gc()
  .jcall("java/lang/System", method = "gc")
  
  sheet_saolta = createSheet(wb, sheetName = 'SAOLTA')
  addDataFrame(tbl_saolta, sheet = sheet_saolta, row.names = F, colnamesStyle = header_style)
  rows = getRows(sheet_saolta)
  cells = getCells(rows)
  
  for(j in c(4:9, 16:18)){
    setCellStyle(cells[[j]], tbl_style_1)
  }
  for(j in c(10:15, 19:24)){
    setCellStyle(cells[[j]], tbl_style_2)
  }
  autoSizeColumn(sheet_saolta, colIndex = 1:3)
  gc()
  .jcall("java/lang/System", method = "gc")
  
  sheet_ssw = createSheet(wb, sheetName = 'SSW')
  addDataFrame(tbl_ssw, sheet = sheet_ssw, row.names = F, colnamesStyle = header_style)
  rows = getRows(sheet_ssw)
  cells = getCells(rows)
  
  for(j in c(4:9, 16:18)){
    setCellStyle(cells[[j]], tbl_style_1)
  }
  for(j in c(10:15, 19:24)){
    setCellStyle(cells[[j]], tbl_style_2)
  }
  autoSizeColumn(sheet_ssw, colIndex = 1:3)
  gc()
  .jcall("java/lang/System", method = "gc")
  
  sheet_ul = createSheet(wb, sheetName = 'UL')
  addDataFrame(tbl_ul, sheet = sheet_ul, row.names = F, colnamesStyle = header_style)
  rows = getRows(sheet_ul)
  cells = getCells(rows)
  
  for(j in c(4:9, 16:18)){
    setCellStyle(cells[[j]], tbl_style_1)
  }
  for(j in c(10:15, 19:24)){
    setCellStyle(cells[[j]], tbl_style_2)
  }
  autoSizeColumn(sheet_ul, colIndex = 1:3)
  
  saveWorkbook(wb, "C:/Users/jack.lyons/Documents/Executive Summary ED.xlsx")
}

write_ed_summary_to_xlsx()

rm(data,trolleygar_data,month_of_interest,year_of_interest,import_biu_data,import_trolleygar_data,month,run_ed_exec_summary,standard_format,write_ed_summary_to_xlsx,year)
