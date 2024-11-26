library(rJava)
library(xlsxjars)
library(xlsx)
library(readxl)
library(extrafont)
options(stringsAsFactors = F)

# Before running this script download NQAIS Spreadsheet.csv
#source("~/xlsx_r_scripts/AVLOS1.R")

import_opwl_data = function(){
  
  df=read.csv('C:/Users/jack.lyons/Documents/Outpatient Data/OPD Waiting List.csv')
  
  provider_codes=read.csv('~/provider_codes_20210625.csv')[,c('hospital_long_name','hospital_group','biu_hospital_id')]
  
  names(provider_codes)=c('Hospital','group','provider_code')
  
  df=merge(x=df,y=provider_codes,by = 'Hospital',all.x = T)
  
  names(df)[names(df)=='hospital_group']='group'
  
  df$Date = as.Date(df$Date)
  
  df$Year = as.numeric(format(df$Date, '%Y'))
  
  df$Month = as.numeric(format(df$Date, '%m'))
  
  return(df)
  
}

import_nqais_data = function(){
  
  df=data.frame(read_excel("C:/Users/jack.lyons/Downloads/AVLOS1.xlsx"))
  
  r = read.csv('~/provider_codes.csv')[,c('biu_hospital_id','hospital_long_name','hospital_group')]
  
  names(r) = c('provider_code', 'Hospital', 'group')
  
  df = merge(x=df,y=r, by = 'Hospital',all.x = T)
  
  return(df)
}

run_non_ed_exec_summary = function(hse_group='NAT'){
  
  months_of_the_year<-data.frame(Month.of.Target.Date=c('January','February','March','April','May','June','July','August','September','October','November','December'),Month=c(1:12))
  
  tf<-"%d/%m/%Y"
  
  last_day = as.Date(waiting_list_last_date)
  
  year_of_interest<-as.numeric(format(last_day, '%Y'))
  
  month_of_interest<-as.numeric(format(last_day, '%m'))
  
  date_to_compare_waiting_list = as.Date("2018-01-04")
  
  op = opwl_data
  avlos = nqais_data[names(nqais_data) != 'Provider.Code']
  #avlos$Month = as.numeric(format(as.Date(avlos$Date), '%m'))
  #avlos$Year = as.numeric(format(as.Date(avlos$Date), '%Y'))
  
  if(month_of_interest != 1){
    date_of_prev_mon_data = max(subset(op, Year == year_of_interest & Month == month_of_interest-1)$Date)
  } else {
    date_of_prev_mon_data = max(subset(op, Year == year_of_interest-1 & Month == 12)$Date)
  }
  
  rf = read.csv(
    
    '~/provider_codes.csv', 
    
    stringsAsFactors = F
    
    )
  
  if(hse_group == "NAT"){
    hospital_of_interest = rf$biu_hospital_id
  } else {
    hospital_of_interest = subset(rf, hospital_group == hse_group)$biu_hospital_id
  }
  
  ## Value: surgical IP discharges
  new_surg_ip_dis<-sum(subset(avlos,avlos$Discipline=="Surgical"
                              & avlos$Month==hipe_month
                              & avlos$Year==hipe_year
                              & avlos$provider_code %in% hospital_of_interest)$In.patient.discharges)
  
  
  ## Value: surgical IP discharges
  new_surg_ip_dis_national<-sum(subset(avlos,avlos$Discipline=="Surgical"
                                       & avlos$Month==hipe_month
                                       & avlos$Year==hipe_year
  )$In.patient.discharges)
  
  
  ## Value: medical IP discharges
  new_med_ip_dis<-sum(subset(avlos,avlos$Discipline=="Medical"
                             & avlos$Month==hipe_month
                             & avlos$Year==hipe_year
                             & avlos$provider_code %in% hospital_of_interest)$In.patient.discharges)
  
  
  ## Value: medical IP discharges
  new_med_ip_dis_national<-sum(subset(avlos,avlos$Discipline=="Medical"
                                      & avlos$Month==hipe_month
                                      & avlos$Year==hipe_year
  )$In.patient.discharges)
  
  
  ## Value: surgical DC discharges
  new_surg_dc_dis<-sum(subset(avlos,avlos$Discipline=="Surgical"
                              & avlos$Month==hipe_month
                              & avlos$Year==hipe_year
                              & avlos$provider_code %in% hospital_of_interest)$Day.case.discharges)
  
  
  ## Value: surgical DC discharges
  new_surg_dc_dis_national<-sum(subset(avlos,avlos$Discipline=="Surgical"
                                       & avlos$Month==hipe_month
                                       & avlos$Year==hipe_year
  )$Day.case.discharges)
  
  
  ## Value: medical DC discharges
  new_med_dc_dis<-sum(subset(avlos,avlos$Discipline=="Medical"
                             & avlos$Month==hipe_month
                             & avlos$Year==hipe_year
                             & avlos$provider_code %in% hospital_of_interest)$Day.case.discharges)
  
  
  ## Value: medical DC discharges
  new_med_dc_dis_national<-sum(subset(avlos,avlos$Discipline=="Medical"
                                      & avlos$Month==hipe_month
                                      & avlos$Year==hipe_year
  )$Day.case.discharges)
  
  
  ## Take a subset of the OP data for the end of the current month of the current year for 
  ## the specified hospital
  op_current_month_current_year<-subset(op, Date == last_day & provider_code %in% hospital_of_interest)
  
  
  ## Take a subset of the OP data for the end of the current month of the current year
  op_current_month_current_year_national<-subset(op, Date == last_day)
  
  
  ## Value: OP waiting list (total)
  new_wait_total<-sum(op_current_month_current_year$Volume)
  
  
  ## Value: national OP waiting list (total)
  new_wait_total_national<-sum(op_current_month_current_year_national$Volume)
  
  
  ## Value: OP waiting list greater than 15
  new_wait_gt_15_number<-sum(subset(op_current_month_current_year,
                                    op_current_month_current_year$Wait.Bin..Months.>15)$Volume)
  
  
  ## Value: national OP waiting list greater than 15
  new_wait_gt_15_number_national<-sum(subset(op_current_month_current_year_national,
                                             op_current_month_current_year_national$Wait.Bin..Months.>15)
                                      $Volume)
  
  
  ## Value: OP waiting list greater than 15 as a percentage of total
  new_wait_gt_15_percentage<-round(sum(subset(op_current_month_current_year,
                                              op_current_month_current_year$Wait.Bin..Months.>15)$Volume)
                                   /sum(op_current_month_current_year$Volume),3)
  
  
  ## Value: national OP waiting list greater than 15 as a percentage of total
  new_wait_gt_15_percentage_national<-round(sum(subset(op_current_month_current_year_national,
                                                       op_current_month_current_year_national
                                                       $Wait.Bin..Months.>15)$Volume)
                                            /sum(op_current_month_current_year_national$Volume),3)
  
  
  ## Value: waiting up to 12 months
  new_wait_up_to_12_number<-sum(subset(op_current_month_current_year,
                                       op_current_month_current_year$Wait.Bin..Months.<=12)$Volume)
  
  
  ## Value: national waiting up to 12 months
  new_wait_up_to_12_number_national<-sum(subset(op_current_month_current_year_national,
                                                op_current_month_current_year_national$Wait.Bin..Months.
                                                <=12)$Volume)
  
  
  ## Value: waiting up to 12 months as a percentage of total
  new_wait_up_to_12_percentage<-round(sum(subset(op_current_month_current_year,
                                                 op_current_month_current_year$Wait.Bin..Months.
                                                 <=12)$Volume)
                                      /sum(op_current_month_current_year$Volume),3)
  
  
  ## Value: national waiting up to 12 months as a percentage of total 
  new_wait_up_to_12_percentage_national<-round(sum(subset(op_current_month_current_year_national,
                                                          op_current_month_current_year_national
                                                          $Wait.Bin..Months.<=12)$Volume)
                                               /sum(op_current_month_current_year_national$Volume),3)
  
  
  ## Value: waiting up to 15 months
  new_wait_up_to_15_number<-sum(subset(op_current_month_current_year,
                                       op_current_month_current_year$Wait.Bin..Months.<=15)$Volume)
  
  
  ## Value: national waiting up to 15 months
  new_wait_up_to_15_number_national<-sum(subset(op_current_month_current_year_national,
                                                op_current_month_current_year_national$Wait.Bin..Months.
                                                <=15)$Volume)
  
  
  ## Value: waiting up to 15 months as a percentage of total
  new_wait_up_to_15_percentage<-round(sum(subset(op_current_month_current_year,
                                                 op_current_month_current_year$Wait.Bin..Months.
                                                 <=15)$Volume)
                                      /sum(op_current_month_current_year$Volume),3)
  
  
  ## Value: national waiting up to 15 months as a percentage of total
  new_wait_up_to_15_percentage_national<-round(sum(subset(op_current_month_current_year_national,
                                                          op_current_month_current_year_national
                                                          $Wait.Bin..Months.<=15)$Volume)
                                               /sum(op_current_month_current_year_national$Volume),3)
  
  ## Calculate values for the previous year ##
  ## Value: surgical IP discharges
  prev_surg_ip_dis<-sum(subset(avlos,avlos$Discipline=="Surgical"
                               & avlos$Month==hipe_month
                               & avlos$Year==hipe_year-1
                               & avlos$provider_code %in% hospital_of_interest)$In.patient.discharges)
  
  
  ## Value: surgical IP discharges
  prev_surg_ip_dis_national<-sum(subset(avlos,avlos$Discipline=="Surgical"
                                        & avlos$Month==hipe_month
                                        & avlos$Year==hipe_year-1
  )$In.patient.discharges)
  
  
  ## Value: medical IP discharges
  prev_med_ip_dis<-sum(subset(avlos,avlos$Discipline=="Medical"
                              & avlos$Month==hipe_month
                              & avlos$Year==hipe_year-1
                              & avlos$provider_code %in% hospital_of_interest)$In.patient.discharges)
  
  
  ## Value: medical IP discharges
  prev_med_ip_dis_national<-sum(subset(avlos,avlos$Discipline=="Medical"
                                       & avlos$Month==hipe_month
                                       & avlos$Year==hipe_year-1
  )$In.patient.discharges)
  
  
  ## Value: surgical DC discharges
  prev_surg_dc_dis<-sum(subset(avlos,avlos$Discipline=="Surgical"
                               & avlos$Month==hipe_month
                               & avlos$Year==hipe_year-1
                               & avlos$provider_code %in% hospital_of_interest)$Day.case.discharges)
  
  
  ## Value: surgical DC discharges
  prev_surg_dc_dis_national<-sum(subset(avlos,avlos$Discipline=="Surgical"
                                        & avlos$Month==hipe_month
                                        & avlos$Year==hipe_year-1
  )$Day.case.discharges)
  
  
  ## Value: medical DC discharges
  prev_med_dc_dis<-sum(subset(avlos,avlos$Discipline=="Medical"
                              & avlos$Month==hipe_month
                              & avlos$Year==hipe_year-1
                              & avlos$provider_code %in% hospital_of_interest)$Day.case.discharges)
  
  
  ## Value: medical DC discharges
  prev_med_dc_dis_national<-sum(subset(avlos,avlos$Discipline=="Medical"
                                       & avlos$Month==hipe_month
                                       & avlos$Year==hipe_year-1
  )$Day.case.discharges)
  
  
  
  
  ## Waiting list (total)
  prev_wait_total<-sum(subset(op, Date == date_to_compare_waiting_list & provider_code %in% hospital_of_interest)$Volume)
  
  
  ## National waiting list (total)
  prev_wait_total_national<-sum(subset(op, Date == date_to_compare_waiting_list)$Volume)
  
  
  ## Waiting list greater than 15 months
  prev_wait_gt_15_number<-sum(subset(op, Date == date_to_compare_waiting_list & provider_code %in% hospital_of_interest & Wait.Bin..Months. > 15)$Volume)
  
  
  ## National waiting list greater than 15 months
  prev_wait_gt_15_number_national<-sum(subset(op, Date == date_to_compare_waiting_list & Wait.Bin..Months.>15)$Volume)
  
  
  ## Waiting list greater than 15 months as a percentage of all
  prev_wait_gt_15_percentage<-round(100*prev_wait_gt_15_number/prev_wait_total,1)
  
  
  ## National waiting list greater than 15 as a percentage of all
  prev_wait_gt_15_percentage_national<-round(100*prev_wait_gt_15_number_national/
                                               prev_wait_total_national,1)
  
  
  ## Waiting list up to 12 months
  prev_wait_up_to_12_number<-sum(subset(op, Date == date_of_prev_mon_data & provider_code %in% hospital_of_interest & Wait.Bin..Months.<=12)$Volume)
  
  ## National waiting list up to 12 months
  prev_wait_up_to_12_number_national<-sum(subset(op, Date == date_of_prev_mon_data & Wait.Bin..Months.<=12)$Volume)
  
  ## Waiting list up to 12 months as a percentage of all
  prev_wait_up_to_12_percentage<-round(prev_wait_up_to_12_number/prev_wait_total,1)
  
  
  ## National waiting list up to 12 months as a percentage of all
  prev_wait_up_to_12_percentage_national<-round(100*prev_wait_up_to_12_number_national/
                                                  sum(subset(op, Date == date_of_prev_mon_data)$Volume,1))
  
  
  ## Waiting list up to 15 months
  prev_wait_up_to_15_number<-sum(subset(op, Date == date_of_prev_mon_data & provider_code %in% hospital_of_interest & Wait.Bin..Months.<=15)$Volume)
  
  
  ## National waiting list up to 15 months
  prev_wait_up_to_15_number_national<-sum(subset(op, Date == date_of_prev_mon_data & Wait.Bin..Months.<=15)$Volume)
  
  
  ## Waiting list up to 15 months as a percentage of all
  prev_wait_up_to_15_percentage<-round(prev_wait_up_to_15_number/prev_wait_total,1)
  
  
  ## National waiting list up to 15 months as a percentage of all
  prev_wait_up_to_15_percentage_national<-round(100*prev_wait_up_to_15_number_national/
                                                  sum(subset(op, Date == date_of_prev_mon_data)$Volume),1)
  
  ## National waiting greater than 15 months at time of last review
  prev_wait_gt_15_number_national_2<-sum(subset(op, Date == date_of_prev_mon_data & Wait.Bin..Months.>15)$Volume)
  
  ## National waiting greater than 15 months percentage at time of last review
  prev_wait_gt_15_percentage_national_2<-round(100*prev_wait_gt_15_number_national_2/
                                                 sum(subset(op, Date == date_of_prev_mon_data)$Volume),1)
  
  ## Calculate the change from previous ##
  
  ## Change in surgical IP discharges
  change_surg_ip_dis<-round((new_surg_ip_dis-prev_surg_ip_dis)*100/prev_surg_ip_dis,1)
  
  
  ## Change in surgical IP discharges
  change_surg_ip_dis_national<-round((new_surg_ip_dis_national-prev_surg_ip_dis_national)*100/
                                       prev_surg_ip_dis_national,1)
  
  
  ## Change in medical IP discharges
  change_med_ip_dis<-round((new_med_ip_dis-prev_med_ip_dis)*100/prev_med_ip_dis,1)
  
  
  ## Change in medical IP discharges
  change_med_ip_dis_national<-round((new_med_ip_dis_national-prev_med_ip_dis_national)*100/
                                      prev_med_ip_dis_national,1)
  
  
  ## Change in surgical DC discharges
  change_surg_dc_dis<-round((new_surg_dc_dis-prev_surg_dc_dis)*100/prev_surg_dc_dis,1)
  
  
  ## Change in surgical DC discharges
  change_surg_dc_dis_national<-round((new_surg_dc_dis_national-prev_surg_dc_dis_national)*100/
                                       prev_surg_dc_dis_national,1)
  
  
  ## Change in medical DC discharges
  change_med_dc_dis<-round((new_med_dc_dis-prev_med_dc_dis)*100/prev_med_dc_dis,1)
  
  
  ## Change in medical DC discharges
  change_med_dc_dis_national<-round((new_med_dc_dis_national-prev_med_dc_dis_national)*100/
                                      prev_med_dc_dis_national,1)
  
  
  ## Change in waiting list (total)
  change_wait_total<-if(hse_group=="National"){
    round(100*(new_wait_total_national-prev_wait_total_national)/prev_wait_total_national,1)
  }  else {
    round(100*(new_wait_total-prev_wait_total)/prev_wait_total,1)
  }
  
  
  ## Change in waiting list greater than 15 months
  change_wait_gt_15<- if(hse_group=="National") {
    round(100*(new_wait_gt_15_number_national-prev_wait_gt_15_number_national)
          /prev_wait_gt_15_number_national,1)
    
  } else {
    round(100*(new_wait_gt_15_number-prev_wait_gt_15_number)
          /prev_wait_gt_15_number,1)
  }
  
  ## Comments ##
  
  if(hse_group!="CHI"){
    ## Comment: surgical IP discharges
    sign=if(change_surg_ip_dis>=0){paste("+")} else {paste("")}
    comment_surg_ip_dis<-paste(sign, change_surg_ip_dis,"% Compared to ",
                               months_of_the_year$Month.of.Target.Date[hipe_month],
                               
                               ' ',
                               
                               hipe_year-1,
                               
                               sep=''
                               
    )
    
    
    # Comment: surgical IP discharges (national)
    sign=if(change_surg_ip_dis_national>=0){paste("+")} else {paste("")}
    comment_surg_ip_dis_national<-paste(sign, change_surg_ip_dis_national,"% Compared to ",
                                        months_of_the_year$Month.of.Target.Date[hipe_month],
                                        
                                        ' ',
                                        
                                        hipe_year-1,
                                        
                                        sep=''
                                        
    )
    
    ## Comment: medical IP discharges
    sign=if(change_med_ip_dis>=0){paste("+")} else {paste("")}
    comment_med_ip_dis<-paste(sign, change_med_ip_dis,"% Compared to ",
                              months_of_the_year$Month.of.Target.Date[hipe_month],
                              
                              ' ',
                              
                              hipe_year-1,
                              
                              sep=''
                              
    )
    
    
    
    ## Comment: medical IP discharges (national)
    sign=if(change_med_ip_dis_national>=0){paste("+")} else {paste("")}
    comment_med_ip_dis_national<-paste(sign, change_med_ip_dis_national,"% Compared to ",
                                       months_of_the_year$Month.of.Target.Date[hipe_month],
                                       
                                       ' ',
                                       
                                       hipe_year-1,
                                       
                                       sep=''
                                       
    )
    
    ## Comment: surgical DC discharges
    sign=if(change_surg_dc_dis>=0){paste("+")} else {paste("")}
    comment_surg_dc_dis<-paste(sign, change_surg_dc_dis,"% Compared to ",
                               months_of_the_year$Month.of.Target.Date[hipe_month],
                               
                               ' ',
                               
                               hipe_year-1,
                               
                               sep=''
                               
    )
    
    
    ## Comment: surgical DC discharges (national)
    sign=if(change_surg_dc_dis_national>=0){paste("+")} else {paste("")}
    comment_surg_dc_dis_national<-paste(sign, change_surg_dc_dis_national,"% Compared to ",
                                        months_of_the_year$Month.of.Target.Date[hipe_month],
                                        
                                        ' ',
                                        
                                        hipe_year-1,
                                        
                                        sep=''
                                        
    )
    
    
    ## Comment: medical DC discharges
    sign=if(change_med_dc_dis>=0){paste("+")} else {paste("")}
    comment_med_dc_dis<-paste(sign, change_med_dc_dis,"% Compared to ",
                              months_of_the_year$Month.of.Target.Date[hipe_month],
                              
                              ' ',
                              
                              hipe_year-1,
                              
                              sep=''
                              
    )
    
    
    ## Comment: medical DC discharges (national)
    sign=if(change_med_dc_dis_national>=0){paste("+")} else {paste("")}
    comment_med_dc_dis_national<-paste(sign, change_med_dc_dis_national,"% Compared to ",
                                       months_of_the_year$Month.of.Target.Date[hipe_month],
                                       
                                       ' ',
                                       
                                       hipe_year-1,
                                       
                                       sep=''
                                       
    )
    
  }
  
  ## Comment: waiting list (total)
  sign<-if(change_wait_total > 0){paste("+")} else {paste("")}
  comment_wait_total<-paste(sign, change_wait_total,"% since ",
                            months_of_the_year$Month.of.Target.Date[
                              as.numeric(format(as.Date.character(date_to_compare_waiting_list,"%Y-%m-%d"), '%m'))],
                            
                            ' ',
                            
                            as.numeric(format(as.Date.character(date_to_compare_waiting_list,"%Y-%m-%d"), '%d')),
                            
                            ', ',
                            
                            as.numeric(format(as.Date.character(date_to_compare_waiting_list,"%Y-%m-%d"), '%Y')),
                            
                            sep=''
                            
  )
  
  
  ## Comment: waiting list greater than 15 months
  sign<-if(change_wait_gt_15 > 0){paste("+")} else {paste("")}
  comment_wait_gt_15_number<-paste(sign,change_wait_gt_15,"% since ",
                                   months_of_the_year$Month.of.Target.Date[
                                     as.numeric(format(as.Date.character(date_to_compare_waiting_list,"%Y-%m-%d"), '%m'))],
                                   
                                   ' ',
                                   
                                   as.numeric(format(as.Date.character(date_to_compare_waiting_list,"%Y-%m-%d"), '%d')),
                                   
                                   ', ',
                                   
                                   as.numeric(format(as.Date.character(date_to_compare_waiting_list,"%Y-%m-%d"), '%Y')),
                                   
                                   sep=''
                                   
  )
  
  
  ## Comment: waiting list up to 12 months
  comment_wait_up_to_12<-if(hse_group=="NAT"){
    paste("Prior Month: ", prev_wait_up_to_12_percentage_national,"% (",
          format(prev_wait_up_to_12_number_national,big.mark = ","),")",
          
          sep=''
          
    )
  } else{
    paste("National rate: ",100*new_wait_up_to_12_percentage_national,"%",
          
          sep=''
          
    )
  }  
  
  
  ## comment: waiting list up to 15 months
  comment_wait_up_to_15<-if(hse_group=="NAT"){
    paste("Prior Month: ", prev_wait_up_to_15_percentage_national,"% (",
          format(prev_wait_up_to_15_number_national,big.mark = ","),")",
          
          sep=''
          
    )
  } else{
    paste("National rate: ",100*new_wait_up_to_15_percentage_national,"%",
          
          sep=''
          
    )
  }
  
  
  ## Comment: waiting list greater than 15 months as a percentage of all
  comment_wait_gt_15_percentage<-if(hse_group=="NAT"){
    paste("Prior Month: ", prev_wait_gt_15_percentage_national_2,"% (",
          format(prev_wait_gt_15_number_national_2,big.mark = ","),")",
          
          sep=''
          
    )
  } else{
    paste("National rate: ",100*new_wait_gt_15_percentage_national,"%",
          
          sep=''
          
    )
  }  
  
  
  ## Create the executive summary table ##
  ## Key performance indicator
  KPI<-if(hse_group=="CHI"){
    c("OP waiting list (Total)",
      "OP waiting list (GT 15 months)",
      "Waiting up to 12 months",
      "Waiting up to 15 months",
      "Waiting more than 15 months")
  } else {
    c("Surgical IP Discharges",
      "Medical IP Discharges",
      "Surgical DC Discharges",
      "Medical DC Discharges",
      "         ",
      "OP waiting list (Total)",
      "OP waiting list (GT 15 months)",
      "Waiting up to 12 months",
      "Waiting up to 15 months",
      "Waiting more than 15 months")
    
    
  }
  
  
  ## Values
  library(stringr)
  Values<-if(hse_group=="CHI") {
    
    c(paste(format(new_wait_total,big.mark = ","),"",sep=""),
      paste(format(new_wait_gt_15_number,big.mark = ","),""),
      paste(100*new_wait_up_to_12_percentage,"% (",format(new_wait_up_to_12_number,big.mark = ","),")",
            
            sep=''
            
      ),
      paste(100*new_wait_up_to_15_percentage,"% (",format(new_wait_up_to_15_number,big.mark = ","),")",
            
            sep=''
            
      ),
      paste(100*new_wait_gt_15_percentage,"% (",format(new_wait_gt_15_number,big.mark = ","),")",
            
            sep=''
            
      )
      
    )
    
  } else {
    c(paste(format(new_surg_ip_dis,big.mark = ","),""),
      paste(format(new_med_ip_dis,big.mark=","),""),
      paste(format(new_surg_dc_dis,big.mark = ","),""),
      paste(format(new_med_dc_dis,big.mark = ","),""),
      "   ",
      paste(format(new_wait_total,big.mark = ","),""),
      paste(format(new_wait_gt_15_number,big.mark = ","),""),
      paste(100*new_wait_up_to_12_percentage,"% (",format(new_wait_up_to_12_number,big.mark = ","),")",sep=''),
      paste(100*new_wait_up_to_15_percentage,"% (",format(new_wait_up_to_15_number,big.mark = ","),")",sep=''),
      paste(100*new_wait_gt_15_percentage,"% (",format(new_wait_gt_15_number,big.mark = ","),")",sep=''))
  }
  
  
  ## Comments
  Comments<-if(hse_group=="CHI"){
    
    c(comment_wait_total,
      comment_wait_gt_15_number,
      comment_wait_up_to_12,
      comment_wait_up_to_15,
      comment_wait_gt_15_percentage)
  } else {
    
    c(comment_surg_ip_dis,
      comment_med_ip_dis,
      comment_surg_dc_dis,
      comment_med_dc_dis,
      "    ",
      comment_wait_total,
      comment_wait_gt_15_number,
      comment_wait_up_to_12,
      comment_wait_up_to_15,
      comment_wait_gt_15_percentage)
  }
  
  
  ## Create the table, call it "summ_tab"
  summ_tab<-data.frame(KPI,Values,Comments)
  return(summ_tab)
  
  # write.csv(summ_tab, paste('~/public_folder/jl_csv/kpi_reviews/kpi_executive_summaries/IPOP Executive Summary ', hse_group, '.csv', sep = ''), row.names = F)
  # paste('IPOP Executive Summary ', hse_group, '.csv', sep = '')
}

nqais_data=import_nqais_data()
opwl_data=import_opwl_data()

hipe_year=as.numeric(max(nqais_data$Year))
hipe_month=as.numeric(max(subset(nqais_data,Year==hipe_year)$Month))
waiting_list_last_date=if(
  hipe_month<11
){
  max(subset(opwl_data,Year==hipe_year & Month==hipe_month+2)$Date)
} else {
  max(subset(opwl_data,Year==hipe_year+1 & Month==hipe_month-10)$Date)
}

write_non_ed_summary_to_xlsx=function(){
  
  gc()
  .jcall("java/lang/System", method = "gc")
  
  wb = createWorkbook(type = "xlsx")
  header_style = CellStyle(wb) + Alignment(horizontal = "ALIGN_LEFT") + Font(wb, name = "Open Sans", heightInPoints = 14, color = "9") + Fill(foregroundColor = "#4F81BD")
  tbl_style = CellStyle(wb) + Alignment(horizontal = "ALIGN_LEFT") + Font(wb, name = "Open Sans", heightInPoints = 14)
  tbl_style_1 = tbl_style + Fill(foregroundColor = "#E9EDF4")
  tbl_style_2 = tbl_style + Fill(foregroundColor = "#D0D8E7")
  
  sheet_nat = createSheet(wb, sheetName = 'National')
  addDataFrame(
    run_non_ed_exec_summary('NAT'), 
    sheet = sheet_nat, 
    row.names = F, 
    colnamesStyle = header_style
    )
  
  rows = getRows(sheet_nat)
  cells = getCells(rows)
  for(j in c(4:6, 10:12, 16:18, 22:24, 28:30)){
    setCellStyle(cells[[j]], tbl_style_1)
  }
  for(j in c(7:9, 13:15, 19:21, 25:27, 31:33)){
    setCellStyle(cells[[j]], tbl_style_2)
  }
  autoSizeColumn(sheet_nat, colIndex = 1:3)
  gc()
  .jcall("java/lang/System", method = "gc")
  
  sheet_chg = createSheet(wb, sheetName = 'CHI')
  addDataFrame(
    run_non_ed_exec_summary('CHI'), 
    sheet = sheet_chg, 
    row.names = F, 
    colnamesStyle = header_style
    )
  
  rows = getRows(sheet_chg)
  cells = getCells(rows)
  for(j in c(4:6, 10:12, 16:18)){
    setCellStyle(cells[[j]], tbl_style_1)
  }
  for(j in c(7:9, 13:15)){
    setCellStyle(cells[[j]], tbl_style_2)
  }
  autoSizeColumn(sheet_chg, colIndex = 1:3)
  gc()
  .jcall("java/lang/System", method = "gc")
  
  sheet_dml = createSheet(wb, sheetName = 'DML')
  addDataFrame(
    run_non_ed_exec_summary('DML'),
    sheet = sheet_dml, 
    row.names = F, 
    colnamesStyle = header_style
    )
  
  rows = getRows(sheet_dml)
  cells = getCells(rows)
  for(j in c(4:6, 10:12, 16:18, 22:24, 28:30)){
    setCellStyle(cells[[j]], tbl_style_1)
  }
  for(j in c(7:9, 13:15, 19:21, 25:27, 31:33)){
    setCellStyle(cells[[j]], tbl_style_2)
  }
  autoSizeColumn(sheet_dml, colIndex = 1:3)
  gc()
  .jcall("java/lang/System", method = "gc")
  
  sheet_iehg = createSheet(wb, sheetName = 'IEHG')
  addDataFrame(run_non_ed_exec_summary('IEHG'), sheet = sheet_iehg, row.names = F, colnamesStyle = header_style)
  rows = getRows(sheet_iehg)
  cells = getCells(rows)
  for(j in c(4:6, 10:12, 16:18, 22:24, 28:30)){
    setCellStyle(cells[[j]], tbl_style_1)
  }
  for(j in c(7:9, 13:15, 19:21, 25:27, 31:33)){
    setCellStyle(cells[[j]], tbl_style_2)
  }
  autoSizeColumn(sheet_iehg, colIndex = 1:3)
  gc()
  .jcall("java/lang/System", method = "gc")
  
  sheet_rcsi = createSheet(wb, sheetName = 'RCSI')
  addDataFrame(run_non_ed_exec_summary('RCSI'), sheet = sheet_rcsi, row.names = F, colnamesStyle = header_style)
  rows = getRows(sheet_rcsi)
  cells = getCells(rows)
  for(j in c(4:6, 10:12, 16:18, 22:24, 28:30)){
    setCellStyle(cells[[j]], tbl_style_1)
  }
  for(j in c(7:9, 13:15, 19:21, 25:27, 31:33)){
    setCellStyle(cells[[j]], tbl_style_2)
  }
  autoSizeColumn(sheet_rcsi, colIndex = 1:3)
  gc()
  .jcall("java/lang/System", method = "gc")
  
  sheet_saolta = createSheet(wb, sheetName = 'SAOLTA')
  addDataFrame(run_non_ed_exec_summary('SAOLTA'), sheet = sheet_saolta, row.names = F, colnamesStyle = header_style)
  rows = getRows(sheet_saolta)
  cells = getCells(rows)
  for(j in c(4:6, 10:12, 16:18, 22:24, 28:30)){
    setCellStyle(cells[[j]], tbl_style_1)
  }
  for(j in c(7:9, 13:15, 19:21, 25:27, 31:33)){
    setCellStyle(cells[[j]], tbl_style_2)
  }
  autoSizeColumn(sheet_saolta, colIndex = 1:3)
  gc()
  .jcall("java/lang/System", method = "gc")
  
  sheet_ssw = createSheet(wb, sheetName = 'SSW')
  addDataFrame(run_non_ed_exec_summary('SSW'), sheet = sheet_ssw, row.names = F, colnamesStyle = header_style)
  rows = getRows(sheet_ssw)
  cells = getCells(rows)
  for(j in c(4:6, 10:12, 16:18, 22:24, 28:30)){
    setCellStyle(cells[[j]], tbl_style_1)
  }
  for(j in c(7:9, 13:15, 19:21, 25:27, 31:33)){
    setCellStyle(cells[[j]], tbl_style_2)
  }
  autoSizeColumn(sheet_ssw, colIndex = 1:3)
  gc()
  .jcall("java/lang/System", method = "gc")
  
  sheet_ul = createSheet(wb, sheetName = 'UL')
  addDataFrame(run_non_ed_exec_summary('UL'), sheet = sheet_ul, row.names = F, colnamesStyle = header_style)
  rows = getRows(sheet_ul)
  cells = getCells(rows)
  for(j in c(4:6, 10:12, 16:18, 22:24, 28:30)){
    setCellStyle(cells[[j]], tbl_style_1)
  }
  for(j in c(7:9, 13:15, 19:21, 25:27, 31:33)){
    setCellStyle(cells[[j]], tbl_style_2)
  }
  autoSizeColumn(sheet_ul, colIndex = 1:3)
  gc()
  .jcall("java/lang/System", method = "gc")
  
  saveWorkbook(wb, "C:/Users/jack.lyons/Documents/Executive Summary Non ED.xlsx")
  
}

write_non_ed_summary_to_xlsx()

rm(nqais_data,opwl_data,hipe_month,hipe_year,waiting_list_last_date,import_nqais_data,import_opwl_data,run_non_ed_exec_summary,write_non_ed_summary_to_xlsx)
