library(dplyr)
library(RODBC)

channel=odbcConnect('MORTALITY')

############ Re-admissions ####################

readmissions = sqlQuery(channel, query = "select hospital_code as provider_code, specialty_group_name as discipline, date_of_admission, date_of_discharge, concat(hospital_code,'.',medical_record_number)::text as mrn, case when type_of_admission in (1,2) then 'elective' else 'emergency' end as type_of_admission, substring(date_of_discharge::text,1,4)::int as discharge_year, substring(date_of_discharge::text,6,2)::int as discharge_month from hipe_casemix_enc, medicine.specialty where hipe_casemix_enc.specialty = medicine.specialty.specialty_code and date_of_discharge >= '2022-01-01' and type_of_admission in (1,2,4,5) and not (date_of_discharge=date_of_admission and type_of_admission in (1,2)) and specialty_group_id in ('MEDICN','SURGRY') and age > 15 and hospital_code in (3,4,5,7,21,22,37,41,44,100,101,103,105,202,203,206,207,235,236,303,305,307,308,401,403,404,405,501,503,506,601,602,701,702,703,705) order by mrn, date_of_admission")

n=nrow(readmissions)

readmissions$mrn = as.character(readmissions$mrn)
readmissions$prv = c(NA,readmissions$mrn[-n])
readmissions$nxt = c(readmissions$mrn[-1],NA)

readmissions$first = ifelse(readmissions$mrn != readmissions$prv | is.na(readmissions$prv), 1, 0)
readmissions$last = ifelse(readmissions$mrn != readmissions$nxt | is.na(readmissions$nxt), 1, 0)

readmissions$type=as.character(readmissions$type)
readmissions$nxt_type = c(readmissions$type[-1],NA)
readmissions$prv_type = c(NA,readmissions$type[-n])

readmissions$emergency_readmission = c(as.character(readmissions$date_of_admission[-1]),NA)
readmissions$emergency_readmission = ifelse(readmissions$last==1 | readmissions$nxt_type=="elective", NA, readmissions$emergency_readmission)
readmissions$emergency_readmission = as.Date(readmissions$emergency_readmission)

readmissions$days = ifelse(is.na(readmissions$emergency_readmission),NA,as.numeric(readmissions$emergency_readmission-readmissions$date_of_discharge))
readmissions$readmission_lt_7 = ifelse(is.na(readmissions$days),0, ifelse(readmissions$days <= 7, 100, 0))
readmissions$readmission_lt_30 = ifelse(is.na(readmissions$days),0, ifelse(readmissions$days <= 30, 100, 0))

readmissions_agg = aggregate(cbind(readmission_lt_7, readmission_lt_30)~discharge_year+discharge_month+provider_code+discipline, data=readmissions, FUN = mean)

merge(sqlQuery(channel, query="select t1.hospital_code as provider_code, hospital_group, hospital, discipline, substring(date_of_discharge::text,1,4)::int as discharge_year, substring(date_of_discharge::text, 6,2)::int as discharge_month, sum(case when type_of_admission in (1,2) and date_of_admission=date_of_discharge then 1 else 0 end) as day_cases, sum(case when type_of_admission in (1,2) and date_of_admission=date_of_discharge then 0 else 1 end) as in_patients, avg(los) as avlos, avg(case when los < 31 then los else null end) as avlos_le_30, avg(case when los > 30 then los else null end) as avlos_gt_30 from (select date_of_admission, date_of_discharge, type_of_admission, specialty, hospital_code, case when date_of_admission=date_of_discharge and type_of_admission in (1,2) then null else case when date_of_admission=date_of_discharge then 0.5 else los end end as los from hipe_casemix_enc where date_of_discharge >= '2022-01-01' and specialty in (select specialty_code from medicine.specialty where specialty_group_id in ('MEDICN','SURGRY')) and type_of_admission in (1,2,4,5) and age > 15 and hospital_code in (3,4,5,7,21,22,37,41,44,100,101,103,105,202,203,206,207,235,236,303,305,307,308,401,403,404,405,501,503,506,601,602,701,702,703,705)) as t1, (select specialty_code, specialty_group_name as discipline from medicine.specialty) as t2, (select hospital_group_long_name as hospital_group, hospital_long_name as hospital, hospital_id as hospital_code from provider_codes) as t3 where t1.specialty=t2.specialty_code and t1.hospital_code=t3.hospital_code group by provider_code, hospital_group, hospital, discipline, discharge_year, discharge_month"), readmissions_agg, by = c('provider_code', 'discipline', 'discharge_year','discharge_month'), all = TRUE) %>% mutate(per_month = "", avlos_trimmed = "") %>% mutate(date = as.Date(paste(discharge_year,discharge_month,1,sep = "-")), discipline = ifelse(discipline=="Adult Medicine","Medical",ifelse(discipline=="Surgery","Surgical",""))) %>% select(provider_code,hospital_group,hospital,discipline,per_month,avlos_trimmed,avlos,avlos_le_30,avlos_gt_30,readmission_lt_7,readmission_lt_30,in_patients,day_cases,discharge_year,discharge_month,date) -> nqais_data

nqais_data$query_date=Sys.Date()

write.csv(x=nqais_data, file = "~/public_folder/jl_csv/NQAIS Spreadsheet.csv", row.names = FALSE)

odbcCloseAll()
rm(readmissions, readmissions_agg, channel, n, nqais_data)
