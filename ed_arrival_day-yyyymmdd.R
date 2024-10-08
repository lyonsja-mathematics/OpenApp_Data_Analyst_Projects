ed_arrival_day_yyyymmdd=function(){
  
  thismonth=ifelse(as.numeric(substring(as.character.Date(Sys.Date()),6,7))==1,12,as.numeric(substring(as.character.Date(Sys.Date()),6,7))-1)
  thisyear=ifelse(as.numeric(substring(as.character.Date(Sys.Date()),6,7))==1,as.numeric(substring(as.character.Date(Sys.Date()),1,4))-1,as.numeric(substring(as.character.Date(Sys.Date()),1,4)))

name_end=ifelse(thismonth != 12, paste(format(as.Date(paste(thisyear,thismonth+1,1,sep="-"))-1,"%Y%m%d"),"csv",sep="."), paste(format(as.Date(paste(thisyear,12,31,sep="-")),"%Y%m%d"),"csv",sep="."))
filename_day <-paste("~/public_folder/jl_csv/edtf_dashboard/ed_arrival_day",name_end,sep="-")

######################################

first <-as.Date(paste(thisyear,"-",thismonth,"-01",sep=""),"%Y-%m-%d")
last <- as.Date(cut(first + 31, "month")) - 1
ndays <-as.numeric(last+1-first)
weekdaylist <-as.data.frame(seq(1,7,by=1))
names(weekdaylist) <-c("weekday")
weekdaylist$countw <-0
for (i in 1:7) weekdaylist$countw[i] <-sum(format(seq(first, last, "day"), "%u") == i) #Counts number of instances of each weekday in the month
month_start=as.POSIXct(paste(thisyear,'-',thismonth,'-01 00:00:00',sep=''))
if(
  thismonth!=12
  ) {
  month_end=as.POSIXct(paste(thisyear,'-',thismonth+1,'-01 00:00:00',sep=''))
} else {
  month_end=as.POSIXct(paste(thisyear+1,'-01-01 00:00:00',sep=''))
}

channel = odbcConnect("SDU")

ED <- sqlQuery(channel, paste("select A.hospital_id as hospital_id, hospital as hosp_name, referral_type, reg, depart, substring(reg::text, 9, 2)::int as day, case when referral_type = 1 then 'GP' else case when referral_type = 2 then 'Self' else 'Other' end end as source, 1 as count, dob, age, case when discharge_code in (20, 30, 80, 100) then 1 else 0 end as admitted from (select *, case when (reg::date < dob or dob is null) then 200 else case when substring(reg::text, 6, 2)::int > substring(dob::text, 6, 2)::int then substring(reg::text, 1, 4)::int-substring(dob::text, 1, 4)::int else case when substring(reg::text, 6, 2)::int < substring(dob::text, 6, 2)::int then substring(reg::text, 1, 4)::int-substring(dob::text, 1, 4)::int-1 else case when substring(reg::text, 6, 2)::int = substring(dob::text, 6, 2)::int and substring(reg::text, 9, 2)::int >= substring(dob::text, 9, 2)::int then substring(reg::text, 1, 4)::int-substring(dob::text, 1, 4)::int else substring(reg::text, 1, 4)::int-substring(dob::text, 1, 4)::int-1 end end end end as age from accident_and_emergency.data where reg between '",month_start,"' and '",month_end,"' and hospital_id != 3) as A, provider_codes as B where A.hospital_id = B.hospital_id and acuteed=1", sep = ""))
odbcCloseAll()
ED$all=1
ED$u75=ifelse(ED$age %in% 0:74,1,0)
ED$gt75=ifelse(ED$age %in% 75:110,1,0)
ED$unknown=ifelse(ED$gt75==0 & ED$u75==0,1,0)
ED$weekday <-as.numeric(format(ED$reg, "%u")) #1 to 7, 1 is Monday

#Profile ED attendances by day of week and source
profile1 <-aggregate(cbind(all,u75,gt75,unknown) ~hospital_id+hosp_name+source+weekday,data=ED,FUN=sum)
profile1n <-aggregate(cbind(all,u75,gt75,unknown) ~source+weekday,data=ED,FUN=sum)
profile1n$hospital_id=9999
profile1n$hosp_name='National'
profile1=rbind(profile1,profile1n)
profile1 <-merge(profile1,weekdaylist,by="weekday",all.x=TRUE)
profile1$all=profile1$all/profile1$countw
profile1$u75=profile1$u75/profile1$countw
profile1$gt75=profile1$gt75/profile1$countw
profile1$unknown=profile1$unknown/profile1$countw

profile1 <-profile1[,c("hospital_id","hosp_name", "source","weekday","all","u75","gt75",'unknown')]
names(profile1)=c("hospital_id","hosp_name", "source","weekday","all","U75","GE75",'Unknown')
setwd('~/')
write.csv(profile1,file=filename_day,row.names=FALSE)
odbcCloseAll()

}

ed_arrival_day_yyyymmdd()