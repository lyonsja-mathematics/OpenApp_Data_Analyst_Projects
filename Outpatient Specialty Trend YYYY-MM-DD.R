library(dplyr)
options(stringsAsFactors = F)
read.csv("~/Tableau/Datasources/OPD Waiting List 2019 and 2020.csv") -> wl
wl[, c('Date','Group','Hospital','Specialty','Volume','Wait.Bin..Months.')] -> wl
names(wl)=c('date','group','hospital','specialty','volume','waitbin')
wl$date=as.Date(wl$date)
aggregate(volume~date+specialty, data=wl, FUN=sum) -> wlst
fn=paste0("~/Jupyter Dashboard Project/Outpatient Specialty Trend ", max(wlst$date), ".csv")
write.csv(x=wlst, file=fn, row.names=F)
