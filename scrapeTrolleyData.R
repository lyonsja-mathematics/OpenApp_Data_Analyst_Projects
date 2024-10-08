# scrapeTrolleyData.R
# Author: Jack Lyons

# Import libraries
library(dplyr)
library(rvest)

# 1.0 --- Define function
scrapeTrolleyData = function(date){
  as.Date(date) -> date
  format(date, "%d") -> day
  format(date, "%m") -> month
  format(date, "%Y") -> year
  paste0("http://137.191.241.85/ed/ED.php?EDDATE=",day,"%2F",month,"%2F",year) -> url
  read_html(url) -> w
  w %>% html_nodes("table td") %>% 
    html_text() %>% 
    gsub(pattern="\r",replacement="") %>% 
    gsub(pattern="\n",replacement="") %>% 
    gsub(pattern="\t", replacement="") -> tableCells
  
  tableCells[3:length(tableCells)] -> tableCells
  
  for(i in 1:length(tableCells)){
    if(grepl(pattern="Tallaght University Hospital", tableCells[i])){as.numeric(tableCells[i+5]) -> tallaght}
    if(grepl(pattern="Portlaoise", tableCells[i])){as.numeric(tableCells[i+5]) -> portlaoise}
    if(grepl(pattern="Tullamore", tableCells[i])){as.numeric(tableCells[i+5]) -> tullamore}
    if(grepl(pattern="Naas", tableCells[i])){as.numeric(tableCells[i+5]) -> naas}
    if(grepl(pattern="James", tableCells[i])){as.numeric(tableCells[i+5]) -> james}
    if(grepl(pattern="Vincent", tableCells[i])){as.numeric(tableCells[i+5]) -> vincent}
    if(grepl(pattern="Mullingar", tableCells[i])){as.numeric(tableCells[i+5]) -> mullingar}
    if(grepl(pattern="Mater", tableCells[i])){as.numeric(tableCells[i+5]) -> mater}
    if(grepl(pattern="Navan", tableCells[i])){as.numeric(tableCells[i+5]) -> navan}
    if(grepl(pattern="Lukes", tableCells[i])){as.numeric(tableCells[i+5]) -> luke}
    if(grepl(pattern="Wexford", tableCells[i])){as.numeric(tableCells[i+5]) -> wexford}
    if(grepl(pattern="Beaumont", tableCells[i])){as.numeric(tableCells[i+5]) -> beaumont}
    if(grepl(pattern="Connolly", tableCells[i])){as.numeric(tableCells[i+5]) -> connolly}
    if(grepl(pattern="Cavan", tableCells[i])){as.numeric(tableCells[i+5]) -> cavan}
    if(grepl(pattern="Lourdes", tableCells[i])){as.numeric(tableCells[i+5]) -> olol}
    if(grepl(pattern="Letterkenny", tableCells[i])){as.numeric(tableCells[i+5]) -> letterkenny}
    if(grepl(pattern="Mayo", tableCells[i])){as.numeric(tableCells[i+5]) -> mayo}
    if(grepl(pattern="Portiuncula", tableCells[i])){as.numeric(tableCells[i+5]) -> portiuncula}
    if(grepl(pattern="Sligo", tableCells[i])){as.numeric(tableCells[i+5]) -> sligo}
    if(grepl(pattern="Galway", tableCells[i])){as.numeric(tableCells[i+5]) -> galway}
    if(grepl(pattern="Cork University Hospital", tableCells[i])){as.numeric(tableCells[i+5]) -> cork}
    if(grepl(pattern="Kerry", tableCells[i])){as.numeric(tableCells[i+5]) -> kerry}
    if(grepl(pattern="Mercy", tableCells[i])){as.numeric(tableCells[i+5]) -> mercy}
    if(grepl(pattern="Tipperary", tableCells[i])){as.numeric(tableCells[i+5]) -> southTipp}
    if(grepl(pattern="Waterford", tableCells[i])){as.numeric(tableCells[i+5]) -> waterford}
    if(grepl(pattern="Limerick", tableCells[i])){as.numeric(tableCells[i+5]) -> limerick}
    if(grepl(pattern="Temple", tableCells[i])){as.numeric(tableCells[i+5]) -> templeStreet}
    if(grepl(pattern="Crumlin", tableCells[i])){as.numeric(tableCells[i+5]) -> crumlin}
    if(grepl(pattern="CHI at Tallaght", tableCells[i])){as.numeric(tableCells[i+5]) -> tallaghtCHI}
  }
  
  data.frame(hospitalID=c(102,201,904,1049,203,
                          908,202,403,601,910,605,
                          923,402,108,209,
                          800,500,802,809,501,
                          724,726,913,607,600,
                          300,
                          941,992,940), 
             total8am=c(naas,portlaoise,james,tallaght,tullamore,
                        mater,mullingar,navan,luke,vincent,wexford,
                        beaumont,cavan,connolly,olol,
                        galway,letterkenny,mayo,portiuncula,sligo,
                        cork,kerry,mercy,southTipp,waterford,
                        limerick,
                        crumlin,tallaghtCHI,templeStreet)) -> df
  
  df$date=format(date, "%d/%m/%Y")
  
  for(j in c(3:7,9:21)){
    df[, paste0("V", j)] = ""
  }
  
  df[, c("hospitalID", "date", paste0("V",3:7),"total8am", paste0("V", 9:21))] -> df
  paste0("~/TrolleyGAR Web Scrape/TrolleygarFile_",date,".csv") -> fileName
  write.table(x=df, sep = ",", file=fileName, row.names = F, col.names = F)
}

# 2.0 --- Run function
first.day="2023-07-01"
last.day="2023-07-31"
dts=seq.Date(from=as.Date(first.day), 
             to=as.Date(last.day), 
             by="1 day")
dts=as.character(dts)
for(k in dts){scrapeTrolleyData(k)}

# 3.0 --- Clean up
rm(dts,first.day,k,last.day,scrapeTrolleyData)
