standard_format = function(time_stamp){
  x = as.character(time_stamp)
  t = data.frame(m = c('01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12'), long_name = c('January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'), short_name = c('Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'))
  
  x = gsub('/', '-', x)
  for(i in 1:nrow(t)){
    x = gsub(t[i, 2], t[i, 1], x)
    x = gsub(t[i, 3], t[i, 1], x)
  }
  
  #if(nrow(data.frame(table(nchar(x)))) > 1){return(warning('Not all elements in this list have the same format. The function standard_format() needs all elements in the list to have the same format.'))}
  y = subset(x, !(x %in% c('', ' ', 'NA', 'N/A', 'NULL')) & complete.cases(x))[1]
  z = data.frame(table(strsplit(y, '')))
  names(z) = c('char', 'count')
  type = ifelse(':' %in% z$char, 'time', 'date')
  if(type == 'date'){
    if(substr(y, 3, 3) == '-' & nchar(y) == 8){f = '%d-%m-%y'}
    if(substr(y, 3, 3) == '-' & nchar(y) == 10){f = '%d-%m-%Y'}
    if(substr(y, 3, 3) != '-' & nchar(y) == 10){f = '%Y-%m-%d'}
  } else {
    if(z[z$char == ':', 'count'] == 1 & substr(y, 3, 3) == '-' & nchar(y) %in% 13:14){f = '%d-%m-%y %H:%M'}
    if(z[z$char == ':', 'count'] == 1 & substr(y, 3, 3) == '-' & nchar(y) %in% 15:16){f = '%d-%m-%Y %H:%M'}
    if(z[z$char == ':', 'count'] == 1 & substr(y, 3, 3) != '-' & nchar(y) %in% 15:16){f = '%Y-%m-%d %H:%M'}
    if(z[z$char == ':', 'count'] == 2 & substr(y, 3, 3) == '-' & nchar(y) %in% 16:17){f = '%d-%m-%y %H:%M:%S'}
    if(z[z$char == ':', 'count'] == 2 & substr(y, 3, 3) == '-' & nchar(y) %in% 18:19){f = '%d-%m-%Y %H:%M:%S'}
    if(z[z$char == ':', 'count'] == 2 & substr(y, 3, 3) != '-' & nchar(y) %in% 18:19){f = '%Y-%m-%d %H:%M:%S'}
  }
  v = as.POSIXct(
    as.character(x), 
    format = f,
    tz = "GMT"
  )
  
  return(v)

}

month = function(time_stamp, format = '%Y-%m-%d'){
  m = as.numeric(format(as.Date(time_stamp, format = format), '%m'))
  return(m)
}

year = function(time_stamp, format = '%Y-%m-%d'){
  y = as.numeric(format(as.Date(time_stamp, format = format), '%Y'))
  return(y)
}
