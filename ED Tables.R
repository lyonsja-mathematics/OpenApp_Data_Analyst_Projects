library(rJava)
library(xlsxjars)
library(xlsx)
library(tidyr)

source('~/xlsx_r_scripts/ed_tables.R')
source('~/xlsx_r_scripts/trolleygar_tables.R')

gc()
.jcall("java/lang/System", method = "gc")
  
wb = createWorkbook(type = "xlsx")
  
header_style_1 = CellStyle(wb) + Alignment(horizontal = "ALIGN_LEFT") + Font(wb, name = "Trebuchet MS", heightInPoints = 18) + Fill(foregroundColor = "#B4C7DC")
  
header_style_2 = CellStyle(wb) + Alignment(horizontal = "ALIGN_LEFT") + Font(wb, name = "Trebuchet MS", heightInPoints = 18) + Fill(foregroundColor = "#B4C7DC") + Border()
  
header_style_3 = CellStyle(wb) + Alignment(horizontal = "ALIGN_LEFT") + Font(wb, name = "Trebuchet MS", heightInPoints = 18) + Fill(foregroundColor = "#B4C7DC") + Border(position=c('TOP','BOTTOM'), pen=c('BORDER_THIN',"BORDER_DOUBLE"))
  
tbl_style = CellStyle(wb) + Alignment(horizontal = "ALIGN_LEFT") + Font(wb, name = "Trebuchet MS", heightInPoints = 18)
  
params=data.frame(

    sheet_name=c('Attendances','Attendances 0-15','Attendances 75+',
                 'Admissions','Admissions 0-15','Admissions 75+',
                 'TrolleyGAR'),

    r_object_name=c('nat1','nat2','nat3',
                    'nat4','nat5','nat6',
                    'nat7'),

    f_call_age_group=as.numeric(c(1,2,3,1,2,3,0)),

    f_call_admissions=as.numeric(c(0,0,0,1,1,1,0)),

    stringsAsFactors = F

    )

  for(j in 1:7){

  assign(

    params$r_object_name[j],

    createSheet(wb, sheetName = params$sheet_name[j])

  )

  addDataFrame(
    
    if(j==7){
      
      trolleygar_table("NAT")
    
      } else {
        
        ed_table(
          
          "NAT",
          
          params$f_call_age_group[j],
          
          params$f_call_admissions[j]
          
          )
        
        },
        
        sheet=get(params$r_object_name[j]),

        row.names = F,
    
        col.names = F,
        
        startRow = 1

      )

  cells=getCells(

    getRows(

      get(params$r_object_name[j])

      )

    )

  for(i in 1:length(cells)){

    setCellStyle(

      cells[[i]],

      tbl_style

      )

    }

  for(

    i in c(11:14,16:17,19:20)

  ){

    setCellStyle(

      cells[[i]],

      header_style_1

      )

  }
  
  for(
    
    i in c(21:24,26:27,29:30)
    
  ){
    
    setCellStyle(
      
      cells[[i]],
      
      header_style_2
      
    )
    
  }
  
  n=length(cells)
  
  for(
    
    i in c((n-9):(n-6),(n-4):(n-3),(n-1):n)
    
  ){
    
    setCellStyle(
      
      cells[[i]],
      
      header_style_3
      
    )
    
  }

  col_widths=data.frame(

    col_nx=1:10,
    width=c(72,30,20,20,5,20,20,5,20,20)
  )

  for(

    i in 1:10

    ){

    setColumnWidth(

      sheet = get(params$r_object_name[j]),

      colIndex = col_widths[i,1],

      colWidth = col_widths[i,2]

      )

    }

  }
  
  # saveWorkbook(
  #   
  #   wb,
  #   
  #   "C:/Users/jack.lyons/Documents/ED Tables.xlsx"
  #   
  # )
  
  #cat("National is finished!",sep="\n")
  
  group_names=c(
    
    'CHI','DML','IEHG','RCSI','SAOLTA','SSW','UL'
    
  )
  
  for(j in 1:7){
  
  assign(
    
    paste0(
      
      'sheet_',
      
      group_names[j]
      
    ),
    
    createSheet(
      
      wb,
      
      group_names[j]
      
    )
    
  )
  
  table1=ed_table(group_names[j],1,0)
  table2=ed_table(group_names[j],1,1)
  table3=ed_table(group_names[j],2,0)
  table4=ed_table(group_names[j],2,1)
  table5=ed_table(group_names[j],3,0)
  table6=ed_table(group_names[j],3,1)
  table7=trolleygar_table(group_names[j])
  
  df1=data.frame(
    
    index=1:7,
    
    stringsAsFactors = F
  )
  
  df1$start_row_i=(nrow(
    
    get(
      
      paste0("table",df1$index)
      
    )
    
  )+2)*df1$index - (nrow(
    
    get(
      
      paste0("table",df1$index)
      
    ))+1)
  
  
  
  for(i in 1:7){
    
    addDataFrame(
      
      get(paste0("table",i)),
      
      sheet=get(
        
        paste0('sheet_',group_names[j])
        
      ),
      
      row.names = F,
      
      col.names = F,
      
      startRow = df1[i,'start_row_i']
      
    )
    
  }
  
  cells=getCells(
    
    getRows(
      
      get(paste0('sheet_',group_names[j]))
      
    )
    
  )
  
  n1=nrow(table1)
  n2=n1+nrow(table2)
  n3=n2+nrow(table3)
  n4=n3+nrow(table4)
  n5=n4+nrow(table5)
  n6=n5+nrow(table6)
  n7=n6+nrow(table7)
  
  header_style_1_cells = c(
    
    11:14,16:17,19:20,
    
    (10*n1 + 11):(10*n1 + 14),(10*n1 + 16):(10*n1 + 17),(10*n1 + 19):(10*n1 + 20),
    
    (10*n2 + 11):(10*n2 + 14),(10*n2 + 16):(10*n2 + 17),(10*n2 + 19):(10*n2 + 20),
    
    (10*n3 + 11):(10*n3 + 14),(10*n3 + 16):(10*n3 + 17),(10*n3 + 19):(10*n3 + 20),
    
    (10*n4 + 11):(10*n4 + 14),(10*n4 + 16):(10*n4 + 17),(10*n4 + 19):(10*n4 + 20),
    
    (10*n5 + 11):(10*n5 + 14),(10*n5 + 16):(10*n5 + 17),(10*n5 + 19):(10*n5 + 20),
    
    (10*n6 + 11):(10*n6 + 14),(10*n6 + 16):(10*n6 + 17),(10*n6 + 19):(10*n6 + 20)
    
  )
  
  header_style_2_cells = c(
    
    21:24,26:27,29:30,
    
    (10*n1 + 21):(10*n1 + 24),(10*n1 + 26):(10*n1 + 27),(10*n1 + 29):(10*n1 + 30),
    
    (10*n2 + 21):(10*n2 + 24),(10*n2 + 26):(10*n2 + 27),(10*n2 + 29):(10*n2 + 30),
    
    (10*n3 + 21):(10*n3 + 24),(10*n3 + 26):(10*n3 + 27),(10*n3 + 29):(10*n3 + 30),
    
    (10*n4 + 21):(10*n4 + 24),(10*n4 + 26):(10*n4 + 27),(10*n4 + 29):(10*n4 + 30),
    
    (10*n5 + 21):(10*n5 + 24),(10*n5 + 26):(10*n5 + 27),(10*n5 + 29):(10*n5 + 30),
    
    (10*n6 + 21):(10*n6 + 24),(10*n6 + 26):(10*n6 + 27),(10*n6 + 29):(10*n6 + 30)
    
    )
  
  header_style_3_cells = c(
    
    (10*n1 - 9):(10*n1 - 6),(10*n1 - 4):(10*n1 - 3),(10*n1 - 1):(10*n1),
    
    (10*n2 - 9):(10*n2 - 6),(10*n2 - 4):(10*n2 - 3),(10*n2 - 1):(10*n2),
    
    (10*n3 - 9):(10*n3 - 6),(10*n3 - 4):(10*n3 - 3),(10*n3 - 1):(10*n3),
    
    (10*n4 - 9):(10*n4 - 6),(10*n4 - 4):(10*n4 - 3),(10*n4 - 1):(10*n4),
    
    (10*n5 - 9):(10*n5 - 6),(10*n5 - 4):(10*n5 - 3),(10*n5 - 1):(10*n5),
    
    (10*n6 - 9):(10*n6 - 6),(10*n6 - 4):(10*n6 - 3),(10*n6 - 1):(10*n6),
    
    (10*n7 - 9):(10*n7 - 6),(10*n7 - 4):(10*n7 - 3),(10*n7 - 1):(10*n7)
    
  )
  
  for(
    
    l in 1:length(cells)
    
  ){
    
    if(
      
      l %in% header_style_1_cells
      
    ){
      
      setCellStyle(
        
        cells[[l]],
        
        header_style_1
        
      )
      
    } else 
    
    if(
      
      l %in% header_style_2_cells
      
    ){
      
      setCellStyle(
        
        cells[[l]],
        
        header_style_2
        
      )
      
    } else
    
    if(
      
      l %in% header_style_3_cells
      
    ){
      
      setCellStyle(
        
        cells[[l]],
        
        header_style_3
        
      )
      
    } else {
      
      setCellStyle(
        
        cells[[l]],
        
        tbl_style
        
      )
      
    }
    
  }
  
  col_widths=data.frame(
    
    col_nx=1:10,
    width=c(72,30,20,20,5,20,20,5,20,20)
  )
  
  for(
    
    k in 1:10
    
  ){
    
    setColumnWidth(
      
      sheet = get(paste0('sheet_',group_names[j])),
      
      colIndex = col_widths[k,1],
      
      colWidth = col_widths[k,2]
      
    )
    
  }
  
}
  
saveWorkbook(wb,"C:/Users/jack.lyons/Documents/ED Tables.xlsx")
  