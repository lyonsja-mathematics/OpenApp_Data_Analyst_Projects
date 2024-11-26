
# If the required packages are not already installed, install them
for(p in c("reshape2","tidyr","shiny","shinydashboard","plotly","formattable","dplyr")){
  if(
    !(p %in% installed.packages())
  ){
    install.packages(pkgs = p, dependencies = TRUE)
  }
}

library(reshape2)
library(tidyr)
library(shiny)
library(shinydashboard)
library(plotly)
library(formattable)
library(dplyr)
options(stringsAsFactors = FALSE)

# Set the script's directory as the working directory
#setwd("C:/Users/jack.lyons/Documents/RShiny/prototype_app2/")

ui = dashboardPage(
  
  dashboardHeader(
    title = "Test Dashboard",
    disable = FALSE,
    
    dropdownMenu(
      messageItem(
        from = "Jack",
        message = "Is the dropdown menu working?",
        icon = icon("gamepad"),
        time = Sys.time()
      )
    )
  ),
  
  dashboardSidebar(
    
    disable = FALSE,
    
    
    sidebarMenu(
      
      menuItem(
        text="Dashboard",
        tabName = "dashboard",
        icon = icon("dashbaord"),
        badgeLabel = "New!",
        badgeColor = "green"
      ),
      
      menuItem(
        text="Widgets",
        tabName="widgets",
        icon = icon("th")
      ),
      
      menuItem(
        text = "HREF",
        icon = icon("file-code-o"),
        href = "https://rstudio.github.io/shinydashboard/structure.html#sidebar"
      ),
      
      selectInput(
        inputId = "hosp",
        label = "Hospital",
        choices = c("Hospital A"=9999, "Hospital B"=923, "Hospital C" = 800, "Hospital D" = 726, "Hospital E"=940),
        selected = 9999
      )
      
      
      
    )
    
  ),
  
  dashboardBody(
    
    # tags$h2(
    #   "TrolleyGAR",
    #   tags$hr(
    #     style = "border-color: #000000; border-width: 3px"
    #   )
    # ),
    
    fluidRow(
      box(
        status = "primary",
        title = "Trolley Total At 8:00 AM Moving Average",
        plotlyOutput(outputId = "Trolley_30DMA"),
        width = 11
      )
    ),
    fluidRow(
      box(
        title = "Trolley Total At 8:00 AM",
        status = "primary",
        selectInput(
          inputId = "Trolley_table_month",
          label = "Month",
          choices = c(
            "January"=1,
            "February"=2,
            "March"=3,
            "April"=4,
            "May"=5,
            "June"=6,
            "July"=7,
            "August"=8,
            "September"=9,
            "October"=10#,
            #"November"=11,
            #"December"=12
          ),
          selected=10
        ),
        formattableOutput(outputId = "Trolley_count_table")
      )
    )
  )
)

server = function(input, output){
  
  output$menu = renderMenu({
    sidebarMenu(
      menuItem(
        text = "Menu Item",
        icon = icon("calender")
      )
    )
  })
  
  output$Trolley_30DMA = renderPlotly({
    
    load(file = "trolley_30dmav.RData")
    load(file = "Trolley_count_thresholds.RData")
    
    trolley_30dmav %>% 
      subset(
        hosp_code==input$hosp 
        #hosp_code == 923
        & year >= max(year)-2
      ) %>% 
      
      merge(
        y = Trolley_count_thresholds %>% rename(hosp_code = biu_hospital_id),
        by = 'hosp_code',
        all.x = TRUE
      ) %>% 
      mutate(
        Chart_date = as.Date(
          paste(
            max(year), 
            month,
            day,
            sep = "-"
          )
        )
      ) %>% 
      arrange(Chart_date) -> Table4
    
    X_ticks = seq.Date(
      from = min(Table4$Chart_date), 
      to = max(Table4$Chart_date), 
      by = 'months'
    )
    
    plot_ly(data = Table4, type = 'scatter', mode = 'lines', x = ~Chart_date, y = ~MovingAverage, 
            name = ~year, hoverinfo = 'text', text = ~paste(format(Chart_date, "%B"), round(MovingAverage,1), sep = "\n")) %>%
      
      add_trace(
        y = ~threshold,
        line = list(
          dash = 'dash',
          color = 'red'
        ),
        showlegend = FALSE,
        text = ~paste(
          "Tipping Point",
          round(
            threshold,
            1
          ),
          sep = "\n"
        )
      ) %>%
      
      layout(title = "", xaxis=list(title = "", tickangle=45, tickformat = "%B", tickangle = 45, tickvals = X_ticks),
             yaxis = list(title = ""),
             legend = list(orientation = 'h', x=0, y=1.1))
    
    
  })
  
  output$Trolley_count_table = renderFormattable({
    
    load(file = "trolley_count_table.RData")
    
    Table = trolley_count_table %>% 
      
      subset(
        biu_hospital_id==input$hosp
        #biu_hospital_id == 923
        & year>=max(year)-1
        #& month == 4
        & month==input$Trolley_table_month
      ) %>% 
      
      select(metric, year, month, threshold, value)
    
    year0 = min(Table$year)
    year1 = max(Table$year)
    month0 = min(Table$month)
    
    Table$year = ifelse(
      Table$year==year0,
      "year0",
      "year1"
    )
    
    Table = Table %>% spread(year, value)
    Table$metric = factor(Table$metric)
    levels(Table$metric) = c("Trolley count 8:00 AM", "Trolley count per day", "Number of days breaching threshold")
    row.names(Table) = Table$metric
    Table = Table %>% select(threshold, year0, year1)
    
    Table$percentchange = ifelse(
      Table$year1 < Table$year0,
      paste0(
        round(
          (Table$year1-Table$year0)*100/Table$year0,
          1
        ),
        "%"),
      paste0(
        "+",
        round(
          (Table$year1-Table$year0)*100/Table$year0,
          1
        ),
        "%"
      )
    )
    
    
    Header2 = format(as.Date(paste(year0,month0,1,sep = "-")),"%b %Y")
    Header3 = format(as.Date(paste(year1,month0,1,sep = "-")),"%b %Y")
    
    names(Table) = c("Threshold", Header2, Header3, "% Change")
    
    Table[
      "Trolley count 8:00 AM",
      1:3
      ] = format(
        as.numeric(
          Table[
            "Trolley count 8:00 AM",
            1:3
            ]
        ),
        big.mark = ','
      )
    
    Table[
      "Trolley count per day",
      1:3
      ] = round(
        as.numeric(
          Table[
            "Trolley count per day",
            1:3
            ]
        ),
        1
      )
    
    Table[
      "Number of days breaching threshold",
      "Threshold"
      ] = 0
    
    formattable(Table)
    
    
  })
  
}

shinyApp(ui = ui, server = server)

