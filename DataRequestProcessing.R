#https://www.r-bloggers.com/2019/07/excel-report-generation-with-shiny/

library(ERICDataProc)
options(shiny.maxRequestSize=50*1024^2)


DR_config <- setup_DR_config_values()
OutputCols <- DR_config["DROutputColsWithDist"]
newColNames <- DR_config["DRColNamesWithDist"]
colCount <- DR_config["colCount"]


DURHAM_DESIGS <- DR_config["DURHAM_DESIGS"]
NORTH_DESIGS <- DR_config["NORTH_DESIGS"]
NP_DESIGS <- DR_config["NP_DESIGS"]
SCOT_DESIGS <- DR_config["SCOT_DESIGS"]
TV_DESIGS <- DR_config["TV_DESIGS"]


  # minimal Shiny UI
ui <- fluidPage(
  titlePanel("Commercial data request"),
  tags$br(),

  fileInput("csvfile", "Choose CSV File",
            accept = c(
              "text/csv",
              "text/comma-separated-values,text/plain",
              ".csv")
  ),

  radioButtons(inputId = "datareq",label="Data required",c("Protected & Notable"="p&n","All species"="all")),

  checkboxInput(inputId = "distcheck",label = "Include distance check?",FALSE),


  textInput(inputId = "gridref",label = "Grid ref",value = ""),
  sliderInput(inputId = "dist",label = "Remove records further than ",value = 2,min=1,max=5),


  checkboxInput(inputId = "sensitivecheck",label = "Highlight sensitive species?",FALSE),

  checkboxGroupInput(inputId="optional",label = "Remove optional designations",
                     choices = c("Durham"="DURHAM_DESIGS",
                                 "Northumberland"="NORTH_DESIGS",
                                 "Northumberland National Park"="NP_DESIGS",
                                 "Scotland"="SCOT_DESIGS",
                                 "Tees Valley"="TV_DESIGS")),

  textInput(inputId = "outputfile",label = "Output filename",value = ""),
  textOutput(outputId = "msg"),



  downloadButton(
    outputId = "okBtn",
    label = "Process data")


)

# minimal Shiny server
server <- function(input, output) {


  output$okBtn <- downloadHandler(
    filename = function() {

      ifelse(stringr::str_ends(input$outputfile,'xlsx'),input$outputfile,paste0(input$outputfile,'.xlsx'))

    },
    content = function(file) {

      req_type <- ifelse((input$datareq=='p&n'),1,0)
      del_dist=1
      del_dist <- ifelse(input$distcheck,1,0)
      if (input$distcheck) {

        OutputCols <- DR_config["DROutputColsWithDist"]
        newColNames <- DR_config["DRColNamesWithDist"]
      }
      else
      {


        OutputCols <- DR_config["DROutputCols"]
        newColNames <- DR_config["DRColNames"]
      }

      grid_ref <- input$gridref
      grid_ref <- stringr::str_replace_all(grid_ref,' ','')
      distance <- input$dist

      optional_desigs <- input$optional


      inFile <- input$csvfile

      if (is.null(inFile))
        return(NULL)

      raw_data <- read.csv(inFile$datapath,header = TRUE)


      #Format the date
      raw_data$Sample.Dat <- formatDates(raw_data$Sample.Dat)

      # #Filter out the records we don't want
      raw_data <- filter_by_survey(raw_data)

      #Remove GCN licence zero counts
      raw_data <- filter_GCN(raw_data)

      #Merge location columns (EA doesn't need this)
      raw_data <- merge_location_cols(raw_data)

      #Update labels, survey names, etc.
      raw_data <- do_data_fixes(raw_data)

      #Sort out designations (EA doesn't need this)
      raw_data <- fix_designations(raw_data)


      # Easting & northing calculations
      raw_data <- E_and_N_calcs(raw_data)

      #Filter by distance if required
      if (del_dist == 1) {

        raw_data <- filter_by_distance(raw_data,distance,grid_ref)
      }


      # Setup protected & notable data
      pandn_data <- raw_data
      pandn_data <- process_pandn_data(pandn_data,DR_config,optional_desigs)

      # If its an all species request get the other records
      if (req_type == 0) {
        other_data <- dplyr::setdiff(raw_data,pandn_data)

      }

      #Format and output to Excel

      outputdata <- format_and_check_data(pandn_data,OutputCols, newColNames, input$sensitivecheck)

      sheet_name = 'Protected & Notable Species'
      XL_wb <- openxlsx::createWorkbook()

      XL_wb <- format_DR_Excel_output(XL_wb, sheet_name, outputdata,input$sensitivecheck,del_dist,DR_config,sheet_number = 1)

      #If its an All species request we want two sheets
      if (req_type != 1) {

        #Format and output to Excel
        outputdata <- format_and_check_data(other_data,OutputCols, newColNames, input$sensitivecheck)
        sheet_name = 'Other Species'
        XL_wb <- format_DR_Excel_output(XL_wb, sheet_name, outputdata,input$sensitivecheck,del_dist,DR_config,sheet_number = 2)

      }


      openxlsx::saveWorkbook(XL_wb,file,overwrite = TRUE)
    }
  )


}

shinyApp(ui, server)
