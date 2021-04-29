library(shiny)
library(writexl)
library(readxl)
library(reticulate)

ui <- fluidPage(
  titlePanel("eBlock Generator for SARS-CoV-2 Mutations of Interest"),
  
  sidebarLayout(
    sidebarPanel(
      fileInput('file1', h4('Choose xlsx file:'),
                accept = c(".xlsx")),
      h4("Click here to download the outputed file:"),
      downloadButton("dl", "Download")),
  
  mainPanel(
    tabsetPanel(
      tabPanel("Input Data:", tableOutput('contents'))#,
      #tabPanel("Output Data", tabelOutput('contents')))
  )

)))

server <- function(input, output) {
  output$contents <- renderTable({
    
    req(input$file1)
    inFile <- input$file1
    read_excel(inFile$datapath, 1)})
    data <- mtcars
    
    
    
    output$dl <- downloadHandler(
      filename = function() { "spike_variants.xlsx"},
      content = function(file) {write_xlsx(data, path = file)
  })
}
shinyApp(ui, server)