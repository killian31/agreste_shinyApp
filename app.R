library(shinythemes)
library(shiny)
library(agreste)
library(openxlsx)
library(shinycssloaders)
library(shinyalert)
library(bslib)
library(shinyWidgets)
library(bs4Dash)
library(shinylogs)
library(fresh)

# library(DT) # for selecting rows in a shiny table


wb <- createWorkbook()
list_sheets_with_note <- c(" ")
list_col_data_types <- list()
index_current_sheet <- 0


shinyApp(
  ui = dashboardPage(
    # freshTheme = theme,
    title = "Mise en forme de tableaux",
    
    header = dashboardHeader(title = "Mise en forme de tableaux",
                             skin = "red"),
    
    sidebar = dashboardSidebar(
      fileInput(
        inputId = "file1",
        label = "Choisir un fichier CSV",
        multiple = TRUE,
        accept = c("text/csv",
                   "text/comma-separated-values,text/plain",
                   ".csv")),
      
      # Horizontal line ----
      tags$hr(),
      
      # Input: Checkbox if file has header ----
      checkboxInput("header", "Header", TRUE),
      
      # Input: Select separator ----
      radioButtons("sep", "Separateur",
                   choices = c(Virgule = ",",
                               "Point Virgule" = ";",
                               Tab = "\t"),
                   selected = ","),
      
      # Input: Select quotes ----
      radioButtons("quote", "Guillemets",
                   choices = c(Aucun = "",
                               "Doubles Guillemets" = '"',
                               "Simples Guillemets" = "'"),
                   selected = '"'),
      
      # Horizontal line ----
      tags$hr(),
      
      # Input: Select number of rows to display ----
      radioButtons("disp", "Afficher",
                   choices = c(Head = "head",
                               Tout = "all"),
                   selected = "head"),
      
      radioButtons("virg", "Type de virgule",
                   choices = c(Virgule = ",",
                               Point = "."),
                   selected = ","),
    ),
    # controlbar = dashboardControlbar(),
    footer = dashboardFooter(),
    body = dashboardBody(
      textInput("feuille", "Nom de la feuille :"),
      textInput("title", "Titre du tableau :"),
      # Output: Data file ----
      tableOutput("contents") %>% withSpinner(),
      uiOutput("deroul"),
      tags$hr(),
      textInput("note", "Note de lecture :"),
      textInput("source", "Source"),
      textInput("champ", "Champ"),
      tags$hr(),
      actionButton("validate", "Valider"),
      textInput("output_file", "Nom du fichier de sortie", value = "output.xlsx"),
      actionButton("enreg", "Enregistrer le fichier Excel")
    ),
    
  ),
  server = function(input, output) {
    
    # track_usage(
    #   storage_mode = store_null(console = FALSE),
    #   what = "input"
    # )
    
    options(shiny.maxRequestSize=15*1024^2) # set maximum file size to 15MB
    output$contents <- renderTable({
      
      # input$file1 will be NULL initially. After the user selects
      # and uploads a file, head of that data file by default,
      # or all rows if selected, will be shown.
      
      req(input$file1)
      
      df <- read.csv(input$file1$datapath,
                     header = input$header,
                     sep = input$sep,
                     quote = input$quote,
                     check.names = FALSE,
                     dec = input$virg)
      
      
      if(input$disp == "head") {
        return(head(df, n = 4L))
      }
      else {
        return(df)
      }
      
    })
    
    output$deroul <- renderUI({
      req(input$file1)
      
      df <- read.csv(input$file1$datapath,
                     header = input$header,
                     sep = input$sep,
                     quote = input$quote,
                     check.names = FALSE,
                     dec = input$virg)
      
      lapply(1:ncol(df), function(i) {
        div(style="display: inline-block;vertical-align:top; width: 100px;",
            selectInput(inputId = paste("col_",
                                        as.character(i),
                                        sep = ""),
                        label = paste("Colonne",
                                      as.character(i),
                                      sep = " "),
                        choices = c("Texte" = "texte",
                                    "Entier" = "numerique",
                                    "Décimal" = "decimal")))
      })
    })
    
    observeEvent(input$validate, {
      # print(input$.shinylogs_input$inputs[[(length(input$.shinylogs_input$inputs) - 1)]])
      req(input$file1)
      vec_data_types <- c()
      new_sheet <- TRUE
      df <- read.csv(input$file1$datapath,
                     header = input$header,
                     sep = input$sep,
                     quote = input$quote,
                     check.names = FALSE,
                     dec = input$virg)
      
      for (i in 1:ncol(df)) {
        id <- eval(parse(text = paste("input$col_", as.character(i), sep = "")))
        vec_data_types <- append(vec_data_types, id)
      }
      
      # vector_col_types <- lapply(vec_data_types, function(x) {
      #   switch (
      #     x,
      #     "texte" = "character",
      #     "numerique" = "numeric",
      #     "decimal" = "double"
      #   ) %>% 
      #     return()
      # })
      
      df <- read.csv(input$file1$datapath,
                     header = input$header,
                     sep = input$sep,
                     quote = input$quote,
                     check.names = FALSE,
                     dec = input$virg)
      
      if (input$feuille != "" & input$title != "") {
        # showModal(modalDialog("Écriture...", footer = NULL))
        # remove sheet before adding it again (fastest way to clean it)
        if (input$feuille %in% names(wb)) {
          removeWorksheet(wb, input$feuille)
          new_sheet <- FALSE
        } else {
          new_sheet <- TRUE
        }
        # add sheet and table
        ajouter_tableau_excel(wb, df, input$feuille)
        
        # remove merge before remerging
        removeCellMerge(wb, input$feuille, 2:(ncol(df)+2-1), 1)
        ajouter_titre_tableau(wb, input$feuille, input$title, fusion = TRUE)
        
        if (input$note != "") {
          list_sheets_with_note <<- append(list_sheets_with_note, input$feuille)
          ajouter_note_lecture(wb, input$feuille, input$note)
          ajouter_source(wb, input$feuille, input$source, avec_note = TRUE)
        } else if (input$source != "") {
          ajouter_source(wb, input$feuille, input$source, avec_note = FALSE)
        }
        if (input$champ != "") {
          ajouter_champ(wb, input$feuille, input$champ)
        }
        
        
        if (isTRUE(new_sheet)) {
          index_current_sheet <<- index_current_sheet + 1
        }
        # list_col_data_types <- append(list_col_data_types, list(vec_data_types))
        list_col_data_types[[index_current_sheet]] <<- vec_data_types
        
        # removeModal()
      } else {
        shinyalert("Veuillez indiquer un nom de feuille et de titre avant de valider.", type = "error")
        print("Veuillez indiquer un nom de feuille et de titre avant de valider.")
      }
      
    })
    
    observeEvent(input$enreg, {
      req(input$file1)
      req(input$feuille)
      req(input$title)
      
      # showModal(modalDialog("Enregistrement...", footer = NULL))
      # formatting
      if (input$source != "" & input$champ != "") {
        formater_auto(classeur = wb,
                      format = "chiffres_et_donnees",
                      liste_feuilles_avec_note = list_sheets_with_note,
                      liste_type_donnees = list_col_data_types)
        # save
        saveWorkbook(wb, file = input$output_file, overwrite = TRUE)
        # removeModal()
        print("Enregistrement réussi.")
      # } else if (input$.shinylogs_input[(length(input$.shinylogs_input) - 1)]$name != "validate") {
      #   shinyalert("Veuillez valider vos modifications avant d'enregistrer", type = "warning")
        } else {
        shinyalert("Veuillez indiquer une source et un champ avant d'enregistrer", type = "error")
      }
    })
  }
)