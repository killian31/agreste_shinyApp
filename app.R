### https://rstudio.github.io/shinythemes/ ###

library(shinythemes)
library(shiny)
library(agreste)
library(openxlsx)
library(shinycssloaders)
library(shinyalert)

wb <- createWorkbook()
liste_f_note <- c(" ")
list_d_types <- list()
ind <- 0

shinyApp(
  ui = fluidPage(
    shinythemes::themeSelector(),# <--- Add this somewhere in the UI
    theme = shinytheme("united"),
    sidebarLayout(
      
      # Sidebar panel for inputs ----
      sidebarPanel(
        
        # Input: Select a file ----
        fileInput("file1", "Choisir un fichier CSV",
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
                     selected = "head")
        
      ),
      
      # Main panel for displaying outputs ----
      mainPanel(
        
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
        
      )
      
    )
  ),
  
  server = function(input, output) {
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
                     check.names = FALSE)
      
      
      if(input$disp == "head") {
        return(head(df))
      }
      else {
        return(df)
      }
      
    })
    
    output$deroul <- renderUI({
      req(input$file1)
      # selectInput("test", "Test", choices = c("Texte" = "texte", 
      #                                         "Numérique" = "numerique",
      #                                         "Décimal" = "decimal"))
      df <- read.csv(input$file1$datapath,
                     header = input$header,
                     sep = input$sep,
                     quote = input$quote,
                     check.names = FALSE)
      
      lapply(1:ncol(df), function(i) {
        div(style="display: inline-block;vertical-align:top; width: 100px;",
            selectInput(inputId = paste("col_",
                                        as.character(i),
                                        sep = ""),
                        label = paste("Colonne",
                                      as.character(i),
                                      sep = " "),
                        choices = c("Texte" = "texte",
                                    "Numérique" = "numerique",
                                    "Décimal" = "decimal")))
      })
    })
    
    observeEvent(input$validate, {
      req(input$file1)
      
      vec_d_types <- c()
      new_sheet <- TRUE
      df <- read.csv(input$file1$datapath,
                     header = input$header,
                     sep = input$sep,
                     quote = input$quote,
                     check.names = FALSE)
      
      if (input$feuille != "" & input$title != "") {
        showModal(modalDialog("Écriture...", footer = NULL))
        # remove sheet before adding it again (fastest way to clean it)
        if (input$feuille %in% names(wb)) {
          removeWorksheet(wb, input$feuille)
          new_sheet <- FALSE
        } else {
          # print(paste("Ajout de ", input$feuille, "réussi."))
          new_sheet <- TRUE
        }
        # add sheet and table
        ajouter_tableau_excel(wb, df, input$feuille)
        
        # remove merge before remerging
        removeCellMerge(wb, input$feuille, 2:(ncol(df)+2-1), 1)
        ajouter_titre_tableau(wb, input$feuille, input$title, fusion = TRUE)

        if (input$note != "") {
          liste_f_note <<- append(liste_f_note, input$feuille)
          ajouter_note_lecture(wb, input$feuille, input$note)
          ajouter_source(wb, input$feuille, input$source, avec_note = TRUE)
        } else if (input$source != "") {
          ajouter_source(wb, input$feuille, input$source, avec_note = FALSE)
        }
        if (input$champ != "") {
          ajouter_champ(wb, input$feuille, input$champ)
        }
        
        
        for (i in 1:ncol(df)) {
        id <- eval(parse(text = paste("input$col_", as.character(i), sep = "")))
        vec_d_types <- append(vec_d_types, id)
        }
        if (isTRUE(new_sheet)) {
          ind <<- ind + 1
        }
        # list_d_types <- append(list_d_types, list(vec_d_types))
        list_d_types[[ind]] <<- vec_d_types
        
        removeModal()
      } else {
        shinyalert("Veuillez indiquer un nom de feuille et de titre avant de valider.", type = "error")
        print("Veuillez indiquer un nom de feuille et de titre avant de valider.")
      }
      
    })
    
    observeEvent(input$enreg, {
      showModal(modalDialog("Enregistrement...", footer = NULL))
      # formatting
      formater_auto(classeur = wb,
                    format = "chiffres_et_donnees",
                    liste_feuilles_avec_note = liste_f_note,
                    liste_type_donnees = list_d_types)
      # save
      saveWorkbook(wb, file = input$output_file, overwrite = TRUE)
      removeModal()
      print("Enregistrement réussi.")
    })
  }
)
