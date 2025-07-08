pacman::p_load(
  shiny,
  shinyjs,
  DT,
  officer,
  flextable,
  readr,
  dplyr,
  writexl,
  lubridate,
  stringr
)

fichier_donnees <- "S:/Projekte/ZIG1_PHIRA/2_tool/evenements_data.csv"
fichier_modele <- "S:/Projekte/ZIG1_PHIRA/2_tool/Template_bulletin_morocco.docx"

serveur <- function(input, output, session) {
  useShinyjs()

  donnees <- reactiveVal({
    tryCatch({
      read_csv(fichier_donnees, show_col_types = FALSE,
               col_types = cols(
                 Version = col_double(),
                 Date = col_date(),
                 Derniere_Modification = col_datetime()
               ))
    }, error = function(e) {
      showModal(modalDialog("Erreur de lecture du fichier CSV : ", e$message))
      data.frame()
    })
  })

  session$userData$indice_edition <- NULL

  observeEvent(input$enregistrer, {
    nouvelle_id <- sprintf("EVT_%s_%03d", format(Sys.Date(), "%Y%m%d"),
                           sum(grepl(format(Sys.Date(), "%Y%m%d"), donnees()$ID_Evenement)) + 1)

    nouvel_evenement <- data.frame(
      ID_Evenement = nouvelle_id,
      Mise_a_jour_de = NA,
      Version = 1,
      Date = input$date,
      Periode_Rapport = paste(input$periode[1], "à", input$periode[2]),
      Derniere_Modification = force_tz(Sys.time(), "Africa/Casablanca"),
      Titre = input$titre,
      Type_Evenement = input$type_evenement,
      Categorie = input$categorie,
      Sources = input$sources,
      Links = input$links,
      Situation = input$situation,
      Evaluation_Risque = input$evaluation_risque,
      Mesures = input$mesures,
      Pertinence_Sante = input$pertinence_sante,
      Pertinence_Voyageurs = input$pertinence_voyageurs,
      Commentaires = input$commentaires,
      Resume = input$resume,
      Autre = input$autre,
      Statut = input$statut,
      stringsAsFactors = FALSE
    )

    donnees(bind_rows(donnees(), nouvel_evenement))

    tryCatch({
      write_csv(donnees(), fichier_donnees)
    }, error = function(e) {
      showModal(modalDialog("Erreur lors de l'enregistrement :", e$message))
    })

    showModal(modalDialog(
      title = "Succès",
      "L'événement a été ajouté avec succès.",
      easyClose = TRUE,
      footer = modalButton("Fermer")
    ))

    updateTextInput(session, "titre", value = "")
    updateSelectInput(session, "type_evenement", selected = "")
    updateSelectInput(session, "categorie", selected = "")
    updateTextInput(session, "sources", value = "")
    updateTextInput(session, "links", value = "")
    updateDateInput(session, "date", value = Sys.Date())
    updateDateRangeInput(session, "periode", start = Sys.Date() - 7, end = Sys.Date())
    updateTextAreaInput(session, "situation", value = "")
    updateTextAreaInput(session, "evaluation_risque", value = "")
    updateTextAreaInput(session, "mesures", value = "")
    updateSelectInput(session, "pertinence_sante", selected = "")
    updateSelectInput(session, "pertinence_voyageurs", selected = "")
    updateTextInput(session, "commentaires", value = "")
    updateTextAreaInput(session, "resume", value = "")
    updateTextInput(session, "autre", value = "")
    updateSelectInput(session, "statut", selected = "actif")
  })

  observeEvent(input$enregistrer_nouveau, {
    sel <- input$tableau_evenements_actifs_rows_selected
    if (length(sel) != 1) {
      showModal(modalDialog("Veuillez sélectionner un événement pour le copier."))
      return()
    }

    ev_copie <- donnees() %>% filter(Statut == "actif") %>% slice(sel)
    nouvelle_id <- sprintf("EVT_%s_%03d", format(Sys.Date(), "%Y%m%d"),
                           sum(grepl(format(Sys.Date(), "%Y%m%d"), donnees()$ID_Evenement)) + 1)

    ev_copie$Mise_a_jour_de <- ev_copie$ID_Evenement
    ev_copie$ID_Evenement <- nouvelle_id
    ev_copie$Derniere_Modification <- force_tz(Sys.time(), "Africa/Casablanca")
    ev_copie$Version <- 1

    donnees(bind_rows(donnees(), ev_copie))
    write_csv(donnees(), fichier_donnees)
  })

  observeEvent(input$count_selected_excel, {
    selected <- input$tableau_evenements_actifs_rows_selected
    session$sendCustomMessage("show_export_count", length(selected))
  })

  observeEvent(input$count_selected_word, {
    selected <- input$tableau_evenements_actifs_rows_selected
    session$sendCustomMessage("show_export_count", length(selected))
  })

  observeEvent(input$mettre_a_jour, {
    indice <- session$userData$indice_edition
    if (is.null(indice)) return()

    donnees_actuelles <- donnees()
    donnees_actuelles[indice, ] <- donnees_actuelles[indice, ] %>%
      mutate(
        Titre = input$titre,
        Type_Evenement = input$type_evenement,
        Categorie = input$categorie,
        Sources = input$sources,
        Links = input$links,
        Date = input$date,
        Periode_Rapport = paste(input$periode[1], "à", input$periode[2]),
        Situation = input$situation,
        Evaluation_Risque = input$evaluation_risque,
        Mesures = input$mesures,
        Pertinence_Sante = input$pertinence_sante,
        Pertinence_Voyageurs = input$pertinence_voyageurs,
        Commentaires = input$commentaires,
        Resume = input$resume,
        Autre = input$autre,
        Statut = input$statut,
        Derniere_Modification = force_tz(Sys.time(), "Africa/Casablanca")
      )

    donnees(donnees_actuelles)
    write_csv(donnees_actuelles, fichier_donnees)
    session$userData$indice_edition <- NULL
  })

  observeEvent(input$editer, {
    selection <- input$tableau_evenements_actifs_rows_selected
    df <- donnees()
    if (!"Statut" %in% names(df)) {
      showModal(modalDialog("La colonne 'Statut' est manquante dans les données. Veuillez vérifier le fichier CSV."))
      return()
    }
    if (length(selection) == 1) {
      evenement <- df %>% filter(Statut == "actif")[selection, ]
      updateTextInput(session, "titre", value = evenement$Titre)
      updateSelectInput(session, "type_evenement", selected = ifelse(is.na(evenement$Type_Evenement) || evenement$Type_Evenement == "", "", evenement$Type_Evenement))
      updateSelectInput(session, "categorie", selected = ifelse(is.na(evenement$Categorie) || evenement$Categorie == "", "", evenement$Categorie))
      updateTextInput(session, "sources", value = evenement$Sources)
      updateTextInput(session, "links", value = evenement$Links)
      updateSelectInput(session, "statut", selected = evenement$Statut)
      updateDateInput(session, "date", value = as.Date(evenement$Date))
      periode <- strsplit(evenement$Periode_Rapport, " à ")[[1]]
      if (length(periode) == 2) {
        updateDateRangeInput(session, "periode", start = as.Date(periode[1]), end = as.Date(periode[2]))
      }
      updateTextAreaInput(session, "situation", value = evenement$Situation)
      updateTextAreaInput(session, "evaluation_risque", value = evenement$Evaluation_Risque)
      updateTextAreaInput(session, "mesures", value = evenement$Mesures)
      updateSelectInput(session, "pertinence_sante", selected = ifelse(is.na(evenement$Pertinence_Sante) || evenement$Pertinence_Sante == "", "", evenement$Pertinence_Sante))
      updateSelectInput(session, "pertinence_voyageurs", selected = ifelse(is.na(evenement$Pertinence_Voyageurs) || evenement$Pertinence_Voyageurs == "", "", evenement$Pertinence_Voyageurs))
      updateTextInput(session, "commentaires", value = evenement$Commentaires)
      updateTextAreaInput(session, "resume", value = evenement$Resume)
      updateTextInput(session, "autre", value = evenement$Autre)
      session$userData$indice_edition <- which(df$ID_Evenement == evenement$ID_Evenement)
    }
  })


  output$tableau_evenements_actifs <- renderDT({
    df <- donnees() %>% filter(Statut == "actif")
    if (!is.null(input$filtre_date[1]) && !is.null(input$filtre_date[2])) {
      df <- df %>% filter(Date >= input$filtre_date[1], Date <= input$filtre_date[2])
    }
    datatable(df, selection = "multiple", editable = TRUE)
  })

  output$tableau_evenements_termines <- renderDT({
    df <- donnees() %>% filter(Statut == "terminé")
    if (!is.null(input$filtre_date[1]) && !is.null(input$filtre_date[2])) {
      df <- df %>% filter(Date >= input$filtre_date[1], Date <= input$filtre_date[2])
    }
    datatable(df, selection = "multiple", editable = TRUE)
  })
}