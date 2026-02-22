pacman::p_load(
  shiny,
  shinyjs,
  DT,
  officer,
  flextable,
  readr,
  dplyr,
  writexl,
  magick,
  lubridate,
  stringr
)

# ---- lien pour enregistrer les données ----
fichier_donnees <- "evenements_data.csv"

server <- function(input, output, session) {
  useShinyjs()
  
  # =========================================================
  # SOURCE DE DONNÉES (CSV) - SYNCHRONISATION MULTI-CLIENTS
  # =========================================================
  donnees <- reactivePoll(
    intervalMillis = 2000L,
    session,
    checkFunc = function() {
      fi <- file.info(fichier_donnees)
      if (is.na(fi$mtime)) return(0)
      paste(fi$mtime, fi$size, sep = "_")
    },
    valueFunc = function() {
      tryCatch({
        df_lu <- readr::read_csv(
          fichier_donnees, show_col_types = FALSE,
          col_types = cols(
            Version = col_double(),
            Date = col_date(),
            Derniere_Modification = col_datetime()
          )
        )
        if (!"Images" %in% names(df_lu)) df_lu$Images <- NA_character_
        df_lu
      }, error = function(e) {
        data.frame()
      })
    }
  )
  
  # =========================================================
  # OUTIL : VIDER LE FORMULAIRE
  # =========================================================
  vider_formulaire <- function() {
    updateTextInput(session, "titre", value = "")
    updateSelectInput(session, "type_evenement", selected = "")
    updateSelectInput(session, "categorie", selected = "")
    updateTextInput(session, "regions", value = "")
    updateTextInput(session, "sources", value = "")
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
  }
  
  # --- Bouton : zoom tableau (affiche/masque le panneau latéral) ---
  observeEvent(input$zoom_table, {
    shinyjs::toggleClass(selector = "body", class = "zoom-dt")
    shinyjs::runjs("setTimeout(function(){ $(window).trigger('resize'); }, 150);")
  }, ignoreInit = TRUE)
  
  # --- Données EIOS en mémoire (onglet EIOS) ---
  eios_donnees <- reactiveVal(
    data.frame(
      Sources   = character(),
      Titre     = character(),
      Situation = character(),
      Autre     = character(),
      `Régions` = character(),
      EIOS_id   = character(),
      stringsAsFactors = FALSE
    )
  )
  
  # --- Dossiers & utilitaires pour les images téléversées ---
  dir.create("Graphs_uploaded", showWarnings = FALSE)
  
  # fonction : récupérer l'ID de l'événement courant (édition prioritaire, sinon sélection)
  id_evenement_courant <- function() {
    df_full <- donnees()
    
    if (!is.null(session$userData$indice_edition)) {
      idx <- session$userData$indice_edition
      if (is.finite(idx) && idx <= nrow(df_full)) {
        return(df_full$ID_Evenement[idx])
      }
    }
    
    onglet <- input$onglets
    if (identical(onglet, "Actifs")) {
      sel <- input$tableau_evenements_actifs_rows_selected
      if (length(sel) == 1) return(donnees_actifs_filtrees()$ID_Evenement[sel])
    } else if (identical(onglet, "Terminés")) {
      sel <- input$tableau_evenements_termines_rows_selected
      if (length(sel) == 1) return(donnees_termines_filtrees()$ID_Evenement[sel])
    }
    
    return(NA_character_)
  }
  
  prochain_index_image <- function(id_evt) {
    dir_evt <- file.path("Graphs_uploaded", id_evt)
    if (!dir.exists(dir_evt)) return(1L)
    exist <- list.files(dir_evt, pattern = "^image_[0-9]+\\.", full.names = FALSE)
    if (length(exist) == 0) return(1L)
    nums <- suppressWarnings(as.integer(gsub("^image_([0-9]+)\\..*$", "\\1", exist)))
    max(nums, na.rm = TRUE) + 1L
  }
  
  taille_img <- function(path, largeur_cible = 6) {
    info <- tryCatch(magick::image_read(path), error = function(e) NULL)
    if (is.null(info)) return(list(w = largeur_cible, h = 4.5))
    dim <- magick::image_info(info)
    h <- (dim$height / dim$width) * largeur_cible
    if (!is.finite(h) || h <= 0) h <- 4.5
    list(w = largeur_cible, h = h)
  }
  
  # =========================================================
  # FILTRAGE / TRI
  # =========================================================
  donnees_actifs_filtrees <- reactive({
    df <- donnees()
    if (nrow(df) == 0) return(df)
    df <- dplyr::filter(df, Statut == "actif")
    if (!is.null(input$filtre_date[1]) && !is.null(input$filtre_date[2])) {
      df <- dplyr::filter(df, Date >= input$filtre_date[1], Date <= input$filtre_date[2])
    }
    df <- dplyr::arrange(df, dplyr::desc(Date), dplyr::desc(Derniere_Modification))
    filtrer_par_texte(df)
  })
  
  donnees_termines_filtrees <- reactive({
    df <- donnees()
    if (nrow(df) == 0) return(df)
    df <- dplyr::filter(df, Statut == "terminé")
    if (!is.null(input$filtre_date[1]) && !is.null(input$filtre_date[2])) {
      df <- dplyr::filter(df, Date >= input$filtre_date[1], Date <= input$filtre_date[2])
    }
    df <- dplyr::arrange(df, dplyr::desc(Date), dplyr::desc(Derniere_Modification))
    filtrer_par_texte(df)
  })
  
  filtrer_par_texte <- function(df) {
    terme <- input$texte_filtre
    champ <- input$champ_filtre
    
    if (is.null(terme) || !nzchar(trimws(terme)) || nrow(df) == 0) {
      return(df)
    }
    
    terme <- trimws(terme)
    cols_tous <- c("Titre","Sources","Situation","Autre","Régions","EIOS_id")
    
    if (identical(champ, "tous")) {
      cols_existe <- intersect(cols_tous, colnames(df))
      if (length(cols_existe) == 0) return(df)
      idx <- apply(
        as.data.frame(df[, cols_existe, drop = FALSE], stringsAsFactors = FALSE),
        1,
        function(ligne) any(grepl(terme, ligne, ignore.case = TRUE))
      )
      df[idx, , drop = FALSE]
    } else {
      if (!champ %in% colnames(df)) return(df)
      df[grepl(terme, df[[champ]], ignore.case = TRUE), , drop = FALSE]
    }
  }
  
  # =========================================================
  # ÉTAT ÉDITION
  # =========================================================
  session$userData$indice_edition <- NULL
  
  # =========================================================
  # AJOUT / MODIFICATION / DUPLICATION
  # =========================================================
  observeEvent(input$enregistrer, {
    nouvelle_id <- sprintf(
      "EVT_%s_%03d",
      format(Sys.Date(), "%Y%m%d"),
      sum(grepl(format(Sys.Date(), "%Y%m%d"), donnees()$ID_Evenement)) + 1
    )
    
    nouvel_evenement <- data.frame(
      ID_Evenement = nouvelle_id,
      Update = "Nouvel événement",
      Version = 1,
      Date = input$date,
      Derniere_Modification = force_tz(Sys.time(), "Africa/Casablanca"),
      Titre = input$titre,
      Type_Evenement = input$type_evenement,
      Categorie = input$categorie,
      `Régions` = input$regions,
      Sources = input$sources,
      Situation = input$situation,
      Evaluation_Risque = input$evaluation_risque,
      Mesures = input$mesures,
      Pertinence_Sante = input$pertinence_sante,
      Pertinence_Voyageurs = input$pertinence_voyageurs,
      Commentaires = input$commentaires,
      Resume = input$resume,
      Autre = input$autre,
      Statut = input$statut,
      EIOS_id = NA_character_,
      stringsAsFactors = FALSE
    )
    
    tryCatch({
      write_csv(dplyr::bind_rows(donnees(), nouvel_evenement), fichier_donnees)
      vider_formulaire()
      session$userData$indice_edition <- NULL
    }, error = function(e) {
      showModal(modalDialog("Erreur lors de l'enregistrement : ", e$message))
      return()
    })
    
    showModal(modalDialog(
      title = "Succès",
      "L'événement a été ajouté avec succès.",
      easyClose = TRUE,
      footer = modalButton("Fermer")
    ))
  })
  
  observeEvent(input$effacer_masque, {
    vider_formulaire()
    session$userData$indice_edition <- NULL
  })
  
  observeEvent(input$enregistrer_nouveau, {
    indice <- session$userData$indice_edition
    
    if (is.null(indice)) {
      showModal(modalDialog("Veuillez d’abord charger un événement à éditer via le bouton « Load for editing »."))
      return()
    }
    
    donnees_actuelles <- donnees()
    evenement_original <- donnees_actuelles[indice, ]
    
    nouvelle_id <- sprintf(
      "EVT_%s_%03d",
      format(Sys.Date(), "%Y%m%d"),
      sum(grepl(format(Sys.Date(), "%Y%m%d"), donnees_actuelles$ID_Evenement)) + 1
    )
    
    original_id <- evenement_original$ID_Evenement[1]
    
    nouvel_evenement <- evenement_original %>%
      mutate(
        ID_Evenement = nouvelle_id,
        Update = paste0("Update of event: ", original_id),
        Version = evenement_original$Version + 1,
        Derniere_Modification = force_tz(Sys.time(), "Africa/Casablanca"),
        Titre = input$titre,
        Type_Evenement = input$type_evenement,
        Categorie = input$categorie,
        `Régions` = input$regions,
        Sources = input$sources,
        Date = input$date,
        Situation = input$situation,
        Evaluation_Risque = input$evaluation_risque,
        Mesures = input$mesures,
        Pertinence_Sante = input$pertinence_sante,
        Pertinence_Voyageurs = input$pertinence_voyageurs,
        Commentaires = input$commentaires,
        Resume = input$resume,
        Autre = input$autre,
        Statut = input$statut
      ) %>%
      select(ID_Evenement, Update, everything())
    
    tryCatch({
      write_csv(dplyr::bind_rows(donnees_actuelles, nouvel_evenement), fichier_donnees)
      vider_formulaire()
      session$userData$indice_edition <- NULL
    }, error = function(e) {
      showModal(modalDialog("Erreur lors de l'enregistrement : ", e$message))
      return()
    })
    
    showModal(modalDialog(
      title = "Succès",
      "Le nouvel événement a été enregistré avec succès.",
      easyClose = TRUE,
      footer = modalButton("Fermer")
    ))
  })
  
  observeEvent(input$mettre_a_jour, {
    indice <- session$userData$indice_edition
    if (is.null(indice)) return()
    
    donnees_actuelles <- donnees()
    
    if (indice > nrow(donnees_actuelles)) {
      showModal(modalDialog("Erreur : l'élément à mettre à jour n'existe plus."))
      return()
    }
    
    donnees_actuelles[indice, ] <- donnees_actuelles[indice, ] %>%
      mutate(
        Titre = input$titre,
        Type_Evenement = input$type_evenement,
        Categorie = input$categorie,
        `Régions` = input$regions,
        Sources = input$sources,
        Date = input$date,
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
    
    tryCatch({
      write_csv(donnees_actuelles, fichier_donnees)
      vider_formulaire()
      session$userData$indice_edition <- NULL
    }, error = function(e) {
      showModal(modalDialog("Erreur lors de la mise à jour : ", e$message))
      return()
    })
    
    showModal(modalDialog(
      title = "Succès",
      "Les modifications ont été enregistrées.",
      easyClose = TRUE,
      footer = modalButton("Fermer")
    ))
  })
  
  observeEvent(input$count_selected_excel, {
    selected <- input$tableau_evenements_actifs_rows_selected
    session$sendCustomMessage("show_export_count", length(selected))
  })
  
  observeEvent(input$count_selected_word, {
    selected <- input$tableau_evenements_actifs_rows_selected
    session$sendCustomMessage("show_export_count", length(selected))
  })
  
  # =========================================================
  # CHARGER POUR ÉDITION
  # =========================================================
  observeEvent(input$editer, {
    onglet <- input$onglets
    df_full <- donnees()
    
    if (identical(onglet, "Actifs")) {
      selection <- input$tableau_evenements_actifs_rows_selected
      df_view <- donnees_actifs_filtrees()
    } else if (identical(onglet, "Terminés")) {
      selection <- input$tableau_evenements_termines_rows_selected
      df_view <- donnees_termines_filtrees()
    } else {
      showModal(modalDialog("Veuillez sélectionner un événement dans l’onglet « Actifs » ou « Terminés »."))
      return()
    }
    
    if (length(selection) != 1) {
      showModal(modalDialog("Veuillez sélectionner exactement un (1) événement dans l’onglet courant."))
      return()
    }
    
    evenement <- df_view[selection, ]
    idx <- which(df_full$ID_Evenement == evenement$ID_Evenement)
    if (length(idx) != 1) {
      showModal(modalDialog("Impossible de trouver l’événement sélectionné dans le jeu de données complet."))
      return()
    }
    
    updateTextInput(session, "titre", value = evenement$Titre)
    updateSelectInput(session, "type_evenement",
                      selected = ifelse(is.na(evenement$Type_Evenement) || evenement$Type_Evenement == "", "", evenement$Type_Evenement))
    updateSelectInput(session, "categorie",
                      selected = ifelse(is.na(evenement$Categorie) || evenement$Categorie == "", "", evenement$Categorie))
    updateSelectInput(session, "regions", selected = evenement$`Régions`)
    updateTextInput(session, "sources", value = evenement$Sources)
    updateSelectInput(session, "statut", selected = evenement$Statut)
    updateDateInput(session, "date", value = as.Date(evenement$Date))
    updateTextAreaInput(session, "situation", value = evenement$Situation)
    updateTextAreaInput(session, "evaluation_risque", value = evenement$Evaluation_Risque)
    updateTextAreaInput(session, "mesures", value = evenement$Mesures)
    updateSelectInput(session, "pertinence_sante",
                      selected = ifelse(is.na(evenement$Pertinence_Sante) || evenement$Pertinence_Sante == "", "", evenement$Pertinence_Sante))
    updateSelectInput(session, "pertinence_voyageurs",
                      selected = ifelse(is.na(evenement$Pertinence_Voyageurs) || evenement$Pertinence_Voyageurs == "", "", evenement$Pertinence_Voyageurs))
    updateTextInput(session, "commentaires", value = evenement$Commentaires)
    updateTextAreaInput(session, "resume", value = evenement$Resume)
    updateTextInput(session, "autre", value = evenement$Autre)
    
    session$userData$indice_edition <- idx
  })
  
  # =========================================================
  # SUPPRESSION : suppression par ID_Evenement + nettoyage formulaire
  # =========================================================
  observeEvent(input$supprimer, {
    onglet <- input$onglets
    
    if (identical(onglet, "Actifs")) {
      
      sel <- shiny::isolate(input$tableau_evenements_actifs_rows_selected)
      df_view <- shiny::isolate(donnees_actifs_filtrees())
      
      if (is.null(sel) || length(sel) == 0) {
        showModal(modalDialog("Aucun événement sélectionné dans l’onglet « Actifs »."))
        return()
      }
      
      sel <- sel[is.finite(sel)]
      sel <- sel[sel >= 1 & sel <= nrow(df_view)]
      if (length(sel) == 0) {
        showModal(modalDialog("La sélection n’est plus valide (mise à jour du tableau). Veuillez re-sélectionner."))
        return()
      }
      
      ids <- unique(as.character(df_view$ID_Evenement[sel]))
      ids <- ids[!is.na(ids) & nzchar(ids)]
      
      if (length(ids) == 0) {
        showModal(modalDialog("Impossible d’identifier les événements à supprimer."))
        return()
      }
      
      session$userData$ids_a_supprimer <- ids
      session$userData$source_suppression <- "csv"
      
      showModal(modalDialog(
        title = "Confirmation",
        paste("Êtes-vous sûr de vouloir supprimer", length(ids), "événement(s) ?"),
        footer = tagList(
          modalButton("Annuler"),
          actionButton("confirm_delete", "Confirmer", class = "btn btn-danger")
        )
      ))
      
    } else if (identical(onglet, "Terminés")) {
      
      sel <- shiny::isolate(input$tableau_evenements_termines_rows_selected)
      df_view <- shiny::isolate(donnees_termines_filtrees())
      
      if (is.null(sel) || length(sel) == 0) {
        showModal(modalDialog("Aucun événement sélectionné dans l’onglet « Terminés »."))
        return()
      }
      
      sel <- sel[is.finite(sel)]
      sel <- sel[sel >= 1 & sel <= nrow(df_view)]
      if (length(sel) == 0) {
        showModal(modalDialog("La sélection n’est plus valide (mise à jour du tableau). Veuillez re-sélectionner."))
        return()
      }
      
      ids <- unique(as.character(df_view$ID_Evenement[sel]))
      ids <- ids[!is.na(ids) & nzchar(ids)]
      
      if (length(ids) == 0) {
        showModal(modalDialog("Impossible d’identifier les événements à supprimer."))
        return()
      }
      
      session$userData$ids_a_supprimer <- ids
      session$userData$source_suppression <- "csv"
      
      showModal(modalDialog(
        title = "Confirmation",
        paste("Êtes-vous sûr de vouloir supprimer", length(ids), "événement(s) ?"),
        footer = tagList(
          modalButton("Annuler"),
          actionButton("confirm_delete", "Confirmer", class = "btn btn-danger")
        )
      ))
      
    } else if (identical(onglet, "EIOS")) {
      
      sel <- shiny::isolate(input$tableau_eios_rows_selected)
      if (is.null(sel) || length(sel) == 0) {
        showModal(modalDialog("Aucun élément sélectionné dans l’onglet « EIOS »."))
        return()
      }
      
      session$userData$rows_eios_a_supprimer <- sel
      session$userData$source_suppression <- "eios"
      
      showModal(modalDialog(
        title = "Confirmation",
        paste("Êtes-vous sûr de vouloir supprimer", length(sel), "ligne(s) de l’onglet EIOS ?"),
        footer = tagList(
          modalButton("Annuler"),
          actionButton("confirm_delete", "Confirmer", class = "btn btn-danger")
        )
      ))
    }
  })
  
  observeEvent(input$confirm_delete, {
    removeModal()
    source <- session$userData$source_suppression
    
    if (identical(source, "csv")) {
      ids <- unique(session$userData$ids_a_supprimer)
      if (is.null(ids) || length(ids) == 0) return()
      
      df <- donnees()
      if (nrow(df) == 0) {
        showModal(modalDialog("Le fichier ne contient aucune donnée à supprimer."))
        return()
      }
      
      df_new <- df[!df$ID_Evenement %in% ids, , drop = FALSE]
      
      tryCatch({
        readr::write_csv(df_new, fichier_donnees)
        vider_formulaire()
        session$userData$indice_edition <- NULL
        session$userData$ids_a_supprimer <- NULL
        session$userData$source_suppression <- NULL
        
        showModal(modalDialog(
          title = "Succès",
          paste(length(ids), "événement(s) ont été supprimé(s)."),
          easyClose = TRUE,
          footer = modalButton("Fermer")
        ))
      }, error = function(e) {
        showModal(modalDialog("Erreur lors de la suppression : ", e$message))
      })
      
    } else if (identical(source, "eios")) {
      sel <- session$userData$rows_eios_a_supprimer
      if (is.null(sel) || length(sel) == 0) return()
      
      df_e <- eios_donnees()
      if (nrow(df_e) == 0) return()
      
      sel <- sort(unique(sel))
      sel <- sel[sel >= 1 & sel <= nrow(df_e)]
      if (length(sel) == 0) return()
      
      df_e_new <- df_e[-sel, , drop = FALSE]
      eios_donnees(df_e_new)
      
      session$userData$rows_eios_a_supprimer <- NULL
      session$userData$source_suppression <- NULL
      
      showModal(modalDialog(
        title = "Succès",
        paste(length(sel), "ligne(s) ont été supprimée(s) de l’onglet EIOS."),
        easyClose = TRUE,
        footer = modalButton("Fermer")
      ))
    }
  })

  # =========================================================
  # TABLES
  # =========================================================
  output$tableau_evenements_actifs <- DT::renderDT({
    df <- donnees_actifs_filtrees() %>%
      dplyr::mutate(across(where(is.character), ~{
        x <- .
        x[is.na(x)] <- ""
        sprintf('<div class="cell-fixed" title="%s">%s</div>', x, x)
      }))
    
    cols <- colnames(df)
    idx <- function(nm) which(cols == nm) - 1L
    
    DT::datatable(
      df,
      selection = "multiple",
      editable = TRUE,
      escape = FALSE,
      rownames = FALSE,
      options = list(
        scrollX = TRUE,
        autoWidth = TRUE,
        pageLength = 15,
        columnDefs = list(
          list(targets = idx("Sources"), width = "250px"),
          list(targets = idx("Situation"), width = "500px"),
          list(targets = idx("Evaluation_Risque"), width = "500px"),
          list(targets = idx("Mesures"), width = "500px")
        )
      )
    )
  })
  
  output$tableau_evenements_termines <- DT::renderDT({
    df <- donnees_termines_filtrees() %>%
      dplyr::mutate(across(where(is.character), ~{
        x <- .
        x[is.na(x)] <- ""
        sprintf('<div class="cell-fixed" title="%s">%s</div>', x, x)
      }))
    
    cols <- colnames(df)
    idx <- function(nm) which(cols == nm) - 1L
    
    DT::datatable(
      df,
      selection = "multiple",
      editable = TRUE,
      escape = FALSE,
      rownames = FALSE,
      options = list(
        scrollX = TRUE,
        autoWidth = TRUE,
        pageLength = 15,
        columnDefs = list(
          list(targets = idx("Sources"), width = "250px"),
          list(targets = idx("Situation"), width = "500px"),
          list(targets = idx("Evaluation_Risque"), width = "500px"),
          list(targets = idx("Mesures"), width = "500px")
        )
      )
    )
  })
  
  output$tableau_eios <- DT::renderDT({
    df <- eios_donnees() %>%
      dplyr::mutate(across(where(is.character), ~{
        x <- .
        x[is.na(x)] <- ""
        sprintf('<div class="cell-fixed" title="%s">%s</div>', x, x)
      }))
    
    DT::datatable(
      df,
      selection = "multiple",
      editable = FALSE,
      escape = FALSE,
      rownames = FALSE,
      options = list(
        scrollX = TRUE,
        autoWidth = TRUE,
        pageLength = 15
      )
    )
  })
  
  # =========================================================
  # PROXYS + BOUTONS "SÉLECTIONNER TOUT" / "EFFACER"
  # =========================================================
  proxy_actifs <- DT::dataTableProxy("tableau_evenements_actifs")
  proxy_termines <- DT::dataTableProxy("tableau_evenements_termines")
  proxy_eios <- DT::dataTableProxy("tableau_eios")
  
  .selectionner_tout <- function(proxy, nrows) {
    if (is.finite(nrows) && nrows > 0) DT::selectRows(proxy, 1:nrows) else DT::selectRows(proxy, NULL)
  }
  .effacer_selection <- function(proxy) DT::selectRows(proxy, NULL)
  
  observeEvent(input$select_all, {
    onglet <- input$onglets
    if (identical(onglet, "Actifs")) {
      .selectionner_tout(proxy_actifs, nrow(donnees_actifs_filtrees()))
    } else if (identical(onglet, "Terminés")) {
      .selectionner_tout(proxy_termines, nrow(donnees_termines_filtrees()))
    } else if (identical(onglet, "EIOS")) {
      .selectionner_tout(proxy_eios, nrow(eios_donnees()))
    }
  }, ignoreInit = TRUE)
  
  observeEvent(input$clear_all, {
    onglet <- input$onglets
    if (identical(onglet, "Actifs")) {
      .effacer_selection(proxy_actifs)
    } else if (identical(onglet, "Terminés")) {
      .effacer_selection(proxy_termines)
    } else if (identical(onglet, "EIOS")) {
      .effacer_selection(proxy_eios)
    }
  }, ignoreInit = TRUE)
  
  # =========================================================
  # FILTRE DATE : RÉINITIALISATION
  # =========================================================
  observeEvent(input$reset_filtre_date, {
    df <- donnees()
    if (!is.null(df) && nrow(df) > 0 && "Date" %in% names(df)) {
      min_date <- suppressWarnings(min(df$Date, na.rm = TRUE))
      if (is.infinite(suppressWarnings(as.numeric(min_date))) || is.na(min_date)) {
        min_date <- Sys.Date() - 30
      }
    } else {
      min_date <- Sys.Date() - 30
    }
    updateDateRangeInput(session, "filtre_date", start = as.Date(min_date), end = Sys.Date())
  })
  
  # =========================================================
  # EXPORT EXCEL
  # =========================================================
  dir.create("excel_outputs", showWarnings = FALSE)
  
  observeEvent(input$exporter_excel, {
    dir.create("excel_outputs", showWarnings = FALSE)
    
    data_view <- donnees_actifs_filtrees()
    sel <- input$tableau_evenements_actifs_rows_selected
    export_data <- if (!is.null(sel) && length(sel) > 0) data_view[sel, ] else data_view
    
    if (nrow(export_data) == 0) {
      showModal(modalDialog("Aucun événement sélectionné ou actif."))
      return()
    }
    
    export_data <- export_data %>%
      mutate(Derniere_Modification = format(Derniere_Modification, "%Y-%m-%d %H:%M:%S"))
    
    date_today <- format(Sys.Date(), "%Y-%m-%d")
    existing_files <- list.files("excel_outputs", pattern = paste0("^evenements_", date_today, "_\\d+\\.xlsx$"))
    existing_nums <- as.integer(gsub(paste0("evenements_", date_today, "_|\\.xlsx"), "", existing_files))
    next_num <- ifelse(length(existing_nums) == 0, 1, max(existing_nums, na.rm = TRUE) + 1)
    filename <- sprintf("excel_outputs/evenements_%s_%d.xlsx", date_today, next_num)
    
    tryCatch({
      writexl::write_xlsx(export_data, filename)
      showModal(modalDialog(
        title = "Succès",
        paste("Le fichier Excel a été créé avec succès :", filename),
        easyClose = TRUE,
        footer = modalButton("Fermer")
      ))
    }, error = function(e) {
      showModal(modalDialog("Erreur lors de l'exportation : ", e$message))
    })
  })
  
  # =========================================================
  # EIOS : IMPORT + COPIE VERS ACTIFS
  # =========================================================
  observeEvent(input$upload_eios, {
    if (is.null(input$upload_eios$datapath)) return()
    path <- input$upload_eios$datapath
    
    df_raw <- tryCatch(
      readr::read_csv(path, show_col_types = FALSE),
      error = function(e) { showModal(modalDialog("Erreur de lecture CSV EIOS : ", e$message)); return(NULL) }
    )
    if (is.null(df_raw)) return()
    
    req_cols <- c("externalLink","title","summary","categories","mentionedCountries","id")
    if (!all(req_cols %in% names(df_raw))) {
      manq <- paste(setdiff(req_cols, names(df_raw)), collapse = ", ")
      showModal(modalDialog(paste("Colonnes manquantes dans le CSV EIOS :", manq)))
      return()
    }
    
    df_eios <- df_raw %>%
      dplyr::transmute(
        Sources   = as.character(.data$externalLink),
        Titre     = as.character(.data$title),
        Situation = as.character(.data$summary),
        Autre     = as.character(.data$categories),
        `Régions` = as.character(.data$mentionedCountries),
        EIOS_id   = as.character(.data$id)
      )
    
    eios_donnees(df_eios)
  })
  
  observeEvent(input$copier_vers_actifs, {
    df_eios <- eios_donnees()
    sel <- input$tableau_eios_rows_selected
    if (is.null(sel) || length(sel) == 0) return()
    
    selection <- df_eios[sel, , drop = FALSE]
    if (nrow(selection) == 0) return()
    
    df_all <- donnees()
    
    to_add <- lapply(seq_len(nrow(selection)), function(i) {
      nouvelle_id <- sprintf(
        "EVT_%s_%03d",
        format(Sys.Date(), "%Y%m%d"),
        sum(grepl(format(Sys.Date(), "%Y%m%d"), df_all$ID_Evenement)) + i
      )
      
      data.frame(
        ID_Evenement          = nouvelle_id,
        Update                = "new event from EIOS",
        Version               = 1,
        Date                  = Sys.Date(),
        Derniere_Modification = force_tz(Sys.time(), "Africa/Casablanca"),
        Titre                 = selection$Titre[i],
        Type_Evenement        = NA_character_,
        Categorie             = NA_character_,
        `Régions`             = selection$`Régions`[i],
        Sources               = selection$Sources[i],
        Situation             = selection$Situation[i],
        Evaluation_Risque     = NA_character_,
        Mesures               = NA_character_,
        Pertinence_Sante      = NA_character_,
        Pertinence_Voyageurs  = NA_character_,
        Commentaires          = NA_character_,
        Resume                = NA_character_,
        Autre                 = selection$Autre[i],
        Statut                = "actif",
        EIOS_id               = selection$EIOS_id[i],
        Images                = if ("Images" %in% names(df_all)) NA_character_ else NULL,
        stringsAsFactors      = FALSE
      )
    })
    
    df_new <- dplyr::bind_rows(df_all, dplyr::bind_rows(to_add))
    
    tryCatch({
      write_csv(df_new, fichier_donnees)
    }, error = function(e) {
      showModal(modalDialog("Erreur lors de l'enregistrement CSV : ", e$message))
    })
  })
  
  # =========================================================
  # EXPORT WORD + IMAGES
  # =========================================================
  observeEvent(input$exporter_word, {
    dir.create("rapport_word_outputs", showWarnings = FALSE)
    
    data_view <- donnees_actifs_filtrees()
    sel <- input$tableau_evenements_actifs_rows_selected
    export_data <- if (!is.null(sel) && length(sel) > 0) data_view[sel, ] else data_view
    
    if (nrow(export_data) == 0) {
      showModal(modalDialog("Aucun événement sélectionné ou actif."))
      return()
    }
    
    rapport_filename <- generer_nom_fichier()
    
    doc <- officer::read_docx("template_rapport/template_word.docx")
    doc <- ajouter_en_tete(doc)
    doc <- ajouter_sources(doc, export_data)
    doc <- officer::body_add_break(doc)
    doc <- ajouter_sommaire(doc, export_data, NULL)
    doc <- ajouter_evenements(doc, export_data, NULL)
    
    tryCatch({
      print(doc, target = rapport_filename)
      showModal(modalDialog(
        title = "Succès",
        paste("Le rapport Word a été généré avec succès :", rapport_filename),
        easyClose = TRUE,
        footer = modalButton("Fermer")
      ))
    }, error = function(e) {
      showModal(modalDialog("Erreur lors de la génération du Word : ", e$message))
    })
  })
  
  observeEvent(input$televerser_graphs, {
    fichiers <- input$televerser_graphs
    if (is.null(fichiers) || nrow(fichiers) == 0) {
      showModal(modalDialog("Veuillez choisir un ou plusieurs fichiers image."))
      return()
    }
    
    id_evt <- id_evenement_courant()
    if (is.na(id_evt) || is.null(id_evt)) {
      showModal(modalDialog("Veuillez d’abord sélectionner un événement (dans le tableau) ou le charger pour édition."))
      return()
    }
    
    dir_evt <- file.path("Graphs_uploaded", id_evt)
    dir.create(dir_evt, showWarnings = FALSE, recursive = TRUE)
    
    idx <- prochain_index_image(id_evt)
    chemins_relatifs <- character(0)
    
    for (i in seq_len(nrow(fichiers))) {
      src <- fichiers$datapath[i]
      ext <- tools::file_ext(fichiers$name[i])
      if (identical(ext, "")) ext <- "png"
      dest_name <- sprintf("image_%d.%s", idx, tolower(ext))
      dest_path <- file.path(dir_evt, dest_name)
      ok <- file.copy(from = src, to = dest_path, overwrite = FALSE)
      if (isTRUE(ok)) {
        chemins_relatifs <- c(chemins_relatifs, dest_path)
        idx <- idx + 1L
      }
    }
    
    if (length(chemins_relatifs) == 0) {
      showModal(modalDialog("Aucun fichier n’a été copié."))
      return()
    }
    
    df <- donnees()
    pos <- which(df$ID_Evenement == id_evt)
    if (length(pos) == 1) {
      exist <- df$Images[pos]
      if (is.null(exist) || is.na(exist)) exist <- ""
      ajout <- paste(chemins_relatifs, collapse = ";")
      df$Images[pos] <- if (nzchar(exist)) paste(exist, ajout, sep = ";") else ajout
      
      tryCatch(
        write_csv(df, fichier_donnees),
        error = function(e) showModal(modalDialog("Erreur lors de l'enregistrement des chemins d’images : ", e$message))
      )
    }
    
    showModal(modalDialog(
      title = "Succès",
      paste0(length(chemins_relatifs), " fichier(s) lié(s) à l’événement ", id_evt, "."),
      easyClose = TRUE,
      footer = modalButton("Fermer")
    ))
  })
  
  generer_nom_fichier <- function() {
    date_today <- format(Sys.Date(), "%Y-%m-%d")
    existing_files <- list.files("rapport_word_outputs", pattern = paste0("^rapport_", date_today, "_\\d+\\.docx$"))
    existing_nums <- as.integer(gsub(paste0("rapport_", date_today, "_|\\.docx"), "", existing_files))
    next_num <- ifelse(length(existing_nums) == 0, 1, max(existing_nums, na.rm = TRUE) + 1)
    sprintf("rapport_word_outputs/rapport_%s_%d.docx", date_today, next_num)
  }
  
  # =========================================================
  # WORD : PAGE 1
  # =========================================================
  ajouter_en_tete <- function(doc) {
    date_ref <- Sys.Date() - lubridate::weeks(1)
    lundi <- lubridate::floor_date(date_ref, unit = "week", week_start = 1)
    dimanche <- lundi + lubridate::days(6)
    
    num_semaine <- lubridate::isoweek(lundi)
    an_iso <- lubridate::isoyear(lundi)
    
    periode_str <- paste0(
      format(lundi, "%d"), "-",
      format(dimanche, "%d"), "/",
      format(dimanche, "%m/%Y")
    )
    
    valeur_semaine <- paste0(
      " |N° ", num_semaine, "/", an_iso,
      " | SEMAINE N°", num_semaine, " (", periode_str, ") |"
    )
    
    doc <- officer::body_replace_text_at_bkm(doc, bookmark = "BM_SEMAINE", value = valeur_semaine)
    
    mois_fr <- c("janvier","février","mars","avril","mai","juin",
                 "juillet","août","septembre","octobre","novembre","décembre")
    aujourd_hui <- Sys.Date()
    mois_nom <- mois_fr[as.integer(format(aujourd_hui, "%m"))]
    mois_nom <- tools::toTitleCase(mois_nom)
    
    valeur_date <- paste0(format(aujourd_hui, "%d"), " ", mois_nom, " ", format(aujourd_hui, "%Y"))
    doc <- officer::body_replace_text_at_bkm(doc, bookmark = "BM_DATE", value = valeur_date)
    
    doc
  }
  
  ajouter_sources <- function(doc, data) {
    
    parse_sources <- function(df) {
      raw <- df$Sources
      raw <- raw[!is.na(raw) & trimws(raw) != ""]
      if (length(raw) == 0) return(character(0))
      
      tokens <- unlist(strsplit(paste(raw, collapse = "\n"), ";|,|\\n"))
      tokens <- trimws(tokens)
      tokens <- tokens[tokens != ""]
      if (length(tokens) == 0) return(character(0))
      
      names_only <- sub("\\s*\\(.*?\\)\\s*$", "", tokens)
      
      is_url <- grepl("^https?://", names_only, ignore.case = TRUE)
      if (any(is_url)) {
        names_only[is_url] <- sub("^https?://(www\\.)?([^/]+).*$", "\\2", names_only[is_url])
      }
      
      names_only <- gsub("^https?://", "", names_only, ignore.case = TRUE)
      names_only <- gsub("^www\\.", "", names_only, ignore.case = TRUE)
      names_only <- sub("/.*$", "", names_only)
      
      names_only <- tools::toTitleCase(names_only)
      unique(names_only[names_only != ""])
    }
    
    join_sources <- function(vec) {
      if (length(vec) == 0) "" else paste(vec, collapse = "; ")
    }
    
    df_nat <- dplyr::filter(data, Categorie == "Alertes nationales")
    df_intl <- dplyr::filter(data, Categorie == "Alertes internationales")
    df_both <- dplyr::filter(data, Categorie == "Alertes internationales et nationales")
    
    src_nat <- parse_sources(df_nat)
    src_intl <- parse_sources(df_intl)
    src_both <- parse_sources(df_both)
    
    base_style <- officer::fp_text(color = "white", italic = TRUE, font.size = 9, font.family = "Century")
    bold_style <- officer::fp_text(color = "white", italic = TRUE, font.size = 9, font.family = "Century", bold = TRUE)
    head_style <- officer::fp_text(color = "white", italic = TRUE, font.size = 9, font.family = "Century", underlined = TRUE)
    
    head_par <- officer::fpar(officer::ftext("Sources de données de cette semaine : ", head_style))
    
    nat_par <- officer::fpar(
      officer::ftext("• Au niveau national : ", bold_style),
      officer::ftext(join_sources(src_nat), base_style)
    )
    intl_par <- officer::fpar(
      officer::ftext("• Au niveau international : ", bold_style),
      officer::ftext(join_sources(src_intl), base_style)
    )
    both_par <- officer::fpar(
      officer::ftext("• Au niveau national et international : ", bold_style),
      officer::ftext(join_sources(src_both), base_style)
    )
    
    doc <- officer::cursor_bookmark(doc, "BM_SOURCES")
    doc <- officer::body_add_fpar(doc, value = head_par, pos = "on")
    doc <- officer::body_add_fpar(doc, value = nat_par, pos = "after")
    doc <- officer::body_add_fpar(doc, value = intl_par, pos = "after")
    doc <- officer::body_add_fpar(doc, value = both_par, pos = "after")
    
    doc
  }
  
  # =========================================================
  # WORD : PAGE 2
  # =========================================================
  ajouter_sommaire <- function(doc, data, styles) {
    cat_map <- list(
      "Alertes internationales et nationales" = "Alertes et Urgences de Santé Publique au niveau international et national",
      "Alertes nationales" = "Alertes et Urgences de Santé Publique au niveau national",
      "Alertes internationales" = "Alertes et Urgences de Santé Publique au niveau international"
    )
    categories <- names(cat_map)
    
    doc <- doc %>% officer::body_add_par("SOMMAIRE DES ÉVÉNEMENTS", style = "m_heading_1")
    
    compteur <- 1
    for (cat in categories) {
      cat_data <- data %>% dplyr::filter(Categorie == cat)
      if (nrow(cat_data) > 0) {
        doc <- doc %>% officer::body_add_par(cat_map[[cat]], style = "m_heading_2")
        for (i in 1:nrow(cat_data)) {
          doc <- doc %>% officer::body_add_par(paste0(compteur, ". ", cat_data$Titre[i]), style = "m_heading_4")
          compteur <- compteur + 1
        }
      }
    }
    doc
  }
  
  ajouter_evenements <- function(doc, data, styles) {
    cat_map <- list(
      "Alertes internationales et nationales" = "Alertes et Urgences de Santé Publique au niveau international et national",
      "Alertes nationales" = "Alertes et Urgences de Santé Publique au niveau national",
      "Alertes internationales" = "Alertes et Urgences de Santé Publique au niveau international"
    )
    categories <- names(cat_map)
    compteur <- 1
    
    p_std <- officer::fp_par(line_spacing = 0.7)
    run_std <- officer::fp_text()
    run_b <- officer::fp_text(bold = TRUE)
    
    champs <- c("Situation", "Evaluation_Risque", "Mesures")
    labels <- c("Situation épidémiologique", "Évaluation des risques", "Mesures entreprises")
    
    for (cat in categories) {
      cat_data <- data %>% dplyr::filter(Categorie == cat)
      if (nrow(cat_data) == 0) next
      
      for (i in seq_len(nrow(cat_data))) {
        row <- cat_data[i, ]
        
        doc <- doc %>% officer::body_add_par(paste0(compteur, ". ", row$Titre), style = "m_heading_1")
        
        for (j in seq_along(champs)) {
          val <- row[[champs[j]]]
          if (!is.null(val) && !is.na(val) && nzchar(trimws(as.character(val)))) {
            bloc <- officer::fpar(
              officer::ftext(paste0(labels[j], " : "), run_b),
              officer::ftext(as.character(val), run_std),
              fp_p = p_std
            )
            doc <- officer::body_add_fpar(doc, value = bloc, style = "m_standard")
          }
        }
        
        if (!is.null(row$Resume) && !is.na(row$Resume) && nzchar(trimws(as.character(row$Resume)))) {
          doc <- officer::body_add_par(doc, "Résumé", style = "m_resume")
          doc <- officer::body_add_par(doc, as.character(row$Resume), style = "m_resume")
        }
        
        if ("Images" %in% names(row)) {
          imgs_raw <- as.character(row$Images)
          if (!is.na(imgs_raw) && nzchar(trimws(imgs_raw))) {
            chemins <- strsplit(imgs_raw, ";")[[1]]
            chemins <- trimws(chemins)
            chemins <- chemins[chemins != ""]
            for (che in chemins) {
              if (file.exists(che)) {
                s <- taille_img(che, 6)
                doc <- officer::body_add_img(doc, src = che, width = s$w, height = s$h)
                doc <- officer::body_add_par(doc, "", style = "m_standard")
              }
            }
          }
        }
        
        compteur <- compteur + 1
      }
    }
    
    doc
  }
  
}