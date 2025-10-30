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
  stringr,
  renv
)


# ---- lien pour enregistrer les donnes ----
fichier_donnees <- "evenements_data.csv"

# Create folders
dir.create("excel_outputs", showWarnings = FALSE, recursive = TRUE)
dir.create("rapport_word_outputs", showWarnings = FALSE, recursive = TRUE)
dir.create("template_rapport", showWarnings = FALSE, recursive = TRUE)

#
server <- function(input, output, session) {
 
  
  # define datasource (csv)
  donnees <- reactiveVal({
    tryCatch({
      readr::read_csv(fichier_donnees, show_col_types = FALSE,
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
  
  # --- Actifs : filtrer + trier du plus récent au plus ancien ---
  donnees_actifs_filtrees <- reactive({
    df <- donnees()
    if (nrow(df) == 0) return(df)
    df <- dplyr::filter(df, Statut == "actif")
    if (!is.null(input$filtre_date[1]) && !is.null(input$filtre_date[2])) {
      df <- dplyr::filter(df, Date >= input$filtre_date[1], Date <= input$filtre_date[2])
    }
    dplyr::arrange(df, dplyr::desc(Date), dplyr::desc(Derniere_Modification))
  })
  
  # --- Terminés : filtrer + trier du plus récent au plus ancien ---
  donnees_termines_filtrees <- reactive({
    df <- donnees()
    if (nrow(df) == 0) return(df)
    df <- dplyr::filter(df, Statut == "terminé")
    if (!is.null(input$filtre_date[1]) && !is.null(input$filtre_date[2])) {
      df <- dplyr::filter(df, Date >= input$filtre_date[1], Date <= input$filtre_date[2])
    }
    dplyr::arrange(df, dplyr::desc(Date), dplyr::desc(Derniere_Modification))
  })
  
  
  
  
  session$userData$indice_edition <- NULL
  filtre_date_active <- reactiveVal(TRUE)
  
  observeEvent(input$enregistrer, {
    nouvelle_id <- sprintf("EVT_%s_%03d", format(Sys.Date(), "%Y%m%d"),
                           sum(grepl(format(Sys.Date(), "%Y%m%d"), donnees()$ID_Evenement)) + 1)
    
    nouvel_evenement <- data.frame(
      ID_Evenement = nouvelle_id,
      Update = "Nouvel événement",
      Version = 1,
      Date = input$date,
      Derniere_Modification = force_tz(Sys.time(), "Africa/Casablanca"),
      Titre = input$titre,
      Type_Evenement = input$type_evenement,
      Categorie = input$categorie,
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
  
  observeEvent(input$effacer_masque, {
    updateTextInput(session, "titre", value = "")
    updateSelectInput(session, "type_evenement", selected = "")
    updateSelectInput(session, "categorie", selected = "")
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
  })
  
  observeEvent(input$enregistrer_nouveau, {
    indice <- session$userData$indice_edition
    
    if (is.null(indice)) {
      showModal(modalDialog("Veuillez d’abord charger un événement à éditer via le bouton 'Charger pour édition'."))
      return()
    }
    
    donnees_actuelles <- donnees()
    evenement_original <- donnees_actuelles[indice, ]
    
    nouvelle_id <- sprintf("EVT_%s_%03d", format(Sys.Date(), "%Y%m%d"),
                           sum(grepl(format(Sys.Date(), "%Y%m%d"), donnees_actuelles$ID_Evenement)) + 1)
    
    
    original_id <- evenement_original$ID_Evenement[1]
    
    nouvel_evenement <- evenement_original %>%
      mutate(
        ID_Evenement = nouvelle_id,
        Update = paste0("Update of event: ", original_id),
        Version = 1,
        Derniere_Modification = force_tz(Sys.time(), "Africa/Casablanca"),
        Titre = input$titre,
        Type_Evenement = input$type_evenement,
        Categorie = input$categorie,
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
    
    
    donnees(bind_rows(donnees_actuelles, nouvel_evenement))
    
    tryCatch({
      write_csv(donnees(), fichier_donnees)
    }, error = function(e) {
      showModal(modalDialog("Erreur lors de l'enregistrement :", e$message))
    })
    
    showModal(modalDialog(
      title = "Succès",
      "Le nouvel événement a été enregistré avec succès.",
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
    
    donnees(donnees_actuelles)
    write_csv(donnees_actuelles, fichier_donnees)
    session$userData$indice_edition <- NULL
  })
  
  observeEvent(input$editer, {
    onglet <- input$onglets
    df_full <- donnees()
    
    # Ermitteln, aus welchem Tab und welche Zeile ausgewählt ist
    if (identical(onglet, "Actifs")) {
      selection <- input$tableau_evenements_actifs_rows_selected
      df_view   <- donnees_actifs_filtrees()
    } else if (identical(onglet, "Terminés")) {
      selection <- input$tableau_evenements_termines_rows_selected
      df_view   <- donnees_termines_filtrees()
    } else {
      showModal(modalDialog("Aucun onglet actif détecté."))
      return()
    }
    
    if (length(selection) != 1) {
      showModal(modalDialog("Veuillez sélectionner exactement un (1) événement dans l’onglet courant."))
      return()
    }
    
    # Ereignis aus dem sichtbaren (gefilterten) Tab holen
    evenement <- df_view[selection, ]
    
    # Globalen Index im vollständigen Datensatz per ID bestimmen
    idx <- which(df_full$ID_Evenement == evenement$ID_Evenement)
    if (length(idx) != 1) {
      showModal(modalDialog("Impossible de trouver l’événement sélectionné dans le jeu de données complet."))
      return()
    }
    
    # Formularfelder füllen
    updateTextInput(session, "titre", value = evenement$Titre)
    updateSelectInput(session, "type_evenement",
                      selected = ifelse(is.na(evenement$Type_Evenement) || evenement$Type_Evenement == "", "", evenement$Type_Evenement))
    updateSelectInput(session, "categorie",
                      selected = ifelse(is.na(evenement$Categorie) || evenement$Categorie == "", "", evenement$Categorie))
    updateTextInput(session, "sources", value = evenement$Sources)
    updateSelectInput(session, "statut", selected = evenement$Statut)
    updateDateInput(session, "date", value = as.Date(evenement$Date))
    updateTextAreaInput(session, "situation",          value = evenement$Situation)
    updateTextAreaInput(session, "evaluation_risque",  value = evenement$Evaluation_Risque)
    updateTextAreaInput(session, "mesures",            value = evenement$Mesures)
    updateSelectInput(session, "pertinence_sante",
                      selected = ifelse(is.na(evenement$Pertinence_Sante) || evenement$Pertinence_Sante == "", "", evenement$Pertinence_Sante))
    updateSelectInput(session, "pertinence_voyageurs",
                      selected = ifelse(is.na(evenement$Pertinence_Voyageurs) || evenement$Pertinence_Voyageurs == "", "", evenement$Pertinence_Voyageurs))
    updateTextInput(session, "commentaires", value = evenement$Commentaires)
    updateTextAreaInput(session, "resume",    value = evenement$Resume)
    updateTextInput(session, "autre",         value = evenement$Autre)
    
    # Merker für "Mettre à jour" / "Supprimer" / "Enregistrer comme Update"
    session$userData$indice_edition <- idx
  })
  
  
  observeEvent(input$supprimer, {
    indice <- session$userData$indice_edition
    if (is.null(indice)) {
      showModal(modalDialog("Veuillez d’abord charger un événement à supprimer via le bouton 'Charger pour édition'."))
      return()
    }
    
    donnees_actuelles <- donnees()
    
    if (indice > nrow(donnees_actuelles)) {
      showModal(modalDialog("Erreur : l'événement à supprimer n'existe plus."))
      return()
    }
    
    donnees_nouvelles <- donnees_actuelles[-indice, ]
    donnees(donnees_nouvelles)
    
    tryCatch({
      write_csv(donnees_nouvelles, fichier_donnees)
      session$userData$indice_edition <- NULL
      showModal(modalDialog(
        title = "Succès",
        "L'événement a été supprimé avec succès.",
        easyClose = TRUE,
        footer = modalButton("Fermer")
      ))
    }, error = function(e) {
      showModal(modalDialog("Erreur lors de la suppression :", e$message))
    })
  })
  
  
  # ---- actifs ----
  output$tableau_evenements_actifs <- DT::renderDT({
    df <- donnees_actifs_filtrees() %>%
      dplyr::mutate(across(where(is.character), ~{
        x <- .
        x[is.na(x)] <- ""
        sprintf('<div class="cell-fixed" title="%s">%s</div>', x, x)
      }))
    
    
    cols <- colnames(df)
    idx  <- function(nm) which(cols == nm) - 1L  # DataTables: 0-basiert
    
    DT::datatable(
      df,
      selection = "multiple",
      editable  = TRUE,
      escape    = FALSE,
      rownames  = FALSE,        # <<< important : aligne les index des colonnes
      options   = list(
        scrollX    = TRUE,      # <<< permet aux largeurs de s'appliquer
        autoWidth  = TRUE,
        pageLength = 15, 
        columnDefs = list(
          list(targets = idx("Sources"),           width = "250px"),
          list(targets = idx("Situation"),         width = "500px"),
          list(targets = idx("Evaluation_Risque"), width = "500px"),
          list(targets = idx("Mesures"),           width = "500px")
        )
      )
    )
  })
  
  # ---- termines ----
  output$tableau_evenements_termines <- DT::renderDT({
    df <- donnees_termines_filtrees() %>%
      dplyr::mutate(across(where(is.character), ~{
        x <- .
        x[is.na(x)] <- ""
        sprintf('<div class="cell-fixed" title="%s">%s</div>', x, x)
      }))
    
    
    cols <- colnames(df)
    idx  <- function(nm) which(cols == nm) - 1L  # DataTables: 0-basiert
    
    DT::datatable(
      df,
      selection = "multiple",
      editable  = TRUE,
      escape    = FALSE,
      rownames  = FALSE,        # <<< important : aligne les index des colonnes
      options   = list(
        scrollX    = TRUE,      # <<< permet aux largeurs de s'appliquer
        autoWidth  = TRUE,
        pageLength = 15, 
        columnDefs = list(
          list(targets = idx("Sources"),           width = "250px"),
          list(targets = idx("Situation"),         width = "500px"),
          list(targets = idx("Evaluation_Risque"), width = "500px"),
          list(targets = idx("Mesures"),           width = "500px")
        )
      )
    )
  })
  
  
  # --- Proxies for both tables ---
  proxy_actifs   <- DT::dataTableProxy("tableau_evenements_actifs")
  proxy_termines <- DT::dataTableProxy("tableau_evenements_termines")
  
  # --- Helper: select/unselect ---
  .select_all <- function(proxy, nrows) {
    if (is.finite(nrows) && nrows > 0) DT::selectRows(proxy, 1:nrows) else DT::selectRows(proxy, NULL)
  }
  .clear_all <- function(proxy) DT::selectRows(proxy, NULL)
  
  # --- Buttons impactibg activ tab ---
  observeEvent(input$select_all, {
    onglet <- input$onglets
    if (identical(onglet, "Actifs")) {
      .select_all(proxy_actifs, nrow(donnees_actifs_filtrees()))
    } else if (identical(onglet, "Terminés")) {
      .select_all(proxy_termines, nrow(donnees_termines_filtrees()))
    }
  }, ignoreInit = TRUE)
  
  observeEvent(input$clear_all, {
    onglet <- input$onglets
    if (identical(onglet, "Actifs")) {
      .clear_all(proxy_actifs)
    } else if (identical(onglet, "Terminés")) {
      .clear_all(proxy_termines)
    }
  }, ignoreInit = TRUE)
  
  
  # filter date
  observeEvent(input$reset_filtre_date, {
    df <- donnees()
    # frühestes Datum im gesamten Datensatz finden
    if (!is.null(df) && nrow(df) > 0 && "Date" %in% names(df)) {
      min_date <- suppressWarnings(min(df$Date, na.rm = TRUE))
      if (is.infinite(suppressWarnings(as.numeric(min_date))) || is.na(min_date)) {
        min_date <- Sys.Date() - 30
      }
    } else {
      min_date <- Sys.Date() - 30
    }
    # Date-Filter im UI zurücksetzen: vom ersten Event bis heute
    updateDateRangeInput(session, "filtre_date", start = as.Date(min_date), end = Sys.Date())
  })
  
  
  
  # excel -------------------------------------------------------------------
  
  dir.create("excel_outputs", showWarnings = FALSE)
  
  # --- Excel : export ---
  observeEvent(input$exporter_excel, {
    dir.create("excel_outputs", showWarnings = FALSE)
    
    # Vue actuelle (actifs filtrés + triés) pour aligner sélection & export
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
      showModal(modalDialog("Erreur lors de l'exportation :", e$message))
    })
  })
  
  
  
  # word rapport ------------------------------------------------------------
  
  # --- Word : export  ---
  observeEvent(input$exporter_word, {
    dir.create("rapport_word_outputs", showWarnings = FALSE)
    
    # Vue actuelle (actifs filtrés + triés)
    data_view <- donnees_actifs_filtrees()
    sel <- input$tableau_evenements_actifs_rows_selected
    export_data <- if (!is.null(sel) && length(sel) > 0) data_view[sel, ] else data_view
    
    if (nrow(export_data) == 0) {
      showModal(modalDialog("Aucun événement sélectionné ou actif."))
      return()
    }
    
    # Nom de fichier
    rapport_filename <- generer_nom_fichier()
    
    # Génération du document
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
      showModal(modalDialog("Erreur lors de la génération du Word :", e$message))
    })
  })
  
  # nombre fichier
  generer_nom_fichier <- function() {
    date_today <- format(Sys.Date(), "%Y-%m-%d")
    existing_files <- list.files("rapport_word_outputs", pattern = paste0("^rapport_", date_today, "_\\d+\\.docx$"))
    existing_nums <- as.integer(gsub(paste0("rapport_", date_today, "_|\\.docx"), "", existing_files))
    next_num <- ifelse(length(existing_nums) == 0, 1, max(existing_nums, na.rm = TRUE) + 1)
    sprintf("rapport_word_outputs/rapport_%s_%d.docx", date_today, next_num)
  }
  
  
  # =========== page 1 ================
  
  ajouter_en_tete <- function(doc) {
    # --- BM_SEMAINE  ---
    date_ref <- Sys.Date() - lubridate::weeks(1)
    lundi    <- lubridate::floor_date(date_ref, unit = "week", week_start = 1)
    dimanche <- lundi + lubridate::days(6)
    
    num_semaine <- lubridate::isoweek(lundi)
    an_iso      <- lubridate::isoyear(lundi)
    
    # "JJ-JJ/MM/AAAA" (ex.: "04-10/08/2025")
    periode_str <- paste0(
      format(lundi, "%d"), "-", 
      format(dimanche, "%d"), "/", 
      format(dimanche, "%m/%Y")
    )
    
    valeur_semaine <- paste0(
      " |N° ", num_semaine, "/", an_iso,
      " | SEMAINE N°", num_semaine, " (", periode_str, ") |"
    )
    
    doc <- officer::body_replace_text_at_bkm(
      doc, bookmark = "BM_SEMAINE", value = valeur_semaine
    )
    
    # ---  BM_DATE ---
    mois_fr <- c("janvier","février","mars","avril","mai","juin",
                 "juillet","août","septembre","octobre","novembre","décembre")
    aujourd_hui <- Sys.Date()
    mois_nom    <- mois_fr[as.integer(format(aujourd_hui, "%m"))]
    mois_nom    <- tools::toTitleCase(mois_nom)  # "août" -> "Août"
    
    valeur_date <- paste0(
      format(aujourd_hui, "%d"), " ", mois_nom, " ", format(aujourd_hui, "%Y")
    )
    
    doc <- officer::body_replace_text_at_bkm(
      doc, bookmark = "BM_DATE", value = valeur_date
    )
    
    doc
  }
  
  
  # ----- BM_SOURCES -----
  ajouter_sources <- function(doc, data) {
    
    # --- Helpers ---
    # Parse la colonne "Sources" où l'utilisateur saisit : "Nom (lien); AutreNom (lien2)"
    parse_sources <- function(df) {
      # retirer d'abord les NA / vides
      raw <- df$Sources
      raw <- raw[!is.na(raw) & trimws(raw) != ""]
      if (length(raw) == 0) return(character(0))
      
      # séparer par ; , ou retour ligne
      brut <- unlist(strsplit(paste(raw, collapse = "\n"), ";|,|\\n"))
      brut <- trimws(brut)
      brut <- brut[brut != ""]
      if (length(brut) == 0) return(character(0))
      
      # garder uniquement le nom avant la parenthèse finale : "ECDC (www...)" -> "ECDC"
      noms <- sub("\\s*\\(.*?\\)\\s*$", "", brut)
      noms <- trimws(noms)
      unique(noms[noms != ""])
    }
    
    join_sources <- function(vec) {
      if (length(vec) == 0) "" else paste(vec, collapse = "; ")
    }
    
    
    # Groupes par catégorie
    df_nat  <- dplyr::filter(data, Categorie == "Alertes nationales")
    df_intl <- dplyr::filter(data, Categorie == "Alertes internationales")
    df_both <- dplyr::filter(data, Categorie == "Alertes internationales et nationales")
    
    src_nat  <- parse_sources(df_nat)
    src_intl <- parse_sources(df_intl)
    src_both <- parse_sources(df_both)
    
    # Styles (tout en blanc, italique, Century, taille 9)
    base_style <- officer::fp_text(
      color = "white", italic = TRUE, font.size = 9, font.family = "Century"
    )
    bold_style <- officer::fp_text(
      color = "white", italic = TRUE, font.size = 9, font.family = "Century", bold = TRUE
    )
    head_style <- officer::fp_text(
      color = "white", italic = TRUE, font.size = 9, font.family = "Century", underlined = TRUE
    )
    
    # Lignes formatées
    head_par <- officer::fpar(officer::ftext("Sources de données de cette semaine : ", head_style))
    
    nat_par  <- officer::fpar(
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
    
    # Aller au signet et remplacer par du contenu riche (compatibilité officer)
    doc <- officer::cursor_bookmark(doc, "BM_SOURCES")
    doc <- officer::body_add_fpar(doc, value = head_par,  pos = "on")
    doc <- officer::body_add_fpar(doc, value = nat_par,   pos = "after")
    doc <- officer::body_add_fpar(doc, value = intl_par,  pos = "after")
    doc <- officer::body_add_fpar(doc, value = both_par,  pos = "after")
    
    doc
  }
  
  
  
  
 # ============== page 2 =================
  
  # --------Sommaire des evenements ----------
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
        # Titre de catégorie (uniquement ici dans le sommaire)
        doc <- doc %>% officer::body_add_par(cat_map[[cat]], style = "m_heading_2")
        
        for (i in 1:nrow(cat_data)) {
          # Titres d’événements en m_heading_4
          doc <- doc %>% officer::body_add_par(
            paste0(compteur, ". ", cat_data$Titre[i]), style = "m_heading_4"
          )
          compteur <- compteur + 1
        }
        
        # <- keine zusätzliche Leerzeile mehr (verhindert den „Balken“)
      }
    }
    
    doc
  }
  
  
  
  # word: ajouter evenements ------------------------------------------------------
  ajouter_evenements <- function(doc, data, styles) {
    # --- Catégories et titres de section (inchangés) ---
    cat_map <- list(
      "Alertes internationales et nationales" = "Alertes et Urgences de Santé Publique au niveau international et national",
      "Alertes nationales"                    = "Alertes et Urgences de Santé Publique au niveau national",
      "Alertes internationales"               = "Alertes et Urgences de Santé Publique au niveau international"
    )
    categories <- names(cat_map)
    compteur   <- 1
    
    # --- Mise en forme simple ---
    p_std   <- officer::fp_par(line_spacing = 0.7)
    run_std <- officer::fp_text()
    run_b   <- officer::fp_text(bold = TRUE)
    
    # --- Champs à inclure pour chaque événement ---
    champs  <- c("Situation", "Evaluation_Risque", "Mesures")
    labels  <- c("Situation épidémiologique", "Évaluation des risques", "Mesures entreprises")
    
    for (cat in categories) {
      cat_data <- data %>% dplyr::filter(Categorie == cat)
      if (nrow(cat_data) == 0) next
      
      for (i in seq_len(nrow(cat_data))) {
        row <- cat_data[i, ]
        
        # 1) Titre de l'événement (niveau 1)
        doc <- doc %>%
          officer::body_add_par(paste0(compteur, ". ", row$Titre), style = "m_heading_1")
        
        # 2) Détails limités aux 3 champs demandés (dans l'ordre)
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
        
        # 3) Résumé : toujours placé à la fin de l'événement (si renseigné)
        if (!is.null(row$Resume) && !is.na(row$Resume) && nzchar(trimws(as.character(row$Resume)))) {
          doc <- officer::body_add_par(doc, "Résumé", style = "m_resume")
          doc <- officer::body_add_par(doc, as.character(row$Resume), style = "m_resume")
        }
        
        compteur <- compteur + 1
      }
    }
    
    doc
  }
  
  
  
}


