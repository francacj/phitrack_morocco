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
  # (anglais -> français) 
  # =========================================================
  normaliser_statut <- function(x) {
    x <- as.character(x)
    dplyr::case_when(
      is.na(x) ~ x,
      x %in% c("active", "Active") ~ "actif",
      x %in% c("closed", "Closed") ~ "terminé",
      TRUE ~ x
    )
  }
  
  normaliser_categorie <- function(x) {
    x <- as.character(x)
    dplyr::case_when(
      is.na(x) ~ x,
      x %in% c("International and national alerts", "International and national") ~ "Alertes internationales et nationales",
      x %in% c("National alerts", "National") ~ "Alertes nationales",
      x %in% c("International alerts", "International") ~ "Alertes internationales",
      TRUE ~ x
    )
  }
  
  normaliser_donnees <- function(df) {
    if (is.null(df) || nrow(df) == 0) return(df)
    if ("Statut" %in% names(df)) df$Statut <- normaliser_statut(df$Statut)
    if ("Categorie" %in% names(df)) df$Categorie <- normaliser_categorie(df$Categorie)
    df
  }
  
  # =========================================================
  # SOURCE DE DONNÉES (CSV) - SYNCHRONISATION MULTI-User
  # =========================================================
  donnees <- reactivePoll(
    intervalMillis = 2000L,
    session = session,
    checkFunc = function() {
      fi <- file.info(fichier_donnees)
      if (is.na(fi$mtime)) return(0)
      paste(fi$mtime, fi$size, sep = "_")
    },
    valueFunc = function() {
      tryCatch({
        df_lu <- readr::read_csv(
          fichier_donnees,
          show_col_types = FALSE,
          col_types = cols(
            Version = col_double(),
            Date = col_date(),
            Derniere_Modification = col_datetime()
          )
        )
        df_lu <- as.data.frame(df_lu, stringsAsFactors = FALSE)
        if (!"Images" %in% names(df_lu)) df_lu$Images <- NA_character_
        df_lu <- normaliser_donnees(df_lu)
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
  
  # =========================================================
  # BOUTON ZOOM TABLEAU
  # =========================================================
  observeEvent(input$zoom_table, {
    shinyjs::toggleClass(selector = "body", class = "zoom-dt")
    shinyjs::runjs("setTimeout(function(){ $(window).trigger('resize'); }, 150);")
  }, ignoreInit = TRUE)
  
  # =========================================================
  # EIOS EN MÉMOIRE
  # =========================================================
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
  
  # =========================================================
  # EIOS : IMPORT (CSV) -> remplit l'onglet EIOS
  # =========================================================
  observeEvent(input$upload_eios, {
    if (is.null(input$upload_eios$datapath)) return()
    chemin <- input$upload_eios$datapath
    
    # --- Détection du séparateur ---
    deviner_separateur <- function(fichier) {
      lignes <- tryCatch(readLines(fichier, n = 10, warn = FALSE), error = function(e) character(0))
      if (length(lignes) == 0) return(",")
      d <- tryCatch(readr::guess_delim(lignes), error = function(e) NULL)
      if (is.null(d) || is.na(d) || !nzchar(d)) return(",")
      d
    }
    separateur <- deviner_separateur(chemin)
    
    # --- Lecture ---
    df_brut <- tryCatch(
      readr::read_delim(chemin, delim = separateur, show_col_types = FALSE),
      error = function(e) {
        showModal(modalDialog("Erreur de lecture du fichier EIOS : ", e$message))
        return(NULL)
      }
    )
    if (is.null(df_brut)) return()
    if (nrow(df_brut) == 0) {
      showModal(modalDialog("Le fichier EIOS ne contient aucune ligne."))
      return()
    }
    
    # --- Utilitaire : choisir la 1ère colonne existante ---
    choisir_col <- function(df, choix) {
      ok <- choix[choix %in% names(df)]
      if (length(ok) == 0) NA_character_ else ok[1]
    }
    
    # Colonnes (priorités adaptées à ton export actuel)
    col_sources <- choisir_col(df_brut, c("externalLink"))
    col_titre   <- choisir_col(df_brut, c("title"))
    col_sit     <- choisir_col(df_brut, c("summary"))
    col_autre   <- choisir_col(df_brut, c("description"))
    col_regions <- choisir_col(df_brut, c("mentionedCountries"))
    col_id      <- choisir_col(df_brut, c("emmId"))  # ✅ FIX ICI
    
    manquantes <- c(
      if (is.na(col_sources)) "externalLink" else NULL,
      if (is.na(col_titre)) "title" else NULL,
      if (is.na(col_sit)) "summary" else NULL,
      if (is.na(col_autre)) "description" else NULL,
      if (is.na(col_regions)) "mentionedCountries" else NULL,
      if (is.na(col_id)) "emmId" else NULL
    )
    
    if (length(manquantes) > 0) {
      showModal(modalDialog(
        title = "Import EIOS impossible",
        paste0(
          "Colonnes introuvables : ", paste(manquantes, collapse = ", "),
          "\n\nColonnes détectées :\n- ", paste(names(df_brut), collapse = "\n- "),
          "\n\nSéparateur détecté : '", separateur, "'"
        ),
        easyClose = TRUE
      ))
      return()
    }
    
    # --- Mapping vers ton onglet EIOS (uniquement ce dont tu as besoin) ---
    df_eios <- df_brut %>%
      dplyr::transmute(
        Sources   = as.character(.data[[col_sources]]),
        Titre     = as.character(.data[[col_titre]]),
        Situation = as.character(.data[[col_sit]]),
        Autre     = as.character(.data[[col_autre]]),
        `Régions` = as.character(.data[[col_regions]]),
        EIOS_id   = as.character(.data[[col_id]])
      )
    
    eios_donnees(df_eios)
    
    showModal(modalDialog(
      title = "Import EIOS réussi",
      paste0("Lignes importées : ", nrow(df_eios)),
      easyClose = TRUE,
      footer = modalButton("Fermer")
    ))
  })
  
  
  # =========================================================
  # EIOS : COPIER LA SÉLECTION -> ajoute des événements dans "Actifs"
  # À COLLER juste après le bloc d'import ci-dessus
  # =========================================================
  observeEvent(input$copier_vers_actifs, {
    df_eios <- eios_donnees()
    sel <- input$tableau_eios_rows_selected
    
    if (is.null(sel) || length(sel) == 0) {
      showModal(modalDialog("Veuillez sélectionner au moins une ligne dans l’onglet « EIOS »."))
      return()
    }
    
    sel <- sel[is.finite(sel)]
    sel <- sel[sel >= 1 & sel <= nrow(df_eios)]
    if (length(sel) == 0) return()
    
    selection <- df_eios[sel, , drop = FALSE]
    if (nrow(selection) == 0) return()
    
    df_all <- donnees()
    if (is.null(df_all) || nrow(df_all) == 0) {
      df_all <- data.frame(
        ID_Evenement = character(),
        Update = character(),
        Version = numeric(),
        Date = as.Date(character()),
        Derniere_Modification = as.POSIXct(character()),
        Titre = character(),
        Type_Evenement = character(),
        Categorie = character(),
        `Régions` = character(),
        Sources = character(),
        Situation = character(),
        Evaluation_Risque = character(),
        Mesures = character(),
        Pertinence_Sante = character(),
        Pertinence_Voyageurs = character(),
        Commentaires = character(),
        Resume = character(),
        Autre = character(),
        Statut = character(),
        EIOS_id = character(),
        Images = character(),
        stringsAsFactors = FALSE
      )
    }
    
    # Créer n nouveaux événements (1 par ligne EIOS sélectionnée)
    date_tag <- format(Sys.Date(), "%Y%m%d")
    deja <- sum(grepl(date_tag, as.character(df_all$ID_Evenement)))
    n <- nrow(selection)
    
    to_add <- lapply(seq_len(n), function(i) {
      nouvelle_id <- sprintf("EVT_%s_%03d", date_tag, deja + i)
      
      data.frame(
        ID_Evenement          = nouvelle_id,
        Update                = "nouvel événement depuis EIOS",
        Version               = 1,
        Date                  = Sys.Date(),
        Derniere_Modification = lubridate::force_tz(Sys.time(), "Africa/Casablanca"),
        Titre                 = as.character(selection$Titre[i]),
        Type_Evenement        = NA_character_,
        Categorie             = NA_character_,
        `Régions`             = as.character(selection$`Régions`[i]),
        Sources               = as.character(selection$Sources[i]),
        Situation             = as.character(selection$Situation[i]),
        Evaluation_Risque     = NA_character_,
        Mesures               = NA_character_,
        Pertinence_Sante      = NA_character_,
        Pertinence_Voyageurs  = NA_character_,
        Commentaires          = NA_character_,
        Resume                = NA_character_,
        Autre                 = as.character(selection$Autre[i]),
        Statut                = "actif",
        EIOS_id               = as.character(selection$EIOS_id[i]),
        Images                = NA_character_,
        stringsAsFactors      = FALSE
      )
    })
    
    df_new <- dplyr::bind_rows(df_all, dplyr::bind_rows(to_add))
    df_new <- normaliser_donnees(df_new)
    
    tryCatch({
      readr::write_csv(df_new, fichier_donnees)
      
      showModal(modalDialog(
        title = "Succès",
        paste0(n, " événement(s) ont été copié(s) vers « Actifs »."),
        easyClose = TRUE,
        footer = modalButton("Fermer")
      ))
    }, error = function(e) {
      showModal(modalDialog("Erreur lors de l'enregistrement CSV : ", e$message))
    })
  })
  
  # =========================================================
  # IMAGES : DOSSIERS + UTILITAIRES
  # =========================================================
  dir.create("Graphs_uploaded", showWarnings = FALSE)
  
  id_evenement_courant <- function() {
    df_full <- donnees()
    
    if (!is.null(session$userData$indice_edition)) {
      idx <- session$userData$indice_edition
      if (is.finite(idx) && idx >= 1 && idx <= nrow(df_full)) {
        return(as.character(df_full$ID_Evenement[idx]))
      }
    }
    
    onglet <- input$onglets
    if (identical(onglet, "Actifs") || identical(onglet, "Active")) {
      sel <- input$tableau_evenements_actifs_rows_selected
      if (length(sel) == 1) return(as.character(donnees_actifs_filtrees()$ID_Evenement[sel]))
    } else if (identical(onglet, "Terminés") || identical(onglet, "Closed")) {
      sel <- input$tableau_evenements_termines_rows_selected
      if (length(sel) == 1) return(as.character(donnees_termines_filtrees()$ID_Evenement[sel]))
    }
    
    NA_character_
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
  filtrer_par_texte <- function(df) {
    terme <- input$texte_filtre
    champ <- input$champ_filtre
    
    if (is.null(terme) || !nzchar(trimws(as.character(terme))) || nrow(df) == 0) return(df)
    terme <- trimws(as.character(terme))
    
    if (identical(champ, "tous")) {
      cols_existe <- colnames(df)
      if (length(cols_existe) == 0) return(df)
      
      mat_txt <- lapply(df[, cols_existe, drop = FALSE], function(x) {
        x <- as.character(x); x[is.na(x)] <- ""; x
      })
      mat_txt <- as.data.frame(mat_txt, stringsAsFactors = FALSE)
      
      idx <- apply(mat_txt, 1, function(ligne) any(grepl(terme, ligne, ignore.case = TRUE)))
      df[idx, , drop = FALSE]
    } else {
      if (!champ %in% colnames(df)) return(df)
      x <- as.character(df[[champ]]); x[is.na(x)] <- ""
      df[grepl(terme, x, ignore.case = TRUE), , drop = FALSE]
    }
  }
  
  donnees_actifs_filtrees <- reactive({
    df <- donnees()
    if (is.null(df) || nrow(df) == 0) return(df)
    
    df <- dplyr::filter(df, Statut == "actif")
    
    if (!is.null(input$filtre_date[1]) && !is.null(input$filtre_date[2])) {
      df <- dplyr::filter(df, Date >= input$filtre_date[1], Date <= input$filtre_date[2])
    }
    
    df <- dplyr::arrange(df, dplyr::desc(Date), dplyr::desc(Derniere_Modification))
    filtrer_par_texte(df)
  })
  
  donnees_termines_filtrees <- reactive({
    df <- donnees()
    if (is.null(df) || nrow(df) == 0) return(df)
    
    df <- dplyr::filter(df, Statut == "terminé")
    
    if (!is.null(input$filtre_date[1]) && !is.null(input$filtre_date[2])) {
      df <- dplyr::filter(df, Date >= input$filtre_date[1], Date <= input$filtre_date[2])
    }
    
    df <- dplyr::arrange(df, dplyr::desc(Date), dplyr::desc(Derniere_Modification))
    filtrer_par_texte(df)
  })
  
  # =========================================================
  # ÉTAT ÉDITION
  # =========================================================
  session$userData$indice_edition <- NULL
  session$userData$ids_a_supprimer <- NULL
  session$userData$source_suppression <- NULL
  session$userData$rows_eios_a_supprimer <- NULL
  
  # =========================================================
  # AJOUT / MODIFICATION / DUPLICATION
  # =========================================================
  observeEvent(input$enregistrer, {
    df_all <- donnees()
    if (is.null(df_all)) df_all <- data.frame()
    
    nouvelle_id <- sprintf(
      "EVT_%s_%03d",
      format(Sys.Date(), "%Y%m%d"),
      sum(grepl(format(Sys.Date(), "%Y%m%d"), as.character(df_all$ID_Evenement))) + 1
    )
    
    nouvel_evenement <- data.frame(
      ID_Evenement = nouvelle_id,
      Update = "Nouvel événement",
      Version = 1,
      Date = input$date,
      Derniere_Modification = lubridate::force_tz(Sys.time(), "Africa/Casablanca"),
      Titre = input$titre,
      Type_Evenement = input$type_evenement,
      Categorie = normaliser_categorie(input$categorie),
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
      Statut = normaliser_statut(input$statut),
      EIOS_id = NA_character_,
      Images = NA_character_,
      stringsAsFactors = FALSE
    )
    
    tryCatch({
      df_new <- dplyr::bind_rows(df_all, nouvel_evenement)
      df_new <- normaliser_donnees(df_new)
      readr::write_csv(df_new, fichier_donnees)
      
      vider_formulaire()
      session$userData$indice_edition <- NULL
      
      showModal(modalDialog(
        title = "Succès",
        "L'événement a été ajouté avec succès.",
        easyClose = TRUE,
        footer = modalButton("Fermer")
      ))
    }, error = function(e) {
      showModal(modalDialog("Erreur lors de l'enregistrement : ", e$message))
    })
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
    
    df_all <- donnees()
    if (!is.finite(indice) || indice < 1 || indice > nrow(df_all)) {
      showModal(modalDialog("L'événement à éditer n'existe plus. Veuillez re-sélectionner."))
      session$userData$indice_edition <- NULL
      return()
    }
    
    evenement_original <- df_all[indice, , drop = FALSE]
    original_id <- as.character(evenement_original$ID_Evenement[1])
    
    nouvelle_id <- sprintf(
      "EVT_%s_%03d",
      format(Sys.Date(), "%Y%m%d"),
      sum(grepl(format(Sys.Date(), "%Y%m%d"), as.character(df_all$ID_Evenement))) + 1
    )
    
    nouvel_evenement <- evenement_original %>%
      dplyr::mutate(
        ID_Evenement = nouvelle_id,
        Update = paste0("Update of event: ", original_id),
        Version = as.numeric(Version) + 1,
        Derniere_Modification = lubridate::force_tz(Sys.time(), "Africa/Casablanca"),
        Titre = input$titre,
        Type_Evenement = input$type_evenement,
        Categorie = normaliser_categorie(input$categorie),
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
        Statut = normaliser_statut(input$statut)
      ) %>%
      dplyr::select(ID_Evenement, Update, dplyr::everything())
    
    tryCatch({
      df_new <- dplyr::bind_rows(df_all, nouvel_evenement)
      df_new <- normaliser_donnees(df_new)
      readr::write_csv(df_new, fichier_donnees)
      
      vider_formulaire()
      session$userData$indice_edition <- NULL
      
      showModal(modalDialog(
        title = "Succès",
        "Le nouvel événement a été enregistré avec succès.",
        easyClose = TRUE,
        footer = modalButton("Fermer")
      ))
    }, error = function(e) {
      showModal(modalDialog("Erreur lors de l'enregistrement : ", e$message))
    })
  })
  
  observeEvent(input$mettre_a_jour, {
    indice <- session$userData$indice_edition
    if (is.null(indice)) return()
    
    df_all <- donnees()
    if (!is.finite(indice) || indice < 1 || indice > nrow(df_all)) {
      showModal(modalDialog("L'événement à mettre à jour n'existe plus. Veuillez re-sélectionner."))
      session$userData$indice_edition <- NULL
      return()
    }
    
    df_all[indice, ] <- df_all[indice, ] %>%
      dplyr::mutate(
        Titre = input$titre,
        Type_Evenement = input$type_evenement,
        Categorie = normaliser_categorie(input$categorie),
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
        Statut = normaliser_statut(input$statut),
        Derniere_Modification = lubridate::force_tz(Sys.time(), "Africa/Casablanca")
      )
    
    tryCatch({
      df_all <- normaliser_donnees(df_all)
      readr::write_csv(df_all, fichier_donnees)
      
      vider_formulaire()
      session$userData$indice_edition <- NULL
      
      showModal(modalDialog(
        title = "Succès",
        "Les modifications ont été enregistrées.",
        easyClose = TRUE,
        footer = modalButton("Fermer")
      ))
    }, error = function(e) {
      showModal(modalDialog("Erreur lors de la mise à jour : ", e$message))
    })
  })
  
  # =========================================================
  # CHARGER POUR ÉDITION
  # =========================================================
  observeEvent(input$editer, {
    onglet <- input$onglets
    df_full <- donnees()
    
    if (identical(onglet, "Actifs") || identical(onglet, "Active")) {
      selection <- input$tableau_evenements_actifs_rows_selected
      df_view <- donnees_actifs_filtrees()
    } else if (identical(onglet, "Terminés") || identical(onglet, "Closed")) {
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
    
    evenement <- df_view[selection, , drop = FALSE]
    idx <- which(as.character(df_full$ID_Evenement) == as.character(evenement$ID_Evenement))
    if (length(idx) != 1) {
      showModal(modalDialog("Impossible de trouver l’événement sélectionné dans le jeu de données complet."))
      return()
    }
    
    updateTextInput(session, "titre", value = as.character(evenement$Titre))
    updateSelectInput(session, "type_evenement",
                      selected = ifelse(is.na(evenement$Type_Evenement) || evenement$Type_Evenement == "", "", as.character(evenement$Type_Evenement)))
    updateSelectInput(session, "categorie",
                      selected = ifelse(is.na(evenement$Categorie) || evenement$Categorie == "", "", as.character(evenement$Categorie)))
    updateSelectInput(session, "regions", selected = as.character(evenement$`Régions`))
    updateTextInput(session, "sources", value = as.character(evenement$Sources))
    updateSelectInput(session, "statut", selected = as.character(evenement$Statut))
    updateDateInput(session, "date", value = as.Date(evenement$Date))
    updateTextAreaInput(session, "situation", value = as.character(evenement$Situation))
    updateTextAreaInput(session, "evaluation_risque", value = as.character(evenement$Evaluation_Risque))
    updateTextAreaInput(session, "mesures", value = as.character(evenement$Mesures))
    updateSelectInput(session, "pertinence_sante",
                      selected = ifelse(is.na(evenement$Pertinence_Sante) || evenement$Pertinence_Sante == "", "", as.character(evenement$Pertinence_Sante)))
    updateSelectInput(session, "pertinence_voyageurs",
                      selected = ifelse(is.na(evenement$Pertinence_Voyageurs) || evenement$Pertinence_Voyageurs == "", "", as.character(evenement$Pertinence_Voyageurs)))
    updateTextInput(session, "commentaires", value = as.character(evenement$Commentaires))
    updateTextAreaInput(session, "resume", value = as.character(evenement$Resume))
    updateTextInput(session, "autre", value = as.character(evenement$Autre))
    
    session$userData$indice_edition <- idx
  })
  
  # =========================================================
  # SUPPRESSION : PAR ID_EVENEMENT
  # =========================================================
  observeEvent(input$supprimer, {
    onglet <- input$onglets
    
    if (identical(onglet, "Actifs") || identical(onglet, "Active")) {
      sel <- shiny::isolate(input$tableau_evenements_actifs_rows_selected)
      df_view <- shiny::isolate(donnees_actifs_filtrees())
      if (is.null(sel) || length(sel) == 0) {
        showModal(modalDialog("Aucun événement sélectionné dans l’onglet « Actifs »."))
        return()
      }
      sel <- sel[is.finite(sel)]
      sel <- sel[sel >= 1 & sel <= nrow(df_view)]
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
      
    } else if (identical(onglet, "Terminés") || identical(onglet, "Closed")) {
      sel <- shiny::isolate(input$tableau_evenements_termines_rows_selected)
      df_view <- shiny::isolate(donnees_termines_filtrees())
      if (is.null(sel) || length(sel) == 0) {
        showModal(modalDialog("Aucun événement sélectionné dans l’onglet « Terminés »."))
        return()
      }
      sel <- sel[is.finite(sel)]
      sel <- sel[sel >= 1 & sel <= nrow(df_view)]
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
      if (is.null(df) || nrow(df) == 0) {
        showModal(modalDialog("Le fichier ne contient aucune donnée à supprimer."))
        return()
      }
      
      df_new <- df[!as.character(df$ID_Evenement) %in% ids, , drop = FALSE]
      
      tryCatch({
        df_new <- normaliser_donnees(df_new)
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
      if (is.null(df_e) || nrow(df_e) == 0) return()
      
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
  # PROXYS + BOUTONS SÉLECTION
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
    if (identical(onglet, "Actifs") || identical(onglet, "Active")) {
      .selectionner_tout(proxy_actifs, nrow(donnees_actifs_filtrees()))
    } else if (identical(onglet, "Terminés") || identical(onglet, "Closed")) {
      .selectionner_tout(proxy_termines, nrow(donnees_termines_filtrees()))
    } else if (identical(onglet, "EIOS")) {
      .selectionner_tout(proxy_eios, nrow(eios_donnees()))
    }
  }, ignoreInit = TRUE)
  
  observeEvent(input$clear_all, {
    onglet <- input$onglets
    if (identical(onglet, "Actifs") || identical(onglet, "Active")) {
      .effacer_selection(proxy_actifs)
    } else if (identical(onglet, "Terminés") || identical(onglet, "Closed")) {
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
    export_data <- if (!is.null(sel) && length(sel) > 0) data_view[sel, , drop = FALSE] else data_view
    
    if (nrow(export_data) == 0) {
      showModal(modalDialog("Aucun événement sélectionné ou actif."))
      return()
    }
    
    export_data <- export_data %>%
      dplyr::mutate(Derniere_Modification = format(Derniere_Modification, "%Y-%m-%d %H:%M:%S"))
    
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
  # EXPORT WORD : NOM DE FICHIER
  # =========================================================
  dir.create("rapport_word_outputs", showWarnings = FALSE)
  
  generer_nom_fichier <- function() {
    date_today <- format(Sys.Date(), "%Y-%m-%d")
    existing_files <- list.files("rapport_word_outputs", pattern = paste0("^rapport_", date_today, "_\\d+\\.docx$"))
    existing_nums <- as.integer(gsub(paste0("rapport_", date_today, "_|\\.docx"), "", existing_files))
    next_num <- ifelse(length(existing_nums) == 0, 1, max(existing_nums, na.rm = TRUE) + 1)
    sprintf("rapport_word_outputs/rapport_%s_%d.docx", date_today, next_num)
  }
  
  # =========================================================
  # EXPORT WORD (ROBUSTE : AJOUTE UNE SECTION "NON CLASSÉ")
  # =========================================================
  observeEvent(input$exporter_word, {
    dir.create("rapport_word_outputs", showWarnings = FALSE)
    
    data_view <- donnees_actifs_filtrees()
    sel <- input$tableau_evenements_actifs_rows_selected
    export_data <- if (!is.null(sel) && length(sel) > 0) data_view[sel, , drop = FALSE] else data_view
    
    if (nrow(export_data) == 0) {
      showModal(modalDialog("Aucun événement sélectionné ou actif."))
      return()
    }
    
    export_data <- normaliser_donnees(export_data)
    
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
  
  # =========================================================
  # TÉLÉVERSEMENT IMAGES
  # =========================================================
  observeEvent(input$televerser_graphs, {
    fichiers <- input$televerser_graphs
    if (is.null(fichiers) || nrow(fichiers) == 0) {
      showModal(modalDialog("Veuillez choisir un ou plusieurs fichiers image."))
      return()
    }
    
    id_evt <- id_evenement_courant()
    if (is.na(id_evt) || is.null(id_evt) || !nzchar(id_evt)) {
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
    pos <- which(as.character(df$ID_Evenement) == as.character(id_evt))
    if (length(pos) == 1) {
      exist <- as.character(df$Images[pos])
      if (is.na(exist)) exist <- ""
      ajout <- paste(chemins_relatifs, collapse = ";")
      df$Images[pos] <- if (nzchar(exist)) paste(exist, ajout, sep = ";") else ajout
      
      tryCatch(
        readr::write_csv(normaliser_donnees(df), fichier_donnees),
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
  
  # =========================================================
  # WORD : SOURCES
  # =========================================================
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
    
    df_nat  <- dplyr::filter(data, Categorie == "Alertes nationales")
    df_intl <- dplyr::filter(data, Categorie == "Alertes internationales")
    df_both <- dplyr::filter(data, Categorie == "Alertes internationales et nationales")
    
    src_nat  <- parse_sources(df_nat)
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
    doc <- officer::body_add_fpar(doc, value = nat_par,  pos = "after")
    doc <- officer::body_add_fpar(doc, value = intl_par, pos = "after")
    doc <- officer::body_add_fpar(doc, value = both_par, pos = "after")
    doc
  }
  
  # =========================================================
  # WORD : SOMMAIRE (AVEC SECTION "NON CLASSÉ")
  # =========================================================
  ajouter_sommaire <- function(doc, data, styles) {
    cat_map <- list(
      "Alertes internationales et nationales" = "Alertes et Urgences de Santé Publique au niveau international et national",
      "Alertes nationales"                   = "Alertes et Urgences de Santé Publique au niveau national",
      "Alertes internationales"              = "Alertes et Urgences de Santé Publique au niveau international"
    )
    categories <- names(cat_map)
    
    doc <- doc %>% officer::body_add_par("SOMMAIRE DES ÉVÉNEMENTS", style = "m_heading_1")
    
    compteur <- 1
    
    for (cat in categories) {
      cat_data <- data %>% dplyr::filter(Categorie == cat)
      if (nrow(cat_data) > 0) {
        doc <- doc %>% officer::body_add_par(cat_map[[cat]], style = "m_heading_2")
        for (i in seq_len(nrow(cat_data))) {
          doc <- doc %>% officer::body_add_par(
            paste0(compteur, ". ", as.character(cat_data$Titre[i])),
            style = "m_heading_4"
          )
          compteur <- compteur + 1
        }
      }
    }
    
    other_data <- data %>%
      dplyr::filter(is.na(Categorie) | trimws(as.character(Categorie)) == "" | !(Categorie %in% categories))
    
    if (nrow(other_data) > 0) {
      doc <- doc %>% officer::body_add_par("Événements non classés", style = "m_heading_2")
      for (i in seq_len(nrow(other_data))) {
        doc <- doc %>% officer::body_add_par(
          paste0(compteur, ". ", as.character(other_data$Titre[i])),
          style = "m_heading_4"
        )
        compteur <- compteur + 1
      }
    }
    
    doc
  }
  
  # =========================================================
  # WORD : ÉVÉNEMENTS (AVEC SECTION "NON CLASSÉ")
  # =========================================================
  ajouter_evenements <- function(doc, data, styles) {
    cat_map <- list(
      "Alertes internationales et nationales" = "Alertes et Urgences de Santé Publique au niveau international et national",
      "Alertes nationales"                    = "Alertes et Urgences de Santé Publique au niveau national",
      "Alertes internationales"               = "Alertes et Urgences de Santé Publique au niveau international"
    )
    categories <- names(cat_map)
    compteur <- 1
    
    p_std   <- officer::fp_par(line_spacing = 0.7)
    run_std <- officer::fp_text()
    run_b   <- officer::fp_text(bold = TRUE)
    
    champs <- c("Situation", "Evaluation_Risque", "Mesures")
    labels <- c("Situation épidémiologique", "Évaluation des risques", "Mesures entreprises")
    
    .print_one <- function(doc, row, compteur) {
      doc <- doc %>% officer::body_add_par(paste0(compteur, ". ", as.character(row$Titre)), style = "m_heading_1")
      
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
      
      list(doc = doc, compteur = compteur + 1)
    }
    
    for (cat in categories) {
      cat_data <- data %>% dplyr::filter(Categorie == cat)
      if (nrow(cat_data) == 0) next
      
      doc <- doc %>% officer::body_add_par(cat_map[[cat]], style = "m_heading_2")
      
      for (i in seq_len(nrow(cat_data))) {
        res <- .print_one(doc, cat_data[i, , drop = FALSE], compteur)
        doc <- res$doc
        compteur <- res$compteur
      }
    }
    
    other_data <- data %>%
      dplyr::filter(is.na(Categorie) | trimws(as.character(Categorie)) == "" | !(Categorie %in% categories))
    
    if (nrow(other_data) > 0) {
      doc <- doc %>% officer::body_add_par("Événements non classés", style = "m_heading_2")
      
      for (i in seq_len(nrow(other_data))) {
        res <- .print_one(doc, other_data[i, , drop = FALSE], compteur)
        doc <- res$doc
        compteur <- res$compteur
      }
    }
    
    doc
  }
}