pacman::p_load(
  shiny,
  DT,
  shinyjs
)

ui <- fluidPage(
  useShinyjs(),
  tags$head(
    tags$style(HTML("
      .large-input input {
        height: 60px !important;
      }
      .table-actions { display: flex; gap: 8px; justify-content: flex-end; margin-bottom: 8px; }
      #annuler_modif {
        background-color: #a8bacc;
        color: white;
        border-color: #a8bacc;
      }
      table.dataTable tbody td { vertical-align: top; }
      .cell-fixed {
        max-height: 78px;
        overflow-y: auto;
        display: block;
        white-space: normal;
      }
      .info-icon{
        display:inline-block; margin-left:6px; width:18px; height:18px;
        line-height:18px; text-align:center; border-radius:50%;
        border:1px solid #17a2b8; color:#17a2b8; font-weight:700;
        font-family: Arial, sans-serif; cursor:help; font-size:12px;
      }
      .info-icon:hover{ background:#e8f7fb; }

      .zoom-dt #sidebar_panel { display: none !important; }
      .zoom-dt #main_panel { width: 100% !important; }
      .zoom-dt .dataTables_scrollBody {
        max-height: calc(100vh - 260px) !important;
      }
      .zoom-dt .col-sm-4 { display: none !important; }
      .zoom-dt .col-sm-8 { width: 100% !important; }

      #effacer_masque {
        background-color: #6c757d !important;
        border-color: #6c757d !important;
        color: #ffffff !important;
      }

      .export-box {
        border: 2px solid #ff8c00;
        padding: 10px;
        border-radius: 0px;
      }
    "))
  ),
  titlePanel("Outil de reporting des événements"),
  tags$script(HTML('Shiny.addCustomMessageHandler("show_export_count", function(message) {
    alert("Nombre d’événements sélectionnés : " + message);
  });')),
  sidebarLayout(
    tags$div(
      id = "sidebar_panel",
      sidebarPanel(
        div(style = "display: flex; gap: 10px; margin-top: 20px;",
            actionButton("editer", label = "Charger pour édition", icon = icon("edit"), class = "btn btn-warning"),
            actionButton("supprimer", label = "Supprimer", icon = icon("trash"), class = "btn btn-danger")
        ),
        
        div(class = "large-input", textInput("titre", "Titre de l’événement", width = "100%")),
        div(class = "large-input", textInput("regions", "Régions (pays / régions mentionnés)", width = "100%")),
        
        selectInput("categorie", "Catégorie d’alerte",
                    choices = c(
                      "",
                      "Alertes internationales et nationales",
                      "Alertes nationales",
                      "Alertes internationales"
                    ),
                    selected = ""),
        
        selectInput("type_evenement", "Type d’événement",
                    choices = c(
                      "",
                      "Humain",
                      "Animal",
                      "Environnement",
                      "Humain et animal",
                      "Autre"
                    ),
                    selected = ""),
        
        div(
          class = "large-input",
          textAreaInput(
            inputId = "sources",
            label = HTML(
              'Sources <span class="info-icon" title="Saisir une source par ligne sous la forme : Nom (lien). Exemple : OMS (https://www.who.int)
ECDC (https://www.ecdc.europa.eu)">i</span>'
            ),
            width = "100%",
            rows = 4,
            placeholder = "Une source par ligne :\nOMS (https://www.who.int)\nECDC (https://www.ecdc.europa.eu)"
          )
        ),
        
        dateInput("date", "Date (statut des données le plus récent)",
                  value = Sys.Date(), format = "yyyy-mm-dd", language = "fr"),
        
        textAreaInput("situation", "Situation épidémiologique", rows = 10),
        textAreaInput("evaluation_risque", "Évaluation des risques", rows = 10),
        textAreaInput("mesures", "Mesures entreprises", rows = 10),
        
        selectInput("pertinence_sante", "Pertinence santé publique",
                    choices = c(
                      "",
                      "Très faible",
                      "Faible",
                      "Modérée",
                      "Élevée"
                    ),
                    selected = ""),
        
        selectInput("pertinence_voyageurs", "Pertinence voyageurs",
                    choices = c(
                      "",
                      "Très faible",
                      "Faible",
                      "Modérée",
                      "Élevée"
                    ),
                    selected = ""),
        
        div(class = "large-input", textInput("commentaires", "Commentaires", width = "100%")),
        
        tags$div(style = "background-color: #ffffcc; padding: 10px; border-radius: 5px;",
                 textAreaInput("resume", "Résumé de la maladie", rows = 10)),
        
        div(class = "large-input", textInput("autre", "Autre", width = "100%")),
        
        selectInput("statut", "Statut de l’événement",
                    choices = c("actif", "terminé"),
                    selected = "actif"),
        
        tags$hr(),
        h4("EIOS"),
        fileInput(
          inputId = "upload_eios",
          label   = "Téléverser EIOS (CSV)",
          accept  = c(".csv"),
          buttonLabel = "Choisir un fichier CSV",
          placeholder = "Aucun fichier sélectionné"
        ),
        actionButton(
          inputId = "copier_vers_actifs",
          label   = "Copier la sélection vers « Actifs »",
          icon    = icon("arrow-right"),
          class   = "btn btn-outline-warning"
        ),
        
        h4("Téléverser des images"),
        tags$div(style = "margin-bottom: 10px;",
                 tags$label(HTML(
                   'Téléverser des graphiques (liés à l’événement sélectionné)
     <span class="info-icon" title="⚠️ Astuce : veuillez sélectionner ou charger un événement dans le tableau avant de téléverser des images. Les fichiers seront liés à cet événement et apparaîtront dans le rapport Word.">i</span>'
                 ), style = "font-weight: bold;"),
                 fileInput(
                   inputId = "televerser_graphs",
                   label   = NULL,
                   multiple = TRUE,
                   accept   = c("image/png","image/jpeg","image/jpg","image/svg+xml"),
                   buttonLabel = "Choisir des fichiers"
                 ),
                 tags$style(HTML("
    #televerser_graphs button, #televerser_graphs .btn {
      background-color: #ff8c00 !important;
      border-color:     #ff8c00 !important;
      color: #ffffff !important;
    }
  "))
        ),
        
        div(style = "margin-top: 20px;",
            h4("Ajouter un nouvel événement"),
            actionButton("effacer_masque", label = "Vider le formulaire", icon = icon("eraser"), class = "btn btn-default"),
            actionButton("enregistrer", label = "Ajouter cet événement", icon = icon("plus"), class = "btn btn-success")
        ),
        
        div(style = "margin-top: 20px;",
            actionButton("mettre_a_jour", label = "Enregistrer les modifications", icon = icon("refresh"), class = "btn btn-primary"),
            br(), br(),
            actionButton("enregistrer_nouveau", label = "Enregistrer comme nouvel événement (mise à jour)", icon = icon("copy"), class = "btn btn-info")
        ),
        
        tags$div(
          class = "export-box",
          style = "margin-top: 20px;",
          h4("Exporter des événements"),
          actionButton("exporter_excel", label = "Exporter vers Excel", icon = icon("file-excel"), class = "btn btn-outline-success",
                       onclick = 'Shiny.setInputValue("count_selected_excel", Math.random())'),
          actionButton("exporter_word", label = "Créer un rapport Word", icon = icon("file-word"), class = "btn btn-outline-primary",
                       onclick = 'Shiny.setInputValue("count_selected_word", Math.random())')
        )
      )
    ),
    
    tags$div(
      id = "main_panel",
      mainPanel(
        dateRangeInput("filtre_date", "Filtrer par date", start = Sys.Date() - 30, end = Sys.Date()),
        actionButton("reset_filtre_date", "Réinitialiser le filtre de date", icon = icon("undo")),
        
        fluidRow(
          column(
            width = 4,
            selectInput(
              inputId = "champ_filtre",
              label   = "Champ de recherche",
              choices = c(
                "Tous les champs"  = "tous",
                "Titre"            = "Titre",
                "Sources"          = "Sources",
                "Situation"        = "Situation",
                "Mesures"          = "Mesures",
                "Résumé"           = "Resume",
                "Évaluation risques" = "Evaluation_Risque",
                "Commentaires"     = "Commentaires",
                "Autre"            = "Autre",
                "Régions"          = "Régions",
                "EIOS_id"          = "EIOS_id"
              ),
              selected = "tous"
            )
          ),
          column(
            width = 4,
            textInput(
              inputId = "texte_filtre",
              label   = "Mot-clé",
              value   = ""
            )
          )
        ),
        
        div(class = "table-actions",
            actionButton("select_all", "Sélectionner tous", icon = icon("check-double"),
                         class = "btn btn-outline-secondary btn-sm"),
            actionButton("clear_all",  "Effacer la sélection", icon = icon("ban"),
                         class = "btn btn-outline-secondary btn-sm"),
            actionButton("zoom_table", "Agrandir le tableau", icon = icon("expand"),
                         class = "btn btn-outline-secondary btn-sm")
        ),
        
        tabsetPanel(id = "onglets",
                    tabPanel("Actifs",
                             DTOutput("tableau_evenements_actifs")
                    ),
                    tabPanel("Terminés",
                             DTOutput("tableau_evenements_termines")
                    ),
                    tabPanel("EIOS",
                             DTOutput("tableau_eios")
                    )
        )
      )
    )
  )
)