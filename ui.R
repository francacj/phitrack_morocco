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
    max-height: 78px;     /* gewünschte Zeilenhöhe */
    overflow-y: auto;     /* Scrollbar bei langem Text */
    display: block;
    white-space: normal;  /* Zeilenumbruch erlaubt */
  }
    .info-icon{
    display:inline-block; margin-left:6px; width:18px; height:18px;
    line-height:18px; text-align:center; border-radius:50%;
    border:1px solid #17a2b8; color:#17a2b8; font-weight:700;
    font-family: Arial, sans-serif; cursor:help; font-size:12px;
  }
  .info-icon:hover{ background:#e8f7fb; }
    "))
  ),
  titlePanel("Outil de rapport d’événements"),
  tags$script(HTML('Shiny.addCustomMessageHandler("show_export_count", function(message) {
    alert("Nombre d\'événements sélectionnés : " + message);
  });')),
  sidebarLayout(
    sidebarPanel(
      div(class = "large-input", textInput("titre", "Titre de l'événement", width = "100%")),
      selectInput("categorie", "Catégorie d’alerte", 
                  choices = c("", "Alertes internationales et nationales", "Alertes nationales", "Alertes internationales"),
                  selected = ""),
      selectInput("type_evenement", "Type d'événement", 
                  choices = c("", "Humain", "Animal", "Environnement"),
                  selected = ""),
      div(class = "large-input", textInput("sources", label = HTML(
        'Sources <span class="info-icon" title="Saisir les sources au format : Nom (lien). Exemple : WHO (https://www.who.int); ECDC (https://www.ecdc.europa.eu)">i</span>'
      ), width = "100%")),
      
      dateInput("date", "Date du dernier état des données",
                value = Sys.Date(), format = "yyyy-mm-dd", language = "fr"),
      textAreaInput("situation", "Situation épidémiologique", rows = 10),
      textAreaInput("evaluation_risque", "Évaluation des risques", rows = 10),
      textAreaInput("mesures", "Mesures entreprises", rows = 10),
      selectInput("pertinence_sante", "Pertinence pour la santé publique", 
                  choices = c("", "Très faible", "Faible", "Modérée", "Élevée"),
                  selected = ""),
      selectInput("pertinence_voyageurs", "Pertinence pour les voyageurs", 
                  choices = c("", "Très faible", "Faible", "Modérée", "Élevée"),
                  selected = ""),
      div(class = "large-input", textInput("commentaires", "Commentaires", width = "100%")),
      tags$div(style = "background-color: #ffffcc; padding: 10px; border-radius: 5px;",
               textAreaInput("resume", "Résumé de la maladie", rows = 10)),
      div(class = "large-input", textInput("autre", "Autre", width = "100%")),
      selectInput("statut", "Statut de l'événement", choices = c("actif", "terminé")),
      
      div(style = "margin-top: 20px;",
          h4("Ajouter un nouvel événement"),
          actionButton("effacer_masque", label = "Effacer le masque", icon = icon("eraser"), class = "btn btn-default"),
          actionButton("enregistrer", label = "Ajouter cet événement", icon = icon("plus"), class = "btn btn-success")
      ),
      div(style = "margin-top: 20px;",
          h4("Supprimer un événement"),
          actionButton("supprimer", label = "Supprimer l'événement", icon = icon("trash"), class = "btn btn-danger")
      ),
      div(style = "margin-top: 20px;",
          h4("Éditer les événements"),
          actionButton("editer", label = "Charger pour édition", icon = icon("edit"), class = "btn btn-warning"),
          br(), br(),
          actionButton("mettre_a_jour", label = "Enregistrer les modifications de l'événement", icon = icon("refresh"), class = "btn btn-primary"),
          br(), br(),
          actionButton("enregistrer_nouveau", label = "Enregistrer comme nouvel événement (update de l’événement précédent)", icon = icon("copy"), class = "btn btn-info")
      ),
      div(style = "margin-top: 20px;",
          h4("Exporter les événements"),
          actionButton("exporter_excel", label = "Exporter en Excel", icon = icon("file-excel"), class = "btn btn-outline-success", 
                       onclick = 'Shiny.setInputValue("count_selected_excel", Math.random())'),
          actionButton("exporter_word", label = "Créer un rapport Word", icon = icon("file-word"), class = "btn btn-outline-primary",
                       onclick = 'Shiny.setInputValue("count_selected_word", Math.random())')
      )
    ),
    mainPanel(
      dateRangeInput("filtre_date", "Filtrer par date", start = Sys.Date() - 30, end = Sys.Date()),
      actionButton("reset_filtre_date", "Réinitialiser le filtre de date", icon = icon("undo")),
      
      # Boutons globaux agissant sur l’onglet actif
      div(class = "table-actions",
          actionButton("select_all", "Sélectionner tous les événements", icon = icon("check-double"),
                       class = "btn btn-outline-secondary btn-sm"),
          actionButton("clear_all",  "Désélectionner tous les événements", icon = icon("ban"),
                       class = "btn btn-outline-secondary btn-sm")
      ),
      
      # Un seul tabsetPanel avec une ID pour savoir quel onglet est actif
      tabsetPanel(id = "onglets",
                  tabPanel("Actifs",
                           DTOutput("tableau_evenements_actifs")
                  ),
                  tabPanel("Terminés",
                           DTOutput("tableau_evenements_termines")
                  )
      )
    )
  )
)
