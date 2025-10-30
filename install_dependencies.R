# install_dependencies.R

# Load the renv package to restore the project environment
if (!requireNamespace("renv", quietly = TRUE)) {
  install.packages("renv")
}

# Restore the environment from the renv.lock file
renv::restore()

# Install other necessary packages 
required_packages <- c(
  "shiny", "shinyjs", "DT", "officer", "flextable", 
  "readr", "dplyr", "writexl", "lubridate", "stringr"
)

