#
#
# Dit script maakt tabellenboeken op basis van gezondheidsmonitors. 
# De configuratie gaat middels een Excelsheet, zie bijgevoegde documentatie.
# Bugs en verzoeken kunnen worden ingediend op:
# -> https://github.com/aartdijkstra/GGData_tabellenboek
#
#

# alle huidige data uit de sessie verwijderen
rm(list=ls())

# benodigde packages installeren als deze afwezig zijn
pkg_nodig = c("tidyverse", "survey", "haven", "this.path",
              "labelled", "openxlsx")

for (pkg in pkg_nodig) {
  if (system.file(package = pkg) == "") {
    install.packages(pkg)
  }
}

suppressPackageStartupMessages(library(tidyverse))
suppressPackageStartupMessages(library(survey))
library(haven)
library(this.path)
library(labelled)
library(openxlsx)

# instellen werkmap voor het laden van de andere bestanden
setwd(dirname(this.path()))
source("tbl_helpers.R")

# het gehele script wordt omgeven door curly brackets, zodat een stop() ook daadwerkelijk het script stopt
{
  # selecteren configuratiebestand en bijbehorende werkmap
  # hierin dienen de configuratie(.xlsx) en de databestanden (in de map data) te staan
  # andere mappen worden automatisch aangemaakt als deze niet bestaan
  config.file = choose.files(caption="Selecteer configuratiebestand...",
                             filters=c("Excel-bestand (*.xlsx)","*.xlsx"),
                             multi=F)
  if (!str_ends(config.file, ".xlsx")) msg("Configuratiebestand dient een Excel-bestand te zijn.", ERR)
  setwd(dirname(config.file))
  
  # algemene instellingen voor de libraries
  options(survey.lonely.psu="certainty")
  options("openxlsx.paperSize" = 9) # A4
  options("openxlsx.orientation" = "landscape")
  options("openxlsx.numFmt" = "0") # standaardformaat zonder decimalen (anders 0.0 invullen)
  
  # controleren of de configuratie leesbaar is
  tmp = read.xlsx(config.file, sheet="datasets")
  if (!exists("tmp")) msg("Configuratiebestand kon niet gelezen worden. Wellicht is deze nog geopend in Excel?", ERR)
  rm(tmp)
  
  # daadwerkelijk inlezen configuratie
  sheets = c("crossings", "datasets", "indeling_rijen", "onderdelen")
  for (sheet in sheets) {
    tmp = read.xlsx(config.file, sheet=sheet)
    if (ncol(tmp) == 1) tmp = tmp[[1]]
    assign(sheet, tmp, envir=.GlobalEnv)
  }
  rm(tmp)
  
  # variabelelijst afleiden uit de indeling van het tabellenboek;
  # iedere regel met (n)var is een variabele die we nodig hebben
  varlist = indeling_rijen$inhoud[indeling_rijen$type %in% c("var", "nvar")]
  varlist = unique(varlist[!is.na(varlist)])
  
  # datasets combineren en de strata en weegfactoren apart opslaan
  data.combined = data.frame()
  for (d in 1:nrow(datasets)) {
    data = read_spss(datasets$bestandsnaam[d], user_na=T) %>% user_na_to_na()
    
    # afwijkende kolommen registreren zodat we deze later kunnen scheiden
    afwijkend = c()
    for (c in 1:ncol(data)) {
      colname = colnames(data)[c]
      
      # strata willen we apart opslaan
      if (!is.na(datasets$stratum[d]) && colname == datasets$stratum[d]) {
        colnames(data)[c] = "tbl_strata"
        next
      }
      
      if (colname %in% colnames(data.combined)) {
        existing = data.combined[[which(colnames(data.combined) == colname)]]
        # controleren of het type en, indien relevant, factors overeenkomen
        if (!(identical(var_label(data[[c]]), var_label(existing)) && all(val_labels(data[[c]]) == val_labels(existing)))) {
          afwijkend = c(afwijkend, c)
        }
      }
    }
    
    # herschrijven kolomnamen zodat ze niet gaan storen
    msg("Afwijkende kolommen: %s", str_c(colnames(data)[afwijkend], collapse=","))
    if (length(afwijkend) > 0) colnames(data)[afwijkend] = paste0("_", d, "_", colnames(data)[afwijkend])
    
    data[,"tbl_dataset"] = d
    data.combined = bind_rows(data.combined, data)
    
    # bind_rows haalt soms willekeurig labels weg, dus die moeten we handmatig terugzetten
    for (c in 1:ncol(data)) {
      colname = colnames(data)[c]
      var_label(data.combined[[colname]]) = var_label(data[[colname]])
    }
  }
  data = data.combined
  rm(data.combined) # twee keer hetzelfde object is zonde van het geheugen
  
  # dummies maken
  for (c in 1:ncol(data)) {
    colname = colnames(data)[c]
    if (!(colname %in% varlist)) next
    if (str_starts(colname, "_") || str_starts(colname, "dummy")) next
    
    if (is.null(val_labels(data[[c]]))) next
    
    labels = val_labels(data[[c]])
    for (i in labels) {
      data[,paste0("dummy.", colname, ".", i)] = data[[c]] == unname(i)
    }
  }
  
  # superstrata maken
  # dit doen we door de bestaande strata te nummeren van 1-<aantal> en daar 1000 * dataset bij te doen
  # dus stratum 40 uit dataset 1 wordt 1040, stratum 25 uit dataset 2 wordt 2025, enz.
  data$superstrata = NA
  for (d in 1:nrow(datasets)) {
    strata = sort(unique(data$tbl_strata[data$tbl_dataset == d]))
    for (i in 1:length(strata)) {
      data$superstrata[data$tbl_dataset == d & data$tbl_strata == strata[i]] = i
    }
    data$superstrata[data$tbl_dataset == d] = d*1000 + data$superstrata[data$tbl_dataset == d]
  }
}