#
#
# Het tabellenboekscript verzamelt enorm veel data en voert een flink aantal testen uit.
# Deze data kan, naast gebruik voor een tabellenboek, ook handig zijn voor intern gebruik
# bij analyses. De resultaten zijn in theorie te lezen vanuit de map resultaten_csv, maar
# dit is niet bijzonder mensvriendelijk. Daarom kan dit script gebruikt worden om de
# resultaten om te zetten naar een fijner formaat.
#
# Let op: dit script kan alleen worden uitgevoerd NADAT de tabellenboeken zijn gemaakt.
# Voer dus eerst tbl_maken.R uit, en daarna pas dit script.
#
#

suppressPackageStartupMessages(library(tidyverse))
library(openxlsx)
library(this.path)

setwd(dirname(this.path()))
source("tbl_helpers.R")

{
  config.file = choose.files(caption="Selecteer configuratiebestand...",
                             filters=c("Excel-bestand (*.xlsx)","*.xlsx"),
                             multi=F)
  #config.file = "configuratieVO.xlsx"
  if (!str_ends(config.file, ".xlsx")) msg("Configuratiebestand dient een Excel-bestand te zijn.", ERR)
  setwd(dirname(config.file))
  
  if (!file.exists(sprintf("resultaten_csv/results_%s.csv", basename(config.file))) ||
      !file.exists(sprintf("resultaten_csv/varlabels_%s.csv", basename(config.file)))) {
    stop("Geen resultaten gevonden voor dit configuratiebestand. Voer eerst tbl_maken.R uit.")
  }
  
  sheets = c("datasets", "onderdelen", "labelcorrectie")
  for (sheet in sheets) {
    tmp = read.xlsx(config.file, sheet=sheet)
    if (ncol(tmp) == 1) tmp = tmp[[1]]
    assign(sheet, tmp, envir=.GlobalEnv)
  }
  rm(tmp)
  
  results = read.csv(sprintf("resultaten_csv/results_%s.csv", basename(config.file)), fileEncoding="UTF-8")
  var_labels = read.csv(sprintf("resultaten_csv/varlabels_%s.csv", basename(config.file)), fileEncoding="UTF-8")
  kolom_opbouw = read.csv(sprintf("resultaten_csv/settings_%s.csv", basename(config.file)), fileEncoding="UTF-8")
  
  # labels toevoegen
  results = results %>%
    mutate(dataset=datasets$naam_dataset[dataset]) %>%
    mutate(crossing.val=as.character(crossing.val), val=as.character(val), subset.val=as.character(subset.val)) %>%
    left_join(var_labels %>% filter(val == "var") %>% select(-val) %>% rename(crossing.lab=label), by=c("crossing"="var")) %>% # crossing
    relocate(crossing.lab, .after=crossing) %>%
    left_join(var_labels %>% filter(val != "var") %>% rename(crossing.val.lab=label), by=c("crossing"="var", "crossing.val"="val")) %>% # crossing - waarde
    unite(crossing.val, crossing.val, crossing.val.lab, sep=" - ") %>%
    left_join(var_labels %>% filter(val == "var") %>% select(-val) %>% rename(var.lab=label), by="var") %>% # variabele
    relocate(var.lab, .after=var) %>%
    left_join(var_labels %>% filter(val != "var") %>% rename(val.lab=label), by=c("var", "val")) %>% # variabele - waarde
    unite(val, val, val.lab, sep=" - ") %>%
    left_join(var_labels %>% filter(val != "var") %>% rename(subset.val.lab=label), by=c("subset"="var", "subset.val"="val")) %>% # subset - waarde
    unite(subset.val, subset.val, subset.val.lab, sep=" - ")
  
  results$sign.vs[which(is.numeric(results$sign.vs) & results$sign.vs > 0)] = sprintf("Totaal %s (subset %s, jaar %d)",
                                                                               datasets$naam_dataset[kolom_opbouw$dataset[results$sign.vs[is.numeric(results$sign.vs) & results$sign.vs > 0]]],
                                                                               kolom_opbouw$subset[results$sign.vs[is.numeric(results$sign.vs) & results$sign.vs > 0]],
                                                                               kolom_opbouw$year[results$sign.vs[is.numeric(results$sign.vs) & results$sign.vs > 0]])
  results$sign.vs[which(results$sign.vs == "0")] = "binnen crossing"
  
  # we gaan door de resultaten loopen per subset, en dan per subsetwaarde
  # dat klinkt abstract, maar in de praktijk betekent dat zoiets als per databestand, per gemeente
  results = results %>%
    group_by(subset, subset.val)
  subsets = group_keys(results)
  for (i in 1:nrow(subsets)) {
    if (!is.na(subsets$subset[i])) {
      filename = sprintf("onderzoek_%s_%s.xlsx", subsets$subset[i], subsets$subset.val[i])
    } else {
      filename = "onderzoek_totalen.xlsx"
    }
    
    output = results[group_rows(results)[[i]],] %>% ungroup() %>% select(-subset, -subset.val)
    
    write.xlsx(output, paste0("output/", filename))
  }
  
  msg("Er zijn %d Excelbestanden geplaatst in de map output, allemaal met de prefix onderzoek_.", nrow(subsets), level=MSG)
}

