#
#
# Dit script maakt tabellenboeken op basis van gezondheidsmonitors. 
# De configuratie gaat middels een Excelsheet, zie bijgevoegde documentatie.
# Problemen en verzoeken kunnen worden ingediend op:
# https://github.com/ggdatascience/GGData_tabellenboek
#
# Versie: concept
#

# alle huidige data uit de sessie verwijderen
if(!(exists("skip_config_popup") && skip_config_popup)){
  rm(list=ls()[!ls() %in%  c("skip_config_popup", "config.file")])
} 
label_problemen <- NULL
# benodigde packages installeren als deze afwezig zijn
pkg_nodig = c("tidyverse", "survey", "haven", "this.path", "textutils",
              "labelled", "openxlsx", "doParallel", "foreach", "knitr")

for (pkg in pkg_nodig) {
  if (system.file(package = pkg) == "") {
    install.packages(pkg)
  }
}

suppressPackageStartupMessages(library(tidyverse))
suppressPackageStartupMessages(library(survey))
library(haven)
library(this.path)
library(textutils)
library(labelled)
library(openxlsx)
suppressPackageStartupMessages(library(doParallel))
library(foreach)
library(knitr)

# instellen werkmap voor het laden van de andere bestanden
setwd(dirname(this.path()))
source("tbl_helpers.R")

# niveau van weergave van meldingen
# mogelijke waarden: ERR / WARN / MSG / DEBUG, waarbij een hoger niveau melding altijd weergegeven wordt
# als het level op MSG staat en er zou een WARN worden weergegeven is deze dus zichtbaar (maar andersom niet)
# in het dagelijks gebruik is MSG ruim voldoende, zet DEBUG alleen aan bij het doorgeven van foutmeldingen aan de werkgroep
log.level = DEBUG
# log opslaan naar een tekstbestand?
log.save = T

# het gehele script wordt omgeven door curly brackets, zodat een stop() ook daadwerkelijk het script stopt
{
  # selecteren configuratiebestand en bijbehorende werkmap
  # hierin dienen de configuratie(.xlsx) en de databestanden (in de map data) te staan
  # andere mappen worden automatisch aangemaakt als deze niet bestaan
  if(!exists("skip_config_popup") || !skip_config_popup){
    config.file = choose.files(caption="Selecteer configuratiebestand...",
                               filters=c("Excel-bestand (*.xlsx)","*.xlsx"),
                               multi=F)    
  } else {
    msg("er wordt een configuratiebestand gebruikt wat buiten tbl_maken.R is gedefinieerd:", config.file)
  }
  
  #config.file = "configuratieVO.xlsx"
  if (!str_ends(config.file, ".xlsx")) msg("Configuratiebestand dient een Excel-bestand te zijn.", ERR)
  setwd(dirname(config.file))
  
  
  # algemene instellingen voor de libraries
  options(survey.lonely.psu="certainty")
  options("openxlsx.paperSize"=9) # A4
  options("openxlsx.orientation"="landscape")
  options("openxlsx.numFmt"="0") # standaardformaat zonder decimalen (anders 0.0 invullen)
  
  # controleren of de configuratie leesbaar is
  tmp = read.xlsx(config.file, sheet="datasets")
  if (!exists("tmp")) msg("Configuratiebestand kon niet gelezen worden. Wellicht is deze nog geopend in Excel?", ERR)
  rm(tmp)
  
  # zorgen dat de juiste mappen bestaan
  if (!dir.exists("output") || !dir.exists("resultaten_csv")) {
    if (!dir.exists("output")) dir.create("output")
    if (!dir.exists("resultaten_csv")) dir.create("resultaten_csv")
    
    msg("Mappenstructuur aangemaakt in de map van het configuratiebestand: /output bevat na afloop de tabellenboeken, /resultaten_csv de resultaten van de berekening.", level=MSG)
  }
  
  # daadwerkelijk inlezen configuratie
  # TODO: onmogelijke waardes checken
  sheets = c("algemeen", "crossings", "datasets", "indeling_rijen", "onderdelen", "opmaak", "labelcorrectie", "logos", "intro_tekst", "headers_afkortingen",
             "dichotoom", "niet_dichotoom", "forceer_datatypen")
  for (sheet in sheets) {
    tmp = read.xlsx(config.file, sheet=sheet)
    
    # om de configuratielast te verlichten is het mogelijk om tabbladen over te erven uit andere configuraties
    # hiervoor dient in cel A1 "KOPIEER" te staan, en in cel A2 een pad naar de gewenste xlsx met een identieke tabbladnaam
    if (ncol(tmp) == 1 && colnames(tmp) == "KOPIEER") {
      if (nrow(tmp) < 1) {
        msg("Tabblad %s zou volgens de configuratie overgenomen moeten worden van een ander bestand, maar er is geen bestand opgegeven.", sheet, level=ERR)
      }
      
      replacement.file = tmp$KOPIEER[1]
      if (!file.exists(replacement.file)) {
        msg("Tabblad %s zou volgens de configuratie overgenomen moeten worden van %s, maar dit bestand is niet gevonden. Let op: relatieve paden worden gelezen vanaf het configuratiebestand.", sheet, replacement.file, level=ERR)
      }
      
      tmp = read.xlsx(replacement.file, sheet=sheet)
    }
    
    if (ncol(tmp) == 1) tmp = tmp[[1]]
    assign(sheet, tmp, envir=.GlobalEnv)
  }
  rm(tmp)
  
  msg("Start verwerking configuratiebestand %s...", config.file, level=MSG)
  
  # controleren of nieuw toegevoegde opties ook zijn ingevuld
  # uiteindelijk zal dit verwijderd worden, maar voor nu is een overgangsperiode wel fijn
  if (!("vraag_verbergen_bij_missend_antwoord" %in% colnames(algemeen) && "afkapwaarde_antwoord" %in% colnames(algemeen))) {
    msg("Let op! Er zijn twee nieuwe opties toegevoegd aan het tabblad algemeen: afkapwaarde_antwoord en vraag_verbergen_bij_missend_antwoord. (Zie de handleiding.)", level=WARN)
    msg("Deze opties missen momenteel in de configuratie. Er wordt nu de standaardwaarde (afkapwaarde = 0, vraag verbergen = WAAR) aangenomen.", level=WARN)
    msg("Het is van belang om deze opties alsnog toe te voegen aan de configuratie; in een volgende versie van het script zal deze waarschuwing vervangen worden door een fout.", level=WARN)
    algemeen$vraag_verbergen_bij_missend_antwoord = F
    algemeen$afkapwaarde_antwoord = 0.0
  }
  
  # bestaat het templatebestand voor digitoegankelijkheid? (indien gewenst)
  if (!is.na(algemeen$template_html)) {
    # in de configuratie mag {script}/ staan, deze vervangen door pad van het script
    algemeen$template_html = str_replace(algemeen$template_html, fixed("{script}"), dirname(this.path()))
    if (!file.exists(algemeen$template_html)){
      msg("Het templatebestand voor digitoegankelijke tabellenboeken is niet gevonden. Controleer de configuratie: tabblad algemeen, kolom template_html.", algemeen$template_html, level=ERR)
    }
    
  }
  
  # sanity checks op indeling_rijen
  if (any(is.na(indeling_rijen$type)) || any(!indeling_rijen$type %in% c("aantallen", "var", "nvar", "titel", "kop", "vraag", "tekst"))) {
    msg("Ongeldige waardes aangetroffen in de kolom type van tabblad indeling_rijen. Het betreft rij(en) %s. Geldige waardes zijn 'aantallen', 'var', 'nvar', 'titel', 'kop', 'vraag', of 'tekst'.",
        str_c(which(is.na(indeling_rijen$type) | !indeling_rijen$type %in% c("aantallen", "var", "nvar", "titel", "kop", "vraag", "tekst")), collapse=", "), level=ERR)
  }
  
  # latere toevoeging: mogelijkheid om kolomkoppen aan en uit te zetten per variabele/kop
  # als deze kolom mist nemen we standaardgedrag aan, dus we kunnen de kolom veilig toevoegen met NA
  if (!"kolomkoppen" %in% colnames(indeling_rijen)) {
    indeling_rijen$kolomkoppen = rep(NA, nrow(indeling_rijen))
    msg("Er is geen kolom met de naam 'kolomkoppen' aanwezig in indeling_rijen. Standaardwaarden (kolomkoppen bij iedere vraag) worden aangenomen.", level=WARN)
  }
  
  # idem voor gewenste waardes; missend is 'normaal', gespecificeerd verandert het gedrag
  # we kunnen dus probleemloos een lege kolom toevoegen
  if (!"waardes" %in% colnames(indeling_rijen)) {
    indeling_rijen$waardes = rep(NA, nrow(indeling_rijen))
    msg("Er is geen kolom met de naam 'waardes' aanwezig in indeling_rijen. Standaardinstelling (alle antwoorden in oplopende volgorde) wordt aangenomen.", level=WARN)
  }
  
  # hovertekst bij significantie is ook later toegevoegd
  if (!"sign_hovertekst" %in% colnames(algemeen)) {
    algemeen$sign_hovertekst = "Deze waarde is significant anders."
    msg("Er is geen kolom met de naam 'sign_hovertekst' aanwezig in algemeen. Standaardinstelling (\"Deze waarde is significant anders.\") wordt aangenomen.", level=WARN)
  }
  
  # idem voor sign_verbergen_wanneer_afgerond_gelijk
  if (!"sign_verbergen_wanneer_afgerond_gelijk" %in% colnames(algemeen)) {
    algemeen$sign_verbergen_wanneer_afgerond_gelijk = F
    msg("Er is geen kolom met de naam 'sign_verbergen_wanneer_afgerond_gelijk' aanwezig in algemeen. Standaardinstelling (ONWAAR) wordt aangenomen.", level=WARN)
  }
  
  # en sign_doelkolom bij onderdelen
  if (!"sign_doelkolom" %in% colnames(onderdelen)) {
    onderdelen$sign_doelkolom = F
    msg("Er is geen kolom met de naam 'sign_doelkolom' aanwezig in onderdelen. Standaardinstelling (ONWAAR) wordt aangenomen.", level=WARN)
  } else {
    onderdelen$sign_doelkolom[is.na(onderdelen$sign_doelkolom)] = F
  }
  
  # en kolomnamen bij onderdelen
  if (!"kolomnaam" %in% colnames(onderdelen)) {
    onderdelen$kolomnaam = NA
  }
  
  # idem voor id bij logos
  if (!"id" %in% colnames(logos)) {
    logos$id = NA
  }
  
  # en verberg_crossings bij indeling_rijen
  if (!"verberg_crossings" %in% colnames(indeling_rijen)) {
    indeling_rijen$verberg_crossings = NA
  }
  
  # variabelelijst afleiden uit de indeling van het tabellenboek;
  # iedere regel met (n)var is een variabele die we nodig hebben
  varlist = indeling_rijen[indeling_rijen$type %in% c("var", "nvar"),]
  if (any(duplicated(varlist$inhoud))) {
    msg("Let op: variabele(n) %s komen meerdere keren voor in indeling_rijen. Controleer of dit de bedoeling is.",
        str_c(varlist$inhoud[duplicated(varlist)], collapse=", "), level=WARN)
  }
  if (any(is.na(varlist$inhoud))) {
    msg("Let op: regel(s) %s bevat geen variabele. Controleer de configuratie.", level=ERR)
  }
  # mogelijke weegfactoren opslaan voor later gebruik (zie de berekeningsloop)
  weight.factors = indeling_rijen %>%
    select(starts_with("weegfactor")) %>%
    lapply(function (v) {
      output = unique(v)
      return(output[!is.na(output)])
    }) %>% unlist() %>% unname()
  
  # in de sheet 'crossings' zijn twee mogelijke kolommen: varname van crossing, en 
  # boolean voor toetsen van deze crossing. We splitsen deze voor backward compatability
  if(length(crossings) == 0){ # scenario 1: een sheet met alleen 'crossings' in A1 (default)
    crossings_toetsen <- NULL
  } else if(is.null(colnames(crossings))){ # scenario 2: een sheet met 'crossings' in A1 en daaronder varnames
    crossings_toetsen <- rep(T, length(crossings))
    names(crossings_toetsen) <- crossings
  } else if("toetsen" %in% colnames(crossings)){ # scenario 3: een sheet met 'crossings' in A1 en daaronder varnames en 'toetsen' in B1 en daaronder lijst met booleans
    crossings_toetsen <- crossings$toetsen
    crossings <- crossings$crossing
    names(crossings_toetsen) <- crossings
  } else{
    msg("Geen kloppend tabblad crossings. Deze moet bestaan uit één of twee kolommen; crossings en (optioneel) toetsen. Standaard is 'crossings' in cel A1 en verder een lege sheet. Voor meer informatie zie documentatie.", level=ERR)
  }

  if (any(crossings %in% onderdelen$subset)) {
    msg("De variabele(n) %s is/zijn ingevuld als crossing en subset. Een subset kan niet met zichzelf gekruist worden.",
        str_c(crossings[crossings %in% onderdelen$subset], ", "), level=WARN)
  }
  
  if (any(crossings %in% varlist$inhoud)) {
    msg("De variabele(n) %s is/zijn ingevuld als crossing en variabele in indeling_rijen. Een variabele kan niet met zichzelf gekruist worden.",
        str_c(crossings[crossings %in% varlist$inhoud], ", "), level=WARN)
  }
  
  # datasets combineren en de strata en weegfactoren apart opslaan
  data.combined = data.frame()
  for (d in 1:nrow(datasets)) {
    
    msg("Inladen dataset %d: %s", d, datasets$naam_dataset[d], level=MSG)
    
    # zoek naar nieuw te maken datasets die een parent dataset hebben, geef ze alvast goed data path
    #splitted_naam_dataset <- strsplit(datasets$naam_dataset[d], "")[[1]]
    if(!is.numeric(datasets$naam_dataset[d]) && strsplit(datasets$naam_dataset[d], "")[[1]][1] == "_"){
      is_filtered_dataset <- T
      # Haal eigenschappen voor de gefilterde dataset uit de string
      parent_dataset <- strsplit(datasets$naam_dataset[d], "_")[[1]][2]
      parent_dataset_row <- which(datasets$naam_dataset == parent_dataset)
      new_name <- strsplit(datasets$naam_dataset[d], "_")[[1]][3]
      # sla de eigenschappen op
      datasets$bestandsnaam[d] <- datasets$bestandsnaam[parent_dataset_row]
      datasets$naam_dataset[d] <- new_name
    } else {
      is_filtered_dataset <- F
    }
    
    # .sav aan het einde kan vergeten worden...
    if (!str_ends(datasets$bestandsnaam[d], fixed(".sav"))) {
      datasets$bestandsnaam[d] = paste0(datasets$bestandsnaam[d], ".sav")
      msg("De opgegeven dataset op rij %d eindigt niet in .sav. Mogelijk vergeten? .sav toegevoegd.", level=WARN, d+1)
    }
    
    # SPSS slaat 'gekke' leestekens (bijv. de Spaanse n met een ~, de Noorse o met een schuine streep, etc.) soms onjuist op
    # dit zorgt dan voor een vastloper
    # dit kunnen we voorkomen door encoding te forceren, maar tekst kan hierdoor wel aangepast worden
    # gelukkig maakt het tabellenboek geen gebruik van open tekst, dus de schade zou minimaal moeten zijn
    # zie ook de issue op https://github.com/tidyverse/haven/issues/615
    data = tryCatch(read_spss(datasets$bestandsnaam[d], user_na=T) %>% user_na_to_na(),
                    error=function(e) { 
                      if (str_detect(as.character(e), "encoding")) {
                        read_sav(datasets$bestandsnaam[d], user_na=T, encoding="latin1") %>% user_na_to_na()
                      } else {
                        msg("Fout tijdens het inlezen van dataset %d: %s", d, e, level=ERR)
                      }
                    })
    
    # als dit een dataset is met een parent dataset, voer hier de filtering uit.
    # deze filtering staat in
    if(is_filtered_dataset){
      data <- data %>% filter(!!rlang::parse_expr(datasets$filter[d]))
      cat(" dit is een gefilterde dataset met filter", datasets$filter[d], "op dataset", parent_dataset, ". Nieuwe naam is", new_name, "\n")
      if(nrow(data) == 0){
        msg("Dit filter resulteert in een lege dataset", level=ERR)
      }
    }
    
    # afwijkende kolommen registreren zodat we deze later kunnen scheiden
    afwijkend = c()
    for (c in 1:ncol(data)) {
      colname = colnames(data)[c]
      
      # in sommige datasets is een eigenlijk numerieke code omgezet naar een string (bijv. voor combinatie met CBS)
      # om dit te corrigeren naar een bruikbare combinatie met oudere datasets kan het datatype geforceerd worden
      if (nrow(forceer_datatypen) > 0 && colname %in% forceer_datatypen$variabele) {
        # mogelijke opties: 'numeric' of 'character'
        desired = forceer_datatypen$type[forceer_datatypen$variabele == colname]
        
        if (desired == "numeric") {
          if (!typeof(data[[c]]) %in% c("double", "integer")) {
            # je zou denken dat hier een functie voor was, maar die is er dus niet
            var_label = var_label(data[[c]])
            if (!is.null(val_labels(data[[c]]))) {
              old_labels = data.frame(val=as.numeric(unname(val_labels(data[[c]]))), label=names(val_labels(data[[c]])))
              old_values = to_character(data[[c]])
              new_values = sapply(old_values, function (value) { return(old_labels$val[which(old_labels$label == value)]) })
              # in sommige versies van R volgt hier alsnog een lijst uit, ook al is sapply() de vectorversie
              # als dat gebeurt, en er een missing is in de waardes, dan komt er een NULL in de lijst, ipv een NA
              # hierdoor komt er uit unlist() dan een vector met minder waardes dan de originele lijst, en dan explodeert alles
              # zie ook https://stackoverflow.com/questions/60812923/unlist-a-list-without-losing-nulls
              # oplossing: indien new_values een lijst is, vervangen we alle NULLs door NA, en dan klopt het weer
              # dit kan door lengths(x) == 0 te doen
              if (is.list(new_values)) {
                new_values[lengths(new_values) == 0] = NA
                new_values = new_values %>% unlist()
              }
              new_labels = old_labels$val
              names(new_labels) = old_labels$label
              new_labels = sort(new_labels)
              data[[c]] = labelled(new_values, labels=new_labels, label=var_label)
            } else {
              new_values = as.numeric(data[[c]])
              data[[c]] = labelled(new_values, label=var_label)
            }
          }
        } else if (desired == "character") {
          if (typeof(data[[c]]) != "character")
            print(c)
          var_label = var_label(data[[c]])
          data[[c]] = to_character(data[[c]])
          var_label(data[[c]]) = var_label
        }
      }
      
      # strata willen we apart opslaan
      if (!is.na(datasets$stratum[d]) && colname == datasets$stratum[d]) {
        data$tbl_strata = data[[c]]
        next
      }
      
      # gewichten apart opslaan
      if (!is.na(datasets$weegfactor[d]) && colname == datasets$weegfactor[d]) {
        data$tbl_weegfactor = data[[c]]
        next
      }
      
      # jaren apart opslaan
      if (!is.na(datasets$jaarvariabele[d]) && colname == datasets$jaarvariabele[d]) {
        data$tbl_jaar = data[[c]]
        next
      }
      
      # checken of de gevraagde subset aanwezig is in de dataset
      if(!is.na(onderdelen$subset[d])){
        if(!onderdelen$subset[d] %in% colnames(data)){
          msg("De variabele(n) %s komen niet voor in de dataset. Deze is in de configuratie opgegeven als subset, en daarmee verplicht. Pas de configuratie aan of voeg de variabele(n) toe aan de dataset.",
              onderdelen$subset[d], level=ERR)
        }
      }
      
      if (colname %in% colnames(data.combined)) {
        existing = data.combined[[which(colnames(data.combined) == colname)]]
        # controleren of het type en, indien relevant, factors overeenkomen
        # hierbij kan er tussen jaren best wat verschil zitten in het label, dus de labelcheck is optioneel
        if (algemeen$vergelijk_variabelelabels) {
          if (!identical(var_label(data[[c]]), var_label(existing))) {
            afwijkend = c(afwijkend, c)
          } 
        } else {
          # zorg dat het oude label bewaard blijft, aangezien de eerste dataset waarschijnlijk de meest recente is
          var_label(data[[c]]) = var_label(existing)
        }
        
        if (is.null(val_labels(data[[c]])) && is.null(val_labels(existing))) next
        
        # weer zo'n gek ding: als je val_labels(x) == val_labels(y) doet met verschillende lengtes
        # telt het alsnog als TRUE, zo lang de overeenkomende waardes maar gelijk zijn
        if (length(val_labels(data[[c]])) != length(val_labels(existing))) {
          msg("Antwoordopties voor variabele %s komen niet overeen tussen datasets. In dataset %s zijn %d optie(s) aanwezig (%s), in de datasets daarvoor %d (%s). Deze waardes worden gecombineerd met de eerder bekende labels als leidraad.",
              colname, datasets$naam_dataset[d], length(val_labels(data[[c]])), str_c(names(val_labels(data[[c]])), collapse=", "),
              length(val_labels(existing)), str_c(names(val_labels(existing)), collapse=", "), level=WARN)
          label_problemen <- bind_rows(
            label_problemen,
            data.frame(
              kolom=colname,
              nieuwe_dataset=datasets$naam_dataset[d],
              nieuwe_labels=paste0(val_labels(data[[c]]), collapse = ";"),
              oude_labels=paste0(val_labels(existing), collapse = ";")
            )
          )
        }
        
        # let op: warnings negeren is niet heel netjes, maar direct hierboven is met een if gecontroleerd of het aantal afwijkt
        # deze if geeft een warning als het aantal ongelijk is... beetje dubbelop dus
        if (suppressWarnings(!all(val_labels(data[[c]]) == val_labels(existing)))) {
          afwijkend = c(afwijkend, c)
        }
      }
    }
    afwijkend = unique(afwijkend)
    
    # herschrijven kolomnamen zodat ze niet gaan storen
    if (length(afwijkend) > 0) {
      msg("Afwijkende kolommen in dataset %d: %s", d, str_c(colnames(data)[afwijkend], collapse=","), level=WARN)
      colnames(data)[afwijkend] = paste0("_", d, "_", colnames(data)[afwijkend])
    }
    
    # zijn er weegfactoren aangetroffen? zo nee, is dat de bedoeling?
    if (!("tbl_weegfactor" %in% colnames(data))) {
      if (is.na(datasets$weegfactor[d]) && is.na(datasets$stratum[d])) {
        # geen stratum en weegfactor opgegeven; alles invullen met 1
        data$tbl_weegfactor = rep(1, nrow(data))
      } else if (is.na(datasets$weegfactor[d]) && !is.na(is.na(datasets$stratum[d]))) {
        data$tbl_weegfactor = rep(1, nrow(data))
        msg("Let op! In dataset %s is wel een stratum aangegeven, maar geen weegfactor. Dit is alleen mogelijk wanneer er per variabele een weegfactor wordt aangegeven in indeling_rijen. Controleer of dit de bedoeling is.",
            datasets$naam_dataset[d], level=WARN)
      } else {
        msg("Kolom %s of %s is niet aangetroffen in dataset %d (%s). Laat voor een ongewogen design beide velden leeg of pas de dataset aan.",
            datasets$weegfactor[d], datasets$stratum[d], d, datasets$naam_dataset[d], level=ERR)
      }
    }
    
    if (!("tbl_strata" %in% colnames(data))) {
      if (is.na(datasets$stratum[d])) {
        data$tbl_strata = rep(1, nrow(data))
      } else {
        msg("In dataset %s is de kolom %s voor strata niet aangetroffen. Hierdoor kan er geen gewogen design worden gebruikt",
            datasets$naam_dataset[d], datasets$stratum[d], level=ERR)
      }
    }
    
    # jaarvariabele aanwezig?
    if (!("tbl_jaar" %in% colnames(data))) {
      if (!is.na(datasets$jaarvariabele[d])) {
        msg("In dataset %s is geen jaarvariabele %s aangetroffen. Controleer de configuratie.",
            datasets$naam_dataset[d], datasets$jaarvariabele[d], level=ERR)
      }
    }
    
    data[,"tbl_dataset"] = d
    data.combined = bind_rows(data.combined, data)
    
    # bind_rows haalt soms willekeurig labels weg, dus die moeten we handmatig terugzetten
    for (c in 1:ncol(data)) {
      colname = colnames(data)[c]
      var_label(data.combined[[colname]]) = var_label(data[[colname]])
    }
  }
  data = data.combined
  rm(data.combined) # twee keer hetzelfde object is zonde van het geheugen, en data.combined blijven gebruiken geeft ons RSI; data is korter
  
  # labels opslaan
  var_labels = data.frame()
  for (i in 1:ncol(data)) {
    colname = colnames(data)[i]
    if (!is.null(var_label(data[[i]]))) {
      var_labels = bind_rows(var_labels, data.frame(var=colname, val="var", label=var_label(data[[i]])))
    } else {
      var_labels = bind_rows(var_labels, data.frame(var=colname, val="var", label=NA))
    }
    labels = val_labels(data[[i]])
    if (!is.null(labels)) {
      var_labels = bind_rows(var_labels, data.frame(var=rep(colname, length(labels)), val=as.character(unname(labels)), label=names(labels)))
    }
  }
  
  # zijn er labels om te veranderen?
  if (nrow(labelcorrectie) > 0) {
    for (i in 1:nrow(labelcorrectie)) {
      if (!is.na(labelcorrectie$var[i]) && !(labelcorrectie$var[i] %in% var_labels$var)) {
        msg("Variabele %s dient volgens de configuratie een nieuw (antwoord)label te krijgen, maar deze is niet aangetroffen in de dataset. (Rij %d.)",
            labelcorrectie$var[i], i, level=WARN)
        next
      }
      
      if (!is.na(labelcorrectie$var[i]) && !is.na(labelcorrectie$var_label[i])) {
        var_labels$label[var_labels$var == labelcorrectie$var[i] & var_labels$val == "var"] = labelcorrectie$var_label[i]
      }
      
      # antwoorden zijn te vervangen op basis van puur tekst, of op basis van variabele
      # hierdoor kan het zijn dat labelcorrectie$var leeg is, maar antwoord_oud niet
      if (!is.na(labelcorrectie$antwoord_nieuw[i])) {
        if (!is.na(labelcorrectie$var[i]) && !(labelcorrectie$var[i] %in% var_labels$var)) {
          msg("Variabele %s dient volgens de configuratie een nieuw antwoordlabel te krijgen, maar deze is niet aangetroffen in de dataset. (Rij %d.)",
              labelcorrectie$var[i], i, level=WARN)
          next
        }
        
        if (is.na(labelcorrectie$antwoord_waarde[i]) && !is.na(labelcorrectie$antwoord_oud[i])) {
          if (is.na(labelcorrectie$var[i])) {
            var_labels$label = str_replace(var_labels$label, fixed(labelcorrectie$antwoord_oud[i]), fixed(labelcorrectie$antwoord_nieuw[i]))
          } else {
            var_labels$label[var_labels$var == labelcorrectie$var[i]] = str_replace(var_labels$label[var_labels$var == labelcorrectie$var[i]],
                                                                                    fixed(labelcorrectie$antwoord_oud[i]), fixed(labelcorrectie$antwoord_nieuw[i]))
          }
        } else if (!is.na(labelcorrectie$antwoord_waarde[i])) {
          var_labels$label[var_labels$var == labelcorrectie$var[i] & var_labels$val == labelcorrectie$antwoord_waarde[i]] = labelcorrectie$antwoord_nieuw[i]
        } else {
          msg("Labelcorrectie in rij %d is onmogelijk; er is geen waarde of oud antwoord opgegeven.", i, level=WARN)
        }
      }
    }
  }
  
  # dummies maken
  for (c in 1:ncol(data)) {
    colname = colnames(data)[c]
    if (!(colname %in% varlist$inhoud) && !(colname %in% crossings)) next
    if (str_starts(colname, "_") || str_starts(colname, "dummy")) next
    
    if (is.null(val_labels(data[[c]]))) next
    
    labels = val_labels(data[[c]])
    for (i in labels) {
      data[,paste0("dummy.", colname, ".", i)] = data[[c]] == unname(i)
    }
  }
  
  # superstrata maken en fpc correctiefactoren
  # superstrata
  # dit doen we door de bestaande strata te nummeren van 1-<aantal> en daar 1000 * dataset bij te doen
  # dus stratum 40 uit dataset 1 wordt 1040, stratum 25 uit dataset 2 wordt 2025, enz.
  # fpc correctie factoren
  # dit doen we uit te zoeken of we een fpc bestand hebben voor elke dataset,
  # dan dit te laden en te matchen met de oorspronkelijke stratumkolom (dus niet superstratum!)
  data$superstrata = NA
  data$fpc = NA
  fpc_data = NULL
  for (d in 1:nrow(datasets)) {
    message(d)
    dataset_columns = which(data$tbl_dataset == d) # welke kolommen in data gaan over dataset d, 'mijn kolommen'?
    strata = data$tbl_strata[dataset_columns] # wat zijn de strata van mijn kolommen?
    data$tbl_strata[dataset_columns] = paste0(strata, "_d", d) # pas strata namen aan zodat ze niet verward worden tussen datasets.
    strata = data$tbl_strata[dataset_columns] # opnieuw binnenhalen voor later gebruik
    unique_strata = sort(unique(strata)) # unique strata
    if("fpc" %in% colnames(datasets) && !is.na(datasets$fpc[d])){
      if(is.na(datasets$stratum[d])){stop("Bij FPC is het verplicht een stratum op te geven.")}
      fpc_data <- table(strata) %>%
        as.data.frame %>% 
        rename(stratum = strata)
      if(file.exists(datasets$fpc[d])){
        # als het een pad is: zoek de fpc data op en berekenen sampling prob per stratum
          fpc_data <- fpc_data %>% left_join(
            read.xlsx(datasets$fpc[d]) %>% 
              mutate(stratum = as.factor(stratum)),
            join_by(stratum == stratum)
          ) %>% 
          mutate(fpc = populatiegrootte)
      } else if(grepl("GROOTGEWICHT_", datasets$fpc[d])){
        if(is.na(datasets$stratum[d])){stop("Bij FPC is het verplicht een stratum op te geven.")}
        # als er grote gewichten inzitten (die weergeven hoeveel mensen een respondent vertegenwoordigd)
        # dan kan je populatiegroottes afleiden daaruit
        
        fpc_data <- fpc_data %>% 
          left_join(
            data %>%
              group_by(tbl_strata) %>%
              summarise(populatiegrootte = sum(!!sym(gsub("GROOTGEWICHT_", "", datasets$fpc[d])), na.rm = TRUE)), 
            join_by(stratum == tbl_strata)
          ) %>% 
          mutate(
            fpc = populatiegrootte
          )
        
      } else if(datasets$fpc[d] %in% colnames(data)){
        # als de fpc correctiefactor gewoon een kolom is, dan die overnemen
        fpc_per_respondent <- data.frame(fpc = data %>% select(!!sym(datasets$fpc[d])))
      } else {
        # anders: onbekende methode, geef fout
        msg("FPC is aangegeven maar onbekende FPC kolom input '%s'. Zie handleiding voor opties.", datasets$fpc[d], level=ERR)
      }
      # merge zodat we voor elke respondent een sampling prob hebben
      fpc_per_respondent <- data.frame(
        stratum = strata %>% as.factor
      ) %>% left_join(
        fpc_data
      )
      
      # voeg de fpc factor toe aan de data
      data$fpc[dataset_columns] <- fpc_per_respondent$fpc
    } else {
      fpc_data <- NULL
    }
    for (i in 1:length(unique_strata)) {
      data$superstrata[dataset_columns & data$tbl_strata == unique_strata[i]] = d*1000 + i
    }
  }
  
  if("fpc" %in% colnames(datasets) && any(is.na(data$fpc))){
    msg("FPC is maar voor een deel van datasets ingegeven. Er wordt aangenomen dat FPC voor andere datasets niet nodig is.", level=MSG)
    data$fpc[is.na(data$fpc)] <- 1e6
  }
  
  data$superweegfactor = data$tbl_weegfactor
  # missende weegfactoren mag niet, wat is het beleid?
  # mogelijke configuraties:
  # - verwijderen (rijen met missend gewicht eruit)
  # - fout (foutmelding geven)
  # - getal (vervangende waarde invoeren)
  if (any(is.na(data$superweegfactor))) {
    n = sum(is.na(data$superweegfactor))
    if (algemeen$missing_weegfactoren == "fout") {
      msg("Er zijn %d missende weegfactoren gevonden. Volgens de configuratie dient er dan te worden gestopt met uitvoering.\nPas dit aan in de data of de configuratie.",
          n, level=ERR)
    } else if (algemeen$missing_weegfactoren == "verwijderen") {
      msg("Er zijn %d missende weegfactoren gevonden. Deze rijen zijn verwijderd uit de dataset, zoals aangegeven in de configuratie.",
          n, level=MSG)
      data = data[-which(is.na(data$superweegfactor)),]
    } else if (str_detect(algemeen$missing_weegfactoren, "^\\d+")) {
      msg("Er zijn %d missende weegfactoren gevonden. Deze waarden zijn vervangen door %s, zoals aangegeven in de configuratie.",
          n, algemeen$missing_weegfactoren, level=MSG)
      data$superweegfactor[is.na(data$superweegfactor)] = as.numeric(algemeen$missing_weegfactoren)
    } else {
      msg("Er zijn %d missende weegfactoren gevonden. Er is geen valide configuratie opgegeven om hiermee om te gaan. Zie de handleiding voor meer informatie.",
          n, level=ERR)
    }
  }
  
  # zijn alle benodigde variabelen aanwezig?
  required.vars = c(weight.factors, crossings, unique(onderdelen$subset))
  required.vars = required.vars[!is.na(required.vars)]
  if (!all(required.vars %in% colnames(data))) {
    msg("De variabele(n) %s komen niet voor in de dataset. Deze zijn in de configuratie opgegeven als weegfactor, crossing of subset, en daarmee verplicht. Pas de configuratie aan of voeg de variabele(n) toe aan de dataset.",
        str_c(required.vars[!required.vars %in% colnames(data)], collapse=", "), level=ERR)
  }
  
  
  msg("Datasets gecombineerd. Start verwerking kolomopbouw.", level=DEBUG)
  
  # door de onderdelen loopen en die verwerken
  kolom_opbouw = data.frame()
  test.col.cache = data.frame()
  for (i in 1:nrow(onderdelen)) {
    if (!(onderdelen$dataset[i] %in% datasets$naam_dataset)) {
      msg("Onderdeel %d (dataset %s, subset %s) bevat een ongeldige datasetnaam. Uitvoering gestopt.",
          i, onderdelen$dataset[i], onderdelen$subset[i], level=ERR)
    }
    d = which(datasets$naam_dataset == onderdelen$dataset[i])
    
    # als met_crossing == WAAR, iedere crossing los toevoegen
    # daarna altijd eentje met crossing = NA voor een totaalkolom
    if (onderdelen$met_crossing[i]) {
      if (algemeen$sign_toetsen && !is.na(onderdelen$sign_crossing[i])) {
        # uitzoeken welke kolom tegenhanger moet zijn van de chi square
        if (onderdelen$sign_crossing[i] == "intern") {
          test.col = 0 # 0 = toetsen met andere waarden van de crossing, als waarde toetsen in sheet crossings != FALSE
        } else {
          # geen toets
          test.col = NA
        }
      } else {
        # geen toets
        test.col = NA
      }
      for (crossing in crossings) {
        crossing.labels = var_labels[var_labels$var == crossing & var_labels$val != "var",]
        n = nrow(crossing.labels)
        
        test.col.current = test.col
        if (!crossings_toetsen[crossing]) {
          test.col.current = NA
        }
        
        kolom_opbouw = bind_rows(kolom_opbouw, data.frame(col.index=nrow(kolom_opbouw)+(1:n), dataset=rep(d, n),
                                                          subset=rep(onderdelen$subset[i], n), year=rep(onderdelen$jaar[i], n),
                                                          crossing=rep(crossing, n), crossing.val=as.numeric(crossing.labels$val),
                                                          crossing.lab=crossing.labels$label, test.col=rep(test.col.current, n), test.display=T))
      }
    }
    
    # totalen toetsen?
    if (algemeen$sign_toetsen && !is.na(onderdelen$sign_totaal[i])) {
      if (str_detect(onderdelen$sign_totaal[i], "^\\d+$")) {
        test.col = as.numeric(onderdelen$sign_totaal[i])
      } else if (onderdelen$sign_totaal[i] %in% datasets$naam_dataset) {
        # we weten nog niet in welke kolom een dataset eindigt
        # daarom een aparte lijst met indexen die later gevuld moeten worden
        test.col = NA
        test.col.cache = bind_rows(test.col.cache, data.frame(col.index=nrow(kolom_opbouw)+1, test.col=onderdelen$sign_totaal[i]))
      } else {
        msg("Let op! Onderdeel %d (dataset %s, subset %s) heeft een ongeldige waarde (%s) in de kolom sign_totaal. Hier is alleen een getal, de naam van een dataset, of een lege cel toegestaan. Indien hier een naam is opgegeven, controleer dan of deze klopt.",
            i, onderdelen$dataset[i], onderdelen$subset[i], onderdelen$sign_totaal[i], level=ERR)
      }
    } else {
      test.col = NA
    }
    
    kolom_opbouw = bind_rows(kolom_opbouw, data.frame(col.index=nrow(kolom_opbouw)+1, dataset=d, subset=onderdelen$subset[i], year=onderdelen$jaar[i],
                                                      crossing=NA, crossing.val=NA, crossing.lab=NA, test.col=test.col, test.display=!onderdelen$sign_doelkolom[i],
                                                      name=onderdelen$kolomnaam[i]))
  }
  
  # moeten er nog testkolommen toegevoegd worden uit de cache?
  if (nrow(test.col.cache) > 0) {
    for (i in 1:nrow(test.col.cache)) {
      # de laatste kolom uit de dataset is altijd totaal, dus hoogste waarde is direct de juiste kolom
      kolom_opbouw$test.col[test.col.cache$col.index[i]] = max(which(kolom_opbouw$dataset == which(datasets$naam_dataset == test.col.cache$test.col[i])))
    }
  }
  
  # dummyvariabelen maken voor iedere kolom
  # dit is nodig om later een chi square uit te kunnen voeren over een vergelijkingsset, omdat survey pure pijn is
  # daarnaast maken we van het moment gebruik om de aantallen even op te slaan
  n_resp = data.frame()
  for (i in 1:nrow(kolom_opbouw)) {
    kolom = data$tbl_dataset == kolom_opbouw$dataset[i]
    # scheiden per jaar?
    if (!is.na(kolom_opbouw$year[i])) {
      kolom = kolom & data$tbl_jaar == kolom_opbouw$year[i]
    }
    # crossing?
    if (!is.na(kolom_opbouw$crossing[i])) {
      kolom = kolom & data[[kolom_opbouw$crossing[i]]] == kolom_opbouw$crossing.val[i]
    }
    
    data[,paste0("dummy._col", i)] = kolom
    
    n_resp = bind_rows(n_resp, data.frame(col=i, year=kolom_opbouw$year[i], crossing=kolom_opbouw$crossing[i], n=sum(kolom, na.rm=T)))
    
    # ook splitsen per subset?
    if (!is.na(kolom_opbouw$subset[i])) {
      subsetvals = val_labels(data[[kolom_opbouw$subset[i]]])
      if (is.null(subsetvals))
        msg("In kolom %d wordt gesplitst op subset %s, maar deze variabele heeft geen labels in de dataset. Hierdoor kunnen de resultaten niet worden berekend. Voeg labels toe aan de dataset, of pas de kolomindeling aan op het tabblad onderdelen.",
            i, kolom_opbouw$subset[i], level=ERR)
      
      for (val in subsetvals) {
        subset = kolom & data[[kolom_opbouw$subset[i]]] == val
        if (sum(subset, na.rm=T) <= 0) next # als er geen data bestaat voor die subset is het een beetje nutteloos
        data[,paste0("dummy._col", i, ".s.", unname(val))] = subset
        
        n_resp = bind_rows(n_resp, data.frame(col=i, year=kolom_opbouw$year[i], crossing=kolom_opbouw$crossing[i], subset=val, n=sum(subset, na.rm=T)))
      }
    }
  }
  # het kan voorkomen dat geen van de kolommen een subset hebben, maar hier wordt in latere functies wel gebruik van gemaakt... indien missend, voeg toe
  if (!"subset" %in% colnames(n_resp)) {
    n_resp$subset = NA
  }
  
  ##### begin berekeningen
  
  # mochten er subsets zijn, dan is het efficiënter om 1x te berekenen welke matches hierin bestaan
  # een tabellenboek wordt meestal gemaakt voor een geografische eenheid (gemeente/regio/GGD-gebied)
  # daarbij kan het zijn dat een vergelijkingskolom wordt toegevoegd van een groter gebied (bijv. gemeente Apeldoorn vs. subregio Midden-IJssel)
  # of er kan een vergelijkingskolom zijn van dezelfde eenheid uit een ander jaar (bijv. Apeldoorn 2022 vs. Apeldoorn 2019)
  # we gaan er daarom vanuit dat de eerste subset (in Excel de bovenste rij) de leidende is, en dat de rest gebaseerd is op die subset
  if (any(!is.na(kolom_opbouw$subset))) {
    leadingsubset = kolom_opbouw$subset[which(!is.na(kolom_opbouw$subset))[1]]
    leadingcol = min(kolom_opbouw$col.index[which(!is.na(kolom_opbouw$subset))[1]])
    
    subsetvals = val_labels(data[[leadingsubset]])
    
    if (sum(!is.na(unique(kolom_opbouw$subset))) > 1) {
      # bij meer dan één subset de tweede af laten hangen van de eerste
      possiblesubsets = unique(kolom_opbouw$subset[which(!is.na(kolom_opbouw$subset) & kolom_opbouw$subset != leadingsubset)])
      
      subsetmatches = matrix(nrow=length(subsetvals), ncol=length(possiblesubsets)+1)
      colnames(subsetmatches) = c(leadingsubset, possiblesubsets)
      subsetmatches[,1] = subsetvals
      rownames(subsetmatches) = names(subsetvals)
      
      # mogelijke matches doorlopen
      for (s in 1:length(subsetvals)) {
        for (p in 1:length(possiblesubsets)) {
          matches = as.numeric(unique(data[[possiblesubsets[p]]][data[[leadingsubset]] == subsetvals[s]]))
          matches = matches[!is.na(matches)]
          
          if (length(matches) > 1) {
            msg("Let op: aanvullende subset %s heeft %d mogelijkheden in de dataset: %s. Dit is niet mogelijk voor berekening; enkel eerste optie gebruikt.",
                possiblesubsets[p], length(matches), str_c(matches, collapse=", "), level=WARN)
            matches = matches[1]
          } else if (length(matches) == 0) {
            matches = NA
          }
          
          subsetmatches[s,p+1] = matches
        }
      }
    } else {
      subsetmatches = matrix(subsetvals, ncol=1)
      colnames(subsetmatches) = leadingsubset
      rownames(subsetmatches) = names(subsetvals)
    }
  } else {
    subsetmatches = NULL
  }
  
  # zijn er eerder resultaten gemaakt uit dit configuratiebestand met een gelijke indeling?
  calc.results = T
  if (file.exists(sprintf("resultaten_csv/results_%s.csv", basename(config.file))) &&
      file.exists(sprintf("resultaten_csv/settings_%s.csv", basename(config.file))) &&
      file.exists(sprintf("resultaten_csv/vars_%s.csv", basename(config.file)))) {
    results = read.csv(sprintf("resultaten_csv/results_%s.csv", basename(config.file)), fileEncoding="UTF-8")
    kolom_opbouw.prev = read.csv(sprintf("resultaten_csv/settings_%s.csv", basename(config.file)), fileEncoding="UTF-8")
    varlist.prev = read.csv(sprintf("resultaten_csv/vars_%s.csv", basename(config.file)), fileEncoding="UTF-8")
    
    # kolomkoppen en waardes mogen verschillen tussen varlist en varlist.prev, dus we maken een 'vergelijkingsset' zonder die kolommen
    varlist.cmp = varlist %>% select(inhoud, starts_with("weeg"))
    varlist.prev.cmp = varlist.prev %>% select(inhoud, starts_with("weeg"))
    
    # sign_doelkolom en kolomnaam zijn nieuw; deze mogen missen
    # daarnaast mag crossing.lab anders zijn; deze kan in de configuratie aangepast zijn
    kolom_opbouw.cmp = kolom_opbouw %>% select(col.index, dataset, subset, year, crossing, crossing.val, test.col)
    kolom_opbouw.prev.cmp = kolom_opbouw.prev %>% select(col.index, dataset, subset, year, crossing, crossing.val, test.col)
    
    if (identical.enough(kolom_opbouw.cmp, kolom_opbouw.prev.cmp) && identical.enough(varlist.cmp, varlist.prev.cmp)) {
      if("forceer_berekening" %in% colnames(algemeen) && algemeen$forceer_berekening){
        msg("Eerdere resultaten aangetroffen vanuit deze configuratie (%s). Deze zullen opnieuw worden gegenereerd omdat de optie forceer_berekening in tabblad algemeen WAAR is.",
            basename(config.file), level=MSG)
      } else {
        msg("Eerdere resultaten aangetroffen vanuit deze configuratie (%s). Berekening wordt overgeslagen. Indien er nieuwe data is toegevoegd, verwijder dan de bestanden uit de map resultaten_csv.",
            basename(config.file), level=MSG)
        
        calc.results = F
        # aangezien een tweede subset vaker kan draaien moeten we hier een distinct op doen - deze is tijdelijk nodig door een ontwerpfoutje
        # deze functie werd eerst meerdere keren aangeroepen in tbl_MakeExcel, wat natuurlijk veel meer resources kost
        # helaas zijn de resultaten die inmiddels opgeslagen zijn wel volgens de oude manier berekend, dus moeten we de correctie voor de zekerheid uitvoeren
        results = results %>% distinct()
      }
      
    } else {
      msg("Er zijn eerdere resultaten aangetroffen vanuit deze configuratie (%s), maar de instellingen waren niet identiek. Berekening wordt opnieuw uitgevoerd.", basename(config.file), level=MSG)
    }
  }
  
  if (calc.results) {
    source(paste0(dirname(this.path()), "/tbl_GetTableRow.R"))
    results = data.frame()
    t.start = proc.time()["elapsed"]
    t.vars = c()
    for (i in 1:nrow(varlist)) {
      var = varlist$inhoud[i]
      
      if (!var %in% colnames(data)) {
        msg("Variabele %s is wel opgegeven in indeling_rijen, maar komt niet voor in de dataset. Controleer de configuratie.", var, level=WARN)
        next
      }
      
      # we kunnen de dataset die de functie in moet flink verkleinen; alle variabelen
      # behalve [var], dummy._col[1:n], dummy.[var], de subsets en de crossings kunnen eruit
      vars = c(var, kolom_opbouw$crossing, kolom_opbouw$subset, colnames(data)[str_starts(colnames(data), "dummy._col")],
               colnames(data)[str_starts(colnames(data), paste0("dummy.", var))], "superstrata", "superweegfactor", "fpc",
               "tbl_dataset", weight.factors)
      vars = unique(vars[!is.na(vars)])
      data.tmp = data %>% select(any_of(vars))
      
      # nu wordt het ingewikkeld: in de monitor VO zijn verschillende weegfactoren nodig per jaar
      # dit betekent dat we per variabele EN per dataset een andere weegfactor kunnen hebben
      # daarvoor kan een combinatieweegfactor gemaakt worden, als er gewerkt wordt met één groot combinatiebestand,
      # of er kan een 'superweegfactor' gemaakt worden, gelijkend aan de superweegfactor die door combinatie hierboven ontstaan is
      # aangezien niet alle GGD'en een overkoepelend bestand hebben is hier gekozen voor de tweede optie
      # gezien de complexiteit wordt deze code lelijk, daar is helaas weinig aan te doen
      
      if (any(str_detect(colnames(varlist), "weegfactor"))) {
        # tijd om te huilen
        # mogelijke opties:
        # "weegfactor" -> override voor die variabele
        # "weegfactor.d[getal]" -> override voor die variabele in dataset [getal]
        # "weegfactor.d_[naam]" -> override voor die variabele in dataset [naam]
        weegfactorvars = str_match(colnames(varlist), "weegfactor(.*)")
        weegfactorvars = matrix(weegfactorvars[!is.na(weegfactorvars[,1]),], ncol=2) # het moet via een matrix met 2 kolommen, omdat R anders een enkele rij omzet naar een vector
        
        for (j in 1:nrow(weegfactorvars)) {
          wfname = weegfactorvars[j,1]
          wfdataset = weegfactorvars[j,2]
          
          if (wfname == "weegfactor" && !is.na(varlist$weegfactor[i])) {
            data.tmp$superweegfactor = data.tmp[[varlist$weegfactor[i]]]
          } else if (!is.na(varlist[[wfname]][i]) && str_detect(wfdataset, "^\\.d(\\d+)$")) { # numeriek, dus d[getal]
            dataset = as.numeric(str_match(wfdataset, "(\\d+)")[,2])
            data.tmp$superweegfactor[data.tmp$tbl_dataset == dataset] = data.tmp[[varlist[[wfname]][i]]][data.tmp$tbl_dataset == dataset]
          } else if (!is.na(varlist[[wfname]][i]) && str_detect(wfdataset, "^\\.d_(.+)")) { # naam van een dataset
            dataset = str_match(wfdataset, "^\\.d_(.+)")[,2]
            if (!dataset %in% datasets$naam_dataset) {
              msg("Er is een weegfactor opgegeven in indeling_rijen voor dataset %s, maar deze dataset is niet bekend. Controleer de configuratie.", dataset, level=ERR)
            }
            dataset = which(datasets$naam_dataset == dataset)
            data.tmp$superweegfactor[data.tmp$tbl_dataset == dataset] = data.tmp[[varlist[[wfname]][i]]][data.tmp$tbl_dataset == dataset]
          } else if (!is.na(varlist[[wfname]][i])) {
            msg("Onbekende weegfactordefinitie opgegeven: %s. Controleer de configuratie.", wfname, level=ERR)
          }
        }
      }

      # aanmaken van design, afhankelijk van of er fpc is.
      if("fpc" %in% colnames(datasets)){
        design = svydesign(ids=~1, strata=~superstrata, weights=~superweegfactor, fpc=~fpc, data=data.tmp)  
      } else {
        design = svydesign(ids=~1, strata=~superstrata, weights=~superweegfactor, data=data.tmp)
      }
      
      
      t.before = proc.time()["elapsed"]
      results = bind_rows(results, GetTableRow(var, design, kolom_opbouw, subsetmatches))
      t.after = proc.time()["elapsed"]
      t.vars = c(t.vars, t.after-t.before)
      
      msg("Variabele %d/%d berekend; rekentijd %0.1f sec.", i, nrow(varlist), t.after-t.before, level=MSG)
    }
    t.end = proc.time()["elapsed"]
    msg("Totale rekentijd %0.2f min voor %d variabelen. Gemiddelde tijd per variabele was %0.1f sec (range %0.1f - %0.1f).",
        (t.end-t.start)/60, nrow(varlist), mean(t.vars), min(t.vars), max(t.vars), level=MSG)
    
    # aangezien een tweede subset vaker kan draaien moeten we hier een distinct op doen
    results = results %>% distinct()
    
    # multiple testing correctie
    if("multiple_testing_correction" %in% colnames(algemeen) && !is.na(algemeen$multiple_testing_correction[1])){
      if("aantal_toetsen" %in% colnames(algemeen) && !is.na(algemeen$aantal_toetsen[1])){
        aantal_toetsen <- algemeen$aantal_toetsen[1]
      } else {
        aantal_toetsen <- length(results$sign[!is.na(results$sign)])
      }
      msg("Correctie voor multiple testing (aantal toetsen %d) toegepast met methode %s", aantal_toetsen, algemeen$multiple_testing_correction[1])
      results$sign <- p.adjust(results$sign, algemeen$multiple_testing_correction[1])
    }
    
    # resultaten opslaan voor hergebruik
    # N.B.: alles in UTF-8 om problemen met een trema o.i.d. te voorkomen
    write.csv(varlist, sprintf("resultaten_csv/vars_%s.csv", basename(config.file)), fileEncoding="UTF-8", row.names=F)
    write.csv(var_labels, sprintf("resultaten_csv/varlabels_%s.csv", basename(config.file)), fileEncoding="UTF-8", row.names=F)
    write.csv(kolom_opbouw, sprintf("resultaten_csv/settings_%s.csv", basename(config.file)), fileEncoding="UTF-8", row.names=F)
    write.csv(results, sprintf("resultaten_csv/results_%s.csv", basename(config.file)), fileEncoding="UTF-8", row.names=F)
    write.csv(fpc_data, sprintf("resultaten_csv/fpc_%s.csv", basename(config.file)), fileEncoding="UTF-8", row.names=F)
  }
  
  ##### begin wegschrijven tabellenboeken in Excel
  basefilename <- opmaak %>% filter(type == "naam_tabellenboek") %>% pull(waarde)
  
  if (!is.na(algemeen$template_html)) {
    # uitdraaien tabellenboeken in HTML-vorm voor digitoegankelijkheid
    source(paste0(dirname(this.path()), "/tbl_MakeHtml.R"))
    
    template_html = read_file(algemeen$template_html)
    
    if (is.null(subsetmatches)) {
      # geen subsets, 1 tabellenboek
      msg("Digitoegankelijk tabellenboek wordt gemaakt.", level=MSG)
      MakeHtml(results, var_labels, kolom_opbouw, NA, NA, subsetmatches, n_resp, template_html, filename=basefilename)
    } else {
      # wel subsets, meerdere tabellenboeken
      subsetvals = subsetmatches[, 1]
      # bij een subset met slechts 1 niveau valt de naam weg bij de bovenstaande selectie
      # om dit te voorkomen voegen we 'm zelf nog een keer toe
      names(subsetvals) = rownames(subsetmatches)
      for (s in 1:length(subsetvals)) {
        if (sum(results$subset == colnames(subsetmatches)[1] & results$subset.val == subsetvals[s], na.rm=T) < 1)
          next # geen data gevonden voor deze subset, overslaan
        
        msg("Digitoegankelijk tabellenboek voor %s wordt gemaakt.", names(subsetvals[s]), level=MSG)
        MakeHtml(results, var_labels, kolom_opbouw, colnames(subsetmatches)[1], subsetvals[s], subsetmatches, n_resp, template_html, filename=paste0(basefilename, " ", names(subsetvals[s])))
      }
    }
  }
  
  # uitdraaien tabellenboeken
  source(paste0(dirname(this.path()), "/tbl_MakeExcel.R"))
  if (is.null(subsetmatches)) {
    # geen subsets, 1 tabellenboek
    msg("Tabellenboek wordt gemaakt.", level=MSG)
    MakeExcel(results, var_labels, kolom_opbouw, NA, NA, subsetmatches, n_resp, filename=basefilename)
  } else {
    # wel subsets, meerdere tabellenboeken
    subsetvals = subsetmatches[, 1]
    names(subsetvals) = rownames(subsetmatches)
    for (s in 1:length(subsetvals)) {
      if (sum(results$subset == colnames(subsetmatches)[1] & results$subset.val == subsetvals[s], na.rm=T) < 1)
        next # geen data gevonden voor deze subset, overslaan
      
      msg("Tabellenboek voor %s wordt gemaakt.", names(subsetvals[s]), level=MSG)
      MakeExcel(results, var_labels, kolom_opbouw, colnames(subsetmatches)[1], subsetvals[s], subsetmatches, n_resp, filename=paste0(basefilename, " ", names(subsetvals[s])))
    }
  }
}
