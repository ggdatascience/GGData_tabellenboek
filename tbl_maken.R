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
              "labelled", "openxlsx", "doParallel", "foreach")

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
suppressPackageStartupMessages(library(doParallel))
library(foreach)

# instellen werkmap voor het laden van de andere bestanden
setwd(dirname(this.path()))
source("tbl_helpers.R")

# het gehele script wordt omgeven door curly brackets, zodat een stop() ook daadwerkelijk het script stopt
{
  # selecteren configuratiebestand en bijbehorende werkmap
  # hierin dienen de configuratie(.xlsx) en de databestanden (in de map data) te staan
  # andere mappen worden automatisch aangemaakt als deze niet bestaan
  #config.file = choose.files(caption="Selecteer configuratiebestand...",
  #                           filters=c("Excel-bestand (*.xlsx)","*.xlsx"),
  #                           multi=F)
  config.file = "configuratie.xlsx"
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
  sheets = c("algemeen", "crossings", "datasets", "indeling_rijen", "onderdelen", "opmaak", "labelcorrectie", "logos", "intro_tekst", "headers_afkortingen")
  for (sheet in sheets) {
    tmp = read.xlsx(config.file, sheet=sheet)
    if (ncol(tmp) == 1) tmp = tmp[[1]]
    assign(sheet, tmp, envir=.GlobalEnv)
  }
  rm(tmp)
  
  # variabelelijst afleiden uit de indeling van het tabellenboek;
  # iedere regel met (n)var is een variabele die we nodig hebben
  varlist = indeling_rijen$inhoud[indeling_rijen$type %in% c("var", "nvar")]
  if (any(duplicated(varlist))) {
    msg("Let op: variabele(n) %s komen meerdere keren voor in indeling_rijen. Controleer of dit de bedoeling is.",
        str_c(varlist[duplicated(varlist)], collapse=", "), level=WARN)
  }
  if (any(is.na(varlist))) {
    msg("Let op: regel(s) %s bevat geen variabele. Controleer de configuratie.", level=ERR)
  }
  varlist = unique(varlist[!is.na(varlist)])
  
  if (any(crossings %in% onderdelen$subset)) {
    msg("De variabele(n) %s is/zijn ingevuld als crossing en subset. Een subset kan niet met zichzelf gekruist worden.",
        str_c(crossings[crossings %in% onderdelen$subset], ", "), level=WARN)
  }
  
  # datasets combineren en de strata en weegfactoren apart opslaan
  data.combined = data.frame()
  for (d in 1:nrow(datasets)) {
    msg("Inladen dataset %d: %s", d, datasets$naam_dataset[d], level=MSG)
    
    # .sav aan het einde kan vergeten worden...
    if (!str_ends(datasets$bestandsnaam[d], fixed(".sav"))) {
      datasets$bestandsnaam[d] = paste0(datasets$bestandsnaam[d], ".sav")
      msg("De opgegeven dataset op rij %d eindigt niet in .sav. Mogelijk vergeten? .sav toegevoegd.", level=WARN, d+1)
    }
    
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
      
      # gewichten apart opslaan
      if (!is.na(datasets$weegfactor[d]) && colname == datasets$weegfactor[d]) {
        colnames(data)[c] = "tbl_weegfactor"
        next
      }
      
      # jaren apart opslaan
      if (!is.na(datasets$jaarvariabele[d]) && colname == datasets$jaarvariabele[d]) {
        colnames(data)[c] = "tbl_jaar"
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
    msg("Afwijkende kolommen in dataset %d: %s", d, str_c(colnames(data)[afwijkend], collapse=","))
    if (length(afwijkend) > 0) colnames(data)[afwijkend] = paste0("_", d, "_", colnames(data)[afwijkend])
    
    # zijn er weegfactoren aangetroffen? zo nee, is dat de bedoeling?
    if (!("tbl_weegfactor" %in% colnames(data))) {
      if (is.na(datasets$weegfactor[d]) && is.na(datasets$stratum[d])) {
        # geen stratum en weegfactor opgegeven; alles invullen met 1
        data$tbl_weegfactor = rep(1, nrow(data))
        data$tbl_strata = rep(1, nrow(data))
      } else {
        msg("Kolom %s of %s is niet aangetroffen in dataset %d (%s). Laat voor een ongewogen design beide velden leeg of pas de dataset aan.",
            datasets$weegfactor[d], datasets$stratum[d], d, datasets$naam_dataset[d], level=ERR)
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
    }
    labels = val_labels(data[[i]])
    if (!is.null(labels)) {
      var_labels = bind_rows(var_labels, data.frame(var=rep(colname, length(labels)), val=as.character(unname(labels)), label=names(labels)))
    }
  }
  
  # dummies maken
  for (c in 1:ncol(data)) {
    colname = colnames(data)[c]
    if (!(colname %in% varlist) && !(colname %in% crossings)) next
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
      data$superstrata[data$tbl_dataset == d & data$tbl_strata == strata[i]] = d*1000 + i
    }
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
          test.col = 0 # 0 betekent andere waarden van de crossing
        } else {
          # geen toets
          test.col = NA
        }
      } else {
        # geen toets
        test.col = NA
      }
      for (crossing in crossings) {
        crossing.labels = val_labels(data[[crossing]])
        n = length(crossing.labels)
        kolom_opbouw = bind_rows(kolom_opbouw, data.frame(col.index=nrow(kolom_opbouw)+(1:n), dataset=rep(d, n),
                                                          subset=rep(onderdelen$subset[i], n), year=rep(onderdelen$jaar[i], n),
                                                          crossing=rep(crossing, n), crossing.val=unname(crossing.labels),
                                                          crossing.lab=names(crossing.labels), test.col=rep(test.col, n)))
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
      }
    } else {
      test.col = NA
    }
    
    kolom_opbouw = bind_rows(kolom_opbouw, data.frame(col.index=nrow(kolom_opbouw)+1, dataset=d, subset=onderdelen$subset[i], year=onderdelen$jaar[i],
                                                      crossing=NA, crossing.val=NA, crossing.lab=NA, test.col=test.col))
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
  #kolom_opbouw = kolom_opbouw %>% group_by(dataset, subset, year)
  #subsets = group_keys(kolom_opbouw)
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
    
    # ook splitsen per subset?
    if (!is.na(kolom_opbouw$subset[i])) {
      subsetvals = val_labels(data[[kolom_opbouw$subset[i]]])
      
      for (val in subsetvals) {
        subset = kolom & data[[kolom_opbouw$subset[i]]] == val
        if (sum(subset, na.rm=T) <= 0) next # als er geen data bestaat voor die subset is het een beetje nutteloos
        data[,paste0("dummy._col", i, ".s.", unname(val))] = subset
      }
    }
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
    varlist.prev = read.csv(sprintf("resultaten_csv/vars_%s.csv", basename(config.file)), fileEncoding="UTF-8")[,1]
    
    kolom_opbouw.identical = T
    for (i in 1:ncol(kolom_opbouw)) {
      # N.B.: je zou verwachten dat identical(x, y) hier zou werken, maar dat blijkt niet zo te zijn
      # bij numerieke waarden kan identical() raar doen, bijv. 1,2,3 != 1,2,3 (volgens identical() dan)
      # andersom doet all.equal het niet goed met kolommen met alleen NA
      # vandaar dit gedrocht
      if (!isTRUE(all.equal(unname(kolom_opbouw[,i]), unname(kolom_opbouw.prev[,i])))) {
        if (!(all(is.na(unname(kolom_opbouw[,i]))) && all(is.na(unname(kolom_opbouw.prev[,i]))))) {
          msg("Kolom %d (%s) is ongelijk aan de vorige configuratie: %s vs. %s", i, colnames(kolom_opbouw)[i],
              str_c(unname(kolom_opbouw[,i]), collapse=", "), str_c(unname(kolom_opbouw.prev[,i]), collapse=", "), level=DEBUG)
          kolom_opbouw.identical = F
        }
      }
    }
    
    if (kolom_opbouw.identical && identical(varlist, varlist.prev)) {
      msg("Eerdere resultaten aangetroffen vanuit deze configuratie (%s). Berekening wordt overgeslagen. Indien er nieuwe data is toegevoegd, verwijder dan de bestanden uit de map resultaten_csv.",
          basename(config.file), level=MSG)
      
      calc.results = F
    } else {
      msg("Er zijn eerdere resultaten aangetroffen vanuit deze configuratie (%s), maar de instellingen waren niet identiek. Berekening wordt opnieuw uitgevoerd.", basename(config.file), level=MSG)
    }
  }
  
  if (calc.results) {
    source("tbl_GetTableRow.R")
    results = data.frame()
    
    # TODO: parallel proberen te maken
    if (F) {
      # berekeningen kunnen parallel worden uitgevoerd met multithreading
      # hierdoor worden meerdere processorkernen ingezet om de berekeningen uit te voeren
      # (dit kan handig zijn, omdat survey een bijzonder traag pakket is)
      
      clus = makeCluster(detectCores())
      registerDoParallel(clus)
      
      
      #### poging met plyr
      results = data.frame()
      
      PrepdataForSurveyAndGetTableRow = function(
      var,
      data.in,
      kolom_opbouw.in,
      subsetmatches.in
      ){
        source("tbl_GetTableRow.R")
        source("tbl_helpers.R")
        options(survey.lonely.psu="certainty")
        vars = c(var, kolom_opbouw.in$crossing, kolom_opbouw.in$subset, colnames(data.in)[str_starts(colnames(data.in), "dummy._col")],
                 colnames(data.in)[str_starts(colnames(data.in), paste0("dummy.", var))], "superstrata", "superweegfactor")
        vars = unique(vars[!is.na(vars)])
        data.in.tmp = data.in[,vars]
        design = svydesign(ids=~1, strata=data.in.tmp$superstrata, weights=data.in.tmp$superweegfactor, data=data.in.tmp)
        return(GetTableRow(var, design, kolom_opbouw.in, subsetmatches.in, F))
      }
      
      out <- alply(
        .data = varlist,
        .margins = 1,
        .fun = PrepdataForSurveyAndGetTableRow,
        data.in = data,
        kolom_opbouw.in = kolom_opbouw,
        subsetmatches.in = subsetmatches,
        .parallel = algemeen$multithreading,
        .paropts = list(
          .packages = c("survey", "tidyverse", "labelled")
        ),
        .progress = progress_time()
      )
      
      results <- out %>% reduce(.f = bind_rows)
      
      stopCluster(clus)
      
      
      #### poging met foreach
      results = system.time({foreach(var=varlist, .combine=bind_rows, .packages=c("tidyverse", "survey", "labelled")) %dopar% {
        # we kunnen de dataset die de functie in moet flink verkleinen; alle variabelen
        # behalve [var], dummy._col[1:n], dummy.[var], de subsets en de crossings kunnen eruit
        vars = c(var, kolom_opbouw$crossing, kolom_opbouw$subset, colnames(data)[str_starts(colnames(data), "dummy._col")],
                 colnames(data)[str_starts(colnames(data), paste0("dummy.", var))], "superstrata", "superweegfactor")
        vars = unique(vars[!is.na(vars)])
        data.tmp = data[,vars]
        design = svydesign(ids=~1, strata=data.tmp$superstrata, weights=data.tmp$superweegfactor, data=data.tmp)
        
        GetTableRow(var, design, kolom_opbouw, subsetmatches, F)
      }})
    }
    
    results = data.frame()
    t.start = proc.time()["elapsed"]
    t.vars = c()
    for (var in varlist) {
      # we kunnen de dataset die de functie in moet flink verkleinen; alle variabelen
      # behalve [var], dummy._col[1:n], dummy.[var], de subsets en de crossings kunnen eruit
      vars = c(var, kolom_opbouw$crossing, kolom_opbouw$subset, colnames(data)[str_starts(colnames(data), "dummy._col")],
               colnames(data)[str_starts(colnames(data), paste0("dummy.", var))], "superstrata", "superweegfactor")
      vars = unique(vars[!is.na(vars)])
      data.tmp = data[,vars]
      if ("weegfactor" %in% colnames(indeling_rijen) && !is.na(indeling_rijen$weegfactor[indeling_rijen$inhoud == var])) {
        design = svydesign(ids=~1, strata=data.tmp$superstrata, weights=indeling_rijen$weegfactor[indeling_rijen$inhoud == var], data=data.tmp)
      } else {
        design = svydesign(ids=~1, strata=data.tmp$superstrata, weights=data.tmp$superweegfactor, data=data.tmp)
      }
      
      t.before = proc.time()["elapsed"]
      results = bind_rows(results, GetTableRow(var, design, kolom_opbouw, subsetmatches, F))
      t.after = proc.time()["elapsed"]
      t.vars = c(t.vars, t.after-t.before)
      
      msg("Variabele %d/%d berekend; rekentijd %0.1f sec.", which(varlist == var), length(varlist), t.after-t.before, level=MSG)
    }
    t.end = proc.time()["elapsed"]
    msg("Totale rekentijd %0.2f min voor %d variabelen. Gemiddelde tijd per variabele was %0.1f sec (range %0.1f - %0.1f).",
        t.end-t.start, length(varlist), mean(t.vars), min(t.vars), max(t.vars), level=MSG)
    
    # resultaten opslaan voor hergebruik
    # N.B.: alles in UTF-8 om problemen met een trema o.i.d. te voorkomen
    write.csv(varlist, sprintf("resultaten_csv/vars_%s.csv", basename(config.file)), fileEncoding="UTF-8", row.names=F)
    write.csv(kolom_opbouw, sprintf("resultaten_csv/settings_%s.csv", basename(config.file)), fileEncoding="UTF-8", row.names=F)
    write.csv(results, sprintf("resultaten_csv/results_%s.csv", basename(config.file)), fileEncoding="UTF-8", row.names=F)
  }
  
  ##### begin wegschrijven tabellenboeken in Excel
  
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
  
  # uitdraaien tabellenboeken
  source("tbl_MakeExcel.R")
  if (is.null(subsetmatches)) {
    # geen subsets, 1 tabellenboek
    msg("Tabellenboek wordt gemaakt: Overzicht.xlsx.", level=MSG)
    MakeExcel(results, var_labels, kolom_opbouw, NA, NA, subsetmatches)
  } else {
    # wel subsets, meerdere tabellenboeken
    subsetvals = subsetmatches[,1]
    for (s in 1:length(subsetvals)) {
      if (sum(results$subset == colnames(subsetmatches)[1] & results$subset.val == subsetvals[s], na.rm=T) < 1)
        next # geen data gevonden voor deze subset, overslaan
      
      msg("Tabellenboek voor %s wordt gemaakt.", names(subsetvals[s]), level=MSG)
      MakeExcel(results, var_labels, kolom_opbouw, colnames(subsetmatches)[1], subsetvals[s], subsetmatches,
                min_observaties_per_vraag = algemeen$min_observaties_per_vraag,
                min_observaties_per_antwoord = algemeen$min_observaties_per_antwoord,
                tekst_min_vraag_niet_gehaald = algemeen$tekst_min_vraag_niet_gehaald,
                tekst_min_antwoord_niet_gehaald = algemeen$tekst_min_antwoord_niet_gehaald)
    }
  }
}
