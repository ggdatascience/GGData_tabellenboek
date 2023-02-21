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
              "labelled", "openxlsx", "doParallel")

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
  options("openxlsx.paperSize" = 9) # A4
  options("openxlsx.orientation" = "landscape")
  options("openxlsx.numFmt" = "0") # standaardformaat zonder decimalen (anders 0.0 invullen)
  
  # controleren of de configuratie leesbaar is
  tmp = read.xlsx(config.file, sheet="datasets")
  if (!exists("tmp")) msg("Configuratiebestand kon niet gelezen worden. Wellicht is deze nog geopend in Excel?", ERR)
  rm(tmp)
  
  # zorgen dat de juiste mappen bestaan
  if (!dir.exists("output")) dir.create("output")
  if (!dir.exists("resultaten_csv")) dir.create("resultaten_csv")
  
  # daadwerkelijk inlezen configuratie
  # TODO: onmogelijke waardes checken
  sheets = c("algemeen", "crossings", "datasets", "indeling_rijen", "onderdelen")
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
  rm(data.combined) # twee keer hetzelfde object is zonde van het geheugen
  
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
      data$superstrata[data$tbl_dataset == d & data$tbl_strata == strata[i]] = i
    }
    data$superstrata[data$tbl_dataset == d] = d*1000 + data$superstrata[data$tbl_dataset == d]
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
          n, level=MSG)
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
        } else if (str_detect(onderdelen$sign_crossing[i], "^\\d+$")) {
          # nummer gegeven; dat moet een kolomindex zijn
          test.col = as.numeric(onderdelen$sign_crossing[i])
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
  # LET OP: daadwerkelijke subsetvariabelen komen pas later aan bod (ivm testen middels chi square)
  kolom_opbouw$global.subset = rep(NA, nrow(kolom_opbouw))
  #kolom_opbouw = kolom_opbouw %>% group_by(dataset, subset, year)
  #subsets = group_keys(kolom_opbouw)
  for (i in 1:nrow(kolom_opbouw)) {
    subset = data$tbl_dataset == kolom_opbouw$dataset[i]
    # scheiden per jaar?
    if (!is.na(kolom_opbouw$year[i])) {
      subset = subset & data$tbl_jaar == kolom_opbouw$year[i]
    }
    # crossing?
    if (!is.na(kolom_opbouw$crossing[i])) {
      subset = subset & data[[kolom_opbouw$crossing[i]]] == kolom_opbouw$crossing.val[i]
    }
    
    data[,paste0("dummy._col", i)] = subset
  }
  
  ##### begin berekeningen
  
  # col.design = kolom_opbouw
  GetTableRow = function (var, design, calculate.ci=T, col.design) {
    msg("Variabele %s wordt uitgevoerd over %d kolommen.", var, nrow(col.design), level=MSG)
    
    results = data.frame()
    
    col.design = col.design %>% group_by(dataset, subset, year, crossing, test.col)
    subsets = group_keys(col.design)
    
    for (i in 1:nrow(subsets)) {
      msg("Subset %d: dataset %d, subset %s, jaar %s, crossing %s", level=DEBUG,
          i, subsets$dataset[i], subsets$subset[i], subsets$year[i], subsets$crossing[i])
      # we loopen op basis van selectiecriteria ipv kolomnummer, omdat er anders meerdere svychisq calls nodig zouden zijn voor crossings
      if (!is.na(subsets$subset[i])) {
        # splitsen op variabele
        
        # TODO: schrijven
      } else {
        # niet splitsen op variabele
        
        # crossing? zo ja, dan kunnen we dit behandelen als totaal indien er NIET intern vergeleken wordt (binnen de else)
        # anders wel behandelen als crossing (dus binnen de if)
        if (!is.na(subsets$crossing[i]) && !is.na(subsets$test.col[i]) && subsets$test.col[i] == 0) {
          # dit betekent meerdere kolommen vullen
          cols = col.design$col.index[group_rows(col.design)[[i]]]
          selection = str_c(paste0("dummy._col", cols), collapse=" | ")
          design.subset = subset(design, eval(parse(text=selection)))

          weighted = svytable(formula=as.formula(paste0("~", var, "+", subsets$crossing[i])),
                              design=design.subset)
          unweighted = table(design.subset$variables[[var]], design.subset$variables[[subsets$crossing[i]]])
          n = length(weighted)
          
          pvals = matrix(NA, nrow=nrow(weighted), ncol=ncol(weighted))
          rownames(pvals) = rownames(weighted)
          
          answers = rownames(weighted)
          for (answer in answers) {
            test = svychisq(formula=as.formula(paste0("~dummy.", var, ".", answer, "+", col.design$crossing[i])), design=design.subset)
            pvals[answer,] = rep(test$p.value, ncol(pvals))
          }
          
          # om de resultaten in een dataframe te krijgen moeten ze door as.numeric() en unname()
          # dit zorgt ervoor dat de structuur verloren gaat; de kolommen worden onder elkaar geplaatst
          # we moeten dus de waardes indelen als colnames[1] * nrow, colnames[2] * nrow, etc.
          vals = as.vector(sapply(colnames(weighted), function (x, nrows) return(rep(x, nrows)), nrows=nrow(weighted)))
          
          results = bind_rows(results, data.frame(dataset=rep(subsets$dataset[i], n), subset=rep(subsets$subset[i], n), subset.val=rep(NA, n),
                                                  year=rep(subsets$year[i], n),
                                                  crossing=rep(subsets$crossing[i], n), crossing.val=vals,
                                                  sign.vs=rep(subsets$test.col[i], n), sign=as.numeric(unname(pvals)),
                                                  var=rep(var, n), val=rownames(weighted),
                                                  n.weighted=as.numeric(unname(weighted)), perc.weighted=as.numeric(unname(proportions(weighted, margin=2)))*100,
                                                  n.unweighted=as.numeric(unname(unweighted)), perc.unweighted=as.numeric(unname(proportions(unweighted, margin=2)))*100))
        } else {
          # geen crossing; totaal
          
          col = col.design$col.index[group_rows(col.design)[[i]]]
          
          weighted = svytable(formula=as.formula(paste0("~", var, "+dummy._col", col)), design=design)
          if (ncol(weighted) == 2 && colnames(weighted)[2] == "TRUE") {
            weighted = weighted[,2]
          } else {
            names = rownames(weighted)
            weighted = rep(NA, nrow(weighted))
            names(weighted) = names
          }
          unweighted = as.numeric(unname(table(design$variables[[var]][design$variables[,paste0("dummy._col", col)]])))
          if (sum(unweighted, na.rm=T) <= 0) unweighted = rep(NA, length(weighted))
          n = length(weighted)
          
          pvals = rep(NA, n)
          names(pvals) = names(weighted)
          
          # significantie berekenen?
          if (!is.na(subsets$test.col[i]) && sum(unweighted, na.rm=T) > 0) {
            # tweede kolom includeren
            selection = str_c(paste0("dummy._col", c(col, subsets$test.col[i])), collapse=" | ")
            design.subset = subset(design, eval(parse(text=selection)))
            
            answers = names(weighted)
            for (answer in answers) {
              test = svychisq(formula=as.formula(paste0("~dummy.", var, ".", answer, "+dummy._col", col)),
                              design=design.subset)
              pvals[answer] = test$p.value
            }
          }
          
          results = bind_rows(results, data.frame(dataset=rep(col.design$dataset[col], n), subset=rep(col.design$subset[col], n), subset.val=rep(NA, n),
                                                  year=rep(col.design$year[col], n),
                                                  crossing=rep(col.design$crossing[col], n), crossing.val=as.character(rep(col.design$crossing.val[col], n)),
                                                  sign.vs=rep(col.design$test.col[col], n), sign=pvals,
                                                  var=rep(var, n), val=names(weighted),
                                                  n.weighted=as.numeric(unname(weighted)), perc.weighted=proportions(as.numeric(unname(weighted)))*100,
                                                  n.unweighted=unweighted, perc.unweighted=proportions(unweighted)*100))
        }
      }
    }
    
    
    return(results)
  }
  
  results = data.frame()
  for (var in varlist) {
    # we kunnen de dataset die de functie in moet flink verkleinen; alle variabelen
    # behalve [var], dummy._col[1:n], dummy.[var] en de crossings kunnen eruit
    vars = c(var, unique(kolom_opbouw$crossing), unique(kolom_opbouw$subset))
    vars = vars[!is.na(vars)]
    data.tmp = data %>% select(all_of(vars), starts_with("dummy._col"), starts_with(paste0("dummy.", var)),
                               superstrata, superweegfactor)
    design = svydesign(ids=~1, strata=data.tmp$superstrata, weights=data.tmp$superweegfactor, data=data.tmp)
    
    results = bind_rows(results, GetTableRow(var, design, F, kolom_opbouw))
  }
  
  # berekeningen kunnen parallel worden uitgevoerd met multithreading
  # hierdoor worden meerdere processorkernen ingezet om de berekeningen uit te voeren
  # (dit kan handig zijn, omdat survey een bijzonder traag pakket is)
}