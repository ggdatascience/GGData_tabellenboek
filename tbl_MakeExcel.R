#
#
# Deze functie schrijft de resultatentabel vanuit survey om naar een bruikbare vorm.
# Opmaak wordt ingelezen vanuit de configuratie (tabblad opmaak) en verdere instellingen
# van de tabbladen indeling_rijen, indeling_kolommen en algemeen.
#
#

# functie om vergelijken met NA mogelijk te maken
# LET OP: de naam is ietwat verwarrend; de werking is juist NIET gelijk aan identical()
# identical() geeft één T/F, NA.identical geeft een lijst, zoals is.na() (zodat dit gecombineerd kan worden met andere selecties)
NA.identical = function (vect, cmp) {
  if (is.na(cmp)) {
    return(is.na(vect))
  } else {
    return(vect == cmp)
  }
}

# standaardwaardes opmaak, mochten deze missen
opmaak.default = read.table(text='"type" "waarde"
"1" "titel_size" "14"
"2" "titel_color" "#FFFFFF"
"3" "titel_decoration" "bold"
"4" "titel_fill" "#D1005D"
"5" "kop_size" "14"
"6" "kop_color" "#FFFFFF"
"7" "kop_decoration" "bold"
"8" "kop_fill" "#1D1756"
"9" "rij_hoogte" "16"
"10" "rij_hoogte_kop" "28"
"11" "kolombreedte_antwoorden" "60"
"12" "kolombreedte" "10"
"13" "font_color" "#000000"
"14" "font_type" "Calibri"
"15" "font_size" "11"
"16" "border_tussen_gegevens" "FALSE"
"17" "rijen_afwisselend_kleuren" "FALSE"
"18" "kolommen_afwisselend_kleuren" "FALSE"
"19" "kolommen_crossings_kleuren" "TRUE"
"20" "header_stijl" "enkel"
"21" "header_template" "Totaal [naam] [jaar]"
"22" "crossing_headers_kleiner" "TRUE"
"23" "label_max_lengte" "66"
"24" "naam_tabellenboek" "Overzicht"
"25" "verberg_lege_kolommen_totaal" "FALSE"
"26" "verberg_lege_kolommen_crossing" "FALSE"')

design = function (var) {
  if (str_length(var) <= 1) {
    msg("Variabelenaam kan niet leeg zijn.", level=ERR)
  }
  
  if (var %in% opmaak$type) {
    ret = opmaak$waarde[opmaak$type == var]
  } else if (str_starts(var, "vraag_") && str_replace(var, "vraag_", "kop_") %in% opmaak$type) {
    # losse opmaak voor vragen is later toegevoegd, wanneer deze ontbreekt nemen we de opmaak van een kop
    ret = opmaak$waarde[opmaak$type == str_replace(var, "vraag_", "kop_")]
  } else if (var %in% opmaak.default$type) {
    ret = opmaak.default$waarde[opmaak.default$type == var]
  } else {
    msg("Opmaakvariabele niet gevonden: %s", var, level=ERR)
  }
  
  # waarden omzetten waar nodig
  if (str_detect(var, "decoration")) {
    # textDecoration moet NULL zijn indien niet gewenst; waardes opschonen
    if (str_length(ret) <= 1) ret = NULL
  } else if (ret == "TRUE" || ret == "FALSE") {
    ret = ifelse(ret == "TRUE", T, F)
  } else if (str_detect(var, "size") || str_detect(var, "hoogte") || str_detect(var, "breedte") || str_detect(var, "lengte")) {
    ret = as.numeric(ret)
  }
  
  return(ret)
}

# col.design = kolom_opbouw
# subset = "Gemeentecode"
# subset.val = 197
# subset.name = "Aalten"
# i = 7

MakeExcel = function (results, var_labels, col.design, subset, subset.val, subsetmatches, n_resp, filename) {
  subset.name = names(subset.val)
  subset.val = unname(subset.val)
  
  if (is.na(subset.name) || is.null(subset.name))
    subset.name = design("naam_tabellenboek")
  
  # Excel heeft een maximum voor de naam van een tabblad; afkorten?
  subset.name.full = subset.name
  if (str_length(subset.name) > 31) {
    subset.name = paste0(str_sub(subset.name, end=27), "...")
  }
  
  # weergave significantie in andere kolommen?
  sign.replacement = col.design[!col.design$test.display, c("col.index", "test.col")]
  colnames(sign.replacement) = c("old", "new")
  
  # moeten er kolommen verborgen worden?
  # let op: aangezien de kolommen hier genummerd worden i.p.v. herkend aan de naam, moet een verwijderde kolom verrekend worden in sign.replacement
  for (j in 1:nrow(col.design)) {
    if (!is.na(col.design$subset[j])) {
      subset.col = subset.val
      if (col.design$subset[j] != subset) {
        subset.col = subsetmatches[subsetmatches[,1] == subset.val, col.design$subset[j]]
      }
      
      n = n_resp$n[which(n_resp$col == j & NA.identical(n_resp$year, col.design$year[j]) & NA.identical(n_resp$crossing, col.design$crossing[j]) & NA.identical(n_resp$subset, subset.col))]
    } else {
      n = n_resp$n[which(n_resp$col == j & NA.identical(n_resp$year, col.design$year[j]) & NA.identical(n_resp$crossing, col.design$crossing[j]) & is.na(n_resp$subset))]
    }
    
    # het is mogelijk dat er helemaal geen deelnemers zijn; dan willen we dat aangeven
    if (length(n) == 0 || is.na(n) || n <= 0) n = 0
    col.design$n[j] = n
  } 
  hide.cols = c()
  if (design("verberg_lege_kolommen_totaal")) {
    hide.cols = c(hide.cols, col.design$col.index[col.design$n == 0 & is.na(col.design$crossing)])
  }
  if (design("verberg_lege_kolommen_crossing")) {
    hide.cols = c(hide.cols, col.design$col.index[col.design$n == 0 & !is.na(col.design$crossing)])
  }
  if (length(hide.cols) > 0) {
    # achterstevoren, zodat we steeds makkelijk -1 kunnen doen
    sign.replacement = sign.replacement[order(sign.replacement$old, decreasing=T),]
    for (i in hide.cols) {
      sign.replacement$old[sign.replacement$old >= i] = sign.replacement$old[sign.replacement$old >= i] - 1
      sign.replacement$new[sign.replacement$new >= i] = sign.replacement$new[sign.replacement$new >= i] - 1
    }
    col.design = col.design[-which(col.design$col.index %in% hide.cols),]
  }
  
  wb = createWorkbook(creator="GGData Tabellenboek")
  addWorksheet(wb, subset.name)
  
  # opmaak
  modifyBaseFont(wb, fontSize=design("font_size"), fontColour=design("font_color"), fontName=design("font_type"))
  
  style.sign = createStyle(textDecoration="bold") # significante resultaten
  style.title = createStyle(fontSize=design("titel_size"), fontColour=design("titel_color"),
                            textDecoration=design("titel_decoration"), fgFill=design("titel_fill")) # titels
  style.subtitle = createStyle(fontSize=design("kop_size"), fontColour=design("kop_color"),
                               textDecoration=design("kop_decoration"), fgFill=design("kop_fill")) # koppen
  style.question = createStyle(fontSize=design("vraag_size"), fontColour=design("vraag_color"),
                               textDecoration=design("vraag_decoration"), fgFill=design("vraag_fill")) # vragen/variabelen/indicatoren
  style.text = createStyle(wrapText=T, valign="center")
  style.header.col = createStyle(halign="center", valign="center", textDecoration="bold") # kolomkoppen
  style.header.col.crossing = createStyle(halign="center", valign="center") # kolomkoppen crossings
  style.perc = createStyle(halign="center", valign="center") # percentagetekens
  style.num = createStyle(numFmt="0", halign="center", valign="center") # cijfers -> 0 betekent hele getallen zonder decimalen, 0.0 -> 1 decimaal, etc.
  style.gray.bg = createStyle(fgFill = "#F2F2F2") # voor afwisselende kolommen/rijen
  style.intro.text = createStyle(wrapText=F)
  style.intro.header = createStyle(wrapText=F, fontSize=design("kop_size"), textDecoration=design("kop_decoration"))
  style.intro.title = createStyle(wrapText=F, fontSize=design("titel_size") + 4, textDecoration=design("titel_decoration"))
  
  # instellingen
  header.col.nrows = ifelse(design("header_stijl") == "dubbel", 2, 1) # aantal rijen per kolomheader
  
  # opslag voor rijen die later aangemaakt of opgemaakt moeten worden
  header.col.rows = numeric(0)
  data.rows = numeric(0)
  title.rows = numeric(0)
  perc.rows = numeric(0)
  label.oversized.rows = numeric(0)
  
  # constantes voor het vervangen van cellen met te weinig antwoorden
  # dit is nodig omdat bij het direct plaatsen van tekst de hele matrix een karakter wordt
  # dan werkt de weergave in procenten in Excel niet goed meer
  A_TOOSMALL = -1
  Q_TOOSMALL = -2
  Q_MISSING = -3
  
  c = 1 # teller voor rijen in Excel
  
  # moet er introtekst bij?
  if (nrow(intro_tekst) > 0) {
    intro_tekst$type = str_to_lower(str_trim(intro_tekst$type))
    # in de digitoegankelijke versie is het soms wenselijk om het logo te plaatsen in de tekst, maar binnen Excel heeft dit geen functie; verwijderen
    intro_tekst$inhoud = str_replace_all(intro_tekst$inhoud, fixed("{logo}"), "")
    intro_tekst$inhoud = str_replace_all(intro_tekst$inhoud, "\\{logo (.*?)\\}", "")
    intro_tekst$inhoud = str_replace_all(intro_tekst$inhoud, fixed("[naam]"), subset.name.full)
    # indien gesplistst per subset: andere subsets ook vervangen?
    if (subset.name.full != design("naam_tabellenboek") && ncol(subsetmatches) > 1) {
      for (i in 1:ncol(subsetmatches)) {
        intro_tekst$inhoud = str_replace_all(intro_tekst$inhoud, fixed(paste0("[", colnames(subsetmatches)[i], "]")), var_labels$label[var_labels$var == colnames(subsetmatches)[i] & var_labels$val == subsetmatches[subset.name.full, i]])
      }
    }
    for (i in 1:nrow(intro_tekst)) {
      output = data.frame(a=intro_tekst$type[i], b=intro_tekst$inhoud[i])
      writeData(wb, subset.name, output, startCol=1, startRow=c, colNames=F)
      
      if (intro_tekst$type[i] == "titel") {
        addStyle(wb, subset.name, style.intro.title, rows=c, cols=1:2, stack=T)
        setRowHeights(wb, subset.name, c, design("titel_size") + 6)
      } else if (intro_tekst$type[i] == "kop") {
        addStyle(wb, subset.name, style.intro.header, rows=c, cols=1:2, stack=T)
        setRowHeights(wb, subset.name, c, design("kop_size") + 2)
      } else {
        addStyle(wb, subset.name, style.intro.text, rows=c, cols=1:2, stack=T)
        setRowHeights(wb, subset.name, c, design("rij_hoogte"))
      }
      
      c = c + 1
    }
    
    # extra witregel voor de volgende sectie, tenzij de laatste regel een witregel is
    if (!is.na(intro_tekst$inhoud[nrow(intro_tekst)])) c = c + 1
  }
  
  indeling_rijen$type = str_to_lower(str_trim(indeling_rijen$type))
  n.col = nrow(col.design) # aantal kolommen in de data
  n.col.total = n.col + 2 # aantal kolommen in de sheet (1: matchcode, 2: label)
  table.start = c
  for (i in 1:nrow(indeling_rijen)) {
    if (indeling_rijen$type[i] %in% c("titel", "kop", "vraag", "tekst")) { # regel met een titel, kop, of tekst
      # een extra witregel voor koppen of titels
      if (indeling_rijen$type[i] != "tekst") {
        if (i > 1)
          c = c + 1
      }
      
      output = data.frame(a=indeling_rijen$type[i], b=str_replace_all(indeling_rijen$inhoud[i], "naam_onderdeel", subset.name.full))
      writeData(wb, subset.name, output, startCol=1, startRow=c, colNames=F)
      
      # TODO: meer opmaak?
      if (indeling_rijen$type[i] == "titel") {
        addStyle(wb, subset.name, style.title, rows=c, cols=1:n.col.total, stack=T)
        title.rows = c(title.rows, c)
      } else if (indeling_rijen$type[i] == "kop") {
        # TODO: willen we hier weer headers toevoegen als de volgende regel een var is?
        addStyle(wb, subset.name, style.subtitle, rows=c, cols=1:n.col.total, stack=T)
        title.rows = c(title.rows, c)
      } else if (indeling_rijen$type[i] == "vraag") {
        addStyle(wb, subset.name, style.question, rows=c, cols=1:n.col.total, stack=T)
        title.rows = c(title.rows, c)
        # kolomkoppen toevoegen?
        if (is.na(indeling_rijen$kolomkoppen[i])) {
          header.col.rows = c(header.col.rows, c+1)
          c = c + header.col.nrows
          
          # we moeten hier c+1 doen, omdat c nog niet bijgewerkt is na het toevoegen van de vraag
          writeData(wb, subset.name, t(rep("%", n.col)), startCol=3, startRow=c+1, colNames=F)
          addStyle(wb, subset.name, style.perc, cols=3:n.col.total, rows=c+1, gridExpand=T, stack=T)
          perc.rows = c(perc.rows, c+1)
          c = c + 1
        }
      } else { # tekst
        addStyle(wb, subset.name, style.text, rows=c, cols=1:n.col.total, stack=T)
        mergeCells(wb, subset.name, cols=2:n.col.total, rows=c)  
      }
      
      c = c + 1 #ifelse(indeling_rijen$type[i] == "tekst", 1, 2) # een extra witregel na een kop of titel
    } else if (indeling_rijen$type[i] == "aantallen") { # regel met aantal deelnemers toevoegen
      output = matrix(data=col.design$n, nrow=1, ncol=nrow(col.design))
      output[,output == 0] = Q_MISSING

      # ruimte vrijhouden voor het later invoegen van headers
      header.col.rows = c(header.col.rows, c)
      c = c + header.col.nrows
      
      writeData(wb, subset.name, "aantallen", startCol=1, startRow=c, colNames=F)
      writeData(wb, subset.name, output, startCol=3, startRow=c, colNames=F)
      
      # missende waardes goed weergeven
      missing = which(output == Q_MISSING | output == Q_TOOSMALL | output == A_TOOSMALL, arr.ind=T)
      if (length(missing) > 0) {
        # vanwege gekke R logica mag de waarde geen naam hebben (unname) en moet het gedwongen een matrix zijn
        replacement = data.frame(row=missing[,1], col=missing[,2], val=matrix(unname(output[missing])))
        
        for (i in 1:nrow(replacement)) {
          val = algemeen$tekst_missende_data
          writeData(wb, subset.name, val, startCol=replacement$col[i]+2, startRow=c+replacement$row[i]-1, colNames=F)
        }
      }
      
      addStyle(wb, subset.name, style.num, rows=c, cols=3:n.col.total, stack=T)
      data.rows = c(data.rows, c)
      
      c = c + 1
    } else if (indeling_rijen$type[i] == "var" || indeling_rijen$type[i] == "nvar") { # variabele toevoegen
      if (!indeling_rijen$inhoud[i] %in% colnames(data)) {
        msg("Variabele %s komt niet voor in de resultaten. Deze wordt overgeslagen. Controleer de configuratie.", indeling_rijen$inhoud[i], level=WARN)
        next
      }
      
      # benodigde resultaten ophalen, zodat we niet steeds belachelijke selectors nodig hebben
      data.var = data.frame()
      for (j in 1:nrow(col.design)) {
        if (!is.na(col.design$subset[j])) {
          subset.col = subset.val
          if (col.design$subset[j] != subset) {
            subset.col = subsetmatches[subsetmatches[,1] == subset.val, col.design$subset[j]]
          }
          data.tmp = results[which(NA.identical(results$dataset, col.design$dataset[j]) & NA.identical(results$subset, col.design$subset[j]) &
                                     NA.identical(results$subset.val, subset.col) &
                                     NA.identical(results$year, col.design$year[j]) & NA.identical(results$crossing, col.design$crossing[j]) &
                                     NA.identical(results$crossing.val, col.design$crossing.val[j]) & NA.identical(results$sign.vs, col.design$test.col[j]) &
                                     results$var == indeling_rijen$inhoud[i]),
                             c("val", "crossing", "crossing.val", "sign", "sign.vs", "n.unweighted", "perc.weighted")]
          if (nrow(data.tmp) == 0) next
          data.tmp$col.index = j
          data.var = bind_rows(data.var, data.tmp)
        }
        else {
          data.tmp = results[which(NA.identical(results$dataset, col.design$dataset[j]) & NA.identical(results$subset, col.design$subset[j]) &
                                     is.na(results$subset.val) &
                                     NA.identical(results$year, col.design$year[j]) & NA.identical(results$crossing, col.design$crossing[j]) &
                                     NA.identical(results$crossing.val, col.design$crossing.val[j]) & NA.identical(results$sign.vs, col.design$test.col[j]) &
                                     results$var == indeling_rijen$inhoud[i]),
                             c("val", "crossing", "crossing.val", "sign", "sign.vs", "n.unweighted", "perc.weighted")]
          if (nrow(data.tmp) == 0) next
          data.tmp$col.index = j
          data.var = bind_rows(data.var, data.tmp)
        }
      }
      
      n_var = data.var %>%
        group_by(col.index) %>%
        summarize(n=sum(n.unweighted, na.rm=T)) %>%
        full_join(data.frame(col.index=col.design$col.index, dummy=0), by="col.index") %>% # zorgen dat alle waardes aanwezig zijn, ook bij lege kolommen
        mutate(n=coalesce(n, 0)) %>%
        arrange(col.index) %>%
        select(-dummy) %>%
        mutate(n=sprintf("n=%d", n)) %>%
        column_to_rownames("col.index") %>%
        t() %>%
        as.data.frame() %>%
        mutate(label="Aantal respondenten",
               matchcode=paste0(indeling_rijen$inhoud[i], "_n")) %>%
        relocate(matchcode, label)
      
      # voor het schrijven naar Excel is een matrix met getallen makkelijker
      # het kan voorkomen dat niet alle antwoordmogelijkheden in elke subset aanwezig zijn
      # daarom nemen we hier de bekende labels, i.p.v. de voorkomende waardes
      output = matrix(nrow=length(var_labels$val[var_labels$var == indeling_rijen$inhoud[i] & var_labels$val != "var"]), ncol=nrow(col.design))
      colnames(output) = col.design$col.index
      # het kan in zeldzame gevallen voorkomen dat er meer dan 10 antwoorden zijn
      # in zo'n geval zal sort() er 1 10 11 12 2 3 4 van maken, omdat het strings zijn
      # voor de indeling zijn we echter wel afhankelijk van een character... dus dubbele omzetting!
      rownames(output) = as.character(sort(as.numeric(var_labels$val[var_labels$var == indeling_rijen$inhoud[i] & var_labels$val != "var"])))
      
      # dit zou in theorie ook zonder for kunnen, maar overzichtelijkheid
      for (j in 1:nrow(col.design)) {
        # N.B.: as.character() is hier nodig omdat R niet om kan gaan met een numerieke rijnaam 0, maar wel met karakter "0"
        vals = as.character(data.var$val[data.var$col.index == j & data.var$val %in% rownames(output)])
        output[vals,j] = data.var$perc.weighted[data.var$col.index == j & data.var$val %in% rownames(output)]
        if (!is.null(algemeen$weergave) && algemeen$weergave == "n")
          output[vals,j] = data.var$n[data.var$col.index == j & data.var$val %in% rownames(output)]
        
        # waarden onder de afkapgrens vervangen
        output[which(output[,j] <= algemeen$afkapwaarde_antwoord),j] = A_TOOSMALL
        
        #PS:
        #Metingen die o.b.v te lage aantallen zijn vervangen 
        if (!is.na(indeling_rijen$verberg_crossings[i]) && !is.na(col.design$crossing[j])) {
          output[,j] = Q_MISSING
        } else if (sum(data.var$n.unweighted[data.var$col.index == j], na.rm=T) == 0) {
          output[,j] = Q_MISSING
        } else if (sum(data.var$n.unweighted[data.var$col.index == j], na.rm=T) < algemeen$min_observaties_per_vraag) {
          #Alle percentages wegstrepen als aantallen per groep te klein zijn.
          output[,j] <- Q_TOOSMALL
        } else if(any(data.var$n.unweighted[data.var$col.index == j] < algemeen$min_observaties_per_antwoord, na.rm=T)) {
          # Bij een cel met te weinig antwoorden zijn er twee opties:
          # 1) De hele kolom verbergen, om herleidbaarheid te voorkomen.
          # 2) Alleen die cel verbergen.
          # De keuze hierin is discutabel, dus we laten het over aan de onderzoekers zelf.
          if (algemeen$vraag_verbergen_bij_missend_antwoord) {
            #Alle percentages wegstrepen als tenminste 1 van de aantallen per antwoord te klein is.
            output[,j] <- A_TOOSMALL
          }
          else {
            # alleen de cel wegstrepen
            data.col = data.var[data.var$col.index == j & data.var$val %in% rownames(output),]
            output[rownames(output) %in% data.col$val[which(data.col$n.unweighted < algemeen$min_observaties_per_antwoord)],j] <- A_TOOSMALL
          }
        }
        
        # is de hele kolom leeg? dan mist er waarschijnlijk data
        if (all(is.na(output[,j])))
          output[is.na(output[,j]),j] = Q_MISSING
        
        # als er nu nog missende getallen zijn betekent dat dat er geen respondenten waren met dat antwoord
        output[is.na(output[,j]),j] = A_TOOSMALL
      }
      
      # significante resultaten zichtbaar maken
      sign = data.var[which(data.var$sign < algemeen$confidence_level), c("val", "col.index", "sign.vs", "perc.weighted")]
      # resultaten verbergen als ze afgerond hetzelfde zijn?
      if (algemeen$sign_verbergen_wanneer_afgerond_gelijk && nrow(sign) > 0) {
        sign.vs = lapply(1:nrow(sign), \(x) data.var[which(data.var$col.index == sign$sign.vs[x] & data.var$val == sign$val[x]), c("val", "col.index", "perc.weighted")]) %>%
          bind_rows() %>%
          rename(perc.vs=perc.weighted) %>%
          distinct()
        sign = sign %>%
          left_join(sign.vs, by=c("val", "sign.vs"="col.index")) %>%
          filter(round(perc.weighted) != round(perc.vs) | is.na(perc.vs))
      }
      # kolomgetal vervangen, indien nodig
      if (nrow(sign.replacement) > 0) {
        for (j in 1:nrow(sign.replacement)) {
          sign$col.index[which(sign$col.index == sign.replacement$old[j])] = sign.replacement$new[j]
        }
      }
      
      # zijn er aparte wensen qua weergave van antwoordmogelijkheden?
      if (!is.na(indeling_rijen$waardes[i])) {
        desired_answers = str_split(indeling_rijen$waardes[i], fixed("|")) %>% unlist() %>% str_trim()
        # wanneer er slechts 1 rij geselecteerd wordt verandert de matrix in een lijst, waardoor alles explodeert
        # oplossing: maak er een data.frame van
        output = as.data.frame(output)
        output = output[desired_answers,]
      }
      
      sign$rij = sapply(sign$val, function (v) ifelse(v %in% rownames(output), which(rownames(output) == v), NA))
      sign = sign[!is.na(sign$rij),]
      
      # dichotoom? zo ja, alleen 1 (= ja) laten zien en geen kop met de vraag
      # zo nee, kop met de vraag en alle waardes laten zien
      levels.var = sort(as.numeric(unique(data.var$val)))
      dichotoom.vals = algemeen$waarden_dichotoom %>% 
        str_split("\\|") %>%
        unlist() %>%
        str_split(",") %>%
        lapply(as.numeric)
      
      if (!indeling_rijen$inhoud[i] %in% niet_dichotoom &&
          (indeling_rijen$inhoud[i] %in% dichotoom ||
           isTRUE(all.equal(levels.var, c(0, 1))) ||
           any(unlist(lapply(dichotoom.vals, function (x) { return(identical(x, levels.var)) }))))) {
        output = output %>% as.data.frame() %>% rownames_to_column("val") %>% filter(val == 1) %>%
          mutate(label=var_labels$label[var_labels$var == indeling_rijen$inhoud[i] & var_labels$val == "var"],
                 matchcode=paste0(indeling_rijen$inhoud[i], val)) %>%
          relocate(matchcode, label) %>%
          select(-val)
        
        # significantie ook aanpassen naar 1 regel
        sign = sign %>% filter(val == 1) %>% mutate(rij=1)
        
        # is de vorige regel een kop of vraag? dan headers en percentages toevoegen
        # als de vorige regel een vraag is dan heeft deze dat al
        if (i > 1 && indeling_rijen$type[i-1] == "kop" && is.na(indeling_rijen$kolomkoppen[i])) {
          # ruimte vrijhouden voor het later invoegen van headers
          header.col.rows = c(header.col.rows, c)
          c = c + header.col.nrows
          
          # regel met aantallen, indien gewenst
          if (indeling_rijen$type[i] == "nvar") {
            writeData(wb, subset.name, n_var, startCol=1, startRow=c, colNames=F)
            addStyle(wb, subset.name, style.num, cols=3:n.col.total, rows=c, gridExpand=T, stack=T)
            data.rows = c(data.rows, c)
            c = c + 1
          }
          
          # regel met procenttekens
          writeData(wb, subset.name, t(rep("%", n.col)), startCol=3, startRow=c, colNames=F)
          addStyle(wb, subset.name, style.perc, cols=3:n.col.total, rows=c, gridExpand=T, stack=T)
          perc.rows = c(perc.rows, c)
          c = c + 1
        } else if (indeling_rijen$type[i] == "nvar") {
          if (i > 1 && indeling_rijen$type[i-1] == "nvar") {
            msg("Let op! De variabele %s op rij %d heeft als type 'nvar', maar dit betreft een dichotome variabele in een lijst. Alleen het aantal respondenten van de eerste variabele wordt weergegeven.",
                indeling_rijen$inhoud[i], i, level=WARN)
          } else {
            writeData(wb, subset.name, n_var, startCol=1, startRow=c, colNames=F)
            addStyle(wb, subset.name, style.num, cols=3:n.col.total, rows=c, gridExpand=T, stack=T)
            data.rows = c(data.rows, c)
            c = c + 1
          }
        }
      } else { # niet dichotoom
        # labels toevoegen
        output = output %>% as.data.frame() %>% rownames_to_column("val") %>%
          mutate(label=sapply(val, function(v) var_labels$label[var_labels$var == indeling_rijen$inhoud[i] & var_labels$val == as.character(v)]),
                 matchcode=paste0(indeling_rijen$inhoud[i], val)) %>%
          relocate(matchcode, label) %>%
          select(-val)
        
        c = c + 1 # witregel na vorig blok
        
        # kop met de vraag toevoegen
        writeData(wb, subset.name, data.frame(col1=indeling_rijen$inhoud[i],
                                              col2=var_labels$label[var_labels$var == indeling_rijen$inhoud[i] & var_labels$val == "var"]),
                  startCol=1, startRow=c, colNames=F)
        addStyle(wb, subset.name, style.question, cols=2:n.col.total, rows=c, stack=T)
        title.rows = c(title.rows, c)
        c = c + 1
        
        # ruimte vrijhouden voor het later invoegen van headers, indien gewenst
        if (is.na(indeling_rijen$kolomkoppen[i])) {
          header.col.rows = c(header.col.rows, c)
          c = c + header.col.nrows
        }
        
        # regel met aantallen, indien gewenst
        if (indeling_rijen$type[i] == "nvar") {
          writeData(wb, subset.name, n_var, startCol=1, startRow=c, colNames=F)
          addStyle(wb, subset.name, style.num, cols=3:n.col.total, rows=c, gridExpand=T, stack=T)
          data.rows = c(data.rows, c)
          c = c + 1
        }
        
        # regel met procenttekens
        writeData(wb, subset.name, t(rep("%", n.col)), startCol=3, startRow=c, colNames=F)
        addStyle(wb, subset.name, style.perc, cols=3:n.col.total, rows=c, gridExpand=T, stack=T)
        perc.rows = c(perc.rows, c)
        
        c = c + 1
      }
      
      labels.oversized = which(str_length(output[,2]) > design("label_max_lengte"))
      if (length(labels.oversized) > 0) {
        label.oversized.rows = c(label.oversized.rows, c + labels.oversized - 1)
      }
      
      # daadwerkelijke data wegschrijven
      writeData(wb, subset.name, output, startCol=1, startRow=c, colNames=F)
      addStyle(wb, subset.name, style.text, cols=2, rows=c:(c+nrow(output)), stack=T)
      addStyle(wb, subset.name, style.num, cols=3:n.col.total, rows=c:(c+nrow(output)), gridExpand=T, stack=T)
      
      # missende waardes goed weergeven
      missing = which(output == Q_MISSING | output == Q_TOOSMALL | output == A_TOOSMALL, arr.ind=T)
      if (length(missing) > 0) {
        # vanwege gekke R logica mag de waarde geen naam hebben (unname) en moet het gedwongen een matrix zijn
        replacement = data.frame(row=missing[,1], col=missing[,2], val=matrix(unname(output[missing])))
        
        for (i in 1:nrow(replacement)) {
          val = algemeen$tekst_missende_data
          if (as.numeric(replacement$val[i]) == Q_TOOSMALL) val = algemeen$tekst_min_vraag_niet_gehaald
          if (as.numeric(replacement$val[i]) == A_TOOSMALL) val = algemeen$tekst_min_antwoord_niet_gehaald
          writeData(wb, subset.name, val, startCol=replacement$col[i], startRow=c+replacement$row[i]-1, colNames=F)
        }
      }
      
      # significantie weergeven
      if (nrow(sign) > 0) {
        addStyle(wb, subset.name, style.sign, rows=c-1+sign$rij, cols=2+sign$col.index, stack=T)
      }
      data.rows = c(data.rows, c:(c+nrow(output)-1))
      c = c + nrow(output)
    }
  }
  
  # laatste regel aanhouden als einde, zodat de kleuren niet doorlopen
  c = c - 1
  
  # algemene opmaak
  setColWidths(wb, subset.name, cols=1, hidden=T) # 1e kolom verbergen, die is alleen voor eigen naslag
  setColWidths(wb, subset.name, cols=2, width=design("kolombreedte_antwoorden")) # 2de kolom breder voor de tekstlabels
  setColWidths(wb, subset.name, cols=3:n.col.total, width=design("kolombreedte")) # 3 tot nde kolom vaste breedte
  setRowHeights(wb, subset.name, rows=table.start:c, heights=design("rij_hoogte")) # normale rijhoogte, ná de intro
  
  if (length(label.oversized.rows) > 0) {
    setRowHeights(wb, subset.name, rows=label.oversized.rows, heights=design("rij_hoogte")*2)
  }
  
  addStyle(wb, subset.name, createStyle(wrapText=T), rows=table.start:c, cols=1:n.col.total, gridExpand=T, stack=T)
  addStyle(wb, subset.name, createStyle(fgFill="#ffffff"), rows=which(!1:c %in% title.rows), cols=1:n.col.total, gridExpand=T, stack=T)
  
  # cellen met koppen/titels/vragen samenvoegen, zodat de tekst de gehele breedte kan gebruiken
  for (i in title.rows) {
    mergeCells(wb, subset.name, cols=2:n.col.total, rows=i)
  }
  
  # headers toevoegen
  # deze kunnen 1 of 2 regels beslaan:
  # <dataset>                   | <dataset 2>    vs.   kolom 1 | kolom 2 | totaal <dataset> | totaal <dataset 2> 
  # kolom 1 | kolom 2 | totaal | totaal 
  output = matrix(nrow=header.col.nrows, ncol=n.col)
  if (header.col.nrows == 2) {
    # datasets op de bovenste rij
    output[1, sapply(unique(col.design$dataset), function (v) {
      return(min(col.design$col.index[col.design$dataset == v]))
    } )] = datasets$naam_dataset[unique(col.design$dataset)]
  }
  for (i in 1:nrow(col.design)) {
    # indien crossing, label van de waarde
    if (!is.na(col.design$crossing.lab[i])) {
      col.name = col.design$crossing.lab[i]
    } else {
      if (!is.na(col.design$name[i])) {
        col.name = col.design$name[i]
        col.name = str_replace(str_replace(col.name, fixed("[naam]"), col.name), fixed("[jaar]"), ifelse(!is.na(col.design$year[i]), col.design$year[i], ""))
      } else {
        # geen naam opgegeven; zelf maken
        col.name = datasets$naam_dataset[col.design$dataset[i]]
        if (!is.na(col.design$subset[i])) {
          col.name = subset.name
          if (col.design$subset[i] != subset) {
            col.name = var_labels$label[var_labels$var == col.design$subset[i] & var_labels$val == subsetmatches[subsetmatches[,1] == subset.val, col.design$subset[i]]]
          }
        }
        
        # tekstopmaak is in te stellen in de configuratie -> [naam] en [jaar] worden vervangen
        col.name = str_replace(str_replace(design("header_template"), fixed("[naam]"), col.name), fixed("[jaar]"), ifelse(!is.na(col.design$year[i]), col.design$year[i], ""))
      }
    }
    
    # als er afkortingen zijn: deze toevoegen
    # deze kunnen normale tekst zijn, of een reguliere expressie (aangegeven met beginteken *)
    if (nrow(headers_afkortingen) > 0) {
      for (j in 1:nrow(headers_afkortingen)) {
        if (str_starts(headers_afkortingen$tekst[j], fixed("*"))) {
          # reguliere expressie
          col.name = str_replace(col.name, str_sub(headers_afkortingen$tekst[j], start=2), headers_afkortingen$vervanging[j])
        } else {
          # normale tekst
          col.name = str_replace(col.name, fixed(headers_afkortingen$tekst[j]), headers_afkortingen$vervanging[j])
        }
      }
    }
    
    output[header.col.nrows, i] = col.name
  }
  
  for (i in header.col.rows) {
    writeData(wb, subset.name, output, startCol=3, startRow=i, colNames=F)
    setRowHeights(wb, subset.name, rows=i:(i+header.col.nrows-1), heights=design("rij_hoogte_kop"))
    addStyle(wb, subset.name, style.header.col, rows=i:(i+header.col.nrows-1), cols=2 + which(is.na(col.design$crossing)), gridExpand=T, stack=T)
    if (design("crossing_headers_kleiner") && sum(!is.na(col.design$crossing)) > 0) {
      addStyle(wb, subset.name, style.header.col.crossing, rows=i:(i+header.col.nrows-1), cols=2 + which(!is.na(col.design$crossing)), gridExpand=T, stack=T)
    } else {
      addStyle(wb, subset.name, style.header.col, rows=i:(i+header.col.nrows-1), cols=2 + which(!is.na(col.design$crossing)), gridExpand=T, stack=T)
    }
    
    if (header.col.nrows > 1) {
      # bovenste rij dikgedrukt, dat zijn de datasets
      addStyle(wb, subset.name, style.header.col, rows=i, cols=2:n.col.total, gridExpand=T, stack=T)
      # cellen samenvoegen, zodat de titel goed over de kolommen staat
      for (dataset in unique(col.design$dataset)) {
        cols = col.design$col.index[col.design$dataset == dataset]
        if (length(cols) > 1 && max(cols - lag(cols), na.rm=T) <= 1) {
          mergeCells(wb, subset.name, cols=cols+2, rows=i)
        }
      }
    }
  }
  
  # afwisselende kleuren?
  if (design("rijen_afwisselend_kleuren")) {
    # data.rows bevat de rijen met data
    # als het verschil tussen twee waardes meer dan 1 is gaat het om een nieuwe categorie, dus daar steeds opnieuw beginnen met tellen
    data.rows.diff = data.rows - lag(data.rows)
    data.rows.blocks = which(data.rows.diff > 1)
    
    data.rows.interval = c()
    for (i in 1:length(data.rows.blocks)) {
      end = ifelse(i < length(data.rows.blocks), data.rows[data.rows.blocks[i + 1]-1], data.rows[length(data.rows)])
      if (data.rows[data.rows.blocks[i]] > end) next
      data.rows.interval = c(data.rows.interval, seq(from=data.rows[data.rows.blocks[i]], to=end, by=2))
    }
    
    addStyle(wb, subset.name, style.gray.bg, rows=data.rows.interval, cols=1:n.col.total, gridExpand=T, stack=T)
  }
  if (design("kolommen_afwisselend_kleuren")) {
    addStyle(wb, subset.name, style.gray.bg, rows=1:c, cols=seq(from=3, to=n.col.total, by=2), gridExpand=T, stack=T)
  }
  if (design("kolommen_crossings_kleuren")) {
    col.design = col.design %>% group_by(dataset, subset, year, crossing)
    cols = group_rows(col.design)[order(sapply(group_rows(col.design),'[[',1))]
    gray = T
    for (i in cols) {
      if (gray) {
        addStyle(wb, subset.name, style.gray.bg, rows=c(data.rows, perc.rows, header.col.rows + ifelse(design("header_stijl") == "enkel", 0, 1)),
                 cols=2+i, gridExpand=T, stack=T)
      }
      gray = !gray
    }
  }
  
  # TODO: borders om de tabellen?
  
  # plaatjes toevoegen?
  if (nrow(logos) > 0) {
    for (i in 1:nrow(logos)) {
      rij = logos$rij[i]
      kolom = logos$kolom[i]
      # bij negatieve waarden doen we aantal - waarde
      if (rij < 0) {
        rij = c - abs(rij)
      }
      if (kolom < 0) {
        kolom = n.col.total - abs(kolom) + 1
      }
      
      insertImage(wb, subset.name, file=logos$bestand[i], width=logos$breedte[i], height=logos$hoogte[i], units="px",
                  startRow=rij, startCol=kolom)
    }
  }

  
  # en als laatste... opslaan!
  tryCatch({
    saveWorkbook(wb, sprintf("output/%s.xlsx", str_replace(filename, "\\\\|/", " ")), overwrite=T)
    msg("Tabellenboek voor %s opgeslagen.", filename, level=MSG)
  },
  error = function(e){
    msg("Er is een fout opgetreden bij het maken van het Excelbestand: %s", e$message, level=MSG)
  },
  warning = function(e){
    msg("Er is een fout opgetreden bij het maken van het Excelbestand: %s", e$message, level=MSG)
  })
  
}
