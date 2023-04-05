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

design = function (var) {
  if (str_length(var) <= 1) {
    msg("Variabelenaam kan niet leeg zijn.", level=ERR)
  }
  
  if (var %in% opmaak$type) {
    ret = opmaak$waarde[opmaak$type == var]
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
  
  return (ret)
}

# col.design = kolom_opbouw
# subset = "Gemeentecode"
# subset.val = 197
# subset.name = "Aalten"
# i = 7

MakeExcel = function (results, var_labels, col.design, subset, subset.val, subsetmatches,
min_observaties_per_vraag,min_observaties_per_antwoord, tekst_min_vraag_niet_gehaald, tekst_min_antwoord_niet_gehaald) {
  subset.name = names(subset.val)
  subset.val = unname(subset.val)
  
  if (is.na(subset.name) || is.null(subset.name))
    subset.name = "Overzicht"
  
  wb = createWorkbook(creator="GGData Tabellenboek")
  addWorksheet(wb, subset.name)
  
  # opmaak
  modifyBaseFont(wb, fontSize=design("font_size"), fontColour=design("font_color"), fontName=design("font_type"))
  
  style.sign = createStyle(textDecoration="bold") # significante resultaten
  style.title = createStyle(fontSize=design("titel_size"), fontColour=design("titel_color"),
                            textDecoration=design("titel_decoration"), fgFill=design("titel_fill")) # titels
  style.subtitle = createStyle(fontSize=design("kop_size"), fontColour=design("kop_color"),
                               textDecoration=design("kop_decoration"), fgFill=design("kop_fill")) # koppen
  style.text = createStyle(wrapText=T)
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
  
  c = 1 # teller voor rijen in Excel
  
  # moet er introtekst bij?
  if (nrow(intro_tekst) > 0) {
    intro_tekst$type = str_to_lower(str_trim(intro_tekst$type))
    intro_tekst$inhoud = str_replace_all(intro_tekst$inhoud, fixed("[naam]"), subset.name)
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
    if (indeling_rijen$type[i] %in% c("titel", "kop", "tekst")) { # regel met een titel, kop, of tekst
      # een extra witregel voor koppen of titels
      if (indeling_rijen$type[i] != "tekst") {
        if (i > 1)
          c = c + 1
      }
      
      output = data.frame(a=indeling_rijen$type[i], b=str_replace_all(indeling_rijen$inhoud[i], "naam_onderdeel", subset.name))
      writeData(wb, subset.name, output, startCol=1, startRow=c, colNames=F)
      
      # TODO: meer opmaak?
      if (indeling_rijen$type[i] == "titel") {
        addStyle(wb, subset.name, style.title, rows=c, cols=1:n.col.total, stack=T)
        title.rows = c(title.rows, c)
      } else if (indeling_rijen$type[i] == "kop") {
        # TODO: willen we hier weer headers toevoegen als de volgende regel een var is?
        addStyle(wb, subset.name, style.subtitle, rows=c, cols=1:n.col.total, stack=T)
        title.rows = c(title.rows, c)
      } else { # tekst
        addStyle(wb, subset.name, style.text, rows=c, cols=1:n.col.total, stack=T)
        mergeCells(wb, subset.name, cols=2:n.col.total, rows=c)  
      }
      
      c = c + 1 #ifelse(indeling_rijen$type[i] == "tekst", 1, 2) # een extra witregel na een kop of titel
    } else if (indeling_rijen$type[i] == "aantallen") { # regel met aantal deelnemers toevoegen
      output = matrix(nrow=1, ncol=nrow(col.design))
      for (j in 1:nrow(col.design)) {
        if (!is.na(col.design$subset[j])) {
          subset.col = subset.val
          if (col.design$subset[j] != subset) {
            subset.col = subsetmatches[subsetmatches[,1] == subset.val, col.design$subset[j]]
          }
          
          n = results[which(NA.identical(results$dataset, col.design$dataset[j]) & NA.identical(results$subset, col.design$subset[j]) &
                              NA.identical(results$subset.val, subset.col) & 
                              NA.identical(results$year, col.design$year[j]) & NA.identical(results$crossing, col.design$crossing[j]) &
                              NA.identical(results$crossing.val, col.design$crossing.val[j])),] %>%
            as.data.frame() %>% distinct() %>% group_by(var) %>% summarize(n=sum(n.unweighted, na.rm=T))
          output[j] = max(n$n, na.rm=T)
        } else {
          # het maximale aantal per vraag is het aantal deelnemers
          # niet iedere vraag is volledig beantwoord, dus we nemen het hoogste getal
          n = results[which(NA.identical(results$dataset, col.design$dataset[j]) & 
                              NA.identical(results$year, col.design$year[j]) & NA.identical(results$crossing, col.design$crossing[j]) &
                              NA.identical(results$crossing.val, col.design$crossing.val[j])),] %>%
            as.data.frame() %>% distinct() %>% group_by(var) %>% summarize(n=sum(n.unweighted, na.rm=T))
          output[j] = max(n$n, na.rm=T)
        }
      }
      #PS: Aantallen die lager dan min. per vraag uitvallen vervangen. 
      #Matrix veranderd door vervanging in string. Geen impact op output in Excel.
      if(!is.null(min_observaties_per_vraag)){
       output[output < min_observaties_per_vraag] <- tekst_min_vraag_niet_gehaald
      }      

      # ruimte vrijhouden voor het later invoegen van headers
      header.col.rows = c(header.col.rows, c)
      c = c + header.col.nrows
      
      writeData(wb, subset.name, "aantallen", startCol=1, startRow=c, colNames=F)
      writeData(wb, subset.name, output, startCol=3, startRow=c, colNames=F)
      addStyle(wb, subset.name, style.num, rows=c, cols=3:n.col.total, stack=T)
      data.rows = c(data.rows, c)
      
      c = c + 1
    } else if (indeling_rijen$type[i] == "var") { # variabele toevoegen
      # benodigde resultaten ophalen, zodat we niet steeds belachelijke selectors nodig hebben
      data.var = data.frame()
      for (j in 1:nrow(col.design)) {
        if (!is.na(col.design$subset[j])) {
          subset.col = subset.val
          if (col.design$subset[j] != subset) {
            subset.col = subsetmatches[subsetmatches[,1] == subset.val, col.design$subset[j]]
          }
          data.tmp = results[which(NA.identical(results$dataset, col.design$dataset[j]) & NA.identical(results$subset, col.design$subset[j]) &
                                     NA.identical(results$year, col.design$year[j]) & NA.identical(results$crossing, col.design$crossing[j]) &
                                     NA.identical(results$crossing.val, col.design$crossing.val[j]) & results$var == indeling_rijen$inhoud[i] &
                                     NA.identical(results$subset.val, subset.col)), c("val", "crossing", "crossing.val", "sign", "sign.vs", "n.unweighted", "perc.weighted")] %>%
            distinct() # aangezien een tweede subset vaker kan draaien moeten we hier een distinct op doen
          data.tmp$col.index = j
          data.var = bind_rows(data.var, data.tmp)
        }
        else {
          data.tmp = results[which(NA.identical(results$dataset, col.design$dataset[j]) & NA.identical(results$subset, col.design$subset[j]) &
                                     NA.identical(results$year, col.design$year[j]) & NA.identical(results$crossing, col.design$crossing[j]) &
                                     NA.identical(results$crossing.val, col.design$crossing.val[j]) &
                                     results$var == indeling_rijen$inhoud[i]), c("val", "crossing", "crossing.val", "sign", "sign.vs", "n.unweighted", "perc.weighted")] %>%
            distinct() # aangezien een tweede subset vaker kan draaien moeten we hier een distinct op doen
          data.tmp$col.index = j
          data.var = bind_rows(data.var, data.tmp)
        }
      }
      
      # voor het schrijven naar Excel is een matrix met getallen makkelijker
      output = matrix(nrow=length(unique(data.var$val)), ncol=nrow(col.design))
      rownames(output) = sort(unique(data.var$val))
      
      # dit zou in theorie ook zonder for kunnen, maar overzichtelijkheid
      for (j in 1:nrow(col.design)) {
        # TODO: afkappunten met minimale metingen toevoegen
        # N.B.: as.character() is hier nodig omdat R niet om kan gaan met een numerieke rijnaam 0, maar wel met karakter "0"
        output[as.character(data.var$val[data.var$col.index == j]),j] = data.var$perc.weighted[data.var$col.index == j]

#PS: round(,2) toegevoegd
        #Vervangen van cellen met te lage aantallen forceert dat de matrix een character matrix wordt
        #Door deze conversie blijven de volledige getallen bewaard als character en zal Excel ze niet
        #zelf afronden. Percentages moeten dus voor deze conversie afgerond worden
        output[data.var$val[data.var$col.index == j],j] = round(data.var$perc.weighted[data.var$col.index == j],2)
        
        
        #PS:
        #Metingen die o.b.v te lage aantallen zijn vervangen        
        #Alle percentages wegstrepen als tenminste 1 van de aantallen per antwoord te klein is.
        if(any(data.var$n.unweighted[data.var$col.index == j] < min_observaties_per_antwoord))
          output[data.var$val[data.var$col.index == j],j] <- tekst_min_antwoord_niet_gehaald
         
        #Alle percentages wegstrepen als aantallen per groep te klein zijn.
        if(sum(data.var$n.unweighted[data.var$col.index == j]) < min_observaties_per_vraag)
          output[data.var$val[data.var$col.index == j],j] <- tekst_min_vraag_niet_gehaald
      }
      
      # significante resultaten zichtbaar maken
      sign = data.var[which(data.var$sign < algemeen$confidence_level), c("val", "col.index")]
      sign$rij = sapply(sign$val, function (v) which(rownames(output) == v))
      
      # dichotoom? zo ja, alleen 1 (= ja) laten zien en geen kop met de vraag
      # zo nee, kop met de vraag en alle waardes laten zien
      levels.var = sort(as.numeric(unique(data.var$val)))
      dichotoom.vals = algemeen$waarden_dichotoom %>% 
        str_split("\\|") %>%
        unlist() %>%
        str_split(",") %>%
        lapply(as.numeric)
      # TODO: tabbladen met dichotoom/niet_dichotoom?
      if (isTRUE(all.equal(levels.var, c(0, 1))) || any(unlist(lapply(dichotoom.vals, function (x) { return(identical(x, levels.var)) })))) {
        output = output %>% as.data.frame() %>% rownames_to_column("val") %>% filter(val == 1) %>%
          mutate(label=var_labels$label[var_labels$var == indeling_rijen$inhoud[i] & var_labels$val == "var"],
                 matchcode=paste0(indeling_rijen$inhoud[i], val)) %>%
          relocate(matchcode, label) %>%
          select(-val)
        
        # significantie ook aanpassen naar 1 regel
        sign = sign %>% filter(val == 1)
        
        # is de vorige regel een kop? dan headers en percentages toevoegen
        if (i > 1 && indeling_rijen$type[i-1] == "kop") {
          # ruimte vrijhouden voor het later invoegen van headers
          header.col.rows = c(header.col.rows, c)
          c = c + header.col.nrows
          
          # TODO: nvar toevoegen?
          # regel met procenttekens
          writeData(wb, subset.name, t(rep("%", n.col)), startCol=3, startRow=c, colNames=F)
          addStyle(wb, subset.name, style.perc, cols=3:n.col.total, rows=c, gridExpand=T, stack=T)
          perc.rows = c(perc.rows, c)
          c = c + 1
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
        addStyle(wb, subset.name, style.subtitle, cols=2:n.col.total, rows=c, stack=T)
        title.rows = c(title.rows, c)
        c = c + 1
        
        # ruimte vrijhouden voor het later invoegen van headers
        header.col.rows = c(header.col.rows, c)
        c = c + header.col.nrows
        
        # TODO: nvar toevoegen?
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
      output[header.col.nrows, i] = col.design$crossing.lab[i]
      next
    }
    
    # totaalkolom - subset?
    col.name = datasets$naam_dataset[col.design$dataset[i]]
    if (!is.na(col.design$subset[i])) {
      col.name = subset.name
      if (col.design$subset[i] != subset) {
        col.name = var_labels$label[var_labels$var == col.design$subset[i] & var_labels$val == subsetmatches[subsetmatches[,1] == subset.val, col.design$subset[i]]]
      }
    }
    
    # als er afkortingen zijn: deze toevoegen
    if (nrow(headers_afkortingen) > 0) {
      for (j in 1:nrow(headers_afkortingen)) {
        col.name = str_replace(col.name, fixed(headers_afkortingen$tekst[j]), headers_afkortingen$vervanging[j])
      }
    }
    
    # tekstopmaak is in te stellen in de configuratie -> [naam] en [jaar] worden vervangen
    output[header.col.nrows, i] = str_replace(str_replace(design("header_template"), fixed("[naam]"), col.name), fixed("[jaar]"), col.design$year[i])
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
        if (length(cols) > 1) {
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
      
      data.rows.interval = c(data.rows.interval, seq(from=data.rows[data.rows.blocks[i]]+1, to=end, by=2))
    }
    
    addStyle(wb, subset.name, style.gray.bg, rows=data.rows.interval, cols=1:n.col.total, gridExpand=T, stack=T)
  }
  if (design("kolommen_afwisselend_kleuren")) {
    addStyle(wb, subset.name, style.gray.bg, rows=1:c, cols=seq(from=3, to=n.col.total, by=2), gridExpand=T, stack=T)
  }
  if (design("kolommen_crossings_kleuren")) {
    col.design = col.design %>% group_by(dataset, subset, year, crossing)
    gray = T
    for (i in group_rows(col.design)) {
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
      
      insertImage(wb, naam_sheet, file=logos$bestand[i], width=logos$breedte[i], height=logos$hoogte[i], units="px",
                  startRow=rij, startCol=kolom)
    }
  }

  
  # en als laatste... opslaan!
  saveWorkbook(wb, sprintf("output/%s.xlsx", subset.name), overwrite=T)
  msg("Tabellenboek voor %s opgeslagen.", subset.name, level=MSG)
}