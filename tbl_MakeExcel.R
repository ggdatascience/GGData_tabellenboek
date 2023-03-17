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
MakeExcel = function (results, col.design, subset, subset.val, subsetmatches) {
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
                            textDecoration=design("titel_decoration"), bgFill=design("titel_fill")) # titels
  style.subtitle = createStyle(fontSize=design("kop_size"), fontColour=design("kop_color"),
                               textDecoration=design("kop_decoration"), bgFill=design("kop_fill")) # koppen
  style.text = createStyle(wrapText=T)
  style.header.col = createStyle(halign="center", valign="center", textDecoration="bold") # kolomkoppen
  style.perc = createStyle(halign="center", valign="center") # percentagetekens
  style.num = createStyle(numFmt="0", halign="center", valign="center") # cijfers -> 0 betekent hele getallen zonder decimalen, 0.0 -> 1 decimaal, etc.
  
  # instellingen
  header.col.nrows = 1 # ifelse(opmaak("dubbel"), 2, 1) # aantal rijen per kolomheader
  header.col.template = "Totaal [naam] [jaar]" # tekstinvulling van de kolomkop -> [naam] en [jaar] worden vervangen.
  
  # opslag voor rijen die later aangemaakt of opgemaakt moeten worden
  header.col.rows = numeric(0)
  
  c = 1 # teller voor rijen in Excel
  indeling_rijen$type = str_to_lower(str_trim(indeling_rijen$type))
  n.col = nrow(col.design) # aantal kolommen in de data
  n.col.total = n.col + 2 # aantal kolommen in de sheet (1: matchcode, 2: label)
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
      } else if (indeling_rijen$type[i] == "kop") {
        # TODO: willen we hier weer headers toevoegen als de volgende regel een var is?
        addStyle(wb, subset.name, style.subtitle, rows=c, cols=1:n.col.total, stack=T)
      } else { # tekst
        addStyle(wb, subset.name, style.text, rows=c, cols=1:n.col.total, stack=T)
        mergeCells(wb, subset.name, cols=2:n.col.total, rows=c)  
      }
      
      c = c + ifelse(indeling_rijen$type[i] == "tekst", 1, 2) # een extra witregel na een kop of titel
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
      
      # ruimte vrijhouden voor het later invoegen van headers
      header.col.rows = c(header.col.rows, c)
      c = c + header.col.nrows
      
      writeData(wb, subset.name, "aantallen", startCol=1, startRow=c, colNames=F)
      writeData(wb, subset.name, output, startCol=3, startRow=c, colNames=F)
      c = c + 2
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
        output[data.var$val[data.var$col.index == j],j] = data.var$perc.weighted[data.var$col.index == j]
      }
      
      # significante resultaten zichtbaar maken
      sign = data.var[which(data.var$sign < algemeen$confidence_level), c("val", "col.index")]
      sign$rij = sapply(sign$val, function (v) which(rownames(output) == v))
      
      # labels toevoegen
      output = output %>% as.data.frame() %>% rownames_to_column("val") %>%
        mutate(label=sapply(val, function(v) val_label(data[[indeling_rijen$inhoud[i]]], v)),
               matchcode=paste0(indeling_rijen$inhoud[i], val)) %>%
        relocate(matchcode, label) %>%
        select(-val)
      
      # ruimte vrijhouden voor het later invoegen van headers
      header.col.rows = c(header.col.rows, c)
      c = c + header.col.nrows
      
      # TODO: nvar toevoegen?
      writeData(wb, subset.name, t(rep("%", n.col)), startCol=3, startRow=c, colNames=F)
      addStyle(wb, subset.name, style.perc, cols=3:n.col.total, rows=c, gridExpand=T, stack=T)
      c = c + 1
      
      writeData(wb, subset.name, output, startCol=1, startRow=c, colNames=F)
      addStyle(wb, subset.name, style.num, cols=3:n.col.total, rows=c:(c+nrow(output)), gridExpand=T, stack=T)
      if (nrow(sign) > 0) {
        addStyle(wb, subset.name, style.sign, rows=c-1+sign$rij, cols=2+sign$col.index, stack=T)
      }
      c = c + nrow(output)
    }
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
      output[header.col.nrows, i] = col.design$crossing.lab[i]
      next
    }
    
    # totaalkolom - subset?
    col.name = datasets$naam_dataset[col.design$dataset[i]]
    if (!is.na(col.design$subset[i])) {
      col.name = subset.name
      if (col.design$subset[i] != subset) {
        col.name = val_label(data[[col.design$subset[i]]], subsetmatches[subsetmatches[,1] == subset.val, col.design$subset[i]])
      }
    }
    
    # tekstopmaak is in te stellen in de configuratie -> [naam] en [jaar] worden vervangen
    output[header.col.nrows, i] = str_replace(str_replace(header.col.template, fixed("[naam]"), col.name), fixed("[jaar]"), col.design$year[i])
  }
  
  for (i in header.col.rows) {
    writeData(wb, subset.name, output, startCol=3, startRow=i, colNames=F)
    # TODO: opmaak
  }
  
  # algemene opmaak
  setColWidths(wb, subset.name, cols=1, hidden=T) # 1e kolom verbergen, die is alleen voor eigen naslag
  
  # en als laatste... opslaan!
  saveWorkbook(wb, sprintf("output/%s.xlsx", subset.name), overwrite=T)
  msg("Tabellenboek voor %s opgeslagen.", subset.name, level=MSG)
}