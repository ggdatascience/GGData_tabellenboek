#
#
# Deze functie schrijft de resultatentabel vanuit survey om naar een bruikbare vorm.
# Opmaak wordt ingelezen vanuit de configuratie (tabblad opmaak) en verdere instellingen
# van de tabbladen indeling_rijen, indeling_kolommen en algemeen.
#
#

# constantes voor het vervangen van cellen met te weinig antwoorden
# dit is nodig omdat bij het direct plaatsen van tekst de hele matrix een karakter wordt
A_TOOSMALL = -1
Q_TOOSMALL = -2
Q_MISSING = -3

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
    # binnen Excel is textDecoration één ding, binnen CSS zijn het twee aparte variabelen
    # daarom moeten we hier even een smerig trucje uitvoeren
    
    if (ret == "bold") {
      ret = "font-weight: bold; text-decoration: none; font-style: normal;"
    } else if (ret == "underline") {
      ret = "font-weight: normal; text-decoration: underline; font-style: normal;"
    } else if (ret == "italic") {
      ret = "font-weight: normal; text-decoration: none; font-style: italic;"
    } else {
      ret = ""
    }
    
    if (str_length(ret) <= 1) ret = NULL
  } else if (ret == "TRUE" || ret == "FALSE") {
    ret = ifelse(ret == "TRUE", T, F)
  } else if (str_detect(var, "size") || str_detect(var, "hoogte") || str_detect(var, "breedte") || str_detect(var, "lengte")) {
    ret = as.numeric(ret)
  }
  
  return(ret)
}

BuildHtmlTableRows = function (input, col.design, n=F) {
  # input is als het goed is een data.frame met twee datakolommen (label en sign), en dan een reeks waarden
  # gewenste output is opgemaakte rijen met significantie aangegeven
  
  # kolomindeling, nodig als kolommen per crossing gekleurd moeten worden
  col.design = col.design %>% group_by(dataset, subset, year, crossing)
  colors = group_rows(col.design)[order(sapply(group_rows(col.design),'[[',1))]
  colors = sapply(1:length(colors), function(i) { return(rep(i, length(colors[[i]]))) }) %>% unlist()
  
  output = ""
  for (i in 1:nrow(input)) {
    htmlclass = ""
    if (design("rijen_afwisselend_kleuren")) {
      if (i %% 2 == 0)
        htmlclass = " class=\"rij_b\""
      else
        htmlclass = " class=\"rij_a\""
    }
    
    output = paste0(output, sprintf('<tr%s><th scope="row">%s</th>', htmlclass, HTMLencode(input$label[i])))
    for (c in 1:(ncol(input) - 2)) {
      val = input[i, c+2]
      htmlclass = c()
      
      if (val == A_TOOSMALL) {
        val = algemeen$tekst_min_antwoord_niet_gehaald
        htmlclass = c(htmlclass, "antwoord_niet_gehaald")
      } else if (val == Q_TOOSMALL) {
        val = algemeen$tekst_min_vraag_niet_gehaald
        htmlclass = c(htmlclass, "vraag_niet_gehaald")
      } else if (val == Q_MISSING) {
        val = algemeen$tekst_missende_data
        htmlclass = c(htmlclass, "data_missend")
      } else {
        val = sprintf("%s%.0f", ifelse(n, "n=", ""), val)
      } 
      
      # significant?
      if (colnames(input)[c+2] %in% unlist(input$sign[i])) {
        htmlclass = c(htmlclass, "sign")
      }
      
      # kolom kleuren o.b.v. crossing of index?
      if (design("kolommen_afwisselend_kleuren")) {
        if (c %% 2 == 0)
          htmlclass = c(htmlclass, "kolom_b")
        else
          htmlclass = c(htmlclass, "kolom_a")
      } else if (design("kolommen_crossings_kleuren")) {
        if (colors[c] %% 2 == 0)
          htmlclass = c(htmlclass, "kolom_b")
        else
          htmlclass = c(htmlclass, "kolom_a")
      }
      
      # klassen toevoegen aan de opmaak
      htmlclass = ifelse(length(htmlclass) > 0, sprintf(' class="%s"', str_c(htmlclass, collapse=" ")), "")
      # significant?
      if (colnames(input)[c+2] %in% unlist(input$sign[i])) {
        htmlclass = paste0(htmlclass, " title=\"", algemeen$sign_hovertekst, "\"")
      }
      
      output = paste0(output, sprintf("<td%s>%s</td>", htmlclass, val))
    }
    output = paste0(output, "</tr>\r\n")
  }
  
  return(output)
}

# col.design = kolom_opbouw
# subset = "Gemeentecode"
# subset.val = 197
# subset.name = "Aalten"
# template = template_html
# i = 7

MakeHtml = function (results, var_labels, col.design, subset, subset.val, subsetmatches, n_resp, template) {
  subset.name = names(subset.val)
  subset.val = unname(subset.val)
  
  if (is.na(subset.name) || is.null(subset.name))
    subset.name = design("naam_tabellenboek")
  
  # basiselementen invullen
  template = str_replace_all(template, fixed("{titel}"), HTMLencode(subset.name))
  
  # opmaak
  design.vars = str_extract_all(template, "\\[[a-zA-Z_]{3,}\\]") %>% unlist() %>% str_sub(start=2, end=-2)
  for (var in design.vars) {
    template = str_replace_all(template, fixed(paste0("[", var, "]")), design(var))
  }
  
  # weergave significantie in andere kolommen?
  sign.replacement = col.design[!col.design$test.display, c("col.index", "test.col")]
  colnames(sign.replacement) = c("old", "new")
  
  # moeten er kolommen verborgen worden?
  # let op: bij deze versie hoeft er geen rekening gehouden te worden met het verspringen van de significantieresultaten, omdat de kolommen vanzelf de juiste naam krijgen
  col.design$n = 0
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
    col.design = col.design[-which(col.design$col.index %in% hide.cols),]
  }
  
  # instellingen
  header.col.nrows = ifelse(design("header_stijl") == "dubbel", 2, 1) # aantal rijen per kolomheader
  n.col = nrow(col.design) # aantal kolommen in de data
  n.col.total = n.col + 1 # aantal kolommen in de sheet (1: label)
  
  # headers klaarzetten
  # deze kunnen 1 of 2 regels beslaan:
  # <dataset>                   | <dataset 2>    vs.   kolom 1 | kolom 2 | totaal <dataset> | totaal <dataset 2> 
  # kolom 1 | kolom 2 | totaal | totaal 
  # hierbij is het belangrijk om de digitoegankelijkheid in de gaten te houden; het eerste type moet goed gescoped worden
  # zie ook: https://www.w3schools.com/tags/att_th_scope.asp
  cols.output = "<colgroup><col /></colgroup>"
  labels.output = "<tr><td />"
  if (header.col.nrows == 2) {
    header.indexes = sapply(unique(col.design$dataset), function (v) {
      return(min(col.design$col.index[col.design$dataset == v]))
    } )
    header.labels = datasets$naam_dataset[unique(col.design$dataset)]
    # kolommenstructuur aangeven
    for (i in 1:length(header.indexes)) {
      if (i < length(header.indexes)) {
        distance = header.indexes[i+1] - header.indexes[i]
      } else {
        distance = n.col - header.indexes[i]
      }
      
      if (distance > 1) {
        cols.output = paste0(cols.output, sprintf("<colgroup span=\"%d\" />", distance))
        labels.output = paste0(labels.output, sprintf("<th colspan=\"%d\" scope=\"colgroup\">%s</th>", distance, HTMLencode(header.labels[i])))
      } else {
        cols.output = paste0(cols.output, "<colgroup><col /></colgroup>")
        labels.output = paste0(labels.output, sprintf("<th scope=\"col\">%s</th>", HTMLencode(header.labels[i])))
      }
    }
    
    cols.output = paste0(cols.output, "\r\n")
    labels.output = paste0(labels.output, "</tr>\r\n<tr><td />")
  }
  
  # afwisselend kleuren van kolommen/groepen?
  col.design = col.design %>% group_by(dataset, subset, year, crossing)
  colors = group_rows(col.design)[order(sapply(group_rows(col.design),'[[',1))]
  colors = sapply(1:length(colors), function(i) { return(rep(i, length(colors[[i]]))) }) %>% unlist()
  # kolomkoppen in HTML gieten
  perc.row.output = "<tr><td />"
  for (i in 1:nrow(col.design)) {
    # kolom kleuren o.b.v. crossing of index?
    htmlclass = NA
    if (design("kolommen_afwisselend_kleuren")) {
      if (i %% 2 == 0)
        htmlclass = "kolom_b"
      else
        htmlclass = "kolom_a"
    } else if (design("kolommen_crossings_kleuren")) {
      if (colors[i] %% 2 == 0)
        htmlclass = "kolom_b"
      else
        htmlclass = "kolom_a"
    }
    htmlclass = ifelse(!is.na(htmlclass), sprintf(' class="%s"', htmlclass), "")
    
    # rij met percentages vullen
    perc.row.output = paste0(perc.row.output, "<td", htmlclass, ">%</td>")
    
    # indien crossing, label van de waarde
    if (!is.na(col.design$crossing.lab[i])) {
      # TODO: iets doen met kleinere labels voor crossings
      # if (design("crossing_headers_kleiner") && sum(!is.na(col.design$crossing)) > 0)
      
      col.name = col.design$crossing.lab[i]
    } else {
      # totaalkolom - naam genereren?
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
    
    labels.output = paste0(labels.output, sprintf("<th scope=\"col\"%s>%s</th>", htmlclass, HTMLencode(col.name)))
  }
  
  # samenvoegen tot een coherent geheel
  header.output = paste0(cols.output, "<thead>\r\n", labels.output, "</thead>\r\n")
  perc.row.output = paste0(perc.row.output, "</tr>\r\n")

  # moet er introtekst bij?
  if (nrow(intro_tekst) > 0) {
    intro_tekst$type = str_to_lower(str_trim(intro_tekst$type))
    intro_tekst$inhoud = str_replace_all(intro_tekst$inhoud, fixed("[naam]"), subset.name)
    # indien gesplistst per subset: andere subsets ook vervangen?
    if (subset.name != design("naam_tabellenboek") && ncol(subsetmatches) > 1) {
      for (i in 1:ncol(subsetmatches)) {
        intro_tekst$inhoud = str_replace_all(intro_tekst$inhoud, fixed(paste0("[", colnames(subsetmatches)[i], "]")), var_labels$label[var_labels$var == colnames(subsetmatches)[i] & var_labels$val == subsetmatches[subset.name,i]])
      }
    }
    
    # html voor titels en koppen toevoegen
    intro_tekst$inhoud[intro_tekst$type == "titel"] = sprintf("<h1>%s</h1>", HTMLencode(intro_tekst$inhoud[intro_tekst$type == "titel"]))
    intro_tekst$inhoud[intro_tekst$type == "kop"] = sprintf("<h2>%s</h2>", HTMLencode(intro_tekst$inhoud[intro_tekst$type == "kop"]))
    intro_tekst$inhoud[intro_tekst$type != "titel" & intro_tekst$type != "kop"] = HTMLencode(intro_tekst$inhoud[intro_tekst$type != "titel" & intro_tekst$type != "kop"])
    # lege waarden worden geprint als NA; dat willen we niet
    intro_tekst$inhoud[is.na(intro_tekst$inhoud)] = ""
    
    # extra witregel voor de volgende sectie, tenzij de laatste regel een witregel is
    if (!is.na(intro_tekst$inhoud[nrow(intro_tekst)])) intro_tekst = bind_rows(intro_tekst, data.frame(type="tekst", inhoud="<br />"))
    
    template = str_replace(template, fixed("{introtekst}"), str_c(intro_tekst$inhoud, collapse="<br />\n"))
  } else {
    # anders is het ook zo jammer...
    template = str_replace(template, fixed("{introtekst}"), "")
  }
  
  # plaatjes toevoegen?
  if (nrow(logos) > 0) {
    img = c()
    
    for (i in 1:nrow(logos)) {
      # rij = logos$rij[i]
      # kolom = logos$kolom[i]
      # # bij negatieve waarden doen we aantal - waarde
      # if (rij < 0) {
      #   rij = c - abs(rij)
      # }
      # if (kolom < 0) {
      #   kolom = n.col.total - abs(kolom) + 1
      # }
      # 
      # insertImage(wb, subset.name, file=logos$bestand[i], width=logos$breedte[i], height=logos$hoogte[i], units="px",
      #             startRow=rij, startCol=kolom)
      
      base_img = sprintf("<img src=\"%s\"%s>",
                         image_uri(logos$bestand[i]),
                         ifelse(!is.na(logos$id[i]), paste0(" id=\"logo_", logos$id[i], "\""), ""))
      
      if (!is.na(logos$id[i]) && str_detect(template, fixed(sprintf("{logo %s}", logos$id[i])))) {
        template = str_replace(template, fixed(sprintf("{logo %s}", logos$id[i])), base_img)
      } else {
        img = c(img, base_img)
      }
    }
    
    template = str_replace(template, fixed("{logo}"), str_c(img, collapse="\n\n"))
    # het kan zijn dat {logo} in de introtekst staat, dan worden de brackets omgezet
    template = str_replace(template, fixed("&lbrace;logo&rbrace;"), str_c(img, collapse="\n\n"))
  }
  
  indeling_rijen$type = str_to_lower(str_trim(indeling_rijen$type))
  table.output = c()
  # er is een tijdelijke opslag nodig voor niet-dichotome variabelen die achter elkaar moeten
  table.cache = NULL # hier is NULL nodig i.p.v. NA, omdat is.na() een vector teruggeeft, en we willen alleen weten of de variabele gevuld is of niet
  question.cache = NA
  n.cache = NULL
  for (i in 1:nrow(indeling_rijen)) {
    # als er iets in de tijdelijke opslag zit en de huidige regel != "var" of "nvar", printen
    if (!is.null(table.cache) && !indeling_rijen$type[i] %in% c("var", "nvar")) {
      table.output = c(table.output, paste0("<table>\r\n",
                                            "<caption>", question.cache, "</caption>\r\n",
                                            header.output, "\r\n",
                                            ifelse(!is.null(n.cache), BuildHtmlTableRows(n.cache, col.design, T), ""),
                                            perc.row.output, "\r\n",
                                            "<tbody>\r\n",
                                            BuildHtmlTableRows(table.cache, col.design),
                                            "</tbody>\r\n",
                                            "</table>\r\n<br />\r\n"))
      question.cache = NA
      table.cache = NULL
      n.cache = NULL
    }
    
    if (indeling_rijen$type[i] %in% c("titel", "kop", "vraag", "tekst")) { # regel met een titel, kop, of tekst
      # een extra witregel voor koppen of titels
      if (!indeling_rijen$type[i] %in% c("tekst", "vraag")) {
        if (i > 1)
          table.output = c(table.output, "<br />\r\n")
      }
      
      output = str_replace_all(indeling_rijen$inhoud[i], "naam_onderdeel", subset.name)
      
      # opmaak toevoegen
      if (indeling_rijen$type[i] == "titel") {
        output = sprintf("<h2 class=\"heading\" id=\"heading_%d\">%s</h2>\r\n", i, HTMLencode(output))
        # als de volgende regel geen kop of vraag is: extra witregel
        if (i < nrow(indeling_rijen) && !indeling_rijen$type[i+1] %in% c("kop", "vraag", "aantallen")) 
          output = paste0(output, "<br />\r\n")
      } else if (indeling_rijen$type[i] == "kop") {
        question.cache = HTMLencode(output)
        output = sprintf("<h3 class=\"heading\" id=\"heading_%d\">%s</h3>", i, HTMLencode(output))
      } else if (indeling_rijen$type[i] == "vraag") {
        question.cache = HTMLencode(output)
        output = sprintf("<h3 class=\"heading vraag\" id=\"heading_%d\">%s</h3>", i, HTMLencode(output))
      } else { # tekst
        output = sprintf("%s<br />", ifelse(is.na(output), "", HTMLencode(output)))
      }
      
      table.output = c(table.output, paste0(output, "\r\n"))
    } else if (indeling_rijen$type[i] == "aantallen") { # regel met aantal deelnemers toevoegen
      output = matrix(data=col.design$n, nrow=1, ncol=nrow(col.design)) %>%
        as.data.frame() %>%
        mutate(label="Aantal deelnemers", sign=NA, .before=1)
      
      table.output = c(table.output, paste0("<table>\r\n",
                                            "<caption>Aantal deelnemers</caption>",
                                            header.output,
                                            "<tbody>\r\n",
                                            BuildHtmlTableRows(output, col.design),
                                            "</tbody>\r\n",
                                            "</table><br />\r\n"))
    } else if (indeling_rijen$type[i] == "var" || indeling_rijen$type[i] == "nvar") { # variabele toevoegen
      if (!indeling_rijen$inhoud[i] %in% colnames(data)) {
        msg("Variabele %s komt niet voor in de resultaten. Deze wordt overgeslagen. Controleer de configuratie.", indeling_rijen$inhoud[i], level=WARN)
        next
      }
      
      # Nu wordt het verhaal een beetje ingewikkeld...
      # We hebben dichotome en niet-dichotome variabelen, waarbij niet-dichotome variabelen vaak onder elkaar worden geplaatst in 1 tabel.
      # Binnen Excel was dit geen probleem; gewoon stiekem een rij toevoegen zonder kop en ergens verzinnen dat er een kop boven moest...
      # Bij HTML-bestanden werkt dit helaas niet zo, daar moeten we de tabel als geheel opzetten.
      # Dit betekent dat we een tijdelijke tabel met voorgaande resultaten moeten opbouwen, welke pas in het document geplakt wordt als
      # er géén dichotome variabele meer volgt. (Let op: er kan een witregel tussen zitten! )
      
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
          data.tmp$col.index = col.design$col.index[j]
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
          data.tmp$col.index = col.design$col.index[j]
          data.var = bind_rows(data.var, data.tmp)
        }
      }
      
      # aantal respondenten per vraag ophalen
      n_var = data.var %>%
        group_by(col.index) %>%
        summarize(n=sum(n.unweighted, na.rm=T)) %>%
        full_join(data.frame(col.index=col.design$col.index, dummy=0), by="col.index") %>% # zorgen dat alle waardes aanwezig zijn, ook bij lege kolommen
        arrange(col.index) %>%
        mutate(n=coalesce(n, 0)) %>%
        select(-dummy) %>%
        column_to_rownames("col.index") %>%
        t() %>%
        as.data.frame() %>%
        mutate(label="Aantal respondenten", sign=NA, .before=1)
      
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
        vals = as.character(data.var$val[data.var$col.index == col.design$col.index[j] & data.var$val %in% rownames(output)])
        output[vals,j] = data.var$perc.weighted[data.var$col.index == col.design$col.index[j] & data.var$val %in% rownames(output)]
        if (!is.null(algemeen$weergave) && algemeen$weergave == "n")
          output[vals,j] = data.var$n[data.var$col.index == col.design$col.index[j] & data.var$val %in% rownames(output)]
        
        # waarden onder de afkapgrens vervangen
        output[which(output[,j] <= algemeen$afkapwaarde_antwoord),j] = A_TOOSMALL
        
        #PS:
        #Metingen die o.b.v te lage aantallen zijn vervangen 
        if (!is.na(indeling_rijen$verberg_crossings[i]) && !is.na(col.design$crossing[j])) {
          output[,j] = Q_MISSING
        } else if (sum(data.var$n.unweighted[data.var$col.index == col.design$col.index[j]], na.rm=T) == 0) {
          output[,j] = Q_MISSING
        } else if (sum(data.var$n.unweighted[data.var$col.index == col.design$col.index[j]], na.rm=T) < algemeen$min_observaties_per_vraag) {
          #Alle percentages wegstrepen als aantallen per groep te klein zijn.
          output[,j] <- Q_TOOSMALL
        } else if(any(data.var$n.unweighted[data.var$col.index == col.design$col.index[j]] < algemeen$min_observaties_per_antwoord, na.rm=T)) {
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
            data.col = data.var[data.var$col.index == col.design$col.index[j] & data.var$val %in% rownames(output),]
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
      
      # output herschrijven naar een bruikbaar formaat
      output = output %>% as.data.frame() %>% rownames_to_column("val") %>%
        mutate(label=sapply(val, function(v) var_labels$label[var_labels$var == indeling_rijen$inhoud[i] & var_labels$val == as.character(v)]),
               sign=sapply(val, function(v) list(sign$col.index[sign$val == v]))) %>%
        relocate(label, sign)
      
      # zijn er aparte wensen qua weergave van antwoordmogelijkheden?
      if (!is.na(indeling_rijen$waardes[i])) {
        desired_answers = str_split(indeling_rijen$waardes[i], fixed("|")) %>% unlist() %>% str_trim()
        desired_rownums = sapply(desired_answers, function (a) { return(which(output$val == a)) })
        output = output[desired_rownums,]
      }
      
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
        output = output %>%
          filter(val == 1) %>%
          mutate(label=var_labels$label[var_labels$var == indeling_rijen$inhoud[i] & var_labels$val == "var"], .after=val) %>%
          select(-val)
        
        if (indeling_rijen$type[i] != "nvar") {
          # aantal hoeft niet in de cache als dit geen nvar is; NULL maken
          n_var = NULL
        }
        
        if (!is.null(table.cache)) {
          table.cache = bind_rows(table.cache, output)
          if (!is.null(n.cache) && indeling_rijen$type[i] == "nvar") {
            msg("Let op! De variabele %s op rij %d heeft als type 'nvar', maar dit betreft een dichotome variabele in een lijst. Alleen het aantal respondenten van de eerste variabele wordt weergegeven.",
                indeling_rijen$inhoud[i], i, level=WARN)
          } else {
            n.cache = n_var
          }
        } else {
          table.cache = output
          n.cache = n_var
        }
      } else { # niet dichotoom
        output = output %>%
          select(-val)
        
        # moet er nog een vorige tabel geprint worden?
        if (!is.null(table.cache)) {
          table.output = c(table.output, paste0("<table>\r\n",
                                                "<caption>", question.cache, "</caption>\r\n",
                                                header.output, "\r\n",
                                                ifelse(!is.null(n.cache), BuildHtmlTableRows(n.cache, col.design, T), ""),
                                                perc.row.output, "\r\n",
                                                "<tbody>\r\n",
                                                BuildHtmlTableRows(table.cache, col.design),
                                                "</tbody>\r\n",
                                                "</table>\r\n<br />\r\n"))
          question.cache = NA
          table.cache = NULL
          n.cache = NULL
        }
        
        # tabel invoegen
        table.output = c(table.output, paste0("<h3 class=\"vraag\">", HTMLencode(var_labels$label[var_labels$var == indeling_rijen$inhoud[i] & var_labels$val == "var"]), "</h3>",
                                              "<table>\r\n",
                                              # titel van de vraag toevoegen
                                              "<caption>", var_labels$label[var_labels$var == indeling_rijen$inhoud[i] & var_labels$val == "var"], "</caption>\r\n",
                                              header.output, "\r\n",
                                              ifelse(indeling_rijen$type[i] == "nvar", BuildHtmlTableRows(n_var, col.design, T), ""),
                                              perc.row.output, "\r\n",
                                              "<tbody>\r\n",
                                              BuildHtmlTableRows(output, col.design),
                                              "</tbody>\r\n",
                                              "</table>\r\n<br />\r\n"))
      }
      
      # TODO: herschrijven
      # labels.oversized = which(str_length(output[,2]) > design("label_max_lengte"))
      # if (length(labels.oversized) > 0) {
      #   label.oversized.rows = c(label.oversized.rows, c + labels.oversized - 1)
      # }
    }
  }
  
  # als er nog iets in de tijdelijke opslag zit, printen
  if (!is.null(table.cache)) {
    table.output = c(table.output, paste0("<table>\r\n",
                                          "<caption>", question.cache, "</caption>\r\n",
                                          header.output, "\r\n",
                                          ifelse(!is.null(n.cache), BuildHtmlTableRows(n.cache, col.design), ""),
                                          perc.row.output,
                                          "<tbody>\r\n",
                                          BuildHtmlTableRows(table.cache, col.design),
                                          "</tbody>\r\n",
                                          "</table>\r\n<br />\r\n"))
    question.cache = NA
    table.cache = NULL
  }
  
  # laatste regel aanhouden als einde, zodat de kleuren niet doorlopen
  # c = c - 1
  # 
  # # algemene opmaak
  # setColWidths(wb, subset.name, cols=1, hidden=T) # 1e kolom verbergen, die is alleen voor eigen naslag
  # setColWidths(wb, subset.name, cols=2, width=design("kolombreedte_antwoorden")) # 2de kolom breder voor de tekstlabels
  # setColWidths(wb, subset.name, cols=3:n.col.total, width=design("kolombreedte")) # 3 tot nde kolom vaste breedte
  # setRowHeights(wb, subset.name, rows=table.start:c, heights=design("rij_hoogte")) # normale rijhoogte, ná de intro
  # 
  # if (length(label.oversized.rows) > 0) {
  #   setRowHeights(wb, subset.name, rows=label.oversized.rows, heights=design("rij_hoogte")*2)
  # }
  # 
  # addStyle(wb, subset.name, createStyle(wrapText=T), rows=table.start:c, cols=1:n.col.total, gridExpand=T, stack=T)
  # addStyle(wb, subset.name, createStyle(fgFill="#ffffff"), rows=which(!1:c %in% title.rows), cols=1:n.col.total, gridExpand=T, stack=T)
  # 
  # # afwisselende kleuren?
  # if (design("rijen_afwisselend_kleuren")) {
  #   # data.rows bevat de rijen met data
  #   # als het verschil tussen twee waardes meer dan 1 is gaat het om een nieuwe categorie, dus daar steeds opnieuw beginnen met tellen
  #   data.rows.diff = data.rows - lag(data.rows)
  #   data.rows.blocks = which(data.rows.diff > 1)
  #   
  #   data.rows.interval = c()
  #   for (i in 1:length(data.rows.blocks)) {
  #     end = ifelse(i < length(data.rows.blocks), data.rows[data.rows.blocks[i + 1]-1], data.rows[length(data.rows)])
  #     if (data.rows[data.rows.blocks[i]] > end) next
  #     data.rows.interval = c(data.rows.interval, seq(from=data.rows[data.rows.blocks[i]], to=end, by=2))
  #   }
  #   
  #   addStyle(wb, subset.name, style.gray.bg, rows=data.rows.interval, cols=1:n.col.total, gridExpand=T, stack=T)
  # }
  # if (design("kolommen_afwisselend_kleuren")) {
  #   addStyle(wb, subset.name, style.gray.bg, rows=1:c, cols=seq(from=3, to=n.col.total, by=2), gridExpand=T, stack=T)
  # }
  # if (design("kolommen_crossings_kleuren")) {
  #   col.design = col.design %>% group_by(dataset, subset, year, crossing)
  #   cols = group_rows(col.design)[order(sapply(group_rows(col.design),'[[',1))]
  #   gray = T
  #   for (i in cols) {
  #     if (gray) {
  #       addStyle(wb, subset.name, style.gray.bg, rows=c(data.rows, perc.rows, header.col.rows + ifelse(design("header_stijl") == "enkel", 0, 1)),
  #                cols=2+i, gridExpand=T, stack=T)
  #     }
  #     gray = !gray
  #   }
  # }
  # 
  # # TODO: borders om de tabellen?
  # 
  # # plaatjes toevoegen?
  # if (nrow(logos) > 0) {
  #   for (i in 1:nrow(logos)) {
  #     rij = logos$rij[i]
  #     kolom = logos$kolom[i]
  #     # bij negatieve waarden doen we aantal - waarde
  #     if (rij < 0) {
  #       rij = c - abs(rij)
  #     }
  #     if (kolom < 0) {
  #       kolom = n.col.total - abs(kolom) + 1
  #     }
  #     
  #     insertImage(wb, subset.name, file=logos$bestand[i], width=logos$breedte[i], height=logos$hoogte[i], units="px",
  #                 startRow=rij, startCol=kolom)
  #   }
  # }

  # tabellen toevoegen en witregel aan het einde
  table.output = c(table.output, "<br>\r\n")
  template = str_replace(template, fixed("{tabellen}"), str_c(table.output, collapse="\r\n"))
  
  # en als laatste... opslaan!
  tryCatch({
    #saveWorkbook(wb, sprintf("output/%s.xlsx", subset.name), overwrite=T)
    cat(template, file=sprintf("output/%s.html", str_replace(subset.name, "\\\\|/", " ")))
    msg("Digitoegankelijk tabellenboek voor %s opgeslagen.", subset.name, level=MSG)
  },
  error = function(e){
    msg("Er is een fout opgetreden bij het maken van het digitoegankelijke tabellenboek: %s", e$message, level=MSG)
  },
  warning = function(e){
    msg("Er is een fout opgetreden bij het maken van het digitoegankelijke tabellenboek: %s", e$message, level=MSG)
  })
  
}
