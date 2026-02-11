#
#
# Deze functie doet de daadwerkelijke berekeningen voor het tabellenboek.
# Op basis van een 'kolomontwerp' worden getallen berekend middels het survey-package
# en de benodigde p-waardes berekend met svychisq(). Vergelijkingen zijn alleen mogelijk
# binnen crossings of tussen totalen. (Het vergelijken van een crossing met een totaal
# zou de boel een stuk complexer maken en dit lijkt ons in de praktijk niet relevant.)
#
#

# functie om resultatentabellen met verschillende lengtes te corrigeren naar het juiste formaat
# dit kan bijvoorbeeld voorkomen als er meerdere antwoordmogelijkheden zijn, maar deze in één (of meer) van de datasets niet allemaal voorkomen
# bijv.:
# dataset 1:      dataset 2:
# 1 - 13          1 - 10
# 2 - 15          3 - 13
# 3 - 16          4 - 8
# hieruit willen we een tabel met 4 rijen waar de resultaten op de goeie locatie staan
# dit is vooral relevant omdat survey wel alle mogelijkheden weergeeft, maar table() alleen als deze ook in de dataset zitten
# we nemen daarom de uitslag van survey als basis, en vullen de ongewogen dan getallen met wat we wel weten
MatchTables = function (weighted, unweighted, is_2d=F) {
  weighted.corr = weighted
  unweighted.corr = unweighted
  
  if (is_2d) {
    if (!identical(rownames(weighted), rownames(unweighted)) || nrow(weighted) != nrow(unweighted) || !identical(colnames(weighted), colnames(unweighted)) || ncol(weighted) != ncol(unweighted)) {
      rownames = unique(c(rownames(weighted), rownames(unweighted)))
      colnames = unique(c(colnames(weighted), colnames(unweighted)))
      weighted.corr = matrix(0, nrow=length(rownames), ncol=length(colnames))
      rownames(weighted.corr) = rownames
      colnames(weighted.corr) = colnames
      weighted.corr[rownames(weighted),colnames(weighted)] = weighted
      
      unweighted.corr = matrix(0, nrow=length(rownames), ncol=length(colnames))
      rownames(unweighted.corr) = rownames
      colnames(unweighted.corr) = colnames
      # nu even creatief zijn:
      # - bij een missende kolom (= crossing), kijken welke kolommen WEL getallen bevatten; die zitten waarschijnlijk ook in survey
      # - bij een missende rij (= antwoordoptie), simpelweg vullen op naam
      if (nrow(weighted) != nrow(unweighted)) {
        unweighted.corr[rownames(unweighted),colnames(unweighted)] = unweighted
      } else if (ncol(weighted) != ncol(unweighted)) {
        unweighted.corr[rownames(unweighted),colnames(unweighted)[colSums(unweighted, na.rm=T) > 0]] = unweighted[,colSums(unweighted, na.rm=T) > 0]
      }
    }
  } else {
    if (!identical(names(weighted), names(unweighted)) || length(weighted) != length(unweighted)) {
      names = unique(c(names(weighted), names(unweighted)))
      weighted.corr = matrix(0, nrow=length(names), ncol=1)
      rownames(weighted.corr) = names
      weighted.corr[names(weighted),1] = weighted
      
      unweighted.corr = matrix(0, nrow=length(names), ncol=1)
      rownames(unweighted.corr) = names
      unweighted.corr[names(unweighted),1] = unweighted
    }
  }
  
  return(list(weighted.corr, unweighted.corr))
}

# col.design = kolom_opbouw
# s = 43
GetTableRow = function (var, design, col.design, subsetmatches) {
  msg("Variabele %s wordt uitgevoerd over %d kolommen.", var, nrow(col.design), level=MSG)
  
  results = data.frame()
  
  col.design = col.design %>% group_by(dataset, subset, year, crossing, test.col)
  colgroups = group_keys(col.design)
  
  # splitsen op subset?
  if (!is.null(subsetmatches)) {
    leadingsubset = colnames(subsetmatches)[1]
    leadingcol = min(col.design$col.index[which(col.design$subset == leadingsubset & is.na(col.design$crossing))])
    
    subsetvals = subsetmatches[,1]
    
    # kolommen met subset en zonder subset worden apart doorlopen om dubbele calls te voorkomen
    # (anders zou voor ieder niveau in de subset totaalkolommen van andere datasets opnieuw berekend worden - zonde)
    for (s in 1:length(subsetvals)) {
      # bestaat hier data voor?
      if (!(paste0("dummy._col", leadingcol, ".s.", unname(subsetvals[s])) %in% names(design$variables)))
        next
      
      msg("Variabele %s wordt uitgevoerd op subset %s met niveau %s (%s)", var, leadingsubset, as.character(unname(subsetvals[s])), names(subsetvals)[s], level=DEBUG)
      
      for (i in 1:nrow(colgroups)) {
        if (is.na(colgroups$subset[i])) next
        
        msg("Subset %d: dataset %d, subset %s, jaar %s, crossing %s", level=DEBUG,
            i, colgroups$dataset[i], colgroups$subset[i], colgroups$year[i], colgroups$crossing[i])
        
        if (!(paste0("dummy._col", leadingcol, ".s.", unname(subsetvals[s])) %in% names(design$variables)))
          next
        
        # er zijn nu twee scenario's:
        # 1) originele subsetvariabele (of dezelfde subsetvariabele in een andere dataset) -> niks doen, subsetvals[s] bevat de nodige informatie
        # 2) een extra subsetvariabele, gebaseerd op de leidende subsetvariabele -> zoek welke overkoepelende subset hierbij hoort en vul die in
        subsetval = subsetvals[s]
        if (colgroups$subset[i] != leadingsubset) {
          # vergelijken van significantie met een 'hoger gelegen' subset kan niet
          # stel bijvoorbeeld dat er een kolom is met gemeente, en een kolom met subregio, dan kan alleen gemeente vs. subregio, niet andersom
          if (!is.na(colgroups$test.col[i]) && (col.design$subset[colgroups$test.col[i]] == leadingsubset || col.design$subset[colgroups$test.col[i]] != colgroups$subset[i])) {
            msg("Let op! P-waardes kunnen alleen berekend worden van klein naar groot niveau. Een vergelijking van %s met %s (kolom %d) is daardoor niet mogelijk. Verander deze volgorde als significantie berekend moet worden.",
                colgroups$subset[i], col.design$subset[colgroups$test.col[i]], colgroups$test.col[i], level=ERR)
          }
          subsetval = subsetmatches[subsetmatches[,1] == subsetvals[s], colgroups$subset[i]]
        }
        
        # crossing?
        if (!is.na(colgroups$crossing[i])) {
          # dit betekent meerdere kolommen vullen, want crossing
          cols = col.design$col.index[group_rows(col.design)[[i]]]
          
          desired.cols = paste0("dummy._col", cols, ".s.", subsetval)
          if (!all(desired.cols %in% names(design$variables))) {
            msg("Er is geen data beschikbaar voor kolom %s bij dataset %d, subset %s, jaar %s, crossing %s, selectie %s.",
                str_c(cols[!(paste0("dummy._col", cols, ".s.", subsetval) %in% names(design$variables))], collapse=", "),
                colgroups$dataset[i], colgroups$subset[i], colgroups$year[i], colgroups$crossing[i], subsetval, level=WARN)
            desired.cols = desired.cols[desired.cols %in% names(design$variables)]
            if (length(desired.cols) == 0)
              next
          }
          
          selection = str_c(desired.cols, collapse=" | ")
          design.subset = subset(design, eval(parse(text=selection)))
          
          weighted.raw = svytable(formula=as.formula(paste0("~", var, "+", colgroups$crossing[i])),
                              design=design.subset)
          unweighted.raw = table(design.subset$variables[[var]], design.subset$variables[[colgroups$crossing[i]]])
          weighted = MatchTables(weighted.raw, unweighted.raw, T)[[1]]
          unweighted = MatchTables(weighted.raw, unweighted.raw, T)[[2]]
          n = length(weighted)
          
          pvals = matrix(NA, nrow=nrow(weighted), ncol=ncol(weighted))
          rownames(pvals) = rownames(weighted)
   
          if (!is.na(colgroups$test.col[i]) && colgroups$test.col[i] == 0) {
            if (min(dim(weighted)) < 2) {
              msg("Bij variabele %s met crossing %s werd maar één rij/kolom in de kruistabel gevonden (dimensies %s). Hierdoor kan geen chi2-test worden uitgevoerd. Controleer de data.",
                  var, colgroups$crossing[i], str_c(dim(weighted), collapse="x"), level=WARN)
            } else {
              answers = rownames(weighted)
              for (answer in answers) {
                tryCatch({
                  test = svychisq(formula=as.formula(paste0("~dummy.", var, ".", answer, "+", colgroups$crossing[i])), design=design.subset)
                  if(is.nan(test$p.value) & algemeen$benader_chisq){
                    test = svychisq(formula=as.formula(paste0("~dummy.", var, ".", answer, "+", colgroups$crossing[i])), design=design.subset, statistic="Chisq")
                    msg("Bij variabele %s met antwoord %s (%s) kon de p-waarde niet met worden berekend met F (Rao–Scott second order).  Nu schatting op basis van Chisq (Rao–Scott first order).",
                        var, answer, var_labels$label[var_labels$var == var & var_labels$val == answer], level=WARN)
                    
                  }
                  pvals[answer,] = rep(test$p.value, ncol(pvals))
                },
                error=function (e) msg("Bij variabele %s met antwoord %s (%s) kon geen p-waarde worden berekend voor crossing %s in subset %s. Foutmelding: %s",
                                       var, answer, var_labels$label[var_labels$var == var & var_labels$val == answer], colgroups$crossing[i], subsetval,
                                       e, level=WARN))
              }
            }
          } else if (!is.na(colgroups$test.col[i])) {
            # tweede kolom includeren om te testen; dit leidt tot een vector van selecties
            selection = paste0(desired.cols, " | dummy._col", colgroups$test.col[i])
            # subset nodig voor testkolom?
            if (!is.na(col.design$subset[colgroups$test.col[i]])) {
              if (col.design$subset[colgroups$test.col[i]] == leadingsubset) {
                selection = paste0(selection, ".s.", subsetvals[s])
              } else {
                selection = paste0(selection, ".s.", subsetmatches[subsetmatches[,1] == subsetvals[s], col.design$subset[colgroups$test.col[i]]])
              }
            }
            for (j in 1:length(selection)) {
              select = selection[j]
              design.subset = subset(design, eval(parse(text=select)))
              source.col = str_split(select, fixed(" | "))[[1]][1]
              
              answers = rownames(weighted)
              for (answer in answers) {
                tablecounts = table(design.subset$variables[[paste0("dummy.", var, ".", answer)]], design.subset$variables[[source.col]])
                if (any(colSums(tablecounts) == 0)) {
                  # er is geen data in één van beide kolommen; p-waarde berekenen is zinloos
                  msg("Bij variabele %s met antwoord %s (%s) kon geen p-waarde worden berekend in vergelijking met kolom %d; één van beide kolommen is leeg.",
                      var, answer, var_labels$label[var_labels$var == var & var_labels$val == answer], colgroups$test.col[i], level=WARN)
                  next
                }
                tryCatch({
                  test = svychisq(formula=as.formula(paste0("~dummy.", var, ".", answer, "+", source.col)),
                                  design=design.subset)
                  if(is.nan(test$p.value) & algemeen$benader_chisq){
                    test = svychisq(formula=as.formula(paste0("~dummy.", var, ".", answer, "+", source.col)),
                                    design=design.subset, statistic = "Chisq")
                    msg("Bij variabele %s met antwoord %s (%s) kon de p-waarde niet met worden berekend met F (Rao–Scott second order).  Nu schatting op basis van Chisq (Rao–Scott first order).",
                        var, answer, var_labels$label[var_labels$var == var & var_labels$val == answer], level=WARN)
                  }
                  pvals[answer, j] = test$p.value
                },
                error=function (e) msg("Bij variabele %s met antwoord %s (%s) kon geen p-waarde worden berekend. Foutmelding: %s",
                                       var, answer, var_labels$label[var_labels$var == var & var_labels$val == answer],
                                       e, level=WARN))
              }
            }
          }
          
          
          # om de resultaten in een dataframe te krijgen moeten ze door as.numeric() en unname()
          # dit zorgt ervoor dat de structuur verloren gaat; de kolommen worden onder elkaar geplaatst
          # we moeten dus de waardes indelen als colnames[1] * nrow, colnames[2] * nrow, etc.
          vals = as.vector(sapply(colnames(weighted), function (x, nrows) return(rep(x, nrows)), nrows=nrow(weighted)))
          
          results = bind_rows(results, data.frame(dataset=rep(colgroups$dataset[i], n), subset=rep(colgroups$subset[i], n), subset.val=rep(subsetval, n),
                                                  year=rep(colgroups$year[i], n),
                                                  crossing=rep(colgroups$crossing[i], n), crossing.val=vals,
                                                  var=rep(var, n), val=rownames(weighted),
                                                  sign=as.numeric(unname(pvals)), sign.vs=rep(colgroups$test.col[i], n),
                                                  n.weighted=as.numeric(unname(weighted)), perc.weighted=as.numeric(unname(proportions(weighted, margin=2)))*100,
                                                  n.unweighted=as.numeric(unname(unweighted)), perc.unweighted=as.numeric(unname(proportions(unweighted, margin=2)))*100))
        } else {
          # geen crossing; totaal
          col = col.design$col.index[group_rows(col.design)[[i]]]
          
          if (!(paste0("dummy._col", col, ".s.", subsetval) %in% names(design$variables))) {
            msg("Er is geen data beschikbaar voor kolom %d bij dataset %d, subset %s, jaar %s, selectie %s.",
                col, colgroups$dataset[i], colgroups$subset[i], colgroups$year[i], subsetval, level=WARN)
            next
          }
          
          weighted = svytable(formula=as.formula(paste0("~", var, "+dummy._col", col, ".s.", subsetval)), design=design)
          if ("TRUE" %in% colnames(weighted)) {
            if(nrow(weighted) == 1){
              answer_name <- rownames(weighted)
              weighted = weighted[,"TRUE"]
              names(weighted) <- answer_name
            } else {
              weighted = weighted[,"TRUE"]
            }
          } else {
            names = rownames(weighted)
            weighted = rep(NA, nrow(weighted))
            names(weighted) = names
          }
          unweighted = MatchTables(weighted, table(design$variables[[var]][design$variables[,paste0("dummy._col", col, ".s.", subsetval)]]))[[2]]
          if (sum(unweighted, na.rm=T) <= 0) unweighted = rep(NA, length(weighted))
          
          n = length(weighted)
          
          pvals = rep(NA, n)
          names(pvals) = names(weighted)
          
          # significantie berekenen?
          if (!is.na(colgroups$test.col[i]) && sum(unweighted, na.rm=T) > 0) {
            # tweede kolom includeren
            selection = paste0("dummy._col", col, ".s.", subsetval, " | dummy._col", colgroups$test.col[i])
            # subset nodig voor testkolom?
            if (!is.na(col.design$subset[colgroups$test.col[i]])) {
              if (col.design$subset[colgroups$test.col[i]] == leadingsubset) {
                selection = paste0(selection, ".s.", subsetvals[s])
              } else {
                selection = paste0(selection, ".s.", subsetmatches[subsetmatches[,1] == subsetvals[s], col.design$subset[colgroups$test.col[i]]])
              }
            }
            design.subset = subset(design, eval(parse(text=selection)))
            
            answers = names(weighted)
            for (answer in answers) {
              tablecounts = table(design.subset$variables[[paste0("dummy.", var, ".", answer)]], design.subset$variables[[paste0("dummy._col", col, ".s.", subsetval)]])
              if (any(colSums(tablecounts) == 0)) {
                # er is geen data in één van beide kolommen; p-waarde berekenen is zinloos
                msg("Bij variabele %s met antwoord %s (%s) kon geen p-waarde worden berekend in vergelijking met kolom %d; één van beide kolommen is leeg.",
                    var, answer, var_labels$label[var_labels$var == var & var_labels$val == answer], colgroups$test.col[i], level=WARN)
                next
              }
              tryCatch({
                test = svychisq(formula=as.formula(paste0("~dummy.", var, ".", answer, "+dummy._col", col, ".s.", subsetval)),
                                design=design.subset)
                if(is.nan(test$p.value) & algemeen$benader_chisq){
                  test = svychisq(formula=as.formula(paste0("~dummy.", var, ".", answer, "+dummy._col", col, ".s.", subsetval)),
                                  design=design.subset, statistic = "Chisq")
                  msg("Bij variabele %s met antwoord %s (%s) kon de p-waarde niet met worden berekend met F (Rao–Scott second order).  Nu schatting op basis van Chisq (Rao–Scott first order).",
                      var, answer, var_labels$label[var_labels$var == var & var_labels$val == answer], level=WARN)
                }
                pvals[answer] = test$p.value
              },
              error=function (e) msg("Bij variabele %s met antwoord %s (%s) kon geen p-waarde worden berekend. Foutmelding: %s",
                                     var, answer, var_labels$label[var_labels$var == var & var_labels$val == answer],
                                     e, level=WARN))
            }
          }
          
          results = bind_rows(results, data.frame(dataset=rep(col.design$dataset[col], n), subset=rep(col.design$subset[col], n), subset.val=rep(subsetval, n),
                                                  year=rep(col.design$year[col], n),
                                                  crossing=rep(col.design$crossing[col], n), crossing.val=as.character(rep(col.design$crossing.val[col], n)),
                                                  var=rep(var, n), val=names(weighted),
                                                  sign=pvals, sign.vs=rep(col.design$test.col[col], n),
                                                  n.weighted=as.numeric(unname(weighted)), perc.weighted=proportions(as.numeric(unname(weighted)))*100,
                                                  n.unweighted=as.numeric(unname(unweighted)), perc.unweighted=proportions(as.numeric(unname(unweighted)))*100))
        }
      }
    }
  }
  
  # nu alle kolommen zonder subset
  for (i in 1:nrow(colgroups)) {
    if (!is.na(colgroups$subset[i]))
      next
    
    msg("Subset %d: dataset %d, subset %s, jaar %s, crossing %s", level=DEBUG,
        i, colgroups$dataset[i], colgroups$subset[i], colgroups$year[i], colgroups$crossing[i])
    
    # crossing?
    if (!is.na(colgroups$crossing[i])) {
      # dit betekent meerdere kolommen vullen, want crossing
      cols = col.design$col.index[group_rows(col.design)[[i]]]
      selection = str_c(paste0("dummy._col", cols), collapse=" | ")
      design.subset = subset(design, eval(parse(text=selection)))
      
      weighted.raw = svytable(formula=as.formula(paste0("~", var, "+", colgroups$crossing[i])),
                          design=design.subset)
      unweighted.raw = table(design.subset$variables[[var]], design.subset$variables[[colgroups$crossing[i]]])
      weighted = MatchTables(weighted.raw, unweighted.raw, T)[[1]]
      unweighted = MatchTables(weighted.raw, unweighted.raw, T)[[2]]
      n = length(weighted)
      
      pvals = matrix(NA, nrow=nrow(weighted), ncol=ncol(weighted))
      rownames(pvals) = rownames(weighted)
      if (!is.na(colgroups$test.col[i]) && colgroups$test.col[i] == 0) {
        if (min(dim(weighted)) < 2) {
          msg("Bij variabele %s met crossing %s werd maar één rij/kolom in de kruistabel gevonden (dimensies %s). Hierdoor kan geen chi2-test worden uitgevoerd. Controleer de data.",
              var, colgroups$crossing[i], str_c(dim(weighted), collapse="x"), level=WARN)
        } else {
          answers = rownames(weighted)
          for (answer in answers) {
            tryCatch({
              test = svychisq(formula=as.formula(paste0("~dummy.", var, ".", answer, "+", colgroups$crossing[i])), design=design.subset)
              if(is.nan(test$p.value) & algemeen$benader_chisq){
                test = svychisq(formula=as.formula(paste0("~dummy.", var, ".", answer, "+", colgroups$crossing[i])), design=design.subset, statistic="Chisq")
                msg("Bij variabele %s met antwoord %s (%s) kon de p-waarde niet met worden berekend met F (Rao–Scott second order).  Nu schatting op basis van Chisq (Rao–Scott first order).",
                    var, answer, var_labels$label[var_labels$var == var & var_labels$val == answer], level=WARN)
              }
              pvals[answer,] = rep(test$p.value, ncol(pvals))
            },
            error=function (e) msg("Bij variabele %s met antwoord %s (%s) kon geen p-waarde worden berekend voor crossing %s. Foutmelding: %s",
                                   var, answer, var_labels$label[var_labels$var == var & var_labels$val == answer], colgroups$crossing[i],
                                   e, level=WARN))
          }
        }
      } else if (!is.na(colgroups$test.col[i])) {
        # let op: testen vanuit een kolom zonder subset naar een kolom MET subset gaat niet
        if (!is.na(col.design$subset[colgroups$test.col[i]])) {
          msg("Let op! In kolom %d (zonder subset) wordt vergeleken met kolom %d (met subset). Dit is niet mogelijk met de opbouw van de code. Als deze verschillen inzichtelijk gemaakt moeten worden moet de berekening andersom worden gezet: kolom MET subset vs. kolom ZONDER subset.",
              col, colgroups$test.col[i], level=WARN)
        } else {
          # tweede kolom includeren om te testen; dit leidt tot een vector van selecties
          selection = paste0("dummy._col", cols, " | dummy._col", colgroups$test.col[i])
          
          for (j in 1:length(selection)) {
            select = selection[j]
            design.subset = subset(design, eval(parse(text=select)))
            source.col = str_split(select, fixed(" | "))[[1]][1]
            
            answers = rownames(weighted)
            for (answer in answers) {
              tablecounts = table(design.subset$variables[[paste0("dummy.", var, ".", answer)]], design.subset$variables[[source.col]])
              if (any(colSums(tablecounts) == 0)) {
                # er is geen data in één van beide kolommen; p-waarde berekenen is zinloos
                msg("Bij variabele %s met antwoord %s (%s) kon geen p-waarde worden berekend in vergelijking met kolom %d; één van beide kolommen is leeg.",
                    var, answer, var_labels$label[var_labels$var == var & var_labels$val == answer], colgroups$test.col[i], level=WARN)
                next
              }
              tryCatch({
                test = svychisq(formula=as.formula(paste0("~dummy.", var, ".", answer, "+", source.col)),
                                design=design.subset)
                if(is.nan(test$p.value) & algemeen$benader_chisq){
                  test = svychisq(formula=as.formula(paste0("~dummy.", var, ".", answer, "+", source.col)),
                                  design=design.subset, statistic="Chisq")
                  msg("Bij variabele %s met antwoord %s (%s) kon de p-waarde niet met worden berekend met F (Rao–Scott second order).  Nu schatting op basis van Chisq (Rao–Scott first order).",
                      var, answer, var_labels$label[var_labels$var == var & var_labels$val == answer], level=WARN)
                  
                }
                pvals[answer, j] = test$p.value
              },
              error=function (e) msg("Bij variabele %s met antwoord %s (%s) kon geen p-waarde worden berekend. Foutmelding: %s",
                                     var, answer, var_labels$label[var_labels$var == var & var_labels$val == answer],
                                     e, level=WARN))
            }
          }
        }
      }
      
      # om de resultaten in een dataframe te krijgen moeten ze door as.numeric() en unname()
      # dit zorgt ervoor dat de structuur verloren gaat; de kolommen worden onder elkaar geplaatst
      # we moeten dus de waardes indelen als colnames[1] * nrow, colnames[2] * nrow, etc.
      vals = as.vector(sapply(colnames(weighted), function (x, nrows) return(rep(x, nrows)), nrows=nrow(weighted)))
      
      results = bind_rows(results, data.frame(dataset=rep(colgroups$dataset[i], n), subset=rep(colgroups$subset[i], n), subset.val=rep(NA, n),
                                              year=rep(colgroups$year[i], n),
                                              crossing=rep(colgroups$crossing[i], n), crossing.val=vals,
                                              var=rep(var, n), val=rownames(weighted),
                                              sign=as.numeric(unname(pvals)), sign.vs=rep(colgroups$test.col[i], n),
                                              n.weighted=as.numeric(unname(weighted)), perc.weighted=as.numeric(unname(proportions(weighted, margin=2)))*100,
                                              n.unweighted=as.numeric(unname(unweighted)), perc.unweighted=as.numeric(unname(proportions(unweighted, margin=2)))*100))
    } else {
      # geen crossing; totaal
      
      col = col.design$col.index[group_rows(col.design)[[i]]]
      
      weighted = svytable(formula=as.formula(paste0("~", var, "+dummy._col", col)), design=design)
      if ("TRUE" %in% colnames(weighted)) {
        if(nrow(weighted) == 1){
          answer_name <- rownames(weighted)
          weighted = weighted[,"TRUE"]
          names(weighted) <- answer_name
        } else {
          weighted = weighted[,"TRUE"]
        }
      } else {
        names = rownames(weighted)
        weighted = rep(NA, nrow(weighted))
        names(weighted) = names
      }
      unweighted = MatchTables(weighted, table(design$variables[[var]][design$variables[,paste0("dummy._col", col)]]))[[2]]
      if (sum(unweighted, na.rm=T) <= 0) unweighted = rep(NA, length(weighted))
      n = length(weighted)
      
      pvals = rep(NA, n)
      names(pvals) = names(weighted)
      
      # significantie berekenen?
      if (!is.na(colgroups$test.col[i]) && sum(unweighted, na.rm=T) > 0) {
        # tweede kolom includeren
        # let op: testen vanuit een kolom zonder subset naar een kolom MET subset gaat niet
        if (!is.na(col.design$subset[colgroups$test.col[i]])) {
          msg("Let op! In kolom %d (zonder subset) wordt vergeleken met kolom %d (met subset). Dit is niet mogelijk met de opbouw van de code. Als deze verschillen inzichtelijk gemaakt moeten worden moet de berekening andersom worden gezet: kolom MET subset vs. kolom ZONDER subset.",
              col, colgroups$test.col[i], level=WARN)
        } else {
          selection = str_c(paste0("dummy._col", c(col, colgroups$test.col[i])), collapse=" | ")
          design.subset = subset(design, eval(parse(text=selection)))
          
          answers = names(weighted)
          for (answer in answers) {
            tablecounts = table(design.subset$variables[[paste0("dummy.", var, ".", answer)]], design.subset$variables[[paste0("dummy._col", col)]])
            if (any(colSums(tablecounts) == 0)) {
              # er is geen data in één van beide kolommen; p-waarde berekenen is zinloos
              msg("Bij variabele %s met antwoord %s (%s) kon geen p-waarde worden berekend in vergelijking met kolom %d; één van beide kolommen is leeg.",
                  var, answer, var_labels$label[var_labels$var == var & var_labels$val == answer], colgroups$test.col[i], level=WARN)
              next
            }
            tryCatch({
              test = svychisq(formula=as.formula(paste0("~dummy.", var, ".", answer, "+dummy._col", col)),
                              design=design.subset)
              if(is.nan(test$p.value) & algemeen$benader_chisq){
                test = svychisq(formula=as.formula(paste0("~dummy.", var, ".", answer, "+dummy._col", col)),
                                design=design.subset, statistic="Chisq")
                msg("Bij variabele %s met antwoord %s (%s) kon de p-waarde niet met worden berekend met F (Rao–Scott second order).  Nu schatting op basis van Chisq (Rao–Scott first order).",
                    var, answer, var_labels$label[var_labels$var == var & var_labels$val == answer], level=WARN)
              }
              pvals[answer] = test$p.value
            },
            error=function (e) msg("Bij variabele %s met antwoord %s (%s) kon geen p-waarde worden berekend. Foutmelding: %s",
                                   var, answer, var_labels$label[var_labels$var == var & var_labels$val == answer],
                                   e, level=WARN))
          }
        }
      }
      
      results = bind_rows(results, data.frame(dataset=rep(col.design$dataset[col], n), subset=rep(col.design$subset[col], n), subset.val=rep(NA, n),
                                              year=rep(col.design$year[col], n),
                                              crossing=rep(col.design$crossing[col], n), crossing.val=as.character(rep(col.design$crossing.val[col], n)),
                                              var=rep(var, n), val=names(weighted),
                                              sign=pvals, sign.vs=rep(col.design$test.col[col], n),
                                              n.weighted=as.numeric(unname(weighted)), perc.weighted=proportions(as.numeric(unname(weighted)))*100,
                                              n.unweighted=as.numeric(unname(unweighted)), perc.unweighted=proportions(as.numeric(unname(unweighted)))*100))
    }
  }
  
  
  return(results)
}