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
MatchTables = function (weighted, unweighted, is_2d=F) {
  weighted.corr = weighted
  unweighted.corr = unweighted
  
  if (is_2d) {
    if (!identical(rownames(weighted), rownames(unweighted)) || nrow(weighted) != nrow(unweighted)) {
      rownames = unique(c(rownames(weighted), rownames(unweighted)))
      colnames = unique(c(colnames(weighted), colnames(unweighted)))
      weighted.corr = matrix(NA, nrow=length(rownames), ncol=length(colnames))
      rownames(weighted.corr) = rownames
      colnames(weighted.corr) = colnames
      weighted.corr[rownames(weighted),colnames(weighted)] = weighted
      
      unweighted.corr = matrix(NA, nrow=length(rownames), ncol=length(colnames))
      rownames(unweighted.corr) = rownames
      colnames(unweighted.corr) = colnames
      unweighted.corr[rownames(unweighted),colnames(unweighted)] = unweighted
    }
  } else {
    if (!identical(names(weighted), names(unweighted)) || length(weighted) != length(unweighted)) {
      names = unique(c(names(weighted), names(unweighted)))
      weighted.corr = matrix(NA, nrow=length(names), ncol=1)
      rownames(weighted.corr) = names
      weighted.corr[names(weighted),1] = weighted
      
      unweighted.corr = matrix(NA, nrow=length(names), ncol=1)
      rownames(unweighted.corr) = names
      unweighted.corr[names(unweighted),1] = unweighted
    }
  }
  
  return(list(weighted.corr, unweighted.corr))
}

# TODO: confidence intervals toevoegen
# col.design = kolom_opbouw
GetTableRow = function (var, design, calculate.ci=T, col.design, subsetmatches) {
  msg("Variabele %s wordt uitgevoerd over %d kolommen.", var, nrow(col.design), level=MSG)
  
  results = data.frame()
  
  col.design = col.design %>% group_by(dataset, subset, year, crossing, test.col)
  colgroups = group_keys(col.design)
  
  # splitsen op subset?
  if (!is.null(subsetmatches)) {
    leadingsubset = colnames(subsetmatches)[1]
    leadingcol = min(col.design$col.index[which(!is.na(col.design$subset))[1]])
    
    subsetvals = subsetmatches[,1]
    
    # kolommen met subset en zonder subset worden apart doorlopen om dubbele calls te voorkomen
    # (anders zou voor ieder niveau in de subset totaalkolommen van andere datasets opnieuw berekend worden - zonde)
    for (s in 1:length(subsetvals)) {
      # bestaat hier data voor?
      if (!(paste0("dummy._col", leadingcol, ".s.", unname(subsetvals[s])) %in% names(design$variables)))
        next
      
      msg("Variabele %s wordt uitgevoerd op subset %s met niveau %d (%s)", var, leadingsubset, unname(subsetvals[s]), names(subsetvals)[s], level=DEBUG)
      
      for (i in 1:nrow(colgroups)) {
        if (is.na(colgroups$subset[i])) next
        
        msg("Subset %d: dataset %d, subset %s, jaar %s, crossing %s", level=DEBUG,
            i, colgroups$dataset[i], colgroups$subset[i], colgroups$year[i], colgroups$crossing[i])
        
        # er zijn nu twee scenario's:
        # 1) originele subsetvariabele (of dezelfde subsetvariabele in een andere dataset) -> niks doen, subsetvals[s] bevat de nodige informatie
        # 2) een extra subsetvariabele, gebaseerd op de leidende subsetvariabele -> zoek welke overkoepelende subset hierbij hoort en vul die in
        subsetval = subsetvals[s]
        if (colgroups$subset[i] != leadingsubset) {
          subsetval = subsetmatches[subsetmatches[,1] == subsetvals[s], colgroups$subset[i]]
        }
        
        # crossing?
        if (!is.na(colgroups$crossing[i]) && !is.na(colgroups$test.col[i]) && colgroups$test.col[i] == 0) {
          # dit betekent meerdere kolommen vullen, want crossing
          cols = col.design$col.index[group_rows(col.design)[[i]]]
          selection = str_c(paste0("dummy._col", cols, ".s.", subsetval), collapse=" | ")
          design.subset = subset(design, eval(parse(text=selection)))
          
          weighted = svytable(formula=as.formula(paste0("~", var, "+", colgroups$crossing[i])),
                              design=design.subset)
          unweighted = MatchTables(weighted, table(design.subset$variables[[var]], design.subset$variables[[colgroups$crossing[i]]]), T)[[2]]
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
          
          weighted = svytable(formula=as.formula(paste0("~", var, "+dummy._col", col, ".s.", subsetval)), design=design)
          if (ncol(weighted) == 2 && colnames(weighted)[2] == "TRUE") {
            weighted = weighted[,2]
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
            if (!is.na(colgroups$subset[colgroups$test.col[i]])) {
              if (colgroups$subset[colgroups$test.col[i]] == leadingsubset) {
                selection = paste0(selection, ".s.", subsetvals[s])
              } else {
                selection = paste0(selection, ".s.", subsetmatches[subsetmatches[,1] == subsetvals[s], colgroups$subset[colgroups$test.col[i]]])
              }
            }
            design.subset = subset(design, eval(parse(text=selection)))
            
            answers = names(weighted)
            for (answer in answers) {
              test = svychisq(formula=as.formula(paste0("~dummy.", var, ".", answer, "+dummy._col", col)),
                              design=design.subset)
              pvals[answer] = test$p.value
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
    
    # crossing? zo ja, dan kunnen we dit behandelen als totaal indien er NIET intern vergeleken wordt (binnen de else)
    # anders wel behandelen als crossing (dus binnen de if)
    if (!is.na(colgroups$crossing[i]) && !is.na(colgroups$test.col[i]) && colgroups$test.col[i] == 0) {
      # dit betekent meerdere kolommen vullen, want crossing
      cols = col.design$col.index[group_rows(col.design)[[i]]]
      selection = str_c(paste0("dummy._col", cols), collapse=" | ")
      design.subset = subset(design, eval(parse(text=selection)))
      
      weighted = svytable(formula=as.formula(paste0("~", var, "+", colgroups$crossing[i])),
                          design=design.subset)
      unweighted = MatchTables(weighted, table(design.subset$variables[[var]], design.subset$variables[[colgroups$crossing[i]]]), T)[[2]]
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
      if (ncol(weighted) == 2 && colnames(weighted)[2] == "TRUE") {
        weighted = weighted[,2]
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
        selection = str_c(paste0("dummy._col", c(col, colgroups$test.col[i])), collapse=" | ")
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
                                              var=rep(var, n), val=names(weighted),
                                              sign=pvals, sign.vs=rep(col.design$test.col[col], n),
                                              n.weighted=as.numeric(unname(weighted)), perc.weighted=proportions(as.numeric(unname(weighted)))*100,
                                              n.unweighted=as.numeric(unname(unweighted)), perc.unweighted=proportions(as.numeric(unname(unweighted)))*100))
    }
  }
  
  
  return(results)
}