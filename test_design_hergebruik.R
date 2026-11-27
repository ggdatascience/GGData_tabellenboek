
## geen hergebruik svydesign

t.start = proc.time()["elapsed"]
results = data.frame()
for (i in 1:5) {
  var = varlist$inhoud[i]
  
  vars_to_select = unique(c(persistent_design_cols, var, colnames(data)[str_starts(colnames(data), paste0("dummy.", var))]))
  data.tmp = data %>% select(any_of(vars_to_select))
  
  for (j in 1:nrow(weegfactorvars)) {
    wfname = weegfactorvars[j, 1]
    wfdataset = weegfactorvars[j, 2]
    
    if (wfname == "weegfactor" && !is.na(varlist$weegfactor[i])) {
      data.tmp$superweegfactor = data[[varlist$weegfactor[i]]]
    } else if (!is.na(varlist[[wfname]][i]) && str_detect(wfdataset, "^\\.d(\\d+)$")) {
      # numeriek, dus d[getal]
      dataset = as.numeric(str_match(wfdataset, "(\\d+)")[, 2])
      data.tmp$superweegfactor[data.tmp$tbl_dataset == dataset] = data[[varlist[[wfname]][i]]][data$tbl_dataset == dataset]
    } else if (!is.na(varlist[[wfname]][i]) && str_detect(wfdataset, "^\\.d_(.+)")) {
      # naam van een dataset
      dataset = str_match(wfdataset, "^\\.d_(.+)")[, 2]
      if (!dataset %in% datasets$naam_dataset) {
        msg("Er is een weegfactor opgegeven in indeling_rijen voor dataset %s, maar deze dataset is niet bekend. Controleer de configuratie.",
            dataset, level = ERR)
      }
      dataset = which(datasets$naam_dataset == dataset)
      data.tmp$superweegfactor[data.tmp$tbl_dataset == dataset] = data[[varlist[[wfname]][i]]][data$tbl_dataset == dataset]
    } else if (!is.na(varlist[[wfname]][i])) {
      msg("Onbekende weegfactordefinitie opgegeven: %s. Controleer de configuratie.",
          wfname, level = ERR)
    }
  }
  
  current_design = svydesign(
    ids = ~1,
    strata = ~superstrata,
    weights = ~superweegfactor,
    data = data.tmp
  )
  
  results = bind_rows(results, GetTableRow(var, current_design, kolom_opbouw, subsetmatches))
}
t.after = proc.time()["elapsed"]

msg("Totale rekentijd zonder recycling %0.2f sec voor %d variabelen.", (t.after - t.start), 5, level = MSG)


## wel hergebruik
t.start = proc.time()["elapsed"]
#results = data.frame()
for (i in 1:5) {
  var = varlist$inhoud[i]
  
  vars_to_select = unique(c(var, colnames(data)[str_starts(colnames(data), paste0("dummy.", var))]))
  superweegfactor = data$superweegfactor
  
  for (j in 1:nrow(weegfactorvars)) {
    wfname = weegfactorvars[j, 1]
    wfdataset = weegfactorvars[j, 2]
    
    if (wfname == "weegfactor" && !is.na(varlist$weegfactor[i])) {
      superweegfactor = data[[varlist$weegfactor[i]]]
    } else if (!is.na(varlist[[wfname]][i]) && str_detect(wfdataset, "^\\.d(\\d+)$")) {
      # numeriek, dus d[getal]
      dataset = as.numeric(str_match(wfdataset, "(\\d+)")[, 2])
      superweegfactor[data$tbl_dataset == dataset] = data[[varlist[[wfname]][i]]][data$tbl_dataset == dataset]
    } else if (!is.na(varlist[[wfname]][i]) && str_detect(wfdataset, "^\\.d_(.+)")) {
      # naam van een dataset
      dataset = str_match(wfdataset, "^\\.d_(.+)")[, 2]
      if (!dataset %in% datasets$naam_dataset) {
        msg("Er is een weegfactor opgegeven in indeling_rijen voor dataset %s, maar deze dataset is niet bekend. Controleer de configuratie.",
            dataset, level = ERR)
      }
      dataset = which(datasets$naam_dataset == dataset)
      superweegfactor[data$tbl_dataset == dataset] = data[[varlist[[wfname]][i]]][data$tbl_dataset == dataset]
    } else if (!is.na(varlist[[wfname]][i])) {
      msg("Onbekende weegfactordefinitie opgegeven: %s. Controleer de configuratie.",
          wfname, level = ERR)
    }
    
    # rval$prob <- apply(probs, 1, prod)
    # rval$allprob <- probs
    # Dit betekent dat we het object kunnen hergebruiken, wat veel rekenkracht scheelt.
    # Hiervoor moeten we dan de variabelen aanpassen en de probs berekenen.
  }
  
  probs = as.data.frame(1/as.matrix(superweegfactor))
  current_design$allprob = probs
  current_design$prob = apply(probs, 1, prod)
  current_design$variables[[var]] = data[[var]]
  
  current_design$variables = cbind(current_design$variables, data[, colnames(data)[str_starts(colnames(data), paste0("dummy.", var))]])
  
  results = bind_rows(results, GetTableRow(var, current_design, kolom_opbouw, subsetmatches))
}
t.after = proc.time()["elapsed"]

msg("Totale rekentijd met recycling %0.2f sec voor %d variabelen.", (t.after - t.start), 5, level = MSG)
