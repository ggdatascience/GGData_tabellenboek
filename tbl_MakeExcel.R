#
#
# Deze functie schrijft de resultatentabel vanuit survey om naar een bruikbare vorm.
# Opmaak wordt ingelezen vanuit de configuratie (tabblad opmaak) en verdere instellingen
# van de tabbladen indeling_rijen, indeling_kolommen en algemeen.
#
#

NA.identical = function (vect, cmp) {
  if (is.na(cmp)) {
    return(is.na(vect))
  } else {
    return(vect == cmp)
  }
}

MakeExcel = function (results, col.design, subset, subset.val, subsetmatches) {
  subset.name = names(subset.val)
  subset.val = unname(subset.val)
  
  wb = createWorkbook(creator="GGData Tabellenboek")
  
  c = 1
  for (i in 1:length(indeling_rijen)) {
    if (indeling_rijen$type[i] == "var") {
      
      # matrix maken met de benodigde resultaten in de juiste volgorde
      output = matrix(nrow=length(unique(results$val[results$var == indeling_rijen$inhoud[i]])), ncol=nrow(col.design))
      for (j in 1:nrow(col.design)) {
        if (!is.na(col.design$subset[j])) {
          subset.col = subset.val
          if (col.design$subset[j] != subset) {
            subset.col = subsetmatches[subsetmatches[,1] == subset.val, col.design$subset[j]]
          }
          output[,j] = results$perc.weighted[which(NA.identical(results$dataset, col.design$dataset[j]) & NA.identical(results$subset, col.design$subset[j]) &
                                               NA.identical(results$year, col.design$year[j]) & NA.identical(results$crossing, col.design$crossing[j]) &
                                               NA.identical(results$crossing.val, col.design$crossing.val[j]) & NA.identical(results$subset.val, subset.col))]
        }
        else {
          output[,j] = results$perc.weighted[which(NA.identical(results$dataset, col.design$dataset[j]) & NA.identical(results$subset, col.design$subset[j]) &
                                               NA.identical(results$year, col.design$year[j]) & NA.identical(results$crossing, col.design$crossing[j]) &
                                               NA.identical(results$crossing.val, col.design$crossing.val[j]))]
        }
      }
    }
  }
  
  saveWorkbook(wb, sprintf("output/%s.xlsx", subset.name), overwrite=T)
}