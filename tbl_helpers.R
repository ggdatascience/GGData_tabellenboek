#
#
# Helpers voor algemeen gebruik.
#
#

# C-style print functie
printf = function (...) cat(paste(sprintf(...),"\n"))

## log levels
ERR = 1
WARN = 2
MSG = 3
DEBUG = 4

# functie om (fout)meldingen weer te geven
# middels het level kunnen verschillende niveau's worden aangegeven, waarbij een lager niveau
# altijd wordt weergegeven (dus msg("test", WARN) wordt ook getoond bij log.level = 3) 
# - msg kan een object of een sprintf-style string zijn
# - level moet een log level zijn
# - overige argumenten worden doorgegeven aan sprintf
msg = function (msg, ..., level = DEBUG) {
  # alleen output geven als het ewenste level gelijk is of hoger dan de gegeven waarde
  if (exists("log.level")) {
    if (level > log.level) {
      return()
    } 
  }
  
  # schrijf naar logbestand
  if (log.save) {
    if (is.character(msg)) {
      tryCatch({
        cat(sprintf(paste0("[%s - %s] ", msg, "\n"), Sys.time(), deparse(substitute(level)), ...), file = "log.txt", append=T)
      },
      error=function(e) { warning(paste0("Fout bij het schrijven naar de log: ", e)) })
      
    } else {
      tryCatch({
        cat(sprintf(paste0("[%s - %s] ", msg, "\n"), Sys.time(), deparse(substitute(level))), file = "log.txt", append=T)
      },
      error=function(e) { warning(paste0("Fout bij het schrijven naar de log: ", e)) })
    }
  }
  
  # geen gewenst level of level hoog genoeg; printen
  if (level == ERR) {
    if (is.character(msg))
      stop(sprintf(msg, ...))
    else
      stop(msg)
  } else if (level == WARN) {
    if (is.character(msg))
      message(sprintf(msg, ...))
    else
      message(msg)
  } else if (level == MSG) {
    if (is.character(msg))
      message(sprintf(msg, ...))
    else
      message(msg)
  } else {
    if (is.character(msg))
      printf(paste0(ifelse(level == DEBUG, "[DEBUG] ", ""), msg), ...)
    else
      print(msg)
  }
}

# identical() blijkt niet altijd goed te werken
# bij numerieke waarden kan identical() raar doen, bijv. 1,2,3 != 1,2,3 (volgens identical() dan)
# andersom doet all.equal het niet goed met kolommen met alleen NA
# vandaar dit gedrocht
identical.enough = function (x, y) {
  if (!all(dim(x) == dim(y))) {
    return(F)
  }
  
  is.identical = T
  for (i in 1:ncol(x)) {
    if (!isTRUE(all.equal(unname(x[,i]), unname(y[,i])))) {
      if (!(all(is.na(unname(x[,i]))) && all(is.na(unname(y[,i]))))) {
        is.identical = F
      }
    }
  }
  return(is.identical)
}

# bepaal welke rijen meegeteld/getoond moeten worden op basis van dichotoom regels
add_dichotoom_flags = function (results, dichotoom, niet_dichotoom, algemeen) {
  if (!"val" %in% colnames(results)) {
    return(results)
  }
  if (is.null(algemeen$waarden_dichotoom)) {
    return(results)
  }

  dichotoom.vals = algemeen$waarden_dichotoom %>%
    str_split("\\|") %>%
    unlist() %>%
    str_split(",") %>%
    lapply(as.numeric)

  var_levels = results %>%
    filter(!is.na(val)) %>%
    group_by(var) %>%
    summarize(levels = list(sort(unique(suppressWarnings(as.numeric(val))))), .groups = "drop")

  var_flags = var_levels %>%
    mutate(
      is_dichotoom = case_when(
        var %in% niet_dichotoom ~ FALSE,
        var %in% dichotoom ~ TRUE,
        TRUE ~ purrr::map_lgl(levels, ~ {
          lv = .x
          if (length(lv) == 0) return(FALSE)
          isTRUE(all.equal(lv, c(0, 1))) ||
            isTRUE(all.equal(lv, c(0))) ||
            isTRUE(all.equal(lv, c(1))) ||
            any(purrr::map_lgl(dichotoom.vals, ~ identical(.x, lv)))
        })
      )
    ) %>%
    select(var, is_dichotoom)

  results %>%
    left_join(var_flags, by = "var") %>%
    mutate(
      is_dichotoom = if_else(
        coalesce(is_dichotoom, FALSE),
        !is.na(val) & suppressWarnings(as.numeric(val)) == 1,
        FALSE
      )
    )
}

count_tests <- function(
    results,
    algemeen,
    crossings_toetsen,
    indeling_rijen,
    dichotoom,
    niet_dichotoom
) {
  if (!"val" %in% colnames(results)) return(results)
  if (is.null(algemeen$waarden_dichotoom)) return(results)
  
  dichotoom.vals = algemeen$waarden_dichotoom %>%
    str_split("\\|") %>%
    unlist() %>%
    str_split(",") %>%
    lapply(as.numeric)
  
  levels = results %>%
    filter(!is.na(val)) %>%
    # Onderdrukte variabelen uitsluiten op basis van voorberekende var_suppressed
    filter(!var_suppressed) %>%
    group_by(dataset, subset, subset.val, var, crossing) %>%
    filter(!all(is.na(sign.vs))) %>%
    filter(!(is.na(sign))) %>%
    summarize(
      levels = list(sort(unique(suppressWarnings(as.numeric(val))))),
      .groups = "drop"
    ) %>%
    mutate(n_levels = purrr::map_int(levels, length))
  
  levels_expanded = levels %>%
    mutate(
      is_dichotoom = case_when(
        var %in% niet_dichotoom ~ FALSE,
        var %in% dichotoom ~ TRUE,
        TRUE ~ purrr::map_lgl(levels, ~ {
          lv = .x
          if (length(lv) == 0) return(FALSE)
          isTRUE(all.equal(lv, c(0, 1))) ||
            isTRUE(all.equal(lv, c(0))) ||
            isTRUE(all.equal(lv, c(1))) ||
            any(purrr::map_lgl(dichotoom.vals, ~ identical(.x, lv)))
        })
      )
    )
  
  indeling_rijen_vars <- indeling_rijen %>%
    filter(type == "var") %>%
    select(inhoud, waardes, verberg_crossings) %>%
    mutate(verberg_crossings = !is.na(verberg_crossings))
  
  
  if(length(crossings_toetsen) > 0){
    crossings_toetsen_tibble <- tibble(
      crossing = names(crossings_toetsen),
      crossings_toetsen = unname(crossings_toetsen)
    )
  } else {
    # Lege tibble maken om errors te voorkomen bij left_join als crossings_toetsen leeg is
    crossings_toetsen_tibble <- tibble(
      crossing = NA,
      crossings_toetsen = NA
    )
  }
  
  
  levels_expanded %>%
    left_join(indeling_rijen_vars, by = join_by(var == inhoud)) %>%
    left_join(crossings_toetsen_tibble, by = join_by(crossing)) %>%
    mutate(
      n_sign_tests = case_when(
        # Volgorde hier is erg belangrijk!
        !is.na(crossing) & coalesce(crossings_toetsen, FALSE) == FALSE ~ 0,
        # verberg_crossings is al verwerkt in var_suppressed, maar ook hier expliciet
        !is.na(crossing) & coalesce(verberg_crossings, FALSE) == TRUE  ~ 0,
        !is.na(waardes) ~ purrr::map_int(stringr::str_split(waardes, "\\|"), length),
        is_dichotoom    ~ 1L,
        .default        = n_levels
      )
    )
}

# Voeg onderdrukkingsvlaggen toe aan results:
# - n_question: totaal aantal respondenten per vraag (som van n.unweighted per kolomgroep)
# - suppression: onderdrukkingscode per rij (zie codes hieronder)
# - var_suppressed: TRUE als de variabele als geheel onderdrukt is
#
# Onderdrukkingscodes:
#   0  = zichtbaar
#  -1  = A_TOOSMALL  : antwoordcategorie te klein of onder afkapwaarde
#  -2  = Q_TOOSMALL  : hele vraag te weinig respondenten
#  -3  = Q_MISSING   : vraag geheel afwezig / verborgen crossing
#  -4  = A_EXACTZERO : exact nul procent (telt niet mee voor var_suppressed)
#
# Parameters:
#   results        - dataframe met berekende percentages
#   algemeen       - configuratietabel met drempelwaarden
#   indeling_rijen - configuratietabel met verberg_crossings per variabele
add_suppression_flags <- function(results, algemeen, indeling_rijen) {
  # Haal verberg_crossings op per variabele uit de tabelindeling
  verberg_crossings_lookup <- indeling_rijen %>%
    filter(type == "var") %>%
    select(inhoud, verberg_crossings) %>%
    mutate(verberg_crossings = !is.na(verberg_crossings))

  results %>%
    left_join(verberg_crossings_lookup, by = c("var" = "inhoud")) %>%
    mutate(verberg_crossings = coalesce(verberg_crossings, FALSE)) %>%
    # Bereken vraag-niveau statistieken per kolomgroep
    group_by(var, dataset, subset, subset.val, year, crossing, crossing.val, sign.vs) %>%
    mutate(
      n_question = sum(n.unweighted, na.rm = TRUE),
      all_na = all(is.na(perc.weighted)),
      # Exact nul (n=0) telt niet mee als 'te klein antwoord'; die krijgen A_EXACTZERO
      any_answer_toosmall = any(
        n.unweighted > 0 & n.unweighted < algemeen$min_observaties_per_antwoord,
        na.rm = TRUE
      )
    ) %>%
    ungroup() %>%
    mutate(
      suppression = case_when(
        # Kolomniveau: verberg_crossings is gezet en dit is een crossing-kolom
        verberg_crossings & !is.na(crossing)                                      ~ -3L, # Q_MISSING
        # Kolomniveau: alle waarden zijn NA
        all_na                                                                    ~ -3L, # Q_MISSING
        # Kolomniveau: geen respondenten
        n_question == 0                                                           ~ -3L, # Q_MISSING
        # Kolomniveau: te weinig respondenten voor de hele vraag
        n_question < algemeen$min_observaties_per_vraag                           ~ -2L, # Q_TOOSMALL
        # Celniveau: exact nul (vóór de n.unweighted check, want n=0 impliceert perc=0)
        perc.weighted == 0                                                        ~ -4L, # A_EXACTZERO
        # Kolomniveau: een antwoord te klein én config zegt hele vraag onderdrukken
        any_answer_toosmall & algemeen$vraag_verbergen_bij_missend_antwoord       ~ -1L, # A_TOOSMALL
        # Celniveau: dit specifieke antwoord te weinig respondenten
        n.unweighted < algemeen$min_observaties_per_antwoord                      ~ -1L, # A_TOOSMALL
        # Celniveau: onder de afkapwaarde
        perc.weighted <= algemeen$afkapwaarde_antwoord & perc.weighted > 0        ~ -1L, # A_TOOSMALL
        TRUE                                                                      ~ 0L
      )
    ) %>%
    # var_suppressed: TRUE als de variabele als geheel onderdrukt is
    # A_EXACTZERO (-4) telt hierbij niet mee als onderdrukking
    group_by(dataset, subset, subset.val, year, crossing, var, val) %>%
    mutate(var_suppressed = !all(suppression == 0 | suppression == -4L)) %>%
    ungroup() %>%
    # Tijdelijke hulpkolommen opruimen
    select(-all_na, -any_answer_toosmall, -verberg_crossings)
}
