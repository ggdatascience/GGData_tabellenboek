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
