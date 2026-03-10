# Functie voor het maken van platte kubusdata voor Swing (ABF).
# Doel: gestandaardiseerde output voor Swing, inclusief labels en dimensies.
library(data.table)

MaakKubusData <- function(
    data,
    configuraties,
    variabelen,
    crossings,
    algemeen = NULL,
    afkapwaarde = -99996,
    swing_output_bestandsnaam = "kubusdata",
    swing_unit = "p",
    output_folder = "output_swing",
    max_char_labels = 100,
    dummy_crossing_var = "dummy_crossing"
) {
  
  # Outputmap aanmaken indien nodig zodat schrijfacties nooit falen op ontbrekende paden.
  if (!dir.exists(output_folder)) dir.create(output_folder, recursive = TRUE)
  
  # Config uit de globale omgeving gebruiken wanneer die niet expliciet is meegegeven.
  # Dit behoudt backward compatibility met bestaande scripts.
  if (is.null(algemeen) && exists("algemeen", envir = .GlobalEnv)) {
    algemeen <- get("algemeen", envir = .GlobalEnv)
  }
  
  # Drempels voor privacy-masking: rijniveau (vraag) en celniveau (antwoord).
  min_obs_vraag <- if(!is.null(algemeen$min_observaties_per_vraag)) as.numeric(algemeen$min_observaties_per_vraag) else 0
  min_obs_cel   <- if(!is.null(algemeen$min_observaties_per_antwoord)) as.numeric(algemeen$min_observaties_per_antwoord) else 0
  
  # Outputnaam uit config overschrijft standaardwaarde voor herkenbare output-map.
  if (!is.na(algemeen$swing_output_bestandsnaam)) {
    swing_output_bestandsnaam <- as.character(algemeen$swing_output_bestandsnaam)
    msg("Outputnaam voor Swing opgehaald uit config: %s", swing_output_bestandsnaam, level = MSG)
  }
  
  # Afkapwaarde uit config overschrijft standaardwaarde zodat masking consistent is met Swing.
  if (!is.na(algemeen$swing_afkapwaarde)) {
    afkapwaarde <- as.numeric(algemeen$swing_afkapwaarde)
    msg("Afkapwaarde voor Swing opgehaald uit config: %s", afkapwaarde, level = MSG)
  }
  
  #Unit-waarde uit config overschrijft standaard unit zodat unitwaarde zoals gewenst is.
  if (!is.na(algemeen$swing_unit)) {  
    swing_unit <- as.character(algemeen$swing_unit) 
    msg("Unit voor Swing opgehaald uit config: %s", swing_unit, level = MSG)
  }
  
  # Variabelen voorbereiden en ontdubbelen om dubbele rijen in configs te voorkomen.
  vars_df <- variabelen
  if (any(duplicated(vars_df$variabelen))) {
    vars_df <- vars_df |> dplyr::distinct()
  }
  # Weegfactor kolommen uit variabelen hernoemen zodat ze niet botsen met configuratie
  # en altijd terug te vinden zijn in de config-rij.
  wf_cols <- grep("^weegfactor", names(vars_df), value=TRUE)
  if (length(wf_cols) > 0) {
    vars_df <- vars_df |> dplyr::rename_with(~paste0("var_", .), dplyr::all_of(wf_cols))
  }
  
  # Hernoemen van de variabele kolom naar 'vars' voor consistente verwijzingen.
  if ("variabelen" %in% names(vars_df)) {
    vars_df <- vars_df |> dplyr::rename(vars = variabelen)
  }
  
  vars <- unique(vars_df$vars)
  
  # Crossings normaliseren naar een karaktervector (ook wanneer input een data.frame is).
  if (missing(crossings) || is.null(crossings)) {
    crossings <- c()
  } else if (is.data.frame(crossings)) {
    crossings <- crossings[[1]]
  }
  crossings <- unique(crossings)
  
  # Dimensies uit configuratie ophalen om outputpaden te bouwen.
  gebiedsniveaus <- unique(configuraties$gebiedsniveau)
  jaren <- unique(configuraties$jaar_voor_analyse)
  
  # NA verwijderen zodat latere selects en filters geen fouten geven.
  vars <- vars[!is.na(vars)] 
  crossings <- crossings[!is.na(crossings)] 
  jaren <- jaren[!is.na(jaren)]
  
  
  # Check of variabelen bestaan om vroegtijdig te stoppen met duidelijke melding.
  vars_niet_in_data <- vars[!vars %in% names(data)]
  if (length(vars_niet_in_data) > 0) {
    msg("Variabele %s niet gevonden in data voor kubus export.\r\n", paste(vars_niet_in_data, collapse = ", "), level = WARN)
    return()
  }
  
  # Check of crossing bestaan
  crossings_niet_in_data <- crossings[!crossings %in% names(data)]
  if (length(crossings_niet_in_data) > 0 &&
      !(length(crossings_niet_in_data) == 1 && crossings_niet_in_data == dummy_crossing_var)) {
    msg("Crossings %s niet gevonden in data voor kubus export.\r\n", paste(crossings_niet_in_data, collapse = ", "), level = WARN)
    return()
  }
  
  if (any(crossings %in% vars)) {
    vars_in_crossing <- crossings[crossings %in% vars]
    msg("Crossings %s komt/komen ook voor in vars. Dit levert problemen op dus wordt niet als var verwerkt.\r\n", vars_in_crossing, level = WARN)
    rm(vars_in_crossing)
    vars <- vars[!vars %in% crossings]
  }
  
  
  # Controleer per gebiedsindeling of de kolom bestaat.
  # Niet-bestaande kolommen worden NA zodat downstream code voorspelbaar blijft.
  gebiedsindelingen_kolom <- sapply(configuraties$gebiedsindeling_kolom, function(kolomnaam) {
    if (is.na(kolomnaam)) {
      return(NA)
    } else if (!(kolomnaam %in% names(data))) {
      msg("Gebiedsindeling_kolom `%s` niet gevonden in data voor kubus export. Deze wordt als leeg beschouwd.\r\n", 
          kolomnaam, level = WARN)
      return(NA)
    } else {
      return(kolomnaam)
    }
  })
  
  # Zorg dat outputmappen bestaan voor alle jaar/gebiedsniveau combinaties.
  for (jaar in jaren) {
    for (gebiedsniveau in gebiedsniveaus) {
      if (!dir.exists(file.path(output_folder, jaar, gebiedsniveau))){
        dir.create(file.path(output_folder, jaar, gebiedsniveau), recursive = TRUE)
        msg("Output folder gemaakt: %s", file.path(output_folder, jaar, gebiedsniveau), level = MSG)
      }
    }
  }
  
  
  # Dummy crossing toevoegen zodat ook zonder echte crossing een dimensie beschikbaar is.
  crossings <- unique(c(crossings, dummy_crossing_var))

  # Configs bouwen: elke combinatie van configuratie, variabele en crossing.
  configs <- tidyr::expand_grid(
    configuraties,
    vars_df,
    crossings
  )
  
  # Verwijder combinaties waarin geen_crossings geldt maar toch een echte crossing staat.
  configs <- configs |> 
    filter(!(geen_crossings & crossings != dummy_crossing_var))

  # Bij geen_crossings alles behalve dummy var weghalen
  configs <- configs |> 
    filter(!(geen_crossings == TRUE & crossings != dummy_crossing_var))
  
  # Bij wel crossings dummy var weghalen
  configs <- configs |> 
    filter(!(geen_crossings == FALSE & crossings == dummy_crossing_var))
  
  # Pre-split data per dataset om herhaalde filtering te vermijden (performance optimalisatie)
  data_by_dataset <- split(data, data$tbl_dataset)
  
  # Pre-cache labels voor variabelen en crossings (performance optimalisatie)
  var_labels_cache <- lapply(stats::setNames(vars, vars), function(v) {
    vl <- labelled::val_labels(data[[v]])
    if (is.null(vl) || length(vl) == 0) {
      unieke_vals <- sort(unique(data[[v]][!is.na(data[[v]])]))
      vl <- stats::setNames(as.numeric(unieke_vals), as.character(unieke_vals))
    }
    vl
  })
  
  var_name_cache <- lapply(stats::setNames(vars, vars), function(v) {
    nm <- labelled::var_label(data[[v]])
    if (is.null(nm) || is.na(nm)) nm <- v
    nm
  })
  
  crossing_labels_cache <- lapply(stats::setNames(crossings, crossings), function(cr) {
    if (!cr %in% names(data)) return(NULL)
    cr_lbls <- labelled::val_labels(data[[cr]])
    if (is.null(cr_lbls) || length(cr_lbls) == 0) {
      vals <- sort(unique(data[[cr]][!is.na(data[[cr]])]))
      cr_lbls <- stats::setNames(as.character(vals), as.character(vals))
    }
    cr_lbls
  })
  
  crossing_name_cache <- lapply(stats::setNames(crossings, crossings), function(cr) {
    if (!cr %in% names(data)) return(cr)
    nm <- labelled::var_label(data[[cr]])
    if (is.null(nm) || is.na(nm)) nm <- cr
    nm
  })
  
  # Over elke regel van de config lopen om de platte kubusdata te maken
  configs |> 
    pwalk(function(...){
      args <- list(...)
      bron <- args$bron
      is_kubus <- if (isTRUE(args$geen_crossings)) 0 else 1
      
      # Weegfactor bepalen met prioriteit:
      # 1) dataset-specifieke override op naam, 2) dataset-specifieke override op id,
      # 3) algemene override uit variabelen, 4) standaard uit configuratie.
      
      current_weegfactor <- args$weegfactor
      
      # Overridekolommen dynamisch bepalen zodat configuraties flexibel zijn.
      ds_name_col <- paste0("var_weegfactor.d_", args$naam_dataset)
      ds_id_col <- paste0("var_weegfactor.d", args$tbl_dataset)
      global_col <- "var_weegfactor"
      
      if (!is.null(args[[ds_name_col]]) && !is.na(args[[ds_name_col]])) {
        current_weegfactor <- args[[ds_name_col]]
      } else if (!is.null(args[[ds_id_col]]) && !is.na(args[[ds_id_col]])) {
        current_weegfactor <- args[[ds_id_col]]
      } else if (!is.null(args[[global_col]]) && !is.na(args[[global_col]])) {
        current_weegfactor <- args[[global_col]]
      }
      
      
      
      # Gebruik pre-split data voor snellere toegang (vermijdt herhaalde filtering)
      kubusdata <- data_by_dataset[[as.character(args$tbl_dataset)]]
      if (is.null(kubusdata) || nrow(kubusdata) == 0) return()
      kubusdata <- kubusdata |> 
        mutate(geolevel = args$gebiedsniveau)
      
      
      
      # Jaar bepalen: vaste waarde wanneer jaarvariabele ontbreekt, anders filteren op jaar.
      if (is.na(args$jaarvariabele)) {
        kubusdata <- kubusdata |> 
          mutate(Jaar = args$jaar_voor_analyse)
        msg("jaarvariabele is NA, dus jaar wordt ingesteld op %s", args$jaar_voor_analyse)
      } else {
        kubusdata <- kubusdata |> 
          mutate(Jaar = .data[[args$jaarvariabele]]) |> 
          filter(Jaar == as.integer(args$jaar_voor_analyse))
      }
      
      # Als Gebiedsindeling_kolom ontbreekt, maak één geocode (1) zodat er toch output is.
      if (is.na(args$gebiedsindeling_kolom)) {
        kubusdata <- kubusdata |> 
          mutate(geoitem = 1)
      } else {
        kubusdata <- kubusdata |> 
          mutate(geoitem = .data[[args$gebiedsindeling_kolom]])
      }
      
      # necessary_cols <- c(args$jaarvariabele, args$gebiedsindeling, args$vars, args$crossings, args$weegfactor)
      # necessary_cols <- necessary_cols[!is.na(necessary_cols)]
      # 
      # kubusdata <- kubusdata |> 
      #   select(all_of(necessary_cols))
      
      
      # Crossing opnemen als geen_crossings == FALSE en deze bestaat (anders niet groeperen op crossing)
      use_crossing <- !isTRUE(args$geen_crossings) &&
        args$crossings != dummy_crossing_var &&
        (args$crossings %in% names(kubusdata))
      
      grouping_cols <- c("Jaar", "geolevel", "geoitem")
      if (use_crossing) {
        grouping_cols <- c(grouping_cols, args$crossings)
      }
      
      kubusdata <- kubusdata |>
        dplyr::filter(!is.na(.data$geoitem),
                      !is.na(.data[[args$vars]]),
                      !is.na(.data[[current_weegfactor]]))
      
      # Alleen crossing filteren wanneer er een echte crossing wordt gebruikt.
      if (use_crossing) {
        kubusdata <- kubusdata |>
          dplyr::filter(!is.na(.data[[args$crossings]]))
      }
      
      # Gebruik data.table voor snellere aggregatie (performance optimalisatie)
      dt <- data.table::as.data.table(kubusdata)
      by_cols <- c(grouping_cols, args$vars)
      wf_col <- current_weegfactor
      kubusdata <- dt[, .(n_gewogen = sum(get(wf_col), na.rm = TRUE), 
                          n_ongewogen = .N), 
                      by = by_cols]
      kubusdata <- tibble::as_tibble(kubusdata)
      
      # Geen records na filtering: log en ga door naar volgende config.
      if (nrow(kubusdata) == 0) {
        if (is_kubus) {
          out_dir <- file.path(output_folder, as.character(args$jaar_voor_analyse), as.character(args$gebiedsniveau))
          bestandsnaam <- file.path(out_dir, paste0(swing_output_bestandsnaam, "_", args$vars, "_", args$crossings, ".xlsx"))
        } else {
          out_dir <- file.path(output_folder, as.character(args$jaar_voor_analyse), paste0(as.character(args$gebiedsniveau), "_totaal"))
          bestandsnaam <- file.path(out_dir, paste0(swing_output_bestandsnaam, "_", args$vars, ".xlsx"))
        }
        msg("Geen data over voor kubus export voor variabele %s: %s (mogelijk alles missing).",args$vars, bestandsnaam, level = WARN)
        return()
      }
      
      # Pivot naar breed: zowel gewogen als ongewogen meenemen voor masking.
      kubusdata <- kubusdata |>
        tidyr::pivot_wider(
          id_cols = dplyr::all_of(grouping_cols),
          names_from = dplyr::all_of(args$vars),
          values_from = c(n_gewogen, n_ongewogen),
          values_fill = 0,
          names_sep = "_"
        )
      
      # Herbereken totaal ongewogen ONG als som van n_ongewogen_* kolommen.
      # Gebruik vectorized rowSums i.p.v. trage rowwise() (performance optimalisatie)
      ongewogen_cols <- grep("^n_ongewogen_", names(kubusdata), value = TRUE)
      ong_col_name <- paste0(args$vars, "_ONG")
      if (length(ongewogen_cols) > 0) {
        kubusdata[[ong_col_name]] <- rowSums(kubusdata[, ongewogen_cols, drop = FALSE], na.rm = TRUE)
      } else {
        # als er om wat voor reden dan ook geen ongewogen kolommen zijn, fallback op 0
        kubusdata[[ong_col_name]] <- 0
      }
      
      # Gewogen kolommen hernoemen naar {variabele}_{waarde} voor Swing-conventie.
      gewogen_cols <- grep("^n_gewogen_", names(kubusdata), value = TRUE)
      if (length(gewogen_cols) > 0) {
        nieuwe_namen <- sub("^n_gewogen_", paste0(args$vars, "_"), gewogen_cols)
        names(kubusdata)[match(gewogen_cols, names(kubusdata))] <- nieuwe_namen
      }
      
      # Labels ophalen uit pre-cached data (performance optimalisatie)
      vl <- var_labels_cache[[args$vars]]
      ans_values <- as.character(unname(vl))   # codes zoals "0","1","2"
      ans_labels <- names(vl)                  # labels zoals "Ja","Nee",...
      
      # Kolomnamen voor gewogen/ongewogen output en totalen.
      antwoordkolommen <- paste0(args$vars, "_", ans_values)     # gewogen data-kolommen
      ongewogen_celkolommen <- paste0("n_ongewogen_", ans_values) # per-antwoord ongewogen aantallen (alleen voor maskering)
      ong_col <- paste0(args$vars, "_ONG")
      
      # Zorg dat alle kolommen aanwezig zijn (ontbrekende aanmaken met 0).
      ontbrekend_w <- setdiff(antwoordkolommen, names(kubusdata))
      if (length(ontbrekend_w) > 0) for (mc in ontbrekend_w) kubusdata[[mc]] <- 0
      ontbrekend_u <- setdiff(ongewogen_celkolommen, names(kubusdata))
      if (length(ontbrekend_u) > 0) for (mc in ontbrekend_u) kubusdata[[mc]] <- 0
      if (!ong_col %in% names(kubusdata)) kubusdata[[ong_col]] <- 0
      
      # Basis-kolommen voor Data: tijd, geo en eventueel crossing.
      base_cols <- c("Jaar", "geolevel", "geoitem")
      if (!isTRUE(args$geen_crossings) && args$crossings %in% names(kubusdata)) {
        base_cols <- c(base_cols, args$crossings)
      }
      
      # Alleen Swing-kolommen behouden; ongewogen celniveau blijft impliciet in kubusdata.
      data_cols <- c(base_cols, antwoordkolommen, ong_col)
      data_cols <- data_cols[data_cols %in% names(kubusdata)]
      
      # Privacy: rijniveau (ONG < min_obs_vraag) -> alle waarden maskeren.
      rij_te_weinig <- kubusdata[[ong_col]] < min_obs_vraag & kubusdata[[ong_col]] != afkapwaarde
      if (any(rij_te_weinig, na.rm = TRUE)) {
        kubusdata[rij_te_weinig, antwoordkolommen] <- afkapwaarde
        kubusdata[rij_te_weinig, ong_col] <- afkapwaarde
      }
      
      # Privacy: celniveau (n_ongewogen_<code> < min_obs_cel) -> alleen die cel maskeren.
      if (min_obs_cel > 0) {
        for (i in seq_along(ans_values)) {
          cel_u_col <- ongewogen_celkolommen[i]
          cel_w_col <- antwoordkolommen[i]
          if (cel_u_col %in% names(kubusdata) && cel_w_col %in% names(kubusdata)) {
            mask <- kubusdata[[cel_u_col]] < min_obs_cel & kubusdata[[cel_u_col]] != afkapwaarde
            if (any(mask, na.rm = TRUE)) {
              kubusdata[mask, cel_w_col] <- afkapwaarde
            }
          }
        }
      }
      
      # Data klaarzetten voor export (zonder n_ongewogen_* hulpkolommen).
      kubus_df <- kubusdata |>
        dplyr::select(dplyr::all_of(data_cols))
      
      # Excel workbook opbouwen met standaard Swing tabbladen.
      wb <- openxlsx::createWorkbook()
      
      # 1) Data: kernoutput met gewogen waarden en ONG.
      openxlsx::addWorksheet(wb, "Data")
      openxlsx::writeData(wb, "Data", kubus_df)
      style_fmt <- openxlsx::createStyle(numFmt = "[=-99996]0;0.000")
      openxlsx::addStyle(wb, "Data", style_fmt, cols = which(names(kubus_df) %in% antwoordkolommen), 
                         rows = 2:(nrow(kubus_df) + 1), gridExpand = TRUE)
      
      # 2) Data_def: kolomdefinities met type (period/geolevel/geoitem/dim/var).
      openxlsx::addWorksheet(wb, "Data_def")
      data_def_cols <- c("Jaar", "geolevel", "geoitem")
      if (!isTRUE(args$geen_crossings) && args$crossings %in% names(kubus_df)) {
        data_def_cols <- c(data_def_cols, args$crossings)
      }
      data_def_cols <- c(data_def_cols, antwoordkolommen, ong_col)
      data_def_cols <- data_def_cols[data_def_cols %in% names(kubus_df)]
      data_def <- data.frame(
        col = data_def_cols,
        type = dplyr::case_when(
          data_def_cols == "Jaar" ~ "period",
          data_def_cols == "geolevel" ~ "geolevel",
          data_def_cols == "geoitem" ~ "geoitem",
          data_def_cols == args$crossings ~ "dim",
          TRUE ~ "var"
        )
      )
      openxlsx::writeData(wb, "Data_def", data_def)
      
      # 3) Label_var: beschrijving van indicators.
      openxlsx::addWorksheet(wb, "Label_var")
      label_var <- data.frame(
        Onderwerpcode = c(antwoordkolommen, ong_col),
        Naam = c(paste0("Aantal ", ans_labels), "Totaal aantal ongewogen"),
        Eenheid = rep("Personen", length(antwoordkolommen) + 1)
      )
      openxlsx::writeData(wb, "Label_var", label_var)
      
      # 4) Dimensies + items (alleen met crossings)
      if (!isTRUE(args$geen_crossings) && args$crossings %in% colnames(data)) {
        openxlsx::addWorksheet(wb, "Dimensies")
        # Gebruik pre-cached crossing naam (performance optimalisatie)
        dim_naam <- crossing_name_cache[[args$crossings]]
        openxlsx::writeData(
          wb, "Dimensies",
          data.frame(Dimensiecode = args$crossings, Naam = dim_naam)
        )
        
        openxlsx::addWorksheet(wb, args$crossings)
        # Gebruik pre-cached crossing labels (performance optimalisatie)
        cr_lbls <- crossing_labels_cache[[args$crossings]]
        df_items <- data.frame(
          Itemcode = unname(cr_lbls),
          Naam = names(cr_lbls),
          Volgnr = seq_along(cr_lbls)
        )
        openxlsx::writeData(wb, args$crossings, df_items)
      }
      
      # 5) Indicators: Swing indicator-definities inclusief formules.
      openxlsx::addWorksheet(wb, "Indicators")
      # Gebruik pre-cached variabele naam (performance optimalisatie)
      variabel_naam <- var_name_cache[[args$vars]]
      
      is_dichotoom <- suppressWarnings(all(as.numeric(ans_values) %in% c(0,1)) && length(ans_values) == 2)
      
      indicator_codes <- c(
        antwoordkolommen,
        ong_col,
        paste0(args$vars, "_GEW"),
        if (is_dichotoom) paste0(args$vars, "_perc") else paste0(args$vars, "_", ans_values, "_perc")
      )
      indicator_names <- c(
        paste0("Aantal ", ans_labels),
        "Totaal aantal ongewogen",
        "Totaal aantal gewogen",
        if (is_dichotoom) substr(variabel_naam, 1, max_char_labels)
        else {
          nm <- paste0(variabel_naam, ", ", ans_labels)
          ifelse(nchar(nm) > max_char_labels, substr(nm, 1, max_char_labels), nm)
        }
      )
      som_formule <- paste(antwoordkolommen, collapse = "+")
      if (is_dichotoom) {
        perc_formulas <- paste0("(", args$vars, "_1/(", som_formule, "))*100")
        formulas <- c(rep("", length(ans_values) + 1), som_formule, perc_formulas)
      } else {
        perc_formulas <- vapply(ans_values, function(v) paste0("(", args$vars, "_", v, "/(", som_formule, "))*100"), character(1))
        formulas <- c(rep("", length(ans_values) + 1), som_formule, perc_formulas)
      }
      
      indicators_df <- data.frame(
        `Indicator code` = indicator_codes,
        Name = indicator_names,
        Unit = c(rep("personen", length(ans_values) + 2),
                 if (is_dichotoom) swing_unit else rep(swing_unit, length(ans_values))),
        `Aggregation indicator` = c(rep("", length(ans_values) + 2),
                                    if (is_dichotoom) paste0(args$vars, "_GEW") else rep(paste0(args$vars, "_GEW"), length(ans_values))),
        Formula = formulas,
        `Data type` = c(rep("Numeric", length(ans_values) + 2),
                        if (is_dichotoom) "Percentage" else rep("Percentage", length(ans_values))),
        Visible = c(rep(0, length(ans_values) + 2),
                    if (is_dichotoom) 1 else rep(1, length(ans_values))),
        `Threshold value` = c(rep("", length(ans_values) + 2),
                              if (is_dichotoom) min_obs_vraag else rep(min_obs_vraag, length(ans_values))),
        `Threshold Indicator` = c(rep("", length(ans_values) + 2),
                                  if (is_dichotoom) paste0(args$vars, "_ONG") else rep(paste0(args$vars, "_ONG"), length(ans_values))),
        Cube = rep(1, length(indicator_codes)), # Cube moet blijkbaar altijd 1 zijn
        Source = rep(bron, length(indicator_codes)),
        check.names = FALSE
      )
      openxlsx::writeData(wb, "Indicators", indicators_df)
      
      # Opslaan in submap {output_folder}/{jaar}/{gebiedsniveau}
      if (is_kubus) {
        out_dir <- file.path(output_folder, as.character(args$jaar_voor_analyse), as.character(args$gebiedsniveau))
        if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)
        bestandsnaam <- file.path(out_dir, paste0(swing_output_bestandsnaam, "_", args$vars, "_", args$crossings, ".xlsx"))
      } else {
        out_dir <- file.path(output_folder, as.character(args$jaar_voor_analyse), paste0(as.character(args$gebiedsniveau), "_totaal"))
        if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)
        bestandsnaam <- file.path(out_dir, paste0(swing_output_bestandsnaam, "_", args$vars, ".xlsx"))
      }
      
      # Schrijfbestand opslaan; errors loggen zonder het hele proces te stoppen.
      tryCatch({
        openxlsx::saveWorkbook(wb, file = bestandsnaam, overwrite = TRUE)
        msg("Kubusdata opgeslagen voor %s: %s", args$vars, bestandsnaam, level = MSG)
      }, error = function(e) {
        msg("Fout bij opslaan kubusdata voor %s: %s", args$vars, e$message, level = ERR)
      })

      },
      .progress = TRUE
      )
}
