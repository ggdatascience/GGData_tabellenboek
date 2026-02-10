#
# Script om specifieke stukken tekst (kolomkoppen of antwoordwaardes) te vervangen bij specifieke vragen.
# Dit kan bijvoorbeeld voorkomen als er voor één vraag een andere leeftijdscategorie is aangehouden, of
# de regioindeling net anders is. Toegevoegd als extra service, geen officieel onderdeel van het grote
# tabellenboekscript. (Lees: kan mogelijk onvoldoende functioneren.)
#

library(tidyverse)
library(openxlsx)
#library(xml2)
library(this.path)

setwd(dirname(this.path()))
source("tbl_helpers.R")

search = read.xlsx(choose.files(caption="Selecteer configuratiebestand...",
                                filters=c("Excel-bestand (*.xlsx)","*.xlsx"),
                                multi=F))
if (any(!c("vraag", "oud", "nieuw") %in% colnames(search))) {
  stop("De benodigde kolommen ontbreken; vraag, oud of nieuw.")
}
#search = data.frame(vraag="Ervaren gezondheid in 3", oud="Vrouw", nieuw="Test")

files = choose.files(caption="Selecteer tabellenboeken", filters=c("HTML-bestand (*.html)","*.html"))

# dit zou een prachtige oplossing zijn, ware het niet dat xml2 weigert om karakters te unescapen...
# als je dit doet krijg je dit soort gedrochten in je tekst:
# Een tekstregel&comma; gevolgd door een lege regel&colon;
#
# recursive_replace = function (basenode, old, new) {
#   for (node in xml_children(basenode)) {
#     if (is_empty(xml_children(node))) {
#       # geen subnodes; alleen tekst om te controleren
#       if (str_detect(xml_text(node), old)) {
#         xml_text(node) = str_replace(xml_text(node), old, new)
#       }
#     } else {
#       # wel subnodes; ga dieper
#       node = recursive_replace(node, old, new)
#     }
#   }
#   return(basenode)
# }
#
# for (f in files) {
#   tabellenboek = read_html(f)
#   
#   for (s in 1:nrow(search)) {
#     tbl = xml_find_first(tabellenboek, paste0("//table/caption[text()[contains(.,'", search$vraag[s], "')]]/.."))
#     if (xml_name(tbl) != "table") {
#       printf("Let op: het eerste element wat gevonden werd met de vraagtekst '%s' is geen tabel. Huidige vervanging wordt overgeslagen.", search$vraag[s])
#       next
#     }
#     
#     tbl = recursive_replace(tbl, search$oud[s], search$nieuw[s])
#   }
#   
#   write_html(tabellenboek, paste0(str_sub(f, end=-6), "_corr.html"), options="as_html")
# }

# poging 2, quick and dirty regex
for (f in files) {
  tabellenboek = read_file(f)
  
  for (s in 1:nrow(search)) {
    old = str_extract(tabellenboek, regex(paste0("<table>\\s+<caption>.*?", search$vraag[s], ".*?</caption>(.|\\s)*?</table>"), multiline=T))
    # waardes zitten altijd in een td of th, dus we kunnen mooi matchen op <tx>tekst</tx>
    new = str_replace_all(old, paste0(">(.*?)", search$oud[s], "(.*?)</t"), paste0(">\\1", search$nieuw[s], "\\2</t"))
    
    if (!is.na(old))
      tabellenboek = str_replace(tabellenboek, fixed(old), new)
  }
  
  write_file(tabellenboek, paste0(str_sub(f, end=-6), "_corr.html"))
}
