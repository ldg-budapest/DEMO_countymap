# Import tools ####

library(tidyr)
library(dplyr)
library(stringr)
library(openxlsx)
library(officedown)
library(officer)
library(mschart)
library(rvg)
library(sf)
library(ggplot2)
library(xml2)

# Change options here, if needed ####
PATH_TO_INPUT_TABLE <- "../data/county_population.csv"
PATH_TO_PPTX_TEMPLATE <- "../configs/grey_template.pptx"
PATH_TO_SHAPES_HUN_COUNTY <- "../configs/hungary_counties"

# Initialize the ppt file ####
main_pptx_output <- PATH_TO_PPTX_TEMPLATE %>%
  read_pptx() %>%
  ph_remove(
    ph_label = ph_location_label(ph_label = "Title 1")
  ) %>%
  ph_with(
    value = "Population of counties in Hungary",
    location = ph_location_label(ph_label = "Title 1")
  ) %>%
  ph_remove(
    ph_label = ph_location_label(ph_label = "Subtitle 2")
  ) %>%
  ph_with(
    value = "LDG Team",
    location = ph_location_label(ph_label = "Subtitle 2")
  )

# Pre-define helpers for plotting ####
theme_borderless <- function() {
  theme(
    panel.border = element_blank(),
    panel.grid.major = element_blank(),
    panel.grid.minor = element_blank(),
    axis.line.x=element_blank(),
    axis.line.y=element_blank(),
    axis.text.x=element_blank(),
    axis.ticks.x=element_blank(),
    axis.text.y=element_blank(),
    axis.ticks.y=element_blank(),
    legend.background = element_rect(),
    legend.box.background = element_rect(color=NA)
  )
}

county_hun_map <- PATH_TO_SHAPES_HUN_COUNTY %>%
  read_sf() %>%
  mutate(
    NAME_1 = recode(
      NAME_1, `Fejér` = "Fejer", `Csongrád` = "Csongrad",
      `Komárom-Esztergom` = "Komarom-Esztergom", `Békés` = "Bekes",
      `Bács-Kiskun` = "Bacs-Kiskun", `Veszprém` = "Veszprem",
      `Nógrád` = "Nograd", `Hajdú-Bihar` = "Hajdu-Bihar",
      `Borsod-Abaúj-Zemplén` = "Borsod-Abauj-Zemplen",
      `Győr-Moson-Sopron` = "Gyor-Moson-Sopron",
      `Jász-Nagykun-Szolnok` = "Jasz-Nagykun-Szolnok",
      `Szabolcs-Szatmár-Bereg` = "Szabolcs-Szatmar-Bereg"
    )
  )

# Hack the default chart margins ####

ooxml_txpr <- function( fptext ){
  out <- format(fptext, type = "pml" )
  out <- gsub("a:rPr", "a:defRPr", out, fixed = TRUE)
  rpr <- "<a:p><a:pPr>%s</a:pPr></a:p>"
  rpr <- sprintf(rpr, out)
  paste0("<c:txPr><a:bodyPr/><a:lstStyle/>", rpr, "</c:txPr>")
}

table_content_xml <- function(x) {
  x_horizontal_id <- ifelse(x$x_table$horizontal, 1, 0)
  x_vertical_id <- ifelse(x$x_table$vertical, 1, 0)
  x_outline_id <- ifelse(x$x_table$outline, 1, 0)
  x_show_keys_id <- ifelse(x$x_table$show_keys, 1, 0)
  
  txpr <- ""
  if (!is.null(x$labels_fp)) {
    txpr <- ooxml_txpr(x$labels_fp)
  } else {
    txpr <- ooxml_txpr(x$theme$table_text)
  }
  
  if (FALSE) { # (x$options$table) {
    table_str <- paste0(
      "<c:dTable>",
      sprintf("<c:showHorzBorder val=\"%s\"/>", x_horizontal_id),
      sprintf("<c:showVertBorder val=\"%s\"/>", x_vertical_id),
      sprintf("<c:showOutline val=\"%s\"/>", x_outline_id),
      sprintf("<c:showKeys val=\"%s\"/>", x_show_keys_id),
      txpr,
      "</c:dTable>"
    )
  } else {
    table_str <- NULL
  }
  table_str
}

new.format.ms_chart <- function(x, id_x, id_y, sheetname = "sheet1", drop_ext_data = FALSE, ...) {
  str_ <- to_pml(x, id_x = id_x, id_y = id_y, sheetname = sheetname, asis = x$asis)
  
  if (is.null(x$x_axis$num_fmt)) {
    x$x_axis$num_fmt <- x$theme[[x$fmt_names$x]]
  }
  if (is.null(x$y_axis$num_fmt)) {
    x$y_axis$num_fmt <- x$theme[[x$fmt_names$y]]
  }
  
  x_axis_str <- axis_content_xml(x$x_axis,
                                 id = id_x, theme = x$theme,
                                 cross_id = id_y, is_x = TRUE,
                                 lab = htmlEscape(x$labels$x), rot = x$theme$title_x_rot
  )
  
  x_axis_str <- sprintf("<%s>%s</%s>", x$axis_tag$x, x_axis_str, x$axis_tag$x)
  
  y_axis_str <- axis_content_xml(x$y_axis,
                                 id = id_y, theme = x$theme,
                                 cross_id = id_x, is_x = FALSE,
                                 lab = htmlEscape(x$labels$y), rot = x$theme$title_y_rot
  )
  
  y_axis_str <- sprintf("<%s>%s</%s>", x$axis_tag$y, y_axis_str, x$axis_tag$y)
  
  
  table_str <- table_content_xml(x)
  extra_layout_text <- '<c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="0.1"/><c:y val="0.0"/><c:w val="0.9"/><c:h val="0.85"/></c:manualLayout></c:layout>'
  
  ns <- "xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
  xml_elt <- paste0("<c:plotArea ", ns, ">", extra_layout_text, str_, x_axis_str, y_axis_str, table_str, "</c:plotArea>")
  xml_doc <- read_xml(system.file(package = "mschart", "template", "chart.xml"))
  
  node <- xml_find_first(xml_doc, "//c:plotArea")
  xml_replace(node, as_xml_document(xml_elt))
  
  if (!is.null(x$labels[["title"]])) {
    chartnode <- xml_find_first(xml_doc, "//c:chart")
    title_ <- "<c:title %s><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r>%s<a:t>%s</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val=\"0\"/></c:title>"
    title_ <- sprintf(title_, ns, format(x$theme[["main_title"]], type = "pml"), htmlEscape(x$labels[["title"]]))
    xml_add_child(chartnode, as_xml_document(title_), .where = 0)
  } else { # null is not enough
    atd_node <- xml_find_first(xml_doc, "//c:chart/c:autoTitleDeleted")
    xml_attr(atd_node, "val") <- "1"
  }
  
  if (x$theme[["legend_position"]] %in% "n") {
    legend_pos <- xml_find_first(xml_doc, "//c:chart/c:legend")
    xml_remove(legend_pos)
  } else {
    legend_pos <- xml_find_first(xml_doc, "//c:chart/c:legend/c:legendPos")
    xml_attr(legend_pos, "val") <- x$theme[["legend_position"]]
    
    rpr <- format(x$theme[["legend_text"]], type = "pml")
    rpr <- gsub("a:rPr", "a:defRPr", rpr)
    labels_text_pr <- "<c:txPr xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:bodyPr/><a:lstStyle/><a:p><a:pPr>%s</a:pPr></a:p></c:txPr>"
    labels_text_pr <- sprintf(labels_text_pr, rpr)
    legend_ <- xml_find_first(xml_doc, "//c:chart/c:legend")
    xml_add_child(legend_, as_xml_document(labels_text_pr))
  }
  if (drop_ext_data) {
    xml_remove(xml_find_first(xml_doc, "//c:externalData"))
  }
  
  as.character(xml_doc)
}

environment(new.format.ms_chart) <- asNamespace("mschart")
assignInNamespace("format.ms_chart", new.format.ms_chart, ns = "mschart")

# Parse input tale ####
main_table <- PATH_TO_INPUT_TABLE %>%
  read.csv()%>%
  filter(County != "National") %>%
  group_by(Year, County) %>%
  summarise(
    value = sum(Population)
  ) %>%
  ungroup()

# Create ggplot version of map ####
p1 <- main_table %>%
  mutate(
    County = recode(
      County, Borsod="Borsod-Abauj-Zemplen", Bacs="Bacs-Kiskun",
      Gyor="Gyor-Moson-Sopron", Hajdu="Hajdu-Bihar", Szolnok="Jasz-Nagykun-Szolnok",
      Komarom="Komarom-Esztergom", Szabolcs="Szabolcs-Szatmar-Bereg"
    ),
    inc_label = format(round(value, 2), nsmall=2)
  ) %>%
  filter(Year == 2019) %>%
  left_join(county_hun_map, ., by = c("NAME_1"="County")) %>% 
  ggplot() + 
  geom_sf(aes(fill = value)) +
  geom_sf_label(
    aes(label=inc_label), size=1.8, fill=NA,
    label.padding = unit(0.15, "lines"),
    label.r = unit(0.05, "lines")
  ) +
  scale_fill_viridis_c() +
  theme_bw() +
  theme_borderless() +
  theme(
    legend.position = "bottom"
  ) +
  labs(
    x = "", y = "", fill = "Population",
  )

# Create an MSO chart for the bar plot ####

p2 <- main_table %>%
  mutate(
    County = recode(
      County, Borsod="Borsod-Abauj-Zemplen", Bacs="Bacs-Kiskun",
      Gyor="Gyor-Moson-Sopron", Hajdu="Hajdu-Bihar", Szolnok="Jasz-Nagykun-Szolnok",
      Komarom="Komarom-Esztergom", Szabolcs="Szabolcs-Szatmar-Bereg"
    ),
    inc_label = format(round(value, 2), nsmall=2)
  ) %>%
  filter(Year %in% c(2011, 2019)) %>%
  ms_barchart(
    x="County", y="value", group="Year", labels="inc_label"
  ) %>%
  # as_bar_stack(dir="horizontal") %>%
  chart_data_labels() %>%
  chart_labels_text(
    fp_text(bold = TRUE, font.size = 8)
  ) %>%
  chart_labels(xlab="", ylab="Population", title="") %>%
  chart_data_labels(position="outEnd") %>%
  chart_data_stroke(list("2011" = "#3399ff","2019" = "#ff3399")) %>%
  chart_data_fill(list("2011" = "#3399ff","2019" = "#ff3399")) %>%
  chart_labels_text(
    list(
      "2011" = fp_text(color="#3399ff", bold=FALSE, font.size=8),
      "2019" = fp_text(color="#ff3399", bold=FALSE, font.size=8)
    )
  ) %>%
  chart_settings(dir="horizontal", overlap=-50) %>%
  chart_theme(
    main_title = fp_text(font.size = 12),
    legend_text = fp_text(font.size = 10),
    axis_title_y = fp_text(font.size = 10),
    title_y_rot = 0,
    legend_position = "b"
  )

# Add plots on a new slide ####
main_pptx_output <- main_pptx_output %>%
  add_slide(
    layout = "Two Content", master = "Office Theme"
  ) %>%
  ph_with(
    value = "Population totals of counties",
    location = ph_location_type(type = "title")
  ) %>%
  ph_with(
    value = "", location = ph_location_type(type = "ftr")
  ) %>%
  ph_remove(type = "ftr") %>%
  ph_with(
    value = format(Sys.Date()),
    location = ph_location_type(type = "dt")
  ) %>%
  ph_with(
    value = "Slide 1",
    location = ph_location_type(type = "sldNum")
  ) %>%
  ph_with(
    value = dml(ggobj=p1),
    location = ph_location_label(
      ph_label = "Content Placeholder 2"
    )
  ) %>%
  ph_with(
    value = p2,
    location = ph_location_label(
      ph_label = "Content Placeholder 3"
    )
  )

# Save ppt file ####
print(main_pptx_output, "map_and_bar.pptx")
