if (!requireNamespace("pacman", quietly = TRUE)) install.packages("pacman")
pacman::p_load(tidyverse, readxl, sf, geobr, leaflet, htmlwidgets, htmltools, base64enc)

options(scipen = 999)
sf::sf_use_s2(FALSE)
options(jsonlite.warn_vec_names = FALSE)

# Se BASE_DIR não existir, assume o diretório atual
if (!exists("BASE_DIR")) BASE_DIR <- getwd()

.gerar_pontos_seguro <- function(sf_poligonos_4326) {
  sf_poligonos_3857 <- sf::st_transform(sf_poligonos_4326, 3857)
  sf_pontos_3857 <- sf_poligonos_3857 |>
    dplyr::mutate(geom = sf::st_point_on_surface(geom)) |>
    sf::st_as_sf()
  sf::st_transform(sf_pontos_3857, 4326)
}

gerar_mapa_individual <- function(aba, titulo = NULL, arquivo = "municipios.xlsx", logo_nome = "FNP_principal.png", out_dir = getwd()) {
  
  logo_path <- file.path(BASE_DIR, logo_nome)
  arquivo_path <- file.path(BASE_DIR, arquivo)
  
  if (!file.exists(logo_path)) stop("Logo não encontrada: ", logo_path)
  if (!file.exists(arquivo_path)) stop("Excel não encontrado: ", arquivo_path)

  if (is.null(titulo)) titulo <- aba

  df <- readxl::read_excel(arquivo_path, sheet = aba)
  
  names(df) <- names(df) |>
    stringr::str_replace_all("\u00A0", " ") |>
    stringr::str_squish() |>
    stringr::str_replace_all(" ", "_") |>
    stringr::str_to_lower()
  
  col_cod <- dplyr::case_when(
    "cod_ibge" %in% names(df) ~ "cod_ibge",
    "cod_ibg" %in% names(df) ~ "cod_ibg",
    "codigo_ibge" %in% names(df) ~ "codigo_ibge",
    TRUE ~ NA_character_
  )
  
  if (is.na(col_cod)) return(NULL) 
  
  municipios_sel <- df |>
    transmute(cod_ibge = as.integer(.data[[col_cod]])) |>
    filter(!is.na(cod_ibge)) |>
    distinct()
  
  if (nrow(municipios_sel) == 0) return(NULL)

  ler_geobr <- function() {
    mapa_estados <- geobr::read_state(year = 2020, showProgress = FALSE) |> sf::st_transform(4326)
    mapa_municipios <- geobr::read_municipality(year = 2020, showProgress = FALSE) |> sf::st_transform(4326)
    list(estados = mapa_estados, municipios = mapa_municipios)
  }
  
  geo <- tryCatch(ler_geobr(), error = function(e) { geobr::cleanup_geobr(); ler_geobr() })
  
  municipios_geo <- geo$municipios |> filter(code_muni %in% municipios_sel$cod_ibge)
  
  municipios_pontos <- .gerar_pontos_seguro(municipios_geo) |>
    mutate(
      label = paste0(name_muni, " (", abbrev_state, ")"),
      popup = paste0("<b>", name_muni, " (", abbrev_state, ")</b><br>IBGE: ", code_muni)
    )
  
  bb_pts <- sf::st_bbox(municipios_pontos)
  pad <- 0.9
  xmin <- as.numeric(unname(bb_pts["xmin"])) - pad
  ymin <- as.numeric(unname(bb_pts["ymin"])) - pad
  xmax <- as.numeric(unname(bb_pts["xmax"])) + pad
  ymax <- as.numeric(unname(bb_pts["ymax"])) + pad
  
  uf_regiao <- tibble::tribble(
    ~abbrev_state, ~regiao,
    "AC","Norte","AL","Nordeste","AP","Norte","AM","Norte","BA","Nordeste",
    "CE","Nordeste","DF","Centro-Oeste","ES","Sudeste","GO","Centro-Oeste",
    "MA","Nordeste","MT","Centro-Oeste","MS","Centro-Oeste","MG","Sudeste",
    "PA","Norte","PB","Nordeste","PR","Sul","PE","Nordeste","PI","Nordeste",
    "RJ","Sudeste","RN","Nordeste","RO","Norte","RR","Norte","RS","Sul",
    "SC","Sul","SE","Nordeste","SP","Sudeste","TO","Norte"
  )
  
  cont_regiao <- municipios_pontos |>
    sf::st_drop_geometry() |>
    dplyr::select(abbrev_state) |>
    dplyr::left_join(uf_regiao, by = "abbrev_state") |>
    dplyr::count(regiao, name = "n") |>
    dplyr::mutate(regiao = factor(regiao, levels = c("Norte","Nordeste","Centro-Oeste","Sudeste","Sul"))) |>
    dplyr::arrange(regiao)
  
  html_regioes <- paste0(
    "<div style='background: rgba(255,255,255,0.92); border: 1px solid rgba(17,24,39,0.12); border-radius: 14px; box-shadow: 0 10px 28px rgba(0,0,0,0.18); padding: 12px 14px; font-family: sans-serif; color: #111827; min-width: 260px;'>",
    "<div style='font-size: 18px; font-weight: 800; margin-bottom: 8px;'>Municípios por região</div>",
    paste0("<div style='display:flex; justify-content:space-between; font-size: 16px; padding: 3px 0;'>", cont_regiao$regiao, "<span style='font-weight:800;'>", cont_regiao$n, "</span></div>", collapse = ""),
    "<div style='height:1px; background: rgba(17,24,39,0.10); margin: 8px 0;'></div>",
    "<div style='display:flex; justify-content:space-between; font-size: 16px;'><span style='font-weight:700;'>Total</span><span style='font-weight:900;'>", nrow(municipios_pontos), "</span></div></div>"
  )
  
  logo_uri <- base64enc::dataURI(file = logo_path, mime = "image/png")
  html_logo <- paste0(
    "<div style='background: rgba(255,255,255,0.92); border: 1px solid rgba(17,24,39,0.12); border-radius: 14px; box-shadow: 0 10px 28px rgba(0,0,0,0.18); padding: 10px 12px; display:flex; align-items:center; justify-content:center;'>",
    "<img src='", logo_uri, "' style='height:64px; width:auto; display:block;'/></div>"
  )
  
  css_cluster <- htmltools::tags$style(htmltools::HTML("
    .marker-cluster.casd-blue-small { background: rgba(30,58,138,0.20); } .marker-cluster.casd-blue-small div { background: rgba(30,58,138,0.85); }
    .marker-cluster.casd-blue-medium { background: rgba(30,58,138,0.24); } .marker-cluster.casd-blue-medium div { background: rgba(30,58,138,0.92); }
    .marker-cluster.casd-blue-large { background: rgba(30,58,138,0.28); } .marker-cluster.casd-blue-large div { background: rgba(30,58,138,0.98); }
    .marker-cluster div { border: 2px solid rgba(255,255,255,0.90); box-shadow: 0 8px 20px rgba(0,0,0,0.22); color: #ffffff; font-weight: 800; font-family: sans-serif; }
    .marker-cluster-small div { width: 34px; height: 34px; line-height: 34px; border-radius: 17px; }
    .marker-cluster-medium div { width: 40px; height: 40px; line-height: 40px; border-radius: 20px; }
    .marker-cluster-large div { width: 46px; height: 46px; line-height: 46px; border-radius: 23px; }
  "))
  
  icon_create_function <- htmlwidgets::JS("
    function (cluster) {
      var count = cluster.getChildCount();
      var size = 'casd-blue-small';
      if (count >= 10 && count < 50) size = 'casd-blue-medium';
      if (count >= 50) size = 'casd-blue-large';
      return new L.DivIcon({ html: '<div><span>' + count + '</span></div>', className: 'marker-cluster ' + size, iconSize: new L.Point(46, 46) });
    }
  ")
  
  html_titulo <- paste0(
    "<div style='background: rgba(255,255,255,0.92); border: 1px solid rgba(17,24,39,0.12); border-radius: 14px; box-shadow: 0 10px 28px rgba(0,0,0,0.18); padding: 14px 18px; font-family: sans-serif; color: #111827;'>",
    "<div style='font-size: 42px; font-weight: 800; line-height: 1.05;'>", titulo, "</div>",
    "<div style='font-size: 24px; font-weight: 600; opacity: 0.85; margin-top: 6px;'>Total: ", nrow(municipios_pontos), "</div></div>"
  )
  
  m <- leaflet(options = leafletOptions(preferCanvas = TRUE)) |>
    addProviderTiles(providers$CartoDB.PositronNoLabels) |>
    addProviderTiles(providers$CartoDB.PositronOnlyLabels, options = providerTileOptions(opacity = 0.85)) |>
    addPolygons(data = geo$estados, weight = 1, color = "white", fillColor = "#2B5DAA", fillOpacity = 0.18, opacity = 1, smoothFactor = 0.5) |>
    addCircleMarkers(data = municipios_pontos, radius = 6, stroke = TRUE, weight = 2, color = "white", fillColor = "#1E3A8A", fillOpacity = 0.95,
                     label = ~label, popup = ~popup, labelOptions = leaflet::labelOptions(direction = "auto", textsize = "13px", opacity = 0.95, style = list("padding" = "4px 6px")),
                     clusterOptions = markerClusterOptions(iconCreateFunction = icon_create_function, spiderfyOnMaxZoom = TRUE, showCoverageOnHover = FALSE, zoomToBoundsOnClick = TRUE, removeOutsideVisibleBounds = TRUE, disableClusteringAtZoom = 8)) |>
    addControl(html = html_titulo, position = "topright") |>
    addControl(html = html_regioes, position = "bottomleft") |>
    addControl(html = html_logo, position = "bottomright") |>
    fitBounds(xmin, ymin, xmax, ymax) |>
    htmlwidgets::prependContent(css_cluster)
  
  aba_slug <- tolower(aba) |> gsub("[^a-z0-9]+", "_", x = _) |> gsub("^_+|_+$", "", x = _)
  filename <- paste0("mapa_", aba_slug, ".html")
  out_html <- file.path(out_dir, filename)
  
  # === AQUI ESTÁ A CORREÇÃO DE MEMÓRIA ===
  # selfcontained = FALSE evita que o Pandoc tente zipar tudo
  htmlwidgets::saveWidget(m, out_html, selfcontained = FALSE)
  
  return(list(nome = aba, arquivo = filename))
}

gerar_site_completo <- function() {
  arquivo_excel <- file.path(BASE_DIR, "municipios.xlsx")
  if (!file.exists(arquivo_excel)) stop("Arquivo municipios.xlsx não encontrado.")
  
  abas <- readxl::excel_sheets(arquivo_excel)
  cat("\nIniciando geração de todos os mapas...\n")
  
  mapas_gerados <- list()
  
  for (aba in abas) {
    cat("-> Processando aba:", aba, "\n")
    res <- gerar_mapa_individual(aba, out_dir = BASE_DIR)
    if (!is.null(res)) {
      mapas_gerados[[length(mapas_gerados) + 1]] <- res
    }
  }
  
  html_links <- ""
  for (m in mapas_gerados) {
    html_links <- paste0(html_links, 
      sprintf('<a href="%s" class="btn">%s</a>', m$arquivo, m$nome)
    )
  }
  
  html_index <- sprintf('
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Painel de Mapas FNP</title>
    <style>
        body { font-family: -apple-system, system-ui, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif; background-color: #f3f4f6; color: #111827; margin: 0; display: flex; flex-direction: column; align-items: center; justify-content: center; min-height: 100vh; }
        .container { background: white; padding: 2rem; border-radius: 1rem; box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1); text-align: center; max-width: 400px; width: 90%%; }
        h1 { font-size: 1.5rem; font-weight: 800; margin-bottom: 0.5rem; color: #1e3a8a; }
        p { margin-bottom: 2rem; color: #6b7280; }
        .btn { display: block; width: 100%%; background-color: #2563eb; color: white; padding: 0.75rem 1rem; margin-bottom: 0.75rem; border-radius: 0.5rem; text-decoration: none; font-weight: 600; transition: background-color 0.2s; box-sizing: border-box; }
        .btn:hover { background-color: #1d4ed8; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Mapas Interativos</h1>
        <p>Selecione uma visualização abaixo:</p>
        %s
    </div>
</body>
</html>', html_links)
  
  writeLines(html_index, file.path(BASE_DIR, "index.html"), useBytes = TRUE)
  
  cat("\n------------------------------------------------------")
  cat("\nSUCESSO! Site gerado com", length(mapas_gerados), "mapas.")
  cat("\n------------------------------------------------------\n")
}

gerar_site_completo()