# ==============================================================================
# Massa salarial x atividade CNAE x dados municipais- RAIS SC (2018-2025)
# ==============================================================================
# Objetivo:
#   Reproduzir, do download dos microdados da RAIS até a exportação em Excel,
#   a base anual agregada por Município x Classe CNAE para Santa Catarina.
#
# Saída principal:
#   Resultados RAIS/Excel Final/RAIS_SC_2018_2025_FINAL.xlsx
#
# O pipeline executa:
#   1) checagem/instalação dos pacotes;
#   2) download e extração da RAIS Sul (2018-2025);
#   3) download e construção do tradutor territorial DTB/IBGE 2025;
#   4) leitura apenas das colunas necessárias da RAIS;
#   5) filtro do município do estabelecimento para Santa Catarina;
#   6) harmonização dos layouts 2018-2022 e 2023-2025;
#   7) reclassificação de Saúde Pública (861911) e Educação Pública (851911);
#   8) harmonização das remunerações mensais;
#   9) correção de escala das remunerações jan-nov em 2018-2021;
#  10) cálculo da massa salarial anual;
#  11) filtro rem_media_sm > 0;
#  12) padronização da CNAE e tratamento de códigos residuais;
#  13) agregação por Município x CNAE;
#  14) inclusão das Regiões Geográficas Imediata e Intermediária;
#  15) auditoria interna e comparação com o benchmark validado;
#  16) exportação do Excel final e dos arquivos RDS anuais.
#
# Fontes:
#   RAIS/MTE:
#   ftp://ftp.mtps.gov.br/pdet/microdados/RAIS/<ANO>/RAIS_VINC_PUB_SUL.7z
#
#   DTB/IBGE 2025:
#   https://geoftp.ibge.gov.br/organizacao_do_territorio/estrutura_territorial/
#   divisao_territorial/2025/DTB_2025.zip
#
# IMPORTANTE:
#   - Execute este script tendo como diretório de trabalho a pasta raiz do
#     projeto RAIS. Se preferir, defina a variável de ambiente MIPSC_RAIS_DIR.
#   - O script NÃO usa o tradutor CNAE -> SCN/MIP. Essa etapa permanece fora
#     deste pipeline, conforme decisão metodológica do projeto.
#   - O benchmark ao final serve para confirmar se uma nova execução reproduz
#     a versão validada da base. Mudanças posteriores nos microdados oficiais
#     podem gerar divergências, por isso o benchmark gera aviso, não altera os
#     dados e, por padrão, não interrompe o processamento.
# ============================================================================== 

rm(list = ls())
options(stringsAsFactors = FALSE)
options(timeout = 7200)

# ==============================================================================
# 0. CONFIGURAÇÕES DO USUÁRIO
# ==============================================================================

ANOS <- 2018:2025

# Se MIPSC_RAIS_DIR estiver definida, ela será usada. Caso contrário, usa getwd().
PASTA_PROJETO <- Sys.getenv("MIPSC_RAIS_DIR", unset = getwd())
PASTA_PROJETO <- normalizePath(PASTA_PROJETO, winslash = "/", mustWork = TRUE)

# Baixa novamente arquivos que já existem? Normalmente, deixe FALSE.
FORCAR_DOWNLOAD <- FALSE

# Salvar uma base individual de SC, já harmonizada, para auditoria/reuso?
# Pode ocupar bastante espaço. A base final agregada sempre será salva.
SALVAR_BASE_INDIVIDUAL_SC <- FALSE

# Incluir a aba de auditoria dentro do Excel final?
# FALSE reproduz a versão final limpa que foi mantida no projeto.
INCLUIR_ABA_AUDITORIA_EXCEL <- FALSE

# Instalar automaticamente pacotes ausentes pelo CRAN?
INSTALAR_PACOTES_AUSENTES <- TRUE

# Comparar o resultado com a execução final já validada?
VERIFICAR_BENCHMARK <- TRUE

# Se TRUE, qualquer divergência do benchmark interrompe o script.
# Recomenda-se FALSE, pois o MTE pode substituir/revisar microdados oficiais.
PARAR_SE_BENCHMARK_DIVERGIR <- FALSE

# Tolerância monetária para comparação com o benchmark (em R$).
TOLERANCIA_MASSA <- 0.05

cat("Diretório do projeto:\n", PASTA_PROJETO, "\n")

# ==============================================================================
# 1. PACOTES
# ==============================================================================

pacotes <- c(
  "data.table",
  "archive",
  "readODS",
  "openxlsx"
)

faltantes <- pacotes[
  !vapply(pacotes, requireNamespace, logical(1), quietly = TRUE)
]

if (length(faltantes) > 0) {
  if (!INSTALAR_PACOTES_AUSENTES) {
    stop(
      "Pacotes ausentes: ",
      paste(faltantes, collapse = ", "),
      ". Instale-os antes de executar o pipeline."
    )
  }

  cat("Instalando pacotes ausentes:", paste(faltantes, collapse = ", "), "\n")
  install.packages(faltantes, dependencies = TRUE)
}

suppressPackageStartupMessages({
  library(data.table)
  library(archive)
  library(readODS)
  library(openxlsx)
})

# ==============================================================================
# 2. ESTRUTURA DE PASTAS
# ==============================================================================

PASTA_RAIS_SUL <- file.path(PASTA_PROJETO, "Dados RAIS Sul")
PASTA_IBGE <- file.path(PASTA_PROJETO, "Dados territoriais IBGE")
PASTA_RESULTADOS <- file.path(PASTA_PROJETO, "Resultados RAIS")
PASTA_RDS_FINAL <- file.path(PASTA_RESULTADOS, "Bases finais RDS")
PASTA_EXCEL <- file.path(PASTA_RESULTADOS, "Excel Final")
PASTA_AUDITORIA <- file.path(PASTA_RESULTADOS, "Auditoria")
PASTA_SC_INDIVIDUAL <- file.path(PASTA_RESULTADOS, "Bases individuais SC")

pastas <- c(
  PASTA_RAIS_SUL,
  PASTA_IBGE,
  PASTA_RESULTADOS,
  PASTA_RDS_FINAL,
  PASTA_EXCEL,
  PASTA_AUDITORIA
)

if (SALVAR_BASE_INDIVIDUAL_SC) {
  pastas <- c(pastas, PASTA_SC_INDIVIDUAL)
}

invisible(lapply(pastas, dir.create, recursive = TRUE, showWarnings = FALSE))

# ==============================================================================
# 3. FUNÇÕES AUXILIARES
# ==============================================================================

normaliza_nome <- function(x) {
  y <- iconv(x, from = "", to = "ASCII//TRANSLIT")
  y[is.na(y)] <- x[is.na(y)]
  y <- tolower(y)
  y <- gsub("[^a-z0-9]+", "_", y)
  y <- gsub("^_+|_+$", "", y)
  y
}

somente_digitos <- function(x) {
  y <- trimws(as.character(x))
  y <- gsub("[^0-9]", "", y)
  y[y == ""] <- NA_character_
  y
}

mensagem_etapa <- function(texto) {
  cat("\n", paste(rep("=", 78), collapse = ""), "\n", sep = "")
  cat(texto, "\n")
  cat(paste(rep("=", 78), collapse = ""), "\n", sep = "")
}

localizar_arquivo_rais_extraido <- function(pasta_ano) {
  arquivos <- list.files(
    pasta_ano,
    pattern = "RAIS_VINC_PUB_SUL\\.(txt|comt)$",
    full.names = TRUE,
    recursive = TRUE,
    ignore.case = TRUE
  )

  arquivos <- unique(arquivos)

  if (length(arquivos) == 0) {
    return(NA_character_)
  }

  if (length(arquivos) > 1) {
    stop(
      "Mais de um arquivo RAIS_VINC_PUB_SUL foi encontrado em: ",
      pasta_ano,
      "\nArquivos:\n",
      paste(arquivos, collapse = "\n")
    )
  }

  arquivos[1]
}

obter_arquivo_rais <- function(ano) {
  pasta_ano <- file.path(PASTA_RAIS_SUL, paste0("RAIS_", ano))
  dir.create(pasta_ano, recursive = TRUE, showWarnings = FALSE)

  arquivo_extraido <- localizar_arquivo_rais_extraido(pasta_ano)

  if (!FORCAR_DOWNLOAD && !is.na(arquivo_extraido)) {
    cat("RAIS", ano, "já está extraída:", basename(arquivo_extraido), "\n")
    return(arquivo_extraido)
  }

  url <- sprintf(
    "ftp://ftp.mtps.gov.br/pdet/microdados/RAIS/%d/RAIS_VINC_PUB_SUL.7z",
    ano
  )

  arquivo_7z <- file.path(pasta_ano, "RAIS_VINC_PUB_SUL.7z")

  if (FORCAR_DOWNLOAD || !file.exists(arquivo_7z)) {
    cat("Baixando RAIS", ano, "...\n")
    utils::download.file(
      url = url,
      destfile = arquivo_7z,
      mode = "wb",
      quiet = FALSE
    )
  } else {
    cat("Arquivo .7z da RAIS", ano, "já existe.\n")
  }

  cat("Extraindo RAIS", ano, "...\n")
  archive::archive_extract(arquivo_7z, dir = pasta_ano)

  arquivo_extraido <- localizar_arquivo_rais_extraido(pasta_ano)

  if (is.na(arquivo_extraido)) {
    stop("A RAIS ", ano, " foi baixada, mas o arquivo extraído não foi localizado.")
  }

  arquivo_extraido
}

aliases_rais <- function(ano) {
  if (ano <= 2022) {
    mapa <- c(
      municipio = "municipio",
      cnae_classe = "cnae_2_0_classe",
      cnae_subclasse = "cnae_2_0_subclasse",
      cbo_2002 = "cbo_ocupacao_2002",
      natureza_juridica = "natureza_juridica",
      vinculo_ativo_3112 = "vinculo_ativo_31_12",
      rem_media_nom = "vl_remun_media_nom",
      rem_media_sm = "vl_remun_media_sm",
      rem_dez = "vl_remun_dezembro_nom"
    )
  } else {
    mapa <- c(
      municipio = "municipio_codigo",
      cnae_classe = "cnae_2_0_classe_codigo",
      cnae_subclasse = "cnae_2_0_subclasse_codigo",
      cbo_2002 = "cbo_2002_ocupacao_codigo",
      natureza_juridica = "natureza_juridica_codigo",
      vinculo_ativo_3112 = "ind_vinculo_ativo_31_12_codigo",
      rem_media_nom = "vl_rem_media_nom",
      rem_media_sm = "vl_rem_media_sm",
      rem_dez = "vl_rem_dezembro_nom"
    )
  }

  sufixo_mensal <- if (ano <= 2019) "cc" else "sc"

  meses_origem <- c(
    jan = "janeiro",
    fev = "fevereiro",
    mar = "marco",
    abr = "abril",
    mai = "maio",
    jun = "junho",
    jul = "julho",
    ago = "agosto",
    set = "setembro",
    out = "outubro",
    nov = "novembro"
  )

  mapa_mensal <- setNames(
    paste0("vl_rem_", meses_origem, "_", sufixo_mensal),
    paste0("rem_", names(meses_origem))
  )

  c(mapa, mapa_mensal)
}

mapear_colunas_rais <- function(arquivo, ano) {
  aliases <- aliases_rais(ano)
  tentativas <- c("Latin-1", "UTF-8")
  diagnosticos <- list()

  for (codificacao in tentativas) {
    cabecalho <- tryCatch(
      data.table::fread(
        arquivo,
        nrows = 0,
        encoding = codificacao,
        showProgress = FALSE,
        check.names = FALSE
      ),
      error = function(e) NULL
    )

    if (is.null(cabecalho)) next

    nomes_originais <- names(cabecalho)
    nomes_normalizados <- normaliza_nome(nomes_originais)

    indices <- match(unname(aliases), nomes_normalizados)
    names(indices) <- names(aliases)

    # Fallback específico para o município do estabelecimento.
    if (is.na(indices["municipio"])) {
      candidatos <- which(
        grepl("^mun", nomes_normalizados) &
          !grepl("trab", nomes_normalizados)
      )
      if (length(candidatos) == 1) {
        indices["municipio"] <- candidatos
      }
    }

    diagnosticos[[codificacao]] <- data.frame(
      original = nomes_originais,
      normalizado = nomes_normalizados,
      stringsAsFactors = FALSE
    )

    if (all(!is.na(indices))) {
      colunas_originais <- nomes_originais[indices]
      names(colunas_originais) <- names(aliases)

      return(list(
        encoding = codificacao,
        colunas = colunas_originais
      ))
    }
  }

  faltantes <- names(aliases)[is.na(indices)]
  stop(
    "Não foi possível mapear todas as colunas da RAIS ", ano, ".\n",
    "Variáveis faltantes: ", paste(faltantes, collapse = ", "), "\n",
    "Revise o layout do arquivo: ", arquivo
  )
}

ler_rais_sul_selecionada <- function(arquivo, ano) {
  mapa <- mapear_colunas_rais(arquivo, ano)

  cat("Codificação escolhida:", mapa$encoding, "\n")
  cat("Lendo apenas as", length(mapa$colunas), "colunas necessárias...\n")

  dado <- data.table::fread(
    arquivo,
    select = unname(mapa$colunas),
    encoding = mapa$encoding,
    dec = ",",
    data.table = TRUE,
    showProgress = TRUE,
    check.names = FALSE,
    na.strings = c("", "NA")
  )

  data.table::setnames(
    dado,
    old = unname(mapa$colunas),
    new = names(mapa$colunas)
  )

  dado
}

padronizar_cnae <- function(cnae, ano, cnae_subclasse) {
  x <- trimws(as.character(cnae))
  x <- gsub("\\s+", "", x)
  x <- gsub("[^0-9]", "", x)
  x[x == ""] <- NA_character_

  # Classes CNAE antigas que perderam o zero à esquerda ao serem lidas como número.
  idx4 <- !is.na(x) & nchar(x) == 4
  x[idx4] <- paste0("0", x[idx4])

  # Residual conhecido na RAIS 2022.
  idx977 <- !is.na(x) & x == "977"
  if (any(idx977)) {
    if (ano != 2022) {
      stop("Código CNAE 977 encontrado fora de 2022. Revisão manual necessária.")
    }

    subclasse <- somente_digitos(cnae_subclasse)
    subclasses_977 <- unique(subclasse[idx977 & !is.na(subclasse)])

    if (length(subclasses_977) > 0 && !all(subclasses_977 == "9999997")) {
      stop(
        "CNAE 977 em 2022 apareceu com subclasse diferente de 9999997: ",
        paste(subclasses_977, collapse = ", ")
      )
    }

    x[idx977] <- "NAO_IDENTIFICADA"
  }

  # Código residual encontrado na série final em 2024/2025.
  # O tratamento é aplicado em qualquer ano caso o código reapareça.
  x[!is.na(x) & x == "99999"] <- "NAO_IDENTIFICADA"

  x
}

validar_cnae_final <- function(x) {
  !is.na(x) & (
    grepl("^[0-9]{5}$", x) |
      x %in% c("851911", "861911", "NAO_IDENTIFICADA")
  )
}

# ==============================================================================
# 4. DTB/IBGE 2025 - DOWNLOAD E TRADUTOR TERRITORIAL
# ==============================================================================

construir_tradutor_territorial <- function() {
  mensagem_etapa("DTB/IBGE 2025 - tradutor territorial")

  url_dtb <- paste0(
    "https://geoftp.ibge.gov.br/organizacao_do_territorio/estrutura_territorial/",
    "divisao_territorial/2025/DTB_2025.zip"
  )

  arquivo_zip <- file.path(PASTA_IBGE, "DTB_2025.zip")
  pasta_dtb <- file.path(PASTA_IBGE, "DTB_2025")

  dir.create(pasta_dtb, recursive = TRUE, showWarnings = FALSE)

  if (FORCAR_DOWNLOAD || !file.exists(arquivo_zip)) {
    cat("Baixando DTB 2025...\n")
    utils::download.file(url_dtb, arquivo_zip, mode = "wb")
  } else {
    cat("DTB_2025.zip já existe.\n")
  }

  arquivo_municipios <- list.files(
    pasta_dtb,
    pattern = "RELATORIO_DTB_BRASIL_2025_MUNICIPIOS\\.ods$",
    full.names = TRUE,
    recursive = TRUE,
    ignore.case = TRUE
  )

  if (FORCAR_DOWNLOAD || length(arquivo_municipios) == 0) {
    cat("Extraindo DTB 2025...\n")
    utils::unzip(arquivo_zip, exdir = pasta_dtb)

    arquivo_municipios <- list.files(
      pasta_dtb,
      pattern = "RELATORIO_DTB_BRASIL_2025_MUNICIPIOS\\.ods$",
      full.names = TRUE,
      recursive = TRUE,
      ignore.case = TRUE
    )
  }

  if (length(arquivo_municipios) != 1) {
    stop(
      "Era esperado exatamente um RELATORIO_DTB_BRASIL_2025_MUNICIPIOS.ods. ",
      "Foram encontrados ", length(arquivo_municipios), "."
    )
  }

  regs <- readODS::read_ods(
    arquivo_municipios[1],
    sheet = 1,
    skip = 6
  )

  regs <- data.table::as.data.table(regs)
  data.table::setnames(regs, normaliza_nome(names(regs)))

  colunas_dtb <- c(
    "uf",
    "codigo_municipio_completo",
    "nome_municipio",
    "nome_regiao_geografica_intermediaria",
    "nome_regiao_geografica_imediata"
  )

  faltantes <- setdiff(colunas_dtb, names(regs))
  if (length(faltantes) > 0) {
    stop(
      "Colunas faltantes na DTB 2025: ",
      paste(faltantes, collapse = ", ")
    )
  }

  regs <- regs[as.character(uf) == "42"]

  codigo_completo <- somente_digitos(regs$codigo_municipio_completo)

  tradutor <- data.table(
    municipio = substr(codigo_completo, 1, 6),
    nome_municipio = as.character(regs$nome_municipio),
    regiao_intermediaria = as.character(regs$nome_regiao_geografica_intermediaria),
    regiao_imediata = as.character(regs$nome_regiao_geografica_imediata)
  )

  tradutor <- unique(tradutor)
  data.table::setorder(tradutor, municipio)

  if (nrow(tradutor) != 295) {
    stop("O tradutor territorial deveria ter 295 municípios, mas possui ", nrow(tradutor), ".")
  }

  if (anyDuplicated(tradutor$municipio) > 0) {
    stop("Há códigos municipais duplicados no tradutor territorial.")
  }

  if (anyNA(tradutor)) {
    stop("Há valores ausentes no tradutor territorial.")
  }

  saveRDS(
    tradutor,
    file.path(PASTA_IBGE, "tradutor_municipios_SC_DTB2025.rds")
  )

  cat("Tradutor territorial validado: 295 municípios.\n")
  tradutor
}

TRADUTOR_MUNICIPIOS <- construir_tradutor_territorial()

# ==============================================================================
# 5. BENCHMARK DA EXECUÇÃO FINAL VALIDADA
# ==============================================================================

benchmark <- data.table(
  ano = ANOS,
  registros_sc = c(
    3425141L, 3550760L, 3581751L, 4026255L,
    4336073L, 4483134L, 4753966L, 4932279L
  ),
  saude_publica = c(
    48707L, 50367L, 54449L, 58667L,
    67921L, 68724L, 71242L, 73101L
  ),
  educacao_publica = c(
    178773L, 182715L, 157547L, 201324L,
    201165L, 221988L, 235835L, 251942L
  ),
  correcoes_salario = c(
    4015844L, 3743354L, 3738621L, 4527983L,
    0L, 0L, 0L, 0L
  ),
  linhas_final = c(
    35465L, 35401L, 35455L, 36602L,
    37877L, 38024L, 38264L, 38344L
  ),
  municipios = rep(295L, 8),
  cnaes_distintas = c(641L, 638L, 638L, 635L, 638L, 636L, 638L, 636L),
  massa_salarial = c(
    72727048229.26,
    74451201046.36,
    75060712763.03,
    86928053340.56,
    101809517561.33,
    116419229717.38,
    128002686821.32,
    139359473944.90
  ),
  vinculos_total = c(
    3326296L, 3353055L, 3393788L, 3832393L,
    4034826L, 4163312L, 4388872L, 4528941L
  ),
  vinculos_ativos = c(
    2198438L, 2231133L, 2263171L, 2421384L,
    2513418L, 2586964L, 2673383L, 2737473L
  ),
  linhas_nao_identificada = c(0L, 0L, 0L, 0L, 63L, 0L, 1L, 1L)
)

# ==============================================================================
# 6. PROCESSAMENTO ANO A ANO
# ==============================================================================

lista_auditoria <- vector("list", length(ANOS))
names(lista_auditoria) <- as.character(ANOS)

rem_jan_nov <- c(
  "rem_jan", "rem_fev", "rem_mar", "rem_abr", "rem_mai", "rem_jun",
  "rem_jul", "rem_ago", "rem_set", "rem_out", "rem_nov"
)

rem_12_meses <- c(rem_jan_nov, "rem_dez")

for (ano in ANOS) {
  mensagem_etapa(paste("PROCESSANDO RAIS", ano))

  arquivo_bruto <- obter_arquivo_rais(ano)
  cat("Arquivo utilizado:", arquivo_bruto, "\n")

  dado <- ler_rais_sul_selecionada(arquivo_bruto, ano)
  registros_sul_selecionados <- nrow(dado)

  # ---------------------------------------------------------------------------
  # 6.1 Padronização básica e filtro territorial
  # ---------------------------------------------------------------------------

  dado[, municipio := somente_digitos(municipio)]
  dado[, cnae_classe := as.character(cnae_classe)]
  dado[, cnae_subclasse := as.character(cnae_subclasse)]
  dado[, cbo_2002 := somente_digitos(cbo_2002)]
  dado[, natureza_juridica := somente_digitos(natureza_juridica)]
  dado[, vinculo_ativo_3112 := suppressWarnings(as.integer(vinculo_ativo_3112))]
  dado[, rem_media_nom := suppressWarnings(as.numeric(rem_media_nom))]
  dado[, rem_media_sm := suppressWarnings(as.numeric(rem_media_sm))]

  for (col in rem_12_meses) {
    data.table::set(
      dado,
      j = col,
      value = suppressWarnings(as.numeric(dado[[col]]))
    )
  }

  dado <- dado[!is.na(municipio) & substr(municipio, 1, 2) == "42"]
  dado[, ano := ano]

  registros_sc <- nrow(dado)

  if (registros_sc == 0) {
    stop("Nenhum registro de Santa Catarina foi encontrado em ", ano, ".")
  }

  if (anyNA(dado$municipio) || !all(substr(dado$municipio, 1, 2) == "42")) {
    stop("Falha no filtro territorial de Santa Catarina em ", ano, ".")
  }

  cat(
    "Registros lidos (Sul, colunas selecionadas): ",
    format(registros_sul_selecionados, big.mark = "."),
    "\nRegistros de SC: ",
    format(registros_sc, big.mark = "."),
    "\n",
    sep = ""
  )

  # ---------------------------------------------------------------------------
  # 6.2 Reclassificação de Saúde Pública e Educação Pública
  # ---------------------------------------------------------------------------

  dado[, cnae_classe_original := cnae_classe]

  natureza_num <- suppressWarnings(as.numeric(dado$natureza_juridica))
  cnae_num <- suppressWarnings(as.numeric(gsub("[^0-9]", "", dado$cnae_classe_original)))
  cbo_txt <- somente_digitos(dado$cbo_2002)

  saude_publica <- (
    !is.na(natureza_num) &
      natureza_num < 2011 &
      (
        substr(cbo_txt, 1, 2) == "22" |
          (!is.na(cnae_num) & cnae_num >= 86100 & cnae_num <= 87309) |
          substr(cbo_txt, 1, 3) %in% as.character(322:328) |
          substr(cbo_txt, 1, 4) %in% c("5151", "5152")
      )
  )

  educacao_publica <- (
    !is.na(natureza_num) &
      natureza_num < 2011 &
      (
        substr(cbo_txt, 1, 2) == "23" |
          (!is.na(cnae_num) & cnae_num >= 85110 & cnae_num <= 85999) |
          substr(cbo_txt, 1, 2) == "33"
      )
  )

  # Reproduz a prioridade do ifelse aninhado do script original: Saúde primeiro.
  educacao_publica <- educacao_publica & !saude_publica

  dado[saude_publica == TRUE, cnae_classe := "861911"]
  dado[educacao_publica == TRUE, cnae_classe := "851911"]

  n_saude <- sum(saude_publica, na.rm = TRUE)
  n_educacao <- sum(educacao_publica, na.rm = TRUE)

  cat(
    "Saúde pública reclassificada: ", format(n_saude, big.mark = "."),
    "\nEducação pública reclassificada: ", format(n_educacao, big.mark = "."),
    "\n",
    sep = ""
  )

  rm(natureza_num, cnae_num, cbo_txt, saude_publica, educacao_publica)

  # ---------------------------------------------------------------------------
  # 6.3 Padronização definitiva da CNAE
  # ---------------------------------------------------------------------------

  dado[, cnae_classe := padronizar_cnae(cnae_classe, ano, cnae_subclasse)]

  cnae_valida <- validar_cnae_final(dado$cnae_classe)
  if (any(!cnae_valida)) {
    invalidas <- sort(unique(dado$cnae_classe[!cnae_valida]))
    stop(
      "CNAEs não tratadas em ", ano, ": ",
      paste(invalidas, collapse = ", ")
    )
  }

  # ---------------------------------------------------------------------------
  # 6.4 Correção de escala das remunerações mensais em 2018-2021
  # ---------------------------------------------------------------------------

  n_corrigidos_salario <- 0L

  if (ano <= 2021) {
    for (col in rem_jan_nov) {
      idx <- !is.na(dado[[col]]) & dado[[col]] > 0 & dado[[col]] < 100
      n_corrigidos_salario <- n_corrigidos_salario + sum(idx)

      if (any(idx)) {
        data.table::set(
          dado,
          i = which(idx),
          j = col,
          value = dado[[col]][idx] * 100
        )
      }
    }
  }

  cat(
    "Valores mensais com escala corrigida: ",
    format(n_corrigidos_salario, big.mark = "."),
    "\n",
    sep = ""
  )

  # ---------------------------------------------------------------------------
  # 6.5 Massa salarial anual por vínculo
  # ---------------------------------------------------------------------------

  dado[, massa_salarial := rowSums(.SD, na.rm = TRUE), .SDcols = rem_12_meses]

  # Opcional: salva a base individual de SC já tratada.
  if (SALVAR_BASE_INDIVIDUAL_SC) {
    arquivo_individual <- file.path(
      PASTA_SC_INDIVIDUAL,
      paste0("RAIS_", ano, "_SC_individual_tratada.rds")
    )
    saveRDS(dado, arquivo_individual)
  }

  # ---------------------------------------------------------------------------
  # 6.6 Filtro metodológico e agregação Município x CNAE
  # ---------------------------------------------------------------------------

  dado_valido <- dado[!is.na(rem_media_sm) & rem_media_sm > 0]

  resultado <- dado_valido[, .(
    massa_salarial = sum(massa_salarial, na.rm = TRUE),
    vinculos_total = .N,
    vinculos_ativos_3112 = sum(vinculo_ativo_3112 == 1L, na.rm = TRUE)
  ), by = .(ano, municipio, cnae_classe)]

  # ---------------------------------------------------------------------------
  # 6.7 Informações territoriais - join somente após a agregação
  # ---------------------------------------------------------------------------

  resultado <- merge(
    resultado,
    TRADUTOR_MUNICIPIOS,
    by = "municipio",
    all.x = TRUE,
    sort = FALSE
  )

  data.table::setcolorder(
    resultado,
    c(
      "ano",
      "municipio",
      "nome_municipio",
      "regiao_intermediaria",
      "regiao_imediata",
      "cnae_classe",
      "massa_salarial",
      "vinculos_total",
      "vinculos_ativos_3112"
    )
  )

  data.table::setorder(resultado, municipio, cnae_classe)

  # ---------------------------------------------------------------------------
  # 6.8 Auditoria interna do ano
  # ---------------------------------------------------------------------------

  if (anyNA(resultado$nome_municipio) ||
      anyNA(resultado$regiao_imediata) ||
      anyNA(resultado$regiao_intermediaria)) {
    stop("Há município sem correspondência territorial em ", ano, ".")
  }

  if (anyDuplicated(resultado[, .(ano, municipio, cnae_classe)]) > 0) {
    stop("Há duplicações Município x CNAE após a agregação em ", ano, ".")
  }

  if (any(resultado$massa_salarial < 0, na.rm = TRUE)) {
    stop("Há massa salarial negativa em ", ano, ".")
  }

  if (any(resultado$vinculos_total < 0, na.rm = TRUE) ||
      any(resultado$vinculos_ativos_3112 < 0, na.rm = TRUE)) {
    stop("Há número de vínculos negativo em ", ano, ".")
  }

  if (any(resultado$vinculos_ativos_3112 > resultado$vinculos_total, na.rm = TRUE)) {
    stop("Há vínculos ativos maiores que vínculos totais em ", ano, ".")
  }

  if (data.table::uniqueN(resultado$municipio) != 295L) {
    stop("A base agregada de ", ano, " não contém os 295 municípios de SC.")
  }

  linhas_nao_identificada <- sum(resultado$cnae_classe == "NAO_IDENTIFICADA")

  lista_auditoria[[as.character(ano)]] <- data.table(
    ano = ano,
    registros_sc = registros_sc,
    saude_publica = n_saude,
    educacao_publica = n_educacao,
    correcoes_salario = n_corrigidos_salario,
    registros_remuneracao_positiva = nrow(dado_valido),
    linhas_final = nrow(resultado),
    municipios = data.table::uniqueN(resultado$municipio),
    cnaes_distintas = data.table::uniqueN(resultado$cnae_classe),
    massa_salarial = sum(resultado$massa_salarial, na.rm = TRUE),
    vinculos_total = sum(resultado$vinculos_total, na.rm = TRUE),
    vinculos_ativos = sum(resultado$vinculos_ativos_3112, na.rm = TRUE),
    linhas_nao_identificada = linhas_nao_identificada
  )

  # ---------------------------------------------------------------------------
  # 6.9 Salva a base final anual
  # ---------------------------------------------------------------------------

  arquivo_rds_final <- file.path(
    PASTA_RDS_FINAL,
    paste0("RAIS_", ano, "_SC_municipio_CNAE_final.rds")
  )

  saveRDS(resultado, arquivo_rds_final)

  cat(
    "Base final salva: ", arquivo_rds_final,
    "\nCombinações Município x CNAE: ", format(nrow(resultado), big.mark = "."),
    "\n",
    sep = ""
  )

  rm(dado, dado_valido, resultado, cnae_valida)
  gc()
}

# ==============================================================================
# 7. AUDITORIA CONSOLIDADA E BENCHMARK
# ==============================================================================

mensagem_etapa("AUDITORIA CONSOLIDADA")

auditoria <- data.table::rbindlist(lista_auditoria, use.names = TRUE, fill = TRUE)
data.table::setorder(auditoria, ano)

print(auditoria)

arquivo_auditoria_csv <- file.path(PASTA_AUDITORIA, "auditoria_pipeline_rais_sc_2018_2025.csv")
data.table::fwrite(auditoria, arquivo_auditoria_csv, sep = ";", dec = ",", bom = TRUE)
saveRDS(auditoria, file.path(PASTA_AUDITORIA, "auditoria_pipeline_rais_sc_2018_2025.rds"))

if (VERIFICAR_BENCHMARK) {
  comparacao <- merge(
    auditoria,
    benchmark,
    by = "ano",
    suffixes = c("_obtido", "_esperado"),
    all.x = TRUE,
    sort = TRUE
  )

  comparacao[, ok_registros_sc := registros_sc_obtido == registros_sc_esperado]
  comparacao[, ok_saude := saude_publica_obtido == saude_publica_esperado]
  comparacao[, ok_educacao := educacao_publica_obtido == educacao_publica_esperado]
  comparacao[, ok_correcao_salario := correcoes_salario_obtido == correcoes_salario_esperado]
  comparacao[, ok_linhas := linhas_final_obtido == linhas_final_esperado]
  comparacao[, ok_municipios := municipios_obtido == municipios_esperado]
  comparacao[, ok_cnaes := cnaes_distintas_obtido == cnaes_distintas_esperado]
  comparacao[, ok_massa := abs(massa_salarial_obtido - massa_salarial_esperado) <= TOLERANCIA_MASSA]
  comparacao[, ok_vinculos := vinculos_total_obtido == vinculos_total_esperado]
  comparacao[, ok_ativos := vinculos_ativos_obtido == vinculos_ativos_esperado]
  comparacao[, ok_residual := linhas_nao_identificada_obtido == linhas_nao_identificada_esperado]

  colunas_ok <- grep("^ok_", names(comparacao), value = TRUE)
  comparacao[, benchmark_confere := Reduce(`&`, .SD), .SDcols = colunas_ok]

  arquivo_benchmark_csv <- file.path(PASTA_AUDITORIA, "comparacao_benchmark_rais_sc_2018_2025.csv")
  data.table::fwrite(comparacao, arquivo_benchmark_csv, sep = ";", dec = ",", bom = TRUE)

  cat("\nComparação com a execução final validada:\n")
  print(comparacao[, c("ano", colunas_ok, "benchmark_confere"), with = FALSE])

  if (!all(comparacao$benchmark_confere)) {
    msg <- paste0(
      "A execução atual divergiu do benchmark validado em pelo menos um teste. ",
      "Consulte: ", arquivo_benchmark_csv
    )

    if (PARAR_SE_BENCHMARK_DIVERGIR) {
      stop(msg)
    } else {
      warning(msg, call. = FALSE)
    }
  } else {
    cat("\nBENCHMARK: todos os anos reproduziram a versão validada.\n")
  }
}

# ==============================================================================
# 8. RESUMO PARA O EXCEL
# ==============================================================================

resumo <- auditoria[, .(
  ano,
  municipios,
  cnaes_distintas,
  combinacoes_municipio_cnae = linhas_final,
  massa_salarial,
  vinculos_total,
  vinculos_ativos_3112 = vinculos_ativos
)]

resumo[, variacao_massa_salarial := massa_salarial / shift(massa_salarial) - 1]
resumo[, variacao_vinculos := vinculos_total / shift(vinculos_total) - 1]
resumo[, variacao_ativos := vinculos_ativos_3112 / shift(vinculos_ativos_3112) - 1]

# ==============================================================================
# 9. EXPORTAÇÃO DO EXCEL FINAL
# ==============================================================================

mensagem_etapa("EXPORTAÇÃO DO EXCEL FINAL")

arquivo_excel <- file.path(PASTA_EXCEL, "RAIS_SC_2018_2025_FINAL.xlsx")
wb <- openxlsx::createWorkbook(creator = "Projeto atualização da MIP-SC")

estilo_titulo <- openxlsx::createStyle(
  fontSize = 14,
  textDecoration = "bold",
  halign = "left"
)

estilo_subtitulo <- openxlsx::createStyle(
  fontSize = 10,
  textDecoration = "italic",
  wrapText = TRUE
)

estilo_cabecalho <- openxlsx::createStyle(
  textDecoration = "bold",
  halign = "center",
  valign = "center",
  border = "Bottom",
  wrapText = TRUE
)

estilo_texto <- openxlsx::createStyle(numFmt = "@")
estilo_inteiro <- openxlsx::createStyle(numFmt = "#,##0")
estilo_monetario <- openxlsx::createStyle(numFmt = "#,##0.00")
estilo_percentual <- openxlsx::createStyle(numFmt = "0.00%")
estilo_wrap <- openxlsx::createStyle(wrapText = TRUE, valign = "top")

# ------------------------------------------------------------------------------
# 9.1 Aba Resumo
# ------------------------------------------------------------------------------

openxlsx::addWorksheet(wb, "Resumo", gridLines = FALSE)
openxlsx::writeData(wb, "Resumo", "RAIS Santa Catarina - 2018 a 2025", startRow = 1, startCol = 1)
openxlsx::addStyle(wb, "Resumo", estilo_titulo, rows = 1, cols = 1)

openxlsx::writeData(
  wb,
  "Resumo",
  "Agregação por município e classe CNAE. Massa salarial em reais correntes.",
  startRow = 2,
  startCol = 1
)
openxlsx::addStyle(
  wb,
  "Resumo",
  estilo_subtitulo,
  rows = 2,
  cols = 1:ncol(resumo),
  gridExpand = TRUE
)

openxlsx::writeData(
  wb,
  "Resumo",
  as.data.frame(resumo),
  startRow = 4,
  startCol = 1,
  withFilter = TRUE
)
openxlsx::addStyle(
  wb,
  "Resumo",
  estilo_cabecalho,
  rows = 4,
  cols = 1:ncol(resumo),
  gridExpand = TRUE
)

linhas_resumo <- 5:(nrow(resumo) + 4)

openxlsx::addStyle(
  wb,
  "Resumo",
  estilo_monetario,
  rows = linhas_resumo,
  cols = which(names(resumo) == "massa_salarial"),
  gridExpand = TRUE
)

openxlsx::addStyle(
  wb,
  "Resumo",
  estilo_inteiro,
  rows = linhas_resumo,
  cols = which(names(resumo) %in% c(
    "municipios",
    "cnaes_distintas",
    "combinacoes_municipio_cnae",
    "vinculos_total",
    "vinculos_ativos_3112"
  )),
  gridExpand = TRUE
)

openxlsx::addStyle(
  wb,
  "Resumo",
  estilo_percentual,
  rows = linhas_resumo,
  cols = which(names(resumo) %in% c(
    "variacao_massa_salarial",
    "variacao_vinculos",
    "variacao_ativos"
  )),
  gridExpand = TRUE
)

openxlsx::freezePane(wb, "Resumo", firstActiveRow = 5)
openxlsx::setColWidths(wb, "Resumo", cols = 1, widths = 10)
openxlsx::setColWidths(wb, "Resumo", cols = 2:ncol(resumo), widths = 22)

# ------------------------------------------------------------------------------
# 9.2 Aba Metodologia
# ------------------------------------------------------------------------------

openxlsx::addWorksheet(wb, "Metodologia", gridLines = FALSE)
openxlsx::writeData(wb, "Metodologia", "Notas metodológicas", startRow = 1, startCol = 1)
openxlsx::addStyle(wb, "Metodologia", estilo_titulo, rows = 1, cols = 1)

metodologia <- data.frame(
  Item = c(
    "Fonte dos dados",
    "Período",
    "Recorte territorial",
    "Unidade de agregação",
    "Informações territoriais",
    "Massa salarial",
    "Correção das remunerações 2018-2021",
    "Vínculos considerados",
    "Vínculos ativos",
    "CNAE",
    "Educação pública",
    "Saúde pública",
    "CNAE não identificada - 2022",
    "CNAE não identificada - 2024/2025",
    "Categoria residual",
    "Valores monetários"
  ),
  Descricao = c(
    "Relação Anual de Informações Sociais (RAIS), Ministério do Trabalho e Emprego.",
    "2018 a 2025.",
    "Estado de Santa Catarina; utiliza-se o município do estabelecimento.",
    "Município x classe CNAE x ano.",
    paste(
      "Nome do município, Região Geográfica Imediata e Região Geográfica Intermediária",
      "obtidos da Divisão Territorial Brasileira (DTB/IBGE) 2025."
    ),
    "Soma das remunerações nominais mensais de janeiro a dezembro dos vínculos considerados.",
    paste(
      "Para janeiro a novembro de 2018 a 2021, valores positivos inferiores a 100",
      "foram multiplicados por 100 após diagnóstico da escala das variáveis.",
      "Dezembro não recebeu essa correção."
    ),
    paste(
      "As agregações de massa salarial e vínculos consideram registros com",
      "remuneração média em salários mínimos (rem_media_sm) maior que zero."
    ),
    paste(
      "Vínculos com indicador de vínculo ativo em 31/12 igual a 1 dentro do conjunto",
      "de registros com rem_media_sm > 0."
    ),
    paste(
      "Classe CNAE armazenada como texto. Códigos de quatro dígitos que haviam perdido",
      "o zero à esquerda foram padronizados para cinco dígitos."
    ),
    "Código auxiliar 851911 utilizado para Educação Pública conforme o procedimento metodológico adotado.",
    "Código auxiliar 861911 utilizado para Saúde Pública conforme o procedimento metodológico adotado.",
    paste(
      "Na RAIS 2022, registros cuja CNAE original apareceu como 977 e cuja subclasse",
      "era 9999997 foram classificados como NAO_IDENTIFICADA."
    ),
    paste(
      "O código residual 99999 encontrado em 2024 e 2025 foi classificado como",
      "NAO_IDENTIFICADA, preservando massa salarial e vínculos nos totais."
    ),
    paste(
      "NAO_IDENTIFICADA é uma categoria residual mantida para evitar imputação arbitrária",
      "de atividade econômica."
    ),
    "Valores monetários expressos em reais correntes do respectivo ano."
  ),
  stringsAsFactors = FALSE
)

openxlsx::writeData(wb, "Metodologia", metodologia, startRow = 3, startCol = 1)
openxlsx::addStyle(wb, "Metodologia", estilo_cabecalho, rows = 3, cols = 1:2, gridExpand = TRUE)
openxlsx::addStyle(
  wb,
  "Metodologia",
  estilo_wrap,
  rows = 4:(nrow(metodologia) + 3),
  cols = 1:2,
  gridExpand = TRUE
)
openxlsx::setColWidths(wb, "Metodologia", cols = 1, widths = 35)
openxlsx::setColWidths(wb, "Metodologia", cols = 2, widths = 100)
openxlsx::setRowHeights(wb, "Metodologia", rows = 4:(nrow(metodologia) + 3), heights = 45)
openxlsx::freezePane(wb, "Metodologia", firstActiveRow = 4)

# ------------------------------------------------------------------------------
# 9.3 Aba Auditoria (opcional)
# ------------------------------------------------------------------------------

if (INCLUIR_ABA_AUDITORIA_EXCEL) {
  openxlsx::addWorksheet(wb, "Auditoria", gridLines = FALSE)
  openxlsx::writeData(wb, "Auditoria", "Auditoria final do pipeline RAIS", startRow = 1, startCol = 1)
  openxlsx::addStyle(wb, "Auditoria", estilo_titulo, rows = 1, cols = 1)
  openxlsx::writeData(wb, "Auditoria", as.data.frame(auditoria), startRow = 3, startCol = 1, withFilter = TRUE)
  openxlsx::addStyle(
    wb,
    "Auditoria",
    estilo_cabecalho,
    rows = 3,
    cols = 1:ncol(auditoria),
    gridExpand = TRUE
  )
  openxlsx::setColWidths(wb, "Auditoria", cols = 1:ncol(auditoria), widths = "auto")
  openxlsx::freezePane(wb, "Auditoria", firstActiveRow = 4)
}

# ------------------------------------------------------------------------------
# 9.4 Abas anuais
# ------------------------------------------------------------------------------

for (ano in ANOS) {
  cat("Criando aba", ano, "...\n")

  arquivo_rds <- file.path(
    PASTA_RDS_FINAL,
    paste0("RAIS_", ano, "_SC_municipio_CNAE_final.rds")
  )

  dado_final <- readRDS(arquivo_rds)
  dado_final <- data.table::as.data.table(dado_final)

  # Garante texto no Excel para códigos com zeros à esquerda.
  dado_final[, municipio := as.character(municipio)]
  dado_final[, cnae_classe := as.character(cnae_classe)]

  nome_aba <- as.character(ano)
  openxlsx::addWorksheet(wb, nome_aba, gridLines = FALSE)

  openxlsx::writeDataTable(
    wb,
    sheet = nome_aba,
    x = as.data.frame(dado_final),
    startRow = 1,
    startCol = 1,
    tableName = paste0("RAIS_SC_", ano),
    tableStyle = "TableStyleMedium2",
    withFilter = TRUE
  )

  linhas_dados <- 2:(nrow(dado_final) + 1)

  openxlsx::freezePane(wb, nome_aba, firstRow = TRUE)
  openxlsx::addStyle(wb, nome_aba, estilo_texto, rows = linhas_dados, cols = 2, gridExpand = TRUE)
  openxlsx::addStyle(wb, nome_aba, estilo_texto, rows = linhas_dados, cols = 6, gridExpand = TRUE)
  openxlsx::addStyle(wb, nome_aba, estilo_monetario, rows = linhas_dados, cols = 7, gridExpand = TRUE)
  openxlsx::addStyle(wb, nome_aba, estilo_inteiro, rows = linhas_dados, cols = 8:9, gridExpand = TRUE)

  openxlsx::setColWidths(wb, nome_aba, cols = 1, widths = 9)
  openxlsx::setColWidths(wb, nome_aba, cols = 2, widths = 12)
  openxlsx::setColWidths(wb, nome_aba, cols = 3, widths = 25)
  openxlsx::setColWidths(wb, nome_aba, cols = 4, widths = 24)
  openxlsx::setColWidths(wb, nome_aba, cols = 5, widths = 30)
  openxlsx::setColWidths(wb, nome_aba, cols = 6, widths = 19)
  openxlsx::setColWidths(wb, nome_aba, cols = 7, widths = 21)
  openxlsx::setColWidths(wb, nome_aba, cols = 8:9, widths = 20)

  rm(dado_final)
  gc()
}

openxlsx::saveWorkbook(wb, arquivo_excel, overwrite = TRUE)

# ==============================================================================
# 10. CONFERÊNCIA FINAL DO ARQUIVO GERADO
# ==============================================================================

if (!file.exists(arquivo_excel)) {
  stop("O arquivo Excel final não foi criado.")
}

cat("\nArquivo Excel criado com sucesso:\n", arquivo_excel, "\n", sep = "")
cat(
  "Tamanho: ",
  round(file.info(arquivo_excel)$size / 1024^2, 2),
  " MB\n",
  sep = ""
)

cat("\nAbas criadas:\n")
print(names(wb))

mensagem_etapa("PIPELINE CONCLUÍDO")
cat(
  "Produtos principais:\n",
  "- Excel: ", arquivo_excel, "\n",
  "- RDS anuais: ", PASTA_RDS_FINAL, "\n",
  "- Auditoria: ", arquivo_auditoria_csv, "\n",
  sep = ""
)
