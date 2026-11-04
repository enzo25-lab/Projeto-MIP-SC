# ==============================================================================
# PIPELINE 01 - PNAD CONTINUA 5a VISITA, 2018-2025
# Projeto Atualizacao MIP-SC
#
# Objetivo:
#   Reproduzir, em um unico script, o fluxo que antes estava dividido em varias
#   etapas: download -> tratamento -> formal/informal -> rendimentos ->
#   salario minimo trimestral -> OCC/FAM -> agregacao -> parcelas -> Excel SEM SCN.
#
# Saida principal:
#   R/PNAD/Resultados Excel/PNADC_2018_2025_SEM_SCN_SM_CORRIGIDO.xlsx
#
# Dependencia do PIPELINE 02:
#   Este script tambem salva os microdados tratados por ano em
#   "Dados PNAD OCC FAM SM Corrigido", pois o PIPELINE 02 precisa de VD4009 para
#   separar educacao/saude publica e privada antes da agregacao SCN.
# ==============================================================================

rm(list = ls())
gc()
options(scipen = 999)

# ==============================================================================
# 0. CONFIGURACAO
# ==============================================================================

# Para usar em outro computador, defina a variavel de ambiente MIP_SC_ROOT.
# Se ela nao existir, usa o caminho atual do projeto.
PASTA_PROJETO <- Sys.getenv(
  "MIP_SC_ROOT",
  unset = "C:/Users/10766048977/Desktop/Projeto atualização da MIP"
)

DIR_PNAD <- file.path(PASTA_PROJETO, "R", "PNAD")
DIR_DOWNLOAD <- file.path(DIR_PNAD, "Dados PNAD Download")
DIR_RDS_BRUTO <- file.path(DIR_PNAD, "Dados PNAD RDS")
DIR_MICRO_CORRIGIDO <- file.path(DIR_PNAD, "Dados PNAD OCC FAM SM Corrigido")
DIR_AGREGADO_CORRIGIDO <- file.path(DIR_PNAD, "Dados PNAD Agregados SM Corrigido")
DIR_PARCELAS_CORRIGIDAS <- file.path(DIR_PNAD, "Dados PNAD Parcelas SM Corrigido")
DIR_RESULTADOS_EXCEL <- file.path(DIR_PNAD, "Resultados Excel")
DIR_LOGS <- file.path(DIR_PNAD, "Logs")

for (d in c(
  DIR_PNAD, DIR_DOWNLOAD, DIR_RDS_BRUTO, DIR_MICRO_CORRIGIDO,
  DIR_AGREGADO_CORRIGIDO, DIR_PARCELAS_CORRIGIDAS,
  DIR_RESULTADOS_EXCEL, DIR_LOGS
)) {
  dir.create(d, recursive = TRUE, showWarnings = FALSE)
}

ANOS <- 2018:2025
BAIXAR_SE_NAO_EXISTIR <- TRUE
VALIDAR_CONTRA_RESULTADO_FINAL <- TRUE

# ==============================================================================
# 1. PACOTES
# ==============================================================================

pacotes <- c("PNADcIBGE", "data.table", "dplyr", "writexl")
faltantes <- pacotes[!pacotes %in% rownames(installed.packages())]
if (length(faltantes) > 0) {
  install.packages(faltantes, dependencies = TRUE)
}
invisible(lapply(pacotes, library, character.only = TRUE))

# ==============================================================================
# 2. FUNCOES AUXILIARES
# ==============================================================================

parar_se <- function(condicao, mensagem) {
  if (isTRUE(condicao)) stop(mensagem, call. = FALSE)
}

normalizar_codigo_5d <- function(x) {
  x <- as.character(x)
  saida <- rep(NA_character_, length(x))
  ok <- !is.na(x) & trimws(x) != ""
  saida[ok] <- sprintf("%05d", as.integer(x[ok]))
  saida
}

mapa_uf <- c(
  "11" = "RO", "12" = "AC", "13" = "AM", "14" = "RR", "15" = "PA",
  "16" = "AP", "17" = "TO", "21" = "MA", "22" = "PI", "23" = "CE",
  "24" = "RN", "25" = "PB", "26" = "PE", "27" = "AL", "28" = "SE",
  "29" = "BA", "31" = "MG", "32" = "ES", "33" = "RJ", "35" = "SP",
  "41" = "PR", "42" = "SC", "43" = "RS", "50" = "MS", "51" = "MT",
  "52" = "GO", "53" = "DF"
)

# Salario minimo usado na classificacao OCC/FAM.
# Nos trimestres com mudanca legal intra-trimestre, usa-se a media mensal.
salarios_minimos <- data.table::data.table(
  Ano = rep(2018:2025, each = 4),
  Trimestre = rep(1:4, times = 8),
  salario_minimo = c(
    954, 954, 954, 954,              # 2018
    998, 998, 998, 998,              # 2019
    1043, 1045, 1045, 1045,          # 2020
    1100, 1100, 1100, 1100,          # 2021
    1212, 1212, 1212, 1212,          # 2022
    1302, 1314, 1320, 1320,          # 2023
    1412, 1412, 1412, 1412,          # 2024
    1518, 1518, 1518, 1518           # 2025
  )
)

# Referencia do arquivo final validado no projeto.
ref_linhas_sem_scn <- c(
  `2018` = 30589,
  `2019` = 31361,
  `2020` = 26690,
  `2021` = 25446,
  `2022` = 28734,
  `2023` = 29081,
  `2024` = 29920,
  `2025` = 29119
)

# ==============================================================================
# 3. DOWNLOAD / CACHE DOS MICRODADOS
# ==============================================================================

for (ano in ANOS) {
  arquivo_bruto <- file.path(
    DIR_RDS_BRUTO,
    paste0("PNADC_", ano, "_5visita_bruta.rds")
  )

  if (file.exists(arquivo_bruto)) {
    cat("[DOWNLOAD]", ano, "- RDS bruto ja existe; usando cache.\n")
    next
  }

  parar_se(
    !BAIXAR_SE_NAO_EXISTIR,
    paste("RDS bruto ausente para", ano, "e download desativado.")
  )

  cat("[DOWNLOAD] Baixando PNAD Continua", ano, "- 5a visita...\n")

  # selected = FALSE foi mantido por ser a forma ja testada no projeto.
  # Isso garante que Trimestre e as demais variaveis necessarias estejam presentes.
  bruto <- PNADcIBGE::get_pnadc(
    year = ano,
    interview = 5,
    selected = FALSE,
    labels = FALSE,
    deflator = FALSE,
    design = FALSE,
    reload = FALSE,
    savedir = DIR_DOWNLOAD
  )

  parar_se(is.null(bruto), paste("Falha no download de", ano))
  bruto <- data.table::as.data.table(bruto)
  saveRDS(bruto, arquivo_bruto, compress = TRUE)
  rm(bruto)
  gc()
}

# ==============================================================================
# 4. PROCESSAMENTO ANO A ANO
# ==============================================================================

lista_agregados <- list()
lista_auditoria <- list()

for (ano in ANOS) {
  cat("\n============================================================\n")
  cat("PROCESSANDO ANO:", ano, "\n")
  cat("============================================================\n")

  arquivo_bruto <- file.path(
    DIR_RDS_BRUTO,
    paste0("PNADC_", ano, "_5visita_bruta.rds")
  )

  dados <- data.table::as.data.table(readRDS(arquivo_bruto))

  obrigatorias <- c(
    "Ano", "Trimestre", "UF", "V4013", "V1032", "VD4002", "VD4009",
    "V4019", "VD4016", "VD4017", "VD4019", "VD4020",
    "VD5004", "VD5007", "VD5008"
  )
  ausentes <- setdiff(obrigatorias, names(dados))
  parar_se(
    length(ausentes) > 0,
    paste("Variaveis ausentes em", ano, ":", paste(ausentes, collapse = ", "))
  )

  # ---------------------------------------------------------------------------
  # 4.1 Tipos e ocupados
  # ---------------------------------------------------------------------------
  dados[, Ano := as.integer(as.character(Ano))]
  dados[, Trimestre := as.integer(as.character(Trimestre))]
  dados[, UF := as.integer(as.character(UF))]
  dados[, VD4002 := as.character(VD4002)]
  dados[, VD4009 := sprintf("%02d", as.integer(as.character(VD4009)))]
  dados[, V4019 := as.character(V4019)]
  dados[, V1032 := as.numeric(V1032)]
  dados[, CNAEDOM := normalizar_codigo_5d(V4013)]

  dados <- dados[VD4002 == "1"]

  # ---------------------------------------------------------------------------
  # 4.2 Formal / informal - proxy oficial usada no projeto
  # 02 = empregado privado sem carteira
  # 04 = domestico sem carteira
  # 10 = trabalhador familiar auxiliar
  # 08/09 + V4019=2 = empregador/conta propria sem CNPJ
  # ---------------------------------------------------------------------------
  dados[, informal := data.table::fifelse(
    VD4009 %in% c("02", "04", "10") |
      (VD4009 %in% c("08", "09") & V4019 == "2"),
    1L,
    0L
  )]

  parar_se(anyNA(dados$informal), paste("Informalidade NA em", ano))
  dados[, formal := 1L - informal]
  dados[, ocupado := 1L]
  dados[, situacao_ocupacao := ifelse(informal == 1L, "Informal", "Formal")]

  parar_se(
    any(dados$formal + dados$informal != 1L),
    paste("Formal + informal diferente de 1 em", ano)
  )

  # ---------------------------------------------------------------------------
  # 4.3 Rendimentos
  # ---------------------------------------------------------------------------
  dados[, renda_ind_princ_habitual := as.numeric(VD4016)]
  dados[, renda_ind_princ_efetiva := as.numeric(VD4017)]
  dados[, renda_ind_total_habitual := as.numeric(VD4019)]
  dados[, renda_ind_total_efetiva := as.numeric(VD4020)]
  dados[, renda_dom_efetiva := as.numeric(VD5004)]
  dados[, renda_dom_habitual := as.numeric(VD5007)]
  dados[, renda_dom_pc_habitual := as.numeric(VD5008)]

  # Mesmos filtros metodologicos usados no resultado final.
  dados <- dados[
    !is.na(renda_ind_princ_efetiva) &
      !is.na(renda_dom_efetiva)
  ]

  # ---------------------------------------------------------------------------
  # 4.4 Salario minimo trimestral + OCC/FAM
  # ---------------------------------------------------------------------------
  dados <- merge(
    dados,
    salarios_minimos,
    by = c("Ano", "Trimestre"),
    all.x = TRUE,
    sort = FALSE
  )

  parar_se(
    anyNA(dados$salario_minimo),
    paste("Salario minimo ausente em", ano)
  )

  dados[, clasOCC := data.table::fcase(
    renda_ind_princ_efetiva <= 2 * salario_minimo, "OCC1",
    renda_ind_princ_efetiva <= 3 * salario_minimo, "OCC2",
    renda_ind_princ_efetiva <= 6 * salario_minimo, "OCC3",
    renda_ind_princ_efetiva <= 10 * salario_minimo, "OCC4",
    renda_ind_princ_efetiva <= 15 * salario_minimo, "OCC5",
    renda_ind_princ_efetiva <= 25 * salario_minimo, "OCC6",
    renda_ind_princ_efetiva > 25 * salario_minimo, "OCC7",
    default = NA_character_
  )]

  dados[, clasFAM := data.table::fcase(
    renda_dom_efetiva <= 2 * salario_minimo, "FAM1",
    renda_dom_efetiva <= 3 * salario_minimo, "FAM2",
    renda_dom_efetiva <= 6 * salario_minimo, "FAM3",
    renda_dom_efetiva <= 10 * salario_minimo, "FAM4",
    renda_dom_efetiva <= 15 * salario_minimo, "FAM5",
    renda_dom_efetiva <= 25 * salario_minimo, "FAM6",
    renda_dom_efetiva > 25 * salario_minimo, "FAM7",
    default = NA_character_
  )]

  parar_se(anyNA(dados$clasOCC), paste("OCC NA em", ano))
  parar_se(anyNA(dados$clasFAM), paste("FAM NA em", ano))
  parar_se(anyNA(dados$CNAEDOM), paste("CNAEDOM NA entre registros validos em", ano))

  # Salvar microdados finais do ano. O PIPELINE 02 le estes arquivos.
  arquivo_micro <- file.path(
    DIR_MICRO_CORRIGIDO,
    paste0("PNADC_", ano, "_5visita_OCC_FAM_SM_corrigido.rds")
  )
  saveRDS(dados, arquivo_micro, compress = TRUE)

  # ---------------------------------------------------------------------------
  # 4.5 Agregacao SEM SCN
  # ---------------------------------------------------------------------------
  agregado <- dados[
    ,
    .(
      n_amostra = .N,
      n_amostra_formal = sum(formal),
      n_amostra_informal = sum(informal),
      TPO_total = sum(V1032, na.rm = TRUE),
      TPO_formal = sum(V1032 * formal, na.rm = TRUE),
      TPO_informal = sum(V1032 * informal, na.rm = TRUE),
      TRP_total = sum(renda_ind_princ_efetiva * V1032, na.rm = TRUE),
      TRP_formal = sum(renda_ind_princ_efetiva * V1032 * formal, na.rm = TRUE),
      TRP_informal = sum(renda_ind_princ_efetiva * V1032 * informal, na.rm = TRUE),
      TRD_total = sum(renda_dom_efetiva * V1032, na.rm = TRUE),
      TRD_formal = sum(renda_dom_efetiva * V1032 * formal, na.rm = TRUE),
      TRD_informal = sum(renda_dom_efetiva * V1032 * informal, na.rm = TRUE)
    ),
    by = .(
      Ano,
      UF_codigo = UF,
      CNAEDOM,
      OCC = clasOCC,
      FAM = clasFAM
    )
  ]

  agregado[, UF := unname(mapa_uf[as.character(UF_codigo)])]
  agregado[, taxa_informalidade := data.table::fifelse(
    TPO_total > 0,
    100 * TPO_informal / TPO_total,
    NA_real_
  )]

  data.table::setcolorder(
    agregado,
    c(
      "Ano", "UF_codigo", "UF", "CNAEDOM", "OCC", "FAM",
      "n_amostra", "n_amostra_formal", "n_amostra_informal",
      "TPO_total", "TPO_formal", "TPO_informal", "taxa_informalidade",
      "TRP_total", "TRP_formal", "TRP_informal",
      "TRD_total", "TRD_formal", "TRD_informal"
    )
  )
  data.table::setorder(agregado, Ano, UF_codigo, CNAEDOM, OCC, FAM)

  # Auditorias de identidade contabil.
  erro_tpo <- max(abs(agregado$TPO_total - agregado$TPO_formal - agregado$TPO_informal))
  erro_trp <- max(abs(agregado$TRP_total - agregado$TRP_formal - agregado$TRP_informal))
  erro_trd <- max(abs(agregado$TRD_total - agregado$TRD_formal - agregado$TRD_informal))

  parar_se(erro_tpo > 1e-6, paste("Falha TPO = formal + informal em", ano))
  parar_se(erro_trp > 1e-4, paste("Falha TRP = formal + informal em", ano))
  parar_se(erro_trd > 1e-4, paste("Falha TRD = formal + informal em", ano))

  duplicadas <- agregado[
    , .N,
    by = .(Ano, UF_codigo, CNAEDOM, OCC, FAM)
  ][N > 1]
  parar_se(nrow(duplicadas) > 0, paste("Chaves duplicadas no agregado de", ano))

  if (VALIDAR_CONTRA_RESULTADO_FINAL) {
    esperado <- unname(ref_linhas_sem_scn[as.character(ano)])
    parar_se(
      nrow(agregado) != esperado,
      paste0(
        "Regressao: ", ano, " gerou ", nrow(agregado),
        " linhas; resultado final validado possui ", esperado, "."
      )
    )
  }

  arquivo_agregado <- file.path(
    DIR_AGREGADO_CORRIGIDO,
    paste0("PNADC_", ano, "_agregada_SM_corrigida.rds")
  )
  saveRDS(agregado, arquivo_agregado, compress = TRUE)

  lista_agregados[[as.character(ano)]] <- agregado
  lista_auditoria[[as.character(ano)]] <- data.frame(
    Ano = ano,
    linhas = nrow(agregado),
    UFs = data.table::uniqueN(agregado$UF_codigo),
    CNAEDOM = data.table::uniqueN(agregado$CNAEDOM),
    TPO_total = sum(agregado$TPO_total),
    TPO_formal = sum(agregado$TPO_formal),
    TPO_informal = sum(agregado$TPO_informal),
    taxa_informalidade = 100 * sum(agregado$TPO_informal) / sum(agregado$TPO_total),
    TRP_total = sum(agregado$TRP_total),
    TRD_total = sum(agregado$TRD_total),
    erro_TPO = erro_tpo,
    erro_TRP = erro_trp,
    erro_TRD = erro_trd
  )

  rm(dados, agregado)
  gc()
}

# ==============================================================================
# 5. CONSOLIDAR 2018-2025
# ==============================================================================

agregado_2018_2025 <- data.table::rbindlist(
  lista_agregados,
  use.names = TRUE,
  fill = TRUE
)
data.table::setorder(agregado_2018_2025, Ano, UF_codigo, CNAEDOM, OCC, FAM)

parar_se(nrow(agregado_2018_2025) != 230940, "Total de linhas SEM SCN diferente de 230.940.")
parar_se(data.table::uniqueN(agregado_2018_2025$CNAEDOM) != 222, "Esperados 222 CNAEDOM.")

saveRDS(
  agregado_2018_2025,
  file.path(DIR_AGREGADO_CORRIGIDO, "PNADC_2018_2025_agregada_SM_corrigida.rds"),
  compress = TRUE
)

# ==============================================================================
# 6. PARCELAS OCC x FAM
# ==============================================================================

parcelas <- agregado_2018_2025[
  ,
  .(
    TRD_total = sum(TRD_total),
    TRD_formal = sum(TRD_formal),
    TRD_informal = sum(TRD_informal)
  ),
  by = .(Ano, UF_codigo, UF, OCC, FAM)
]

parcelas[, parcela_total := {
  z <- sum(TRD_total)
  if (z > 0) TRD_total / z else rep(NA_real_, .N)
}, by = .(Ano, UF_codigo, OCC)]

parcelas[, parcela_formal := {
  z <- sum(TRD_formal)
  if (z > 0) TRD_formal / z else rep(NA_real_, .N)
}, by = .(Ano, UF_codigo, OCC)]

parcelas[, parcela_informal := {
  z <- sum(TRD_informal)
  if (z > 0) TRD_informal / z else rep(NA_real_, .N)
}, by = .(Ano, UF_codigo, OCC)]

data.table::setorder(parcelas, Ano, UF_codigo, OCC, FAM)

# Auditoria das parcelas somente onde existe massa no grupo.
aud_parcelas <- parcelas[
  ,
  .(
    TRD_formal_grupo = sum(TRD_formal),
    TRD_informal_grupo = sum(TRD_informal),
    soma_total = sum(parcela_total, na.rm = TRUE),
    soma_formal = if (sum(TRD_formal) > 0) sum(parcela_formal, na.rm = TRUE) else NA_real_,
    soma_informal = if (sum(TRD_informal) > 0) sum(parcela_informal, na.rm = TRUE) else NA_real_
  ),
  by = .(Ano, UF_codigo, OCC)
]

parar_se(any(abs(aud_parcelas$soma_total - 1) > 1e-10, na.rm = TRUE), "Parcelas totais nao fecham em 1.")
parar_se(any(abs(aud_parcelas$soma_formal - 1) > 1e-10, na.rm = TRUE), "Parcelas formais nao fecham em 1.")
parar_se(any(abs(aud_parcelas$soma_informal - 1) > 1e-10, na.rm = TRUE), "Parcelas informais nao fecham em 1.")

saveRDS(
  parcelas,
  file.path(DIR_PARCELAS_CORRIGIDAS, "PNADC_2018_2025_parcelas_OCC_FAM_SM_corrigido.rds"),
  compress = TRUE
)

# ==============================================================================
# 7. ABAS DO EXCEL SEM SCN
# ==============================================================================

dados_pnad <- data.table::copy(agregado_2018_2025)
dados_sc <- dados_pnad[UF_codigo == 42]

if (VALIDAR_CONTRA_RESULTADO_FINAL) {
  parar_se(nrow(dados_sc) != 15037, "Dados_SC SEM SCN deveria ter 15.037 linhas.")
}

pessoal_ocupado <- dados_pnad[
  , .(
    Ano, UF_codigo, UF, CNAEDOM, OCC, FAM, n_amostra,
    TPO_total, TPO_formal, TPO_informal, taxa_informalidade
  )
]

renda_trabalho <- dados_pnad[
  , .(Ano, UF_codigo, UF, CNAEDOM, OCC, FAM, TRP_total, TRP_formal, TRP_informal)
]

renda_domiciliar <- dados_pnad[
  , .(Ano, UF_codigo, UF, CNAEDOM, OCC, FAM, TRD_total, TRD_formal, TRD_informal)
]

totais_brasil <- dados_pnad[
  , .(
    Escopo = "Brasil",
    TPO_total = sum(TPO_total),
    TPO_formal = sum(TPO_formal),
    TPO_informal = sum(TPO_informal),
    TRP_total = sum(TRP_total),
    TRP_formal = sum(TRP_formal),
    TRP_informal = sum(TRP_informal),
    TRD_total = sum(TRD_total),
    TRD_formal = sum(TRD_formal),
    TRD_informal = sum(TRD_informal)
  ),
  by = Ano
]
totais_brasil[, taxa_informalidade := 100 * TPO_informal / TPO_total]

totais_sc <- dados_sc[
  , .(
    Escopo = "Santa Catarina",
    TPO_total = sum(TPO_total),
    TPO_formal = sum(TPO_formal),
    TPO_informal = sum(TPO_informal),
    TRP_total = sum(TRP_total),
    TRP_formal = sum(TRP_formal),
    TRP_informal = sum(TRP_informal),
    TRD_total = sum(TRD_total),
    TRD_formal = sum(TRD_formal),
    TRD_informal = sum(TRD_informal)
  ),
  by = Ano
]
totais_sc[, taxa_informalidade := 100 * TPO_informal / TPO_total]

totais_ano <- data.table::rbindlist(list(totais_brasil, totais_sc), use.names = TRUE)
data.table::setcolorder(
  totais_ano,
  c(
    "Ano", "Escopo", "TPO_total", "TPO_formal", "TPO_informal",
    "taxa_informalidade", "TRP_total", "TRP_formal", "TRP_informal",
    "TRD_total", "TRD_formal", "TRD_informal"
  )
)

auditoria_dados <- dados_pnad[
  , .(
    linhas = .N,
    ufs = data.table::uniqueN(UF_codigo),
    cnaes = data.table::uniqueN(CNAEDOM),
    diferenca_TPO = sum(TPO_total) - sum(TPO_formal) - sum(TPO_informal),
    diferenca_TRP = sum(TRP_total) - sum(TRP_formal) - sum(TRP_informal),
    diferenca_TRD = sum(TRD_total) - sum(TRD_formal) - sum(TRD_informal)
  ),
  by = Ano
]

auditoria_parcelas_ano <- aud_parcelas[
  , .(
    max_dif_parcela_total = max(abs(soma_total - 1), na.rm = TRUE),
    max_dif_parcela_formal = max(abs(soma_formal - 1), na.rm = TRUE),
    max_dif_parcela_informal = max(abs(soma_informal - 1), na.rm = TRUE),
    grupos_sem_formal = sum(TRD_formal_grupo == 0),
    grupos_sem_informal = sum(TRD_informal_grupo == 0)
  ),
  by = Ano
]

auditoria_excel <- merge(
  auditoria_dados,
  auditoria_parcelas_ano,
  by = "Ano",
  all.x = TRUE,
  sort = TRUE
)

# Dicionario igual ao arquivo corrigido que foi validado.
dicionario <- data.frame(
  Variavel = c(
    "Ano", "UF_codigo", "UF", "CNAEDOM", "OCC", "FAM",
    "n_amostra", "n_amostra_formal", "n_amostra_informal",
    "TPO_total", "TPO_formal", "TPO_informal", "taxa_informalidade",
    "TRP_total", "TRP_formal", "TRP_informal",
    "TRD_total", "TRD_formal", "TRD_informal",
    "parcela_total", "parcela_formal", "parcela_informal",
    "NOTA 1", "NOTA 2", "NOTA 3", "NOTA 4", "NOTA 5"
  ),
  Descricao = c(
    "Ano de referência da PNAD Contínua.",
    "Código numérico da Unidade da Federação.",
    "Sigla da Unidade da Federação.",
    "Código CNAE Domiciliar da atividade do trabalho principal, preservado com 5 dígitos.",
    "Classe de rendimento individual do trabalhador: OCC1 a OCC7.",
    "Classe de rendimento domiciliar: FAM1 a FAM7.",
    "Número de pessoas efetivamente observadas na amostra na célula.",
    "Número de observações classificadas como formais.",
    "Número de observações classificadas como informais.",
    "Estimativa ponderada de pessoas ocupadas.",
    "Estimativa ponderada de pessoas ocupadas formais.",
    "Estimativa ponderada de pessoas ocupadas informais.",
    "Participação percentual dos informais no total de ocupados da célula.",
    "Massa ponderada mensal do rendimento efetivo do trabalho principal.",
    "Parcela da massa de rendimento do trabalho associada aos trabalhadores formais.",
    "Parcela da massa de rendimento do trabalho associada aos trabalhadores informais.",
    "Renda domiciliar efetiva ponderada associada aos trabalhadores da célula.",
    "Renda domiciliar efetiva associada aos trabalhadores formais.",
    "Renda domiciliar efetiva associada aos trabalhadores informais.",
    "Participação da classe FAM na renda domiciliar associada a determinado grupo OCC.",
    "Participação da FAM na renda domiciliar associada aos trabalhadores formais de determinado OCC.",
    "Participação da FAM na renda domiciliar associada aos trabalhadores informais de determinado OCC.",
    "O arquivo utiliza a 5ª visita da PNAD Contínua em todos os anos de 2018 a 2025.",
    "Formal e informal seguem a proxy de informalidade utilizada pelo IBGE.",
    "OCC é construída com VD4017, rendimento efetivamente recebido no trabalho principal.",
    "FAM utiliza VD5004. TRD deve ser interpretada como renda domiciliar associada aos trabalhadores da célula, e não como uma estimativa sem duplicação da renda total dos domicílios.",
    paste0(
      "As classes OCC e FAM utilizam o salário mínimo correspondente ao período. ",
      "Nos anos com alteração intra-anual, foi adotada aproximação trimestral: ",
      "em 2020, R$ 1.043 no 1º trimestre e R$ 1.045 nos demais; ",
      "em 2023, R$ 1.302 no 1º trimestre, R$ 1.314 no 2º trimestre ",
      "e R$ 1.320 no 3º e 4º trimestres."
    )
  ),
  stringsAsFactors = FALSE
)

# ==============================================================================
# 8. EXPORTAR EXCEL SEM SCN
# ==============================================================================

arquivo_excel <- file.path(
  DIR_RESULTADOS_EXCEL,
  "PNADC_2018_2025_SEM_SCN_SM_CORRIGIDO.xlsx"
)

writexl::write_xlsx(
  list(
    Dados_PNAD = as.data.frame(dados_pnad),
    Dados_SC = as.data.frame(dados_sc),
    Pessoal_Ocupado = as.data.frame(pessoal_ocupado),
    Renda_Trabalho = as.data.frame(renda_trabalho),
    Renda_Domiciliar = as.data.frame(renda_domiciliar),
    Parcelas_OCC_FAM = as.data.frame(parcelas),
    Totais_Ano = as.data.frame(totais_ano),
    Auditoria = as.data.frame(auditoria_excel),
    Dicionario = dicionario
  ),
  path = arquivo_excel
)

# ==============================================================================
# 9. AUDITORIA FINAL E LOGS
# ==============================================================================

auditoria_pipeline <- data.table::rbindlist(lista_auditoria, fill = TRUE)
data.table::fwrite(
  auditoria_pipeline,
  file.path(DIR_LOGS, "AUDITORIA_pipeline_01_PNADC_dados.csv")
)

capture.output(
  sessionInfo(),
  file = file.path(DIR_LOGS, "sessionInfo_pipeline_01_PNADC.txt")
)

cat("\n============================================================\n")
cat("PIPELINE 01 CONCLUIDO\n")
cat("============================================================\n")
cat("Linhas consolidadas SEM SCN:", nrow(dados_pnad), "\n")
cat("CNAEDOM distintos:", data.table::uniqueN(dados_pnad$CNAEDOM), "\n")
cat("Dados_SC SEM SCN:", nrow(dados_sc), "linhas\n")
cat("Excel:\n", arquivo_excel, "\n", sep = "")
cat("\nExecute em seguida: 02_pipeline_PNADC_CNAE_SCN_excel_final.R\n")
