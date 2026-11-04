# ==============================================================================
# PIPELINE 02 - PNAD 2018-2025: CNAEDOM -> CNAE 2.0 -> SCN + EXCEL FINAL
# Projeto Atualizacao MIP-SC
#
# Pre-requisito:
#   Executar antes o arquivo 01_pipeline_PNADC_2018_2025_dados.R
#
# Objetivos:
#   1) Ler o tradutor original do professor;
#   2) construir o tradutor metodologico usado em Dados_PNAD/Dados_SC;
#   3) construir separadamente o "Tradutor atualizado" usado pela tabela PNAD SC,
#      preservando as convencoes do professor;
#   4) separar educacao e saude publica/privada usando VD4009 no microdado;
#   5) gerar o Excel final no mesmo desenho do arquivo fechado no projeto;
#   6) executar auditorias e testes de regressao antes de salvar.
#
# Saida principal:
#   R/PNAD/Resultados Excel/PNADC_2018_2025_FINAL_CNAE_SCN.xlsx
# ==============================================================================

rm(list = ls())
gc()
options(scipen = 999)

# ==============================================================================
# 0. CONFIGURACAO
# ==============================================================================

PASTA_PROJETO <- Sys.getenv(
  "MIP_SC_ROOT",
  unset = "C:/Users/10766048977/Desktop/Projeto atualização da MIP"
)

DIR_PNAD <- file.path(PASTA_PROJETO, "R", "PNAD")
DIR_MICRO_CORRIGIDO <- file.path(DIR_PNAD, "Dados PNAD OCC FAM SM Corrigido")
DIR_AGREGADO_CORRIGIDO <- file.path(DIR_PNAD, "Dados PNAD Agregados SM Corrigido")
DIR_RESULTADOS_EXCEL <- file.path(DIR_PNAD, "Resultados Excel")
DIR_TRADUTOR <- file.path(DIR_PNAD, "Tradutor")
DIR_LOGS <- file.path(DIR_PNAD, "Logs")

for (d in c(DIR_RESULTADOS_EXCEL, DIR_TRADUTOR, DIR_LOGS)) {
  dir.create(d, recursive = TRUE, showWarnings = FALSE)
}

ANOS <- 2018:2025
VALIDAR_CONTRA_RESULTADO_FINAL <- TRUE

# Caminho preferencial. Para outro computador, pode-se definir PNAD_TRADUTOR_XLSX.
ARQUIVO_TRADUTOR <- Sys.getenv(
  "PNAD_TRADUTOR_XLSX",
  unset = file.path(DIR_TRADUTOR, "Tradutor_CNAEDOM_SCN_professor.xlsx")
)

# Se o nome padrao nao existir, aceita automaticamente um unico xlsx na pasta Tradutor.
if (!file.exists(ARQUIVO_TRADUTOR)) {
  candidatos <- list.files(
    DIR_TRADUTOR,
    pattern = "\\.xlsx$",
    full.names = TRUE,
    ignore.case = TRUE
  )
  if (length(candidatos) == 1) {
    ARQUIVO_TRADUTOR <- candidatos[1]
  }
}

if (!file.exists(ARQUIVO_TRADUTOR)) {
  stop(
    paste0(
      "Tradutor nao encontrado. Coloque o arquivo original do professor em:\n",
      file.path(DIR_TRADUTOR, "Tradutor_CNAEDOM_SCN_professor.xlsx"),
      "\nou defina a variavel de ambiente PNAD_TRADUTOR_XLSX."
    ),
    call. = FALSE
  )
}

# ==============================================================================
# 1. PACOTES
# ==============================================================================

pacotes <- c("data.table", "dplyr", "readxl", "openxlsx")
faltantes <- pacotes[!pacotes %in% rownames(installed.packages())]
if (length(faltantes) > 0) {
  install.packages(faltantes, dependencies = TRUE)
}
invisible(lapply(pacotes, library, character.only = TRUE))

# ==============================================================================
# 2. FUNCOES E REFERENCIAS
# ==============================================================================

parar_se <- function(condicao, mensagem) {
  if (isTRUE(condicao)) stop(mensagem, call. = FALSE)
}

col_letra <- function(n) {
  letras <- character(0)
  while (n > 0) {
    resto <- (n - 1) %% 26
    letras <- c(LETTERS[resto + 1], letras)
    n <- (n - 1) %/% 26
  }
  paste0(letras, collapse = "")
}

normalizar_vazio <- function(x) {
  x <- trimws(as.character(x))
  x[x %in% c("", "#N/A", "NA", "N/A")] <- NA_character_
  x
}

normalizar_codigo_5d <- function(x) {
  x <- as.character(x)
  out <- rep(NA_character_, length(x))
  ok <- !is.na(x) & trimws(x) != ""
  out[ok] <- sprintf("%05d", as.integer(x[ok]))
  out
}

mapa_uf <- c(
  "11" = "RO", "12" = "AC", "13" = "AM", "14" = "RR", "15" = "PA",
  "16" = "AP", "17" = "TO", "21" = "MA", "22" = "PI", "23" = "CE",
  "24" = "RN", "25" = "PB", "26" = "PE", "27" = "AL", "28" = "SE",
  "29" = "BA", "31" = "MG", "32" = "ES", "33" = "RJ", "35" = "SP",
  "41" = "PR", "42" = "SC", "43" = "RS", "50" = "MS", "51" = "MT",
  "52" = "GO", "53" = "DF"
)

# Resultados finais validados no projeto - usados como teste de regressao.
ref_linhas_final_scn <- c(
  `2018` = 31976,
  `2019` = 32786,
  `2020` = 27866,
  `2021` = 26544,
  `2022` = 29980,
  `2023` = 30353,
  `2024` = 31254,
  `2025` = 30445
)

ref_sc <- data.frame(
  Ano = 2018:2025,
  TPO_total = c(
    3618368.427142114, 3715724.010693646, 3650196.110942547,
    3751329.020245892, 3968684.565933999, 4129448.855939473,
    4304311.477805644, 4392960.098665712
  ),
  TPO_informal = c(
    932102.300192039, 927704.297934421, 891117.243363730,
    905661.799195190, 1046662.662938559, 1055927.476136260,
    1093870.353688889, 1029753.704660440
  ),
  TRP_total = c(
    8905465021.828259, 9218122426.621334, 9284214537.438492,
    10205855040.180338, 12309450133.038265, 13695125669.592205,
    15529979358.394962, 17762532577.710236
  )
)

# ==============================================================================
# 3. LER E LIMPAR O TRADUTOR DO PROFESSOR
# ==============================================================================

abas_tradutor <- readxl::excel_sheets(ARQUIVO_TRADUTOR)

aba_origem <- if ("CNAE Dom x SCN" %in% abas_tradutor) {
  "CNAE Dom x SCN"
} else if ("Tradutor" %in% abas_tradutor) {
  "Tradutor"
} else if ("Tradutor atualizado" %in% abas_tradutor) {
  "Tradutor atualizado"
} else {
  stop("Nao encontrei uma aba de tradutor reconhecivel no arquivo informado.", call. = FALSE)
}

# Snapshot bruto para a aba Tradutor do Excel final.
tradutor_snapshot <- readxl::read_excel(
  ARQUIVO_TRADUTOR,
  sheet = aba_origem,
  col_names = FALSE,
  col_types = "text"
) %>% as.data.frame(stringsAsFactors = FALSE)

# A limpeza nao depende da posicao exata dos cabecalhos: mantem apenas linhas cujo
# primeiro campo seja um CNAEDOM de 5 digitos.
tradutor_raw <- tradutor_snapshot[, seq_len(min(8, ncol(tradutor_snapshot))), drop = FALSE]
while (ncol(tradutor_raw) < 8) tradutor_raw[[ncol(tradutor_raw) + 1]] <- NA_character_

linhas_validas <- grepl("^[0-9]{5}$", trimws(as.character(tradutor_raw[[1]])))
tradutor_professor <- tradutor_raw[linhas_validas, 1:8, drop = FALSE]

names(tradutor_professor) <- c(
  "CNAEDOM", "descricao_CNAEDOM", "CNAE20", "descricao_CNAE20",
  "SCN_atividade", "descricao_SCN", "SCN_produto", "descricao_produto_SCN"
)

tradutor_professor <- tradutor_professor %>%
  mutate(across(everything(), normalizar_vazio)) %>%
  mutate(CNAEDOM = normalizar_codigo_5d(CNAEDOM)) %>%
  arrange(CNAEDOM)

parar_se(anyDuplicated(tradutor_professor$CNAEDOM) > 0, "CNAEDOM duplicado no tradutor do professor.")

# ==============================================================================
# 4. TRADUTOR ATUALIZADO - VERSAO USADA NA ABA PNAD SC
#
# IMPORTANTE: esta versao reproduz o tradutor antigo/professor e aplica SOMENTE as
# atualizacoes manuais aprovadas para a tabelinha PNAD SC. Ela nao e a mesma coisa
# que o tradutor metodologico da base Dados_PNAD.
# ==============================================================================

tradutor_atualizado <- tradutor_professor

tradutor_atualizado <- tradutor_atualizado %>%
  mutate(
    SCN_atividade = ifelse(CNAEDOM == "01119", "0191", SCN_atividade),
    descricao_SCN = ifelse(
      CNAEDOM == "01119",
      "Agricultura, inclusive o apoio à agricultura e a pós-colheita",
      descricao_SCN
    )
  )

novos_tradutor_manual <- data.frame(
  CNAEDOM = c("19010", "29003", "32002"),
  descricao_CNAEDOM = c(
    "Coquerias",
    "Reconstrução, em fábrica, de motores de veículos automotores",
    "Fabricação de instrumentos musicais"
  ),
  CNAE20 = c("1910", "2950", "3220"),
  descricao_CNAE20 = c(
    "Coquerias",
    "Recondicionamento e recuperação de motores para veículos automotores",
    "Fabricação de instrumentos musicais"
  ),
  SCN_atividade = c("1991", "2992", "31802"),
  descricao_SCN = c(
    "Refino de petróleo e coquerias",
    "Fabricação de peças e acessórios para veículos automotores",
    "Fabricação de produtos de indústrias diversas"
  ),
  SCN_produto = c("19916", "29921", "31802"),
  descricao_produto_SCN = c(
    "Outros produtos do refino do petróleo",
    "Peças e acessórios para veículos automotores",
    "Produtos de industrias diversas"
  ),
  stringsAsFactors = FALSE
)

# Acrescentar apenas se ausentes na fonte.
tradutor_atualizado <- tradutor_atualizado %>%
  filter(!CNAEDOM %in% novos_tradutor_manual$CNAEDOM) %>%
  bind_rows(novos_tradutor_manual) %>%
  arrange(CNAEDOM)

# Auditoria: entre os codigos que ja existiam, a unica alteracao intencional e 01119.
comparacao_manual <- tradutor_professor %>%
  select(CNAEDOM, SCN_antigo = SCN_atividade) %>%
  inner_join(
    tradutor_atualizado %>% select(CNAEDOM, SCN_novo = SCN_atividade),
    by = "CNAEDOM"
  ) %>%
  mutate(
    antigo_cmp = ifelse(is.na(SCN_antigo), "<NA>", SCN_antigo),
    novo_cmp = ifelse(is.na(SCN_novo), "<NA>", SCN_novo)
  ) %>%
  filter(antigo_cmp != novo_cmp)

parar_se(
  !all(comparacao_manual$CNAEDOM %in% "01119"),
  paste(
    "O Tradutor atualizado alterou codigos antigos alem de 01119:",
    paste(comparacao_manual$CNAEDOM, collapse = ", ")
  )
)

# ==============================================================================
# 5. TRADUTOR METODOLOGICO - USADO EM Dados_PNAD / Dados_SC
# ==============================================================================

tradutor_metodologico <- tradutor_professor %>%
  mutate(
    SCN_atividade = case_when(
      CNAEDOM %in% c("10010", "10030") ~ "1091",
      CNAEDOM %in% c("10021", "10022", "10091", "10093", "10099") ~ "1093",
      CNAEDOM %in% c("31000", "32001", "32003", "32009") ~ "3180",
      CNAEDOM == "01113" ~ "0191",
      CNAEDOM == "01119" ~ "0191",
      CNAEDOM == "01209" ~ "0192",
      TRUE ~ SCN_atividade
    ),
    descricao_SCN = case_when(
      SCN_atividade == "1091" ~ "Abate e produtos de carne, inclusive os produtos do laticínio e da pesca",
      SCN_atividade == "1093" ~ "Outros produtos alimentares",
      SCN_atividade == "3180" ~ "Fabricação de móveis e de produtos de indústrias diversas",
      CNAEDOM %in% c("01113", "01119") ~ "Agricultura, inclusive o apoio à agricultura e a pós-colheita",
      CNAEDOM == "01209" ~ "Pecuária, inclusive o apoio à pecuária",
      TRUE ~ descricao_SCN
    )
  )

complementares_metodologia <- data.frame(
  CNAEDOM = c("19010", "29003", "32002", "35022"),
  descricao_CNAEDOM = c(
    "Coquerias",
    "Recondicionamento e recuperação de motores para veículos automotores",
    "Fabricação de instrumentos musicais",
    "Produção e distribuição de vapor, água quente e ar condicionado"
  ),
  CNAE20 = c("1910", "2950", "3220", "3530"),
  descricao_CNAE20 = c(
    "Coquerias",
    "Recondicionamento e recuperação de motores para veículos automotores",
    "Fabricação de instrumentos musicais",
    "Produção e distribuição de vapor, água quente e ar condicionado"
  ),
  SCN_atividade = c("1991", "2992", "3180", "3500"),
  descricao_SCN = c(
    "Refino de petróleo e coquerias",
    "Fabricação de peças e acessórios para veículos automotores",
    "Fabricação de móveis e de produtos de indústrias diversas",
    "Energia elétrica, gás natural e outras utilidades"
  ),
  SCN_produto = NA_character_,
  descricao_produto_SCN = NA_character_,
  stringsAsFactors = FALSE
)

tradutor_metodologico <- tradutor_metodologico %>%
  filter(!CNAEDOM %in% complementares_metodologia$CNAEDOM) %>%
  bind_rows(complementares_metodologia) %>%
  arrange(CNAEDOM)

parar_se(anyDuplicated(tradutor_metodologico$CNAEDOM) > 0, "Duplicacao no tradutor metodologico.")

# ==============================================================================
# 6. BASE SEM SCN - REFERENCIA PARA PRESERVACAO DOS TOTAIS
# ==============================================================================

arquivo_sem_scn_rds <- file.path(
  DIR_AGREGADO_CORRIGIDO,
  "PNADC_2018_2025_agregada_SM_corrigida.rds"
)

parar_se(!file.exists(arquivo_sem_scn_rds), "Execute primeiro o PIPELINE 01.")
base_sem_scn <- as.data.frame(readRDS(arquivo_sem_scn_rds))

codigos_pnad <- sort(unique(as.character(base_sem_scn$CNAEDOM)))
sem_tradutor <- setdiff(codigos_pnad, tradutor_metodologico$CNAEDOM)
parar_se(
  length(sem_tradutor) > 0,
  paste("CNAEDOM da PNAD ausentes no tradutor metodologico:", paste(sem_tradutor, collapse = ", "))
)

parar_se(length(codigos_pnad) != 222, "A base de referencia deveria conter 222 CNAEDOM.")

# ==============================================================================
# 7. PROCESSAR MICRODADOS COM CNAE/SCN
# ==============================================================================

lista_final <- list()

for (ano in ANOS) {
  cat("[SCN] Processando", ano, "...\n")

  arquivo_micro <- file.path(
    DIR_MICRO_CORRIGIDO,
    paste0("PNADC_", ano, "_5visita_OCC_FAM_SM_corrigido.rds")
  )
  parar_se(!file.exists(arquivo_micro), paste("Microdado corrigido ausente para", ano))

  base <- as.data.frame(readRDS(arquivo_micro))

  obrigatorias <- c(
    "Ano", "UF", "V4013", "VD4009", "V1032", "formal", "informal", "ocupado",
    "renda_ind_princ_efetiva", "renda_dom_efetiva", "clasOCC", "clasFAM"
  )
  faltantes <- setdiff(obrigatorias, names(base))
  parar_se(length(faltantes) > 0, paste("Variaveis ausentes em", ano, ":", paste(faltantes, collapse = ", ")))

  base <- base %>%
    mutate(
      Ano = as.integer(as.character(Ano)),
      UF_codigo = as.integer(as.character(UF)),
      UF = unname(mapa_uf[as.character(UF_codigo)]),
      CNAEDOM = normalizar_codigo_5d(V4013),
      VD4009 = sprintf("%02d", as.integer(as.character(VD4009))),
      OCC = as.character(clasOCC),
      FAM = as.character(clasFAM)
    ) %>%
    left_join(
      tradutor_metodologico %>%
        select(CNAEDOM, CNAE20, descricao_CNAE20, SCN_atividade, descricao_SCN),
      by = "CNAEDOM"
    ) %>%
    mutate(
      SCN_atividade = case_when(
        substr(CNAEDOM, 1, 2) == "85" & VD4009 %in% c("05", "06", "07") ~ "8591",
        substr(CNAEDOM, 1, 2) == "85" ~ "8592",
        substr(CNAEDOM, 1, 2) %in% c("86", "87") & VD4009 %in% c("05", "06", "07") ~ "8691",
        substr(CNAEDOM, 1, 2) %in% c("86", "87") ~ "8692",
        CNAEDOM == "88000" ~ "8692",
        TRUE ~ SCN_atividade
      ),
      descricao_SCN = case_when(
        SCN_atividade == "8591" ~ "Educação pública",
        SCN_atividade == "8592" ~ "Educação privada",
        SCN_atividade == "8691" ~ "Saúde pública",
        SCN_atividade == "8692" ~ "Saúde privada",
        TRUE ~ descricao_SCN
      )
    )

  resultado <- base %>%
    group_by(
      Ano, UF_codigo, UF, CNAEDOM,
      CNAE20, descricao_CNAE20, SCN_atividade, descricao_SCN,
      OCC, FAM
    ) %>%
    summarise(
      n_amostra = n(),
      n_amostra_formal = sum(formal == 1, na.rm = TRUE),
      n_amostra_informal = sum(informal == 1, na.rm = TRUE),
      TPO_total = sum(ocupado * V1032, na.rm = TRUE),
      TPO_formal = sum(ifelse(formal == 1, ocupado * V1032, 0), na.rm = TRUE),
      TPO_informal = sum(ifelse(informal == 1, ocupado * V1032, 0), na.rm = TRUE),
      TRP_total = sum(renda_ind_princ_efetiva * V1032, na.rm = TRUE),
      TRP_formal = sum(ifelse(formal == 1, renda_ind_princ_efetiva * V1032, 0), na.rm = TRUE),
      TRP_informal = sum(ifelse(informal == 1, renda_ind_princ_efetiva * V1032, 0), na.rm = TRUE),
      TRD_total = sum(renda_dom_efetiva * V1032, na.rm = TRUE),
      TRD_formal = sum(ifelse(formal == 1, renda_dom_efetiva * V1032, 0), na.rm = TRUE),
      TRD_informal = sum(ifelse(informal == 1, renda_dom_efetiva * V1032, 0), na.rm = TRUE),
      .groups = "drop"
    ) %>%
    mutate(
      taxa_informalidade = ifelse(TPO_total > 0, 100 * TPO_informal / TPO_total, NA_real_)
    ) %>%
    select(
      Ano, UF_codigo, UF, CNAEDOM,
      CNAE20, descricao_CNAE20, SCN_atividade, descricao_SCN,
      OCC, FAM,
      n_amostra, n_amostra_formal, n_amostra_informal,
      TPO_total, TPO_formal, TPO_informal, taxa_informalidade,
      TRP_total, TRP_formal, TRP_informal,
      TRD_total, TRD_formal, TRD_informal
    ) %>%
    arrange(Ano, UF_codigo, CNAEDOM, SCN_atividade, OCC, FAM)

  # 00000: nao imputar; tornar a ausencia explicita no Excel.
  idx_00000 <- resultado$CNAEDOM == "00000"
  resultado$CNAE20[idx_00000] <- "Não identificado"
  resultado$descricao_CNAE20[idx_00000] <- "Não identificado"
  resultado$SCN_atividade[idx_00000] <- "Não identificado"
  resultado$descricao_SCN[idx_00000] <- "Não identificado"

  if (VALIDAR_CONTRA_RESULTADO_FINAL) {
    esperado <- unname(ref_linhas_final_scn[as.character(ano)])
    parar_se(
      nrow(resultado) != esperado,
      paste0("Regressao SCN: ", ano, " gerou ", nrow(resultado), " linhas; esperado = ", esperado)
    )
  }

  lista_final[[as.character(ano)]] <- resultado
  rm(base, resultado)
  gc()
}

PNAD_FINAL <- bind_rows(lista_final) %>%
  arrange(Ano, UF_codigo, CNAEDOM, SCN_atividade, OCC, FAM)

DADOS_SC <- PNAD_FINAL %>% filter(UF_codigo == 42)

parar_se(nrow(PNAD_FINAL) != 241204, "Resultado final deveria ter 241.204 linhas.")
parar_se(nrow(DADOS_SC) != 15647, "Dados_SC final deveria ter 15.647 linhas.")
parar_se(n_distinct(PNAD_FINAL$CNAEDOM) != 222, "Resultado final deveria preservar 222 CNAEDOM.")
parar_se(n_distinct(DADOS_SC$CNAEDOM) != 216, "Santa Catarina deveria conter 216 CNAEDOM.")

# ==============================================================================
# 8. AUDITORIA: PRESERVACAO DOS TOTAIS APOS TRADUCAO / DESDOBRAMENTO
# ==============================================================================

aud_scn <- PNAD_FINAL %>%
  group_by(Ano) %>%
  summarise(
    n_amostra_novo = sum(n_amostra),
    TPO_novo = sum(TPO_total),
    TRP_novo = sum(TRP_total),
    TRD_novo = sum(TRD_total),
    .groups = "drop"
  ) %>%
  left_join(
    base_sem_scn %>%
      group_by(Ano) %>%
      summarise(
        n_amostra_original = sum(n_amostra),
        TPO_original = sum(TPO_total),
        TRP_original = sum(TRP_total),
        TRD_original = sum(TRD_total),
        .groups = "drop"
      ),
    by = "Ano"
  ) %>%
  mutate(
    dif_amostra = n_amostra_novo - n_amostra_original,
    dif_TPO = TPO_novo - TPO_original,
    dif_TRP = TRP_novo - TRP_original,
    dif_TRD = TRD_novo - TRD_original
  )

parar_se(any(abs(aud_scn$dif_amostra) > 0), "Amostra mudou apos traducao SCN.")
parar_se(any(abs(aud_scn$dif_TPO) > 1e-5), "TPO mudou apos traducao SCN.")
parar_se(any(abs(aud_scn$dif_TRP) > 1e-3), "TRP mudou apos traducao SCN.")
parar_se(any(abs(aud_scn$dif_TRD) > 1e-3), "TRD mudou apos traducao SCN.")

# Fechamentos internos.
parar_se(
  max(abs(PNAD_FINAL$TPO_total - PNAD_FINAL$TPO_formal - PNAD_FINAL$TPO_informal)) > 1e-6,
  "Falha TPO total = formal + informal."
)
parar_se(
  max(abs(PNAD_FINAL$TRP_total - PNAD_FINAL$TRP_formal - PNAD_FINAL$TRP_informal)) > 1e-4,
  "Falha TRP total = formal + informal."
)
parar_se(
  max(abs(PNAD_FINAL$TRD_total - PNAD_FINAL$TRD_formal - PNAD_FINAL$TRD_informal)) > 1e-4,
  "Falha TRD total = formal + informal."
)

# Teste de regressao dos totais de Santa Catarina.
sc_check <- DADOS_SC %>%
  group_by(Ano) %>%
  summarise(
    TPO_total = sum(TPO_total),
    TPO_informal = sum(TPO_informal),
    TRP_total = sum(TRP_total),
    .groups = "drop"
  ) %>%
  left_join(ref_sc, by = "Ano", suffix = c("_novo", "_ref"))

if (VALIDAR_CONTRA_RESULTADO_FINAL) {
  parar_se(max(abs(sc_check$TPO_total_novo - sc_check$TPO_total_ref)) > 1e-5, "TPO de SC divergiu do arquivo final validado.")
  parar_se(max(abs(sc_check$TPO_informal_novo - sc_check$TPO_informal_ref)) > 1e-5, "TPO informal de SC divergiu do arquivo final validado.")
  parar_se(max(abs(sc_check$TRP_total_novo - sc_check$TRP_total_ref)) > 1e-3, "TRP de SC divergiu do arquivo final validado.")
}

# ==============================================================================
# 9. TABELA SC POR CNAEDOM x ANO - BASE DE Dinamic Table SC E PNAD SC
# ==============================================================================

resumo_sc <- DADOS_SC %>%
  group_by(CNAEDOM, Ano) %>%
  summarise(
    TPO_total = sum(TPO_total),
    TRP_total = sum(TRP_total),
    .groups = "drop"
  )

cnaes_sc <- sort(unique(resumo_sc$CNAEDOM))
parar_se(length(cnaes_sc) != 216, "A tabela PNAD SC deveria conter 216 CNAEDOM.")

# Conferir se, no Tradutor atualizado, so 00000 e 01999 ficam nao identificados em SC.
manual_check <- data.frame(CNAEDOM = cnaes_sc, stringsAsFactors = FALSE) %>%
  left_join(
    tradutor_atualizado %>% select(CNAEDOM, SCN_manual = SCN_atividade),
    by = "CNAEDOM"
  ) %>%
  mutate(
    SCN_manual = ifelse(CNAEDOM %in% c("00000", "01999", "99000"), NA_character_, SCN_manual)
  )

na_manual_sc <- manual_check$CNAEDOM[is.na(manual_check$SCN_manual)]
parar_se(
  !setequal(na_manual_sc, c("00000", "01999")),
  paste("SCN manual nao identificado em codigos inesperados:", paste(na_manual_sc, collapse = ", "))
)

# Matriz com pares TPO/TRP para 2018-2025.
mat_sc <- data.frame(CNAEDOM = cnaes_sc, stringsAsFactors = FALSE)
for (ano in ANOS) {
  tmp <- resumo_sc %>% filter(Ano == ano)
  mat_sc[[paste0("TPO_", ano)]] <- tmp$TPO_total[match(cnaes_sc, tmp$CNAEDOM)]
  mat_sc[[paste0("TRP_", ano)]] <- tmp$TRP_total[match(cnaes_sc, tmp$CNAEDOM)]
}

# ==============================================================================
# 10. DICIONARIO E METODOLOGIA - TEXTO DO ARQUIVO FINAL
# ==============================================================================

dicionario <- data.frame(
  Variavel = c(
    "Ano", "UF_codigo", "UF", "CNAEDOM", "CNAE20", "descricao_CNAE20",
    "SCN_atividade", "descricao_SCN", "OCC", "FAM", "n_amostra",
    "n_amostra_formal", "n_amostra_informal", "TPO_total", "TPO_formal",
    "TPO_informal", "taxa_informalidade", "TRP_total", "TRP_formal",
    "TRP_informal", "TRD_total", "TRD_formal", "TRD_informal",
    "parcela_total", "parcela_formal", "parcela_informal",
    "NOTA 1", "NOTA 2", "NOTA 3", "NOTA 4", "NOTA 5", "NOTA 6", "NOTA 7"
  ),
  Descricao = c(
    "Ano de referência da PNAD Contínua.",
    "Código numérico da Unidade da Federação.",
    "Sigla da Unidade da Federação.",
    "Código CNAE Domiciliar da atividade do trabalho principal, preservado com 5 dígitos.",
    "Código da classe CNAE 2.0 associada ao CNAEDOM, conforme tradutor utilizado no projeto.",
    "Descrição da classe CNAE 2.0 associada ao CNAEDOM.",
    "Código da atividade do Sistema de Contas Nacionais utilizado na compatibilização da MIP-SC.",
    "Descrição da atividade do Sistema de Contas Nacionais.",
    "Classe de rendimento individual do trabalhador, em salários mínimos: OCC1 = até 2 SM; OCC2 = mais de 2 até 3 SM; OCC3 = mais de 3 até 6 SM; OCC4 = mais de 6 até 10 SM; OCC5 = mais de 10 até 15 SM; OCC6 = mais de 15 até 25 SM; OCC7 = mais de 25 SM.",
    "Classe de rendimento domiciliar, em salários mínimos: FAM1 = até 2 SM; FAM2 = mais de 2 até 3 SM; FAM3 = mais de 3 até 6 SM; FAM4 = mais de 6 até 10 SM; FAM5 = mais de 10 até 15 SM; FAM6 = mais de 15 até 25 SM; FAM7 = mais de 25 SM.",
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
    "Participação da classe FAM na renda domiciliar associada a determinado grupo OCC (variável presente nas tabelas derivadas do arquivo anterior; pode ser recalculada a partir dos dados finais).",
    "Participação da FAM na renda domiciliar associada aos trabalhadores formais de determinado OCC (variável presente nas tabelas derivadas do arquivo anterior; pode ser recalculada).",
    "Participação da FAM na renda domiciliar associada aos trabalhadores informais de determinado OCC (variável presente nas tabelas derivadas do arquivo anterior; pode ser recalculada).",
    "O arquivo utiliza a 5ª visita da PNAD Contínua em todos os anos de 2018 a 2025.",
    "Formal e informal seguem a proxy de informalidade utilizada no processamento do projeto, baseada nas variáveis da PNAD Contínua.",
    "OCC é construída com VD4017, rendimento efetivamente recebido no trabalho principal.",
    "FAM utiliza VD5004. TRD deve ser interpretada como renda domiciliar associada aos trabalhadores da célula, e não como uma estimativa sem duplicação da renda total dos domicílios.",
    "As classes OCC e FAM utilizam o salário mínimo correspondente ao período. Nos anos com alteração intra-anual, foi adotada aproximação trimestral: em 2020, R$ 1.043 no 1º trimestre e R$ 1.045 nos demais; em 2023, R$ 1.302 no 1º trimestre, R$ 1.314 no 2º trimestre e R$ 1.320 no 3º e 4º trimestres.",
    "Educação e saúde são desdobradas entre atividades públicas e privadas a partir de VD4009 antes da agregação final; por isso o arquivo com SCN possui mais linhas que o arquivo sem SCN, sem alteração dos totais.",
    "O CNAEDOM 00000 é utilizado pelo IBGE para atividades mal definidas, não declaradas ou não identificadas. Por não permitir identificar uma atividade econômica específica, ele não pode ser compatibilizado de forma não arbitrária com a CNAE 2.0 nem com uma atividade do SCN. Por isso, no arquivo final, os campos CNAE20, descricao_CNAE20, SCN_atividade e descricao_SCN dessas linhas são preenchidos com “Não identificado” e destacados em amarelo."
  ),
  stringsAsFactors = FALSE
)

metodologia <- data.frame(
  Item = c(
    "Fonte dos dados", "Período", "Visita utilizada", "Cobertura territorial",
    "Unidade de agregação", "Atividade econômica", "Peso amostral",
    "Pessoal ocupado (TPO)", "Rendimento do trabalho principal (TRP)",
    "Rendimento domiciliar (TRD)", "Formalidade", "Classes OCC", "Classes FAM",
    "Salário mínimo", "Compatibilização CNAEDOM-CNAE 2.0",
    "Compatibilização CNAE-SCN", "Educação pública e privada",
    "Saúde pública e privada", "Categorias residuais", "Código SCN 4580",
    "Preservação dos totais", "Estrutura do arquivo final", "Tratamento do CNAEDOM 00000"
  ),
  Descricao = c(
    "Pesquisa Nacional por Amostra de Domicílios Contínua (PNAD Contínua), IBGE.",
    "2018 a 2025.",
    "5ª visita da PNAD Contínua em todos os anos.",
    "Todas as Unidades da Federação do Brasil; a variável UF identifica o recorte territorial.",
    "Ano × UF × CNAEDOM × CNAE 2.0 × atividade SCN × OCC × FAM. Para educação e saúde, uma mesma combinação original pode ser desdobrada entre atividade pública e privada.",
    "A atividade do trabalho principal é identificada por V4013 (CNAE Domiciliar 2.0), mantida na base como CNAEDOM com 5 dígitos.",
    "V1032 é utilizado como peso de expansão nas estimativas ponderadas.",
    "TPO_total é calculado pela soma de ocupado × V1032. TPO_formal e TPO_informal aplicam o mesmo peso aos respectivos grupos.",
    "TRP utiliza o rendimento efetivo do trabalho principal (VD4017 / renda_ind_princ_efetiva) ponderado por V1032.",
    "TRD utiliza VD5004 / renda_dom_efetiva ponderada por V1032. Deve ser interpretada como renda domiciliar associada aos trabalhadores da célula, podendo haver repetição do rendimento do mesmo domicílio entre trabalhadores.",
    "As observações são separadas entre formal e informal segundo a proxy utilizada no processamento da PNAD do projeto; n_amostra, TPO, TRP e TRD são apresentados para total, formal e informal.",
    "OCC1: até 2 SM; OCC2: >2 a 3 SM; OCC3: >3 a 6 SM; OCC4: >6 a 10 SM; OCC5: >10 a 15 SM; OCC6: >15 a 25 SM; OCC7: >25 SM. A classificação usa o rendimento efetivo do trabalho principal.",
    "FAM1: até 2 SM; FAM2: >2 a 3 SM; FAM3: >3 a 6 SM; FAM4: >6 a 10 SM; FAM5: >10 a 15 SM; FAM6: >15 a 25 SM; FAM7: >25 SM. A classificação usa o rendimento domiciliar efetivo.",
    "As classes OCC e FAM usam o salário mínimo do período. Para anos com mudança intra-anual, a correção é trimestral: 2020 = R$ 1.043 no 1º trimestre e R$ 1.045 nos demais; 2023 = R$ 1.302 no 1º trimestre, R$ 1.314 no 2º e R$ 1.320 no 3º e 4º trimestres.",
    "A CNAE Domiciliar 2.0 foi vinculada à classe CNAE 2.0 por meio do tradutor utilizado no projeto, preservando os 222 CNAEDOM observados na base.",
    "A classe CNAE 2.0 foi compatibilizada com as atividades do SCN pelo tradutor fornecido para o projeto e auditado durante o processamento.",
    "Para CNAEDOM do grupo 85, VD4009 = 05, 06 ou 07 é classificado em Educação Pública (8591); os demais casos são classificados em Educação Privada (8592).",
    "Para CNAEDOM dos grupos 86 e 87, VD4009 = 05, 06 ou 07 é classificado em Saúde Pública (8691); os demais casos são classificados em Saúde Privada (8692). A assistência social CNAEDOM 88000 é direcionada para 8692.",
    "CNAEDOM 00000 (atividades mal definidas), 01999 (agropecuária não especificada, ambígua entre agricultura e pecuária) e 99000 (organismos internacionais/extraterritoriais) são mantidos na base sem código SCN, evitando imputação arbitrária.",
    "O código 4580 é mantido conforme o tradutor e a convenção metodológica adotados no projeto, em alinhamento com o tratamento utilizado na base RAIS e com orientação do professor.",
    "A tradução e o desdobramento público/privado preservam os totais anuais de n_amostra, TPO, TRP e TRD. Diferenças residuais eventualmente observadas são apenas de precisão numérica de ponto flutuante.",
    "A aba Dados_PNAD concentra os dados necessários ao uso posterior. As abas derivadas do arquivo sem SCN podem ser reconstruídas a partir dela. Dicionario documenta as variáveis e Metodologia registra as decisões de processamento e compatibilização.",
    "O código CNAEDOM 00000 representa atividades mal definidas, não declaradas ou não identificadas na PNAD Contínua. Como não há informação suficiente para associá-lo com segurança a uma classe CNAE 2.0 ou a uma atividade SCN, não é realizada imputação setorial. Os respectivos campos de CNAE 2.0 e SCN são registrados como “Não identificado” e destacados em amarelo, preservando a informação original e tornando explícita a ausência de compatibilização."
  ),
  stringsAsFactors = FALSE
)

# ==============================================================================
# 11. GERAR EXCEL FINAL
# ==============================================================================

arquivo_final <- file.path(
  DIR_RESULTADOS_EXCEL,
  "PNADC_2018_2025_FINAL_CNAE_SCN.xlsx"
)

wb <- openxlsx::createWorkbook()

# Estilos simples e estaveis.
st_titulo <- openxlsx::createStyle(fontSize = 14, textDecoration = "bold")
st_header <- openxlsx::createStyle(
  textDecoration = "bold",
  fgFill = "D9EAF7",
  halign = "center",
  valign = "center",
  border = "Bottom"
)
st_amarelo <- openxlsx::createStyle(fgFill = "FFF200")
st_codigo <- openxlsx::createStyle(numFmt = "@")
st_num <- openxlsx::createStyle(numFmt = "#,##0.0000")
st_moeda <- openxlsx::createStyle(numFmt = "#,##0.00")

# -----------------------------------------------------------------------------
# 11.1 Metodologia
# -----------------------------------------------------------------------------
openxlsx::addWorksheet(wb, "Metodologia")
openxlsx::writeData(wb, "Metodologia", "Notas metodológicas", startRow = 1, startCol = 1)
openxlsx::addStyle(wb, "Metodologia", st_titulo, rows = 1, cols = 1)
openxlsx::writeData(wb, "Metodologia", metodologia, startRow = 3, startCol = 1)
openxlsx::addStyle(wb, "Metodologia", st_header, rows = 3, cols = 1:2, gridExpand = TRUE)
openxlsx::setColWidths(wb, "Metodologia", cols = 1, widths = 32)
openxlsx::setColWidths(wb, "Metodologia", cols = 2, widths = 120)
openxlsx::setRowHeights(wb, "Metodologia", rows = 4:(3 + nrow(metodologia)), heights = 35)
openxlsx::addStyle(
  wb, "Metodologia",
  openxlsx::createStyle(wrapText = TRUE, valign = "top"),
  rows = 3:(3 + nrow(metodologia)), cols = 1:2, gridExpand = TRUE, stack = TRUE
)

# -----------------------------------------------------------------------------
# 11.2 Dados_PNAD e Dados_SC
# -----------------------------------------------------------------------------
for (nome in c("Dados_PNAD", "Dados_SC")) {
  openxlsx::addWorksheet(wb, nome)
  obj <- if (nome == "Dados_PNAD") PNAD_FINAL else DADOS_SC
  openxlsx::writeData(wb, nome, obj, startRow = 1, startCol = 1, withFilter = FALSE)
  openxlsx::addStyle(wb, nome, st_header, rows = 1, cols = 1:ncol(obj), gridExpand = TRUE)
  openxlsx::freezePane(wb, nome, firstRow = TRUE)
  openxlsx::setColWidths(wb, nome, cols = c(1:5, 7, 9:23), widths = "auto")
  openxlsx::setColWidths(wb, nome, cols = c(6, 8), widths = 45)
  openxlsx::addStyle(wb, nome, st_codigo, rows = 2:(nrow(obj) + 1), cols = c(4, 5, 7), gridExpand = TRUE, stack = TRUE)

  linhas_00000 <- which(obj$CNAEDOM == "00000") + 1
  if (length(linhas_00000) > 0) {
    openxlsx::addStyle(
      wb, nome, st_amarelo,
      rows = linhas_00000, cols = 5:8,
      gridExpand = TRUE, stack = TRUE
    )
  }
}

# -----------------------------------------------------------------------------
# 11.3 Dicionario
# -----------------------------------------------------------------------------
openxlsx::addWorksheet(wb, "Dicionario")
openxlsx::writeData(wb, "Dicionario", dicionario)
openxlsx::addStyle(wb, "Dicionario", st_header, rows = 1, cols = 1:2, gridExpand = TRUE)
openxlsx::setColWidths(wb, "Dicionario", cols = 1, widths = 24)
openxlsx::setColWidths(wb, "Dicionario", cols = 2, widths = 120)
openxlsx::addStyle(
  wb, "Dicionario",
  openxlsx::createStyle(wrapText = TRUE, valign = "top"),
  rows = 1:(nrow(dicionario) + 1), cols = 1:2, gridExpand = TRUE, stack = TRUE
)

# -----------------------------------------------------------------------------
# 11.4 Dinamic Table SC - reproducao estatica da tabela dinamica usada no arquivo
# -----------------------------------------------------------------------------
openxlsx::addWorksheet(wb, "Dinamic Table SC")
openxlsx::writeData(wb, "Dinamic Table SC", "Rótulos de Coluna", startRow = 3, startCol = 2)

for (i in seq_along(ANOS)) {
  col_tpo <- 2 + (i - 1) * 2
  openxlsx::writeData(wb, "Dinamic Table SC", ANOS[i], startRow = 4, startCol = col_tpo, colNames = FALSE)
  openxlsx::writeData(wb, "Dinamic Table SC", "Soma de TPO_total", startRow = 5, startCol = col_tpo, colNames = FALSE)
  openxlsx::writeData(wb, "Dinamic Table SC", "Soma de TRP_total", startRow = 5, startCol = col_tpo + 1, colNames = FALSE)
}
openxlsx::writeData(wb, "Dinamic Table SC", "Total Soma de TPO_total", startRow = 4, startCol = 18, colNames = FALSE)
openxlsx::writeData(wb, "Dinamic Table SC", "Total Soma de TRP_total", startRow = 4, startCol = 19, colNames = FALSE)
openxlsx::writeData(wb, "Dinamic Table SC", "Rótulos de Linha", startRow = 5, startCol = 1, colNames = FALSE)

# Dados das 216 CNAEDOM.
for (i in seq_along(cnaes_sc)) {
  linha <- 5 + i
  cod <- cnaes_sc[i]
  openxlsx::writeData(wb, "Dinamic Table SC", cod, startRow = linha, startCol = 1, colNames = FALSE)
  total_tpo <- 0
  total_trp <- 0
  for (j in seq_along(ANOS)) {
    ano <- ANOS[j]
    col_tpo <- 2 + (j - 1) * 2
    v_tpo <- mat_sc[[paste0("TPO_", ano)]][i]
    v_trp <- mat_sc[[paste0("TRP_", ano)]][i]
    if (!is.na(v_tpo)) {
      openxlsx::writeData(wb, "Dinamic Table SC", v_tpo, startRow = linha, startCol = col_tpo, colNames = FALSE)
      total_tpo <- total_tpo + v_tpo
    }
    if (!is.na(v_trp)) {
      openxlsx::writeData(wb, "Dinamic Table SC", v_trp, startRow = linha, startCol = col_tpo + 1, colNames = FALSE)
      total_trp <- total_trp + v_trp
    }
  }
  openxlsx::writeData(wb, "Dinamic Table SC", total_tpo, startRow = linha, startCol = 18, colNames = FALSE)
  openxlsx::writeData(wb, "Dinamic Table SC", total_trp, startRow = linha, startCol = 19, colNames = FALSE)
}

linha_total <- 6 + length(cnaes_sc)
openxlsx::writeData(wb, "Dinamic Table SC", "Total Geral", startRow = linha_total, startCol = 1, colNames = FALSE)
for (j in seq_along(ANOS)) {
  ano <- ANOS[j]
  col_tpo <- 2 + (j - 1) * 2
  openxlsx::writeData(wb, "Dinamic Table SC", sum(mat_sc[[paste0("TPO_", ano)]], na.rm = TRUE), startRow = linha_total, startCol = col_tpo, colNames = FALSE)
  openxlsx::writeData(wb, "Dinamic Table SC", sum(mat_sc[[paste0("TRP_", ano)]], na.rm = TRUE), startRow = linha_total, startCol = col_tpo + 1, colNames = FALSE)
}
openxlsx::writeData(wb, "Dinamic Table SC", sum(resumo_sc$TPO_total), startRow = linha_total, startCol = 18, colNames = FALSE)
openxlsx::writeData(wb, "Dinamic Table SC", sum(resumo_sc$TRP_total), startRow = linha_total, startCol = 19, colNames = FALSE)
openxlsx::addStyle(wb, "Dinamic Table SC", st_header, rows = 4:5, cols = 1:19, gridExpand = TRUE)
openxlsx::addStyle(wb, "Dinamic Table SC", st_header, rows = linha_total, cols = 1:19, gridExpand = TRUE)
openxlsx::setColWidths(wb, "Dinamic Table SC", cols = 1:19, widths = "auto")

# -----------------------------------------------------------------------------
# 11.5 PNAD SC - tabela usada manualmente no projeto
# -----------------------------------------------------------------------------
openxlsx::addWorksheet(wb, "PNAD SC")

# Anos na linha 4.
for (i in seq_along(ANOS)) {
  openxlsx::writeData(
    wb, "PNAD SC", ANOS[i],
    startRow = 4, startCol = 3 + (i - 1) * 2,
    colNames = FALSE
  )
}

cab <- c("SCN", "CNAE DOM")
for (ano in ANOS) cab <- c(cab, "Pessoal ocupadas", "Massa salarial")
openxlsx::writeData(wb, "PNAD SC", t(cab), startRow = 5, startCol = 1, colNames = FALSE)

# CNAEDOM + valores anuais.
openxlsx::writeData(wb, "PNAD SC", cnaes_sc, startRow = 6, startCol = 2, colNames = FALSE)
for (i in seq_along(ANOS)) {
  ano <- ANOS[i]
  openxlsx::writeData(wb, "PNAD SC", mat_sc[[paste0("TPO_", ano)]], startRow = 6, startCol = 3 + (i - 1) * 2, colNames = FALSE, keepNA = FALSE)
  openxlsx::writeData(wb, "PNAD SC", mat_sc[[paste0("TRP_", ano)]], startRow = 6, startCol = 4 + (i - 1) * 2, colNames = FALSE, keepNA = FALSE)
}

# Formula SCN - mesma logica do arquivo final.
formulas_scn <- vapply(
  6:(5 + length(cnaes_sc)),
  function(r) {
    paste0(
      '=IF(B', r, '="","",IF(OR(B', r, '="00000",B', r,
      '="01999",B', r, '="99000"),"Não identificado",',
      'IFERROR(VLOOKUP(B', r, ",'Tradutor atualizado'!$A:$F,5,FALSE),",
      '"Não identificado")))'
    )
  },
  character(1)
)
openxlsx::writeFormula(wb, "PNAD SC", formulas_scn, startRow = 6, startCol = 1)

# Total Geral.
openxlsx::writeData(wb, "PNAD SC", "Total Geral", startRow = linha_total, startCol = 2, colNames = FALSE)
for (i in seq_along(ANOS)) {
  ano <- ANOS[i]
  openxlsx::writeData(wb, "PNAD SC", sum(mat_sc[[paste0("TPO_", ano)]], na.rm = TRUE), startRow = linha_total, startCol = 3 + (i - 1) * 2, colNames = FALSE)
  openxlsx::writeData(wb, "PNAD SC", sum(mat_sc[[paste0("TRP_", ano)]], na.rm = TRUE), startRow = linha_total, startCol = 4 + (i - 1) * 2, colNames = FALSE)
}

# Testes abaixo da tabela, como no arquivo final.
linha_teste_label <- linha_total + 3
linha_teste <- linha_total + 4
for (i in seq_along(ANOS)) {
  ano <- ANOS[i]
  c_tpo <- 3 + (i - 1) * 2
  c_trp <- 4 + (i - 1) * 2
  openxlsx::writeData(wb, "PNAD SC", paste("Teste ocupados", ano), startRow = linha_teste_label, startCol = c_tpo, colNames = FALSE)
  openxlsx::writeData(wb, "PNAD SC", paste("Teste massa", ano), startRow = linha_teste_label, startCol = c_trp, colNames = FALSE)

  letra_tpo <- col_letra(c_tpo)
  letra_trp <- col_letra(c_trp)
  formula_tpo <- paste0("=", letra_tpo, linha_total, "-SUMIFS(Dados_SC!$N:$N,Dados_SC!$A:$A,", letra_tpo, "$4)")
  formula_trp <- paste0("=", letra_trp, linha_total, "-SUMIFS(Dados_SC!$R:$R,Dados_SC!$A:$A,", letra_tpo, "$4)")
  openxlsx::writeFormula(wb, "PNAD SC", formula_tpo, startRow = linha_teste, startCol = c_tpo)
  openxlsx::writeFormula(wb, "PNAD SC", formula_trp, startRow = linha_teste, startCol = c_trp)
}

openxlsx::addStyle(wb, "PNAD SC", st_header, rows = 4:5, cols = 1:18, gridExpand = TRUE)
openxlsx::addStyle(wb, "PNAD SC", st_header, rows = linha_total, cols = 2:18, gridExpand = TRUE)
openxlsx::addStyle(wb, "PNAD SC", st_header, rows = linha_teste_label, cols = 3:18, gridExpand = TRUE)
openxlsx::addStyle(wb, "PNAD SC", st_codigo, rows = 6:(5 + length(cnaes_sc)), cols = 1:2, gridExpand = TRUE, stack = TRUE)
openxlsx::setColWidths(wb, "PNAD SC", cols = 1:18, widths = "auto")

# -----------------------------------------------------------------------------
# 11.6 Tradutor - snapshot da fonte do professor
# -----------------------------------------------------------------------------
openxlsx::addWorksheet(wb, "Tradutor")
openxlsx::writeData(wb, "Tradutor", tradutor_snapshot, colNames = FALSE, keepNA = TRUE, na.string = "#N/A")
openxlsx::setColWidths(wb, "Tradutor", cols = 1:min(10, ncol(tradutor_snapshot)), widths = "auto")

# -----------------------------------------------------------------------------
# 11.7 Tradutor atualizado - mesma logica usada na tabela PNAD SC
# -----------------------------------------------------------------------------
openxlsx::addWorksheet(wb, "Tradutor atualizado")
openxlsx::writeData(wb, "Tradutor atualizado", "CNAE domiciliar...2.0 5D", startRow = 1, startCol = 1, colNames = FALSE)
openxlsx::writeData(wb, "Tradutor atualizado", "Soma de 2015", startRow = 2, startCol = 1, colNames = FALSE)
openxlsx::writeData(
  wb, "Tradutor atualizado",
  t(c("Rótulos de Linha", "CNAE domiciliar...2.0", "CNAE 2.0", "Atividades CNAE 2.0", "COD ATIV SCN", "ATIVIDADE SCN", "COD PROD SCN", "PRODUTOS SCN")),
  startRow = 3, startCol = 1, colNames = FALSE
)
openxlsx::writeData(
  wb, "Tradutor atualizado",
  tradutor_atualizado,
  startRow = 4, startCol = 1,
  colNames = FALSE,
  keepNA = TRUE,
  na.string = "#N/A"
)

# Linhas finais da tabela/pivot de origem, preservadas por compatibilidade visual.
openxlsx::writeData(wb, "Tradutor atualizado", "(vazio)", startRow = 4 + nrow(tradutor_atualizado), startCol = 1, colNames = FALSE)
openxlsx::writeData(wb, "Tradutor atualizado", "Total Geral", startRow = 5 + nrow(tradutor_atualizado), startCol = 1, colNames = FALSE)

openxlsx::addStyle(wb, "Tradutor atualizado", st_header, rows = 3, cols = 1:8, gridExpand = TRUE)
openxlsx::addStyle(wb, "Tradutor atualizado", st_codigo, rows = 4:(3 + nrow(tradutor_atualizado)), cols = c(1, 3, 5, 7), gridExpand = TRUE, stack = TRUE)
openxlsx::setColWidths(wb, "Tradutor atualizado", cols = c(1, 3, 5, 7), widths = 14)
openxlsx::setColWidths(wb, "Tradutor atualizado", cols = c(2, 4, 6, 8), widths = 45)

# Amarelo: 01119 apenas E:F; novas linhas 19010/29003/32002 A:H.
linha_01119 <- which(tradutor_atualizado$CNAEDOM == "01119") + 3
linhas_novas <- which(tradutor_atualizado$CNAEDOM %in% c("19010", "29003", "32002")) + 3
openxlsx::addStyle(wb, "Tradutor atualizado", st_amarelo, rows = linha_01119, cols = 5:6, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, "Tradutor atualizado", st_amarelo, rows = linhas_novas, cols = 1:8, gridExpand = TRUE, stack = TRUE)

# ==============================================================================
# 12. SALVAR E REGISTRAR AUDITORIAS
# ==============================================================================

openxlsx::saveWorkbook(wb, arquivo_final, overwrite = TRUE)

saveRDS(
  PNAD_FINAL,
  file.path(DIR_RESULTADOS_EXCEL, "PNADC_2018_2025_FINAL_CNAE_SCN.rds"),
  compress = TRUE
)

data.table::fwrite(data.table::as.data.table(aud_scn), file.path(DIR_LOGS, "AUDITORIA_pipeline_02_preservacao_totais.csv"))
data.table::fwrite(data.table::as.data.table(sc_check), file.path(DIR_LOGS, "AUDITORIA_pipeline_02_SC_referencia.csv"))
data.table::fwrite(data.table::as.data.table(comparacao_manual), file.path(DIR_LOGS, "AUDITORIA_pipeline_02_mudancas_tradutor_manual.csv"))

capture.output(
  sessionInfo(),
  file = file.path(DIR_LOGS, "sessionInfo_pipeline_02_PNADC.txt")
)

cat("\n============================================================\n")
cat("PIPELINE 02 CONCLUIDO\n")
cat("============================================================\n")
cat("Dados_PNAD:", nrow(PNAD_FINAL), "linhas\n")
cat("Dados_SC:", nrow(DADOS_SC), "linhas\n")
cat("CNAEDOM nacionais:", n_distinct(PNAD_FINAL$CNAEDOM), "\n")
cat("CNAEDOM de SC:", n_distinct(DADOS_SC$CNAEDOM), "\n")
cat("Nao identificados na tabela PNAD SC:", paste(na_manual_sc, collapse = ", "), "\n")
cat("Excel final:\n", arquivo_final, "\n", sep = "")
cat("\nTodas as auditorias obrigatorias passaram.\n")
