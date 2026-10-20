#============================================================#
# COMPATIBILIZAÇÃO CNAE 2.0 -> SCN
# RAIS SC 2018-2025 | Projeto de atualização da MIP-SC
#
# Entrada:
#   1) RAIS_SC_2018_2025_FINAL_TODAS_CNAES_CORRIGIDAS.xlsx
#   2) Tradutor CNAE 20 para SCN.xlsx
#
# Saída:
#   RAIS_SC_2018_2025_FINAL_CNAE_SCN.xlsx
#
# O script:
#   - NÃO altera massa salarial, vínculos ou dados municipais;
#   - mantém a classe CNAE original de 5 dígitos;
#   - cria a classe CNAE de 4 dígitos para o casamento com o tradutor;
#   - adiciona SCN padrão IBGE e SCN desagregado da MIP-SC;
#   - trata separadamente Educação Pública e Saúde Pública;
#   - audita correspondências e interrompe se houver CNAE sem tradução.
#============================================================#


rm(list = ls())

library(dplyr)
library(openxlsx)


#============================================================#
# 1. CONFIGURAÇÕES
#============================================================#

pasta_projeto <-
  "C:/Users/10766048977/Desktop/Projeto atualização da MIP/R/RAIS"

anos <- 2018:2025


#------------------------------------------------------------#
# Arquivo RAIS final já corrigido
#------------------------------------------------------------#

arquivo_rais <- file.path(
  pasta_projeto,
  "Resultados RAIS",
  "Excel Final",
  "RAIS_SC_2018_2025_FINAL_TODAS_CNAES_CORRIGIDAS.xlsx"
)


#------------------------------------------------------------#
# Tradutor enviado pelo professor
#
# Neste código estou assumindo que o arquivo foi colocado
# diretamente na pasta principal do projeto RAIS.
#------------------------------------------------------------#

arquivo_tradutor <-
  "C:/Users/10766048977/Desktop/Projeto atualização da MIP/Tradutores/Traduores Arlei/Tradutor CNAE 20 para SCN.xlsx"


#------------------------------------------------------------#
# Arquivo final
#------------------------------------------------------------#

pasta_saida <- file.path(
  pasta_projeto,
  "Resultados RAIS",
  "Excel Final"
)

dir.create(
  pasta_saida,
  recursive = TRUE,
  showWarnings = FALSE
)


arquivo_saida <- file.path(
  pasta_saida,
  "RAIS_SC_2018_2025_FINAL_CNAE_SCN.xlsx"
)


#------------------------------------------------------------#
# Pasta de auditoria
#------------------------------------------------------------#

pasta_auditoria <- file.path(
  pasta_projeto,
  "Resultados RAIS",
  "Auditoria"
)

dir.create(
  pasta_auditoria,
  recursive = TRUE,
  showWarnings = FALSE
)


#------------------------------------------------------------#
# Confere arquivos
#------------------------------------------------------------#

if (!file.exists(arquivo_rais)) {
  stop(
    paste(
      "Arquivo RAIS final não encontrado:",
      arquivo_rais
    )
  )
}


if (!file.exists(arquivo_tradutor)) {
  stop(
    paste(
      "Tradutor CNAE-SCN não encontrado:",
      arquivo_tradutor
    )
  )
}


#============================================================#
# 2. FUNÇÕES AUXILIARES
#============================================================#

#------------------------------------------------------------#
# Normaliza CNAE de 4 dígitos
#
# Ex.:
#   111  -> 0111
#   4120 -> 4120
#------------------------------------------------------------#

normalizar_cnae4 <- function(x) {

  x <- trimws(
    as.character(x)
  )

  x <- gsub(
    "[^0-9]",
    "",
    x
  )

  x[
    x == ""
  ] <- NA_character_


  numero <- suppressWarnings(
    as.integer(x)
  )


  resultado <- ifelse(
    !is.na(numero),
    sprintf(
      "%04d",
      numero
    ),
    NA_character_
  )


  resultado
}


#------------------------------------------------------------#
# Normaliza código SCN padrão IBGE
#
# O SCN padrão utilizado no tradutor possui 4 dígitos.
#
# Ex.:
#   191  -> 0191
#   4180 -> 4180
#------------------------------------------------------------#

normalizar_scn_ibge <- function(x) {

  x <- trimws(
    as.character(x)
  )

  x <- gsub(
    "[^0-9]",
    "",
    x
  )

  x[
    x == ""
  ] <- NA_character_


  numero <- suppressWarnings(
    as.integer(x)
  )


  resultado <- ifelse(
    !is.na(numero),
    sprintf(
      "%04d",
      numero
    ),
    NA_character_
  )


  resultado
}


#------------------------------------------------------------#
# Normaliza os códigos SCN desagregados do tradutor.
#
# Aqui o código pode ter 4, 5 ou 6 dígitos.
# Para códigos menores que 1000, recupera o zero inicial.
#------------------------------------------------------------#

normalizar_scn_desag <- function(x) {

  x <- trimws(
    as.character(x)
  )

  x <- gsub(
    "[^0-9]",
    "",
    x
  )

  x[
    x == ""
  ] <- NA_character_


  numero <- suppressWarnings(
    as.integer(x)
  )


  resultado <- rep(
    NA_character_,
    length(numero)
  )


  indice <- !is.na(numero)


  resultado[
    indice & numero < 1000
  ] <- sprintf(
    "%04d",
    numero[
      indice & numero < 1000
    ]
  )


  resultado[
    indice & numero >= 1000
  ] <- as.character(
    numero[
      indice & numero >= 1000
    ]
  )


  resultado
}


#============================================================#
# 3. LÊ O TRADUTOR CNAE -> SCN
#============================================================#

# A aba do professor possui cabeçalho na linha 4.
# Os dados começam na linha 5.
#
# Colunas usadas:
# C = Classe CNAE 2.0 (4 dígitos)
# D = Descrição da classe CNAE
# E = Atividade SCN 2010 - IBGE
# F = Descrição SCN 2010 - IBGE
# G = Atividade SCN desagregada
# H = Descrição SCN desagregada

tradutor_raw <- read.xlsx(
  arquivo_tradutor,
  sheet = "TRAD SCN_CNAE20",
  startRow = 5,
  cols = 3:8,
  colNames = FALSE,
  skipEmptyRows = FALSE
)


names(
  tradutor_raw
) <- c(
  "cnae_classe_4dig",
  "descricao_cnae_classe",
  "scn_ibge",
  "descricao_scn_ibge",
  "scn_mip_sc_tradutor",
  "descricao_scn_mip_sc_tradutor"
)


tradutor <- tradutor_raw |>
  mutate(

    cnae_classe_4dig =
      normalizar_cnae4(
        cnae_classe_4dig
      ),

    descricao_cnae_classe =
      trimws(
        as.character(
          descricao_cnae_classe
        )
      ),

    scn_ibge =
      normalizar_scn_ibge(
        scn_ibge
      ),

    descricao_scn_ibge =
      trimws(
        as.character(
          descricao_scn_ibge
        )
      ),

    scn_mip_sc_tradutor =
      normalizar_scn_desag(
        scn_mip_sc_tradutor
      ),

    descricao_scn_mip_sc_tradutor =
      trimws(
        as.character(
          descricao_scn_mip_sc_tradutor
        )
      )

  ) |>

  filter(
    !is.na(
      cnae_classe_4dig
    )
  )


#============================================================#
# 4. AUDITA DUPLICAÇÕES DO TRADUTOR
#============================================================#

# O arquivo do professor possui algumas CNAEs repetidas.
# Isso não é problema se a correspondência for idêntica.
# O código para caso haja uma mesma CNAE associada a SCNs
# diferentes.

conflitos_tradutor <- tradutor |>
  mutate(
    chave_traducao =
      paste(
        scn_ibge,
        scn_mip_sc_tradutor,
        sep = "|"
      )
  ) |>
  group_by(
    cnae_classe_4dig
  ) |>
  summarise(
    traducoes_distintas =
      n_distinct(
        chave_traducao
      ),
    .groups = "drop"
  ) |>
  filter(
    traducoes_distintas > 1
  )


if (nrow(conflitos_tradutor) > 0) {

  write.csv(
    conflitos_tradutor,
    file.path(
      pasta_auditoria,
      "CONFLITOS_TRADUTOR_CNAE_SCN.csv"
    ),
    row.names = FALSE,
    fileEncoding = "UTF-8"
  )


  stop(
    paste(
      "Existem CNAEs com mais de uma correspondência SCN no tradutor.",
      "Confira CONFLITOS_TRADUTOR_CNAE_SCN.csv."
    )
  )
}


# Remove repetições idênticas
tradutor <- tradutor |>
  distinct(
    cnae_classe_4dig,
    .keep_all = TRUE
  )


cat(
  "\nClasses CNAE distintas no tradutor:",
  nrow(tradutor),
  "\n"
)


#============================================================#
# 5. FUNÇÃO DE COMPATIBILIZAÇÃO
#============================================================#

compatibilizar_cnae_scn <- function(dado) {


  #----------------------------------------------------------#
  # 5.1 Preserva CNAE como texto
  #----------------------------------------------------------#

  dado <- dado |>
    mutate(

      cnae_classe =
        trimws(
          as.character(
            cnae_classe
          )
        )

    )


  #----------------------------------------------------------#
  # 5.2 Cria CNAE classe de 4 dígitos
  #
  # A base RAIS possui a classe completa com o dígito
  # verificador:
  #
  #   41204 -> 4120
  #   01113 -> 0111
  #
  # Os códigos auxiliares criados para separar setor público
  # são tratados depois:
  #
  #   851911 = Educação Pública
  #   861911 = Saúde Pública
  #----------------------------------------------------------#

  dado <- dado |>
    mutate(

      cnae_classe_4dig =
        case_when(

          cnae_classe %in%
            c(
              "851911",
              "861911"
            ) ~ NA_character_,

          grepl(
            "^[0-9]{5}$",
            cnae_classe
          ) ~ substr(
            cnae_classe,
            1,
            4
          ),

          TRUE ~ NA_character_

        )

    )


  #----------------------------------------------------------#
  # 5.3 Junta o tradutor
  #----------------------------------------------------------#

  dado <- dado |>
    left_join(
      tradutor,
      by = "cnae_classe_4dig"
    )


  #----------------------------------------------------------#
  # 5.4 Tratamento dos códigos auxiliares de setor público
  #
  # IMPORTANTE:
  #
  # A classificação original do SCN agrupa educação pública
  # e privada em 8591 e saúde pública e privada em 8691.
  #
  # No projeto da MIP-SC, o arquivo do professor apresenta
  # atividades separadas:
  #
  #   8591 = Educação pública
  #   8592 = Educação privada
  #   8691 = Saúde pública
  #   8692 = Saúde privada
  #
  # Como a RAIS já foi tratada anteriormente para separar
  # os vínculos públicos nos códigos auxiliares 851911 e
  # 861911, os registros restantes das atividades de
  # educação/saúde são direcionados aos setores privados.
  #----------------------------------------------------------#

  dado <- dado |>
    mutate(

      #--------------------------------------#
      # SCN padrão IBGE
      #--------------------------------------#

      scn_ibge =
        case_when(

          cnae_classe == "851911" ~ "8591",

          cnae_classe == "861911" ~ "8691",

          TRUE ~ scn_ibge

        ),


      descricao_scn_ibge =
        case_when(

          cnae_classe == "851911" ~
            "Educação pública e privada",

          cnae_classe == "861911" ~
            "Saúde pública e privada",

          cnae_classe == "99008" ~
            "Não se aplica - organismo internacional ou instituição extraterritorial",

          TRUE ~ descricao_scn_ibge

        ),


      #--------------------------------------#
      # CNAE desagregada utilizada na MIP-SC
      #--------------------------------------#

      scn_mip_sc =
        case_when(

          # Educação pública
          cnae_classe == "851911" ~
            "8591",

          # Saúde pública
          cnae_classe == "861911" ~
            "8691",

          # Demais atividades de educação
          # que no SCN padrão estavam em 8591
          scn_ibge == "8591" ~
            "8592",

          # Demais atividades de saúde
          # que no SCN padrão estavam em 8691
          scn_ibge == "8691" ~
            "8692",

          TRUE ~
            scn_mip_sc_tradutor

        ),


      descricao_scn_mip_sc =
        case_when(

          cnae_classe == "851911" ~
            "Educação pública",

          cnae_classe == "861911" ~
            "Saúde pública",

          scn_ibge == "8591" ~
            "Educação privada",

          scn_ibge == "8691" ~
            "Saúde privada",

          cnae_classe == "99008" ~
            "Não se aplica - organismo internacional ou instituição extraterritorial",

          TRUE ~
            descricao_scn_mip_sc_tradutor

        ),


      #--------------------------------------#
      # Descrição dos códigos auxiliares
      #--------------------------------------#

      descricao_cnae_classe =
        case_when(

          cnae_classe == "851911" ~
            "Educação pública - código auxiliar",

          cnae_classe == "861911" ~
            "Saúde pública - código auxiliar",

          cnae_classe == "99008" ~
            "Organismos internacionais e outras instituições extraterritoriais",

          TRUE ~
            descricao_cnae_classe

        ),


      #--------------------------------------#
      # Status
      #--------------------------------------#

      status_compatibilizacao =
        case_when(

          cnae_classe %in%
            c(
              "851911",
              "861911"
            ) ~ "REGRA_ESPECIAL_PUBLICO",

          cnae_classe == "99008" ~
            "EXTRATERRITORIAL_SEM_SCN",

          !is.na(
            scn_mip_sc
          ) ~ "OK",

          TRUE ~
            "SEM_CORRESPONDENCIA"

        )

    )


  #----------------------------------------------------------#
  # 5.5 Remove colunas auxiliares internas do tradutor
  #----------------------------------------------------------#

  dado <- dado |>
    select(
      -scn_mip_sc_tradutor,
      -descricao_scn_mip_sc_tradutor
    )


  dado
}


#============================================================#
# 6. PROCESSA 2018-2025
#============================================================#

bases_compatibilizadas <- list()

auditoria <- list()

cnaes_sem_correspondencia <- list()


for (ano_atual in anos) {


  cat(
    "\n========================================\n"
  )

  cat(
    "COMPATIBILIZANDO:",
    ano_atual,
    "\n"
  )

  cat(
    "========================================\n"
  )


  #----------------------------------------------------------#
  # Abre a aba anual
  #----------------------------------------------------------#

  dado_original <- read.xlsx(
    arquivo_rais,
    sheet = as.character(
      ano_atual
    ),
    check.names = FALSE
  )


  #----------------------------------------------------------#
  # Garante tipos das variáveis principais
  #----------------------------------------------------------#

  dado_original <- dado_original |>
    mutate(

      ano =
        as.integer(
          ano
        ),

      municipio =
        as.character(
          municipio
        ),

      nome_municipio =
        as.character(
          nome_municipio
        ),

      regiao_intermediaria =
        as.character(
          regiao_intermediaria
        ),

      regiao_imediata =
        as.character(
          regiao_imediata
        ),

      cnae_classe =
        as.character(
          cnae_classe
        ),

      massa_salarial =
        as.numeric(
          massa_salarial
        ),

      vinculos_total =
        as.integer(
          vinculos_total
        ),

      vinculos_ativos_3112 =
        as.integer(
          vinculos_ativos_3112
        )

    )


  #----------------------------------------------------------#
  # Totais antes
  #----------------------------------------------------------#

  linhas_antes <-
    nrow(
      dado_original
    )

  massa_antes <-
    sum(
      dado_original$massa_salarial,
      na.rm = TRUE
    )

  vinculos_antes <-
    sum(
      dado_original$vinculos_total,
      na.rm = TRUE
    )

  ativos_antes <-
    sum(
      dado_original$vinculos_ativos_3112,
      na.rm = TRUE
    )


  #----------------------------------------------------------#
  # Compatibiliza
  #----------------------------------------------------------#

  dado <- compatibilizar_cnae_scn(
    dado_original
  )


  #----------------------------------------------------------#
  # Totais depois
  #----------------------------------------------------------#

  linhas_depois <-
    nrow(
      dado
    )

  massa_depois <-
    sum(
      dado$massa_salarial,
      na.rm = TRUE
    )

  vinculos_depois <-
    sum(
      dado$vinculos_total,
      na.rm = TRUE
    )

  ativos_depois <-
    sum(
      dado$vinculos_ativos_3112,
      na.rm = TRUE
    )


  #----------------------------------------------------------#
  # Identifica CNAEs sem correspondência
  #----------------------------------------------------------#

  sem_correspondencia <- dado |>
    filter(
      status_compatibilizacao ==
        "SEM_CORRESPONDENCIA"
    ) |>
    distinct(
      cnae_classe,
      cnae_classe_4dig
    )


  cnaes_sem_correspondencia[[
    as.character(
      ano_atual
    )
  ]] <- sem_correspondencia


  #----------------------------------------------------------#
  # Auditoria
  #----------------------------------------------------------#

  auditoria[[
    as.character(
      ano_atual
    )
  ]] <- data.frame(

    ano =
      ano_atual,

    linhas_antes =
      linhas_antes,

    linhas_depois =
      linhas_depois,

    diferenca_linhas =
      linhas_depois -
      linhas_antes,

    massa_antes =
      massa_antes,

    massa_depois =
      massa_depois,

    diferenca_massa =
      massa_depois -
      massa_antes,

    vinculos_antes =
      vinculos_antes,

    vinculos_depois =
      vinculos_depois,

    diferenca_vinculos =
      vinculos_depois -
      vinculos_antes,

    ativos_antes =
      ativos_antes,

    ativos_depois =
      ativos_depois,

    diferenca_ativos =
      ativos_depois -
      ativos_antes,

    cnaes_sem_correspondencia =
      nrow(
        sem_correspondencia
      )

  )


  #----------------------------------------------------------#
  # Impede alterações indevidas
  #----------------------------------------------------------#

  if (
    linhas_antes !=
      linhas_depois
  ) {

    stop(
      paste(
        "ERRO:",
        ano_atual,
        "- o número de linhas foi alterado."
      )
    )
  }


  if (
    abs(
      massa_depois -
        massa_antes
    ) > 0.01
  ) {

    stop(
      paste(
        "ERRO:",
        ano_atual,
        "- a massa salarial foi alterada."
      )
    )
  }


  if (
    vinculos_depois !=
      vinculos_antes
  ) {

    stop(
      paste(
        "ERRO:",
        ano_atual,
        "- o total de vínculos foi alterado."
      )
    )
  }


  if (
    ativos_depois !=
      ativos_antes
  ) {

    stop(
      paste(
        "ERRO:",
        ano_atual,
        "- o total de vínculos ativos foi alterado."
      )
    )
  }


  #----------------------------------------------------------#
  # Guarda base
  #----------------------------------------------------------#

  bases_compatibilizadas[[
    as.character(
      ano_atual
    )
  ]] <- dado


  rm(
    dado_original
  )

  gc()

}


#============================================================#
# 7. AUDITORIA FINAL
#============================================================#

auditoria_final <- bind_rows(
  auditoria
)


print(
  auditoria_final,
  row.names = FALSE
)


cnaes_sem_correspondencia_final <-
  bind_rows(
    lapply(
      names(
        cnaes_sem_correspondencia
      ),
      function(ano_txt) {

        dado <- cnaes_sem_correspondencia[[
          ano_txt
        ]]

        if (
          nrow(
            dado
          ) == 0
        ) {

          return(
            NULL
          )
        }


        dado$ano <-
          as.integer(
            ano_txt
          )


        dado

      }
    )
  )


#------------------------------------------------------------#
# Salva auditoria
#------------------------------------------------------------#

write.csv(
  auditoria_final,
  file.path(
    pasta_auditoria,
    "Auditoria_compatibilizacao_CNAE_SCN.csv"
  ),
  row.names = FALSE,
  fileEncoding = "UTF-8"
)


#------------------------------------------------------------#
# Se existir CNAE sem correspondência, salva e PARA.
# Não cria o Excel final incompleto.
#------------------------------------------------------------#

if (
  !is.null(
    cnaes_sem_correspondencia_final
  ) &&
    nrow(
      cnaes_sem_correspondencia_final
    ) > 0
) {


  write.csv(
    cnaes_sem_correspondencia_final,
    file.path(
      pasta_auditoria,
      "CNAEs_sem_correspondencia_SCN.csv"
    ),
    row.names = FALSE,
    fileEncoding = "UTF-8"
  )


  print(
    cnaes_sem_correspondencia_final,
    row.names = FALSE
  )


  stop(
    paste(
      "Existem CNAEs sem correspondência no tradutor.",
      "Confira CNAEs_sem_correspondencia_SCN.csv."
    )
  )

}


cat(
  "\nTODAS AS CNAEs FORAM COMPATIBILIZADAS COM O SCN.\n"
)


#============================================================#
# 8. ABRE O EXCEL FINAL ORIGINAL
#
# Assim preservamos as abas Resumo e Metodologia.
#============================================================#

wb <- loadWorkbook(
  arquivo_rais
)


#============================================================#
# 9. SUBSTITUI SOMENTE AS ABAS ANUAIS
#    POR VERSÕES COM AS NOVAS COLUNAS SCN
#============================================================#

estilo_monetario <- createStyle(
  numFmt = '#,##0.00'
)


estilo_texto <- createStyle(
  numFmt = "@"
)


for (ano_atual in anos) {


  nome_aba <-
    as.character(
      ano_atual
    )


  dado <- bases_compatibilizadas[[
    nome_aba
  ]]


  cat(
    "Exportando",
    ano_atual,
    "...\n"
  )


  #----------------------------------------------------------#
  # Remove a tabela antiga da aba, caso exista
  #----------------------------------------------------------#

  tabelas_existentes <- getTables(
    wb,
    sheet = nome_aba
  )


  if (
    length(
      tabelas_existentes
    ) > 0
  ) {

    for (
      tabela_atual
      in tabelas_existentes
    ) {

      removeTable(
        wb,
        sheet = nome_aba,
        table = tabela_atual
      )

    }

  }


  #----------------------------------------------------------#
  # Reescreve a aba com as informações antigas + SCN
  #----------------------------------------------------------#

  writeDataTable(
    wb,
    sheet = nome_aba,
    x = dado,
    startRow = 1,
    startCol = 1,
    tableName = paste0(
      "RAIS",
      ano_atual,
      "SCN"
    ),
    tableStyle = "TableStyleMedium2",
    withFilter = TRUE
  )


  #----------------------------------------------------------#
  # Congela cabeçalho
  #----------------------------------------------------------#

  freezePane(
    wb,
    sheet = nome_aba,
    firstRow = TRUE
  )


  #----------------------------------------------------------#
  # Formatos
  #----------------------------------------------------------#

  # Município
  addStyle(
    wb,
    sheet = nome_aba,
    style = estilo_texto,
    rows = 2:(
      nrow(
        dado
      ) + 1
    ),
    cols = 2,
    gridExpand = TRUE,
    stack = TRUE
  )


  # CNAE original
  addStyle(
    wb,
    sheet = nome_aba,
    style = estilo_texto,
    rows = 2:(
      nrow(
        dado
      ) + 1
    ),
    cols = 6,
    gridExpand = TRUE,
    stack = TRUE
  )


  # Massa salarial
  addStyle(
    wb,
    sheet = nome_aba,
    style = estilo_monetario,
    rows = 2:(
      nrow(
        dado
      ) + 1
    ),
    cols = 7,
    gridExpand = TRUE,
    stack = TRUE
  )


  # Novos códigos
  colunas_texto_novas <-
    which(
      names(
        dado
      ) %in%
        c(
          "cnae_classe_4dig",
          "scn_ibge",
          "scn_mip_sc"
        )
    )


  addStyle(
    wb,
    sheet = nome_aba,
    style = estilo_texto,
    rows = 2:(
      nrow(
        dado
      ) + 1
    ),
    cols = colunas_texto_novas,
    gridExpand = TRUE,
    stack = TRUE
  )


  #----------------------------------------------------------#
  # Largura das colunas
  #----------------------------------------------------------#

  setColWidths(
    wb,
    sheet = nome_aba,
    cols = 1,
    widths = 10
  )

  setColWidths(
    wb,
    sheet = nome_aba,
    cols = 2,
    widths = 14
  )

  setColWidths(
    wb,
    sheet = nome_aba,
    cols = 3,
    widths = 25
  )

  setColWidths(
    wb,
    sheet = nome_aba,
    cols = 4:5,
    widths = 30
  )

  setColWidths(
    wb,
    sheet = nome_aba,
    cols = 6,
    widths = 18
  )

  setColWidths(
    wb,
    sheet = nome_aba,
    cols = 7,
    widths = 21
  )

  setColWidths(
    wb,
    sheet = nome_aba,
    cols = 8:9,
    widths = 20
  )


  # Novas colunas de compatibilização
  setColWidths(
    wb,
    sheet = nome_aba,
    cols = 10,
    widths = 18
  )

  setColWidths(
    wb,
    sheet = nome_aba,
    cols = 11,
    widths = 45
  )

  setColWidths(
    wb,
    sheet = nome_aba,
    cols = 12,
    widths = 15
  )

  setColWidths(
    wb,
    sheet = nome_aba,
    cols = 13,
    widths = 45
  )

  setColWidths(
    wb,
    sheet = nome_aba,
    cols = 14,
    widths = 15
  )

  setColWidths(
    wb,
    sheet = nome_aba,
    cols = 15,
    widths = 45
  )

  setColWidths(
    wb,
    sheet = nome_aba,
    cols = 16,
    widths = 25
  )

}


#============================================================#
# 10. ACRESCENTA NOTA À ABA METODOLOGIA
#============================================================#

metodologia_atual <- read.xlsx(
  arquivo_rais,
  sheet = "Metodologia",
  colNames = FALSE
)


ultima_linha_metodologia <-
  nrow(
    metodologia_atual
  ) + 2


writeData(
  wb,
  sheet = "Metodologia",
  x = data.frame(
    Item =
      "Compatibilização CNAE-SCN",
    Descricao =
      paste(
        "A classe CNAE 2.0 foi compatibilizada com o SCN utilizando o arquivo",
        "'Tradutor CNAE 20 para SCN.xlsx' fornecido para o projeto.",
        "Para as classes CNAE usuais, utilizam-se os quatro primeiros dígitos",
        "da classe CNAE (retirada do dígito verificador).",
        "Os códigos auxiliares 851911 e 861911 foram compatibilizados,",
        "respectivamente, com Educação Pública (8591) e Saúde Pública (8691).",
        "As demais atividades de educação e saúde foram direcionadas,",
        "na classificação desagregada da MIP-SC, para Educação Privada (8592)",
        "e Saúde Privada (8692).",
        "A CNAE 99008 (organismos internacionais e instituições extraterritoriais)",
        "é mantida na base, mas sem código SCN doméstico, com status",
        "EXTRATERRITORIAL_SEM_SCN, pois não possui correspondência no tradutor."
      ),
    stringsAsFactors = FALSE
  ),
  startRow =
    ultima_linha_metodologia,
  startCol = 1,
  colNames = FALSE
)


setColWidths(
  wb,
  sheet = "Metodologia",
  cols = 1,
  widths = 35
)

setColWidths(
  wb,
  sheet = "Metodologia",
  cols = 2,
  widths = 100
)


#============================================================#
# 11. SALVA O EXCEL FINAL
#============================================================#

saveWorkbook(
  wb,
  arquivo_saida,
  overwrite = TRUE
)


#============================================================#
# 12. CONFERÊNCIA FINAL
#============================================================#

if (
  !file.exists(
    arquivo_saida
  )
) {

  stop(
    "O arquivo Excel final não foi criado."
  )
}


tamanho_mb <-
  file.info(
    arquivo_saida
  )$size /
  1024^2


cat(
  "\n========================================\n"
)

cat(
  "COMPATIBILIZAÇÃO CONCLUÍDA COM SUCESSO\n"
)

cat(
  "========================================\n"
)


cat(
  "\nArquivo criado em:\n",
  arquivo_saida,
  "\n"
)


cat(
  "\nTamanho:",
  round(
    tamanho_mb,
    2
  ),
  "MB\n"
)


cat(
  "\nResultado esperado da auditoria:\n"
)

cat(
  "- diferença de linhas = 0 em todos os anos\n"
)

cat(
  "- diferença de massa salarial = 0 em todos os anos\n"
)

cat(
  "- diferença de vínculos = 0 em todos os anos\n"
)

cat(
  "- diferença de ativos = 0 em todos os anos\n"
)

cat(
  "- CNAEs sem correspondência = 0\n"
)

cat(
  "\nArquivo final pronto para uso na compatibilização RAIS -> SCN/MIP-SC.\n"
)
