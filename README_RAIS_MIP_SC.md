# Pipeline RAIS-SC 2018–2025 — Projeto MIP-SC

Este repositório contém o pipeline utilizado para construir e compatibilizar a base da **RAIS de Santa Catarina (2018–2025)** para o projeto de atualização da **Matriz Insumo-Produto de Santa Catarina (MIP-SC)**.

O processamento foi organizado em duas etapas principais:

1. construção da base RAIS de Santa Catarina por município e classe CNAE;
2. compatibilização da CNAE 2.0 com as atividades do SCN/MIP-SC.

---

## Scripts principais

### 1. `Scrit completo RAIS SC (2018-2025) CNAE x Massa salarial x Dados muncipais.R`

Pipeline principal da RAIS.

Executa:

- download dos microdados da RAIS da Região Sul para 2018–2025;
- leitura apenas das variáveis necessárias;
- filtro dos estabelecimentos localizados em Santa Catarina;
- harmonização das diferenças de layout entre 2018–2022 e 2023–2025;
- reclassificação de Educação Pública e Saúde Pública;
- tratamento e harmonização das remunerações mensais;
- correção de escala das remunerações de janeiro a novembro em 2018–2021;
- cálculo da massa salarial anual;
- filtro de registros com remuneração média em salários mínimos maior que zero;
- padronização das classes CNAE;
- tratamento de códigos residuais;
- agregação por Município × CNAE × Ano;
- incorporação do nome do município e das Regiões Geográficas Imediata e Intermediária;
- auditorias internas;
- comparação com o benchmark da versão final validada;
- exportação do Excel final e dos arquivos RDS anuais.

### 2. `Compatibilizacao_CNAE_SCN_RAIS_SC_2018_2025.R`

Executa a compatibilização da base final da RAIS com o **SCN**.

Principais procedimentos:

- preserva a classe CNAE original de 5 dígitos;
- cria a CNAE de 4 dígitos para ligação com o tradutor;
- adiciona o código de atividade SCN padrão do IBGE;
- adiciona o código SCN desagregado utilizado na MIP-SC;
- trata separadamente Educação Pública, Educação Privada, Saúde Pública e Saúde Privada;
- mantém organismos internacionais/instituições extraterritoriais como categoria sem SCN doméstico;
- audita a correspondência CNAE–SCN;
- verifica se massa salarial, número de vínculos e vínculos ativos permanecem inalterados após a compatibilização.

---

## Ordem de execução

Execute os scripts nesta ordem:

```text
1. Scrit completo RAIS SC (2018-2025) CNAE x Massa salarial x Dados muncipais.R
2. Compatibilizacao_CNAE_SCN_RAIS_SC_2018_2025.R
```

O primeiro script constrói a base RAIS final por município e CNAE.

O segundo usa essa base como entrada e acrescenta as classificações SCN.

---

## Principais fontes

### RAIS — Ministério do Trabalho e Emprego

Os microdados utilizados são os arquivos da RAIS Vínculos Públicos da Região Sul:

```text
ftp://ftp.mtps.gov.br/pdet/microdados/RAIS/<ANO>/RAIS_VINC_PUB_SUL.7z
```

### Divisão Territorial Brasileira — IBGE

O tradutor territorial utiliza a **DTB 2025**, para identificação de municípios e das Regiões Geográficas Imediata e Intermediária de Santa Catarina.

### Tradutor CNAE → SCN

A compatibilização CNAE–SCN utiliza o arquivo:

```text
Tradutor CNAE 20 para SCN.xlsx
```

fornecido para o projeto.

---

## Estrutura principal de saída

O pipeline principal cria, entre outros produtos:

```text
Resultados RAIS/
├── Bases finais RDS/
├── Excel Final/
│   └── RAIS_SC_2018_2025_FINAL.xlsx
└── Auditoria/
```

Após a compatibilização CNAE–SCN, é criado:

```text
Resultados RAIS/Excel Final/
└── RAIS_SC_2018_2025_FINAL_CNAE_SCN.xlsx
```

---

## Unidade de análise

A base final antes da compatibilização com o SCN está agregada em:

```text
Ano × Município × Classe CNAE
```

As principais variáveis são:

- massa salarial anual;
- número total de vínculos;
- vínculos ativos em 31/12;
- município;
- nome do município;
- Região Geográfica Imediata;
- Região Geográfica Intermediária;
- classe CNAE.

Após a compatibilização, também são adicionados os códigos e descrições do SCN.

---

## Decisões metodológicas importantes

### Santa Catarina

O recorte territorial utiliza o **município do estabelecimento**.

### Massa salarial

A massa salarial corresponde à soma das remunerações nominais mensais de janeiro a dezembro dos vínculos considerados.

### Correção das remunerações de 2018–2021

Para janeiro a novembro de 2018–2021, valores positivos inferiores a 100 foram multiplicados por 100 após a identificação de problema de escala nas variáveis originais.

Dezembro não recebe essa correção.

### Filtro dos vínculos

São mantidos na agregação os registros com:

```text
rem_media_sm > 0
```

### CNAE

A classe CNAE é preservada como texto para evitar perda de zeros à esquerda.

Também são tratados códigos residuais identificados durante a auditoria dos microdados.

### Educação e Saúde públicas

Foram criados códigos auxiliares no processamento da RAIS:

```text
851911 = Educação Pública
861911 = Saúde Pública
```

Na compatibilização com o SCN/MIP-SC:

```text
8591 = Educação pública
8592 = Educação privada
8691 = Saúde pública
8692 = Saúde privada
```

As demais atividades de educação e saúde são direcionadas para os setores privados na classificação desagregada da MIP-SC.

### Organismos internacionais

A CNAE `99008`, referente a organismos internacionais e instituições extraterritoriais, é preservada na base, mas permanece sem código SCN doméstico.

---

## Auditorias

O pipeline possui diversas verificações automáticas.

Entre elas:

- número de municípios de Santa Catarina;
- duplicações Município × CNAE;
- CNAEs inválidas;
- massa salarial negativa;
- vínculos negativos;
- vínculos ativos maiores que vínculos totais;
- correspondência territorial;
- número de linhas;
- massa salarial;
- vínculos totais;
- vínculos ativos;
- quantidade de CNAEs distintas;
- códigos residuais;
- hashes SHA-256 da versão final validada.

O pipeline possui um benchmark da execução final validada para 2018–2025.

Na etapa CNAE–SCN, o script também verifica se a compatibilização altera indevidamente:

```text
massa salarial
vínculos totais
vínculos ativos
número de linhas
```

Se houver CNAE sem correspondência no tradutor, a exportação final é interrompida para evitar a geração de uma base incompleta.

---

## Dependências principais

Os scripts utilizam principalmente os pacotes:

```r
data.table
archive
readODS
openxlsx
digest
dplyr
```

---

## Reprodutibilidade

O objetivo desses scripts é permitir que a base final da RAIS utilizada no projeto MIP-SC possa ser reproduzida a partir dos microdados oficiais.

O pipeline inclui auditorias e benchmarks para identificar eventuais diferenças caso os microdados oficiais sejam revisados ou substituídos futuramente.

---

## Projeto

**Atualização da Matriz Insumo-Produto de Santa Catarina — MIP-SC**

Período da RAIS utilizado:

```text
2018–2025
```
