[README_pipeline_PNADC_MIP_SC.md](https://github.com/user-attachments/files/34384962/README_pipeline_PNADC_MIP_SC.md)
# Pipeline PNAD Contínua 2018-2025 — MIP-SC

Este diretório consolida o processamento da PNAD Contínua usado na atualização da MIP-SC.
O fluxo foi reduzido a dois scripts reproduzíveis.

## Ordem de execução

1. `01_pipeline_PNADC_2018_2025_dados.R`
   - baixa/usa cache da 5ª visita anual da PNAD Contínua;
   - mantém ocupados;
   - cria formal/informal pela proxy adotada no projeto;
   - cria rendimentos;
   - aplica salário mínimo trimestral corrigido em 2020 e 2023;
   - cria OCC/FAM;
   - agrega por Ano × UF × CNAEDOM × OCC × FAM;
   - calcula parcelas OCC/FAM;
   - gera `PNADC_2018_2025_SEM_SCN_SM_CORRIGIDO.xlsx`;
   - salva os microdados corrigidos necessários ao segundo script.

2. `02_pipeline_PNADC_CNAE_SCN_excel_final.R`
   - lê o tradutor original do professor;
   - monta o tradutor metodológico usado em `Dados_PNAD`/`Dados_SC`;
   - monta separadamente o `Tradutor atualizado` usado pela tabela `PNAD SC`;
   - desdobra educação e saúde entre público/privado usando `VD4009`;
   - preserva os totais do arquivo sem SCN;
   - gera `PNADC_2018_2025_FINAL_CNAE_SCN.xlsx` com as abas finais.

## Configuração de caminhos

Por padrão, os scripts usam:

`C:/Users/10766048977/Desktop/Projeto atualização da MIP`

Em outro computador, defina a variável de ambiente `MIP_SC_ROOT` com a pasta raiz do projeto.

O segundo script espera o tradutor original em:

`R/PNAD/Tradutor/Tradutor_CNAEDOM_SCN_professor.xlsx`

Também é possível informar outro caminho pela variável de ambiente `PNAD_TRADUTOR_XLSX`.

## Testes de regressão

Os scripts param automaticamente se deixarem de reproduzir os principais resultados do arquivo final validado, incluindo:

- 230.940 linhas no agregado sem SCN;
- 222 CNAEDOM nacionais;
- 241.204 linhas no `Dados_PNAD` final com SCN;
- 15.647 linhas no `Dados_SC` final;
- 216 CNAEDOM em Santa Catarina;
- preservação anual de amostra, TPO, TRP e TRD após a tradução;
- fechamento `total = formal + informal`;
- totais de Santa Catarina iguais aos valores do arquivo final validado;
- no `Tradutor atualizado`, nenhuma classificação preexistente pode mudar além da atualização deliberada de `01119`;
- na tabela `PNAD SC`, somente `00000` e `01999` ficam sem SCN entre os CNAEDOM observados em Santa Catarina.

## Observação importante sobre os dois tradutores

O arquivo final contém duas lógicas de tradução por decisão metodológica do projeto:

- `Dados_PNAD` e `Dados_SC`: usam a compatibilização metodológica corrigida, inclusive a separação de educação/saúde pública e privada;
- `PNAD SC`: reproduz o `Tradutor atualizado` do professor, mantendo as convenções adotadas no projeto (inclusive códigos como `4580` e a convenção histórica do tradutor manual).

Essa separação é intencional e está documentada no script 02.

## Reprodutibilidade

Cada script salva um `sessionInfo()` e arquivos CSV de auditoria na pasta `R/PNAD/Logs`.
Para máxima reprodutibilidade no GitHub, recomenda-se também versionar o tradutor de entrada e registrar a versão dos pacotes usada na execução final.
