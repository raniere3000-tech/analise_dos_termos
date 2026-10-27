# Painel de termos aplicados em 2026 e adimplência

Contexto para continuar o desenvolvimento no Claude Code.

## Regras de trabalho

- **Pergunte antes de executar.** Quando houver dúvida sobre regra de negócio, layout ou formato, pergunte antes de mudar o código.
- **Valide os números.** Confira toda mudança de cálculo com `tests/conferir_totais.py`, que recalcula o Resumo em Python de forma independente.
- **Mantenha um arquivo só.** O painel é um único `index.html`, sem build e sem dependências além das Google Fonts.
- **Escreva sempre em português do Brasil.** Interface, comentários, mensagens de commit e **toda a comunicação com o usuário**: o passo a passo do que está sendo feito, os resumos, as explicações e as perguntas. O usuário pediu isso mais de uma vez; nunca responda em inglês.
- **Aumente a versão a cada atualização.** A constante `VERSAO_SITE` no `index.html` aparece no cabeçalho.
- **Estilo:** o mesmo do painel de economias de esgoto (repositório economias), que copia o painel de faturamento: IBM Plex Sans, marinho `#102840`, azul `#176b9c`, fundo `#f3f6f9`; positivo em marinho e azul, negativo em laranja (`#E8710A`, texto `#C2560C`). O leitor de .xlsx em fluxo, os filtros, o gráfico com tabela alinhada e a dica ao passar o mouse vieram de lá.

## Objetivo

Identificar os clientes que receberam termos (sanções) em 2026 e, a partir do termo, deixaram de ser adimplentes, para avaliar se vale a pena aplicar o termo com o risco de perder o cliente pagante.

## Dados de entrada

- **Atualização:** o usuário só acrescenta os arquivos novos nas pastas e clica em ↻ Atualizar; pagamentos repetidos entre a arrecadação diária e a acumulada entram uma vez. O faturamento vem até o último ciclo liberado.
- **Arrecadação:** de jan/2025 até o mês atual. Meses fechados na pasta `arrecadação acumulada`, mês atual na pasta `arrecadação diária`. Colunas da exportação: `Nº Ligação, Nº Nota, Mês, Ano, Valor Pago, Vencimento, Data Pagto., Data Atualiz., Data Crédito, Val. Autenticado, Utilização (Sub. Categ.), Zona (Localidade), Classificação, Nº Contrato, Cliente, Tipo Fatura, Situação Pagto, Banco, Tipo Pagamento`. Formatos .xlsx e .csv.
- **Termos (desde a v1.1):** arquivo com "termo" no nome (hoje `termos_todos_os_termos_AAAA-MM-DD.xlsx`). Lê **só a aba Termos**; a aba Resumo é ignorada (`Termos.abas()`). Colunas da exportação: `Nº O.S., Matrícula, Frente, Cidade, Zona, Equipe, Tipo de atividade, Status da atividade, Serviço adicional identificado, Data de aplicação, Data de notificação, Data de negociação, Nome do solicitante, Categoria, Situação da ligação, Situação do termo, Faturado, Dias, Faixa, Valor, Isento, Motivo da isenção, Origem, Arquivo da aplicação, Arquivo de status, Arquivo de faturamento`.
  - `Matrícula` é o mesmo número que o `Nº Ligação` da arrecadação (confirmado pelo usuário).
  - **Mês do termo = Data de negociação** (padrão, definido pelo usuário), com chave **Data do termo** na barra superior para usar a Data de aplicação (v2.6). Quando a data escolhida está vazia, vale a outra. O worker guarda as linhas de termo com as duas datas (`tMat, tNeg, tApl, tNegTxt, tAplTxt, tValor, tTipo, tCid, tExtra`) e a página monta os eventos em `montaEventos()`; só entram termos de 2026 pela data escolhida.
  - **Termo com Valor = 0 fica fora da análise** (v2.1, definido pelo usuário: são isentos ou cancelados). O rodapé mostra quantos.
  - Tipo de termo = `Tipo de atividade`; valor = `Valor`. Fora o Valor 0, nenhum termo é excluído por status, situação ou isenção.
  - Filtros vindos da base de termos (`EXTRAS`): Situação do termo, Faturado, Isento, Frente, Equipe, Categoria e Origem. Coluna ausente no arquivo não vira filtro. Também aparecem no Consolidado e no CSV.
  - A coluna `Faixa` da base de termos não tem relação com as faixas de adimplência do painel e não é lida.
- **Faturamento (v2.1):** subpasta `Faturamento`, reconhecida pelo cabeçalho (tabela com a coluna `Rubrica`). Na base real são CSV de ~82 MB por mês (`Fatura de ciclo MM-AAAA.csv`, de 01-2025 em diante); o formato exato (separador, codificação, datas) ainda não foi confirmado. Colunas da exportação: `N. da Ligacao, Situacao Lancamento, Nome Cliente, Categoria, Qtd. Economia Residencial, Qtd. Economia Comercial, Qtd. Economia Industrial, Qtd. Economia Publica, Qtd. Economia Outros, Nome Logradouro, Numero da Porta, Desc. Complemento, Nome do Bairro, Nome da Localidade, Endereco Ligacao, Grupo, Rota, Rubrica, Valor Parcela, Ocorrencia, Situacao de Repasse, Faturamento, Situacao Nota Liberada, Referencia de Leitura, Data de Vencimento, Situacao Ligacao, Situacao Conta`.
  - Só as rubricas `VALOR DE AGUA` e `VALOR DE ESGOTO` (coluna Valor Parcela); quem paga esgoto tem as duas linhas e a conta do mês é a soma (definido pelo usuário). O usuário escreveu "coluna I", mas a Rubrica é a coluna R; o painel lê pelo nome.
  - **Mês da conta = mês da Data de Vencimento** (definido pelo usuário). Todas as situações de conta entram (os meses analisados são fechados).
  - Nome Cliente do faturamento preenche o nome de quem não tem pagamento na arrecadação.
  - **Ligação cortada (v2.6, definido pelo usuário):** pela coluna `Situacao Ligacao` da **última conta** da matrícula no faturamento (`fatUlt` = mês da última conta, `fatSitLig` = situação). Os valores que indicam corte são escolhidos em "Situações de corte" nos parâmetros (`P.corteVals`; padrão: os que têm CORT no texto).
- Nomes de coluna são comparados sem acento, maiúsculas e pontuação (`norm()`: "Nº Ligação" vira `N_LIGACAO`).
- Número da ligação normalizado em `ligKey()` (sem zeros à esquerda), igual nas duas bases.

## Arquitetura (index.html)

- **Worker inline** (`#worker-src`): lê primeiro a base de termos e depois os demais arquivos, guardando **só os dados das matrículas com termo** (a base inteira teria milhões de linhas). Pagamentos repetidos entre arquivos (mesma ligação, nota, referência, data e valor) entram uma vez. Matrizes matrícula × mês (`m * nR + r`): `valRef` (referência Mês/Ano), `valPag` (Data Pagto.), `valCred` (Data Crédito; sem ela, Data Pagto.), `fat` (conta do mês pelo vencimento) e `credVenc` (v2.2: para cada matrícula e mês de vencimento, o primeiro mês em que entrou crédito daquela conta, pela coluna Vencimento da arrecadação; -1 se nunca; `temVenc` diz se a coluna existe). `porMes`, `porRef` e `porCred` (linhas da base inteira por mês) e `stats.fat.porMes` servem para achar meses sem dados.
- **Faturamento no worker:** `tabela()` escolhe os nomes de coluna pelo cabeçalho (com `Rubrica` é faturamento). `preFiltro()` (.xlsx) e `preFiltroCsv()` (.csv) leem só a ligação e pulam as linhas de matrículas sem termo antes de ler o resto (o faturamento inteiro tem dezenas de milhões de linhas).
- **Leitura tolerante (v2.3):** CSV em UTF-8, Windows-1252 ou UTF-16 (pelo BOM); separador `;`, tabulação, `,` ou `|`, fixado na primeira linha com 3 ou mais separadores (pula títulos); `toNum()` aceita `1.234,56` e `1,234.56`; `dataInfo()` aceita `dd/mm/aaaa`, `dd-mm-aaaa`, `dd.mm.aaaa`, `dd/mm/aa`, `aaaa-mm-dd` (com hora) e `aaaammdd`; `ligKey()` aceita `1.234.567`.
- **Diagnóstico do faturamento (v2.3):** `Faturamento.st` guarda linhas lidas, linhas de matrículas com termo, rubricas encontradas, exemplos de ligação e de vencimento não lidos e colunas faltando; `motivoSemFat()` mostra no aviso amarelo e no rodapé por que o faturamento não foi usado. `stats.exLig` traz exemplos de ligação dos termos e da arrecadação para comparar.
- **Período do estudo (v2.4, definido pelo usuário):** de jan/2025 (`INICIO_ESTUDO`, ano dos termos − 1) até o mês de hoje (`HOJE_AM`). Arrecadação com data fora disso (o usuário tinha datas de 2022 a 2024, sem dados reais, e pode haver datas futuras digitadas errado) é ignorada e contada em `stats.foraPeriodo`; vencimentos do faturamento fora do período também (`stats.fat.fora`); termos só de 2026. Na página, `recortaPeriodo()` aplica o mesmo recorte a uma base guardada por versão anterior, sem precisar ler os arquivos de novo (foi o que o usuário viu: a base guardada pela v2.2 abria com meses desde 2022).
- **Termos como eventos (v2.0):** cada matrícula guarda a lista dos seus termos; termos da mesma matrícula no mesmo mês viram um evento só (`eQtd`). Arrays por evento: `eMat, eAm, eData, eSeq` (ordem na matrícula), `eNum`, `eProx` (mês do termo seguinte ou -1), `eTipo, eCid, eValor, extra[k]`.
- **IndexedDB** `painel-termos-2026`, chaves `atual` (base) e `pasta`. Base lida por versão anterior (sem `eMat`) pede nova leitura. Parâmetros em `localStorage` (`painel-termos-param-v4`). O mês vigente (`st.vig`) não é guardado: o painel abre sempre no mês de hoje.
- **`derive()`** monta as unidades (`D.U`: eventos analisados) e classifica: `nA/pA/sA` meses, meses pagos e valor antes; `nD/pD/sD` depois; `cA` (0 Adimplente, 1 Flutuante, 2 Inadimplente, 3 Nunca pagou, -1 sem histórico); `cD` (0 a 3, 4 acompanhamento insuficiente); `res` (0 permaneceu, 1 piorou, 2 melhorou, 3 insuficiente, -1 sem histórico); `mA/mD` médias mensais; `fimD` último mês do período depois; `corte` (encurtado por termo seguinte); `semDados` por mês.
- **`agrega(k, chave, nK)`** soma por grupo (total, mês do termo, classe anterior, município, tipo, zona, utilização).
- **Filtros (v2.7, `defFiltros`)**: cada filtro tem grupo (Termo, Resultado, Cadastro e ligação, Equipe e origem), lista e `get(u)`; os que existem para a amostra sem termo têm também `getC(u)`. Os **principais** (`PRINCIPAIS`, escolhidos pelo usuário: Mês do termo, Situação do termo, Isento, Antes, Depois, Resultado) ficam à vista na faixa "Filtros"; **todos** ficam na gaveta "Todos os filtros" (`montaGaveta()`), agrupados. Filtros ativos aparecem como etiquetas removíveis (`marcaAtivos()`). A aba Clientes usa os filtros gerais (não tem mais filtros próprios); clicar num número da matriz aplica os filtros Antes e Depois.
- **Fluxo antes → depois (v2.8, pedido do usuário com um modelo do SankeyMATIC):** diagrama de Sankey desenhado em SVG por `renderSankey(M, linhas)`, sem biblioteca, no cartão da matriz, logo abaixo dela. Usa a mesma contagem `M[antes][depois]` da matriz (os totais sempre batem): à esquerda as classificações antes (e Sem histórico antes), à direita as classificações depois (e Acompanhamento insuficiente). Cor pela classificação antes (padrão; faixas em `SK_FAIXA`, blocos em `SK_COR`) ou pelo resultado (`st.sk`). Dica ao passar o mouse; clique aplica os filtros Antes e Depois e abre a aba Clientes.
- **Abas e configuração (v2.7):** barra de abas própria (Resumo, Clientes, Comparação sem termo), separada da faixa "Configuração da análise" (Base de cálculo, Mês vigente, Data do termo, Mínimo de meses depois) e da faixa "Filtros".
- **Amostra de comparação (v2.7):** o worker registra também matrículas sem termo sorteadas por `hashLig()` (FNV-1a do número da ligação, 5%, até 60 mil; `AMOSTRA_PCT`, `AMOSTRA_MAX`), vistas na arrecadação ou no faturamento (`Arrecadacao.registra()`). `D.nTermo` separa as matrículas com termo (índices menores) das da amostra. Na página, cada matrícula da amostra recebe um mês de referência sorteado (pelo mesmo hash) entre os meses das unidades com termo e passa por `classifica()`, a mesma função dos clientes com termo; o resultado fica em `D.C`.

## Regras de negócio (v2.0, definidas com o usuário)

- **Unidade de análise:** padrão é a matrícula única com o primeiro termo do ano, **acompanhado até o fim mesmo que venha outro termo** (v2.6, definido pelo usuário); o painel avisa quantas matrículas são reincidentes e tem o filtro Reincidente. Por evento (parâmetro), o período depois termina no mês anterior ao termo seguinte da mesma matrícula.
- **Ligação cortada depois do termo (v2.6, definido pelo usuário):** com faturamento, para matrícula cortada, os meses do período depois sem conta e posteriores à última conta contam como observados e **não pagos** (`D.ncD`). Assim, quem não pagou nada fica Nunca pagou e quem pagou parte fica Inadimplente ou Flutuante. Só vale para o período depois. Coluna e filtro "Ligação cortada" e "corte" na grade de clientes.
- **Mínimo de meses depois:** na barra superior (v2.6), padrão 3.
- **Base de cálculo (v2.7, definido pelo usuário):** botões "Com faturamento" / "Só arrecadação" na faixa de configuração (`P.fat`, guardado no navegador: abre com a última escolha). Só arrecadação = todo mês conta e o mês é pago se houve crédito nele (Data Crédito); sem meses sem conta, sem conta paga pelo vencimento e sem regra de corte.
- **Comparação sem termo (v2.7, definido pelo usuário, em aba separada):** mesmos indicadores para os clientes com termo e para a amostra sem termo, com a diferença em pontos percentuais; gráficos de quem piorou e melhorou por classificação antes; matrizes em % da linha. Na aba só valem os filtros que existem para os dois grupos (os outros são listados como ignorados). Aviso quando a amostra tem menos de 200 comparáveis.
- **Mês vigente (v2.2, definido pelo usuário):** seletor na barra superior, abre no mês de hoje. É o mês ainda não fechado: termos com mês ≥ mês vigente ficam fora (`D.foraVig`), a análise vai até o mês anterior (`idxUltimo()`), e os créditos do mês vigente (arrecadação diária) só servem para pagar com atraso contas dos meses anteriores. Contas com vencimento no mês vigente não entram.
- **Conta paga (v2.2, definido pelo usuário):** com faturamento e a coluna Vencimento na arrecadação (`D.porConta`), a conta que vence no mês r está paga se entrou crédito daquela conta (mesmo mês de vencimento na arrecadação) até o fim do mês r + tolerância e no máximo no mês vigente. Tolerância padrão: até o fim do mês seguinte (parâmetro "Pagamento com atraso aceito": só no mês, mês seguinte, 2 meses, qualquer data). Vale igual para todos os meses. Pagamento muito atrasado (dívida negociada depois do termo) não limpa o período antes do termo.
- **Mês pago sem faturamento** (ou sem a coluna Vencimento): soma do Valor Pago no mês > 0, pela **Data Crédito** (definido pelo usuário na v2.1); parâmetros permitem a Data Pagto. ou a referência da fatura (Mês/Ano). Com faturamento e Vencimento, esse parâmetro fica desligado na tela.
- **Meses observados (v2.1):** com faturamento, só os meses em que a matrícula tinha conta a pagar (conta > 0 com vencimento no mês). Mês sem conta não conta como "não pagou" nem como pago; o valor arrecadado nele entra só na média mensal. Parâmetro "Meses observados" permite usar todos os meses. Matrícula sem nenhuma conta antes do termo fica "Sem histórico antes".
- **Mês sem dados:** mês sem nenhum pagamento de nenhum cliente na base inteira (ou, com faturamento, sem nenhuma conta); fica fora das janelas e aparece no aviso.
- **Média mensal:** valor arrecadado no período dividido pelos meses do período com dados (inclui meses sem conta).
- **Antes:** os 12 meses anteriores ao mês do termo (o mês do termo não entra). 10 a 12 pagos: Adimplente; 6 a 9: Flutuante; 1 a 5: Inadimplente; 0: Nunca pagou. Com meses sem dados, a mesma regra em proporção (≥ 10/12, ≥ 6/12).
- **Depois:** do mês seguinte ao termo até o último mês completo, no máximo 12 meses. % de meses pagos = pagos / observados. Adimplente ≥ 83,3%; Flutuante ≥ 50%; Inadimplente > 0%; Nunca pagou 0%. Menos de 3 meses observados: **Acompanhamento insuficiente**, fora de melhora, piora e permanência. Limiares e mínimo são parâmetros.
- **Resultado:** ordem Adimplente → Flutuante → Inadimplente → Nunca pagou; ir para trás é melhora, para frente é piora.
- **Acompanhamento insuficiente (v2.5):** conta todos com menos meses observados que o mínimo depois do termo, **inclusive os sem histórico antes**; é o mesmo número no card, na coluna da matriz, no gráfico de perfil e nas tabelas (antes da v2.5 o card e o gráfico deixavam os sem histórico de fora e o usuário viu 835 na matriz e 677 no gráfico). `D.motivo` diz o porquê: 1 termo recente (menos meses fechados que o mínimo entre o termo e o último mês), 2 período encurtado por outro termo, 3 meses sem dados, 4 sem conta a pagar no faturamento depois do termo. `D.motivoA`: sem histórico por falta de dados (1) ou por não ter conta a pagar nos 12 meses antes (2). Os motivos aparecem abaixo da matriz e na coluna Motivo da tabela de clientes.
- **Premissa revista (v2.6):** na base real (v2.5), 835 de 1.328 ficaram com acompanhamento insuficiente: 459 termo recente, 66 período encurtado por outro termo, 310 sem conta a pagar depois do termo; e 183 sem histórico antes (todos sem conta nos 12 meses antes). Decisões do usuário: primeiro termo até o fim (acaba com os 66), cortadas contam meses sem conta como não pagos (parte dos 310), chave entre Data de negociação e de aplicação, mínimo de meses na barra superior, período depois começa no mês seguinte ao termo (já era assim), e análise à parte dos sem histórico.
- **Sem histórico antes (v2.6):** cartão próprio no Resumo com a classificação depois do termo, já que não há um antes para comparar.
- **Comparáveis:** com classificação antes e depois (não insuficiente). São o denominador de melhora, piora e permanência.
- **Recuperados:** Inadimplente ou Nunca pagou antes que passaram a Adimplente ou Flutuante, sobre os comparáveis desses dois grupos. "Pagaram ao menos 1 mês" é mostrado separado (registro de pagamento, não recuperação).
- **Sem faturamento** (base ausente ou parâmetro desligado): todo mês é observado, e mês sem pagamento conta como não pago mesmo sem conta a pagar; o painel mostra esse aviso.

## Layout (v2.0)

- Barra de filtros com os 7 principais e "Mais filtros"; botão "Limpar filtros".
- Resumo: 7 cards (total, acompanhamento suficiente, melhoraram, pioraram, permaneceram, insuficiente, % recuperados), com denominadores; card de regras e limitação, com parâmetros; **matriz antes × depois** (quantidade ou % da linha; verde melhora, amarelo permanece, vermelho piora, cinza insuficiente; clique abre a tabela de clientes); gráfico A (perfil antes e depois, quantidade ou %, na mesma base da matriz, com as categorias Sem histórico antes e Acompanhamento insuficiente); gráfico C (recuperação); gráfico B (resultado por mês de aplicação, com meses acompanhados); gráfico D (média mensal por cliente antes e depois, por classe anterior); tabela por Município / Tipo / Zona / Utilização; % que pagou mês a mês em relação ao termo; rodapé.
- Clientes: uma linha por unidade, com classificações, meses observados e pagos, meses sem conta, % pago depois, resultado, valores e o valor pago em cada mês (laranja-claro sem pagamento no período, contorno no mês do termo, "s/d" sem dados, "s/c" sem conta). Exporta CSV (`;`, BOM, vírgula decimal), com colunas PAGO e CONTA por mês.
- Favicon (v2.1): SVG dentro do próprio `index.html` (lupa branca com gota verde-água num quadrado marinho `#14375e`), simplificado a partir do modelo do usuário (sem a moeda e a segunda gota, que somem em 16 px).
- Cores de resultado em `:root`: `--melhora #1F8A4C`, `--permanece #E0B43C`, `--piora #C62828`, `--insuf #9AA6B2` (pedido do usuário na v2.0).

## Versões anteriores

- v1.1 guardada em `versoes/v1.1/index.html` (abre pelo GitHub Pages em `/versoes/v1.1/`) e no branch `versao-1.1`.
- v2.0 guardada em `versoes/v2.0/index.html` e no branch `versao-2.0`; v2.1 a v2.7 em `versoes/` e nos branches `versao-2.1` a `versao-2.7`. O branch `v25-envio` sobrou de um teste de envio (o proxy não deixa apagar branches); pode ser apagado pelo GitHub. Guarde a versão anterior assim a cada mudança grande. As tags `v1.0` e `v1.1` existem no repositório local, mas o envio de tags é bloqueado pelo proxy da sessão.

## Pendências

- **Faturamento não reconhecido na base real (v2.2):** o aviso dizia "a pasta não tem base de faturamento", com os CSV na pasta. A v2.3 lê mais formatos e mostra o motivo; pedir ao usuário o texto do aviso ou as primeiras linhas de um CSV, se continuar.
- Conferir com a base real: nomes das rubricas (`VALOR DE AGUA`, `VALOR DE ESGOTO`), formato da Data de Vencimento e da Data Crédito, se o Vencimento da arrecadação bate com a Data de Vencimento do faturamento (na grade de Clientes, a dica do mês mostra em que mês a conta foi paga), e as situações de conta listadas no rodapé.

## Como validar

```bash
python tests/gerar_base_teste.py base_teste                   # base sintética (requer openpyxl)
python tests/conferir_totais.py base_teste 2026-10                  # padrão: mês vigente out/26, crédito, matrícula, faturamento, atraso até o mês seguinte
python tests/conferir_totais.py base_teste 2026-10 cred mat 1 0     # conta paga só no mês do vencimento
python tests/conferir_totais.py base_teste 2026-10 cred mat 0       # sem faturamento (todos os meses observados)
python tests/conferir_totais.py base_teste 2026-10 cred evento 1 1  # cada termo como evento
python tests/conferir_totais.py base_teste 2026-10 cred mat 1 1 apl    # mês do termo pela Data de aplicação
python tests/conferir_totais.py base_teste 2026-10 cred mat 1 1 neg 2  # mínimo de 2 meses depois
```

Valores de referência (v2.6 e v2.7) com a base sintética padrão (3000 ligações, semente 7, com Faturamento e Situacao Ligacao CORTADA ou SUPRIMIDA na última conta antes do corte), mês vigente out/26:
- padrão: 636 clientes (23 sem histórico), 407 comparáveis, 32 melhoraram, 197 pioraram, 178 permaneceram, 225 insuficientes (213 termo recente, 12 sem conta), 13 recuperados de 114; 52 cortadas com meses de corte, 37 em Nunca pagou; 28 reincidentes;
- Data de aplicação: 597, 400, 32, 191, 177, 193; 12 de 111;
- mínimo de 2 meses: 467 comparáveis, 37, 222, 208, 161; 16 de 128;
- por evento: 657 termos, 400, 31, 194, 175, 252 (223 recente, 17 outro termo, 12 sem conta); 12 de 110;
- sem faturamento: 423, 62, 190, 171, 213; 25 de 136.
- amostra sem termo (v2.7): 103 matrículas, 62 comparáveis; com faturamento 3 melhoraram, 20 pioraram, 39 permaneceram, 1 recuperado de 16; só arrecadação 9, 22, 31, 4 de 19.

No teste com Playwright, `setInputFiles` descarta arquivos cujo caminho tem acento (as pastas "Arrecadação…"): copie a base para pastas sem acento antes de testar.

Nunca suba arquivos `.xlsx` ou `.csv` para o repositório: o `.gitignore` já bloqueia.
