# Painel de termos aplicados em 2026 e adimplência

Contexto para continuar o desenvolvimento no Claude Code.

## Regras de trabalho

- **Pergunte antes de executar.** Quando houver dúvida sobre regra de negócio, layout ou formato, pergunte antes de mudar o código.
- **Valide os números.** Confira toda mudança de cálculo com `tests/conferir_totais.py`, que recalcula o Resumo em Python de forma independente.
- **Mantenha um arquivo só.** O painel é um único `index.html`, sem build e sem dependências além das Google Fonts.
- **Escreva em português do Brasil.** Interface e comentários em português.
- **Aumente a versão a cada atualização.** A constante `VERSAO_SITE` no `index.html` aparece no cabeçalho.
- **Estilo:** o mesmo do painel de economias de esgoto (repositório economias), que copia o painel de faturamento: IBM Plex Sans, marinho `#102840`, azul `#176b9c`, fundo `#f3f6f9`; positivo em marinho e azul, negativo em laranja (`#E8710A`, texto `#C2560C`). O leitor de .xlsx em fluxo, os filtros, o gráfico com tabela alinhada e a dica ao passar o mouse vieram de lá.

## Objetivo

Identificar os clientes que receberam termos (sanções) em 2026 e, a partir do termo, deixaram de ser adimplentes, para avaliar se vale a pena aplicar o termo com o risco de perder o cliente pagante.

## Dados de entrada

- **Arrecadação:** de jan/2025 até o mês atual. Meses fechados na pasta `arrecadação acumulada`, mês atual na pasta `arrecadação diária`. Colunas da exportação: `Nº Ligação, Nº Nota, Mês, Ano, Valor Pago, Vencimento, Data Pagto., Data Atualiz., Data Crédito, Val. Autenticado, Utilização (Sub. Categ.), Zona (Localidade), Classificação, Nº Contrato, Cliente, Tipo Fatura, Situação Pagto, Banco, Tipo Pagamento`. Formatos .xlsx e .csv.
- **Termos (desde a v1.1):** arquivo com "termo" no nome (hoje `termos_todos_os_termos_AAAA-MM-DD.xlsx`). Lê **só a aba Termos**; a aba Resumo é ignorada (`Termos.abas()`). Colunas da exportação: `Nº O.S., Matrícula, Frente, Cidade, Zona, Equipe, Tipo de atividade, Status da atividade, Serviço adicional identificado, Data de aplicação, Data de notificação, Data de negociação, Nome do solicitante, Categoria, Situação da ligação, Situação do termo, Faturado, Dias, Faixa, Valor, Isento, Motivo da isenção, Origem, Arquivo da aplicação, Arquivo de status, Arquivo de faturamento`.
  - `Matrícula` é o mesmo número que o `Nº Ligação` da arrecadação (confirmado pelo usuário).
  - **Mês do termo = Data de negociação** (definido pelo usuário). Linha sem Data de negociação usa a Data de aplicação (confirmado na v2.1; em geral são termos isentos ou cancelados, com Valor 0).
  - **Termo com Valor = 0 fica fora da análise** (v2.1, definido pelo usuário: são isentos ou cancelados). O rodapé mostra quantos.
  - Tipo de termo = `Tipo de atividade`; valor = `Valor`. Fora o Valor 0, nenhum termo é excluído por status, situação ou isenção.
  - Filtros vindos da base de termos (`EXTRAS`): Situação do termo, Faturado, Isento, Frente, Equipe, Categoria e Origem. Coluna ausente no arquivo não vira filtro. Também aparecem no Consolidado e no CSV.
  - A coluna `Faixa` da base de termos não tem relação com as faixas de adimplência do painel e não é lida.
- **Faturamento (v2.1):** subpasta `Faturamento`, reconhecida pelo cabeçalho (tabela com a coluna `Rubrica`). Colunas da exportação: `N. da Ligacao, Situacao Lancamento, Nome Cliente, Categoria, Qtd. Economia Residencial, Qtd. Economia Comercial, Qtd. Economia Industrial, Qtd. Economia Publica, Qtd. Economia Outros, Nome Logradouro, Numero da Porta, Desc. Complemento, Nome do Bairro, Nome da Localidade, Endereco Ligacao, Grupo, Rota, Rubrica, Valor Parcela, Ocorrencia, Situacao de Repasse, Faturamento, Situacao Nota Liberada, Referencia de Leitura, Data de Vencimento, Situacao Ligacao, Situacao Conta`.
  - Só as rubricas `VALOR DE AGUA` e `VALOR DE ESGOTO` (coluna Valor Parcela); quem paga esgoto tem as duas linhas e a conta do mês é a soma (definido pelo usuário). O usuário escreveu "coluna I", mas a Rubrica é a coluna R; o painel lê pelo nome.
  - **Mês da conta = mês da Data de Vencimento** (definido pelo usuário). Todas as situações de conta entram (os meses analisados são fechados).
  - Nome Cliente do faturamento preenche o nome de quem não tem pagamento na arrecadação.
- Nomes de coluna são comparados sem acento, maiúsculas e pontuação (`norm()`: "Nº Ligação" vira `N_LIGACAO`).
- Número da ligação normalizado em `ligKey()` (sem zeros à esquerda), igual nas duas bases.

## Arquitetura (index.html)

- **Worker inline** (`#worker-src`): lê primeiro a base de termos e depois os demais arquivos, guardando **só os dados das matrículas com termo** (a base inteira teria milhões de linhas). Pagamentos repetidos entre arquivos (mesma ligação, nota, referência, data e valor) entram uma vez. Matrizes matrícula × mês (`m * nR + r`): `valRef` (referência Mês/Ano), `valPag` (Data Pagto.), `valCred` (Data Crédito; sem ela, Data Pagto.) e `fat` (conta do mês pelo vencimento). `porMes`, `porRef` e `porCred` (linhas da base inteira por mês) e `stats.fat.porMes` servem para achar meses sem dados.
- **Faturamento no worker:** `tabela()` escolhe os nomes de coluna pelo cabeçalho (com `Rubrica` é faturamento). Nos .xlsx do faturamento, `preFiltro()` lê só a célula da ligação e pula as linhas de matrículas sem termo antes de ler o resto (o faturamento inteiro tem dezenas de milhões de linhas).
- **Termos como eventos (v2.0):** cada matrícula guarda a lista dos seus termos; termos da mesma matrícula no mesmo mês viram um evento só (`eQtd`). Arrays por evento: `eMat, eAm, eData, eSeq` (ordem na matrícula), `eNum`, `eProx` (mês do termo seguinte ou -1), `eTipo, eCid, eValor, extra[k]`.
- **IndexedDB** `painel-termos-2026`, chaves `atual` (base) e `pasta`. Base lida por versão anterior (sem `eMat`) pede nova leitura. Parâmetros em `localStorage` (`painel-termos-param-v2`).
- **`derive()`** monta as unidades (`D.U`: eventos analisados) e classifica: `nA/pA/sA` meses, meses pagos e valor antes; `nD/pD/sD` depois; `cA` (0 Adimplente, 1 Flutuante, 2 Inadimplente, 3 Nunca pagou, -1 sem histórico); `cD` (0 a 3, 4 acompanhamento insuficiente); `res` (0 permaneceu, 1 piorou, 2 melhorou, 3 insuficiente, -1 sem histórico); `mA/mD` médias mensais; `fimD` último mês do período depois; `corte` (encurtado por termo seguinte); `semDados` por mês.
- **`agrega(k, chave, nK)`** soma por grupo (total, mês do termo, classe anterior, município, tipo, zona, utilização).
- **Filtros (`defFiltros`)**: principais (Mês do termo, Município, Tipo de irregularidade, Antes, Depois, Resultado, Meses acompanhados) e "Mais filtros" (Zona, Utilização, Situação do termo, Faturado, Isento, Frente, Equipe, Categoria, Origem). Todos são listas de seleção e valem para cards, matriz, gráficos e tabela.

## Regras de negócio (v2.0, definidas com o usuário)

- **Unidade de análise:** padrão é a matrícula única com o primeiro termo do ano; parâmetro permite cada termo (evento). Em qualquer caso, o período depois termina no mês anterior ao termo seguinte da mesma matrícula (não mistura dois termos).
- **Mês pago:** soma do Valor Pago no mês > 0, pela **Data Crédito** (v2.1, definido pelo usuário). Parâmetros permitem a Data Pagto. ou a referência da fatura (Mês/Ano).
- **Meses observados (v2.1):** com faturamento, só os meses em que a matrícula tinha conta a pagar (conta > 0 com vencimento no mês). Mês sem conta não conta como "não pagou" nem como pago; o valor arrecadado nele entra só na média mensal. Conta e pagamento são comparados pelo mês (não há número de conta para ligar o pagamento à conta). Parâmetro "Meses observados" permite usar todos os meses. Matrícula sem nenhuma conta antes do termo fica "Sem histórico antes".
- **Mês sem dados:** mês sem nenhum pagamento de nenhum cliente na base inteira (ou, com faturamento, sem nenhuma conta); fica fora das janelas e aparece no aviso.
- **Média mensal:** valor arrecadado no período dividido pelos meses do período com dados (inclui meses sem conta).
- **Antes:** os 12 meses anteriores ao mês do termo (o mês do termo não entra). 10 a 12 pagos: Adimplente; 6 a 9: Flutuante; 1 a 5: Inadimplente; 0: Nunca pagou. Com meses sem dados, a mesma regra em proporção (≥ 10/12, ≥ 6/12).
- **Depois:** do mês seguinte ao termo até o último mês completo, no máximo 12 meses. % de meses pagos = pagos / observados. Adimplente ≥ 83,3%; Flutuante ≥ 50%; Inadimplente > 0%; Nunca pagou 0%. Menos de 3 meses observados: **Acompanhamento insuficiente**, fora de melhora, piora e permanência. Limiares e mínimo são parâmetros.
- **Resultado:** ordem Adimplente → Flutuante → Inadimplente → Nunca pagou; ir para trás é melhora, para frente é piora.
- **Comparáveis:** com classificação antes e depois (não insuficiente). São o denominador de melhora, piora e permanência.
- **Recuperados:** Inadimplente ou Nunca pagou antes que passaram a Adimplente ou Flutuante, sobre os comparáveis desses dois grupos. "Pagaram ao menos 1 mês" é mostrado separado (registro de pagamento, não recuperação).
- **Último mês completo:** padrão é o mês mais recente da base, menos o mês em andamento.
- **Sem faturamento** (base ausente ou parâmetro desligado): todo mês é observado, e mês sem pagamento conta como não pago mesmo sem conta a pagar; o painel mostra esse aviso.

## Layout (v2.0)

- Barra de filtros com os 7 principais e "Mais filtros"; botão "Limpar filtros".
- Resumo: 7 cards (total, acompanhamento suficiente, melhoraram, pioraram, permaneceram, insuficiente, % recuperados), com denominadores; card de regras e limitação, com parâmetros; **matriz antes × depois** (quantidade ou % da linha; verde melhora, amarelo permanece, vermelho piora, cinza insuficiente; clique abre a tabela de clientes); gráfico A (perfil antes e depois, quantidade ou %); gráfico C (recuperação); gráfico B (resultado por mês de aplicação, com meses acompanhados); gráfico D (média mensal por cliente antes e depois, por classe anterior); tabela por Município / Tipo / Zona / Utilização; % que pagou mês a mês em relação ao termo; rodapé.
- Clientes: uma linha por unidade, com classificações, meses observados e pagos, meses sem conta, % pago depois, resultado, valores e o valor pago em cada mês (laranja-claro sem pagamento no período, contorno no mês do termo, "s/d" sem dados, "s/c" sem conta). Exporta CSV (`;`, BOM, vírgula decimal), com colunas PAGO e CONTA por mês.
- Favicon (v2.1): SVG dentro do próprio `index.html` (lupa branca com gota verde-água num quadrado marinho `#14375e`), simplificado a partir do modelo do usuário (sem a moeda e a segunda gota, que somem em 16 px).
- Cores de resultado em `:root`: `--melhora #1F8A4C`, `--permanece #E0B43C`, `--piora #C62828`, `--insuf #9AA6B2` (pedido do usuário na v2.0).

## Versões anteriores

- v1.1 guardada em `versoes/v1.1/index.html` (abre pelo GitHub Pages em `/versoes/v1.1/`) e no branch `versao-1.1`.
- v2.0 guardada em `versoes/v2.0/index.html` e no branch `versao-2.0`. Guarde a versão anterior assim a cada mudança grande. As tags `v1.0` e `v1.1` existem no repositório local, mas o envio de tags é bloqueado pelo proxy da sessão.

## Pendências

- Conferir com a base real: nomes das rubricas (`VALOR DE AGUA`, `VALOR DE ESGOTO`), formato da Data de Vencimento e da Data Crédito, e as situações de conta listadas no rodapé.
- Ideia não pedida (perguntar antes): comparar com um grupo de clientes sem termo.

## Como validar

```bash
python tests/gerar_base_teste.py base_teste                   # base sintética (requer openpyxl)
python tests/conferir_totais.py base_teste 2026-09                 # padrão: data de crédito, matrícula única, com faturamento
python tests/conferir_totais.py base_teste 2026-09 cred mat 0      # sem faturamento (todos os meses observados)
python tests/conferir_totais.py base_teste 2026-09 pag evento 1    # data do pagamento, cada termo como evento
```

Valores de referência (v2.1) com a base sintética padrão (3000 ligações, semente 7, com Faturamento), análise até set/26:
- padrão (crédito, matrícula, com faturamento): 636 clientes (29 sem histórico), 354 comparáveis, 51 melhoraram, 152 pioraram, 151 permaneceram, 253 insuficientes, 18 recuperados de 82; média mensal R$ 130,93 antes e R$ 103,08 depois; 142 termos com Valor zero fora;
- sem faturamento: 406 comparáveis, 57, 160, 189, 230; 24 de 133;
- data do pagamento, por evento, com faturamento: 664 termos, 365 comparáveis, 52, 156, 157, 268; 19 de 85.

No teste com Playwright, `setInputFiles` descarta arquivos cujo caminho tem acento (as pastas "Arrecadação…"): copie a base para pastas sem acento antes de testar.

Nunca suba arquivos `.xlsx` ou `.csv` para o repositório: o `.gitignore` já bloqueia.
