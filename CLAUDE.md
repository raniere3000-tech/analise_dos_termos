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
  - **Mês do termo = Data de negociação** (definido pelo usuário). Linha sem Data de negociação usa a Data de aplicação (decisão provisória, a confirmar com o usuário; o rodapé mostra quantas).
  - Tipo de termo = `Tipo de atividade`; valor = `Valor`. Todos os termos entram (nenhum é excluído por status, situação ou isenção).
  - Filtros vindos da base de termos (`EXTRAS`): Situação do termo, Faturado, Isento, Frente, Equipe, Categoria e Origem. Coluna ausente no arquivo não vira filtro. Também aparecem no Consolidado e no CSV.
  - A coluna `Faixa` da base de termos não tem relação com as faixas de adimplência do painel e não é lida.
- Nomes de coluna são comparados sem acento, maiúsculas e pontuação (`norm()`: "Nº Ligação" vira `N_LIGACAO`).
- Número da ligação normalizado em `ligKey()` (sem zeros à esquerda), igual nas duas bases.

## Arquitetura (index.html)

- **Worker inline** (`#worker-src`): lê primeiro a base de termos e depois a arrecadação, guardando **só os pagamentos das matrículas com termo** (a base inteira teria milhões de linhas). Pagamentos repetidos entre arquivos (mesma ligação, nota, referência, data e valor) entram uma vez. Monta duas matrizes matrícula × mês (`m * nR + r`): `valRef` (pela referência Mês/Ano) e `valPag` (pela data do pagamento). Também devolve `porMes` e `porRef` (linhas da base inteira por mês), usados para achar meses sem dados.
- **Termos como eventos (v2.0):** cada matrícula guarda a lista dos seus termos; termos da mesma matrícula no mesmo mês viram um evento só (`eQtd`). Arrays por evento: `eMat, eAm, eData, eSeq` (ordem na matrícula), `eNum`, `eProx` (mês do termo seguinte ou -1), `eTipo, eCid, eValor, extra[k]`.
- **IndexedDB** `painel-termos-2026`, chaves `atual` (base) e `pasta`. Base lida por versão anterior (sem `eMat`) pede nova leitura. Parâmetros em `localStorage` (`painel-termos-param-v2`).
- **`derive()`** monta as unidades (`D.U`: eventos analisados) e classifica: `nA/pA/sA` meses, meses pagos e valor antes; `nD/pD/sD` depois; `cA` (0 Adimplente, 1 Flutuante, 2 Inadimplente, 3 Nunca pagou, -1 sem histórico); `cD` (0 a 3, 4 acompanhamento insuficiente); `res` (0 permaneceu, 1 piorou, 2 melhorou, 3 insuficiente, -1 sem histórico); `mA/mD` médias mensais; `fimD` último mês do período depois; `corte` (encurtado por termo seguinte); `semDados` por mês.
- **`agrega(k, chave, nK)`** soma por grupo (total, mês do termo, classe anterior, município, tipo, zona, utilização).
- **Filtros (`defFiltros`)**: principais (Mês do termo, Município, Tipo de irregularidade, Antes, Depois, Resultado, Meses acompanhados) e "Mais filtros" (Zona, Utilização, Situação do termo, Faturado, Isento, Frente, Equipe, Categoria, Origem). Todos são listas de seleção e valem para cards, matriz, gráficos e tabela.

## Regras de negócio (v2.0, definidas com o usuário)

- **Unidade de análise:** padrão é a matrícula única com o primeiro termo do ano; parâmetro permite cada termo (evento). Em qualquer caso, o período depois termina no mês anterior ao termo seguinte da mesma matrícula (não mistura dois termos).
- **Mês pago:** soma do Valor Pago no mês > 0. Padrão pela **data do pagamento** (pagamento recebido no mês, pedido na v2.0); parâmetro permite a referência da fatura (Mês/Ano). Na base sintética, a data do pagamento deixa bem menos adimplentes (quem paga com atraso tem meses com 0 e meses com 2 pagamentos).
- **Mês sem dados:** mês sem nenhum pagamento de nenhum cliente na base inteira; fica fora das janelas (não vira "não pagou") e aparece no aviso de limitação.
- **Antes:** os 12 meses anteriores ao mês do termo (o mês do termo não entra). 10 a 12 pagos: Adimplente; 6 a 9: Flutuante; 1 a 5: Inadimplente; 0: Nunca pagou. Com meses sem dados, a mesma regra em proporção (≥ 10/12, ≥ 6/12).
- **Depois:** do mês seguinte ao termo até o último mês completo, no máximo 12 meses. % de meses pagos = pagos / observados. Adimplente ≥ 83,3%; Flutuante ≥ 50%; Inadimplente > 0%; Nunca pagou 0%. Menos de 3 meses observados: **Acompanhamento insuficiente**, fora de melhora, piora e permanência. Limiares e mínimo são parâmetros.
- **Resultado:** ordem Adimplente → Flutuante → Inadimplente → Nunca pagou; ir para trás é melhora, para frente é piora.
- **Comparáveis:** com classificação antes e depois (não insuficiente). São o denominador de melhora, piora e permanência.
- **Recuperados:** Inadimplente ou Nunca pagou antes que passaram a Adimplente ou Flutuante, sobre os comparáveis desses dois grupos. "Pagaram ao menos 1 mês" é mostrado separado (registro de pagamento, não recuperação).
- **Último mês completo:** padrão é o mês mais recente da base, menos o mês em andamento.
- **Limitação:** sem base de faturamento, mês sem pagamento conta como não pago mesmo sem conta a pagar. Matrícula sem nenhum pagamento na base conta como Nunca pagou (rodapé avisa).

## Layout (v2.0)

- Barra de filtros com os 7 principais e "Mais filtros"; botão "Limpar filtros".
- Resumo: 7 cards (total, acompanhamento suficiente, melhoraram, pioraram, permaneceram, insuficiente, % recuperados), com denominadores; card de regras e limitação, com parâmetros; **matriz antes × depois** (quantidade ou % da linha; verde melhora, amarelo permanece, vermelho piora, cinza insuficiente; clique abre a tabela de clientes); gráfico A (perfil antes e depois, quantidade ou %); gráfico C (recuperação); gráfico B (resultado por mês de aplicação, com meses acompanhados); gráfico D (média mensal por cliente antes e depois, por classe anterior); tabela por Município / Tipo / Zona / Utilização; % que pagou mês a mês em relação ao termo; rodapé.
- Clientes: uma linha por unidade, com classificações, meses, % pago depois, resultado, valores e o valor pago em cada mês (laranja-claro sem pagamento no período, contorno no mês do termo, cinza "s/d" sem dados). Exporta CSV (`;`, BOM, vírgula decimal).
- Cores de resultado em `:root`: `--melhora #1F8A4C`, `--permanece #E0B43C`, `--piora #C62828`, `--insuf #9AA6B2` (pedido do usuário na v2.0).

## Versões anteriores

- v1.1 guardada em `versoes/v1.1/index.html` (abre pelo GitHub Pages em `/versoes/v1.1/`) e no branch `versao-1.1`. As tags `v1.0` e `v1.1` existem no repositório local, mas o envio de tags é bloqueado pelo proxy da sessão.

## Pendências

- Confirmar com o usuário o que fazer com termos sem Data de negociação (hoje usam a Data de aplicação).
- Base de faturamento (ou débitos): ainda não está na pasta. Quando chegar, usar para marcar os meses com conta a pagar e só contar como "não pagou" os meses com cobrança. Perguntar o layout antes.
- Ideia não pedida (perguntar antes): comparar com um grupo de clientes sem termo.

## Como validar

```bash
python tests/gerar_base_teste.py base_teste                   # base sintética (requer openpyxl)
python tests/conferir_totais.py base_teste 2026-09            # números de referência (data do pagamento, matrícula única)
python tests/conferir_totais.py base_teste 2026-09 ref mat    # referência da fatura
python tests/conferir_totais.py base_teste 2026-09 pag evento # cada termo como evento
```

Valores de referência (v2.0) com a base sintética padrão (3000 ligações, semente 7), análise até set/26:
- data do pagamento, matrícula única: 757 clientes, 469 comparáveis, 74 melhoraram, 211 pioraram, 184 permaneceram, 288 insuficientes, 20 recuperados de 143; média mensal por cliente R$ 124,20 antes e R$ 98,37 depois;
- referência da fatura: 43, 179, 247, 288; 17 de 130; R$ 124,67 e R$ 96,91;
- data do pagamento, por evento: 797 termos, 488 comparáveis, 78, 215, 195, 309; 22 de 150;
- sem o arquivo de fev/26 (mês sem dados): 93, 188, 188, 288; 24 de 152.

No teste com Playwright, `setInputFiles` descarta arquivos cujo caminho tem acento (as pastas "Arrecadação…"): copie a base para pastas sem acento antes de testar.

Nunca suba arquivos `.xlsx` ou `.csv` para o repositório: o `.gitignore` já bloqueia.
