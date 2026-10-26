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
- **Termos (v1.1):** arquivo com "termo" no nome (hoje `termos_todos_os_termos_AAAA-MM-DD.xlsx`). Lê **só a aba Termos**; a aba Resumo é ignorada (`Termos.abas()`). Colunas da exportação: `Nº O.S., Matrícula, Frente, Cidade, Zona, Equipe, Tipo de atividade, Status da atividade, Serviço adicional identificado, Data de aplicação, Data de notificação, Data de negociação, Nome do solicitante, Categoria, Situação da ligação, Situação do termo, Faturado, Dias, Faixa, Valor, Isento, Motivo da isenção, Origem, Arquivo da aplicação, Arquivo de status, Arquivo de faturamento`.
  - `Matrícula` é o mesmo número que o `Nº Ligação` da arrecadação (confirmado pelo usuário).
  - **Mês do termo = Data de negociação** (definido pelo usuário). Linha sem Data de negociação usa a Data de aplicação (decisão provisória, a confirmar com o usuário; o rodapé mostra quantas).
  - Tipo de termo = `Tipo de atividade`; valor = `Valor`. Todos os termos entram (nenhum é excluído por status, situação ou isenção).
  - Filtros vindos da base de termos (`EXTRAS`): Situação do termo, Faturado, Isento, Frente, Equipe, Categoria e Origem. Coluna ausente no arquivo não vira filtro. Também aparecem no Consolidado e no CSV.
  - A coluna `Faixa` da base de termos não tem relação com as faixas de adimplência do painel e não é lida.
- Nomes de coluna são comparados sem acento, maiúsculas e pontuação (`norm()`: "Nº Ligação" vira `N_LIGACAO`).
- Número da ligação normalizado em `ligKey()` (sem zeros à esquerda), igual nas duas bases.

## Arquitetura (index.html)

- **Worker inline** (`#worker-src`): lê primeiro a base de termos e depois a arrecadação, guardando **só os pagamentos dos clientes com termo** (a base inteira teria milhões de linhas). Pagamentos repetidos entre arquivos (mesma ligação, nota, referência, data e valor) entram uma vez. Monta duas matrizes cliente × mês (`m * nR + r`): `valRef` (pela referência Mês/Ano) e `valPag` (pela data do pagamento). O eixo de meses é o período das datas de pagamento (`am0`, `nR`); referências fora dele (dívidas antigas) só contam pela data do pagamento.
- **IndexedDB** `painel-termos-2026`, chaves `atual` (base) e `pasta` (handle da pasta). Parâmetros em `localStorage` (`painel-termos-param`).
- **`derive()`** classifica cada cliente com os parâmetros `P`: `nA/pA` meses e meses pagos antes, `nD/pD` depois, `cA/cD` faixas (0 Adimplente, 1 Flutuante, 2 Inadimplente, 3 Nunca pagou, -1 sem período), `mA/mD` arrecadação média mensal, `mov` (0 manteve, 1 piorou, 2 melhorou, -1 fora da análise).
- **`agrega(k, chave, nK)`** soma por grupo: usado no total, por mês do termo, tipo, zona e utilização.

## Regras de negócio (definidas com o usuário)

- **Mês pago:** soma do Valor Pago no mês > 0. Padrão pela referência da fatura (Mês/Ano); parâmetro permite usar a data do pagamento.
- **Faixa no mês do termo:** 12 meses anteriores ao mês do termo. Até 2 meses sem pagar: Adimplente; até 6: Flutuante; pagou de 1 a 5: Inadimplente; nenhum: Nunca pagou.
- **Faixa depois:** do mês do termo (parâmetro: ou do seguinte) até o último mês da análise, no máximo 12 meses. Como não há 12 meses depois, os limites por quantidade de meses são parâmetros (tabela em "Ajustar parâmetros"); o padrão é proporcional (`Math.round(n*2/12)` e `Math.round(n*6/12)`). A coluna 12 vale também para o antes.
- **Último mês da análise:** padrão é o mês mais recente da base, menos o mês atual quando ele está incompleto.
- **Deixou de ser adimplente:** Adimplente antes e outra faixa depois. **Parou de pagar:** pagava antes (não era Nunca pagou) e não pagou nenhum mês depois.
- Cliente com termo sem nenhum pagamento na base conta como Nunca pagou (o rodapé avisa quantos são).

## Layout

- Resumo: 8 indicadores; regra de classificação com parâmetros; matriz faixa antes × depois (números abrem o Consolidado); gráfico por mês de aplicação (melhorou/manteve/piorou) com tabela alinhada; gráfico de % que pagou em cada mês relativo ao termo (−12 a +11), com seletor Todos / Adimplentes antes / Deixaram de ser adimplentes; tabelas por tipo de termo, zona e utilização (ordenáveis); rodapé com a origem da base.
- Consolidado: um cliente por linha, faixas, movimento, arrecadação média e o valor pago em cada mês (laranja-claro = mês sem pagamento dentro do período, contorno = mês do termo). Exporta CSV (`;`, BOM, vírgula decimal).

## Pendências

- Confirmar com o usuário o que fazer com termos sem Data de negociação (hoje usam a Data de aplicação).
- Ideia não pedida (perguntar antes): comparar com um grupo de clientes sem termo, para separar o efeito do termo da inadimplência normal da base.

## Como validar

```bash
python tests/gerar_base_teste.py base_teste              # base sintética (requer openpyxl)
python tests/conferir_totais.py base_teste 2026-09       # números de referência (modo ref, início no mês do termo)
python tests/conferir_totais.py base_teste 2026-09 pag 1 # data do pagamento, período depois a partir do mês seguinte
```

Valores de referência com a base sintética padrão (3000 ligações, semente 7, layout real de termos desde a v1.1), análise até set/26: 757 clientes com termo, 415 adimplentes antes, 159 deixaram de ser adimplentes, 177 pararam de pagar (de 663), 87 melhoraram, arrecadação mensal R$ 90.997 antes e R$ 69.494 depois, R$ 20.473 perdidos de quem deixou de ser adimplente; 86 termos sem Data de negociação. Com data do pagamento e início no mês seguinte: 211, 110, 160, 136, R$ 80.213, R$ 62.785, R$ 10.417. Com o filtro Isento = NÃO: 680 clientes.

No teste com Playwright, `setInputFiles` descarta arquivos cujo caminho tem acento (as pastas "Arrecadação…"): copie a base para pastas sem acento antes de testar.

Nunca suba arquivos `.xlsx` ou `.csv` para o repositório: o `.gitignore` já bloqueia.
