# Termos aplicados em 2026 e adimplência

Painel em HTML que cruza a base de termos aplicados em 2026 com a base de arrecadação para mostrar quantos clientes deixaram de ser adimplentes depois do termo e quanto isso representa na arrecadação. O objetivo é avaliar até que ponto vale a pena aplicar um termo, com o risco de perder o cliente pagante.

## Como usar

1. Abra o painel no Chrome ou no Edge.
2. Clique em **📁 Conectar pasta** e escolha a pasta com:
   - `arrecadação acumulada/`: os meses fechados (a partir de jan/2025);
   - `arrecadação diária/`: o mês atual;
   - a base de termos, um arquivo com **"termo" no nome** (por exemplo, `Termos aplicados 2026.xlsx`).
3. Aguarde a leitura. A base fica guardada no navegador e abre sozinha nas próximas vezes. Para atualizar, clique em **↻ Atualizar**.

Arquivos .xlsx e .csv são aceitos (CSV com `;`, `,` ou tabulação, em UTF-8 ou Windows-1252). As subpastas são lidas e os arquivos `~$` ignorados.

## Colunas esperadas

**Arrecadação:** Nº Ligação, Nº Nota, Mês, Ano, Valor Pago, Data Pagto., Utilização (Sub. Categ.), Zona (Localidade), Cliente, Situação Pagto. As demais colunas da exportação são lidas e ignoradas. O cabeçalho pode estar em qualquer uma das 30 primeiras linhas.

**Termos:** Nº Ligação (ou Matrícula) e a data do termo (Data do Termo, Data de Aplicação ou outra coluna com "Data" no nome). Tipo de termo e Valor do termo são opcionais. Só entram os termos de 2026; se a ligação tiver mais de um termo no ano, vale o primeiro.

## Classificação

- **Faixa no mês do termo:** os 12 meses anteriores. Até 2 meses sem pagar: Adimplente; até 6: Flutuante; pagou de 1 a 5: Inadimplente; nenhum: Nunca pagou.
- **Faixa depois:** do mês do termo até o último mês da análise (no máximo 12 meses), com limites ajustáveis para cada quantidade de meses em **Ajustar parâmetros**.

## Privacidade

Os arquivos são lidos só no seu computador. Nenhum dado é enviado para a internet nem para este repositório. Nunca suba arquivos .xlsx ou .csv da base para cá.
