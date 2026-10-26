# Termos aplicados em 2026 e adimplência

Painel em HTML que cruza a base de termos aplicados em 2026 com a base de arrecadação para mostrar quantos clientes deixaram de ser adimplentes depois do termo e quanto isso representa na arrecadação. O objetivo é avaliar até que ponto vale a pena aplicar um termo, com o risco de perder o cliente pagante.

## Como usar

1. Abra o painel no Chrome ou no Edge.
2. Clique em **📁 Conectar pasta** e escolha a pasta com:
   - `Arrecadação Acumulada/`: os meses fechados (a partir de jan/2025);
   - `Arrecadação Diaria/`: o mês atual;
   - a base de termos, um arquivo com **"termo" no nome** (por exemplo, `termos_todos_os_termos_2026-10-23.xlsx`).
3. Aguarde a leitura. A base fica guardada no navegador e abre sozinha nas próximas vezes. Para atualizar, clique em **↻ Atualizar**.

Arquivos .xlsx e .csv são aceitos (CSV com `;`, `,` ou tabulação, em UTF-8 ou Windows-1252). As subpastas são lidas e os arquivos `~$` ignorados.

## Colunas esperadas

**Arrecadação:** Nº Ligação, Nº Nota, Mês, Ano, Valor Pago, Data Pagto., Utilização (Sub. Categ.), Zona (Localidade), Cliente, Situação Pagto. As demais colunas da exportação são lidas e ignoradas. O cabeçalho pode estar em qualquer uma das 30 primeiras linhas.

**Termos:** só a aba **Termos** é lida (a aba Resumo é ignorada). Usa Matrícula (o mesmo número do Nº Ligação), Data de negociação (mês do termo; sem ela, a Data de aplicação), Tipo de atividade e Valor. Situação do termo, Faturado, Isento, Frente, Equipe, Categoria e Origem viram filtros. Só entram os termos de 2026; se a matrícula tiver mais de um termo no ano, vale o primeiro.

## Classificação

- **Antes do termo:** os 12 meses anteriores ao mês do termo. Adimplente: pagou de 10 a 12 meses; Flutuante: de 6 a 9; Inadimplente: de 1 a 5; Nunca pagou: nenhum.
- **Depois do termo:** do mês seguinte ao termo até o último mês completo (no máximo 12 meses), pelo % de meses pagos. Adimplente: 83,3% ou mais; Flutuante: de 50% a 83,3%; Inadimplente: acima de 0% e abaixo de 50%; Nunca pagou: 0%.
- Com menos de 3 meses depois do termo, o cliente fica em **Acompanhamento insuficiente**, sem contar como melhora, piora ou permanência.
- O mês do termo fica fora dos dois períodos. Se a matrícula recebeu outro termo, o período depois termina antes dele.
- Os parâmetros (unidade de análise, mês do pagamento, último mês, limites) ficam em **Ajustar parâmetros**.

**Limitação:** sem base de faturamento, um mês sem pagamento conta como não pago mesmo que não tenha havido conta a pagar.

## Versão anterior

A versão 1.1 continua disponível em `versoes/v1.1/index.html` e no branch `versao-1.1`.

## Privacidade

Os arquivos são lidos só no seu computador. Nenhum dado é enviado para a internet nem para este repositório. Nunca suba arquivos .xlsx ou .csv da base para cá.
