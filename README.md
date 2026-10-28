# Termos aplicados em 2026 e adimplência

Painel em HTML que cruza a base de termos aplicados em 2026 com a base de arrecadação para mostrar quantos clientes deixaram de ser adimplentes depois do termo e quanto isso representa na arrecadação. O objetivo é avaliar até que ponto vale a pena aplicar um termo, com o risco de perder o cliente pagante.

## Como usar

1. Abra o painel no Chrome ou no Edge.
2. Clique em **📁 Conectar pasta** e escolha a pasta com:
   - `Arrecadação Acumulada/`: os meses fechados (a partir de jan/2025);
   - `Arrecadação Diaria/`: o mês atual;
   - `Faturamento/`: as contas de jan/2025 até o último mês fechado;
   - a base de termos, um arquivo com **"termo" no nome** (por exemplo, `termos_todos_os_termos_2026-10-23.xlsx`).
3. Aguarde a leitura. A base fica guardada no navegador e abre sozinha nas próximas vezes.
4. Para atualizar, coloque os arquivos novos nas pastas (arrecadação diária, acumulada, faturamento e a base de termos nova) e clique em **↻ Atualizar**. Um pagamento que esteja na diária e na acumulada conta uma vez só.
5. **Configuração da análise** (faixa logo abaixo das abas): **Base de cálculo** (com faturamento ou só arrecadação), **Mês vigente**, **Data do termo** (negociação ou aplicação) e **Mínimo de meses depois** do termo.
   **Filtros**: os principais ficam à vista; o botão **Todos os filtros** abre uma gaveta com todos, agrupados. Os filtros ativos aparecem como etiquetas; o ✕ tira o filtro.
   No cartão **Perfil antes × depois**, abaixo da matriz, o **fluxo antes → depois** mostra em faixas para onde foram os clientes de cada classificação.
   O cartão **Ticket médio de arrecadação** mostra, mês a mês, o valor arrecadado dividido pelos clientes que pagaram e a quantidade de clientes que pagaram, com o botão **Relativo ao termo** (cada cliente alinhado pelo mês do próprio termo) ou **Calendário** (meses reais; a linha do termo aparece quando há um só mês do termo filtrado).
   A aba **Comparação sem termo** compara os clientes com termo com uma amostra de clientes sem termo (5% da base), com as mesmas regras, para separar o efeito do termo da inadimplência normal.
6. **Mês vigente** (ao lado das abas): abre no mês atual. Os termos desse mês ficam fora, a análise vai até o mês anterior, e os pagamentos do mês vigente só contam como pagamento com atraso de contas dos meses anteriores.

Arquivos .xlsx e .csv são aceitos (CSV com `;`, `,` ou tabulação, em UTF-8 ou Windows-1252). As subpastas são lidas e os arquivos `~$` ignorados.

## Colunas esperadas

**Arrecadação:** Nº Ligação, Nº Nota, Mês, Ano, Valor Pago, Data Pagto., Utilização (Sub. Categ.), Zona (Localidade), Cliente, Situação Pagto. As demais colunas da exportação são lidas e ignoradas. O cabeçalho pode estar em qualquer uma das 30 primeiras linhas.

**Faturamento:** .xlsx ou .csv (qualquer separador e codificação comuns), com N. da Ligacao, Rubrica, Valor Parcela, Data de Vencimento e Nome Cliente. Se o faturamento não puder ser usado, o aviso amarelo do Resumo diz o motivo. Só as rubricas VALOR DE AGUA e VALOR DE ESGOTO entram; a conta do mês é a soma das duas, pelo mês do vencimento.

**Termos:** só a aba **Termos** é lida (a aba Resumo é ignorada). Usa Matrícula (o mesmo número do Nº Ligação), Data de negociação (mês do termo; sem ela, a Data de aplicação), Tipo de atividade e Valor. Situação do termo, Faturado, Isento, Frente, Equipe, Categoria e Origem viram filtros. Só entram os termos de 2026 com Valor maior que zero; se a matrícula tiver mais de um termo no ano, vale o primeiro.

## Período do estudo

Só entram arrecadação e faturamento de **jan/2025 até o mês atual** e termos de **2026** (sem os do mês vigente). Datas fora disso são ignoradas, e o rodapé informa quantas linhas ficaram de fora.

## Classificação

- **Antes do termo:** os 12 meses anteriores ao mês do termo. Adimplente: pagou de 10 a 12 meses; Flutuante: de 6 a 9; Inadimplente: de 1 a 5; Nunca pagou: nenhum.
- **Depois do termo:** do mês seguinte ao termo até o último mês completo (no máximo 12 meses), pelo % de meses pagos. Adimplente: 83,3% ou mais; Flutuante: de 50% a 83,3%; Inadimplente: acima de 0% e abaixo de 50%; Nunca pagou: 0%.
- Com menos de 3 meses depois do termo, o cliente fica em **Acompanhamento insuficiente**, sem contar como melhora, piora ou permanência.
- O mês do termo fica fora dos dois períodos. Matrícula com mais de um termo (reincidente): vale o primeiro, acompanhado até o fim.
- **Ligação cortada** (Situacao Ligacao da última conta no faturamento): depois do termo, os meses sem conta após o corte contam como não pagos.
- Clientes **sem histórico antes** do termo (sem conta nos 12 meses anteriores) aparecem num cartão à parte, só com a situação depois.
- **Mês observado:** a matrícula tinha conta a pagar (vencimento) no mês; mês sem conta não conta como "não pagou".
- **Conta paga:** entrou crédito daquela conta (mesmo vencimento na arrecadação) até o fim do mês seguinte ao vencimento. Pagamento mais atrasado conta como não pago naquele mês.
- Os parâmetros (unidade de análise, mês do pagamento, meses observados, último mês, limites) ficam em **Ajustar parâmetros**.

Sem a pasta Faturamento, todos os meses são observados, e um mês sem pagamento conta como não pago mesmo que não tenha havido conta a pagar.

## Versão anterior

As versões anteriores (1.1 e 2.0 a 2.10) continuam disponíveis em `versoes/` e nos branches `versao-1.1` e `versao-2.0` a `versao-2.10`.

## Privacidade

Os arquivos são lidos só no seu computador. Nenhum dado é enviado para a internet nem para este repositório. Nunca suba arquivos .xlsx ou .csv da base para cá.
