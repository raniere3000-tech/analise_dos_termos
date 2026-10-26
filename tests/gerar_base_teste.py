"""Gera uma base sintética no formato da exportação de arrecadação e uma base de termos.

Uso: python tests/gerar_base_teste.py PASTA_SAIDA [ligacoes] [semente]

Cria, dentro de PASTA_SAIDA:
  Arrecadação Acumulada/Arrecadacao_AAAA_MM.xlsx   um arquivo por mês de pagamento, de jan/25 a set/26
  Arrecadação Diaria/Arrecadacao_2026_10.csv       mês atual (out/26, até o dia 23), em CSV ; e Windows-1252
  termos_todos_os_termos_2026-10-23.xlsx           abas Resumo (ignorada) e Termos, no layout da exportação real.
                                                   O mês do termo é a Data de negociação; sem ela, a Data de aplicação.
                                                   Termos sem negociação e isentos têm Valor 0 (ficam fora da análise)
  Faturamento/Faturamento_AAAA_MM.xlsx             uma conta por ligação e mês de vencimento (jan/25 a set/26), com as
                                                   rubricas VALOR DE AGUA e VALOR DE ESGOTO e outras que o painel ignora.
                                                   Parte das ligações fica meses sem conta (cortadas)
"""
import csv
import os
import random
import sys
from datetime import date, timedelta

from openpyxl import Workbook

pasta = sys.argv[1] if len(sys.argv) > 1 else 'base_teste'
N = int(sys.argv[2]) if len(sys.argv) > 2 else 3000
rnd = random.Random(int(sys.argv[3]) if len(sys.argv) > 3 else 7)
HOJE = date(2026, 10, 23)

CAB = ['Nº Ligação', 'Nº Nota', 'Mês', 'Ano', 'Valor Pago', 'Vencimento', 'Data Pagto.', 'Data Atualiz.', 'Data Crédito',
       'Val. Autenticado', 'Utilização (Sub. Categ.)', 'Zona (Localidade)', 'Classificação', 'Nº Contrato', 'Cliente',
       'Tipo Fatura', 'Situação Pagto', 'Banco', 'Tipo Pagamento']
ZONAS = ['Miracema', 'Paracambi', 'Seropédica', 'Itaguaí', 'Mangaratiba', 'Rio Claro']
UTIL = ['Residencial', 'Residencial Social', 'Comercial', 'Industrial', 'Pública']
TIPOS = ['Termo de irregularidade', 'Ligação clandestina', 'Violação de hidrômetro']
CAB_TERMO = ['Nº O.S.', 'Matrícula', 'Frente', 'Cidade', 'Zona', 'Equipe', 'Tipo de atividade', 'Status da atividade', 'Serviço adicional identificado',
             'Data de aplicação', 'Data de notificação', 'Data de negociação', 'Nome do solicitante', 'Categoria', 'Situação da ligação', 'Situação do termo',
             'Faturado', 'Dias', 'Faixa', 'Valor', 'Isento', 'Motivo da isenção', 'Origem', 'Arquivo da aplicação', 'Arquivo de status', 'Arquivo de faturamento']
PERFIS = {'adimplente': .96, 'flutuante': .62, 'inadimplente': .25, 'nunca': 0.0}


def am(y, m):
    return y * 12 + m - 1


def ym(a):
    return a // 12, a % 12 + 1


A0, A_FIM = am(2025, 1), am(2026, 10)
lig = []
for i in range(N):
    perfil = rnd.choices(list(PERFIS), [.55, .2, .15, .1])[0]
    lig.append({'num': 100000 + i * 7, 'nome': f'CLIENTE {i:05d}', 'zona': rnd.choice(ZONAS), 'util': rnd.choices(UTIL, [6, 2, 2, .5, .5])[0],
                'perfil': perfil, 'valor': round(rnd.uniform(40, 320), 2), 'contrato': 500000 + i})

def linha_termo(num, t):
    """Linha no layout real. t é o mês do termo (Data de negociação); 10% sem negociação usam a Data de aplicação."""
    y, m = ym(t)
    neg = date(y, m, rnd.randint(1, 28))
    sem_neg = rnd.random() < .1
    apl = neg if sem_neg else neg - timedelta(days=rnd.randint(5, 25))
    isento = rnd.random() < .1
    return [rnd.randint(10 ** 6, 10 ** 7), str(num), rnd.choice(['Frente A', 'Frente B', 'Frente C']), rnd.choice(ZONAS), rnd.choice(ZONAS),
            rnd.choice(['Equipe 01', 'Equipe 02', 'Equipe 03', 'Equipe 04']), rnd.choice(TIPOS), 'Concluída', rnd.choice(['', 'Troca de hidrômetro']),
            apl, apl + timedelta(days=2), None if sem_neg else neg, 'SOLICITANTE', rnd.choice(['Residencial', 'Comercial']), 'Ligada',
            rnd.choice(['Negociado', 'Pago', 'Em aberto']), rnd.choice(['SIM', 'NÃO']), rnd.randint(0, 300), rnd.choice(['0-30', '31-60', '61+']),
            0.0 if isento or (sem_neg and rnd.random() < .8) else round(rnd.uniform(300, 4000), 2), 'SIM' if isento else 'NÃO', 'Decisão administrativa' if isento else '',
            rnd.choice(['Fiscalização', 'Denúncia']), 'apl.pdf', 'status.pdf', 'fat.pdf']


# termos: 1 em cada 4 ligações, aplicados de jan/26 a set/26; o comportamento muda depois do termo
termos = []
for L in lig:
    if rnd.random() < .25:
        t = am(2026, rnd.randint(1, 9))
        L['termo'] = t
        reacao = rnd.choices(['igual', 'piora', 'para', 'melhora'], [.45, .25, .15, .15])[0]
        L['depois'] = {'igual': PERFIS[L['perfil']], 'piora': PERFIS[L['perfil']] * .45, 'para': 0.0,
                       'melhora': min(.97, PERFIS[L['perfil']] + .5)}[reacao]
        termos.append(linha_termo(L['num'], t))
        if rnd.random() < .05:   # segundo termo no ano: o painel usa o primeiro
            termos.append(linha_termo(L['num'], min(am(2026, 10), t + rnd.randint(1, 3))))
for _ in range(15):   # termos de 2025 (fora da análise) e ligações sem arrecadação
    termos.append(linha_termo(rnd.choice(lig)['num'], am(2025, rnd.randint(1, 12))))
for i in range(8):
    termos.append(linha_termo(900000 + i, am(2026, rnd.randint(1, 9))))
rnd.shuffle(termos)

# pagamentos: a fatura de cada mês de referência é paga ou não pelo perfil; 20% pagam no mês seguinte
por_mes = {}
nota = 1
for L in lig:
    for r in range(A0, A_FIM + 1):
        p = PERFIS[L['perfil']] if 'termo' not in L or r < L['termo'] else L['depois']
        nota += 1
        if rnd.random() >= p:
            continue
        pg = r + (1 if rnd.random() < .2 else 0)
        y, m = ym(r); yp, mp = ym(pg)
        dp = date(yp, mp, rnd.randint(1, 28))
        if dp > HOJE:
            continue
        valor = round(L['valor'] * rnd.uniform(.9, 1.1), 2)
        dc = dp + timedelta(days=rnd.randint(0, 2))   # Data Crédito: até 2 dias depois do pagamento (às vezes vira o mês)
        if dc > HOJE:
            dc = dp
        sit = 'PAGO'
        if rnd.random() < .01:
            valor, sit = 0.0, 'CANCELADO'
        por_mes.setdefault(pg, []).append([L['num'], nota, m, y, valor, date(y, m, 10), dp, dp, dc, valor, L['util'], L['zona'],
                                           'NORMAL', L['contrato'], L['nome'], 'CONTA', sit, 'BANCO 001', rnd.choice(['BOLETO', 'PIX', 'DÉBITO'])])
    if rnd.random() < .02:   # pagamento de dívida antiga (referência de 2023), fora do período
        por_mes.setdefault(am(2025, 6), []).append([L['num'], 1, 5, 2023, 99.9, date(2023, 5, 10), date(2025, 6, 3), date(2025, 6, 3), date(2025, 6, 3),
                                                     99.9, L['util'], L['zona'], 'NORMAL', L['contrato'], L['nome'], 'PARCELAMENTO', 'PAGO', 'BANCO 001', 'BOLETO'])

ACUM, DIARIA = 'Arrecadação Acumulada', 'Arrecadação Diaria'
os.makedirs(os.path.join(pasta, ACUM), exist_ok=True)
os.makedirs(os.path.join(pasta, DIARIA), exist_ok=True)
for a in range(A0, A_FIM):
    y, m = ym(a)
    wb = Workbook(); ws = wb.active; ws.title = f'{m:02d}-{y}'
    ws.append([f'Relatório de arrecadação {m:02d}/{y}'])
    ws.append([])
    ws.append(CAB)
    for row in por_mes.get(a, []):
        ws.append(row)
    for c in ('F', 'G', 'H', 'I'):
        for cell in ws[c][3:]:
            cell.number_format = 'dd/mm/yyyy'
    wb.save(os.path.join(pasta, ACUM, f'Arrecadacao_{y}_{m:02d}.xlsx'))

# mês atual em CSV, com algumas linhas de setembro repetidas (o painel descarta pagamentos repetidos)
fmt = lambda v: v.strftime('%d/%m/%Y') if isinstance(v, date) else (f'{v:.2f}'.replace('.', ',') if isinstance(v, float) else v)
with open(os.path.join(pasta, DIARIA, 'Arrecadacao_2026_10.csv'), 'w', newline='', encoding='cp1252') as f:
    w = csv.writer(f, delimiter=';')
    w.writerow(CAB)
    for row in por_mes.get(A_FIM, []) + por_mes.get(A_FIM - 1, [])[:12]:
        w.writerow([fmt(v) for v in row])

# aba Resumo antes da aba Termos, com uma tabela que tem Matrícula: precisa ser ignorada pelo painel
wb = Workbook(); rs = wb.active; rs.title = 'Resumo'
rs.append(['Matrícula', 'Data de negociação', 'Valor'])
rs.append(['999999', date(2026, 5, 1), 1.0])
ws = wb.create_sheet('Termos')
ws.append(CAB_TERMO)
for t in termos:
    ws.append(t)
for c in ('J', 'K', 'L'):
    for cell in ws[c][1:]:
        cell.number_format = 'dd/mm/yyyy'
wb.save(os.path.join(pasta, 'termos_todos_os_termos_2026-10-23.xlsx'))
print(f'{N} ligações, {sum(len(v) for v in por_mes.values())} pagamentos, {len(termos)} linhas de termos em {pasta}')

# faturamento: uma conta por ligação e mês de vencimento (dia 10), de jan/25 a set/26 (meses fechados)
CAB_FAT = ['N. da Ligacao', 'Situacao Lancamento', 'Nome Cliente', 'Categoria', 'Qtd. Economia Residencial', 'Qtd. Economia Comercial',
           'Qtd. Economia Industrial', 'Qtd. Economia Publica', 'Qtd. Economia Outros', 'Nome Logradouro', 'Numero da Porta', 'Desc. Complemento',
           'Nome do Bairro', 'Nome da Localidade', 'Endereco Ligacao', 'Grupo', 'Rota', 'Rubrica', 'Valor Parcela', 'Ocorrencia',
           'Situacao de Repasse', 'Faturamento', 'Situacao Nota Liberada', 'Referencia de Leitura', 'Data de Vencimento', 'Situacao Ligacao', 'Situacao Conta']
os.makedirs(os.path.join(pasta, 'Faturamento'), exist_ok=True)
for L in lig:
    L['esgoto'] = rnd.random() < .5
    # parte dos clientes que não pagam fica sem conta (ligação cortada) a partir de algum mês
    L['corte'] = rnd.randint(A0, A_FIM) if L['perfil'] in ('nunca', 'inadimplente') and rnd.random() < .5 else None
for a in range(A0, A_FIM):
    y, m = ym(a)
    wb = Workbook(); ws = wb.active; ws.title = 'Faturamento'
    ws.append(CAB_FAT)
    for L in lig:
        if L['corte'] is not None and a >= L['corte']:
            continue
        base = [L['num'], 'NORMAL', L['nome'], L['util'], 1, 0, 0, 0, 0, 'RUA A', '10', '', 'CENTRO', L['zona'], 'RUA A, 10', 1, 1]
        fim = [f'{m:02d}/{y}', 'NORMAL', 'NORMAL', 'LIBERADA', f'{ym(a - 1)[1]:02d}/{ym(a - 1)[0]}', date(y, m, 10), 'LIGADA', 'NORMAL']
        agua = round(L['valor'] * (.5 if L['esgoto'] else 1), 2)
        ws.append(base + ['VALOR DE AGUA', agua] + fim)
        if L['esgoto']:
            ws.append(base + ['VALOR DE ESGOTO', round(L['valor'] - agua, 2)] + fim)
        if rnd.random() < .05:
            ws.append(base + ['MULTA POR ATRASO', 5.0] + fim)
    for cell in ws['Y'][1:]:
        cell.number_format = 'dd/mm/yyyy'
    wb.save(os.path.join(pasta, 'Faturamento', f'Faturamento_{y}_{m:02d}.xlsx'))
print('faturamento gerado')
