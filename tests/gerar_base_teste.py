"""Gera uma base sintética no formato da exportação de arrecadação e uma base de termos.

Uso: python tests/gerar_base_teste.py PASTA_SAIDA [ligacoes] [semente]

Cria, dentro de PASTA_SAIDA:
  arrecadacao acumulada/Arrecadacao_AAAA_MM.xlsx   um arquivo por mês de pagamento, de jan/25 a set/26
  arrecadacao diaria/Arrecadacao_2026_10.csv       mês atual (out/26, até o dia 23), em CSV ; e Windows-1252
  Termos aplicados 2026.xlsx                       termos aplicados (com alguns de 2025 e repetidos)
"""
import csv
import os
import random
import sys
from datetime import date

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

# termos: 1 em cada 4 ligações, aplicados de jan/26 a set/26; o comportamento muda depois do termo
termos = []
for L in lig:
    if rnd.random() < .25:
        t = am(2026, rnd.randint(1, 9))
        L['termo'] = t
        reacao = rnd.choices(['igual', 'piora', 'para', 'melhora'], [.45, .25, .15, .15])[0]
        L['depois'] = {'igual': PERFIS[L['perfil']], 'piora': PERFIS[L['perfil']] * .45, 'para': 0.0,
                       'melhora': min(.97, PERFIS[L['perfil']] + .5)}[reacao]
        y, m = ym(t)
        termos.append([L['num'], date(y, m, rnd.randint(1, 28)), rnd.choice(TIPOS), round(rnd.uniform(300, 4000), 2)])
        if rnd.random() < .05:   # segundo termo no ano: o painel usa o primeiro
            t2 = min(am(2026, 10), t + rnd.randint(1, 3)); y2, m2 = ym(t2)
            termos.append([L['num'], date(y2, m2, rnd.randint(1, 28)), rnd.choice(TIPOS), round(rnd.uniform(300, 4000), 2)])
for _ in range(15):   # termos de 2025 (fora da análise) e ligações sem arrecadação
    termos.append([rnd.choice(lig)['num'], date(2025, rnd.randint(1, 12), 10), rnd.choice(TIPOS), 500.0])
for i in range(8):
    termos.append([900000 + i, date(2026, rnd.randint(1, 9), 5), rnd.choice(TIPOS), 800.0])
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
        sit = 'PAGO'
        if rnd.random() < .01:
            valor, sit = 0.0, 'CANCELADO'
        por_mes.setdefault(pg, []).append([L['num'], nota, m, y, valor, date(y, m, 10), dp, dp, dp, valor, L['util'], L['zona'],
                                           'NORMAL', L['contrato'], L['nome'], 'CONTA', sit, 'BANCO 001', rnd.choice(['BOLETO', 'PIX', 'DÉBITO'])])
    if rnd.random() < .02:   # pagamento de dívida antiga (referência de 2023), fora do período
        por_mes.setdefault(am(2025, 6), []).append([L['num'], 1, 5, 2023, 99.9, date(2023, 5, 10), date(2025, 6, 3), date(2025, 6, 3), date(2025, 6, 3),
                                                     99.9, L['util'], L['zona'], 'NORMAL', L['contrato'], L['nome'], 'PARCELAMENTO', 'PAGO', 'BANCO 001', 'BOLETO'])

os.makedirs(os.path.join(pasta, 'arrecadacao acumulada'), exist_ok=True)
os.makedirs(os.path.join(pasta, 'arrecadacao diaria'), exist_ok=True)
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
    wb.save(os.path.join(pasta, 'arrecadacao acumulada', f'Arrecadacao_{y}_{m:02d}.xlsx'))

# mês atual em CSV, com algumas linhas de setembro repetidas (o painel descarta pagamentos repetidos)
fmt = lambda v: v.strftime('%d/%m/%Y') if isinstance(v, date) else (f'{v:.2f}'.replace('.', ',') if isinstance(v, float) else v)
with open(os.path.join(pasta, 'arrecadacao diaria', 'Arrecadacao_2026_10.csv'), 'w', newline='', encoding='cp1252') as f:
    w = csv.writer(f, delimiter=';')
    w.writerow(CAB)
    for row in por_mes.get(A_FIM, []) + por_mes.get(A_FIM - 1, [])[:12]:
        w.writerow([fmt(v) for v in row])

wb = Workbook(); ws = wb.active; ws.title = 'Termos'
ws.append(['Nº Ligação', 'Data do Termo', 'Tipo de Termo', 'Valor do Termo'])
for t in termos:
    ws.append(t)
for cell in ws['B'][1:]:
    cell.number_format = 'dd/mm/yyyy'
wb.save(os.path.join(pasta, 'Termos aplicados 2026.xlsx'))
print(f'{N} ligações, {sum(len(v) for v in por_mes.values())} pagamentos, {len(termos)} linhas de termos em {pasta}')
