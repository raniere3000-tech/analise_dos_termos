"""Recalcula em Python, de forma independente, os números do Resumo do painel (parâmetros padrão).

Uso: python tests/conferir_totais.py PASTA [ULTIMO_MES=AAAA-MM] [modo=ref|pag] [inicio=0|1]

Lê todos os .xlsx e .csv da pasta e subpastas. O arquivo com "termo" no nome é a base de termos.
"""
import csv
import os
import re
import sys
import unicodedata
from datetime import date, datetime, timedelta

from openpyxl import load_workbook

pasta = sys.argv[1]
ULT = sys.argv[2] if len(sys.argv) > 2 else None
MODO = sys.argv[3] if len(sys.argv) > 3 else 'ref'
INICIO = int(sys.argv[4]) if len(sys.argv) > 4 else 0
FAIXAS = ['Adimplente', 'Flutuante', 'Inadimplente', 'Nunca pagou']


def norm(s):
    s = unicodedata.normalize('NFD', str(s or '')).encode('ascii', 'ignore').decode().upper()
    return re.sub(r'[^A-Z0-9]+', '_', s).strip('_')


def lig(v):
    if v is None or v == '':
        return ''
    if isinstance(v, (int, float)):
        return str(round(v))
    s = str(v).strip()
    return str(int(float(s.replace(',', '.')))) if re.fullmatch(r'\d+([.,]0+)?', s) else s


def data(v):
    if v in (None, ''):
        return None
    if isinstance(v, datetime) or isinstance(v, date):
        return v.year * 12 + v.month - 1, v.strftime('%d/%m/%Y')
    if isinstance(v, (int, float)):
        d = datetime(1899, 12, 30) + timedelta(days=v)
        return d.year * 12 + d.month - 1, d.strftime('%d/%m/%Y')
    m = re.match(r'(\d{1,2})/(\d{1,2})/(\d{4})', str(v).strip())
    return (int(m[3]) * 12 + int(m[2]) - 1, str(v).strip()[:10]) if m else None


def num(v):
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v or '').strip()
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return 0.0


def tabelas(caminho):
    if caminho.lower().endswith('.csv'):
        raw = open(caminho, 'rb').read()
        try:
            txt = raw.decode('utf-8-sig')
        except UnicodeDecodeError:
            txt = raw.decode('cp1252')
        yield list(csv.reader(txt.splitlines(), delimiter=';'))
    else:
        wb = load_workbook(caminho, read_only=True)
        for ws in wb.worksheets:
            yield list(ws.iter_rows(values_only=True))


def com_cabecalho(linhas):
    for i, l in enumerate(linhas[:30]):
        nomes = [norm(c) for c in l]
        if 'N_LIGACAO' in nomes:
            return nomes, linhas[i + 1:]
    return None, []


arquivos = sorted(os.path.join(r, f) for r, _, fs in os.walk(pasta) for f in fs if f.lower().endswith(('.xlsx', '.csv')))
termos = {}
for a in arquivos:
    if 'termo' not in os.path.basename(a).lower():
        continue
    for t in tabelas(a):
        h, linhas = com_cabecalho(t)
        if not h:
            continue
        il, idt = h.index('N_LIGACAO'), next(i for i, n in enumerate(h) if 'DATA' in n)
        for l in linhas:
            k, d = lig(l[il]), data(l[idt])
            if not k or not d or d[0] // 12 != 2026:
                continue
            ordem = d[1][6:] + d[1][3:5] + d[1][:2]
            if k not in termos or (d[0], ordem) < termos[k][:2]:
                termos[k] = (d[0], ordem)

pag, vistos, min_am, max_am, max_data = {}, set(), 10 ** 9, -1, ''
for a in arquivos:
    if 'termo' in os.path.basename(a).lower():
        continue
    for t in tabelas(a):
        h, linhas = com_cabecalho(t)
        if not h:
            continue
        ix = {n: h.index(n) for n in ('N_LIGACAO', 'N_NOTA', 'MES', 'ANO', 'VALOR_PAGO', 'DATA_PAGTO')}
        for l in linhas:
            k = lig(l[ix['N_LIGACAO']])
            if not k:
                continue
            ref = int(num(l[ix['ANO']])) * 12 + int(num(l[ix['MES']])) - 1
            dp = data(l[ix['DATA_PAGTO']])
            min_am, max_am = min(min_am, dp[0]), max(max_am, dp[0])
            if k not in termos:
                continue
            v = num(l[ix['VALOR_PAGO']])
            chave = (k, str(l[ix['N_NOTA']]).strip(), ref, dp[1], v)
            if chave in vistos:
                continue
            vistos.add(chave)
            if v > 0:
                mes = ref if MODO == 'ref' else dp[0]
                pag[(k, mes)] = pag.get((k, mes), 0) + v

ult = int(ULT[:4]) * 12 + int(ULT[5:7]) - 1 if ULT else max_am - 1


def classe(p, n):
    if n == 0:
        return -1
    if p == 0:
        return 3
    sem = n - p
    ad, fl = int(n * 2 / 12 + .5), int(n * 6 / 12 + .5)   # arredonda 0,5 para cima, como o Math.round do painel
    return 0 if sem <= ad else 1 if sem <= fl else 2


M = [[0] * 4 for _ in range(4)]
tot = dict(termos=0, anal=0, adA=0, deixou=0, parou=0, pagava=0, melhorou=0, piorou=0, mA=0.0, mD=0.0, perda=0.0)
for k, (t, _) in termos.items():
    tot['termos'] += 1
    antes = [m for m in range(t - 12, t) if min_am <= m <= ult]
    depois = [m for m in range(t + INICIO, min(ult, t + INICIO + 11) + 1) if m >= min_am]
    pa = [pag.get((k, m), 0) for m in antes]
    pd = [pag.get((k, m), 0) for m in depois]
    ca, cd = classe(sum(v > 0 for v in pa), len(pa)), classe(sum(v > 0 for v in pd), len(pd))
    if ca < 0 or cd < 0:
        continue
    ma, md = sum(pa) / len(pa), sum(pd) / len(pd)
    tot['anal'] += 1; M[ca][cd] += 1; tot['mA'] += ma; tot['mD'] += md
    tot['adA'] += ca == 0
    if ca == 0 and cd > 0:
        tot['deixou'] += 1; tot['perda'] += ma - md
    if ca < 3:
        tot['pagava'] += 1; tot['parou'] += cd == 3
    tot['melhorou'] += cd < ca; tot['piorou'] += cd > ca

rot = lambda a: ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'][a % 12] + '/' + str(a // 12)[2:]
print(f'Pagamentos de {rot(min_am)} a {rot(max_am)}, análise até {rot(ult)}, modo {MODO}, início {INICIO}')
print(f"Clientes com termo em 2026: {tot['termos']} (analisados {tot['anal']})")
print(f"Adimplentes antes: {tot['adA']} | deixaram de ser adimplentes: {tot['deixou']} | pararam de pagar: {tot['parou']} de {tot['pagava']} | melhoraram: {tot['melhorou']} | pioraram: {tot['piorou']}")
print(f"Arrecadação mensal antes: R$ {tot['mA']:,.2f} | depois: R$ {tot['mD']:,.2f} | perdida (deixaram): R$ {tot['perda']:,.2f}")
print('Matriz (linhas = antes, colunas = depois):')
for a in range(4):
    print(f'  {FAIXAS[a]:<13}', ' '.join(f'{x:>6}' for x in M[a]), f'| {sum(M[a]):>6}')
