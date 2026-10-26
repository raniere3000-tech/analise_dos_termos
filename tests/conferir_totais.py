"""Recalcula em Python, de forma independente, os números do Resumo do painel (v2, parâmetros padrão).

Uso: python tests/conferir_totais.py PASTA [MES_VIGENTE=AAAA-MM] [modo=cred|pag|ref] [unidade=mat|evento] [fat=1|0] [tolerancia=1]

Regras (as mesmas do painel):
- Termos: arquivo com "termo" no nome, só a aba Termos. Mês do termo = Data de negociação (sem ela, Data de aplicação).
  Termos com Valor 0 ficam fora. Termos da mesma matrícula no mesmo mês são um evento só. Unidade "mat" usa o
  primeiro termo do ano de cada matrícula.
- Mês pago: soma do Valor Pago no mês > 0 (pela Data Crédito, pela Data Pagto. ou pela referência Mês/Ano).
- Mês vigente: termos dele (ou depois) ficam fora; a análise vai até o mês anterior.
- Faturamento (tabelas com a coluna Rubrica): conta do mês = VALOR DE AGUA + VALOR DE ESGOTO pela Data de Vencimento.
  Com fat=1, só os meses com conta > 0 são observados, e a conta do mês r é paga se houve crédito daquela conta
  (Vencimento da arrecadação no mês r) até o mês r + tolerância, e no máximo no mês vigente.
- Mês sem dados: mês sem nenhum pagamento na base inteira (ou sem faturamento, com fat=1); fica fora das janelas.
- Média mensal: valor arrecadado no período / meses do período com dados.
- Antes: 12 meses anteriores ao mês do termo. 10 a 12 pagos (proporcional, ≥ 10/12) Adimplente, ≥ 6/12 Flutuante, > 0 Inadimplente, 0 Nunca pagou.
- Depois: do mês seguinte ao termo até o último mês completo, no máximo 12 meses, e até o mês anterior ao próximo termo
  da mesma matrícula. Menos de 3 meses: acompanhamento insuficiente. ≥ 83,3% Adimplente, ≥ 50% Flutuante, > 0 Inadimplente, 0 Nunca pagou.
"""
import csv
import os
import re
import sys
import unicodedata
from datetime import date, datetime, timedelta

from openpyxl import load_workbook

pasta = sys.argv[1]
VIG = sys.argv[2] if len(sys.argv) > 2 else None
MODO = sys.argv[3] if len(sys.argv) > 3 else 'cred'
UNIDADE = sys.argv[4] if len(sys.argv) > 4 else 'mat'
USA_FAT = (sys.argv[5] if len(sys.argv) > 5 else '1') == '1'
TOL = int(sys.argv[6]) if len(sys.argv) > 6 else 1
MIN, ADIM, FLUT = 3, 83.3, 50.0
FAIXAS = ['Adimplente', 'Flutuante', 'Inadimplente', 'Nunca pagou', 'Acomp. insuf.']


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
    if isinstance(v, (datetime, date)):
        return v.year * 12 + v.month - 1, v.strftime('%Y%m%d')
    if isinstance(v, (int, float)):
        d = datetime(1899, 12, 30) + timedelta(days=v)
        return d.year * 12 + d.month - 1, d.strftime('%Y%m%d')
    m = re.match(r'(\d{1,2})/(\d{1,2})/(\d{4})', str(v).strip())
    return (int(m[3]) * 12 + int(m[2]) - 1, m[3] + m[2].zfill(2) + m[1].zfill(2)) if m else None


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


def tabelas(caminho, so_termos=False):
    if caminho.lower().endswith('.csv'):
        raw = open(caminho, 'rb').read()
        try:
            txt = raw.decode('utf-8-sig')
        except UnicodeDecodeError:
            txt = raw.decode('cp1252')
        yield list(csv.reader(txt.splitlines(), delimiter=';'))
    else:
        wb = load_workbook(caminho, read_only=True)
        abas = [ws for ws in wb.worksheets if so_termos and norm(ws.title) == 'TERMOS'] or wb.worksheets
        for ws in abas:
            yield list(ws.iter_rows(values_only=True))


def com_cabecalho(linhas):
    for i, l in enumerate(linhas[:30]):
        nomes = [norm(c) for c in l]
        if 'N_LIGACAO' in nomes or 'MATRICULA' in nomes or 'N_DA_LIGACAO' in nomes:
            return nomes, linhas[i + 1:]
    return None, []


arquivos = sorted(os.path.join(r, f) for r, _, fs in os.walk(pasta) for f in fs if f.lower().endswith(('.xlsx', '.csv')))
termos = {}   # matrícula -> {mês: data mais antiga}
for a in arquivos:
    if 'termo' not in os.path.basename(a).lower():
        continue
    for t in tabelas(a, so_termos=True):
        h, linhas = com_cabecalho(t)
        if not h:
            continue
        il = h.index('MATRICULA') if 'MATRICULA' in h else h.index('N_LIGACAO')
        ineg, iapl, ival = h.index('DATA_DE_NEGOCIACAO'), h.index('DATA_DE_APLICACAO'), h.index('VALOR')
        for l in linhas:
            k, d = lig(l[il]), data(l[ineg]) or data(l[iapl])
            if not k or not d or d[0] // 12 != 2026 or not num(l[ival]) > 0:
                continue
            ev = termos.setdefault(k, {})
            ev[d[0]] = min(ev.get(d[0], d[1]), d[1])

pag, vistos, min_am, max_am, cobertura, cred_venc = {}, set(), 10 ** 9, -1, set(), {}
fat, vistos_fat, cob_fat = {}, set(), set()
for a in arquivos:
    if 'termo' in os.path.basename(a).lower():
        continue
    for t in tabelas(a):
        h, linhas = com_cabecalho(t)
        if not h:
            continue
        if 'RUBRICA' in h:
            il, ir, iv, ive, irl = (h.index(n) for n in ('N_DA_LIGACAO', 'RUBRICA', 'VALOR_PARCELA', 'DATA_DE_VENCIMENTO', 'REFERENCIA_DE_LEITURA'))
            for l in linhas:
                k = lig(l[il])
                if k not in termos or norm(l[ir]) not in ('VALOR_DE_AGUA', 'VALOR_DE_ESGOTO'):
                    continue
                d, v = data(l[ive]), num(l[iv])
                chave = (k, norm(l[ir]), d[1], str(l[irl]), v)
                if chave in vistos_fat:
                    continue
                vistos_fat.add(chave)
                cob_fat.add(d[0])
                fat[(k, d[0])] = fat.get((k, d[0]), 0) + v
            continue
        ix = {n: h.index(n) for n in ('N_LIGACAO', 'N_NOTA', 'MES', 'ANO', 'VALOR_PAGO', 'DATA_PAGTO', 'DATA_CREDITO', 'VENCIMENTO')}
        for l in linhas:
            k = lig(l[ix['N_LIGACAO']])
            if not k:
                continue
            ref = int(num(l[ix['ANO']])) * 12 + int(num(l[ix['MES']])) - 1
            dp = data(l[ix['DATA_PAGTO']])
            dc = data(l[ix['DATA_CREDITO']]) or dp
            min_am, max_am = min(min_am, dp[0]), max(max_am, dp[0])
            mes = {'pag': dp[0], 'ref': ref, 'cred': dc[0]}[MODO]
            cobertura.add(mes)
            if k not in termos:
                continue
            v = num(l[ix['VALOR_PAGO']])
            chave = (k, str(l[ix['N_NOTA']]).strip(), ref, dp[1], v)
            if chave in vistos:
                continue
            vistos.add(chave)
            if v > 0:
                pag[(k, mes)] = pag.get((k, mes), 0) + v
                ve = data(l[ix['VENCIMENTO']])
                if ve:
                    cred_venc[(k, ve[0])] = min(cred_venc.get((k, ve[0]), dc[0]), dc[0])

vig = int(VIG[:4]) * 12 + int(VIG[5:7]) - 1 if VIG else max_am
ult = vig - 1
USA_FAT = USA_FAT and bool(fat)
ok = lambda m: min_am <= m <= ult and m in cobertura and (not USA_FAT or m in cob_fat)
conta = lambda k, m: not USA_FAT or fat.get((k, m), 0) > 0
# conta paga: crédito daquela conta (mesmo vencimento) até o mês + tolerância e até o mês vigente; sem faturamento, crédito no mês
pago = lambda k, m: (cred_venc.get((k, m), 10 ** 9) <= min(m + TOL, vig)) if USA_FAT else pag.get((k, m), 0) > 0


def antes(p, n):
    if n == 0:
        return -1
    if p == 0:
        return 3
    return 0 if p / n >= 10 / 12 - 1e-9 else 1 if p / n >= .5 - 1e-9 else 2


def depois(p, n):
    if n < MIN:
        return 4
    if p == 0:
        return 3
    x = p * 100 / n
    return 0 if x >= ADIM - 1e-9 else 1 if x >= FLUT - 1e-9 else 2


M = [[0] * 5 for _ in range(4)]
c = dict(total=0, comp=0, melhorou=0, piorou=0, igual=0, insuf=0, semhist=0, naopag=0, recup=0, mA=0.0, mD=0.0)
for k, evs in termos.items():
    meses = sorted(evs)
    for i, t in enumerate(meses):
        if UNIDADE == 'mat' and i > 0:
            break
        if t >= vig:   # termo do mês vigente: fora
            continue
        c['total'] += 1
        fim = min(ult, t + 12)
        if i + 1 < len(meses):
            fim = min(fim, meses[i + 1] - 1)
        ja = [m for m in range(t - 12, t) if ok(m)]
        jd = [m for m in range(t + 1, fim + 1) if ok(m)]
        va, vd = [pag.get((k, m), 0) for m in ja], [pag.get((k, m), 0) for m in jd]   # valores: todos os meses com dados
        oa, od = [pago(k, m) for m in ja if conta(k, m)], [pago(k, m) for m in jd if conta(k, m)]   # meses observados
        ca, cd = antes(sum(oa), len(oa)), depois(sum(od), len(od))
        if ca < 0:
            c['semhist'] += 1
            continue
        M[ca][cd] += 1
        if cd == 4:
            c['insuf'] += 1
            continue
        c['comp'] += 1
        c['melhorou' if cd < ca else 'piorou' if cd > ca else 'igual'] += 1
        c['mA'] += sum(va) / len(va); c['mD'] += sum(vd) / len(vd)
        if ca >= 2:
            c['naopag'] += 1; c['recup'] += cd <= 1

rot = lambda a: ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'][a % 12] + '/' + str(a // 12)[2:]
sem = [rot(m) for m in range(min_am, ult + 1) if not ok(m)]
print(f'Pagamentos de {rot(min_am)} a {rot(max_am)}, mês vigente {rot(vig)}, análise até {rot(ult)}, tolerância {TOL}, modo {MODO}, unidade {UNIDADE}, faturamento {"sim" if USA_FAT else "não"}' + (f', meses sem dados: {", ".join(sem)}' if sem else ''))
print(f"Total: {c['total']} | suficiente: {c['comp']} | melhoraram: {c['melhorou']} | pioraram: {c['piorou']} | permaneceram: {c['igual']} | insuficiente: {c['insuf']} | sem histórico: {c['semhist']}")
print(f"Recuperados: {c['recup']} de {c['naopag']} | média mensal por cliente comparável: antes R$ {c['mA'] / max(1, c['comp']):,.2f}, depois R$ {c['mD'] / max(1, c['comp']):,.2f}")
print('Matriz (linhas = antes, colunas = depois):', ' '.join(f'{f[:6]:>7}' for f in FAIXAS))
for a in range(4):
    print(f'  {FAIXAS[a]:<13}', ' '.join(f'{x:>7}' for x in M[a]), f'| {sum(M[a]):>6}')
