"""Recalcula em Python, de forma independente, os números do Resumo do painel (v2, parâmetros padrão).

Uso: python tests/conferir_totais.py PASTA [MES_VIGENTE=AAAA-MM] [modo=cred|pag|ref] [unidade=mat|evento] [fat=1|0] [tolerancia=1] [data=neg|apl] [minimo=3]

unidade=mat é a contagem Por cliente; unidade=evento, a contagem Por termo (v3.0).

Regras (as mesmas do painel):
- Termos: arquivo com "termo" no nome, só a aba Termos. Mês do termo = Data de negociação (sem ela, Data de aplicação),
  ou o contrário com data=apl.
  Termos com Valor 0 ficam fora. Termos da mesma matrícula no mesmo mês são um evento só. Unidade "mat" usa o
  primeiro termo do ano de cada matrícula; "evento" usa cada termo, com o antes do primeiro termo da matrícula (v3.0).
- Reincidência (v3.0): termos de 2025 e com Valor 0 não entram na análise, mas marcam termo anterior ou posterior.
  Reincidentes (aba própria): matrículas com mais de um termo analisado, nas visões "antes do 1º termo" e "cada termo
  como ponto de partida" (antes do próprio termo).
- Mês pago: soma do Valor Pago no mês > 0 (pela Data Crédito, pela Data Pagto. ou pela referência Mês/Ano).
- Mês vigente: termos dele (ou depois) ficam fora; a análise vai até o mês anterior.
- Faturamento (tabelas com a coluna Rubrica): conta do mês = VALOR DE AGUA + VALOR DE ESGOTO pela Data de Vencimento.
  Com fat=1, só os meses com conta > 0 são observados, e a conta do mês r é paga se houve crédito daquela conta
  (Vencimento da arrecadação no mês r) até o mês r + tolerância, e no máximo no mês vigente.
- Mês sem dados: mês sem nenhum pagamento na base inteira (ou sem faturamento, com fat=1); fica fora das janelas.
- Média mensal: valor arrecadado no período / meses do período com dados.
- Antes: 12 meses anteriores ao mês do termo (na contagem por termo, ao mês do primeiro termo da matrícula). 10 a 12 pagos (proporcional, ≥ 10/12) Adimplente, ≥ 6/12 Flutuante, > 0 Inadimplente, 0 Nunca pagou.
- Depois: do mês seguinte ao termo até o último mês completo, no máximo 12 meses (sem cortar no termo seguinte). Menos de 3 meses: acompanhamento insuficiente. ≥ 83,3% Adimplente, ≥ 50% Flutuante, > 0 Inadimplente, 0 Nunca pagou.
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
DATA = sys.argv[7] if len(sys.argv) > 7 else 'neg'
ADIM, FLUT = 83.3, 50.0
MIN = int(sys.argv[8]) if len(sys.argv) > 8 else 3
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


def hash_lig(k):
    """FNV-1a de 32 bits, igual ao do painel (amostra de comparação e mês de referência)."""
    h = 0x811c9dc5
    for ch in k:
        h ^= ord(ch)
        h = (h * 0x01000193) & 0xFFFFFFFF
    return h


arquivos = sorted(os.path.join(r, f) for r, _, fs in os.walk(pasta) for f in fs if f.lower().endswith(('.xlsx', '.csv')))
termos = {}   # matrícula -> {mês: data mais antiga}
outros = {}   # matrícula -> meses dos termos de 2025 e com Valor 0 (só marcam reincidência)
com_termo = set()   # matrículas com alguma linha válida de 2026 (qualquer data): não entram na amostra de comparação
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
            k = lig(l[il])
            if k and num(l[ival]) > 0 and any(x and x[0] // 12 == 2026 for x in (data(l[ineg]), data(l[iapl]))):
                com_termo.add(k)
            dn, dl = data(l[ineg]), data(l[iapl])
            d = (dn or dl) if DATA == 'neg' else (dl or dn)
            em = lambda a: any(x and x[0] // 12 == a for x in (dn, dl))
            # termo de 2025 (nenhuma data em 2026, alguma em 2025), com Valor 0 (alguma data em 2026) ou válido com a
            # data escolhida em 2025: marca reincidência
            if k and d and ((not em(2026) and em(2025)) or (em(2026) and not num(l[ival]) > 0) or (em(2026) and num(l[ival]) > 0 and d[0] // 12 == 2025)):
                outros.setdefault(k, []).append(d[0])
            if not k or not d or d[0] // 12 != 2026 or not num(l[ival]) > 0:
                continue
            ev = termos.setdefault(k, {})
            ev[d[0]] = min(ev.get(d[0], d[1]), d[1])

pag, vistos, min_am, max_am, cobertura, cred_venc = {}, set(), 10 ** 9, -1, set(), {}
fat, vistos_fat, cob_fat, ult_conta = {}, set(), set(), {}
amostra = set()
for a in arquivos:
    if 'termo' in os.path.basename(a).lower():
        continue
    for t in tabelas(a):
        h, linhas = com_cabecalho(t)
        if not h:
            continue
        if 'RUBRICA' in h:
            il, ir, iv, ive, irl, isl = (h.index(n) for n in ('N_DA_LIGACAO', 'RUBRICA', 'VALOR_PARCELA', 'DATA_DE_VENCIMENTO', 'REFERENCIA_DE_LEITURA', 'SITUACAO_LIGACAO'))
            for l in linhas:
                k = lig(l[il])
                if k and k not in com_termo and hash_lig(k) % 100 < 5:
                    amostra.add(k)
                if (k not in termos and k not in amostra) or norm(l[ir]) not in ('VALOR_DE_AGUA', 'VALOR_DE_ESGOTO'):
                    continue
                d, v = data(l[ive]), num(l[iv])
                chave = (k, norm(l[ir]), d[1], str(l[irl]), v)
                if chave in vistos_fat:
                    continue
                vistos_fat.add(chave)
                cob_fat.add(d[0])
                fat[(k, d[0])] = fat.get((k, d[0]), 0) + v
                if d[0] >= ult_conta.get(k, (-1, ''))[0]:
                    ult_conta[k] = (d[0], str(l[isl] or '').upper())
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
            if k not in com_termo and hash_lig(k) % 100 < 5:
                amostra.add(k)
            if k not in termos and k not in amostra:
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
motivos = [0] * 5   # insuficiente: 1 recente, 2 outro termo, 3 sem dados, 4 sem conta


def classifica(k, t, ta):
    """Classificação de uma unidade (mesmas regras do painel): termo no mês t, período antes definido pelo mês ta.
    Devolve (antes, depois, média antes, média depois, motivo)."""
    fim = min(ult, t + 12)
    ja = [m for m in range(ta - 12, ta) if ok(m)]
    jd = [m for m in range(t + 1, fim + 1) if ok(m)]
    va, vd = [pag.get((k, m), 0) for m in ja], [pag.get((k, m), 0) for m in jd]   # valores: todos os meses com dados
    uc = ult_conta.get(k)
    cortada = USA_FAT and uc is not None and 'CORT' in uc[1]
    # ligação cortada: depois do termo, mês sem conta depois da última conta conta como não pago
    oa = [pago(k, m) for m in ja if conta(k, m)]
    od = [pago(k, m) if conta(k, m) else False for m in jd if conta(k, m) or (cortada and m > uc[0])]
    ca, cd = antes(sum(oa), len(oa)), depois(sum(od), len(od))
    mot = 0
    if cd == 4:
        mot = 1 if fim - t < MIN else 3 if len(jd) < MIN else 4
    return ca, cd, (sum(va) / len(va) if va else 0), (sum(vd) / len(vd) if vd else 0), mot


def soma(unidades):
    c = dict(total=0, comp=0, melhorou=0, piorou=0, igual=0, insuf=0, semhist=0, naopag=0, recup=0, mA=0.0, mD=0.0)
    M = [[0] * 5 for _ in range(4)]
    for k, t, ta in unidades:
        c['total'] += 1
        ca, cd, ma, md, mot = classifica(k, t, ta)
        if cd == 4:
            motivos[mot] += 1 if unidades is un_termo else 0
        if ca < 0:
            c['semhist'] += 1
            c['insuf'] += cd == 4   # insuficiente conta também quem não tem histórico antes (como a matriz)
            continue
        M[ca][cd] += 1
        if cd == 4:
            c['insuf'] += 1
            continue
        c['comp'] += 1
        c['melhorou' if cd < ca else 'piorou' if cd > ca else 'igual'] += 1
        c['mA'] += ma; c['mD'] += md
        if ca >= 2:
            c['naopag'] += 1; c['recup'] += cd <= 1
    return c, M


un_termo = []
for k, evs in termos.items():
    meses = sorted(evs)
    for i, t in enumerate(meses):
        if UNIDADE == 'mat' and i > 0:
            break
        if t >= vig:   # termo do mês vigente: fora
            continue
        un_termo.append((k, t, meses[0]))
c, M = soma(un_termo)
# amostra de comparação: mês de referência sorteado entre os meses das unidades com termo
dist = sorted(t for _, t, _ in un_termo)
un_ctrl = [(k, dist[(hash_lig(k) // 100) % len(dist)], dist[(hash_lig(k) // 100) % len(dist)]) for k in sorted(amostra)]
cc, MC = soma(un_ctrl)

rot = lambda a: ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'][a % 12] + '/' + str(a // 12)[2:]
sem = [rot(m) for m in range(min_am, ult + 1) if not ok(m)]
print(f'Pagamentos de {rot(min_am)} a {rot(max_am)}, mês vigente {rot(vig)}, análise até {rot(ult)}, tolerância {TOL}, modo {MODO}, unidade {UNIDADE}, faturamento {"sim" if USA_FAT else "não"}' + (f', meses sem dados: {", ".join(sem)}' if sem else ''))
print(f"Total: {c['total']} | suficiente: {c['comp']} | melhoraram: {c['melhorou']} | pioraram: {c['piorou']} | permaneceram: {c['igual']} | insuficiente: {c['insuf']} | sem histórico: {c['semhist']}")
print(f"Recuperados: {c['recup']} de {c['naopag']} | média mensal por cliente comparável: antes R$ {c['mA'] / max(1, c['comp']):,.2f}, depois R$ {c['mD'] / max(1, c['comp']):,.2f}")
print(f'Motivos do acompanhamento insuficiente: recente {motivos[1]}, outro termo {motivos[2]}, sem dados {motivos[3]}, sem conta {motivos[4]}')
print(f"Sem termo (amostra): {cc['total']} | comparáveis: {cc['comp']} | melhoraram: {cc['melhorou']} | pioraram: {cc['piorou']} | permaneceram: {cc['igual']} | recuperados: {cc['recup']} de {cc['naopag']}")
print('Matriz (linhas = antes, colunas = depois):', ' '.join(f'{f[:6]:>7}' for f in FAIXAS))
for a in range(4):
    print(f'  {FAIXAS[a]:<13}', ' '.join(f'{x:>7}' for x in M[a]), f'| {sum(M[a]):>6}')


def ticket(unidades):
    """Ticket médio (v2.11): valor arrecadado no mês ÷ clientes que pagaram no mês, nos eixos relativo e calendário.
    No eixo relativo, os meses antes contam a partir do mês que define o período antes (v3.0)."""
    rel = {}   # deslocamento -> [clientes no mês, pagaram, valor]
    for k, t, ta in unidades:
        fim = min(ult, t + 12)
        for o in [*range(-12, 0), *range(1, 13)]:
            m = (ta if o < 0 else t) + o
            if not ok(m) or (o > 0 and m > fim):
                continue
            x = rel.setdefault(o, [0, 0, 0.0])
            v = pag.get((k, m), 0)
            x[0] += 1
            if v > 0:
                x[1] += 1; x[2] += v
    cal = {}
    for k in {k for k, _, _ in unidades}:
        for m in range(min_am, ult + 1):
            if not ok(m):
                continue
            x = cal.setdefault(m, [0, 0, 0.0])
            v = pag.get((k, m), 0)
            x[0] += 1
            if v > 0:
                x[1] += 1; x[2] += v
    per = lambda d, ks: sum(d[o][2] for o in ks if o in d) / max(1, sum(d[o][1] for o in ks if o in d))
    return rel, cal, per(rel, range(-12, 0)), per(rel, range(1, 13))


for nome, un in (('Com termo', un_termo), ('Sem termo', un_ctrl)):
    rel, cal, ta, td = ticket(un)
    print(f'Ticket {nome}: antes R$ {ta:,.2f}, depois R$ {td:,.2f} ({(td / ta - 1) * 100:+.1f}%)')
    print('  relativo:', ' '.join(f'{o:+d}:{rel[o][1]}/{rel[o][0]}={rel[o][2] / rel[o][1]:.2f}' for o in sorted(rel) if rel[o][1] and o in (-12, -1, 1, 3, 6)))
    print('  calendário:', ' '.join(f'{rot(m)}:{cal[m][1]}/{cal[m][0]}={cal[m][2] / cal[m][1]:.2f}' for m in sorted(cal) if cal[m][1] and m in (min_am, min_am + 12, ult)))


# reincidência (v3.0): termo anterior e posterior de cada unidade
def marcas(k, t):
    meses = sorted(termos[k])
    o = outros.get(k, [])
    return (meses[0] < t or any(a < t for a in o)), (meses[-1] > t or any(a > t for a in o))


ant = sum(marcas(k, t)[0] for k, t, _ in un_termo)
post = sum(marcas(k, t)[1] for k, t, _ in un_termo)
print(f'Termo anterior: {ant} | termo posterior: {post} (de {len(un_termo)}); termos que só marcam reincidência: {sum(len(v) for k, v in outros.items() if k in termos)}')

# aba Reincidentes: matrículas com mais de um termo analisado; visão A (antes do 1º termo) e B (antes do próprio termo)
analisados = {k: [t for t in sorted(ev) if t < vig] for k, ev in termos.items()}
reinc = {k: ts for k, ts in analisados.items() if len(ts) > 1}
un_a = [(k, t, ts[0]) for k, ts in sorted(reinc.items()) for t in ts]
un_b = [(k, t, t) for k, ts in sorted(reinc.items()) for t in ts]
for nome, un in (('A antes do 1º termo', un_a), ('B ponto de partida', un_b)):
    r, _ = soma(un)
    ordem = {}
    for k, t, ta in un:
        ca, cd, *_ = classifica(k, t, ta)
        g = ordem.setdefault(min(2, reinc[k].index(t)), [0, 0, 0, 0])   # termos, melhorou, permaneceu, piorou
        g[0] += 1
        if ca >= 0 and cd != 4:
            g[1 if cd < ca else 2 if cd == ca else 3] += 1
    print(f"Reincidentes {nome}: {len(reinc)} matrículas, {r['total']} termos | comparáveis {r['comp']} | melhoraram {r['melhorou']} | pioraram {r['piorou']} | permaneceram {r['igual']} | insuficiente {r['insuf']} | sem histórico {r['semhist']} | recuperados {r['recup']} de {r['naopag']}")
    print('  por ordem (termos, melhorou, permaneceu, piorou):', ' '.join(f'{i + 1}º:{ordem[i]}' for i in sorted(ordem)))
traj, fora = {}, 0
for k, ts in sorted(reinc.items()):
    ca, cd1, *_ = classifica(k, ts[0], ts[0])
    _, cd2, *_ = classifica(k, ts[1], ts[0])
    if cd1 == 4 or cd2 == 4:
        fora += 1
        continue
    ch = (4 if ca < 0 else ca, cd1, cd2)
    traj[ch] = traj.get(ch, 0) + 1
print(f'Trajetória: {sum(traj.values())} clientes, {fora} fora; mais comuns:', ', '.join(f'{FAIXAS[a][:5] if a < 4 else "SemH"}>{FAIXAS[b][:5]}>{FAIXAS[c][:5]}:{v}' for (a, b, c), v in sorted(traj.items(), key=lambda x: -x[1])[:4]))


# Resumo executivo (v3.1): destino dos adimplentes antes do termo e piora líquida (com termo − sem termo)
def adim(Mx):
    linha = Mx[0][:4]
    n = sum(linha)
    return n, linha, (n - linha[0]) / n if n else None, (linha[2] + linha[3]) / n if n else None


nT, lT, pT, gT = adim(M)
nCa, lC, pC, gC = adim(MC)
print(f'Executivo: adimplentes com termo {nT} {lT} piora {pT:.4f} grave {gT:.4f} | sem termo {nCa} {lC} piora {pC:.4f} grave {gC:.4f} | piora líquida {(pT - pC) * 100:+.2f} p.p.')
