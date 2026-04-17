"""
Extrator de Razão Analítico gerado pelo RDprint (PRAXIO)
Converte o PDF em um DataFrame Pandas com todas as linhas de lançamento.

Dependências (100% Python, sem executáveis externos):
    pip install pdfplumber pandas

Uso:
    df = extrair_razao_analitico("Gerado_por_RDprint__1_.PDF")
    print(df.head())
    df.to_csv("razao_analitico.csv", index=False)
    df.to_excel("razao_analitico.xlsx", index=False)
"""

import re
import pdfplumber
import pandas as pd


# ---------------------------------------------------------------------------
# Limites de coluna (coordenada x) baseados no layout do RDprint/PRAXIO
# Obtidos inspecionando a posição das palavras do cabeçalho no PDF.
# ---------------------------------------------------------------------------
COL_FIL_X       = (88,  105)   # "Fil"       x0≈91
COL_DATA_X      = (105, 152)   # "Data"      x0≈108
COL_LOT_X       = (152, 168)   # "Lot"       x0≈155
COL_DOC_X       = (168, 216)   # "Documento" x0≈172
COL_ITEM_X      = (216, 237)   # "Item"      x0≈218
COL_CPART_X     = (237, 263)   # "CPart"     x0≈239
COL_HIST_X      = (263, 475)   # "Historico" x0≈265
COL_DEBITO_X    = (475, 543)   # "Debito"    x0≈481
COL_CREDITO_X   = (543, 600)   # "Credito"   x0≈549
COL_SALDO_X     = (600, 680)   # "Saldo Atual" x0≈604

# Faixa Y de linhas de dados (ignora cabeçalho e rodapé da página)
Y_DADOS_MIN = 115
Y_DADOS_MAX = 820

# Regex para valor numérico BR: 1.234,56 ou 234,56
VALOR_RE = re.compile(r'^-?[\d]+(?:\.\d{3})*,\d{2}$')


def _br_para_float(texto: str):
    """Converte '1.234,56' → 1234.56. Retorna None se inválido."""
    if not texto:
        return None
    try:
        return float(texto.strip().replace('.', '').replace(',', '.'))
    except ValueError:
        return None


def _agrupar_por_linha(words: list, tolerancia_y: float = 3.0) -> list:
    """
    Agrupa palavras em linhas lógicas pela proximidade vertical (top).
    Retorna lista de (top_medio, [words]).
    """
    if not words:
        return []

    words_sorted = sorted(words, key=lambda w: (w['top'], w['x0']))
    grupos = []
    grupo_atual = [words_sorted[0]]
    top_atual = words_sorted[0]['top']

    for w in words_sorted[1:]:
        if abs(w['top'] - top_atual) <= tolerancia_y:
            grupo_atual.append(w)
        else:
            grupos.append((top_atual, grupo_atual))
            grupo_atual = [w]
            top_atual = w['top']

    if grupo_atual:
        grupos.append((top_atual, grupo_atual))

    return grupos


def _col(ws: list, x_min: float, x_max: float) -> str:
    """Concatena texto das palavras dentro da faixa X, ordenadas por x0."""
    return ' '.join(
        w['text'] for w in sorted(ws, key=lambda w: w['x0'])
        if x_min <= w['x0'] < x_max
    )


def _e_linha_dados(ws: list) -> bool:
    """Verifica se linha começa com Fil (3 dígitos na faixa correta)."""
    faixa = [w for w in ws if COL_FIL_X[0] <= w['x0'] < COL_FIL_X[1]]
    return bool(faixa) and re.fullmatch(r'\d{3}', faixa[0]['text'])


def _e_linha_conta(ws: list) -> bool:
    """Verifica se a linha é cabeçalho de CONTA."""
    return any(
        w['text'] == 'CONTA' for w in ws
        if COL_FIL_X[0] <= w['x0'] < 130
    )


def _e_continuacao_historico(ws: list) -> bool:
    """
    Linha de continuação: tem palavras na faixa do histórico,
    mas não tem Fil nem Documento.
    """
    tem_hist = any(COL_HIST_X[0] <= w['x0'] < COL_HIST_X[1] for w in ws)
    tem_fil  = any(COL_FIL_X[0]  <= w['x0'] < COL_FIL_X[1]  for w in ws)
    tem_doc  = any(COL_DOC_X[0]  <= w['x0'] < COL_DOC_X[1]  for w in ws)
    return tem_hist and not tem_fil and not tem_doc


def extrair_razao_analitico(caminho_pdf: str) -> pd.DataFrame:
    """
    Extrai os lançamentos do Razão Analítico (RDprint/PRAXIO) para DataFrame.

    Parâmetros
    ----------
    caminho_pdf : str
        Caminho para o arquivo PDF.

    Retorna
    -------
    pd.DataFrame com colunas:
        conta_codigo, conta_nome, fil, data, lot, documento, item,
        cpart, historico, debito, credito, saldo_atual
    """
    registros = []
    conta_codigo = None
    conta_nome   = None
    registro_atual = None

    with pdfplumber.open(caminho_pdf) as pdf:
        total = len(pdf.pages)
        print(f"Processando {total} páginas...", flush=True)

        for num_pagina, page in enumerate(pdf.pages, 1):
            if num_pagina % 20 == 0:
                print(f"  Página {num_pagina}/{total}", flush=True)

            # Extrai palavras com coordenadas de posição
            words = page.extract_words(
                x_tolerance=3,
                y_tolerance=3,
                keep_blank_chars=False,
                use_text_flow=False,
            )

            # Filtra apenas a área de dados (exclui cabeçalho/rodapé da página)
            words = [w for w in words if Y_DADOS_MIN <= w['top'] <= Y_DADOS_MAX]

            # Agrupa por linha lógica
            linhas = _agrupar_por_linha(words)

            for top, ws in linhas:

                # ── Cabeçalho de CONTA ────────────────────────────────────
                if _e_linha_conta(ws):
                    if registro_atual:
                        registros.append(registro_atual)
                        registro_atual = None

                    # Filtra palavras até a coluna de Débito (evita "Saldo Anterior")
                    ws_desc = [w for w in ws if w['x0'] < COL_DEBITO_X[0]]
                    linha_str = ' '.join(
                        w['text'] for w in sorted(ws_desc, key=lambda w: w['x0'])
                    )
                    # "CONTA 00230.5 - 3.1.01.01.1001 - PRESTAÇÃO DE SERVIÇOS..."
                    m = re.match(
                        r'CONTA\s+([\d.]+)\s+-\s+[\w.]+\s+-\s+(.+)$',
                        linha_str
                    )
                    if m:
                        conta_codigo = m.group(1)
                        conta_nome   = m.group(2).strip()
                    continue

                # ── Linha de lançamento ───────────────────────────────────
                if _e_linha_dados(ws):
                    if registro_atual:
                        registros.append(registro_atual)

                    debito_txt  = _col(ws, *COL_DEBITO_X)
                    credito_txt = _col(ws, *COL_CREDITO_X)
                    saldo_txt   = _col(ws, *COL_SALDO_X)

                    registro_atual = {
                        'conta_codigo': conta_codigo,
                        'conta_nome':   conta_nome,
                        'fil':          _col(ws, *COL_FIL_X),
                        'data':         _col(ws, *COL_DATA_X),
                        'lot':          _col(ws, *COL_LOT_X),
                        'documento':    _col(ws, *COL_DOC_X),
                        'item':         _col(ws, *COL_ITEM_X),
                        'cpart':        _col(ws, *COL_CPART_X),
                        'historico':    _col(ws, *COL_HIST_X),
                        'debito':    _br_para_float(debito_txt)  if VALOR_RE.match(debito_txt)  else None,
                        'credito':   _br_para_float(credito_txt) if VALOR_RE.match(credito_txt) else None,
                        'saldo_atual':_br_para_float(saldo_txt)  if VALOR_RE.match(saldo_txt)   else None,
                    }
                    continue

                # ── Continuação do histórico ──────────────────────────────
                if _e_continuacao_historico(ws) and registro_atual:
                    cont = _col(ws, *COL_HIST_X)
                    if cont:
                        registro_atual['historico'] = (
                            registro_atual['historico'] + ' ' + cont
                        ).strip()

    # Adiciona último registro
    if registro_atual:
        registros.append(registro_atual)

    # ── Monta DataFrame ───────────────────────────────────────────────────
    df = pd.DataFrame(registros, columns=[
        'conta_codigo', 'conta_nome', 'fil', 'data', 'lot',
        'documento', 'item', 'cpart', 'historico',
        'debito', 'credito', 'saldo_atual',
    ])

    df['data']        = pd.to_datetime(df['data'], format='%d/%m/%Y', errors='coerce')
    df['debito']      = pd.to_numeric(df['debito'],     errors='coerce')
    df['credito']     = pd.to_numeric(df['credito'],    errors='coerce')
    df['saldo_atual'] = pd.to_numeric(df['saldo_atual'],errors='coerce')

    return df


# ---------------------------------------------------------------------------
# Execução direta
# ---------------------------------------------------------------------------
if __name__ == '__main__':
    import sys

    caminho = sys.argv[1] if len(sys.argv) > 1 else 'Gerado_por_RDprint__1_.PDF'

    df = extrair_razao_analitico(caminho)

    print(f"\n✅ {len(df)} lançamentos extraídos")
    print(f"   Contas    : {df['conta_codigo'].nunique()}")
    print(f"   Período   : {df['data'].min().date()} → {df['data'].max().date()}")
    print(f"   Total déb : R$ {df['debito'].sum():,.2f}")
    print(f"   Total créd: R$ {df['credito'].sum():,.2f}")
    print(f"\nPrimeiras linhas:\n")
    print(df.head(10).to_string(index=False))

    saida = 'razao_analitico_excel.xlsx'
    # df.to_csv(saida, index=False, encoding='utf-8-sig')
    df.to_excel(saida, index=False)
    print(f"\n💾 Salvo em: {saida}")