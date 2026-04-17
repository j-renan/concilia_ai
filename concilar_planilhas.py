import pandas as pd
import re
from utils import ler_arquivo

def comparar_fretes(arquivo_credito, arquivo_frete, mapping=None):

    if mapping is None:
        mapping = {
            'credito': {'historico': 'Histórico', 'valor': 'Crédito'},
            'frete': {'documento': 'Documento', 'valor': 'Valor Frete', 'destinatario': 'Destinatário'}
        }

    df_credito = ler_arquivo(arquivo_credito)
    df_frete = ler_arquivo(arquivo_frete)

    df_credito.columns = [c.strip() for c in df_credito.columns]
    df_frete.columns = [c.strip() for c in df_frete.columns]

    col_hist_cred = str(mapping['credito']['historico']).strip()
    col_val_cred = str(mapping['credito']['valor']).strip()

    col_doc_frete = str(mapping['frete']['documento']).strip()
    col_val_frete = str(mapping['frete']['valor']).strip()
    col_dest_frete = str(mapping['frete']['destinatario']).strip() if mapping['frete']['destinatario'] else None

    # ==================== FUNÇÕES ====================

    def limpar_valor(val):
        if pd.isna(val) or val == '':
            return None

        if isinstance(val, (int, float)):
            return float(val)

        s = str(val).strip().upper().replace('R$', '').replace('$', '')

        if s.startswith('(') and s.endswith(')'):
            s = '-' + s[1:-1]

        s = re.sub(r'[^0-9,.-]', '', s)

        if not s or s in ['-', '.', ',']:
            return None

        try:
            if ',' in s and '.' in s:
                if s.rfind(',') > s.rfind('.'):
                    s = s.replace('.', '').replace(',', '.')
                else:
                    s = s.replace(',', '')
            elif ',' in s:
                s = s.replace(',', '.')
            elif s.count('.') > 1:
                s = s.replace('.', '')

            return float(s)
        except:
            return None

    def extrair_numero_cte(historico):
        if pd.isna(historico):
            return None

        texto = str(historico)
        match = re.search(r'(?:CTE|NFP|NF)[^\d]*(\d+)', texto)

        if match:
            numero = match.group(1).lstrip('0')
            return int(numero) if numero else None

        return None

    # ==================== PREPARAÇÃO ====================

    df_credito['ID_DOC'] = df_credito[col_hist_cred].apply(extrair_numero_cte)
    df_credito = df_credito.dropna(subset=['ID_DOC'])
    df_credito['ID_DOC'] = df_credito['ID_DOC'].astype(int)
    df_credito['VALOR_LIMPO'] = df_credito[col_val_cred].apply(limpar_valor)

    df_frete['ID_DOC'] = pd.to_numeric(df_frete[col_doc_frete], errors='coerce')
    df_frete = df_frete.dropna(subset=['ID_DOC'])
    df_frete['ID_DOC'] = df_frete['ID_DOC'].astype(int)
    df_frete['VALOR_LIMPO'] = df_frete[col_val_frete].apply(limpar_valor)

    # ==================== SEQ ====================

    df_credito['SEQ'] = df_credito.groupby('ID_DOC').cumcount()
    df_frete['SEQ'] = df_frete.groupby('ID_DOC').cumcount()

    # ==================== MERGE ====================

    df_comparacao = df_credito.merge(
        df_frete,
        on=['ID_DOC', 'SEQ'],
        how='outer',
        suffixes=('_CRED', '_FRETE'),
        indicator=True
    )

    # ==================== CAMPOS ====================

    df_comparacao['Documento'] = df_comparacao['ID_DOC']

    df_comparacao['Valor Crédito'] = df_comparacao['VALOR_LIMPO_CRED']
    df_comparacao['Valor Frete'] = df_comparacao['VALOR_LIMPO_FRETE']

    if col_dest_frete and col_dest_frete in df_comparacao.columns:
        df_comparacao['Destinatário'] = df_comparacao[col_dest_frete].fillna('-')
    else:
        df_comparacao['Destinatário'] = '-'

    df_comparacao['Diferença'] = (
        df_comparacao['Valor Crédito'].fillna(0) -
        df_comparacao['Valor Frete'].fillna(0)
    )

    df_comparacao['Diferença_Abs'] = df_comparacao['Diferença'].abs()

    # ==================== OBS ====================

    def aplicar_obs(row):
        if row['_merge'] == 'left_only':
            return '❌ Não encontrado na Planilha de Frete'
        if row['_merge'] == 'right_only':
            return '❌ Não encontrado na Planilha de Crédito'
        if row['Diferença_Abs'] > 0.01:
            return '⚠️ Divergência de Valor'
        return '✅ OK'

    df_comparacao['Observação'] = df_comparacao.apply(aplicar_obs, axis=1)

    # ==================== FILTRO ====================

    divergencias = df_comparacao[
        (df_comparacao['Diferença_Abs'] > 0.01) |
        (df_comparacao['_merge'] != 'both')
    ].copy()

    if not divergencias.empty:
        divergencias = divergencias.sort_values('Diferença_Abs', ascending=False)
        divergencias['Documento'] = divergencias['Documento'].astype(str)

        colunas_final = [
            'Documento',
            'Valor Crédito',
            'Valor Frete',
            'Diferença',
            'Destinatário',
            'Observação'
        ]

        divergencias = divergencias[[c for c in colunas_final if c in divergencias.columns]]

    docs_credito_sem_frete = df_comparacao[df_comparacao['_merge'] == 'left_only']['ID_DOC'].astype(str).unique().tolist()
    docs_frete_sem_credito = df_comparacao[df_comparacao['_merge'] == 'right_only']['ID_DOC'].astype(str).unique().tolist()

    return divergencias, docs_credito_sem_frete, docs_frete_sem_credito