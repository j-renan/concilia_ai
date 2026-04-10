import pandas as pd
import re
import unicodedata


def comparar_fretes(arquivo_credito, arquivo_frete, mapping=None):
    """
    Compara os valores de crédito e frete entre duas planilhas com mapeamento flexível.

    Parâmetros:
    - arquivo_credito: caminho ou buffer da planilha de crédito
    - arquivo_frete: caminho ou buffer da planilha de frete
    - mapping: dicionário com o mapeamento das colunas
        {
            'credito': {'historico': '...', 'valor': '...'},
            'frete': {'documento': '...', 'valor': '...', 'destinatario': '...'}
        }
    """
    
    # Mapeamento padrão caso não seja fornecido
    if mapping is None:
        mapping = {
            'credito': {'historico': 'Histórico', 'valor': 'Crédito'},
            'frete': {'documento': 'Documento', 'valor': 'Valor Frete', 'destinatario': 'Destinatário'}
        }

    # ==================== LEITURA DAS PLANILHAS ====================
    df_credito = pd.read_excel(arquivo_credito)
    df_frete = pd.read_excel(arquivo_frete)

    # ==================== NORMALIZAÇÃO DE COLUNAS ====================
    # Remove espaços extras dos nomes das colunas
    df_credito.columns = [c.strip() for c in df_credito.columns]
    df_frete.columns = [c.strip() for c in df_frete.columns]
    
    # ==================== NORMALIZAÇÃO DE MAPEAMENTO ====================
    # Garante que os nomes das colunas vindos do usuário sejam limpos
    col_hist_cred = str(mapping['credito']['historico']).strip()
    col_val_cred = str(mapping['credito']['valor']).strip()
    
    col_doc_frete = str(mapping['frete']['documento']).strip()
    col_val_frete = str(mapping['frete']['valor']).strip()
    col_dest_frete = str(mapping['frete']['destinatario']).strip() if mapping['frete']['destinatario'] else None

    # ==================== FUNÇÕES AUXILIARES ====================
    def limpar_valor(val):
        """Converte qualquer formato monetário para float com inteligência de separadores."""
        if pd.isna(val) or val == '':
            return 0.0
        if isinstance(val, (int, float)):
            return float(val)
        
        # Stringify e limpeza básica
        s = str(val).strip().upper().replace('R$', '').replace('$', '')
        
        # Tratar parênteses como sinal negativo: (100,00) -> -100,00
        if s.startswith('(') and s.endswith(')'):
            s = '-' + s[1:-1]
            
        # Manter apenas números, pontos, vírgulas e sinal de menos
        s = re.sub(r'[^0-9,.-]', '', s)
        if not s or s == '-':
            return 0.0

        # Lógica de decisão de separador decimal
        # Se houver ponto e vírgula, o último decide
        if ',' in s and '.' in s:
            if s.rfind(',') > s.rfind('.'):
                s = s.replace('.', '').replace(',', '.') # 1.234,56 -> 1234.56
            else:
                s = s.replace(',', '') # 1,234.56 -> 1234.56
        elif ',' in s:
            # Se houver mais de uma vírgula (milhar)
            if s.count(',') > 1:
                parts = s.split(',')
                s = "".join(parts[:-1]) + "." + parts[-1]
            else:
                # Uma vírgula só: assumimos decimal em PT-BR
                s = s.replace(',', '.')
        elif s.count('.') > 1:
            # Múltiplos pontos (milhar sem decimal): 1.000.000
            s = s.replace('.', '')

        try:
            return float(s)
        except:
            return 0.0

    # ==================== EXTRAIR ID DO HISTÓRICO ====================
    def extrair_numero_cte(historico):
        if pd.isna(historico):
            return None
        match = re.search(r'CTE N\.\s*(\d+)', str(historico))
        if match:
            try:
                return int(match.group(1))
            except:
                return None
        return None

    # ==================== PREPARAÇÃO DOS DADOS ====================
    # Crédito
    df_credito['ID_DOC'] = df_credito[col_hist_cred].apply(extrair_numero_cte)
    df_credito = df_credito.dropna(subset=['ID_DOC'])
    df_credito['ID_DOC'] = df_credito['ID_DOC'].astype(int)
    df_credito['VALOR_LIMPO'] = df_credito[col_val_cred].apply(limpar_valor)
    
    # Frete
    df_frete['ID_DOC'] = pd.to_numeric(df_frete[col_doc_frete], errors='coerce')
    df_frete = df_frete.dropna(subset=['ID_DOC'])
    df_frete['ID_DOC'] = df_frete['ID_DOC'].astype(int)
    df_frete['VALOR_LIMPO'] = df_frete[col_val_frete].apply(limpar_valor)

    # ==================== AGREGAÇÃO (SAFA E EXPLÍCITA) ====================
    # Crédito: Agrupar por ID e somar o valor limpo
    df_cred_agg = df_credito.groupby('ID_DOC', as_index=False).agg({'VALOR_LIMPO': 'sum'})
    df_cred_agg = df_cred_agg.rename(columns={'ID_DOC': 'ID_JOIN', 'VALOR_LIMPO': 'VAL_CRED_FINAL'})

    # Frete: Agrupar por ID, somar valor e pegar primeiro destinatário
    map_agg_frete = {'VALOR_LIMPO': 'sum'}
    if col_dest_frete and col_dest_frete in df_frete.columns:
        map_agg_frete[col_dest_frete] = 'first'
        
    df_frete_agg = df_frete.groupby('ID_DOC', as_index=False).agg(map_agg_frete)
    
    # Renomeação explícita usando dicionário para evitar confusão de ordem
    rename_dict_frete = {'ID_DOC': 'ID_JOIN_FRETE', 'VALOR_LIMPO': 'VAL_FRETE_FINAL'}
    if col_dest_frete:
        rename_dict_frete[col_dest_frete] = 'DEST_FINAL'
    df_frete_agg = df_frete_agg.rename(columns=rename_dict_frete)

    # ==================== OUTER MERGE ====================
    df_comparacao = df_cred_agg.merge(
        df_frete_agg,
        left_on='ID_JOIN',
        right_on='ID_JOIN_FRETE',
        how='outer'
    )

    # Consolidação do Documento e Preenchimento de Zeros
    df_comparacao['Documento'] = df_comparacao['ID_JOIN'].fillna(df_comparacao['ID_JOIN_FRETE']).astype(int)
    df_comparacao['Valor Crédito'] = df_comparacao['VAL_CRED_FINAL'].fillna(0.0)
    df_comparacao['Valor Frete'] = df_comparacao['VAL_FRETE_FINAL'].fillna(0.0)
    
    if 'DEST_FINAL' in df_comparacao.columns:
        df_comparacao['Destinatário'] = df_comparacao['DEST_FINAL'].fillna('-')
    else:
        df_comparacao['Destinatário'] = '-'

    # Diferenças
    df_comparacao['Diferença'] = df_comparacao['Valor Crédito'] - df_comparacao['Valor Frete']
    df_comparacao['Diferença_Abs'] = df_comparacao['Diferença'].abs()

    # Observações
    def aplicar_obs(row):
        if pd.isna(row['ID_JOIN']):
            return '❌ Não encontrado na Planilha de Crédito'
        if pd.isna(row['ID_JOIN_FRETE']):
            return '❌ Não encontrado na Planilha de Frete'
        if row['Diferença_Abs'] > 0.01:
            return '⚠️ Divergência de Valor'
        return '✅ OK'

    df_comparacao['Observação'] = df_comparacao.apply(aplicar_obs, axis=1)

    # Filtrar apenas o que importa
    divergencias = df_comparacao[
        (df_comparacao['Diferença_Abs'] > 0.01) | 
        (df_comparacao['Observação'].str.contains('Não encontrado', na=False))
    ].copy()

    # Ordenação e Seleção Final das Colunas
    if not divergencias.empty:
        divergencias = divergencias.sort_values('Diferença_Abs', ascending=False)
        # CONVERSÃO CRÍTICA PARA JSON: NumPy int64 -> str
        divergencias['Documento'] = divergencias['Documento'].astype(str)
        colunas_final = ['Documento', 'Valor Crédito', 'Valor Frete', 'Diferença', 'Destinatário', 'Observação']
        existentes = [c for c in colunas_final if c in divergencias.columns]
        divergencias = divergencias[existentes]

    # Documentos faltantes (também convertidos para string para evitar erro de serialização)
    missing_cred = [str(x) for x in (set(df_credito['ID_DOC'].unique()) - set(df_frete['ID_DOC'].unique()))]
    missing_frete = [str(x) for x in (set(df_frete['ID_DOC'].unique()) - set(df_credito['ID_DOC'].unique()))]

    return divergencias, missing_cred, missing_frete


def gerar_planilha_diferencas(divergencias, arquivo_saida='divergencias_frete.xlsx'):
    """
    Gera uma planilha com as divergências encontradas.
    """
    if divergencias.empty:
        print("✅ Nenhuma divergência encontrada para exportar.")
        return

    divergencias.to_excel(arquivo_saida, index=False)
    print(f"\n📁 Planilha com divergências salva em: {arquivo_saida}")


def gerar_relatorio_documentos_nao_encontrados(docs_credito_sem_frete, docs_frete_sem_credito):
    """
    Gera um relatório detalhado dos documentos não encontrados.
    """
    print("\n" + "=" * 80)
    print("📋 RELATÓRIO DE DOCUMENTOS NÃO ENCONTRADOS:")
    print("=" * 80)
    
    if docs_credito_sem_frete:
        print(f"\n📄 Documentos na Planilha de CRÉDITO que NÃO foram encontrados na Planilha de FRETE:")
        print(f"   Total: {len(docs_credito_sem_frete)} documento(s)")
        print(f"   Documentos: {sorted(docs_credito_sem_frete)}")
    else:
        print("\n✅ Todos os documentos da Planilha de CRÉDITO foram encontrados na Planilha de FRETE")
    
    if docs_frete_sem_credito:
        print(f"\n📄 Documentos na Planilha de FRETE que NÃO foram encontrados na Planilha de CRÉDITO:")
        print(f"   Total: {len(docs_frete_sem_credito)} documento(s)")
        print(f"   Documentos: {sorted(docs_frete_sem_credito)}")
    else:
        print("\n✅ Todos os documentos da Planilha de FRETE foram encontrados na Planilha de CRÉDITO")
    
    total_nao_encontrados = len(docs_credito_sem_frete) + len(docs_frete_sem_credito)
    print(f"\n📊 TOTAL DE DOCUMENTOS NÃO ENCONTRADOS: {total_nao_encontrados}")


# ==================== EXEMPLO DE USO (CLI) ====================
if __name__ == "__main__":
    arquivo_credito = "./uploads/PLANILHA CREDITO.xlsx"
    arquivo_frete = "./uploads/PLANILHA DE FRETE.xlsx"

    divergencias, docs_credito_sem_frete, docs_frete_sem_credito = comparar_fretes(arquivo_credito, arquivo_frete)

    gerar_relatorio_documentos_nao_encontrados(docs_credito_sem_frete, docs_frete_sem_credito)
    
    if not divergencias.empty:
        print("\n" + "=" * 80)
        print("⚠️ DIVERGÊNCIAS DETALHADAS (Inclui documentos não encontrados):")
        print("=" * 80)
        print(divergencias.to_string(index=False))
        gerar_planilha_diferencas(divergencias)
    else:
        print("\n✅ Nenhuma divergência encontrada!")

    print("\n" + "=" * 80)
    print("ANÁLISE CONCLUÍDA!")
    print("=" * 80)