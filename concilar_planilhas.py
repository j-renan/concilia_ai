import pandas as pd
import re


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

    # Identificar colunas baseadas no mapeamento
    col_hist_cred = mapping['credito']['historico']
    col_val_cred = mapping['credito']['valor']
    
    col_doc_frete = mapping['frete']['documento']
    col_val_frete = mapping['frete']['valor']
    col_dest_frete = mapping['frete']['destinatario']

    # ==================== EXTRAIR NÚMERO DO DOCUMENTO DO HISTÓRICO ====================
    def extrair_numero_cte(historico):
        if pd.isna(historico):
            return None
        match = re.search(r'CTE N\.\s*(\d+)', str(historico))
        if match:
            return int(match.group(1))
        return None

    df_credito['Documento_Extraido'] = df_credito[col_hist_cred].apply(extrair_numero_cte)

    # ==================== LIMPEZA DOS DADOS ====================
    df_credito = df_credito.dropna(subset=['Documento_Extraido'])
    df_credito['Documento_Extraido'] = df_credito['Documento_Extraido'].astype(int)

    # Garantir que os documentos no df_frete sejam inteiros e remover NaNs
    df_frete = df_frete.dropna(subset=[col_doc_frete])
    df_frete[col_doc_frete] = df_frete[col_doc_frete].astype(int)

    # ==================== REALIZAR O MERGE ====================
    df_comparacao = df_credito.merge(
        df_frete[[col_doc_frete, col_val_frete, col_dest_frete]],
        left_on='Documento_Extraido',
        right_on=col_doc_frete,
        how='inner'
    )

    # ==================== CALCULAR DIFERENÇAS ====================
    df_comparacao['Diferença'] = df_comparacao[col_val_cred] - df_comparacao[col_val_frete]
    df_comparacao['Diferença_Abs'] = df_comparacao['Diferença'].abs()

    # ==================== IDENTIFICAR DIVERGÊNCIAS ====================
    divergencias = df_comparacao[df_comparacao['Diferença_Abs'] > 0.01].copy()

    # ==================== VERIFICAR DOCUMENTOS FALTANTES ====================
    documentos_credito_sem_frete = set(df_credito['Documento_Extraido']) - set(df_frete[col_doc_frete])
    documentos_frete_sem_credito = set(df_frete[col_doc_frete]) - set(df_credito['Documento_Extraido'])

    # ==================== FORMATAR RESULTADOS ====================
    if not divergencias.empty:
        divergencias = divergencias.sort_values('Diferença_Abs', ascending=False)
        # Renomear para colunas amigáveis no resultado final
        divergencias = divergencias.rename(columns={
            'Documento_Extraido': 'Documento',
            col_val_cred: 'Crédito',
            col_val_frete: 'Valor Frete',
            col_dest_frete: 'Destinatário'
        })
        divergencias = divergencias[['Documento', 'Crédito', 'Valor Frete', 'Diferença', 'Destinatário']]

    return divergencias, documentos_credito_sem_frete, documentos_frete_sem_credito


def gerar_planilha_diferencas(divergencias, arquivo_saida='divergencias_frete.xlsx'):
    """
    Gera uma planilha com as divergências encontradas.
    """
    if divergencias.empty:
        print("✅ Nenhuma divergência encontrada para exportar.")
        return

    divergencias.to_excel(arquivo_saida, index=False)
    print(f"\n📁 Planilha com divergências salva em: {arquivo_saida}")


# ==================== EXEMPLO DE USO ====================
if __name__ == "__main__":
    # Substitua pelos caminhos reais dos seus arquivos
    arquivo_credito = "01-102025.xlsx"
    arquivo_frete = "val.xlsx"

    # Executar a comparação
    divergencias, docs_credito_sem_frete, docs_frete_sem_credito = comparar_fretes(
        arquivo_credito,
        arquivo_frete
    )

    # Exibir divergências detalhadas
    if not divergencias.empty:
        print("\n" + "=" * 80)
        print("DIVERGÊNCIAS DETALHADAS:")
        print("=" * 80)
        print(divergencias.to_string(index=False))

        # Salvar planilha com divergências
        gerar_planilha_diferencas(divergencias)

    print("\n" + "=" * 80)
    print("ANÁLISE CONCLUÍDA!")
    print("=" * 80)