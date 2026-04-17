import pandas as pd
import pdfplumber
import os

def ler_arquivo(caminho, nrows=None):
    """
    Lê planilhas de arquivos com as extensões: .xlsx, .xls, .xlt, .csv, .pdf.
    Se nrows=0, retorna apenas o cabeçalho (DataFrame vazio com colunas).
    """
    ext = os.path.splitext(caminho)[1].lower()
    
    if ext in ['.xlsx', '.xls', '.xlt']:
        return pd.read_excel(caminho, nrows=nrows)

    elif ext == '.csv':
        try:
            return pd.read_csv(caminho, nrows=nrows, sep=';', encoding='utf-8')
        except:
            return pd.read_csv(caminho, nrows=nrows, sep=',', encoding='utf-8')

    elif ext == '.pdf':
        all_data = []
        headers = None
        
        with pdfplumber.open(caminho) as pdf:
            # Pega as tabelas da primeira página
            if len(pdf.pages) > 0:
                primeiro_table = pdf.pages[0].extract_table()
                if primeiro_table and len(primeiro_table) > 0:
                    headers = primeiro_table[0]
                    # Limpa quebras de linha em headers
                    headers = [str(h).replace('\n', ' ') if h else '' for h in headers]
                    
                    if nrows == 0:
                        return pd.DataFrame(columns=headers)
                    else:
                        all_data.extend(primeiro_table[1:])
                
                # Se precisa de mais linhas, varre as outras páginas
                if nrows != 0:
                    for i in range(1, len(pdf.pages)):
                        page = pdf.pages[i]
                        table = page.extract_table()
                        if table:
                            # Se a primeira linha da página for igual ao header, ignora ela
                            cleaned_first_row = [str(cell).replace('\n', ' ') if cell else '' for cell in table[0]]
                            if cleaned_first_row == headers:
                                all_data.extend(table[1:])
                            else:
                                all_data.extend(table)
        
        if headers:
            df = pd.DataFrame(all_data, columns=headers)
            return df
        return pd.DataFrame()
        
    else:
        raise ValueError(f"Formato de arquivo não suportado: {ext}")
