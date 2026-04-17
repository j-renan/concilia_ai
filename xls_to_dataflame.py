import pandas as pd
import os
from bs4 import BeautifulSoup
import re


def extrair_html_workbook_excel(caminho_arquivo_principal):
    """
    Extrai dados de um HTML Workbook do Excel (arquivo .xls que na verdade é HTML).

    Parâmetros:
    -----------
    caminho_arquivo_principal : str
        Caminho para o arquivo HTML principal (Janeiro 2025 (1).xls)

    Retorna:
    --------
    pandas.DataFrame ou None
        DataFrame com os dados extraídos
    """

    # Obter o diretório base
    diretorio_base = os.path.dirname(caminho_arquivo_principal)
    nome_base = os.path.splitext(os.path.basename(caminho_arquivo_principal))[0]

    # Possíveis caminhos para a pasta de arquivos
    possiveis_pastas = [
        os.path.join(diretorio_base, f"{nome_base}_arquivos"),
        os.path.join(diretorio_base, f"{nome_base} arquivos"),
        os.path.join(diretorio_base, f"{nome_base}_files"),
        os.path.join(diretorio_base, "sheet001_files"),
        diretorio_base
    ]

    # Possíveis nomes para o arquivo da planilha
    possiveis_planilhas = [
        "sheet001.htm",
        "sheet001.html",
        "Sheet001.htm",
        "Sheet001.html",
        "index.htm",
        "index.html"
    ]

    # Tentar encontrar a planilha
    planilha_encontrada = None
    pasta_encontrada = None

    for pasta in possiveis_pastas:
        for planilha in possiveis_planilhas:
            caminho_planilha = os.path.join(pasta, planilha)
            if os.path.exists(caminho_planilha):
                planilha_encontrada = caminho_planilha
                pasta_encontrada = pasta
                break
        if planilha_encontrada:
            break

    if not planilha_encontrada:
        print(f"❌ Não foi possível encontrar a pasta de arquivos auxiliares")
        print(f"   Procurado em: {possiveis_pastas}")
        return None

    print(f"✓ Planilha encontrada: {planilha_encontrada}")

    # Ler o arquivo HTML da planilha
    try:
        with open(planilha_encontrada, 'r', encoding='utf-8') as f:
            html_content = f.read()

        # Extrair tabelas do HTML
        df = extrair_tabela_do_html(html_content)

        if df is not None and not df.empty:
            print(f"✓ Dados extraídos com sucesso!")
            print(f"  - Linhas: {len(df)}")
            print(f"  - Colunas: {len(df.columns)}")
            return df
        else:
            print("❌ Nenhuma tabela encontrada no arquivo da planilha")
            return None

    except Exception as e:
        print(f"❌ Erro ao ler a planilha: {str(e)}")
        return None


def extrair_tabela_do_html(html_content):
    """
    Extrai tabelas de um arquivo HTML, especificamente para HTML do Excel.
    """
    soup = BeautifulSoup(html_content, 'html.parser')

    # Procurar por tabelas
    tabelas = soup.find_all('table')

    if not tabelas:
        print("⚠ Nenhuma tabela encontrada no HTML")
        return None

    # Para HTML do Excel, a tabela principal geralmente tem atributos específicos
    tabela_principal = None

    # Procurar tabela com atributos do Excel
    for tabela in tabelas:
        if tabela.get('x:str') is not None or \
                tabela.get('cellspacing') == '0' or \
                'Excel' in str(tabela.get('class', [])):
            tabela_principal = tabela
            break

    # Se não encontrar, pegar a primeira tabela não vazia
    if not tabela_principal:
        for tabela in tabelas:
            if len(tabela.find_all('tr')) > 1:  # Mais que uma linha (cabeçalho + dados)
                tabela_principal = tabela
                break

    if not tabela_principal:
        tabela_principal = tabelas[0]

    # Extrair dados da tabela
    dados = []
    cabecalhos = []

    # Encontrar todas as linhas
    linhas = tabela_principal.find_all('tr')

    if not linhas:
        return None

    # Primeiro, tentar encontrar cabeçalhos
    primeira_linha = linhas[0]
    celulas_cabecalho = primeira_linha.find_all(['th', 'td'])

    # Verificar se a primeira linha parece ser cabeçalho
    tem_cabecalho = False
    for celula in celulas_cabecalho:
        texto = celula.get_text().strip()
        # Se tiver texto e não for número vazio
        if texto and not texto.replace('.', '').replace(',', '').isdigit():
            tem_cabecalho = True
            break

    if tem_cabecalho:
        # Usar primeira linha como cabeçalho
        cabecalhos = [celula.get_text().strip() for celula in celulas_cabecalho]
        linhas_para_processar = linhas[1:]
    else:
        # Gerar cabeçalhos automáticos
        num_colunas = len(celulas_cabecalho)
        cabecalhos = [f"Coluna_{i + 1}" for i in range(num_colunas)]
        linhas_para_processar = linhas

    # Processar cada linha
    for linha in linhas_para_processar:
        celulas = linha.find_all(['td', 'th'])
        linha_dados = []

        for i, celula in enumerate(celulas):
            # Extrair texto
            texto = celula.get_text().strip()

            # Verificar se é número
            if texto:
                # Remover formatação de moeda e espaços
                texto_limpo = texto.replace('R$', '').replace(' ', '').replace('.', '').replace(',', '.')
                try:
                    # Tentar converter para float
                    valor = float(texto_limpo)
                    linha_dados.append(valor)
                except:
                    # Se não for número, manter como string
                    linha_dados.append(texto)
            else:
                linha_dados.append(None)

        # Preencher colunas faltantes
        while len(linha_dados) < len(cabecalhos):
            linha_dados.append(None)

        # Só adicionar se tiver pelo menos um dado não vazio
        if any(dado is not None and str(dado).strip() for dado in linha_dados):
            dados.append(linha_dados)

    # Criar DataFrame
    if dados:
        df = pd.DataFrame(dados, columns=cabecalhos[:len(dados[0])])
        return df
    else:
        return None


def extrair_todas_as_planilhas(caminho_arquivo_principal):
    """
    Extrai dados de todas as planilhas do HTML Workbook.
    """
    diretorio_base = os.path.dirname(caminho_arquivo_principal)
    nome_base = os.path.splitext(os.path.basename(caminho_arquivo_principal))[0]

    # Procurar por todas as planilhas (sheet001.htm, sheet002.htm, etc.)
    pasta_arquivos = os.path.join(diretorio_base, f"{nome_base}_arquivos")

    if not os.path.exists(pasta_arquivos):
        # Tentar outros nomes de pasta
        possiveis_pastas = [
            f"{nome_base}_arquivos",
            f"{nome_base} arquivos",
            f"{nome_base}_files",
            "sheet001_files"
        ]
        for pasta in possiveis_pastas:
            caminho_pasta = os.path.join(diretorio_base, pasta)
            if os.path.exists(caminho_pasta):
                pasta_arquivos = caminho_pasta
                break

    if not os.path.exists(pasta_arquivos):
        print(f"❌ Pasta de arquivos não encontrada")
        return {}

    # Encontrar todas as planilhas
    planilhas = {}
    import glob

    for arquivo in sorted(glob.glob(os.path.join(pasta_arquivos, "sheet*.htm*"))):
        nome_planilha = os.path.basename(arquivo)
        print(f"\nProcessando {nome_planilha}...")

        with open(arquivo, 'r', encoding='utf-8') as f:
            html_content = f.read()

        df = extrair_tabela_do_html(html_content)
        if df is not None and not df.empty:
            planilhas[nome_planilha] = df
            print(f"  ✓ {len(df)} linhas, {len(df.columns)} colunas")

    return planilhas


# Exemplo de uso
if __name__ == "__main__":
    caminho = "Janeiro_2025.xls"

    print("=== Extraindo dados do HTML Workbook ===\n")

    # Método 1: Extrair planilha principal
    df = extrair_html_workbook_excel(caminho)

    if df is not None and not df.empty:
        print("\n" + "=" * 50)
        print("DADOS EXTRAÍDOS:")
        print("=" * 50)
        print(f"\nPrimeiras 10 linhas:")
        print(df.head(10))

        print(f"\nInformações do DataFrame:")
        print(df.info())

        print(f"\nEstatísticas descritivas (colunas numéricas):")
        print(df.describe())

        # Salvar como Excel válido
        df.to_excel("Janeiro_2025_extraido.xlsx", index=False)
        print(f"\n✓ Dados salvos em: Janeiro_2025_extraido.xlsx")

        # Salvar como CSV
        df.to_csv("Janeiro_2025_extraido.csv", index=False, encoding='utf-8-sig')
        print(f"✓ Dados salvos em: Janeiro_2025_extraido.csv")

    else:
        print("\n❌ Não foi possível extrair os dados")
        print("\nVerifique se a pasta 'Janeiro 2025_arquivos' existe no mesmo diretório")
        print("e contém o arquivo 'sheet001.htm'")

    # Método 2: Extrair todas as planilhas (opcional)
    print("\n" + "=" * 50)
    print("EXTRAINDO TODAS AS PLANILHAS:")
    print("=" * 50)
    todas = extrair_todas_as_planilhas(caminho)

    if todas:
        print(f"\n✓ Total de planilhas extraídas: {len(todas)}")
        for nome, df_planilha in todas.items():
            print(f"\nPlanilha: {nome}")
            print(df_planilha.head())