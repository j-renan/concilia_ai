import os
import pandas as pd
from flask import Flask, render_template, request, jsonify, send_file, session
from concilar_planilhas import comparar_fretes, gerar_planilha_diferencas
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'concilia_ai_secret_key'
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def limpar_uploads():
    """Remove todos os arquivos da pasta uploads para economizar espaço."""
    for filename in os.listdir(UPLOAD_FOLDER):
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                import shutil
                shutil.rmtree(file_path)
        except Exception as e:
            print(f'Erro ao deletar {file_path}. Razão: {e}')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    # Limpar envios anteriores antes de começar um novo
    limpar_uploads()
    
    if 'file_credito' not in request.files or 'file_frete' not in request.files:
        return jsonify({'error': 'Arquivos ausentes'}), 400
    
    file_credito = request.files['file_credito']
    file_frete = request.files['file_frete']
    
    if file_credito.filename == '' or file_frete.filename == '':
        return jsonify({'error': 'Nenhum arquivo selecionado'}), 400

    # Salvar temporariamente para ler cabeçalhos
    path_credito = os.path.join(UPLOAD_FOLDER, f"credito_{datetime.now().timestamp()}.xlsx")
    path_frete = os.path.join(UPLOAD_FOLDER, f"frete_{datetime.now().timestamp()}.xlsx")
    
    file_credito.save(path_credito)
    file_frete.save(path_frete)
    
    try:
        df_credito = pd.read_excel(path_credito, nrows=0)
        df_frete = pd.read_excel(path_frete, nrows=0)
        
        headers_credito = df_credito.columns.tolist()
        headers_frete = df_frete.columns.tolist()
        
        session['path_credito'] = path_credito
        session['path_frete'] = path_frete
        
        return jsonify({
            'headers_credito': headers_credito,
            'headers_frete': headers_frete
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/process', methods=['POST'])
def process():
    mapping = request.json.get('mapping')
    path_credito = session.get('path_credito')
    path_frete = session.get('path_frete')
    
    if not all([mapping, path_credito, path_frete]):
        return jsonify({'error': 'Dados de processamento ausentes'}), 400
    
    try:
        divergencias, docs_credito_sem_frete, docs_frete_sem_credito = comparar_fretes(
            path_credito, path_frete, mapping
        )
        
        # Guardar divergências para exportação
        export_path = os.path.join(UPLOAD_FOLDER, f"divergencias_{datetime.now().timestamp()}.xlsx")
        divergencias.to_excel(export_path, index=False)
        session['export_path'] = export_path
        
        return jsonify({
            'summary': {
                'total_divergencias': len(divergencias),
                'total_credito_sem_frete': len(docs_credito_sem_frete),
                'total_frete_sem_credito': len(docs_frete_sem_credito),
                'valor_total_divergencia': float(divergencias['Diferença'].sum()) if not divergencias.empty else 0
            },
            'divergencias': divergencias.to_dict(orient='records'),
            'docs_credito_sem_frete': list(docs_credito_sem_frete),
            'docs_frete_sem_credito': list(docs_frete_sem_credito)
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/export')
def export():
    export_path = session.get('export_path')
    if export_path and os.path.exists(export_path):
        return send_file(export_path, as_attachment=True, download_name='divergencias_frete.xlsx')
    return "Nenhum resultado disponível para exportação", 404

if __name__ == '__main__':
    app.run(debug=True)
