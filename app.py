import os
import pandas as pd
from flask import Flask, render_template, request, jsonify, send_file, session, redirect, url_for, flash
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from models import db, User
from concilar_planilhas import comparar_fretes
from datetime import datetime
from utils import ler_arquivo

app = Flask(__name__)
app.secret_key = 'concilia_ai_secret_key'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///concilia.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db.init_app(app)
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

def get_user_folder():
    """Retorna o caminho da pasta de upload exclusiva do usuário atual."""
    user_folder = os.path.join(UPLOAD_FOLDER, f'user_{current_user.id}')
    os.makedirs(user_folder, exist_ok=True)
    return user_folder

def limpar_uploads(folder):
    """Remove arquivos apenas da pasta do usuário atual."""
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
        except Exception as e:
            print(f'Erro ao deletar {file_path}. Razão: {e}')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('index'))
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        remember = True if request.form.get('remember') else False
        
        user = User.query.filter_by(username=username).first()
        if user and user.check_password(password):
            login_user(user, remember=remember)
            return redirect(url_for('index'))
        flash('Usuário ou senha inválidos')
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

@app.route('/register', methods=['GET', 'POST'])
@login_required
def register():
    # Apenas administradores podem cadastrar novos usuários
    if not current_user.is_admin:
        flash('Acesso negado: apenas administradores podem cadastrar usuários.')
        return redirect(url_for('index'))
        
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        is_admin = True if request.form.get('is_admin') else False
        
        if User.query.filter_by(username=username).first():
            flash('Este usuário já existe.')
        else:
            new_user = User(username=username, is_admin=is_admin)
            new_user.set_password(password)
            db.session.add(new_user)
            db.session.commit()
            flash('Usuário cadastrado com sucesso!')
            return redirect(url_for('index'))
    return render_template('register.html')

@app.route('/')
@login_required
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
@login_required
def upload_files():
    user_folder = get_user_folder()
    limpar_uploads(user_folder)
    
    if 'file_credito' not in request.files or 'file_frete' not in request.files:
        return jsonify({'error': 'Arquivos ausentes'}), 400
    
    file_credito = request.files['file_credito']
    file_frete = request.files['file_frete']
    
    if file_credito.filename == '' or file_frete.filename == '':
        return jsonify({'error': 'Nenhum arquivo selecionado'}), 400

    ext_credito = os.path.splitext(file_credito.filename)[1]
    ext_frete = os.path.splitext(file_frete.filename)[1]
    
    path_credito = os.path.join(user_folder, f"credito_{datetime.now().timestamp()}{ext_credito}")
    path_frete = os.path.join(user_folder, f"frete_{datetime.now().timestamp()}{ext_frete}")
    
    file_credito.save(path_credito)
    file_frete.save(path_frete)
    
    try:
        df_credito = ler_arquivo(path_credito, nrows=0)
        df_frete = ler_arquivo(path_frete, nrows=0)
        
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


import json

@app.route('/process', methods=['POST'])
@login_required
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

        user_folder = get_user_folder()
        export_path = os.path.join(user_folder, f"divergencias_{datetime.now().timestamp()}.xlsx")
        divergencias.to_excel(export_path, index=False)
        session['export_path'] = export_path

        return jsonify({
            'summary': {
                'total_divergencias': len(divergencias),
                'total_credito_sem_frete': len(docs_credito_sem_frete),
                'total_frete_sem_credito': len(docs_frete_sem_credito),
                'valor_total_divergencia': float(divergencias['Diferença'].sum()) if not divergencias.empty else 0
            },
            # ✅ CORREÇÃO DEFINITIVA AQUI
            'divergencias': json.loads(divergencias.to_json(orient='records')),
            'docs_credito_sem_frete': docs_credito_sem_frete,
            'docs_frete_sem_credito': docs_frete_sem_credito
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/export')
@login_required
def export():
    export_path = session.get('export_path')
    if export_path and os.path.exists(export_path):
        return send_file(export_path, as_attachment=True, download_name='divergencias_frete.xlsx')
    return "Nenhum resultado disponível para exportação", 404

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(debug=True)

if __name__ == '__main__':
    app.run(debug=True)
