import os
from app import app
from models import db, User

def create_initial_admin():
    with app.app_context():
        # Criar tabelas se não existirem
        db.create_all()
        
        # Verificar se admin já existe
        admin = User.query.filter_by(username='admin').first()
        if not admin:
            admin = User(username='admin', is_admin=True)
            admin.set_password('admin123') # Senha padrão para alteração posterior
            db.session.add(admin)
            db.session.commit()
            print("Usuário administrador criado: admin / admin123")
        else:
            print("Administrador já existe.")

if __name__ == '__main__':
    create_initial_admin()
