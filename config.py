import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Configuração do Banco de Dados
DB_TYPE = os.getenv('DB_TYPE', 'sqlite')

DB_CONFIG = {
    'type': DB_TYPE,
    'sqlite_path': os.path.join(BASE_DIR, 'gestao_contratos.db'),
    'host': 'localhost',
    'user': 'root',
    'password': '',
    'database': 'gestao_contratos',
    'port': 3306
}

# Configurações da Aplicação
APP_CONFIG = {
    'title': 'Prefeitura Municipal de Ribeirãozinho do Maranhão - MA | Gestão de Contratos',
    'icon': '🏛️',
    'municipio': 'PREFEITURA MUNICIPAL DE RIBEIRÃOZINHO DO MARANHÃO - MA',
    'estado': 'MA',
    'dias_alerta': 30
}

SECRET_KEY = 'gestao-contratos-secret-key-2026'
