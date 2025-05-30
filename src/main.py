import os
import sys
from dotenv import load_dotenv
from flask import Flask
from flask_cors import CORS

# Adicionar o diretório pai ao path para importações relativas
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

# Importar a função create_app do módulo routes
from src.routes import create_app

# Carregar variáveis de ambiente
load_dotenv()

# Criar a aplicação Flask
app = create_app()

# Configurar CORS para permitir apenas o domínio do frontend
CORS(app, resources={r"/*": {"origins": [
    "https://frontend-production-73ab.up.railway.app",
    "http://localhost:5000",  # Para desenvolvimento local
    "http://127.0.0.1:5000"   # Para desenvolvimento local
]}})

# Executar a aplicação se este arquivo for executado diretamente
if __name__ == "__main__":
    port = int(os.getenv("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=False)
