import os
from dotenv import load_dotenv
from src.routes import create_app

# Carregar variáveis de ambiente
load_dotenv()

# Criar aplicação Flask
app = create_app()

if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    app.run(debug=True, host='0.0.0.0', port=port)
