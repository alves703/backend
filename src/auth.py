import os
import msal
from dotenv import load_dotenv
import time

# Carregar variáveis de ambiente
load_dotenv()

# Configurações do Microsoft Graph API
TENANT_ID = os.getenv('TENANT_ID')
CLIENT_ID = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
USER_ID = os.getenv('USER_ID')

# Cache para token
_token_cache = {
    "access_token": None,
    "expires_at": 0
}

def get_access_token():
    """
    Obtém um token de acesso para a Microsoft Graph API.
    Implementa cache para evitar requisições desnecessárias.
    """
    global _token_cache
    
    # Verificar se o token em cache ainda é válido (com margem de segurança de 5 minutos)
    current_time = time.time()
    if _token_cache["access_token"] and _token_cache["expires_at"] > current_time + 300:
        return _token_cache["access_token"]
    
    # Token expirado ou não existe, obter um novo
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET
    )
    
    scopes = ["https://graph.microsoft.com/.default"]
    result = app.acquire_token_for_client(scopes)
    
    if "access_token" in result:
        # Armazenar token em cache com tempo de expiração
        _token_cache["access_token"] = result["access_token"]
        _token_cache["expires_at"] = current_time + result.get("expires_in", 3600)
        return result["access_token"]
    else:
        print(f"Erro ao obter token: {result.get('error')}")
        print(f"Descrição: {result.get('error_description')}")
        return None
