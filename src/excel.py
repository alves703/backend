import os
import requests
import threading
from dotenv import load_dotenv
from .auth import get_access_token

# Carregar variáveis de ambiente
load_dotenv()

# Configurações do Microsoft Graph API
EXCEL_FILE_PATH = os.getenv('EXCEL_FILE_PATH')
EXCEL_WORKSHEET_NAME = os.getenv('EXCEL_WORKSHEET_NAME')
USER_ID = os.getenv('USER_ID')

# Cache para file_id
_file_id_cache = None
_file_id_lock = threading.Lock()

# Lock para operações concorrentes
_excel_operation_lock = threading.Lock()

def get_excel_file_id():
    """
    Obtém o ID do arquivo Excel.
    Implementa cache para evitar requisições desnecessárias.
    """
    global _file_id_cache
    
    # Verificar se o file_id já está em cache
    if _file_id_cache:
        return _file_id_cache
    
    with _file_id_lock:
        # Verificar novamente dentro do lock (para evitar condição de corrida)
        if _file_id_cache:
            return _file_id_cache
        
        token = get_access_token()
        if not token:
            return None
        
        # Extrair o site e o caminho do arquivo
        parts = EXCEL_FILE_PATH.split('/')
        site_path = '/'.join(parts[:2])  # /personal/acesso_nuvemedge_com
        file_path = '/'.join(parts[2:])  # Documents/formula.xlsx
        
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        
        # Obter o ID do site
        site_url = f"https://graph.microsoft.com/v1.0/sites/nuvemedge.sharepoint.com:{site_path}"
        response = requests.get(site_url, headers=headers)
        
        if response.status_code != 200:
            print(f"Erro ao obter ID do site: {response.status_code}")
            print(response.text)
            return None
        
        site_id = response.json().get('id')
        
        # Obter o ID do arquivo
        drive_item_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_path}"
        response = requests.get(drive_item_url, headers=headers)
        
        if response.status_code != 200:
            print(f"Erro ao obter ID do arquivo: {response.status_code}")
            print(response.text)
            return None
        
        # Armazenar file_id em cache
        _file_id_cache = response.json().get('id')
        return _file_id_cache

def update_cell(cell, value):
    """
    Atualiza uma célula na planilha.
    Usa lock para evitar operações concorrentes.
    """
    with _excel_operation_lock:
        token = get_access_token()
        file_id = get_excel_file_id()
        
        if not token or not file_id:
            return False
        
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        
        # Endpoint para atualizar célula (usando USER_ID em vez de me)
        url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/drive/items/{file_id}/workbook/worksheets/{EXCEL_WORKSHEET_NAME}/range(address='{cell}')"
        
        data = {
            "values": [[value]]
        }
        
        response = requests.patch(url, headers=headers, json=data)
        
        if response.status_code != 200:
            print(f"Erro ao atualizar célula {cell}: {response.status_code}")
            print(response.text)
            return False
        
        return True

def get_cell_value(cell):
    """
    Obtém o valor de uma célula.
    """
    token = get_access_token()
    file_id = get_excel_file_id()
    
    if not token or not file_id:
        return None
    
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    
    # Endpoint para obter valor da célula (usando USER_ID em vez de me)
    url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/drive/items/{file_id}/workbook/worksheets/{EXCEL_WORKSHEET_NAME}/range(address='{cell}')"
    
    response = requests.get(url, headers=headers)
    
    if response.status_code != 200:
        print(f"Erro ao obter valor da célula {cell}: {response.status_code}")
        print(response.text)
        return None
    
    values = response.json().get('values', [[None]])
    return values[0][0]

def get_range_values(cell_range):
    """
    Obtém valores de um intervalo de células.
    """
    token = get_access_token()
    file_id = get_excel_file_id()
    
    if not token or not file_id:
        return None
    
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    
    # Endpoint para obter valores do intervalo (usando USER_ID em vez de me)
    url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/drive/items/{file_id}/workbook/worksheets/{EXCEL_WORKSHEET_NAME}/range(address='{cell_range}')"
    
    response = requests.get(url, headers=headers)
    
    if response.status_code != 200:
        print(f"Erro ao obter valores do intervalo {cell_range}: {response.status_code}")
        print(response.text)
        return None
    
    return response.json().get('values', [])

def find_next_empty_cell(column, start_row, end_row):
    """
    Encontra a próxima célula vazia em uma coluna.
    """
    range_values = get_range_values(f"{column}{start_row}:{column}{end_row}")
    
    if not range_values:
        return start_row
    
    for i, cell_value in enumerate(range_values):
        if not cell_value[0]:
            return start_row + i
    
    return None  # Todas as células estão preenchidas

def clear_range(cell_range):
    """
    Limpa um intervalo de células.
    Usa lock para evitar operações concorrentes.
    """
    with _excel_operation_lock:
        token = get_access_token()
        file_id = get_excel_file_id()
        
        if not token or not file_id:
            return False
        
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        
        # Endpoint para limpar intervalo (usando USER_ID em vez de me)
        url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/drive/items/{file_id}/workbook/worksheets/{EXCEL_WORKSHEET_NAME}/range(address='{cell_range}')/clear"
        
        response = requests.post(url, headers=headers)
        
        if response.status_code != 200:
            print(f"Erro ao limpar intervalo {cell_range}: {response.status_code}")
            print(response.text)
            return False
        
        return True

def check_connection():
    """
    Verifica a conexão com a planilha.
    Útil para monitoramento e diagnóstico.
    """
    # Tenta ler uma célula simples (A1) para verificar a conexão
    result = get_cell_value("A1")
    return result is not None
