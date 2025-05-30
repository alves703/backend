import os
import requests
import threading
from dotenv import load_dotenv
from .auth import get_access_token

# Carregar variáveis de ambiente
load_dotenv()

# Configurações do Microsoft Graph API
EXCEL_WORKSHEET_NAME = os.getenv('EXCEL_WORKSHEET_NAME')
USER_ID = os.getenv('USER_ID')

# Lock para operações concorrentes
_excel_operation_lock = threading.Lock()

def get_file_id():
    """
    Obtém o ID do arquivo Excel diretamente da raiz do OneDrive.
    """
    token = get_access_token()
    if not token:
        print("[Excel] ERRO: Não foi possível obter token de acesso")
        return None
    
    if not USER_ID:
        print("[Excel] ERRO: USER_ID não está definido nas variáveis de ambiente")
        return None
    
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    
    # Acessar o arquivo diretamente na raiz do OneDrive
    url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/drive/root:/formula.xlsx"
    
    print(f"[Excel] Obtendo ID do arquivo Excel na raiz do OneDrive...")
    response = requests.get(url, headers=headers)
    
    if response.status_code != 200:
        print(f"[Excel] ERRO: Falha ao obter ID do arquivo: {response.status_code}")
        print(f"[Excel] Resposta: {response.text}")
        return None
    
    file_id = response.json().get("id")
    print(f"[Excel] ID do arquivo obtido com sucesso: {file_id}")
    return file_id

def update_cell(cell, value):
    """
    Atualiza uma célula na planilha.
    Usa lock para evitar operações concorrentes.
    """
    with _excel_operation_lock:
        token = get_access_token()
        file_id = get_file_id()
        
        if not token or not file_id:
            print(f"[Excel] ERRO: Não foi possível atualizar célula {cell}. Token ou file_id inválidos.")
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
        
        print(f"[Excel] Atualizando célula {cell} com valor: {value}")
        response = requests.patch(url, headers=headers, json=data)
        
        if response.status_code != 200:
            print(f"[Excel] ERRO: Falha ao atualizar célula {cell}: {response.status_code}")
            print(f"[Excel] Resposta: {response.text}")
            return False
        
        print(f"[Excel] Célula {cell} atualizada com sucesso")
        return True

def get_cell_value(cell):
    """
    Obtém o valor de uma célula.
    """
    token = get_access_token()
    file_id = get_file_id()
    
    if not token or not file_id:
        print(f"[Excel] ERRO: Não foi possível ler célula {cell}. Token ou file_id inválidos.")
        return None
    
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    
    # Endpoint para obter valor da célula (usando USER_ID em vez de me)
    url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/drive/items/{file_id}/workbook/worksheets/{EXCEL_WORKSHEET_NAME}/range(address='{cell}')"
    
    print(f"[Excel] Lendo valor da célula {cell}")
    response = requests.get(url, headers=headers)
    
    if response.status_code != 200:
        print(f"[Excel] ERRO: Falha ao ler célula {cell}: {response.status_code}")
        print(f"[Excel] Resposta: {response.text}")
        return None
    
    values = response.json().get('values', [[None]])
    print(f"[Excel] Valor lido da célula {cell}: {values[0][0]}")
    return values[0][0]

def get_range_values(cell_range):
    """
    Obtém valores de um intervalo de células.
    """
    token = get_access_token()
    file_id = get_file_id()
    
    if not token or not file_id:
        print(f"[Excel] ERRO: Não foi possível ler intervalo {cell_range}. Token ou file_id inválidos.")
        return None
    
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    
    # Endpoint para obter valores do intervalo (usando USER_ID em vez de me)
    url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/drive/items/{file_id}/workbook/worksheets/{EXCEL_WORKSHEET_NAME}/range(address='{cell_range}')"
    
    print(f"[Excel] Lendo valores do intervalo {cell_range}")
    response = requests.get(url, headers=headers)
    
    if response.status_code != 200:
        print(f"[Excel] ERRO: Falha ao ler intervalo {cell_range}: {response.status_code}")
        print(f"[Excel] Resposta: {response.text}")
        return None
    
    values = response.json().get('values', [])
    print(f"[Excel] {len(values)} linhas lidas do intervalo {cell_range}")
    return values

def find_next_empty_cell(column, start_row, end_row):
    """
    Encontra a próxima célula vazia em uma coluna.
    """
    range_values = get_range_values(f"{column}{start_row}:{column}{end_row}")
    
    if not range_values:
        print(f"[Excel] Nenhum valor encontrado no intervalo {column}{start_row}:{column}{end_row}, retornando linha inicial {start_row}")
        return start_row
    
    for i, cell_value in enumerate(range_values):
        if not cell_value[0]:
            next_row = start_row + i
            print(f"[Excel] Próxima célula vazia encontrada: {column}{next_row}")
            return next_row
    
    print(f"[Excel] Nenhuma célula vazia encontrada no intervalo {column}{start_row}:{column}{end_row}")
    return None  # Todas as células estão preenchidas

def clear_range(cell_range):
    """
    Limpa um intervalo de células.
    Usa lock para evitar operações concorrentes.
    """
    with _excel_operation_lock:
        token = get_access_token()
        file_id = get_file_id()
        
        if not token or not file_id:
            print(f"[Excel] ERRO: Não foi possível limpar intervalo {cell_range}. Token ou file_id inválidos.")
            return False
        
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        
        # Endpoint para limpar intervalo (usando USER_ID em vez de me)
        url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/drive/items/{file_id}/workbook/worksheets/{EXCEL_WORKSHEET_NAME}/range(address='{cell_range}')/clear"
        
        print(f"[Excel] Limpando intervalo {cell_range}")
        response = requests.post(url, headers=headers)
        
        if response.status_code != 200:
            print(f"[Excel] ERRO: Falha ao limpar intervalo {cell_range}: {response.status_code}")
            print(f"[Excel] Resposta: {response.text}")
            return False
        
        print(f"[Excel] Intervalo {cell_range} limpo com sucesso")
        return True

def check_connection():
    """
    Verifica a conexão com a planilha.
    Útil para monitoramento e diagnóstico.
    """
    print("[Excel] Verificando conexão com a planilha...")
    
    # Verificar se as variáveis de ambiente estão configuradas
    if not EXCEL_WORKSHEET_NAME:
        print("[Excel] ERRO: EXCEL_WORKSHEET_NAME não está definido nas variáveis de ambiente")
        return False
    
    if not USER_ID:
        print("[Excel] ERRO: USER_ID não está definido nas variáveis de ambiente")
        return False
    
    # Tenta obter o ID do arquivo para verificar a conexão
    file_id = get_file_id()
    if not file_id:
        print("[Excel] ERRO: Não foi possível obter o ID do arquivo Excel")
        return False
    
    print("[Excel] Conexão com a planilha estabelecida com sucesso")
    return True
