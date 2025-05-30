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
    global _file_id_cache
    if _file_id_cache:
        return _file_id_cache

    with _file_id_lock:
        if _file_id_cache:
            return _file_id_cache

        token = get_access_token()
        if not token:
            return None

        headers = {
            "Authorization": f"Bearer {token}"
        }

        # Acessar diretamente o OneDrive do usuário
        url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/drive/root:/Documents/formula.xlsx"
        response = requests.get(url, headers=headers)

        if response.status_code != 200:
            print(f"[Arquivo] Erro {response.status_code} - {response.text}")
            return None

        file_id = response.json().get("id")
        _file_id_cache = file_id
        return file_id


def update_cell(cell, value):
    with _excel_operation_lock:
        token = get_access_token()
        file_id = get_excel_file_id()
        if not token or not file_id:
            return False

        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }

        url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/drive/items/{file_id}/workbook/worksheets/{EXCEL_WORKSHEET_NAME}/range(address='{cell}')"
        data = { "values": [[value]] }

        response = requests.patch(url, headers=headers, json=data)
        if response.status_code != 200:
            print(f"Erro ao atualizar célula {cell}: {response.status_code}")
            print(response.text)
            return False
        return True

def get_cell_value(cell):
    token = get_access_token()
    file_id = get_excel_file_id()
    if not token or not file_id:
        return None

    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }

    url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/drive/items/{file_id}/workbook/worksheets/{EXCEL_WORKSHEET_NAME}/range(address='{cell}')"
    response = requests.get(url, headers=headers)
    if response.status_code != 200:
        print(f"Erro ao obter valor da célula {cell}: {response.status_code}")
        print(response.text)
        return None

    values = response.json().get('values', [[None]])
    return values[0][0]

def get_range_values(cell_range):
    token = get_access_token()
    file_id = get_excel_file_id()
    if not token or not file_id:
        return None

    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }

    url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/drive/items/{file_id}/workbook/worksheets/{EXCEL_WORKSHEET_NAME}/range(address='{cell_range}')"
    response = requests.get(url, headers=headers)
    if response.status_code != 200:
        print(f"Erro ao obter valores do intervalo {cell_range}: {response.status_code}")
        print(response.text)
        return None

    return response.json().get('values', [])

def find_next_empty_cell(column, start_row, end_row):
    range_values = get_range_values(f"{column}{start_row}:{column}{end_row}")
    if not range_values:
        return start_row

    for i, cell_value in enumerate(range_values):
        if not cell_value[0]:
            return start_row + i
    return None

def clear_range(cell_range):
    with _excel_operation_lock:
        token = get_access_token()
        file_id = get_excel_file_id()
        if not token or not file_id:
            return False

        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }

        url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/drive/items/{file_id}/workbook/worksheets/{EXCEL_WORKSHEET_NAME}/range(address='{cell_range}')/clear"
        response = requests.post(url, headers=headers)
        if response.status_code != 200:
            print(f"Erro ao limpar intervalo {cell_range}: {response.status_code}")
            print(response.text)
            return False
        return True

def check_connection():
    result = get_cell_value("A1")
    return result is not None
