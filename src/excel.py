# -*- coding: utf-8 -*-
import os
import requests
import threading
from dotenv import load_dotenv
from .auth import get_access_token

# Carregar variáveis de ambiente
load_dotenv()

# Configurações do Microsoft Graph API
EXCEL_WORKSHEET_NAME = os.getenv("EXCEL_WORKSHEET_NAME")
USER_ID = os.getenv("USER_ID")

# Cache simples para o ID do arquivo
_file_id_cache = None
_file_id_lock = threading.Lock()

# Lock para operações concorrentes de escrita/limpeza
_excel_write_lock = threading.Lock()

def get_cached_file_id():
    """
    Obtém o ID do arquivo Excel, usando cache para evitar chamadas repetidas.
    """
    global _file_id_cache
    if _file_id_cache:
        print("[Excel] Usando ID do arquivo em cache")
        return _file_id_cache

    with _file_id_lock:
        # Verificar novamente após adquirir o lock, caso outra thread já tenha preenchido
        if _file_id_cache:
            print("[Excel] Usando ID do arquivo em cache (após lock)")
            return _file_id_cache

        print("[Excel] Cache do ID do arquivo vazio, buscando na API...")
        token = get_access_token()
        if not token:
            print("[Excel] ERRO: Não foi possível obter token de acesso para buscar file_id")
            return None

        if not USER_ID:
            print("[Excel] ERRO: USER_ID não está definido nas variáveis de ambiente")
            return None

        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }

        # Acessar o arquivo diretamente na raiz do OneDrive
        url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/drive/root:/formula.xlsx"

        print("[Excel] Obtendo ID do arquivo Excel na raiz do OneDrive...")
        try:
            response = requests.get(url, headers=headers, timeout=15) # Adicionado timeout
            response.raise_for_status() # Levanta exceção para erros HTTP
        except requests.exceptions.RequestException as e:
            print(f"[Excel] ERRO: Falha na requisição ao obter ID do arquivo: {e}")
            return None

        file_id = response.json().get("id")
        if file_id:
            print(f"[Excel] ID do arquivo obtido e cacheado com sucesso: {file_id}")
            _file_id_cache = file_id
            return file_id
        else:
            print(f"[Excel] ERRO: Resposta da API não continha ID do arquivo. Resposta: {response.text}")
            return None

def update_cell(cell, value):
    """
    Atualiza uma célula na planilha.
    Usa lock para evitar operações concorrentes.
    """
    with _excel_write_lock:
        token = get_access_token()
        file_id = get_cached_file_id() # Usa cache

        if not token or not file_id:
            print(f"[Excel] ERRO: Não foi possível atualizar célula {cell}. Token ou file_id inválidos.")
            return False

        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }

        url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/drive/items/{file_id}/workbook/worksheets/{EXCEL_WORKSHEET_NAME}/range(address=\'{cell}\')"

        data = {
            "values": [[value]]
        }

        print(f"[Excel] Atualizando célula {cell} com valor: {value}")
        try:
            response = requests.patch(url, headers=headers, json=data, timeout=20) # Adicionado timeout
            response.raise_for_status()
        except requests.exceptions.RequestException as e:
            print(f"[Excel] ERRO: Falha na requisição ao atualizar célula {cell}: {e}")
            print(f"[Excel] Resposta (se disponível): {response.text if 'response' in locals() else 'N/A'}")
            return False

        print(f"[Excel] Célula {cell} atualizada com sucesso")
        return True

def get_cell_value(cell):
    """
    Obtém o valor de uma célula, tratando erros e convertendo para float.
    Retorna 0.0 em caso de erro ou valor não numérico.
    """
    token = get_access_token()
    file_id = get_cached_file_id() # Usa cache

    if not token or not file_id:
        print(f"[Excel] ERRO: Não foi possível ler célula {cell}. Token ou file_id inválidos.")
        return 0.0 # Retorna 0.0 em caso de erro de setup

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/drive/items/{file_id}/workbook/worksheets/{EXCEL_WORKSHEET_NAME}/range(address=\'{cell}\')"

    print(f"[Excel] Lendo valor da célula {cell}")
    try:
        response = requests.get(url, headers=headers, timeout=15) # Adicionado timeout
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"[Excel] ERRO: Falha na requisição ao ler célula {cell}: {e}")
        print(f"[Excel] Resposta (se disponível): {response.text if 'response' in locals() else 'N/A'}")
        return 0.0 # Retorna 0.0 em caso de erro de API

    values = response.json().get("values", [[None]])
    raw_value = values[0][0]
    print(f"[Excel] Valor bruto lido da célula {cell}: {repr(raw_value)}")

    # Tratar valor lido
    if raw_value is None:
        print(f"[Excel] Célula {cell} está vazia, retornando 0.0")
        return 0.0

    if isinstance(raw_value, (int, float)):
        print(f"[Excel] Valor numérico {raw_value} lido da célula {cell}")
        return float(raw_value)

    if isinstance(raw_value, str):
        # Handle specific Excel errors like #VALUE!, #N/A, etc.
        if raw_value.startswith("#"):
            print(f"[Excel] Aviso: Erro de fórmula \'{raw_value}\' lido da célula {cell}, retornando 0.0")
            return 0.0

        # Try converting string to float (handle potential commas, currency symbols)
        try:
            # Remove R$, spaces, thousands separators (.), then replace comma decimal separator
            cleaned_value = raw_value.replace("R$", "").strip().replace(".", "").replace(",", ".")
            numeric_value = float(cleaned_value)
            print(f"[Excel] Valor string \'{raw_value}\' convertido para float {numeric_value} da célula {cell}")
            return numeric_value
        except ValueError:
            print(f"[Excel] Aviso: Valor não numérico \'{raw_value}\' lido da célula {cell}, retornando 0.0")
            return 0.0

    # Fallback for unexpected types
    print(f"[Excel] Aviso: Tipo de valor inesperado ({type(raw_value)}) lido da célula {cell}, retornando 0.0")
    return 0.0

def get_range_values(cell_range):
    """
    Obtém valores de um intervalo de células.
    """
    token = get_access_token()
    file_id = get_cached_file_id() # Usa cache

    if not token or not file_id:
        print(f"[Excel] ERRO: Não foi possível ler intervalo {cell_range}. Token ou file_id inválidos.")
        return None

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/drive/items/{file_id}/workbook/worksheets/{EXCEL_WORKSHEET_NAME}/range(address=\'{cell_range}\')"

    print(f"[Excel] Lendo valores do intervalo {cell_range}")
    try:
        response = requests.get(url, headers=headers, timeout=15) # Adicionado timeout
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"[Excel] ERRO: Falha na requisição ao ler intervalo {cell_range}: {e}")
        print(f"[Excel] Resposta (se disponível): {response.text if 'response' in locals() else 'N/A'}")
        return None

    values = response.json().get("values", [])
    print(f"[Excel] {len(values)} linhas lidas do intervalo {cell_range}")
    return values

def find_next_empty_row(column, start_row, end_row):
    """
    Encontra o número da próxima linha vazia em uma coluna.
    """
    print(f"[Excel] Procurando próxima linha vazia na coluna {column} ({start_row}-{end_row})")
    range_values = get_range_values(f"{column}{start_row}:{column}{end_row}")

    if range_values is None: # Erro ao ler o intervalo
        print(f"[Excel] ERRO: Falha ao ler intervalo para encontrar linha vazia.")
        return None

    if not range_values: # Intervalo completamente vazio
        print(f"[Excel] Intervalo {column}{start_row}:{column}{end_row} vazio, retornando linha inicial {start_row}")
        return start_row

    for i, row_data in enumerate(range_values):
        # Considera a linha vazia se a célula na coluna estiver vazia ou for None
        if not row_data or row_data[0] is None or str(row_data[0]).strip() == "":
            next_row_num = start_row + i
            print(f"[Excel] Próxima linha vazia encontrada: {next_row_num}")
            return next_row_num

    # Se chegou aqui, todas as linhas no intervalo estão preenchidas
    print(f"[Excel] Nenhuma linha vazia encontrada no intervalo {column}{start_row}:{column}{end_row}. Verifique se o intervalo é suficiente.")
    return None

def write_operation(row_num, result):
    """
    Escreve o resultado (W/L) na linha especificada da coluna C.
    Assume que os valores de entrada (B, D, E) são calculados pela planilha.
    """
    print(f"[Excel] Escrevendo operação '{result}' na linha {row_num}")
    return update_cell(f"C{row_num}", result)

def clear_range(cell_range):
    """
    Limpa um intervalo de células.
    Usa lock para evitar operações concorrentes.
    """
    with _excel_write_lock:
        token = get_access_token()
        file_id = get_cached_file_id() # Usa cache

        if not token or not file_id:
            print(f"[Excel] ERRO: Não foi possível limpar intervalo {cell_range}. Token ou file_id inválidos.")
            return False

        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }

        url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/drive/items/{file_id}/workbook/worksheets/{EXCEL_WORKSHEET_NAME}/range(address=\'{cell_range}\')/clear"

        print(f"[Excel] Limpando intervalo {cell_range}")
        try:
            response = requests.post(url, headers=headers, timeout=20) # Adicionado timeout
            response.raise_for_status()
        except requests.exceptions.RequestException as e:
            print(f"[Excel] ERRO: Falha na requisição ao limpar intervalo {cell_range}: {e}")
            print(f"[Excel] Resposta (se disponível): {response.text if 'response' in locals() else 'N/A'}")
            return False

        print(f"[Excel] Intervalo {cell_range} limpo com sucesso")
        return True

def get_summary_data():
    """
    Obtém os dados resumidos necessários para atualizar a UI.
    Lê as células N25, N26, N29, N30.
    """
    print("[Excel] Obtendo dados resumidos (N25, N26, N29, N30)")
    # Ler em paralelo pode ser mais rápido, mas por simplicidade, lemos sequencialmente
    # Idealmente, usaríamos a API Batch do Graph, mas requer mais complexidade
    capital_atual = get_cell_value("N25")
    lucro_acumulado = get_cell_value("N26")
    acertos = get_cell_value("N29")
    erros = get_cell_value("N30")
    # Adicionar leitura de N16 e N17 se necessário para Valor Entrada e Lucro Operação
    valor_entrada = get_cell_value("N16")
    lucro_operacao = get_cell_value("N17")

    return {
        "capital_atual": capital_atual,
        "lucro_acumulado": lucro_acumulado,
        "acertos": acertos,
        "erros": erros,
        "valor_entrada": valor_entrada,
        "lucro_operacao": lucro_operacao
    }

def get_history_data(max_rows=100):
    """
    Obtém os dados do histórico das últimas operações.
    Lê as colunas B, C, D, E.
    """
    start_row = 3
    end_row = start_row + max_rows - 1
    print(f"[Excel] Obtendo dados do histórico (Linhas {start_row}-{end_row})")

    # Ler colunas necessárias em lote (idealmente seria uma única chamada batch)
    numeros = get_range_values(f"B{start_row}:B{end_row}")
    resultados = get_range_values(f"C{start_row}:C{end_row}")
    entradas = get_range_values(f"D{start_row}:D{end_row}")
    lucros = get_range_values(f"E{start_row}:E{end_row}")

    if numeros is None or resultados is None or entradas is None or lucros is None:
        print("[Excel] ERRO: Falha ao ler um ou mais intervalos do histórico")
        return [] # Retorna lista vazia em caso de erro

    # Formatar histórico
    historico = []
    num_items = min(len(numeros), len(resultados), len(entradas), len(lucros))

    for i in range(num_items):
        # Só adiciona ao histórico se o número da operação (coluna B) existir e não for vazio
        num_op = numeros[i][0] if numeros[i] else None
        if num_op is not None and str(num_op).strip() != "":
            historico.append({
                "numero": num_op,
                "valor": entradas[i][0] if entradas[i] else None,
                "resultado": resultados[i][0] if resultados[i] else None,
                "lucro": lucros[i][0] if lucros[i] else None
            })
        else:
            # Para de adicionar ao encontrar a primeira linha sem número de operação
            # Assume que o histórico é contíguo
            break

    print(f"[Excel] {len(historico)} itens de histórico formatados")
    return historico

def check_connection():
    """
    Verifica a conexão com a planilha tentando obter o file_id.
    """
    print("[Excel] Verificando conexão com a planilha...")

    if not EXCEL_WORKSHEET_NAME:
        print("[Excel] ERRO: EXCEL_WORKSHEET_NAME não está definido")
        return False

    if not USER_ID:
        print("[Excel] ERRO: USER_ID não está definido")
        return False

    file_id = get_cached_file_id()
    if not file_id:
        print("[Excel] ERRO: Não foi possível obter o ID do arquivo Excel")
        return False

    print("[Excel] Conexão com a planilha parece OK (ID do arquivo obtido)")
    return True

