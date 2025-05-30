import os
from flask import Flask, request, jsonify
from flask_cors import CORS
from dotenv import load_dotenv
from .excel import (
    update_cell, 
    get_cell_value, 
    get_range_values, 
    find_next_empty_row, 
    clear_range, 
    check_connection,
    get_summary_data,
    get_history_data,
    write_operation
)

# Carregar variáveis de ambiente
load_dotenv()

def create_app():
    app = Flask(__name__)
    CORS(app)
    
    # Endpoint para atualizar células da planilha
    @app.route('/update', methods=['POST'])
    def update():
        try:
            data = request.json
            print(f"[API] Recebido pedido de atualização: {data}")
            
            # Mapeamento de campos para células com suporte a múltiplos formatos
            cell_mapping = {
                'N12': ['capital_inicial'],
                'N13': ['total_operacoes'],
                'N14': ['operacoes_ganho', 'operacoes_com_ganho'],
                'N15': ['payout_fixo', 'payout']
            }
            
            success = True
            cells_updated = []
            
            # Processar cada célula e tentar todos os possíveis nomes de campo
            for cell, field_names in cell_mapping.items():
                updated = False
                
                # Tentar cada possível nome de campo
                for field_name in field_names:
                    value = data.get(field_name)
                    if value is not None:
                        print(f"[API] Atualizando {field_name} ({cell}) com valor: {value}")
                        if update_cell(cell, value):
                            cells_updated.append(cell)
                            updated = True
                            break  # Campo encontrado e atualizado, não precisa tentar outros nomes
                
                # Se nenhum dos possíveis nomes de campo foi encontrado ou atualizado
                if not updated:
                    print(f"[API] Aviso: Nenhum valor fornecido para célula {cell} (campos possíveis: {field_names})")
            
            # Após atualizar as células, obter os dados atualizados para retornar ao frontend
            summary_data = get_summary_data()
            
            if cells_updated:
                print(f"[API] Células atualizadas com sucesso: {cells_updated}")
                return jsonify({
                    "status": "success", 
                    "message": "Células atualizadas com sucesso", 
                    "cells_updated": cells_updated,
                    **summary_data  # Incluir dados atualizados na resposta
                }), 200
            else:
                print("[API] Nenhuma célula foi atualizada")
                return jsonify({
                    "status": "warning", 
                    "message": "Nenhuma célula foi atualizada",
                    **summary_data  # Incluir dados atualizados na resposta
                }), 200
        except Exception as e:
            print(f"[API] Exceção ao processar /update: {str(e)}")
            return jsonify({"status": "error", "message": f"Erro ao atualizar células: {str(e)}"}), 500

    # Endpoint para registrar vitória (WIN)
    @app.route('/win', methods=['POST'])
    def win():
        try:
            print("[API] Recebido pedido para registrar vitória")
            next_row = find_next_empty_row('C', 3, 102)
            
            if next_row is None:
                print("[API] Não há células vazias disponíveis para registrar vitória")
                return jsonify({"status": "error", "message": "Não há células vazias disponíveis"}), 400
            
            print(f"[API] Registrando vitória na célula C{next_row}")
            if write_operation(next_row, "W"):
                print(f"[API] Vitória registrada com sucesso na célula C{next_row}")
                
                # Obter dados atualizados após registrar a vitória
                summary_data = get_summary_data()
                history_data = get_history_data(10)  # Limitar a 10 itens mais recentes
                
                # Retornar todos os dados necessários para atualizar o frontend
                return jsonify({
                    "status": "success", 
                    "message": f"Vitória registrada na célula C{next_row}",
                    **summary_data,
                    "historico": history_data
                }), 200
            else:
                print("[API] Erro ao registrar vitória")
                return jsonify({"status": "error", "message": "Erro ao registrar vitória"}), 500
        except Exception as e:
            print(f"[API] Exceção ao processar /win: {str(e)}")
            return jsonify({"status": "error", "message": f"Erro ao registrar vitória: {str(e)}"}), 500

    # Endpoint para registrar derrota (LOSS)
    @app.route('/loss', methods=['POST'])
    def loss():
        try:
            print("[API] Recebido pedido para registrar derrota")
            next_row = find_next_empty_row('C', 3, 102)
            
            if next_row is None:
                print("[API] Não há células vazias disponíveis para registrar derrota")
                return jsonify({"status": "error", "message": "Não há células vazias disponíveis"}), 400
            
            print(f"[API] Registrando derrota na célula C{next_row}")
            if write_operation(next_row, "L"):
                print(f"[API] Derrota registrada com sucesso na célula C{next_row}")
                
                # Obter dados atualizados após registrar a derrota
                summary_data = get_summary_data()
                history_data = get_history_data(10)  # Limitar a 10 itens mais recentes
                
                # Retornar todos os dados necessários para atualizar o frontend
                return jsonify({
                    "status": "success", 
                    "message": f"Derrota registrada na célula C{next_row}",
                    **summary_data,
                    "historico": history_data
                }), 200
            else:
                print("[API] Erro ao registrar derrota")
                return jsonify({"status": "error", "message": "Erro ao registrar derrota"}), 500
        except Exception as e:
            print(f"[API] Exceção ao processar /loss: {str(e)}")
            return jsonify({"status": "error", "message": f"Erro ao registrar derrota: {str(e)}"}), 500

    # Endpoint para zerar (limpar células)
    @app.route('/reset', methods=['POST'])
    def reset():
        try:
            print("[API] Recebido pedido para zerar dados")
            # Limpar resultados (W/L)
            print("[API] Limpando resultados (C3:C102)")
            if not clear_range("C3:C102"):
                print("[API] Erro ao limpar resultados")
                return jsonify({"status": "error", "message": "Erro ao limpar resultados"}), 500
            
            # Limpar células de entrada
            cell_mapping = {
                'N12': '',
                'N13': '',
                'N14': '',
                'N15': ''
            }
            
            success = True
            for cell, value in cell_mapping.items():
                print(f"[API] Limpando célula {cell}")
                if not update_cell(cell, value):
                    success = False
                    print(f"[API] Erro ao limpar célula {cell}")
            
            if success:
                print("[API] Dados zerados com sucesso")
                
                # Obter dados atualizados após zerar
                summary_data = get_summary_data()
                
                # Retornar todos os dados necessários para atualizar o frontend
                return jsonify({
                    "status": "success", 
                    "message": "Dados zerados com sucesso",
                    **summary_data,
                    "historico": []  # Histórico vazio após zerar
                }), 200
            else:
                print("[API] Erro ao zerar dados")
                return jsonify({"status": "error", "message": "Erro ao zerar dados"}), 500
        except Exception as e:
            print(f"[API] Exceção ao processar /reset: {str(e)}")
            return jsonify({"status": "error", "message": f"Erro ao zerar dados: {str(e)}"}), 500

    # Endpoint para obter dados da planilha (usado apenas no carregamento inicial)
    @app.route('/dados', methods=['GET'])
    def get_data():
        try:
            print("[API] Recebido pedido para obter dados iniciais")
            
            # Obter dados resumidos
            summary_data = get_summary_data()
            
            # Obter histórico
            history_data = get_history_data()
            
            # Formatar resposta
            response = {
                **summary_data,
                "historico": history_data
            }
            
            print("[API] Dados iniciais obtidos com sucesso")
            return jsonify(response), 200
        
        except Exception as e:
            print(f"[API] Exceção ao processar /dados: {str(e)}")
            return jsonify({"status": "error", "message": f"Erro ao obter dados: {str(e)}"}), 500

    # Endpoint de status para monitoramento
    @app.route('/status', methods=['GET'])
    def status():
        try:
            print("[API] Verificando status da conexão com a planilha")
            if check_connection():
                print("[API] Conexão com a planilha estabelecida com sucesso")
                return jsonify({"status": "online", "message": "Conexão com a planilha estabelecida com sucesso"}), 200
            else:
                print("[API] Erro ao conectar com a planilha")
                return jsonify({"status": "offline", "message": "Erro ao conectar com a planilha"}), 503
        except Exception as e:
            print(f"[API] Exceção ao verificar status: {str(e)}")
            return jsonify({"status": "error", "message": f"Erro ao verificar status: {str(e)}"}), 500

    # Rota de teste para verificar se a API está funcionando
    @app.route('/test', methods=['GET'])
    def test():
        print("[API] Teste de funcionamento da API")
        return jsonify({"status": "success", "message": "API funcionando corretamente"}), 200
    
    return app
