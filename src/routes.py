import os
from flask import Flask, request, jsonify
from flask_cors import CORS
from dotenv import load_dotenv
from .excel import update_cell, get_cell_value, get_range_values, find_next_empty_cell, clear_range, check_connection

# Carregar variáveis de ambiente
load_dotenv()

def create_app():
    app = Flask(__name__)
    CORS(app)
    
    # Endpoint para atualizar células da planilha
    @app.route('/update', methods=['POST'])
    def update():
        data = request.json
        
        # Mapeamento de campos para células
        cell_mapping = {
            'capital_inicial': 'N12',
            'total_operacoes': 'N13',
            'operacoes_ganho': 'N14',
            'payout_fixo': 'N15'
        }
        
        success = True
        for field, cell in cell_mapping.items():
            value = data.get(field)
            if value is not None:
                if not update_cell(cell, value):
                    success = False
        
        if success:
            return jsonify({"status": "success", "message": "Células atualizadas com sucesso"}), 200
        else:
            return jsonify({"status": "error", "message": "Erro ao atualizar células"}), 500

    # Endpoint para registrar vitória (WIN)
    @app.route('/win', methods=['POST'])
    def win():
        next_row = find_next_empty_cell('C', 3, 102)
        
        if next_row is None:
            return jsonify({"status": "error", "message": "Não há células vazias disponíveis"}), 400
        
        if update_cell(f"C{next_row}", "W"):
            return jsonify({"status": "success", "message": f"Vitória registrada na célula C{next_row}"}), 200
        else:
            return jsonify({"status": "error", "message": "Erro ao registrar vitória"}), 500

    # Endpoint para registrar derrota (LOSS)
    @app.route('/loss', methods=['POST'])
    def loss():
        next_row = find_next_empty_cell('C', 3, 102)
        
        if next_row is None:
            return jsonify({"status": "error", "message": "Não há células vazias disponíveis"}), 400
        
        if update_cell(f"C{next_row}", "L"):
            return jsonify({"status": "success", "message": f"Derrota registrada na célula C{next_row}"}), 200
        else:
            return jsonify({"status": "error", "message": "Erro ao registrar derrota"}), 500

    # Endpoint para zerar (limpar células)
    @app.route('/reset', methods=['POST'])
    def reset():
        # Limpar resultados (W/L)
        if not clear_range("C3:C102"):
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
            if not update_cell(cell, value):
                success = False
        
        if success:
            return jsonify({"status": "success", "message": "Dados zerados com sucesso"}), 200
        else:
            return jsonify({"status": "error", "message": "Erro ao zerar dados"}), 500

    # Endpoint para obter dados da planilha
    @app.route('/dados', methods=['GET'])
    def get_data():
        try:
            # Obter valores das células individuais
            capital_atual = get_cell_value("N25")
            lucro_acumulado = get_cell_value("N26")
            acertos = get_cell_value("N29")
            erros = get_cell_value("N30")
            
            # Obter valores dos intervalos
            entradas = get_range_values("D3:D102")
            lucros = get_range_values("E3:E102")
            numeros = get_range_values("B3:B102")
            resultados = get_range_values("C3:C102")
            
            # Formatar histórico
            historico = []
            for i in range(min(len(numeros), len(entradas), len(resultados), len(lucros))):
                if numeros[i][0] is not None:  # Se o número da operação existe
                    historico.append({
                        "numero": numeros[i][0],
                        "valor": entradas[i][0],
                        "resultado": resultados[i][0],
                        "lucro": lucros[i][0]
                    })
            
            # Formatar resposta
            response = {
                "capital_atual": capital_atual,
                "lucro_acumulado": lucro_acumulado,
                "acertos": acertos,
                "erros": erros,
                "entradas": [row[0] for row in entradas if row and row[0] is not None],
                "lucros": [row[0] for row in lucros if row and row[0] is not None],
                "historico": historico
            }
            
            return jsonify(response), 200
        
        except Exception as e:
            print(f"Erro ao obter dados: {str(e)}")
            return jsonify({"status": "error", "message": f"Erro ao obter dados: {str(e)}"}), 500

    # Endpoint de status para monitoramento
    @app.route('/status', methods=['GET'])
    def status():
        if check_connection():
            return jsonify({"status": "online", "message": "Conexão com a planilha estabelecida com sucesso"}), 200
        else:
            return jsonify({"status": "offline", "message": "Erro ao conectar com a planilha"}), 503

    # Rota de teste para verificar se a API está funcionando
    @app.route('/test', methods=['GET'])
    def test():
        return jsonify({"status": "success", "message": "API funcionando corretamente"}), 200
    
    return app
