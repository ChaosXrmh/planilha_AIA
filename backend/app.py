# ARQUIVO: app.py
from flask import Flask, request, jsonify
import os
from pathlib import Path
# Importamos a função que criamos no passo anterior
from backend_logic import processar_arquivo_excel

app = Flask(__name__, static_folder='.', static_url_path='')

# Configuração de onde salvar arquivos temporários
UPLOAD_FOLDER = Path('temp_uploads')
UPLOAD_FOLDER.mkdir(exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Rota para servir o seu HTML principal
@app.route('/')
def index():
    # Garanta que seu arquivo HTML se chame 'index.html'
    return app.send_static_file('index.html')

# Rota da API que vai receber os dados do formulário
@app.route('/api/processar', methods=['POST'])
def api_processar():
    try:
        # 1. Verificar se o arquivo foi enviado
        if 'file' not in request.files:
            return jsonify({"success": False, "error": "Nenhum arquivo enviado"}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({"success": False, "error": "Nome de arquivo inválido"}), 400

        # 2. Pegar os dados do formulário (Action, Company, Batch Size)
        acao = request.form.get('action')
        empresa = request.form.get('company')
        # Tenta converter batchSize para inteiro, usa 100 como fallback
        try:
            tamanho_lote = int(request.form.get('batchSize', 100))
        except ValueError:
             tamanho_lote = 100

        # 3. Salvar o arquivo Excel temporariamente
        caminho_temp_arquivo = app.config['UPLOAD_FOLDER'] / file.filename
        file.save(caminho_temp_arquivo)

        # 4. CHAMAR A SUA COZINHA (seu script de backend adaptado)
        # Passamos os dados que vieram do front e onde queremos que salve
        resultado = processar_arquivo_excel(
            caminho_arquivo_entrada=caminho_temp_arquivo,
            acao=acao,
            empresa_raw=empresa,
            tamanho_lote=tamanho_lote,
            pasta_base_saida=Path('.') # Salva na raiz do projeto
        )

        # 5. Limpar o arquivo temporário (opcional, mas boa prática)
        try:
            os.remove(caminho_temp_arquivo)
        except:
            pass

        # 6. Retornar o resultado para o Front-end como JSON
        if resultado['success']:
            return jsonify(resultado)
        else:
            # Se sua função retornou success: False, envia erro 500
            return jsonify(resultado), 500

    except Exception as e:
        # Erro genérico no servidor
        return jsonify({"success": False, "error": str(e)}), 500

if __name__ == '__main__':
    # Roda o servidor na porta 5000
    app.run(debug=True, port=5000)