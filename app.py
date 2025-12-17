from flask import Flask, request, jsonify, send_from_directory
from pathlib import Path
import os
from werkzeug.utils import secure_filename

# importa a função de processamento
from backend.aia import processar_arquivo_excel

app = Flask(__name__, static_folder='frontend', static_url_path='')

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / 'data'
DATA_DIR.mkdir(exist_ok=True)

@app.route('/')
def index():
    return send_from_directory(app.static_folder, 'index.html')

# Serve arquivos estáticos (css/js/img)
@app.route('/<path:filename>')
def static_files(filename):
    return send_from_directory(app.static_folder, filename)

@app.route('/api/processar', methods=['POST'])
def api_processar():
    try:
        if 'file' not in request.files:
            return jsonify({"success": False, "error": "Arquivo não enviado."}), 400
        f = request.files['file']
        if f.filename == '':
            return jsonify({"success": False, "error": "Arquivo sem nome."}), 400

        action = request.form.get('action', 'criar')
        company = request.form.get('company', '')
        batchSize = request.form.get('batchSize', '')

        # salva arquivo temporariamente em data/
        filename = secure_filename(f.filename)
        temp_path = DATA_DIR / filename
        f.save(str(temp_path))

        # determina pasta base de saída (opcional) fornecida pelo usuário
        output_base = request.form.get('outputBase', '').strip()
        if output_base:
            # sanitização básica: impedir caminhos absolutos e traversals
            if '..' in output_base or output_base.startswith(('/', '\\')):
                return jsonify({"success": False, "error": "Caminho de saída inválido."}), 400
            candidate = (BASE_DIR / output_base).resolve()
            base_resolved = BASE_DIR.resolve()
            if not str(candidate).startswith(str(base_resolved)):
                return jsonify({"success": False, "error": "Caminho de saída fora do diretório do projeto."}), 400
            pasta_base = str(candidate)
        else:
            pasta_base = str(BASE_DIR)

        # chama a função de processamento
        result = processar_arquivo_excel(str(temp_path), action, company, batchSize, pasta_base)

        # opcional: remover arquivo temporário
        try:
            os.remove(temp_path)
        except Exception:
            pass

        if result.get('success'):
            return jsonify(result)
        else:
            return jsonify(result), 500

    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
