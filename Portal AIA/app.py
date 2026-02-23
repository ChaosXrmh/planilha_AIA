from flask import Flask, request, jsonify, send_from_directory, send_file
from pathlib import Path
import os
from werkzeug.utils import secure_filename

# importa a função de processamento
from backend.aia import processar_arquivo_excel
import socket
import netifaces
import shutil
import tempfile

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
        output_format = request.form.get('output_format', 'planilha')
        # mapeamento explícito enviado pelo frontend (opcional)
        explicit_mapping = {
            'numero_col': request.form.get('numero_col', '') or None,
            'cnpj_col': request.form.get('cnpj_col', '') or None,
            'acao_col': request.form.get('acao_col', '') or None
        }
        # normaliza None se não preenchido
        for k in list(explicit_mapping.keys()):
            if not explicit_mapping[k]:
                explicit_mapping[k] = None

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
        result = processar_arquivo_excel(str(temp_path), action, company, batchSize, pasta_base, explicit_mapping, output_format=output_format)

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


@app.route('/api/download_zip', methods=['POST'])
def api_download_zip():
    """Compacta a pasta solicitada e retorna um ZIP para download.
    Recebe JSON: { "folder": "<absolute-or-relative-path>" }
    Por segurança, só permite pastas dentro do BASE_DIR.
    """
    try:
        data = request.get_json() or {}
        folder = data.get('folder') or request.form.get('folder')
        if not folder:
            return jsonify({"success": False, "error": "Parâmetro 'folder' é necessário."}), 400

        # resolve e valida caminho
        candidate = (Path(folder) if Path(folder).is_absolute() else (BASE_DIR / folder)).resolve()
        base_resolved = BASE_DIR.resolve()
        if not str(candidate).startswith(str(base_resolved)):
            return jsonify({"success": False, "error": "Caminho de pasta inválido ou fora do projeto."}), 400

        if not candidate.exists() or not candidate.is_dir():
            return jsonify({"success": False, "error": "Pasta não encontrada."}), 404

        # cria zip temporário
        tmpdir = tempfile.mkdtemp()
        zip_name = f"{candidate.name}.zip"
        zip_path = Path(tmpdir) / zip_name
        shutil.make_archive(str(zip_path.with_suffix('')), 'zip', root_dir=str(candidate))

        return send_file(str(zip_path), as_attachment=True, download_name=zip_name)
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


@app.route('/api/hostinfo', methods=['GET'])
def api_hostinfo():
    """Retorna IPs IPv4 do host para montar um link de rede."""
    try:
        ips = []
        # tenta usar netifaces para listar endereços de interfaces
        try:
            for iface in netifaces.interfaces():
                addrs = netifaces.ifaddresses(iface)
                ipv4 = addrs.get(netifaces.AF_INET, [])
                for a in ipv4:
                    ip = a.get('addr')
                    if ip and not ip.startswith('127.') and ':' not in ip:
                        ips.append(ip)
        except Exception:
            # fallback: tenta resolver hostname
            try:
                hostname = socket.gethostname()
                resolved = socket.gethostbyname_ex(hostname)[2]
                for ip in resolved:
                    if ip and not ip.startswith('127.') and ':' not in ip:
                        ips.append(ip)
            except Exception:
                pass

        # dedupe
        ips = list(dict.fromkeys(ips))
        return jsonify({"success": True, "ips": ips, "port": 5000})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
