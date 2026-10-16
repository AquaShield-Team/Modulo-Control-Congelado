import os
import sys
import shutil
import subprocess
import threading
from flask import Flask, jsonify, request, send_from_directory

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__, static_folder=BASE_DIR)
RULES_LOCK = threading.Lock()

ALLOWED_STATIC_EXTS = {'.html', '.js', '.css', '.png', '.jpg', '.jpeg', '.ico', '.svg', '.json', '.woff', '.woff2', '.ttf'}

@app.after_request
def add_cors(response):
    origin = request.headers.get('Origin', '')
    allowed_origins = ['http://localhost:8765', 'http://127.0.0.1:8765', 'null']
    if origin in allowed_origins or not origin:
        response.headers['Access-Control-Allow-Origin'] = origin if origin else '*'
    else:
        response.headers['Access-Control-Allow-Origin'] = 'http://localhost:8765'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type,Authorization'
    response.headers['Access-Control-Allow-Methods'] = 'GET,POST,OPTIONS'
    return response

@app.route('/')
def index():
    return send_from_directory(BASE_DIR, 'index.html')

@app.route('/<path:filename>')
def static_files(filename):
    ext = os.path.splitext(filename)[1].lower()
    if ext not in ALLOWED_STATIC_EXTS:
        return jsonify({'error': 'Acceso no permitido al archivo'}), 403
    return send_from_directory(BASE_DIR, filename)

@app.route('/health')
@app.route('/api/health')
def health():
    return jsonify({'status': 'ok', 'service': 'AquaShield Synapse Sync'})

@app.route('/api/upload', methods=['POST', 'OPTIONS'])
def upload():
    if request.method == 'OPTIONS':
        return jsonify({'status': 'ok'})
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'No se recibio ningun archivo'}), 400
        
        file = request.files['file']
        filename = (file.filename or '').strip()
        file_type = request.form.get('type', 'auto').lower()
        
        # Reconocimiento inteligente de archivo
        if file_type == 'vf05n' or 'vf05' in filename.lower() or 'factura' in filename.lower():
            target_name = 'vf05n.xlsx'
        elif 'control congelado' in filename.lower() or 'congelado' in filename.lower():
            target_name = 'Control Congelado Comex 2026.xlsx'
        elif 'flujo' in filename.lower():
            target_name = 'flujo v2.xlsx'
        elif file_type == 'cartera' or 'cartera' in filename.lower() or 'asignacion' in filename.lower():
            target_name = 'Asignacion Clientes.xlsx'
        elif file_type == 'exportaciones' or 'exportacion' in filename.lower() or 'gestion pedidos' in filename.lower() or 'pedidos de vta' in filename.lower():
            target_name = 'Exportaciones gestion Pedidos de Vta.xlsx'
        else:
            safe_name = os.path.basename(filename).strip()
            ext = os.path.splitext(safe_name)[1].lower()
            if ext not in ('.xlsx', '.xls'):
                return jsonify({'success': False, 'error': 'Solo se permiten archivos Excel (.xlsx, .xls)'}), 400
            target_name = safe_name if safe_name else 'archivo_subido.xlsx'
        
        dest_path = os.path.abspath(os.path.join(BASE_DIR, target_name))
        if not dest_path.startswith(BASE_DIR):
            return jsonify({'success': False, 'error': 'Ruta de destino inválida'}), 400

        # Protección crítica: si es Control Congelado Comex 2026.xlsx, verificar que NO sea un archivo parcial/filtrado
        if target_name == 'Control Congelado Comex 2026.xlsx' and os.path.exists(dest_path):
            tmp_upload = os.path.join(BASE_DIR, 'temp_upload_check.xlsx')
            file.save(tmp_upload)
            try:
                import openpyxl
                wb_new = openpyxl.load_workbook(tmp_upload, read_only=True)
                new_rows = wb_new.active.max_row or 0
                wb_new.close()
                wb_old = openpyxl.load_workbook(dest_path, read_only=True)
                old_rows = wb_old.active.max_row or 0
                wb_old.close()
                if old_rows > 300 and new_rows < (old_rows * 0.6):
                    os.remove(tmp_upload)
                    print(f"-> [PROTECCION] Archivo parcial detectado ({new_rows} filas vs {old_rows} de la base maestra). Se mantienen las {old_rows} filas maestras y se sincronizan cambios vía Firebase/cambios_manuales.json.")
                    cmd = [sys.executable, os.path.join(BASE_DIR, 'actualizar_datos.py')]
                    r = subprocess.run(cmd, capture_output=True, text=True, cwd=BASE_DIR, timeout=120)
                    return jsonify({
                        'success': True,
                        'message': f'Planilla parcial ({new_rows} filas) aplicada como actualización sin borrar la base maestra ({old_rows} filas).',
                        'target_file': target_name,
                        'log': r.stdout[-300:] if r.stdout else r.stderr[-300:]
                    }), 200
            except Exception as e_chk:
                print(f"-> Aviso al verificar filas de archivo subido: {e_chk}")
            shutil.copyfile(dest_path, os.path.join(BASE_DIR, f"{os.path.splitext(target_name)[0]}_BACKUP.xlsx"))
            os.replace(tmp_upload, dest_path)
            print(f"-> Archivo maestro guardado con exito como: {target_name}")
        else:
            if os.path.exists(dest_path):
                backup_path = os.path.join(BASE_DIR, f"{os.path.splitext(target_name)[0]}_BACKUP.xlsx")
                try:
                    shutil.copyfile(dest_path, backup_path)
                except Exception:
                    pass
            file.save(dest_path)
            print(f"-> Archivo guardado con exito como: {target_name}")

        # Ejecutar consolidación automática
        cmd = [sys.executable, os.path.join(BASE_DIR, 'actualizar_datos.py')]
        r = subprocess.run(cmd, capture_output=True, text=True, cwd=BASE_DIR, timeout=120)

        success = (r.returncode == 0)
        return jsonify({
            'success': success,
            'message': f'Archivo guardado como {target_name} y datos triangulados con exito.' if success else 'Error al procesar archivo en consolidacion',
            'target_file': target_name,
            'log': r.stdout[-300:] if r.stdout else r.stderr[-300:]
        }), (200 if success else 500)
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/sync', methods=['POST', 'OPTIONS'])
def sync():
    if request.method == 'OPTIONS':
        return jsonify({'status': 'ok'})

    print("-> Sincronizacion solicitada desde la interfaz web...")
    try:
        # 1. Extraer pedidos en vivo desde Azure Synapse Data Lake
        cmd1 = [sys.executable, os.path.join(BASE_DIR, 'extraer_synapse.py')]
        r1 = subprocess.run(cmd1, capture_output=True, text=True, cwd=BASE_DIR, timeout=120)

        # 2. Triangular con la planilla, carteras y VF05N
        cmd2 = [sys.executable, os.path.join(BASE_DIR, 'actualizar_datos.py')]
        r2 = subprocess.run(cmd2, capture_output=True, text=True, cwd=BASE_DIR, timeout=120)

        success = (r1.returncode == 0 and r2.returncode == 0)
        return jsonify({
            'success': success,
            'message': 'Sincronizacion completada con exito' if success else 'Error al procesar datos',
            'synapse_log': r1.stdout[-300:] if r1.stdout else r1.stderr[-300:],
            'actualizar_log': r2.stdout[-300:] if r2.stdout else r2.stderr[-300:]
        }), (200 if success else 500)
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/save_rule', methods=['POST', 'OPTIONS'])
def save_rule():
    if request.method == 'OPTIONS':
        return jsonify({'status': 'ok'})
    try:
        import json
        data = request.get_json() or {}
        c_name = data.get('cliente_patron', '').strip().upper()
        if not c_name:
            return jsonify({'success': False, 'error': 'Nombre de cliente requerido'}), 400
        
        rules_path = os.path.join(BASE_DIR, 'reglas_clientes.json')
        with RULES_LOCK:
            current_data = {'_meta': {}, 'reglas': []}
            if os.path.exists(rules_path):
                try:
                    with open(rules_path, 'r', encoding='utf-8') as f:
                        current_data = json.load(f)
                except Exception as e_read:
                    return jsonify({'success': False, 'error': f'Error al leer archivo de reglas: {e_read}'}), 500
            
            reglas = current_data.get('reglas', [])
            found = False
            for i, r in enumerate(reglas):
                if r.get('cliente_patron', '').upper() == c_name or (data.get('codigo_sap') and r.get('codigo_sap') == data.get('codigo_sap')):
                    reglas[i] = data
                    found = True
                    break
            if not found:
                reglas.append(data)
            current_data['reglas'] = reglas
            
            tmp_rules = rules_path + '.tmp'
            with open(tmp_rules, 'w', encoding='utf-8') as f:
                json.dump(current_data, f, indent=2, ensure_ascii=False)
            os.replace(tmp_rules, rules_path)
            
        print(f"-> Regla de cliente actualizada en reglas_clientes.json: {c_name}")
        return jsonify({'success': True, 'message': f'Regla guardada para {c_name}'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

if __name__ == '__main__':
    print("==================================================================")
    print("   AQUASHIELD - SERVIDOR DE SINCRONIZACION SYNAPSE EN VIVO")
    print("   Disponible en: http://localhost:8765")
    print("==================================================================")
    app.run(host='127.0.0.1', port=8765, debug=False)
