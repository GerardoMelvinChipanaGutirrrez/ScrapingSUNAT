from flask import Flask, render_template, request, jsonify
from werkzeug.utils import secure_filename
import os
import pandas as pd
from .scraper.sunat import SunatScraper

app = Flask(__name__)
app.config.from_object('app.config.Config')

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
        
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Procesar archivo
        try:
            df = pd.read_excel(filepath)
            scraper = SunatScraper()
            results = []
            
            for razon in df['razon_social']:
                data = scraper.buscar_en_sunat(nombre=razon)
                if data:
                    ruc, nombre, reps = data
                    results.append({
                        'razon': razon,
                        'ruc': ruc,
                        'nombre': nombre,
                        'representantes': reps
                    })
                    
            return jsonify({'results': results})
            
        except Exception as e:
            return jsonify({'error': str(e)}), 500
            
    return jsonify({'error': 'File type not allowed'}), 400