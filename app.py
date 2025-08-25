import os
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
import processor

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'

# Crear el directorio de subidas si no existe
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate_report():
    try:
        # 1. Recoger datos del formulario
        pais = request.form['pais']
        anyo = request.form['anyo']
        mes = request.form['mes']
        dia = request.form['dia']
        idioma = request.form['idioma']
        segmentos = request.form['segmentos']
        segmentation = request.form['segmentation']

        # 2. Guardar archivos subidos
        modelos_file = request.files['libro_modelos']
        marcas_file = request.files['libro_marcas']

        modelos_filename = secure_filename(modelos_file.filename)
        marcas_filename = secure_filename(marcas_file.filename)

        modelos_path = os.path.join(app.config['UPLOAD_FOLDER'], modelos_filename)
        marcas_path = os.path.join(app.config['UPLOAD_FOLDER'], marcas_filename)

        modelos_file.save(modelos_path)
        marcas_file.save(marcas_path)

        # 3. Llamar a la lógica de procesamiento
        output_path = processor.generate_report(
            pais, anyo, mes, dia, idioma, segmentos, segmentation,
            modelos_path, marcas_path
        )

        # 4. Enviar el archivo generado al usuario
        return send_file(output_path, as_attachment=True)

    except Exception as e:
        # Manejo básico de errores
        return str(e), 500
    finally:
        # 5. Limpieza (opcional, podrías querer un cron job para esto)
        if 'modelos_path' in locals() and os.path.exists(modelos_path):
            os.remove(modelos_path)
        if 'marcas_path' in locals() and os.path.exists(marcas_path):
            os.remove(marcas_path)
        # El archivo de salida (output_path) también debería limpiarse después de ser enviado.
        # send_file puede manejar esto con un wrapper si es necesario.

if __name__ == '__main__':
    app.run(debug=True)
