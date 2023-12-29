import os
from flask import Flask, request, send_file
from flask.templating import render_template
from werkzeug.utils import secure_filename
import plan_ecus

UPLOAD_FOLDER = 'static'

app = Flask(__name__)
app.config['SECRET_KEY'] = 'secret'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

archivo_excel, archivo_solucion = '', ''

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    global archivo_excel
    if request.method == 'POST':
        if archivo_excel == '':
            if 'archivo_excel' in request.files:
                datos = request.files["archivo_excel"]
                # Guardar archivo de datos
                archivo_excel = secure_filename(datos.filename)
                datos.save(os.path.join(app.config['UPLOAD_FOLDER'], archivo_excel))

    return render_template("index.html", archivo_excel=archivo_excel)

@app.route('/descarga_plan', methods=['GET', 'POST'])
def desargar_archivo():
    global archivo_excel
    if request.method == 'POST':
        archivo_datos = os.path.join(app.config['UPLOAD_FOLDER'], archivo_excel)
        archivo_solucion = os.path.join(app.config['UPLOAD_FOLDER'], 'solucion.xlsx')
        plan_ecus.crea_plan(archivo_datos, archivo_solucion)
        archivo_excel = ''
        return send_file(os.path.join(app.config['UPLOAD_FOLDER'], 'solucion.xlsx'), mimetype='text/csv', as_attachment=True)

@app.route('/borrarArchivos', methods = ['GET', 'POST'])
def borar_archivos():
    for document in os.listdir('./static/'):
            if not document.startswith('style'):
                os.remove('./static/' + document)
    return 'Archivos de texto borrados correctamente'

if __name__ == '__main__':
    app.run(debug = True)