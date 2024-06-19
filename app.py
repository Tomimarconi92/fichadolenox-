from flask import Flask, render_template, request, redirect, url_for
import openpyxl
from datetime import datetime

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('formulario.html')

@app.route('/submit', methods=['POST'])
def submit():
    if request.method == 'POST':
        nombre = request.form['nombre']
        apellido = request.form['apellido']
        documento = request.form['documento']
        legajo = request.form['legajo']
        sector = request.form['sector']
        if sector == '1':
            sector = 'Administracion'
        elif sector == '2':
            sector = "Cabina Operativa"
        elif sector == '3':
            sector = "Coordinacion"
        elif sector == '4':
            sector = 'otros'
        puesto = request.form['puesto']
        horario_inicio_guardia = request.form['horario_inicio_guardia']
        horario_fin_guardia = request.form['horario_fin_guardia']
        dias_laborables = request.form['dias_laborables']
        if dias_laborables == '1':
            dias_laborables = 'Dias de semana'
        elif dias_laborables == '2':
            dias_laborables = 'SaDoFe'

        direccion = request.form['direccion']
        localidad = request.form['localidad']
        provincia = request.form['provincia']
        codigo_postal = request.form['codigo_postal']
        fecha_nacimiento = request.form['fecha_nacimiento']
        email = request.form['email']
        cuil = request.form['cuil']
        numero_celular = request.form['numero_celular']
        observaciones = request.form['observaciones']

        # Guardar datos en un archivo Excel
        guardar_datos_en_excel(nombre, 
        apellido,
        documento,
        legajo,
        sector,
        puesto,
        horario_inicio_guardia,
        horario_fin_guardia,
        dias_laborables,
        direccion,
        localidad,
        provincia,
        codigo_postal,
        fecha_nacimiento,
        email,
        cuil,
        numero_celular,
        observaciones)

        # Redirigir a la página de agradecimiento
        return redirect(url_for('agradecimiento'))

@app.route('/agradecimiento')
def agradecimiento():
    return render_template('agradecimiento.html')

def guardar_datos_en_excel(nombre, 
        apellido,
        documento,
        legajo,
        sector,
        puesto,
        horario_inicio_guardia,
        horario_fin_guardia,
        dias_laborables,
        direccion,
        localidad,
        provincia,
        codigo_postal,
        fecha_nacimiento,
        email,
        cuil,
        numero_celular,
        observaciones):
    # Abrir el archivo Excel existente o crear uno nuevo
    try:
        wb = openpyxl.load_workbook('datos.xlsx')
        sheet = wb.active
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        sheet = wb.active
        # Agregar encabezados si es un nuevo archivo
        sheet.append(['Nombre', 
        'Apellido',
        'Documento',
        'Legajo',
        'Sector',
        'Puesto',
        'Horario inicio guardia',
        'Horario Fin Guardia',
        'Dias Laborables',
        'Direccion',
        'Localidad',
        'Provincia',
        'Codigo_postal',
        'Fecha Nacimiento',
        'Email',
        'Cuil',
        'Numero celular',
        'Observaciones'])

    # Agregar nueva fila con datos del formulario
    sheet.append([nombre, 
        apellido,
        documento,
        legajo,
        sector,
        puesto,
        horario_inicio_guardia,
        horario_fin_guardia,
        dias_laborables,
        direccion,
        localidad,
        provincia,
        codigo_postal,
        fecha_nacimiento,
        email,
        cuil,
        numero_celular,
        observaciones])

    # Guardar cambios en el archivo Excel
    wb.save('datos.xlsx')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
   