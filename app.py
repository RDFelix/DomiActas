from flask import Flask, render_template, request, redirect, url_for, send_file
import os
from convert_pdf import process_document_and_convert_to_pdf
from database import create_database, insert_acta, insert_descuento, get_all_actas, get_all_descuentos, delete_acta, delete_descuento
from export_excel import generar_excel
from datetime import datetime

app = Flask(__name__)
# app.secret_key = 'tu_clave_secreta_aqui' # Descomentar si usas flash messages

@app.route('/')
def index():
    actas = get_all_actas()
    return render_template('inicio.html', actas=actas)

@app.route('/inicio')
def inicio():
    actas = get_all_actas()
    return render_template('inicio.html', actas=actas)


@app.route('/eliminar', methods=['POST'])
def eliminar():
    if request.method == 'POST':
        codigo = request.form.get('codigo')
        if codigo:
            print(f"Attempting to delete acta with ID: {codigo}")
            delete_acta(codigo)
        else:
            print("No 'codigo' provided in the form data for deletion.")
    return redirect(url_for('inicio'))

@app.route('/asignacion')
def asignacion():
    return render_template('asignacion.html')

@app.route('/devolucion')
def devolucion():
    return render_template('devolucion.html')

@app.route('/descuento')
def descuento():
    return render_template('descuento.html')

@app.route('/mantenimiento')
def mantenimiento():
    return render_template('mantenimiento.html')

@app.route('/generar_seguimiento')
def generar_seguimiento():
    excel_data = generar_excel()
    if excel_data:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"SEGUIMIENTO_DE_ACTAS_{timestamp}.xlsx"
        return send_file(
            excel_data,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=output_filename
        )
    else:
        return "Error al generar el archivo excel", 500


@app.route('/generar_asignar', methods=['POST'])
def generar_asignar():
    if request.method == 'POST':
        ciudad = request.form.get('ciudad')
        dia = request.form.get('dia')
        mes = request.form.get('mes')
        año = request.form.get('año')
        nombre_completo = request.form.get('nombre_completo', '').upper()
        cedula = request.form.get('cedula')
        cesantias = request.form.get('cesantias')
        
        marca_celular = request.form.get('marca_celular') or "N/A"
        modelo_celular = request.form.get('modelo_celular') or "N/A"
        serial_celular = request.form.get('serial_celular') or "N/A"
        imei_celular = request.form.get('imei_celular') or "N/A"
        cargador_celular = request.form.get('cargador_celular') or "N/A"
        linea_celular = request.form.get('linea_celular') or "N/A"

        marca_portatil = request.form.get('marca_portatil') or "N/A"
        modelo_portatil = request.form.get('modelo_portatil') or "N/A"
        serial_portatil = request.form.get('serial_portatil') or "N/A"
        cargador_portatil = request.form.get('cargador_portatil') or "N/A"

        teclado_accesorio = request.form.get('teclado_accesorio') or "N/A"
        mouse_accesorio = request.form.get('mouse_accesorio') or "N/A"
        base_accesorio = request.form.get('base_accesorio') or "N/A"
        diadema_accesorio = request.form.get('diadema_accesorio') or "N/A"

        marca_monitor = request.form.get('marca_monitor') or "N/A"
        modelo_monitor = request.form.get('modelo_monitor') or "N/A"
        serial_monitor = request.form.get('serial_monitor') or "N/A"
        cargador_monitor = request.form.get('cargador_monitor') or "N/A"
        
        observacion = request.form.get('observacion')

        template_path = os.path.join(app.root_path, 'static', 'actas', 'asignacion.docx')
        if not os.path.exists(template_path):
            return "Error: La plantilla 'asignacion.docx' no se encontró.", 404
        
        context = {
            'ciudad': ciudad, 'dia': dia, 'mes': mes, 'año': año,
            'marca_celular': marca_celular, 'modelo_celular': modelo_celular, 'serial_celular': serial_celular,
            'imei_celular': imei_celular, 'cargador_celular': cargador_celular, 'linea_celular': linea_celular,
            'marca_portatil': marca_portatil, 'modelo_portatil': modelo_portatil, 'serial_portatil': serial_portatil,
            'cargador_portatil': cargador_portatil,
            'teclado_accesorio': teclado_accesorio, 'mouse_accesorio': mouse_accesorio, 'base_accesorio': base_accesorio,
            'diadema_accesorio': diadema_accesorio,
            'marca_monitor': marca_monitor, 'modelo_monitor': modelo_monitor, 'serial_monitor': serial_monitor,
            'cargador_monitor': cargador_monitor,
            'observaciones': observacion, 'cesantias': cesantias,
            'nombre_completo': nombre_completo, 'cedula': cedula
        }



        values = [cedula, nombre_completo, ciudad, "ASIGNACION", f"{dia} de {mes} de {año}", (observacion or 'Ninguna'), "PENDIENTE"]

        if insert_acta(values):
            return process_document_and_convert_to_pdf(template_path, context, "ASIGNACION", nombre_completo, cedula)
        else:
            return "Error al guardar los datos en la base de datos.", 500

@app.route('/generar_devolver', methods=['POST'])
def generar_devolver():
    if request.method == 'POST':
        ciudad = request.form.get('ciudad')
        dia = request.form.get('dia')
        mes = request.form.get('mes')
        año = request.form.get('año')
        nombre_completo = request.form.get('nombre_completo', '').upper()
        cedula = request.form.get('cedula')
        
        marca_1 = request.form.get('marca_1')
        modelo_1 = request.form.get('modelo_1')
        serial_1 = request.form.get('serial_1')
        cargador_1 = request.form.get('cargador_1')
        estuche_1 = request.form.get('estuche_1')

        marca_2 = request.form.get('marca_2')
        modelo_2 = request.form.get('modelo_2')
        serial_2 = request.form.get('serial_2')
        cargador_2 = request.form.get('cargador_2')
        estuche_2 = request.form.get('estuche_2')

        marca_3 = request.form.get('marca_3')
        modelo_3 = request.form.get('modelo_3')
        serial_3 = request.form.get('serial_3')
        cargador_3 = request.form.get('cargador_3')
        estuche_3 = request.form.get('estuche_3')

        marca_4 = request.form.get('marca_4')
        modelo_4 = request.form.get('modelo_4')
        serial_4 = request.form.get('serial_4')
        cargador_4 = request.form.get('cargador_4')
        estuche_4 = request.form.get('estuche_4')

        marca_5 = request.form.get('marca_5')
        modelo_5 = request.form.get('modelo_5')
        serial_5 = request.form.get('serial_5')
        cargador_5 = request.form.get('cargador_5')
        estuche_5 = request.form.get('estuche_5')

        observacion = request.form.get('observacion')

        template_path = os.path.join(app.root_path, 'static', 'actas', 'devolucion.docx')
        if not os.path.exists(template_path):
            return "Error: La plantilla 'devolucion.docx' no se encontró.", 404
        
        context = {
            'ciudad': ciudad, 'dia': dia, 'mes': mes, 'año': año,
            'nombre_completo': nombre_completo, 'cedula': cedula, 'observaciones': observacion,
            'marca_1': marca_1, 'modelo_1': modelo_1, 'serial_1': serial_1, 'cargador_1': cargador_1, 'estuche_1': estuche_1,
            'marca_2': marca_2, 'modelo_2': modelo_2, 'serial_2': serial_2, 'cargador_2': cargador_2, 'estuche_2': estuche_2,
            'marca_3': marca_3, 'modelo_3': modelo_3, 'serial_3': serial_3, 'cargador_3': cargador_3, 'estuche_3': estuche_3,
            'marca_4': marca_4, 'modelo_4': modelo_4, 'serial_4': serial_4, 'cargador_4': cargador_4, 'estuche_4': estuche_4,
            'marca_5': marca_5, 'modelo_5': modelo_5, 'serial_5': serial_5, 'cargador_5': cargador_5, 'estuche_5': estuche_5
        }
        values = [cedula, nombre_completo, ciudad, "DEVOLUCION", f"{dia} de {mes} de {año}", observacion, "PENDIENTE"]

        if insert_acta(values):
            return process_document_and_convert_to_pdf(template_path, context, "DEVOLUCION", nombre_completo, cedula)
        else:
            return "Error al guardar los datos en la base de datos.", 500

@app.route('/generar_descontar', methods=['POST'])
def generar_descontar():
    if request.method == 'POST':
        ciudad = request.form.get('ciudad')
        dia = request.form.get('dia')
        mes = request.form.get('mes')
        año = request.form.get('año')
        nombre_completo = request.form.get('nombre_completo', '').upper()
        cedula = request.form.get('cedula')
        expedicion = request.form.get('expedicion')
        cesantias = request.form.get('cesantias')

        valor = request.form.get('valor')
        precio = request.form.get('precio')
        razon = request.form.get('razon')

        valor_cuota = request.form.get('valor_cuota')
        precio_cuota = request.form.get('precio_cuota')
        cuotas = request.form.get('cuotas')

        template_path = os.path.join(app.root_path, 'static', 'actas', 'descuento.docx')
        if not os.path.exists(template_path):
            return "Error: La plantilla 'descuento.docx' no se encontró.", 404
        
        context = {
            'ciudad': ciudad, 'dia': dia, 'mes': mes, 'año': año,
            'nombre_completo': nombre_completo, 'cedula': cedula,
            'expedicion': expedicion, 'cesantias': cesantias, 'razon': razon,
            'valor': valor, 'precio': precio,
            'valor_cuota': valor_cuota, 'precio_cuota': precio_cuota, 'cuotas': cuotas
        }

        values = [cedula, nombre_completo, ciudad, "ASIGNACION", f"{dia} de {mes} de {año}", (razon or 'Ninguna'), "PENDIENTE"]

        if insert_acta(values):
            return process_document_and_convert_to_pdf(template_path, context, "DESCUENTO", nombre_completo, cedula)
        else:
            return "Error al guardar los datos en la base de datos.", 500

                   
@app.route('/generar_mantenimiento', methods=['POST'])
def generar_mantenimiento():
    if request.method == 'POST':
        ciudad = request.form.get('ciudad')
        dia = request.form.get('dia')
        mes = request.form.get('mes')
        año = request.form.get('año')

        mes_numero = request.form.get('mes_numero')
        fecha = f"{str(dia).zfill(2)}/{str(mes_numero).zfill(2)}/{año}"
        nombre_completo = request.form.get('nombre_completo', '').upper()
        cedula = request.form.get('cedula')
        cargo = request.form.get('cargo')


        if request.form.get('tipo_mantenimiento') == 'preventivo':
            preventivo, correctivo = "✔", ""
        elif request.form.get('tipo_mantenimiento') == 'correctivo':
            preventivo, correctivo = "", "✔"
        
        print(f"\n\n----------------------->TIPO DE MANTENIMIENTO recibido: {request.form.get('tipo_mantenimiento')}")

        serial = request.form.get('serial')

        formateo = "✔" if request.form.get('formateo') == 'on' else ""
        instalacion = "✔" if request.form.get('instalacion') == 'on' else ""
        limpieza = "✔" if request.form.get('limpieza') == 'on' else ""
        eliminacion_temporales = "✔" if request.form.get('eliminacion_temporales') == 'on' else ""
        actualizacion = "✔" if request.form.get('actualizacion') == 'on' else ""
        eliminacion_programas = "✔" if request.form.get('eliminacion_programas') == 'on' else ""
        cambio = "✔" if request.form.get('cambio') == 'on' else ""
        configuracion = "✔" if request.form.get('configuracion') == 'on' else ""
        observacion = request.form.get('observacion') or ''

        print(f"\n\n----------------------->FORMATEO recibido: {formateo}")

        template_path = os.path.join(app.root_path, 'static', 'actas', 'mantenimiento.docx')
        if not os.path.exists(template_path):
            return "Error: La plantilla 'mantenimiento.docx' no se encontró.", 404
        
        context = {
            'ciudad': ciudad, 'dia': dia, 'mes': mes, 'año': año, 'fecha': fecha,
            'nombre_completo': nombre_completo, 'cedula': cedula, 'cargo': cargo,
            'serial': serial,
            'preventivo': preventivo, 'correctivo': correctivo,
            'formateo': formateo, 'instalacion': instalacion, 'limpieza': limpieza,
            'eliminacion_temporales': eliminacion_temporales, 'actualizacion': actualizacion,
            'eliminacion_programas': eliminacion_programas, 'cambio': cambio, 'configuracion': configuracion,
            'observaciones': observacion,
        }


        return process_document_and_convert_to_pdf(template_path, context, "MANTENIMIENTO", nombre_completo, cedula)



if __name__ == '__main__':
    create_database()
    app.run(debug=False)