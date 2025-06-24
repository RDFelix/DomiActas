from flask import Flask, render_template, request
from docxtpl import DocxTemplate
import io
from flask import send_file
import os

app = Flask(__name__)

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

@app.route('/generar_entrega', methods=['POST'])
def generar_entrega():
    """
    Genera un acta de entrega a partir de los datos del formulario y una plantilla DOCX.
    """
    if request.method == 'POST':
        # Obtener los datos del formulario de entrega
        ciudad = request.form.get('ciudad_entrega')
        dia = request.form.get('dia_entrega')
        mes = request.form.get('mes_entrega')
        nombre_completo = request.form.get('nombre_completo_entrega')
        cedula = request.form.get('cedula_entrega')
        marca_equipo = request.form.get('marca_equipo')
        modelo_equipo = request.form.get('modelo_equipo')
        serial_equipo = request.form.get('serial_equipo_entrega')
        imei_equipo = request.form.get('imei_equipo')
        linea_celular = request.form.get('linea_celular')
        cesantias = request.form.get('cesantias_entrega')

        # Ruta a la plantilla DOCX
        template_path = os.path.join(app.root_path, 'static', 'actas', 'acta_entrega.docx')
        doc = DocxTemplate(template_path)

        # Crear el contexto para la plantilla
        context = {
            'ciudad': ciudad,
            'dia': dia,
            'mes': mes,
            'nombre_completo': nombre_completo,
            'cedula': cedula,
            'marca_equipo': marca_equipo,
            'modelo_equipo': modelo_equipo,
            'serial_equipo': serial_equipo,
            'imei_equipo': imei_equipo,
            'linea_celular': linea_celular,
            'cesantias': cesantias
        }

        # Renderizar la plantilla con el contexto
        doc.render(context)

        # Guardar el documento generado en un stream de bytes
        file_stream = io.BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)

        # Nombre del archivo para la descarga
        nombre_archivo_descarga = f"ACTA_ENTREGA_{nombre_completo}_{cedula}.docx"

        # Enviar el archivo para descarga
        return send_file(
            file_stream,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=nombre_archivo_descarga
        )

# --- Ruta para generar Acta de Devolución ---
@app.route('/generar_devolucion', methods=['POST'])
def generar_devolucion():
    """
    Genera un acta de devolución a partir de los datos del formulario y una plantilla DOCX.
    """
    if request.method == 'POST':
        # Obtener los datos del formulario de devolución
        nombre_completo = request.form.get('nombre_devolucion')
        cedula = request.form.get('cedula_devolucion')
        ciudad = request.form.get('ciudad_devolucion')
        dia = request.form.get('dia_devolucion')
        mes = request.form.get('mes_devolucion')
        marca_equipo_dev = request.form.get('marca_equipo_dev')
        modelo_equipo_dev = request.form.get('modelo_equipo_dev')
        serial_equipo = request.form.get('serial_equipo_devolucion')
        accesorios_devolucion = request.form.get('accesorios_devolucion')

        # Ruta a la plantilla DOCX
        template_path = os.path.join(app.root_path, 'static', 'actas', 'acta_devolucion.docx')
        doc = DocxTemplate(template_path)

        # Crear el contexto para la plantilla
        context = {
            'nombre_completo': nombre_completo,
            'cedula': cedula,
            'ciudad': ciudad,
            'dia': dia,
            'mes': mes,
            'marca_equipo_dev': marca_equipo_dev,
            'modelo_equipo_dev': modelo_equipo_dev,
            'serial_equipo': serial_equipo,
            'accesorios_devolucion': accesorios_devolucion
        }

        # Renderizar la plantilla con el contexto
        doc.render(context)

        # Guardar el documento generado en un stream de bytes
        file_stream = io.BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)

        # Nombre del archivo para la descarga
        nombre_archivo_descarga = f"ACTA_DEVOLUCION_{nombre_completo}_{cedula}.docx"

        # Enviar el archivo para descarga
        return send_file(
            file_stream,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=nombre_archivo_descarga
        )

# --- Ruta para generar Acta de Descuento ---
@app.route('/generar_descuento', methods=['POST'])
def generar_descuento():
    """
    Genera un acta de descuento a partir de los datos del formulario y una plantilla DOCX.
    """
    if request.method == 'POST':
        # Obtener los datos del formulario de descuento
        nombre_completo = request.form.get('nombre_descuento')
        cedula = request.form.get('cedula_descuento')
        ciudad = request.form.get('ciudad_descuento')
        dia = request.form.get('dia_descuento')
        mes = request.form.get('mes_descuento')
        razon = request.form.get('razon_descuento')
        valor = request.form.get('valor_descuento')
        cuotas = request.form.get('cuotas_descuento')
        cesantias = request.form.get('cesantias_descuento')

        # Ruta a la plantilla DOCX
        template_path = os.path.join(app.root_path, 'static', 'actas', 'acta_descuento.docx')
        doc = DocxTemplate(template_path)

        # Crear el contexto para la plantilla
        context = {
            'nombre_completo': nombre_completo,
            'cedula': cedula,
            'ciudad': ciudad,
            'dia': dia,
            'mes': mes,
            'razon': razon,
            'valor': valor,
            'cuotas': cuotas,
            'cesantias': cesantias
        }

        # Renderizar la plantilla con el contexto
        doc.render(context)

        # Guardar el documento generado en un stream de bytes
        file_stream = io.BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)

        # Nombre del archivo para la descarga
        nombre_archivo_descarga = f"ACTA_DESCUENTO_{nombre_completo}_{cedula}.docx"

        # Enviar el archivo para descarga
        return send_file(
            file_stream,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=nombre_archivo_descarga
        )

# --- Ruta para generar Acta de Mantenimiento ---
@app.route('/generar_mantenimiento', methods=['POST'])
def generar_mantenimiento():
    """
    Genera un acta de mantenimiento a partir de los datos del formulario y una plantilla DOCX.
    """
    if request.method == 'POST':
        # Obtener los datos del formulario de mantenimiento
        ciudad = request.form.get('ciudad_mantenimiento')
        dia = request.form.get('dia_mantenimiento')
        mes = request.form.get('mes_mantenimiento')
        serial = request.form.get('serial_mantenimiento')
        observaciones = request.form.get('observaciones_mantenimiento')

        # Ruta a la plantilla DOCX
        template_path = os.path.join(app.root_path, 'static', 'actas', 'acta_mantenimiento.docx')
        doc = DocxTemplate(template_path)

        # Crear el contexto para la plantilla
        context = {
            'ciudad': ciudad,
            'dia': dia,
            'mes': mes,
            'serial': serial,
            'observaciones': observaciones
        }

        # Renderizar la plantilla con el contexto
        doc.render(context)

        # Guardar el documento generado en un stream de bytes
        file_stream = io.BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)

        # Nombre del archivo para la descarga
        nombre_archivo_descarga = f"ACTA_MANTENIMIENTO_{serial}.docx"

        # Enviar el archivo para descarga
        return send_file(
            file_stream,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=nombre_archivo_descarga
        )
    
if __name__ == '__main__':
    app.run(debug=True)