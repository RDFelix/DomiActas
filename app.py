from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate
import io
import os
import tempfile # Importar el módulo tempfile para archivos temporales
from docx2pdf import convert # Importar la función convert de docx2pdf

app = Flask(__name__)

# Definir las rutas para renderizar los formularios HTML
@app.route('/asignacion_celular')
def asignacion():
    return render_template('asignacion_celular.html')

@app.route('/devolucion')
def devolucion():
    return render_template('devolucion.html')

@app.route('/descuento')
def descuento():
    return render_template('descuento.html')

@app.route('/mantenimiento')
def mantenimiento():
    return render_template('mantenimiento.html')

@app.route('/asignar_celular', methods=['POST'])
def asignar_celular():
    if request.method == 'POST':
        ciudad = request.form.get('ciudad')
        dia = request.form.get('dia')
        mes = request.form.get('mes')
        año = request.form.get('año')
        nombre_completo = (f"{request.form.get('nombres')} {request.form.get('apellidos')}").upper()
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
        doc = DocxTemplate(template_path)

        context = {
            'ciudad': ciudad,
            'dia': dia,
            'mes': mes,
            'año': año,
            'marca_celular': marca_celular,
            'modelo_celular': modelo_celular,
            'serial_celular': serial_celular,
            'imei_celular': imei_celular,
            'cargador_celular': cargador_celular,
            'linea_celular': linea_celular,

            'marca_portatil': marca_portatil,
            'modelo_portatil': modelo_portatil,
            'serial_portatil': serial_portatil,
            'cargador_portatil': cargador_portatil,

            'teclado_accesorio': teclado_accesorio,
            'mouse_accesorio': mouse_accesorio,
            'base_accesorio': base_accesorio,
            'diadema_accesorio': diadema_accesorio,

            'marca_monitor': marca_monitor,
            'modelo_monitor': modelo_monitor,
            'serial_monitor': serial_monitor,
            'cargador_monitor': cargador_monitor,

            'observacion': observacion,

            'cesantias': cesantias,
            'nombre': nombre_completo,
            'cedula': cedula
        }

        doc.render(context)

        temp_docx_path = None
        temp_pdf_path = None
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                doc.save(tmp_docx.name)
                temp_docx_path = tmp_docx.name

            temp_pdf_path = temp_docx_path.replace(".docx", ".pdf")

            convert(temp_docx_path, temp_pdf_path)

            with open(temp_pdf_path, 'rb') as f:
                file_stream = io.BytesIO(f.read())
            file_stream.seek(0)

            nombre_archivo_descarga = f"ASIGNACION_{nombre_completo}_{cedula}.pdf"

            return send_file(
                file_stream,
                mimetype='application/pdf',
                as_attachment=True,
                download_name=nombre_archivo_descarga
            )
        except Exception as e:
            print(f"Error al convertir DOCX a PDF en generar_entrega: {e}")
            return "Error al generar el PDF. Asegúrate de que LibreOffice (o Microsoft Word en Windows) esté instalado en el servidor.", 500
        finally:
            if temp_docx_path and os.path.exists(temp_docx_path):
                os.remove(temp_docx_path)
            if temp_pdf_path and os.path.exists(temp_pdf_path):
                os.remove(temp_pdf_path)

# --- Ruta para generar Acta de Devolución y convertir a PDF ---
@app.route('/generar_devolucion', methods=['POST'])
def generar_devolucion():
    """
    Genera un acta de devolución a partir de los datos del formulario,
    la renderiza en una plantilla DOCX y la convierte a PDF para descarga.
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

        # Verificar si la plantilla existe
        if not os.path.exists(template_path):
            return "Error: La plantilla 'acta_devolucion.docx' no se encontró.", 404

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

        # Usar archivos temporales para guardar DOCX y luego convertir a PDF
        temp_docx_path = None
        temp_pdf_path = None
        try:
            # Guardar el documento DOCX generado en un archivo temporal
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                doc.save(tmp_docx.name)
                temp_docx_path = tmp_docx.name

            # Definir la ruta para el archivo PDF temporal
            temp_pdf_path = temp_docx_path.replace(".docx", ".pdf")

            # Convertir DOCX a PDF
            convert(temp_docx_path, temp_pdf_path)

            # Leer el PDF generado en un stream de bytes
            with open(temp_pdf_path, 'rb') as f:
                file_stream = io.BytesIO(f.read())
            file_stream.seek(0)

            # Nombre del archivo para la descarga (ahora con extensión .pdf)
            nombre_archivo_descarga = f"ACTA_DEVOLUCION_{nombre_completo}_{cedula}.pdf"

            # Enviar el archivo PDF para descarga
            return send_file(
                file_stream,
                mimetype='application/pdf',
                as_attachment=True,
                download_name=nombre_archivo_descarga
            )
        except Exception as e:
            print(f"Error al convertir DOCX a PDF en generar_devolucion: {e}")
            return "Error al generar el PDF. Asegúrate de que LibreOffice (o Microsoft Word en Windows) esté instalado en el servidor.", 500
        finally:
            # Limpiar los archivos temporales
            if temp_docx_path and os.path.exists(temp_docx_path):
                os.remove(temp_docx_path)
            if temp_pdf_path and os.path.exists(temp_pdf_path):
                os.remove(temp_pdf_path)

# --- Ruta para generar Acta de Descuento y convertir a PDF ---
@app.route('/generar_descuento', methods=['POST'])
def generar_descuento():
    """
    Genera un acta de descuento a partir de los datos del formulario,
    la renderiza en una plantilla DOCX y la convierte a PDF para descarga.
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
        
        # Verificar si la plantilla existe
        if not os.path.exists(template_path):
            return "Error: La plantilla 'acta_descuento.docx' no se encontró.", 404

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

        # Usar archivos temporales para guardar DOCX y luego convertir a PDF
        temp_docx_path = None
        temp_pdf_path = None
        try:
            # Guardar el documento DOCX generado en un archivo temporal
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                doc.save(tmp_docx.name)
                temp_docx_path = tmp_docx.name

            # Definir la ruta para el archivo PDF temporal
            temp_pdf_path = temp_docx_path.replace(".docx", ".pdf")

            # Convertir DOCX a PDF
            convert(temp_docx_path, temp_pdf_path)

            # Leer el PDF generado en un stream de bytes
            with open(temp_pdf_path, 'rb') as f:
                file_stream = io.BytesIO(f.read())
            file_stream.seek(0)

            # Nombre del archivo para la descarga (ahora con extensión .pdf)
            nombre_archivo_descarga = f"ACTA_DESCUENTO_{nombre_completo}_{cedula}.pdf"

            # Enviar el archivo PDF para descarga
            return send_file(
                file_stream,
                mimetype='application/pdf',
                as_attachment=True,
                download_name=nombre_archivo_descarga
            )
        except Exception as e:
            print(f"Error al convertir DOCX a PDF en generar_descuento: {e}")
            return "Error al generar el PDF. Asegúrate de que LibreOffice (o Microsoft Word en Windows) esté instalado en el servidor.", 500
        finally:
            # Limpiar los archivos temporales
            if temp_docx_path and os.path.exists(temp_docx_path):
                os.remove(temp_docx_path)
            if temp_pdf_path and os.path.exists(temp_pdf_path):
                os.remove(temp_pdf_path)

# --- Ruta para generar Acta de Mantenimiento y convertir a PDF ---
@app.route('/generar_mantenimiento', methods=['POST'])
def generar_mantenimiento():
    """
    Genera un acta de mantenimiento a partir de los datos del formulario,
    la renderiza en una plantilla DOCX y la convierte a PDF para descarga.
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
        
        # Verificar si la plantilla existe
        if not os.path.exists(template_path):
            return "Error: La plantilla 'acta_mantenimiento.docx' no se encontró.", 404

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

        # Usar archivos temporales para guardar DOCX y luego convertir a PDF
        temp_docx_path = None
        temp_pdf_path = None
        try:
            # Guardar el documento DOCX generado en un archivo temporal
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                doc.save(tmp_docx.name)
                temp_docx_path = tmp_docx.name

            # Definir la ruta para el archivo PDF temporal
            temp_pdf_path = temp_docx_path.replace(".docx", ".pdf")

            # Convertir DOCX a PDF
            convert(temp_docx_path, temp_pdf_path)

            # Leer el PDF generado en un stream de bytes
            with open(temp_pdf_path, 'rb') as f:
                file_stream = io.BytesIO(f.read())
            file_stream.seek(0)

            # Nombre del archivo para la descarga (ahora con extensión .pdf)
            nombre_archivo_descarga = f"ACTA_MANTENIMIENTO_{serial}.pdf"

            # Enviar el archivo PDF para descarga
            return send_file(
                file_stream,
                mimetype='application/pdf',
                as_attachment=True,
                download_name=nombre_archivo_descarga
            )
        except Exception as e:
            print(f"Error al convertir DOCX a PDF en generar_mantenimiento: {e}")
            return "Error al generar el PDF. Asegúrate de que LibreOffice (o Microsoft Word en Windows) esté instalado en el servidor.", 500
        finally:
            # Limpiar los archivos temporales
            if temp_docx_path and os.path.exists(temp_docx_path):
                os.remove(temp_docx_path)
            if temp_pdf_path and os.path.exists(temp_pdf_path):
                os.remove(temp_pdf_path)

if __name__ == '__main__':
    app.run(debug=True)
