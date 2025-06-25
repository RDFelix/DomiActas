from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from docxtpl import DocxTemplate
import io
import os
import tempfile
from docx2pdf import convert
import traceback
import pythoncom


app = Flask(__name__)

@app.route('/')
def index():
    """Ruta para mostrar el formulario de asignación."""
    return render_template('asignacion.html')


# Definir las rutas para renderizar los formularios HTML
@app.route('/asignacion')
def asignacion():
    """Ruta para mostrar el formulario de asignación."""
    return render_template('asignacion.html')

@app.route('/devolucion')
def devolucion():
    """Ruta para mostrar el formulario de devolución."""
    return render_template('devolucion.html')

@app.route('/descuento')
def descuento():
    """Ruta para mostrar el formulario de descuento."""
    return render_template('descuento.html')

@app.route('/mantenimiento')
def mantenimiento():
    """Ruta para mostrar el formulario de mantenimiento."""
    return render_template('mantenimiento.html')



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
            'observaciones': observacion,
            'cesantias': cesantias,
            'nombre_completo': nombre_completo,
            'cedula': cedula
        }

        doc.render(context)

        temp_docx_path = None
        temp_pdf_path = None
        try:
            # Inicializar COM para el hilo actual. Es crucial para cada operación.
            # Se usa un try/finally para asegurar que CoUninitialize sea llamado.
            pythoncom.CoInitialize() 

            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                doc.save(tmp_docx.name)
                temp_docx_path = tmp_docx.name

            temp_pdf_path = os.path.join(os.path.dirname(temp_docx_path), os.path.basename(temp_docx_path).replace(".docx", ".pdf"))

            # Convertir DOCX a PDF
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
            print(f"Error al convertir DOCX a PDF en generar_asignacion: {e}")
            traceback.print_exc()
            return "Error al generar el PDF. Asegúrate de que Microsoft Word esté instalado y no haya procesos de Word colgados.", 500
        finally:
            # Desinicializar COM en el mismo hilo.
            try:
                pythoncom.CoUninitialize()
            except Exception as e:
                print(f"Error al desinicializar COM: {e}")
            
            # Limpiar los archivos temporales
            if temp_docx_path and os.path.exists(temp_docx_path):
                try:
                    os.remove(temp_docx_path)
                except OSError as e:
                    print(f"Error al eliminar el archivo temporal DOCX {temp_docx_path}: {e}")
            if temp_pdf_path and os.path.exists(temp_pdf_path):
                try:
                    os.remove(temp_pdf_path)
                except OSError as e:
                    print(f"Error al eliminar el archivo temporal PDF {temp_pdf_path}: {e}")

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
        doc = DocxTemplate(template_path)

        context = {
            'ciudad': ciudad,
            'dia': dia,
            'mes': mes,
            'año': año,
            'nombre_completo': nombre_completo,
            'cedula': cedula,
            'observaciones': observacion,

            'marca_1': marca_1,
            'modelo_1': modelo_1,
            'serial_1': serial_1,
            'cargador_1': cargador_1,
            'estuche_1': estuche_1,

            'marca_2': marca_2,
            'modelo_2': modelo_2,
            'serial_2': serial_2,
            'cargador_2': cargador_2,
            'estuche_2': estuche_2,

            'marca_3': marca_3,
            'modelo_3': modelo_3,
            'serial_3': serial_3,
            'cargador_3': cargador_3,
            'estuche_3': estuche_3,

            'marca_4': marca_4,
            'modelo_4': modelo_4,
            'serial_4': serial_4,
            'cargador_4': cargador_4,
            'estuche_4': estuche_4,

            'marca_5': marca_5,
            'modelo_5': modelo_5,
            'serial_5': serial_5,
            'cargador_5': cargador_5,
            'estuche_5': estuche_5
        }

        doc.render(context)

        temp_docx_path = None
        temp_pdf_path = None
        try:
            # Inicializar COM para el hilo actual. Es crucial para cada operación.
            # Se usa un try/finally para asegurar que CoUninitialize sea llamado.
            pythoncom.CoInitialize() 

            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                doc.save(tmp_docx.name)
                temp_docx_path = tmp_docx.name

            temp_pdf_path = os.path.join(os.path.dirname(temp_docx_path), os.path.basename(temp_docx_path).replace(".docx", ".pdf"))

            # Convertir DOCX a PDF
            convert(temp_docx_path, temp_pdf_path)

            with open(temp_pdf_path, 'rb') as f:
                file_stream = io.BytesIO(f.read())
            file_stream.seek(0)

            nombre_archivo_descarga = f"DEVOLUCION_{nombre_completo}_{cedula}.pdf"

            return send_file(
                file_stream,
                mimetype='application/pdf',
                as_attachment=True,
                download_name=nombre_archivo_descarga
            )
        except Exception as e:
            print(f"Error al convertir DOCX a PDF en generar_devolucion: {e}")
            traceback.print_exc()
            return "Error al generar el PDF. Asegúrate de que Microsoft Word esté instalado y no haya procesos de Word colgados.", 500
        finally:
            # Desinicializar COM en el mismo hilo.
            try:
                pythoncom.CoUninitialize()
            except Exception as e:
                print(f"Error al desinicializar COM: {e}")
            
            # Limpiar los archivos temporales
            if temp_docx_path and os.path.exists(temp_docx_path):
                try:
                    os.remove(temp_docx_path)
                except OSError as e:
                    print(f"Error al eliminar el archivo temporal DOCX {temp_docx_path}: {e}")
            if temp_pdf_path and os.path.exists(temp_pdf_path):
                try:
                    os.remove(temp_pdf_path)
                except OSError as e:
                    print(f"Error al eliminar el archivo temporal PDF {temp_pdf_path}: {e}")
                    
@app.route('/generar_descuento', methods=['POST'])
def generar_descuento():
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
        cantidad = request.form.get('cantidad')
        razon = request.form.get('razon')

        valor_cuota = request.form.get('valor_cuota')
        cantidad_cuota = request.form.get('cantidad_cuota')
        cuotas = request.form.get('cuotas')

        template_path = os.path.join(app.root_path, 'static', 'actas', 'descuento.docx')
        if not os.path.exists(template_path):
            return "Error: La plantilla 'descuento.docx' no se encontró.", 404
        doc = DocxTemplate(template_path)

        context = {
            'ciudad': ciudad,
            'dia': dia,
            'mes': mes,
            'año': año,
            'nombre_completo': nombre_completo,
            'cedula': cedula,
            'expedicion': expedicion,
            'cesantias': cesantias,
            'valor': valor,
            'cantidad': cantidad,
            'razon': razon,
            'valor_cuota': valor_cuota,
            'cantidad_cuota': cantidad_cuota,
            'cuotas': cuotas
        }

        doc.render(context)

        temp_docx_path = None
        temp_pdf_path = None
        try:
            pythoncom.CoInitialize()
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                doc.save(tmp_docx.name)
                temp_docx_path = tmp_docx.name

            temp_pdf_path = os.path.join(os.path.dirname(temp_docx_path), os.path.basename(temp_docx_path).replace(".docx", ".pdf"))
            convert(temp_docx_path, temp_pdf_path)

            with open(temp_pdf_path, 'rb') as f:
                file_stream = io.BytesIO(f.read())
            file_stream.seek(0)

            nombre_archivo_descarga = f"DESCUENTO_{nombre_completo}_{cedula}.pdf"

            return send_file(
                file_stream,
                mimetype='application/pdf',
                as_attachment=True,
                download_name=nombre_archivo_descarga
            )
        except Exception as e:
            print(f"Error al convertir DOCX a PDF en generar_descuento: {e}")
            traceback.print_exc()
            return "Error al generar el PDF. Asegúrate de que Microsoft Word esté instalado y no haya procesos de Word colgados.", 500
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception as e:
                print(f"Error al desinicializar COM: {e}")
            if temp_docx_path and os.path.exists(temp_docx_path):
                try:
                    os.remove(temp_docx_path)
                except OSError as e:
                    print(f"Error al eliminar el archivo temporal DOCX {temp_docx_path}: {e}")
            if temp_pdf_path and os.path.exists(temp_pdf_path):
                try:
                    os.remove(temp_pdf_path)
                except OSError as e:
                    print(f"Error al eliminar el archivo temporal PDF {temp_pdf_path}: {e}")

if __name__ == '__main__':
    app.run(debug=True)  # Cambia debug=False en producción
    #app.run(host='0.0.0.0', port=5000)