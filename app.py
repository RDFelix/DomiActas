from flask import Flask, render_template, request, send_file, redirect, url_for, flash, g
from docxtpl import DocxTemplate
import io
import os
import tempfile
from docx2pdf import convert
import traceback
import pythoncom
import time # Para medir tiempos y depurar

app = Flask(__name__)
# app.secret_key = 'tu_clave_secreta_aqui' # Descomentar si usas flash messages

# --- Función Auxiliar para Manejar el Proceso de Conversión ---
def process_document_and_convert_to_pdf(template_path, context, file_prefix, nombre_completo, cedula):
    """
    Función auxiliar para procesar la plantilla DOCX, convertir a PDF y enviar el archivo.
    Maneja la inicialización/desinicialización de COM y la creación/limpieza de archivos temporales.
    """
    doc = DocxTemplate(template_path)
    
    # 1. Renderizar la plantilla (esto suele ser rápido)
    start_render_time = time.time()
    doc.render(context)
    end_render_time = time.time()
    print(f"Tiempo de renderizado de DOCX: {end_render_time - start_render_time:.4f} segundos")

    temp_docx_path = None
    temp_pdf_path = None
    
    # --- BLOQUE CLAVE: Gestión de COM ---
    try:
        # Inicializar COM para el hilo actual. Es CRUCIAL para cada operación con Word.
        # Esto asegura que el hilo que procesa la solicitud tenga un entorno COM válido.
        pythoncom.CoInitialize() 

        # Crea el archivo DOCX temporal
        start_save_docx_time = time.time()
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
            doc.save(tmp_docx.name)
            temp_docx_path = tmp_docx.name
        end_save_docx_time = time.time()
        print(f"Tiempo de guardado de DOCX temporal: {end_save_docx_time - start_save_docx_time:.4f} segundos")

        # Define la ruta del PDF temporal
        temp_pdf_path = os.path.join(os.path.dirname(temp_docx_path), os.path.basename(temp_docx_path).replace(".docx", ".pdf"))

        # 2. Convertir DOCX a PDF (este es el paso lento, depende de Word)
        start_convert_time = time.time()
        # Mantenemos solo los argumentos esenciales
        convert(temp_docx_path, temp_pdf_path) 
        end_convert_time = time.time()
        print(f"Tiempo de conversión DOCX a PDF: {end_convert_time - start_convert_time:.4f} segundos")

        # Lee el PDF generado en un stream para enviarlo
        with open(temp_pdf_path, 'rb') as f:
            file_stream = io.BytesIO(f.read())
        file_stream.seek(0)

        nombre_archivo_descarga = f"{file_prefix}_{nombre_completo}_{cedula}.pdf"

        return send_file(
            file_stream,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=nombre_archivo_descarga
        )
    except Exception as e:
        print(f"Error en process_document_and_convert_to_pdf: {e}")
        traceback.print_exc()
        return "Error al generar el PDF. Asegúrate de que Microsoft Word esté instalado y no haya procesos de Word colgados.", 500
    finally:
        # Desinicializar COM en el mismo hilo. Esto es CRUCIAL para liberar los recursos.
        try:
            pythoncom.CoUninitialize()
        except Exception as e_uninit:
            # Captura errores si CoUninitialize es llamado sin CoInitialize previo o si hay otros problemas
            print(f"Advertencia: Error al desinicializar COM (puede ser normal si COM no se inicializó correctamente): {e_uninit}")
            
        # Limpiar los archivos temporales, independientemente del éxito o error
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


# --- Rutas para renderizar los formularios HTML (sin cambios) ---
@app.route('/')
def index():
    return render_template('asignacion.html')

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

# --- Rutas para Generar y Convertir Documentos (simplificadas) ---

@app.route('/generar_asignar', methods=['POST'])
def generar_asignar():
    if request.method == 'POST':
        # Recopilación de datos del formulario (sin cambios)
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

        return process_document_and_convert_to_pdf(template_path, context, "ASIGNACION", nombre_completo, cedula)

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

        return process_document_and_convert_to_pdf(template_path, context, "DEVOLUCION", nombre_completo, cedula)
                   
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

        return process_document_and_convert_to_pdf(template_path, context, "DESCUENTO", nombre_completo, cedula)
                   
@app.route('/generar_mantenimiento', methods=['POST'])
def generar_mantenimiento():
    if request.method == 'POST':
        ciudad = request.form.get('ciudad')
        fecha = request.form.get('fecha')
        nombre_completo = request.form.get('nombre_completo', '').upper()
        cedula = request.form.get('cedula')
        cargo = request.form.get('cargo')
        
        preventivo = request.form.get('preventivo') == 'on'
        correctivo = request.form.get('correctivo') == 'on'

        serial = request.form.get('serial')

        formateo = request.form.get('formateo') == 'on'
        instalacion = request.form.get('instalacion') == 'on'
        limpieza = request.form.get('limpieza') == 'on'
        eliminacion_temporales = request.form.get('eliminacion_temporales') == 'on'
        actualizacion = request.form.get('actualizacion') == 'on'
        eliminacion_programas = request.form.get('eliminacion_programas') == 'on'
        cambio = request.form.get('cambio') == 'on'
        configuracion = request.form.get('configuracion') == 'on'
        observacion = request.form.get('observacion')

        template_path = os.path.join(app.root_path, 'static', 'actas', 'mantenimiento.docx')
        if not os.path.exists(template_path):
            return "Error: La plantilla 'mantenimiento.docx' no se encontró.", 404
        
        context = {
            'ciudad': ciudad, 'fecha': fecha,
            'nombre_completo': nombre_completo, 'cedula': cedula, 'cargo': cargo,
            'serial': serial,
            'preventivo': preventivo, 'correctivo': correctivo,
            'formateo': formateo, 'instalacion': instalacion, 'limpieza': limpieza,
            'eliminacion_temporales': eliminacion_temporales, 'actualizacion': actualizacion,
            'eliminacion_programas': eliminacion_programas, 'cambio': cambio, 'configuracion': configuracion,
            'observaciones': observacion,
        }

        return process_document_and_convert_to_pdf(template_path, context, "MANTENIMIENTO", nombre_completo, cedula)

# --- Punto de Entrada de la Aplicación ---
if __name__ == '__main__':
    app.run(debug=True) # Cambia debug=False en producción