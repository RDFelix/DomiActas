from flask import send_file
from docxtpl import DocxTemplate
import io
import os
import tempfile
from docx2pdf import convert
import traceback
import pythoncom
import time

def process_document_and_convert_to_pdf(template_path, context, file_prefix, nombre_completo, cedula):
    doc = DocxTemplate(template_path)
    
    start_render_time = time.time()
    doc.render(context)
    end_render_time = time.time()
    print(f"Tiempo de renderizado de DOCX: {end_render_time - start_render_time:.4f} segundos")

    temp_docx_path = None
    temp_pdf_path = None
    
    try:
        pythoncom.CoInitialize() 

        start_save_docx_time = time.time()
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
            doc.save(tmp_docx.name)
            temp_docx_path = tmp_docx.name
        end_save_docx_time = time.time()
        print(f"Tiempo de guardado de DOCX temporal: {end_save_docx_time - start_save_docx_time:.4f} segundos")

        temp_pdf_path = os.path.join(os.path.dirname(temp_docx_path), os.path.basename(temp_docx_path).replace(".docx", ".pdf"))

        start_convert_time = time.time()
        convert(temp_docx_path, temp_pdf_path) 
        end_convert_time = time.time()
        print(f"Tiempo de conversión DOCX a PDF: {end_convert_time - start_convert_time:.4f} segundos")

        with open(temp_pdf_path, 'rb') as f:
            file_stream = io.BytesIO(f.read())
        file_stream.seek(0)

        nombre_archivo_descarga = f"{file_prefix} - {nombre_completo} - {cedula}.pdf"

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
        try:
            pythoncom.CoUninitialize()
        except Exception as e_uninit:
            print(f"Advertencia: Error al desinicializar COM (puede ser normal si COM no se inicializó correctamente): {e_uninit}")
            
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
