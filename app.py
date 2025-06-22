from flask import Flask, render_template, request
from docxtpl import DocxTemplate
import io
from flask import send_file

app = Flask(__name__)

@app.route('/')
def mostrar_formulario():
    return render_template('form.html')


@app.route('/', methods=['POST'])
def generar_acta():
    if request.method == 'POST':
        doc = DocxTemplate("./static/actas/acta_entrega.docx")
        nombre = request.form.get('nombre')
        cedula = request.form.get('cedula')
        ciudad_expedicion = request.form.get('ciudad_expedicion')
        ciudad = request.form.get('ciudad')
        dia = request.form.get('dia')
        mes = request.form.get('mes')

        context = {
            'ciudad': ciudad,
            'dia': dia,
            'mes': mes,
            'nombre': nombre,
            'cedula': cedula,
            'ciudad_expedicion': ciudad_expedicion
        }

        doc.render(context)

        file_stream = io.BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)

        nombre_archivo_descarga = f"ACTA - {context['nombre']} - {context['cedula']}.docx"

        return send_file(
            file_stream,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=nombre_archivo_descarga
        )
    
if __name__ == '__main__':
    app.run(debug=True)