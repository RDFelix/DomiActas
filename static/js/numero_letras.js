const $ = element => document.getElementById(element);

function format(numero) {
    const num = parseFloat(numero);
    if (isNaN(num)) {
        return "";
    }
    return num.toLocaleString('es-CO', {
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
    });
}

function numeroALetras(num) {
    const unidades = ['', 'UNO', 'DOS', 'TRES', 'CUATRO', 'CINCO', 'SEIS', 'SIETE', 'OCHO', 'NUEVE'];
    const decenas = ['', '', 'VEINTE', 'TREINTA', 'CUARENTA', 'CINCUENTA', 'SESENTA', 'SETENTA', 'OCHENTA', 'NOVENTA'];
    const dieces = ['DIEZ', 'ONCE', 'DOCE', 'TRECE', 'CATORCE', 'QUINCE', 'DIECISÉIS', 'DIECISIETE', 'DIECIOCHO', 'DIECINUEVE'];
    const centenas = ['', 'CIENTO', 'DOSCIENTOS', 'TRESCIENTOS', 'CUATROCIENTOS', 'QUINIENTOS', 'SEISCIENTOS', 'SETECIENTOS', 'OCHOCIENTOS', 'NOVECIENTOS'];

    if (num === 0) {
        return 'CERO';
    }

    if (num < 0) {
        return 'MENOS ' + numeroALetras(Math.abs(num));
    }

    let palabras = '';

    if (num >= 1000000000) {
        palabras += numeroALetras(Math.floor(num / 1000000000)) + ' MIL MILLONES ';
        num %= 1000000000;
    }

    if (num >= 1000000) {
        if (num >= 2000000) {
            palabras += numeroALetras(Math.floor(num / 1000000)) + ' MILLONES ';
        } else {
            palabras += 'UN MILLÓN ';
        }
        num %= 1000000;
    }

    if (num >= 1000) {
        if (num === 1000) {
            palabras += 'MIL ';
        } else {
            palabras += numeroALetras(Math.floor(num / 1000)) + ' MIL ';
        }
        num %= 1000;
    }

    if (num >= 100) {
        if (num === 100) {
            palabras += 'CIEN ';
        } else {
            palabras += centenas[Math.floor(num / 100)] + ' ';
        }
        num %= 100;
    }

    if (num >= 20) {
        palabras += decenas[Math.floor(num / 10)];
        if (num % 10 !== 0) {
            palabras += ' y ' + unidades[num % 10];
        }
    } else if (num >= 10) {
        palabras += dieces[num - 10];
    } else if (num > 0) {
        palabras += unidades[num];
    }

    return palabras.trim();
}

$('cuotas').addEventListener('input', function() {
    $('valor').value = numeroALetras(parseInt($('valor_input').value, 0));
    $('precio').value = format(parseInt($('valor_input').value, 10))
    $('precio_cuota').value = numeroALetras(parseInt($('valor_input').value, 10) / parseInt(this.value, 10));
    $('valor_cuota').value = format(parseInt($('valor_input').value, 10) / parseInt(this.value, 10));
    $('valor_cuota_disabled').value = format(parseInt($('valor_input').value, 10) / parseInt(this.value, 10));
});

$('valor_input').addEventListener('input', function() {
    $('valor').value = numeroALetras(parseInt($('valor_input').value, 0));
    $('precio').value = format(parseInt($('valor_input').value, 10))
    $('precio_cuota').value = numeroALetras(parseInt(this.value, 10) / parseInt($('cuotas').value, 10));
    $('valor_cuota').value = format(parseInt(this.value, 10) / parseInt($('cuotas').value, 10));
    $('valor_cuota_disabled').value = format(parseInt(this.value, 10) / parseInt($('cuotas').value, 10));
});

function actualizarTexto() {

    $('ciudad_txt').textContent = $('ciudad').value || '____';
    $('dia_txt').textContent = $('dia').value || '____';
    $('mes_txt').textContent = $('mes').value || '________';
    $('año_txt').textContent = $('año').value || '______';
    $('nombre_txt').textContent = $('nombre_completo').value || '________________________';
    $('cedula_txt').textContent = $('cedula').value || '___________';
    $('ciudad_expedicion_txt').textContent = $('expedicion').value || '_____________';
    $('valor_txt').textContent = numeroALetras(parseInt($('valor_input').value, 0)) || '___________';
    $('precio_txt').textContent = format(parseInt($('valor_input').value, 10)) || '___________';
    $('razon_txt').textContent = $('razon').value || '________________________';

    $('valor_cuota_txt').textContent = numeroALetras(parseInt($('valor_input').value, 10) / parseInt($('cuotas').value, 10)) || '___________';
    $('precio_cuota_txt').textContent = format(parseInt($('valor_input').value, 10) / parseInt($('cuotas').value, 10)) || '___________';

    $('cuotas_txt').textContent = $('cuotas').value || '___';
    $('cesantias_txt').textContent = $('cesantias').value || '___________';
}

setInterval(actualizarTexto, 100);