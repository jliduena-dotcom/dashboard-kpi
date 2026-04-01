// ==============================
// 1. LEER EL EXCEL
// ==============================
fetch('BASE DE KPI CCOT.xlsx')
    .then(res => res.arrayBuffer())
    .then(data => {
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        crearTabla(json);
        crearGrafico(json);
    });

// ==============================
// 2. CREAR TABLA (MISMA ESTRUCTURA EXCEL)
// ==============================
function crearTabla(data) {
    const table = document.getElementById('tablaExcel');
    table.innerHTML = "";

    // 1. ANOTA AQUÍ LAS COLUMNAS QUE NO SON PORCENTAJE (0 es la primera, 1 la segunda...)
    const columnasSoloNumeros = [0, 1]; // Ejemplo: Columna A y C no serán porcentaje

    data
        .filter(row => row.some(cell => cell !== "" && cell !== null))
        .forEach((row, rowIndex) => {
            const tr = document.createElement('tr');

            // Agregamos colIndex para saber en qué columna estamos
            row.forEach((cell, colIndex) => {
                const td = document.createElement(rowIndex === 0 ? 'th' : 'td');

                // Si es la fila de encabezado (rowIndex === 0), solo ponemos el texto
                if (rowIndex === 0) {
                    td.textContent = cell;
                } 
                // Si es un número y NO está en la lista de exclusión:
                else if (typeof cell === "number" && !columnasSoloNumeros.includes(colIndex)) {
                    const porcentaje = cell * 100;
                    td.textContent = porcentaje.toFixed(2) + " %";

                    if (porcentaje >= 100) {
                        td.style.backgroundColor = "#c8e6c9"; // verde
                    } else if (porcentaje >= 90) {
                        td.style.backgroundColor = "#fff3cd"; // amarillo
                    } else {
                        td.style.backgroundColor = "#f8d7da"; // rojo
                    }

                    td.style.textAlign = "center";
                    td.style.fontWeight = "bold";
                } 
                // Resto de celdas (texto o números excluidos)
                else {
                    td.textContent = cell;
                }

                tr.appendChild(td);
            });

            table.appendChild(tr);
        });
}

// ==============================
// 3. CREAR GRÁFICO META VS REAL
// ==============================
function crearGrafico(data) {
    const headers = data[0];
    const selectMes = document.getElementById('mes');

    // Detectar columnas que contengan REALIZADO
    const meses = headers
        .map((h, i) => ({ nombre: h, index: i }))
        .filter(h => h.nombre && h.nombre.toUpperCase().includes("REALIZADO"));

    if (meses.length === 0) {
        console.error("No se encontraron columnas REALIZADO");
        return;
    }

    // Llenar combo
    selectMes.innerHTML = "";
    meses.forEach(m => {
        const opt = document.createElement('option');
        opt.value = m.index;
        opt.textContent = m.nombre.replace(/REALIZADO/i, '').trim();
        selectMes.appendChild(opt);
    });

    // Gráfico inicial
    dibujarGraficoPorMes(data, meses[0].index, meses[0].nombre);

    selectMes.onchange = function () {
        const idx = parseInt(this.value);
        const label = this.options[this.selectedIndex].text;
        dibujarGraficoPorMes(data, idx, label);
    };
}

function dibujarGraficoPorMes(data, colReal, mesNombre) {
    const headers = data[0];
    const colUsuario = headers.indexOf("USUARIO");
    const colMeta = headers.indexOf("META");

    const usuarios = [];
    const metas = [];
    const realizados = [];

    for (let i = 1; i < data.length; i++) {
        if (data[i][colUsuario]) {
            usuarios.push(data[i][colUsuario]);
            metas.push(Number(data[i][colMeta]) || 0);
            realizados.push(Number(data[i][colReal]) || 0);
        }
    }

    const canvas = document.getElementById('graficoCumplimiento');
    canvas.remove();
    const nuevoCanvas = document.createElement('canvas');
    nuevoCanvas.id = 'graficoCumplimiento';
    document.querySelector('#contenedorGrafico').appendChild(nuevoCanvas);

    new Chart(nuevoCanvas, {
        type: 'bar',
        data: {
            labels: usuarios,
            datasets: [
                { label: 'Meta', data: metas, backgroundColor: '#A7C7E7' },
                { label: 'Realizado ' + mesNombre, data: realizados, backgroundColor: '#004b87' }
            ]
        }
    });
}
// ==============================
// 4. CONTADOR VISIBLE DE VISITAS
// ==============================
let visitas = localStorage.getItem('contadorVisitas');

if (!visitas) {
    visitas = 1;
} else {
    visitas = parseInt(visitas) + 1;
}

localStorage.setItem('contadorVisitas', visitas);

const spanVisitas = document.getElementById('visitas');
if (spanVisitas) {
    spanVisitas.textContent = visitas;
}
