// ======================================================================
//  MAPEO EXACTO DE ENTRADAS, TESAUROS Y CERTIFICADOS
// ======================================================================

const MAPEO_ENTRADAS = {
    "GFD KickOff": {
        certificado: "Certificado de asistencia a sesión online KickOff Developers",
        campos: [
            {
                nombre: "Asistencia Kick Off online Developers (Sí/No)",
                tipo: "Sí/No",
                valor: "Sí"
            },
            {
                nombre: "Fecha y hora de asistencia a Kick off online Gestiona Developers",
                tipo: "Fecha"
            }
        ]
    },
    "GFD Taller01": {
        certificado: 'Certificado asistencia a taller online 01 - "Autorización de un addon, listados paginados y filtrados, carga Archivos-Bus de Eventos y Conectores Externos"',
        campos: [
            {
                nombre: "Asiste al taller online 01",
                tipo: "Sí/No",
                valor: "Sí"
            },
            {
                nombre: "Fecha y hora asistencia taller online 01",
                tipo: "Fecha"
            }
        ]
    },
    "GFD Taller02": {
        certificado: 'Certificado asistencia a taller online 02 - "Bus de Eventos y Conectores Externos"',
        campos: [
            {
                nombre: "Asiste al taller online 02",
                tipo: "Sí/No",
                valor: "Sí"
            },
            {
                nombre: "Fecha y hora asistencia taller online 02",
                tipo: "Fecha"
            }
        ]
    },
    "GFD Taller03": {
        certificado: 'Certificado asistencia a taller online 03 - "Terceros"',
        campos: [
            {
                nombre: "Asiste al taller online 03",
                tipo: "Sí/No",
                valor: "Sí"
            },
            {
                nombre: "Fecha y hora asistencia taller online 03",
                tipo: "Fecha"
            }
        ]
    },
    "GFD Taller04": {
        certificado: 'Certificado asistencia a taller online 04 - "Registro Electrónico"',
        campos: [
            {
                nombre: "Asiste al taller online 04",
                tipo: "Sí/No",
                valor: "Sí"
            },
            {
                nombre: "Fecha y hora asistencia taller online 04",
                tipo: "Fecha"
            }
        ]
    },
    "GFD Taller05": {
        certificado: 'Certificado asistencia a taller online 05 - "Expedientes"',
        campos: [
            {
                nombre: "Asiste al taller online 05",
                tipo: "Sí/No",
                valor: "Sí"
            },
            {
                nombre: "Fecha y hora asistencia taller online 05",
                tipo: "Fecha"
            }
        ]
    },
    "GFD Taller06": {
        certificado: 'Certificado asistencia a taller online 06 - "Tramitación"',
        campos: [
            {
                nombre: "Asiste al taller online 06",
                tipo: "Sí/No",
                valor: "Sí"
            },
            {
                nombre: "Fecha y hora asistencia taller online 06",
                tipo: "Fecha"
            }
        ]
    },
    "ADD Kickoff": {
        certificado: "Certificado de asistencia a sesión online KickOff Analiza",
        campos: [
            {
                nombre: "Asistencia Kick Off online Certificación Analiza (Sí/No)",
                tipo: "Sí/No",
                valor: "Sí"
            },
            {
                nombre: 'Fecha y hora de asistencia a Kick off online de "Gestiona Analiza"',
                tipo: "Fecha"
            }
        ]
    },
    "ADD Taller01": {
        certificado: "Certificado asistencia a Taller 01: Representación de la información y configuración básica",
        campos: [
            {
                nombre: "Asiste al Taller 01 Introducción y conceptos básicos de Analiza",
                tipo: "Sí/No",
                valor: "Sí"
            },
            {
                nombre: "Fecha y hora de asistencia a taller 01 del curso Analiza",
                tipo: "Fecha"
            }
        ]
    },
    "ADD Taller02": {
        certificado: "Certificado asistencia a Taller 02: Configuración de dimensiones personalizadas y operaciones con fechas",
        campos: [
            {
                nombre: "Asiste al Taller 02 Edición de expresiones",
                tipo: "Sí/No",
                valor: "Sí"
            },
            {
                nombre: "Fecha y hora de asistencia a taller 02 del curso Analiza",
                tipo: "Fecha"
            }
        ]
    },
    "ADD Taller03": {
        certificado: "Certificado asistencia a Taller 03 Funciones avanzadas y variables",
        campos: [
            {
                nombre: "Asiste al Taller 03 Preparación evaluación",
                tipo: "Sí/No",
                valor: "Sí"
            },
            {
                nombre: "Fecha y hora de asistencia a taller 03 del curso Analiza",
                tipo: "Fecha"
            }
        ]
    }
};


// ======================================================================
//  VARIABLES
// ======================================================================

let datosOrigen = [];
let datosSalida = [];


// ======================================================================
//  LEER N EXCELS
// ======================================================================

document.getElementById("inputFile").addEventListener("change", async (e) => {
    const files = Array.from(e.target.files);
    datosOrigen = [];

    for (const file of files) {
        const content = await file.arrayBuffer();
        const workbook = XLSX.read(content, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet);
        datosOrigen.push(...json);
    }

    alert(`Se han cargado ${files.length} archivo(s). Total filas: ${datosOrigen.length}`);
});


// ======================================================================
//  PROCESAR TODO
// ======================================================================

document.getElementById("btnProcesar").onclick = () => {
    if (datosOrigen.length === 0) {
        alert("Primero sube uno o varios archivos Excel.");
        return;
    }
    procesar();
};


// ======================================================================
//  TRANSFORMAR 1 FILA ORIGEN → 2 FILAS SALIDA (A–G)
// ======================================================================

function procesar() {
    datosSalida = [];

    datosOrigen.forEach(row => {

        // === Columna de entrada ===
        const nombreInput = (
            row["Nombre input"] ||
            row["Nombre Input"] ||
            row["Taller"]
        );
        if (!nombreInput) return;

        const entrada = MAPEO_ENTRADAS[nombreInput.trim()];

        if (!entrada) {
            console.warn("Entrada NO encontrada:", nombreInput);
            return;
        }


        // === Valores origen ===
        const dni = row["InteresadoIdentificador"];
        const exp = row["ExpedienteCodigo"];
        const fecha = row["Fecha -Hora"];
        entrada.campos.forEach((campo) => {
            const valorCampo = campo.tipo === "Fecha" ? fecha : campo.valor;

            datosSalida.push({
                NombreEntidad: "ESPUBLICO SERVICIOS PARA LA ADMINISTRACIÓN",
                CódigoExpediente: exp,
                NombreTarea: entrada.certificado,
                CrearTarea: "Sí",
                NombreCampoCastellano: campo.nombre,
                TipoCampoTesauro: campo.tipo,
                ValorCampo: valorCampo,
                ValorCampoAdicional: "",
                NIFTercero: dni
            });
        });
    });

    mostrarTabla();
}


// ======================================================================
//  MOSTRAR TABLA EN HTML
// ======================================================================

function mostrarTabla() {
    const div = document.getElementById("resultado");
    div.innerHTML = "";

    if (datosSalida.length === 0) {
        div.innerHTML = "<p>No hay datos procesados.</p>";
        return;
    }

    const tabla = document.createElement("table");
    const cols = Object.keys(datosSalida[0]);

    let thead = "<tr>";
    cols.forEach(c => thead += `<th>${c}</th>`);
    thead += "</tr>";

    let tbody = "";
    datosSalida.forEach(r => {
        tbody += "<tr>";
        cols.forEach(c => tbody += `<td>${r[c]}</td>`);
        tbody += "</tr>";
    });

    tabla.innerHTML = thead + tbody;
    div.appendChild(tabla);
}


// ======================================================================
//  EXPORTAR XLSX FINAL
// ======================================================================

document.getElementById("btnDescargar").onclick = () => {
    if (datosSalida.length === 0) {
        alert("Nada que descargar.");
        return;
    }

    const ws = XLSX.utils.json_to_sheet(datosSalida);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Salida");

    XLSX.writeFile(wb, "Salida.xlsx");
};
