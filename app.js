// ======================================================================
//  MAPEO EXACTO DE TESAUROS  (texto largo que va en la columna "Campo")
// ======================================================================

const MAPEO_TESAUROS = {
    'Asistencia a taller online "Búsquedas avanzadas"': "Taller 09",
    'Asistencia a taller online "Analiza. Conceptos básicos y configuración"': "Taller 10",
    'Asistencia a taller online "Circuitos de resolución con gasto"': "Taller 08",
    'Asistencia a taller online "Módulo de diseño"': "Taller 07",
    'Asistencia a taller online "Simplificación administrativa"': "Taller 06",
    'Asistencia a taller online "Configuración de circuitos de resolución plural"': "Taller 05",
    'Asistencia a taller online "Configuración de circuitos de resolución singulares"': "Taller 04",
    'Asistencia taller online "Construcción de documentos inteligentes"': "Taller 03",
    'Asistencia a taller online "Taller de campos personalizados del tesauro"': "Taller 02",
    'Asistencia a taller online "Contextualización y configuración del catálogo"': "Taller 01",
    'Asistencia taller online "Condición de usuario apoderado"': "Taller 00",
    "Asistencia Kick Off online Developers (Sí/No)": "GFD KickOff",
    "Fecha y hora de asistencia a Kick off online Gestiona Developers": "GFD KickOff",
    "Asiste al taller online 01": "GFD Taller01",
    "Fecha y hora asistencia taller online 01": "GFD Taller01",
    "Asiste al taller online 02": "GFD Taller02",
    "Fecha y hora asistencia taller online 02": "GFD Taller02",
    "Asiste al taller online 03": "GFD Taller03",
    "Fecha y hora asistencia taller online 03": "GFD Taller03",
    "Asiste al taller online 04": "GFD Taller04",
    "Fecha y hora asistencia taller online 04": "GFD Taller04",
    "Asiste al taller online 05": "GFD Taller05",
    "Fecha y hora asistencia taller online 05": "GFD Taller05",
    "Asiste al taller online 06": "GFD Taller06",
    "Fecha y hora asistencia taller online 06": "GFD Taller06",
    "Asistencia Kick Off online Certificación Analiza (Sí/No)": "ADD Kickoff",
    'Fecha y hora de asistencia a Kick off online de "Gestiona Analiza"': "ADD Kickoff",
    "Asiste al Taller 01 Introducción y conceptos básicos de Analiza": "ADD Taller01",
    "Fecha y hora de asistencia a taller 01 del curso Analiza": "ADD Taller01",
    "Asiste al Taller 02 Edición de expresiones": "ADD Taller02",
    "Fecha y hora de asistencia a taller 02 del curso Analiza": "ADD Taller02",
    "Asiste al Taller 03 Preparación evaluación": "ADD Taller03",
    "Fecha y hora de asistencia a taller 03 del curso Analiza": "ADD Taller03"
};


// ======================================================================
//  MAPEO EXACTO DE CERTIFICADOS (columna C)
// ======================================================================

const MAPEO_CERTIFICADOS = {
    "Taller 10": `Certificado asistencia a taller online 12 - "Analiza. Conceptos básicos y configuración"`,
    "Taller 09": `Certificado asistencia a taller online 11 - "Búsquedas avanzadas"`,
    "Taller 08": `Certificado asistencia a taller online 09 - "Circuitos de resolución con gasto"`,
    "Taller 07": `Certificado asistencia a taller online 08 - "Módulo de diseño de Control Interno"`,
    "Taller 06": `Certificado asistencia a taller online 07 - "Simplificación administrativa"`,
    "Taller 05": `Certificado asistencia a taller online 06 - "Configuración de circuitos de resolución plural"`,
    "Taller 04": `Certificado asistencia a taller online 05 - "Configuración de circuitos de resolución singulares"`,
    "Taller 03": `Certificado asistencia a taller online 04 - "Construcción de documentos inteligentes"`,
    "Taller 02": `Certificado asistencia a taller online 03 - "Taller de campos personalizados del tesauro"`,
    "Taller 01": `Certificado asistencia a taller online 02 - "Contextualización y configuración del catálogo"`,
    "Taller 00": `Certificado asistencia a taller online 01 - "Condición de usuario apoderado"`,
    "GFD KickOff": "Certificado de asistencia a sesión online KickOff Developers",
    "GFD Taller01": 'Certificado asistencia a taller online 01 - "Autorización de un addon, listados paginados y filtrados, carga Archivos-Bus de Eventos y Conectores Externos"',
    "GFD Taller02": 'Certificado asistencia a taller online 02 - "Bus de Eventos y Conectores Externos"',
    "GFD Taller03": 'Certificado asistencia a taller online 03 - "Terceros"',
    "GFD Taller04": 'Certificado asistencia a taller online 04 - "Registro Electrónico"',
    "GFD Taller05": 'Certificado asistencia a taller online 05 - "Expedientes"',
    "GFD Taller06": 'Certificado asistencia a taller online 06 - "Tramitación"',
    "ADD Kickoff": "Certificado de asistencia a sesión online KickOff Analiza",
    "ADD Taller01": "Certificado asistencia a Taller 01: Representación de la información y configuración básica",
    "ADD Taller02": "Certificado asistencia a Taller 02: Configuración de dimensiones personalizadas y operaciones con fechas",
    "ADD Taller03": "Certificado asistencia a Taller 03 Funciones avanzadas y variables"
};

const MAPEO_TIPOS_TESAUROS = {
    "Asistencia Kick Off online Developers (Sí/No)": "Sí/No",
    "Asiste al taller online 01": "Sí/No",
    "Asiste al taller online 02": "Sí/No",
    "Asiste al taller online 03": "Sí/No",
    "Asiste al taller online 04": "Sí/No",
    "Asiste al taller online 05": "Sí/No",
    "Asiste al taller online 06": "Sí/No",
    "Asistencia Kick Off online Certificación Analiza (Sí/No)": "Sí/No",
    "Asiste al Taller 01 Introducción y conceptos básicos de Analiza": "Sí/No",
    "Asiste al Taller 02 Edición de expresiones": "Sí/No",
    "Asiste al Taller 03 Preparación evaluación": "Sí/No"
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
        const codigoEntrada = row["Taller"]
            || row["Nombre input"]
            || row["Nombre Input"];

        if (!codigoEntrada) return;

        const codigo = codigoEntrada.trim();

        // === Valores origen ===
        const dni = row["InteresadoIdentificador"];
        const exp = row["ExpedienteCodigo"];
        const fecha = row["Fecha -Hora"];

        if (codigo.startsWith("Taller")) {
            // === Buscar el tesauro correcto ===
            const entradaTesauro = Object.entries(MAPEO_TESAUROS)
                .find(([texto, codigoMapa]) => codigoMapa === codigo);

            if (!entradaTesauro) {
                console.warn("Taller NO encontrado:", codigo);
                return;
            }

            const tesauroTexto = entradaTesauro[0];
            const tesauroSN = tesauroTexto + " (Sí/No)";


            // === CERTIFICADO (columna C) ===
            const certificadoTexto = MAPEO_CERTIFICADOS[codigo];


            // ==================================================
            //  FILA 1 — Tipo Fecha (Taller 08 → Texto)
            // ==================================================
            datosSalida.push({
                NombreEntidad: "ESPUBLICO SERVICIOS PARA LA ADMINISTRACIÓN",
                CódigoExpediente: exp,
                NombreTarea: certificadoTexto,
                CrearTarea: "Sí",
                NombreCampoCastellano: tesauroTexto,
                TipoCampoTesauro: (codigo === "Taller 08" ? "Texto" : "Fecha"),
                ValorCampo: fecha,
                ValorCampoAdicional: "",
                NIFTercero: dni
            });


            // ==================================================
            //  FILA 2 — Tipo Sí/No
            // ==================================================
            datosSalida.push({
                NombreEntidad: "ESPUBLICO SERVICIOS PARA LA ADMINISTRACIÓN",
                CódigoExpediente: exp,
                NombreTarea: certificadoTexto,
                CrearTarea: "Sí",
                NombreCampoCastellano: tesauroSN,
                TipoCampoTesauro: "Sí/No",
                ValorCampo: "Sí",
                ValorCampoAdicional: "",
                NIFTercero: dni
            });

            return;
        }

        // === Nuevos casos GFD/ADD ===
        const tesauros = Object.entries(MAPEO_TESAUROS)
            .filter(([, codigoMapa]) => codigoMapa === codigo);

        if (tesauros.length === 0) {
            console.warn("Entrada NO encontrada:", codigo);
            return;
        }

        const certificadoTexto = MAPEO_CERTIFICADOS[codigo];

        tesauros.forEach(([tesauroTexto]) => {
            const tipoTesauro = MAPEO_TIPOS_TESAUROS[tesauroTexto] || "Fecha";
            const valorCampo = tipoTesauro === "Sí/No" ? "Sí" : fecha;

            datosSalida.push({
                NombreEntidad: "ESPUBLICO SERVICIOS PARA LA ADMINISTRACIÓN",
                CódigoExpediente: exp,
                NombreTarea: certificadoTexto,
                CrearTarea: "Sí",
                NombreCampoCastellano: tesauroTexto,
                TipoCampoTesauro: tipoTesauro,
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
