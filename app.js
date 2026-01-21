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
    'Asistencia taller online "Condición de usuario apoderado"': "Taller 00"
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
    "Taller 00": `Certificado asistencia a taller online 01 - "Condición de usuario apoderado"`
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

        // === Columna G del origen (Taller XX) ===
        let codigoTaller = row["Taller"];
        if (!codigoTaller) return;

        codigoTaller = codigoTaller.trim();  // Ej: "Taller 10"


        // === Buscar el tesauro correcto ===
        const entradaTesauro = Object.entries(MAPEO_TESAUROS)
            .find(([texto, codigo]) => codigo === codigoTaller);

        if (!entradaTesauro) {
            console.warn("Taller NO encontrado:", codigoTaller);
            return;
        }

        const tesauroTexto = entradaTesauro[0];
        const tesauroSN = tesauroTexto + " (Sí/No)";


        // === CERTIFICADO (columna C) ===
        const certificadoTexto = MAPEO_CERTIFICADOS[codigoTaller];


        // === Valores origen ===
        const dni = row["InteresadoIdentificador"];
        const exp = row["ExpedienteCodigo"];
        const fecha = row["Fecha -Hora"];
        const nombre = row["Tercero"];


// ==================================================
//  FILA 1 — Tipo Fecha (Taller 08 → Texto)
// ==================================================
datosSalida.push({
    NombreEntidad: "ESPUBLICO SERVICIOS PARA LA ADMINISTRACIÓN",
    CódigoExpediente: exp,
    NombreTarea: certificadoTexto,
    CrearTarea: "Sí",
    NombreCampoCastellano: tesauroTexto,
    TipoCampoTesauro: (codigoTaller === "Taller 08" ? "Texto" : "Fecha"),
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
