// ======================================================================
//  MAPEO TESAUROS — MODO ANTIGUO (NO TOCAR)
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
//  MAPEO NUEVO — GFD / ADD  (NORMALIZADO)
// ======================================================================

const MAPEO_INPUT_TALLERES = {
    "gfd kickoff": {
        certificado: "Certificado de asistencia a sesión online KickOff Developers",
        campos: [
            { nombre: "Asistencia Kick Off online Developers (Sí/No)", tipo: "Sí/No" },
            { nombre: "Fecha y hora de asistencia a Kick off online Gestiona Developers", tipo: "Fecha" }
        ]
    },
    "add kickoff": {
        certificado: "Certificado de asistencia a sesión online KickOff Analiza",
        campos: [
            { nombre: "Asistencia Kick Off online Certificación Analiza (Sí/No)", tipo: "Sí/No" },
            { nombre: "Fecha y hora de asistencia a Kick off online de \"Gestiona Analiza\"", tipo: "Fecha" }
        ]
    }
};

// ======================================================================
//  VARIABLES
// ======================================================================

let datosOrigen = [];
let datosSalida = [];

// ======================================================================
//  LEER EXCEL
// ======================================================================

document.getElementById("inputFile").addEventListener("change", async (e) => {
    const files = Array.from(e.target.files);
    datosOrigen = [];

    for (const file of files) {
        const buffer = await file.arrayBuffer();
        const wb = XLSX.read(buffer, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws);
        datosOrigen.push(...json);
    }

    alert(`Cargadas ${datosOrigen.length} filas`);
});

// ======================================================================
//  BOTÓN PROCESAR
// ======================================================================

document.getElementById("btnProcesar").onclick = () => {
    if (datosOrigen.length === 0) {
        alert("Primero carga un Excel");
        return;
    }
    procesar();
};

// ======================================================================
//  PROCESAR
// ======================================================================

function procesar() {
    datosSalida = [];

    datosOrigen.forEach(row => {

        if (!row["Taller"]) return;

        const valorTallerRaw = row["Taller"].toString().trim();
        const valorTallerNorm = valorTallerRaw.toLowerCase();

        const dni = row["InteresadoIdentificador"];
        const exp = row["ExpedienteCodigo"];
        const fecha = row["Fecha -Hora"];

        // ==================================================
        //  MODO NUEVO (GFD / ADD)
        // ==================================================
        if (MAPEO_INPUT_TALLERES[valorTallerNorm]) {

            const cfg = MAPEO_INPUT_TALLERES[valorTallerNorm];

            cfg.campos.forEach(campo => {
                datosSalida.push({
                    NombreEntidad: "ESPUBLICO SERVICIOS PARA LA ADMINISTRACIÓN",
                    CódigoExpediente: exp,
                    NombreTarea: cfg.certificado,
                    CrearTarea: "Sí",
                    NombreCampoCastellano: campo.nombre,
                    TipoCampoTesauro: campo.tipo,
                    ValorCampo: campo.tipo === "Sí/No" ? "Sí" : fecha,
                    ValorCampoAdicional: "",
                    NIFTercero: dni
                });
            });

            return;
        }

        // ==================================================
        //  MODO ANTIGUO (Taller XX)
        // ==================================================

        const entradaTesauro = Object.entries(MAPEO_TESAUROS)
            .find(([_, codigo]) => codigo === valorTallerRaw);

        if (!entradaTesauro) return;

        const tesauroTexto = entradaTesauro[0];
        const tesauroSN = tesauroTexto + " (Sí/No)";
        const certificadoTexto = MAPEO_CERTIFICADOS[valorTallerRaw];

        datosSalida.push({
            NombreEntidad: "ESPUBLICO SERVICIOS PARA LA ADMINISTRACIÓN",
            CódigoExpediente: exp,
            NombreTarea: certificadoTexto,
            CrearTarea: "Sí",
            NombreCampoCastellano: tesauroTexto,
            TipoCampoTesauro: "Fecha",
            ValorCampo: fecha,
            ValorCampoAdicional: "",
            NIFTercero: dni
        });

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
//  MOSTRAR TABLA
// ======================================================================

function mostrarTabla() {
    const div = document.getElementById("resultado");
    div.innerHTML = "";

    if (datosSalida.length === 0) {
        div.innerHTML = "<p>No hay datos procesados</p>";
        return;
    }

    const tabla = document.createElement("table");
    const cols = Object.keys(datosSalida[0]);

    tabla.innerHTML =
        "<tr>" + cols.map(c => `<th>${c}</th>`).join("") + "</tr>" +
        datosSalida.map(r =>
            "<tr>" + cols.map(c => `<td>${r[c]}</td>`).join("") + "</tr>"
        ).join("");

    div.appendChild(tabla);
}

// ======================================================================
//  DESCARGAR XLSX
// ======================================================================

document.getElementById("btnDescargar").onclick = () => {
    if (datosSalida.length === 0) {
        alert("Nada que descargar");
        return;
    }

    const ws = XLSX.utils.json_to_sheet(datosSalida);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Salida");
    XLSX.writeFile(wb, "Salida.xlsx");
};

// ======================================================================
//  DESCARGAR CSV UTF-8
// ======================================================================

document.getElementById("btnDescargarCsv").onclick = () => {
    if (datosSalida.length === 0) {
        alert("Nada que descargar");
        return;
    }

    const ws = XLSX.utils.json_to_sheet(datosSalida);
    const csv = XLSX.utils.sheet_to_csv(ws);
    const csvConBOM = "\uFEFF" + csv;
    const blob = new Blob([csvConBOM], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const enlace = document.createElement("a");
    enlace.href = url;
    enlace.download = "Salida.csv";
    document.body.appendChild(enlace);
    enlace.click();
    document.body.removeChild(enlace);
    URL.revokeObjectURL(url);
};
