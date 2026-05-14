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


const TIPO_CAMPO_TESAURO_SALIDA = "Texto";

const crearCamposAsistencia = (nombreBase) => [
    { nombre: `${nombreBase} (Sí/No)`, valor: "Sí" },
    { nombre: `Fecha y hora de ${nombreBase}`, valor: "__FECHA__" }
];

const MAPEO_NOMBRE_TAREA_CSV = {
    "agg kickoff": "Certificado de asistencia a sesión online KickOff",
    "taller 01": `Certificado asistencia a taller online 01 - "Condición de usuario apoderado"`,
    "taller 02": `Certificado asistencia a taller online 02 - "Contextualización y configuración del catálogo"`,
    "taller 03": `Certificado asistencia a taller online 03 - "Taller de campos personalizados del tesauro"`,
    "taller 04": `Certificado asistencia a taller online 04 - "Construcción de documentos inteligentes"`,
    "taller 05": `Certificado asistencia a taller online 05 - "Configuración de circuitos de resolución singulares"`,
    "taller 06": `Certificado asistencia a taller online 06 - "Configuración de circuitos de resolución plural"`,
    "taller 07": `Certificado asistencia a taller online 07 - "Simplificación administrativa"`,
    "taller 08": `Certificado asistencia a taller online 08 - "Módulo de diseño de Control Interno"`,
    "taller 09": `Certificado asistencia a taller online 09 - "Circuitos de resolución con gasto"`,
    "taller 10": `Certificado asistencia a taller online 10 - "Búsquedas avanzadas"`,
    "taller 11": `Certificado asistencia a taller online 11 - "Analiza. Conceptos básicos y configuración"`,
    "add kickoff": "Certificado de asistencia a sesión online KickOff Analiza",
    "add taller01": "Certificado asistencia a Taller 01: Representación de la información y configuración básica",
    "add taller02": "Certificado asistencia a Taller 02: Configuración de dimensiones personalizadas y operaciones con fechas",
    "add taller03": "Certificado asistencia a Taller 03 Funciones avanzadas y variables",
    "gfd kickoff": "Certificado de asistencia a sesión online KickOff Developers",
    "gfd taller01": `Certificado asistencia a taller online 01 - "Autorización + Creación Expediente"`,
    "gfd taller02": `Certificado asistencia a taller online 02 - "Gestiona Code y Tramitación reglada"`,
    "gfd taller03": `Certificado asistencia a taller online 03 - "Carga de archivos + Gestión de expediente"`,
    "gfd taller04": `Certificado asistencia a taller online 04 - "Terceros"`,
    "gfd taller05": `Certificado asistencia a taller online 05 - "Registro de entradas + Listado y paginado"`,
    "gfd taller06": `Certificado asistencia a taller online 06 - "Tramitación y registro de salida"`,
    "gfd taller07": `Certificado asistencia a taller online 07 - "Bus de Eventos y Conector externo"`,
    "gfd taller08": `Certificado asistencia a taller online 08 - "Operaciones externas"`
};

const crearConfiguracionTarea = (certificado) => ({
    certificado,
    campos: crearCamposAsistencia(certificado.replace(/^Certificado de asistencia a |^Certificado asistencia a /, ""))
});

// ======================================================================
//  MAPEO NUEVO — GFD / ADD  (NORMALIZADO)
// ======================================================================

const MAPEO_INPUT_TALLERES = Object.fromEntries(
    Object.entries(MAPEO_NOMBRE_TAREA_CSV).map(([input, certificado]) => [
        input,
        crearConfiguracionTarea(certificado)
    ])
);

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
                    TipoCampoTesauro: TIPO_CAMPO_TESAURO_SALIDA,
                    ValorCampo: campo.valor === "__FECHA__" ? fecha : campo.valor,
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
            TipoCampoTesauro: TIPO_CAMPO_TESAURO_SALIDA,
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
            TipoCampoTesauro: TIPO_CAMPO_TESAURO_SALIDA,
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

    const columnasCsv = [
        { key: "NombreEntidad", label: "Nombre entidad" },
        { key: "CódigoExpediente", label: "Código expediente" },
        { key: "NombreTarea", label: "Nombre tarea" },
        { key: "CrearTarea", label: "Crear tarea" },
        { key: "NombreCampoCastellano", label: "Nombre campo castellano" },
        { key: "TipoCampoTesauro", label: "Tipo campo tesauro" },
        { key: "ValorCampo", label: "Valor campo" },
        { key: "ValorCampoAdicional", label: "Valor campo adicional" },
        { key: "NIFTercero", label: "NIF Tercero" }
    ];
    const separador = ";";
    const saltoLinea = "\r\n";
    const escaparValor = (valor) => {
        const texto = valor ?? "";
        const textoStr = texto.toString();
        const necesitaComillas = textoStr.includes(separador) ||
            textoStr.includes('"') ||
            textoStr.includes("\n") ||
            textoStr.includes("\r");
        const textoEscapado = textoStr.replace(/"/g, '""');
        return necesitaComillas ? `"${textoEscapado}"` : textoEscapado;
    };

    const filas = [
        columnasCsv.map(col => escaparValor(col.label)).join(separador),
        ...datosSalida.map(row =>
            columnasCsv.map(col => escaparValor(row[col.key])).join(separador)
        )
    ];
    const csv = filas.join(saltoLinea);
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
