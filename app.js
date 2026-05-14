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

const MAPEO_NOMBRE_TAREA_POR_PROGRAMA = {
    CAAG: {
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
        "taller 11": `Certificado asistencia a taller online 11 - "Analiza. Conceptos básicos y configuración"`
    },
    CAZ: {
        "kickoff": "Certificado de asistencia a sesión online KickOff Analiza",
        "taller 01": "Certificado asistencia a Taller 01: Representación de la información y configuración básica",
        "taller 02": "Certificado asistencia a Taller 02: Configuración de dimensiones personalizadas y operaciones con fechas",
        "taller 03": "Certificado asistencia a Taller 03 Funciones avanzadas y variables"
    },
    GFD: {
        "kickoff": "Certificado de asistencia a sesión online KickOff Developers",
        "taller 01": `Certificado asistencia a taller online 01 - "Autorización + Creación Expediente"`,
        "taller 02": `Certificado asistencia a taller online 02 - "Gestiona Code y Tramitación reglada"`,
        "taller 03": `Certificado asistencia a taller online 03 - "Carga de archivos + Gestión de expediente"`,
        "taller 04": `Certificado asistencia a taller online 04 - "Terceros"`,
        "taller 05": `Certificado asistencia a taller online 05 - "Registro de entradas + Listado y paginado"`,
        "taller 06": `Certificado asistencia a taller online 06 - "Tramitación y registro de salida"`,
        "taller 07": `Certificado asistencia a taller online 07 - "Bus de Eventos y Conector externo"`,
        "taller 08": `Certificado asistencia a taller online 08 - "Operaciones externas"`
    }
};

const crearConfiguracionTarea = (certificado) => ({
    certificado,
    campos: crearCamposAsistencia(certificado.replace(/^Certificado de asistencia a |^Certificado asistencia a /, ""))
});

// ======================================================================
//  MAPEO NUEVO POR PROGRAMA (NORMALIZADO)
// ======================================================================

const MAPEO_INPUT_TALLERES_POR_PROGRAMA = Object.fromEntries(
    Object.entries(MAPEO_NOMBRE_TAREA_POR_PROGRAMA).map(([programa, mapeo]) => [
        programa,
        Object.fromEntries(
            Object.entries(mapeo).map(([input, certificado]) => [
                input,
                crearConfiguracionTarea(certificado)
            ])
        )
    ])
);

// ======================================================================
//  VARIABLES
// ======================================================================

let datosOrigen = [];
let datosSalida = [];
let programaSeleccionado = null;
let avisosDatosFaltantes = [];
let avisosMapeo = [];

const CAMPOS_OBLIGATORIOS_INPUT = ["Taller", "InteresadoIdentificador", "ExpedienteCodigo", "Fecha -Hora"];
const COLUMNAS_IGNORADAS_DETECCION_ERRORES = ["ValorCampoAdicional"];
const PROGRAMAS_CERTIFICACION = ["CAAG", "GFD", "CAZ"];


function normalizarValor(valor) {
    return (valor ?? "").toString().trim().toLowerCase();
}

function tieneDato(valor) {
    return valor !== undefined && valor !== null && valor.toString().trim() !== "";
}

function obtenerCamposFaltantes(row) {
    return CAMPOS_OBLIGATORIOS_INPUT.filter(campo => !tieneDato(row[campo]));
}


function completarEncabezadoExpedienteCodigo(ws) {
    const celdaColumnaC = "C1";
    const encabezado = ws[celdaColumnaC]?.v;

    if (!tieneDato(encabezado)) {
        ws[celdaColumnaC] = { t: "s", v: "ExpedienteCodigo" };
    }
}

function obtenerProgramaCertificacionSeleccionado() {
    const selector = document.getElementById("programaCertificacion");
    const programa = selector.value;

    if (!PROGRAMAS_CERTIFICACION.includes(programa)) {
        alert("Selecciona un programa de certificación antes de cargar el Excel.");
        selector.focus();
        return null;
    }

    return programa;
}

function escaparHtml(valor) {
    return (valor ?? "").toString()
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

// ======================================================================
//  LEER EXCEL
// ======================================================================

document.getElementById("inputFile").addEventListener("change", async (e) => {
    const files = Array.from(e.target.files);
    if (files.length === 0) return;

    const programa = obtenerProgramaCertificacionSeleccionado();
    if (!programa) {
        e.target.value = "";
        return;
    }

    datosOrigen = [];
    datosSalida = [];
    avisosDatosFaltantes = [];
    avisosMapeo = [];
    programaSeleccionado = programa;

    for (const file of files) {
        const buffer = await file.arrayBuffer();
        const wb = XLSX.read(buffer, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        completarEncabezadoExpedienteCodigo(ws);
        const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
        datosOrigen.push(...json);
    }

    document.getElementById("programaSeleccionado").textContent = `Programa seleccionado: ${programaSeleccionado}`;
    document.getElementById("avisos").innerHTML = "";
    document.getElementById("resultado").innerHTML = "";
    alert(`Cargadas ${datosOrigen.length} filas para el programa ${programaSeleccionado}`);
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

    avisosDatosFaltantes = [];
    avisosMapeo = [];

    const mapeoPrograma = MAPEO_INPUT_TALLERES_POR_PROGRAMA[programaSeleccionado] ?? {};

    datosOrigen.forEach((row, index) => {

        const filaExcel = index + 2;
        const camposFaltantes = obtenerCamposFaltantes(row);
        if (camposFaltantes.length > 0) {
            avisosDatosFaltantes.push({ fila: filaExcel, campos: camposFaltantes, row });
        }

        if (!tieneDato(row["Taller"])) {
            return;
        }

        const valorTallerRaw = row["Taller"].toString().trim();
        const valorTallerNorm = normalizarValor(valorTallerRaw);

        const dni = row["InteresadoIdentificador"];
        const exp = row["ExpedienteCodigo"];
        const fecha = row["Fecha -Hora"];

        // ==================================================
        //  MODO POR PROGRAMA (CAAG / CAZ / GFD)
        // ==================================================
        if (mapeoPrograma[valorTallerNorm]) {

            const cfg = mapeoPrograma[valorTallerNorm];

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

        if (programaSeleccionado) {
            avisosMapeo.push({
                fila: filaExcel,
                taller: valorTallerRaw,
                programa: programaSeleccionado
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
    mostrarAvisos();

    const div = document.getElementById("resultado");
    div.innerHTML = "";

    if (datosSalida.length === 0) {
        div.innerHTML = "<p>No hay datos procesados</p>";
        return;
    }

    const tabla = document.createElement("table");
    const cols = Object.keys(datosSalida[0]);
    const columnaIgnorada = (columna) => COLUMNAS_IGNORADAS_DETECCION_ERRORES.includes(columna);
    const filaConErrores = (row) => cols.some(columna => !columnaIgnorada(columna) && !tieneDato(row[columna]));

    tabla.innerHTML =
        "<tr>" + cols.map(c => `<th>${escaparHtml(c)}</th>`).join("") + "</tr>" +
        datosSalida.map(r =>
            `<tr${filaConErrores(r) ? ' class="row-missing"' : ""}>` +
            cols.map(c => `<td>${escaparHtml(r[c])}</td>`).join("") +
            "</tr>"
        ).join("");

    div.appendChild(tabla);
}

function mostrarAvisos() {
    const div = document.getElementById("avisos");
    div.innerHTML = "";

    if (avisosDatosFaltantes.length === 0 && avisosMapeo.length === 0) {
        return;
    }

    const partes = [];

    if (avisosDatosFaltantes.length > 0) {
        const columnas = CAMPOS_OBLIGATORIOS_INPUT;
        partes.push(`
            <div class="alert alert-warning">
                <strong>Advertencia:</strong> se detectaron ${avisosDatosFaltantes.length} filas con datos faltantes.
                Las filas con cualquier campo obligatorio vacío se muestran con fondo rojo.
            </div>
            <div class="table-wrapper">
                <table>
                    <tr><th>Fila Excel</th>${columnas.map(c => `<th>${escaparHtml(c)}</th>`).join("")}</tr>
                    ${avisosDatosFaltantes.map(aviso => `
                        <tr class="row-missing">
                            <td>${aviso.fila}</td>
                            ${columnas.map(c => `<td>${escaparHtml(aviso.row[c])}</td>`).join("")}
                        </tr>
                    `).join("")}
                </table>
            </div>
        `);
    }

    if (avisosMapeo.length > 0) {
        partes.push(`
            <div class="alert alert-warning">
                <strong>Advertencia:</strong> ${avisosMapeo.length} filas tienen un taller sin mapeo para el programa ${escaparHtml(programaSeleccionado)}.
            </div>
            <ul class="warning-list">
                ${avisosMapeo.map(aviso => `<li>Fila ${aviso.fila}: "${escaparHtml(aviso.taller)}" no existe en ${escaparHtml(aviso.programa)}.</li>`).join("")}
            </ul>
        `);
    }

    div.innerHTML = partes.join("");
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
