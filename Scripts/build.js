const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, AlignmentType, HeadingLevel, ImageRun,
  ShadingType, PageOrientation, VerticalAlign, Header, Footer,
  PageNumber, TabStopType, TabStopPosition,
} = require("docx");

const data = JSON.parse(fs.readFileSync("/tmp/aval/data.json", "utf-8"));
const studies = data.studies;
const agg = data.agg;

const FONT = "Arial";
const NOBORDER = {
  top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  insideVertical: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
};
const THINBORDER = {
  top: { style: BorderStyle.SINGLE, size: 4, color: "999999" },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: "999999" },
  left: { style: BorderStyle.SINGLE, size: 4, color: "999999" },
  right: { style: BorderStyle.SINGLE, size: 4, color: "999999" },
  insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "999999" },
  insideVertical: { style: BorderStyle.SINGLE, size: 4, color: "999999" },
};

function p(text, opts = {}) {
  const { bold = false, italic = false, size = 22, before = 0, after = 120,
           align = AlignmentType.JUSTIFIED, firstLine = 0, color } = opts;
  return new Paragraph({
    spacing: { before, after, line: 300, lineRule: "auto" },
    alignment: align,
    indent: firstLine ? { firstLine } : undefined,
    children: [new TextRun({ text, bold, italics: italic, size, font: FONT, color })],
  });
}

function heading(text, opts = {}) {
  const { size = 28, before = 240, after = 120 } = opts;
  return new Paragraph({
    spacing: { before, after },
    children: [new TextRun({ text, bold: true, size, font: FONT })],
  });
}

function cell(text, { width, bold = false, italic = false, size = 18, align = AlignmentType.LEFT,
                       shade, valign = VerticalAlign.CENTER, colSpan } = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    verticalAlign: valign,
    columnSpan: colSpan,
    shading: shade ? { type: ShadingType.CLEAR, fill: shade } : undefined,
    margins: { top: 40, bottom: 40, left: 60, right: 60 },
    children: [new Paragraph({
      alignment: align,
      spacing: { before: 0, after: 0, line: 240, lineRule: "auto" },
      children: [new TextRun({ text: String(text), bold, italics: italic, size, font: FONT,
                                color: shade ? "FFFFFF" : undefined })],
    })],
  });
}

function makeTable(widths, headerCells, bodyRows, headerShade = "1F4E78") {
  const rows = [];
  rows.push(new TableRow({
    tableHeader: true,
    children: headerCells.map((h, i) => cell(h, { width: widths[i], bold: true, size: 18,
                                                     shade: headerShade, align: AlignmentType.CENTER })),
  }));
  for (const r of bodyRows) {
    rows.push(new TableRow({
      children: r.map((v, i) => cell(v, { width: widths[i], size: 17 })),
    }));
  }
  return new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
                      columnWidths: widths, borders: THINBORDER, rows });
}

// ---------- PORTADA ----------
const portada = [
  new Paragraph({ spacing: { before: 2200, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "AVAL DE INFORMACIÓN", bold: true, size: 48, font: FONT })] }),
  new Paragraph({ spacing: { before: 200, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Material de respaldo y evidencia documental del proceso de revisión sistemática",
                              size: 26, font: FONT, italics: true })] }),
  new Paragraph({ spacing: { before: 600, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Chatbots basados en inteligencia artificial:", bold: true, size: 24, font: FONT })] }),
  new Paragraph({ spacing: { before: 0, after: 800 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Una revisión sistemática de la literatura, 2021-2025", bold: true, size: 24, font: FONT })] }),
  new Paragraph({ spacing: { before: 1200, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Autor: Pittman Medina, Edward Miguel", size: 22, font: FONT })] }),
  new Paragraph({ spacing: { before: 100, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Asesor(a): Vidal Rischmoller, Julio Cesar", size: 22, font: FONT })] }),
  new Paragraph({ spacing: { before: 100, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Universidad Autónoma del Perú — Desarrollo de Investigación", size: 22, font: FONT })] }),
  new Paragraph({ spacing: { before: 100, after: 1200 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Octubre de 2026", size: 22, font: FONT })] }),
  new Paragraph({ spacing: { before: 400, after: 0 }, alignment: AlignmentType.JUSTIFIED,
    children: [new TextRun({ text:
      "Este documento reúne, de forma separada al artículo de revisión, el material probatorio del proceso " +
      "metodológico: el diagrama de flujo PRISMA, las tablas de cribado y exclusión con sus conteos, la matriz " +
      "completa de extracción de los estudios incluidos y las fichas de extracción individuales. Su propósito " +
      "es dejar constancia verificable de cómo se llegó a cada cifra y a cada decisión reportadas en el " +
      "artículo, y servir de respaldo ante cualquier solicitud de sustento adicional.",
      size: 20, italics: true, font: FONT })] }),
  new Paragraph({ children: [new TextRun({ text: "", break: 1 })], pageBreakBefore: true }),
];

// ---------- TABLA 1: resumen PRISMA ----------
const resumenWidths = [4200, 1400, 3426];
const resumenRows = [
  ["Scopus", "840", ""],
  ["ScienceDirect", "186", "Incluye 1 título repetido dentro de la propia exportación"],
  ["SciELO", "104", "Sin filtro de año en la búsqueda"],
  ["IEEE Xplore", "78", ""],
  ["Google Scholar (inglés)", "150", "Primeros 150 de 2060 por relevancia (Haddaway et al., 2015)"],
  ["Google Scholar (español)", "222", "Cobertura completa"],
  ["TOTAL IDENTIFICADOS", "1580", ""],
  ["Eliminados antes del cribado: SciELO fuera de año o solo portugués", "42", "33 fuera de 2021-2025 y 9 solo en portugués"],
  ["Eliminados antes del cribado: duplicados", "367", "Coincidencia exacta de título normalizado"],
  ["REGISTROS PARA CRIBADO (título/resumen)", "1171", ""],
  ["Excluidos por título/resumen", "913", "913 excluidos en 1171 cribados; 258 pasan a evaluación a texto completo"],
  ["Evaluados a texto completo", "159", "De 258 elegibles; 99 no pudieron evaluarse por limitaciones de acceso (ver nota)"],
  ["Excluidos a texto completo", "83", "FT1=56, FT4=14, FT5=8, FT2=3, DUP=1, FT3=1"],
  ["Incluidos en la síntesis", "76", "Estudios incluidos (Figura 1 y Tabla 6 / fichas de extracción)"],
];

const seccion1 = [
  heading("1. Resumen numérico del proceso PRISMA"),
  p("La Tabla 1 reproduce, fase por fase, los conteos exactos del proceso de identificación, cribado, " +
    "elegibilidad e inclusión, tal como constan en el archivo de trabajo " +
    "PRISMA_identificacion_consolidado.xlsx (hoja Resumen_PRISMA).", { after: 160 }),
  makeTable(resumenWidths, ["Fase / fuente", "N", "Nota"], resumenRows),
  p("Nota sobre limitaciones de cobertura. De los 258 registros elegibles para evaluación a texto completo, " +
    "99 (38%) no pudieron revisarse dentro del periodo de trabajo por limitaciones de acceso a las plataformas " +
    "editoriales (límites de uso de las herramientas de recuperación disponibles). No se trata de una exclusión " +
    "por criterios de elegibilidad, sino de una limitación metodológica declarada del estudio. Asimismo, durante " +
    "la evaluación a texto completo se detectó un registro adicional (id 1393) que resultó ser un duplicado " +
    "cross-idioma del id 411 (mismo DOI, el mismo estudio publicado con título en español y en inglés en la " +
    "Revista Venezolana de Gerencia); este duplicado no fue identificado por el algoritmo de deduplicación por " +
    "título normalizado, por lo que se excluyó manualmente con código DUP al momento de su detección.",
    { before: 160, italic: true, size: 20 }),
];

// ---------- FIGURA 1 ----------
const imgBuffer = fs.readFileSync("/tmp/claude-0/-home-claude/cd0e5fa1-4ddf-5d67-96cf-b0025679cb0b/scratchpad/figura1_prisma.png");
const IMG_W_IN = 6.1, IMG_H_IN = IMG_W_IN * (2948 / 2398);
const seccion2 = [
  new Paragraph({ pageBreakBefore: true, children: [] }),
  heading("2. Figura 1. Diagrama de flujo PRISMA 2020"),
  p("Elaboración propia a partir de los conteos de la Tabla 1 y de las hojas Cribado y Texto_completo del " +
    "archivo de trabajo. Las cifras del diagrama coinciden exactamente con las de dicha tabla.", { after: 160 }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new ImageRun({ type: "png", data: imgBuffer,
      transformation: { width: Math.round(IMG_W_IN * 96), height: Math.round(IMG_H_IN * 96) } })],
  }),
];

// ---------- TABLA 2: ecuaciones de búsqueda ----------
const eqWidths = [1900, 5926, 1200];
const eqRows = [
  ["Scopus", 'TITLE(chatbot OR chatbots OR "conversational agent") AND TITLE("artificial intelligence") AND PUBYEAR > 2020 AND PUBYEAR < 2026', "840"],
  ["ScienceDirect", 'Título: (chatbot OR chatbots OR "conversational agent") AND ("artificial intelligence")', "186"],
  ["SciELO", 'Sin ecuación verificable: falla técnica persistente del buscador durante el periodo de trabajo; todos los índices, sin filtro de año', "104"],
  ["IEEE Xplore", '"Document Title": (chatbot OR chatbots OR "conversational agent") AND "artificial intelligence"', "78"],
  ["Google Scholar (inglés)", 'intitle:(chatbot OR chatbots OR "conversational agent") intitle:"artificial intelligence"', "150*"],
  ["Google Scholar (español)", 'intitle:(chatbot OR chatbots OR "agente conversacional") intitle:"inteligencia artificial"', "222"],
  ["Total", "", "1580"],
];
const seccion3 = [
  heading("3. Tabla 2. Ecuaciones de búsqueda y registros identificados por fuente", { before: 400 }),
  p("Reproduce, para efectos de trazabilidad, la misma información que la Tabla 1 del apartado Método del " +
    "artículo.", { after: 160 }),
  makeTable(eqWidths, ["Fuente", "Ecuación de búsqueda", "Registros"], eqRows),
  p("Nota. *Primeros 150 de 2060 resultados, por relevancia. Salvo SciELO (sin filtro de año), todas las " +
    "búsquedas se limitaron a 2021-2025. Elaboración propia.", { before: 100, italic: true, size: 18, after: 160 }),
];

// ---------- TABLA 3 y 4: motivos de exclusión ----------
const LEYENDA = {
  DUP: "Duplicado",
  TD: "Tipo de documento no elegible (ponencia, carta, capítulo, editorial, tesis, libro, preprint, nota de prensa)",
  OA: "Sin acceso abierto al texto completo",
  T1: "El chatbot basado en IA no es el objeto central",
  T2: "Sin método explícito (opinión, comentario, ensayo, revisión narrativa)",
  T3: "Prototipo o desarrollo sin evaluación empírica",
  T4: "Solo desempeño frente a bancos de preguntas o consultas simuladas, sin usuarios",
  FT1: "Texto completo no recuperado",
  FT2: "Tipo de documento no elegible (detectado en texto completo)",
  FT3: "Idioma distinto de español o inglés",
  FT4: "No cumple criterios al leer el texto completo",
  FT5: "Calidad metodológica insuficiente (puntaje < 4 de 5)",
};
const exclWidths = [1200, 6626, 1200];
function exclTable(codes, total) {
  const rows = codes.map(([c, n]) => [c, LEYENDA[c], String(n)]);
  rows.push(["Total", "", String(total)]);
  return makeTable(exclWidths, ["Código", "Significado", "n"], rows);
}
const seccion4 = [
  heading("4. Motivos de exclusión por fase", { before: 400 }),
  p("Tabla 3. Exclusiones en el cribado por título y resumen (n = 913, sobre 1171 registros cribados). " +
    "Conteos verificados directamente sobre la hoja Cribado del archivo de trabajo.", { after: 120, size: 20, bold: true }),
  exclTable([["TD", 447], ["OA", 175], ["T4", 124], ["T2", 65], ["DUP", 42], ["T1", 32], ["T3", 28]], 913),
  p("Tabla 4. Exclusiones en la evaluación a texto completo (n = 83, sobre 258 registros elegibles). " +
    "Conteos verificados directamente sobre la hoja Texto_completo del archivo de trabajo.",
    { before: 240, after: 120, size: 20, bold: true }),
  exclTable([["FT1", 56], ["FT4", 14], ["FT5", 8], ["FT2", 3], ["DUP", 1], ["FT3", 1]], 83),
];

// ---------- TABLA 5: síntesis agregada ----------
function aggTable(title, pairs) {
  const widths = [5500, 1452];
  return [
    p(title, { before: 200, after: 100, size: 20, bold: true, align: AlignmentType.LEFT }),
    makeTable(widths, ["Categoría", "n"], pairs.map(([k, v]) => [k, String(v)])),
  ];
}
const latamLabel = { si: "Sí (estudio con muestra o caso en América Latina)", no: "No" };
const puntajeLabel = { "5": "5 / 5 (cumple todos los criterios de calidad)", "4": "4 / 5" };
const seccion5 = [
  heading("5. Síntesis agregada de los 76 estudios incluidos", { before: 400 }),
  p("Conteos obtenidos directamente de la hoja Matriz_extraccion del archivo de trabajo " +
    "(76 estudios, sin el registro duplicado id 1393).", { after: 160 }),
  ...aggTable("Por sector", agg.sector),
  ...aggTable("Por tipo de estudio", agg.tipo_estudio),
  ...aggTable("Por madurez de implementación", agg.madurez),
  ...aggTable("Estudios con presencia de América Latina", agg.latam.map(([k, v]) => [latamLabel[k] || k, v])),
  ...aggTable("Puntaje de calidad metodológica (escala 0-5, aplicada en el cribado a texto completo)",
              agg.puntaje.map(([k, v]) => [puntajeLabel[k] || k, v])),
];

// ---------- TABLA 6: matriz de extracción compacta (landscape) ----------
const matrixWidths = [500, 700, 1600, 1900, 1800, 4200, 1800, 700, 758];
const matrixHeader = ["N", "ID", "País", "Sector", "Tipo de estudio", "Chatbot evaluado", "Madurez", "Calidad", "LATAM"];
const matrixRows = studies.map(s => [
  s.n, s.id, s.pais, s.sector_label, s.tipo_label, s.chatbot, s.madurez_label,
  `${s.puntaje}/5`, s.latam === "si" ? "Sí" : "No",
]);
const tabla6 = makeTable(matrixWidths, matrixHeader, matrixRows);

// ---------- FICHAS DE EXTRACCIÓN DETALLADA ----------
function ficha(s) {
  return [
    new Paragraph({
      spacing: { before: 220, after: 40 }, keepNext: true,
      children: [new TextRun({ text: `[${s.n}] ID ${s.id} — `, bold: true, size: 20, font: FONT }),
                 new TextRun({ text: s.ref_ieee, size: 20, font: FONT })],
    }),
    new Paragraph({
      spacing: { before: 0, after: 60 },
      children: [new TextRun({
        text: `País: ${s.pais}  ·  Sector: ${s.sector_label}  ·  Tipo de estudio: ${s.tipo_label}  ·  ` +
              `Madurez: ${s.madurez_label}  ·  Calidad: ${s.puntaje}/5  ·  LATAM: ${s.latam === "si" ? "Sí" : "No"}`,
        italics: true, size: 18, font: FONT,
      })],
    }),
    p(`Muestra / unidad de análisis: ${s.muestra || "No reportado por los autores"}`, { size: 19, after: 40, align: AlignmentType.LEFT }),
    p(`Chatbot evaluado: ${s.chatbot || "No reportado por los autores"}`, { size: 19, after: 40, align: AlignmentType.LEFT }),
    p(`Hallazgo principal: ${s.hallazgo || "No reportado por los autores"}`, { size: 19, after: 40 }),
    p(`Barreras / riesgos: ${s.barreras || "No reportado por los autores"}`, { size: 19, after: 40 }),
    p(`Limitaciones declaradas por los autores: ${s.limitaciones || "No reportado por los autores"}`, { size: 19, after: 120 }),
  ];
}

const bySector = {};
for (const s of studies) {
  (bySector[s.sector_label] = bySector[s.sector_label] || []).push(s);
}
const sectorOrder = agg.sector.map(([k]) => k);

let fichasContent = [];
for (const sec of sectorOrder) {
  fichasContent.push(heading(`${sec} (n = ${bySector[sec].length})`, { size: 24, before: 300, after: 120 }));
  for (const s of bySector[sec]) fichasContent = fichasContent.concat(ficha(s));
}

const notaFinal = [
  heading("8. Nota metodológica final", { before: 400 }),
  p("Este documento reitera, para constancia, las tres limitaciones metodológicas declaradas en el artículo " +
    "principal: (1) SciELO no pudo consultarse mediante una ecuación de búsqueda verificable debido a una falla " +
    "técnica persistente del buscador durante el periodo de trabajo, por lo que sus 104 registros se obtuvieron " +
    "sin filtro de año y se depuraron manualmente; (2) de los 258 registros elegibles para evaluación a texto " +
    "completo, 99 (38%) no pudieron revisarse por limitaciones de acceso a las plataformas editoriales dentro " +
    "del periodo de trabajo, y se documentan como no evaluados, no como excluidos; (3) se detectó y corrigió un " +
    "duplicado cross-idioma (id 1393, mismo estudio que el id 411) no capturado por el algoritmo de " +
    "deduplicación por título normalizado, lo que redujo el número final de estudios incluidos de 77 a 76.",
    { after: 0 }),
];

// ---------- SECCIONES DEL DOCUMENTO ----------
const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  sections: [
    {
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
      children: [...portada, ...seccion1, ...seccion2, ...seccion3, ...seccion4, ...seccion5],
    },
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE },
          margin: { top: 1080, bottom: 1080, left: 1080, right: 1080 },
        },
      },
      children: [
        heading("6. Tabla 6. Matriz de extracción completa — 76 estudios incluidos", { before: 0 }),
        p("Ordenada por sector y luego por ID. El ID permite ubicar la referencia IEEE completa y el detalle de " +
          "hallazgos, barreras y limitaciones en la sección 7 (fichas de extracción detallada).", { after: 160 }),
        tabla6,
      ],
    },
    {
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
      children: [
        heading("7. Fichas de extracción detallada", { before: 0 }),
        p("Una ficha por cada uno de los 76 estudios incluidos, agrupadas por sector en el mismo orden que la " +
          "Tabla 6. Cada ficha reproduce la referencia IEEE completa y los campos extraídos (hallazgo principal, " +
          "barreras o riesgos, y limitaciones declaradas por los propios autores del estudio).", { after: 200 }),
        ...fichasContent,
        ...notaFinal,
      ],
    },
  ],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("/tmp/aval/Aval_de_informacion_revision_sistematica.docx", buf);
  console.log("done, bytes:", buf.length);
});
