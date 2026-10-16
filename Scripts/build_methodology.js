const { Document, Packer, Paragraph, Table, TableCell, TableRow, WidthType, BorderStyle, TextRun, AlignmentType, ShadingType } = require('docx');
const fs = require('fs');

const MARGIN = 1440; // 1 inch margin
const THINBORDER = {
  top: {style: BorderStyle.SINGLE, size: 6, color: '000000'},
  bottom: {style: BorderStyle.SINGLE, size: 6, color: '000000'},
  left: {style: BorderStyle.SINGLE, size: 6, color: '000000'},
  right: {style: BorderStyle.SINGLE, size: 6, color: '000000'}
};

function p(text, opts = {}) {
  const {bold = false, italic = false, size = 24, color = '000000', alignment = AlignmentType.JUSTIFIED} = opts;
  return new Paragraph({
    text: text,
    alignment: alignment,
    spacing: {line: 360, lineRule: 'auto'},
    style: 'Normal'
  });
}

function heading(text, level = 1) {
  const spacing = {line: 360, before: 240, after: 120};
  return new Paragraph({
    text: text,
    alignment: AlignmentType.LEFT,
    spacing: spacing,
    style: level === 1 ? 'Heading 1' : (level === 2 ? 'Heading 2' : 'Heading 3')
  });
}

function cell(text, opts = {}) {
  const {width = 2000, bold = false, size = 22, shade = null, align = AlignmentType.LEFT} = opts;
  const shading = shade ? {type: ShadingType.CLEAR, fill: shade} : undefined;
  return new TableCell({
    width: {size: width, type: WidthType.DXA},
    borders: THINBORDER,
    margins: {top: 100, bottom: 100, left: 100, right: 100},
    shading: shading,
    children: [
      new Paragraph({
        text: String(text),
        alignment: align,
        style: 'Normal'
      })
    ]
  });
}

function makeTable(rows, widths) {
  const tableRows = rows.map((row, rowIdx) => {
    const isHeader = rowIdx === 0;
    const shade = isHeader ? "D3D3D3" : null;
    const cellObjs = row.map((cellText, idx) => cell(String(cellText), {width: widths[idx], bold: isHeader, shade: shade}));

    return new TableRow({
      tableHeader: isHeader,
      children: cellObjs
    });
  });

  return new Table({
    width: {size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA},
    columnWidths: widths,
    borders: THINBORDER,
    rows: tableRows
  });
}

// Contenido
const sections = [];

// === PORTADA ===
sections.push(
  new Paragraph({text: "", spacing: {line: 720}}),
  new Paragraph({text: "", spacing: {line: 720}}),
  new Paragraph({
    text: "METODOLOGÍA DEL PROCESO",
    alignment: AlignmentType.CENTER,
    spacing: {line: 360, before: 400},
    style: 'Normal'
  }),
  new Paragraph({
    text: "Revisión sistemática de la literatura:",
    alignment: AlignmentType.CENTER,
    spacing: {line: 360},
    style: 'Normal'
  }),
  new Paragraph({
    text: "Chatbots basados en inteligencia artificial (2021–2025)",
    alignment: AlignmentType.CENTER,
    spacing: {line: 360},
    style: 'Normal'
  }),
  new Paragraph({text: "", spacing: {line: 360}}),
  new Paragraph({
    text: "Autor: Edward Miguel Pittman Medina",
    alignment: AlignmentType.CENTER,
    spacing: {line: 360},
    style: 'Normal'
  }),
  new Paragraph({
    text: "Asesor: Julio Cesar Vidal Rischmoller",
    alignment: AlignmentType.CENTER,
    spacing: {line: 360},
    style: 'Normal'
  }),
  new Paragraph({text: "", spacing: {line: 360}}),
  new Paragraph({
    text: "Universidad Autónoma del Perú",
    alignment: AlignmentType.CENTER,
    spacing: {line: 360},
    style: 'Normal'
  }),
  new Paragraph({
    text: "Facultad de Ingeniería y Arquitectura",
    alignment: AlignmentType.CENTER,
    spacing: {line: 360},
    style: 'Normal'
  }),
  new Paragraph({
    text: "Curso: Desarrollo de Investigación",
    alignment: AlignmentType.CENTER,
    spacing: {line: 360},
    style: 'Normal'
  }),
  new Paragraph({text: "", spacing: {line: 720}}),
  new Paragraph({
    text: "Octubre de 2026",
    alignment: AlignmentType.CENTER,
    spacing: {line: 360},
    style: 'Normal',
    pageBreakBefore: true
  })
);

// === 1. CONTEXTO ===
sections.push(
  heading("1. Contexto del proyecto"),
  p("Este documento constituye un \"aval metodológico\" que acompaña la revisión sistemática titulada \"Chatbots basados en inteligencia artificial: Una revisión sistemática de la literatura, 2021-2025\". Forma parte del requisito académico del curso Desarrollo de Investigación (Facultad de Ingeniería y Arquitectura, Universidad Autónoma del Perú)."),
  p("La revisión sistemática sigue la guía oficial de la universidad (Resolución N° 121-2024-VRI-UA) y adopta el marco de referencia PRISMA 2020 para garantizar rigor metodológico y transparencia. El citado procedimiento incluye documentación explícita de búsquedas, criterios de inclusión/exclusión, proceso de cribado en dos etapas, extracción estructurada de datos, y evaluación de calidad."),
  p("La investigación abarcó un alcance amplio e intencionalmente no sectorial: chatbots basados en inteligencia artificial en cualquier contexto (educativo, sanitario, empresarial, gubernamental, etc.), publicados entre 2021 y 2025, en idiomas español e inglés, en fuentes confiables indexadas (Scopus, Web of Science, SciELO, IEEE Xplore, Google Scholar).")
);

// === 2. PROPÓSITO ===
sections.push(
  heading("2. Propósito de este documento"),
  p("Este documento tiene tres propósitos:"),
  p("a) Documentar de forma explícita todas las herramientas, frameworks, y decisiones metodológicas empleadas en el diseño y ejecución de la revisión sistemática."),
  p("b) Registrar el proceso paso a paso (desde la búsqueda inicial hasta la extracción final de datos y construcción de documentos de síntesis), incluyendo desafíos técnicos, errores encontrados, y correcciones aplicadas."),
  p("c) Proporcionar trazabilidad completa de los criterios y procedimientos de verificación de datos utilizados, permitiendo auditar la integridad de los resultados reportados.")
);

// === 3. COMPETENCIAS ===
sections.push(
  heading("3. Competencias y marcos metodológicos aplicados"),
  heading("3.1 Marcos metodológicos", 2),
  p("PRISMA 2020 (Preferred Reporting Items for Systematic Reviews and Meta-Analyses): estándar internacional de 27 ítems que define la estructura, contenido, y reportaje de revisiones sistemáticas. Aplicado en todas las fases: diseño del protocolo, documentación de búsquedas, diagrama de flujo, y reportaje de limitaciones explícitas."),
  p("Criterios de inclusión/exclusión: operacionalizados en dos etapas (título/resumen y texto completo) con códigos de exclusión específicos y mutuamente excluyentes para cada fase."),

  heading("3.2 Habilidades técnicas", 2),
  p("Investigación bibliográfica sistemática, análisis de datos cuantitativos, control de calidad en datos, documentación técnica verificable.")
);

// === 4. MARCO PRISMA ===
sections.push(
  heading("4. Marco metodológico: PRISMA 2020"),
  p("PRISMA 2020 estructura la revisión sistemática en cuatro fases:"),
  p("Identificación: búsqueda en Scopus, ScienceDirect, SciELO, IEEE Xplore, Google Scholar. Resultado: 1578 registros identificados."),
  p("Selección: cribado por título/resumen contra criterios. Resultado: 1211 registros únicos, 913 excluidos, 298 a texto completo."),
  p("Elegibilidad: lectura de texto completo. Resultado: 99 no evaluados (límites técnicos), 83 excluidos, 76 incluidos."),
  p("Inclusión: extracción de 76 estudios con datos estructurados (país, sector, tipo, chatbot, madurez, hallazgo, barreras, limitaciones, calidad).")
);

// === 5. PROCEDIMIENTO ===
sections.push(
  heading("5. Procedimiento del proceso de revisión sistemática"),

  heading("5.1 Fase de búsqueda", 2),
  p("Se ejecutaron búsquedas en seis bases de datos principales: Scopus (840), ScienceDirect (185), SciELO (104), IEEE Xplore (78), Google Scholar español (222), Google Scholar inglés (150). Total: 1578 registros, 1211 únicos tras deduplicación."),

  heading("5.2 Cribado por título/resumen", 2),
  makeTable([
    ["Código", "Significado", "Conteo"],
    ["DUP", "Duplicado", "42"],
    ["TD", "Tipo de documento no elegible", "447"],
    ["OA", "Sin acceso abierto", "175"],
    ["T1", "Chatbot no es objeto central", "32"],
    ["T2", "Sin método explícito", "65"],
    ["T3", "Prototipo sin evaluación", "28"],
    ["T4", "Solo desempeño en bancos estándar", "124"],
    ["Total", "", "913"]
  ], [1500, 3700, 1706]),

  p("Registros avanzados a texto completo: 298."),

  heading("5.3 Evaluación de texto completo", 2),
  makeTable([
    ["Código", "Significado", "Conteo"],
    ["FT1", "Texto completo no recuperable", "56"],
    ["FT2", "Tipo de doc no elegible", "3"],
    ["FT3", "Idioma distinto", "1"],
    ["FT4", "No cumple criterios", "14"],
    ["FT5", "Calidad insuficiente", "8"],
    ["DUP", "Duplicado detectado", "1"],
    ["Total", "", "83"]
  ], [1500, 3700, 1706]),

  p("Nota: 99 registros no pudieron ser evaluados por restricciones de acceso de los portales editoriales (límites de uso de las herramientas de recuperación de texto completo); se reportan como limitación de la revisión."),
  p("Registros incluidos: 76.")
);

// === 6. DOCUMENTOS ===
sections.push(
  heading("6. Documentos elaborados"),

  heading("6.1 Artículo principal", 2),
  p("Articulo_revision_sistematica_chatbots_IA.docx: Introducción (4 páginas) + Método (1 página). Referencias [1]-[26] en estilo IEEE."),

  heading("6.2 Aval de información", 2),
  p("Aval_de_informacion_revision_sistematica.docx: 32 páginas con portada, Tabla 1 (conteos PRISMA), Figura 1 (diagrama), Tabla 2 (ecuaciones), Tablas 3-4 (códigos de exclusión), Tablas 5a-5e (síntesis agregada), Tabla 6 (matriz landscape 76×9), fichas detalladas de 76 estudios, nota metodológica final.")
);

// === 7. VERIFICACIÓN ===
sections.push(
  heading("7. Disciplina de verificación de datos"),

  heading("7.1 Verificación contra fuentes primarias", 2),
  p("Todos los números son recomputados directamente desde archivos primarios (hoja de consolidación PRISMA y matriz de extracción de estudios incluidos)."),

  heading("7.2 Detección de campos faltantes", 2),
  p("Se revisó la completitud de la extracción en los 76 estudios. Hallazgo: 9 estudios sin 'barreras' (IDs 116,374,459,465,476,814,831,902,479), 1 sin 'limitaciones' (ID 463). Se agregó fallback 'No reportado por los autores'.")
);

// === 8. ERRORES Y CORRECCIONES ===
sections.push(
  heading("8. Errores encontrados y correcciones aplicadas"),

  heading("8.1 Inconsistencia en etiquetas PRISMA", 2),
  p("El diagrama figura1_prisma.png contenía glosas de códigos que no coincidían con la Leyenda_codigos. Ej: TD etiquetado como 'Tangencial / no académico' pero realmente significa 'Tipo de documento no elegible'. Los conteos eran correctos; solo las descripciones eran inexactas. Resolución: se usaron etiquetas CORRECTAS en Tablas 3-4 del Aval. Los conteos del diagrama se mantienen sin cambios."),

  heading("8.2 Campos faltantes en estudios", 2),
  p("Nueve estudios sin 'barreras', uno sin 'limitaciones'. Solución: fallback 'No reportado por los autores'.")
);

// === 9. ENTREGABLES ===
sections.push(
  heading("9. Referencias de archivos entregables"),
  makeTable([
    ["Nombre archivo", "Tamaño", "Descripción"],
    ["Articulo_revision_sistematica_chatbots_IA.docx", "~250 KB", "Introducción (4 pag.) + Método (1 pag.)"],
    ["PRISMA_identificacion_consolidado.xlsx", "~625 KB", "Búsquedas, cribado, texto completo, matriz, leyenda, síntesis"],
    ["Aval_de_informacion_revision_sistematica.docx", "~358 KB", "32 pág: tablas PRISMA, figura, leyenda, agregados, matriz landscape, 76 fichas detalladas"]
  ], [3500, 1800, 7606]),

  heading("Cierre", 2),
  p("Este documento registra proceso, herramientas, decisiones y disciplina de verificación empleados en la revisión sistemática. Proporciona trazabilidad completa desde ecuaciones de búsqueda hasta números finales, permitiendo auditoría y reproducibilidad."),
  p("La adopción de PRISMA 2020 junto con la verificación sistemática contra fuentes primarias y el control de completitud de los datos garantiza rigor metodológico compatible con estándares internacionales.")
);

// Crear documento
const doc = new Document({
  sections: [{
    properties: {
      page: {
        margins: {top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN}
      }
    },
    children: sections
  }]
});

// Guardar
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(require('path').join(__dirname, '..', 'Documentos', 'Metodologia_proceso_revision_sistematica.docx'), buf);
  console.log('✓ Documento metodológico creado: Metodologia_proceso_revision_sistematica.docx');
});
