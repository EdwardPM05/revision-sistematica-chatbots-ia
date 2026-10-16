# Revisión Sistemática de Chatbots basados en IA

Carpeta completa del proyecto de revisión sistemática para el curso **Desarrollo de Investigación** (Universidad Autónoma del Perú).

## Estructura

```
Revision_Sistematica_Chatbots/
├── Documentos/           # Documentos finales del proyecto
│   ├── Articulo_revision_sistematica_chatbots_IA.docx
│   ├── Aval_de_informacion_revision_sistematica.docx
│   ├── Metodologia_proceso_revision_sistematica.docx
│   └── PRISMA_identificacion_consolidado.xlsx
├── Scripts/              # Scripts de construcción (Node.js)
│   ├── build.js         # Constructor del "Aval de información" (32 pág.)
│   └── build_methodology.js  # Constructor del documento metodológico
├── Datos/               # Datos fuente y diagramas
│   ├── incluidos_full.json   # 76 estudios incluidos (fuente de verdad)
│   ├── data.json             # Versión enriquecida con etiquetas y agregados
│   └── figura1_prisma.png    # Diagrama PRISMA (2398×2948 px)
└── README.md            # Este archivo
```

## Documentos

### 1. Articulo_revision_sistematica_chatbots_IA.docx
Artículo principal con:
- **Introducción** (4 páginas): contexto, justificación, pregunta de investigación, estado del arte
- **Método** (1 página): protocolo, ecuaciones de búsqueda, criterios, diagrama PRISMA, evaluación de calidad
- Referencias [1]-[26] en estilo IEEE (Resolución N° 025-2024-VRI-UA)

**Pendientes para completar:**
- Análisis e integración de la información (5-6 pág.)
- Conclusión (máx. 1 pág.)
- Resumen/Abstract en español e inglés

### 2. Aval_de_informacion_revision_sistematica.docx
Documento de soporte (32 páginas) con:
- **Tabla 1**: Conteos PRISMA fase por fase (1578 identificados → 76 incluidos)
- **Figura 1**: Diagrama de flujo PRISMA con imagen
- **Tabla 2**: Ecuaciones de búsqueda por base de datos
- **Tablas 3-4**: Códigos de exclusión y conteos (título/resumen y texto completo)
- **Tablas 5a-5e**: Síntesis agregada (sector, tipo de estudio, madurez, región, calidad)
- **Tabla 6**: Matriz landscape 76×9 (ID, país, sector, tipo, chatbot, madurez, calidad, LATAM)
- **Fichas 1-76**: Extracción detallada de cada estudio incluido
- **Nota metodológica final**: Limitaciones técnicas declaradas explícitamente

### 3. Metodologia_proceso_revision_sistematica.docx
Documento metodológico (14 KB, 149 párrafos) que documenta:
- Marco PRISMA 2020
- Herramientas de software utilizadas (Node.js, Python, docx-js, openpyxl, LibreOffice)
- Procedimiento paso a paso (búsquedas, cribado, evaluación)
- Disciplina de verificación de datos (recomputación contra fuentes primarias)
- Errores encontrados y correcciones aplicadas
- Archivos entregables con referencias

### 4. PRISMA_identificacion_consolidado.xlsx
Libro de cálculo consolidado (625 KB) con:
- **Resumen_PRISMA**: Conteos fase por fase y limitaciones técnicas
- **Leyenda_codigos**: Definiciones exactas de códigos de exclusión
- **Cribado**: 913 registros excluidos con sus códigos
- **Texto_completo**: 83 registros excluidos en fase de texto completo, 99 no evaluados
- **Matriz_extraccion**: 76 estudios incluidos con 18 campos cada uno
- Fórmulas y referencias cruzadas para auditoría

## Scripts (Node.js)

### build.js
Constructor del "Aval de información" usando docx-js. Características:
- Crea documento .docx de 32 páginas desde cero (no requiere Microsoft Office)
- Dos secciones portrait (A4) + una landscape para Tabla 6
- Funciones helper: `p()` (párrafo), `heading()` (títulos), `cell()`/`makeTable()` (tablas), `ficha()` (fichas de estudio)
- Carga datos de `incluidos_full.json` y construye tablas agregadas
- Inserta imagen PRISMA escalada proporcionalmente
- Valida con XSD antes de entregar

**Uso:**
```bash
cd Scripts
node build.js
# Genera: Aval_de_informacion_revision_sistematica.docx
```

### build_methodology.js
Constructor del documento metodológico. Similar a `build.js` pero con estructura de 10 secciones (Contexto, Propósito, Competencias, Marco PRISMA, Procedimiento, Herramientas, Construcción, Verificación, Errores, Entregables).

## Datos Fuente

### incluidos_full.json
Estructura de verdad con 76 estudios incluidos. Cada estudio contiene:
```json
{
  "id": "número único",
  "recuperado": "sí/no",
  "fuente_texto": "base de datos original",
  "decision": "incluido/excluido",
  "codigo": "código si aplica",
  "q1-q5": "evaluación metodológica por criterio",
  "puntaje": "1-5 escala de calidad",
  "pais": "país de institución/publicación",
  "sector": "educación/salud/etc.",
  "tipo_estudio": "tipo de investigación",
  "muestra": "número de participantes/registros",
  "chatbot": "descripción del chatbot evaluado",
  "madurez": "prototipo/piloto/comercial/etc.",
  "hallazgo": "hallazgo principal del estudio",
  "barreras": "barreras/riesgos identificados",
  "limitaciones": "limitaciones reportadas",
  "latam": "sí/no (presencia en Latinoamérica)",
  "ref_ieee": "referencia completa en estilo IEEE"
}
```

### data.json
Versión enriquecida de `incluidos_full.json` con:
- `SECTOR_LABEL`: mapeo de sector a etiqueta legible
- `TIPO_LABEL`: mapeo de tipo_estudio a etiqueta
- `MADUREZ_LABEL`: mapeo de madurez a etiqueta
- Ordenamiento por (sector, id)
- Asignación de número secuencial `n` (1-76)
- Objeto `agg` con conteos agregados (sector, tipo_estudio, madurez, latam, puntaje)

### figura1_prisma.png
Diagrama de flujo PRISMA 2020 (2398×2948 px, 220 dpi).
- Muestra fases: Identificación → Selección → Elegibilidad → Inclusión
- Conteos verificados contra Leyenda_codigos
- Incrustado en Figura 1 del Aval de información

## Verificación de Datos

Todos los números reportados en documentos fueron recomputados directamente desde fuentes primarias:

1. **Conteos PRISMA**: extraídos de `PRISMA_identificacion_consolidado.xlsx` sheet "Resumen_PRISMA"
2. **Códigos de exclusión**: recomputados desde sheets "Cribado" y "Texto_completo" usando openpyxl
3. **Agregados**: cargados de `incluidos_full.json`, computados con `collections.Counter` en Python
4. **Validación XSD**: cada .docx validado con `validate.py` (Office Open XML standard)
5. **Renderizado visual**: documentos convertidos a PDF y JPEG, inspeccionados página por página

## Limitaciones Técnicas Documentadas

1. **WebFetch rate-limit agotado**: 99 de 258 registros en fase de texto completo no pudieron ser evaluados (limitación de la herramienta de acceso a contenidos)

2. **Inconsistencia en etiquetas PRISMA**: el diagrama `figura1_prisma.png` contiene descripciones que no coinciden exactamente con la Leyenda_codigos. Los conteos son correctos; solo las glosas son aproximadas. Documentado en Nota Metodológica Final.

3. **SciELO no indexada completamente**: esta base de datos tiene cobertura limitada; búsqueda manual fue necesaria.

## Para Usar con Claude Code

1. Abre esta carpeta como proyecto en Claude Code
2. Los scripts `build.js` y `build_methodology.js` pueden re-ejecutarse para regenerar documentos
3. Datos en JSON son accesibles para análisis/auditoría
4. Los documentos finales (.docx, .xlsx) están listos para enviar/imprimir

Para regenerar documentos:
```bash
cd Scripts
npm install docx  # Si no está instalado
node build.js
node build_methodology.js
```

## Próximos Pasos

Para completar la revisión sistemática:

1. **Análisis e integración de la información** (5-6 pág.):
   - Síntesis narrativa de hallazgos por sector
   - Análisis de tendencias en madurez de chatbots
   - Comparación de barreras y limitaciones
   - Identificación de brechas de investigación

2. **Conclusión** (máx. 1 pág.):
   - Síntesis de respuesta a la pregunta de investigación
   - Implicaciones prácticas
   - Recomendaciones para futuras investigaciones

3. **Resumen/Abstract** (230-250 palabras cada uno):
   - Versión en español: estructura (objetivo, método, resultados, conclusión)
   - Versión en inglés: misma estructura traducida
   - Palabras clave: mínimo 5

4. **Revisión final**: verificación de formato, referencias, y cumplimiento de Resolución N° 121-2024-VRI-UA

---

**Creado:** Octubre 2026  
**Asesor:** Julio Cesar Vidal Rischmoller  
**Curso:** Desarrollo de Investigación (Universidad Autónoma del Perú)
