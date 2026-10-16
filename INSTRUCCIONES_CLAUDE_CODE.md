# Instrucciones para Usar con Claude Code

Esta carpeta contiene toda la metodología, datos, scripts, y documentos de la revisión sistemática. Puedes usarla con **Claude Code** para continuar desarrollando o auditar el proyecto.

## Configuración Inicial

1. Abre la carpeta en Claude Code:
   ```bash
   cd "C:\Users\epitt\Downloads\Claude outputs\Revision_Sistematica_Chatbots"
   ```

2. Verifica que tienes Node.js instalado:
   ```bash
   node --version
   npm --version
   ```

3. Instala `docx` si necesitas regenerar documentos:
   ```bash
   npm install docx
   ```

## Entendiendo el Proyecto

- **Curso:** Desarrollo de Investigación (Universidad Autónoma del Perú)
- **Tema:** Chatbots basados en inteligencia artificial (2021-2025)
- **Marco:** PRISMA 2020 (Preferred Reporting Items for Systematic Reviews)
- **Estudios incluidos:** 76 (de 1578 identificados inicialmente)

Lee `README.md` para una descripción completa de cada documento y archivo.

## Estructura de Decisiones

### Fase 1: Identificación (1578 registros)
Búsquedas en 6 bases de datos: Scopus, ScienceDirect, SciELO, IEEE Xplore, Google Scholar español, Google Scholar inglés.

### Fase 2: Selección por Título/Resumen (1211 únicos)
**913 excluidos** con códigos mutuamente excluyentes:
- DUP (42): Duplicado
- TD (447): Tipo de documento no elegible (editorial, comentario)
- OA (175): Sin acceso abierto
- T1 (32): Chatbot no es objeto central
- T2 (65): Sin método explícito
- T3 (28): Prototipo sin evaluación empírica
- T4 (124): Solo desempeño en bancos de preguntas estándar

**298 avanzados** a texto completo.

### Fase 3: Evaluación de Texto Completo (199 evaluados)
- 99 no evaluados (agotamiento de WebFetch rate-limit — documentado como limitación)
- **83 excluidos** con códigos:
  - FT1 (56): Texto completo no recuperable
  - FT2 (3): Tipo de documento no elegible
  - FT3 (1): Idioma distinto
  - FT4 (14): No cumple criterios
  - FT5 (8): Calidad metodológica insuficiente
  - DUP (1): Duplicado detectado

### Fase 4: Inclusión (76 estudios)
Extracción estructurada de datos: país, sector, tipo de estudio, chatbot, madurez, hallazgo, barreras, limitaciones, puntaje de calidad (1-5).

## Trabajar con Datos

### Cargar datos JSON en Python
```python
import json

# Cargar estudios incluidos
with open('Datos/incluidos_full.json', 'r', encoding='utf-8') as f:
    estudios = json.load(f)

# Acceder a un estudio
primer_estudio = estudios[0]
print(f"ID: {primer_estudio['id']}, País: {primer_estudio['pais']}, Sector: {primer_estudio['sector']}")

# Contar estudios por sector
from collections import Counter
sectores = Counter(e['sector'] for e in estudios)
print(sectores.most_common())
```

### Cargar datos Excel en Python
```python
from openpyxl import load_workbook

wb = load_workbook('Documentos/PRISMA_identificacion_consolidado.xlsx')

# Leer conteos PRISMA
resumen = wb['Resumen_PRISMA']
for row in resumen.iter_rows(min_row=2, max_row=5, values_only=True):
    print(row)

# Leer leyenda de códigos
leyenda = wb['Leyenda_codigos']
for row in leyenda.iter_rows(min_row=2, max_row=10, values_only=True):
    print(row)
```

## Regenerar Documentos

### Regenerar "Aval de Información" (32 páginas)
```bash
cd Scripts
node build.js
# Genera: Aval_de_informacion_revision_sistematica.docx
# Valida con XSD automáticamente
```

### Regenerar Documento Metodológico
```bash
cd Scripts
node build_methodology.js
# Genera: Metodologia_proceso_revision_sistematica.docx
```

### Validar documentos .docx
```bash
# Necesitas tener librerías de validación instaladas (vea docx skill)
python3 scripts/office/validate.py ../Documentos/Aval_de_informacion_revision_sistematica.docx
```

## Auditar Datos

### Verificar consistencia de conteos

```python
import json
from collections import Counter
from openpyxl import load_workbook

# Cargar estudios
with open('Datos/incluidos_full.json', 'r') as f:
    estudios = json.load(f)

# Verificar totales
print(f"Total estudios incluidos: {len(estudios)}")

# Contar por categoría
print("\nPor sector:")
print(Counter(e['sector'] for e in estudios).most_common())

print("\nPor tipo de estudio:")
print(Counter(e['tipo_estudio'] for e in estudios).most_common())

print("\nPor madurez:")
print(Counter(e['madurez'] for e in estudios).most_common())

print("\nPor puntaje de calidad:")
print(Counter(e['puntaje'] for e in estudios).most_common())

print("\nPresencia en Latinoamérica:")
print(Counter(e['latam'] for e in estudios).most_common())

# Buscar campos faltantes
campos = ['muestra', 'chatbot', 'hallazgo', 'barreras', 'limitaciones']
for campo in campos:
    vacios = [e['id'] for e in estudios if not e.get(campo) or e.get(campo).strip() == '']
    if vacios:
        print(f"\n⚠ {len(vacios)} estudios sin {campo}: {vacios}")
```

### Verificar contra Excel
```python
from openpyxl import load_workbook

wb = load_workbook('Documentos/PRISMA_identificacion_consolidado.xlsx')

# Leer tabla de extracción
matriz = wb['Matriz_extraccion']
conteo_excel = matriz.max_row - 1  # Resta header
print(f"Estudios en Excel: {conteo_excel}")
print(f"Estudios en JSON: {len(estudios)}")
print(f"¿Coinciden? {conteo_excel == len(estudios)}")
```

## Completar el Artículo

Los tres documentos pendientes (Análisis e integración, Conclusión, Resumen/Abstract) deben insertarse en `Articulo_revision_sistematica_chatbots_IA.docx` continuando desde la página después de la sección Método.

### Estructura esperada
```
[Portada + Introducción + Método]
↓
[3. Análisis e integración de la información] (5-6 pág.) ← PENDIENTE
[4. Conclusión] (máx. 1 pág.) ← PENDIENTE
[Referencias]
[Resumen/Abstract español] (230-250 palabras) ← PENDIENTE
[Abstract inglés] (230-250 palabras) ← PENDIENTE
```

## Contactar con el Asesor

**Asesor:** Julio Cesar Vidal Rischmoller

Cuando envíes documentos parciales o finales al asesor, incluye siempre:
- El artículo principal (.docx)
- El "Aval de información" (.docx) como documento de soporte
- El libro de cálculo consolidado (.xlsx) como anexo técnico
- Esta carpeta completa si necesita revisar metodología/datos

## Resoluciones Universales Aplicadas

- **Resolución N° 121-2024-VRI-UA:** Guía oficial de tesis/revisiones sistemáticas de la universidad
- **Resolución N° 025-2024-VRI-UA:** Requiere IEEE citation style (ya aplicado en todos los documentos)

## Notas Técnicas

### Orientación Landscape en Tabla 6
La Tabla 6 (matriz 76×9) está en orientación landscape para caber los 9 columnas (ID, País, Sector, Tipo, Chatbot, Madurez, Calidad, LATAM). En `build.js`, se usa:
```javascript
size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }
```
**Importante:** Pasar dimensiones PORTRAIT + orientation LANDSCAPE (docx-js invierte internamente).

### Diagrama PRISMA
El archivo `figura1_prisma.png` tiene descripción de códigos que no coincide exactamente con `Leyenda_codigos.xlsx`. Los conteos son correctos; solo las glosas son aproximadas. Esto está documentado en la Nota Metodológica Final del Aval.

### WebFetch Rate-Limit
99 de 258 registros no pudieron ser evaluados porque WebFetch (herramienta de acceso a PDF/HTML) agotó sus límites después de ~150 documentos. Esto está documentado explícitamente en:
- Tabla 1 del Aval (nota sobre registros no evaluados)
- Nota Metodológica Final
- Sección Limitaciones del documento metodológico

## Troubleshooting

**Problema:** El script `build.js` falla al cargar `incluidos_full.json`  
**Solución:** Verifica la ruta relativa. Ejecuta desde `Scripts/`:
```bash
cd Scripts
node build.js
```

**Problema:** Documentos .docx no abren en Word  
**Solución:** Valida con `validate.py` para detectar errores XSD. Si la validación pasa, el documento es válido pero Word podría tener un bug; intenta abrirlo con LibreOffice.

**Problema:** Tabla 6 (landscape) sale cortada  
**Solución:** No está cortada; necesita escala de zoom menor en el visor. Intenta zoom 80-90% en Word/PDF reader.

---

**Última actualización:** Octubre 2026  
**Para preguntas o mejoras:** Abre esta carpeta en Claude Code y describe el cambio.
