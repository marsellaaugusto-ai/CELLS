# Relevamiento de apps similares y hoja de ruta de Cells

Octubre 2026. Este documento resume qué ofrecen las hojas de cálculo más usadas, qué se puede
incorporar a Cells sin problemas legales, y en qué orden conviene hacerlo antes de publicar.

> No es asesoramiento legal. Antes de lanzar comercialmente, conviene una consulta con un
> abogado de propiedad intelectual, sobre todo por marcas y patentes.

---

## 1. Apps relevadas

| App | Tipo | Lo que la distingue |
|---|---|---|
| **Microsoft Excel** (escritorio y web) | Estándar del mercado | Motor de fórmulas más completo (matrices dinámicas, LET, LAMBDA), tablas dinámicas, Power Query, gráficos, macros VBA, Python en celdas |
| **Google Sheets** | Colaborativa en la nube | Edición simultánea, comentarios con menciones, historial de versiones, Apps Script, IMPORTRANGE, casillas y chips desplegables, formularios |
| **LibreOffice Calc** | Libre, escritorio | Formato abierto ODS, alta compatibilidad con Excel, sin costo |
| **Apple Numbers** | Escritorio / iPad | Lienzo libre: varias tablas y gráficos sueltos en una misma hoja, diseño cuidado |
| **Zoho Sheet**, **OnlyOffice**, **WPS** | Suites en la nube | Clones cercanos de Excel con colaboración; OnlyOffice es de código abierto (AGPL) |
| **Airtable**, **Smartsheet** | Base de datos con aspecto de planilla | Tipos de columna (adjuntos, selección, vínculos), vistas kanban/calendario/Gantt, formularios, automatizaciones |
| **Grist** | Planilla relacional, código abierto | Columnas con fórmulas en Python, tablas vinculadas, permisos por fila |
| **Rows**, **Equals** | Planillas conectadas | Integraciones nativas con APIs y bases de datos, asistentes de IA |
| **Quadratic** | Lienzo infinito con código | Celdas con Python, SQL y JavaScript |
| **Coda**, **Notion** (tablas) | Documento + tabla | Tablas dentro de documentos, botones y automatizaciones |
| **Univer**, **EtherCalc** | Componentes de código abierto | Base para construir planillas propias (Univer es Apache 2.0) |

---

## 2. Qué se puede tomar de otras apps (y qué no)

### Se puede usar libremente

- **Ideas, funciones y comportamientos.** El derecho de autor protege la *expresión* (el código,
  los textos, los gráficos), no las ideas ni los métodos de funcionamiento. Este principio está en
  el Acuerdo ADPIC (art. 9.2), el Tratado de la OMPI sobre Derecho de Autor (art. 2) y en la ley
  argentina de propiedad intelectual (Ley 11.723, que protege el software desde la Ley 25.036).
  Casos de referencia:
  - *Lotus v. Borland* (EE. UU., 1996): la estructura de menús de una planilla es un "método de
    operación" y no se protege.
  - *SAS Institute v. World Programming* (Tribunal de Justicia de la UE, 2012): la funcionalidad de
    un programa, su lenguaje y sus formatos de archivo no se protegen por derecho de autor.
  - *Google v. Oracle* (Corte Suprema de EE. UU., 2021): reimplementar una interfaz para que los
    usuarios puedan usar lo que ya saben fue uso legítimo.
- **Nombres de funciones** como `SUMA`, `BUSCARV`, `XLOOKUP`. Son funcionales y necesarios para la
  compatibilidad.
- **El formato .xlsx.** Es un estándar abierto (ECMA-376 / ISO/IEC 29500), y Microsoft lo incluye en
  su *Open Specification Promise*, que permite implementarlo. ODS también es estándar (ISO/IEC 26300).
- **Atajos de teclado comunes** (Ctrl+C, Ctrl+Z, F2, Ctrl+flechas). Son convenciones de uso.

### No se debe copiar

- **Código** de otras apps, salvo que su licencia lo permita. Cells usa solo bibliotecas MIT o
  Apache 2.0 (ver `THIRD_PARTY.md`). Hay que evitar las GPL/AGPL (por ejemplo HyperFormula y
  OnlyOffice), que obligarían a publicar el código de Cells.
- **Íconos, logos, ilustraciones, sonidos, plantillas y textos de ayuda** de Microsoft, Google u
  otros. Hay que hacer los propios.
- **La apariencia exacta** de otra app (combinación distintiva de colores, disposición y estilo).
  Puede reclamarse como "imagen comercial" (*trade dress*) o competencia desleal. El diseño propio
  de Cells (riel con triángulos, tarjetas) ya es distintivo.
- **Marcas** como "Excel" y "Google Sheets", y los nombres comerciales de funciones (por ejemplo
  "Flash Fill", "Smart Fill" o "Power Query"). Se pueden *mencionar* para decir que Cells es
  compatible ("abre archivos de Excel"), pero no usarlas como nombre de funciones propias, ni dar
  a entender una relación con esas empresas.

### Puntos a revisar con un profesional

- **Patentes.** Son independientes del derecho de autor: algunas técnicas concretas (por ejemplo,
  ciertas formas de autocompletar o de recalcular) pueden estar patentadas, sobre todo en EE. UU.
  El riesgo es bajo para funciones clásicas de planillas, que existen desde hace décadas.
- **Registro de la marca "Cells"** en Argentina (INPI) y en los países donde se ofrezca. "Cells" es
  una palabra común y puede haber registros previos en la clase 9/42; conviene una búsqueda de
  antecedentes antes de invertir en la marca.
- **Datos personales**: al guardar datos de usuarios en la nube hacen falta términos de uso y una
  política de privacidad (Ley 25.326 en Argentina; RGPD si hay usuarios en Europa).

---

## 3. Funciones recomendadas

Prioridad 1 es lo que un usuario de Excel o Google Sheets espera desde el primer día. El esfuerzo
es una estimación relativa (S/M/L).

### Prioridad 1: antes de publicar

| Función | Inspiración | Esfuerzo |
|---|---|---|
| **Las fórmulas se ajustan al copiar, pegar y arrastrar** (`A1` relativo, `$A$1` absoluto). Hoy se copian tal cual | Todas | M |
| **Controlador de relleno**: arrastrar la esquina de la celda para copiar fórmulas o continuar series (1, 2, 3… / enero, febrero…) | Todas | M |
| **Insertar y eliminar filas y columnas**, ajustando las fórmulas que las referencian | Todas | M |
| **Hojas grandes rápidas**: hoy se dibujan todas las celdas; hace falta dibujar solo lo visible para archivos de 10.000+ filas | Todas | M |
| **Autocompletar funciones** con ayuda de argumentos, y elegir celdas con el mouse mientras se escribe la fórmula | Excel, Sheets | M |
| **Buscar y reemplazar** | Todas | S |
| **Filtros por columna** y orden por varias columnas | Todas | M |
| **Formato condicional visible** (ya se conserva al importar; falta mostrarlo y poder crearlo) | Todas | M |
| **Escribir fechas y horas** (18/10/2026, 14:30) y que se reconozcan como fecha | Todas | S |
| **Matrices dinámicas**: `UNIQUE`, `FILTER`, `SORT`, `SEQUENCE`, `LET`, `INDIRECT`, `OFFSET` | Excel 365, Sheets | L |
| **Atajos de navegación**: Ctrl+flechas, Ctrl+Shift+flechas, Inicio/Fin, Ctrl+Inicio | Todas | S |
| **Permisos por persona**: lector, comentador, editor; y enlace para compartir | Sheets | M |
| **Historial de versiones** con restaurar | Sheets | M |
| **Imprimir y exportar a PDF** | Todas | M |
| **Términos de uso y política de privacidad** | Requisito legal | S |

### Prioridad 2: primeros meses

| Función | Inspiración | Esfuerzo |
|---|---|---|
| Más gráficos (líneas, torta, dispersión, combinados) e importar los de Excel | Todas | L |
| Comentarios con hilos y menciones (@persona) | Sheets | M |
| Tablas con nombre y referencias estructuradas (`Tabla1[Total]`) | Excel | L |
| Tablas dinámicas | Excel, Sheets | L |
| Validación completa (números, fechas, rangos) y casillas de verificación | Sheets | M |
| Proteger hojas y rangos | Excel, Sheets | M |
| Imágenes en celdas y flotantes | Excel, Sheets | M |
| Modo claro (además del oscuro) | — | S |
| Uso sin conexión (instalable como app, con sincronización al volver) | Sheets | M |
| Abrir archivos `.ods` (LibreOffice) y `.numbers` | Calc, Numbers | M |
| Abrir y guardar directo en Google Drive, sin descargar | Integraciones | M |

### Prioridad 3: diferenciales

Estas funciones distinguirían a Cells de un clon:

| Función | Inspiración | Esfuerzo |
|---|---|---|
| **Automatizaciones**: "cuando cambie esta columna, avisar a tal app" (webhooks), en línea con la idea de sincronizar con otras apps | Airtable, Smartsheet | M |
| **API de escritura** para que otras apps carguen datos, no solo los lean | Rows, Airtable | M |
| **Formularios** que agregan filas a una hoja | Google Forms, Airtable | M |
| **Asistente con IA**: escribir fórmulas en lenguaje natural, explicar una fórmula, limpiar datos | Excel Copilot, Sheets Gemini | M |
| **Vistas** kanban, calendario y galería sobre la misma tabla | Airtable, Smartsheet | L |
| **Columnas con tipo** (selección, persona, adjunto, vínculo a otra tabla) | Airtable, Grist | L |
| **Celdas con código** (JavaScript o Python) | Quadratic, Excel Python | L |
| **Plantillas propias** de Cells (presupuesto de obra, inventario, cómputo métrico) | — | S |

---

## 4. Compatibilidad con Excel y Google Sheets

**Objetivo:** que todo lo que Cells entiende viaje sin cambios en ambos sentidos, y que lo que
todavía no entiende se avise claramente al abrir el archivo.

Una conversión 100 % idéntica no la logra ningún programa; ni siquiera Google Sheets con los
archivos de Excel. Cuando se abre un archivo, Cells muestra un informe con tres listas: qué
funciona igual, qué se conserva con límites y qué no se trae todavía.

Google Sheets no tiene un formato de archivo propio para descargar: exporta e importa `.xlsx`.
Por eso el `.xlsx` es el puente en los dos sentidos.

| Elemento | Estado en Cells |
|---|---|
| Valores, textos, números, fechas, booleanos | ✅ Se conservan |
| Textos que parecen números (`00123`) | ✅ Se conservan como texto |
| Fórmulas, entre hojas y con rangos con nombre | ✅ Se calculan |
| Fórmulas con funciones no soportadas (`UNIQUE`, `FILTER`…) | ⚠️ Muestran el último resultado de Excel (con borde punteado) y se exportan intactas |
| Formato de números (moneda, %, fechas, colores `[Red]`) | ✅ Igual, con separadores en español |
| Fuente, tamaño, negrita, cursiva, subrayado, tachado, color | ✅ |
| Relleno, bordes, alineación, ajuste de texto | ✅ |
| Ancho de columnas, alto de filas, filas y columnas ocultas | ✅ |
| Celdas combinadas, filas y columnas fijas | ✅ |
| Notas, enlaces | ✅ |
| Listas desplegables | ✅ Se usan al escribir |
| Otras validaciones | ⚠️ Se conservan, todavía no se aplican |
| Formato condicional | ⚠️ Se conserva, todavía no se muestra |
| Gráficos, imágenes | ❌ Se cuentan en el informe; no se traen |
| Tablas dinámicas, tablas con formato | ❌ Quedan los valores |
| Macros (VBA), Apps Script | ❌ No se ejecutan (por seguridad, tampoco conviene) |
| Texto con varios formatos dentro de una celda | ❌ Queda el texto, sin los formatos parciales |

### Pruebas que respaldan esta tabla

- Un archivo de prueba con todos los elementos anteriores (`test/fixtures/completo.xlsx`).
- Ese mismo archivo guardado por LibreOffice (`test/fixtures/libreoffice.xlsx`), que escribe los
  archivos de otra forma que Excel.
- Viaje de ida y vuelta: abrir → guardar desde Cells → volver a abrir da el mismo contenido.
- El `.xlsx` exportado por Cells se verificó abriéndolo con LibreOffice, que recalculó bien todas
  las fórmulas, y con openpyxl (lector de Python independiente).

Falta probar con archivos reales guardados por Excel y por Google Sheets, con un uso típico de
la empresa.

---

## 5. Notas técnicas pendientes

- `npm audit` marca una vulnerabilidad moderada en `uuid`, dependencia de ExcelJS. El caso
  afectado (generar UUID v3/v5/v6 en un buffer propio) no se usa en Cells, pero conviene
  forzar una versión corregida cuando ExcelJS lo permita.
- ExcelJS (929 KB) se carga solo cuando alguien abre o guarda un Excel, para no hacer lenta la
  primera carga de la app.
