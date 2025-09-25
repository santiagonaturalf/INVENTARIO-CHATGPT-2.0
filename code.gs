/**
 * @OnlyCurrentDoc
 *
 * El script principal para el sistema de Inventario 2.0 en Google Sheets.
 * Este script maneja la creación de hojas, importación de datos, cálculo de inventario diario,
 * y la automatización de tareas.
 */

// =====================================================================================
// CONSTANTES Y CONFIGURACIÓN GLOBAL
// =====================================================================================

// URLs de las hojas de cálculo de origen
const URL_ORIGEN_ORDERS_SKU = "https://docs.google.com/spreadsheets/d/1-gVCyrB57thPhC-4TlsA10ifWlFd78GSGUFYVCYqeXk/edit";
const URL_ORIGEN_ADQUISICIONES = "https://docs.google.com/spreadsheets/d/1vCZejbBPMh73nbAhdZNYFOlvJvRoMA7PVSCUiLl8MMQ/edit";

// Nombres de las hojas
const HOJA_ORDERS = "Orders";
const HOJA_SKU = "SKU";
const HOJA_ADQUISICIONES = "Adquisiciones";
const HOJA_HISTORICO = "Inventario Histórico";
const HOJA_REPORTE_HOY = "Reporte Hoy";

// Zona horaria para cálculos de fecha
const TIMEZONE = "America/Santiago";

// =====================================================================================
// FUNCIONES DE AUTOMATIZACIÓN Y MENÚ
// =====================================================================================

/**
 * Normaliza una cadena de texto: la convierte a minúsculas y elimina espacios en blanco al inicio y al final.
 * @param {string} s La cadena a normalizar.
 * @returns {string} La cadena normalizada.
 */
const norm = s => (s ?? '').toString().trim().toLowerCase();

/**
 * Se ejecuta cuando se abre la hoja de cálculo.
 * Crea un menú personalizado para ejecutar las funciones principales.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Inventario 2.0')
      .addItem('INICIAR DIA', 'calcularInventarioDiario')
      .addItem('CERRAR DIA', 'cerrarDia')
      .addSeparator()
      .addItem('Abrir Dashboard de Inventario', 'showDashboard')
      .addSeparator()
      .addItem('Contactar Cliente (dashboard)', 'openContactarCliente')
      .addToUi();
}

/**
 * Se ejecuta automáticamente cuando un usuario edita una celda en la hoja de cálculo.
 * Si la edición ocurre en la columna "Stock Real" de "Reporte Hoy", actualiza
 * el estado del producto en la hoja "Estados".
 * @param {object} e El objeto de evento de Google Apps Script.
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();

  // Verificar la hoja, la columna (F = 6) y la fila (mayor que 1 para evitar el encabezado)
  if (sheetName === HOJA_REPORTE_HOY && range.getColumn() === 6 && range.getRow() > 1) {
    const productBase = sheet.getRange(range.getRow(), 1).getValue();
    const newValue = e.value;

    if (!productBase) {
      return; // No hay producto base en esta fila, no hacer nada.
    }

    // Determinar el nuevo estado. Si la celda está vacía o no es un número, es 'pendiente'.
    const newState = (newValue !== null && newValue !== "" && !isNaN(parseFloat(newValue))) ? 'aprobado' : 'pendiente';

    // Actualizar la hoja de Estados usando la función helper existente.
    try {
      setEstadoProducto(productBase, newState, 'Actualizado por onEdit');
    } catch (error) {
      Logger.log(`Error al actualizar estado para ${productBase} vía onEdit: ${error.message}`);
    }
  }
}

/**
 * Archiva el estado actual de "Reporte Hoy" en "Inventario Histórico".
 * Esta función está pensada para ser ejecutada manualmente al final del día.
 */
function cerrarDia() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    const reporteSheet = ss.getSheetByName(HOJA_REPORTE_HOY);
    const histSheet = ss.getSheetByName(HOJA_HISTORICO);
    const skuSheet = ss.getSheetByName(HOJA_SKU);

    if (!reporteSheet || !histSheet || !skuSheet) {
      throw new Error("No se encontraron una o más hojas requeridas: Reporte Hoy, Inventario Histórico, SKU.");
    }

    if (reporteSheet.getLastRow() < 2) {
      ui.alert("La hoja 'Reporte Hoy' está vacía. No hay nada que archivar.");
      return;
    }

    // 1. Read data from sheets
    const reporteData = reporteSheet.getRange(2, 1, reporteSheet.getLastRow() - 1, 6).getValues(); // A:F
    const skuData = skuSheet.getRange(2, 1, skuSheet.getLastRow() - 1, 8).getValues(); // A:H

    // 2. Create a map for SKU units for quick lookup
    const skuUnitMap = new Map(skuData.map(row => [norm(row[1]), row[7]])); // Map: normalized product base -> unit

    // 3. Prepare new historical records
    const newHistoricalRecords = [];
    const timestamp = new Date();

    reporteData.forEach(row => {
      const productBase = row[0];
      const stockReal = row[5]; // Column F

      // Only process rows where a "Stock Real" has been entered
      if (productBase && (stockReal !== null && stockReal !== "" && !isNaN(parseFloat(stockReal)))) {
        const unit = skuUnitMap.get(norm(productBase)) || '';
        newHistoricalRecords.push([
          timestamp,
          productBase,
          parseFloat(stockReal),
          unit
        ]);
      }
    });

    if (newHistoricalRecords.length === 0) {
      ui.alert("No hay productos con 'Stock Real' para archivar. No se realizó ninguna acción.");
      return;
    }

    // 4. Append new records in a single batch
    histSheet.getRange(histSheet.getLastRow() + 1, 1, newHistoricalRecords.length, 4).setValues(newHistoricalRecords);

    // 5. Clean up old historical data (keep last 5 per product)
    const allHistData = histSheet.getDataRange().getValues();
    allHistData.shift(); // Remove header for processing

    const histMap = new Map();
    // Re-map all historical data including the newly added ones
    allHistData.forEach((row, index) => {
      const base = row[1];
      if (base) {
        const baseNorm = norm(base);
        if (!histMap.has(baseNorm)) histMap.set(baseNorm, []);
        // Store original row index (add 2 because we shifted header and it's 1-based)
        histMap.get(baseNorm).push({ rowIndex: index + 2, timestamp: new Date(row[0]) });
      }
    });

    const rowsToDelete = new Set();
    histMap.forEach(entries => {
      if (entries.length > 5) {
        // Sort by date ascending to find the oldest
        entries.sort((a, b) => a.timestamp - b.timestamp);
        const toDeleteCount = entries.length - 5;
        for (let i = 0; i < toDeleteCount; i++) {
          rowsToDelete.add(entries[i].rowIndex);
        }
      }
    });

    // Delete rows in reverse order to avoid index shifts
    if (rowsToDelete.size > 0) {
      Array.from(rowsToDelete).sort((a, b) => b - a).forEach(rowIndex => {
        histSheet.deleteRow(rowIndex);
      });
    }

    ui.alert(`¡Día cerrado con éxito! Se han archivado ${newHistoricalRecords.length} registros en el Inventario Histórico.`);

  } catch (e) {
    Logger.log(`Error en cerrarDia: ${e.message}`);
    ui.alert(`Ocurrió un error al cerrar el día: ${e.message}`);
  }
}

/**
 * Configura el entorno inicial del sistema de inventario.
 * Crea las hojas necesarias, establece las fórmulas de importación y los encabezados.
 * Es una función idempotente: se puede ejecutar varias veces sin causar problemas.
 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. Hoja de Orders ---
  const hojaOrders = obtenerOCrearHoja(ss, HOJA_ORDERS);
  const formulaOrders = `=IMPORTRANGE("${URL_ORIGEN_ORDERS_SKU}"; "Orders!A:K")`;
  // Solo escribir la fórmula si la celda A1 está vacía para no sobrescribir
  if (hojaOrders.getRange("A1").getFormula() === "") {
    hojaOrders.getRange("A1").setFormula(formulaOrders);
  }

  // --- 2. Hoja de SKU ---
  const hojaSKU = obtenerOCrearHoja(ss, HOJA_SKU);
  const formulaSKU = `=IMPORTRANGE("${URL_ORIGEN_ORDERS_SKU}"; "SKU!A:K")`;
  if (hojaSKU.getRange("A1").getFormula() === "") {
    hojaSKU.getRange("A1").setFormula(formulaSKU);
  }

  // --- 3. Hoja de Adquisiciones ---
  const hojaAdquisiciones = obtenerOCrearHoja(ss, HOJA_ADQUISICIONES);
  const formulaAdquisiciones = `=IMPORTRANGE("${URL_ORIGEN_ADQUISICIONES}"; "RESUMEN_Adquisiciones!A:M")`;
  if (hojaAdquisiciones.getRange("A1").getFormula() === "") {
    hojaAdquisiciones.getRange("A1").setFormula(formulaAdquisiciones);
  }

  // --- 4. Hoja de Inventario Histórico ---
  const hojaHistorico = obtenerOCrearHoja(ss, HOJA_HISTORICO);
  const encabezadosHistorico = ["Timestamp", "Producto Base", "Stock Real", "Unidad Venta"];
  // Solo escribir encabezados si la fila 1 está vacía
  if (hojaHistorico.getRange("A1").getValue() === "") {
      hojaHistorico.getRange(1, 1, 1, encabezadosHistorico.length).setValues([encabezadosHistorico]).setFontWeight("bold");
  }

  // --- 5. Hoja de Reporte Hoy ---
  const hojaReporte = obtenerOCrearHoja(ss, HOJA_REPORTE_HOY);
  const encabezadosReporte = ["Producto Base", "Inventario Ayer", "Compras del Día", "Ventas del Día", "Inventario Hoy (estimado)", "Stock Real", "Discrepancias"];
  if (hojaReporte.getRange("A1").getValue() === "") {
    hojaReporte.getRange(1, 1, 1, encabezadosReporte.length).setValues([encabezadosReporte]).setFontWeight("bold");
  }

  // --- 6. Hoja de Discrepancias ---
  const hojaDiscrepancias = obtenerOCrearHoja(ss, "Discrepancias");
  const encabezadosDiscrepancias = ["Timestamp", "Producto Base", "Inventario Estimado", "Inventario Real", "Discrepancia"];
   if (hojaDiscrepancias.getRange("A1").getValue() === "") {
    hojaDiscrepancias.getRange(1, 1, 1, encabezadosDiscrepancias.length).setValues([encabezadosDiscrepancias]).setFontWeight("bold");
  }

  // Hoja Estados (persistencia de verificación/aprobación)
  const hojaEstados = obtenerOCrearHoja(ss, 'Estados');
  const headers = ['Producto Base','Estado','Notas','Usuario','Timestamp'];
  if (hojaEstados.getRange(1,1,1,headers.length).isBlank()) {
    hojaEstados.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  }

  // --- Hoja de Clientes Notificados (para el dashboard de contacto) ---
  const hojaNotificados = obtenerOCrearHoja(ss, "ClientesNotificados");
  const encabezadosNotificados = ["Link de Notificación", "Timestamp"];
  if (hojaNotificados.getRange("A1").getValue() === "") {
    hojaNotificados.getRange(1, 1, 1, encabezadosNotificados.length).setValues([encabezadosNotificados]).setFontWeight("bold");
  }

  SpreadsheetApp.getUi().alert("¡Configuración completada! Las hojas han sido creadas y configuradas.");
}

/**
 * Crea o actualiza el disparador (trigger) programado para ejecutar el cálculo de inventario diariamente.
 * Primero elimina cualquier disparador existente para esta función para evitar duplicados.
 */
function crearDisparadorDiario() {
  const nombreFuncion = 'calcularInventarioDiario';

  // 1. Eliminar triggers existentes para esta función
  const todosLosTriggers = ScriptApp.getProjectTriggers();
  for (const trigger of todosLosTriggers) {
    if (trigger.getHandlerFunction() === nombreFuncion) {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // 2. Crear el nuevo trigger
  ScriptApp.newTrigger(nombreFuncion)
      .timeBased()
      .everyDays(1)
      .atHour(2) // Se ejecuta todos los días entre las 2 y 3 AM
      .inTimezone(TIMEZONE)
      .create();

  // 3. Notificar al usuario
  SpreadsheetApp.getUi().alert(`¡Disparador configurado! La función '${nombreFuncion}' se ejecutará automáticamente todos los días entre las 2:00 y 3:00 AM (hora de Chile).`);
}

// =====================================================================================
// LÓGICA PRINCIPAL - CÁLCULO DE INVENTARIO
// =====================================================================================

/**
 * Función principal que calcula el inventario del día.
 * Orquesta la lectura de datos, el procesamiento y la escritura de resultados.
 * Se ejecuta diariamente a través de un disparador.
 */
function calcularInventarioDiario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(HtmlService.createHtmlOutput('<h3>Procesando inventario...</h3><p>Este proceso puede tardar unos minutos. Por favor, no cierres la hoja.</p>'), 'Cálculo de Inventario en Curso');

  try {
    // --- INICIO: Limpiar estados aprobados al iniciar el día ---
    const hojaEstados = ss.getSheetByName('Estados');
    if (hojaEstados && hojaEstados.getLastRow() > 1) {
      const data = hojaEstados.getDataRange().getValues();
      const rowsToDelete = [];
      // Itera desde la segunda fila (índice 1) para saltar el encabezado.
      for (let i = 1; i < data.length; i++) {
        // El estado está en la columna B (índice 1).
        if (data[i][1] && data[i][1].toString().toLowerCase() === 'aprobado') {
          // Se guarda el número de fila real (i + 1).
          rowsToDelete.push(i + 1);
        }
      }
      // Elimina las filas en orden inverso para evitar problemas con los índices.
      for (let i = rowsToDelete.length - 1; i >= 0; i--) {
        hojaEstados.deleteRow(rowsToDelete[i]);
      }
    }
    // --- FIN: Limpieza ---

    // --- 1. OBTENER DATOS ---
    const hojaSku = ss.getSheetByName(HOJA_SKU);
    const hojaAdquisiciones = ss.getSheetByName(HOJA_ADQUISICIONES);
    const hojaHistorico = ss.getSheetByName(HOJA_HISTORICO);
    const hojaReporteHoy = ss.getSheetByName(HOJA_REPORTE_HOY);

    const datosSku = hojaSku.getRange("A2:K" + hojaSku.getLastRow()).getValues();
    const datosAdquisiciones = hojaAdquisiciones.getRange("A2:M" + hojaAdquisiciones.getLastRow()).getValues();
    const datosHistorico = hojaHistorico.getLastRow() > 1 ? hojaHistorico.getRange("A2:E" + hojaHistorico.getLastRow()).getValues() : [];

    // --- 2. PREPARAR MAPAS DE BÚSQUEDA (Lookups) con normalización mejorada ---
    const mapaCompraSku = new Map();
    const baseOriginalNames = new Map(); // Mapa de nombre normalizado a nombre original

    datosSku.forEach(fila => {
      const [productoBase, formatoAdq, cantAdq] = [fila[1], fila[2], fila[3]];
      if (productoBase) {
        const productoBaseNorm = normalizeText(productoBase);
        if (!baseOriginalNames.has(productoBaseNorm)) {
          baseOriginalNames.set(productoBaseNorm, productoBase);
        }
        if (formatoAdq) {
          const claveCompra = `${productoBaseNorm}-${formatoAdq.toString().trim()}`;
          mapaCompraSku.set(claveCompra, parseFloat(String(cantAdq).replace(',', '.')) || 0);
        }
      }
    });

    // --- 3. PROCESAR VENTAS DEL DÍA (NUEVA LÓGICA) ---
    const { ventasPorProductoBase, filasIgnoradas } = calculateVentasDelDia();
    Logger.log(`Se ignoraron ${filasIgnoradas} filas de 'Orders' marcadas con 'E'.`);

    // --- 4. PROCESAR COMPRAS DEL DÍA ---
    const comprasDelDia = new Map();
    datosAdquisiciones.forEach(fila => {
      const productoBase = fila[1];
      const formatoCompra = fila[2];
      const cantidadComprada = parseFloat(String(fila[3]).replace(',', '.')) || 0;

      if (!productoBase || !formatoCompra) return;

      const productoBaseNorm = normalizeText(productoBase);
      const formatoAdq = _getFormatoAdquisicionBase(formatoCompra);
      const claveCompra = `${productoBaseNorm}-${formatoAdq}`;
      const cantAdquisicion = mapaCompraSku.get(claveCompra);

      if (cantAdquisicion) {
        const compraEnUnidadBase = cantidadComprada * cantAdquisicion;
        comprasDelDia.set(productoBaseNorm, (comprasDelDia.get(productoBaseNorm) || 0) + compraEnUnidadBase);
      }
    });

    // --- 5. OBTENER INVENTARIO DE AYER ---
    const inventarioAyer = new Map();
    const productosVistos = new Set();
    for (let i = datosHistorico.length - 1; i >= 0; i--) {
      const fila = datosHistorico[i];
      const productoBase = fila[1];
      if (productoBase) {
        const key = normalizeText(productoBase);
        if (!productosVistos.has(key)) {
          inventarioAyer.set(key, parseFloat(fila[2]) || 0);
          productosVistos.add(key);
        }
      }
    }

    // --- 6. CONSTRUIR Y ESCRIBIR EL REPORTE COMPLETO ---
    const reporteFinal = [];
    // Iterar sobre todos los productos base definidos en la hoja SKU
    for (const [productoNorm, productoOriginal] of baseOriginalNames.entries()) {
      const invAyer = inventarioAyer.get(productoNorm) || 0;
      const compras = comprasDelDia.get(productoNorm) || 0;
      const ventas = ventasPorProductoBase.get(productoNorm) || 0;
      const invHoyEstimado = invAyer + compras - ventas;

      reporteFinal.push([
        productoOriginal,    // Producto Base
        invAyer,             // Inventario Ayer
        compras,             // Compras del Día
        ventas,              // Ventas del Día
        invHoyEstimado,      // Inventario Hoy (estimado)
        '',                  // Stock Real (vacío para llenado manual)
        ''                   // Discrepancias (vacío para llenado manual)
      ]);
    }

    // Ordenar el reporte alfabéticamente por el nombre del producto
    reporteFinal.sort((a, b) => a[0].localeCompare(b[0]));

    // Limpiar la hoja (excepto encabezados) y escribir los nuevos datos
    if (hojaReporteHoy.getLastRow() > 1) {
      hojaReporteHoy.getRange(2, 1, hojaReporteHoy.getLastRow() - 1, 7).clearContent();
    }

    if (reporteFinal.length > 0) {
      hojaReporteHoy.getRange(2, 1, reporteFinal.length, 7).setValues(reporteFinal);
    }

    ui.showModalDialog(HtmlService.createHtmlOutput('<h3>¡Éxito!</h3><p>El "Reporte Hoy" ha sido calculado y actualizado correctamente.</p>'), 'Proceso Completado');
    Utilities.sleep(4000);
    const activeDoc = SpreadsheetApp.getActive();
    const html = HtmlService.createHtmlOutput("<script>google.script.host.close()</script>");
    ui.showModalDialog(html, "Cerrando...");

  } catch (e) {
    Logger.log(e);
    ui.showModalDialog(HtmlService.createHtmlOutput(`<h3>Error</h3><p>Ocurrió un error durante el cálculo: ${e.message}</p><pre>${e.stack}</pre>`), 'Error en el Proceso');
  }
}

/**
 * Completa las columnas 'Producto Base', 'Cantidad (venta)' y 'Unidad Venta' en la hoja 'Orders'
 * basándose en el mapeo de la hoja 'SKU'. Solo rellena las celdas que están vacías.
 * Es útil para enriquecer datos de forma manual o bajo demanda.
 * Esta versión es robusta: crea los encabezados si no existen.
 */
function completarSKUenOrders() {
  const ss = SpreadsheetApp.getActive();
  const shOrders = ss.getSheetByName('Orders');
  const shSKU = ss.getSheetByName('SKU');
  if (!shOrders || !shSKU) throw new Error('Faltan hojas Orders o SKU');

  // --- Asegurar encabezados en columnas L, M, N ---
  shOrders.getRange('L1').setValue('Producto Base');
  shOrders.getRange('M1').setValue('Cantidad (venta)');
  shOrders.getRange('N1').setValue('Unidad Venta');
  SpreadsheetApp.flush(); // Asegurar que los encabezados se escriban antes de continuar

  // --- Lectura de datos (después de asegurar encabezados) ---
  const orders = shOrders.getDataRange().getValues();
  const sku = shSKU.getDataRange().getValues();
  if (orders.length < 2 || sku.length < 2) return;

  // --- Definición de Índices ---
  const hdrO = orders[0].map(String);
  const idxNombre = hdrO.indexOf('Nombre Producto');
  const idxCantOrd = hdrO.indexOf('Cantidad');
  const idxBase = hdrO.indexOf('Producto Base');
  const idxCantVen = hdrO.indexOf('Cantidad (venta)');
  const idxUniVen = hdrO.indexOf('Unidad Venta');

  if (idxNombre === -1 || idxCantOrd === -1) {
    throw new Error('Faltan las columnas de origen críticas: "Nombre Producto" y/o "Cantidad" en la hoja "Orders".');
  }

  // Mapa SKU
  const hdrS = sku[0].map(String);
  const sNombre = hdrS.indexOf('Nombre Producto');
  const sBase = hdrS.indexOf('Producto Base');
  const sCantV = hdrS.indexOf('Cantidad Venta');
  const sUniV = hdrS.indexOf('Unidad Venta');
  if ([sNombre, sBase, sCantV, sUniV].some(i => i === -1)) {
    throw new Error('En la hoja "SKU" faltan columnas críticas: Nombre Producto, Producto Base, Cantidad Venta, Unidad Venta');
  }

  const map = {};
  for (let i = 1; i < sku.length; i++) {
    const n = (sku[i][sNombre] ?? '').toString().trim().toLowerCase();
    if (!n) continue;
    map[n] = {
      base: sku[i][sBase],
      fac: Number(sku[i][sCantV]) || 0,
      uni: sku[i][sUniV]
    };
  }

  // --- Procesamiento y enriquecimiento ---
  const out = orders.slice(); // copia editable
  const faltantes = new Set();

  for (let r = 1; r < out.length; r++) {
    const nombre = (out[r][idxNombre] ?? '').toString().trim().toLowerCase();
    if (!nombre) continue;
    const qtyOrd = Number(out[r][idxCantOrd]) || 0;
    const s = map[nombre];

    // Solo escribir si las celdas de destino están vacías
    if (out[r][idxBase] === "" || out[r][idxBase] === null) {
      if (s) {
        out[r][idxBase] = s.base;
      } else {
        faltantes.add(out[r][idxNombre]);
        continue; // Si no hay SKU, no podemos rellenar el resto
      }
    }
    if (out[r][idxCantVen] === "" || out[r][idxCantVen] === null) {
      if (s) out[r][idxCantVen] = qtyOrd * (s.fac || 0);
    }
    if (out[r][idxUniVen] === "" || out[r][idxUniVen] === null) {
      if (s) out[r][idxUniVen] = s.uni;
    }
  }

  // --- Volcado de datos ---
  shOrders.getRange(1, 1, out.length, out[0].length).setValues(out);

  if (faltantes.size > 0) {
    const shMiss = ss.getSheetByName('SKU_FALTANTES') || ss.insertSheet('SKU_FALTANTES');
    shMiss.clear();
    shMiss.getRange(1,1,1,1).setValue('Nombre Producto no mapeado en SKU (desde Enriquecer)');
    shMiss.getRange(2,1,faltantes.size,1).setValues([...faltantes].map(v => [v]));
    SpreadsheetApp.getUi().alert(`Proceso completado. Se encontraron ${faltantes.size} productos sin mapeo en SKU. Revisa la hoja 'SKU_FALTANTES'.`);
  } else {
    SpreadsheetApp.getUi().alert('Proceso de enriquecimiento completado exitosamente.');
  }
}


// =====================================================================================
// FUNCIONES AUXILIARES (Helpers)
// =====================================================================================

/**
 * Suma un valor a una propiedad de un objeto. Si la propiedad no existe, la inicializa.
 * @param {object} obj El objeto al que se le sumará el valor.
 * @param {string} key La clave o propiedad del objeto.
 * @param {number} value El valor numérico a sumar.
 */
function sumarAObjeto(obj, key, value) {
  if (obj[key]) {
    obj[key] += value;
  } else {
    obj[key] = value;
  }
}

/**
 * Obtiene o crea una hoja de cálculo por su nombre.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss La hoja de cálculo activa.
 * @param {string} nombreHoja El nombre de la hoja a obtener o crear.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} La hoja de cálculo encontrada o creada.
 */
function obtenerOCrearHoja(ss, nombreHoja) {
  let hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) {
    hoja = ss.insertSheet(nombreHoja);
  }
  return hoja;
}

function unitToKg_(unidadVenta) {
  const u = norm(unidadVenta);
  if (u === 'kilo' || u === 'kg') return 1;
  if (u === 'gramo' || u === 'g' || u === 'gramos') return 0.001;
  if (u === 'unidad' || u === 'unidades') return 1; // asumir factor ya está en Kg
  throw new Error(`Unidad Venta no soportada: "${unidadVenta}"`);
}

// =====================================================================================
// SERVIDOR WEB PARA DASHBOARD
// =====================================================================================

/**
 * Sirve la aplicación web del dashboard.
 * Se ejecuta cuando un usuario visita la URL de la aplicación.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('dashboard.html')
      .setTitle('Dashboard de Inventario 2.0');
}

/**
 * Lanza el dashboard v3 en un diálogo modal (ventana emergente).
 */

/**
 * Calcula y puebla la columna N de la hoja "Adquisiciones" con el total de la compra en unidades base (Kg).
 * Se ejecuta antes de mostrar el dashboard para asegurar que los datos estén visibles y actualizados.
 */
function populateTotalCompradoEnAdquisiciones() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adqSheet = ss.getSheetByName(HOJA_ADQUISICIONES);
  const skuSheet = ss.getSheetByName(HOJA_SKU);

  if (!adqSheet || adqSheet.getLastRow() < 2 || !skuSheet || skuSheet.getLastRow() < 2) {
    Logger.log("No se pudieron encontrar las hojas Adquisiciones o SKU, o están vacías.");
    return;
  }

  // 1. Crear mapa de búsqueda desde SKU
  const mapaCompraSku = new Map();
  const skuValues = skuSheet.getRange(2, 1, skuSheet.getLastRow() - 1, 4).getValues(); // A:D
  skuValues.forEach(fila => {
    const [, productoBase, formatoAdq, cantAdq] = fila; // B, C, D
    if (productoBase && formatoAdq) {
      const pBase = norm(productoBase);
      const fAdq = formatoAdq.toString().trim();
      mapaCompraSku.set(`${pBase}-${fAdq}`, parseFloat(String(cantAdq).replace(',','.')) || 0);
    }
  });

  // 2. Preparar los cálculos para la columna N
  const adqValues = adqSheet.getRange(2, 1, adqSheet.getLastRow() - 1, 4).getValues(); // A:D
  const resultadosColumnaN = adqValues.map(fila => {
    const [, productoBase, formatoCompra, cantidadCompradaStr] = fila; // B, C, D
    if (!productoBase || !formatoCompra) {
      return [""]; // Devuelve un array con un elemento para que se escriba una celda vacía
    }

    const cantidadComprada = parseFloat(String(cantidadCompradaStr).replace(',','.')) || 0;

    const formatoAdq = _getFormatoAdquisicionBase(formatoCompra);

    const claveCompra = `${productoBase.toString().trim()}-${formatoAdq}`;
    const cantAdquisicion = mapaCompraSku.get(claveCompra);

    if (cantAdquisicion) {
      const compraEnUnidadBase = cantidadComprada * cantAdquisicion;
      return [compraEnUnidadBase];
    } else {
      return [""]; // Si no se encuentra, celda vacía
    }
  });

  // 3. Escribir los resultados en la columna N
  // Primero, el encabezado si es necesario. La columna N es la 14.
  const headerCell = adqSheet.getRange("N1");
  if (headerCell.getValue() !== "Total Comprado (Kg)") {
    headerCell.setValue("Total Comprado (Kg)").setFontWeight("bold");
  }

  // Limpiar contenido anterior y escribir nuevos valores
  if (resultadosColumnaN.length > 0) {
    const targetRange = adqSheet.getRange(2, 14, resultadosColumnaN.length, 1);
    targetRange.clearContent();
    targetRange.setValues(resultadosColumnaN);
  }
}

function showDashboard() {
  populateTotalCompradoEnAdquisiciones(); // Ejecutar el cálculo antes de mostrar

  const html = HtmlService.createHtmlOutputFromFile('dashboard.html')
      .setWidth(1200)
      .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard de Inventario v3');
}

/**
 * Extrae el formato de adquisición base de una cadena de texto.
 * Maneja casos como "Paquete (4 Kg)" -> "Paquete" y "Paquete de compra" -> "Paquete".
 * @param {string} formatoCompra El texto completo del formato.
 * @returns {string} El formato base normalizado.
 */
function _getFormatoAdquisicionBase(formatoCompra) {
  if (!formatoCompra) return '';
  const formatoStr = formatoCompra.toString();
  const match = formatoStr.match(/(.*) \(/);
  if (match && match[1]) {
    return match[1].trim();
  }
  return formatoStr.trim().split(' ')[0];
}


// =========================
// Helpers de normalización
// =========================

/**
 * Normaliza un texto de forma robusta: quita acentos, convierte a minúsculas,
 * recorta espacios y elimina espacios duplicados.
 * @param {string} text El texto a normalizar.
 * @returns {string} El texto normalizado.
 */
function normalizeText(text) {
  if (!text) return '';
  return text.toString()
    .trim()
    .toLowerCase()
    .normalize("NFD") // Descompone acentos y caracteres especiales
    .replace(/[\u0300-\u036f]/g, "") // Elimina los diacríticos (acentos)
    .replace(/\s+/g, ' '); // Reemplaza uno o más espacios por uno solo
}

function normalizeUnit(uRaw) {
  if (!uRaw) return '';
  const u = ('' + uRaw).trim().toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/[áä]/g,'a').replace(/[éë]/g,'e').replace(/[íï]/g,'i')
    .replace(/[óö]/g,'o').replace(/[úü]/g,'u');

  // equivalencias comunes
  if (/(^| )kg(s)?($)/.test(u) || /(kilo|kilos)/.test(u)) return 'kg';
  if (/^g(ramos)?$/.test(u) || /gramo(s)?/.test(u)) return 'g';
  if (/^l(itro)?s?$/.test(u) || /litro(s)?/.test(u)) return 'lt';
  if (/botella(s)?/.test(u)) return 'botella';
  if (/bandeja(s)?/.test(u)) return 'bandeja';
  if (/paquete(s)?/.test(u)) return 'paquete';
  if (/malla(s)?/.test(u)) return 'malla';
  if (/caja(s)?/.test(u)) return 'caja';
  if (/unidad(es)?|unid/.test(u)) return 'unidad';
  if (/trio/.test(u)) return 'trio';
  if (/docena/.test(u)) return 'docena';
  if (/envase/.test(u)) return 'envase';

  return u; // devolver tal cual si no reconocimos
}

/**
 * Parseo de "Formato de Compra".
 * Ejemplos de entrada:
 *  - "Kilo (1 Kg)"     -> {unit:'kg', qtyPerFormato: 1}
 *  - "Paquete (4 Kg)"  -> {unit:'kg', qtyPerFormato: 4}
 *  - "Caja (10 Unidad)"-> {unit:'unidad', qtyPerFormato: 10}
 *  - "Malla (14 Kg)"   -> {unit:'kg', qtyPerFormato: 14}
 */
function parseFormatoCompra(formatoRaw) {
  if (!formatoRaw) return { unit: '', qtyPerFormato: NaN, source: '' };
  const txt = ('' + formatoRaw).trim();

  // Intentar capturar "Algo (N Unidad)" o "Algo (N Kg)" etc.
  const m = txt.match(/\(([\d.,]+)\s*([A-Za-zÁÉÍÓÚáéíóúñÑ]+)\s*\)/);
  if (m) {
    const qty = parseFloat(m[1].replace(',', '.'));
    const insideUnit = normalizeUnit(m[2]);
    return { unit: insideUnit, qtyPerFormato: qty, source: 'parentesis' };
  }

  // Si no hay paréntesis, intentar casos simples: "Kilo", "Unidad", etc.
  const u = normalizeUnit(txt);
  if (u) return { unit: u, qtyPerFormato: 1, source: 'simple' };

  return { unit: '', qtyPerFormato: NaN, source: 'unknown' };
}

/**
 * Convierte una cantidad a la unidad de dashboard (unidad SKU) si es posible.
 * - Si las unidades ya coinciden -> cantidad * qtyPerFormato
 * - Si vienen en kg y piden kg -> cantidad * qtyPerFormato
 * - Si vienen en "paquete (X Kg)" y dashboard es kg -> cantidad * X
 * - Incompatibles (ej. kg -> unidad sin factor) => inconsistencia
 */
function convertAcquisitionToDashUnit(qtyCompra, parsedFormato, dashUnit) {
  const dash = normalizeUnit(dashUnit);
  const src = parsedFormato.unit; // ya normalizada
  const factor = parsedFormato.qtyPerFormato;

  // Conversión directa: misma unidad base
  if (dash === src && !isNaN(factor)) {
    return { qty: qtyCompra * factor, ok: true };
  }

  // Caso más común: formato empaquetado que especifica kg y dashboard usa kg
  if (src === 'kg' && dash === 'kg' && !isNaN(factor)) {
    return { qty: qtyCompra * factor, ok: true };
  }

  // Otros equivalentes triviales (litros, unidad)
  if (src === 'lt' && dash === 'lt' && !isNaN(factor)) {
    return { qty: qtyCompra * factor, ok: true };
  }
  if (src === 'unidad' && dash === 'unidad' && !isNaN(factor)) {
    return { qty: qtyCompra * factor, ok: true };
  }

  // Si el formato es "Caja (10 Unidad)" y dashboard es "unidad"
  if (src === 'unidad' && dash === 'unidad' && !isNaN(factor)) {
    return { qty: qtyCompra * factor, ok: true };
  }

  // Si el formato NO trae especificación dentro de paréntesis (factor)
  if (isNaN(factor)) {
    return { qty: NaN, ok: false, reason: `Formato de compra sin factor interpretable` };
  }

  // Unidades no compatibles / no cubiertas
  return {
    qty: NaN,
    ok: false,
    reason: `Unidad incompatible: formato "${src}" -> dashboard "${dash}"`
  };
}

// Índice de columna por encabezado (robusto a cambios de orden)
function indexByHeader(sheet, headerName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => (''+h).trim());
  const idx = headers.findIndex(h => h.toLowerCase() === (''+headerName).trim().toLowerCase());
  return idx; // -1 si no existe
}

/**
 * Devuelve un Map<string, string> de estados por producto base.
 * Estados posibles: 'pendiente' (por defecto), 'verificando', 'aprobado'
 */
function getEstadosProductos_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Estados');
  if (!sh || sh.getLastRow() < 2) return new Map();
  const rows = sh.getRange(2,1, sh.getLastRow()-1, 5).getValues();
  const map = new Map();
  rows.forEach(([base, estado]) => {
    // Usar el nombre normalizado como clave para que no haya distinción de mayúsculas/minúsculas
    if (base) map.set(norm(base), (estado || 'pendiente').toString().toLowerCase());
  });
  return map;
}

/** Devuelve los estados como objeto { [baseProduct]: estado } para el dashboard. */
function getEstadosParaUI() {
  const m = getEstadosProductos_();
  const obj = {};
  m.forEach((v,k)=> obj[k]=v);
  return obj;
}

/**
 * Actualiza el estado de un producto base.
 * @param {string} baseProduct
 * @param {'pendiente'|'verificando'|'aprobado'} estado
 * @param {string=} notas
 */
function setEstadoProducto(baseProduct, estado, notas) {
  if (!baseProduct) throw new Error('Producto Base vacío.');
  const baseProductNorm = norm(baseProduct);
  estado = (estado || 'pendiente').toLowerCase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Estados') || ss.insertSheet('Estados');

  // Asegurar headers
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,5).setValues([['Producto Base','Estado','Notas','Usuario','Timestamp']]);
  }

  // Buscar si ya existe
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    const data = sh.getRange(2,1,lastRow-1,5).getValues();
    for (let i=0;i<data.length;i++){
      if (norm(data[i][0]) === baseProductNorm){
        sh.getRange(i+2,2).setValue(estado);
        if (notas !== undefined) sh.getRange(i+2,3).setValue(notas);
        sh.getRange(i+2,4).setValue(Session.getActiveUser().getEmail() || 'Sistema');
        sh.getRange(i+2,5).setValue(new Date());
        return {ok:true};
      }
    }
  }

  // Insertar nueva fila
  sh.appendRow([
    baseProduct, // Guardar el nombre original
    estado,
    notas || '',
    Session.getActiveUser().getEmail() || 'Sistema',
    new Date()
  ]);
  return {ok:true};
}

// =====================================================================================
// LÓGICA DE CÁLCULO DE VENTAS (NUEVO)
// =====================================================================================

/**
 * Calcula las "Ventas del Día" de forma determinística desde las hojas Orders y SKU.
 * Esta función procesa todas las filas de Orders, excepto las marcadas como eliminadas.
 * Normaliza textos, maneja errores de parseo y agrupa las ventas por Producto Base.
 *
 * @returns {{ventasPorProductoBase: Map<string, number>, filasIgnoradas: number}}
 * Objeto con un mapa de ventas acumuladas por producto base y el conteo de filas ignoradas.
 */
function calculateVentasDelDia() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaSku = ss.getSheetByName(HOJA_SKU);
  const hojaOrders = ss.getSheetByName(HOJA_ORDERS);

  if (!hojaSku || !hojaOrders) {
    throw new Error("No se encontraron las hojas SKU u Orders.");
  }

  // Paso 2a: Construir el mapa de SKU.
  const productToSkuMap = new Map();
  const datosSku = hojaSku.getRange(2, 1, hojaSku.getLastRow() - 1, 7).getValues(); // A:G

  datosSku.forEach(row => {
    const nombreProducto = row[0]; // Col A
    const productoBase = row[1];   // Col B
    const cantidadVentaRaw = row[6]; // Col G

    if (nombreProducto) {
      const cantidadVenta = parseFloat(String(cantidadVentaRaw || '0').replace(',', '.')) || 0;
      productToSkuMap.set(normalizeText(nombreProducto), {
        productoBaseNormalizado: normalizeText(productoBase),
        cantidadVenta: cantidadVenta
      });
    }
  });

  // Paso 2b y 2c: Procesar la hoja Orders, filtrar y agregar.
  const ventasPorProductoBase = new Map();
  let filasIgnoradas = 0;
  const datosOrders = hojaOrders.getRange(2, 1, hojaOrders.getLastRow() - 1, 26).getValues(); // A:Z

  datosOrders.forEach(row => {
    const cantidadRaw = String(row[10] || ''); // Col K: Cantidad

    if (cantidadRaw.trim().toUpperCase().startsWith('E')) {
      filasIgnoradas++;
      return; // Ignorar fila
    }

    const nombreProducto = row[9];  // Col J: Nombre Producto
    const productoBase = row[25]; // Col Z: Producto Base

    if (!nombreProducto || !productoBase) {
      return; // Si falta el nombre o el base, no se puede procesar
    }

    const cantidad = parseFloat(cantidadRaw.replace(',', '.')) || 0;
    const nombreProductoNormalizado = normalizeText(nombreProducto);
    const productoBaseNormalizado = normalizeText(productoBase);

    const skuInfo = productToSkuMap.get(nombreProductoNormalizado);

    if (skuInfo) {
      const ventaBaseSKU = cantidad * skuInfo.cantidadVenta;
      const totalActual = ventasPorProductoBase.get(productoBaseNormalizado) || 0;
      ventasPorProductoBase.set(productoBaseNormalizado, totalActual + ventaBaseSKU);
    }
    // Si no hay SKU, la venta es 0 y no se suma, como se especificó.
  });

  return { ventasPorProductoBase, filasIgnoradas };
}


/***** CONFIG *****/
const CFG = {
  timezone: 'America/Santiago',
  sheetOrders: 'Orders',
  sheetSKU: 'SKU',
  sheetVentasBase: 'VENTAS_BASE_HOY',
  estadosPermitidos: ['Procesando', 'En Espera de Pago'], // ajustar si es necesario
  // Encabezados esperados:
  ORDERS_HEADERS: {
    estado: 'Estado',
    fecha: 'Fecha',
    nombreProducto: 'Nombre Producto',
    cantidad: 'Cantidad',
    productoBase: 'Producto Base', // Se añade para leer la columna Z de Orders
  },
  SKU_HEADERS: {
    nombreProducto: 'Nombre Producto',
    productoBase: 'Producto Base',
    cantVenta: 'Cantidad Venta',
    unidadVenta: 'Unidad Venta',
  }
};

/***** HELPERS *****/

function getHeaderIndexes_(headerRow, headerMap) {
  const idx = {};
  const mapKeys = Object.keys(headerMap);
  mapKeys.forEach(k => {
    const wanted = norm(headerMap[k]);
    const pos = headerRow.findIndex(h => norm(h) === wanted);
    if (pos === -1) throw new Error(`No se encontró la columna "${headerMap[k]}"`);
    idx[k] = pos;
  });
  return idx;
}

function startOfToday_(tz) {
  const now = new Date();
  const str = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  return new Date(`${str}T00:00:00`);
}
function startOfTomorrow_(tz) {
  const t0 = startOfToday_(tz);
  return new Date(t0.getTime() + 24*60*60*1000);
}

/***** CORE *****/
/**
 * Construye un mapa desde la hoja SKU.
 * La clave es el "Nombre Producto" y el valor contiene el factor de conversión y la unidad.
 * El "Producto Base" ya no se necesita en el mapa porque se lee directamente de la hoja Orders.
 */
function buildSkuMap_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.sheetSKU);
  if (!sh) throw new Error(`No existe la hoja ${CFG.sheetSKU}`);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return {};

  const hdr = values[0].map(String);
  const H = getHeaderIndexes_(hdr, CFG.SKU_HEADERS);
  const map = {};
  const warnings = [];

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const name = norm(row[H.nombreProducto]);
    const base = row[H.productoBase];
    const fac  = row[H.cantVenta];
    const uni  = row[H.unidadVenta];

    if (!name || !base) continue;
    if (map[name]) warnings.push(`Duplicado en SKU: ${row[H.nombreProducto]}`);

    map[name] = {
      // "base" ya no es necesario aquí, se leerá de la hoja Orders.
      factor: Number(fac) || 0,
      unidad: uni
    };
  }

  if (warnings.length) Logger.log(warnings.join('\n'));
  return map;
}

/**
 * Calcula las ventas por "Producto Base" para un rango de fechas y estados de pedido.
 * La lógica ahora toma el "Producto Base" directamente de la hoja "Orders" (columna Z).
 */
function getVentasPorBase_EntreFechas_(fromDate, toDate, estadosPermitidos) {
  const skuMap = buildSkuMap_();

  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.sheetOrders);
  if (!sh) throw new Error(`No existe la hoja ${CFG.sheetOrders}`);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { ventasPorBase:{}, missing:[] };

  const hdr = values[0].map(String);
  const H = getHeaderIndexes_(hdr, CFG.ORDERS_HEADERS);
  const estadosOK = new Set(estadosPermitidos.map(norm));

  const ventasPorBase = {};
  const missing = [];
  const badUnits = [];

  for (let r = 1; r < values.length; r++) {
    const row = values[r];

    const estado = norm(row[H.estado]);
    if (!estadosOK.has(estado)) continue;

    const fechaRaw = row[H.fecha];
    if (!fechaRaw) continue;
    const fecha = (fechaRaw instanceof Date) ? fechaRaw : new Date(fechaRaw);
    if (!(fecha instanceof Date) || isNaN(fecha)) continue;
    if (!(fecha >= fromDate && fecha < toDate)) continue;

    const cantidadRaw = String(row[H.cantidad]);
    // Si la cantidad contiene 'E', es un item eliminado y se debe omitir.
    if (cantidadRaw.toUpperCase().includes('E')) {
      continue;
    }
    const nombre = norm(row[H.nombreProducto]);
    const cantidad = Number(cantidadRaw.replace(',', '.')) || 0;
    const productoBase = row[H.productoBase]; // Se obtiene el "Producto Base" de la columna Z.

    if (!nombre || cantidad <= 0 || !productoBase) continue;

    const s = skuMap[nombre];
    if (!s) { missing.push(row[H.nombreProducto]); continue; }

    let mult;
    try {
      mult = unitToKg_(s.unidad);
    } catch (e) {
      badUnits.push(`${row[H.nombreProducto]} -> Unidad "${s.unidad}"`);
      continue;
    }

    const kg = cantidad * (Number(s.factor) || 0) * mult;
    const baseKey = norm(productoBase); // Se usa el "Producto Base" de la orden.
    if (!ventasPorBase[baseKey]) ventasPorBase[baseKey] = 0;
    ventasPorBase[baseKey] += kg;
  }

  if (missing.length) Logger.log('SKU faltante para: \n' + Array.from(new Set(missing)).join('\n'));
  if (badUnits.length) Logger.log('Unidad no soportada en: \n' + Array.from(new Set(badUnits)).join('\n'));

  return { ventasPorBase, missing: Array.from(new Set(missing)) };
}

function getVentasPorBaseHoy_() {
  const tz = CFG.timezone;
  const from = startOfToday_(tz);
  const to   = startOfTomorrow_(tz);
  return getVentasPorBase_EntreFechas_(from, to, CFG.estadosPermitidos);
}

/***** OUTPUTS *****/
function writeVentasBaseHoy() {
  const ss = SpreadsheetApp.getActive();
  const shOut = ss.getSheetByName(CFG.sheetVentasBase) || ss.insertSheet(CFG.sheetVentasBase);

  // calcular
  const { ventasPorBase } = getVentasPorBaseHoy_();

  // volcar
  shOut.clear();
  const rows = [['Producto Base', 'Ventas Kg']];
  Object.keys(ventasPorBase).sort().forEach(k => {
    const baseOriginalCase = k; // normalizado; si necesitan el case original, guardar en skuMap
    rows.push([baseOriginalCase, Number(ventasPorBase[k])]);
  });
  if (rows.length === 1) rows.push(['(sin ventas)', 0]);

  shOut.getRange(1,1,rows.length,rows[0].length).setValues(rows);
  shOut.autoResizeColumns(1, 2);
}

/***** INTEGRACIÓN CON EL DASHBOARD *****/
// Si el Dashboard es HTML (HtmlService) y ya existe una función que arma los datos,
// exportar esta función y sumarle "ventasKg" por Producto Base.
function getVentasPorBaseHoy_JSON() {
  const data = getVentasPorBaseHoy_().ventasPorBase;
  // Devuelve: { 'acelga': 6, 'naranja': 27, ... }
  return data;
}

/** Config de compras por día (activa SOLO_HOY si la hoja Adquisiciones tiene columna "Fecha") */
const SOLO_HOY = false;

/** Devuelve el inicio del día en la zona horaria dada */
function _startOfDay_(tz) {
  const now = new Date();
  const s = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  return new Date(`${s}T00:00:00`);
}
function _startOfTomorrow_(tz) {
  const t0 = _startOfDay_(tz);
  return new Date(t0.getTime() + 24 * 60 * 60 * 1000);
}

/**
 * Suma compras por Producto Base con la regla F + H − E.
 * Usa nombres de encabezado para encontrar columnas. Si existe "Fecha" y SOLO_HOY = true, filtra solo el día actual.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} adqSheet Hoja "Adquisiciones"
 * @returns {Map<string, number>} comprasPorBase
 */
function getComprasPorBase_SUMARSI_(adqSheet) {
  const comprasPorBase = new Map();
  if (!adqSheet || adqSheet.getLastRow() < 2) return comprasPorBase;

  const values = adqSheet.getDataRange().getValues();
  const headers = values[0].map(h => ('' + h).trim());
  const idxBase = headers.indexOf('Producto Base');

  // Ajusta estos nombres según tus encabezados reales
  const idxE = headers.indexOf('Inventario Actual');      // Col. E
  const idxF = headers.indexOf('Necesidad de Venta');     // Col. F
  const idxH = headers.indexOf('Inventario al Finalizar'); // Col. H

  if (idxBase === -1 || idxE === -1 || idxF === -1 || idxH === -1) {
    Logger.log('No se encontraron las columnas requeridas (Producto Base / Inventario Actual / Necesidad de Venta / Inventario al Finalizar).');
    return comprasPorBase;
  }

  // Filtrado opcional por fecha
  const idxFecha = headers.indexOf('Fecha');
  let t0 = null, t1 = null;
  if (SOLO_HOY && idxFecha !== -1) {
    t0 = _startOfDay_(TIMEZONE);
    t1 = _startOfTomorrow_(TIMEZONE);
  }

  for (let r = 1; r < values.length; r++) {
    if (SOLO_HOY && idxFecha !== -1) {
      const fr = values[r][idxFecha];
      if (!fr) continue;
      const f = (fr instanceof Date) ? fr : new Date(fr);
      if (!(f >= t0 && f < t1)) continue;
    }

    const base = (values[r][idxBase] ?? '').toString().trim();
    if (!base) continue;

    const e = parseFloat(('' + values[r][idxE]).replace(',', '.')) || 0;
    const f = parseFloat(('' + values[r][idxF]).replace(',', '.')) || 0;
    const h = parseFloat(('' + values[r][idxH]).replace(',', '.')) || 0;
    const total = f + h - e;

    if (total !== 0) {
      comprasPorBase.set(base, (comprasPorBase.get(base) || 0) + total);
    }
  }
  return comprasPorBase;
}

// =====================================
// getDashboardData CON COMPRAS INCLUIDO
// =====================================

/**
 * Calcula las compras del día según la lógica solicitada:
 * Para cada adquisición, multiplica la 'Cantidad a Comprar' por el multiplicador del SKU.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} adqSheet Hoja "Adquisiciones"
 * @param {GoogleAppsScript.Spreadsheet.Sheet} skuSheet Hoja "SKU"
 * @returns {Map<string, number>} comprasPorBase
 */
function getComprasPorBase_Correcto(adqSheet, skuSheet) {
    const comprasPorBase = new Map();
    if (!adqSheet || adqSheet.getLastRow() < 2 || !skuSheet || skuSheet.getLastRow() < 2) {
        return comprasPorBase;
    }

    // 1. Crear mapa de búsqueda desde SKU
    const mapaCompraSku = new Map();
    const skuValues = skuSheet.getRange(2, 1, skuSheet.getLastRow() - 1, 4).getValues(); // A:D
    skuValues.forEach(fila => {
        const [, productoBase, formatoAdq, cantAdq] = fila; // B, C, D
        if (productoBase && formatoAdq) {
            const pBase = norm(productoBase);
            const fAdq = formatoAdq.toString().trim();
            mapaCompraSku.set(`${pBase}-${fAdq}`, parseFloat(String(cantAdq).replace(',','.')) || 0);
        }
    });

    // 2. Procesar adquisiciones con filtro de fecha
    const adqData = adqSheet.getDataRange().getValues();
    const adqHeaders = adqData[0].map(h => norm(h));
    const idxFecha = adqHeaders.indexOf('fecha');
    const idxProductoBase = adqHeaders.indexOf('producto base');
    const idxFormatoCompra = adqHeaders.indexOf('formato de compra');
    const idxCantidad = adqHeaders.indexOf('cantidad a comprar');

    if (idxProductoBase === -1 || idxFormatoCompra === -1 || idxCantidad === -1) {
        Logger.log("ADVERTENCIA: Faltan columnas esenciales en 'Adquisiciones'. Se requieren 'Producto Base', 'Formato de Compra' y 'Cantidad a Comprar'.");
        return comprasPorBase;
    }

    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);

    for (let i = 1; i < adqData.length; i++) {
        const row = adqData[i];
        const fechaAdqRaw = idxFecha !== -1 ? row[idxFecha] : null;

        if (fechaAdqRaw) {
            const fechaAdq = new Date(fechaAdqRaw);
            if (fechaAdq.getTime() < hoy.getTime()) {
                continue; // Omitir si es de un día anterior
            }
        }

        const productoBase = row[idxProductoBase];
        const formatoCompra = row[idxFormatoCompra];
        const cantidadCompradaStr = row[idxCantidad];

        if (!productoBase || !formatoCompra) continue;

        const pBase = norm(productoBase);
        const cantidadComprada = parseFloat(String(cantidadCompradaStr).replace(',','.')) || 0;
        const formatoAdq = _getFormatoAdquisicionBase(formatoCompra);
        const claveCompra = `${pBase}-${formatoAdq}`;
        const cantAdquisicion = mapaCompraSku.get(claveCompra);

        if (cantAdquisicion) {
            const compraEnUnidadBase = cantidadComprada * cantAdquisicion;
            comprasPorBase.set(pBase, (comprasPorBase.get(pBase) || 0) + compraEnUnidadBase);
        }
    }

    return comprasPorBase;
}

function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const skuSheet = ss.getSheetByName(HOJA_SKU);
  const reporteSheet = ss.getSheetByName(HOJA_REPORTE_HOY);
  const histSheet = ss.getSheetByName(HOJA_HISTORICO);

  if (!skuSheet || !reporteSheet) {
    throw new Error(`No se encontraron las hojas requeridas: ${HOJA_SKU} o ${HOJA_REPORTE_HOY}.`);
  }

  if (skuSheet.getLastRow() < 2) {
    return { inventory: [], estados: {}, error: `La hoja "${HOJA_SKU}" está vacía o solo contiene encabezados.` };
  }
  if (reporteSheet.getLastRow() < 2) {
    return { inventory: [], estados: {}, error: `La hoja "${HOJA_REPORTE_HOY}" está vacía. Ejecute el cálculo diario primero.` };
  }

  // === 1) Mapa SKU: {original, unit, category} por producto base normalizado
  const skuData = skuSheet.getRange(2, 1, Math.max(0, skuSheet.getLastRow() - 1), 8).getValues();
  const baseInfoMap = new Map();
  skuData.forEach(r => {
    const productoBase = r[1]; // Col B
    const categoria = r[5];    // Col F
    const unidadVenta = r[7];  // Col H
    if (productoBase && !baseInfoMap.has(norm(productoBase))) {
      baseInfoMap.set(norm(productoBase), { original: productoBase, unit: unidadVenta || '', category: categoria || '' });
    }
  });

  // === 2) Leer datos del "Reporte Hoy" y crear un mapa (incluyendo Stock Real)
  const reporteData = reporteSheet.getRange(2, 1, reporteSheet.getLastRow() - 1, 6).getValues(); // Leer hasta la columna F
  const reporteMap = new Map();
  reporteData.forEach(row => {
    const [productoBase, invAyer, compras, ventas, invEstimado, stockReal] = row;
    if (productoBase) {
      reporteMap.set(norm(productoBase), {
        lastInventory: parseFloat(String(invAyer || '0').replace(',', '.')) || 0,
        purchases: parseFloat(String(compras || '0').replace(',', '.')) || 0,
        sales: parseFloat(String(ventas || '0').replace(',', '.')) || 0,
        expectedStock: parseFloat(String(invEstimado || '0').replace(',', '.')) || 0,
        stockReal: stockReal // Guardar el valor de Stock Real
      });
    }
  });

  // === 3) Armar inventory[] para el Dashboard
  const inventory = [];
  baseInfoMap.forEach((info, key) => {
    const reporteInfo = reporteMap.get(key);
    const hasError = !reporteInfo;

    inventory.push({
      baseProduct: info.original,
      lastInventory: hasError ? 0 : reporteInfo.lastInventory,
      purchases: hasError ? 0 : reporteInfo.purchases,
      sales: hasError ? 0 : reporteInfo.sales,
      expectedStock: hasError ? 0 : reporteInfo.expectedStock,
      unit: info.unit,
      category: info.category,
      error: hasError,
      errorMsg: hasError ? "Producto no encontrado en el reporte de hoy." : ""
    });
  });

  // === 4) Lógica de estados (simplificada) ---
  // La fuente de la verdad para el estado de un producto es la hoja "Estados".
  // Esta función simplemente lee los estados persistidos y los devuelve.
  // El reseteo de estados "aprobados" ahora se maneja exclusivamente en `calcularInventarioDiario`.
  const persistedStates = getEstadosParaUI(); // Devuelve { [normalizedBase]: state }
  const finalStates = {};

  baseInfoMap.forEach((info, baseNorm) => {
    // El nombre original del producto (con mayúsculas/minúsculas) se usa como clave en el objeto final.
    // El estado se busca usando el nombre normalizado. Si no se encuentra, se asume 'pendiente'.
    finalStates[info.original] = persistedStates[baseNorm] || 'pendiente';
  });

  return {
    inventory: inventory,
    sales: [], // Estos ya no se calculan aquí, se devuelven vacíos
    acquisitions: [], // Estos ya no se calculan aquí, se devuelven vacíos
    estados: finalStates
  };
}

/**
 * Guarda las actualizaciones de stock desde el dashboard.
 * Esta función ahora solo actualiza la columna "Stock Real" en "Reporte Hoy".
 * La actualización de "Estados" se maneja por el disparador onEdit.
 * La actualización de "Inventario Histórico" se maneja por "CERRAR DIA".
 */
function saveStockUpdates(updates) {
  if (!updates || !Array.isArray(updates) || updates.length === 0) {
    return { success: false, error: "No se proporcionaron datos para actualizar." };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const reporteSheet = ss.getSheetByName(HOJA_REPORTE_HOY);

    if (!reporteSheet) {
      throw new Error(`La hoja "${HOJA_REPORTE_HOY}" no fue encontrada.`);
    }

    // 1. Leer los datos de la hoja de reporte para encontrar las filas correctas.
    const reporteRange = reporteSheet.getDataRange();
    const reporteData = reporteRange.getValues();

    // 2. Crear un mapa para buscar la fila de cada producto rápidamente.
    const productRowMap = new Map();
    reporteData.forEach((row, index) => {
      const productName = row[0]; // Columna A
      if (productName) {
        productRowMap.set(norm(productName), index); // Guardar índice 0-based
      }
    });

    // 3. Procesar las actualizaciones en la matriz en memoria.
    const updatedProductsForFe = []; // Para la respuesta al frontend
    updates.forEach(update => {
      const { productBase, quantity } = update;
      const productNorm = norm(productBase);

      if (productRowMap.has(productNorm)) {
        const rowIndex = productRowMap.get(productNorm);
        // Columna F (Stock Real) es el índice 5 en el array 0-based.
        reporteData[rowIndex][5] = (quantity === null || quantity === undefined) ? '' : quantity;

        // El frontend todavía espera un objeto de respuesta, así que lo preparamos.
        updatedProductsForFe.push({
          productBase: productBase,
          quantity: quantity,
          // El estado se actualizará visualmente por el `onEdit` y el siguiente `getDashboardData`,
          // pero podemos devolver 'aprobado' para una respuesta más rápida de la UI.
          state: (quantity !== null && quantity !== '') ? 'aprobado' : 'pendiente'
        });
      }
    });

    // 4. Escribir toda la matriz de datos actualizada de vuelta a la hoja.
    reporteRange.setValues(reporteData);

    return { success: true, message: `${updates.length} productos actualizados en 'Reporte Hoy'.`, updatedProducts: updatedProductsForFe };

  } catch (e) {
    Logger.log(`Error en saveStockUpdates: ${e.stack}`);
    return { success: false, error: e.message };
  }
}

// =====================================================================================
// FUNCIONES PARA DASHBOARD "CONTACTAR CLIENTE" (NUEVA VERSIÓN)
// =====================================================================================

/***** CONFIG *****/
const SHEET_NAME = 'Orders';
const DEFAULT_COUNTRY_CODE = '56'; // Chile
const DEFAULT_FORM_URL = 'https://forms.gle/8x3bzfwL2oZyqcou6';

const DEFAULT_TEMPLATE = [
  '👋 ¡Hola! Te contactamos desde Santiago Natural Food 🙂.',
  '',
  'Te informamos que lamentablemente no pudimos enviar:',
  '',
  '{PRODUCTOS}',
  'Pedido N° {PEDIDO}',
  '',
  '💳 Para solucionar este inconveniente, por favor, completa el siguiente formulario para elegir la forma en que deseas tu devolución',
  '🔗 {FORM_URL}',
  '',
  '🙌 Si tienes cualquier duda o necesitas asistencia adicional, estamos aquí para ayudarte. ¡Gracias por tu comprensión y confianza!'
].join('\n');

/***** MENU (function defined above, this is the implementation) *****/
function openContactarCliente() {
  const html = HtmlService.createTemplateFromFile('ContactarCliente').evaluate()
    .setTitle('Contactar Cliente')
    .setWidth(1200)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Contactar Cliente');
}

/***** DATA LAYER *****/
function fetchOrdersAggregated() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`No existe la hoja "${SHEET_NAME}"`);

  const values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) return { orders: [], template: DEFAULT_TEMPLATE, formURL: DEFAULT_FORM_URL, productNames: [] };

  // Mapear encabezados sin depender del orden de columnas
  const headers = values[0].map(h => (h || '').toString().trim());
  const col = (name) => headers.indexOf(name);

  const idx = {
    pedido: col('Número de pedido'),
    nombre: col('Nombre completo'),
    email: col('Email'),
    telefono: col('Teléfono'),
    direccion: col('Direccion'),
    depto: col('Depto/Condominio'),
    comuna: col('Comuna'),
    estado: col('Estado'),
    fecha: col('Fecha'),
    producto: col('Nombre Producto'),
    cantidad: col('Cantidad')
  };

  // Validación mínima
  const required = ['Número de pedido','Nombre completo','Teléfono','Nombre Producto','Cantidad'];
  required.forEach(k => {
    if (headers.indexOf(k) === -1) throw new Error(`Falta la columna requerida: "${k}"`);
  });

  /** Agrupar por pedido y recolectar nombres de productos **/
  const byOrder = new Map();
  const productNames = new Set();

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const nPedido = (row[idx.pedido] || '').toString().trim();
    if (!nPedido) continue;

    if (!byOrder.has(nPedido)) {
      byOrder.set(nPedido, {
        pedido: nPedido,
        nombre: (row[idx.nombre] || '').toString().trim(),
        email: idx.email >= 0 ? (row[idx.email] || '').toString().trim() : '',
        telefono: (row[idx.telefono] || '').toString().trim(),
        direccion: idx.direccion >= 0 ? (row[idx.direccion] || '').toString().trim() : '',
        depto: idx.depto >= 0 ? (row[idx.depto] || '').toString().trim() : '',
        comuna: idx.comuna >= 0 ? (row[idx.comuna] || '').toString().trim() : '',
        estado: idx.estado >= 0 ? (row[idx.estado] || '').toString().trim() : '',
        fecha: idx.fecha >= 0 ? (row[idx.fecha] || '').toString().trim() : '',
        items: []
      });
    }

    const item = {
      nombreProducto: (row[idx.producto] || '').toString().trim(),
      cantidad: (row[idx.cantidad] || '').toString().trim()
    };

    if (item.nombreProducto) {
      productNames.add(item.nombreProducto);
    }
    byOrder.get(nPedido).items.push(item);
  }

  // Orden opcional por número de pedido descendente
  const orders = Array.from(byOrder.values()).sort((a, b) => {
    const na = Number(a.pedido), nb = Number(b.pedido);
    if (isNaN(na) || isNaN(nb)) return (''+b.pedido).localeCompare(''+a.pedido);
    return nb - na;
  });

  const uniqueProductNames = Array.from(productNames).sort();

  return { orders, template: DEFAULT_TEMPLATE, formURL: DEFAULT_FORM_URL, productNames: uniqueProductNames };
}

/***** HELPERS *****/
function normalizePhoneForWa(phoneRaw) {
  // Solo dígitos
  const digits = (phoneRaw || '').toString().replace(/\D+/g, '');
  if (!digits) return '';

  // Si ya viene con 56* lo dejamos, si no, le anteponemos 56
  if (digits.startsWith(DEFAULT_COUNTRY_CODE)) return digits;
  // Caso típico Chile: 9 dígitos móviles sin prefijo
  if (digits.length === 8 || digits.length === 9) return DEFAULT_COUNTRY_CODE + digits;
  // Si viene con 0 inicial (líneas fijas antiguas), lo removemos y anteponemos 56
  return DEFAULT_COUNTRY_CODE + digits.replace(/^0+/, '');
}

function buildWaLink(phoneRaw, text, useWeb) {
  const phone = normalizePhoneForWa(phoneRaw);
  const encoded = encodeURIComponent(text || '');
  const domain = useWeb ? 'web.whatsapp.com' : 'api.whatsapp.com';
  // api.whatsapp.com es más tolerante y recomendado para el link que abre la app
  return `https://${domain}/send?phone=${phone}&text=${encoded}`;
}

/***** Exponer utilidades a front *****/
function getWhatsAppLinkFromTemplate(order, selectedItems, template, formURL) {
  const items = selectedItems || [];
  const formattedProds = items.map(item => {
    return `Producto *${item.name}*\nCantidad *${item.quantity}*`;
  }).join('\n\n');

  const tpl = (template || DEFAULT_TEMPLATE)
    .replaceAll('{PRODUCTOS}', formattedProds)
    .replaceAll('{PEDIDO}', order?.pedido ?? '')
    .replaceAll('{NOMBRE}', order?.nombre ?? '')
    .replaceAll('{DIRECCION}', order?.direccion ?? '')
    .replaceAll('{COMUNA}', order?.comuna ?? '')
    .replaceAll('{FORM_URL}', formURL || DEFAULT_FORM_URL);

  const link = buildWaLink(order?.telefono || '', tpl, false);
  const webLink = buildWaLink(order?.telefono || '', tpl, true);
  return { message: tpl, link, webLink };
}
