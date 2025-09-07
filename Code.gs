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
 * Se ejecuta cuando se abre la hoja de cálculo.
 * Crea un menú personalizado para ejecutar las funciones principales.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Inventario 2.0')
      .addItem('Abrir Dashboard', 'showDashboard')
      .addSeparator()
      .addItem('1. Configurar Hojas y Fórmulas', 'setup')
      .addSeparator()
      .addItem('2. Calcular Inventario de Hoy (Manual)', 'calcularInventarioDiario')
      .addSeparator()
      .addItem('3. Activar/Actualizar Trigger Diario', 'crearDisparadorDiario')
      .addSeparator()
      .addItem('Crear Historico', 'crearInventarioHistoricoDePrueba')
      .addToUi();
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
  const encabezadosHistorico = ["Timestamp", "Producto Base", "Cantidad", "Stock Real", "Unidad Venta"];
  // Solo escribir encabezados si la fila 1 está vacía
  if (hojaHistorico.getRange("A1").getValue() === "") {
      hojaHistorico.getRange(1, 1, 1, encabezadosHistorico.length).setValues([encabezadosHistorico]).setFontWeight("bold");
  }

  // --- 5. Hoja de Reporte Hoy ---
  const hojaReporte = obtenerOCrearHoja(ss, HOJA_REPORTE_HOY);
  const encabezadosReporte = ["Producto Base", "Inventario Ayer", "Compras del Día", "Ventas del Día", "Inventario Hoy", "Stock Real"];
  if (hojaReporte.getRange("A1").getValue() === "") {
    hojaReporte.getRange(1, 1, 1, encabezadosReporte.length).setValues([encabezadosReporte]).setFontWeight("bold");
  }

  // --- 6. Hoja de Discrepancias ---
  const hojaDiscrepancias = obtenerOCrearHoja(ss, "Discrepancias");
  const encabezadosDiscrepancias = ["Timestamp", "Producto Base", "Inventario Estimado", "Inventario Real", "Discrepancia"];
   if (hojaDiscrepancias.getRange("A1").getValue() === "") {
    hojaDiscrepancias.getRange(1, 1, 1, encabezadosDiscrepancias.length).setValues([encabezadosDiscrepancias]).setFontWeight("bold");
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
    // --- 1. OBTENER DATOS ---
    const hojaSku = ss.getSheetByName(HOJA_SKU);
    const hojaOrders = ss.getSheetByName(HOJA_ORDERS);
    const hojaAdquisiciones = ss.getSheetByName(HOJA_ADQUISICIONES);
    const hojaHistorico = ss.getSheetByName(HOJA_HISTORICO);
    const hojaReporteHoy = ss.getSheetByName(HOJA_REPORTE_HOY);

    // Obtener datos de las hojas, omitiendo encabezados (fila 1)
    const datosSku = hojaSku.getRange("A2:K" + hojaSku.getLastRow()).getValues();
    const datosOrders = hojaOrders.getRange("A2:K" + hojaOrders.getLastRow()).getValues();
    const datosAdquisiciones = hojaAdquisiciones.getRange("A2:M" + hojaAdquisiciones.getLastRow()).getValues();
    const datosHistorico = hojaHistorico.getLastRow() > 1 ? hojaHistorico.getRange("A2:E" + hojaHistorico.getLastRow()).getValues() : [];

    // --- 2. PREPARAR MAPAS DE BÚSQUEDA (Lookups) ---
    const mapaVentaSku = new Map(); // Key: Nombre Producto, Value: { productoBase, cantVenta, unidadVenta }
    const mapaCompraSku = new Map(); // Key: 'Producto Base-Formato Adquisición', Value: cantAdquisicion

    datosSku.forEach(fila => {
      const [nombreProducto, productoBase, formatoAdq, cantAdq, , , cantVenta, unidadVenta] = fila;
      if (nombreProducto) {
        mapaVentaSku.set(nombreProducto, {
          productoBase: productoBase,
          cantVenta: parseFloat(cantVenta) || 0,
          unidadVenta: unidadVenta
        });
      }
      if (productoBase && formatoAdq) {
        const claveCompra = `${productoBase}-${formatoAdq}`;
        mapaCompraSku.set(claveCompra, parseFloat(cantAdq) || 0);
      }
    });

    // --- 3. PROCESAR VENTAS DEL DÍA ---
    const ventasDelDia = {}; // { "Producto Base": cantidad, ... }
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0); // Inicio del día en la zona horaria del script

    datosOrders.forEach(fila => {
      const fechaPedido = new Date(fila[8]); // Columna I: Fecha
      if (fechaPedido >= hoy) {
        const nombreProductoVendido = fila[9]; // Columna J: Nombre Producto
        const cantidadVendida = parseFloat(fila[10]) || 0; // Columna K: Cantidad

        const skuInfo = mapaVentaSku.get(nombreProductoVendido);
        if (skuInfo) {
          const ventaEnUnidadBase = cantidadVendida * skuInfo.cantVenta;
          sumarAObjeto(ventasDelDia, skuInfo.productoBase, ventaEnUnidadBase);
        }
      }
    });

    // --- 4. PROCESAR COMPRAS DEL DÍA ---
    const comprasDelDia = {}; // { "Producto Base": cantidad, ... }
    datosAdquisiciones.forEach(fila => {
      // Asumimos que la fecha relevante está en la columna B (Producto Base) y que es la fecha de adquisición
      // NOTA: La lógica de fecha aquí puede necesitar ajuste según la estructura real de "Adquisiciones".
      // Por ahora, procesaremos todas las adquisiciones como si fueran del día.
      // Para una implementación real, se necesitaría una columna de fecha en la hoja Adquisiciones.
      // Si la columna C es la fecha, sería: const fechaAdq = new Date(fila[2]);
      const productoBase = fila[1]; // Columna B: Producto Base
      const formatoCompra = fila[2]; // Columna C: Formato de Compra
      const cantidadComprada = parseFloat(fila[3]) || 0; // Columna D: Cantidad a Comprar

      const claveCompra = `${productoBase}-${formatoCompra}`;
      const cantAdquisicion = mapaCompraSku.get(claveCompra);

      if (cantAdquisicion) {
        const compraEnUnidadBase = cantidadComprada * cantAdquisicion;
        sumarAObjeto(comprasDelDia, productoBase, compraEnUnidadBase);
      }
    });

    // --- 5. OBTENER INVENTARIO DE AYER ---
    const inventarioAyer = {}; // { "Producto Base": cantidad, ... }
    const productosVistos = new Set();
    // Recorrer el histórico desde el final para encontrar la última entrada de cada producto
    for (let i = datosHistorico.length - 1; i >= 0; i--) {
        const fila = datosHistorico[i];
        const productoBase = fila[1];
        if (!productosVistos.has(productoBase)) {
            const stockReal = parseFloat(fila[3]); // Columna D: Stock Real
            const cantidadCalculada = parseFloat(fila[2]); // Columna C: Cantidad (estimada)

            // Priorizar el stock real si existe y es un número válido
            if (!isNaN(stockReal) && fila[3] !== '') {
                inventarioAyer[productoBase] = stockReal;
            } else {
                inventarioAyer[productoBase] = cantidadCalculada || 0;
            }
            productosVistos.add(productoBase);
        }
    }

    // --- 6. CALCULAR INVENTARIO DE HOY Y PREPARAR REPORTE ---
    const todosLosProductos = new Set([...Object.keys(inventarioAyer), ...Object.keys(comprasDelDia), ...Object.keys(ventasDelDia)]);
    const reporteHoy = [];
    const nuevoHistorico = [];
    const timestamp = new Date();

    todosLosProductos.forEach(producto => {
      const ayer = inventarioAyer[producto] || 0;
      const compras = comprasDelDia[producto] || 0;
      const ventas = ventasDelDia[producto] || 0;
      const hoy = ayer + compras - ventas;

      const skuInfo = mapaVentaSku.get(producto) || {}; // Para obtener unidad de venta

      reporteHoy.push([
        producto,
        ayer,
        compras,
        ventas,
        hoy,
        "" // Stock Real (editable por el usuario)
      ]);

      nuevoHistorico.push([
        timestamp,
        producto,
        hoy,
        "", // Stock Real
        skuInfo.unidadVenta || 'N/A'
      ]);
    });

    // --- 7. ESCRIBIR RESULTADOS EN LAS HOJAS ---
    // Limpiar reporte anterior y escribir el nuevo
    if (hojaReporteHoy.getLastRow() > 1) {
      hojaReporteHoy.getRange(2, 1, hojaReporteHoy.getLastRow() - 1, 6).clearContent();
    }
    if (reporteHoy.length > 0) {
      hojaReporteHoy.getRange(2, 1, reporteHoy.length, 6).setValues(reporteHoy);
    }

    // Añadir al histórico
    if (nuevoHistorico.length > 0) {
      hojaHistorico.getRange(hojaHistorico.getLastRow() + 1, 1, nuevoHistorico.length, 5).setValues(nuevoHistorico);
    }

    ui.showModalDialog(HtmlService.createHtmlOutput('<h3>¡Éxito!</h3><p>El cálculo del inventario ha finalizado.</p>'), 'Proceso Completado');
    // Cerrar el diálogo automáticamente después de unos segundos
    Utilities.sleep(4000);
    const activeDoc = SpreadsheetApp.getActive();
    const html = HtmlService.createHtmlOutput("<script>google.script.host.close()</script>");
    ui.showModalDialog(html, "Cerrando...");


  } catch (e) {
    Logger.log(e);
    ui.showModalDialog(HtmlService.createHtmlOutput(`<h3>Error</h3><p>Ocurrió un error durante el cálculo: ${e.message}</p><pre>${e.stack}</pre>`), 'Error en el Proceso');
  }
}


// =====================================================================================
// FUNCIONES DE PRUEBA Y CONFIGURACIÓN
// =====================================================================================

/**
 * Crea un inventario histórico de prueba para los productos base de la hoja SKU.
 * Borra el histórico existente y genera nuevas entradas con fechas aleatorias
 * en los últimos 7 días y cantidades iniciales aleatorias.
 * Esta función es útil para la configuración inicial y pruebas.
 */
function crearInventarioHistoricoDePrueba() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const respuesta = ui.alert(
    'Confirmación',
    'Esta acción borrará TODOS los datos existentes en la hoja "Inventario Histórico" y los reemplazará con datos de prueba. ¿Desea continuar?',
    ui.ButtonSet.YES_NO
  );

  if (respuesta !== ui.Button.YES) {
    ui.alert('Operación cancelada por el usuario.');
    return;
  }

  try {
    const modalTitle = 'Generando Histórico de Prueba';
    ui.showModalDialog(HtmlService.createHtmlOutput('<h3>Procesando...</h3><p>Generando el inventario histórico de prueba. Por favor, espere.</p>'), modalTitle);

    const hojaSku = ss.getSheetByName(HOJA_SKU);
    const hojaHistorico = ss.getSheetByName(HOJA_HISTORICO);

    if (!hojaSku || !hojaHistorico) {
      throw new Error(`No se encontraron las hojas "${HOJA_SKU}" o "${HOJA_HISTORICO}". Ejecute la configuración "1. Configurar Hojas y Fórmulas" primero.`);
    }

    // 1. Obtener productos base únicos de la hoja SKU
    const datosSku = hojaSku.getRange("B2:H" + hojaSku.getLastRow()).getValues();
    const productosBaseUnicos = new Map(); // Usar un Map para obtener productos únicos y su unidad de venta
    datosSku.forEach(fila => {
      const productoBase = fila[0]; // Columna B
      const unidadVenta = fila[6];   // Columna H
      if (productoBase && productoBase.trim() !== "" && !productosBaseUnicos.has(productoBase)) {
        productosBaseUnicos.set(productoBase, unidadVenta || 'N/A');
      }
    });

    if (productosBaseUnicos.size === 0) {
        throw new Error(`No se encontraron "Producto Base" en la hoja "${HOJA_SKU}". Asegúrese de que los datos estén cargados.`);
    }

    // 2. Preparar datos históricos
    const datosNuevosHistorico = [];
    // Usar la fecha fija del 5 de septiembre de 2025 como "hoy" para consistencia
    const hoy = new Date("2025-09-05T12:00:00");
    const sieteDiasMs = 7 * 24 * 60 * 60 * 1000;

    productosBaseUnicos.forEach((unidad, producto) => {
      // Generar una fecha aleatoria en los últimos 7 días desde "hoy"
      const timestampAleatorio = new Date(hoy.getTime() - Math.random() * sieteDiasMs);

      // Generar una cantidad de stock inicial aleatoria (ej. entre 50 y 200)
      const cantidadInicial = Math.floor(Math.random() * 151) + 50;

      datosNuevosHistorico.push([
        timestampAleatorio,
        producto,
        cantidadInicial,
        '', // Stock Real se deja vacío
        unidad // Unidad de Venta
      ]);
    });

    // Ordenar por fecha para que el histórico tenga sentido cronológico
    datosNuevosHistorico.sort((a, b) => a[0] - b[0]);

    // 3. Escribir en la hoja Histórico
    // Limpiar datos antiguos (manteniendo el encabezado)
    if (hojaHistorico.getLastRow() > 1) {
      hojaHistorico.getRange(2, 1, hojaHistorico.getLastRow() - 1, 5).clearContent();
    }

    // Escribir los nuevos datos
    if (datosNuevosHistorico.length > 0) {
      hojaHistorico.getRange(2, 1, datosNuevosHistorico.length, 5).setValues(datosNuevosHistorico);
    }

    // Cerrar diálogo y mostrar éxito
    const successHtml = HtmlService.createHtmlOutput('<h3>¡Éxito!</h3><p>Se ha generado el inventario histórico de prueba correctamente.</p><p>Cerrando en 3 segundos...</p><script>setTimeout(function(){ google.script.host.close(); }, 3000);</script>')
      .setWidth(400).setHeight(150);
    ui.showModalDialog(successHtml, modalTitle);

  } catch (e) {
    Logger.log(e);
    const errorHtml = HtmlService.createHtmlOutput(`<style>p{font-family:sans-serif;}</style><h3>Error</h3><p>${e.message}</p>`)
      .setWidth(400).setHeight(150);
    ui.showModalDialog(errorHtml, 'Error en el Proceso');
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

// =====================================================================================
// SERVIDOR WEB PARA DASHBOARD
// =====================================================================================

/**
 * Sirve la aplicación web del dashboard.
 * Se ejecuta cuando un usuario visita la URL de la aplicación.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('dashboard')
      .setTitle('Dashboard de Inventario 2.0');
}

/**
 * Lanza el dashboard en un diálogo modal (ventana emergente).
 */
function showDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('dashboard')
      .setWidth(1200)
      .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard de Inventario');
}

/***********************/
/* HELPERS NUMÉRICOS   */
/***********************/
function _toNumber(v) {
  if (typeof v === 'number') return v;
  if (typeof v === 'string') {
    // Soporta comas decimales "0,5"
    const t = v.replace(',', '.').trim();
    const n = parseFloat(t);
    return isNaN(n) ? 0 : n;
  }
  return 0;
}

function _dayString(dateLike) {
  // Devuelve 'yyyy-MM-dd' en America/Santiago a partir de Date o string (p.ej. '2025-09-02 19:53')
  if (dateLike instanceof Date) {
    return Utilities.formatDate(dateLike, TIMEZONE, 'yyyy-MM-dd');
  }
  if (typeof dateLike === 'string') {
    // normaliza 'YYYY-MM-DD HH:mm' -> 'YYYY-MM-DDTHH:mm'
    const normalized = dateLike.replace(' ', 'T');
    const d = new Date(normalized);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, TIMEZONE, 'yyyy-MM-dd');
    }
    // fallback: toma sólo la parte de la fecha antes del espacio
    return (dateLike.split(' ')[0] || '').trim();
  }
  return '';
}

/********************************************/
/* ULTIMO INVENTARIO POR PRODUCTO BASE (B,C) */
/********************************************/
function _mapUltimoInventario(preferirStockReal) {
  // preferirStockReal=true => si hay Stock Real (D) válido, úsalo; si no, usa Cantidad (C)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hist = ss.getSheetByName(HOJA_HISTORICO);
  const mapa = new Map();
  if (!hist || hist.getLastRow() < 2) return mapa;

  const values = hist.getRange(2, 1, hist.getLastRow() - 1, 5).getValues();
  // Estructura: [Timestamp(A), Producto Base(B), Cantidad(C), Stock Real(D), Unidad Venta(E)]
  for (let i = 0; i < values.length; i++) {
    const [ts, base, cant, stockReal, unidad] = values[i];
    if (!base) continue;
    const when = (ts instanceof Date) ? ts : new Date(ts);
    if (isNaN(when.getTime())) continue;

    const qty = (preferirStockReal && stockReal !== '' && !isNaN(_toNumber(stockReal)))
      ? _toNumber(stockReal)
      : _toNumber(cant);

    const prev = mapa.get(base);
    if (!prev || when > prev.ts) {
      mapa.set(base, { ts: when, qty: qty, unit: unidad || '' });
    }
  }
  return mapa;
}

/********************************************/
/* LOOKUPS DE SKU                           */
/********************************************/
function _buildSkuLookups() {
  // Devuelve:
  //  - baseInfoMap: Map(Producto Base -> { unit, category })
  //  - skuByNombre: Map(Nombre Producto -> { productoBase, cantidadVenta, unidadVenta })
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sku = ss.getSheetByName(HOJA_SKU);
  if (!sku || sku.getLastRow() < 2) return { baseInfoMap: new Map(), skuByNombre: new Map() };

  // Tomamos A:H para cubrir: Nombre, Base, Categoria (F), Cantidad Venta(G), Unidad Venta(H)
  const data = sku.getRange(2, 1, sku.getLastRow() - 1, 8).getValues();

  const baseInfoMap = new Map();
  const skuByNombre = new Map();

  data.forEach(r => {
    const nombreProducto = r[0] || '';       // A
    const productoBase   = r[1] || '';       // B
    const categoria      = r[5] || 'Sin Categoría'; // F
    const cantidadVenta  = _toNumber(r[6]);  // G
    const unidadVenta    = r[7] || '';       // H

    if (productoBase && !baseInfoMap.has(productoBase)) {
      baseInfoMap.set(productoBase, { unit: unidadVenta, category: categoria });
    }
    if (nombreProducto) {
      skuByNombre.set(nombreProducto, { productoBase, cantidadVenta, unidadVenta });
    }
  });

  return { baseInfoMap, skuByNombre };
}


/********************************************/
/* REEMPLAZA TU getDashboardData()          */
/********************************************/
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orders = ss.getSheetByName(HOJA_ORDERS);
  if (!orders) throw new Error('No se encontró la hoja Orders.');
  const { baseInfoMap, skuByNombre } = _buildSkuLookups();

  // 1) Último inventario por Producto Base (Inv. Ayer teórico)
  const preferirStockReal = false;
  const lastInvMap = _mapUltimoInventario(preferirStockReal);

  // 2) Ventas de Hoy (por Producto Base)
  const hoy = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
  const rows = orders.getLastRow() > 1 ? orders.getRange(2, 1, orders.getLastRow() - 1, 11).getValues() : [];
  const ventasPorBase = new Map();

  rows.forEach(r => {
    const fecha    = r[8];
    const skuName  = r[9];
    const qtyVenta = _toNumber(r[10]);

    if (!skuName || !qtyVenta) return;
    if (_dayString(fecha) !== hoy) return;

    const sku = skuByNombre.get(skuName);
    if (!sku || !sku.productoBase) return;

    const cantidadBase = qtyVenta * _toNumber(sku.cantidadVenta);
    ventasPorBase.set(sku.productoBase, (ventasPorBase.get(sku.productoBase) || 0) + cantidadBase);
  });

  // 3) Ensamblar el arreglo para el dashboard
  const inventory = [];
  baseInfoMap.forEach(({ unit, category }, baseProduct) => {
    const last = lastInvMap.get(baseProduct);
    const invAyer = last ? _toNumber(last.qty) : 0;
    const sales   = ventasPorBase.get(baseProduct) || 0;

    inventory.push({
      baseProduct: baseProduct,
      category: category,
      lastInventory: invAyer,
      purchases: 0,
      sales: sales,
      expectedStock: invAyer - sales,
      unit: unit || (last ? last.unit : '')
    });
  });

  // Ordenar por categoría, y luego por nombre de producto
  inventory.sort((a, b) => {
    if (a.category < b.category) return -1;
    if (a.category > b.category) return 1;
    return a.baseProduct.localeCompare(b.baseProduct);
  });

  return {
    inventory: inventory,
    sales: [],
    acquisitions: []
  };
}

/***********************************************************/
/* OPCIONAL: “VLOOKUP del último” para una sola clave      */
/***********************************************************/
function lookupUltimoInventarioCantidad(productoBase) {
  const m = _mapUltimoInventario(false); // false: usa SIEMPRE Cantidad(C)
  return m.get(productoBase)?.qty || 0;
}

/**
 * Guarda el inventario real introducido por el usuario.
 * @param {Array<Object>} inventoryData Un array de objetos, cada uno con {productoBase, cantidad}.
 */
function saveRealInventory(inventoryData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaReporte = ss.getSheetByName(HOJA_REPORTE_HOY);
    const hojaDiscrepancias = ss.getSheetByName("Discrepancias");
    const hojaHistorico = ss.getSheetByName(HOJA_HISTORICO);

    const datosReporte = hojaReporte.getRange("A2:F" + hojaReporte.getLastRow()).getValues();
    const reporteMap = new Map(datosReporte.map(row => [row[0], row])); // Map by Producto Base

    const datosHistorico = hojaHistorico.getRange("A2:E" + hojaHistorico.getLastRow()).getValues();

    const discrepanciasNuevas = [];
    const timestamp = new Date();

    inventoryData.forEach(item => {
      const productoBase = item.productoBase;
      const cantidadReal = parseFloat(item.cantidad);

      if (!isNaN(cantidadReal) && reporteMap.has(productoBase)) {
        const filaReporte = reporteMap.get(productoBase);
        const inventarioEstimado = parseFloat(filaReporte[4]); // Columna E: Inventario Hoy
        const discrepancia = cantidadReal - inventarioEstimado;

        // 1. Log en Hoja Discrepancias
        discrepanciasNuevas.push([
          timestamp,
          productoBase,
          inventarioEstimado,
          cantidadReal,
          discrepancia
        ]);

        // 2. Actualizar Stock Real en Hoja Reporte Hoy
        // Encontrar la fila correcta y actualizar solo la columna F (Stock Real)
        for (let i = 0; i < datosReporte.length; i++) {
          if (datosReporte[i][0] === productoBase) {
            hojaReporte.getRange(i + 2, 6).setValue(cantidadReal); // Fila i+2, Columna 6
            break;
          }
        }

        // 3. Actualizar Stock Real en la última entrada del Histórico para ese producto
        // Esto es crucial para que el cálculo del día siguiente sea correcto.
        let ultimaFilaProducto = -1;
        for (let i = datosHistorico.length - 1; i >= 0; i--) {
          if (datosHistorico[i][1] === productoBase) {
             ultimaFilaProducto = i;
             break;
          }
        }
        if(ultimaFilaProducto !== -1) {
            // La columna D (4) es 'Stock Real' en Histórico
            hojaHistorico.getRange(ultimaFilaProducto + 2, 4).setValue(cantidadReal);
        }
      }
    });

    if (discrepanciasNuevas.length > 0) {
      hojaDiscrepancias.getRange(hojaDiscrepancias.getLastRow() + 1, 1, discrepanciasNuevas.length, 5).setValues(discrepanciasNuevas);
    }

    // Forzar un recálculo para que el "Inventario Ayer" del próximo ciclo sea el real
    calcularInventarioDiario();

    return { success: true, message: "Inventario real guardado correctamente." };

  } catch (e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}
