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
      .addItem('Abrir Dashboard (Antiguo)', 'showDashboard')
      .addItem('Abrir Dashboard v3', 'showDashboardV3')
      .addSeparator()
      .addItem('1. Configurar Hojas y Fórmulas', 'setup')
      .addItem('Enriquecer Datos de Orders', 'completarSKUenOrders')
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
  const encabezadosReporte = ["Producto Base", "Inventario Ayer", "Compras del Día", "Ventas del Día", "Inventario Hoy", "Stock Real", "Discrepancias"];
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
      hojaReporteHoy.getRange(2, 1, hojaReporteHoy.getLastRow() - 1, 7).clearContent();
    }
    if (reporteHoy.length > 0) {
      hojaReporteHoy.getRange(2, 1, reporteHoy.length, 6).setValues(reporteHoy);
      // Añadir fórmula de discrepancia en la columna G
      const formulaRange = hojaReporteHoy.getRange(2, 7, reporteHoy.length);
      formulaRange.setFormulaR1C1('=IF(RC[-1]<>"", RC[-1]-RC[-2], "")');
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
 * Lanza el dashboard en un diálogo modal (ventana emergente).
 */
function showDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('dashboard.html')
      .setWidth(1200)
      .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard de Inventario v2');
}

/**
 * Lanza el dashboard v3 en un diálogo modal (ventana emergente).
 */
function showDashboardV3() {
  const html = HtmlService.createHtmlOutputFromFile('dashboard_v3.html')
      .setWidth(1200)
      .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard de Inventario v3');
}


// =========================
// Helpers de normalización
// =========================
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
    if (base) map.set(base, (estado || 'pendiente').toString().toLowerCase());
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
      if ((data[i][0]||'').toString().trim().toLowerCase() === baseProduct.toString().trim().toLowerCase()){
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
    baseProduct,
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
  },
  SKU_HEADERS: {
    nombreProducto: 'Nombre Producto',
    productoBase: 'Producto Base',
    cantVenta: 'Cantidad Venta',
    unidadVenta: 'Unidad Venta',
  }
};

/***** HELPERS *****/
const norm = s => (s ?? '').toString().trim().toLowerCase();

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
      base: base.toString().trim(),
      factor: Number(fac) || 0,
      unidad: uni
    };
  }

  if (warnings.length) Logger.log(warnings.join('\n'));
  return map;
}

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

    const nombre = norm(row[H.nombreProducto]);
    const cantidad = Number(row[H.cantidad]) || 0;
    if (!nombre || cantidad <= 0) continue;

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
    const baseKey = norm(s.base);
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
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const skuSheet       = ss.getSheetByName(HOJA_SKU);
  const ordersSheet    = ss.getSheetByName(HOJA_ORDERS);
  const adqSheet       = ss.getSheetByName(HOJA_ADQUISICIONES);
  const histSheet      = ss.getSheetByName(HOJA_HISTORICO);

  if (!skuSheet || !ordersSheet) {
    throw new Error('No se encontraron las hojas SKU u Orders.');
  }

  // === 1) Mapa SKU: Nombre Producto -> {productoBase, cantidadVenta, unidadVenta}
  const skuData = skuSheet.getRange(2, 1, Math.max(0, skuSheet.getLastRow()-1), 8).getValues();
  const skuMap = new Map();
  const baseUnitMap = new Map();
  skuData.forEach(r => {
    const nombreProd = r[0];
    const productoBase = r[1];
    const cantidadVenta = parseFloat((r[6] || '0').toString().replace(',', '.')) || 0;
    const unidadVenta   = r[7] || '';
    if (nombreProd) skuMap.set(nombreProd, { productoBase, cantidadVenta, unidadVenta });
    if (productoBase && !baseUnitMap.has(productoBase)) baseUnitMap.set(productoBase, unidadVenta);
  });

  // === 2) Último inventario por Producto Base (Inv. Ayer)
  const lastInvMap = new Map();
  if (histSheet && histSheet.getLastRow() > 1) {
    const histData = histSheet.getRange(2, 1, histSheet.getLastRow() - 1, 4).getValues();
    histData.forEach(r => {
      const ts = r[0];
      const base = r[1];
      const qty = parseFloat((r[2] || '0').toString().replace(',', '.'));
      const realQty = parseFloat((r[3] || '').toString().replace(',', '.'));

      const when = (ts instanceof Date) ? ts : new Date(ts);
      if (!base || isNaN(when.getTime())) return;

      const currentQty = !isNaN(realQty) ? realQty : qty;
      const prev = lastInvMap.get(base);
      if (!prev || when > prev.ts) {
        lastInvMap.set(base, { ts: when, qty: currentQty });
      }
    });
  }

  // === 3) Ventas del día (lógica simple SUMAR.SI)
  const ventasPorBase = new Map();
  const ordersData = ordersSheet.getDataRange().getValues();
  const ordersHeaders = ordersData[0].map(h => (''+h).trim());

  const idxProductoBase = ordersHeaders.indexOf('Producto Base'); // Col L
  const idxCantidadVenta = ordersHeaders.indexOf('Cantidad (venta)'); // Col M

  if (idxProductoBase !== -1 && idxCantidadVenta !== -1) {
    for (let i = 1; i < ordersData.length; i++) {
      const row = ordersData[i];
      const productoBase = row[idxProductoBase];
      const cantidad = parseFloat(row[idxCantidadVenta]) || 0;

      if (productoBase && cantidad > 0) {
        const key = (''+productoBase).trim();
        ventasPorBase.set(key, (ventasPorBase.get(key) || 0) + cantidad);
      }
    }
  } else {
      Logger.log("No se encontraron las columnas 'Producto Base' (L) o 'Cantidad (venta)' (M) en la hoja 'Orders'. Las ventas aparecerán en 0. Ejecuta 'Enriquecer Datos de Orders' primero.");
  }

  // === 4) Compras hoy (calculo F + H − E por Producto Base)
  const comprasPorBase = getComprasPorBase_SUMARSI_(adqSheet);

  // === 5) Armar inventory[] para el Dashboard
  const inventory = [];
  baseUnitMap.forEach((unidad, base) => {
    const lastInv = lastInvMap.get(base);
    const lastInventory = lastInv ? lastInv.qty : 0;
    const sales    = ventasPorBase.get(base)   || 0;
    const purchases= comprasPorBase.get(base)  || 0;

    inventory.push({
      baseProduct:  base,
      lastInventory: lastInventory,
      purchases:    purchases,
      sales:        sales,
      expectedStock: lastInventory + purchases - sales,
      unit:         unidad,
      error:        false,
      errorMsg:     ""
    });
  });

  const estados = getEstadosParaUI(); // <-- NUEVO
  return {
    inventory: inventory,
    sales: [],
    acquisitions: [],
    estados: estados // <-- NUEVO
  };
}


/**
 * Guarda el stock real para un único producto.
 * @param {string} productBase El producto base a actualizar.
 * @param {number} quantity La cantidad de stock real.
 */
function saveSingleRealStock(productBase, quantity) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaReporte = ss.getSheetByName(HOJA_REPORTE_HOY);
    const hojaDiscrepancias = ss.getSheetByName("Discrepancias");
    const hojaHistorico = ss.getSheetByName(HOJA_HISTORICO);

    if (!hojaReporte || hojaReporte.getLastRow() < 2) {
      throw new Error(`La hoja "${HOJA_REPORTE_HOY}" no está lista o está vacía.`);
    }

    const data = hojaReporte.getDataRange().getValues();
    const headers = data[0];
    const baseProductCol = headers.indexOf("Producto Base");
    const invHoyCol = headers.indexOf("Inventario Hoy");
    const stockRealCol = headers.indexOf("Stock Real");

    if (baseProductCol === -1 || invHoyCol === -1 || stockRealCol === -1) {
      throw new Error("No se encontraron las columnas 'Producto Base', 'Inventario Hoy' o 'Stock Real' en 'Reporte Hoy'.");
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][baseProductCol] === productBase) {
        const rowNum = i + 1;

        hojaReporte.getRange(rowNum, stockRealCol + 1).setValue(quantity);

        const inventarioEstimado = parseFloat(data[i][invHoyCol]);
        const discrepancia = quantity - inventarioEstimado;
        if (hojaDiscrepancias) {
          hojaDiscrepancias.appendRow([new Date(), productBase, inventarioEstimado, quantity, discrepancia]);
        }

        if (hojaHistorico && hojaHistorico.getLastRow() > 1) {
            const histData = hojaHistorico.getDataRange().getValues();
            let ultimaFilaProducto = -1;
            for (let j = histData.length - 1; j >= 1; j--) {
                if (histData[j][1] === productBase) {
                    ultimaFilaProducto = j + 1;
                    break;
                }
            }
            if (ultimaFilaProducto !== -1) {
                hojaHistorico.getRange(ultimaFilaProducto, 4).setValue(quantity);
            }
        }

        return { success: true, message: `Stock para ${productBase} actualizado.` };
      }
    }

    return { success: false, error: `Producto ${productBase} no encontrado en 'Reporte Hoy'.` };

  } catch (e) {
    Logger.log(e.stack);
    return { success: false, error: e.message };
  }
}

/**
 * Aprueba un producto y guarda su stock real en una sola operación.
 * @param {string} productBase El producto base a actualizar.
 * @param {number} quantity La cantidad de stock real.
 */
function approveProductAndSaveStock(productBase, quantity) {
  try {
    setEstadoProducto(productBase, 'aprobado', 'Aprobado desde dashboard');
    const saveResult = saveSingleRealStock(productBase, quantity);
    if (!saveResult.success) {
      throw new Error(saveResult.error);
    }
    return { success: true, message: `Producto ${productBase} aprobado y stock actualizado.` };
  } catch (e) {
    Logger.log(e.stack);
    return { success: false, error: e.message };
  }
}
