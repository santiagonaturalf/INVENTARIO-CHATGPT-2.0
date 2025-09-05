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

/**
 * Devuelve todos los productos base registrados en la hoja SKU para mostrarlos en el dashboard,
 * aunque todavía no tengan datos de inventario.
 *
 * Estructura de la hoja SKU (columnas relevantes):
 *   A: Nombre Producto
 *   B: Producto Base
 *   C: Formato Adquisición
 *   D: Cantidad Adquisición
 *   E: Unidad Adquisición
 *   F: Categoría
 *   G: Cantidad Venta
 *   H: Unidad Venta
 *
 * @returns {Object} Un objeto con:
 *   - inventory: arreglo con cada producto base y valores iniciales en cero
 *   - sales: []
 *   - acquisitions: []
 */
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const skuSheet = ss.getSheetByName(HOJA_SKU);
  if (!skuSheet) {
    throw new Error('No se encontró la hoja "SKU". Verifica el nombre de la hoja.');
  }

  const lastRow = skuSheet.getLastRow();
  if (lastRow < 2) {
    return { inventory: [], sales: [], acquisitions: [] };
  }

  // Leer columnas de la hoja SKU: B (Producto Base) y H (Unidad Venta)
  const data = skuSheet.getRange(2, 2, lastRow - 1, 7).getValues();
  // data = matriz de filas, cada fila = [Producto Base, Formato Adq, Cantidad Adq, Unidad Adq, Categoría, Cantidad Venta, Unidad Venta]

  // Construir un map único por producto base para no duplicar
  const uniqueBases = new Map();
  data.forEach(row => {
    const productoBase = row[0];
    const unidadVenta = row[6] || ''; // Columna H (índice 6) es Unidad Venta
    if (productoBase && !uniqueBases.has(productoBase)) {
      uniqueBases.set(productoBase, unidadVenta);
    }
  });

  // Crear el arreglo de inventario con valores iniciales en 0
  const inventory = [];
  uniqueBases.forEach((unidad, producto) => {
    inventory.push({
      baseProduct: producto,
      lastInventory: 0,
      purchases: 0,
      sales: 0,
      expectedStock: 0,
      unit: unidad
    });
  });

  return {
    inventory: inventory,
    sales: [],
    acquisitions: []
  };
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
