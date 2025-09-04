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

// Datos de inventario de prueba
// Cada objeto representa: producto base, unidad y cantidad (inventario hoy)
const INVENTARIO_PRUEBA = [
  { producto: 'frutilla', unidad: 'Kilo', cantidad: 2 },
  { producto: 'kiwi', unidad: 'Kilo', cantidad: 16 },
  { producto: 'limon', unidad: 'Kilo', cantidad: 40 },
  { producto: 'limon sutil', unidad: 'Kilo', cantidad: 17 },
  { producto: 'mango', unidad: 'Kilo', cantidad: 1 },
  { producto: 'Zapallo Cubo PREELABORADO', unidad: 'Envase', cantidad: 2 },
  { producto: 'Carbonada PREELABORADO', unidad: 'kilo', cantidad: 2 }
];

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
      .addItem('2. Generar Inventario Histórico Automático', 'generarInventarioHistoricoAutomatico')
      .addSeparator()
      .addItem('3. Calcular Inventario de Hoy (Manual)', 'calcularInventarioDiario')
      .addSeparator()
      .addItem('4. Activar/Actualizar Trigger Diario', 'crearDisparadorDiario')
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

/**
 * Genera entradas de inventario histórico basadas en INVENTARIO_PRUEBA.
 * Asigna fechas consecutivas hacia atrás desde hoy para cada producto.
 */
function generarInventarioHistoricoAutomatico() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaHistorico = ss.getSheetByName(HOJA_HISTORICO); // Reutiliza tu constante HOJA_HISTORICO
  if (!hojaHistorico) {
    SpreadsheetApp.getUi().alert('No existe la hoja "' + HOJA_HISTORICO + '".');
    return;
  }

  // Calcular una fecha distinta para cada producto, restando i días al día actual
  const baseDate = new Date();
  baseDate.setHours(0, 0, 0, 0); // Hora cero
  const timezone = TIMEZONE || Session.getScriptTimeZone();

  const filas = INVENTARIO_PRUEBA.map((item, index) => {
    // Restar index días a la fecha base
    const fecha = new Date(baseDate);
    fecha.setDate(fecha.getDate() - index);
    const timestamp = Utilities.formatDate(fecha, timezone, 'yyyy-MM-dd\'T\'HH:mm:ss');

    return [
      timestamp,           // Timestamp
      item.producto,       // Producto Base
      item.cantidad,       // Cantidad inventariada (Stock Real)
      '',                  // Stock Real (vacío por ahora)
      item.unidad          // Unidad de venta
    ];
  });

  // Insertar las filas en la hoja, a partir de la primera fila libre
  const startRow = hojaHistorico.getLastRow() + 1;
  hojaHistorico.getRange(startRow, 1, filas.length, 5).setValues(filas);

  SpreadsheetApp.getUi().alert('Inventario histórico de prueba cargado correctamente.');
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
 * Lanza el dashboard en una barra lateral en la hoja de cálculo.
 */
function showDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('dashboard')
      .setTitle('Dashboard de Inventario');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Obtiene todos los datos necesarios para renderizar el dashboard.
 * Esta función es llamada desde el cliente (JavaScript en dashboard.html).
 * @returns {Object} Un objeto con los datos de inventario, ventas, adquisiciones y KPIs.
 */
function getDashboardData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaReporte = ss.getSheetByName(HOJA_REPORTE_HOY);
    const hojaOrders = ss.getSheetByName(HOJA_ORDERS);
    const hojaAdquisiciones = ss.getSheetByName(HOJA_ADQUISICIONES);
    const hojaSku = ss.getSheetByName(HOJA_SKU);

    // --- 1. OBTENER MAPA SKU PARA UNIDADES Y CONVERSIONES ---
    const datosSku = hojaSku.getRange("A2:K" + hojaSku.getLastRow()).getValues();
    const mapaVentaSku = new Map();
    datosSku.forEach(fila => {
      const [nombreProducto, productoBase, , , , , , unidadVenta] = fila;
      if (nombreProducto && productoBase) {
        mapaVentaSku.set(nombreProducto, { productoBase, unidadVenta });
      }
    });

    const mapaUnidades = new Map();
     datosSku.forEach(fila => {
      const [, productoBase, , , , , , unidadVenta] = fila;
      if (productoBase && unidadVenta && !mapaUnidades.has(productoBase)) {
        mapaUnidades.set(productoBase, unidadVenta);
      }
    });


    // --- 2. OBTENER INVENTARIO ACTUAL ---
    const datosInventario = hojaReporte.getLastRow() > 1 ? hojaReporte.getRange("A2:E" + hojaReporte.getLastRow()).getValues() : [];
    const inventory = datosInventario.map(fila => ({
      productoBase: fila[0],
      inventarioAyer: fila[1],
      compras: fila[2],
      ventas: fila[3],
      inventarioHoy: fila[4],
      unidad: mapaUnidades.get(fila[0]) || 'N/A'
    }));

    // --- 3. OBTENER VENTAS DE HOY ---
    const datosOrders = hojaOrders.getLastRow() > 1 ? hojaOrders.getRange("A2:K" + hojaOrders.getLastRow()).getValues() : [];
    const sales = [];
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);

    datosOrders.forEach(fila => {
      const fechaPedido = new Date(fila[8]); // Columna I: Fecha
      if (fechaPedido >= hoy) {
        const nombreProductoVendido = fila[9]; // Columna J: Nombre Producto
        const skuInfo = mapaVentaSku.get(nombreProductoVendido);
        sales.push({
          pedido: fila[0], // Columna A: Order Number
          cliente: fila[2], // Columna C: Billing Name
          productoVendido: nombreProductoVendido,
          productoBase: skuInfo ? skuInfo.productoBase : 'No Encontrado',
          cantidad: fila[10] // Columna K: Cantidad
        });
      }
    });

    // --- 4. OBTENER ADQUISICIONES DE HOY ---
    // Nota: Asumimos que todas las adquisiciones listadas son para el día.
    const datosAdquisiciones = hojaAdquisiciones.getLastRow() > 1 ? hojaAdquisiciones.getRange("A2:D" + hojaAdquisiciones.getLastRow()).getValues() : [];
    const acquisitions = datosAdquisiciones.map(fila => ({
      productoBase: fila[1], // Columna B: Producto Base
      formato: fila[2],      // Columna C: Formato de Compra
      cantidad: fila[3]      // Columna D: Cantidad a Comprar
    }));

    // --- 5. CALCULAR KPIs ---
    const kpis = {
      totalProducts: inventory.length,
      lowStock: inventory.filter(p => p.inventarioHoy <= 0).length,
      discrepancies: 0 // Se implementará en el futuro
    };

    return {
      inventory: inventory,
      sales: sales.reverse(), // Mostrar las más recientes primero
      acquisitions: acquisitions,
      kpis: kpis
    };

  } catch (e) {
    Logger.log(e);
    // Devolver un error estructurado al cliente
    return { error: e.message, stack: e.stack };
  }
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
