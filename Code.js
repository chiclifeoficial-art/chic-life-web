/**
 * CHIC LIFE — Captura de clientas + reseñas de productos
 * ─────────────────────────────────────────────────────
 * Este script convierte una Google Sheet en tu "base de datos" de
 * clientas y reseñas, y en la API que consume el sitio web.
 *
 * Hojas que administra:
 *   - "Clientas": popup de bienvenida / código 10% OFF (la crea sola).
 *   - "Reseñas":  calificaciones y comentarios por producto (la crea sola).
 *   - "Apeiron - Inventario": catálogo de productos. La sincroniza
 *     apeiron-erp/backend/services/googleSyncService.js cada 15 min
 *     (bidireccional Apeiron ↔ Sheets, con CLEAR + reescritura completa
 *     de la hoja en cada corrida). Este script SOLO LEE de ahí, nunca
 *     escribe — cero riesgo de chocar con esa sincronización. Precio
 *     Venta viene incluido en ese export, así que se lee directo de ahí
 *     (no se duplica en ninguna otra hoja).
 *
 * IMPORTANTE — aislamiento de "Apeiron - Inventario":
 *   Este script identifica la hoja de cálculo por ID fijo (SPREADSHEET_ID,
 *   vía abrirLibro_()), no por "hoja activa". Eso permite instalarlo como
 *   proyecto de Apps Script independiente (no necesita ser el script
 *   contenedor de esta Sheet), así que puede coexistir sin conflicto con
 *   cualquier integración que Apeiron ERP tenga con la hoja.
 *
 * INSTALACIÓN (ver guía GUIA_DESPLIEGUE.md paso a paso):
 * 1. Ve a script.google.com > Proyecto nuevo (o Extensiones > Apps Script
 *    desde la Sheet, si prefieres que quede ligado a ella).
 * 2. Borra el contenido de Code.gs y pega TODO este archivo.
 * 3. Implementar > Nueva implementación > Aplicación web.
 *    - Ejecutar como: Yo (tu cuenta)
 *    - Quién tiene acceso: Cualquier usuario
 * 4. Copia la URL que te da ("Web App URL") y pégala en
 *    APPS_SCRIPT_URL dentro del index.html (popup Y reseñas).
 *
 * Cada vez que edites este archivo, tienes que volver a publicar:
 * Implementar > Gestionar implementaciones > lápiz > Nueva versión > Implementar.
 */

// ID fijo de la Google Sheet — el script SIEMPRE apunta aquí, sin
// importar si queda instalado como script contenedor o independiente.
const SPREADSHEET_ID = '1-OWb-eIYu2UNqP4_IT1hNRv5CZfjZcJ0EWs05lkD7xU';

function abrirLibro_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

const SHEET_CLIENTAS = 'Clientas';
const SHEET_RESENAS = 'Reseñas';
const SHEET_INVENTARIO = 'Apeiron - Inventario';
const SHEET_CARRITOS = 'Carritos';
const SHEET_PEDIDOS = 'Pedidos';
const CODE_PREFIX = 'BIENVENIDA-';

// ═══════════════════════════════════════════════════════════
// ENRUTAMIENTO
// ═══════════════════════════════════════════════════════════

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const data = JSON.parse(e.postData.contents);

    if (data.accion === 'resena') {
      return crearResena_(data);
    }
    if (data.accion === 'carrito') {
      return registrarCarrito_(data);
    }
    if (data.accion === 'pedido') {
      return registrarPedido_(data);
    }
    if (data.accion === 'actualizar-pedido') {
      return actualizarEstadoPedido_(data);
    }
    return crearSuscripcion_(data);

  } catch (err) {
    return jsonResponse_({ success: false, error: err.message });
  } finally {
    lock.releaseLock();
  }
}

// Permite abrir la URL en el navegador para confirmar que el
// despliegue quedó activo, y sirve los datos de reseñas (lectura pública).
function doGet(e) {
  const params = (e && e.parameter) || {};

  try {
    if (params.accion === 'resenas-resumen') {
      return jsonResponse_({ success: true, resumen: resumenPorProducto_() });
    }
    if (params.accion === 'resenas') {
      const productoId = limpiar_(params.producto);
      if (!productoId) {
        return jsonResponse_({ success: false, error: 'Falta el parámetro producto.' });
      }
      return jsonResponse_(obtenerResenasProducto_(productoId, params.orden));
    }
    if (params.accion === 'inventario') {
      return jsonResponse_({ success: true, productos: leerInventario_() });
    }
    if (params.accion === 'pedidos') {
      return jsonResponse_({ success: true, pedidos: leerPedidos_() });
    }
    return ContentService
      .createTextOutput('Chic Life — API de clientas y reseñas activa ✅')
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return jsonResponse_({ success: false, error: err.message });
  }
}

// ═══════════════════════════════════════════════════════════
// SUSCRIPCIÓN AL POPUP (10% OFF)
// ═══════════════════════════════════════════════════════════

function crearSuscripcion_(data) {
  const sheet = getOrCreateSheetClientas_();

  const nombre = limpiar_(data.nombre);
  const email = limpiar_(data.email).toLowerCase();
  const whatsapp = limpiar_(data.whatsapp);
  const pais = limpiar_(data.pais);
  const fechaNacimiento = limpiar_(data.fechaNacimiento);

  if (!nombre || !email || !isEmailValido_(email)) {
    return jsonResponse_({ success: false, error: 'Nombre y correo válido son obligatorios.' });
  }

  // ¿Ya existe esta clienta? Si sí, le reenviamos su mismo código
  // en vez de crear un duplicado o dejarla pedir uno nuevo.
  const existente = buscarClientaPorEmail_(sheet, email);
  if (existente) {
    return jsonResponse_({ success: true, code: existente.code, duplicate: true });
  }

  const code = generarCodigoUnico_(sheet);

  // Google Sheets interpreta un texto que empieza con "+", "-" o "="
  // como el inicio de una fórmula (por eso salía #ERROR! en WhatsApp).
  // Forzamos la columna E (WhatsApp) a texto plano antes de escribir,
  // como defensa aunque el valor ya venga sin "+" desde el popup.
  sheet.getRange('E2:E').setNumberFormat('@');

  sheet.appendRow([
    new Date(),      // Fecha de registro
    nombre,          // Nombre
    email,           // Correo
    pais,            // País (bandera / indicativo)
    whatsapp,        // WhatsApp
    fechaNacimiento, // Fecha de nacimiento
    code,            // Código único asignado
    'NO'             // ¿Usado? (lo cambias a SI manualmente al validar la compra)
  ]);

  return jsonResponse_({ success: true, code: code, duplicate: false });
}

function getOrCreateSheetClientas_() {
  const ss = abrirLibro_();
  let sheet = ss.getSheetByName(SHEET_CLIENTAS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CLIENTAS);
    sheet.appendRow(['Fecha registro', 'Nombre', 'Correo', 'País', 'WhatsApp', 'Fecha nacimiento', 'Código', 'Usado']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function buscarClientaPorEmail_(sheet, email) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][2]).toLowerCase() === email) {
      return { row: i + 2, code: values[i][6] };
    }
  }
  return null;
}

function generarCodigoUnico_(sheet) {
  const lastRow = sheet.getLastRow();
  const existentes = lastRow < 2 ? [] : sheet.getRange(2, 7, lastRow - 1, 1).getValues().flat();
  let code;
  do {
    const sufijo = Math.random().toString(36).substring(2, 7).toUpperCase();
    code = CODE_PREFIX + sufijo;
  } while (existentes.indexOf(code) !== -1);
  return code;
}

// ═══════════════════════════════════════════════════════════
// RESEÑAS DE PRODUCTOS
// Columnas: Fecha | ProductoId | Producto | Calificación | Comentario
//           | Nombre | Correo | Estado (Pendiente / Aprobada / Rechazada)
//
// Las reseñas SIEMPRE entran como "Pendiente". Solo cuentan para el
// promedio y aparecen en el sitio cuando cambias esa celda a "Aprobada"
// manualmente en la hoja — es tu filtro contra spam o comentarios falsos.
// ═══════════════════════════════════════════════════════════

const COL_RESENA = {
  FECHA: 0, PRODUCTO_ID: 1, PRODUCTO: 2, CALIFICACION: 3,
  COMENTARIO: 4, NOMBRE: 5, CORREO: 6, ESTADO: 7
};

function crearResena_(data) {
  const sheet = getOrCreateSheetResenas_();

  const productoId = limpiar_(data.productoId);
  const producto = limpiar_(data.productoNombre);
  const calificacion = Number(data.calificacion);
  const comentario = limpiar_(data.comentario);
  const nombre = limpiar_(data.nombre);
  const email = limpiar_(data.email).toLowerCase();

  if (!productoId || !producto) {
    return jsonResponse_({ success: false, error: 'Falta identificar el producto.' });
  }
  if (!calificacion || calificacion < 1 || calificacion > 5) {
    return jsonResponse_({ success: false, error: 'Selecciona una calificación de 1 a 5 estrellas.' });
  }
  if (!comentario) {
    return jsonResponse_({ success: false, error: 'Escribe un comentario.' });
  }
  if (comentario.length > 2000) {
    return jsonResponse_({ success: false, error: 'El comentario es demasiado largo (máx. 2000 caracteres).' });
  }
  if (!nombre || !email || !isEmailValido_(email)) {
    return jsonResponse_({ success: false, error: 'Nombre y correo válido son obligatorios.' });
  }

  sheet.appendRow([
    new Date(),   // Fecha
    productoId,   // ProductoId (slug interno)
    producto,     // Producto (nombre legible)
    calificacion, // Calificación 1-5
    comentario,   // Comentario
    nombre,       // Nombre
    email,        // Correo
    'Pendiente'   // Estado — cámbialo a "Aprobada" para publicarla
  ]);

  return jsonResponse_({
    success: true,
    mensaje: '¡Gracias por tu reseña! Quedará publicada en el sitio en cuanto la revisemos.'
  });
}

function getOrCreateSheetResenas_() {
  const ss = abrirLibro_();
  let sheet = ss.getSheetByName(SHEET_RESENAS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_RESENAS);
    sheet.appendRow(['Fecha', 'ProductoId', 'Producto', 'Calificación', 'Comentario', 'Nombre', 'Correo', 'Estado']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function leerResenasAprobadas_() {
  const sheet = getOrCreateSheetResenas_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  return values
    .filter(row => String(row[COL_RESENA.ESTADO]).trim().toLowerCase() === 'aprobada')
    .map(row => ({
      fecha: row[COL_RESENA.FECHA],
      productoId: String(row[COL_RESENA.PRODUCTO_ID]),
      producto: String(row[COL_RESENA.PRODUCTO]),
      calificacion: Number(row[COL_RESENA.CALIFICACION]),
      comentario: String(row[COL_RESENA.COMENTARIO]),
      nombre: String(row[COL_RESENA.NOMBRE])
    }));
}

// Promedio y total por producto — usado para pintar el badge de
// estrellas en cada tarjeta del catálogo sin tener que pedir un
// endpoint distinto por cada uno de los 18 productos.
function resumenPorProducto_() {
  const aprobadas = leerResenasAprobadas_();
  const resumen = {};
  aprobadas.forEach(r => {
    if (!resumen[r.productoId]) {
      resumen[r.productoId] = { total: 0, suma: 0 };
    }
    resumen[r.productoId].total += 1;
    resumen[r.productoId].suma += r.calificacion;
  });
  Object.keys(resumen).forEach(id => {
    const item = resumen[id];
    item.promedio = Math.round((item.suma / item.total) * 10) / 10;
    delete item.suma;
  });
  return resumen;
}

function obtenerResenasProducto_(productoId, orden) {
  const todas = leerResenasAprobadas_().filter(r => r.productoId === productoId);

  const distribucion = { 5: 0, 4: 0, 3: 0, 2: 0, 1: 0 };
  let suma = 0;
  todas.forEach(r => {
    const c = Math.min(5, Math.max(1, Math.round(r.calificacion)));
    distribucion[c] += 1;
    suma += r.calificacion;
  });
  const total = todas.length;
  const promedio = total ? Math.round((suma / total) * 10) / 10 : 0;

  let ordenadas = todas.slice();
  if (orden === 'calif-alta') {
    ordenadas.sort((a, b) => b.calificacion - a.calificacion);
  } else if (orden === 'calif-baja') {
    ordenadas.sort((a, b) => a.calificacion - b.calificacion);
  } else {
    ordenadas.sort((a, b) => new Date(b.fecha) - new Date(a.fecha)); // más nuevo primero (default)
  }

  return {
    success: true,
    stats: { promedio: promedio, total: total, distribucion: distribucion },
    resenas: ordenadas.map(r => ({
      nombre: r.nombre,
      calificacion: r.calificacion,
      comentario: r.comentario,
      fecha: Utilities.formatDate(new Date(r.fecha), Session.getScriptTimeZone(), 'dd/MM/yyyy')
    }))
  };
}

// ═══════════════════════════════════════════════════════════
// CARRITOS (para ver carritos abandonados)
// Hoja "Carritos" — propia de este script, sin relación con
// "Apeiron - Inventario" (no hay ningún riesgo de choque con la
// sincronización de Apeiron: es una hoja más, igual que Clientas
// y Reseñas).
//
// El sitio manda una fila cada vez que:
//   a) la clienta cierra o cambia de pestaña con el carrito lleno
//      (estado "Abandonado" — no hay garantía de que realmente lo
//      haya abandonado, solo que se fue sin dar clic en "Finalizar
//      compra"), o
//   b) da clic en "Finalizar compra por WhatsApp" (estado "Enviado
//      a WhatsApp" — tampoco es garantía de compra, solo de que
//      llegó a ese paso).
// No hay forma de saber con certeza cuál se convirtió en venta real
// sin cruzarlo tú manualmente contra tus pedidos — esta hoja es una
// señal de intención, no un reporte de ventas.
// ═══════════════════════════════════════════════════════════

function registrarCarrito_(data) {
  const items = Array.isArray(data.productos) ? data.productos : [];
  if (!items.length) {
    return jsonResponse_({ success: false, error: 'Carrito vacío.' });
  }

  const sheet = getOrCreateSheetCarritos_();

  const resumen = items
    .map(it => `${limpiar_(it.nombre)} x${Number(it.cantidad) || 0}`)
    .join(', ');
  const total = Number(data.total) || 0;
  const sesion = limpiar_(data.sesionId);
  const estado = limpiar_(data.estado) || 'Abandonado';

  sheet.appendRow([
    new Date(),  // Fecha
    sesion,      // Sesión (id anónimo del navegador, no identifica a la persona)
    resumen,     // Productos
    total,       // Total
    estado       // Abandonado / Enviado a WhatsApp
  ]);

  return jsonResponse_({ success: true });
}

function getOrCreateSheetCarritos_() {
  const ss = abrirLibro_();
  let sheet = ss.getSheetByName(SHEET_CARRITOS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CARRITOS);
    sheet.appendRow(['Fecha', 'Sesión', 'Productos', 'Total', 'Estado']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ═══════════════════════════════════════════════════════════
// INVENTARIO DE PRODUCTOS (Apeiron ERP → Google Sheets → sitio web)
//
// El script SOLO LEE la hoja "Apeiron - Inventario" — ninguna función de
// este archivo la modifica (no hay .setValue/.appendRow/.clear sobre
// SHEET_INVENTARIO en ningún lugar). Esto es lo que garantiza que la
// comunicación Google Sheets → sitio web sea estrictamente unidireccional,
// y que la sincronización bidireccional Apeiron ERP ↔ Sheets (el cron de
// apeiron-erp/backend/server.js que hace CLEAR + reescritura completa de
// esta hoja cada 15 min) nunca vea interferencia de este script.
//
// Lee la fila 1 como encabezados y busca las columnas POR NOMBRE (no por
// posición), así que Apeiron puede reordenar sus 22 columnas sin romper
// nada, siempre que los nombres de encabezado se mantengan.
//
// Columnas reales de "Apeiron - Inventario" (confirmadas contra un export
// de la hoja en julio 2026): Código Apeiron, Referencia Fábrica, Nombre,
// Marca, Categoría, Subcategoría, Precio Venta, Precio Costo, Stock
// Actual, Stock Mínimo, Stock Máximo, Nivel de Stock, Unidad, Aplica IVA,
// Estatus Catálogo, Descripción, Modo de Uso, Beneficios, Link Drive,
// Foto/Video 1 a Foto/Video 7. El sitio solo usa un subconjunto (ver
// abajo); el resto se ignora sin problema gracias a la búsqueda por
// nombre — si Apeiron agrega, quita o reordena columnas que el sitio no
// usa, no rompe nada.
//
// Precio Venta SÍ viene directo de Apeiron (ya no vive en una hoja aparte):
// como Apeiron reescribe esa columna en cada sync, es la fuente de verdad
// única — mantener una copia paralela solo crearía dos precios que se
// pueden desincronizar.
//
// Reglas automáticas:
//   - Estatus Catálogo: se OCULTA solo si el valor normalizado está en
//     ESTADOS_OCULTOS_ (hoy: "inactivo"). Cualquier otro valor —
//     "Activo" (el real, confirmado contra tu hoja), vacío, o algo que
//     Apeiron agregue después— se publica por defecto. Antes era al
//     revés (solo se publicaba si decía exactamente "Publicado"), y como
//     tu columna real usa "Activo"/"Inactivo", esa versión ocultaba el
//     catálogo completo — este es el fix a ese bug.
//   - Stock Actual <= 0 (o vacío/no numérico) NO oculta el producto:
//     se publica igual, marcado como "Agotado" en el sitio, sin botón de
//     carrito, para que la clienta lo conozca y consulte disponibilidad
//     por WhatsApp.
//   - La Categoría se reconoce por palabra clave (sin importar
//     mayúsculas/tildes): "capilar"/"cabello" → Cuidado Capilar,
//     "facial"/"rostro"/"piel" → Cuidado Facial,
//     "corporal"/"cuerpo" → Cuidado Corporal,
//     "maquillaje" → Maquillaje. Si no coincide con ninguna (ej.
//     "Vitaminas", "Servicios"), el producto se omite (no aparece en
//     ningún tab) — el sitio solo tiene esas 4 pestañas hoy.
//   - Las fotos se arman con las columnas `Foto/Video 1` a `Foto/Video 7`
//     (enlaces de archivo individual `drive.google.com/file/d/...`).
//     `Link Drive` NO se usa como imagen: en tu hoja es un link de
//     CARPETA (`drive.google.com/drive/folders/...`), y una carpeta no
//     se puede mostrar como `<img>`.
//   - Precio Costo NUNCA se envía al sitio — solo Precio Venta.
// ═══════════════════════════════════════════════════════════

const ESTADOS_OCULTOS_ = ['inactivo', 'oculto', 'borrador', 'descontinuado'];
const COLUMNAS_FOTO_ = ['Foto/Video 1', 'Foto/Video 2', 'Foto/Video 3', 'Foto/Video 4', 'Foto/Video 5', 'Foto/Video 6', 'Foto/Video 7'];

function leerInventario_() {
  const ss = abrirLibro_();
  const sheet = ss.getSheetByName(SHEET_INVENTARIO);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = {};
  headers.forEach((h, i) => { idx[normalizarTexto_(h)] = i; });
  const tieneEstatus = idx[normalizarTexto_('Estatus Catálogo')] !== undefined;

  function col(fila, nombreColumna) {
    const i = idx[normalizarTexto_(nombreColumna)];
    return (i === undefined) ? '' : fila[i];
  }

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const productos = [];

  values.forEach(fila => {
    const nombre = limpiar_(col(fila, 'Nombre'));
    if (!nombre) return; // fila vacía o incompleta, se omite

    if (tieneEstatus) {
      const estatus = normalizarTexto_(col(fila, 'Estatus Catálogo'));
      if (estatus && ESTADOS_OCULTOS_.indexOf(estatus) !== -1) return; // Inactivo/Oculto/etc.
    }

    // Sin stock YA NO oculta el producto: se publica igual, marcado como
    // agotado, para que la clienta lo conozca y consulte disponibilidad
    // por WhatsApp. El frontend decide cómo mostrarlo según este número.
    const stock = Number(col(fila, 'Stock Actual')) || 0;

    const categoria = mapearCategoria_(limpiar_(col(fila, 'Categoría')));
    if (!categoria) return; // categoría no reconocida, se omite

    const fotos = COLUMNAS_FOTO_
      .map(c => limpiar_(col(fila, c)))
      .filter(Boolean)
      .map(convertirUrlImagen_);

    productos.push({
      codigo: limpiar_(col(fila, 'Código Apeiron')),
      nombre: nombre,
      marca: limpiar_(col(fila, 'Marca')),
      categoria: categoria,
      precioVenta: Number(col(fila, 'Precio Venta')) || 0,
      stock: stock,
      descripcion: limpiar_(col(fila, 'Descripción')),
      modoUso: limpiar_(col(fila, 'Modo de Uso')),
      beneficios: limpiar_(col(fila, 'Beneficios')),
      fotos: fotos
    });
  });

  return productos;
}

function mapearCategoria_(texto) {
  const t = normalizarTexto_(texto);
  if (!t) return null;
  if (t.indexOf('capilar') !== -1 || t.indexOf('cabello') !== -1) return 'capilar';
  if (t.indexOf('facial') !== -1 || t.indexOf('rostro') !== -1 || t.indexOf('piel') !== -1) return 'skincare';
  if (t.indexOf('corporal') !== -1 || t.indexOf('cuerpo') !== -1) return 'corporal';
  if (t.indexOf('maquillaje') !== -1 || t.indexOf('maquillage') !== -1) return 'maquillaje';
  return null;
}

// Convierte un link de "compartir" de Google Drive en uno que sí sirve
// como imagen directa. Si no es un link de Drive, lo deja tal cual.
function convertirUrlImagen_(url) {
  const match = url.match(/drive\.google\.com\/file\/d\/([^/]+)/);
  if (match && match[1]) {
    return 'https://drive.google.com/uc?export=view&id=' + match[1];
  }
  return url;
}

function normalizarTexto_(texto) {
  return String(texto === undefined || texto === null ? '' : texto)
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().trim();
}

// ═══════════════════════════════════════════════════════════
// UTILIDADES COMPARTIDAS
// ═══════════════════════════════════════════════════════════

function limpiar_(valor) {
  return (valor === undefined || valor === null) ? '' : String(valor).trim();
}

function isEmailValido_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════════
// PEDIDOS WEB — compras completadas desde el carrito del sitio
//
// Hoja "Pedidos" — se crea automáticamente la primera vez.
// Columnas: Fecha | NumeroPedido | Estado | MetodoPago |
//   Nombre | Cedula | Telefono | Email | Direccion | Barrio |
//   Municipio | Departamento | Indicaciones | Horario |
//   Productos | TotalItems | Total
//
// El módulo Pedidos del ERP lee esta hoja vía ?accion=pedidos
// y puede actualizar el estado de cada pedido vía POST
// accion=actualizar-pedido.
// ═══════════════════════════════════════════════════════════

function registrarPedido_(data) {
  const sheet = getOrCreateSheetPedidos_();

  const numeroPedido = limpiar_(data.numeroPedido);
  const metodo       = limpiar_(data.metodo);
  const cliente      = data.cliente || {};
  const productos    = Array.isArray(data.productos) ? data.productos : [];
  const total        = Number(data.total) || 0;
  const totalItems   = Number(data.totalItems) || productos.reduce(function(s, p){ return s + (Number(p.cantidad)||1); }, 0);

  if (!numeroPedido) {
    return jsonResponse_({ success: false, error: 'Falta numeroPedido.' });
  }
  if (!productos.length) {
    return jsonResponse_({ success: false, error: 'El pedido no tiene productos.' });
  }

  // Resumen legible de productos: "Nombre x2 — $58.000"
  const resumenProductos = productos
    .map(function(p) {
      return limpiar_(p.nombre) + ' x' + (Number(p.cantidad)||1) + ' — $' + Number(p.total||0).toLocaleString('es-CO');
    })
    .join(' | ');

  sheet.appendRow([
    new Date(),                      // Fecha
    numeroPedido,                    // NumeroPedido
    'Pendiente',                     // Estado inicial
    metodo,                          // MetodoPago
    limpiar_(cliente.nombre),        // Nombre
    limpiar_(cliente.cedula),        // Cedula
    limpiar_(cliente.telefono),      // Telefono
    limpiar_(cliente.email),         // Email
    limpiar_(cliente.direccion),     // Direccion
    limpiar_(cliente.barrio),        // Barrio
    limpiar_(cliente.municipio),     // Municipio
    limpiar_(cliente.departamento),  // Departamento
    limpiar_(cliente.indicaciones),  // Indicaciones
    limpiar_(cliente.horario),       // Horario
    resumenProductos,                // Productos (resumen)
    totalItems,                      // TotalItems
    total                            // Total
  ]);

  return jsonResponse_({ success: true, numeroPedido: numeroPedido });
}

function actualizarEstadoPedido_(data) {
  const sheet = getOrCreateSheetPedidos_();
  const numeroPedido = limpiar_(data.numeroPedido);
  const nuevoEstado  = limpiar_(data.estado);

  if (!numeroPedido || !nuevoEstado) {
    return jsonResponse_({ success: false, error: 'Faltan numeroPedido o estado.' });
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return jsonResponse_({ success: false, error: 'No se encontró el pedido.' });
  }

  // Columna B (índice 2) = NumeroPedido; Columna C (índice 3) = Estado
  const numeros = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  for (var i = 0; i < numeros.length; i++) {
    if (String(numeros[i][0]).trim() === numeroPedido) {
      sheet.getRange(i + 2, 3).setValue(nuevoEstado);  // Columna 3 = Estado
      return jsonResponse_({ success: true });
    }
  }
  return jsonResponse_({ success: false, error: 'Pedido no encontrado: ' + numeroPedido });
}

function leerPedidos_() {
  const sheet = getOrCreateSheetPedidos_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 17).getValues();
  var pedidos = [];

  values.forEach(function(row) {
    var numeroPedido = String(row[1] || '').trim();
    if (!numeroPedido) return;

    // Reconstruir la lista de productos desde el resumen de texto
    // Formato guardado: "Nombre xN — $29.000 | Nombre2 x2 — $58.000"
    var resumenProductos = String(row[14] || '');
    var productosArr = resumenProductos.split(' | ').map(function(p) {
      var match = p.match(/^(.+?) x(\d+) — \$(.+)$/);
      if (match) {
        var cant  = parseInt(match[2], 10) || 1;
        var total = parseInt(match[3].replace(/\./g, ''), 10) || 0;
        return {
          nombre:   match[1].trim(),
          cantidad: cant,
          precio:   cant > 0 ? Math.round(total / cant) : 0,
          total:    total
        };
      }
      return { nombre: p, cantidad: 1, precio: 0, total: 0 };
    });

    pedidos.push({
      timestamp:     row[0] ? new Date(row[0]).toISOString() : '',
      numeroPedido:  numeroPedido,
      estado:        String(row[2]  || 'Pendiente').trim(),
      metodo:        String(row[3]  || '').trim(),
      cliente: {
        nombre:       String(row[4]  || '').trim(),
        cedula:       String(row[5]  || '').trim(),
        telefono:     String(row[6]  || '').trim(),
        email:        String(row[7]  || '').trim(),
        direccion:    String(row[8]  || '').trim(),
        barrio:       String(row[9]  || '').trim(),
        municipio:    String(row[10] || '').trim(),
        departamento: String(row[11] || '').trim(),
        indicaciones: String(row[12] || '').trim(),
        horario:      String(row[13] || '').trim()
      },
      productos:    productosArr,
      totalItems:   Number(row[15]) || 0,
      total:        Number(row[16]) || 0
    });
  });

  // Más reciente primero
  pedidos.sort(function(a, b) { return b.timestamp.localeCompare(a.timestamp); });
  return pedidos;
}

function getOrCreateSheetPedidos_() {
  const ss = abrirLibro_();
  var sheet = ss.getSheetByName(SHEET_PEDIDOS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PEDIDOS);
    var headers = [
      'Fecha', 'NumeroPedido', 'Estado', 'MetodoPago',
      'Nombre', 'Cedula', 'Telefono', 'Email',
      'Direccion', 'Barrio', 'Municipio', 'Departamento',
      'Indicaciones', 'Horario', 'Productos', 'TotalItems', 'Total'
    ];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    // Ancho cómodo para columnas clave
    sheet.setColumnWidth(2, 160);  // NumeroPedido
    sheet.setColumnWidth(5, 200);  // Nombre
    sheet.setColumnWidth(15, 400); // Productos
    sheet.setColumnWidth(17, 100); // Total
    // Formato moneda en columna Total (col 17)
    sheet.getRange('Q2:Q').setNumberFormat('$#,##0');
    // Color de encabezado
    sheet.getRange(1, 1, 1, 17)
      .setBackground('#FF1493')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
  }
  return sheet;
}
