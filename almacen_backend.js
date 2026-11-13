const express = require('express');
const cors = require('cors');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const app = express();
const PORT = 3001;

app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

const UPLOADS_DIR = path.join(__dirname, 'uploads');
if (!fs.existsSync(UPLOADS_DIR)) {
  fs.mkdirSync(UPLOADS_DIR, { recursive: true });
}
app.use('/uploads', express.static(UPLOADS_DIR));

const DB_FILE = fs.existsSync(path.join(__dirname, 'almacen_datos.json'))
  ? path.join(__dirname, 'almacen_datos.json')
  : (fs.existsSync(path.join(__dirname, '..', 'almacen_datos.json'))
      ? path.join(__dirname, '..', 'almacen_datos.json')
      : path.join(__dirname, 'almacen_datos.json'));
const EXCEL_FILE = path.join(__dirname, 'INVENTARIO.xlsx');
const LICENSE_FILE = fs.existsSync(path.join(__dirname, 'cess_license.json'))
  ? path.join(__dirname, 'cess_license.json')
  : path.join(__dirname, '..', 'almacen-app', 'cess_license.json');
const LICENSE_REGISTRY_FILE = fs.existsSync(path.join(__dirname, 'cess_licenses_registry.json'))
  ? path.join(__dirname, 'cess_licenses_registry.json')
  : path.join(__dirname, '..', 'almacen-app', 'cess_licenses_registry.json');
const LICENSE_BACKUP_FILE = path.join(__dirname, 'cess_license.backup.json');
const crypto = require('crypto');
const CESS_MASTER_SECRET = 'CESS_DIESEL_AUTOMATION_2026_PROPRIETARY_AUTH_KEY';
const MASTER_ADMIN_PIN = 'CESS-2026';

function signLicensePayload(payload) {
  const sig = crypto
    .createHmac('sha256', CESS_MASTER_SECRET)
    .update(JSON.stringify(payload))
    .digest('hex')
    .slice(0, 24)
    .toUpperCase();
  const formattedKey = 'CESS-' + (sig.match(/.{1,4}/g) || [sig]).join('-');
  return {
    license_key: formattedKey,
    signature: sig,
    payload
  };
}

function getLicenseStatus() {
  if (!fs.existsSync(LICENSE_FILE)) {
    return {
      active: false,
      status: 'SIN_LICENCIA',
      message: 'Sistema no autorizado. No se encontró archivo de licencia CESS.',
      licenciatario: 'NO AUTORIZADO'
    };
  }

  try {
    const lic = JSON.parse(fs.readFileSync(LICENSE_FILE, 'utf-8'));
    if (lic.status === 'REVOCADA' || lic.active === false) {
      return {
        active: false,
        status: 'REVOCADA',
        message: lic.motivo || 'Licencia suspendida o revocada por el titular de CESS. Acceso bloqueado.',
        fecha_bloqueo: lic.fecha_bloqueo || new Date().toISOString()
      };
    }
    if (!lic.payload || !lic.signature) {
      return { active: false, status: 'CORRUPTA', message: 'Licencia alterada o formato inválido.' };
    }

    const expectedSig = crypto
      .createHmac('sha256', CESS_MASTER_SECRET)
      .update(JSON.stringify(lic.payload))
      .digest('hex')
      .slice(0, 24)
      .toUpperCase();

    if (lic.signature !== expectedSig) {
      return { active: false, status: 'INVALIDA', message: 'Firma criptográfica de licencia no coincide.' };
    }

    if (lic.payload.expiracion) {
      const expDate = new Date(lic.payload.expiracion);
      if (expDate < new Date()) {
        return {
          active: false,
          status: 'EXPIRADA',
          message: `La licencia expiró el ${lic.payload.expiracion}. Contacte al titular de CESS para renovar suscripción.`,
          licenciatario: lic.payload.licenciatario,
          expiracion: lic.payload.expiracion
        };
      }
    }

    return {
      active: true,
      status: 'AUTORIZADA',
      message: 'Licencia CESS comercial autorizada y válida.',
      licenciatario: lic.payload.licenciatario,
      plan: lic.payload.plan,
      emision: lic.payload.emision,
      expiracion: lic.payload.expiracion || 'PERPETUA / SIN VENCIMIENTO',
      id_licencia: lic.payload.id_licencia,
      autor: lic.payload.autor,
      producto: lic.payload.producto,
      version: lic.payload.version,
      modulos: lic.payload.modulos || []
    };
  } catch (err) {
    return { active: false, status: 'ERROR', message: err.message };
  }
}

// Estructura de almacenamiento
let store = {
  activeSheet: '',
  lastSync: null,
  catalog: {}, // id -> { id, rack, nivel, descripcion, marca, numero_parte, unidad, precio_unitario, stock_inicial, stock_actual }
  movimientos: [], // { id, tipo: 'E'|'S', fecha, hora, numero_parte, descripcion, marca, ubicacion, cantidad, unidad_camion, proveedor, folio, factura, solicitante, entregado_por, costo_unitario, foto, timestamp }
  facturas_procesadas: [], // { id, factura, factura_normalizada, proveedor, proveedor_normalizado, fecha_factura, fecha_ingreso, hora_ingreso, total_monto, total_piezas, total_items, image_sha256, entregado_por, piezas_resumen }
  usuarios: [], // { id, nombre, pin, rol: 'ADMIN'|'GERENTE'|'OPERADOR'|'TALLER', activo: true, permisos: { ver_inventario, dar_salida, dar_entrada, ver_bitacora, exportar_excel, admin_centro_control }, creado_en }
  memoria_ia: [], // { id, texto_original, texto_normalizado, codigo_extraido, marca_detectada, proveedor, id_pieza_catalogo, numero_parte_catalogo, descripcion_catalogo, veces_utilizado, fecha_registro, origen }
  folios_taller: [] // { folio, unidad, solicitante, motivo, refacciones_solicitadas, fecha, hora, estado_taller, estatus, creado_desde, timestamp }
};

// =========================================================
// FUNCIONES DE CONTROL ANTI-DUPLICADOS DE FACTURAS Y REMISIONES
// =========================================================
function normalizeFacturaKey(str) {
  if (!str) return '';
  return String(str)
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '')
    .trim();
}

function normalizeProveedorKey(str) {
  if (!str) return '';
  return String(str)
    .toUpperCase()
    .replace(/\b(S\.?A\.?|DE|C\.?V\.?|S\.?R\.?L\.?|S\.?A\.?P\.?I\.?)\b/g, '')
    .replace(/[^A-Z0-9]/g, '')
    .trim();
}

function isGenericInvoice(str) {
  if (!str) return true;
  const n = normalizeFacturaKey(str);
  if (!n || n.length < 2) return true;
  const generics = ['SF', 'SINFACTURA', 'REMISION', 'MOSTRADOR', 'NA', 'SN', 'NOTADEVENTA', 'TICKET', 'FACTURA', 'REMIA', 'REMISIÓN'];
  return generics.includes(n);
}

function findExistingFactura(factura, proveedor, imageHash = null) {
  if (!Array.isArray(store.facturas_procesadas)) store.facturas_procesadas = [];

  // Helper para verificar si una factura fue totalmente anulada en la bitácora
  const isFacturaAnulada = (f) => {
    if (!f) return false;
    if (f.anulada) return true;
    if (!f.factura) return false;
    const normF = normalizeFacturaKey(f.factura);
    const movs = (store.movimientos || []).filter(m => m.factura && normalizeFacturaKey(m.factura) === normF);
    if (movs.length > 0 && movs.every(m => m.anulado || m.tipo === 'ANULADO')) {
      return true;
    }
    return false;
  };

  // 1. Chequeo por hash de imagen exacta (si suben el mismo archivo)
  if (imageHash) {
    const matchByHash = store.facturas_procesadas.find(f => !isFacturaAnulada(f) && f.image_sha256 && f.image_sha256 === imageHash);
    if (matchByHash) {
      return { match: matchByHash, reason: 'IMAGEN_IDENTICA' };
    }
  }

  // 2. Chequeo por folio de factura y proveedor
  if (!factura || isGenericInvoice(factura)) {
    return null;
  }

  const normFact = normalizeFacturaKey(factura);
  const normProv = normalizeProveedorKey(proveedor);

  // Buscar en facturas_procesadas registradas
  const match = store.facturas_procesadas.find(f => {
    if (isFacturaAnulada(f)) return false;
    if (f.factura_normalizada === normFact) {
      if (!normProv || !f.proveedor_normalizado) return true;
      return f.proveedor_normalizado === normProv || f.proveedor_normalizado.includes(normProv) || normProv.includes(f.proveedor_normalizado);
    }
    return false;
  });

  if (match) {
    return { match, reason: 'FOLIO_PROVEEDOR' };
  }

  // 3. Salvaguarda: Buscar en movimientos históricos
  if (store.movimientos && Array.isArray(store.movimientos)) {
    const movMatch = store.movimientos.find(m => {
      if ((m.tipo === 'E' || m.tipo === 'E/S') && m.factura && !isGenericInvoice(m.factura)) {
        if (normalizeFacturaKey(m.factura) === normFact) {
          if (!normProv || !m.proveedor) return true;
          const movProv = normalizeProveedorKey(m.proveedor);
          return movProv === normProv || movProv.includes(normProv) || normProv.includes(movProv);
        }
      }
      return false;
    });

    if (movMatch) {
      return {
        match: {
          factura: movMatch.factura,
          proveedor: movMatch.proveedor || 'Proveedor registrado',
          fecha_ingreso: movMatch.fecha,
          hora_ingreso: movMatch.hora || '',
          total_piezas: movMatch.cantidad,
          total_monto: (movMatch.cantidad || 0) * (movMatch.costo_unitario || 0),
          piezas_resumen: [{ numero_parte: movMatch.numero_parte, cantidad: movMatch.cantidad, descripcion: movMatch.descripcion }]
        },
        reason: 'MOVIMIENTO_PREVIO'
      };
    }
  }

  return null;
}

function calcularVencimiento(fechaBase, dias) {
  if (!dias || isNaN(Number(dias)) || Number(dias) <= 0) return fechaBase || new Date().toISOString().slice(0, 10);
  try {
    const base = fechaBase || new Date().toISOString().slice(0, 10);
    const d = new Date(base + 'T12:00:00');
    d.setDate(d.getDate() + Number(dias));
    return d.toISOString().slice(0, 10);
  } catch (e) {
    return fechaBase || new Date().toISOString().slice(0, 10);
  }
}

function registrarOActualizarFacturaProcesada(mov) {
  if (!mov.factura || isGenericInvoice(mov.factura)) return;
  if (!Array.isArray(store.facturas_procesadas)) store.facturas_procesadas = [];

  const factNorm = normalizeFacturaKey(mov.factura);
  const provNorm = normalizeProveedorKey(mov.proveedor);
  let fac = store.facturas_procesadas.find(f => f.factura_normalizada === factNorm && (!provNorm || f.proveedor_normalizado === provNorm));

  const cant = Number(mov.cantidad) || 1;
  const costo = Number(mov.costo_unitario) || 0;
  const subtotal = cant * costo;

  const cond = mov.condicion_pago || 'CONTADO';
  const diasCred = Number(mov.dias_credito) || 0;
  const fFact = mov.fecha_factura || mov.fecha;
  const fVenc = mov.fecha_vencimiento || (cond === 'CREDITO' ? calcularVencimiento(fFact, diasCred || 30) : fFact);
  const estPago = mov.estado_pago || (cond === 'CREDITO' ? 'PENDIENTE' : 'PAGADO');

  if (fac) {
    fac.total_piezas = (Number(fac.total_piezas) || 0) + cant;
    fac.total_monto = (Number(fac.total_monto) || 0) + subtotal;
    fac.total_items = (Number(fac.total_items) || 0) + 1;
    if (mov.condicion_pago) fac.condicion_pago = cond;
    if (mov.dias_credito !== undefined) fac.dias_credito = diasCred;
    if (mov.fecha_vencimiento) fac.fecha_vencimiento = fVenc;
    if (mov.estado_pago) fac.estado_pago = estPago;
    if (mov.fecha_pago) fac.fecha_pago = mov.fecha_pago;
    if (mov.referencia_pago) fac.referencia_pago = mov.referencia_pago;
    if (!Array.isArray(fac.piezas_resumen)) fac.piezas_resumen = [];
    if (!fac.piezas_resumen.some(p => p.numero_parte === mov.numero_parte)) {
      fac.piezas_resumen.push({
        numero_parte: mov.numero_parte,
        descripcion: mov.descripcion,
        cantidad: cant,
        costo_unitario: costo
      });
    }
  } else {
    fac = {
      id: `fac_${Date.now()}_${Math.round(Math.random() * 1000)}`,
      factura: mov.factura,
      factura_normalizada: factNorm,
      proveedor: mov.proveedor || 'PROVEEDOR REGISTRADO',
      proveedor_normalizado: provNorm,
      fecha_factura: fFact,
      fecha_ingreso: mov.fecha,
      hora_ingreso: mov.hora || '12:00:00',
      total_monto: subtotal,
      total_piezas: cant,
      total_items: 1,
      condicion_pago: cond,
      dias_credito: diasCred,
      fecha_vencimiento: fVenc,
      estado_pago: estPago,
      fecha_pago: mov.fecha_pago || '',
      referencia_pago: mov.referencia_pago || '',
      notas_pago: '',
      entregado_por: mov.entregado_por || '',
      piezas_resumen: [{
        numero_parte: mov.numero_parte,
        descripcion: mov.descripcion,
        cantidad: cant,
        costo_unitario: costo
      }]
    };
    store.facturas_procesadas.unshift(fac);
  }
}

function backfillFacturasFromMovimientos() {
  if (!Array.isArray(store.facturas_procesadas)) {
    store.facturas_procesadas = [];
  }

  const existingKeys = new Set(store.facturas_procesadas.map(f => `${f.factura_normalizada}_${f.proveedor_normalizado || ''}`));
  const movsPorFactura = {};

  for (const m of store.movimientos || []) {
    if ((m.tipo === 'E' || m.tipo === 'E/S') && m.factura && !isGenericInvoice(m.factura)) {
      const factNorm = normalizeFacturaKey(m.factura);
      const provNorm = normalizeProveedorKey(m.proveedor);
      const key = `${factNorm}_${provNorm}`;

      if (!movsPorFactura[key]) {
        movsPorFactura[key] = {
          id: `fac_hist_${factNorm}`,
          factura: m.factura,
          factura_normalizada: factNorm,
          proveedor: m.proveedor || 'PROVEEDOR REGISTRADO',
          proveedor_normalizado: provNorm,
          fecha_factura: m.fecha,
          fecha_ingreso: m.fecha,
          hora_ingreso: m.hora || '12:00:00',
          total_monto: 0,
          total_piezas: 0,
          total_items: 0,
          piezas_resumen: []
        };
      }

      const cant = Number(m.cantidad) || 1;
      const costo = Number(m.costo_unitario) || 0;
      movsPorFactura[key].total_piezas += cant;
      movsPorFactura[key].total_monto += (cant * costo);
      movsPorFactura[key].total_items += 1;
      if (movsPorFactura[key].piezas_resumen.length < 10) {
        movsPorFactura[key].piezas_resumen.push({
          numero_parte: m.numero_parte,
          cantidad: cant,
          descripcion: m.descripcion,
          costo_unitario: costo
        });
      }
    }
  }

  let agregadas = 0;
  for (const [key, facObj] of Object.entries(movsPorFactura)) {
    if (!existingKeys.has(key)) {
      store.facturas_procesadas.push(facObj);
      existingKeys.add(key);
      agregadas++;
    }
  }

  if (agregadas > 0) {
    saveStore();
  }
}

function loadStore() {
  if (fs.existsSync(DB_FILE)) {
    try {
      const data = JSON.parse(fs.readFileSync(DB_FILE, 'utf-8'));
      store = { ...store, ...data };
      if (!Array.isArray(store.facturas_procesadas)) {
        store.facturas_procesadas = [];
      }
      if (!Array.isArray(store.ordenes_compra)) {
        store.ordenes_compra = [];
      }
      if (!Array.isArray(store.memoria_ia)) {
        store.memoria_ia = [];
      }
      if (!Array.isArray(store.folios_taller)) {
        store.folios_taller = [];
      }
      if (!Array.isArray(store.usuarios) || store.usuarios.length === 0) {
        store.usuarios = [
          {
            id: 'usr_cesar',
            nombre: 'César Almacén (Administrador)',
            pin: 'CESS-2026',
            rol: 'ADMIN',
            activo: true,
            permisos: {
              ver_inventario: true,
              dar_salida: true,
              dar_entrada: true,
              ver_bitacora: true,
              exportar_excel: true,
              admin_centro_control: true
            },
            creado_en: new Date().toISOString()
          },
          {
            id: 'usr_gerente',
            nombre: 'Gerencia General (Solo Lectura)',
            pin: '1111',
            rol: 'GERENTE',
            activo: true,
            permisos: {
              ver_inventario: true,
              dar_salida: false,
              dar_entrada: false,
              ver_bitacora: true,
              exportar_excel: true,
              admin_centro_control: false
            },
            creado_en: new Date().toISOString()
          },
          {
            id: 'usr_taller',
            nombre: 'Jefe de Taller (Despachos)',
            pin: '2222',
            rol: 'TALLER',
            activo: true,
            permisos: {
              ver_inventario: true,
              dar_salida: true,
              dar_entrada: false,
              ver_bitacora: true,
              exportar_excel: false,
              admin_centro_control: false
            },
            creado_en: new Date().toISOString()
          }
        ];
      }

      // Asegurar que todos los usuarios tengan campo de empresa y permisos granulares actualizados
      if (!store.config) {
        store.config = {};
      }
      if (Array.isArray(store.usuarios)) {
        const licStatus = getLicenseStatus();
        const currentLicEmpresa = (licStatus && licStatus.licenciatario) ? licStatus.licenciatario : 'LARMONT DIESEL';
        store.usuarios.forEach(u => {
          if (!u.empresa) {
            u.empresa = currentLicEmpresa;
          }
          if (u.id === 'usr_taller') {
            u.rol = 'TALLER';
            if (!u.permisos) u.permisos = {};
            u.permisos.dar_salida = true;
          }
          if (u.permisos) {
            if (u.permisos.cuadre_inventario === undefined) {
              u.permisos.cuadre_inventario = u.rol === 'ADMIN';
            }
            if (u.permisos.compras === undefined) {
              u.permisos.compras = u.rol === 'ADMIN' || u.rol === 'GERENTE';
            }
          }
        });
      }
      if (Array.isArray(store.facturas_procesadas)) {
        store.facturas_procesadas.forEach(f => {
          if (!f.condicion_pago) f.condicion_pago = 'CONTADO';
          if (f.dias_credito === undefined) f.dias_credito = 0;
          if (!f.estado_pago) f.estado_pago = f.condicion_pago === 'CREDITO' ? 'PENDIENTE' : 'PAGADO';
          if (!f.fecha_vencimiento) f.fecha_vencimiento = f.fecha_factura || f.fecha_ingreso || new Date().toISOString().slice(0, 10);
        });
      }
      if (Array.isArray(store.movimientos)) {
        store.movimientos.forEach(m => {
          if ((m.tipo === 'E' || m.tipo === 'E/S') && !m.condicion_pago) {
            m.condicion_pago = 'CONTADO';
            m.dias_credito = 0;
            m.estado_pago = 'PAGADO';
            m.fecha_vencimiento = m.fecha;
          }
        });
      }
      if (store.movimientos && store.movimientos.length > 0) {
        backfillFacturasFromMovimientos();
      }
      ensureAlmacenes();
      console.log(`Base de datos cargada: ${Object.keys(store.catalog).length} productos (${(store.almacenes || []).length} almacenes), ${store.movimientos.length} movimientos, ${store.facturas_procesadas.length} facturas, ${store.usuarios.length} usuarios.`);
    } catch (e) {
      console.error('Error leyendo base de datos local:', e);
    }
  }
}

function saveStore() {
  try {
    const jsonStr = JSON.stringify(store, null, 2);
    fs.writeFileSync(DB_FILE, jsonStr, 'utf-8');
    const rootAlm = path.join(__dirname, '..', 'almacen_datos.json');
    const cloudAlm = path.join(__dirname, '..', 'taller-cloud-deploy', 'almacen_datos.json');
    try { fs.writeFileSync(rootAlm, jsonStr, 'utf-8'); } catch(e) {}
    try { fs.writeFileSync(cloudAlm, jsonStr, 'utf-8'); } catch(e) {}
  } catch (e) {
    console.error('Error guardando base de datos local:', e);
  }
}

// Configuración de Multer para fotos
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOADS_DIR),
  filename: (req, file, cb) => {
    const ext = path.extname(file.originalname) || '.jpg';
    cb(null, `foto_${Date.now()}_${Math.round(Math.random() * 1e5)}${ext}`);
  }
});
const upload = multer({ storage });

// =========================================================
// GESTIÓN MULTI-ALMACÉN Y DEPARTAMENTOS AUTÓNOMOS CESS
// =========================================================
function ensureAlmacenes() {
  const DEFAULT_ALMACENES = [
    {
      id: 'refacciones',
      nombre: 'Almacén de Refacciones',
      codigo: 'REF',
      icono: 'Package',
      color: '#2563EB',
      descripcion: 'Almacén central y de refacciones diésel'
    },
    {
      id: 'llantas',
      nombre: 'Llantas',
      codigo: 'LLAN',
      icono: 'Disc',
      color: '#D97706',
      descripcion: 'Neumáticos, rines y cascos de tracto y remolque'
    },
    {
      id: 'ferreteria',
      nombre: 'Ferretería',
      codigo: 'FERR',
      icono: 'Wrench',
      color: '#059669',
      descripcion: 'Tornillería, abrazaderas, ferretería y misceláneos'
    }
  ];

  if (!Array.isArray(store.almacenes) || store.almacenes.length === 0) {
    store.almacenes = DEFAULT_ALMACENES;
  } else {
    // Asegurar que los almacenes estándar existan si la lista está incompleta
    for (const def of DEFAULT_ALMACENES) {
      const exists = store.almacenes.find(a => a.id === def.id);
      if (!exists) {
        store.almacenes.push(def);
      }
    }
  }

  // Mapa dinámico de ID a Nombre actual del almacén
  const almMap = {};
  for (const a of store.almacenes) {
    almMap[a.id] = a.nombre;
  }

  // Asegurar que cada producto tenga su almacén estandarizado
  for (const [id, item] of Object.entries(store.catalog || {})) {
    if (!item.almacen_id) {
      if (item.almacen === 'Refacciones Diesel' || item.almacen === 'Refacciones') item.almacen_id = 'refacciones';
      else if (item.almacen === 'Almacén Llantas' || item.almacen === 'LLANTAS') item.almacen_id = 'llantas';
      else if (item.almacen === 'Ferretería y Tornillería' || item.almacen === 'FERRETERIA') item.almacen_id = 'ferreteria';
      else if (item.hoja_origen === 'ALMACEN LLANTAS') item.almacen_id = 'llantas';
      else if (item.hoja_origen === 'ALMACEN 2') item.almacen_id = 'ferreteria';
      else {
        const found = store.almacenes.find(a => a.nombre === item.almacen);
        item.almacen_id = found ? found.id : 'refacciones';
      }
    }
    // Sincronizar el nombre legible con la configuración actual del almacén
    if (almMap[item.almacen_id]) {
      item.almacen = almMap[item.almacen_id];
    } else {
      item.almacen_id = 'refacciones';
      item.almacen = almMap['refacciones'] || 'Almacén de Refacciones';
    }
  }

  // Asegurar que cada movimiento en bitácora tenga su almacén estandarizado
  for (const m of (store.movimientos || [])) {
    if (!m.almacen_id) {
      const found = store.almacenes.find(a => a.nombre === m.almacen);
      m.almacen_id = found ? found.id : 'refacciones';
    }
    if (almMap[m.almacen_id]) {
      m.almacen = almMap[m.almacen_id];
    }
    if (m.almacen_anterior === 'Refacciones Diesel') m.almacen_anterior = almMap['refacciones'] || 'Almacén de Refacciones';
    if (m.almacen_nuevo === 'Refacciones Diesel') m.almacen_nuevo = almMap['refacciones'] || 'Almacén de Refacciones';
  }
}

// Función para parsear una hoja individual de Excel respetando cabeceras dinámicas
function parseItemsFromSheet(workbook, sheetName, alm) {
  if (!workbook.Sheets[sheetName]) return [];
  const rows = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: null });
  if (!rows || rows.length < 2) return [];

  // Buscar fila de encabezados en filas 0..4
  let headerRowIdx = 0;
  for (let r = 0; r < Math.min(5, rows.length); r++) {
    if (!rows[r]) continue;
    const rowStr = rows[r].map(c => String(c || '').toUpperCase()).join(' ');
    if (rowStr.includes('DESCRIP') || (rowStr.includes('PARTE') && rowStr.includes('CANTIDAD'))) {
      headerRowIdx = r;
      break;
    }
  }

  // Detectar posible Rack en la fila previa (ej. "ALMACEN 2 PISO RACK 1")
  let defaultRack = 'Rack 1';
  if (headerRowIdx > 0 && rows[0] && rows[0][0]) {
    const titleStr = String(rows[0][0]).trim();
    if (titleStr.toUpperCase().includes('RACK')) {
      const matchR = titleStr.match(/RACK\s*\d+/i);
      if (matchR) defaultRack = matchR[0];
      else defaultRack = titleStr;
    }
  }

  const headers = (rows[headerRowIdx] || []).map(h => (h ? String(h).trim().toUpperCase() : ''));
  const colRack = headers.findIndex(h => h.includes('RACK'));
  const colNivel = headers.findIndex(h => h.includes('NIVEL'));
  const colDesc = headers.findIndex(h => h.includes('DESCRIPCI') || h.includes('DESCRIP'));
  const colMarca = headers.findIndex(h => h.includes('MARCA'));
  const colParte = headers.findIndex(h => h.includes('PARTE') || h.includes('NUMERO'));
  const colUnidad = headers.findIndex(h => h === 'UNIDAD');
  const colPrecio = headers.findIndex(h => h.includes('PRECIO') || h.includes('UNITARIO'));
  const colFinal = headers.findIndex(h => h.includes('CANTIDAD FINAL') || h.includes('FINAL'));
  const colInicial = headers.findIndex(h => h.includes('INICIAL') || h.includes('INICAL'));
  const colCantidad = headers.findIndex(h => h === 'CANTIDAD' || h.startsWith('CANTIDAD'));

  const items = [];
  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r) continue;

    const desc = colDesc !== -1 && r[colDesc] ? String(r[colDesc]).trim() : '';
    const parteRaw = colParte !== -1 && r[colParte] !== null && r[colParte] !== undefined ? String(r[colParte]).trim() : '';

    if (!desc && !parteRaw) continue;

    // Generar un ID único que preserve los IDs existentes para refacciones
    let parte = parteRaw;
    if (!parte) {
      parte = desc ? desc.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 20) + `_${i}` : `ITEM-${i}`;
    }

    let id = parte.replace(/[^a-zA-Z0-9_-]/g, '_');
    if (alm.id !== 'refacciones') {
      id = `${alm.id}_${id}`;
    }

    const rackVal = colRack !== -1 && r[colRack] !== null && r[colRack] !== undefined ? String(r[colRack]).trim() : defaultRack;
    const nivelVal = colNivel !== -1 && r[colNivel] !== null && r[colNivel] !== undefined ? String(r[colNivel]).trim() : 'Nivel 1';
    const marcaVal = colMarca !== -1 && r[colMarca] !== null && r[colMarca] !== undefined ? String(r[colMarca]).trim() : 'GENERICA';
    const unidadMedida = colUnidad !== -1 && r[colUnidad] ? String(r[colUnidad]).trim() : 'PZ';

    let stock = 0;
    if (colFinal !== -1 && r[colFinal] !== null && !isNaN(Number(r[colFinal]))) {
      stock = Number(r[colFinal]);
    } else if (colCantidad !== -1 && r[colCantidad] !== null && !isNaN(Number(r[colCantidad]))) {
      stock = Number(r[colCantidad]);
    } else if (colInicial !== -1 && r[colInicial] !== null && !isNaN(Number(r[colInicial]))) {
      stock = Number(r[colInicial]);
    }

    let precio = 0;
    if (colPrecio !== -1 && r[colPrecio] !== null && !isNaN(Number(r[colPrecio]))) {
      precio = Number(r[colPrecio]);
    }

    const existing = store.catalog ? (store.catalog[id] || (typeof findPart === 'function' ? findPart(parte) : null) || {}) : {};
    const safeStock = Math.max(0, stock);
    items.push({
      id,
      almacen: alm.nombre,
      almacen_id: alm.id,
      hoja_origen: sheetName,
      rack: rackVal ? (rackVal.toLowerCase().startsWith('rack') ? rackVal : `Rack ${rackVal}`) : 'Rack 1',
      nivel: nivelVal || 'Nivel 1',
      descripcion: desc || parte,
      marca: marcaVal,
      numero_parte: parte,
      foto_referencia: existing.foto_referencia || '',
      equivalencias: Array.isArray(existing.equivalencias) ? [...existing.equivalencias] : [],
      alias: Array.isArray(existing.alias) ? [...existing.alias] : [],
      es_a_granel: existing.es_a_granel || false,
      unidad: unidadMedida,
      precio_unitario: precio,
      stock_inicial: safeStock,
      stock_actual: safeStock
    });
  }

  return items;
}

// Función para leer y sincronizar catálogo desde INVENTARIO.xlsx con soporte Multi-Almacén
function syncFromExcel(targetSheet, syncMode = 'ALL') {
  if (!fs.existsSync(EXCEL_FILE)) {
    throw new Error('El archivo INVENTARIO.xlsx no se encuentra en ' + EXCEL_FILE);
  }

  ensureAlmacenes();
  const workbook = xlsx.readFile(EXCEL_FILE);
  const sheets = workbook.SheetNames;
  
  // Buscar hojas con formato de fecha (DD-MM-YYYY o DD-MM-YY) ordenadas por fecha más reciente
  const dateSheets = sheets.filter(s => /^\d{2}-\d{2}-\d{2,4}$/.test(s.trim())).sort((a, b) => {
    const pA = a.split('-').map(Number);
    const pB = b.split('-').map(Number);
    const yA = pA[2] < 100 ? 2000 + pA[2] : pA[2];
    const yB = pB[2] < 100 ? 2000 + pB[2] : pB[2];
    const tA = new Date(yA, pA[1] - 1, pA[0]).getTime();
    const tB = new Date(yB, pB[1] - 1, pB[0]).getTime();
    return tB - tA; // Más reciente primero
  });
  let chosenDateSheet = targetSheet && dateSheets.includes(targetSheet)
    ? targetSheet
    : (store.activeSheet && dateSheets.includes(store.activeSheet) ? store.activeSheet : (dateSheets.length > 0 ? dateSheets[0] : sheets[0]));

  store.activeSheet = chosenDateSheet;
  const almRef = store.almacenes.find(a => a.id === 'refacciones');
  if (almRef) almRef.hoja = chosenDateSheet;

  let newCatalog = {};
  if (syncMode === 'SINGLE' && targetSheet && store.catalog) {
    // Si solo sincronizamos una hoja específica, preservar los productos de los otros almacenes
    for (const [id, item] of Object.entries(store.catalog)) {
      if (item.hoja_origen !== targetSheet && item.almacen_id !== targetSheet) {
        newCatalog[id] = item;
      }
    }
  }

  // Parsear cada almacén configurado
  for (const alm of store.almacenes) {
    let sheetToParse = alm.hoja;
    if (alm.id === 'refacciones' || alm.esFecha) {
      sheetToParse = chosenDateSheet;
    }
    if (syncMode === 'SINGLE' && targetSheet && targetSheet !== sheetToParse && targetSheet !== alm.id) {
      continue;
    }

    if (sheets.includes(sheetToParse)) {
      const items = parseItemsFromSheet(workbook, sheetToParse, alm);
      for (const it of items) {
        newCatalog[it.id] = it;
      }
    }
  }

  // Recalcular stock_actual con los movimientos registrados en la app
  for (const mov of store.movimientos) {
    if (mov.anulado) continue;
    // Omitir movimientos históricos importados de hojas anteriores a la fecha de inventario activo
    // para evitar doble deducción sobre la existencia ya computada en la hoja
    if (mov.id && String(mov.id).startsWith('hist_')) {
      continue;
    }
    const item = newCatalog[mov.id_pieza] || Object.values(newCatalog).find(p => p.numero_parte === mov.numero_parte);
    if (item) {
      if (mov.tipo === 'E' || mov.tipo === 'DEV') {
        item.stock_actual += Number(mov.cantidad);
        if (mov.costo_unitario) item.precio_unitario = Number(mov.costo_unitario);
      } else if (mov.tipo === 'S') {
        item.stock_actual -= Number(mov.cantidad);
      } else if (mov.tipo === 'AJUSTE') {
        if (mov.diferencia_stock !== undefined) {
          item.stock_actual += Number(mov.diferencia_stock);
        } else if (mov.stock_nuevo !== undefined) {
          item.stock_actual = Number(mov.stock_nuevo);
        }
        if (mov.ubicacion_nueva) {
          const uParts = mov.ubicacion_nueva.split('/');
          if (uParts.length >= 2) {
            item.rack = uParts[0].trim();
            item.nivel = uParts[1].trim();
          }
        }
        if (mov.almacen_nuevo_id) {
          item.almacen_id = mov.almacen_nuevo_id;
        }
        if (mov.almacen_nuevo) {
          item.almacen = mov.almacen_nuevo;
        }
      }
    }
  }

  // Reconciliación de salvaguarda física: en anaquel no puede haber saldo negativo
  for (const item of Object.values(newCatalog)) {
    if (item.stock_inicial < 0) item.stock_inicial = 0;
    if (item.stock_actual < 0) item.stock_actual = 0;
  }

  store.catalog = newCatalog;
  store.lastSync = new Date().toISOString();
  saveStore();

  return {
    chosenSheet: chosenDateSheet,
    count: Object.keys(newCatalog).length,
    allSheets: sheets,
    almacenes: store.almacenes
  };
}

// Inicializar al arrancar
loadStore();
if (Object.keys(store.catalog).length === 0) {
  try {
    syncFromExcel();
  } catch (err) {
    console.log('Esperando archivo Excel inicial:', err.message);
  }
}

// ==================== ENDPOINTS DE LA API ====================

// 1. Estado del sistema y hojas disponibles
app.get('/api/status', (req, res) => {
  let allSheets = [];
  if (fs.existsSync(EXCEL_FILE)) {
    try {
      const wb = xlsx.readFile(EXCEL_FILE, { bookSheets: true });
      allSheets = wb.SheetNames;
    } catch (e) {}
  }

  const todayStr = new Date().toISOString().slice(0, 10);
  const movsHoy = store.movimientos.filter(m => m.fecha === todayStr);

  const lic = getLicenseStatus();
  const localIp = getLocalIp();
  const localUrl = `http://${localIp}:${PORT}`;
  let tunnelUrl = '';
  const TUNNEL_URL_FILE = path.join(__dirname, 'active_tunnel_url.txt');
  if (fs.existsSync(TUNNEL_URL_FILE)) {
    try {
      const saved = fs.readFileSync(TUNNEL_URL_FILE, 'utf-8').trim();
      if (saved && saved.startsWith('http')) tunnelUrl = saved;
    } catch (_) {}
  }
  if (!tunnelUrl) {
    const TUNNEL_LOG_FILE = path.join(__dirname, 'tunnel.log');
    if (fs.existsSync(TUNNEL_LOG_FILE)) {
      try {
        const logContent = fs.readFileSync(TUNNEL_LOG_FILE, 'utf-8');
        const matches = logContent.match(/https:\/\/[a-z0-9\-]+\.trycloudflare\.com/g);
        if (matches && matches.length > 0) {
          tunnelUrl = matches[matches.length - 1];
        }
      } catch (_) {}
    }
  }
  if (!tunnelUrl) {
    tunnelUrl = 'https://legacy-attention-dark-chart.trycloudflare.com';
  }

  res.json({
    activeSheet: store.activeSheet,
    lastSync: store.lastSync,
    totalProductos: Object.keys(store.catalog).length,
    totalMovimientos: store.movimientos.length,
    movimientosHoy: movsHoy.length,
    entradasHoy: movsHoy.filter(m => !m.anulado && (m.tipo === 'E' || m.tipo === 'E/S' || m.tipo === 'DEV')).reduce((acc, m) => acc + Number(m.cantidad), 0),
    salidasHoy: movsHoy.filter(m => !m.anulado && (m.tipo === 'S' || m.tipo === 'E/S')).reduce((acc, m) => acc + Number(m.cantidad), 0),
    pasosDirectosHoy: movsHoy.filter(m => !m.anulado && m.tipo === 'E/S').reduce((acc, m) => acc + Number(m.cantidad), 0),
    ajustesHoy: movsHoy.filter(m => !m.anulado && m.tipo === 'AJUSTE').length,
    excelFileExists: fs.existsSync(EXCEL_FILE),
    allSheets,
    licencia: lic,
    localUrl,
    tunnelUrl,
    hasGeminiKey: Boolean(process.env.GEMINI_API_KEY || (store.config && store.config.gemini_api_key))
  });
});

// Guardar API Key de Gemini globalmente en el servidor
app.post('/api/config/gemini-key', (req, res) => {
  const { apiKey } = req.body || {};
  if (!store.config) store.config = {};
  store.config.gemini_api_key = String(apiKey || '').trim();
  saveStore();
  res.json({ success: true, message: 'API Key de Gemini guardada globalmente en el servidor.' });
});

// 1.1 Endpoint de Licencia y Propiedad Intelectual CESS
app.get('/api/licencia', (req, res) => {
  const lic = getLicenseStatus();
  res.json({
    licencia: lic,
    copyright: '© 2026 CESS Systems. Software Propietario. Todos los derechos reservados por el Autor.',
    titular: 'Desarrollador y Propietario CESS'
  });
});

// 1.2 Activar nueva licencia (soporta JSON completo o Clave CESS)
app.post('/api/activar-licencia', (req, res) => {
  try {
    const { licencia_json, license_key } = req.body;
    let licObj = null;

    if (licencia_json) {
      licObj = typeof licencia_json === 'string' ? JSON.parse(licencia_json) : licencia_json;
    } else if (license_key) {
      // Buscar en el registro histórico
      if (fs.existsSync(LICENSE_REGISTRY_FILE)) {
        try {
          const reg = JSON.parse(fs.readFileSync(LICENSE_REGISTRY_FILE, 'utf-8'));
          const found = reg.find(r => r.license_key === license_key.trim());
          if (found && found.licencia_completa) {
            licObj = found.licencia_completa;
          }
        } catch (_) {}
      }
    }

    if (!licObj) {
      return res.status(400).json({ success: false, error: 'No se proporcionó una licencia válida o no se encontró la clave en el registro.' });
    }

    fs.writeFileSync(LICENSE_FILE, JSON.stringify(licObj, null, 2), 'utf-8');
    const status = getLicenseStatus();
    if (!status.active) {
      return res.status(400).json({ success: false, error: status.message });
    }
    res.json({ success: true, message: `Licencia activada con éxito para ${status.licenciatario}`, licencia: status });
  } catch (e) {
    res.status(400).json({ success: false, error: 'Formato de licencia inválido: ' + e.message });
  }
});

// 1.3 Verificar PIN Maestro del Propietario CESS
app.post('/api/admin/verificar-pin', (req, res) => {
  const { pin } = req.body || {};
  if (pin === MASTER_ADMIN_PIN || pin === 'CESS-2026' || pin === '2026') {
    return res.json({ success: true, authorized: true });
  }
  return res.status(401).json({ success: false, error: 'PIN Maestro incorrecto.' });
});

// 1.4 Bloquear o Suspender Licencia ("Quitarla")
app.post('/api/admin/bloquear-licencia', (req, res) => {
  const { pin, motivo } = req.body || {};
  if (pin !== MASTER_ADMIN_PIN && pin !== 'CESS-2026' && pin !== '2026') {
    return res.status(401).json({ success: false, error: 'No autorizado. PIN Maestro incorrecto.' });
  }

  // Respaldar licencia actual si existe y es válida
  if (fs.existsSync(LICENSE_FILE)) {
    try {
      const current = fs.readFileSync(LICENSE_FILE, 'utf-8');
      fs.writeFileSync(LICENSE_BACKUP_FILE, current, 'utf-8');
    } catch (_) {}
  }

  const revocationData = {
    active: false,
    status: 'REVOCADA',
    motivo: motivo || 'Acceso suspendido manualmente por el Titular / Administrador Maestro CESS.',
    fecha_bloqueo: new Date().toISOString()
  };

  fs.writeFileSync(LICENSE_FILE, JSON.stringify(revocationData, null, 2), 'utf-8');
  const status = getLicenseStatus();
  res.json({ success: true, message: 'Sistema bloqueado. Licencia revocada con éxito.', licencia: status });
});

// 1.5 Reactivar Licencia Vitalicia de Propietario ("Ponerla")
app.post('/api/admin/activar-vitalicia', (req, res) => {
  const { pin } = req.body || {};
  if (pin !== MASTER_ADMIN_PIN && pin !== 'CESS-2026' && pin !== '2026') {
    return res.status(401).json({ success: false, error: 'No autorizado. PIN Maestro incorrecto.' });
  }

  // Si hay backup válido, intentar restaurarlo
  if (fs.existsSync(LICENSE_BACKUP_FILE)) {
    try {
      const backupContent = fs.readFileSync(LICENSE_BACKUP_FILE, 'utf-8');
      const backupObj = JSON.parse(backupContent);
      if (backupObj.payload && backupObj.signature) {
        fs.writeFileSync(LICENSE_FILE, backupContent, 'utf-8');
        const st = getLicenseStatus();
        if (st.active) {
          return res.json({ success: true, message: 'Licencia restaurada desde respaldo anterior.', licencia: st });
        }
      }
    } catch (_) {}
  }

  // Generar licencia de Propietario Vitalicio oficial
  const ownerPayload = {
    producto: "Sistema de Control CESS - Almacén & Mantenimiento",
    version: "2.0.0",
    autor: "Desarrollador / Propietario CESS",
    licenciatario: "LARMONT DIESEL",
    rfc: "N/A",
    plan: "PROPIETARIO_VITALICIO",
    emision: new Date().toISOString().split('T')[0],
    expiracion: null,
    modulos: ["almacen", "entradas", "salidas", "bitacora", "reportes", "ocr_ia"],
    id_licencia: "CESS-MASTER-AUTH"
  };
  const signed = signLicensePayload(ownerPayload);
  fs.writeFileSync(LICENSE_FILE, JSON.stringify(signed, null, 2), 'utf-8');
  res.json({ success: true, message: 'Licencia vitalicia de Propietario activada con éxito.', licencia: getLicenseStatus() });
});

// 1.6 Generar Nueva Licencia Comercial para Clientes
app.post('/api/admin/generar-licencia', (req, res) => {
  const { pin, licenciatario, rfc, plan, diasVigencia, modulos, notas } = req.body || {};
  if (pin !== MASTER_ADMIN_PIN && pin !== 'CESS-2026' && pin !== '2026') {
    return res.status(401).json({ success: false, error: 'No autorizado. PIN Maestro incorrecto.' });
  }

  if (!licenciatario || !licenciatario.trim()) {
    return res.status(400).json({ success: false, error: 'Se requiere el nombre del cliente o empresa.' });
  }

  let expiracion = null;
  const hoy = new Date();
  const planUpper = (plan || 'MENSUAL').toUpperCase();

  if (planUpper === 'DEMO') {
    const exp = new Date(hoy.getTime() + 15 * 24 * 60 * 60 * 1000);
    expiracion = exp.toISOString().split('T')[0];
  } else if (planUpper === 'MENSUAL') {
    const exp = new Date(hoy.getTime() + (Number(diasVigencia) || 30) * 24 * 60 * 60 * 1000);
    expiracion = exp.toISOString().split('T')[0];
  } else if (planUpper === 'TRIMESTRAL') {
    const exp = new Date(hoy.getTime() + (Number(diasVigencia) || 90) * 24 * 60 * 60 * 1000);
    expiracion = exp.toISOString().split('T')[0];
  } else if (planUpper === 'ANUAL') {
    const exp = new Date(hoy.getTime() + (Number(diasVigencia) || 365) * 24 * 60 * 60 * 1000);
    expiracion = exp.toISOString().split('T')[0];
  } else if (planUpper === 'VITALICIO') {
    expiracion = null;
  } else if (diasVigencia && Number(diasVigencia) > 0) {
    const exp = new Date(hoy.getTime() + Number(diasVigencia) * 24 * 60 * 60 * 1000);
    expiracion = exp.toISOString().split('T')[0];
  }

  const idLic = 'CESS-' + Math.random().toString(36).substring(2, 8).toUpperCase() + '-' + Math.random().toString(36).substring(2, 6).toUpperCase();

  const payload = {
    producto: "Sistema de Control CESS - Almacén & Mantenimiento",
    version: "2.0.0",
    autor: "Desarrollador / Propietario CESS",
    licenciatario: licenciatario.trim(),
    rfc: (rfc || 'N/A').trim(),
    plan: planUpper,
    emision: hoy.toISOString().split('T')[0],
    expiracion,
    modulos: Array.isArray(modulos) && modulos.length ? modulos : ["almacen", "entradas", "salidas", "bitacora", "reportes", "ocr_ia"],
    id_licencia: idLic
  };

  const signed = signLicensePayload(payload);

  // Guardar en el registro histórico
  let registry = [];
  if (fs.existsSync(LICENSE_REGISTRY_FILE)) {
    try {
      registry = JSON.parse(fs.readFileSync(LICENSE_REGISTRY_FILE, 'utf-8'));
    } catch (_) {}
  }
  const registroEntry = {
    id_licencia: idLic,
    licenciatario: payload.licenciatario,
    plan: payload.plan,
    emision: payload.emision,
    expiracion: payload.expiracion || 'PERPETUA / SIN VENCIMIENTO',
    license_key: signed.license_key,
    notas: notas || '',
    licencia_completa: signed,
    creado_el: new Date().toISOString()
  };
  registry.unshift(registroEntry);
  fs.writeFileSync(LICENSE_REGISTRY_FILE, JSON.stringify(registry, null, 2), 'utf-8');

  res.json({
    success: true,
    message: `Licencia generada con éxito para ${payload.licenciatario}`,
    licencia: signed,
    license_key: signed.license_key,
    expiracion: payload.expiracion || 'PERPETUA / SIN VENCIMIENTO'
  });
});

// 1.7 Historial de Licencias y Plataformas CESS
app.get('/api/admin/licencias-historial', (req, res) => {
  let registry = [];
  if (fs.existsSync(LICENSE_REGISTRY_FILE)) {
    try {
      registry = JSON.parse(fs.readFileSync(LICENSE_REGISTRY_FILE, 'utf-8'));
    } catch (_) {}
  }
  res.json({
    plataformas_cess: [
      { id: 'almacen_taller', nombre: 'CESS Almacén & Taller Unificado', estado: 'Instalado / Activo', version: '2.0.0', tipo: 'Web / LAN' },
      { id: 'control_rendimientos', nombre: 'CESS Control de Rendimientos Diésel', estado: 'Plataforma Vinculada', version: '2.0.0', tipo: 'Desktop / Python' },
      { id: 'mantenimiento_flota', nombre: 'CESS Mantenimiento & Flotillas', estado: 'Módulo Compatible', version: '2.1.0', tipo: 'Suite CESS' }
    ],
    licencias_generadas: registry
  });
});

// 1.8 Aplicar una licencia del registro a la plataforma actual
app.post('/api/admin/aplicar-licencia-registro', (req, res) => {
  const { pin, id_licencia } = req.body || {};
  if (pin !== MASTER_ADMIN_PIN && pin !== 'CESS-2026' && pin !== '2026') {
    return res.status(401).json({ success: false, error: 'No autorizado. PIN Maestro incorrecto.' });
  }

  if (!fs.existsSync(LICENSE_REGISTRY_FILE)) {
    return res.status(404).json({ success: false, error: 'No hay registro de licencias.' });
  }

  try {
    const reg = JSON.parse(fs.readFileSync(LICENSE_REGISTRY_FILE, 'utf-8'));
    const item = reg.find(r => r.id_licencia === id_licencia);
    if (!item || !item.licencia_completa) {
      return res.status(404).json({ success: false, error: 'Licencia no encontrada en el registro.' });
    }

    fs.writeFileSync(LICENSE_FILE, JSON.stringify(item.licencia_completa, null, 2), 'utf-8');
    const status = getLicenseStatus();
    res.json({ success: true, message: `Licencia aplicada para ${status.licenciatario}`, licencia: status });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// 1.9 Editar y Re-firmar Licencia Activa (Licenciatario, RFC, Plan, Vencimiento)
app.post('/api/admin/editar-licencia-activa', (req, res) => {
  const { pin, licenciatario, rfc, plan, expiracion, modulos } = req.body || {};
  if (pin !== MASTER_ADMIN_PIN && pin !== 'CESS-2026' && pin !== '2026') {
    return res.status(401).json({ success: false, error: 'No autorizado. PIN Maestro CESS incorrecto.' });
  }

  if (!licenciatario || !licenciatario.trim()) {
    return res.status(400).json({ success: false, error: 'Se requiere el nombre o razón social del cliente/licenciatario.' });
  }

  try {
    let currentLic = {};
    if (fs.existsSync(LICENSE_FILE)) {
      try {
        currentLic = JSON.parse(fs.readFileSync(LICENSE_FILE, 'utf-8'));
      } catch (_) {}
    }

    const currentPayload = currentLic.payload || {};
    const hoy = new Date().toISOString().split('T')[0];
    const planUpper = (plan || currentPayload.plan || 'VITALICIO').toUpperCase();

    let finalExpiracion = expiracion !== undefined ? expiracion : currentPayload.expiracion;
    if (
      !finalExpiracion ||
      finalExpiracion === 'PERPETUA' ||
      finalExpiracion === 'VITALICIO' ||
      finalExpiracion === 'PERPETUA / SIN VENCIMIENTO' ||
      planUpper === 'VITALICIO' ||
      planUpper === 'PROPIETARIO_VITALICIO'
    ) {
      finalExpiracion = null;
    }

    const payload = {
      producto: "Sistema de Control CESS - Almacén & Mantenimiento",
      version: "2.0.0",
      autor: "Desarrollador / Propietario CESS",
      licenciatario: licenciatario.trim(),
      rfc: (rfc || currentPayload.rfc || 'N/A').trim(),
      plan: planUpper,
      emision: currentPayload.emision || hoy,
      expiracion: finalExpiracion,
      modulos: Array.isArray(modulos) && modulos.length ? modulos : (currentPayload.modulos || ["almacen", "entradas", "salidas", "bitacora", "reportes", "ocr_ia", "auditoria", "compras"]),
      id_licencia: currentPayload.id_licencia || ('CESS-' + Math.random().toString(36).substring(2, 8).toUpperCase() + '-' + Math.random().toString(36).substring(2, 6).toUpperCase())
    };

    const signed = signLicensePayload(payload);
    fs.writeFileSync(LICENSE_FILE, JSON.stringify(signed, null, 2), 'utf-8');

    // Actualizar también en el registro histórico
    let registry = [];
    if (fs.existsSync(LICENSE_REGISTRY_FILE)) {
      try {
        registry = JSON.parse(fs.readFileSync(LICENSE_REGISTRY_FILE, 'utf-8'));
      } catch (_) {}
    }
    const regIdx = registry.findIndex(r => r.id_licencia === payload.id_licencia);
    const regEntry = {
      id_licencia: payload.id_licencia,
      licenciatario: payload.licenciatario,
      plan: payload.plan,
      emision: payload.emision,
      expiracion: payload.expiracion || 'PERPETUA / SIN VENCIMIENTO',
      license_key: signed.license_key,
      notas: 'Actualizado directamente desde Centro de Control Maestro CESS',
      licencia_completa: signed,
      actualizado_el: new Date().toISOString()
    };
    if (regIdx >= 0) {
      registry[regIdx] = { ...registry[regIdx], ...regEntry };
    } else {
      registry.unshift(regEntry);
    }
    fs.writeFileSync(LICENSE_REGISTRY_FILE, JSON.stringify(registry, null, 2), 'utf-8');

    // Sincronizar razón social / empresa en usuarios de este sistema
    if (Array.isArray(store.usuarios)) {
      store.usuarios.forEach(u => {
        if (!u.empresa || u.empresa === 'LARMONT DIESEL' || u.empresa === currentPayload.licenciatario) {
          u.empresa = payload.licenciatario;
        }
      });
      saveStore();
    }

    const updatedStatus = getLicenseStatus();
    res.json({
      success: true,
      message: `Licencia actualizada y re-firmada exitosamente para "${payload.licenciatario}"`,
      licencia: updatedStatus
    });
  } catch (err) {
    res.status(500).json({ success: false, error: 'Error al actualizar licencia: ' + err.message });
  }
});

// =============================================================
// GESTIÓN DE USUARIOS Y CONTROL DE ACCESOS (CENTRO DE CONTROL)
// =============================================================

// Listar usuarios públicos para selector de perfil
app.get('/api/usuarios', (req, res) => {
  const currentLic = getLicenseStatus();
  const defaultEmpresa = currentLic?.licenciatario || 'LARMONT DIESEL';
  const users = (store.usuarios || [])
    .filter(u => u.activo !== false)
    .map(u => ({
      id: u.id,
      nombre: u.nombre,
      rol: u.rol,
      empresa: u.empresa || defaultEmpresa,
      permisos: u.permisos
    }));
  res.json({ success: true, usuarios: users });
});

// Inicio de sesión / cambio de usuario mediante clave alfanumérica o PIN
app.post('/api/login-usuario', (req, res) => {
  const { pin, userId, username } = req.body || {};
  if (!pin) {
    return res.status(400).json({ success: false, error: 'Ingresa tu clave de acceso o contraseña' });
  }

  const cleanPin = String(pin).trim().toUpperCase();
  const activeUsers = (store.usuarios || []).filter(u => u.activo !== false);

  let user = null;
  if (userId) {
    user = activeUsers.find(
      u => u.id === userId && String(u.pin).trim().toUpperCase() === cleanPin
    );
  } else if (username) {
    const cleanUser = String(username).trim().toUpperCase();
    user = activeUsers.find(
      u => (u.nombre.toUpperCase().includes(cleanUser) || u.id.toUpperCase() === cleanUser) &&
           String(u.pin).trim().toUpperCase() === cleanPin
    );
  } else {
    user = activeUsers.find(
      u => String(u.pin).trim().toUpperCase() === cleanPin
    );
  }

  if (!user) {
    if (userId || username) {
      return res.status(401).json({ success: false, error: 'Contraseña o clave de acceso incorrecta para el usuario' });
    }
    return res.status(401).json({ success: false, error: 'Clave de acceso no reconocida o usuario inactivo' });
  }

  const currentLic = getLicenseStatus();
  const defaultEmpresa = currentLic?.licenciatario || 'LARMONT DIESEL';

  res.json({
    success: true,
    usuario: {
      id: user.id,
      nombre: user.nombre,
      rol: user.rol,
      empresa: user.empresa || defaultEmpresa,
      permisos: user.permisos,
      creado_en: user.creado_en
    }
  });
});

// Obtener usuarios completos para el Centro de Control Maestro (César)
app.get('/api/admin/usuarios', (req, res) => {
  const authHeader = req.headers['x-master-pin'] || req.query.masterPin;
  if (authHeader !== MASTER_ADMIN_PIN && authHeader !== 'CESS-2026' && authHeader !== '2026') {
    return res.status(403).json({ error: 'Acceso no autorizado al Centro de Control Maestro' });
  }
  const currentLic = getLicenseStatus();
  const defaultEmpresa = currentLic?.licenciatario || 'LARMONT DIESEL';
  const usuarios = (store.usuarios || []).map(u => ({
    ...u,
    empresa: u.empresa || defaultEmpresa
  }));
  res.json({ success: true, usuarios });
});

// Crear, editar, alternar estado o eliminar usuarios desde el Centro de Control
app.post('/api/admin/usuarios', (req, res) => {
  const { masterPin, action, user } = req.body || {};
  if (masterPin !== MASTER_ADMIN_PIN && masterPin !== 'CESS-2026' && masterPin !== '2026') {
    return res.status(403).json({ error: 'PIN Maestro CESS inválido' });
  }

  if (!Array.isArray(store.usuarios)) store.usuarios = [];
  const currentLic = getLicenseStatus();
  const defaultEmpresa = currentLic?.licenciatario || 'LARMONT DIESEL';

  if (action === 'create') {
    if (!user || !user.nombre || !user.pin) {
      return res.status(400).json({ error: 'Nombre y Clave de Acceso son obligatorios' });
    }
    const cleanPin = String(user.pin).trim();
    if (store.usuarios.some(u => String(u.pin).trim().toUpperCase() === cleanPin.toUpperCase())) {
      return res.status(400).json({ error: 'Esa clave de acceso ya está asignada a otro usuario. Elige una clave diferente.' });
    }

    const rol = user.rol || 'GERENTE';
    const newUser = {
      id: `usr_${Date.now()}`,
      nombre: String(user.nombre).trim(),
      pin: cleanPin,
      rol: rol,
      empresa: (user.empresa && String(user.empresa).trim()) || defaultEmpresa,
      activo: true,
      permisos: user.permisos || {
        ver_inventario: true,
        dar_salida: rol !== 'GERENTE',
        dar_entrada: rol !== 'GERENTE' && rol !== 'TALLER',
        ver_bitacora: true,
        exportar_excel: rol === 'ADMIN' || rol === 'GERENTE',
        cuadre_inventario: rol === 'ADMIN',
        compras: rol === 'ADMIN' || rol === 'GERENTE',
        admin_centro_control: rol === 'ADMIN'
      },
      creado_en: new Date().toISOString()
    };
    store.usuarios.push(newUser);
    saveStore();
    return res.json({ success: true, message: `Usuario "${newUser.nombre}" vinculado a "${newUser.empresa}" creado exitosamente`, usuarios: store.usuarios });
  }

  if (action === 'update') {
    if (!user || !user.id) return res.status(400).json({ error: 'ID de usuario requerido' });
    const idx = store.usuarios.findIndex(u => u.id === user.id);
    if (idx === -1) return res.status(404).json({ error: 'Usuario no encontrado' });

    if (user.pin) {
      const cleanPin = String(user.pin).trim();
      const collision = store.usuarios.find(u => u.id !== user.id && String(u.pin).trim().toUpperCase() === cleanPin.toUpperCase());
      if (collision) {
        return res.status(400).json({ error: `Esa clave de acceso ya está asignada a "${collision.nombre}"` });
      }
    }

    store.usuarios[idx] = {
      ...store.usuarios[idx],
      nombre: user.nombre ? String(user.nombre).trim() : store.usuarios[idx].nombre,
      pin: user.pin ? String(user.pin).trim() : store.usuarios[idx].pin,
      rol: user.rol || store.usuarios[idx].rol,
      empresa: user.empresa !== undefined ? String(user.empresa).trim() : (store.usuarios[idx].empresa || defaultEmpresa),
      permisos: user.permisos ? { ...store.usuarios[idx].permisos, ...user.permisos } : store.usuarios[idx].permisos
    };
    saveStore();
    return res.json({ success: true, message: `Usuario "${store.usuarios[idx].nombre}" actualizado correctamente`, usuarios: store.usuarios });
  }

  if (action === 'toggle') {
    const target = store.usuarios.find(u => u.id === user.id);
    if (!target) return res.status(404).json({ error: 'Usuario no encontrado' });
    if (target.id === 'usr_cesar') {
      return res.status(400).json({ error: 'No se puede desactivar la cuenta principal de Administrador' });
    }
    target.activo = !target.activo;
    saveStore();
    return res.json({ success: true, message: `Usuario ${target.activo ? 'activado' : 'suspendido'}`, usuarios: store.usuarios });
  }

  if (action === 'delete') {
    if (user.id === 'usr_cesar') {
      return res.status(400).json({ error: 'No se puede eliminar la cuenta principal de Administrador' });
    }
    store.usuarios = store.usuarios.filter(u => u.id !== user.id);
    saveStore();
    return res.json({ success: true, message: 'Usuario eliminado del sistema', usuarios: store.usuarios });
  }

  res.status(400).json({ error: 'Acción inválida' });
});

// 2. Obtener lista de hojas en INVENTARIO.xlsx
app.get('/api/sheets', (req, res) => {
  if (!fs.existsSync(EXCEL_FILE)) {
    return res.status(404).json({ error: 'Archivo INVENTARIO.xlsx no encontrado' });
  }
  const wb = xlsx.readFile(EXCEL_FILE, { bookSheets: true });
  res.json({ sheets: wb.SheetNames, activeSheet: store.activeSheet });
});

// 2.1 Obtener lista de Almacenes con métricas y existencias
app.get('/api/almacenes', (req, res) => {
  ensureAlmacenes();
  const stats = store.almacenes.map(alm => {
    const items = Object.values(store.catalog).filter(p => p.almacen_id === alm.id || p.almacen === alm.nombre);
    const totalItems = items.length;
    const valorTotal = items.reduce((sum, p) => sum + ((Number(p.stock_actual) || 0) * (Number(p.precio_unitario) || 0)), 0);
    const stockTotal = items.reduce((sum, p) => sum + (Number(p.stock_actual) || 0), 0);
    const agotados = items.filter(p => (Number(p.stock_actual) || 0) <= 0).length;
    return {
      ...alm,
      totalItems,
      valorTotal,
      stockTotal,
      agotados
    };
  });
  res.json({ almacenes: stats });
});

// 2.2 Crear nuevo Almacén / Departamento Autónomo
app.post('/api/almacenes', (req, res) => {
  const { nombre, codigo, icono, color, descripcion } = req.body;
  if (!nombre || !nombre.trim()) {
    return res.status(400).json({ error: 'El nombre del almacén es obligatorio.' });
  }

  ensureAlmacenes();
  const id = nombre.toLowerCase().trim().replace(/[^a-z0-9]/g, '_');
  if (store.almacenes.some(a => a.id === id)) {
    return res.status(400).json({ error: 'Ya existe un almacén con este identificador o nombre.' });
  }

  const cod = (codigo || nombre.slice(0, 4)).trim().toUpperCase();

  const nuevoAlm = {
    id,
    nombre: nombre.trim(),
    codigo: cod,
    icono: icono || 'Boxes',
    color: color || '#2563EB',
    descripcion: descripcion ? descripcion.trim() : ''
  };

  store.almacenes.push(nuevoAlm);
  saveStore();

  res.json({
    success: true,
    almacen: nuevoAlm,
    almacenes: store.almacenes,
    message: `Almacén '${nombre}' creado con éxito en la base de datos.`
  });
});

// 2.3 Editar Almacén Autónomo
app.put('/api/almacenes/:id', (req, res) => {
  const { id } = req.params;
  const { nombre, codigo, icono, color, descripcion } = req.body;
  ensureAlmacenes();

  const alm = store.almacenes.find(a => a.id === id);
  if (!alm) return res.status(404).json({ error: 'Almacén no encontrado' });

  const oldNombre = alm.nombre;
  if (nombre && nombre.trim()) alm.nombre = nombre.trim();
  if (codigo && codigo.trim()) alm.codigo = codigo.trim().toUpperCase();
  if (icono) alm.icono = icono;
  if (color) alm.color = color;
  if (descripcion !== undefined) alm.descripcion = descripcion.trim();

  // Si cambió el nombre, actualizar productos asociados y bitácora
  if (nombre && oldNombre !== nombre.trim()) {
    for (const item of Object.values(store.catalog)) {
      if (item.almacen_id === id || item.almacen === oldNombre) {
        item.almacen = nombre.trim();
        item.almacen_id = id;
      }
    }
    for (const m of (store.movimientos || [])) {
      if (m.almacen_id === id || m.almacen === oldNombre) {
        m.almacen = nombre.trim();
      }
    }
  }

  saveStore();
  res.json({ success: true, almacen: alm, almacenes: store.almacenes, message: `Almacén '${alm.nombre}' actualizado con éxito.` });
});

// 2.4 Eliminar Almacén (Reasigna productos a Refacciones Diesel)
app.delete('/api/almacenes/:id', (req, res) => {
  const { id } = req.params;
  if (id === 'refacciones') {
    return res.status(400).json({ error: 'No se puede eliminar el almacén principal de Refacciones.' });
  }

  ensureAlmacenes();
  const idx = store.almacenes.findIndex(a => a.id === id);
  if (idx === -1) return res.status(404).json({ error: 'Almacén no encontrado' });

  // Reasignar productos huérfanos a Almacén de Refacciones
  for (const item of Object.values(store.catalog)) {
    if (item.almacen_id === id) {
      item.almacen_id = 'refacciones';
      item.almacen = 'Almacén de Refacciones';
    }
  }

  const eliminado = store.almacenes.splice(idx, 1)[0];
  saveStore();
  res.json({ success: true, message: `Almacén '${eliminado.nombre}' eliminado. Los productos fueron reasignados a Almacén de Refacciones.` });
});

// 3. Forzar sincronización desde Excel (Multi-Almacén o Hoja específica)
app.post('/api/sync-excel', (req, res) => {
  try {
    const { sheetName, syncMode } = req.body || {};
    const result = syncFromExcel(sheetName, syncMode || (sheetName ? 'SINGLE' : 'ALL'));
    res.json({
      success: true,
      message: `Sincronizado con éxito desde Excel (${result.count} productos cargados de ${result.almacenes.length} almacenes).`,
      ...result
    });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// 4. Catálogo de inventario con existencias y almacenes
app.get('/api/inventory', (req, res) => {
  ensureAlmacenes();
  const items = Object.values(store.catalog);
  res.json({
    total: items.length,
    activeSheet: store.activeSheet,
    items,
    almacenes: store.almacenes
  });
});

// Stopwords y normalización morfológica (lematización/stemming en español) para refacciones diesel
const STOPWORDS = new Set(['de', 'la', 'el', 'los', 'las', 'para', 'con', 'en', 'un', 'una', 'por', 'del', 'al', 'y', 'o', 'que', 'par', 'pares', 'sin', 'se', 'su', 'sus', 'pza', 'pzas', 'pz', 'jgo', 'juego', 'c/u', 'cada', 'uno']);

const PHYSICAL_PART_KEYWORDS = [
  'filtro', 'filtros', 'aceite', 'lubricante', 'grasa', 'reten', 'balero', 'rodamiento',
  'sensor', 'disco', 'campana', 'tambor', 'banda', 'correa', 'manguera', 'valvula',
  'empaque', 'junta', 'cinta', 'terminal', 'cruceta', 'buje', 'amortiguador',
  'zapata', 'freno', 'balata', 'tornillo', 'tuerca', 'arandela', 'huasa', 'rondana',
  'foco', 'lampara', 'llanta', 'neumatico', 'camara', 'bateria', 'acumulador',
  'relevador', 'fusible', 'abrazadera', 'conector', 'piston', 'camisa', 'anillo',
  'bomba', 'termostato', 'polea', 'tensor', 'rotula', 'varilla', 'resorte', 'donaldson',
  'fleetguard', 'baldwin', 'wix', 'mann', 'caterpillar', 'cummins', 'detroit', 'volvo'
];

const SERVICE_KEYWORDS = [
  'rimar', 'rima', 'torno', 'rectificar', 'rectificado', 'mandrinado', 'barrenado',
  'soldar', 'soldadura', 'alineacion', 'balanceo', 'vulcanizado', 'parche',
  'mano de obra', 'maquila', 'maquinado', 'trabajo externo', 'calibracion de inyectores',
  'embobinado', 'servicio de torno', 'servicio externo', 'reparacion externa'
];

function isServiceDescription(desc) {
  if (!desc) return false;
  const d = String(desc).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  // Si contiene explícitamente palabras de refacciones o insumos físicos de almacén, es una pieza física, NO un servicio
  if (PHYSICAL_PART_KEYWORDS.some(p => new RegExp(`\\b${p}\\b`, 'i').test(d))) {
    return false;
  }
  return SERVICE_KEYWORDS.some(k => {
    const regex = new RegExp(`\\b${k}\\b`, 'i');
    return regex.test(d) || d.startsWith(k);
  });
}

function stemWord(w) {
  if (!w || typeof w !== 'string') return '';
  let s = w.toLowerCase().trim().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  if (s.endsWith('ces') && s.length > 4) s = s.slice(0, -3) + 'z';
  else if (s.endsWith('es') && s.length > 4) s = s.slice(0, -2);
  else if (s.endsWith('s') && s.length > 3 && !s.endsWith('is') && !s.endsWith('us')) s = s.slice(0, -1);
  return s;
}

// Diccionario de sinónimos, apodos y variantes comunes de taller mecánico diésel
const WORKSHOP_SYNONYMS = [
  {
    targetPart: 'TI6008',
    keys: ['cinta gris', 'cinta ducto', 'cinta de ducto', 'cinta para ducto', 'cinta para ductos', 'cinta gris para ductos', 'cinta ductos', 'duct tape', 'cinta plateada', 'tape gris', 'cinta gris ducto']
  },
  {
    targetPart: '165',
    keys: ['cinta de aislar', 'cinta aislar', 'cinta aislante', 'cinta electrica', 'cinta negra', 'cinta 3m', 'cinta aisladora', 'cinta de aislar 3m']
  },
  {
    targetPart: '27010',
    keys: ['cinta de teflon', 'cinta teflon', 'teflon', 'cinta blanca', 'cinta para roscas', 'cinta selladora', 'teflon pretul']
  }
];

// Helper robusto para buscar refacción por ID, número de parte, equivalencias cruzadas o alias
function isServicioExterno(text) {
  if (!text) return false;
  const t = String(text).toUpperCase();
  const kw = ['RIMA', 'RIMADO', 'RIMAR', 'TORNO', 'TORNEADO', 'TORNEAR', 'RECTIFICADO', 'RECTIFICAR', 'MAQUINADO', 'MAQUINAR', 'SOLDADURA', 'SOLDAR', 'EMBOCINADO', 'EMBOCINAR', 'ALINEACION', 'BALANCEO', 'SERVICIO EXTERNO', 'MANO DE OBRA', 'FLETE', 'MANTENIMIENTO EXTERNO'];
  return kw.some(k => t.includes(k));
}

function findPart(idOrCode) {
  if (!idOrCode) return null;
  const raw = String(idOrCode).trim();
  const q = raw.toLowerCase();
  const qClean = raw.toUpperCase().replace(/[^A-Z0-9]/g, '');

  // Códigos genéricos NUNCA deben asociar una pieza única ni acumular inventario cruzado
  const genericCodes = ['s/n', 'sn', 'n/a', 'na', 'sin numero', 'sin parte', 'sin codigo', 'generica', 'generico', 'pza', 'pz', 'pieza', 'refaccion', ''];
  if (genericCodes.includes(q)) return null;

  if (store.catalog[raw]) return store.catalog[raw];
  if (store.catalog[raw.toUpperCase()]) return store.catalog[raw.toUpperCase()];

  const allItems = Object.values(store.catalog);

  // 1. Coincidencia exacta por ID o Número de Parte (excluyendo números genéricos)
  let exact = allItems.find(p => {
    if (p.id && p.id.toLowerCase() === q) return true;
    if (p.numero_parte && p.numero_parte.toLowerCase() === q) {
      const pNumLow = String(p.numero_parte).toLowerCase();
      if (genericCodes.includes(pNumLow)) return false;
      return true;
    }
    return false;
  });
  if (exact) return exact;

  // 2. Coincidencia normalizada (sin guiones, espacios ni barras)
  if (qClean.length >= 3) {
    let norm = allItems.find(p => {
      const pIdClean = (p.id || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
      if (pIdClean && pIdClean === qClean) return true;
      const pParteClean = (p.numero_parte || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
      if (pParteClean && pParteClean === qClean) return true;
      return false;
    });
    if (norm) return norm;
  }

  // 3. Coincidencia en Equivalencias (exacta o normalizada)
  let eqMatch = allItems.find(p => {
    if (!Array.isArray(p.equivalencias)) return false;
    return p.equivalencias.some(eq => {
      const eqRaw = String(eq || '').trim();
      if (eqRaw.toLowerCase() === q) return true;
      if (qClean.length >= 3) {
        const eqClean = eqRaw.toUpperCase().replace(/[^A-Z0-9]/g, '');
        if (eqClean === qClean) return true;
      }
      return false;
    });
  });
  if (eqMatch) return eqMatch;

  // 4. Coincidencia en Alias registrados (exacta o normalizada)
  let aliasMatch = allItems.find(p => {
    if (!Array.isArray(p.alias)) return false;
    return p.alias.some(a => {
      const aRaw = String(a || '').trim();
      if (aRaw.toLowerCase() === q) return true;
      if (qClean.length >= 3) {
        const aClean = aRaw.toUpperCase().replace(/[^A-Z0-9]/g, '');
        if (aClean === qClean) return true;
      }
      return false;
    });
  });
  if (aliasMatch) return aliasMatch;

  return null;
}

function normalizeSearchText(txt) {
  if (!txt) return '';
  return String(txt)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();
}

function matchPartMultiWord(p, query) {
  if (!p) return false;
  if (!query || !query.trim()) return true;
  const rawTokens = normalizeSearchText(query).split(/\s+/).filter(Boolean);
  if (rawTokens.length === 0) return true;

  const fields = [
    p.numero_parte,
    p.descripcion,
    p.marca,
    p.rack,
    p.nivel,
    p.almacen,
    ...(Array.isArray(p.equivalencias) ? p.equivalencias : []),
    ...(Array.isArray(p.alias) ? p.alias : [])
  ].filter(Boolean).map(normalizeSearchText);

  const fullText = fields.join(' ');
  const cleanFullText = fullText.replace(/[^a-z0-9]/g, '');

  return rawTokens.every(tok => {
    if (fullText.includes(tok)) return true;
    const cleanTok = tok.replace(/[^a-z0-9]/g, '');
    if (cleanTok && cleanFullText.includes(cleanTok)) return true;
    return false;
  });
}

// =========================================================
// MOTOR DE APRENDIZAJE Y MEMORIA PERMANENTE IA (CESS ALMACÉN)
// =========================================================
function getMemoriaIAPromptContext() {
  if (!Array.isArray(store.memoria_ia) || store.memoria_ia.length === 0) return "";
  const topMem = store.memoria_ia.slice(0, 30).map(m => {
    const raw = m.codigo_extraido || m.texto_original || "";
    const cat = (m.numero_parte_catalogo || "") + " (" + (m.descripcion_catalogo || "") + ")";
    const marca = m.marca_catalogo ? (" [" + m.marca_catalogo + "]") : "";
    return "- \"" + raw + "\" ➔ Corresponde a: \"" + cat + "\"" + marca;
  }).join("\n");
  return "\nREGLAS DE APRENDIZAJE Y ASOCIACIÓN PREVIA RECORDADAS POR EL ALMACÉN (Usa estas asociaciones si detectas textos o piezas afines):\n" + topMem + "\n";
}
function buscarEnMemoriaIA(numParteRaw, descRaw, marcaRaw, proveedorRaw) {
  if (!Array.isArray(store.memoria_ia) || store.memoria_ia.length === 0) return null;

  const numParteClean = String(numParteRaw || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  const descNorm = normalizeSearchText(descRaw || '');

  // 1. Coincidencia por código extraído memorizado
  if (numParteClean && numParteClean.length >= 3) {
    const memCod = store.memoria_ia.find(m => {
      const mCodClean = String(m.codigo_extraido || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
      return mCodClean && mCodClean === numParteClean;
    });
    if (memCod && store.catalog[memCod.id_pieza_catalogo]) {
      return {
        match: store.catalog[memCod.id_pieza_catalogo],
        tipo: 'MEMORIA_IA_CODIGO',
        confianza: 'ALTA',
        detalle_aprendizaje: `Aprendido por corrección previa del usuario (${memCod.fecha_registro || ''})`
      };
    }
  }

  // 2. Coincidencia por texto completo memorizado
  if (descNorm && descNorm.length >= 4) {
    // 2.1 Coincidencia exacta o contenida
    const memExact = store.memoria_ia.find(m => {
      const mText = normalizeSearchText(m.texto_original || m.texto_normalizado || '');
      return mText && (mText === descNorm || descNorm.includes(mText) || mText.includes(descNorm));
    });
    if (memExact && store.catalog[memExact.id_pieza_catalogo]) {
      return {
        match: store.catalog[memExact.id_pieza_catalogo],
        tipo: 'MEMORIA_IA_TEXTO',
        confianza: 'ALTA',
        detalle_aprendizaje: `Aprendido por corrección previa del usuario ("${memExact.texto_original}")`
      };
    }

    // 2.2 Coincidencia multi-palabra de alta similitud (todas las palabras significativas coinciden)
    const descWords = descNorm.split(/\s+/).filter(w => w.length >= 3 && !['PARA', 'CON', 'SIN', 'DEL', 'POR', 'LOS', 'LAS', 'REFACCION', 'PIEZA'].includes(w));
    if (descWords.length >= 2) {
      for (const m of store.memoria_ia) {
        const mWords = normalizeSearchText(m.texto_original || '').split(/\s+/).filter(Boolean);
        const matchesAll = descWords.every(w => mWords.some(mw => mw.includes(w) || w.includes(mw)));
        if (matchesAll && store.catalog[m.id_pieza_catalogo]) {
          return {
            match: store.catalog[m.id_pieza_catalogo],
            tipo: 'MEMORIA_IA_SIMILITUD',
            confianza: 'ALTA',
            detalle_aprendizaje: `Asociación inteligente aprendida ("${m.texto_original}")`
          };
        }
      }
    }
  }

  return null;
}

function registrarAprendizajeIA({
  texto_detectado,
  codigo_detectado,
  marca_detectada,
  proveedor,
  id_pieza_catalogo,
  numero_parte_catalogo,
  origen = 'correccion_usuario'
}) {
  if (!id_pieza_catalogo && !numero_parte_catalogo) return null;
  if (!Array.isArray(store.memoria_ia)) store.memoria_ia = [];

  let targetItem = id_pieza_catalogo ? store.catalog[id_pieza_catalogo] : null;
  if (!targetItem && numero_parte_catalogo) {
    targetItem = Object.values(store.catalog).find(p => p.numero_parte === numero_parte_catalogo || p.id === numero_parte_catalogo);
  }
  if (!targetItem) return null;

  const rawText = String(texto_detectado || '').trim();
  const rawCode = String(codigo_detectado || '').trim().toUpperCase();
  const normText = normalizeSearchText(rawText);

  // Buscar si ya existe una regla similar
  let existingIndex = store.memoria_ia.findIndex(m => {
    const sameCode = rawCode && rawCode.length >= 3 && String(m.codigo_extraido || '').toUpperCase() === rawCode;
    const sameText = normText && normText.length >= 4 && normalizeSearchText(m.texto_original || '') === normText;
    return (sameCode || sameText) && m.id_pieza_catalogo === targetItem.id;
  });

  const record = {
    id: existingIndex !== -1 ? store.memoria_ia[existingIndex].id : `mem_${Date.now()}_${Math.floor(Math.random() * 1000)}`,
    texto_original: rawText || targetItem.descripcion,
    texto_normalizado: normText,
    codigo_extraido: rawCode || targetItem.numero_parte,
    marca_detectada: String(marca_detectada || targetItem.marca || '').toUpperCase(),
    proveedor: String(proveedor || '').toUpperCase(),
    id_pieza_catalogo: targetItem.id,
    numero_parte_catalogo: targetItem.numero_parte,
    descripcion_catalogo: targetItem.descripcion,
    marca_catalogo: targetItem.marca,
    rack: targetItem.rack,
    nivel: targetItem.nivel,
    origen,
    veces_utilizado: existingIndex !== -1 ? (store.memoria_ia[existingIndex].veces_utilizado || 1) + 1 : 1,
    fecha_registro: new Date().toISOString().slice(0, 10),
    timestamp: new Date().toISOString()
  };

  if (existingIndex !== -1) {
    store.memoria_ia[existingIndex] = record;
  } else {
    store.memoria_ia.unshift(record);
  }

  // Enriquecer el catálogo para búsquedas rápidas (equivalencias y alias)
  if (rawCode && rawCode.length >= 3 && rawCode !== targetItem.numero_parte) {
    if (!Array.isArray(targetItem.equivalencias)) targetItem.equivalencias = [];
    if (!targetItem.equivalencias.includes(rawCode)) targetItem.equivalencias.push(rawCode);
  }
  if (rawText && rawText.length >= 4 && rawText.toUpperCase() !== targetItem.descripcion.toUpperCase()) {
    if (!Array.isArray(targetItem.alias)) targetItem.alias = [];
    if (!targetItem.alias.includes(rawText)) targetItem.alias.push(rawText);
  }
  targetItem.aprendizaje_ia = (targetItem.aprendizaje_ia || 0) + 1;

  saveStore();
  console.log(`[IA Aprendizaje] Record memorizado: "${rawText || rawCode}" ➔ ${targetItem.numero_parte} (${targetItem.descripcion})`);
  return record;
}

// Helper inteligente de coincidencia contra catálogo con lematización, sinónimos de taller, multi-palabra y ranking difuso
function matchItemToCatalog(numParteRaw, descRaw, marcaRaw, proveedorRaw, allowServices = true) {
  // LAYER 0: CONSULTAR MEMORIA PERMANENTE DE APRENDIZAJE IA
  const learnedMatch = buscarEnMemoriaIA(numParteRaw, descRaw, marcaRaw, proveedorRaw);
  if (learnedMatch) return learnedMatch;

  // DETECCIÓN ESPECIAL DE SERVICIOS EXTERNOS (Torno, Rectificado, Rima, Mano de Obra, etc.)
  // Sólo si allowServices es true
  const isSrv = allowServices ? isServiceDescription(descRaw) : false;
  if (isSrv) {
    // Si es un servicio, buscar ÚNICAMENTE en el catálogo de servicios externos para no confundir con refacciones físicas
    const allServices = Object.values(store.catalog).filter(p => p.almacen_id === 'servicios_externos' || p.tipo === 'SERVICIO');
    const srvMatch = allServices.find(p => {
      const pDesc = (p.descripcion || '').toLowerCase();
      const dRaw = String(descRaw || '').toLowerCase();
      return pDesc.includes(dRaw) || dRaw.includes(pDesc);
    });
    if (srvMatch) {
      return { match: srvMatch, es_servicio: true, tipo: 'SERVICIO_CATALOGO', confianza: 'ALTA' };
    }
    // Si no está registrado en el catálogo de servicios, devolver como nuevo servicio externo
    return {
      match: null,
      es_servicio: true,
      sugerencia_almacen: 'servicios_externos',
      sugerencia_codigo: 'SRV-' + Math.floor(100 + Math.random() * 900),
      tipo: 'NUEVO_SERVICIO_EXTERNO',
      confianza: 'ALTA'
    };
  }

  const numParte = String(numParteRaw || '').trim().toUpperCase();
  const numParteClean = numParte.replace(/[^A-Z0-9]/g, '');
  const desc = String(descRaw || '').trim().toLowerCase();
  // IMPORTANTÍSIMO: allItems para refacciones físicas NUNCA debe incluir servicios externos ni dummies 'SERVICIO'
  const allItems = Object.values(store.catalog).filter(p => p.almacen_id !== 'servicios_externos' && p.numero_parte !== 'SERVICIO' && p.tipo !== 'SERVICIO');

  const isGenericNum = !numParte || ['S/N', 'SN', 'N/A', 'NA', 'SIN NUMERO', 'SIN CODIGO', 'GENERICA', 'GENERICO', 'SERVICIO'].includes(numParte);

  // 1. Coincidencia exacta por número de parte o ID
  if (!isGenericNum && numParte) {
    let exact = allItems.find(p => (p.numero_parte && p.numero_parte.toUpperCase() === numParte) || (p.id && p.id.toUpperCase() === numParte));
    if (exact) return { match: exact, tipo: 'EXACTA', confianza: 'ALTA' };

    // 2. Coincidencia normalizada (sin guiones o espacios)
    if (numParteClean.length >= 3) {
      let cleanMatch = allItems.find(p => {
        const pClean = (p.numero_parte || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
        const pIdClean = (p.id || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
        return pClean === numParteClean || pIdClean === numParteClean;
      });
      if (cleanMatch) return { match: cleanMatch, tipo: 'NORMALIZADA', confianza: 'ALTA' };
    }

    // 3. Coincidencia por equivalencias cruzadas de otras marcas
    let eqMatch = allItems.find(p =>
      Array.isArray(p.equivalencias) &&
      p.equivalencias.some(eq => {
        const eqUp = String(eq || '').toUpperCase();
        return eqUp === numParte || (numParteClean.length >= 3 && eqUp.replace(/[^A-Z0-9]/g, '') === numParteClean);
      })
    );
    if (eqMatch) return { match: eqMatch, tipo: 'EQUIVALENCIA', confianza: 'ALTA' };

    // 3.1 Coincidencia por alias registrados
    let aliasMatch = allItems.find(p =>
      Array.isArray(p.alias) &&
      p.alias.some(a => {
        const aUp = String(a || '').toUpperCase();
        return aUp === numParte || (numParteClean.length >= 3 && aUp.replace(/[^A-Z0-9]/g, '') === numParteClean);
      })
    );
    if (aliasMatch) return { match: aliasMatch, tipo: 'ALIAS', confianza: 'ALTA' };

    // 3.2 Coincidencia por prefijo o sufijo de proveedor (ej. FLT-LF9009 o LF9009-B)
    if (numParteClean.length >= 4) {
      let subMatch = allItems.find(p => {
        const pClean = (p.numero_parte || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
        return pClean.length >= 4 && (numParteClean.includes(pClean) || pClean.includes(numParteClean));
      });
      if (subMatch) return { match: subMatch, tipo: 'CODIGO_PROVEEDOR', confianza: 'ALTA' };
    }
  }

  // 4. Códigos de parte incrustados en la descripción (ej. "FILTRO COMBUSTIBLE FS20385")
  const fullTextToInspect = `${descRaw || ''} ${numParteRaw || ''}`;
  const candidateCodes = fullTextToInspect.match(/[A-Z0-9-]{3,}/gi) || [];
  for (const c of candidateCodes) {
    const cClean = c.toUpperCase().replace(/[^A-Z0-9]/g, '');
    if (cClean.length >= 4 && !['FILTRO', 'VALVULA', 'RETEN', 'ACEITE', 'DIESEL', 'PIEZA', 'REFACCION'].includes(cClean)) {
      const matchFound = allItems.find(p => {
        const pClean = (p.numero_parte || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
        return pClean === cClean || (Array.isArray(p.equivalencias) && p.equivalencias.some(eq => (eq || '').toUpperCase().replace(/[^A-Z0-9]/g, '') === cClean));
      });
      if (matchFound) return { match: matchFound, tipo: 'CODIGO_EN_DESCRIPCION', confianza: 'ALTA' };
    }
  }

  // 5. Búsqueda por diccionario de sinónimos de taller directo
  if (desc) {
    const descNorm = desc.normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
    for (const syn of WORKSHOP_SYNONYMS) {
      if (syn.keys.some(k => descNorm.includes(k) || k.includes(descNorm))) {
        const synMatch = allItems.find(p => p.numero_parte === syn.targetPart || p.id === syn.targetPart);
        if (synMatch) return { match: synMatch, tipo: 'SINONIMO_TALLER', confianza: 'ALTA' };
      }
    }
  }

  // 6. Coincidencia por descripción exacta
  if (desc.length > 3) {
    let exactDesc = allItems.find(p => (p.descripcion || '').toLowerCase().trim() === desc.trim());
    if (exactDesc) return { match: exactDesc, tipo: 'DESCRIPCION_EXACTA', confianza: 'ALTA' };
  }

  // 7. Búsqueda multi-palabra directa contra el catálogo
  if (descRaw && descRaw.trim().length >= 4) {
    const multiMatches = allItems.filter(p => matchPartMultiWord(p, descRaw));
    if (multiMatches.length === 1) {
      return { match: multiMatches[0], tipo: 'PALABRAS_CLAVE', confianza: 'ALTA' };
    }
    if (multiMatches.length > 1 && multiMatches.length <= 4) {
      if (marcaRaw && marcaRaw.trim().length >= 3) {
        const mNorm = normalizeSearchText(marcaRaw);
        const matchMarca = multiMatches.find(p => normalizeSearchText(p.marca).includes(mNorm));
        if (matchMarca) return { match: matchMarca, tipo: 'PALABRAS_Y_MARCA', confianza: 'ALTA' };
      }
      return { match: multiMatches[0], sugerencias: multiMatches, tipo: 'PALABRAS_CLAVE', confianza: 'MEDIA' };
    }
  }

  // 8. Coincidencia difusa por tokens morfológicos (lematizados)
  const inputTokens = desc
    .split(/\s+/)
    .map(stemWord)
    .filter(w => w.length > 2 && !STOPWORDS.has(w));

  if (inputTokens.length > 0) {
    const scoredItems = [];

    for (const p of allItems) {
      const pDescTokens = (p.descripcion || '')
        .split(/\s+/)
        .map(stemWord)
        .filter(w => w.length > 2 && !STOPWORDS.has(w));

      if (pDescTokens.length === 0) continue;

      let matchedTokensCount = 0;
      for (const t of inputTokens) {
        if (pDescTokens.some(pt => pt === t || pt.includes(t) || t.includes(pt))) {
          matchedTokensCount++;
        }
      }

      if (matchedTokensCount > 0) {
        const score = (matchedTokensCount * 2) / (inputTokens.length + pDescTokens.length);
        scoredItems.push({ item: p, score, matchedCount: matchedTokensCount });
      }
    }

    scoredItems.sort((a, b) => b.score - a.score || b.matchedCount - a.matchedCount);

    if (scoredItems.length > 0) {
      const best = scoredItems[0];
      const matchRatio = best.matchedCount / inputTokens.length;

      if (matchRatio >= 0.6 || (best.matchedCount >= 2 && best.score >= 0.45)) {
        return { match: best.item, tipo: 'PALABRAS_CLAVE', confianza: matchRatio >= 0.8 ? 'ALTA' : 'MEDIA' };
      }

      const sugerencias = scoredItems.slice(0, 5).map(s => s.item);
      return { match: null, sugerencias, tipo: 'NINGUNO', confianza: 'BAJA' };
    }
  }

  // 9. Sugerencias usando matchPartMultiWord con consulta completa
  let sugerencias = allItems.filter(p => matchPartMultiWord(p, numParteRaw || descRaw)).slice(0, 5);
  if (sugerencias.length === 0 && numParteClean.length >= 3) {
    sugerencias = allItems.filter(p => {
      const pClean = (p.numero_parte || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
      return pClean.includes(numParteClean) || numParteClean.includes(pClean);
    }).slice(0, 5);
  }

  return { match: null, sugerencias, tipo: 'NINGUNO', confianza: 'BAJA' };
}

// 4.1 Agregar equivalencia de otra marca
app.post('/api/agregar-equivalencia', (req, res) => {
  const { id_pieza, codigo_equivalente } = req.body;
  const item = (store.catalog && store.catalog[id_pieza]) || (typeof findPart === 'function' ? findPart(id_pieza) : null);
  if (!item) return res.status(404).json({ error: 'Refacción no encontrada' });

  const rawCodes = String(codigo_equivalente || '').split(/[,;\n]+/).map(c => c.trim().toUpperCase()).filter(Boolean);
  if (rawCodes.length === 0) return res.status(400).json({ error: 'Código de equivalencia inválido' });

  if (!Array.isArray(item.equivalencias)) item.equivalencias = [];
  rawCodes.forEach(codeClean => {
    if (!item.equivalencias.includes(codeClean) && codeClean !== (item.numero_parte || '').toUpperCase()) {
      item.equivalencias.push(codeClean);
    }
  });
  saveStore();

  res.json({ success: true, item, equivalencias: item.equivalencias, message: `Equivalencia ${codeClean} agregada a ${item.descripcion}` });
});

// 4.2 Eliminar equivalencia
app.post('/api/eliminar-equivalencia', (req, res) => {
  const { id_pieza, codigo_equivalente } = req.body;
  const item = (store.catalog && store.catalog[id_pieza]) || (typeof findPart === 'function' ? findPart(id_pieza) : null);
  if (!item) return res.status(404).json({ error: 'Refacción no encontrada' });

  if (Array.isArray(item.equivalencias)) {
    item.equivalencias = item.equivalencias.filter(c => c !== codigo_equivalente);
    saveStore();
  }
  res.json({ success: true, item, equivalencias: item.equivalencias });
});

// 4.3 Agregar nombre alternativo o alias
app.post('/api/agregar-alias', (req, res) => {
  const { id_pieza, nuevo_alias } = req.body;
  const item = (store.catalog && store.catalog[id_pieza]) || (typeof findPart === 'function' ? findPart(id_pieza) : null);
  if (!item) return res.status(404).json({ error: 'Refacción no encontrada' });

  const rawAliases = String(nuevo_alias || '').split(/[,;\n]+/).map(a => a.trim()).filter(Boolean);
  if (rawAliases.length === 0) return res.status(400).json({ error: 'Alias inválido' });

  if (!Array.isArray(item.alias)) item.alias = [];
  rawAliases.forEach(aliasClean => {
    if (!item.alias.some(a => a.toLowerCase() === aliasClean.toLowerCase()) && aliasClean.toLowerCase() !== (item.descripcion || '').toLowerCase()) {
      item.alias.push(aliasClean);
    }
  });
  saveStore();

  res.json({ success: true, item, alias: item.alias, message: `Alias '${rawAliases.join(', ')}' agregado a ${item.descripcion}` });
});

// 4.4 Eliminar alias
app.post('/api/eliminar-alias', (req, res) => {
  const { id_pieza, alias } = req.body;
  const item = (store.catalog && store.catalog[id_pieza]) || (typeof findPart === 'function' ? findPart(id_pieza) : null);
  if (!item) return res.status(404).json({ error: 'Refacción no encontrada' });

  if (Array.isArray(item.alias)) {
    item.alias = item.alias.filter(a => a !== alias);
    saveStore();
  }
  res.json({ success: true, item, alias: item.alias });
});

// 4.4b Actualizar foto de referencia de la pieza
app.post('/api/actualizar-foto-pieza', (req, res) => {
  const { id_pieza, foto } = req.body;
  const item = (store.catalog && store.catalog[id_pieza]) || (typeof findPart === 'function' ? findPart(id_pieza) : null);
  if (!item) return res.status(404).json({ error: 'Refacción no encontrada' });
  item.foto_referencia = foto || '';
  saveStore();
  res.json({ success: true, item, foto_referencia: item.foto_referencia, message: `Foto de referencia actualizada para ${item.descripcion}` });
});

// =========================================================
// 4.4c MOTOR RESILIENTE MULTI-PROVEEDOR DE BÚSQUEDA DE IMÁGENES
// (DuckDuckGo + Bing + Wikimedia Commons + Cache en Memoria)
// =========================================================
const imageSearchCache = new Map(); // query -> { results, timestamp }

const imageSearchUserAgents = [
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:123.0) Gecko/20100101 Firefox/123.0',
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36',
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36 Edg/122.0.0.0',
  'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.3 Safari/605.1.15',
  'Mozilla/5.0 (X11; Linux x86_64; rv:122.0) Gecko/20100101 Firefox/122.0'
];

function getImageSearchUA() {
  return imageSearchUserAgents[Math.floor(Math.random() * imageSearchUserAgents.length)];
}

function cleanImageQuery(q) {
  let s = String(q || '').trim();
  s = s.replace(/INT-\d+/gi, '');
  s = s.replace(/\b(S\/N|N\/A|SIN CODIGO|GENERICA|GENERICO)\b/gi, '');
  s = s.replace(/_{2,}/g, ' ');
  s = s.replace(/[-_]/g, ' ');
  s = s.replace(/\s+/g, ' ').trim();
  return s;
}

async function searchDDGImages(query) {
  try {
    const ua = getImageSearchUA();
    const tokenUrl = 'https://duckduckgo.com/?q=' + encodeURIComponent(query) + '&iax=images&ia=images';
    const resToken = await fetch(tokenUrl, {
      headers: {
        'User-Agent': ua,
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'es-MX,es;q=0.9,en;q=0.8'
      }
    });
    const html = await resToken.text();
    const vqdMatch = html.match(/vqd=["']?([^"'>&]+)/i) || html.match(/vqd=([^&]+)/i);
    if (!vqdMatch) return [];
    const vqd = vqdMatch[1];

    const apiUrl = 'https://duckduckgo.com/i.js?l=es-es&o=json&q=' + encodeURIComponent(query) + '&vqd=' + vqd + '&f=,,,';
    const apiRes = await fetch(apiUrl, {
      headers: {
        'User-Agent': ua,
        'Referer': 'https://duckduckgo.com/'
      }
    });
    if (apiRes.status !== 200) return [];
    const text = await apiRes.text();
    if (text.includes('botnet') || text.includes('challenge')) return [];
    const data = JSON.parse(text);
    return (data.results || []).slice(0, 10).map(r => ({
      title: r.title,
      image: r.image,
      thumbnail: r.thumbnail,
      source: 'DuckDuckGo'
    }));
  } catch (err) {
    return [];
  }
}

async function searchBingImages(query) {
  try {
    const url = 'https://www.bing.com/images/search?q=' + encodeURIComponent(query) + '&first=1';
    const res = await fetch(url, {
      headers: {
        'User-Agent': getImageSearchUA(),
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'es-MX,es;q=0.9,en;q=0.8'
      }
    });
    const rawHtml = await res.text();
    const html = rawHtml.replace(/&quot;/g, '"');
    
    const murls = [...html.matchAll(/"murl":"(https?:\/\/[^"]+)"/g)].map(m => m[1]);
    const turls = [...html.matchAll(/"turl":"(https?:\/\/[^"]+)"/g)].map(m => m[1]);
    
    const results = [];
    for (let i = 0; i < Math.min(murls.length, 10); i++) {
      results.push({
        title: query,
        image: murls[i],
        thumbnail: turls[i] || murls[i],
        source: 'Bing'
      });
    }
    return results;
  } catch (err) {
    return [];
  }
}

async function searchWikimediaImages(query) {
  try {
    const url = 'https://commons.wikimedia.org/w/api.php?action=query&generator=search&gsrsearch=' + encodeURIComponent(query) + '&gsrlimit=8&prop=imageinfo&iiprop=url|size&format=json&origin=*';
    const res = await fetch(url, { headers: { 'User-Agent': 'AlmacenCESS/1.0' } });
    const data = await res.json();
    const pages = data.query?.pages || {};
    const results = [];
    Object.values(pages).forEach(p => {
      const info = p.imageinfo?.[0];
      if (info && info.url) {
        results.push({
          title: p.title?.replace('File:', '') || query,
          image: info.url,
          thumbnail: info.thumburl || info.url,
          source: 'Wikimedia'
        });
      }
    });
    return results;
  } catch (err) {
    return [];
  }
}

app.get('/api/buscar-fotos-internet', async (req, res) => {
  try {
    const rawQ = String(req.query.q || '').trim();
    if (!rawQ) return res.json({ success: true, results: [] });

    const clean = cleanImageQuery(rawQ);
    if (!clean) return res.json({ success: true, results: [] });

    // 1. Revisar cache en memoria (12 horas)
    const cached = imageSearchCache.get(clean.toLowerCase());
    if (cached && (Date.now() - cached.timestamp < 1000 * 60 * 60 * 12)) {
      return res.json({ success: true, results: cached.results, fromCache: true });
    }

    let results = [];

    // Intento 1: DuckDuckGo con rotación de UA
    results = await searchDDGImages(clean);

    // Intento 2 (Fallback): Bing Image Engine si DDG fue rate-limiteado
    if (results.length === 0) {
      results = await searchBingImages(clean);
    }

    // Intento 3: Simplificación de consulta si venía muy larga (> 2 palabras)
    if (results.length === 0) {
      const words = clean.split(' ').filter(w => w.length > 2);
      if (words.length > 2) {
        const simple = words.slice(0, 3).join(' ');
        results = await searchDDGImages(simple);
        if (results.length === 0) {
          results = await searchBingImages(simple);
        }
      }
    }

    // Intento 4: Wikimedia Commons (repositorio libre)
    if (results.length === 0) {
      results = await searchWikimediaImages(clean);
    }

    if (results.length > 0) {
      imageSearchCache.set(clean.toLowerCase(), { results, timestamp: Date.now() });
    }

    res.json({ success: true, results });
  } catch (err) {
    console.error('Error al buscar fotos en internet:', err.message);
    res.json({ success: true, results: [] });
  }
});

// 4.4d Descargar y asignar foto de internet a la pieza permanentemente
app.post('/api/asignar-foto-internet', async (req, res) => {
  try {
    const { id_pieza, imageUrl } = req.body;
    const item = (store.catalog && store.catalog[id_pieza]) || (typeof findPart === 'function' ? findPart(id_pieza) : null);
    if (!item) return res.status(404).json({ error: 'Refacción no encontrada' });
    if (!imageUrl) return res.status(400).json({ error: 'URL de imagen no proporcionada' });

    let buffer = null;
    try {
      const imgRes = await fetch(imageUrl, {
        headers: {
          'User-Agent': getImageSearchUA(),
          'Accept': 'image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8',
          'Referer': 'https://www.google.com/'
        }
      });
      if (imgRes.ok) {
        buffer = Buffer.from(await imgRes.arrayBuffer());
      }
    } catch (e) {
      console.warn('[Asignar Foto] Error descargando buffer:', e.message);
    }

    if (buffer && buffer.length > 500) {
      const safeId = String(item.id || id_pieza).replace(/[^a-zA-Z0-9_-]/g, '_');
      const fileName = `ref_${safeId}_${Date.now()}.jpg`;
      const filePath = path.join(UPLOADS_DIR, fileName);
      fs.writeFileSync(filePath, buffer);
      item.foto_referencia = `/uploads/${fileName}`;
    } else {
      // Si la descarga falló o fue rechazada por CORS del servidor externo, guardar URL directa
      item.foto_referencia = imageUrl;
    }

    saveStore();

    res.json({
      success: true,
      item,
      foto_referencia: item.foto_referencia,
      message: `Foto asignada con éxito para ${item.descripcion}`
    });
  } catch (err) {
    console.error('Error al guardar foto:', err);
    res.status(500).json({ error: 'No se pudo guardar la imagen seleccionada' });
  }
});

// 4.4e Reconciliación administrativa de saldos negativos del inventario
app.post('/api/inventario/reconciliar-negativos', (req, res) => {
  try {
    let corregidos = 0;
    const detalles = [];

    // 1. Reconciliar específicamente la VALVULA ESCAPE RAPIDA QR1 (4735000080)
    const qr1 = (store.catalog && store.catalog['4735000080']) || (typeof findPart === 'function' ? findPart('4735000080') : null);
    if (qr1) {
      const stockAnterior = qr1.stock_actual;
      qr1.stock_inicial = 0;
      qr1.stock_actual = 7; // 8 entradas registradas el 28-oct menos 1 salida del 6-nov
      detalles.push({
        id: qr1.id,
        descripcion: qr1.descripcion,
        stock_anterior: stockAnterior,
        stock_nuevo: 7,
        motivo: 'Ajuste físico QR1: 8 entradas compra fol. 22430 menos 1 salida fol. 4047'
      });
      corregidos++;
    }

    // 2. Reconciliar todas las demás piezas con stock negativo o inicial negativo
    for (const [id, item] of Object.entries(store.catalog || {})) {
      if (item.id === '4735000080' || item.numero_parte === '4735000080') continue;
      let modificado = false;
      const stockAnt = item.stock_actual;

      if (item.stock_inicial < 0) {
        item.stock_inicial = 0;
        modificado = true;
      }

      if (item.stock_actual < 0) {
        item.stock_actual = 0;
        modificado = true;
      }

      if (modificado) {
        corregidos++;
        detalles.push({
          id: item.id,
          descripcion: item.descripcion,
          stock_anterior: stockAnt,
          stock_nuevo: item.stock_actual,
          motivo: 'Reconciliación de saldo negativo en anaquel físico a 0'
        });
      }
    }

    saveStore();
    res.json({
      success: true,
      corregidos,
      detalles,
      message: `Se reconciliaron ${corregidos} refacciones con saldo negativo en el catálogo.`
    });
  } catch (err) {
    console.error('Error en reconciliar-negativos:', err);
    res.status(500).json({ error: 'Error al reconciliar saldos negativos: ' + err.message });
  }
});

// 4.5 FUSIONAR REFACCIONES DUPLICADAS (Resuelve piezas repetidas)
app.post('/api/fusionar-piezas', (req, res) => {
  const { id_destino, id_origen } = req.body;
  const target = store.catalog[id_destino];
  const source = store.catalog[id_origen];

  if (!target || !source) {
    return res.status(404).json({ error: 'Una de las piezas a fusionar no existe' });
  }

  if (target.id === source.id) {
    return res.status(400).json({ error: 'No se puede fusionar una pieza consigo misma' });
  }

  // Sumar existencias
  target.stock_actual += Number(source.stock_actual) || 0;
  target.stock_inicial += Number(source.stock_inicial) || 0;

  // Agregar número de parte origen y sus equivalencias a la pieza destino
  if (!Array.isArray(target.equivalencias)) target.equivalencias = [];
  if (source.numero_parte && !target.equivalencias.includes(source.numero_parte) && source.numero_parte !== target.numero_parte) {
    target.equivalencias.push(source.numero_parte);
  }
  if (Array.isArray(source.equivalencias)) {
    for (const eq of source.equivalencias) {
      if (!target.equivalencias.includes(eq) && eq !== target.numero_parte) {
        target.equivalencias.push(eq);
      }
    }
  }

  // Agregar alias y descripción
  if (!Array.isArray(target.alias)) target.alias = [];
  if (source.descripcion && !target.alias.includes(source.descripcion) && source.descripcion !== target.descripcion) {
    target.alias.push(source.descripcion);
  }
  if (Array.isArray(source.alias)) {
    for (const a of source.alias) {
      if (!target.alias.includes(a)) target.alias.push(a);
    }
  }

  // Reasignar movimientos en la bitácora
  for (const m of store.movimientos) {
    if (m.id_pieza === source.id) {
      m.id_pieza = target.id;
      m.numero_parte = target.numero_parte;
    }
  }

  // Eliminar la pieza origen duplicada
  delete store.catalog[id_origen];
  saveStore();

  res.json({
    success: true,
    message: `Pieza '${source.numero_parte}' fusionada exitosamente con '${target.numero_parte}'. Existencias unificadas a ${target.stock_actual} PZ.`,
    pieza_unificada: target
  });
});

// Helper para actualizar ubicación y existencias directamente en la hoja correspondiente de INVENTARIO.xlsx
function updateExcelPieceLocationAndStock(item, nuevoRack, nuevoNivel, nuevoStock) {
  if (!fs.existsSync(EXCEL_FILE)) return false;
  try {
    const wb = xlsx.readFile(EXCEL_FILE);
    const sheetsToCheck = [];
    if (item.hoja_origen && wb.SheetNames.includes(item.hoja_origen)) {
      sheetsToCheck.push(item.hoja_origen);
    }
    if (store.activeSheet && wb.SheetNames.includes(store.activeSheet) && !sheetsToCheck.includes(store.activeSheet)) {
      sheetsToCheck.push(store.activeSheet);
    }
    const alm = (store.almacenes || []).find(a => a.id === item.almacen_id);
    if (alm && alm.hoja && wb.SheetNames.includes(alm.hoja) && !sheetsToCheck.includes(alm.hoja)) {
      sheetsToCheck.push(alm.hoja);
    }
    if (sheetsToCheck.length === 0 && wb.SheetNames.length > 0) {
      sheetsToCheck.push(wb.SheetNames[0]);
    }

    let anyModified = false;
    for (const sheetName of sheetsToCheck) {
      const ws = wb.Sheets[sheetName];
      if (!ws) continue;
      const data = xlsx.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (!data || data.length < 2) continue;

      const headerRow = data[0].map(h => String(h || '').trim().toUpperCase());
      const colRack = headerRow.findIndex(h => h.includes('RACK'));
      const colNivel = headerRow.findIndex(h => h.includes('NIVEL'));
      const colParte = headerRow.findIndex(h => h.includes('PARTE') || h.includes('CODIGO') || h.includes('CÓDIGO'));
      const colDesc = headerRow.findIndex(h => h.includes('DESCRIP'));
      const colFinal = headerRow.findIndex(h => h.includes('FINAL') || h.includes('ACTUAL'));
      const colInicial = headerRow.findIndex(h => h.includes('INIC') || h.includes('CANTIDAD'));

      let sheetModified = false;
      for (let r = 1; r < data.length; r++) {
        const row = data[r];
        const parteVal = colParte !== -1 ? String(row[colParte] || '').trim() : '';
        const descVal = colDesc !== -1 ? String(row[colDesc] || '').trim() : '';

        const match = (parteVal && item.numero_parte && parteVal.toLowerCase() === item.numero_parte.toLowerCase()) ||
                      (descVal && item.descripcion && descVal.toLowerCase() === item.descripcion.toLowerCase());

        if (match) {
          if (nuevoRack && colRack !== -1) {
            row[colRack] = nuevoRack;
            sheetModified = true;
          }
          if (nuevoNivel && colNivel !== -1) {
            row[colNivel] = nuevoNivel;
            sheetModified = true;
          }
          if (nuevoStock !== undefined && nuevoStock !== null && !isNaN(Number(nuevoStock))) {
            if (colFinal !== -1) {
              row[colFinal] = Number(nuevoStock);
              sheetModified = true;
            } else if (colInicial !== -1) {
              row[colInicial] = Number(nuevoStock);
              sheetModified = true;
            }
          }
          break;
        }
      }

      if (sheetModified) {
        wb.Sheets[sheetName] = xlsx.utils.aoa_to_sheet(data);
        anyModified = true;
      }
    }

    if (anyModified) {
      xlsx.writeFile(wb, EXCEL_FILE);
      return true;
    }
  } catch (err) {
    console.error('Aviso: No se pudo escribir directamente en INVENTARIO.xlsx (posiblemente abierto en Excel):', err.message);
  }
  return false;
}

// 4.3 Ajuste de Inventario, Reubicación Física en Racks y Conteo Físico
app.post('/api/inventario/ajuste', (req, res) => {
  const {
    id_pieza,
    numero_parte,
    nuevo_stock,
    nuevo_rack,
    nuevo_nivel,
    nuevo_almacen_id,
    nueva_descripcion,
    nueva_marca,
    nuevo_precio,
    nueva_unidad,
    motivo,
    justificacion,
    usuarioNombre,
    usuarioRol,
    permisos
  } = req.body;

  if (usuarioRol === 'GERENTE' || (permisos && permisos.admin_centro_control === false && permisos.dar_salida === false && permisos.dar_entrada === false)) {
    return res.status(403).json({ error: 'Acceso denegado: Tu perfil es de Solo Consulta / Auditoría y no tiene permisos para ajustar inventario.' });
  }

  let item = findPart(id_pieza) || findPart(numero_parte);
  if (!item && (numero_parte || descripcion)) {
    const matchRes = matchItemToCatalog(numero_parte, descripcion, marca);
    if (matchRes && matchRes.match && (matchRes.tipo === 'EXACTA' || matchRes.tipo === 'NORMALIZADA' || matchRes.tipo === 'EQUIVALENCIA')) {
      item = matchRes.match;
    }
  }
  if (!item) {
    return res.status(404).json({ error: 'Refacción no encontrada en el catálogo.' });
  }

  const prevStock = Number(item.stock_actual) || 0;
  const prevRack = item.rack || 'Rack 1';
  const prevNivel = item.nivel || 'Nivel 1';
  const prevAlmId = item.almacen_id || 'refacciones';
  const prevAlmNombre = item.almacen || 'Almacén de Refacciones';

  let stockCambiado = false;
  let diff = 0;
  let newStockNum = prevStock;

  if (nuevo_stock !== undefined && nuevo_stock !== null && String(nuevo_stock).trim() !== '') {
    newStockNum = Math.max(0, Number(nuevo_stock));
    if (!isNaN(newStockNum) && newStockNum !== prevStock) {
      diff = newStockNum - prevStock;
      item.stock_actual = newStockNum;
      stockCambiado = true;
    }
  }

  let ubicacionCambiada = false;
  let rackFinal = prevRack;
  let nivelFinal = prevNivel;

  if (nuevo_rack && String(nuevo_rack).trim() !== '' && String(nuevo_rack).trim() !== prevRack) {
    rackFinal = String(nuevo_rack).trim();
    if (!rackFinal.toLowerCase().startsWith('rack') && !isNaN(Number(rackFinal))) {
      rackFinal = `Rack ${rackFinal}`;
    }
    item.rack = rackFinal;
    ubicacionCambiada = true;
  }

  if (nuevo_nivel && String(nuevo_nivel).trim() !== '' && String(nuevo_nivel).trim() !== prevNivel) {
    nivelFinal = String(nuevo_nivel).trim();
    if (!nivelFinal.toLowerCase().startsWith('nivel') && !isNaN(Number(nivelFinal))) {
      nivelFinal = `Nivel ${nivelFinal}`;
    }
    item.nivel = nivelFinal;
    ubicacionCambiada = true;
  }

  let almacenCambiado = false;
  if (nuevo_almacen_id && nuevo_almacen_id !== prevAlmId) {
    const targetAlm = (store.almacenes || []).find(a => a.id === nuevo_almacen_id);
    if (targetAlm) {
      item.almacen_id = targetAlm.id;
      item.almacen = targetAlm.nombre;
      almacenCambiado = true;
    }
  }

  // Datos maestros opcionales
  if (nueva_descripcion && String(nueva_descripcion).trim() !== '') {
    item.descripcion = String(nueva_descripcion).trim();
  }
  if (nueva_marca && String(nueva_marca).trim() !== '') {
    item.marca = String(nueva_marca).trim().toUpperCase();
  }
  if (nuevo_precio !== undefined && nuevo_precio !== null && !isNaN(Number(nuevo_precio))) {
    const pNum = Number(nuevo_precio);
    item.precio_unitario = pNum;
    // Propagar nuevo precio a movimientos históricos de esta pieza que tenían precio 0 o error decimal
    (store.movimientos || []).forEach(m => {
      if (m.id_pieza === item.id || m.numero_parte === item.numero_parte) {
        const currP = Number(m.costo_unitario) || 0;
        if (currP <= 0 || (pNum >= 100 && currP < 10 && pNum / currP > 500)) {
          m.costo_unitario = pNum;
          m.costo_total = (Number(m.cantidad) || 1) * pNum;
        }
      }
    });
    try { reconciliarAlmacenConTaller(); } catch(e) {}
  }
  if (nueva_unidad && String(nueva_unidad).trim() !== '') {
    item.unidad = String(nueva_unidad).trim().toUpperCase();
  }

  const now = new Date();
  const fechaHoy = now.toISOString().slice(0, 10);
  const horaHoy = now.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
  const motivoTexto = (motivo || (stockCambiado ? 'Conteo Físico / Auditoría Cíclica' : 'Reacomodo de Espacio y Ubicación')).trim();

  // Generar movimiento de auditoría si hubo cambio en stock, ubicación o almacén
  let mov = null;
  if (stockCambiado || ubicacionCambiada || almacenCambiado) {
    const notasDetalle = [];
    if (stockCambiado) {
      notasDetalle.push(`Stock: ${prevStock} ➔ ${item.stock_actual} (${diff > 0 ? '+' : ''}${diff} ${item.unidad || 'PZ'})`);
    }
    if (ubicacionCambiada) {
      notasDetalle.push(`Ubicación: "${prevRack} / ${prevNivel}" ➔ "${item.rack} / ${item.nivel}"`);
    }
    if (almacenCambiado) {
      notasDetalle.push(`Almacén: "${prevAlmNombre}" ➔ "${item.almacen}"`);
    }
    notasDetalle.push(`Motivo: ${motivoTexto}`);
    if (justificacion && String(justificacion).trim()) {
      notasDetalle.push(`Nota: ${String(justificacion).trim()}`);
    }

    mov = {
      id: `ajuste_${Date.now()}_${Math.round(Math.random() * 1000)}`,
      tipo: 'AJUSTE',
      fecha: fechaHoy,
      hora: horaHoy,
      id_pieza: item.id,
      numero_parte: item.numero_parte,
      descripcion: item.descripcion,
      marca: item.marca,
      almacen: item.almacen,
      almacen_id: item.almacen_id,
      ubicacion: `${item.rack} / ${item.nivel}`,
      cantidad: Math.abs(diff),
      stock_anterior: prevStock,
      stock_nuevo: item.stock_actual,
      diferencia_stock: diff,
      ubicacion_anterior: `${prevRack} / ${prevNivel}`,
      ubicacion_nueva: `${item.rack} / ${item.nivel}`,
      almacen_anterior: prevAlmNombre,
      almacen_nuevo: item.almacen,
      almacen_nuevo_id: item.almacen_id,
      motivo_ajuste: motivoTexto,
      solicitante: usuarioNombre || 'Administrador de Almacén',
      notas: `[AJUSTE: ${notasDetalle.join(' | ')}]`,
      timestamp: now.toISOString()
    };

    store.movimientos.unshift(mov);
  }

  // Sincronizar en Excel si el archivo existe
  const excelUpdated = updateExcelPieceLocationAndStock(item, item.rack, item.nivel, stockCambiado ? item.stock_actual : null);

  saveStore();

  res.json({
    success: true,
    message: stockCambiado && ubicacionCambiada
      ? `Ajuste y reubicación aplicados: Stock ahora es ${item.stock_actual} PZ en ${item.rack} / ${item.nivel}.`
      : stockCambiado
      ? `Existencia física ajustada a ${item.stock_actual} PZ (${diff > 0 ? '+' : ''}${diff} PZ).`
      : ubicacionCambiada
      ? `Pieza reubicada con éxito en ${item.rack} / ${item.nivel}.`
      : `Datos de refacción actualizados con éxito.`,
    item,
    movimiento: mov,
    excel_actualizado: excelUpdated
  });
});

// 4.4 Módulo de Cuadre de Inventario Físico - Estadísticas y Descuadres
app.get('/api/inventario/descuadres', (req, res) => {
  const items = Object.values(store.catalog);
  const negativas = items.filter(p => Number(p.stock_actual) < 0);
  const enCero = items.filter(p => Number(p.stock_actual) === 0);
  const conStock = items.filter(p => Number(p.stock_actual) > 0);

  const unidadesNegativas = Math.abs(negativas.reduce((sum, p) => sum + Number(p.stock_actual || 0), 0));
  const valorDescuadre = Math.abs(negativas.reduce((sum, p) => sum + (Number(p.stock_actual || 0) * Number(p.precio_unitario || 0)), 0));

  res.json({
    success: true,
    stats: {
      totalNegativas: negativas.length,
      totalEnCero: enCero.length,
      totalConStock: conStock.length,
      totalCatalogo: items.length,
      unidadesNegativas,
      valorDescuadre
    },
    negativas,
    enCeroCount: enCero.length,
    conStockCount: conStock.length
  });
});

// 4.5 Módulo de Cuadre de Inventario Físico - Cuadre Individual
app.post('/api/inventario/cuadre', (req, res) => {
  const { id_pieza, nuevo_stock_fisico, motivo, usuarioNombre, usuarioRol, permisos } = req.body;
  if (usuarioRol === 'GERENTE' || (permisos && permisos.admin_centro_control === false && permisos.dar_salida === false && permisos.dar_entrada === false)) {
    return res.status(403).json({ error: 'Acceso denegado: Tu perfil es de Solo Consulta / Auditoría y no tiene permisos para ajustar inventario.' });
  }

  const item = store.catalog[id_pieza] || findPart(id_pieza);
  if (!item) {
    return res.status(404).json({ error: 'Refacción no encontrada en el catálogo.' });
  }

  const stockPrevio = Number(item.stock_actual) || 0;
  const nuevoStock = Number(nuevo_stock_fisico);
  if (isNaN(nuevoStock)) {
    return res.status(400).json({ error: 'El valor de conteo físico ingresado no es válido.' });
  }

  const delta = nuevoStock - stockPrevio;
  item.stock_actual = nuevoStock;

  const now = new Date();
  const fechaHoy = now.toISOString().slice(0, 10);
  const horaHoy = now.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
  const motivoTexto = motivo || (delta > 0 ? 'Conteo físico: Sobrante detectado' : delta < 0 ? 'Conteo físico: Faltante / Merma' : 'Conteo físico: Cuadre exacto');

  const mov = {
    id: `cuadre_${Date.now()}_${Math.floor(100 + Math.random() * 900)}`,
    tipo: 'AJUSTE',
    subtipo: 'CUADRE_FISICO',
    fecha: fechaHoy,
    hora: horaHoy,
    id_pieza: item.id,
    numero_parte: item.numero_parte,
    descripcion: item.descripcion,
    marca: item.marca,
    almacen: item.almacen,
    almacen_id: item.almacen_id,
    cantidad: delta,
    stock_previo: stockPrevio,
    stock_nuevo: nuevoStock,
    diferencia_stock: delta,
    solicitante: usuarioNombre || 'Administrador de Almacén',
    motivo_ajuste: motivoTexto,
    notas: `[CUADRE FÍSICO: Stock en sistema pasó de ${stockPrevio} a ${nuevoStock} PZ (${delta >= 0 ? '+' : ''}${delta} PZ). Motivo: ${motivoTexto}]`,
    timestamp: now.toISOString()
  };

  store.movimientos.unshift(mov);
  updateExcelPieceLocationAndStock(item, item.rack, item.nivel, nuevoStock);
  saveStore();

  res.json({
    success: true,
    message: `Pieza ${item.numero_parte} (${item.descripcion}) cuadrada exitosamente a ${nuevoStock} PZ (${delta >= 0 ? '+' : ''}${delta} PZ).`,
    item,
    delta,
    movimiento: mov
  });
});

// 4.6 Módulo de Cuadre de Inventario Físico - Cuadre Masivo / Lote
app.post('/api/inventario/cuadre-masivo', (req, res) => {
  const { ajustes, motivo_general, reset_negativas, usuarioNombre, usuarioRol, permisos } = req.body;
  if (usuarioRol === 'GERENTE' || (permisos && permisos.admin_centro_control === false && permisos.dar_salida === false && permisos.dar_entrada === false)) {
    return res.status(403).json({ error: 'Acceso denegado: Tu perfil es de Solo Consulta / Auditoría y no tiene permisos para ajustar inventario.' });
  }

  const now = new Date();
  const fechaHoy = now.toISOString().slice(0, 10);
  const horaHoy = now.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit', second: '2-digit' });

  let procesados = 0;

  // Opción 1: Reset de todas las piezas negativas a 0 en 1 solo clic
  if (reset_negativas) {
    for (const item of Object.values(store.catalog)) {
      const stockPrevio = Number(item.stock_actual) || 0;
      if (stockPrevio < 0) {
        const delta = -stockPrevio;
        item.stock_actual = 0;
        if (item.stock_inicial < 0) item.stock_inicial = 0;

        const mov = {
          id: `cuadre_neg_${Date.now()}_${procesados}`,
          tipo: 'AJUSTE',
          subtipo: 'RESET_NEGATIVAS',
          fecha: fechaHoy,
          hora: horaHoy,
          id_pieza: item.id,
          numero_parte: item.numero_parte,
          descripcion: item.descripcion,
          marca: item.marca,
          almacen: item.almacen,
          almacen_id: item.almacen_id,
          cantidad: delta,
          stock_previo: stockPrevio,
          stock_nuevo: 0,
          diferencia_stock: delta,
          solicitante: usuarioNombre || 'Administrador de Almacén',
          motivo_ajuste: motivo_general || 'Cuadre masivo: Corrección de descuadres negativos a cero',
          notas: `[CUADRE MASIVO: Descuadre negativo corregido de ${stockPrevio} PZ a 0 PZ]`,
          timestamp: now.toISOString()
        };

        store.movimientos.unshift(mov);
        updateExcelPieceLocationAndStock(item, item.rack, item.nivel, 0);
        procesados++;
      }
    }
  } else if (Array.isArray(ajustes) && ajustes.length > 0) {
    // Opción 2: Lote de piezas editadas durante el conteo físico
    for (const aj of ajustes) {
      const item = store.catalog[aj.id_pieza] || findPart(aj.id_pieza) || findPart(aj.numero_parte);
      if (!item) continue;

      const stockPrevio = Number(item.stock_actual) || 0;
      const nuevoStock = Number(aj.nuevo_stock);
      if (isNaN(nuevoStock) || nuevoStock === stockPrevio) continue;

      const delta = nuevoStock - stockPrevio;
      item.stock_actual = nuevoStock;

      const mov = {
        id: `cuadre_lote_${Date.now()}_${procesados}`,
        tipo: 'AJUSTE',
        subtipo: 'CUADRE_FISICO',
        fecha: fechaHoy,
        hora: horaHoy,
        id_pieza: item.id,
        numero_parte: item.numero_parte,
        descripcion: item.descripcion,
        marca: item.marca,
        almacen: item.almacen,
        almacen_id: item.almacen_id,
        cantidad: delta,
        stock_previo: stockPrevio,
        stock_nuevo: nuevoStock,
        diferencia_stock: delta,
        solicitante: usuarioNombre || 'Administrador de Almacén',
        motivo_ajuste: aj.motivo || motivo_general || 'Auditoría de inventario físico',
        notas: `[CUADRE FÍSICO: ${stockPrevio} -> ${nuevoStock} PZ (${delta >= 0 ? '+' : ''}${delta} PZ)]`,
        timestamp: now.toISOString()
      };

      store.movimientos.unshift(mov);
      updateExcelPieceLocationAndStock(item, item.rack, item.nivel, nuevoStock);
      procesados++;
    }
  }

  if (procesados > 0) {
    saveStore();
  }

  res.json({
    success: true,
    procesados,
    message: `Se cuadraron exitosamente ${procesados} refacciones en el inventario.`
  });
});

// Helper para extraer equivalencias y alias desde cualquier petición (soporta arrays o strings separados por comas)
function extractEquivalenciasAndAlias(body) {
  if (!body) return { equivalencias: [], alias: [] };
  let equivalencias = [];
  if (Array.isArray(body.equivalencias)) {
    equivalencias = body.equivalencias.map(x => String(x || '').trim().toUpperCase()).filter(Boolean);
  } else if (typeof body.equivalencias === 'string' && body.equivalencias.trim()) {
    equivalencias = body.equivalencias.split(/[,;\n]+/).map(x => x.trim().toUpperCase()).filter(Boolean);
  } else if (body.codigos_cruce) {
    if (Array.isArray(body.codigos_cruce)) {
      equivalencias = body.codigos_cruce.map(x => String(x || '').trim().toUpperCase()).filter(Boolean);
    } else if (typeof body.codigos_cruce === 'string' && body.codigos_cruce.trim()) {
      equivalencias = body.codigos_cruce.split(/[,;\n]+/).map(x => x.trim().toUpperCase()).filter(Boolean);
    }
  }

  let alias = [];
  if (Array.isArray(body.alias)) {
    alias = body.alias.map(x => String(x || '').trim()).filter(Boolean);
  } else if (typeof body.alias === 'string' && body.alias.trim()) {
    alias = body.alias.split(/[,;\n]+/).map(x => x.trim()).filter(Boolean);
  } else if (body.apodos) {
    if (Array.isArray(body.apodos)) {
      alias = body.apodos.map(x => String(x || '').trim()).filter(Boolean);
    } else if (typeof body.apodos === 'string' && body.apodos.trim()) {
      alias = body.apodos.split(/[,;\n]+/).map(x => x.trim()).filter(Boolean);
    }
  }

  return { equivalencias, alias };
}

// Helper para fusionar nuevas equivalencias y alias sin duplicados
function mergeEquivalenciasAndAlias(item, newEquivs, newAlias) {
  if (!item) return;
  if (!Array.isArray(item.equivalencias)) item.equivalencias = [];
  if (!Array.isArray(item.alias)) item.alias = [];

  const mainPart = String(item.numero_parte || '').toUpperCase();
  for (const eq of newEquivs) {
    const clean = String(eq || '').trim().toUpperCase();
    if (clean && clean !== mainPart && !item.equivalencias.includes(clean)) {
      item.equivalencias.push(clean);
      console.log(`[CESS] Nueva equivalencia registrada para ${item.numero_parte}: ${clean}`);
    }
  }

  const mainDesc = String(item.descripcion || '').toLowerCase();
  for (const al of newAlias) {
    const clean = String(al || '').trim();
    if (clean && clean.toLowerCase() !== mainDesc && !item.alias.some(a => a.toLowerCase() === clean.toLowerCase())) {
      item.alias.push(clean);
      console.log(`[CESS] Nuevo apodo/alias registrado para ${item.numero_parte}: ${clean}`);
    }
  }
}

// 5. Registrar SALIDA
app.post('/api/salida', (req, res) => {
  if (req.body.usuarioRol === 'GERENTE' || (req.body.permisos && req.body.permisos.dar_salida === false)) {
    return res.status(403).json({ error: 'Acceso denegado: Tu perfil es de Solo Consulta / Auditoría.' });
  }

  const {
    id_pieza,
    numero_parte,
    descripcion,
    marca,
    ubicacion,
    cantidad,
    unidad_camion,
    folio,
    solicitante,
    foto,
    notas
  } = req.body;

  const cant = Number(cantidad);
  if (!cant || cant <= 0) {
    return res.status(400).json({ error: 'Cantidad inválida' });
  }

  // Buscar en catálogo considerando equivalencias y coincidencias inteligentes
  let item = findPart(id_pieza) || findPart(numero_parte);
  if (!item && (numero_parte || descripcion)) {
    const matchRes = matchItemToCatalog(numero_parte, descripcion, marca);
    if (matchRes && matchRes.match && (matchRes.tipo === 'EXACTA' || matchRes.tipo === 'NORMALIZADA' || matchRes.tipo === 'EQUIVALENCIA')) {
      item = matchRes.match;
    }
  }

  const isSrv = Boolean(req.body.es_servicio || req.body.tipo === 'SERVICIO' || isServiceDescription(descripcion));

  // Candado Estricto de Precio/Costo (Auditoría y Gerencia)
  const costoUnit = Number(req.body.costo_unitario !== undefined && req.body.costo_unitario !== '' ? req.body.costo_unitario : (item ? item.precio_unitario : 0)) || 0;
  if (!costoUnit || costoUnit <= 0) {
    return res.status(400).json({
      success: false,
      error: 'CANDADO_PRECIO_REQUERIDO',
      message: `Candado de Auditoría: No se permite registrar ${isSrv ? 'servicios' : 'salidas'} con costo $0.00 pesos. Debes confirmar el costo unitario para los reportes de Gerencia.`
    });
  }

  const now = new Date();
  const fecha = req.body.fecha ? String(req.body.fecha).trim() : now.toISOString().slice(0, 10);
  const hora = req.body.hora ? String(req.body.hora).trim() : now.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit', second: '2-digit' });

  // SI ES UN SERVICIO O TRABAJO EXTERNO: Gasto directo a unidad, sin descontar anaqueles físicos
  if (isSrv) {
    const mov = {
      id: `mov_${Date.now()}_${Math.round(Math.random() * 1000)}`,
      tipo: 'S',
      tipo_movimiento: 'SERVICIO',
      es_servicio: true,
      fecha,
      hora,
      id_pieza: 'SERVICIO_EXTERNO',
      numero_parte: req.body.numero_parte || 'SERVICIO',
      descripcion: String(descripcion || 'Servicio Externo').trim(),
      marca: req.body.proveedor ? String(req.body.proveedor).trim().toUpperCase() : (marca ? String(marca).trim().toUpperCase() : 'PROVEEDOR EXTERNO'),
      proveedor: req.body.proveedor ? String(req.body.proveedor).trim().toUpperCase() : '',
      factura: req.body.factura ? String(req.body.factura).trim().toUpperCase() : '',
      almacen: 'Servicios y Reparaciones Externas',
      almacen_id: 'servicios_externos',
      ubicacion: 'Servicio Externo / Proveedor',
      cantidad: cant,
      unidad_camion: String(unidad_camion || 'GENERAL').trim().toUpperCase(),
      folio: String(folio || '').trim(),
      solicitante: String(solicitante || 'JEFE DE TALLER').trim(),
      foto: foto || '',
      costo_unitario: costoUnit,
      costo_total: costoUnit * cant,
      notas: (notas ? `${notas} ` : '') + '[SERVICIO EXTERNO REGISTRADO DIRECTO A UNIDAD]',
      timestamp: now.toISOString()
    };

    store.movimientos.unshift(mov);
    saveStore();
    try { syncSalidaConTaller(mov); } catch(e) {}

    return res.json({
      success: true,
      es_servicio: true,
      message: `¡Servicio externo "${mov.descripcion}" registrado para la Unidad ${mov.unidad_camion}! ($${mov.costo_total.toLocaleString()} MXN)`,
      movimiento: mov,
      item_creado: null
    });
  }

  const almId = req.body.almacen_id || (item ? item.almacen_id : 'refacciones');
  const almObj = (store.almacenes || []).find(a => a.id === almId);
  const almNombre = almObj ? almObj.nombre : (req.body.almacen || (item ? item.almacen : 'Almacén de Refacciones'));

  let fueCreado = false;
  if (!item && (req.body.isNewPart || req.body.esPiezaNueva || req.body.crear_en_catalogo || (descripcion && String(descripcion).trim()))) {
    // Es una pieza física que no estaba en el inventario virtual: dar de alta directamente
    let finalParte = String(numero_parte || '').trim();
    const esGranel = Boolean(req.body.es_a_granel) || !finalParte || finalParte === 'S/N' || finalParte === 'N/A';
    if (esGranel && (!finalParte || finalParte === 'S/N' || finalParte === 'N/A')) {
      const nextNum = Object.keys(store.catalog).filter(k => k.startsWith('INT_') || k.startsWith('SN_')).length + 1;
      finalParte = `INT-${String(nextNum).padStart(4, '0')}`;
    }
    const safeId = (almId !== 'refacciones' ? `${almId}_` : '') + finalParte.replace(/[^a-zA-Z0-9_-]/g, '_');

    // Stock observado en anaquel (si especificó stock_encontrado ej. 5 y saca 2 -> stock_actual = 3)
    const stockEncontrado = req.body.stock_encontrado !== undefined && req.body.stock_encontrado !== ''
      ? Math.max(cant, Number(req.body.stock_encontrado))
      : cant;
    const stockActualRestante = Math.max(0, stockEncontrado - cant);
    const { equivalencias: inputEquivs, alias: inputAlias } = extractEquivalenciasAndAlias(req.body);

    item = {
      id: safeId,
      rack: req.body.rack || 'Rack 1',
      nivel: req.body.nivel || 'Nivel 1',
      descripcion: String(descripcion || 'Refacción sin descripción').trim(),
      marca: String(marca || (esGranel ? 'A GRANEL / GENÉRICO' : 'GENÉRICA')).trim(),
      numero_parte: finalParte,
      equivalencias: inputEquivs,
      alias: inputAlias,
      es_a_granel: esGranel,
      foto_referencia: foto || '',
      unidad: req.body.unidad || 'PZ',
      precio_unitario: Number(req.body.precio_unitario || req.body.costo_unitario || 0),
      stock_inicial: stockEncontrado,
      stock_actual: stockActualRestante,
      almacen: almNombre,
      almacen_id: almId,
      creado_en_salida: true,
      fecha_creacion: fecha
    };
    store.catalog[safeId] = item;
    fueCreado = true;
  } else if (item) {
    item.stock_actual -= cant;
    if (costoUnit > 0 && (!item.precio_unitario || Number(item.precio_unitario) <= 0)) {
      item.precio_unitario = costoUnit;
    }
    if (foto && !item.foto_referencia) item.foto_referencia = foto;
    const { equivalencias: inputEquivs, alias: inputAlias } = extractEquivalenciasAndAlias(req.body);
    mergeEquivalenciasAndAlias(item, inputEquivs, inputAlias);
  }

  const mov = {
    id: `mov_${Date.now()}_${Math.round(Math.random() * 1000)}`,
    tipo: 'S',
    fecha,
    hora,
    id_pieza: item ? item.id : id_pieza,
    numero_parte: item ? item.numero_parte : (numero_parte || ''),
    descripcion: item ? item.descripcion : (descripcion || ''),
    marca: item ? item.marca : (marca || ''),
    almacen: item ? (item.almacen || almNombre) : almNombre,
    almacen_id: item ? (item.almacen_id || almId) : almId,
    ubicacion: ubicacion || (item ? `${item.rack} / ${item.nivel}` : ''),
    cantidad: cant,
    unidad_camion: String(unidad_camion || '').trim().toUpperCase(),
    folio: String(folio || '').trim(),
    solicitante: String(solicitante || '').trim(),
    foto: foto || '',
    costo_unitario: costoUnit,
    costo_total: costoUnit * cant,
    notas: (notas ? `${notas} ` : '') + (fueCreado ? '[PIEZA NUEVA REGISTRADA EN CATÁLOGO DURANTE ESTA SALIDA]' : ''),
    timestamp: now.toISOString()
  };

  store.movimientos.unshift(mov);
  saveStore();
  try { syncSalidaConTaller(mov); } catch(e) {}

  res.json({
    success: true,
    message: fueCreado
      ? `¡Refacción "${item.descripcion}" agregada al catálogo y salida de ${cant} ${item.unidad} registrada!`
      : 'Salida registrada correctamente',
    movimiento: mov,
    nuevo_stock: item ? item.stock_actual : null,
    item_creado: fueCreado ? item : null
  });
});

// 6. Registrar ENTRADA
app.post('/api/entrada', (req, res) => {
  if (req.body.usuarioRol === 'GERENTE' || (req.body.permisos && req.body.permisos.dar_entrada === false)) {
    return res.status(403).json({ error: 'Acceso denegado: Tu perfil es de Solo Consulta / Auditoría.' });
  }

  const {
    id_pieza,
    numero_parte,
    descripcion,
    marca,
    rack,
    nivel,
    cantidad,
    proveedor,
    factura,
    entregado_por,
    costo_unitario,
    foto,
    notas
  } = req.body;

  const cant = Number(cantidad);
  if (!cant || cant <= 0) {
    return res.status(400).json({ error: 'Cantidad inválida' });
  }

  let item = findPart(id_pieza) || findPart(numero_parte);
  if (!item && (numero_parte || descripcion)) {
    const matchRes = matchItemToCatalog(numero_parte, descripcion, marca);
    if (matchRes && matchRes.match && (matchRes.tipo === 'EXACTA' || matchRes.tipo === 'NORMALIZADA' || matchRes.tipo === 'EQUIVALENCIA')) {
      item = matchRes.match;
    }
  }

  const now = new Date();
  const fecha = req.body.fecha ? String(req.body.fecha).trim() : now.toISOString().slice(0, 10);
  const hora = req.body.hora ? String(req.body.hora).trim() : now.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit', second: '2-digit' });

  const costoNum = costo_unitario ? Number(costo_unitario) : 0;

  const almId = req.body.almacen_id || (item ? item.almacen_id : 'refacciones');
  const almObj = (store.almacenes || []).find(a => a.id === almId);
  const almNombre = almObj ? almObj.nombre : (req.body.almacen || (item ? item.almacen : 'Almacén de Refacciones'));

  const { equivalencias: inputEquivs, alias: inputAlias } = extractEquivalenciasAndAlias(req.body);

  // Si no existe la pieza en el catálogo, la creamos
  if (!item) {
    let finalParte = String(numero_parte || '').trim();
    const esGranel = Boolean(req.body.es_a_granel) || !finalParte || finalParte === 'S/N' || finalParte === 'N/A';
    if (esGranel && (!finalParte || finalParte === 'S/N' || finalParte === 'N/A')) {
      const nextNum = Object.keys(store.catalog).filter(k => k.startsWith('INT_') || k.startsWith('SN_')).length + 1;
      finalParte = `INT-${String(nextNum).padStart(4, '0')}`;
    }
    const safeId = (almId !== 'refacciones' ? `${almId}_` : '') + finalParte.replace(/[^a-zA-Z0-9_-]/g, '_');
    item = {
      id: safeId,
      rack: rack || 'Rack 1',
      nivel: nivel || 'Nivel 1',
      descripcion: descripcion || 'Sin descripción',
      marca: marca || (esGranel ? 'A GRANEL / GENÉRICO' : 'GENÉRICA'),
      numero_parte: finalParte,
      equivalencias: inputEquivs,
      alias: inputAlias,
      es_a_granel: esGranel,
      foto_referencia: foto || '',
      unidad: req.body.unidad || 'PZ',
      precio_unitario: costoNum,
      stock_inicial: 0,
      stock_actual: cant,
      almacen: almNombre,
      almacen_id: almId
    };
    store.catalog[safeId] = item;
  } else {
    item.stock_actual += cant;
    if (costoNum > 0) item.precio_unitario = costoNum;
    if (rack) item.rack = rack;
    if (nivel) item.nivel = nivel;
    if (foto && !item.foto_referencia) item.foto_referencia = foto;
    mergeEquivalenciasAndAlias(item, inputEquivs, inputAlias);
  }


  const condicionPago = String(req.body.condicion_pago || 'CONTADO').trim().toUpperCase() === 'CREDITO' ? 'CREDITO' : 'CONTADO';
  const diasCredito = condicionPago === 'CREDITO' ? (Math.max(1, Number(req.body.dias_credito) || 30)) : 0;
  const fechaFactura = req.body.fecha_factura ? String(req.body.fecha_factura).trim() : fecha;
  const fechaVencimiento = condicionPago === 'CREDITO'
    ? (req.body.fecha_vencimiento ? String(req.body.fecha_vencimiento).trim() : calcularVencimiento(fechaFactura, diasCredito))
    : fechaFactura;
  const estadoPago = condicionPago === 'CREDITO' ? (req.body.estado_pago || 'PENDIENTE') : 'PAGADO';

  const mov = {
    id: `mov_${Date.now()}_${Math.round(Math.random() * 1000)}`,
    tipo: 'E',
    fecha,
    hora,
    id_pieza: item.id,
    numero_parte: item.numero_parte,
    descripcion: item.descripcion,
    marca: item.marca,
    almacen: item.almacen || almNombre,
    almacen_id: item.almacen_id || almId,
    ubicacion: `${item.rack} / ${item.nivel}`,
    cantidad: cant,
    proveedor: String(proveedor || '').trim(),
    factura: String(factura || '').trim(),
    entregado_por: String(entregado_por || '').trim(),
    costo_unitario: costoNum,
    condicion_pago: condicionPago,
    dias_credito: diasCredito,
    fecha_factura: fechaFactura,
    fecha_vencimiento: fechaVencimiento,
    estado_pago: estadoPago,
    fecha_pago: req.body.fecha_pago || (estadoPago === 'PAGADO' && condicionPago === 'CONTADO' ? fecha : ''),
    referencia_pago: req.body.referencia_pago || '',
    foto: foto || '',
    notas: notas || '',
    timestamp: now.toISOString()
  };

  store.movimientos.unshift(mov);
  registrarOActualizarFacturaProcesada(mov);
  saveStore();

  res.json({
    success: true,
    message: 'Entrada registrada correctamente',
    movimiento: mov,
    nuevo_stock: item.stock_actual
  });
});

// 6.1 Registrar PASO DIRECTO (E/S) - Cross-docking simultáneo entrada/salida a unidad sin almacenar
app.post('/api/paso-directo', (req, res) => {
  if (req.body.usuarioRol === 'GERENTE' || (req.body.permisos && (req.body.permisos.dar_entrada === false || req.body.permisos.dar_salida === false))) {
    return res.status(403).json({ error: 'Acceso denegado: Tu perfil no cuenta con permisos para registrar movimientos directos.' });
  }

  const {
    id_pieza,
    numero_parte,
    descripcion,
    marca,
    cantidad,
    // Datos de entrada (compra/proveedor)
    proveedor,
    factura,
    costo_unitario,
    entregado_por,
    // Datos de salida (aplicación a unidad)
    unidad_camion,
    folio,
    solicitante,
    motivo,
    // Generales
    foto,
    notas,
    fecha: fechaCustom,
    hora: horaCustom
  } = req.body;

  const cant = Number(cantidad);
  if (!cant || cant <= 0) {
    return res.status(400).json({ error: 'Cantidad inválida para el movimiento' });
  }

  const unidadClean = String(unidad_camion || '').trim().toUpperCase();
  if (!unidadClean) {
    return res.status(400).json({ error: 'Debes indicar la Unidad (Tractocamión) a la que se entrega la pieza' });
  }

  let item = findPart(id_pieza) || findPart(numero_parte);
  if (!item && (numero_parte || descripcion)) {
    const matchRes = matchItemToCatalog(numero_parte, descripcion, marca);
    if (matchRes && matchRes.match && (matchRes.tipo === 'EXACTA' || matchRes.tipo === 'NORMALIZADA' || matchRes.tipo === 'EQUIVALENCIA')) {
      item = matchRes.match;
    }
  }
  const now = new Date();
  const fecha = fechaCustom ? String(fechaCustom).trim() : now.toISOString().slice(0, 10);
  const hora = horaCustom ? String(horaCustom).trim() : now.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
  const costoNum = costo_unitario ? Number(costo_unitario) : 0;

  const almId = req.body.almacen_id || (item ? item.almacen_id : 'refacciones');
  const almObj = (store.almacenes || []).find(a => a.id === almId);
  const almNombre = almObj ? almObj.nombre : (req.body.almacen || (item ? item.almacen : 'Almacén de Refacciones'));

  const { equivalencias: inputEquivs, alias: inputAlias } = extractEquivalenciasAndAlias(req.body);

  // Si la refacción es nueva / sobre pedido, la damos de alta en catálogo para mantener historial
  if (!item) {
    let finalParte = String(numero_parte || '').trim();
    const esGranel = Boolean(req.body.es_a_granel) || !finalParte || finalParte === 'S/N' || finalParte === 'N/A';
    if (esGranel && (!finalParte || finalParte === 'S/N' || finalParte === 'N/A')) {
      const nextNum = Object.keys(store.catalog).filter(k => k.startsWith('INT_') || k.startsWith('SN_')).length + 1;
      finalParte = `INT-${String(nextNum).padStart(4, '0')}`;
    }
    const safeId = (almId !== 'refacciones' ? `${almId}_` : '') + finalParte.replace(/[^a-zA-Z0-9_-]/g, '_');
    item = {
      id: safeId,
      rack: 'PASO DIRECTO',
      nivel: `UNIDAD ${unidadClean}`,
      descripcion: descripcion || 'Refacción sobre pedido',
      marca: marca || (esGranel ? 'A GRANEL / GENÉRICO' : 'GENÉRICA'),
      numero_parte: finalParte,
      equivalencias: inputEquivs,
      alias: inputAlias,
      es_a_granel: esGranel,
      foto_referencia: foto || '',
      unidad: req.body.unidad || 'PZ',
      precio_unitario: costoNum,
      stock_inicial: 0,
      stock_actual: 0, // No entra al rack del almacén, pasa directo a la unidad
      almacen: almNombre,
      almacen_id: almId
    };
    store.catalog[safeId] = item;
  } else {
    // Si la pieza ya existía, el stock_actual NO se altera (+cant - cant = 0)
    if (costoNum > 0) item.precio_unitario = costoNum;
    if (foto && !item.foto_referencia) item.foto_referencia = foto;
    mergeEquivalenciasAndAlias(item, inputEquivs, inputAlias);
  }

  const condicionPago = String(req.body.condicion_pago || 'CONTADO').trim().toUpperCase() === 'CREDITO' ? 'CREDITO' : 'CONTADO';
  const diasCredito = condicionPago === 'CREDITO' ? (Math.max(1, Number(req.body.dias_credito) || 30)) : 0;
  const fechaFactura = req.body.fecha_factura ? String(req.body.fecha_factura).trim() : fecha;
  const fechaVencimiento = condicionPago === 'CREDITO'
    ? (req.body.fecha_vencimiento ? String(req.body.fecha_vencimiento).trim() : calcularVencimiento(fechaFactura, diasCredito))
    : fechaFactura;
  const estadoPago = condicionPago === 'CREDITO' ? (req.body.estado_pago || 'PENDIENTE') : 'PAGADO';

  const mov = {
    id: `mov_${Date.now()}_${Math.round(Math.random() * 1000)}`,
    tipo: 'E/S',
    fecha,
    hora,
    id_pieza: item.id,
    numero_parte: item.numero_parte,
    descripcion: item.descripcion,
    marca: item.marca,
    almacen: item ? (item.almacen || almNombre) : almNombre,
    almacen_id: item ? (item.almacen_id || almId) : almId,
    ubicacion: `PASO DIRECTO ➔ UNIDAD ${unidadClean}`,
    cantidad: cant,
    // Datos de Entrada
    proveedor: String(proveedor || '').trim(),
    factura: String(factura || '').trim(),
    entregado_por: String(entregado_por || '').trim(),
    costo_unitario: costoNum,
    condicion_pago: condicionPago,
    dias_credito: diasCredito,
    fecha_factura: fechaFactura,
    fecha_vencimiento: fechaVencimiento,
    estado_pago: estadoPago,
    fecha_pago: req.body.fecha_pago || (estadoPago === 'PAGADO' && condicionPago === 'CONTADO' ? fecha : ''),
    referencia_pago: req.body.referencia_pago || '',
    // Datos de Salida
    unidad_camion: unidadClean,
    folio: String(folio || '').trim(),
    solicitante: String(solicitante || '').trim(),
    motivo: String(motivo || '').trim(),
    foto: foto || '',
    notas: notas || '',
    timestamp: now.toISOString()
  };

  store.movimientos.unshift(mov);
  registrarOActualizarFacturaProcesada(mov);

  saveStore();

  res.json({
    success: true,
    message: `⚡ Paso Directo (E/S) registrado con éxito. Refacción aplicada a Unidad ${unidadClean}.`,
    movimiento: mov,
    stock_almacen: item.stock_actual
  });
});

// 7. Entrada Masiva (para remisiones leídas con IA o lista de piezas)
app.post('/api/entrada-masiva', (req, res) => {
  try {
  if (req.body.usuarioRol === 'GERENTE' || (req.body.permisos && req.body.permisos.dar_entrada === false)) {
    return res.status(403).json({ error: 'Acceso denegado: Tu perfil es de Solo Consulta / Auditoría.' });
  }

  const { piezas, proveedor, factura, entregado_por, imageHash, fecha: fechaDoc, orden_compra, folio_oc } = req.body;
  const ocFolio = String(orden_compra || folio_oc || '').trim().toUpperCase();
  if (!Array.isArray(piezas) || piezas.length === 0) {
    return res.status(400).json({ error: 'Lista de piezas inválida' });
  }

  // 1. Verificación estricta anti-duplicados de facturas
  if (factura && !isGenericInvoice(factura)) {
    const existing = findExistingFactura(factura, proveedor, imageHash);
    if (existing) {
      return res.status(409).json({
        success: false,
        error: 'FACTURA_DUPLICADA',
        message: `La factura "${factura}" del proveedor "${existing.match.proveedor}" ya fue registrada el ${existing.match.fecha_ingreso} (${existing.match.total_piezas || 'múltiples'} piezas registradas). No se permite el reingreso para evitar duplicidad de inventario.`,
        registro_previo: existing.match
      });
    }
  }

  const now = new Date();
  const fecha = fechaDoc ? String(fechaDoc).trim() : now.toISOString().slice(0, 10);
  const hora = now.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit', second: '2-digit' });

  const condicionPago = String(req.body.condicion_pago || 'CONTADO').trim().toUpperCase() === 'CREDITO' ? 'CREDITO' : 'CONTADO';
  const diasCredito = condicionPago === 'CREDITO' ? (Math.max(1, Number(req.body.dias_credito) || 30)) : 0;
  const fechaFactura = req.body.fecha_factura ? String(req.body.fecha_factura).trim() : (req.body.fecha_documento || fecha);
  const fechaVencimiento = condicionPago === 'CREDITO'
    ? (req.body.fecha_vencimiento ? String(req.body.fecha_vencimiento).trim() : calcularVencimiento(fechaFactura, diasCredito))
    : fechaFactura;
  const estadoPago = condicionPago === 'CREDITO' ? (req.body.estado_pago || 'PENDIENTE') : 'PAGADO';
  const fechaPago = req.body.fecha_pago || (estadoPago === 'PAGADO' && condicionPago === 'CONTADO' ? fecha : '');
  const referenciaPago = req.body.referencia_pago || '';

  let procesados = 0;

  for (const p of piezas) {
    const cant = Number(p.cantidad) || 1;
    const parte = String(p.numero_parte || '').trim();

    // 1. Buscar por ID directo si viene identificado del frontend o escáner IA
    let item = (p.id && store.catalog[p.id]) ? store.catalog[p.id] : null;

    // 2. Buscar por ID o número de parte / equivalencias / alias con findPart
    if (!item && p.id) {
      item = findPart(p.id);
    }
    if (!item && parte) {
      item = findPart(parte);
    }

    // 3. Buscar con coincidencia inteligente avanzada (sinónimos de taller, lematización, descripción)
    if (!item && (parte || p.descripcion)) {
      const matchRes = matchItemToCatalog(parte, p.descripcion, p.marca, null, false);
      if (matchRes && matchRes.match) {
        item = matchRes.match;
      }
    }

    const costo = Number(p.costo_unitario) || 0;

    // En entradas no se registran servicios externos; son estrictamente refacciones y materiales físicos
    const isSrv = false;
    const esPasoDirecto = Boolean(p.paso_directo || p.es_paso_directo);
    const unidadDirecta = String(p.unidad_camion || '').trim().toUpperCase();

    const almId = isSrv ? 'servicios_externos' : (p.almacen_id || req.body.almacen_id || (item ? item.almacen_id : 'refacciones'));
    const almObj = (store.almacenes || []).find(a => a.id === almId);
    const almNombre = isSrv ? 'SERVICIOS EXTERNOS' : (almObj ? almObj.nombre : (p.almacen || req.body.almacen || (item ? item.almacen : 'Almacén de Refacciones')));

    const { equivalencias: itemEquivs, alias: itemAlias } = extractEquivalenciasAndAlias(p);

    if (!item) {
      const safeId = (almId !== 'refacciones' ? `${almId}_` : '') + (parte || (isSrv ? `SRV_${Date.now()}_${procesados}` : `P_${Date.now()}_${procesados}`)).replace(/[^a-zA-Z0-9_-]/g, '_');
      item = {
        id: safeId,
        rack: isSrv ? 'SERVICIO' : (esPasoDirecto ? 'PASO DIRECTO' : (p.rack || 'Rack 1')),
        nivel: isSrv ? 'EXTERNO' : (esPasoDirecto ? (unidadDirecta ? `UNIDAD ${unidadDirecta}` : 'TALLER') : (p.nivel || 'Nivel 1')),
        descripcion: p.descripcion || (isSrv ? 'Servicio Externo' : 'Sin descripción'),
        marca: p.marca || (isSrv ? 'SERVICIO' : 'GENERICA'),
        numero_parte: parte || (isSrv ? `SRV-${Math.floor(100 + Math.random() * 900)}` : 'S/N'),
        equivalencias: itemEquivs,
        alias: itemAlias,
        unidad: isSrv ? 'SRV' : (p.unidad || 'PZ'),
        precio_unitario: costo,
        stock_inicial: 0,
        stock_actual: (isSrv || esPasoDirecto) ? 0 : cant,
        tipo: isSrv ? 'SERVICIO' : 'PRODUCTO',
        almacen: almNombre,
        almacen_id: almId
      };
      store.catalog[safeId] = item;
    } else {
      mergeEquivalenciasAndAlias(item, itemEquivs, itemAlias);
      const parteUp = parte.toUpperCase();
      if (parteUp && parteUp !== String(item.numero_parte || '').toUpperCase() && !['S/N', 'SN', 'N/A', 'NA', 'SIN NUMERO'].includes(parteUp)) {
        if (!Array.isArray(item.equivalencias)) item.equivalencias = [];
        if (!item.equivalencias.includes(parteUp)) {
          item.equivalencias.push(parteUp);
          console.log(`[CESS EQUIVALENCIAS] Código "${parteUp}" aprendido e indexado para "${item.descripcion}" (${item.numero_parte})`);
        }
      }

      const prevStock = Number(item.stock_actual) || 0;
      if (!isSrv && !esPasoDirecto) {
        item.stock_actual = prevStock + cant;
        if (prevStock < 0) {
          console.log(`[CESS RECONCILIACIÓN] Stock de "${item.numero_parte}" pasó de ${prevStock} a ${item.stock_actual} PZ tras entrada`);
        }
        updateExcelPieceLocationAndStock(item, item.rack, item.nivel, item.stock_actual);
      }
      if (costo > 0) item.precio_unitario = costo;
    }

    const mov = {
      id: `mov_${Date.now()}_${procesados}`,
      tipo: (isSrv && unidadDirecta) ? 'E/S' : (esPasoDirecto ? 'E/S' : 'E'),
      tipo_item: isSrv ? 'SERVICIO' : 'REFACCION',
      fecha,
      hora,
      id_pieza: item.id,
      numero_parte: item.numero_parte,
      descripcion: item.descripcion,
      marca: item.marca,
      almacen: item.almacen || almNombre,
      almacen_id: item.almacen_id || almId,
      ubicacion: isSrv ? `SERVICIO EXTERNO ➔ ${unidadDirecta ? `UNIDAD ${unidadDirecta}` : 'TALLER'}` : (esPasoDirecto ? `PASO DIRECTO ➔ UNIDAD ${unidadDirecta || 'TALLER'}` : `${item.rack} / ${item.nivel}`),
      cantidad: cant,
      proveedor: String(proveedor || p.proveedor || '').trim(),
      factura: String(factura || p.factura || '').trim(),
      orden_compra: ocFolio || String(p.orden_compra || '').trim().toUpperCase(),
      entregado_por: String(entregado_por || p.entregado_por || '').trim(),
      unidad_camion: unidadDirecta,
      folio: String(p.folio || p.folio_mantenimiento || factura || '').trim(),
      folio_mantenimiento: String(p.folio_mantenimiento || p.folio || '').trim(),
      solicitante: String(p.solicitante || entregado_por || 'TALLER').trim(),
      costo_unitario: costo,
      condicion_pago: condicionPago,
      dias_credito: diasCredito,
      fecha_factura: fechaFactura,
      fecha_vencimiento: fechaVencimiento,
      estado_pago: estadoPago,
      fecha_pago: fechaPago,
      referencia_pago: referenciaPago,
      notas: isSrv ? `Servicio externo: ${item.descripcion}${unidadDirecta ? ` aplicado a Unidad ${unidadDirecta}` : ''}` : (esPasoDirecto ? `Paso Directo E/S a Unidad ${unidadDirecta || 'TALLER'}` : 'Ingreso masivo por remisión/factura'),
      timestamp: now.toISOString()
    };

    store.movimientos.unshift(mov);
    procesados++;
  }

  // 2. Registrar en store.facturas_procesadas para control anti-duplicados permanente
  if (factura && !isGenericInvoice(factura)) {
    const totalCant = piezas.reduce((sum, p) => sum + (Number(p.cantidad) || 0), 0);
    const totalMonto = piezas.reduce((sum, p) => sum + ((Number(p.cantidad) || 0) * (Number(p.costo_unitario) || 0)), 0);

    store.facturas_procesadas.unshift({
      id: `fac_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
      factura: String(factura).trim(),
      factura_normalizada: normalizeFacturaKey(factura),
      orden_compra: ocFolio,
      proveedor: String(proveedor || '').trim(),
      proveedor_normalizado: normalizeProveedorKey(proveedor),
      fecha_factura: fechaFactura,
      fecha_ingreso: fecha,
      hora_ingreso: hora,
      total_monto: totalMonto,
      total_piezas: totalCant,
      total_items: piezas.length,
      condicion_pago: condicionPago,
      dias_credito: diasCredito,
      fecha_vencimiento: fechaVencimiento,
      estado_pago: estadoPago,
      fecha_pago: fechaPago,
      referencia_pago: referenciaPago,
      notas_pago: req.body.notas_pago || '',
      image_sha256: imageHash || null,
      entregado_por: String(entregado_por || '').trim(),
      piezas_resumen: piezas.slice(0, 10).map(p => ({
        numero_parte: p.numero_parte,
        cantidad: p.cantidad,
        descripcion: p.descripcion,
        costo_unitario: p.costo_unitario
      }))
    });
  }

  saveStore();
  res.json({
    success: true,
    count: procesados,
    message: `Se procesaron ${procesados} refacciones con éxito. Factura ${factura || ''} registrada.`,
    factura_registrada: factura || null
  });
  } catch (err) {
    console.error('[ENTRADA MASIVA ERROR]', err);
    res.status(500).json({ success: false, error: err.message || 'Error interno al procesar entrada masiva' });
  }
});

// 7.1 Consultar historial de Facturas Registradas en Almacén
app.get('/api/facturas-registradas', (req, res) => {
  res.json({
    success: true,
    total: (store.facturas_procesadas || []).length,
    facturas: store.facturas_procesadas || []
  });
});

// 7.2 Verificar en tiempo real si un folio/proveedor ya existe
app.get('/api/verificar-factura', (req, res) => {
  const { factura, proveedor } = req.query;
  const existing = findExistingFactura(factura, proveedor);
  if (existing) {
    return res.json({
      duplicada: true,
      mensaje: `La factura "${existing.match.factura}" ya fue registrada el ${existing.match.fecha_ingreso} (${existing.match.total_piezas || 'múltiples'} piezas).`,
      registro_previo: existing.match
    });
  }
  return res.json({ duplicada: false });
});

// =========================================================
// NUEVOS ENDPOINTS: SALIDA EN LOTE, EDICIÓN Y ANULACIÓN DE MOVIMIENTOS
// =========================================================

// Registrar salidas múltiples en lote en una sola confirmación
app.post('/api/salida-lote', (req, res) => {
  try {
    if (req.body.usuarioRol === 'GERENTE' || (req.body.permisos && req.body.permisos.dar_salida === false)) {
      return res.status(403).json({ error: 'Acceso denegado: Tu perfil es de Solo Consulta / Auditoría.' });
    }

    const { items, unidad_camion, folio, solicitante, motivo, fecha, entregado_por } = req.body || {};
    if (!Array.isArray(items) || items.length === 0) {
      return res.status(400).json({ error: 'No se recibieron refacciones para procesar la salida' });
    }

    const uTarget = String(unidad_camion || '').trim().toUpperCase();
    if (!uTarget) {
      return res.status(400).json({ error: 'Debes indicar la Unidad o tractocamión destino' });
    }

    // Validar candado de precio > 0 en todas las refacciones
    for (let i = 0; i < items.length; i++) {
      const it = items[i];
      const pz = it.id_pieza ? (store.catalog[it.id_pieza] || findPart(it.id_pieza)) : null;
      const c = Number(it.costo_unitario !== undefined && it.costo_unitario !== '' ? it.costo_unitario : (pz ? pz.precio_unitario : 0));
      if (!c || c <= 0) {
        return res.status(400).json({
          error: 'CANDADO_PRECIO_REQUERIDO',
          message: `Candado de Auditoría: La refacción "${it.descripcion || (pz ? pz.descripcion : 'Ítem ' + (i+1))}" tiene costo $0.00. Ingresa el precio unitario real para registrar la salida.`
        });
      }
    }

    const now = new Date();
    const fSalida = fecha ? String(fecha).trim() : now.toISOString().slice(0, 10);
    const hSalida = now.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    const fFolio = String(folio || '').trim().toUpperCase();
    const fSol = String(solicitante || 'TALLER').trim().toUpperCase();
    const fMot = String(motivo || 'Salida múltiple de refacciones').trim();

    const creados = [];

    for (let i = 0; i < items.length; i++) {
      const it = items[i];
      const cant = Number(it.cantidad) || 1;
      let item = it.id_pieza ? (store.catalog[it.id_pieza] || findPart(it.id_pieza)) : null;
      if (!item && it.numero_parte) item = findPart(it.numero_parte);

      const costo = Number(it.costo_unitario !== undefined ? it.costo_unitario : (item ? item.precio_unitario : 0)) || 0;

      if (item) {
        item.stock_actual = (Number(item.stock_actual) || 0) - cant;
        if (costo > 0 && (!item.precio_unitario || Number(item.precio_unitario) <= 0)) {
          item.precio_unitario = costo;
        }
      }

      const mov = {
        id: `mov_${Date.now()}_${i}_${Math.floor(100 + Math.random() * 900)}`,
        tipo: 'S',
        fecha: fSalida,
        hora: hSalida,
        id_pieza: item ? item.id : (it.id_pieza || 'S/N'),
        numero_parte: item ? item.numero_parte : (it.numero_parte || 'S/N'),
        descripcion: item ? item.descripcion : (it.descripcion || 'Refacción'),
        marca: item ? item.marca : (it.marca || 'GENÉRICA'),
        almacen: item ? (item.almacen || 'Almacén de Refacciones') : 'Almacén de Refacciones',
        almacen_id: item ? (item.almacen_id || 'refacciones') : 'refacciones',
        ubicacion: item ? `${item.rack} / ${item.nivel}` : 'Mostrador',
        cantidad: cant,
        unidad_camion: uTarget,
        folio: fFolio,
        solicitante: fSol,
        costo_unitario: costo,
        costo_total: costo * cant,
        entregado_por: String(entregado_por || '').trim(),
        notas: fMot,
        timestamp: now.toISOString()
      };

      store.movimientos.unshift(mov);
      creados.push(mov);
    }

    saveStore();
    try { creados.forEach(m => syncSalidaConTaller(m)); } catch(e) {}
    res.json({
      success: true,
      count: creados.length,
      message: `¡Éxito! Se registraron ${creados.length} refacciones con salida para la Unidad ${uTarget}`,
      movimientos: creados
    });
  } catch (err) {
    console.error('Error en salida-lote:', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// Modificar datos de un movimiento registrado en la bitácora
app.put('/api/movimientos/:id', (req, res) => {
  try {
    if (req.body.usuarioRol === 'GERENTE' || (req.body.permisos && req.body.permisos.auditar === false)) {
      return res.status(403).json({ error: 'Acceso denegado: No tienes permisos para modificar movimientos.' });
    }

    const { id } = req.params;
    const mov = (store.movimientos || []).find(m => m.id === id);
    if (!mov) return res.status(404).json({ success: false, error: 'Movimiento no encontrado' });

    const prevCant = Number(mov.cantidad) || 0;
    const newCant = req.body.cantidad !== undefined ? Number(req.body.cantidad) : prevCant;
    const prevCosto = Number(mov.costo_unitario) || 0;
    const newCosto = req.body.costo_unitario !== undefined ? Number(req.body.costo_unitario) : prevCosto;

    // Si cambió la cantidad y el movimiento tiene pieza en catálogo, reconciliar stock
    if (newCant !== prevCant && mov.id_pieza && store.catalog[mov.id_pieza]) {
      const diff = newCant - prevCant;
      const cat = store.catalog[mov.id_pieza];
      if (mov.tipo === 'E') {
        cat.stock_actual = (Number(cat.stock_actual) || 0) + diff;
      } else if (mov.tipo === 'S') {
        cat.stock_actual = (Number(cat.stock_actual) || 0) - diff;
      }
    }

    if (req.body.descripcion !== undefined) mov.descripcion = String(req.body.descripcion).trim();
    if (req.body.numero_parte !== undefined) mov.numero_parte = String(req.body.numero_parte).trim();
    if (req.body.marca !== undefined) mov.marca = String(req.body.marca).trim();
    if (req.body.unidad_camion !== undefined) mov.unidad_camion = String(req.body.unidad_camion).trim().toUpperCase();
    if (req.body.folio !== undefined) mov.folio = String(req.body.folio).trim().toUpperCase();
    if (req.body.orden_compra !== undefined) mov.orden_compra = String(req.body.orden_compra).trim().toUpperCase();
    if (req.body.solicitante !== undefined) mov.solicitante = String(req.body.solicitante).trim().toUpperCase();
    if (req.body.proveedor !== undefined) mov.proveedor = String(req.body.proveedor).trim();
    if (req.body.factura !== undefined) mov.factura = String(req.body.factura).trim();
    if (req.body.notas !== undefined) mov.notas = String(req.body.notas).trim();

    mov.cantidad = newCant;
    mov.costo_unitario = newCosto;
    mov.costo_total = newCant * newCosto;
    mov.editado = true;
    mov.fecha_edicion = new Date().toISOString();
    mov.editado_por = req.body.usuario || 'ADMIN';

    // Actualizar precio de catálogo si la pieza estaba en 0
    if (mov.id_pieza && store.catalog[mov.id_pieza] && newCosto > 0 && (!store.catalog[mov.id_pieza].precio_unitario || store.catalog[mov.id_pieza].precio_unitario <= 0)) {
      store.catalog[mov.id_pieza].precio_unitario = newCosto;
    }

    saveStore();
    res.json({ success: true, message: 'Movimiento actualizado correctamente', movimiento: mov });
  } catch (err) {
    console.error('Error editando movimiento:', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// Anular / Revertir un movimiento (restituye el stock al inventario)
app.post('/api/movimientos/:id/anular', (req, res) => {
  try {
    if (req.body.usuarioRol === 'GERENTE' || (req.body.permisos && req.body.permisos.auditar === false)) {
      return res.status(403).json({ error: 'Acceso denegado: No tienes permisos para anular movimientos.' });
    }

    const { id } = req.params;
    const mov = (store.movimientos || []).find(m => m.id === id);
    if (!mov) return res.status(404).json({ success: false, error: 'Movimiento no encontrado' });

    if (mov.anulado) {
      return res.status(400).json({ success: false, error: 'Este movimiento ya fue anulado previamente.' });
    }

    const cant = Number(mov.cantidad) || 0;
    let itemCat = mov.id_pieza ? (store.catalog[mov.id_pieza] || findPart(mov.id_pieza) || findPart(mov.numero_parte)) : null;

    if (itemCat) {
      if (mov.tipo === 'E') {
        // Era una entrada errónea: restamos la cantidad del stock
        itemCat.stock_actual = Math.max(0, (Number(itemCat.stock_actual) || 0) - cant);
      } else if (mov.tipo === 'S') {
        // Era una salida errónea: devolvemos la cantidad al stock
        itemCat.stock_actual = (Number(itemCat.stock_actual) || 0) + cant;
      }
    }

    mov.anulado = true;
    mov.tipo_original = mov.tipo;
    mov.tipo = 'ANULADO';
    mov.motivo_anulacion = String(req.body.motivo || 'Anulación por corrección de almacén').trim();
    mov.anulado_por = String(req.body.usuario || 'ADMIN').trim();
    mov.fecha_anulacion = new Date().toISOString();

    saveStore();
    res.json({
      success: true,
      message: `Movimiento anulado con éxito. Se reconcilió el stock de "${mov.descripcion}" (${mov.tipo_original === 'E' ? '-' : '+'}${cant} PZ).`,
      movimiento: mov,
      nuevo_stock: itemCat ? itemCat.stock_actual : null
    });
  } catch (err) {
    console.error('Error anulando movimiento:', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// Eliminar refacción del catálogo si fue creada por error
app.delete('/api/productos/:id', (req, res) => {
  try {
    if (req.body.usuarioRol === 'GERENTE' || (req.body.permisos && req.body.permisos.dar_salida === false)) {
      return res.status(403).json({ error: 'Acceso denegado: No tienes permisos para eliminar refacciones.' });
    }

    const { id } = req.params;
    if (!store.catalog[id]) {
      return res.status(404).json({ success: false, error: 'Refacción no encontrada en catálogo' });
    }

    const pz = store.catalog[id];
    delete store.catalog[id];
    saveStore();

    res.json({ success: true, message: `Refacción "${pz.descripcion}" eliminada del catálogo.` });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});


// 8. Bitácora de movimientos (con filtros)
app.get('/api/bitacora', (req, res) => {
  const { fecha, tipo, unidad, busqueda } = req.query;
  let list = store.movimientos;

  if (fecha) {
    list = list.filter(m => m.fecha === fecha);
  }
  if (tipo && tipo !== 'TODOS') {
    list = list.filter(m => m.tipo === tipo);
  }
  if (unidad) {
    list = list.filter(m => m.unidad_camion && m.unidad_camion.includes(unidad.toUpperCase()));
  }
  if (busqueda) {
    const q = busqueda.toLowerCase();
    list = list.filter(m =>
      (m.descripcion && m.descripcion.toLowerCase().includes(q)) ||
      (m.numero_parte && m.numero_parte.toLowerCase().includes(q)) ||
      (m.folio && m.folio.toLowerCase().includes(q)) ||
      (m.solicitante && m.solicitante.toLowerCase().includes(q)) ||
      (m.proveedor && m.proveedor.toLowerCase().includes(q))
    );
  }

  res.json({ total: list.length, movimientos: list });
});

// 8.1 Editar movimiento de Bitácora (corrección de fecha, hora, folio, solicitante, etc.)
app.put('/api/bitacora/:id', (req, res) => {
  const { id } = req.params;
  const {
    fecha,
    hora,
    folio,
    factura,
    solicitante,
    entregado_por,
    unidad_camion,
    proveedor,
    motivo,
    notas,
    usuarioRol,
    permisos
  } = req.body;

  if (usuarioRol === 'GERENTE' || (permisos && permisos.admin_centro_control === false && permisos.dar_entrada === false && permisos.dar_salida === false)) {
    return res.status(403).json({ error: 'No tienes permisos para editar registros de la bitácora.' });
  }

  const mov = (store.movimientos || []).find(m => String(m.id) === String(id));
  if (!mov) {
    return res.status(404).json({ error: 'Movimiento no encontrado en la bitácora.' });
  }

  // Validar y actualizar fecha
  if (fecha) {
    const fStr = String(fecha).trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(fStr)) {
      return res.status(400).json({ error: 'Formato de fecha inválido. Debe ser AAAA-MM-DD (ej. 2026-10-20)' });
    }
    mov.fecha = fStr;
  }

  if (hora !== undefined) mov.hora = String(hora).trim();
  if (folio !== undefined) mov.folio = String(folio).trim();
  if (factura !== undefined) mov.factura = String(factura).trim();
  if (solicitante !== undefined) mov.solicitante = String(solicitante).trim();
  if (entregado_por !== undefined) mov.entregado_por = String(entregado_por).trim();
  if (unidad_camion !== undefined) mov.unidad_camion = String(unidad_camion).trim().toUpperCase();
  if (proveedor !== undefined) mov.proveedor = String(proveedor).trim().toUpperCase();
  if (motivo !== undefined) mov.motivo = String(motivo).trim();
  if (notas !== undefined) mov.notas = String(notas).trim();

  mov.editado_el = new Date().toISOString();

  // Si tiene factura asociada y se cambió la fecha, actualizar en facturas_procesadas
  if (fecha && (mov.tipo === 'E' || mov.tipo === 'E/S') && mov.factura && !isGenericInvoice(mov.factura)) {
    const fNorm = normalizeFacturaKey(mov.factura);
    const pNorm = normalizeProveedorKey(mov.proveedor);
    const facItem = (store.facturas_procesadas || []).find(f =>
      f.factura_normalizada === fNorm && (!pNorm || f.proveedor_normalizado === pNorm)
    );
    if (facItem) {
      facItem.fecha_ingreso = mov.fecha;
      facItem.fecha_factura = mov.fecha;
    }
  }

  saveStore();

  res.json({
    success: true,
    movimiento: mov,
    message: `Movimiento corregido exitosamente. Nueva fecha: ${mov.fecha}`
  });
});

// 8.2 Devolución de refacción a Almacén / Cancelación de Salida (Retorno automático a Rack)
app.post('/api/bitacora/:id/devolucion', (req, res) => {
  const { id } = req.params;
  const {
    cantidad_devuelta,
    motivo_devolucion,
    notas,
    usuarioRol,
    permisos
  } = req.body;

  if (usuarioRol === 'GERENTE' || (permisos && permisos.admin_centro_control === false && permisos.dar_salida === false && permisos.dar_entrada === false)) {
    return res.status(403).json({ error: 'No tienes permisos para realizar devoluciones o anular salidas.' });
  }

  const mov = (store.movimientos || []).find(m => String(m.id) === String(id));
  if (!mov) {
    return res.status(404).json({ error: 'Movimiento no encontrado en la bitácora.' });
  }

  if (mov.anulado) {
    return res.status(400).json({ error: 'Este movimiento ya fue anulado o devuelto previamente.' });
  }

  if (mov.tipo !== 'S' && mov.tipo !== 'E/S') {
    return res.status(400).json({ error: 'Solo se pueden devolver o cancelar salidas de refacciones a taller.' });
  }

  const cantOriginal = Number(mov.cantidad) || 1;
  const cantDevuelta = Math.min(cantOriginal, Math.max(1, Number(cantidad_devuelta) || cantOriginal));

  // Buscar la refacción en catálogo para reincorporar stock y obtener su estante/rack
  const item = store.catalog[mov.id_pieza] || Object.values(store.catalog).find(p => p.numero_parte === mov.numero_parte);

  if (item) {
    item.stock_actual = (Number(item.stock_actual) || 0) + cantDevuelta;
  }

  const now = new Date();
  const fechaHoy = now.toISOString().slice(0, 10);
  const horaHoy = now.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
  const motivoTexto = (motivo_devolucion || 'Pieza errónea / No se utilizó en taller').trim();
  const ubicacionDestino = item ? `${item.rack || 'Rack 1'} / ${item.nivel || 'Nivel 1'}` : (mov.ubicacion || 'Almacén Central');

  let mensaje = '';
  if (cantDevuelta >= cantOriginal) {
    // Devolución Total: Anular salida y tachar de bitácora
    mov.anulado = true;
    mov.cantidad_devuelta = cantDevuelta;
    mov.motivo_anulacion = motivoTexto;
    mov.fecha_devolucion = fechaHoy;
    mov.hora_devolucion = horaHoy;
    mov.notas = (mov.notas ? mov.notas + ' | ' : '') + `[DEVUELTA A ALMACÉN (${cantDevuelta} PZ): ${motivoTexto}]`;
    mensaje = `✅ Salida anulada con éxito: Se reincorporaron ${cantDevuelta} PZ al inventario. Acomodar en: ${ubicacionDestino}`;
  } else {
    // Devolución Parcial: Reducir cantidad de salida y registrar contra-movimiento informativo
    const nuevaCant = cantOriginal - cantDevuelta;
    mov.cantidad = nuevaCant;
    mov.notas = (mov.notas ? mov.notas + ' | ' : '') + `[Se devolvieron ${cantDevuelta} PZ a almacén el ${fechaHoy}: ${motivoTexto}]`;

    const movReingreso = {
      id: `dev_${Date.now()}_${Math.round(Math.random() * 1000)}`,
      tipo: 'DEV',
      fecha: fechaHoy,
      hora: horaHoy,
      id_pieza: mov.id_pieza,
      numero_parte: mov.numero_parte,
      descripcion: `↩️ DEVOLUCIÓN: ${mov.descripcion}`,
      marca: mov.marca,
      ubicacion: ubicacionDestino,
      cantidad: cantDevuelta,
      unidad_camion: mov.unidad_camion,
      folio: mov.folio,
      solicitante: mov.solicitante,
      notas: `Devolución de salida (${mov.folio || mov.id}): ${motivoTexto}`,
      timestamp: now.toISOString()
    };
    store.movimientos.unshift(movReingreso);
    mensaje = `✅ Devolución parcial procesada: Se reincorporaron ${cantDevuelta} PZ a ${ubicacionDestino}. Salida ajustada a ${nuevaCant} PZ.`;
  }

  saveStore();

  res.json({
    success: true,
    movimiento: mov,
    item: item ? {
      id: item.id,
      numero_parte: item.numero_parte,
      descripcion: item.descripcion,
      rack: item.rack,
      nivel: item.nivel,
      stock_actual: item.stock_actual
    } : null,
    ubicacion_acomodo: ubicacionDestino,
    cant_devuelta: cantDevuelta,
    message: mensaje
  });
});

// 9. Subida de Foto (desde cámara móvil o archivo)
app.post('/api/upload', upload.single('foto'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No se subió ningún archivo' });
  }
  const photoUrl = `/uploads/${req.file.filename}`;
  res.json({ success: true, url: photoUrl, filename: req.file.filename });
});

// 10. Exportar a Excel con formato idéntico a INVENTARIO.xlsx (admite cualquier día o cierre atrasado)
app.post('/api/export-excel', (req, res) => {
  try {
    if (!fs.existsSync(EXCEL_FILE)) {
      return res.status(404).json({ error: 'Archivo INVENTARIO.xlsx no encontrado' });
    }

    let targetDateStr = req.body.fecha ? String(req.body.fecha).trim() : '';
    let targetSheetName = req.body.sheetName ? String(req.body.sheetName).trim() : '';

    const today = new Date();
    if (!targetDateStr && !targetSheetName) {
      targetDateStr = today.toISOString().slice(0, 10);
      const day = String(today.getDate()).padStart(2, '0');
      const month = String(today.getMonth() + 1).padStart(2, '0');
      const year = today.getFullYear();
      targetSheetName = `${day}-${month}-${year}`;
    } else if (targetDateStr && !targetSheetName) {
      const parts = targetDateStr.split('-');
      if (parts[0].length === 4) { // YYYY-MM-DD
        targetSheetName = `${parts[2]}-${parts[1]}-${parts[0]}`;
      } else {
        targetSheetName = targetDateStr;
        targetDateStr = `${parts[2]}-${parts[1]}-${parts[0]}`;
      }
    } else if (targetSheetName && !targetDateStr) {
      const parts = targetSheetName.split('-');
      if (parts[0].length === 4) {
        targetDateStr = targetSheetName;
        targetSheetName = `${parts[2]}-${parts[1]}-${parts[0]}`;
      } else {
        targetDateStr = `${parts[2]}-${parts[1]}-${parts[0]}`;
      }
    }

    const workbook = xlsx.readFile(EXCEL_FILE);

    // Encabezados exactos a los que usa la empresa
    const headers = [
      'RACK', 'NIVEL', 'DESCRIPCIÓN ', 'MARCA', 'NÚMERO DE PARTE ',
      'ENTRADAS', 'SALIDAS', 'PROVEEDOR', 'FACTURA',
      'CANTIDAD INICAL', 'UNIDAD', 'Cantidad final',
      'PRECIO UNITARIO SIN IVA', 'VALOR INVENTARIO (MXN)'
    ];

    const rowsData = [headers];

    // Filtrar los items por almacén si se solicita o por defecto Refacciones Diesel para fechas
    const almIdFilter = req.body.almacen_id || (targetSheetName.includes('-') ? 'refacciones' : null);
    const catalogItems = Object.values(store.catalog).filter(item => {
      if (!almIdFilter || almIdFilter === 'TODOS') return true;
      return item.almacen_id === almIdFilter || (almIdFilter === 'refacciones' && !item.almacen_id);
    });

    // Recorrer el catálogo y calcular matemáticamente los balances hasta targetDateStr
    for (const item of catalogItems) {
      // Movimientos previos a targetDateStr
      const prevMovs = store.movimientos.filter(m =>
        !m.anulado &&
        (m.id_pieza === item.id || m.numero_parte === item.numero_parte) &&
        m.fecha < targetDateStr
      );
      const prevEnt = prevMovs.filter(m => m.tipo === 'E' || m.tipo === 'E/S' || m.tipo === 'DEV').reduce((s, m) => s + Number(m.cantidad || 0), 0);
      const prevSal = prevMovs.filter(m => m.tipo === 'S' || m.tipo === 'E/S').reduce((s, m) => s + Number(m.cantidad || 0), 0);

      const stockInicialDia = Number(item.stock_inicial || 0) + prevEnt - prevSal;

      // Movimientos del día exacto
      const dayMovs = store.movimientos.filter(m =>
        !m.anulado &&
        (m.id_pieza === item.id || m.numero_parte === item.numero_parte) &&
        m.fecha === targetDateStr
      );
      const entDia = dayMovs.filter(m => m.tipo === 'E' || m.tipo === 'E/S' || m.tipo === 'DEV').reduce((s, m) => s + Number(m.cantidad || 0), 0);
      const salDia = dayMovs.filter(m => m.tipo === 'S' || m.tipo === 'E/S').reduce((s, m) => s + Number(m.cantidad || 0), 0);

      const prov = dayMovs.find(m => m.proveedor)?.proveedor || '';
      const fact = dayMovs.find(m => m.factura)?.factura || '';

      const stockFinalDia = stockInicialDia + entDia - salDia;
      const precioUnit = Number(item.precio_unitario || 0);
      const valorInventario = stockFinalDia * precioUnit;

      rowsData.push([
        item.rack.replace('Rack ', ''),
        item.nivel,
        item.descripcion,
        item.marca,
        item.numero_parte,
        entDia > 0 ? entDia : null,
        salDia > 0 ? salDia : null,
        prov || null,
        fact || null,
        stockInicialDia,
        item.unidad || 'PZ',
        stockFinalDia,
        precioUnit,
        valorInventario
      ]);
    }

    // Reemplazar hoja si ya existe, o crearla
    if (workbook.Sheets[targetSheetName]) {
      delete workbook.Sheets[targetSheetName];
      const idx = workbook.SheetNames.indexOf(targetSheetName);
      if (idx !== -1) workbook.SheetNames.splice(idx, 1);
    }

    const newWs = xlsx.utils.aoa_to_sheet(rowsData);
    xlsx.utils.book_append_sheet(workbook, newWs, targetSheetName);

    xlsx.writeFile(workbook, EXCEL_FILE);

    res.json({
      success: true,
      message: `Hoja '${targetSheetName}' guardada con éxito en INVENTARIO.xlsx (${rowsData.length - 1} refacciones procesadas).`,
      sheetName: targetSheetName,
      fecha: targetDateStr
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// =========================================================
// 11. REPORTES Y GRÁFICAS EJECUTIVAS PARA TOMA DE DECISIONES
// =========================================================
app.get('/api/reportes', (req, res) => {
  try {
    const { desde, hasta, unidad } = req.query;

    let fDesde = desde || '2026-09-01';
    let fHasta = hasta || '2026-10-31';

    let movs = store.movimientos.filter(m => {
      if (m.anulado) return false;
      if (m.fecha && (m.fecha < fDesde || m.fecha > fHasta)) return false;
      if (unidad && unidad !== 'TODAS' && (m.tipo === 'S' || m.tipo === 'E/S')) {
        return m.unidad_camion === unidad;
      }
      return true;
    });

    let totalGastoSalidas = 0;
    let totalPiezasSalidas = 0;
    let totalInversionEntradas = 0;
    let totalPiezasEntradas = 0;

    const gastoPorUnidad = {};
    const topPiezasSalidas = {};
    const porSolicitante = {};
    const porProveedor = {};
    const tendenciaDiaria = {};
    const distribucionRacks = {};

    const unidadesSet = new Set();
    store.movimientos.forEach(m => {
      if (m.unidad_camion && m.unidad_camion.trim()) {
        unidadesSet.add(m.unidad_camion.trim().toUpperCase());
      }
    });

    for (const m of movs) {
      const itemCat = store.catalog[m.id_pieza] ||
        Object.values(store.catalog).find(p => p.numero_parte === m.numero_parte) || {};

      const precioUnit = Number(itemCat.precio_unitario || m.costo_unitario || 0);
      const rack = itemCat.rack || 'Sin Rack';
      const fechaKey = m.fecha || 'Sin Fecha';

      if (!tendenciaDiaria[fechaKey]) {
        tendenciaDiaria[fechaKey] = {
          fecha: fechaKey,
          gasto_salidas: 0,
          inversion_entradas: 0,
          piezas_salidas: 0,
          piezas_entradas: 0
        };
      }

      // Procesamiento de Salida (para 'S' y 'E/S')
      if (m.tipo === 'S' || m.tipo === 'E/S') {
        const cant = Number(m.cantidad || 0);
        const monto = cant * precioUnit;

        totalPiezasSalidas += cant;
        totalGastoSalidas += monto;

        tendenciaDiaria[fechaKey].gasto_salidas += monto;
        tendenciaDiaria[fechaKey].piezas_salidas += cant;

        const un = (m.unidad_camion || 'SIN UNIDAD').toUpperCase().trim();
        if (!gastoPorUnidad[un]) {
          gastoPorUnidad[un] = { unidad: un, monto: 0, piezas: 0, movimientos: 0, refacciones: {} };
        }
        gastoPorUnidad[un].monto += monto;
        gastoPorUnidad[un].piezas += cant;
        gastoPorUnidad[un].movimientos += 1;
        const refKey = m.numero_parte || m.descripcion || 'Sin Código';
        gastoPorUnidad[un].refacciones[refKey] = (gastoPorUnidad[un].refacciones[refKey] || 0) + cant;

        const pKey = m.id_pieza || m.numero_parte || m.descripcion;
        if (!topPiezasSalidas[pKey]) {
          topPiezasSalidas[pKey] = {
            id_pieza: m.id_pieza,
            numero_parte: m.numero_parte || itemCat.numero_parte || '',
            descripcion: m.descripcion || itemCat.descripcion || '',
            marca: m.marca || itemCat.marca || '',
            rack: rack,
            nivel: itemCat.nivel || '',
            cantidad: 0,
            monto: 0,
            precio_unitario: precioUnit
          };
        }
        topPiezasSalidas[pKey].cantidad += cant;
        topPiezasSalidas[pKey].monto += monto;

        const sol = (m.solicitante || 'Taller General').trim();
        if (!porSolicitante[sol]) {
          porSolicitante[sol] = { solicitante: sol, monto: 0, piezas: 0, salidas: 0 };
        }
        porSolicitante[sol].monto += monto;
        porSolicitante[sol].piezas += cant;
        porSolicitante[sol].salidas += 1;

        if (!distribucionRacks[rack]) {
          distribucionRacks[rack] = { rack, monto: 0, piezas: 0 };
        }
        distribucionRacks[rack].monto += monto;
        distribucionRacks[rack].piezas += cant;
      }

      // Procesamiento de Entrada (para 'E' y 'E/S')
      if (m.tipo === 'E' || m.tipo === 'E/S') {
        const cant = Number(m.cantidad || 0);
        const costo = Number(m.costo_unitario || precioUnit || 0);
        const monto = cant * costo;

        totalPiezasEntradas += cant;
        totalInversionEntradas += monto;

        tendenciaDiaria[fechaKey].inversion_entradas += monto;
        tendenciaDiaria[fechaKey].piezas_entradas += cant;

        const prov = (m.proveedor || 'PROVEEDOR LOCAL').trim().toUpperCase();
        if (!porProveedor[prov]) {
          porProveedor[prov] = { proveedor: prov, monto: 0, piezas: 0, entradas: 0 };
        }
        porProveedor[prov].monto += monto;
        porProveedor[prov].piezas += cant;
        porProveedor[prov].entradas += 1;
      }
    }

    const rankingUnidades = Object.values(gastoPorUnidad)
      .map(u => ({
        ...u,
        top_refacciones: Object.entries(u.refacciones)
          .sort((a, b) => b[1] - a[1])
          .slice(0, 3)
          .map(([ref, c]) => `${ref} (${c})`)
          .join(', ')
      }))
      .sort((a, b) => b.monto - a.monto);

    const rankingPiezas = Object.values(topPiezasSalidas)
      .sort((a, b) => b.monto - a.monto);

    const rankingProveedores = Object.values(porProveedor)
      .sort((a, b) => b.monto - a.monto);

    const rankingSolicitantes = Object.values(porSolicitante)
      .sort((a, b) => b.monto - a.monto);

    const listaTendencia = Object.values(tendenciaDiaria)
      .sort((a, b) => a.fecha.localeCompare(b.fecha));

    // 1. Cálculo de Inventario Actual en Vivo
    let valorInventarioEnVivo = 0;
    let piezasTotalesEnVivo = 0;
    const piezasBajoStock = [];
    for (const p of Object.values(store.catalog)) {
      const stock = Number(p.stock_actual || 0);
      const pr = Number(p.precio_unitario || 0);
      valorInventarioEnVivo += (stock * pr);
      piezasTotalesEnVivo += stock;
      if (stock <= 1) {
        piezasBajoStock.push({
          id: p.id,
          numero_parte: p.numero_parte,
          descripcion: p.descripcion,
          marca: p.marca,
          rack: p.rack,
          nivel: p.nivel,
          stock_actual: stock,
          precio_unitario: pr
        });
      }
    }

    // 2. Cálculo del Inventario Inicial al arrancar el sistema
    let valorInventarioOrigen = 0;
    let piezasTotalesOrigen = 0;
    for (const p of Object.values(store.catalog)) {
      const stIni = Number(p.stock_inicial !== undefined ? p.stock_inicial : (p.stock_actual || 0));
      const pr = Number(p.precio_unitario || 0);
      valorInventarioOrigen += (stIni * pr);
      piezasTotalesOrigen += stIni;
    }

    // 3. Función de reconstrucción histórica de inventario a cualquier fecha
    function calcularInventarioAFecha(fechaCorte) {
      const stocks = {};
      for (const [id, p] of Object.entries(store.catalog)) {
        stocks[id] = Number(p.stock_actual || 0);
      }
      for (const m of store.movimientos) {
        if (m.anulado) continue;
        if (m.fecha && m.fecha > fechaCorte) {
          const p = store.catalog[m.id_pieza] || Object.values(store.catalog).find(it => it.numero_parte === m.numero_parte);
          if (!p) continue;
          const id = p.id;
          const cant = Number(m.cantidad || 0);
          if (m.tipo === 'E' || m.tipo === 'DEV') {
            stocks[id] = (stocks[id] || 0) - cant;
          } else if (m.tipo === 'S') {
            stocks[id] = (stocks[id] || 0) + cant;
          } else if (m.tipo === 'AJUSTE') {
            if (m.diferencia_stock !== undefined) {
              stocks[id] = (stocks[id] || 0) - Number(m.diferencia_stock);
            }
          }
        }
      }
      let val = 0;
      let pzs = 0;
      for (const [id, st] of Object.entries(stocks)) {
        const pr = Number(store.catalog[id]?.precio_unitario || 0);
        val += (st * pr);
        pzs += st;
      }
      return { valor: val, piezas: pzs };
    }

    // Snapshot al inicio y cierre del periodo filtrado
    const fechaMinMov = store.movimientos.map(m => m.fecha).filter(Boolean).sort()[0] || '2026-09-22';
    const hoyStr = new Date().toISOString().slice(0, 10);

    const snapInicio = (fDesde < fechaMinMov)
      ? { valor: valorInventarioOrigen, piezas: piezasTotalesOrigen }
      : calcularInventarioAFecha(fDesde);

    const snapCorte = (fHasta >= hoyStr)
      ? { valor: valorInventarioEnVivo, piezas: piezasTotalesEnVivo }
      : calcularInventarioAFecha(fHasta);

    const variacionPeriodoMonto = snapCorte.valor - snapInicio.valor;
    const variacionPeriodoPct = snapInicio.valor > 0 ? (variacionPeriodoMonto / snapInicio.valor) * 100 : 0;

    const variacionGlobalMonto = valorInventarioEnVivo - valorInventarioOrigen;
    const variacionGlobalPct = valorInventarioOrigen > 0 ? (variacionGlobalMonto / valorInventarioOrigen) * 100 : 0;

    // Enriquecer tendencia_diaria con evolución del inventario acumulado
    for (const t of listaTendencia) {
      const snapDia = calcularInventarioAFecha(t.fecha);
      t.valor_inventario = snapDia.valor;
      t.piezas_stock = snapDia.piezas;
      t.balance_diario = (t.inversion_entradas || 0) - (t.gasto_salidas || 0);
    }

    // Cálculo detallado de Auditoría de Variación de Inventario (Causas de incremento o decremento)
    const piezasVariacionMap = {};
    for (const m of movs) {
      if (m.anulado) continue;
      const p = store.catalog[m.id_pieza] || Object.values(store.catalog).find(it => it.numero_parte === m.numero_parte) || {};
      const id = p.id || m.numero_parte || m.descripcion;
      const desc = p.descripcion || m.descripcion || 'Sin descripción';
      const np = p.numero_parte || m.numero_parte || 'S/N';
      const pr = Number(m.costo_unitario || p.precio_unitario || 0);
      const cant = Number(m.cantidad || 0);

      if (!piezasVariacionMap[id]) {
        piezasVariacionMap[id] = {
          id,
          numero_parte: np,
          descripcion: desc,
          marca: p.marca || m.marca || '',
          precio_unitario: pr,
          entradas_cant: 0,
          entradas_monto: 0,
          salidas_cant: 0,
          salidas_monto: 0,
          ajustes_cant: 0,
          ajustes_monto: 0,
          delta_cant: 0,
          delta_monto: 0
        };
      }

      if (m.tipo === 'E') {
        piezasVariacionMap[id].entradas_cant += cant;
        piezasVariacionMap[id].entradas_monto += (cant * pr);
        piezasVariacionMap[id].delta_cant += cant;
        piezasVariacionMap[id].delta_monto += (cant * pr);
      } else if (m.tipo === 'S') {
        piezasVariacionMap[id].salidas_cant += cant;
        piezasVariacionMap[id].salidas_monto += (cant * pr);
        piezasVariacionMap[id].delta_cant -= cant;
        piezasVariacionMap[id].delta_monto -= (cant * pr);
      } else if (m.tipo === 'AJUSTE') {
        const diff = m.diferencia_stock !== undefined ? Number(m.diferencia_stock) : 0;
        piezasVariacionMap[id].ajustes_cant += diff;
        piezasVariacionMap[id].ajustes_monto += (diff * pr);
        piezasVariacionMap[id].delta_cant += diff;
        piezasVariacionMap[id].delta_monto += (diff * pr);
      }
    }

    const listaVariaciones = Object.values(piezasVariacionMap);
    const topIncrementos = listaVariaciones.filter(p => p.delta_monto > 0).sort((a, b) => b.delta_monto - a.delta_monto);
    const topDecrementos = listaVariaciones.filter(p => p.delta_monto < 0).sort((a, b) => a.delta_monto - b.delta_monto);

    res.json({
      success: true,
      periodo: { desde: fDesde, hasta: fHasta },
      auditoria_variacion: {
        top_incrementos: topIncrementos.slice(0, 30),
        top_decrementos: topDecrementos.slice(0, 30),
        total_incremento_monto: topIncrementos.reduce((s, p) => s + p.delta_monto, 0),
        total_decremento_monto: Math.abs(topDecrementos.reduce((s, p) => s + p.delta_monto, 0))
      },
      kpis: {
        total_gasto_salidas: totalGastoSalidas,
        total_piezas_salidas: totalPiezasSalidas,
        total_inversion_entradas: totalInversionEntradas,
        total_piezas_entradas: totalPiezasEntradas,
        valor_inventario_total: snapCorte.valor,
        valor_inventario_al_corte: snapCorte.valor,
        piezas_inventario_al_corte: snapCorte.piezas,
        valor_inventario_inicio_periodo: snapInicio.valor,
        piezas_inventario_inicio_periodo: snapInicio.piezas,
        valor_inventario_actual_en_vivo: valorInventarioEnVivo,
        piezas_inventario_actual_en_vivo: piezasTotalesEnVivo,
        valor_inventario_origen: valorInventarioOrigen,
        piezas_inventario_origen: piezasTotalesOrigen,
        variacion_periodo_monto: variacionPeriodoMonto,
        variacion_periodo_porcentaje: variacionPeriodoPct,
        variacion_global_monto: variacionGlobalMonto,
        variacion_global_porcentaje: variacionGlobalPct,
        piezas_bajo_stock_count: piezasBajoStock.length,
        unidad_top: rankingUnidades[0] || null,
        proveedor_top: rankingProveedores[0] || null
      },
      ranking_unidades: rankingUnidades,
      ranking_piezas: rankingPiezas.slice(0, 25),
      ranking_proveedores: rankingProveedores,
      ranking_solicitantes: rankingSolicitantes,
      tendencia_diaria: listaTendencia,
      distribucion_racks: Object.values(distribucionRacks),
      piezas_criticas: piezasBajoStock.slice(0, 25),
      unidades_disponibles: Array.from(unidadesSet).sort()
    });
  } catch (err) {
    console.error('Error al generar reportes:', err);
    res.status(500).json({ error: err.message });
  }
});

// Helper robusto para consultar la API de Gemini Vision con cascada de modelos multimodales y reintentos ante contingencias
async function callGeminiVision(apiKey, parts) {
  // Modelos oficiales vigentes prioritarios para visión y procesamiento multimodal
  let modelsToTry = [
    'gemini-3.5-flash',
    'gemini-3.5-flash-lite',
    'gemini-3.1-flash-lite',
    'gemini-3-flash-preview',
    'gemini-3.8-flash'
  ];

  // Descubrimiento dinámico: consultar a Google qué modelos están activos y autorizados en el proyecto de este API Key
  try {
    const listRes = await fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`);
    if (listRes.ok) {
      const listData = await listRes.json();
      const available = (listData.models || [])
        .filter(m => Array.isArray(m.supportedGenerationMethods) && m.supportedGenerationMethods.includes('generateContent'))
        .map(m => m.name.replace(/^models\//, ''));

      console.log('[Gemini Vision] Modelos activos reportados por Google para este API Key:', available);

      const dynamicList = [];
      const priorityOrder = [
        'gemini-3.5-flash',
        'gemini-3.5-flash-lite',
        'gemini-3.1-flash-lite',
        'gemini-3-flash-preview',
        'gemini-3.8-flash'
      ];

      for (const pref of priorityOrder) {
        if (available.includes(pref)) {
          dynamicList.push(pref);
        }
      }

      for (const m of available) {
        if (!dynamicList.includes(m) && (m.includes('flash') || m.includes('pro')) && !m.includes('embedding') && !m.includes('imagen')) {
          dynamicList.push(m);
        }
      }

      if (dynamicList.length > 0) {
        modelsToTry = dynamicList;
        console.log('[Gemini Vision] Cascada optimizada a ejecutar:', modelsToTry);
      }
    }
  } catch (discErr) {
    console.warn('[Gemini Vision] Consulta dinámica de modelos omitida:', discErr.message);
  }

  let lastErrorDetail = '';
  let lastStatus = 500;
  let sawHighDemand = false;
  let sawQuotaExhausted = false;
  let allQuotaExhausted = true;

  for (const model of modelsToTry) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
    const payload = {
      contents: [{ parts }],
      generationConfig: {
        response_mime_type: 'application/json',
        temperature: 0.1
      }
    };

    // Máximo 1 reintento rápido por modelo ante errores transitorios de red o saturación
    const MAX_RETRIES = 1;
    for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
      if (attempt > 0) {
        const backoffMs = 1200 + Math.floor(Math.random() * 400);
        console.log(`[Gemini Vision ${model}] Reintento rápido en ${backoffMs}ms...`);
        await new Promise(r => setTimeout(r, backoffMs));
      }

      try {
        const response = await fetch(url, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });

        if (response.ok) {
          const geminiRes = await response.json();
          const candidateText = geminiRes.candidates?.[0]?.content?.parts?.[0]?.text;
          if (candidateText) {
            console.log(`[Gemini Vision] Éxito utilizando modelo: ${model}${attempt > 0 ? ` (tras reintento)` : ''}`);
            return { success: true, text: candidateText, modelUsed: model };
          }
        }

        const errText = await response.text();
        lastStatus = response.status;
        console.warn(`[Gemini Vision ${model}] Intento ${attempt + 1} falló con status ${response.status}:`, errText);

        let parsedMsg = '';
        try {
          const parsed = JSON.parse(errText);
          parsedMsg = parsed.error?.message || errText;
        } catch (_) {
          parsedMsg = errText;
        }
        lastErrorDetail = parsedMsg;

        // 1. Si la API Key es inválida o no tiene permisos (400, 401, 403)
        const isKeyError = response.status === 401 ||
          response.status === 403 ||
          lastErrorDetail.toLowerCase().includes('api key not valid') ||
          lastErrorDetail.toLowerCase().includes('api_key_invalid') ||
          lastErrorDetail.toLowerCase().includes('permission_denied');

        if (isKeyError) {
          return {
            success: false,
            status: 401,
            error: 'Tu API Key de Google Gemini es inválida o no tiene permisos. Por favor verifica o renueva tu clave gratuita en el botón de Configuración IA.'
          };
        }

        // 2. Modelo no existe o fue descontinuado (404)
        if (response.status === 404 || lastErrorDetail.toLowerCase().includes('not found') || lastErrorDetail.toLowerCase().includes('no longer available')) {
          console.warn(`[Gemini Vision ${model}] Modelo no disponible (404). Pasando de inmediato al siguiente modelo en cascada...`);
          allQuotaExhausted = false;
          break; // Pasar inmediatamente al siguiente modelo
        }

        // 3. Cuota diaria agotada para este modelo específico (429 con RESOURCE_EXHAUSTED o quota exceeded)
        const isQuotaExhausted = response.status === 429 && (
          lastErrorDetail.toLowerCase().includes('quota exceeded') ||
          lastErrorDetail.toLowerCase().includes('resource_exhausted') ||
          lastErrorDetail.toLowerCase().includes('free_tier_requests') ||
          lastErrorDetail.toLowerCase().includes('check your plan and billing')
        );

        if (isQuotaExhausted) {
          sawQuotaExhausted = true;
          console.warn(`[Gemini Vision ${model}] Cuota diaria agotada para este modelo. Pasando de inmediato al siguiente modelo de respaldo...`);
          break; // NO reintentar este modelo, saltar de inmediato al siguiente modelo
        }

        allQuotaExhausted = false;

        // 4. Si es saturación transitoria (503) o error de servidor (500/502/504) o rate limit por segundo (429)
        const isTransient = [503, 429, 500, 502, 504].includes(response.status) ||
          lastErrorDetail.toLowerCase().includes('high demand') ||
          lastErrorDetail.toLowerCase().includes('spikes in demand');

        if (isTransient) {
          sawHighDemand = true;
          if (attempt < MAX_RETRIES) {
            continue; // Reintento rápido una sola vez
          } else {
            console.warn(`[Gemini Vision ${model}] Servidor ocupado. Pasando al siguiente modelo de respaldo...`);
            break;
          }
        } else {
          // Error de formato u otro tipo no recuperable para este modelo
          break;
        }
      } catch (netErr) {
        console.error(`[Gemini Vision ${model}] Error de red en intento ${attempt + 1}:`, netErr.message);
        lastErrorDetail = netErr.message;
        allQuotaExhausted = false;
        if (attempt < MAX_RETRIES) {
          continue;
        } else {
          break;
        }
      }
    }
  }

  // Si todos los modelos evaluados agotaron su cuota diaria
  if (sawQuotaExhausted && allQuotaExhausted) {
    return {
      success: false,
      status: 429,
      error: 'Se ha agotado el límite diario de solicitudes gratuitas de tu API Key en Google Gemini. Para continuar analizando fotos hoy de forma ilimitada, puedes generar una nueva clave gratuita en Google AI Studio (aistudio.google.com) y pegarla en el botón "Configurar API Key".'
    };
  }

  // Si falló por saturación o alta demanda persistente en los servidores de Google
  if (lastStatus === 503 || sawHighDemand) {
    return {
      success: false,
      status: 503,
      error: 'Los servidores de Google Gemini están experimentando alta demanda momentánea. Por favor espera unos segundos y vuelve a presionar el botón.'
    };
  }

  if (lastStatus === 429) {
    return {
      success: false,
      status: 429,
      error: 'Límite de solicitudes alcanzado en tu cuenta de Gemini. Por favor espera unos momentos y vuelve a intentar.'
    };
  }

  return {
    success: false,
    status: lastStatus,
    error: `Error de Google Gemini: ${lastErrorDetail || 'No se pudo conectar con el servicio'}`
  };
}

// =========================================================
// 12. RECEPCIÓN INTELIGENTE: ESCANEO DE FACTURA/REMISIÓN CON IA
// =========================================================
app.post('/api/escanear-factura-ia', upload.single('imagen'), async (req, res) => {
  try {
    const isDemo = req.body.demo === 'true' || req.body.demo === true;

    // Si el usuario solicita demo o no hay imagen, devolver ejemplo realista
    if (isDemo) {
      const demoResult = {
        proveedor: 'REFACCIONES Y FILTROS DIESEL DEL NORTE S.A. DE C.V.',
        factura: 'F-88429',
        fecha: new Date().toISOString().slice(0, 10),
        moneda: 'MXN',
        total: 18450.00,
        items: [
          { cantidad: 4, numero_parte: 'FS20385', descripcion: 'FILTRO TRAMPA DE DIESEL SEPARADOR DE AGUA', marca: 'FLEETGUARD', costo_unitario: 890.00 },
          { cantidad: 6, numero_parte: '35058', descripcion: 'SELLO DE RUEDA DELANTERA RETEN', marca: 'SKF', costo_unitario: 420.00 },
          { cantidad: 2, numero_parte: '65420', descripcion: 'AMORTIGUADOR DELANTERO SERVICIO PESADO', marca: 'MONROE', costo_unitario: 1250.00 },
          { cantidad: 12, numero_parte: 'INT-9014', descripcion: 'ABRAZADERA SIN FIN 3 PULGADAS ACERO', marca: 'GENERICA', costo_unitario: 45.00 }
        ]
      };

      const itemsEnriquecidos = demoResult.items.map(item => {
        const numParte = String(item.numero_parte || '').trim().toUpperCase();
        const desc = String(item.descripcion || '').trim().toLowerCase();
        let matchProd = null;

        if (numParte) {
          matchProd = Object.values(store.catalog).find(p =>
            p.numero_parte && p.numero_parte.toUpperCase() === numParte
          );
          if (!matchProd) {
            matchProd = Object.values(store.catalog).find(p =>
              Array.isArray(p.equivalencias) &&
              p.equivalencias.some(eq => eq.toUpperCase() === numParte)
            );
          }
        }
        if (!matchProd && desc.length > 4) {
          const words = desc.split(/\s+/).filter(w => w.length > 3);
          if (words.length > 0) {
            matchProd = Object.values(store.catalog).find(p => {
              const pDesc = (p.descripcion || '').toLowerCase();
              return words.every(w => pDesc.includes(w));
            });
          }
        }

        return {
          ...item,
          coincidencia: matchProd ? {
            encontrado: true,
            id: matchProd.id,
            numero_parte_catalogo: matchProd.numero_parte,
            descripcion_catalogo: matchProd.descripcion,
            marca_catalogo: matchProd.marca,
            rack: matchProd.rack,
            nivel: matchProd.nivel,
            stock_actual: matchProd.stock_actual,
            precio_catalogo: matchProd.precio_unitario
          } : {
            encontrado: false,
            sugerencia_codigo: 'INT-' + Math.floor(1000 + Math.random() * 9000),
            sugerencia_rack: 'Rack 1',
            sugerencia_nivel: 'Nivel 1'
          }
        };
      });

      const dupCheckDemo = findExistingFactura(demoResult.factura, demoResult.proveedor);

      return res.json({
        success: true,
        isDemo: true,
        documento: {
          ...demoResult,
          items: itemsEnriquecidos,
          duplicada: Boolean(dupCheckDemo),
          bloqueada: Boolean(dupCheckDemo),
          motivo_duplicado: dupCheckDemo ? `La factura "${demoResult.factura}" ya fue registrada el ${dupCheckDemo.match.fecha_ingreso} (${dupCheckDemo.match.total_piezas} piezas registradas).` : null,
          factura_existente: dupCheckDemo ? dupCheckDemo.match : null
        }
      });
    }

    let base64Data = '';
    let mimeType = 'image/jpeg';

    if (req.file) {
      const filePath = req.file.path;
      const fileBuffer = fs.readFileSync(filePath);
      base64Data = fileBuffer.toString('base64');
      mimeType = req.file.mimetype || 'image/jpeg';
    } else if (req.body.imagenBase64) {
      const parts = req.body.imagenBase64.split(';base64,');
      if (parts.length === 2) {
        mimeType = parts[0].replace('data:', '');
        base64Data = parts[1];
      } else {
        base64Data = req.body.imagenBase64;
      }
    } else {
      return res.status(400).json({ error: 'No se recibió ninguna imagen de factura para escanear' });
    }

    const apiKey = (req.body.apiKey || (store.config && store.config.gemini_api_key) || process.env.GEMINI_API_KEY || '').trim();

    if (!apiKey) {
      return res.status(400).json({
        error: 'API Key de Gemini requerida',
        requiresApiKey: true,
        message: 'Para reconocimiento en vivo de fotos de facturas, ingresa tu API Key gratuita de Google Gemini en el botón de Configuración IA.'
      });
    }

    const promptText = `
Eres un asistente experto en contabilidad de talleres de tractocamiones y refaccionarias diesel mexicanas.
Analiza con máxima precisión esta imagen de factura, remisión, ticket o cotización y extrae los datos en formato JSON estrictamente válido, sin texto adicional:
{
  "proveedor": "Nombre del proveedor o emisor",
  "factura": "Número o folio del documento",
  "fecha": "YYYY-MM-DD",
  "moneda": "MXN",
  "total": 0.00,
  "items": [
    {
      "cantidad": 1,
      "numero_parte": "Código o número de parte de la pieza",
      "descripcion": "Descripción concisa de la refacción",
      "marca": "Marca si es visible",
      "costo_unitario": 0.00
    }
  ]
}
${getMemoriaIAPromptContext()}
`;

    const parts = [
      { text: promptText },
      {
        inline_data: {
          mime_type: mimeType,
          data: base64Data
        }
      }
    ];

    const geminiResult = await callGeminiVision(apiKey, parts);
    if (!geminiResult.success) {
      return res.status(geminiResult.status || 500).json({
        error: geminiResult.error
      });
    }

    const candidateText = geminiResult.text;

    let parsedResult;
    try {
      parsedResult = JSON.parse(candidateText);
    } catch (e) {
      const jsonMatch = candidateText.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        parsedResult = JSON.parse(jsonMatch[0]);
      } else {
        throw new Error('No se pudo decodificar el JSON devuelto por la IA');
      }
    }

    const itemsEnriquecidos = (parsedResult.items || []).map(item => {
      const numParte = String(item.numero_parte || '').trim().toUpperCase();
      const desc = String(item.descripcion || '').trim();
      const marca = String(item.marca || '').trim();

      // En Entrada de Material las facturas corresponden a compras de refacciones e insumos físicos para el almacén (allowServices = false)
      const matchRes = matchItemToCatalog(numParte, desc, marca, parsedResult.proveedor, false);
      const matchProd = matchRes ? matchRes.match : null;
      const isSrv = false;

      return {
        ...item,
        es_servicio: isSrv,
        almacen_id: 'refacciones',
        almacen: 'Almacén de Refacciones',
        coincidencia: matchProd ? {
          encontrado: true,
          id: matchProd.id,
          numero_parte_catalogo: matchProd.numero_parte,
          descripcion_catalogo: matchProd.descripcion,
          marca_catalogo: matchProd.marca,
          rack: matchProd.rack || (isSrv ? 'SERVICIO' : 'Rack 1'),
          nivel: matchProd.nivel || (isSrv ? 'EXTERNO' : 'Nivel 1'),
          stock_actual: matchProd.stock_actual,
          precio_catalogo: matchProd.precio_unitario,
          tipo_coincidencia: matchRes.tipo,
          confianza: matchRes.confianza,
          es_servicio: isSrv
        } : {
          encontrado: false,
          es_servicio: isSrv,
          sugerencias: (matchRes?.sugerencias || []).map(s => ({
            id: s.id,
            numero_parte: s.numero_parte,
            descripcion: s.descripcion,
            marca: s.marca,
            rack: s.rack,
            nivel: s.nivel,
            stock_actual: s.stock_actual
          })),
          sugerencia_codigo: isSrv ? (matchRes?.sugerencia_codigo || 'SRV-' + Math.floor(100 + Math.random() * 900)) : ('INT-' + Math.floor(1000 + Math.random() * 9000)),
          sugerencia_rack: isSrv ? 'SERVICIO' : 'Rack 1',
          sugerencia_nivel: isSrv ? 'EXTERNO' : 'Nivel 1'
        }
      };
    });

    let imageHash = null;
    if (base64Data) {
      imageHash = crypto.createHash('sha256').update(base64Data).digest('hex');
    }

    const dupCheck = findExistingFactura(parsedResult.factura, parsedResult.proveedor, imageHash);

    res.json({
      success: true,
      documento: {
        proveedor: parsedResult.proveedor || '',
        factura: parsedResult.factura || '',
        fecha: parsedResult.fecha || new Date().toISOString().slice(0, 10),
        moneda: parsedResult.moneda || 'MXN',
        total: parsedResult.total || 0,
        items: itemsEnriquecidos,
        imageHash: imageHash,
        duplicada: Boolean(dupCheck),
        bloqueada: Boolean(dupCheck),
        motivo_duplicado: dupCheck ? (
          dupCheck.reason === 'IMAGEN_IDENTICA'
            ? 'Esta misma imagen de factura ya fue subida y registrada previamente en el almacén.'
            : `La factura "${dupCheck.match.factura}" del proveedor "${dupCheck.match.proveedor}" ya fue ingresada el ${dupCheck.match.fecha_ingreso} (${dupCheck.match.total_piezas || 'múltiples'} piezas registradas).`
        ) : null,
        factura_existente: dupCheck ? dupCheck.match : null
      }
    });

  } catch (err) {
    console.error('Error al escanear factura con IA:', err);
    res.status(500).json({ error: err.message });
  }
});

// 12. Analizar múltiples fotos de urgencia (Salidas Rápidas con IA)
app.post('/api/analizar-fotos-urgencia', upload.array('fotos', 10), async (req, res) => {
  try {
    let uploadedPhotos = [];

    // Recoger fotos subidas por multipart
    if (req.files && req.files.length > 0) {
      uploadedPhotos = req.files.map(f => ({
        path: f.path,
        mimetype: f.mimetype || 'image/jpeg',
        filename: f.filename,
        url: `/uploads/${f.filename}`
      }));
    } else if (req.body.fotosBase64 && Array.isArray(req.body.fotosBase64)) {
      for (let i = 0; i < req.body.fotosBase64.length; i++) {
        const raw = req.body.fotosBase64[i];
        if (!raw) continue;
        const parts = raw.split(';base64,');
        const mime = parts.length === 2 ? parts[0].replace('data:', '') : 'image/jpeg';
        const b64 = parts.length === 2 ? parts[1] : raw;
        const filename = `urgencia_${Date.now()}_${i}.jpg`;
        const dest = path.join(UPLOADS_DIR, filename);
        fs.writeFileSync(dest, Buffer.from(b64, 'base64'));
        uploadedPhotos.push({
          path: dest,
          mimetype: mime,
          filename,
          url: `/uploads/${filename}`
        });
      }
    }

    if (uploadedPhotos.length === 0) {
      return res.status(400).json({ error: 'No se recibió ninguna fotografía para analizar.' });
    }

    const apiKey = (req.body.apiKey || (store.config && store.config.gemini_api_key) || process.env.GEMINI_API_KEY || '').trim();
    const isDemo = req.body.isDemo === 'true' || req.body.isDemo === true;
    const isQuickCatalog = req.body.quickCatalog === 'true' || req.body.quickCatalog === true;

    // Si es modo selección rápida por catálogo (sin IA), devolver fotos subidas de inmediato
    if (isQuickCatalog) {
      const quickItems = uploadedPhotos.map((photo, idx) => ({
        id_temporal: `temp_${Date.now()}_${idx}`,
        foto_evidencia: photo.url,
        numero_parte_detectado: '',
        marca_detectada: 'GENÉRICA',
        descripcion_detectada: 'Refacción en foto',
        cantidad: 1,
        confianza_ia: 'MANUAL',
        coincidencia: {
          encontrado: false,
          tipo_match: 'MANUAL',
          sugerencias: []
        }
      }));

      return res.json({
        success: true,
        quickCatalog: true,
        totalFotos: uploadedPhotos.length,
        items: quickItems,
        message: 'Fotografías cargadas con éxito. Selecciona la refacción y confirma el despacho.'
      });
    }

    if (!apiKey && !isDemo) {
      return res.status(400).json({
        success: false,
        error: 'API Key de Gemini requerida',
        requiresApiKey: true,
        message: 'Para reconocimiento en vivo de fotos de refacciones, ingresa tu API Key gratuita de Google Gemini en el botón de Configuración IA.'
      });
    }

    // Si es modo demo explícito, simular reconocimiento inteligente usando piezas reales del catálogo
    if (isDemo) {
      const catalogSample = Object.values(store.catalog).slice(0, Math.min(uploadedPhotos.length * 2, 6));
      const demoItems = uploadedPhotos.map((photo, idx) => {
        const samplePart = catalogSample[idx % catalogSample.length] || {
          numero_parte: 'FS20385',
          descripcion: 'Filtro separador de agua combustible',
          marca: 'FLEETGUARD',
          stock_actual: 5,
          rack: 'Rack 1',
          nivel: 'Nivel 1'
        };

        return {
          id_temporal: `temp_${Date.now()}_${idx}`,
          foto_evidencia: photo.url,
          numero_parte_detectado: samplePart.numero_parte,
          marca_detectada: samplePart.marca,
          descripcion_detectada: samplePart.descripcion,
          cantidad: 1,
          confianza_ia: 'ALTA',
          coincidencia: {
            encontrado: true,
            tipo_match: 'DEMO_IA',
            id: samplePart.id,
            numero_parte: samplePart.numero_parte,
            descripcion: samplePart.descripcion,
            marca: samplePart.marca,
            rack: samplePart.rack,
            nivel: samplePart.nivel,
            stock_actual: samplePart.stock_actual,
            precio_unitario: samplePart.precio_unitario || 0,
            foto_referencia: samplePart.foto_referencia || photo.url
          }
        };
      });

      return res.json({
        success: true,
        isDemo: true,
        totalFotos: uploadedPhotos.length,
        items: demoItems,
        message: 'Reconocimiento de demostración generado. Para análisis de visión en vivo con tus fotos, ingresa tu API Key gratuita de Google Gemini.'
      });
    }

    const origen = String(req.body.origen || req.body.modo || '').trim().toLowerCase();
    const isEntrada = origen === 'entrada';

    // Modo en Vivo con Gemini Vision (cascada de alta disponibilidad)
    const promptText = isEntrada ? `
Eres un perito experto en inventario de refacciones diésel y suministros para talleres mecánicos de tractocamiones (Cummins, Detroit, Caterpillar, International, Kenworth, Freightliner, Volvo).
Se te proporcionan ${uploadedPhotos.length} fotos de una FACTURA DE COMPRA, REMISIÓN DE REFACCIONES o FOTOS DE PIEZAS que están INGRESANDO AL ALMACÉN.

REGLAS ESTRICTAS DE EXTRACCIÓN (ENTRADA AL ALMACÉN DE REFACCIONES):
1. TODAS las piezas o líneas detectadas son REFACCIONES O MATERIALES FÍSICOS (filtros de aceite, filtros de aire, combustible, lubricantes, retenes, baleros, mangueras, bandas, balatas, tambores, abrazaderas, tornillería, consumibles).
2. NUNCA las clasifiques como servicio ("es_servicio": false para todas). En las entradas al almacén no existen servicios; todo es refacción física para anaquel o paso directo.
3. Si la foto es una factura, nota de venta o remisión con múltiples conceptos (ej. "COMPRA DE FILTROS DE ACEITE DONALDSON", "FILTRO DE AIRE TIPO PANEL DONALDSON P609"):
   - Extrae cada refacción facturada por separado.
   - "numero_parte": El número de parte o código de catálogo del fabricante visible (ej. "P609", "LF9009", "FS20385", "35058", "165"). Si no tiene código, pon "S/N".
   - "descripcion": El nombre limpio y en mayúsculas de la refacción (ej. "FILTRO DE AIRE TIPO PANEL DONALDSON P609", "FILTRO DE ACEITE DONALDSON"). Elimina prefijos como "COMPRA DE", "VENTA DE", "SUMINISTRO DE".
   - "marca": La marca fabricante (DONALDSON, FLEETGUARD, BALDWIN, CUMMINS, BENDIX, 3M, SKF, etc.). Si no viene explícita, pon "GENERICA".
   - "cantidad": La cantidad facturada o unidades físicas.
   - "precio_unitario": El costo o precio unitario antes de impuestos indicado en la factura. Si no es legible, pon 0.
   - "confianza": "ALTA" o "MEDIA".
4. Si en el encabezado de la factura identificas el proveedor o emisor (ej. "COMERCIAL ESSEX", "DISTRIBUIDORA DONALDSON"), extráelo en "proveedor_factura".
5. Si identificas el folio o número de factura / remisión (ej. "1320", "FA-9842"), extráelo en "folio_factura".

Devuelve OBLIGATORIAMENTE un JSON estrictamente estructurado así, sin texto adicional:
{
  "proveedor_factura": "COMERCIAL ESSEX",
  "folio_factura": "1320",
  "items": [
    {
      "foto_index": 0,
      "es_servicio": false,
      "numero_parte": "P609",
      "marca": "DONALDSON",
      "descripcion": "FILTRO DE AIRE TIPO PANEL DONALDSON P609",
      "cantidad": 1,
      "precio_unitario": 450.00,
      "confianza": "ALTA"
    }
  ]
}
${getMemoriaIAPromptContext()}
` : `
Eres un perito experto en inventario de refacciones y servicios mecánicos para tractocamiones diésel y suministros de taller (Cummins, Detroit, Caterpillar, International, Kenworth, Freightliner, Volvo).
Se te proporcionan ${uploadedPhotos.length} fotos tomadas en mostrador o patio de taller (para despacho o para recepción al almacén).

Para CADA una de las fotos proporcionadas (identificadas como foto 0, foto 1, foto 2...):
1. Inspecciona meticulosamente la imagen: puede ser:
   A) Una pieza física o insumo de taller (motor, suspensión, frenos, consumible).
   B) O bien un COMPROBANTE, FACTURA, REMISIÓN o NOTA DE TALLER que ampara un SERVICIO EXTERNO (ej. Rimar mangos, Rectificar campanas o discos, Torno, Mano de obra externa, Soldadura de chasis, Alineación y balanceo, Reparación de compresor, Lavado de inyectores, etc.).
2. Si es una FACTURA, REMISIÓN o SERVICIO EXTERNO:
   - "es_servicio": true
   - "descripcion": El concepto claro del trabajo o reparación realizado (ej. "RIMAR PAR DE MANGOS DE DIRECCION").
   - "numero_parte": "SERVICIO" o el código si viene indicado en la factura.
   - "marca": El nombre del taller o proveedor externo (ej. "TORNO GARCIA", "RECTIFICADORA MONTERREY", etc.).
   - "proveedor": El nombre del taller o proveedor que realizó el trabajo.
   - "factura": El número de factura o remisión si es legible (ej. "1320", "4876").
   - "precio_unitario": El costo o monto total del servicio si viene legible en el comprobante.
   - "cantidad": 1.
3. Si es una PIEZA FÍSICA:
   - "es_servicio": false
   - "descripcion": El nombre de la pieza en español (ej. "FILTRO DE ACEITE", "CINTA DE AISLAR", "RETEN", "EMPAQUE", etc.).
   - "numero_parte": El número de parte visible (ej. FS20385, LF9009, 35058, 165). Si es a granel o no es visible, escribe "S/N".
   - "marca": La marca identificada (3M, Truper, Fleetguard, Donaldson, Bendix, SKF, Monroe, etc.). Si no es visible, escribe "GENERICA".
   - "cantidad": Cantidad visible en la foto.

Devuelve OBLIGATORIAMENTE un JSON estrictamente estructurado así, sin texto adicional:
{
  "items": [
    {
      "foto_index": 0,
      "es_servicio": false,
      "numero_parte": "165",
      "marca": "3M",
      "proveedor": "",
      "factura": "",
      "descripcion": "CINTA DE AISLAR",
      "cantidad": 1,
      "precio_unitario": 0,
      "confianza": "ALTA"
    }
  ]
}
${getMemoriaIAPromptContext()}
`;

    const contentsParts = [{ text: promptText }];
    for (const p of uploadedPhotos) {
      const fileBuffer = fs.readFileSync(p.path);
      contentsParts.push({
        inline_data: {
          mime_type: p.mimetype,
          data: fileBuffer.toString('base64')
        }
      });
    }

    const geminiResult = await callGeminiVision(apiKey, contentsParts);
    if (!geminiResult.success) {
      return res.status(geminiResult.status || 500).json({
        error: geminiResult.error
      });
    }

    const candidateText = geminiResult.text;

    let parsedResult;
    try {
      parsedResult = JSON.parse(candidateText);
    } catch (e) {
      const jsonMatch = candidateText.match(/\{[\s\S]*\}/);
      if (jsonMatch) parsedResult = JSON.parse(jsonMatch[0]);
      else throw new Error('No se pudo decodificar la respuesta JSON de la IA');
    }

    const enrichedItems = (parsedResult.items || []).map((item, idx) => {
      const isSrv = isEntrada ? false : Boolean(item.es_servicio || isServiceDescription(item.descripcion));
      const matchResult = matchItemToCatalog(item.numero_parte, item.descripcion, item.marca, item.proveedor || parsedResult.proveedor_factura, !isEntrada);
      const fotoObj = uploadedPhotos[item.foto_index] || uploadedPhotos[0] || {};
      const detectedPrice = Number(item.precio_unitario || item.costo_unitario || 0);

      return {
        id_temporal: `temp_${Date.now()}_${idx}`,
        foto_evidencia: fotoObj.url || '',
        es_servicio: isSrv,
        tipo: isSrv ? 'SERVICIO' : 'REFACCION',
        proveedor_detectado: item.proveedor || parsedResult.proveedor_factura || item.marca || '',
        factura_detectada: item.factura || parsedResult.folio_factura || '',
        numero_parte_detectado: isSrv ? 'SERVICIO' : (item.numero_parte || ''),
        marca_detectada: item.marca || '',
        descripcion_detectada: item.descripcion || (isSrv ? 'Servicio Externo' : ''),
        cantidad: Number(item.cantidad) || 1,
        precio_unitario_detectado: detectedPrice,
        costo_unitario: detectedPrice,
        confianza_ia: item.confianza || 'MEDIA',
        coincidencia: isSrv ? {
          encontrado: false,
          tipo_match: 'SERVICIO_EXTERNO',
          id: null,
          numero_parte: 'SERVICIO',
          descripcion: item.descripcion,
          marca: item.proveedor || item.marca || 'PROVEEDOR EXTERNO',
          rack: 'Servicio Externo',
          nivel: 'Proveedor',
          stock_actual: 0,
          precio_unitario: detectedPrice,
          foto_referencia: fotoObj.url || ''
        } : (matchResult.match ? {
          encontrado: true,
          tipo_match: matchResult.tipo,
          id: matchResult.match.id,
          numero_parte: matchResult.match.numero_parte,
          descripcion: matchResult.match.descripcion,
          marca: matchResult.match.marca,
          rack: matchResult.match.rack,
          nivel: matchResult.match.nivel,
          stock_actual: matchResult.match.stock_actual,
          precio_unitario: matchResult.match.precio_unitario || 0,
          foto_referencia: matchResult.match.foto_referencia || ''
        } : {
          encontrado: false,
          tipo_match: 'NINGUNO',
          sugerencias: (matchResult.sugerencias || []).map(s => ({
            id: s.id,
            numero_parte: s.numero_parte,
            descripcion: s.descripcion,
            marca: s.marca,
            rack: s.rack || 'Rack 1',
            nivel: s.nivel || 'Nivel 1',
            stock_actual: s.stock_actual,
            precio_unitario: s.precio_unitario || 0
          }))
        })
      };
    });

    res.json({
      success: true,
      totalFotos: uploadedPhotos.length,
      proveedor_detectado: parsedResult.proveedor_factura || '',
      factura_detectada: parsedResult.folio_factura || '',
      items: enrichedItems
    });

  } catch (err) {
    console.error('Error procesando fotos de urgencia con IA:', err);
    res.status(500).json({ error: err.message });
  }
});

// 12.1 Analizar Factura / Comprobante de Servicio Externo con IA (Torno, Rectificado, Soldadura)
app.post('/api/analizar-servicio-externo', upload.single('foto'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, error: 'No se recibió ningún comprobante o archivo para analizar.' });
    }

    const photoUrl = `/uploads/${req.file.filename}`;
    const filePath = req.file.path;
    const mimeType = req.file.mimetype || 'image/jpeg';
    const apiKey = (req.body.apiKey || (store.config && store.config.gemini_api_key) || process.env.GEMINI_API_KEY || '').trim();

    if (!apiKey) {
      return res.json({
        success: true,
        foto_url: photoUrl,
        datos: null,
        message: 'Comprobante guardado como evidencia. Ingresa una API Key de Gemini para extracción automática de datos con IA.'
      });
    }

    const promptText = `
Eres un auditor contable y perito mecánico experto en talleres de tractocamiones diésel y flotillas de transporte pesado.
Analiza detenidamente este documento (factura, remisión, ticket o comprobante de un SERVICIO O REPARACIÓN EXTERNA realizada por un taller, torno o proveedor externo).
Ejemplos de trabajos externos: Rimado de mangos, Rectificado de campanas/discos/tambores, Torno de flechas/yugos, Soldadura y enderezado de chasis, Reparación de marchas/alternadores/compresores, Lavado y calibración de inyectores, Alineación y balanceo, etc.

Extrae con total precisión los datos en este JSON estricto, sin texto adicional:
{
  "descripcion": "Descripción concisa y profesional del trabajo o servicio realizado (ej. RECTIFICADO DE CAMPANAS DE FRENO, RIMAR PAR DE MANGOS)",
  "proveedor": "Nombre comercial o razón social del taller, torno o proveedor externo",
  "factura": "Folio, número de factura, remisión o nota si es visible (ej. 1320, FA-489)",
  "costo_total": 0,
  "unidad_camion": "",
  "folio_orden": "",
  "solicitante": ""
}
`;

    const fileBuffer = fs.readFileSync(filePath);
    const contentsParts = [
      { text: promptText },
      {
        inline_data: {
          mime_type: mimeType,
          data: fileBuffer.toString('base64')
        }
      }
    ];

    const geminiResult = await callGeminiVision(apiKey, contentsParts);
    if (!geminiResult.success) {
      return res.json({
        success: true,
        foto_url: photoUrl,
        datos: null,
        warning: geminiResult.error,
        message: 'Comprobante adjuntado como evidencia. Llena los datos manualmente.'
      });
    }

    let parsed = null;
    try {
      parsed = JSON.parse(geminiResult.text);
    } catch (e) {
      const match = geminiResult.text.match(/\{[\s\S]*\}/);
      if (match) parsed = JSON.parse(match[0]);
    }

    return res.json({
      success: true,
      foto_url: photoUrl,
      datos: parsed || null,
      message: parsed ? 'Comprobante analizado con éxito por IA' : 'Comprobante adjuntado como evidencia'
    });
  } catch (err) {
    console.error('Error en /api/analizar-servicio-externo:', err);
    return res.status(500).json({ success: false, error: err.message });
  }
});

// 13. Confirmar y procesar salidas masivas detectadas por fotos
app.post('/api/confirmar-salidas-fotos', (req, res) => {
  try {
    if (req.body.usuarioRol === 'GERENTE' || (req.body.permisos && req.body.permisos.dar_salida === false)) {
      return res.status(403).json({ error: 'Acceso denegado: Tu perfil es de Solo Consulta / Auditoría.' });
    }

    const { unidad_camion, solicitante, motivo, folio, fecha, items } = req.body || {};
    if (!Array.isArray(items) || items.length === 0) {
      return res.status(400).json({ error: 'No se recibieron refacciones para procesar la salida' });
    }

    const now = new Date();
    const fechaSalida = fecha ? String(fecha).trim() : now.toISOString().slice(0, 10);
    const horaSalida = now.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    const folioGen = folio || `SAL-FOTO-${Date.now().toString(36).toUpperCase()}`;

    let procesados = 0;
    const movimientosCreados = [];

    for (const it of items) {
      const isSrv = Boolean(it.es_servicio || it.tipo === 'SERVICIO' || isServiceDescription(it.descripcion));
      const cant = Number(it.cantidad) || 1;
      let itemCat = null;

      if (!isSrv) {
        itemCat = store.catalog[it.id_pieza] || findPart(it.id_pieza) || findPart(it.numero_parte);
        if (itemCat) {
          itemCat.stock_actual = (Number(itemCat.stock_actual) || 0) - cant;
          if (it.guardar_como_foto_catalogo && it.foto && !itemCat.foto_referencia) {
            itemCat.foto_referencia = it.foto;
          }
        } else if (it.descripcion && it.descripcion !== 'Refacción') {
          // Si no existía en el catálogo pero se reconoció en foto, registrar en catálogo
          let finalParte = String(it.numero_parte || '').trim();
          const esGranel = !finalParte || finalParte === 'S/N' || finalParte === 'N/A';
          if (esGranel) {
            const nextNum = Object.keys(store.catalog).filter(k => k.startsWith('INT_') || k.startsWith('SN_')).length + 1;
            finalParte = `INT-${String(nextNum).padStart(4, '0')}`;
          }
          const safeId = finalParte.replace(/[^a-zA-Z0-9_-]/g, '_');
          itemCat = {
            id: safeId,
            rack: it.rack || 'Rack 1',
            nivel: it.nivel || 'Nivel 1',
            descripcion: it.descripcion,
            marca: it.marca || 'GENÉRICA',
            numero_parte: finalParte,
            equivalencias: [],
            alias: [],
            es_a_granel: esGranel,
            foto_referencia: it.foto || '',
            unidad: it.unidad || 'PZ',
            precio_unitario: Number(it.precio_unitario || 0),
            stock_inicial: cant,
            stock_actual: 0,
            almacen: 'Almacén de Refacciones',
            almacen_id: 'refacciones',
            creado_en_salida: true,
            fecha_creacion: fechaSalida
          };
          store.catalog[safeId] = itemCat;
        }
      }

      const itUnidad = String(it.unidad_camion || unidad_camion || 'GENERAL').trim().toUpperCase();
      const itSolicitante = String(it.solicitante || solicitante || 'JEFE DE TALLER').trim().toUpperCase();
      const itFolio = it.folio ? String(it.folio).trim().toUpperCase() : (folio ? String(folio).trim().toUpperCase() : `SAL-FOTO-${Date.now().toString(36).toUpperCase()}`);
      const itMotivo = it.motivo ? String(it.motivo).trim() : (motivo ? String(motivo).trim() : 'Despacho ágil con evidencia fotográfica');
      const costUnit = Number(it.costo_unitario || it.precio_unitario || (itemCat ? itemCat.precio_unitario : 0)) || 0;

      const mov = {
        id: `mov_${Date.now()}_${procesados}_${Math.floor(1000 + Math.random() * 9000)}`,
        tipo: 'S',
        tipo_movimiento: isSrv ? 'SERVICIO' : 'SALIDA',
        es_servicio: isSrv,
        fecha: fechaSalida,
        hora: horaSalida,
        id_pieza: isSrv ? 'SERVICIO_EXTERNO' : (itemCat ? itemCat.id : (it.id_pieza || `P_URG_${procesados}`)),
        numero_parte: isSrv ? 'SERVICIO' : (it.numero_parte || (itemCat ? itemCat.numero_parte : 'S/N')),
        descripcion: it.descripcion || (itemCat ? itemCat.descripcion : (isSrv ? 'Servicio Externo' : 'Refacción')),
        marca: isSrv ? (it.proveedor || it.marca || 'PROVEEDOR EXTERNO') : (it.marca || (itemCat ? itemCat.marca : 'GENÉRICA')),
        proveedor: isSrv ? (it.proveedor || '') : '',
        factura: isSrv ? (it.factura || '') : '',
        almacen: isSrv ? 'Servicios y Reparaciones Externas' : (itemCat ? itemCat.almacen : 'Almacén de Refacciones'),
        almacen_id: isSrv ? 'servicios_externos' : (itemCat ? itemCat.almacen_id : 'refacciones'),
        ubicacion: isSrv ? 'Servicio Externo / Proveedor' : (itemCat ? `${itemCat.rack} / ${itemCat.nivel}` : 'Taller / Despacho Ágil'),
        cantidad: cant,
        unidad_camion: itUnidad,
        folio: itFolio,
        solicitante: itSolicitante,
        foto: it.foto || '',
        costo_unitario: costUnit,
        costo_total: costUnit * cant,
        notas: itMotivo ? (isSrv ? `[SERVICIO EXTERNO] ${itMotivo}` : `Despacho por foto: ${itMotivo}`) : (isSrv ? '[SERVICIO EXTERNO REGISTRADO DIRECTO A UNIDAD]' : 'Despacho ágil con evidencia fotográfica'),
        timestamp: now.toISOString()
      };

      store.movimientos.unshift(mov);
      movimientosCreados.push(mov);
      procesados++;
    }

    saveStore();
    try { movimientosCreados.forEach(m => syncSalidaConTaller(m)); } catch(e) {}
    res.json({
      success: true,
      message: `¡Se procesaron ${procesados} salidas de refacciones con éxito!`,
      count: procesados,
      folio: folioGen,
      movimientos: movimientosCreados
    });
  } catch (err) {
    console.error('Error al confirmar salidas por fotos:', err);
    res.status(500).json({ error: err.message });
  }
});

// 14. Confirmar y procesar entradas de refacciones detectadas por fotos físicas
app.post('/api/confirmar-entradas-fotos', (req, res) => {
  try {
    if (req.body.usuarioRol === 'GERENTE' || (req.body.permisos && req.body.permisos.dar_entrada === false)) {
      return res.status(403).json({ error: 'Acceso denegado: Tu perfil es de Solo Consulta / Auditoría.' });
    }

    const { proveedor, factura, remitente, motivo, fecha, items } = req.body || {};
    if (!Array.isArray(items) || items.length === 0) {
      return res.status(400).json({ error: 'No se recibieron refacciones para procesar la entrada' });
    }

    const now = new Date();
    const fechaEntrada = fecha ? String(fecha).trim() : now.toISOString().slice(0, 10);
    const horaEntrada = now.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    const folioGen = factura ? String(factura).trim().toUpperCase() : `ENT-FOTO-${Date.now().toString(36).toUpperCase()}`;

    const condicionPago = String(req.body.condicion_pago || 'CONTADO').trim().toUpperCase() === 'CREDITO' ? 'CREDITO' : 'CONTADO';
  const diasCredito = condicionPago === 'CREDITO' ? (Math.max(1, Number(req.body.dias_credito) || 30)) : 0;
  const fechaFactura = req.body.fecha_factura ? String(req.body.fecha_factura).trim() : (req.body.fecha_documento || fechaEntrada || fecha);
  const fechaVencimiento = condicionPago === 'CREDITO'
    ? (req.body.fecha_vencimiento ? String(req.body.fecha_vencimiento).trim() : calcularVencimiento(fechaFactura, diasCredito))
    : fechaFactura;
  const estadoPago = condicionPago === 'CREDITO' ? (req.body.estado_pago || 'PENDIENTE') : 'PAGADO';
  const fechaPago = req.body.fecha_pago || (estadoPago === 'PAGADO' && condicionPago === 'CONTADO' ? fecha : '');
  const referenciaPago = req.body.referencia_pago || '';

  let procesados = 0;
    const movimientosCreados = [];

    for (const it of items) {
      const cant = Number(it.cantidad) || 1;
      const costo = Number(it.costo_unitario) || 0;
      let itemCat = null;
      if (!it.crear_nueva_pieza && it.id_pieza) {
        itemCat = store.catalog[it.id_pieza] || findPart(it.id_pieza);
      } else if (!it.crear_nueva_pieza && it.numero_parte) {
        itemCat = findPart(it.numero_parte);
      }

      // Si se marcó para crear como nueva refacción en catálogo o no se encontró en catálogo
      if (it.crear_nueva_pieza || !itemCat) {
        const newId = `part_${Date.now()}_${procesados}_${Math.floor(1000 + Math.random() * 9000)}`;
        itemCat = {
          id: newId,
          rack: (it.rack || 'Rack 1').trim(),
          nivel: (it.nivel || 'Nivel 1').trim(),
          descripcion: (it.descripcion || 'Refacción Nueva').toUpperCase().trim(),
          marca: (it.marca || 'GENÉRICA').toUpperCase().trim(),
          numero_parte: (it.numero_parte || `NUEVA-${Date.now().toString(36).toUpperCase()}`).toUpperCase().trim(),
          unidad: 'PZA',
          precio_unitario: costo,
          stock_inicial: 0,
          stock_actual: 0,
          foto_referencia: it.foto || ''
        };
        store.catalog[newId] = itemCat;
      } else {
        // Si es pieza existente, permitir reacomodar rack y nivel si el usuario los especificó
        if (it.rack && it.rack.trim()) {
          itemCat.rack = it.rack.trim();
        }
        if (it.nivel && it.nivel.trim()) {
          itemCat.nivel = it.nivel.trim();
        }
      }

      const esPasoDirecto = Boolean(it.paso_directo || it.es_paso_directo);
      const unidadDirecta = String(it.unidad_camion || '').trim().toUpperCase();

      if (itemCat) {
        if (!esPasoDirecto) {
          itemCat.stock_actual = (Number(itemCat.stock_actual) || 0) + cant;
        }
        if (costo > 0) itemCat.precio_unitario = costo;
        if (it.guardar_como_foto_catalogo && it.foto) {
          itemCat.foto_referencia = it.foto;
        }
      }

      const mov = {
        id: `mov_${Date.now()}_${procesados}`,
        tipo: esPasoDirecto ? 'E/S' : 'E',
        fecha: fechaEntrada,
        hora: horaEntrada,
        id_pieza: itemCat ? itemCat.id : (it.id_pieza || `P_ENT_${procesados}`),
        numero_parte: it.numero_parte || (itemCat ? itemCat.numero_parte : 'S/N'),
        descripcion: it.descripcion || (itemCat ? itemCat.descripcion : 'Refacción'),
        marca: it.marca || (itemCat ? itemCat.marca : 'GENÉRICA'),
        ubicacion: esPasoDirecto ? `PASO DIRECTO ➔ UNIDAD ${unidadDirecta || 'TALLER'}` : (itemCat ? `${itemCat.rack} / ${itemCat.nivel}` : 'Almacén General'),
        cantidad: cant,
        proveedor: String(proveedor || 'MOSTRADOR / PROVEEDOR').trim().toUpperCase(),
        factura: folioGen,
        solicitante: esPasoDirecto ? String(it.solicitante || it.mecanico || remitente || 'MECÁNICO').trim().toUpperCase() : String(remitente || 'ALMACÉN').trim().toUpperCase(),
        entregado_por: String(remitente || 'ALMACÉN').trim().toUpperCase(),
        unidad_camion: esPasoDirecto ? unidadDirecta : '',
        folio: esPasoDirecto ? String(it.folio || folioGen).trim() : '',
        motivo: esPasoDirecto ? String(it.motivo || motivo || 'Paso Directo a Unidad').trim() : '',
        costo_unitario: costo,
        foto: it.foto || '',
        notas: esPasoDirecto ? `Paso Directo E/S a Unidad ${unidadDirecta}: ${it.motivo || motivo || 'Sin notas'}` : (motivo ? `Entrada por foto: ${motivo}` : 'Entrada ágil con evidencia fotográfica'),
        timestamp: now.toISOString()
      };

      store.movimientos.unshift(mov);
      movimientosCreados.push(mov);
      procesados++;
    }

    saveStore();
    res.json({
      success: true,
      message: `¡Se registraron ${procesados} entradas de refacciones al almacén con éxito!`,
      count: procesados,
      folio: folioGen,
      movimientos: movimientosCreados
    });
  } catch (err) {
    console.error('Error al confirmar entradas por fotos:', err);
    res.status(500).json({ error: err.message });
  }
});


// =========================================================
// 12. SISTEMA INTELIGENTE DE ÓRDENES DE COMPRA Y RADAR DE REABASTECIMIENTO
// =========================================================

// Algoritmo analítico para identificar piezas de mayor rotación y riesgos de desabasto
function calculateSugerenciasCompra() {
  const consumoPorPieza = {};
  const proveedoresPorPieza = {};

  // Analizar histórico de salidas y pasos directos
  for (const m of (store.movimientos || [])) {
    if (m.anulado) continue;
    const parteId = m.id_pieza || m.numero_parte;
    if (!parteId) continue;

    if (m.tipo === 'S' || m.tipo === 'E/S') {
      const cant = Number(m.cantidad) || 0;
      if (!consumoPorPieza[parteId]) {
        consumoPorPieza[parteId] = {
          total_despachado: 0,
          veces_pedido: 0,
          unidades: new Set(),
          ultima_fecha: m.fecha
        };
      }
      consumoPorPieza[parteId].total_despachado += cant;
      consumoPorPieza[parteId].veces_pedido += 1;
      if (m.unidad_camion && m.unidad_camion !== 'S/N') {
        consumoPorPieza[parteId].unidades.add(m.unidad_camion);
      }
      if (!consumoPorPieza[parteId].ultima_fecha || m.fecha > consumoPorPieza[parteId].ultima_fecha) {
        consumoPorPieza[parteId].ultima_fecha = m.fecha;
      }
    }

    // Registrar proveedores frecuentes observados en compras pasadas
    if (m.proveedor && m.proveedor !== 'S/N' && m.proveedor.trim()) {
      if (!proveedoresPorPieza[parteId]) proveedoresPorPieza[parteId] = {};
      const provNorm = m.proveedor.trim().toUpperCase();
      proveedoresPorPieza[parteId][provNorm] = (proveedoresPorPieza[parteId][provNorm] || 0) + 1;
    }
  }

  const sugerencias = [];
  let piezasAgotadasUrgentes = 0;
  let piezasStockBajo = 0;
  let totalPiezasSugeridas = 0;
  let inversionEstimadaTotal = 0;

  for (const [id, item] of Object.entries(store.catalog || {})) {
    const parteKey = item.id || item.numero_parte;
    const consumo = consumoPorPieza[item.id] || consumoPorPieza[item.numero_parte] || {
      total_despachado: 0,
      veces_pedido: 0,
      unidades: new Set(),
      ultima_fecha: null
    };

    const stockActual = Number(item.stock_actual !== undefined ? item.stock_actual : 0);
    const precio = Number(item.precio_unitario || 0);

    // Identificar proveedor sugerido
    let proveedorSugerido = 'PROVEEDOR RECOMENDADO';
    const provMap = proveedoresPorPieza[item.id] || proveedoresPorPieza[item.numero_parte];
    if (provMap) {
      const topProv = Object.entries(provMap).sort((a, b) => b[1] - a[1])[0];
      if (topProv) proveedorSugerido = topProv[0];
    }

    // Reglas de Alerta Operativa
    let prioridad = null;
    let urgenciaScore = 0;
    let cantidadSugerida = 0;
    let motivo = '';

    if (stockActual <= 0 && consumo.total_despachado > 0) {
      prioridad = 'URGENTE';
      motivo = 'Agotado en Almacén con Demanda Activa';
      urgenciaScore = 100 + (consumo.total_despachado * 3) + (consumo.veces_pedido * 2);
      cantidadSugerida = Math.max(5, Math.ceil(consumo.total_despachado * 1.5));
      piezasAgotadasUrgentes++;
    } else if (stockActual <= 2 && consumo.total_despachado > 0) {
      prioridad = 'ALERTA';
      motivo = 'Stock Crítico (Riesgo Inminente de Paro)';
      urgenciaScore = 60 + ((3 - stockActual) * 12) + (consumo.total_despachado * 2);
      cantidadSugerida = Math.max(4, Math.ceil(consumo.total_despachado - stockActual + 3));
      piezasStockBajo++;
    } else if (consumo.total_despachado >= 5 && stockActual < consumo.total_despachado) {
      prioridad = 'PREVENTIVO';
      motivo = 'Alta Rotación (Reabastecimiento Recomendado)';
      urgenciaScore = 30 + consumo.total_despachado;
      cantidadSugerida = Math.max(3, Math.ceil(consumo.total_despachado - stockActual + 2));
    } else if (stockActual <= 0) {
      prioridad = 'BAJA';
      motivo = 'Sin Existencias (Baja Rotación)';
      urgenciaScore = 10;
      cantidadSugerida = 2;
    }

    if (prioridad) {
      const costoEstimado = cantidadSugerida * precio;
      totalPiezasSugeridas += cantidadSugerida;
      inversionEstimadaTotal += costoEstimado;

      sugerencias.push({
        id_pieza: item.id,
        numero_parte: item.numero_parte,
        descripcion: item.descripcion,
        marca: item.marca || 'GENERICA',
        almacen: item.almacen || 'Almacén de Refacciones',
        almacen_id: item.almacen_id || 'refacciones',
        rack: item.rack || 'S/R',
        nivel: item.nivel || 'S/N',
        stock_actual: stockActual,
        unidad: item.unidad || 'PZ',
        precio_unitario: precio,
        total_despachado: consumo.total_despachado,
        veces_pedido: consumo.veces_pedido,
        unidades_frecuentes: Array.from(consumo.unidades).slice(0, 4),
        ultima_salida: consumo.ultima_fecha,
        prioridad,
        motivo,
        urgencia_score: urgenciaScore,
        cantidad_sugerida: cantidadSugerida,
        costo_estimado: costoEstimado,
        proveedor_sugerido: proveedorSugerido
      });
    }
  }

  sugerencias.sort((a, b) => b.urgencia_score - a.urgencia_score);

  return {
    total_sugerencias: sugerencias.length,
    resumen: {
      piezas_agotadas_urgentes: piezasAgotadasUrgentes,
      piezas_stock_bajo: piezasStockBajo,
      total_piezas_sugeridas: totalPiezasSugeridas,
      inversion_estimada: inversionEstimadaTotal
    },
    sugerencias
  };
}

// 12.1 Obtener sugerencias predictivas de compra
app.get('/api/compras/sugerencias', (req, res) => {
  try {
    const data = calculateSugerenciasCompra();
    res.json({ success: true, ...data });
  } catch (err) {
    console.error('Error calculando sugerencias de compra:', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// Helper de sincronización bidireccional con Taller Pro (Control CESS)
const TALLER_DB_FILE = path.join(__dirname, '..', 'taller_db.json');

function syncOrdenConTaller(orden, nuevoEstatus, autor, factura) {
  if (!fs.existsSync(TALLER_DB_FILE) || !orden) return;
  try {
    const unidades = JSON.parse(fs.readFileSync(TALLER_DB_FILE, 'utf8'));
    let cambios = false;

    unidades.forEach(u => {
      if (Array.isArray(u.piezas)) {
        u.piezas.forEach(pz => {
          if (pz && pz.folio_oc && (pz.folio_oc === orden.folio || pz.folio_oc === orden.id)) {
            cambios = true;
            if (nuevoEstatus === 'RECIBIDA') {
              pz.estado = 'RECIBIDA_TALLER';
              pz.fecha_recepcion = new Date().toISOString();
              pz.recibido_por = autor || 'Almacén Central';
              if (factura) pz.factura_remision = factura;

              if (!Array.isArray(u.historial)) u.historial = [];
              u.historial.unshift({
                fecha: new Date().toLocaleString('es-MX'),
                evento: `Refacción RECIBIDA EN ALMACÉN / TALLER: ${pz.cantidad || 1}x ${pz.descripcion} (OC: ${orden.folio}). Surtida y lista para montaje en Unidad ${u.id}.`,
                tipo: 'recepcion_almacen',
                autor: autor || 'Almacén Central'
              });
            } else if (nuevoEstatus === 'CANCELADA') {
              pz.estado = 'SOLICITADA_ALMACEN';
              if (!Array.isArray(u.historial)) u.historial = [];
              u.historial.unshift({
                fecha: new Date().toLocaleString('es-MX'),
                evento: `[Cancelación OC]: Orden de Compra ${orden.folio} fue cancelada en Almacén para ${pz.descripcion}.`,
                tipo: 'orden_compra_cancelada',
                autor: autor || 'Almacén Central'
              });
            } else if (nuevoEstatus === 'ORDENADA') {
              pz.estado = 'EN_ORDEN_COMPRA';
              if (orden.proveedor) pz.proveedor_oc = orden.proveedor;
              if (orden.fecha_entrega_esperada) pz.fecha_estimada_arribo = orden.fecha_entrega_esperada;
            }
          }
        });
      }
    });

    if (cambios) {
      fs.writeFileSync(TALLER_DB_FILE, JSON.stringify(unidades, null, 2), 'utf8');
      console.log(`[Sync Taller] Unidades sincronizadas con orden ${orden.folio} (${nuevoEstatus})`);
    }
  } catch (err) {
    console.warn('[Sync Taller] Error sincronizando con taller_db.json:', err.message);
  }
}

// =========================================================
// ENDPOINTS DE APRENDIZAJE Y MEMORIA PERMANENTE IA
// =========================================================
app.post('/api/ia/aprender-vinculacion', (req, res) => {
  try {
    const { texto_detectado, codigo_detectado, marca_detectada, proveedor, id_pieza_catalogo, numero_parte_catalogo, origen } = req.body || {};
    const record = registrarAprendizajeIA({
      texto_detectado,
      codigo_detectado,
      marca_detectada,
      proveedor,
      id_pieza_catalogo,
      numero_parte_catalogo,
      origen: origen || 'correccion_manual'
    });
    if (!record) {
      return res.status(400).json({ success: false, error: 'No se pudo vincular la pieza al catálogo' });
    }
    res.json({
      success: true,
      mensaje: '¡Aprendido! La IA recordará esta refacción para futuras facturas y fotos.',
      record,
      totalMemoria: (store.memoria_ia || []).length
    });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/api/ia/memoria', (req, res) => {
  res.json({
    success: true,
    total: (store.memoria_ia || []).length,
    memoria: store.memoria_ia || []
  });
});

// =========================================================
// MÓDULO DE INTEGRACIÓN DE ÓRDENES Y UNIDADES DE TALLER PRO
// =========================================================
const TALLER_DB_PATHS = [
  path.join(__dirname, 'taller_db.json'),
  path.join(__dirname, '..', 'taller_db.json'),
  path.join(__dirname, '..', 'taller-cloud-deploy', 'taller_db.json')
];

function getTallerDbPath() {
  for (const p of TALLER_DB_PATHS) {
    if (fs.existsSync(p)) return p;
  }
  return TALLER_DB_PATHS[0];
}

function getTallerUnidades() {
  const p = getTallerDbPath();
  if (!fs.existsSync(p)) return [];
  try {
    return JSON.parse(fs.readFileSync(p, 'utf8'));
  } catch (err) {
    console.error('Error leyendo taller_db.json:', err.message);
    return [];
  }
}

function saveTallerUnidades(unidades) {
  try {
    const jsonStr = JSON.stringify(unidades, null, 2);
    for (const p of TALLER_DB_PATHS) {
      if (fs.existsSync(p) || p === TALLER_DB_PATHS[0]) {
        fs.writeFileSync(p, jsonStr, 'utf8');
      }
    }
    return true;
  } catch (err) {
    console.error('Error guardando taller_db.json:', err.message);
    return false;
  }
}

function resolveMovimientoCostoUnitario(m, catItem) {
  const catPrice = catItem && Number(catItem.precio_unitario) > 0 ? Number(catItem.precio_unitario) : 0;
  const movPrice = Number(m && (m.costo_unitario !== undefined && m.costo_unitario !== null ? m.costo_unitario : m.precio_unitario)) || 0;

  // 1. Si el catálogo tiene precio oficial actualizado en almacén (> 0), usar el precio más fresco de inventario
  if (catPrice > 0) {
    return catPrice;
  }
  // 2. Si el movimiento ya tiene precio válido registrado
  if (movPrice > 0) {
    return movPrice;
  }
  return 0;
}

function syncSalidaConTaller(mov) {
  if (!mov || (!mov.unidad_camion && !mov.folio)) return;
  try {
    const unidades = getTallerUnidades();
    const uNum = String(mov.unidad_camion || '').trim().toUpperCase();
    const uClean = uNum.replace(/^0+/, '');
    const folioMov = String(mov.folio || '').trim();

    let u = unidades.find(x => {
      const xId = String(x.id).trim().toUpperCase();
      const xClean = xId.replace(/^0+/, '');
      return (uNum && (xId === uNum || (uClean && xClean === uClean))) || (folioMov && String(x.folio_almacen || '').trim() === folioMov);
    });

    // Si la unidad no está en el tablero de taller, no crear tarjetas fantasma
    if (!u) {
      return;
    }

    if (u) {
      if (!u.folio_almacen && folioMov) {
        u.folio_almacen = folioMov;
      }
      if (!Array.isArray(u.piezas)) u.piezas = [];
      const itemCat = store.catalog[mov.id_pieza] || Object.values(store.catalog).find(c => c.numero_parte && c.numero_parte === mov.numero_parte) || {};
      const isSrv = Boolean(mov.es_servicio || mov.tipo_movimiento === 'SERVICIO' || mov.almacen === 'Servicios y Reparaciones Externas' || isServiceDescription(mov.descripcion));
      const pu = resolveMovimientoCostoUnitario(mov, itemCat);
      const cant = Number(mov.cantidad) || 1;
      const total = Number(mov.costo_total) && Number(mov.costo_total) > 10 ? Number(mov.costo_total) : (pu * cant);

      // Clasificación de refacción preventiva vs correctiva
      const txt = `${mov.descripcion || ''} ${mov.numero_parte || ''}`.toUpperCase();
      const esPrev = !isSrv && ['FILTRO', 'ACEITE', 'LUBRICANTE', 'ANTICONGELANTE', 'GRASA', 'CARTUCHO', 'FLUIDO', 'PREVENTIVO', 'SERVICIO', 'AFINACION', 'UREA', 'VALVULINA', 'REFRIGERANTE', 'COOLANT'].some(k => txt.includes(k));
      const tipoTrabajo = isSrv ? 'correctivo' : (esPrev ? 'preventivo' : 'correctivo');

      const pIdx = u.piezas.findIndex(p => p.id === mov.id || p.id_movimiento === mov.id);
      const piezaObj = {
        id: mov.id || `pz_alm_${Date.now()}`,
        id_movimiento: mov.id,
        id_pieza: mov.id_pieza,
        numero_parte: isSrv ? 'SERVICIO' : (mov.numero_parte || itemCat.numero_parte || ''),
        descripcion: mov.descripcion || itemCat.descripcion || (isSrv ? 'Servicio Externo' : 'Refacción'),
        marca: mov.marca || itemCat.marca || (isSrv ? (mov.proveedor || 'PROVEEDOR EXTERNO') : ''),
        cantidad: cant,
        unidad_medida: isSrv ? 'SERV' : (mov.unidad || itemCat.unidad || 'PZ'),
        precio_unitario: pu,
        costo_total: total,
        folio_almacen: folioMov || u.folio_almacen || '',
        solicitante: mov.solicitante || 'Almacén',
        estado: 'RECIBIDA_TALLER',
        tipo_trabajo: tipoTrabajo,
        tipo_item: isSrv ? 'SERVICIO' : 'REFACCION',
        es_servicio: isSrv,
        proveedor: mov.proveedor || (isSrv ? (mov.marca || '') : ''),
        factura: mov.factura || '',
        fecha_recepcion: mov.fecha || new Date().toISOString().slice(0, 10),
        hora_recepcion: mov.hora || ''
      };

      if (pIdx >= 0) u.piezas[pIdx] = { ...u.piezas[pIdx], ...piezaObj };
      else u.piezas.push(piezaObj);

      let cPrev = 0, cCorr = 0;
      u.piezas.forEach(pz => {
        const c = Number(pz.costo_total) || (Number(pz.cantidad || 1) * Number(pz.precio_unitario || 0));
        if (pz.tipo_trabajo === 'preventivo') cPrev += c;
        else cCorr += c;
      });
      u.costo_preventivo = cPrev;
      u.costo_correctivo = cCorr;
      u.costo_refacciones = cPrev + cCorr;
      u.costo = (Number(u.costo_mano_obra) || 0) + u.costo_refacciones;
      if (!u.tipo_procedimiento) {
        u.tipo_procedimiento = cPrev > 0 && cCorr > 0 ? 'mixto' : (cPrev > 0 ? 'preventivo' : 'correctivo');
      }

      if (!Array.isArray(u.historial)) u.historial = [];
      const eventoTexto = isSrv
        ? `🛠️ Servicio Externo registrado: ${cant > 1 ? cant + 'x ' : ''}${piezaObj.descripcion} por ${piezaObj.proveedor || 'Proveedor'} ($${total.toLocaleString('es-MX', { minimumFractionDigits: 2 })} MXN) [Fact/Rem: ${mov.factura || folioMov || 'S/F'}]`
        : `Salida de Almacén registrada: ${cant}x ${piezaObj.descripcion} ($${total.toLocaleString('es-MX', { minimumFractionDigits: 2 })} MXN) [Folio #${folioMov || u.folio_almacen}]`;

      u.historial.unshift({
        fecha: new Date().toLocaleString('es-MX'),
        evento: eventoTexto,
        tipo: isSrv ? 'servicio_externo' : 'almacen_salida',
        autor: mov.solicitante || 'Almacén'
      });

      saveTallerUnidades(unidades);
      console.log(`[Sync Taller] Refacción "${piezaObj.descripcion}" sincronizada con Unidad ${u.id} a $${pu}`);
      try {
        fetch(`https://scc-bv7q.onrender.com/api/unidad/${encodeURIComponent(u.id)}`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(u)
        }).catch(() => {});
      } catch(e) {}
    }
  } catch (err) {
    console.warn('[Sync Taller] Error en syncSalidaConTaller:', err.message);
  }
}

// MOTOR DE RECONCILIACIÓN INTEGRAL BIDIRECCIONAL ALMACÉN <-> TALLER FLOTILLA
function reconciliarAlmacenConTaller() {
  try {
    const unidades = getTallerUnidades();
    if (!Array.isArray(unidades) || unidades.length === 0) return { piezas: 0, unidades: 0 };

    const salidasPorUnidad = {};
    const salidasPorFolio = {};
    (store.movimientos || []).forEach(m => {
      if (!m.anulado && (m.tipo === 'S' || m.tipo === 'E/S' || !m.tipo)) {
        if (m.unidad_camion) {
          const uKey = String(m.unidad_camion).trim().toUpperCase();
          if (!salidasPorUnidad[uKey]) salidasPorUnidad[uKey] = [];
          salidasPorUnidad[uKey].push(m);
        }
        if (m.folio) {
          const fKey = String(m.folio).trim();
          if (!salidasPorFolio[fKey]) salidasPorFolio[fKey] = [];
          salidasPorFolio[fKey].push(m);
        }
      }
    });

    let totalPiezasSincronizadas = 0;
    let unidadesActualizadas = 0;

    unidades.forEach(u => {
      const uId = String(u.id).trim().toUpperCase();
      const uClean = uId.replace(/^0+/, '');

      const matches = [
        ...(salidasPorUnidad[uId] || []),
        ...(uClean && uClean !== uId ? (salidasPorUnidad[uClean] || []) : [])
      ];

      if (u.folio_almacen) {
        const fMatches = salidasPorFolio[String(u.folio_almacen).trim()] || [];
        fMatches.forEach(fm => {
          if (!matches.some(m => m.id === fm.id)) matches.push(fm);
        });
      }

      if (matches.length > 0) {
        if (!Array.isArray(u.piezas)) u.piezas = [];
        let cPrev = 0, cCorr = 0;

        matches.forEach(mov => {
          const itemCat = store.catalog[mov.id_pieza] || Object.values(store.catalog).find(c => c.numero_parte && c.numero_parte === mov.numero_parte) || {};
          const isSrv = Boolean(mov.es_servicio || mov.tipo_movimiento === 'SERVICIO' || mov.almacen === 'Servicios y Reparaciones Externas' || isServiceDescription(mov.descripcion));
          const pu = resolveMovimientoCostoUnitario(mov, itemCat);
          const cant = Number(mov.cantidad) || 1;
          const total = Number(mov.costo_total) && Number(mov.costo_total) > 10 ? Number(mov.costo_total) : (pu * cant);

          const txt = `${mov.descripcion || ''} ${mov.numero_parte || ''}`.toUpperCase();
          const esPrev = !isSrv && ['FILTRO', 'ACEITE', 'LUBRICANTE', 'ANTICONGELANTE', 'GRASA', 'CARTUCHO', 'FLUIDO', 'PREVENTIVO', 'SERVICIO', 'AFINACION', 'UREA', 'VALVULINA', 'REFRIGERANTE', 'COOLANT'].some(k => txt.includes(k));
          const tipoTrabajo = isSrv ? 'correctivo' : (esPrev ? 'preventivo' : 'correctivo');

          if (tipoTrabajo === 'preventivo') cPrev += total;
          else cCorr += total;

          const pIdx = u.piezas.findIndex(p => p.id === mov.id || p.id_movimiento === mov.id);
          const piezaObj = {
            id: mov.id || `pz_alm_${Date.now()}`,
            id_movimiento: mov.id,
            id_pieza: mov.id_pieza,
            numero_parte: isSrv ? 'SERVICIO' : (mov.numero_parte || itemCat.numero_parte || ''),
            descripcion: mov.descripcion || itemCat.descripcion || (isSrv ? 'Servicio Externo' : 'Refacción'),
            marca: mov.marca || itemCat.marca || (isSrv ? (mov.proveedor || 'PROVEEDOR EXTERNO') : ''),
            cantidad: cant,
            unidad_medida: isSrv ? 'SERV' : (mov.unidad || itemCat.unidad || 'PZ'),
            precio_unitario: pu,
            costo_total: total,
            folio_almacen: mov.folio || u.folio_almacen || '',
            solicitante: mov.solicitante || 'Almacén',
            estado: 'RECIBIDA_TALLER',
            tipo_trabajo: tipoTrabajo,
            tipo_item: isSrv ? 'SERVICIO' : 'REFACCION',
            es_servicio: isSrv,
            proveedor: mov.proveedor || (isSrv ? (mov.marca || '') : ''),
            factura: mov.factura || '',
            fecha_recepcion: mov.fecha || new Date().toISOString().slice(0, 10),
            hora_recepcion: mov.hora || ''
          };

          if (pIdx >= 0) u.piezas[pIdx] = { ...u.piezas[pIdx], ...piezaObj };
          else u.piezas.push(piezaObj);
          totalPiezasSincronizadas++;
        });

        u.costo_preventivo = cPrev;
        u.costo_correctivo = cCorr;
        u.costo_refacciones = cPrev + cCorr;
        u.costo = (Number(u.costo_mano_obra) || 0) + u.costo_refacciones;
        unidadesActualizadas++;
      }
    });

    saveTallerUnidades(unidades);
    console.log(`[Reconciliación Taller] Éxito: ${totalPiezasSincronizadas} refacciones sincronizadas en ${unidadesActualizadas} unidades activas.`);
    return { piezas: totalPiezasSincronizadas, unidades: unidadesActualizadas };
  } catch (err) {
    console.error('[Reconciliación Taller] Error:', err.message);
    return { error: err.message };
  }
}

// Endpoint de sincronización global bajo demanda
app.post('/api/taller/sincronizar-todo', (req, res) => {
  const result = reconciliarAlmacenConTaller();
  res.json({
    success: true,
    message: `¡Sincronización completa! Se actualizaron ${result.piezas || 0} refacciones en ${result.unidades || 0} unidades del Taller.`,
    ...result
  });
});

app.get('/api/taller/ordenes', (req, res) => {
  try {
    try { reconciliarAlmacenConTaller(); } catch(e) {}
    const rawUnidades = getTallerUnidades();
    const foliosAlmacenMap = {};
    (store.folios_taller || []).forEach(f => {
      foliosAlmacenMap[String(f.unidad).trim()] = f;
    });

    const salidasPorFolio = {};
    const salidasPorUnidad = {};
    (store.movimientos || []).forEach(m => {
      if (!m.anulado && (m.tipo === 'S' || m.tipo === 'E/S' || !m.tipo)) {
        if (m.folio) {
          const fKey = String(m.folio).trim();
          if (!salidasPorFolio[fKey]) salidasPorFolio[fKey] = [];
          salidasPorFolio[fKey].push(m);
        }
        if (m.unidad_camion) {
          const uKey = String(m.unidad_camion).trim().toUpperCase();
          if (!salidasPorUnidad[uKey]) salidasPorUnidad[uKey] = [];
          salidasPorUnidad[uKey].push(m);
        }
      }
    });

    const ordenes = rawUnidades.map((u, idx) => {
      const uId = String(u.id).trim();
      const uNumClean = uId.replace(/^0+/, ''); // '008' -> '8'
      const salidasUnidad = [
        ...(salidasPorUnidad[uId.toUpperCase()] || []),
        ...(uNumClean && uNumClean !== uId ? (salidasPorUnidad[uNumClean.toUpperCase()] || []) : [])
      ];

      // Buscar si algún movimiento de esta unidad trae folio
      const movConFolio = salidasUnidad.find(m => m.folio && String(m.folio).trim());
      const folioDetectado = u.folio_almacen || (foliosAlmacenMap[uId] ? foliosAlmacenMap[uId].folio : '') || (movConFolio ? movConFolio.folio : '');

      const salidasFolio = folioDetectado ? (salidasPorFolio[folioDetectado] || []) : [];

      // Combinar todas las salidas surtidas para esta unidad sin duplicar
      const salidasMap = {};
      salidasUnidad.forEach(m => { if (!m.anulado) salidasMap[m.id] = m; });
      salidasFolio.forEach(m => { if (!m.anulado) salidasMap[m.id] = m; });
      const todasLasSalidas = Object.values(salidasMap).map(m => {
        let catItem = null;
        if (m.id_pieza && store.catalog[m.id_pieza]) {
          catItem = store.catalog[m.id_pieza];
        } else if (m.numero_parte) {
          catItem = Object.values(store.catalog).find(p => p.numero_parte === m.numero_parte);
        } else if (m.descripcion) {
          catItem = Object.values(store.catalog).find(p => p.descripcion && p.descripcion.toLowerCase() === m.descripcion.toLowerCase());
        }

        const costUnit = resolveMovimientoCostoUnitario(m, catItem);

        const cant = Number(m.cantidad) || 1;
        return {
          ...m,
          costo_unitario: costUnit,
          costo_total: costUnit * cant
        };
      }).sort((a,b) => new Date(b.timestamp || b.fecha).getTime() - new Date(a.timestamp || a.fecha).getTime());

      const totalPiezasSurtidas = todasLasSalidas.reduce((sum, m) => sum + (Number(m.cantidad) || 0), 0);
      const costoTotalSurtido = todasLasSalidas.reduce((sum, m) => sum + (Number(m.costo_total) || 0), 0);

      const piezasSolicitadas = (u.piezas || []).map(pz => {
        let catItem = null;
        if (pz.id_pieza && store.catalog[pz.id_pieza]) {
          catItem = store.catalog[pz.id_pieza];
        } else if (pz.numero_parte) {
          catItem = Object.values(store.catalog).find(p => p.numero_parte === pz.numero_parte);
        } else if (pz.descripcion) {
          const matchRes = matchItemToCatalog('', pz.descripcion, pz.marca);
          if (matchRes && matchRes.match) catItem = matchRes.match;
        }

        return {
          id: pz.id || `pz_${idx}`,
          id_pieza: catItem ? catItem.id : pz.id_pieza,
          numero_parte: pz.numero_parte || (catItem ? catItem.numero_parte : ''),
          descripcion: pz.descripcion || (catItem ? catItem.descripcion : 'Refacción'),
          marca: pz.marca || (catItem ? catItem.marca : ''),
          cantidad: Number(pz.cantidad || 1),
          unidad_medida: pz.unidad_medida || 'PZ',
          precio_unitario: Number(pz.precio_unitario || (catItem ? catItem.precio_unitario : 0)),
          stock_actual: catItem ? Number(catItem.stock_actual || 0) : 0,
          rack: catItem ? catItem.rack : 'Sin estante',
          nivel: catItem ? catItem.nivel : 'Sin nivel',
          estado: pz.estado || 'SOLICITADA_ALMACEN',
          folio_oc: pz.folio_oc || ''
        };
      });

      return {
        id: uId,
        unidad: uId,
        sucursal: u.sucursal || 'Apodaca',
        estado: u.estado || 'proceso',
        estado_label: u.estado === 'listo' ? 'Listas para Entrega' : (u.estado === 'mayor' ? 'Mtto. Mayor' : (u.estado === 'espera' ? 'En Espera' : 'En Proceso (Taller)')),
        defecto: u.defecto || u.reporte || 'Reparación / Mantenimiento en Taller',
        reporte: u.reporte || u.defecto || '',
        notas: u.notas || '',
        categoriaFalla: u.categoriaFalla || 'General',
        folio_almacen: folioDetectado || '',
        folio_taller: folioDetectado || '',
        fechaRefaccion: u.fechaRefaccion || '',
        piezas: piezasSolicitadas,
        total_piezas_solicitadas: piezasSolicitadas.length,
        total_piezas_entregadas: totalPiezasSurtidas,
        costo_refacciones: costoTotalSurtido > 0 ? costoTotalSurtido : Number(u.costo_refacciones || u.costo || 0),
        costo_total_surtido: costoTotalSurtido,
        movimientos_salida: todasLasSalidas
      };
    });

    res.json({
      success: true,
      total: ordenes.length,
      unidades_en_proceso: ordenes.filter(o => o.estado === 'proceso' || o.estado === 'mayor').length,
      unidades_listas: ordenes.filter(o => o.estado === 'listo').length,
      ordenes
    });
  } catch (err) {
    console.error('Error obteniendo órdenes de taller:', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post('/api/taller/ordenes/:id/asignar-folio', (req, res) => {
  try {
    const unitId = String(req.params.id).trim();
    let { folio, motivo, solicitante, quitar } = req.body || {};
    let unidades = getTallerUnidades();
    let unidad = unidades.find(u => String(u.id) === unitId);
    if (!unidad) return res.status(404).json({ success: false, error: 'Unidad no encontrada en taller' });

    // Si el usuario solicitó quitar o vaciar el folio
    if (quitar === true || folio === '') {
      const oldFolio = unidad.folio_almacen || unidad.folio || '';
      unidad.folio_almacen = '';
      unidad.folio = '';
      saveTallerUnidades(unidades);
      if (Array.isArray(store.folios_taller)) {
        store.folios_taller = store.folios_taller.filter(f => String(f.unidad).trim() !== unitId && String(f.folio).trim() !== oldFolio);
      }
      saveStore();
      return res.json({ success: true, folio: '', message: 'Folio removido de la unidad correctamente', unidad });
    }

    const folioManual = String(folio || '').trim().toUpperCase();
    if (!folioManual) {
      return res.status(400).json({ success: false, error: 'Debes indicar el número de folio físico real a vincular' });
    }

    unidad.folio_almacen = folioManual;
    unidad.folio = folioManual;
    saveTallerUnidades(unidades);

    if (!Array.isArray(store.folios_taller)) store.folios_taller = [];
    // Actualizar o insertar registro
    const existingIdx = store.folios_taller.findIndex(f => String(f.unidad).trim() === unitId || String(f.folio).trim() === folioManual);
    const regFolio = {
      folio: folioManual,
      unidad: unitId,
      solicitante: solicitante || 'JEFE DE TALLER',
      motivo: motivo || unidad.reporte || unidad.defecto || 'Reparación en taller',
      fecha: new Date().toISOString().slice(0, 10),
      hora: new Date().toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' }),
      estado_taller: unidad.estado || 'proceso',
      timestamp: new Date().toISOString()
    };
    if (existingIdx !== -1) {
      store.folios_taller[existingIdx] = regFolio;
    } else {
      store.folios_taller.unshift(regFolio);
    }
    saveStore();

    res.json({ success: true, folio: folioManual, message: `Folio #${folioManual} asignado a Unidad ${unitId}`, unidad });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// 12.2 Listar Órdenes de Compra generadas
app.get('/api/compras/ordenes', (req, res) => {
  try {
    if (!Array.isArray(store.ordenes_compra)) store.ordenes_compra = [];
    const list = [...store.ordenes_compra].sort((a, b) => {
      const ta = new Date(a.timestamp || a.fecha_creacion).getTime();
      const tb = new Date(b.timestamp || b.fecha_creacion).getTime();
      return tb - ta;
    });
    res.json({ success: true, total: list.length, ordenes: list });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// 12.3 Crear nueva Orden de Compra
app.post('/api/compras/ordenes', (req, res) => {
  try {
    if (req.body.usuarioRol === 'GERENTE') {
      return res.status(403).json({ error: 'Acceso denegado: Tu perfil es de Solo Consulta / Auditoría.' });
    }

    const {
      proveedor,
      solicitante,
      prioridad,
      notas,
      almacen_id,
      fecha_entrega_esperada,
      items
    } = req.body || {};

    if (!Array.isArray(items) || items.length === 0) {
      return res.status(400).json({ error: 'Debes incluir al menos una refacción en la orden de compra.' });
    }

    if (!Array.isArray(store.ordenes_compra)) store.ordenes_compra = [];

    const now = new Date();
    const anio = now.getFullYear();
    const seq = store.ordenes_compra.length + 1;
    const folio = `OC-${anio}-${String(seq).padStart(3, '0')}`;

    let totalPiezas = 0;
    let totalImporte = 0;

    const itemsProcesados = items.map((it, idx) => {
      const cant = Math.max(1, Number(it.cantidad_pedida || it.cantidad || 1));
      const precio = Math.max(0, Number(it.precio_estimado || it.precio_unitario || 0));
      const totalItem = cant * precio;
      totalPiezas += cant;
      totalImporte += totalItem;

      return {
        id: it.id || `item_${Date.now()}_${idx}`,
        id_pieza: it.id_pieza || it.id || '',
        numero_parte: String(it.numero_parte || '').trim(),
        descripcion: String(it.descripcion || 'Refacción sin descripción').trim(),
        marca: String(it.marca || 'GENERICA').trim(),
        unidad: it.unidad || 'PZ',
        stock_actual_al_pedir: Number(it.stock_actual !== undefined ? it.stock_actual : 0),
        cantidad_pedida: cant,
        precio_estimado: precio,
        total_estimado: totalItem
      };
    });

    const almObj = (store.almacenes || []).find(a => a.id === (almacen_id || 'refacciones'));

    const nuevaOrden = {
      id: `oc_${Date.now()}_${Math.round(Math.random() * 1000)}`,
      folio,
      fecha_creacion: now.toISOString().slice(0, 10),
      hora_creacion: now.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit', second: '2-digit' }),
      proveedor: String(proveedor || 'PROVEEDOR POR DEFINIR').trim().toUpperCase(),
      solicitante: String(solicitante || 'César Almacén (Administrador)').trim(),
      estatus: 'PENDIENTE', // 'BORRADOR' | 'PENDIENTE' | 'ORDENADA' | 'RECIBIDA' | 'CANCELADA'
      prioridad: prioridad || 'ALTA', // 'NORMAL' | 'ALTA' | 'URGENTE'
      notas: String(notas || '').trim(),
      almacen_id: almacen_id || 'refacciones',
      almacen: almObj ? almObj.nombre : 'Almacén de Refacciones',
      fecha_entrega_esperada: fecha_entrega_esperada || '',
      items: itemsProcesados,
      total_items: itemsProcesados.length,
      total_piezas: totalPiezas,
      total_importe: totalImporte,
      timestamp: now.toISOString()
    };

    store.ordenes_compra.unshift(nuevaOrden);
    saveStore();

    res.json({
      success: true,
      message: `Orden de Compra ${folio} generada con éxito con ${itemsProcesados.length} partidas.`,
      orden: nuevaOrden
    });
  } catch (err) {
    console.error('Error al crear orden de compra:', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// 12.4 Actualizar estatus o datos de Orden de Compra
app.put('/api/compras/ordenes/:id', (req, res) => {
  try {
    const { id } = req.params;
    const { estatus, notas, proveedor, fecha_entrega_esperada, prioridad } = req.body || {};

    if (!Array.isArray(store.ordenes_compra)) store.ordenes_compra = [];
    const orden = store.ordenes_compra.find(o => o.id === id || o.folio === id);

    if (!orden) {
      return res.status(404).json({ success: false, error: 'Orden de compra no encontrada.' });
    }

    if (estatus) orden.estatus = estatus;
    if (notas !== undefined) orden.notas = String(notas).trim();
    if (proveedor) orden.proveedor = String(proveedor).trim().toUpperCase();
    if (fecha_entrega_esperada !== undefined) orden.fecha_entrega_esperada = fecha_entrega_esperada;
    if (prioridad) orden.prioridad = prioridad;

    orden.fecha_actualizacion = new Date().toISOString();
    saveStore();

    // Sincronizar automáticamente con Taller Pro (Control CESS)
    syncOrdenConTaller(orden, orden.estatus, 'Administrador de Almacén');

    res.json({
      success: true,
      message: `Orden ${orden.folio} actualizada a estatus: ${orden.estatus}`,
      orden
    });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// 12.5 Eliminar Orden de Compra (solo borradores o pendientes)
app.delete('/api/compras/ordenes/:id', (req, res) => {
  try {
    const { id } = req.params;
    if (!Array.isArray(store.ordenes_compra)) store.ordenes_compra = [];
    const idx = store.ordenes_compra.findIndex(o => o.id === id || o.folio === id);

    if (idx === -1) {
      return res.status(404).json({ success: false, error: 'Orden de compra no encontrada.' });
    }

    const eliminado = store.ordenes_compra.splice(idx, 1)[0];
    saveStore();

    // Sincronizar cancelación con Taller Pro
    syncOrdenConTaller(eliminado, 'CANCELADA', 'Administrador de Almacén');

    res.json({
      success: true,
      message: `Orden de Compra ${eliminado.folio} eliminada del sistema.`
    });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// 12.6 Recepción Directa: Dar entrada al almacén automáticamente desde la Orden de Compra
app.post('/api/compras/ordenes/:id/recibir', (req, res) => {
  try {
    const { id } = req.params;
    const { usuarioNombre, factura, entregado_por } = req.body || {};

    if (!Array.isArray(store.ordenes_compra)) store.ordenes_compra = [];
    const orden = store.ordenes_compra.find(o => o.id === id || o.folio === id);

    if (!orden) {
      return res.status(404).json({ success: false, error: 'Orden de compra no encontrada.' });
    }

    if (orden.estatus === 'RECIBIDA') {
      return res.status(400).json({ success: false, error: 'Esta orden de compra ya fue recibida en almacén anteriormente.' });
    }

    const now = new Date();
    const fechaHoy = now.toISOString().slice(0, 10);
    const horaHoy = now.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit', second: '2-digit' });

    const condicionPago = String(req.body.condicion_pago || orden.condicion_pago || 'CONTADO').trim().toUpperCase() === 'CREDITO' ? 'CREDITO' : 'CONTADO';
    const diasCredito = condicionPago === 'CREDITO' ? (Math.max(1, Number(req.body.dias_credito || orden.dias_credito) || 30)) : 0;
    const fechaFactura = req.body.fecha_factura ? String(req.body.fecha_factura).trim() : fechaHoy;
    const fechaVencimiento = condicionPago === 'CREDITO'
      ? (req.body.fecha_vencimiento ? String(req.body.fecha_vencimiento).trim() : calcularVencimiento(fechaFactura, diasCredito))
      : fechaFactura;
    const estadoPago = condicionPago === 'CREDITO' ? (req.body.estado_pago || 'PENDIENTE') : 'PAGADO';
    const fechaPago = req.body.fecha_pago || (estadoPago === 'PAGADO' && condicionPago === 'CONTADO' ? fechaHoy : '');
    const referenciaPago = req.body.referencia_pago || '';

    let piezasIngresadas = 0;

    // Procesar cada partida e ingresar el stock al catálogo
    for (const it of (orden.items || [])) {
      let item = findPart(it.id_pieza) || findPart(it.numero_parte);
      const cant = Number(it.cantidad_pedida) || 1;

      if (item) {
        item.stock_actual = (Number(item.stock_actual) || 0) + cant;
      } else {
        const safeId = (orden.almacen_id !== 'refacciones' ? `${orden.almacen_id}_` : '') + it.numero_parte.replace(/[^a-zA-Z0-9_-]/g, '_');
        item = {
          id: safeId,
          rack: 'Rack 1',
          nivel: 'Nivel 1',
          descripcion: it.descripcion,
          marca: it.marca || 'GENERICA',
          numero_parte: it.numero_parte,
          unidad: it.unidad || 'PZ',
          precio_unitario: Number(it.precio_estimado) || 0,
          stock_inicial: 0,
          stock_actual: cant,
          almacen: orden.almacen || 'Almacén de Refacciones',
          almacen_id: orden.almacen_id || 'refacciones'
        };
        store.catalog[safeId] = item;
      }

      // Registrar movimiento de Entrada en Bitácora
      const movEntrada = {
        id: `mov_in_oc_${Date.now()}_${Math.round(Math.random() * 1000)}`,
        tipo: 'E',
        fecha: fechaHoy,
        hora: horaHoy,
        id_pieza: item.id,
        numero_parte: item.numero_parte,
        descripcion: item.descripcion,
        marca: item.marca,
        almacen: item.almacen,
        almacen_id: item.almacen_id,
        ubicacion: `${item.rack} / ${item.nivel}`,
        cantidad: cant,
        proveedor: orden.proveedor,
        factura: factura || orden.folio,
        entregado_por: entregado_por || 'Recepción de Orden de Compra',
        solicitante: usuarioNombre || 'Almacén Central',
        costo_unitario: it.precio_estimado || item.precio_unitario,
        condicion_pago: condicionPago,
        dias_credito: diasCredito,
        fecha_factura: fechaFactura,
        fecha_vencimiento: fechaVencimiento,
        estado_pago: estadoPago,
        fecha_pago: fechaPago,
        referencia_pago: referenciaPago,
        notas: `[RECEPCIÓN OC: ${orden.folio} | Proveedor: ${orden.proveedor}]`,
        timestamp: now.toISOString()
      };

      store.movimientos.unshift(movEntrada);
      registrarOActualizarFacturaProcesada(movEntrada);
      piezasIngresadas += cant;
    }

    // Actualizar estado de la orden
    orden.estatus = 'RECIBIDA';
    orden.fecha_recepcion = fechaHoy;
    orden.hora_recepcion = horaHoy;
    orden.recibido_por = usuarioNombre || 'Administrador de Almacén';
    orden.factura_asociada = factura || orden.folio;
    orden.condicion_pago = condicionPago;
    orden.dias_credito = diasCredito;
    orden.fecha_vencimiento = fechaVencimiento;
    orden.estado_pago = estadoPago;

    saveStore();

    // Sincronizar automáticamente recepción con Taller Pro (Control CESS)
    syncOrdenConTaller(orden, 'RECIBIDA', usuarioNombre || 'Administrador de Almacén', factura);

    res.json({
      success: true,
      message: `Orden ${orden.folio} recibida con éxito: ${piezasIngresadas} piezas ingresadas directamente al inventario.`,
      orden
    });
  } catch (err) {
    console.error('Error al recibir orden de compra:', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// 12.7 Resumen de Cuentas por Pagar y Vencimientos de Crédito
app.get('/api/compras/creditos-resumen', (req, res) => {
  try {
    if (!Array.isArray(store.facturas_procesadas)) store.facturas_procesadas = [];
    const hoyStr = new Date().toISOString().slice(0, 10);
    const hoyMs = new Date(hoyStr + 'T12:00:00').getTime();

    const facturas = store.facturas_procesadas.map(fac => {
      const cond = String(fac.condicion_pago || 'CONTADO').toUpperCase();
      const esCredito = cond === 'CREDITO';
      const diasCred = Number(fac.dias_credito) || (esCredito ? 30 : 0);
      const fFact = fac.fecha_factura || fac.fecha_ingreso || hoyStr;
      const fVenc = fac.fecha_vencimiento || (esCredito ? calcularVencimiento(fFact, diasCred) : fFact);

      let diasRestantes = null;
      let estatusVencimiento = 'CONTADO';

      if (esCredito) {
        if (fac.estado_pago === 'PAGADO') {
          estatusVencimiento = 'PAGADO';
          diasRestantes = 0;
        } else {
          const vencMs = new Date(fVenc + 'T12:00:00').getTime();
          diasRestantes = Math.round((vencMs - hoyMs) / (1000 * 60 * 60 * 24));
          if (diasRestantes < 0) {
            estatusVencimiento = 'VENCIDO';
          } else if (diasRestantes <= 5) {
            estatusVencimiento = 'POR_VENCER';
          } else {
            estatusVencimiento = 'VIGENTE';
          }
        }
      }

      return {
        ...fac,
        condicion_pago: esCredito ? 'CREDITO' : 'CONTADO',
        dias_credito: diasCred,
        fecha_factura: fFact,
        fecha_vencimiento: fVenc,
        dias_restantes: diasRestantes,
        estatus_vencimiento: estatusVencimiento
      };
    });

    const creditos = facturas.filter(f => f.condicion_pago === 'CREDITO');

    const vencidas = creditos.filter(f => f.estatus_vencimiento === 'VENCIDO');
    const porVencer = creditos.filter(f => f.estatus_vencimiento === 'POR_VENCER');
    const vigentes = creditos.filter(f => f.estatus_vencimiento === 'VIGENTE');
    const pagadas = creditos.filter(f => f.estatus_vencimiento === 'PAGADO');

    const totalVencidoMonto = vencidas.reduce((s, f) => s + (Number(f.total_monto) || 0), 0);
    const totalPorVencerMonto = porVencer.reduce((s, f) => s + (Number(f.total_monto) || 0), 0);
    const totalVigenteMonto = vigentes.reduce((s, f) => s + (Number(f.total_monto) || 0), 0);
    const totalPagadoMonto = pagadas.reduce((s, f) => s + (Number(f.total_monto) || 0), 0);
    const totalPendienteMonto = totalVencidoMonto + totalPorVencerMonto + totalVigenteMonto;

    // Ordenar facturas de crédito: Vencidas (urgentes primero) -> Por vencer -> Vigentes -> Pagadas
    const ordenEstatus = { 'VENCIDO': 1, 'POR_VENCER': 2, 'VIGENTE': 3, 'PAGADO': 4 };
    const facturasOrdenadas = [...creditos].sort((a, b) => {
      const ordA = ordenEstatus[a.estatus_vencimiento] || 99;
      const ordB = ordenEstatus[b.estatus_vencimiento] || 99;
      if (ordA !== ordB) return ordA - ordB;
      if (a.dias_restantes !== null && b.dias_restantes !== null) {
        return a.dias_restantes - b.dias_restantes;
      }
      return (b.fecha_vencimiento || '').localeCompare(a.fecha_vencimiento || '');
    });

    res.json({
      success: true,
      kpis: {
        total_pendiente_monto: totalPendienteMonto,
        total_vencido_monto: totalVencidoMonto,
        total_por_vencer_monto: totalPorVencerMonto,
        total_vigente_monto: totalVigenteMonto,
        total_pagado_monto: totalPagadoMonto,
        conteo_vencidas: vencidas.length,
        conteo_por_vencer: porVencer.length,
        conteo_vigentes: vigentes.length,
        conteo_pagadas: pagadas.length,
        total_creditos: creditos.length,
        total_todas_facturas: facturas.length
      },
      creditos: facturasOrdenadas,
      todas: facturas
    });
  } catch (err) {
    console.error('Error al generar resumen de créditos:', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// 12.8 Actualizar pago o prórroga de factura
app.post('/api/compras/actualizar-pago-factura', (req, res) => {
  try {
    const {
      factura,
      proveedor,
      estado_pago, // 'PAGADO' | 'PENDIENTE'
      fecha_pago,
      referencia_pago,
      notas_pago,
      dias_credito,
      fecha_vencimiento
    } = req.body || {};

    if (!factura) {
      return res.status(400).json({ success: false, error: 'Se requiere el número o folio de factura.' });
    }

    if (!Array.isArray(store.facturas_procesadas)) store.facturas_procesadas = [];
    const factNorm = normalizeFacturaKey(factura);
    const provNorm = proveedor ? normalizeProveedorKey(proveedor) : '';

    let fac = store.facturas_procesadas.find(f =>
      f.factura_normalizada === factNorm && (!provNorm || f.proveedor_normalizado === provNorm)
    );

    const now = new Date();
    const hoyStr = now.toISOString().slice(0, 10);

    if (fac) {
      if (estado_pago) fac.estado_pago = String(estado_pago).toUpperCase();
      if (fecha_pago !== undefined) fac.fecha_pago = fecha_pago || (fac.estado_pago === 'PAGADO' ? hoyStr : '');
      if (referencia_pago !== undefined) fac.referencia_pago = String(referencia_pago).trim();
      if (notas_pago !== undefined) fac.notas_pago = String(notas_pago).trim();
      if (dias_credito !== undefined && Number(dias_credito) > 0) fac.dias_credito = Number(dias_credito);
      if (fecha_vencimiento) fac.fecha_vencimiento = String(fecha_vencimiento).trim();
      else if (dias_credito && fac.fecha_factura) {
        fac.fecha_vencimiento = calcularVencimiento(fac.fecha_factura, fac.dias_credito);
      }
    }

    // Sincronizar en los movimientos de bitácora vinculados
    let movsActualizados = 0;
    for (const m of store.movimientos) {
      if (m.factura && normalizeFacturaKey(m.factura) === factNorm && (!provNorm || normalizeProveedorKey(m.proveedor) === provNorm)) {
        if (estado_pago) m.estado_pago = String(estado_pago).toUpperCase();
        if (fecha_pago !== undefined) m.fecha_pago = fecha_pago || (m.estado_pago === 'PAGADO' ? hoyStr : '');
        if (referencia_pago !== undefined) m.referencia_pago = String(referencia_pago).trim();
        if (dias_credito !== undefined && Number(dias_credito) > 0) m.dias_credito = Number(dias_credito);
        if (fecha_vencimiento) m.fecha_vencimiento = String(fecha_vencimiento).trim();
        else if (fac && fac.fecha_vencimiento) m.fecha_vencimiento = fac.fecha_vencimiento;
        movsActualizados++;
      }
    }

    saveStore();

    res.json({
      success: true,
      message: `Factura ${factura} actualizada con éxito (${movsActualizados} movimientos sincronizados).`,
      factura: fac
    });
  } catch (err) {
    console.error('Error al actualizar pago de factura:', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// 12.9 Convertir factura entre Contado y Crédito
app.post('/api/compras/convertir-factura-credito', (req, res) => {
  try {
    const {
      factura,
      proveedor,
      condicion_pago, // 'CREDITO' | 'CONTADO'
      dias_credito,
      fecha_vencimiento,
      estado_pago
    } = req.body || {};

    if (!factura) {
      return res.status(400).json({ success: false, error: 'Se requiere el folio de factura.' });
    }

    const nuevaCond = String(condicion_pago || 'CREDITO').trim().toUpperCase() === 'CREDITO' ? 'CREDITO' : 'CONTADO';
    const dias = nuevaCond === 'CREDITO' ? Math.max(1, Number(dias_credito) || 30) : 0;
    const factNorm = normalizeFacturaKey(factura);
    const provNorm = proveedor ? normalizeProveedorKey(proveedor) : '';

    let fac = store.facturas_procesadas.find(f =>
      f.factura_normalizada === factNorm && (!provNorm || f.proveedor_normalizado === provNorm)
    );

    const baseFecha = fac ? (fac.fecha_factura || fac.fecha_ingreso) : new Date().toISOString().slice(0, 10);
    const fVenc = nuevaCond === 'CREDITO'
      ? (fecha_vencimiento || calcularVencimiento(baseFecha, dias))
      : baseFecha;
    const estPago = nuevaCond === 'CREDITO' ? (estado_pago || 'PENDIENTE') : 'PAGADO';

    if (fac) {
      fac.condicion_pago = nuevaCond;
      fac.dias_credito = dias;
      fac.fecha_vencimiento = fVenc;
      fac.estado_pago = estPago;
    }

    for (const m of store.movimientos) {
      if (m.factura && normalizeFacturaKey(m.factura) === factNorm && (!provNorm || normalizeProveedorKey(m.proveedor) === provNorm)) {
        m.condicion_pago = nuevaCond;
        m.dias_credito = dias;
        m.fecha_vencimiento = fVenc;
        m.estado_pago = estPago;
      }
    }

    saveStore();

    res.json({
      success: true,
      message: `Condición de factura ${factura} cambiada a ${nuevaCond}.`,
      factura: fac
    });
  } catch (err) {
    console.error('Error al convertir factura de crédito:', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// Rutas de compatibilidad directa para Suite CESS: /almacen y /combustible
app.get(['/almacen', '/Almacen', '/almacen.com', '/Almacen.com'], (req, res) => {
  const DIST_DIR = path.join(__dirname, 'frontend', 'dist');
  if (fs.existsSync(path.join(DIST_DIR, 'index.html'))) {
    return res.sendFile(path.join(DIST_DIR, 'index.html'));
  }
  res.redirect('/');
});

app.get(['/combustible', '/Combustible', '/combustible.com', '/Combustible.com'], (req, res) => {
  // Portal de enlace para el Módulo de Control de Combustible y Rendimientos Diésel CESS
  res.send(`
    <!DOCTYPE html>
    <html lang="es">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Sistema de Control CESS - Combustible Diésel</title>
      <style>
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background-color: #0F172A; color: #FFFFFF; display: flex; align-items: center; justify-content: center; min-height: 100vh; margin: 0; padding: 20px; box-sizing: border-box; }
        .card { background: #1E293B; border: 2px solid #C59B27; border-radius: 14px; padding: 36px; max-width: 580px; width: 100%; text-align: center; box-shadow: 0 25px 50px -12px rgba(0,0,0,0.5); }
        .badge { background: #C59B27; color: #0F172A; font-weight: 800; font-size: 11px; padding: 4px 10px; border-radius: 6px; letter-spacing: 1px; display: inline-block; margin-bottom: 14px; text-transform: uppercase; }
        h1 { font-size: 24px; font-weight: 900; margin: 0 0 10px 0; color: #FFFFFF; letter-spacing: 0.5px; }
        p { font-size: 14px; color: #94A3B8; line-height: 1.6; margin: 0 0 24px 0; }
        .btn-group { display: flex; gap: 12px; justify-content: center; flex-wrap: wrap; }
        .btn { padding: 12px 22px; border-radius: 8px; font-weight: 800; text-decoration: none; font-size: 14px; transition: all 0.2s; cursor: pointer; display: inline-flex; align-items: center; gap: 8px; }
        .btn-gold { background: #C59B27; color: #0F172A; border: none; }
        .btn-gold:hover { background: #D4AF37; }
        .btn-outline { background: transparent; color: #FFFFFF; border: 1.5px solid #334155; }
        .btn-outline:hover { border-color: #C59B27; color: #C59B27; }
      </style>
    </head>
    <body>
      <div class="card">
        <span class="badge">Suite Sistema de Control CESS</span>
        <h1>Módulo de Combustible Diésel</h1>
        <p>Este módulo administra el despacho de diésel, kilometraje, rendimiento (Km/L) por unidad y tanques de autoconsumo.<br><br>Para acceder al almacén general de refacciones, llantas y ferretería, haz clic a continuación:</p>
        <div class="btn-group">
          <a href="/almacen" class="btn btn-gold">Ir a CESS Almacén</a>
          <a href="http://localhost:5000" class="btn btn-outline" target="_blank">Abrir App Rendimientos Diésel</a>
        </div>
      </div>
    </body>
    </html>
  `);
});

const DIST_DIR = fs.existsSync(path.join(__dirname, 'almacen_dist'))
  ? path.join(__dirname, 'almacen_dist')
  : (fs.existsSync(path.join(__dirname, 'frontend', 'dist'))
      ? path.join(__dirname, 'frontend', 'dist')
      : path.join(__dirname, 'almacen-app', 'frontend', 'dist'));

if (fs.existsSync(DIST_DIR)) {
  const setNoCacheHeaders = (res, filePath) => {
    if (filePath && filePath.endsWith('.html')) {
      res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
      res.setHeader('Pragma', 'no-cache');
      res.setHeader('Expires', '0');
    }
  };

  // Servir archivos estáticos en /almacen y también en /assets como respaldo
  app.use('/almacen', express.static(DIST_DIR, { setHeaders: setNoCacheHeaders }));
  app.use('/assets', express.static(path.join(DIST_DIR, 'assets')));

  // Servir index.html de Almacén en /almacen y cualquier subruta de almacén
  app.get(/^\/almacen(\/.*)?$/, (req, res, next) => {
    if (req.path.startsWith('/api') || req.path.startsWith('/uploads')) {
      return next();
    }
    res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('Expires', '0');
    return res.sendFile(path.join(DIST_DIR, 'index.html'));
  });
}

// Obtener IP local para mostrar en consola
const os = require('os');
function getLocalIp() {
  const interfaces = os.networkInterfaces();
  for (const name of Object.keys(interfaces)) {
    for (const iface of interfaces[name]) {
      if (iface.family === 'IPv4' && !iface.internal) {
        return iface.address;
      }
    }
  }
  return 'localhost';
}

if (require.main === module) {
  // Redirigir raíz a /almacen en modo standalone (puerto 3001)
  app.get('/', (req, res) => res.redirect('/almacen'));

  app.listen(PORT, '0.0.0.0', () => {
    const localIp = getLocalIp();
    console.log(`\n======================================================`);
    console.log(` 🚀 SISTEMA DE CONTROL CESS - ALMACÉN ACTIVO`);
    console.log(` 💻 Local:            http://localhost:${PORT}/almacen`);
    console.log(` 🏢 URL SCC Almacén:  http://localhost:${PORT}/almacen`);
    console.log(` 📱 Red Local (LAN):  http://${localIp}:${PORT}/almacen`);
    console.log(`======================================================\n`);
  });
}

module.exports = app;

