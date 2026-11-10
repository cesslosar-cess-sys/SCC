let express, cors, multer, xlsx;
try {
    express = require('express');
} catch (e) {
    express = require('./almacen-app/node_modules/express');
}
try {
    cors = require('cors');
} catch (e) {
    cors = require('./almacen-app/node_modules/cors');
}
try {
    multer = require('multer');
} catch (e) {
    multer = require('./almacen-app/node_modules/multer');
}
try {
    xlsx = require('xlsx');
} catch (e) {
    xlsx = require('./almacen-app/node_modules/xlsx');
}
const fs = require('fs');
const path = require('path');
const os = require('os');

const app = express();
const PORT = process.env.PORT || 8080;

app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Anti-Caché estricto para todas las llamadas API
app.use('/api', (req, res, next) => {
    res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate, max-age=0');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('Expires', '0');
    res.setHeader('Surrogate-Control', 'no-store');
    next();
});

const UPLOADS_DIR = path.join(__dirname, 'uploads');
if (!fs.existsSync(UPLOADS_DIR)) {
    fs.mkdirSync(UPLOADS_DIR, { recursive: true });
}
app.use('/uploads', express.static(UPLOADS_DIR));

const DB_FILE = path.join(__dirname, 'taller_db.json');
const BITACORAS_FILE = path.join(__dirname, 'taller_bitacoras_historico.json');
const JORNADA_FILE = path.join(__dirname, 'taller_jornada_actual.json');
const EXCEL_FILE = path.join(__dirname, 'Taller Flotilla-v2.xlsx');
const ALMACEN_DB_FILE = fs.existsSync(path.join(__dirname, 'almacen_datos.json'))
    ? path.join(__dirname, 'almacen_datos.json')
    : path.join(__dirname, 'almacen-app', 'almacen_datos.json');
const PREVENTIVOS_FILE = path.join(__dirname, 'taller_preventivos.json');
const RESCATES_FILE = path.join(__dirname, 'taller_rescates.json');
const TITULOS_COLUMNAS_FILE = path.join(__dirname, 'taller_config_columnas.json');
const REMOLQUES_FILE = path.join(__dirname, 'taller_remolques.json');
const ARCHIVADAS_FILE = path.join(__dirname, 'taller_unidades_archivadas.json');

function getUnidadesArchivadas() {
    if (!fs.existsSync(ARCHIVADAS_FILE)) return [];
    try {
        return JSON.parse(fs.readFileSync(ARCHIVADAS_FILE, 'utf8'));
    } catch(e) {
        return [];
    }
}

function saveUnidadesArchivadas(archivadas) {
    try {
        fs.writeFileSync(ARCHIVADAS_FILE, JSON.stringify(archivadas, null, 2), 'utf8');
        return true;
    } catch(e) {
        return false;
    }
}

function getRemolques() {
    if (!fs.existsSync(REMOLQUES_FILE)) {
        return [];
    }
    try {
        return JSON.parse(fs.readFileSync(REMOLQUES_FILE, 'utf8'));
    } catch(e) {
        return [];
    }
}

function saveRemolques(remolques) {
    try {
        fs.writeFileSync(REMOLQUES_FILE, JSON.stringify(remolques, null, 2), 'utf8');
        broadcastSSE('remolques_updated', { remolques, count: remolques.length });
        return true;
    } catch(e) {
        return false;
    }
}

function getTitulosColumnas() {
    const defaultTitulos = {
        considerar: "A Considerar / Detecciones",
        mayor: "Reparación Mayor / Presupuesto",
        proceso: "En Reparación Activa",
        listo: "Listas para Entrega"
    };
    if (!fs.existsSync(TITULOS_COLUMNAS_FILE)) return defaultTitulos;
    try {
        return { ...defaultTitulos, ...JSON.parse(fs.readFileSync(TITULOS_COLUMNAS_FILE, 'utf8')) };
    } catch(e) {
        return defaultTitulos;
    }
}

function saveTitulosColumnas(titulos) {
    try {
        fs.writeFileSync(TITULOS_COLUMNAS_FILE, JSON.stringify(titulos, null, 2), 'utf8');
        return true;
    } catch(e) {
        return false;
    }
}

function getPreventivosData() {
    const freshCachePath = path.join(__dirname, 'combustible_preventivos_cache.json');
    if (fs.existsSync(freshCachePath)) {
        try {
            return JSON.parse(fs.readFileSync(freshCachePath, 'utf8'));
        } catch(e) {}
    }
    const webCachePath = path.join(__dirname, 'web', 'src', 'data', 'initial_cache.json');
    if (fs.existsSync(webCachePath)) {
        try {
            const raw = JSON.parse(fs.readFileSync(webCachePath, 'utf8'));
            return {
                preventivos: raw.preventivos || [],
                bitacora: raw.preventivos_historial || [],
                stats: { total: (raw.preventivos || []).length }
            };
        } catch(e) {}
    }
    const dieselCachePath = path.join(__dirname, 'diesel_local_cache.json');
    if (fs.existsSync(dieselCachePath)) {
        try {
            const cache = JSON.parse(fs.readFileSync(dieselCachePath, 'utf8'));
            if (cache && Array.isArray(cache.preventivos)) {
                return {
                    preventivos: cache.preventivos,
                    bitacora: cache.preventivos_historial || [],
                    stats: { total: cache.preventivos.length }
                };
            }
        } catch(e) {}
    }
    return { preventivos: [], bitacora: [], stats: { total: 0 } };
}

function getPreventivos() {
    const data = getPreventivosData();
    return data.preventivos || [];
}

function getBitacoraPreventivos() {
    const data = getPreventivosData();
    return data.bitacora || [];
}

function getRescates() {
    if (!fs.existsSync(RESCATES_FILE)) return [];
    try {
        return JSON.parse(fs.readFileSync(RESCATES_FILE, 'utf8'));
    } catch (e) {
        console.error('Error leyendo taller_rescates.json:', e);
        return [];
    }
}

function saveRescates(rescates) {
    try {
        fs.writeFileSync(RESCATES_FILE, JSON.stringify(rescates, null, 2), 'utf8');
        if (typeof broadcastSSE === 'function') {
            broadcastSSE('rescates_updated', { rescates, count: rescates.length });
        }
        return true;
    } catch (e) {
        console.error('Error guardando taller_rescates.json:', e);
        return false;
    }
}

// --- Helper Funciones de Base de Datos y Fechas ---
function getFechaActualISO() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function formatearFechaTitulo(fechaISO) {
    try {
        const [y, m, d] = fechaISO.split('-').map(Number);
        const date = new Date(y, m - 1, d);
        return date.toLocaleDateString('es-MX', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
    } catch (e) {
        return fechaISO;
    }
}

// Normalización de Refacciones y Materiales
function normalizarPieza(p, idx = 0) {
    if (!p) return null;
    if (typeof p === 'string') {
        return {
            id: `pz_leg_${idx}_${Date.now()}`,
            id_pieza: null,
            descripcion: p,
            numero_parte: '',
            marca: '',
            cantidad: 1,
            unidad_medida: 'PZ',
            precio_unitario: 0,
            costo_total: 0,
            rack: '',
            nivel: '',
            stock_disponible: 0,
            estado: 'SOLICITADA_ALMACEN', // SOLICITADA_ALMACEN, SALIDA_AUTORIZADA, EN_ORDEN_COMPRA, RECIBIDA_TALLER, INSTALADA
            fecha_solicitud: new Date().toISOString(),
            folio_oc: '',
            proveedor_oc: '',
            fecha_estimada_arribo: '',
            autorizado_por: '',
            fecha_autorizacion: '',
            recibido_por: '',
            fecha_recepcion: '',
            instalado_por: '',
            fecha_instalacion: '',
            notas: ''
        };
    }
    const cant = Number(p.cantidad) > 0 ? Number(p.cantidad) : 1;
    const pu = Number(p.precio_unitario) || 0;
    return {
        id: p.id || `pz_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
        id_pieza: p.id_pieza || null,
        descripcion: p.descripcion || 'Refacción sin descripción',
        numero_parte: p.numero_parte || '',
        marca: p.marca || '',
        cantidad: cant,
        unidad_medida: p.unidad_medida || p.unidad || 'PZ',
        precio_unitario: pu,
        costo_total: p.costo_total !== undefined ? Number(p.costo_total) : (pu * cant),
        rack: p.rack || '',
        nivel: p.nivel || '',
        stock_disponible: p.stock_disponible !== undefined ? Number(p.stock_disponible) : (p.stock_actual !== undefined ? Number(p.stock_actual) : 0),
        estado: p.estado || 'SOLICITADA_ALMACEN',
        fecha_solicitud: p.fecha_solicitud || new Date().toISOString(),
        folio_oc: p.folio_oc || '',
        proveedor_oc: p.proveedor_oc || '',
        fecha_estimada_arribo: p.fecha_estimada_arribo || '',
        autorizado_por: p.autorizado_por || '',
        fecha_autorizacion: p.fecha_autorizacion || '',
        recibido_por: p.recibido_por || '',
        fecha_recepcion: p.fecha_recepcion || '',
        instalado_por: p.instalado_por || '',
        fecha_instalacion: p.fecha_instalacion || '',
        notas: p.notas || ''
    };
}

function getAlmacenData() {
    if (!fs.existsSync(ALMACEN_DB_FILE)) return { catalog: {}, movimientos: [], ordenes_compra: [] };
    try {
        return JSON.parse(fs.readFileSync(ALMACEN_DB_FILE, 'utf8'));
    } catch (e) {
        console.error('Error leyendo almacen_datos.json:', e);
        return { catalog: {}, movimientos: [], ordenes_compra: [] };
    }
}

function saveAlmacenData(data) {
    try {
        fs.writeFileSync(ALMACEN_DB_FILE, JSON.stringify(data, null, 2), 'utf8');
        return true;
    } catch (e) {
        console.error('Error guardando almacen_datos.json:', e);
        return false;
    }
}

// =========================================================================
// SINCRONIZADOR BIDIRECCIONAL EN TIEMPO REAL: TALLER PRO <-> ALMACÉN CESS
// =========================================================================
function sincronizarFoliosConAlmacen(origen = 'Watcher Almacén') {
    const almacenData = getAlmacenData();
    const movimientos = almacenData.movimientos || [];
    const catalog = almacenData.catalog || {};
    let unidades = getUnidades();
    let huboCambios = false;

    unidades.forEach(unidad => {
        if (!unidad.folio_almacen) return;
        const folioBuscado = String(unidad.folio_almacen).trim();

        // Movimientos de salida vinculados a este folio o unidad
        const movsFolio = movimientos.filter(m => 
            String(m.folio || '').trim() === folioBuscado && (m.tipo === 'S' || !m.tipo)
        );

        let nuevoCostoRefacciones = 0;
        const piezasActualizadas = movsFolio.map((m, idx) => {
            const itemCat = catalog[m.id_pieza] || {};
            const pu = Number(m.precio_unitario !== undefined ? m.precio_unitario : (itemCat.precio_unitario || 0));
            const cant = Number(m.cantidad) || 1;
            const total = pu * cant;
            nuevoCostoRefacciones += total;
            return {
                id: m.id || `pz_alm_${Date.now()}_${idx}`,
                id_pieza: m.id_pieza,
                numero_parte: m.numero_parte || itemCat.numero_parte || '',
                descripcion: m.descripcion || itemCat.descripcion || 'Refacción',
                marca: m.marca || itemCat.marca || '',
                cantidad: cant,
                unidad_medida: m.unidad || itemCat.unidad || 'PZ',
                precio_unitario: pu,
                costo_total: total,
                folio_almacen: folioBuscado,
                solicitante: m.solicitante || 'Almacén CESS',
                estado: 'RECIBIDA_TALLER',
                fecha_recepcion: m.fecha || new Date().toISOString().slice(0, 10),
                hora_recepcion: m.hora || ''
            };
        });

        const costoPrevio = Number(unidad.costo_refacciones) || 0;
        const cantPrevias = Array.isArray(unidad.piezas) ? unidad.piezas.filter(p => p.folio_almacen === folioBuscado).length : 0;

        // Comprobar si hubo cambios en cantidad, piezas o costo
        if (Math.abs(costoPrevio - nuevoCostoRefacciones) > 0.01 || cantPrevias !== piezasActualizadas.length) {
            huboCambios = true;
            unidad.costo_refacciones = nuevoCostoRefacciones;
            
            // Conservar piezas de otros folios o manuales
            const piezasOtras = (unidad.piezas || []).filter(p => p.folio_almacen !== folioBuscado);
            unidad.piezas = [...piezasOtras, ...piezasActualizadas];
            
            // Recalcular costo total de la unidad
            unidad.costo = (Number(unidad.costo_mano_obra) || 0) + nuevoCostoRefacciones;
            
            if (!Array.isArray(unidad.historial)) unidad.historial = [];
            unidad.historial.unshift({
                fecha: new Date().toLocaleDateString('es-MX') + ' ' + new Date().toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' }),
                evento: `Almacén sincronizó Folio #${folioBuscado}: ${piezasActualizadas.length} piezas surtidas (${nuevoCostoRefacciones.toLocaleString()} MXN)`,
                tipo: 'almacen_sync',
                autor: origen
            });

            console.log(`>>> [Sincronización Bidireccional] Unidad ${unidad.id} vinculada a Folio #${folioBuscado} actualizada: ${piezasActualizadas.length} refacciones (${nuevoCostoRefacciones} MXN)`);
        }
    });

    if (huboCambios) {
        saveUnidades(unidades, origen);
        if (typeof broadcastSSE === 'function') {
            broadcastSSE('almacen_folio_updated', {
                mensaje: 'Refacciones y costos actualizados en vivo desde Almacén CESS',
                timestamp: Date.now()
            });
        }
    }
}

// Observador continuo de cambios físicos en almacen_datos.json
try {
    if (fs.existsSync(ALMACEN_DB_FILE)) {
        fs.watchFile(ALMACEN_DB_FILE, { interval: 1000 }, (curr, prev) => {
            if (curr.mtimeMs !== prev.mtimeMs) {
                console.log('[Almacén Watcher] Cambio detectado en base de datos de almacén. Sincronizando con Taller Pro...');
                sincronizarFoliosConAlmacen('Almacén CESS en tiempo real');
            }
        });
        console.log('[Almacén Watcher] Monitoreo bidireccional activo sobre:', ALMACEN_DB_FILE);
    }
} catch(e) {
    console.warn('[Almacén Watcher] No se pudo inicializar watchFile:', e.message);
}

function getUnidades() {
    if (!fs.existsSync(DB_FILE)) {
        if (fs.existsSync(path.join(__dirname, 'taller_datos_iniciales.json'))) {
            fs.copyFileSync(path.join(__dirname, 'taller_datos_iniciales.json'), DB_FILE);
        } else {
            return [];
        }
    }
    try {
        const unidades = JSON.parse(fs.readFileSync(DB_FILE, 'utf8'));
        if (Array.isArray(unidades)) {
            unidades.forEach(u => {
                if (Array.isArray(u.piezas) && u.piezas.length > 0) {
                    u.piezas = u.piezas.map((p, idx) => normalizarPieza(p, idx)).filter(Boolean);
                } else if (!Array.isArray(u.piezas)) {
                    u.piezas = [];
                }
            });
        }
        return unidades;
    } catch (e) {
        console.error('Error leyendo taller_db.json:', e);
        return [];
    }
}

// --- MOTOR SERVER-SENT EVENTS (SSE) EN TIEMPO REAL ---
const sseClients = new Set();
let ultimaMarcaTiempoDatos = Date.now();

function broadcastSSE(type, payload = {}) {
    ultimaMarcaTiempoDatos = Date.now();
    const dataStr = `data: ${JSON.stringify({ type, timestamp: ultimaMarcaTiempoDatos, ...payload })}\n\n`;
    for (const client of sseClients) {
        try {
            client.write(dataStr);
        } catch (e) {
            sseClients.delete(client);
        }
    }
}

function saveUnidades(unidades, autor = 'Sistema') {
    if (Array.isArray(unidades)) {
        // Garantizar unicidad estricta por ID (evitar duplicados entre columnas)
        const mapa = new Map();
        unidades.forEach(u => {
            if (u && u.id) {
                const key = String(u.id).trim();
                if (!mapa.has(key)) {
                    mapa.set(key, u);
                } else {
                    const exist = mapa.get(key);
                    // Si una unidad duplicada tiene un estado más avanzado, adoptarlo
                    if (u.estado === 'listo' || (u.estado === 'proceso' && exist.estado === 'mayor')) {
                        exist.estado = u.estado;
                    }
                    if (u.reporte && !exist.reporte) exist.reporte = u.reporte;
                    if (Array.isArray(u.piezas) && u.piezas.length > 0) {
                        exist.piezas = [...(exist.piezas || []), ...u.piezas];
                    }
                    if (Array.isArray(u.historial) && u.historial.length > 0) {
                        exist.historial = [...(exist.historial || []), ...u.historial];
                    }
                }
            }
        });
        unidades = Array.from(mapa.values());

        unidades.forEach(u => {
            if (Array.isArray(u.piezas) && u.piezas.length > 0) {
                u.piezas = u.piezas.map((p, idx) => normalizarPieza(p, idx)).filter(Boolean);
            }
        });
    }

    // PROTECCIÓN DE DATOS HISTÓRICOS: Si una unidad ya no está en el tablero activo, archivarla sin perder avances
    if (fs.existsSync(DB_FILE)) {
        try {
            const previas = JSON.parse(fs.readFileSync(DB_FILE, 'utf8'));
            if (Array.isArray(previas)) {
                let archivadas = getUnidadesArchivadas();
                let cambioArchivadas = false;
                previas.forEach(p => {
                    const sigueEnTablero = unidades.some(u => String(u.id).trim() === String(p.id).trim());
                    if (!sigueEnTablero) {
                        const idxArch = archivadas.findIndex(a => String(a.id).trim() === String(p.id).trim());
                        const copia = { ...p, archivada: true, fecha_salida: p.fecha_salida || new Date().toLocaleDateString('es-MX'), motivo_salida: p.motivo_salida || 'Retirada del tablero activo / Salida a viaje' };
                        if (idxArch >= 0) {
                            archivadas[idxArch] = { ...archivadas[idxArch], ...copia };
                        } else {
                            archivadas.unshift(copia);
                        }
                        cambioArchivadas = true;
                    }
                });
                if (cambioArchivadas) {
                    saveUnidadesArchivadas(archivadas);
                }
            }
        } catch(e) {}
    }

    fs.writeFileSync(DB_FILE, JSON.stringify(unidades, null, 2), 'utf8');
    broadcastSSE('units_updated', { units: unidades, autor, count: unidades.length });
}

function getBitacoras() {
    if (!fs.existsSync(BITACORAS_FILE)) {
        return [];
    }
    try {
        return JSON.parse(fs.readFileSync(BITACORAS_FILE, 'utf8'));
    } catch (e) {
        console.error('Error leyendo taller_bitacoras_historico.json:', e);
        return [];
    }
}

function saveBitacoras(bitacoras) {
    fs.writeFileSync(BITACORAS_FILE, JSON.stringify(bitacoras, null, 2), 'utf8');
}

// --- GESTIÓN DE JORNADA DIARIA & ROLLOVER DE MEDIANOCHE ---
function getJornadaActual() {
    const fechaHoy = getFechaActualISO();
    if (!fs.existsSync(JORNADA_FILE)) {
        const nueva = {
            fecha: fechaHoy,
            hora_inicio: "00:00:00",
            titulo: `Jornada Activa - ${formatearFechaTitulo(fechaHoy)}`,
            estado: 'ACTIVA',
            jornada_iniciada_con: 'Cierre del día anterior',
            proximo_corte_automatico: 'Medianoche (00:00:00)',
            movimientos: [
                {
                    id: `mov_init_${Date.now()}`,
                    hora: "00:00:00",
                    unidad_id: "TODAS",
                    tipo: "apertura",
                    descripcion: `Apertura de jornada diaria. Se arranca con el estado del cierre anterior.`,
                    autor: "Sistema Taller Pro"
                }
            ]
        };
        fs.writeFileSync(JORNADA_FILE, JSON.stringify(nueva, null, 2), 'utf8');
        return nueva;
    }

    try {
        let jornada = JSON.parse(fs.readFileSync(JORNADA_FILE, 'utf8'));
        if (jornada.fecha && jornada.fecha !== fechaHoy) {
            console.log(`[Medianoche] Rollover automático de ${jornada.fecha} hacia ${fechaHoy}`);
            jornada = ejecutarRolloverMedianoche(jornada.fecha, fechaHoy);
        }
        return jornada;
    } catch (e) {
        console.error('Error leyendo jornada:', e);
        return { fecha: fechaHoy, hora_inicio: "00:00:00", movimientos: [] };
    }
}

function saveJornada(jornada) {
    fs.writeFileSync(JORNADA_FILE, JSON.stringify(jornada, null, 2), 'utf8');
}

function ejecutarRolloverMedianoche(fechaAnterior, fechaNueva) {
    const unidades = getUnidades();
    const bitacoras = getBitacoras();

    let oldMovimientos = [];
    let notasCierrePrevias = 'Corte automático de medianoche (12:00 AM). Las unidades continúan su ciclo de trabajo hacia el nuevo día.';
    if (fs.existsSync(JORNADA_FILE)) {
        try {
            const oldJ = JSON.parse(fs.readFileSync(JORNADA_FILE, 'utf8'));
            if (oldJ.movimientos) oldMovimientos = oldJ.movimientos;
            if (oldJ.notas_cierre) notasCierrePrevias = oldJ.notas_cierre;
        } catch (e) {}
    }

    const yaExiste = bitacoras.some(b => b.fecha === fechaAnterior);
    if (!yaExiste) {
        const counts = { total: unidades.length, mayor: 0, proceso: 0, listo: 0, costo_acumulado: 0 };
        unidades.forEach(u => {
            if (counts[u.estado] !== undefined) counts[u.estado]++;
            counts.costo_acumulado += (Number(u.costo) || 0);
        });

        const snapshot = {
            id: `snap_${fechaAnterior.replace(/-/g, '')}_2359`,
            fecha: fechaAnterior,
            hora: '23:59:59',
            titulo: `Cierre Oficial de Bitácora - ${formatearFechaTitulo(fechaAnterior)}`,
            responsable: 'Cierre Automático de Medianoche (00:00 hrs)',
            resumen: counts,
            notas_cierre: notasCierrePrevias,
            movimientos: oldMovimientos,
            unidades: unidades.map(u => ({
                id: u.id,
                sucursal: u.sucursal,
                estado: u.estado,
                defecto: u.defecto,
                notas: u.notas,
                reporte: u.reporte,
                costo: u.costo,
                fechaRefaccion: u.fechaRefaccion
            }))
        };
        bitacoras.unshift(snapshot);
        saveBitacoras(bitacoras);
        console.log(`[Rollover Medianoche] Corte guardado para ${fechaAnterior}`);
    }

    const nuevaJornada = {
        fecha: fechaNueva,
        hora_inicio: '00:00:00',
        titulo: `Jornada Activa - ${formatearFechaTitulo(fechaNueva)}`,
        estado: 'ACTIVA',
        jornada_iniciada_con: `Cierre del día anterior (${fechaAnterior})`,
        proximo_corte_automatico: 'Medianoche (00:00:00)',
        movimientos: [
            {
                id: `mov_init_${Date.now()}`,
                hora: '00:00:00',
                unidad_id: 'TODAS',
                tipo: 'apertura',
                descripcion: `Apertura de jornada diaria. Se arranca con el estatus del cierre de ayer (${fechaAnterior}).`,
                autor: 'Sistema Taller Pro'
            }
        ]
    };
    saveJornada(nuevaJornada);
    return nuevaJornada;
}

function registrarMovimientoJornada(unidad_id, tipo, descripcion, autor = 'Jefe de Taller') {
    const jornada = getJornadaActual();
    const now = new Date();
    const hora = now.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    const nuevoMov = {
        id: `mov_${Date.now()}_${Math.floor(Math.random() * 1000)}`,
        fecha: jornada.fecha,
        hora,
        unidad_id: String(unidad_id),
        tipo,
        descripcion,
        autor
    };
    if (!jornada.movimientos) jornada.movimientos = [];
    jornada.movimientos.unshift(nuevoMov);
    saveJornada(jornada);
    broadcastSSE('jornada_updated', { jornada, nuevoMovimiento: nuevoMov, autor, unidad_id });
    return nuevoMov;
}

// --- 0. STREAM SERVER-SENT EVENTS (SSE) EN TIEMPO REAL ---
app.get('/api/stream', (req, res) => {
    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache, no-transform');
    res.setHeader('Connection', 'keep-alive');
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('X-Accel-Buffering', 'no');
    if (typeof res.flushHeaders === 'function') res.flushHeaders();

    const unidades = getUnidades();
    const jornada = getJornadaActual();
    res.write(`data: ${JSON.stringify({ type: 'connected', units: unidades, jornada, timestamp: ultimaMarcaTiempoDatos })}\n\n`);

    sseClients.add(res);

    req.on('close', () => {
        sseClients.delete(res);
    });
});

// Ping de mantenimiento cada 15 segundos para túneles y routers móviles
setInterval(() => {
    for (const client of sseClients) {
        try {
            client.write(': ping\n\n');
        } catch (e) {
            sseClients.delete(client);
        }
    }
}, 15000);

// Timer en segundo plano: verificar cada 30 segundos si cruzamos la medianoche
setInterval(() => {
    try {
        const fechaHoy = getFechaActualISO();
        if (fs.existsSync(JORNADA_FILE)) {
            const data = JSON.parse(fs.readFileSync(JORNADA_FILE, 'utf8'));
            if (data.fecha && data.fecha !== fechaHoy) {
                console.log(`[Medianoche] Detectado cambio de día de ${data.fecha} a ${fechaHoy}. Ejecutando cierre...`);
                ejecutarRolloverMedianoche(data.fecha, fechaHoy);
            }
        } else {
            getJornadaActual();
        }
    } catch (err) {
        console.error('Error en monitor de medianoche:', err);
    }
}, 30000);

// Configuración de almacenamiento para fotos de evidencia
const storage = multer.diskStorage({
    destination: (req, file, cb) => cb(null, UPLOADS_DIR),
    filename: (req, file, cb) => {
        const ext = path.extname(file.originalname) || '.jpg';
        cb(null, `evidencia_${Date.now()}_${Math.floor(Math.random() * 10000)}${ext}`);
    }
});
const upload = multer({ storage, limits: { fileSize: 30 * 1024 * 1024 } });

// --- ENDPOINTS API ---

// 1. Obtener jornada activa y cuenta regresiva a medianoche
app.get('/api/jornada', (req, res) => {
    const jornada = getJornadaActual();
    const unidades = getUnidades();
    const counts = { total: unidades.length, mayor: 0, proceso: 0, listo: 0, costo_acumulado: 0 };
    unidades.forEach(u => {
        if (counts[u.estado] !== undefined) counts[u.estado]++;
        counts.costo_acumulado += (Number(u.costo) || 0);
    });

    const now = new Date();
    const medianoche = new Date(now);
    medianoche.setHours(24, 0, 0, 0);
    const msRestantes = Math.max(0, medianoche - now);
    const horasRestantes = Math.floor(msRestantes / (1000 * 60 * 60));
    const minsRestantes = Math.floor((msRestantes % (1000 * 60 * 60)) / (1000 * 60));

    res.json({
        success: true,
        jornada,
        conteo_actual: counts,
        corte_automatico_en: `${horasRestantes}h ${minsRestantes}m (12:00 AM Medianoche)`
    });
});

// 2. Registrar movimiento en la jornada de hoy
app.post('/api/jornada/movimiento', (req, res) => {
    const { unidad_id, tipo, descripcion, autor } = req.body;
    if (!descripcion) return res.status(400).json({ error: 'Falta descripción' });
    const mov = registrarMovimientoJornada(unidad_id || 'GENERAL', tipo || 'nota', descripcion, autor);
    res.json({ success: true, movimiento: mov });
});

// 3. Obtener todas las unidades
app.get('/api/unidades', (req, res) => {
    getJornadaActual(); // valida medianoche si aplica
    const unidades = getUnidades();
    res.json({ success: true, count: unidades.length, data: unidades });
});

// 4. Guardar / sincronizar todas las unidades (Full Sync)
app.post('/api/unidades', (req, res) => {
    const unidades = req.body;
    if (!Array.isArray(unidades)) {
        return res.status(400).json({ error: 'Se esperaba un arreglo de unidades' });
    }
    saveUnidades(unidades);
    res.json({ success: true, count: unidades.length });
});

// 5. Actualizar una unidad específica
app.post('/api/unidad/:id', (req, res) => {
    const unitId = String(req.params.id).trim();
    const updates = req.body;
    let unidades = getUnidades();
    let index = unidades.findIndex(u => String(u.id) === unitId);

    if (index === -1) {
        const nueva = {
            id: unitId,
            sucursal: updates.sucursal || 'Apodaca',
            estado: updates.estado || 'proceso',
            defecto: updates.defecto || updates.reporte || 'Sin especificar',
            notas: updates.notas || '',
            reporte: updates.reporte || updates.defecto || 'Sin especificar',
            costo: updates.costo || 0,
            fechaRefaccion: updates.fechaRefaccion || '',
            categoriaFalla: updates.categoriaFalla || 'General',
            piezas: updates.piezas || [],
            historial: updates.historial || [{
                fecha: new Date().toLocaleString('es-MX'),
                evento: `Unidad registrada (${updates.estado || 'proceso'})`,
                tipo: 'alta',
                autor: 'Móvil'
            }]
        };
        unidades.unshift(nueva);
        saveUnidades(unidades);
        registrarMovimientoJornada(unitId, 'alta', `Unidad ${unitId} registrada en taller (${nueva.estado.toUpperCase()})`);
        return res.json({ success: true, action: 'created', data: nueva });
    }

    const estadoPrevio = unidades[index].estado;
    unidades[index] = { ...unidades[index], ...updates };
    saveUnidades(unidades);

    // Si cambió de estado, registrar en movimientos de hoy
    if (updates.estado && updates.estado !== estadoPrevio) {
        registrarMovimientoJornada(unitId, 'cambio_estatus', `Unidad ${unitId} cambió de '${estadoPrevio}' a '${updates.estado}'`);
    } else {
        registrarMovimientoJornada(unitId, 'actualizacion', `Unidad ${unitId}: ${updates.reporte || 'Avance actualizado'}`);
    }

    res.json({ success: true, action: 'updated', data: unidades[index] });
});

// 6. HISTORIAL CLÍNICO DÍA POR DÍA DE UNA UNIDAD
app.get('/api/unidad/:id/historial-diario', (req, res) => {
    const unitId = String(req.params.id).trim();
    const unidades = getUnidades();
    const archivadas = getUnidadesArchivadas();
    let u = unidades.find(x => String(x.id) === unitId) || archivadas.find(x => String(x.id) === unitId);

    // Búsqueda profunda en bitácoras históricas si la unidad fue atendida en jornadas anteriores
    if (!u) {
        const bitacoras = getBitacoras();
        for (const b of bitacoras) {
            const uSnap = (b.unidades || []).find(x => String(x.id) === unitId);
            if (uSnap) {
                u = { ...uSnap, historial: uSnap.historial || [], avances: uSnap.avances || [] };
                break;
            }
        }
    }

    if (!u) {
        return res.status(404).json({ error: `Unidad ${unitId} no encontrada` });
    }

    const { desde, hasta } = req.query;
    let historial = u.historial ? [...u.historial] : [];

    const mapaDias = new Map();

    historial.forEach(ev => {
        let fechaISO = '';
        let fechaTexto = 'Sin Fecha';

        const matchFecha = String(ev.fecha || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
        if (matchFecha) {
            const d = matchFecha[1].padStart(2, '0');
            const m = matchFecha[2].padStart(2, '0');
            let y = matchFecha[3];
            if (y.length === 2) y = '20' + y;
            fechaISO = `${y}-${m}-${d}`;
            fechaTexto = `${d}/${m}/${y}`;
        } else {
            fechaISO = getFechaActualISO();
            fechaTexto = fechaISO;
        }

        if (desde && fechaISO < desde) return;
        if (hasta && fechaISO > hasta) return;

        if (!mapaDias.has(fechaISO)) {
            mapaDias.set(fechaISO, {
                fechaISO,
                fechaTexto,
                fechaTitulo: formatearFechaTitulo(fechaISO),
                eventos: []
            });
        }
        mapaDias.get(fechaISO).eventos.push(ev);
    });

    const diasOrdenados = Array.from(mapaDias.values()).sort((a, b) => b.fechaISO.localeCompare(a.fechaISO));

    res.json({
        success: true,
        unidad: {
            id: u.id,
            sucursal: u.sucursal,
            estado: u.estado,
            defecto: u.defecto,
            notas: u.notas,
            reporte: u.reporte,
            costo: u.costo,
            fechaRefaccion: u.fechaRefaccion,
            categoriaFalla: u.categoriaFalla,
            piezas: u.piezas || [],
            avances: u.avances || [],
            archivada: u.archivada || false,
            fecha_salida: u.fecha_salida || null,
            motivo_salida: u.motivo_salida || null
        },
        resumen: {
            dias_con_actividad: diasOrdenados.length,
            costo_acumulado: Number(u.costo) || 0,
            total_eventos: historial.length
        },
        dias: diasOrdenados
    });
});


// --- 6.5 GESTIÓN DE SALIDA A VIAJE Y UNIDADES ARCHIVADAS ---
app.get('/api/unidades/archivadas', (req, res) => {
    const archivadas = getUnidadesArchivadas();
    res.json({ success: true, count: archivadas.length, archivadas });
});

app.post('/api/unidad/:id/liberar', (req, res) => {
    const unitId = String(req.params.id).trim();
    const { motivo, notas, km_salida } = req.body || {};
    let unidades = getUnidades();
    const idx = unidades.findIndex(u => String(u.id).trim() === unitId);

    if (idx === -1) {
        return res.status(404).json({ error: `Unidad ${unitId} no encontrada en el tablero activo` });
    }

    const unidad = unidades[idx];
    const fechaSalida = new Date().toLocaleDateString('es-MX');
    const horaSalida = new Date().toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' });

    const eventoSalida = {
        fecha: fechaSalida,
        hora: horaSalida,
        evento: `Salida de taller: ${motivo || 'Asignada a viaje / Liberada'}. ${notas || ''}`.trim(),
        tipo: 'liberacion',
        km: km_salida || null
    };

    if (!Array.isArray(unidad.historial)) unidad.historial = [];
    unidad.historial.push(eventoSalida);

    if (!Array.isArray(unidad.avances)) unidad.avances = [];
    unidad.avances.unshift({
        fecha: fechaSalida,
        hora: horaSalida,
        texto: `🏁 Unidad liberada de taller: ${motivo || 'Asignada a viaje'}. ${notas || ''}`.trim(),
        mecanico: 'Jefe de Taller',
        porcentaje: 100
    });

    unidad.archivada = true;
    unidad.fecha_salida = fechaSalida;
    unidad.hora_salida = horaSalida;
    unidad.motivo_salida = motivo || 'Asignada a viaje / Liberada';
    unidad.notas_salida = notas || '';
    if (km_salida) unidad.km_salida = km_salida;

    let archivadas = getUnidadesArchivadas();
    const archIdx = archivadas.findIndex(a => String(a.id).trim() === unitId);
    if (archIdx >= 0) {
        archivadas[archIdx] = { ...archivadas[archIdx], ...unidad };
    } else {
        archivadas.unshift(unidad);
    }
    saveUnidadesArchivadas(archivadas);

    // Retirar del tablero activo
    unidades.splice(idx, 1);
    saveUnidades(unidades);

    // Registrar en movimientos de jornada
    try {
        const jornada = getJornadaActual();
        if (!jornada.movimientos) jornada.movimientos = [];
        jornada.movimientos.unshift({
            id: `mov_lib_${Date.now()}`,
            hora: horaSalida,
            unidad_id: unitId,
            tipo: 'liberacion',
            descripcion: `Unidad ${unitId} liberada de taller: ${motivo || 'Salida a viaje'}`,
            autor: 'Jefe de Taller'
        });
        saveJornada(jornada);
    } catch(e) {}

    res.json({ success: true, mensaje: `Unidad ${unitId} liberada y archivada correctamente`, unidad });
});

app.post('/api/unidad/:id/reincorporar', (req, res) => {
    const unitId = String(req.params.id).trim();
    let archivadas = getUnidadesArchivadas();
    const idx = archivadas.findIndex(a => String(a.id).trim() === unitId);

    if (idx === -1) {
        return res.status(404).json({ error: `Unidad ${unitId} no encontrada en archivo` });
    }

    const unidad = archivadas[idx];
    unidad.archivada = false;
    unidad.estado = unidad.estado || 'proceso';

    const fechaReinc = new Date().toLocaleDateString('es-MX');
    const horaReinc = new Date().toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' });
    if (!Array.isArray(unidad.historial)) unidad.historial = [];
    unidad.historial.push({
        fecha: fechaReinc,
        hora: horaReinc,
        evento: 'Reincorporada al tablero de taller activo',
        tipo: 'reincorporacion'
    });

    let unidades = getUnidades();
    const existIdx = unidades.findIndex(u => String(u.id).trim() === unitId);
    if (existIdx >= 0) {
        unidades[existIdx] = unidad;
    } else {
        unidades.unshift(unidad);
    }
    saveUnidades(unidades);

    archivadas.splice(idx, 1);
    saveUnidadesArchivadas(archivadas);

    res.json({ success: true, mensaje: `Unidad ${unitId} reincorporada al tablero`, unidad });
});

// --- 23.5 ACTUALIZACIÓN BIDIRECCIONAL DE SERVICIOS PREVENTIVOS ---
app.post('/api/preventivo/:unidad/actualizar-servicio', (req, res) => {
    const unidadId = String(req.params.unidad).trim();
    const { km_actual, fecha_realizado, intervalo, tipo_tracto, costo, observaciones } = req.body || {};

    const targetFiles = [
        'combustible_preventivos_cache.json',
        'taller_preventivos.json',
        'diesel_local_cache.json',
        'web/src/data/initial_cache.json'
    ];

    let unidadActualizada = null;

    targetFiles.forEach(relPath => {
        const fullPath = path.join(__dirname, relPath);
        if (!fs.existsSync(fullPath)) return;
        try {
            let content = JSON.parse(fs.readFileSync(fullPath, 'utf8'));
            let list = Array.isArray(content) ? content : (content.preventivos || []);
            let item = list.find(x => String(x.unidad).trim() === unidadId);

            const tracto = tipo_tracto || (item ? item.tipo_tracto : 'Cascadia');
            const isSitrak = String(tracto).toLowerCase().includes('sitrak');
            const inter = Number(intervalo) || (isSitrak ? 30000 : 35000);
            const km = Number(km_actual || (item ? item.km_actual : 0));
            const prox = km + inter;
            const fecha = fecha_realizado || new Date().toLocaleDateString('es-MX');

            const updatedData = {
                unidad: unidadId,
                division: item ? item.division : 'RL',
                marca: item ? item.marca : (isSitrak ? 'Sinotruk' : 'Freightliner'),
                tipo_tracto: tracto,
                km_actual: km,
                km_mtto_anterior: km,
                fecha_realizado: fecha,
                proximo_mtto_km: prox,
                km_restantes: inter,
                estatus: 'VIGENTE',
                estatusLabel: 'Vigente',
                ubicacion: item ? item.ubicacion : 'PATIO APODACA'
            };

            unidadActualizada = updatedData;

            if (item) {
                Object.assign(item, updatedData);
            } else {
                list.push(updatedData);
            }

            const histEntry = {
                id: `bit_${Date.now()}`,
                unidad: unidadId,
                fecha: fecha,
                km_actual: km,
                intervalo: inter,
                tipo_tracto: tracto,
                marca: updatedData.marca,
                costo_total: Number(costo) || 0,
                observaciones: observaciones || 'Servicio preventivo registrado en sistema'
            };

            if (content.preventivos_historial && Array.isArray(content.preventivos_historial)) {
                content.preventivos_historial.unshift(histEntry);
            }
            if (content.bitacora && Array.isArray(content.bitacora)) {
                content.bitacora.unshift(histEntry);
            }

            if (!Array.isArray(content)) {
                content.stats = {
                    total: list.length,
                    vigentes: list.filter(p => (p.estatus || '').toUpperCase() === 'VIGENTE').length,
                    vencidos: list.filter(p => (p.estatus || '').toUpperCase() === 'VENCIDO').length,
                    proximos: list.filter(p => (p.estatus || '').toUpperCase() === 'PRÓXIMO' || (p.estatus || '').toUpperCase() === 'PROXIMO').length
                };
                content.resumen = content.stats;
            }

            fs.writeFileSync(fullPath, JSON.stringify(content, null, 2), 'utf8');
        } catch (e) {
            console.error(`Error actualizando preventivo en ${relPath}:`, e.message);
        }
    });

    broadcastSSE('preventivos_updated', { unidad: unidadActualizada });
    res.json({ success: true, mensaje: `Servicio preventivo de Unidad ${unidadId} actualizado con éxito`, preventivo: unidadActualizada });
});

// 7. Sincronizar desde archivo Excel Taller Flotilla-v2.xlsx
app.post('/api/sincronizar-excel', (req, res) => {
    if (!fs.existsSync(EXCEL_FILE)) {
        return res.status(404).json({ error: 'No se encontró el archivo Taller Flotilla-v2.xlsx' });
    }

    try {
        const wb = xlsx.readFile(EXCEL_FILE);
        const ws = wb.Sheets['Requisiciones'] || wb.Sheets[wb.SheetNames[0]];
        const raw = xlsx.utils.sheet_to_json(ws);
        const existentes = getUnidades();
        const mapExistentes = new Map(existentes.map(u => [String(u.id), u]));

        const actualizadas = raw.map(r => {
            const id = String(r['Número de Unidad'] || '').trim();
            const sucursal = String(r['Terminal'] || 'Apodaca').trim();
            const estatusRaw = String(r['Estatus Actual'] || '').toLowerCase();
            
            let estado = 'proceso';
            if (estatusRaw.includes('$$') || estatusRaw.includes('presupuesto') || estatusRaw.includes('pendiente')) {
                estado = 'mayor';
            } else if (estatusRaw.includes('listo') || estatusRaw.includes('finalizado')) {
                estado = 'listo';
            } else {
                estado = 'proceso';
            }

            const defecto = String(r['Defecto Mecánico'] || '').trim();
            const notas = String(r['Notas adicionales (Datos del Excel de Gabriel)'] || '').trim();

            let costo = 0;
            const matchCosto = (notas + ' ' + defecto).match(/\$\s*([\d,]+(?:\.\d+)?)/);
            if (matchCosto) {
                costo = parseFloat(matchCosto[1].replace(/,/g, '')) || 0;
            }

            let fechaRefaccion = '';
            const matchFecha = (notas + ' ' + defecto).match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
            if (matchFecha) {
                let dia = matchFecha[1].padStart(2, '0');
                let mes = matchFecha[2].padStart(2, '0');
                let anio = matchFecha[3];
                if (anio.length === 2) anio = '20' + anio;
                fechaRefaccion = anio + '-' + mes + '-' + dia;
            }

            let categoriaFalla = 'General';
            const txt = (defecto + ' ' + notas).toLowerCase();
            if (txt.includes('motor') || txt.includes('cárter') || txt.includes('fan clutch')) categoriaFalla = 'Motor';
            else if (txt.includes('freno') || txt.includes('cáliper') || txt.includes('caliper')) categoriaFalla = 'Frenos';
            else if (txt.includes('suspensión') || txt.includes('amortiguador') || txt.includes('válvula niveladora') || txt.includes('soporte')) categoriaFalla = 'Suspensión';
            else if (txt.includes('hidráulico') || txt.includes('hidraulico')) categoriaFalla = 'Hidráulico';
            else if (txt.includes('aire') || txt.includes('válvula de caudal') || txt.includes('egr')) categoriaFalla = 'Neumático/Emisiones';
            else if (txt.includes('llanta') || txt.includes('enllantado')) categoriaFalla = 'Llantas';
            else if (txt.includes('clima') || txt.includes('aire acondicionado')) categoriaFalla = 'Clima';

            const reporte = defecto && notas ? (defecto + ' • ' + notas) : (defecto || notas || 'Sin novedades');

            const previa = mapExistentes.get(id);
            const historial = previa && previa.historial ? previa.historial : [{
                fecha: new Date().toLocaleString('es-MX'),
                evento: 'Sincronización Excel: ' + reporte,
                tipo: 'sistema',
                categoria: categoriaFalla,
                autor: 'Excel Sincronizador'
            }];

            return {
                id,
                sucursal,
                estado: (previa && previa.estadoModificadoManualmente) ? previa.estado : estado,
                defecto,
                notas,
                reporte,
                categoriaFalla,
                costo: costo > 0 ? costo : (previa ? previa.costo : 0),
                fechaRefaccion: fechaRefaccion || (previa ? previa.fechaRefaccion : ''),
                piezas: (previa && previa.piezas && previa.piezas.length > 0) ? previa.piezas : [],
                historial
            };
        });

        saveUnidades(actualizadas);
        registrarMovimientoJornada('TODAS', 'sync_excel', 'Sincronización masiva desde archivo Taller Flotilla-v2.xlsx completada.');
        res.json({ success: true, count: actualizadas.length, message: 'Sincronización con Excel completada' });
    } catch (err) {
        console.error('Error al sincronizar con Excel:', err);
        res.status(500).json({ error: 'Error procesando Excel: ' + err.message });
    }
});

// 8. Historial de Bitácoras Pasadas (Snapshots)
app.get('/api/bitacoras', (req, res) => {
    const bitacoras = getBitacoras();
    res.json({ success: true, data: bitacoras });
});

// 9. Consultar Bitácora por Fecha Específica (Hoy o Pasado)
app.get('/api/bitacora/fecha/:fecha', (req, res) => {
    const fecha = String(req.params.fecha).trim();
    const bitacoras = getBitacoras();
    const encontrada = bitacoras.find(b => b.fecha === fecha);

    if (encontrada) {
        return res.json({ success: true, encontrada: true, es_en_curso: false, data: encontrada });
    }

    const jornada = getJornadaActual();
    if (jornada.fecha === fecha) {
        const unidades = getUnidades();
        const counts = { total: unidades.length, mayor: 0, proceso: 0, listo: 0, costo_acumulado: 0 };
        unidades.forEach(u => {
            if (counts[u.estado] !== undefined) counts[u.estado]++;
            counts.costo_acumulado += (Number(u.costo) || 0);
        });

        return res.json({
            success: true,
            encontrada: true,
            es_en_curso: true,
            data: {
                id: `en_curso_${fecha}`,
                fecha,
                hora: 'EN CURSO',
                titulo: `Bitácora de Hoy (En Vivo) - ${formatearFechaTitulo(fecha)}`,
                responsable: 'Jornada activa en patio',
                resumen: counts,
                notas_cierre: 'Jornada abierta. Cierre automático programado para las 12:00 AM (Medianoche).',
                movimientos: jornada.movimientos || [],
                unidades: unidades.map(u => ({
                    id: u.id,
                    sucursal: u.sucursal,
                    estado: u.estado,
                    defecto: u.defecto,
                    notas: u.notas,
                    reporte: u.reporte,
                    costo: u.costo,
                    fechaRefaccion: u.fechaRefaccion
                }))
            }
        });
    }

    res.status(404).json({ success: false, error: `No se encontró bitácora para la fecha ${fecha}` });
});

// 10. Cierre Manual de Bitácora (Guardar Corte Oficial On-Demand)
app.post('/api/bitacora/cierre-manual', (req, res) => {
    const { notas_cierre, responsable } = req.body;
    const jornada = getJornadaActual();
    const unidades = getUnidades();
    const bitacoras = getBitacoras();

    const counts = { total: unidades.length, mayor: 0, proceso: 0, listo: 0, costo_acumulado: 0 };
    unidades.forEach(u => {
        if (counts[u.estado] !== undefined) counts[u.estado]++;
        counts.costo_acumulado += (Number(u.costo) || 0);
    });

    const now = new Date();
    const hora = now.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit', second: '2-digit' });

    const snapshot = {
        id: `snap_${jornada.fecha.replace(/-/g, '')}_${Date.now()}`,
        fecha: jornada.fecha,
        hora,
        titulo: `Cierre Diario de Bitácora - ${formatearFechaTitulo(jornada.fecha)} (${hora})`,
        responsable: responsable || 'Jefe de Taller',
        resumen: counts,
        notas_cierre: notas_cierre || 'Corte formal de bitácora realizado antes de medianoche.',
        movimientos: jornada.movimientos || [],
        unidades: unidades.map(u => ({
            id: u.id,
            sucursal: u.sucursal,
            estado: u.estado,
            defecto: u.defecto,
            notas: u.notas,
            reporte: u.reporte,
            costo: u.costo,
            fechaRefaccion: u.fechaRefaccion
        }))
    };

    bitacoras.unshift(snapshot);
    saveBitacoras(bitacoras);

    registrarMovimientoJornada('TODAS', 'cierre_manual', `Corte formal de bitácora archivado a las ${hora} por ${responsable || 'Jefe de Taller'}.`);

    res.json({ success: true, snapshot });
});

// Snapshot backward-compatible
app.post('/api/bitacora/snapshot', (req, res) => {
    const { notas_cierre, responsable } = req.body;
    const jornada = getJornadaActual();
    const unidades = getUnidades();
    const bitacoras = getBitacoras();

    const counts = { total: unidades.length, mayor: 0, proceso: 0, listo: 0, costo_acumulado: 0 };
    unidades.forEach(u => {
        if (counts[u.estado] !== undefined) counts[u.estado]++;
        counts.costo_acumulado += (Number(u.costo) || 0);
    });

    const now = new Date();
    const hora = now.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit', second: '2-digit' });

    const nuevoSnapshot = {
        id: `snap_${Date.now()}`,
        fecha: jornada.fecha,
        hora,
        titulo: `Cierre de Bitácora - ${formatearFechaTitulo(jornada.fecha)} (${hora})`,
        responsable: responsable || 'Jefe de Taller',
        resumen: counts,
        notas_cierre: notas_cierre || 'Corte de operaciones y avance de patio.',
        movimientos: jornada.movimientos || [],
        unidades: unidades.map(u => ({
            id: u.id,
            sucursal: u.sucursal,
            estado: u.estado,
            defecto: u.defecto,
            notas: u.notas,
            reporte: u.reporte,
            costo: u.costo,
            fechaRefaccion: u.fechaRefaccion
        }))
    };

    bitacoras.unshift(nuevoSnapshot);
    saveBitacoras(bitacoras);
    res.json({ success: true, snapshot: nuevoSnapshot });
});

// 11. Subir foto de evidencia
app.post('/api/upload-foto', upload.single('foto'), (req, res) => {
    if (!req.file) {
        return res.status(400).json({ error: 'No se recibió ninguna imagen' });
    }
    const publicUrl = `/uploads/${req.file.filename}`;
    res.json({ success: true, url: publicUrl, filename: req.file.filename });
});

// 12. Integración con Almacén CESS (Consultar Inventario)
app.get('/api/almacen/inventario', (req, res) => {
    if (!fs.existsSync(ALMACEN_DB_FILE)) {
        return res.json({ success: true, catalog: [], message: 'Almacén CESS no vinculado' });
    }
    try {
        const almacenData = JSON.parse(fs.readFileSync(ALMACEN_DB_FILE, 'utf8'));
        const catalogObj = almacenData.catalog || {};
        let items = Object.values(catalogObj);

        const q = String(req.query.q || '').trim().toLowerCase();
        if (q) {
            items = items.filter(it => 
                (it.numero_parte && it.numero_parte.toLowerCase().includes(q)) ||
                (it.descripcion && it.descripcion.toLowerCase().includes(q)) ||
                (it.marca && it.marca.toLowerCase().includes(q))
            );
        }

        res.json({
            success: true,
            totalEnAlmacen: Object.keys(catalogObj).length,
            resultados: items.slice(0, 50).map(it => ({
                id: it.id,
                numero_parte: it.numero_parte,
                descripcion: it.descripcion,
                marca: it.marca,
                rack: it.rack,
                nivel: it.nivel,
                stock_actual: it.stock_actual,
                precio_unitario: it.precio_unitario,
                unidad: it.unidad,
                foto_referencia: it.foto_referencia
            }))
        });
    } catch (err) {
        console.error('Error consultando catálogo de almacén:', err);
        res.status(500).json({ error: err.message });
    }
});

// 13. Integración con Almacén CESS (Órdenes de Compra)
app.get('/api/almacen/ordenes', (req, res) => {
    if (!fs.existsSync(ALMACEN_DB_FILE)) {
        return res.json({ success: true, ordenes: [] });
    }
    try {
        const almacenData = JSON.parse(fs.readFileSync(ALMACEN_DB_FILE, 'utf8'));
        const ordenes = almacenData.ordenes_compra || [];
        res.json({ success: true, ordenes });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// 13.1 Actualizar estatus o datos de Orden de Compra desde Taller
app.put('/api/almacen/ordenes/:id', (req, res) => {
    try {
        const { id } = req.params;
        const { estatus, notas, proveedor, fecha_entrega_esperada, prioridad, recibido_por, factura, pin_almacen, origen } = req.body || {};

        // REGLA: En el taller no se pueden autorizar requisiciones; es tarea exclusiva de Almacén
        if (estatus === 'AUTORIZADA' || estatus === 'ORDENADA') {
            const esAlmacen = (origen === 'almacen' || pin_almacen === '0000');
            if (!esAlmacen) {
                return res.status(403).json({
                    success: false,
                    error: 'En el taller no se pueden autorizar requisiciones. Esta es una tarea exclusiva del departamento de Almacén y Compras.'
                });
            }
        }

        const almacendata = getAlmacenData();
        if (!Array.isArray(almacendata.ordenes_compra)) almacendata.ordenes_compra = [];
        const orden = almacendata.ordenes_compra.find(o => o.id === id || o.folio === id);

        if (!orden) {
            return res.status(404).json({ success: false, error: 'Orden de compra no encontrada' });
        }

        if (estatus) orden.estatus = estatus;
        if (notas !== undefined) orden.notas = String(notas).trim();
        if (proveedor) orden.proveedor = String(proveedor).trim().toUpperCase();
        if (fecha_entrega_esperada !== undefined) orden.fecha_entrega_esperada = fecha_entrega_esperada;
        if (prioridad) orden.prioridad = prioridad;
        if (estatus === 'RECIBIDA') {
            orden.fecha_recepcion = new Date().toISOString().slice(0, 10);
            orden.hora_recepcion = new Date().toLocaleTimeString('es-MX');
            orden.recibido_por = recibido_por || 'Jefe de Taller CESS';
            if (factura) orden.factura_asociada = factura;
        }

        orden.fecha_actualizacion = new Date().toISOString();
        saveAlmacenData(almacendata);

        // Sincronizar piezas asociadas en taller_db.json
        const unidades = getUnidades();
        let cambiosTaller = false;
        unidades.forEach(u => {
            if (Array.isArray(u.piezas)) {
                u.piezas.forEach(pz => {
                    if (pz && pz.folio_oc && (pz.folio_oc === orden.folio || pz.folio_oc === orden.id)) {
                        cambiosTaller = true;
                        if (estatus === 'RECIBIDA') {
                            pz.estado = 'RECIBIDA_TALLER';
                            pz.fecha_recepcion = new Date().toISOString();
                            pz.recibido_por = recibido_por || 'Jefe de Taller';
                            if (factura) pz.factura_remision = factura;
                            if (!Array.isArray(u.historial)) u.historial = [];
                            u.historial.unshift({
                                fecha: new Date().toLocaleString('es-MX'),
                                evento: `Refacción RECIBIDA EN TALLER: ${pz.cantidad || 1}x ${pz.descripcion} (OC: ${orden.folio}). Surtida y lista para montaje en Unidad ${u.id}.`,
                                tipo: 'recepcion_almacen',
                                autor: recibido_por || 'Jefe de Taller'
                            });
                        } else if (estatus === 'CANCELADA') {
                            pz.estado = 'SOLICITADA_ALMACEN';
                        } else if (estatus === 'ORDENADA') {
                            pz.estado = 'EN_ORDEN_COMPRA';
                            if (proveedor) pz.proveedor_oc = proveedor;
                        }
                    }
                });
            }
        });
        if (cambiosTaller) saveUnidades(unidades);

        res.json({ success: true, message: `Orden ${orden.folio} actualizada a ${orden.estatus}`, orden });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// 13.2 Recepción directa de Orden de Compra desde Taller
app.post('/api/almacen/ordenes/:id/recibir', (req, res) => {
    try {
        const { id } = req.params;
        const { recibido_por, factura } = req.body || {};

        const almacendata = getAlmacenData();
        if (!Array.isArray(almacendata.ordenes_compra)) almacendata.ordenes_compra = [];
        const orden = almacendata.ordenes_compra.find(o => o.id === id || o.folio === id);

        if (!orden) return res.status(404).json({ error: 'Orden no encontrada' });

        // REGLA: Si la orden está PENDIENTE, no se puede recibir en taller hasta ser autorizada por Almacén
        if (orden.estatus === 'PENDIENTE') {
            return res.status(400).json({
                error: `La requisición ${orden.folio} está PENDIENTE. Debe ser autorizada previamente por el Departamento de Almacén y Compras.`
            });
        }

        orden.estatus = 'RECIBIDA';
        orden.fecha_recepcion = new Date().toISOString().slice(0, 10);
        orden.hora_recepcion = new Date().toLocaleTimeString('es-MX');
        orden.recibido_por = recibido_por || 'Jefe de Taller CESS';
        orden.factura_asociada = factura || orden.folio;

        if (!almacendata.movimientos) almacendata.movimientos = [];
        (orden.items || []).forEach(it => {
            almacendata.movimientos.unshift({
                id: `mov_rec_${Date.now()}_${Math.floor(Math.random() * 1000)}`,
                tipo: 'E',
                fecha: orden.fecha_recepcion,
                hora: orden.hora_recepcion,
                id_pieza: it.id_pieza || '',
                numero_parte: it.numero_parte || '',
                descripcion: it.descripcion,
                marca: it.marca || '',
                cantidad: Number(it.cantidad_pedida) || 1,
                unidad_camion: it.unidad_camion || '',
                folio: factura || orden.folio,
                solicitante: orden.recibido_por,
                notas: `Recepción de Orden ${orden.folio}`,
                timestamp: new Date().toISOString()
            });
        });
        saveAlmacenData(almacendata);

        // Sincronizar en taller_db.json
        const unidades = getUnidades();
        let cambiosTaller = false;
        unidades.forEach(u => {
            if (Array.isArray(u.piezas)) {
                u.piezas.forEach(pz => {
                    if (pz && pz.folio_oc && (pz.folio_oc === orden.folio || pz.folio_oc === orden.id)) {
                        cambiosTaller = true;
                        pz.estado = 'RECIBIDA_TALLER';
                        pz.fecha_recepcion = new Date().toISOString();
                        pz.recibido_por = orden.recibido_por;
                        if (factura) pz.factura_remision = factura;
                        if (!Array.isArray(u.historial)) u.historial = [];
                        u.historial.unshift({
                            fecha: new Date().toLocaleString('es-MX'),
                            evento: `Refacción RECIBIDA EN TALLER: ${pz.cantidad || 1}x ${pz.descripcion} (OC: ${orden.folio}). Surtida y lista para montaje en Unidad ${u.id}.`,
                            tipo: 'recepcion_almacen',
                            autor: orden.recibido_por
                        });
                    }
                });
            }
        });
        if (cambiosTaller) saveUnidades(unidades);

        res.json({ success: true, message: `Orden ${orden.folio} recibida con éxito`, orden });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// 13.3 Métricas y Analítica Gerencial para Toma de Decisiones
app.get('/api/analitica/resumen', (req, res) => {
    try {
        const unidades = getUnidades();
        const { sucursal } = req.query;

        let lista = [...unidades];
        if (sucursal && sucursal !== 'TODAS') {
            lista = lista.filter(u => u.sucursal && u.sucursal.toLowerCase() === sucursal.toLowerCase());
        }

        const totalUnidades = lista.length;
        let conteoMayor = 0;
        let conteoProceso = 0;
        let conteoListo = 0;
        let costoTotal = 0;
        let totalPiezasAsignadas = 0;

        const categoriasMap = {};
        const sucursalesMap = {};

        lista.forEach(u => {
            if (u.estado === 'mayor') conteoMayor++;
            else if (u.estado === 'listo') conteoListo++;
            else conteoProceso++;

            const costoUnit = Number(u.costo) || 0;
            costoTotal += costoUnit;

            // Sucursales
            const suc = u.sucursal || 'Apodaca';
            if (!sucursalesMap[suc]) sucursalesMap[suc] = { unidades: 0, costo: 0, mayor: 0, proceso: 0, listo: 0 };
            sucursalesMap[suc].unidades++;
            sucursalesMap[suc].costo += costoUnit;
            if (u.estado === 'mayor') sucursalesMap[suc].mayor++;
            else if (u.estado === 'listo') sucursalesMap[suc].listo++;
            else sucursalesMap[suc].proceso++;

            // Categorías de falla
            const cat = u.categoriaFalla || 'General';
            if (!categoriasMap[cat]) categoriasMap[cat] = { conteo: 0, costo: 0 };
            categoriasMap[cat].conteo++;
            categoriasMap[cat].costo += costoUnit;

            // Piezas
            if (Array.isArray(u.piezas)) {
                totalPiezasAsignadas += u.piezas.length;
            }
        });

        // Top 5 unidades más costosas
        const topUnidades = [...lista]
            .sort((a, b) => (Number(b.costo) || 0) - (Number(a.costo) || 0))
            .slice(0, 5)
            .map(u => ({
                id: u.id,
                sucursal: u.sucursal,
                estado: u.estado,
                defecto: u.defecto,
                costo: Number(u.costo) || 0,
                categoriaFalla: u.categoriaFalla || 'General',
                totalPiezas: Array.isArray(u.piezas) ? u.piezas.length : 0
            }));

        // Datos de Almacén y Compras
        let resumenAlmacen = { total_ordenes: 0, pendientes: 0, autorizadas: 0, recibidas: 0, importe_total: 0 };
        const almacendata = getAlmacenData();
        if (Array.isArray(almacendata.ordenes_compra)) {
            resumenAlmacen.total_ordenes = almacendata.ordenes_compra.length;
            almacendata.ordenes_compra.forEach(o => {
                if (o.estatus === 'PENDIENTE') resumenAlmacen.pendientes++;
                else if (o.estatus === 'AUTORIZADA' || o.estatus === 'ORDENADA') resumenAlmacen.autorizadas++;
                else if (o.estatus === 'RECIBIDA') resumenAlmacen.recibidas++;
                resumenAlmacen.importe_total += (Number(o.total_importe) || 0);
            });
        }

        // Diagnóstico ejecutivo automático
        const hallazgos = [];
        const catMasFrecuente = Object.entries(categoriasMap).sort((a, b) => b[1].conteo - a[1].conteo)[0];
        if (catMasFrecuente) {
            const pct = Math.round((catMasFrecuente[1].conteo / (totalUnidades || 1)) * 100);
            hallazgos.push(`El sistema de **${catMasFrecuente[0]}** concentra el ${pct}% de las fallas mecánicas reportadas.`);
        }
        if (conteoMayor > 0) {
            const pctMayor = Math.round((conteoMayor / (totalUnidades || 1)) * 100);
            hallazgos.push(`Hay **${conteoMayor} unidad${conteoMayor > 1 ? 'es' : ''}** (${pctMayor}% del patio) detenidas por reparación mayor o autorización presupuestal.`);
        }
        if (conteoListo > 0) {
            hallazgos.push(`**${conteoListo} camión${conteoListo > 1 ? 'es están listos' : ' está listo'}** para entrega inmediata a operación.`);
        }
        if (resumenAlmacen.pendientes > 0) {
            hallazgos.push(`Existen **${resumenAlmacen.pendientes} requisición${resumenAlmacen.pendientes > 1 ? 'es' : ''}** en espera de autorización por parte de Almacén.`);
        }

        res.json({
            success: true,
            filtro_sucursal: sucursal || 'TODAS',
            totales: {
                total_unidades: totalUnidades,
                mayor: conteoMayor,
                proceso: conteoProceso,
                listo: conteoListo,
                disponibilidad_pct: totalUnidades > 0 ? Math.round((conteoListo / totalUnidades) * 100) : 0,
                costo_total: costoTotal,
                costo_promedio: totalUnidades > 0 ? Math.round(costoTotal / totalUnidades) : 0,
                total_piezas_asignadas: totalPiezasAsignadas
            },
            por_sucursal: sucursalesMap,
            por_categoria_falla: categoriasMap,
            top_unidades_costosas: topUnidades,
            almacen: resumenAlmacen,
            hallazgos
        });
    } catch (err) {
        console.error('Error generando analítica gerencial:', err);
        res.status(500).json({ error: err.message });
    }
});

// 14. Registrar Salida de Almacén al instalar pieza en unidad
app.post('/api/almacen/salida-taller', (req, res) => {
    const { unidad_id, id_pieza, cantidad, solicitante, notas } = req.body;
    if (!fs.existsSync(ALMACEN_DB_FILE)) {
        return res.status(404).json({ error: 'Archivo de almacén no encontrado' });
    }

    try {
        const almacenData = JSON.parse(fs.readFileSync(ALMACEN_DB_FILE, 'utf8'));
        const item = almacenData.catalog[id_pieza];
        if (!item) {
            return res.status(404).json({ error: 'Pieza no encontrada en catálogo de Almacén' });
        }

        const cant = Number(cantidad) || 1;
        item.stock_actual = Math.max(0, (item.stock_actual || 0) - cant);

        if (!almacenData.movimientos) almacenData.movimientos = [];
        const movimiento = {
            id: `mov_${Date.now()}`,
            tipo: 'SALIDA_TALLER',
            id_pieza,
            numero_parte: item.numero_parte,
            descripcion: item.descripcion,
            cantidad: cant,
            unidad_camion: String(unidad_id),
            solicitante: solicitante || 'Jefe de Taller',
            fecha: new Date().toISOString().slice(0, 10),
            hora: new Date().toLocaleTimeString('es-MX'),
            notas: notas || `Instalada en Unidad ${unidad_id} en Taller Diésel`
        };
        almacenData.movimientos.unshift(movimiento);

        fs.writeFileSync(ALMACEN_DB_FILE, JSON.stringify(almacenData, null, 2), 'utf8');

        let unidades = getUnidades();
        let unidad = unidades.find(u => String(u.id) === String(unidad_id));
        if (unidad) {
            if (!unidad.piezas) unidad.piezas = [];
            unidad.piezas.push(`${item.descripcion} (${cant} ${item.unidad || 'PZ'})`);
            unidad.historial.unshift({
                fecha: new Date().toLocaleString('es-MX'),
                evento: `Salida de Almacén CESS: ${cant}x ${item.descripcion} (Parte: ${item.numero_parte}) instalada en unidad.`,
                tipo: 'refaccion_almacen',
                costo: (item.precio_unitario || 0) * cant,
                autor: 'Almacén CESS'
            });
            saveUnidades(unidades);
        }

        registrarMovimientoJornada(unidad_id, 'refaccion', `Refacción instalada en Unidad ${unidad_id}: ${cant}x ${item.descripcion} (Parte: ${item.numero_parte}).`);

        res.json({
            success: true,
            message: `Pieza descontada de almacén e instalada en Unidad ${unidad_id}`,
            stock_restante: item.stock_actual,
            movimiento
        });
    } catch (err) {
        console.error('Error registrando salida:', err);
        res.status(500).json({ error: err.message });
    }
});

// --- WORKSHOP REFACCIONES & ALMACÉN CESS INTEGRATION ---

// A. Agregar refacción a una orden de taller (Con stock o enviada a Orden de Compra)
app.post('/api/taller/pieza/agregar', (req, res) => {
    try {
        const { 
            unidad_id, 
            id_pieza, 
            descripcion, 
            numero_parte, 
            marca, 
            cantidad, 
            precio_unitario, 
            rack, 
            nivel, 
            stock_actual, 
            solicitar_oc, 
            proveedor,
            fecha_estimada_arribo,
            notas, 
            solicitante 
        } = req.body;

        if (!unidad_id) return res.status(400).json({ error: 'Falta unidad_id' });
        let unidades = getUnidades();
        let unidad = unidades.find(u => String(u.id) === String(unidad_id));
        if (!unidad) return res.status(404).json({ error: `Unidad ${unidad_id} no encontrada` });

        if (!Array.isArray(unidad.piezas)) unidad.piezas = [];
        unidad.piezas = unidad.piezas.map((p, i) => normalizarPieza(p, i)).filter(Boolean);

        const cant = Number(cantidad) > 0 ? Number(cantidad) : 1;
        const pu = Number(precio_unitario) || 0;
        const stock = stock_actual !== undefined ? Number(stock_actual) : null;
        const almacendata = getAlmacenData();

        // Determinar si va a Orden de Compra o a Almacén
        const vaACompra = Boolean(solicitar_oc) || (stock !== null && stock <= 0);

        let folio_oc = '';
        let nuevoEstado = vaACompra ? 'EN_ORDEN_COMPRA' : 'SOLICITADA_ALMACEN';

        if (vaACompra) {
            if (!almacendata.ordenes_compra) almacendata.ordenes_compra = [];
            const nextNum = almacendata.ordenes_compra.length + 1;
            folio_oc = `OC-${new Date().getFullYear()}-${String(nextNum).padStart(3, '0')}`;
            
            const nuevaOC = {
                id: `oc_${Date.now()}_${Math.floor(Math.random() * 1000)}`,
                folio: folio_oc,
                fecha_creacion: new Date().toISOString().slice(0, 10),
                hora_creacion: new Date().toLocaleTimeString('es-MX'),
                proveedor: proveedor || marca || 'Proveedor Requisición Taller',
                solicitante: solicitante || 'Jefe de Taller CESS',
                estatus: 'PENDIENTE',
                prioridad: 'ALTA',
                notas: `Requisición directa desde Taller para Unidad ${unidad_id}. ${notas || ''}`.trim(),
                almacen_id: 'refacciones',
                almacen: 'Almacén de Refacciones',
                fecha_entrega_esperada: fecha_estimada_arribo || '',
                items: [{
                    id: `item_${Date.now()}`,
                    id_pieza: id_pieza || '',
                    numero_parte: numero_parte || 'S/N',
                    descripcion: descripcion || 'Refacción de taller',
                    marca: marca || '',
                    unidad: 'PZ',
                    stock_actual_al_pedir: stock === null ? 0 : stock,
                    cantidad_pedida: cant,
                    precio_estimado: pu,
                    total_estimado: pu * cant,
                    unidad_camion: String(unidad_id)
                }],
                total_items: 1,
                total_piezas: cant,
                total_importe: pu * cant,
                timestamp: new Date().toISOString()
            };
            almacendata.ordenes_compra.unshift(nuevaOC);
            saveAlmacenData(almacendata);
        }

        const nuevaPieza = normalizarPieza({
            id_pieza: id_pieza || null,
            descripcion: descripcion || 'Refacción sin descripción',
            numero_parte: numero_parte || '',
            marca: marca || '',
            cantidad: cant,
            unidad_medida: 'PZ',
            precio_unitario: pu,
            costo_total: pu * cant,
            rack: rack || '',
            nivel: nivel || '',
            stock_disponible: stock === null ? 0 : stock,
            estado: nuevoEstado,
            folio_oc: folio_oc,
            proveedor_oc: proveedor || marca || '',
            fecha_estimada_arribo: fecha_estimada_arribo || '',
            notas: notas || ''
        });

        unidad.piezas.push(nuevaPieza);

        if (pu > 0) {
            unidad.costo = (Number(unidad.costo) || 0) + (pu * cant);
        }

        if (!Array.isArray(unidad.historial)) unidad.historial = [];
        const eventoHistorial = vaACompra 
            ? `[Requisición OC (${folio_oc})]: ${cant}x ${nuevaPieza.descripcion} (P/N: ${nuevaPieza.numero_parte || 'S/N'}). Pieza sin existencia en almacén.`
            : `[Solicitud Almacén]: ${cant}x ${nuevaPieza.descripcion} (P/N: ${nuevaPieza.numero_parte || 'S/N'}) en espera de autorización de salida.`;

        unidad.historial.unshift({
            fecha: new Date().toLocaleString('es-MX'),
            evento: eventoHistorial,
            tipo: vaACompra ? 'orden_compra' : 'solicitud_almacen',
            autor: solicitante || 'Jefe de Taller',
            costo: pu * cant
        });

        saveUnidades(unidades);
        registrarMovimientoJornada(unidad_id, vaACompra ? 'orden_compra' : 'solicitud_almacen', eventoHistorial, solicitante || 'Jefe de Taller');

        res.json({
            success: true,
            pieza: nuevaPieza,
            folio_oc,
            costo_total_unidad: unidad.costo,
            unidad
        });
    } catch (err) {
        console.error('Error al agregar pieza a orden de taller:', err);
        res.status(500).json({ error: err.message });
    }
});

// B. Mandar pieza a Orden de Compra / Requisición
app.post('/api/taller/pieza/solicitar-compra', (req, res) => {
    try {
        const { unidad_id, pieza_id, proveedor, fecha_estimada_arribo, notas, solicitante } = req.body;
        let unidades = getUnidades();
        let unidad = unidades.find(u => String(u.id) === String(unidad_id));
        if (!unidad) return res.status(404).json({ error: 'Unidad no encontrada' });

        unidad.piezas = (unidad.piezas || []).map((p, i) => normalizarPieza(p, i));
        let pieza = unidad.piezas.find(p => String(p.id) === String(pieza_id));
        if (!pieza) return res.status(404).json({ error: 'Pieza no encontrada en la unidad' });

        const almacendata = getAlmacenData();
        if (!almacendata.ordenes_compra) almacendata.ordenes_compra = [];
        const nextNum = almacendata.ordenes_compra.length + 1;
        const folio_oc = `OC-${new Date().getFullYear()}-${String(nextNum).padStart(3, '0')}`;

        const nuevaOC = {
            id: `oc_${Date.now()}_${Math.floor(Math.random() * 1000)}`,
            folio: folio_oc,
            fecha_creacion: new Date().toISOString().slice(0, 10),
            hora_creacion: new Date().toLocaleTimeString('es-MX'),
            proveedor: proveedor || pieza.marca || 'Proveedor Requisición Taller',
            solicitante: solicitante || 'Jefe de Taller CESS',
            estatus: 'PENDIENTE',
            prioridad: 'ALTA',
            notas: `Requisición desde Taller para Unidad ${unidad_id}. ${notas || ''}`.trim(),
            almacen_id: 'refacciones',
            almacen: 'Almacén de Refacciones',
            fecha_entrega_esperada: fecha_estimada_arribo || '',
            items: [{
                id: `item_${Date.now()}`,
                id_pieza: pieza.id_pieza || '',
                numero_parte: pieza.numero_parte || 'S/N',
                descripcion: pieza.descripcion,
                marca: pieza.marca,
                cantidad_pedida: pieza.cantidad,
                precio_estimado: pieza.precio_unitario,
                total_estimado: pieza.costo_total,
                unidad_camion: String(unidad_id)
            }],
            total_items: 1,
            total_piezas: pieza.cantidad,
            total_importe: pieza.costo_total,
            timestamp: new Date().toISOString()
        };
        almacendata.ordenes_compra.unshift(nuevaOC);
        saveAlmacenData(almacendata);

        pieza.estado = 'EN_ORDEN_COMPRA';
        pieza.folio_oc = folio_oc;
        pieza.proveedor_oc = proveedor || pieza.marca || '';
        pieza.fecha_estimada_arribo = fecha_estimada_arribo || '';
        if (notas) pieza.notas = (pieza.notas ? pieza.notas + ' | ' : '') + notas;

        const evento = `[Requisición OC (${folio_oc})]: ${pieza.cantidad}x ${pieza.descripcion} (P/N: ${pieza.numero_parte || 'S/N'}) enviada a Compras.`;
        if (!Array.isArray(unidad.historial)) unidad.historial = [];
        unidad.historial.unshift({
            fecha: new Date().toLocaleString('es-MX'),
            evento,
            tipo: 'orden_compra',
            autor: solicitante || 'Jefe de Taller'
        });

        saveUnidades(unidades);
        registrarMovimientoJornada(unidad_id, 'orden_compra', evento, solicitante || 'Jefe de Taller');

        res.json({ success: true, folio_oc, pieza, unidad });
    } catch (err) {
        console.error('Error al generar orden de compra:', err);
        res.status(500).json({ error: err.message });
    }
});

// C. Autorización de Salida por parte del Encargado de Inventario del Almacén
app.post('/api/taller/pieza/autorizar-salida', (req, res) => {
    try {
        const { unidad_id, pieza_id, autorizado_por, notas, pin_almacen } = req.body;

        // REGLA: La autorización de salida de refacciones requiere PIN del Encargado de Almacén
        if (pin_almacen && pin_almacen.trim() !== '0000') {
            return res.status(403).json({
                error: 'PIN de Almacén inválido. Solo el Encargado de Almacén puede autorizar salidas de inventario.'
            });
        }

        let unidades = getUnidades();
        let unidad = unidades.find(u => String(u.id) === String(unidad_id));
        if (!unidad) return res.status(404).json({ error: 'Unidad no encontrada' });

        unidad.piezas = (unidad.piezas || []).map((p, i) => normalizarPieza(p, i));
        let pieza = unidad.piezas.find(p => String(p.id) === String(pieza_id));
        if (!pieza) return res.status(404).json({ error: 'Pieza no encontrada en la unidad' });

        const almacendata = getAlmacenData();
        let itemCatalog = null;
        if (pieza.id_pieza && almacendata.catalog && almacendata.catalog[pieza.id_pieza]) {
            itemCatalog = almacendata.catalog[pieza.id_pieza];
        } else if (pieza.numero_parte && almacendata.catalog) {
            itemCatalog = Object.values(almacendata.catalog).find(it => it.numero_parte && String(it.numero_parte).trim().toLowerCase() === String(pieza.numero_parte).trim().toLowerCase());
        }

        if (itemCatalog) {
            itemCatalog.stock_actual = Math.max(0, (itemCatalog.stock_actual || 0) - pieza.cantidad);
        }

        if (!almacendata.movimientos) almacendata.movimientos = [];
        const folioSalida = `SAL-${Date.now().toString().slice(-6)}`;
        const mov = {
            id: `mov_${Date.now()}_${Math.floor(Math.random() * 1000)}`,
            tipo: 'S',
            fecha: new Date().toISOString().slice(0, 10),
            hora: new Date().toLocaleTimeString('es-MX'),
            id_pieza: pieza.id_pieza || (itemCatalog ? itemCatalog.id : ''),
            numero_parte: pieza.numero_parte || (itemCatalog ? itemCatalog.numero_parte : ''),
            descripcion: pieza.descripcion,
            marca: pieza.marca,
            ubicacion: `${pieza.rack || ''} ${pieza.nivel ? '/ ' + pieza.nivel : ''}`.trim() || 'Almacén Central',
            cantidad: pieza.cantidad,
            unidad_camion: String(unidad_id),
            folio: folioSalida,
            solicitante: autorizado_por || 'Encargado de Inventario',
            notas: notas || `Salida autorizada de almacén para Unidad ${unidad_id}`,
            timestamp: new Date().toISOString()
        };
        almacendata.movimientos.unshift(mov);
        saveAlmacenData(almacendata);

        pieza.estado = 'SALIDA_AUTORIZADA';
        pieza.autorizado_por = autorizado_por || 'Encargado de Inventario';
        pieza.fecha_autorizacion = new Date().toISOString();
        if (itemCatalog) pieza.stock_disponible = itemCatalog.stock_actual;

        const evento = `[Salida Autorizada]: ${pieza.cantidad}x ${pieza.descripcion} (Folio ${folioSalida}) despachada hacia Taller por ${pieza.autorizado_por}.`;
        if (!Array.isArray(unidad.historial)) unidad.historial = [];
        unidad.historial.unshift({
            fecha: new Date().toLocaleString('es-MX'),
            evento,
            tipo: 'salida_autorizada',
            autor: pieza.autorizado_por
        });

        saveUnidades(unidades);
        registrarMovimientoJornada(unidad_id, 'salida_autorizada', evento, pieza.autorizado_por);

        res.json({
            success: true,
            pieza,
            stock_restante: itemCatalog ? itemCatalog.stock_actual : null,
            movimiento: mov,
            unidad
        });
    } catch (err) {
        console.error('Error autorizando salida:', err);
        res.status(500).json({ error: err.message });
    }
});

// D. Dar entrada física en taller al recibir pieza traída por proveedor/compras
app.post('/api/taller/pieza/dar-entrada', (req, res) => {
    try {
        const { unidad_id, pieza_id, recibido_por, folio_factura, notas } = req.body;
        let unidades = getUnidades();
        let unidad = unidades.find(u => String(u.id) === String(unidad_id));
        if (!unidad) return res.status(404).json({ error: 'Unidad no encontrada' });

        unidad.piezas = (unidad.piezas || []).map((p, i) => normalizarPieza(p, i));
        let pieza = unidad.piezas.find(p => String(p.id) === String(pieza_id));
        if (!pieza) return res.status(404).json({ error: 'Pieza no encontrada en la unidad' });

        const almacendata = getAlmacenData();

        // Actualizar Orden de Compra si existe
        if (pieza.folio_oc && almacendata.ordenes_compra) {
            const oc = almacendata.ordenes_compra.find(o => o.folio === pieza.folio_oc);
            if (oc) {
                oc.estatus = 'RECIBIDA';
                oc.fecha_recepcion = new Date().toISOString().slice(0, 10);
                oc.hora_recepcion = new Date().toLocaleTimeString('es-MX');
                oc.recibido_por = recibido_por || 'Jefe de Taller';
                if (folio_factura) oc.factura_asociada = folio_factura;
            }
        }

        // Registrar movimiento de Entrada física en almacén/taller
        if (!almacendata.movimientos) almacendata.movimientos = [];
        const folioEntrada = folio_factura || pieza.folio_oc || `ENT-TALLER-${Date.now().toString().slice(-4)}`;
        const movEntrada = {
            id: `mov_${Date.now()}_${Math.floor(Math.random() * 1000)}`,
            tipo: 'E',
            fecha: new Date().toISOString().slice(0, 10),
            hora: new Date().toLocaleTimeString('es-MX'),
            id_pieza: pieza.id_pieza || '',
            numero_parte: pieza.numero_parte || '',
            descripcion: pieza.descripcion,
            marca: pieza.marca || '',
            cantidad: pieza.cantidad,
            unidad_camion: String(unidad_id),
            folio: folioEntrada,
            solicitante: recibido_por || 'Jefe de Taller',
            notas: `Recepción física de refacción para Unidad ${unidad_id}. ${notas || ''}`.trim(),
            timestamp: new Date().toISOString()
        };
        almacendata.movimientos.unshift(movEntrada);
        saveAlmacenData(almacendata);

        pieza.estado = 'RECIBIDA_TALLER';
        pieza.recibido_por = recibido_por || 'Jefe de Taller';
        pieza.fecha_recepcion = new Date().toISOString();
        pieza.factura_remision = folio_factura || '';

        const evento = `[Recepción en Taller]: ${pieza.cantidad}x ${pieza.descripcion} (${pieza.folio_oc ? 'OC: ' + pieza.folio_oc : 'Directa'}). Lista para instalación en Unidad ${unidad_id}.`;
        if (!Array.isArray(unidad.historial)) unidad.historial = [];
        unidad.historial.unshift({
            fecha: new Date().toLocaleString('es-MX'),
            evento,
            tipo: 'recepcion_taller',
            autor: pieza.recibido_por
        });

        saveUnidades(unidades);
        registrarMovimientoJornada(unidad_id, 'recepcion_taller', evento, pieza.recibido_por);

        res.json({
            success: true,
            pieza,
            unidad
        });
    } catch (err) {
        console.error('Error al dar entrada a pieza:', err);
        res.status(500).json({ error: err.message });
    }
});

// E. Marcar pieza como INSTALADA en el tractocamión
app.post('/api/taller/pieza/instalar', (req, res) => {
    try {
        const { unidad_id, pieza_id, mecanico, notas } = req.body;
        let unidades = getUnidades();
        let unidad = unidades.find(u => String(u.id) === String(unidad_id));
        if (!unidad) return res.status(404).json({ error: 'Unidad no encontrada' });

        unidad.piezas = (unidad.piezas || []).map((p, i) => normalizarPieza(p, i));
        let pieza = unidad.piezas.find(p => String(p.id) === String(pieza_id));
        if (!pieza) return res.status(404).json({ error: 'Pieza no encontrada en la unidad' });

        pieza.estado = 'INSTALADA';
        pieza.instalado_por = mecanico || 'Mecánico de Patio';
        pieza.fecha_instalacion = new Date().toISOString();
        if (notas) pieza.notas = (pieza.notas ? pieza.notas + ' | ' : '') + notas;

        const evento = `[Instalación]: ${pieza.cantidad}x ${pieza.descripcion} (P/N: ${pieza.numero_parte || 'S/N'}) instalada por ${pieza.instalado_por}.`;
        if (!Array.isArray(unidad.historial)) unidad.historial = [];
        unidad.historial.unshift({
            fecha: new Date().toLocaleString('es-MX'),
            evento,
            tipo: 'pieza_instalada',
            autor: pieza.instalado_por
        });

        saveUnidades(unidades);
        registrarMovimientoJornada(unidad_id, 'pieza_instalada', evento, pieza.instalado_por);

        res.json({ success: true, pieza, unidad });
    } catch (err) {
        console.error('Error al marcar pieza como instalada:', err);
        res.status(500).json({ error: err.message });
    }
});

// F. Eliminar refacción de la orden de taller
app.delete('/api/taller/pieza/:unidad_id/:pieza_id', (req, res) => {
    try {
        const { unidad_id, pieza_id } = req.params;
        let unidades = getUnidades();
        let unidad = unidades.find(u => String(u.id) === String(unidad_id));
        if (!unidad) return res.status(404).json({ error: 'Unidad no encontrada' });

        unidad.piezas = (unidad.piezas || []).map((p, i) => normalizarPieza(p, i));
        const idx = unidad.piezas.findIndex(p => String(p.id) === String(pieza_id));
        if (idx === -1) return res.status(404).json({ error: 'Pieza no encontrada' });

        const removida = unidad.piezas[idx];
        unidad.piezas.splice(idx, 1);

        if (removida.costo_total && unidad.costo) {
            unidad.costo = Math.max(0, (Number(unidad.costo) || 0) - removida.costo_total);
        }

        const evento = `[Baja de Orden]: Refacción ${removida.descripcion} (P/N: ${removida.numero_parte || 'S/N'}) retirada de orden.`;
        if (!Array.isArray(unidad.historial)) unidad.historial = [];
        unidad.historial.unshift({
            fecha: new Date().toLocaleString('es-MX'),
            evento,
            tipo: 'pieza_retirada',
            autor: 'Jefe de Taller'
        });

        saveUnidades(unidades);
        registrarMovimientoJornada(unidad_id, 'pieza_retirada', evento, 'Jefe de Taller');

        res.json({ success: true, unidad });
    } catch (err) {
        console.error('Error retirando pieza:', err);
        res.status(500).json({ error: err.message });
    }
});

// 15. Motor IA: Procesar Dictado por Voz (NLP + Redacción Técnica)
app.post('/api/ia/procesar-dictado', async (req, res) => {
    const { transcript, unitId, apiKey } = req.body;
    if (!transcript) {
        return res.status(400).json({ error: 'Falta el texto de dictado' });
    }

    const key = (apiKey || process.env.GEMINI_API_KEY || '').trim();

    if (key) {
        try {
            const prompt = `Actúa como el Ingeniero en Jefe de Mantenimiento de Flotillas de Tractocamiones Diésel.
Analiza la siguiente transcripción verbal del mecánico en patio:
"${transcript}"

${unitId ? `Unidad en contexto: ${unitId}` : 'Determina si menciona un número de unidad.'}

Devuelve un JSON estrictamente con esta estructura:
{
  "unidadDetectada": "string o null si no se menciona y no había contexto",
  "nuevoEstado": "mayor" | "proceso" | "listo" | null,
  "redaccionTecnica": "Redacción estructurada, profesional y técnica para bitácora ejecutiva. Corrige jerga y modismos sin perder datos clave.",
  "refaccionesFaltantes": ["lista de refacciones mencionadas"],
  "fechaProgramada": "YYYY-MM-DD o null si se mencionó una fecha",
  "costoEstimado": number o 0,
  "categoriaFalla": "Motor" | "Frenos" | "Suspensión" | "Hidráulico" | "Neumático/Emisiones" | "Llantas" | "Clima" | "General"
}
Solo entrega el JSON válido, sin markdown ni explicaciones.`;

            const models = ['gemini-2.0-flash', 'gemini-1.5-flash', 'gemini-2.5-flash'];
            for (const model of models) {
                try {
                    const geminiRes = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${key}`, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            contents: [{ parts: [{ text: prompt }] }],
                            generationConfig: { response_mime_type: 'application/json', temperature: 0.1 }
                        })
                    });
                    if (geminiRes.ok) {
                        const data = await geminiRes.json();
                        const text = data.candidates?.[0]?.content?.parts?.[0]?.text;
                        if (text) {
                            const parsed = JSON.parse(text);
                            if (parsed.unidadDetectada) {
                                registrarMovimientoJornada(parsed.unidadDetectada, 'dictado_ia', parsed.redaccionTecnica);
                            }
                            return res.json({ success: true, engine: 'gemini', model, result: parsed });
                        }
                    }
                } catch (e) {
                    console.warn(`Fallo con modelo ${model}:`, e.message);
                }
            }
        } catch (iaErr) {
            console.warn('Error en llamada a Gemini:', iaErr.message);
        }
    }

    const localResult = procesarDictadoOffline(transcript, unitId);
    if (localResult.unidadDetectada) {
        registrarMovimientoJornada(localResult.unidadDetectada, 'dictado_voz', localResult.redaccionTecnica);
    }
    res.json({ success: true, engine: 'offline_nlp', result: localResult });
});

function procesarDictadoOffline(transcript, contextUnitId) {
    const txt = transcript.toLowerCase();

    let unidadDetectada = contextUnitId || null;
    const matchUnidad = txt.match(/(?:unidad|camion|camión|la|el)\s*#?\s*(\d{1,4})/);
    if (matchUnidad) {
        unidadDetectada = matchUnidad[1];
    }

    let nuevoEstado = null;
    if (txt.includes('listo') || txt.includes('lista') || txt.includes('terminado') || txt.includes('entregada') || txt.includes('liberada')) {
        nuevoEstado = 'listo';
    } else if (txt.includes('proceso') || txt.includes('montado') || txt.includes('armando') || txt.includes('trabajando') || txt.includes('taller')) {
        nuevoEstado = 'proceso';
    } else if (txt.includes('presupuesto') || txt.includes('dinero') || txt.includes('parada por') || txt.includes('mayor')) {
        nuevoEstado = 'mayor';
    }

    let fechaProgramada = null;
    const matchFecha = txt.match(/(\d{1,2})\s*(?:de|\/)\s*([a-z]+|\d{1,2})(?:\s*(?:de|\/)\s*(\d{2,4}))?/);
    if (matchFecha) {
        const meses = { enero: '01', febrero: '02', marzo: '03', abril: '04', mayo: '05', junio: '06', julio: '07', agosto: '08', septiembre: '09', octubre: '10', noviembre: '11', diciembre: '12' };
        let dia = matchFecha[1].padStart(2, '0');
        let mes = meses[matchFecha[2]] || matchFecha[2].padStart(2, '0');
        let anio = matchFecha[3] || '2026';
        if (anio.length === 2) anio = '20' + anio;
        fechaProgramada = `${anio}-${mes}-${dia}`;
    }

    let costoEstimado = 0;
    const matchCosto = txt.match(/(?:\$|pesos|costo|cotizacion|cotización)\s*([\d,]+)/);
    if (matchCosto) {
        costoEstimado = parseFloat(matchCosto[1].replace(/,/g, '')) || 0;
    }

    let redaccion = transcript
        .replace(/la verdad es que/gi, '')
        .replace(/oye|mira|fíjate que/gi, '')
        .replace(/está tirando agüita|tira agua/gi, 'fuga en circuito de refrigeración')
        .replace(/trae un ruido raro/gi, 'emisión de ruido anómalo')
        .replace(/están como amarrados|amarrados/gi, 'bloqueados/atascados')
        .replace(/pídelos para|pídelo para/gi, 'requisición requerida para fecha:')
        .trim();

    if (unidadDetectada) {
        redaccion = `Unidad ${unidadDetectada}: ${redaccion.charAt(0).toUpperCase() + redaccion.slice(1)}`;
    }

    return {
        unidadDetectada,
        nuevoEstado,
        redaccionTecnica: redaccion,
        refaccionesFaltantes: [],
        fechaProgramada,
        costoEstimado,
        categoriaFalla: 'General'
    };
}

// 16. Motor IA: Escáner de Documentos / Cotizaciones (Vision AI)
app.post('/api/ia/escanear-documento', upload.single('imagen'), async (req, res) => {
    const key = (req.body.apiKey || process.env.GEMINI_API_KEY || '').trim();
    if (!key) {
        return res.status(400).json({ error: 'Se requiere Gemini API Key para Vision AI OCR' });
    }

    let base64Image = '';
    let mimeType = 'image/jpeg';

    if (req.file) {
        base64Image = fs.readFileSync(req.file.path).toString('base64');
        mimeType = req.file.mimetype;
    } else if (req.body.base64) {
        base64Image = req.body.base64.replace(/^data:image\/[a-z]+;base64,/, '');
    } else {
        return res.status(400).json({ error: 'No se envió imagen' });
    }

    const prompt = `Analiza esta fotografía de una nota, cotización de refacciones o factura para taller mecánico de tractocamiones.
Extrae los siguientes datos en JSON estricto:
{
  "proveedor": "Nombre del proveedor o tienda de refacciones",
  "folio": "Número de nota o cotización",
  "totalMxn": number con el monto total en pesos mexicanos,
  "unidadSugerida": "Número de camión o unidad si aparece escrito, o null",
  "piezas": [
    { "descripcion": "Nombre de la refacción", "cantidad": number, "precio": number }
  ],
  "resumen": "Resumen conciso del documento"
}
Solo JSON válido.`;

    try {
        const geminiRes = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${key}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                contents: [{
                    parts: [
                        { text: prompt },
                        { inline_data: { mime_type: mimeType, data: base64Image } }
                    ]
                }],
                generationConfig: { response_mime_type: 'application/json', temperature: 0.1 }
            })
        });

        if (!geminiRes.ok) {
            const errBody = await geminiRes.text();
            return res.status(500).json({ error: 'Error de Gemini Vision: ' + errBody });
        }

        const data = await geminiRes.json();
        const text = data.candidates?.[0]?.content?.parts?.[0]?.text;
        const parsed = JSON.parse(text);
        res.json({ success: true, data: parsed });
    } catch (err) {
        console.error('Error Vision OCR:', err);
        res.status(500).json({ error: err.message });
    }
});

// 17. Motor IA: Resúmenes Ejecutivos On-Demand
app.post('/api/ia/resumen-ejecutivo', async (req, res) => {
    const { pregunta, apiKey } = req.body;
    const unidades = getUnidades();

    const metricas = {
        total: unidades.length,
        porEstado: { mayor: 0, proceso: 0, listo: 0 },
        porSucursal: {},
        inversionTotal: 0,
        unidadesDetenidasPresupuesto: [],
        unidadesListas: [],
        comprasProgramadas: []
    };

    unidades.forEach(u => {
        if (metricas.porEstado[u.estado] !== undefined) metricas.porEstado[u.estado]++;
        metricas.porSucursal[u.sucursal] = (metricas.porSucursal[u.sucursal] || 0) + 1;
        metricas.inversionTotal += (Number(u.costo) || 0);

        if (u.estado === 'mayor') {
            metricas.unidadesDetenidasPresupuesto.push(`Unidad ${u.id} (${u.sucursal}): ${u.reporte} [Costo: $${(u.costo||0).toLocaleString()} MXN]`);
        }
        if (u.estado === 'listo') {
            metricas.unidadesListas.push(`Unidad ${u.id} (${u.sucursal})`);
        }
        if (u.fechaRefaccion) {
            metricas.comprasProgramadas.push(`Unidad ${u.id}: Prog ${u.fechaRefaccion} • ${u.reporte}`);
        }
    });

    const key = (apiKey || process.env.GEMINI_API_KEY || '').trim();
    if (key && pregunta) {
        try {
            const prompt = `Eres el Asistente Ejecutivo de Gerencia para Flotillas Diésel.
Datos actuales en tiempo real del Taller:
${JSON.stringify(metricas, null, 2)}

Pregunta del directivo: "${pregunta}"

Responde con precisión ejecutiva, cifras exactas, tono corporativo profesional y bullets concisos. No inventes datos.`;

            const geminiRes = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${key}`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] })
            });

            if (geminiRes.ok) {
                const data = await geminiRes.json();
                const respuesta = data.candidates?.[0]?.content?.parts?.[0]?.text;
                return res.json({ success: true, respuesta, metricas });
            }
        } catch (e) {
            console.warn('Fallo Gemini en resumen:', e.message);
        }
    }

    let respuestaAutomatica = `*Resumen de Estado de Flotilla (${new Date().toLocaleDateString('es-MX')}):*\n\n`;
    respuestaAutomatica += `• Total en Taller: ${metricas.total} unidades\n`;
    respuestaAutomatica += `• En Reparación Mayor / Presupuesto: ${metricas.porEstado.mayor} unidades\n`;
    respuestaAutomatica += `• En Proceso de Reparación: ${metricas.porEstado.proceso} unidades\n`;
    respuestaAutomatica += `• Listas para Entrega: ${metricas.porEstado.listo} unidades\n`;
    respuestaAutomatica += `• Inversión Cotizada Registrada: $${metricas.inversionTotal.toLocaleString()} MXN\n\n`;
    respuestaAutomatica += `*Unidades Detenidas por Presupuesto:*\n${metricas.unidadesDetenidasPresupuesto.map(u => '  - ' + u).join('\n') || '  (Ninguna)'}`;

    res.json({ success: true, respuesta: respuestaAutomatica, metricas });
});

// --- 18. EXPORTACIÓN PROFESIONAL EXCEL PARA DIRECCIÓN / JEFE ---
app.get('/api/exportar-excel-ejecutivo', (req, res) => {
    try {
        const unidades = getUnidades();
        const fechaHoy = getFechaActualISO();

        const prevList = getPreventivos();
        const orden = { considerar: 0, mayor: 1, proceso: 2, listo: 3 };
        const ordenadas = [...unidades].sort((a, b) => (orden[a.estado] ?? 9) - (orden[b.estado] ?? 9));

        const rows = ordenadas.map(u => {
            const p = prevList.find(x => String(x.unidad).trim() === String(u.id).trim() || (Number(x.unidad) > 0 && Number(x.unidad) === Number(u.id)));
            let prevTag = 'Al Día';
            let kmRestPrev = '';
            if (p) {
                const estNorm = (p.estatus || '').toUpperCase().trim();
                kmRestPrev = Number(p.km_restantes) || 0;
                if (estNorm === 'VENCIDO') {
                    prevTag = `🚨 VENCIDO (${kmRestPrev.toLocaleString()} km)`;
                } else if (estNorm.includes('PROXIM')) {
                    prevTag = `⚠️ PRÓXIMO (+${kmRestPrev.toLocaleString()} km)`;
                } else {
                    prevTag = `✅ Al Día (+${kmRestPrev.toLocaleString()} km)`;
                }
            } else {
                prevTag = 'Sin registro';
            }

            const estTablero = u.estado === 'considerar' 
                ? 'A Considerar / Detección' 
                : (u.estado === 'mayor' 
                    ? 'Reparación Mayor / Presupuesto' 
                    : (u.estado === 'proceso' ? 'En Reparación Activa' : 'Lista para Entrega'));

            return {
                'No. Unidad': u.id,
                'Terminal': u.sucursal || 'Apodaca',
                'Estatus Tablero': estTablero,
                'Diagnóstico / Defecto Mecánico': u.defecto || '',
                'Avances Técnicos / Notas': u.notas || u.reporte || '',
                'Folio Almacén': u.folio_almacen ? `#${u.folio_almacen}` : 'Sin folio',
                'Costo Estimado (MXN)': (Number(u.costo) || 0) + (Number(u.costo_refacciones) || 0),
                '¿Le toca Preventivo?': prevTag,
                'Km Restantes Preventivo': kmRestPrev,
                'Fecha Prog. Entrega': u.fechaRefaccion || 'Sin programar'
            };
        });

        // Fila de resumen total
        const totalCosto = unidades.reduce((acc, u) => acc + (Number(u.costo) || 0) + (Number(u.costo_refacciones) || 0), 0);
        rows.push({
            'No. Unidad': 'TOTALES',
            'Terminal': `${unidades.length} Unidades`,
            'Estatus Tablero': `Mayor: ${unidades.filter(u=>u.estado==='mayor').length} | Proceso: ${unidades.filter(u=>u.estado==='proceso').length} | Listas: ${unidades.filter(u=>u.estado==='listo').length}`,
            'Diagnóstico / Defecto Mecánico': `Generado el ${new Date().toLocaleString('es-MX')}`,
            'Avances Técnicos / Notas': 'Reporte Ejecutivo Oficial CESS Flotilla',
            'Folio Almacén': `${unidades.filter(u => u.folio_almacen).length} con folio`,
            'Costo Estimado (MXN)': totalCosto,
            '¿Le toca Preventivo?': 'Corte Oficial',
            'Km Restantes Preventivo': '',
            'Fecha Prog. Entrega': 'Corte Oficial'
        });

        const wb = xlsx.utils.book_new();
        const ws = xlsx.utils.json_to_sheet(rows);

        ws['!cols'] = [
            { wch: 12 }, // Unidad
            { wch: 14 }, // Terminal
            { wch: 32 }, // Estatus Tablero
            { wch: 38 }, // Defecto
            { wch: 48 }, // Notas
            { wch: 16 }, // Folio Almacén
            { wch: 22 }, // Costo
            { wch: 30 }, // ¿Le toca Preventivo?
            { wch: 22 }, // Km Restantes Preventivo
            { wch: 22 }  // Fecha Prog
        ];

        xlsx.utils.book_append_sheet(wb, ws, 'Bitácora Flotilla CESS');
        const buf = xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });

        const filename = `Bitacora_Ejecutiva_Flotilla_CESS_${fechaHoy}.xlsx`;
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.send(buf);
    } catch (err) {
        console.error('Error generando Excel ejecutivo:', err);
        res.status(500).json({ error: 'Error al generar Excel: ' + err.message });
    }
});

// --- 19. GESTIÓN DE TÚNEL REMOTO CLOUDFLARE PARA COMPARTIR EN VIVO (PERSISTENTE CON WATCHDOG) ---
const https = require('https');
const TUNNEL_URL_FILE = path.join(__dirname, 'active_taller_tunnel_url.txt');
let tunnelProcess = null;
let currentTunnelUrl = null;
let watchdogFallosConsecutivos = 0;

if (fs.existsSync(TUNNEL_URL_FILE)) {
    try {
        const saved = fs.readFileSync(TUNNEL_URL_FILE, 'utf8').trim();
        if (saved && saved.startsWith('https://') && !saved.includes('api.trycloudflare.com')) currentTunnelUrl = saved;
    } catch (e) {}
}

async function verificarSiTunelEstaVivo(url) {
    if (!url || !url.startsWith('https://')) return false;
    return new Promise((resolve) => {
        try {
            const req = https.get(url, { timeout: 12000 }, (res) => {
                if (res.statusCode < 500) resolve(true);
                else resolve(false);
            });
            req.on('error', () => resolve(false));
            req.on('timeout', () => { req.destroy(); resolve(false); });
        } catch (e) {
            resolve(false);
        }
    });
}

let tunnelFailCount = 0;
function iniciarTunelAutomatico() {
    // 1. Si ya tenemos una URL guardada en archivo y el proceso sigue corriendo, conservarlo
    if (tunnelProcess) return;

    const cloudflaredPath = path.join(__dirname, 'almacen-app', 'cloudflared.exe');
    if (!fs.existsSync(cloudflaredPath) || process.env.RENDER || process.env.IS_CLOUD) {
        return;
    }

    if (tunnelFailCount > 5) {
        // Pausar reintentos agresivos para evitar sobrecarga si no hay salida a internet
        setTimeout(() => { tunnelFailCount = 0; iniciarTunelAutomatico(); }, 60000);
        return;
    }

    try {
        const { spawn } = require('child_process');
        const logFile = path.join(__dirname, 'cloudflared_taller.log');
        const outStream = fs.createWriteStream(logFile, { flags: 'a' });

        // --ha-connections 4 y --retries 20 para máxima resiliencia sin rotación de URL
        tunnelProcess = spawn(cloudflaredPath, [
            'tunnel',
            '--ha-connections', '4',
            '--retries', '20',
            '--url', `http://localhost:${PORT}`
        ], {
            detached: true,
            stdio: ['ignore', 'pipe', 'pipe']
        });

        if (typeof tunnelProcess.unref === 'function') {
            tunnelProcess.unref();
        }

        const checkOutput = (data) => {
            const str = data.toString();
            try { outStream.write(str); } catch (e) {}

            if (str.includes('Unauthorized: Tunnel not found') || str.includes('Tunnel not found') || str.includes('failed to request quick Tunnel')) {
                console.warn('[Cloudflare Tunnel] Sesión del túnel expirada por Cloudflare. Reiniciando proceso de túnel en 2 segundos...');
                if (tunnelProcess) {
                    try { tunnelProcess.kill(); } catch (e) {}
                    try {
                        const { execSync } = require('child_process');
                        execSync('taskkill /f /im cloudflared.exe', { stdio: 'ignore' });
                    } catch (e) {}
                    tunnelProcess = null;
                }
                setTimeout(() => iniciarTunelAutomatico(), 2000);
                return;
            }

            const matches = str.match(/https:\/\/[a-zA-Z0-9-]+\.trycloudflare\.com/g);
            if (matches) {
                for (const m of matches) {
                    if (!m.includes('api.trycloudflare.com')) {
                        currentTunnelUrl = m;
                        try { fs.writeFileSync(TUNNEL_URL_FILE, currentTunnelUrl, 'utf8'); } catch (e) {}
                        console.log('>>> [Cloudflare Tunnel] Enlace público en vivo 100% operativo:', currentTunnelUrl);
                        break;
                    }
                }
            }
        };

        tunnelProcess.stdout.on('data', checkOutput);
        tunnelProcess.stderr.on('data', checkOutput);

        tunnelProcess.on('close', () => {
            tunnelFailCount++;
            console.log('[Cloudflare Tunnel] Proceso finalizado. Reintento en 5s...');
            tunnelProcess = null;
            setTimeout(() => iniciarTunelAutomatico(), 5000);
        });
    } catch (e) {
        console.warn('[Cloudflare] Error iniciando túnel automático:', e.message);
    }
}

// Monitor no destructivo: NUNCA matar el proceso cloudflared. Cloudflare es auto-recuperable.
setInterval(() => {
    const cloudflaredPath = path.join(__dirname, 'almacen-app', 'cloudflared.exe');
    if (!fs.existsSync(cloudflaredPath) || process.env.RENDER || process.env.IS_CLOUD) return;
    if (!tunnelProcess) {
        iniciarTunelAutomatico();
    }
}, 15000);

app.get('/api/version', (req, res) => {
    let htmlMtime = Date.now();
    try {
        const stats = fs.statSync(path.join(__dirname, 'index.html'));
        htmlMtime = stats.mtimeMs;
    } catch (e) {}

    const host = req.get('host') || '';
    const isLocalhost = host.includes('localhost') || host.includes('127.0.0.1') || host.startsWith('192.168.');
    let effectiveUrl = currentTunnelUrl;
    if (!isLocalhost && host) {
        const proto = req.headers['x-forwarded-proto'] || req.protocol || 'https';
        effectiveUrl = `${proto}://${host}`;
    }

    const unidades = getUnidades() || [];
    const jornada = getJornadaActual() || {};

    res.json({
        success: true,
        version: '1.4.0',
        lastUpdated: htmlMtime,
        dataTimestamp: ultimaMarcaTiempoDatos,
        unitsCount: unidades.length,
        activeTunnel: effectiveUrl,
        jornadaFecha: jornada.fecha || null
    });
});

app.get('/api/tunnel-status', async (req, res) => {
    if (fs.existsSync(TUNNEL_URL_FILE)) {
        try {
            const saved = fs.readFileSync(TUNNEL_URL_FILE, 'utf8').trim();
            if (saved && saved.startsWith('http')) currentTunnelUrl = saved;
        } catch (e) {}
    }

    const host = req.get('host') || '';
    const isLocalhost = host.includes('localhost') || host.includes('127.0.0.1') || host.startsWith('192.168.');
    let effectiveUrl = currentTunnelUrl;
    if (!isLocalhost && host) {
        const proto = req.headers['x-forwarded-proto'] || req.protocol || 'https';
        effectiveUrl = `${proto}://${host}`;
    }

    if (!effectiveUrl && !tunnelProcess) {
        iniciarTunelAutomatico();
    }

    const currentIp = getLocalIp();
    res.json({
        active: Boolean(effectiveUrl),
        url: effectiveUrl || null,
        localUrl: `http://${currentIp}:${PORT}`,
        localHostUrl: `http://localhost:${PORT}`
    });
});

app.post('/api/iniciar-tunel', async (req, res) => {
    if (currentTunnelUrl) {
        return res.json({ success: true, url: currentTunnelUrl, alreadyRunning: true });
    }

    iniciarTunelAutomatico();

    let attempts = 0;
    const interval = setInterval(() => {
        attempts++;
        if (currentTunnelUrl) {
            clearInterval(interval);
            return res.json({ success: true, url: currentTunnelUrl });
        }
        if (attempts >= 25) { // 10 segundos
            clearInterval(interval);
            return res.json({ 
                success: true, 
                url: currentTunnelUrl || `http://${getLocalIp()}:${PORT}`, 
                note: 'Túnel en proceso, usando IP local como respaldo' 
            });
        }
    }, 400);
});

// --- 20. ENDPOINT RÁPIDO PARA DRAG & DROP Y EDICIÓN INMEDIATA ---
app.post('/api/unidad/:id/mover', (req, res) => {
    const unitId = String(req.params.id).trim();
    const { nuevoEstado } = req.body;
    if (!['considerar', 'mayor', 'proceso', 'listo'].includes(nuevoEstado)) {
        return res.status(400).json({ error: 'Estado inválido' });
    }

    let unidades = getUnidades();
    let unidad = unidades.find(u => String(u.id) === unitId);
    if (!unidad) return res.status(404).json({ error: 'Unidad no encontrada' });

    const anterior = unidad.estado;
    unidad.estado = nuevoEstado;
    unidad.estadoModificadoManualmente = true;

    const nombres = {
        considerar: 'A Considerar / En Evaluación',
        mayor: 'Reparación Mayor ($$)',
        proceso: 'En Reparación Activa',
        listo: 'Listas para Entrega'
    };
    const desc = `Movimiento rápido: De '${nombres[anterior] || anterior}' a '${nombres[nuevoEstado]}'`;

    if (!unidad.historial) unidad.historial = [];
    unidad.historial.unshift({
        fecha: new Date().toLocaleString('es-MX'),
        evento: desc,
        tipo: 'movimiento_kanban',
        autor: req.body.autor || 'Jefe de Taller'
    });

    saveUnidades(unidades);
    registrarMovimientoJornada(unitId, 'movimiento_kanban', desc, req.body.autor || 'Jefe de Taller');

    res.json({ success: true, unidad });
});

// --- 21. REGISTRO RÁPIDO DE AVANCES EN TARJETAS ---
app.post('/api/unidad/:id/avance', (req, res) => {
    const unitId = String(req.params.id).trim();
    const { texto, mecanico, porcentaje, autor } = req.body;
    if (!texto || !texto.trim()) {
        return res.status(400).json({ error: 'El texto del avance es requerido' });
    }

    let unidades = getUnidades();
    let unidad = unidades.find(u => String(u.id) === unitId);
    if (!unidad) return res.status(404).json({ error: 'Unidad no encontrada' });

    if (!Array.isArray(unidad.avances)) unidad.avances = [];
    const now = new Date();
    const hora = now.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' });
    const fecha = now.toLocaleDateString('es-MX');

    const nuevoAvance = {
        id: `av_${Date.now()}_${Math.random().toString(36).substr(2, 4)}`,
        texto: texto.trim(),
        mecanico: (mecanico || '').trim(),
        porcentaje: Number(porcentaje) || null,
        autor: autor || 'Jefe de Taller',
        fecha,
        hora,
        timestamp: Date.now()
    };

    unidad.avances.unshift(nuevoAvance);
    unidad.ultimo_avance = {
        texto: nuevoAvance.texto,
        mecanico: nuevoAvance.mecanico,
        hora: nuevoAvance.hora,
        fecha: nuevoAvance.fecha,
        porcentaje: nuevoAvance.porcentaje
    };

    if (!Array.isArray(unidad.historial)) unidad.historial = [];
    unidad.historial.unshift({
        fecha: now.toLocaleString('es-MX'),
        evento: `Avance registrado: "${nuevoAvance.texto}"${nuevoAvance.mecanico ? ' (Mec: ' + nuevoAvance.mecanico + ')' : ''}`,
        tipo: 'avance_reparacion',
        autor: nuevoAvance.autor
    });

    saveUnidades(unidades);
    registrarMovimientoJornada(unitId, 'avance_reparacion', `Avance: ${nuevoAvance.texto}`, nuevoAvance.autor);

    res.json({ success: true, unidad, avance: nuevoAvance });
});

// --- 22. VINCULACIÓN DE FOLIOS DE ALMACÉN Y COSTOS POR UNIDAD ---
app.get('/api/almacen/folio/:folio', (req, res) => {
    const folioBuscado = String(req.params.folio || '').trim();
    if (!folioBuscado) return res.status(400).json({ error: 'Folio requerido' });

    const almacenData = getAlmacenData();
    const movimientos = (almacenData.movimientos || []).filter(m => 
        String(m.folio || '').trim() === folioBuscado && (m.tipo === 'S' || !m.tipo)
    );

    const catalog = almacenData.catalog || {};
    let costoTotal = 0;
    const items = movimientos.map(m => {
        const itemCat = catalog[m.id_pieza] || {};
        const pu = Number(m.precio_unitario || itemCat.precio_unitario || 0);
        const cant = Number(m.cantidad) || 1;
        const total = pu * cant;
        costoTotal += total;
        return {
            ...m,
            descripcion: m.descripcion || itemCat.descripcion || 'Refacción sin descripción',
            numero_parte: m.numero_parte || itemCat.numero_parte || '',
            marca: m.marca || itemCat.marca || '',
            precio_unitario: pu,
            costo_total: total
        };
    });

    res.json({
        success: true,
        folio: folioBuscado,
        count: items.length,
        costoTotal,
        items,
        unidad_camion: items[0]?.unidad_camion || null,
        solicitante: items[0]?.solicitante || null,
        fecha: items[0]?.fecha || null
    });
});

app.post('/api/unidad/:id/vincular-folio', (req, res) => {
    const unitId = String(req.params.id).trim();
    const { folio } = req.body;
    if (!folio) return res.status(400).json({ error: 'Folio requerido' });

    let unidades = getUnidades();
    let unidad = unidades.find(u => String(u.id) === unitId);
    if (!unidad) return res.status(404).json({ error: 'Unidad no encontrada' });

    const almacenData = getAlmacenData();
    const movimientos = (almacenData.movimientos || []).filter(m => 
        String(m.folio || '').trim() === String(folio).trim()
    );

    const catalog = almacenData.catalog || {};
    let costoRefacciones = 0;
    const piezasSurtidas = movimientos.map((m, idx) => {
        const itemCat = catalog[m.id_pieza] || {};
        const pu = Number(m.precio_unitario || itemCat.precio_unitario || 0);
        const cant = Number(m.cantidad) || 1;
        const total = pu * cant;
        costoRefacciones += total;
        return {
            id: `pz_alm_${m.id || Date.now()}_${idx}`,
            id_pieza: m.id_pieza,
            descripcion: m.descripcion || itemCat.descripcion || 'Refacción',
            numero_parte: m.numero_parte || itemCat.numero_parte || '',
            marca: m.marca || itemCat.marca || '',
            cantidad: cant,
            unidad_medida: m.unidad || 'PZ',
            precio_unitario: pu,
            costo_total: total,
            folio_almacen: folio,
            estado: 'RECIBIDA_TALLER',
            fecha_recepcion: m.fecha || new Date().toISOString()
        };
    });

    unidad.folio_almacen = String(folio).trim();
    unidad.costo_refacciones = costoRefacciones;
    if (!Array.isArray(unidad.piezas)) unidad.piezas = [];

    // Agregar piezas sin duplicar por id_pieza y folio
    piezasSurtidas.forEach(np => {
        const yaExiste = unidad.piezas.some(p => p.id_pieza === np.id_pieza && p.folio_almacen === folio);
        if (!yaExiste) unidad.piezas.push(np);
    });

    // Actualizar costo total de la unidad
    unidad.costo = (Number(unidad.costo_mano_obra) || 0) + costoRefacciones;

    const desc = `Vinculado Folio Almacén #${folio}: ${piezasSurtidas.length} refacciones ($${costoRefacciones.toLocaleString()} MXN)`;
    if (!Array.isArray(unidad.historial)) unidad.historial = [];
    unidad.historial.unshift({
        fecha: new Date().toLocaleString('es-MX'),
        evento: desc,
        tipo: 'vinculacion_almacen',
        autor: req.body.autor || 'Encargado de Taller'
    });

    saveUnidades(unidades);
    registrarMovimientoJornada(unitId, 'vinculacion_almacen', desc, req.body.autor || 'Encargado de Taller');

    res.json({ success: true, unidad, itemsVinculados: piezasSurtidas.length, costoRefacciones });
});

// --- GENERAR O SOLICITAR NUEVO FOLIO DE ALMACÉN DESDE TALLER ---
app.post('/api/unidad/:id/generar-folio-almacen', (req, res) => {
    const unitId = String(req.params.id).trim();
    const { folio, motivo, refacciones, solicitante } = req.body;
    let unidades = getUnidades();
    let unidad = unidades.find(u => String(u.id) === unitId);
    if (!unidad) return res.status(404).json({ error: 'Unidad no encontrada' });

    const almacenData = getAlmacenData();
    let nuevoFolio = String(folio || '').trim();
    if (!nuevoFolio) {
        // Generar consecutivo inteligente en rango 4000
        let maxF = 4075;
        (almacenData.movimientos || []).forEach(m => {
            const n = parseInt(m.folio, 10);
            if (!isNaN(n) && n >= 4000 && n < 9999 && n > maxF) maxF = n;
        });
        (almacenData.folios_taller || []).forEach(f => {
            const n = parseInt(f.folio, 10);
            if (!isNaN(n) && n >= 4000 && n < 9999 && n > maxF) maxF = n;
        });
        nuevoFolio = String(maxF + 1);
    }

    unidad.folio_almacen = nuevoFolio;
    const desc = `Generado Folio de Almacén #${nuevoFolio}: ${motivo || 'Requisición de refacciones para reparación'}. Refacciones solicitadas: ${refacciones || 'Por surtir en almacén'}`;
    
    if (!Array.isArray(unidad.historial)) unidad.historial = [];
    unidad.historial.unshift({
        fecha: new Date().toLocaleDateString('es-MX') + ' ' + new Date().toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' }),
        evento: desc,
        tipo: 'generacion_folio_almacen',
        autor: solicitante || 'Encargado de Taller'
    });

    // Registrar en folios_taller de almacen_datos.json
    if (!Array.isArray(almacenData.folios_taller)) almacenData.folios_taller = [];
    const yaExisteIndex = almacenData.folios_taller.findIndex(f => String(f.folio).trim() === nuevoFolio);
    const regFolio = {
        folio: nuevoFolio,
        unidad: unitId,
        solicitante: solicitante || 'Encargado de Taller',
        motivo: motivo || unidad.reporte || 'Reparación en taller',
        refacciones_solicitadas: refacciones || 'Evaluación técnica',
        fecha: new Date().toISOString().slice(0, 10),
        hora: new Date().toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' }),
        estado_taller: unidad.estado || 'proceso',
        estatus: 'ABIERTO',
        creado_desde: 'Taller Pro CESS',
        timestamp: new Date().toISOString()
    };

    if (yaExisteIndex !== -1) {
        almacenData.folios_taller[yaExisteIndex] = regFolio;
    } else {
        almacenData.folios_taller.unshift(regFolio);
    }
    saveAlmacenData(almacenData);
    saveUnidades(unidades, solicitante || 'Encargado de Taller');
    registrarMovimientoJornada(unitId, 'generacion_folio_almacen', desc, solicitante || 'Encargado de Taller');

    res.json({
        success: true,
        folio: nuevoFolio,
        unidad,
        mensaje: `Folio #${nuevoFolio} registrado en Almacén CESS y vinculado a Unidad ${unitId}`
    });
});

// --- LISTAR TODOS LOS FOLIOS DE TALLER ACTIVOS PARA ALMACÉN ---
app.get('/api/almacen/folios-taller', (req, res) => {
    const almacenData = getAlmacenData();
    const movimientos = almacenData.movimientos || [];
    const catalog = almacenData.catalog || {};
    const unidades = getUnidades();
    const foliosRegistrados = almacenData.folios_taller || [];

    // Mapear cada unidad con folio vinculado
    const foliosMap = new Map();

    // 1. Agregar folios que están en el tablero de taller
    unidades.forEach(u => {
        if (!u.folio_almacen) return;
        const fol = String(u.folio_almacen).trim();
        foliosMap.set(fol, {
            folio: fol,
            unidad: u.id,
            estado_taller: u.estado || 'proceso',
            defecto: u.defecto || u.reporte || 'Reparación en taller',
            solicitante: 'Encargado de Taller',
            fecha_apertura: u.fecha_ingreso || new Date().toISOString().slice(0, 10),
            piezas: [],
            costo_acumulado: 0
        });
    });

    // 2. Agregar folios registrados en almacen_datos.json
    foliosRegistrados.forEach(fr => {
        const fol = String(fr.folio).trim();
        if (!foliosMap.has(fol)) {
            foliosMap.set(fol, {
                folio: fol,
                unidad: fr.unidad || '--',
                estado_taller: fr.estado_taller || 'proceso',
                defecto: fr.motivo || fr.refacciones_solicitadas || 'Reparación en taller',
                solicitante: fr.solicitante || 'Encargado de Taller',
                fecha_apertura: fr.fecha || new Date().toISOString().slice(0, 10),
                piezas: [],
                costo_acumulado: 0
            });
        }
    });

    // 3. Asociar salidas y calcular costos
    const foliosLista = Array.from(foliosMap.values()).map(f => {
        const movs = movimientos.filter(m => String(m.folio).trim() === f.folio && (m.tipo === 'S' || !m.tipo));
        let costoTotal = 0;
        const piezas = movs.map(m => {
            const itemCat = catalog[m.id_pieza] || {};
            const pu = Number(m.precio_unitario !== undefined ? m.precio_unitario : (itemCat.precio_unitario || 0));
            const cant = Number(m.cantidad) || 1;
            const sub = pu * cant;
            costoTotal += sub;
            return {
                id_movimiento: m.id,
                id_pieza: m.id_pieza,
                numero_parte: m.numero_parte || itemCat.numero_parte || '',
                descripcion: m.descripcion || itemCat.descripcion || 'Refacción',
                cantidad: cant,
                precio_unitario: pu,
                subtotal: sub,
                fecha: m.fecha || '',
                hora: m.hora || ''
            };
        });

        f.piezas = piezas;
        f.costo_acumulado = costoTotal;
        f.estatus = piezas.length > 0 ? 'CON_REFACCIONES_SURTIDAS' : 'PENDIENTE_SURTIR';
        return f;
    });

    res.json({ success: true, count: foliosLista.length, folios: foliosLista });
});

// --- SURTIR REFACCIÓN DIRECTAMENTE A UN FOLIO DE TALLER ---
app.post('/api/almacen/folio/:folio/surtir-pieza', (req, res) => {
    const folio = String(req.params.folio).trim();
    const { id_pieza, cantidad, solicitante, notas } = req.body;
    const cant = Number(cantidad) || 1;

    const almacenData = getAlmacenData();
    const item = (almacenData.catalog || {})[id_pieza];
    if (!item) {
        return res.status(404).json({ error: 'Refacción no encontrada en el catálogo de almacén' });
    }

    // Descontar inventario
    item.stock_actual = (Number(item.stock_actual) || 0) - cant;

    const now = new Date();
    const fecha = now.toISOString().slice(0, 10);
    const hora = now.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' });

    // Determinar unidad asociada
    const unidades = getUnidades();
    const uMatch = unidades.find(u => String(u.folio_almacen).trim() === folio);
    const unidadCamion = uMatch ? uMatch.id : (req.body.unidad_camion || '');

    const mov = {
        id: `mov_${Date.now()}_${Math.round(Math.random() * 1000)}`,
        tipo: 'S',
        fecha,
        hora,
        id_pieza: item.id || id_pieza,
        numero_parte: item.numero_parte || '',
        descripcion: item.descripcion || '',
        marca: item.marca || '',
        almacen: item.almacen || 'Almacén de Refacciones',
        almacen_id: item.almacen_id || 'refacciones',
        ubicacion: `${item.rack || 'Rack 1'} / ${item.nivel || 'Nivel 1'}`,
        cantidad: cant,
        unidad_camion: String(unidadCamion).toUpperCase(),
        folio: folio,
        solicitante: solicitante || (uMatch ? `Taller Unidad ${uMatch.id}` : 'Taller Pro'),
        precio_unitario: Number(item.precio_unitario || 0),
        notas: notas || 'Surtido desde Centro de Vinculación Taller-Almacén',
        timestamp: now.toISOString()
    };

    if (!Array.isArray(almacenData.movimientos)) almacenData.movimientos = [];
    almacenData.movimientos.unshift(mov);
    saveAlmacenData(almacenData);

    // Sincronizar inmediatamente con Taller
    sincronizarFoliosConAlmacen('Despacho Inmediato Almacén');

    res.json({
        success: true,
        movimiento: mov,
        mensaje: `${cant} PZ de "${item.descripcion}" surtidas al Folio #${folio} de Unidad ${unidadCamion || ''}`
    });
});

// --- MODIFICAR PIEZA O CANTIDAD EN FOLIO DE ALMACÉN ---
app.post('/api/almacen/folio/:folio/modificar-pieza', (req, res) => {
    const folio = String(req.params.folio).trim();
    const { id_movimiento, nueva_cantidad, nuevo_precio_unitario } = req.body;
    const almacenData = getAlmacenData();
    const mov = (almacenData.movimientos || []).find(m => String(m.id) === String(id_movimiento));
    if (!mov) {
        return res.status(404).json({ error: 'Movimiento no encontrado en la bitácora de almacén' });
    }

    const cantAnterior = Number(mov.cantidad) || 0;
    const cantNueva = Number(nueva_cantidad);
    const item = (almacenData.catalog || {})[mov.id_pieza];

    if (cantNueva <= 0) {
        // Devolución completa de la refacción
        if (item) item.stock_actual = (Number(item.stock_actual) || 0) + cantAnterior;
        almacenData.movimientos = almacenData.movimientos.filter(m => String(m.id) !== String(id_movimiento));
    } else {
        const dif = cantNueva - cantAnterior;
        if (item) item.stock_actual = (Number(item.stock_actual) || 0) - dif;
        mov.cantidad = cantNueva;
        if (nuevo_precio_unitario !== undefined) {
            mov.precio_unitario = Number(nuevo_precio_unitario);
        }
    }

    saveAlmacenData(almacenData);
    sincronizarFoliosConAlmacen('Ajuste de Pieza en Almacén');

    res.json({
        success: true,
        mensaje: cantNueva <= 0 
            ? 'Pieza devuelta al stock de almacén y retirada de la unidad'
            : 'Cantidad y costo actualizados en Almacén y reflejados en Taller'
    });
});

// --- 23. MÓDULO DE MANTENIMIENTOS PREVENTIVOS (VINCULADO A SISTEMA DE CONTROL CESS COMBUSTIBLE) ---
app.get('/api/preventivos', (req, res) => {
    const cacheData = getPreventivosData();
    const list = cacheData.preventivos || [];
    const stats = cacheData.stats || {
        total: list.length,
        vigentes: list.filter(p => (p.estatus || '').toUpperCase() === 'VIGENTE').length,
        vencidos: list.filter(p => (p.estatus || '').toUpperCase() === 'VENCIDO').length,
        proximos: list.filter(p => (p.estatus || '').toUpperCase() === 'PRÓXIMO' || (p.estatus || '').toUpperCase() === 'PROXIMO').length
    };
    
    const unidadParam = req.query.unidad ? String(req.query.unidad).trim() : null;
    if (unidadParam) {
        const item = list.find(p => String(p.unidad).trim() === unidadParam);
        return res.json({ success: true, item: item || null });
    }

    res.json({
        success: true,
        source: 'Control Oficial de Flota CESS',
        data: list,
        preventivos: list,
        stats,
        resumen: stats,
        count: list.length
    });
});

app.get('/api/preventivos/bitacora', (req, res) => {
    let bitacora = getBitacoraPreventivos();
    const unidadParam = req.query.unidad ? String(req.query.unidad).trim() : null;
    if (unidadParam) {
        bitacora = bitacora.filter(b => String(b.unidad).trim() === unidadParam);
    }
    res.json({
        success: true,
        source: 'Bitácora Oficial de Mantenimientos CESS',
        data: bitacora,
        bitacora: bitacora,
        count: bitacora.length
    });
});

// --- GESTOR DE TÍTULOS DE COLUMNAS DEL TABLERO ---
app.get('/api/columnas/titulos', (req, res) => {
    res.json({ success: true, titulos: getTitulosColumnas() });
});

app.post('/api/columnas/titulos', (req, res) => {
    const titulos = req.body;
    if (!titulos || typeof titulos !== 'object') {
        return res.status(400).json({ success: false, error: 'Datos de títulos inválidos' });
    }
    const current = getTitulosColumnas();
    const updated = { ...current, ...titulos };
    saveTitulosColumnas(updated);
    broadcastSSE('columnas_updated', { titulos: updated });
    res.json({ success: true, titulos: updated });
});

// --- REGISTRO DE NUEVA DETECCIÓN / UNIDAD A CONSIDERAR ---
app.post('/api/unidad/deteccion', (req, res) => {
    const { unidad, sucursal, situacion, refacciones, trabajos, prioridad, reportado_por, costo_estimado } = req.body;
    if (!unidad) {
        return res.status(400).json({ success: false, error: 'Número de unidad requerido' });
    }

    const unidades = getUnidades();
    let u = unidades.find(x => String(x.id).trim() === String(unidad).trim());
    const hora = new Date().toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' });
    const fecha = new Date().toLocaleDateString('es-MX');

    const descDeteccion = `[DETECCIÓN ${prioridad || 'MEDIA'}] ${situacion || ''}\nRefacciones: ${refacciones || 'Por evaluar'}\nTrabajos: ${trabajos || 'Revisión en patio'}\nReportó: ${reportado_por || 'Supervisor'}`;

    if (u) {
        u.estado = 'considerar';
        u.sucursal = sucursal || u.sucursal || 'Apodaca';
        u.reporte = descDeteccion;
        u.defecto = situacion || u.defecto;
        u.situacion_detectada = situacion;
        u.refacciones_requeridas = refacciones;
        u.trabajos_requeridos = trabajos;
        u.prioridad = prioridad || 'MEDIA';
        u.reportado_por = reportado_por || 'Supervisor';
        if (costo_estimado) u.costo = Number(costo_estimado);
        if (!Array.isArray(u.historial)) u.historial = [];
        u.historial.push({
            fecha: `${fecha} ${hora}`,
            accion: `Nueva Detección registrada para evaluación: ${situacion}`,
            usuario: reportado_por || 'Supervisor'
        });
    } else {
        u = {
            id: String(unidad).trim(),
            sucursal: sucursal || 'Apodaca',
            estado: 'considerar',
            defecto: situacion || 'Situación detectada para evaluación',
            reporte: descDeteccion,
            situacion_detectada: situacion,
            refacciones_requeridas: refacciones,
            trabajos_requeridos: trabajos,
            prioridad: prioridad || 'MEDIA',
            reportado_por: reportado_por || 'Supervisor',
            costo: Number(costo_estimado) || 0,
            piezas: [],
            avances: [],
            historial: [{
                fecha: `${fecha} ${hora}`,
                accion: `Ingreso a Considerar / Detecciones: ${situacion}`,
                usuario: reportado_por || 'Supervisor'
            }]
        };
        unidades.push(u);
    }

    saveUnidades(unidades);
    broadcastSSE('units_updated', { units: unidades, autor: reportado_por || 'Supervisor' });
    res.json({ success: true, unidad: u });
});

// --- CONSULTA DETALLADA DE FOLIO DE ALMACÉN ---
app.get('/api/almacen/folio/:folio/detalle', (req, res) => {
    const folio = String(req.params.folio).trim();
    const data = getAlmacenData();
    const movimientos = data.movimientos || data.salidas || [];

    const matches = movimientos.filter(m => String(m.folio).trim() === folio);
    if (matches.length === 0) {
        return res.status(404).json({ success: false, error: 'Folio no encontrado en Almacén CESS' });
    }

    const items = matches.map(m => ({
        id_item: m.id || m.numero_parte,
        numero_parte: m.numero_parte || '--',
        descripcion: m.descripcion || m.concepto || 'Refacción',
        cantidad: Number(m.cantidad) || 1,
        precio_unitario: Number(m.precio_unitario || m.costo_unitario) || 0,
        total_costo: (Number(m.cantidad) || 1) * (Number(m.precio_unitario || m.costo_unitario) || 0),
        fecha: m.fecha || m.fecha_salida || '--',
        solicitante: m.solicitante || m.entregado_a || '--',
        unidad: m.unidad_camion || m.unidad || '--'
    }));

    const costoTotal = items.reduce((acc, it) => acc + it.total_costo, 0);

    res.json({
        success: true,
        folio,
        unidad: matches[0].unidad_camion || matches[0].unidad || '--',
        solicitante: matches[0].solicitante || matches[0].entregado_a || '--',
        fecha: matches[0].fecha || matches[0].fecha_salida || '--',
        items,
        costo_total: costoTotal
    });
});

// --- 24. MÓDULO DE RESCATES CARRETEROS (AUXILIO EN RUTA) ---
app.get('/api/rescates', (req, res) => {
    const list = getRescates();
    const activos = list.filter(r => r.estado !== 'CONCLUIDO');
    const concluidos = list.filter(r => r.estado === 'CONCLUIDO');
    const costoTotal = list.reduce((acc, r) => acc + (Number(r.costo_real || r.costo_estimado) || 0), 0);

    res.json({
        success: true,
        stats: {
            total: list.length,
            activos: activos.length,
            concluidos: concluidos.length,
            costoTotal
        },
        rescates: list
    });
});

app.post('/api/rescates', (req, res) => {
    const data = req.body;
    if (!data.unidad || !data.carretera) {
        return res.status(400).json({ error: 'Unidad y Carretera son campos obligatorios' });
    }

    const list = getRescates();
    const now = new Date();
    const id = data.id || `rc_${Date.now()}`;
    const folio = data.folio || `RC-${Math.floor(1000 + Math.random() * 9000)}`;

    const nuevoRescate = {
        id,
        folio,
        unidad: String(data.unidad).trim(),
        operador: data.operador || 'Operador CESS',
        telefono_operador: data.telefono_operador || '',
        fecha: data.fecha || now.toLocaleDateString('es-MX'),
        hora_reporte: data.hora_reporte || now.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' }),
        carretera: data.carretera,
        kilometro: data.kilometro || '',
        tramo: data.tramo || '',
        sentido: data.sentido || 'Norte a Sur',
        falla_reportada: data.falla_reportada || 'Falla mecánica en carretera',
        tipo_falla: data.tipo_falla || 'Mecánica',
        estado: data.estado || 'EN_CAMINO',
        mecanico_asignado: data.mecanico_asignado || 'Auxilio Móvil CESS',
        proveedor_externo: data.proveedor_externo || '',
        costo_estimado: Number(data.costo_estimado) || 0,
        costo_real: Number(data.costo_real) || 0,
        tiempo_llegada: data.tiempo_llegada || 'En trayecto',
        tiempo_resolucion: data.tiempo_resolucion || 'En proceso',
        hora_conclusion: data.hora_conclusion || null,
        notas: data.notas || '',
        sucursal: data.sucursal || 'Apodaca'
    };

    const idx = list.findIndex(r => r.id === id);
    if (idx >= 0) {
        list[idx] = { ...list[idx], ...nuevoRescate };
    } else {
        list.unshift(nuevoRescate);
    }
    saveRescates(list);

    // Si la unidad está en el tablero taller, marcarla como rescate carretero
    let unidades = getUnidades();
    let uTaller = unidades.find(u => String(u.id) === String(nuevoRescate.unidad));
    if (uTaller) {
        uTaller.es_rescate = nuevoRescate.estado !== 'CONCLUIDO';
        uTaller.rescate_folio = nuevoRescate.folio;
        uTaller.rescate_tramo = `${nuevoRescate.carretera} ${nuevoRescate.kilometro}`.trim();
        saveUnidades(unidades);
    }

    res.json({ success: true, rescate: nuevoRescate });
});

app.post('/api/rescates/:id/concluir', (req, res) => {
    const { id } = req.params;
    const { costo_real, notas_conclusion, tiempo_resolucion } = req.body;

    const list = getRescates();
    const rescate = list.find(r => r.id === id);
    if (!rescate) return res.status(404).json({ error: 'Rescate no encontrado' });

    rescate.estado = 'CONCLUIDO';
    rescate.costo_real = Number(costo_real) || rescate.costo_estimado;
    rescate.hora_conclusion = new Date().toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' });
    if (tiempo_resolucion) rescate.tiempo_resolucion = tiempo_resolucion;
    if (notas_conclusion) rescate.notas = `${rescate.notas || ''}\n[Conclusión]: ${notas_conclusion}`.trim();

    saveRescates(list);

    // Actualizar estado en unidad si existe
    let unidades = getUnidades();
    let uTaller = unidades.find(u => String(u.id) === String(rescate.unidad));
    if (uTaller) {
        uTaller.es_rescate = false;
        saveUnidades(unidades);
    }

    res.json({ success: true, rescate });
});

// --- 25. MÓDULO DE CAJAS Y REMOLQUES CESS ---
app.get('/api/remolques', (req, res) => {
    const list = getRemolques();
    const stats = {
        total: list.length,
        considerar: list.filter(r => r.estado === 'considerar').length,
        mayor: list.filter(r => r.estado === 'mayor').length,
        proceso: list.filter(r => r.estado === 'proceso').length,
        listo: list.filter(r => r.estado === 'listo').length,
        costoTotal: list.reduce((acc, r) => acc + (Number(r.costo) || 0), 0)
    };
    res.json({ success: true, stats, remolques: list, count: list.length });
});

app.post('/api/remolques', (req, res) => {
    const remolques = req.body;
    if (!Array.isArray(remolques)) {
        return res.status(400).json({ error: 'Se esperaba un arreglo de remolques' });
    }
    saveRemolques(remolques);
    res.json({ success: true, count: remolques.length });
});

app.post('/api/remolque/:id', (req, res) => {
    const remolqueId = String(req.params.id).trim();
    const data = req.body;
    let list = getRemolques();
    let idx = list.findIndex(r => String(r.id).trim().toUpperCase() === remolqueId.toUpperCase());

    const item = {
        id: remolqueId,
        tipo: data.tipo || "Caja Seca 53'",
        marca: data.marca || '',
        anio: data.anio || '',
        placas: data.placas || '',
        estado: data.estado || 'proceso',
        sistema: data.sistema || 'General',
        defecto: data.defecto || data.reporte || 'Mantenimiento preventivo / correctivo',
        reporte: data.reporte || data.defecto || '',
        mecanico: data.mecanico || 'Sin asignar',
        costo: Number(data.costo) || 0,
        sucursal: data.sucursal || 'Apodaca',
        folio_almacen: data.folio_almacen || '',
        fechaIngreso: data.fechaIngreso || new Date().toLocaleDateString('es-MX')
    };

    if (idx >= 0) {
        list[idx] = { ...list[idx], ...item };
    } else {
        list.unshift(item);
    }
    saveRemolques(list);
    res.json({ success: true, remolque: item });
});

app.post('/api/remolque/:id/estado', (req, res) => {
    const remolqueId = String(req.params.id).trim();
    const { nuevoEstado } = req.body;
    let list = getRemolques();
    let item = list.find(r => String(r.id).trim().toUpperCase() === remolqueId.toUpperCase());
    if (!item) return res.status(404).json({ error: 'Remolque no encontrado' });

    item.estado = nuevoEstado;
    saveRemolques(list);
    res.json({ success: true, remolque: item });
});

app.delete('/api/remolque/:id', (req, res) => {
    const remolqueId = String(req.params.id).trim();
    let list = getRemolques();
    const nuevaLista = list.filter(r => String(r.id).trim().toUpperCase() !== remolqueId.toUpperCase());
    saveRemolques(nuevaLista);
    res.json({ success: true, count: nuevaLista.length });
});

app.get('/api/exportar-excel-remolques', (req, res) => {
    try {
        const remolques = getRemolques();
        const fechaHoy = getFechaActualISO();
        const rows = remolques.map(r => ({
            'No. Remolque / Caja': r.id,
            'Tipo de Remolque': r.tipo,
            'Marca / Fabricante': r.marca || '',
            'Año': r.anio || '',
            'Placas': r.placas || '',
            'Estatus en Taller': r.estado === 'mayor' ? 'Reparación Mayor / Hojalatería' : (r.estado === 'proceso' ? 'En Reparación Activa' : (r.estado === 'considerar' ? 'A Considerar / Daño' : 'Lista para Carga')),
            'Sistema Afectado': r.sistema || 'General',
            'Defecto / Trabajo Reportado': r.defecto || '',
            'Avances Técnicos / Notas': r.reporte || '',
            'Mecánico Asignado': r.mecanico || 'Sin asignar',
            'Folio Almacén': r.folio_almacen ? `#${r.folio_almacen}` : 'Sin folio',
            'Costo Estimado (MXN)': Number(r.costo) || 0,
            'Patio / Ubicación': r.sucursal || 'Apodaca',
            'Fecha Ingreso': r.fechaIngreso || ''
        }));

        const wb = xlsx.utils.book_new();
        const ws = xlsx.utils.json_to_sheet(rows);
        ws['!cols'] = [
            { wch: 20 }, { wch: 24 }, { wch: 18 }, { wch: 10 }, { wch: 14 },
            { wch: 30 }, { wch: 22 }, { wch: 38 }, { wch: 42 }, { wch: 20 },
            { wch: 16 }, { wch: 20 }, { wch: 18 }, { wch: 16 }
        ];
        xlsx.utils.book_append_sheet(wb, ws, 'Remolques CESS');
        const buf = xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="Cajas_y_Remolques_CESS_${fechaHoy}.xlsx"`);
        res.send(buf);
    } catch(err) {
        res.status(500).json({ error: 'Error exportando Excel remolques: ' + err.message });
    }
});

// --- Servir Aplicación SPA Frontend con Anti-Cache para HTML y Service Worker ---
app.use((req, res, next) => {
    if (req.url === '/' || req.url.endsWith('.html') || req.url.includes('sw.js')) {
        res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
        res.setHeader('Pragma', 'no-cache');
        res.setHeader('Expires', '0');
    }
    next();
});

app.use(express.static(__dirname, {
    setHeaders: (res, filePath) => {
        if (filePath.endsWith('.html') || filePath.endsWith('sw.js')) {
            res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
            res.setHeader('Pragma', 'no-cache');
            res.setHeader('Expires', '0');
        }
    }
}));

// Fallback para SPA HTML5
app.use((req, res) => {
    res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('Expires', '0');
    res.sendFile(path.join(__dirname, 'index.html'));
});

// Obtener IP local de la máquina
function getLocalIp() {
    const ifaces = os.networkInterfaces();
    for (let dev in ifaces) {
        for (let details of ifaces[dev]) {
            if (details.family === 'IPv4' && !details.internal) {
                return details.address;
            }
        }
    }
    return 'localhost';
}

const localIp = getLocalIp();

app.listen(PORT, '0.0.0.0', () => {
    console.log('================================================================');
    console.log('       SISTEMA SMART WORKSHOP - SPA INTELIGENTE CESS');
    console.log('================================================================');
    console.log(`- En esta computadora:   http://localhost:${PORT}`);
    console.log(`- En tu celular / patio: http://${localIp}:${PORT}`);
    console.log('================================================================');
    iniciarTunelAutomatico();
});
