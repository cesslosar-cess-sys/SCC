const { Pool } = require('pg');
const fs = require('fs');
const path = require('path');

// =========================================================================
// MÓDULO DE PERSISTENCIA EMPRESARIAL (ERP) - ADAPTADOR POSTGRESQL & DUAL-STORAGE
// Cubre: Taller, Cajas y Remolques, Almacén y Sistema de Diésel & Rendimientos
// =========================================================================

let pool = null;
let isPostgresAvailable = false;
let isInitialized = false;

// Configuración de conexión (compatible con Render DATABASE_URL y variables estándar)
const connectionString = process.env.DATABASE_URL || process.env.PG_CONNECTION_STRING || null;

if (connectionString) {
    try {
        const isLocalhost = connectionString.includes('localhost') || connectionString.includes('127.0.0.1');
        pool = new Pool({
            connectionString: connectionString,
            ssl: isLocalhost ? false : { rejectUnauthorized: false },
            max: 10,
            idleTimeoutMillis: 30000,
            connectionTimeoutMillis: 5000
        });

        pool.on('error', (err) => {
            console.error('[PostgreSQL Error en Pool]:', err.message);
        });

        isPostgresAvailable = true;
        console.log('[PostgreSQL] Conexión configurada con base de datos remota.');
    } catch (e) {
        console.warn('[PostgreSQL] Error al configurar Pool:', e.message);
        isPostgresAvailable = false;
    }
} else {
    console.log('[PostgreSQL] DATABASE_URL no detectada en variables de entorno.');
    console.log('[PostgreSQL] Operando en modo Híbrido: Archivos Locales + Memoria Transaccional con Respaldo Automático.');
}

// -------------------------------------------------------------------------
// 1. INICIALIZACIÓN DE ESQUEMA RELACIONAL (DDL SEGURO)
// -------------------------------------------------------------------------
async function inicializarEsquemaPostgres() {
    if (!isPostgresAvailable || !pool) return false;

    const client = await pool.connect();
    try {
        await client.query('BEGIN');

        // A) MÓDULO TALLER Y FLOTA
        await client.query(`
            CREATE TABLE IF NOT EXISTS taller_unidades (
                id VARCHAR(64) PRIMARY KEY,
                sucursal VARCHAR(64) DEFAULT 'Apodaca',
                estado VARCHAR(64) DEFAULT 'proceso',
                defecto TEXT,
                reporte TEXT,
                notas TEXT,
                costo NUMERIC DEFAULT 0,
                costo_refacciones NUMERIC DEFAULT 0,
                costo_preventivo NUMERIC DEFAULT 0,
                costo_correctivo NUMERIC DEFAULT 0,
                tipo_procedimiento VARCHAR(64) DEFAULT 'correctivo',
                folio_almacen VARCHAR(64),
                fecha_ingreso VARCHAR(64),
                archivada BOOLEAN DEFAULT false,
                avances JSONB DEFAULT '[]'::jsonb,
                piezas JSONB DEFAULT '[]'::jsonb,
                historial JSONB DEFAULT '[]'::jsonb,
                datos_extra JSONB DEFAULT '{}'::jsonb,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS taller_unidades_archivadas (
                id VARCHAR(64) PRIMARY KEY,
                sucursal VARCHAR(64),
                estado VARCHAR(64),
                defecto TEXT,
                reporte TEXT,
                notas TEXT,
                costo NUMERIC DEFAULT 0,
                costo_refacciones NUMERIC DEFAULT 0,
                folio_almacen VARCHAR(64),
                fecha_ingreso VARCHAR(64),
                fecha_salida VARCHAR(64),
                folio_orden VARCHAR(64),
                motivo_salida TEXT,
                avances JSONB DEFAULT '[]'::jsonb,
                piezas JSONB DEFAULT '[]'::jsonb,
                historial JSONB DEFAULT '[]'::jsonb,
                datos_extra JSONB DEFAULT '{}'::jsonb,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS taller_remolques (
                id VARCHAR(64) PRIMARY KEY,
                tipo VARCHAR(64) DEFAULT 'Caja Seca 53',
                marca VARCHAR(64),
                anio VARCHAR(16),
                placas VARCHAR(32),
                estado VARCHAR(64) DEFAULT 'proceso',
                sistema VARCHAR(64) DEFAULT 'General',
                defecto TEXT,
                reporte TEXT,
                mecanico VARCHAR(128) DEFAULT 'Sin asignar',
                costo NUMERIC DEFAULT 0,
                costo_refacciones NUMERIC DEFAULT 0,
                sucursal VARCHAR(64) DEFAULT 'Apodaca',
                folio_almacen VARCHAR(64),
                fecha VARCHAR(64),
                fecha_ingreso VARCHAR(64),
                piezas JSONB DEFAULT '[]'::jsonb,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS taller_rescates (
                id VARCHAR(64) PRIMARY KEY,
                folio VARCHAR(64),
                unidad VARCHAR(64),
                operador VARCHAR(128),
                telefono VARCHAR(64),
                carretera TEXT,
                kilometro VARCHAR(64),
                km VARCHAR(64),
                tramo TEXT,
                sentido VARCHAR(64),
                falla_reportada TEXT,
                tipo_falla VARCHAR(64),
                costo_estimado NUMERIC DEFAULT 0,
                costo_real NUMERIC DEFAULT 0,
                estado VARCHAR(64) DEFAULT 'EN_CAMINO',
                estatus VARCHAR(64) DEFAULT 'EN_CAMINO',
                mecanico_asignado VARCHAR(128),
                proveedor_externo VARCHAR(128),
                tiempo_llegada VARCHAR(64),
                tiempo_resolucion VARCHAR(64),
                hora_conclusion VARCHAR(64),
                notas TEXT,
                sucursal VARCHAR(64) DEFAULT 'Apodaca',
                fecha VARCHAR(64),
                hora_reporte VARCHAR(64),
                avances JSONB DEFAULT '[]'::jsonb,
                raw_data JSONB,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
            ALTER TABLE taller_rescates ADD COLUMN IF NOT EXISTS raw_data JSONB;

            CREATE TABLE IF NOT EXISTS taller_preventivos (
                unidad VARCHAR(64) PRIMARY KEY,
                division VARCHAR(64) DEFAULT 'RL',
                marca VARCHAR(64),
                tracto VARCHAR(64),
                km_actual NUMERIC DEFAULT 0,
                km_ant NUMERIC DEFAULT 0,
                fecha_realizado VARCHAR(64),
                proximo_mtto NUMERIC DEFAULT 0,
                km_restantes NUMERIC DEFAULT 0,
                estatus VARCHAR(64) DEFAULT 'VIGENTE',
                ubicacion TEXT,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS taller_bitacoras (
                id VARCHAR(64) PRIMARY KEY,
                fecha VARCHAR(32),
                hora VARCHAR(32),
                titulo TEXT,
                responsable VARCHAR(128),
                resumen JSONB DEFAULT '{}'::jsonb,
                notas_cierre TEXT,
                movimientos JSONB DEFAULT '[]'::jsonb,
                unidades JSONB DEFAULT '[]'::jsonb,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS taller_usuarios (
                id VARCHAR(64) PRIMARY KEY,
                nombre VARCHAR(128),
                rol VARCHAR(64),
                pin VARCHAR(32),
                permisos JSONB DEFAULT '{}'::jsonb,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
        `);

        // B) MÓDULO DE DIÉSEL Y RENDIMIENTOS (100% PROTEGIDO E INTEGRADO)
        await client.query(`
            CREATE TABLE IF NOT EXISTS diesel_registros_rendimiento (
                id VARCHAR(64) PRIMARY KEY,
                timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                fecha VARCHAR(64),
                unidad VARCHAR(64),
                operador VARCHAR(128),
                ruta VARCHAR(64),
                km_recorridos NUMERIC DEFAULT 0,
                km_total NUMERIC DEFAULT 0,
                rend_ecm NUMERIC DEFAULT 0,
                rend_fisico NUMERIC DEFAULT 0,
                ecm NUMERIC DEFAULT 0,
                real_total NUMERIC DEFAULT 0,
                real_apodaca NUMERIC DEFAULT 0,
                real_laredo NUMERIC DEFAULT 0,
                real_foraneo NUMERIC DEFAULT 0,
                urea NUMERIC DEFAULT 0,
                rend_urea NUMERIC DEFAULT 0,
                litros_ralenti NUMERIC DEFAULT 0,
                costo_diesel NUMERIC DEFAULT 0,
                costo_urea NUMERIC DEFAULT 0,
                costo_ralenti NUMERIC DEFAULT 0,
                bono NUMERIC DEFAULT 0,
                diferencia NUMERIC DEFAULT 0,
                carga_manejo NUMERIC DEFAULT 0,
                sellos TEXT,
                comentarios_sellos TEXT,
                dialogo TEXT,
                auxiliar VARCHAR(128),
                raw_data JSONB DEFAULT '{}'::jsonb,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );

            CREATE INDEX IF NOT EXISTS idx_diesel_unidad ON diesel_registros_rendimiento(unidad);
            CREATE INDEX IF NOT EXISTS idx_diesel_operador ON diesel_registros_rendimiento(operador);

            CREATE TABLE IF NOT EXISTS diesel_operadores (
                id VARCHAR(64) PRIMARY KEY,
                nombre VARCHAR(128),
                numero_empleado VARCHAR(64),
                activo BOOLEAN DEFAULT true,
                ruta_habitual VARCHAR(64),
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS diesel_autorizaciones (
                id VARCHAR(64) PRIMARY KEY,
                fecha VARCHAR(64),
                unidad VARCHAR(64),
                operador VARCHAR(128),
                litros_autorizados NUMERIC DEFAULT 0,
                km_actual NUMERIC DEFAULT 0,
                destino TEXT,
                estatus VARCHAR(64) DEFAULT 'AUTORIZADO',
                autorizador VARCHAR(128),
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS diesel_pipas (
                id VARCHAR(64) PRIMARY KEY,
                fecha VARCHAR(64),
                litros_recibidos NUMERIC DEFAULT 0,
                proveedor VARCHAR(128),
                factura VARCHAR(64),
                costo_total NUMERIC DEFAULT 0,
                notas TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS diesel_config (
                clave VARCHAR(64) PRIMARY KEY,
                valor JSONB DEFAULT '{}'::jsonb,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
        `);

        // C) MÓDULO ALMACÉN E INVENTARIO
        await client.query(`
            CREATE TABLE IF NOT EXISTS almacen_catalog (
                id_pieza VARCHAR(64) PRIMARY KEY,
                numero_parte VARCHAR(64),
                descripcion TEXT,
                marca VARCHAR(64),
                stock_actual NUMERIC DEFAULT 0,
                stock_minimo NUMERIC DEFAULT 0,
                rack VARCHAR(32),
                nivel VARCHAR(32),
                precio_unitario NUMERIC DEFAULT 0,
                almacen VARCHAR(64) DEFAULT 'Almacén de Refacciones',
                almacen_id VARCHAR(64) DEFAULT 'refacciones',
                unidad_medida VARCHAR(32) DEFAULT 'PZ',
                datos_extra JSONB DEFAULT '{}'::jsonb,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );

            CREATE INDEX IF NOT EXISTS idx_almacen_numpat ON almacen_catalog(numero_parte);

            CREATE TABLE IF NOT EXISTS almacen_movimientos (
                id VARCHAR(64) PRIMARY KEY,
                tipo VARCHAR(16) DEFAULT 'S',
                folio VARCHAR(64),
                fecha VARCHAR(32),
                hora VARCHAR(32),
                id_pieza VARCHAR(64),
                numero_parte VARCHAR(64),
                descripcion TEXT,
                marca VARCHAR(64),
                cantidad NUMERIC DEFAULT 1,
                unidad_camion VARCHAR(64),
                solicitante VARCHAR(128),
                precio_unitario NUMERIC DEFAULT 0,
                costo_total NUMERIC DEFAULT 0,
                factura VARCHAR(64),
                proveedor VARCHAR(128),
                ubicacion TEXT,
                motivo TEXT,
                anulado BOOLEAN DEFAULT false,
                datos_extra JSONB DEFAULT '{}'::jsonb,
                timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );

            CREATE INDEX IF NOT EXISTS idx_almacen_mov_folio ON almacen_movimientos(folio);
            CREATE INDEX IF NOT EXISTS idx_almacen_mov_unidad ON almacen_movimientos(unidad_camion);

            CREATE TABLE IF NOT EXISTS almacen_folios_taller (
                folio VARCHAR(64) PRIMARY KEY,
                unidad VARCHAR(64),
                solicitante VARCHAR(128),
                motivo TEXT,
                fecha VARCHAR(32),
                hora VARCHAR(32),
                estado VARCHAR(64) DEFAULT 'abierto',
                timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
        `);

        await client.query('COMMIT');
        console.log('[PostgreSQL] Esquema ERP verificado e inicializado correctamente.');
        return true;
    } catch (err) {
        await client.query('ROLLBACK');
        console.error('[PostgreSQL] Error inicializando esquema:', err.message);
        return false;
    } finally {
        client.release();
    }
}

// -------------------------------------------------------------------------
// 2. MIGRACIÓN AUTOMÁTICA DESDE ARCHIVOS EXISTENTES A POSTGRESQL (SEMBRADO SEGURO)
// -------------------------------------------------------------------------
async function sembrarDatosInicialesSiVacio(rootDir = __dirname) {
    if (!isPostgresAvailable || !pool) return;

    try {
        const client = await pool.connect();
        try {
            // A) Sembrado de Taller Unidades
            const resUnidades = await client.query('SELECT COUNT(*) FROM taller_unidades;');
            if (parseInt(resUnidades.rows[0].count) === 0) {
                const fPath = path.join(rootDir, 'taller_db.json');
                if (fs.existsSync(fPath)) {
                    const unidades = JSON.parse(fs.readFileSync(fPath, 'utf8'));
                    if (Array.isArray(unidades) && unidades.length > 0) {
                        for (const u of unidades) {
                            await client.query(`
                                INSERT INTO taller_unidades 
                                (id, sucursal, estado, defecto, reporte, notas, costo, costo_refacciones, costo_preventivo, costo_correctivo, tipo_procedimiento, folio_almacen, fecha_ingreso, archivada, avances, piezas, historial)
                                VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17)
                                ON CONFLICT (id) DO NOTHING;
                            `, [
                                String(u.id), u.sucursal || 'Apodaca', u.estado || 'proceso', u.defecto || '', u.reporte || '', u.notas || '',
                                Number(u.costo) || 0, Number(u.costo_refacciones) || 0, Number(u.costo_preventivo) || 0, Number(u.costo_correctivo) || 0,
                                u.tipo_procedimiento || 'correctivo', u.folio_almacen || null, u.fecha_ingreso || null, Boolean(u.archivada),
                                JSON.stringify(u.avances || []), JSON.stringify(u.piezas || []), JSON.stringify(u.historial || [])
                            ]);
                        }
                        console.log(`[PostgreSQL Migración] Se migraron ${unidades.length} unidades activas de Taller.`);
                    }
                }
            }

            // B) Sembrado de Cajas y Remolques
            const resRem = await client.query('SELECT COUNT(*) FROM taller_remolques;');
            if (parseInt(resRem.rows[0].count) === 0) {
                const fPath = path.join(rootDir, 'taller_remolques.json');
                if (fs.existsSync(fPath)) {
                    const remolques = JSON.parse(fs.readFileSync(fPath, 'utf8'));
                    if (Array.isArray(remolques) && remolques.length > 0) {
                        for (const r of remolques) {
                            await client.query(`
                                INSERT INTO taller_remolques 
                                (id, tipo, marca, anio, placas, estado, sistema, defecto, reporte, mecanico, costo, costo_refacciones, sucursal, folio_almacen, fecha, fecha_ingreso, piezas)
                                VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17)
                                ON CONFLICT (id) DO NOTHING;
                            `, [
                                String(r.id), r.tipo || 'Caja Seca 53', r.marca || '', r.anio || '', r.placas || '', r.estado || 'proceso',
                                r.sistema || 'General', r.defecto || '', r.reporte || '', r.mecanico || 'Sin asignar',
                                Number(r.costo) || 0, Number(r.costo_refacciones) || 0, r.sucursal || 'Apodaca', r.folio_almacen || null,
                                r.fecha || null, r.fechaIngreso || null, JSON.stringify(r.piezas || [])
                            ]);
                        }
                        console.log(`[PostgreSQL Migración] Se migraron ${remolques.length} cajas y remolques reales.`);
                    }
                }
            }

            // C) Sembrado de Sistema de Diésel & Rendimientos
            const resDiesel = await client.query('SELECT COUNT(*) FROM diesel_registros_rendimiento;');
            if (parseInt(resDiesel.rows[0].count) === 0) {
                const fPath = path.join(rootDir, 'diesel_local_cache.json');
                if (fs.existsSync(fPath)) {
                    const dieselCache = JSON.parse(fs.readFileSync(fPath, 'utf8'));
                    const rendimientos = dieselCache.registros_rendimiento || [];
                    if (Array.isArray(rendimientos) && rendimientos.length > 0) {
                        for (const r of rendimientos) {
                            const numId = String(r.id || `d_${Date.now()}_${Math.random()}`);
                            const numCostoDiesel = typeof r.costo_diesel === 'string' ? Number(r.costo_diesel.replace(/[^0-9.-]+/g, '')) || 0 : Number(r.costo_diesel) || 0;
                            const numBono = typeof r.bono === 'string' ? Number(r.bono.replace(/[^0-9.-]+/g, '')) || 0 : Number(r.bono) || 0;
                            const numCostoRalenti = typeof r.costo_ralenti === 'string' ? Number(r.costo_ralenti.replace(/[^0-9.-]+/g, '')) || 0 : Number(r.costo_ralenti) || 0;

                            await client.query(`
                                INSERT INTO diesel_registros_rendimiento 
                                (id, fecha, unidad, operador, ruta, km_recorridos, km_total, rend_ecm, rend_fisico, ecm, real_total, real_apodaca, real_laredo, real_foraneo, urea, rend_urea, litros_ralenti, costo_diesel, costo_ralenti, bono, diferencia, carga_manejo, sellos, comentarios_sellos, dialogo, auxiliar, raw_data)
                                VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17, $18, $19, $20, $21, $22, $23, $24, $25, $26, $27)
                                ON CONFLICT (id) DO NOTHING;
                            `, [
                                numId, r.fecha || '', String(r.unidad || ''), r.operador || '', r.ruta || '',
                                Number(r.km_recorridos) || 0, Number(r.km_total) || 0, Number(r.rend_ecm) || 0, Number(r.rend_fisico) || 0,
                                Number(r.ecm) || 0, Number(r.real_total) || 0, Number(r.real_apodaca) || 0, Number(r.real_laredo) || 0, Number(r.real_foraneo) || 0,
                                Number(r.urea) || 0, Number(r.rend_urea) || 0, Number(r.litros_ralenti) || 0, numCostoDiesel, numCostoRalenti, numBono,
                                Number(r.diferencia) || 0, Number(r.carga_manejo) || 0, r.sellos || '', r.comentarios_sellos || '', r.dialogo || '', r.auxiliar || '',
                                JSON.stringify(r)
                            ]);
                        }
                        console.log(`[PostgreSQL Migración] Se migraron ${rendimientos.length} registros históricos de Diésel y Rendimientos.`);
                    }

                    // Operadores
                    const operadores = dieselCache.operadores || [];
                    for (const op of operadores) {
                        const opId = String(op.id || op.nombre || Math.random());
                        await client.query(`
                            INSERT INTO diesel_operadores (id, nombre, numero_empleado, activo)
                            VALUES ($1, $2, $3, $4)
                            ON CONFLICT (id) DO NOTHING;
                        `, [opId, op.nombre || opId, op.numero_empleado || '', true]);
                    }

                    // Autorizaciones de Diesel
                    const auts = dieselCache.autorizaciones || [];
                    for (const a of auts) {
                        await client.query(`
                            INSERT INTO diesel_autorizaciones (id, fecha, unidad, operador, litros_autorizados, km_actual, destino, estatus, autorizador)
                            VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9)
                            ON CONFLICT (id) DO NOTHING;
                        `, [
                            String(a.id || Math.random()), a.fecha || '', String(a.unidad || ''), a.operador || '',
                            Number(a.litros_autorizados) || 0, Number(a.km_actual) || 0, a.destino || '', a.estatus || 'AUTORIZADO', a.autorizador || ''
                        ]);
                    }
                }
            }

            // D) Sembrado de Almacén
            const resAlm = await client.query('SELECT COUNT(*) FROM almacen_movimientos;');
            if (parseInt(resAlm.rows[0].count) === 0) {
                const fPath = path.join(rootDir, 'almacen_datos.json');
                if (fs.existsSync(fPath)) {
                    const almData = JSON.parse(fs.readFileSync(fPath, 'utf8'));
                    const catalog = almData.catalog || {};
                    for (const [idPieza, item] of Object.entries(catalog)) {
                        await client.query(`
                            INSERT INTO almacen_catalog (id_pieza, numero_parte, descripcion, marca, stock_actual, stock_minimo, rack, nivel, precio_unitario, almacen, almacen_id, unidad_medida)
                            VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12)
                            ON CONFLICT (id_pieza) DO NOTHING;
                        `, [
                            idPieza, item.numero_parte || '', item.descripcion || '', item.marca || '',
                            Number(item.stock_actual) || 0, Number(item.stock_minimo) || 0, item.rack || '', item.nivel || '',
                            Number(item.precio_unitario) || 0, item.almacen || 'Almacén de Refacciones', item.almacen_id || 'refacciones', item.unidad || 'PZ'
                        ]);
                    }

                    const movs = almData.movimientos || almData.salidas || [];
                    for (const m of movs) {
                        const pu = Number(m.costo_unitario || m.precio_unitario) || 0;
                        const tot = Number(m.costo_total !== undefined ? m.costo_total : (m.total_costo || (m.cantidad * pu))) || 0;
                        await client.query(`
                            INSERT INTO almacen_movimientos 
                            (id, tipo, folio, fecha, hora, id_pieza, numero_parte, descripcion, marca, cantidad, unidad_camion, solicitante, precio_unitario, costo_total, factura, proveedor, ubicacion, motivo, anulado)
                            VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17, $18, $19)
                            ON CONFLICT (id) DO NOTHING;
                        `, [
                            String(m.id), m.tipo || 'S', m.folio || '', m.fecha || '', m.hora || '',
                            m.id_pieza || '', m.numero_parte || '', m.descripcion || '', m.marca || '',
                            Number(m.cantidad) || 1, String(m.unidad_camion || m.unidad || ''), m.solicitante || '',
                            pu, tot, m.factura || '', m.proveedor || '', m.ubicacion || '', m.motivo || '', Boolean(m.anulado)
                        ]);
                    }
                    console.log(`[PostgreSQL Migración] Se migraron ${Object.keys(catalog).length} productos y ${movs.length} movimientos de Almacén.`);
                }
            }

            // E) Sembrado de Rescates Carreteros
            const resResc = await client.query('SELECT COUNT(*) FROM taller_rescates;');
            if (parseInt(resResc.rows[0].count) === 0) {
                const fPath = path.join(rootDir, 'taller_rescates.json');
                if (fs.existsSync(fPath)) {
                    const rescates = JSON.parse(fs.readFileSync(fPath, 'utf8'));
                    if (Array.isArray(rescates) && rescates.length > 0) {
                        for (const r of rescates) {
                            if (String(r.unidad).trim() === '75') continue; // Ignorar prueba de unidad 75
                            await client.query(`
                                INSERT INTO taller_rescates 
                                (id, folio, unidad, operador, telefono, carretera, kilometro, km, tramo, sentido, falla_reportada, tipo_falla, costo_estimado, costo_real, estado, estatus, mecanico_asignado, proveedor_externo, tiempo_llegada, tiempo_resolucion, hora_conclusion, notas, sucursal, fecha, hora_reporte, avances, raw_data)
                                VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17, $18, $19, $20, $21, $22, $23, $24, $25, $26, $27)
                                ON CONFLICT (id) DO NOTHING;
                            `, [
                                String(r.id), r.folio || '', String(r.unidad || ''), r.operador || '', r.telefono_operador || r.telefono || '',
                                r.carretera || '', r.kilometro || '', r.km || '', r.tramo || '', r.sentido || '',
                                r.falla_reportada || r.falla || '', r.tipo_falla || 'Mecánica', Number(r.costo_estimado || r.costo) || 0,
                                Number(r.costo_real) || 0, r.estado || 'EN_CAMINO', r.estatus || r.estado || 'EN_CAMINO',
                                r.mecanico_asignado || r.mecanico || '', r.proveedor_externo || '', r.tiempo_llegada || '',
                                r.tiempo_resolucion || '', r.hora_conclusion || null, r.notas || '', r.sucursal || 'Apodaca',
                                r.fecha || '', r.hora_reporte || '', JSON.stringify(r.avances || []), JSON.stringify(r)
                            ]);
                        }
                        console.log(`[PostgreSQL Migración] Se migraron ${rescates.length} rescates carreteros.`);
                    }
                }
            }

        } finally {
            client.release();
        }
    } catch (err) {
        console.error('[PostgreSQL Migración Error]:', err.message);
    }
}

// -------------------------------------------------------------------------
// 3. CAPA DE ACCESO DUAL-STORAGE (POSTGRESQL + LOCAL BACKUP REFLEJO)
// -------------------------------------------------------------------------

// TALLER UNIDADES
async function obtenerUnidadesTaller() {
    if (isPostgresAvailable && pool) {
        try {
            const res = await pool.query('SELECT * FROM taller_unidades WHERE archivada = false ORDER BY id ASC;');
            if (res.rows.length > 0) {
                return res.rows.map(r => ({
                    id: r.id,
                    sucursal: r.sucursal,
                    estado: r.estado,
                    defecto: r.defecto,
                    reporte: r.reporte,
                    notas: r.notas,
                    costo: Number(r.costo) || 0,
                    costo_refacciones: Number(r.costo_refacciones) || 0,
                    costo_preventivo: Number(r.costo_preventivo) || 0,
                    costo_correctivo: Number(r.costo_correctivo) || 0,
                    tipo_procedimiento: r.tipo_procedimiento,
                    folio_almacen: r.folio_almacen,
                    fecha_ingreso: r.fecha_ingreso,
                    archivada: r.archivada,
                    avances: r.avances || [],
                    piezas: r.piezas || [],
                    historial: r.historial || []
                }));
            }
        } catch (e) {
            console.warn('[PostgreSQL Read Unidades Error, usando disco]:', e.message);
        }
    }
    // Fallback a disco
    try {
        const f = path.join(__dirname, 'taller_db.json');
        if (fs.existsSync(f)) return JSON.parse(fs.readFileSync(f, 'utf8'));
    } catch(e) {}
    return [];
}

async function guardarUnidadesTaller(unidades) {
    if (!Array.isArray(unidades)) return false;

    // Respaldo en disco local garantizado siempre (Zero Data Loss)
    try {
        const f = path.join(__dirname, 'taller_db.json');
        fs.writeFileSync(f, JSON.stringify(unidades, null, 2), 'utf8');
    } catch (e) {}

    if (isPostgresAvailable && pool) {
        try {
            const client = await pool.connect();
            try {
                await client.query('BEGIN');
                for (const u of unidades) {
                    await client.query(`
                        INSERT INTO taller_unidades 
                        (id, sucursal, estado, defecto, reporte, notas, costo, costo_refacciones, costo_preventivo, costo_correctivo, tipo_procedimiento, folio_almacen, fecha_ingreso, archivada, avances, piezas, historial, updated_at)
                        VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17, CURRENT_TIMESTAMP)
                        ON CONFLICT (id) DO UPDATE SET
                            sucursal = EXCLUDED.sucursal,
                            estado = EXCLUDED.estado,
                            defecto = EXCLUDED.defecto,
                            reporte = EXCLUDED.reporte,
                            notas = EXCLUDED.notas,
                            costo = EXCLUDED.costo,
                            costo_refacciones = EXCLUDED.costo_refacciones,
                            costo_preventivo = EXCLUDED.costo_preventivo,
                            costo_correctivo = EXCLUDED.costo_correctivo,
                            tipo_procedimiento = EXCLUDED.tipo_procedimiento,
                            folio_almacen = EXCLUDED.folio_almacen,
                            archivada = EXCLUDED.archivada,
                            avances = EXCLUDED.avances,
                            piezas = EXCLUDED.piezas,
                            historial = EXCLUDED.historial,
                            updated_at = CURRENT_TIMESTAMP;
                    `, [
                        String(u.id), u.sucursal || 'Apodaca', u.estado || 'proceso', u.defecto || '', u.reporte || '', u.notas || '',
                        Number(u.costo) || 0, Number(u.costo_refacciones) || 0, Number(u.costo_preventivo) || 0, Number(u.costo_correctivo) || 0,
                        u.tipo_procedimiento || 'correctivo', u.folio_almacen || null, u.fecha_ingreso || null, Boolean(u.archivada),
                        JSON.stringify(u.avances || []), JSON.stringify(u.piezas || []), JSON.stringify(u.historial || [])
                    ]);
                }
                await client.query('COMMIT');
                return true;
            } catch (err) {
                await client.query('ROLLBACK');
                console.error('[PostgreSQL Save Unidades Error]:', err.message);
            } finally {
                client.release();
            }
        } catch(e) {}
    }
    return true;
}

// CAJAS Y REMOLQUES
async function obtenerRemolquesTaller() {
    if (isPostgresAvailable && pool) {
        try {
            const res = await pool.query('SELECT * FROM taller_remolques ORDER BY id ASC;');
            if (res.rows.length > 0) {
                return res.rows.map(r => ({
                    id: r.id,
                    tipo: r.tipo,
                    marca: r.marca,
                    anio: r.anio,
                    placas: r.placas,
                    estado: r.estado,
                    sistema: r.sistema,
                    defecto: r.defecto,
                    reporte: r.reporte,
                    mecanico: r.mecanico,
                    costo: Number(r.costo) || 0,
                    costo_refacciones: Number(r.costo_refacciones) || 0,
                    sucursal: r.sucursal,
                    folio_almacen: r.folio_almacen,
                    fecha: r.fecha,
                    fechaIngreso: r.fecha_ingreso,
                    piezas: r.piezas || []
                }));
            }
        } catch (e) {
            console.warn('[PostgreSQL Read Remolques Error]:', e.message);
        }
    }
    try {
        const f = path.join(__dirname, 'taller_remolques.json');
        if (fs.existsSync(f)) return JSON.parse(fs.readFileSync(f, 'utf8'));
    } catch(e) {}
    return [];
}

async function guardarRemolquesTaller(remolques) {
    if (!Array.isArray(remolques)) return false;

    try {
        const f = path.join(__dirname, 'taller_remolques.json');
        fs.writeFileSync(f, JSON.stringify(remolques, null, 2), 'utf8');
    } catch(e) {}

    if (isPostgresAvailable && pool) {
        try {
            const client = await pool.connect();
            try {
                await client.query('BEGIN');
                for (const r of remolques) {
                    await client.query(`
                        INSERT INTO taller_remolques 
                        (id, tipo, marca, anio, placas, estado, sistema, defecto, reporte, mecanico, costo, costo_refacciones, sucursal, folio_almacen, fecha, fecha_ingreso, piezas, updated_at)
                        VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17, CURRENT_TIMESTAMP)
                        ON CONFLICT (id) DO UPDATE SET
                            tipo = EXCLUDED.tipo,
                            marca = EXCLUDED.marca,
                            anio = EXCLUDED.anio,
                            placas = EXCLUDED.placas,
                            estado = EXCLUDED.estado,
                            sistema = EXCLUDED.sistema,
                            defecto = EXCLUDED.defecto,
                            reporte = EXCLUDED.reporte,
                            mecanico = EXCLUDED.mecanico,
                            costo = EXCLUDED.costo,
                            costo_refacciones = EXCLUDED.costo_refacciones,
                            sucursal = EXCLUDED.sucursal,
                            folio_almacen = EXCLUDED.folio_almacen,
                            fecha = EXCLUDED.fecha,
                            fecha_ingreso = EXCLUDED.fecha_ingreso,
                            piezas = EXCLUDED.piezas,
                            updated_at = CURRENT_TIMESTAMP;
                    `, [
                        String(r.id), r.tipo || 'Caja Seca 53', r.marca || '', r.anio || '', r.placas || '', r.estado || 'proceso',
                        r.sistema || 'General', r.defecto || '', r.reporte || '', r.mecanico || 'Sin asignar',
                        Number(r.costo) || 0, Number(r.costo_refacciones) || 0, r.sucursal || 'Apodaca', r.folio_almacen || null,
                        r.fecha || null, r.fechaIngreso || null, JSON.stringify(r.piezas || [])
                    ]);
                }
                await client.query('COMMIT');
                return true;
            } catch (err) {
                await client.query('ROLLBACK');
                console.error('[PostgreSQL Save Remolques Error]:', err.message);
            } finally {
                client.release();
            }
        } catch(e) {}
    }
    return true;
}

// =========================================================================
// MÓDULO RESCATES CARRETEROS (AUXILIOS EN RUTA)
// =========================================================================

async function obtenerRescatesTaller() {
    if (isPostgresAvailable && pool) {
        try {
            const res = await pool.query('SELECT * FROM taller_rescates ORDER BY updated_at DESC;');
            if (res.rows.length > 0) {
                return res.rows.map(r => {
                    const raw = (r.raw_data && typeof r.raw_data === 'object') ? r.raw_data : {};
                    return {
                        id: r.id,
                        folio: r.folio || raw.folio || `RC-${r.id}`,
                        unidad: r.unidad || raw.unidad || '',
                        operador: r.operador || raw.operador || '',
                        telefono_operador: r.telefono || raw.telefono_operador || '',
                        carretera: r.carretera || raw.carretera || '',
                        kilometro: r.kilometro || raw.kilometro || '',
                        km: r.km || raw.km || '',
                        tramo: r.tramo || raw.tramo || '',
                        sentido: r.sentido || raw.sentido || '',
                        falla_reportada: r.falla_reportada || raw.falla_reportada || r.falla || '',
                        tipo_falla: r.tipo_falla || raw.tipo_falla || 'Mecánica',
                        estado: r.estado || raw.estado || 'EN_CAMINO',
                        estatus: r.estatus || raw.estatus || r.estado || 'EN_CAMINO',
                        mecanico_asignado: r.mecanico_asignado || raw.mecanico_asignado || r.mecanico || '',
                        proveedor_externo: r.proveedor_externo || raw.proveedor_externo || '',
                        costo_estimado: Number(r.costo_estimado || raw.costo_estimado || r.costo) || 0,
                        costo_real: Number(r.costo_real || raw.costo_real) || 0,
                        tiempo_llegada: r.tiempo_llegada || raw.tiempo_llegada || '',
                        tiempo_resolucion: r.tiempo_resolucion || raw.tiempo_resolucion || '',
                        hora_conclusion: r.hora_conclusion || raw.hora_conclusion || null,
                        notas: r.notas || raw.notas || '',
                        sucursal: r.sucursal || raw.sucursal || 'Apodaca',
                        fecha: r.fecha || raw.fecha || '',
                        hora_reporte: r.hora_reporte || raw.hora_reporte || '',
                        avances: r.avances || raw.avances || []
                    };
                }).filter(r => String(r.unidad).trim() !== '75');
            }
        } catch (err) {
            console.error('[PostgreSQL Read Rescates Error]:', err.message);
        }
    }

    // Fallback archivo local
    try {
        const f = path.join(__dirname, 'taller_rescates.json');
        if (fs.existsSync(f)) {
            const data = JSON.parse(fs.readFileSync(f, 'utf8'));
            if (Array.isArray(data)) return data.filter(r => String(r.unidad).trim() !== '75');
        }
    } catch(e) {}
    return [];
}

async function guardarRescatesTaller(rescates) {
    if (!Array.isArray(rescates)) return false;
    const limpios = rescates.filter(r => String(r.unidad).trim() !== '75');

    try {
        const f = path.join(__dirname, 'taller_rescates.json');
        fs.writeFileSync(f, JSON.stringify(limpios, null, 2), 'utf8');
    } catch(e) {}

    if (isPostgresAvailable && pool) {
        try {
            const client = await pool.connect();
            try {
                await client.query('BEGIN');
                for (const r of limpios) {
                    await client.query(`
                        INSERT INTO taller_rescates 
                        (id, folio, unidad, operador, telefono, carretera, kilometro, km, tramo, sentido, falla_reportada, tipo_falla, costo_estimado, costo_real, estado, estatus, mecanico_asignado, proveedor_externo, tiempo_llegada, tiempo_resolucion, hora_conclusion, notas, sucursal, fecha, hora_reporte, avances, raw_data, updated_at)
                        VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17, $18, $19, $20, $21, $22, $23, $24, $25, $26, $27, CURRENT_TIMESTAMP)
                        ON CONFLICT (id) DO UPDATE SET
                            folio = EXCLUDED.folio,
                            unidad = EXCLUDED.unidad,
                            operador = EXCLUDED.operador,
                            telefono = EXCLUDED.telefono,
                            carretera = EXCLUDED.carretera,
                            kilometro = EXCLUDED.kilometro,
                            km = EXCLUDED.km,
                            tramo = EXCLUDED.tramo,
                            sentido = EXCLUDED.sentido,
                            falla_reportada = EXCLUDED.falla_reportada,
                            tipo_falla = EXCLUDED.tipo_falla,
                            costo_estimado = EXCLUDED.costo_estimado,
                            costo_real = EXCLUDED.costo_real,
                            estado = EXCLUDED.estado,
                            estatus = EXCLUDED.estatus,
                            mecanico_asignado = EXCLUDED.mecanico_asignado,
                            proveedor_externo = EXCLUDED.proveedor_externo,
                            tiempo_llegada = EXCLUDED.tiempo_llegada,
                            tiempo_resolucion = EXCLUDED.tiempo_resolucion,
                            hora_conclusion = EXCLUDED.hora_conclusion,
                            notas = EXCLUDED.notas,
                            sucursal = EXCLUDED.sucursal,
                            fecha = EXCLUDED.fecha,
                            hora_reporte = EXCLUDED.hora_reporte,
                            avances = EXCLUDED.avances,
                            raw_data = EXCLUDED.raw_data,
                            updated_at = CURRENT_TIMESTAMP;
                    `, [
                        String(r.id), r.folio || '', String(r.unidad || ''), r.operador || '', r.telefono_operador || r.telefono || '',
                        r.carretera || '', r.kilometro || '', r.km || '', r.tramo || '', r.sentido || '',
                        r.falla_reportada || r.falla || '', r.tipo_falla || 'Mecánica', Number(r.costo_estimado || r.costo) || 0,
                        Number(r.costo_real) || 0, r.estado || 'EN_CAMINO', r.estatus || r.estado || 'EN_CAMINO',
                        r.mecanico_asignado || r.mecanico || '', r.proveedor_externo || '', r.tiempo_llegada || '',
                        r.tiempo_resolucion || '', r.hora_conclusion || null, r.notas || '', r.sucursal || 'Apodaca',
                        r.fecha || '', r.hora_reporte || '', JSON.stringify(r.avances || []), JSON.stringify(r)
                    ]);
                }
                await client.query('COMMIT');
                return true;
            } catch (err) {
                await client.query('ROLLBACK');
                console.error('[PostgreSQL Save Rescates Error]:', err.message);
            } finally {
                client.release();
            }
        } catch(e) {}
    }
    return true;
}

async function guardarRescateTaller(rescate) {
    if (!rescate || !rescate.id) return false;
    if (String(rescate.unidad).trim() === '75') return false; // Bloquear permanentemente prueba de 75
    const rescates = await obtenerRescatesTaller();
    const idx = rescates.findIndex(r => String(r.id) === String(rescate.id));
    if (idx >= 0) {
        rescates[idx] = { ...rescates[idx], ...rescate };
    } else {
        rescates.unshift(rescate);
    }
    return await guardarRescatesTaller(rescates);
}

async function eliminarRescateTaller(id) {
    if (!id) return false;
    const rescates = await obtenerRescatesTaller();
    const filtrados = rescates.filter(r => String(r.id) !== String(id));
    
    try {
        const f = path.join(__dirname, 'taller_rescates.json');
        fs.writeFileSync(f, JSON.stringify(filtrados, null, 2), 'utf8');
    } catch(e) {}

    if (isPostgresAvailable && pool) {
        try {
            await pool.query('DELETE FROM taller_rescates WHERE id = $1;', [String(id)]);
        } catch (err) {
            console.error('[PostgreSQL Delete Rescate Error]:', err.message);
        }
    }
    return true;
}

// MÓDULO DIÉSEL & RENDIMIENTOS
async function obtenerRegistrosDiesel(limit = 1000) {
    if (isPostgresAvailable && pool) {
        try {
            const res = await pool.query('SELECT * FROM diesel_registros_rendimiento ORDER BY timestamp DESC LIMIT $1;', [limit]);
            if (res.rows.length > 0) {
                return res.rows.map(r => ({
                    id: r.id,
                    fecha: r.fecha,
                    unidad: r.unidad,
                    operador: r.operador,
                    ruta: r.ruta,
                    km_recorridos: Number(r.km_recorridos) || 0,
                    km_total: Number(r.km_total) || 0,
                    rend_ecm: Number(r.rend_ecm) || 0,
                    rend_fisico: Number(r.rend_fisico) || 0,
                    ecm: Number(r.ecm) || 0,
                    real_total: Number(r.real_total) || 0,
                    real_apodaca: Number(r.real_apodaca) || 0,
                    real_laredo: Number(r.real_laredo) || 0,
                    real_foraneo: Number(r.real_foraneo) || 0,
                    urea: Number(r.urea) || 0,
                    rend_urea: Number(r.rend_urea) || 0,
                    litros_ralenti: Number(r.litros_ralenti) || 0,
                    costo_diesel: Number(r.costo_diesel) || 0,
                    costo_urea: Number(r.costo_urea) || 0,
                    costo_ralenti: Number(r.costo_ralenti) || 0,
                    bono: Number(r.bono) || 0,
                    diferencia: Number(r.diferencia) || 0,
                    carga_manejo: Number(r.carga_manejo) || 0,
                    sellos: r.sellos,
                    comentarios_sellos: r.comentarios_sellos,
                    dialogo: r.dialogo,
                    auxiliar: r.auxiliar
                }));
            }
        } catch(e) {
            console.warn('[PostgreSQL Read Diesel Error]:', e.message);
        }
    }
    try {
        const f = path.join(__dirname, 'diesel_local_cache.json');
        if (fs.existsSync(f)) {
            const d = JSON.parse(fs.readFileSync(f, 'utf8'));
            return d.registros_rendimiento || [];
        }
    } catch(e) {}
    return [];
}

async function guardarRegistroDiesel(r) {
    if (!r || !r.unidad) return false;
    const numId = String(r.id || `d_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`);

    // Actualizar cache local
    try {
        const f = path.join(__dirname, 'diesel_local_cache.json');
        if (fs.existsSync(f)) {
            const cache = JSON.parse(fs.readFileSync(f, 'utf8'));
            cache.registros_rendimiento = cache.registros_rendimiento || [];
            const idx = cache.registros_rendimiento.findIndex(x => x.id === numId);
            if (idx >= 0) cache.registros_rendimiento[idx] = { ...cache.registros_rendimiento[idx], ...r };
            else cache.registros_rendimiento.unshift({ ...r, id: numId });
            fs.writeFileSync(f, JSON.stringify(cache, null, 2), 'utf8');
        }
    } catch(e) {}

    if (isPostgresAvailable && pool) {
        try {
            const numCostoDiesel = typeof r.costo_diesel === 'string' ? Number(r.costo_diesel.replace(/[^0-9.-]+/g, '')) || 0 : Number(r.costo_diesel) || 0;
            const numBono = typeof r.bono === 'string' ? Number(r.bono.replace(/[^0-9.-]+/g, '')) || 0 : Number(r.bono) || 0;
            const numCostoRalenti = typeof r.costo_ralenti === 'string' ? Number(r.costo_ralenti.replace(/[^0-9.-]+/g, '')) || 0 : Number(r.costo_ralenti) || 0;

            await pool.query(`
                INSERT INTO diesel_registros_rendimiento 
                (id, fecha, unidad, operador, ruta, km_recorridos, km_total, rend_ecm, rend_fisico, ecm, real_total, real_apodaca, real_laredo, real_foraneo, urea, rend_urea, litros_ralenti, costo_diesel, costo_ralenti, bono, diferencia, carga_manejo, sellos, comentarios_sellos, dialogo, auxiliar, raw_data, created_at)
                VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17, $18, $19, $20, $21, $22, $23, $24, $25, $26, $27, CURRENT_TIMESTAMP)
                ON CONFLICT (id) DO UPDATE SET
                    fecha = EXCLUDED.fecha,
                    unidad = EXCLUDED.unidad,
                    operador = EXCLUDED.operador,
                    ruta = EXCLUDED.ruta,
                    km_recorridos = EXCLUDED.km_recorridos,
                    km_total = EXCLUDED.km_total,
                    rend_ecm = EXCLUDED.rend_ecm,
                    rend_fisico = EXCLUDED.rend_fisico,
                    ecm = EXCLUDED.ecm,
                    real_total = EXCLUDED.real_total,
                    real_apodaca = EXCLUDED.real_apodaca,
                    real_laredo = EXCLUDED.real_laredo,
                    real_foraneo = EXCLUDED.real_foraneo,
                    urea = EXCLUDED.urea,
                    rend_urea = EXCLUDED.rend_urea,
                    litros_ralenti = EXCLUDED.litros_ralenti,
                    costo_diesel = EXCLUDED.costo_diesel,
                    costo_ralenti = EXCLUDED.costo_ralenti,
                    bono = EXCLUDED.bono,
                    diferencia = EXCLUDED.diferencia,
                    carga_manejo = EXCLUDED.carga_manejo,
                    sellos = EXCLUDED.sellos,
                    comentarios_sellos = EXCLUDED.comentarios_sellos,
                    dialogo = EXCLUDED.dialogo,
                    auxiliar = EXCLUDED.auxiliar,
                    raw_data = EXCLUDED.raw_data;
            `, [
                numId, r.fecha || '', String(r.unidad || ''), r.operador || '', r.ruta || '',
                Number(r.km_recorridos) || 0, Number(r.km_total) || 0, Number(r.rend_ecm) || 0, Number(r.rend_fisico) || 0,
                Number(r.ecm) || 0, Number(r.real_total) || 0, Number(r.real_apodaca) || 0, Number(r.real_laredo) || 0, Number(r.real_foraneo) || 0,
                Number(r.urea) || 0, Number(r.rend_urea) || 0, Number(r.litros_ralenti) || 0, numCostoDiesel, numCostoRalenti, numBono,
                Number(r.diferencia) || 0, Number(r.carga_manejo) || 0, r.sellos || '', r.comentarios_sellos || '', r.dialogo || '', r.auxiliar || '',
                JSON.stringify(r)
            ]);
            return true;
        } catch(e) {
            console.error('[PostgreSQL Save Diesel Error]:', e.message);
        }
    }
    return true;
}

// OPERADORES DE DIÉSEL
async function obtenerOperadoresDiesel() {
    if (isPostgresAvailable && pool) {
        try {
            const res = await pool.query('SELECT * FROM diesel_operadores WHERE activo = true ORDER BY nombre ASC;');
            if (res.rows.length > 0) {
                return res.rows.map(r => ({
                    id: r.id,
                    nombre: r.nombre,
                    numero_empleado: r.numero_empleado,
                    activo: r.activo,
                    ruta_habitual: r.ruta_habitual
                }));
            }
        } catch (e) {
            console.warn('[PostgreSQL Read Operadores Error]:', e.message);
        }
    }
    try {
        const f = path.join(__dirname, 'diesel_local_cache.json');
        if (fs.existsSync(f)) {
            const d = JSON.parse(fs.readFileSync(f, 'utf8'));
            return d.operadores || [];
        }
    } catch(e) {}
    return [];
}

async function guardarOperadorDiesel(op) {
    if (!op || !op.nombre) return false;
    const opId = String(op.id || op.numero_empleado || op.nombre.toLowerCase().replace(/\s+/g, '_'));

    try {
        const f = path.join(__dirname, 'diesel_local_cache.json');
        if (fs.existsSync(f)) {
            const cache = JSON.parse(fs.readFileSync(f, 'utf8'));
            cache.operadores = cache.operadores || [];
            const idx = cache.operadores.findIndex(x => String(x.id) === opId || x.nombre === op.nombre);
            if (idx >= 0) cache.operadores[idx] = { ...cache.operadores[idx], ...op };
            else cache.operadores.push({ ...op, id: opId });
            fs.writeFileSync(f, JSON.stringify(cache, null, 2), 'utf8');
        }
    } catch(e) {}

    if (isPostgresAvailable && pool) {
        try {
            await pool.query(`
                INSERT INTO diesel_operadores (id, nombre, numero_empleado, activo, ruta_habitual, updated_at)
                VALUES ($1, $2, $3, $4, $5, CURRENT_TIMESTAMP)
                ON CONFLICT (id) DO UPDATE SET
                    nombre = EXCLUDED.nombre,
                    numero_empleado = EXCLUDED.numero_empleado,
                    activo = EXCLUDED.activo,
                    ruta_habitual = EXCLUDED.ruta_habitual,
                    updated_at = CURRENT_TIMESTAMP;
            `, [opId, op.nombre, op.numero_empleado || '', op.activo !== false, op.ruta_habitual || '']);
            return true;
        } catch(e) {
            console.error('[PostgreSQL Save Operador Error]:', e.message);
        }
    }
    return true;
}

// AUTORIZACIONES DE DIÉSEL
async function obtenerAutorizacionesDiesel(limit = 100) {
    if (isPostgresAvailable && pool) {
        try {
            const res = await pool.query('SELECT * FROM diesel_autorizaciones ORDER BY created_at DESC LIMIT $1;', [limit]);
            if (res.rows.length > 0) {
                return res.rows.map(r => ({
                    id: r.id,
                    fecha: r.fecha,
                    unidad: r.unidad,
                    operador: r.operador,
                    litros_autorizados: Number(r.litros_autorizados) || 0,
                    km_actual: Number(r.km_actual) || 0,
                    destino: r.destino,
                    estatus: r.estatus,
                    autorizador: r.autorizador
                }));
            }
        } catch (e) {
            console.warn('[PostgreSQL Read Autorizaciones Error]:', e.message);
        }
    }
    try {
        const f = path.join(__dirname, 'diesel_local_cache.json');
        if (fs.existsSync(f)) {
            const d = JSON.parse(fs.readFileSync(f, 'utf8'));
            return d.autorizaciones || [];
        }
    } catch(e) {}
    return [];
}

async function guardarAutorizacionDiesel(aut) {
    if (!aut || !aut.unidad) return false;
    const autId = String(aut.id || `aut_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`);

    try {
        const f = path.join(__dirname, 'diesel_local_cache.json');
        if (fs.existsSync(f)) {
            const cache = JSON.parse(fs.readFileSync(f, 'utf8'));
            cache.autorizaciones = cache.autorizaciones || [];
            cache.autorizaciones.unshift({ ...aut, id: autId });
            fs.writeFileSync(f, JSON.stringify(cache, null, 2), 'utf8');
        }
    } catch(e) {}

    if (isPostgresAvailable && pool) {
        try {
            await pool.query(`
                INSERT INTO diesel_autorizaciones (id, fecha, unidad, operador, litros_autorizados, km_actual, destino, estatus, autorizador, created_at)
                VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, CURRENT_TIMESTAMP)
                ON CONFLICT (id) DO UPDATE SET
                    fecha = EXCLUDED.fecha,
                    unidad = EXCLUDED.unidad,
                    operador = EXCLUDED.operador,
                    litros_autorizados = EXCLUDED.litros_autorizados,
                    km_actual = EXCLUDED.km_actual,
                    destino = EXCLUDED.destino,
                    estatus = EXCLUDED.estatus,
                    autorizador = EXCLUDED.autorizador;
            `, [
                autId, aut.fecha || new Date().toISOString().slice(0, 10), String(aut.unidad), aut.operador || '',
                Number(aut.litros_autorizados) || 0, Number(aut.km_actual) || 0, aut.destino || '', aut.estatus || 'AUTORIZADO', aut.autorizador || ''
            ]);
            return true;
        } catch(e) {
            console.error('[PostgreSQL Save Autorizacion Error]:', e.message);
        }
    }
    return true;
}

// PIPAS Y RECARGAS DE DIÉSEL
async function obtenerPipasDiesel(limit = 100) {
    if (isPostgresAvailable && pool) {
        try {
            const res = await pool.query('SELECT * FROM diesel_pipas ORDER BY created_at DESC LIMIT $1;', [limit]);
            if (res.rows.length > 0) {
                return res.rows.map(r => ({
                    id: r.id,
                    fecha: r.fecha,
                    litros_recibidos: Number(r.litros_recibidos) || 0,
                    proveedor: r.proveedor,
                    factura: r.factura,
                    costo_total: Number(r.costo_total) || 0,
                    notas: r.notas
                }));
            }
        } catch (e) {
            console.warn('[PostgreSQL Read Pipas Error]:', e.message);
        }
    }
    try {
        const f = path.join(__dirname, 'diesel_local_cache.json');
        if (fs.existsSync(f)) {
            const d = JSON.parse(fs.readFileSync(f, 'utf8'));
            return d.registros_pipas || [];
        }
    } catch(e) {}
    return [];
}

async function guardarPipaDiesel(p) {
    if (!p) return false;
    const pId = String(p.id || `pipa_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`);

    try {
        const f = path.join(__dirname, 'diesel_local_cache.json');
        if (fs.existsSync(f)) {
            const cache = JSON.parse(fs.readFileSync(f, 'utf8'));
            cache.registros_pipas = cache.registros_pipas || [];
            cache.registros_pipas.unshift({ ...p, id: pId });
            fs.writeFileSync(f, JSON.stringify(cache, null, 2), 'utf8');
        }
    } catch(e) {}

    if (isPostgresAvailable && pool) {
        try {
            await pool.query(`
                INSERT INTO diesel_pipas (id, fecha, litros_recibidos, proveedor, factura, costo_total, notas, created_at)
                VALUES ($1, $2, $3, $4, $5, $6, $7, CURRENT_TIMESTAMP)
                ON CONFLICT (id) DO UPDATE SET
                    fecha = EXCLUDED.fecha,
                    litros_recibidos = EXCLUDED.litros_recibidos,
                    proveedor = EXCLUDED.proveedor,
                    factura = EXCLUDED.factura,
                    costo_total = EXCLUDED.costo_total,
                    notas = EXCLUDED.notas;
            `, [
                pId, p.fecha || new Date().toISOString().slice(0, 10), Number(p.litros_recibidos) || 0,
                p.proveedor || '', p.factura || '', Number(p.costo_total) || 0, p.notas || ''
            ]);
            return true;
        } catch(e) {
            console.error('[PostgreSQL Save Pipa Error]:', e.message);
        }
    }
    return true;
}

// ESTADÍSTICAS GERENCIALES DE DIÉSEL
async function obtenerStatsDiesel() {
    if (isPostgresAvailable && pool) {
        try {
            const res = await pool.query(`
                SELECT 
                    COUNT(*) as total_registros,
                    COALESCE(SUM(km_recorridos), 0) as km_totales,
                    COALESCE(SUM(real_total), 0) as litros_totales,
                    COALESCE(ROUND(AVG(NULLIF(rend_fisico, 0)), 2), 0) as rend_fisico_promedio,
                    COALESCE(ROUND(AVG(NULLIF(rend_ecm, 0)), 2), 0) as rend_ecm_promedio,
                    COALESCE(SUM(costo_diesel), 0) as costo_diesel_total,
                    COALESCE(SUM(litros_ralenti), 0) as litros_ralenti_total,
                    COALESCE(SUM(costo_ralenti), 0) as costo_ralenti_total,
                    COALESCE(SUM(bono), 0) as bono_total
                FROM diesel_registros_rendimiento;
            `);
            if (res.rows.length > 0) {
                const r = res.rows[0];
                return {
                    total_registros: parseInt(r.total_registros) || 0,
                    km_totales: Number(r.km_totales) || 0,
                    litros_totales: Number(r.litros_totales) || 0,
                    rend_fisico_promedio: Number(r.rend_fisico_promedio) || 0,
                    rend_ecm_promedio: Number(r.rend_ecm_promedio) || 0,
                    costo_diesel_total: Number(r.costo_diesel_total) || 0,
                    litros_ralenti_total: Number(r.litros_ralenti_total) || 0,
                    costo_ralenti_total: Number(r.costo_ralenti_total) || 0,
                    bono_total: Number(r.bono_total) || 0
                };
            }
        } catch (e) {
            console.warn('[PostgreSQL Stats Diesel Error]:', e.message);
        }
    }

    try {
        const f = path.join(__dirname, 'diesel_local_cache.json');
        if (fs.existsSync(f)) {
            const d = JSON.parse(fs.readFileSync(f, 'utf8'));
            const recs = d.registros_rendimiento || [];
            let km = 0, lit = 0, sumRendF = 0, countRendF = 0, sumRendE = 0, countRendE = 0, cDiesel = 0, litRal = 0, cRal = 0, bono = 0;
            recs.forEach(r => {
                km += Number(r.km_recorridos) || 0;
                lit += Number(r.real_total) || 0;
                const rf = Number(r.rend_fisico) || 0;
                if (rf > 0) { sumRendF += rf; countRendF++; }
                const re = Number(r.rend_ecm) || 0;
                if (re > 0) { sumRendE += re; countRendE++; }
                const cd = typeof r.costo_diesel === 'string' ? Number(r.costo_diesel.replace(/[^0-9.-]+/g, '')) || 0 : Number(r.costo_diesel) || 0;
                cDiesel += cd;
                litRal += Number(r.litros_ralenti) || 0;
                const cr = typeof r.costo_ralenti === 'string' ? Number(r.costo_ralenti.replace(/[^0-9.-]+/g, '')) || 0 : Number(r.costo_ralenti) || 0;
                cRal += cr;
                const b = typeof r.bono === 'string' ? Number(r.bono.replace(/[^0-9.-]+/g, '')) || 0 : Number(r.bono) || 0;
                bono += b;
            });
            return {
                total_registros: recs.length,
                km_totales: km,
                litros_totales: lit,
                rend_fisico_promedio: countRendF > 0 ? Math.round((sumRendF / countRendF) * 100) / 100 : 0,
                rend_ecm_promedio: countRendE > 0 ? Math.round((sumRendE / countRendE) * 100) / 100 : 0,
                costo_diesel_total: Math.round(cDiesel * 100) / 100,
                litros_ralenti_total: Math.round(litRal * 100) / 100,
                costo_ralenti_total: Math.round(cRal * 100) / 100,
                bono_total: Math.round(bono * 100) / 100
            };
        }
    } catch(e) {}

    return {
        total_registros: 0,
        km_totales: 0,
        litros_totales: 0,
        rend_fisico_promedio: 0,
        rend_ecm_promedio: 0,
        costo_diesel_total: 0,
        litros_ralenti_total: 0,
        costo_ralenti_total: 0,
        bono_total: 0
    };
}

// SINCRONIZACIÓN INVERSA (POSTGRESQL -> DISCO LOCAL PARA COMPATIBILIDAD 100%)
async function sincronizarTodoHaciaDisco(rootDir = __dirname) {
    if (!isPostgresAvailable || !pool) return false;
    try {
        const unidades = await obtenerUnidadesTaller();
        if (Array.isArray(unidades) && unidades.length > 0) {
            fs.writeFileSync(path.join(rootDir, 'taller_db.json'), JSON.stringify(unidades, null, 2), 'utf8');
            console.log(`[ERP Sync] ${unidades.length} unidades actualizadas desde PostgreSQL a disco local.`);
        }
        const remolques = await obtenerRemolquesTaller();
        if (Array.isArray(remolques) && remolques.length > 0) {
            fs.writeFileSync(path.join(rootDir, 'taller_remolques.json'), JSON.stringify(remolques, null, 2), 'utf8');
            console.log(`[ERP Sync] ${remolques.length} remolques actualizados desde PostgreSQL a disco local.`);
        }
        const rescates = await obtenerRescatesTaller();
        if (Array.isArray(rescates) && rescates.length > 0) {
            const limpios = rescates.filter(r => String(r.unidad).trim() !== '75');
            fs.writeFileSync(path.join(rootDir, 'taller_rescates.json'), JSON.stringify(limpios, null, 2), 'utf8');
            console.log(`[ERP Sync] ${limpios.length} rescates carreteros actualizados desde PostgreSQL a disco local.`);
        }
        return true;
    } catch(e) {
        console.error('[ERP Sync Error]:', e.message);
        return false;
    }
}

// Inicialización de arranque segura
async function init() {
    if (isInitialized) return;
    isInitialized = true;
    if (isPostgresAvailable && pool) {
        const ok = await inicializarEsquemaPostgres();
        if (ok) {
            await sembrarDatosInicialesSiVacio();
            await sincronizarTodoHaciaDisco();
        }
    }
}

// Ejecutar init al requerir
init().catch(err => console.error('[Database Init Error]:', err));

module.exports = {
    pool,
    isPostgresAvailable: () => isPostgresAvailable,
    inicializarEsquemaPostgres,
    sembrarDatosInicialesSiVacio,
    sincronizarTodoHaciaDisco,
    // Taller
    obtenerUnidadesTaller,
    guardarUnidadesTaller,
    obtenerRemolquesTaller,
    guardarRemolquesTaller,
    obtenerRescatesTaller,
    guardarRescatesTaller,
    guardarRescateTaller,
    eliminarRescateTaller,
    // Diésel & Rendimientos
    obtenerRegistrosDiesel,
    guardarRegistroDiesel,
    obtenerOperadoresDiesel,
    guardarOperadorDiesel,
    obtenerAutorizacionesDiesel,
    guardarAutorizacionDiesel,
    obtenerPipasDiesel,
    guardarPipaDiesel,
    obtenerStatsDiesel
};
