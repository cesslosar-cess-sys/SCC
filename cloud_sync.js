/**
 * CLOUD SYNC MODULE - CESS SISTEMA INTEGRAL
 * Sincronización en tiempo real y persistencia automática en Google Cloud Firestore
 * Proyecto: diesel-cess-40bbf
 */

const fs = require('fs');
const path = require('path');

let admin = null;
let getFirestore = null;
let FieldValue = null;
let db = null;
let isInitialized = false;

function initFirebase() {
  if (isInitialized) return db;

  try {
    admin = require('firebase-admin');
    const firestorePkg = require('firebase-admin/firestore');
    getFirestore = firestorePkg.getFirestore;
    FieldValue = firestorePkg.FieldValue;

    const candidates = [];

    // 1. Prioridad: Archivos locales certificados en disco (inmunes a errores de copiar/pegar en web)
    const possibleKeyPaths = [
      path.join(__dirname, 'firebase_key.json'),
      path.join(process.cwd(), 'firebase_key.json'),
      '/etc/secrets/firebase_key.json',
      '/etc/secrets/FIREBASE_SERVICE_ACCOUNT',
      path.join(__dirname, '..', 'firebase_key.json'),
      path.join(__dirname, 'almacen-app', 'firebase_key.json')
    ];

    for (const p of possibleKeyPaths) {
      if (fs.existsSync(p)) {
        try {
          const saData = JSON.parse(fs.readFileSync(p, 'utf8'));
          if (saData && saData.project_id) {
            candidates.push({ source: `archivo local (${path.basename(p)})`, data: saData });
          }
        } catch(e) {}
      }
    }

    // 2. Variable de Entorno (Render / Hosting)
    if (process.env.FIREBASE_SERVICE_ACCOUNT) {
      try {
        let envVal = process.env.FIREBASE_SERVICE_ACCOUNT.trim();
        if (fs.existsSync(envVal)) {
          const saData = JSON.parse(fs.readFileSync(envVal, 'utf8'));
          if (saData && saData.project_id) {
            candidates.push({ source: `ruta en env (${envVal})`, data: saData });
          }
        } else {
          if ((envVal.startsWith('"') && envVal.endsWith('"')) || (envVal.startsWith("'") && envVal.endsWith("'"))) {
            envVal = envVal.slice(1, -1).trim();
          }
          let saData = null;
          if (envVal.startsWith('{')) {
            saData = JSON.parse(envVal);
          } else {
            const cleanB64 = envVal.replace(/\s+/g, '');
            const decoded = Buffer.from(cleanB64, 'base64').toString('utf8');
            saData = JSON.parse(decoded);
          }
          if (saData && saData.project_id) {
            candidates.push({ source: 'process.env.FIREBASE_SERVICE_ACCOUNT', data: saData });
          }
        }
      } catch (errEnv) {
        console.warn('️ [CloudSync] Aviso al leer env variable:', errEnv.message);
      }
    }

    // Probar candidatos hasta lograr conexión exitosa
    for (const cand of candidates) {
      try {
        const sa = cand.data;
        if (sa.private_key) {
          sa.private_key = sa.private_key.replace(/\\n/g, '\n');
        }

        if (admin.getApps().length === 0) {
          admin.initializeApp({
            credential: admin.cert(sa)
          });
        }
        db = getFirestore();
        isInitialized = true;
        console.log(` [CloudSync] Conectado exitosamente a Google Cloud Firestore (${sa.project_id}) vía ${cand.source}`);
        return db;
      } catch (candErr) {
        console.warn(`️ [CloudSync] Falló candidato ${cand.source}:`, candErr.message);
      }
    }

    console.warn('️ [CloudSync] No se logró conectar a Firebase. Operando en modo disco local.');
    return null;
  } catch (err) {
    console.error('️ [CloudSync] Error inicializando Firebase Admin:', err.message);
    return null;
  }
}

const COLL = 'cess_cloud_database';
const CHUNK_SIZE = 150;

function isCloudEnabled() {
  if (!isInitialized) initFirebase();
  return Boolean(db);
}

/**
 * PULL de datos de Almacén desde Firestore
 */
async function pullAlmacenData() {
  if (!isCloudEnabled()) return null;

  try {
    console.log(' [CloudSync] Descargando datos de Almacén desde Firestore...');
    const catDoc = await db.collection(COLL).doc('almacen_catalog').get();
    const metaDoc = await db.collection(COLL).doc('almacen_movimientos_meta').get();
    const extraMetaDoc = await db.collection(COLL).doc('almacen_metadata').get();

    if (!catDoc.exists && !metaDoc.exists) {
      console.warn('️ [CloudSync] No se encontraron documentos de Almacén en Firestore.');
      return null;
    }

    const catalog = catDoc.exists ? JSON.parse(catDoc.data().json_data || '{}') : {};
    let movimientos = [];

    if (metaDoc.exists) {
      const meta = metaDoc.data();
      const totalChunks = meta.total_chunks || 1;
      const chunkPromises = [];
      for (let i = 0; i < totalChunks; i++) {
        chunkPromises.push(db.collection(COLL).doc(`almacen_movs_chunk_${i}`).get());
      }
      const chunkSnaps = await Promise.all(chunkPromises);
      for (const snap of chunkSnaps) {
        if (snap.exists) {
          const chunkData = JSON.parse(snap.data().json_data || '[]');
          if (Array.isArray(chunkData)) {
            movimientos = movimientos.concat(chunkData);
          }
        }
      }
    }

    let extraMeta = {};
    if (extraMetaDoc.exists) {
      extraMeta = JSON.parse(extraMetaDoc.data().json_data || '{}');
    }

    const reconstructed = {
      activeSheet: extraMeta.activeSheet || '',
      lastSync: extraMeta.lastSync || new Date().toISOString(),
      catalog: catalog,
      movimientos: movimientos,
      facturas_procesadas: extraMeta.facturas_procesadas || [],
      usuarios: extraMeta.usuarios || [],
      memoria_ia: extraMeta.memoria_ia || [],
      folios_taller: extraMeta.folios_taller || [],
      almacenes: extraMeta.almacenes || [],
      ordenes_compra: extraMeta.ordenes_compra || [],
      config: extraMeta.config || {},
      facturas_registradas: extraMeta.facturas_registradas || []
    };

    console.log(` [CloudSync] Almacén descargado de la nube: ${Object.keys(catalog).length} refacciones, ${movimientos.length} movimientos.`);
    return reconstructed;
  } catch (err) {
    console.error(' [CloudSync] Error descargando Almacén desde Firestore:', err.message);
    return null;
  }
}

/**
 * PUSH de datos de Almacén a Firestore
 */
let pushAlmacenTimer = null;
async function pushAlmacenData(data) {
  if (!isCloudEnabled() || !data) return;

  // Debounce para no saturar escrituras si se hacen múltiples cambios consecutivos
  if (pushAlmacenTimer) clearTimeout(pushAlmacenTimer);

  return new Promise((resolve) => {
    pushAlmacenTimer = setTimeout(async () => {
      try {
        const catalog = data.catalog || {};
        const movimientos = data.movimientos || [];
        const catalogJson = JSON.stringify(catalog);

        const totalMovs = movimientos.length;
        const totalChunks = Math.ceil(totalMovs / CHUNK_SIZE) || 1;

        const batch = db.batch();

        // 1. Catalog
        const catRef = db.collection(COLL).doc('almacen_catalog');
        batch.set(catRef, {
          count: Object.keys(catalog).length,
          json_data: catalogJson,
          updatedAt: FieldValue.serverTimestamp()
        });

        // 2. Chunks de Movimientos
        for (let i = 0; i < totalChunks; i++) {
          const chunkItems = movimientos.slice(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
          const chunkRef = db.collection(COLL).doc(`almacen_movs_chunk_${i}`);
          batch.set(chunkRef, {
            chunk_index: i,
            count: chunkItems.length,
            json_data: JSON.stringify(chunkItems),
            updatedAt: FieldValue.serverTimestamp()
          });
        }

        // 3. Meta movimientos
        const metaRef = db.collection(COLL).doc('almacen_movimientos_meta');
        batch.set(metaRef, {
          total_count: totalMovs,
          chunk_size: CHUNK_SIZE,
          total_chunks: totalChunks,
          updatedAt: FieldValue.serverTimestamp()
        });

        // 4. Metadata
        const metaKeys = ['activeSheet', 'lastSync', 'facturas_procesadas', 'usuarios', 'memoria_ia', 'folios_taller', 'almacenes', 'ordenes_compra', 'config', 'facturas_registradas'];
        const metadataPayload = {};
        for (const k of metaKeys) {
          if (data[k] !== undefined) metadataPayload[k] = data[k];
        }
        const extraMetaRef = db.collection(COLL).doc('almacen_metadata');
        batch.set(extraMetaRef, {
          json_data: JSON.stringify(metadataPayload),
          updatedAt: FieldValue.serverTimestamp()
        });

        await batch.commit();
        console.log(`️ [CloudSync] Almacén guardado en la nube: ${Object.keys(catalog).length} partes, ${totalMovs} movimientos.`);
        resolve(true);
      } catch (err) {
        console.error(' [CloudSync] Error subiendo Almacén a Firestore:', err.message);
        resolve(false);
      }
    }, 250);
  });
}

/**
 * PULL de datos de Taller desde Firestore
 */
async function pullTallerData() {
  if (!isCloudEnabled()) return null;

  try {
    console.log(' [CloudSync] Descargando datos de Taller desde Firestore...');
    const [unidadesSnap, archivadasSnap, bitacorasSnap, remolquesSnap, rescatesSnap, colsSnap, usuariosSnap] = await Promise.all([
      db.collection(COLL).doc('taller_unidades').get(),
      db.collection(COLL).doc('taller_archivadas').get(),
      db.collection(COLL).doc('taller_bitacoras').get(),
      db.collection(COLL).doc('taller_remolques').get(),
      db.collection(COLL).doc('taller_rescates').get(),
      db.collection(COLL).doc('taller_columnas').get(),
      db.collection(COLL).doc('taller_usuarios').get()
    ]);

    const result = {
      unidades: unidadesSnap.exists ? JSON.parse(unidadesSnap.data().json_data || '[]') : null,
      archivadas: archivadasSnap.exists ? JSON.parse(archivadasSnap.data().json_data || '[]') : null,
      bitacoras: bitacorasSnap.exists ? JSON.parse(bitacorasSnap.data().json_data || '[]') : null,
      remolques: remolquesSnap.exists ? JSON.parse(remolquesSnap.data().json_data || '[]') : null,
      rescates: rescatesSnap.exists ? JSON.parse(rescatesSnap.data().json_data || '[]') : null,
      columnas: colsSnap.exists ? JSON.parse(colsSnap.data().json_data || '{}') : null,
      usuarios: usuariosSnap.exists ? JSON.parse(usuariosSnap.data().json_data || '[]') : null
    };

    console.log(` [CloudSync] Taller descargado de la nube: ${Array.isArray(result.unidades) ? result.unidades.length : 0} unidades activas.`);
    return result;
  } catch (err) {
    console.error(' [CloudSync] Error descargando Taller desde Firestore:', err.message);
    return null;
  }
}

/**
 * PUSH de unidades de Taller a Firestore
 */
let pushTallerTimer = null;
async function pushTallerData(unidades) {
  if (!isCloudEnabled() || !Array.isArray(unidades)) return;

  if (pushTallerTimer) clearTimeout(pushTallerTimer);

  return new Promise((resolve) => {
    pushTallerTimer = setTimeout(async () => {
      try {
        await db.collection(COLL).doc('taller_unidades').set({
          count: unidades.length,
          json_data: JSON.stringify(unidades),
          updatedAt: FieldValue.serverTimestamp()
        });
        console.log(`️ [CloudSync] Taller guardado en la nube: ${unidades.length} unidades.`);
        resolve(true);
      } catch (err) {
        console.error(' [CloudSync] Error subiendo unidades de Taller a Firestore:', err.message);
        resolve(false);
      }
    }, 250);
  });
}

/**
 * PUSH de archivadas de Taller a Firestore
 */
async function pushTallerArchivadas(archivadas) {
  if (!isCloudEnabled() || !Array.isArray(archivadas)) return;
  try {
    await db.collection(COLL).doc('taller_archivadas').set({
      count: archivadas.length,
      json_data: JSON.stringify(archivadas),
      updatedAt: FieldValue.serverTimestamp()
    });
  } catch (err) {
    console.error(' [CloudSync] Error subiendo archivadas a Firestore:', err.message);
  }
}

/**
 * PUSH de Cajas y Remolques a Firestore
 */
async function pushRemolquesData(remolques) {
  if (!isCloudEnabled() || !Array.isArray(remolques)) return;
  try {
    await db.collection(COLL).doc('taller_remolques').set({
      count: remolques.length,
      json_data: JSON.stringify(remolques),
      updatedAt: FieldValue.serverTimestamp()
    });
    console.log(`[CloudSync] Remolques guardados en la nube: ${remolques.length} unidades.`);
  } catch (err) {
    console.error('[CloudSync] Error subiendo remolques a Firestore:', err.message);
  }
}

/**
 * Sincronización Inicial: Compara y descarga de Firestore a disco local si la nube tiene datos más completos
 */
async function syncAllFromCloudToDisk(baseDir) {
  if (!isCloudEnabled()) return false;

  const rootDir = baseDir || __dirname;
  console.log(' [CloudSync] Verificando sincronización de arranque con Firestore...');

  try {
    // 1. Sincronizar Almacén
    const cloudAlm = await pullAlmacenData();
    if (cloudAlm && cloudAlm.movimientos) {
      const almFile = path.join(rootDir, 'almacen_datos.json');
      let localMovsCount = 0;
      if (fs.existsSync(almFile)) {
        try {
          const local = JSON.parse(fs.readFileSync(almFile, 'utf8'));
          localMovsCount = (local.movimientos || []).length;
        } catch (e) {}
      }

      // Si la nube tiene más o iguales movimientos, hidratar disco local
      if (cloudAlm.movimientos.length >= localMovsCount) {
        const jsonStr = JSON.stringify(cloudAlm, null, 2);
        fs.writeFileSync(almFile, jsonStr, 'utf8');
        // También actualizar copias en subdirectorios si existen
        [
          path.join(rootDir, 'almacen-app', 'almacen_datos.json'),
          path.join(rootDir, 'taller-cloud-deploy', 'almacen_datos.json')
        ].forEach(target => {
          try {
            if (fs.existsSync(path.dirname(target))) {
              fs.writeFileSync(target, jsonStr, 'utf8');
            }
          } catch(e) {}
        });
        console.log(` [CloudSync] Almacén hidratado en disco local con ${cloudAlm.movimientos.length} movimientos.`);
      }
    }

    // 2. Sincronizar Taller
    const cloudTaller = await pullTallerData();
    if (cloudTaller) {
      if (Array.isArray(cloudTaller.unidades)) {
        const tallerFile = path.join(rootDir, 'taller_db.json');
        fs.writeFileSync(tallerFile, JSON.stringify(cloudTaller.unidades, null, 2), 'utf8');
        console.log(` [CloudSync] Taller hidratado en disco local con ${cloudTaller.unidades.length} unidades activas.`);
      }
      if (Array.isArray(cloudTaller.archivadas)) {
        fs.writeFileSync(path.join(rootDir, 'taller_unidades_archivadas.json'), JSON.stringify(cloudTaller.archivadas, null, 2), 'utf8');
      }
      if (Array.isArray(cloudTaller.bitacoras)) {
        fs.writeFileSync(path.join(rootDir, 'taller_bitacoras_historico.json'), JSON.stringify(cloudTaller.bitacoras, null, 2), 'utf8');
      }
      if (Array.isArray(cloudTaller.remolques)) {
        fs.writeFileSync(path.join(rootDir, 'taller_remolques.json'), JSON.stringify(cloudTaller.remolques, null, 2), 'utf8');
      }
      if (Array.isArray(cloudTaller.rescates)) {
        fs.writeFileSync(path.join(rootDir, 'taller_rescates.json'), JSON.stringify(cloudTaller.rescates, null, 2), 'utf8');
      }
      if (cloudTaller.columnas) {
        fs.writeFileSync(path.join(rootDir, 'taller_config_columnas.json'), JSON.stringify(cloudTaller.columnas, null, 2), 'utf8');
      }
      if (Array.isArray(cloudTaller.usuarios) && cloudTaller.usuarios.length > 0) {
        fs.writeFileSync(path.join(rootDir, 'taller_usuarios.json'), JSON.stringify(cloudTaller.usuarios, null, 2), 'utf8');
        console.log(` [CloudSync] Roles y Usuarios de Taller hidratados desde Firestore.`);
      }
    }

    return true;
  } catch (err) {
    console.error(' [CloudSync] Error en syncAllFromCloudToDisk:', err.message);
    return false;
  }
}

/**
 * PUSH de usuarios y roles a Firestore
 */
async function pushTallerUsuarios(usuarios) {
  if (!isCloudEnabled() || !Array.isArray(usuarios)) return;
  try {
    await db.collection(COLL).doc('taller_usuarios').set({
      count: usuarios.length,
      json_data: JSON.stringify(usuarios),
      updatedAt: FieldValue.serverTimestamp()
    });
    console.log(`️ [CloudSync] Usuarios y Roles guardados en Firestore: ${usuarios.length} cuentas.`);
  } catch (err) {
    console.error(' [CloudSync] Error subiendo usuarios a Firestore:', err.message);
  }
}

function reloadFirebase() {
  isInitialized = false;
  db = null;
  return initFirebase();
}

// Inicializar de inmediato al importar
initFirebase();

module.exports = {
  isCloudEnabled,
  reloadFirebase,
  pullAlmacenData,
  pushAlmacenData,
  pullTallerData,
  pushTallerData,
  pushTallerArchivadas,
  pushTallerUsuarios,
  pushRemolquesData,
  syncAllFromCloudToDisk
};
