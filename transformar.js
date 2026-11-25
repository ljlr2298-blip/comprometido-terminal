const xlsx = require('xlsx');
const Database = require('better-sqlite3');
const path = require('path');
const fs = require('fs');

const DATA_DIR = path.join(__dirname, 'data');
const BASES_DIR = path.join(__dirname, 'bases');
const INFO_BASE_PATH = path.join(__dirname, 'info base.xlsx');
const DB_PATH = path.join(__dirname, 'comparacion.db');

if (!fs.existsSync(BASES_DIR)) fs.mkdirSync(BASES_DIR);

// ===== NORMALIZADOR DE CODIGOS =====
function norm(v, len) {
    let s = String(v === null || v === undefined ? '' : v).trim();
    if (s === '') return '';
    if (/^\d+(\.0+)?$/.test(s)) s = String(Math.floor(Number(s)));
    return s.padStart(len, '0');
}

// ===== EXTRAER FECHA DEL NOMBRE DEL ARCHIVO =====
// Detecta: DD-MM-AAAA o DD-MM-AA
function extraerFechaDelNombre(nombreArchivo) {
    // Primero intenta DD-MM-AAAA (8 dígitos)
    let match = nombreArchivo.match(/(\d{2})-(\d{2})-(\d{4})/);
    if (match) {
        const [, dd, mm, yyyy] = match;
        const d = new Date(`${yyyy}-${mm}-${dd}T00:00:00`);
        if (!isNaN(d.getTime())) return `${yyyy}-${mm}-${dd}`;
    }
    
    // Si no, intenta DD-MM-AA (6 dígitos) y asume 2000+
    match = nombreArchivo.match(/(\d{2})-(\d{2})-(\d{2})/);
    if (match) {
        const [, dd, mm, aa] = match;
        const yyyy = `20${aa}`;
        const d = new Date(`${yyyy}-${mm}-${dd}T00:00:00`);
        if (!isNaN(d.getTime())) return `${yyyy}-${mm}-${dd}`;
    }
    
    return null;
}

// ===== 1. LISTAR TODOS LOS ARCHIVOS CRUDO EN data/ =====
const archivosData = fs.readdirSync(DATA_DIR).filter(f =>
    f.endsWith('.xlsx') && !f.startsWith('~') &&
    f !== 'periodo1.xlsx' && f !== 'periodo2.xlsx'
);

if (archivosData.length === 0) {
    console.error('❌ No hay archivos .xlsx en la carpeta data/');
    process.exit(1);
}

console.log(`📂 Archivos detectados en data/: ${archivosData.length}\n`);

// ===== 2. CARGAR Y LIMPIAR CATALOGO =====
console.log('📚 Cargando catálogo info base.xlsx...');
const wbInfo = xlsx.readFile(INFO_BASE_PATH);
const filasCatalogo = [];
wbInfo.SheetNames.forEach(nombre => {
    const datos = xlsx.utils.sheet_to_json(wbInfo.Sheets[nombre], { defval: '' });
    filasCatalogo.push(...datos);
});

// FILL DOWN en codigos Y descripciones
const columnasFill = [
    'UP', 'UR', 'PPAR', 'PP', 'COM', 'PARTIDA', 'DEST', 'FF', 'SFF',
    'DEPENDENCIA', 'UNIDAD RESPONSABLE', 'ASIGNACION DE RECURSOS',
    'PROGRAMA PRESUPUESTARIO', 'COMPONENTE', 'OBJETO DEL GASTO', 'DESTINO'
];
const ultimoValor = {};
filasCatalogo.forEach(fila => {
    columnasFill.forEach(col => {
        const val = String(fila[col] ?? '').trim();
        if (val !== '') {
            ultimoValor[col] = val;
        } else if (ultimoValor[col] !== undefined) {
            fila[col] = ultimoValor[col];
        }
    });
});

// Mapas con claves NORMALIZADAS
const mapDependencia = new Map();
const mapConceptoUR = new Map();
const mapAsignacion = new Map();
const mapDescripPP = new Map();
const mapDescripComp = new Map();
const mapConceptoPartida = new Map();
const mapDescripDest = new Map();

filasCatalogo.forEach(f => {
    const up = norm(f.UP, 2);
    const ur = norm(f.UR, 3);
    const ppar = String(f.PPAR ?? '').trim();
    const pp = String(f.PP ?? '').trim();
    const com = String(f.COM ?? '').trim();
    const part = norm(f.PARTIDA, 4);
    const dest = String(f.DEST ?? '').trim();

    if (up && f.DEPENDENCIA) mapDependencia.set(up, f.DEPENDENCIA);
    if (up && ur && f['UNIDAD RESPONSABLE']) mapConceptoUR.set(`${up}-${ur}`, f['UNIDAD RESPONSABLE']);
    if (ppar && f['ASIGNACION DE RECURSOS']) mapAsignacion.set(ppar, f['ASIGNACION DE RECURSOS']);
    if (pp && f['PROGRAMA PRESUPUESTARIO']) mapDescripPP.set(pp, f['PROGRAMA PRESUPUESTARIO']);
    if (up && ur && pp && com && f['COMPONENTE']) mapDescripComp.set(`${up}-${ur}-${pp}-${com}`, f['COMPONENTE']);
    if (part && f['OBJETO DEL GASTO']) mapConceptoPartida.set(part, f['OBJETO DEL GASTO']);
    if (up && ur && pp && com && part && dest && f['DESTINO']) {
        mapDescripDest.set(`${up}-${ur}-${pp}-${com}-${part}-${dest}`, f['DESTINO']);
    }
});

console.log(`✅ Catálogo: ${filasCatalogo.length} filas`);
console.log(`   UR=${mapConceptoUR.size}, Partidas=${mapConceptoPartida.size}, Dest=${mapDescripDest.size}\n`);

// ===== 3. FUNCION DE EXTRACCION DE CLAVE =====
function extraerDeClave(cve) {
    const s = String(cve || '').trim();
    if (s.length < 41) return null;
    return {
        up: s.substring(5, 7),
        ur: s.substring(7, 10),
        ar: s.substring(20, 21),
        pp: s.substring(21, 24),
        comp: s.substring(24, 26),
        part: s.substring(26, 30),
        dest: s.substring(30, 32),
        ff: s.substring(32, 34),
        sff: s.substring(34, 41),
        capitulo: s.substring(26, 27) + '000'
    };
}

// ===== 4. PROCESAR CADA ARCHIVO =====
const db = new Database(DB_PATH);
db.pragma('journal_mode = WAL');

const tablasExistentes = new Set(
    db.prepare("SELECT name FROM sqlite_master WHERE type='table' AND name LIKE 'base_%'").all().map(t => t.name)
);

let procesados = 0, saltados = 0, errores = 0;

archivosData.forEach(nombreArchivo => {
    const rutaArchivo = path.join(DATA_DIR, nombreArchivo);
    const fechaArchivo = extraerFechaDelNombre(nombreArchivo);

    console.log(`──────────────────────────────────────────────`);
    console.log(`📄 ${nombreArchivo}`);

    if (!fechaArchivo) {
        console.log(`   ⚠️  No se detectó fecha en el nombre. SE OMITE.`);
        console.log(`   💡 Renómbralo para que tenga DD-MM-AA o DD-MM-AAAA, ej: "Comprometido_23-11-26.xlsx"`);
        errores++;
        return;
    }

    const nombreTabla = `base_${fechaArchivo.replace(/-/g, '_')}`;
    console.log(`   📅 Fecha detectada: ${fechaArchivo}`);

    if (tablasExistentes.has(nombreTabla)) {
        console.log(`   ⏭️  Ya existe en la base de datos. SE OMITE.`);
        saltados++;
        return;
    }

    try {
        const wbCrudo = xlsx.readFile(rutaArchivo);
        const datosCrudos = xlsx.utils.sheet_to_json(wbCrudo.Sheets[wbCrudo.SheetNames[0]], { defval: '' });
        console.log(`   📥 ${datosCrudos.length} filas leídas`);

        const filasTransformadas = [];
        let sinClave = 0;

        datosCrudos.forEach(fila => {
            const codes = extraerDeClave(fila.cvepresupuestal);
            if (!codes) { sinClave++; return; }
            const { up, ur, ar, pp, comp, part, dest, ff, sff, capitulo } = codes;

            filasTransformadas.push({
                'Justificación': fila.justificacion || '',
                'Núm Justificación': fila.numjustificacion || '',
                'Nombre Fiscal': fila.nombrefiscal || '',
                'Concepto': fila.concepto || '',
                'Clave Presupuestal': fila.cvepresupuestal || '',
                'Importe Original': Number(fila.importeoriginal) || 0,
                'Importe Pagado': Number(fila.importepagado) || 0,
                'Saldo a Pagar': Number(fila.saldoapagar) || 0,
                'Estatus': fila.estatus || '',
                'Idu': String(fila.idur || ''),
                'Dependencia (código)': String(fila.dependencia || ''),
                'Fecha Emis': String(fila.fechaemision || ''),
                'UP': up, 'UR': ur, 'AR': ar, 'PP': pp, 'Componente': comp,
                'Partida': part, 'Destino': dest, 'FF': ff, 'SFF': sff, 'Capítulo': capitulo,
                'Dependencia': mapDependencia.get(up) || '',
                'Concepto UR': mapConceptoUR.get(`${up}-${ur}`) || '',
                'Asignación de Recurso': mapAsignacion.get(ar) || '',
                'Descrip PP': mapDescripPP.get(pp) || '',
                'Descrip Componente': mapDescripComp.get(`${up}-${ur}-${pp}-${comp}`) || '',
                'Concepto Partida': mapConceptoPartida.get(part) || '',
                'Descrip. Destino': mapDescripDest.get(`${up}-${ur}-${pp}-${comp}-${part}-${dest}`) || ''
            });
        });

        console.log(`   ⚙️  ${filasTransformadas.length} transformadas (${sinClave} sin clave válida)`);

        db.exec(`CREATE TABLE ${nombreTabla} (
            clave_unica TEXT PRIMARY KEY,
            justificacion TEXT, num_justificacion TEXT, nombre_fiscal TEXT, concepto TEXT,
            cve_presupuestal TEXT, importe_original REAL, importe_pagado REAL, saldo_a_pagar REAL,
            estatus TEXT, idu TEXT, dependencia_cod TEXT, fecha_emision TEXT,
            up TEXT, ur TEXT, ar TEXT, pp TEXT, componente TEXT, partida TEXT,
            destino TEXT, ff TEXT, sff TEXT, capitulo TEXT,
            dependencia TEXT, concepto_ur TEXT, asignacion_recurso TEXT,
            descrip_pp TEXT, descrip_componente TEXT, concepto_partida TEXT, descrip_destino TEXT
        )`);

        const stmt = db.prepare(`INSERT OR REPLACE INTO ${nombreTabla} VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`);
        const insertMany = db.transaction((rows) => { for (const r of rows) stmt.run(...r); });

        const filasParaDB = filasTransformadas.map(f => [
            `${f['Núm Justificación']}-${f['Clave Presupuestal']}`,
            f['Justificación'], f['Núm Justificación'], f['Nombre Fiscal'], f['Concepto'],
            f['Clave Presupuestal'], f['Importe Original'], f['Importe Pagado'], f['Saldo a Pagar'],
            f['Estatus'], f['Idu'], f['Dependencia (código)'], f['Fecha Emis'],
            f['UP'], f['UR'], f['AR'], f['PP'], f['Componente'], f['Partida'],
            f['Destino'], f['FF'], f['SFF'], f['Capítulo'],
            f['Dependencia'], f['Concepto UR'], f['Asignación de Recurso'],
            f['Descrip PP'], f['Descrip Componente'], f['Concepto Partida'], f['Descrip. Destino']
        ]);

        insertMany(filasParaDB);
        console.log(`   💾 Guardado en SQLite: '${nombreTabla}' (${filasParaDB.length} registros)`);

        const ws = xlsx.utils.json_to_sheet(filasTransformadas);
        const wb = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(wb, ws, 'Base Comprometido');
        const excelPath = path.join(BASES_DIR, `Base_Comprometido_${fechaArchivo}.xlsx`);
        xlsx.writeFile(wb, excelPath);
        console.log(`   📁 Excel: ${path.basename(excelPath)}`);

        procesados++;
    } catch (err) {
        console.log(`   ❌ ERROR: ${err.message}`);
        errores++;
    }
});

db.close();

console.log(`\n══════════════════════════════════════════`);
console.log(`🎉 RESUMEN FINAL`);
console.log(`   ✅ Procesados: ${procesados}`);
console.log(`   ⏭️  Saltados (ya existían): ${saltados}`);
console.log(`   ❌ Con error/sin fecha: ${errores}`);
console.log(`══════════════════════════════════════════\n`);