// ═══════════════════════════════════════════════════════════
// COMPROMETIDO TERMINAL — Servidor (Supabase + transformar.js)
// ═══════════════════════════════════════════════════════════
require('dotenv').config();
const express = require('express');
const { createClient } = require('@supabase/supabase-js');
const xlsx = require('xlsx');
const multer = require('multer');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 8080;

// ── Supabase ──────────────────────────────────────────────
const supabaseUrl = process.env.SUPABASE_URL;
const supabaseKey = process.env.SUPABASE_KEY;
const supabase = createClient(supabaseUrl, supabaseKey);

// ── Uploads ───────────────────────────────────────────────
const upload = multer({ dest: 'uploads/' });

// ── Static ────────────────────────────────────────────────
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

// ═══════════════════════════════════════════════════════════
// FUNCIONES AUXILIARES
// ═══════════════════════════════════════════════════════════

function validaFecha(f) {
    return /^\d{4}-\d{2}-\d{2}$/.test(f);
}

function norm(v, len) {
    let s = String(v === null || v === undefined ? '' : v).trim();
    if (s === '') return '';
    if (/^\d+(\.0+)?$/.test(s)) s = String(Math.floor(Number(s)));
    return s.padStart(len, '0');
}

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

function excelDateToISO(serial) {
    const n = Number(serial);
    if (isNaN(n) || n < 1) return String(serial || '').trim();
    const epoch = new Date(Date.UTC(1899, 11, 30));
    const d = new Date(epoch.getTime() + n * 86400000);
    if (isNaN(d.getTime())) return String(serial || '').trim();
    const yyyy = d.getUTCFullYear();
    const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
    const dd = String(d.getUTCDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
}

// ── Catálogo ──────────────────────────────────────────────
const INFO_BASE_PATH = path.join(__dirname, 'info base.xlsx');
let cacheCatalogo = null;

async function cargarCatalogo() {
    if (cacheCatalogo) return cacheCatalogo;
    if (!fs.existsSync(INFO_BASE_PATH)) {
        throw new Error('No se encuentra info base.xlsx en el servidor');
    }

    const wbInfo = xlsx.readFile(INFO_BASE_PATH);
    const filasCatalogo = [];
    wbInfo.SheetNames.forEach(nombre => {
        const datos = xlsx.utils.sheet_to_json(wbInfo.Sheets[nombre], { defval: '' });
        filasCatalogo.push(...datos);
    });

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

    cacheCatalogo = {
        mapDependencia, mapConceptoUR, mapAsignacion,
        mapDescripPP, mapDescripComp, mapConceptoPartida, mapDescripDest
    };
    console.log(`📚 Catálogo cargado: UR=${mapConceptoUR.size}, Partidas=${mapConceptoPartida.size}, Dest=${mapDescripDest.size}`);
    return cacheCatalogo;
}

// ═══════════════════════════════════════════════════════════
// ENDPOINTS
// ═══════════════════════════════════════════════════════════

// 1. Lista de fechas disponibles
app.get('/api/bases', async (req, res) => {
    const { data, error } = await supabase
        .from('comprometidos')
        .select('fecha')
        .order('fecha', { ascending: false });
    if (error) return res.status(500).json({ error: error.message });
    const fechasUnicas = [...new Set(data.map(item => item.fecha))];
    res.json(fechasUnicas);
});

// 2. UP/UR únicos de una fecha
app.get('/api/upur/:fecha', async (req, res) => {
    if (!validaFecha(req.params.fecha)) return res.status(400).json({ error: 'Fecha inválida' });
    const { data, error } = await supabase
        .from('comprometidos')
        .select('up, ur, concepto_ur, dependencia')
        .eq('fecha', req.params.fecha)
        .not('up', 'is', null)
        .not('ur', 'is', null)
        .range(0, 100000);
    if (error) return res.status(500).json({ error: error.message });

    const unique = [];
    const seen = new Set();
    data.forEach(row => {
        const key = `${row.up}-${row.ur}`;
        if (!seen.has(key)) {
            seen.add(key);
            unique.push({ up: row.up, ur: row.ur, concepto_ur: row.concepto_ur, dependencia: row.dependencia });
        }
    });
    res.json(unique.sort((a, b) => a.up.localeCompare(b.up) || a.ur.localeCompare(b.ur)));
});

// 3. Justificaciones únicas
app.get('/api/justificaciones/:fecha', async (req, res) => {
    if (!validaFecha(req.params.fecha)) return res.status(400).json({ error: 'Fecha inválida' });
    const { data, error } = await supabase
        .from('comprometidos')
        .select('justificacion')
        .eq('fecha', req.params.fecha)
        .not('justificacion', 'is', null)
        .range(0, 100000);
    if (error) return res.status(500).json({ error: error.message });
    const unique = [...new Set(data.map(r => r.justificacion))].filter(Boolean).sort();
    res.json(unique.map(j => ({ justificacion: j })));
});

// 4. Partidas únicas
app.get('/api/partidas/:fecha', async (req, res) => {
    if (!validaFecha(req.params.fecha)) return res.status(400).json({ error: 'Fecha inválida' });
    const { data, error } = await supabase
        .from('comprometidos')
        .select('partida, concepto_partida')
        .eq('fecha', req.params.fecha)
        .not('partida', 'is', null)
        .range(0, 100000);
    if (error) return res.status(500).json({ error: error.message });

    const unique = [];
    const seen = new Set();
    data.forEach(row => {
        if (!seen.has(row.partida)) {
            seen.add(row.partida);
            unique.push({ partida: row.partida, concepto_partida: row.concepto_partida });
        }
    });
    res.json(unique.sort((a, b) => a.partida.localeCompare(b.partida)));
});

// ═══════════════════════════════════════════════════════════
// 5. DRILL-DOWN: Nivel 1 — Dependencias (UP+UR)
//    Versión sin RPC: loop con bloques de 1000
// ═══════════════════════════════════════════════════════════
app.get('/api/dependencias/:fecha', async (req, res) => {
    if (!validaFecha(req.params.fecha)) return res.status(400).json({ error: 'Fecha inválida' });

    try {
        let todosLosDatos = [];
        let offset = 0;
        const batchSize = 1000;
        let hayMas = true;

        while (hayMas) {
            const { data, error } = await supabase
                .from('comprometidos')
                .select('up, ur, concepto_ur, dependencia, importe_original, importe_pagado, saldo_a_pagar')
                .eq('fecha', req.params.fecha)
                .range(offset, offset + batchSize - 1);

            if (error) throw error;
            if (!data || data.length === 0) { hayMas = false; break; }

            todosLosDatos.push(...data);
            if (data.length < batchSize) { hayMas = false; }
            else { offset += batchSize; }
        }

        const mapa = new Map();
        let totalRegistros = 0;
        let totalOriginal = 0;
        let totalPagado = 0;
        let totalSaldo = 0;

        todosLosDatos.forEach(row => {
            const key = `${row.up}-${row.ur}`;
            if (!mapa.has(key)) {
                mapa.set(key, {
                    up: row.up,
                    ur: row.ur,
                    concepto_ur: row.concepto_ur || '',
                    dependencia: row.dependencia || '',
                    registros: 0,
                    importe_original: 0,
                    importe_pagado: 0,
                    saldo_a_pagar: 0
                });
            }
            const item = mapa.get(key);
            item.registros += 1;
            item.importe_original += parseFloat(row.importe_original) || 0;
            item.importe_pagado += parseFloat(row.importe_pagado) || 0;
            item.saldo_a_pagar += parseFloat(row.saldo_a_pagar) || 0;

            totalRegistros += 1;
            totalOriginal += parseFloat(row.importe_original) || 0;
            totalPagado += parseFloat(row.importe_pagado) || 0;
            totalSaldo += parseFloat(row.saldo_a_pagar) || 0;
        });

        const dependencias = Array.from(mapa.values()).sort((a, b) => {
            if (a.up !== b.up) return a.up.localeCompare(b.up);
            return a.ur.localeCompare(b.ur);
        });

        res.json({
            dependencias,
            totales: {
                total_registros: totalRegistros,
                total_original: totalOriginal,
                total_pagado: totalPagado,
                total_saldo: totalSaldo
            }
        });
    } catch (error) {
        console.error('❌ Error en /api/dependencias:', error);
        res.status(500).json({ error: error.message });
    }
});

// ═══════════════════════════════════════════════════════════
// 6. DRILL-DOWN: Nivel 2 — Registros de una dependencia
// ═══════════════════════════════════════════════════════════
app.get('/api/base/:fecha', async (req, res) => {
    if (!validaFecha(req.params.fecha)) return res.status(400).json({ error: 'Fecha inválida' });

    const { up, ur, justificacion, partida, estatus, busqueda, limite, offset } = req.query;

    if (!up || !ur) {
        return res.status(400).json({ error: 'Faltan parámetros up y ur' });
    }

    const lim = Math.min(parseInt(limite) || 1000, 5000);
    const off = parseInt(offset) || 0;

    const construirQuery = () => {
        let q = supabase
            .from('comprometidos')
            .select('*')
            .eq('fecha', req.params.fecha)
            .eq('up', up)
            .eq('ur', ur);

        if (justificacion) q = q.eq('justificacion', justificacion);
        if (partida) q = q.eq('partida', partida);
        if (estatus) q = q.eq('estatus', estatus);
        if (busqueda) {
            const b = `%${busqueda}%`;
            q = q.or(`num_justificacion.ilike.${b},nombre_fiscal.ilike.${b},concepto.ilike.${b},cve_presupuestal.ilike.${b}`);
        }
        return q;
    };

    const { data: todas, error: errT } = await construirQuery().range(0, 100000);
    if (errT) return res.status(500).json({ error: errT.message });

    const totales = {
        total_registros: todas.length,
        total_original: todas.reduce((s, r) => s + (parseFloat(r.importe_original) || 0), 0),
        total_pagado: todas.reduce((s, r) => s + (parseFloat(r.importe_pagado) || 0), 0),
        total_saldo: todas.reduce((s, r) => s + (parseFloat(r.saldo_a_pagar) || 0), 0)
    };

    const { data: filas, error } = await construirQuery()
        .order('importe_original', { ascending: false })
        .range(off, off + lim - 1);

    if (error) return res.status(500).json({ error: error.message });

    res.json({
        filas,
        totales,
        pagina: Math.floor(off / lim) + 1,
        totalPaginas: Math.ceil(todas.length / lim),
        offset: off,
        limite: lim
    });
});

// ═══════════════════════════════════════════════════════════
// 7. Comparativa Nivel 1 — por Justificación
// ═══════════════════════════════════════════════════════════
app.get('/api/comparar/nivel1/:f1/:f2', async (req, res) => {
    if (!validaFecha(req.params.f1) || !validaFecha(req.params.f2)) return res.status(400).json({ error: 'Fechas inválidas' });

    const { data: dataA, error: errA } = await supabase
        .from('comprometidos')
        .select('justificacion, importe_original, importe_pagado, saldo_a_pagar')
        .eq('fecha', req.params.f1)
        .range(0, 100000);
    const { data: dataB, error: errB } = await supabase
        .from('comprometidos')
        .select('justificacion, importe_original, importe_pagado, saldo_a_pagar')
        .eq('fecha', req.params.f2)
        .range(0, 100000);
    if (errA || errB) return res.status(500).json({ error: 'Error al obtener datos' });

    const mapa = new Map();
    const procesar = (data, origen) => {
        data.forEach(row => {
            const just = row.justificacion || 'SIN_JUSTIFICACION';
            if (!mapa.has(just)) mapa.set(just, { justificacion: just, orig_a: 0, orig_b: 0, pag_a: 0, pag_b: 0, saldo_a: 0, saldo_b: 0, regs_a: 0, regs_b: 0 });
            const item = mapa.get(just);
            const orig = parseFloat(row.importe_original) || 0;
            const pag = parseFloat(row.importe_pagado) || 0;
            const saldo = parseFloat(row.saldo_a_pagar) || 0;
            if (origen === 'A') { item.orig_a += orig; item.pag_a += pag; item.saldo_a += saldo; item.regs_a += 1; }
            else { item.orig_b += orig; item.pag_b += pag; item.saldo_b += saldo; item.regs_b += 1; }
        });
    };
    procesar(dataA, 'A');
    procesar(dataB, 'B');

    const filas = Array.from(mapa.values()).map(f => ({
        ...f,
        dif_orig: f.orig_b - f.orig_a,
        dif_pag: f.pag_b - f.pag_a,
        dif_saldo: f.saldo_b - f.saldo_a
    })).sort((a, b) => Math.abs(b.dif_orig) - Math.abs(a.dif_orig));

    const totales = {
        total_orig_a: filas.reduce((s, f) => s + f.orig_a, 0),
        total_orig_b: filas.reduce((s, f) => s + f.orig_b, 0),
        total_pag_a: filas.reduce((s, f) => s + f.pag_a, 0),
        total_pag_b: filas.reduce((s, f) => s + f.pag_b, 0),
        total_saldo_a: filas.reduce((s, f) => s + f.saldo_a, 0),
        total_saldo_b: filas.reduce((s, f) => s + f.saldo_b, 0),
        total_grupos: filas.length,
        con_cambios: filas.filter(f => f.dif_orig !== 0 || f.dif_pag !== 0 || f.dif_saldo !== 0).length,
        solo_a: filas.filter(f => f.regs_b === 0).length,
        solo_b: filas.filter(f => f.regs_a === 0).length
    };
    res.json({ filas, totales });
});

// 8. Comparativa Nivel 2 — por UP/UR dentro de una Justificación
app.get('/api/comparar/nivel2/:f1/:f2', async (req, res) => {
    if (!validaFecha(req.params.f1) || !validaFecha(req.params.f2)) return res.status(400).json({ error: 'Fechas inválidas' });
    const { justificacion } = req.query;
    if (!justificacion) return res.status(400).json({ error: 'Falta parámetro justificacion' });

    const { data: dataA } = await supabase
        .from('comprometidos')
        .select('up, ur, concepto_ur, dependencia, importe_original, importe_pagado, saldo_a_pagar')
        .eq('fecha', req.params.f1).eq('justificacion', justificacion)
        .range(0, 100000);
    const { data: dataB } = await supabase
        .from('comprometidos')
        .select('up, ur, concepto_ur, dependencia, importe_original, importe_pagado, saldo_a_pagar')
        .eq('fecha', req.params.f2).eq('justificacion', justificacion)
        .range(0, 100000);

    const mapa = new Map();
    const procesar = (data, origen) => {
        data.forEach(row => {
            const key = `${row.up}-${row.ur}`;
            if (!mapa.has(key)) mapa.set(key, { up: row.up, ur: row.ur, concepto_ur: row.concepto_ur, dependencia: row.dependencia, orig_a: 0, orig_b: 0, pag_a: 0, pag_b: 0, saldo_a: 0, saldo_b: 0, regs_a: 0, regs_b: 0 });
            const item = mapa.get(key);
            const orig = parseFloat(row.importe_original) || 0;
            const pag = parseFloat(row.importe_pagado) || 0;
            const saldo = parseFloat(row.saldo_a_pagar) || 0;
            if (origen === 'A') { item.orig_a += orig; item.pag_a += pag; item.saldo_a += saldo; item.regs_a += 1; }
            else { item.orig_b += orig; item.pag_b += pag; item.saldo_b += saldo; item.regs_b += 1; }
        });
    };
    procesar(dataA, 'A');
    procesar(dataB, 'B');

    const filas = Array.from(mapa.values()).map(f => ({
        ...f,
        dif_orig: f.orig_b - f.orig_a,
        dif_pag: f.pag_b - f.pag_a,
        dif_saldo: f.saldo_b - f.saldo_a
    })).sort((a, b) => Math.abs(b.dif_orig) - Math.abs(a.dif_orig));

    const totales = {
        total_orig_a: filas.reduce((s, f) => s + f.orig_a, 0),
        total_orig_b: filas.reduce((s, f) => s + f.orig_b, 0),
        total_pag_a: filas.reduce((s, f) => s + f.pag_a, 0),
        total_pag_b: filas.reduce((s, f) => s + f.pag_b, 0),
        total_saldo_a: filas.reduce((s, f) => s + f.saldo_a, 0),
        total_saldo_b: filas.reduce((s, f) => s + f.saldo_b, 0),
        total_grupos: filas.length,
        con_cambios: filas.filter(f => f.dif_orig !== 0 || f.dif_pag !== 0 || f.dif_saldo !== 0).length,
        solo_a: filas.filter(f => f.regs_b === 0).length,
        solo_b: filas.filter(f => f.regs_a === 0).length
    };
    res.json({ filas, totales });
});

// 9. Comparativa Nivel 3 — Detalle de registros
app.get('/api/comparar/nivel3/:f1/:f2', async (req, res) => {
    if (!validaFecha(req.params.f1) || !validaFecha(req.params.f2)) return res.status(400).json({ error: 'Fechas inválidas' });
    const { up, ur, justificacion } = req.query;
    if (!up || !ur || !justificacion) return res.status(400).json({ error: 'Faltan parámetros' });

    const { data: filasA } = await supabase
        .from('comprometidos').select('*')
        .eq('fecha', req.params.f1).eq('up', up).eq('ur', ur).eq('justificacion', justificacion)
        .order('importe_original', { ascending: false })
        .range(0, 100000);
    const { data: filasB } = await supabase
        .from('comprometidos').select('*')
        .eq('fecha', req.params.f2).eq('up', up).eq('ur', ur).eq('justificacion', justificacion)
        .order('importe_original', { ascending: false })
        .range(0, 100000);

    res.json({ filasA: filasA || [], filasB: filasB || [] });
});

// ═══════════════════════════════════════════════════════════
// 10. ENDPOINT SUBIR BASE (con loteSize = 250 para evitar timeout)
// ═══════════════════════════════════════════════════════════
app.post('/api/subir-base', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'No se proporcionó archivo' });

        const { mapDependencia, mapConceptoUR, mapAsignacion, mapDescripPP, mapDescripComp, mapConceptoPartida, mapDescripDest } = await cargarCatalogo();

        const workbook = xlsx.readFile(req.file.path);
        const datosCrudos = xlsx.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { defval: '' });

        if (datosCrudos.length === 0) {
            fs.unlinkSync(req.file.path);
            return res.status(400).json({ error: 'El archivo está vacío' });
        }

        let fechaBase = req.body.fecha;
        if (!fechaBase) {
            const match = req.file.originalname.match(/(\d{2})-(\d{2})-(\d{2,4})/);
            if (match) {
                const [, dd, mm, aa] = match;
                const yyyy = aa.length === 2 ? `20${aa}` : aa;
                fechaBase = `${yyyy}-${mm}-${dd}`;
            }
        }
        if (!fechaBase || !validaFecha(fechaBase)) {
            fs.unlinkSync(req.file.path);
            return res.status(400).json({ error: 'No se pudo determinar la fecha. Envíala en el campo "fecha" o renombra el archivo como DD-MM-AAAA.' });
        }

        await supabase.from('comprometidos').delete().eq('fecha', fechaBase);

        const datosParaInsertar = [];
        let sinClave = 0;

        datosCrudos.forEach(fila => {
            const codes = extraerDeClave(fila.cvepresupuestal);
            if (!codes) { sinClave++; return; }
            const { up: upRaw, ur: urRaw, ar, pp, comp, part: partRaw, dest, ff, sff, capitulo } = codes;

            const up = norm(upRaw, 2);
            const ur = norm(urRaw, 3);
            const part = norm(partRaw, 4);
            const fechaEmisionISO = excelDateToISO(fila.fechaemision);

            datosParaInsertar.push({
                fecha: fechaBase,
                clave_unica: `${fila.numjustificacion || ''}-${fila.cvepresupuestal || ''}`,
                justificacion: String(fila.justificacion || ''),
                num_justificacion: String(fila.numjustificacion || ''),
                nombre_fiscal: String(fila.nombrefiscal || ''),
                concepto: String(fila.concepto || ''),
                cve_presupuestal: String(fila.cvepresupuestal || ''),
                importe_original: Number(fila.importeoriginal) || 0,
                importe_pagado: Number(fila.importepagado) || 0,
                saldo_a_pagar: Number(fila.saldoapagar) || 0,
                estatus: String(fila.estatus || ''),
                idu: String(fila.idur || ''),
                dependencia_cod: String(fila.dependencia || ''),
                fecha_emision: fechaEmisionISO,
                up, ur, ar, pp,
                componente: comp,
                partida: part,
                destino: dest,
                ff, sff, capitulo,
                dependencia: mapDependencia.get(up) || '',
                concepto_ur: mapConceptoUR.get(`${up}-${ur}`) || '',
                asignacion_recurso: mapAsignacion.get(ar) || '',
                descrip_pp: mapDescripPP.get(pp) || '',
                descrip_componente: mapDescripComp.get(`${up}-${ur}-${pp}-${comp}`) || '',
                concepto_partida: mapConceptoPartida.get(part) || '',
                descrip_destino: mapDescripDest.get(`${up}-${ur}-${pp}-${comp}-${part}-${dest}`) || ''
            });
        });

        console.log(`📥 ${datosCrudos.length} filas leídas, ${datosParaInsertar.length} transformadas (${sinClave} sin clave válida)`);

        if (datosParaInsertar.length === 0) {
            fs.unlinkSync(req.file.path);
            return res.status(400).json({ error: 'Ninguna fila tenía clave presupuestal válida (mínimo 41 caracteres)' });
        }

        // ✅ Lote de 250 para evitar timeout de Supabase (30s)
        const loteSize = 250;
        for (let i = 0; i < datosParaInsertar.length; i += loteSize) {
            const lote = datosParaInsertar.slice(i, i + loteSize);
            const { error } = await supabase.from('comprometidos').insert(lote);
            if (error) throw new Error(`Error al insertar lote: ${error.message}`);
        }

        fs.unlinkSync(req.file.path);
        res.json({ message: `Base ${fechaBase} subida con ${datosParaInsertar.length} registros (${sinClave} omitidas sin clave válida).` });
    } catch (error) {
        console.error('❌ Error al subir base:', error);
        if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.status(500).json({ error: error.message });
    }
});

// ═══════════════════════════════════════════════════════════
// ARRANQUE
// ═══════════════════════════════════════════════════════════
app.listen(PORT, () => {
    console.log(`\n🚀 Servidor Comprometido (Supabase) corriendo en puerto ${PORT}\n`);
});