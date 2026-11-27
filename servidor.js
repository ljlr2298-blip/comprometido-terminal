   // === MIGRACION SUPABASE 2026-11-26 ===
require('dotenv').config();
const express = require('express');
const { createClient } = require('@supabase/supabase-js');
const xlsx = require('xlsx');
const multer = require('multer');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

// Configurar Supabase
const supabaseUrl = process.env.SUPABASE_URL;
const supabaseKey = process.env.SUPABASE_KEY;
const supabase = createClient(supabaseUrl, supabaseKey);

// Configurar subida de archivos
const upload = multer({ dest: 'uploads/' });

app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

// Función auxiliar para validar fecha
function validaFecha(f) { return /^\d{4}-\d{2}-\d{2}$/.test(f); }

// 1. Obtener lista de fechas disponibles
app.get('/api/bases', async (req, res) => {
    const { data, error } = await supabase.from('comprometidos').select('fecha').order('fecha', { ascending: false });
    if (error) return res.status(500).json({ error: error.message });
    
    // Obtener fechas únicas
    const fechasUnicas = [...new Set(data.map(item => item.fecha))];
    res.json(fechasUnicas);
});

// 2. Obtener UP/UR únicos de una fecha
app.get('/api/upur/:fecha', async (req, res) => {
    if (!validaFecha(req.params.fecha)) return res.status(400).json({ error: 'Fecha inválida' });
    const { data, error } = await supabase.from('comprometidos').select('up, ur, concepto_ur, dependencia').eq('fecha', req.params.fecha).not('up', 'is', null).not('ur', 'is', null);
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

// 3. Obtener justificaciones únicas
app.get('/api/justificaciones/:fecha', async (req, res) => {
    if (!validaFecha(req.params.fecha)) return res.status(400).json({ error: 'Fecha inválida' });
    const { data, error } = await supabase.from('comprometidos').select('justificacion').eq('fecha', req.params.fecha).not('justificacion', 'is', null);
    if (error) return res.status(500).json({ error: error.message });
    
    const unique = [...new Set(data.map(r => r.justificacion))].filter(Boolean).sort();
    res.json(unique.map(j => ({ justificacion: j })));
});

// 4. Obtener partidas únicas
app.get('/api/partidas/:fecha', async (req, res) => {
    if (!validaFecha(req.params.fecha)) return res.status(400).json({ error: 'Fecha inválida' });
    const { data, error } = await supabase.from('comprometidos').select('partida, concepto_partida').eq('fecha', req.params.fecha).not('partida', 'is', null);
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

// 5. Consulta principal con filtros
app.get('/api/base/:fecha', async (req, res) => {
    if (!validaFecha(req.params.fecha)) return res.status(400).json({ error: 'Fecha inválida' });
    
    let query = supabase.from('comprometidos').select('*').eq('fecha', req.params.fecha);
    
    const { upur, justificacion, partida, estatus, busqueda, limite } = req.query;
    const lim = Math.min(parseInt(limite) || 1000, 5000);

    if (upur) {
        const [up, ur] = upur.split('-');
        query = query.eq('up', up).eq('ur', ur);
    }
    if (justificacion) query = query.eq('justificacion', justificacion);
    if (partida) query = query.eq('partida', partida);
    if (estatus) query = query.eq('estatus', estatus);
    if (busqueda) {
        const b = `%${busqueda}%`;
        query = query.or(`num_justificacion.ilike.${b},nombre_fiscal.ilike.${b},concepto.ilike.${b},cve_presupuestal.ilike.${b}`);
    }

    // Nota: Supabase no permite COUNT() y LIMIT() en la misma consulta fácilmente sin RPC, 
    // así que hacemos dos consultas rápidas o contamos en JS si el límite es bajo.
    // Para optimizar, pedimos los datos y calculamos totales en JS (es rápido para <5000 filas).
    const { data: filas, error } = await query.limit(lim).order('importe_original', { ascending: false });
    
    if (error) return res.status(500).json({ error: error.message });

    const totales = {
        total_registros: filas.length,
        total_original: filas.reduce((sum, r) => sum + (parseFloat(r.importe_original) || 0), 0),
        total_pagado: filas.reduce((sum, r) => sum + (parseFloat(r.importe_pagado) || 0), 0),
        total_saldo: filas.reduce((sum, r) => sum + (parseFloat(r.saldo_a_pagar) || 0), 0)
    };

    res.json({ filas, totales });
});

// 6. Comparativa Nivel 1 (Agrupado por Justificación)
app.get('/api/comparar/nivel1/:f1/:f2', async (req, res) => {
    if (!validaFecha(req.params.f1) || !validaFecha(req.params.f2)) return res.status(400).json({ error: 'Fechas inválidas' });
    
    const { data: dataA, error: errA } = await supabase.from('comprometidos').select('justificacion, importe_original, importe_pagado, saldo_a_pagar').eq('fecha', req.params.f1);
    const { data: dataB, error: errB } = await supabase.from('comprometidos').select('justificacion, importe_original, importe_pagado, saldo_a_pagar').eq('fecha', req.params.f2);
    
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
            
            if (origen === 'A') {
                item.orig_a += orig; item.pag_a += pag; item.saldo_a += saldo; item.regs_a += 1;
            } else {
                item.orig_b += orig; item.pag_b += pag; item.saldo_b += saldo; item.regs_b += 1;
            }
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

// 7. Comparativa Nivel 2 (Agrupado por UP/UR dentro de una Justificación)
app.get('/api/comparar/nivel2/:f1/:f2', async (req, res) => {
    if (!validaFecha(req.params.f1) || !validaFecha(req.params.f2)) return res.status(400).json({ error: 'Fechas inválidas' });
    const { justificacion } = req.query;
    if (!justificacion) return res.status(400).json({ error: 'Falta parámetro justificacion' });

    const { data: dataA } = await supabase.from('comprometidos').select('up, ur, concepto_ur, dependencia, importe_original, importe_pagado, saldo_a_pagar').eq('fecha', req.params.f1).eq('justificacion', justificacion);
    const { data: dataB } = await supabase.from('comprometidos').select('up, ur, concepto_ur, dependencia, importe_original, importe_pagado, saldo_a_pagar').eq('fecha', req.params.f2).eq('justificacion', justificacion);

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
        ...f, dif_orig: f.orig_b - f.orig_a, dif_pag: f.pag_b - f.pag_a, dif_saldo: f.saldo_b - f.saldo_a
    })).sort((a, b) => Math.abs(b.dif_orig) - Math.abs(a.dif_orig));

    const totales = {
        total_orig_a: filas.reduce((s, f) => s + f.orig_a, 0), total_orig_b: filas.reduce((s, f) => s + f.orig_b, 0),
        total_pag_a: filas.reduce((s, f) => s + f.pag_a, 0), total_pag_b: filas.reduce((s, f) => s + f.pag_b, 0),
        total_saldo_a: filas.reduce((s, f) => s + f.saldo_a, 0), total_saldo_b: filas.reduce((s, f) => s + f.saldo_b, 0),
        total_grupos: filas.length, con_cambios: filas.filter(f => f.dif_orig !== 0 || f.dif_pag !== 0 || f.dif_saldo !== 0).length,
        solo_a: filas.filter(f => f.regs_b === 0).length, solo_b: filas.filter(f => f.regs_a === 0).length
    };

    res.json({ filas, totales });
});

// 8. Comparativa Nivel 3 (Detalle de registros)
app.get('/api/comparar/nivel3/:f1/:f2', async (req, res) => {
    if (!validaFecha(req.params.f1) || !validaFecha(req.params.f2)) return res.status(400).json({ error: 'Fechas inválidas' });
    const { up, ur, justificacion } = req.query;
    if (!up || !ur || !justificacion) return res.status(400).json({ error: 'Faltan parámetros' });

    const { data: filasA } = await supabase.from('comprometidos').select('*').eq('fecha', req.params.f1).eq('up', up).eq('ur', ur).eq('justificacion', justificacion).order('importe_original', { ascending: false });
    const { data: filasB } = await supabase.from('comprometidos').select('*').eq('fecha', req.params.f2).eq('up', up).eq('ur', ur).eq('justificacion', justificacion).order('importe_original', { ascending: false });

    res.json({ filasA: filasA || [], filasB: filasB || [] });
});

// 9. 🚀 Endpoint para subir una base nueva desde la web
// --- Normalizar encabezados: minúsculas, sin acentos, sin espacios/puntos ---
function normalizarClave(k) {
    return String(k)
        .toLowerCase()
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
        .replace(/[^a-z0-9]/g, '');
}

function mapearFila(row) {
    const n = {};
    for (const [k, v] of Object.entries(row)) n[normalizarClave(k)] = v;

    const txt = (...cands) => {
        for (const c of cands) {
            const v = n[c];
            if (v !== undefined && v !== null && String(v).trim() !== '') return String(v).trim();
        }
        return '';
    };
    const num = (...cands) => {
        for (const c of cands) {
            const v = n[c];
            if (v !== undefined && v !== null && String(v).trim() !== '') {
                const x = parseFloat(String(v).replace(/[$,\s]/g, ''));
                if (!isNaN(x)) return x;
            }
        }
        return 0;
    };

    return {
        justificacion:      txt('justificacion','tipodejustificacion','justificacions'),
        num_justificacion:  txt('numjustificacion','nojustificacion','numerojustificacion','justificacionnumero','nofolio','folio'),
        nombre_fiscal:      txt('nombrefiscal','fiscal','proveedor','beneficiario'),
        cve_presupuestal:   txt('clavepresupuestal','cvepresupuestal','clavereal','clavepresupuestaria'),
        up:                 txt('up','unidadprogramatica'),
        ur:                 txt('ur','unidadresponsable'),
        partida:            txt('partida','partidadelgasto','partidadegasto'),
        concepto_partida:   txt('conceptopartida','descripcionpartida'),
        concepto:           txt('concepto','conceptodelpago','descripcion'),
        estatus:            txt('estatus','estado','estatuspago'),
        importe_original:   num('importeoriginal','original','montooriginal','solicitado'),
        importe_pagado:     num('importepagado','pagado','montopagado','amortizado'),
        saldo_a_pagar:      num('saldopagar','saldoapagar','saldo','saldorestante','pendiente'),
        fecha_emision:      txt('fechaemision','fechadeldocumento','fechadeemision'),
        dependencia:        txt('dependencia','unidadadministrativa','entidad'),
        concepto_ur:        txt('conceptour','nombreur')
    };
}

app.post('/api/subir-base', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'No se proporcionó archivo' });

        const workbook = xlsx.readFile(req.file.path);
        const sheetName = workbook.SheetNames[0];
        const jsonData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

        if (jsonData.length === 0) {
            fs.unlinkSync(req.file.path);
            return res.status(400).json({ error: 'El archivo está vacío' });
        }

        // Fecha: viene del formulario, o se infiere de la primera fila
        let fechaBase = req.body.fecha;
        if (!fechaBase && jsonData[0].fecha) fechaBase = String(jsonData[0].fecha).substring(0, 10);
        if (!fechaBase || !validaFecha(fechaBase)) {
            fs.unlinkSync(req.file.path);
            return res.status(400).json({ error: 'No se pudo determinar la fecha. Envíala en el campo "fecha".' });
        }

        // 🧹 Idempotencia: borrar cualquier intento previo de esta fecha
        await supabase.from('comprometidos').delete().eq('fecha', fechaBase);

        // Mapear filas con encabezados flexibles
        const datosParaInsertar = jsonData.map(row => ({
            fecha: fechaBase,
            ...mapearFila(row)
        }));

        // Insertar en lotes de 1000
        const loteSize = 1000;
        for (let i = 0; i < datosParaInsertar.length; i += loteSize) {
            const lote = datosParaInsertar.slice(i, i + loteSize);
            const { error } = await supabase.from('comprometidos').insert(lote);
            if (error) throw new Error(`Error al insertar lote: ${error.message}`);
        }

        fs.unlinkSync(req.file.path);
        res.json({ message: `Base ${fechaBase} subida con ${datosParaInsertar.length} registros.` });
    } catch (error) {
        console.error('Error al subir base:', error);
        if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.status(500).json({ error: error.message });
    }
});

app.listen(PORT, () => {
    console.log(`\n🚀 Servidor Comprometido (Supabase) corriendo en puerto ${PORT}\n`);
});