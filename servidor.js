const express = require('express');
const Database = require('better-sqlite3');
const xlsx = require('xlsx');
const path = require('path');

const app = express();
const DB_PATH = path.join(__dirname, 'comparacion.db');
const PORT = process.env.PORT || 3000;

app.use(express.static(path.join(__dirname, 'public')));

function validaFecha(f) { return /^\d{4}-\d{2}-\d{2}$/.test(f); }
function tablaDe(f) { return `base_${f.replace(/-/g, '_')}`; }

app.get('/', (req, res) => { res.sendFile(path.join(__dirname, 'public', 'index.html')); });

app.get('/api/bases', (req, res) => {
    const db = new Database(DB_PATH);
    const tablas = db.prepare("SELECT name FROM sqlite_master WHERE type='table' AND name LIKE 'base_%' ORDER BY name DESC").all();
    db.close();
    res.json(tablas.map(t => t.name.replace('base_', '').replace(/_/g, '-')));
});

app.get('/api/upur/:fecha', (req, res) => {
    if (!validaFecha(req.params.fecha)) return res.status(400).json({ error: 'Fecha inválida' });
    const tabla = tablaDe(req.params.fecha);
    const db = new Database(DB_PATH);
    const upurList = db.prepare(`SELECT DISTINCT up, ur, concepto_ur, dependencia FROM ${tabla} WHERE up != '' AND ur != '' ORDER BY up, ur`).all();
    db.close();
    res.json(upurList);
});

app.get('/api/justificaciones/:fecha', (req, res) => {
    if (!validaFecha(req.params.fecha)) return res.status(400).json({ error: 'Fecha inválida' });
    const tabla = tablaDe(req.params.fecha);
    const db = new Database(DB_PATH);
    const justificaciones = db.prepare(`SELECT DISTINCT justificacion FROM ${tabla} WHERE justificacion != '' ORDER BY justificacion`).all();
    db.close();
    res.json(justificaciones);
});

app.get('/api/partidas/:fecha', (req, res) => {
    if (!validaFecha(req.params.fecha)) return res.status(400).json({ error: 'Fecha inválida' });
    const tabla = tablaDe(req.params.fecha);
    const db = new Database(DB_PATH);
    const partidas = db.prepare(`SELECT DISTINCT partida, concepto_partida FROM ${tabla} WHERE partida != '' ORDER BY partida`).all();
    db.close();
    res.json(partidas);
});

app.get('/api/base/:fecha', (req, res) => {
    if (!validaFecha(req.params.fecha)) return res.status(400).json({ error: 'Fecha inválida' });
    const tabla = tablaDe(req.params.fecha);
    const db = new Database(DB_PATH);
    const existe = db.prepare(`SELECT name FROM sqlite_master WHERE type='table' AND name=?`).get(tabla);
    if (!existe) { db.close(); return res.status(404).json({ error: 'Base no encontrada' }); }

    const { upur, justificacion, partida, estatus, busqueda, limite } = req.query;
    const lim = Math.min(parseInt(limite) || 1000, 5000);

    let where = 'WHERE 1=1';
    const params = [];
    if (upur) { const [up, ur] = upur.split('-'); where += ` AND up = ? AND ur = ?`; params.push(up, ur); }
    if (justificacion) { where += ` AND justificacion = ?`; params.push(justificacion); }
    if (partida) { where += ` AND partida = ?`; params.push(partida); }
    if (estatus) { where += ` AND estatus = ?`; params.push(estatus); }
    if (busqueda) {
        where += ` AND (num_justificacion LIKE ? OR nombre_fiscal LIKE ? OR concepto LIKE ? OR cve_presupuestal LIKE ?)`;
        const b = `%${busqueda}%`;
        params.push(b, b, b, b);
    }

    const totales = db.prepare(`SELECT COUNT(*) as total_registros, SUM(importe_original) as total_original, SUM(importe_pagado) as total_pagado, SUM(saldo_a_pagar) as total_saldo FROM ${tabla} ${where}`).get(...params);
    const filas = db.prepare(`SELECT * FROM ${tabla} ${where} ORDER BY importe_original DESC LIMIT ?`).all(...params, lim);

    db.close();
    res.json({ filas, totales });
});

app.get('/api/comparar/nivel1/:f1/:f2', (req, res) => {
    if (!validaFecha(req.params.f1) || !validaFecha(req.params.f2)) return res.status(400).json({ error: 'Fechas inválidas' });
    const t1 = tablaDe(req.params.f1);
    const t2 = tablaDe(req.params.f2);
    const db = new Database(DB_PATH);
    const t = db.prepare(`SELECT name FROM sqlite_master WHERE type='table' AND name IN (?,?)`).all(t1, t2);
    if (t.length < 2) { db.close(); return res.status(404).json({ error: 'Una de las bases no existe' }); }
    try { db.exec(`CREATE INDEX IF NOT EXISTS idx_${t1}_just ON ${t1}(justificacion)`); db.exec(`CREATE INDEX IF NOT EXISTS idx_${t2}_just ON ${t2}(justificacion)`); } catch(e) {}
    const filas = db.prepare(`SELECT justificacion, SUM(CASE WHEN origen='A' THEN importe_original ELSE 0 END) as orig_a, SUM(CASE WHEN origen='B' THEN importe_original ELSE 0 END) as orig_b, SUM(CASE WHEN origen='B' THEN importe_original ELSE 0 END) - SUM(CASE WHEN origen='A' THEN importe_original ELSE 0 END) as dif_orig, SUM(CASE WHEN origen='A' THEN importe_pagado ELSE 0 END) as pag_a, SUM(CASE WHEN origen='B' THEN importe_pagado ELSE 0 END) as pag_b, SUM(CASE WHEN origen='B' THEN importe_pagado ELSE 0 END) - SUM(CASE WHEN origen='A' THEN importe_pagado ELSE 0 END) as dif_pag, SUM(CASE WHEN origen='A' THEN saldo_a_pagar ELSE 0 END) as saldo_a, SUM(CASE WHEN origen='B' THEN saldo_a_pagar ELSE 0 END) as saldo_b, SUM(CASE WHEN origen='B' THEN saldo_a_pagar ELSE 0 END) - SUM(CASE WHEN origen='A' THEN saldo_a_pagar ELSE 0 END) as dif_saldo, SUM(CASE WHEN origen='A' THEN 1 ELSE 0 END) as regs_a, SUM(CASE WHEN origen='B' THEN 1 ELSE 0 END) as regs_b FROM (SELECT justificacion, importe_original, importe_pagado, saldo_a_pagar, 'A' as origen FROM ${t1} UNION ALL SELECT justificacion, importe_original, importe_pagado, saldo_a_pagar, 'B' as origen FROM ${t2}) GROUP BY justificacion ORDER BY ABS(dif_orig) DESC`).all();
    const totales = { total_orig_a: filas.reduce((s, f) => s + (f.orig_a || 0), 0), total_orig_b: filas.reduce((s, f) => s + (f.orig_b || 0), 0), total_pag_a: filas.reduce((s, f) => s + (f.pag_a || 0), 0), total_pag_b: filas.reduce((s, f) => s + (f.pag_b || 0), 0), total_saldo_a: filas.reduce((s, f) => s + (f.saldo_a || 0), 0), total_saldo_b: filas.reduce((s, f) => s + (f.saldo_b || 0), 0), total_grupos: filas.length, con_cambios: filas.filter(f => f.dif_orig !== 0 || f.dif_pag !== 0 || f.dif_saldo !== 0).length, solo_a: filas.filter(f => f.regs_b === 0).length, solo_b: filas.filter(f => f.regs_a === 0).length };
    db.close();
    res.json({ filas, totales });
});

app.get('/api/comparar/nivel2/:f1/:f2', (req, res) => {
    if (!validaFecha(req.params.f1) || !validaFecha(req.params.f2)) return res.status(400).json({ error: 'Fechas inválidas' });
    const t1 = tablaDe(req.params.f1);
    const t2 = tablaDe(req.params.f2);
    const db = new Database(DB_PATH);
    const { justificacion } = req.query;
    if (!justificacion) { db.close(); return res.status(400).json({ error: 'Falta parámetro justificacion' }); }
    const t = db.prepare(`SELECT name FROM sqlite_master WHERE type='table' AND name IN (?,?)`).all(t1, t2);
    if (t.length < 2) { db.close(); return res.status(404).json({ error: 'Una de las bases no existe' }); }
    try { db.exec(`CREATE INDEX IF NOT EXISTS idx_${t1}_justupur ON ${t1}(justificacion, up, ur)`); db.exec(`CREATE INDEX IF NOT EXISTS idx_${t2}_justupur ON ${t2}(justificacion, up, ur)`); } catch(e) {}
    const filas = db.prepare(`SELECT up, ur, concepto_ur, dependencia, SUM(CASE WHEN origen='A' THEN importe_original ELSE 0 END) as orig_a, SUM(CASE WHEN origen='B' THEN importe_original ELSE 0 END) as orig_b, SUM(CASE WHEN origen='B' THEN importe_original ELSE 0 END) - SUM(CASE WHEN origen='A' THEN importe_original ELSE 0 END) as dif_orig, SUM(CASE WHEN origen='A' THEN importe_pagado ELSE 0 END) as pag_a, SUM(CASE WHEN origen='B' THEN importe_pagado ELSE 0 END) as pag_b, SUM(CASE WHEN origen='B' THEN importe_pagado ELSE 0 END) - SUM(CASE WHEN origen='A' THEN importe_pagado ELSE 0 END) as dif_pag, SUM(CASE WHEN origen='A' THEN saldo_a_pagar ELSE 0 END) as saldo_a, SUM(CASE WHEN origen='B' THEN saldo_a_pagar ELSE 0 END) as saldo_b, SUM(CASE WHEN origen='B' THEN saldo_a_pagar ELSE 0 END) - SUM(CASE WHEN origen='A' THEN saldo_a_pagar ELSE 0 END) as dif_saldo, SUM(CASE WHEN origen='A' THEN 1 ELSE 0 END) as regs_a, SUM(CASE WHEN origen='B' THEN 1 ELSE 0 END) as regs_b FROM (SELECT up, ur, concepto_ur, dependencia, importe_original, importe_pagado, saldo_a_pagar, 'A' as origen FROM ${t1} WHERE justificacion = ? UNION ALL SELECT up, ur, concepto_ur, dependencia, importe_original, importe_pagado, saldo_a_pagar, 'B' as origen FROM ${t2} WHERE justificacion = ?) GROUP BY up, ur, concepto_ur, dependencia ORDER BY ABS(dif_orig) DESC`).all(justificacion, justificacion);
    const totales = { total_orig_a: filas.reduce((s, f) => s + (f.orig_a || 0), 0), total_orig_b: filas.reduce((s, f) => s + (f.orig_b || 0), 0), total_pag_a: filas.reduce((s, f) => s + (f.pag_a || 0), 0), total_pag_b: filas.reduce((s, f) => s + (f.pag_b || 0), 0), total_saldo_a: filas.reduce((s, f) => s + (f.saldo_a || 0), 0), total_saldo_b: filas.reduce((s, f) => s + (f.saldo_b || 0), 0), total_grupos: filas.length, con_cambios: filas.filter(f => f.dif_orig !== 0 || f.dif_pag !== 0 || f.dif_saldo !== 0).length, solo_a: filas.filter(f => f.regs_b === 0).length, solo_b: filas.filter(f => f.regs_a === 0).length };
    db.close();
    res.json({ filas, totales });
});

app.get('/api/comparar/nivel3/:f1/:f2', (req, res) => {
    if (!validaFecha(req.params.f1) || !validaFecha(req.params.f2)) return res.status(400).json({ error: 'Fechas inválidas' });
    const t1 = tablaDe(req.params.f1);
    const t2 = tablaDe(req.params.f2);
    const db = new Database(DB_PATH);
    const { up, ur, justificacion } = req.query;
    if (!up || !ur || !justificacion) { db.close(); return res.status(400).json({ error: 'Faltan parámetros' }); }
    const filasA = db.prepare(`SELECT * FROM ${t1} WHERE up = ? AND ur = ? AND justificacion = ? ORDER BY importe_original DESC`).all(up, ur, justificacion);
    const filasB = db.prepare(`SELECT * FROM ${t2} WHERE up = ? AND ur = ? AND justificacion = ? ORDER BY importe_original DESC`).all(up, ur, justificacion);
    db.close();
    res.json({ filasA, filasB });
});

app.get('/api/comparar/export/:f1/:f2', (req, res) => {
    if (!validaFecha(req.params.f1) || !validaFecha(req.params.f2)) return res.status(400).send('Fechas inválidas');
    const t1 = tablaDe(req.params.f1);
    const t2 = tablaDe(req.params.f2);
    const db = new Database(DB_PATH);
    const { nivel, justificacion } = req.query;
    let filas = [];
    let nombreHoja = 'Comparativa';
    if (nivel === '1' || !nivel) {
        filas = db.prepare(`SELECT justificacion, SUM(CASE WHEN origen='A' THEN importe_original ELSE 0 END) as orig_a, SUM(CASE WHEN origen='B' THEN importe_original ELSE 0 END) as orig_b, SUM(CASE WHEN origen='B' THEN importe_original ELSE 0 END) - SUM(CASE WHEN origen='A' THEN importe_original ELSE 0 END) as dif_orig, SUM(CASE WHEN origen='A' THEN importe_pagado ELSE 0 END) as pag_a, SUM(CASE WHEN origen='B' THEN importe_pagado ELSE 0 END) as pag_b, SUM(CASE WHEN origen='B' THEN importe_pagado ELSE 0 END) - SUM(CASE WHEN origen='A' THEN importe_pagado ELSE 0 END) as dif_pag, SUM(CASE WHEN origen='A' THEN saldo_a_pagar ELSE 0 END) as saldo_a, SUM(CASE WHEN origen='B' THEN saldo_a_pagar ELSE 0 END) as saldo_b, SUM(CASE WHEN origen='B' THEN saldo_a_pagar ELSE 0 END) - SUM(CASE WHEN origen='A' THEN saldo_a_pagar ELSE 0 END) as dif_saldo, SUM(CASE WHEN origen='A' THEN 1 ELSE 0 END) as regs_a, SUM(CASE WHEN origen='B' THEN 1 ELSE 0 END) as regs_b FROM (SELECT justificacion, importe_original, importe_pagado, saldo_a_pagar, 'A' as origen FROM ${t1} UNION ALL SELECT justificacion, importe_original, importe_pagado, saldo_a_pagar, 'B' as origen FROM ${t2}) GROUP BY justificacion ORDER BY ABS(dif_orig) DESC`).all();
        nombreHoja = 'Nivel1-Justificaciones';
    } else if (nivel === '2' && justificacion) {
        filas = db.prepare(`SELECT up, ur, concepto_ur, dependencia, SUM(CASE WHEN origen='A' THEN importe_original ELSE 0 END) as orig_a, SUM(CASE WHEN origen='B' THEN importe_original ELSE 0 END) as orig_b, SUM(CASE WHEN origen='B' THEN importe_original ELSE 0 END) - SUM(CASE WHEN origen='A' THEN importe_original ELSE 0 END) as dif_orig, SUM(CASE WHEN origen='A' THEN importe_pagado ELSE 0 END) as pag_a, SUM(CASE WHEN origen='B' THEN importe_pagado ELSE 0 END) as pag_b, SUM(CASE WHEN origen='B' THEN importe_pagado ELSE 0 END) - SUM(CASE WHEN origen='A' THEN importe_pagado ELSE 0 END) as dif_pag, SUM(CASE WHEN origen='A' THEN saldo_a_pagar ELSE 0 END) as saldo_a, SUM(CASE WHEN origen='B' THEN saldo_a_pagar ELSE 0 END) as saldo_b, SUM(CASE WHEN origen='B' THEN saldo_a_pagar ELSE 0 END) - SUM(CASE WHEN origen='A' THEN saldo_a_pagar ELSE 0 END) as dif_saldo, SUM(CASE WHEN origen='A' THEN 1 ELSE 0 END) as regs_a, SUM(CASE WHEN origen='B' THEN 1 ELSE 0 END) as regs_b FROM (SELECT up, ur, concepto_ur, dependencia, importe_original, importe_pagado, saldo_a_pagar, 'A' as origen FROM ${t1} WHERE justificacion = ? UNION ALL SELECT up, ur, concepto_ur, dependencia, importe_original, importe_pagado, saldo_a_pagar, 'B' as origen FROM ${t2} WHERE justificacion = ?) GROUP BY up, ur, concepto_ur, dependencia ORDER BY ABS(dif_orig) DESC`).all(justificacion, justificacion);
        nombreHoja = `Nivel2-${justificacion.substring(0, 20)}`;
    }
    db.close();
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, xlsx.utils.json_to_sheet(filas), nombreHoja);
    const buffer = xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Disposition', `attachment; filename=Comparativa_${req.params.f1}_vs_${req.params.f2}_${nombreHoja}.xlsx`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buffer);
});

app.listen(PORT, () => { console.log(`\n🚀 Servidor: http://localhost:${PORT}\n`); });
