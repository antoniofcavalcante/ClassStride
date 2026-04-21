// ═══════════════════════════════════════════════════════════
//  CONSELHO DE CLASSE — app.js
//  Abas: "Notas", "FALTAS", "DIGITAL"
//  Cruzamento: ESTUDANTE + TURMA + DISCIPLINA
//  Média de aprovação: 5,0 | Frequência mínima: 75%
//  Frequência crítica: < 50% (risco de retenção)
// ═══════════════════════════════════════════════════════════

const APP = { notas: [], faltas: [], digital: [], charts: {}, turmaSel: null, discSel: null };

const CFG = {
  mediaAprov: 5.0,
  freqMin: 75,
  freqRisco: 50,
  escola: 'E.E. Professora Célia Vasques Ferrari Duch',
  etapa: 'Ensino Médio',
};

// ── Utilitários ────────────────────────────────────────────

function toNum(v) {
  if (v === null || v === undefined || v === '' || v === '-' || v === '—') return null;
  if (typeof v === 'number' && isFinite(v)) return v;
  const n = parseFloat(String(v).replace(',', '.'));
  return isFinite(n) ? n : null;
}

function unique(arr) { return [...new Set(arr.filter(Boolean))].sort(); }

function iniciais(nome) {
  const p = String(nome).trim().split(/\s+/);
  return ((p[0]?.[0] || '') + (p[1]?.[0] || '')).toUpperCase();
}

function destroyChart(id) {
  if (APP.charts[id]) { APP.charts[id].destroy(); delete APP.charts[id]; }
}

function distribuirPct(valores) {
  const total = valores.reduce((a, b) => a + b, 0);
  if (!total) return valores.map(() => 0);
  const exatos = valores.map(v => v / total * 100);
  const floors = exatos.map(v => Math.floor(v));
  let resto = 100 - floors.reduce((a, b) => a + b, 0);
  const indices = exatos.map((v, i) => ({ i, rem: v - floors[i] })).sort((a, b) => b.rem - a.rem);
  for (let k = 0; k < resto; k++) floors[indices[k].i]++;
  return floors;
}

function chave(e, t, d) { return `${e}|${t}|${d}`; }

// Normaliza strings de identificação para evitar mismatch silencioso entre abas
function normalizar(v) {
  return String(v || '').trim().replace(/\s+/g, ' ').toUpperCase();
}

// ── Plugins Chart.js ───────────────────────────────────────

const pluginBarPct = {
  id: 'barPct',
  afterDatasetsDraw(chart) {
    if (chart.config.type !== 'bar') return;
    const { ctx, data } = chart;
    const isHoriz = chart.config.options?.indexAxis === 'y';
    const numCols = data.labels?.length || 0;
    const colPcts = [];
    for (let i = 0; i < numCols; i++) colPcts.push(distribuirPct(data.datasets.map(ds => ds.data[i] || 0)));
    data.datasets.forEach((ds, di) => {
      const meta = chart.getDatasetMeta(di);
      if (meta.hidden) return;
      meta.data.forEach((bar, i) => {
        const val = ds.data[i];
        if (!val) return;
        const pct = colPcts[i][di];
        if (pct < 8) return;
        ctx.save();
        ctx.fillStyle = '#fff'; ctx.font = 'bold 11px sans-serif';
        ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
        if (isHoriz) { const cx = bar.x - (bar.x - bar.base) / 2; if (Math.abs(bar.x - bar.base) > 34) ctx.fillText(pct + '%', cx, bar.y); }
        else { if (bar.height > 18) ctx.fillText(pct + '%', bar.x, bar.y + bar.height / 2); }
        ctx.restore();
      });
    });
  }
};

const pluginBarValor = {
  id: 'barValor',
  afterDatasetsDraw(chart) {
    if (chart.config.type !== 'bar' || chart.config.options?.indexAxis !== 'y' || chart.data.datasets.length !== 1) return;
    const { ctx, data } = chart;
    data.datasets[0].data.forEach((val, i) => {
      if (!val) return;
      const bar = chart.getDatasetMeta(0).data[i];
      ctx.save(); ctx.fillStyle = '#374151'; ctx.font = 'bold 11px sans-serif';
      ctx.textAlign = 'left'; ctx.textBaseline = 'middle';
      ctx.fillText(val.toFixed(1).replace('.', ','), bar.x + 6, bar.y); ctx.restore();
    });
  }
};

const pluginLinePct = {
  id: 'linePct',
  afterDatasetsDraw(chart) {
    if (chart.config.type !== 'line') return;
    const { ctx, data } = chart;
    data.datasets.forEach((ds, di) => {
      const meta = chart.getDatasetMeta(di);
      if (meta.hidden) return;
      meta.data.forEach((pt, idx) => {
        if (idx === 0) return;
        let prevIdx = idx - 1;
        while (prevIdx >= 0 && (ds.data[prevIdx] === null || ds.data[prevIdx] === undefined)) prevIdx--;
        if (prevIdx < 0) return;
        const prev = ds.data[prevIdx], curr = ds.data[idx];
        if (prev === null || curr === null || prev === undefined || curr === undefined) return;
        const delta = curr - prev;
        if (Math.abs(delta) < 0.05) return;
        ctx.save(); ctx.font = 'bold 10px sans-serif';
        ctx.fillStyle = delta > 0 ? '#057a55' : '#c81e1e';
        ctx.textAlign = 'center'; ctx.textBaseline = 'bottom';
        ctx.fillText((delta > 0 ? '+' : '') + delta.toFixed(1).replace('.', ','), pt.x, pt.y - 6); ctx.restore();
      });
    });
  }
};

function makePctLabelPlugin(getTotal) {
  return {
    id: 'pctLabel_' + Math.random(),
    afterDraw(chart) {
      const total = typeof getTotal === 'function' ? getTotal() : getTotal;
      if (!total) return;
      const ctx2 = chart.ctx;
      chart.data.datasets.forEach((ds, di) => {
        const meta = chart.getDatasetMeta(di);
        const pcts = distribuirPct(ds.data.map(v => v || 0));
        meta.data.forEach((arc, i) => {
          if (!ds.data[i] || pcts[i] < 5) return;
          const mid = arc.startAngle + (arc.endAngle - arc.startAngle) / 2;
          const r = (arc.innerRadius + arc.outerRadius) / 2;
          ctx2.save(); ctx2.fillStyle = '#fff'; ctx2.font = 'bold 13px sans-serif';
          ctx2.textAlign = 'center'; ctx2.textBaseline = 'middle';
          ctx2.fillText(pcts[i] + '%', arc.x + r * Math.cos(mid), arc.y + r * Math.sin(mid)); ctx2.restore();
        });
      });
    }
  };
}

// ── Opções de gráfico ──────────────────────────────────────

function optsBarEmpilhado(extra) {
  return Object.assign({
    responsive: true, maintainAspectRatio: false,
    plugins: { legend: { position: 'top', labels: { font: { size: 12 }, boxWidth: 12 } } },
    scales: { x: { stacked: true }, y: { stacked: true, ticks: { stepSize: 1 } } }
  }, extra);
}
function optsBarHoriz(extra) {
  return Object.assign({
    indexAxis: 'y', responsive: true, maintainAspectRatio: false,
    plugins: { legend: { display: false } }, scales: { x: { min: 0, max: 10 } }
  }, extra);
}
function optsLinha(extra) {
  return Object.assign({
    responsive: true, maintainAspectRatio: false,
    plugins: { legend: { position: 'top', labels: { font: { size: 12 }, boxWidth: 12 } } },
    scales: { y: { min: 0, max: 10 } }
  }, extra);
}

// ── Cálculos pedagógicos ───────────────────────────────────

function calcMedia(n) {
  const vals = [toNum(n.b1), toNum(n.b2), toNum(n.b3), toNum(n.b4)].filter(v => v !== null);
  if (!vals.length) return null;
  return +(vals.reduce((a, b) => a + b, 0) / vals.length).toFixed(2);
}

function isParcial(n) { return toNum(n.b4) === null; }

function evolucao(n) {
  const vals = [toNum(n.b1), toNum(n.b2), toNum(n.b3), toNum(n.b4)].filter(v => v !== null);
  if (vals.length < 2) return null;
  return +(vals[vals.length - 1] - vals[0]).toFixed(1);
}

function calcFreq(f) {
  if (!f) return null;
  const pares = [
    { aulas: toNum(f.a1), faltas: toNum(f.f1) }, { aulas: toNum(f.a2), faltas: toNum(f.f2) },
    { aulas: toNum(f.a3), faltas: toNum(f.f3) }, { aulas: toNum(f.a4), faltas: toNum(f.f4) },
  ].filter(p => p.aulas !== null && p.aulas > 0);
  if (!pares.length) return null;
  const totalAulas = pares.reduce((s, p) => s + p.aulas, 0);
  const totalFaltas = pares.reduce((s, p) => s + (p.faltas || 0), 0);
  const bims = [
    { aulas: toNum(f.a1), faltas: toNum(f.f1) }, { aulas: toNum(f.a2), faltas: toNum(f.f2) },
    { aulas: toNum(f.a3), faltas: toNum(f.f3) }, { aulas: toNum(f.a4), faltas: toNum(f.f4) },
  ].map(p => (!p.aulas || p.aulas <= 0) ? null : +((1 - (p.faltas || 0) / p.aulas) * 100).toFixed(1));
  return { aulas: totalAulas, faltas: totalFaltas, pct: +((1 - totalFaltas / totalAulas) * 100).toFixed(1), bims };
}

function calcFreqGeral(nome, turma) {
  const regs = APP.notas.filter(r => r.ESTUDANTE === nome && r.TURMA === turma);
  let totalAulas = 0, totalFaltas = 0;
  regs.forEach(r => {
    const f = getFaltas(r.ESTUDANTE, r.TURMA, r.DISCIPLINA);
    if (!f) return;
    const freq = calcFreq(f);
    if (freq) { totalAulas += freq.aulas; totalFaltas += freq.faltas; }
  });
  if (!totalAulas) return null;
  return { aulas: totalAulas, faltas: totalFaltas, pct: +((1 - totalFaltas / totalAulas) * 100).toFixed(1) };
}

function situacao(media) {
  if (media === null) return '—';
  return media >= CFG.mediaAprov ? 'Aprovado' : 'Reprovado';
}

function alertaFreq(freq) { return !!(freq && freq.pct < CFG.freqMin); }

function nivelRiscoFreq(freq) {
  if (!freq) return 'sem-dados';
  if (freq.pct < CFG.freqRisco) return 'critico';
  if (freq.pct < CFG.freqMin) return 'risco';
  if (freq.pct < 80) return 'atencao';
  return 'ok';
}

function tendenciaPioraFreq(freq) {
  if (!freq) return false;
  const bims = freq.bims.filter(b => b !== null);
  if (bims.length < 2) return false;
  return (bims[0] - bims[bims.length - 1]) > 10;
}

function getFaltas(estudante, turma, disc) {
  return APP.faltas.find(f => f.ESTUDANTE === estudante && f.TURMA === turma && f.DISCIPLINA === disc) || null;
}

function dadosReg(n) {
  const f = getFaltas(n.ESTUDANTE, n.TURMA, n.DISCIPLINA);
  const media = calcMedia(n);
  const freq = calcFreq(f);
  const sit = situacao(media);
  const alFreq = alertaFreq(freq);
  const parc = isParcial(n);
  const nivelFreq = nivelRiscoFreq(freq);
  const piora = tendenciaPioraFreq(freq);
  return { n, f, media, freq, sit, alFreq, parc, nivelFreq, piora };
}

// ── HTML helpers ───────────────────────────────────────────

function fmtNota(v) {
  if (v === null) return '<span class="nota n-ne">—</span>';
  const cls = v >= CFG.mediaAprov ? 'n-ok' : 'n-re';
  return `<span class="nota ${cls}">${v.toFixed(1).replace('.', ',')}</span>`;
}

function fmtMedia(media, parcial) {
  if (media === null) return '<span class="nota n-ne">—</span>';
  const cls = media >= CFG.mediaAprov ? 'n-ok' : 'n-re';
  const tag = parcial ? ' <span style="font-size:10px;color:var(--cinza);">(parc.)</span>' : '';
  return `<span class="nota ${cls}">${media.toFixed(1).replace('.', ',')}</span>${tag}`;
}

function fmtFreqCell(freq) {
  if (!freq) return '<span class="nota n-ne">—</span>';
  const nivel = nivelRiscoFreq(freq);
  const cls = (nivel === 'ok' || nivel === 'atencao') ? (freq.pct >= 85 ? 'freq-ok' : 'freq-al') : 'freq-ruim';
  let extra = '';
  if (nivel === 'critico') extra = '<span class="alerta-freq"><i class="bi bi-x-circle-fill"></i> Crítico</span>';
  else if (nivel === 'risco') extra = '<span class="alerta-freq"><i class="bi bi-exclamation-triangle-fill"></i> &lt;75%</span>';
  return `<span class="${cls}">${freq.pct.toFixed(1)}%</span>${extra}`;
}

function fmtSit(media, freq, parcial) {
  const sit = situacao(media);
  const al = alertaFreq(freq);
  const nivel = nivelRiscoFreq(freq);
  const suf = parcial ? ' <span style="font-size:10px;">(parc.)</span>' : '';
  if (sit === 'Aprovado') {
    if (nivel === 'critico') return `<span class="badge b-re"><i class="bi bi-person-x-fill"></i> Aprov. ⚠ ret.freq.</span>${suf}`;
    if (al) return `<span class="badge b-al"><i class="bi bi-exclamation-triangle-fill"></i> Aprovado ⚠ freq.</span>${suf}`;
    return `<span class="badge b-ap"><i class="bi bi-check-circle-fill"></i> Aprovado</span>${suf}`;
  }
  if (sit === 'Reprovado') return `<span class="badge b-re"><i class="bi bi-x-circle-fill"></i> Reprovado</span>${suf}`;
  return `<span class="badge b-pa">—</span>`;
}

function fmtEvol(n) {
  const e = evolucao(n);
  if (e === null) return '<span class="evol-neu">—</span>';
  if (e > 0) return `<span class="evol-pos">▲ ${e.toFixed(1)}</span>`;
  if (e < 0) return `<span class="evol-neg">▼ ${Math.abs(e).toFixed(1)}</span>`;
  return '<span class="evol-neu">= 0</span>';
}

// ── Material Digital ───────────────────────────────────────

function calcDigital(registros) {
  let totalPrev = 0, totalConc = 0;
  registros.forEach(r => {
    [1, 2, 3, 4].forEach(b => {
      const prev = toNum(r[`prev${b}`]), conc = toNum(r[`conc${b}`]);
      if (prev && prev > 0) { totalPrev += prev; totalConc += (conc || 0); }
    });
  });
  if (!totalPrev) return null;
  return { previsto: totalPrev, concluido: totalConc, pct: +((totalConc / totalPrev) * 100).toFixed(1) };
}

function calcDigitalBimestres(registros) {
  return [1, 2, 3, 4].map(b => {
    let prev = 0, conc = 0;
    registros.forEach(r => {
      const p = toNum(r[`prev${b}`]), c = toNum(r[`conc${b}`]);
      if (p && p > 0) { prev += p; conc += (c || 0); }
    });
    if (!prev) return null;
    return { bim: b, previsto: prev, concluido: conc, pct: +((conc / prev) * 100).toFixed(1) };
  });
}

function fmtDigitalCard(turma, disc, registros) {
  const calc = calcDigital(registros);
  if (!calc) return '';
  const pct = Math.min(100, calc.pct);
  const cls = pct >= 80 ? 'digital-prog-ok' : pct >= 50 ? 'digital-prog-al' : 'digital-prog-re';
  const bims = calcDigitalBimestres(registros);
  const nomeBim = ['1º BIM', '2º BIM', '3º BIM', '4º BIM'];
  const bimRows = bims.map((b, i) => {
    if (!b) return '';
    const bPct = Math.min(100, b.pct);
    const bCls = bPct >= 80 ? 'digital-prog-ok' : bPct >= 50 ? 'digital-prog-al' : 'digital-prog-re';
    return `<div class="digital-bim-row">
      <div class="digital-bim-label"><span class="digital-bim-nome">${nomeBim[i]}</span><span class="digital-bim-info">${b.concluido} de ${b.previsto}</span></div>
      <div class="digital-bim-bar-wrap"><div class="digital-bim-bar ${bCls}" style="width:${bPct}%"></div></div>
      <span class="digital-bim-pct ${bCls}">${b.pct.toFixed(0)}%</span>
    </div>`;
  }).join('');
  return `<div class="digital-card">
    <div class="digital-card-head">
      <span class="digital-card-turma">${turma}</span>
      <span class="digital-card-disc">${disc}</span>
      <span class="digital-pct ${cls}">${calc.pct.toFixed(1)}%</span>
    </div>
    <div class="digital-prog-bar-wrap"><div class="digital-prog-bar ${cls}" style="width:${pct}%"></div></div>
    <div class="digital-card-info">${calc.concluido} de ${calc.previsto} aulas concluídas</div>
    <div class="digital-bim-breakdown">${bimRows}</div>
  </div>`;
}

// ── Parsers ────────────────────────────────────────────────

function colIdx(header, ...keys) {
  return header.findIndex(h => keys.some(k => String(h).toUpperCase().includes(k.toUpperCase())));
}

function parseNotasSheet(rows) {
  if (!rows.length) return { records: [], invalidos: 0 };
  const h = rows[0].map(c => String(c || '').trim());
  const iE = colIdx(h, 'ESTUDANTE', 'ALUNO', 'NOME');
  const iT = colIdx(h, 'TURMA');
  const iD = colIdx(h, 'DISCIPLINA');
  const iB1 = colIdx(h, '1º BIMESTRE', '1 BIMESTRE', 'NOTA 1', 'B1', 'BIM1', '1BIM', 'NOTA1');
  const iB2 = colIdx(h, '2º BIMESTRE', '2 BIMESTRE', 'NOTA 2', 'B2', 'BIM2', '2BIM', 'NOTA2');
  const iB3 = colIdx(h, '3º BIMESTRE', '3 BIMESTRE', 'NOTA 3', 'B3', 'BIM3', '3BIM', 'NOTA3');
  const iB4 = colIdx(h, '4º BIMESTRE', '4 BIMESTRE', 'NOTA 4', 'B4', 'BIM4', '4BIM', 'NOTA4');
  if (iE < 0 || iT < 0 || iD < 0) return { records: [], invalidos: 0 };
  let invalidos = 0;
  const records = rows.slice(1).map(r => {
    const rec = {
      ESTUDANTE: normalizar(r[iE]),
      TURMA: normalizar(r[iT]),
      DISCIPLINA: normalizar(r[iD]),
      b1: iB1 >= 0 ? toNum(r[iB1]) : null,
      b2: iB2 >= 0 ? toNum(r[iB2]) : null,
      b3: iB3 >= 0 ? toNum(r[iB3]) : null,
      b4: iB4 >= 0 ? toNum(r[iB4]) : null,
    };
    // Validar range de notas [0, 10]
    ['b1', 'b2', 'b3', 'b4'].forEach(b => {
      if (rec[b] !== null && (rec[b] < 0 || rec[b] > 10)) { rec[b] = null; invalidos++; }
    });
    return rec;
  }).filter(r => r.ESTUDANTE && r.TURMA && r.DISCIPLINA);
  return { records, invalidos };
}

function parseFaltasSheet(rows) {
  if (!rows.length) return { records: [], invalidos: 0 };
  const h = rows[0].map(c => String(c || '').trim());
  const iE = colIdx(h, 'ESTUDANTE', 'ALUNO', 'NOME');
  const iT = colIdx(h, 'TURMA');
  const iD = colIdx(h, 'DISCIPLINA');
  const iA1 = colIdx(h, 'AULAS 1', 'A1', 'AULAS1');
  const iA2 = colIdx(h, 'AULAS 2', 'A2', 'AULAS2');
  const iA3 = colIdx(h, 'AULAS 3', 'A3', 'AULAS3');
  const iA4 = colIdx(h, 'AULAS 4', 'A4', 'AULAS4');
  const iF1 = colIdx(h, 'FALTAS 1', 'F1', 'FALTAS1');
  const iF2 = colIdx(h, 'FALTAS 2', 'F2', 'FALTAS2');
  const iF3 = colIdx(h, 'FALTAS 3', 'F3', 'FALTAS3');
  const iF4 = colIdx(h, 'FALTAS 4', 'F4', 'FALTAS4');
  const iTotalAulas = colIdx(h, 'TOTAL DE AULAS', 'TOTAL AULAS');
  if (iE < 0 || iT < 0 || iD < 0) return { records: [], invalidos: 0 };
  let invalidos = 0;
  const parsed = rows.slice(1).map(r => ({
    ESTUDANTE: normalizar(r[iE]),
    TURMA: normalizar(r[iT]),
    DISCIPLINA: normalizar(r[iD]),
    a1: iA1 >= 0 ? toNum(r[iA1]) : null, a2: iA2 >= 0 ? toNum(r[iA2]) : null,
    a3: iA3 >= 0 ? toNum(r[iA3]) : null, a4: iA4 >= 0 ? toNum(r[iA4]) : null,
    f1: iF1 >= 0 ? toNum(r[iF1]) : null, f2: iF2 >= 0 ? toNum(r[iF2]) : null,
    f3: iF3 >= 0 ? toNum(r[iF3]) : null, f4: iF4 >= 0 ? toNum(r[iF4]) : null,
    totalAulas: iTotalAulas >= 0 ? toNum(r[iTotalAulas]) : null,
  })).filter(r => r.ESTUDANTE && r.TURMA && r.DISCIPLINA);
  parsed.forEach(r => {
    if (!r.a1 && !r.a2 && !r.a3 && !r.a4 && r.totalAulas) {
      const aulasP = +(r.totalAulas / 4).toFixed(0);
      if (r.f1 !== null) r.a1 = aulasP;
      if (r.f2 !== null) r.a2 = aulasP;
      if (r.f3 !== null) r.a3 = aulasP;
      if (r.f4 !== null) r.a4 = aulasP;
    }
    // Validar faltas <= aulas por bimestre
    [1, 2, 3, 4].forEach(b => {
      const a = r[`a${b}`], f = r[`f${b}`];
      if (a !== null && f !== null && f > a) { r[`f${b}`] = null; invalidos++; }
    });
  });
  return { records: parsed, invalidos };
}

function parseDigitalSheet(rows) {
  if (!rows.length) return { records: [], invalidos: 0 };
  const h = rows[0].map(c => String(c || '').trim().toUpperCase());
  const iT = colIdx(h, 'TURMA'), iD = colIdx(h, 'DISCIPLINA');
  const iPrev1 = colIdx(h, 'PREVISTO 1', 'PREV1', 'PREV 1');
  const iConc1 = colIdx(h, 'CONCLUIDO 1', 'CONC1', 'CONC 1');
  const iPrev2 = colIdx(h, 'PREVISTO 2', 'PREV2', 'PREV 2');
  const iConc2 = colIdx(h, 'CONCLUIDO 2', 'CONC2', 'CONC 2');
  const iPrev3 = colIdx(h, 'PREVISTO 3', 'PREV3', 'PREV 3');
  const iConc3 = colIdx(h, 'CONCLUIDO 3', 'CONC3', 'CONC 3');
  const iPrev4 = colIdx(h, 'PREVISTO 4', 'PREV4', 'PREV 4');
  const iConc4 = colIdx(h, 'CONCLUIDO 4', 'CONC4', 'CONC 4');
  if (iT < 0 || iD < 0) return { records: [], invalidos: 0 };
  const records = rows.slice(1).map(r => ({
    TURMA: normalizar(r[iT]),
    DISCIPLINA: normalizar(r[iD]),
    prev1: iPrev1 >= 0 ? toNum(r[iPrev1]) : null, conc1: iConc1 >= 0 ? toNum(r[iConc1]) : null,
    prev2: iPrev2 >= 0 ? toNum(r[iPrev2]) : null, conc2: iConc2 >= 0 ? toNum(r[iConc2]) : null,
    prev3: iPrev3 >= 0 ? toNum(r[iPrev3]) : null, conc3: iConc3 >= 0 ? toNum(r[iConc3]) : null,
    prev4: iPrev4 >= 0 ? toNum(r[iPrev4]) : null, conc4: iConc4 >= 0 ? toNum(r[iConc4]) : null,
  })).filter(r => r.TURMA && r.DISCIPLINA);
  return { records, invalidos: 0 };
}

// ── Upload ─────────────────────────────────────────────────

function lerArquivo(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'binary', raw: false });
        let notas = [], faltas = [], digital = [], invalidos = 0;
        wb.SheetNames.forEach(name => {
          const ws = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
          const nUp = name.trim().toUpperCase();
          if (nUp === 'NOTAS' || nUp.includes('NOTA')) {
            const r = parseNotasSheet(ws); notas = r.records; invalidos += r.invalidos;
          } else if (nUp === 'FALTAS' || nUp.includes('FALT')) {
            const r = parseFaltasSheet(ws); faltas = r.records; invalidos += r.invalidos;
          } else if (nUp === 'DIGITAL' || nUp.includes('DIGITAL')) {
            const r = parseDigitalSheet(ws); digital = r.records;
          } else {
            const r = parseNotasSheet(ws); if (r.records.length) notas.push(...r.records); invalidos += r.invalidos;
          }
        });
        resolve({ notas, faltas, digital, invalidos });
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsBinaryString(file);
  });
}

function dragOver(e, id) { e.preventDefault(); document.getElementById(id)?.classList.add('over'); }
function dragLeave(id) { document.getElementById(id)?.classList.remove('over'); }
function drop(e, id) { e.preventDefault(); dragLeave(id); processarArquivos(e.dataTransfer.files); }
async function handleUpload(input) { if (input.files.length) await processarArquivos(input.files); }

async function processarArquivos(files) {
  loader(true, 'Lendo planilha...');
  const statusEl = document.getElementById('status-up');
  let totalInvalidos = 0;
  try {
    for (const f of files) {
      const { notas, faltas, digital, invalidos } = await lerArquivo(f);
      totalInvalidos += invalidos;
      const chN = new Set(APP.notas.map(r => chave(r.ESTUDANTE, r.TURMA, r.DISCIPLINA) + r.b1 + r.b2 + r.b3 + r.b4));
      notas.forEach(r => { if (!chN.has(chave(r.ESTUDANTE, r.TURMA, r.DISCIPLINA) + r.b1 + r.b2 + r.b3 + r.b4)) APP.notas.push(r); });
      const chF = new Set(APP.faltas.map(r => chave(r.ESTUDANTE, r.TURMA, r.DISCIPLINA)));
      faltas.forEach(r => { if (!chF.has(chave(r.ESTUDANTE, r.TURMA, r.DISCIPLINA))) APP.faltas.push(r); });
      const chD = new Set(APP.digital.map(r => `${r.TURMA}|${r.DISCIPLINA}`));
      digital.forEach(r => { if (!chD.has(`${r.TURMA}|${r.DISCIPLINA}`)) APP.digital.push(r); });
    }
    if (!APP.notas.length) throw new Error('Nenhuma nota encontrada. Verifique se a aba se chama "Notas".');
    const turmas = unique(APP.notas.map(r => r.TURMA));
    const discs = unique(APP.notas.map(r => r.DISCIPLINA));
    const digTxt = APP.digital.length ? ` · ${APP.digital.length} registro(s) digital` : '';
    const invTxt = totalInvalidos > 0
      ? `<br><span style="color:var(--amber);">⚠ ${totalInvalidos} valor(es) inválido(s) descartado(s) — notas fora do intervalo [0–10] ou faltas maiores que aulas registradas.</span>`
      : '';
    statusEl.innerHTML = `<span style="color:var(--verde);">✓ ${APP.notas.length} registros de notas · ${APP.faltas.length} de frequência${digTxt} · ${turmas.length} turma(s) · ${discs.length} disciplina(s)</span>${invTxt}`;
    buildNav();
    showSection('geral');
  } catch (err) {
    statusEl.innerHTML = `<span style="color:var(--verm);">Erro: ${err.message}</span>`;
  }
  loader(false);
}

function loader(sim, msg) {
  const el = document.getElementById('loader');
  if (!el) return;
  el.style.display = sim ? 'flex' : 'none';
  if (msg) { const m = document.getElementById('loader-msg'); if (m) m.textContent = msg; }
}

// ── Navegação ──────────────────────────────────────────────

function buildNav() {
  const turmas = unique(APP.notas.map(r => r.TURMA));
  const discs = unique(APP.notas.map(r => r.DISCIPLINA));
  document.getElementById('nav-turmas').innerHTML = turmas.map(t =>
    `<a class="nav-link" data-t="${t}" onclick="selTurma('${t}')"><i class="bi bi-people"></i> ${t}</a>`).join('');
  document.getElementById('nav-discs').innerHTML = discs.map(d =>
    `<a class="nav-link" data-d="${encodeURIComponent(d)}" onclick="selDisc('${encodeURIComponent(d)}')"><i class="bi bi-journal-bookmark"></i> ${d}</a>`).join('');
  const countT = document.getElementById('acc-turmas-count');
  const countD = document.getElementById('acc-discs-count');
  if (countT) { countT.textContent = turmas.length; countT.classList.add('visible'); }
  if (countD) { countD.textContent = discs.length; countD.classList.add('visible'); }
  if (turmas.length) openAcc('acc-turmas');
  if (discs.length) openAcc('acc-discs');
  const btn = document.getElementById('btn-pptx-topbar');
  if (btn) btn.style.display = '';
  // Popular selects globais
  [['filtro-turma-medalhas', 'Todas as turmas'], ['filtro-digital-turma', 'Todas as turmas'], ['radar-sel-turma', 'Todas as turmas']].forEach(([id, label]) => {
    const el = document.getElementById(id);
    if (el) el.innerHTML = `<option value="">${label}</option>` + turmas.map(t => `<option value="${t}">${t}</option>`).join('');
  });
  [['filtro-digital-disc', 'Todas as disciplinas'], ['radar-sel-disc', 'Todas as disciplinas']].forEach(([id, label]) => {
    const el = document.getElementById(id);
    if (el) el.innerHTML = `<option value="">${label}</option>` + discs.map(d => `<option value="${d}">${d}</option>`).join('');
  });
}

function openAcc(id) {
  const acc = document.getElementById(id);
  if (!acc) return;
  const toggle = acc.querySelector('.sb-acc-toggle'), body = acc.querySelector('.sb-acc-body');
  if (!toggle || !body) return;
  toggle.classList.add('open'); toggle.setAttribute('aria-expanded', 'true'); body.classList.add('open');
}

function toggleAcc(id) {
  const acc = document.getElementById(id);
  if (!acc) return;
  const toggle = acc.querySelector('.sb-acc-toggle'), body = acc.querySelector('.sb-acc-body');
  if (!toggle || !body) return;
  const isOpen = body.classList.contains('open');
  toggle.classList.toggle('open', !isOpen); toggle.setAttribute('aria-expanded', String(!isOpen)); body.classList.toggle('open', !isOpen);
}

function showSection(id) {
  document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
  document.querySelectorAll('#sidebar .nav-link').forEach(l => l.classList.remove('active'));
  document.getElementById('section-' + id)?.classList.add('active');
  document.querySelectorAll(`[data-s="${id}"]`).forEach(l => l.classList.add('active'));
  const titles = {
    upload: ['Importar dados', 'Selecione os arquivos para começar'],
    geral: ['Painel geral', 'Visão consolidada'],
    turma: ['Turma ' + (APP.turmaSel || ''), 'Análise por turma'],
    disc: [APP.discSel || 'Disciplina', 'Análise e intervenções pedagógicas'],
    medalhas: ['🏆 Medalhistas', 'Ranking por desempenho e frequência'],
    digital: ['💻 Material Digital', 'Progresso do conteúdo digital'],
    radar: ['🎯 Radar Preditivo', 'Triagem e priorização para o conselho'],
    relatorios: ['Relatórios PDF', 'Gere relatórios para impressão'],
    pptx: ['Apresentação PPTX', 'Gerar relatório em PowerPoint'],
  };
  const [t, s] = titles[id] || ['', ''];
  const tb = document.getElementById('tb-title'), sb = document.getElementById('tb-sub');
  if (tb) tb.textContent = t; if (sb) sb.textContent = s;
  if (id === 'geral') renderGeral();
  if (id === 'turma') renderTurma();
  if (id === 'disc') renderDisc();
  if (id === 'medalhas') renderMedalhas();
  if (id === 'digital') renderDigital();
  if (id === 'radar') renderRadar();
  if (id === 'relatorios') popSelects();
  if (id === 'pptx') popSelectsPPTX();
}

function selTurma(t) {
  APP.turmaSel = t;
  document.querySelectorAll('[data-t]').forEach(el => el.classList.remove('active'));
  document.querySelectorAll(`[data-t="${t}"]`).forEach(el => el.classList.add('active'));
  showSection('turma');
}
function selDisc(d) {
  APP.discSel = decodeURIComponent(d);
  document.querySelectorAll('[data-d]').forEach(el => el.classList.remove('active'));
  document.querySelectorAll(`[data-d="${d}"]`).forEach(el => el.classList.add('active'));
  showSection('disc');
}

// ── PAINEL GERAL ───────────────────────────────────────────

function renderGeral() {
  if (!APP.notas.length) return;
  const turmas = unique(APP.notas.map(r => r.TURMA));
  const discs = unique(APP.notas.map(r => r.DISCIPLINA));
  const alunos = unique(APP.notas.map(r => r.ESTUDANTE));
  let aprov = 0, reprov = 0, alFreqs = 0, freqCritica = 0;
  alunos.forEach(a => {
    const regs = APP.notas.filter(r => r.ESTUDANTE === a);
    const sits = regs.map(r => dadosReg(r));
    if (sits.some(d => d.sit === 'Reprovado')) reprov++; else aprov++;
    if (sits.some(d => d.alFreq)) alFreqs++;
    const fg = calcFreqGeral(a, regs[0]?.TURMA || '');
    if (fg && fg.pct < CFG.freqRisco) freqCritica++;
  });
  const medias = APP.notas.map(r => calcMedia(r)).filter(m => m !== null);
  const mg = medias.length ? +(medias.reduce((a, b) => a + b, 0) / medias.length).toFixed(1) : 0;
  const [pctAp, pctRp] = distribuirPct([aprov, reprov]);

  document.getElementById('mg-geral').innerHTML = `
    <div class="mc bl"><div class="ml">Total de alunos</div><div class="mv c-bl">${alunos.length}</div><div class="ms">${turmas.length} turmas · ${discs.length} disciplinas</div></div>
    <div class="mc gr"><div class="ml">Aprovados</div><div class="mv c-gr">${aprov}</div><div class="ms">${pctAp}% do total</div></div>
    <div class="mc vm"><div class="ml">Reprovados</div><div class="mv c-vm">${reprov}</div><div class="ms">${pctRp}% do total</div></div>
    <div class="mc am"><div class="ml">Alerta de frequência</div><div class="mv c-am">${alFreqs}</div><div class="ms">freq. &lt; 75%</div></div>
    <div class="mc rx"><div class="ml">Média geral</div><div class="mv c-rx">${mg.toFixed(1).replace('.', ',')}</div></div>`;

  // Métricas de frequência — bloco mg-freq-geral
  let totalAulasGlobal = 0, totalFaltasGlobal = 0;
  APP.faltas.forEach(r => {
    [1, 2, 3, 4].forEach(b => {
      const a = toNum(r[`a${b}`]), f = toNum(r[`f${b}`]);
      if (a && a > 0) { totalAulasGlobal += a; totalFaltasGlobal += (f || 0); }
    });
  });
  const freqMediaGlobal = totalAulasGlobal > 0 ? +((1 - totalFaltasGlobal / totalAulasGlobal) * 100).toFixed(1) : null;
  const fmgCls = freqMediaGlobal === null ? 'c-rx' : freqMediaGlobal >= 85 ? 'c-gr' : freqMediaGlobal >= CFG.freqMin ? 'c-am' : 'c-vm';
  const mgFreqEl = document.getElementById('mg-freq-geral');
  if (mgFreqEl) {
    mgFreqEl.style.display = '';
    mgFreqEl.innerHTML = `
      <div class="mc am" style="border-left:4px solid var(--amber);"><div class="ml"><i class="bi bi-exclamation-triangle-fill"></i> Alerta freq. (&lt;75%)</div><div class="mv c-am">${alFreqs}</div><div class="ms">alunos abaixo do mínimo legal</div></div>
      <div class="mc vm" style="border-left:4px solid var(--verm);"><div class="ml"><i class="bi bi-calendar-x-fill"></i> Freq. crítica (&lt;${CFG.freqRisco}%)</div><div class="mv c-vm">${freqCritica}</div><div class="ms">encaminhar para a gestão</div></div>
      <div class="mc" style="border-left:4px solid var(--azul);"><div class="ml"><i class="bi bi-calendar-check"></i> Presença média geral</div><div class="mv ${fmgCls}">${freqMediaGlobal !== null ? freqMediaGlobal.toFixed(1) + '%' : '—'}</div><div class="ms">média de todas as disciplinas</div></div>`;
  }

  // Gráfico situação por turma
  destroyChart('ch-sit-turmas');
  const apT = [], rpT = [];
  turmas.forEach(t => {
    const al = unique(APP.notas.filter(r => r.TURMA === t).map(r => r.ESTUDANTE));
    let ap = 0, rp = 0;
    al.forEach(a => { const regs = APP.notas.filter(r => r.ESTUDANTE === a && r.TURMA === t); if (regs.some(r => dadosReg(r).sit === 'Reprovado')) rp++; else ap++; });
    apT.push(ap); rpT.push(rp);
  });
  APP.charts['ch-sit-turmas'] = new Chart(document.getElementById('ch-sit-turmas'), {
    type: 'bar', plugins: [pluginBarPct],
    data: { labels: turmas, datasets: [{ label: 'Aprovado', data: apT, backgroundColor: '#057a55' }, { label: 'Reprovado', data: rpT, backgroundColor: '#c81e1e' }] },
    options: optsBarEmpilhado()
  });

  // Gráfico média por disciplina
  destroyChart('ch-media-disc');
  const mD = discs.map(d => { const ms = APP.notas.filter(r => r.DISCIPLINA === d).map(r => calcMedia(r)).filter(m => m !== null); return ms.length ? +(ms.reduce((a, b) => a + b, 0) / ms.length).toFixed(2) : 0; });
  APP.charts['ch-media-disc'] = new Chart(document.getElementById('ch-media-disc'), {
    type: 'bar', plugins: [pluginBarValor],
    data: { labels: discs, datasets: [{ label: 'Média', data: mD, backgroundColor: mD.map(m => m >= CFG.mediaAprov ? '#057a55' : '#c81e1e') }] },
    options: optsBarHoriz()
  });

  // Gráfico frequência por turma
  const chFreqEl = document.getElementById('ch-freq-turmas');
  if (chFreqEl) {
    destroyChart('ch-freq-turmas');
    const freqPorTurma = turmas.map(t => {
      const regsT = APP.faltas.filter(r => r.TURMA === t);
      if (!regsT.length) return null;
      let totalA = 0, totalF = 0;
      regsT.forEach(r => { [1, 2, 3, 4].forEach(b => { const a = toNum(r[`a${b}`]), f = toNum(r[`f${b}`]); if (a && a > 0) { totalA += a; totalF += (f || 0); } }); });
      return totalA ? +((1 - totalF / totalA) * 100).toFixed(1) : null;
    });
    APP.charts['ch-freq-turmas'] = new Chart(chFreqEl, {
      type: 'bar', plugins: [pluginBarValor],
      data: { labels: turmas, datasets: [{ label: 'Frequência média (%)', data: freqPorTurma, backgroundColor: freqPorTurma.map(f => f === null ? '#9ca3af' : f >= 85 ? '#057a55' : f >= CFG.freqMin ? '#b45309' : '#c81e1e') }] },
      options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => ctx.raw !== null ? ctx.raw.toFixed(1) + '%' : '—' } } }, scales: { x: { min: 0, max: 100, ticks: { callback: v => v + '%' } } } }
    });
  }

  // Gráfico Material Digital
  if (APP.digital.length) {
    const cardDig = document.getElementById('card-digital-geral');
    if (cardDig) cardDig.style.display = '';
    const discsDig = unique(APP.digital.map(r => r.DISCIPLINA));
    const pctsPorDisc = discsDig.map(d => { const r = APP.digital.filter(x => x.DISCIPLINA === d); const c = calcDigital(r); return c ? c.pct : 0; });
    destroyChart('ch-digital-geral');
    APP.charts['ch-digital-geral'] = new Chart(document.getElementById('ch-digital-geral'), {
      type: 'bar',
      data: { labels: discsDig, datasets: [{ label: '% Cumprimento Digital', data: pctsPorDisc, backgroundColor: pctsPorDisc.map(p => p >= 80 ? '#6c2bd9' : p >= 50 ? '#a855f7' : '#c084fc'), borderRadius: 4 }] },
      options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => `${ctx.raw.toFixed(1)}%` } } }, scales: { x: { min: 0, max: 100, ticks: { callback: v => v + '%' } } } }
    });
  } else {
    const cardDig = document.getElementById('card-digital-geral');
    if (cardDig) cardDig.style.display = 'none';
  }

  document.getElementById('tb-resumo').innerHTML = turmas.map(t => {
    const al = unique(APP.notas.filter(r => r.TURMA === t).map(r => r.ESTUDANTE));
    let ap = 0, rp = 0, alF = 0, fCrit = 0;
    al.forEach(a => {
      const ds = APP.notas.filter(r => r.ESTUDANTE === a && r.TURMA === t).map(r => dadosReg(r));
      if (ds.some(d => d.sit === 'Reprovado')) rp++; else ap++;
      if (ds.some(d => d.alFreq)) alF++;
      const fg = calcFreqGeral(a, t); if (fg && fg.pct < CFG.freqRisco) fCrit++;
    });
    const ms = APP.notas.filter(r => r.TURMA === t).map(r => calcMedia(r)).filter(m => m !== null);
    const mg2 = ms.length ? +(ms.reduce((a, b) => a + b, 0) / ms.length).toFixed(1) : 0;
    // Freq. média da turma
    const regsF = APP.faltas.filter(r => r.TURMA === t);
    let tA = 0, tF = 0;
    regsF.forEach(r => { [1, 2, 3, 4].forEach(b => { const a = toNum(r[`a${b}`]), f = toNum(r[`f${b}`]); if (a && a > 0) { tA += a; tF += (f || 0); } }); });
    const fmT = tA ? +((1 - tF / tA) * 100).toFixed(1) : null;
    const fmTCls = fmT === null ? 'n-ne' : fmT >= 85 ? 'freq-ok' : fmT >= CFG.freqMin ? 'freq-al' : 'freq-ruim';
    return `<tr class="click" onclick="selTurma('${t}')">
      <td><strong>${t}</strong></td><td>${al.length}</td>
      <td><span class="badge b-ap">${ap}</span></td>
      <td><span class="badge b-re">${rp}</span></td>
      <td>${alF > 0 ? `<span class="badge b-al">${alF}</span>` : '<span style="color:var(--cinza)">—</span>'}</td>
      <td>${fCrit > 0 ? `<span class="badge b-re">${fCrit}</span>` : '<span style="color:var(--cinza)">—</span>'}</td>
      <td><span class="${fmTCls}">${fmT !== null ? fmT.toFixed(1) + '%' : '—'}</span></td>
      <td class="nota ${mg2 >= CFG.mediaAprov ? 'n-ok' : 'n-re'}">${mg2.toFixed(1).replace('.', ',')}</td>
    </tr>`;
  }).join('');
}

// ── VISÃO TURMA ────────────────────────────────────────────

function renderTurma() {
  const t = APP.turmaSel;
  if (!t) return;
  const todasRegs = APP.notas.filter(r => r.TURMA === t);
  if (!todasRegs.length) return;
  const todasDiscs = unique(todasRegs.map(r => r.DISCIPLINA));
  const titEl = document.getElementById('titulo-alunos'), tbEl = document.getElementById('tb-title');
  if (titEl) titEl.textContent = `Alunos — Turma ${t}`;
  if (tbEl) tbEl.textContent = `Turma ${t}`;
  const selD = document.getElementById('filtro-disc-turma');
  if (selD) { const prev = selD.value; selD.innerHTML = '<option value="">Todas as disciplinas</option>' + todasDiscs.map(d => `<option value="${d}"${d === prev ? ' selected' : ''}>${d}</option>`).join(''); }
  const buscaEl = document.getElementById('busca'), sitEl = document.getElementById('filtro-sit');
  if (buscaEl) buscaEl.value = ''; if (sitEl) sitEl.value = '';
  renderTurmaComFiltro();
  renderDigitalTurma(t);
}

function renderDigitalTurma(turma) {
  const sec = document.getElementById('digital-turma-section'), body = document.getElementById('digital-turma-body');
  const regs = APP.digital.filter(r => r.TURMA === turma);
  if (!regs.length) { if (sec) sec.style.display = 'none'; return; }
  if (sec) sec.style.display = '';
  const discs = unique(regs.map(r => r.DISCIPLINA));
  if (body) body.innerHTML = `<div class="digital-grid">${discs.map(d => fmtDigitalCard(turma, d, regs.filter(r => r.DISCIPLINA === d))).join('')}</div>`;
}

function renderTurmaComFiltro() {
  const t = APP.turmaSel;
  const disc = document.getElementById('filtro-disc-turma')?.value || '';
  const regs = APP.notas.filter(r => r.TURMA === t && (!disc || r.DISCIPLINA === disc));
  if (!regs.length) return;
  const discs = unique(regs.map(r => r.DISCIPLINA));
  const alunos = unique(regs.map(r => r.ESTUDANTE));
  let ap = 0, rp = 0, alF = 0;
  alunos.forEach(a => { const ds = regs.filter(r => r.ESTUDANTE === a).map(r => dadosReg(r)); if (ds.some(d => d.sit === 'Reprovado')) rp++; else ap++; if (ds.some(d => d.alFreq)) alF++; });
  const ms = regs.map(r => calcMedia(r)).filter(m => m !== null);
  const mg = ms.length ? +(ms.reduce((a, b) => a + b, 0) / ms.length).toFixed(1) : 0;
  const discLabel = disc || `${unique(APP.notas.filter(r => r.TURMA === t).map(r => r.DISCIPLINA)).length} disciplinas`;
  const [pctApT, pctRpT] = distribuirPct([ap, rp]);
  document.getElementById('mg-turma').innerHTML = `
    <div class="mc bl"><div class="ml">Alunos</div><div class="mv c-bl">${alunos.length}</div><div class="ms">${discLabel}</div></div>
    <div class="mc gr"><div class="ml">Aprovados</div><div class="mv c-gr">${ap}</div><div class="ms">${pctApT}%</div></div>
    <div class="mc vm"><div class="ml">Reprovados</div><div class="mv c-vm">${rp}</div><div class="ms">${pctRpT}%</div></div>
    <div class="mc am"><div class="ml">Alerta freq.</div><div class="mv c-am">${alF}</div></div>
    <div class="mc rx"><div class="ml">Média geral</div><div class="mv c-rx">${mg.toFixed(1).replace('.', ',')}</div></div>`;
  destroyChart('ch-turma-disc');
  const mD = discs.map(d => { const ms2 = regs.filter(r => r.DISCIPLINA === d).map(r => calcMedia(r)).filter(m => m !== null); return ms2.length ? +(ms2.reduce((a, b) => a + b, 0) / ms2.length).toFixed(2) : 0; });
  APP.charts['ch-turma-disc'] = new Chart(document.getElementById('ch-turma-disc'), { type: 'bar', plugins: [pluginBarValor], data: { labels: discs, datasets: [{ label: 'Média', data: mD, backgroundColor: mD.map(m => m >= CFG.mediaAprov ? '#057a55' : '#c81e1e') }] }, options: optsBarHoriz() });
  destroyChart('ch-turma-sit');
  const total = ap + rp || 1;
  APP.charts['ch-turma-sit'] = new Chart(document.getElementById('ch-turma-sit'), {
    type: 'doughnut', plugins: [makePctLabelPlugin(total)],
    data: { labels: ['Aprovado', 'Reprovado'], datasets: [{ data: [ap, rp], backgroundColor: ['#057a55', '#c81e1e'], borderWidth: 2 }] },
    options: { responsive: true, maintainAspectRatio: false, cutout: '52%', plugins: { legend: { position: 'bottom', labels: { font: { size: 12 }, boxWidth: 12 } }, tooltip: { callbacks: { label: ctx => { const vals = ctx.chart.data.datasets[ctx.datasetIndex].data; const pcts = distribuirPct(vals.map(v => v || 0)); return `${ctx.label}: ${ctx.raw} (${pcts[ctx.dataIndex]}%)`; } } } } }
  });
  renderAlunos();
}

function renderAlunos() {
  const t = APP.turmaSel;
  const disc = document.getElementById('filtro-disc-turma')?.value || '';
  const sit = document.getElementById('filtro-sit')?.value || '';
  const bsc = (document.getElementById('busca')?.value || '').toLowerCase();
  const tbody = document.getElementById('tb-alunos');
  const thead = document.getElementById('thead-alunos');
  if (!tbody || !t) return;

  if (!disc) {
    // ── MODO RESUMO: uma linha por aluno ─────────────────────
    if (thead) thead.innerHTML = `<tr>
      <th>Aluno</th><th style="text-align:center;">Discs.</th>
      <th style="text-align:center;">Reprov.</th><th>Freq. Geral</th>
      <th>Situação Geral</th><th>Prioridade</th>
    </tr>`;
    const alunos = unique(
      APP.notas.filter(r => r.TURMA === t && (!bsc || r.ESTUDANTE.toLowerCase().includes(bsc))).map(r => r.ESTUDANTE)
    );
    const linhas = [];
    alunos.forEach(nome => {
      const regs = APP.notas.filter(r => r.ESTUDANTE === nome && r.TURMA === t);
      const dados = regs.map(r => dadosReg(r));
      const reprovCount = dados.filter(d => d.sit === 'Reprovado').length;
      const totalDiscs = dados.length;
      const temAlFreq = dados.some(d => d.alFreq);
      const freqGeral = calcFreqGeral(nome, t);
      const nivelFG = nivelRiscoFreq(freqGeral);
      const sitGeral = reprovCount > 0 ? 'Reprovado' : 'Aprovado';
      // Filtros de situação
      if (sit === 'Aprovado' && sitGeral !== 'Aprovado') return;
      if (sit === 'Reprovado' && sitGeral !== 'Reprovado') return;
      if (sit === 'Alerta freq.' && !temAlFreq) return;
      if (sit === 'Risco combinado' && !(sitGeral === 'Reprovado' && temAlFreq)) return;
      // Badge situação geral
      let sitLabel, sitCls;
      if (reprovCount === 0) {
        if (nivelFG === 'critico') { sitLabel = 'Aprovado ⚠ ret. freq.'; sitCls = 'b-re'; }
        else if (temAlFreq) { sitLabel = 'Aprovado ⚠ freq.'; sitCls = 'b-al'; }
        else { sitLabel = 'Aprovado'; sitCls = 'b-ap'; }
      } else {
        sitLabel = `${reprovCount} reprov. de ${totalDiscs}`; sitCls = 'b-re';
      }
      // Badge prioridade
      const colapso = reprovCount >= Math.ceil(totalDiscs * 0.5) && nivelFG === 'critico';
      let prioridade;
      if (colapso || nivelFG === 'critico') {
        prioridade = `<span class="badge b-re"><i class="bi bi-exclamation-octagon-fill"></i> Urgente</span>`;
      } else if (reprovCount > 0 && temAlFreq) {
        prioridade = `<span class="badge b-al"><i class="bi bi-exclamation-triangle-fill"></i> Aten\u00e7\u00e3o</span>`;
      } else if (reprovCount > 0 || temAlFreq) {
        prioridade = `<span class="badge b-al" style="opacity:.75;"><i class="bi bi-eye-fill"></i> Monitorar</span>`;
      } else {
        prioridade = `<span class="badge b-ap"><i class="bi bi-check-circle-fill"></i> Regular</span>`;
      }
      linhas.push(`<tr class="click modal-trigger" data-nome="${encodeURIComponent(nome)}" data-turma="${encodeURIComponent(t)}">
        <td style="font-size:12px;"><strong>${nome}</strong></td>
        <td style="text-align:center;font-size:12px;color:var(--cinza);">${totalDiscs}</td>
        <td style="text-align:center;">${reprovCount > 0 ? `<span class="nota n-re">${reprovCount}</span>` : `<span class="nota n-ok">0</span>`}</td>
        <td>${fmtFreqCell(freqGeral)}</td>
        <td><span class="badge ${sitCls}">${sitLabel}</span></td>
        <td>${prioridade}</td>
      </tr>`);
    });
    tbody.innerHTML = linhas.join('') || '<tr><td colspan="6" style="text-align:center;color:var(--cinza);padding:20px;">Nenhum registro encontrado</td></tr>';
  } else {
    // ── MODO DETALHE: uma linha por aluno × disciplina ───────
    if (thead) thead.innerHTML = `<tr>
      <th>Aluno</th><th>Disciplina</th>
      <th>B1</th><th>B2</th><th>B3</th><th>B4</th>
      <th>Evolu\u00e7\u00e3o</th><th>M\u00e9dia</th><th>Frequ\u00eancia</th><th>Situa\u00e7\u00e3o</th>
    </tr>`;
    const regs = APP.notas.filter(r => r.TURMA === t && r.DISCIPLINA === disc && (!bsc || r.ESTUDANTE.toLowerCase().includes(bsc)));
    const linhas = [];
    regs.forEach(r => {
      const d = dadosReg(r);
      if (sit === 'Aprovado' && d.sit !== 'Aprovado') return;
      if (sit === 'Reprovado' && d.sit !== 'Reprovado') return;
      if (sit === 'Alerta freq.' && !d.alFreq) return;
      if (sit === 'Risco combinado' && !(d.sit === 'Reprovado' && d.alFreq)) return;
      linhas.push(`<tr class="click modal-trigger" data-nome="${encodeURIComponent(r.ESTUDANTE)}" data-turma="${encodeURIComponent(t)}">
        <td style="font-size:12px;"><strong>${r.ESTUDANTE}</strong></td>
        <td style="font-size:12px;color:var(--cinza);">${r.DISCIPLINA}</td>
        <td>${fmtNota(toNum(r.b1))}</td><td>${fmtNota(toNum(r.b2))}</td><td>${fmtNota(toNum(r.b3))}</td><td>${fmtNota(toNum(r.b4))}</td>
        <td>${fmtEvol(r)}</td><td>${fmtMedia(d.media, d.parc)}</td><td>${fmtFreqCell(d.freq)}</td><td>${fmtSit(d.media, d.freq, d.parc)}</td>
      </tr>`);
    });
    tbody.innerHTML = linhas.join('') || '<tr><td colspan="10" style="text-align:center;color:var(--cinza);padding:20px;">Nenhum registro encontrado</td></tr>';
  }
}


// ── VISÃO DISCIPLINA ───────────────────────────────────────

function renderDisc() {
  const disc = APP.discSel;
  if (!disc) return;
  const todasRegs = APP.notas.filter(r => r.DISCIPLINA === disc);
  if (!todasRegs.length) return;
  const turmas = unique(todasRegs.map(r => r.TURMA));
  const titEl = document.getElementById('titulo-disc'), tbEl = document.getElementById('tb-title');
  if (titEl) titEl.textContent = disc; if (tbEl) tbEl.textContent = disc;
  const sel = document.getElementById('filtro-turma-disc');
  if (sel) { const prev = sel.value; sel.innerHTML = '<option value="">Todas as turmas</option>' + turmas.map(t => `<option value="${t}"${t === prev ? ' selected' : ''}>${t}</option>`).join(''); }
  renderDiscComFiltro();
}

function renderDiscComFiltro() {
  const disc = APP.discSel;
  const ft = document.getElementById('filtro-turma-disc')?.value || '';
  const regs = APP.notas.filter(r => r.DISCIPLINA === disc && (!ft || r.TURMA === ft));
  if (!regs.length) return;
  const turmas = unique(APP.notas.filter(r => r.DISCIPLINA === disc).map(r => r.TURMA));
  const label = document.getElementById('iv-turma-label');
  if (label) label.textContent = ft ? `— ${ft}` : '';
  let ap = 0, rp = 0, alF = 0;
  const ms = [];
  regs.forEach(r => { const d = dadosReg(r); if (d.media !== null) ms.push(d.media); if (d.sit === 'Reprovado') rp++; else if (d.sit === 'Aprovado') ap++; if (d.alFreq) alF++; });
  const mg = ms.length ? +(ms.reduce((a, b) => a + b, 0) / ms.length).toFixed(1) : 0;
  const [pctApD, pctRpD] = distribuirPct([ap, rp]);
  document.getElementById('mg-disc').innerHTML = `
    <div class="mc bl"><div class="ml">Registros</div><div class="mv c-bl">${regs.length}</div><div class="ms">${ft || turmas.length + ' turma(s)'}</div></div>
    <div class="mc gr"><div class="ml">Aprovados</div><div class="mv c-gr">${ap}</div><div class="ms">${ms.length ? pctApD + '%' : ''}</div></div>
    <div class="mc vm"><div class="ml">Reprovados</div><div class="mv c-vm">${rp}</div><div class="ms">${ms.length ? pctRpD + '%' : ''}</div></div>
    <div class="mc am"><div class="ml">Alerta freq.</div><div class="mv c-am">${alF}</div></div>
    <div class="mc rx"><div class="ml">Média geral</div><div class="mv c-rx">${mg.toFixed(1).replace('.', ',')}</div></div>`;
  destroyChart('ch-disc-evol');
  const cores = ['#1a56db', '#057a55', '#b45309', '#c81e1e', '#6c2bd9'];
  const turmasGraf = ft ? [ft] : turmas;
  APP.charts['ch-disc-evol'] = new Chart(document.getElementById('ch-disc-evol'), {
    type: 'line', plugins: [pluginLinePct],
    data: {
      labels: ['1º Bim', '2º Bim', '3º Bim', '4º Bim'], datasets: turmasGraf.map((t, i) => {
        const tr = regs.filter(r => r.TURMA === t);
        return {
          label: t, borderColor: cores[i % cores.length], backgroundColor: 'transparent', pointBackgroundColor: cores[i % cores.length], tension: 0.3, spanGaps: true,
          data: ['b1', 'b2', 'b3', 'b4'].map(b => { const vs = tr.map(r => toNum(r[b])).filter(v => v !== null); return vs.length ? +(vs.reduce((a, c) => a + c, 0) / vs.length).toFixed(2) : null; })
        };
      })
    },
    options: optsLinha()
  });
  destroyChart('ch-disc-dist');
  const faixas = ['0–2', '2–4', '4–5', '5–7', '7–10'], cnt = [0, 0, 0, 0, 0];
  ms.forEach(m => { if (m < 2) cnt[0]++; else if (m < 4) cnt[1]++; else if (m < 5) cnt[2]++; else if (m < 7) cnt[3]++; else cnt[4]++; });
  const pctsDist = distribuirPct(cnt);
  const pluginDistPct = { id: 'distPct', afterDatasetsDraw(chart) { const { ctx, data } = chart; data.datasets[0].data.forEach((val, i) => { if (!val) return; const meta = chart.getDatasetMeta(0), bar = meta.data[i], pct = pctsDist[i]; ctx.save(); ctx.fillStyle = '#374151'; ctx.font = 'bold 11px sans-serif'; ctx.textAlign = 'center'; ctx.textBaseline = 'bottom'; ctx.fillText(val, bar.x, bar.y - 2); if (bar.height > 20) { ctx.fillStyle = '#fff'; ctx.textBaseline = 'middle'; ctx.fillText(pct + '%', bar.x, bar.y + bar.height / 2); } ctx.restore(); }); } };
  APP.charts['ch-disc-dist'] = new Chart(document.getElementById('ch-disc-dist'), { type: 'bar', plugins: [pluginDistPct], data: { labels: faixas, datasets: [{ label: 'Alunos', data: cnt, backgroundColor: ['#c81e1e', '#c81e1e', '#b45309', '#057a55', '#057a55'] }] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { ticks: { stepSize: 1 } } }, layout: { padding: { top: 18 } } } });
  renderIntervencoes(regs);
  renderListaDisc();
}

function renderIntervencoes(regs) {
  const criticos = regs.filter(r => { const d = dadosReg(r); return d.media !== null && d.media < 3; });
  const reprovN = regs.filter(r => { const d = dadosReg(r); return d.media !== null && d.media >= 3 && d.media < CFG.mediaAprov; });
  const freqCrit = regs.filter(r => dadosReg(r).nivelFreq === 'critico');
  const alertFreq = regs.filter(r => dadosReg(r).alFreq && dadosReg(r).nivelFreq !== 'critico');
  let html = '';
  if (criticos.length) html += `<div class="iv iv-crit"><i class="bi bi-exclamation-octagon-fill" style="font-size:20px;color:var(--verm);margin-top:2px;flex-shrink:0;"></i><div><div class="iv-title" style="color:var(--verm-esc);">Desempenho crítico — ${criticos.length} aluno(s)</div><div class="iv-desc">${criticos.map(r => `<strong>${r.ESTUDANTE}</strong> (${r.TURMA}) — média ${(calcMedia(r) || 0).toFixed(1).replace('.', ',')}`).join('<br>')}</div><div class="iv-desc" style="margin-top:5px;">→ Encaminhar para reforço intensivo. Verificar compensação de nota e contato com a família.</div></div></div>`;
  if (reprovN.length) html += `<div class="iv iv-alert"><i class="bi bi-exclamation-triangle-fill" style="font-size:20px;color:var(--amber);margin-top:2px;flex-shrink:0;"></i><div><div class="iv-title" style="color:var(--amber-esc);">Abaixo da média — ${reprovN.length} aluno(s)</div><div class="iv-desc">${reprovN.map(r => `<strong>${r.ESTUDANTE}</strong> (${r.TURMA}) — média ${(calcMedia(r) || 0).toFixed(1).replace('.', ',')}`).join('<br>')}</div><div class="iv-desc" style="margin-top:5px;">→ Solicitar compensação de nota. Reforço antes do próximo bimestre.</div></div></div>`;
  if (freqCrit.length) html += `<div class="iv" style="background:var(--verm-cl);border-left:4px solid var(--verm);border-radius:var(--r);padding:11px 14px;margin-bottom:8px;display:flex;gap:10px;align-items:flex-start;"><i class="bi bi-person-x-fill" style="font-size:20px;color:var(--verm);margin-top:2px;flex-shrink:0;"></i><div><div class="iv-title" style="color:var(--verm-esc);">Risco de retenção por frequência — ${freqCrit.length} aluno(s)</div><div class="iv-desc">${freqCrit.map(r => { const freq = calcFreq(getFaltas(r.ESTUDANTE, r.TURMA, r.DISCIPLINA)); return `<strong>${r.ESTUDANTE}</strong> (${r.TURMA}) — ${freq ? freq.pct.toFixed(1) + '%' : '-'} de presença`; }).join('<br>')}</div><div class="iv-desc" style="margin-top:5px;">→ <strong>Encaminhar para a gestão.</strong> Frequência abaixo de ${CFG.freqRisco}% — verificar compensação de ausências.</div></div></div>`;
  if (alertFreq.length) html += `<div class="iv iv-freq"><i class="bi bi-calendar-x-fill" style="font-size:20px;color:#f97316;margin-top:2px;flex-shrink:0;"></i><div><div class="iv-title" style="color:#7c2d12;">Frequência abaixo de 75% — ${alertFreq.length} aluno(s)</div><div class="iv-desc">${alertFreq.map(r => { const freq = calcFreq(getFaltas(r.ESTUDANTE, r.TURMA, r.DISCIPLINA)); return `<strong>${r.ESTUDANTE}</strong> (${r.TURMA}) — ${freq ? freq.pct.toFixed(1) + '%' : '-'}`; }).join('<br>')}</div><div class="iv-desc" style="margin-top:5px;">→ Comunicar família. Encaminhar para acompanhamento da gestão.</div></div></div>`;
  if (!html) html = `<div class="iv iv-ok"><i class="bi bi-check-circle-fill" style="font-size:20px;color:var(--verde);flex-shrink:0;"></i><div><div class="iv-title" style="color:var(--verde-esc);">Sem alertas críticos nesta disciplina</div><div class="iv-desc">Todos os alunos com média ≥ 5,0 e sem alerta de frequência.</div></div></div>`;
  const divIv = document.getElementById('div-iv');
  if (divIv) divIv.innerHTML = html;
}

function onFiltroTurmaDisc() { renderDiscComFiltro(); }

function renderListaDisc() {
  const disc = APP.discSel;
  const ft = document.getElementById('filtro-turma-disc')?.value || '';
  const regs = APP.notas.filter(r => r.DISCIPLINA === disc && (!ft || r.TURMA === ft));
  const tbody = document.getElementById('tb-disc');
  if (!tbody) return;
  tbody.innerHTML = regs.map(r => {
    const d = dadosReg(r), nivel = nivelRiscoFreq(d.freq);
    const alerta = d.media !== null && d.media < 3 ? `<span class="badge b-re">Crítico</span>` :
      d.media !== null && d.media < CFG.mediaAprov ? `<span class="badge b-al">Comp. nota</span>` :
        nivel === 'critico' ? `<span class="badge b-re">Ret. freq.</span>` :
          d.alFreq ? `<span class="badge b-al">Comp. freq.</span>` : `<span style="color:var(--cinza);font-size:11px;">—</span>`;
    return `<tr class="click modal-trigger" data-nome="${encodeURIComponent(r.ESTUDANTE)}" data-turma="${encodeURIComponent(r.TURMA)}">
      <td style="font-size:12px;"><strong>${r.ESTUDANTE}</strong></td>
      <td><span class="badge b-info">${r.TURMA}</span></td>
      <td>${fmtNota(toNum(r.b1))}</td><td>${fmtNota(toNum(r.b2))}</td><td>${fmtNota(toNum(r.b3))}</td><td>${fmtNota(toNum(r.b4))}</td>
      <td>${fmtMedia(d.media, d.parc)}</td><td>${fmtFreqCell(d.freq)}</td><td>${fmtSit(d.media, d.freq, d.parc)}</td><td>${alerta}</td>
    </tr>`;
  }).join('') || '<tr><td colspan="10" style="text-align:center;color:var(--cinza);padding:20px;">Nenhum registro</td></tr>';
}

// ── MATERIAL DIGITAL ───────────────────────────────────────

function renderDigital() {
  const mgDig = document.getElementById('mg-digital');
  if (!APP.digital.length) {
    if (mgDig) mgDig.innerHTML = `<div class="mc" style="grid-column:1/-1;border-left:4px solid var(--roxo);"><div class="ml">Sem dados de material digital</div><div class="mv c-rx" style="font-size:14px;font-weight:500;">Importe uma planilha com a aba <strong>DIGITAL</strong></div></div>`;
    const body = document.getElementById('digital-cards-body');
    if (body) body.innerHTML = '';
    return;
  }
  const ft = document.getElementById('filtro-digital-turma')?.value || '';
  const fd = document.getElementById('filtro-digital-disc')?.value || '';
  const regs = APP.digital.filter(r => (!ft || r.TURMA === ft) && (!fd || r.DISCIPLINA === fd));
  const calcTotal = calcDigital(regs);
  const turmasDig = unique(regs.map(r => r.TURMA)), discsDig = unique(regs.map(r => r.DISCIPLINA));
  const pctGlobal = calcTotal ? calcTotal.pct : 0;
  const pctCls = pctGlobal >= 80 ? 'c-gr' : pctGlobal >= 50 ? 'c-am' : 'c-vm';
  if (mgDig) mgDig.innerHTML = `
    <div class="mc" style="border-left:4px solid var(--roxo);"><div class="ml">Cumprimento geral</div><div class="mv ${pctCls}">${calcTotal ? calcTotal.pct.toFixed(1) + '%' : '—'}</div><div class="ms">${calcTotal ? calcTotal.concluido + ' de ' + calcTotal.previsto + ' aulas' : ''}</div></div>
    <div class="mc" style="border-left:4px solid var(--roxo);"><div class="ml">Turmas</div><div class="mv c-rx">${turmasDig.length}</div><div class="ms">com material digital</div></div>
    <div class="mc" style="border-left:4px solid var(--roxo);"><div class="ml">Disciplinas</div><div class="mv c-rx">${discsDig.length}</div></div>
    <div class="mc" style="border-left:4px solid var(--roxo);"><div class="ml">Registros</div><div class="mv c-rx">${regs.length}</div><div class="ms">turma × disciplina</div></div>`;
  const cardChart = document.getElementById('card-digital-disc-chart');
  if (cardChart) cardChart.style.display = '';
  const pctsPorDisc = discsDig.map(d => { const r = regs.filter(x => x.DISCIPLINA === d); const c = calcDigital(r); return c ? c.pct : 0; });
  destroyChart('ch-digital-disc');
  APP.charts['ch-digital-disc'] = new Chart(document.getElementById('ch-digital-disc'), {
    type: 'bar', data: { labels: discsDig, datasets: [{ label: '% Cumprimento', data: pctsPorDisc, backgroundColor: pctsPorDisc.map(p => p >= 80 ? '#6c2bd9' : p >= 50 ? '#a855f7' : '#c084fc'), borderRadius: 4 }] },
    options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => `${ctx.raw.toFixed(1)}%` } } }, scales: { x: { min: 0, max: 100, ticks: { callback: v => v + '%' } } } }
  });
  const body = document.getElementById('digital-cards-body');
  if (!body) return;
  if (!regs.length) { body.innerHTML = '<p style="color:var(--cinza);font-size:13px;">Nenhum dado para o filtro selecionado.</p>'; return; }
  const grupos = {};
  regs.forEach(r => { if (!grupos[r.TURMA]) grupos[r.TURMA] = {}; if (!grupos[r.TURMA][r.DISCIPLINA]) grupos[r.TURMA][r.DISCIPLINA] = []; grupos[r.TURMA][r.DISCIPLINA].push(r); });
  let html = '';
  Object.keys(grupos).sort().forEach(turma => {
    html += `<div style="margin-bottom:18px;"><div style="font-size:13px;font-weight:700;color:var(--roxo-esc);margin-bottom:10px;padding-bottom:6px;border-bottom:2px solid var(--roxo-cl);">Turma ${turma}</div><div class="digital-grid">`;
    Object.keys(grupos[turma]).sort().forEach(disc => { html += fmtDigitalCard(turma, disc, grupos[turma][disc]); });
    html += `</div></div>`;
  });
  body.innerHTML = html;
}

// ── MODAL ALUNO ────────────────────────────────────────────

function abrirModal(nome, turma) {
  const regs = APP.notas.filter(r => r.ESTUDANTE === nome && r.TURMA === turma);
  if (!regs.length) return;
  document.getElementById('m-av').textContent = iniciais(nome);
  document.getElementById('m-nome').textContent = `${nome} — ${turma}`;
  const todos = regs.map(r => dadosReg(r));
  const temReprov = todos.some(d => d.sit === 'Reprovado'), temAlFreq = todos.some(d => d.alFreq);
  const freqGeral = calcFreqGeral(nome, turma), nivelFG = nivelRiscoFreq(freqGeral);
  const sitEl = document.getElementById('m-sit');
  let sitTxt = 'Aprovado', sitCls = 'b-ap';
  if (temReprov && temAlFreq) { sitTxt = 'Reprovado + ⚠ freq.'; sitCls = 'b-re'; }
  else if (temReprov) { sitTxt = 'Reprovado em disciplina(s)'; sitCls = 'b-re'; }
  else if (nivelFG === 'critico') { sitTxt = 'Aprovado ⚠ Ret.freq.'; sitCls = 'b-re'; }
  else if (temAlFreq) { sitTxt = 'Aprovado ⚠ freq.'; sitCls = 'b-al'; }
  sitEl.className = 'badge ' + sitCls; sitEl.textContent = sitTxt;
  const cores = ['#1a56db', '#057a55', '#b45309', '#c81e1e', '#6c2bd9', '#0e9f6e'];
  let tabela = `<table style="width:100%;border-collapse:collapse;font-size:12px;"><thead><tr style="background:#f9fafb;">
    <th style="padding:8px 10px;text-align:left;border-bottom:1px solid #e5e7eb;font-size:11px;">Disciplina</th>
    <th style="padding:8px 10px;text-align:center;border-bottom:1px solid #e5e7eb;font-size:11px;">B1</th><th style="padding:8px 10px;text-align:center;border-bottom:1px solid #e5e7eb;font-size:11px;">B2</th>
    <th style="padding:8px 10px;text-align:center;border-bottom:1px solid #e5e7eb;font-size:11px;">B3</th><th style="padding:8px 10px;text-align:center;border-bottom:1px solid #e5e7eb;font-size:11px;">B4</th>
    <th style="padding:8px 10px;text-align:center;border-bottom:1px solid #e5e7eb;font-size:11px;">Média</th>
    <th style="padding:8px 10px;text-align:center;border-bottom:1px solid #e5e7eb;font-size:11px;">Frequência</th>
    <th style="padding:8px 10px;text-align:center;border-bottom:1px solid #e5e7eb;font-size:11px;">Situação</th>
  </tr></thead><tbody>`;
  regs.forEach(r => { const d = dadosReg(r); tabela += `<tr><td style="padding:8px 10px;border-bottom:1px solid #f3f4f6;font-weight:600;">${r.DISCIPLINA}</td><td style="padding:8px 10px;border-bottom:1px solid #f3f4f6;text-align:center;">${fmtNota(toNum(r.b1))}</td><td style="padding:8px 10px;border-bottom:1px solid #f3f4f6;text-align:center;">${fmtNota(toNum(r.b2))}</td><td style="padding:8px 10px;border-bottom:1px solid #f3f4f6;text-align:center;">${fmtNota(toNum(r.b3))}</td><td style="padding:8px 10px;border-bottom:1px solid #f3f4f6;text-align:center;">${fmtNota(toNum(r.b4))}</td><td style="padding:8px 10px;border-bottom:1px solid #f3f4f6;text-align:center;">${fmtMedia(d.media, d.parc)}</td><td style="padding:8px 10px;border-bottom:1px solid #f3f4f6;text-align:center;">${fmtFreqCell(d.freq)}</td><td style="padding:8px 10px;border-bottom:1px solid #f3f4f6;text-align:center;">${fmtSit(d.media, d.freq, d.parc)}</td></tr>`; });
  tabela += `</tbody></table>`;
  let htmlFreqGeral = '';
  if (freqGeral) {
    const fCls = freqGeral.pct < CFG.freqMin ? 'var(--verm)' : freqGeral.pct < 80 ? 'var(--amber)' : 'var(--verde)';
    const fTxt = nivelFG === 'critico' ? '⚠ RISCO DE RETENÇÃO — encaminhar para gestão' : nivelFG === 'risco' ? '⚠ Abaixo do mínimo — acompanhar' : 'Regular';
    htmlFreqGeral = `<div style="background:#f9fafb;border:1px solid #e5e7eb;border-radius:8px;padding:10px 14px;margin-bottom:14px;display:flex;align-items:center;gap:12px;"><i class="bi bi-calendar-check" style="color:${fCls};font-size:18px;"></i><div><div style="font-size:11px;font-weight:700;color:var(--cinza-esc);">Frequência geral na turma</div><div style="font-size:18px;font-weight:800;color:${fCls};">${freqGeral.pct.toFixed(1)}%</div><div style="font-size:11px;color:var(--cinza);">${freqGeral.faltas} falta(s) em ${freqGeral.aulas} aulas · ${fTxt}</div></div></div>`;
  }
  const regsDigModal = APP.digital.filter(r => r.TURMA === turma);
  let htmlDig = '';
  if (regsDigModal.length) {
    const discsDig = unique(regsDigModal.map(r => r.DISCIPLINA));
    htmlDig = `<p style="font-size:12px;font-weight:600;color:var(--cinza-esc);margin:14px 0 8px;"><i class="bi bi-laptop" style="color:var(--roxo);margin-right:4px;"></i> Material Digital — Turma ${turma}</p><div class="digital-grid">${discsDig.map(d => fmtDigitalCard(turma, d, regsDigModal.filter(r => r.DISCIPLINA === d))).join('')}</div>`;
  }
  const nomeEsc = nome.replace(/'/g, "\\'");
  document.getElementById('m-body').innerHTML = htmlFreqGeral +
    `<p style="font-size:12px;font-weight:600;color:var(--cinza-esc);margin-bottom:8px;">Evolução bimestral por disciplina</p>` +
    `<div style="position:relative;height:210px;margin-bottom:18px;"><canvas id="ch-modal-evol"></canvas></div>` +
    `<p style="font-size:12px;font-weight:600;color:var(--cinza-esc);margin:0 0 8px;">Notas e frequência por disciplina</p>` +
    tabela + htmlDig +
    `<div style="margin-top:14px;display:flex;gap:8px;justify-content:flex-end;"><button class="btn btn-success btn-sm pdf-trigger" data-nome="${encodeURIComponent(nome)}" data-turma="${encodeURIComponent(turma)}"><i class="bi bi-file-earmark-person"></i> Gerar boletim PDF</button></div>`;
  document.getElementById('modal-aluno').classList.add('open');
  setTimeout(() => {
    const ctx = document.getElementById('ch-modal-evol');
    if (!ctx) return;
    if (APP.charts['ch-modal-evol']) { APP.charts['ch-modal-evol'].destroy(); delete APP.charts['ch-modal-evol']; }
    APP.charts['ch-modal-evol'] = new Chart(ctx, {
      type: 'line', plugins: [pluginLinePct],
      data: { labels: ['1º Bim', '2º Bim', '3º Bim', '4º Bim'], datasets: regs.map((r, i) => ({ label: r.DISCIPLINA, borderColor: cores[i % cores.length], backgroundColor: 'transparent', pointBackgroundColor: cores[i % cores.length], pointRadius: 5, pointHoverRadius: 7, tension: 0.3, spanGaps: true, data: [toNum(r.b1), toNum(r.b2), toNum(r.b3), toNum(r.b4)] })) },
      options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'top', labels: { font: { size: 11 }, boxWidth: 10, padding: 12 } }, tooltip: { callbacks: { label: c => `${c.dataset.label}: ${c.raw !== null ? Number(c.raw).toFixed(1).replace('.', ',') : '—'}` } } }, scales: { y: { min: 0, max: 10, ticks: { stepSize: 1 }, grid: { color: 'rgba(0,0,0,.05)' }, afterDraw(chart) { const ctx2 = chart.ctx, yS = chart.scales.y, xS = chart.scales.x, y5 = yS.getPixelForValue(CFG.mediaAprov); ctx2.save(); ctx2.beginPath(); ctx2.setLineDash([5, 4]); ctx2.strokeStyle = 'rgba(220,38,38,.35)'; ctx2.lineWidth = 1; ctx2.moveTo(xS.left, y5); ctx2.lineTo(xS.right, y5); ctx2.stroke(); ctx2.restore(); } }, x: { grid: { color: 'rgba(0,0,0,.04)' } } } }
    });
  }, 60);
}

function fecharModal(e) { if (e.target === document.getElementById('modal-aluno')) document.getElementById('modal-aluno').classList.remove('open'); }

// ── MEDALHISTAS ────────────────────────────────────────────

function renderMedalhas() {
  const turmaSel = document.getElementById('filtro-turma-medalhas')?.value || '';
  const turmas = turmaSel ? [turmaSel] : unique(APP.notas.map(r => r.TURMA));
  const candidatos = [];
  turmas.forEach(t => {
    unique(APP.notas.filter(r => r.TURMA === t).map(r => r.ESTUDANTE)).forEach(a => {
      const regs = APP.notas.filter(r => r.ESTUDANTE === a && r.TURMA === t);
      const ms = regs.map(r => calcMedia(r)).filter(m => m !== null);
      const media = ms.length ? +(ms.reduce((a, b) => a + b, 0) / ms.length).toFixed(2) : null;
      const fg = calcFreqGeral(a, t);
      if (media !== null) candidatos.push({ nome: a, turma: t, media, freq: fg ? fg.pct : 0 });
    });
  });
  const ouro = candidatos.filter(c => c.media >= 9.0 && c.freq >= 90);
  const prata = candidatos.filter(c => c.media >= 8.0 && c.media < 9.0 && c.freq >= 90);
  const bronze = candidatos.filter(c => c.media >= 7.0 && c.media < 8.0 && c.freq >= 90);
  document.getElementById('mg-medalhas').innerHTML = `
    <div class="mc" style="border-left:4px solid #f59e0b;"><div class="ml">🥇 Ouro</div><div class="mv" style="color:#f59e0b;">${ouro.length}</div><div class="ms">média ≥ 9,0 · freq. ≥ 90%</div></div>
    <div class="mc" style="border-left:4px solid #94a3b8;"><div class="ml">🥈 Prata</div><div class="mv" style="color:#64748b;">${prata.length}</div><div class="ms">média ≥ 8,0 · freq. ≥ 90%</div></div>
    <div class="mc" style="border-left:4px solid #b45309;"><div class="ml">🥉 Bronze</div><div class="mv" style="color:#b45309;">${bronze.length}</div><div class="ms">média ≥ 7,0 · freq. ≥ 90%</div></div>
    <div class="mc bl"><div class="ml">Total premiados</div><div class="mv c-bl">${ouro.length + prata.length + bronze.length}</div><div class="ms">de ${candidatos.length} alunos</div></div>`;
  const sort = arr => [...arr].sort((a, b) => b.media - a.media || b.freq - a.freq);
  const listaHTML = arr => !arr.length ? `<div class="medal-empty">Nenhum aluno nesta categoria</div>` :
    `<ul class="medal-list">${sort(arr).map((c, i) => `<li class="medal-item" onclick="abrirModal('${c.nome.replace(/'/g, "\\'")}','${c.turma}')"><span class="medal-rank">${i + 1}</span><span class="medal-nome">${c.nome} <span class="medal-turma">${c.turma}</span></span><span class="medal-stats"><span class="medal-nota">${c.media.toFixed(1).replace('.', ',')}</span><span class="medal-freq">${c.freq.toFixed(0)}%</span></span></li>`).join('')}</ul>`;
  document.getElementById('medalhas-content').innerHTML = `<div class="medal-grid">
    <div class="medal-card medal-ouro"><div class="medal-header"><span class="medal-emoji">🥇</span><div class="medal-info"><div class="medal-title">Medalha de Ouro</div><div class="medal-sub">Média ≥ 9,0 e frequência ≥ 90%</div></div><span class="medal-count">${ouro.length}</span></div>${listaHTML(ouro)}</div>
    <div class="medal-card medal-prata"><div class="medal-header"><span class="medal-emoji">🥈</span><div class="medal-info"><div class="medal-title">Medalha de Prata</div><div class="medal-sub">Média ≥ 8,0 e frequência ≥ 90%</div></div><span class="medal-count">${prata.length}</span></div>${listaHTML(prata)}</div>
    <div class="medal-card medal-bronze"><div class="medal-header"><span class="medal-emoji">🥉</span><div class="medal-info"><div class="medal-title">Medalha de Bronze</div><div class="medal-sub">Média ≥ 7,0 e frequência ≥ 90%</div></div><span class="medal-count">${bronze.length}</span></div>${listaHTML(bronze)}</div>
  </div>`;
}

// ── RADAR PREDITIVO ────────────────────────────────────────

function renderRadar() {
  const mgRadar = document.getElementById('mg-radar');
  if (!APP.notas.length) { if (mgRadar) mgRadar.innerHTML = `<div class="mc" style="grid-column:1/-1;border-left:4px solid #c2410c;"><div class="ml">Sem dados</div><div class="mv" style="font-size:14px;font-weight:500;color:#c2410c;">Importe dados para ativar o Radar Preditivo</div></div>`; return; }
  const turmaFiltro = document.getElementById('radar-sel-turma')?.value || '';
  const discFiltro = document.getElementById('radar-sel-disc')?.value || '';
  const regs = APP.notas.filter(r => (!turmaFiltro || r.TURMA === turmaFiltro) && (!discFiltro || r.DISCIPLINA === discFiltro));
  if (!regs.length) { if (mgRadar) mgRadar.innerHTML = `<div class="mc" style="grid-column:1/-1;border-left:4px solid var(--cinza);"><div class="ml">Nenhum registro</div><div class="mv c-rx" style="font-size:14px;">Ajuste os filtros</div></div>`; return; }
  const turmas = unique(regs.map(r => r.TURMA)), discs = unique(regs.map(r => r.DISCIPLINA)), alunos = unique(regs.map(r => r.ESTUDANTE));
  let totalRisco = 0, totalColapso = 0, totalDesistencia = 0, totalFreqCrit = 0;
  alunos.forEach(a => {
    const aRegs = regs.filter(r => r.ESTUDANTE === a), aDados = aRegs.map(r => dadosReg(r));
    const reprovCount = aDados.filter(d => d.sit === 'Reprovado').length;
    const alFreqCount = aDados.filter(d => d.alFreq).length;
    const medias = aDados.map(d => d.media).filter(m => m !== null);
    const mediaGeral = medias.length ? +(medias.reduce((a, b) => a + b, 0) / medias.length).toFixed(2) : null;
    const fg = calcFreqGeral(a, aRegs[0]?.TURMA || '');
    if (reprovCount > 0) totalRisco++;
    if (reprovCount >= Math.ceil(aDados.length * 0.5) && alFreqCount > 0) totalColapso++;
    if (mediaGeral !== null && mediaGeral >= CFG.mediaAprov && alFreqCount > 0) totalDesistencia++;
    if (fg && fg.pct < CFG.freqRisco) totalFreqCrit++;
  });
  const pctRisco = alunos.length ? Math.round(totalRisco / alunos.length * 100) : 0;
  if (mgRadar) mgRadar.innerHTML = `
    <div class="mc bl"><div class="ml">Alunos no filtro</div><div class="mv c-bl">${alunos.length}</div><div class="ms">${turmas.length} turma(s) · ${discs.length} disciplina(s)</div></div>
    <div class="mc vm"><div class="ml">Em risco de reprovação</div><div class="mv c-vm">${totalRisco}</div><div class="ms">${pctRisco}% do filtro</div></div>
    <div class="mc am"><div class="ml">Desistência silenciosa</div><div class="mv c-am">${totalDesistencia}</div><div class="ms">nota ok, faltas crescentes</div></div>
    <div class="mc vm"><div class="ml">Freq. crítica (&lt;${CFG.freqRisco}%)</div><div class="mv c-vm">${totalFreqCrit}</div><div class="ms">encaminhar para gestão</div></div>`;

  // Triagem de frequência
  const bodyFreq = document.getElementById('radar-freq-body');
  if (bodyFreq) {
    const criticos = [], riscos = [], atencao = [];
    alunos.forEach(a => {
      const turmaAluno = regs.find(r => r.ESTUDANTE === a)?.TURMA || '';
      const fg = calcFreqGeral(a, turmaAluno);
      if (!fg) return;
      const nivel = nivelRiscoFreq(fg);
      const faltasFiltradas = APP.faltas.filter(r => r.ESTUDANTE === a && r.TURMA === turmaAluno);
      const piora = faltasFiltradas.some(f => tendenciaPioraFreq(calcFreq(f)));
      const obj = { nome: a, turma: turmaAluno, pct: fg.pct, faltas: fg.faltas, aulas: fg.aulas, piora };
      if (nivel === 'critico') criticos.push(obj);
      else if (nivel === 'risco') riscos.push(obj);
      else if (piora) atencao.push(obj);
    });
    const sortFreq = arr => [...arr].sort((a, b) => a.pct - b.pct);
    const freqItem = (item, cls, titulo) => `<div style="display:flex;align-items:center;gap:10px;padding:7px 12px;background:rgba(255,255,255,.7);border-radius:8px;border:1px solid rgba(0,0,0,.06);margin-bottom:5px;"><span class="freq-ruim" style="font-size:15px;font-weight:800;min-width:50px;">${item.pct.toFixed(1)}%</span><div style="flex:1;min-width:0;"><div style="font-size:12px;font-weight:700;color:#111827;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${item.nome}</div><div style="font-size:11px;color:var(--cinza);">${item.turma} · ${item.faltas} falta(s) em ${item.aulas} aulas${item.piora ? ' · <span style="color:var(--verm);">↗ piora crescente</span>' : ''}</div></div><span class="badge ${cls}" style="font-size:10px;">${titulo}</span></div>`;
    let htmlFreq = '';
    if (criticos.length) htmlFreq += `<div class="radar-alerta radar-alerta-risco" style="margin-bottom:10px;"><div class="radar-alerta-icon">🚨</div><div class="radar-alerta-body"><div class="radar-alerta-titulo">Risco de Retenção por Frequência — ${criticos.length} aluno(s)</div><div class="radar-alerta-desc">Frequência abaixo de ${CFG.freqRisco}% — <strong>encaminhar imediatamente para a gestão</strong>.</div><div style="margin-top:8px;">${sortFreq(criticos).map(i => freqItem(i, 'b-re', 'Crítico')).join('')}</div></div></div>`;
    if (riscos.length) htmlFreq += `<div class="radar-alerta radar-alerta-disc" style="margin-bottom:10px;"><div class="radar-alerta-icon">⚠️</div><div class="radar-alerta-body"><div class="radar-alerta-titulo">Abaixo do Mínimo (75%) — ${riscos.length} aluno(s)</div><div class="radar-alerta-desc">Frequência insuficiente. Encaminhar para gestão.</div><div style="margin-top:8px;">${sortFreq(riscos).map(i => freqItem(i, 'b-al', 'Risco')).join('')}</div></div></div>`;
    if (atencao.length) htmlFreq += `<div class="radar-alerta radar-alerta-info" style="margin-bottom:10px;"><div class="radar-alerta-icon">📉</div><div class="radar-alerta-body"><div class="radar-alerta-titulo">Tendência de Piora na Frequência — ${atencao.length} aluno(s)</div><div class="radar-alerta-desc">Frequência ainda acima de 75%, mas com crescimento nas faltas.</div><div style="margin-top:8px;">${sortFreq(atencao).map(i => freqItem(i, 'b-info', 'Atenção')).join('')}</div></div></div>`;
    if (!htmlFreq) htmlFreq = `<div class="radar-alerta radar-alerta-ok"><div class="radar-alerta-icon">✅</div><div class="radar-alerta-body"><div class="radar-alerta-titulo">Nenhum aluno com frequência em risco no filtro atual</div></div></div>`;
    bodyFreq.innerHTML = htmlFreq;
  }

  // Discrepância Digital vs Notas
  const bodyDisc = document.getElementById('radar-discrepancia-body');
  if (bodyDisc) {
    if (!APP.digital.length) { bodyDisc.innerHTML = `<p style="color:var(--cinza);font-size:13px;"><i class="bi bi-info-circle"></i> Importe a aba <strong>DIGITAL</strong> para ativar esta análise.</p>`; }
    else {
      let htmlDisc = '';
      turmas.forEach(t => discs.forEach(d => {
        if (!regs.some(r => r.TURMA === t && r.DISCIPLINA === d)) return;
        const digRegs = APP.digital.filter(r => r.TURMA === t && r.DISCIPLINA === d);
        if (!digRegs.length) return;
        const calc = calcDigital(digRegs);
        if (!calc) return;
        const notasCombo = regs.filter(r => r.TURMA === t && r.DISCIPLINA === d);
        const mediasCombo = notasCombo.map(r => calcMedia(r)).filter(m => m !== null);
        if (!mediasCombo.length) return;
        const mediaMedia = +(mediasCombo.reduce((a, b) => a + b, 0) / mediasCombo.length).toFixed(2);
        if (calc.pct > 70 && mediaMedia < CFG.mediaAprov) {
          const alunosBaixoRend = notasCombo.filter(r => { const m = calcMedia(r); return m !== null && m < CFG.mediaAprov; }).map(r => r.ESTUDANTE);
          htmlDisc += `<div class="radar-alerta radar-alerta-disc"><div class="radar-alerta-icon">⚠️</div><div class="radar-alerta-body"><div class="radar-alerta-titulo">Discrepância em ${d} — Turma ${t}</div><div class="radar-alerta-desc">Alto volume de material digital (<strong>${calc.pct.toFixed(0)}%</strong>) com baixo rendimento (média <strong>${mediaMedia.toFixed(1).replace('.', ',')}</strong>). Esta discrepância <em>não implica falha do professor</em>.</div><div class="radar-alerta-alunos"><span class="radar-alunos-label">Alunos com baixo rendimento:</span><div class="radar-alunos-lista">${alunosBaixoRend.map(n => `<span class="radar-aluno-tag radar-tag-amber">${n}</span>`).join('')}</div></div></div></div>`;
        }
        if (calc.pct < 30 && mediaMedia >= CFG.mediaAprov) {
          htmlDisc += `<div class="radar-alerta radar-alerta-info"><div class="radar-alerta-icon">ℹ️</div><div class="radar-alerta-body"><div class="radar-alerta-titulo">Ponto de curiosidade em ${d} — Turma ${t}</div><div class="radar-alerta-desc">Baixo uso de material digital (<strong>${calc.pct.toFixed(0)}%</strong>) com bom rendimento (média <strong>${mediaMedia.toFixed(1).replace('.', ',')}</strong>). Verificar metodologias complementares.</div></div></div>`;
        }
      }));
      bodyDisc.innerHTML = htmlDisc || `<div class="radar-alerta radar-alerta-ok"><div class="radar-alerta-icon">✅</div><div class="radar-alerta-body"><div class="radar-alerta-titulo">Nenhuma discrepância digital detectada</div></div></div>`;
    }
  }

  // Volume de alunos em risco por turma/disciplina
  const bodyRisco = document.getElementById('radar-risco-turma-body');
  if (bodyRisco) {
    let htmlRisco = '';
    turmas.forEach(t => discs.forEach(d => {
      const notasCombo = regs.filter(r => r.TURMA === t && r.DISCIPLINA === d);
      if (!notasCombo.length) return;
      const dados = notasCombo.map(r => dadosReg(r));
      const emRisco = dados.filter(d2 => d2.sit === 'Reprovado');
      const pct = Math.round(emRisco.length / dados.length * 100);
      if (pct > 30) htmlRisco += `<div class="radar-alerta radar-alerta-risco"><div class="radar-alerta-icon">📊</div><div class="radar-alerta-body"><div class="radar-alerta-titulo">Ponto de Atenção em ${d} — Turma ${t}</div><div class="radar-alerta-desc"><strong>${emRisco.length} alunos</strong> (${pct}% da turma nesta disciplina) em risco. Alinhar plano de recuperação com o professor.</div><div class="radar-alerta-alunos"><span class="radar-alunos-label">Estudantes em risco:</span><div class="radar-alunos-lista">${emRisco.map(d2 => `<span class="radar-aluno-tag radar-tag-verm">${d2.n.ESTUDANTE}</span>`).join('')}</div></div></div></div>`;
    }));
    bodyRisco.innerHTML = htmlRisco || `<div class="radar-alerta radar-alerta-ok"><div class="radar-alerta-icon">✅</div><div class="radar-alerta-body"><div class="radar-alerta-titulo">Nenhuma turma/disciplina com mais de 30% em risco</div></div></div>`;
  }

  // Perfis de risco individual
  const desistentes = [], localizados = [], colapsos = [];
  alunos.forEach(a => {
    const aRegs = regs.filter(r => r.ESTUDANTE === a), aDados = aRegs.map(r => dadosReg(r));
    const turmaAluno = aRegs[0]?.TURMA || '';
    const reprovCount = aDados.filter(d => d.sit === 'Reprovado').length;
    const alFreqCount = aDados.filter(d => d.alFreq).length;
    const medias = aDados.map(d => d.media).filter(m => m !== null);
    const mediaGeral = medias.length ? +(medias.reduce((x, y) => x + y, 0) / medias.length).toFixed(2) : null;
    if (reprovCount >= Math.ceil(aDados.length * 0.5) && alFreqCount > 0) colapsos.push({ nome: a, turma: turmaAluno, reprov: reprovCount, totalDiscs: aDados.length, alFreq: alFreqCount });
    else if (mediaGeral !== null && mediaGeral >= CFG.mediaAprov && alFreqCount > 0) desistentes.push({ nome: a, turma: turmaAluno, media: mediaGeral, alFreq: alFreqCount });
    else if (reprovCount === 1 || reprovCount === 2) { const discsReprov = aRegs.filter(r => dadosReg(r).sit === 'Reprovado').map(r => r.DISCIPLINA); localizados.push({ nome: a, turma: turmaAluno, discs: discsReprov }); }
  });
  const perfilItem = (item, tipo) => {
    if (tipo === 'desistencia') return `<div class="radar-perfil-item"><strong>${item.nome}</strong><span class="radar-perfil-detalhe">${item.turma} · Média ${item.media.toFixed(1).replace('.', ',')} · ${item.alFreq} disc. com alerta freq.</span></div>`;
    if (tipo === 'localizada') return `<div class="radar-perfil-item"><strong>${item.nome}</strong><span class="radar-perfil-detalhe">${item.turma} · Reprovando em: ${item.discs.join(', ')}</span></div>`;
    if (tipo === 'colapso') return `<div class="radar-perfil-item"><strong>${item.nome}</strong><span class="radar-perfil-detalhe">${item.turma} · ${item.reprov}/${item.totalDiscs} disc. reprovadas · ${item.alFreq} com alerta freq.</span></div>`;
    return '';
  };
  const listaDesEl = document.getElementById('radar-lista-desistencia');
  const listaLocEl = document.getElementById('radar-lista-localizada');
  const listaColEl = document.getElementById('radar-lista-colapso');
  if (listaDesEl) listaDesEl.innerHTML = desistentes.length ? desistentes.map(i => perfilItem(i, 'desistencia')).join('') : `<div class="radar-perfil-vazio">Nenhum aluno neste perfil.</div>`;
  if (listaLocEl) listaLocEl.innerHTML = localizados.length ? localizados.map(i => perfilItem(i, 'localizada')).join('') : `<div class="radar-perfil-vazio">Nenhum aluno neste perfil.</div>`;
  if (listaColEl) listaColEl.innerHTML = colapsos.length ? colapsos.map(i => perfilItem(i, 'colapso')).join('') : `<div class="radar-perfil-vazio">Nenhum aluno neste perfil.</div>`;

  // Viabilidade de recuperação
  const bodyViab = document.getElementById('radar-viabilidade-body');
  if (!bodyViab) return;
  const viabRows = [];
  regs.forEach(r => {
    const { media, sit, parc } = dadosReg(r);
    if (media === null) return;
    if (sit !== 'Reprovado' && !parc) return;
    const bims = [toNum(r.b1), toNum(r.b2), toNum(r.b3), toNum(r.b4)];
    const lancados = bims.filter(v => v !== null).length, restantes = 4 - lancados;
    if (restantes === 0) { if (sit === 'Reprovado') viabRows.push({ nome: r.ESTUDANTE, turma: r.TURMA, disc: r.DISCIPLINA, mediaAtual: media, notaNecessaria: null, impossivel: true }); return; }
    const somaAtual = bims.filter(v => v !== null).reduce((a, b) => a + b, 0);
    const notaPorBim = (CFG.mediaAprov * 4 - somaAtual) / restantes;
    if (notaPorBim > 10) viabRows.push({ nome: r.ESTUDANTE, turma: r.TURMA, disc: r.DISCIPLINA, mediaAtual: media, notaNecessaria: notaPorBim, restantes, impossivel: true });
    else if (notaPorBim > 8.5) viabRows.push({ nome: r.ESTUDANTE, turma: r.TURMA, disc: r.DISCIPLINA, mediaAtual: media, notaNecessaria: notaPorBim, restantes, impossivel: false, dificil: true });
    else if (notaPorBim > 0) viabRows.push({ nome: r.ESTUDANTE, turma: r.TURMA, disc: r.DISCIPLINA, mediaAtual: media, notaNecessaria: notaPorBim, restantes, impossivel: false, dificil: false });
  });
  if (!viabRows.length) { bodyViab.innerHTML = `<div class="radar-alerta radar-alerta-ok"><div class="radar-alerta-icon">✅</div><div class="radar-alerta-body"><div class="radar-alerta-titulo">Nenhum aluno em situação de risco de recuperação</div></div></div>`; return; }
  const impossiveis = viabRows.filter(v => v.impossivel), dificeis = viabRows.filter(v => !v.impossivel && v.dificil), possiveis = viabRows.filter(v => !v.impossivel && !v.dificil);
  const viabRow = v => {
    const notaStr = v.notaNecessaria === null ? 'Reprovado definitivo (sem bimestres restantes)' :
      v.notaNecessaria > 10 ? `${v.notaNecessaria.toFixed(1).replace('.', ',')} (impossível)` :
        `${v.notaNecessaria.toFixed(1).replace('.', ',')} em cada bimestre restante (${v.restantes} restante${v.restantes > 1 ? 's' : ''})`;
    return `<tr><td><strong>${v.nome}</strong></td><td><span class="badge b-info">${v.turma}</span></td><td style="font-size:12px;">${v.disc}</td><td>${fmtMedia(v.mediaAtual, true)}</td>
      <td style="font-size:12px;${v.impossivel ? 'color:var(--verm);font-weight:600;' : v.dificil ? 'color:var(--amber);font-weight:600;' : 'color:var(--verde);'}">${notaStr}</td>
      <td>${v.impossivel ? `<span class="badge b-re"><i class="bi bi-x-circle-fill"></i> Intervenção externa</span>` : v.dificil ? `<span class="badge b-al"><i class="bi bi-exclamation-triangle-fill"></i> Recuperação difícil</span>` : `<span class="badge b-ap"><i class="bi bi-check-circle-fill"></i> Recuperação viável</span>`}</td></tr>`;
  };
  bodyViab.innerHTML = `<div class="ib" style="margin-bottom:14px;"><i class="bi bi-info-circle-fill"></i><span>Cálculo baseado na nota mínima necessária para atingir média ${CFG.mediaAprov.toFixed(1).replace('.', ',')} nos bimestres restantes.</span></div>
    ${impossiveis.length ? `<div style="font-size:12px;font-weight:700;color:var(--verm);margin-bottom:6px;"><i class="bi bi-x-circle-fill"></i> Recuperação Improvável — Necessita Intervenção Externa (${impossiveis.length})</div>` : ''}
    <div class="tw"><table><thead><tr><th>Aluno</th><th>Turma</th><th>Disciplina</th><th>Média atual</th><th>Nota necessária</th><th>Viabilidade</th></tr></thead><tbody>${[...impossiveis, ...dificeis, ...possiveis].map(viabRow).join('')}</tbody></table></div>`;
}

// ── RELATÓRIOS PDF ─────────────────────────────────────────

function popSelects() {
  const turmas = unique(APP.notas.map(r => r.TURMA)), discs = unique(APP.notas.map(r => r.DISCIPLINA));
  ['sel-turma-pdf', 'sel-turma-aluno'].forEach(id => { const el = document.getElementById(id); if (el) el.innerHTML = turmas.map(t => `<option value="${t}">${t}</option>`).join(''); });
  const sdp = document.getElementById('sel-disc-pdf'); if (sdp) sdp.innerHTML = discs.map(d => `<option value="${d}">${d}</option>`).join('');
  const stdp = document.getElementById('sel-turma-disc-pdf'); if (stdp) stdp.innerHTML = '<option value="">Todas</option>' + turmas.map(t => `<option value="${t}">${t}</option>`).join('');
  popAluno();
}

function popAluno() {
  const t = document.getElementById('sel-turma-aluno')?.value;
  const al = unique(APP.notas.filter(r => r.TURMA === t).map(r => r.ESTUDANTE));
  const el = document.getElementById('sel-aluno-pdf');
  if (el) el.innerHTML = al.map(a => `<option value="${a}">${a}</option>`).join('');
}

function estilosPDF() {
  return `body{font-family:Arial,sans-serif;padding:22px;color:#111;}h1{font-size:17px;margin:0 0 2px;}h2{font-size:12px;color:#6b7280;margin:0 0 14px;}.stats{display:flex;gap:10px;margin-bottom:18px;flex-wrap:wrap;}.stat{padding:9px 16px;border-radius:7px;text-align:center;font-size:12px;min-width:80px;}table{width:100%;border-collapse:collapse;}th{background:#f9fafb;padding:6px 9px;text-align:left;font-size:11px;color:#374151;border-bottom:1px solid #e5e7eb;white-space:nowrap;}td{padding:6px 9px;border-bottom:1px solid #f3f4f6;font-size:12px;}.ok{color:#057a55;font-weight:700;}.re{color:#c81e1e;font-weight:700;}.am{color:#b45309;font-weight:700;}.ne{color:#9ca3af;}.b-ap{background:#def7ec;color:#014737;padding:1px 7px;border-radius:99px;font-size:10px;font-weight:700;}.b-re{background:#fde8e8;color:#771d1d;padding:1px 7px;border-radius:99px;font-size:10px;font-weight:700;}.b-al{background:#fef3c7;color:#78350f;padding:1px 7px;border-radius:99px;font-size:10px;font-weight:700;}@media print{body{padding:0;}}`;
}

function fmtNotaPDF(v) { if (v === null || v === undefined) return '<span class="ne">—</span>'; return `<span class="${v >= CFG.mediaAprov ? 'ok' : 're'}">${v.toFixed(1).replace('.', ',')}</span>`; }
function fmtSitPDF(media, alFreq, parc) { const p = parc ? ' (parc.)' : ''; if (media === null) return '—'; if (media >= CFG.mediaAprov) return alFreq ? `<span class="b-al">Aprov. ⚠freq.${p}</span>` : `<span class="b-ap">Aprovado${p}</span>`; return `<span class="b-re">Reprovado${p}</span>`; }

function pdfTurma() {
  const t = document.getElementById('sel-turma-pdf')?.value;
  if (!t) return alert('Selecione uma turma.');
  const regs = APP.notas.filter(r => r.TURMA === t), discs = unique(regs.map(r => r.DISCIPLINA)), alunos = unique(regs.map(r => r.ESTUDANTE));
  let ap = 0, rp = 0, alF = 0;
  alunos.forEach(a => { const ds = regs.filter(r => r.ESTUDANTE === a).map(r => dadosReg(r)); if (ds.some(d => d.sit === 'Reprovado')) rp++; else ap++; if (ds.some(d => d.alFreq)) alF++; });
  const rows = alunos.map(a => discs.map(d => { const r = regs.find(x => x.ESTUDANTE === a && x.DISCIPLINA === d); if (!r) return null; const dd = dadosReg(r); return `<tr><td style="font-weight:600;">${a}</td><td>${d}</td>${['b1', 'b2', 'b3', 'b4'].map(b => `<td style="text-align:center;">${fmtNotaPDF(toNum(r[b]))}</td>`).join('')}<td style="text-align:center;">${fmtNotaPDF(dd.media)}</td><td style="text-align:center;">${dd.freq ? `<span class="${dd.freq.pct < CFG.freqMin ? 'am' : 'ok'}">${dd.freq.pct.toFixed(1)}%</span>` : '<span class="ne">—</span>'}</td><td style="text-align:center;">${fmtSitPDF(dd.media, dd.alFreq, dd.parc)}</td></tr>`; }).filter(Boolean).join('')).join('');
  const w = window.open('', '_blank'); if (!w) return;
  w.document.write(`<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><title>Turma ${t}</title><style>${estilosPDF()}</style></head><body><h1>${CFG.escola} — Turma ${t}</h1><h2>${CFG.etapa} · Conselho de Classe · ${new Date().toLocaleDateString('pt-BR')}</h2><div class="stats"><div class="stat" style="background:#def7ec;color:#014737;"><strong>${ap}</strong><br>Aprovados</div><div class="stat" style="background:#fde8e8;color:#771d1d;"><strong>${rp}</strong><br>Reprovados</div><div class="stat" style="background:#fef3c7;color:#78350f;"><strong>${alF}</strong><br>Alerta freq.</div><div class="stat" style="background:#f3f4f6;color:#374151;"><strong>${alunos.length}</strong><br>Alunos</div></div><table><thead><tr><th>Aluno</th><th>Disciplina</th><th style="text-align:center;">B1</th><th style="text-align:center;">B2</th><th style="text-align:center;">B3</th><th style="text-align:center;">B4</th><th style="text-align:center;">Média</th><th style="text-align:center;">Frequência</th><th>Situação</th></tr></thead><tbody>${rows}</tbody></table><script>window.onload=()=>window.print();<\/script></body></html>`);
  w.document.close();
}

function pdfAluno() { const nome = document.getElementById('sel-aluno-pdf')?.value, turma = document.getElementById('sel-turma-aluno')?.value; if (!nome) return alert('Selecione um aluno.'); pdfAlunoNome(nome, turma); }

function pdfAlunoNome(nome, turma) {
  const regs = APP.notas.filter(r => r.ESTUDANTE === nome && r.TURMA === turma); if (!regs.length) return;
  const todos = regs.map(r => dadosReg(r)), temRep = todos.some(d => d.sit === 'Reprovado'), temAlF = todos.some(d => d.alFreq);
  const freqGeral = calcFreqGeral(nome, turma), sitGeral = temRep ? 'Reprovado em disciplina(s)' : temAlF ? 'Aprovado ⚠ frequência' : 'Aprovado';
  const rows = regs.map(r => { const d = dadosReg(r); return `<tr><td style="font-weight:600;">${r.DISCIPLINA}</td>${['b1', 'b2', 'b3', 'b4'].map(b => `<td style="text-align:center;">${fmtNotaPDF(toNum(r[b]))}</td>`).join('')}<td style="text-align:center;">${fmtNotaPDF(d.media)}</td><td style="text-align:center;">${d.freq ? `<span class="${d.freq.pct < CFG.freqMin ? 'am' : 'ok'}">${d.freq.pct.toFixed(1)}% (${d.freq.faltas} falta${d.freq.faltas !== 1 ? 's' : ''}/${d.freq.aulas} aulas)</span>` : '<span class="ne">—</span>'}</td><td style="text-align:center;">${fmtSitPDF(d.media, d.alFreq, d.parc)}</td><td style="font-size:11px;color:#6b7280;">${d.media !== null && d.media < CFG.mediaAprov ? 'Solicitar comp. nota' : d.alFreq ? 'Verificar comp. freq.' : ''}</td></tr>`; }).join('');
  const freqLine = freqGeral ? `<div class="stat" style="background:${freqGeral.pct < CFG.freqMin ? '#fde8e8;color:#771d1d' : '#def7ec;color:#014737'};"><strong>${freqGeral.pct.toFixed(1)}%</strong><br>Freq. geral</div>` : '';
  const w = window.open('', '_blank'); if (!w) return;
  w.document.write(`<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><title>Boletim ${nome}</title><style>${estilosPDF()}.header{background:#1a56db;color:#fff;padding:14px 18px;border-radius:8px;margin-bottom:18px;}.header h1{color:#fff;font-size:16px;}.header p{font-size:11px;opacity:.85;margin:0;}</style></head><body><div class="header"><h1>${nome}</h1><p>Turma ${turma} · ${CFG.escola} · ${new Date().toLocaleDateString('pt-BR')}</p></div><div class="stats"><div class="stat" style="background:${temRep ? '#fde8e8;color:#771d1d' : temAlF ? '#fef3c7;color:#78350f' : '#def7ec;color:#014737'};"><strong>${sitGeral}</strong><br>Situação</div>${freqLine}</div><table><thead><tr><th>Disciplina</th><th style="text-align:center;">B1</th><th style="text-align:center;">B2</th><th style="text-align:center;">B3</th><th style="text-align:center;">B4</th><th style="text-align:center;">Média</th><th>Frequência</th><th>Situação</th><th>Observação</th></tr></thead><tbody>${rows}</tbody></table><script>window.onload=()=>window.print();<\/script></body></html>`);
  w.document.close();
}

function pdfDisc() {
  const disc = document.getElementById('sel-disc-pdf')?.value, turma = document.getElementById('sel-turma-disc-pdf')?.value;
  if (!disc) return alert('Selecione uma disciplina.');
  const regs = APP.notas.filter(r => r.DISCIPLINA === disc && (!turma || r.TURMA === turma)); if (!regs.length) return alert('Nenhum dado encontrado.');
  let ap = 0, rp = 0, alF = 0; regs.forEach(r => { const d = dadosReg(r); if (d.sit === 'Reprovado') rp++; else if (d.sit === 'Aprovado') ap++; if (d.alFreq) alF++; });
  const rows = regs.map(r => { const d = dadosReg(r), obs = d.media !== null && d.media < CFG.mediaAprov ? 'Solicitar comp. de nota' : d.alFreq ? 'Verificar comp. de frequência' : ''; return `<tr><td style="font-weight:600;">${r.ESTUDANTE}</td><td style="text-align:center;"><span class="b-ap" style="background:#e8f0fe;color:#1e3a8a;">${r.TURMA}</span></td>${['b1', 'b2', 'b3', 'b4'].map(b => `<td style="text-align:center;">${fmtNotaPDF(toNum(r[b]))}</td>`).join('')}<td style="text-align:center;">${fmtNotaPDF(d.media)}</td><td style="text-align:center;">${d.freq ? `<span class="${d.freq.pct < CFG.freqMin ? 'am' : 'ok'}">${d.freq.pct.toFixed(1)}%</span>` : '<span class="ne">—</span>'}</td><td style="text-align:center;">${fmtSitPDF(d.media, d.alFreq, d.parc)}</td><td style="font-size:11px;color:#b45309;font-weight:600;">${obs}</td></tr>`; }).join('');
  const w = window.open('', '_blank'); if (!w) return;
  w.document.write(`<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><title>${disc}</title><style>${estilosPDF()}</style></head><body><h1>${disc}${turma ? ' — Turma ' + turma : ' — Todas as turmas'}</h1><h2>${CFG.escola} · ${CFG.etapa} · Conselho de Classe · ${new Date().toLocaleDateString('pt-BR')}</h2><div class="stats"><div class="stat" style="background:#def7ec;color:#014737;"><strong>${ap}</strong><br>Aprovados</div><div class="stat" style="background:#fde8e8;color:#771d1d;"><strong>${rp}</strong><br>Reprovados</div><div class="stat" style="background:#fef3c7;color:#78350f;"><strong>${alF}</strong><br>Alerta freq.</div><div class="stat" style="background:#f3f4f6;"><strong>${regs.length}</strong><br>Total</div></div><table><thead><tr><th>Aluno</th><th>Turma</th><th style="text-align:center;">B1</th><th style="text-align:center;">B2</th><th style="text-align:center;">B3</th><th style="text-align:center;">B4</th><th style="text-align:center;">Média</th><th style="text-align:center;">Freq.</th><th>Situação</th><th>Observação</th></tr></thead><tbody>${rows}</tbody></table><script>window.onload=()=>window.print();<\/script></body></html>`);
  w.document.close();
}

// ── MÓDULO PPTX ────────────────────────────────────────────

const PPTX_CORES = { roxo: '6C2BD9', roxoEsc: '4A1D96', roxoCl: 'F3E8FF', azul: '1A56DB', azulEsc: '1E3A8A', azulCl: 'EFF6FF', verde: '057A55', verdeCl: 'DEF7EC', verdeEsc: '014737', verm: 'C81E1E', vermCl: 'FDE8E8', amber: 'B45309', amberCl: 'FEF3C7', cinza: '6B7280', cinzaCl: 'F3F4F6', cinzaEsc: '374151', branco: 'FFFFFF', escuro: '0F172A' };

function pptxAddHeader(slide, prs, cor, texto, h = 0.65) {
  slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: '100%', h, fill: { color: cor } });
  slide.addText(texto, { x: 0.3, y: 0, w: 9.4, h, fontSize: h > 0.6 ? 16 : 13, bold: true, color: PPTX_CORES.branco, fontFace: 'Calibri', valign: 'middle' });
}

function pptxCapa(prs, titulo, subtitulo, detalhe) {
  const slide = prs.addSlide();
  slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: '100%', h: '100%', fill: { color: PPTX_CORES.escuro } });
  slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.08, fill: { color: PPTX_CORES.roxo } });
  slide.addShape(prs.ShapeType.rect, { x: 0, y: 3.42, w: '100%', h: 0.08, fill: { color: PPTX_CORES.roxo } });
  slide.addShape(prs.ShapeType.ellipse, { x: 0.45, y: 0.7, w: 0.9, h: 0.9, fill: { color: PPTX_CORES.roxo }, line: { color: PPTX_CORES.roxo } });
  slide.addText('📊', { x: 0.45, y: 0.7, w: 0.9, h: 0.9, align: 'center', valign: 'middle', fontSize: 22 });
  slide.addText(titulo, { x: 0.45, y: 1.75, w: 9.1, h: 0.85, fontSize: 32, bold: true, color: PPTX_CORES.branco, fontFace: 'Calibri' });
  slide.addText(subtitulo, { x: 0.45, y: 2.55, w: 9.1, h: 0.5, fontSize: 16, color: 'A5B4FC', fontFace: 'Calibri' });
  slide.addText(detalhe, { x: 0.45, y: 3.12, w: 9.1, h: 0.35, fontSize: 11, color: '94A3B8', fontFace: 'Calibri' });
  return slide;
}

function pptxSlideResumoEscola(prs, dados) {
  const slide = prs.addSlide();
  const { turmas, totalAlunos, aprov, reprov, alFreqs, mg } = dados;
  pptxAddHeader(slide, prs, PPTX_CORES.roxo, 'Painel Consolidado — Escola');
  const cards = [
    { label: 'Total de Alunos', valor: String(totalAlunos), cor: PPTX_CORES.azulCl, corTexto: PPTX_CORES.azulEsc, borda: PPTX_CORES.azul },
    { label: 'Aprovados', valor: String(aprov), cor: PPTX_CORES.verdeCl, corTexto: PPTX_CORES.verdeEsc, borda: PPTX_CORES.verde },
    { label: 'Reprovados', valor: String(reprov), cor: PPTX_CORES.vermCl, corTexto: PPTX_CORES.verm, borda: PPTX_CORES.verm },
    { label: 'Alerta Freq.', valor: String(alFreqs), cor: PPTX_CORES.amberCl, corTexto: PPTX_CORES.amber, borda: PPTX_CORES.amber },
    { label: 'Média Geral', valor: mg.toFixed(1).replace('.', ','), cor: PPTX_CORES.roxoCl, corTexto: PPTX_CORES.roxoEsc, borda: PPTX_CORES.roxo },
  ];
  const cW = 1.82, cH = 1.0, cY = 0.82, cGap = 0.065;
  cards.forEach((c, i) => { const cX = 0.25 + i * (cW + cGap); slide.addShape(prs.ShapeType.rect, { x: cX, y: cY, w: cW, h: cH, fill: { color: c.cor }, line: { color: c.borda, pt: 1.5 }, rectRadius: 0.08 }); slide.addText(c.valor, { x: cX, y: cY + 0.12, w: cW, h: 0.52, fontSize: 26, bold: true, color: c.corTexto, align: 'center', fontFace: 'Calibri' }); slide.addText(c.label, { x: cX, y: cY + 0.62, w: cW, h: 0.3, fontSize: 10, color: PPTX_CORES.cinzaEsc, align: 'center', fontFace: 'Calibri' }); });
  slide.addText('Resumo por Turma', { x: 0.3, y: 2.05, w: 9.4, h: 0.32, fontSize: 12, bold: true, color: PPTX_CORES.cinzaEsc, fontFace: 'Calibri' });
  const hdr = ['Turma', 'Alunos', 'Aprovados', 'Reprovados', 'Alerta Freq.', 'Média'].map(t => ({ text: t, options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.roxo, align: 'center' } }));
  const tabRows = [hdr, ...turmas.map((info, idx) => {
    const bg = idx % 2 === 0 ? PPTX_CORES.branco : PPTX_CORES.cinzaCl, medCor = info.mg >= CFG.mediaAprov ? PPTX_CORES.verde : PPTX_CORES.verm;
    return [{ text: info.turma, options: { bold: true, fill: bg, align: 'center' } }, { text: String(info.alunos), options: { fill: bg, align: 'center' } }, { text: String(info.aprov), options: { fill: bg, align: 'center', color: PPTX_CORES.verde, bold: true } }, { text: String(info.reprov), options: { fill: bg, align: 'center', color: info.reprov > 0 ? PPTX_CORES.verm : PPTX_CORES.cinzaEsc, bold: info.reprov > 0 } }, { text: String(info.alFreq), options: { fill: bg, align: 'center', color: info.alFreq > 0 ? PPTX_CORES.amber : PPTX_CORES.cinzaEsc } }, { text: info.mg.toFixed(1).replace('.', ','), options: { fill: bg, align: 'center', color: medCor, bold: true } }];
  })];
  slide.addTable(tabRows, { x: 0.3, y: 2.38, w: 9.4, h: Math.min(1.4, 0.28 * (turmas.length + 1)), fontSize: 10, fontFace: 'Calibri', color: PPTX_CORES.cinzaEsc, border: { pt: 0.5, color: 'E5E7EB' }, rowH: 0.28, colW: [1.2, 1.0, 1.4, 1.4, 1.4, 1.0] });
  return slide;
}

function pptxSlideResumoPorTurma(prs, turma) {
  const slide = prs.addSlide(), regs = APP.notas.filter(r => r.TURMA === turma), alunos = unique(regs.map(r => r.ESTUDANTE));
  let ap = 0, rp = 0, alF = 0; alunos.forEach(a => { const ds = regs.filter(r => r.ESTUDANTE === a).map(r => dadosReg(r)); if (ds.some(d => d.sit === 'Reprovado')) rp++; else ap++; if (ds.some(d => d.alFreq)) alF++; });
  const ms = regs.map(r => calcMedia(r)).filter(m => m !== null), mg = ms.length ? +(ms.reduce((a, b) => a + b, 0) / ms.length).toFixed(1) : 0, discs = unique(regs.map(r => r.DISCIPLINA));
  pptxAddHeader(slide, prs, PPTX_CORES.azul, `Turma ${turma} — Resumo`);
  const cards = [{ label: 'Alunos', valor: String(alunos.length), cor: PPTX_CORES.azulCl, borda: PPTX_CORES.azul, corTexto: PPTX_CORES.azulEsc }, { label: 'Aprovados', valor: String(ap), cor: PPTX_CORES.verdeCl, borda: PPTX_CORES.verde, corTexto: PPTX_CORES.verdeEsc }, { label: 'Reprovados', valor: String(rp), cor: PPTX_CORES.vermCl, borda: PPTX_CORES.verm, corTexto: PPTX_CORES.verm }, { label: 'Alerta Freq.', valor: String(alF), cor: PPTX_CORES.amberCl, borda: PPTX_CORES.amber, corTexto: PPTX_CORES.amber }, { label: 'Média', valor: mg.toFixed(1).replace('.', ','), cor: PPTX_CORES.roxoCl, borda: PPTX_CORES.roxo, corTexto: PPTX_CORES.roxoEsc }];
  const cW = 1.82, cH = 0.9, cY = 0.82, cGap = 0.065;
  cards.forEach((c, i) => { const cX = 0.25 + i * (cW + cGap); slide.addShape(prs.ShapeType.rect, { x: cX, y: cY, w: cW, h: cH, fill: { color: c.cor }, line: { color: c.borda, pt: 1.5 }, rectRadius: 0.07 }); slide.addText(c.valor, { x: cX, y: cY + 0.08, w: cW, h: 0.48, fontSize: 24, bold: true, color: c.corTexto, align: 'center', fontFace: 'Calibri' }); slide.addText(c.label, { x: cX, y: cY + 0.55, w: cW, h: 0.28, fontSize: 10, color: PPTX_CORES.cinzaEsc, align: 'center', fontFace: 'Calibri' }); });
  slide.addText('Média por disciplina', { x: 0.3, y: 1.9, w: 9.4, h: 0.3, fontSize: 11, bold: true, color: PPTX_CORES.cinzaEsc, fontFace: 'Calibri' });
  const discRows = [['Disciplina', 'Média', 'Situação'].map(t => ({ text: t, options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.azul, align: 'center' } })), ...discs.map((d, idx) => { const mds = regs.filter(r => r.DISCIPLINA === d).map(r => calcMedia(r)).filter(m => m !== null), mdg = mds.length ? +(mds.reduce((a, b) => a + b, 0) / mds.length).toFixed(2) : null, bg = idx % 2 === 0 ? PPTX_CORES.branco : PPTX_CORES.cinzaCl, cor = mdg === null ? PPTX_CORES.cinza : mdg >= CFG.mediaAprov ? PPTX_CORES.verde : PPTX_CORES.verm; return [{ text: d, options: { fill: bg, bold: true } }, { text: mdg !== null ? mdg.toFixed(1).replace('.', ',') : '—', options: { fill: bg, align: 'center', color: cor, bold: true } }, { text: mdg === null ? '—' : mdg >= CFG.mediaAprov ? 'Acima da média' : 'Abaixo da média', options: { fill: bg, align: 'center', color: cor } }]; })];
  slide.addTable(discRows, { x: 0.3, y: 2.22, w: 5.5, h: Math.min(1.8, 0.28 * (discs.length + 1)), fontSize: 10, fontFace: 'Calibri', color: PPTX_CORES.cinzaEsc, border: { pt: 0.5, color: 'E5E7EB' }, rowH: 0.28, colW: [3.2, 1.2, 1.1] });
  const regsDigTurma = APP.digital.filter(r => r.TURMA === turma);
  if (regsDigTurma.length) { const calcDig = calcDigital(regsDigTurma); if (calcDig) { slide.addShape(prs.ShapeType.rect, { x: 5.95, y: 2.22, w: 3.75, h: 1.5, fill: { color: 'FDF4FF' }, line: { color: 'E9D5FF', pt: 1 }, rectRadius: 0.08 }); slide.addText('💻 Material Digital', { x: 6.1, y: 2.3, w: 3.5, h: 0.28, fontSize: 11, bold: true, color: PPTX_CORES.roxoEsc, fontFace: 'Calibri' }); const digCor = calcDig.pct >= 80 ? PPTX_CORES.verde : calcDig.pct >= 50 ? PPTX_CORES.amber : PPTX_CORES.verm; slide.addText(`${calcDig.pct.toFixed(0)}%`, { x: 6.1, y: 2.55, w: 3.5, h: 0.52, fontSize: 30, bold: true, color: digCor, fontFace: 'Calibri', align: 'center' }); slide.addText(`${calcDig.concluido} de ${calcDig.previsto} aulas concluídas`, { x: 6.1, y: 3.05, w: 3.5, h: 0.25, fontSize: 10, color: PPTX_CORES.cinza, fontFace: 'Calibri', align: 'center' }); } }
  return slide;
}

function pptxSlidesDetalhado(prs, turma, disc) {
  const regs = APP.notas.filter(r => (!turma || r.TURMA === turma) && (!disc || r.DISCIPLINA === disc));
  if (!regs.length) return;
  const alunos = unique(regs.map(r => r.ESTUDANTE));
  let ap = 0, rp = 0, alF = 0; alunos.forEach(a => { const ds = regs.filter(r => r.ESTUDANTE === a).map(r => dadosReg(r)); if (ds.some(d => d.sit === 'Reprovado')) rp++; else ap++; if (ds.some(d => d.alFreq)) alF++; });
  const ms = regs.map(r => calcMedia(r)).filter(m => m !== null), mg = ms.length ? +(ms.reduce((a, b) => a + b, 0) / ms.length).toFixed(1) : 0;
  const headerTitulo = [turma, disc].filter(Boolean).join(' · ') || 'Relatório Detalhado';
  const slideSumario = prs.addSlide();
  pptxAddHeader(slideSumario, prs, PPTX_CORES.roxo, `Resumo — ${headerTitulo}`);
  const metCards = [{ label: 'Total de Alunos', valor: String(alunos.length), cor: PPTX_CORES.azulCl, borda: PPTX_CORES.azul, corV: PPTX_CORES.azulEsc }, { label: 'Aprovados', valor: String(ap), cor: PPTX_CORES.verdeCl, borda: PPTX_CORES.verde, corV: PPTX_CORES.verdeEsc }, { label: 'Reprovados', valor: String(rp), cor: PPTX_CORES.vermCl, borda: PPTX_CORES.verm, corV: PPTX_CORES.verm }, { label: 'Alerta Freq.', valor: String(alF), cor: PPTX_CORES.amberCl, borda: PPTX_CORES.amber, corV: PPTX_CORES.amber }, { label: 'Média Geral', valor: mg.toFixed(1).replace('.', ','), cor: PPTX_CORES.roxoCl, borda: PPTX_CORES.roxo, corV: PPTX_CORES.roxoEsc }];
  const cW = 1.82, cH = 0.95, cY = 0.8, cGap = 0.065;
  metCards.forEach((c, i) => { const cX = 0.25 + i * (cW + cGap); slideSumario.addShape(prs.ShapeType.rect, { x: cX, y: cY, w: cW, h: cH, fill: { color: c.cor }, line: { color: c.borda, pt: 1.5 }, rectRadius: 0.07 }); slideSumario.addText(c.valor, { x: cX, y: cY + 0.1, w: cW, h: 0.52, fontSize: 26, bold: true, color: c.corV, align: 'center', fontFace: 'Calibri' }); slideSumario.addText(c.label, { x: cX, y: cY + 0.62, w: cW, h: 0.26, fontSize: 10, color: PPTX_CORES.cinzaEsc, align: 'center', fontFace: 'Calibri' }); });
  const [txAp, txRp] = distribuirPct([ap, rp]), barY = 1.95, barX = 0.3, barW = 9.4, barH = 0.32;
  slideSumario.addShape(prs.ShapeType.rect, { x: barX, y: barY, w: barW, h: barH, fill: { color: PPTX_CORES.vermCl }, line: { color: PPTX_CORES.verm, pt: 0.5 } });
  if (txAp > 0) slideSumario.addShape(prs.ShapeType.rect, { x: barX, y: barY, w: +(barW * txAp / 100).toFixed(2), h: barH, fill: { color: PPTX_CORES.verde }, line: { color: PPTX_CORES.verde, pt: 0 } });
  slideSumario.addText(`Aprovados: ${txAp}%`, { x: barX + 0.1, y: barY, w: 3, h: barH, fontSize: 10, bold: true, color: PPTX_CORES.branco, valign: 'middle', fontFace: 'Calibri' });
  slideSumario.addText(`Reprovados: ${txRp}%`, { x: barX + barW - 2.5, y: barY, w: 2.4, h: barH, fontSize: 10, bold: true, color: PPTX_CORES.verm, align: 'right', valign: 'middle', fontFace: 'Calibri' });
  if (!disc) { const discsRegs = unique(regs.map(r => r.DISCIPLINA)); slideSumario.addText('Médias por disciplina', { x: 0.3, y: 2.42, w: 9.4, h: 0.28, fontSize: 11, bold: true, color: PPTX_CORES.cinzaEsc, fontFace: 'Calibri' }); const discTabRows = [[{ text: 'Disciplina', options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.roxo } }, { text: 'Média', options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.roxo, align: 'center' } }], ...discsRegs.map((d, idx) => { const mds = regs.filter(r => r.DISCIPLINA === d).map(r => calcMedia(r)).filter(m => m !== null), mdg = mds.length ? +(mds.reduce((a, b) => a + b, 0) / mds.length).toFixed(2) : null, bg = idx % 2 === 0 ? PPTX_CORES.branco : PPTX_CORES.cinzaCl; return [{ text: d, options: { fill: bg, bold: true } }, { text: mdg !== null ? mdg.toFixed(1).replace('.', ',') : '—', options: { fill: bg, align: 'center', color: mdg !== null && mdg >= CFG.mediaAprov ? PPTX_CORES.verde : PPTX_CORES.verm, bold: true } }]; })]; slideSumario.addTable(discTabRows, { x: 0.3, y: 2.72, w: 5.2, h: Math.min(1.2, 0.27 * (discsRegs.length + 1)), fontSize: 10, fontFace: 'Calibri', color: PPTX_CORES.cinzaEsc, border: { pt: 0.5, color: 'E5E7EB' }, rowH: 0.27, colW: [3.8, 1.4] }); }
  const ALUNOS_POR_PAGINA = 14, paginas = Math.ceil(regs.length / ALUNOS_POR_PAGINA);
  for (let p = 0; p < paginas; p++) {
    const slideD = prs.addSlide(), pageRegs = regs.slice(p * ALUNOS_POR_PAGINA, (p + 1) * ALUNOS_POR_PAGINA), pageLabel = paginas > 1 ? ` (${p + 1}/${paginas})` : '';
    pptxAddHeader(slideD, prs, PPTX_CORES.azulEsc, `Lista de Alunos${pageLabel} — ${headerTitulo}`, 0.55);
    const hasDisc = !!disc;
    const cabs = hasDisc ? ['Aluno', 'Turma', 'B1', 'B2', 'B3', 'B4', 'Média', 'Freq.', 'Situação'] : ['Aluno', 'Média', 'Freq.', 'Situação'];
    const cabecalho = cabs.map(t => ({ text: t, options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.azul, align: 'center' } }));
    const dataRows = pageRegs.map((r, idx) => {
      const d = dadosReg(r), bg = idx % 2 === 0 ? PPTX_CORES.branco : PPTX_CORES.cinzaCl, medCor = d.media === null ? PPTX_CORES.cinza : d.media >= CFG.mediaAprov ? PPTX_CORES.verde : PPTX_CORES.verm, freqCor = !d.freq ? PPTX_CORES.cinza : d.freq.pct >= CFG.freqMin ? PPTX_CORES.verde : PPTX_CORES.verm, freqStr = d.freq ? `${d.freq.pct.toFixed(1)}%` : '—', sitStr = d.sit === 'Aprovado' ? (d.alFreq ? 'Apr. ⚠Freq' : 'Aprovado') : d.sit === 'Reprovado' ? 'Reprovado' : '—', sitCor = d.sit === 'Aprovado' ? PPTX_CORES.verde : d.sit === 'Reprovado' ? PPTX_CORES.verm : PPTX_CORES.cinza;
      const notaCell = v => ({ text: v !== null ? v.toFixed(1).replace('.', ',') : '—', options: { fill: v !== null && v < CFG.mediaAprov ? PPTX_CORES.vermCl : bg, align: 'center', color: v !== null && v < CFG.mediaAprov ? PPTX_CORES.verm : PPTX_CORES.cinzaEsc, bold: v !== null && v < CFG.mediaAprov } });
      if (hasDisc) return [{ text: r.ESTUDANTE, options: { fill: bg, fontSize: 9 } }, { text: r.TURMA, options: { fill: bg, align: 'center' } }, notaCell(toNum(r.b1)), notaCell(toNum(r.b2)), notaCell(toNum(r.b3)), notaCell(toNum(r.b4)), { text: d.media !== null ? d.media.toFixed(1).replace('.', ',') : '—', options: { fill: d.media !== null && d.media < CFG.mediaAprov ? PPTX_CORES.vermCl : bg, align: 'center', color: medCor, bold: true } }, { text: freqStr, options: { fill: bg, align: 'center', color: freqCor, bold: d.alFreq } }, { text: sitStr, options: { fill: bg, align: 'center', color: sitCor, bold: true } }];
      return [{ text: r.ESTUDANTE, options: { fill: bg, fontSize: 9 } }, { text: d.media !== null ? d.media.toFixed(1).replace('.', ',') : '—', options: { fill: d.media !== null && d.media < CFG.mediaAprov ? PPTX_CORES.vermCl : bg, align: 'center', color: medCor, bold: true } }, { text: freqStr, options: { fill: bg, align: 'center', color: freqCor, bold: d.alFreq } }, { text: sitStr, options: { fill: bg, align: 'center', color: sitCor, bold: true } }];
    });
    const colWidths = hasDisc ? [3.0, 0.7, 0.7, 0.7, 0.7, 0.7, 0.85, 0.8, 1.05] : [6.2, 1.2, 1.2, 1.8];
    slideD.addTable([cabecalho, ...dataRows], { x: 0.3, y: 0.65, w: colWidths.reduce((a, b) => a + b, 0), h: 3.7, fontSize: 10, fontFace: 'Calibri', color: PPTX_CORES.cinzaEsc, border: { pt: 0.5, color: 'E5E7EB' }, rowH: 0.245, colW: colWidths });
  }
  const regsDigFilt = APP.digital.filter(r => (!turma || r.TURMA === turma) && (!disc || r.DISCIPLINA === disc));
  if (regsDigFilt.length) {
    const slideD2 = prs.addSlide(); pptxAddHeader(slideD2, prs, PPTX_CORES.roxo, '💻 Material Digital — Progresso por Bimestre'); const discsD = unique(regsDigFilt.map(r => r.DISCIPLINA)), turmasD = unique(regsDigFilt.map(r => r.TURMA)), combos = []; turmasD.forEach(t => discsD.forEach(d => { const rr = regsDigFilt.filter(x => x.TURMA === t && x.DISCIPLINA === d); if (rr.length) combos.push({ turma: t, disc: d, regs: rr }); }));
    const digTabRows = [[{ text: 'Turma', options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.roxo } }, { text: 'Disciplina', options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.roxo } }, { text: '1º Bim', options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.roxo, align: 'center' } }, { text: '2º Bim', options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.roxo, align: 'center' } }, { text: '3º Bim', options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.roxo, align: 'center' } }, { text: '4º Bim', options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.roxo, align: 'center' } }, { text: 'Total %', options: { bold: true, color: PPTX_CORES.branco, fill: PPTX_CORES.roxo, align: 'center' } }], ...combos.map((c, idx) => { const bims = calcDigitalBimestres(c.regs), total = calcDigital(c.regs), bg = idx % 2 === 0 ? PPTX_CORES.branco : 'F3E8FF'; const bimCell = b => { if (!b) return { text: '—', options: { fill: bg, align: 'center', color: PPTX_CORES.cinza } }; const cor = b.pct >= 80 ? PPTX_CORES.verde : b.pct >= 50 ? PPTX_CORES.amber : PPTX_CORES.verm; return { text: `${b.pct.toFixed(0)}%\n(${b.concluido}/${b.previsto})`, options: { fill: bg, align: 'center', color: cor, bold: true, fontSize: 9 } }; }; const totalCor = total ? (total.pct >= 80 ? PPTX_CORES.verde : total.pct >= 50 ? PPTX_CORES.amber : PPTX_CORES.verm) : PPTX_CORES.cinza; return [{ text: c.turma, options: { fill: bg, bold: true, align: 'center' } }, { text: c.disc, options: { fill: bg } }, bimCell(bims[0]), bimCell(bims[1]), bimCell(bims[2]), bimCell(bims[3]), { text: total ? `${total.pct.toFixed(0)}%` : '—', options: { fill: bg, align: 'center', color: totalCor, bold: true } }]; })];
    slideD2.addTable(digTabRows, { x: 0.3, y: 0.75, w: 9.4, h: Math.min(3.5, 0.38 * (combos.length + 1)), fontSize: 10, fontFace: 'Calibri', color: PPTX_CORES.cinzaEsc, border: { pt: 0.5, color: 'E9D5FF' }, rowH: 0.38, colW: [0.9, 2.3, 1.1, 1.1, 1.1, 1.1, 1.0] });
  }
}

function popSelectsPPTX() {
  if (!APP.notas.length) return;
  const turmas = unique(APP.notas.map(r => r.TURMA)), discs = unique(APP.notas.map(r => r.DISCIPLINA));
  const selT = document.getElementById('pptx-sel-turma'), selD = document.getElementById('pptx-sel-disc');
  if (selT) selT.innerHTML = '<option value="">Todas as turmas (Relatório Geral)</option>' + turmas.map(t => `<option value="${t}">${t}</option>`).join('');
  if (selD) selD.innerHTML = '<option value="">Todas as disciplinas</option>' + discs.map(d => `<option value="${d}">${d}</option>`).join('');
  atualizarContextoPPTX();
}

function atualizarContextoPPTX() {
  if (!APP.notas.length) return;
  const turmaSel = document.getElementById('pptx-sel-turma')?.value || '', discSel = document.getElementById('pptx-sel-disc')?.value || '';
  const modoDetalhado = !!(turmaSel || discSel), turmas = unique(APP.notas.map(r => r.TURMA));
  const alunos = unique(APP.notas.filter(r => (!turmaSel || r.TURMA === turmaSel) && (!discSel || r.DISCIPLINA === discSel)).map(r => r.ESTUDANTE));
  const prevTitulo = document.getElementById('pptx-prev-titulo'), prevDesc = document.getElementById('pptx-prev-desc'), prevSlides = document.getElementById('pptx-prev-slides');
  if (!modoDetalhado) {
    if (prevTitulo) prevTitulo.textContent = 'Relatório Geral da Escola';
    if (prevDesc) prevDesc.textContent = `${turmas.length} turmas · ${unique(APP.notas.map(r => r.ESTUDANTE)).length} alunos`;
    if (prevSlides) prevSlides.innerHTML = `<div class="pptx-slide-item"><span class="pptx-slide-num">1</span> Capa</div><div class="pptx-slide-item"><span class="pptx-slide-num">2</span> Painel consolidado</div>${turmas.map((t, i) => `<div class="pptx-slide-item"><span class="pptx-slide-num">${i + 3}</span> Turma ${t}</div>`).join('')}`;
  } else {
    const label = [turmaSel && `Turma ${turmaSel}`, discSel].filter(Boolean).join(' · ');
    if (prevTitulo) prevTitulo.textContent = `Relatório Detalhado — ${label}`;
    if (prevDesc) prevDesc.textContent = `${alunos.length} alunos · dados filtrados`;
    const regs = APP.notas.filter(r => (!turmaSel || r.TURMA === turmaSel) && (!discSel || r.DISCIPLINA === discSel));
    const pags = Math.ceil(regs.length / 14), temDig = APP.digital.some(r => (!turmaSel || r.TURMA === turmaSel) && (!discSel || r.DISCIPLINA === discSel));
    if (prevSlides) prevSlides.innerHTML = `<div class="pptx-slide-item"><span class="pptx-slide-num">1</span> Capa</div><div class="pptx-slide-item"><span class="pptx-slide-num">2</span> Resumo e indicadores</div>${Array.from({ length: pags }, (_, i) => `<div class="pptx-slide-item"><span class="pptx-slide-num">${i + 3}</span> Lista de alunos${pags > 1 ? ` (${i + 1}/${pags})` : ''}</div>`).join('')}${temDig ? `<div class="pptx-slide-item"><span class="pptx-slide-num">${3 + pags}</span> Material Digital</div>` : ''}`;
  }
}

function gerarPPTX() {
  if (!APP.notas.length) { alert('Importe dados antes de gerar a apresentação.'); return; }
  loader(true, 'Gerando apresentação PPTX...');
  setTimeout(() => {
    try {
      const turmaSel = document.getElementById('pptx-sel-turma')?.value || '', discSel = document.getElementById('pptx-sel-disc')?.value || '';
      const modoDetalhado = !!(turmaSel || discSel), data = new Date().toLocaleDateString('pt-BR');
      const prs = new PptxGenJS(); prs.layout = 'LAYOUT_WIDE'; prs.author = CFG.escola; prs.company = CFG.escola; prs.subject = 'Conselho de Classe';
      if (!modoDetalhado) {
        pptxCapa(prs, 'Relatório Geral — Conselho de Classe', `${CFG.escola} · ${CFG.etapa}`, `Gerado em ${data} · Todas as turmas`);
        const turmas = unique(APP.notas.map(r => r.TURMA)), alunos = unique(APP.notas.map(r => r.ESTUDANTE));
        let ap = 0, rp = 0, alF = 0; alunos.forEach(a => { const regs = APP.notas.filter(r => r.ESTUDANTE === a), ds = regs.map(r => dadosReg(r)); if (ds.some(d => d.sit === 'Reprovado')) rp++; else ap++; if (ds.some(d => d.alFreq)) alF++; });
        const medias = APP.notas.map(r => calcMedia(r)).filter(m => m !== null), mg = medias.length ? +(medias.reduce((a, b) => a + b, 0) / medias.length) : 0;
        const resumoTurmas = turmas.map(t => { const al = unique(APP.notas.filter(r => r.TURMA === t).map(r => r.ESTUDANTE)); let ta = 0, tr = 0, taf = 0; al.forEach(a => { const ds = APP.notas.filter(r => r.ESTUDANTE === a && r.TURMA === t).map(r => dadosReg(r)); if (ds.some(d => d.sit === 'Reprovado')) tr++; else ta++; if (ds.some(d => d.alFreq)) taf++; }); const ms2 = APP.notas.filter(r => r.TURMA === t).map(r => calcMedia(r)).filter(m => m !== null), tmg = ms2.length ? +(ms2.reduce((a, b) => a + b, 0) / ms2.length).toFixed(1) : 0; return { turma: t, alunos: al.length, aprov: ta, reprov: tr, alFreq: taf, mg: +tmg }; });
        pptxSlideResumoEscola(prs, { turmas: resumoTurmas, totalAlunos: alunos.length, aprov: ap, reprov: rp, alFreqs: alF, mg });
        turmas.forEach(t => pptxSlideResumoPorTurma(prs, t));
        prs.writeFile({ fileName: `Conselho_Classe_Geral_${data.replace(/\//g, '-')}.pptx` });
      } else {
        const subtitulo = [turmaSel && `Turma ${turmaSel}`, discSel].filter(Boolean).join(' · ');
        pptxCapa(prs, 'Relatório Detalhado — Conselho de Classe', `${CFG.escola} · ${subtitulo}`, `Gerado em ${data} · ${CFG.etapa}`);
        pptxSlidesDetalhado(prs, turmaSel, discSel);
        const nomeSafe = subtitulo.replace(/[^a-zA-Z0-9À-ÿ\s]/g, '').replace(/\s+/g, '_');
        prs.writeFile({ fileName: `Conselho_Classe_${nomeSafe}_${data.replace(/\//g, '-')}.pptx` });
      }
    } catch (err) { alert('Erro ao gerar PPTX: ' + err.message); console.error(err); }
    loader(false);
  }, 80);
}

// ── DEMO ───────────────────────────────────────────────────

function carregarDemo() {
  const nomes = ['ALEXIA RAISSA ALMEIDA ALVES', 'ALYANDRO ARIEL SANTOS FREITAS', 'CRISTOFFER MIGUEL FERREIRA LIMA', 'IRANI CAIRAC PRESTES NETO', 'JOAO LUCAS BORGES DE OLIVEIRA', 'LEANDRO HENRIQUE DE ALMEIDA', 'LUIS GUSTAVO DE OLIVEIRA SANTOS', 'MARIA FERNANDA COSTA SILVA', 'PEDRO HENRIQUE RAMOS JUNIOR', 'ANA CAROLINA MENDES SOUSA', 'GABRIEL TORRES NASCIMENTO', 'BEATRIZ LIMA CAVALCANTE', 'LUCAS AUGUSTO FERREIRA DIAS', 'JULIA APARECIDA MOREIRA COSTA', 'THIAGO CAMARGO BORGES', 'LARISSA BRITO CARVALHO', 'VINICIUS MARTINS PRADO', 'CAMILA RODRIGUES PEREIRA', 'FELIPE SOUZA ALBUQUERQUE', 'NATHALIA GOMES RIBEIRO'];
  const turmas = ['2ªA', '2ªC', '3ªA'], discs = ['MATEMÁTICA', 'PROGRAMAÇÃO', 'ROBÓTICA'];
  APP.notas = []; APP.faltas = []; APP.digital = [];
  turmas.forEach(t => {
    const al = nomes.slice(0, t === '3ªA' ? 18 : 20);
    discs.forEach(d => {
      al.forEach(a => {
        const rnd = Math.random();
        const b1 = +(rnd < .1 ? 1.5 + Math.random() * 2 : rnd < .25 ? 3 + Math.random() * 1.5 : 5 + Math.random() * 5).toFixed(1);
        const b2 = +(Math.min(10, Math.max(0, b1 + (Math.random() * 2 - 0.8)))).toFixed(1);
        const b3 = +(Math.min(10, Math.max(0, b2 + (Math.random() * 2 - 0.7)))).toFixed(1);
        const b4 = Math.random() > .3 ? +(Math.min(10, Math.max(0, b3 + (Math.random() * 2 - 0.5)))).toFixed(1) : null;
        APP.notas.push({ ESTUDANTE: a, TURMA: t, DISCIPLINA: d, b1, b2, b3, b4 });
        const a1 = 38 + Math.floor(Math.random() * 4), a2 = 39 + Math.floor(Math.random() * 4), a3 = 38 + Math.floor(Math.random() * 5), a4 = b4 !== null ? 39 + Math.floor(Math.random() * 4) : null;
        const faltas = Math.random() < .15 ? Math.floor(Math.random() * 15) + 5 : Math.floor(Math.random() * 5);
        APP.faltas.push({ ESTUDANTE: a, TURMA: t, DISCIPLINA: d, a1, f1: Math.floor(Math.random() * 4), a2, f2: Math.floor(faltas * .4), a3, f3: Math.floor(faltas * .4), a4, f4: a4 ? Math.floor(faltas * .2) : null });
      });
      APP.digital.push({ TURMA: t, DISCIPLINA: d, prev1: 12, conc1: Math.round(12 * (0.9 + Math.random() * .2)), prev2: 12, conc2: Math.round(12 * (0.85 + Math.random() * .15)), prev3: 12, conc3: Math.round(12 * (0.6 + Math.random() * .2)), prev4: 12, conc4: null });
    });
  });
  const statusEl = document.getElementById('status-up');
  if (statusEl) statusEl.innerHTML = `<span style="color:var(--verde);">✓ Dados de exemplo carregados: ${APP.notas.length} registros em ${turmas.length} turmas · ${APP.digital.length} registros de material digital.</span>`;
  buildNav();
  showSection('geral');
}

// ── MODELO EXCEL ───────────────────────────────────────────

function baixarModelo() {
  const wb = XLSX.utils.book_new();
  const wsNotas = XLSX.utils.aoa_to_sheet([['ESTUDANTE', 'TURMA', 'DISCIPLINA', 'NOTA 1º BIMESTRE', 'NOTA 2º BIMESTRE', 'NOTA 3º BIMESTRE', 'NOTA 4º BIMESTRE'], ['NOME DO ALUNO', '2ªA', 'MATEMÁTICA', 7, 8, 6, 7], ['OUTRO ALUNO', '2ªA', 'MATEMÁTICA', 5, 4, '', '']]);
  wsNotas['!cols'] = [{ wch: 35 }, { wch: 8 }, { wch: 20 }, { wch: 18 }, { wch: 18 }, { wch: 18 }, { wch: 18 }];
  XLSX.utils.book_append_sheet(wb, wsNotas, 'Notas');
  const wsFaltas = XLSX.utils.aoa_to_sheet([['ESTUDANTE', 'TURMA', 'DISCIPLINA', 'AULAS 1º BIM', 'FALTAS 1º BIM', 'AULAS 2º BIM', 'FALTAS 2º BIM', 'AULAS 3º BIM', 'FALTAS 3º BIM', 'AULAS 4º BIM', 'FALTAS 4º BIM'], ['NOME DO ALUNO', '2ªA', 'MATEMÁTICA', 38, 0, 40, 2, 39, 5, 39, 3], ['OUTRO ALUNO', '2ªA', 'MATEMÁTICA', 38, 4, 40, 14, '', '', '', '']]);
  wsFaltas['!cols'] = [{ wch: 35 }, { wch: 8 }, { wch: 20 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }];
  XLSX.utils.book_append_sheet(wb, wsFaltas, 'FALTAS');
  const wsDigital = XLSX.utils.aoa_to_sheet([['TURMA', 'DISCIPLINA', 'PREVISTO 1º BIM', 'CONCLUIDO 1º BIM', 'PREVISTO 2º BIM', 'CONCLUIDO 2º BIM', 'PREVISTO 3º BIM', 'CONCLUIDO 3º BIM', 'PREVISTO 4º BIM', 'CONCLUIDO 4º BIM'], ['2ªA', 'MATEMÁTICA', 10, 10, 10, 8, 10, 5, 10, ''], ['2ªA', 'ROBÓTICA', 8, 8, 8, 6, 8, 4, 8, '']]);
  wsDigital['!cols'] = [{ wch: 8 }, { wch: 20 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 }];
  XLSX.utils.book_append_sheet(wb, wsDigital, 'DIGITAL');
  XLSX.writeFile(wb, 'modelo_conselho_classe.xlsx');
}

// ── LIMPAR ─────────────────────────────────────────────────

function limpar() {
  if (!confirm('Limpar todos os dados?')) return;
  APP.notas = []; APP.faltas = []; APP.digital = []; APP.turmaSel = null; APP.discSel = null;
  Object.values(APP.charts).forEach(c => c.destroy()); APP.charts = {};
  ['nav-turmas', 'nav-discs'].forEach(id => { const el = document.getElementById(id); if (el) el.innerHTML = ''; });
  const statusEl = document.getElementById('status-up'); if (statusEl) statusEl.innerHTML = '';
  const inpEl = document.getElementById('inp-p'); if (inpEl) inpEl.value = '';
  ['acc-turmas', 'acc-discs'].forEach(id => { const acc = document.getElementById(id); if (!acc) return; const toggle = acc.querySelector('.sb-acc-toggle'), body = acc.querySelector('.sb-acc-body'); if (toggle) { toggle.classList.remove('open'); toggle.setAttribute('aria-expanded', 'false'); } if (body) body.classList.remove('open'); });
  ['acc-turmas-count', 'acc-discs-count'].forEach(id => { const el = document.getElementById(id); if (el) { el.textContent = ''; el.classList.remove('visible'); } });
  showSection('upload');
}

// ── SIDEBAR MOBILE ─────────────────────────────────────────

function abrirSidebar() { document.getElementById('sidebar').classList.add('open'); document.getElementById('sb-overlay').classList.add('open'); document.body.style.overflow = 'hidden'; }
function fecharSidebar() { document.getElementById('sidebar').classList.remove('open'); document.getElementById('sb-overlay').classList.remove('open'); document.body.style.overflow = ''; }
document.addEventListener('click', e => { if (window.innerWidth <= 768 && e.target.closest('#sidebar .nav-link')) fecharSidebar(); });

// ── LGPD ───────────────────────────────────────────────────

function aceitarCookies() { localStorage.setItem('lgpd_consent', 'accepted'); document.getElementById('lgpd-banner').classList.add('hidden'); }
function recusarCookies() { localStorage.setItem('lgpd_consent', 'essential'); document.getElementById('lgpd-banner').classList.add('hidden'); }
function abrirPolitica() { document.getElementById('lgpd-modal').classList.add('open'); }
function fecharPolitica(e) { if (e.target === document.getElementById('lgpd-modal')) document.getElementById('lgpd-modal').classList.remove('open'); }

// ── INIT ───────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', () => {
  const loaderEl = document.getElementById('loader');
  if (loaderEl) loaderEl.style.display = 'none';

  // Verificar dependências de CDN
  const depsMissing = [
    { nome: 'Chart.js', ok: typeof Chart !== 'undefined' },
    { nome: 'SheetJS (XLSX)', ok: typeof XLSX !== 'undefined' },
    { nome: 'PptxGenJS', ok: typeof PptxGenJS !== 'undefined' },
  ].filter(d => !d.ok);
  if (depsMissing.length > 0) {
    const statusEl = document.getElementById('status-up');
    if (statusEl) statusEl.innerHTML =
      `<div style="background:var(--verm-cl);border:1px solid var(--verm);border-radius:var(--r);padding:12px 16px;color:var(--verm-esc);font-size:13px;line-height:1.7;">
        <strong>⚠ Erro de carregamento</strong><br>
        As seguintes dependências não foram carregadas. Verifique sua conexão com a internet e recarregue a página:<br>
        <strong>${depsMissing.map(d => d.nome).join(', ')}</strong>
      </div>`;
    const uzEl = document.getElementById('uz-p');
    if (uzEl) uzEl.style.pointerEvents = 'none';
  }

  // Listener para consentimento LGPD
  const consent = localStorage.getItem('lgpd_consent');
  if (consent) document.getElementById('lgpd-banner').classList.add('hidden');

  // Event delegation: abrirModal via data-nome + data-turma (modal-trigger)
  document.addEventListener('click', e => {
    const mt = e.target.closest('.modal-trigger');
    if (mt && mt.dataset.nome && mt.dataset.turma) {
      abrirModal(decodeURIComponent(mt.dataset.nome), decodeURIComponent(mt.dataset.turma));
      return;
    }
    const pt = e.target.closest('.pdf-trigger');
    if (pt && pt.dataset.nome && pt.dataset.turma) {
      pdfAlunoNome(decodeURIComponent(pt.dataset.nome), decodeURIComponent(pt.dataset.turma));
    }
  });
});
